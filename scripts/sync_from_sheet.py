#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SDC x HMP Production Planning - Google Sheets -> index.html sync

Reads the Roster / Product_Capability / Labor_Cost / QC_Records tabs from the
shared Google Sheet, rebuilds DEFAULT_PRODUCERS / DEFAULT_LABOR /
DEFAULT_QC_HISTORY and the DATA_VERSION stamp inside index.html, and leaves
everything else in the file untouched.

Required environment variables:
  GOOGLE_SERVICE_ACCOUNT_JSON  - full JSON key of the service account (string)
  SPREADSHEET_ID                - the Google Sheet's ID (the long string in its URL)

Unmatched QC rows (no HMP_ID and no name match against the active roster) are
written to the Sync_Errors tab in the sheet itself, and also printed to the
GitHub Actions log, instead of silently being dropped.
"""
import json
import os
import re
import sys
from datetime import datetime, timezone
from collections import defaultdict

import gspread
from google.oauth2.service_account import Credentials

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

INDEX_HTML_PATH = os.environ.get("INDEX_HTML_PATH", "index.html")


def die(msg):
    print(f"::error::{msg}")
    sys.exit(1)


def get_client():
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    if not raw:
        die("GOOGLE_SERVICE_ACCOUNT_JSON is not set")
    try:
        info = json.loads(raw)
    except json.JSONDecodeError as e:
        die(f"GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON: {e}")
    creds = Credentials.from_service_account_info(info, scopes=SCOPES)
    return gspread.authorize(creds)


def get_sheet_id():
    sid = os.environ.get("SPREADSHEET_ID")
    if not sid:
        die("SPREADSHEET_ID is not set")
    return sid


def ws_records(sh, title):
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        die(f'Tab "{title}" not found in the spreadsheet')
    return ws, ws.get_all_records(default_blank="")


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def norm_name_key(s):
    return norm(s).upper()


def truthy_capable(v):
    v = norm(v).lower()
    return v not in ("", "n", "no", "false", "0")


def parse_date(s):
    s = norm(s)
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def build_producers_and_roster(roster_rows, cap_rows):
    """Returns (PRODUCERS list, roster_by_id dict, roster_by_name dict) - only Active people."""
    roster_by_id = {}
    roster_by_name = {}
    order = []
    for r in roster_rows:
        hid = norm(r.get("HMP_ID"))
        name = norm(r.get("Name"))
        if not hid or not name:
            continue
        status = norm(r.get("Status")).lower()
        active = status != "inactive"
        entry = {
            "id": hid,
            "name": name,
            "category": norm(r.get("Category")),
            "batch": r.get("Batch"),
            "active": active,
        }
        roster_by_id[hid] = entry
        roster_by_name[norm_name_key(name)] = entry
        if active:
            order.append(hid)

    caps_by_id = defaultdict(list)
    for r in cap_rows:
        hid = norm(r.get("HMP_ID"))
        pname = norm(r.get("Product_Name"))
        if not hid or not pname:
            continue
        if truthy_capable(r.get("Capable")):
            caps_by_id[hid].append(pname)

    producers = []
    for hid in order:
        entry = roster_by_id[hid]
        try:
            batch = int(entry["batch"])
        except (TypeError, ValueError):
            batch = entry["batch"] or ""
        producers.append({
            "category": entry["category"],
            "batch": batch,
            "name": entry["name"],
            "capable": caps_by_id.get(hid, []),
        })
    return producers, roster_by_id, roster_by_name


def build_labor(labor_rows, run_date):
    by_product = defaultdict(list)
    for r in labor_rows:
        pname = norm(r.get("Product_Name"))
        if not pname:
            continue
        try:
            cost = float(r.get("Labor_Cost"))
        except (TypeError, ValueError):
            continue
        eff = parse_date(r.get("Effective_From"))
        by_product[pname].append((eff, cost))

    labor = {}
    for pname, rows in by_product.items():
        dated = [r for r in rows if r[0] is not None]
        past_or_today = [r for r in dated if r[0] <= run_date]
        if past_or_today:
            chosen = max(past_or_today, key=lambda r: r[0])
        elif dated:
            chosen = min(dated, key=lambda r: r[0])
        else:
            chosen = rows[-1]
        labor[pname] = chosen[1]
    return labor


def build_qc_history(qc_rows, roster_by_id, roster_by_name, run_date):
    groups = defaultdict(lambda: defaultdict(lambda: {
        "batch": None, "submitted": 0, "qcPass": 0, "products": []
    }))
    sync_errors = []

    for r in qc_rows:
        entry_id = norm(r.get("Entry_ID"))
        try:
            year = int(r.get("Year"))
            month = int(r.get("Month"))
        except (TypeError, ValueError):
            sync_errors.append((entry_id, "QC_Records", norm(r.get("HMP_Name_Raw")),
                                 "Year/Month missing or not a number"))
            continue
        product = norm(r.get("Product_Name"))
        try:
            submitted = float(r.get("Submitted_Qty") or 0)
        except (TypeError, ValueError):
            submitted = 0
        try:
            qc_pass = float(r.get("QC_Pass_Qty") or 0)
        except (TypeError, ValueError):
            qc_pass = 0
        raw_name = norm(r.get("HMP_Name_Raw"))
        hid = norm(r.get("HMP_ID"))
        batch = r.get("Batch")

        person = None
        if hid and hid in roster_by_id and roster_by_id[hid]["active"]:
            person = roster_by_id[hid]
        elif raw_name:
            person = roster_by_name.get(norm_name_key(raw_name))
            if person and not person["active"]:
                person = None

        if not person:
            sync_errors.append((entry_id, "QC_Records", raw_name or hid,
                                 "No matching active person found in Roster (check ID/name spelling)"))
            continue

        key = (year, month)
        agg = groups[key][person["name"]]
        if agg["batch"] is None:
            try:
                agg["batch"] = int(batch) if batch not in (None, "") else person["batch"]
            except (TypeError, ValueError):
                agg["batch"] = person["batch"]
        agg["submitted"] += submitted
        agg["qcPass"] += qc_pass
        agg["products"].append({
            "product": product,
            "submitted": submitted,
            "qc_pass": qc_pass,
        })

    history = []
    for (year, month) in sorted(groups.keys(), reverse=True):
        data = {}
        for name, agg in groups[(year, month)].items():
            rate = round(agg["qcPass"] / agg["submitted"] * 100, 1) if agg["submitted"] else 0
            data[name] = {
                "batch": agg["batch"],
                "submitted": agg["submitted"],
                "qcPass": agg["qcPass"],
                "rate": rate,
                "products": agg["products"],
            }
        sheet = MONTH_ABBR[month - 1] if 1 <= month <= 12 else str(month)
        label = f"{year}년 {month}월"  # "YYYY년 M월"
        uploaded_at = datetime(year, month, 1, tzinfo=timezone.utc).isoformat().replace("+00:00", "Z")
        history.append({
            "sheet": sheet,
            "label": label,
            "uploadedAt": uploaded_at,
            "data": data,
        })
    return history, sync_errors


def build_engagement(qc_rows, roster_by_id, roster_by_name, labor, sync_errors):
    """Mirrors the app's own deriveEngagementFromQcMonth(): for each person/month,
    engagement value = sum over that month's QC rows of (Submitted_Qty * Labor_Cost[product])."""
    monthly = defaultdict(lambda: defaultdict(float))  # name -> periodKey -> value

    for r in qc_rows:
        try:
            year = int(r.get("Year"))
            month = int(r.get("Month"))
        except (TypeError, ValueError):
            continue
        product = norm(r.get("Product_Name"))
        try:
            submitted = float(r.get("Submitted_Qty") or 0)
        except (TypeError, ValueError):
            submitted = 0
        raw_name = norm(r.get("HMP_Name_Raw"))
        hid = norm(r.get("HMP_ID"))

        person = None
        if hid and hid in roster_by_id and roster_by_id[hid]["active"]:
            person = roster_by_id[hid]
        elif raw_name:
            person = roster_by_name.get(norm_name_key(raw_name))
            if person and not person["active"]:
                person = None
        if not person:
            continue  # already logged as a Sync_Errors row by build_qc_history()

        fee = labor.get(product)
        if fee is None:
            sync_errors.append(("", "QC_Records", product,
                                 "Product not found in Labor_Cost (excluded from engagement value)"))
            continue

        period_key = f"{year:04d}-{month:02d}"
        monthly[person["name"]][period_key] += submitted * fee

    engagement = {}
    for name, months in monthly.items():
        rounded = {pk: round(v) for pk, v in months.items()}
        engagement[name] = {
            "monthly": rounded,
            "total": sum(rounded.values()),
            "months": len(rounded),
        }
    return engagement


def write_sync_errors(sh, sync_errors):
    try:
        ws = sh.worksheet("Sync_Errors")
    except gspread.exceptions.WorksheetNotFound:
        print("Sync_Errors tab not found - skipping error write-back")
        return
    ws.clear()
    rows = [["Entry_ID", "Sheet", "Raw_HMP_Text", "Reason"]]
    for e in sync_errors:
        rows.append(list(e))
    ws.update(rows, value_input_option="RAW")


def replace_js_const(html, const_name, new_value_json):
    pattern = re.compile(
        r"const " + re.escape(const_name) + r"\s*=\s*.*?;\n", re.DOTALL
    )
    replacement = f"const {const_name}={new_value_json};\n"
    new_html, count = pattern.subn(replacement, html, count=1)
    if count != 1:
        die(f'Could not find "const {const_name}=...;" in {INDEX_HTML_PATH}')
    return new_html


def main():
    client = get_client()
    sh = client.open_by_key(get_sheet_id())

    _, roster_rows = ws_records(sh, "Roster")
    _, cap_rows = ws_records(sh, "Product_Capability")
    _, labor_rows = ws_records(sh, "Labor_Cost")
    _, qc_rows = ws_records(sh, "QC_Records")

    run_date = datetime.now(timezone.utc)

    producers, roster_by_id, roster_by_name = build_producers_and_roster(roster_rows, cap_rows)
    labor = build_labor(labor_rows, run_date)
    qc_history, sync_errors = build_qc_history(qc_rows, roster_by_id, roster_by_name, run_date)
    engagement = build_engagement(qc_rows, roster_by_id, roster_by_name, labor, sync_errors)

    if not producers:
        die("Roster produced 0 active producers - aborting so we never overwrite index.html with empty data")
    if not labor:
        die("Labor_Cost produced 0 entries - aborting so we never overwrite index.html with empty data")

    with open(INDEX_HTML_PATH, encoding="utf-8") as f:
        html = f.read()

    html = replace_js_const(html, "DEFAULT_PRODUCERS", json.dumps(producers, ensure_ascii=False))
    html = replace_js_const(html, "DEFAULT_LABOR", json.dumps(labor, ensure_ascii=False))
    html = replace_js_const(html, "DEFAULT_QC_HISTORY", json.dumps(qc_history, ensure_ascii=False))
    html = replace_js_const(html, "DEFAULT_ENGAGEMENT_DATA", json.dumps(engagement, ensure_ascii=False))

    version_stamp = run_date.isoformat().replace("+00:00", "Z")
    version_pattern = re.compile(r'const DATA_VERSION\s*=\s*"[^"]*";')
    html, count = version_pattern.subn(f'const DATA_VERSION="{version_stamp}";', html, count=1)
    if count != 1:
        die("Could not find DATA_VERSION constant in index.html")

    with open(INDEX_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    if sync_errors:
        print(f"::warning::{len(sync_errors)} QC_Records row(s) could not be matched to an active roster entry:")
        for e in sync_errors:
            print(f"  - {e}")
        write_sync_errors(sh, sync_errors)
    else:
        write_sync_errors(sh, [])

    print(f"Synced {len(producers)} active producers, {len(labor)} labor-cost entries, "
          f"{len(qc_history)} QC month(s), engagement for {len(engagement)} people. "
          f"DATA_VERSION={version_stamp}")


if __name__ == "__main__":
    main()
