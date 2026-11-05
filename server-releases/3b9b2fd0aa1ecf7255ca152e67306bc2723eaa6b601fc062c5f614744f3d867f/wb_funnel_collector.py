#!/usr/bin/env python3
"""Collect seller funnel metrics from the authenticated WB portal session."""
import json
import math
import os
import sys
import time
import websocket
import fcntl
from datetime import date, timedelta
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

from selenium import webdriver
from selenium.webdriver.firefox.options import Options

PORTAL_URL = "https://seller.wildberries.ru/content-analytics/interactive-report/main"
DETAILS_URL = "https://seller-content.wildberries.ru/ns/analytics-api/content-analytics/api/v1/sales-funnel/report"
RECEIVER_URL = "https://script.google.com/macros/s/AKfycbyj_4_nwplUCYyEjDwHDOZqDhfSZjkpB57PBlhunkMOkpnPvAeTUozb1rxgrUciE2Dc/exec"
PAGE_SIZE = 100

FETCH_SCRIPT = r"""
const [url, payload, headers] = arguments;
const done = arguments[arguments.length - 1];
fetch(url, {method:'POST', credentials:'include',
  headers:headers, body:JSON.stringify(payload)})
  .then(async r => done({status:r.status, text:await r.text()}))
  .catch(() => done({status:0, text:'Browser fetch failed'}));
"""

def fail(message):
    print(message, file=sys.stderr, flush=True)
    raise SystemExit(message)

def report_number(item, names, integer=False):
    # Read product-level current values ONLY. Chart points and WB Club/b2b
    # subtotals must never replace the corresponding whole-product counters.
    for name in names:
        if name not in item:
            continue
        value = item[name]
        if isinstance(value, dict):
            value = value.get("current")
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            continue
        if not math.isfinite(value) or value < 0:
            continue
        if integer and int(value) != value:
            continue
        return int(value) if integer else float(value)
    return None


def flatten_rows(root):
    data = root.get("data") if isinstance(root, dict) else None
    if not isinstance(data, dict) or not isinstance(data.get("groups"), list):
        fail("Unexpected report structure: product groups missing. Nothing was sent.")
    found, conflicts = {}, set()
    skipped, missing_names = 0, set()
    for group in data["groups"]:
        if not isinstance(group, dict):
            continue
        items = group.get("itemsGroup", [])
        if not isinstance(items, list):
            fail("Unexpected product list structure. Nothing was sent.")
        for item in items:
            if not isinstance(item, dict):
                continue
            raw_id = item.get("nmId")
            product_id = str(raw_id) if isinstance(raw_id, (str, int)) and not isinstance(raw_id, bool) else ""
            if not product_id.isdigit() or not 1 <= len(product_id) <= 16:
                skipped += 1
                continue
            counts = {
                "impressions": report_number(item, ("viewCount", "impressions"), True),
                "clicks": report_number(item, ("openCard", "openCardCount", "openCount"), True),
                "carts": report_number(item, ("addToCart", "addToCartCount", "cartCount"), True),
                "orders": report_number(item, ("orders", "orderCount", "ordersCount"), True),
                "buyouts": report_number(item, ("buyoutCount", "buyoutsCount"), True),
            }
            missing = [key for key, value in counts.items() if value is None]
            if missing:
                skipped += 1
                missing_names.update(missing)
                continue
            impressions, clicks = counts["impressions"], counts["clicks"]
            if clicks > impressions:
                skipped += 1
                missing_names.add("comparable impressions/clicks")
                continue
            row = {"id": product_id, **counts,
                   "ctr": clicks / impressions if impressions else 0,
                   "revenue": report_number(item, ("buyoutSum",))}
            for key, field in (("cartRate", "openToCart"), ("orderRate", "cartToOrder"), ("buyoutRate", "buyoutPercent")):
                percent = report_number(item, (field,))
                row[key] = percent / 100 if percent is not None and percent <= 100 else None
            if product_id in found and found[product_id] != row:
                conflicts.add(product_id)
            else:
                found[product_id] = row
    for product_id in conflicts:
        found.pop(product_id, None)
    if skipped or conflicts:
        print(f"Excluded incomplete/conflicting report rows: {skipped + len(conflicts)}; unavailable fields: {', '.join(sorted(missing_names)) or 'conflicting duplicates'}. Previous sheet values retained.", flush=True)
    return list(found.values())


def get_portal_data(driver, start, end, target_ids):
    ws = websocket.create_connection(driver.capabilities["webSocketUrl"], timeout=45, suppress_origin=True)
    captured_headers = None
    try:
        ws.send(json.dumps({"id": 1, "method": "session.subscribe",
                            "params": {"events": ["network.beforeRequestSent", "network.responseCompleted"]}}))
        ack = json.loads(ws.recv())
        if ack.get("type") == "error":
            fail("Firefox network subscription failed. Nothing was sent.")
        driver.get(PORTAL_URL)
        deadline = time.monotonic() + 45
        while time.monotonic() < deadline:
            event = json.loads(ws.recv())
            req = event.get("params", {}).get("request", {})
            if req.get("method") != "POST" or req.get("url", "").split("?")[0] != DETAILS_URL:
                continue
            captured_headers = {}
            for header in req.get("headers", []):
                name = header["name"]
                value = header["value"]
                if isinstance(value, dict):
                    value = value.get("value", "")
                if not name.startswith(":") and name.lower() not in ("host", "content-length", "accept-encoding", "connection"):
                    captured_headers[name] = value
            if req.get("cookies"):
                def cookie_value(c):
                    v = c["value"]
                    return v.get("value", "") if isinstance(v, dict) else v
                captured_headers["Cookie"] = "; ".join(c["name"] + "=" + cookie_value(c) for c in req["cookies"])
            request_id = req.get("request")
            while time.monotonic() < deadline:
                completed = json.loads(ws.recv())
                params = completed.get("params", {})
                if completed.get("method") == "network.responseCompleted" and params.get("request", {}).get("request") == request_id:
                    live_status = params.get("response", {}).get("status")
                    print(f"Live Firefox report request: HTTP {live_status}", flush=True)
                    if live_status != 200:
                        fail("The automation Firefox profile itself is not receiving a successful report. Nothing was sent.")
                    break
            break
    except (websocket.WebSocketTimeoutException, KeyError):
        fail("No live report request was captured. Open the report in the automation profile and verify sign-in.")
    finally:
        ws.close()
    if not captured_headers:
        fail("No live report request was captured. Nothing was sent.")
    result = []
    for batch_start in range(0, len(target_ids), 25):
        batch = target_ids[batch_start:batch_start + 25]
        payload = {
            "currentPeriod": {"start": start, "end": end},
            "subjects": [], "brands": [], "nms": [int(i) for i in batch], "tagIds": [],
            "skipDeletedNm": False,
            "prevPeriod": {"start": (date.fromisoformat(start) - timedelta(days=(date.fromisoformat(end)-date.fromisoformat(start)).days+1)).isoformat(),
                           "end": (date.fromisoformat(start) - timedelta(days=1)).isoformat()},
            "orderBy": {"field": "orders", "mode": "desc"},
        }
        headers = {k: v for k, v in captured_headers.items()
                   if k.lower() not in ("cookie", "origin", "referer", "user-agent", "content-length", "host", "connection", "accept-encoding")
                   and not k.lower().startswith("sec-")}
        headers["Content-Type"] = "application/json"
        response = driver.execute_async_script(FETCH_SCRIPT, DETAILS_URL, payload, headers)
        if not response or response.get("status") != 200:
            status = response.get("status", "no response") if response else "no response"
            fail(f"WB funnel request failed (HTTP {status}). No spreadsheet data was sent.")
        try:
            data = json.loads(response["text"])
        except (KeyError, json.JSONDecodeError):
            fail("WB funnel response was not JSON. No spreadsheet data was sent.")
        if data.get("error"):
            fail("WB reported an error in a product batch. No spreadsheet data was sent.")
        page = flatten_rows(data)
        if any(row["id"] not in batch for row in page):
            fail("WB did not respect the article filter. No spreadsheet data was sent.")
        result.extend(page)
        print(f"Batch {batch_start // 25 + 1}: requested {len(batch)}, received {len(page)}", flush=True)
        if batch_start + 25 < len(target_ids):
            time.sleep(3)
    if not result:
        fail("No complete product rows with comparable impressions and funnel counters for this period. Nothing was sent.")
    dedup = {row["id"]: row for row in result}
    return list(dedup.values())

def send_to_sheet(start, end, rows, job_id=None):
    token = os.environ.get("WB_COST_RECEIVER_TOKEN", "").strip()
    if len(token) < 32:
        fail("Receiver secret is not configured. No spreadsheet data was sent.")
    body = json.dumps({"schema": 1, "action": "funnel", "token": token,
                       "start": start, "end": end, "rows": rows, "jobId": job_id}).encode()
    request = Request(RECEIVER_URL, data=body, headers={"Content-Type": "application/json"}, method="POST")
    try:
        with urlopen(request, timeout=90) as response:
            result = json.loads(response.read().decode("utf-8"))
    except (HTTPError, URLError, TimeoutError, json.JSONDecodeError):
        fail("Sheet receiver request failed. Check Apps Script deployment and server network.")
    if not result.get("ok"):
        fail("Sheet receiver declined the update. Check its response in Apps Script execution log.")
    print(f"OK: {result.get('received', 0)} products sent; forecast cells updated: {result.get('updatedForecast', 0)}; source cells updated: {result.get('updatedSource', 0)}; period {start} — {end}", flush=True)

def get_catalog():
    token = os.environ.get("WB_COST_RECEIVER_TOKEN", "").strip()
    if len(token) < 32:
        fail("Receiver secret is not configured.")
    body = json.dumps({"schema": 1, "action": "catalog", "token": token}).encode()
    try:
        with urlopen(Request(RECEIVER_URL, data=body, headers={"Content-Type":"application/json"}), timeout=90) as r:
            catalog = json.loads(r.read().decode())
    except (HTTPError, URLError, TimeoutError, json.JSONDecodeError):
        fail("Cannot read product IDs from sheet receiver. Nothing was sent to WB.")
    ids = catalog.get("ids")
    if not catalog.get("ok") or not isinstance(ids, list) or not ids or len(ids) > 300:
        fail("Update the deployed Apps Script receiver: catalog action unavailable or sheet has no IDs.")
    if any(not isinstance(i, str) or not i.isdigit() or len(i) > 16 for i in ids):
        fail("Invalid product ID list. Nothing was sent to WB.")
    print(f"Catalog: {len(ids)} selected of {catalog.get('total', len(ids))}; limit 300", flush=True)
    return ids

def receiver_action(action, **values):
    token = os.environ.get("WB_COST_RECEIVER_TOKEN", "").strip()
    if len(token) < 32:
        fail("Receiver secret is not configured.")
    body = json.dumps({"schema": 1, "action": action, "token": token, **values}).encode()
    try:
        with urlopen(Request(RECEIVER_URL, data=body, headers={"Content-Type":"application/json"}), timeout=90) as r:
            result = json.loads(r.read().decode())
    except (HTTPError, URLError, TimeoutError, json.JSONDecodeError):
        fail("Comparison receiver request failed. Check the deployed Apps Script version.")
    if not result.get("ok"):
        fail("Comparison receiver declined the action. Check its deployment.")
    return result


def report_job_error(job, message):
    if job:
        try:
            receiver_action("jobError", jobId=job["id"], message=message[:250])
        except (Exception, SystemExit):
            print("Could not update comparison error status; inspect server journal.", flush=True)


def main(job=None):
    profile = os.environ.get("WB_FIREFOX_PROFILE", "/home/reportdeploy/wb-automation/profile")
    if not os.path.isdir(profile):
        report_job_error(job, "Dedicated Firefox profile directory is missing.")
        fail("Dedicated Firefox profile directory is missing.")
    from datetime import datetime
    day = (datetime.now(ZoneInfo("Europe/Moscow")).date() - timedelta(days=1)).isoformat()
    start, end = (job["start"], job["end"]) if job else (day, day)
    target_ids = job["ids"] if job else get_catalog()
    if len(target_ids) > 300 or not target_ids or any(not isinstance(i,str) or not i.isdigit() for i in target_ids):
        report_job_error(job, "Invalid comparison article list.")
        fail("Invalid comparison article list.")
    options = Options()
    options.set_capability("webSocketUrl", True)
    options.binary_location = "/usr/bin/firefox"
    options.add_argument("--headless")
    options.add_argument("--no-remote")
    options.add_argument("-profile")
    options.add_argument(profile)
    driver = None
    try:
        print(f"Collecting {len(target_ids)} articles: {start} — {end}", flush=True)
        driver = webdriver.Firefox(options=options)
        driver.set_page_load_timeout(60)
        driver.set_script_timeout(90)
        rows = get_portal_data(driver, start, end, target_ids)
        received_ids = {row["id"] for row in rows}
        print(f"Coverage: {len(received_ids)}/{len(target_ids)}; missing articles retain previous values", flush=True)
    except SystemExit as error:
        report_job_error(job, str(error))
        raise SystemExit(1)
    except Exception as error:
        message = f"Browser collection failed ({type(error).__name__}). No spreadsheet data was sent."
        report_job_error(job, message)
        fail(message)
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
    try:
        send_to_sheet(start, end, rows, job["id"] if job else None)
    except SystemExit as error:
        report_job_error(job, str(error))
        raise SystemExit(1)


def run_service():
    lock_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".collector.lock")
    with open(lock_path, "a") as collector_lock:
        job = None
        if "--poll" in sys.argv:
            # Do not claim a comparison while the daily collector is busy.
            try:
                fcntl.flock(collector_lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                return
            fcntl.flock(collector_lock, fcntl.LOCK_UN)
            # No profile lock, Firefox, or WB request while waiting for a command.
            job = receiver_action("task", waitSeconds=20).get("job")
            if not job:
                return
        try:
            fcntl.flock(collector_lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            if job:
                receiver_action("jobDefer", jobId=job["id"])
            print("Another collection owns the profile; next timer will retry.", flush=True)
            return
        main(job)


if __name__ == "__main__":
    if "--self-check" in sys.argv:
        print("WB collector imports OK", flush=True)
        raise SystemExit(0)
    run_service()
