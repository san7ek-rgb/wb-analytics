#!/usr/bin/env python3
"""Collect seller funnel metrics from the authenticated WB portal session."""
import json
import os
import sys
import time
import websocket
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
    raise SystemExit(1)

def find_direct_id(obj):
    if not isinstance(obj, dict):
        return None
    keys = ("nmId", "nmID", "nm_id", "nmid", "articleId", "vendorCode")
    for key in keys:
        value = obj.get(key)
        if isinstance(value, (str, int)) and str(value).isdigit() and 1 <= len(str(value)) <= 16:
            return str(value)
    for key in ("product", "card", "nm", "item"):
        nested = obj.get(key)
        if isinstance(nested, dict):
            result = find_direct_id(nested)
            if result:
                return result
    return None

def find_metric(obj, names):
    wanted = {name.lower() for name in names}
    stack = [obj]
    while stack:
        value = stack.pop()
        if isinstance(value, dict):
            for key, child in value.items():
                if key.lower() in wanted and isinstance(child, (int, float)) and not isinstance(child, bool):
                    return float(child)
                if key.lower() in wanted and isinstance(child, dict):
                    current = child.get("current")
                    if isinstance(current, (int, float)) and not isinstance(current, bool):
                        return float(current)
                if isinstance(child, (dict, list)):
                    stack.append(child)
        elif isinstance(value, list):
            stack.extend(value)
    return None

def flatten_rows(root):
    found = {}
    stack = [root]
    while stack:
        item = stack.pop()
        if isinstance(item, list):
            stack.extend(reversed(item))
            continue
        if not isinstance(item, dict):
            continue
        product_id = find_direct_id(item)
        has_child_array = any(isinstance(v, list) for v in item.values())
        if product_id and not has_child_array:
            metrics = {
                "impressions": find_metric(item, ("viewCount", "views", "impressions", "impressionCount", "showCount")),
                "ctr": find_metric(item, ("viewToOpenConversion", "ctr", "ctrPercent", "clickThroughRate")),
                "clicks": find_metric(item, ("openCard", "openCardCount", "openCount", "clicks", "clickCount")),
                "carts": find_metric(item, ("addToCart", "addToCartCount", "cartCount", "addToCartCountAll")),
                "orders": find_metric(item, ("orderCount", "ordersCount", "orders")),
                "buyouts": find_metric(item, ("buyoutCount", "buyoutsCount", "buyouts")),
            }
            if metrics["impressions"] is not None:
                found[product_id] = metrics
        for value in item.values():
            if isinstance(value, (dict, list)):
                stack.append(value)
    rows = []
    for product_id, m in found.items():
        impressions = int(m["impressions"] or 0)
        clicks = int(m["clicks"] or 0)
        if any(m[k] is None for k in ("clicks", "carts", "orders", "buyouts")):
            fail("A product is missing required funnel fields; nothing was sent.")
        ctr = clicks / impressions if impressions else 0
        if ctr is None:
            ctr = clicks / impressions if impressions else 0
        elif ctr > 1:
            ctr /= 100
        rows.append({
            "id": product_id,
            "impressions": impressions,
            "ctr": min(max(float(ctr), 0), 1),
            "clicks": clicks,
            "carts": int(m["carts"] or 0),
            "orders": int(m["orders"] or 0),
            "buyouts": int(m["buyouts"] or 0),
        })
    return rows

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
            "prevPeriod": {"start": (date.fromisoformat(start) - timedelta(days=1)).isoformat(),
                           "end": (date.fromisoformat(end) - timedelta(days=1)).isoformat()},
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
        fail("No product rows with impressions found in WB response. Nothing was sent.")
    dedup = {row["id"]: row for row in result}
    return list(dedup.values())

def send_to_sheet(start, end, rows):
    token = os.environ.get("WB_COST_RECEIVER_TOKEN", "").strip()
    if len(token) < 32:
        fail("Receiver secret is not configured. No spreadsheet data was sent.")
    body = json.dumps({"schema": 1, "action": "funnel", "token": token,
                       "start": start, "end": end, "rows": rows}).encode()
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

def main():
    profile = os.environ.get("WB_FIREFOX_PROFILE", "/home/reportdeploy/wb-automation/profile")
    if not os.path.isdir(profile):
        fail("Dedicated Firefox profile directory is missing.")
    # The server uses UTC; WB reporting period is Moscow calendar time.
    from datetime import datetime
    moscow_yesterday = datetime.now(ZoneInfo("Europe/Moscow")).date() - timedelta(days=1)
    day = moscow_yesterday.isoformat()
    options = Options()
    options.set_capability("webSocketUrl", True)
    options.binary_location = "/usr/bin/firefox"
    options.add_argument("--headless")
    options.add_argument("--no-remote")
    options.add_argument("-profile")
    options.add_argument(profile)
    target_ids = get_catalog()
    driver = None
    try:
        driver = webdriver.Firefox(options=options)
        driver.set_page_load_timeout(60)
        rows = get_portal_data(driver, day, day, target_ids)
        received_ids = {row["id"] for row in rows}
        print(f"Coverage: {len(received_ids)}/{len(target_ids)}; missing articles retain previous values", flush=True)
    except SystemExit:
        raise
    except Exception as exc:
        # Avoid printing driver/browser exceptions: they can include local paths.
        fail(f"Browser collection failed ({type(exc).__name__}). No spreadsheet data was sent.")
    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass
    send_to_sheet(day, day, rows)

if __name__ == "__main__":
    main()
