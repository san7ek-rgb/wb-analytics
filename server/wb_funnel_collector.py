#!/usr/bin/env python3
"""Collect seller funnel metrics from the authenticated WB portal session."""
import json
import os
import sys
import time
from datetime import date, timedelta
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

from selenium import webdriver
from selenium.webdriver.firefox.options import Options

PORTAL_URL = "https://seller.wildberries.ru/content-analytics/interactive-report/main"
DETAILS_URL = "https://seller-content.wildberries.ru/ns/analytics-api/content-analytics/api/v1/sales-funnel/report/details"
RECEIVER_URL = "https://script.google.com/macros/s/AKfycbyj_4_nwplUCYyEjDwHDOZqDhfSZjkpB57PBlhunkMOkpnPvAeTUozb1rxgrUciE2Dc/exec"
PAGE_SIZE = 100

FETCH_SCRIPT = r"""
const [url, payload] = arguments;
const done = arguments[arguments.length - 1];
fetch(url, {method:'POST', credentials:'include',
  headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)})
  .then(async r => done({status:r.status, text:await r.text()}))
  .catch(e => done({status:0, text:String(e)}));
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
                "clicks": find_metric(item, ("openCardCount", "openCount", "clicks", "clickCount")),
                "carts": find_metric(item, ("addToCartCount", "cartCount", "addToCartCountAll")),
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
        ctr = m["ctr"]
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

def get_portal_data(driver, start, end):
    driver.get(PORTAL_URL)
    time.sleep(8)
    if "login" in driver.current_url.lower() or "auth" in driver.current_url.lower():
        fail("WB session needs sign-in again. Open Firefox on the server and sign in.")
    result = []
    offset = 0
    for _ in range(20):
        payload = {
            "currentPeriod": {"start": start, "end": end},
            "subjects": [], "brands": [], "nms": [], "tagIds": [],
            "skipDeletedNm": True,
            "orderBy": {"field": "openCardCount", "mode": "desc"},
            "limit": PAGE_SIZE, "offset": offset,
        }
        response = driver.execute_async_script(FETCH_SCRIPT, DETAILS_URL, payload)
        if not response or response.get("status") != 200:
            status = response.get("status", "no response") if response else "no response"
            fail(f"WB funnel request failed (HTTP {status}). No spreadsheet data was sent.")
        try:
            data = json.loads(response["text"])
        except (KeyError, json.JSONDecodeError):
            fail("WB funnel response was not JSON. No spreadsheet data was sent.")
        page = flatten_rows(data)
        result.extend(page)
        if len(page) < PAGE_SIZE:
            break
        offset += PAGE_SIZE
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

def main():
    profile = os.environ.get("WB_FIREFOX_PROFILE", "/home/reportdeploy/wb-automation/profile")
    if not os.path.isdir(profile):
        fail("Dedicated Firefox profile directory is missing.")
    # The server uses UTC; WB reporting period is Moscow calendar time.
    from datetime import datetime
    moscow_yesterday = datetime.now(ZoneInfo("Europe/Moscow")).date() - timedelta(days=1)
    day = moscow_yesterday.isoformat()
    options = Options()
    options.binary_location = "/usr/bin/firefox"
    options.add_argument("--headless")
    options.add_argument("--no-remote")
    options.add_argument("-profile")
    options.add_argument(profile)
    driver = None
    try:
        driver = webdriver.Firefox(options=options)
        driver.set_page_load_timeout(60)
        rows = get_portal_data(driver, day, day)
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
