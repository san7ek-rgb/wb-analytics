#!/usr/bin/env python3
"""Collect seller funnel metrics from the authenticated WB portal session."""
import json
import math
import os
import sys
import time
import websocket
import fcntl
import sqlite3
import hashlib
import uuid
from pathlib import Path
from datetime import datetime, timezone
from http.cookies import SimpleCookie
from email.utils import parsedate_to_datetime
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


class HistoryStore:
    """Private day-grain archive; range reports are stored separately."""
    FIELDS = ('impressions', 'clicks', 'carts', 'orders', 'buyouts', 'revenue',
              'cartRate', 'orderRate', 'buyoutRate')

    def __init__(self, folder=None):
        self.folder = Path(folder) if folder else Path(__file__).resolve().parent / 'history'
        self.folder.mkdir(mode=0o700, parents=True, exist_ok=True)
        self.folder.chmod(0o700)
        self.path = self.folder / 'history.sqlite3'
        self.db = sqlite3.connect(self.path, timeout=20)
        self.path.chmod(0o600)
        version=self.db.execute('PRAGMA user_version').fetchone()[0]
        if version > 1:
            self.db.close()
            raise ValueError('History schema newer than this collector')
        self.db.execute('PRAGMA journal_mode=WAL')
        self.db.execute('PRAGMA foreign_keys=ON')
        if version == 0:
            self.db.executescript('''
          BEGIN;
          CREATE TABLE IF NOT EXISTS history_meta (key TEXT PRIMARY KEY, value TEXT NOT NULL);
          CREATE TABLE IF NOT EXISTS products (
            shop_key TEXT NOT NULL, article_id TEXT NOT NULL, title TEXT NOT NULL DEFAULT '',
            product_type TEXT NOT NULL DEFAULT '', first_seen TEXT NOT NULL, last_seen TEXT NOT NULL,
            PRIMARY KEY(shop_key,article_id));
          CREATE TABLE IF NOT EXISTS collection_runs (
            id TEXT PRIMARY KEY, shop_key TEXT NOT NULL, period_start TEXT NOT NULL,
            period_end TEXT NOT NULL, observed_at TEXT NOT NULL, requested INTEGER NOT NULL,
            received INTEGER NOT NULL, grain TEXT NOT NULL);
          CREATE TABLE IF NOT EXISTS daily_funnel (
            shop_key TEXT NOT NULL, article_id TEXT NOT NULL, day TEXT NOT NULL,
            impressions INTEGER, clicks INTEGER, carts INTEGER,
            orders INTEGER, buyouts INTEGER, revenue REAL,
            cart_rate REAL, order_rate REAL, buyout_rate REAL, observed_at TEXT NOT NULL,
            run_id TEXT NOT NULL REFERENCES collection_runs(id),
            PRIMARY KEY(shop_key,article_id,day));
          CREATE TABLE IF NOT EXISTS daily_coverage (
            shop_key TEXT NOT NULL, article_id TEXT NOT NULL, day TEXT NOT NULL,
            status TEXT NOT NULL CHECK(status IN ('ok','partial','missing','unavailable')), attempted_at TEXT NOT NULL,
            PRIMARY KEY(shop_key,article_id,day));
          CREATE TABLE IF NOT EXISTS period_reports (
            shop_key TEXT NOT NULL, article_id TEXT NOT NULL, period_start TEXT NOT NULL,
            period_end TEXT NOT NULL, impressions INTEGER, clicks INTEGER,
            carts INTEGER, orders INTEGER, buyouts INTEGER,
            revenue REAL, cart_rate REAL, order_rate REAL, buyout_rate REAL,
            observed_at TEXT NOT NULL, run_id TEXT NOT NULL REFERENCES collection_runs(id),
            PRIMARY KEY(shop_key,article_id,period_start,period_end));
          CREATE TABLE IF NOT EXISTS cost_observations (
            shop_key TEXT NOT NULL, article_id TEXT NOT NULL, observed_at TEXT NOT NULL,
            amount REAL NOT NULL CHECK(amount>=0), source TEXT NOT NULL,
            PRIMARY KEY(shop_key,article_id,observed_at));
          CREATE INDEX IF NOT EXISTS daily_day_idx ON daily_funnel(shop_key,day);
          CREATE INDEX IF NOT EXISTS cost_latest_idx ON cost_observations(shop_key,article_id,observed_at);
          PRAGMA user_version=1;
          COMMIT;
        ''')

    def close(self):
        self.db.close()

    def meta(self, key, value=None):
        if value is not None:
            self.db.execute('INSERT INTO history_meta VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value', (key, value))
            self.db.commit()
        row = self.db.execute('SELECT value FROM history_meta WHERE key=?', (key,)).fetchone()
        return row[0] if row else None

    def record(self, shop, start, end, requested, rows, products=None, unavailable=False):
        if not shop or not start or not end or start > end or date.fromisoformat(start).isoformat()!=start or date.fromisoformat(end).isoformat()!=end or len(set(requested)) != len(requested) or any(not isinstance(i,str) or not i.isdigit() or len(i)>16 for i in requested):
            raise ValueError('Invalid history scope')
        if any(r['id'] not in requested for r in rows) or len({r['id'] for r in rows}) != len(rows):
            raise ValueError('History article mismatch')
        now = datetime.now(timezone.utc).isoformat(timespec='milliseconds')
        today = datetime.now(ZoneInfo('Europe/Moscow')).date().isoformat()
        daily = start == end and end < today
        run = uuid.uuid4().hex
        metadata = {str(p['id']): p for p in (products or [])}
        returned = {r['id']:r for r in rows}
        with self.db:
            self.db.execute('INSERT INTO collection_runs VALUES (?,?,?,?,?,?,?,?)',
                            (run,shop,start,end,now,len(requested),len(rows),'day' if daily else 'period'))
            for article in requested:
                product = metadata.get(article, {})
                self.db.execute('''INSERT INTO products VALUES (?,?,?,?,?,?)
                    ON CONFLICT(shop_key,article_id) DO UPDATE SET
                    title=CASE WHEN excluded.title<>'' THEN excluded.title ELSE products.title END,
                    product_type=CASE WHEN excluded.product_type<>'' THEN excluded.product_type ELSE products.product_type END,
                    last_seen=excluded.last_seen''',
                    (shop,article,str(product.get('name','')),str(product.get('type','')),now,now))
                if daily:
                    coverage='unavailable' if unavailable else 'missing' if article not in returned else 'partial' if any(returned[article].get(k) is None for k in self.FIELDS[:5]) else 'ok'
                    self.db.execute('''INSERT INTO daily_coverage VALUES (?,?,?,?,?)
                      ON CONFLICT(shop_key,article_id,day) DO UPDATE SET status=excluded.status,attempted_at=excluded.attempted_at''',
                      (shop,article,start,coverage,now))
                amount = product.get('cost')
                if isinstance(amount,(int,float)) and not isinstance(amount,bool) and math.isfinite(amount) and amount>=0:
                    prior = self.db.execute('SELECT amount FROM cost_observations WHERE shop_key=? AND article_id=? ORDER BY observed_at DESC LIMIT 1',(shop,article)).fetchone()
                    if not prior or prior[0] != amount:
                        self.db.execute('INSERT INTO cost_observations VALUES (?,?,?,?,?)',(shop,article,datetime.now(timezone.utc).isoformat(),amount,'sheet-observed'))
            for row in rows:
                numbers = [row.get(field) for field in self.FIELDS]
                if daily:
                    self.db.execute('''INSERT INTO daily_funnel VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                      ON CONFLICT(shop_key,article_id,day) DO UPDATE SET
                      impressions=COALESCE(excluded.impressions,daily_funnel.impressions),clicks=COALESCE(excluded.clicks,daily_funnel.clicks),carts=COALESCE(excluded.carts,daily_funnel.carts),
                      orders=COALESCE(excluded.orders,daily_funnel.orders),buyouts=COALESCE(excluded.buyouts,daily_funnel.buyouts),revenue=COALESCE(excluded.revenue,daily_funnel.revenue),
                      cart_rate=COALESCE(excluded.cart_rate,daily_funnel.cart_rate),order_rate=COALESCE(excluded.order_rate,daily_funnel.order_rate),buyout_rate=COALESCE(excluded.buyout_rate,daily_funnel.buyout_rate),
                      observed_at=excluded.observed_at,run_id=excluded.run_id''', (shop,row['id'],start,*numbers,now,run))
                else:
                    self.db.execute('''INSERT INTO period_reports VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                      ON CONFLICT(shop_key,article_id,period_start,period_end) DO UPDATE SET
                      impressions=excluded.impressions,clicks=excluded.clicks,carts=excluded.carts,
                      orders=excluded.orders,buyouts=excluded.buyouts,revenue=excluded.revenue,
                      cart_rate=excluded.cart_rate,order_rate=excluded.order_rate,buyout_rate=excluded.buyout_rate,
                      observed_at=excluded.observed_at,run_id=excluded.run_id''', (shop,row['id'],start,end,*numbers,now,run))
            self.db.execute('INSERT INTO history_meta VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',('active_shop',shop))
            self.db.execute('INSERT INTO history_meta VALUES (?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value',('updated_at',now))

    def covered(self, shop, day, ids):
        if not ids:
            return False
        query="SELECT COUNT(*) FROM daily_coverage WHERE shop_key=? AND day=? AND status='ok' AND article_id IN ("+','.join('?' for _ in ids)+")"
        return self.db.execute(query,(shop,day,*ids)).fetchone()[0]==len(ids)

    def range_rows(self, start, end, ids):
        shop=self.meta('active_shop')
        if not shop or not ids:return []
        days=(date.fromisoformat(end)-date.fromisoformat(start)).days+1
        if days<=0:raise ValueError('Invalid history dates')
        placeholders=','.join('?' for _ in ids)
        query='''SELECT d.article_id,COUNT(*),SUM(d.impressions),SUM(d.clicks),SUM(d.carts),
          SUM(d.orders),SUM(d.buyouts),CASE WHEN COUNT(d.revenue)=COUNT(*) THEN SUM(d.revenue) ELSE NULL END
          FROM daily_funnel d JOIN daily_coverage c ON d.shop_key=c.shop_key AND d.article_id=c.article_id AND d.day=c.day
          WHERE d.shop_key=? AND d.day BETWEEN ? AND ? AND c.status='ok' AND d.article_id IN ('''+placeholders+''') GROUP BY d.article_id'''
        rows=[]
        for result in self.db.execute(query,(shop,start,end,*ids)):
            article,covered,impressions,clicks,carts,orders,buyouts,revenue=result
            if covered!=days or clicks>impressions:continue
            rows.append({'id':article,'impressions':impressions,'clicks':clicks,'carts':carts,'orders':orders,'buyouts':buyouts,
                         'revenue':revenue,'ctr':clicks/impressions if impressions else 0,
                         'cartRate':None,'orderRate':None,'buyoutRate':None})
        return rows

    def backup_if_due(self, force=False):
        if not self.db.execute('SELECT 1 FROM collection_runs LIMIT 1').fetchone():
            return
        folder = self.folder / 'backups'; folder.mkdir(mode=0o700,exist_ok=True);folder.chmod(0o700)
        day = datetime.now(ZoneInfo('Europe/Moscow')).date().isoformat()
        target = folder / ('history-'+day+'.sqlite3')
        if target.exists() and not force:
            return
        temporary = folder / '.backup-new.sqlite3'
        try:
            destination = sqlite3.connect(temporary)
            try:
                self.db.backup(destination)
            finally:
                destination.close()
            temporary.chmod(0o600);os.replace(temporary,target)
            self.meta('backup_at',datetime.now(timezone.utc).isoformat(timespec='milliseconds'))
            for old in sorted(folder.glob('history-????-??-??.sqlite3'))[:-7]:
                old.unlink()
        finally:
            temporary.unlink(missing_ok=True)

    def summary(self):
        shop = self.meta('active_shop')
        if not shop:
            return {'dailyRows':0,'days':0,'articles':0,'periodRows':0,'missingRows':0,'firstDay':'','lastDay':'','updatedAt':'','backupAt':'','error':self.meta('last_error') or ''}
        row = self.db.execute('SELECT COUNT(*),COUNT(DISTINCT day),COUNT(DISTINCT article_id),MIN(day),MAX(day) FROM daily_funnel WHERE shop_key=?',(shop,)).fetchone()
        periods = self.db.execute('SELECT COUNT(*) FROM period_reports WHERE shop_key=?',(shop,)).fetchone()[0]
        missing = self.db.execute("SELECT COUNT(*) FROM daily_coverage WHERE shop_key=? AND status<>'ok'",(shop,)).fetchone()[0]
        unavailable = self.db.execute("SELECT COUNT(DISTINCT day) FROM daily_coverage WHERE shop_key=? AND status='unavailable'",(shop,)).fetchone()[0]
        return {'dailyRows':row[0],'days':row[1],'articles':row[2],'firstDay':row[3] or '', 'lastDay':row[4] or '',
                'periodRows':periods,'missingRows':missing,'unavailableDays':unavailable,'updatedAt':self.meta('updated_at') or '', 'backupAt':self.meta('backup_at') or '', 'error':self.meta('last_error') or ''}


def history_shop_key(request):
    cookies = {}
    for c in request.get('cookies',[]):
        value=c.get('value','');cookies[c['name']]=value.get('value','') if isinstance(value,dict) else value
    headers={}
    for h in request.get('headers',[]):
        value=h.get('value','');headers[h['name'].lower()]=value.get('value','') if isinstance(value,dict) else value
    if 'x-supplier-id-external' not in cookies and headers.get('cookie'):
        jar=SimpleCookie()
        try:
            jar.load(headers['cookie'])
            cookies.update({key:morsel.value for key,morsel in jar.items()})
        except Exception:
            pass
    supplier=cookies.get('x-supplier-id-external') or headers.get('x-supplier-id-external') or headers.get('x-supplier-id')
    return hashlib.sha256(('wb-supplier:'+supplier).encode()).hexdigest() if isinstance(supplier,str) and supplier else None

FETCH_SCRIPT = r"""
const [url, payload, headers] = arguments;
const done = arguments[arguments.length - 1];
fetch(url, {method:'POST', credentials:'include',
  headers:headers, body:JSON.stringify(payload)})
  .then(async r => done({status:r.status, text:await r.text(), retryAfter:r.headers.get('Retry-After') || ''}))
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


class HistoryRetry(Exception):
    def __init__(self, seconds=60):
        self.seconds=seconds

class HistoryUnavailable(Exception):
    pass


def history_retry_seconds(response):
    value=str(response.get('retryAfter') or '').strip()
    try:
        seconds=int(value)
    except ValueError:
        try:
            when=parsedate_to_datetime(value)
            if when.tzinfo is None:when=when.replace(tzinfo=timezone.utc)
            seconds=int((when-datetime.now(timezone.utc)).total_seconds())
        except (TypeError,ValueError,OverflowError):
            seconds=60
    return min(86400,max(15,seconds))


def history_rows(root):
    data=root.get('data') if isinstance(root,dict) else None
    if not isinstance(data,dict) or not isinstance(data.get('groups'),list):
        raise HistoryUnavailable('Product groups unavailable')
    found={};conflicts=set()
    fields={'impressions':('viewCount','impressions'),'clicks':('openCard','openCardCount'),
            'carts':('addToCart','addToCartCount'),'orders':('orders','orderCount'),'buyouts':('buyoutCount',)}
    for group in data['groups']:
        for item in group.get('itemsGroup',[]) if isinstance(group,dict) else []:
            if not isinstance(item,dict):continue
            raw=item.get('nmId');identity=str(raw) if isinstance(raw,(str,int)) and not isinstance(raw,bool) else ''
            if not identity.isdigit() or len(identity)>16:continue
            row={'id':identity,**{key:report_number(item,names,True) for key,names in fields.items()},'revenue':report_number(item,('buyoutSum',))}
            for key,field in (('cartRate','openToCart'),('orderRate','cartToOrder'),('buyoutRate','buyoutPercent')):
                value=report_number(item,(field,));row[key]=value/100 if value is not None and value<=100 else None
            if any(row.get(key) is not None for key in HistoryStore.FIELDS):
                if identity in found and found[identity]!=row:conflicts.add(identity)
                else:found[identity]=row
    for identity in conflicts:found.pop(identity,None)
    return list(found.values())


def get_portal_data(driver, start, end, target_ids, captured_headers=None, history_mode=False):
    if captured_headers is None:
        ws = websocket.create_connection(driver.capabilities["webSocketUrl"], timeout=45, suppress_origin=True)
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
                driver._wb_history_shop_key = history_shop_key(req)
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
                        if history_mode and live_status in (429,500,502,503,504):
                            raise HistoryRetry(60)
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
    driver._wb_report_headers=captured_headers
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
            if history_mode and status in (429,500,502,503,504):
                raise HistoryRetry(history_retry_seconds(response))
            if history_mode and status in (400,422):
                raise HistoryUnavailable('WB declined this day')
            fail(f"WB funnel request failed (HTTP {status}). No spreadsheet data was sent.")
        try:
            data = json.loads(response["text"])
        except (KeyError, json.JSONDecodeError):
            fail("WB funnel response was not JSON. No spreadsheet data was sent.")
        if data.get("error"):
            if history_mode:raise HistoryUnavailable('WB reported an error for this day')
            fail("WB reported an error in a product batch. No spreadsheet data was sent.")
        page = history_rows(data) if history_mode else flatten_rows(data)
        if any(row["id"] not in batch for row in page):
            fail("WB did not respect the article filter. No spreadsheet data was sent.")
        result.extend(page)
        print(f"Batch {batch_start // 25 + 1}: requested {len(batch)}, received {len(page)}", flush=True)
        if batch_start + 25 < len(target_ids):
            time.sleep(3)
    if not result and not history_mode:
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

def get_catalog_info():
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
    return catalog

def get_catalog():
    return get_catalog_info()["ids"]

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


def publish_history_status(store, error=None):
    if error is not None:
        store.meta('last_error',error)
    summary=store.summary()
    digest=hashlib.sha256(json.dumps(summary,sort_keys=True).encode()).hexdigest()
    if store.meta('reported_summary')==digest:
        return
    try:
        receiver_action('historyStatus',summary=summary)
        store.meta('reported_summary',digest)
    except (Exception,SystemExit):
        print('History status is not yet available in Apps Script; archive remains local.',flush=True)


def prepare_history():
    store=None
    try:
        store=HistoryStore()
        publish_history_status(store)
    except Exception as error:
        print(f'History initialization failed ({type(error).__name__}); collection is not blocked.',flush=True)
    finally:
        if store:
            store.close()


def save_history(shop, start, end, ids, rows, products):
    store=None
    try:
        store=HistoryStore()
        if not shop:
            publish_history_status(store,'Не удалось определить магазин: история этого запуска не сохранена.')
            print('History not saved: supplier identity unavailable.',flush=True)
            return
        store.record(shop,start,end,ids,rows,products)
        store.backup_if_due(force=start==end)
        publish_history_status(store,'')
        print(f'History saved: {len(rows)} article records; grain={"day" if start==end and end<datetime.now(ZoneInfo("Europe/Moscow")).date().isoformat() else "period"}.',flush=True)
    except Exception as error:
        print(f'History write failed ({type(error).__name__}); spreadsheet collection continues.',flush=True)
        if store:
            publish_history_status(store,'Ошибка записи истории или резервной копии; проверьте журнал сервера.')
    finally:
        if store:
            store.close()


def report_job_error(job, message):
    if job:
        try:
            receiver_action("historyError" if job.get('kind')=='history' else "jobError", jobId=job["id"], message=message[:250])
        except (Exception, SystemExit):
            print("Could not update comparison error status; inspect server journal.", flush=True)


def import_history(job):
    store=None;driver=None;cursor=job.get('cursor',job['start'])
    streak=int(job.get('unavailableStreak',0));headers=None
    try:
        store=HistoryStore();catalog=get_catalog_info();ids=job['ids']
        if not ids or len(ids)>300 or cursor<job['start'] or cursor>job['end']:
            fail('Invalid history import range or article list.')
        options=Options();options.set_capability('webSocketUrl',True);options.binary_location='/usr/bin/firefox'
        for value in ('--headless','--no-remote','-profile',os.environ.get('WB_FIREFOX_PROFILE','/home/reportdeploy/wb-automation/profile')):options.add_argument(value)
        driver=webdriver.Firefox(options=options);driver.set_page_load_timeout(60);driver.set_script_timeout(90)
        deadline=time.monotonic()+90
        while cursor<=job['end']:
            unavailable=False;shop=getattr(driver,'_wb_history_shop_key',None)
            covered=bool(shop and store.covered(shop,cursor,ids))
            if not covered:
                try:
                    rows=get_portal_data(driver,cursor,cursor,ids,captured_headers=headers,history_mode=True)
                    streak=0
                except HistoryUnavailable:
                    rows=[];unavailable=True;streak+=1
                headers=getattr(driver,'_wb_report_headers',None);shop=getattr(driver,'_wb_history_shop_key',None)
                if not shop:fail('Cannot identify the WB shop; history was not saved.')
                store.record(shop,cursor,cursor,ids,rows,catalog.get('products',[]),unavailable=unavailable)
                store.meta('last_error','')
            cursor=(date.fromisoformat(cursor)+timedelta(days=1)).isoformat()
            progress=receiver_action('historyProgress',jobId=job['id'],cursor=cursor,unavailableStreak=streak,summary=store.summary())
            if not progress.get('continue'):return
            if progress.get('comparisonWaiting') or time.monotonic()>=deadline:
                break
            if not covered:time.sleep(3)
        receiver_action('historyProgress',jobId=job['id'],cursor=cursor,unavailableStreak=streak,
                        **{'yield':True},delaySeconds=20)
    except HistoryRetry as error:
        delay=min(86400,max(error.seconds,int(job.get('lastRetryDelay',0))*2))
        receiver_action('historyProgress',jobId=job['id'],cursor=cursor,unavailableStreak=streak,
                        **{'yield':True},delaySeconds=delay,retryDelay=delay)
        print(f'History paused for WB retry: {delay} seconds; cursor saved.',flush=True)
    except SystemExit as error:
        report_job_error(job,str(error));raise SystemExit(1)
    except Exception as error:
        report_job_error(job,f'History import failed ({type(error).__name__}); saved days are retained.')
        raise SystemExit(1)
    finally:
        if driver:
            try:driver.quit()
            except Exception:pass
        if store:
            try:
                store.backup_if_due(force=True);publish_history_status(store)
            except Exception as error:
                print(f'History backup/status failed ({type(error).__name__}).',flush=True)
            store.close()


def main(job=None):
    if job and job.get('kind')=='history':
        return import_history(job)
    if job and job.get('source')=='history':
        store=None
        try:
            store=HistoryStore();rows=store.range_rows(job['start'],job['end'],job['ids'])
            if not rows:fail('В архиве нет полной дневной истории выбранных товаров за весь период. Пропуски не заменены нулями.')
            send_to_sheet(job['start'],job['end'],rows,job['id'])
            print(f'Compared {len(rows)} articles from complete daily history; no WB request.',flush=True)
        except SystemExit as error:
            report_job_error(job,str(error));raise SystemExit(1)
        except Exception as error:
            report_job_error(job,f'History comparison failed ({type(error).__name__}).');raise SystemExit(1)
        finally:
            if store:store.close()
        return
    profile = os.environ.get("WB_FIREFOX_PROFILE", "/home/reportdeploy/wb-automation/profile")
    if not os.path.isdir(profile):
        report_job_error(job, "Dedicated Firefox profile directory is missing.")
        fail("Dedicated Firefox profile directory is missing.")
    from datetime import datetime
    day = (datetime.now(ZoneInfo("Europe/Moscow")).date() - timedelta(days=1)).isoformat()
    start, end = (job["start"], job["end"]) if job else (day, day)
    try:
        catalog = get_catalog_info()
    except SystemExit as error:
        report_job_error(job,str(error))
        raise SystemExit(1)
    target_ids = job["ids"] if job else catalog['ids']
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
    shop = None
    try:
        print(f"Collecting {len(target_ids)} articles: {start} — {end}", flush=True)
        driver = webdriver.Firefox(options=options)
        driver.set_page_load_timeout(60)
        driver.set_script_timeout(90)
        rows = get_portal_data(driver, start, end, target_ids)
        shop = getattr(driver,"_wb_history_shop_key",None)
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
    save_history(shop,start,end,target_ids,rows,catalog.get("products",[]))
    try:
        send_to_sheet(start, end, rows, job["id"] if job else None)
    except SystemExit as error:
        report_job_error(job, str(error))
        raise SystemExit(1)


def run_service():
    prepare_history()
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
            job = receiver_action("task", waitSeconds=20, historyWorker=1).get("job")
            if not job:
                return
        try:
            fcntl.flock(collector_lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            if job:
                receiver_action("historyDefer" if job.get('kind')=='history' else "jobDefer", jobId=job["id"])
            print("Another collection owns the profile; next timer will retry.", flush=True)
            if '--poll' not in sys.argv:
                raise SystemExit(75)
            return
        main(job)


if __name__ == "__main__":
    if "--self-check" in sys.argv:
        print("WB collector imports OK", flush=True)
        raise SystemExit(0)
    run_service()
