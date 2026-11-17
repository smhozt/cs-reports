"""
Bot Goals — Comm100 AI Agent "Goals" raporunu (botun koyduğu taglar) Google Sheet'e yazar.

Kaynak : Reporting → AI Agent → Digital AI Agent → Goals  (Report API: reportingquery/reports/query)
Hedef  : 1) AYRI bir Google Sheet (GOALS_SPREADSHEET_ID), sekme "Goals_Günlük":
            [Tarih, Marka, Goal, Sayı]   (yalnızca sayısı > 0 olan goal'ler)
         2) (opsiyonel) elle tutulan "WRAP UP / BOT TAGS" Sheet'inin (MANUAL_SPREADSHEET_ID) bu haftaki
            sekmesi: BOT TAGS sütunları (B, G, L) Goals'tan, AGENT'S TAGS sütunları (C, H, M) Tag
            Sheet'indeki Haftalık sekmesinden (MEMBER_TAGS_SPREADSHEET_ID) doldurulur. % formüllerine
            ve TOTAL'e dokunulmaz. Yalnızca bu hafta ve geçen hafta güncellenir (eski haftalar korunur).

Kullanım:
  python bot_goals.py                       → dün
  python bot_goals.py 2026-11-02            → tek gün
  python bot_goals.py 2026-11-02 2026-11-08 → aralık (geriye dönük doldurma)
  python bot_goals.py probe [YYYY-MM-DD]    → yalnızca API'yi dener, ham cevabı yazdırır (Sheet'e dokunmaz)

Aynı gün+marka tekrar çalıştırılırsa eski satırlar silinip yenisi yazılır (idempotent).
"""

import os
import re
import sys
import json
import time
import base64
import smtplib
import traceback
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pytz
import requests

# ── CONFIG ───────────────────────────────────────────────────
SITE_ID = os.environ["COMM100_SITE_ID"]
API_KEY = os.environ["COMM100_API_KEY"]
COMM100_EMAIL = os.environ["COMM100_EMAIL"]
SPREADSHEET_ID = os.environ.get("GOALS_SPREADSHEET_ID", "")
GCP_CREDS_JSON = os.environ.get("GCP_CREDENTIALS", "")
MANUAL_SPREADSHEET_ID = os.environ.get("MANUAL_SPREADSHEET_ID", "")          # elle tutulan haftalık Sheet
TAGS_SPREADSHEET_ID = os.environ.get("MEMBER_TAGS_SPREADSHEET_ID", "")       # AGENT'S TAGS kaynağı
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_PASS = os.environ.get("GMAIL_APP_PASSWORD", "")
REPORT_EMAILS = [e.strip() for e in os.environ.get("REPORT_EMAILS", "").split(",") if e.strip()]

HOST = "https://dash15.lively-chat.com"
REPORT_URLS = [
    f"{HOST}/api/reportingquery/reports/query?siteId={SITE_ID}",      # tarayıcının kullandığı adres
    f"{HOST}/api/v4/reportingquery/reports/query?siteId={SITE_ID}",   # yedek (404 gelirse)
]
TZ = pytz.timezone("Europe/Sofia")
COMM100_TZ = "FLE Standard Time"        # raporun timezone alanı (Sofia/Kyiv/Helsinki)

# Marka → AI Agent (chatbot) ID. Gerekirse BOT_IDS_JSON ile ezilebilir: {"BS":"...","SB":"...","TB":"..."}
BOTS = {
    "BS": "ff1a7879-ff4d-4469-837a-065f6416f1b0",   # Betsat Canlı Yardım Asistanı
    "SB": "8298c1fa-c16d-40ef-9b3b-65dfd01609a2",   # Superbetin
    "TB": "6a2130f6-5067-44f5-b299-893ac062b14e",   # Turkbet
}
if os.environ.get("BOT_IDS_JSON"):
    BOTS.update(json.loads(os.environ["BOT_IDS_JSON"]))

# Goal adları. Raporda her goal ayrı bir sayaç alanı; botlarda goal listesi farklı olabildiği
# için üst küme kullanılıyor (olmayan goal 0 döner). Yeni bir goal eklenirse buraya ekle.
GOALS = [
    "ac_open_request", "affiliate", "betting_rules_query", "birthday_bonus_query",
    "boosted_odds_query", "cashout_issue", "casino_cashback_query", "casino_instant_cashback",
    "casino_issue", "contacts", "deposit_bonus_query", "deposit_check_query", "deposit_issue",
    "deposit_missing", "deposit_query", "docs_query", "fake_site", "fdb_query", "free_spin_query",
    "game_result_query", "game_rules_query", "general_info", "live_chat_ban", "login_request",
    "personal_data_change", "raf_bonus_query", "registration_query", "settlement_query",
    "sms_email_verification", "social_media_bonus", "Special Promotions", "sport_cashback_query",
    "sport_issue", "trial_bonus_query", "wd_cancellation", "wd_check", "wd_issue", "wd_missing",
    "wr_query",
    # eskiden botta olan, sonradan kaldırılmış goal'ler (geçmiş veri tam çıksın diye)
    "reload_bonus_query", "monthly_boost_query", "daily_reactivation_query", "combo_bonus_query",
]

TAB = "Goals_Günlük"
HEADER = ["Tarih", "Marka", "Goal", "Sayı"]
MAX_DAYS_PER_CALL = 31

# Elle tutulan Sheet'in haftalık sekme düzeni: marka → (BOT TAGS sütunu, AGENT'S TAGS sütunu)
# Sıra: Betsat | Turkbet | Superbetin  (TB/SB sırası casino_instant_cashback'e göre çıkarıldı)
MANUAL_BLOCKS = {"BS": ("B", "C"), "TB": ("G", "H"), "SB": ("L", "M")}
MANUAL_LABEL_COLS = ["B", "C", "G", "H", "L", "M", "Q", "R"]     # 2. satırdaki hafta etiketleri
MANUAL_FIRST_ROW = 3
MANUAL_NAME_COLS = ["A", "F", "K", "P"]                          # her blokta tag adı sütunu
# Goals raporundaki adı wrap-up tag adından farklı olanlar (manuel Sheet / Tag Sheet'teki ada çevrilir)
GOAL_TO_TAG = {"deposit_check_query": "deposit_check"}
# Bot tek "Special Promotions" goal'ü atar; agent'lar bunun yerine bu özel tag'leri kullanır → manuel Sheet'te
# "Special Promotions" satırında toplanır (bu tag'lerin kendi satırları boş bırakılır).
PROMO_ROW = "special promotions"
PROMO_TAGS = {"casino_private_bonus_query", "sport_private_bonus_query", "big_reactivation_query",
              "monthly_boost_query", "daily_reactivation_query", "taraftar_bonus_query",
              "calendar_bonus_query", "special_event_bonus_query", "lapsed_reactivation_query"}

def roll_promo(counts):
    """{tag: sayı} → PROMO_TAGS toplamı PROMO_ROW'a eklenir, tek tek tag'ler çıkarılır."""
    out, extra = {}, 0
    for k, n in counts.items():
        if k in PROMO_TAGS:
            extra += n
        else:
            out[k] = out.get(k, 0) + n
    if extra:
        out[PROMO_ROW] = out.get(PROMO_ROW, 0) + extra
    return out
MANUAL_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "June", "July", "Aug", "Sep", "Oct", "Nov", "Dec"]
TAGS_MONTHS3 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ── TARİH ────────────────────────────────────────────────────
def sofia_yesterday():
    return (datetime.now(TZ) - timedelta(days=1)).strftime("%Y-%m-%d")

def date_range(a, b):
    d0 = datetime.strptime(a, "%Y-%m-%d")
    d1 = datetime.strptime(b, "%Y-%m-%d")
    out = []
    while d0 <= d1:
        out.append(d0.strftime("%Y-%m-%d"))
        d0 += timedelta(days=1)
    return out


def week_start(date_str):
    d = datetime.strptime(date_str, "%Y-%m-%d")
    return (d - timedelta(days=d.weekday())).strftime("%Y-%m-%d")

def week_end(date_str):
    d = datetime.strptime(week_start(date_str), "%Y-%m-%d")
    return (d + timedelta(days=6)).strftime("%Y-%m-%d")

def tags_week_label(ws, we):
    """Tag Sheet'indeki Haftalık etiketi: '2 - 8 Nov 2026', '26 Oct - 1 Nov 2026'."""
    d0, d1 = datetime.strptime(ws, "%Y-%m-%d"), datetime.strptime(we, "%Y-%m-%d")
    m0, m1 = TAGS_MONTHS3[d0.month - 1], TAGS_MONTHS3[d1.month - 1]
    if d0.year == d1.year and d0.month == d1.month:
        return f"{d0.day} - {d1.day} {m1} {d1.year}"
    if d0.year == d1.year:
        return f"{d0.day} {m0} - {d1.day} {m1} {d1.year}"
    return f"{d0.day} {m0} {d0.year} - {d1.day} {m1} {d1.year}"

def manual_short_label(ws, we):
    """Elle tutulan Sheet'in 2. satır etiketi: '2-8 Nov', '26 Oct - 1 Nov', '29 Dec - 4 Jan'."""
    d0, d1 = datetime.strptime(ws, "%Y-%m-%d"), datetime.strptime(we, "%Y-%m-%d")
    m0, m1 = MANUAL_MONTHS[d0.month - 1], MANUAL_MONTHS[d1.month - 1]
    if d0.year == d1.year and d0.month == d1.month:
        return f"{d0.day}-{d1.day} {m1}"
    return f"{d0.day} {m0} - {d1.day} {m1}"

def manual_tab_name(ws, we):
    """'2-8 Nov 2026', '26 Oct - 1 Nov 2026' (yıl = haftanın bitiş yılı, mevcut sekmelerdeki gibi)."""
    return f"{manual_short_label(ws, we)} {datetime.strptime(we, '%Y-%m-%d').year}"

def norm_tab(name):
    return re.sub(r"\s+", "", str(name)).lower()

def norm_tag(name):
    return str(name or "").strip().lower()

# ── COMM100 REPORT API ───────────────────────────────────────
def auth_header():
    return "Basic " + base64.b64encode(f"{COMM100_EMAIL}:{API_KEY}".encode()).decode()

def field_name(goal):
    return goal.lower()          # raporun kendi adlandırması: "special promotions"

def build_payload(bot_id, d_from, d_to):
    """Goals sayfasının gönderdiği gövdenin aynısı; tek fark goal listesi ve tarih/bot."""
    fields = [{
        "name": "time", "valueType": "string", "calculationType": "originalValue",
        "conditionExpression": "", "conditionMatchType": "all", "conditions": [],
        "expression": "", "fieldName": "",
    }]
    for g in GOALS:
        fields.append({
            "name": field_name(g), "calculationType": "count",
            "conditionExpression": "currentGoal",
            "conditions": [{"name": "currentGoal", "fieldName": "Goal",
                            "operate": "equals", "values": [g]}],
            "expression": "", "fieldName": "Goal", "valueType": "int",
        })
    return {
        "cubeEntities": [{
            "name": "ChatbotGoalTriggerRecord",
            "fields": fields,
            "filters": [
                {"fieldName": "TriggeredTime", "matchType": "between",
                 "value": [d_from.replace("-", "/"), d_to.replace("-", "/")]},
                {"fieldName": "ChatbotId", "matchType": "equals", "value": [bot_id]},
            ],
            "rowGroups": [{"fieldName": "TriggeredTime", "isFull": True,
                           "name": "time", "timeDisplayType": "day"}],
        }],
        "mergeType": "column",
        "timezone": COMM100_TZ,
    }

def call_report(payload):
    """Report API çağrısı. Geçici hatalarda tekrar dener; 401/403/404'te anlamlı hata verir."""
    last = None
    for url in REPORT_URLS:
        for attempt in range(3):
            try:
                r = requests.post(
                    url, json=payload, timeout=90,
                    headers={"Authorization": auth_header(),
                             "Content-Type": "application/json",
                             "Accept": "application/json"})
            except Exception as e:
                last = f"{type(e).__name__}: {e}"
                time.sleep(2 * (attempt + 1))
                continue
            if r.status_code == 200:
                try:
                    return r.json()
                except Exception:
                    last = f"HTTP 200 ama JSON değil: {r.text[:200]}"
                    break
            last = f"HTTP {r.status_code}: {r.text[:300]}"
            if r.status_code == 404:
                break                       # yedek adrese geç
            if r.status_code in (401, 403):
                raise RuntimeError(
                    f"Report API yetki hatası ({last}). API anahtarının Reporting yetkisi yok "
                    f"veya Basic auth bu uçta kabul edilmiyor.")
            time.sleep(2 * (attempt + 1))
    raise RuntimeError(f"Report API çağrısı başarısız: {last}")

def parse_series(resp):
    """Cevap → {gün: {goal: sayı}} (yalnızca sayısı > 0 olanlar). Gün listede yoksa dönmez."""
    by_name = {field_name(g): g for g in GOALS}
    out = {}
    for item in (resp or {}).get("series", []) or []:
        day = str(item.get("time") or "")[:10]
        if len(day) != 10:
            continue
        counts = {}
        for k, v in item.items():
            g = by_name.get(k)
            if g is None:
                continue
            try:
                n = int(v or 0)
            except (TypeError, ValueError):
                n = 0
            if n > 0:
                counts[g] = n
        out[day] = counts
    return out

def fetch_goals(brand, days):
    """Bir marka için günleri ≤31'lik parçalarla çeker → {gün: {goal: sayı}}."""
    bot_id = BOTS[brand]
    out = {}
    for i in range(0, len(days), MAX_DAYS_PER_CALL):
        chunk = days[i:i + MAX_DAYS_PER_CALL]
        resp = call_report(build_payload(bot_id, chunk[0], chunk[-1]))
        parsed = parse_series(resp)
        for d in chunk:
            if d in parsed:
                out[d] = parsed[d]
        print(f"[API] {brand} {chunk[0]}..{chunk[-1]}: {len(parsed)} gün döndü, "
              f"{sum(sum(c.values()) for c in parsed.values())} goal")
        time.sleep(0.3)
    return out


# ── GOOGLE SHEETS ────────────────────────────────────────────
def sheets_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_info(
        json.loads(GCP_CREDS_JSON),
        scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return build("sheets", "v4", credentials=creds, cache_discovery=False)

def _exec(req, tries=5):
    last = None
    for i in range(tries):
        try:
            return req.execute()
        except Exception as e:
            last = e
            time.sleep(2 ** i)
    raise last

def ensure_tab(svc):
    meta = _exec(svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID))
    ids = {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}
    if TAB in ids:
        cur = _exec(svc.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"'{TAB}'!A1:D1")).get("values", [[]])
        if (cur[0] if cur else []) != HEADER:
            _exec(svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=f"'{TAB}'!A1",
                valueInputOption="RAW", body={"values": [HEADER]}))
        return ids[TAB]
    resp = _exec(svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": [
        {"addSheet": {"properties": {"title": TAB, "gridProperties": {"frozenRowCount": 1}}}}]}))
    sid = resp["replies"][0]["addSheet"]["properties"]["sheetId"]
    _exec(svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"'{TAB}'!A1",
        valueInputOption="RAW", body={"values": [HEADER]}))
    widths = [95, 70, 220, 80]
    reqs = [{"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                  "startColumnIndex": 0, "endColumnIndex": len(HEADER)},
        "cell": {"userEnteredFormat": {
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
            "backgroundColor": {"red": 0.17, "green": 0.24, "blue": 0.35}}},
        "fields": "userEnteredFormat(textFormat,backgroundColor)"}}]
    for i, w in enumerate(widths):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    _exec(svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": reqs}))
    print(f"[SCHEMA] '{TAB}' sekmesi oluşturuldu")
    return sid

def _norm_date(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime("%Y-%m-%d")
    return str(v)

def read_keys(svc):
    """A (tarih) ve B (marka) sütunları → [(tarih, marka)] satır sırasıyla (başlık dahil)."""
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{TAB}'!A:B",
        valueRenderOption="UNFORMATTED_VALUE"))
    out = []
    for r in resp.get("values", []):
        a = _norm_date(r[0]) if r else ""
        b = str(r[1]) if len(r) > 1 else ""
        out.append((a, b))
    return out

def _contiguous(idx):
    out = []
    for i in sorted(idx):
        if out and out[-1][1] == i:
            out[-1][1] = i + 1
        else:
            out.append([i, i + 1])
    return [tuple(x) for x in out]

def delete_rows(svc, sid, idx):
    ranges = _contiguous(idx)
    reqs = [{"deleteDimension": {"range": {
        "sheetId": sid, "dimension": "ROWS", "startIndex": a, "endIndex": b}}}
        for a, b in sorted(ranges, reverse=True)]
    for i in range(0, len(reqs), 100):
        _exec(svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body={"requests": reqs[i:i + 100]}))

def append_rows(svc, rows, chunk=4000):
    for i in range(0, len(rows), chunk):
        _exec(svc.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range=f"'{TAB}'!A1",
            valueInputOption="RAW", insertDataOption="OVERWRITE",
            body={"values": rows[i:i + chunk]}))

def write_day_brand(svc, sid, day_brand_rows):
    """day_brand_rows: {(gün, marka): [[Tarih, Marka, Goal, Sayı], ...]}
    Önce yeni satırlar eklenir, sonra aynı (gün, marka) için eski satırlar silinir."""
    keys = read_keys(svc)
    targets = set(day_brand_rows)
    old_idx = [i for i, k in enumerate(keys) if i > 0 and k in targets]
    new_rows = [r for k in sorted(day_brand_rows) for r in day_brand_rows[k]]
    if new_rows:
        append_rows(svc, new_rows)
    if old_idx:
        delete_rows(svc, sid, old_idx)
    return len(new_rows), len(old_idx)

def sort_tab(svc, sid):
    _exec(svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": [{
        "sortRange": {
            "range": {"sheetId": sid, "startRowIndex": 1, "startColumnIndex": 0, "endColumnIndex": len(HEADER)},
            "sortSpecs": [
                {"dimensionIndex": 0, "sortOrder": "ASCENDING"},
                {"dimensionIndex": 1, "sortOrder": "ASCENDING"},
                {"dimensionIndex": 3, "sortOrder": "DESCENDING"},
            ]}}]}))


# ── ELLE TUTULAN HAFTALIK SHEET ──────────────────────────────
def read_agent_tags(svc, week_start_str):
    """Tag Sheet'i → {marka: {tag(küçük harf): chat}} (Haftalık, Departman = Tümü). Okunamazsa None."""
    if not TAGS_SPREADSHEET_ID:
        return None
    label = tags_week_label(week_start_str, week_end(week_start_str))
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=TAGS_SPREADSHEET_ID, range="'Haftalık'!A2:E",
        valueRenderOption="UNFORMATTED_VALUE"))
    out = {}
    for r in resp.get("values", []):
        if len(r) < 5 or str(r[0]).strip() != label or str(r[2]).strip() != "Tümü":
            continue
        try:
            n = int(r[4])
        except (TypeError, ValueError):
            continue
        d = out.setdefault(str(r[1]).strip(), {})
        d[norm_tag(r[3])] = d.get(norm_tag(r[3]), 0) + n
    return out

def manual_find_or_create_tab(svc, ws, allow_create):
    """Haftanın sekmesini bulur (ad biçimi esnek); yoksa en yeni sekmenin kopyasından açar.
    (sheetId, başlık) döner; sekme yoksa ve açılamazsa None."""
    meta = _exec(svc.spreadsheets().get(spreadsheetId=MANUAL_SPREADSHEET_ID))
    sheets = meta["sheets"]
    we = week_end(ws)
    want = norm_tab(manual_tab_name(ws, we))
    for sh in sheets:
        if norm_tab(sh["properties"]["title"]) == want:
            return sh["properties"]["sheetId"], sh["properties"]["title"]
    if not allow_create or len(sheets) < 2:
        return None
    src = sheets[1]["properties"]        # Yearly'nin yanındaki = en yeni hafta
    title = manual_tab_name(ws, we)
    resp = _exec(svc.spreadsheets().batchUpdate(spreadsheetId=MANUAL_SPREADSHEET_ID, body={"requests": [
        {"duplicateSheet": {"sourceSheetId": src["sheetId"], "insertSheetIndex": 1, "newSheetName": title}}]}))
    sid = resp["replies"][0]["duplicateSheet"]["properties"]["sheetId"]
    # eski haftanın girdi sütunlarını temizle (formüllere dokunma) + 2. satır etiketleri
    clear = [f"'{title}'!{a}{MANUAL_FIRST_ROW}:{a}" for pair in MANUAL_BLOCKS.values() for a in pair]
    _exec(svc.spreadsheets().values().batchClear(
        spreadsheetId=MANUAL_SPREADSHEET_ID, body={"ranges": clear}))
    label = manual_short_label(ws, we)
    _exec(svc.spreadsheets().values().batchUpdate(spreadsheetId=MANUAL_SPREADSHEET_ID, body={
        "valueInputOption": "RAW",
        "data": [{"range": f"'{title}'!{c}2", "values": [[label]]} for c in MANUAL_LABEL_COLS]}))
    print(f"[MANUAL] yeni hafta sekmesi açıldı: '{title}' (kaynak: '{src['title']}')")
    return sid, title

def update_manual_week(svc, ws, bot_by_brand, agent_by_brand):
    """bot_by_brand: {marka: {goal: sayı}} (None/yoksa o markanın BOT sütunu elleme),
    agent_by_brand: {marka: {tag(küçük): chat}} ya da None.
    Sekmede satırı olmayan tag'ler (sayısı > 0) tabloya, son satırın biçim/formülleri kopyalanarak eklenir."""
    yesterday_ws = week_start(sofia_yesterday())
    prev_ws = (datetime.strptime(yesterday_ws, "%Y-%m-%d") - timedelta(days=7)).strftime("%Y-%m-%d")
    allow = ws >= yesterday_ws
    if ws < prev_ws:
        print(f"[MANUAL] {ws} haftası eski, elle girilmiş veri korunuyor → atlandı")
        return
    found = manual_find_or_create_tab(svc, ws, allow)
    if not found:
        print(f"[MANUAL] {ws} haftasının sekmesi yok → atlandı")
        return
    sid, title = found
    col = _exec(svc.spreadsheets().values().get(
        spreadsheetId=MANUAL_SPREADSHEET_ID, range=f"'{title}'!A{MANUAL_FIRST_ROW}:A")).get("values", [])
    row_of = {}
    for i, r in enumerate(col):
        k = norm_tag(r[0]) if r else ""
        if k and k not in row_of:           # aynı adlı ikinci satır (payment_method_request) boş kalır
            row_of[k] = i
    n_rows = len(col)

    # kaynaklar: (sütun, {tag: sayı}) — veri alınamayanlar atlanır
    sources = []
    for brand, (bot_col, agent_col) in MANUAL_BLOCKS.items():
        bot = (bot_by_brand or {}).get(brand)
        if bot is not None:
            norm = {}
            for g, n in bot.items():
                k = norm_tag(GOAL_TO_TAG.get(g, g))
                norm[k] = norm.get(k, 0) + n
            sources.append((bot_col, roll_promo(norm)))
        agent = (agent_by_brand or {}).get(brand)
        if agent is not None:
            sources.append((agent_col, roll_promo(agent)))

    # sekmede satırı olmayan ama sayısı olan tag'ler → tabloya ekle
    new_tags = sorted({k for _, src in sources for k, n in src.items() if n > 0 and k not in row_of})
    if new_tags and n_rows > 0:
        tmpl = MANUAL_FIRST_ROW - 1 + n_rows - 1          # son tag satırı (0 tabanlı)
        reqs, names = [], []
        for j, k in enumerate(new_tags):
            dest = MANUAL_FIRST_ROW - 1 + n_rows + j
            reqs.append({"copyPaste": {
                "source": {"sheetId": sid, "startRowIndex": tmpl, "endRowIndex": tmpl + 1,
                           "startColumnIndex": 0, "endColumnIndex": 20},
                "destination": {"sheetId": sid, "startRowIndex": dest, "endRowIndex": dest + 1,
                                "startColumnIndex": 0, "endColumnIndex": 20},
                "pasteType": "PASTE_NORMAL"}})
            for c in MANUAL_NAME_COLS:
                names.append({"range": f"'{title}'!{c}{dest + 1}", "values": [[k]]})
            row_of[k] = n_rows + j
        _exec(svc.spreadsheets().batchUpdate(spreadsheetId=MANUAL_SPREADSHEET_ID, body={"requests": reqs}))
        _exec(svc.spreadsheets().values().batchUpdate(spreadsheetId=MANUAL_SPREADSHEET_ID, body={
            "valueInputOption": "RAW", "data": names}))
        n_rows += len(new_tags)
        print(f"[MANUAL] '{title}': tabloya yeni tag satırı eklendi: {', '.join(new_tags)}")

    data = []
    for cname, src in sources:
        vals = [[""] for _ in range(n_rows)]
        for k, n in src.items():
            i = row_of.get(k)
            if i is not None:
                vals[i] = [n]
        data.append({"range": f"'{title}'!{cname}{MANUAL_FIRST_ROW}:{cname}{MANUAL_FIRST_ROW + n_rows - 1}",
                     "values": vals})
    if data:
        _exec(svc.spreadsheets().values().batchUpdate(spreadsheetId=MANUAL_SPREADSHEET_ID, body={
            "valueInputOption": "RAW", "data": data}))
    print(f"[MANUAL] '{title}': {len(data)} sütun yazıldı")
    repair_pct_formulas(svc, only_title=title)


# ── % FORMÜLÜ ONARIMI ────────────────────────────────────────
PCT_BLOCKS = [("B", "C", "D"), ("G", "H", "I"), ("L", "M", "N"), ("Q", "R", "S")]   # (BOT, AGENT, %)

def repair_pct_formulas(svc, only_title=None):
    """Her haftalık sekmede % sütunlarının formülünü =(BOT-AGENT)/BOT kalıbına zorlar
    (ör. D81 yanlışlıkla =(B81-C80)/B81 / #REF! olmuş). Yalnızca formülü kalıba uymayan hücrelere dokunur."""
    meta = _exec(svc.spreadsheets().get(spreadsheetId=MANUAL_SPREADSHEET_ID))
    titles = [sh["properties"]["title"] for sh in meta["sheets"][1:]]      # ilk sekme Yearly
    if only_title:
        titles = [t for t in titles if t == only_title]
    fixed_total, data = 0, []
    for i in range(0, len(titles), 20):
        chunk = titles[i:i + 20]
        resp = _exec(svc.spreadsheets().values().batchGet(
            spreadsheetId=MANUAL_SPREADSHEET_ID, valueRenderOption="FORMULA",
            ranges=[f"'{t}'!A{MANUAL_FIRST_ROW}:S" for t in chunk]))
        for t, vr in zip(chunk, resp.get("valueRanges", [])):
            fixed = []
            for ri, row in enumerate(vr.get("values", [])):
                r = MANUAL_FIRST_ROW + ri
                for b, a, pc in PCT_BLOCKS:
                    ci = ord(pc) - ord("A")
                    v = row[ci] if ci < len(row) else ""
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue                      # boş / elle girilmiş hücreye dokunma
                    exp = f"=({b}{r}-{a}{r})/{b}{r}"
                    if v.replace(" ", "") != exp:
                        data.append({"range": f"'{t}'!{pc}{r}", "values": [[exp]]})
                        fixed.append(f"{pc}{r}")
            if fixed:
                fixed_total += len(fixed)
                print(f"[REPAIR] '{t}': {', '.join(fixed)}")
    if data:
        for i in range(0, len(data), 200):
            _exec(svc.spreadsheets().values().batchUpdate(spreadsheetId=MANUAL_SPREADSHEET_ID, body={
                "valueInputOption": "USER_ENTERED", "data": data[i:i + 200]}))
    print(f"[REPAIR] {len(titles)} sekme tarandı, {fixed_total} hücre düzeltildi")

# ── HATA MAİLİ ───────────────────────────────────────────────
def send_error_email(label, error_msg, tb_str):
    if not (GMAIL_USER and GMAIL_PASS and REPORT_EMAILS):
        return
    html = f"""<!DOCTYPE html><html><body style="font-family:monospace;padding:20px;">
<h2 style="color:#c0392b;">Bot Goals — Hata Raporu</h2>
<p><strong>Tarih:</strong> {label}</p>
<p><strong>Hata:</strong> {error_msg}</p>
<pre style="background:#f8f9fa;padding:16px;border-radius:6px;font-size:12px;
     white-space:pre-wrap;word-break:break-all;">{tb_str}</pre>
<p style="color:#666;font-size:11px;">CS Operations &amp; Technology</p>
</body></html>"""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"⚠️ Bot Goals ERROR | {label}"
    msg["From"] = GMAIL_USER
    msg["To"] = REPORT_EMAILS[0]
    if len(REPORT_EMAILS) > 1:
        msg["Cc"] = ", ".join(REPORT_EMAILS[1:])
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.sendmail(GMAIL_USER, REPORT_EMAILS, msg.as_string())
        print("[EMAIL] Error mail sent")
    except Exception as e:
        print(f"[EMAIL] Could not send: {e}")


# ── ANA AKIŞ ─────────────────────────────────────────────────
def probe(day):
    """Sheet'e dokunmadan API'yi dener; ham cevabın başını ve parse sonucunu yazdırır."""
    for brand, bot_id in BOTS.items():
        print(f"\n--- {brand} ({bot_id}) {day}")
        resp = call_report(build_payload(bot_id, day, day))
        print(json.dumps(resp, ensure_ascii=False)[:600])
        parsed = parse_series(resp)
        for d, c in parsed.items():
            print(f"{d}: {sum(c.values())} goal, {len(c)} çeşit; ilk 5: {sorted(c.items(), key=lambda x: -x[1])[:5]}")

def sync_manual(svc, days):
    """İlgili haftaları (yalnızca bu hafta / geçen hafta) elle tutulan Sheet'e yazar."""
    yesterday = sofia_yesterday()
    prev_ws = (datetime.strptime(week_start(yesterday), "%Y-%m-%d") - timedelta(days=7)).strftime("%Y-%m-%d")
    for ws in sorted({week_start(d) for d in days}):
        if ws < prev_ws:                       # eski hafta: elle girilmiş veri korunur, API'ye gitme
            print(f"[MANUAL] {ws} haftası eski, elle girilmiş veri korunuyor → atlandı")
            continue
        wk_days = [d for d in date_range(ws, week_end(ws)) if d <= yesterday]
        if not wk_days:
            continue
        bot_by_brand = {}
        for brand in BOTS:
            try:
                got = fetch_goals(brand, wk_days)
            except Exception as e:
                print(f"[MANUAL] {brand} {ws}: goals alınamadı ({e}) → BOT sütunu olduğu gibi kaldı")
                continue
            if not got:                        # API o hafta için gün döndürmedi → sütunu elleme
                print(f"[MANUAL] {brand} {ws}: goals verisi yok → BOT sütunu olduğu gibi kaldı")
                continue
            tot = {}
            for counts in got.values():
                for g, n in counts.items():
                    tot[g] = tot.get(g, 0) + n
            bot_by_brand[brand] = tot
        try:
            agent_by_brand = read_agent_tags(svc, ws)
        except Exception as e:
            print(f"[MANUAL] AGENT'S TAGS okunamadı ({e}) → o sütunlar olduğu gibi kaldı")
            agent_by_brand = None
        update_manual_week(svc, ws, bot_by_brand, agent_by_brand)

def main():
    t0 = time.time()
    args = sys.argv[1:]
    if args and args[0] == "repair":
        if not MANUAL_SPREADSHEET_ID:
            raise SystemExit("MANUAL_SPREADSHEET_ID tanımlı değil.")
        repair_pct_formulas(sheets_service())
        return
    if args and args[0] == "probe":
        probe(args[1] if len(args) > 1 else sofia_yesterday())
        return
    if len(args) >= 2:
        days = date_range(args[0], args[1])
    elif len(args) == 1:
        days = [args[0]]
    else:
        days = [sofia_yesterday()]
    label = days[0] if len(days) == 1 else f"{days[0]}..{days[-1]}"
    print(f"\n{'='*55}\nBOT GOALS — {label}\n{'='*55}\n")
    try:
        if not SPREADSHEET_ID:
            raise RuntimeError("GOALS_SPREADSHEET_ID tanımlı değil (GitHub secret eksik).")
        day_brand_rows, empty = {}, []
        for brand in BOTS:
            got = fetch_goals(brand, days)
            for d in days:
                counts = got.get(d)
                if not counts:                  # gün dönmedi ya da hepsi 0 → yazma, eski veri varsa koru
                    empty.append((d, brand))
                    continue
                day_brand_rows[(d, brand)] = [
                    [d, brand, g, n] for g, n in sorted(counts.items(), key=lambda x: (-x[1], x[0]))]
        for d, b in empty:
            print(f"[UYARI] {d} {b}: veri yok / tüm goal'ler 0 → yazılmadı")
        svc = sheets_service()
        sid = ensure_tab(svc)
        added, removed = write_day_brand(svc, sid, day_brand_rows)
        sort_tab(svc, sid)
        print(f"\n[DONE] {len(day_brand_rows)} gün×marka, +{added} satır, -{removed} eski satır, "
              f"{time.time()-t0:.0f} sn")
        if MANUAL_SPREADSHEET_ID:
            try:
                sync_manual(svc, days)
            except Exception as e:             # Goals Sheet yazıldı; elle Sheet hatası işi düşürmesin
                tb = traceback.format_exc()
                print(f"[MANUAL] HATA: {e}\n{tb}")
                send_error_email(label + " (yalnızca elle Sheet adımı)", str(e), tb)
        if empty and len(empty) == len(days) * len(BOTS):
            raise RuntimeError(f"{label}: hiçbir marka için veri dönmedi.")
    except Exception as e:
        tb = traceback.format_exc()
        print(tb)
        send_error_email(label, str(e), tb)
        sys.exit(1)

if __name__ == "__main__":
    main()
