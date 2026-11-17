#!/usr/bin/env python3
# ============================================================
# TAG SHEET — v1 (GitHub Actions)
# Comm100 wrap-up tag sayımlarını (marka x tag x hafta) Reporting API'den
# çekip "ALL CS Agents Tags" sheet'ine yazar.
#
# Modlar:
#   calibrate YYYY-MM   Ayı farklı filtre varyantlarıyla çek, sheet'teki mevcut
#                       sekmelerle karşılaştır, sonucu "_calibration" sekmesine yaz.
#                       (Sheet'teki hiçbir ay sekmesine DOKUNMAZ)
#   fill YYYY-MM [...]  Belirtilen ay(lar)ın 4 sekmesini (All Brands, Betsat,
#                       Superbetin, Turkbet) yeniden üretir. Sekme yoksa en yeni
#                       ay sekmesini şablon alıp kopyalar.
#   summary YYYY-MM[,..] Biten ayları yıllık özet sekmesine ekler (var olan bloğa dokunmaz;
#                       SUMMARY_OVERWRITE=1 ile üzerine yazar). fill zaten bunu otomatik yapar.
#   top10 YYYY-MM[,..]  Mevcut sekmelerin Top 10 bloğunu doğru satıra (ALL: K5, markalar: K7)
#                       yeniden yazar. API çağrısı yapmaz; kayma onarımı için.
#   daily               Dünün ait olduğu ayı (ayın ilk 3 günü ise önceki ayı da)
#                       yeniden üretir. İdempotent.
#
# Sayım Reporting API ile yapılır (rowGroups: campaign x wrap-up option), bu yüzden
# chat chat taramaya gerek yok; ay başına birkaç istek yeterli.
# ============================================================
import os
import re
import sys
import json
import time
import base64
import calendar
from collections import defaultdict
from datetime import date, datetime, timedelta

import requests
import pytz
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import HttpRequest

# Sheets API dakika başı yazma kotası (60) aşılırsa 429 döner; bekleyip tekrar dene.
_orig_execute = HttpRequest.execute
def _execute_with_retry(self, *args, **kwargs):
    for attempt in range(8):
        try:
            return _orig_execute(self, *args, **kwargs)
        except HttpError as e:
            status = getattr(e.resp, "status", 0)
            if status in (429, 500, 502, 503) and attempt < 7:
                wait = 25 if status == 429 else 5
                print(f"[SHEETS] HTTP {status}, {wait}s bekleyip tekrar denenecek ({attempt+1}/7)")
                time.sleep(wait)
                continue
            raise
HttpRequest.execute = _execute_with_retry

# ── CONFIG ───────────────────────────────────────────────────
SITE_ID        = os.environ["COMM100_SITE_ID"]
API_KEY        = os.environ["COMM100_API_KEY"]
COMM100_EMAIL  = os.environ["COMM100_EMAIL"]
SPREADSHEET_ID = os.environ["TAGS_SPREADSHEET_ID"]
GCP_CREDS_JSON = os.environ["GCP_CREDENTIALS"]

CS_DEPT_ID = "29098b6e-0acb-40cc-8cd7-ebf662defc42"

# Sheet'teki sekme sırası: All Brands, Betsat, Superbetin, Turkbet
CAMPAIGNS = {
    "Superbetin": "3cd6532a-0d60-4d7d-9331-5dd5be71f527",
    "Betsat":     "2b71ad44-6a12-4073-bc6c-8edb22eabcc6",
    "Turkbet":    "6712311a-2268-408c-a803-b338a1308010",
}
CAMPAIGN_TO_BRAND = {v.lower(): k for k, v in CAMPAIGNS.items()}
BRAND_ORDER = ["Betsat", "Superbetin", "Turkbet"]
ALL_PREFIX  = "TAGs All Brands"
TAB_KINDS   = [ALL_PREFIX] + BRAND_ORDER

TR_MONTHS = ["Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
             "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"]
EN_MON    = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
             "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# Kalibrasyon sonucuna göre güncellenecek varsayılan varyant
DEFAULT_VARIANT = os.environ.get("TAG_VARIANT", "A_dur")

# Filtre varyantları (hepsinde RequestedTime filtresi var)
VARIANTS = {
    "A_dur":          {"duration": True,  "dept": False, "agent": False},
    "B_dur_cs":       {"duration": True,  "dept": True,  "agent": False},
    "C_dur_agent":    {"duration": True,  "dept": False, "agent": True},
    "D_dur_agent_cs": {"duration": True,  "dept": True,  "agent": True},
    "E_nofilter":     {"duration": False, "dept": False, "agent": False},
}

N_WEEK_SLOTS = 6   # B..G sütunları
SHEET_ROWS   = 300
CAL_SHEET    = "_calibration"

unknown_campaigns = defaultdict(int)
raw_samples = []

# ── DATE ─────────────────────────────────────────────────────
def sofia_today():
    return datetime.now(pytz.timezone("Europe/Sofia")).date()

def month_weeks(y, m):
    """Pazartesi-Pazar haftaları, ay sınırlarına kırpılmış."""
    last = date(y, m, calendar.monthrange(y, m)[1])
    d = date(y, m, 1)
    weeks = []
    while d <= last:
        end = min(d + timedelta(days=6 - d.weekday()), last)
        weeks.append((d, end))
        d = end + timedelta(days=1)
    return weeks

def week_label(a, b):
    mon = EN_MON[a.month - 1]
    return f"{a.day} {mon}" if a == b else f"{a.day}-{b.day} {mon}"

def api_date(d):
    return d.strftime("%Y/%m/%d")

def parse_ym(s):
    m = re.match(r"^(\d{4})-(\d{1,2})$", s.strip())
    if not m:
        raise SystemExit(f"Geçersiz ay: {s} (format YYYY-MM)")
    return int(m.group(1)), int(m.group(2))

# ── COMM100 REPORTING API ────────────────────────────────────
def comm100_auth():
    return "Basic " + base64.b64encode(f"{COMM100_EMAIL}:{API_KEY}".encode()).decode()

def post_report(payload):
    url = f"https://dash15.lively-chat.com/api/reportingquery/reports/query?siteId={SITE_ID}"
    last = ""
    for attempt in range(4):
        try:
            r = requests.post(
                url,
                headers={"Authorization": comm100_auth(), "Content-Type": "application/json"},
                json=payload, timeout=120,
            )
            if r.status_code == 200:
                return r.json()
            last = f"HTTP {r.status_code}: {r.text[:300]}"
            if r.status_code not in (429, 500, 502, 503, 504):
                break
        except requests.RequestException as e:
            last = str(e)
        time.sleep(2 * (2 ** attempt))
    raise RuntimeError(f"Reporting API hatası: {last}")

def tag_payload(d_from, d_to, variant):
    v = VARIANTS[variant]
    filters = [{"fieldName": "RequestedTime", "matchType": "between",
                "value": [api_date(d_from), api_date(d_to)]}]
    if v["duration"]:
        filters.append({"fieldName": "Duration", "matchType": "notEquals", "value": ["0"]})
    if v["dept"]:
        filters.append({"fieldName": "DepartmentId", "matchType": "equals", "value": [CS_DEPT_ID]})
    if v["agent"]:
        filters.append({"fieldName": "ChatType", "matchType": "equals", "value": ["0", "2"]})
        filters.append({"fieldName": "Status", "matchType": "equals", "value": ["0"]})

    def of(name, field):
        return {"name": name, "calculationType": "originalValue", "valueType": "string",
                "fieldName": field, "conditionExpression": "", "conditionMatchType": "all",
                "conditions": [], "enumItems": [], "expression": ""}

    return {
        "cubeEntities": [{
            "name": "Chat",
            "fields": [
                of("campaignId", "Campaign.Id"),
                of("categoryOptionId", "ChatWrapupCategory.CategoryOption.Id"),
                of("categoryOptionName", "ChatWrapupCategory.CategoryOption.Name"),
                {"name": "count", "calculationType": "count", "valueType": "int", "fieldName": "Id",
                 "conditionExpression": "", "conditionMatchType": "all", "conditions": [],
                 "enumItems": [], "expression": ""},
            ],
            "filters": filters,
            "rowGroups": [
                {"name": "campaignId", "fieldName": "CampaignId", "isFull": False},
                {"name": "optionId", "fieldName": "ChatWrapupCategory.CategoryOptionId", "isFull": False},
            ],
        }],
        "mergeType": "column",
        "timezone": "FLE Standard Time",
    }

def fetch_counts(d_from, d_to, variant):
    """-> {brand: {tag: count}} (sadece 3 marka; bilinmeyen campaign'ler loglanır)."""
    series = post_report(tag_payload(d_from, d_to, variant)).get("series") or []
    if len(raw_samples) < 3:
        raw_samples.extend(series[: 3 - len(raw_samples)])
    out = {b: defaultdict(int) for b in CAMPAIGNS}
    for row in series:
        tag = str(row.get("categoryOptionName") or "").strip()
        if not tag or not row.get("categoryOptionId"):
            continue
        cnt = int(row.get("count") or 0)
        brand = CAMPAIGN_TO_BRAND.get(str(row.get("campaignId") or "").lower())
        if not brand:
            unknown_campaigns[str(row.get("campaignId"))] += cnt
            continue
        out[brand][tag] += cnt
    time.sleep(0.3)
    return out

# ── GOOGLE SHEETS ────────────────────────────────────────────
_svc = None
def svc():
    global _svc
    if _svc is None:
        creds = service_account.Credentials.from_service_account_info(
            json.loads(GCP_CREDS_JSON),
            scopes=["https://www.googleapis.com/auth/spreadsheets"],
        )
        _svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    return _svc

def list_tabs():
    meta = svc().spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return [(s["properties"]["title"], s["properties"]["sheetId"], s["properties"]["index"])
            for s in meta.get("sheets", [])]

TAB_RE = re.compile(r"^(%s)\s+(%s)\s+(\d{4})\s*$" % (
    "|".join(re.escape(k) for k in TAB_KINDS), "|".join(TR_MONTHS)))

def parse_tab_title(title):
    m = TAB_RE.match(title)
    if not m:
        return None
    return m.group(1), int(m.group(3)), TR_MONTHS.index(m.group(2)) + 1

def find_tab(tabs, kind, y, m):
    for title, sid, idx in tabs:
        p = parse_tab_title(title)
        if p and p == (kind, y, m):
            return title, sid
    return None

def template_tab_of_kind(tabs, kind, y, m):
    """Hedef aydan ÖNCEKİ en yakın aynı-tür sekme; yoksa sonraki en yakın. -> (key, title, sid)"""
    target = (y, m)
    before, after = None, None
    for title, sid, idx in tabs:
        p = parse_tab_title(title)
        if not p or p[0] != kind:
            continue
        key = (p[1], p[2])
        if key < target and (before is None or key > before[0]):
            before = (key, title, sid)
        elif key > target and (after is None or key < after[0]):
            after = (key, title, sid)
    return before or after

def read_tab(title):
    """-> {tag: [w1..w6]} (aynı isimli satırlar toplanır)."""
    res = svc().spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1:I{SHEET_ROWS}",
        valueRenderOption="UNFORMATTED_VALUE").execute()
    rows = res.get("values", [])
    header = rows[0] if rows else []
    data = {}
    for r in rows[1:]:
        if not r or not str(r[0]).strip():
            continue
        tag = str(r[0]).strip()
        vals = []
        for i in range(1, 1 + N_WEEK_SLOTS):
            v = r[i] if i < len(r) else 0
            vals.append(int(v) if isinstance(v, (int, float)) else 0)
        if tag in data:
            data[tag] = [a + b for a, b in zip(data[tag], vals)]
        else:
            data[tag] = vals
    return header, data

# ── WEEKLY TABLES ────────────────────────────────────────────
def collect_month(y, m, variant, upto=None):
    """-> weeks, {brand: {tag: [w1..w6]}}  (upto: bu tarihten sonrası sorgulanmaz)"""
    weeks = month_weeks(y, m)
    table = {b: defaultdict(lambda: [0] * N_WEEK_SLOTS) for b in BRAND_ORDER}
    for wi, (a, b) in enumerate(weeks):
        if upto and a > upto:
            break
        end = min(b, upto) if upto else b
        counts = fetch_counts(a, end, variant)
        for brand in BRAND_ORDER:
            for tag, c in counts[brand].items():
                table[brand][tag][wi] += c
    return weeks, table

def sum_brands(table):
    allb = defaultdict(lambda: [0] * N_WEEK_SLOTS)
    for brand in BRAND_ORDER:
        for tag, ws in table[brand].items():
            for i, v in enumerate(ws):
                allb[tag][i] += v
    return allb

# ── CALIBRATE ────────────────────────────────────────────────
def calibrate(y, m):
    tabs = list_tabs()
    weeks = month_weeks(y, m)
    nw = len(weeks)
    sheet = {}
    for kind in TAB_KINDS:
        t = find_tab(tabs, kind, y, m)
        if not t:
            raise SystemExit(f"Sekme bulunamadı: {kind} {TR_MONTHS[m-1]} {y}")
        header, data = read_tab(t[0])
        exp = [week_label(a, b) for a, b in weeks]
        got = [str(h).strip() for h in header[1:1 + nw]]
        if got != exp:
            print(f"[UYARI] {t[0]} hafta başlıkları farklı: sheet={got} beklenen={exp}")
        sheet[kind] = data

    summary, best_diffs = [], {}
    for variant in VARIANTS:
        print(f"[CAL] varyant {variant} çekiliyor ({nw} hafta)...")
        _, table = collect_month(y, m, variant)
        api = {b: {t: list(w) for t, w in table[b].items()} for b in BRAND_ORDER}
        api[ALL_PREFIX] = {t: list(w) for t, w in sum_brands(table).items()}
        tot_abs = tot_sheet = tot_api = exact = cells = 0
        diffs = []
        for kind in TAB_KINDS:
            tags = set(sheet[kind]) | set(api[kind])
            for tag in tags:
                sw = sheet[kind].get(tag, [0] * N_WEEK_SLOTS)
                aw = api[kind].get(tag, [0] * N_WEEK_SLOTS)
                for wi in range(nw):
                    cells += 1
                    d = aw[wi] - sw[wi]
                    tot_abs += abs(d)
                    tot_sheet += sw[wi]
                    tot_api += aw[wi]
                    if d == 0:
                        exact += 1
                    else:
                        diffs.append((abs(d), kind, week_label(*weeks[wi]), tag, sw[wi], aw[wi], d))
        summary.append((variant, tot_abs, tot_sheet, tot_api, exact, cells, len(diffs)))
        best_diffs[variant] = sorted(diffs, reverse=True)
        print(f"[CAL] {variant}: |fark| toplamı={tot_abs} sheet={tot_sheet} api={tot_api} "
              f"birebir={exact}/{cells}")

    summary.sort(key=lambda s: s[1])
    best = summary[0][0]
    print(f"[CAL] EN İYİ VARYANT: {best}")

    rows = [[f"Kalibrasyon {TR_MONTHS[m-1]} {y}", "", "", "", "", "", ""],
            ["Varyant", "|Fark| toplamı", "Sheet toplamı", "API toplamı", "Birebir hücre", "Toplam hücre", "Farklı hücre"]]
    rows += [list(s) for s in summary]
    rows += [[], [f"En iyi varyant: {best} — en büyük farklar"],
             ["Sekme", "Hafta", "Tag", "Sheet", "API", "Fark"]]
    rows += [[d[1], d[2], d[3], d[4], d[5], d[6]] for d in best_diffs[best][:80]]
    rows += [[], ["Bilinmeyen campaign'ler (3 markaya eşleşmeyen)", "Toplam sayı"]]
    rows += [[k, v] for k, v in unknown_campaigns.items()] or [["yok", 0]]
    rows += [[], ["Ham API örneği (ilk 3 satır)"]]
    rows += [[json.dumps(s, ensure_ascii=False)[:300]] for s in raw_samples]
    write_calibration(rows)

def write_calibration(rows):
    tabs = list_tabs()
    if not any(t[0] == CAL_SHEET for t in tabs):
        svc().spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": CAL_SHEET}}}]}).execute()
    svc().spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=f"'{CAL_SHEET}'", body={}).execute()
    svc().spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"'{CAL_SHEET}'!A1",
        valueInputOption="RAW", body={"values": rows}).execute()
    print(f"[CAL] '{CAL_SHEET}' sekmesine yazıldı ({len(rows)} satır)")

# ── FILL ─────────────────────────────────────────────────────
def sort_key(tag):
    return (tag.lower(), tag)

def ensure_tab(kind, y, m, tabs):
    """Hedef sekmeyi döndür; yoksa en yeni aynı-tür sekmeyi kopyalayıp oluştur."""
    found = find_tab(tabs, kind, y, m)
    if found:
        return found[0], found[1], False
    tpl = template_tab_of_kind(tabs, kind, y, m)
    if not tpl:
        raise SystemExit(f"Şablon sekme yok: {kind}")
    title = f"{kind} {TR_MONTHS[m-1]} {y}"
    resp = svc().spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"duplicateSheet": {
            "sourceSheetId": tpl[2], "newSheetName": title, "insertSheetIndex": 1}}]}).execute()
    sid = resp["replies"][0]["duplicateSheet"]["properties"]["sheetId"]
    print(f"[FILL] Yeni sekme: {title} (şablon: {tpl[1]})")
    return title, sid, True

def top10_block(y, m, table_by_tag):
    totals = {t: sum(ws) for t, ws in table_by_tag.items() if not t.upper().startswith("VIP")}
    base = sum(totals.values()) or 1
    top = sorted(totals.items(), key=lambda kv: (-kv[1], kv[0]))[:10]
    rows = [[f"{EN_MON[m-1]} {y} / Top 10 Tags", "Totals", "%"]]
    rows += [[t, c, c / base] for t, c in top]
    return rows

TOP10_ROW = {ALL_PREFIX: 5, "Betsat": 7, "Superbetin": 7, "Turkbet": 7}

def write_top10(title, sid, kind, y, m, data):
    """Top 10 bloğunu K:M'e yazar. Şablondaki biçim ALL sekmesinde 5. satırda,
    marka sekmelerinde 7. satırda başlar (logo üstte durur) — o yüzden tür bazlı."""
    row = TOP10_ROW[kind]
    svc().spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!K1:M{SHEET_ROWS}", body={}).execute()
    block = top10_block(y, m, data)
    svc().spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!K{row}", valueInputOption="USER_ENTERED",
        body={"values": block}).execute()
    svc().spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": row, "endRowIndex": row + len(block) - 1,
                      "startColumnIndex": 12, "endColumnIndex": 13},
            "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0.00%"}}},
            "fields": "userEnteredFormat.numberFormat"}}]}).execute()

def top10_repair(y, m):
    """Mevcut sekmelerin kendi verisinden Top 10 bloğunu doğru satıra yeniden yazar (API çağrısı yok)."""
    tabs = list_tabs()
    for kind in TAB_KINDS:
        t = find_tab(tabs, kind, y, m)
        if not t:
            print(f"[TOP10] sekme yok: {kind} {TR_MONTHS[m-1]} {y}")
            continue
        _, data = read_tab(t[0])
        write_top10(t[0], t[1], kind, y, m, data)
        print(f"[TOP10] {t[0]}: K{TOP10_ROW[kind]} satırına yazıldı")
        time.sleep(4)

def fill_month(y, m, variant, today=None):
    today = today or sofia_today()
    yesterday = today - timedelta(days=1)
    first = date(y, m, 1)
    last = date(y, m, calendar.monthrange(y, m)[1])
    if first > yesterday:
        print(f"[FILL] {y}-{m:02d} henüz başlamadı, atlanıyor")
        return
    upto = yesterday if yesterday < last else None
    days = (upto or last).day
    print(f"[FILL] {TR_MONTHS[m-1]} {y} | varyant {variant} | gün sayısı {days}"
          f"{' (ay devam ediyor)' if upto else ''}")

    weeks, table = collect_month(y, m, variant, upto)
    tabs = list_tabs()

    # Şablon sekmelerdeki tag listesi satır sürekliliği için birleşime katılır
    per_kind = {ALL_PREFIX: sum_brands(table)}
    per_kind.update({b: table[b] for b in BRAND_ORDER})

    # Yeni sekmeler ters sırada eklenir ki sıra All, Betsat, Superbetin, Turkbet kalsın
    order = list(reversed(TAB_KINDS))
    infos = {}
    for kind in order:
        infos[kind] = ensure_tab(kind, y, m, list_tabs())

    for kind in TAB_KINDS:
        title, sid, is_new = infos[kind]
        tpl = template_tab_of_kind([t for t in list_tabs() if t[0] != title], kind, y, m)
        tpl_tags = set()
        if tpl:
            try:
                _, tpl_data = read_tab(tpl[1])
                tpl_tags = set(tpl_data)
            except Exception as e:
                print(f"[FILL] şablon tag listesi okunamadı ({tpl[1]}): {e}")
        data = per_kind[kind]
        tags = sorted(set(data) | tpl_tags, key=sort_key)

        header = [""] + [week_label(a, b) for a, b in weeks]
        header += [""] * (1 + N_WEEK_SLOTS - len(header)) + ["Totals", "Daily Avg."]
        values = [header]
        for i, tag in enumerate(tags):
            r = i + 2
            ws = list(data[tag]) if tag in data else [0] * N_WEEK_SLOTS
            values.append([tag] + ws + [f"=SUM(B{r}:G{r})", f"=H{r}/{days}"])

        svc().spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1:I{SHEET_ROWS}", body={}).execute()
        svc().spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1", valueInputOption="USER_ENTERED",
            body={"values": values}).execute()

        write_top10(title, sid, kind, y, m, data)
        total = sum(sum(ws) for ws in data.values())
        print(f"[FILL] {title}: {len(tags)} tag satırı, toplam {total}")
        time.sleep(4)

    if unknown_campaigns:
        print(f"[FILL] Eşleşmeyen campaign'ler: {dict(unknown_campaigns)}")

    # Ay bittiyse yıllık özet sekmesine de ekle (yoksa; varsa dokunmaz)
    try:
        update_summary(y, m)
    except Exception as e:
        print(f"[ÖZET] güncellenemedi (aylık sekmeler yine de yazıldı): {e}")

# ── YILLIK ÖZET SEKMESİ ("TAGs All Brands <yıl>") ────────────
# Düzen: her ay 3 sütun (ay adı + tag listesi | Totals | boşluk), Ocak A:B'den başlar.
EN_FULL = ["January", "February", "March", "April", "May", "June",
           "July", "August", "September", "October", "November", "December"]

def col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s

def grid_cols(sid):
    meta = svc().spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID, fields="sheets.properties").execute()
    for s in meta.get("sheets", []):
        if s["properties"]["sheetId"] == sid:
            return s["properties"]["gridProperties"].get("columnCount", 26)
    return 26

def update_summary(y, m, overwrite=False):
    """Biten ay için özet sekmesine blok ekler. Mevcut ay bloğuna (elle girilmiş eski aylar dahil)
    overwrite=False iken dokunmaz. Değerler hedef ay sekmesine formülle bağlanır."""
    last = date(y, m, calendar.monthrange(y, m)[1])
    if sofia_today() <= last:
        print(f"[ÖZET] {EN_FULL[m-1]} {y} henüz bitmedi, özete eklenmiyor")
        return
    tabs = list_tabs()
    s_title = f"{ALL_PREFIX} {y}"
    summ = [t for t in tabs if t[0].strip() == s_title]
    src = find_tab(tabs, ALL_PREFIX, y, m)
    if not summ or not src:
        print(f"[ÖZET] atlandı: özet sekmesi ({s_title}) veya ay sekmesi yok")
        return
    s_name, s_id = summ[0][0], summ[0][1]
    hdr = svc().spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{s_name}'!1:1").execute().get("values", [[]])
    row = hdr[0] if hdr else []
    # Düzen sabit: Ocak A, Şubat D, ... → ay sütunu = 3 * (ay-1) (sıradan bağımsız)
    pos = 3 * (m - 1)
    exists = pos < len(row) and str(row[pos]).strip() == EN_FULL[m - 1]
    if exists and not overwrite:
        print(f"[ÖZET] {EN_FULL[m-1]} bloğu zaten var, dokunulmadı")
        return
    new_block = not exists
    # biçim kaynağı: hedefin solundaki en yakın dolu ay bloğu
    last_month_col = None
    for i in range(pos - 3, -1, -3):
        if i < len(row) and str(row[i]).strip() in EN_FULL:
            last_month_col = i
            break

    colA = svc().spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{src[0]}'!A1:A{SHEET_ROWS}").execute().get("values", [])
    n = len([r for r in colA[1:] if r and str(r[0]).strip()])
    values = [[EN_FULL[m - 1], "Totals"]]
    for r in range(2, n + 2):
        values.append([f"='{src[0]}'!A{r}", f"='{src[0]}'!H{r}"])

    need = pos + 3
    reqs = []
    if grid_cols(s_id) < need:
        reqs.append({"appendDimension": {"sheetId": s_id, "dimension": "COLUMNS",
                                         "length": need - grid_cols(s_id)}})
    if new_block and last_month_col is not None:
        reqs.append({"copyPaste": {
            "source": {"sheetId": s_id, "startRowIndex": 0, "endRowIndex": n + 1,
                       "startColumnIndex": last_month_col, "endColumnIndex": last_month_col + 2},
            "destination": {"sheetId": s_id, "startRowIndex": 0, "endRowIndex": n + 1,
                            "startColumnIndex": pos, "endColumnIndex": pos + 2},
            "pasteType": "PASTE_FORMAT"}})
    # sütunlar genişletilmeden değer yazılamaz → önce yapı, sonra değer
    pre = [r for r in reqs if "appendDimension" in r]
    post = [r for r in reqs if "copyPaste" in r]
    if pre:
        svc().spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": pre}).execute()
    svc().spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"'{s_name}'!{col_letter(pos)}1",
        valueInputOption="USER_ENTERED", body={"values": values}).execute()
    if post:
        svc().spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": post}).execute()
    print(f"[ÖZET] {EN_FULL[m-1]} {y} bloğu yazıldı: {col_letter(pos)}:{col_letter(pos+1)} ({n} tag)")

# ── MAIN ─────────────────────────────────────────────────────
def main():
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    mode = sys.argv[1]
    if mode == "calibrate":
        y, m = parse_ym(sys.argv[2])
        calibrate(y, m)
    elif mode == "fill":
        variant = DEFAULT_VARIANT
        if variant not in VARIANTS:
            raise SystemExit(f"Bilinmeyen varyant: {variant}")
        months = [parse_ym(s) for s in sys.argv[2].replace(" ", "").split(",") if s]
        for y, m in months:
            fill_month(y, m, variant)
    elif mode == "summary":
        overwrite = os.environ.get("SUMMARY_OVERWRITE") == "1"
        for y, m in [parse_ym(s) for s in sys.argv[2].replace(" ", "").split(",") if s]:
            update_summary(y, m, overwrite=overwrite)
    elif mode == "top10":
        for y, m in [parse_ym(s) for s in sys.argv[2].replace(" ", "").split(",") if s]:
            top10_repair(y, m)
    elif mode == "daily":
        yday = sofia_today() - timedelta(days=1)
        months = [(yday.year, yday.month)]
        if yday.day <= 3:
            prev = date(yday.year, yday.month, 1) - timedelta(days=1)
            months.insert(0, (prev.year, prev.month))
        for y, m in months:
            fill_month(y, m, DEFAULT_VARIANT)
    else:
        raise SystemExit(f"Bilinmeyen mod: {mode}")

if __name__ == "__main__":
    main()
