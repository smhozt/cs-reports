#!/usr/bin/env python3
# ============================================================
# MEMBER TAGS — v1 (GitHub Actions)
# Günlük: dünün chatleri → marka x departman x tag bazında
#         chat sayısı + TEKİL ÜYE sayısı (Google Sheets)
# Haftalık özet ve yoğun üye listesi (Risk Level sütunuyla)
# her çalışmada o haftanın başından bugüne yeniden hesaplanır.
#
# Kullanım:
#   python member_tags.py                      → dün (Sofia)
#   python member_tags.py 2026-11-09           → tek gün
#   python member_tags.py 2026-11-02 2026-11-08 → aralık (geriye dönük doldurma)
#   python member_tags.py rebuild-daily 2026-11-02 2026-11-08 → Günlük_Özet + Haftalık + Yoğun/Top20'yi Ham_Veri'den yeniden hesaplar (API yok)
#   python member_tags.py ratings 2026-11-02 2026-11-12 → yalnızca Rating 1-2 tablolarını (Rating_Ham / Rating_Günlük / Rating_Agent) doldurur
#   python member_tags.py fix-format           → eski '(etiketsiz)' satırlarını siler + başlık dışı biçimi sıfırlar (tek seferlik)
# ============================================================

import os
import re
import sys
import json
import time
import base64
import smtplib
import traceback
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pytz
import requests

# ── CONFIG ───────────────────────────────────────────────────
SITE_ID        = os.environ["COMM100_SITE_ID"]
API_KEY        = os.environ["COMM100_API_KEY"]
COMM100_EMAIL  = os.environ["COMM100_EMAIL"]
# Sheets / mail değişkenleri yalnızca yazma ve hata maili için gerekir; 'diag' modu (salt okunur)
# yerelde sadece üstteki 3 Comm100 değişkeniyle çalışabilsin diye boş varsayılır.
SPREADSHEET_ID = os.environ.get("MEMBER_TAGS_SPREADSHEET_ID", "")
GCP_CREDS_JSON = os.environ.get("GCP_CREDENTIALS", "")
GMAIL_USER     = os.environ.get("GMAIL_USER", "")
GMAIL_PASS     = os.environ.get("GMAIL_APP_PASSWORD", "")
REPORT_EMAILS  = [e.strip() for e in os.environ.get("REPORT_EMAILS", "").split(",") if e.strip()]

TZ = pytz.timezone("Europe/Sofia")

# Departmanlar (hassas değil — diğer scriptlerle aynı sabitler)
CS_DEPT_ID  = "29098b6e-0acb-40cc-8cd7-ebf662defc42"
VIP_DEPT_ID = "433cb80f-4e0d-4e6d-96d3-bff700296a6f"
DEPT_NAMES = {
    CS_DEPT_ID.lower():  "Customer Support",
    VIP_DEPT_ID.lower(): "VIP",
}

CAMPAIGNS = {
    "SB": "3cd6532a-0d60-4d7d-9331-5dd5be71f527",
    "BS": "2b71ad44-6a12-4073-bc6c-8edb22eabcc6",
    "TB": "6712311a-2268-408c-a803-b338a1308010",
}

# Custom Variable alanları (ad soneki ile de bulunur — id değişirse bozulmaz)
USER_CODE_FIELD_ID = "7192c568-9cca-4e6d-b68b-657d3654ee98"   # 04 User Code
USERNAME_FIELD_ID  = "99c9204e-8cfc-44f6-b3d1-4b855df01cd5"   # 01 Username
RISK_FIELD_ID      = ""                                       # 05 Risk Level (ad ile bulunur)
DETAIL_WORKERS     = 4

# Ayarlar
RAW_RETENTION_DAYS = 35      # Ham_Veri'de tutulacak gün sayısı (Sheets 10M hücre sınırı)
HEAVY_MIN_CHATS    = 4       # Aynı tagde bu kadar ve üzeri chat açan üye "yoğun" sayılır
INCLUDE_UNTAGGED   = False   # Tagsiz chatler (botta kalanlar) sayılmaz / listelenmez
# "VIP" tagi ve VIP TIER tagleri (Wild card → VIP TIER3 gibi) Risk Level ile aynı
# bilgiyi taşır; tag olarak sayılmaz / listelenmez.
EXCLUDE_TAG_RE     = re.compile(r"^VIP(\s*TIER.*)?$", re.I)

TAB_RAW     = "Ham_Veri"
TAB_MEMBER  = "Üye_Günlük"
TAB_DAILY   = "Günlük_Özet"
TAB_WEEKLY  = "Haftalık"
TAB_HEAVY   = "Yoğun_Üyeler"
TAB_DTOP    = "Günlük_Top20"
TAB_WTOP    = "Haftalık_Top20"
TAB_RHAM    = "Rating_Ham"       # puanı 1 veya 2 olan chatler (satır = chat)
TAB_RDAY    = "Rating_Günlük"    # gün x marka: toplam chat, rating 1, rating 2
TAB_RAGT    = "Rating_Agent"     # gün x marka x agent: toplam chat, rating 1, rating 2
TAB_RTAG    = "Rating_Tag"       # gün x marka x tag: rating 1, rating 2 (geçmiş veri de burada)
RATING_TABS = (TAB_RHAM, TAB_RDAY, TAB_RTAG, TAB_RAGT)
RATING_RETENTION_DAYS = 100          # Rating_Ham (chat satırları)
RATING_AGG_RETENTION_DAYS = 400      # günlük / tag / agent toplamları (geçmiş veri korunsun)
HISTORY_BRAND = "ALL"                # eski Rating Dashboard'dan gelen, marka kırılımı olmayan günler
ALL_TABS    = (TAB_RAW, TAB_MEMBER, TAB_DAILY, TAB_WEEKLY, TAB_HEAVY, TAB_DTOP, TAB_WTOP) + RATING_TABS
TOP_N       = 20          # her marka için Top N üye

HEADERS = {
    TAB_RAW:    ["Tarih", "Marka", "Departman", "Kullanıcı Adı", "User Code",
                 "Risk Level", "Tag", "Chat"],
    TAB_MEMBER: ["Tarih", "Marka", "Departman", "Kullanıcı Adı", "User Code",
                 "Risk Level", "Chat"],
    TAB_DAILY:  ["Tarih", "Marka", "Departman", "Tag", "Chat", "Tekil Üye", "Chat/Üye"],
    TAB_WEEKLY: ["Hafta", "Marka", "Departman", "Tag", "Chat", "Tekil Üye", "Chat/Üye"],
    TAB_HEAVY:  ["Hafta", "Marka", "Departman", "Kullanıcı Adı", "User Code",
                 "Risk Level", "Tag", "Chat"],
    TAB_DTOP:   ["Tarih", "Marka", "Sıra", "Departman", "Kullanıcı Adı", "User Code",
                 "Risk Level", "Chat", "Tag Dağılımı (ilk 3)"],
    TAB_WTOP:   ["Hafta", "Marka", "Sıra", "Departman", "Kullanıcı Adı", "User Code",
                 "Risk Level", "Chat", "Tag Dağılımı (ilk 3)"],
    TAB_RHAM:   ["Tarih", "Marka", "Departman", "Agent", "Tag", "Rating",
                 "User Code", "Kullanıcı Adı", "Chat ID"],
    TAB_RDAY:   ["Tarih", "Marka", "Toplam Chat", "Rating 1", "Rating 2"],
    TAB_RTAG:   ["Tarih", "Marka", "Tag", "Rating 1", "Rating 2"],
    TAB_RAGT:   ["Tarih", "Marka", "Agent", "Toplam Chat", "Rating 1", "Rating 2"],
}

# ── AUTH / DATE ──────────────────────────────────────────────
def comm100_auth():
    return "Basic " + base64.b64encode(f"{COMM100_EMAIL}:{API_KEY}".encode()).decode()

def sofia_yesterday():
    return (datetime.now(TZ) - timedelta(days=1)).strftime("%Y-%m-%d")

def sofia_to_utc_range(date_str):
    start = TZ.localize(datetime.strptime(date_str + " 00:00:00", "%Y-%m-%d %H:%M:%S"))
    end   = TZ.localize(datetime.strptime(date_str + " 23:59:59", "%Y-%m-%d %H:%M:%S"))
    fmt = "%Y-%m-%dT%H:%M:%SZ"
    return (start.astimezone(pytz.utc).strftime(fmt),
            end.astimezone(pytz.utc).strftime(fmt))

def in_sofia_range(ts_str, date_str):
    if not ts_str:
        return True
    try:
        dt = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        return dt.astimezone(TZ).strftime("%Y-%m-%d") == date_str
    except Exception:
        return True

def week_start(date_str):
    d = datetime.strptime(date_str, "%Y-%m-%d")
    return (d - timedelta(days=d.weekday())).strftime("%Y-%m-%d")

def week_end(date_str):
    d = datetime.strptime(week_start(date_str), "%Y-%m-%d")
    return (d + timedelta(days=6)).strftime("%Y-%m-%d")

EN_MONTHS3 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

def week_label(ws, we):
    """('2026-11-02', '2026-11-08') → '2 - 8 Nov 2026'"""
    d0 = datetime.strptime(ws, "%Y-%m-%d")
    d1 = datetime.strptime(we, "%Y-%m-%d")
    m0, m1 = EN_MONTHS3[d0.month - 1], EN_MONTHS3[d1.month - 1]
    if d0.year == d1.year and d0.month == d1.month:
        return f"{d0.day} - {d1.day} {m1} {d1.year}"
    if d0.year == d1.year:
        return f"{d0.day} {m0} - {d1.day} {m1} {d1.year}"
    return f"{d0.day} {m0} {d0.year} - {d1.day} {m1} {d1.year}"

def date_range(a, b):
    d0 = datetime.strptime(a, "%Y-%m-%d")
    d1 = datetime.strptime(b, "%Y-%m-%d")
    out = []
    while d0 <= d1:
        out.append(d0.strftime("%Y-%m-%d"))
        d0 += timedelta(days=1)
    return out

# ── COMM100 FETCH ────────────────────────────────────────────
SEARCH_BASE = (
    f"https://dash15.lively-chat.com/api/LiveChat/chats:search"
    f"?siteId={SITE_ID}"
    f"&include=chatWrapupCategory&include=chatFieldResult&include=campaign&include=chatAgent&include=postChatSurvey"
    f"&sortBy=startTime&sortOrder=asc"
)

def fetch_chats(date_str):
    start_time, end_time = sofia_to_utc_range(date_str)
    auth = comm100_auth()
    result, seen, page = [], set(), 1
    print(f"[FETCH] {date_str} ...")
    while page <= 60:
        url = f"{SEARCH_BASE}&pageIndex={page}&pageSize=500"
        last = None
        chats = None
        for attempt in range(3):
            try:
                t_req = time.time()
                r = requests.post(
                    url,
                    headers={"Authorization": auth, "Content-Type": "application/json"},
                    json={"startTime": start_time, "endTime": end_time},
                    timeout=90,
                )
                if r.status_code == 200:
                    chats = r.json().get("list", [])
                    print(f"[API] chat sayfa {page}: {len(chats)} kayıt, {time.time()-t_req:.1f} sn")
                    break
                last = f"HTTP {r.status_code}: {r.text[:200]}"
            except Exception as e:
                last = type(e).__name__
            time.sleep(2 * (attempt + 1))
        if chats is None:
            raise RuntimeError(f"fetch sayfa {page}: {last}")
        if not chats:
            break
        before = len(result)
        for c in chats:
            cid = str(c.get("id") or c.get("chatId") or "")
            if cid and cid not in seen:
                seen.add(cid)
                result.append(c)
        if len(chats) < 500:
            break
        # Comm100 son sayfadan sonra da dolu (tekrarlı) sayfa döndürüyor: yeni chat gelmiyorsa bitti
        if len(result) == before:
            print(f"[FETCH] sayfa {page}: yeni chat yok, tarama bitti")
            break
        page += 1
        time.sleep(0.2)
    result = [c for c in result if in_sofia_range(c.get("startTime") or c.get("start_time"), date_str)]
    print(f"[FETCH] {date_str}: {len(result)} chat")
    return result

# ── RATING 1-2 ───────────────────────────────────────────────
def rating_grade(c):
    g = (c.get("postChatSurvey") or {}).get("ratingGrade")
    try:
        return int(g) if g is not None else None
    except (TypeError, ValueError):
        return None

def agent_names(c):
    """Chatte konuşan insan agent adları (bot hariç)."""
    names = []
    def add(n):
        n = str(n or "").strip()
        if n and n not in names:
            names.append(n)
    for a in c.get("chatAgents") or []:
        if isinstance(a, dict) and not a.get("botId"):
            add((a.get("agent") or {}).get("displayName") or a.get("displayName") or a.get("name"))
    if not names:
        for a in c.get("agents") or []:
            if isinstance(a, dict) and not a.get("botId"):
                add(a.get("name") or a.get("displayName"))
    if not names:
        add(c.get("agentName") or c.get("agent_name"))
    return names

def primary_tag(c):
    """Rating tablosunda chat başına tek tag: '_' içeren ilk tag, yoksa ilk tag, yoksa '—'."""
    tags = get_tags(c)
    for t in tags:
        if "_" in t:
            return t
    return tags[0] if tags else "—"

BOT_AGENT = "BOT"    # hiçbir insan agent'ın cevaplamadığı chatler (botta kalanlar)

def agents_or_bot(c):
    return agent_names(c) or [BOT_AGENT]

def build_rating_rows(date_str, all_chats, low_chats):
    """→ (Rating_Ham, Rating_Günlük, Rating_Agent) satırları. Agent'sız chatler 'BOT' satırında toplanır."""
    ham = []
    for c in low_chats:
        ham.append([date_str, get_brand(c), get_dept(c), ", ".join(agent_names(c)) or BOT_AGENT,
                    primary_tag(c), rating_grade(c), c.get("_uc", ""), c.get("_un", ""),
                    str(c.get("id") or c.get("chatId") or "")])
    ham.sort(key=lambda r: (r[1], r[4], r[3]))

    day = defaultdict(lambda: [0, 0, 0])       # marka → [toplam, r1, r2]
    agt = defaultdict(lambda: [0, 0, 0])       # (marka, agent) → [toplam, r1, r2]
    for c in all_chats:
        b = get_brand(c)
        day[b][0] += 1
        for a in agents_or_bot(c):
            agt[(b, a)][0] += 1
    for c in low_chats:
        b, g = get_brand(c), rating_grade(c)
        day[b][g] += 1
        for a in agents_or_bot(c):
            agt[(b, a)][g] += 1
            if agt[(b, a)][0] < agt[(b, a)][1] + agt[(b, a)][2]:   # toplam, düşük puanlıdan az görünmesin
                agt[(b, a)][0] = agt[(b, a)][1] + agt[(b, a)][2]
    dayrows = [[date_str, b, v[0], v[1], v[2]] for b, v in day.items()]
    dayrows.sort(key=lambda r: _brand_order(r[1]))
    agtrows = [[date_str, b, a, v[0], v[1], v[2]] for (b, a), v in agt.items()]
    agtrows.sort(key=lambda r: (_brand_order(r[1]), -(r[4] + r[5]), -r[3], r[2]))
    return ham, dayrows, agtrows

def build_rating_tag_rows(date_str, low_chats):
    """→ Rating_Tag satırları: [Tarih, Marka, Tag, Rating 1, Rating 2] (tag '—' = etiketsiz)."""
    agg = defaultdict(lambda: [0, 0])
    for c in low_chats:
        g = rating_grade(c)
        agg[(get_brand(c), primary_tag(c))][g - 1] += 1
    rows = [[date_str, b, t, v[0], v[1]] for (b, t), v in agg.items()]
    rows.sort(key=lambda r: (_brand_order(r[1]), -(r[3] + r[4]), r[2]))
    return rows

# ── CUSTOM VARIABLES ─────────────────────────────────────────
def extract_custom_var(c, field_id, name_suffix):
    """customVariables / fieldResults içinden değer. fieldId ile, olmazsa
    'NN Ad' biçimindeki ad sonekiyle bulunur."""
    for key in ("customVariables", "fieldResults"):
        for it in c.get(key) or []:
            if not isinstance(it, dict):
                continue
            fname = str(it.get("fieldName") or "").strip().lower()
            if not ((field_id and it.get("fieldId") == field_id) or fname.endswith(name_suffix)):
                continue
            val = it.get("value")
            vals = val if isinstance(val, list) else [val]
            for v in vals:
                if v is not None and str(v).strip():
                    return str(v).strip()
    return ""

def get_member_fields(c):
    return (
        extract_custom_var(c, USERNAME_FIELD_ID, " username"),
        extract_custom_var(c, USER_CODE_FIELD_ID, "user code"),
        extract_custom_var(c, RISK_FIELD_ID, "risk level"),
    )

def _fetch_chat_detail(cid, auth):
    url = (f"https://dash15.lively-chat.com/api/LiveChat/chats/{cid}"
           f"?siteId={SITE_ID}&include=chatFieldResult&include=campaign")
    last = None
    for _ in range(3):
        try:
            r = requests.get(url, headers={"Authorization": auth}, timeout=30)
            if r.status_code == 200:
                return r.json()
            last = f"HTTP {r.status_code}"
        except Exception as e:
            last = type(e).__name__
        time.sleep(0.5)
    raise RuntimeError(last or "unknown")

def get_typed_name(c):
    """Login olmadan gelen üyenin chat formuna yazdığı ad (commqa.py ile aynı alan önceliği)."""
    return str(c.get("preChatName") or c.get("name") or
               c.get("visitorName") or c.get("visitor_name") or "").strip()

def attach_member_fields(chats):
    """Her chate c['_un'], c['_uc'], c['_risk'] ekler.
    1) Custom Variable (01 Username / 04 User Code / 05 Risk Level)
    2) Aramada custom variable hiç gelmiyorsa detay çağrısı
    3) Login olmadan gelenlerde (User Code ve Username boş) chat formuna
       yazılan ad kullanıcı adı olarak kabul edilir; User Code boş kalır."""
    for c in chats:
        c["_un"], c["_uc"], c["_risk"] = get_member_fields(c)

    n = len(chats)
    with_vars = sum(1 for c in chats if c["_un"] or c["_uc"])
    # Arama custom variable döndürüyorsa (chatlerin en az %5'inde varsa) boş olanlar
    # login olmamış sayılır; detay çağrısı yapılmaz. Hiç dönmüyorsa hepsi için detay çekilir.
    search_has_vars = n > 0 and with_vars / n >= 0.05
    todo = [] if search_has_vars else [c for c in chats if not c["_un"] and not c["_uc"]]

    errors = 0
    if todo:
        auth = comm100_auth()

        def one(c):
            cid = str(c.get("id") or c.get("chatId") or "")
            try:
                return c, get_member_fields(_fetch_chat_detail(cid, auth)), False
            except Exception:
                return c, ("", "", ""), True

        with ThreadPoolExecutor(max_workers=DETAIL_WORKERS) as ex:
            for c, (un, uc, risk), err in ex.map(one, todo):
                c["_un"], c["_uc"], c["_risk"] = un, uc, risk
                errors += int(err)

    typed = 0
    for c in chats:
        if not c["_un"] and not c["_uc"]:
            name = get_typed_name(c)
            if name:
                c["_un"] = name
                typed += 1
    print(f"[MEMBER] custom var: {with_vars}/{n} (search_has_vars={search_has_vars}) | "
          f"risk {sum(1 for c in chats if c['_risk'])}/{n} | "
          f"formdan ad (login yok): {typed} | detail calls {len(todo)}, errors {errors} | "
          f"kimliksiz: {sum(1 for c in chats if not c['_un'] and not c['_uc'])}")
    return chats

# ── CHAT → ALANLAR ───────────────────────────────────────────
def get_brand(c):
    cid = str(c.get("campaignId") or (c.get("campaign") or {}).get("id") or "").lower()
    for brand, camp_id in CAMPAIGNS.items():
        if cid == camp_id.lower():
            return brand
    name = str((c.get("campaign") or {}).get("name") or "").lower()
    for key, brand in (("superbetin", "SB"), ("betsat", "BS"), ("turkbet", "TB")):
        if key in name:
            return brand
    url = (c.get("requestingPageURL") or "").lower()
    for key, brand in (("superbetin", "SB"), ("betsat", "BS"), ("turkbet", "TB")):
        if key in url:
            return brand
    return "?"

def get_dept(c):
    did = str(c.get("departmentId") or (c.get("department") or {}).get("id") or "").lower()
    if did in DEPT_NAMES:
        return DEPT_NAMES[did]
    return "Diğer" if did else "Bilinmiyor"

def split_tags(text):
    """'a, b' → ['a', 'b']; VIP TIER* ayıklanır, tekrarlar atılır."""
    out = []
    for t in str(text).split(","):
        t = t.strip()
        if t and not EXCLUDE_TAG_RE.search(t) and t not in out:
            out.append(t)
    return out

def get_tags(c):
    """chatWrapup.categoriesName → tag listesi.
    'Promotions (goodwill_query), Wild card (VIP TIER3)' → ['goodwill_query']
    'Wild card (joker, VIP TIER3)' → ['joker']  (parantez içi virgülle ayrılabilir)
    Parantez yoksa kategori adı tag sayılır."""
    cat = (c.get("chatWrapup") or {}).get("categoriesName") or ""
    tags = []
    if cat:
        inner = [m.group(1) for m in re.finditer(r"\(([^)]+)\)", cat)]
        if not inner:
            inner = [cat]
        for chunk in inner:
            for t in split_tags(chunk):
                if t not in tags:
                    tags.append(t)
    if not tags and INCLUDE_UNTAGGED:
        tags = ["(etiketsiz)"]
    return tags

def member_key(brand, un, uc):
    ident = (uc or un or "").strip().lower()
    return f"{brand}|{ident}" if ident else ""

# ── AGREGASYON ───────────────────────────────────────────────
def resolve_identities_chats(chats):
    """Login olmadan formdan ad yazan chatler: aynı markada aynı kullanıcı adıyla
    (büyük/küçük harf fark etmez) User Code'u olan bir üye varsa o üyeye bağlanır."""
    known = {}
    for c in chats:
        if c["_uc"] and c["_un"]:
            known[(get_brand(c), c["_un"].strip().lower())] = (c["_uc"], c["_un"], c["_risk"])
    linked = 0
    for c in chats:
        if not c["_uc"] and c["_un"]:
            hit = known.get((get_brand(c), c["_un"].strip().lower()))
            if hit:
                c["_uc"], c["_un"] = hit[0], hit[1]
                if not c["_risk"]:
                    c["_risk"] = hit[2]
                linked += 1
    return linked

def resolve_identities_rows(rows):
    """Ham satırlar (haftalık hesap için) üzerinde aynı işlem."""
    known = {}
    for r in rows:
        if str(r[4]).strip() and str(r[3]).strip():
            known[(r[1], str(r[3]).strip().lower())] = (str(r[3]), str(r[4]), r[5])
    for r in rows:
        if not str(r[4]).strip() and str(r[3]).strip():
            hit = known.get((r[1], str(r[3]).strip().lower()))
            if hit:
                r[3], r[4] = hit[0], hit[1]
                if r[5] == "(bilinmiyor)":
                    r[5] = hit[2]
    return rows

def normalize_raw_rows(rows):
    """Eski satırları temizler: virgüllü tag hücrelerini ('joker, VIP TIER3') ayırır,
    dışlanan tagleri (VIP, VIP TIER*) atar."""
    out = []
    for r in rows:
        tag = str(r[6])
        if tag == "(etiketsiz)":
            out.append(r)
            continue
        for t in split_tags(tag):
            out.append(r[:6] + [t] + r[7:])
    return out

def build_raw_rows(date_str, chats):
    """(tarih, marka, dept, username, user code, risk, tag) → chat sayısı.
    Ham_Veri satırları: [Tarih, Marka, Departman, Kullanıcı, UserCode, Risk, Tag, Chat]"""
    agg = defaultdict(int)
    risk_of = {}   # (marka, kimlik) → son görülen risk (gün içinde)
    for c in sorted(chats, key=lambda x: x.get("startTime") or ""):
        brand = get_brand(c)
        un, uc, risk = c["_un"], c["_uc"], c["_risk"]
        ident = (uc or un).strip().lower()
        if ident and risk:
            risk_of[(brand, ident)] = risk
    for c in chats:
        brand, dept = get_brand(c), get_dept(c)
        un, uc = c["_un"], c["_uc"]
        ident = (uc or un).strip().lower()
        risk = risk_of.get((brand, ident), c["_risk"]) if ident else c["_risk"]
        for tag in get_tags(c):
            agg[(brand, dept, un, uc, risk or "(bilinmiyor)", tag)] += 1
    rows = [[date_str, b, d, un, uc, rk, tg, n] for (b, d, un, uc, rk, tg), n in agg.items()]
    rows.sort(key=lambda r: (r[1], r[2], r[6], -r[7]))
    return rows

def _ratio(chats, members):
    return round(chats / members, 2) if members else ""

def summarize(raw_rows, group_idx):
    """raw_rows üzerinden group_idx alanlarına göre (chat, tekil üye) hesaplar.
    group_idx: ham satırdaki sütun indeksleri (1=marka, 2=dept, 5=risk, 6=tag)."""
    chats = defaultdict(int)
    members = defaultdict(set)
    for r in raw_rows:
        key = tuple(r[i] for i in group_idx)
        chats[key] += int(r[7])
        mk = member_key(r[1], r[3], r[4])
        if mk:
            members[key].add(mk)
    return chats, members

def daily_summary_rows(date_str, raw_rows):
    out = []
    for dept_mode in ("per", "all"):
        rows = raw_rows if dept_mode == "per" else [
            [r[0], r[1], "Tümü", r[3], r[4], r[5], r[6], r[7]] for r in raw_rows]
        chats, members = summarize(rows, (1, 2, 6))
        for (b, d, tg), n in chats.items():
            m = len(members[(b, d, tg)])
            out.append([date_str, b, d, tg, n, m, _ratio(n, m)])
    out.sort(key=lambda r: (r[1], r[2] != "Tümü", r[2], -r[4]))
    return out

def weekly_rows(label, raw_rows):
    """raw_rows: o haftanın ham satırları. (haftalık satırlar, üye→son risk seviyesi) döner."""
    dept_rows = []
    for mode in ("per", "all"):
        rows = raw_rows if mode == "per" else [
            [r[0], r[1], "Tümü", r[3], r[4], r[5], r[6], r[7]] for r in raw_rows]
        chats, members = summarize(rows, (1, 2, 6))
        for (b, d, tg), n in chats.items():
            m = len(members[(b, d, tg)])
            dept_rows.append([label, b, d, tg, n, m, _ratio(n, m)])
    dept_rows.sort(key=lambda r: (r[1], r[2] != "Tümü", r[2], -r[5]))

    # Üyenin haftadaki SON risk seviyesi (yoğun üye listesinde gösterilir)
    last_risk = {}
    for r in sorted(raw_rows, key=lambda x: x[0]):
        mk = member_key(r[1], r[3], r[4])
        if mk and r[5] != "(bilinmiyor)":
            last_risk[mk] = r[5]
    return dept_rows, last_risk

def heavy_rows(label, raw_rows, last_risk):
    """Haftada aynı tagde >= HEAVY_MIN_CHATS chat açan üyeler (departmanlar birleşik)."""
    cnt = defaultdict(int)
    depts = defaultdict(set)
    info = {}
    for r in raw_rows:
        mk = member_key(r[1], r[3], r[4])
        if not mk or r[6] == "(etiketsiz)":
            continue
        k = (mk, r[6])
        cnt[k] += int(r[7])
        depts[k].add(r[2])
        info[mk] = (r[1], r[3], r[4])
    out = []
    for (mk, tag), n in cnt.items():
        if n < HEAVY_MIN_CHATS:
            continue
        b, un, uc = info[mk]
        dept = "/".join(sorted(depts[(mk, tag)]))
        out.append([label, b, dept, un, uc, last_risk.get(mk, "(bilinmiyor)"), tag, n])
    out.sort(key=lambda r: (-r[7], r[1], r[6]))
    return out

def build_member_rows(date_str, chats):
    """Üye başına TEKİL chat sayısı (tag'li chatler). Satır:
    [Tarih, Marka, Departman, Kullanıcı Adı, User Code, Risk Level, Chat]"""
    risk_of = {}
    for c in sorted(chats, key=lambda x: x.get("startTime") or ""):
        ident = (c["_uc"] or c["_un"]).strip().lower()
        if ident and c["_risk"]:
            risk_of[(get_brand(c), ident)] = c["_risk"]
    agg = {}
    for c in chats:
        if not get_tags(c):
            continue
        brand, dept = get_brand(c), get_dept(c)
        ident = (c["_uc"] or c["_un"]).strip().lower()
        if not ident:
            continue
        k = (brand, dept, ident)
        rec = agg.setdefault(k, {"un": c["_un"], "uc": c["_uc"],
                                 "risk": risk_of.get((brand, ident), c["_risk"]), "n": 0})
        rec["n"] += 1
    rows = [[date_str, b, d, v["un"], v["uc"], v["risk"] or "(bilinmiyor)", v["n"]]
            for (b, d, _), v in agg.items()]
    rows.sort(key=lambda r: (r[1], r[2], -r[6]))
    return rows

def _brand_order(b):
    order = {"SB": 0, "BS": 1, "TB": 2}
    return (order.get(b, 9), b)

def top20_rows(label, member_rows, raw_rows, n=TOP_N):
    """Her marka için en çok chat açan n üye.
    member_rows: [Tarih, Marka, Dept, Kullanıcı, UserCode, Risk, Chat]
    raw_rows:    ham satırlar (tag dağılımı için)
    Satır: [label, Marka, Sıra, Departman, Kullanıcı, UserCode, Risk, Chat, Tag Dağılımı]"""
    tot = defaultdict(int)
    depts = defaultdict(set)
    info = {}
    last_risk = {}
    for r in sorted(member_rows, key=lambda x: str(x[0])):
        mk = member_key(r[1], r[3], r[4])
        if not mk:
            continue
        tot[mk] += int(r[6])
        depts[mk].add(r[2])
        if mk not in info or (r[4] and not info[mk][2]):
            info[mk] = (r[1], r[3], r[4])
        if r[5] != "(bilinmiyor)":
            last_risk[mk] = r[5]
    tags = defaultdict(lambda: defaultdict(int))
    for r in raw_rows:
        mk = member_key(r[1], r[3], r[4])
        if mk:
            tags[mk][r[6]] += int(r[7])
    by_brand = defaultdict(list)
    for mk, total in tot.items():
        by_brand[info[mk][0]].append((total, mk))
    out = []
    for brand in sorted(by_brand, key=_brand_order):
        ranked = sorted(by_brand[brand], key=lambda x: (-x[0], info[x[1]][1].lower()))[:n]
        for i, (total, mk) in enumerate(ranked, 1):
            b, un, uc = info[mk]
            td = sorted(tags[mk].items(), key=lambda x: (-x[1], x[0]))[:3]
            out.append([label, b, i, "/".join(sorted(depts[mk])), un, uc,
                        last_risk.get(mk, "(bilinmiyor)"), total,
                        " · ".join(f"{t} ({c})" for t, c in td)])
    return out

# ── GOOGLE SHEETS ────────────────────────────────────────────
def sheets_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_info(
        json.loads(GCP_CREDS_JSON),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
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

def get_tab_ids(svc):
    meta = _exec(svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID))
    return {s["properties"]["title"]: s["properties"]["sheetId"] for s in meta["sheets"]}

def _header_format_requests(sid, ncols):
    return [
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {}, "fields": "userEnteredFormat"}},
        {"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                      "startColumnIndex": 0, "endColumnIndex": ncols},
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "backgroundColor": {"red": 0.17, "green": 0.24, "blue": 0.35}}},
            "fields": "userEnteredFormat(textFormat,backgroundColor)"}},
    ]

def ensure_tab(svc, title, ids):
    """Sekme yoksa oluşturur (başlık + dondurma + kalın). Varsa ve başlık satırı
    güncel şemayla uyuşmuyorsa içeriği temizleyip yeni başlığı yazar (şema göçü);
    veri sonraki hesaplamada yeniden dolar. sheetId döner."""
    hdr = HEADERS[title]
    if title in ids:
        cur = _exec(svc.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1:Z1")).get("values", [[]])
        cur = cur[0] if cur else []
        # yalnızca tablo sütunlarına bak: K1'deki Top 10 bloğu başlık sayılmasın
        if cur[:len(hdr)] != hdr:
            _exec(svc.spreadsheets().values().clear(
                spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1:{chr(64 + len(hdr))}", body={}))
            _exec(svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1",
                valueInputOption="RAW", body={"values": [hdr]}))
            _exec(svc.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={"requests": _header_format_requests(ids[title], len(hdr))}))
            print(f"[SCHEMA] '{title}' başlığı güncellendi, içerik yeniden hesaplanacak")
        return ids[title]
    resp = _exec(svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"addSheet": {"properties": {
            "title": title, "gridProperties": {"frozenRowCount": 1}}}}]}))
    sid = resp["replies"][0]["addSheet"]["properties"]["sheetId"]
    ids[title] = sid
    _exec(svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1",
        valueInputOption="RAW", body={"values": [hdr]}))
    _exec(svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests":
        _header_format_requests(sid, len(hdr)) + [
        {"autoResizeDimensions": {"dimensions": {
            "sheetId": sid, "dimension": "COLUMNS",
            "startIndex": 0, "endIndex": len(hdr)}}}]}))
    return sid

def read_col(svc, title, col="A"):
    """Sütunu okur. Elle yapıştırılan gerçek tarih hücreleri (seri numarası) da
    'YYYY-MM-DD' metnine çevrilir; böylece eşleştirme/silme her durumda çalışır."""
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!{col}:{col}",
        valueRenderOption="UNFORMATTED_VALUE"))
    out = []
    for r in resp.get("values", []):
        v = r[0] if r else ""
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            v = _norm_date(v)
        out.append(str(v))
    return out

def delete_row_ranges(svc, sid, ranges):
    """ranges: [(start0, end0_excl)] — 0 tabanlı; sondan başa silinir (kayma olmasın)."""
    if not ranges:
        return
    reqs = [{"deleteDimension": {"range": {
        "sheetId": sid, "dimension": "ROWS", "startIndex": a, "endIndex": b}}}
        for a, b in sorted(ranges, reverse=True)]
    for i in range(0, len(reqs), 100):
        _exec(svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body={"requests": reqs[i:i + 100]}))

def _contiguous(idx):
    out = []
    for i in sorted(idx):
        if out and out[-1][1] == i:
            out[-1][1] = i + 1
        else:
            out.append([i, i + 1])
    return [tuple(x) for x in out]

def delete_rows_where(svc, title, sid, pred):
    """A sütunundaki değere göre pred(value) True olan satırları siler (başlık hariç)."""
    col = read_col(svc, title)
    idx = [i for i, v in enumerate(col) if i > 0 and pred(v)]
    delete_row_ranges(svc, sid, _contiguous(idx))
    return len(idx)

def append_rows(svc, title, rows, chunk=4000):
    for i in range(0, len(rows), chunk):
        _exec(svc.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A1",
            valueInputOption="RAW", insertDataOption="OVERWRITE",
            body={"values": rows[i:i + chunk]}))

def replace_rows(svc, title, sid, pred, rows):
    """A sütunu pred(value) olan eski satırları yenileriyle değiştirir.
    Önce yeni satırlar eklenir, sonra eskiler silinir: Sheets 'tüm veri satırlarını
    silemez' hatası vermesin ve yazma yarıda kesilirse veri kaybolmasın."""
    col = read_col(svc, title)
    idx = [i for i, v in enumerate(col) if i > 0 and pred(v)]
    append_rows(svc, title, rows)
    delete_row_ranges(svc, sid, _contiguous(idx))
    return len(idx)

def _norm_date(v):
    """Sheets'e elle yapıştırılan tarih seri numarası (46000 gibi) gelirse YYYY-MM-DD'ye çevirir."""
    if isinstance(v, (int, float)):
        return (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime("%Y-%m-%d")
    return str(v)

def read_raw_for_week(svc, wk_start, wk_end):
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{TAB_RAW}'!A2:H",
        valueRenderOption="UNFORMATTED_VALUE"))
    out = []
    for r in resp.get("values", []):
        r = list(r) + [""] * (8 - len(r))
        r[0] = _norm_date(r[0])
        if not INCLUDE_UNTAGGED and r[6] == "(etiketsiz)":
            continue
        if wk_start <= r[0] <= wk_end:
            r[7] = int(r[7] or 0)
            r[3] = str(r[3]); r[4] = str(r[4])
            out.append(r)
    return normalize_raw_rows(resolve_identities_rows(out))

def read_members_for_week(svc, wk_start, wk_end):
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{TAB_MEMBER}'!A2:G",
        valueRenderOption="UNFORMATTED_VALUE"))
    out = []
    for r in resp.get("values", []):
        r = list(r) + [""] * (7 - len(r))
        r[0] = _norm_date(r[0])
        if wk_start <= r[0] <= wk_end:
            r[6] = int(r[6] or 0)
            r[3] = str(r[3]); r[4] = str(r[4])
            out.append(r)
    return resolve_identities_rows(out)

# ── ANA AKIŞ ─────────────────────────────────────────────────
def process_day(svc, ids, date_str):
    chats = attach_member_fields(fetch_chats(date_str))
    print(f"[MEMBER] formdan yazılan ad → User Code'lu üyeye bağlandı: {resolve_identities_chats(chats)}")
    raw = build_raw_rows(date_str, chats)
    daily = daily_summary_rows(date_str, raw)
    members = build_member_rows(date_str, chats)
    dtop = top20_rows(date_str, members, raw)
    print(f"[DAY] {date_str}: {len(chats)} chat → ham {len(raw)}, üye {len(members)}, "
          f"günlük özet {len(daily)}, top20 {len(dtop)} satır")

    # idempotent: aynı günün satırlarını yenisiyle değiştir
    for title, rows in ((TAB_RAW, raw), (TAB_MEMBER, members),
                        (TAB_DAILY, daily), (TAB_DTOP, dtop)):
        replace_rows(svc, title, ids[title], lambda v, d=date_str: v == d, rows)
    try:
        process_ratings(svc, ids, date_str, chats)
    except Exception as e:
        RATING_ERRORS.append((date_str, str(e), traceback.format_exc()))
        print(f"[RATING] {date_str} HATA (tag verisi etkilenmedi): {e}")
    return len(chats)

RATING_ERRORS = []

def has_history_day(svc, date_str):
    """Rating_Günlük'te bu güne ait 'ALL' (içe aktarılmış geçmiş) satırı var mı?"""
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{TAB_RDAY}'!A2:B",
        valueRenderOption="UNFORMATTED_VALUE"))
    for r in resp.get("values", []):
        if len(r) >= 2 and _norm_date(r[0]) == date_str and str(r[1]).strip() == HISTORY_BRAND:
            return True
    return False

def process_ratings(svc, ids, date_str, chats=None):
    """Rating 1-2 tablolarını o gün için yeniden yazar. chats verilmişse (üye alanları eklenmiş)
    toplamlar oradan alınır; verilmemişse tüm chatler yalnızca sayım için çekilir."""
    t_a = time.time()
    all_chats = chats if chats is not None else fetch_chats(date_str)
    t_b = time.time()
    low = [c for c in all_chats if rating_grade(c) in (1, 2)]   # Comm100 ratingGrade filtresini uygulamıyor: yerelde süz
    t_c = time.time()
    missing = [c for c in low if "_uc" not in c]
    if missing:
        attach_member_fields(missing)
    if not low and has_history_day(svc, date_str):
        print(f"[RATING] {date_str}: API düşük puan döndürmedi ama bu gün eski dashboard'dan "
              f"içe aktarılmış ('{HISTORY_BRAND}') → tarihsel veri korundu, yazılmadı")
        return
    ham, dayrows, agtrows = build_rating_rows(date_str, all_chats, low)
    tagrows = build_rating_tag_rows(date_str, low)
    for title, rows in ((TAB_RHAM, ham), (TAB_RDAY, dayrows), (TAB_RTAG, tagrows), (TAB_RAGT, agtrows)):
        if rows:
            replace_rows(svc, title, ids[title], lambda v, d=date_str: v == d, rows)
        else:
            delete_rows_where(svc, title, ids[title], lambda v, d=date_str: v == d)
    print(f"[RATING] {date_str}: {len(low)} düşük puan / {len(all_chats)} chat → ham {len(ham)}, "
          f"günlük {len(dayrows)}, tag {len(tagrows)}, agent {len(agtrows)} satır")
    print(f"[TIME] {date_str}: chat çekme {t_b-t_a:.0f} sn · rating süzme {t_c-t_b:.0f} sn · "
          f"Sheet yazma {time.time()-t_c:.0f} sn")

def refresh_week(svc, ids, any_date_in_week, upto):
    ws, we = week_start(any_date_in_week), week_end(any_date_in_week)
    label = week_label(ws, we)
    raw = read_raw_for_week(svc, ws, we)
    if not raw:
        print(f"[WEEK] {label}: ham veri yok, atlandı")
        return
    members = read_members_for_week(svc, ws, we)
    dept_rows, last_risk = weekly_rows(label, raw)
    heavy = heavy_rows(label, raw, last_risk)
    wtop = top20_rows(label, members, raw)
    for title, rows in ((TAB_WEEKLY, dept_rows), (TAB_HEAVY, heavy), (TAB_WTOP, wtop)):
        replace_rows(svc, title, ids[title], lambda v, lb=label: v == lb, rows)
    print(f"[WEEK] {label}: haftalık {len(dept_rows)}, yoğun {len(heavy)}, top20 {len(wtop)} satır "
          f"(üye satırı: {len(members)})")

def prune_raw(svc, ids, today_str):
    cutoff = (datetime.strptime(today_str, "%Y-%m-%d")
              - timedelta(days=RAW_RETENTION_DAYS)).strftime("%Y-%m-%d")
    for title in (TAB_RAW, TAB_MEMBER):
        n = delete_rows_where(svc, title, ids[title], lambda v: bool(v) and v < cutoff)
        if n:
            print(f"[PRUNE] {title}: {n} satır silindi (< {cutoff})")
    rcut = (datetime.strptime(today_str, "%Y-%m-%d")
            - timedelta(days=RATING_RETENTION_DAYS)).strftime("%Y-%m-%d")
    acut = (datetime.strptime(today_str, "%Y-%m-%d")
            - timedelta(days=RATING_AGG_RETENTION_DAYS)).strftime("%Y-%m-%d")
    for title in RATING_TABS:
        cut = rcut if title == TAB_RHAM else acut
        n = delete_rows_where(svc, title, ids[title], lambda v, cut=cut: bool(v) and v < cut)
        if n:
            print(f"[PRUNE] {title}: {n} satır silindi (< {cut})")

def send_error_email(date_str, error_msg, tb_str):
    html = f"""<!DOCTYPE html><html><body style="font-family:monospace;padding:20px;">
<h2 style="color:#c0392b;">Member Tags — Hata Raporu</h2>
<p><strong>Tarih:</strong> {date_str}</p>
<p><strong>Hata:</strong> {error_msg}</p>
<pre style="background:#f8f9fa;padding:16px;border-radius:6px;font-size:12px;
     white-space:pre-wrap;word-break:break-all;">{tb_str}</pre>
<p style="color:#666;font-size:11px;">CS Operations &amp; Technology</p>
</body></html>"""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"⚠️ Member Tags ERROR | {date_str}"
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

def remove_untagged(svc, ids):
    """Eski '(etiketsiz)' ve dışlanan tag (VIP, VIP TIER*) satırlarını Ham_Veri ve Günlük_Özet'ten siler (tek seferlik)."""
    for title, col in ((TAB_RAW, "G"), (TAB_DAILY, "D")):
        vals = read_col(svc, title, col)
        idx = [i for i, v in enumerate(vals)
               if i > 0 and (v == "(etiketsiz)" or EXCLUDE_TAG_RE.search(str(v).strip()))]
        delete_row_ranges(svc, ids[title], _contiguous(idx))
        print(f"[CLEAN] {title}: {len(idx)} satır silindi ('(etiketsiz)', 'VIP', 'VIP TIER*')")

STYLE = {
    # widths: piksel · align: sütun → CENTER/RIGHT · risk: Risk Level sütunu · grad: renk skalası sütunları
    TAB_RAW:    dict(widths=[95, 60, 150, 170, 90, 100, 210, 60],       align={1: "CENTER", 7: "RIGHT"}, risk=5, grad=[], zebra="day"),
    TAB_MEMBER: dict(widths=[95, 60, 150, 170, 90, 100, 60],            align={1: "CENTER", 6: "RIGHT"}, risk=5, grad=[], zebra="day"),
    TAB_DAILY:  dict(widths=[95, 60, 110, 230, 60, 80, 80],             align={1: "CENTER", 4: "RIGHT", 5: "RIGHT", 6: "RIGHT"}, risk=None, grad=[4, 5], zebra="day", ratio=6),
    TAB_WEEKLY: dict(widths=[130, 60, 110, 230, 60, 80, 80],            align={1: "CENTER", 4: "RIGHT", 5: "RIGHT", 6: "RIGHT"}, risk=None, grad=[4, 5], zebra=True, ratio=6),
    TAB_HEAVY:  dict(widths=[130, 60, 150, 170, 90, 100, 210, 60],      align={1: "CENTER", 7: "RIGHT"}, risk=5, grad=[7], zebra=True),
    TAB_DTOP:   dict(widths=[95, 60, 50, 150, 170, 90, 100, 60, 520],   align={1: "CENTER", 2: "CENTER", 7: "RIGHT"}, risk=6, grad=[7], zebra="day"),
    TAB_WTOP:   dict(widths=[130, 60, 50, 150, 170, 90, 100, 60, 520],  align={1: "CENTER", 2: "CENTER", 7: "RIGHT"}, risk=6, grad=[7], zebra=True),
    TAB_RHAM:   dict(widths=[95, 60, 150, 200, 210, 60, 90, 170, 260],  align={1: "CENTER", 5: "CENTER"}, risk=None, grad=[], zebra="day"),
    TAB_RDAY:   dict(widths=[95, 60, 100, 80, 80],                      align={1: "CENTER", 2: "RIGHT", 3: "RIGHT", 4: "RIGHT"}, risk=None, grad=[], zebra="day"),
    TAB_RTAG:   dict(widths=[95, 60, 230, 80, 80],                      align={1: "CENTER", 3: "RIGHT", 4: "RIGHT"}, risk=None, grad=[], zebra="day"),
    TAB_RAGT:   dict(widths=[95, 60, 200, 100, 80, 80],                 align={1: "CENTER", 3: "RIGHT", 4: "RIGHT", 5: "RIGHT"}, risk=None, grad=[], zebra="day"),
}

def _rgb(h):
    h = h.lstrip("#")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}

def style_tabs(svc, ids):
    """Görünüm: Arial 10, sütun genişlikleri, hizalama, satır şeritleri, Risk Level vurgusu
    ve Chat sütunlarında renk skalası. Koşullu biçim kuralları her seferinde baştan kurulur."""
    meta = _exec(svc.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        fields="sheets(properties(sheetId,title),conditionalFormats)"))
    existing = {sh["properties"]["title"]: len(sh.get("conditionalFormats", []))
                for sh in meta["sheets"]}
    reqs = []
    for title, st in STYLE.items():
        sid, n = ids[title], len(HEADERS[title])
        for _ in range(existing.get(title, 0)):
            reqs.append({"deleteConditionalFormatRule": {"sheetId": sid, "index": 0}})
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid},
            "cell": {"userEnteredFormat": {"textFormat": {"fontFamily": "Arial", "fontSize": 10}}},
            "fields": "userEnteredFormat.textFormat.fontFamily,userEnteredFormat.textFormat.fontSize"}})
        # başlık: kalın beyaz (yazı tipi ayarı üstüne yeniden uygulanır)
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1, "endColumnIndex": n},
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 10,
                               "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "backgroundColor": _rgb("#2b3d59"), "verticalAlignment": "MIDDLE"}},
            "fields": "userEnteredFormat(textFormat,backgroundColor,verticalAlignment)"}})
        for i, w in enumerate(st["widths"]):
            reqs.append({"updateDimensionProperties": {
                "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": i, "endIndex": i + 1},
                "properties": {"pixelSize": w}, "fields": "pixelSize"}})
        for col, al in st["align"].items():
            reqs.append({"repeatCell": {
                "range": {"sheetId": sid, "startColumnIndex": col, "endColumnIndex": col + 1},
                "cell": {"userEnteredFormat": {"horizontalAlignment": al}},
                "fields": "userEnteredFormat.horizontalAlignment"}})
        if st.get("ratio") is not None:
            c = st["ratio"]
            reqs.append({"repeatCell": {
                "range": {"sheetId": sid, "startRowIndex": 1, "startColumnIndex": c, "endColumnIndex": c + 1},
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.00"}}},
                "fields": "userEnteredFormat.numberFormat"}})
        def rng(c0, c1):
            return {"sheetId": sid, "startRowIndex": 1, "startColumnIndex": c0, "endColumnIndex": c1}
        if st["zebra"]:
            # "day": her gün kendi renginde blok (Tarih metni → DATEVALUE, virgülsüz); aksi halde satır şeridi
            zf = "=ISEVEN(DATEVALUE($A2))" if st["zebra"] == "day" else "=ISEVEN(ROW())"
            reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
                "ranges": [rng(0, n)],
                "booleanRule": {"condition": {"type": "CUSTOM_FORMULA",
                                              "values": [{"userEnteredValue": zf}]},
                                "format": {"backgroundColor": _rgb("#e8eef7")}}}}})
        if st["risk"] is not None:
            L = chr(65 + st["risk"])
            reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
                "ranges": [rng(st["risk"], st["risk"] + 1)],
                "booleanRule": {"condition": {"type": "CUSTOM_FORMULA", "values": [{"userEnteredValue":
                                # virgülsüz yazım: Sheet dili Türkçe ise ',' ayracı geçersiz sayılır
                                f'=({L}2<>"")*({L}2<>"Normal")*({L}2<>"(bilinmiyor)")'}]},
                                "format": {"backgroundColor": _rgb("#fde9c8"),
                                           "textFormat": {"foregroundColor": _rgb("#8a4b00")}}}}}})
        for c in st["grad"]:
            reqs.append({"addConditionalFormatRule": {"index": 0, "rule": {
                "ranges": [rng(c, c + 1)],
                "gradientRule": {"minpoint": {"color": _rgb("#ffffff"), "type": "MIN"},
                                 "maxpoint": {"color": _rgb("#9fc0e6"), "type": "MAX"}}}}})
    for i in range(0, len(reqs), 80):
        _exec(svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body={"requests": reqs[i:i + 80]}))
    print("[STYLE] Görünüm uygulandı (yazı tipi, genişlikler, şeritler, Risk Level, renk skalası)")

def fix_format(svc, ids):
    """Tek seferlik: başlık hariç tüm satırların biçimini sıfırlar, başlığı yeniden boyar."""
    reqs = []
    for t in ALL_TABS:
        sid = ids[t]
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 1},
            "cell": {}, "fields": "userEnteredFormat"}})
        reqs.append({"repeatCell": {
            "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
            "cell": {"userEnteredFormat": {
                "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                "backgroundColor": {"red": 0.17, "green": 0.24, "blue": 0.35}}},
            "fields": "userEnteredFormat(textFormat,backgroundColor)"}})
    _exec(svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": reqs}))
    print("[FORMAT] Başlık dışındaki satırların biçimi sıfırlandı")
    normalize_date_cols(svc, ids)
    style_tabs(svc, ids)

def normalize_date_cols(svc, ids):
    """Elle yapıştırılan günlerde Tarih sütunu seri numarası (46328) olarak durur.
    A sütununu 'YYYY-MM-DD' metnine çevirip geri yazar (RAW → metin kalır)."""
    for title in (TAB_RAW, TAB_MEMBER, TAB_DAILY, TAB_DTOP) + RATING_TABS:
        col = read_col(svc, title)[1:]          # başlık hariç; seri numaraları zaten çevrilmiş gelir
        if not col:
            continue
        for i in range(0, len(col), 10000):
            part = col[i:i + 10000]
            _exec(svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"'{title}'!A{i + 2}:A{i + 1 + len(part)}",
                valueInputOption="RAW", body={"values": [[v] for v in part]}))
        print(f"[DATES] {title}: {len(col)} tarih hücresi metne çevrildi")

SORT_TABS = {TAB_RAW: 8, TAB_MEMBER: 7, TAB_DAILY: 7, TAB_DTOP: 9, TAB_RHAM: 9, TAB_RDAY: 5, TAB_RTAG: 5, TAB_RAGT: 6}

def sort_by_date(svc, ids):
    """Günlük sekmeleri Tarih'e göre eskiden yeniye sıralar (yalnızca sıra bozuksa).
    Python sıralaması kararlıdır: gün içindeki sıra (chat azalan, Top20 sırası) korunur."""
    for title, n in SORT_TABS.items():
        L = chr(64 + n)
        resp = _exec(svc.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A2:{L}",
            valueRenderOption="UNFORMATTED_VALUE"))
        rows = [list(r) + [""] * (n - len(r)) for r in resp.get("values", [])]
        if not rows:
            continue
        keys = [_norm_date(r[0]) if r[0] != "" else "" for r in rows]
        if keys == sorted(keys):
            print(f"[SORT] {title}: sıra zaten doğru")
            continue
        order = sorted(range(len(rows)), key=lambda i: keys[i])
        out = [[keys[i]] + rows[i][1:] for i in order]
        for i in range(0, len(out), 10000):
            part = out[i:i + 10000]
            _exec(svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=f"'{title}'!A{i + 2}:{L}{i + 1 + len(part)}",
                valueInputOption="RAW", body={"values": part}))
        print(f"[SORT] {title}: {len(out)} satır tarihe göre sıralandı")

BRAND_NAMES = {"SB": "Superbetin", "BS": "Betsat", "TB": "Turkbet"}
BLOCK_OFFSET = {"SB": 0, "BS": 5, "TB": 10}      # K, P, U sütunları
BLOCK_COL0 = 10                                  # K (0 tabanlı)
BLOCK_WIDTHS = [40, 190, 55, 75, 18, 40, 190, 55, 75, 18, 40, 190, 55, 75]

def fmt_day(d):
    dt = datetime.strptime(d, "%Y-%m-%d")
    return f"{dt.day} {EN_MONTHS3[dt.month - 1]} {dt.year}"

def _read_summary(svc, title):
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A2:G",
        valueRenderOption="UNFORMATTED_VALUE"))
    out = []
    for r in resp.get("values", []):
        r = list(r) + [""] * (7 - len(r))
        r[0] = _norm_date(r[0]) if r[0] != "" else ""
        out.append(r)
    return out

MEMBER_SEC = 14          # üye bloğunun başlık satırı (0 tabanlı → sayfada 15. satır)
RISK_PLAIN = ("", "Normal", "(bilinmiyor)")

def write_top10(svc, ids, title, caption, rows, mcaption, mrows):
    """Sekmenin K1'inden başlayan iki blok:
    1) Top 10 Tag (marka başına: Sıra, Tag, Chat, Tekil Üye) — 'Tümü' departman satırlarından
    2) Top 10 Üye (marka başına: Sıra, Kullanıcı Adı, Chat, Risk Level) — Top20 sekmesinden
    rows: özet satırları [anahtar, Marka, Departman, Tag, Chat, Tekil Üye, Oran]
    mrows: Top20 satırları [anahtar, Marka, Sıra, Dept, Kullanıcı, UserCode, Risk, Chat, Tag]"""
    sid = ids[title]
    W = len(BLOCK_WIDTHS)
    grid = [[""] * W for _ in range(MEMBER_SEC + 13)]
    grid[0][0] = caption
    grid[MEMBER_SEC][0] = mcaption
    risk_cells = []
    for b, o in BLOCK_OFFSET.items():
        sel = [r for r in rows if r[1] == b and r[2] == "Tümü"]
        sel.sort(key=lambda r: (-int(r[4] or 0), str(r[3])))
        grid[1][o] = f"{BRAND_NAMES[b]} ({b})"
        grid[2][o:o + 4] = ["Sıra", "Tag", "Chat", "Tekil Üye"]
        for i, r in enumerate(sel[:10]):
            grid[3 + i][o:o + 4] = [i + 1, r[3], int(r[4] or 0), int(r[5] or 0)]
        msel = sorted([r for r in mrows if r[1] == b], key=lambda r: int(r[2] or 0))[:10]
        m0 = MEMBER_SEC
        grid[m0 + 1][o] = f"{BRAND_NAMES[b]} ({b})"
        grid[m0 + 2][o:o + 4] = ["Sıra", "Kullanıcı Adı", "Chat", "Risk Level"]
        for i, r in enumerate(msel):
            grid[m0 + 3 + i][o:o + 4] = [i + 1, str(r[4]), int(r[7] or 0), str(r[6])]
            if str(r[6]) not in RISK_PLAIN:
                risk_cells.append((m0 + 3 + i, o + 3))
    _exec(svc.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!K1:Z60", body={}))
    _exec(svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!K1", valueInputOption="RAW",
        body={"values": grid}))

    def rg(r0, r1, c0, c1):
        return {"sheetId": sid, "startRowIndex": r0, "endRowIndex": r1,
                "startColumnIndex": BLOCK_COL0 + c0, "endColumnIndex": BLOCK_COL0 + c1}
    BOLD = {"bold": True, "fontFamily": "Arial", "fontSize": 10}
    reqs = [
        {"repeatCell": {"range": rg(1, 60, 0, 16), "cell": {}, "fields": "userEnteredFormat"}},
        {"repeatCell": {"range": rg(1, MEMBER_SEC + 13, 0, W), "cell": {"userEnteredFormat": {
            "textFormat": {"fontFamily": "Arial", "fontSize": 10}}},
            "fields": "userEnteredFormat.textFormat"}},
    ]
    for t0 in (0, MEMBER_SEC):                       # iki bölümün koyu başlık çubuğu
        reqs.append({"repeatCell": {"range": rg(t0, t0 + 1, 0, W), "cell": {"userEnteredFormat": {
            "backgroundColor": _rgb("#2b3d59"), "verticalAlignment": "MIDDLE",
            "textFormat": {"bold": True, "fontFamily": "Arial", "fontSize": 10,
                           "foregroundColor": {"red": 1, "green": 1, "blue": 1}}}},
            "fields": "userEnteredFormat(backgroundColor,verticalAlignment,textFormat)"}})
    for t0 in (0, MEMBER_SEC):
        for b, o in BLOCK_OFFSET.items():
            last_col_align = "CENTER" if t0 == MEMBER_SEC else "RIGHT"
            reqs += [
                {"repeatCell": {"range": rg(t0 + 1, t0 + 2, o, o + 4), "cell": {"userEnteredFormat": {
                    "backgroundColor": _rgb("#c9d8ee"), "textFormat": BOLD}},
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
                {"repeatCell": {"range": rg(t0 + 2, t0 + 3, o, o + 4), "cell": {"userEnteredFormat": {
                    "backgroundColor": _rgb("#eef3fa"), "textFormat": BOLD}},
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"}},
                {"repeatCell": {"range": rg(t0 + 2, t0 + 13, o, o + 1), "cell": {"userEnteredFormat": {
                    "horizontalAlignment": "CENTER"}}, "fields": "userEnteredFormat.horizontalAlignment"}},
                {"repeatCell": {"range": rg(t0 + 2, t0 + 13, o + 2, o + 3), "cell": {"userEnteredFormat": {
                    "horizontalAlignment": "RIGHT"}}, "fields": "userEnteredFormat.horizontalAlignment"}},
                {"repeatCell": {"range": rg(t0 + 2, t0 + 13, o + 3, o + 4), "cell": {"userEnteredFormat": {
                    "horizontalAlignment": last_col_align}}, "fields": "userEnteredFormat.horizontalAlignment"}},
                {"repeatCell": {"range": rg(t0 + 3, t0 + 13, o + 2, o + 3), "cell": {"userEnteredFormat": {
                    "textFormat": BOLD}}, "fields": "userEnteredFormat.textFormat"}},
                {"updateBorders": {"range": rg(t0 + 1, t0 + 13, o, o + 4),
                    "bottom": {"style": "SOLID", "color": _rgb("#c9d8ee")},
                    "innerHorizontal": {"style": "SOLID", "color": _rgb("#e6ebf3")}}},
            ]
    for r, c in risk_cells:                          # Normal dışı risk seviyeleri turuncu
        reqs.append({"repeatCell": {"range": rg(r, r + 1, c, c + 1), "cell": {"userEnteredFormat": {
            "backgroundColor": _rgb("#fde9c8"),
            "textFormat": {"fontFamily": "Arial", "fontSize": 10, "foregroundColor": _rgb("#8a4b00")}}},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"}})
    for c in (7, 8, 9):   # H-J boşluk sütunları: veri ile blok arası dar kalsın
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": c, "endIndex": c + 1},
            "properties": {"pixelSize": 14}, "fields": "pixelSize"}})
    for i, w in enumerate(BLOCK_WIDTHS):
        reqs.append({"updateDimensionProperties": {
            "range": {"sheetId": sid, "dimension": "COLUMNS",
                      "startIndex": BLOCK_COL0 + i, "endIndex": BLOCK_COL0 + i + 1},
            "properties": {"pixelSize": w}, "fields": "pixelSize"}})
    _exec(svc.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": reqs}))

def _read_top20(svc, title):
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{title}'!A2:I",
        valueRenderOption="UNFORMATTED_VALUE"))
    out = []
    for r in resp.get("values", []):
        r = list(r) + [""] * (9 - len(r))
        r[0] = _norm_date(r[0]) if r[0] != "" else ""
        out.append(r)
    return out

def refresh_top10(svc, ids):
    """Günlük_Özet K1: en yeni günün Top 10 tag + Top 10 üyesi;
    Haftalık K1: o günün haftasının Top 10 tag + Top 10 üyesi."""
    daily = _read_summary(svc, TAB_DAILY)
    dates = [r[0] for r in daily if r[0]]
    if not dates:
        return
    last = max(dates)
    dtop = [r for r in _read_top20(svc, TAB_DTOP) if r[0] == last]
    write_top10(svc, ids, TAB_DAILY, f"Top 10 Tag · {fmt_day(last)}",
                [r for r in daily if r[0] == last],
                f"Top 10 Üye · {fmt_day(last)}", dtop)
    lbl = week_label(week_start(last), week_end(last))
    weekly = _read_summary(svc, TAB_WEEKLY)
    wtop = [r for r in _read_top20(svc, TAB_WTOP) if r[0] == lbl]
    write_top10(svc, ids, TAB_WEEKLY, f"Top 10 Tag · {lbl}",
                [r for r in weekly if r[0] == lbl],
                f"Top 10 Üye · {lbl}", wtop)
    print(f"[TOP10] Günlük_Özet ({fmt_day(last)}) ve Haftalık ({lbl}): Top 10 tag + Top 10 üye blokları yazıldı")

def rebuild_daily(svc, ids, days):
    """API'ye gitmeden: Günlük_Özet satırlarını Ham_Veri'den yeniden hesaplar
    (elle yapıştırılmış eski günleri düzeltmek için)."""
    for d in days:
        raw = read_raw_for_week(svc, d, d)
        if not raw:
            print(f"[REBUILD] {d}: ham veri yok, atlandı")
            continue
        daily = daily_summary_rows(d, raw)
        replace_rows(svc, TAB_DAILY, ids[TAB_DAILY], lambda v, d=d: v == d, daily)
        print(f"[REBUILD] {d}: Günlük_Özet {len(daily)} satır yeniden yazıldı")

def diag(d0, d1, brand="TB", tag="casino_issue"):
    """Tanı: API'den gelen chat'lerde bir tag'in kaç chatte geçtiğini farklı yollarla sayar.
    Sheet'e YAZMAZ. Kullanım: python member_tags.py diag 2026-11-09 2026-11-15 [TB] [casino_issue]"""
    print(f"[DIAG] {d0}..{d1} marka={brand} tag={tag}")
    tot = {"chat": 0, "brand": 0, "parsed": 0, "raw": 0}
    samples, keys_seen, fmt = [], set(), {}
    for d in date_range(d0, d1):
        chats = fetch_chats(d)
        ch_b = [c for c in chats if get_brand(c) == brand]
        parsed = [c for c in ch_b if tag in get_tags(c)]
        raw = [c for c in ch_b if tag.lower() in json.dumps(c, ensure_ascii=False).lower()]
        miss = [c for c in raw if c not in parsed]
        for c in chats[:50]:
            keys_seen.update(k for k in c.keys() if "rap" in k.lower() or "tag" in k.lower() or "categ" in k.lower())
        print(f"  {d}: tüm {len(chats)} | marka {len(ch_b)} | get_tags ile {len(parsed)} | ham metinde geçen {len(raw)} | fark {len(miss)}")
        tot["chat"] += len(chats); tot["brand"] += len(ch_b); tot["parsed"] += len(parsed); tot["raw"] += len(raw)
        for c in raw:
            v = str((c.get("chatWrapup") or {}).get("categoriesName") or "(boş)")
            fmt[v] = fmt.get(v, 0) + 1
        for c in miss:
            if len(samples) >= 3:
                break
            samples.append({k: v for k, v in c.items()
                            if k in ("id", "startTime", "chatType", "status", "departmentId")
                            or any(x in k.lower() for x in ("rap", "categ", "tag"))})
    print(f"[DIAG] TOPLAM: marka chat {tot['brand']} | get_tags {tot['parsed']} | ham metinde {tot['raw']}")
    print(f"[DIAG] wrap-up/tag/kategori alan adları: {sorted(keys_seen)}")
    print("[DIAG] categoriesName biçim örnekleri (tag'i içeren chat'lerden):")
    for v, n in sorted(fmt.items(), key=lambda x: -x[1])[:8]:
        print(f"    {n:4}  {v[:140]}")
    for sm in samples:
        print("[DIAG] örnek (ham'da var, get_tags'te yok):", json.dumps(sm, ensure_ascii=False)[:500])

def main():
    t0 = time.time()
    args = sys.argv[1:]
    if args and args[0] == "fix-format":
        svc = sheets_service()
        ids = get_tab_ids(svc)
        for t in ALL_TABS:
            ensure_tab(svc, t, ids)
        remove_untagged(svc, ids)
        fix_format(svc, ids)
        sort_by_date(svc, ids)
        refresh_top10(svc, ids)
        return
    if args and args[0] == "ratings":
        if len(args) < 2:
            print("kullanım: python member_tags.py ratings BAŞLANGIÇ [BİTİŞ]"); sys.exit(2)
        days = date_range(args[1], args[2]) if len(args) >= 3 else [args[1]]
        label = f"ratings {days[0]}..{days[-1]}"
        try:
            svc = sheets_service()
            ids = get_tab_ids(svc)
            for t in ALL_TABS:
                ensure_tab(svc, t, ids)
            for d in days:
                process_ratings(svc, ids, d)
            prune_raw(svc, ids, days[-1])
            sort_by_date(svc, ids)
            style_tabs(svc, ids)
            print(f"\n[DONE] {len(days)} gün rating, {time.time()-t0:.0f} sn")
        except Exception as e:
            tb = traceback.format_exc()
            print(tb)
            send_error_email(label, str(e), tb)
            sys.exit(1)
        return
    if args and args[0] == "diag":
        diag(args[1], args[2] if len(args) > 2 else args[1],
             args[3] if len(args) > 3 else "TB", args[4] if len(args) > 4 else "casino_issue")
        return
    if args and args[0] == "rebuild-daily":
        svc = sheets_service()
        ids = get_tab_ids(svc)
        for t in ALL_TABS:
            ensure_tab(svc, t, ids)
        days = date_range(args[1], args[2]) if len(args) >= 3 else [args[1]]
        rebuild_daily(svc, ids, days)
        for ws in sorted({week_start(d) for d in days}):
            refresh_week(svc, ids, ws, upto=days[-1])
        sort_by_date(svc, ids)
        refresh_top10(svc, ids)
        return
    if len(args) >= 2:
        days = date_range(args[0], args[1])
    elif len(args) == 1:
        days = [args[0]]
    else:
        # Dün + önceki LOOKBACK_DAYS gün yeniden çekilir: geç düşen wrap-up'lar ve atlanan
        # (örn. başarısız) günler kendiliğinden düzelir. Günler idempotent yazıldığı için güvenli.
        back = max(0, int(os.environ.get("LOOKBACK_DAYS", "2") or 2))
        y = datetime.strptime(sofia_yesterday(), "%Y-%m-%d")
        days = [(y - timedelta(days=k)).strftime("%Y-%m-%d") for k in range(back, -1, -1)]
    label = days[0] if len(days) == 1 else f"{days[0]}..{days[-1]}"
    print(f"\n{'='*55}\nMEMBER TAGS — {label}\n{'='*55}\n")

    try:
        svc = sheets_service()
        ids = get_tab_ids(svc)
        new_tabs = [t for t in ALL_TABS if t not in ids]
        for t in ALL_TABS:
            ensure_tab(svc, t, ids)
        if new_tabs:
            style_tabs(svc, ids)    # yeni açılan sekmelere görünüm uygula

        total = 0
        for d in days:
            total += process_day(svc, ids, d)

        weeks = sorted({week_start(d) for d in days})
        for ws in weeks:
            refresh_week(svc, ids, ws, upto=days[-1])

        prune_raw(svc, ids, days[-1])
        sort_by_date(svc, ids)
        refresh_top10(svc, ids)
        print(f"\n[DONE] {len(days)} gün, {total} chat, {time.time()-t0:.0f} sn")
        if RATING_ERRORS:
            d0, msg0, tb0 = RATING_ERRORS[0]
            send_error_email(f"{label} (yalnızca Rating adımı)",
                             f"{len(RATING_ERRORS)} günde rating yazılamadı, tag verisi tamam. İlk hata ({d0}): {msg0}", tb0)
    except Exception as e:
        tb = traceback.format_exc()
        print(tb)
        send_error_email(label, str(e), tb)
        sys.exit(1)

if __name__ == "__main__":
    main()
