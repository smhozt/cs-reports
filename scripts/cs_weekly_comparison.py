"""
CS Weekly Comparison — her hafta "cs_weekly_comparison_<hafta>.xlsx" dosyasını üretir ve mail atar.

Şablon : templates/cs_weekly_comparison_template.xlsx  (elle doldurduğun dosyanın kendisi; biçim/formüller korunur)
Kaynak : • Comm100 Report API  → chat hacmi, bekleme, süre, rating, queue, bot-only, bot→agent (marka bazlı)
         • Tag Sheet (MEMBER_TAGS_SPREADSHEET_ID, 'Haftalık')  → agent tag top 10 (hafta + geçen hafta)
         • Goals Sheet (GOALS_SPREADSHEET_ID, 'Goals_Günlük')   → BOT TAGS ve BOT TAGS SUCCESS top 10
Çıktı  : out/cs_weekly_comparison_<hafta>.xlsx  → mail eki (+ opsiyonel Drive klasörü, DRIVE_FOLDER_ID)
         Elle girilen / API'de olmayan hücreler SARI boş bırakılır (aşağıda MANUAL_CELLS).

Kullanım:
  python cs_weekly_comparison.py                  → geçen tam hafta (Pzt–Paz), mail atar
  python cs_weekly_comparison.py 2026-11-09       → o tarihin haftası
  python cs_weekly_comparison.py calibrate 2026-11-02
        → o haftanın API değerlerini (3 departman filtresi varyantıyla) şablondaki 2. satırla yan yana yazar
Ortam: MAIL_MODE = all | me | none  (varsayılan all) · WEEKLY_DEPT_MODE = cs | cs_vip | all (varsayılan cs)
"""

import os
import re
import sys
import json
import time
import smtplib
import zipfile
import traceback
from copy import copy
from datetime import datetime, timedelta, time as dtime
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pytz
import requests

# ── CONFIG ───────────────────────────────────────────────────
SITE_ID = os.environ.get("COMM100_SITE_ID", "")
API_KEY = os.environ.get("COMM100_API_KEY", "")
COMM100_EMAIL = os.environ.get("COMM100_EMAIL", "")
GCP_CREDS_JSON = os.environ.get("GCP_CREDENTIALS", "")
TAGS_SPREADSHEET_ID = os.environ.get("MEMBER_TAGS_SPREADSHEET_ID", "")
GOALS_SPREADSHEET_ID = os.environ.get("GOALS_SPREADSHEET_ID", "")
GMAIL_USER = os.environ.get("GMAIL_USER", "")
GMAIL_PASS = os.environ.get("GMAIL_APP_PASSWORD", "")
_mails = os.environ.get("WEEKLY_REPORT_EMAILS") or os.environ.get("REPORT_EMAILS", "")
REPORT_EMAILS = [e.strip() for e in _mails.split(",") if e.strip()]
MAIL_MODE = (os.environ.get("MAIL_MODE") or "all").strip().lower()          # all | me | none
DEPT_MODE = (os.environ.get("WEEKLY_DEPT_MODE") or "cs").strip().lower()    # cs | cs_vip | all
EFF_DEPT_MODE = (os.environ.get("WEEKLY_EFF_DEPT_MODE") or "cs").strip().lower()   # yanıt/süre için: sadece CS (CS+VIP istenirse cs_vip)
BOT_TO_AGENT_MODE = (os.environ.get("BOT_TO_AGENT_MODE") or "accepted").strip().lower()   # accepted | request
DRIVE_FOLDER_ID = os.environ.get("DRIVE_FOLDER_ID", "")
TEMPLATE = os.environ.get("WEEKLY_TEMPLATE") or "templates/cs_weekly_comparison_template.xlsx"
OUT_DIR = os.environ.get("WEEKLY_OUT_DIR") or "out"

HOST = "https://dash15.lively-chat.com"
REPORT_URLS = [
    f"{HOST}/api/reportingquery/reports/query?siteId={SITE_ID}",
    f"{HOST}/api/v4/reportingquery/reports/query?siteId={SITE_ID}",
]
TZ = pytz.timezone("Europe/Sofia")
COMM100_TZ = "FLE Standard Time"

CS_DEPT = "29098b6e-0acb-40cc-8cd7-ebf662defc42"
VIP_DEPT = "433cb80f-4e0d-4e6d-96d3-bff700296a6f"
DEPTS = {"cs": [CS_DEPT], "cs_vip": [CS_DEPT, VIP_DEPT], "all": None}

# (kod, şablon sekmesi, Comm100 CampaignId LİSTESİ: eski + yeni "Copy:" kampanya)
# Eski chat'ler eski ID'yi taşıdığı için eskileri SİLME. Turkbet kopyası açılınca yeni ID'yi listeye ekle.
BRANDS = [
    ("SB", "SUPERBETIN", ["3cd6532a-0d60-4d7d-9331-5dd5be71f527", "2c1c5c27-4f31-4ee6-90f6-1da8747e6786"]),
    ("BS", "BETSAT", ["2b71ad44-6a12-4073-bc6c-8edb22eabcc6", "ec7a1e4d-0822-4682-a25f-e7a2b29637f8"]),
    ("TB", "TURKBET", ["6712311a-2268-408c-a803-b338a1308010"]),
]

GOALS_TAB = "Goals_Günlük"
TOP_N = 10
MIN_BOT_SUCCESS = 400          # BOT TAGS SUCCESS listesine girmek için en az bot adedi

# Bot tek "Special Promotions" goal'ü atar; agent'lar özel tag'ler kullanır → tek satırda toplanır
PROMO_ROW = "special promotions"
PROMO_TAGS = {"casino_private_bonus_query", "sport_private_bonus_query", "big_reactivation_query",
              "monthly_boost_query", "daily_reactivation_query", "taraftar_bonus_query",
              "calendar_bonus_query", "special_event_bonus_query", "lapsed_reactivation_query"}
GOAL_TO_TAG = {"deposit_check_query": "deposit_check"}
SKIP_TAGS = {"", "—", "-"}

# Şablondaki konumlar
ALL_SHEET = "ALL CS"
ROWS = {"prev_stat": 7, "cur_stat": 8, "prev_bot": 13, "cur_bot": 14}
TAG_TABLE = {            # sekme → (başlık satırı, ilk veri satırı)
    "ALL CS": (27, 28), "SUPERBETIN": (25, 26), "BETSAT": (24, 25), "TURKBET": (24, 25)}
BOT_TABLE = (46, 47)      # ALL CS: BOT TAGS
SUCC_TABLE = (58, 59)     # ALL CS: BOT TAGS SUCCESS
# API'den gelmeyen, elle girilen hücreler (her iki hafta satırında sarı + boş bırakılır)
MANUAL_CELLS = {
    "ALL CS": ["S", "T", "U"],                                   # shifts, call attempts, call reached
}
BOT_MANUAL = {}           # chatbot bloğunda elle kalan hücre yok (F, I, J, K hepsi API'den)
# AI Agent > Rating raporu: BotId + ChatType 1,4 (sadece bot chat'leri), puan > 0
BOT_IDS = {"SB": "8298c1fa-c16d-40ef-9b3b-65dfd01609a2",
           "BS": "ff1a7879-ff4d-4469-837a-065f6416f1b0",
           "TB": "6a2130f6-5067-44f5-b299-893ac062b14e"}

MONTHS3 = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ── TARİH ────────────────────────────────────────────────────
def sofia_today():
    return datetime.now(TZ).strftime("%Y-%m-%d")

def week_start(date_str):
    d = datetime.strptime(date_str, "%Y-%m-%d")
    return (d - timedelta(days=d.weekday())).strftime("%Y-%m-%d")

def week_end(ws):
    return (datetime.strptime(ws, "%Y-%m-%d") + timedelta(days=6)).strftime("%Y-%m-%d")

def add_days(date_str, n):
    return (datetime.strptime(date_str, "%Y-%m-%d") + timedelta(days=n)).strftime("%Y-%m-%d")

def last_full_week():
    return add_days(week_start(sofia_today()), -7)

def label_short(ws):
    """Şablondaki etiket: '2 - 8 Nov', '26 Oct - 1 Nov', '29 Dec - 4 Jan'."""
    we = week_end(ws)
    d0, d1 = datetime.strptime(ws, "%Y-%m-%d"), datetime.strptime(we, "%Y-%m-%d")
    if d0.month == d1.month:
        return f"{d0.day} - {d1.day} {MONTHS3[d1.month - 1]}"
    return f"{d0.day} {MONTHS3[d0.month - 1]} - {d1.day} {MONTHS3[d1.month - 1]}"

def label_tags(ws):
    """Tag Sheet'indeki Haftalık etiketi: '2 - 8 Nov 2026'."""
    we = week_end(ws)
    d0, d1 = datetime.strptime(ws, "%Y-%m-%d"), datetime.strptime(we, "%Y-%m-%d")
    if d0.year == d1.year:
        return f"{label_short(ws)} {d1.year}"
    return f"{d0.day} {MONTHS3[d0.month - 1]} {d0.year} - {d1.day} {MONTHS3[d1.month - 1]} {d1.year}"

def file_name(ws):
    we = week_end(ws)
    base = label_short(ws).replace(" - ", "-").replace(" ", "_")
    return f"cs_weekly_comparison_{base}_{we[:4]}.xlsx"


# ── COMM100 REPORT API ───────────────────────────────────────
def auth_header():
    import base64
    return "Basic " + base64.b64encode(f"{COMM100_EMAIL}:{API_KEY}".encode()).decode()

def post(payload):
    last = None
    for url in REPORT_URLS:
        for attempt in range(3):
            try:
                r = requests.post(url, json=payload, timeout=120, headers={
                    "Authorization": auth_header(), "Content-Type": "application/json",
                    "Accept": "application/json"})
            except Exception as e:
                last = f"{type(e).__name__}: {e}"
                time.sleep(2 ** attempt)
                continue
            if r.status_code == 200:
                return r.json()
            if r.status_code in (401, 403):
                raise RuntimeError(f"Report API {r.status_code}: kimlik/izin hatası (COMM100_EMAIL/API_KEY). {r.text[:200]}")
            last = f"HTTP {r.status_code}: {r.text[:200]}"
            if r.status_code == 404:
                break
            time.sleep(2 ** attempt)
    raise RuntimeError(f"Report API cevap vermedi: {last}")

def total(payload):
    return post(payload).get("total", {}) or {}

def cnd(name, field, op, values):
    return {"name": name, "fieldName": field, "operate": op, "values": values}

def f_count(name, field, conds):
    return {"name": name, "calculationType": "count", "valueType": "int", "fieldName": field,
            "conditionExpression": " & ".join(c["name"] for c in conds),
            "conditionMatchType": "all", "conditions": conds}

def flt(field, match, value):
    return {"fieldName": field, "matchType": match, "value": value}

def dstr(d):
    return d.replace("-", "/")

def dept_filter(dept_ids):
    return [flt("DepartmentId", "equals", dept_ids)] if dept_ids else []

def wrap(fields, filters, name="Chat"):
    return {"cubeEntities": [{"name": name, "fields": fields, "filters": filters}],
            "mergeType": "column", "timezone": COMM100_TZ}

def p_volume_bot(ws, we, camp):
    fields = [
        f_count("chats", "Status", [cnd("Status0", "Status", "equals", ["0"]),
                                    cnd("Duration>0", "Duration", "notEquals", ["0"])]),
        f_count("missedChats", "Status", [cnd("missedChatStatus", "Status", "equals", ["2", "3"])]),
        f_count("refusedChats", "Status", [cnd("refusedChatStatus", "Status", "equals", ["1"])]),
        {"name": "chatRequests", "calculationType": "expression", "valueType": "int",
         "expression": "chats + missedChats + refusedChats"},
        f_count("botToAgentAccepted", "Status", [cnd("Status=0", "Status", "equals", ["0"]),
                                                 cnd("ChatType=2", "ChatType", "equals", ["2"])]),
        f_count("botToAgentRefused", "Status", [cnd("Status=1", "Status", "equals", ["1"]),
                                                cnd("ChatType=1", "ChatType", "equals", ["1"])]),
        f_count("botToAgentVisitorClose", "Status", [cnd("Status=2", "Status", "equals", ["2"]),
                                                     cnd("ChatType=1", "ChatType", "equals", ["1"])]),
        f_count("botToAgentLeaveMsg", "Status", [cnd("Status=2", "Status", "equals", ["2"]),
                                                 cnd("ChatType=4", "ChatType", "equals", ["4"])]),
        {"name": "chatFromBotToAgentRequest", "calculationType": "expression", "valueType": "int",
         "expression": "botToAgentAccepted+botToAgentRefused+botToAgentVisitorClose+botToAgentLeaveMsg"},
        f_count("botOnlyChats", "Id", [cnd("Status", "Status", "equals", ["0", "1", "2"]),
                                       cnd("ChatType", "ChatType", "equals", ["1"])]),
    ]
    return wrap(fields, [flt("RequestedTime", "between", [dstr(ws), dstr(we)]),
                         flt("CampaignId", "equals", list(camp))])

def p_rating(ws, we, camp=None, dept_ids=None, chat_types=("0", "2"), statuses=("0",)):
    fields = [{"name": "ratingTimes", "calculationType": "count", "valueType": "int",
               "fieldName": "PostChatSurvey.ChatId"}]
    for n in (5, 4, 3, 2, 1):
        fields.append(f_count(f"score{n}", "PostChatSurvey.RatingGrade",
                              [cnd(f"RatingGrade{n}", "PostChatSurvey.RatingGrade", "equals", [str(n)])]))
    fields.append({"name": "avgScore", "calculationType": "average", "valueType": "decimal",
                   "fieldName": "PostChatSurvey.RatingGrade"})
    filters = [flt("RequestedTime", "between", [dstr(ws), dstr(we)]),
               flt("PostChatSurvey.RatingGrade", "greaterThan", ["0"])]
    if chat_types:                           # None/boş = filtre yok (Comm100 özel raporu gibi)
        filters.append(flt("ChatType", "equals", list(chat_types)))
    if statuses:
        filters.append(flt("Status", "equals", list(statuses)))
    if camp:
        filters.append(flt("CampaignId", "equals", list(camp)))
    filters += dept_filter(dept_ids)
    return wrap(fields, filters)

def p_bot_rating(ws, we, bot_id):
    """Comm100 AI Agent > Rating raporunun birebir sorgusu (bot başına)."""
    fields = [{"name": "ratingTimes", "calculationType": "count", "valueType": "int",
               "fieldName": "PostChatSurvey.ChatId"}]
    for n in (5, 4, 3, 2, 1):
        fields.append(f_count(f"score{n}", "PostChatSurvey.RatingGrade",
                              [cnd(f"RatingGrade{n}", "PostChatSurvey.RatingGrade", "equals", [str(n)])]))
    fields.append({"name": "avgScore", "calculationType": "average", "valueType": "decimal",
                   "fieldName": "PostChatSurvey.RatingGrade"})
    return wrap(fields, [flt("RequestedTime", "between", [dstr(ws), dstr(we)]),
                         flt("BotId", "equals", [bot_id]),
                         flt("ChatType", "equals", ["1", "4"]),
                         flt("PostChatSurvey.RatingGrade", "greaterThan", ["0"])])

def p_action_usage(ws, we, bot_id):
    """AI Agent > Usage raporu: kullanılan AI yanıt (action) sayısı, bot başına."""
    fields = [{"name": "usedAIReplies", "calculationType": "sum", "valueType": "decimal", "fieldName": "UsedAIReplies"}]
    return {"cubeEntities": [{"name": "AIReplyRecord", "fields": fields,
                              "filters": [flt("CreatedTime", "between", [dstr(ws), dstr(we)]),
                                          flt("AIAgentId", "equals", [bot_id]),
                                          flt("UsedBy", "equals", ["0"])]}],
            "mergeType": "column", "timezone": COMM100_TZ}

# Bot süreleri: Chat History export'undaki W/X/Y kolonları (Duration / Agent Duration / Bot Duration) = Chat cube'unda
# Duration / AgentDuration / BotDuration. Elle yapılan hesap: o botun tüm chat'lerinin kolon ortalaması (agent'ı olmayanlar 0 sayılır).
# mod → (ChatType listesi (None = hepsi), Duration>0 filtresi)
BOT_DUR_MODES = {"12": (["1", "2"], False), "124": (["1", "2", "4"], False), "all": (None, False),
                 "12d": (["1", "2"], True), "124d": (["1", "2", "4"], True), "alld": (None, True)}
BOT_DUR_MODE = os.environ.get("BOT_DUR_MODE", "124")

def p_bot_dur(ws, we, bot_id, chat_types, dur_gt0):
    fields = [
        {"name": "avgAgentDuration", "calculationType": "average", "valueType": "timespan", "fieldName": "AgentDuration"},
        {"name": "avgBotDuration", "calculationType": "average", "valueType": "timespan", "fieldName": "BotDuration"},
        {"name": "avgDuration", "calculationType": "average", "valueType": "timespan", "fieldName": "Duration"},
        {"name": "n", "calculationType": "count", "valueType": "int", "fieldName": "Id"},
    ]
    filters = [flt("RequestedTime", "between", [dstr(ws), dstr(we)]), flt("BotId", "equals", [bot_id])]
    if chat_types:
        filters.append(flt("ChatType", "equals", list(chat_types)))
    if dur_gt0:
        filters.append(flt("Duration", "greaterThan", ["0"]))
    return wrap(fields, filters)

def fetch_bot_dur(ws, we, bot_id, mode=None):
    types, d0 = BOT_DUR_MODES.get(mode or BOT_DUR_MODE, BOT_DUR_MODES["124"])
    try:
        t = total(p_bot_dur(ws, we, bot_id, types, d0))
    except Exception as e:                   # BotDuration alanı kabul edilmezse tüm rapor patlamasın → hücreler sarı/boş kalır
        print(f"[UYARI] bot süresi alınamadı ({str(e)[:120]}) → I/J elle doldurulacak")
        return {"agent": 0.0, "bot": 0.0, "dur": 0.0, "n": 0}
    return {"agent": num(t.get("avgAgentDuration")), "bot": num(t.get("avgBotDuration")),
            "dur": num(t.get("avgDuration")), "n": int(num(t.get("n")))}

def bot_avg(r):
    """Puan sayılarından ağırlıklı ortalama (1–5)."""
    n = r["rated"]
    # Comm100 raporu ortalamayı 2 haneye yuvarlayıp gösteriyor; dosyadaki değerler de öyle (4,08/5 = 0,816)
    return round((5 * r["r5"] + 4 * r["r4"] + 3 * r["r3"] + 2 * r["r2"] + 1 * r["r1"]) / n, 2) if n else 0.0

def p_chat_dur(ws, we, camp, chat_types, statuses):
    """Belirli ChatType'taki chat'lerin ortalama toplam süresi ve agent süresi (bot süreleri için deneme)."""
    fields = [
        {"name": "avgDuration", "calculationType": "average", "valueType": "timespan", "fieldName": "Duration"},
        {"name": "avgAgentDuration", "calculationType": "average", "valueType": "timespan", "fieldName": "AgentDuration"},
        f_count("n", "Id", [cnd("Status", "Status", "equals", list(statuses)),
                            cnd("ChatType", "ChatType", "equals", list(chat_types))]),
    ]
    return wrap(fields, [flt("RequestedTime", "between", [dstr(ws), dstr(we)]), flt("CampaignId", "equals", list(camp)),
                         flt("Duration", "greaterThan", ["0"]),
                         flt("ChatType", "equals", list(chat_types)), flt("Status", "equals", list(statuses))])

def p_wait(ws, we, dept_ids):
    fields = [
        {"name": "avgWaitingTime", "calculationType": "average", "valueType": "timespan", "fieldName": "WaitingTime",
         "conditionExpression": "WaitingTime>0", "conditionMatchType": "all",
         "conditions": [cnd("WaitingTime>0", "WaitingTime", "notEquals", ["0"])]},
        {"name": "avgWaitingTimeOfMissedChats", "calculationType": "average", "valueType": "timespan",
         "fieldName": "WaitingTime", "conditionExpression": "Status2 & WaitingTime>0", "conditionMatchType": "all",
         "conditions": [cnd("Status2", "Status", "equals", ["2"]), cnd("WaitingTime>0", "WaitingTime", "notEquals", ["0"])]},
    ]
    return wrap(fields, [flt("RequestedTime", "between", [dstr(ws), dstr(we)])] + dept_filter(dept_ids))

def p_eff(ws, we, dept_ids):
    fields = [
        {"name": "avgAgentResponseTime", "calculationType": "average", "valueType": "timespan", "fieldName": "AvgResponseTime"},
        {"name": "avgChatTime", "calculationType": "average", "valueType": "timespan", "fieldName": "AgentDuration"},
    ]
    return wrap(fields, [flt("RequestedTime", "between", [dstr(ws), dstr(we)])] + dept_filter(dept_ids) + [
        flt("Duration", "greaterThan", ["0"]), flt("ChatType", "equals", ["0", "2"]), flt("Status", "equals", ["0"])])

def p_queue(ws, we, dept_ids):
    chat_fields = [
        f_count("chatsFromQueue", "Status", [cnd("Status0", "Status", "equals", ["0"])]),
        f_count("switchedToMsg", "Status", [cnd("Status2", "Status", "equals", ["2"]),
                                            cnd("ChatType4or8or16", "ChatType", "equals", ["4", "8", "16"])]),
        f_count("abandonedChats", "Status", [cnd("Status2", "Status", "equals", ["2"]),
                                             cnd("ChatTypeNot", "ChatType", "equals", ["0", "1", "2", "32"])]),
        f_count("refusedChats", "Status", [cnd("Status1", "Status", "equals", ["1"])]),
        {"name": "queuedChatRequests", "calculationType": "expression", "valueType": "int",
         "expression": "chatsFromQueue+switchedToMsg+abandonedChats+refusedChats"},
    ]
    return {"cubeEntities": [
        {"name": "Chat", "fields": chat_fields,
         "filters": [flt("RequestedTime", "between", [dstr(ws), dstr(we)])] + dept_filter(dept_ids)
                    + [flt("WaitingTime", "notEquals", ["0"])]},
        {"name": "QueueStatistics",
         "fields": [{"name": "maxQueueSize", "calculationType": "max", "valueType": "int", "fieldName": "QueueLength"}],
         "filters": [flt("Time", "between", [dstr(ws), dstr(we)])] + dept_filter(dept_ids)},
    ], "mergeType": "column", "timezone": COMM100_TZ}

def num(v):
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        return 0.0

def parse_rating(t):
    return {"rated": int(num(t.get("ratingTimes"))), "r5": int(num(t.get("score5"))),
            "r4": int(num(t.get("score4"))), "r3": int(num(t.get("score3"))),
            "r2": int(num(t.get("score2"))), "r1": int(num(t.get("score1"))),
            "avg": num(t.get("avgScore"))}

def fetch_dept_metrics(ws, we, dept_ids, eff_ids=None):
    """Bekleme + queue: dept_ids · yanıt + süre: eff_ids (verilmezse dept_ids)."""
    eff_ids = dept_ids if eff_ids is None else eff_ids
    w, e, q = total(p_wait(ws, we, dept_ids)), total(p_eff(ws, we, eff_ids)), total(p_queue(ws, we, dept_ids))
    return {"wait": num(w.get("avgWaitingTime")), "wait_missed": num(w.get("avgWaitingTimeOfMissedChats")),
            "resp": num(e.get("avgAgentResponseTime")), "dur": num(e.get("avgChatTime")),
            "q_req": int(num(q.get("queuedChatRequests"))), "q_srv": int(num(q.get("chatsFromQueue"))),
            "q_max": int(num(q.get("maxQueueSize")))}

def fetch_brands(ws, we):
    out = {}
    for code, _, camp in BRANDS:
        v = total(p_volume_bot(ws, we, camp))
        out[code] = {"total": int(num(v.get("chatRequests"))), "served": int(num(v.get("chats"))),
                     "missed": int(num(v.get("missedChats"))),
                     "bot_only": int(num(v.get("botOnlyChats"))),
                     "b2a_parts": {"accepted": int(num(v.get("botToAgentAccepted"))),
                                   "refused": int(num(v.get("botToAgentRefused"))),
                                   "visitor_close": int(num(v.get("botToAgentVisitorClose"))),
                                   "leave_msg": int(num(v.get("botToAgentLeaveMsg"))),
                                   "request": int(num(v.get("chatFromBotToAgentRequest")))},
                     "bot_to_agent": int(num(v.get("botToAgentAccepted" if BOT_TO_AGENT_MODE == "accepted"
                                                  else "chatFromBotToAgentRequest"))),
                     # Comm100 "rating scores per brand" raporu gibi: sadece campaign filtresi (ChatType/Status filtresiz)
                     "rating": parse_rating(total(p_rating(ws, we, camp, chat_types=None, statuses=None))),
                     "bot_rating": parse_rating(total(p_bot_rating(ws, we, BOT_IDS[code]))),
                     "bot_dur": fetch_bot_dur(ws, we, BOT_IDS[code]),
                     "action_usage": int(num(total(p_action_usage(ws, we, BOT_IDS[code])).get("usedAIReplies")))}
    return out

def fetch_stats(ws):
    we = week_end(ws)
    d = fetch_dept_metrics(ws, we, DEPTS.get(DEPT_MODE, DEPTS["cs"]), DEPTS.get(EFF_DEPT_MODE, DEPTS["cs"]))
    d["brands"] = fetch_brands(ws, we)
    # ALL CS: Comm100 Live Chat > Rating raporu gibi, sadece agent'ın yanıtladığı chat'ler (ChatType 0,2 + Status 0).
    # Marka satırları ise 'rating scores per brand' özel raporu gibi tüm puanlı chat'leri (agent + bot) sayar.
    d["rating_all"] = parse_rating(total(p_rating(ws, we)))
    return d


# ── GOOGLE SHEETS ────────────────────────────────────────────
def sheets_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    creds = service_account.Credentials.from_service_account_info(
        json.loads(GCP_CREDS_JSON), scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
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

def norm_tag(s):
    return str(s or "").strip().lower()

def read_agent_week(svc, ws):
    """Tag Sheet → {marka: {tag: chat}} (Haftalık, Departman = Tümü)."""
    if not TAGS_SPREADSHEET_ID:
        raise RuntimeError("MEMBER_TAGS_SPREADSHEET_ID tanımlı değil.")
    label = label_tags(ws)
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=TAGS_SPREADSHEET_ID, range="'Haftalık'!A2:E", valueRenderOption="UNFORMATTED_VALUE"))
    out = {}
    for r in resp.get("values", []):
        if len(r) < 5 or str(r[0]).strip() != label or str(r[2]).strip() != "Tümü":
            continue
        try:
            n = int(r[4])
        except (TypeError, ValueError):
            continue
        t = norm_tag(r[3])
        if t in SKIP_TAGS:
            continue
        d = out.setdefault(str(r[1]).strip(), {})
        d[t] = d.get(t, 0) + n
    return out

def _norm_date(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return (datetime(1899, 12, 30) + timedelta(days=int(v))).strftime("%Y-%m-%d")
    return str(v)[:10]

def read_goals_week(svc, ws):
    """Goals Sheet → {marka: {goal(tag adı, küçük harf): sayı}} o haftanın günleri toplamı."""
    if not GOALS_SPREADSHEET_ID:
        raise RuntimeError("GOALS_SPREADSHEET_ID tanımlı değil.")
    we = week_end(ws)
    resp = _exec(svc.spreadsheets().values().get(
        spreadsheetId=GOALS_SPREADSHEET_ID, range=f"'{GOALS_TAB}'!A2:D", valueRenderOption="UNFORMATTED_VALUE"))
    out = {}
    for r in resp.get("values", []):
        if len(r) < 4:
            continue
        d = _norm_date(r[0])
        if not (ws <= d <= we):
            continue
        try:
            n = int(r[3])
        except (TypeError, ValueError):
            continue
        g = norm_tag(r[2])
        g = GOAL_TO_TAG.get(g, g)
        b = out.setdefault(str(r[1]).strip(), {})
        b[g] = b.get(g, 0) + n
    return out


# ── TABLOLAR (top 10 / success) ──────────────────────────────
def roll_promo(counts):
    out, extra = {}, 0
    for k, n in counts.items():
        if k in PROMO_TAGS:
            extra += n
        else:
            out[k] = out.get(k, 0) + n
    if extra:
        out[PROMO_ROW] = out.get(PROMO_ROW, 0) + extra
    return out

def merge(dicts):
    out = {}
    for d in dicts:
        for k, n in d.items():
            out[k] = out.get(k, 0) + n
    return out

def disp(tag):
    return "Special Promotions" if tag == PROMO_ROW else tag

def top(counts, n=TOP_N):
    items = [(k, v) for k, v in counts.items() if v > 0]
    items.sort(key=lambda x: (-x[1], x[0]))
    return items[:n]

def build_tables(agent_by_brand, goals_by_brand):
    """agent_by_brand: {marka:{tag:n}}, goals_by_brand: {marka:{goal:n}} → şablona yazılacak tablolar."""
    t = {"agent": {}, "bot": [], "success": []}
    for code, sheet, _ in BRANDS:
        t["agent"][sheet] = top(agent_by_brand.get(code, {}))
    all_agent = merge(agent_by_brand.values())
    t["agent"][ALL_SHEET] = top(all_agent)
    bot_all = roll_promo(merge(goals_by_brand.values()))
    t["bot"] = top(bot_all)
    agent_rolled = roll_promo(all_agent)
    cand = []
    for tag, b in bot_all.items():
        if not all_agent:        # Tag Sheet'te o hafta yok (geçmiş 2 Kasım'dan başlıyor) → SUCCESS uydurma %100 olmasın
            break
        if b >= MIN_BOT_SUCCESS:
            a = agent_rolled.get(tag, 0)
            cand.append((tag, b, round((b - a) / b, 4)))
    cand.sort(key=lambda x: (-x[2], -x[1], x[0]))
    t["success"] = cand[:TOP_N]
    return t


# ── ŞABLONU DOLDUR ───────────────────────────────────────────
def to_time(sec):
    s = int(round(sec or 0))
    return dtime((s // 3600) % 24, (s % 3600) // 60, s % 60)

def sat(avg5):
    return round(avg5 / 5, 3) if avg5 else None

def mark_manual(ws_, cell):
    from openpyxl.styles import PatternFill
    ws_[cell].value = None
    ws_[cell].fill = PatternFill("solid", fgColor="FFFF00")

def fill_workbook(wb, weeks):
    """weeks = [(ws_prev, stats_prev, tables_prev), (ws_cur, stats_cur, tables_cur)]"""
    from openpyxl.styles import PatternFill
    (p_ws, p_st, p_tb), (c_ws, c_st, c_tb) = weeks
    pl, cl = label_short(p_ws), label_short(c_ws)
    sheets = {n: wb[n] for n in [ALL_SHEET] + [b[1] for b in BRANDS]}

    # etiketler
    for ws_ in sheets.values():
        ws_[f"A{ROWS['prev_stat']}"] = pl;  ws_[f"A{ROWS['cur_stat']}"] = cl
        ws_[f"A{ROWS['prev_bot']}"] = pl;   ws_[f"A{ROWS['cur_bot']}"] = cl
    # istatistik + chatbot blokları
    for (row_s, row_b, st) in ((ROWS["prev_stat"], ROWS["prev_bot"], p_st), (ROWS["cur_stat"], ROWS["cur_bot"], c_st)):
        a = sheets[ALL_SHEET]
        a[f"G{row_s}"] = to_time(st["wait"]);        a[f"H{row_s}"] = to_time(st["wait_missed"])
        a[f"I{row_s}"] = to_time(st["dur"]);         a[f"J{row_s}"] = to_time(st["resp"])
        a[f"K{row_s}"] = sat(st["rating_all"]["avg"])
        a[f"L{row_s}"] = st["q_req"];                a[f"N{row_s}"] = st["rating_all"]["rated"]
        a[f"P{row_s}"] = st["rating_all"]["r4"] + st["rating_all"]["r5"]     # Total Good Like = 4+5
        for col in MANUAL_CELLS["ALL CS"]:
            mark_manual(a, f"{col}{row_s}")
        for code, sheet, _ in BRANDS:
            b, s = st["brands"][code], sheets[sheet]
            s[f"B{row_s}"] = b["total"];   s[f"C{row_s}"] = b["served"]
            s[f"G{row_s}"] = sat(b["rating"]["avg"])
            s[f"B{row_b}"] = b["bot_only"]; s[f"C{row_b}"] = b["bot_to_agent"]
            s[f"K{row_b}"] = sat(bot_avg(b["bot_rating"]))
            s[f"F{row_b}"] = b["action_usage"]
            if b["bot_dur"]["n"]:
                s[f"I{row_b}"] = to_time(b["bot_dur"]["agent"]);  s[f"J{row_b}"] = to_time(b["bot_dur"]["bot"])
            else:
                mark_manual(s, f"I{row_b}");  mark_manual(s, f"J{row_b}")
        # ALL CS bot memnuniyeti = üç botun puanlarının ağırlıklı ortalaması
        agg = {k: sum(st["brands"][c]["bot_rating"][k] for c in st["brands"]) for k in ("rated", "r5", "r4", "r3", "r2", "r1")}
        a[f"K{row_b}"] = sat(bot_avg(agg))
        n_all = sum(st["brands"][c]["bot_dur"]["n"] for c in st["brands"])
        for col, key in (("I", "agent"), ("J", "bot")):       # chat sayısına göre ağırlıklı ortalama
            w = sum(st["brands"][c]["bot_dur"][key] * st["brands"][c]["bot_dur"]["n"] for c in st["brands"])
            if n_all:
                a[f"{col}{row_b}"] = to_time(w / n_all)
            else:
                mark_manual(a, f"{col}{row_b}")
        for name, cols in BOT_MANUAL.items():
            for col in cols:
                mark_manual(sheets[name], f"{col}{row_b}")

    # agent tag tabloları (A/B = geçen hafta, E/F = bu hafta)
    for name, (hdr, first) in TAG_TABLE.items():
        s = sheets[name]
        s[f"A{hdr}"] = pl;  s[f"E{hdr}"] = cl
        for off in range(TOP_N):
            for (cn, cc, tb) in (("A", "B", p_tb), ("E", "F", c_tb)):
                rows = tb["agent"].get(name, [])
                tag, n = (rows[off] if off < len(rows) else (None, None))
                s[f"{cn}{first + off}"] = tag
                s[f"{cc}{first + off}"] = n
    # BOT TAGS + SUCCESS (yalnızca ALL CS)
    a = sheets[ALL_SHEET]
    hdr, first = BOT_TABLE
    a[f"A{hdr}"] = f"BOT TAGS {pl}";  a[f"E{hdr}"] = f"BOT TAGS {cl}"
    for off in range(TOP_N):
        for (cn, cc, tb) in (("A", "B", p_tb), ("E", "F", c_tb)):
            tag, n = tb["bot"][off] if off < len(tb["bot"]) else (None, None)
            a[f"{cn}{first + off}"] = disp(tag) if tag else None
            a[f"{cc}{first + off}"] = n
    hdr, first = SUCC_TABLE
    a[f"A{hdr}"] = f"BOT TAGS SUCCESS {pl}";  a[f"E{hdr}"] = f"BOT TAGS SUCCESS {cl}"
    for off in range(TOP_N):
        for (cn, cc, cp, tb) in (("A", "B", "C", p_tb), ("E", "F", "G", c_tb)):
            row = tb["success"][off] if off < len(tb["success"]) else None
            a[f"{cn}{first + off}"] = disp(row[0]) if row else None
            a[f"{cc}{first + off}"] = row[1] if row else None
            a[f"{cp}{first + off}"] = row[2] if row else None


# ── DEĞİŞİM (%) RENKLERİ ─────────────────────────────────────
# Delta hücreleri (istatistik satırı 9, chatbot satırı 15) metriğin yönüne göre boyanır:
#   HIGHER = artış iyi (yeşil), LOWER = artış kötü (kırmızı); listede olmayanlar nötr (siyah).
# Renk, hücre değerine bağlı koşullu biçimlendirmedir → elle düzeltme yapılsa da doğru kalır.
DELTA_RULES = {
    "ALL CS":     {"stat": {"higher": "DKNOPQ", "lower": "EFGHIJLM"}, "bot": {"higher": "KL", "lower": "EFGHIJ"}},
    "SUPERBETIN": {"stat": {"higher": "DG",     "lower": "EF"},       "bot": {"higher": "KL", "lower": "EFGHIJ"}},
    "BETSAT":     {"stat": {"higher": "DG",     "lower": "EF"},       "bot": {"higher": "KL", "lower": "EFGHIJ"}},
    "TURKBET":    {"stat": {"higher": "DG",     "lower": "EF"},       "bot": {"higher": "KL", "lower": "EFGHIJ"}},
}
GREEN, RED = "FF00B050", "FFFF0000"

def apply_delta_colors(wb):
    import copy
    from openpyxl.styles import Font
    from openpyxl.styles.colors import Color
    from openpyxl.formatting.rule import CellIsRule
    for name, blocks in DELTA_RULES.items():
        if name not in wb.sheetnames:
            continue
        ws_ = wb[name]
        for block, delta_row in (("stat", ROWS["cur_stat"] + 1), ("bot", ROWS["cur_bot"] + 1)):
            for direction, cols in blocks[block].items():
                up, down = (GREEN, RED) if direction == "higher" else (RED, GREEN)
                for col in cols:
                    cell = ws_[f"{col}{delta_row}"]
                    f = copy.copy(cell.font)
                    cell.font = Font(name=f.name, size=f.size, bold=f.bold, italic=f.italic, color=Color(rgb="FF000000"))
                    ref = f"{col}{delta_row}"
                    ws_.conditional_formatting.add(ref, CellIsRule(operator="greaterThan", formula=["0"], font=Font(bold=True, color=up)))
                    ws_.conditional_formatting.add(ref, CellIsRule(operator="lessThan", formula=["0"], font=Font(bold=True, color=down)))
            # kuralı olmayan (nötr) delta hücreleri de siyaha çek: şablondaki eski sabit renk kalmasın
            ruled = set("".join(blocks[block].values()))
            for cell in ws_[delta_row]:
                v = cell.value
                if isinstance(v, str) and v.startswith("=(") and cell.column_letter not in ruled:
                    f = copy.copy(cell.font)
                    cell.font = Font(name=f.name, size=f.size, bold=f.bold, italic=f.italic, color=Color(rgb="FF000000"))


# ── MAİL / DRIVE ─────────────────────────────────────────────
def fmt_n(n):
    return f"{int(round(n)):,}".replace(",", ".")

def fmt_t(sec):
    s = int(round(sec or 0))
    return f"{s // 60}:{s % 60:02d}"

def summary(st):
    tot = sum(b["total"] for b in st["brands"].values())
    srv = sum(b["served"] for b in st["brands"].values())
    bot = sum(b["bot_only"] for b in st["brands"].values())
    b2a = sum(b["bot_to_agent"] for b in st["brands"].values())
    return {"Total chats": tot, "Served": srv, "Missed": tot - srv,
            "Acceptance %": (srv / tot * 100) if tot else 0, "Chats (Only Bot)": bot,
            "Chats (Bot to Agent)": b2a, "% Bot chat volume": (bot / srv * 100) if srv else 0,
            "Avg WT on Queue": st["wait"], "Avg response": st["resp"], "Avg duration": st["dur"],
            "Queued requests": st["q_req"], "Rated chats": st["rating_all"]["rated"],
            "Good (4+5)": st["rating_all"]["r4"] + st["rating_all"]["r5"]}

LOWER_BETTER = {"Missed", "Avg WT on Queue", "Avg response", "Avg duration", "Queued requests"}   # artış = kötü (kırmızı)
NEUTRAL_METRICS = {"Total chats", "Served", "Chats (Only Bot)", "Chats (Bot to Agent)"}      # yön yorumlanmaz → siyah (Excel ile aynı)

# Şablondaki marka logoları (xl/media) → mailde inline (cid) gösterilir; bulunamazsa logosuz devam eder
LOGO_FILES = {"sb_word": "image4.png", "bs_word": "image5.png", "tb_word": "image6.png",
              "sb": "image3.png", "bs": "image2.png", "tb": "image1.png"}

def load_logos():
    out = {}
    try:
        with zipfile.ZipFile(TEMPLATE) as z:
            names = set(z.namelist())
            for key, fn in LOGO_FILES.items():
                if f"xl/media/{fn}" in names:
                    out[key] = z.read(f"xl/media/{fn}")
    except Exception as e:
        print(f"[MAIL] logolar okunamadı: {e}")
    return out

def img_tag(logos, key, h):
    if key not in logos:
        return ""
    return f'<img src="cid:logo_{key}" alt="" style="height:{h}px;width:auto;vertical-align:middle;border:0">'

def brand_rows(p_st, c_st, logos):
    names = {"SB": "Superbetin", "BS": "Betsat", "TB": "Turkbet"}
    metrics = [("Total chats", lambda b: b["total"], fmt_n), ("Served", lambda b: b["served"], fmt_n),
               ("Only Bot", lambda b: b["bot_only"], fmt_n), ("Bot→Agent", lambda b: b["bot_to_agent"], fmt_n),
               ("Action Usage", lambda b: b["action_usage"], fmt_n),
               ("Bot satisfaction", lambda b: sat(bot_avg(b["bot_rating"])) or 0, lambda v: f"{v * 100:.2f}%")]
    BRAND_DIR = {"Action Usage": "lower", "Bot satisfaction": "higher"}      # diğerleri nötr (siyah), Excel ile aynı
    head = "".join(f"<th style='padding:8px 10px;text-align:right'>{m[0]}</th>" for m in metrics)
    rows = ""
    for code, _, _ in BRANDS:
        cb, pb = c_st["brands"][code], p_st["brands"][code]
        cells = ""
        for mname, get, fmt in metrics:
            cv, pv = get(cb), get(pb)
            ch = ((cv - pv) / pv * 100) if pv else 0
            d = BRAND_DIR.get(mname)
            if d is None:
                color = "#111827"
            elif abs(ch) < 0.05:
                color = "#6b7280"
            else:
                color = "#16a34a" if ((ch > 0) == (d == "higher")) else "#dc2626"
            cells += (f"<td style='padding:6px 10px;border-bottom:1px solid #eee;text-align:right'>"
                      f"<b>{fmt(cv)}</b><br><span style='font-size:11px;color:{color}'>{ch:+.1f}%</span></td>")
        rows += (f"<tr><td style='padding:6px 10px;border-bottom:1px solid #eee;white-space:nowrap'>"
                 f"{img_tag(logos, code.lower(), 22)} {names[code]}</td>{cells}</tr>")
    return (f"<table style='border-collapse:collapse;font-size:13px;margin-top:20px'>"
            f"<tr style='background:#111827;color:#fff'><th style='padding:8px 10px;text-align:left'>Brand</th>{head}</tr>"
            f"{rows}</table>")

def build_mail_html(p_ws, p_st, c_ws, c_st, logos=None):
    logos = logos or {}
    ps, cs = summary(p_st), summary(c_st)
    rows = ""
    for k in cs:
        cv, pv = cs[k], ps[k]
        if k in ("Avg WT on Queue", "Avg response", "Avg duration"):
            a, b = fmt_t(cv), fmt_t(pv)
        elif "%" in k:
            a, b = f"{cv:.2f}%", f"{pv:.2f}%"
        else:
            a, b = fmt_n(cv), fmt_n(pv)
        ch = ((cv - pv) / pv * 100) if pv else 0
        if k in NEUTRAL_METRICS:
            color = "#111827"                # nötr metrik: siyah
        elif abs(ch) < 0.05:
            color = "#6b7280"                # değişim yok: gri
        else:
            good = (ch < 0) if k in LOWER_BETTER else (ch > 0)
            color = "#16a34a" if good else "#dc2626"
        rows += (f"<tr><td style='padding:6px 12px;border-bottom:1px solid #eee'>{k}</td>"
                 f"<td style='padding:6px 12px;border-bottom:1px solid #eee;text-align:right;color:#6b7280'>{b}</td>"
                 f"<td style='padding:6px 12px;border-bottom:1px solid #eee;text-align:right;font-weight:700'>{a}</td>"
                 f"<td style='padding:6px 12px;border-bottom:1px solid #eee;text-align:right;color:{color}'>{ch:+.1f}%</td></tr>")
    head_logos = " &nbsp;&nbsp; ".join(x for x in (img_tag(logos, "sb_word", 30), img_tag(logos, "bs_word", 30),
                                                   img_tag(logos, "tb_word", 30)) if x)
    head_logos = f'<div style="margin:0 0 14px">{head_logos}</div>' if head_logos else ""
    return f"""<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;color:#111827;">
{head_logos}<h2 style="margin:0 0 4px">CS Weekly Comparison — {label_short(c_ws)} {week_end(c_ws)[:4]}</h2>
<p style="margin:0 0 14px;color:#6b7280">Ekteki dosya hazır. <b>Sarı hücreler</b> API'de olmayan alanlardır
(shifts, call attempts/reached) — elle doldurulacak.</p>
<table style="border-collapse:collapse;font-size:13px;min-width:520px">
<tr style="background:#111827;color:#fff"><th style="padding:8px 12px;text-align:left">Metric</th>
<th style="padding:8px 12px;text-align:right">{label_short(p_ws)}</th>
<th style="padding:8px 12px;text-align:right">{label_short(c_ws)}</th>
<th style="padding:8px 12px;text-align:right">Change</th></tr>{rows}</table>
{brand_rows(p_st, c_st, logos)}
<p style="color:#9ca3af;font-size:11px;margin-top:18px">CS Operations &amp; Technology · otomatik üretildi</p>
</body></html>"""

def recipients():
    if MAIL_MODE == "none" or not REPORT_EMAILS:
        return []
    if MAIL_MODE == "me":
        return [GMAIL_USER] if GMAIL_USER else REPORT_EMAILS[:1]
    return REPORT_EMAILS

def send_mail(path, subject, html, logos=None):
    to = recipients()
    if not to:
        print("[MAIL] alıcı yok / MAIL_MODE=none → mail atılmadı")
        return
    if not (GMAIL_USER and GMAIL_PASS):
        print("[MAIL] GMAIL_USER / GMAIL_APP_PASSWORD yok → mail atılmadı")
        return
    msg = MIMEMultipart("mixed")
    msg["Subject"], msg["From"], msg["To"] = subject, GMAIL_USER, to[0]
    if len(to) > 1:
        msg["Cc"] = ", ".join(to[1:])
    rel = MIMEMultipart("related")
    rel.attach(MIMEText(html, "html", "utf-8"))
    for key, data in (logos or {}).items():
        im = MIMEImage(data, _subtype="png")
        im.add_header("Content-ID", f"<logo_{key}>")
        im.add_header("Content-Disposition", "inline", filename=f"{key}.png")
        rel.attach(im)
    msg.attach(rel)
    with open(path, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=os.path.basename(path))
    msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(GMAIL_USER, GMAIL_PASS)
        s.sendmail(GMAIL_USER, to, msg.as_string())
    print(f"[MAIL] gönderildi → {', '.join(to)}")

def upload_drive(path):
    """Opsiyonel: DRIVE_FOLDER_ID tanımlıysa dosyayı oraya yükler (Paylaşımlı Drive klasörü önerilir). Hata işi düşürmez."""
    if not DRIVE_FOLDER_ID:
        return
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        creds = service_account.Credentials.from_service_account_info(
            json.loads(GCP_CREDS_JSON), scopes=["https://www.googleapis.com/auth/drive"])
        drv = build("drive", "v3", credentials=creds, cache_discovery=False)
        media = MediaFileUpload(path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        f = drv.files().create(body={"name": os.path.basename(path), "parents": [DRIVE_FOLDER_ID]},
                               media_body=media, fields="id,webViewLink", supportsAllDrives=True).execute()
        print(f"[DRIVE] yüklendi: {f.get('webViewLink', f.get('id'))}")
    except Exception as e:
        print(f"[DRIVE] yüklenemedi (mail etkilenmedi): {e}")

def send_error_email(label, msg_, tb):
    if not (GMAIL_USER and GMAIL_PASS and REPORT_EMAILS):
        return
    m = MIMEText(f"<h3 style='color:#c0392b'>CS Weekly Comparison — Hata</h3><p><b>{label}</b></p>"
                 f"<p>{msg_}</p><pre style='font-size:12px;white-space:pre-wrap'>{tb}</pre>", "html", "utf-8")
    m["Subject"], m["From"], m["To"] = f"⚠️ CS Weekly Comparison ERROR | {label}", GMAIL_USER, REPORT_EMAILS[0]
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.sendmail(GMAIL_USER, [REPORT_EMAILS[0]], m.as_string())
    except Exception as e:
        print(f"[EMAIL] hata maili gönderilemedi: {e}")


# ── ANA AKIŞ ─────────────────────────────────────────────────
def build_week_inputs(svc, ws):
    stats = fetch_stats(ws)
    agents = read_agent_week(svc, ws)
    if not agents:
        print(f"[UYARI] Tag Sheet'te {label_tags(ws)} haftası yok → agent tag ve SUCCESS tabloları boş kalır")
    tables = build_tables(agents, read_goals_week(svc, ws))
    return stats, tables

def generate(ws):
    import openpyxl
    p_ws = add_days(ws, -7)
    svc = sheets_service()
    print(f"[HAFTA] {label_short(ws)} {week_end(ws)[:4]} (önceki: {label_short(p_ws)}) · dept={DEPT_MODE}")
    c_st, c_tb = build_week_inputs(svc, ws)
    p_st, p_tb = build_week_inputs(svc, p_ws)
    if not os.path.exists(TEMPLATE):
        raise RuntimeError(f"Şablon bulunamadı: {TEMPLATE} (repo'ya templates/ altına koy)")
    wb = openpyxl.load_workbook(TEMPLATE)
    fill_workbook(wb, [(p_ws, p_st, p_tb), (ws, c_st, c_tb)])
    apply_delta_colors(wb)
    os.makedirs(OUT_DIR, exist_ok=True)
    path = os.path.join(OUT_DIR, file_name(ws))
    wb.save(path)
    print(f"[DOSYA] {path}")
    try:                                    # Pillow yoksa openpyxl logoları sessizce düşürür → uyar
        with zipfile.ZipFile(TEMPLATE) as zt, zipfile.ZipFile(path) as zo:
            n_t = sum(1 for n in zt.namelist() if n.startswith("xl/media/"))
            n_o = sum(1 for n in zo.namelist() if n.startswith("xl/media/"))
        if n_t and n_o < n_t:
            print(f"[UYARI] şablondaki {n_t} logodan {n_o} tanesi dosyaya yazıldı → 'pip install pillow' gerekli")
    except Exception as e:
        print(f"[UYARI] logo kontrolü yapılamadı: {e}")
    return path, (p_ws, p_st), (ws, c_st)

def calibrate(date_str):
    """API değerlerini (3 departman varyantı) şablondaki 2. satırla (bu hafta) karşılaştırır."""
    import openpyxl
    ws = week_start(date_str)
    we = week_end(ws)
    wb = openpyxl.load_workbook(TEMPLATE, data_only=False)
    a = wb[ALL_SHEET]
    print(f"[CALIBRATE] hafta {ws}..{we} — şablonun 8. satırı bu haftaya ait varsayılır\n")
    brands = fetch_brands(ws, we)
    for code, sheet, _ in BRANDS:
        b, s = brands[code], wb[sheet]
        pr = b["b2a_parts"]
        rc = parse_rating(total(p_rating(ws, we, dict((c, k) for c, _, k in BRANDS)[code], DEPTS["cs"])))
        print(f"{code}: total API {b['total']} / dosya {s['B8'].value} | served {b['served']} / {s['C8'].value} | "
              f"botOnly {b['bot_only']} / {s['B14'].value}")
        print(f"    bot→agent dosya {s['C14'].value} | API parçalar: accepted {pr['accepted']} + refused {pr['refused']} "
              f"+ visitor_close {pr['visitor_close']} + leave_msg {pr['leave_msg']} = {pr['request']}")
        print(f"    sat dosya {s['G8'].value} | API tüm dept {sat(b['rating']['avg'])} | API sadece CS {sat(rc['avg'])} "
              f"| rated {b['rating']['rated']}/CS {rc['rated']}")
    print("\nMEMNUNİYET varyantları (marka, campaign filtreli) — Comm100 'rating scores per brand' raporuyla kıyasla:")
    variants = {"mevcut (ChatType 0,2 + Status 0)": dict(),
                "filtresiz (rapor gibi)": dict(chat_types=None, statuses=None),
                "ChatType 0,2 / status filtresiz": dict(statuses=None),
                "ChatType filtresiz / status 0": dict(chat_types=None)}
    for code, sheet, camp in BRANDS:
        print(f"{code}: dosya G7={wb[sheet]['G7'].value} G8={wb[sheet]['G8'].value}")
        for name, kw in variants.items():
            r = parse_rating(total(p_rating(ws, we, camp, **kw)))
            print(f"    {name:36}: rated {r['rated']:4} | avg {r['avg']:.3f} | avg/5 {sat(r['avg'])} | 4+5 {r['r4'] + r['r5']}")
    print("\nBOT SÜRELERİ + BOT-ONLY MEMNUNİYET adayları (dosyadaki 13/14. satırlarla karşılaştır):")
    for code, sheet, camp in BRANDS:
        s = wb[sheet]
        rb = brands[code]["bot_rating"]
        print(f"{code}: dosya agent süresi I14={s['I14'].value} bot süresi J14={s['J14'].value} bot memnuniyeti K14={s['K14'].value}")
        for mode in BOT_DUR_MODES:
            try:
                d = fetch_bot_dur(ws, we, BOT_IDS[code], mode)
                print(f"    mod {mode:5}: agent {to_time(d['agent'])} | bot {to_time(d['bot'])} | toplam {to_time(d['dur'])} | chat {d['n']}")
            except Exception as e:
                print(f"    mod {mode:5}: HATA {str(e)[:150]}")
        print(f"    Action Usage: dosya F14={s['F14'].value} | API {brands[code]['action_usage']}")
        print(f"    API (AI Agent Rating raporu): bot memnuniyeti avg/5 {sat(bot_avg(rb))} (rated {rb['rated']}, avgScore {rb['avg']})")
    ra = parse_rating(total(p_rating(ws, we)))
    print(f"\nRating(all): rated {ra['rated']} / {a['N8'].value} | good 4+5 {ra['r4'] + ra['r5']} / {a['P8'].value} | "
          f"avg/5 {sat(ra['avg'])} / K8 {a['K8'].value}")
    good, bad = ra["r4"] + ra["r5"], ra["r1"] + ra["r2"]
    print(f"  satisfaction adayları: 4+5/rated={good / ra['rated']:.3f} | 4+5/(4+5+1+2)={good / max(good + bad, 1):.3f}")
    print("\nDepartman varyantları (API) → dosya  [wait/queue ve response/duration ayrı seçiliyor]:")
    print(f"  dosya: wait {a['G8'].value} | wait missed {a['H8'].value} | duration {a['I8'].value} | "
          f"response {a['J8'].value} | queued {a['L8'].value}")
    for mode, ids in DEPTS.items():
        d = fetch_dept_metrics(ws, we, ids)
        print(f"  {mode:7}: wait {to_time(d['wait'])} | wait missed {to_time(d['wait_missed'])} | "
              f"duration {to_time(d['dur'])} | response {to_time(d['resp'])} | queued {d['q_req']}")

def main():
    args = sys.argv[1:]
    if args and args[0] == "calibrate":
        calibrate(args[1] if len(args) > 1 else last_full_week())
        return
    ws = week_start(args[0]) if args else last_full_week()
    label = f"{label_short(ws)} {week_end(ws)[:4]}"
    try:
        path, (p_ws, p_st), (c_ws, c_st) = generate(ws)
        logos = load_logos()
        send_mail(path, f"CS Weekly Comparison | {label}", build_mail_html(p_ws, p_st, c_ws, c_st, logos), logos)
        upload_drive(path)
    except Exception as e:
        tb = traceback.format_exc()
        print(tb)
        send_error_email(label, str(e), tb)
        sys.exit(1)

if __name__ == "__main__":
    main()
