"""
rating_dashboard.py — kullanıcı kırılımı testleri.

Çalıştırma:  python -m unittest discover -s tests -v
Google kütüphaneleri kurulu değilse stub'lanır; ağ/API çağrısı yapılmaz.
"""
import contextlib
import importlib.util
import io
import json
import os
import sys
import types
import unittest

ROOT   = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(ROOT, "scripts", "rating_dashboard.py")

for k in ("COMM100_SITE_ID", "COMM100_API_KEY", "COMM100_EMAIL",
          "RATING_SPREADSHEET_ID", "GCP_CREDENTIALS", "GMAIL_USER",
          "GMAIL_APP_PASSWORD"):
    os.environ.setdefault(k, "test")
os.environ.setdefault("REPORT_EMAILS", "a@example.com")


def _stub(name, **attrs):
    mod = types.ModuleType(name)
    for k, v in attrs.items():
        setattr(mod, k, v)
    sys.modules[name] = mod
    return mod


try:
    import google.oauth2.service_account  # noqa: F401
    import googleapiclient.discovery       # noqa: F401
except ImportError:
    _stub("google")
    _stub("google.oauth2")
    _stub("google.oauth2.service_account", Credentials=object)
    _stub("googleapiclient")
    _stub("googleapiclient.discovery", build=lambda *a, **k: None)
    sys.modules["google"].oauth2 = sys.modules["google.oauth2"]
    sys.modules["google.oauth2"].service_account = sys.modules["google.oauth2.service_account"]
    sys.modules["googleapiclient"].discovery = sys.modules["googleapiclient.discovery"]

_spec = importlib.util.spec_from_file_location("rating_dashboard", SCRIPT)
rd = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(rd)


# ── Sahte Sheets servisi ─────────────────────────────────────
class _Exec:
    def __init__(self, result):
        self._result = result

    def execute(self):
        return self._result


class FakeSheets:
    def __init__(self):
        self.values_updates = []   # (range, values)
        self.batch_calls    = []   # her batchUpdate'in request listesi
        self._sheets        = {}
        self._next_id       = 100

    # svc.spreadsheets()
    def spreadsheets(self):
        return self

    def values(self):
        return self

    def get(self, spreadsheetId):
        return _Exec({"sheets": [{"properties": {"title": t, "sheetId": i}}
                                 for t, i in self._sheets.items()]})

    def update(self, spreadsheetId, range, valueInputOption, body):
        self.values_updates.append((range, body["values"]))
        return _Exec({})

    def batchUpdate(self, spreadsheetId, body):
        reqs = body["requests"]
        replies = []
        for r in reqs:
            if "addSheet" in r:
                title = r["addSheet"]["properties"]["title"]
                self._sheets[title] = self._next_id
                replies.append({"addSheet": {"properties": {"sheetId": self._next_id}}})
                self._next_id += 1
            elif "deleteSheet" in r:
                sid = r["deleteSheet"]["sheetId"]
                self._sheets = {t: i for t, i in self._sheets.items() if i != sid}
            else:
                replies.append({})
        self.batch_calls.append(reqs)
        return _Exec({"replies": replies})

    def format_requests(self):
        return [r for call in self.batch_calls for r in call
                if "addSheet" not in r and "deleteSheet" not in r]


# ── Sahte chat verisi ────────────────────────────────────────
def chat(cid, grade=None, name=None, agent="Agent A", tag="(Payment_Delay)", name_field="preChatName",
         code=None, visitor=None, brand=None):
    c = {"id": str(cid), "chatAgents": [{"agent": {"displayName": agent}}],
         "chatWrapup": {"categoriesName": tag}}
    if name is not None:
        c[name_field] = name
    if code is not None:
        c["_userCode"] = code
    if visitor is not None:
        c["visitorId"] = visitor
    if brand is not None:
        c["campaignId"] = f"camp-{brand.lower()}"
        c["campaign"] = {"id": f"camp-{brand.lower()}", "name": brand}
    if grade is not None:
        c["postChatSurvey"] = {"ratingGrade": grade}
    return c


def make_day(n_users=160, seed_offset=0):
    """
    - 'Alice' 3 low chat (2x R1, 1x R2) farklı yazımlarla + 2 normal chat;
      tag'ler: Payment_Delay x2 (R1 + R2), tag'siz x1
    - 'bob' 1x R1, 'Carl' 1x R2 (name alanından), boş isimli low chat
    - n_users adet tek seferlik R2 kullanıcısı
    """
    low, allc = [], []
    cid = seed_offset * 100000
    def add(grade, name, **kw):
        nonlocal cid
        cid += 1
        c = chat(cid, grade, name, **kw)
        allc.append(c)
        if grade in (1, 2):
            low.append(c)
    add(1, "Alice")                        # Payment_Delay
    add(1, "ALICE ", tag="")               # tag'siz → "—"
    add(2, "alice")                        # Payment_Delay (aynı tag ikinci kez)
    add(5, "Alice", tag="(Other_Topic)")   # yüksek puan — tag'lere girmemeli
    add(None, "aLiCe")
    add(1, "bob")
    add(2, "Carl", name_field="name")
    add(1, "", name_field="preChatName")
    add(2, None)
    add(4, "Dora")  # sadece yüksek puan — listede olmamalı
    for i in range(n_users):
        add(2, f"user{i:04d}")
        add(5, f"USER{i:04d}")
    return low, allc


def rows_by_range(fake):
    assert len(fake.values_updates) == 1
    return fake.values_updates[0][1]


def req_rows(req):
    """Bir format isteğinin kapsadığı satır aralığı (start, end_exclusive)."""
    body = next(iter(req.values()))
    r = body.get("range")
    if r is None or "startRowIndex" not in r:
        return None
    return r["startRowIndex"], r["endRowIndex"]


class BuildPivotsTest(unittest.TestCase):
    def setUp(self):
        self.low, self.allc = make_day()
        (self.tag_rows, self.agent_rows, self.user_rows,
         self.total_low, self.total_all) = rd.build_pivots(self.low, self.allc)

    def test_return_shape(self):
        self.assertEqual(self.total_low, len(self.low))
        self.assertEqual(self.total_all, len(self.allc))

    def test_case_insensitive_grouping_first_spelling(self):
        alice = [r for r in self.user_rows if r[0].lower() == "alice"]
        self.assertEqual(len(alice), 1)
        # ilk görülen yazım; Total Chats tüm yazımları kapsar (5 chat)
        self.assertEqual(alice[0][:10], ["Alice", "", "", "", 2, 1, 3, 5, "60.0%",
                                        "Payment_Delay (2), —"])
        self.assertEqual(alice[0][10], {"Payment_Delay": 2, "—": 1})

    def test_empty_names_skipped_and_high_only_excluded(self):
        names = [r[0] for r in self.user_rows]
        self.assertNotIn("", names)
        self.assertNotIn("Dora", names)
        # 3 isimli düşük puan kullanıcısı + 160 tekil kullanıcı, üst sınır yok
        self.assertEqual(len(self.user_rows), 3 + 160)

    def test_name_field_fallback(self):
        carl = next(r for r in self.user_rows if r[0] == "Carl")
        self.assertEqual(carl[:10], ["Carl", "", "", "", 0, 1, 1, 1, "100.0%", "Payment_Delay"])

    def test_sorting(self):
        self.assertEqual(self.user_rows[0][0], "Alice")
        self.assertEqual(self.user_rows[1][0], "bob")   # low 1, r1 1
        self.assertEqual(self.user_rows[2][0], "Carl")  # low 1, r1 0, alfabetik önce
        keys = [(-r[6], -r[4], r[0].lower()) for r in self.user_rows]
        self.assertEqual(keys, sorted(keys))

    def test_user_totals_consistent(self):
        named_low = sum(1 for c in self.low if rd.get_username(c))
        self.assertEqual(sum(r[6] for r in self.user_rows), named_low)

    def test_tag_format(self):
        self.assertEqual(rd.format_tag_counts({"b_x": 2, "a_x": 2, "c_x": 1, "—": 3}),
                         "— (3), a_x (2), b_x (2), c_x")
        self.assertEqual(rd.format_tag_counts({}), "")

    def test_no_low_ratings(self):
        _, _, ur, tl, ta = rd.build_pivots([], [chat(1, 5, "x")])
        self.assertEqual((ur, tl, ta), ([], 0, 1))


class DailySheetTest(unittest.TestCase):
    def setUp(self):
        low, allc = make_day()
        self.tag_rows, self.agent_rows, self.user_rows, tl, ta = rd.build_pivots(low, allc)
        self.fake = FakeSheets()
        rd.write_daily_sheet(self.fake, "2026-10-14", self.tag_rows, self.agent_rows,
                             self.user_rows, tl, ta)
        self.rows = rows_by_range(self.fake)

    def test_layout(self):
        rows = self.rows
        agent_total = 3 + len(self.tag_rows) + 1 + 1 + 1 + 1 + len(self.agent_rows)
        self.assertEqual(rows[agent_total - len(self.agent_rows) - 2][0], "👤 BREAKDOWN BY AGENT")
        self.assertEqual(rows[agent_total][0], "TOTAL")
        self.assertEqual(rows[agent_total + 1], [""] * 6)
        self.assertEqual(rows[agent_total + 2][0], "👤 BREAKDOWN BY USER")
        self.assertEqual(rows[agent_total + 3],
                         ["Chat Name", "Username", "User Code", "Brand", "Rating 1 ❌",
                          "Rating 2 🟠", "Low Total", "Total Chats", "% Rate", "Low Rating Tags"])
        data = rows[agent_total + 4: agent_total + 4 + len(self.user_rows)]
        self.assertEqual(data, [r[:10] for r in self.user_rows])
        self.assertEqual(data[0][9], "Payment_Delay (2), —")
        total = rows[agent_total + 4 + len(self.user_rows)]
        self.assertEqual(total[:4], ["TOTAL", "", "", ""])
        self.assertEqual(total[9], "")
        self.assertEqual(total[6], sum(r[6] for r in self.user_rows))
        self.assertEqual(len(rows), agent_total + 5 + len(self.user_rows))
        # mevcut bölümler 6 sütun, kullanıcı bölümü (başlık hariç) 10 sütun
        self.assertTrue(all(len(r) == 6 for r in rows[:agent_total + 3]))
        self.assertTrue(all(len(r) == 10 for r in rows[agent_total + 3:]))
        # başlık satırı değişmedi
        self.assertTrue(rows[0][0].startswith("📊 Rating 1 & 2 Dashboard"))

    def test_format_rows_and_batches(self):
        reqs = self.fake.format_requests()
        user_header = next(i for i, r in enumerate(self.rows) if r[0] == "👤 BREAKDOWN BY USER")
        upd = [r for r in reqs if "updateCells" in r]
        self.assertEqual(len(upd), 1)
        rr = upd[0]["updateCells"]["range"]
        self.assertEqual(rr["startRowIndex"], user_header + 2)
        self.assertEqual(rr["endRowIndex"], user_header + 2 + len(self.user_rows))
        self.assertEqual(len(upd[0]["updateCells"]["rows"]), len(self.user_rows))
        # Önceki bölümlere ait istekler kullanıcı bölümüne taşmıyor
        user_reqs = [r for r in reqs if req_rows(r) and req_rows(r)[0] >= user_header]
        for r in reqs:
            if r in user_reqs:
                continue
            span = req_rows(r)
            if span:
                self.assertLessEqual(span[1], user_header - 1, r)
        # Kullanıcı istekleri satır başına değil: başlık, kolon, veri, TOTAL,
        # yatay hizalama (C:I ortalı, A:B sola, J sola) ve dikey ortalama = 8
        self.assertEqual(len(user_reqs), 8)
        # batch boyutu
        for call in self.fake.batch_calls:
            self.assertLessEqual(len(call), rd.BATCH_MAX_REQS)
            self.assertLessEqual(len(json.dumps({"requests": call})), rd.BATCH_MAX_BYTES + 1000)

    def test_cell_colors(self):
        upd = next(r for r in self.fake.format_requests() if "updateCells" in r)["updateCells"]
        first = upd["rows"][0]["values"]
        for i in (1, 2, 3):
            self.assertEqual(first[i]["userEnteredFormat"]["backgroundColor"], rd.rgb("#f0f4ff"))
        self.assertEqual(first[4]["userEnteredFormat"]["backgroundColor"], rd.rgb("#fde8e8"))
        self.assertEqual(first[5]["userEnteredFormat"]["backgroundColor"], rd.rgb("#fff3cd"))
        self.assertEqual(first[0]["userEnteredFormat"]["backgroundColor"], rd.rgb("#f0f4ff"))
        second = upd["rows"][1]["values"]
        self.assertEqual(second[0]["userEnteredFormat"]["backgroundColor"], rd.rgb("#ffffff"))
        tags_fmt = first[9]["userEnteredFormat"]
        self.assertEqual(tags_fmt["wrapStrategy"], "WRAP")
        self.assertEqual(tags_fmt["backgroundColor"], rd.rgb("#f0f4ff"))
        self.assertEqual(second[9]["userEnteredFormat"]["backgroundColor"], rd.rgb("#ffffff"))
        self.assertEqual(upd["range"]["endColumnIndex"], 10)
        self.assertIn("userEnteredFormat.wrapStrategy", upd["fields"])
        reqs = self.fake.format_requests()
        widths = {r["updateDimensionProperties"]["range"]["startIndex"]:
                  r["updateDimensionProperties"]["properties"]["pixelSize"]
                  for r in reqs if "updateDimensionProperties" in r
                  and r["updateDimensionProperties"]["range"]["dimension"] == "COLUMNS"}
        self.assertEqual(widths, {0: 220, 1: 160, 2: 110, 3: 100, 4: 110, 5: 160, 6: 100, 7: 100,
                                  8: 100, 9: 280})
        # başlık / kolon başlığı / TOTAL J sütununu da kapsıyor
        for r in reqs:
            if ("repeatCell" in r and "backgroundColor" in r["repeatCell"]["fields"]
                    and r["repeatCell"]["range"]["startRowIndex"] >= upd["range"]["startRowIndex"] - 2):
                self.assertEqual(r["repeatCell"]["range"]["endColumnIndex"], 10)
        # Alice 60% → >30 kırmızı
        self.assertEqual(first[8]["userEnteredFormat"]["backgroundColor"], rd.rgb("#f28b82"))
        for pct, bg in (("31%", "#f28b82"), ("20%", "#fbbc04"), ("6%", "#fff475"),
                        ("5.0%", "#d4edda"), ("—", "#ffffff")):
            self.assertEqual(rd.pct_colors(pct)[0], bg)

    def test_large_user_count_chunks(self):
        low, allc = make_day(n_users=1200)
        tr, ar, ur, tl, ta = rd.build_pivots(low, allc)
        fake = FakeSheets()
        rd.write_daily_sheet(fake, "2026-10-14", tr, ar, ur, tl, ta)
        upd = [r for r in fake.format_requests() if "updateCells" in r]
        self.assertEqual(len(upd), 3)  # 1203 satır → 500+500+203
        covered = []
        for r in upd:
            rr = r["updateCells"]["range"]
            self.assertLessEqual(len(r["updateCells"]["rows"]), rd.USER_FMT_CHUNK_ROWS)
            covered.extend(range(rr["startRowIndex"], rr["endRowIndex"]))
        self.assertEqual(len(covered), len(set(covered)))
        self.assertEqual(len(covered), len(ur))
        for call in fake.batch_calls:
            self.assertLessEqual(len(call), rd.BATCH_MAX_REQS)
            self.assertLessEqual(len(json.dumps({"requests": call})), rd.BATCH_MAX_BYTES + 1000)


class ChunkRequestsTest(unittest.TestCase):
    def test_count_and_size_limits(self):
        reqs = [{"x": i} for i in range(1234)]
        chunks = rd.chunk_requests(reqs)
        self.assertEqual([len(c) for c in chunks], [500, 500, 234])
        self.assertEqual([r for c in chunks for r in c], reqs)
        big = [{"x": "a" * 400}] * 10
        chunks = rd.chunk_requests(big, max_bytes=1000)
        self.assertTrue(all(len(c) <= 2 for c in chunks))
        self.assertEqual(sum(len(c) for c in chunks), 10)


class WeeklySheetTest(unittest.TestCase):
    def setUp(self):
        self.week = []
        for i in range(7):
            ds = f"2026-10-{5 + i:02d}"
            low, allc = make_day(n_users=160, seed_offset=i)
            # 'Eve': 1. gün tek düşük puan, sonraki 5 gün yalnızca normal puanlı
            # chat'ler (farklı yazımlarla) — haftalık Total Chats bunları da saymalı
            if i == 0:
                eve = chat(f"eve-{i}", 1, "Eve", tag="(Refund_Request)")
                low.append(eve)
                allc.append(eve)
            elif i <= 5:
                allc.append(chat(f"eve-{i}-a", 5, "eve"))
                allc.append(chat(f"eve-{i}-b", 4, " EVE"))
            tr, ar, ur, tl, ta = rd.build_pivots(low, allc)
            self.week.append({"date": ds, "low_chats": low, "all_count": allc,
                              "tag_rows": tr, "agent_rows": ar, "user_rows": ur,
                              "total_low": tl, "total_all": ta})
        # Bir gün verisi eksik (hata yolu) + eski formatta user_rows anahtarı yok
        self.week[6] = {"date": "2026-10-11", "low_chats": [], "all_count": [],
                        "tag_rows": [], "agent_rows": [], "total_low": 0, "total_all": 0}
        self.fake = FakeSheets()
        rd.write_weekly_sheet(self.fake, self.week, "2026-10-05", "2026-10-11", None)
        self.rows = rows_by_range(self.fake)

    def test_aggregation_and_layout(self):
        rows = self.rows
        idx = {r[0]: i for i, r in enumerate(rows) if r and isinstance(r[0], str)}
        agent_hdr = idx["👤 AGENT BREAKDOWN — WEEKLY"]
        user_hdr  = idx["👤 USER BREAKDOWN — WEEKLY"]
        self.assertGreater(user_hdr, agent_hdr)
        self.assertEqual(rows[user_hdr - 2][0], "TOTAL")      # agent TOTAL
        self.assertEqual(rows[user_hdr - 1], [""] * 10)
        self.assertEqual(rows[user_hdr + 1][0], "Chat Name")
        data = rows[user_hdr + 2:-1]
        self.assertEqual(len(data), 164)
        # 6 gün × (Payment_Delay x2 + tag'siz x1), R1 ve R2 tag'leri birleşik
        self.assertEqual(data[0], ["Alice", "", "", "", 12, 6, 18, 30, "60.0%", "Payment_Delay (12), — (6)"])
        # Total Chats tüm haftanın chat'lerinden: 1 (düşük) + 5 gün × 2 (normal)
        eve = next(r for r in data if r[0] == "Eve")
        self.assertEqual(eve, ["Eve", "", "", "", 1, 0, 1, 11, "9.1%", "Refund_Request"])
        self.assertEqual(rows[user_hdr + 1][9], "Low Rating Tags")
        self.assertEqual(rows[-1][9], "")
        # Agent haftalık tablosu eski hesapla aynı kalıyor (günlük satır toplamı)
        agent_hdr = next(i for i, r in enumerate(rows) if r[0] == "👤 AGENT BREAKDOWN — WEEKLY")
        agent_a = next(r for r in rows[agent_hdr + 2:] if r[0] == "Agent A")
        self.assertEqual(agent_a[4], sum(next(a for a in d["agent_rows"] if a[0] == "Agent A")[4]
                                         for d in self.week if d["agent_rows"]))
        self.assertEqual(rows[-1][0], "TOTAL")
        self.assertEqual(rows[-1][6], sum(r[6] for r in data))
        self.assertTrue(all(len(r) == 10 for r in rows))
        keys = [(-r[6], -r[4], r[0].lower()) for r in data]
        self.assertEqual(keys, sorted(keys))

    def test_existing_sections_unchanged(self):
        # kullanıcı bölümü olmadan üretilen satırlar mevcut çıktının ön eki olmalı
        week_wo_users = [{k: v for k, v in d.items() if k not in ("user_rows", "low_chats")}
                         for d in self.week]
        fake = FakeSheets()
        rd.write_weekly_sheet(fake, week_wo_users, "2026-10-05", "2026-10-11", None)
        base = rows_by_range(fake)
        user_hdr = next(i for i, r in enumerate(self.rows) if r[0] == "👤 USER BREAKDOWN — WEEKLY")
        self.assertEqual(self.rows[:user_hdr - 1], base[:user_hdr - 1])
        # mevcut format istekleri aynen korunuyor (kullanıcı istekleri sona eklendi)
        base_reqs = fake.format_requests()
        new_reqs  = self.fake.format_requests()

        def in_user(r):
            span = req_rows(r)
            return bool(span) and span[0] >= user_hdr

        # kullanıcı bölümü dışındaki istekler kullanıcı verisinden bağımsız ve ona taşmıyor
        self.assertEqual([r for r in new_reqs if not in_user(r)],
                         [r for r in base_reqs if not in_user(r)])
        for r in new_reqs:
            span = req_rows(r)
            if span and not in_user(r):
                self.assertLessEqual(span[1], user_hdr - 1, r)
        user_reqs = [r for r in new_reqs if in_user(r)]
        # boş tabloya göre fark: veri updateCells + dikey ortalama
        self.assertEqual(len(user_reqs), len([r for r in base_reqs if in_user(r)]) + 2)
        for r in user_reqs:
            self.assertGreaterEqual(req_rows(r)[0], user_hdr)
            self.assertLessEqual(req_rows(r)[1], len(self.rows))
        for call in self.fake.batch_calls:
            self.assertLessEqual(len(call), rd.BATCH_MAX_REQS)


class MainModeTest(unittest.TestCase):
    """Haftalık rapor sadece mode=weekly ile; Pazartesi otomatik tetiklenmez."""

    def _run_main(self, argv, monday):
        from unittest import mock
        low, allc = make_day(n_users=5)
        calls = {"daily": [], "weekly": 0, "fetch_days": []}

        def fake_daily(svc, date_str, *a):
            calls["daily"].append(date_str)

        def fake_weekly(*a, **k):
            calls["weekly"] += 1

        def fetch_low(ds):
            calls["fetch_days"].append(ds)
            return list(low)

        with mock.patch.object(rd, "sheets_service", return_value=FakeSheets()), \
             mock.patch.object(rd, "sofia_yesterday", return_value="2026-10-18"), \
             mock.patch.object(rd, "is_monday", return_value=monday), \
             mock.patch.object(rd, "fetch_low_rating_chats", side_effect=fetch_low), \
             mock.patch.object(rd, "fetch_all_chats_count", return_value=list(allc)), \
             mock.patch.object(rd, "attach_user_codes", side_effect=lambda cs: cs), \
             mock.patch.object(rd, "write_daily_sheet", side_effect=fake_daily), \
             mock.patch.object(rd, "write_weekly_sheet", side_effect=fake_weekly), \
             mock.patch.object(rd, "send_error_email"), \
             mock.patch.object(rd.time, "sleep"), \
             mock.patch.object(sys, "argv", argv), \
             contextlib.redirect_stdout(io.StringIO()):
            rd.main()
        return calls

    def test_monday_scheduled_daily_only(self):
        for argv in (["rating_dashboard.py"], ["rating_dashboard.py", "daily"]):
            calls = self._run_main(argv, monday=True)
            self.assertEqual(calls["daily"], ["2026-10-18"])
            self.assertEqual(calls["weekly"], 0)
            self.assertEqual(calls["fetch_days"], ["2026-10-18"])  # sadece dün çekildi

    def test_manual_weekly_still_works(self):
        calls = self._run_main(["rating_dashboard.py", "weekly"], monday=False)
        self.assertEqual(calls["daily"], ["2026-10-18"])
        self.assertEqual(calls["weekly"], 1)
        self.assertEqual(len(calls["fetch_days"]), 1 + 7 + 7)


def section_blocks(rows, headers):
    """Bölüm başlığı → (başlık satırı, sütun başlığı satırı, TOTAL satırı)."""
    out = {}
    for h in headers:
        hi = next(i for i, r in enumerate(rows) if r[0] == h)
        ti = next(i for i in range(hi + 2, len(rows)) if rows[i][0] == "TOTAL")
        out[h] = (hi, hi + 1, ti)
    return out


def aligns(reqs, key):
    """key ('horizontalAlignment' / 'verticalAlignment') sadece hizalama değiştiren istekler."""
    res = []
    for r in reqs:
        rc = r.get("repeatCell")
        if rc and rc["fields"] == f"userEnteredFormat.{key}":
            rr = rc["range"]
            res.append((rr["startRowIndex"], rr["endRowIndex"], rr["startColumnIndex"],
                        rr["endColumnIndex"], rc["cell"]["userEnteredFormat"][key]))
    return res


class FormattingLayoutTest(unittest.TestCase):
    DAILY_TABLES = {  # bölüm → (ortalı sütun aralıkları, sola dayalı sütun aralıkları)
        "🏷️ BREAKDOWN BY TAG":   ({(1, 5)}, {(0, 1)}),
        "👤 BREAKDOWN BY AGENT": ({(1, 6)}, {(0, 1)}),
        "👤 BREAKDOWN BY USER":  ({(2, 9)}, {(0, 2), (9, 10)}),
    }
    WEEKLY_TABLES = {
        "📅 DAILY BREAKDOWN":           ({(2, 7)}, {(0, 2)}),
        "🏷️ TAG BREAKDOWN — WEEKLY":    ({(1, 5)}, {(0, 1)}),
        "👤 AGENT BREAKDOWN — WEEKLY":  ({(1, 6)}, {(0, 1)}),
        "👤 USER BREAKDOWN — WEEKLY":   ({(2, 9)}, {(0, 2), (9, 10)}),
    }

    @classmethod
    def setUpClass(cls):
        low, allc = make_day(n_users=160)
        tr, ar, ur, tl, ta = rd.build_pivots(low, allc)
        cls.daily = FakeSheets()
        rd.write_daily_sheet(cls.daily, "2026-10-14", tr, ar, ur, tl, ta)
        cls.daily_rows = rows_by_range(cls.daily)
        week = [{"date": f"2026-10-{5 + i:02d}", "all_count": allc, "tag_rows": tr,
                 "agent_rows": ar, "user_rows": ur, "total_low": tl, "total_all": ta}
                for i in range(7)]
        cls.weekly = FakeSheets()
        rd.write_weekly_sheet(cls.weekly, week, "2026-10-05", "2026-10-11", None)
        cls.weekly_rows = rows_by_range(cls.weekly)

    def _check_title_and_freeze(self, fake, height):
        reqs = fake.format_requests()
        merges = [r["mergeCells"] for r in reqs if "mergeCells" in r]
        self.assertEqual(len(merges), 1)
        m = merges[0]["range"]
        self.assertEqual((m["startRowIndex"], m["endRowIndex"],
                          m["startColumnIndex"], m["endColumnIndex"]), (0, 1, 0, 10))
        self.assertEqual(merges[0]["mergeType"], "MERGE_ALL")
        title = [r["repeatCell"] for r in reqs if "repeatCell" in r
                 and r["repeatCell"]["range"]["startRowIndex"] == 0
                 and "backgroundColor" in r["repeatCell"]["fields"]]
        self.assertEqual(len(title), 1)
        fmt = title[0]["cell"]["userEnteredFormat"]
        self.assertEqual(title[0]["range"]["endColumnIndex"], 10)
        self.assertEqual(fmt["horizontalAlignment"], "CENTER")
        self.assertEqual(fmt["verticalAlignment"], "MIDDLE")
        self.assertEqual(fmt["wrapStrategy"], "WRAP")
        heights = [r["updateDimensionProperties"] for r in reqs if "updateDimensionProperties" in r
                   and r["updateDimensionProperties"]["range"]["dimension"] == "ROWS"]
        self.assertEqual([(h["range"]["startIndex"], h["properties"]["pixelSize"]) for h in heights],
                         [(0, height)])
        freezes = [r["updateSheetProperties"] for r in reqs if "updateSheetProperties" in r]
        self.assertEqual(len(freezes), 1)
        self.assertEqual(freezes[0]["properties"]["gridProperties"],
                         {"frozenRowCount": 1, "frozenColumnCount": 0})
        # Başlık satırına yalnızca başlık istekleri dokunuyor
        for r in reqs:
            span = req_rows(r)
            if span and span[0] == 0:
                self.assertTrue("mergeCells" in r or r["repeatCell"] is title[0], r)

    def test_daily_title_and_freeze(self):
        self._check_title_and_freeze(self.daily, 36)

    def test_weekly_title_and_freeze(self):
        self._check_title_and_freeze(self.weekly, 40)

    def _check_alignment(self, fake, rows, tables):
        reqs   = fake.format_requests()
        blocks = section_blocks(rows, tables)
        h_al   = aligns(reqs, "horizontalAlignment")
        for name, (center, left) in tables.items():
            _, col_row, total_row = blocks[name]
            mine = [a for a in h_al if a[0] == col_row]
            for a in mine:
                self.assertEqual(a[1], total_row + 1, (name, a))
            self.assertEqual({(a[2], a[3]) for a in mine if a[4] == "CENTER"}, center, name)
            self.assertEqual({(a[2], a[3]) for a in mine if a[4] == "LEFT"}, left, name)
        # Hiçbir hizalama aralığı iki tablo bloğunu kesmiyor
        spans = [(b[1], b[2] + 1) for b in blocks.values()]
        for a in h_al:
            self.assertIn((a[0], a[1]), spans, a)
        self.assertEqual(len(h_al), sum(len(c) + len(l) for c, l in tables.values()))
        return blocks

    def test_daily_alignment(self):
        blocks = self._check_alignment(self.daily, self.daily_rows, self.DAILY_TABLES)
        _, col_row, total_row = blocks["👤 BREAKDOWN BY USER"]
        v_al = aligns(self.daily.format_requests(), "verticalAlignment")
        self.assertEqual(v_al, [(col_row + 1, total_row, 0, 10, "MIDDLE")])

    def test_weekly_alignment(self):
        blocks = self._check_alignment(self.weekly, self.weekly_rows, self.WEEKLY_TABLES)
        _, col_row, total_row = blocks["👤 USER BREAKDOWN — WEEKLY"]
        v_al = aligns(self.weekly.format_requests(), "verticalAlignment")
        self.assertEqual(v_al, [(col_row + 1, total_row, 0, 10, "MIDDLE")])

    def _bars(self, fake, rows, headers):
        reqs = fake.format_requests()
        for h in headers:
            hi = next(i for i, r in enumerate(rows) if r[0] == h)
            bar = [r["repeatCell"]["range"] for r in reqs if "repeatCell" in r
                   and r["repeatCell"]["range"]["startRowIndex"] == hi
                   and "backgroundColor" in r["repeatCell"]["fields"]]
            self.assertEqual(len(bar), 1, h)
            self.assertEqual((bar[0]["startColumnIndex"], bar[0]["endColumnIndex"]), (0, 10), h)

    def test_section_bars_full_width(self):
        self._bars(self.daily, self.daily_rows, self.DAILY_TABLES)
        self._bars(self.weekly, self.weekly_rows,
                   ["📌 WEEKLY SUMMARY"] + list(self.WEEKLY_TABLES))

    def test_col_header_and_total_match_table_width(self):
        widths = {"🏷️ BREAKDOWN BY TAG": 5, "👤 BREAKDOWN BY AGENT": 6, "👤 BREAKDOWN BY USER": 10}
        reqs = self.daily.format_requests()
        for name, (_, col_row, total_row) in section_blocks(self.daily_rows, widths).items():
            for row in (col_row, total_row):
                bg = [r["repeatCell"]["range"] for r in reqs if "repeatCell" in r
                      and r["repeatCell"]["range"]["startRowIndex"] == row
                      and "backgroundColor" in r["repeatCell"]["fields"]]
                self.assertEqual([(b["startColumnIndex"], b["endColumnIndex"]) for b in bg],
                                 [(0, widths[name])], (name, row))

    def test_weekly_column_widths_and_agent_total_chats_color(self):
        reqs = self.weekly.format_requests()
        widths = {r["updateDimensionProperties"]["range"]["startIndex"]:
                  r["updateDimensionProperties"]["properties"]["pixelSize"]
                  for r in reqs if "updateDimensionProperties" in r
                  and r["updateDimensionProperties"]["range"]["dimension"] == "COLUMNS"}
        self.assertEqual(widths, {0: 220, 1: 160, 2: 110, 3: 120, 4: 110, 5: 100, 6: 100, 7: 100,
                                  8: 100, 9: 280})
        _, col_row, total_row = section_blocks(self.weekly_rows, ["👤 AGENT BREAKDOWN — WEEKLY"])[
            "👤 AGENT BREAKDOWN — WEEKLY"]
        grey = [r for r in reqs if "repeatCell" in r
                and col_row < r["repeatCell"]["range"]["startRowIndex"] < total_row
                and r["repeatCell"]["range"]["startColumnIndex"] == 4
                and r["repeatCell"]["cell"]["userEnteredFormat"].get("textFormat", {})
                    .get("foregroundColor") == rd.rgb("#444444")]
        self.assertEqual(len(grey), total_row - col_row - 1)

    def test_weekly_summary_zebra_covers_label_column(self):
        reqs = self.weekly.format_requests()
        hdr = next(i for i, r in enumerate(self.weekly_rows) if r[0] == "📌 WEEKLY SUMMARY")
        first = hdr + 2
        bg = [r["repeatCell"] for r in reqs if "repeatCell" in r
              and r["repeatCell"]["range"]["startRowIndex"] == first
              and "backgroundColor" in r["repeatCell"]["fields"]]
        self.assertEqual(len(bg), 1)
        self.assertEqual((bg[0]["range"]["startColumnIndex"], bg[0]["range"]["endColumnIndex"]), (0, 3))
        self.assertEqual(bg[0]["cell"]["userEnteredFormat"]["backgroundColor"], rd.rgb("#f0f4ff"))


class WeeklyUserTotalFallbackTest(unittest.TestCase):
    def test_without_all_count_uses_daily_rows(self):
        low, allc = make_day(n_users=3)
        tr, ar, ur, tl, ta = rd.build_pivots(low, allc)
        day = {"date": "2026-10-14", "tag_rows": tr, "agent_rows": ar,
               "user_rows": ur, "total_low": tl, "total_all": ta}
        fake = FakeSheets()
        rd.write_weekly_sheet(fake, [day, dict(day, date="2026-10-15")],
                              "2026-10-14", "2026-10-15", None)
        rows = rows_by_range(fake)
        alice = next(r for r in rows if r[0] == "Alice")
        self.assertEqual(alice, ["Alice", "", "", "", 4, 2, 6, 10, "60.0%", "Payment_Delay (4), — (2)"])


class NoPrintTest(unittest.TestCase):
    def test_usernames_not_logged(self):
        import io
        import contextlib
        low, allc = make_day(n_users=5)
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            tr, ar, ur, tl, ta = rd.build_pivots(low, allc)
            rd.write_daily_sheet(FakeSheets(), "2026-10-14", tr, ar, ur, tl, ta)
            rd.write_weekly_sheet(FakeSheets(), [{"date": "2026-10-14", "tag_rows": tr,
                                  "agent_rows": ar, "user_rows": ur,
                                  "total_low": tl, "total_all": ta}],
                                  "2026-10-14", "2026-10-14", None)
        out = buf.getvalue().lower()
        for name in ("alice", "bob", "carl", "user0000", "payment_delay", "other_topic"):
            self.assertNotIn(name, out)
        # kullanıcı tablosu özeti sadece sayı olarak loglanıyor
        named_low = sum(1 for c in low if rd.get_username(c))
        self.assertIn(f"user table: {len(ur)} rows, 10 cols, "
                      f"low rating tags filled: {len(ur)}/{len(ur)} cells, "
                      f"tagged low chats: {named_low}, with username: 0/{len(ur)}, "
                      f"with user code: 0/{len(ur)}, "
                      f"with brand: 0/{len(ur)}", out)


def detail(code, field_id=None, field_name="04 User Code"):
    """chats/{id} detay yanıtı biçiminde customVariables."""
    return {"customVariables": [
        {"formType": "customVariable", "fieldId": "x-1", "fieldName": "01 Username",
         "value": ["someone"]},
        {"formType": "customVariable", "fieldId": field_id or rd.USER_CODE_FIELD_ID,
         "fieldName": field_name, "value": [code]},
    ]}


class UserCodeTest(unittest.TestCase):
    def test_extract(self):
        self.assertEqual(rd.extract_user_code(detail("123")), "123")
        # fieldId değişse de ada göre bulunur; fieldResults da taranır
        self.assertEqual(rd.extract_user_code(
            {"fieldResults": detail(" 77 ", field_id="other")["customVariables"]}), "77")
        self.assertEqual(rd.extract_user_code(detail("")), "")
        self.assertEqual(rd.extract_user_code({"customVariables": []}), "")
        self.assertEqual(rd.extract_user_code({}), "")

    def test_grouping_by_code(self):
        low = [
            chat(1, 1, "Ann", code="100", visitor="v1"),
            chat(2, 2, "ann_new", code="100", visitor="v2"),  # aynı kod, farklı ad → tek satır
            chat(3, 1, "Sam", code="200"),
            chat(4, 2, "sam", code="201"),                    # aynı ad, farklı kod → ayrı satır
            chat(5, 1, "ANN"),                                # kodsuz, ad tek kod grubuyla eşleşiyor
            chat(6, 2, "sam"),                                # kodsuz, ad belirsiz → ada göre
            chat(7, 1, None, code="300"),                     # adsız ama kodlu
            chat(8, 2, "", ),                                 # ne ad ne kod → atlanır
        ]
        allc = list(low) + [
            chat(20, 5, "someone_else", visitor="v1"),        # aynı ziyaretçi → Ann grubuna
            chat(21, 5, "Ann_New"),                           # ad eşleşmesi → Ann grubuna
            chat(22, 5, "SAM"),                               # belirsiz ad → kodsuz sam grubuna
        ]
        _, _, ur, _, _ = rd.build_pivots(low, allc)
        by_code = {r[2]: r for r in ur if r[2]}
        self.assertEqual(by_code["100"][:9], ["Ann", "", "100", "", 2, 1, 3, 5, "60.0%"])
        self.assertEqual(by_code["200"][:8], ["Sam", "", "200", "", 1, 0, 1, 1])
        self.assertEqual(by_code["201"][:8], ["sam", "", "201", "", 0, 1, 1, 1])
        self.assertEqual(by_code["300"][:8], ["", "", "300", "", 1, 0, 1, 1])
        nocode = [r for r in ur if not r[2]]
        self.assertEqual([r[:8] for r in nocode], [["sam", "", "", "", 0, 1, 1, 2]])
        self.assertEqual(len(ur), 5)
        self.assertEqual(sum(r[6] for r in ur), 7)

    def test_weekly_merges_by_code(self):
        week = []
        for i, name in enumerate(["Ann", "ann_renamed"]):
            low = [chat(f"d{i}", 1, name, code="100")]
            allc = low + [chat(f"n{i}", 5, name)]
            tr, ar, ur, tl, ta = rd.build_pivots(low, allc)
            week.append({"date": f"2026-10-0{5 + i}", "low_chats": low, "all_count": allc,
                         "tag_rows": tr, "agent_rows": ar, "user_rows": ur,
                         "total_low": tl, "total_all": ta})
        fake = FakeSheets()
        rd.write_weekly_sheet(fake, week, "2026-10-05", "2026-10-06", None)
        rows = rows_by_range(fake)
        hdr = next(i for i, r in enumerate(rows) if r[0] == "👤 USER BREAKDOWN — WEEKLY")
        self.assertEqual(rows[hdr + 1][:4], ["Chat Name", "Username", "User Code", "Brand"])
        self.assertEqual(rows[hdr + 2][:9], ["Ann", "", "100", "", 2, 0, 2, 4, "50.0%"])
        self.assertEqual(rows[hdr + 3][:4], ["TOTAL", "", "", ""])
        # Ham chat yoksa günlük satırlar koda göre birleşir
        old = [{k: v for k, v in d.items() if k not in ("low_chats", "all_count")} for d in week]
        fake = FakeSheets()
        rd.write_weekly_sheet(fake, old, "2026-10-05", "2026-10-06", None)
        rows = rows_by_range(fake)
        hdr = next(i for i, r in enumerate(rows) if r[0] == "👤 USER BREAKDOWN — WEEKLY")
        self.assertEqual(rows[hdr + 2][:9], ["Ann", "", "100", "", 2, 0, 2, 4, "50.0%"])

    def test_attach_user_codes_fetches_details_without_logging(self):
        from unittest import mock
        rd._user_code_cache.clear()
        codes = {"1": "555001", "2": "", "3": None}
        urls = []

        class Resp:
            def __init__(self, status, body=None):
                self.status_code, self._body = status, body

            def json(self):
                return self._body

        def fake_get(url, headers, timeout):
            urls.append(url)
            cid = url.split("/chats/")[1].split("?")[0]
            if codes[cid] is None:
                return Resp(500)
            return Resp(200, detail(codes[cid]))

        chats = [chat(1, 1, "a"), chat(2, 1, "b"), chat(3, 2, "c"),
                 dict(chat(4, 1, "d"), customVariables=detail("999")["customVariables"])]
        buf = io.StringIO()
        with mock.patch.object(rd.requests, "get", side_effect=fake_get), \
             mock.patch.object(rd.time, "sleep"), \
             contextlib.redirect_stdout(buf):
            rd.attach_user_codes(chats)
        self.assertEqual([c["_userCode"] for c in chats], ["555001", "", "", "999"])
        self.assertTrue(all("/api/LiveChat/chats/" in u and "include=chatFieldResult" in u
                            for u in urls))
        self.assertEqual(sum(1 for u in urls if "/chats/3?" in u), 3)  # hata → yeniden dene
        self.assertNotIn("/chats/4?", " ".join(urls))                  # zaten kodlu → istek yok
        out = buf.getvalue()
        self.assertIn("2/4 chats have User Code, 3/4 have Username "
                      "(detail requests: 3, errors: 1)", out)
        for secret in ("555001", "999", "someone"):
            self.assertNotIn(secret, out)
        # Önbellek: başarılı olanlar tekrar istenmez, hatalı olan yeniden denenir
        urls.clear()
        again = [chat(1, 1, "a"), chat(3, 2, "c")]
        with mock.patch.object(rd.requests, "get", side_effect=fake_get), \
             mock.patch.object(rd.time, "sleep"), \
             contextlib.redirect_stdout(io.StringIO()):
            rd.attach_user_codes(again)
        self.assertEqual(again[0]["_userCode"], "555001")
        self.assertTrue(urls and all("/chats/3?" in u for u in urls))

    def test_user_codes_not_logged_by_sheet_writers(self):
        low = [chat(1, 1, "Ann", code="8675309"), chat(2, 2, "Bo", code="4242")]
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            tr, ar, ur, tl, ta = rd.build_pivots(low, low)
            rd.write_daily_sheet(FakeSheets(), "2026-10-14", tr, ar, ur, tl, ta)
            rd.write_weekly_sheet(FakeSheets(), [{"date": "2026-10-14", "low_chats": low,
                                  "all_count": low, "tag_rows": tr, "agent_rows": ar,
                                  "user_rows": ur, "total_low": tl, "total_all": ta}],
                                  "2026-10-14", "2026-10-14", None)
        out = buf.getvalue()
        for secret in ("8675309", "4242", "Ann", "Bo"):
            self.assertNotIn(secret, out)
        self.assertIn("with User Code: 2/2", out)


class BrandTest(unittest.TestCase):
    def test_brand_display(self):
        self.assertEqual(rd.get_brand(chat(1, 1, "a", brand="SUPERBETIN")), "Superbetin")
        self.assertEqual(rd.get_brand(chat(1, 1, "a", brand="TurkBet")), "TurkBet")
        self.assertEqual(rd.get_brand(chat(1, 1, "a")), "")
        self.assertEqual(rd.get_brand_key(chat(1, 1, "a", brand="BETSAT")), "camp-betsat")
        self.assertEqual(rd.get_brand_key({"campaignId": "x9"}), "x9")

    def test_same_code_or_name_in_two_brands_stays_separate(self):
        low = [
            chat(1, 1, "ann", code="100", brand="SUPERBETIN"),
            chat(2, 2, "ann", code="100", brand="BETSAT"),
            chat(3, 1, "zed", brand="SUPERBETIN"),
            chat(4, 1, "zed", brand="TURKBET"),
        ]
        allc = list(low) + [
            dict(chat(10, 5, "ANN"), campaignId="camp-betsat"),   # Betsat'taki ann'e
            dict(chat(11, 5, "zed"), campaignId="camp-turkbet"),  # Turkbet'teki zed'e
            chat(12, 5, "zed"),                                   # markasız, ad belirsiz → yok
            chat(13, 5, "Ann"),                                   # markasız, ad belirsiz → yok
        ]
        _, _, ur, _, _ = rd.build_pivots(low, allc)
        got = sorted([r[0]] + r[2:8] for r in ur)
        self.assertEqual(got, [
            ["ann", "100", "Betsat", 0, 1, 1, 2],
            ["ann", "100", "Superbetin", 1, 0, 1, 1],
            ["zed", "", "Superbetin", 1, 0, 1, 1],
            ["zed", "", "Turkbet", 1, 0, 1, 2],
        ])

    def test_brandless_all_chats_match_unique_name(self):
        low = [chat(1, 1, "solo", code="5", brand="BETSAT")]
        allc = list(low) + [chat(2, 5, "SOLO"), chat(3, 4, "other")]
        _, _, ur, _, _ = rd.build_pivots(low, allc)
        self.assertEqual(ur[0][:8], ["solo", "", "5", "Betsat", 1, 0, 1, 2])

    def test_detail_fills_campaign(self):
        from unittest import mock
        rd._user_code_cache.clear()
        urls = []

        class Resp:
            status_code = 200

            def json(self):
                return dict(detail("42"), campaign={"id": "camp-turkbet", "name": "TURKBET"})

        def fake_get(url, headers, timeout):
            urls.append(url)
            return Resp()

        c = chat(77, 1, "x")
        with mock.patch.object(rd.requests, "get", side_effect=fake_get), \
             contextlib.redirect_stdout(io.StringIO()):
            rd.attach_user_codes([c])
        self.assertEqual((rd.get_user_code(c), rd.get_brand(c)), ("42", "Turkbet"))
        self.assertIn("include=campaign", urls[0])


class RealUsernameTest(unittest.TestCase):
    def test_extract_and_columns(self):
        d = detail("100")
        self.assertEqual(rd.extract_real_username(d), "someone")
        # '01 Username' ile '04 User Code' karışmıyor
        self.assertEqual(rd.extract_user_code(d), "100")
        low = [dict(chat(1, 1, "typed name", code="100"), _realUsername="acct_1"),
               dict(chat(2, 2, "other typing", code="100"), _realUsername="acct_1"),
               dict(chat(3, 1, "guest"), _realUsername="")]
        _, _, ur, _, _ = rd.build_pivots(low, low)
        self.assertEqual([r[:4] for r in ur],
                         [["typed name", "acct_1", "100", ""], ["guest", "", "", ""]])
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            rd.write_daily_sheet(FakeSheets(), "2026-10-14", *rd.build_pivots(low, low))
        self.assertNotIn("acct_1", buf.getvalue())
        self.assertIn("with Username: 1/2", buf.getvalue())


if __name__ == "__main__":
    unittest.main()
