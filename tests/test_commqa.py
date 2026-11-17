"""
commqa.py — görsel düzeltme testleri.

Çalıştırma:  python -m unittest discover -s tests -v

Golden test: tests/golden/commqa_before.py, değişiklik öncesi commqa.py'nin
birebir kopyasıdır. Aynı sahte chat verisiyle eski ve yeni kodun sayfalara
yazdığı değerler (write_values matrisleri + "Chati Ac" link hücreleri)
birebir aynı olmalı. Google kütüphaneleri kurulu değilse stub'lanır.
"""
import contextlib
import copy
import importlib.util
import io
import os
import sys
import types
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

for k in ("COMM100_SITE_ID", "COMM100_API_KEY", "COMM100_EMAIL",
          "COMMQA_SPREADSHEET_ID", "GCP_CREDENTIALS", "GMAIL_USER",
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


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


new = _load("commqa_new", os.path.join(ROOT, "scripts", "commqa.py"))
old = _load("commqa_old", os.path.join(ROOT, "tests", "golden", "commqa_before.py"))


# ── Sahte Sheets servisi ─────────────────────────────────────
class _Exec:
    def __init__(self, fn):
        self._fn = fn

    def execute(self):
        return self._fn()


class SheetsError(Exception):
    pass


def _overlap(a, b):
    def span(r, k):
        return r.get(f"start{k}Index", 0), r.get(f"end{k}Index", 10 ** 9)
    for k in ("Row", "Column"):
        a1, a2 = span(a, k)
        b1, b2 = span(b, k)
        if a2 <= b1 or b2 <= a1:
            return False
    return True


class FakeSheets:
    """
    Gerçek API'ye yakın davranış: mergeCells mevcut merge veya filtre
    aralığıyla çakışırsa hata verir; unmergeCells aralıktaki merge'leri siler.
    """
    def __init__(self, filter_start_row=1):
        self.events  = []       # ("clear"|"values"|"batch", payload) sırası
        self.sheets  = {}
        self.merges  = []
        self.filter  = {"startRowIndex": filter_start_row, "startColumnIndex": 0,
                        "endColumnIndex": 10}
        self._next   = 500

    def spreadsheets(self):
        return self

    def values(self):
        return self

    def get(self, spreadsheetId, ranges=None, fields=None, range=None):
        if range is not None:          # values().get (log sayfası)
            return _Exec(lambda: {"values": []})
        if ranges is not None:         # report_site_format
            return _Exec(lambda: {"sheets": []})
        return _Exec(lambda: {"sheets": [{"properties": {"title": t, "sheetId": i}}
                                         for t, i in self.sheets.items()]})

    def clear(self, spreadsheetId, range, body):
        self.events.append(("clear", range))
        return _Exec(lambda: {})

    def update(self, spreadsheetId, range, valueInputOption, body):
        self.events.append(("values", (range, valueInputOption, copy.deepcopy(body["values"]))))
        return _Exec(lambda: {})

    def append(self, **kw):
        return _Exec(lambda: {})

    def batchUpdate(self, spreadsheetId, body):
        reqs = copy.deepcopy(body["requests"])

        def run():
            replies = []
            for r in reqs:
                if "addSheet" in r:
                    title = r["addSheet"]["properties"]["title"]
                    self.sheets[title] = self._next
                    replies.append({"addSheet": {"properties": {"sheetId": self._next}}})
                    self._next += 1
                    continue
                if "mergeCells" in r:
                    m = r["mergeCells"]["range"]
                    if any(x["sheetId"] == m["sheetId"] and _overlap(x, m) for x in self.merges):
                        raise SheetsError("You can't merge cells that are already merged")
                    if _overlap(self.filter, m):
                        raise SheetsError("You can't merge cells that intersect a filter")
                    self.merges.append(m)
                if "unmergeCells" in r:
                    u = r["unmergeCells"]["range"]
                    self.merges = [x for x in self.merges
                                   if not (x["sheetId"] == u["sheetId"] and _overlap(x, u))]
                replies.append({})
            self.events.append(("batch", reqs))
            return {"replies": replies}
        return _Exec(run)

    # yardımcılar
    def value_writes(self):
        return [p for kind, p in self.events if kind == "values"]

    def requests(self):
        return [r for kind, p in self.events if kind == "batch" for r in p]

    def link_cells(self):
        """apply_hyperlinks'in yazdığı hücre değerleri: (satır, sütun, değer)."""
        out = []
        for r in self.requests():
            uc = r.get("updateCells")
            if not uc:
                continue
            for i, row in enumerate(uc["rows"]):
                for j, cell in enumerate(row["values"]):
                    if "userEnteredValue" in cell:
                        out.append((uc["range"]["startRowIndex"] + i,
                                    uc["range"]["startColumnIndex"] + j,
                                    cell["userEnteredValue"],
                                    cell["userEnteredFormat"]["textFormat"]))
        return out


# ── Sahte chat verisi ────────────────────────────────────────
SECRET_NAMES = ["Zeynep Uzunisimli Soyadıyla Birlikte Çok Uzun Bir Ziyaretçi", "visitor_x"]
SECRET_COMMENTS = ["Çok yavaş cevap verildi, " * 10, "teşekkürler"]


def make_chats(n=40):
    chats = []
    sites = ["https://sitea.example/tr/page?x=1", "www.siteb.example/login",
             "siteb.example", "", "http://sub.sitec.example"]
    cats = ["General, Deposit (deposit_check)", "VIP TIER3, Casino",
            "VIP, Sport (bonus_query)", "", "Wild card (login_request), VIP TIER1",
            "Promotions"]
    for i in range(n):
        c = {
            "id": f"c{i}" if i % 13 else None,
            "chatId": None if i % 13 else f"alt{i}",
            "startTime": "2026-10-14T08:00:00Z",
            "endTime": "2026-10-14T08:%02d:%02dZ" % (i % 60, (i * 7) % 60),
            "chatAgents": [{"agent": {"displayName": f"Agent {i % 4}"}}],
            "chatWrapup": {"categoriesName": cats[i % len(cats)]},
            "requestingPageURL": sites[i % len(sites)],
            "postChatSurvey": {"ratingGrade": (i % 5) + 1 if i % 7 else None,
                               "ratingComment": SECRET_COMMENTS[i % 2] if i % 3 == 0 else ""},
        }
        if i % 4 == 0:
            c["preChatName"] = SECRET_NAMES[0]
        elif i % 4 == 1:
            c["visitorName"] = SECRET_NAMES[1]
        if c["id"] is None and i % 26 == 0:
            c["chatId"] = None   # link'siz satır
        chats.append(c)
    return chats


def run_writer(mod, fake, which, chats, date_str="2026-10-14"):
    fn = mod.write_rating_sheet if which == "rating" else mod.write_tag_sheet
    with contextlib.redirect_stdout(io.StringIO()) as buf:
        fn(fake, chats, date_str)
    return buf.getvalue()


class GoldenValuesTest(unittest.TestCase):
    """Eski ve yeni kodun yazdığı değerler birebir aynı."""

    def _compare(self, which, chats):
        f_old, f_new = FakeSheets(), FakeSheets()
        run_writer(old, f_old, which, chats)
        run_writer(new, f_new, which, chats)
        self.assertEqual(f_new.value_writes(), f_old.value_writes())
        old_links = [(r, c, v) for r, c, v, _ in f_old.link_cells()]
        new_links = [(r, c, v) for r, c, v, _ in f_new.link_cells()]
        self.assertEqual(new_links, old_links)
        # Chat Linki biçimi aynen (mavi, altı çizili, bağlantılı)
        self.assertEqual([t for *_, t in f_new.link_cells()], [t for *_, t in f_old.link_cells()])
        # Sekme adları ve temizlenen aralıklar aynı
        self.assertEqual(sorted(f_new.sheets), sorted(f_old.sheets))
        self.assertEqual([p for k, p in f_new.events if k == "clear"],
                         [p for k, p in f_old.events if k == "clear"])
        return f_new

    def test_rating_sheet_values_identical(self):
        chats = make_chats()
        f = self._compare("rating", chats)
        rng_, opt, values = f.value_writes()[0]
        self.assertEqual(rng_, "'Rating & Comment'!A1")
        self.assertEqual(opt, "RAW")
        self.assertEqual(len(values), 2 + len(chats))
        self.assertTrue(all(len(r) == 10 for r in values[1:]))

    def test_tag_sheet_values_identical(self):
        f = self._compare("tag", make_chats())
        self.assertEqual(f.value_writes()[0][0], "'🏷️ Tag Listesi'!A1")

    def test_empty_and_single(self):
        self._compare("rating", [])
        self._compare("tag", make_chats(1))

    def test_constants_unchanged(self):
        for name in ("RATING_SHEET", "TAG_SHEET", "LOG_SHEET", "QA_TAGS", "VIP_TAGS", "PORTAL_BASE"):
            self.assertEqual(getattr(new, name), getattr(old, name), name)

    def test_no_filter_or_structure_requests(self):
        f = FakeSheets()
        run_writer(new, f, "rating", make_chats())
        run_writer(new, f, "tag", make_chats())
        allowed = {"repeatCell", "updateDimensionProperties", "updateSheetProperties",
                   "mergeCells", "unmergeCells", "updateCells", "addSheet"}
        for r in f.requests():
            self.assertTrue(set(r) <= allowed, r)
            if "updateSheetProperties" in r:
                usp = r["updateSheetProperties"]
                self.assertEqual(usp["fields"], "gridProperties.frozenRowCount")
                self.assertEqual(usp["properties"]["gridProperties"], {"frozenRowCount": 2})
            if "updateCells" in r:   # sadece Chat Linki sütunu
                self.assertEqual(r["updateCells"]["range"]["startColumnIndex"], 9)
            if "updateDimensionProperties" in r:
                self.assertEqual(r["updateDimensionProperties"]["fields"], "pixelSize")
            if "repeatCell" in r:
                self.assertNotIn("userEnteredValue", r["repeatCell"]["cell"])


def _rc(r):
    rc = r["repeatCell"]
    return rc["range"], rc["cell"].get("userEnteredFormat", {}), rc["fields"]


class FormatRequestsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.chats = make_chats()
        cls.fakes = {}
        for which in ("rating", "tag"):
            f = FakeSheets()
            run_writer(new, f, which, cls.chats)
            cls.fakes[which] = f

    def each(self):
        for which, f in self.fakes.items():
            with self.subTest(sheet=which):
                yield f, f.sheets[new.RATING_SHEET if which == "rating" else new.TAG_SHEET]

    def test_reset_before_values(self):
        for f, sid in self.each():
            kinds = [k for k, _ in f.events]
            vi = kinds.index("values")
            resets = [i for i, (k, p) in enumerate(f.events) if k == "batch"
                      and any("repeatCell" in r and r["repeatCell"]["fields"] == "userEnteredFormat"
                              for r in p)]
            self.assertEqual(len(resets), 1)
            self.assertLess(kinds.index("clear"), resets[0])
            self.assertLess(resets[0], vi)
            req = f.events[resets[0]][1][0]["repeatCell"]
            # satır 3'ten sayfa sonuna, tüm sütunlar (uçlar açık)
            self.assertEqual(req["range"], {"sheetId": sid, "startRowIndex": 2})
            self.assertEqual(req["cell"], {})

    def test_title_merged_centered(self):
        for f, sid in self.each():
            merges = [r["mergeCells"] for r in f.requests() if "mergeCells" in r]
            self.assertEqual(len(merges), 1)
            self.assertEqual(merges[0]["range"], {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                                                  "startColumnIndex": 0, "endColumnIndex": 10})
            self.assertEqual(f.merges, [merges[0]["range"]])
            title = [_rc(r) for r in f.requests() if "repeatCell" in r
                     and r["repeatCell"]["range"].get("startRowIndex") == 0]
            self.assertEqual(len(title), 1)
            rng_, fmt, _ = title[0]
            self.assertEqual((rng_["startColumnIndex"], rng_["endColumnIndex"]), (0, 10))
            self.assertEqual(fmt["horizontalAlignment"], "CENTER")
            self.assertEqual(fmt["verticalAlignment"], "MIDDLE")
            self.assertEqual(fmt["wrapStrategy"], "CLIP")
            rows = [r["updateDimensionProperties"] for r in f.requests()
                    if "updateDimensionProperties" in r
                    and r["updateDimensionProperties"]["range"]["dimension"] == "ROWS"]
            self.assertEqual(len(rows), 1)
            self.assertEqual((rows[0]["range"]["startIndex"], rows[0]["range"]["endIndex"],
                              rows[0]["properties"]["pixelSize"]), (0, 1, 34))

    def test_column_widths_same_on_both_sheets(self):
        for f, sid in self.each():
            widths = {r["updateDimensionProperties"]["range"]["startIndex"]:
                      r["updateDimensionProperties"]["properties"]["pixelSize"]
                      for r in f.requests() if "updateDimensionProperties" in r
                      and r["updateDimensionProperties"]["range"]["dimension"] == "COLUMNS"}
            self.assertEqual(widths, dict(enumerate([160, 170, 160, 180, 110, 160, 100, 105, 320, 200])))

    def test_data_wrap_align(self):
        n = len(self.chats)
        for f, sid in self.each():
            layout = [_rc(r) for r in f.requests() if "repeatCell" in r
                      and "wrapStrategy" in r["repeatCell"]["cell"].get("userEnteredFormat", {})
                      and r["repeatCell"]["range"].get("startRowIndex") == 2]
            by_cols = {(g["startColumnIndex"], g["endColumnIndex"]): (fmt, fields) for g, fmt, fields in layout}
            for g, _, _ in layout:
                self.assertEqual((g["startRowIndex"], g["endRowIndex"]), (2, 2 + n))
            fmt, fields = by_cols[(0, 10)]
            self.assertEqual(fmt, {"wrapStrategy": "CLIP", "verticalAlignment": "MIDDLE"})
            self.assertNotIn("backgroundColor", fields)   # renkleri ezmez
            fmt, fields = by_cols[(8, 9)]
            self.assertEqual(fmt, {"wrapStrategy": "WRAP", "horizontalAlignment": "LEFT"})
            # WRAP yalnızca User Comment sütununda
            for r in f.requests():
                if "repeatCell" not in r:
                    continue
                g, fmt, _ = _rc(r)
                if fmt.get("wrapStrategy") == "WRAP":
                    self.assertEqual((g["startColumnIndex"], g["endColumnIndex"]), (8, 9))
            rating = [fmt for g, fmt, _ in (_rc(r) for r in f.requests() if "repeatCell" in r)
                      if g.get("startColumnIndex") == 7 and g.get("endRowIndex") == 2 + n
                      and g.get("startRowIndex") == 2]
            self.assertEqual(rating, [{"horizontalAlignment": "CENTER"}])

    def test_site_column_uniform(self):
        n = len(self.chats)
        for f, sid in self.each():
            site = [_rc(r) for r in f.requests() if "repeatCell" in r
                    and r["repeatCell"]["range"].get("startColumnIndex") == 5]
            self.assertEqual(len(site), 1)
            g, fmt, fields = site[0]
            self.assertEqual((g["startRowIndex"], g["endRowIndex"], g["endColumnIndex"]), (2, 2 + n, 6))
            self.assertEqual(fmt["textFormat"], {"foregroundColor": new.rgb("#000000"), "underline": False})
            self.assertIn("userEnteredFormat.textFormat.link", fields)

    def test_zebra_vip_rating_colors_unchanged(self):
        """Zebra/VIP/rating renk istekleri eski kodla aynı (kayma yok)."""
        for which in ("rating", "tag"):
            f_old = FakeSheets()
            run_writer(old, f_old, which, self.chats)

            def colors(f):
                return [r for r in f.requests() if "repeatCell" in r
                        and "backgroundColor" in r["repeatCell"]["fields"]
                        and r["repeatCell"]["range"].get("startRowIndex", 0) >= 2]
            self.assertEqual(colors(self.fakes[which]), colors(f_old))

    def test_chat_link_format_kept(self):
        for f, _ in self.each():
            for _, col, val, tf in f.link_cells():
                self.assertEqual(col, 9)
                self.assertEqual(val, {"stringValue": "Chati Ac"})
                self.assertTrue(tf["underline"])
                self.assertEqual(tf["foregroundColor"], new.rgb("#1155cc"))

    def test_no_customer_data_in_logs(self):
        f = FakeSheets()
        out = run_writer(new, f, "rating", self.chats) + run_writer(new, f, "tag", self.chats)
        for s in SECRET_NAMES + [c.strip() for c in SECRET_COMMENTS]:
            self.assertNotIn(s.strip()[:20], out)


class SiteCheckLogTest(unittest.TestCase):
    def test_logs_only_counts_and_shapes(self):
        cells = [
            {"formattedValue": "https://secret-site.example/p", "hyperlink": "https://secret-site.example/p",
             "effectiveFormat": {"textFormat": {"foregroundColor": {"blue": 0.8}, "underline": True}}},
            {"formattedValue": "www.secret-two.example",
             "effectiveFormat": {"textFormat": {"foregroundColor": {}}}},
            {"formattedValue": "secret-three.example",
             "effectiveFormat": {"textFormat": {"foregroundColor": {}}}},
            {},
        ]

        class F(FakeSheets):
            def get(self, spreadsheetId, ranges=None, fields=None, range=None):
                if ranges is not None:
                    return _Exec(lambda: {"sheets": [{"data": [{"rowData": [
                        {"values": [c]} for c in cells]}]}]})
                return super().get(spreadsheetId)

        with contextlib.redirect_stdout(io.StringIO()) as buf:
            new.report_site_format(F(), new.RATING_SHEET, 4)
        out = buf.getvalue()
        self.assertNotIn("secret", out)
        self.assertIn("bağlantılı hücre: 1", out)
        self.assertIn("'http': 1", out)
        self.assertIn("'www': 1", out)
        self.assertIn("'domain': 1", out)
        self.assertIn("'bos': 1", out)
        self.assertIn("#0000cc/alti-cizili", out)


class IdempotencyTest(unittest.TestCase):
    def test_same_day_rerun(self):
        f = FakeSheets()
        chats = make_chats()
        out1 = run_writer(new, f, "rating", chats)
        out2 = run_writer(new, f, "rating", chats[:10])   # ikinci çalıştırma, daha az satır
        out3 = run_writer(new, f, "rating", chats)
        for out in (out1, out2, out3):
            self.assertIn("Başlık: A:J birleşik", out)
            self.assertNotIn("merge yapılamadı", out)
        self.assertEqual(len(f.merges), 1)
        # her çalıştırmada format sıfırlama tekrarlandı
        resets = [r for r in f.requests() if "repeatCell" in r
                  and r["repeatCell"]["fields"] == "userEnteredFormat"]
        self.assertEqual(len(resets), 3)

    def test_both_sheets_rerun(self):
        f = FakeSheets()
        for _ in range(2):
            run_writer(new, f, "rating", make_chats())
            run_writer(new, f, "tag", make_chats())
        self.assertEqual(len(f.merges), 2)
        self.assertEqual({m["sheetId"] for m in f.merges}, set(f.sheets.values()))

    def test_merge_failure_falls_back(self):
        f = FakeSheets(filter_start_row=0)   # filtre başlık satırını kapsıyor
        out = run_writer(new, f, "rating", make_chats())
        self.assertIn("merge yapılamadı", out)
        self.assertEqual(f.merges, [])
        fb = [_rc(r) for r in f.requests() if "repeatCell" in r
              and r["repeatCell"]["cell"].get("userEnteredFormat", {}).get("wrapStrategy") == "OVERFLOW_CELL"]
        self.assertEqual(len(fb), 1)
        g, fmt, _ = fb[0]
        self.assertEqual((g["startRowIndex"], g["endRowIndex"], g["startColumnIndex"], g["endColumnIndex"]),
                         (0, 1, 0, 1))
        self.assertEqual(fmt["horizontalAlignment"], "LEFT")
        # değerler yine yazıldı ve linkler uygulandı
        self.assertEqual(len(f.value_writes()), 1)
        self.assertTrue(f.link_cells())


if __name__ == "__main__":
    unittest.main()
