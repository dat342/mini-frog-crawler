"""Nghien cuu volume tu khoa: tong theo thang, theo nhom, theo chu de, top tu khoa.

Dung (CLI):
    .venv/bin/python -m xu_ly_data.volume_trend --out kq.xlsx [--rules rules.json] file1.csv [file2.xlsx ...]

Dau vao (tu nhan dang):
 (a) CSV xuat tu Google Keyword Planner: UTF-16 (hoac UTF-8), tab/phay, vai dong tieu de dau;
     cot "Keyword"/"Từ khóa", "Avg. monthly searches"/"Số lượt tìm kiếm trung bình hằng tháng",
     "Searches: Oct 2025"/"Lượt tìm kiếm: thg 10 2025".
 (b) Bang xlsx/csv chung: cot Tu khoa (+ Nhom, Chu de neu co), Trung binh, cac cot thang
     ("Tháng 10/2025", "10/2025", "2025-10", o ngay...).
Nhom/Chu de: lay tu file neu co; khong co (hoac o trong) thi gan bang opts["rules"]:
    {"nhom":[{"name","pattern"}], "chu_de":[{"name","pattern"}], "default":"Khác"}
    regex tren tu khoa bo dau, first match wins.
Gia tri dang khoang ("0-10", "1K – 10K") lay can duoi.

- Sheet an 'Volume data' (doc): Từ khoá | Nhóm | Chủ đề | Volume TB | mm/yyyy x (toi da 12 thang gan nhat).
- Sheet 'Nghiên cứu volume': moi bang la HAM (SUM/SUMIFS) tro vao 'Volume data'.
Khoa (CONTRACT 8.2): bang vol_thang, vol_nhom, vol_chu_de, vol_top_tu_khoa; bieu_do vol_thang (col),
vol_nhom (pie), vol_chu_de (bar); du_lieu vol_data; so vol_tong, vol_so_tu.
Python 3.9: khong dung cu phap X | None.
"""
import argparse
import csv
import datetime
import io
import json
import os
import re
import sys
import unicodedata
from typing import Dict, List, Optional, Tuple

from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from . import common

KEY = "volume_trend"
SHEET = "Nghiên cứu volume"
INPUT_KIND = "keyword_volume"
DESC = "Nghiên cứu volume từ khoá: xu hướng theo tháng, theo nhóm, chủ đề, top từ khoá"

DATA_SHEET = "Volume data"
CHUA_PL = "Chưa phân loại"
SO_THANG = 12
TOP_N = 50
MAX_CHU_DE = 55       # vung tong hop <= 60 dong
NGUONG_TU_LON = 20.0  # % tong volume

_THANG_EN = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], 1)}


# ---------------------------------------------------------------- tien ich
def khong_dau(s):
    # type: (str) -> str
    s = str(s).replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def _chuan(s):
    """Ten cot da chuan hoa: bo dau, lower, gop khoang trang."""
    return re.sub(r"\s+", " ", khong_dau(s or "").strip())


def _nhan(s):
    """Nhan dung lam dieu kien SUMIFS: bo ky tu wildcard va toan tu dau chuoi."""
    s = re.sub(r"[~*?]", "", str(s)).strip()
    return s.lstrip("=<>").strip()


_SO_K = re.compile(r"^([\d.,]+)\s*([kKmM]?)$")


def so_volume(v):
    # type: (object) -> Tuple[float, bool]
    """(gia tri, la_khoang). Chap nhan so, '1,000', '1.000', '0-10', '1K – 10K', '' -> 0."""
    if v is None:
        return 0.0, False
    if isinstance(v, (int, float)):
        return float(v), False
    s = str(v).strip().replace(" ", " ")
    if not s or s in ("-", "—", "–"):
        return 0.0, False
    khoang = False
    parts = re.split(r"\s*[-–—]\s*", s)
    if len(parts) == 2 and parts[0] and parts[1]:
        s, khoang = parts[0], True
    m = _SO_K.match(s.strip())
    if not m:
        return 0.0, False
    num, unit = m.group(1), m.group(2).lower()
    if unit:
        num = num.replace(",", ".")
    elif re.match(r"^\d{1,3}([.,]\d{3})+$", num):
        num = re.sub(r"[.,]", "", num)
    else:
        num = num.replace(",", "")
    try:
        x = float(num)
    except ValueError:
        return 0.0, False
    x *= {"k": 1e3, "m": 1e6}.get(unit, 1)
    return x, khoang


def thang_cua_cot(h):
    # type: (object) -> Optional[Tuple[int, int]]
    """(nam, thang) neu tieu de cot la 1 thang, nguoc lai None."""
    if isinstance(h, (datetime.date, datetime.datetime)):
        return h.year, h.month
    s = _chuan(h)
    if not s:
        return None
    # Keyword Planner EN: "searches: oct 2025"
    m = re.search(r"\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?\s+(\d{4})\b", s)
    if m:
        return int(m.group(2)), _THANG_EN[m.group(1)]
    # Keyword Planner VI: "luot tim kiem: thg 10 2025"
    m = re.search(r"\bthg\.?\s*(\d{1,2})[\s/.-]+(\d{4})\b", s)
    if m:
        return int(m.group(2)), int(m.group(1))
    # "thang 10/2025", "10/2025", "t10/2025"
    m = re.search(r"(?:^|thang\s*|t)(\d{1,2})\s*[/.-]\s*(\d{4})$", s)
    if m and 1 <= int(m.group(1)) <= 12:
        return int(m.group(2)), int(m.group(1))
    m = re.match(r"^(\d{4})[/.-](\d{1,2})(?:[/.-]\d{1,2})?$", s)
    if m and 1 <= int(m.group(2)) <= 12:
        return int(m.group(1)), int(m.group(2))
    return None


_TEN_KW = ("keyword", "keywords", "tu khoa", "search term", "query")
_TEN_NHOM = ("nhom", "nhom tu khoa", "group", "keyword group")
_TEN_CD = ("chu de", "topic", "chu de tu khoa")
_TEN_TB = ("avg. monthly searches", "avg monthly searches", "so luot tim kiem trung binh hang thang",
           "luot tim kiem trung binh hang thang", "trung binh", "volume", "search volume", "volume tb",
           "average monthly searches")


# ---------------------------------------------------------------- doc file
def _doc_csv_rows(path):
    # type: (str) -> Tuple[List[List[str]], str]
    with open(path, "rb") as f:
        raw = f.read()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        txt, enc = raw.decode("utf-16"), "utf-16"
    elif raw[:3] == b"\xef\xbb\xbf":
        txt, enc = raw[3:].decode("utf-8"), "utf-8-sig"
    elif len(raw) > 1 and raw[1:2] == b"\x00":
        txt, enc = raw.decode("utf-16-le"), "utf-16-le"
    else:
        try:
            txt, enc = raw.decode("utf-8"), "utf-8"
        except UnicodeDecodeError:
            txt, enc = raw.decode("cp1258", errors="replace"), "cp1258"
    lines = txt.splitlines()
    sample = "\n".join(lines[:20])
    delim = "\t" if sample.count("\t") >= sample.count(",") and "\t" in sample else ","
    if delim == "," and sample.count(";") > sample.count(","):
        delim = ";"
    return [r for r in csv.reader(io.StringIO(txt), delimiter=delim)], "%s/%s" % (enc, "tab" if delim == "\t" else delim)


def _tim_tieu_de(rows):
    # type: (List[list]) -> int
    for i, r in enumerate(rows[:15]):
        if any(_chuan(c) in _TEN_KW for c in r if c is not None):
            return i
    return -1


def _doc_xlsx_rows(path, sheet=None):
    # type: (str, Optional[str]) -> Tuple[List[list], str]
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True, read_only=True)
    names = [sheet] if sheet else wb.sheetnames
    best = None
    for n in names:
        ws = wb[n]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        h = _tim_tieu_de(rows)
        if h >= 0:
            sc = sum(1 for c in rows[h] if thang_cua_cot(c))
            if best is None or sc > best[0]:
                best = (sc, n, rows)
    wb.close()
    if not best:
        raise ValueError("%s: không thấy sheet nào có cột Từ khoá/Keyword" % path)
    return best[2], "xlsx:%s" % best[1]


def doc_file(path, sheet=None):
    # type: (str, Optional[str]) -> Dict
    """Doc 1 file -> {"fmt", "rows": [{"kw","nhom","chu_de","avg","thang":{(y,m):v}}], "co_nhom","co_cd",
    "so_khoang"}."""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".txt":  # danh sach tu khoa, moi dong 1 tu (dung voi nguon keywordtool)
        with open(path, encoding="utf-8-sig") as f:
            kws = [ln.strip() for ln in f if ln.strip()]
        return {"fmt": "danh sách từ khoá (.txt)", "rows": [{"kw": k, "nhom": "", "chu_de": "", "avg": 0.0, "thang": {}}
                                                          for k in kws],
                "co_nhom": False, "co_cd": False, "so_khoang": 0, "co_tb": False, "so_thang": 0}
    if ext in (".xlsx", ".xlsm"):
        rows, fmt = _doc_xlsx_rows(path, sheet)
    else:
        rows, fmt = _doc_csv_rows(path)
    h = _tim_tieu_de(rows)
    if h < 0:
        raise ValueError("%s: không thấy dòng tiêu đề có cột Keyword/Từ khoá" % path)
    hdr = rows[h]
    norm = [_chuan(c) if c is not None else "" for c in hdr]

    def col(names):
        for n in names:
            if n in norm:
                return norm.index(n)
        return None

    c_kw, c_nhom, c_cd, c_tb = col(_TEN_KW), col(_TEN_NHOM), col(_TEN_CD), col(_TEN_TB)
    c_thang = [(j, thang_cua_cot(c)) for j, c in enumerate(hdr) if thang_cua_cot(c)]
    kp = any(n.startswith("searches:") or n.startswith("luot tim kiem:") for n in norm) or \
        "avg. monthly searches" in norm or "so luot tim kiem trung binh hang thang" in norm
    if kp:
        fmt = "Keyword Planner (%s)" % fmt
    out, so_khoang = [], 0
    for r in rows[h + 1:]:
        if c_kw is None or c_kw >= len(r):
            continue
        kw = r[c_kw]
        kw = str(kw).strip() if kw is not None else ""
        if not kw:
            continue
        g = lambda j: r[j] if (j is not None and j < len(r)) else None
        thang = {}
        for j, ym in c_thang:
            v, kh = so_volume(g(j))
            so_khoang += kh
            thang[ym] = v
        if c_tb is not None:
            avg, kh = so_volume(g(c_tb))
            so_khoang += kh
        else:
            vs = list(thang.values())
            avg = sum(vs) / len(vs) if vs else 0.0
        out.append({"kw": kw, "nhom": (str(g(c_nhom)).strip() if g(c_nhom) not in (None, "") else ""),
                    "chu_de": (str(g(c_cd)).strip() if g(c_cd) not in (None, "") else ""),
                    "avg": avg, "thang": thang})
    return {"fmt": fmt, "rows": out, "co_nhom": c_nhom is not None, "co_cd": c_cd is not None,
            "so_khoang": so_khoang, "co_tb": c_tb is not None, "so_thang": len(c_thang)}


# ---------------------------------------------------------------- luat
def doc_luat(rules):
    # type: (object) -> Optional[Dict]
    if not rules:
        return None
    if isinstance(rules, str):
        with open(rules, encoding="utf-8") as f:
            rules = json.load(f)
    out = {"default": rules.get("default") or "Khác"}
    for k in ("nhom", "chu_de"):
        out[k] = [(t["name"], re.compile(t["pattern"], re.I)) for t in rules.get(k, [])]
    return out


def _gan(kw, luat, k):
    # type: (str, Optional[Dict], str) -> str
    if not luat or not luat.get(k):
        return CHUA_PL if not luat else luat["default"]
    s = khong_dau(kw)
    for name, rx in luat[k]:
        if rx.search(s):
            return name
    return luat["default"]


# ---------------------------------------------------------------- nguon KeywordTool.io
KT_URL = "https://api.keywordtool.io/v2/search/volume/google"
KT_LO = 500          # so tu khoa moi luot goi
KT_CACH_GIAY = 6.5   # gioi han 10 luot/phut


def doc_key_keywordtool():
    k = os.environ.get("KEYWORDTOOL_API_KEY", "").strip()
    if k:
        return k
    p = os.path.expanduser("~/.config/seongon/keywordtool.key")
    if os.path.exists(p):
        with open(p) as f:
            return f.read().strip()
    raise ValueError("Chưa có key KeywordTool: đặt KEYWORDTOOL_API_KEY hoặc tạo ~/.config/seongon/keywordtool.key (chmod 600)")


def lay_volume_keywordtool(rows, opts, ctx):
    # type: (List[Dict], Dict, Dict) -> Tuple[int, int]
    """Dien avg + 12 thang cho rows tu KeywordTool.io (du lieu Google Ads). Cache theo tu khoa/vung/ngon ngu trong
    ctx['cache_dir']/keywordtool.json: dung lai file khong ton luot goi. Tra ve (so luot goi, so tu tu cache)."""
    import time
    import urllib.request
    loc = int(opts.get("location", 2704))
    lang = str(opts.get("language", "vi"))
    cache_dir = ctx.get("cache_dir") or ""
    cpath = os.path.join(cache_dir, "keywordtool.json") if cache_dir else ""
    cache = {}
    if cpath and os.path.exists(cpath):
        with open(cpath, encoding="utf-8") as f:
            cache = json.load(f)
    ck = lambda kw: "%s|%d|%s" % (kw.lower(), loc, lang)
    can = sorted({r["kw"] for r in rows if ck(r["kw"]) not in cache})
    luot = 0
    if can:
        key = doc_key_keywordtool()
        for i in range(0, len(can), KT_LO):
            if luot:
                time.sleep(KT_CACH_GIAY)
            lo = can[i:i + KT_LO]
            payload = {"metrics_network": "googlesearchnetwork", "metrics_currency": "VND", "complete": True,
                       "output": "json", "apikey": key, "metrics_location": [loc], "metrics_language": [lang],
                       "keyword": lo}
            req = urllib.request.Request(KT_URL, data=json.dumps(payload).encode("utf-8"),
                                         headers={"content-type": "application/json"})
            with urllib.request.urlopen(req, timeout=180) as f:
                res = (json.loads(f.read().decode("utf-8")) or {}).get("results") or {}
            luot += 1
            for kw in lo:
                v = res.get(kw) or {}
                cache[ck(kw)] = {"volume": v.get("volume") or 0,
                                 "thang": [[v.get("m%d_year" % j), v.get("m%d_month" % j), v.get("m%d" % j) or 0]
                                           for j in range(1, 13) if v.get("m%d_year" % j)]}
            if cpath:
                os.makedirs(cache_dir, exist_ok=True)
                with open(cpath, "w", encoding="utf-8") as f:
                    json.dump(cache, f, ensure_ascii=False)
    tu_cache = len(rows) - len(can)
    for r in rows:
        v = cache.get(ck(r["kw"])) or {}
        r["avg"] = float(v.get("volume") or 0)
        r["thang"] = {(int(y), int(m)): float(x) for y, m, x in v.get("thang", [])}
    return luot, tu_cache


# ---------------------------------------------------------------- build
def build(wb, inputs, opts, ctx):
    opts = opts or {}
    ctx = ctx or {}
    luat = doc_luat(opts.get("rules"))
    mac_dinh = (luat or {}).get("default", "Khác")
    files = [doc_file(p, opts.get("sheet")) for p in inputs]
    nguon = str(opts.get("nguon") or "").lower()
    kt_info = None
    if nguon == "keywordtool":
        # input chi can danh sach tu khoa (+ Nhom/Chu de neu co); volume lay tu KeywordTool.io
        kt_info = lay_volume_keywordtool([r for f in files for r in f["rows"]], opts, ctx)
        for f in files:
            f["fmt"] = "KeywordTool.io API (Google Ads) ← " + f["fmt"]

    # gop, bo trung tu khoa (giu lan dau)
    seen, rows, trung, so_khoang = set(), [], 0, 0
    for f in files:
        so_khoang += f["so_khoang"]
        for r in f["rows"]:
            k = khong_dau(r["kw"]).strip()
            if k in seen:
                trung += 1
                continue
            seen.add(k)
            r["nhom"] = _nhan(r["nhom"] or _gan(r["kw"], luat, "nhom")) or mac_dinh
            r["chu_de"] = _nhan(r["chu_de"] or _gan(r["kw"], luat, "chu_de")) or mac_dinh
            rows.append(r)
    all_m = sorted({ym for r in rows for ym in r["thang"]})
    months = all_m[-SO_THANG:]
    rows.sort(key=lambda r: (-r["avg"], r["kw"]))
    n = len(rows)
    last = max(n + 1, 2)

    # ---------- Volume data (an) ----------
    dws = wb.create_sheet(DATA_SHEET)
    hdr = ["Từ khoá", "Nhóm", "Chủ đề", "Volume TB"] + ["%02d/%d" % (m, y) for y, m in months]
    common.header_row(dws, 1, 1, hdr)
    for r in rows:
        dws.append([r["kw"], r["nhom"], r["chu_de"], r["avg"]] + [r["thang"].get(ym, 0) for ym in months])
    for i, w in enumerate([40, 20, 28, 12] + [10] * len(months), 1):
        dws.column_dimensions[get_column_letter(i)].width = w
    for j in range(4, 5 + len(months)):
        for cell in dws[get_column_letter(j)][1:]:
            cell.number_format = "#,##0"
    dws.freeze_panes = "B2"
    dws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(hdr)), last)
    dws.sheet_state = "hidden"
    D = lambda c: "'%s'!$%s$2:$%s$%d" % (DATA_SHEET, c, c, last)

    # ---------- thong ke Python ----------
    tong = sum(r["avg"] for r in rows)
    v_nhom, v_cd = {}, {}  # type: Dict[str, float], Dict[str, float]
    for r in rows:
        v_nhom[r["nhom"]] = v_nhom.get(r["nhom"], 0) + r["avg"]
        v_cd[r["chu_de"]] = v_cd.get(r["chu_de"], 0) + r["avg"]
    nhoms = sorted(v_nhom, key=lambda t: (-v_nhom[t], t))[:MAX_CHU_DE]
    cds = sorted(v_cd, key=lambda t: (-v_cd[t], t))
    cds_hien = cds[:MAX_CHU_DE]
    v_thang = [sum(r["thang"].get(ym, 0) for r in rows) for ym in months]

    # ---------- sheet ket qua ----------
    ws = wb.create_sheet(SHEET)
    q = common.quote_sheet(SHEET)
    common.title(ws, "NGHIÊN CỨU VOLUME TỪ KHOÁ",
                 "Nguồn: %s. Bảng dưới là hàm SUM/SUMIFS trỏ vào sheet ẩn '%s'."
                 % ("; ".join(sorted({f["fmt"] for f in files})) or "-", DATA_SHEET))
    ws["A4"] = "Tổng volume TB/tháng"
    ws["B4"] = "=SUM(%s)" % D("D")
    ws["A5"] = "Số từ khoá"
    ws["B5"] = "=COUNTA(%s)" % D("A")
    for c in ("A4", "A5"):
        ws[c].font = Font(bold=True)
    for c in ("B4", "B5"):
        ws[c].number_format = "#,##0"
        ws[c].font = Font(bold=True, color=common.XANH)
        ws[c].border = common.BOX
    sums, charts = [], []

    def nhan_bang(row, text):
        ws.cell(row, 1, text).font = Font(bold=True, color=common.XANH)

    def reg(key, ref, desc):
        e = common.summary(wb, ws, key, ref, desc)
        e["key"] = key
        sums.append(e)

    def o(r, c, v, fmt="#,##0"):
        cell = ws.cell(r, c, v)
        cell.border = common.BOX
        if fmt:
            cell.number_format = fmt
        return cell

    # (1) vol_thang
    row = 7
    nhan_bang(row, "(1) Tổng volume theo tháng (%d tháng gần nhất)" % len(months))
    common.header_row(ws, row + 1, 1, ["Tháng", "Tổng volume"])
    r0 = row + 2
    for i, (y, m) in enumerate(months):
        o(r0 + i, 1, "%02d/%d" % (m, y), None)
        cl = get_column_letter(5 + i)
        o(r0 + i, 2, "=SUM(%s)" % D(cl))
    r_end_thang = r0 + max(len(months), 1) - 1
    reg("vol_thang", "A%d:B%d" % (row + 1, r_end_thang), "Tổng volume tìm kiếm theo tháng (12 tháng gần nhất)")
    blk_thang = (row + 1, r_end_thang)

    # (2) vol_nhom
    row = r_end_thang + 3
    nhan_bang(row, "(2) Volume TB theo nhóm")
    common.header_row(ws, row + 1, 1, ["Nhóm", "Volume TB", "Tỷ trọng"])
    r0 = row + 2
    rT = r0 + len(nhoms)
    for i, t in enumerate(nhoms):
        r = r0 + i
        o(r, 1, t, None)
        o(r, 2, "=SUMIFS(%s,%s,$A%d)" % (D("D"), D("B"), r))
        o(r, 3, "=IFERROR(B%d/$B$4,0)" % r, "0.0%")
    ws.cell(rT, 1, "Tổng").font = Font(bold=True)
    o(rT, 2, "=SUM(B%d:B%d)" % (r0, rT - 1)).font = Font(bold=True)
    o(rT, 3, "=SUM(C%d:C%d)" % (r0, rT - 1), "0.0%").font = Font(bold=True)
    reg("vol_nhom", "A%d:C%d" % (row + 1, rT - 1), "Volume TB theo nhóm và tỷ trọng % trên tổng (dòng Tổng nằm ngay dưới, ngoài vùng)")
    blk_nhom = (row + 1, rT - 1)

    # (3) vol_chu_de
    row = rT + 3
    nhan_bang(row, "(3) Volume TB theo chủ đề (giảm dần)")
    common.header_row(ws, row + 1, 1, ["Chủ đề", "Volume TB", "Tỷ trọng"])
    r0 = row + 2
    for i, t in enumerate(cds_hien):
        r = r0 + i
        o(r, 1, t, None)
        o(r, 2, "=SUMIFS(%s,%s,$A%d)" % (D("D"), D("C"), r))
        o(r, 3, "=IFERROR(B%d/$B$4,0)" % r, "0.0%")
    r_end_cd = r0 + max(len(cds_hien), 1) - 1
    reg("vol_chu_de", "A%d:C%d" % (row + 1, r_end_cd), "Volume TB theo chủ đề, sắp giảm dần")
    blk_cd = (row + 1, r_end_cd)

    # nhan xet
    row = r_end_cd + 3
    notes = [common.note_area(ws, "A%d" % (row + 1), 4, 8, "NHẬN XÉT")]

    # (4) vol_top_tu_khoa
    row += 11
    top = rows[:TOP_N]
    nhan_bang(row, "(4) Top %d từ khoá theo volume TB" % len(top))
    common.header_row(ws, row + 1, 1, ["Từ khoá", "Nhóm", "Chủ đề", "Volume TB"])
    r0 = row + 2
    for i, r in enumerate(top):
        o(r0 + i, 1, r["kw"], None)
        o(r0 + i, 2, r["nhom"], None)
        o(r0 + i, 3, r["chu_de"], None)
        o(r0 + i, 4, "='%s'!D%d" % (DATA_SHEET, i + 2))
    reg("vol_top_tu_khoa", "A%d:D%d" % (row + 1, r0 + max(len(top), 1) - 1),
        "Top %d từ khoá theo volume TB (từ khoá/nhóm/chủ đề là giá trị, volume tham chiếu '%s'!D)" % (len(top), DATA_SHEET))

    ws.column_dimensions["A"].width = 40
    for c, w in (("B", 18), ("C", 24), ("D", 14)):
        ws.column_dimensions[c].width = w

    # ---------- bieu do ----------
    def ve(key, loai, blk, anchor, title, cot=None):
        h0, h1 = blk
        if loai == "pie":
            ch = PieChart()
            ch.dataLabels = DataLabelList()
            ch.dataLabels.showPercent = True
        else:
            ch = BarChart()
            ch.type = loai
            ch.legend = None
        ch.title = title
        ch.height, ch.width = (9, 16) if loai != "bar" else (max(9, 0.5 * (h1 - h0) + 3), 16)
        ch.add_data(Reference(ws, min_col=2, min_row=h0, max_row=h1), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=h0 + 1, max_row=h1))
        if loai == "bar":
            ch.y_axis.scaling.orientation = "minMax"
            ch.x_axis.scaling.orientation = "maxMin"
        ws.add_chart(ch, anchor)
        e_ch = {"key": key, "loai": loai, "bang": key, "anchor": "%s!%s" % (q, anchor)}
        if cot:
            e_ch["cot"] = cot
        charts.append(e_ch)

    ve("vol_thang", "col", blk_thang, "F4", "Tổng volume theo tháng")
    ve("vol_nhom", "pie", blk_nhom, "F23", "Tỷ trọng volume theo nhóm")
    ve("vol_chu_de", "bar", blk_cd, "F42", "Volume TB theo chủ đề", cot="Volume TB")

    # ---------- caveats ----------
    caveats = ["%s từ khoá, tổng volume TB %s/tháng; %d tháng dữ liệu (%s → %s)." % (
        format(n, ","), format(int(round(tong)), ","), len(months),
        "%02d/%d" % (months[0][1], months[0][0]) if months else "-",
        "%02d/%d" % (months[-1][1], months[-1][0]) if months else "-")]
    if kt_info is not None:
        caveats.append("Nguồn volume: KeywordTool.io (dữ liệu Google Ads/Keyword Planner, số làm tròn theo thang Google); "
                       "vùng %s, ngôn ngữ %s; %d lượt gọi API, %d từ lấy từ cache." % (
                           opts.get("location", 2704), opts.get("language", "vi"), kt_info[0], kt_info[1]))
    if trung:
        caveats.append("Bỏ %d từ khoá trùng (giữ lần xuất hiện đầu)." % trung)
    if tong > 0:
        caveats.append("Nhóm '%s' chiếm %.1f%% tổng volume." % (mac_dinh, v_nhom.get(mac_dinh, 0) / tong * 100))
        for lab in sorted({mac_dinh, CHUA_PL}):
            if lab in v_nhom and lab != mac_dinh:
                caveats.append("Nhóm '%s' chiếm %.1f%% tổng volume." % (lab, v_nhom[lab] / tong * 100))
            if lab in v_cd:
                caveats.append("Chủ đề '%s' chiếm %.1f%% tổng volume." % (lab, v_cd[lab] / tong * 100))
        lon = [r for r in rows if r["avg"] / tong * 100 > NGUONG_TU_LON]
        for r in lon:
            caveats.append("Từ khoá '%s' chiếm %.1f%% tổng volume (> %.0f%%), xu hướng tổng bị nó chi phối."
                           % (r["kw"], r["avg"] / tong * 100, NGUONG_TU_LON))
    if months and any(v_thang):
        imax = max(range(len(months)), key=lambda i: v_thang[i])
        imin = min(range(len(months)), key=lambda i: v_thang[i])
        tb = sum(v_thang) / len(v_thang)
        caveats.append("Tháng cao nhất %02d/%d: %s (%.1fx TB tháng); thấp nhất %02d/%d: %s." % (
            months[imax][1], months[imax][0], format(int(v_thang[imax]), ","), v_thang[imax] / tb if tb else 0,
            months[imin][1], months[imin][0], format(int(v_thang[imin]), ",")))
        tb_cot = sum(v_thang) / len(v_thang)
        if tong > 0 and abs(tb_cot - tong) / tong > 0.15:
            caveats.append("Volume TB (cột nguồn) = %s khác trung bình 12 tháng = %s (lệch %.0f%%) — "
                           "Volume TB do nguồn tính trên khoảng khác/làm tròn." % (
                               format(int(tong), ","), format(int(tb_cot), ","), (tb_cot - tong) / tong * 100))
    elif not months:
        caveats.append("File không có cột theo tháng: bảng vol_thang trống.")
    if so_khoang:
        caveats.append("%d ô volume dạng khoảng (vd '0-10'), đã lấy cận dưới." % so_khoang)
    if len(cds) > MAX_CHU_DE:
        caveats.append("Có %d chủ đề, bảng chỉ hiện %d chủ đề đầu." % (len(cds), MAX_CHU_DE))
    if not luat and not all(f["co_nhom"] and f["co_cd"] for f in files):
        caveats.append("File thiếu cột Nhóm/Chủ đề và chưa có opts.rules: phần thiếu = '%s'." % CHUA_PL)

    columns = [{"name": "Từ khoá", "desc": "Từ khoá (đã bỏ trùng)"},
               {"name": "Nhóm", "desc": "Nhóm từ khoá (từ file hoặc luật regex)"},
               {"name": "Chủ đề", "desc": "Chủ đề từ khoá (từ file hoặc luật regex)"},
               {"name": "Volume TB", "desc": "Lượt tìm kiếm trung bình/tháng (cột nguồn; không có thì TB các tháng)"}] + \
              [{"name": "%02d/%d" % (m, y), "desc": "Lượt tìm kiếm tháng %02d/%d" % (m, y)} for y, m in months]
    so = {"vol_tong": {"site": None, "cell": "%s!B4" % q},
          "vol_so_tu": {"site": None, "cell": "%s!B5" % q}}
    return common.entry(
        SHEET, KEY, "Volume tìm kiếm của bộ từ khoá: xu hướng theo tháng, tỷ trọng theo nhóm/chủ đề, top từ khoá",
        inputs, data_sheet=DATA_SHEET, rows=n, columns=columns, summaries=sums,
        options=[], caveats=caveats, notes=notes,
        extra={"so": so, "charts": charts, "du_lieu": {"key": "vol_data", "sheet": DATA_SHEET},
               "formats": [f["fmt"] for f in files]})


# ---------------------------------------------------------------- CLI
def main(argv=None):
    ap = argparse.ArgumentParser(description=DESC)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--rules")
    ap.add_argument("--sheet")
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    opts = {}
    if a.rules:
        opts["rules"] = a.rules
    if a.sheet:
        opts["sheet"] = a.sheet
    e = build(wb, a.files, opts, {})
    wb.save(a.out)
    print(json.dumps(e, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
