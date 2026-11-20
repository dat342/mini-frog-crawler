"""Engine "mau gan the" (CONTRACT muc 8): do ket qua cac module vao 1 file mau .xlsx co the {{...}}.

    .venv/bin/python -m xu_ly_data.the_mau quet <mau.xlsx>                 # liet ke the + kiem khoa
    .venv/bin/python -m xu_ly_data.the_mau dien --built tam.xlsx --manifest tam.xlsx.manifest.json \
        --mau xu_ly_data/mau/bd-seo.xlsx --out "<Khach> - BD SEO.xlsx" [--recipe r.json] [--notes n.json]
    .venv/bin/python -m xu_ly_data.the_mau chuyen-notes --tu "<file cu>.xlsx" --sang "<file moi>.xlsx"

Python: dien_mau(built_xlsx, manifest, mau_path, out_path, recipe, notes=None) -> manifest moi.

Nguyen tac:
- openpyxl chi DOC (mau + file built). Moi thay doi tren mau di qua GridCraft CLI (giu bieu do/anh/logo cua mau).
- Sheet du lieu an cua module (vd 'RD data') duoc chep nguyen ten sang file ra (an) -> cong thuc tro vao do giu nguyen.
- Cong thuc trong bang tro toi o cua bang khac cung sheet ket qua cua module -> engine doi dia chi sang vi tri bang do
  trong file ra; bang (hoac o le) duoc tham chieu ma mau khong dat -> engine dat ban an o sheet '_bang'.
- Bang dat tai the: dong tieu de lay dinh dang o the (tung cot: o cung dong ben phai neu co dinh dang, khong thi o gan
  nhat ben trai); dong than lay dinh dang dong ngay duoi the (cung quy tac). Dinh dang so lay tu module.
- Bang dai hon khoang trong ben duoi (toi o co noi dung/the tiep theo trong cung cot) -> chen dong.
Python 3.9: khong dung cu phap X | None.
"""
import argparse
import datetime as dt
import json
import os
import re
import subprocess
import sys
import tempfile
import unicodedata
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

GRIDCRAFT = os.environ.get(
    "GRIDCRAFT_CLI",
    os.path.expanduser("~/.local/gridcraft/gridcraft-cli-0.19.1-macos-universal/gridcraft-cli"))
MAU_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mau")
BANG_SHEET = "_bang"
MANIFEST_SHEET = "_manifest"
KHUNG = "Dữ liệu nghiên cứu:\n\nĐánh giá:\n\nĐịnh hướng:"
PALETTE = ["#004AEF", "#12D6F7", "#FFCE04", "#1F2329", "#7FA6F7", "#C00000", "#BFBFBF", "#00B050"]
MAX_SUM_ROWS = 55

# ---------------------------------------------------------------- danh muc khoa (CONTRACT 8.2 + bo sung)
CATALOG = {
    "backlink_rd": {"bang": ["bl_theo_nam", "bl_chat_luong", "bl_chat_luong_pct", "bl_site", "bl_tuy_chon"],
                    "bieu_do": ["bl_theo_nam", "bl_chat_luong", "bl_site"], "du_lieu": ["bl_data"],
                    "so": ["bl_tong", "bl_dr70", "bl_spam_pct"]},
    "traffic_ahrefs": {"bang": ["tr_thang", "tr_tong_hop"], "bieu_do": ["tr_thang"], "du_lieu": ["tr_data"],
                       "so": ["tr_tb12", "tr_tang_truong", "tr_dinh"]},
    "content_topic": {"bang": ["ct_tong", "ct_topic_sotrang", "ct_topic_traffic", "ct_loaitrang_sotrang",
                               "ct_loaitrang_traffic", "ct_loaitrang_pct"],
                      "bieu_do": ["ct_tong", "ct_loaitrang", "ct_topic"], "du_lieu": ["ct_data"],
                      "so": ["ct_so_trang", "ct_traffic"]},
    "onpage": {"bang": ["op_diem", "op_tieu_chi", "op_da_cham"], "bieu_do": ["op_diem"], "du_lieu": ["op_data"],
               "so": ["op_diem_tong"]},
    "volume_trend": {"bang": ["vol_thang", "vol_nhom", "vol_chu_de", "vol_top_tu_khoa"],
                     "bieu_do": ["vol_thang", "vol_nhom", "vol_chu_de"], "du_lieu": ["vol_data"],
                     "so": ["vol_tong", "vol_so_tu"]},
    "batch_overview": {"bang": ["bo_tong_quan"], "bieu_do": ["bo_traffic"], "du_lieu": ["bo_data"],
                       "so": ["bo_traffic", "bo_dr", "bo_keywords", "bo_rd", "bo_ai_responses"]},
}
DATA_TYPES = ("bang", "bieu_do", "du_lieu", "so")
PARAMS = {
    "bang": {"site", "top", "cot", "bo_tieu_de"},
    "bieu_do": {"loai", "site", "buoc", "rong", "cao", "tieu_de", "dong", "cot", "chuoi", "mau"},
    "du_lieu": {"kieu", "top"},
    "nhan_xet": {"khung"},
    "so": {"site"},
    "site": set(), "ten": set(), "khach": set(), "nam": set(), "ngay": set(),
}
TEXT_TYPES = ("khach", "nam", "ngay", "site", "ten", "so")
LOAI_CHART = {"bar": ("bar", "clustered"), "bar_stacked": ("bar", "stacked"), "col": ("column", "clustered"),
              "col_stacked": ("column", "stacked"), "col_stacked100": ("column", "stacked100"),
              "line": ("line", None), "line_markers": ("line", "markers"), "pie": ("pie", None),
              "doughnut": ("doughnut", None), "radar": ("radar", None), "area": ("area", None),
              "area_stacked": ("area", "stacked")}


class LoiThe(Exception):
    pass


def L(c):
    return get_column_letter(c)


def q(sheet):
    return "'%s'" % sheet.replace("'", "''")


def module_cua_khoa(loai, khoa):
    for m, d in CATALOG.items():
        if khoa in d.get(loai, []) or (loai == "bieu_do" and khoa in d["bang"]):
            return m
    return None


def _slug(s):
    s = (s or "").replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]+", "_", s).strip("_")


# ---------------------------------------------------------------- doc the
def tach_the(text):
    """Tra ve list (start, end, inner) cac the {{...}} cap ngoai cung (cho phep long nhau)."""
    out, i, n = [], 0, len(text)
    while i < n:
        a = text.find("{{", i)
        if a < 0:
            break
        depth, j = 0, a
        while j < n:
            if text.startswith("{{", j):
                depth += 1
                j += 2
            elif text.startswith("}}", j):
                depth -= 1
                j += 2
                if depth == 0:
                    break
            else:
                j += 1
        if depth != 0:
            raise LoiThe("thiếu '}}' đóng thẻ: %s" % text[a:a + 60])
        out.append((a, j, text[a + 2:j - 2]))
        i = j
    return out


def phan_tich(inner):
    """'bang:bl_site|site=1' -> (loai, khoa, params)."""
    parts, depth, cur = [], 0, ""
    k = 0
    while k < len(inner):
        if inner.startswith("{{", k):
            depth += 1
            cur += "{{"
            k += 2
            continue
        if inner.startswith("}}", k):
            depth -= 1
            cur += "}}"
            k += 2
            continue
        ch = inner[k]
        if ch == "|" and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
        k += 1
    parts.append(cur)
    head = parts[0].strip()
    loai, _, khoa = head.partition(":")
    params = {}
    for p in parts[1:]:
        if "=" not in p:
            raise LoiThe("tham số '%s' thiếu dấu '='" % p)
        a, b = p.split("=", 1)
        params[a.strip()] = b.strip().replace("\\n", "\n")
    return loai.strip(), khoa.strip(), params


class The(object):
    def __init__(self, sheet, row, col, loai, khoa, params, raw, text):
        self.sheet, self.row, self.col = sheet, row, col
        self.loai, self.khoa, self.params, self.raw, self.text = loai, khoa, params, raw, text

    @property
    def cell(self):
        return "%s%d" % (L(self.col), self.row)

    def __repr__(self):
        return "%s!%s %s" % (q(self.sheet), self.cell, self.raw)


def doc_mau(mau_path, loi=None):
    """Doc the cua moi sheet. Tra ve (tags, info). info[sheet] = {max_row, widths, styled, values, merges, objs}.
    Nem LoiThe neu the sai cu phap."""
    wb = load_workbook(mau_path)
    tags, info, errs = [], {}, []
    for ws in wb.worksheets:
        vals, styled = {}, set()
        for row in ws.iter_rows():
            for c in row:
                if c.has_style and (c.fill.fill_type or c.border.left.style or c.border.top.style or c.font.b
                                    or (c.font.color is not None and c.font.color.rgb not in (None, "FF000000"))):
                    styled.add((c.row, c.column))
                if c.value is None:
                    continue
                vals[(c.row, c.column)] = c.value
                if not isinstance(c.value, str) or "{{" not in c.value:
                    continue
                try:
                    found = tach_the(c.value)
                except LoiThe as e:
                    errs.append("%s!%s: %s" % (q(ws.title), c.coordinate, e))
                    continue
                whole = len(found) == 1 and found[0][0] == 0 and found[0][1] == len(c.value.strip()) \
                    and c.value.strip() == c.value[:found[0][1]]
                for a, b, inner in found:
                    try:
                        loai, khoa, params = phan_tich(inner)
                    except LoiThe as e:
                        errs.append("%s!%s: %s (%s)" % (q(ws.title), c.coordinate, e, c.value[a:b]))
                        continue
                    if loai in ("bang", "bieu_do", "du_lieu", "nhan_xet") and not whole:
                        errs.append("%s!%s: thẻ %s phải nằm một mình trong ô: %s" % (q(ws.title), c.coordinate, loai,
                                                                                  c.value[a:b]))
                        continue
                    tags.append(The(ws.title, c.row, c.column, loai, khoa, params, c.value[a:b], c.value))
        widths = {}
        for k, d in ws.column_dimensions.items():
            if d.width:
                try:
                    lo, hi = column_index_from_string(d.min and L(d.min) or k), column_index_from_string(d.max and L(d.max) or k)
                except ValueError:
                    continue
                for ci in range(lo, hi + 1):
                    widths[ci] = d.width
        objs = []
        for ch in list(ws._charts) + list(ws._images):
            try:
                a = ch.anchor._from
                objs.append((a.row + 1, a.col + 1))
            except AttributeError:
                pass
        info[ws.title] = {"max_row": ws.max_row, "widths": widths, "styled": styled, "values": vals,
                          "merges": [(m.min_row, m.min_col, m.max_row, m.max_col) for m in ws.merged_cells.ranges],
                          "objs": objs, "state": ws.sheet_state}
    sheets = wb.sheetnames
    wb.close()
    if errs:
        if loi is not None:
            loi.extend(errs)
        else:
            raise LoiThe("Thẻ sai cú pháp:\n  " + "\n  ".join(errs))
    return tags, info, sheets


def kiem_the(tags, so_site_max=None):
    """Kiem khoa/tham so. Tra ve list loi (chuoi)."""
    errs = []
    for t in tags:
        if t.loai not in PARAMS:
            errs.append("%r: loại thẻ không hỗ trợ '%s'" % (t, t.loai))
            continue
        bad = set(t.params) - PARAMS[t.loai]
        if bad:
            errs.append("%r: tham số không hỗ trợ %s" % (t, ", ".join(sorted(bad))))
        if t.loai in DATA_TYPES:
            if not t.khoa:
                errs.append("%r: thiếu khóa" % t)
            elif module_cua_khoa(t.loai, t.khoa) is None:
                errs.append("%r: khóa '%s' không có trong danh mục %s (CONTRACT 8.2)" % (t, t.khoa, t.loai))
        if t.loai in ("site", "ten"):
            if not re.match(r"^\d+$", t.khoa or ""):
                errs.append("%r: cần số thứ tự site, vd {{site:1}}" % t)
            elif so_site_max and int(t.khoa) > so_site_max:
                errs.append("%r: site %s vượt số site tối đa %d của mẫu" % (t, t.khoa, so_site_max))
        if t.loai == "nhan_xet" and not t.khoa:
            errs.append("%r: thiếu khóa nhận xét" % t)
        s = t.params.get("site")
        if s is not None and not (re.match(r"^\d+$", s) or (s == "*" and t.loai == "bieu_do")):
            errs.append("%r: site=%s không hợp lệ (số thứ tự, hoặc * với biểu đồ)" % (t, s))
        if s == "*" and not re.match(r"^\d+$", t.params.get("buoc", "")):
            errs.append("%r: site=* cần buoc=<số cột>" % t)
        if t.loai == "bieu_do" and t.params.get("loai") and t.params["loai"] not in LOAI_CHART:
            errs.append("%r: loai=%s không hỗ trợ (%s)" % (t, t.params["loai"], ", ".join(sorted(LOAI_CHART))))
        if t.loai == "du_lieu" and t.params.get("kieu", "canh_nhau") not in ("canh_nhau", "doc"):
            errs.append("%r: kieu phải là canh_nhau hoặc doc" % t)
        for k in ("top", "buoc", "rong", "cao"):
            if k in t.params and not re.match(r"^\d+$", t.params[k]):
                errs.append("%r: %s phải là số" % (t, k))
        if "dong" in t.params and not re.match(r"^\d+-\d+$", t.params["dong"]):
            errs.append("%r: dong phải dạng a-b" % t)
    return errs


# ---------------------------------------------------------------- doc file built
class Bang(object):
    def __init__(self, key, site, module, sheet, r1, c1, r2, c2, total_rows=0, desc="", name="", loose=False):
        self.key, self.site, self.module, self.sheet = key, site, module, sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
        self.total_rows, self.desc, self.name, self.loose = total_rows, desc, name, loose

    def contains(self, sheet, r, c):
        return sheet == self.sheet and self.r1 <= r <= self.r2 and self.c1 <= c <= self.c2


def _range(ref):
    m = re.match(r"^(?:'((?:[^']|'')+)'|([^!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$", ref.strip())
    if not m:
        raise ValueError("Tham chiếu không hợp lệ: %s" % ref)
    sheet = (m.group(1) or "").replace("''", "'") or m.group(2)
    c1, r1 = column_index_from_string(m.group(3)), int(m.group(4))
    c2 = column_index_from_string(m.group(5)) if m.group(5) else c1
    r2 = int(m.group(6)) if m.group(6) else r1
    return sheet, r1, c1, r2, c2


def norm(ref):
    try:
        s, r1, c1, r2, c2 = _range(ref)
    except ValueError:
        return ref.strip()
    if (r1, c1) == (r2, c2):
        return "%s!%s%d" % (q(s), L(c1), r1)
    return "%s!%s%d:%s%d" % (q(s), L(c1), r1, L(c2), r2)


def _list(x):
    if x is None:
        return []
    if isinstance(x, (list, tuple)):
        return list(x)
    return [x]


class Built(object):
    def __init__(self, path, manifest):
        self.wb = load_workbook(path)
        self.manifest = manifest
        self.entries = {e["module"]: e for e in manifest.get("sheets", []) if e.get("module")}
        self.sites = list(manifest.get("sites") or [])
        self.result_sheets = {e["sheet"]: e["module"] for e in manifest.get("sheets", [])}
        self.tables = []
        for e in manifest.get("sheets", []):
            for s in e.get("summaries", []):
                if not s.get("key"):
                    continue
                sh, r1, c1, r2, c2 = _range(s["range"])
                self.tables.append(Bang(s["key"], s.get("site"), e["module"], sh, r1, c1, r2, c2,
                                        int(s.get("total_rows") or 0), s.get("desc", ""), s.get("name", "")))
        self.data_sheets = []
        for e in manifest.get("sheets", []):
            for d in [e.get("data_sheet")] + [x.get("sheet") for x in self.du_lieu_specs(e["module"])]:
                if d and d in self.wb.sheetnames and d not in self.data_sheets and d not in self.result_sheets:
                    self.data_sheets.append(d)

    def module_sites(self, module):
        e = self.entries.get(module) or {}
        return list(e.get("sites") or self.sites)

    def charts(self, module):
        return [c for c in _list((self.entries.get(module) or {}).get("charts"))]

    def du_lieu_specs(self, module):
        return [d for d in _list((self.entries.get(module) or {}).get("du_lieu"))]

    def so(self, module):
        return (self.entries.get(module) or {}).get("so") or {}

    def find_table(self, key, site=None):
        c = [t for t in self.tables if t.key == key and (site is None or t.site == site)]
        return c

    def table_at(self, sheet, r, c):
        for t in self.tables:
            if t.contains(sheet, r, c):
                return t
        return None

    def cell(self, sheet, r, c):
        return self.wb[sheet].cell(r, c)


# ---------------------------------------------------------------- dat bang
class Dat(object):
    """1 lan dat bang: src rows/cols -> o ra (out_sheet, r0, c0)."""

    def __init__(self, bang, out_sheet, r0, c0, rows, cols, hidden=False, tag=None, header=True):
        self.bang, self.out_sheet, self.r0, self.c0 = bang, out_sheet, r0, c0
        self.rows, self.cols, self.hidden, self.tag, self.header = rows, cols, hidden, tag, header
        self.full = rows == list(range(bang.r1, bang.r2 + 1)) and cols == list(range(bang.c1, bang.c2 + 1))
        self._ri = {r: i for i, r in enumerate(rows)}
        self._ci = {c: i for i, c in enumerate(cols)}

    def map(self, r, c):
        if r in self._ri and c in self._ci:
            return self.r0 + self._ri[r], self.c0 + self._ci[c]
        return None

    def contiguous(self, r1, c1, r2, c2):
        a, b = self.map(r1, c1), self.map(r2, c2)
        if not a or not b:
            return False
        return b[0] - a[0] == r2 - r1 and b[1] - a[1] == c2 - c1

    @property
    def r_end(self):
        return self.r0 + len(self.rows) - 1

    @property
    def c_end(self):
        return self.c0 + len(self.cols) - 1

    def out_range(self):
        return "%s!$%s$%d:$%s$%d" % (q(self.out_sheet), L(self.c0), self.r0, L(self.c_end), self.r_end)


REF_RE = re.compile(
    r"(?<![A-Za-z0-9_.$'])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?"
    r"(\$?)([A-Z]{1,3})(\$?)([0-9]+)(?::(\$?)([A-Z]{1,3})(\$?)([0-9]+))?(?![A-Za-z0-9_(!])")


class Planner(object):
    def __init__(self, built, sites):
        self.b = built
        self.sites = sites
        self.dats = []            # visible + hidden
        self.bang_row = 1         # con tro sheet _bang
        self.queue = []

    def visible_for(self, sheet, r, c):
        best = None
        for d in self.dats:
            if d.bang.sheet == sheet and d.map(r, c):
                if best is None or (best.hidden and not d.hidden) or (not best.full and d.full and best.hidden == d.hidden):
                    best = d
        return best

    def hidden_full(self, bang):
        for d in self.dats:
            if d.bang is bang and d.full:
                return d
        rows = list(range(bang.r1, bang.r2 + 1))
        cols = list(range(bang.c1, bang.c2 + 1))
        d = Dat(bang, BANG_SHEET, self.bang_row, 1, rows, cols, hidden=True)
        self.bang_row += len(rows) + 1
        self.dats.append(d)
        self.queue.append(d)
        return d

    def loose(self, sheet, r, c):
        for d in self.dats:
            if d.bang.loose and d.bang.contains(sheet, r, c):
                return d
        t = Bang("_o_%s%d" % (L(c), r), None, self.b.result_sheets.get(sheet), sheet, r, c, r, c, loose=True)
        return self.hidden_full(t)

    def locate(self, sheet, r1, c1, r2=None, c2=None):
        """Tra ve Dat chua vung (r1,c1)-(r2,c2) lien tuc, tao ban an neu can."""
        r2 = r1 if r2 is None else r2
        c2 = c1 if c2 is None else c2
        cands = [d for d in self.dats if d.bang.sheet == sheet and d.contiguous(r1, c1, r2, c2)]
        if cands:
            cands.sort(key=lambda d: (d.hidden, not d.full))
            return cands[0]
        t = self.b.table_at(sheet, r1, c1)
        if t is not None and t.contains(sheet, r2, c2):
            return self.hidden_full(t)
        if (r1, c1) == (r2, c2):
            return self.loose(sheet, r1, c1)
        raise LoiThe("Công thức trỏ vùng %s!%s%d:%s%d nằm ngoài các bảng đăng ký của module (sheet kết quả)."
                     % (q(sheet), L(c1), r1, L(c2), r2))

    def rewrite(self, formula, host_sheet, out_sheet):
        """Doi dia chi tham chieu trong cong thuc (built -> file ra)."""
        parts = formula.split('"')
        for i in range(0, len(parts), 2):
            parts[i] = REF_RE.sub(lambda m: self._sub(m, host_sheet, out_sheet), parts[i])
        return '"'.join(parts)

    def _sub(self, m, host, out_sheet):
        qual = m.group(1)
        sheet = host
        if qual:
            s = qual[:-1]
            sheet = s[1:-1].replace("''", "'") if s.startswith("'") else s
        if sheet not in self.b.result_sheets:
            return m.group(0)
        c1, r1 = column_index_from_string(m.group(3)), int(m.group(5))
        if m.group(7):
            c2, r2 = column_index_from_string(m.group(7)), int(m.group(9))
        else:
            c2, r2 = c1, r1
        d = self.locate(sheet, r1, c1, r2, c2)
        a = d.map(r1, c1)
        pre = "" if d.out_sheet == out_sheet else q(d.out_sheet) + "!"
        s = "%s%s%s%s%d" % (pre, m.group(2), L(a[1]), m.group(4), a[0])
        if m.group(7):
            b = d.map(r2, c2)
            s += ":%s%s%s%d" % (m.group(6), L(b[1]), m.group(8), b[0])
        return s

    def out_ref(self, ref):
        """'Sheet'!B5 (built) -> 'Sheet ra'!$B$5."""
        sheet, r1, c1, r2, c2 = _range(ref)
        if sheet not in self.b.result_sheets:
            return "%s!$%s$%d" % (q(sheet), L(c1), r1)
        d = self.locate(sheet, r1, c1, r2, c2)
        a = d.map(r1, c1)
        s = "%s!$%s$%d" % (q(d.out_sheet), L(a[1]), a[0])
        if (r1, c1) != (r2, c2):
            b = d.map(r2, c2)
            s += ":$%s$%d" % (L(b[1]), b[0])
        return s


def _excel_val(v):
    if isinstance(v, dt.datetime):
        if v.hour or v.minute or v.second:
            return (v - dt.datetime(1899, 12, 30)).total_seconds() / 86400.0
        v = v.date()
    if isinstance(v, dt.date):
        return (v - dt.date(1899, 12, 30)).days
    if isinstance(v, dt.time):
        return None
    return v


# ---------------------------------------------------------------- script GridCraft
class Script(object):
    def __init__(self):
        self.lines = []
        self.active = None

    def cmd(self, command, **params):
        self.lines.append({"command": command, "params": params})

    def sheet(self, name):
        if name != self.active:
            self.cmd("sheet.activate", sheet=name)
            self.active = name

    def values(self, sheet, top_left, rows, chunk=1500):
        if not rows:
            return
        self.sheet(sheet)
        m = re.match(r"([A-Z]+)(\d+)$", top_left)
        c, r = m.group(1), int(m.group(2))
        for i in range(0, len(rows), chunk):
            self.cmd("range.setValues", range="%s%d" % (c, r + i), values=rows[i:i + chunk], parse=False)

    def paint(self, sheet, src, dst):
        self.sheet(sheet)
        self.cmd("selection.set", range=src)
        self.cmd("edit.formatPainter")
        self.cmd("edit.formatPainter", apply=dst)

    def dump(self, path):
        with open(path, "w", encoding="utf-8") as f:
            for x in self.lines:
                f.write(json.dumps(x, ensure_ascii=False) + "\n")


def gc_run(src, script_path, out):
    p = subprocess.run([GRIDCRAFT, "run", "--in", src, "--script", script_path, "--out", out, "--quiet"],
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    err = "\n".join(x for x in p.stderr.splitlines() if not x.startswith("Warning"))
    if p.returncode != 0:
        raise RuntimeError("GridCraft lỗi (rc=%d): %s" % (p.returncode, err[-2000:]))
    return err


# ---------------------------------------------------------------- tien ich dien
def _site_index(params, n_sites, tag):
    s = params.get("site")
    if s is None:
        return None
    i = int(s)
    if i < 1 or i > n_sites:
        raise LoiThe("%r: site=%d nhưng recipe chỉ có %d site" % (tag, i, n_sites))
    return i - 1


def _cot_list(x):
    if x is None:
        return None
    if isinstance(x, str):
        return [s.strip() for s in x.split(",") if s.strip()]
    return list(x)


def _style_src(info, sheet, row, c0, j):
    """Cot nguon dinh dang cho cot thu j (tinh tu c0) o dong row: chinh o do neu co dinh dang, khong thi o gan nhat
    ben trai (trong pham vi bang) co dinh dang, cuoi cung la o c0."""
    st = info[sheet]["styled"]
    for k in range(j, -1, -1):
        if (row, c0 + k) in st:
            return c0 + k
    return c0


class Dien(object):
    def __init__(self, built_xlsx, manifest, mau_path, recipe, notes):
        self.b = Built(built_xlsx, manifest)
        self.mau_path = mau_path
        self.recipe = recipe or {}
        self.notes = {norm(k): v for k, v in (notes or {}).items()}
        self.sites = list(self.recipe.get("sites") or self.b.sites)
        self.names = list(self.recipe.get("names") or [])
        self.P = Planner(self.b, self.sites)
        self.S = Script()
        self.warnings = list(manifest.get("warnings") or [])
        self.notes_out = {}       # out sheet -> list notes
        self.charts_out = {}      # module -> list
        self.blocks_out = {}      # module -> data_blocks
        self.applied = 0

    def ten(self, i):
        return self.names[i] if i < len(self.names) and self.names[i] else self.sites[i]

    # ---------- chu
    def text_val(self, t, cur_site=None):
        """Gia tri the chu (khong gom so)."""
        if t.loai == "khach":
            return self.recipe.get("client") or self.b.manifest.get("client", "")
        if t.loai == "nam":
            return str(dt.date.today().year)
        if t.loai == "ngay":
            return dt.date.today().strftime("%d/%m/%Y")
        if t.loai in ("site", "ten"):
            if not t.khoa:
                if cur_site is None:
                    raise LoiThe("%r: {{%s}} không số thứ tự chỉ dùng trong tieu_de biểu đồ site" % (t, t.loai))
                i = cur_site
            else:
                i = int(t.khoa) - 1
            if i >= len(self.sites):
                return ""
            return self.sites[i] if t.loai == "site" else self.ten(i)
        raise LoiThe("%r: không phải thẻ chữ" % t)

    def sub_text(self, s, cur_site=None):
        out, last = "", 0
        for a, b, inner in tach_the(s):
            loai, khoa, params = phan_tich(inner)
            t = The("", 0, 0, loai, khoa, params, s[a:b], s)
            out += s[last:a] + self.text_val(t, cur_site)
            last = b
        return out + s[last:]

    def so_ref(self, t, out_sheet):
        """-> (bieu thuc Excel, num_fmt)."""
        mod = module_cua_khoa("so", t.khoa)
        spec = self.b.so(mod).get(t.khoa)
        if spec is None:
            raise LoiThe("%r: module %s chưa đăng ký số '%s' (extra.so)" % (t, mod, t.khoa))
        lst = _list(spec)
        si = _site_index(t.params, len(self.sites), t)
        if si is not None:
            pick = [x for x in lst if x.get("site") == self.sites[si]]
        else:
            pick = lst if len(lst) == 1 else [x for x in lst if not x.get("site")]
        if not pick:
            raise LoiThe("%r: không có số '%s' cho site %s" % (t, t.khoa, t.params.get("site", "(chung)")))
        x = pick[0]
        if x.get("cell"):
            ref = self.P.out_ref(x["cell"])
            sh, r, c, _r2, _c2 = _range(x["cell"])
            fmt = x.get("num_fmt") or self.b.cell(sh, r, c).number_format
            return ref, fmt
        host = x.get("sheet") or (self.b.entries.get(mod) or {}).get("sheet")
        f = self.P.rewrite(x["formula"], host, "\x00")
        return f.lstrip("="), x.get("num_fmt") or "General"

    # ---------- ke hoach
    def plan(self, tags, info, sheets):
        # module chay
        present = set(self.b.entries)
        keep_sheets = self.recipe.get("sheets")
        deleted = []
        if keep_sheets:
            for s in sheets:
                if s not in keep_sheets:
                    deleted.append(s)
            miss = [s for s in keep_sheets if s not in sheets]
            if miss:
                self.warnings.append("recipe.sheets có sheet không có trong mẫu: %s" % ", ".join(miss))
        by_sheet = {}
        for t in tags:
            by_sheet.setdefault(t.sheet, []).append(t)
        live = []
        for s in sheets:
            if s in deleted:
                continue
            ts = by_sheet.get(s, [])
            data = [t for t in ts if t.loai in DATA_TYPES]
            missing = [t for t in data if module_cua_khoa(t.loai, t.khoa) not in present]
            if data and len(missing) == len(data):
                deleted.append(s)
                self.warnings.append("Xoá sheet '%s': module %s chưa chạy." % (
                    s, ", ".join(sorted({module_cua_khoa(t.loai, t.khoa) for t in missing}))))
                continue
            for t in missing:
                self.warnings.append("%r: module %s chưa chạy, để trống ô." % (t, module_cua_khoa(t.loai, t.khoa)))
                t.loai = "_bo"
            live += ts
        # the site=i vuot so site cua recipe (mau thiet ke cho nhieu site hon) -> de trong o
        n = len(self.sites)
        for t in live:
            s_ = t.params.get("site")
            if t.loai in DATA_TYPES and t.loai != "so" and s_ and s_.isdigit() and int(s_) > n:
                t.loai = "_bo"
        self.deleted = deleted
        self.tags = [t for t in live if t.loai != "_bo"]
        self.blank = [t for t in live if t.loai == "_bo"]
        self.info = info
        # kich thuoc + chen dong
        sizes = {}
        for t in self.tags:
            if t.loai == "bang":
                rows, cols, bang = self._chon_bang(t)
                sizes[id(t)] = (len(rows) - (0 if t.params.get("bo_tieu_de") != "1" else 0), len(cols))
                t._sel = (bang, rows, cols)
            elif t.loai == "du_lieu":
                sizes[id(t)] = self._kich_thuoc_du_lieu(t)
        self.inserts = {}
        for s in {t.sheet for t in self.tags}:
            self._chen_dong(s, [t for t in self.tags if t.sheet == s], sizes)

    def _chon_bang(self, t):
        si = _site_index(t.params, len(self.sites), t)
        cands = self.b.find_table(t.khoa)
        if not cands:
            raise LoiThe("%r: module chưa đăng ký bảng '%s'" % (t, t.khoa))
        per_site = any(c.site for c in cands)
        if per_site:
            if si is None:
                raise LoiThe("%r: bảng '%s' là bảng theo từng site, cần site=<i>" % (t, t.khoa))
            cs = [c for c in cands if c.site == self.sites[si]]
            if not cs:
                raise LoiThe("%r: không có bảng '%s' của site %s" % (t, t.khoa, self.sites[si]))
            bang = cs[0]
        else:
            bang = cands[0]
        rows = list(range(bang.r1, bang.r2 + 1))
        cols = list(range(bang.c1, bang.c2 + 1))
        ws = self.b.wb[bang.sheet]
        heads = {str(ws.cell(bang.r1, c).value or "").strip(): c for c in cols}
        cot = _cot_list(t.params.get("cot"))
        if si is not None and not per_site:
            dom = self.sites[si]
            if dom not in heads:
                raise LoiThe("%r: bảng '%s' không có cột site %s" % (t, t.khoa, dom))
            cols = [bang.c1, heads[dom]]
        if cot:
            sel = []
            for n in cot:
                if n not in heads:
                    raise LoiThe("%r: bảng '%s' không có cột '%s' (có: %s)" % (t, t.khoa, n, ", ".join(heads)))
                sel.append(heads[n])
            if bang.c1 not in sel:
                sel = [bang.c1] + sel
            cols = sel
        if t.params.get("top"):
            rows = rows[:1 + int(t.params["top"])]
        if t.params.get("bo_tieu_de") == "1":
            rows = rows[1:]
        return rows, cols, bang

    def _du_lieu_spec(self, t):
        mod = module_cua_khoa("du_lieu", t.khoa)
        for d in self.b.du_lieu_specs(mod):
            if d.get("key") == t.khoa:
                return mod, d
        raise LoiThe("%r: module %s chưa đăng ký dữ liệu '%s' (extra.du_lieu)" % (t, mod, t.khoa))

    def _doc_du_lieu(self, t):
        if hasattr(t, "_dl"):
            return t._dl
        mod, d = self._du_lieu_spec(t)
        ws = self.b.wb[d["sheet"]]
        head = [str(c.value or "").strip() for c in ws[1]]
        site_col = d.get("site_col") or ("Website" if "Website" in head else None)
        cot = d.get("cot")
        if cot:
            pairs = [(x, x) if isinstance(x, str) else (x[0], x[1]) for x in cot]
        else:
            pairs = [(h, h) for h in head if h and h != site_col]
        idx = []
        for src, _disp in pairs:
            if src not in head:
                raise LoiThe("%r: sheet '%s' không có cột '%s'" % (t, d["sheet"], src))
            idx.append(head.index(src))
        si = head.index(site_col) if site_col in head else None
        groups, order = {}, []
        fmts = [ws.cell(2, i + 1).number_format for i in idx]
        top = int(t.params["top"]) if t.params.get("top") else None
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(v is not None for v in row):
                continue
            k = row[si] if si is not None else "(tất cả)"
            if k not in groups:
                groups[k] = []
                order.append(k)
            if top and len(groups[k]) >= top:
                continue
            groups[k].append([_excel_val(row[i]) if i < len(row) else None for i in idx])
        ks = [s for s in self.sites if s in groups] + [k for k in order if k not in self.sites]
        t._dl = (mod, d, pairs, groups, ks, fmts, site_col)
        return t._dl

    def _kich_thuoc_du_lieu(self, t):
        mod, d, pairs, groups, ks, fmts, site_col = self._doc_du_lieu(t)
        W = len(pairs)
        if t.params.get("kieu", "canh_nhau") == "doc":
            return (1 + sum(len(groups[k]) for k in ks), W + 1)
        return (2 + max([len(groups[k]) for k in ks] or [0]), len(ks) * (W + 1) - 1)

    def _chen_dong(self, sheet, ts, sizes):
        inf = self.info[sheet]
        occupied = {}
        for (r, c), v in inf["values"].items():
            occupied.setdefault(c, []).append(r)
        for (r, c) in inf["objs"]:
            occupied.setdefault(c, []).append(r)
        ins = {}
        for t in ts:
            if id(t) not in sizes:
                continue
            h, w = sizes[id(t)]
            obs = [r for c in range(t.col, t.col + max(w, 1)) for r in occupied.get(c, []) if r > t.row]
            if not obs:
                continue
            ob = min(obs)
            avail = ob - t.row
            if h + 1 > avail:
                ins[ob] = max(ins.get(ob, 0), h + 1 - avail)
        self.inserts[sheet] = ins

    def shift(self, sheet, r):
        return r + sum(n for at, n in self.inserts.get(sheet, {}).items() if r >= at)

    # ---------- phat lenh
    def emit(self):
        S = self.S
        for s in self.deleted:
            S.cmd("home.deleteSheet", sheet=s)
            S.active = None
        # sheet du lieu an (tao truoc de cong thuc khong #REF)
        for ds in self.b.data_sheets:
            if ds in self.info and ds not in self.deleted:
                raise LoiThe("Mẫu đã có sheet trùng tên sheet dữ liệu '%s'" % ds)
            S.cmd("home.insertSheet", name=ds, before=9999)
            S.active = ds
        S.cmd("home.insertSheet", name=BANG_SHEET, before=9999)
        S.active = BANG_SHEET
        for ds in self.b.data_sheets:
            self._chep_sheet(ds)
        # chen dong
        for sh, ins in self.inserts.items():
            for at in sorted(ins, reverse=True):
                S.sheet(sh)
                S.cmd("home.insertRows", rows="%d:%d" % (at, at + ins[at] - 1))
                self.warnings.append("'%s': chèn %d dòng trước dòng %d để đủ chỗ cho bảng." % (sh, ins[at], at))
        for t in self.tags:
            t.row = self.shift(t.sheet, t.row)
        for t in self.blank:
            t.row = self.shift(t.sheet, t.row)
            S.values(t.sheet, t.cell, [[None]])
        # dat bang hien
        for t in self.tags:
            if t.loai == "bang":
                bang, rows, cols = t._sel
                d = Dat(bang, t.sheet, t.row, t.col, rows, cols, tag=t, header=t.params.get("bo_tieu_de") != "1")
                self.P.dats.append(d)
        # noi dung bang hien (tham chieu co the sinh ban an)
        for d in [x for x in self.P.dats if not x.hidden]:
            self._viet_bang(d)
        # bieu do: lenh de rieng, chen SAU khi moi bang (ke ca ban an) da co gia tri, de GridCraft nhan dung
        # dong tieu de/chuoi
        saved, S.lines, S.active = S.lines, [], None
        for t in self.tags:
            if t.loai == "bieu_do":
                self._bieu_do(t)
        chart_lines, S.lines, S.active = S.lines, saved, None
        for t in self.tags:
            if t.loai == "du_lieu":
                self._du_lieu(t)
        for t in self.tags:
            if t.loai == "nhan_xet":
                self._nhan_xet(t)
        done = set()
        for t in self.tags:
            if t.loai in TEXT_TYPES and (t.sheet, t.row, t.col) not in done:
                done.add((t.sheet, t.row, t.col))
                self._chu(t)
        # ban an (lap toi khi het hang doi)
        while self.P.queue:
            d = self.P.queue.pop(0)
            self._viet_bang(d)
        S.active = None
        S.lines.extend(chart_lines)
        S.active = None
        self._xac_nhan_o_claude()
        self._validations()
        return S

    def _chep_sheet(self, ds):
        ws = self.b.wb[ds]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append([_excel_val(v) for v in row])
        self.S.values(ds, "A1", rows)
        n = len(rows)
        if n >= 2:
            for c in range(1, ws.max_column + 1):
                f = ws.cell(2, c).number_format
                if f and f != "General":
                    self.S.cmd("home.numberFormat", range="%s2:%s%d" % (L(c), L(c), n), code=f)
        self.S.cmd("home.bold", range="A1:%s1" % L(max(ws.max_column, 1)), on=True)

    def _viet_bang(self, d):
        b, S = self.P.b, self.S
        ws = b.wb[d.bang.sheet]
        out = []
        fmts = {}
        for i, r in enumerate(d.rows):
            line = []
            for j, c in enumerate(d.cols):
                cell = ws.cell(r, c)
                v = cell.value
                if isinstance(v, str) and v.startswith("="):
                    v = self.P.rewrite(v, d.bang.sheet, d.out_sheet)
                else:
                    v = _excel_val(v)
                line.append(v)
                f = cell.number_format
                if f and f != "General" and not (d.header and i == 0):
                    fmts.setdefault((j, f), []).append(i)
            out.append(line)
        top = "%s%d" % (L(d.c0), d.r0)
        if not d.hidden:
            self._paint_bang(d)
        S.values(d.out_sheet, top, out)
        for (j, f), idx in fmts.items():
            for a, z in _runs(idx):
                S.cmd("home.numberFormat", range="%s%d:%s%d" % (L(d.c0 + j), d.r0 + a, L(d.c0 + j), d.r0 + z), code=f)
        if d.bang.total_rows and d.rows[-1] == d.bang.r2 and not d.hidden:
            r = d.r_end - d.bang.total_rows + 1
            S.cmd("home.bold", range="%s%d:%s%d" % (L(d.c0), r, L(d.c_end), d.r_end), on=True)

    def _paint_bang(self, d):
        t, S, inf = d.tag, self.S, self.info[d.out_sheet]
        W = len(d.cols)
        hr = t.row
        br = t.row + 1 if d.header else t.row
        body0 = d.r0 + (1 if d.header else 0)
        # style nguon doc tu mau goc (toa do truoc chen dong) -> dung toa do da dich
        o_hr = self._orig_row(t, 0)
        o_br = self._orig_row(t, 1 if d.header else 0)
        for j in range(W):
            if d.header:
                src = _style_src(self.info, t.sheet, o_hr, t.col, j)
                if src != t.col + j:
                    S.paint(t.sheet, "%s%d" % (L(src), hr), "%s%d" % (L(t.col + j), hr))
            if d.r_end >= body0:
                src = _style_src(self.info, t.sheet, o_br, t.col, j)
                S.paint(t.sheet, "%s%d" % (L(src), br), "%s%d:%s%d" % (L(t.col + j), body0, L(t.col + j), d.r_end))
        # xoa dinh dang thua ben phai (mau chuan bi nhieu cot hon)
        for rr, orr in ((hr, o_hr), (br, o_br)):
            c = t.col + W
            while (orr, c) in inf["styled"] and (orr, c) not in inf["values"]:
                S.cmd("edit.clearFormats", range="%s%d" % (L(c), rr))
                c += 1
        # chieu cao dong khi o than xuong dong
        if self._wrap(t.sheet, o_br, t.col, W) and d.r_end >= body0:
            ws = self.P.b.wb[d.bang.sheet]
            widths = [inf["widths"].get(t.col + j, 8.43) for j in range(W)]
            for i, r in enumerate(d.rows[(1 if d.header else 0):]):
                lines = 1
                for j, c in enumerate(d.cols):
                    s = ws.cell(r, c).value
                    s = "" if s is None or (isinstance(s, str) and s.startswith("=")) else str(s)
                    est = sum(max(1, -(-len(x) // max(int(widths[j] * 1.1), 1))) for x in s.split("\n")) if s else 1
                    lines = max(lines, est)
                S.cmd("home.rowHeight", rows="%d:%d" % (body0 + i, body0 + i), points=min(12.5 * lines + 4, 300))

    def _orig_row(self, t, k):
        """Dong goc trong mau (truoc chen dong) cua the t, cong k."""
        ins = self.inserts.get(t.sheet, {})
        for r0 in range(max(1, t.row - sum(ins.values())), t.row + 1):
            if self.shift(t.sheet, r0) == t.row:
                return r0 + k
        return t.row + k

    def _wrap(self, sheet, row, c0, W):
        return getattr(self, "_wrapinfo", {}).get((sheet, row), False)

    # ---------- bieu do
    def _chart_spec(self, t):
        mod = module_cua_khoa("bieu_do", t.khoa)
        for c in self.b.charts(mod):
            if c.get("key") == t.khoa:
                return dict(c)
        if self.b.find_table(t.khoa):
            return {"key": t.khoa, "bang": t.khoa, "loai": t.params.get("loai") or "col"}
        raise LoiThe("%r: module %s chưa đăng ký biểu đồ/bảng '%s'" % (t, mod, t.khoa))

    def _bieu_do(self, t):
        S = self.S
        spec = self._chart_spec(t)
        loai = t.params.get("loai") or spec.get("loai") or "col"
        typ, sub = LOAI_CHART[loai]
        site_param = t.params.get("site")
        targets = []
        if site_param == "*":
            step = int(t.params.get("buoc", "0"))
            for i in range(len(self.sites)):
                targets.append((i, t.col + i * step))
        else:
            targets.append((_site_index(t.params, len(self.sites), t), t.col))
        S.values(t.sheet, t.cell, [[None]])
        for si, col in targets:
            cands = self.b.find_table(spec.get("bang") or t.khoa)
            if not cands:
                raise LoiThe("%r: không có bảng nguồn '%s'" % (t, spec.get("bang")))
            per_site = any(c.site for c in cands)
            if per_site:
                if si is None:
                    raise LoiThe("%r: biểu đồ '%s' theo từng site, cần site=<i> hoặc site=*" % (t, t.khoa))
                cs = [c for c in cands if c.site == self.sites[si]]
                if not cs:
                    raise LoiThe("%r: không có bảng của site %s" % (t, self.sites[si]))
                bang = cs[0]
            else:
                bang = cands[0]
            ws = self.b.wb[bang.sheet]
            heads = [str(ws.cell(bang.r1, c).value or "").strip() for c in range(bang.c1, bang.c2 + 1)]
            body = list(range(bang.r1 + 1, bang.r2 + 1 - bang.total_rows))
            if t.params.get("dong"):
                a, z = [int(x) for x in t.params["dong"].split("-")]
                body = body[a - 1:z]
            cot = _cot_list(t.params.get("cot")) or _cot_list(spec.get("cot"))
            if cot:
                miss = [x for x in cot if x not in heads]
                if miss:
                    raise LoiThe("%r: bảng '%s' không có cột %s" % (t, bang.key, ", ".join(miss)))
                scols = [bang.c1 + heads.index(x) for x in cot]
            elif si is not None and not per_site:
                if self.sites[si] not in heads:
                    raise LoiThe("%r: bảng '%s' không có cột %s" % (t, bang.key, self.sites[si]))
                scols = [bang.c1 + heads.index(self.sites[si])]
            else:
                scols = list(range(bang.c1 + 1, bang.c2 + 1))
                if typ in ("pie", "doughnut"):
                    scols = scols[:1]          # tron chi ve 1 chuoi: cot so dau tien
            # site=i tren bang chung (cot = site): ve cot cua site do, khong theo 'chuoi' mac dinh cua module
            chuoi = t.params.get("chuoi") or (spec.get("chuoi") if si is None or per_site else None) or "cot"
            by_rows = chuoi == "dong"
            rng, cats = self._nguon_chart(bang, body, scols, by_rows)
            title = t.params.get("tieu_de") or spec.get("tieu_de") or ""
            if si is None and ("{{ten}}" in title or "{{site}}" in title):
                title = re.sub(r"\s*[–-]?\s*\{\{(ten|site)\}\}\s*[–-]?\s*", " ", title).strip()
            title = self.sub_text(title, si)
            w = int(t.params.get("rong") or 480)
            h = int(t.params.get("cao") or 290)
            S.sheet(t.sheet)
            S.cmd("selection.set", range="%s%d" % (L(col), t.row))
            p = {"type": typ, "at": "%s%d" % (L(col), t.row), "width": w, "height": h}
            if sub:
                p["subtype"] = sub
            if title:
                p["title"] = title
            S.cmd("insert.chart", **p)
            sd = {"range": rng}
            if by_rows:
                sd["byRows"] = True
            else:
                sd["categories"] = cats
            S.cmd("chart.selectData", **sd)
            colors = _cot_list(t.params.get("mau")) or spec.get("mau")
            if typ not in ("pie", "doughnut"):
                S.cmd("chart.changeColors", colors=colors or PALETTE)
            S.cmd("chart.set", legend="bottom" if typ != "radar" or len(scols) > 1 else "none")
            self.charts_out.setdefault(bang.module, []).append(
                {"key": t.khoa, "site": self.sites[si] if si is not None else None,
                 "sheet": t.sheet, "at": "%s%d" % (L(col), t.row), "loai": loai, "range": rng})

    def _nguon_chart(self, bang, body, scols, by_rows):
        """Tra ve (range du lieu gom dong tieu de [+ cot nhan neu theo dong], categories)."""
        need_rows = [bang.r1] + body
        need_cols = ([bang.c1] if by_rows else []) + scols
        d = None
        for x in sorted([x for x in self.P.dats if x.bang is bang], key=lambda x: (x.hidden, not x.full)):
            if all(x.map(r, c) for r in need_rows for c in need_cols + [bang.c1]):
                d = x
                break
        if d is None:
            d = self.P.hidden_full(bang)
        rr = [d.map(r, bang.c1)[0] for r in need_rows]
        cc = [d.map(bang.r1, c)[1] for c in need_cols]
        contiguous = rr[1:] == list(range(rr[1], rr[1] + len(rr) - 1)) and cc == list(range(cc[0], cc[0] + len(cc)))
        if not contiguous or rr[1] != rr[0] + 1 and not by_rows:
            # ban sao lien tuc trong _bang (cong thuc tro o goc)
            rows = []
            for r in need_rows:
                line = []
                for c in [bang.c1] + scols:
                    a = d.map(r, c)
                    line.append("=%s!%s%d" % (q(d.out_sheet), L(a[1]), a[0]))
                rows.append(line)
            r0 = self.P.bang_row
            self.P.bang_row += len(rows) + 1
            self.S.values(BANG_SHEET, "A%d" % r0, rows)
            n = len(rows)
            sc = 2 if not by_rows else 1
            rng = "%s!$%s$%d:$%s$%d" % (q(BANG_SHEET), L(sc), r0, L(1 + len(scols)), r0 + n - 1)
            cats = "%s!$A$%d:$A$%d" % (q(BANG_SHEET), r0 + 1, r0 + n - 1)
            return rng, cats
        sh = q(d.out_sheet)
        rng = "%s!$%s$%d:$%s$%d" % (sh, L(cc[0]), rr[0], L(cc[-1]), rr[-1])
        lab = d.map(bang.r1, bang.c1)[1]
        cats = "%s!$%s$%d:$%s$%d" % (sh, L(lab), rr[1], L(lab), rr[-1])
        return rng, cats

    # ---------- du lieu
    def _du_lieu(self, t):
        S = self.S
        mod, d, pairs, groups, ks, fmts, site_col = self._doc_du_lieu(t)
        W = len(pairs)
        names = [p[1] for p in pairs]
        o_r = self._orig_row(t, 0)
        widths = self.info[t.sheet]["widths"]
        blocks = []
        if t.params.get("kieu", "canh_nhau") == "doc":
            rows = [["Website"] + names]
            for k in ks:
                rows += [[k] + v for v in groups[k]]
            for j in range(W + 1):
                src = _style_src(self.info, t.sheet, o_r, t.col, j)
                if src != t.col + j:
                    S.paint(t.sheet, "%s%d" % (L(src), t.row), "%s%d" % (L(t.col + j), t.row))
                if len(rows) > 1:
                    src = _style_src(self.info, t.sheet, o_r + 1, t.col, j)
                    S.paint(t.sheet, "%s%d" % (L(src), t.row + 1),
                            "%s%d:%s%d" % (L(t.col + j), t.row + 1, L(t.col + j), t.row + len(rows) - 1))
            S.values(t.sheet, t.cell, rows)
            for j, f in enumerate(fmts):
                if f and f != "General" and len(rows) > 1:
                    S.cmd("home.numberFormat", range="%s%d:%s%d" % (L(t.col + 1 + j), t.row + 1, L(t.col + 1 + j),
                                                                    t.row + len(rows) - 1), code=f)
            blocks.append({"site": "(tất cả)", "sheet": t.sheet,
                           "range": "%s%d:%s%d" % (L(t.col), t.row + 1, L(t.col + W), t.row + max(len(rows) - 1, 1)),
                           "header_row": t.row, "rows": len(rows) - 1, "columns": ["Website"] + names})
        else:
            merges = []
            for i, k in enumerate(ks):
                c0 = t.col + i * (W + 1)
                vals = groups[k]
                last = t.row + 1 + max(len(vals), 1)
                # dinh dang: hang ten site <- o the, hang tieu de <- dong duoi, than <- dong duoi nua (tung cot)
                for j in range(W):
                    for k_, dst in ((0, "%s%d" % (L(c0 + j), t.row)), (1, "%s%d" % (L(c0 + j), t.row + 1)),
                                    (2, "%s%d:%s%d" % (L(c0 + j), t.row + 2, L(c0 + j), last))):
                        src = _style_src(self.info, t.sheet, o_r + k_, t.col, j if k_ else 0)
                        srcc = "%s%d" % (L(src), t.row + k_)
                        if i == 0 and srcc == dst:
                            continue
                        S.paint(t.sheet, srcc, dst)
                if i > 0:
                    for j in range(W + 1):
                        w = widths.get(t.col + j)
                        if w:
                            S.cmd("home.columnWidth", cols="%s:%s" % (L(c0 + j), L(c0 + j)), chars=w)
                S.values(t.sheet, "%s%d" % (L(c0), t.row), [[k] + [None] * (W - 1), names])
                if W > 1:
                    merges.append("%s%d:%s%d" % (L(c0), t.row, L(c0 + W - 1), t.row))
                S.values(t.sheet, "%s%d" % (L(c0), t.row + 2), vals)
                for j, f in enumerate(fmts):
                    if f and f != "General" and vals:
                        S.cmd("home.numberFormat", range="%s%d:%s%d" % (L(c0 + j), t.row + 2, L(c0 + j), last), code=f)
                blocks.append({"site": k, "sheet": t.sheet,
                               "range": "%s%d:%s%d" % (L(c0), t.row + 2, L(c0 + W - 1), last),
                               "header_row": t.row + 1, "rows": len(vals), "columns": names})
            # gop o ten site SAU khi da quet dinh dang (quet tu o da gop se tran sang phai)
            S.sheet(t.sheet)
            for m in merges:
                S.cmd("home.mergeCells", range=m)
        self.blocks_out.setdefault(mod, []).extend(blocks)

    # ---------- nhan xet + chu
    def _nhan_xet(self, t):
        S = self.S
        ref = "%s!%s" % (q(t.sheet), t.cell)
        label = t.params.get("nhan") or t.khoa
        self.notes_out.setdefault(t.sheet, []).append({"cell": ref, "label": label})
        n = self.notes.get(norm(ref))
        if n and n.get("text"):
            S.values(t.sheet, t.cell, [[n["text"]]])
            lines = _nguon(n)
            if lines:
                S.sheet(t.sheet)
                S.cmd("review.newNote", cell=t.cell, text="Nguồn:\n" + "\n".join(lines), author="BD SEO")
            self.applied += 1
        else:
            S.values(t.sheet, t.cell, [[t.params.get("khung") or KHUNG]])

    def _chu(self, t):
        s = t.text
        parts, last, has_so, fmt1 = [], 0, False, None
        for a, b, inner in tach_the(s):
            loai, khoa, params = phan_tich(inner)
            tt = The(t.sheet, t.row, t.col, loai, khoa, params, s[a:b], s)
            if s[last:a]:
                parts.append(("t", s[last:a]))
            if loai == "so":
                expr, fmt = self.so_ref(tt, t.sheet)
                parts.append(("f", expr, fmt))
                has_so = True
                fmt1 = fmt
            else:
                parts.append(("t", self.text_val(tt)))
            last = b
        if s[last:]:
            parts.append(("t", s[last:]))
        if not has_so:
            v = "".join(p[1] for p in parts)
            self.S.values(t.sheet, t.cell, [[v]])
            return
        if len(parts) == 1:
            self.S.values(t.sheet, t.cell, [["=" + parts[0][1]]])
            if fmt1 and fmt1 != "General":
                self.S.cmd("home.numberFormat", range=t.cell, code=fmt1)
            return
        segs = []
        for p in parts:
            if p[0] == "t":
                segs.append('"%s"' % p[1].replace('"', '""'))
            else:
                f = p[2] if p[2] and p[2] != "General" else "#,##0.##"
                segs.append('TEXT(%s,"%s")' % (p[1], f.replace('"', '""')))
        self.S.values(t.sheet, t.cell, [["=" + "&".join(segs)]])

    def _xac_nhan_o_claude(self):
        """Doi claude_cells / pairs sang file ra; chen notes vao o Claude cham."""
        self.claude = {}
        for mod, e in self.b.entries.items():
            cells, items, pairs = [], [], {}
            pr = e.get("claude_pairs") or {}
            for c in e.get("claude_cells") or []:
                c = c["cell"] if isinstance(c, dict) else c
                out = self._map_cell(c)
                if out is None:
                    continue
                cells.append(out)
                if c in pr:
                    po = self._map_cell(pr[c])
                    if po:
                        pairs[out] = po
            for it in e.get("claude_items") or []:
                cs = [self._map_cell(c) for c in it.get("cells", [])]
                cs = [c for c in cs if c]
                if cs:
                    items.append({"row": int(re.search(r"(\d+)$", cs[0]).group(1)), "van_de": it.get("van_de"),
                                  "yeu_cau": it.get("yeu_cau"), "cells": cs})
            self.claude[mod] = (cells, items, pairs)
            for c in cells:
                n = self.notes.get(norm(c))
                if not n:
                    continue
                sh, r, col, _r2, _c2 = _range(c)
                self.S.values(sh, "%s%d" % (L(col), r), [[n.get("text", "")]])
                vals = n.get("values") or []
                if c in pairs and vals:
                    psh, pr_, pc, _a, _b = _range(pairs[c])
                    self.S.values(psh, "%s%d" % (L(pc), pr_), [["; ".join(str(v) for v in vals)]])
                elif _nguon(n):
                    self.S.sheet(sh)
                    self.S.cmd("review.newNote", cell="%s%d" % (L(col), r), text="Nguồn:\n" + "\n".join(_nguon(n)),
                               author="BD SEO")
                self.applied += 1

    def _map_cell(self, ref):
        sh, r, c, _r2, _c2 = _range(ref)
        for d in sorted(self.P.dats, key=lambda x: x.hidden):
            if d.bang.sheet == sh and d.map(r, c):
                a = d.map(r, c)
                return "%s!%s%d" % (q(d.out_sheet), L(a[1]), a[0])
        return None

    def _validations(self):
        self.options_out = {}
        for mod, e in self.b.entries.items():
            vs = list(e.get("validations") or [])
            opts = []
            for o in e.get("options") or []:
                vs.append({"range": o["cell"], "values": o.get("values") or []})
                out = self._map_cell(o["cell"])
                if out:
                    o2 = dict(o)
                    o2["cell"] = out
                    opts.append(o2)
            self.options_out[mod] = opts
            for v in vs:
                sh, r1, c1, r2, c2 = _range(v["range"])
                f1 = '"%s"' % ",".join(v["values"])
                for d in self.P.dats:
                    if d.bang.sheet != sh or d.hidden:
                        continue
                    cells = [d.map(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if d.map(r, c)]
                    if not cells:
                        continue
                    ra, rz = min(x[0] for x in cells), max(x[0] for x in cells)
                    ca, cz = min(x[1] for x in cells), max(x[1] for x in cells)
                    self.S.sheet(d.out_sheet)
                    self.S.cmd("data.validation", range="%s%d:%s%d" % (L(ca), ra, L(cz), rz), type="list",
                               formula1=f1)


def _nguon(n):
    src = n.get("sources") or []
    val = n.get("values") or []
    return ["%s = %s" % (s, val[i] if i < len(val) else "?") for i, s in enumerate(src)]


def _runs(idx):
    out, a, prev = [], None, None
    for i in idx:
        if a is None:
            a = prev = i
        elif i == prev + 1:
            prev = i
        else:
            out.append((a, prev))
            a = prev = i
    if a is not None:
        out.append((a, prev))
    return out


# ---------------------------------------------------------------- ham chinh
def tim_mau(name):
    """Ten mau ('bd-seo') hoac duong dan -> duong dan tuyet doi."""
    p = os.path.abspath(os.path.expanduser(name))
    if os.path.isfile(p):
        return p
    for cand in (os.path.join(MAU_DIR, name), os.path.join(MAU_DIR, name + ".xlsx")):
        if os.path.isfile(cand):
            return cand
    raise LoiThe("Không thấy mẫu '%s' (đường dẫn hoặc tên trong %s)" % (name, MAU_DIR))


def doc_mau_json(mau_path):
    p = os.path.splitext(mau_path)[0] + ".mau.json"
    if os.path.isfile(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)
    return {}


def dien_mau(built_xlsx, manifest, mau_path, out_path, recipe, notes=None):
    """Dien file built vao mau gan the. Tra ve manifest moi (CONTRACT muc 3 + 8)."""
    mau_path = tim_mau(mau_path)
    meta = doc_mau_json(mau_path)
    tags, info, sheets = doc_mau(mau_path)
    errs = kiem_the(tags, meta.get("so_site_toi_da"))
    if errs:
        raise LoiThe("Mẫu %s có thẻ sai:\n  %s" % (os.path.basename(mau_path), "\n  ".join(errs)))
    D = Dien(built_xlsx, manifest, mau_path, recipe, notes)
    mx = meta.get("so_site_toi_da")
    if mx and len(D.sites) > mx:
        D.warnings.append("Recipe có %d site, mẫu %s chỉ thiết kế cho %d site: site thừa có thể tràn bố cục."
                          % (len(D.sites), os.path.basename(mau_path), mx))
    # dinh dang xuong dong cua dong than (de tinh chieu cao)
    wbm = load_workbook(mau_path)
    D._wrapinfo = {}
    for t in tags:
        if t.loai == "bang":
            ws = wbm[t.sheet]
            D._wrapinfo[(t.sheet, t.row + 1)] = bool(ws.cell(t.row + 1, t.col).alignment.wrap_text)
    wbm.close()
    D.plan(tags, info, sheets)
    S = D.emit()
    # ten vung tong hop
    entries = _entries(D, out_path)
    used = set()
    for e in entries:
        for s in e.get("summaries", []):
            nm = s["name"]
            k = 2
            while nm in used:
                nm = "%s_%d" % (s["name"], k)
                k += 1
            s["name"] = nm
            used.add(nm)
            S.cmd("formulas.defineName", name=nm, refersTo="=" + s["range"])
    # an sheet
    for ds in D.b.data_sheets:
        S.cmd("sheet.hide", sheet=ds)
    if D.P.bang_row == 1:
        S.cmd("home.deleteSheet", sheet=BANG_SHEET)
        S.active = None
    else:
        S.cmd("sheet.hide", sheet=BANG_SHEET)
    res = {
        "version": 1, "file": os.path.abspath(out_path), "client": (recipe or {}).get("client") or manifest.get("client", ""),
        "created": dt.datetime.now().replace(microsecond=0).isoformat(),
        "sites": D.sites, "layout": "mau_the", "mau": mau_path, "mau_ten": meta.get("ten", ""),
        "sheets": entries, "warnings": D.warnings, "notes_applied": D.applied, "deleted_sheets": D.deleted,
    }
    known = {norm(n["cell"]) for e in entries for n in e.get("notes", [])}
    known |= {norm(c) for e in entries for c in e.get("claude_cells", [])}
    for k in D.notes:
        if k not in known:
            res["warnings"].append("Nhận xét không có ô tương ứng trong file theo mẫu, bỏ qua: %s" % k)
    text = json.dumps(res, ensure_ascii=False)
    S.cmd("home.insertSheet", name=MANIFEST_SHEET, before=9999)
    S.active = MANIFEST_SHEET
    S.values(MANIFEST_SHEET, "A1", [[text[i:i + 30000]] for i in range(0, len(text), 30000)])
    S.cmd("sheet.hide", sheet=MANIFEST_SHEET)
    first = [s for s in sheets if s not in D.deleted and info[s]["state"] == "visible"]
    if first:
        S.sheet(first[0])
        S.cmd("selection.set", range="A1")
    fd, script_path = tempfile.mkstemp(suffix=".jsonl", prefix="the_mau_")
    os.close(fd)
    try:
        S.dump(script_path)
        out_dir = os.path.dirname(os.path.abspath(out_path))
        os.makedirs(out_dir, exist_ok=True)
        tmp_out = os.path.join(out_dir, ".~gc-" + os.path.basename(out_path))
        gc_run(mau_path, script_path, tmp_out)
        os.replace(tmp_out, out_path)
    finally:
        if os.environ.get("BD_KEEP_SCRIPT"):
            print("script:", script_path, file=sys.stderr)
        else:
            os.remove(script_path)
    return res


def _entries(D, out_path):
    b = D.b
    entries = []
    for mod, e in b.entries.items():
        dats = [d for d in D.P.dats if d.bang.module == mod and not d.bang.loose]
        sums = []
        for d in sorted(dats, key=lambda x: (x.hidden, x.out_sheet, x.r0)):
            if d.hidden and any(x.bang is d.bang and not x.hidden for x in dats):
                continue
            base = "TH_" + re.sub(r"[^0-9A-Za-z_]", "_", d.bang.key + ("_" + d.bang.site if d.bang.site else ""))
            desc = d.bang.desc + (" (bản ẩn ở sheet _bang)" if d.hidden else "")
            s = {"name": base, "key": d.bang.key, "range": d.out_range(), "desc": desc}
            if d.bang.site:
                s["site"] = d.bang.site
            sums.append(s)
            n_body = len(d.rows) - 1
            if n_body + 1 > MAX_SUM_ROWS:
                for p, a in enumerate(range(d.r0 + 1, d.r_end + 1, 50)):
                    z = min(a + 49, d.r_end)
                    sums.append({"name": "%s_p%d" % (base, p + 1), "key": d.bang.key, "part": p + 1,
                                 "range": "%s!$%s$%d:$%s$%d" % (q(d.out_sheet), L(d.c0), a, L(d.c_end), z),
                                 "desc": "%s — phần %d (dòng %d-%d; tiêu đề cột ở dòng %d)" % (
                                     d.bang.key, p + 1, a, z, d.r0)})
        vis = [d for d in dats if not d.hidden]
        cnt = {}
        for d in vis:
            cnt[d.out_sheet] = cnt.get(d.out_sheet, 0) + 1
        sheet = max(cnt, key=lambda k: cnt[k]) if cnt else (e.get("data_sheet") or e["sheet"])
        cells, items, pairs = D.claude.get(mod, ([], [], {}))
        ent = {"sheet": sheet, "module": mod, "purpose": e.get("purpose", ""), "inputs": e.get("inputs"),
               "data_sheet": e.get("data_sheet"), "rows": e.get("rows"), "columns": e.get("columns") or [],
               "summaries": sums, "options": D.options_out.get(mod, []), "caveats": e.get("caveats") or [],
               "notes": [], "data_blocks": D.blocks_out.get(mod, []), "charts": D.charts_out.get(mod, []),
               "sites": e.get("sites") or D.sites}
        if cells:
            ent["claude_cells"] = cells
            ent["claude_items"] = items
            ent["claude_pairs"] = pairs
        if e.get("so"):
            so = {}
            for k, v in e["so"].items():
                lst = []
                for x in _list(v):
                    try:
                        if x.get("cell"):
                            m = D._map_cell(x["cell"])
                            lst.append({"site": x.get("site"), "cell": m} if m else dict(x))
                        else:
                            lst.append(dict(x))
                    except Exception:  # noqa
                        lst.append(dict(x))
                so[k] = lst
            ent["so"] = so
        entries.append(ent)
    for sh, ns in D.notes_out.items():
        entries.insert(0, {"sheet": sh, "module": "mau", "purpose": "Ô nhận xét theo mẫu (Claude viết, có dẫn nguồn)",
                           "data_sheet": None, "rows": 0, "columns": [], "summaries": [], "options": [],
                           "caveats": [], "notes": ns})
    return entries


# ---------------------------------------------------------------- chuyen notes
def chuyen_notes(old_xlsx, new_xlsx):
    """Chuyen notes tu file cu sang file moi: o nhan xet khop theo label (khong dau, bo ky tu la); o Claude cham onpage
    khop theo (van_de, thu tu lap, site). Ghi <new>.notes.json. Tra ve (so nhan xet, so o onpage)."""
    om = json.load(open(old_xlsx + ".manifest.json", encoding="utf-8"))
    on = json.load(open(old_xlsx + ".notes.json", encoding="utf-8")) if os.path.exists(old_xlsx + ".notes.json") else {}
    nm = json.load(open(new_xlsx + ".manifest.json", encoding="utf-8"))
    p_new = new_xlsx + ".notes.json"
    out = json.load(open(p_new, encoding="utf-8")) if os.path.exists(p_new) else {}
    on = {norm(k): v for k, v in on.items()}
    # nhan xet theo label
    new_lab = {}
    for e in nm["sheets"]:
        for n in e.get("notes", []):
            new_lab[_slug(n.get("label"))] = n["cell"]
    n1 = 0
    for e in om["sheets"]:
        for n in e.get("notes", []):
            v = on.get(norm(n["cell"]))
            k = _slug(n.get("label"))
            if v and k in new_lab:
                out[new_lab[k]] = v
                n1 += 1
    # o Claude cham: (van_de, lan xuat hien) + chi so site
    def index(m):
        res = {}
        for e in m["sheets"]:
            if e.get("module") != "onpage":
                continue
            sites = e.get("sites") or m.get("sites") or []
            seen = {}
            for it in e.get("claude_items") or []:
                k = (it.get("van_de") or "").strip()
                seen[k] = seen.get(k, 0) + 1
                for c in it.get("cells", []):
                    col = re.search(r"!\$?([A-Z]+)\$?\d+$", c).group(1)
                    res[norm(c)] = (k, seen[k], col)
            # cot danh gia -> site: thu tu cot tang dan
            cols = sorted({v[2] for v in res.values()}, key=column_index_from_string)
            res = {kk: (v[0], v[1], sites[cols.index(v[2])] if cols.index(v[2]) < len(sites) else v[2])
                   for kk, v in res.items()}
        return res
    oi, ni = index(om), index(nm)
    rev = {v: k for k, v in ni.items()}
    n2 = 0
    for k, v in on.items():
        if k in oi and oi[k] in rev:
            out[rev[oi[k]]] = v
            n2 += 1
    with open(p_new, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    return n1, n2


# ---------------------------------------------------------------- CLI
def quet(mau_path):
    mau_path = tim_mau(mau_path)
    meta = doc_mau_json(mau_path)
    loi = []
    tags, info, sheets = doc_mau(mau_path, loi)
    lines = ["Mẫu: %s%s" % (mau_path, (" — " + meta["ten"]) if meta.get("ten") else "")]
    mods = set()
    for s in sheets:
        ts = [t for t in tags if t.sheet == s]
        lines.append("## %s (%d thẻ)%s" % (s, len(ts), " [ẩn]" if info[s]["state"] != "visible" else ""))
        for t in ts:
            m = module_cua_khoa(t.loai, t.khoa) if t.loai in DATA_TYPES else None
            if m:
                mods.add(m)
            lines.append("  %-6s %-9s %-22s %s%s" % (t.cell, t.loai, t.khoa,
                                                     " ".join("%s=%s" % kv for kv in t.params.items()),
                                                     ("  [%s]" % m) if m else ""))
    errs = loi + kiem_the(tags, meta.get("so_site_toi_da"))
    lines.append("Module cần: %s" % (", ".join(sorted(mods)) or "-"))
    can = meta.get("modules") or meta.get("module_can")
    if can and not mods <= set(can) | set(meta.get("module_tuy_chon") or []):
        lines.append("! mau.json khai module %s, thẻ cần %s" % (can, sorted(mods)))
    lines.append("LỖI (%d):\n  %s" % (len(errs), "\n  ".join(errs)) if errs else "Không có lỗi thẻ.")
    return "\n".join(lines), errs


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    a0 = sub.add_parser("quet")
    a0.add_argument("mau")
    a1 = sub.add_parser("dien")
    a1.add_argument("--built", required=True)
    a1.add_argument("--manifest", required=True)
    a1.add_argument("--mau", required=True)
    a1.add_argument("--out", required=True)
    a1.add_argument("--recipe")
    a1.add_argument("--notes")
    a2 = sub.add_parser("chuyen-notes")
    a2.add_argument("--tu", required=True)
    a2.add_argument("--sang", required=True)
    a = ap.parse_args(argv)
    try:
        if a.cmd == "quet":
            text, errs = quet(a.mau)
            print(text)
            return 1 if errs else 0
        if a.cmd == "dien":
            m = json.load(open(a.manifest, encoding="utf-8"))
            r = json.load(open(a.recipe, encoding="utf-8")) if a.recipe else {"sites": m.get("sites"), "client": m.get("client")}
            n = json.load(open(a.notes, encoding="utf-8")) if a.notes else None
            res = dien_mau(a.built, m, a.mau, a.out, r, n)
            with open(a.out + ".manifest.json", "w", encoding="utf-8") as f:
                json.dump(res, f, ensure_ascii=False, indent=1)
            print("Đã ghi %s" % a.out)
            for w in res["warnings"]:
                print("CẢNH BÁO: " + w)
            return 0
        n1, n2 = chuyen_notes(os.path.abspath(a.tu), os.path.abspath(a.sang))
        print("Đã chuyển %d nhận xét, %d ô Claude chấm" % (n1, n2))
        return 0
    except LoiThe as e:
        print("LỖI: %s" % e, file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
