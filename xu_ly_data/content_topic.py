"""Sheet 1.7 So luong content: phan loai Top pages Ahrefs theo Loai trang + Topic cho N website.

Dung (CLI):
    .venv/bin/python -m xu_ly_data.content_topic --out kq.xlsx [--rules rules.json] a-top-pages.csv b-top-pages.csv
    .venv/bin/python -m xu_ly_data.content_topic --sample-unclassified --rules rules.json -n 150 a-top-pages.csv

- Sheet "TP data": du lieu xep doc, Loai trang + Topic la GIA TRI (Python tinh san), con lai la so lieu goc.
- Sheet "1.7 So luong content": moi bang tong hop la HAM (COUNTIFS/SUMIFS tro vao 'TP data').
- rules.json: {"topics":[{"name":"Bong da","pattern":"bong-da|football"}], "default":"Khac"}
  so tren slug da bo dau (unquote + bo dau + lower), first match wins.
Python 3.9: khong dung cu phap X | None.
"""
import argparse
import csv
import json
import os
import re
import sys
import unicodedata
from typing import Dict, List, Optional, Tuple
from urllib.parse import unquote, urlparse

from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from . import common

KEY = "content_topic"
SHEET = "1.7 Số lượng content"
INPUT_KIND = "ahrefs_top_pages_csv"
DESC = "Số lượng content theo loại trang và topic (Top pages Ahrefs)"

DATA_SHEET = "TP data"
LOAI_TRANG = ["Trang chủ", "Danh mục", "Sản phẩm/Dịch vụ", "Blog", "Landing", "Chính sách/hỗ trợ", "Khác"]
CHUA_PL = "Chưa phân loại"
NGUONG_TRANG_CHU = 30.0  # % traffic dồn vào trang chủ thì cảnh báo

DATA_COLS = [
    ("Website", "Tên website (lấy từ tên file Ahrefs)"),
    ("URL", "URL trang"),
    ("Loại trang", "Trang chủ/Danh mục/Sản phẩm-Dịch vụ/Blog/Landing/Chính sách-hỗ trợ/Khác (suy từ URL + Page type Ahrefs, giá trị)"),
    ("Topic", "Chủ đề theo file luật (giá trị)"),
    ("Page type (Ahrefs)", "Page type gốc của Ahrefs"),
    ("Traffic", "Organic traffic ước tính/tháng"),
    ("Traffic value", "Giá trị traffic (Ahrefs)"),
    ("Keywords", "Số từ khóa trang xếp hạng"),
    ("Top keyword", "Từ khóa kéo traffic nhiều nhất"),
    ("Top KW volume", "Volume của top keyword"),
]


# ---------------------------------------------------------------- tien ich
def khong_dau(s):
    # type: (str) -> str
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def _num(s):
    try:
        return float(str(s).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _nhan(s):
    """Nhan dung lam dieu kien COUNTIFS/SUMIFS: bo ky tu wildcard va ky tu mo dau la toan tu."""
    s = re.sub(r"[~*?]", "", str(s)).strip()
    return s.lstrip("=<>") or "?"


def site_tu_ten_file(path):
    name = os.path.basename(path)
    m = re.match(r"(.+?)-top-pages", name)
    return m.group(1) if m else os.path.splitext(name)[0]


_LANG_REGION = re.compile(r"^[a-z]{2}-[A-Z]{2}$")  # en-VN, vi-VN
_LANG_SHORT = ("en", "vi")


def duong_dan(url):
    # type: (str) -> List[str]
    """Cac doan path (da unquote), bo tien to ngon ngu (en, vi-VN...). Bo query/fragment."""
    p = urlparse(url if "//" in url else "//" + url)
    seg = [s for s in unquote(p.path or "").split("/") if s]
    if seg and (_LANG_REGION.match(seg[0]) or seg[0] in _LANG_SHORT):
        seg = seg[1:]
    return seg


def slug_khong_dau(url):
    return khong_dau("/".join(duong_dan(url)))


# ---------------------------------------------------------------- phan loai trang
_BLOG = {"blog", "blogs", "tin-tuc", "tintuc", "tin", "kien-thuc", "kienthuc", "news", "bai-viet", "baiviet",
         "cam-nang", "camnang", "article", "articles", "post", "posts", "magazine", "tin-tuc-su-kien",
         "chia-se", "tu-van", "review", "reviews", "huong-dan", "su-kien", "wiki", "hoi-dap", "thu-vien"}
_DANH_MUC = {"c", "cat", "category", "categories", "danh-muc", "danhmuc", "collections", "collection",
             "product-category", "thuong-hieu", "brand", "brands", "chuyen-muc", "nganh-hang", "shop"}
_SAN_PHAM = {"p", "sp", "san-pham", "sanpham", "product", "products", "item", "items", "dich-vu", "dichvu",
             "service", "services", "pdp", "giai-phap", "solutions"}
_LANDING = {"s", "lp", "landing", "landing-page", "khuyen-mai", "promotion", "promotions", "sale", "campaign",
            "event", "uu-dai", "chuong-trinh", "deal", "flash-sale"}
_CHINH_SACH = re.compile(
    r"chinh-sach|policy|privacy|terms|dieu-khoan|dieu-kien|bao-mat|giao-hang|van-chuyen|doi-tra|hoan-tra|"
    r"bao-hanh|thanh-toan|huong-dan-(mua|dat|thanh)|ho-tro|support|help|faq|cau-hoi|lien-he|contact|"
    r"gioi-thieu|about|tuyen-dung|careers?|cua-hang|he-thong|click-and-collect|sitemap|stores?")


def phan_loai_trang(url, page_type=""):
    # type: (str, str) -> str
    seg = [khong_dau(s) for s in duong_dan(url)]
    pt = (page_type or "").strip().lower()
    if not seg:
        return "Trang chủ"
    joined = "/".join(seg[:2])
    s0 = seg[0]
    # 1) thu muc cap 1 la dau hieu ro nhat
    if s0 in _BLOG:
        return "Blog"
    if s0 in _DANH_MUC:
        return "Danh mục"
    if s0 in _SAN_PHAM:
        return "Sản phẩm/Dịch vụ"
    # 2) trang chinh sach/ho tro (ke ca nam trong /s/, /landing/...)
    if _CHINH_SACH.search(joined):
        return "Chính sách/hỗ trợ"
    if s0 in _LANDING:
        return "Landing"
    # 3) thu muc cap 2
    if len(seg) > 1:
        if seg[1] in _BLOG:
            return "Blog"
        if seg[1] in _DANH_MUC:
            return "Danh mục"
        if seg[1] in _SAN_PHAM:
            return "Sản phẩm/Dịch vụ"
    # khong nhan ra qua URL -> dung Page type cua Ahrefs
    if pt.startswith("article"):
        return "Blog"
    if pt.startswith("listing"):
        return "Danh mục"
    if pt.startswith("product"):
        return "Sản phẩm/Dịch vụ"
    if pt.startswith("landing"):
        return "Landing"
    if pt.startswith("site page"):
        return "Trang chủ" if "home" in pt else "Chính sách/hỗ trợ"
    return "Khác"


# ---------------------------------------------------------------- topic
def doc_luat(rules):
    # type: (object) -> Optional[Dict]
    """rules: duong dan file JSON, dict, hoac None -> {"topics":[(name, regex)], "default": str} hoac None."""
    if not rules:
        return None
    if isinstance(rules, str):
        with open(rules, encoding="utf-8") as f:
            rules = json.load(f)
    topics = []
    for t in rules.get("topics", []):
        topics.append((t["name"], re.compile(t["pattern"], re.I)))
    return {"topics": topics, "default": rules.get("default") or "Khác"}


def phan_loai_topic(url, luat):
    # type: (str, Optional[Dict]) -> str
    if not luat:
        return CHUA_PL
    slug = slug_khong_dau(url)
    for name, rx in luat["topics"]:
        if rx.search(slug):
            return name
    return luat["default"]


# ---------------------------------------------------------------- doc file
def doc_file(path, luat=None, site=None):
    # type: (str, Optional[Dict], Optional[str]) -> Tuple[str, List[list]]
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    site = site or site_tu_ten_file(path)
    if not rows:
        return site, []
    cols = {k.strip().lower(): k for k in rows[0].keys()}

    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None

    c_url, c_tr = col("url", "page url"), col("traffic", "organic traffic")
    if not (c_url and c_tr):
        raise ValueError("%s: thiếu cột URL/Traffic — không phải file Top pages?" % path)
    c_pt, c_tv, c_kw = col("page type"), col("traffic value"), col("keywords")
    c_tk, c_vol = col("top keyword"), col("top keyword: volume", "top keyword volume")
    out = []
    for r in rows:
        url = (r[c_url] or "").strip()
        if not url:
            continue
        pt = (r.get(c_pt) or "").strip() if c_pt else ""
        out.append([
            site, url, phan_loai_trang(url, pt), phan_loai_topic(url, luat), pt,
            _num(r[c_tr]), _num(r.get(c_tv)) if c_tv else 0, int(_num(r.get(c_kw))) if c_kw else 0,
            (r.get(c_tk) or "").strip() if c_tk else "", int(_num(r.get(c_vol))) if c_vol else 0,
        ])
    return site, out


def _ten_site_duy_nhat(sites):
    seen, out = {}, []
    for s in sites:
        k = seen.get(s, 0)
        seen[s] = k + 1
        out.append(s if k == 0 else "%s (%d)" % (s, k + 1))
    return out


def xuat_mau_chua_phan_loai(inputs, rules, n=150):
    # type: (List[str], object, int) -> List[str]
    """Slug (bo dau) cua cac URL chua phan loai (Topic = default/Chua phan loai), xep theo traffic giam dan."""
    luat = doc_luat(rules)
    tong = {}  # type: Dict[str, float]
    for p in inputs:
        _site, rows = doc_file(p, luat)
        for r in rows:
            if luat is None or r[3] == luat["default"]:
                s = slug_khong_dau(r[1])
                tong[s] = tong.get(s, 0) + r[5]
    return [s for s, _ in sorted(tong.items(), key=lambda kv: -kv[1])[:n]]


# ---------------------------------------------------------------- build
def build(wb, inputs, opts, ctx):
    opts = opts or {}
    ctx = ctx or {}
    luat = doc_luat(opts.get("rules"))
    parsed = [doc_file(p, luat) for p in inputs]
    names = opts.get("sites") or None
    if not names and ctx.get("sites") and len(ctx["sites"]) == len(inputs):
        names = common.ten_site_theo_file(inputs, ctx["sites"])
    if names and len(names) == len(inputs):
        parsed = [(names[i], rows) for i, (_, rows) in enumerate(parsed)]
        for i, (s, rows) in enumerate(parsed):
            for r in rows:
                r[0] = s
    sites = _ten_site_duy_nhat([_nhan(s) for s, _ in parsed])
    for i, (_, rows) in enumerate(parsed):
        for r in rows:
            r[0] = sites[i]
    nguong_tc = float(opts.get("nguong_trang_chu", NGUONG_TRANG_CHU))

    # ---------- TP data ----------
    dws = wb.create_sheet(DATA_SHEET)
    common.header_row(dws, 1, 1, [c[0] for c in DATA_COLS])
    n = 1
    for _s, rows in parsed:
        for r in sorted(rows, key=lambda x: -x[5]):
            n += 1
            dws.append(r)
    last = max(n, 2)
    for i, w in enumerate([18, 60, 18, 24, 22, 10, 12, 9, 30, 12], 1):
        dws.column_dimensions[get_column_letter(i)].width = w
    for c in ("F", "G"):
        for cell in dws[c][1:]:
            cell.number_format = "#,##0"
    dws.freeze_panes = "A2"
    dws.auto_filter.ref = "A1:J%d" % last

    D = lambda c: "'%s'!$%s$2:$%s$%d" % (DATA_SHEET, c, c, last)

    # ---------- thong ke Python (de tinh caveats + sap xep topic) ----------
    all_rows = [r for _s, rows in parsed for r in rows]
    tr_site = {s: 0.0 for s in sites}
    tr_topic = {}  # type: Dict[str, float]
    for r in all_rows:
        tr_site[r[0]] += r[5]
        tr_topic[r[3]] = tr_topic.get(r[3], 0) + r[5]
    topics = sorted(tr_topic, key=lambda t: (-tr_topic[t], t))
    if len(topics) > 40:  # giu vung tong hop <= 60 dong
        topics = topics[:40]

    # ---------- sheet ket qua ----------
    ws = wb.create_sheet(SHEET)
    common.title(ws, "SỐ LƯỢNG CONTENT THEO LOẠI TRANG VÀ TOPIC",
                 "Nguồn: Ahrefs › Top pages (export CSV). Loại trang/Topic là giá trị tính sẵn trong sheet '%s'; "
                 "các bảng dưới là hàm COUNTIFS/SUMIFS." % DATA_SHEET)
    ncol = len(sites)
    sums = []

    def bang(row, label, hdr0, labels, kind, tot_label="Tổng", data_col=None, ref_block=None):
        """Dung 1 bang: nhan o cot A, moi site 1 cot. kind: 'count' | 'sum' | 'pct'.
        Tra ve (dong dau data, dong cuoi data, dong tong)."""
        ws.cell(row, 1, label).font = Font(bold=True, color=common.XANH)
        common.header_row(ws, row + 1, 1, [hdr0] + sites)
        r0 = row + 2
        for i, lab in enumerate(labels):
            r = r0 + i
            ws.cell(r, 1, lab).border = common.BOX
            for j in range(ncol):
                cl = get_column_letter(2 + j)
                crit_col = "C" if data_col == "loai" else "D"
                if kind == "count":
                    f = '=COUNTIFS(%s,%s$%d,%s,$A%d)' % (D("A"), cl, row + 1, D(crit_col), r)
                elif kind == "sum":
                    f = '=SUMIFS(%s,%s,%s$%d,%s,$A%d)' % (D("F"), D("A"), cl, row + 1, D(crit_col), r)
                else:  # pct: tham chieu bang traffic (ref_block) cung vi tri
                    rb0, rbt = ref_block
                    f = '=IFERROR(%s%d/%s$%d,0)' % (cl, rb0 + i, cl, rbt)
                c = ws.cell(r, 2 + j, f)
                c.border = common.BOX
                c.number_format = "0.0%" if kind == "pct" else "#,##0"
        rt = r0 + len(labels)
        ws.cell(rt, 1, tot_label).font = Font(bold=True)
        ws.cell(rt, 1).border = common.BOX
        for j in range(ncol):
            cl = get_column_letter(2 + j)
            f = "=SUM(%s%d:%s%d)" % (cl, r0, cl, rt - 1)
            c = ws.cell(rt, 2 + j, f)
            c.font = Font(bold=True)
            c.border = common.BOX
            c.number_format = "0.0%" if kind == "pct" else "#,##0"
        return r0, rt - 1, rt

    def ref(r_hdr, r_end, key, desc, name_key):
        sums.append(common.bang(wb, ws, key, "A%d:%s%d" % (r_hdr, get_column_letter(1 + ncol), r_end), desc,
                                total_rows=1, name=name_key))

    so = {"ct_so_trang": [], "ct_traffic": []}
    # ct_tong: site (dong) x So trang, Traffic
    row = 4
    ws.cell(row, 1, "Tổng số trang có traffic và traffic theo website").font = Font(bold=True, color=common.XANH)
    common.header_row(ws, row + 1, 1, ["Website", "Số trang", "Traffic"])
    for k, s_ in enumerate(sites):
        r = row + 2 + k
        ws.cell(r, 1, s_).border = common.BOX
        c = ws.cell(r, 2, "=COUNTIFS(%s,$A%d)" % (D("A"), r))
        c.number_format = "#,##0"
        c.border = common.BOX
        c = ws.cell(r, 3, "=SUMIFS(%s,%s,$A%d)" % (D("F"), D("A"), r))
        c.number_format = "#,##0"
        c.border = common.BOX
        so["ct_so_trang"].append(common.so_cell(ws, "B%d" % r, s_))
        so["ct_traffic"].append(common.so_cell(ws, "C%d" % r, s_))
    sums.append(common.bang(wb, ws, "ct_tong", "A%d:C%d" % (row + 1, row + 1 + ncol),
                            "Số trang có traffic và tổng traffic theo website (dòng = website)."))
    row = row + 3 + ncol
    # (b) Topic x Website
    b1 = bang(row, "(b1) Topic × Website: số trang", "Topic", topics, "count", data_col="topic")
    ref(row + 1, b1[2], "ct_topic_sotrang", "Số trang theo topic × website (dòng cuối = tổng)", "Topic_SoTrang")
    row = b1[2] + 2
    b2 = bang(row, "(b2) Topic × Website: traffic", "Topic (traffic)", topics, "sum", data_col="topic")
    ref(row + 1, b2[2], "ct_topic_traffic", "Traffic theo topic × website (dòng cuối = tổng)", "Topic_Traffic")
    row = b2[2] + 2
    # (a) Loai trang x Website
    a1 = bang(row, "(a1) Loại trang × Website: số trang", "Loại trang", LOAI_TRANG, "count", data_col="loai")
    ref(row + 1, a1[2], "ct_loaitrang_sotrang", "Số trang theo loại trang × website (dòng cuối = tổng)", "LoaiTrang_SoTrang")
    row = a1[2] + 2
    a2 = bang(row, "(a2) Loại trang × Website: traffic", "Loại trang (traffic)", LOAI_TRANG, "sum", data_col="loai")
    ref(row + 1, a2[2], "ct_loaitrang_traffic", "Traffic theo loại trang × website (dòng cuối = tổng)", "LoaiTrang_Traffic")
    row = a2[2] + 2
    a3 = bang(row, "(a3) Loại trang × Website: % traffic", "Loại trang (% traffic)", LOAI_TRANG, "pct",
              tot_label="Tổng", ref_block=(a2[0], a2[2]))
    ref(row + 1, a3[2], "ct_loaitrang_pct", "% traffic theo loại trang trong từng website", "LoaiTrang_PctTraffic")
    row = a3[2] + 2

    # nhan xet
    notes = [common.note_area(ws, "A%d" % (row + 1), max(ncol + 1, 4), 8, "NHẬN XÉT")]
    row += 11

    # (c) top 10 URL moi website
    ws.cell(row, 1, "(c) Top 10 URL theo traffic của từng website (giá trị)").font = Font(bold=True, color=common.XANH)
    row += 1
    for s, (_n, rows) in zip(sites, parsed):
        ws.cell(row, 1, s).font = Font(bold=True)
        common.header_row(ws, row + 1, 1, ["URL", "Loại trang", "Topic", "Traffic"])
        for i, r in enumerate(sorted(rows, key=lambda x: -x[5])[:10]):
            rr = row + 2 + i
            for j, v in enumerate([r[1], r[2], r[3], r[5]]):
                c = ws.cell(rr, 1 + j, v)
                c.border = common.BOX
            ws.cell(rr, 4).number_format = "#,##0"
        row += 13

    ws.column_dimensions["A"].width = 55
    for j in range(max(ncol, 3)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 22
    ws.freeze_panes = "B4"

    # bieu do cot cum: traffic theo Topic x Website
    ch = BarChart()
    ch.type = "col"
    ch.grouping = "clustered"
    ch.title = "Traffic theo Topic × Website"
    ch.y_axis.title = "Traffic"
    ch.height, ch.width = 11, 24
    for j in range(ncol):
        ch.add_data(Reference(ws, min_col=2 + j, min_row=b2[0] - 1, max_row=b2[1]), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=b2[0], max_row=b2[1]))
    ws.add_chart(ch, "%s4" % get_column_letter(ncol + 3))

    # ---------- caveats ----------
    caveats = []
    mac_dinh = {CHUA_PL, (luat or {}).get("default", "Khác")}
    for s in sites:
        rs = [r for r in all_rows if r[0] == s]
        tt = tr_site[s]
        if not rs:
            caveats.append("%s: file không có dòng nào." % s)
            continue
        caveats.append("%s: %s URL, tổng traffic %s." % (s, format(len(rs), ","), format(int(round(tt)), ",")))
        if tt > 0:
            pl = sum(r[5] for r in rs if r[3] in mac_dinh) / tt * 100
            lk = sum(r[5] for r in rs if r[2] == "Khác") / tt * 100
            nhan = "Chưa phân loại" if luat is None else "'%s'" % luat["default"]
            caveats.append("%s: %.1f%% traffic Topic %s; %.1f%% traffic Loại trang 'Khác'." % (s, pl, nhan, lk))
            home = sum(r[5] for r in rs if r[2] == "Trang chủ") / tt * 100
            if home > nguong_tc:
                caveats.append("%s: %.1f%% traffic dồn vào trang chủ (> %.0f%%), tỷ trọng các loại trang khác bị méo." % (s, home, nguong_tc))
    if luat is None:
        caveats.append("Chưa có file luật topic (opts.rules): toàn bộ Topic = 'Chưa phân loại'. Dùng xuat_mau_chua_phan_loai để lấy mẫu viết luật.")
    if len(tr_topic) > 40:
        caveats.append("Có %d topic, bảng chỉ liệt kê 40 topic đầu theo traffic; phần còn lại không hiện trong (b)." % len(tr_topic))

    columns = [{"name": n_, "desc": d} for n_, d in DATA_COLS]
    return common.entry(
        SHEET, KEY, "Số trang và traffic theo loại trang và topic của từng website (Ahrefs Top pages)",
        inputs, data_sheet=DATA_SHEET, rows=len(all_rows), columns=columns, summaries=sums,
        options=[], caveats=caveats, notes=notes,
        extra={"sites": sites, "so": so,
               "charts": [{"key": "ct_tong", "bang": "ct_tong", "loai": "bar", "cot": ["Số trang"],
                           "tieu_de": "Số trang có traffic theo website"},
                          {"key": "ct_loaitrang", "bang": "ct_loaitrang_sotrang", "loai": "pie",
                           "tieu_de": "{{ten}} – số trang theo loại trang"},
                          {"key": "ct_topic", "bang": "ct_topic_sotrang", "loai": "pie",
                           "tieu_de": "{{ten}} – số trang theo topic"}],
               "du_lieu": [{"key": "ct_data", "sheet": DATA_SHEET, "site_col": "Website",
                            "cot": [["URL", "URL"], ["Traffic", "Traffic"], ["Top keyword", "Top keyword"],
                                    ["Loại trang", "Loại Trang"], ["Topic", "Topic"], ["Top KW volume", "Top KW volume"]]}]})


# ---------------------------------------------------------------- CLI
def main(argv=None):
    ap = argparse.ArgumentParser(description=DESC)
    ap.add_argument("files", nargs="+", help="file Top pages CSV (mỗi website một file)")
    ap.add_argument("--rules", help="file luật topic JSON")
    ap.add_argument("--out", help="file xlsx đầu ra")
    ap.add_argument("--sample-unclassified", action="store_true", help="in slug chưa phân loại theo traffic")
    ap.add_argument("-n", type=int, default=150)
    a = ap.parse_args(argv)
    if a.sample_unclassified:
        for s in xuat_mau_chua_phan_loai(a.files, a.rules, a.n):
            print(s)
        return 0
    if not a.out:
        ap.error("cần --out hoặc --sample-unclassified")
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    e = build(wb, a.files, {"rules": a.rules} if a.rules else {}, {})
    wb.save(a.out)
    print(json.dumps(e, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
