"""Sheet 1.5 Du lieu Onpage: cham tieu chi onpage cho N website tu ket qua crawl Mini Frog.

Dung (CLI):
    .venv/bin/python -m xu_ly_data.onpage --out kq.xlsx a-crawl.xlsx b-records.json ...

Dau vao moi website: file xlsx xuat tu /api/export/xlsx (sheet co cot URL + Status + So H1...),
hoac JSON: list records /api/results, hoac dict {"records": [...], "site": "...", "sitemap_count": n,
"orphan_count": n, "max_urls": n, "limited": bool}.

- Tieu chi lay tu xu_ly_data/onpage_tieuchi.json (mau BD SEO, dong 9-117 + vai tieu chi bo sung extra=true).
  skip=true (CMS, GSC, GA4) khong dua vao sheet. auto=<id luat> -> tool tu cham; auto=null -> de trong,
  cot "Nguoi cham" = Claude, Claude ghi vao o (entry extra.claude_cells).
- Bang diem nhom o tren dung COUNTIFS "Dat"; bang tieu chi o duoi. Cac tieu chi duoc xep theo nhom (lien tuc)
  de vung tong hop doc gon, thu tu trong nhom giu nhu mau.
Python 3.9: khong dung cu phap X | None.
"""
import argparse
import json
import os
import re
import statistics
import sys
import unicodedata
from typing import Dict, List, Optional, Tuple
from urllib.parse import unquote, urlparse

from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from . import common

KEY = "onpage"
SHEET = "1.5 Dữ liệu Onpage"
INPUT_KIND = "minifrog_xlsx"
DESC = "Chấm tiêu chí onpage từ kết quả crawl Mini Frog"

DAT, CHUA_DAT = "Đạt", "Chưa đạt"
TIEUCHI_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "onpage_tieuchi.json")
HDR_ROW_TC = 14  # dong tieu de bang tieu chi
MAX_ROWS_SUMMARY = 55


def _khong_dau(s):
    s = s.replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


# ---------------------------------------------------------------- doc dau vao
_XLSX_MAP = {
    "URL": "url", "Loại trang": "page_type_label", "Độ sâu thư mục": "folder_depth", "Status": "status",
    "Redirect": "_redirect", "Depth": "depth", "Title": "title", "Dài title": "title_length",
    "Meta description": "meta_description", "Dài meta": "meta_desc_length", "Số H1": "h1_count",
    "Số H2": "h2_count", "Canonical": "canonical", "Canonical đúng": "_canonical_ok",
    "Meta robots": "meta_robots", "Thumbnail (og:image)": "og_image", "Số ảnh": "image_count",
    "Ảnh thiếu alt": "images_missing_alt", "Số từ": "word_count", "Inlinks": "inlink_count",
    "Phản hồi (ms)": "response_ms", "Lỗi": "error",
}
_LABEL_KEY = (("trang chủ", "homepage"), ("danh mục", "category"), ("sản phẩm", "product"), ("blog", "blog"))


def _page_type(label):
    l = (label or "").strip().lower()
    for pre, k in _LABEL_KEY:
        if l.startswith(pre):
            return k
    return "other"


def _int(v, default=0):
    try:
        return int(float(str(v).replace(",", "")))
    except (TypeError, ValueError):
        return default


def _host(url):
    h = (urlparse(url).hostname or "")
    return h[4:] if h.startswith("www.") else h


def _doc_xlsx(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    meta = {}  # type: Dict
    ws_url = None
    for ws in wb.worksheets:
        head = [str(c.value or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1), [])]
        if "URL" in head and "Status" in head:
            ws_url, head_url = ws, head
            break
    if ws_url is None:
        raise ValueError("%s: không thấy sheet có cột URL + Status (file xuất Excel của Mini Frog)" % path)
    idx = {name: i for i, name in enumerate(head_url)}
    recs = []
    for row in ws_url.iter_rows(min_row=2, values_only=True):
        if not row or not row[idx["URL"]]:
            continue

        def g(name, d=""):
            i = idx.get(name)
            v = row[i] if i is not None and i < len(row) else None
            return d if v is None else v
        st = g("Status")
        r = {
            "url": str(g("URL")), "page_type": _page_type(g("Loại trang")),
            "status": int(st) if str(st).strip().isdigit() else (str(st) or None),
            "redirect_chain": [x for x in str(g("Redirect")).split(" ; ") if x.strip()],
            "title": str(g("Title")).strip() or None, "title_length": _int(g("Dài title")),
            "meta_description": str(g("Meta description")).strip() or None, "meta_desc_length": _int(g("Dài meta")),
            "h1_count": _int(g("Số H1")), "h2_count": _int(g("Số H2")),
            "canonical": str(g("Canonical")).strip() or None,
            "canonical_ok": {"Đúng": True, "Sai/Thiếu": False}.get(str(g("Canonical đúng")).strip()),
            "meta_robots": str(g("Meta robots")).strip() or None,
            "og_image": str(g("Thumbnail (og:image)")).strip() or None,
            "image_count": _int(g("Số ảnh")), "images_missing_alt": _int(g("Ảnh thiếu alt")),
            "word_count": _int(g("Số từ")), "inlink_count": _int(g("Inlinks")),
            "depth": _int(g("Depth")), "folder_depth": _int(g("Độ sâu thư mục")),
            "response_ms": _int(g("Phản hồi (ms)"), None), "error": str(g("Lỗi")) or None,
        }
        r["noindex"] = "noindex" in (r["meta_robots"] or "").lower()
        r["final_url"] = r["url"]
        recs.append(r)
    for ws in wb.worksheets:  # sheet tong quan + orphan
        if ws.title == "Tổng quan":
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or row[0] is None:
                    continue
                k, v = str(row[0]), row[1] if len(row) > 1 else None
                if k == "Website":
                    meta["site"] = _host(str(v)) if "//" in str(v) else str(v)
                elif k.startswith("URL trong sitemap"):
                    meta["sitemap_count"] = _int(v)
                elif k.startswith("Orphan pages"):
                    meta["orphan_count"] = _int(v)
    wb.close()
    return recs, meta


def doc_site(path):
    # type: (str) -> Dict
    """Doc 1 website -> {"site", "records", "meta"}."""
    if path.lower().endswith(".json"):
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        meta = {}
        if isinstance(d, dict):
            meta = {k: d[k] for k in ("site", "sitemap_count", "orphan_count", "max_urls", "limited") if k in d}
            recs = d.get("records", [])
        else:
            recs = d
        for r in recs:
            r.setdefault("final_url", r.get("url"))
    else:
        recs, meta = _doc_xlsx(path)
    site = meta.get("site") or (_host(recs[0]["url"]) if recs else os.path.splitext(os.path.basename(path))[0])
    return {"site": site, "records": recs, "meta": meta}


# ---------------------------------------------------------------- luat tu cham
def _p(n, d):
    return 100.0 * n / d if d else 0.0


def _kq(ok, ev):
    return (DAT if ok else CHUA_DAT), ev


def _is_int(r):
    return isinstance(r.get("status"), int)


def _path(r):
    return urlparse(r["url"]).path or "/"


class Ctx(object):
    def __init__(self, recs, meta):
        self.recs = recs
        self.meta = meta
        self.n = len(recs)
        self.r200 = [r for r in recs if r.get("status") == 200]
        self.home = [r for r in self.r200 if r.get("page_type") == "homepage"]
        self.non_home = [r for r in self.r200 if r.get("page_type") != "homepage"]


def _by_type(c, t, label):
    rs = [r for r in c.r200 if r.get("page_type") == t]
    if not rs:
        return None
    bad = [r for r in rs if not (r.get("h1_count") == 1 and r.get("h2_count", 0) >= 1)]
    nguong = 100.0 if t == "homepage" else 90.0
    return _kq(_p(len(rs) - len(bad), len(rs)) >= nguong,
               "%d/%d trang %s chưa đúng 1 H1 + >= 1 H2" % (len(bad), len(rs), label))


def r_https(c):
    rs = [r for r in c.recs if _is_int(r)]
    if not rs:
        return None
    bad = [r for r in rs if not (r.get("final_url") or r["url"]).lower().startswith("https")]
    return _kq(_p(len(bad), len(rs)) <= 1, "%d/%d URL không dùng https" % (len(bad), len(rs)))


def r_host(c):
    if not c.recs:
        return None
    err = [r for r in c.recs if not _is_int(r) or r["status"] >= 500]
    ms = [r["response_ms"] for r in c.recs if r.get("response_ms")]
    med = int(statistics.median(ms)) if ms else 0
    ok = _p(len(err), c.n) <= 2 and med <= 1500
    return _kq(ok, "%d/%d URL lỗi kết nối/5xx; phản hồi trung vị %d ms" % (len(err), c.n, med))


def r_sitemap(c):
    n = c.meta.get("sitemap_count")
    if n is None:
        return None
    return _kq(n > 0, "Sitemap có %d URL" % n if n else "Không tìm thấy sitemap/không có URL")


def r_broken(c):
    if not c.recs:
        return None
    b404 = sum(1 for r in c.recs if r.get("status") == 404)
    b4 = sum(1 for r in c.recs if _is_int(r) and 400 <= r["status"] < 500 and r["status"] != 404)
    b5 = sum(1 for r in c.recs if _is_int(r) and r["status"] >= 500)
    be = sum(1 for r in c.recs if not _is_int(r))
    bad = b404 + b4 + b5 + be
    return _kq(_p(bad, c.n) <= 1, "%d/%d URL lỗi (404: %d, 4xx khác: %d, 5xx: %d, không truy cập: %d)"
               % (bad, c.n, b404, b4, b5, be))


def _url_rule(pred, nguong, mo_ta, dao=False):
    def f(c):
        if not c.recs:
            return None
        bad = [r for r in c.recs if pred(r)]
        pb = _p(len(bad), c.n)
        ok = pb <= nguong
        return _kq(ok, "%d/%d URL %s" % (len(bad), c.n, mo_ta))
    return f


def _path_unq(r):
    return unquote(_path(r))


r_url_param = _url_rule(lambda r: "?" in r["url"], 5, "có tham số (?)")
r_url_upper = _url_rule(lambda r: _path_unq(r) != _path_unq(r).lower(), 1, "có chữ hoa")
r_url_special = _url_rule(lambda r: not re.match(r"^[a-z0-9/_.\-]*$", _path_unq(r).lower()), 2,
                          "có dấu/ký tự đặc biệt")
r_url_length = _url_rule(lambda r: len(r["url"]) > 115, 5, "dài > 115 ký tự")
r_url_depth = _url_rule(lambda r: (r.get("folder_depth") or len([s for s in _path(r).split("/") if s])) > 4, 5,
                        "sâu > 4 cấp thư mục")
r_url_underscore = _url_rule(lambda r: "_" in _path(r), 1, "chứa gạch dưới")
r_url_ext = _url_rule(lambda r: re.search(r"\.(html?|php\d?|aspx?|jsp)$", _path(r).lower()) is not None, 5,
                      "có đuôi file")


def r_url_trailing(c):
    ps = [_path(r) for r in c.recs if _path(r) != "/"]
    if not ps:
        return None
    sl = sum(1 for p in ps if p.endswith("/"))
    ma = max(sl, len(ps) - sl)
    return _kq(_p(ma, len(ps)) >= 95, "%d/%d URL có dấu / cuối, %d không có" % (sl, len(ps), len(ps) - sl))


def r_internal_link(c):
    if not c.non_home:
        return None
    zero = [r for r in c.non_home if not r.get("inlink_count")]
    return _kq(_p(len(zero), len(c.non_home)) <= 5, "%d/%d trang không có inlink" % (len(zero), len(c.non_home)))


def r_title(c):
    rs = c.r200
    if not rs:
        return None
    miss = [r for r in rs if not r.get("title")]
    cnt = {}  # type: Dict[str, int]
    for r in rs:
        t = (r.get("title") or "").strip()
        if t:
            cnt[t] = cnt.get(t, 0) + 1
    dup = sum(v for v in cnt.values() if v > 1)
    withp = [r for r in rs if r.get("title")]
    lbad = [r for r in withp if not (30 <= r.get("title_length", 0) <= 65)]
    ok = _p(len(miss), len(rs)) <= 5 and _p(dup, len(rs)) <= 5 and (not withp or _p(len(lbad), len(withp)) <= 20)
    return _kq(ok, "Thiếu title %d/%d; trùng %d/%d; ngoài 30-65 ký tự %d/%d"
               % (len(miss), len(rs), dup, len(rs), len(lbad), len(withp)))


def r_meta(c):
    rs = c.r200
    if not rs:
        return None
    miss = [r for r in rs if not r.get("meta_description")]
    cnt = {}  # type: Dict[str, int]
    for r in rs:
        t = (r.get("meta_description") or "").strip()
        if t:
            cnt[t] = cnt.get(t, 0) + 1
    dup = sum(v for v in cnt.values() if v > 1)
    withp = [r for r in rs if r.get("meta_description")]
    lbad = [r for r in withp if r.get("meta_desc_length", 0) > 160]
    ok = _p(len(miss), len(rs)) <= 10 and _p(dup, len(rs)) <= 5 and (not withp or _p(len(lbad), len(withp)) <= 10)
    return _kq(ok, "Thiếu meta %d/%d; trùng %d/%d; dài > 160 ký tự %d/%d"
               % (len(miss), len(rs), dup, len(rs), len(lbad), len(withp)))


def r_noindex(c):
    if not c.r200:
        return None
    ni = [r for r in c.r200 if r.get("noindex")]
    return _kq(_p(len(ni), len(c.r200)) <= 5, "%d/%d trang 200 có noindex" % (len(ni), len(c.r200)))


def r_og(c):
    if not c.r200:
        return None
    no = [r for r in c.r200 if not r.get("og_image")]
    return _kq(_p(len(no), len(c.r200)) <= 20, "%d/%d trang thiếu og:image" % (len(no), len(c.r200)))


def r_canonical(c):
    rs = [r for r in c.r200 if r.get("canonical_ok") is not None]
    if not rs:
        return None
    bad = [r for r in rs if r["canonical_ok"] is False]
    return _kq(_p(len(bad), len(rs)) <= 5, "%d/%d trang canonical sai/thiếu" % (len(bad), len(rs)))


def r_canonical_dyn(c):
    if not c.r200:
        return None
    dyn = [r for r in c.r200 if "?" in r["url"]]
    if not dyn:
        return _kq(True, "Không có URL động (có ?) nào trong dữ liệu crawl")
    good = [r for r in dyn if r.get("canonical") and r.get("canonical_ok") is False]
    return _kq(_p(len(good), len(dyn)) >= 90, "%d/%d URL động có canonical trỏ về URL gốc" % (len(good), len(dyn)))


def r_alt(c):
    tot = sum(r.get("image_count", 0) for r in c.r200)
    if not tot:
        return None
    miss = sum(r.get("images_missing_alt", 0) for r in c.r200)
    pages = sum(1 for r in c.r200 if r.get("images_missing_alt", 0) > 0)
    return _kq(_p(miss, tot) <= 5, "%d/%d ảnh thiếu alt (%d trang)" % (miss, tot, pages))


def r_redirect_chain(c):
    if not c.recs:
        return None
    ch = sum(1 for r in c.recs if len(r.get("redirect_chain") or []) >= 2)
    one = sum(1 for r in c.recs if len(r.get("redirect_chain") or []) == 1)
    return _kq(_p(ch, c.n) <= 2, "%d/%d URL chuyển hướng >= 2 bước (%d URL chuyển hướng 1 bước)" % (ch, c.n, one))


def r_thin(c):
    if not c.non_home:
        return None
    th = [r for r in c.non_home if r.get("word_count", 0) < 300]
    return _kq(_p(len(th), len(c.non_home)) <= 20, "%d/%d trang dưới 300 từ" % (len(th), len(c.non_home)))


def r_click_depth(c):
    if not c.r200:
        return None
    deep = [r for r in c.r200 if r.get("depth", 0) > 3]
    return _kq(_p(len(deep), len(c.r200)) <= 10, "%d/%d trang cách trang chủ > 3 click" % (len(deep), len(c.r200)))


def r_orphan(c):
    o, s = c.meta.get("orphan_count"), c.meta.get("sitemap_count")
    if o is None or not s or c.meta.get("limited"):
        return None  # crawl bi gioi han: orphan khong dang tin
    return _kq(_p(o, s) <= 5, "%d/%d URL trong sitemap không có link nội bộ" % (o, s))


RULES = {
    "https": r_https, "host": r_host, "sitemap": r_sitemap, "broken": r_broken,
    "url_param": r_url_param, "url_upper": r_url_upper, "url_special": r_url_special,
    "url_length": r_url_length, "url_depth": r_url_depth, "url_underscore": r_url_underscore,
    "url_ext": r_url_ext, "url_trailing": r_url_trailing, "internal_link": r_internal_link,
    "title": r_title, "meta": r_meta, "noindex": r_noindex, "og": r_og, "canonical": r_canonical,
    "canonical_dyn": r_canonical_dyn,
    "h_home": lambda c: _by_type(c, "homepage", "chủ"),
    "h_product": lambda c: _by_type(c, "product", "sản phẩm"),
    "h_category": lambda c: _by_type(c, "category", "danh mục"),
    "h_blog": lambda c: _by_type(c, "blog", "blog"),
    "h_other": lambda c: _by_type(c, "other", "loại khác"),
    "alt": r_alt, "redirect_chain": r_redirect_chain, "thin": r_thin, "click_depth": r_click_depth,
    "orphan": r_orphan,
}


def doc_tieuchi():
    with open(TIEUCHI_PATH, encoding="utf-8") as f:
        return json.load(f)


def cham_site(site, tieu_chi):
    # type: (Dict, List[Dict]) -> Dict[str, Optional[Tuple[str, str]]]
    c = Ctx(site["records"], site["meta"])
    out = {}
    for t in tieu_chi:
        fn = RULES.get(t.get("auto")) if t.get("auto") else None
        out[t["id"]] = fn(c) if fn else None
    return out


# ---------------------------------------------------------------- sheet du lieu an
DATA_SHEET = "Onpage data"
DATA_COLS = [("Website", None), ("URL", "url"), ("Loại trang", "page_type"), ("Status", "status"),
             ("Title", "title"), ("Dài title", "title_length"), ("Meta description", "meta_description"),
             ("Dài meta", "meta_desc_length"), ("Số H1", "h1_count"), ("Số H2", "h2_count"),
             ("Canonical", "canonical"), ("Canonical đúng", "canonical_ok"), ("Noindex", "noindex"),
             ("og:image", "og_image"), ("Số ảnh", "image_count"), ("Ảnh thiếu alt", "images_missing_alt"),
             ("Số từ", "word_count"), ("Inlinks", "inlink_count"), ("Depth", "depth"), ("Phản hồi (ms)", "response_ms")]
DATA_COL_DESC = [("Website", "Website"), ("URL", "URL đã crawl"),
                 ("Loại trang", "homepage/category/product/blog/other"), ("Status", "HTTP status (chữ = lỗi kết nối)"),
                 ("Title", "Thẻ title"), ("Dài title", "Số ký tự title"), ("Meta description", "Meta description"),
                 ("Dài meta", "Số ký tự meta"), ("Số H1", "Số thẻ H1"), ("Số H2", "Số thẻ H2"), ("Canonical", "URL canonical"),
                 ("Canonical đúng", "Có/Không/trống"), ("Noindex", "Có nếu meta robots noindex"), ("og:image", "Ảnh og"),
                 ("Số ảnh", "Số ảnh"), ("Ảnh thiếu alt", "Số ảnh thiếu alt"), ("Số từ", "Số từ nội dung"),
                 ("Inlinks", "Số link nội bộ trỏ đến"), ("Depth", "Số click từ trang chủ"), ("Phản hồi (ms)", "Thời gian phản hồi")]


def _gia_tri(r, k):
    v = r.get(k)
    if isinstance(v, bool):
        return "Có" if v else "Không"
    if isinstance(v, (list, dict)):
        return json.dumps(v, ensure_ascii=False)[:500]
    if isinstance(v, str) and len(v) > 1000:
        return v[:1000]
    return v


# ---------------------------------------------------------------- build
def build(wb, inputs, opts, ctx):
    opts = opts or {}
    ctx = ctx or {}
    cfg = doc_tieuchi()
    groups = cfg["nhom_diem"]
    tcs = [t for t in cfg["tieu_chi"] if not t.get("skip")]
    # xep lien tuc theo nhom (stable)
    tcs.sort(key=lambda t: groups.index(t["nhom"]) if t["nhom"] in groups else 99)
    data = [doc_site(p) for p in inputs]
    names = opts.get("sites") or None
    if not names and ctx.get("sites") and len(ctx["sites"]) == len(inputs):
        names = common.ten_site_theo_file(inputs, ctx["sites"])
    if names and len(names) == len(inputs):
        for d, nm in zip(data, names):
            d["site"] = nm
    seen = {}
    for d in data:
        k = seen.get(d["site"], 0)
        seen[d["site"]] = k + 1
        if k:
            d["site"] = "%s (%d)" % (d["site"], k + 1)
    sites = [d["site"] for d in data]
    N = len(sites)
    kq = [cham_site(d, tcs) for d in data]

    ws = wb.create_sheet(SHEET)
    # ---- sheet du lieu an: ban ghi crawl
    dws = wb.create_sheet(DATA_SHEET)
    common.header_row(dws, 1, 1, [c[0] for c in DATA_COLS])
    n_data = 1
    for d in data:
        for r in d["records"]:
            n_data += 1
            dws.append([d["site"]] + [_gia_tri(r, k) for _n, k in DATA_COLS[1:]])
    dws.freeze_panes = "B2"
    common.title(ws, "ĐÁNH GIÁ ONPAGE THEO TIÊU CHÍ (tool tự chấm phần đo được từ crawl, phần còn lại để Claude chấm)",
                 "Nguồn: crawl Mini Frog (sheet ẩn '%s'). Điểm nhóm = Đạt / (Đạt + Chưa đạt); ô trống, 'Chưa kiểm' "
                 "không tính. Tiêu chí CMS, GSC, GA4 đã loại." % DATA_SHEET)
    n_tc = len(tcs)
    L = get_column_letter
    r_diem = 4
    r_tot = r_diem + 1 + len(groups)
    r_dc = r_tot + 2                  # op_da_cham
    hdr = r_dc + 4                    # tieu de bang tieu chi
    r_first, r_last = hdr + 1, hdr + n_tc
    c_eval0 = 6                       # F
    c_evid0 = c_eval0 + N
    c_ng = c_evid0 + N
    c_who = c_ng + 1

    def rng(col):
        return "$%s$%d:$%s$%d" % (col, r_first, col, r_last)

    # ---- op_diem
    common.header_row(ws, r_diem, 1, ["Nhóm"] + sites)
    so = {"op_diem_tong": []}
    for i, g in enumerate(groups + [None]):
        r = r_diem + 1 + i
        ws.cell(r, 1, g if g else "Tổng (mọi tiêu chí đã chấm)").border = common.BOX
        for j in range(N):
            col = L(c_eval0 + j)
            if g is None:
                d_ = 'COUNTIF(%s,"%s")' % (rng(col), DAT)
                x_ = 'COUNTIF(%s,"%s")' % (rng(col), CHUA_DAT)
            else:
                d_ = 'COUNTIFS(%s,$A%d,%s,"%s")' % (rng("A"), r, rng(col), DAT)
                x_ = 'COUNTIFS(%s,$A%d,%s,"%s")' % (rng("A"), r, rng(col), CHUA_DAT)
            c = ws.cell(r, 2 + j, '=IFERROR(%s/(%s+%s),"")' % (d_, d_, x_))
            c.number_format = "0.0%"
            c.border = common.BOX
            if g is None:
                c.font = Font(bold=True)
                so["op_diem_tong"].append(common.so_cell(ws, "%s%d" % (L(2 + j), r), sites[j]))
    sums = [common.bang(wb, ws, "op_diem", "A%d:%s%d" % (r_diem, L(1 + N), r_tot),
                        "Điểm onpage theo nhóm × website (Đạt / (Đạt + Chưa đạt)); dòng cuối = tổng.",
                        total_rows=1, name="Onpage_Diem")]
    # ---- op_da_cham
    common.header_row(ws, r_dc, 1, ["Tiến độ chấm"] + sites)
    ws.cell(r_dc + 1, 1, "% tiêu chí đã chấm").border = common.BOX
    ws.cell(r_dc + 2, 1, "Số tiêu chí chưa chấm").border = common.BOX
    for j in range(N):
        col = L(c_eval0 + j)
        d_ = 'COUNTIF(%s,"%s")' % (rng(col), DAT)
        x_ = 'COUNTIF(%s,"%s")' % (rng(col), CHUA_DAT)
        c = ws.cell(r_dc + 1, 2 + j, "=(%s+%s)/ROWS(%s)" % (d_, x_, rng(col)))
        c.number_format = "0%"
        c.border = common.BOX
        ws.cell(r_dc + 2, 2 + j, "=ROWS(%s)-%s-%s" % (rng(col), d_, x_)).border = common.BOX
    sums.append(common.bang(wb, ws, "op_da_cham", "A%d:%s%d" % (r_dc, L(1 + N), r_dc + 2),
                            "Tiến độ chấm: % tiêu chí đã chấm (Đạt/Chưa đạt) và số tiêu chí còn trống, theo website."))
    notes = [common.note_area(ws, "%s5" % L(N + 3), 6, 7, "NHẬN XÉT")]

    # ---- op_tieu_chi
    heads = ["Nhóm", "Nhóm con", "Vấn đề cần kiểm tra", "Vì sao cần làm", "Yêu cầu đầu ra"]
    heads += ["Đánh giá %s" % s for s in sites] + ["Bằng chứng %s" % s for s in sites] + ["Ngưỡng tự chấm", "Người chấm"]
    common.header_row(ws, hdr, 1, heads)
    dv = DataValidation(type="list", formula1='"%s,%s,Chưa kiểm"' % (DAT, CHUA_DAT), allow_blank=True)
    ws.add_data_validation(dv)
    wrap = Alignment(wrap_text=True, vertical="top")
    claude_cells, claude_items, claude_pairs = [], [], {}
    auto_cnt = [0] * N
    q = common.quote_sheet(SHEET)
    for i, t in enumerate(tcs):
        r = r_first + i
        vals = [t["nhom"], t["nhom_con"], t["van_de"], t["vi_sao"], t["yeu_cau"]]
        for j, v in enumerate(vals):
            c = ws.cell(r, 1 + j, v)
            c.alignment = wrap
            c.border = common.BOX
        empty_here = []
        for j in range(N):
            res = kq[j].get(t["id"])
            ce = ws.cell(r, c_eval0 + j)
            cv = ws.cell(r, c_evid0 + j)
            ce.border = cv.border = common.BOX
            ce.alignment = Alignment(horizontal="center", vertical="top")
            cv.alignment = wrap
            dv.add(ce.coordinate)
            if res:
                ce.value, cv.value = res[0], res[1]
                auto_cnt[j] += 1
            else:
                ref = "%s!%s" % (q, ce.coordinate)
                empty_here.append(ref)
                claude_pairs[ref] = "%s!%s" % (q, cv.coordinate)
        ws.cell(r, c_ng, t.get("nguong") or ("Claude chấm tay" if not t.get("auto") else "")).alignment = wrap
        n_auto = N - len(empty_here)
        who = "Tool" if not empty_here else ("Claude" if n_auto == 0 else "Tool + Claude")
        c = ws.cell(r, c_who, who)
        c.border = ws.cell(r, c_ng).border = common.BOX
        c.alignment = Alignment(horizontal="center", vertical="top")
        if empty_here:
            claude_cells += empty_here
            claude_items.append({"row": r, "van_de": t["van_de"], "yeu_cau": (t["yeu_cau"] or "")[:200],
                                 "cells": empty_here})
    green = PatternFill("solid", bgColor="C6EFCE", fgColor="C6EFCE")
    red = PatternFill("solid", bgColor="FFC7CE", fgColor="FFC7CE")
    area = "%s%d:%s%d" % (L(c_eval0), r_first, L(c_eval0 + N - 1), r_last)
    ws.conditional_formatting.add(area, CellIsRule(operator="equal", formula=['"%s"' % DAT], fill=green))
    ws.conditional_formatting.add(area, CellIsRule(operator="equal", formula=['"%s"' % CHUA_DAT], fill=red))
    for col, w in zip("ABCDE", [20, 18, 38, 45, 45]):
        ws.column_dimensions[col].width = w
    for j in range(N):
        ws.column_dimensions[L(c_eval0 + j)].width = 14
        ws.column_dimensions[L(c_evid0 + j)].width = 42
    ws.column_dimensions[L(c_ng)].width = 30
    ws.column_dimensions[L(c_who)].width = 14
    ws.freeze_panes = "D%d" % r_first
    sums.append(common.bang(wb, ws, "op_tieu_chi", "A%d:%s%d" % (hdr, L(c_who), r_last),
                            "Bảng tiêu chí: Nhóm, Nhóm con, Vấn đề, Vì sao, Yêu cầu, Đánh giá từng website (Đạt/Chưa đạt/"
                            "Chưa kiểm, trống = chờ Claude), Bằng chứng từng website, Ngưỡng tự chấm, Người chấm.",
                            name="Onpage_TieuChi"))
    validations = [{"range": "%s!%s" % (q, area), "values": [DAT, CHUA_DAT, "Chưa kiểm"]}]

    # ---- caveats
    caveats = []
    n_active = len(tcs)
    for j, d in enumerate(data):
        nrec = len(d["records"])
        m = d["meta"]
        lim = m.get("limited")
        mx = opts.get("max_urls") or m.get("max_urls")
        if lim is None and mx:
            lim = nrec >= int(mx)
        txt = "%s: đã crawl %d URL" % (d["site"], nrec)
        if lim is True:
            txt += ", crawl BỊ GIỚI HẠN số URL (kết quả chỉ là mẫu, không phải toàn site)"
        elif lim is False:
            txt += ", crawl chạy hết (không chạm giới hạn)"
        else:
            txt += ", không rõ crawl có bị giới hạn không (truyền opts.max_urls hoặc JSON có 'limited' để biết)"
        caveats.append(txt + ".")
        caveats.append("%s: tool tự chấm %d/%d tiêu chí (%.0f%%), còn %d tiêu chí chờ Claude."
                       % (d["site"], auto_cnt[j], n_active, _p(auto_cnt[j], n_active), n_active - auto_cnt[j]))
        if m.get("sitemap_count") is None:
            caveats.append("%s: không có thông tin sitemap/orphan (đầu vào JSON) nên tiêu chí sitemap, orphan để Claude." % d["site"])
        if nrec == 0:
            caveats.append("%s: file không có URL nào." % d["site"])
    caveats.append("Chưa đo được từ crawl: robots.txt, schema, breadcrumb, viewport, charset, hreflang, favicon, tốc độ, "
                   "giao diện và chức năng; các dòng này để trống chờ Claude/người chấm.")

    n_claude_rows = len(claude_items)
    return common.entry(
        SHEET, KEY, "Chấm tiêu chí onpage (nhóm, đạt/chưa đạt kèm bằng chứng) cho từng website từ crawl Mini Frog",
        inputs, data_sheet=DATA_SHEET, rows=n_data - 1,
        columns=[{"name": n_, "desc": d_} for n_, d_ in DATA_COL_DESC],
        summaries=sums, options=[], caveats=caveats, notes=notes,
        extra={"sites": sites, "claude_cells": claude_cells, "claude_items": claude_items,
               "claude_pairs": claude_pairs, "validations": validations, "so": so,
               "charts": [{"key": "op_diem", "bang": "op_diem", "loai": "radar", "tieu_de": "{{ten}}"}],
               "du_lieu": [{"key": "op_data", "sheet": DATA_SHEET, "site_col": "Website"}],
               "criteria_header_row": hdr, "criteria_rows": [r_first, r_last],
               "stats": {"tong_tieu_chi": n_tc, "auto_rule": sum(1 for t in tcs if t.get("auto")),
                         "claude_rows": n_claude_rows}})


# ---------------------------------------------------------------- CLI
def main(argv=None):
    ap = argparse.ArgumentParser(description=DESC)
    ap.add_argument("files", nargs="+", help="file xlsx xuất của Mini Frog hoặc JSON records (mỗi website một file)")
    ap.add_argument("--out", required=True)
    ap.add_argument("--max-urls", type=int, default=None)
    a = ap.parse_args(argv)
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    e = build(wb, a.files, {"max_urls": a.max_urls} if a.max_urls else {}, {})
    wb.save(a.out)
    print(json.dumps({k: v for k, v in e.items() if k != "claude_items"}, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
