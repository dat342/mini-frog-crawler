"""Tong quan doi thu tu Ahrefs Batch Analysis (Export CSV): DR, ref domains, organic traffic/keywords, traffic value.

Dung (CLI):
    .venv/bin/python -m xu_ly_data.batch_overview --out kq.xlsx [--sites a.com,b.com] batch-analysis.csv

- Doc CSV UTF-8/UTF-16, phay/tab; ten cot linh hoat (Target, Domain Rating, Ref. domains / Ref domains Dofollow /
  Referring domains, Organic traffic, Organic keywords / Total Keywords, Traffic value / Organic Traffic Value...).
- Sheet an 'Batch data': moi dong 1 target (gia tri goc).
- Sheet 'Tổng quan đối thủ': bang bo_tong_quan (site x chi so) bang INDEX/MATCH tro 'Batch data'.
Khoa (CONTRACT 8.2): bang bo_tong_quan; bieu_do bo_traffic (bar); du_lieu bo_data; so bo_traffic, bo_dr (theo site).
Python 3.9: khong dung cu phap X | None.
"""
import argparse
import csv
import io
import json
import re
import sys
from typing import Dict, List, Optional, Tuple

from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from . import common

KEY = "batch_overview"
SHEET = "Tổng quan đối thủ"
INPUT_KIND = "ahrefs_batch_csv"
DESC = "Tổng quan đối thủ: DR, referring domains, organic traffic/keywords, traffic value (Ahrefs Batch Analysis)"

DATA_SHEET = "Batch data"
# (ma, ten hien thi, mo ta, dinh dang)
CHI_SO = [
    ("dr", "DR", "Domain Rating 0–100", "0"),
    ("rd", "Ref. domains", "Số referring domains", "#,##0"),
    ("traffic", "Organic traffic", "Organic traffic ước tính/tháng", "#,##0"),
    ("kw", "Organic keywords", "Số từ khoá organic", "#,##0"),
    ("value", "Traffic value", "Giá trị traffic organic (USD/tháng)", "#,##0"),
]


# ---------------------------------------------------------------- tien ich
def _k(s):
    return re.sub(r"\s+", " ", re.sub(r"[^0-9a-z]+", " ", str(s or "").lower())).strip()


def _num(v):
    # type: (object) -> float
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("$", "").replace(" ", "").replace(" ", "")
    if not s or s in ("-", "—", "n/a", "N/A"):
        return 0.0
    m = re.match(r"^([\d.,]+)([kKmM]?)$", s)
    if not m:
        return 0.0
    num, unit = m.group(1), m.group(2).lower()
    if unit:
        num = num.replace(",", ".")
    else:
        num = num.replace(",", "")
    try:
        return float(num) * {"k": 1e3, "m": 1e6}.get(unit, 1)
    except ValueError:
        return 0.0


def chuan_domain(s):
    # type: (str) -> str
    s = str(s or "").strip().lower()
    s = re.sub(r"^[a-z]+://", "", s)
    s = re.sub(r"^www\.", "", s)
    return s.split("/")[0] if s.count("/") <= 1 else s.rstrip("/")


def _doc_csv(path):
    # type: (str) -> Tuple[List[Dict[str, str]], List[str]]
    with open(path, "rb") as f:
        raw = f.read()
    if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
        txt = raw.decode("utf-16")
    elif raw[:3] == b"\xef\xbb\xbf":
        txt = raw[3:].decode("utf-8")
    elif len(raw) > 1 and raw[1:2] == b"\x00":
        txt = raw.decode("utf-16-le")
    else:
        txt = raw.decode("utf-8", errors="replace")
    first = txt.split("\n", 1)[0]
    delim = "\t" if first.count("\t") > first.count(",") else ","
    rows = list(csv.reader(io.StringIO(txt), delimiter=delim))
    h = 0
    for i, r in enumerate(rows[:10]):
        if any(_k(c) in ("target", "domain", "url", "website") for c in r):
            h = i
            break
    hdr = [c.strip() for c in rows[h]] if rows else []
    out = []
    for r in rows[h + 1:]:
        if not any(str(c).strip() for c in r):
            continue
        out.append({hdr[j]: r[j] for j in range(min(len(hdr), len(r)))})
    return out, hdr


def _chon_cot(hdr, exact, need=(), avoid=()):
    # type: (List[str], Tuple[str, ...], Tuple[str, ...], Tuple[str, ...]) -> Optional[str]
    """Cot co ten chuan hoa trung exact (theo thu tu uu tien); khong co thi cot dau tien chua du moi
    tu trong need va khong chua tu nao trong avoid."""
    norm = {h: _k(h) for h in hdr}
    for e in exact:
        for h in hdr:
            if norm[h] == e:
                return h
    if need:
        for h in hdr:
            toks = norm[h].split()
            if all(t in toks for t in need) and not any(t in toks for t in avoid):
                return h
    return None


def map_cot(hdr):
    # type: (List[str]) -> Dict[str, Optional[str]]
    return {
        "target": _chon_cot(hdr, ("target", "domain", "url", "website", "site")),
        "mode": _chon_cot(hdr, ("mode",)),
        "dr": _chon_cot(hdr, ("domain rating", "dr")),
        "rd": _chon_cot(hdr, ("ref domains", "referring domains", "ref domains all", "ref domains total",
                              "referring domains all", "ref domains dofollow", "referring domains dofollow"),
                        ("domains", "ref"), ("gov", "governmental", "edu", "educational", "nofollow", "lost", "new")),
        "traffic": _chon_cot(hdr, ("organic traffic", "organic total traffic", "organic traffic total",
                                   "organic search traffic", "traffic"),
                             ("organic", "traffic"), ("value", "cost", "paid", "top", "change", "country")),
        "kw": _chon_cot(hdr, ("organic keywords", "total keywords", "organic total keywords",
                              "organic keywords total", "keywords"),
                        ("keywords",), ("paid", "top", "1", "3", "4", "10", "11", "100", "change", "position")),
        "value": _chon_cot(hdr, ("traffic value", "organic traffic value", "organic value", "organic cost",
                                 "organic traffic cost", "value"),
                           ("value",), ("paid",)),
    }


def doc_file(path):
    # type: (str) -> Dict
    rows, hdr = _doc_csv(path)
    m = map_cot(hdr)
    if not m["target"]:
        raise ValueError("%s: thiếu cột Target — không phải file Batch Analysis? Cột: %s" % (path, hdr))
    out = []
    for r in rows:
        t = (r.get(m["target"]) or "").strip()
        if not t or _k(t) in ("total", "tong"):
            continue
        rec = {"target": t, "site": chuan_domain(t), "mode": (r.get(m["mode"]) or "").strip() if m["mode"] else ""}
        for c, _n, _d, _f in CHI_SO:
            rec[c] = _num(r.get(m[c])) if m[c] else 0.0
        out.append(rec)
    return {"rows": out, "hdr": hdr, "map": m}


# ---------------------------------------------------------------- build
def build(wb, inputs, opts, ctx):
    opts = opts or {}
    ctx = ctx or {}
    files = [doc_file(p) for p in inputs]
    recs, seen = [], {}
    for f in files:
        for r in f["rows"]:
            if r["site"] in seen:
                continue
            seen[r["site"]] = r
            recs.append(r)
    sites_cfg = opts.get("sites") or ctx.get("sites") or []
    if isinstance(sites_cfg, str):
        sites_cfg = [s.strip() for s in sites_cfg.split(",") if s.strip()]
    if sites_cfg:
        sites = list(sites_cfg)
        thua = [r["target"] for r in recs if r["site"] not in {chuan_domain(s) for s in sites}]
    else:
        sites = [r["site"] for r in recs]
        thua = []
    # nhan dung trong bang = domain dung nhu recipe; trong Batch data ghi cot Website = nhan nay
    nhan_cua = {}
    for s in sites:
        nhan_cua[chuan_domain(s)] = s
    for r in recs:
        r["website"] = nhan_cua.get(r["site"], r["site"])

    # ---------- Batch data (an) ----------
    dws = wb.create_sheet(DATA_SHEET)
    hdr = ["Website", "Target", "Mode"] + [n for _c, n, _d, _f in CHI_SO]
    common.header_row(dws, 1, 1, hdr)
    for r in recs:
        dws.append([r["website"], r["target"], r["mode"]] + [r[c] for c, _n, _d, _f in CHI_SO])
    last = max(len(recs) + 1, 2)
    for i, w in enumerate([24, 30, 12, 8, 14, 16, 16, 14], 1):
        dws.column_dimensions[get_column_letter(i)].width = w
    for j, (_c, _n, _d, fm) in enumerate(CHI_SO):
        for cell in dws[get_column_letter(4 + j)][1:]:
            cell.number_format = fm
    dws.freeze_panes = "B2"
    dws.sheet_state = "hidden"
    D = lambda c: "'%s'!$%s$2:$%s$%d" % (DATA_SHEET, c, c, last)

    # ---------- sheet ket qua ----------
    ws = wb.create_sheet(SHEET)
    q = common.quote_sheet(SHEET)
    common.title(ws, "TỔNG QUAN ĐỐI THỦ",
                 "Nguồn: Ahrefs › Batch Analysis (Export CSV). Bảng dưới là hàm INDEX/MATCH trỏ sheet ẩn '%s'."
                 % DATA_SHEET)
    row = 4
    ws.cell(row, 1, "(1) Chỉ số tổng quan theo website").font = Font(bold=True, color=common.XANH)
    common.header_row(ws, row + 1, 1, ["Website"] + [n for _c, n, _d, _f in CHI_SO])
    r0 = row + 2
    so = {"bo_traffic": [], "bo_dr": [], "bo_rd": [], "bo_keywords": [], "bo_value": []}  # type: Dict[str, List[Dict]]
    for i, s in enumerate(sites):
        r = r0 + i
        c = ws.cell(r, 1, s)
        c.border = common.BOX
        c.font = Font(bold=True) if i == 0 else Font()
        for j, (code, _n, _d, fm) in enumerate(CHI_SO):
            cl = get_column_letter(4 + j)
            f = "=IFERROR(INDEX(%s,MATCH($A%d,%s,0)),0)" % (D(cl), r, D("A"))
            cell = ws.cell(r, 2 + j, f)
            cell.border = common.BOX
            cell.number_format = fm
        so["bo_dr"].append({"site": s, "cell": "%s!B%d" % (q, r)})
        so["bo_traffic"].append({"site": s, "cell": "%s!D%d" % (q, r)})
        so["bo_rd"].append({"site": s, "cell": "%s!C%d" % (q, r)})
        so["bo_keywords"].append({"site": s, "cell": "%s!E%d" % (q, r)})
        so["bo_value"].append({"site": s, "cell": "%s!F%d" % (q, r)})
    r_end = r0 + max(len(sites), 1) - 1
    e = common.summary(wb, ws, "bo_tong_quan", "A%d:F%d" % (row + 1, r_end),
                       "Website × DR, Ref. domains, Organic traffic, Organic keywords, Traffic value")
    e["key"] = "bo_tong_quan"
    sums = [e]

    notes = [common.note_area(ws, "A%d" % (r_end + 4), 6, 8, "NHẬN XÉT")]
    ws.column_dimensions["A"].width = 28
    for j in range(len(CHI_SO)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 17

    ch = BarChart()
    ch.type = "bar"
    ch.title = "Organic traffic theo website"
    ch.legend = None
    ch.height, ch.width = max(7, 1.2 * len(sites) + 4), 16
    ch.add_data(Reference(ws, min_col=4, min_row=row + 1, max_row=r_end), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=r0, max_row=r_end))
    ch.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(ch, "H4")
    charts = [{"key": "bo_traffic", "loai": "bar", "bang": "bo_tong_quan", "cot": "Organic traffic",
               "anchor": "%s!H4" % q}]

    # ---------- caveats ----------
    caveats = []
    by = {chuan_domain(s): seen.get(chuan_domain(s)) for s in sites}
    co = [(s, by[chuan_domain(s)]) for s in sites if by[chuan_domain(s)]]
    thieu = [s for s in sites if not by[chuan_domain(s)]]
    if thieu:
        caveats.append("Không thấy trong file Batch Analysis: %s (bảng = 0)." % ", ".join(thieu))
    if thua:
        caveats.append("Target trong file nhưng không có trong danh sách site: %s (không đưa vào bảng)." % ", ".join(thua))
    if co:
        tong_tr = sum(r["traffic"] for _s, r in co)
        s_top, r_top = max(co, key=lambda x: x[1]["traffic"])
        if tong_tr > 0:
            caveats.append("Dẫn đầu organic traffic: %s với %s/tháng (%.1f%% tổng traffic các site)." % (
                s_top, format(int(r_top["traffic"]), ","), r_top["traffic"] / tong_tr * 100))
        s_hi, r_hi = max(co, key=lambda x: x[1]["dr"])
        s_lo, r_lo = min(co, key=lambda x: x[1]["dr"])
        caveats.append("DR cao nhất %s (%g), thấp nhất %s (%g), chênh %g điểm." % (
            s_hi, r_hi["dr"], s_lo, r_lo["dr"], r_hi["dr"] - r_lo["dr"]))
        for s, r in co:
            z = [n for c, n in (("traffic", "organic traffic"), ("kw", "organic keywords")) if r[c] == 0]
            if z:
                caveats.append("%s: %s = 0 (kiểm tra mode/target hoặc site chưa có organic)." % (s, ", ".join(z)))
        modes = {r["mode"] for _s, r in co if r["mode"]}
        if len(modes) > 1:
            caveats.append("Các target dùng mode khác nhau (%s): số liệu không cùng phạm vi." % ", ".join(sorted(modes)))
    for f in files:
        m = f["map"]
        miss = [n for c, n, _d, _f in CHI_SO if not m[c]]
        if miss:
            caveats.append("File thiếu cột %s: ghi 0." % ", ".join(miss))
        if m["rd"] and "dofollow" in _k(m["rd"]):
            caveats.append("Ref. domains lấy từ cột '%s' (chỉ dofollow)." % m["rd"])

    columns = [{"name": "Website", "desc": "Domain (khớp recipe sites nếu có)"},
               {"name": "Target", "desc": "Target gốc trong file Ahrefs"},
               {"name": "Mode", "desc": "Chế độ Batch Analysis (subdomains/domain/prefix/exact)"}] + \
              [{"name": n, "desc": d} for _c, n, d, _f in CHI_SO]
    return common.entry(
        SHEET, KEY, "So sánh chỉ số tổng quan Ahrefs (DR, ref domains, organic traffic/keywords, traffic value) giữa các website",
        inputs, data_sheet=DATA_SHEET, rows=len(recs), columns=columns, summaries=sums,
        options=[], caveats=caveats, notes=notes,
        extra={"sites": sites, "so": so, "charts": charts, "du_lieu": {"key": "bo_data", "sheet": DATA_SHEET},
               "cot_nguon": [f["map"] for f in files]})


# ---------------------------------------------------------------- CLI
def main(argv=None):
    ap = argparse.ArgumentParser(description=DESC)
    ap.add_argument("files", nargs="+")
    ap.add_argument("--sites", help="danh sách domain, cách nhau dấu phẩy (thứ tự trong bảng)")
    ap.add_argument("--out", required=True)
    a = ap.parse_args(argv)
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    e = build(wb, a.files, {"sites": a.sites} if a.sites else {}, {})
    wb.save(a.out)
    print(json.dumps(e, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    sys.exit(main())
