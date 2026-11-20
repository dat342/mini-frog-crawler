"""Xu ly file Referring domains (Ahrefs export) cho N website -> sheet danh gia backlink bang ham.

Dung:
    .venv/bin/python -m xu_ly_data.backlink_rd --out ket-qua.xlsx a-refdomains.csv b-refdomains.csv ...

- Sheet "RD data": du lieu tho da chuan hoa, xep doc, co cot Website -> them site chi can them file.
- Sheet "1.9 Backlink": bang theo nam x nhom chat luong (SUMIFS/COUNTIFS luy ke), 2 o tuy chon
  (chi so: So refdomain / Links to target; pham vi: chi domain con link / ca domain da mat),
  1 bieu do cot chong moi site (chieu cao = tong, mau = ty le chat luong), bang so sanh cac site.
Python 3.9: khong dung cu phap X | None.
"""
import argparse
import csv
import datetime as dt
import os
import re
from typing import List, Optional

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

try:
    from . import common
except ImportError:  # chay truc tiep
    import common  # type: ignore

KEY = "backlink_rd"
SHEET = "1.9 Backlink"
INPUT_KIND = "ahrefs_refdomains_csv"
DESC = "Đánh giá backlink theo Referring domains"
DATA_SHEET = "RD data"

NHOM = ["Spam", "0", "1-30", "30-50", "50-70", "70-100"]
MAU_NHOM = ["C0504D", "BFBFBF", "F2C894", "9DC3E6", "4A90D9", "1F3864"]
CHI_SO = ["Số refdomain", "Links to target"]
PHAM_VI = ["Chỉ domain còn link", "Cả domain đã mất"]
XANH = "0B3D91"
HEADER_FILL = PatternFill("solid", fgColor="DCE6F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _date(s):
    # type: (str) -> Optional[dt.datetime]
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d %b %Y", "%A, %d %B %Y"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def _num(s):
    try:
        return float(str(s).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def _site_tu_ten_file(path):
    name = os.path.basename(path)
    m = re.match(r"(.+?)-refdomains", name)
    return m.group(1) if m else os.path.splitext(name)[0]


def doc_file(path, site=None):
    rows = list(csv.DictReader(open(path, encoding="utf-8-sig")))
    if not rows:
        return site or _site_tu_ten_file(path), []
    cols = {k.strip().lower(): k for k in rows[0].keys()}

    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None

    c_dom, c_spam, c_dr = col("domain", "referring domain"), col("is spam", "spam"), col("dr", "domain rating")
    c_ltt, c_df = col("links to target"), col("dofollow links")
    c_tr, c_fs, c_lost = col("traffic", "domain traffic"), col("first seen"), col("lost")
    if not (c_dom and c_dr and c_fs):
        raise ValueError("%s: thiếu cột Domain/DR/First seen — không phải file Referring domains?" % path)
    out = []
    for r in rows:
        out.append([
            r[c_dom],
            (r.get(c_spam, "") or "").strip().lower() == "true" if c_spam else False,
            _num(r[c_dr]),
            _num(r.get(c_ltt, 0)) if c_ltt else 1,
            _num(r.get(c_df, 0)) if c_df else 0,
            _num(r.get(c_tr, 0)) if c_tr else 0,
            _date(r[c_fs]),
            _date(r.get(c_lost, "")) if c_lost else None,
        ])
    return site or _site_tu_ten_file(path), out


COLUMNS = [
    {"name": "Website", "desc": "Website (domain) mà refdomain trỏ về, lấy từ tên file"},
    {"name": "Domain", "desc": "Referring domain"},
    {"name": "Spam", "desc": "'Spam' nếu Ahrefs gắn nhãn spam, trống nếu không"},
    {"name": "DR", "desc": "Domain Rating 0-100 của refdomain"},
    {"name": "Links to target", "desc": "Số link từ domain này tới website"},
    {"name": "Dofollow links", "desc": "Số link dofollow"},
    {"name": "Traffic", "desc": "Organic traffic ước tính của refdomain"},
    {"name": "First seen", "desc": "Ngày Ahrefs thấy link đầu tiên"},
    {"name": "Lost", "desc": "Ngày mất link (trống = còn link)"},
    {"name": "Trạng thái", "desc": "Còn/Mất (giá trị, từ cột Lost)"},
    {"name": "Nhóm chất lượng", "desc": "Spam, 0, 1-30, 30-50, 50-70, 70-100 (theo Spam rồi DR)"},
    {"name": "Năm first seen", "desc": "Năm của First seen"},
]


def _caveats(sites, nams):
    out = []
    for site, rows in sites:
        n = len(rows)
        if not n:
            out.append("%s: file không có dòng nào." % site)
            continue
        spam = sum(1 for d in rows if d[1])
        lost = sum(1 for d in rows if d[7])
        out.append("%s: %.0f%% refdomain bị Ahrefs gắn spam (%d/%d)." % (site, 100.0 * spam / n, spam, n))
        out.append("%s: %.0f%% refdomain đã mất link (%d/%d)." % (site, 100.0 * lost / n, lost, n))
        by_year = {}
        for d in rows:
            if d[6]:
                by_year[d[6].year] = by_year.get(d[6].year, 0) + 1
        tot_y = sum(by_year.values())
        if tot_y:
            y, c = max(by_year.items(), key=lambda kv: kv[1])
            if c * 2 > tot_y:
                out.append("%s: năm %d chiếm %.0f%% số refdomain mới (%d/%d), tăng trưởng dồn vào một năm."
                           % (site, y, 100.0 * c / tot_y, c, tot_y))
        ltt = sum(d[3] for d in rows)
        if ltt:
            top = max(rows, key=lambda d: d[3])
            if top[3] / ltt > 0.2:
                out.append("%s: domain %s chiếm %.0f%% Links to target (%d/%d), kéo lệch chỉ số Links to target."
                           % (site, top[0], 100.0 * top[3] / ltt, top[3], ltt))
    return out


def _nhom(spam, dr):
    if spam:
        return "Spam"
    if dr == 0:
        return "0"
    if dr <= 30:
        return "1-30"
    if dr <= 50:
        return "30-50"
    if dr <= 70:
        return "50-70"
    return "70-100"


DU_LIEU_COT = [["Domain", "Domain"], ["Spam", "Spam"], ["DR", "DR"], ["Links to target", "Links to target"],
               ["First seen", "First seen"], ["Lost", "Lost"], ["Trạng thái", "Trạng thái"],
               ["Nhóm chất lượng", "Nhóm"], ["Năm first seen", "Năm"]]
MAU_NHOM_BRAND = ["#C00000", "#BFBFBF", "#12D6F7", "#7FA6F7", "#004AEF", "#1F2329"]


def build(wb, inputs, opts, ctx):
    """Sheet 1.9 Backlink theo CONTRACT muc 8.2: bang bl_tuy_chon (3 o chon), bl_theo_nam, bl_chat_luong,
    bl_chat_luong_pct, bl_site (moi site); moi cong thuc tro 'RD data' hoac o trong cac bang tren."""
    opts = opts or {}
    ctx = ctx or {}
    names = opts.get("sites") or []
    if not names and ctx.get("sites") and len(ctx["sites"]) == len(inputs):
        names = common.ten_site_theo_file(inputs, ctx["sites"])
    sites = []
    for i, p in enumerate(inputs):
        s, rows = doc_file(p, names[i] if i < len(names) else None)
        sites.append((s, rows))
    nam_dau = opts.get("nam_dau")
    ws = wb.create_sheet(SHEET)
    rd = wb.create_sheet(DATA_SHEET)

    # ---------- RD data (Trang thai/Nhom/Nam la gia tri tinh san) ----------
    rd.append([c["name"] for c in COLUMNS])
    n = 1
    years = []
    for site, rows in sites:
        for d in rows:
            n += 1
            rd.append([site, d[0], "Spam" if d[1] else "", d[2], d[3], d[4], d[5], d[6], d[7],
                       "Mất" if d[7] else "Còn", _nhom(d[1], d[2]), d[6].year if d[6] else ""])
            if d[6]:
                years.append(d[6].year)
    last = max(n, 2)
    for c in ("H", "I"):
        for cell in rd[c][1:]:
            cell.number_format = "dd/mm/yyyy"
    for i, w in enumerate([16, 30, 7, 6, 10, 10, 14, 12, 12, 10, 14, 10], 1):
        rd.column_dimensions[get_column_letter(i)].width = w
    for c in rd[1]:
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
    rd.freeze_panes = "A2"
    rd.auto_filter.ref = "A1:L%d" % last

    nam_cuoi = max(years) if years else dt.date.today().year
    nam_dau = nam_dau or max(min(years) if years else nam_cuoi, nam_cuoi - 6)
    nams = list(range(nam_dau, nam_cuoi + 1))
    ny = len(nams)
    R = lambda col: "'%s'!$%s$2:$%s$%d" % (DATA_SHEET, col, col, last)
    site_names = [s for s, _ in sites]
    N = len(site_names)
    L = get_column_letter

    common.title(ws, "ĐÁNH GIÁ BACKLINK THEO REFERRING DOMAINS",
                 "Nguồn: Ahrefs › Referring domains (export CSV). Nhóm chất lượng: Spam theo nhãn Ahrefs, còn lại theo DR.")
    chi_so = opts.get("chi_so") if opts.get("chi_so") in CHI_SO else CHI_SO[0]
    pham_vi = opts.get("pham_vi") if opts.get("pham_vi") in PHAM_VI else PHAM_VI[0]
    hien_spam = opts.get("hien_spam") if opts.get("hien_spam") in ("Có", "Không") else "Không"
    summaries = []
    # ---- bl_tuy_chon A4:B7 (o chon B5:B7)
    common.header_row(ws, 4, 1, ["Tùy chọn", "Giá trị"])
    opt_vals = [("Chỉ số", CHI_SO, chi_so), ("Phạm vi", PHAM_VI, pham_vi), ("Hiện Spam", ["Có", "Không"], hien_spam)]
    options = []
    for k, (lab, vals, df) in enumerate(opt_vals):
        r = 5 + k
        ws.cell(r, 1, lab).font = Font(bold=True)
        ws.cell(r, 2, df).fill = common.OPTION_FILL
        dv = DataValidation(type="list", formula1='"%s"' % ",".join(vals), allow_blank=False)
        ws.add_data_validation(dv)
        dv.add("B%d" % r)
        options.append({"cell": "%s!B%d" % (common.quote_sheet(SHEET), r), "label": lab, "values": vals, "default": df})
    ws["C5"] = "← chọn ở ô vàng; Không (Hiện Spam) = bỏ Spam khỏi Tổng và biểu đồ"
    ws["C5"].font = common.GREY_ITALIC
    summaries.append(common.bang(wb, ws, "bl_tuy_chon", "A4:B7",
                                 "3 ô chọn: Chỉ số (Số refdomain/Links to target), Phạm vi (chỉ domain còn link/cả domain đã mất), "
                                 "Hiện Spam (Có/Không). Mọi bảng backlink tính theo 3 ô này."))
    trang_thai = 'IF($B$6="%s","Còn","*")' % PHAM_VI[0]

    def cong_thuc(site_txt, nhom_cell, nam_cell):
        dk = '%s,"%s",%s,%s,%s,"<"&DATE(%s+1,1,1),%s,%s' % (
            R("A"), site_txt.replace('"', '""'), R("K"), nhom_cell, R("H"), nam_cell, R("J"), trang_thai)
        return 'IF($B$5="Links to target",SUMIFS(%s,%s),COUNTIFS(%s))' % (R("E"), dk, dk)

    # vi tri
    r_nam = 9
    r_cl = r_nam + ny + 2
    r_pct = r_cl + len(NHOM) + 3
    r_site0 = r_pct + len(NHOM) + 3
    site_rows = []
    row = r_site0
    for s in site_names:
        site_rows.append((row, row + 1, row + 1 + ny))  # (nhan, tieu de, dong cuoi)
        row = row + ny + 4
    heads = ["Năm"] + NHOM + ["Tổng", "% DR>30 không spam", "% Spam"]
    so = {"bl_tong": [], "bl_dr70": [], "bl_spam_pct": []}
    # ---- bl_site
    for si, s in enumerate(site_names):
        r0, h, e = site_rows[si]
        ws.cell(r0, 1, "%s – lũy kế đến hết năm (first seen trước 1/1 năm sau)" % s).font = Font(bold=True, color=XANH)
        common.header_row(ws, h, 1, heads)
        for k, y in enumerate(nams):
            r = h + 1 + k
            ws.cell(r, 1, y).border = BOX
            spam_f = cong_thuc(s, "B$%d" % h, "$A%d" % r)
            for j in range(len(NHOM)):
                col = L(2 + j)
                f = "=" + cong_thuc(s, "%s$%d" % (col, h), "$A%d" % r)
                if j == 0:
                    f = '=IF($B$7="Không",0,%s)' % spam_f
                c = ws.cell(r, 2 + j, f)
                c.number_format = "#,##0"
                c.border = BOX
            ws.cell(r, 8, "=SUM(B%d:G%d)" % (r, r)).number_format = "#,##0"
            ws.cell(r, 9, "=IFERROR(SUM(E%d:G%d)/H%d,0)" % (r, r, r)).number_format = "0.0%"
            ws.cell(r, 10, "=IFERROR((%s)/(SUM(C%d:G%d)+%s),0)" % (spam_f, r, r, spam_f)).number_format = "0.0%"
            for j in (8, 9, 10):
                ws.cell(r, j).border = BOX
            ws.cell(r, 8).font = Font(bold=True)
        summaries.append(common.bang(
            wb, ws, "bl_site", "A%d:J%d" % (h, e),
            "%s: refdomain lũy kế theo năm (dòng) × nhóm chất lượng (cột B:G), Tổng, %% DR>30 không spam, %% Spam. "
            "Theo 3 ô chọn bl_tuy_chon." % s, site=s, name="BL_" + s))
        so["bl_spam_pct"].append(common.so_cell(ws, "J%d" % e, s))
    # ---- bl_theo_nam
    ws.cell(r_nam - 1, 1, "Tổng refdomain lũy kế theo năm").font = Font(bold=True, color=XANH)
    common.header_row(ws, r_nam, 1, ["Năm"] + site_names)
    for k, y in enumerate(nams):
        r = r_nam + 1 + k
        ws.cell(r, 1, y).border = BOX
        for si in range(N):
            c = ws.cell(r, 2 + si, "=H%d" % (site_rows[si][1] + 1 + k))
            c.number_format = "#,##0"
            c.border = BOX
    e_nam = r_nam + ny
    summaries.append(common.bang(wb, ws, "bl_theo_nam", "A%d:%s%d" % (r_nam, L(1 + N), e_nam),
                                 "Tổng refdomain lũy kế đến hết từng năm (dòng) × website (cột), theo 3 ô chọn."))
    for si, s in enumerate(site_names):
        so["bl_tong"].append(common.so_cell(ws, "%s%d" % (L(2 + si), e_nam), s))
    # ---- bl_chat_luong (nam cuoi) + pct
    ws.cell(r_cl - 1, 1, "Chất lượng refdomain năm %d (lũy kế)" % nam_cuoi).font = Font(bold=True, color=XANH)
    common.header_row(ws, r_cl, 1, ["Nhóm (%d)" % nam_cuoi] + site_names)
    common.header_row(ws, r_pct, 1, ["Nhóm (%d) %%" % nam_cuoi] + site_names)
    for g, nh in enumerate(NHOM + ["Tổng"]):
        r = r_cl + 1 + g
        rp = r_pct + 1 + g
        ws.cell(r, 1, nh).border = BOX
        ws.cell(rp, 1, nh).border = BOX
        for si in range(N):
            e = site_rows[si][2]
            c = ws.cell(r, 2 + si, "=%s%d" % (L(2 + g), e))
            c.number_format = "#,##0"
            c.border = BOX
            col = L(2 + si)
            c = ws.cell(rp, 2 + si, "=IFERROR(%s%d/%s$%d,0)" % (col, r, col, r_cl + 1 + len(NHOM)))
            c.number_format = "0.0%"
            c.border = BOX
    e_cl = r_cl + 1 + len(NHOM)
    summaries.append(common.bang(wb, ws, "bl_chat_luong", "A%d:%s%d" % (r_cl, L(1 + N), e_cl),
                                 "Năm %d: số refdomain theo nhóm chất lượng (dòng) × website (cột), dòng cuối Tổng." % nam_cuoi,
                                 total_rows=1))
    summaries.append(common.bang(wb, ws, "bl_chat_luong_pct", "A%d:%s%d" % (r_pct, L(1 + N), r_pct + 1 + len(NHOM)),
                                 "Năm %d: tỉ trọng nhóm chất lượng trong từng website." % nam_cuoi, total_rows=1))
    for si, s in enumerate(site_names):
        so["bl_dr70"].append(common.so_cell(ws, "%s%d" % (L(2 + si), r_cl + len(NHOM)), s))
    note = common.note_area(ws, "L5", 10, 8, "NHẬN XÉT")
    ws.column_dimensions["A"].width = 18
    for j in range(2, 11):
        ws.column_dimensions[L(j)].width = 11
    charts = [
        {"key": "bl_theo_nam", "bang": "bl_theo_nam", "loai": "col", "tieu_de": "Số lượng backlink theo thời gian"},
        {"key": "bl_chat_luong", "bang": "bl_chat_luong", "loai": "col", "chuoi": "dong", "mau": MAU_NHOM_BRAND,
         "tieu_de": "Chất lượng backlink"},
        {"key": "bl_site", "bang": "bl_site", "loai": "col_stacked", "cot": NHOM, "mau": MAU_NHOM_BRAND,
         "tieu_de": "{{ten}} – chất lượng refdomain lũy kế theo năm"},
    ]
    du_lieu = [{"key": "bl_data", "sheet": DATA_SHEET, "site_col": "Website", "cot": DU_LIEU_COT}]
    return common.entry(
        SHEET, KEY, DESC, list(inputs), data_sheet=DATA_SHEET, rows=last - 1, columns=COLUMNS,
        summaries=summaries, options=options, caveats=_caveats(sites, nams), notes=[note],
        extra={"sites": site_names, "years": nams, "so": so, "charts": charts, "du_lieu": du_lieu})


def dung_file(inputs, out_path, nam_dau=None):
    # type: (List[str], str, Optional[int]) -> dict
    wb = Workbook()
    wb.remove(wb.active)
    e = build(wb, inputs, {"nam_dau": nam_dau}, {"out_path": out_path})
    wb.save(out_path)
    return {"sites": e["sites"], "rows": e["rows"], "years": e["years"], "out": out_path}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("inputs", nargs="+", help="file CSV Referring domains (Ahrefs), mỗi website 1 file")
    ap.add_argument("--out", required=True)
    ap.add_argument("--nam-dau", type=int, default=None, help="năm đầu bảng (mặc định: 6 năm gần nhất)")
    a = ap.parse_args()
    print(dung_file(a.inputs, a.out, a.nam_dau))


if __name__ == "__main__":
    main()
