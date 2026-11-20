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
    {"name": "Trạng thái", "desc": "Còn/Mất, tính từ cột Lost"},
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


def build(wb, inputs, opts, ctx):
    opts = opts or {}
    names = opts.get("sites") or []
    sites = []
    for i, p in enumerate(inputs):
        s, rows = doc_file(p, names[i] if i < len(names) else None)
        sites.append((s, rows))
    nam_dau = opts.get("nam_dau")
    ws = wb.create_sheet(SHEET)
    rd = wb.create_sheet(DATA_SHEET)

    # ---------- RD data ----------
    rd.append([c["name"] for c in COLUMNS])
    n = 1
    years = []
    for site, rows in sites:
        for d in rows:
            n += 1
            rd.append([site, d[0], "Spam" if d[1] else "", d[2], d[3], d[4], d[5], d[6], d[7],
                       '=IF(I%d="","Còn","Mất")' % n,
                       '=IF(C{0}="Spam","Spam",IF(D{0}=0,"0",IF(D{0}<=30,"1-30",IF(D{0}<=50,"30-50",'
                       'IF(D{0}<=70,"50-70","70-100")))))'.format(n),
                       '=IF(H%d="","",YEAR(H%d))' % (n, n)])
            if d[6]:
                years.append(d[6].year)
    last = n
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
    R = lambda col: "'%s'!$%s$2:$%s$%d" % (DATA_SHEET, col, col, last)

    # ---------- 1.9 Backlink ----------
    common.title(ws, "ĐÁNH GIÁ BACKLINK THEO REFERRING DOMAINS",
                 "Nguồn: Ahrefs › Referring domains (export CSV). Nhóm chất lượng: Spam theo nhãn Ahrefs, còn lại theo DR.")
    chi_so = opts.get("chi_so") if opts.get("chi_so") in CHI_SO else CHI_SO[0]
    pham_vi = opts.get("pham_vi") if opts.get("pham_vi") in PHAM_VI else PHAM_VI[0]
    hien_spam = opts.get("hien_spam") if opts.get("hien_spam") in ("Có", "Không") else "Không"
    options = [
        common.option(ws, "B3", "Chỉ số", CHI_SO, chi_so, "← chọn ở ô vàng"),
        common.option(ws, "B4", "Phạm vi", PHAM_VI, pham_vi),
        common.option(ws, "B5", "Hiện Spam", ["Có", "Không"], hien_spam,
                      "Không = bỏ Spam khỏi cột Tổng và biểu đồ (cột % Spam vẫn tính)"),
    ]
    ws["D3"] = '=IF($B$4="%s","Còn","*")' % PHAM_VI[0]
    ws["D3"].font = Font(color="FFFFFF")  # o phu: dieu kien trang thai

    def cong_thuc(site_cell, nhom_cell, nam_cell):
        dk = "%s,%s,%s,%s,%s,\"<\"&DATE(%s+1,1,1),%s,$D$3" % (
            R("A"), site_cell, R("K"), nhom_cell, R("H"), nam_cell, R("J"))
        return '=IF($B$3="Links to target",SUMIFS(%s,%s),COUNTIFS(%s))' % (R("E"), dk, dk)

    row = 7
    blocks = []
    summaries = []
    for si, (site, _) in enumerate(sites):
        ws.cell(row, 1, site).font = Font(bold=True, size=12, color=XANH)
        ws.cell(row, 3, "Lũy kế đến hết năm (domain có first seen trước 1/1 năm sau)").font = Font(italic=True, color="7F7F7F")
        h = row + 1
        heads = ["Năm"] + NHOM + ["Tổng", "% DR>30 không spam", "% Spam"]
        for j, t in enumerate(heads, 1):
            c = ws.cell(h, j, t)
            c.font = Font(bold=True)
            c.fill = HEADER_FILL
            c.alignment = Alignment(horizontal="center", wrap_text=True)
            c.border = BOX
        for k, y in enumerate(nams):
            r = h + 1 + k
            ws.cell(r, 1, y).border = BOX
            spam_f = cong_thuc("$A$%d" % row, "B$%d" % h, "$A%d" % r)[1:]
            for j in range(len(NHOM)):
                col = get_column_letter(2 + j)
                f = cong_thuc("$A$%d" % row, "%s$%d" % (col, h), "$A%d" % r)
                if j == 0:
                    f = '=IF($B$5="Không",0,%s)' % spam_f
                c = ws.cell(r, 2 + j, f)
                c.number_format = "#,##0"
                c.border = BOX
            ws.cell(r, 8, "=SUM(B%d:G%d)" % (r, r)).number_format = "#,##0"
            ws.cell(r, 9, "=IFERROR(SUM(E%d:G%d)/H%d,0)" % (r, r, r)).number_format = "0.0%"
            ws.cell(r, 10, "=IFERROR((%s)/(SUM(C%d:G%d)+%s),0)" % (spam_f, r, r, spam_f)).number_format = "0.0%"
            for j in (8, 9, 10):
                ws.cell(r, j).border = BOX
            ws.cell(r, 8).font = Font(bold=True)
        end = h + len(nams)
        summaries.append(common.summary(
            wb, ws, "BL_" + site, "A%d:J%d" % (h, end),
            "%s: refdomain lũy kế theo năm (cột) x nhóm chất lượng; cột Tổng, %% DR>30 không spam, %% Spam. "
            "Số liệu theo 2 ô chọn Chỉ số/Phạm vi và ô Hiện Spam." % site))
        ch = BarChart()
        ch.type = "col"
        ch.grouping = "stacked"
        ch.overlap = 100
        ch.gapWidth = 60
        ch.title = "%s – chất lượng refdomain lũy kế theo năm" % site
        ch.y_axis.title = None
        ch.y_axis.numFmt = "#,##0"
        ch.y_axis.majorGridlines = None
        ch.add_data(Reference(ws, min_col=2, max_col=1 + len(NHOM), min_row=h, max_row=end), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=h + 1, max_row=end))
        for s, color in zip(ch.series, MAU_NHOM):
            s.graphicalProperties.solidFill = color
            s.graphicalProperties.line.solidFill = color
        ch.legend.position = "b"
        ch.height, ch.width = 7.5, 15
        ws.add_chart(ch, "L%d" % row)
        blocks.append((site, row, h, end))
        row = max(end + 3, row + 17)

    # So sanh cac site o nam cuoi
    ws.cell(row, 1, "So sánh tỉ trọng hiện tại (năm %d)" % nam_cuoi).font = Font(bold=True, size=12, color=XANH)
    h = row + 1
    for j, t in enumerate(["Website"] + NHOM + ["Tổng"], 1):
        c = ws.cell(h, j, t)
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
        c.border = BOX
    for k, (site, srow, sh, send) in enumerate(blocks):
        r = h + 1 + k
        ws.cell(r, 1, "=A%d" % srow).border = BOX
        for j in range(len(NHOM) + 1):
            col = get_column_letter(2 + j)
            c = ws.cell(r, 2 + j, "=IFERROR(%s%d/$H%d,0)" % (col, send, send) if j < len(NHOM) else "=H%d" % send)
            c.number_format = "0.0%" if j < len(NHOM) else "#,##0"
            c.border = BOX
    cmp_end = h + len(blocks)
    summaries.append(common.summary(
        wb, ws, "BL_so_sanh", "A%d:H%d" % (h, cmp_end),
        "So sánh các site ở năm %d: tỉ trọng từng nhóm chất lượng và Tổng." % nam_cuoi))
    note = common.note_area(ws, "A%d" % (cmp_end + 3), 10, 8, "NHẬN XÉT")
    ws.column_dimensions["A"].width = 18
    for j in range(2, 11):
        ws.column_dimensions[get_column_letter(j)].width = 11
    ws.column_dimensions["I"].width = 13
    ws.freeze_panes = "A6"
    return common.entry(
        SHEET, KEY, DESC, list(inputs), data_sheet=DATA_SHEET, rows=last - 1, columns=COLUMNS,
        summaries=summaries, options=options, caveats=_caveats(sites, nams), notes=[note],
        extra={"sites": [s for s, _ in sites], "years": nams})


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
