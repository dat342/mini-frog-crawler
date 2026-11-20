"""Phan dung chung cho moi module xu ly data BD SEO. Xem CONTRACT.md.
Python 3.9: khong dung cu phap X | None.
"""
import re
from typing import Dict, List, Optional

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName

XANH = "0B3D91"
HEADER_FILL = PatternFill("solid", fgColor="DCE6F2")
OPTION_FILL = PatternFill("solid", fgColor="FFF2CC")
NOTE_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
GREY_ITALIC = Font(italic=True, color="7F7F7F")


def quote_sheet(name):
    # type: (str) -> str
    return "'%s'" % name.replace("'", "''")


def ten_hop_le(s):
    # type: (str) -> str
    """Ten defined name hop le: chu, so, gach duoi; bat dau bang chu."""
    s = re.sub(r"[^0-9A-Za-z_]", "_", s)
    return s if re.match(r"[A-Za-z_]", s) else "_" + s


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=14, color=XANH)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = GREY_ITALIC


def header_row(ws, row, col, labels):
    for j, t in enumerate(labels):
        c = ws.cell(row, col + j, t)
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX


def summary(wb, ws, name, ref, desc):
    # type: (object, object, str, str, str) -> Dict
    """Dang ky 1 vung tong hop: tao defined name TH_<name> va tra ve muc cho manifest.
    ref dang 'A7:J15' (trong ws)."""
    dn = "TH_" + ten_hop_le(name)
    full ="%s!%s" % (quote_sheet(ws.title), _abs(ref))
    wb.defined_names[dn] = DefinedName(dn, attr_text=full)
    return {"name": dn, "range": full, "desc": desc}


def _abs(ref):
    out = []
    for part in ref.split(":"):
        m = re.match(r"([A-Z]+)(\d+)$", part)
        out.append("$%s$%s" % (m.group(1), m.group(2)) if m else part)
    return ":".join(out)


def note_area(ws, top_left, cols, rows, label):
    # type: (object, str, int, int, str) -> Dict
    """Vung de Claude ghi nhan xet. O nhan (label) nam ngay tren top_left; vung duoc merge.
    Tra ve muc cho manifest: {"cell": "'Sheet'!A30", "label": ...}"""
    m = re.match(r"([A-Z]+)(\d+)$", top_left)
    col_letter, r = m.group(1), int(m.group(2))
    from openpyxl.utils import column_index_from_string, get_column_letter
    c0 = column_index_from_string(col_letter)
    lab = ws.cell(r - 1, c0, label)
    lab.font = Font(bold=True, color=XANH)
    end = "%s%d" % (get_column_letter(c0 + cols - 1), r + rows - 1)
    ws.merge_cells("%s:%s" % (top_left, end))
    cell = ws[top_left]
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    cell.fill = NOTE_FILL
    return {"cell": "%s!%s" % (quote_sheet(ws.title), top_left), "label": label}


def option(ws, cell, label, values, default, hint=None):
    # type: (object, str, str, List[str], str, Optional[str]) -> Dict
    """O tuy chon (dropdown). Nhan dat o cot ben trai cell."""
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.utils import column_index_from_string, get_column_letter
    m = re.match(r"([A-Z]+)(\d+)$", cell)
    c0, r = column_index_from_string(m.group(1)), int(m.group(2))
    ws.cell(r, c0 - 1, label).font = Font(bold=True)
    ws[cell] = default
    ws[cell].fill = OPTION_FILL
    ws[cell].font = Font(bold=True)
    dv = DataValidation(type="list", formula1='"%s"' % ",".join(values), allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(cell)
    if hint:
        ws.cell(r, c0 + 1, hint).font = GREY_ITALIC
    return {"cell": "%s!%s" % (quote_sheet(ws.title), cell), "label": label, "values": values, "default": default}


def entry(sheet, module, purpose, inputs, data_sheet=None, rows=0, columns=None, summaries=None,
          options=None, caveats=None, notes=None, extra=None):
    """Muc manifest cho 1 sheet ket qua. Xem CONTRACT.md muc 3."""
    e = {
        "sheet": sheet,
        "module": module,
        "purpose": purpose,
        "inputs": inputs,
        "data_sheet": data_sheet,
        "rows": rows,
        "columns": columns or [],
        "summaries": summaries or [],
        "options": options or [],
        "caveats": caveats or [],
        "notes": notes or [],
    }
    if extra:
        e.update(extra)
    return e


# ---------------------------------------------------------------- muc 8 (mau gan the)
def bang(wb, ws, key, ref, desc, site=None, total_rows=0, name=None):
    """Dang ky 1 bang tong hop co khoa (CONTRACT muc 8.2). ref 'A5:D12' trong ws, dong 1 = tieu de cot,
    cot 1 = nhan. site: bang rieng cua 1 site (vd bl_site). total_rows: so dong tong o cuoi (bieu do bo qua)."""
    nm = name or (key + ("_" + site if site else ""))
    s = summary(wb, ws, nm, ref, desc)
    s["key"] = key
    if site:
        s["site"] = site
    if total_rows:
        s["total_rows"] = total_rows
    return s


def so_cell(ws, cell, site=None):
    """Muc extra.so: {"site": site, "cell": "'Sheet'!B5"}."""
    return {"site": site, "cell": "%s!%s" % (quote_sheet(ws.title), cell)}


def ten_site_theo_file(inputs, sites):
    # type: (List[str], List[str]) -> List[str]
    """Gan ten site cho tung file dau vao: uu tien domain xuat hien trong TEN FILE (vd 'seongon.com-refdomains-...csv',
    'minifrog-seongon.com.json'); chi khi khong khop duoc tung file 1-1 moi gan theo thu tu (truong hop cu).
    Tranh loi file sap theo ABC lam lech nhan site."""
    import os
    def norm(s):
        s = str(s or "").lower().strip()
        s = re.sub(r"^[a-z]+://", "", s)
        return re.sub(r"^www\.", "", s).rstrip("/")
    ns = [norm(s) for s in sites]
    out = []
    for p in inputs:
        base = os.path.basename(str(p)).lower()
        hit = [i for i, s in enumerate(ns) if s and s in base]
        if len(hit) > 1:  # vd 'seo.com' nam trong 'abcseo.com': chon domain dai nhat
            hit = [max(hit, key=lambda i: len(ns[i]))]
        out.append(hit[0] if len(hit) == 1 else None)
    if None not in out and len(set(out)) == len(out):
        return [sites[i] for i in out]
    return list(sites)
