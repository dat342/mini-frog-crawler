"""[CU — thay boi xu_ly_data/the_mau.py + mau gan the xu_ly_data/mau/bd-seo.xlsx tu 20/11/2026.
bd_seo khong goi file nay nua: recipe "template" duoc chuyen sang mau "bd-seo". Giu lai de doc/chuyen notes file cu;
bo cuc module da doi (RD data gia tri, onpage co sheet 'Onpage data') nen do-vao-mau co the khong con chay dung.]

Do ket qua BD SEO vao MAU Excel cua nguoi dung (file xuat tu Google Sheets). Xem CONTRACT.md muc 7.

    .venv/bin/python -m xu_ly_data.template_bd do-vao-mau --built tam.xlsx --manifest tam.manifest.json \
        --template "BD SEO _ c06_2026).xlsx" --out "<Khach> - BD SEO.xlsx"
    .venv/bin/python -m xu_ly_data.template_bd chuyen-notes --tu "<file cu>.xlsx" --sang "<file theo mau>.xlsx"

Nguyen tac:
- File mau CHI DOC. openpyxl chi dung de DOC (mau + file da build). Moi thay doi tren mau di qua
  GridCraft CLI (`gridcraft-cli run --in <mau> --script steps.jsonl --out <file>`) de giu bieu do, dinh dang, pivot.
- Du lieu lay tu file da build boi cac module (sheet 'RD data', 'Traffic data', 'TP data', bang tieu chi 1.5),
  KHONG goi lai API/crawl.
- Bo cuc: giu 6 sheet (Muc luc, 1. Dinh huong, 1.5, 1.6, 1.7, 1.9). Bang thong ke ben trai, du lieu tho tung site
  xep canh nhau ben phai trong cung sheet (moi site 1 khoi: hang 1 ten site, hang 2 tieu de cot).
- Nhan xet nam o "1. Dinh huong" (4 o khung "Du lieu nghien cuu / Danh gia / Dinh huong").
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
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter

GRIDCRAFT = os.environ.get(
    "GRIDCRAFT_CLI",
    os.path.expanduser("~/.local/gridcraft/gridcraft-cli-0.19.1-macos-universal/gridcraft-cli"))

SH_MUCLUC = "Mục lục"
SH_DH = "1. Định hướng"
SH_ONPAGE = "1.5 Dữ liệu Onpage"
SH_TRAFFIC = "1.6 Dữ liệu traffic"
SH_CONTENT = "1.7 Số lượng content"
SH_BACKLINK = "1.9 Backlink"
MODULE_SHEET = {"onpage": SH_ONPAGE, "traffic_ahrefs": SH_TRAFFIC, "content_topic": SH_CONTENT,
                "backlink_rd": SH_BACKLINK}
# hang muc trong Muc luc (nhom "Dinh huong") -> sheet du lieu
MUC_LUC = {"check onpage sơ bộ": SH_ONPAGE, "so sánh traffic đối thủ": SH_TRAFFIC,
           "nghiên cứu số lượng content": SH_CONTENT, "nghiên cứu số lượng offpage": SH_BACKLINK,
           "nghiên cứu chất lượng offpage": SH_BACKLINK}
KHUNG_NHAN_XET = "Dữ liệu nghiên cứu:\n\nĐánh giá: \n\nĐịnh hướng:"

# mau sac theo mau
XANH_DAM = "#004AEF"
XANH_NHAT = "#12D6F7"
VIEN = "#1F2329"
VIEN_DATA = "#666666"
FONT = "Google Sans"
# o danh gia onpage: nen theo cot site (dung nhu mau F/G/H), tieu de cot G/H cua mau
EVAL_FILL = ["#12D6F7", "#FFCE04", "#D3D3D3"]
EVAL_HDR_FILL = ["#004AEF", "#7F6000", "#000000"]

NHOM_BL = ["Spam", "0", "1-30", "30-50", "50-70", "70-100"]
CHI_SO = ["Số refdomain", "Links to target"]
PHAM_VI = ["Chỉ domain còn link", "Cả domain đã mất"]
LOAI_TRANG = ["Trang chủ", "Danh mục", "Sản phẩm/Dịch vụ", "Blog", "Landing", "Chính sách/hỗ trợ", "Khác"]
DAT, CHUA_DAT = "Đạt", "Chưa đạt"


# ---------------------------------------------------------------- tien ich
def L(c):
    return get_column_letter(c)


def q(sheet):
    return "'%s'" % sheet.replace("'", "''")


def absref(sheet, c1, r1, c2=None, r2=None):
    if c2 is None:
        return "%s!$%s$%d" % (q(sheet), L(c1), r1)
    return "%s!$%s$%d:$%s$%d" % (q(sheet), L(c1), r1, L(c2), r2)


def serial(d):
    """datetime/date -> so serial Excel (ngay, bo gio)."""
    if d is None:
        return None
    if isinstance(d, dt.datetime):
        d = d.date()
    return (d - dt.date(1899, 12, 30)).days


def split_ref(ref):
    m = re.match(r"^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)$", ref.strip())
    if not m:
        raise ValueError("Tham chiếu ô không hợp lệ: %s" % ref)
    return m.group(1).replace("''", "'"), "%s%s" % (m.group(2), m.group(3))


def norm(ref):
    try:
        s, c = split_ref(ref)
        return "%s!%s" % (s, c)
    except ValueError:
        return ref.strip()


def ten_hop_le(s):
    s = re.sub(r"[^0-9A-Za-z_]", "_", s)
    return s if re.match(r"[A-Za-z_]", s) else "_" + s


class Script(object):
    """Gom lenh GridCraft (jsonl)."""

    def __init__(self):
        self.lines = []

    def cmd(self, command, **params):
        self.lines.append({"command": command, "params": params})

    def sheet(self, name):
        self.cmd("sheet.activate", sheet=name)

    def values(self, top_left, rows):
        if rows:
            self.cmd("range.setValues", range=top_left, values=rows, parse=False)

    def set(self, cell, v):
        self.values(cell, [[v]])

    def style(self, rng, kind, border=True, wrap=None, num=None, align="center"):
        """kind: h1 (nen xanh dam, chu trang dam) | h2 (nen xanh nhat, chu trang dam) | lab (nen xanh nhat, chu trang)
        | data (chu den nen trang) | hint (chu xam nghieng) | link (chu xanh)."""
        self.cmd("home.fontName", range=rng, name=FONT)
        self.cmd("home.fontSize", range=rng, size=9)
        if kind in ("h1", "h2", "lab"):
            self.cmd("home.fillColor", range=rng, color=XANH_DAM if kind == "h1" else XANH_NHAT)
            self.cmd("home.fontColor", range=rng, color="#FFFFFF")
            self.cmd("home.bold", range=rng, on=kind != "lab")
        elif kind == "hint":
            self.cmd("home.fillColor", range=rng, color="none")
            self.cmd("home.fontColor", range=rng, color="#7F7F7F")
            self.cmd("home.italic", range=rng, on=True)
            border = False
        else:
            self.cmd("home.fillColor", range=rng, color="#FFFFFF")
            self.cmd("home.fontColor", range=rng, color=XANH_DAM if kind == "link" else "#000000")
            self.cmd("home.bold", range=rng, on=False)
        if align == "center":
            self.cmd("home.alignCenter", range=rng)
        elif align == "left":
            self.cmd("home.alignLeft", range=rng)
        self.cmd("home.alignMiddle" if kind != "datatop" else "home.alignTop", range=rng)
        if wrap is not None:
            self.cmd("home.wrapText", range=rng, on=wrap)
        if border:
            self.cmd("home.borders", range=rng, preset="all", style="thin",
                     color=VIEN_DATA if kind == "data" and border == "data" else VIEN)
        if num:
            self.cmd("home.numberFormat", range=rng, code=num)

    def width(self, cols, chars):
        self.cmd("home.columnWidth", cols=cols, chars=chars)

    def name(self, nm, ref):
        self.cmd("formulas.defineName", name=nm, refersTo="=" + ref)

    def dump(self, path):
        with open(path, "w", encoding="utf-8") as f:
            for x in self.lines:
                f.write(json.dumps(x, ensure_ascii=False) + "\n")


def gc_run(src, script_path, out):
    p = subprocess.run([GRIDCRAFT, "run", "--in", src, "--script", script_path, "--out", out, "--quiet"],
                       stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    err = "\n".join(l for l in p.stderr.splitlines() if not l.startswith("Warning"))
    if p.returncode != 0:
        raise RuntimeError("GridCraft lỗi (rc=%d): %s" % (p.returncode, err[-1500:]))
    return err


def gc_info(path):
    p = subprocess.run([GRIDCRAFT, "info", path, "--json"], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                       universal_newlines=True)
    if p.returncode != 0:
        raise RuntimeError("GridCraft info lỗi: %s" % p.stderr[-500:])
    return json.loads(p.stdout)


# ---------------------------------------------------------------- doc mau
def _chart_refs(ch):
    refs = []
    for s in ch.series:
        for part in (s.tx and s.tx.strRef, s.val and s.val.numRef, s.cat and (s.cat.numRef or s.cat.strRef)):
            if part is not None and part.f:
                refs.append(part.f)
    return refs


def _ref_sheet(f):
    m = re.match(r"^'?(.*?)'?!", f)
    return m.group(1).replace("''", "'") if m else None


def doc_mau(template_path):
    """Doc cau truc mau: sheet, chart (id GridCraft + tham chieu), o nhan xet Dinh huong, Muc luc..."""
    info = gc_info(template_path)
    wb = load_workbook(template_path)  # chi doc
    out = {"sheets": [s["name"] for s in info["sheets"]], "charts": {}, "max_row": {}, "merges": {}}
    for s in info["sheets"]:
        name = s["name"]
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        by_at = {}
        for ch in ws._charts:
            a = ch.anchor._from
            by_at.setdefault("%s%d" % (L(a.col + 1), a.row + 1), []).append(ch)
        lst = []
        for c in s["charts"]:
            cands = by_at.get(c["at"]) or []
            ch = cands.pop(0) if cands else None
            refs = _chart_refs(ch) if ch is not None else []
            lst.append({"id": c["id"], "at": c["at"], "kind": c["kind"], "title": c.get("title"), "refs": refs,
                        "ref_sheets": sorted({_ref_sheet(f) for f in refs if _ref_sheet(f)})})
        out["charts"][name] = lst
        out["max_row"][name] = ws.max_row
        out["merges"][name] = [str(m) for m in ws.merged_cells.ranges]
    # Dinh huong: muc II.1, cac khung nhan xet, hang ten site cua radar
    ws = wb[SH_DH]
    a_text = {}
    for r in range(1, ws.max_row + 1):
        for c in range(1, 14):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.strip():
                a_text[(r, c)] = v
    r_ii1 = r_ii2 = None
    for (r, c), v in sorted(a_text.items()):
        t = v.strip()
        if c == 1 and t.startswith("II.1") and r_ii1 is None:
            r_ii1 = r
        if c == 1 and t.startswith("II.2") and r_ii2 is None:
            r_ii2 = r
    notes = []
    sec = None
    for (r, c), v in sorted(a_text.items()):
        t = v.strip()
        m = re.match(r"^II\.2\.\d+\s+(.*)$", t)
        if m:
            sec = (r, m.group(1).strip())
        if t.startswith("Dữ liệu nghiên cứu") and r_ii2 and r > r_ii2:
            notes.append({"row": r, "col": c, "label": sec[1] if sec else "", "sec_row": sec[0] if sec else None})
    out["dh"] = {"r_ii1": r_ii1, "r_ii2": r_ii2, "notes": notes,
                 "title": ws["A5"].value, "radar_row": None}
    for (r, c), v in sorted(a_text.items()):
        if v.strip().startswith("II.2.1"):
            out["dh"]["radar_row"] = r + 1
            break
    # Muc luc
    ws = wb[SH_MUCLUC]
    rows = []
    for r in range(2, ws.max_row + 1):
        b, cc = ws.cell(r, 2).value, ws.cell(r, 3).value
        if b is None and cc is None:
            continue
        rows.append((r, (b or "").strip(), (cc or "").strip()))
    out["muc_luc"] = rows
    out["onpage_hdr"] = [ws_val for ws_val in (wb[SH_ONPAGE].cell(8, c).value for c in range(1, 9))]
    wb.close()
    return out


# ---------------------------------------------------------------- doc du lieu da build
def doc_built(built_xlsx, manifest):
    """Doc du lieu tu file da build (openpyxl, chi doc). Tra ve dict theo module."""
    wb = load_workbook(built_xlsx, read_only=True)
    ent = {e["module"]: e for e in manifest.get("sheets", [])}
    data = {"sites": list(manifest.get("sites") or []), "entries": ent}
    if "backlink_rd" in ent and "RD data" in wb.sheetnames:
        rows = list(wb["RD data"].iter_rows(min_row=2, values_only=True))
        by = {}
        for r in rows:
            if r[0] is None:
                continue
            by.setdefault(r[0], []).append(r)
        data["rd"] = by
        data["rd_years"] = ent["backlink_rd"].get("years") or []
        data["rd_sites"] = ent["backlink_rd"].get("sites") or list(by)
    if "traffic_ahrefs" in ent and "Traffic data" in wb.sheetnames:
        rows = list(wb["Traffic data"].iter_rows(min_row=2, values_only=True))
        by = {}
        for s, t, v in rows:
            if s is None:
                continue
            by.setdefault(s, {})[t] = v or 0
        data["traffic"] = by
        data["traffic_sites"] = list(ent["traffic_ahrefs"].get("inputs") or by)
    if "content_topic" in ent and "TP data" in wb.sheetnames:
        rows = list(wb["TP data"].iter_rows(min_row=2, values_only=True))
        by = {}
        for r in rows:
            if r[0] is None:
                continue
            by.setdefault(r[0], []).append(r)
        data["tp"] = by
        data["tp_sites"] = ent["content_topic"].get("sites") or list(by)
    if "onpage" in ent and SH_ONPAGE in wb.sheetnames:
        e = ent["onpage"]
        sites = e.get("sites") or []
        n = len(sites)
        r1, r2 = e.get("criteria_rows") or [15, 14]
        ws = wb[SH_ONPAGE]
        crit = []
        for r in ws.iter_rows(min_row=r1, max_row=r2, values_only=True):
            r = list(r) + [None] * (7 + 2 * n)
            ev = []
            for j in range(n):
                v = r[6 + j]
                if isinstance(v, str) and v.startswith('="') and v.endswith('"'):
                    v = v[2:-1]
                ev.append(v)
            crit.append({"nhom": r[0], "nhom_con": r[1], "van_de": r[2], "vi_sao": r[3], "yeu_cau": r[4],
                         "nguong": r[5], "eval": ev, "evid": [r[6 + n + j] for j in range(n)],
                         "who": r[6 + 2 * n]})
        data["onpage"] = crit
        data["onpage_sites"] = sites
        data["onpage_groups"] = []
        for c in crit:
            if c["nhom"] not in data["onpage_groups"]:
                data["onpage_groups"].append(c["nhom"])
    wb.close()
    return data


# ---------------------------------------------------------------- dung tung sheet
def _chart_role(sheet, ch):
    """Xac dinh vai tro chart cua mau theo tham chieu. Tra ve ma vai tro hoac None."""
    refs = ch["refs"]
    rs = set(ch["ref_sheets"])
    if not refs:
        return None
    if rs == {SH_ONPAGE} and ch["kind"].startswith("Radar"):
        m = re.search(r"!\$?([A-Z]+)\$?\d+:", [f for f in refs if ":" in f][0])
        return "radar:%d" % (column_index_from_string(m.group(1)) - 2)
    if rs == {SH_TRAFFIC}:
        cols = {re.search(r"!\$?([A-Z]+)", f).group(1) for f in refs}
        return "traffic_organic" if cols <= {"J", "K", "L", "M"} else "traffic_kenh"
    if rs == {SH_CONTENT}:
        return "content_topic"
    if rs == {SH_BACKLINK}:
        return "bl_stack" if "Stacked" in ch["kind"] else "bl_time"
    return None


def dung_onpage(S, data, notes, mau):
    sites = data["onpage_sites"]
    n = len(sites)
    groups = data["onpage_groups"]
    crit = data["onpage"]
    sh = SH_ONPAGE
    S.sheet(sh)
    S.cmd("home.unmergeCells", range="A1:Z%d" % max(mau["max_row"].get(sh, 120), 120))
    S.cmd("edit.clearAll", range="A1:Z%d" % max(mau["max_row"].get(sh, 120), 120))
    hdr_row = len(groups) + 4          # 5 nhom -> bang diem A1:D6, Tong dong 7, trong dong 8, tieu de dong 9
    r1 = hdr_row + 1
    r2 = hdr_row + len(crit)
    c_ev = 6                           # F
    c_bc = c_ev + n
    c_ng = c_bc + n
    c_who = c_ng + 1
    # --- bang diem nhom (A1:D6 nhu mau) + dong Tong
    S.values("A1", [["Nhóm"] + sites])
    S.style("A1:%s1" % L(1 + n), "h1", wrap=True)
    S.values("A2", [[g] for g in groups] + [["Tổng (mọi tiêu chí đã chấm)"]])
    S.style("A2:A%d" % (2 + len(groups)), "lab", wrap=True, align="left")
    S.cmd("home.bold", range="A2:A%d" % (2 + len(groups)), on=True)
    rng = lambda col: "$%s$%d:$%s$%d" % (col, r1, col, r2)
    rows = []
    for i, g in enumerate(groups + [None]):
        r = 2 + i
        row = []
        for j in range(n):
            col = L(c_ev + j)
            if g is None:
                d_ = 'COUNTIF(%s,"%s")' % (rng(col), DAT)
                x_ = 'COUNTIF(%s,"%s")' % (rng(col), CHUA_DAT)
            else:
                d_ = 'COUNTIFS(%s,$A%d,%s,"%s")' % (rng("A"), r, rng(col), DAT)
                x_ = 'COUNTIFS(%s,$A%d,%s,"%s")' % (rng("A"), r, rng(col), CHUA_DAT)
            row.append('=IFERROR(%s/(%s+%s),"")' % (d_, d_, x_))
        rows.append(row)
    S.values("B2", rows)
    last_sc = 2 + len(groups)
    S.style("B2:%s%d" % (L(1 + n), last_sc), "data", num="0.0%")
    S.cmd("home.bold", range="B%d:%s%d" % (last_sc, L(1 + n), last_sc), on=True)
    # --- bang tieu chi
    heads = ["Nhóm", "Các vấn đề cần kiểm tra", "", "Vì sao cần làm", "Yêu cầu đầu ra"]
    heads += ["Đánh giá %s" % s for s in sites] + ["Bằng chứng %s" % s for s in sites]
    heads += ["Ngưỡng tự chấm", "Người chấm"]
    S.values("A%d" % hdr_row, [heads])
    S.style("A%d:%s%d" % (hdr_row, L(c_who), hdr_row), "h1", wrap=True)
    for j in range(n):
        S.cmd("home.fillColor", range="%s%d" % (L(c_ev + j), hdr_row), color=EVAL_HDR_FILL[j % 3])
    S.cmd("home.mergeCells", range="B%d:C%d" % (hdr_row, hdr_row))
    claude_cells, claude_items, body = [], [], []
    applied = 0
    for i, t in enumerate(crit):
        r = r1 + i
        ev, bc = list(t["eval"]), list(t["evid"])
        empty_here = []
        for j in range(n):
            ref = "%s!%s%d" % (q(sh), L(c_ev + j), r)
            if t["who"] in ("Claude", "Tool + Claude") and not ev[j]:
                empty_here.append(ref)
            nt = notes.get(norm(ref))
            if nt is not None:
                ev[j] = nt.get("text", "")
                vals = nt.get("values") or []
                bc[j] = "; ".join(str(v) for v in vals) if vals else bc[j]
                applied += 1
        if empty_here:
            claude_cells += empty_here
            claude_items.append({"row": r, "van_de": t["van_de"], "yeu_cau": (t["yeu_cau"] or "")[:200],
                                 "cells": empty_here})
        body.append([t["nhom"], t["nhom_con"], t["van_de"], t["vi_sao"] or "", t["yeu_cau"] or ""] + ev + bc +
                    [t["nguong"] or "", t["who"] or ""])
    S.values("A%d" % r1, body)
    S.style("A%d:%s%d" % (r1, L(c_who), r2), "data", wrap=True, align="left")
    S.cmd("home.alignTop", range="A%d:%s%d" % (r1, L(c_who), r2))
    for j in range(n):
        rg = "%s%d:%s%d" % (L(c_ev + j), r1, L(c_ev + j), r2)
        S.cmd("home.fillColor", range=rg, color=EVAL_FILL[j % 3])
        S.cmd("home.fontColor", range=rg, color="#FFFFFF")  # chu trang tren nen mau nhu mau
        S.cmd("home.alignCenter", range=rg)
    S.cmd("data.validation", range="%s%d:%s%d" % (L(c_ev), r1, L(c_ev + n - 1), r2), type="list",
          formula1='"%s,%s,Chưa kiểm"' % (DAT, CHUA_DAT))
    for col, w in zip("ABCDE", [18, 18, 26, 34, 34]):
        S.width("%s:%s" % (col, col), w)
    for j in range(n):
        S.width("%s:%s" % (L(c_ev + j), L(c_ev + j)), 16)
        S.width("%s:%s" % (L(c_bc + j), L(c_bc + j)), 38)
    S.width("%s:%s" % (L(c_ng), L(c_ng)), 32)
    S.width("%s:%s" % (L(c_who), L(c_who)), 12)
    # chieu cao dong theo do dai chu (GridCraft khong tu co)
    for i, b in enumerate(body):
        widths = [18, 8, 26, 34, 34] + [16] * n + [38] * n + [32, 12]
        lines = 1
        for v, w in zip(b, widths):
            s = str(v or "")
            est = sum(max(1, -(-len(x) // max(int(w * 1.15), 1))) for x in s.split("\n")) if s else 1
            lines = max(lines, est)
        S.cmd("home.rowHeight", rows="%d:%d" % (r1 + i, r1 + i), points=min(12 * lines + 4, 300))
    # vung tong hop
    summaries = [{"name": "TH_Onpage_Diem", "range": absref(sh, 1, 1, 1 + n, last_sc),
                  "desc": "Điểm onpage theo nhóm × website (tiêu chí Đạt / (Đạt + Chưa đạt)); dòng cuối = tổng. "
                          "Radar ở '1. Định hướng' trỏ vào đây."}]
    r = r1
    while r <= r2:
        g = crit[r - r1]["nhom"]
        e = r
        while e < r2 and crit[e + 1 - r1]["nhom"] == g and e - r + 1 < 55:
            e += 1
        nm = "TH_Onpage_TC_%s_%d" % (ten_hop_le(_khong_dau(g)).strip("_"), r)
        summaries.append({"name": nm, "range": absref(sh, 1, r, c_who, e),
                          "desc": "Tiêu chí nhóm '%s' dòng %d-%d: đánh giá, bằng chứng từng website (tiêu đề cột ở dòng %d)"
                                  % (g, r, e, hdr_row)})
        r = e + 1
    for s_ in summaries:
        S.name(s_["name"], s_["range"])
    entry = {"sheet": sh, "module": "onpage",
             "purpose": "Chấm tiêu chí onpage (Đạt/Chưa đạt kèm bằng chứng) cho từng website; bảng điểm nhóm A1 theo mẫu",
             "data_sheet": None, "rows": len(crit),
             "columns": [{"name": "Đánh giá <website>", "desc": "Đạt / Chưa đạt / Chưa kiểm (trống = chờ Claude chấm)"},
                         {"name": "Bằng chứng <website>", "desc": "Con số đo được hoặc nguồn kiểm tay"},
                         {"name": "Người chấm", "desc": "Tool / Claude / Tool + Claude"}],
             "summaries": summaries, "options": [], "notes": [],
             "sites": sites, "claude_cells": claude_cells, "claude_items": claude_items,
             "criteria_header_row": hdr_row, "criteria_rows": [r1, r2],
             "eval_cols": [L(c_ev + j) for j in range(n)], "evidence_cols": [L(c_bc + j) for j in range(n)]}
    return entry, applied


def _khong_dau(s):
    import unicodedata
    s = (s or "").replace("đ", "d").replace("Đ", "D")
    s = unicodedata.normalize("NFD", s)
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def dung_traffic(S, data, mau):
    sh = SH_TRAFFIC
    sites = data["traffic_sites"]
    n = len(sites)
    S.sheet(sh)
    mr = max(mau["max_row"].get(sh, 40), 60)
    S.cmd("home.unmergeCells", range="A1:Z%d" % mr)
    S.cmd("edit.clearAll", range="A1:Z%d" % mr)
    # khoi kenh A:H theo mau: moi site 1 khoi 13 dong, chi giu tieu de
    kenh = ["Time Period", "Direct", "Display Ads", "Email", "Organic Search", "Paid Search", "Referrals", "Social"]
    for i, s in enumerate(sites):
        r = 1 + 13 * i
        S.set("A%d" % r, s)
        S.cmd("home.mergeCells", range="A%d:H%d" % (r, r))
        S.style("A%d:H%d" % (r, r), "h1")
        S.values("A%d" % (r + 1), [kenh])
        S.style("A%d:H%d" % (r + 1, r + 1), "h2")
        S.set("A%d" % (r + 2), "Chưa có dữ liệu kênh (cần Similarweb). Organic từ Ahrefs ở bảng bên phải (J:%s)." % L(10 + n))
        S.style("A%d" % (r + 2), "hint", align="left")
    # bang Organic Search J:M
    months = sorted({t for s in sites for t in data["traffic"].get(s, {})})
    S.set("J1", "Organic Search")
    S.cmd("home.mergeCells", range="J1:%s1" % L(10 + n))
    S.style("J1:%s1" % L(10 + n), "h1")
    S.values("J2", [["Time Period"] + sites])
    S.style("J2:%s2" % L(10 + n), "h2")
    # thang ghi dang chu "mm/yyyy" de truc bieu do hien dung (GridCraft chua dat duoc dinh dang so cho truc)
    rows = [[t.strftime("%m/%Y")] + [data["traffic"].get(s, {}).get(t, 0) for s in sites] for t in months]
    S.values("J3", rows)
    end = 2 + len(months)
    S.style("J3:J%d" % end, "lab")
    S.style("K3:%s%d" % (L(10 + n), end), "data", num="#,##0")
    s0 = end + 2
    S.values("J%d" % s0, [["Tổng hợp"] + sites])
    S.style("J%d:%s%d" % (s0, L(10 + n), s0), "h1")
    labels = ["TB 12 tháng gần nhất", "TB 12 tháng trước đó", "Tăng trưởng"]
    srows = []
    for k, lb in enumerate(labels):
        row = [lb]
        for j in range(n):
            col = L(11 + j)
            if k == 0:
                row.append("=AVERAGE(%s%d:%s%d)" % (col, end - 11, col, end) if len(months) >= 12 else "")
            elif k == 1:
                row.append("=AVERAGE(%s%d:%s%d)" % (col, end - 23, col, end - 12) if len(months) >= 24 else "")
            else:
                row.append("=IFERROR(%s%d/%s%d-1,0)" % (col, s0 + 1, col, s0 + 2) if len(months) >= 24 else "")
        srows.append(row)
    S.values("J%d" % (s0 + 1), srows)
    S.style("J%d:J%d" % (s0 + 1, s0 + 3), "lab", align="left")
    S.style("K%d:%s%d" % (s0 + 1, L(10 + n), s0 + 2), "data", num="#,##0")
    S.style("K%d:%s%d" % (s0 + 3, L(10 + n), s0 + 3), "data", num="0.0%")
    S.width("A:H", 11)
    S.width("I:I", 4)
    S.width("J:J", 20)
    S.width("K:%s" % L(10 + n), 16)
    summaries = [
        {"name": "TH_TRAFFIC_thang", "range": absref(sh, 10, 2, 10 + n, end),
         "desc": "Organic traffic theo tháng (dòng, cột J = tháng mm/yyyy) × website (cột K…); nguồn Ahrefs."},
        {"name": "TH_TRAFFIC_tong_hop", "range": absref(sh, 10, s0, 10 + n, s0 + 3),
         "desc": "TB 12 tháng gần nhất, TB 12 tháng trước đó và tăng trưởng, theo website."}]
    for s_ in summaries:
        S.name(s_["name"], s_["range"])
    # chart organic cua sheet
    for ch in mau["charts"].get(sh, []):
        role = _chart_role(sh, ch)
        if role == "traffic_organic":
            S.cmd("chart.selectData", chart=ch["id"], range="%s!K2:%s%d" % (q(sh), L(10 + n), end),
                  categories="%s!J3:J%d" % (q(sh), end))
            S.cmd("chart.move", chart=ch["id"], at="%s2" % L(12 + n))
        else:
            S.cmd("chart.delete", chart=ch["id"])
    entry = {"sheet": sh, "module": "traffic_ahrefs",
             "purpose": "Organic traffic theo tháng (Ahrefs, %d tháng) của các website, bảng J:%s theo mẫu" % (len(months), L(10 + n)),
             "data_sheet": None, "rows": len(months) * n,
             "columns": [{"name": "Time Period", "desc": "Tháng, chữ dạng mm/yyyy (cột J)"}] +
                        [{"name": s, "desc": "Organic traffic tháng của %s (cột %s)" % (s, L(11 + j))} for j, s in enumerate(sites)],
             "summaries": summaries, "options": [], "notes": [],
             "data_blocks": [{"site": "(tất cả)", "sheet": sh, "range": "J3:%s%d" % (L(10 + n), end), "header_row": 2,
                              "rows": len(months), "columns": ["Time Period"] + sites}]}
    return entry, {"organic": (sh, end, n)}


def dung_content(S, data, mau):
    sh = SH_CONTENT
    sites = data["tp_sites"]
    n = len(sites)
    S.sheet(sh)
    mr = max(mau["max_row"].get(sh, 100), 100)
    S.cmd("home.unmergeCells", range="A1:AZ%d" % mr)
    S.cmd("edit.clearAll", range="A1:AZ%d" % mr)
    # khoi du lieu tung site: URL | Traffic | Top keyword | Loai Trang | Topic | Top KW volume (cot H tro di, cach 1 cot)
    cols = ["URL", "Traffic", "Top keyword", "Loại Trang", "Topic", "Top KW volume"]
    W = len(cols)
    blocks = []
    tr_topic = {}
    for i, s in enumerate(sites):
        c0 = 8 + i * (W + 1)
        rows = sorted(data["tp"].get(s, []), key=lambda r: -(r[5] or 0))
        S.set("%s1" % L(c0), s)
        S.cmd("home.mergeCells", range="%s1:%s1" % (L(c0), L(c0 + W - 1)))
        S.style("%s1:%s1" % (L(c0), L(c0 + W - 1)), "h1")
        S.values("%s2" % L(c0), [cols])
        S.style("%s2:%s2" % (L(c0), L(c0 + W - 1)), "h2")
        vals = [[r[1], r[5] or 0, r[8] or "", r[2], r[3], r[9] or 0] for r in rows]
        for r in rows:
            tr_topic[r[3]] = tr_topic.get(r[3], 0) + (r[5] or 0)
        S.values("%s3" % L(c0), vals)
        last = max(3, 2 + len(vals))
        S.style("%s3:%s%d" % (L(c0), L(c0 + W - 1), last), "data", border="data", align=None)
        S.cmd("home.fontColor", range="%s3:%s%d" % (L(c0), L(c0), last), color=XANH_DAM)
        S.cmd("home.numberFormat", range="%s3:%s%d" % (L(c0 + 1), L(c0 + 1), last), code="#,##0")
        S.cmd("home.numberFormat", range="%s3:%s%d" % (L(c0 + 5), L(c0 + 5), last), code="#,##0")
        S.width("%s:%s" % (L(c0), L(c0)), 40)
        S.width("%s:%s" % (L(c0 + 1), L(c0 + 1)), 9)
        S.width("%s:%s" % (L(c0 + 2), L(c0 + 2)), 24)
        S.width("%s:%s" % (L(c0 + 3), L(c0 + 4)), 16)
        S.width("%s:%s" % (L(c0 + 5), L(c0 + 5)), 10)
        S.width("%s:%s" % (L(c0 + W), L(c0 + W)), 3)
        blocks.append({"site": s, "c0": c0, "last": last, "n": len(vals)})
    topics = sorted(tr_topic, key=lambda t: (-tr_topic[t], t))[:40]

    def col_rng(b, k):
        return "$%s$3:$%s$%d" % (L(b["c0"] + k), L(b["c0"] + k), b["last"])

    # so luong index (nhap tay)
    S.set("B1", "Số lượng index")
    S.style("B1", "h1")
    S.values("A2", [[s] for s in sites])
    S.style("A2:A%d" % (1 + n), "data", align="left")
    S.cmd("home.fontColor", range="A2:A%d" % (1 + n), color="#1155CC")
    S.style("B2:B%d" % (1 + n), "data")
    S.set("C1", "← nhập tay (Google site: hoặc GSC)")
    S.style("C1", "hint", align="left")
    summaries = []

    def bang(row, hdr0, labels, kind, key_k, name, desc):
        """Bang nhan × site. kind count|sum|pct(ref=(r0,rt)). key_k: cot dieu kien trong khoi (3=Loai, 4=Topic)."""
        S.values("A%d" % row, [[hdr0] + sites])
        S.style("A%d:%s%d" % (row, L(1 + n), row), "h1")
        r0 = row + 1
        body = []
        for i, lab in enumerate(labels):
            r = r0 + i
            line = [lab]
            for j, b in enumerate(blocks):
                if kind == "count":
                    line.append("=COUNTIFS(%s,$A%d)" % (col_rng(b, key_k), r))
                elif kind == "sum":
                    line.append("=SUMIFS(%s,%s,$A%d)" % (col_rng(b, 1), col_rng(b, key_k), r))
                else:
                    ref0, reft = kind[1]
                    line.append("=IFERROR(%s%d/%s$%d,0)" % (L(2 + j), ref0 + i, L(2 + j), reft))
            body.append(line)
        rt = r0 + len(labels)
        body.append(["Tổng"] + ["=SUM(%s%d:%s%d)" % (L(2 + j), r0, L(2 + j), rt - 1) for j in range(n)])
        S.values("A%d" % r0, body)
        S.style("A%d:A%d" % (r0, rt), "lab", align="left")
        num = "0.0%" if isinstance(kind, tuple) else "#,##0"
        S.style("B%d:%s%d" % (r0, L(1 + n), rt), "data", num=num)
        S.cmd("home.bold", range="A%d:%s%d" % (rt, L(1 + n), rt), on=True)
        rng = absref(sh, 1, row, 1 + n, rt)
        summaries.append({"name": name, "range": rng, "desc": desc})
        return r0, rt

    row = 6
    t1 = bang(row, "Topic", topics, "count", 4, "TH_Topic_SoTrang", "Số trang theo topic × website (dòng cuối = tổng)")
    chart_topic = (row, t1[1] - 1)       # header row .. last topic row
    row = t1[1] + 2
    chart_row = row
    row += 22                              # cho bieu do "So luong content theo chu de"
    t2 = bang(row, "Topic (traffic)", topics, "sum", 4, "TH_Topic_Traffic", "Traffic theo topic × website (dòng cuối = tổng)")
    row = t2[1] + 2
    l1 = bang(row, "Loại trang", LOAI_TRANG, "count", 3, "TH_LoaiTrang_SoTrang", "Số trang theo loại trang × website (dòng cuối = tổng)")
    row = l1[1] + 2
    l2 = bang(row, "Loại trang (traffic)", LOAI_TRANG, "sum", 3, "TH_LoaiTrang_Traffic", "Traffic theo loại trang × website (dòng cuối = tổng)")
    row = l2[1] + 2
    bang(row, "Loại trang (% traffic)", LOAI_TRANG, ("pct", (l2[0], l2[1])), 3, "TH_LoaiTrang_PctTraffic",
         "% traffic theo loại trang trong từng website")
    for s_ in summaries:
        S.name(s_["name"], s_["range"])
    S.width("A:A", 26)
    S.width("B:%s" % L(1 + n), 13)
    S.width("%s:G" % L(2 + n), 4)
    for ch in mau["charts"].get(sh, []):
        role = _chart_role(sh, ch)
        if role == "content_topic":
            S.cmd("chart.selectData", chart=ch["id"], range="%s!B%d:%s%d" % (q(sh), chart_topic[0], L(1 + n), chart_topic[1]),
                  categories="%s!A%d:A%d" % (q(sh), chart_topic[0] + 1, chart_topic[1]))
            S.cmd("chart.move", chart=ch["id"], at="A%d" % chart_row)
            S.cmd("chart.set", chart=ch["id"], width=560, height=330)
        else:
            S.cmd("chart.delete", chart=ch["id"])
    names = ["URL", "Traffic", "Top keyword", "Loại Trang", "Topic", "Top KW volume"]
    data_blocks = [{"site": b["site"], "sheet": sh,
                    "range": "%s3:%s%d" % (L(b["c0"]), L(b["c0"] + W - 1), b["last"]), "header_row": 2,
                    "rows": b["n"], "columns": names} for b in blocks]
    entry = {"sheet": sh, "module": "content_topic",
             "purpose": "Số trang và traffic theo topic và loại trang của từng website (Ahrefs Top pages), theo mẫu",
             "data_sheet": None, "rows": sum(b["n"] for b in blocks),
             "columns": [{"name": "URL", "desc": "URL trang"}, {"name": "Traffic", "desc": "Organic traffic ước tính/tháng"},
                         {"name": "Top keyword", "desc": "Từ khóa kéo traffic nhiều nhất"},
                         {"name": "Loại Trang", "desc": "Trang chủ/Danh mục/Sản phẩm-Dịch vụ/Blog/Landing/Chính sách-hỗ trợ/Khác"},
                         {"name": "Topic", "desc": "Chủ đề theo file luật"},
                         {"name": "Top KW volume", "desc": "Volume của top keyword"}],
             "summaries": summaries, "options": [], "notes": [], "data_blocks": data_blocks}
    return entry, {"topic": (sh, chart_topic[0], chart_topic[1], n)}


def dung_backlink(S, data, mau, opts):
    sh = SH_BACKLINK
    sites = data["rd_sites"]
    n = len(sites)
    nams = data["rd_years"]
    S.sheet(sh)
    mr = max(mau["max_row"].get(sh, 100), 100)
    S.cmd("home.unmergeCells", range="A1:BZ%d" % mr)
    S.cmd("edit.clearAll", range="A1:BZ%d" % mr)
    # ---- khoi du lieu tung site (cot M tro di)
    cols = ["Domain", "Spam", "DR", "Links to target", "First seen", "Lost", "Trạng thái", "Nhóm", "Năm"]
    W = len(cols)
    C0 = 13
    blocks = []
    for i, s in enumerate(sites):
        c0 = C0 + i * (W + 1)
        rows = data["rd"].get(s, [])
        S.set("%s1" % L(c0), s)
        S.cmd("home.mergeCells", range="%s1:%s1" % (L(c0), L(c0 + W - 1)))
        S.style("%s1:%s1" % (L(c0), L(c0 + W - 1)), "h1")
        S.values("%s2" % L(c0), [cols])
        S.style("%s2:%s2" % (L(c0), L(c0 + W - 1)), "h2", wrap=True)
        cD, cS, cDR, cFS, cLo = (L(c0 + k) for k in (0, 1, 2, 4, 5))
        vals = []
        for k, d in enumerate(rows):
            r = 3 + k
            # RD data: Website, Domain, Spam, DR, LTT, Dofollow, Traffic, First seen, Lost
            vals.append([d[1], d[2] or "", d[3] or 0, d[4] or 0, serial(d[7]), serial(d[8]),
                         '=IF(%s%d="","Còn","Mất")' % (cLo, r),
                         '=IF({S}{r}="Spam","Spam",IF({D}{r}=0,"0",IF({D}{r}<=30,"1-30",IF({D}{r}<=50,"30-50",'
                         'IF({D}{r}<=70,"50-70","70-100")))))'.format(S=cS, D=cDR, r=r),
                         '=IF(%s%d="","",YEAR(%s%d))' % (cFS, r, cFS, r)])
        S.values("%s3" % L(c0), vals)
        last = max(3, 2 + len(vals))
        S.style("%s3:%s%d" % (L(c0), L(c0 + W - 1), last), "data", border="data", align=None)
        S.cmd("home.numberFormat", range="%s3:%s%d" % (cFS, cLo, last), code="dd/mm/yyyy")
        for k, w in enumerate([24, 6, 5, 8, 10, 10, 8, 8, 6]):
            S.width("%s:%s" % (L(c0 + k), L(c0 + k)), w)
        S.width("%s:%s" % (L(c0 + W), L(c0 + W)), 3)
        blocks.append({"site": s, "c0": c0, "last": last, "n": len(vals)})
    S.width("L:L", 3)

    def R(b, k):
        return "$%s$3:$%s$%d" % (L(b["c0"] + k), L(b["c0"] + k), b["last"])

    # ---- o tuy chon
    chi_so = opts.get("chi_so") if opts.get("chi_so") in CHI_SO else CHI_SO[0]
    pham_vi = opts.get("pham_vi") if opts.get("pham_vi") in PHAM_VI else PHAM_VI[0]
    hien_spam = opts.get("hien_spam") if opts.get("hien_spam") in ("Có", "Không") else "Không"
    S.values("A1", [["Chỉ số", chi_so], ["Phạm vi", pham_vi], ["Hiện Spam", hien_spam]])
    S.style("A1:A3", "lab", align="left")
    S.style("B1:B3", "data")
    S.cmd("home.bold", range="B1:B3", on=True)
    for cell, vals in (("B1", CHI_SO), ("B2", PHAM_VI), ("B3", ["Có", "Không"])):
        S.cmd("data.validation", range=cell, type="list", formula1='"%s"' % ",".join(vals))
    S.set("D1", '=IF($B$2="%s","Còn","*")' % PHAM_VI[0])
    S.cmd("home.fontColor", range="D1", color="#FFFFFF")
    S.set("C1", "← chọn ở cột B (danh sách)")
    S.set("C3", "Không = bỏ Spam khỏi Tổng và biểu đồ (cột % Spam vẫn tính)")
    S.style("C1", "hint", align="left")
    S.style("C3", "hint", align="left")
    options = [
        {"cell": "%s!B1" % q(sh), "label": "Chỉ số", "values": CHI_SO, "default": chi_so},
        {"cell": "%s!B2" % q(sh), "label": "Phạm vi", "values": PHAM_VI, "default": pham_vi},
        {"cell": "%s!B3" % q(sh), "label": "Hiện Spam", "values": ["Có", "Không"], "default": hien_spam}]

    def cong_thuc(b, nhom_cell, nam_cell):
        dk = "%s,%s,%s,\"<\"&DATE(%s+1,1,1),%s,$D$1" % (R(b, 7), nhom_cell, R(b, 4), nam_cell, R(b, 6))
        return 'IF($B$1="Links to target",SUMIFS(%s,%s),COUNTIFS(%s))' % (R(b, 3), dk, dk)

    # ---- bang chi tiet tung site (tinh truoc vi tri)
    ny = len(nams)
    r_time = 5                                  # bang Tong theo nam × site
    r_q = r_time + ny + 2                       # bang chat luong nam cuoi × site
    r_charts = r_q + len(NHOM_BL) + 3           # 2 bieu do mau (thoi gian, chat luong)
    r_site0 = r_charts + 21
    site_tbl = []
    summaries = []
    row = r_site0
    for b in blocks:
        h = row + 1
        site_tbl.append((b, row, h, h + ny))
        row = h + ny + 2 + 20                   # bang + bieu do cot chong ben duoi
    # bang tong theo nam
    S.values("A%d" % r_time, [["Năm"] + sites])
    S.style("A%d:%s%d" % (r_time, L(1 + n), r_time), "h1")
    body = []
    for k, y in enumerate(nams):
        body.append([y] + ["=H%d" % (h + 1 + k) for (_b, _r, h, _e) in site_tbl])
    S.values("A%d" % (r_time + 1), body)
    S.style("A%d:A%d" % (r_time + 1, r_time + ny), "lab")
    S.style("B%d:%s%d" % (r_time + 1, L(1 + n), r_time + ny), "data", num="#,##0")
    summaries.append({"name": "TH_BL_theo_nam", "range": absref(sh, 1, r_time, 1 + n, r_time + ny),
                      "desc": "Tổng refdomain lũy kế đến hết từng năm × website (theo 3 ô chọn B1:B3)."})
    # bang chat luong nam cuoi
    nam_cuoi = nams[-1] if nams else dt.date.today().year
    S.values("A%d" % r_q, [["Nhóm (%d)" % nam_cuoi] + sites + ["%% %s" % s for s in sites]])
    S.style("A%d:%s%d" % (r_q, L(1 + 2 * n), r_q), "h1", wrap=True)
    body = []
    for g_i, g in enumerate(NHOM_BL + ["Tổng"]):
        line = [g]
        for (b, _r, h, e) in site_tbl:
            line.append("=%s%d" % (L(2 + g_i), e))
        for j in range(n):
            line.append("=IFERROR(%s%d/%s$%d,0)" % (L(2 + j), r_q + 1 + g_i, L(2 + j), r_q + 1 + len(NHOM_BL)))
        body.append(line)
    S.values("A%d" % (r_q + 1), body)
    rq_end = r_q + 1 + len(NHOM_BL)
    S.style("A%d:A%d" % (r_q + 1, rq_end), "lab")
    S.style("B%d:%s%d" % (r_q + 1, L(1 + n), rq_end), "data", num="#,##0")
    S.style("%s%d:%s%d" % (L(2 + n), r_q + 1, L(1 + 2 * n), rq_end), "data", num="0.0%")
    S.cmd("home.bold", range="A%d:%s%d" % (rq_end, L(1 + 2 * n), rq_end), on=True)
    summaries.append({"name": "TH_BL_so_sanh", "range": absref(sh, 1, r_q, 1 + 2 * n, rq_end),
                      "desc": "Năm %d: số refdomain theo nhóm chất lượng × website và tỉ trọng (%%) trong từng website." % nam_cuoi})
    # bang tung site
    heads = ["Năm"] + NHOM_BL + ["Tổng", "% DR>30 không spam", "% Spam"]
    for (b, r0, h, e) in site_tbl:
        S.set("A%d" % r0, "%s – lũy kế đến hết năm (first seen trước 1/1 năm sau)" % b["site"])
        S.cmd("home.mergeCells", range="A%d:J%d" % (r0, r0))
        S.style("A%d:J%d" % (r0, r0), "h1", align="left")
        S.values("A%d" % h, [heads])
        S.style("A%d:J%d" % (h, h), "h2", wrap=True)
        body = []
        for k, y in enumerate(nams):
            r = h + 1 + k
            spam_f = cong_thuc(b, "B$%d" % h, "$A%d" % r)
            line = [y, '=IF($B$3="Không",0,%s)' % spam_f]
            for j in range(1, len(NHOM_BL)):
                line.append("=" + cong_thuc(b, "%s$%d" % (L(2 + j), h), "$A%d" % r))
            line += ["=SUM(B%d:G%d)" % (r, r), "=IFERROR(SUM(E%d:G%d)/H%d,0)" % (r, r, r),
                     "=IFERROR((%s)/(SUM(C%d:G%d)+%s),0)" % (spam_f, r, r, spam_f)]
            body.append(line)
        S.values("A%d" % (h + 1), body)
        S.style("A%d:A%d" % (h + 1, e), "lab")
        S.style("B%d:H%d" % (h + 1, e), "data", num="#,##0")
        S.style("I%d:J%d" % (h + 1, e), "data", num="0.0%")
        S.cmd("home.bold", range="H%d:H%d" % (h + 1, e), on=True)
        summaries.append({"name": "TH_BL_" + ten_hop_le(b["site"]), "range": absref(sh, 1, h, 10, e),
                          "desc": "%s: refdomain lũy kế theo năm (dòng) × nhóm chất lượng (cột B:G), Tổng, %% DR>30 không spam, "
                                  "%% Spam. Theo 3 ô chọn B1:B3." % b["site"]})
        # bieu do cot chong cua site
        S.cmd("insert.chart", range="B%d:G%d" % (h, e), type="column", subtype="stacked",
              title="%s – chất lượng refdomain lũy kế theo năm" % b["site"], at="A%d" % (e + 2))
        S.cmd("chart.selectData", range="%s!B%d:G%d" % (q(sh), h, e), categories="%s!A%d:A%d" % (q(sh), h + 1, e))
        S.cmd("chart.changeColors", colors=["#C00000", "#BFBFBF", "#12D6F7", "#7FA6F7", "#004AEF", "#1F2329"])
        S.cmd("chart.set", legend="bottom", width=560, height=300)
    for s_ in summaries:
        S.name(s_["name"], s_["range"])
    S.width("A:A", 13)
    S.width("B:J", 10)
    S.width("K:K", 3)
    # 2 bieu do cua mau
    for ch in mau["charts"].get(sh, []):
        role = _chart_role(sh, ch)
        if role == "bl_time":
            S.cmd("chart.selectData", chart=ch["id"], range="%s!B%d:%s%d" % (q(sh), r_time, L(1 + n), r_time + ny),
                  categories="%s!A%d:A%d" % (q(sh), r_time + 1, r_time + ny))
            S.cmd("chart.move", chart=ch["id"], at="A%d" % r_charts)
            S.cmd("chart.set", chart=ch["id"], width=385, height=300)
        elif role == "bl_stack":
            S.cmd("chart.selectData", chart=ch["id"], range="%s!A%d:%s%d" % (q(sh), r_q, L(1 + n), rq_end - 1), byRows=True)
            S.cmd("chart.move", chart=ch["id"], at="F%d" % r_charts)
            S.cmd("chart.set", chart=ch["id"], width=385, height=300)
        else:
            S.cmd("chart.delete", chart=ch["id"])
    data_blocks = [{"site": b["site"], "sheet": sh,
                    "range": "%s3:%s%d" % (L(b["c0"]), L(b["c0"] + W - 1), b["last"]), "header_row": 2,
                    "rows": b["n"], "columns": cols} for b in blocks]
    entry = {"sheet": sh, "module": "backlink_rd",
             "purpose": "Đánh giá backlink theo Referring domains: lũy kế theo năm × nhóm chất lượng (bảng bên trái), "
                        "dữ liệu từng site ở các khối bên phải (cột %s trở đi)" % L(C0),
             "data_sheet": None, "rows": sum(b["n"] for b in blocks),
             "columns": [{"name": "Domain", "desc": "Referring domain"},
                         {"name": "Spam", "desc": "'Spam' nếu Ahrefs gắn nhãn spam, trống nếu không"},
                         {"name": "DR", "desc": "Domain Rating 0-100"},
                         {"name": "Links to target", "desc": "Số link từ domain này tới website"},
                         {"name": "First seen", "desc": "Ngày Ahrefs thấy link đầu tiên (số serial ngày)"},
                         {"name": "Lost", "desc": "Ngày mất link (trống = còn link)"},
                         {"name": "Trạng thái", "desc": "Còn/Mất (hàm từ Lost)"},
                         {"name": "Nhóm", "desc": "Spam, 0, 1-30, 30-50, 50-70, 70-100 (hàm từ Spam, DR)"},
                         {"name": "Năm", "desc": "Năm first seen (hàm)"}],
             "summaries": summaries, "options": options, "notes": [], "data_blocks": data_blocks,
             "sites": sites, "years": nams}
    return entry, {"time": (sh, r_time, ny, n), "stack": (sh, r_q, rq_end - 1, n)}


def dung_dinh_huong(S, mau, refs, client, sites, notes, deleted_sheets):
    sh = SH_DH
    S.sheet(sh)
    d = mau["dh"]
    dropped = 0
    # bo chart tro sheet da xoa, chart kenh traffic, chart khong con du lieu
    keep_roles = {}
    for ch in mau["charts"].get(sh, []):
        role = _chart_role(sh, ch)
        bad = any(s in deleted_sheets or s not in mau["sheets"] for s in ch["ref_sheets"])
        if bad or role in (None, "traffic_kenh"):
            S.cmd("chart.delete", chart=ch["id"])
            continue
        if role == "content_topic" and "topic" not in refs or role == "traffic_organic" and "organic" not in refs \
                or role in ("bl_time", "bl_stack") and "time" not in refs or role.startswith("radar") and "radar" not in refs:
            S.cmd("chart.delete", chart=ch["id"])
            continue
        keep_roles.setdefault(role, []).append(ch)
    # xoa muc II.1
    if d["r_ii1"] and d["r_ii2"] and d["r_ii2"] > d["r_ii1"]:
        S.cmd("home.deleteRows", rows="%d:%d" % (d["r_ii1"], d["r_ii2"] - 1))
        dropped = d["r_ii2"] - d["r_ii1"]

    def sh_row(r):
        return r - dropped if d["r_ii1"] and r >= d["r_ii2"] else r

    # tieu de
    title = d.get("title") or ""
    if title:
        t = re.sub(r"DỰ ÁN\s*\.+", "DỰ ÁN %s" % client.upper(), title)
        t = re.sub(r"\b20\d\d\b", str(dt.date.today().year), t)
        S.set("A5", t)
    # ten site tren radar
    if d.get("radar_row"):
        rr = sh_row(d["radar_row"])
        for j, col in enumerate(["B", "F", "J"]):
            S.set("%s%d" % (col, rr), sites["onpage"][j] if j < len(sites.get("onpage", [])) else "")
    # repoint chart
    for role, lst in keep_roles.items():
        for ch in lst:
            cid = ch["id"]
            if role.startswith("radar"):
                j = int(role.split(":")[1])
                o_sh, n_groups, n_sites = refs["radar"]
                if j >= n_sites:
                    S.cmd("chart.delete", chart=cid)
                    continue
                S.cmd("chart.selectData", chart=cid, range="%s!%s1:%s%d" % (q(o_sh), L(2 + j), L(2 + j), 1 + n_groups),
                      categories="%s!A2:A%d" % (q(o_sh), 1 + n_groups))
            elif role == "traffic_organic":
                t_sh, end, n = refs["organic"]
                S.cmd("chart.selectData", chart=cid, range="%s!K2:%s%d" % (q(t_sh), L(10 + n), end),
                      categories="%s!J3:J%d" % (q(t_sh), end))
                at = sh_row(int(re.match(r"[A-Z]+(\d+)", ch["at"]).group(1)))
                S.cmd("chart.move", chart=cid, at="A%d" % at)
                S.cmd("chart.set", chart=cid, width=1050)
            elif role == "content_topic":
                c_sh, r_h, r_e, n = refs["topic"]
                S.cmd("chart.selectData", chart=cid, range="%s!B%d:%s%d" % (q(c_sh), r_h, L(1 + n), r_e),
                      categories="%s!A%d:A%d" % (q(c_sh), r_h + 1, r_e))
                S.cmd("chart.set", chart=cid, width=1050)
            elif role == "bl_time":
                b_sh, r_t, ny, n = refs["time"]
                S.cmd("chart.selectData", chart=cid, range="%s!B%d:%s%d" % (q(b_sh), r_t, L(1 + n), r_t + ny),
                      categories="%s!A%d:A%d" % (q(b_sh), r_t + 1, r_t + ny))
            elif role == "bl_stack":
                b_sh, r_q, r_qe, n = refs["stack"]
                S.cmd("chart.selectData", chart=cid, range="%s!A%d:%s%d" % (q(b_sh), r_q, L(1 + n), r_qe), byRows=True)
    # o nhan xet
    notes_out = []
    applied = 0
    for nt in d["notes"]:
        r = sh_row(nt["row"])
        cell = "%s%d" % (L(nt["col"]), r)
        ref = "%s!%s" % (q(sh), cell)
        notes_out.append({"cell": ref, "label": nt["label"]})
        n_ = notes.get(norm(ref))
        if n_ and n_.get("text"):
            S.set(cell, n_["text"])
            lines = ["%s = %s" % (s, (n_.get("values") or ["?"] * 99)[i] if i < len(n_.get("values") or []) else "?")
                     for i, s in enumerate(n_.get("sources") or [])]
            if lines:
                S.cmd("review.newNote", cell=cell, text="Nguồn:\n" + "\n".join(lines), author="BD SEO")
            applied += 1
        else:
            S.set(cell, KHUNG_NHAN_XET)
    entry = {"sheet": sh, "module": "dinh_huong",
             "purpose": "Trang tổng hợp theo mẫu: biểu đồ từng mục (Onpage, Organic traffic, Content, Offpage) + 4 ô nhận xét "
                        "khung 'Dữ liệu nghiên cứu / Đánh giá / Định hướng'",
             "data_sheet": None, "rows": 0, "columns": [], "summaries": [], "options": [], "caveats": [],
             "notes": notes_out, "dropped_rows": dropped}
    return entry, applied


def dung_muc_luc(S, mau, kept):
    sh = SH_MUCLUC
    S.sheet(sh)
    keep, drop = [], []
    for r, nhom, hm in mau["muc_luc"]:
        target = MUC_LUC.get(hm.lower()) if nhom.lower().startswith("định hướng") else None
        if target and target in kept:
            keep.append((r, hm, target))
        else:
            drop.append(r)
    for r in sorted(drop, reverse=True):
        S.cmd("home.deleteRows", rows="%d:%d" % (r, r))
    for i, (_r, hm, target) in enumerate(keep):
        rr = 2 + i
        S.set("A%d" % rr, i + 1)
        S.cmd("insert.link", cell="C%d" % rr, target="%s!A1" % q(target), text=hm)


# ---------------------------------------------------------------- ham chinh
def do_vao_mau(built_xlsx, manifest, template_path, out_path, notes=None, opts=None):
    """Dung file theo mau tu file da build. Tra ve manifest moi (CONTRACT muc 3 + muc 7).
    manifest: manifest cua file da build (bo cuc tu dat). notes: notes.json cua file THEO MAU (o Dinh huong + o Claude
    cham onpage, toa do file theo mau). opts: {module_key: opts cua recipe} (vd backlink_rd.chi_so)."""
    notes = {norm(k): v for k, v in (notes or {}).items()}
    opts = opts or {}
    mau = doc_mau(template_path)
    data = doc_built(built_xlsx, manifest)
    present = set(data["entries"])
    kept = [SH_MUCLUC, SH_DH] + [MODULE_SHEET[k] for k in ("onpage", "traffic_ahrefs", "content_topic", "backlink_rd")
                                 if k in present and MODULE_SHEET[k] in mau["sheets"]]
    deleted = [s for s in mau["sheets"] if s not in kept]
    S = Script()
    for s in deleted:
        S.cmd("home.deleteSheet", sheet=s)
    entries, refs, warnings = [], {}, list(manifest.get("warnings") or [])
    applied = 0
    old = data["entries"]
    if "onpage" in present:
        e, a = dung_onpage(S, data, notes, mau)
        applied += a
        e["caveats"] = list(old["onpage"].get("caveats") or [])
        e["inputs"] = old["onpage"].get("inputs")
        entries.append(e)
        refs["radar"] = (SH_ONPAGE, len(data["onpage_groups"]), len(data["onpage_sites"]))
    if "traffic_ahrefs" in present:
        e, r = dung_traffic(S, data, mau)
        e["caveats"] = list(old["traffic_ahrefs"].get("caveats") or [])
        e["inputs"] = old["traffic_ahrefs"].get("inputs")
        entries.append(e)
        refs.update(r)
    if "content_topic" in present:
        e, r = dung_content(S, data, mau)
        e["caveats"] = list(old["content_topic"].get("caveats") or [])
        e["inputs"] = old["content_topic"].get("inputs")
        entries.append(e)
        refs.update(r)
    if "backlink_rd" in present:
        e, r = dung_backlink(S, data, mau, dict(opts.get("backlink_rd") or {}))
        e["caveats"] = list(old["backlink_rd"].get("caveats") or [])
        e["inputs"] = old["backlink_rd"].get("inputs")
        entries.append(e)
        refs.update(r)
    sites = {"onpage": data.get("onpage_sites") or []}
    e_dh, a = dung_dinh_huong(S, mau, refs, manifest.get("client", ""), sites, notes, deleted)
    applied += a
    dung_muc_luc(S, mau, kept)
    S.sheet(SH_DH)
    S.cmd("selection.set", range="A1")
    # ghi va chay
    fd, script_path = tempfile.mkstemp(suffix=".jsonl", prefix="bd_mau_")
    os.close(fd)
    try:
        S.dump(script_path)
        out_dir = os.path.dirname(os.path.abspath(out_path))
        os.makedirs(out_dir, exist_ok=True)
        tmp_out = os.path.join(out_dir, ".~gc-" + os.path.basename(out_path))
        gc_run(template_path, script_path, tmp_out)
        os.replace(tmp_out, out_path)
    finally:
        if os.environ.get("BD_KEEP_SCRIPT"):
            print("script:", script_path, file=sys.stderr)
        else:
            os.remove(script_path)
    entries = [e_dh] + entries
    # cac o tu cham khong con trong neu notes da dien
    known = {norm(n_["cell"]) for e in entries for n_ in e.get("notes", [])}
    known |= {norm(c) for e in entries for c in e.get("claude_cells", [])}
    for k in notes:
        if k not in known:
            warnings.append("Nhận xét không có ô tương ứng trong file theo mẫu, bỏ qua: %s" % k)
    return {
        "version": 1, "file": os.path.abspath(out_path), "client": manifest.get("client", ""),
        "created": dt.datetime.now().replace(microsecond=0).isoformat(),
        "sites": manifest.get("sites") or [], "layout": "mau_bd", "template": os.path.abspath(template_path),
        "sheets": entries, "warnings": warnings, "notes_applied": applied,
        "deleted_sheets": deleted,
    }


# ---------------------------------------------------------------- chuyen notes tu file bo cuc cu
def chuyen_notes(old_xlsx, new_xlsx, chi_onpage=True):
    """Chuyen notes onpage (o Claude cham) tu file bo cuc cu sang file theo mau, theo thu tu tieu chi + thu tu site.
    Nhan xet NHAN XET cu khong chuyen (khung moi o Dinh huong). Ghi <new>.notes.json, tra so o da chuyen."""
    om = json.load(open(old_xlsx + ".manifest.json", encoding="utf-8"))
    on = json.load(open(old_xlsx + ".notes.json", encoding="utf-8")) if os.path.exists(old_xlsx + ".notes.json") else {}
    nm = json.load(open(new_xlsx + ".manifest.json", encoding="utf-8"))
    o = [e for e in om["sheets"] if e["module"] == "onpage"]
    n = [e for e in nm["sheets"] if e["module"] == "onpage"]
    new_notes = json.load(open(new_xlsx + ".notes.json", encoding="utf-8")) if os.path.exists(new_xlsx + ".notes.json") else {}
    cnt = 0
    if o and n:
        o, n = o[0], n[0]
        or1 = o["criteria_rows"][0]
        nr1 = n["criteria_rows"][0]
        n_sites = len(o.get("sites") or [])
        oc_ev = 7  # cot G trong bo cuc cu
        for k, v in on.items():
            s, c = split_ref(k)
            if s != SH_ONPAGE:
                continue
            m = re.match(r"([A-Z]+)(\d+)$", c)
            ci, r = column_index_from_string(m.group(1)), int(m.group(2))
            if r < or1 or not (oc_ev <= ci < oc_ev + n_sites):
                continue
            j = ci - oc_ev
            new_ref = "%s!%s%d" % (q(SH_ONPAGE), n["eval_cols"][j], nr1 + (r - or1))
            new_notes[new_ref] = v
            cnt += 1
    with open(new_xlsx + ".notes.json", "w", encoding="utf-8") as f:
        json.dump(new_notes, f, ensure_ascii=False, indent=1)
    return cnt


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    a1 = sub.add_parser("do-vao-mau")
    a1.add_argument("--built", required=True)
    a1.add_argument("--manifest", required=True)
    a1.add_argument("--template", required=True)
    a1.add_argument("--out", required=True)
    a2 = sub.add_parser("chuyen-notes")
    a2.add_argument("--tu", required=True)
    a2.add_argument("--sang", required=True)
    a = ap.parse_args(argv)
    if a.cmd == "do-vao-mau":
        m = json.load(open(a.manifest, encoding="utf-8"))
        res = do_vao_mau(a.built, m, a.template, a.out)
        with open(a.out + ".manifest.json", "w", encoding="utf-8") as f:
            json.dump(res, f, ensure_ascii=False, indent=1)
        print("Đã ghi %s" % a.out)
    else:
        print("Đã chuyển %d ô" % chuyen_notes(os.path.abspath(a.tu), os.path.abspath(a.sang)))


if __name__ == "__main__":
    main()
