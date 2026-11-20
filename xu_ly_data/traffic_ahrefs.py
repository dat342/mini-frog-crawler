"""Keo organic traffic theo thang (Ahrefs API v3, metrics-history) cho N website -> sheet 1.6 bang ham.

Dung:
    AHREFS_API_KEY=... .venv/bin/python -m xu_ly_data.traffic_ahrefs --out traffic.xlsx seongon.com seosona.com gtvseo.com
Key doc theo thu tu: bien moi truong AHREFS_API_KEY -> file ~/.config/seongon/ahrefs.key (KHONG luu trong repo).
Chi phi do thuc te 20/11/2026: ~11 don vi/thang/site -> 36 thang ~407 don vi/site.
Ket qua tho luu cache JSON canh file output de khong goi lai (ton don vi) khi dung lai bang.
Python 3.9: khong dung cu phap X | None.
"""
import argparse
import datetime as dt
import json
import os

import httpx
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from . import common
except ImportError:  # chay truc tiep
    import common  # type: ignore

KEY = "traffic_ahrefs"
SHEET = "1.6 Dữ liệu traffic"
INPUT_KIND = "domains_api"
DESC = "Organic traffic theo tháng của các website (Ahrefs API, 36 tháng)"
DATA_SHEET = "Traffic data"

API = "https://api.ahrefs.com/v3/site-explorer/metrics-history"
XANH = "0B3D91"
HEADER_FILL = PatternFill("solid", fgColor="DCE6F2")


def doc_key():
    k = os.environ.get("AHREFS_API_KEY", "").strip()
    if k:
        return k
    p = os.path.expanduser("~/.config/seongon/ahrefs.key")
    if os.path.exists(p):
        return open(p).read().strip()
    raise SystemExit("Chưa có API key: đặt AHREFS_API_KEY hoặc tạo ~/.config/seongon/ahrefs.key (chmod 600)")


def keo(site, date_from, key):
    r = httpx.get(API, params={"target": site, "mode": "subdomains", "date_from": date_from,
                               "history_grouping": "monthly", "select": "date,org_traffic"},
                  headers={"Authorization": "Bearer " + key}, timeout=60)
    r.raise_for_status()
    return r.json()["metrics"], int(r.headers.get("x-api-units-cost-total", 0))


def _date_from(thang):
    today = dt.date.today().replace(day=1)
    y, m = today.year, today.month - (thang - 1)
    while m <= 0:
        y, m = y - 1, m + 12
    return "%04d-%02d-01" % (y, m)


def _lay_cache(sites, thang, cache_path):
    """Doc cache theo site; chi goi API khi cache chua phu tu date_from. Tra ve (cache, units)."""
    date_from = _date_from(thang)
    cache = json.load(open(cache_path, encoding="utf-8")) if os.path.exists(cache_path) else {}
    for k in list(cache):
        if "|" in k:  # dinh dang cache cu "site|date_from"
            cache.setdefault(k.split("|")[0], cache.pop(k))
    key = None
    units = 0
    for s in sites:
        rows = cache.get(s) or []
        if not rows or min(r["date"][:10] for r in rows) > date_from:
            try:
                key = key or doc_key()
            except SystemExit as e:
                raise RuntimeError("Cache traffic thiếu dữ liệu cho %s từ %s và chưa có API key: %s" % (s, date_from, e))
            cache[s], u = keo(s, date_from, key)
            units += u
    os.makedirs(os.path.dirname(os.path.abspath(cache_path)), exist_ok=True)
    json.dump(cache, open(cache_path, "w", encoding="utf-8"))
    return cache, units


def _caveats(sites, cache, date_from):
    out = ["Chỉ organic (Ahrefs ước tính), không có Direct/Social/Referral/Paid."]
    for s in sites:
        rows = sorted((r["date"][:7], r.get("org_traffic") or 0) for r in cache[s] if r["date"][:10] >= date_from)
        if not rows:
            out.append("%s: không có dữ liệu traffic." % s)
            continue
        v = [x[1] for x in rows]
        if len(v) >= 24:
            cur, prev = sum(v[-12:]), sum(v[-24:-12])
            if prev:
                out.append("%s: traffic 12 tháng gần nhất %s so với 12 tháng trước (%s%.0f%%)." % (
                    s, "cao hơn" if cur >= prev else "thấp hơn", "+" if cur >= prev else "", 100.0 * (cur / prev - 1)))
        hi = max(rows, key=lambda x: x[1])
        lo = min(rows, key=lambda x: x[1])
        out.append("%s: tháng đỉnh %s (%s), tháng đáy %s (%s)." % (
            s, hi[0][5:] + "/" + hi[0][:4], format(hi[1], ","), lo[0][5:] + "/" + lo[0][:4], format(lo[1], ",")))
        if len(v) >= 2 and v[-1] < 0.5 * v[-2]:
            out.append("%s: tháng cuối (%s) giảm hơn một nửa so với tháng trước, có thể là tháng chưa trọn." % (s, rows[-1][0]))
    return out


def build(wb, inputs, opts, ctx):
    opts = opts or {}
    sites = list(inputs)
    thang = int(opts.get("thang", 36))
    date_from = _date_from(thang)
    cache_dir = (ctx or {}).get("cache_dir") or "."
    cache, units = _lay_cache(sites, thang, os.path.join(cache_dir, "traffic.json"))

    ws = wb.create_sheet(SHEET)
    raw = wb.create_sheet(DATA_SHEET)
    raw.append(["Website", "Tháng", "Organic traffic", "Kỳ"])
    n = 1
    kys = {}
    for s in sites:
        for row in cache[s]:
            if row["date"][:10] < date_from:
                continue
            n += 1
            d = dt.datetime.strptime(row["date"][:10], "%Y-%m-%d")
            ky = d.strftime("%m/%Y")
            kys[ky] = d
            raw.append([s, d, row.get("org_traffic") or 0, ky])
            raw.cell(n, 2).number_format = "mm/yyyy"
    for c in raw[1]:
        c.font = Font(bold=True)
        c.fill = HEADER_FILL
    last = max(n, 2)

    common.title(ws, "ORGANIC TRAFFIC THEO THÁNG – %d THÁNG GẦN NHẤT (Ahrefs)" % thang,
                 "Nguồn: Ahrefs API v3 › metrics-history (org_traffic, theo tháng). Dữ liệu thô ở sheet Traffic data.")
    h = 4
    common.header_row(ws, h, 1, ["Time Period"] + sites)
    thangs = sorted(kys, key=lambda k: kys[k])
    R = lambda col: "'%s'!$%s$2:$%s$%d" % (DATA_SHEET, col, col, last)
    for i, t in enumerate(thangs):
        r = h + 1 + i
        ws.cell(r, 1, t)
        for j in range(len(sites)):
            col = get_column_letter(2 + j)
            ws.cell(r, 2 + j, "=SUMIFS(%s,%s,%s$%d,%s,$A%d)" % (R("C"), R("A"), col, h, R("D"), r)).number_format = "#,##0"
    end = h + len(thangs)
    lastcol = get_column_letter(1 + len(sites))
    summaries = [common.bang(
        wb, ws, "tr_thang", "A%d:%s%d" % (h, lastcol, end),
        "Organic traffic theo tháng (dòng, cột 1 = tháng dạng chữ mm/yyyy) × website (cột).", name="TRAFFIC_thang")]
    so = {"tr_tb12": [], "tr_tang_truong": [], "tr_dinh": []}
    s0 = end + 2
    common.header_row(ws, s0, 1, ["Tổng hợp"] + sites)
    labels = ["TB 12 tháng gần nhất", "TB 12 tháng trước đó", "Tăng trưởng"]
    for k, lb in enumerate(labels):
        ws.cell(s0 + 1 + k, 1, lb).font = Font(bold=True)
    for j in range(len(sites)):
        col = get_column_letter(2 + j)
        if len(thangs) >= 12:
            ws.cell(s0 + 1, 2 + j, "=AVERAGE(%s%d:%s%d)" % (col, end - 11, col, end)).number_format = "#,##0"
        if len(thangs) >= 24:
            ws.cell(s0 + 2, 2 + j, "=AVERAGE(%s%d:%s%d)" % (col, end - 23, col, end - 12)).number_format = "#,##0"
            ws.cell(s0 + 3, 2 + j, "=IFERROR(%s%d/%s%d-1,0)" % (col, s0 + 1, col, s0 + 2)).number_format = "0.0%"
        so["tr_tb12"].append(common.so_cell(ws, "%s%d" % (col, s0 + 1), sites[j]))
        so["tr_tang_truong"].append(common.so_cell(ws, "%s%d" % (col, s0 + 3), sites[j]))
        so["tr_dinh"].append({"site": sites[j], "formula": "=MAX(%s%d:%s%d)" % (col, h + 1, col, end),
                              "sheet": SHEET, "num_fmt": "#,##0"})
    summaries.append(common.bang(
        wb, ws, "tr_tong_hop", "A%d:%s%d" % (s0, lastcol, s0 + 3),
        "TB traffic 12 tháng gần nhất, TB 12 tháng trước đó và tăng trưởng, theo website.", name="TRAFFIC_tong_hop"))
    ch = LineChart()
    ch.title = "Organic traffic theo tháng"
    ch.y_axis.numFmt = "#,##0"
    ch.y_axis.majorGridlines = None
    ch.add_data(Reference(ws, min_col=2, max_col=1 + len(sites), min_row=h, max_row=end), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=h + 1, max_row=end))
    ch.legend.position = "b"
    ch.height, ch.width = 9, 22
    ws.add_chart(ch, "%s4" % get_column_letter(len(sites) + 3))
    note = common.note_area(ws, "A%d" % (s0 + 6), 10, 8, "NHẬN XÉT")
    ws.column_dimensions["A"].width = 22
    for j in range(len(sites)):
        ws.column_dimensions[get_column_letter(2 + j)].width = 16
    columns = [{"name": "Website", "desc": "Website"},
               {"name": "Tháng", "desc": "Ngày đầu tháng"},
               {"name": "Organic traffic", "desc": "Organic traffic ước tính của Ahrefs trong tháng"},
               {"name": "Kỳ", "desc": "Tháng dạng chữ mm/yyyy (khóa nối với bảng tr_thang)"}]
    charts = [{"key": "tr_thang", "bang": "tr_thang", "loai": "line", "tieu_de": "So sánh Organic traffic"}]
    du_lieu = [{"key": "tr_data", "sheet": DATA_SHEET, "site_col": "Website",
                "cot": [["Kỳ", "Tháng"], ["Organic traffic", "Organic traffic"]]}]
    return common.entry(
        SHEET, KEY, DESC, sites, data_sheet=DATA_SHEET, rows=last - 1, columns=columns,
        summaries=summaries, caveats=_caveats(sites, cache, date_from), notes=[note],
        extra={"units_used": units, "sites": sites, "so": so, "charts": charts, "du_lieu": du_lieu})


def dung_file(sites, out_path, thang=36):
    """CLI cu: cache JSON canh file output (<out>.cache/traffic.json)."""
    wb = Workbook()
    wb.remove(wb.active)
    cache_dir = os.path.splitext(out_path)[0] + ".cache"
    e = build(wb, sites, {"thang": thang}, {"out_path": out_path, "cache_dir": cache_dir})
    wb.save(out_path)
    return {"sites": sites, "months": e["rows"] // max(len(sites), 1), "units_used": e["units_used"], "out": out_path}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("sites", nargs="+")
    ap.add_argument("--out", required=True)
    ap.add_argument("--thang", type=int, default=36)
    a = ap.parse_args()
    print(dung_file(a.sites, a.out, a.thang))


if __name__ == "__main__":
    main()
