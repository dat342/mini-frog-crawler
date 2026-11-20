"""Tao lai mau gan the xu_ly_data/mau/bd-seo.xlsx tu file mau BD goc cua nguoi dung (chi doc file goc).

    .venv/bin/python -m xu_ly_data.mau.tao_bd_seo ["/abs/BD SEO _ c06_2026).xlsx"]

Giu 6 sheet (Muc luc, 1. Dinh huong, 1.5, 1.6, 1.7, 1.9) theo bo cuc da duyet 20/11/2026, xoa du lieu khach cu
va moi bieu do, dat the CONTRACT muc 8. Moi thay doi qua GridCraft (giu logo, mau, dinh dang cua mau goc).
"""
import os
import sys
import tempfile

from xu_ly_data import template_bd as T

GOC = os.environ.get("BD_SEO_GOC", os.path.expanduser("~/Downloads/Claude/SEONGON/5-inbox/BD SEO _ c06_2026).xlsx"))
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bd-seo.xlsx")
KEEP = [T.SH_MUCLUC, T.SH_DH, T.SH_ONPAGE, T.SH_TRAFFIC, T.SH_CONTENT, T.SH_BACKLINK]


def tao(goc=GOC, out=OUT):
    mau = T.doc_mau(goc)
    S = T.Script()
    for s in mau["sheets"]:
        if s not in KEEP:
            S.cmd("home.deleteSheet", sheet=s)
    for sh in KEEP[1:]:
        for ch in mau["charts"].get(sh, []):
            S.sheet(sh)
            S.cmd("chart.delete", chart=ch["id"])
    # ---- Muc luc: chi giu hang muc Dinh huong ung voi 4 sheet (kem link)
    T.dung_muc_luc(S, mau, KEEP)
    # ---- 1. Dinh huong
    d = mau["dh"]
    sh = T.SH_DH
    S.sheet(sh)
    S.cmd("home.deleteRows", rows="%d:%d" % (d["r_ii1"], d["r_ii2"] - 1))
    drop = d["r_ii2"] - d["r_ii1"]
    S.set("A5", "ĐỊNH HƯỚNG SEO | DỰ ÁN {{khach}} | {{nam}}")
    rr = d["radar_row"] - drop                       # hang ten site tren radar (20)
    for i, col in enumerate(["B", "F", "J"]):
        S.set("%s%d" % (col, rr), "{{site:%d}}" % (i + 1))
    S.set("A%d" % (rr + 1), "{{bieu_do:op_diem|loai=radar|site=*|buoc=4|rong=300|cao=255|tieu_de={{ten}}}}")
    keys = {"Onpage": "onpage", "Organic traffic": "organic_traffic", "Content": "content", "Offpage": "offpage"}
    rows = {}
    for n in d["notes"]:
        r = n["row"] - drop
        rows[n["label"]] = r
        S.set("%s%d" % (T.L(n["col"]), r), "{{nhan_xet:%s}}" % keys.get(n["label"], n["label"]))
    # bieu do: dong ngay tren o nhan xet cua tung muc
    S.set("A%d" % (rows["Organic traffic"] - 2),
          "{{bieu_do:tr_thang|loai=line|rong=1050|cao=280|tieu_de=So sánh Organic traffic}}")
    S.set("A%d" % (rows["Content"] - 2),
          "{{bieu_do:ct_topic_sotrang|loai=col|rong=1050|cao=350|tieu_de=Số lượng content theo chủ đề}}")
    S.set("B%d" % (rows["Offpage"] - 3),
          "{{bieu_do:bl_theo_nam|loai=col|rong=420|cao=250|tieu_de=Số lượng backlink theo thời gian}}")
    S.set("H%d" % (rows["Offpage"] - 3),
          "{{bieu_do:bl_chat_luong|loai=col_stacked|chuoi=dong|rong=420|cao=250|tieu_de=Chất lượng backlink}}")

    def reset(sh, cols, mr):
        S.sheet(sh)
        S.cmd("home.unmergeCells", range="A1:%s%d" % (cols, mr))
        S.cmd("edit.clearAll", range="A1:%s%d" % (cols, mr))
        S.cmd("home.rowHeight", rows="1:%d" % min(mr, 300), points=15)

    # ---- 1.5 Onpage
    sh = T.SH_ONPAGE
    reset(sh, "Z", max(mau["max_row"][sh], 120))
    S.set("A1", "{{bang:op_diem}}")
    S.style("A1:D1", "h1", wrap=True)
    S.style("A2", "lab", wrap=True, align="left")
    S.cmd("home.bold", range="A2", on=True)
    S.style("B2", "data")
    S.set("A9", "{{bang:op_tieu_chi}}")
    S.style("A9:M9", "h1", wrap=True)
    for j, c in enumerate("FGH"):
        S.cmd("home.fillColor", range="%s9" % c, color=T.EVAL_HDR_FILL[j])
    S.style("A10:M10", "data", wrap=True, align="left")
    S.cmd("home.alignTop", range="A10:M10")
    for j, c in enumerate("FGH"):
        S.cmd("home.fillColor", range="%s10" % c, color=T.EVAL_FILL[j])
        S.cmd("home.fontColor", range="%s10" % c, color="#FFFFFF")
        S.cmd("home.alignCenter", range="%s10" % c)
    for col, w in zip("ABCDE", [18, 18, 26, 34, 34]):
        S.width("%s:%s" % (col, col), w)
    S.width("F:H", 16)
    S.width("I:K", 38)
    S.width("L:L", 32)
    S.width("M:M", 12)
    # ---- 1.6 traffic
    sh = T.SH_TRAFFIC
    reset(sh, "Z", 60)
    kenh = ["Time Period", "Direct", "Display Ads", "Email", "Organic Search", "Paid Search", "Referrals", "Social"]
    for i in range(3):
        r = 1 + 13 * i
        S.set("A%d" % r, "{{site:%d}}" % (i + 1))
        S.cmd("home.mergeCells", range="A%d:H%d" % (r, r))
        S.style("A%d:H%d" % (r, r), "h1")
        S.values("A%d" % (r + 1), [kenh])
        S.style("A%d:H%d" % (r + 1, r + 1), "h2")
        S.set("A%d" % (r + 2), "Chưa có dữ liệu kênh (cần Similarweb). Organic từ Ahrefs ở bảng bên phải (cột J).")
        S.style("A%d" % (r + 2), "hint", align="left")
    S.set("J1", "Organic Search")
    S.cmd("home.mergeCells", range="J1:M1")
    S.style("J1:M1", "h1")
    S.set("J2", "{{bang:tr_thang}}")
    S.style("J2:M2", "h2")
    S.style("J3", "lab")
    S.style("K3", "data")
    S.set("J40", "{{bang:tr_tong_hop}}")
    S.style("J40:M40", "h1")
    S.style("J41", "lab", align="left")
    S.style("K41", "data")
    S.set("O2", "{{bieu_do:tr_thang|loai=line|rong=640|cao=320|tieu_de=So sánh Organic traffic}}")
    S.width("A:H", 11)
    S.width("I:I", 4)
    S.width("J:J", 20)
    S.width("K:M", 16)
    # ---- 1.7 content
    sh = T.SH_CONTENT
    reset(sh, "AZ", max(mau["max_row"][sh], 120))
    S.set("B1", "Số lượng index")
    S.style("B1", "h1")
    S.values("A2", [["{{site:1}}"], ["{{site:2}}"], ["{{site:3}}"]])
    S.style("A2:A4", "data", align="left")
    S.cmd("home.fontColor", range="A2:A4", color="#1155CC")
    S.style("B2:B4", "data")
    S.set("C1", "← nhập tay (Google site: hoặc GSC)")
    S.style("C1", "hint", align="left")
    for r, key in ((6, "ct_topic_sotrang"), (48, "ct_topic_traffic"), (68, "ct_loaitrang_sotrang"),
                   (78, "ct_loaitrang_traffic"), (88, "ct_loaitrang_pct")):
        S.set("A%d" % r, "{{bang:%s}}" % key)
        S.style("A%d:D%d" % (r, r), "h1")
        S.style("A%d" % (r + 1), "lab", align="left")
        S.style("B%d" % (r + 1), "data")
    S.set("A26", "{{bieu_do:ct_topic_sotrang|loai=col|rong=560|cao=330|tieu_de=Số lượng content theo chủ đề}}")
    S.set("H1", "{{du_lieu:ct_data|kieu=canh_nhau}}")
    S.style("H1:M1", "h1")
    S.style("H2:M2", "h2")
    S.style("H3:M3", "data", border="data", align=None)
    S.cmd("home.fontColor", range="H3", color=T.XANH_DAM)
    for c, w in zip("HIJKLMN", [40, 9, 24, 16, 16, 10, 3]):
        S.width("%s:%s" % (c, c), w)
    S.width("A:A", 26)
    S.width("B:D", 13)
    S.width("E:G", 4)
    # ---- 1.9 backlink
    sh = T.SH_BACKLINK
    reset(sh, "BZ", 200)
    S.cmd("edit.clearAll", range="A1:BZ%d" % mau["max_row"][sh])
    S.set("A1", "{{bang:bl_tuy_chon|bo_tieu_de=1}}")
    S.style("A1", "lab", align="left")
    S.style("B1", "data")
    S.cmd("home.bold", range="B1", on=True)
    S.set("C1", "← chọn ở cột B (danh sách)")
    S.set("C3", "Không = bỏ Spam khỏi Tổng và biểu đồ (cột % Spam vẫn tính)")
    S.style("C1", "hint", align="left")
    S.style("C3", "hint", align="left")
    for r, key, c0 in ((5, "bl_theo_nam", "A"), (14, "bl_chat_luong", "A"), (14, "bl_chat_luong_pct", "F")):
        c1 = chr(ord(c0) + 3)
        S.set("%s%d" % (c0, r), "{{bang:%s}}" % key)
        S.style("%s%d:%s%d" % (c0, r, c1, r), "h1", wrap=True)
        S.style("%s%d" % (c0, r + 1), "lab")
        S.style("%s%d" % (chr(ord(c0) + 1), r + 1), "data")
    S.set("A23", "{{bieu_do:bl_theo_nam|loai=col|rong=385|cao=300|tieu_de=Số lượng backlink theo thời gian}}")
    S.set("F23", "{{bieu_do:bl_chat_luong|loai=col_stacked|chuoi=dong|rong=385|cao=300|tieu_de=Chất lượng backlink}}")
    for i in range(3):
        r = 44 + 30 * i
        S.set("A%d" % r, "{{site:%d}} – lũy kế đến hết năm (first seen trước 1/1 năm sau)" % (i + 1))
        S.cmd("home.mergeCells", range="A%d:J%d" % (r, r))
        S.style("A%d:J%d" % (r, r), "h1", align="left")
        S.set("A%d" % (r + 1), "{{bang:bl_site|site=%d}}" % (i + 1))
        S.style("A%d:J%d" % (r + 1, r + 1), "h2", wrap=True)
        S.style("A%d" % (r + 2), "lab")
        S.style("B%d" % (r + 2), "data")
        S.set("A%d" % (r + 10), "{{bieu_do:bl_site|site=%d|loai=col_stacked|rong=560|cao=300}}" % (i + 1))
    S.set("M1", "{{du_lieu:bl_data|kieu=canh_nhau}}")
    S.style("M1:U1", "h1")
    S.style("M2:U2", "h2", wrap=True)
    S.style("M3:U3", "data", border="data", align=None)
    for k, w in enumerate([24, 6, 5, 8, 10, 10, 8, 8, 6, 3]):
        S.width("%s:%s" % (T.L(13 + k), T.L(13 + k)), w)
    S.width("L:L", 3)
    S.width("A:A", 13)
    S.width("B:J", 10)
    S.width("K:K", 3)
    S.sheet(T.SH_MUCLUC)
    S.cmd("selection.set", range="A1")
    fd, sp = tempfile.mkstemp(suffix=".jsonl")
    os.close(fd)
    S.dump(sp)
    try:
        T.gc_run(goc, sp, out)
    finally:
        os.remove(sp)
    return out


if __name__ == "__main__":
    print(tao(*(sys.argv[1:2] or [GOC])))
