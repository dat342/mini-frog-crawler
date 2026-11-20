"""Orchestrator file BD SEO: dung nhieu sheet tu 1 recipe. Xem CONTRACT.md muc 2, 3, 5.

    .venv/bin/python -m xu_ly_data.bd_seo build --recipe recipe.json --out "<Khach> - BD SEO.xlsx"
    .venv/bin/python -m xu_ly_data.bd_seo rebuild "<Khach> - BD SEO.xlsx"

Python: build_from_recipe(recipe_dict, out_path) -> manifest dict.
Python 3.9: khong dung cu phap X | None.
"""
import argparse
import datetime as dt
import importlib
import json
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.comments import Comment

# thu tu sheet theo mau BD: 1.5 -> 1.6 -> 1.7 -> 1.9
MODULE_ORDER = ["onpage", "traffic_ahrefs", "content_topic", "backlink_rd", "volume_trend", "batch_overview"]
MANIFEST_SHEET = "_manifest"


def load_module(key):
    """Import dong xu_ly_data.<key>. Tra ve module hoac nem ImportError."""
    pkg = __package__ or "xu_ly_data"
    if key not in MODULE_ORDER:
        raise ImportError("Module không có trong registry: %s" % key)
    try:
        mod = importlib.import_module("%s.%s" % (pkg, key))
    except ImportError:
        if pkg == "xu_ly_data":
            raise
        mod = importlib.import_module("xu_ly_data.%s" % key)
    if not hasattr(mod, "build"):
        raise ImportError("Module %s chưa có hàm build()" % key)
    return mod


def available_modules():
    out = {}
    for k in MODULE_ORDER:
        try:
            m = load_module(k)
            out[k] = {"sheet": getattr(m, "SHEET", ""), "input_kind": getattr(m, "INPUT_KIND", ""),
                      "desc": getattr(m, "DESC", "")}
        except Exception as e:  # noqa
            out[k] = {"error": str(e)}
    return out


def _paths(out_path):
    return {"manifest": out_path + ".manifest.json", "recipe": out_path + ".recipe.json",
            "notes": out_path + ".notes.json", "cache": out_path + ".cache"}


def _split_ref(ref):
    m = re.match(r"^'?(.*?)'?!\$?([A-Z]+)\$?(\d+)$", ref.strip())
    if not m:
        raise ValueError("Tham chiếu ô không hợp lệ: %s" % ref)
    return m.group(1).replace("''", "'"), "%s%s" % (m.group(2), m.group(3))


def apply_notes(wb, notes):
    """Chen nhan xet: text vao o (o merge = o tren trai), comment liet ke sources = values."""
    applied, skipped = 0, []
    for ref, n in (notes or {}).items():
        try:
            sheet, cell = _split_ref(ref)
            ws = wb[sheet]
            ws[cell] = n.get("text", "")
            src = n.get("sources") or []
            val = n.get("values") or []
            lines = []
            for i, s in enumerate(src):
                v = val[i] if i < len(val) else "?"
                lines.append("%s = %s" % (s, v))
            if lines:
                c = Comment("Nguồn:\n" + "\n".join(lines), "BD SEO")
                c.width, c.height = 420, 40 + 18 * len(lines)
                ws[cell].comment = c
            applied += 1
        except Exception as e:  # noqa
            skipped.append("%s: %s" % (ref, e))
    return applied, skipped


def _sort_sheets(wb, entries):
    results = [e["sheet"] for e in entries]
    data = [e["data_sheet"] for e in entries if e.get("data_sheet")]
    order = results + data
    rest = [ws.title for ws in wb.worksheets if ws.title not in order and ws.title != MANIFEST_SHEET]
    names = [n for n in order + rest if n in wb.sheetnames]
    wb._sheets = [wb[n] for n in names]


def build_from_recipe(recipe, out_path, notes=None, cache_dir=None):
    """Dung file BD SEO tu recipe (dict). Ghi xlsx + manifest/recipe cua file. Tra ve manifest dict.
    recipe co "mau" (ten trong xu_ly_data/mau/ hoac duong dan) -> dung ra file tam roi dien vao mau gan the
    (CONTRACT muc 8). "template" cu (muc 7) duoc chuyen sang mau bd-seo."""
    if recipe.get("mau") or recipe.get("template"):
        return _build_mau_the(recipe, out_path, notes)
    out_path = os.path.abspath(out_path)
    p = _paths(out_path)
    if cache_dir:
        p["cache"] = cache_dir
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    os.makedirs(p["cache"], exist_ok=True)
    mods = recipe.get("modules") or []
    sites = list(recipe.get("sites") or [])
    ctx = {"out_path": out_path, "cache_dir": p["cache"], "sites": sites}
    wb = Workbook()
    wb.remove(wb.active)
    entries, warnings = [], []
    done = set()
    # xep theo thu tu sheet de nhat quan, module la thu tu MODULE_ORDER
    mods = sorted(mods, key=lambda m: MODULE_ORDER.index(m["key"]) if m.get("key") in MODULE_ORDER else 99)
    for m in mods:
        key = m.get("key")
        if key in done:
            warnings.append("Module %s khai báo nhiều lần, bỏ qua lần sau." % key)
            continue
        try:
            mod = load_module(key)
        except Exception as e:  # noqa
            warnings.append("Bỏ qua module %s: không import được (%s)" % (key, e))
            continue
        before = set(wb.sheetnames)
        try:
            inputs = list(m.get("inputs") or [])
            if sites and len(inputs) == len(sites) and all(os.path.sep in str(x) or "." in os.path.basename(str(x)) for x in inputs):
                # sap file theo thu tu site cua recipe (khach truoc) dua tren domain trong ten file
                from .common import ten_site_theo_file
                gan = ten_site_theo_file(inputs, sites)
                if sorted(gan) == sorted(sites) and gan != list(sites):
                    inputs = [inputs[gan.index(s)] for s in sites]
            entry = mod.build(wb, inputs, m.get("opts") or {}, ctx)
        except Exception as e:  # noqa
            for n in set(wb.sheetnames) - before:
                del wb[n]
            warnings.append("Module %s lỗi: %s" % (key, e))
            continue
        done.add(key)
        entries.append(entry)
    if notes is None and os.path.exists(p["notes"]):
        notes = json.load(open(p["notes"], encoding="utf-8"))
    notes_applied = 0
    if notes:
        notes_applied, skipped = apply_notes(wb, notes)
        for s in skipped:
            warnings.append("Nhận xét bỏ qua: " + s)
    _sort_sheets(wb, entries)
    manifest = {
        "version": 1, "file": out_path, "client": recipe.get("client", ""),
        "created": dt.datetime.now().replace(microsecond=0).isoformat(),
        "sites": sites, "sheets": entries, "warnings": warnings, "notes_applied": notes_applied,
    }
    text = json.dumps(manifest, ensure_ascii=False)
    ms = wb.create_sheet(MANIFEST_SHEET)
    # o Excel toi da 32767 ky tu: manifest dai thi chia nhieu o A1, A2, ...
    for i in range(0, len(text), 30000):
        ms.cell(1 + i // 30000, 1, text[i:i + 30000])
    ms.sheet_state = "hidden"
    wb.save(out_path)
    with open(p["manifest"], "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    with open(p["recipe"], "w", encoding="utf-8") as f:
        json.dump(recipe, f, ensure_ascii=False, indent=1)
    return manifest


def _build_mau_the(recipe, out_path, notes=None):
    """Dung ban tam bang cac module (khong chen nhan xet) roi the_mau.dien_mau. Nhan xet (notes.json cua file ra)
    chen o buoc dien mau: o nhan_xet (text + note nguon) va o Claude cham (text, gia tri vao cot bang chung)."""
    try:
        from . import the_mau
    except ImportError:
        import the_mau  # type: ignore
    out_path = os.path.abspath(out_path)
    p = _paths(out_path)
    recipe = dict(recipe)
    warn = []
    if not recipe.get("mau"):
        # tuong thich nguoc: "template" (do vao mau viet cung, template_bd) -> mau gan the bd-seo
        warn.append("recipe 'template' đã chuyển sang mẫu gắn thẻ 'bd-seo' (xu_ly_data/mau/bd-seo.xlsx).")
        recipe["mau"] = "bd-seo"
        recipe.pop("template", None)
    mau = the_mau.tim_mau(recipe["mau"])
    os.makedirs(p["cache"], exist_ok=True)
    tmp = os.path.join(p["cache"], "_ban-tam.xlsx")
    plain = dict(recipe)
    plain.pop("mau", None)
    m0 = build_from_recipe(plain, tmp, notes={}, cache_dir=p["cache"])
    if notes is None:
        notes = json.load(open(p["notes"], encoding="utf-8")) if os.path.exists(p["notes"]) else {}
    manifest = the_mau.dien_mau(tmp, m0, mau, out_path, recipe, notes=notes)
    manifest["warnings"] = warn + manifest["warnings"]
    with open(p["manifest"], "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    with open(p["recipe"], "w", encoding="utf-8") as f:
        json.dump(recipe, f, ensure_ascii=False, indent=1)
    if not os.path.exists(p["notes"]):
        with open(p["notes"], "w", encoding="utf-8") as f:
            json.dump({}, f)
    for f_ in (tmp, tmp + ".manifest.json", tmp + ".recipe.json", tmp + ".notes.json"):
        if os.path.exists(f_):
            os.remove(f_)
    return manifest


def rebuild(path):
    path = os.path.abspath(path)
    p = _paths(path)
    if not os.path.exists(p["recipe"]):
        raise SystemExit("Không thấy recipe: %s" % p["recipe"])
    recipe = json.load(open(p["recipe"], encoding="utf-8"))
    return build_from_recipe(recipe, path)


def _summary(m):
    lines = ["Đã ghi: %s%s" % (m["file"], " (mẫu: %s)" % os.path.basename(m["mau"]) if m.get("layout") == "mau_the" else "")]
    for e in m["sheets"]:
        lines.append("  - %s (%s): %d dòng data, %d vùng tổng hợp, %d ô nhận xét" % (
            e["sheet"], e["module"], e.get("rows") or 0, len(e.get("summaries") or []), len(e.get("notes") or [])))
        for nt in e.get("notes") or []:
            lines.append("      nhận xét %s [%s]" % (nt["cell"], nt.get("label", "")))
        for c in e.get("caveats") or []:
            lines.append("      ! " + c)
    for w in m["warnings"]:
        lines.append("CẢNH BÁO: " + w)
    return "\n".join(lines)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build")
    b.add_argument("--recipe", required=True)
    b.add_argument("--out", required=True)
    r = sub.add_parser("rebuild")
    r.add_argument("file")
    sub.add_parser("modules")
    a = ap.parse_args(argv)
    if a.cmd == "build":
        recipe = json.load(open(a.recipe, encoding="utf-8"))
        print(_summary(build_from_recipe(recipe, a.out)))
    elif a.cmd == "rebuild":
        print(_summary(rebuild(a.file)))
    else:
        print(json.dumps(available_modules(), ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
