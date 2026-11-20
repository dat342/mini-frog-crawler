#!/usr/bin/env python3
"""MCP server "seongon-bd" (stdio, JSON-RPC 2.0 newline-delimited). Python 3.9, chỉ thư viện chuẩn + openpyxl.

Claude đọc manifest + bảng tổng hợp, hỏi sâu bằng hàm (GridCraft CLI eval), ghi nhận xét có dẫn nguồn.
"""
import csv
import datetime
import io
import json
import os
import re
import subprocess
import sys

SERVER_NAME = "seongon-bd"
SERVER_VERSION = "0.1.0"
SUPPORTED = ["2025-06-18", "2025-03-26", "2024-11-05"]

ROOT = os.path.abspath(os.path.expanduser(os.environ.get("SEONGON_ROOT", "~/Downloads/Claude/SEONGON")))
REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VENV_PY = os.path.join(REPO, ".venv", "bin", "python")
GRIDCRAFT = os.environ.get(
    "GRIDCRAFT_CLI",
    os.path.expanduser("~/.local/gridcraft/gridcraft-cli-0.19.1-macos-universal/gridcraft-cli"))
CLI_TIMEOUT = 60

MAX_TH_ROWS, MAX_TH_COLS = 60, 15
MAX_Q_ROWS, MAX_Q_CHARS = 50, 4000
MAX_SAMPLE = 20


class ToolError(Exception):
    pass


def log(*a):
    print("[seongon-bd]", *a, file=sys.stderr, flush=True)


# ---------------------------------------------------------------- tiện ích
def check_path(path):
    if not isinstance(path, str) or not path:
        raise ToolError("path phải là chuỗi đường dẫn tuyệt đối tới file .xlsx")
    if not os.path.isabs(path):
        raise ToolError("path phải là đường dẫn tuyệt đối: %s" % path)
    real = os.path.realpath(path)
    if not (real == ROOT or real.startswith(ROOT + os.sep)):
        raise ToolError("path phải nằm dưới %s" % ROOT)
    if not real.lower().endswith(".xlsx"):
        raise ToolError("chỉ chấp nhận file .xlsx")
    if not os.path.isfile(real):
        raise ToolError("không tìm thấy file: %s" % path)
    return real


def run_cli(args):
    try:
        p = subprocess.run([GRIDCRAFT] + args, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                           timeout=CLI_TIMEOUT, universal_newlines=True)
    except subprocess.TimeoutExpired:
        raise ToolError("GridCraft quá %ds, hãy thu hẹp công thức/vùng." % CLI_TIMEOUT)
    except OSError as e:
        raise ToolError("không chạy được GridCraft CLI (%s): %s" % (GRIDCRAFT, e))
    return p.returncode, p.stdout, p.stderr.strip()


def eval_formula(path, formula, sheet=None):
    """Trả giá trị Python (số, chuỗi, list 2D, None). Lỗi hàm -> ToolError nguyên văn."""
    if not isinstance(formula, str) or not formula.strip():
        raise ToolError("công thức rỗng")
    f = formula.strip()
    if not f.startswith("="):
        f = "=" + f
    args = ["eval", f, "--in", path, "--json"]
    if sheet:
        args += ["--sheet", sheet]
    rc, out, err = run_cli(args)
    out = out.strip()
    if not out:
        raise ToolError("GridCraft lỗi: %s" % (err or "không có kết quả"))
    try:
        val = json.loads(out)
    except ValueError:
        raise ToolError("GridCraft trả kết quả lạ: %s %s" % (out[:300], err))
    if isinstance(val, dict) and "error" in val:
        raise ToolError("hàm trả lỗi %s" % val["error"])
    return val


def fmt_num(x):
    if isinstance(x, float):
        if x == int(x) and abs(x) < 1e15:
            return str(int(x))
        return ("%.6g" % x) if abs(x) < 1e6 else repr(x)
    return str(x)


def fmt_val(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return fmt_num(v)
    return str(v)


def col_letter(n):  # 1-based
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def col_index(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


REF_RE = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$")


def parse_ref(ref):
    """'Sheet'!$A$8:$J$15 -> (sheet|None, c1, r1, c2, r2)."""
    m = REF_RE.match(ref.strip())
    if not m:
        raise ToolError("địa chỉ không hợp lệ: %s (dạng 'Sheet'!A1:J20)" % ref)
    sheet = m.group(1).replace("''", "'") if m.group(1) else m.group(2)
    c1, r1 = col_index(m.group(3)), int(m.group(4))
    c2 = col_index(m.group(5)) if m.group(5) else c1
    r2 = int(m.group(6)) if m.group(6) else r1
    return sheet, c1, r1, c2, r2


def norm_cell(ref):
    try:
        sheet, c1, r1, c2, r2 = parse_ref(ref)
    except ToolError:
        return ref.strip()
    return "%s!%s%d" % (sheet, col_letter(c1), r1)


def load_manifest(real):
    mp = real + ".manifest.json"
    if os.path.isfile(mp):
        try:
            with open(mp, encoding="utf-8") as f:
                return json.load(f)
        except ValueError as e:
            raise ToolError("manifest.json hỏng: %s" % e)
    try:
        import openpyxl
        wb = openpyxl.load_workbook(real, read_only=True, data_only=True)
        if "_manifest" in wb.sheetnames:
            v = wb["_manifest"]["A1"].value
            wb.close()
            if v:
                return json.loads(v)
        else:
            wb.close()
    except Exception as e:  # noqa
        raise ToolError("không đọc được manifest (%s)" % e)
    raise ToolError("file chưa có manifest (<file>.manifest.json hoặc sheet ẩn _manifest). "
                    "File này không phải đầu ra của xu_ly_data.bd_seo?")


def load_notes(real):
    p = real + ".notes.json"
    if os.path.isfile(p):
        try:
            with open(p, encoding="utf-8") as f:
                return json.load(f)
        except ValueError:
            return {}
    return {}


def all_sheets(manifest):
    return manifest.get("sheets", [])


def note_cells(manifest):
    """cell(chuẩn hoá) -> label, gồm notes + claude_cells."""
    res = {}
    for sh in all_sheets(manifest):
        for n in sh.get("notes", []) or []:
            res[norm_cell(n["cell"])] = n.get("label", "")
        for n in sh.get("claude_cells", []) or []:
            if isinstance(n, dict):
                res[norm_cell(n["cell"])] = n.get("label", "")
            else:
                res[norm_cell(str(n))] = ""
    return res


def q_sheet(name):
    return "'%s'" % name.replace("'", "''")


def find_sheet(manifest, name):
    for sh in all_sheets(manifest):
        if sh.get("sheet") == name or sh.get("data_sheet") == name:
            return sh
    return None


# ---------------------------------------------------------------- tool 1
def t_mo_ta_file(a):
    real = check_path(a.get("path"))
    m = load_manifest(real)
    notes = load_notes(real)
    notes_n = {norm_cell(k): v for k, v in notes.items()}
    L = []
    L.append("Khách: %s | Site: %s | Tạo: %s" % (m.get("client", "?"), ", ".join(m.get("sites", [])),
                                                  m.get("created", "?")))
    for sh in all_sheets(m):
        L.append("")
        L.append("## Sheet '%s' (%s): %s" % (sh.get("sheet"), sh.get("module", ""), sh.get("purpose", "")))
        if sh.get("data_sheet"):
            cols = "; ".join("%s%s" % (c["name"], (" = " + c["desc"][:45]) if c.get("desc") else "")
                             for c in sh.get("columns", []))
            letters = ""
            L.append("Dữ liệu thô: sheet '%s', %s dòng (dòng 1 tiêu đề). Cột theo thứ tự A,B,C...: %s"
                     % (sh["data_sheet"], sh.get("rows", "?"), cols))
        for b in sh.get("data_blocks", []) or []:
            try:
                _s, c1, r1, c2, r2 = parse_ref("%s!%s" % (q_sheet(b["sheet"]), b["range"]))
                cols = ", ".join("%s=%s" % (col_letter(c1 + i), n) for i, n in enumerate(b.get("columns", [])))
            except ToolError:
                cols = ", ".join(b.get("columns", []))
            L.append("- Khối dữ liệu [%s]: %s!%s (%s dòng, tiêu đề ở dòng %s). Cột: %s" % (
                b.get("site", ""), q_sheet(b["sheet"]), b["range"], b.get("rows", "?"), b.get("header_row", "?"), cols))
        for s in sh.get("summaries", []):
            tag = ""
            if s.get("key"):
                tag = " [key=%s%s]" % (s["key"], (", site=" + s["site"]) if s.get("site") else "")
            L.append("- Tổng hợp %s%s: %s | %s" % (s["name"], tag, s["range"], s.get("desc", "")))
        for o in sh.get("options", []):
            cur = None
            try:
                cur = eval_formula(real, "=" + o["cell"])
            except ToolError as e:
                cur = "?(%s)" % e
            L.append("- Tùy chọn %s (%s) hiện = %s; chọn trong %s" % (o["cell"], o.get("label", ""),
                                                                      fmt_val(cur), o.get("values")))
        for ch in sh.get("charts", []) or []:
            L.append("- Biểu đồ %s%s (%s) tại %s!%s, dữ liệu %s" % (
                ch.get("key"), (" [" + ch["site"] + "]") if ch.get("site") else "", ch.get("loai", ""),
                q_sheet(ch.get("sheet", "")), ch.get("at", ""), ch.get("range", "")))
        for c in sh.get("caveats", []):
            L.append("! Lưu ý: %s" % c)
        for n in sh.get("notes", []) or []:
            has = norm_cell(n["cell"]) in notes_n
            L.append("- Ô nhận xét %s [%s]: %s" % (n["cell"], n.get("label", ""),
                                                   "ĐÃ có" if has else "chưa có"))
        cc = sh.get("claude_cells", []) or []
        if cc:
            # gom gon theo cot de mo_ta_file khong phinh (onpage co ~100 o)
            pending = {}
            done = 0
            for n in cc:
                cell = n["cell"] if isinstance(n, dict) else str(n)
                if norm_cell(cell) in notes_n:
                    done += 1
                    continue
                m = re.search(r"!\$?([A-Z]+)\$?(\d+)$", cell)
                if m:
                    pending.setdefault(m.group(1), []).append(int(m.group(2)))
            parts = ["%s: %s" % (col, ",".join(str(r) for r in sorted(rows))) for col, rows in sorted(pending.items())]
            L.append("- Ô Claude chấm: %d ô, đã có %d. Còn trống (cột: dòng) — %s" % (len(cc), done, "; ".join(parts) or "không"))
    return "\n".join(L)


# ---------------------------------------------------------------- tool 2
def read_csv_range(real, sheet, ref_range):
    rc, out, err = run_cli(["cat", real, "--sheet", sheet, "--range", ref_range, "--csv"])
    if rc != 0:
        raise ToolError("GridCraft cat lỗi: %s %s" % (err, out[:200]))
    return list(csv.reader(io.StringIO(out)))


def t_doc_tong_hop(a):
    real = check_path(a.get("path"))
    name = a.get("name")
    if not isinstance(name, str) or not name:
        raise ToolError("thiếu name")
    ref = None
    desc = ""
    if "!" in name:
        ref = name
    else:
        m = load_manifest(real)
        for sh in all_sheets(m):
            for s in sh.get("summaries", []):
                if s["name"] == name:
                    ref, desc = s["range"], s.get("desc", "")
        if not ref:
            names = [s["name"] for sh in all_sheets(m) for s in sh.get("summaries", [])]
            raise ToolError("không có vùng tổng hợp '%s'. Có: %s" % (name, ", ".join(names)))
    sheet, c1, r1, c2, r2 = parse_ref(ref)
    if not sheet:
        raise ToolError("range cần có tên sheet: 'Sheet'!A1:J20")
    cut = []
    if r2 - r1 + 1 > MAX_TH_ROWS:
        cut.append("dòng %d-%d" % (r1 + MAX_TH_ROWS, r2))
        r2 = r1 + MAX_TH_ROWS - 1
    if c2 - c1 + 1 > MAX_TH_COLS:
        cut.append("cột %s-%s" % (col_letter(c1 + MAX_TH_COLS), col_letter(c2)))
        c2 = c1 + MAX_TH_COLS - 1
    rng = "%s%d:%s%d" % (col_letter(c1), r1, col_letter(c2), r2)
    rows = read_csv_range(real, sheet, rng)
    out = ["Vùng '%s'!%s%s (giá trị đã tính; ô = cột + dòng ở hàng/cột đầu)" % (
        sheet, rng, (" - " + desc) if desc else "")]
    out.append("dòng\t" + "\t".join(col_letter(c) for c in range(c1, c2 + 1)))
    for i in range(r2 - r1 + 1):
        r = rows[i] if i < len(rows) else []
        r = list(r) + [""] * (c2 - c1 + 1 - len(r))
        out.append(str(r1 + i) + "\t" + "\t".join(x.replace("\t", " ").replace("\n", " ") for x in r))
    if cut:
        out.append("[đã cắt %s vì vượt giới hạn %d dòng x %d cột; hãy đọc range nhỏ hơn]" % (
            ", ".join(cut), MAX_TH_ROWS, MAX_TH_COLS))
    return "\n".join(out)


# ---------------------------------------------------------------- tool 3
def render_value(val):
    if isinstance(val, list):
        total = len(val)
        lines = []
        for r in val:
            lines.append("\t".join(fmt_val(x) for x in r) if isinstance(r, list) else fmt_val(r))
        text = "\n".join(lines)
        note = ""
        if total > MAX_Q_ROWS:
            lines = lines[:MAX_Q_ROWS]
            text = "\n".join(lines)
            note = "[đã cắt: %d dòng, chỉ hiện %d; hãy dùng TAKE/FILTER chặt hơn]" % (total, MAX_Q_ROWS)
        if len(text) > MAX_Q_CHARS:
            text = text[:MAX_Q_CHARS]
            note = "[đã cắt ở %d ký tự; hãy dùng TAKE/FILTER chặt hơn]" % MAX_Q_CHARS
        return text + ("\n" + note if note else ""), total
    t = fmt_val(val)
    if len(t) > MAX_Q_CHARS:
        t = t[:MAX_Q_CHARS] + "\n[đã cắt, hãy dùng TAKE/FILTER chặt hơn]"
    return t, 1


def t_hoi(a):
    real = check_path(a.get("path"))
    val = eval_formula(real, a.get("formula"), a.get("sheet"))
    text, n = render_value(val)
    if isinstance(val, list):
        return "(mảng %d dòng)\n%s" % (n, text)
    return text


# ---------------------------------------------------------------- tool 4
OPS = {"=": "=", "==": "=", "!=": "<>", "<>": "<>", ">": ">", ">=": ">=", "<": "<", "<=": "<="}


def xl_lit(v):
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return repr(v)
    return '"%s"' % str(v).replace('"', '""')


def _vung_du_lieu(real, m, a):
    """Xác định vùng dữ liệu cho xem_mau -> (sheet, c1, r1, c2, r2, names).
    Bố cục cũ: data_sheet dạng bảng dọc (A1 tiêu đề). Bố cục theo mẫu: data_blocks (mỗi site 1 khối) -> chọn bằng
    tham số site hoặc range."""
    sheet = a.get("sheet")
    rng = a.get("range")
    site = a.get("site")
    sh = find_sheet(m, sheet)
    if sh is None and rng and "!" in rng:
        sh = find_sheet(m, parse_ref(rng)[0])
    blocks = (sh or {}).get("data_blocks") or []
    if blocks:
        cand = None
        if rng:
            rs, c1, r1, c2, r2 = parse_ref(rng if "!" in rng else "%s!%s" % (q_sheet(sh["sheet"]), rng))
            for b in blocks:
                _s, bc1, br1, bc2, br2 = parse_ref("%s!%s" % (q_sheet(b["sheet"]), b["range"]))
                if b["sheet"] == (rs or sh["sheet"]) and bc1 <= c1 <= bc2:
                    cand = b
                    break
            if cand is None:
                hdr = read_csv_range(real, rs or sh["sheet"], "%s%d:%s%d" % (col_letter(c1), r1 - 1, col_letter(c2), r1 - 1))
                names = [x.strip() for x in (hdr[0] if hdr else [])] or [col_letter(c) for c in range(c1, c2 + 1)]
                return rs or sh["sheet"], c1, r1, c2, r2, names
        elif site:
            cand = next((b for b in blocks if b.get("site") == site), None)
            if cand is None:
                raise ToolError("không có khối dữ liệu của site '%s'. Có: %s" % (site, ", ".join(b.get("site", "") for b in blocks)))
        elif len(blocks) == 1:
            cand = blocks[0]
        else:
            raise ToolError("sheet '%s' có %d khối dữ liệu (mỗi site 1 khối): truyền site (%s) hoặc range." % (
                sh["sheet"], len(blocks), ", ".join(b.get("site", "") for b in blocks)))
        _s, c1, r1, c2, r2 = parse_ref("%s!%s" % (q_sheet(cand["sheet"]), cand["range"]))
        return cand["sheet"], c1, r1, c2, r2, list(cand.get("columns") or [])
    if not sh or not sh.get("data_sheet"):
        raise ToolError("sheet '%s' không có data_sheet/data_blocks trong manifest (truyền tên sheet dữ liệu, vd 'RD data')" % sheet)
    names = [c["name"] for c in sh.get("columns", [])]
    if not names:
        raise ToolError("manifest không có danh sách cột")
    last = int(sh.get("rows", 0)) + 1
    if last < 2:
        raise ToolError("data_sheet không có dòng dữ liệu")
    return sh["data_sheet"], 1, 2, len(names), last, names


def t_xem_mau(a):
    real = check_path(a.get("path"))
    m = load_manifest(real)
    ds, c1, r1, c2, last, names = _vung_du_lieu(real, m, a)
    if len(names) != c2 - c1 + 1:
        names = (names + [col_letter(c) for c in range(c1 + len(names), c2 + 1)])[:c2 - c1 + 1]
    q = q_sheet(ds)
    full = "%s!$%s$%d:$%s$%d" % (q, col_letter(c1), r1, col_letter(c2), last)

    def colrng(name):
        if name not in names:
            raise ToolError("không có cột '%s'. Cột: %s" % (name, ", ".join(names)))
        L = col_letter(c1 + names.index(name))
        return "%s!$%s$%d:$%s$%d" % (q, L, r1, L, last)

    conds = a.get("dieu_kien") or []
    if not isinstance(conds, list):
        raise ToolError("dieu_kien phải là list")
    parts = []
    for c in conds:
        if not isinstance(c, dict) or "cot" not in c or "gt" not in c:
            raise ToolError("mỗi điều kiện cần cot, op, gt")
        op = c.get("op", "=")
        r = colrng(c["cot"])
        if op in ("chua", "contains"):
            parts.append("ISNUMBER(SEARCH(%s,%s))" % (xl_lit(c["gt"]), r))
        elif op in OPS:
            parts.append("(%s%s%s)" % (r, OPS[op], xl_lit(c["gt"])))
        else:
            raise ToolError("op không hợp lệ: %s (=, !=, >, >=, <, <=, chua)" % op)
    cond = "*".join("(%s)" % p for p in parts) if parts else None

    # đếm
    cnt_f = "=SUMPRODUCT(%s)" % ("--" + cond if cond else "--(%s<>\"\")" % colrng(names[0]))
    if cond:
        cnt_f = "=SUMPRODUCT(%s*1)" % cond
    total = int(eval_formula(real, cnt_f))

    cots = a.get("cot") or names
    if not isinstance(cots, list):
        raise ToolError("cot phải là list tên cột")
    idx = [names.index(c) + 1 if c in names else None for c in cots]
    if None in idx:
        bad = [c for c, i in zip(cots, idx) if i is None]
        raise ToolError("không có cột %s. Cột: %s" % (bad, ", ".join(names)))
    n = a.get("n", 10)
    try:
        n = max(1, min(MAX_SAMPLE, int(n)))
    except (TypeError, ValueError):
        raise ToolError("n phải là số nguyên")
    rows = []
    if total > 0:
        inner = "FILTER(%s,%s)" % (full, cond) if cond else full
        sort = a.get("sap_xep")
        if sort:
            if isinstance(sort, dict):
                sc, so = sort.get("cot"), sort.get("thu_tu", "giam")
            else:
                sc, _, so = str(sort).partition(":")
                so = so or "giam"
            if sc not in names:
                raise ToolError("sap_xep: không có cột '%s'" % sc)
            order = -1 if so.lower() in ("giam", "desc", "-1") else 1
            inner = "SORT(%s,%d,%d)" % (inner, names.index(sc) + 1, order)
        if total == 1 and not cond:
            pass
        f = "=TAKE(CHOOSECOLS(%s,%s),%d)" % (inner, ",".join(str(i) for i in idx), n)
        val = eval_formula(real, f)
        if not isinstance(val, list):
            val = [[val]]
        rows = val[:MAX_SAMPLE]
    out = ["Có %d dòng thỏa điều kiện; hiện %d dòng (dữ liệu %s!%s%d:%s%d; ngày hiển thị dạng số serial)." % (
        total, len(rows), q, col_letter(c1), r1, col_letter(c2), last), "\t".join(cots)]
    for r in rows:
        out.append("\t".join(fmt_val(x) for x in r))
    return "\n".join(out)


# ---------------------------------------------------------------- tool 5
def source_empty(v):
    if v is None or v == "":
        return True
    if isinstance(v, list):
        return all(source_empty(x) for r in v for x in (r if isinstance(r, list) else [r])) if v else True
    return False


def compact_value(v):
    if isinstance(v, list):
        flat = []
        for r in v:
            flat.append(",".join(fmt_val(x) for x in r) if isinstance(r, list) else fmt_val(r))
        s = " | ".join(flat)
        return s if len(s) <= 300 else s[:300] + "..."
    return fmt_val(v)


def t_ghi_nhan_dinh(a):
    real = check_path(a.get("path"))
    cell = a.get("cell")
    text = a.get("text")
    sources = a.get("sources")
    if not isinstance(text, str) or not text.strip():
        raise ToolError("text rỗng")
    if not isinstance(sources, list) or len(sources) < 1:
        raise ToolError("sources phải có ≥1 phần tử (địa chỉ ô hoặc công thức bắt đầu bằng '=')")
    m = load_manifest(real)
    allowed = note_cells(m)
    key = norm_cell(cell) if isinstance(cell, str) else None
    if key not in allowed:
        raise ToolError("cell '%s' không phải ô nhận xét trong manifest. Ô hợp lệ: %s" % (cell, ", ".join(allowed)))
    vals = []
    bad = []
    for s in sources:
        if not isinstance(s, str) or not s.strip():
            bad.append("nguồn rỗng")
            continue
        s = s.strip()
        f = s if s.startswith("=") else "=" + s
        if not s.startswith("="):
            parse_ref(s)  # phải là địa chỉ hợp lệ
        try:
            v = eval_formula(real, f)
        except ToolError as e:
            bad.append("%s -> %s" % (s, e))
            continue
        if source_empty(v):
            bad.append("%s -> rỗng" % s)
            continue
        vals.append(compact_value(v))
    if bad:
        raise ToolError("TỪ CHỐI ghi, nguồn lỗi/rỗng: " + "; ".join(bad))
    notes = load_notes(real)
    # xoá khoá cũ cùng ô nếu viết khác dạng
    for k in list(notes):
        if norm_cell(k) == key:
            del notes[k]
    notes[cell.strip()] = {"text": text.strip(), "sources": [s.strip() for s in sources], "values": vals,
                           "updated": datetime.datetime.now().replace(microsecond=0).isoformat()}
    tmp = real + ".notes.json.tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(notes, f, ensure_ascii=False, indent=1)
    os.replace(tmp, real + ".notes.json")
    lines = ["Đã ghi nhận xét vào %s.notes.json (chưa sửa xlsx; gọi chot_file để chèn). Đối chiếu số trong câu chữ với giá trị nguồn:" % os.path.basename(real)]
    for s, v in zip(sources, vals):
        lines.append("- %s = %s" % (s, v))
    return "\n".join(lines)


# ---------------------------------------------------------------- tool 6
def t_chot_file(a):
    real = check_path(a.get("path"))
    if not os.path.isfile(VENV_PY):
        raise ToolError("không thấy venv python: %s" % VENV_PY)
    try:
        p = subprocess.run([VENV_PY, "-m", "xu_ly_data.bd_seo", "rebuild", real], cwd=REPO,
                           stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=600, universal_newlines=True)
    except subprocess.TimeoutExpired:
        raise ToolError("rebuild quá 600s")
    so, se = p.stdout.strip(), p.stderr.strip()
    if p.returncode != 0:
        if "No module named" in se and "bd_seo" in se:
            raise ToolError("module xu_ly_data.bd_seo chưa có (đang được viết). stderr: %s" % se[-300:])
        raise ToolError("rebuild lỗi (rc=%d). stdout: %s | stderr: %s" % (p.returncode, so[-600:], se[-600:]))
    m = load_manifest(real)
    notes = {norm_cell(k) for k in load_notes(real)}
    empty = [c for c in note_cells(m) if c not in notes]
    return "rebuild OK.\n%s%s\nÔ nhận xét còn trống: %s" % (
        so[-800:], ("\n[stderr] " + se[-400:]) if se else "", ", ".join(empty) if empty else "không còn")


# ---------------------------------------------------------------- tool 7
GUIDE = """Quy trình chuẩn làm nhận xét cho file BD SEO (không đọc dữ liệu thô):
1. mo_ta_file(path): đọc mục đích, vùng tổng hợp, tùy chọn hiện tại, LƯU Ý dữ liệu (caveats), ô nhận xét chưa có.
2. doc_tong_hop(path, name) cho từng vùng TH_...: đọc số đã tính. Ghi nhớ địa chỉ ô (cột + dòng) để dẫn nguồn.
3. hoi(path, formula): kiểm điểm lạ bằng COUNTIFS/SUMIFS/AVERAGEIFS/FILTER/SORT/TAKE/UNIQUE trên data_sheet (file theo mẫu: trên các "Khối dữ liệu" mỗi site, địa chỉ cột in trong mo_ta_file). Ví dụ: tỉ lệ spam theo năm, top domain DR cao, đột biến một năm. Không kéo cả bảng thô; xem_mau chỉ để xem vài dòng minh họa (tối đa 20).
4. ghi_nhan_dinh(path, cell, text, sources): mỗi câu nhận định phải có số cụ thể, nguồn là ô ('Sheet'!G13) hoặc công thức '=...'. Server tự tính lại nguồn; nguồn lỗi/rỗng bị từ chối. So số trong text với values trả về. Nêu cả điểm bất lợi cho khách, không chỉ điểm tốt. Đọc caveats trước khi kết luận (vd dữ liệu spam dồn một năm làm sai xu hướng).
5. Khi người dùng đưa sẵn kết luận (vd "đối thủ A mạnh hơn"), PHẢI thử luận điểm ngược bằng hàm: tìm số liệu bác lại, rồi mới ghi; nếu số không ủng hộ thì nói thẳng.
6. chot_file(path): chèn nhận xét vào xlsx (rebuild), kiểm tra ô nhận xét còn trống, rồi báo người dùng.
Quy ước: công thức chạy trên cả Excel và Google Sheets; ngày trong eval là số serial; tránh ~ * ? trong điều kiện COUNTIFS."""


def t_huong_dan(a):
    return GUIDE


# ---------------------------------------------------------------- đăng ký tool
PATH_PROP = {"type": "string", "description": "Đường dẫn tuyệt đối tới file .xlsx (dưới thư mục gốc SEONGON_ROOT, mặc định ~/Downloads/Claude/SEONGON)."}


def schema(props, required):
    return {"type": "object", "properties": props, "required": required, "additionalProperties": False}


TOOLS = [
    ("mo_ta_file",
     "Mô tả gọn file BD SEO: khách, site, mỗi sheet (mục đích, vùng tổng hợp TH_..., tùy chọn và giá trị hiện tại, lưu ý dữ liệu, "
     "ô nhận xét đã có/chưa có, sheet dữ liệu thô + số dòng + cột). LUÔN gọi đầu tiên. Không trả dữ liệu thô.",
     schema({"path": PATH_PROP}, ["path"]), t_mo_ta_file),
    ("doc_tong_hop",
     "Đọc giá trị đã tính của một vùng tổng hợp (name = TH_... trong manifest, hoặc range dạng 'Sheet'!A1:J20). Trả bảng TSV có hàng "
     "chữ cái cột và cột đầu là số dòng để dẫn nguồn ô. Tối đa 60 dòng x 15 cột.",
     schema({"path": PATH_PROP, "name": {"type": "string", "description": "Tên vùng TH_... hoặc range 'Sheet'!A1:J20"}},
            ["path", "name"]), t_doc_tong_hop),
    ("hoi",
     "Chạy MỘT hàm Excel/Google Sheets (qua GridCraft) trên file và trả kết quả đã tính. Dùng COUNTIFS/SUMIFS/AVERAGEIFS/FILTER/SORT/TAKE/UNIQUE/"
     "CHOOSECOLS/SUMPRODUCT/GROUPBY trên data_sheet (tham chiếu kiểu 'RD data'!A:A). KHÔNG dùng để kéo cả bảng thô: mảng >50 dòng hoặc "
     ">4000 ký tự sẽ bị cắt, hãy dùng TAKE/FILTER chặt. Lỗi hàm được trả nguyên văn để sửa. Mỗi lần chạy khoảng 2 giây trên file 13k dòng.",
     schema({"path": PATH_PROP, "formula": {"type": "string", "description": "Công thức, có thể bắt đầu bằng '='"},
             "sheet": {"type": "string", "description": "Sheet mặc định cho ô không ghi tên sheet (tùy chọn)"}},
            ["path", "formula"]), t_hoi),
    ("xem_mau",
     "Xem tối đa 20 dòng mẫu từ dữ liệu thô (sheet = tên sheet kết quả hoặc data_sheet; file theo mẫu: thêm site hoặc range "
     "để chọn khối dữ liệu của site), có lọc/sắp xếp, kèm tổng số dòng thỏa điều kiện. "
     "Chỉ để xem ví dụ minh họa; muốn thống kê thì dùng hoi.",
     schema({"path": PATH_PROP, "sheet": {"type": "string", "description": "Tên sheet dữ liệu thô (vd 'RD data') hoặc sheet kết quả"},
             "dieu_kien": {"type": "array", "description": "Điều kiện AND trên tên cột",
                           "items": schema({"cot": {"type": "string"},
                                            "op": {"type": "string", "enum": ["=", "!=", ">", ">=", "<", "<=", "chua"]},
                                            "gt": {"type": ["string", "number"]}}, ["cot", "gt"])},
             "site": {"type": "string", "description": "File theo mẫu: chọn khối dữ liệu của site (xem 'Khối dữ liệu' trong mo_ta_file)"},
             "range": {"type": "string", "description": "File theo mẫu: vùng dữ liệu (không gồm tiêu đề), vd '1.9 Backlink'!M3:U5906"},
             "cot": {"type": "array", "items": {"type": "string"}, "description": "Các cột cần hiện (mặc định tất cả)"},
             "sap_xep": {"type": "string", "description": "'Tên cột:giam' hoặc 'Tên cột:tang'"},
             "n": {"type": "integer", "minimum": 1, "maximum": 20, "default": 10}},
            ["path", "sheet"]), t_xem_mau),
    ("ghi_nhan_dinh",
     "Ghi nhận xét vào MỘT ô nhận xét/claude_cells của manifest, kèm nguồn. Server tự tính lại từng nguồn; nguồn lỗi hoặc rỗng thì từ chối. "
     "Lưu vào <file>.notes.json (chưa sửa xlsx; gọi chot_file sau cùng). Trả values đã tính để đối chiếu câu chữ.",
     schema({"path": PATH_PROP, "cell": {"type": "string", "description": "Ô nhận xét theo manifest, vd '1.9 Backlink'!L70"},
             "text": {"type": "string", "description": "Nhận xét tiếng Việt, mỗi ý có số"},
             "sources": {"type": "array", "minItems": 1, "items": {"type": "string"},
                         "description": "Địa chỉ ô/range nhỏ ('Sheet'!G13) hoặc công thức bắt đầu bằng '='"}},
            ["path", "cell", "text", "sources"]), t_ghi_nhan_dinh),
    ("chot_file",
     "Chèn các nhận xét đã ghi vào file xlsx (xu_ly_data.bd_seo rebuild) và liệt kê ô nhận xét còn trống. Gọi sau khi ghi xong.",
     schema({"path": PATH_PROP}, ["path"]), t_chot_file),
    ("huong_dan", "Trả quy trình làm việc chuẩn với bộ tool seongon-bd (đọc trước khi bắt đầu).",
     schema({}, []), t_huong_dan),
]
TOOL_MAP = dict((t[0], t) for t in TOOLS)


def call_tool(name, args):
    if name not in TOOL_MAP:
        return None
    try:
        text = TOOL_MAP[name][3](args or {})
        return {"content": [{"type": "text", "text": text}], "isError": False}
    except ToolError as e:
        return {"content": [{"type": "text", "text": "LỖI: %s" % e}], "isError": True}
    except Exception as e:  # noqa
        log("exception", repr(e))
        return {"content": [{"type": "text", "text": "LỖI nội bộ: %r" % e}], "isError": True}


# ---------------------------------------------------------------- JSON-RPC
def send(obj):
    sys.stdout.write(json.dumps(obj, ensure_ascii=False) + "\n")
    sys.stdout.flush()


def err(id_, code, msg):
    send({"jsonrpc": "2.0", "id": id_, "error": {"code": code, "message": msg}})


def handle(msg):
    if not isinstance(msg, dict) or msg.get("jsonrpc") != "2.0" or "method" not in msg:
        if isinstance(msg, dict) and "id" in msg and "method" not in msg:
            return  # response từ client, bỏ qua
        err(msg.get("id") if isinstance(msg, dict) else None, -32600, "Invalid Request")
        return
    method, id_ = msg["method"], msg.get("id")
    params = msg.get("params") or {}
    is_note = "id" not in msg
    if method == "initialize":
        want = params.get("protocolVersion")
        send({"jsonrpc": "2.0", "id": id_, "result": {
            "protocolVersion": want if want in SUPPORTED else SUPPORTED[0],
            "capabilities": {"tools": {"listChanged": False}},
            "serverInfo": {"name": SERVER_NAME, "version": SERVER_VERSION}}})
    elif method == "ping":
        send({"jsonrpc": "2.0", "id": id_, "result": {}})
    elif method == "tools/list":
        send({"jsonrpc": "2.0", "id": id_, "result": {"tools": [
            {"name": t[0], "description": t[1], "inputSchema": t[2]} for t in TOOLS]}})
    elif method == "tools/call":
        if not isinstance(params, dict) or not isinstance(params.get("name"), str) or \
                not isinstance(params.get("arguments", {}), dict):
            err(id_, -32602, "Invalid params: cần name (string) và arguments (object)")
            return
        res = call_tool(params["name"], params.get("arguments", {}))
        if res is None:
            err(id_, -32602, "Unknown tool: %s" % params["name"])
            return
        send({"jsonrpc": "2.0", "id": id_, "result": res})
    elif is_note:
        return  # notifications/initialized, cancelled...
    else:
        err(id_, -32601, "Method not found: %s" % method)


def main():
    log("khởi động; gridcraft=%s" % GRIDCRAFT)
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            msg = json.loads(line)
        except ValueError:
            err(None, -32700, "Parse error")
            continue
        try:
            handle(msg)
        except Exception as e:  # noqa
            log("lỗi xử lý", repr(e))
            err(msg.get("id") if isinstance(msg, dict) else None, -32603, "Internal error: %r" % e)


if __name__ == "__main__":
    main()
