#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Xuất cửa sổ Claude Code đang làm việc (session mới nhất) ra file .txt dễ đọc.

Cách dùng:
    python3 "export-hoi-thoai.py"              # tự tìm session mới nhất, lưu ra Desktop
    python3 "export-hoi-thoai.py" ten-file.txt # đặt tên file tùy ý
"""
import glob
import json
import os
import sys
from datetime import datetime

PROJECTS_DIR = os.path.expanduser("~/.claude/projects")


def find_latest_session():
    """Tìm file .jsonl được sửa gần nhất trong mọi project = cửa sổ đang làm việc."""
    files = glob.glob(os.path.join(PROJECTS_DIR, "*", "*.jsonl"))
    if not files:
        sys.exit("Không tìm thấy bản ghi hội thoại nào trong ~/.claude/projects")
    return max(files, key=os.path.getmtime)


def text_of(content):
    txts, tools, is_tr = [], [], False
    if isinstance(content, str):
        return content.strip(), [], False
    if isinstance(content, list):
        for b in content:
            if not isinstance(b, dict):
                continue
            t = b.get("type")
            if t == "text":
                txts.append(b.get("text", ""))
            elif t == "tool_use":
                name = b.get("name", "tool")
                inp = b.get("input", {}) if isinstance(b.get("input"), dict) else {}
                desc = inp.get("description") or inp.get("command") or inp.get("file_path") or ""
                if isinstance(desc, str) and len(desc) > 90:
                    desc = desc[:90] + "..."
                tools.append(f"{name}: {desc}".strip().rstrip(":").strip())
            elif t == "tool_result":
                is_tr = True
    return "\n".join(txts).strip(), tools, is_tr


def is_noise(s):
    if not s:
        return True
    n = s.lstrip()
    return (n.startswith("<system-reminder>") or n.startswith("<command-")
            or n.startswith("<local-command") or n.startswith("Caveat:")
            or n.startswith("<bash-") or n.startswith("[Request interrupted"))


def main():
    src = find_latest_session()
    default_name = "hoi-thoai-" + datetime.now().strftime("%Y%m%d-%H%M") + ".txt"
    out_arg = sys.argv[1] if len(sys.argv) > 1 else default_name
    out = out_arg if os.path.isabs(out_arg) else os.path.expanduser(f"~/Desktop/{out_arg}")

    blocks = []
    for line in open(src, encoding="utf-8"):
        try:
            o = json.loads(line)
        except Exception:
            continue
        if o.get("type") not in ("user", "assistant"):
            continue
        msg = o.get("message") or {}
        role = msg.get("role")
        if role not in ("user", "assistant"):
            continue
        text, tools, is_tr = text_of(msg.get("content"))
        if role == "user":
            if is_tr or is_noise(text) or not text:
                continue
            blocks.append(("user", text, []))
        else:
            if not text and not tools:
                continue
            blocks.append(("assistant", text, tools))

    sep = "=" * 70
    lines = ["LỊCH SỬ TRÒ CHUYỆN — Claude Code",
             f"Nguồn: {os.path.basename(src)}",
             "(Thao tác công cụ được tóm tắt 1 dòng.)", sep, ""]
    nu = na = 0
    for role, text, tools in blocks:
        if role == "user":
            nu += 1
            lines.append("BẠN:")
            lines.append(text)
        else:
            na += 1
            lines.append("CLAUDE:")
            for t in tools:
                lines.append("   [đã dùng] " + t)
            if text:
                lines.append(text)
        lines += ["", "-" * 70, ""]

    with open(out, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    print(f"✓ Đã xuất: {out}")
    print(f"  {nu} lượt bạn, {na} lượt Claude — từ {os.path.basename(src)}")


if __name__ == "__main__":
    main()
