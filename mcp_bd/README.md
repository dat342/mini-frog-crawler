# seongon-bd (MCP server, stdio)

Giúp Claude phân tích file BD SEO .xlsx mà không đọc dữ liệu thô: đọc manifest + bảng tổng hợp, hỏi sâu bằng hàm (GridCraft CLI), ghi nhận xét có dẫn nguồn.
Runtime: Python 3.9 của venv, chỉ thư viện chuẩn + openpyxl, tự cài JSON-RPC 2.0 (không dùng SDK `mcp`).

Tools: `huong_dan`, `mo_ta_file`, `doc_tong_hop`, `hoi`, `xem_mau`, `ghi_nhan_dinh`, `chot_file`.
Giới hạn: chỉ nhận file .xlsx tồn tại, đường dẫn tuyệt đối, dưới $SEONGON_ROOT (mặc định ~/Downloads/Claude/SEONGON). Mỗi lần gọi CLI timeout 60s.

## Đăng ký

```
claude mcp add seongon-bd -- <THU_MUC_TOOL>/.venv/bin/python <THU_MUC_TOOL>/mcp_bd/server.py
```

Đổi GridCraft: thêm `-e GRIDCRAFT_CLI=/đường/dẫn/gridcraft-cli` (mặc định ~/.local/gridcraft/gridcraft-cli-0.19.1-macos-universal/gridcraft-cli).

## .mcp.json mẫu

```json
{
  "mcpServers": {
    "seongon-bd": {
      "command": "<THU_MUC_TOOL>/.venv/bin/python",
      "args": ["<THU_MUC_TOOL>/mcp_bd/server.py"],
      "env": {}
    }
  }
}
```

## Test

```
4-tools/tool-quet-web-url/.venv/bin/python 4-tools/tool-quet-web-url/mau-test/mcp/test_mcp.py
```
Dùng fixture `mau-test/mcp/fixture.xlsx` (+ `.manifest.json`); test sẽ ghi `fixture.xlsx.notes.json`.
