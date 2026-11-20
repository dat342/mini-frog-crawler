# Giao kèo xu_ly_data (BD SEO) — tool ↔ MCP ↔ skill

Mục tiêu: người dùng yêu cầu "làm file BD SEO cho <khách>, các sheet X, Y" ở bất kỳ phiên Claude nào → nhận 1 file .xlsx
đủ số liệu (tính bằng hàm Excel/Google Sheets) + nhận xét do Claude viết, mỗi nhận xét dẫn nguồn ô/hàm.

Nguyên tắc: **data được tool tính trước, Claude chỉ đọc bảng tổng hợp và hỏi bằng hàm**, không đọc dòng thô.

Python 3.9 (`.venv/`): KHÔNG dùng `X | None`, `match`, `list[str]` trong annotation chạy lúc runtime.
Công thức phải chạy được trên cả Excel và Google Sheets (SUMIFS/COUNTIFS/AVERAGEIFS/IFERROR/DATE…).
Tránh dấu `~`, `*`, `?` trong nhãn dùng làm điều kiện SUMIFS/COUNTIFS (là ký tự wildcard).

## 1. API module

Mỗi module là một file trong `xu_ly_data/`, có:

```python
KEY = "backlink_rd"                 # mã module
SHEET = "1.9 Backlink"              # tên sheet kết quả chính (theo mẫu BD SEO)
INPUT_KIND = "ahrefs_refdomains_csv" # loại đầu vào, xem mục 4
DESC = "Đánh giá backlink theo Referring domains"

def build(wb, inputs, opts, ctx):
    """wb: openpyxl Workbook dùng chung (đã có sẵn, có thể đã có sheet khác).
    inputs: list đường dẫn file (hoặc list domain với module gọi API).
    opts: dict tùy chọn của module (vd {"nam_dau": 2020}).
    ctx: {"out_path": ..., "cache_dir": ..., "sites": [...]}.
    Tạo sheet kết quả + sheet dữ liệu (nếu cần) trong wb. Trả về common.entry(...) (mục 3)."""
```

- Không tự tạo Workbook, không tự save. Không đọc API key trực tiếp — dùng `ctx` hoặc hàm doc_key của module traffic.
- Mỗi sheet kết quả phải có: tiêu đề (common.title), ít nhất 1 vùng tổng hợp đăng ký bằng `common.summary`
  (tạo defined name `TH_...`), ít nhất 1 vùng nhận xét `common.note_area` (mặc định nhãn "NHẬN XÉT"),
  các ô tùy chọn qua `common.option`.
- Vùng tổng hợp ≤ 60 dòng × 15 cột (để MCP đọc gọn).
- `caveats`: các lưu ý dữ liệu tool tự phát hiện được (vd "86% refdomain là spam, dồn vào năm 2026",
  "12% traffic chưa phân loại topic"). Đây là thứ giúp Claude không kết luận sai — tính bằng code, ghi bằng câu ngắn có số.

## 2. Orchestrator `bd_seo.py`

```
.venv/bin/python -m xu_ly_data.bd_seo build --recipe recipe.json --out "<Khách> - BD SEO.xlsx"
.venv/bin/python -m xu_ly_data.bd_seo rebuild "<Khách> - BD SEO.xlsx"   # dựng lại từ recipe + cache + notes
```

recipe.json:
```json
{"client": "Khach A", "sites": ["khach-a.vn", "a.com", "b.com"],
 "modules": [
   {"key": "traffic_ahrefs", "inputs": ["khach-a.vn", "a.com", "b.com"], "opts": {"thang": 36}},
   {"key": "backlink_rd", "inputs": ["/abs/a-refdomains.csv", "..."], "opts": {}},
   {"key": "content_topic", "inputs": ["/abs/a-top-pages.csv", "..."], "opts": {"rules": "/abs/rules.json"}},
   {"key": "onpage", "inputs": ["/abs/minifrog-a.xlsx", "..."], "opts": {}}
 ]}
```
Đầu ra cạnh file xlsx:
- `<file>.manifest.json` (mục 3) — đồng thời ghi vào sheet ẩn `_manifest` (ô A1, JSON).
- `<file>.recipe.json` (bản sao recipe đã dùng), `<file>.notes.json` (nhận xét, mục 5), cache API trong `<file>.cache/`.
Thứ tự sheet theo mẫu BD: 1.5 → 1.6 → 1.7 → 1.9; sheet dữ liệu thô xếp cuối.

## 3. Manifest (`<file>.manifest.json`)

```json
{"version": 1, "file": "/abs/....xlsx", "client": "...", "created": "2026-11-20T12:00:00",
 "sites": ["..."],
 "sheets": [ {
   "sheet": "1.9 Backlink", "module": "backlink_rd", "purpose": "…1 câu…",
   "inputs": ["..."], "data_sheet": "RD data", "rows": 13709,
   "columns": [{"name": "DR", "desc": "Domain Rating 0–100"}, ...],          // cột của data_sheet
   "summaries": [{"name": "TH_BL_seongon_com", "range": "'1.9 Backlink'!$A$8:$J$15", "desc": "..."}],
   "options": [{"cell": "'1.9 Backlink'!B3", "label": "Chỉ số", "values": [...], "default": "..."}],
   "caveats": ["..."],
   "notes": [{"cell": "'1.9 Backlink'!L60", "label": "NHẬN XÉT"}]
 } ]}
```

## 4. Loại đầu vào

| INPUT_KIND | Nguồn | Cột bắt buộc |
|---|---|---|
| ahrefs_refdomains_csv | Ahrefs › Referring domains › Export CSV UTF-8 | Domain, DR, First seen (có thì dùng: Is spam, Links to target, Lost) |
| ahrefs_top_pages_csv | Ahrefs › Top pages › Export CSV UTF-8 | URL, Traffic (có thì dùng: Page type, Top keyword, Top keyword: Volume, Keywords, Traffic value) |
| domains_api | Ahrefs API v3 (key: env AHREFS_API_KEY hoặc ~/.config/seongon/ahrefs.key) | — |
| minifrog_xlsx | Mini Frog › Xuất Excel (sheet URL) hoặc /api/results JSON | xem URL_COLUMNS trong app.py |

Tên site lấy từ tên file Ahrefs (`<domain>-refdomains-...csv`, `<domain>-top-pages-...csv`) hoặc opts["sites"].

## 5. Nhận xét (`<file>.notes.json`) — do MCP ghi, orchestrator chèn khi rebuild

```json
{"'1.9 Backlink'!L60": {"text": "…nhận xét…", "sources": ["'1.9 Backlink'!G13", "=COUNTIFS(...)"],
                        "values": ["176", "29"], "updated": "2026-11-20T12:30:00"}}
```
Rebuild: ghi `text` vào ô; ghi chú (comment) của ô liệt kê `sources` = `values`. Nhận xét không có nguồn bị từ chối ở MCP.

## 6. Bổ sung sau khi ráp (20/11/2026)
- File phụ đặt tên theo đường dẫn đầy đủ của xlsx: `<file>.xlsx.manifest.json`, `.xlsx.recipe.json`, `.xlsx.notes.json`, thư mục cache `<file>.xlsx.cache/`.
- Manifest có thêm `warnings` (module lỗi bị bỏ qua) và `notes_applied`. Sheet `_manifest` có thể chia JSON sang A2, A3… khi > 30.000 ký tự.
- `ctx` có `sites`. Module onpage nhận JSON dạng list records hoặc `{"records": [...], "total": n}` (UI Mini Frog lưu dạng này).
- Onpage: entry có `claude_cells`; điểm nhóm = Đạt / (Đạt + Chưa đạt), ô "Chưa kiểm" không tính.
- MCP `seongon-bd` đăng ký ở user scope (`claude mcp add --scope user`), dùng được ở mọi phiên.

## 7. Đổ vào mẫu (20/11/2026) — `template_bd.py`
Người dùng chỉ nhận file trình bày ĐÚNG mẫu Excel của họ (xuất từ Google Sheets, vd `5-inbox/BD SEO _ c06_2026).xlsx`).
- recipe có `"template": "/abs/mẫu.xlsx"` → `bd_seo build/rebuild` dựng bản tạm bằng các module như cũ
  (`<file>.cache/_ban-tam.xlsx`, dùng chung cache traffic) rồi gọi `template_bd.do_vao_mau(built, manifest, template, out,
  notes, opts)` → file cuối + manifest mới. Bản tạm bị xoá sau khi xong.
- File mẫu CHỈ ĐỌC. openpyxl chỉ để ĐỌC (mẫu + bản tạm). Mọi thay đổi trên mẫu đi qua GridCraft
  (`gridcraft-cli run --in <mẫu> --script steps.jsonl --out <file>`): giữ biểu đồ, định dạng của mẫu. KHÔNG lưu file mẫu bằng openpyxl
  (mất biểu đồ/ảnh). Đặt `BD_KEEP_SCRIPT=1` để giữ lại file jsonl khi gỡ lỗi.
- Giữ 6 sheet: Mục lục, 1. Định hướng, 1.5, 1.6, 1.7, 1.9 (sheet của module không có trong recipe cũng bị xoá).
  Mục lục chỉ giữ hạng mục nhóm "Định hướng" ứng với sheet còn lại (kèm link sang sheet).
- "1. Định hướng": xoá mục II.1 (dòng + biểu đồ), xoá biểu đồ trỏ sheet đã xoá và biểu đồ kênh traffic; trỏ lại radar (1.5 A1:D6),
  line organic (1.6 J:M), cột topic (1.7), offpage (1.9: tổng theo năm × site và chất lượng năm cuối × site).
  4 ô nhận xét (khung "Dữ liệu nghiên cứu / Đánh giá / Định hướng") là `notes` của sheet này, label Onpage / Organic traffic /
  Content / Offpage. Không còn ô "NHẬN XÉT" trong sheet dữ liệu.
- Bố cục sheet dữ liệu: bảng thống kê bên trái (cột A…), dữ liệu thô từng site xếp cạnh nhau bên phải trong cùng sheet
  (hàng 1 tên site nền #004AEF, hàng 2 tiêu đề cột nền #12D6F7, cách 1 cột). Không còn sheet 'RD data', 'TP data', 'Traffic data'.
  - 1.5: bảng điểm A1 (Nhóm × site, dòng Tổng), bảng tiêu chí từ dòng 9: Nhóm | Các vấn đề (B:C) | Vì sao | Yêu cầu |
    Đánh giá site… | Bằng chứng site… | Ngưỡng tự chấm | Người chấm. Ô Claude chấm = `claude_cells`; khi có notes thì `text`
    vào ô đánh giá, `values` vào cột Bằng chứng (không dùng comment).
  - 1.6: khối kênh A:H chỉ giữ tiêu đề (chưa có Similarweb); bảng Organic Search J:… (tháng dạng chữ mm/yyyy) + Tổng hợp.
  - 1.7: Số lượng index (nhập tay) A1, bảng Topic × site A6 (COUNTIFS), biểu đồ, Topic traffic, Loại trang (số trang, traffic, %);
    khối site từ cột H: URL | Traffic | Top keyword | Loại Trang | Topic | Top KW volume.
  - 1.9: ô chọn B1:B3 (Chỉ số/Phạm vi/Hiện Spam, ô phụ D1), bảng tổng theo năm A5, chất lượng năm cuối, 2 biểu đồ mẫu,
    rồi bảng lũy kế năm × nhóm + biểu đồ cột chồng từng site; khối site từ cột M:
    Domain | Spam | DR | Links to target | First seen | Lost | Trạng thái | Nhóm | Năm (3 cột cuối là hàm).
- Manifest file theo mẫu: thêm `layout: "mau_bd"`, `template`, `deleted_sheets`; mỗi sheet có thể có
  `data_blocks: [{"site","sheet","range" (không gồm tiêu đề),"header_row","rows","columns"}]` thay cho `data_sheet`.
  Vùng tổng hợp vẫn là defined name `TH_…` (tạo bằng `formulas.defineName` của GridCraft).
- MCP: `mo_ta_file` in "Khối dữ liệu" (địa chỉ từng cột); `xem_mau` nhận `site` hoặc `range` để chọn khối; `hoi` dùng địa chỉ khối,
  vd `=COUNTIFS('1.9 Backlink'!$T$3:$T$5906,"70-100",'1.9 Backlink'!$S$3:$S$5906,"Còn")`.
- Nhận xét Định hướng: rebuild ghi `text` vào ô và tạo note (comment) "Nguồn: <ô> = <giá trị>" bằng `review.newNote`.
- Chuyển nhận xét onpage từ file bố cục cũ sang file theo mẫu:
  `.venv/bin/python -m xu_ly_data.template_bd chuyen-notes --tu "<cũ>.xlsx" --sang "<theo mẫu>.xlsx"` rồi `bd_seo rebuild`.
- Hạn chế: style biểu đồ chèn mới (cột chồng từng site ở 1.9) là style mặc định GridCraft; trục ngày của chart chưa đặt được
  định dạng số nên tháng ở 1.6 ghi dạng chữ; mẫu không ẩn gridlines nên giữ nguyên; tối đa 3 site có radar/màu cột đánh giá theo mẫu.
