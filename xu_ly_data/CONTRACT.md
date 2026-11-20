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

## 8. Mẫu gắn thẻ (thay template_bd viết cứng) — đặc tả chung cho engine, module và file mẫu

Mục tiêu: mỗi cách trình bày = 1 file mẫu .xlsx do người dùng thiết kế (Google Sheets/Excel) + thẻ gõ trong ô.
Engine `xu_ly_data/the_mau.py` đọc thẻ, đặt bảng/biểu đồ/dữ liệu/nhận xét đúng chỗ, giữ màu/font/ảnh/logo của mẫu.
Mẫu lưu ở `xu_ly_data/mau/` (vd `bd-seo.xlsx`, `dinh-huong-du-an.xlsx`) kèm `<ten>.mau.json` mô tả (tên, module cần, số site tối đa).
Ghi file bằng GridCraft CLI (openpyxl chỉ để ĐỌC mẫu — openpyxl save làm mất chart/ảnh).

### 8.1 Cú pháp thẻ — cả ô chỉ chứa 1 thẻ, trừ thẻ chữ (8.1.5) có thể nằm giữa câu
`{{loai:khoa|tham_so=gia_tri|...}}`

1. `{{bang:<khoa>}}` — đặt bảng tổng hợp, ô thẻ = góc trên trái. Style: dòng tiêu đề lấy style ô thẻ; các dòng thân lấy
   style ô ngay dưới ô thẻ (người thiết kế định dạng sẵn 2 ô đó). Tham số: `site=<i>` (bảng riêng của site thứ i),
   `top=<n>` (chỉ n dòng đầu), `cot=<danh sách tên cột>` (chọn cột).
2. `{{bieu_do:<khoa>|loai=bar|col|col_stacked|line|pie|doughnut|radar|...}}` — chèn biểu đồ, neo tại ô thẻ.
   Dữ liệu = bảng `<khoa>` (engine tự đặt bảng nguồn ở sheet ẩn `_bang` nếu mẫu không có thẻ bảng tương ứng).
   Tham số: `site=<i>` (chuỗi = cột của site i — cột có tiêu đề là domain site i), `site=*|buoc=<số cột>` (lặp 1 biểu đồ
   mỗi site, dịch phải `buoc` cột), `rong=<px>`, `cao=<px>`, `tieu_de=<chữ, cho phép {{site}}/{{ten}}>`, `dong=<a-b>`.
3. `{{du_lieu:<khoa>|kieu=canh_nhau|doc}}` — hiển thị dữ liệu thô: `canh_nhau` = mỗi site 1 khối cạnh nhau
   (hàng 1 tên site, hàng 2 tiêu đề cột, cách 1 cột trống); `doc` = 1 bảng dài có cột Website. Đây là bản HIỂN THỊ (giá trị).
   Công thức của mọi bảng luôn trỏ vào sheet dữ liệu ẩn chuẩn của module (vd `RD data`) — engine luôn chép sheet đó vào file
   ra (ẩn) với đúng tên, nên công thức module giữ nguyên.
4. `{{nhan_xet:<khoa>|khung=...}}` — ô Claude viết nhận xét (đăng ký vào manifest.notes, label = khoa). Engine ghi sẵn khung,
   mặc định "Dữ liệu nghiên cứu:\n\nĐánh giá:\n\nĐịnh hướng:".
5. Thẻ chữ, thay tại chỗ, nằm được giữa câu: `{{khach}}`, `{{nam}}`, `{{ngay}}` (dd/mm/yyyy), `{{site:i}}` (domain site i),
   `{{ten:i}}` (tên thương hiệu site i, recipe `names`, mặc định = domain), `{{so:<khoa_so>|site=i}}` (1 con số, vd tb12 traffic).
6. `{{lap_site|tu=i}}…` — KHÔNG làm ở bản đầu; lặp theo site dùng `site=*` của biểu đồ, `site=i` của bảng.
Thẻ không nhận ra → báo lỗi rõ (vị trí ô + thẻ), không bỏ qua im lặng. Sheet trong mẫu không có thẻ → giữ nguyên.
Recipe có `"sheets": [...]` → chỉ giữ các sheet đó (cộng sheet ẩn dữ liệu); sheet có thẻ mà module chưa chạy → xoá sheet
và báo.

### 8.2 Danh mục khóa (module phải đăng ký đúng khóa này trong manifest.summaries[].key / charts / data)
Mọi bảng: dòng 1 = tiêu đề cột, cột 1 = nhãn; với bảng theo site, tiêu đề cột = domain đúng như recipe `sites`.
| Module | bang | bieu_do mặc định (loai) | du_lieu | so |
|---|---|---|---|---|
| backlink_rd | `bl_theo_nam` (năm × site, tổng lũy kế), `bl_chat_luong` (nhóm DR × site, năm cuối), `bl_chat_luong_pct`, `bl_site` (site=i: năm × nhóm) | `bl_theo_nam` (col), `bl_chat_luong` (col), `bl_site` (col_stacked) | `bl_data` | `bl_tong`, `bl_dr70`, `bl_spam_pct` |
| traffic_ahrefs | `tr_thang` (tháng × site), `tr_tong_hop` (TB12, TB12 trước, tăng trưởng × site) | `tr_thang` (line) | `tr_data` | `tr_tb12`, `tr_tang_truong`, `tr_dinh` |
| content_topic | `ct_tong` (site × số trang, traffic), `ct_topic_sotrang`, `ct_topic_traffic`, `ct_loaitrang_sotrang`, `ct_loaitrang_traffic`, `ct_loaitrang_pct` | `ct_tong` (bar), `ct_loaitrang` (pie, site=i), `ct_topic` (pie, site=i) | `ct_data` | `ct_so_trang`, `ct_traffic` |
| onpage | `op_diem` (nhóm × site), `op_tieu_chi` | `op_diem` (radar, site=i) | `op_data` | `op_diem_tong` |
| volume_trend (mới) | `vol_thang` (tháng × tổng volume, 12 tháng), `vol_nhom` (nhóm × volume, %), `vol_chu_de` (chủ đề × volume, giảm dần), `vol_top_tu_khoa` (từ khóa, nhóm, chủ đề, volume TB) | `vol_thang` (col), `vol_nhom` (pie), `vol_chu_de` (bar) | `vol_data` | `vol_tong`, `vol_so_tu` |
| batch_overview (mới) | `bo_tong_quan` (site × DR, refdomain, organic traffic, organic keywords, traffic value) | `bo_traffic` (bar) | `bo_data` | `bo_traffic`, `bo_dr` |
Module mới phải theo API mục 1, có summaries với trường `key` như bảng trên; khóa `so:` khai báo trong entry `extra.so`
= {"khoa": {"site": "<domain>", "cell": "'Sheet'!B5"}} hoặc công thức.

### 8.3 Hiện thực (engine `the_mau.py`, 20/11/2026) — chi tiết bổ sung, không đổi 8.1/8.2
Lệnh:
```
.venv/bin/python -m xu_ly_data.the_mau quet <mau.xlsx | tên trong mau/>     # liệt kê thẻ + kiểm khóa/tham số (rc≠0 nếu lỗi)
.venv/bin/python -m xu_ly_data.the_mau dien --built tam.xlsx --manifest tam.xlsx.manifest.json --mau bd-seo --out ra.xlsx [--recipe r.json] [--notes n.json]
.venv/bin/python -m xu_ly_data.the_mau chuyen-notes --tu "<cũ>.xlsx" --sang "<mới>.xlsx"   # nhận xét theo label, ô onpage theo (vấn đề, site)
.venv/bin/python -m xu_ly_data.mau.tao_bd_seo        # tạo lại mau/bd-seo.xlsx từ file mẫu BD gốc
```
- bd_seo: recipe `"mau": "bd-seo"` (tên trong `xu_ly_data/mau/` hoặc đường dẫn) → build bản tạm `<file>.cache/_ban-tam.xlsx`
  rồi `dien_mau`. Recipe cũ có `"template"` được CHUYỂN sang mẫu `bd-seo` (cảnh báo trong manifest); `template_bd.py` là bản cũ.
  `rebuild` = build lại + chèn `notes.json`: ô nhan_xet ghi text + note "Nguồn: ô = giá trị" (GridCraft `review.newNote`);
  ô Claude chấm ghi text, `values` vào ô cặp (`claude_pairs`, vd cột Bằng chứng) nếu có, không thì note.
  Registry module: `MODULE_ORDER` gồm cả `volume_trend`, `batch_overview` (import động, thiếu/lỗi → cảnh báo, bỏ qua).
- Dạng đăng ký chuẩn của module (thống nhất với volume_trend/batch_overview):
  - `summaries[]`: `key` (khóa 8.2), `site` (bảng riêng 1 site, vd bl_site), `total_rows` (số dòng tổng cuối bảng; biểu đồ bỏ qua).
    Helper: `common.bang(wb, ws, key, ref, desc, site=None, total_rows=0, name=None)`.
  - `extra.so`: `{khoa: {"site", "cell"}}` hoặc `{khoa: [{"site", "cell"}, ...]}`; thay `cell` bằng `"formula"` (+ `"sheet"`,
    `"num_fmt"`) khi là công thức. `common.so_cell(ws, "B5", site)`.
  - `extra.charts`: `[{key, bang, loai, cot?: tên cột|list, chuoi?: "cot"|"dong", tieu_de?: (cho phép {{ten}}), mau?: [hex]}]`
    (`anchor` nếu có bị bỏ qua). Khóa bieu_do không đăng ký mà trùng khóa bảng → vẽ thẳng bảng đó.
  - `extra.du_lieu`: dict hoặc list `{key, sheet, site_col?: "Website", cot?: [[cột nguồn, tên hiển thị], ...]}`
    (mặc định mọi cột trừ cột site).
  - `extra.validations`: `[{"range": "'Sheet'!F10:H90", "values": [...]}]`; `options` (ô chọn) cũng thành dropdown ở vị trí mới.
  - `extra.claude_cells`, `claude_items`, `claude_pairs` (ô chấm → ô bằng chứng) — engine đổi sang toạ độ file ra.
- Công thức trong bảng: chỉ trỏ (a) sheet dữ liệu ẩn của module, (b) ô thuộc bảng đã đăng ký cùng sheet kết quả. Engine đổi địa chỉ
  (b) sang vị trí bảng trong file ra; bảng được tham chiếu mà mẫu không đặt (hoặc đặt thiếu dòng/cột do top/cot) → bản đầy đủ
  ở sheet ẩn `_bang`; ô lẻ ngoài bảng (vd ô tổng B4 của volume_trend) cũng được chép vào `_bang`. `_bang` không dùng thì xoá.
- Sheet dữ liệu ẩn: 'RD data' (Trạng thái/Nhóm/Năm là GIÁ TRỊ), 'Traffic data' (+ cột `Kỳ` = mm/yyyy chữ, khóa của tr_thang),
  'TP data', 'Onpage data' (bản ghi crawl), 'Volume data', 'Batch data'.
- Bổ sung danh mục: bang `bl_tuy_chon` (3 ô chọn Chỉ số/Phạm vi/Hiện Spam — công thức backlink trỏ vào đây), `op_da_cham`
  (tiến độ chấm); so `bo_keywords`, `bo_rd`, `bo_ai_responses` (batch_overview — chưa đăng ký thì engine báo lỗi rõ).
  `op_diem` chỉ gồm các nhóm + dòng Tổng. `op_tieu_chi` cột: Nhóm | Nhóm con | Vấn đề | Vì sao | Yêu cầu | Đánh giá×site |
  Bằng chứng×site | Ngưỡng tự chấm | Người chấm. bieu_do mặc định: `ct_loaitrang` = ct_loaitrang_sotrang, `ct_topic` = ct_topic_sotrang,
  `bl_chat_luong` chuỗi theo dòng (nhóm), `bl_site` chỉ 6 cột nhóm.
- Tham số bổ sung: bang `bo_tieu_de=1` (không ghi dòng tiêu đề, thân lấy style dòng thẻ); bieu_do `chuoi=dong|cot`, `cot=`, `mau=#..,#..`.
  `site=i` trên bảng chung (cột = domain) → chỉ cột site i (bảng) / chuỗi site i (biểu đồ). Biểu đồ tròn mặc định 1 chuỗi (cột số đầu).
- Style: dòng tiêu đề = ô thẻ, dòng thân = ô ngay dưới, xét TỪNG CỘT (ô cùng dòng bên phải nếu có định dạng, không thì ô gần nhất
  bên trái); định dạng số luôn lấy từ module; dòng tổng in đậm; ô thân xuống dòng → engine tính chiều cao dòng. Ô định dạng thừa bên
  phải (mẫu chuẩn bị nhiều cột site hơn) bị xoá định dạng. du_lieu canh_nhau: hàng tên site = ô thẻ, tiêu đề cột = ô dưới, thân = ô dưới nữa;
  độ rộng cột khối 1 chép sang các khối sau.
- Chỗ trống: bảng/du_lieu dài hơn khoảng trống bên dưới (tới ô có nội dung/thẻ/ảnh kế tiếp trong cùng cột) → chèn dòng (cảnh báo).
- Thẻ site=i > số site recipe → để trống ô. Sheet mà mọi thẻ dữ liệu thuộc module chưa chạy → xoá sheet + cảnh báo; sheet có một phần
  thiếu → để trống các thẻ đó + cảnh báo. Khóa `nhan_xet` là nhãn tự do (không kiểm danh mục).
- Manifest file ra: `layout: "mau_the"`, `mau`, `deleted_sheets`; mỗi module 1 entry (`sheet` = sheet hiện nhiều bảng nhất,
  `data_sheet` = sheet ẩn, `summaries` có `key` [+ `site`], bảng > 55 dòng chia thêm `TH_<key>_p1..`, `data_blocks`, `charts`, `so`
  đã đổi toạ độ, `claude_cells`…); mỗi sheet có thẻ nhan_xet thêm 1 entry `module: "mau"` với `notes` (label = khóa).
- Mẫu kèm `<ten>.mau.json`: `ten`, `mo_ta`, `modules` (hoặc `module_can`/`module_tuy_chon`), `so_site_toi_da`, `nhan_xet`.
