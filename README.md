# 🐸 Mini Frog — Tool quét Web URL

Bản mô phỏng Screaming Frog chạy local: nhập website → crawl → xem báo cáo SEO on-page → xuất Excel/CSV. **Không cần API key, không tốn phí, dữ liệu nằm hoàn toàn trên máy bạn.**

![Python](https://img.shields.io/badge/Python-3.9%2B-004aef) ![License](https://img.shields.io/badge/Free-100%25-0def9b)

---

## 📥 Bước 0 — Tải tool về máy (chung cho Mac & Windows)

**Cách 1 (không cần biết Git):** bấm nút **Code** (màu xanh) ở đầu trang này → **Download ZIP** → giải nén ra Desktop hoặc Documents.

**Cách 2 (có Git):**
```bash
git clone https://github.com/dat342/mini-frog-crawler.git
```

---

## 🍎 Cài đặt trên macOS

### Bước 1 — Mở tool lần đầu

1. Mở thư mục vừa giải nén trong **Finder**
2. **Chuột phải** (hoặc giữ Control + click) vào file **`Chạy tool.command`** → chọn **Open**
3. macOS hiện cảnh báo "app từ nhà phát triển không xác định" → bấm **Open** lần nữa
   *(chỉ cần làm 1 lần duy nhất, các lần sau nhấp đúp bình thường)*

### Bước 2 — Chờ tự cài đặt (chỉ lần đầu)

- Cửa sổ Terminal mở ra, hiện `🔧 Lần đầu chạy: đang cài đặt môi trường (1-2 phút, cần mạng)...`
- **Nếu máy chưa có Python**: macOS tự bật hộp thoại đề nghị cài *Command Line Tools* → bấm **Install** → chờ cài xong (5-10 phút) → nhấp đúp file `.command` chạy lại

### Bước 3 — Dùng

- Trình duyệt tự mở **http://localhost:8765** → nhập website → bấm **▶ Bắt đầu crawl**
- Tắt tool: đóng cửa sổ Terminal

### Lỗi thường gặp trên Mac

| Hiện tượng | Cách xử lý |
|---|---|
| Nhấp đúp không mở được, báo "unidentified developer" | Chuột phải → Open → Open (bước 1 ở trên) |
| Báo `Không tìm thấy Python 3` | Bấm Install ở hộp thoại Command Line Tools, hoặc tải Python tại python.org |
| Trình duyệt báo "can't connect" | Chờ 5 giây rồi tải lại trang (server đang khởi động) |
| Cổng 8765 bị chiếm | Mở Terminal chạy: `lsof -ti:8765 \| xargs kill` rồi mở lại tool |

---

## 🪟 Cài đặt trên Windows

### Bước 1 — Cài Python (chỉ 1 lần, bỏ qua nếu đã có)

1. Nhấp đúp file **`Chay tool - Windows.bat`** — nếu máy chưa có Python, tool tự mở trang tải
   (hoặc tự vào https://www.python.org/downloads/ → bấm **Download Python 3.x**)
2. Chạy file cài đặt Python, ở màn hình ĐẦU TIÊN:
   - ⚠️ **QUAN TRỌNG: tick vào ô "Add python.exe to PATH"** ở cuối màn hình (quên bước này tool sẽ không chạy)
   - Bấm **Install Now**
3. Cài xong bấm Close

### Bước 2 — Chạy tool

1. Nhấp đúp **`Chay tool - Windows.bat`**
2. Nếu Windows SmartScreen hiện cảnh báo xanh "Windows protected your PC" → bấm **More info** → **Run anyway** *(chỉ lần đầu)*
3. Lần đầu chạy: cửa sổ đen hiện `Lan dau chay: dang cai dat moi truong (1-2 phut, can mang)...` → chờ
4. Trình duyệt tự mở **http://localhost:8765** → dùng thôi!

### Bước 3 — Tắt tool

Đóng cửa sổ đen (Command Prompt) là tool tắt.

### Lỗi thường gặp trên Windows

| Hiện tượng | Cách xử lý |
|---|---|
| Báo `'python' is not recognized` | Quên tick "Add python.exe to PATH" → gỡ Python, cài lại, nhớ tick |
| Cài Python rồi mà vẫn báo chưa có | Đóng cửa sổ đen, nhấp đúp file `.bat` lại (PATH chỉ nhận sau khi mở lại) |
| SmartScreen chặn | More info → Run anyway |
| Cửa sổ đen tắt ngay lập tức | Chuột phải file `.bat` → Edit xem có lỗi, hoặc mở Command Prompt tại thư mục và gõ tên file để đọc lỗi |
| Trình duyệt báo "can't connect" | Chờ 5 giây rồi F5 |

---

## 🚀 Cách dùng

### 4 chế độ

| Chế độ | Dùng khi nào |
|---|---|
| 🌐 **Tự động từ trang chủ** | Nhập 1 domain, tool tự crawl theo ưu tiên: Trang chủ → Danh mục/Dịch vụ → Sản phẩm → Blog (mặc định 1.000 URL) |
| 📂 **Theo thư mục mẹ** | Dán URL thư mục (vd `site.com/may-giat/`), tool chỉ cào URL con bên trong |
| 📋 **Chỉ đúng URL tôi dán** | Dán danh sách URL (copy từ Excel/GSC), tool quét đúng các URL đó, không lần theo link — dán được nhiều domain |
| 🧭 **Tái tạo URL & phân loại** | Chức năng riêng cho URL phẳng (không có category trong URL). Đọc breadcrumb/schema/WordPress API để suy ra category, tái tạo thành `domain/%category%/%post%` và phân loại trang. Không bóc SEO — chỉ 2 việc này. Nguồn ưu tiên: Schema breadcrumb → Breadcrumb HTML → article:section → WordPress REST API |

### Dữ liệu thu được trên mỗi URL

Status code + chuỗi redirect · Title + độ dài · Meta description + độ dài · Số H1/H2 + nội dung H1 · Canonical + kiểm tra self-canonical · Meta robots/noindex · **Ngày đăng + ngày cập nhật** (từ meta/schema/sitemap) · Thumbnail og:image · Toàn bộ ảnh + alt (đánh dấu thiếu alt) · Số từ (phát hiện thin content) · **Inlinks + anchor text** · Phân tầng thư mục (Tầng 1/2/3) + tự phân loại danh mục/sản phẩm theo cây URL

### Phát hiện vấn đề SEO

Link gãy 404 · Lỗi 5xx · Redirect · Thiếu/trùng title & meta · Thiếu/nhiều H1 · Canonical sai · Noindex · Title/meta quá dài · Ảnh thiếu alt · Thin content (<150 từ) · Trang không có inlink · **Orphan pages** (có trong sitemap.xml nhưng không có link nội bộ trỏ tới)

### Xuất báo cáo

Bấm **⬇ Excel** (7 sheet: Tổng quan / URLs / Hình ảnh / Trùng lặp / Inlinks / Orphan / Vấn đề) hoặc **⬇ CSV** — file tải về thư mục Downloads.

> ⚠️ Dữ liệu crawl chỉ nằm trong bộ nhớ khi tool đang chạy — **xuất Excel trước khi tắt tool**.

### Mẹo

- Site đặt URL không theo chuẩn VN → sửa pattern trong mục **⚙ Pattern phân loại trang**
- Loại bỏ URL rác (giỏ hàng, filter...) → điền mục **Loại trừ URL** (regex hoặc chuỗi con, vd `/tag/`, `\?sort=`)
- Muốn orphan pages chính xác → đặt Giới hạn URL lớn hơn tổng số trang của site
- Site chậm/hosting yếu → giảm **Luồng song song** xuống 2-3

---

## ⚠️ Giới hạn

- Crawl HTML thuần — site render hoàn toàn bằng JavaScript (React/Vue thuần) sẽ thiếu dữ liệu (site VN đa số không bị)
- Ngày đăng/cập nhật chỉ lấy được khi site có khai báo (meta article, schema, sitemap lastmod)
- Cần Python 3.9 trở lên

## 🗂 Cấu trúc code

```
crawler.py          # Lõi crawl: asyncio + httpx + BeautifulSoup
app.py              # Backend FastAPI + xuất Excel/CSV
static/index.html   # Giao diện web (HTML/CSS/JS thuần)
requirements.txt    # Thư viện Python
```
