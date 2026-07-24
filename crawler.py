# -*- coding: utf-8 -*-
"""Lõi crawler: quét URL theo ưu tiên loại trang, bóc tách dữ liệu SEO on-page."""
import asyncio
import html as htmllib
import json
import re
import time
from urllib.parse import urljoin, urlparse, urlunparse, parse_qsl, urlencode
from urllib import robotparser

import httpx
from bs4 import BeautifulSoup

USER_AGENT = "Mozilla/5.0 (compatible; MiniFrog/1.0; local SEO crawler)"

PAGE_TYPE_LABELS = {
    "homepage": "Trang chủ",
    "category": "Danh mục / Dịch vụ",
    "product": "Sản phẩm",
    "blog": "Blog / Tin tức",
    "other": "Khác",
}

# Thứ tự ưu tiên crawl: trang chủ -> danh mục/dịch vụ -> sản phẩm -> blog -> khác
TYPE_PRIORITY = {"homepage": 0, "category": 1, "product": 2, "blog": 3, "other": 4}

DEFAULT_PATTERNS = {
    "category": [
        "/danh-muc", "/danhmuc", "/category", "/categories", "/product-category",
        "/dich-vu", "/dichvu", "/service", "/collections", "/c/", "/cat/",
        "/thuong-hieu", "/brand",
    ],
    "product": [
        "/san-pham", "/sanpham", "/product", "/sp/", "/p/", "/item",
    ],
    "blog": [
        "/blog", "/tin-tuc", "/tintuc", "/news", "/bai-viet", "/baiviet",
        "/kien-thuc", "/huong-dan", "/cam-nang", "/post", "/article", "/tu-van",
    ],
}

SKIP_EXTENSIONS = re.compile(
    r"\.(jpg|jpeg|png|gif|webp|svg|ico|bmp|css|js|mjs|json|xml|txt|pdf|zip|rar|7z|gz|"
    r"doc|docx|xls|xlsx|ppt|pptx|mp3|mp4|avi|mov|webm|wav|woff2?|ttf|eot|otf|apk|exe|dmg)"
    r"(\?|#|$)",
    re.I,
)

TRACKING_PARAMS = {"fbclid", "gclid", "msclkid", "zanpid", "zarsrc", "srsltid", "yclid"}

DATE_RE = re.compile(r"\d{4}-\d{2}-\d{2}")
VISIBLE_DATE_RE = re.compile(r"\b(\d{1,2}[/-]\d{1,2}[/-]20\d{2}|20\d{2}-\d{2}-\d{2})\b")

# Các meta name khai báo ngày mà site VN hay dùng:
# - pubdate/publishdate: VnExpress, nhiều báo điện tử
# - dc.* / dcterms.*: Dublin Core - site .gov.vn, thư viện, đại học
# - cxenseparse: hệ thống cá nhân hóa cXense của nhiều tòa soạn
# - sailthru.date: nền tảng email/analytics báo chí
META_PUB_NAMES = frozenset((
    "pubdate", "publishdate", "publish-date", "publish_date", "publication_date",
    "date", "dc.date", "dc.date.issued", "dcterms.date", "dcterms.created",
    "sailthru.date", "cxenseparse:recs:publishtime", "article:published_time",
    "og:article:published_time", "datepublished", "creation_date", "ptime",
))
META_MOD_NAMES = frozenset((
    "lastmod", "last-modified", "lastmodified", "revised", "dc.date.modified",
    "dcterms.modified", "article:modified_time", "datemodified", "updated_time",
))
# Nhãn tiếng Việt/Anh để phân biệt ngày đăng vs ngày cập nhật trong text hiển thị
MOD_LABEL_RE = re.compile(r"cập nhật|cap nhat|updated|sửa đổi|chỉnh sửa|last modified", re.I)


def clean_date(value):
    """Rút gọn về dạng YYYY-MM-DD nếu nhận diện được."""
    if not value:
        return None
    value = str(value).strip()
    m = DATE_RE.search(value)
    return m.group(0) if m else (value[:40] or None)


def find_jsonld_dates(obj, found):
    """Tìm đệ quy datePublished / dateModified trong JSON-LD."""
    if isinstance(obj, dict):
        for key in ("datePublished", "dateModified"):
            v = obj.get(key)
            if isinstance(v, str) and v and key not in found:
                found[key] = v
        for v in obj.values():
            find_jsonld_dates(v, found)
    elif isinstance(obj, list):
        for v in obj:
            find_jsonld_dates(v, found)


def strip_www(host: str) -> str:
    return host[4:] if host.startswith("www.") else host


def normalize_url(base: str, href: str):
    """Chuẩn hóa URL để khử trùng lặp. Trả về None nếu không crawl được."""
    if not href:
        return None
    href = href.strip()
    low = href.lower()
    if low.startswith(("mailto:", "tel:", "javascript:", "data:", "#", "sms:", "callto:")):
        return None
    absolute = urljoin(base, href)
    parsed = urlparse(absolute)
    if parsed.scheme not in ("http", "https"):
        return None
    if SKIP_EXTENSIONS.search(parsed.path or ""):
        return None
    host = (parsed.hostname or "").lower()
    if not host:
        return None
    # Bỏ tham số tracking, giữ tham số nội dung (vd: ?page=2)
    query_pairs = [
        (k, v) for k, v in parse_qsl(parsed.query, keep_blank_values=True)
        if k.lower() not in TRACKING_PARAMS and not k.lower().startswith("utm_")
    ]
    path = parsed.path or "/"
    if path != "/" and path.endswith("/"):
        path = path.rstrip("/")
    port = ""
    if parsed.port and not (
        (parsed.scheme == "http" and parsed.port == 80)
        or (parsed.scheme == "https" and parsed.port == 443)
    ):
        port = f":{parsed.port}"
    return urlunparse((parsed.scheme, host + port, path, "", urlencode(query_pairs), ""))


def classify(url: str, patterns: dict) -> str:
    """Phân loại trang theo pattern trên đường dẫn URL."""
    parsed = urlparse(url)
    path = (parsed.path or "/").lower()
    if path in ("", "/") and not parsed.query:
        return "homepage"
    full = path + ("?" + parsed.query.lower() if parsed.query else "")
    for page_type in ("category", "blog", "product"):
        for pat in patterns.get(page_type, []):
            pat = pat.strip().lower()
            if pat and pat in full:
                return page_type
    return "other"


def compute_duplicates(records: list) -> dict:
    """Gom nhóm title / meta description trùng lặp giữa các trang 200."""
    titles, descs = {}, {}
    for r in records:
        if r.get("status") != 200:
            continue
        t = (r.get("title") or "").strip()
        d = (r.get("meta_description") or "").strip()
        if t:
            titles.setdefault(t, []).append(r["url"])
        if d:
            descs.setdefault(d, []).append(r["url"])
    return {
        "titles": {k: v for k, v in titles.items() if len(v) > 1},
        "descriptions": {k: v for k, v in descs.items() if len(v) > 1},
    }


# ===== Tái tạo URL %category%/%post% từ breadcrumb (chế độ reconstruct) =====

GENERIC_CRUMBS = {
    "trang chủ", "trang chu", "home", "homepage", "trang chính", "trang chinh", "#", "",
}
_ACCENT_PAIRS = [
    ("àáảãạăằắẳẵặâầấẩẫậ", "a"), ("èéẻẽẹêềếểễệ", "e"), ("ìíỉĩị", "i"),
    ("òóỏõọôồốổỗộơờớởỡợ", "o"), ("ùúủũụưừứửữự", "u"), ("ỳýỷỹỵ", "y"), ("đ", "d"),
]


def slugify_vn(text: str) -> str:
    """Chuyển tên category tiếng Việt thành slug: 'Dinh dưỡng mẹ' -> 'dinh-duong-me'."""
    text = htmllib.unescape(text or "").strip().lower().replace("#", "")
    for chars, repl in _ACCENT_PAIRS:
        for ch in chars:
            text = text.replace(ch, repl)
    return re.sub(r"[^a-z0-9]+", "-", text).strip("-")


def _same_page(u1: str, u2: str) -> bool:
    """So sánh 2 URL bỏ qua scheme/www/dấu / cuối."""
    if not u1 or not u2:
        return False
    a, b = urlparse(u1), urlparse(u2)
    return (strip_www(a.hostname or "") == strip_www(b.hostname or "")
            and (a.path or "/").rstrip("/") == (b.path or "/").rstrip("/"))


def _is_page_title(name: str, title: str) -> bool:
    """Phần tử breadcrumb có phải chính tiêu đề trang không (khớp gần như hoàn toàn).

    Dùng tỉ lệ độ dài thay vì 'chuỗi con' vì nhiều sản phẩm có tên bắt đầu bằng
    đúng tên danh mục — vd danh mục 'Tủ bảo quản rượu vang' và sản phẩm
    'Tủ bảo quản rượu vang Vinocave 80 chai' là hai thứ khác nhau.
    """
    ns, ts = slugify_vn(name), slugify_vn(title)
    if not ns or not ts:
        return False
    if ns == ts:
        return True
    longer, shorter = (ts, ns) if len(ts) >= len(ns) else (ns, ts)
    return longer.startswith(shorter) and len(shorter) >= 0.85 * len(longer)


def clean_breadcrumb(items, domain: str, title: str, page_url: str = "") -> list:
    """Làm sạch breadcrumb, giữ TỐI ĐA số cấp thư mục.

    items: list các (tên, url) — url có thể None nếu cấp đó không phải link.
    Chỉ bỏ: đoạn rỗng/generic, tên site ở đầu, và cấp cuối nếu đó chính là trang
    hiện tại (không có link, hoặc link trỏ về chính nó, hoặc tên trùng tiêu đề).
    """
    out = []
    for it in items:
        name, url = it if isinstance(it, (tuple, list)) else (it, None)
        name = htmllib.unescape((name or "").strip()).lstrip("#").strip()
        if not name or name.lower() in GENERIC_CRUMBS:
            continue
        if out and out[-1][0].lower() == name.lower():
            continue
        out.append((name, url))
    # Bỏ phần tử đầu nếu là tên site (slug gần trùng domain) — vd "Thế Giới Di Động"
    if out:
        dom = domain.replace("www.", "").split(".")[0].replace("-", "")
        first = slugify_vn(out[0][0]).replace("-", "")
        if first and (first in dom or dom in first):
            out = out[1:]
    # Bỏ cấp cuối CHỈ KHI nó là chính trang hiện tại, không phải danh mục con
    if len(out) >= 2:
        name, url = out[-1]
        is_current = (not url) or _same_page(url, page_url) or _is_page_title(name, title)
        if is_current:
            out = out[:-1]
    return [name for name, _url in out]


class CrawlJob:
    """Một phiên crawl. Chạy bằng asyncio, kết quả gom vào self.records."""

    def __init__(self, config: dict):
        self.mode = config.get("mode", "auto")  # auto | seeds | list | reconstruct
        self.use_wp_api = bool(config.get("use_wp_api", True))
        self.max_urls = min(int(config.get("max_urls", 1000) or 1000), 5000)
        self.concurrency = min(int(config.get("concurrency", 5) or 5), 10)
        self.delay = max(float(config.get("delay", 0.2) or 0.2), 0.05)
        self.respect_robots = bool(config.get("respect_robots", True))
        self.use_wayback = bool(config.get("use_wayback", False))

        patterns = {**DEFAULT_PATTERNS}
        for key in ("category", "product", "blog"):
            user_pats = config.get("patterns", {}).get(key)
            if user_pats:
                patterns[key] = [p for p in user_pats if p.strip()]
        self.patterns = patterns

        # Mẫu loại trừ URL: regex, nếu regex sai thì coi như chuỗi con
        self.exclude_patterns = []
        for pat in config.get("exclude", []) or []:
            pat = pat.strip()
            if not pat:
                continue
            try:
                self.exclude_patterns.append(re.compile(pat, re.I))
            except re.error:
                self.exclude_patterns.append(re.compile(re.escape(pat), re.I))

        raw_seeds = []
        if self.mode in ("seeds", "list", "reconstruct"):
            raw_seeds = [s for s in config.get("seeds", []) if s.strip()]
        else:
            raw_seeds = [config.get("url", "")]
        seeds = []
        for s in raw_seeds:
            s = s.strip()
            if s and not s.lower().startswith(("http://", "https://")):
                s = "https://" + s
            n = normalize_url(s, s)
            if n:
                seeds.append(n)
        if not seeds:
            raise ValueError("Không có URL hợp lệ để bắt đầu")
        self.seeds = seeds
        self.root_host = strip_www(urlparse(seeds[0]).hostname or "")
        # Chế độ list/reconstruct: chỉ xử lý đúng danh sách, giới hạn = số URL đã dán
        if self.mode in ("list", "reconstruct"):
            self.max_urls = len(seeds)
        self._wp_cat_cache = {}  # (host, cat_id) -> {"name","slug","parent"}
        # Chế độ seeds: chỉ crawl URL con nằm trong các thư mục mẹ đã khai báo
        self.seed_paths = [urlparse(s).path or "/" for s in seeds] if self.mode == "seeds" else []

        self.state = "idle"      # idle | running | done | stopped | error
        self.message = ""
        self.records = []        # kết quả, append theo thứ tự crawl
        self.seen = set()
        self.inlinks = {}        # url đích -> {"count": n, "sources": [{"from","anchor"}]}
        self.sitemap_lastmod = {}  # url -> lastmod từ sitemap
        self.sitemap_urls = set()
        self.sitemaps_found = []
        self.orphan_urls = []
        self.queued_count = 0
        self.started_at = None
        self.finished_at = None
        self._stop = False
        self._seq = 0
        self._robots = None

    # ---------- API trạng thái ----------
    def snapshot(self) -> dict:
        by_type = {}
        errors = 0
        for r in self.records:
            by_type[r["page_type"]] = by_type.get(r["page_type"], 0) + 1
            status = r.get("status")
            if not isinstance(status, int) or status >= 400:
                errors += 1
        elapsed = 0
        if self.started_at:
            elapsed = (self.finished_at or time.time()) - self.started_at
        return {
            "state": self.state,
            "message": self.message,
            "fetched": len(self.records),
            "queued": max(self.queued_count - len(self.records), 0),
            "max_urls": self.max_urls,
            "by_type": by_type,
            "errors": errors,
            "elapsed": round(elapsed, 1),
            "recent": [r["url"] for r in self.records[-5:]],
            "sitemap_count": len(self.sitemap_urls),
            "orphan_count": len(self.orphan_urls),
        }

    def stop(self):
        self._stop = True

    # ---------- Lõi crawl ----------
    def _same_site(self, url: str) -> bool:
        return strip_www(urlparse(url).hostname or "") == self.root_host

    def _in_seed_scope(self, url: str) -> bool:
        if self.mode != "seeds":
            return True
        path = urlparse(url).path or "/"
        return any(
            path == sp or path.startswith(sp.rstrip("/") + "/") or (sp == "/" )
            for sp in self.seed_paths
        )

    def _allowed_by_robots(self, url: str) -> bool:
        if not self.respect_robots or self._robots is None:
            return True
        try:
            return self._robots.can_fetch(USER_AGENT, url)
        except Exception:
            return True

    async def _load_robots(self, client: httpx.AsyncClient):
        if not self.respect_robots:
            return
        robots_url = f"{urlparse(self.seeds[0]).scheme}://{urlparse(self.seeds[0]).netloc}/robots.txt"
        try:
            resp = await client.get(robots_url, timeout=10)
            if resp.status_code == 200 and resp.text:
                rp = robotparser.RobotFileParser()
                rp.parse(resp.text.splitlines())
                self._robots = rp
        except Exception:
            self._robots = None

    def _enqueue(self, queue: asyncio.PriorityQueue, url: str, depth: int, referrer: str):
        if url in self.seen or self.queued_count >= self.max_urls * 5:
            return
        # Chế độ list: nhận mọi URL người dùng dán (kể cả khác domain)
        if self.mode != "list" and (not self._same_site(url) or not self._in_seed_scope(url)):
            return
        if self.mode != "list" and any(p.search(url) for p in self.exclude_patterns):
            return
        self.seen.add(url)
        self.queued_count += 1
        page_type = classify(url, self.patterns)
        self._seq += 1
        queue.put_nowait((TYPE_PRIORITY[page_type], depth, self._seq, url, depth, referrer, page_type))

    async def _fetch_one(self, client, url, depth, referrer, page_type):
        path_segments = [s for s in urlparse(url).path.split("/") if s]
        record = {
            "url": url,
            "page_type": page_type,
            "page_type_label": PAGE_TYPE_LABELS[page_type],
            "path_segments": path_segments,
            "folder_depth": len(path_segments),
            "depth": depth,
            "found_on": referrer,
            "status": None,
            "final_url": url,
            "redirect_chain": [],
            "response_ms": None,
            "title": None, "title_length": 0,
            "meta_description": None, "meta_desc_length": 0,
            "h1_count": 0, "h1_texts": [], "h2_count": 0,
            "canonical": None, "canonical_ok": None,
            "meta_robots": None, "noindex": False,
            "og_image": None,
            "images": [], "image_count": 0, "images_missing_alt": 0,
            "word_count": 0, "inlink_count": 0, "inlink_sources": [],
            "date_published": None, "date_modified": None, "date_source": None,
            "error": None,
        }
        links = []
        start = time.time()
        try:
            resp = await client.get(url)
            record["response_ms"] = round((time.time() - start) * 1000)
            record["status"] = resp.status_code
            record["final_url"] = str(resp.url)
            if resp.history:
                record["redirect_chain"] = [
                    f"{h.status_code} → {h.headers.get('location', '')}" for h in resp.history
                ]
            content_type = resp.headers.get("content-type", "")
            if resp.status_code == 200 and "text/html" in content_type and len(resp.text) < 3_000_000:
                links = self._parse_html(resp.text, record)
        except httpx.HTTPError as e:
            record["status"] = "Lỗi"
            record["error"] = f"{type(e).__name__}: {e}"
        except Exception as e:
            record["status"] = "Lỗi"
            record["error"] = f"{type(e).__name__}: {e}"
        return record, links

    def _parse_html(self, html: str, record: dict) -> list:
        soup = BeautifulSoup(html, "lxml")

        if soup.title and soup.title.string:
            record["title"] = soup.title.string.strip()
            record["title_length"] = len(record["title"])

        meta_desc = soup.find("meta", attrs={"name": lambda v: v and v.lower() == "description"})
        if meta_desc and meta_desc.get("content"):
            record["meta_description"] = meta_desc["content"].strip()
            record["meta_desc_length"] = len(record["meta_description"])

        meta_robots = soup.find("meta", attrs={"name": lambda v: v and v.lower() == "robots"})
        if meta_robots and meta_robots.get("content"):
            record["meta_robots"] = meta_robots["content"].strip()
            record["noindex"] = "noindex" in record["meta_robots"].lower()

        canonical_tag = soup.find("link", rel="canonical")
        if canonical_tag and canonical_tag.get("href"):
            canonical = canonical_tag["href"].strip()
            record["canonical"] = canonical
            norm_canon = normalize_url(record["final_url"], canonical)
            norm_self = normalize_url(record["final_url"], record["final_url"])
            record["canonical_ok"] = bool(norm_canon and norm_canon == norm_self)
        else:
            record["canonical_ok"] = False

        og_image = soup.find("meta", attrs={"property": "og:image"})
        if og_image and og_image.get("content"):
            record["og_image"] = urljoin(record["final_url"], og_image["content"].strip())

        self._extract_dates(soup, record)

        h1_texts = [h.get_text(" ", strip=True)[:200] for h in soup.find_all("h1")]
        record["h1_texts"] = h1_texts
        record["h1_count"] = len(h1_texts)
        record["h2_count"] = len(soup.find_all("h2"))

        images, seen_src = [], set()
        for img in soup.find_all("img"):
            src = (
                img.get("src") or img.get("data-src")
                or img.get("data-original") or img.get("data-lazy-src") or ""
            ).strip()
            if not src or src.startswith("data:"):
                continue
            src = urljoin(record["final_url"], src)
            if src in seen_src:
                continue
            seen_src.add(src)
            alt = img.get("alt")
            missing_alt = alt is None or not alt.strip()
            images.append({"src": src, "alt": (alt or "").strip(), "missing_alt": missing_alt})
        record["images"] = images
        record["image_count"] = len(images)
        record["images_missing_alt"] = sum(1 for i in images if i["missing_alt"])

        links, seen_targets = [], set()
        for a in soup.find_all("a", href=True):
            n = normalize_url(record["final_url"], a["href"])
            if not n or n in seen_targets:
                continue
            seen_targets.add(n)
            anchor = a.get_text(" ", strip=True)
            if not anchor:
                img = a.find("img")
                anchor = (img.get("alt") or "").strip() if img else ""
            links.append((n, anchor[:120]))

        # Đếm số từ của nội dung hiển thị (bỏ script/style)
        for tag in soup(["script", "style", "noscript", "template"]):
            tag.decompose()
        record["word_count"] = len(soup.get_text(" ", strip=True).split())
        return links

    def _extract_dates(self, soup, record):
        """Dò ngày đăng / ngày cập nhật qua 7 tầng, tin cậy cao -> thấp.

        1. Meta property Open Graph (article:published_time...) — WordPress, báo chí
        2. Meta name (pubdate, Dublin Core, cXense...) — VnExpress, site .gov.vn
        3. Microdata itemprop (strip khoảng trắng thừa — techcombank)
        4. Schema JSON-LD — site làm SEO tốt
        5. Thẻ <time datetime> có phân biệt nhãn "Cập nhật" — dantri, tuoitre...
        6. Thuộc tính CMS tùy chỉnh (date-created...) — vib
        7. Ngày dạng text quanh H1 (có nhãn cập nhật/đăng) — bidv, vpbank
        """
        sources = []

        def set_pub(value, src):
            if value and not record["date_published"]:
                record["date_published"] = clean_date(value)
                if src not in sources:
                    sources.append(src)

        def set_mod(value, src):
            if value and not record["date_modified"]:
                record["date_modified"] = clean_date(value)
                if src not in sources:
                    sources.append(src)

        # Tầng 1: meta property (Open Graph)
        tag = soup.find("meta", attrs={"property": "article:published_time"})
        if tag:
            set_pub(tag.get("content"), "meta")
        for prop in ("article:modified_time", "og:updated_time"):
            tag = soup.find("meta", attrs={"property": prop})
            if tag:
                set_mod(tag.get("content"), "meta")

        # Tầng 2: meta name (pubdate, Dublin Core dc.*, cXense... — báo chí, site .gov)
        if not record["date_published"]:
            tag = soup.find("meta", attrs={
                "name": lambda v: v and v.strip().lower() in META_PUB_NAMES})
            if tag:
                set_pub(tag.get("content"), f"meta {(tag.get('name') or '').strip()}")
        if not record["date_modified"]:
            tag = soup.find("meta", attrs={
                "name": lambda v: v and v.strip().lower() in META_MOD_NAMES})
            if tag:
                set_mod(tag.get("content"), f"meta {(tag.get('name') or '').strip()}")

        # Tầng 3: itemprop — nhận mọi thẻ (meta/time/span), strip khoảng trắng thừa
        if not record["date_published"]:
            tag = soup.find(attrs={
                "itemprop": lambda v: v and v.strip().lower() == "datepublished"})
            if tag:
                set_pub(tag.get("content") or tag.get("datetime")
                        or tag.get_text(" ", strip=True)[:40], "itemprop")
        if not record["date_modified"]:
            tag = soup.find(attrs={
                "itemprop": lambda v: v and v.strip().lower() == "datemodified"})
            if tag:
                set_mod(tag.get("content") or tag.get("datetime")
                        or tag.get_text(" ", strip=True)[:40], "itemprop")

        # Tầng 4: Schema JSON-LD
        if not record["date_published"] or not record["date_modified"]:
            found = {}
            for script in soup.find_all("script", type="application/ld+json"):
                try:
                    find_jsonld_dates(json.loads(script.string or ""), found)
                except Exception:
                    continue
                if "datePublished" in found and "dateModified" in found:
                    break
            set_pub(found.get("datePublished"), "JSON-LD")
            set_mod(found.get("dateModified"), "JSON-LD")

        # Tầng 5: thẻ <time datetime> — nhãn chứa "cập nhật/updated" thì là ngày sửa
        if not record["date_published"] or not record["date_modified"]:
            for t in soup.find_all("time", attrs={"datetime": True})[:5]:
                dt = (t.get("datetime") or "").strip()
                if not dt:
                    continue
                context = t.get_text(" ", strip=True)
                if t.parent:
                    context += " " + t.parent.get_text(" ", strip=True)[:80]
                if MOD_LABEL_RE.search(context):
                    set_mod(dt, "thẻ time")
                else:
                    set_pub(dt, "thẻ time")

        # Tầng 6: thuộc tính CMS tùy chỉnh (vib.com.vn: date-created / date-up)
        if not record["date_published"]:
            for attr in ("date-created", "data-created", "data-published",
                         "data-date", "data-publish-date"):
                tag = soup.find(attrs={attr: True})
                if tag and tag.get(attr, "").strip():
                    set_pub(tag[attr], f"attr {attr}")
                    break
        if not record["date_modified"]:
            for attr in ("date-up", "date-updated", "data-updated", "data-modified"):
                tag = soup.find(attrs={attr: True})
                if tag and tag.get(attr, "").strip():
                    set_mod(tag[attr], f"attr {attr}")
                    break

        # Tầng 7: ngày dạng text hiển thị quanh H1 (bidv: <i>16/04/2024</i>,
        # vpbank: <p>04/12/2019</p>, vnexpress: "Thứ năm, 4/7/2024, 06:00 (GMT+7)")
        # Nhãn "Cập nhật..." -> ngày sửa; còn lại -> ngày đăng.
        if not record["date_published"] or not record["date_modified"]:
            h1 = soup.find("h1")
            if h1:
                candidates = list(h1.find_all_next(True))[:25]
                candidates += list(h1.find_all_previous(True))[:10]
                for el in candidates:
                    if el.name in ("script", "style", "a", "img", "nav", "header"):
                        continue
                    text = el.get_text(" ", strip=True)
                    if not text or len(text) > 90:
                        continue
                    m = VISIBLE_DATE_RE.search(text)
                    if not m:
                        continue
                    if MOD_LABEL_RE.search(text):
                        set_mod(m.group(0), "text gần H1")
                    else:
                        set_pub(m.group(0), "text gần H1")
                    if record["date_published"] and record["date_modified"]:
                        break

        if sources:
            record["date_source"] = " + ".join(sources)

    async def _worker(self, queue: asyncio.PriorityQueue, client: httpx.AsyncClient):
        while True:
            _, _, _, url, depth, referrer, page_type = await queue.get()
            try:
                if self._stop or len(self.records) >= self.max_urls:
                    continue  # xả queue, không fetch nữa
                if not self._allowed_by_robots(url):
                    continue
                record, links = await self._fetch_one(client, url, depth, referrer, page_type)
                if len(self.records) < self.max_urls:
                    self.records.append(record)
                    for link, anchor in links:
                        self._add_inlink(link, url, anchor)
                    # Chế độ list: không lần theo link, chỉ quét đúng danh sách đã dán
                    if self.mode != "list" and not self._stop and len(self.records) < self.max_urls:
                        for link, _anchor in links:
                            self._enqueue(queue, link, depth + 1, url)
                await asyncio.sleep(self.delay)
            finally:
                queue.task_done()

    def _add_inlink(self, target: str, source: str, anchor: str):
        """Ghi nhận 1 liên kết nội bộ trỏ đến target (giữ tối đa 30 nguồn mẫu)."""
        if self.mode != "list" and not self._same_site(target):
            return
        info = self.inlinks.get(target)
        if info is None:
            info = {"count": 0, "sources": []}
            self.inlinks[target] = info
        info["count"] += 1
        if len(info["sources"]) < 30:
            info["sources"].append({"from": source, "anchor": anchor})

    def _attach_inlinks(self):
        for r in self.records:
            info = self.inlinks.get(r["url"])
            if info:
                r["inlink_count"] = info["count"]
                r["inlink_sources"] = info["sources"]
            # Ngày cập nhật: nếu trang không khai báo thì lấy từ sitemap <lastmod>
            if not r.get("date_modified"):
                lastmod = self.sitemap_lastmod.get(r["url"])
                if lastmod:
                    r["date_modified"] = lastmod
                    src = r.get("date_source")
                    r["date_source"] = (src + " + sitemap") if src else "sitemap"

    async def _load_sitemaps(self, client: httpx.AsyncClient):
        """Đọc sitemap của mọi domain trong seeds (list mode dán được nhiều domain).

        Mục đích: đối chiếu orphan pages (chế độ auto/seeds) và lấy <lastmod>
        làm nguồn dự phòng cho ngày cập nhật (mọi chế độ, kể cả list).
        """
        origins, seen_hosts = [], set()
        for s in self.seeds:
            p = urlparse(s)
            if p.netloc and p.netloc not in seen_hosts:
                seen_hosts.add(p.netloc)
                origins.append((f"{p.scheme}://{p.netloc}", strip_www(p.hostname or "")))
            if len(origins) >= 5:  # giới hạn 5 domain/lần
                break
        for base, root_host in origins:
            await self._load_sitemaps_origin(client, base, root_host)
        # Orphan = có trong sitemap nhưng crawler không hề gặp qua link nội bộ.
        # List mode không tính orphan (chỉ quét URL được dán nên so sánh vô nghĩa).
        if self.sitemap_urls and self.mode != "list":
            self.orphan_urls = sorted(self.sitemap_urls - self.seen)[:1000]

    async def _load_sitemaps_origin(self, client: httpx.AsyncClient, base: str, root_host: str):
        candidates = []
        try:
            resp = await client.get(base + "/robots.txt", timeout=10)
            if resp.status_code == 200:
                candidates.extend(re.findall(r"(?im)^\s*sitemap:\s*(\S+)", resp.text))
        except Exception:
            pass
        candidates.extend([base + "/sitemap.xml", base + "/sitemap_index.xml"])

        def same_origin(url):
            return strip_www(urlparse(url).hostname or "") == root_host

        fetched = set()
        pending = [c for c in candidates if c]
        while pending and len(fetched) < 15 and len(self.sitemap_urls) < 50000:
            sm_url = pending.pop(0)
            if sm_url in fetched:
                continue
            fetched.add(sm_url)
            try:
                resp = await client.get(sm_url, timeout=15)
                if resp.status_code != 200:
                    continue
                text = resp.text
                locs = re.findall(r"<loc>\s*(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?\s*</loc>", text, re.I | re.S)
                if not locs:
                    continue
                self.sitemaps_found.append(sm_url)
                if "<sitemapindex" in text.lower():
                    pending.extend(locs[:30])  # sitemap index -> nạp các sitemap con
                else:
                    # Lấy cả <lastmod> theo từng khối <url> nếu có
                    blocks = re.findall(r"<url>(.*?)</url>", text, re.I | re.S)
                    if blocks:
                        for block in blocks:
                            locm = re.search(
                                r"<loc>\s*(?:<!\[CDATA\[)?(.*?)(?:\]\]>)?\s*</loc>", block, re.I | re.S)
                            if not locm:
                                continue
                            n = normalize_url(locm.group(1), locm.group(1))
                            if n and same_origin(n):
                                self.sitemap_urls.add(n)
                                lastmod = re.search(r"<lastmod>\s*(.*?)\s*</lastmod>", block, re.I | re.S)
                                if lastmod:
                                    self.sitemap_lastmod[n] = clean_date(lastmod.group(1))
                    else:
                        for loc in locs:
                            n = normalize_url(loc, loc)
                            if n and same_origin(n):
                                self.sitemap_urls.add(n)
            except Exception:
                continue

    async def _lookup_wayback(self, client: httpx.AsyncClient):
        """Tra Internet Archive (Wayback Machine) cho các trang thiếu ngày.

        Bản chụp đầu tiên ~ cận trên của ngày đăng (trang tồn tại muộn nhất từ đó).
        Bản chụp cuối ~ lần cuối archive.org thấy trang. Giới hạn 100 trang/lần.
        """
        targets = [
            r for r in self.records
            if r.get("status") == 200 and (not r.get("date_published") or not r.get("date_modified"))
        ][:100]
        if not targets:
            return
        api = "https://web.archive.org/cdx/search/cdx"

        async def first_or_last_snapshot(url, last=False):
            params = {"url": url, "output": "json", "fl": "timestamp",
                      "filter": "statuscode:200", "limit": "-1" if last else "1"}
            # archive.org giới hạn tốc độ rất gắt -> retry với backoff khi bị 429
            for attempt in range(3):
                resp = await client.get(api, params=params, timeout=30)
                if resp.status_code == 429:
                    await asyncio.sleep(10 * (attempt + 1))
                    continue
                if resp.status_code == 200:
                    try:
                        data = resp.json()
                    except ValueError:
                        return None
                    if len(data) > 1 and data[1]:
                        ts = data[1][0]
                        return f"{ts[0:4]}-{ts[4:6]}-{ts[6:8]}"
                return None
            return None

        # Gọi TUẦN TỰ, nghỉ 2.5s giữa các request để không bị archive.org chặn
        for i, r in enumerate(targets):
            if self._stop:
                break
            self.message = f"Đang tra Wayback Machine: {i + 1}/{len(targets)} trang..."
            try:
                if not r.get("date_published"):
                    d = await first_or_last_snapshot(r["url"])
                    if d:
                        r["date_published"] = d
                        src = r.get("date_source")
                        r["date_source"] = ((src + " + ") if src else "") + "Wayback bản lưu đầu"
                    await asyncio.sleep(2.5)
                if not r.get("date_modified"):
                    d = await first_or_last_snapshot(r["url"], last=True)
                    if d and d != r.get("date_published"):
                        r["date_modified"] = d
                        src = r.get("date_source")
                        r["date_source"] = ((src + " + ") if src else "") + "Wayback bản lưu cuối"
                    await asyncio.sleep(2.5)
            except Exception:
                pass
        self.message = ""

    def _reclassify_tree(self):
        """Phân tầng theo cây URL: trang 'Sản phẩm' nhưng có trang con bên dưới
        (là thư mục cha của URL khác) thực chất là trang danh mục."""
        paths = set()
        for r in self.records:
            p = urlparse(r["url"]).path.rstrip("/")
            if p:
                paths.add(p)
        for r in self.records:
            if r["page_type"] != "product":
                continue
            prefix = urlparse(r["url"]).path.rstrip("/") + "/"
            if prefix != "/" and any(other.startswith(prefix) for other in paths):
                r["page_type"] = "category"
                r["page_type_label"] = PAGE_TYPE_LABELS["category"]

    # ===== Chế độ RECONSTRUCT: chỉ tái tạo URL %category%/%post% + phân loại =====

    def _extract_breadcrumb(self, soup, url, title):
        """Trích chuỗi category. Trả về (trail[list], nguồn). Ưu tiên schema -> HTML -> section."""
        domain = urlparse(url).netloc
        # 1. Schema BreadcrumbList (JSON-LD) — sạch và tin cậy nhất
        for s in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(s.string or "")
            except Exception:
                continue
            stack = [data]
            while stack:
                o = stack.pop()
                if isinstance(o, dict):
                    if o.get("@type") == "BreadcrumbList":
                        items = sorted(
                            o.get("itemListElement", []),
                            key=lambda x: x.get("position", 0) if isinstance(x, dict) else 0)
                        pairs = []
                        for it in items:
                            nm = it.get("name")
                            itm = it.get("item")
                            link = None
                            if isinstance(itm, dict):
                                nm = nm or itm.get("name")
                                link = itm.get("@id") or itm.get("url")
                            elif isinstance(itm, str):
                                link = itm
                            if nm:
                                pairs.append((nm, link))
                        trail = clean_breadcrumb(pairs, domain, title, url)
                        if trail:
                            return trail, "schema"
                    stack.extend(o.values())
                elif isinstance(o, list):
                    stack.extend(o)
        # 2. Breadcrumb HTML — lấy <a> kèm href; thêm cả phần tử cuối không có link
        for bc in soup.find_all(class_=re.compile(r"breadcrumb|crumb", re.I)):
            pairs = []
            for a in bc.find_all("a"):
                href = a.get("href")
                pairs.append((a.get_text(" ", strip=True),
                              urljoin(url, href) if href else None))
            trail = clean_breadcrumb(pairs, domain, title, url)
            if trail:
                return trail, "HTML breadcrumb"
        # 3. meta article:section
        m = soup.find("meta", attrs={"property": "article:section"})
        if m and m.get("content"):
            trail = clean_breadcrumb([(m["content"], None)], domain, title, url)
            if trail:
                return trail, "article:section"
        return [], None

    async def _wp_breadcrumb(self, client, url):
        """Dự phòng cho site WordPress: lấy category qua REST API /wp-json."""
        parsed = urlparse(url)
        host = parsed.netloc
        slug = parsed.path.strip("/").split("/")[-1]
        if not slug:
            return [], None
        base = f"{parsed.scheme}://{host}"
        try:
            resp = await client.get(
                f"{base}/wp-json/wp/v2/posts",
                params={"slug": slug, "_fields": "categories"}, timeout=15)
            if resp.status_code != 200:
                return [], None
            posts = resp.json()
            if not posts or not isinstance(posts, list):
                return [], None
            cat_ids = posts[0].get("categories") or []
            if not cat_ids:
                return [], None
            # Lần theo chuỗi cha để dựng đúng thứ tự category (cha -> con)
            cat_id = cat_ids[0]
            chain = []
            for _ in range(5):
                key = (host, cat_id)
                if key not in self._wp_cat_cache:
                    r2 = await client.get(
                        f"{base}/wp-json/wp/v2/categories/{cat_id}",
                        params={"_fields": "name,slug,parent"}, timeout=15)
                    if r2.status_code != 200:
                        break
                    self._wp_cat_cache[key] = r2.json()
                info = self._wp_cat_cache[key]
                if not info or not info.get("name"):
                    break
                chain.insert(0, info["name"])
                parent = info.get("parent") or 0
                if not parent:
                    break
                cat_id = parent
            if chain:
                return chain, "WordPress API"
        except Exception:
            pass
        return [], None

    def _classify_page(self, url, trail):
        """Phân loại trang từ breadcrumb + URL. Trả về (key, label)."""
        p = urlparse(url)
        path = p.path.strip("/")
        blob = (" ".join(trail) + " " + url).lower()
        if not path:
            return "homepage", PAGE_TYPE_LABELS["homepage"]
        blog_kw = ("tin tức", "tin tuc", "/tin-tuc", "blog", "bài viết", "bai viet",
                   "news", "kinh nghiệm", "kinh nghiem", "cẩm nang", "cam nang",
                   "tư vấn", "tu van", "review", "thủ thuật", "thu thuat", "góc",
                   "chia sẻ", "chia se", "sức khỏe", "suc khoe")
        prod_kw = ("sản phẩm", "san pham", "/san-pham", "product", "/collections/",
                   "/products/", "/p/", "/dat-mua")
        if any(k in blob for k in prod_kw):
            return "product", PAGE_TYPE_LABELS["product"]
        if any(k in blob for k in blog_kw):
            return "blog", PAGE_TYPE_LABELS["blog"]
        # Có category + slug bài dài (nhiều từ hoặc có ID số) => trang chi tiết/bài viết
        last_seg = path.split("/")[-1]
        looks_article = last_seg.count("-") >= 3 or bool(re.search(r"-\d{4,}", last_seg))
        if trail and looks_article:
            return "blog", PAGE_TYPE_LABELS["blog"]
        if len(trail) <= 1:
            return "category", PAGE_TYPE_LABELS["category"]
        return "other", PAGE_TYPE_LABELS["other"]

    def _reconstruct_url(self, url, trail):
        """Ghép domain/%category%/%post% từ trail category + slug bài."""
        p = urlparse(url)
        slug = p.path.strip("/").split("/")[-1] or ""
        cat_slugs = [slugify_vn(c) for c in trail if slugify_vn(c)]
        if not cat_slugs:
            return None
        tail = f"/{slug}" if slug else ""
        return f"{p.netloc}/" + "/".join(cat_slugs) + tail

    async def _reconstruct_one(self, client, url):
        record = {
            "url": url, "status": None, "title": None,
            "page_type": "other", "page_type_label": PAGE_TYPE_LABELS["other"],
            "path_segments": [s for s in urlparse(url).path.split("/") if s],
            "folder_depth": len([s for s in urlparse(url).path.split("/") if s]),
            "breadcrumb": [], "breadcrumb_source": None,
            "reconstructed_url": None, "is_wordpress": False, "error": None,
        }
        try:
            resp = await client.get(url)
            record["status"] = resp.status_code
            record["final_url"] = str(resp.url)
            ct = resp.headers.get("content-type", "")
            if resp.status_code == 200 and "text/html" in ct and len(resp.text) < 3_000_000:
                soup = BeautifulSoup(resp.text, "lxml")
                h1 = soup.find("h1")
                title = (h1.get_text(" ", strip=True) if h1
                         else (soup.title.string if soup.title else "")) or ""
                record["title"] = title.strip()[:200]
                record["is_wordpress"] = ("wp-content" in resp.text or "/wp-json" in resp.text)
                trail, src = self._extract_breadcrumb(soup, str(resp.url), title)
                # Dự phòng WordPress API khi không có breadcrumb trong HTML
                if not trail and self.use_wp_api and record["is_wordpress"]:
                    trail, src = await self._wp_breadcrumb(client, str(resp.url))
                record["breadcrumb"] = trail
                record["breadcrumb_source"] = src
                record["reconstructed_url"] = self._reconstruct_url(str(resp.url), trail)
                key, label = self._classify_page(str(resp.url), trail)
                record["page_type"], record["page_type_label"] = key, label
        except httpx.HTTPError as e:
            record["status"] = "Lỗi"
            record["error"] = f"{type(e).__name__}: {e}"
        except Exception as e:
            record["status"] = "Lỗi"
            record["error"] = f"{type(e).__name__}: {e}"
        return record

    async def _run_reconstruct(self, client):
        sem = asyncio.Semaphore(self.concurrency)

        async def process(i, url):
            async with sem:
                if self._stop:
                    return
                rec = await self._reconstruct_one(client, url)
                self.records.append(rec)
                await asyncio.sleep(self.delay)

        await asyncio.gather(*(process(i, u) for i, u in enumerate(self.seeds)))

    async def run(self):
        self.state = "running"
        self.started_at = time.time()
        queue: asyncio.PriorityQueue = asyncio.PriorityQueue()
        try:
            async with httpx.AsyncClient(
                headers={"User-Agent": USER_AGENT, "Accept-Language": "vi,en;q=0.8"},
                follow_redirects=True,
                timeout=httpx.Timeout(20.0),
                limits=httpx.Limits(max_connections=self.concurrency + 2),
            ) as client:
                if self.mode == "reconstruct":
                    await self._run_reconstruct(client)
                    self.state = "stopped" if self._stop else "done"
                    return
                await self._load_robots(client)
                for seed in self.seeds:
                    self._enqueue(queue, seed, 0, "(seed)")
                workers = [
                    asyncio.create_task(self._worker(queue, client))
                    for _ in range(self.concurrency)
                ]
                await queue.join()
                for w in workers:
                    w.cancel()
                await self._load_sitemaps(client)
                self._reclassify_tree()
                self._attach_inlinks()
                if self.use_wayback:
                    await self._lookup_wayback(client)
            self.state = "stopped" if self._stop else "done"
        except Exception as e:
            self.state = "error"
            self.message = f"{type(e).__name__}: {e}"
        finally:
            self.finished_at = time.time()
