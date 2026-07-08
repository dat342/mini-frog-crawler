# -*- coding: utf-8 -*-
"""Lõi crawler: quét URL theo ưu tiên loại trang, bóc tách dữ liệu SEO on-page."""
import asyncio
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


class CrawlJob:
    """Một phiên crawl. Chạy bằng asyncio, kết quả gom vào self.records."""

    def __init__(self, config: dict):
        self.mode = config.get("mode", "auto")  # auto | seeds | list
        self.max_urls = min(int(config.get("max_urls", 1000) or 1000), 5000)
        self.concurrency = min(int(config.get("concurrency", 5) or 5), 10)
        self.delay = max(float(config.get("delay", 0.2) or 0.2), 0.05)
        self.respect_robots = bool(config.get("respect_robots", True))

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
        if self.mode in ("seeds", "list"):
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
        # Chế độ list: chỉ quét đúng danh sách, giới hạn = số URL đã dán
        if self.mode == "list":
            self.max_urls = len(seeds)
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

        # Ngày đăng / ngày cập nhật: meta article:* -> itemprop -> JSON-LD
        sources = []
        for prop in ("article:published_time",):
            tag = soup.find("meta", attrs={"property": prop})
            if tag and tag.get("content"):
                record["date_published"] = clean_date(tag["content"])
                sources.append("meta")
                break
        for prop in ("article:modified_time", "og:updated_time"):
            tag = soup.find("meta", attrs={"property": prop})
            if tag and tag.get("content"):
                record["date_modified"] = clean_date(tag["content"])
                if "meta" not in sources:
                    sources.append("meta")
                break
        # So khớp itemprop có strip() vì nhiều site viết thừa khoảng trắng
        # (vd techcombank.com: itemprop=" datePublished")
        if not record["date_published"]:
            tag = soup.find("meta", attrs={
                "itemprop": lambda v: v and v.strip().lower() == "datepublished"})
            if tag and tag.get("content"):
                record["date_published"] = clean_date(tag["content"])
                sources.append("itemprop")
        if not record["date_modified"]:
            tag = soup.find("meta", attrs={
                "itemprop": lambda v: v and v.strip().lower() == "datemodified"})
            if tag and tag.get("content"):
                record["date_modified"] = clean_date(tag["content"])
                if "itemprop" not in sources:
                    sources.append("itemprop")
        if not record["date_published"] or not record["date_modified"]:
            found = {}
            for script in soup.find_all("script", type="application/ld+json"):
                try:
                    find_jsonld_dates(json.loads(script.string or ""), found)
                except Exception:
                    continue
                if "datePublished" in found and "dateModified" in found:
                    break
            if not record["date_published"] and found.get("datePublished"):
                record["date_published"] = clean_date(found["datePublished"])
                sources.append("JSON-LD")
            if not record["date_modified"] and found.get("dateModified"):
                record["date_modified"] = clean_date(found["dateModified"])
                if "JSON-LD" not in sources:
                    sources.append("JSON-LD")
        # Dự phòng cuối: thẻ <time datetime> hoặc thuộc tính ngày tùy chỉnh
        # của CMS (vd vib.com.vn dùng date-created / date-up)
        if not record["date_published"]:
            tag = soup.find("time", attrs={"datetime": True})
            if tag and tag.get("datetime", "").strip():
                record["date_published"] = clean_date(tag["datetime"])
                sources.append("thẻ time")
        if not record["date_published"]:
            for attr in ("date-created", "data-created", "data-published", "data-date"):
                tag = soup.find(attrs={attr: True})
                if tag and tag.get(attr, "").strip():
                    record["date_published"] = clean_date(tag[attr])
                    sources.append(f"attr {attr}")
                    break
        if not record["date_modified"]:
            for attr in ("date-up", "date-updated", "data-updated", "data-modified"):
                tag = soup.find(attrs={attr: True})
                if tag and tag.get(attr, "").strip():
                    record["date_modified"] = clean_date(tag[attr])
                    sources.append(f"attr {attr}")
                    break
        if sources:
            record["date_source"] = " + ".join(sources)

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
        """Đọc sitemap.xml (kể cả sitemap index) để đối chiếu tìm orphan pages."""
        candidates = []
        if self._robots is not None:
            try:
                candidates.extend(self._robots.site_maps() or [])
            except Exception:
                pass
        parsed = urlparse(self.seeds[0])
        base = f"{parsed.scheme}://{parsed.netloc}"
        candidates.extend([base + "/sitemap.xml", base + "/sitemap_index.xml"])

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
                            if n and self._same_site(n):
                                self.sitemap_urls.add(n)
                                lastmod = re.search(r"<lastmod>\s*(.*?)\s*</lastmod>", block, re.I | re.S)
                                if lastmod:
                                    self.sitemap_lastmod[n] = clean_date(lastmod.group(1))
                    else:
                        for loc in locs:
                            n = normalize_url(loc, loc)
                            if n and self._same_site(n):
                                self.sitemap_urls.add(n)
            except Exception:
                continue
        # Orphan = có trong sitemap nhưng crawler không hề gặp qua link nội bộ
        if self.sitemap_urls:
            self.orphan_urls = sorted(self.sitemap_urls - self.seen)[:1000]

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
                if self.mode != "list":
                    await self._load_sitemaps(client)
            self._reclassify_tree()
            self._attach_inlinks()
            self.state = "stopped" if self._stop else "done"
        except Exception as e:
            self.state = "error"
            self.message = f"{type(e).__name__}: {e}"
        finally:
            self.finished_at = time.time()
