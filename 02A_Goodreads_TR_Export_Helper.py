from __future__ import annotations

import csv
import html
import json
import random
import re
import time
from dataclasses import asdict, dataclass
from datetime import UTC, datetime
from html.parser import HTMLParser
from http.cookiejar import CookieJar
from pathlib import Path
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse
from urllib.request import HTTPCookieProcessor, Request, build_opener


# ==================== 1. GOODREADS AYARLARI VE DESENLER ====================
GOODREADS_BASE_URL = "https://www.goodreads.com"
NEXT_DATA_RE = re.compile(
    r"<script id=[\"']__NEXT_DATA__[\"'] type=[\"']application/json[\"']>(.*?)</script>",
    re.DOTALL,
)
APP_BUNDLE_RE = re.compile(
    r"<script[^>]+src=([\"'])(?P<src>[^\"']*/_next/static/chunks/pages/_app-[^\"']+\.js)\1",
    re.IGNORECASE,
)
PRODUCTION_GRAPHQL_RE = re.compile(
    r'"Production":\{"auth":\{.*?\},"graphql":\{"apiKey":"(?P<api_key>[^"]+)","endpoint":"(?P<endpoint>https://[^"]+/graphql)"',
    re.DOTALL,
)
BOOK_ID_RE = re.compile(r"/book/show/(?P<book_id>\d+)")
WHITESPACE_RE = re.compile(r"\s+")
BLOCK_TAG_RE = re.compile(r"(?i)<\s*(br|/p|/div|/li|/h\d)\b[^>]*>")
TAG_RE = re.compile(r"<[^>]+>")
URL_RE = re.compile(r"https?://\S+|www\.\S+", re.IGNORECASE)

REVIEWS_QUERY = """query getReviews($filters: BookReviewsFilterInput!, $pagination: PaginationInput) {
  getReviews(filters: $filters, pagination: $pagination) {
    totalCount
    pageInfo {
      prevPageToken
      nextPageToken
    }
    edges {
      node {
        id
        rating
        text
        createdAt
        creator {
          name
        }
      }
    }
  }
}"""

BOOK_HEADERS = [
    "kitap_adi",
    "yayin_yili",
    "yazar",
    "ortalama_puan",
    "puan_sayisi",
    "metin_yorum_sayisi",
    "work_id",
    "kaynak_url",
]

REVIEW_HEADERS = [
    "kitap_adi",
    "yayin_yili",
    "yazar",
    "yorumcu_adi",
    "puan",
    "yorum_tarihi",
    "yorum_dili",
    "yorum",
    "review_id",
    "review_url",
        "yorum_sayfasi_url",
]


# ==================== 2. VERI MODELLERI ====================
@dataclass(slots=True)
class ExportConfig:
    list_url: str
    review_language_code: str = "tr"
    output_dir: str = "./output_list_tr"
    max_books: int | None = None
    request_delay_seconds: float = 0.15
    http_retry_limit: int = 5
    graphql_page_retry_limit: int = 6
    book_retry_limit: int = 5
    review_page_size: int = 30
    request_timeout_seconds: int = 30
    user_agent: str = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )


@dataclass(slots=True)
class GraphQLConfig:
    endpoint: str
    api_key: str


@dataclass(slots=True)
class BookMetadata:
    source_url: str
    legacy_book_id: str
    work_id: str
    title: str
    author: str
    publication_year: int | None
    average_rating: float | None
    ratings_count: int | None
    text_reviews_count: int | None
    stats_language_reviews_count: int


class BookLinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.book_urls: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag.lower() != "a":
            return
        values = dict(attrs)
        class_attr = values.get("class") or ""
        href = values.get("href")
        if href and "bookTitle" in class_attr.split():
            self.book_urls.append(urljoin(GOODREADS_BASE_URL, href))


# ==================== 3. URL VE SAYFA PARSE YARDIMCILARI ====================
def normalize_goodreads_url(url: str, keep_query: bool = False) -> str:
    parsed = urlparse(url.strip())
    if not parsed.scheme:
        parsed = urlparse(urljoin(GOODREADS_BASE_URL, url.strip()))
    clean = parsed._replace(
        scheme="https",
        netloc=parsed.netloc or urlparse(GOODREADS_BASE_URL).netloc,
        path=(parsed.path or "/").rstrip("/") or parsed.path or "/",
        query=parsed.query if keep_query else "",
        fragment="",
    )
    return urlunparse(clean)


def build_list_page_url(list_url: str, page_number: int) -> str:
    parsed = urlparse(normalize_goodreads_url(list_url, keep_query=True))
    query_items = [(key, value) for key, value in parse_qsl(parsed.query, keep_blank_values=True) if key != "page"]
    query_items.append(("page", str(page_number)))
    return urlunparse(parsed._replace(query=urlencode(query_items)))


def extract_next_data(page_html: str) -> dict[str, Any]:
    match = NEXT_DATA_RE.search(page_html)
    if not match:
        raise ValueError("Goodreads __NEXT_DATA__ bulunamadi.")
    return json.loads(match.group(1))


def extract_app_bundle_url(page_html: str) -> str:
    match = APP_BUNDLE_RE.search(page_html)
    if not match:
        raise ValueError("Goodreads _app bundle URL bulunamadi.")
    return urljoin(GOODREADS_BASE_URL, match.group("src"))


def parse_graphql_config(bundle_js: str) -> GraphQLConfig:
    match = PRODUCTION_GRAPHQL_RE.search(bundle_js)
    if not match:
        raise ValueError("GraphQL config parse edilemedi.")
    return GraphQLConfig(endpoint=match.group("endpoint"), api_key=match.group("api_key"))


# ==================== 4. TARIH VE YORUM TEMIZLEME ====================
def timestamp_ms_to_iso_date(timestamp_ms: int | None) -> str:
    if timestamp_ms in (None, ""):
        return ""
    return datetime.fromtimestamp(int(timestamp_ms) / 1000, tz=UTC).strftime("%Y-%m-%d")


def timestamp_ms_to_year(timestamp_ms: int | None) -> int | None:
    if timestamp_ms in (None, ""):
        return None
    return datetime.fromtimestamp(int(timestamp_ms) / 1000, tz=UTC).year


def clean_review_html(review_html: str) -> str:
    if not review_html:
        return ""
    text = BLOCK_TAG_RE.sub("\n", review_html)
    text = TAG_RE.sub(" ", text)
    text = html.unescape(text)
    text = URL_RE.sub(" ", text)
    text = WHITESPACE_RE.sub(" ", text)
    return text.strip()


# ==================== 5. KITAP METADATA CIKARMA ====================
def _iter_apollo_entries(apollo_state: dict[str, Any], prefix: str) -> list[dict[str, Any]]:
    return [value for key, value in apollo_state.items() if key.startswith(prefix) and isinstance(value, dict)]


def _extract_requested_book_id(source_url: str, next_data: dict[str, Any]) -> str | None:
    match = BOOK_ID_RE.search(source_url)
    if match:
        return match.group("book_id")
    for query in (next_data.get("query") or {}, next_data.get("props", {}).get("pageProps", {}).get("query") or {}):
        raw = query.get("book_id")
        if raw:
            return str(raw).split(".", 1)[0]
    return None


def _pick_book_object(apollo_state: dict[str, Any], requested_book_id: str | None, source_url: str) -> dict[str, Any]:
    book_objects = _iter_apollo_entries(apollo_state, "Book:")
    if not book_objects:
        raise ValueError("Book object bulunamadi.")
    if requested_book_id:
        for book in book_objects:
            if str(book.get("legacyId") or "") == requested_book_id:
                return book
    normalized_source = normalize_goodreads_url(source_url)
    for book in book_objects:
        web_url = normalize_goodreads_url(str(book.get("webUrl") or normalized_source))
        if web_url == normalized_source:
            return book
    return book_objects[0]


def _pick_work_object(apollo_state: dict[str, Any], book_object: dict[str, Any]) -> dict[str, Any]:
    works = _iter_apollo_entries(apollo_state, "Work:")
    if not works:
        raise ValueError("Work object bulunamadi.")
    best_book_ref = (book_object.get("work") or {}).get("__ref")
    if best_book_ref:
        for work in works:
            if work.get("id") == best_book_ref:
                return work
    selected_book_ref = book_object.get("id")
    for work in works:
        if ((work.get("bestBook") or {}).get("__ref")) == selected_book_ref:
            return work
    return works[0]


def _unique_preserve_order(values: list[str]) -> list[str]:
    unique: list[str] = []
    seen: set[str] = set()
    for value in values:
        cleaned = value.strip()
        if not cleaned or cleaned in seen:
            continue
        seen.add(cleaned)
        unique.append(cleaned)
    return unique


def _extract_author_names(apollo_state: dict[str, Any], book_object: dict[str, Any]) -> str:
    authors: list[str] = []
    edge = book_object.get("primaryContributorEdge") or {}
    node = edge.get("node") or {}
    if isinstance(node, dict) and node.get("name"):
        authors.append(str(node["name"]))
    contributor = book_object.get("primaryContributor") or {}
    if contributor.get("name"):
        authors.append(str(contributor["name"]))
    for contributor_object in _iter_apollo_entries(apollo_state, "Contributor:"):
        if contributor_object.get("name"):
            authors.append(str(contributor_object["name"]))
    return "; ".join(_unique_preserve_order(authors))


def _extract_publication_year(book_object: dict[str, Any], work_object: dict[str, Any]) -> int | None:
    candidate_times = [
        ((book_object.get("details") or {}).get("publicationTime")),
        ((work_object.get("details") or {}).get("publicationTime")),
    ]
    for raw_value in candidate_times:
        if raw_value in (None, ""):
            continue
        try:
            return timestamp_ms_to_year(int(raw_value))
        except Exception:
            continue
    return None


# ==================== 6. GOODREADS EXPORTER SINIFI ====================
class GoodreadsTurkishReviewExporter:
    def __init__(self, config: ExportConfig) -> None:
        self.config = config
        self._cookie_jar = CookieJar()
        self._opener = self._build_opener()
        self._graphql_config: GraphQLConfig | None = None
        self._graphql_seed_url: str | None = None

    def _build_opener(self):
        return build_opener(HTTPCookieProcessor(self._cookie_jar))

    def reset_transport(self) -> None:
        self._cookie_jar = CookieJar()
        self._opener = self._build_opener()

    def _sleep(self, seconds: float | None = None) -> None:
        duration = self.config.request_delay_seconds if seconds is None else seconds
        if duration > 0:
            time.sleep(duration)

    def request_text(
        self,
        method: str,
        url: str,
        data: bytes | None = None,
        headers: dict[str, str] | None = None,
        retry_limit: int | None = None,
    ) -> str:
        attempt_limit = retry_limit or self.config.http_retry_limit
        request_url = url if urlparse(url).scheme else urljoin(GOODREADS_BASE_URL, url)
        base_headers = {"User-Agent": self.config.user_agent}
        if headers:
            base_headers.update(headers)

        last_error: Exception | None = None
        for attempt in range(1, attempt_limit + 1):
            if attempt > 1:
                self._sleep(min(6.0, (1.4 ** (attempt - 1)) + random.random()))
            else:
                self._sleep()

            request = Request(request_url, data=data, headers=base_headers, method=method.upper())
            try:
                with self._opener.open(request, timeout=self.config.request_timeout_seconds) as response:
                    charset = response.headers.get_content_charset() or "utf-8"
                    return response.read().decode(charset, errors="replace")
            except HTTPError as exc:
                last_error = exc
                if exc.code in (408, 425, 429, 500, 502, 503, 504):
                    continue
                raise
            except URLError as exc:
                last_error = exc
                continue

        if last_error is None:
            raise RuntimeError(f"HTTP request failed without an exception: {request_url}")
        raise last_error

    def request_json(
        self,
        method: str,
        url: str,
        data: bytes | None = None,
        headers: dict[str, str] | None = None,
        retry_limit: int | None = None,
    ) -> dict[str, Any]:
        return json.loads(self.request_text(method=method, url=url, data=data, headers=headers, retry_limit=retry_limit))

    def discover_graphql_config(self, seed_book_url: str) -> GraphQLConfig:
        page_html = self.request_text("GET", seed_book_url)
        bundle_url = extract_app_bundle_url(page_html)
        bundle_js = self.request_text("GET", bundle_url)
        config = parse_graphql_config(bundle_js)
        self._graphql_config = config
        self._graphql_seed_url = seed_book_url
        return config

    def ensure_graphql_config(self, seed_book_url: str) -> GraphQLConfig:
        if self._graphql_config is None:
            return self.discover_graphql_config(seed_book_url)
        return self._graphql_config

    def refresh_graphql_config(self, seed_book_url: str) -> GraphQLConfig:
        self._graphql_config = None
        return self.discover_graphql_config(seed_book_url)

    def extract_book_urls_from_list_html(self, page_html: str) -> list[str]:
        parser = BookLinkParser()
        parser.feed(page_html)
        unique_urls: list[str] = []
        seen: set[str] = set()
        for raw_url in parser.book_urls:
            normalized = normalize_goodreads_url(raw_url)
            if normalized in seen:
                continue
            seen.add(normalized)
            unique_urls.append(normalized)
        return unique_urls

    def iter_list_book_urls(self) -> list[str]:
        all_urls: list[str] = []
        seen: set[str] = set()
        page_number = 1
        while True:
            page_url = build_list_page_url(self.config.list_url, page_number)
            page_html = self.request_text("GET", page_url)
            page_book_urls = self.extract_book_urls_from_list_html(page_html)
            if not page_book_urls:
                break
            new_urls = [url for url in page_book_urls if url not in seen]
            if not new_urls:
                break
            for url in new_urls:
                seen.add(url)
                all_urls.append(url)
                if self.config.max_books is not None and len(all_urls) >= self.config.max_books:
                    return all_urls
            page_number += 1
        return all_urls

    def fetch_book_metadata(self, book_url: str) -> BookMetadata:
        page_html = self.request_text("GET", book_url)
        next_data = extract_next_data(page_html)
        page_props = next_data.get("props", {}).get("pageProps", {})
        apollo_state = page_props.get("apolloState") or {}
        if not isinstance(apollo_state, dict):
            raise ValueError("apolloState bulunamadi.")
        requested_book_id = _extract_requested_book_id(book_url, next_data)
        book_object = _pick_book_object(apollo_state, requested_book_id, book_url)
        work_object = _pick_work_object(apollo_state, book_object)
        stats = work_object.get("stats") or {}
        work_id = str(work_object.get("id") or "").strip()
        if not work_id:
            raise ValueError("work_id bulunamadi.")

        language_counts = stats.get("textReviewsLanguageCounts") or []
        stats_language_reviews_count = 0
        for item in language_counts:
            if item.get("isoLanguageCode") == self.config.review_language_code:
                stats_language_reviews_count = int(item.get("count") or 0)
                break

        legacy_book_id = requested_book_id or str(book_object.get("legacyId") or "").strip()
        return BookMetadata(
            source_url=normalize_goodreads_url(book_url),
            legacy_book_id=legacy_book_id,
            work_id=work_id,
            title=str(book_object.get("title") or book_object.get("titleComplete") or "").strip(),
            author=_extract_author_names(apollo_state, book_object),
            publication_year=_extract_publication_year(book_object, work_object),
            average_rating=stats.get("averageRating"),
            ratings_count=stats.get("ratingsCount"),
            text_reviews_count=stats.get("textReviewsCount"),
            stats_language_reviews_count=stats_language_reviews_count,
        )

    def fetch_review_page(
        self,
        graph_config: GraphQLConfig,
        work_id: str,
        language_code: str,
        after_token: str | None,
    ) -> dict[str, Any]:
        filters: dict[str, Any] = {"resourceType": "WORK", "resourceId": work_id, "languageCode": language_code}
        pagination: dict[str, Any] = {"limit": self.config.review_page_size}
        if after_token:
            pagination["after"] = after_token

        payload = json.dumps(
            {
                "query": REVIEWS_QUERY,
                "operationName": "getReviews",
                "variables": {"filters": filters, "pagination": pagination},
            }
        ).encode("utf-8")

        response = self.request_json(
            "POST",
            graph_config.endpoint,
            data=payload,
            headers={"Content-Type": "application/json", "x-api-key": graph_config.api_key},
            retry_limit=self.config.http_retry_limit,
        )
        errors = response.get("errors") or []
        if errors:
            message = errors[0].get("message") or "Unknown Goodreads GraphQL error"
            raise RuntimeError(message)
        data = response.get("data") or {}
        return data.get("getReviews") or {}

    def fetch_review_page_with_retries(
        self,
        seed_book_url: str,
        work_id: str,
        after_token: str | None,
    ) -> dict[str, Any]:
        last_error: Exception | None = None
        for attempt in range(1, self.config.graphql_page_retry_limit + 1):
            if attempt == 1:
                graph_config = self.ensure_graphql_config(seed_book_url)
            else:
                self.reset_transport()
                graph_config = self.refresh_graphql_config(seed_book_url)
                self._sleep(min(10.0, (1.8 ** (attempt - 1)) + random.random()))
            try:
                return self.fetch_review_page(
                    graph_config=graph_config,
                    work_id=work_id,
                    language_code=self.config.review_language_code,
                    after_token=after_token,
                )
            except Exception as exc:
                last_error = exc
        if last_error is None:
            raise RuntimeError("Goodreads review page retry exhausted without exception.")
        raise last_error

    def collect_reviews_for_book(self, book: BookMetadata) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        last_error: str | None = None
        aggregate_reviews: dict[str, dict[str, Any]] = {}
        aggregate_graphql_counts: set[int] = set()
        best_audit: dict[str, Any] = {
            "success": False,
            "attempts_used": 0,
            "graphql_total_count": None,
            "graphql_total_candidates": [],
            "stats_language_reviews_count": book.stats_language_reviews_count,
            "pages_fetched": 0,
            "reviews_fetched": 0,
            "last_error": None,
        }

        for attempt in range(1, self.config.book_retry_limit + 1):
            if attempt > 1:
                self.reset_transport()
                self.refresh_graphql_config(book.source_url)
                self._sleep(min(15.0, (2.0 ** (attempt - 1)) + random.random()))

            reviews: list[dict[str, Any]] = []
            seen_review_ids: set[str] = set()
            next_token: str | None = None
            graphql_total_count: int | None = None
            graphql_total_candidates: set[int] = set()
            pages_fetched = 0

            try:
                while True:
                    review_page = self.fetch_review_page_with_retries(
                        seed_book_url=book.source_url,
                        work_id=book.work_id,
                        after_token=next_token,
                    )
                    pages_fetched += 1

                    page_total = int(review_page.get("totalCount") or 0)
                    if graphql_total_count is None:
                        graphql_total_count = page_total
                    graphql_total_candidates.add(page_total)

                    for edge in review_page.get("edges") or []:
                        review_node = (edge or {}).get("node") or {}
                        review_id = str(review_node.get("id") or "").strip()
                        if not review_id or review_id in seen_review_ids:
                            continue

                        cleaned = clean_review_html(review_node.get("text") or "")
                        creator = review_node.get("creator") or {}
                        reviews.append(
                            {
                                "kitap_adi": book.title,
                                "yayin_yili": book.publication_year,
                                "yazar": book.author,
                                "yorumcu_adi": creator.get("name") or "",
                                "puan": review_node.get("rating"),
                                "yorum_tarihi": timestamp_ms_to_iso_date(review_node.get("createdAt")),
                                "yorum_dili": self.config.review_language_code,
                                "yorum": cleaned,
                                "review_id": review_id,
                                "review_url": "",
                                "yorum_sayfasi_url": "",
                            }
                        )
                        seen_review_ids.add(review_id)

                    next_token = (review_page.get("pageInfo") or {}).get("nextPageToken")
                    if not next_token:
                        break

                verified_total = max(graphql_total_candidates) if graphql_total_candidates else (graphql_total_count or 0)
                aggregate_graphql_counts.update(graphql_total_candidates or {verified_total})
                for row in reviews:
                    review_id = str(row["review_id"])
                    if review_id not in aggregate_reviews:
                        aggregate_reviews[review_id] = row

                aggregate_count = len(aggregate_reviews)
                best_audit = {
                    "success": False,
                    "attempts_used": attempt,
                    "graphql_total_count": verified_total,
                    "graphql_total_candidates": sorted(aggregate_graphql_counts),
                    "stats_language_reviews_count": book.stats_language_reviews_count,
                    "pages_fetched": pages_fetched,
                    "reviews_fetched": aggregate_count,
                    "last_error": None,
                }

                if len(reviews) == verified_total:
                    best_audit["success"] = True
                    return reviews, best_audit

                if aggregate_count == max(aggregate_graphql_counts or {verified_total}):
                    best_audit["success"] = True
                    return list(aggregate_reviews.values()), best_audit

                last_error = (
                    "Tamlik kontrolu basarisiz: "
                    f"graphql_total_count={verified_total}, "
                    f"graphql_total_candidates={sorted(aggregate_graphql_counts)}, "
                    f"stats_total_count={book.stats_language_reviews_count}, "
                    f"reviews_fetched={aggregate_count}"
                )
                best_audit["last_error"] = last_error
            except Exception as exc:
                last_error = str(exc)
                best_audit = {
                    "success": False,
                    "attempts_used": attempt,
                    "graphql_total_count": best_audit.get("graphql_total_count"),
                    "graphql_total_candidates": sorted(aggregate_graphql_counts),
                    "stats_language_reviews_count": book.stats_language_reviews_count,
                    "pages_fetched": best_audit.get("pages_fetched", 0),
                    "reviews_fetched": len(aggregate_reviews),
                    "last_error": last_error,
                }

        return [], {
            "success": False,
            "attempts_used": self.config.book_retry_limit,
            "graphql_total_count": best_audit.get("graphql_total_count"),
            "graphql_total_candidates": sorted(aggregate_graphql_counts),
            "stats_language_reviews_count": book.stats_language_reviews_count,
            "pages_fetched": best_audit.get("pages_fetched", 0),
            "reviews_fetched": len(aggregate_reviews),
            "last_error": last_error,
        }


# ==================== 7. LISTE EXPORT AKISI ====================
def export_goodreads_list(config: ExportConfig) -> dict[str, Any]:
    exporter = GoodreadsTurkishReviewExporter(config)
    output_dir = Path(config.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    books_csv_path = output_dir / "books.csv"
    reviews_csv_path = output_dir / "reviews.csv"
    summary_json_path = output_dir / "run_summary.json"

    book_urls = exporter.iter_list_book_urls()
    run_errors: list[dict[str, Any]] = []
    per_book_audit: list[dict[str, Any]] = []
    unique_reviews_written = 0

    with books_csv_path.open("w", newline="", encoding="utf-8-sig") as books_handle, reviews_csv_path.open(
        "w", newline="", encoding="utf-8-sig"
    ) as reviews_handle:
        books_writer = csv.DictWriter(books_handle, fieldnames=BOOK_HEADERS, delimiter=";")
        reviews_writer = csv.DictWriter(reviews_handle, fieldnames=REVIEW_HEADERS, delimiter=";")
        books_writer.writeheader()
        reviews_writer.writeheader()

        for index, book_url in enumerate(book_urls, start=1):
            print(f"[{index}/{len(book_urls)}] Book metadata: {book_url}")

            try:
                book = exporter.fetch_book_metadata(book_url)
            except Exception as exc:
                error_item = {
                    "stage": "metadata",
                    "book_url": book_url,
                    "error": str(exc),
                }
                run_errors.append(error_item)
                per_book_audit.append(
                    {
                        "book_url": book_url,
                        "title": None,
                        "work_id": None,
                        "success": False,
                        "last_error": str(exc),
                    }
                )
                print(f"  [!] Metadata hatasi: {exc}")
                continue

            books_writer.writerow(
                {
                    "kitap_adi": book.title,
                    "yayin_yili": book.publication_year,
                    "yazar": book.author,
                    "ortalama_puan": book.average_rating,
                    "puan_sayisi": book.ratings_count,
                    "metin_yorum_sayisi": book.text_reviews_count,
                    "work_id": book.work_id,
                    "kaynak_url": book.source_url,
                }
            )
            books_handle.flush()

            reviews, audit = exporter.collect_reviews_for_book(book)
            audit.update(
                {
                    "book_url": book.source_url,
                    "title": book.title,
                    "work_id": book.work_id,
                }
            )
            per_book_audit.append(audit)

            if not audit["success"]:
                run_errors.append(
                    {
                        "stage": "reviews",
                        "book_url": book.source_url,
                        "work_id": book.work_id,
                        "error": audit["last_error"],
                    }
                )
                print(f"  [!] Yorumlar tamamlanamadi: {audit['last_error']}")
                continue

            for row in reviews:
                reviews_writer.writerow(row)
            reviews_handle.flush()
            unique_reviews_written += len(reviews)

            print(
                "  reviews written for book: "
                f"{len(reviews)} (graphql_total={audit['graphql_total_count']}, "
                f"stats_total={audit['stats_language_reviews_count']}, attempts={audit['attempts_used']})"
            )

    successful_books = [item for item in per_book_audit if item.get("success")]
    failed_books = [item for item in per_book_audit if not item.get("success")]
    stats_graphql_mismatches = [
        {
            "book_url": item["book_url"],
            "title": item["title"],
            "stats_language_reviews_count": item["stats_language_reviews_count"],
            "graphql_total_count": item["graphql_total_count"],
        }
        for item in successful_books
        if item.get("graphql_total_count") != item.get("stats_language_reviews_count")
    ]

    summary = {
        "config": asdict(config),
        "books_discovered": len(book_urls),
        "books_completed": len(successful_books),
        "books_failed": len(failed_books),
        "unique_reviews_written": unique_reviews_written,
        "errors": run_errors,
        "stats_graphql_mismatches": stats_graphql_mismatches,
        "per_book_audit": per_book_audit,
        "output_dir": str(output_dir.resolve()),
        "books_csv_path": str(books_csv_path.resolve()),
        "reviews_csv_path": str(reviews_csv_path.resolve()),
        "summary_json_path": str(summary_json_path.resolve()),
    }

    summary_json_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    if failed_books:
        failed_titles = [item.get("title") or item["book_url"] for item in failed_books]
        raise RuntimeError(
            "Goodreads export tamamlanamadi. Eksik kalan kitaplar: " + ", ".join(failed_titles)
        )

    return summary
