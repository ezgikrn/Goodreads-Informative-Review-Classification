from __future__ import annotations

import argparse
import html
import importlib.util
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any
from urllib.parse import urlencode, urljoin
from xml.sax.saxutils import escape


# ==================== 1. AYARLAR VE DESENLER ====================


INPUT_XLSX_NAME = "MEB 100 temel eser listesi (ortaöğretim).xlsx"
OUTPUT_XLSX_NAME = "MEB 100 temel eser listesi (ortaöğretim) Goodreads Türkçe yorumlar.xlsx"
EXPORTER_MODULE_PATH = Path(__file__).resolve().parent / "02A_Goodreads_TR_Export_Helper.py"
GOODREADS_BASE_URL = "https://www.goodreads.com"

XML_NS = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL_NS = {"r": "http://schemas.openxmlformats.org/package/2006/relationships"}

TAG_RE = re.compile(r"<[^>]+>")
WHITESPACE_RE = re.compile(r"\s+")
SEARCH_ROW_RE = re.compile(
    r"<tr[^>]+itemscope[^>]+itemtype=[\"']http://schema\.org/Book[\"'][^>]*>(?P<body>.*?)</tr>",
    re.IGNORECASE | re.DOTALL,
)
BOOK_TITLE_RE = re.compile(
    r'<a[^>]+class="bookTitle"[^>]+href="(?P<href>[^"]+)"[^>]*>.*?(?:<span[^>]*>)?(?P<title>.*?)(?:</span>)?\s*</a>',
    re.IGNORECASE | re.DOTALL,
)
AUTHOR_RE = re.compile(
    r'<a[^>]+class="authorName"[^>]*>.*?(?:<span[^>]*>)?(?P<author>.*?)(?:</span>)?\s*</a>',
    re.IGNORECASE | re.DOTALL,
)
YEAR_RE = re.compile(r"published\s*(?P<year>\d{3,4})", re.IGNORECASE)
TURKISH_ASCII_MAP = str.maketrans(
    {
        "ç": "c",
        "Ç": "C",
        "ğ": "g",
        "Ğ": "G",
        "ı": "i",
        "İ": "I",
        "ö": "o",
        "Ö": "O",
        "ş": "s",
        "Ş": "S",
        "ü": "u",
        "Ü": "U",
        "â": "a",
        "Â": "A",
        "î": "i",
        "Î": "I",
        "û": "u",
        "Û": "U",
    }
)
TURKISH_ASCII_MAP.update(
    str.maketrans(
        {
            "\u00e7": "c",
            "\u00c7": "C",
            "\u011f": "g",
            "\u011e": "G",
            "\u0131": "i",
            "\u0130": "I",
            "\u00f6": "o",
            "\u00d6": "O",
            "\u015f": "s",
            "\u015e": "S",
            "\u00fc": "u",
            "\u00dc": "U",
            "\u00e2": "a",
            "\u00c2": "A",
            "\u00ee": "i",
            "\u00ce": "I",
            "\u00fb": "u",
            "\u00db": "U",
        }
    )
)
MANUAL_GOODREADS_MATCHES: dict[tuple[str, str], dict[str, str]] = {
    (
        "babalar ve ogullar",
        "ivan turgenyev",
    ): {
        "title": "Babalar ve Ogullar",
        "author": "Ivan Turgenev",
        "url": "https://www.goodreads.com/book/show/21569718-babalar-ve-o-ullar",
        "published_year": "1862",
        "query": "manual_override",
    },
    (
        "gora",
        "rabindranath tagore",
    ): {
        "title": "Gora",
        "author": "Rabindranath Tagore",
        "url": "https://www.goodreads.com/book/show/1268541.Gora",
        "published_year": "1909",
        "query": "manual_override",
    },
    (
        "robinson crusoe",
        "daniel defoe",
    ): {
        "title": "Robinson Crusoe",
        "author": "Daniel Defoe",
        "url": "https://www.goodreads.com/book/show/2932.Robinson_Crusoe",
        "published_year": "1719",
        "query": "manual_override",
    },
    (
        "dede korkut hikayeleri",
        "",
    ): {
        "title": "Dede Korkut Hikayeleri - Kitab-i Dedem Korkut",
        "author": "Anonymous",
        "url": "https://www.goodreads.com/book/show/39893426-dede-korkut-hik-yeleri-kitab--dedem-korkut",
        "published_year": "",
        "query": "manual_override",
    },
    (
        "haldun taner hikayelerinden secmeler",
        "haldun taner",
    ): {
        "title": "Haldun Taner Hikayeleri",
        "author": "Haldun Taner",
        "url": "https://www.goodreads.com/book/show/29619096",
        "published_year": "2018",
        "query": "manual_override",
    },
    (
        "divan siirinden secmeler",
        "",
    ): {
        "title": "Divan Siirinden Secmeler",
        "author": "Erol Battal",
        "url": "https://www.goodreads.com/book/show/23059009-divan-iirinden-se-meler",
        "published_year": "2008",
        "query": "manual_override",
    },
    (
        "halk siirinden secmeler",
        "",
    ): {
        "title": "Halk Siirinden Secmeler",
        "author": "Sennur Sezer",
        "url": "https://www.goodreads.com/book/show/29391340-halk-iirinden-se-meler",
        "published_year": "2009",
        "query": "manual_override",
    },
    (
        "kutadgu bilig den secmeler",
        "yusuf has hacib",
    ): {
        "title": "Kutadgu Biligden Secmeler",
        "author": "Yusuf Has Hacip",
        "url": "https://www.goodreads.com/book/show/67958624-kutadgu-biligden-secmeler",
        "published_year": "",
        "query": "manual_override",
    },
    (
        "mesnevi den secmeler",
        "mevlana",
    ): {
        "title": "Mesnevi'den Secmeler",
        "author": "Mevlana",
        "url": "https://www.goodreads.com/book/show/34814525",
        "published_year": "",
        "query": "manual_override",
    },
    (
        "devlet",
        "platon",
    ): {
        "title": "Devlet",
        "author": "Plato",
        "url": "https://www.goodreads.com/book/show/21566685-devlet",
        "published_year": "",
        "query": "manual_override",
    },
    (
        "seyahatname",
        "evliya celebi",
    ): {
        "title": "Seyahatname'den Secmeler",
        "author": "Evliya Celebi",
        "url": "https://www.goodreads.com/book/show/16248609-seyahatname-den-se-meler",
        "published_year": "",
        "query": "manual_override",
    },
    (
        "esir sehrin insanlari",
        "kemal tahir",
    ): {
        "title": "Esir Şehrin İnsanları",
        "author": "Kemal Tahir",
        "url": "https://www.goodreads.com/book/show/6562214-esir-ehrin-i-nsanlar",
        "published_year": "1956",
        "query": "manual_override",
    },
    (
        "gurbet hikayeleri",
        "refik halit karay",
    ): {
        "title": "Gurbet Hikayeleri",
        "author": "Refik Halid Karay",
        "url": "https://www.goodreads.com/book/show/17840893-gurbet-hikayeleri",
        "published_year": "1940",
        "query": "manual_override",
    },
    (
        "orhan veli kanik in butun siirleri",
        "",
    ): {
        "title": "Butun Siirleri",
        "author": "Orhan Veli Kanik",
        "url": "https://www.goodreads.com/book/show/7667952",
        "published_year": "1951",
        "query": "manual_override",
    },
    (
        "sait faik abasiyanik hikayelerinden secmeler",
        "sait faik abasiyanik",
    ): {
        "title": "Seçme Hikâyeler",
        "author": "Sait Faik Abasıyanık",
        "url": "https://www.goodreads.com/book/show/16095499-se-me-hik-yeler",
        "published_year": "2012",
        "query": "manual_override",
    },
    (
        "ahmet muhip diranas siirlerinden secmeler",
        "ahmet muhip diranas",
    ): {
        "title": "Şiirler",
        "author": "Ahmet Muhip Dıranas",
        "url": "https://www.goodreads.com/book/show/12707874-iirler",
        "published_year": "1974",
        "query": "manual_override",
    },
    (
        "ahmet kutsi tecer siirlerinden secmeler",
        "ahmet kutsi tecer",
    ): {
        "title": "Şiirler",
        "author": "Ahmet Kutsi Tecer",
        "url": "https://www.goodreads.com/book/show/18517041-b-t-n-iirleri",
        "published_year": "",
        "query": "manual_override",
    },
}


# ==================== 2. VERI MODELLERI ====================
@dataclass(slots=True)
class InputBook:
    eser: str
    yazar: str
    milliyet: str
    ilk_basim_yili: str


@dataclass(slots=True)
class SearchCandidate:
    title: str
    author: str
    url: str
    published_year: str
    query: str
    score: float


# ==================== 3. METIN NORMALIZASYONU VE SKORLAMA ====================
def log(message: str) -> None:
    safe_message = message.encode("cp1254", errors="replace").decode("cp1254")
    print(safe_message)


def ascii_fold(value: str) -> str:
    translated = value.translate(TURKISH_ASCII_MAP)
    normalized = unicodedata.normalize("NFKD", translated)
    return normalized.encode("ascii", "ignore").decode("ascii")


def clean_text(value: str) -> str:
    value = html.unescape(value)
    value = TAG_RE.sub(" ", value)
    value = value.replace("\xa0", " ")
    value = WHITESPACE_RE.sub(" ", value)
    return value.strip()


def normalize_for_match(value: str) -> str:
    value = ascii_fold(value).lower()
    value = re.sub(r"[^a-z0-9\s]", " ", value)
    return WHITESPACE_RE.sub(" ", value).strip()


def similarity(left: str, right: str) -> float:
    if not left or not right:
        return 0.0
    return SequenceMatcher(None, left, right).ratio()


def token_overlap(left: str, right: str) -> float:
    left_tokens = set(left.split())
    right_tokens = set(right.split())
    if not left_tokens or not right_tokens:
        return 0.0
    intersection = len(left_tokens & right_tokens)
    return intersection / max(len(left_tokens), len(right_tokens))


def score_candidate(book: InputBook, candidate_title: str, candidate_author: str) -> float:
    expected_title = normalize_for_match(book.eser)
    expected_author = normalize_for_match(book.yazar)
    found_title = normalize_for_match(candidate_title)
    found_author = normalize_for_match(candidate_author)

    title_score = max(similarity(expected_title, found_title), token_overlap(expected_title, found_title))
    author_score = max(similarity(expected_author, found_author), token_overlap(expected_author, found_author))

    score = (title_score * 0.78) + (author_score * 0.22)

    if expected_title and found_title == expected_title:
        score += 0.20
    elif expected_title and expected_title in found_title:
        score += 0.08

    if expected_author:
        if found_author == expected_author:
            score += 0.10
        elif expected_author in found_author or found_author in expected_author:
            score += 0.05
    else:
        score += 0.03

    return min(score, 1.0)


# ==================== 4. MODUL VE EXCEL OKUMA ====================
def load_exporter_module() -> Any:
    if not EXPORTER_MODULE_PATH.exists():
        raise FileNotFoundError(f"Exporter dosyasi bulunamadi: {EXPORTER_MODULE_PATH}")

    spec = importlib.util.spec_from_file_location("goodreads_tr_exporter_local", EXPORTER_MODULE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("Goodreads exporter modulu yuklenemedi.")

    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def read_shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []

    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    values: list[str] = []
    for string_item in root.findall("x:si", XML_NS):
        values.append("".join(string_item.itertext()))
    return values


def find_first_sheet_path(workbook: zipfile.ZipFile) -> str:
    workbook_root = ET.fromstring(workbook.read("xl/workbook.xml"))
    rels_root = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    relationship_map = {
        rel.attrib["Id"]: rel.attrib["Target"] for rel in rels_root.findall("r:Relationship", REL_NS)
    }

    sheet = workbook_root.find("x:sheets/x:sheet", XML_NS)
    if sheet is None:
        raise RuntimeError("Excel icinde sayfa bulunamadi.")

    relation_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
    if not relation_id or relation_id not in relationship_map:
        raise RuntimeError("Excel sayfa iliskisi bulunamadi.")

    target = relationship_map[relation_id]
    return f"xl/{target}" if not target.startswith("xl/") else target


def read_cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t", "")

    if cell_type == "inlineStr":
        inline = cell.find("x:is", XML_NS)
        return "".join(inline.itertext()).strip() if inline is not None else ""

    if cell_type == "s":
        value = cell.findtext("x:v", default="", namespaces=XML_NS)
        if value.isdigit():
            index = int(value)
            if 0 <= index < len(shared_strings):
                return shared_strings[index].strip()
        return ""

    return cell.findtext("x:v", default="", namespaces=XML_NS).strip()


def read_input_books(path: Path) -> list[InputBook]:
    with zipfile.ZipFile(path) as workbook:
        shared_strings = read_shared_strings(workbook)
        sheet_path = find_first_sheet_path(workbook)
        sheet_root = ET.fromstring(workbook.read(sheet_path))

    rows = sheet_root.findall("x:sheetData/x:row", XML_NS)
    if not rows:
        return []

    header = [read_cell_value(cell, shared_strings) for cell in rows[0].findall("x:c", XML_NS)]
    header_map = {name: index for index, name in enumerate(header)}

    def get_value(cells: list[str], key: str) -> str:
        index = header_map.get(key)
        if index is None or index >= len(cells):
            return ""
        return cells[index]

    books: list[InputBook] = []
    for row in rows[1:]:
        values = [read_cell_value(cell, shared_strings) for cell in row.findall("x:c", XML_NS)]
        eser = get_value(values, "Eser")
        if not eser:
            continue
        books.append(
            InputBook(
                eser=eser,
                yazar=get_value(values, "Yazar"),
                milliyet=get_value(values, "Milliyet"),
                ilk_basim_yili=get_value(values, "İlk basım yılı"),
            )
        )
    return books


# ==================== 5. GOODREADS ARAMA VE ESLESTIRME ====================
def build_search_queries(book: InputBook) -> list[str]:
    raw_queries = [
        " ".join(part for part in [book.eser, book.yazar] if part),
        " ".join(part for part in [ascii_fold(book.eser), ascii_fold(book.yazar)] if part),
        book.eser,
        ascii_fold(book.eser),
    ]

    unique: list[str] = []
    seen: set[str] = set()
    for item in raw_queries:
        cleaned = WHITESPACE_RE.sub(" ", item).strip()
        if not cleaned:
            continue
        lowered = cleaned.lower()
        if lowered in seen:
            continue
        seen.add(lowered)
        unique.append(cleaned)
    return unique


def parse_search_candidates(page_html: str, query: str, book: InputBook) -> list[SearchCandidate]:
    candidates: list[SearchCandidate] = []

    for match in SEARCH_ROW_RE.finditer(page_html):
        body = match.group("body")
        title_match = BOOK_TITLE_RE.search(body)
        author_match = AUTHOR_RE.search(body)
        if not title_match:
            continue

        title = clean_text(title_match.group("title"))
        author = clean_text(author_match.group("author")) if author_match else ""
        href = html.unescape(title_match.group("href"))
        year_match = YEAR_RE.search(body)
        published_year = year_match.group("year") if year_match else ""
        url = urljoin(GOODREADS_BASE_URL, href)
        score = score_candidate(book, title, author)

        candidates.append(
            SearchCandidate(
                title=title,
                author=author,
                url=url,
                published_year=published_year,
                query=query,
                score=score,
            )
        )

    deduped: dict[str, SearchCandidate] = {}
    for candidate in candidates:
        current = deduped.get(candidate.url)
        if current is None or candidate.score > current.score:
            deduped[candidate.url] = candidate

    return sorted(deduped.values(), key=lambda item: item.score, reverse=True)


def find_best_goodreads_match(exporter: Any, book: InputBook) -> tuple[SearchCandidate | None, list[SearchCandidate]]:
    manual_key = (normalize_for_match(book.eser), normalize_for_match(book.yazar))
    manual_match = MANUAL_GOODREADS_MATCHES.get(manual_key)
    if manual_match is not None:
        candidate = SearchCandidate(
            title=manual_match["title"],
            author=manual_match["author"],
            url=manual_match["url"],
            published_year=manual_match["published_year"],
            query=manual_match["query"],
            score=1.0,
        )
        return candidate, [candidate]

    all_candidates: dict[str, SearchCandidate] = {}

    for query in build_search_queries(book):
        search_url = f"{GOODREADS_BASE_URL}/search?{urlencode({'q': query, 'search_type': 'books'})}"
        page_html = ""
        last_error: Exception | None = None
        for _ in range(3):
            try:
                page_html = exporter.request_text("GET", search_url)
                last_error = None
                break
            except Exception as exc:
                last_error = exc
                exporter.reset_transport()
                exporter._sleep(1.0)
        if last_error is not None:
            raise last_error
        for candidate in parse_search_candidates(page_html, query, book):
            current = all_candidates.get(candidate.url)
            if current is None or candidate.score > current.score:
                all_candidates[candidate.url] = candidate

        ranked = sorted(all_candidates.values(), key=lambda item: item.score, reverse=True)
        if ranked and ranked[0].score >= 0.92:
            return ranked[0], ranked

    ranked = sorted(all_candidates.values(), key=lambda item: item.score, reverse=True)
    if not ranked:
        return None, []

    best = ranked[0]
    second = ranked[1] if len(ranked) > 1 else None
    if best.score >= 0.75 and (second is None or (best.score - second.score) >= 0.03):
        return best, ranked

    title_exact = normalize_for_match(best.title) == normalize_for_match(book.eser)
    author_ok = not book.yazar or similarity(normalize_for_match(best.author), normalize_for_match(book.yazar)) >= 0.55
    if best.score >= 0.68 and title_exact and author_ok:
        return best, ranked

    return None, ranked


# ==================== 6. EXCEL YAZMA YARDIMCILARI ====================
def column_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def inline_string_cell(cell_ref: str, value: str) -> str:
    safe_value = escape(value)
    preserve = ' xml:space="preserve"' if value != value.strip() else ""
    return f'<c r="{cell_ref}" t="inlineStr"><is><t{preserve}>{safe_value}</t></is></c>'


def build_sheet_xml(headers: list[str], rows: list[list[str]]) -> str:
    xml_rows: list[str] = []
    all_rows = [headers] + rows
    for row_index, row_values in enumerate(all_rows, start=1):
        cells = []
        for column_index, value in enumerate(row_values, start=1):
            cell_ref = f"{column_letter(column_index)}{row_index}"
            cells.append(inline_string_cell(cell_ref, str(value)))
        xml_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(xml_rows)}</sheetData>"
        "</worksheet>"
    )


def write_xlsx(output_path: Path, sheets: list[tuple[str, list[str], list[list[str]]]]) -> None:
    workbook_sheets = []
    workbook_rels = []
    content_type_overrides = []

    for index, (sheet_name, _, _) in enumerate(sheets, start=1):
        relationship_id = f"rId{index}"
        workbook_sheets.append(
            f'<sheet name="{escape(sheet_name[:31])}" sheetId="{index}" '
            f'r:id="{relationship_id}"/>'
        )
        workbook_rels.append(
            f'<Relationship Id="{relationship_id}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )
        content_type_overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    styles_relationship_id = f"rId{len(sheets) + 1}"
    workbook_rels.append(
        f'<Relationship Id="{styles_relationship_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f"<sheets>{''.join(workbook_sheets)}</sheets>"
        "</workbook>"
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        f"{''.join(workbook_rels)}"
        "</Relationships>"
    )

    root_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        f"{''.join(content_type_overrides)}"
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        "</Types>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1"><xf/></cellStyleXfs>'
        '<cellXfs count="1"><xf xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types_xml)
        workbook.writestr("_rels/.rels", root_rels_xml)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        workbook.writestr("xl/styles.xml", styles_xml)

        for index, (_, headers, rows) in enumerate(sheets, start=1):
            workbook.writestr(f"xl/worksheets/sheet{index}.xml", build_sheet_xml(headers, rows))


# ==================== 7. YORUM CEKME ====================
def fetch_reviews_for_match(exporter: Any, match: SearchCandidate) -> tuple[Any, list[dict[str, Any]], dict[str, Any]]:
    last_error: Exception | None = None
    for attempt in range(3):
        try:
            metadata = exporter.fetch_book_metadata(match.url)
            reviews, audit = exporter.collect_reviews_for_book(metadata)
            return metadata, reviews, audit
        except Exception as exc:
            last_error = exc
            exporter.reset_transport()
            exporter._sleep(min(4.0, 1.5 + attempt))
    if last_error is None:
        raise RuntimeError("Goodreads veri cekimi bilinmeyen bir nedenle basarisiz oldu.")
    raise last_error


# ==================== 8. CALISTIRMA AKISI ====================
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--max-books", type=int, default=None, help="Sadece test icin ilk N kitabi isle.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    base_dir = Path(__file__).resolve().parent
    input_path = base_dir / INPUT_XLSX_NAME
    output_path = base_dir / OUTPUT_XLSX_NAME

    if not input_path.exists():
        raise FileNotFoundError(f"Girdi Excel dosyasi bulunamadi: {input_path}")

    module = load_exporter_module()
    config = module.ExportConfig(
        list_url="",
        review_language_code="tr",
        output_dir=str(base_dir),
        max_books=None,
        request_delay_seconds=0.05,
        http_retry_limit=5,
        graphql_page_retry_limit=6,
        book_retry_limit=5,
        review_page_size=30,
        request_timeout_seconds=30,
    )
    exporter = module.GoodreadsTurkishReviewExporter(config)

    input_books = read_input_books(input_path)
    if args.max_books is not None:
        input_books = input_books[: args.max_books]

    review_rows: list[list[str]] = []
    book_rows: list[list[str]] = []
    error_rows: list[list[str]] = []

    total_reviews = 0
    matched_books = 0

    for index, book in enumerate(input_books, start=1):
        log(f"[{index}/{len(input_books)}] Araniyor: {book.eser} / {book.yazar}")

        try:
            match, ranked_candidates = find_best_goodreads_match(exporter, book)
        except Exception as exc:
            error_rows.append(
                [book.eser, book.yazar, book.ilk_basim_yili, "arama", str(exc), "", ""]
            )
            book_rows.append(
                ["arama_hatasi", book.eser, book.yazar, book.ilk_basim_yili, "", "", "", "", "", "", "", str(exc)]
            )
            log(f"  [!] Arama hatasi: {exc}")
            continue

        if match is None:
            candidate_hint = ranked_candidates[0].url if ranked_candidates else ""
            error_rows.append(
                [
                    book.eser,
                    book.yazar,
                    book.ilk_basim_yili,
                    "eslestirme",
                    "Guvenilir Goodreads eslesmesi bulunamadi.",
                    ranked_candidates[0].query if ranked_candidates else "",
                    candidate_hint,
                ]
            )
            book_rows.append(
                [
                    "eslesme_yok",
                    book.eser,
                    book.yazar,
                    book.ilk_basim_yili,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "Guvenilir Goodreads eslesmesi bulunamadi.",
                ]
            )
            log("  [!] Guvenilir eslesme bulunamadi.")
            continue

        log(f"  -> Eslesme: {match.title} / {match.author} (score={match.score:.3f})")

        try:
            metadata, reviews, audit = fetch_reviews_for_match(exporter, match)
        except Exception as exc:
            error_rows.append(
                [book.eser, book.yazar, book.ilk_basim_yili, "yorumlar", str(exc), match.query, match.url]
            )
            book_rows.append(
                [
                    "yorum_hatasi",
                    book.eser,
                    book.yazar,
                    book.ilk_basim_yili,
                    match.title,
                    match.author,
                    match.published_year,
                    match.url,
                    f"{match.score:.3f}",
                    "",
                    "",
                    str(exc),
                ]
            )
            log(f"  [!] Yorum hatasi: {exc}")
            continue

        if not audit.get("success"):
            error_rows.append(
                [
                    book.eser,
                    book.yazar,
                    book.ilk_basim_yili,
                    "yorumlar",
                    str(audit.get("last_error") or "Goodreads yorumlari eksiksiz cekilemedi."),
                    match.query,
                    match.url,
                ]
            )
            book_rows.append(
                [
                    "yorum_hatasi",
                    book.eser,
                    book.yazar,
                    book.ilk_basim_yili,
                    metadata.title,
                    metadata.author,
                    str(metadata.publication_year or ""),
                    metadata.source_url,
                    f"{match.score:.3f}",
                    str(len(reviews)),
                    str(audit.get("stats_language_reviews_count") or 0),
                    str(audit.get("last_error") or ""),
                ]
            )
            log(f"  [!] Goodreads tamamlama hatasi: {audit.get('last_error')}")
            continue

        matched_books += 1
        total_reviews += len(reviews)
        stats_count = audit.get("stats_language_reviews_count") or 0
        verified_count = audit.get("graphql_total_count")
        displayed_count = verified_count if verified_count is not None else stats_count
        note = ""
        if verified_count is not None and stats_count != verified_count:
            note = (
                f"Goodreads sayfa istatistigi {stats_count}, "
                f"dogrulanan GraphQL toplam {verified_count}."
            )
        book_rows.append(
            [
                "tamam",
                book.eser,
                book.yazar,
                book.ilk_basim_yili,
                metadata.title,
                metadata.author,
                str(metadata.publication_year or ""),
                metadata.source_url,
                f"{match.score:.3f}",
                str(len(reviews)),
                str(displayed_count),
                note,
            ]
        )

        for review in reviews:
            review_rows.append(
                [
                    book.eser,
                    book.yazar,
                    book.ilk_basim_yili,
                    metadata.title,
                    metadata.author,
                    str(metadata.publication_year or ""),
                    metadata.source_url,
                    str(review.get("yorumcu_adi") or ""),
                    str(review.get("puan") or ""),
                    str(review.get("yorum_tarihi") or ""),
                    str(review.get("yorum_dili") or ""),
                    str(review.get("yorum") or ""),
                    str(review.get("review_id") or ""),
                ]
            )

        log(f"  yorumlar yazilmaya hazir: {len(reviews)}")

    summary_rows = [
        ["Toplam MEB kitabi", str(len(input_books))],
        ["Basarili eslesen kitap", str(matched_books)],
        ["Toplam cekilen Turkce yorum", str(total_reviews)],
        ["Hata/eslesemeyen kitap", str(len(error_rows))],
        ["Cikti Excel", str(output_path)],
    ]

    sheets = [
        (
            "Yorumlar",
            [
                "MEB eser",
                "MEB yazar",
                "MEB ilk basim yili",
                "Goodreads eser",
                "Goodreads yazar",
                "Goodreads yayin yili",
                "Goodreads URL",
                "Yorumcu",
                "Puan",
                "Yorum tarihi",
                "Yorum dili",
                "Yorum",
                "Review ID",
            ],
            review_rows,
        ),
        (
            "Kitaplar",
            [
                "Durum",
                "MEB eser",
                "MEB yazar",
                "MEB ilk basim yili",
                "Goodreads eser",
                "Goodreads yazar",
                "Goodreads yayin yili",
                "Goodreads URL",
                "Arama skoru",
                "Cekilen yorum sayisi",
                "Goodreads stats tr yorum",
                "Not",
            ],
            book_rows,
        ),
        (
            "Hatalar",
            ["MEB eser", "MEB yazar", "MEB ilk basim yili", "Asama", "Hata", "Arama sorgusu", "Aday URL"],
            error_rows if error_rows else [["", "", "", "", "", "", ""]],
        ),
        ("Ozet", ["Alan", "Deger"], summary_rows),
    ]
    write_xlsx(output_path, sheets)

    log("")
    log(f"Excel kaydedildi: {output_path}")
    log(f"Basarili kitap: {matched_books}/{len(input_books)}")
    log(f"Toplam Turkce yorum: {total_reviews}")
    log(f"Hata sayisi: {len(error_rows)}")


if __name__ == "__main__":
    main()
