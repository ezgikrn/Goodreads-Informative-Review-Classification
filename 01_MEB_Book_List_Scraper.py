from __future__ import annotations

import re
import zipfile
from html import unescape
from html.parser import HTMLParser
from pathlib import Path
from urllib.request import Request, urlopen
from xml.sax.saxutils import escape


# ==================== 1. AYARLAR ====================
WIKIPEDIA_URL = (
    "https://tr.wikipedia.org/wiki/"
    "MEB_100_temel_eser_listesi_(orta%C3%B6%C4%9Fretim)"
)
OUTPUT_NAME = "MEB 100 temel eser listesi (ortaöğretim).xlsx"
OUTPUT_DIR = Path(__file__).resolve().parent
OUTPUT_PATH = OUTPUT_DIR / OUTPUT_NAME


# ==================== 2. WIKIPEDIA TABLO PARSER ====================
class WikiTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.tables: list[list[list[str]]] = []
        self._table_depth = 0
        self._in_target_table = False
        self._current_table: list[list[str]] = []
        self._current_row: list[str] | None = None
        self._current_cell: list[str] | None = None
        self._capture_cell = False

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr_map = dict(attrs)
        class_name = attr_map.get("class", "")

        if tag == "table" and "wikitable" in class_name:
            self._table_depth += 1
            if not self._in_target_table:
                self._in_target_table = True
                self._current_table = []
            return

        if not self._in_target_table:
            return

        if tag == "tr":
            self._current_row = []
        elif tag in {"td", "th"}:
            self._capture_cell = True
            self._current_cell = []
        elif tag == "br" and self._capture_cell and self._current_cell is not None:
            self._current_cell.append("\n")
        elif tag == "img" and self._capture_cell and self._current_cell is not None:
            alt_text = normalize_country(attr_map.get("alt", "") or "")
            if alt_text:
                self._current_cell.append(alt_text)

    def handle_endtag(self, tag: str) -> None:
        if not self._in_target_table:
            return

        if tag in {"td", "th"} and self._capture_cell and self._current_cell is not None:
            text = normalize_text("".join(self._current_cell))
            self._current_row = self._current_row or []
            self._current_row.append(text)
            self._capture_cell = False
            self._current_cell = None
        elif tag == "tr" and self._current_row is not None:
            if any(cell.strip() for cell in self._current_row):
                self._current_table.append(self._current_row)
            self._current_row = None
        elif tag == "table":
            self._table_depth -= 1
            if self._table_depth == 0:
                self.tables.append(self._current_table)
                self._current_table = []
                self._in_target_table = False

    def handle_data(self, data: str) -> None:
        if self._capture_cell and self._current_cell is not None:
            self._current_cell.append(data)


# ==================== 3. METIN TEMIZLEME ====================
def normalize_text(value: str) -> str:
    value = unescape(value)
    value = value.replace("\xa0", " ")
    value = re.sub(r"\[[^\]]+\]", "", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


# ==================== 4. HTML CEKME VE TABLO SECME ====================
def fetch_html(url: str) -> str:
    request = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(request, timeout=30) as response:
        return response.read().decode("utf-8")


def find_target_table(html: str) -> list[list[str]]:
    parser = WikiTableParser()
    parser.feed(html)

    for table in parser.tables:
        if not table:
            continue
        header = table[0]
        if {"Eser", "Yazar", "Milliyet"}.issubset(set(header)):
            return table

    raise RuntimeError("Eserler tablosu bulunamadı.")


# ==================== 5. ESER KAYITLARINI OLUSTURMA ====================
def build_records(table: list[list[str]]) -> list[dict[str, str]]:
    header = table[0]
    rows = table[1:]
    normalized_header = [normalize_text(cell) for cell in header]
    expected_column_count = len(normalized_header)

    records: list[dict[str, str]] = []
    for row in rows:
        padded_row = row[:expected_column_count] + [""] * max(0, expected_column_count - len(row))
        record = dict(zip(normalized_header, padded_row))

        milliyet = normalize_country(record.get("Milliyet", ""))

        aciklama = normalize_text(record.get("Açıklama", ""))
        if "çıkarıldı" in aciklama.lower():
            continue

        ilk_basim = extract_year(record.get("İlk basım", ""))
        eser = record.get("Eser", "")
        yazar = record.get("Yazar", "")
        if not eser:
            continue

        records.append(
            {
                "Eser": eser,
                "Yazar": yazar,
                "Milliyet": milliyet,
                "İlk basım yılı": ilk_basim,
            }
        )

    return records


def extract_year(value: str) -> str:
    match = re.search(r"\b(\d{3,4})\b", value)
    return match.group(1) if match else ""


def normalize_country(value: str) -> str:
    value = normalize_text(value)
    value = re.sub(r"^Image:\s*", "", value, flags=re.IGNORECASE)
    return value


# ==================== 6. EXCEL XML YARDIMCILARI ====================
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


def build_sheet_xml(headers: list[str], records: list[dict[str, str]]) -> str:
    rows_xml: list[str] = []
    table_rows = [headers] + [[record.get(header, "") for header in headers] for record in records]

    for row_index, row_values in enumerate(table_rows, start=1):
        cells_xml = []
        for column_index, value in enumerate(row_values, start=1):
            cell_ref = f"{column_letter(column_index)}{row_index}"
            cells_xml.append(inline_string_cell(cell_ref, value))
        rows_xml.append(f'<row r="{row_index}">{"".join(cells_xml)}</row>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(rows_xml)}</sheetData>"
        "</worksheet>"
    )


# ==================== 7. EXCEL DOSYASINI YAZMA ====================
def write_xlsx(output_path: Path, records: list[dict[str, str]]) -> None:
    headers = ["Eser", "Yazar", "Milliyet", "İlk basım yılı"]
    sheet_xml = build_sheet_xml(headers, records)

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Eserler" sheetId="1" r:id="rId1"/></sheets>'
        "</workbook>"
    )

    workbook_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
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
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
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
        workbook.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        workbook.writestr("xl/styles.xml", styles_xml)


# ==================== 8. CALISTIRMA AKISI ====================
def main() -> None:
    html = fetch_html(WIKIPEDIA_URL)
    table = find_target_table(html)
    records = build_records(table)
    write_xlsx(OUTPUT_PATH, records)
    print(f"{len(records)} kayıt yazıldı: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
