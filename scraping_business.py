"""Pencarian direktori bisnis lokal dari OpenStreetMap Overpass API."""

from __future__ import annotations

import argparse
import re
import threading
import time
from pathlib import Path
from typing import Callable

import requests
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


OVERPASS_URL = "https://overpass-api.de/api/interpreter"
USER_AGENT = "ScraperDataLocal/1.0 (local business directory research)"
AREA_ID_OFFSET = 3_600_000_000
SOUTH_SUMATRA = "Sumatera Selatan"
ProgressCallback = Callable[[str, int], None]

BUSINESS_CATEGORIES = [
    {"value": "retail", "label": "Toko dan ritel"},
    {"value": "food", "label": "Makanan dan minuman"},
    {"value": "health", "label": "Kesehatan dan farmasi"},
    {"value": "professional", "label": "Jasa profesional dan kantor"},
    {"value": "finance", "label": "Keuangan dan perbankan"},
    {"value": "automotive", "label": "Otomotif dan transportasi"},
    {"value": "accommodation", "label": "Akomodasi dan pariwisata"},
    {"value": "education", "label": "Pendidikan dan pelatihan"},
    {"value": "construction", "label": "Konstruksi dan perlengkapan"},
    {"value": "agriculture", "label": "Pertanian dan peternakan"},
    {"value": "other", "label": "Kerajinan dan industri"},
]

CATEGORY_SELECTORS: dict[str, list[tuple[str, str | None]]] = {
    "retail": [("shop", None)],
    "food": [
        ("amenity", "restaurant|cafe|fast_food|bar|pub|food_court|ice_cream"),
        ("shop", "bakery|beverages|butcher|cheese|chocolate|coffee|confectionery|deli|seafood|tea"),
    ],
    "health": [
        ("amenity", "clinic|doctors|dentist|pharmacy|hospital|veterinary"),
        ("healthcare", None), ("shop", "medical|optician|hearing_aids"),
    ],
    "professional": [("office", None)],
    "finance": [("amenity", "bank|atm|bureau_de_change|payment_centre"), ("office", "financial|insurance")],
    "automotive": [
        ("shop", "car|car_repair|car_parts|motorcycle|motorcycle_repair|tyres|rental"),
        ("amenity", "fuel|car_wash|charging_station|vehicle_inspection"),
    ],
    "accommodation": [
        ("tourism", "hotel|guest_house|hostel|motel|resort|apartment|camp_site|chalet|attraction"),
        ("leisure", "amusement_arcade|escape_game|water_park"),
    ],
    "education": [
        ("amenity", "school|college|university|kindergarten|language_school|driving_school|training"),
        ("office", "educational_institution"),
    ],
    "construction": [
        ("office", "architect|construction_company|estate_agent|surveyor"),
        ("shop", "building_materials|hardware|trade|paint|doors|flooring|tiles"),
        ("craft", "builder|carpenter|electrician|glaziery|plumber|roofer"),
    ],
    "agriculture": [
        ("shop", "agrarian|farm|garden_centre|seeds|agricultural_machinery"),
        ("office", "agricultural"), ("craft", "beekeeper"),
    ],
    "other": [("craft", None), ("industrial", None)],
}

XLSX_COLUMNS = [
    "Nama Usaha", "Kategori", "Kabupaten/Kota", "Kecamatan", "Alamat",
    "Latitude", "Longitude", "Link Peta", "Telepon", "Email", "Situs Web",
    "Skor Ulasan", "Jumlah Ulasan", "Sumber Data",
]

_overpass_lock = threading.Lock()
_cache_lock = threading.Lock()
_area_cache: dict[str, tuple[float, list[dict[str, str]]]] = {}
_known_areas: dict[str, tuple[float, dict[str, str]]] = {}
_business_cache: dict[str, tuple[float, list[list[object]]]] = {}
AREA_CACHE_TTL = 86_400
BUSINESS_CACHE_TTL = 3_600
BUSINESS_CACHE_MAX_ITEMS = 128
MIN_OVERPASS_INTERVAL = 1.0
_last_overpass_request_at = 0.0


class OverpassError(RuntimeError):
    """Kesalahan kueri ke layanan Overpass."""


def _clean(value: object) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def _xlsx_safe(value: object) -> object:
    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def _fetch_overpass(query: str, timeout: int = 90) -> dict:
    global _last_overpass_request_at
    headers = {"User-Agent": USER_AGENT, "Accept": "application/json"}
    try:
        # Serialize requests: public Overpass asks clients not to run queries in parallel.
        with _overpass_lock:
            pause = MIN_OVERPASS_INTERVAL - (time.monotonic() - _last_overpass_request_at)
            if pause > 0:
                time.sleep(pause)
            _last_overpass_request_at = time.monotonic()
            response = requests.post(
                OVERPASS_URL,
                data={"data": query},
                headers=headers,
                timeout=timeout,
            )
    except requests.Timeout:
        raise OverpassError("Pencarian OpenStreetMap melewati batas waktu. Persempit wilayah atau coba lagi nanti.") from None
    except requests.RequestException as error:
        raise OverpassError(f"Layanan OpenStreetMap tidak dapat dijangkau ({type(error).__name__}).") from None

    if response.status_code in {429, 502, 503, 504}:
        raise OverpassError(
            f"Layanan OpenStreetMap sedang sibuk (HTTP {response.status_code}). Tunggu sebentar sebelum mencoba lagi."
        )
    try:
        response.raise_for_status()
        payload = response.json()
    except (requests.RequestException, ValueError):
        raise OverpassError("Respons layanan OpenStreetMap tidak dapat dibaca.") from None
    if not isinstance(payload, dict) or not isinstance(payload.get("elements"), list):
        raise OverpassError("Format respons layanan OpenStreetMap tidak sesuai.")
    return payload


def _area_id_from_relation(relation_id: object) -> str:
    try:
        relation_number = int(relation_id)
    except (TypeError, ValueError):
        return ""
    return str(AREA_ID_OFFSET + relation_number)


def get_known_area(area_id: str) -> dict[str, str] | None:
    """Return an area previously discovered beneath the South Sumatra boundary."""
    now = time.monotonic()
    with _cache_lock:
        cached = _known_areas.get(str(area_id))
        if not cached:
            return None
        expires_at, area = cached
        if now >= expires_at:
            _known_areas.pop(str(area_id), None)
            return None
        return dict(area)


def fetch_admin_areas(level: str, parent_area_id: str = "") -> list[dict[str, str]]:
    """Return kabupaten/kota in Sumsel, or kecamatan inside one of them."""
    if level not in {"regency", "district"}:
        raise ValueError("Tingkat wilayah tidak dikenal.")

    if level == "regency":
        cache_key = "sumsel:regency"
        query = f"""[out:json][timeout:25];
rel["boundary"="administrative"]["admin_level"="4"]["name"="{SOUTH_SUMATRA}"];
map_to_area->.sumsel;
rel(area.sumsel)["boundary"="administrative"]["admin_level"="5"];
out tags;
"""
    else:
        if not str(parent_area_id).isdigit():
            raise ValueError("Pilih kabupaten atau kota terlebih dahulu.")
        parent_number = int(parent_area_id)
        parent_area = get_known_area(str(parent_area_id))
        if not parent_area or parent_area["admin_level"] != "5":
            raise ValueError("Pilih kabupaten atau kota Sumatera Selatan yang tersedia.")
        cache_key = f"sumsel:district:{parent_number}"
        query = f"""[out:json][timeout:25];
area({parent_number})->.parent;
rel(area.parent)["boundary"="administrative"]["admin_level"="6"];
out tags;
"""

    now = time.monotonic()
    with _cache_lock:
        cached = _area_cache.get(cache_key)
        if cached and now - cached[0] < AREA_CACHE_TTL:
            return [dict(item) for item in cached[1]]

    payload = _fetch_overpass(query, timeout=45)
    areas = []
    for element in payload["elements"]:
        tags = element.get("tags") or {}
        name = _clean(tags.get("name") or tags.get("name:id"))
        area_id = _area_id_from_relation(element.get("id"))
        if not name or not area_id:
            continue
        areas.append({
            "id": area_id,
            "name": name,
            "admin_level": str(tags.get("admin_level") or ("5" if level == "regency" else "6")),
        })

    # Keep the results deterministic and remove same-name boundary duplicates.
    unique = {item["id"]: item for item in areas}
    ordered = sorted(unique.values(), key=lambda item: item["name"].casefold())
    with _cache_lock:
        cached_at = time.monotonic()
        expires_at = cached_at + AREA_CACHE_TTL
        _area_cache[cache_key] = (cached_at, ordered)
        for item in ordered:
            _known_areas[item["id"]] = (expires_at, {
                **item,
                "parent_id": str(parent_area_id) if level == "district" else "",
            })
    return [dict(item) for item in ordered]


def _overpass_regex(value: str) -> str:
    # Escape the regex first, then encode its backslashes for the QL string literal.
    value = value.replace("\r", " ").replace("\n", " ")
    regex_chars = set(r"\.^$*+?{}[]|()")
    escaped_regex = "".join(f"\\{char}" if char in regex_chars else char for char in value)
    return escaped_regex.replace("\\", "\\\\").replace('"', '\\"')


def _build_business_query(area_id: str, category: str, keyword: str, max_results: int) -> str:
    if category not in CATEGORY_SELECTORS:
        raise ValueError("Pilih kategori bisnis yang tersedia.")
    if not str(area_id).isdigit() or int(area_id) <= AREA_ID_OFFSET:
        raise ValueError("Pilih wilayah kabupaten/kota atau kecamatan terlebih dahulu.")

    selectors = CATEGORY_SELECTORS[category]
    keyword_filter = ""
    term = _clean(keyword)[:80]
    if term:
        keyword_filter = f'[name~"{_overpass_regex(term)}",i]'

    clauses = []
    for key, values in selectors:
        if values:
            clauses.append(f'nwr(area.searchArea)["{key}"~"^({values})$"]["name"]{keyword_filter};')
        else:
            clauses.append(f'nwr(area.searchArea)["{key}"]["name"]{keyword_filter};')

    return "\n".join([
        "[out:json][timeout:45];",
        f"area({int(area_id)})->.searchArea;",
        "(",
        *clauses,
        ");",
        f"out tags center {max_results};",
    ])


def _first_tag(tags: dict, *keys: str) -> str:
    for key in keys:
        value = _clean(tags.get(key))
        if value:
            return value
    return ""


def _business_category(tags: dict) -> str:
    category_keys = ("shop", "amenity", "office", "craft", "industrial", "tourism", "healthcare", "leisure")
    return ", ".join(f"{key}={_clean(tags[key])}" for key in category_keys if _clean(tags.get(key)))


def _address(tags: dict) -> str:
    full = _first_tag(tags, "addr:full")
    if full:
        return full
    address_parts = [
        _first_tag(tags, "addr:housenumber"),
        _first_tag(tags, "addr:street", "addr:place"),
        _first_tag(tags, "addr:suburb", "addr:neighbourhood", "addr:village"),
        _first_tag(tags, "addr:district"),
        _first_tag(tags, "addr:city", "addr:town", "addr:county"),
        _first_tag(tags, "addr:province"),
        _first_tag(tags, "addr:postcode"),
    ]
    return ", ".join(dict.fromkeys(part for part in address_parts if part))


def _coordinates(element: dict) -> tuple[float | None, float | None]:
    if element.get("type") == "node":
        lat, lon = element.get("lat"), element.get("lon")
    else:
        center = element.get("center") or {}
        lat, lon = center.get("lat"), center.get("lon")
    try:
        return float(lat), float(lon)
    except (TypeError, ValueError):
        return None, None


def _element_source_url(element_type: object, element_id: object) -> str:
    kind = str(element_type or "")
    if kind not in {"node", "way", "relation"} or not str(element_id).isdigit():
        return ""
    return f"https://www.openstreetmap.org/{kind}/{element_id}"


def _business_row(
    element: dict,
    scope_label: str,
    regency_label: str = "",
    district_label: str = "",
) -> list[object] | None:
    tags = element.get("tags") or {}
    name = _first_tag(tags, "name", "brand", "operator")
    if not name:
        return None
    latitude, longitude = _coordinates(element)
    map_url = ""
    if latitude is not None and longitude is not None:
        map_url = f"https://www.openstreetmap.org/?mlat={latitude:.6f}&mlon={longitude:.6f}#map=18/{latitude:.6f}/{longitude:.6f}"
    review_score = _first_tag(tags, "review:rating", "rating")
    review_count = _first_tag(tags, "review:count", "rating_count")

    return [
        name,
        _business_category(tags),
        _first_tag(tags, "addr:city", "addr:town", "addr:county") or regency_label or scope_label,
        _first_tag(tags, "addr:district", "addr:subdistrict") or district_label,
        _address(tags),
        latitude if latitude is not None else "",
        longitude if longitude is not None else "",
        map_url,
        _first_tag(tags, "contact:phone", "phone", "telephone", "contact:mobile", "mobile"),
        _first_tag(tags, "contact:email", "email"),
        _first_tag(tags, "contact:website", "website", "url"),
        review_score,
        review_count,
        _element_source_url(element.get("type"), element.get("id")),
    ]


def search_businesses(
    area_id: str,
    scope_label: str,
    category: str,
    keyword: str = "",
    max_results: int = 500,
    progress: ProgressCallback | None = None,
    regency_label: str = "",
    district_label: str = "",
) -> list[list[object]]:
    if not 1 <= max_results <= 2000:
        raise ValueError("Batas hasil harus antara 1 dan 2.000.")
    query = _build_business_query(area_id, category, keyword, max_results)
    cache_key = f"{area_id}|{category}|{_clean(keyword).casefold()}|{max_results}|{regency_label}|{district_label}"
    now = time.monotonic()
    with _cache_lock:
        cached = _business_cache.get(cache_key)
        if cached and now - cached[0] < BUSINESS_CACHE_TTL:
            rows = [list(row) for row in cached[1]]
            if progress:
                progress(f"Menggunakan cache OpenStreetMap ({len(rows)} usaha)", 95)
            return rows

    if progress:
        progress(f"Mencari {dict((x['value'], x['label']) for x in BUSINESS_CATEGORIES)[category]} di {scope_label}", 15)
    payload = _fetch_overpass(query)
    rows: list[list[object]] = []
    seen: set[tuple] = set()
    for element in payload["elements"]:
        row = _business_row(element, scope_label, regency_label, district_label)
        if row is None:
            continue
        latitude, longitude = row[5], row[6]
        dedupe = (
            _clean(row[0]).casefold(),
            round(latitude, 5) if isinstance(latitude, float) else "",
            round(longitude, 5) if isinstance(longitude, float) else "",
        )
        if dedupe in seen:
            continue
        seen.add(dedupe)
        rows.append(row)
        if len(rows) >= max_results:
            break

    rows.sort(key=lambda row: str(row[0]).casefold())
    with _cache_lock:
        if cache_key not in _business_cache and len(_business_cache) >= BUSINESS_CACHE_MAX_ITEMS:
            oldest_key = min(_business_cache, key=lambda key: _business_cache[key][0])
            _business_cache.pop(oldest_key, None)
        _business_cache[cache_key] = (time.monotonic(), [list(row) for row in rows])
    if progress:
        progress(f"Menyiapkan {len(rows)} hasil untuk file Excel", 95)
    return rows


def save_xlsx(rows: list[list[object]], output_path: Path, metadata: dict[str, object] | None = None) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Direktori Bisnis"
    sheet.sheet_view.showGridLines = False
    sheet.append(XLSX_COLUMNS)
    for row in rows:
        sheet.append([_xlsx_safe(value) for value in row])

    header_fill = PatternFill(fill_type="solid", fgColor="17365D")
    for cell in sheet[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:N{max(len(rows) + 1, 1)}"
    sheet.row_dimensions[1].height = 32
    for column, width in zip("ABCDEFGHIJKLMN", [32, 30, 28, 24, 55, 14, 14, 52, 23, 32, 42, 16, 18, 45]):
        sheet.column_dimensions[column].width = width
    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
        for col_index in (8, 11, 14):
            cell = row[col_index - 1]
            if cell.value:
                cell.hyperlink = str(cell.value)
                cell.style = "Hyperlink"
    if rows:
        table = Table(displayName="DirektoriBisnisOSM", ref=f"A1:N{len(rows) + 1}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
            showRowStripes=True, showColumnStripes=False,
        )
        sheet.add_table(table)

    info = workbook.create_sheet("Sumber dan Cakupan")
    info.append(["Informasi", "Nilai"])
    for key, value in (metadata or {}).items():
        info.append([_xlsx_safe(key), _xlsx_safe(value)])
    info.append(["Sumber data", "OpenStreetMap contributors"])
    info.append(["Lisensi", "Open Database License (ODbL)"])
    attribution_row = info.max_row + 1
    info.append(["Atribusi", "© OpenStreetMap contributors"])
    info.append(["Catatan rating", "Skor dan jumlah ulasan hanya tersedia jika dicatat secara eksplisit pada data sumber; OpenStreetMap tidak menyediakan rating ulasan umum."])
    for cell in info[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
    info.column_dimensions["A"].width = 28
    info.column_dimensions["B"].width = 110
    info.freeze_panes = "A2"
    attribution_cell = info.cell(row=attribution_row, column=2)
    attribution_cell.hyperlink = "https://www.openstreetmap.org/copyright"
    attribution_cell.style = "Hyperlink"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Cari direktori bisnis Sumatera Selatan dari OpenStreetMap.")
    parser.add_argument("--area-id", required=True, help="ID area OSM kabupaten/kota atau kecamatan.")
    parser.add_argument("--area-label", required=True, help="Nama cakupan wilayah.")
    parser.add_argument("--category", choices=[item["value"] for item in BUSINESS_CATEGORIES], required=True)
    parser.add_argument("--keyword", default="", help="Kata yang dicari pada nama usaha.")
    parser.add_argument("--max-results", type=int, default=500)
    parser.add_argument("--output", type=Path, default=Path("direktori_bisnis_sumsel.xlsx"))
    args = parser.parse_args()
    rows = search_businesses(args.area_id, args.area_label, args.category, args.keyword, args.max_results,
                             progress=lambda message, _: print(message))
    save_xlsx(rows, args.output, {"Cakupan wilayah": args.area_label, "Kategori": args.category,
                                  "Kata kunci": args.keyword or "(semua nama dalam kategori)", "Jumlah usaha": len(rows)})
    print(f"Usaha ditemukan: {len(rows)}; file: {args.output.resolve()}")


if __name__ == "__main__":
    main()
