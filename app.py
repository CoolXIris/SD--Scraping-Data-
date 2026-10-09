"""Aplikasi Flask lokal untuk scraper anime, lowongan, BPS, dan direktori bisnis."""

from __future__ import annotations

import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path

import requests
from flask import Flask, jsonify, render_template, request, send_file

import scraping
import scraping_karirhub
import scraping_bps
import scraping_business


ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "outputs"
WEEKDAYS = {"Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"}
SOURCE_VALUES = {"all", "top", "glints", "jobstreet", "other"}
ALLOWED_STATUS = {"ongoing", "complete", "both"}

app = Flask(__name__, template_folder=str(ROOT / "templates"), static_folder=str(ROOT / "static"))
executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="scraper")
jobs: dict[str, dict] = {}
jobs_lock = threading.Lock()
options_lock = threading.Lock()
options_cache: dict | None = None
options_cached_at = 0.0
OPTIONS_TTL_SECONDS = 1800


def load_options(force: bool = False) -> dict:
    global options_cache, options_cached_at
    with options_lock:
        if options_cache is not None and not force and time.monotonic() - options_cached_at < OPTIONS_TTL_SECONDS:
            return options_cache

        warnings: list[str] = []
        try:
            genres = scraping.fetch_genre_options()
        except requests.RequestException as error:
            genres = scraping.GENRE_OPTIONS_FALLBACK.copy()
            warnings.append(f"Daftar genre sementara memakai pilihan bawaan: {error}")

        try:
            job_options = scraping_karirhub.fetch_filter_options()
        except (requests.RequestException, ValueError, KeyError, TypeError) as error:
            if options_cache is not None:
                job_options = {
                    key: options_cache[key]
                    for key in ("provinces", "educations", "job_types", "sources")
                }
                warnings.append(f"Pilihan Karirhub memakai cache terakhir: {error}")
            else:
                job_options = {"provinces": [], "educations": [], "job_types": [], "sources": scraping_karirhub.SOURCE_CATEGORIES.copy()}
                warnings.append(f"Pilihan wilayah Karirhub belum dapat dimuat: {error}")

        options_cache = {"genres": genres, **job_options, "warnings": warnings}
        options_cached_at = time.monotonic()
        return options_cache


def update_job(job_id: str, **changes) -> None:
    with jobs_lock:
        if job_id in jobs:
            jobs[job_id].update(changes)


def progress_callback(job_id: str):
    def update(message: str, percent: int) -> None:
        with jobs_lock:
            record = jobs.get(job_id)
            if record is None:
                return
            record["progress"] = max(record["progress"], min(99, int(percent)))
            record["message"] = str(message)
    return update


def run_anime_job(job_id: str, filters: dict) -> None:
    output = OUTPUT_DIR / f"anime_{job_id}.xlsx"
    update_job(job_id, state="running", message="Memulai pengambilan data anime")
    try:
        records = scraping.scrape_anime(
            months=filters["months"],
            status=filters["status"],
            genre=filters["genre"],
            release_days=filters["release_days"],
            max_pages=filters["max_pages"],
            delay=0.75,
            progress=progress_callback(job_id),
        )
        scraping.save_anime_xlsx(records, output)
        update_job(
            job_id,
            state="complete",
            progress=100,
            message="Selesai. File XLSX siap diunduh.",
            count=len(records),
            output=str(output),
            filename="hasil_scraping_anime.xlsx",
        )
    except Exception as error:
        update_job(job_id, state="error", message=f"Pengambilan anime gagal: {error}", error=str(error))


def run_karirhub_job(job_id: str, filters: dict) -> None:
    output = OUTPUT_DIR / f"karirhub_{job_id}.xlsx"
    update_job(job_id, state="running", message="Memulai pengambilan lowongan Karirhub")
    try:
        records, failures = scraping_karirhub.scrape_jobs_filtered(
            province_id=filters["province_id"],
            province_label=filters["province_label"],
            job_type_id=filters["job_type_id"],
            education_id=filters["education_id"],
            source=filters["source"],
            page_size=18,
            max_jobs=filters["max_jobs"],
            delay=0.5,
            progress=progress_callback(job_id),
        )
        scraping_karirhub.save_xlsx(records, output)
        message = "Selesai. File XLSX siap diunduh."
        if failures:
            message += f" {failures} detail lowongan gagal dimuat; data daftar yang tersedia tetap disimpan."
        update_job(
            job_id,
            state="complete",
            progress=100,
            message=message,
            count=len(records),
            output=str(output),
            filename="hasil_scraping_lowongan_karirhub.xlsx",
        )
    except Exception as error:
        update_job(job_id, state="error", message=f"Pengambilan lowongan gagal: {error}", error=str(error))


def run_bps_job(job_id: str, filters: dict) -> None:
    output = OUTPUT_DIR / f"bps_{job_id}.xlsx"
    update_job(job_id, state="running", message="Memulai pengambilan statistik BPS")
    try:
        rows, missing_years = scraping_bps.scrape_table_by_year(
            api_key=filters["api_key"],
            category_name=filters["category_name"],
            subject_name=filters["subject_name"],
            table_id=filters["table_id"],
            table_title=filters["table_title"],
            table_source=filters["table_source"],
            start_year=filters["start_year"],
            end_year=filters["end_year"],
            delay=0.25,
            progress=progress_callback(job_id),
        )
        scraping_bps.save_xlsx(rows, output, {
            "Kategori subjek": filters["category_name"],
            "Jenis data": filters["subject_name"],
            "Nama tabel": filters["table_title"],
            "Sumber tabel": filters["table_source_name"],
            "Rentang tahun": f"{filters['start_year']}–{filters['end_year']}",
            "Tahun tanpa data": missing_years,
        })
        message = "Selesai. File XLSX siap diunduh."
        if missing_years:
            message += f" Data tidak tersedia untuk {missing_years} tahun dalam rentang pilihan."
        update_job(
            job_id,
            state="complete",
            progress=100,
            message=message,
            count=len(rows),
            output=str(output),
            filename="hasil_statistik_bps.xlsx",
        )
    except Exception as error:
        update_job(job_id, state="error", message=f"Pengambilan statistik BPS gagal: {error}", error=str(error))


def run_business_job(job_id: str, filters: dict) -> None:
    output = OUTPUT_DIR / f"business_{job_id}.xlsx"
    update_job(job_id, state="running", message="Memulai pencarian direktori bisnis OpenStreetMap")
    try:
        rows = scraping_business.search_businesses(
            area_id=filters["area_id"],
            scope_label=filters["scope_label"],
            category=filters["category"],
            keyword=filters["keyword"],
            max_results=filters["max_results"],
            progress=progress_callback(job_id),
            regency_label=filters["regency_label"],
            district_label=filters["district_label"],
        )
        scraping_business.save_xlsx(rows, output, {
            "Wilayah": filters["scope_label"],
            "Kabupaten/kota": filters["regency_label"],
            "Kecamatan": filters["district_label"] or "(seluruh kabupaten/kota)",
            "Kategori": filters["category_label"],
            "Kata kunci nama": filters["keyword"] or "(semua nama dalam kategori)",
            "Batas hasil": filters["max_results"],
            "Jumlah usaha ditemukan": len(rows),
        })
        update_job(
            job_id,
            state="complete",
            progress=100,
            message="Selesai. File XLSX siap diunduh.",
            count=len(rows),
            output=str(output),
            filename="direktori_bisnis_sumatera_selatan.xlsx",
        )
    except Exception as error:
        update_job(job_id, state="error", message=f"Pencarian direktori bisnis gagal: {error}", error=str(error))


def public_job(record: dict) -> dict:
    return {
        "id": record["id"],
        "kind": record["kind"],
        "state": record["state"],
        "progress": record["progress"],
        "message": record["message"],
        "count": record.get("count"),
        "filename": record.get("filename"),
        "download_url": f"/api/jobs/{record['id']}/download" if record.get("output") else None,
        "error": record.get("error"),
    }


def create_job(kind: str, worker, filters: dict) -> dict:
    job_id = uuid.uuid4().hex
    with jobs_lock:
        jobs[job_id] = {
            "id": job_id,
            "kind": kind,
            "state": "queued",
            "progress": 0,
            "message": "Menunggu proses dimulai",
        }
    executor.submit(worker, job_id, filters)
    with jobs_lock:
        return public_job(jobs[job_id])


def parse_int(value, label: str, minimum: int, maximum: int) -> int:
    try:
        result = int(value)
    except (TypeError, ValueError):
        raise ValueError(f"Nilai {label} harus berupa angka.") from None
    if not minimum <= result <= maximum:
        raise ValueError(f"Nilai {label} harus berada di antara {minimum} dan {maximum}.")
    return result


@app.get("/")
def index():
    return render_template("index.html", business_categories=scraping_business.BUSINESS_CATEGORIES)


@app.get("/api/options")
def api_options():
    force = request.args.get("refresh") == "1"
    return jsonify(load_options(force=force))


@app.get("/api/business/areas")
def business_areas():
    level = str(request.args.get("level") or "")
    parent_area_id = str(request.args.get("parent_area_id") or "")
    try:
        areas = scraping_business.fetch_admin_areas(level, parent_area_id)
    except ValueError as error:
        return jsonify({"error": str(error)}), 400
    except scraping_business.OverpassError as error:
        return jsonify({"error": str(error)}), 503
    return jsonify({"areas": areas})


@app.post("/api/scrape/anime")
def submit_anime():
    body = request.get_json(silent=True) or {}
    if not isinstance(body, dict):
        return jsonify({"error": "Format filter anime tidak valid."}), 400
    try:
        months = parse_int(body.get("months"), "rentang bulan", 1, 12)
        if months not in {1, 6, 12}:
            raise ValueError("Rentang bulan anime harus 1, 6, atau 12.")
        status = str(body.get("status") or "ongoing")
        if status not in ALLOWED_STATUS:
            raise ValueError("Status anime tidak dikenal.")
        days = body.get("release_days") or []
        if not isinstance(days, list) or any(day not in WEEKDAYS for day in days):
            raise ValueError("Pilihan hari rilis tidak valid.")
        max_pages = parse_int(body.get("max_pages", 0), "batas halaman", 0, 1000)
        selected_genre = str(body.get("genre") or "").strip()
        available_genres = load_options().get("genres", [])
        if selected_genre and selected_genre not in available_genres:
            raise ValueError("Genre yang dipilih tidak tersedia. Muat ulang daftar genre.")
    except ValueError as error:
        return jsonify({"error": str(error)}), 400

    filters = {
        "months": months,
        "status": status,
        "genre": selected_genre,
        "release_days": days,
        "max_pages": max_pages,
    }
    return jsonify(create_job("anime", run_anime_job, filters)), 202


@app.post("/api/scrape/karirhub")
def submit_karirhub():
    body = request.get_json(silent=True) or {}
    if not isinstance(body, dict):
        return jsonify({"error": "Format filter lowongan tidak valid."}), 400
    options = load_options()
    try:
        province_id = str(body.get("province_id") or "")
        province = next((item for item in options["provinces"] if item["id"] == province_id), None)
        if province is None:
            raise ValueError("Pilih salah satu provinsi yang tersedia.")

        job_type_id = str(body.get("job_type_id") or "")
        if job_type_id and job_type_id not in {item["id"] for item in options["job_types"]}:
            raise ValueError("Jenis pekerjaan yang dipilih tidak tersedia.")
        education_id = str(body.get("education_id") or "")
        if education_id and education_id not in {item["id"] for item in options["educations"]}:
            raise ValueError("Pendidikan yang dipilih tidak tersedia.")
        source = str(body.get("source") or "all")
        if source not in SOURCE_VALUES:
            raise ValueError("Sumber lowongan yang dipilih tidak tersedia.")
        max_jobs = parse_int(body.get("max_jobs", 0), "batas lowongan", 0, 5000)
    except (KeyError, ValueError) as error:
        return jsonify({"error": str(error)}), 400

    filters = {
        "province_id": province["id"],
        "province_label": province["name"],
        "job_type_id": job_type_id,
        "education_id": education_id,
        "source": source,
        "max_jobs": max_jobs,
    }
    return jsonify(create_job("karirhub", run_karirhub_job, filters)), 202


@app.post("/api/bps/subjects")
def bps_subjects():
    body = request.get_json(silent=True) or {}
    if not isinstance(body, dict):
        return jsonify({"error": "Format pilihan subjek BPS tidak valid."}), 400
    try:
        subjects = scraping_bps.fetch_subject_types(body.get("api_key", ""), str(body.get("category_id") or ""))
    except (ValueError, scraping_bps.BPSAPIError) as error:
        return jsonify({"error": str(error)}), 400
    return jsonify({"subjects": subjects})


@app.post("/api/bps/tables")
def bps_tables():
    body = request.get_json(silent=True) or {}
    if not isinstance(body, dict):
        return jsonify({"error": "Format pilihan tabel BPS tidak valid."}), 400
    try:
        tables = scraping_bps.fetch_tables(body.get("api_key", ""), str(body.get("subject_id") or ""))
    except (ValueError, scraping_bps.BPSAPIError) as error:
        return jsonify({"error": str(error)}), 400
    return jsonify({"tables": tables})


@app.post("/api/scrape/bps")
def submit_bps():
    body = request.get_json(silent=True) or {}
    if not isinstance(body, dict):
        return jsonify({"error": "Format filter statistik BPS tidak valid."}), 400
    try:
        api_key = str(body.get("api_key") or "").strip()
        if not api_key:
            raise ValueError("Masukkan API key WebAPI BPS terlebih dahulu.")
        category_id = str(body.get("category_id") or "")
        category = next((item for item in scraping_bps.SUBJECT_CATEGORIES if item["id"] == category_id), None)
        if category is None:
            raise ValueError("Kategori subjek BPS tidak dikenal.")
        subject_id = str(body.get("subject_id") or "")
        subjects = scraping_bps.fetch_subject_types(api_key, category_id)
        subject = next((item for item in subjects if item["id"] == subject_id), None)
        if subject is None:
            raise ValueError("Pilih jenis data BPS dari kategori yang dipilih.")
        tables = scraping_bps.fetch_tables(api_key, subject_id)
        table_id = str(body.get("table_id") or "")
        table = next((item for item in tables if item["id"] == table_id), None)
        if table is None:
            raise ValueError("Pilih tabel dari jenis data yang dipilih.")
        this_year = datetime.now().year
        start_year = parse_int(body.get("start_year"), "tahun awal", 1900, this_year)
        end_year = parse_int(body.get("end_year"), "tahun akhir", 1900, this_year)
        if start_year > end_year:
            raise ValueError("Tahun awal harus sama dengan atau lebih kecil dari tahun akhir.")
    except (ValueError, scraping_bps.BPSAPIError) as error:
        return jsonify({"error": str(error)}), 400

    filters = {
        "api_key": api_key,
        "category_name": category["name"],
        "subject_name": subject["name"],
        "table_id": table["id"],
        "table_title": table["title"],
        "table_source": table["source"],
        "table_source_name": table["source_name"],
        "start_year": start_year,
        "end_year": end_year,
    }
    return jsonify(create_job("bps", run_bps_job, filters)), 202


@app.post("/api/scrape/business")
def submit_business():
    body = request.get_json(silent=True)
    if not isinstance(body, dict):
        return jsonify({"error": "Format filter direktori bisnis tidak valid."}), 400
    try:
        area_id = str(body.get("area_id") or "")
        area = scraping_business.get_known_area(area_id)
        if area is None:
            raise ValueError("Pilih wilayah dari daftar Sumatera Selatan yang tersedia.")
        if area["admin_level"] == "6":
            regency = scraping_business.get_known_area(area.get("parent_id", ""))
            if regency is None or regency["admin_level"] != "5":
                raise ValueError("Daftar wilayah sudah kedaluwarsa. Muat ulang tab Bisnis lalu pilih wilayah lagi.")
            regency_label = regency["name"]
            district_label = area["name"]
        else:
            regency_label = area["name"]
            district_label = ""

        category = str(body.get("category") or "")
        category_data = next((item for item in scraping_business.BUSINESS_CATEGORIES if item["value"] == category), None)
        if category_data is None:
            raise ValueError("Pilih kategori bisnis yang tersedia.")
        keyword = str(body.get("keyword") or "").strip()
        if len(keyword) > 80:
            raise ValueError("Kata kunci maksimal 80 karakter.")
        max_results = parse_int(body.get("max_results", 500), "batas hasil", 1, 2000)
    except ValueError as error:
        return jsonify({"error": str(error)}), 400

    filters = {
        "area_id": area["id"],
        "scope_label": area["name"],
        "regency_label": regency_label,
        "district_label": district_label,
        "category": category,
        "category_label": category_data["label"],
        "keyword": keyword,
        "max_results": max_results,
    }
    return jsonify(create_job("business", run_business_job, filters)), 202


@app.get("/api/jobs/<job_id>")
def job_status(job_id: str):
    with jobs_lock:
        record = jobs.get(job_id)
        if record is None:
            return jsonify({"error": "Proses tidak ditemukan. Jika server dimulai ulang, proses lama sudah dihapus."}), 404
        return jsonify(public_job(record))


@app.get("/api/jobs/<job_id>/download")
def download_job(job_id: str):
    with jobs_lock:
        record = jobs.get(job_id)
        output = Path(record["output"]) if record and record.get("output") else None
        filename = record.get("filename", "hasil.xlsx") if record else "hasil.xlsx"
    if output is None or not output.is_file() or output.parent.resolve() != OUTPUT_DIR.resolve():
        return jsonify({"error": "File hasil belum tersedia."}), 404
    return send_file(output, as_attachment=True, download_name=filename)


if __name__ == "__main__":
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    print("Buka http://127.0.0.1:5000 di browser untuk menggunakan scraper.")
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)
