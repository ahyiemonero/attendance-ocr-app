import os
import re
import uuid
import sqlite3
import threading
from datetime import datetime
from dateutil import parser

from flask import Flask, render_template, request, redirect, url_for, send_file, send_from_directory, abort, jsonify
from werkzeug.utils import secure_filename

from PIL import Image, ImageEnhance, ImageFilter
import pytesseract

import cv2
import numpy as np

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.drawing.image import Image as ExcelImage


app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_TIME_IN = os.path.join(BASE_DIR, "uploads", "time_in")
UPLOAD_TIME_OUT = os.path.join(BASE_DIR, "uploads", "time_out")
GENERATED_DIR = os.path.join(BASE_DIR, "generated")
LOGO_PATH = os.path.join(BASE_DIR, "static", "henderson_logo.png")
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "attendance.db")

os.makedirs(UPLOAD_TIME_IN, exist_ok=True)
os.makedirs(UPLOAD_TIME_OUT, exist_ok=True)
os.makedirs(GENERATED_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)

# Windows Tesseract path.
# Later on Ubuntu, comment this line.
if os.name == "nt":
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp"}

TIME_IN_RECORDS = []
TIME_OUT_RECORDS = []
JOBS = {}

EMPLOYEES_SEED = [
    ("15829A", "SAMUEL ONG KAI WEN"),
    ("17327", "SANJAY"),
    ("14464", "GENIEVE CAROLE ADRIANO"),
    ("17769", "GOVINDHRAJ MURUGAN"),
]


def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db_connection()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            emp_code TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL,
            active INTEGER NOT NULL DEFAULT 1
        )
    """)

    for emp_code, name in EMPLOYEES_SEED:
        conn.execute("""
            INSERT OR IGNORE INTO employees (emp_code, name, active)
            VALUES (?, ?, 1)
        """, (emp_code, name))

    conn.commit()
    conn.close()


def get_employees():
    conn = get_db_connection()

    employees = conn.execute("""
        SELECT emp_code, name
        FROM employees
        WHERE active = 1
        ORDER BY name ASC
    """).fetchall()

    conn.close()
    return employees


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def save_uploaded_file(file, target_folder):
    original_filename = secure_filename(file.filename)
    ext = original_filename.rsplit(".", 1)[1].lower()
    unique_filename = f"{uuid.uuid4().hex}.{ext}"
    file_path = os.path.join(target_folder, unique_filename)
    file.save(file_path)

    return file_path, original_filename, unique_filename


def clean_ocr_text(text):
    if not text:
        return ""

    text = text.replace("\n", " ")
    text = text.replace("\r", " ")
    text = re.sub(r"\s+", " ", text)

    replacements = {
        "S aturday": "Saturday",
        "M onday": "Monday",
        "T uesday": "Tuesday",
        "W ednesday": "Wednesday",
        "T hursday": "Thursday",
        "F riday": "Friday",
        "S unday": "Sunday",

        "J anuary": "January",
        "F ebruary": "February",
        "M arch": "March",
        "A pril": "April",
        "M ay": "May",
        "J une": "June",
        "J uly": "July",
        "A ugust": "August",
        "S eptember": "September",
        "O ctober": "October",
        "N ovember": "November",
        "D ecember": "December",

        "Sept ": "September ",
        "Sept,": "September,",

        "O:": "0:",
        "o:": "0:",
        "I:": "1:",
        "l:": "1:",
        "|:": "1:",
        " :": ":",
    }

    for wrong, correct in replacements.items():
        text = text.replace(wrong, correct)

    return text.strip()


def extract_datetime_from_filename(filename):
    """
    Extract datetime from WhatsApp-style filename.

    Supports:
    - WhatsApp Image 2026-05-28 at 18.57.22.jpeg
    - WhatsApp_Image_2026-05-28_at_18.57.22.jpeg
    - WhatsApp_Image_2026-05-28_at_18.57.22_1.jpeg

    Used mainly for the DATE and chronological sorting.
    OCR time from the photo overlay is still preferred when detected.
    """
    if not filename:
        return None

    patterns = [
        r"(\d{4})-(\d{2})-(\d{2})[\s_]+at[\s_]+(\d{1,2})[.\-:](\d{2})[.\-:](\d{2})",
        r"(\d{4})-(\d{2})-(\d{2}).*?(\d{1,2})[.\-:](\d{2})[.\-:](\d{2})",
    ]

    for pattern in patterns:
        match = re.search(pattern, filename, flags=re.IGNORECASE)
        if match:
            year, month, day, hour, minute, second = match.groups()

            try:
                return datetime(
                    int(year),
                    int(month),
                    int(day),
                    int(hour),
                    int(minute),
                    int(second)
                )
            except Exception:
                return None

    return None


def extract_time_from_text(text, filename_dt=None):
    """
    Strict OCR time extractor.

    Supports:
    - 06:40:03
    - 6:40:03
    - 18:34:24
    - 06.40.03
    - 18.34.24
    - 06:34pm
    - 06:34 pm
    - 7:03 PM
    - 07:03 PM
    """

    if not text:
        return None

    text = clean_ocr_text(text)

    # Normalize separators
    text = text.replace(".", ":")
    text = text.replace(";", ":")
    text = text.replace(",", ":")

    candidates = []

    # 24-hour format with seconds: 18:34:24
    pattern_24h_seconds = r"\b([01]?\d|2[0-3]):([0-5]\d):([0-5]\d)\b"

    for match in re.finditer(pattern_24h_seconds, text):
        hour, minute, second = match.groups()

        try:
            detected_time = datetime.strptime(
                f"{int(hour):02d}:{minute}:{second}",
                "%H:%M:%S"
            ).time()

            if filename_dt:
                ocr_dt = datetime.combine(filename_dt.date(), detected_time)
                diff_seconds = abs((ocr_dt - filename_dt).total_seconds())
                candidates.append((diff_seconds, detected_time))
            else:
                candidates.append((0, detected_time))

        except Exception:
            continue

    # 12-hour format: 06:34pm / 06:34 pm / 7:03 PM
    pattern_12h = r"\b(1[0-2]|0?[1-9]):([0-5]\d)(?::([0-5]\d))?\s*(AM|PM|am|pm)\b"

    for match in re.finditer(pattern_12h, text):
        hour, minute, second, ampm = match.groups()
        second = second or "00"

        try:
            detected_time = datetime.strptime(
                f"{int(hour):02d}:{minute}:{second} {ampm.upper()}",
                "%I:%M:%S %p"
            ).time()

            if filename_dt:
                ocr_dt = datetime.combine(filename_dt.date(), detected_time)
                diff_seconds = abs((ocr_dt - filename_dt).total_seconds())
                candidates.append((diff_seconds, detected_time))
            else:
                candidates.append((0, detected_time))

        except Exception:
            continue

    if not candidates:
        return None

    # Choose OCR time closest to filename time.
    candidates.sort(key=lambda x: x[0])

    best_diff, best_time = candidates[0]

    # Reject OCR time if too far from WhatsApp filename time.
    # 30 minutes window.
    if filename_dt and best_diff > 1800:
        return None

    return best_time

def extract_datetime_from_text(text):
    """
    Supported examples:
    - Saturday, May 2, 2026 06:40:34
    - May 2, 2026 06:40:34
    - 2 May 2026 18:48:07
    - 3 May 2026 18:45
    - 01/05/2026 06:31:37
    - 01-05-2026 06:31:37
    - 2026-05-01 06:31:37
    """

    text = clean_ocr_text(text)

    text = re.sub(
        r"\b(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),?\s+",
        "",
        text,
        flags=re.IGNORECASE
    )

    patterns = [
        r"\b[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?\b",
        r"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?\b",
        r"\b\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?\b",
        r"\b\d{1,2}-\d{1,2}-\d{4}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?\b",
        r"\b\d{4}-\d{1,2}-\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM|am|pm)?\b",
    ]

    candidates = []

    for pattern in patterns:
        matches = re.findall(pattern, text, flags=re.IGNORECASE)

        for raw_datetime in matches:
            try:
                dt = parser.parse(raw_datetime, dayfirst=True)

                if 2020 <= dt.year <= 2035:
                    candidates.append(dt)

            except Exception:
                continue

    if candidates:
        return candidates[0]

    return None


def get_image_regions(image):
    """
    Scan full image first, then common timestamp zones.
    """
    width, height = image.size

    regions = []

    regions.append(("full", image))

    regions.append((
        "bottom_left",
        image.crop((0, int(height * 0.50), int(width * 0.70), height))
    ))

    regions.append((
        "bottom_right",
        image.crop((int(width * 0.30), int(height * 0.50), width, height))
    ))

    regions.append((
        "top_left",
        image.crop((0, 0, int(width * 0.70), int(height * 0.55)))
    ))

    regions.append((
        "top_right",
        image.crop((int(width * 0.30), 0, width, int(height * 0.55)))
    ))

    regions.append((
        "lower_half",
        image.crop((0, int(height * 0.45), width, height))
    ))

    regions.append((
        "upper_half",
        image.crop((0, 0, width, int(height * 0.60)))
    ))

    return regions


def pil_to_cv2(pil_image):
    rgb = np.array(pil_image.convert("RGB"))
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


def cv2_to_pil(cv2_image):
    if len(cv2_image.shape) == 2:
        return Image.fromarray(cv2_image)

    rgb = cv2.cvtColor(cv2_image, cv2.COLOR_BGR2RGB)
    return Image.fromarray(rgb)


def preprocess_for_ocr(region_image):
    """
    Creates multiple OCR-friendly versions using PIL and OpenCV.
    """
    processed_images = []

    region_image = region_image.convert("RGB")

    enlarged = region_image.resize(
        (region_image.width * 2, region_image.height * 2),
        Image.Resampling.LANCZOS
    )
    processed_images.append(("pil_enlarged", enlarged))

    gray = enlarged.convert("L")
    processed_images.append(("pil_gray", gray))

    sharp = gray.filter(ImageFilter.SHARPEN)
    processed_images.append(("pil_sharp", sharp))

    contrast = ImageEnhance.Contrast(gray).enhance(2.5)
    processed_images.append(("pil_contrast_2_5", contrast))

    contrast_high = ImageEnhance.Contrast(gray).enhance(4.0)
    processed_images.append(("pil_contrast_4_0", contrast_high))

    bw_130 = gray.point(lambda x: 0 if x < 130 else 255, "1")
    processed_images.append(("pil_bw_130", bw_130))

    bw_160 = gray.point(lambda x: 0 if x < 160 else 255, "1")
    processed_images.append(("pil_bw_160", bw_160))

    bw_190 = gray.point(lambda x: 0 if x < 190 else 255, "1")
    processed_images.append(("pil_bw_190", bw_190))

    # OpenCV preprocessing
    cv_img = pil_to_cv2(enlarged)
    cv_gray = cv2.cvtColor(cv_img, cv2.COLOR_BGR2GRAY)

    cv_gray = cv2.bilateralFilter(cv_gray, 9, 75, 75)
    processed_images.append(("cv_gray_bilateral", cv2_to_pil(cv_gray)))

    _, cv_thresh_150 = cv2.threshold(cv_gray, 150, 255, cv2.THRESH_BINARY)
    processed_images.append(("cv_thresh_150", cv2_to_pil(cv_thresh_150)))

    _, cv_thresh_180 = cv2.threshold(cv_gray, 180, 255, cv2.THRESH_BINARY)
    processed_images.append(("cv_thresh_180", cv2_to_pil(cv_thresh_180)))

    cv_adaptive = cv2.adaptiveThreshold(
        cv_gray,
        255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        31,
        2
    )
    processed_images.append(("cv_adaptive", cv2_to_pil(cv_adaptive)))

    cv_inverted = cv2.bitwise_not(cv_adaptive)
    processed_images.append(("cv_adaptive_inverted", cv2_to_pil(cv_inverted)))

    return processed_images


def ocr_full_image(image_path):
    """
    Optimized OCR:
    - scans full image
    - scans likely timestamp zones
    - tries PIL and OpenCV preprocessing
    - stops early once a valid timestamp is found
    """
    try:
        image = Image.open(image_path).convert("RGB")

        all_ocr_text = []

        configs = [
            "--psm 6",
        ]

        regions = get_image_regions(image)

        for region_name, region_image in regions:
            processed_versions = preprocess_for_ocr(region_image)

            for version_name, processed_image in processed_versions:
                for config in configs:
                    try:
                        text = pytesseract.image_to_string(
                            processed_image,
                            config=config
                        )

                        if text:
                            labelled_text = f"[{region_name}-{version_name}-{config}] {text}"
                            all_ocr_text.append(labelled_text)

                            detected_dt = extract_datetime_from_text(text)

                            if detected_dt:
                                return detected_dt, " ".join(all_ocr_text)

                    except Exception:
                        continue

        combined_text = " ".join(all_ocr_text)
        detected_dt = extract_datetime_from_text(combined_text)

        return detected_dt, combined_text

    except Exception as e:
        return None, f"OCR error: {str(e)}"



def ocr_time_only_fast(image_path, filename_dt=None):
    """
    Safer OCR mode:
    - Uses filename datetime only as sanity reference.
    - OCR still provides the actual time.
    - Rejects random/wrong OCR times that are too far from filename time.
    """

    try:
        image = Image.open(image_path).convert("RGB")
        width, height = image.size

        regions = [
             # Bottom-left timestamp, like 06:34pm
             ("bottom_left", image.crop((0, int(height * 0.55), int(width * 0.85), height))),

             # Bottom-right timestamp
             ("bottom_right", image.crop((int(width * 0.15), int(height * 0.55), width, height))),

             # Top-right timestamp, like "27 May 2026 at 7:03 PM"
             ("top_right", image.crop((int(width * 0.35), 0, width, int(height * 0.35)))),
         ]

        all_ocr_text = []

        # Do NOT whitelist only numbers. It can destroy the timestamp structure.
        configs = [
            "--psm 6",
            "--psm 11",
        ]

        for region_name, region in regions:
            enlarged = region.resize(
                (region.width * 3, region.height * 3),
                Image.Resampling.LANCZOS
            )

            gray = enlarged.convert("L")
            contrast = ImageEnhance.Contrast(gray).enhance(3.5)
            sharp = contrast.filter(ImageFilter.SHARPEN)

            cv_img = np.array(sharp)
            inverted = Image.fromarray(cv2.bitwise_not(cv_img))

            versions = [
                ("contrast", contrast),
            ]

            for version_name, version_img in versions:
                for config in configs:
                    try:
                        text = pytesseract.image_to_string(
                            version_img,
                            config=config
                        )
                    except Exception:
                        continue

                    if text:
                        all_ocr_text.append(f"[{region_name}-{version_name}-{config}] {text}")

                        detected_time = extract_time_from_text(text, filename_dt=filename_dt)

                        if detected_time:
                            return detected_time, " ".join(all_ocr_text)

        return None, " ".join(all_ocr_text)

    except Exception as e:
        return None, f"OCR error: {str(e)}"


def format_date(dt):
    if not dt:
        return ""
    return f"{dt.day}/{dt.month}/{dt.year}"


def format_time(dt):
    if not dt:
        return ""
    return dt.strftime("%H:%M:%S")

def time_difference_seconds(filename_dt, detected_time):
    """
    Compare filename datetime against OCR-detected time.
    Returns difference in seconds.
    """
    if not filename_dt or not detected_time:
        return None

    ocr_dt = datetime.combine(filename_dt.date(), detected_time)

    return abs((ocr_dt - filename_dt).total_seconds())

def process_uploads(files, upload_type):
    """
    Processing logic:

    1. Extract DATE from WhatsApp filename.
    2. OCR reads TIME from the photo overlay.
    3. If OCR succeeds:
       date = filename date
       time = OCR time
    4. If OCR fails:
       date = filename date
       time = blank for manual correction
    5. Sorting:
       - If OCR time exists: sort by filename date + OCR time
       - If OCR fails: sort by filename datetime, but display blank time
    """
    records = []

    target_folder = UPLOAD_TIME_IN if upload_type == "time_in" else UPLOAD_TIME_OUT

    for file in files:
        if file and allowed_file(file.filename):
            file_path, original_filename, stored_filename = save_uploaded_file(file, target_folder)

            filename_dt = extract_datetime_from_filename(original_filename)

            detected_time = None
            ocr_text = ""
            display_date = ""
            display_time = ""
            sort_dt = None
            source_used = ""

            if filename_dt:
                detected_time, ocr_text = ocr_time_only_fast(file_path, filename_dt=filename_dt)

                display_date = format_date(filename_dt)

                if detected_time:
                    # Use filename DATE only + OCR TIME
                    final_dt = datetime.combine(filename_dt.date(), detected_time)

                    display_time = format_time(final_dt)
                    sort_dt = final_dt
                    source_used = "OCR_TIME"
                else:
                    # Do not use filename time as display time
                    # Keep date only, leave time blank for manual correction
                    display_time = ""
                    sort_dt = filename_dt
                    source_used = "OCR_FAILED_TIME_BLANK"
                    ocr_text = ocr_text or "OCR failed. Time left blank for manual correction."

            else:
                final_dt = None
                display_date = ""
                display_time = ""
                sort_dt = datetime.max
                source_used = "FILENAME_FAILED"
                ocr_text = "Filename date could not be detected. Manual entry required."

            records.append({
                "id": uuid.uuid4().hex,
                "original_filename": original_filename,
                "stored_filename": stored_filename,
                "upload_type": upload_type,
                "file_path": file_path,
                "filename_datetime": filename_dt,
                "detected_datetime": sort_dt,
                "date": display_date,
                "time": display_time,
                "ocr_text": ocr_text,
                "source_used": source_used,
            })

    records.sort(key=lambda x: x["detected_datetime"] or datetime.max)
    return records

def process_saved_files(saved_files, upload_type, job_id, start_progress, end_progress):
    records = []
    total_files = len(saved_files)

    for index, saved in enumerate(saved_files):
        file_path = saved["file_path"]
        original_filename = saved["original_filename"]
        stored_filename = saved["stored_filename"]

        filename_dt = extract_datetime_from_filename(original_filename)

        detected_time = None
        ocr_text = ""
        display_date = ""
        display_time = ""
        sort_dt = None
        source_used = ""

        if filename_dt:
            detected_time, ocr_text = ocr_time_only_fast(file_path, filename_dt=filename_dt)

            display_date = format_date(filename_dt)

            if detected_time:
                final_dt = datetime.combine(filename_dt.date(), detected_time)
                display_time = format_time(final_dt)
                sort_dt = final_dt
                source_used = "OCR_TIME"
            else:
                display_time = ""
                sort_dt = filename_dt
                source_used = "OCR_FAILED_TIME_BLANK"
                ocr_text = ocr_text or "OCR failed. Time left blank for manual correction."
        else:
            display_date = ""
            display_time = ""
            sort_dt = datetime.max
            source_used = "FILENAME_FAILED"
            ocr_text = "Filename date could not be detected. Manual entry required."

        records.append({
            "id": uuid.uuid4().hex,
            "original_filename": original_filename,
            "stored_filename": stored_filename,
            "upload_type": upload_type,
            "file_path": file_path,
            "filename_datetime": filename_dt,
            "detected_datetime": sort_dt,
            "date": display_date,
            "time": display_time,
            "ocr_text": ocr_text,
            "source_used": source_used,
        })

        if total_files > 0:
            progress_range = end_progress - start_progress
            progress = start_progress + int(((index + 1) / total_files) * progress_range)
            JOBS[job_id]["progress"] = progress
            JOBS[job_id]["message"] = f"Processing {upload_type.replace('_', ' ').upper()} photos: {index + 1}/{total_files}"

    records.sort(key=lambda x: x["detected_datetime"] or datetime.max)
    return records


def process_job_background(job_id, saved_time_in_files, saved_time_out_files):
    try:
        JOBS[job_id]["message"] = "Processing TIME IN photos..."
        JOBS[job_id]["progress"] = 5

        time_in_records = process_saved_files(
            saved_time_in_files,
            "time_in",
            job_id,
            5,
            50
        )

        JOBS[job_id]["message"] = "Processing TIME OUT photos..."
        JOBS[job_id]["progress"] = 50

        time_out_records = process_saved_files(
            saved_time_out_files,
            "time_out",
            job_id,
            50,
            95
        )

        JOBS[job_id]["time_in_records"] = time_in_records
        JOBS[job_id]["time_out_records"] = time_out_records
        JOBS[job_id]["progress"] = 100
        JOBS[job_id]["message"] = "Processing completed."
        JOBS[job_id]["status"] = "done"

    except Exception as e:
        JOBS[job_id]["status"] = "error"
        JOBS[job_id]["error"] = str(e)
        JOBS[job_id]["message"] = f"Error: {str(e)}"

def get_sorted_records():
    time_in_sorted = sorted(
        TIME_IN_RECORDS,
        key=lambda x: x["detected_datetime"] or datetime.max
    )

    time_out_sorted = sorted(
        TIME_OUT_RECORDS,
        key=lambda x: x["detected_datetime"] or datetime.max
    )

    return time_in_sorted, time_out_sorted


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        site = request.form.get("site", "").strip()
        month_year = request.form.get("month_year", "").strip()

        time_in_files = request.files.getlist("time_in_photos")
        time_out_files = request.files.getlist("time_out_photos")

        job_id = uuid.uuid4().hex

        JOBS[job_id] = {
            "status": "processing",
            "progress": 0,
            "message": "Starting upload processing...",
            "site": site,
            "month_year": month_year,
            "time_in_records": [],
            "time_out_records": [],
            "error": None,
        }

        # Save files first while request is active
        saved_time_in_files = []
        saved_time_out_files = []

        for file in time_in_files:
            if file and allowed_file(file.filename):
                file_path, original_filename, stored_filename = save_uploaded_file(file, UPLOAD_TIME_IN)
                saved_time_in_files.append({
                    "file_path": file_path,
                    "original_filename": original_filename,
                    "stored_filename": stored_filename,
                })

        for file in time_out_files:
            if file and allowed_file(file.filename):
                file_path, original_filename, stored_filename = save_uploaded_file(file, UPLOAD_TIME_OUT)
                saved_time_out_files.append({
                    "file_path": file_path,
                    "original_filename": original_filename,
                    "stored_filename": stored_filename,
                })

        thread = threading.Thread(
            target=process_job_background,
            args=(job_id, saved_time_in_files, saved_time_out_files),
            daemon=True
        )
        thread.start()

        return redirect(url_for("processing", job_id=job_id))

    return render_template("index.html")

@app.route("/uploaded/<upload_type>/<filename>")
def uploaded_file(upload_type, filename):
    if upload_type == "time_in":
        folder = UPLOAD_TIME_IN
    elif upload_type == "time_out":
        folder = UPLOAD_TIME_OUT
    else:
        abort(404)

    return send_from_directory(folder, filename)

@app.route("/processing/<job_id>")
def processing(job_id):
    if job_id not in JOBS:
        abort(404)

    return render_template("processing.html", job_id=job_id)


@app.route("/job_status/<job_id>")
def job_status(job_id):
    job = JOBS.get(job_id)

    if not job:
        return jsonify({
            "status": "missing",
            "progress": 0,
            "message": "Job not found."
        }), 404

    return jsonify({
        "status": job["status"],
        "progress": job["progress"],
        "message": job["message"],
        "error": job["error"],
    })

@app.route("/preview", methods=["GET"])
def preview():
    job_id = request.args.get("job_id")

    if job_id:
        job = JOBS.get(job_id)
        if not job:
            abort(404)

        site = job["site"]
        month_year = job["month_year"]
        time_in_sorted = sorted(
            job["time_in_records"],
            key=lambda x: x["detected_datetime"] or datetime.max
        )
        time_out_sorted = sorted(
            job["time_out_records"],
            key=lambda x: x["detected_datetime"] or datetime.max
        )
    else:
        site = request.args.get("site", "")
        month_year = request.args.get("month_year", "")
        time_in_sorted, time_out_sorted = get_sorted_records()

    max_rows = max(len(time_in_sorted), len(time_out_sorted))

    rows = []

    for i in range(max_rows):
        time_in = time_in_sorted[i] if i < len(time_in_sorted) else None
        time_out = time_out_sorted[i] if i < len(time_out_sorted) else None

        date_value = ""
        if time_in and time_in.get("date"):
            date_value = time_in["date"]
        elif time_out and time_out.get("date"):
            date_value = time_out["date"]

        rows.append({
            "index": i,
            "date": date_value,
            "time_in": time_in,
            "time_out": time_out,
            "site": site,
        })
    employees = get_employees()
    return render_template(
        "preview.html",
        rows=rows,
        site=site,
        month_year=month_year,
        employees=employees,
        job_id=job_id
    )
    
@app.route("/employees", methods=["GET", "POST"])
def employees_page():
    if request.method == "POST":
        emp_code = request.form.get("emp_code", "").strip()
        name = request.form.get("name", "").strip().upper()

        if emp_code and name:
            conn = get_db_connection()
            conn.execute("""
                INSERT INTO employees (emp_code, name, active)
                VALUES (?, ?, 1)
                ON CONFLICT(emp_code) DO UPDATE SET
                    name = excluded.name,
                    active = 1
            """, (emp_code, name))
            conn.commit()
            conn.close()

        return redirect(url_for("employees_page"))

    conn = get_db_connection()

    employees = conn.execute("""
    SELECT id, emp_code, name
    FROM employees
    ORDER BY name ASC
""").fetchall()

    conn.close()

    return render_template("employees.html", employees=employees)


@app.route("/employees/delete/<int:employee_id>", methods=["POST"])
def delete_employee(employee_id):
    conn = get_db_connection()
    conn.execute("""
        DELETE FROM employees
        WHERE id = ?
    """, (employee_id,))
    conn.commit()
    conn.close()

    return redirect(url_for("employees_page"))

def resize_image_portrait_for_excel(image_path, max_width=120, max_height=160):
    """
    Keeps the image aspect ratio and fits it inside a portrait-style Excel cell.
    Returns width and height in pixels.
    """
    with Image.open(image_path) as pil_img:
        original_width, original_height = pil_img.size

    if original_width == 0 or original_height == 0:
        return max_width, max_height

    # Scale image to fit inside max_width x max_height
    width_ratio = max_width / original_width
    height_ratio = max_height / original_height
    scale = min(width_ratio, height_ratio)

    new_width = int(original_width * scale)
    new_height = int(original_height * scale)

    return new_width, new_height

def resize_logo_for_excel(image_path, max_width=360, max_height=110):
    """
    Keeps logo aspect ratio and fits it into the top merged row.
    """
    with Image.open(image_path) as pil_img:
        original_width, original_height = pil_img.size

    if original_width == 0 or original_height == 0:
        return max_width, max_height

    width_ratio = max_width / original_width
    height_ratio = max_height / original_height
    scale = min(width_ratio, height_ratio)

    new_width = int(original_width * scale)
    new_height = int(original_height * scale)

    return new_width, new_height

@app.route("/generate_excel", methods=["POST"])
def generate_excel():
    month_year = request.form.get("month_year", "").strip()
    site = request.form.get("site", "").strip()
    row_count = int(request.form.get("row_count", 0))

    job_id = request.form.get("job_id", "").strip()

if job_id and job_id in JOBS:
    job = JOBS[job_id]

    time_in_sorted = sorted(
        job["time_in_records"],
        key=lambda x: x["detected_datetime"] or datetime.max
    )

    time_out_sorted = sorted(
        job["time_out_records"],
        key=lambda x: x["detected_datetime"] or datetime.max
    )
else:
    time_in_sorted, time_out_sorted = get_sorted_records()

    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"

    title = f"ATTENDANCE {month_year.upper()}"

    headers = [
        "SHIFT",
        "DATE",
        "PHOTO IN",
        "TIME IN",
        "PHOTO OUT",
        "TIME OUT",
        "EMP CODE",
        "NAME",
        "SITE",
    ]

    # Logo row
    ws.merge_cells("A1:I1")
    ws.row_dimensions[1].height = 105

    if os.path.exists(LOGO_PATH):
        logo = ExcelImage(LOGO_PATH)

        logo_width, logo_height = resize_logo_for_excel(
            LOGO_PATH,
            max_width=360,
            max_height=100
        )

        logo.width = logo_width
        logo.height = logo_height

        # Anchor logo around the center area of merged A1:I1
        ws.add_image(logo, "E1")

    # Title row
    ws.merge_cells("A2:I2")
    ws["A2"] = title
    ws["A2"].font = Font(size=16, bold=True, underline="single", color="000000")
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 30

    # Header row
    header_row = 3
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=header_row, column=col_num)
        cell.value = header
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="2E75B6")
    widths = {
        "A": 10,
        "B": 15,
        "C": 18,   # PHOTO IN - portrait size
        "D": 15,
        "E": 18,   # PHOTO OUT - portrait size
        "F": 15,
        "G": 15,
        "H": 32,
        "I": 35,
    }

    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    thin = Side(border_style="thin", color="B7C9E2")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    excel_row = 4

    for i in range(row_count):
        shift = request.form.get(f"shift_{i}", "").strip()
        date_value = request.form.get(f"date_{i}", "").strip()
        time_in_value = request.form.get(f"time_in_{i}", "").strip()
        time_out_value = request.form.get(f"time_out_{i}", "").strip()
        emp_code = request.form.get(f"emp_code_{i}", "").strip()
        name = request.form.get(f"name_{i}", "").strip()
        site_value = request.form.get(f"site_{i}", site).strip()

        ws.cell(row=excel_row, column=1).value = shift
        ws.cell(row=excel_row, column=2).value = date_value
        ws.cell(row=excel_row, column=4).value = time_in_value
        ws.cell(row=excel_row, column=6).value = time_out_value
        ws.cell(row=excel_row, column=7).value = emp_code
        ws.cell(row=excel_row, column=8).value = name
        ws.cell(row=excel_row, column=9).value = site_value

        ws.row_dimensions[excel_row].height = 125

        # Insert Photo In
        if i < len(time_in_sorted):
            img_path = time_in_sorted[i]["file_path"]
            if os.path.exists(img_path):
                img = ExcelImage(img_path)

                new_width, new_height = resize_image_portrait_for_excel(
                    img_path,
                    max_width=115,
                    max_height=160
                )

                img.width = new_width
                img.height = new_height

                # Anchor image to the PHOTO IN cell
                ws.add_image(img, f"C{excel_row}")

        # Insert Photo Out
        if i < len(time_out_sorted):
            img_path = time_out_sorted[i]["file_path"]
            if os.path.exists(img_path):
                img = ExcelImage(img_path)

                new_width, new_height = resize_image_portrait_for_excel(
                    img_path,
                    max_width=115,
                    max_height=160
                )

                img.width = new_width
                img.height = new_height

                # Anchor image to the PHOTO OUT cell
                ws.add_image(img, f"E{excel_row}")

        for col in range(1, 10):
            cell = ws.cell(row=excel_row, column=col)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        excel_row += 1

    # Prepared By row at bottom
    prepared_row = excel_row + 1

    # Merge A to D like your screenshot
    ws.merge_cells(start_row=prepared_row, start_column=1, end_row=prepared_row, end_column=4)

    prepared_cell = ws.cell(row=prepared_row, column=1)
    prepared_cell.value = "Prepared By : Thaqif Kasmani"
    prepared_cell.font = Font(size=12, bold=False, color="000000")
    prepared_cell.alignment = Alignment(horizontal="left", vertical="center")

    ws.row_dimensions[prepared_row].height = 25

    # Add border to merged prepared row area
    for col in range(1, 5):
        cell = ws.cell(row=prepared_row, column=col)
        cell.border = border

    output_filename = f"attendance_{uuid.uuid4().hex}.xlsx"
    output_path = os.path.join(GENERATED_DIR, output_filename)

    wb.save(output_path)

    return send_file(output_path, as_attachment=True, download_name="attendance.xlsx")


if __name__ == "__main__":
    init_db()
    app.run(debug=True)