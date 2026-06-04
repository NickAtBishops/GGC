"""
GGC Deal Engine — Backend Server v3

CHANGES vs. v4:
- Uses GGC's official blank template directly (GGC_Blank_Underwriting_Sizer.xlsx)
- Maps to GGC's exact category names (matching what they use in SUMIFS, including
  the "Electrcitiy" typo which is in their model)
- Adds a NEW "Miscellaneous" tab on top of GGC's template (since their model
  doesn't have one) — preserves Google Maps imagery, landmarks, demographics
- Adds Demographics section per Michael's feedback
- Lists individual rent roll units (one row per unit) instead of flattening counts

Run: python backend.py
Then open http://localhost:5001
"""

import os
import json
import time
import base64
import hashlib
import statistics
import traceback
import re
from pathlib import Path
from threading import Thread
from concurrent.futures import ThreadPoolExecutor, as_completed
from collections import Counter
from io import BytesIO
from urllib.parse import quote_plus
from google.cloud import documentai
from google.api_core.client_options import ClientOptions

import requests
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from dotenv import load_dotenv
from pydantic import BaseModel, Field, ValidationError, field_validator, model_validator

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════════════
ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
# Two-stage financial pipeline:
#   EXTRACTION  → Sonnet 4.6 at temperature=0. Deterministic. Just reads the
#                 documents and pulls clean numbers. Same input → same output.
#   METHODOLOGY → Opus 4.8 (effort=high by default). Judgment-heavy: GGC
#                 categorization, collections, POH bifurcation, taxes. Opus 4.8
#                 flags its own uncertainty better, which matters for underwriting.
# NOTE: Opus 4.8 (like 4.7) does NOT accept temperature/top_p/top_k — only
# adaptive thinking. So temperature=0 is only used on the Sonnet extraction call.
#
# Pin model versions explicitly. Bare aliases like "claude-sonnet-4-6" silently
# re-point to new snapshots over time — a hidden source of run-to-run variance.
# Override via env once a dated snapshot is confirmed in the Anthropic docs.
MODEL_EXTRACTION  = os.environ.get("MODEL_EXTRACTION",  "claude-sonnet-4-6")
MODEL_METHODOLOGY = os.environ.get("MODEL_METHODOLOGY", "claude-opus-4-8")
MODEL_MARKET      = os.environ.get("MODEL_MARKET",      "claude-opus-4-8")
API_VERSION       = "2023-06-01"
MAX_TOKENS             = 32000  # default; safe for Sonnet 4.6 (64k cap) and non-thinking calls
# Opus 4.8 with adaptive thinking + effort=high spends most of the budget on
# thinking — a complex deal can burn 30-50k thinking tokens before emitting
# any visible JSON. max_tokens is the COMBINED ceiling for thinking + output,
# so the methodology + market stages need much more headroom than extraction.
# Opus supports up to 128k output via streaming (which we already use).
MAX_TOKENS_METHODOLOGY = 96000  # ~80k thinking headroom + ~16k for JSON
MAX_TOKENS_MARKET      = 64000  # thinking + web_search results + comp tables
MAX_RETRIES       = 6
BASE_BACKOFF_SEC  = 2

# Anthropic Structured Outputs is now GA (verified against
# platform.claude.com/docs/en/build-with-claude/structured-outputs as of
# 2026-06). When a JSON Schema is passed to call_claude(), it's compiled into
# a token grammar and invalid tokens are masked at inference — guarantees
# schema compliance and eliminates the class of variance where Claude emits
# the wrong field types / out-of-enum categories / malformed JSON.
# Confirmed supported (GA list): Sonnet 4.5/4.6, Opus 4.5/4.6/4.7/4.8,
# Haiku 4.5, Mythos Preview. No beta header required.
# STRUCTURED_OUTPUTS_BETA is kept for backward compatibility with deployments
# pinned to a model snapshot where only the beta path is recognized.
STRUCTURED_OUTPUTS_BETA = "structured-outputs-2025-11-13"
USE_STRUCTURED_OUTPUTS  = os.environ.get("USE_STRUCTURED_OUTPUTS", "1") == "1"
# Default to off — GA path doesn't need it. Flip on if your pinned snapshot
# rejects output_config.format without the legacy beta header.
SO_LEGACY_BETA_HEADER   = os.environ.get("SO_LEGACY_BETA_HEADER", "0") == "1"

# Retries when verify_extraction returns any status=fail. Each retry feeds the
# failure messages back to Claude (Instructor pattern).
MAX_PARSE_RETRIES = int(os.environ.get("MAX_PARSE_RETRIES", "2"))

# N parallel extraction runs to field-merge when deep_search=on (Wang et al.
# self-consistency). 1 disables it. Costs ~N× extraction tokens.
FINANCIAL_PARSE_RUNS_DEEP = int(os.environ.get("FINANCIAL_PARSE_RUNS_DEEP", "3"))

# Versioned extraction cache: (file content + property_info + config) hash →
# the full financial_pipeline output. A hit returns the exact same numbers
# every time for previously-seen inputs. Lives in EXTRACTION_CACHE_DIR.
EXTRACTION_CACHE_ENABLED = os.environ.get("EXTRACTION_CACHE_ENABLED", "1") == "1"

# Parser backend: which OCR/layout engine converts PDFs to text for Claude.
# Options: "docai" (Google, current), "azure" (Document Intelligence Layout,
# deterministic), "tensorlake" (vision-first, 91.7% F1 per vendor),
# "reducto" (vision-first, 90.2% on RD-TableBench per vendor).
# Per the 2026 playbook, Doc AI scores 64.6% on complex tables — swap once a
# 50-doc benchmark confirms a better fit. Backend name + version go into the
# extraction cache key, so swapping invalidates the cache.
PARSER_BACKEND = os.environ.get("PARSER_BACKEND", "docai").lower()
PARSER_VERSION = "v1"  # bump when changing parser configuration

GOOGLE_MAPS_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "")
GOOGLE_STATIC_MAPS_URL       = "https://maps.googleapis.com/maps/api/staticmap"
GOOGLE_STATIC_STREETVIEW_URL = "https://maps.googleapis.com/maps/api/streetview"

# ─────────────────────────────────────────────────────────────────────────────
# DEFAULT ANTHROPIC KEY (optional)
# Paste your key here so you don't have to enter it in the UI every time.
# Leave as empty string "" if you want to type it in the UI manually.
# Example: DEFAULT_ANTHROPIC_KEY = "sk-ant-api03-AbCd1234..."
# ─────────────────────────────────────────────────────────────────────────────
load_dotenv()

DEFAULT_ANTHROPIC_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
GOOGLE_MAPS_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "")

GCP_PROJECT_ID         = os.environ.get("GCP_PROJECT_ID", "")
GCP_LOCATION           = os.environ.get("GCP_LOCATION", "us")
GCP_LAYOUT_PROCESSOR_ID = os.environ.get("GCP_LAYOUT_PROCESSOR_ID", "")
DOC_AI_ENABLED = bool(GCP_PROJECT_ID and GCP_LAYOUT_PROCESSOR_ID)

# Alternate parser backends (Phase 2 of the 2026 playbook). Set the env vars
# for whichever backend you select in PARSER_BACKEND.
AZURE_DOC_INTEL_ENDPOINT = os.environ.get("AZURE_DOC_INTEL_ENDPOINT", "")
AZURE_DOC_INTEL_KEY      = os.environ.get("AZURE_DOC_INTEL_KEY", "")
TENSORLAKE_API_KEY       = os.environ.get("TENSORLAKE_API_KEY", "")
REDUCTO_API_KEY          = os.environ.get("REDUCTO_API_KEY", "")

# IMPORTANT: GGC's official blank template, extended to 1000 rent roll rows
TEMPLATE_PATH        = Path(__file__).parent / "GGC_Blank_Underwriting_Sizer_Extended.xlsx"
JOBS_DIR             = Path(__file__).parent.parent / "jobs"
IMG_CACHE_DIR        = Path(__file__).parent.parent / "img_cache"
EXTRACTION_CACHE_DIR = Path(__file__).parent.parent / "extraction_cache"
JOBS_DIR.mkdir(exist_ok=True)
IMG_CACHE_DIR.mkdir(exist_ok=True)
EXTRACTION_CACHE_DIR.mkdir(exist_ok=True)

# GGC's exact category strings — must match column A in Data Consolidation
# (these feed the SUMIFS in the GGC Underwriting tab)
GGC_INCOME_CATEGORIES = [
    "Gross Potential Rent", "Less: Vacancy", "Less: Concessions", "Less: Bad Debt",
    "Utility Reimbursement", "Home Rent Income", "RV Site Rental Income",
    "Storage Income", "Retail Income", "Other Income", "Employee Allowance",
    "Model Units",
]

GGC_EXPENSE_CATEGORIES = [
    "RE Taxes", "Insurance", "Gas/Fuel", "Electrcitiy",  # GGC's spelling
    "Water and Sewer", "Trash Removal", "Repair and Maintenance",
    "Ground Maintenance", "Recreational Amenities", "Management Fee",
    "Payroll", "General and Administrative", "Professional Fees",
    "Advertising", "Home Rent Expense (MH)", "Other", "Cap-Ex Reserve",
]

# Canonical unit-type taxonomy. Unit Mix Summary COUNTIFS/SUMIFS in the
# template key on these exact strings — any drift breaks unit counts,
# which cascades to # of Units (Underwriting!N7) and every per-unit
# expense formula downstream.
CANONICAL_UNIT_TYPES = [
    "TOH MH Site",         # Tenant-owned manufactured-home lots
    "POH-Infilled units",  # Park-owned home rentals
    "Long term RV Site",   # Annual RV lots
    "Retail/Commercial",   # Storefronts, commercial space, storage
]

import re
import secrets
import threading
from collections import OrderedDict

# Job state. Mutated from the request thread (insert) and the worker
# thread (status / progress / result writes); read from /api/status.
# CPython dict ops are atomic, but read-modify-write on nested values is
# not, so all mutations go through JOBS_LOCK. OrderedDict + LRU
# eviction keeps memory and disk bounded.
JOBS = OrderedDict()
JOBS_LOCK = threading.Lock()
JOBS_MAX = 50  # cap concurrent + recent finished jobs in memory

def _evict_old_jobs():
    """Remove the oldest completed/errored jobs (and their .xlsx files)
    once we exceed JOBS_MAX. Running/queued jobs are skipped — we never
    evict work in flight. Caller MUST already hold JOBS_LOCK."""
    if len(JOBS) <= JOBS_MAX:
        return
    overflow = len(JOBS) - JOBS_MAX
    victims = []
    for jid, job in JOBS.items():
        if job.get("status") in ("complete", "error"):
            victims.append(jid)
            if len(victims) >= overflow:
                break
    for jid in victims:
        JOBS.pop(jid, None)
        try:
            (JOBS_DIR / f"{jid}.xlsx").unlink(missing_ok=True)
        except Exception:
            pass

# Upload guard rails. Catches obvious abuse before parsing.
MAX_UPLOAD_BYTES = 150 * 1024 * 1024  # 150 MB total per request
ALLOWED_UPLOAD_EXTS = {
    ".pdf", ".xlsx", ".xls", ".csv",
    ".png", ".jpg", ".jpeg",
    ".txt", ".md",
}

# Job IDs are 256-bit random tokens — not enumerable by timestamp.
# The validation pattern below is also used as a path-traversal guard.
JOB_ID_RE = re.compile(r"^[A-Za-z0-9_-]{16,64}$")

def _new_job_id():
    return secrets.token_urlsafe(24)

def _valid_job_id(job_id):
    return isinstance(job_id, str) and bool(JOB_ID_RE.match(job_id))

app = Flask(__name__, static_folder=".", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES
CORS(app)

# ═══════════════════════════════════════════════════════════════════════════
# GOOGLE MAPS HELPERS
# ═══════════════════════════════════════════════════════════════════════════
def fetch_google_static_map(address, map_type="satellite", zoom=17, size="600x400"):
    if not GOOGLE_MAPS_API_KEY:
        return None

    geo = geocode_address(address)
    if not geo:
        # Fallback to address-based
        center_param = address
        cache_key = f"map_{abs(hash(address))}_{map_type}_{zoom}_{size}"
    else:
        center_param = f"{geo['lat']},{geo['lng']}"
        cache_key = f"map_{geo['lat']:.6f}_{geo['lng']:.6f}_{map_type}_{zoom}_{size}"

    cache_path = IMG_CACHE_DIR / f"{cache_key}.png"
    if cache_path.exists():
        return str(cache_path)

    params = {
        "center": center_param,
        "zoom": zoom,
        "size": size,
        "maptype": map_type,
        "markers": f"color:red|{center_param}",
        "key": GOOGLE_MAPS_API_KEY,
    }
    try:
        r = requests.get(GOOGLE_STATIC_MAPS_URL, params=params, timeout=15)
        if r.status_code == 200 and r.headers.get("content-type", "").startswith("image/"):
            cache_path.write_bytes(r.content)
            return str(cache_path)
    except Exception as e:
        print(f"[GoogleMaps] Static map error: {e}")
    return None


GOOGLE_STREETVIEW_METADATA_URL = "https://maps.googleapis.com/maps/api/streetview/metadata"

def fetch_google_streetview(address, heading=0, pitch=0, fov=90, size="600x400",
                             radius=200):
    """
    Fetch a Street View image, intelligently handling cases where the address
    doesn't sit on a road with panorama coverage.

    Strategy:
    1. Geocode the address to lat/lng (with centroid offset for MHCs)
    2. Query the Street View Metadata API to find the nearest panorama
       within `radius` meters
    3. If found, request the image using the panorama's actual lat/lng
       and compute a heading that points BACK TOWARD the property
    4. If no panorama within radius, fall back to address-based request
    """
    if not GOOGLE_MAPS_API_KEY:
        return None

    geo = geocode_address(address)
    if not geo:
        return _fetch_streetview_by_address(address, heading, pitch, fov, size)

    cache_key = f"sv_{geo['lat']:.6f}_{geo['lng']:.6f}_{heading}_{pitch}_{fov}_{size}"
    cache_path = IMG_CACHE_DIR / f"{cache_key}.png"
    if cache_path.exists():
        return str(cache_path)

    # Find the closest panorama within `radius` meters of the geocoded point
    try:
        meta = requests.get(GOOGLE_STREETVIEW_METADATA_URL,
                             params={
                                 "location": f"{geo['lat']},{geo['lng']}",
                                 "radius": radius,
                                 "source": "outdoor",
                                 "key": GOOGLE_MAPS_API_KEY,
                             }, timeout=10).json()
    except Exception as e:
        print(f"[StreetView] Metadata fetch failed: {e}")
        meta = {"status": "ERROR"}

    if meta.get("status") == "OK":
        # The panorama exists. Get the panorama's actual location, then
        # compute a heading that points from the panorama AT the property
        pano_lat = meta["location"]["lat"]
        pano_lng = meta["location"]["lng"]
        bearing_to_property = _bearing(pano_lat, pano_lng, geo["lat"], geo["lng"])

        # The `heading` arg lets the caller request "looking left", "looking
        # right" etc. Treat it as an offset from the property-facing bearing.
        effective_heading = (bearing_to_property + heading) % 360

        params = {
            "location": f"{pano_lat},{pano_lng}",
            "size": size,
            "heading": effective_heading,
            "pitch": pitch,
            "fov": fov,
            "key": GOOGLE_MAPS_API_KEY,
            "source": "outdoor",
        }
        try:
            r = requests.get(GOOGLE_STATIC_STREETVIEW_URL, params=params, timeout=15)
            if r.status_code == 200 and r.headers.get("content-type", "").startswith("image/") and len(r.content) > 5000:
                cache_path.write_bytes(r.content)
                return str(cache_path)
        except Exception as e:
            print(f"[StreetView] Image fetch failed: {e}")

    # Fallback: address-based, original behavior — better than nothing
    return _fetch_streetview_by_address(address, heading, pitch, fov, size)


def _fetch_streetview_by_address(address, heading, pitch, fov, size):
    """Original fallback path — used when geocoding or metadata fails."""
    cache_key = f"sv_addr_{abs(hash(address))}_{heading}_{pitch}_{fov}_{size}"
    cache_path = IMG_CACHE_DIR / f"{cache_key}.png"
    if cache_path.exists():
        return str(cache_path)
    params = {"location": address, "size": size, "heading": heading,
              "pitch": pitch, "fov": fov, "key": GOOGLE_MAPS_API_KEY,
              "source": "outdoor"}
    try:
        r = requests.get(GOOGLE_STATIC_STREETVIEW_URL, params=params, timeout=15)
        if r.status_code == 200 and r.headers.get("content-type", "").startswith("image/"):
            if len(r.content) < 5000:
                return None
            cache_path.write_bytes(r.content)
            return str(cache_path)
    except Exception as e:
        print(f"[StreetView] Error: {e}")
    return None


def _bearing(lat1, lng1, lat2, lng2):
    """Calculate compass bearing in degrees from point 1 to point 2."""
    import math
    lat1_r, lat2_r = math.radians(lat1), math.radians(lat2)
    diff = math.radians(lng2 - lng1)
    y = math.sin(diff) * math.cos(lat2_r)
    x = math.cos(lat1_r) * math.sin(lat2_r) - math.sin(lat1_r) * math.cos(lat2_r) * math.cos(diff)
    return (math.degrees(math.atan2(y, x)) + 360) % 360


def embed_image_in_cell(ws, image_path, anchor_cell, width_px=400, height_px=280):
    if not image_path or not Path(image_path).exists():
        return False
    try:
        img = XLImage(image_path)
        img.width = width_px
        img.height = height_px
        img.anchor = anchor_cell
        ws.add_image(img)
        return True
    except Exception as e:
        print(f"[Excel] Image embed failed: {e}")
        return False


def to_decimal_pct(value):
    """Coerce a percentage-ish value to its DECIMAL form for Excel cells
    formatted as ``0.00%``. The LLM is inconsistent: it may return 7.5 to
    mean 7.5% (percent form) or 0.075 (decimal form). Excel's percent format
    multiplies by 100, so a percent-form 7.5 would render as 750%.

    Heuristic: anything with abs(v) < 1.5 is already decimal (0.075 → 0.075).
    Anything larger is presumed percent form and divided by 100 (7.5 → 0.075).
    Returns None if the value can't be parsed.
    """
    if value is None or value == "":
        return None
    if isinstance(value, str):
        s = value.strip().rstrip("%").strip()
        try:
            value = float(s)
        except (TypeError, ValueError):
            return None
    try:
        v = float(value)
    except (TypeError, ValueError):
        return None
    return v if abs(v) < 1.5 else v / 100


def safe_pct(value):
    """
    Format a value as a percentage string, regardless of whether Claude returned:
      - None              → ""
      - a number 0.045    → "4.5%"
      - a number 4.5      → "4.5%"  (already in percent form)
      - a string "4.5%"   → "4.5%"  (passthrough)
      - a string "4.5"    → "4.5%"
    """
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        s = value.strip()
        if s.endswith("%"):
            return s
        try:
            value = float(s)
        except ValueError:
            return s  # not a number, return as-is
    try:
        v = float(value)
    except (TypeError, ValueError):
        return str(value)
    # Heuristic: if absolute value is < 1, treat as decimal fraction (0.045 → 4.5%)
    # Otherwise treat as already in percent form (4.5 → 4.5%)
    if abs(v) < 1:
        v = v * 100
    return f"{v:.1f}%"


def safe_money(value, suffix=""):
    """Format a value as $X,XXX. Handles strings like '$78,000' or numbers or None."""
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        s = value.strip().replace("$", "").replace(",", "").replace("/mo", "").strip()
        if s.endswith("%"):
            return value  # weird input, return as-is
        try:
            value = float(s)
        except ValueError:
            return value  # not numeric, return as-is
    try:
        return f"${float(value):,.0f}{suffix}"
    except (TypeError, ValueError):
        return str(value)

# ═══════════════════════════════════════════════════════════════════════════
# PDF PARSER ABSTRACTION
# Single entry point `parse_pdf(bytes, filename) → markdown | None`.
# Dispatches by PARSER_BACKEND. Parser name + version are baked into the
# cache filename so swapping backends or bumping config invalidates stale
# parses (the file SHA stays stable so identical bytes hit the cache).
#
# Backends:
#   docai      — Google Document AI Layout Parser (current default).
#                Per RD-TableBench Nov 2024, scored 64.6% on complex tables
#                — lowest of the major cloud providers.
#   azure      — Azure Document Intelligence Layout. ~82.7% on the same
#                benchmark; deterministic given identical input.
#   tensorlake — Vision-first parser. 91.7% F1 per vendor (Nov 2025).
#   reducto    — Vision-first parser. 90.2% on RD-TableBench per vendor.
# ═══════════════════════════════════════════════════════════════════════════

def _stable_pdf_hash(pdf_bytes):
    return hashlib.sha256(pdf_bytes).hexdigest()[:16]


def parse_pdf(pdf_bytes, filename):
    """
    Parse a PDF using whichever backend PARSER_BACKEND names. Returns the
    markdown string on success, None on failure (caller falls back to
    Anthropic's native PDF handling).
    """
    backend = PARSER_BACKEND
    cache_key  = f"{backend}_{PARSER_VERSION}_{_stable_pdf_hash(pdf_bytes)}"
    cache_path = IMG_CACHE_DIR / f"{cache_key}.md"
    if cache_path.exists():
        print(f"[Parser:{backend}] Cache hit for {filename}")
        return cache_path.read_text()

    try:
        if backend == "docai":
            markdown = _parse_via_docai(pdf_bytes, filename)
        elif backend == "azure":
            markdown = _parse_via_azure(pdf_bytes, filename)
        elif backend == "tensorlake":
            markdown = _parse_via_tensorlake(pdf_bytes, filename)
        elif backend == "reducto":
            markdown = _parse_via_reducto(pdf_bytes, filename)
        else:
            print(f"[Parser] Unknown PARSER_BACKEND='{backend}' — "
                  "supported: docai, azure, tensorlake, reducto")
            return None
    except Exception as e:
        print(f"[Parser:{backend}] Failed for {filename}: {e}")
        traceback.print_exc()
        return None

    if markdown:
        # Diagnostic hash — if this changes across runs of the same file,
        # the parser itself is non-deterministic (and Claude variance has
        # nothing to do with run-to-run drift)
        md_hash = hashlib.md5(markdown.encode()).hexdigest()[:12]
        print(f"[Parser:{backend}] Markdown hash for {filename}: "
              f"{md_hash} ({len(markdown)} chars)")
        cache_path.write_text(markdown)
    return markdown


# Backward-compat alias — encode_file_for_claude callers and any external
# wrappers expecting the old name still work.
parse_pdf_with_document_ai = parse_pdf


def _parse_via_docai(pdf_bytes, filename):
    """Google Document AI Layout Parser — markdown with tables/headers."""
    if not DOC_AI_ENABLED:
        print("[Parser:docai] GCP_PROJECT_ID / GCP_LAYOUT_PROCESSOR_ID not "
              "set — skipping Doc AI")
        return None

    opts = ClientOptions(api_endpoint=f"{GCP_LOCATION}-documentai.googleapis.com")
    client = documentai.DocumentProcessorServiceClient(client_options=opts)
    processor_name = (
        f"projects/{GCP_PROJECT_ID}/locations/{GCP_LOCATION}"
        f"/processors/{GCP_LAYOUT_PROCESSOR_ID}"
    )
    raw_document = documentai.RawDocument(content=pdf_bytes, mime_type="application/pdf")
    process_options = documentai.ProcessOptions(
        layout_config=documentai.ProcessOptions.LayoutConfig(
            chunking_config=documentai.ProcessOptions.LayoutConfig.ChunkingConfig(
                chunk_size=1000,
                include_ancestor_headings=True,
            )
        )
    )
    request = documentai.ProcessRequest(
        name=processor_name,
        raw_document=raw_document,
        process_options=process_options,
    )

    print(f"[Parser:docai] Parsing {filename} ({len(pdf_bytes)//1024} KB)...")
    t0 = time.time()
    result = client.process_document(request=request)
    print(f"[Parser:docai] Parsed in {time.time() - t0:.1f}s")

    doc = result.document
    parts = [f"# Document: {filename}\n"]
    if doc.chunked_document and doc.chunked_document.chunks:
        for chunk in doc.chunked_document.chunks:
            if chunk.page_headers:
                for hdr in chunk.page_headers:
                    parts.append(f"\n## {hdr.text}\n")
            parts.append(chunk.content)
            parts.append("\n")
    else:
        parts.append(doc.text)
    return "\n".join(parts)


def _parse_via_azure(pdf_bytes, filename):
    """
    Azure Document Intelligence Layout (prebuilt-layout model) — returns
    markdown when outputContentFormat=markdown is requested. Async: POST
    returns 202 with an operation-location header; poll until done.

    Verified 2026-06 against
      learn.microsoft.com/en-us/azure/ai-services/document-intelligence/
    REST API v4.0 (api-version=2024-11-30) is the current GA. v4.0 changed
    table representation to HTML inside the markdown stream to support
    merged cells + multirow headers, which is desirable for T12s.
    """
    if not AZURE_DOC_INTEL_ENDPOINT or not AZURE_DOC_INTEL_KEY:
        print("[Parser:azure] AZURE_DOC_INTEL_ENDPOINT / AZURE_DOC_INTEL_KEY "
              "not set — skipping")
        return None

    base = AZURE_DOC_INTEL_ENDPOINT.rstrip("/")
    api_version = os.environ.get("AZURE_DOC_INTEL_API_VERSION", "2024-11-30")
    start_url = (
        f"{base}/documentintelligence/documentModels/"
        f"prebuilt-layout:analyze"
        f"?api-version={api_version}&outputContentFormat=markdown"
    )
    headers = {
        "Ocp-Apim-Subscription-Key": AZURE_DOC_INTEL_KEY,
        "Content-Type": "application/pdf",
    }

    print(f"[Parser:azure] Parsing {filename} ({len(pdf_bytes)//1024} KB)...")
    t0 = time.time()
    r = requests.post(start_url, headers=headers, data=pdf_bytes, timeout=120)
    if r.status_code not in (200, 202):
        print(f"[Parser:azure] start failed: {r.status_code} {r.text[:300]}")
        return None
    op_url = r.headers.get("operation-location") or r.headers.get("Operation-Location")
    if not op_url:
        print(f"[Parser:azure] no operation-location header: {dict(r.headers)}")
        return None

    poll_headers = {"Ocp-Apim-Subscription-Key": AZURE_DOC_INTEL_KEY}
    deadline = time.time() + 300
    while time.time() < deadline:
        time.sleep(2)
        pr = requests.get(op_url, headers=poll_headers, timeout=30)
        if pr.status_code != 200:
            print(f"[Parser:azure] poll {pr.status_code}: {pr.text[:200]}")
            return None
        data = pr.json()
        status = data.get("status")
        if status == "succeeded":
            print(f"[Parser:azure] Parsed in {time.time() - t0:.1f}s")
            content = data.get("analyzeResult", {}).get("content", "")
            return f"# Document: {filename}\n\n{content}"
        if status in ("failed", "canceled"):
            print(f"[Parser:azure] terminal status {status}: {data}")
            return None
    print("[Parser:azure] polling timed out (5 min)")
    return None


def _parse_via_tensorlake(pdf_bytes, filename):
    """
    Tensorlake document parsing (v2 async) per docs.tensorlake.ai/api-reference
    (verified 2026-06):
      1) POST /documents/v2/files (multipart) → {"file_id": "..."}
         NOTE: the upload endpoint URL is the one piece NOT explicitly shown
         in the public docs preview — verify against your account's actual
         endpoint or use the official Tensorlake Python SDK for upload.
         Override via TENSORLAKE_UPLOAD_PATH if needed.
      2) POST /documents/v2/parse with {"file_id", "mime_type",
         "parsing_options": {"table_output_mode": "markdown"}} → {"parse_id"}
      3) GET /documents/v2/parse/{parse_id} until status="successful"
      4) Markdown is the joined chunks[].content
    Pricing: $0.01/page flat ($10/1k pages).
    """
    if not TENSORLAKE_API_KEY:
        print("[Parser:tensorlake] TENSORLAKE_API_KEY not set — skipping")
        return None

    base = os.environ.get("TENSORLAKE_BASE_URL", "https://api.tensorlake.ai").rstrip("/")
    upload_path = os.environ.get("TENSORLAKE_UPLOAD_PATH", "/documents/v2/files")
    auth = {"Authorization": f"Bearer {TENSORLAKE_API_KEY}"}

    # Step 1 — upload, get file_id
    print(f"[Parser:tensorlake] Uploading {filename} "
          f"({len(pdf_bytes)//1024} KB)...")
    t0 = time.time()
    up = requests.post(f"{base}{upload_path}", headers=auth,
                       files={"file": (filename, pdf_bytes, "application/pdf")},
                       timeout=300)
    if up.status_code not in (200, 201):
        print(f"[Parser:tensorlake] upload {up.status_code}: {up.text[:300]} "
              f"— if 404/405 the upload endpoint differs; set "
              "TENSORLAKE_UPLOAD_PATH or use the Tensorlake SDK")
        return None
    try:
        file_id = (up.json().get("file_id") or up.json().get("id"))
    except json.JSONDecodeError:
        print(f"[Parser:tensorlake] upload non-JSON: {up.text[:200]}")
        return None
    if not file_id:
        print(f"[Parser:tensorlake] upload missing file_id: {up.text[:200]}")
        return None

    # Step 2 — submit parse job
    body = {
        "file_id":   file_id,
        "mime_type": "application/pdf",
        "parsing_options": {"table_output_mode": "markdown"},
    }
    sub = requests.post(f"{base}/documents/v2/parse",
                        headers={**auth, "Content-Type": "application/json"},
                        json=body, timeout=120)
    if sub.status_code not in (200, 201, 202):
        print(f"[Parser:tensorlake] parse submit {sub.status_code}: {sub.text[:300]}")
        return None
    try:
        parse_id = sub.json().get("parse_id") or sub.json().get("id")
    except json.JSONDecodeError:
        print(f"[Parser:tensorlake] submit non-JSON: {sub.text[:200]}")
        return None
    if not parse_id:
        print(f"[Parser:tensorlake] submit missing parse_id: {sub.text[:200]}")
        return None

    # Step 3 — poll until successful (or timeout at 5 min)
    poll_url = f"{base}/documents/v2/parse/{parse_id}"
    deadline = time.time() + 300
    while time.time() < deadline:
        time.sleep(3)
        pr = requests.get(poll_url, headers=auth, timeout=30)
        if pr.status_code != 200:
            print(f"[Parser:tensorlake] poll {pr.status_code}: {pr.text[:200]}")
            return None
        data = pr.json()
        status = (data.get("status") or "").lower()
        if status in ("successful", "succeeded", "completed", "complete"):
            chunks = (data.get("chunks") or data.get("result", {}).get("chunks")
                      or [])
            parts = [ch.get("content", "") for ch in chunks if ch.get("content")]
            if not parts:
                print(f"[Parser:tensorlake] no chunks in result; keys="
                      f"{list(data.keys())}")
                return None
            print(f"[Parser:tensorlake] Parsed in {time.time() - t0:.1f}s "
                  f"({len(parts)} chunks)")
            return f"# Document: {filename}\n\n" + "\n\n".join(parts)
        if status in ("failed", "error", "errored", "canceled"):
            print(f"[Parser:tensorlake] terminal status {status}: {data}")
            return None
    print("[Parser:tensorlake] polling timed out (5 min)")
    return None


def _parse_via_reducto(pdf_bytes, filename):
    """
    Reducto vision-first parser. Two-step synchronous flow per
    docs.reducto.ai/quickstart (verified 2026-06):
      1) POST /upload (multipart) → {"file_id": "reducto://..."}
      2) POST /parse (JSON, input=file_id) → {"result": {"chunks":[{"content":...}]}}
    Markdown is the concatenated chunks[].content. Published pricing
    starts at $0.015/page.
    """
    if not REDUCTO_API_KEY:
        print("[Parser:reducto] REDUCTO_API_KEY not set — skipping")
        return None

    base = os.environ.get("REDUCTO_BASE_URL", "https://platform.reducto.ai").rstrip("/")
    auth = {"Authorization": f"Bearer {REDUCTO_API_KEY}"}

    # Step 1 — upload bytes, get file_id
    print(f"[Parser:reducto] Uploading {filename} "
          f"({len(pdf_bytes)//1024} KB)...")
    t0 = time.time()
    up = requests.post(f"{base}/upload", headers=auth,
                       files={"file": (filename, pdf_bytes, "application/pdf")},
                       timeout=300)
    if up.status_code not in (200, 201):
        print(f"[Parser:reducto] upload {up.status_code}: {up.text[:300]}")
        return None
    try:
        file_id = up.json().get("file_id")
    except json.JSONDecodeError:
        print(f"[Parser:reducto] upload returned non-JSON: {up.text[:200]}")
        return None
    if not file_id:
        print(f"[Parser:reducto] upload response missing file_id: {up.text[:200]}")
        return None

    # Step 2 — parse with markdown table formatting
    body = {"input": file_id, "formatting": {"table_output_format": "md"}}
    pr = requests.post(f"{base}/parse",
                       headers={**auth, "Content-Type": "application/json"},
                       json=body, timeout=600)
    if pr.status_code not in (200, 201):
        print(f"[Parser:reducto] parse {pr.status_code}: {pr.text[:300]}")
        return None

    try:
        payload = pr.json()
    except json.JSONDecodeError:
        print(f"[Parser:reducto] parse non-JSON: {pr.text[:200]}")
        return None

    chunks = (payload.get("result") or {}).get("chunks") or payload.get("chunks") or []
    parts = []
    for ch in chunks:
        # Per docs, ch.content is the full text; ch.blocks[] is per-element.
        # Prefer content; fall back to joined block contents.
        txt = ch.get("content")
        if not txt and ch.get("blocks"):
            txt = "\n\n".join(b.get("content", "") for b in ch["blocks"]
                              if b.get("content"))
        if txt:
            parts.append(txt)

    if not parts:
        print(f"[Parser:reducto] parse response had no chunks: "
              f"keys={list(payload.keys())}")
        return None

    print(f"[Parser:reducto] Parsed in {time.time() - t0:.1f}s "
          f"({payload.get('usage', {}).get('num_pages', '?')} pages)")
    return f"# Document: {filename}\n\n" + "\n\n".join(parts)

# ═══════════════════════════════════════════════════════════════════════════
# VERSIONED EXTRACTION CACHE
# Key = SHA-256 of (file_blocks + property_info + model pins + parser config +
# prompt hashes + schema hashes + n_runs). On a hit, returns the exact same
# financial_pipeline output that was produced last time the same inputs were
# seen. This is the operational fix for the "same PDF produces different
# output" complaint — fingerprinted memoization.
# Bumping any prompt, schema, or model invalidates all entries automatically
# (their hash changes). To force-clear, delete EXTRACTION_CACHE_DIR.
# ═══════════════════════════════════════════════════════════════════════════

def _hash_obj(obj):
    """Stable SHA-256 hex digest of any JSON-serializable object."""
    return hashlib.sha256(
        json.dumps(obj, sort_keys=True, default=str).encode()
    ).hexdigest()


def extraction_cache_key(file_blocks, property_info, n_extraction_runs,
                         extraction_prompt, methodology_prompt):
    """
    Build the cache key tuple, then hash it. Anything that changes the output
    must be in here. PARSER_BACKEND is in via PARSER_VERSION + the fact that
    PARSER_BACKEND affects what's *in* file_blocks (the parser's markdown is
    embedded directly).
    """
    key_obj = {
        "files":        _hash_obj(file_blocks),
        "property":     _hash_obj(property_info),
        "model_x":      MODEL_EXTRACTION,
        "model_m":      MODEL_METHODOLOGY,
        "use_so":       USE_STRUCTURED_OUTPUTS,
        "parser":       PARSER_BACKEND,
        "parser_v":     PARSER_VERSION,
        "ext_prompt":   _hash_obj(extraction_prompt)[:16],
        "meth_prompt":  _hash_obj(methodology_prompt)[:16],
        "ext_schema":   _hash_obj(EXTRACTION_OUTPUT_SCHEMA)[:16],
        "meth_schema":  _hash_obj(METHODOLOGY_OUTPUT_SCHEMA)[:16],
        "n_runs":       n_extraction_runs,
    }
    return _hash_obj(key_obj)[:32]


def extraction_cache_get(key):
    if not EXTRACTION_CACHE_ENABLED:
        return None
    path = EXTRACTION_CACHE_DIR / f"{key}.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text())
    except Exception as e:
        print(f"[ExtractionCache] read failed for {key[:8]}: {e}")
        return None


def extraction_cache_put(key, payload):
    if not EXTRACTION_CACHE_ENABLED:
        return
    path = EXTRACTION_CACHE_DIR / f"{key}.json"
    try:
        path.write_text(json.dumps(payload))
    except Exception as e:
        print(f"[ExtractionCache] write failed for {key[:8]}: {e}")


GOOGLE_GEOCODE_URL = "https://maps.googleapis.com/maps/api/geocode/json"

def geocode_address(address):
    """
    Resolve an address to its best lat/lng, using the property centroid
    when the geocoder returns a bounding box (common for MHCs and complexes).
    Returns dict with 'lat', 'lng', 'precision', and 'place_id'.
    """
    if not GOOGLE_MAPS_API_KEY or not address:
        return None

    cache_key = f"geo_{abs(hash(address))}"
    cache_path = IMG_CACHE_DIR / f"{cache_key}.json"
    if cache_path.exists():
        try:
            return json.loads(cache_path.read_text())
        except Exception:
            pass

    try:
        r = requests.get(GOOGLE_GEOCODE_URL,
                          params={"address": address, "key": GOOGLE_MAPS_API_KEY},
                          timeout=15)
        if r.status_code != 200:
            return None
        data = r.json()
        if data.get("status") != "OK" or not data.get("results"):
            print(f"[Geocode] No result for {address}: {data.get('status')}")
            return None

        result = data["results"][0]
        loc = result["geometry"]["location"]
        precision = result["geometry"].get("location_type", "UNKNOWN")

        # If the geocoder returns a bounding box (common for MHCs, apartment
        # complexes, business parks), prefer the box centroid over the snap
        # point — it's typically further inside the property
        bounds = result["geometry"].get("bounds") or result["geometry"].get("viewport")
        if bounds:
            ne = bounds["northeast"]
            sw = bounds["southwest"]
            centroid_lat = (ne["lat"] + sw["lat"]) / 2
            centroid_lng = (ne["lng"] + sw["lng"]) / 2
            # Only use centroid if it's reasonably close to the geocoder pin
            # (sanity check — sometimes viewports are huge city-level boxes)
            lat_diff = abs(centroid_lat - loc["lat"])
            lng_diff = abs(centroid_lng - loc["lng"])
            if lat_diff < 0.005 and lng_diff < 0.005:  # ~500m
                loc = {"lat": centroid_lat, "lng": centroid_lng}

        out = {
            "lat": loc["lat"],
            "lng": loc["lng"],
            "precision": precision,
            "place_id": result.get("place_id"),
            "formatted_address": result.get("formatted_address"),
        }
        cache_path.write_text(json.dumps(out))
        return out
    except Exception as e:
        print(f"[Geocode] Error for {address}: {e}")
        return None

# ═══════════════════════════════════════════════════════════════════════════
# CLAUDE API CLIENT
# ═══════════════════════════════════════════════════════════════════════════
def call_claude(api_key, system_prompt, user_content, tools=None,
                use_thinking=True, temperature=None, model=None,
                output_schema=None, max_tokens=None):
    """
    Call Claude with streaming enabled. Streaming keeps the connection alive
    during long-running requests (which can hit 3+ minutes when Claude is doing
    heavy thinking + web search) instead of timing out at the request level.

    Model routing:
    - Defaults to MODEL_MARKET (Opus 4.8) for market research with adaptive thinking
    - Pass model=MODEL_METHODOLOGY (Opus 4.8) for GGC categorization + methodology
    - Pass model=MODEL_EXTRACTION (Sonnet 4.6) with use_thinking=False, temperature=0
      for deterministic document extraction

    NOTE: Opus 4.7 and 4.8 deprecated temperature/top_p/top_k entirely. Only pass
    temperature when targeting Sonnet 4.6 (or other pre-4.7 models). Opus 4.8
    defaults to effort=high, which is what we want for judgment-heavy work.
    """
    headers = {"x-api-key": api_key, "anthropic-version": API_VERSION,
               "content-type": "application/json"}
    # Wrap the system prompt in Anthropic's prompt-cache marker. The
    # FINANCIAL_PARSE_PROMPT is ~50K tokens of static methodology that
    # changes only when the codebase is redeployed; cached input is billed
    # at ~10% of the normal rate after the first hit (5-minute TTL). This
    # cuts the dominant Opus call's input cost ~80-90% on cache hits.
    # Strings shorter than the cache-eligibility floor are passed through
    # in their normal form to avoid wasting a cache slot.
    CACHE_MIN_TOKENS = 1024  # Anthropic's ephemeral cache floor
    if isinstance(system_prompt, str) and len(system_prompt) >= CACHE_MIN_TOKENS * 4:
        # ~4 chars/token heuristic — only cache prompts long enough to amortize.
        system_field = [{"type": "text", "text": system_prompt,
                         "cache_control": {"type": "ephemeral"}}]
    else:
        system_field = system_prompt

    body = {"model": model or MODEL_MARKET,
            "max_tokens": max_tokens or MAX_TOKENS,
            "system": system_field,
            "messages": [{"role": "user", "content": user_content}],
            "stream": True,}
    if use_thinking:
        body["thinking"] = {"type": "adaptive"}
        body["output_config"] = {"effort": "high"}
    elif temperature is not None:
        body["temperature"] = temperature

    # Structured outputs (GA). The schema is compiled into a token-level
    # grammar — Claude literally cannot emit a property outside it.
    # Composes with thinking by sharing output_config.
    use_so = bool(output_schema) and USE_STRUCTURED_OUTPUTS
    if use_so:
        # The GA path doesn't need a beta header; SO_LEGACY_BETA_HEADER opts
        # back in if a pinned snapshot still requires the transition header.
        if SO_LEGACY_BETA_HEADER:
            headers["anthropic-beta"] = STRUCTURED_OUTPUTS_BETA
        body.setdefault("output_config", {})["format"] = {
            "type": "json_schema",
            "schema": output_schema,
        }
    if tools:
        body["tools"] = tools

    last_err = None
    for attempt in range(MAX_RETRIES):
        try:
            # connect_timeout = 30s, read_timeout = 600s (10 min) per chunk
            # Streaming keeps each chunk's idle time low, so this is safe
            resp = requests.post(ANTHROPIC_API_URL, headers=headers,
                                 json=body, timeout=(30, 600), stream=True)
            if resp.status_code != 200:
                # If structured outputs is rejected for the pinned snapshot,
                # drop it and retry — prompt + Pydantic still enforce the
                # shape, just without grammar-level masking.
                if resp.status_code == 400 and use_so:
                    err_text = resp.text[:500].lower()
                    if any(t in err_text for t in ("structured", "beta", "output_config", "schema")):
                        print(f"[Claude] Structured outputs rejected for "
                              f"{body['model']} ({resp.status_code}: "
                              f"{resp.text[:160]}); falling back to "
                              f"prompt-only schema")
                        headers.pop("anthropic-beta", None)
                        if isinstance(body.get("output_config"), dict):
                            body["output_config"].pop("format", None)
                            if not body["output_config"]:
                                body.pop("output_config", None)
                        use_so = False
                        continue
                if resp.status_code in (429, 500, 502, 503, 504, 529):
                    retry_after = resp.headers.get("retry-after")
                    delay = int(retry_after) if retry_after else BASE_BACKOFF_SEC * (2 ** attempt)
                    print(f"[Claude] {resp.status_code} — retrying in {delay}s")
                    time.sleep(delay)
                    last_err = f"HTTP {resp.status_code}"
                    continue
                raise RuntimeError(f"Claude API error {resp.status_code}: {resp.text[:500]}")

            # Parse the SSE stream and reassemble the message
            return _parse_stream(resp)

        except requests.exceptions.Timeout:
            print(f"[Claude] Request timeout — retrying...")
            time.sleep(BASE_BACKOFF_SEC * (2 ** attempt))
            last_err = "Timeout"
        except requests.exceptions.ChunkedEncodingError as e:
            # Mid-stream connection drop — retry
            print(f"[Claude] Stream interrupted — retrying...")
            time.sleep(BASE_BACKOFF_SEC * (2 ** attempt))
            last_err = f"Stream interrupted: {e}"
        except requests.exceptions.ConnectionError as e:
            print(f"[Claude] Connection error (likely DNS/network) — retrying...")
            time.sleep(BASE_BACKOFF_SEC * (2 ** attempt))
            last_err = f"Connection error: {e}"
        except Exception as e:
            last_err = str(e)
            if attempt == MAX_RETRIES - 1:
                raise
            time.sleep(BASE_BACKOFF_SEC * (2 ** attempt))
    raise RuntimeError(f"Claude API failed after {MAX_RETRIES} retries: {last_err}")

def _parse_stream(resp):
    """
    Parse Anthropic's Server-Sent Events stream into the same dict shape
    that the non-streaming endpoint returns.
    """
    content_blocks = []   # accumulated content blocks
    current_block = None  # block currently being streamed
    stop_reason = None
    usage = {}

    for line in resp.iter_lines(decode_unicode=True):
        if not line or not line.startswith("data: "):
            continue
        try:
            event = json.loads(line[6:])
        except json.JSONDecodeError:
            continue

        etype = event.get("type")
        if etype == "content_block_start":
            current_block = dict(event.get("content_block", {}))
            content_blocks.append(current_block)
        elif etype == "content_block_delta":
            delta = event.get("delta", {})
            dtype = delta.get("type")
            if dtype == "text_delta" and current_block is not None:
                current_block["text"] = current_block.get("text", "") + delta.get("text", "")
            elif dtype == "input_json_delta" and current_block is not None:
                current_block["partial_json"] = current_block.get("partial_json", "") + delta.get("partial_json", "")
            elif dtype == "thinking_delta" and current_block is not None:
                current_block["thinking"] = current_block.get("thinking", "") + delta.get("thinking", "")
        elif etype == "content_block_stop":
            current_block = None
        elif etype == "message_delta":
            md = event.get("delta", {})
            if "stop_reason" in md:
                stop_reason = md["stop_reason"]
            if "usage" in event:
                usage.update(event["usage"])
        elif etype == "message_stop":
            break
        elif etype == "error":
            err = event.get("error", {})
            raise RuntimeError(f"Claude stream error: {err.get('type', '?')}: {err.get('message', '')}")

    return {"content": content_blocks, "stop_reason": stop_reason, "usage": usage}


def extract_text(response):
    return "\n".join(b.get("text", "") for b in response.get("content", []) if b.get("type") == "text")


def extract_json(response):
    text = extract_text(response)
    stop_reason = response.get("stop_reason", "")

    if stop_reason == "max_tokens":
        raise RuntimeError(
            "Claude hit max_tokens before finishing. The response was truncated. "
            "This usually means the deal has a very large rent roll. "
            "Try analyzing without an OM PDF, or split the analysis across two runs."
        )

    # Strip code fences if present
    text = re.sub(r"```(?:json)?\s*", "", text).replace("```", "")
    start, end = text.find("{"), text.rfind("}")
    if start == -1 or end == -1:
        raise ValueError(f"No JSON object found in response:\n{text[:500]}")

    candidate = text[start:end + 1]

    # Strategy 1: Try as-is
    try:
        return json.loads(candidate)
    except json.JSONDecodeError as e1:
        first_err = e1

    # Strategy 2: Escape unescaped control characters inside string values.
    # Claude sometimes inserts literal newlines/tabs inside strings, which is invalid JSON.
    # We walk the text character-by-character, tracking whether we're inside a quoted
    # string, and escape any raw control characters we encounter.
    def escape_control_chars(s):
        out = []
        in_string = False
        prev_was_backslash = False
        for ch in s:
            if in_string:
                if ch == '"' and not prev_was_backslash:
                    in_string = False
                    out.append(ch)
                elif ch == '\n':
                    out.append('\\n')
                elif ch == '\r':
                    out.append('\\r')
                elif ch == '\t':
                    out.append('\\t')
                elif ord(ch) < 0x20:
                    out.append(f'\\u{ord(ch):04x}')
                else:
                    out.append(ch)
            else:
                if ch == '"' and not prev_was_backslash:
                    in_string = True
                out.append(ch)
            prev_was_backslash = (ch == '\\') and not prev_was_backslash
        return ''.join(out)

    try:
        escaped = escape_control_chars(candidate)
        return json.loads(escaped)
    except json.JSONDecodeError:
        pass

    # Strategy 3: Strip trailing commas before } or ]
    cleaned = re.sub(r",(\s*[}\]])", r"\1", candidate)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    # Strategy 4: Combine — escape control chars AND strip trailing commas
    try:
        combo = re.sub(r",(\s*[}\]])", r"\1", escape_control_chars(candidate))
        return json.loads(combo)
    except json.JSONDecodeError:
        pass

    # Strategy 5: Truncate at the failure point and auto-close any open brackets/braces
    # This recovers a partial-but-valid JSON when Claude went off the rails mid-output
    try:
        truncate_at = first_err.pos
        partial = candidate[:truncate_at]
        depth_obj = partial.count("{") - partial.count("}")
        depth_arr = partial.count("[") - partial.count("]")
        partial = re.sub(r',\s*"[^"]*"?\s*:?\s*$', '', partial)
        partial = re.sub(r',\s*$', '', partial)
        partial = partial.rstrip()
        partial += "]" * max(0, depth_arr) + "}" * max(0, depth_obj)
        return json.loads(escape_control_chars(partial))
    except (json.JSONDecodeError, Exception):
        pass

    # All strategies failed — give a useful error with context around the failure point
    err_pos = first_err.pos
    context_start = max(0, err_pos - 200)
    context_end = min(len(candidate), err_pos + 200)
    context = candidate[context_start:err_pos] + " >>>HERE>>> " + candidate[err_pos:context_end]
    raise RuntimeError(
        f"Failed to parse JSON from Claude. Stop reason: {stop_reason}. "
        f"JSON error at position {err_pos}: {first_err.msg}.\n"
        f"Context around error:\n...{context}..."
    )


# ═══════════════════════════════════════════════════════════════════════════
# FILE ENCODING
# ═══════════════════════════════════════════════════════════════════════════
def encode_file_for_claude(file_storage):
    filename = file_storage.filename
    ext = Path(filename).suffix.lower()
    data = file_storage.read()
    b64 = base64.standard_b64encode(data).decode("utf-8")

    if ext == ".pdf":
        # Try the configured PARSER_BACKEND first for deterministic,
        # high-accuracy table extraction. Falls back to Anthropic's native
        # PDF handling if the parser is disabled, misconfigured, or fails.
        markdown = parse_pdf(data, filename)
        if markdown:
            return {"type": "text",
                    "text": f"[PDF parsed by {PARSER_BACKEND}: {filename}]\n\n{markdown[:200000]}"}
        print(f"[Encode] Parser {PARSER_BACKEND} unavailable; "
              f"falling back to Anthropic native PDF for {filename}")
        return {"type": "document",
                "source": {"type": "base64", "media_type": "application/pdf", "data": b64},
                "title": filename}
    elif ext in (".png", ".jpg", ".jpeg", ".webp", ".gif"):
        media_type = f"image/{'jpeg' if ext == '.jpg' else ext[1:]}"
        return {"type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": b64}}
    elif ext in (".xlsx", ".xls"):
        # Preserve column letters, row numbers, and merged-cell ranges so
        # the LLM can reason about the geometry of the seller's workbook
        # ("the monthly columns are C-N, the total is O, the broker
        # proforma is Q"). The old flatten-to-tabs approach lost all of
        # that and was the root cause of the categorization mismatch on
        # multi-column P&Ls.
        try:
            from openpyxl.utils import get_column_letter
            wb = load_workbook(BytesIO(data), data_only=True)
            PER_SHEET_CAP = 200_000
            sheet_blocks = [f"[Spreadsheet: {filename}]"]
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                max_row, max_col = ws.max_row or 0, ws.max_column or 0
                max_col_letter = get_column_letter(max_col) if max_col else "A"
                header = (
                    f"\n## Sheet: {sheet_name}  "
                    f"(rows 1-{max_row}, cols A-{max_col_letter})\n"
                    f"Format: <ColLetter><RowNum>: <value>  |  blanks shown as '·'  "
                    f"|  merged ranges expanded (same value repeated in every cell)"
                )
                # Map every cell inside a merged range to the top-left value.
                merged_map = {}
                merged_anchors = set()
                for mr in ws.merged_cells.ranges:
                    top_left = ws.cell(mr.min_row, mr.min_col).value
                    merged_anchors.add((mr.min_row, mr.min_col, mr.coord))
                    for r in range(mr.min_row, mr.max_row + 1):
                        for c in range(mr.min_col, mr.max_col + 1):
                            merged_map[(r, c)] = top_left
                anchor_lookup = {(r, c): coord for (r, c, coord) in merged_anchors}
                lines, used = [header], len(header)
                truncated = False
                for r in range(1, max_row + 1):
                    row_cells, row_has_value = [], False
                    for c in range(1, max_col + 1):
                        if (r, c) in merged_map:
                            v = merged_map[(r, c)]
                            suffix = (f"[merged {anchor_lookup[(r, c)]}]"
                                      if (r, c) in anchor_lookup else "")
                        else:
                            v, suffix = ws.cell(r, c).value, ""
                        if v is None:
                            token = "·"
                        else:
                            token = str(v).replace("\n", " ").replace("\t", " ").strip() or "·"
                            row_has_value = True
                        row_cells.append(f"{get_column_letter(c)}{r}: {token}{suffix}")
                    if row_has_value:
                        line = "  |  ".join(row_cells)
                        if used + len(line) + 1 > PER_SHEET_CAP:
                            truncated = True
                            break
                        lines.append(line)
                        used += len(line) + 1
                if truncated:
                    lines.append(f"... [TRUNCATED at {PER_SHEET_CAP} chars; "
                                 f"sheet has {max_row} rows total]")
                sheet_blocks.append("\n".join(lines))
            return {"type": "text", "text": "\n".join(sheet_blocks)}
        except Exception as e:
            return {"type": "text", "text": f"[Could not parse {filename}: {e}]"}
    elif ext in (".txt", ".csv", ".md"):
        return {"type": "text",
                "text": f"[File: {filename}]\n{data.decode('utf-8', errors='replace')[:200000]}"}
    else:
        return {"type": "text",
                "text": f"[File: {filename}]\n{data.decode('utf-8', errors='replace')[:100000]}"}


# ═══════════════════════════════════════════════════════════════════════════
# JSON SCHEMAS + PYDANTIC MODELS
# Two correctness layers:
#   (1) JSON Schema sent to Anthropic Structured Outputs → token-grammar
#       compliance. Eliminates malformed JSON + category-mapping drift.
#   (2) Pydantic models run after extract_json → mirror the schema for the
#       Python side and give the validator something to type-check against.
# verify_extraction (the pure-Python tie-out function) runs in addition and
# catches numeric/cross-field problems the schema layer can't see.
# ═══════════════════════════════════════════════════════════════════════════

CONFIDENCE_LEVELS = ["high", "medium", "low"]
SEVERITY_LEVELS   = ["high", "medium", "low"]
PROPERTY_TYPES    = ["MHC", "RV", "Hybrid"]
RECONCILE_UNIT    = ["match", "rent_roll_short", "rent_roll_long"]
RECONCILE_POH     = ["match", "stated_high", "stated_low", "stated_zero_but_found"]
SECTION_INCOME    = ["income"]
SECTION_EXPENSE   = ["expense"]


# ── EXTRACTION SCHEMA (Sonnet 4.6 output) ─────────────────────────────────

def _extracted_line_schema(section_values):
    """Income/expense row from the extraction stage."""
    return {
        "type": "object",
        "additionalProperties": False,
        "required": ["sellerLabel", "annualTotal", "monthly", "section", "isSubtotal"],
        "properties": {
            "sellerLabel": {"type": "string"},
            "annualTotal": {"type": ["number", "null"]},
            # monthly: 12-element array OR null. Anthropic schema accepts the
            # type-array form for nullables.
            "monthly": {
                "type": ["array", "null"],
                "items":    {"type": "number"},
                "minItems": 12,
                "maxItems": 12,
            },
            "section":    {"type": "string", "enum": section_values},
            "isSubtotal": {"type": "boolean"},
        },
    }


EXTRACTION_OUTPUT_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["reportingPeriod", "income", "expenses", "rentRoll",
                 "documentsSeen", "extractionNotes"],
    "properties": {
        "reportingPeriod": {
            "type": "object",
            "additionalProperties": False,
            "required": ["periodUsed", "dateRange", "monthsCovered",
                         "candidatePeriodsSeen", "notes"],
            "properties": {
                "periodUsed":           {"type": "string"},
                "dateRange":            {"type": "string"},
                "monthsCovered":        {"type": ["integer", "null"]},
                "candidatePeriodsSeen": {"type": "array", "items": {"type": "string"}},
                "notes":                {"type": "string"},
            },
        },
        "income":   {"type": "array", "items": _extracted_line_schema(SECTION_INCOME)},
        "expenses": {"type": "array", "items": _extracted_line_schema(SECTION_EXPENSE)},
        "rentRoll": {
            "type": "object",
            "additionalProperties": False,
            "required": ["totalRowsInRentRoll", "statedTotalRentMonthly",
                         "statedTotalIsMonthly", "occupiedCount", "vacantCount",
                         "unitTypes"],
            "properties": {
                "totalRowsInRentRoll":    {"type": ["integer", "null"]},
                "statedTotalRentMonthly": {"type": ["number", "null"]},
                "statedTotalIsMonthly":   {"type": "boolean"},
                "occupiedCount":          {"type": ["integer", "null"]},
                "vacantCount":            {"type": ["integer", "null"]},
                "unitTypes": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["unitType", "count", "occupiedCount",
                                     "vacantCount", "avgLotRentOccupied",
                                     "hasHomeRentEntries", "avgHomeRent"],
                        "properties": {
                            "unitType":           {"type": "string"},
                            "count":              {"type": "integer"},
                            "occupiedCount":      {"type": "integer"},
                            "vacantCount":        {"type": "integer"},
                            "avgLotRentOccupied": {"type": ["number", "null"]},
                            "hasHomeRentEntries": {"type": "boolean"},
                            "avgHomeRent":        {"type": ["number", "null"]},
                        },
                    },
                },
            },
        },
        "documentsSeen":   {"type": "array", "items": {"type": "string"}},
        "extractionNotes": {"type": "string"},
    },
}


# ── METHODOLOGY SCHEMA (Opus 4.8 output) ───────────────────────────────────

def _methodology_line_schema(category_enum):
    return {
        "type": "object",
        "additionalProperties": False,
        "required": ["ggcCategory", "sellerName", "fyPrior", "fyCurrent",
                     "brokerProforma", "t12Total", "monthly",
                     "ggcUnderwritten", "confidence", "notes"],
        "properties": {
            "ggcCategory":     {"type": "string", "enum": category_enum},
            "sellerName":      {"type": "string"},
            "fyPrior":         {"type": "number"},
            "fyCurrent":       {"type": "number"},
            "brokerProforma":  {"type": "number"},
            "t12Total":        {"type": "number"},
            "monthly":         {"type": "array", "items": {"type": "number"},
                                "minItems": 12, "maxItems": 12},
            "ggcUnderwritten": {"type": "number"},
            "confidence":      {"type": "string", "enum": CONFIDENCE_LEVELS},
            "notes":           {"type": "string"},
        },
    }


METHODOLOGY_OUTPUT_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": ["propertyInfo", "income", "expenses", "rentRoll",
                 "flags", "dataQualityChecks", "questions", "dataQuality"],
    "properties": {
        "propertyInfo": {
            "type": "object",
            "additionalProperties": False,
            "required": ["name", "address", "city", "state", "county",
                         "totalUnits", "askingPrice", "propertyType",
                         "ingoingCapRate", "stabilizedYieldOnCost",
                         "spreadBps", "meetsInvestmentCriteria"],
            "properties": {
                "name":        {"type": "string"},
                "address":     {"type": "string"},
                "city":        {"type": "string"},
                "state":       {"type": "string"},
                "county":      {"type": "string"},
                "totalUnits":  {"type": "integer"},
                "askingPrice": {"type": "number"},
                "propertyType":            {"type": "string", "enum": PROPERTY_TYPES},
                "ingoingCapRate":          {"type": "number"},
                "stabilizedYieldOnCost":   {"type": "number"},
                "spreadBps":               {"type": "integer"},
                "meetsInvestmentCriteria": {"type": "boolean"},
            },
        },
        "income":   {"type": "array", "items": _methodology_line_schema(GGC_INCOME_CATEGORIES)},
        "expenses": {"type": "array", "items": _methodology_line_schema(GGC_EXPENSE_CATEGORIES)},
        "rentRoll": {
            "type": "object",
            "additionalProperties": False,
            "required": ["totalUnits", "occupiedUnits", "vacantUnits",
                         "occupancyRate", "avgLotRent", "parkOwnedHomes",
                         "pohPercent", "unitGroups", "unitMixSummary"],
            "properties": {
                "totalUnits":     {"type": "integer"},
                "occupiedUnits":  {"type": "integer"},
                "vacantUnits":    {"type": "integer"},
                "occupancyRate":  {"type": "number"},
                "avgLotRent":     {"type": "number"},
                "parkOwnedHomes": {"type": "integer"},
                "pohPercent":     {"type": "number"},
                "unitGroups": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["unitType", "occupiedCount", "vacantCount",
                                     "lotRent", "pohRent", "ltoPremium",
                                     "tenantNamePattern"],
                        "properties": {
                            "unitType":          {"type": "string", "enum": CANONICAL_UNIT_TYPES},
                            "occupiedCount":     {"type": "integer"},
                            "vacantCount":       {"type": "integer"},
                            "lotRent":           {"type": "number"},
                            "pohRent":           {"type": "number"},
                            "ltoPremium":        {"type": "number"},
                            "tenantNamePattern": {"type": "string"},
                            "sellerUnitLabel":   {"type": "string"},
                        },
                    },
                },
                "rentRollRows": {
                    "type": "array",
                    "description": "Per-row extraction (preserves real Unit IDs and tenant names).",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["unitId", "unitType", "status",
                                     "tenantName", "lotRent", "homeRent"],
                        "properties": {
                            "unitId":     {"type": "string"},
                            "unitType":   {"type": "string", "enum": CANONICAL_UNIT_TYPES},
                            "status":     {"type": "string", "enum": ["Occupied", "Vacant"]},
                            "tenantName": {"type": "string"},
                            "lotRent":    {"type": "number"},
                            "homeRent":   {"type": "number"},
                            "sellerUnitLabel": {"type": "string"},
                        },
                    },
                },
                "unitMixSummary": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["unitType", "count", "occupied",
                                     "avgRent", "marketRent"],
                        "properties": {
                            "unitType":   {"type": "string", "enum": CANONICAL_UNIT_TYPES},
                            "count":      {"type": "integer"},
                            "occupied":   {"type": "integer"},
                            "avgRent":    {"type": "number"},
                            "marketRent": {"type": "number"},
                        },
                    },
                },
            },
        },
        "flags": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["item", "issue", "severity", "recommendation"],
                "properties": {
                    "item":           {"type": "string"},
                    "issue":          {"type": "string"},
                    "severity":       {"type": "string", "enum": SEVERITY_LEVELS},
                    "recommendation": {"type": "string"},
                },
            },
        },
        "dataQualityChecks": {
            "type": "object",
            "additionalProperties": False,
            "required": ["totalUnitsStated", "rentRollRowsFound",
                         "unitCountReconciliation", "pohCountStated",
                         "pohRowsFound", "pohReconciliation",
                         "finalPohCount", "finalPohPercent"],
            "properties": {
                "totalUnitsStated":         {"type": "integer"},
                "rentRollRowsFound":        {"type": "integer"},
                "unitCountReconciliation":  {"type": "string", "enum": RECONCILE_UNIT},
                "pohCountStated":           {"type": "integer"},
                "pohRowsFound":             {"type": "integer"},
                "pohReconciliation":        {"type": "string", "enum": RECONCILE_POH},
                "finalPohCount":            {"type": "integer"},
                "finalPohPercent":          {"type": "number"},
            },
        },
        "questions": {"type": "array", "items": {"type": "string"}},
        "dataQuality": {
            "type": "object",
            "additionalProperties": False,
            "required": ["hasT12", "hasT3", "hasRentRoll",
                         "hasMonthlyBreakdown", "t12Period", "missingData"],
            "properties": {
                "hasT12":              {"type": "boolean"},
                "hasT3":               {"type": "boolean"},
                "hasRentRoll":         {"type": "boolean"},
                "hasMonthlyBreakdown": {"type": "boolean"},
                "t12Period":           {"type": "string"},
                "missingData":         {"type": "array", "items": {"type": "string"}},
            },
        },
    },
}


# ── Pydantic mirrors (extraction stage) ───────────────────────────────────

class ExtractedLineItem(BaseModel):
    sellerLabel: str
    annualTotal: float | None = None
    monthly:     list[float] | None = None
    section:     str
    isSubtotal:  bool = False


class ExtractedRentRoll(BaseModel):
    totalRowsInRentRoll:    int | None = None
    statedTotalRentMonthly: float | None = None
    statedTotalIsMonthly:   bool = True
    occupiedCount:          int | None = None
    vacantCount:            int | None = None
    unitTypes:              list[dict] = Field(default_factory=list)


class ExtractedReportingPeriod(BaseModel):
    periodUsed:           str = ""
    dateRange:            str = ""
    monthsCovered:        int | None = None
    candidatePeriodsSeen: list[str] = Field(default_factory=list)
    notes:                str = ""


class ExtractedFinancials(BaseModel):
    reportingPeriod: ExtractedReportingPeriod
    income:          list[ExtractedLineItem]
    expenses:        list[ExtractedLineItem]
    rentRoll:        ExtractedRentRoll
    documentsSeen:   list[str] = Field(default_factory=list)
    extractionNotes: str = ""


# ── Pydantic mirrors (methodology stage) ──────────────────────────────────

class MethodologyLineItem(BaseModel):
    ggcCategory: str
    sellerName: str = ""
    fyPrior: float = 0
    fyCurrent: float = 0
    brokerProforma: float = 0
    t12Total: float = 0
    monthly: list[float] = Field(default_factory=lambda: [0.0] * 12)
    ggcUnderwritten: float = 0
    confidence: str = "medium"
    notes: str = ""

    @model_validator(mode="after")
    def monthly_ties_to_total(self):
        # Skip when there's nothing to tie out OR when the LLM declined
        # to emit a monthly series (some derived/synthetic lines like
        # GGC's Cap-Ex Reserve don't have monthly history).
        if not self.monthly or len(self.monthly) != 12 or abs(self.t12Total) < 1:
            return self
        # 5% tolerance matches the methodology's other thresholds (spike
        # detection, run-to-run disagreement). 1% was too tight — every
        # methodology-adjusted line item failed on rounding noise and the
        # Extraction Check tab drowned in false positives. Above 5% the
        # mismatch is real (LLM confused monthly with adjusted total, or
        # arithmetic error worth surfacing).
        total = sum(self.monthly)
        tolerance = max(50.0, abs(self.t12Total) * 0.05)
        if abs(total - self.t12Total) > tolerance:
            raise ValueError(
                f"{self.ggcCategory} ({self.sellerName}): monthly sum "
                f"${total:,.2f} != t12Total ${self.t12Total:,.2f} "
                f"(diff ${abs(total - self.t12Total):,.2f}, "
                f"{abs(total - self.t12Total) / abs(self.t12Total):.1%})"
            )
        return self


class MethodologyIncomeItem(MethodologyLineItem):
    @field_validator("ggcCategory")
    @classmethod
    def category_in_enum(cls, v):
        if v not in GGC_INCOME_CATEGORIES:
            raise ValueError(
                f"income.ggcCategory='{v}' not in GGC_INCOME_CATEGORIES"
            )
        return v


class MethodologyExpenseItem(MethodologyLineItem):
    @field_validator("ggcCategory")
    @classmethod
    def category_in_enum(cls, v):
        if v not in GGC_EXPENSE_CATEGORIES:
            raise ValueError(
                f"expense.ggcCategory='{v}' not in GGC_EXPENSE_CATEGORIES"
            )
        return v


# ═══════════════════════════════════════════════════════════════════════════
# CLAUDE CALL #1 — EXTRACTION (deterministic, Sonnet 4.6 @ temp=0)
# ═══════════════════════════════════════════════════════════════════════════
# This call does ONE job: read the documents and pull out clean numbers.
# No GGC categorization, no underwriting methodology, no judgment. Just faithful
# transcription of what's actually in the source, with the correct reporting
# period identified. Keeping this separate from the methodology call is what
# makes the numbers reliable — the model isn't juggling extraction AND analysis
# at the same time, which is what was causing wrong periods and wrong values.
EXTRACTION_PROMPT = """You are a financial data extraction engine for a real estate underwriting team. Your ONLY job is to faithfully transcribe the numbers from the attached seller documents into clean structured JSON. You do NOT categorize, analyze, or apply any methodology. You transcribe exactly what is in the documents.

CRITICAL OUTPUT RULES:
- Your response MUST start with `{` and end with `}`. NOTHING before or after.
- DO NOT add preamble, explanation, or commentary.
- DO NOT wrap the JSON in markdown code fences.
- Transcribe numbers EXACTLY as they appear. Do not round, adjust, or "correct" them.
- If a value is genuinely not present in the documents, use null. Never invent a number.

## STEP 1 — IDENTIFY THE CORRECT REPORTING PERIOD (most important)

Seller financials frequently contain MULTIPLE time periods side by side. You MUST identify which one is the operative trailing-12-month (T12) period before extracting anything.

- Look for column headers like "T-12 Ended 5/23", "T-12 Ended 9/22", "Oct 2022 - May 2023", "FY2023", individual month columns, etc.
- The operative T12 is the MOST RECENT complete trailing-12-month column. If there are several "T-12 Ended X" columns, use the one with the latest ending date.
- A column covering fewer than 12 months (e.g. "Oct 2022 - May 2023" = 8 months) is a PARTIAL period. NEVER treat a partial period as the annual figure.
- If individual monthly columns are present, the 12 most recent consecutive months ARE the T12. Sum them to cross-check against any stated T12 total.
- Report exactly which period you used and its date range in the "reportingPeriod" field.

## STEP 2 — EXTRACT THE INCOME STATEMENT (P&L / T12)

For EVERY line item in the seller's operating statement, transcribe:
- "sellerLabel": the exact label as written (e.g. "6950 · UTILITIES", "4010 · RENTAL INCOME")
- "annualTotal": the value from the operative T12 column you identified in Step 1
- "monthly": an array of the 12 monthly values for that line, IN CHRONOLOGICAL ORDER, if monthly detail exists. If no monthly detail exists, use null.
- "section": your best read of whether this is "income" or "expense" based on where it sits in the statement (this is structural, NOT GGC categorization — just income vs expense)

Rules:
- If monthly values exist, they should sum to (or very close to) the annualTotal. If they don't, still transcribe both faithfully — the verification step will flag the discrepancy.
- Include EVERY line, even ones that look like subtotals or totals. Mark subtotals/totals with "isSubtotal": true so they can be excluded from sums later.
- Preserve the seller's account numbers in the label if present.

## STEP 3 — EXTRACT THE RENT ROLL

Transcribe the rent roll into structured form:
- "totalRowsInRentRoll": the number of unit/space rows you found (integer)
- "statedTotalRentMonthly": if the rent roll shows a "Total Possible Rent" or "Totals" figure, transcribe it here. Note whether it is monthly or annual.
- "occupiedCount": number of occupied units
- "vacantCount": number of vacant units
- "unitTypes": array of distinct unit types found, each with:
    - "unitType": the label (e.g. "TURQUOISE SPACE", "Standard Lot")
    - "count": how many of this type
    - "occupiedCount": occupied of this type
    - "vacantCount": vacant of this type
    - "avgLotRentOccupied": average lot rent of the OCCUPIED units of this type
    - "hasHomeRentEntries": true if any unit of this type shows a home rent / POH rent value
    - "avgHomeRent": average home rent among units of this type that have one (null if none)

## OUTPUT SCHEMA (JSON only)

{
  "reportingPeriod": {
    "periodUsed": "string — exact label of the column you used, e.g. 'T-12 Ended 5/23'",
    "dateRange": "string — e.g. 'Jun 2022 - May 2023'",
    "monthsCovered": 12,
    "candidatePeriodsSeen": ["list every period column header you saw in the document"],
    "notes": "string — explain any ambiguity in choosing the period"
  },
  "income": [
    {"sellerLabel": "string", "annualTotal": number|null, "monthly": [12 numbers]|null, "section": "income", "isSubtotal": false}
  ],
  "expenses": [
    {"sellerLabel": "string", "annualTotal": number|null, "monthly": [12 numbers]|null, "section": "expense", "isSubtotal": false}
  ],
  "rentRoll": {
    "totalRowsInRentRoll": integer|null,
    "statedTotalRentMonthly": number|null,
    "statedTotalIsMonthly": true,
    "occupiedCount": integer|null,
    "vacantCount": integer|null,
    "unitTypes": [
      {"unitType": "string", "count": integer, "occupiedCount": integer, "vacantCount": integer,
       "avgLotRentOccupied": number|null, "hasHomeRentEntries": false, "avgHomeRent": number|null}
    ]
  },
  "documentsSeen": ["list each document by what it appears to be, e.g. 'T12 operating statement', 'rent roll', 'offering memorandum'"],
  "extractionNotes": "string — anything that was hard to read, ambiguous, or that the downstream analyst should know"
}"""


def call_extract_financials(api_key, file_blocks, property_info):
    """
    Call 1 of the financial pipeline: deterministic extraction.
    Sonnet 4.6 at temperature=0 reads the documents and returns clean numbers
    with the correct reporting period identified. No GGC methodology applied.

    Two correctness layers:
      (a) Structured outputs grammar (Anthropic beta) — guarantees schema
          compliance; the model literally cannot emit a property outside
          EXTRACTION_OUTPUT_SCHEMA.
      (b) Pydantic structural validation + verify_extraction tie-outs. On
          either kind of failure, the errors are appended to the prompt and
          extraction is re-run (up to MAX_PARSE_RETRIES). Failures past the
          last retry are surfaced by verify_extraction downstream — not
          silently dropped.
    """
    base_context = {
        "type": "text",
        "text": f"""Documents are attached above. Property context for reference only:
- Name: {property_info.get('name', 'N/A')}
- Total Units (user-stated): {property_info.get('units', 'N/A')}
- Park-Owned Home Count (user-stated): {property_info.get('pohCount', '0')}

Extract the income statement, rent roll, and reporting period into the structured JSON. Transcribe faithfully — do not categorize or analyze."""
    }
    base_user_blocks = file_blocks + [base_context]

    user_blocks    = base_user_blocks
    last_extracted = None

    for attempt in range(MAX_PARSE_RETRIES + 1):
        print(f"[Claude] Stage 1/2 — EXTRACTION attempt "
              f"{attempt+1}/{MAX_PARSE_RETRIES+1} ({MODEL_EXTRACTION}, temp=0)...")
        t0 = time.time()
        response = call_claude(api_key, EXTRACTION_PROMPT, user_blocks,
                               use_thinking=False, temperature=0,
                               model=MODEL_EXTRACTION,
                               output_schema=EXTRACTION_OUTPUT_SCHEMA)
        elapsed = time.time() - t0
        print(f"[Claude] Extraction returned in {elapsed:.1f}s "
              f"(stop_reason: {response.get('stop_reason', '?')})")
        extracted = extract_json(response)
        last_extracted = extracted

        # Layer (b1): Pydantic structural check
        pydantic_errors = []
        try:
            ExtractedFinancials(**extracted)
        except ValidationError as e:
            for err in e.errors():
                loc = ".".join(str(x) for x in err.get("loc", []))
                pydantic_errors.append(f"{loc}: {err.get('msg', 'invalid')}")

        # Layer (b2): pure-Python tie-out checks
        checks   = verify_extraction(extracted, property_info)
        failures = [c for c in checks if c.get("status") == "fail"]

        if not pydantic_errors and not failures:
            print(f"[Extract] Validation passed on attempt {attempt+1}")
            return extracted

        print(f"[Extract] Attempt {attempt+1} issues: "
              f"{len(pydantic_errors)} schema, {len(failures)} tie-out")
        for e in pydantic_errors[:5]:
            print(f"  - schema: {e}")
        for c in failures[:5]:
            print(f"  - tie-out: {c.get('item', '')}: {c.get('detail', '')}")

        if attempt < MAX_PARSE_RETRIES:
            err_lines = ([f"- schema: {e}" for e in pydantic_errors[:8]]
                         + [f"- {c.get('item','?')}: {c.get('detail','?')}"
                            for c in failures[:10]])
            user_blocks = base_user_blocks + [{
                "type": "text",
                "text": (
                    "Your previous extraction failed these checks. "
                    "Re-read the documents and produce a CORRECTED "
                    "extraction:\n\n"
                    + "\n".join(err_lines)
                    + "\n\nPay attention to: (a) the monthly array of each "
                    "income/expense line summing to its annualTotal, "
                    "(b) the rent roll occupiedCount + vacantCount equaling "
                    "totalRowsInRentRoll, (c) picking the MOST RECENT "
                    "12-month period (not a partial period), (d) section "
                    "labels exactly matching 'income' or 'expense'."
                ),
            }]

    # Retries exhausted. Return the last extraction. Downstream
    # verify_extraction surfaces the failures on the Extraction Check tab —
    # they are visible to the reviewer, not silently passed.
    print(f"[Extract] {MAX_PARSE_RETRIES+1} attempts exhausted; "
          "downstream verify_extraction will surface residual failures")
    return last_extracted


def verify_extraction(extracted, property_info):
    """
    Pure-Python verification of the extracted data — NO AI involved, so this is
    fully deterministic. Catches the failure modes that showed up in testing:
    monthly values not summing to the annual total, rent roll row count not
    matching the stated unit count, partial periods used as annual, etc.

    Returns a list of check dicts: {item, check, status, detail}
    status is "ok" | "warn" | "fail". These get surfaced on the Extraction
    Check tab so the reviewer can confirm the numbers tie before trusting
    anything downstream.
    """
    checks = []

    # ── Reporting period sanity ──────────────────────────────────────────
    rp = extracted.get("reportingPeriod", {}) or {}
    months = rp.get("monthsCovered")
    if months == 12:
        checks.append({"item": "Reporting period", "check": "T12 (12 months)",
                       "status": "ok",
                       "detail": f"{rp.get('periodUsed', '?')} ({rp.get('dateRange', '?')})"})
    elif months is not None:
        checks.append({"item": "Reporting period", "check": "Should be 12 months",
                       "status": "fail",
                       "detail": f"Period used covers {months} months — this is a PARTIAL period, not a T12. {rp.get('periodUsed', '?')}"})
    else:
        checks.append({"item": "Reporting period", "check": "12 months",
                       "status": "warn",
                       "detail": "Could not confirm months covered. Verify the period manually."})
    candidates = rp.get("candidatePeriodsSeen") or []
    if len(candidates) > 1:
        checks.append({"item": "Multiple periods in doc", "check": "Picked most recent T12",
                       "status": "warn",
                       "detail": f"Document had {len(candidates)} period columns: {', '.join(str(c) for c in candidates[:6])}. Confirm the right one was used."})

    # ── Monthly vs annual reconciliation for each line ───────────────────
    def _check_lines(lines, label):
        for ln in lines or []:
            if ln.get("isSubtotal"):
                continue
            monthly = ln.get("monthly")
            annual = ln.get("annualTotal")
            name = ln.get("sellerLabel", "(unlabeled)")
            if isinstance(monthly, list) and len(monthly) == 12 and isinstance(annual, (int, float)):
                msum = sum(v for v in monthly if isinstance(v, (int, float)))
                if annual == 0:
                    pct_off = 0 if msum == 0 else 100
                else:
                    pct_off = abs(msum - annual) / abs(annual) * 100
                if pct_off <= 1:
                    checks.append({"item": f"{label}: {name}", "check": "Monthlies sum to annual",
                                   "status": "ok", "detail": f"Σmonthly={msum:,.0f} vs annual={annual:,.0f}"})
                elif pct_off <= 5:
                    checks.append({"item": f"{label}: {name}", "check": "Monthlies ≈ annual",
                                   "status": "warn",
                                   "detail": f"Σmonthly={msum:,.0f} vs annual={annual:,.0f} ({pct_off:.1f}% off)"})
                else:
                    checks.append({"item": f"{label}: {name}", "check": "Monthlies vs annual MISMATCH",
                                   "status": "fail",
                                   "detail": f"Σmonthly={msum:,.0f} vs annual={annual:,.0f} ({pct_off:.1f}% off) — check extraction"})

    _check_lines(extracted.get("income"), "Income")
    _check_lines(extracted.get("expenses"), "Expense")

    # ── Rent roll row count vs user-stated unit count ────────────────────
    rr = extracted.get("rentRoll", {}) or {}
    rows = rr.get("totalRowsInRentRoll")
    try:
        stated_units = int(str(property_info.get("units", "")).strip() or 0)
    except (ValueError, TypeError):
        stated_units = 0
    if rows is not None and stated_units:
        if rows == stated_units:
            checks.append({"item": "Rent roll rows vs unit count", "check": "Match",
                           "status": "ok", "detail": f"{rows} rows = {stated_units} units"})
        elif rows < stated_units:
            checks.append({"item": "Rent roll rows vs unit count", "check": "Rent roll short",
                           "status": "warn",
                           "detail": f"{rows} rows but {stated_units} units stated — {stated_units - rows} likely vacant lots omitted by seller. Should be imputed at market rent."})
        else:
            checks.append({"item": "Rent roll rows vs unit count", "check": "Rent roll long",
                           "status": "fail",
                           "detail": f"{rows} rows but only {stated_units} units stated — verify unit count or check for non-unit rows."})

    # ── Rent roll stated total vs sum of unit-type rents (rough) ─────────
    stated_total = rr.get("statedTotalRentMonthly")
    annual_from_monthly = None
    if isinstance(stated_total, (int, float)) and stated_total:
        annual_from_monthly = stated_total * 12 if rr.get("statedTotalIsMonthly", True) else stated_total
        checks.append({"item": "Rent roll total", "check": "Monthly → annual",
                       "status": "ok",
                       "detail": f"Stated rent roll total {stated_total:,.0f}/mo → {annual_from_monthly:,.0f}/yr (this should approximate GPR)"})

    # ── Rent anomaly across unit types (hybrid 2σ + ratio-to-median) ─────
    # With only 3 unit types and one extreme outlier, the outlier itself
    # inflates the stdev and z stays under 2. So we also check ratio to
    # the median (robust statistic): flag any rent ≥3× or ≤1/3 the median.
    # Catches both subtle and obvious mispricings.
    rents = [ut.get("avgLotRentOccupied") for ut in (rr.get("unitTypes") or [])
             if isinstance(ut.get("avgLotRentOccupied"), (int, float))
             and ut.get("avgLotRentOccupied") > 0]
    if len(rents) >= 3:
        mean = sum(rents) / len(rents)
        median = statistics.median(rents)
        try:
            sd = statistics.stdev(rents)
        except statistics.StatisticsError:
            sd = 0
        for ut in rr.get("unitTypes") or []:
            rent = ut.get("avgLotRentOccupied")
            if not isinstance(rent, (int, float)) or rent <= 0:
                continue
            z = (rent - mean) / sd if sd > 0 else 0
            ratio = rent / median if median > 0 else 1
            if abs(z) >= 2 or ratio >= 3 or ratio <= 1/3:
                checks.append({
                    "item":  f"Rent anomaly: {ut.get('unitType', '?')}",
                    "check": "Lot rent within property range",
                    "status": "warn",
                    "detail": (
                        f"avgLotRent ${rent:,.0f} vs median ${median:,.0f} "
                        f"(ratio {ratio:.2f}×, z={z:+.2f}σ). "
                        "Verify this isn't a data entry error or a "
                        "mispriced unit type."
                    ),
                })

    # ── POH cross-check vs user-stated count ─────────────────────────────
    try:
        stated_poh = int(str(property_info.get("pohCount", "")).strip() or 0)
    except (ValueError, TypeError):
        stated_poh = 0
    any_home_entries = any(ut.get("hasHomeRentEntries")
                           for ut in (rr.get("unitTypes") or []))
    if stated_poh > 0 and not any_home_entries:
        checks.append({
            "item": "POH count vs rent roll",
            "check": f"User stated {stated_poh} POH",
            "status": "warn",
            "detail": (
                f"User stated {stated_poh} park-owned homes but no unit "
                "type in the rent roll has home rent entries. The seller "
                "may be hiding home rent income (or POH lots are comped "
                "for on-site staff). Verify with seller."
            ),
        })
    elif stated_poh == 0 and any_home_entries:
        checks.append({
            "item": "POH count vs rent roll",
            "check": "User stated 0 POH",
            "status": "fail",
            "detail": (
                "User stated 0 park-owned homes but the rent roll has "
                "home rent entries. Use the rent-roll count as "
                "authoritative — the user input was wrong."
            ),
        })

    # ── Cross-document parity: rent-roll annual GPR vs sum of income ─────
    docs = extracted.get("documentsSeen") or []
    if len(docs) >= 2 and annual_from_monthly:
        income_sum = sum(
            ln.get("annualTotal") or 0
            for ln in (extracted.get("income") or [])
            if not ln.get("isSubtotal")
            and isinstance(ln.get("annualTotal"), (int, float))
        )
        if income_sum > 0:
            spread = abs(annual_from_monthly - income_sum) / max(income_sum, 1)
            if spread > 0.20:
                checks.append({
                    "item": "Cross-doc revenue parity",
                    "check": "Rent roll annual ≈ Σ income lines",
                    "status": "warn",
                    "detail": (
                        f"Rent roll implies ${annual_from_monthly:,.0f}/yr "
                        f"GPR but income lines sum to ${income_sum:,.0f}/yr "
                        f"({spread*100:.0f}% gap). Either one source is "
                        "stale (different period) or income lines are missing."
                    ),
                })

    # ── Run-to-run disagreement, if N=3 merge ran ────────────────────────
    note = (extracted.get("extractionNotes") or "")
    if "Run-to-run disagreement" in note:
        checks.append({
            "item": "Run-to-run disagreement",
            "check": "N=3 merge: line-item annualTotal spread",
            "status": "warn",
            "detail": note.split("Run-to-run disagreement")[-1].strip()[:300],
        })

    return checks


# ═══════════════════════════════════════════════════════════════════════════
# ═══════════════════════════════════════════════════════════════════════════
# verify_methodology — post-categorization parity checks
# ═══════════════════════════════════════════════════════════════════════════
# Runs after call_parse_financials. Compares the rent roll's canonical
# unit types against the income categories the methodology assigned —
# if the rent roll has Long-term RV sites but RV Site Rental Income is
# missing or zero, the LLM almost certainly collapsed it into Gross
# Potential Rent (the bug that prompted this whole investigation).
def verify_methodology(financials):
    checks = []
    income = financials.get("income") or []
    expenses = financials.get("expenses") or []
    rr = financials.get("rentRoll") or {}

    # Index income by category (multiple rows per category aggregate).
    by_cat = {}
    for it in income:
        c = (it.get("ggcCategory") or "").strip()
        by_cat.setdefault(c, []).append(it)

    def _cat_total(cat):
        return sum(float(i.get("t12Total") or i.get("ggcUnderwritten") or 0)
                   for i in by_cat.get(cat, []))

    # Map canonical unit types → required income categories.
    unit_to_cat = {
        "TOH MH Site":        "Gross Potential Rent",
        "POH-Infilled units": "Home Rent Income",
        "Long term RV Site":  "RV Site Rental Income",
        "Retail/Commercial":  "Retail Income",
    }
    unit_groups = rr.get("unitGroups") or []
    rows = rr.get("rentRollRows") or []
    type_counts = {}
    for g in unit_groups:
        t = (g.get("unitType") or "").strip()
        type_counts[t] = type_counts.get(t, 0) + int(g.get("occupiedCount") or 0)
    for r in rows:
        if (r.get("status") or "").lower() == "occupied":
            t = (r.get("unitType") or "").strip()
            type_counts[t] = type_counts.get(t, 0) + 1

    for unit_type, required_cat in unit_to_cat.items():
        n = type_counts.get(unit_type, 0)
        if n == 0:
            continue
        cat_total = _cat_total(required_cat)
        if cat_total == 0:
            checks.append({
                "item": f"Rent-roll vs income parity: {unit_type}",
                "check": f"Expect non-zero '{required_cat}' income",
                "status": "fail",
                "detail": (f"Rent roll has {n} occupied {unit_type} unit(s) "
                           f"but '{required_cat}' income totals $0. The LLM "
                           f"likely collapsed this revenue into another "
                           f"category (most often Gross Potential Rent)."),
            })
        else:
            checks.append({
                "item": f"Rent-roll vs income parity: {unit_type}",
                "check": f"'{required_cat}' present",
                "status": "ok",
                "detail": f"{n} unit(s) → ${cat_total:,.0f} income"
            })

    # GPR row-count sanity: if rent roll has multiple revenue streams,
    # GPR (lot rent only) should be ONE row, not a sum of all streams.
    gpr_rows = by_cat.get("Gross Potential Rent", [])
    if len(gpr_rows) > 1:
        checks.append({
            "item": "Gross Potential Rent line count",
            "check": "Expect 1 GL line under GPR (lot rent only)",
            "status": "warn",
            "detail": (f"GPR has {len(gpr_rows)} line items. Multiple GL "
                       f"accounts under one GGC category sometimes "
                       f"indicates duplication — confirm none of these are "
                       f"actually RV / Storage / Retail."),
        })

    # NOI sanity — must be positive at the underwritten level.
    inc_sum = sum(float(i.get("ggcUnderwritten") or 0) for i in income)
    exp_sum = sum(float(e.get("ggcUnderwritten") or 0) for e in expenses)
    noi = inc_sum - exp_sum
    if inc_sum > 0:
        if noi <= 0:
            checks.append({
                "item": "Underwritten NOI",
                "check": "NOI > 0",
                "status": "fail",
                "detail": (f"NOI = ${noi:,.0f} (income ${inc_sum:,.0f} − "
                           f"expenses ${exp_sum:,.0f}). Deal does not cover "
                           f"its operating cost."),
            })
        else:
            ratio = exp_sum / inc_sum if inc_sum else 0
            status = ("warn" if (ratio < 0.20 or ratio > 0.65) else "ok")
            checks.append({
                "item": "Expense ratio",
                "check": "25–55% typical for MHC, 35–50% for hybrid",
                "status": status,
                "detail": f"Underwritten expense ratio = {ratio:.1%}",
            })

    # Surface any methodology Pydantic validation issues. These are
    # advisory — a monthly-tie mismatch within ~5-15% usually means the
    # LLM applied a methodology adjustment to an intermediate field, not
    # that the deal is broken. We treat the structural ones (bad
    # ggcCategory enum) as warns, and elevate to fail only when many
    # items share the same root issue (signals the LLM was confused
    # across the board, not noisy on a single line).
    methodology_issues = financials.get("_methodologyValidation") or []
    severity_for_issues = "warn" if len(methodology_issues) < 20 else "fail"
    for msg in methodology_issues:
        checks.append({
            "item": "Methodology schema validation",
            "check": "Per-line Pydantic check",
            "status": severity_for_issues,
            "detail": msg,
        })

    return checks


# ═══════════════════════════════════════════════════════════════════════════
# N=3 EXTRACTION MAJORITY VOTE (Wang et al. self-consistency)
# Runs extraction in parallel N times and field-merges. Median across runs
# for numerics, mode for strings, union for nested lists. Surfaces a
# "Run-to-run disagreement" check when annualTotals diverge >5% between
# runs — that signal is more useful than the median itself.
# Gated by deep_search so the cost (≈N× extraction tokens) is opt-in.
# ═══════════════════════════════════════════════════════════════════════════

def call_extract_financials_merged(api_key, file_blocks, property_info, n_runs=3):
    print(f"[Claude] Starting {n_runs}× extraction merge...")
    t0 = time.time()
    with ThreadPoolExecutor(max_workers=n_runs) as executor:
        futures = [executor.submit(call_extract_financials, api_key,
                                   file_blocks, property_info)
                   for _ in range(n_runs)]
        results = []
        for i, f in enumerate(as_completed(futures)):
            try:
                results.append(f.result())
                print(f"[Claude] Extraction run {i+1}/{n_runs} complete")
            except Exception as e:
                print(f"[Claude] Extraction run {i+1}/{n_runs} FAILED: {e}")

    if not results:
        raise RuntimeError(f"All {n_runs} extraction runs failed")
    print(f"[Claude] Merged {len(results)} runs in {time.time() - t0:.1f}s")
    return _merge_extraction(results)


def _median_of(vs):
    nums = [v for v in vs if isinstance(v, (int, float))]
    return statistics.median(nums) if nums else None


# Confidence-weighted-median helpers. The LLM emits a confidence enum
# {high, medium, low} per line item but the original merge step
# discarded it — so a high-confidence outlier counted the same as two
# low-confidence inliers (and vice versa). Re-introducing the weight
# pushes the merged value toward whichever runs the model was most
# certain about.
_CONF_WEIGHT = {"high": 3, "medium": 2, "low": 1}

def _weighted_median(values_with_weights):
    """Weighted median: lowest-key value v where cumulative weight
    crosses 50% of total. Falls back to plain median when all weights
    are equal or missing. values_with_weights is an iterable of
    (value, weight) — values that are None are skipped."""
    pairs = [(float(v), max(int(w or 0), 0))
             for v, w in values_with_weights
             if isinstance(v, (int, float))]
    pairs = [(v, w) for v, w in pairs if w > 0]
    if not pairs:
        return None
    pairs.sort(key=lambda p: p[0])
    total = sum(w for _, w in pairs)
    half = total / 2
    running = 0
    for v, w in pairs:
        running += w
        if running >= half:
            return v
    return pairs[-1][0]


def _median_with_confidence(items, value_field, confidence_field="confidence"):
    """Confidence-weighted median across a list of dicts. Each item
    contributes its value_field weighted by _CONF_WEIGHT[confidence_field].
    Items missing both weight and confidence default to medium."""
    pairs = []
    for it in items:
        if not isinstance(it, dict):
            continue
        v = it.get(value_field)
        conf = (it.get(confidence_field) or "medium").lower()
        w = _CONF_WEIGHT.get(conf, _CONF_WEIGHT["medium"])
        pairs.append((v, w))
    return _weighted_median(pairs)


def _mode_with_confidence(items, value_field, confidence_field="confidence"):
    """Mode of a categorical field, weighted by confidence so a
    high-confidence run can outvote two low-confidence runs that agree."""
    weighted = Counter()
    for it in items:
        if not isinstance(it, dict):
            continue
        v = it.get(value_field)
        if v in (None, ""):
            continue
        conf = (it.get(confidence_field) or "medium").lower()
        w = _CONF_WEIGHT.get(conf, _CONF_WEIGHT["medium"])
        weighted[v] += w
    return weighted.most_common(1)[0][0] if weighted else ""


def _mode_of(vs):
    nons = [v for v in vs if v not in (None, "")]
    return Counter(nons).most_common(1)[0][0] if nons else ""


def _merge_extraction(runs):
    """Field-level median for numerics, mode for strings."""
    # ── reportingPeriod: take the modal periodUsed across runs ───────────
    rp_runs = [r.get("reportingPeriod", {}) for r in runs]
    merged_rp = {
        "periodUsed": _mode_of([r.get("periodUsed") for r in rp_runs]),
        "dateRange":  _mode_of([r.get("dateRange")  for r in rp_runs]),
        "monthsCovered": int(_median_of(
            [r.get("monthsCovered") for r in rp_runs]) or 12),
        "candidatePeriodsSeen": sorted({
            p for r in rp_runs for p in (r.get("candidatePeriodsSeen") or [])
        }),
        "notes": " | ".join(sorted({
            r.get("notes", "") for r in rp_runs if r.get("notes")
        })),
    }

    # ── income / expenses: match line items by sellerLabel, median fields ──
    def _merge_lines(runs_lists, section_value):
        by_label = {}
        for run_list in runs_lists:
            for ln in run_list or []:
                key = (ln.get("sellerLabel") or "").strip().lower()
                if not key:
                    continue
                by_label.setdefault(key, []).append(ln)

        merged = []
        disagreements = []
        for key, items in by_label.items():
            # Confidence-weighted median: a high-confidence run pulls the
            # merged value toward its number; two low-confidence runs that
            # agree on a wrong figure no longer outvote one high-confidence
            # correct figure.
            annual = _median_with_confidence(items, "annualTotal")
            # detect run-to-run disagreement on this line
            non_null_annuals = [i.get("annualTotal") for i in items
                                if isinstance(i.get("annualTotal"), (int, float))]
            if len(non_null_annuals) >= 2 and max(non_null_annuals) != 0:
                spread = ((max(non_null_annuals) - min(non_null_annuals))
                          / abs(max(non_null_annuals)))
                if spread > 0.05:
                    disagreements.append({
                        "label":  items[0].get("sellerLabel", key),
                        "range":  (min(non_null_annuals), max(non_null_annuals)),
                        "spread": spread,
                    })

            # monthly: confidence-weighted median per index across runs that
            # returned 12 values. Each contributing run carries its own
            # confidence weight at every monthly position.
            monthly_with_weights = [
                (i.get("monthly"),
                 _CONF_WEIGHT.get((i.get("confidence") or "medium").lower(),
                                   _CONF_WEIGHT["medium"]))
                for i in items
                if isinstance(i.get("monthly"), list) and len(i.get("monthly")) == 12
            ]
            if monthly_with_weights:
                monthly = [
                    _weighted_median([(m[k], w) for m, w in monthly_with_weights])
                    for k in range(12)
                ]
            else:
                monthly = None

            # Surface the merged confidence so the methodology stage knows
            # how much to trust each line. Use the minimum: if any
            # contributing run was "low", the merged item is at best "low".
            confidences = [(i.get("confidence") or "").lower() for i in items]
            confidences = [c for c in confidences if c in _CONF_WEIGHT]
            if confidences:
                merged_conf = min(confidences,
                                  key=lambda c: _CONF_WEIGHT.get(c, 0))
            else:
                merged_conf = "medium"

            merged.append({
                "sellerLabel": items[0].get("sellerLabel", key),
                "annualTotal": annual,
                "monthly":     monthly,
                "section":     section_value,
                "isSubtotal":  any(i.get("isSubtotal") for i in items),
                "confidence":  merged_conf,
            })
        return merged, disagreements

    income, inc_disagree = _merge_lines(
        [r.get("income") for r in runs], "income")
    expenses, exp_disagree = _merge_lines(
        [r.get("expenses") for r in runs], "expense")

    # ── rentRoll: median scalars, union of unitTypes by name ─────────────
    rr_runs = [r.get("rentRoll", {}) for r in runs]
    merged_rr = {
        "totalRowsInRentRoll":   int(_median_of([r.get("totalRowsInRentRoll") for r in rr_runs]) or 0) or None,
        "statedTotalRentMonthly": _median_of([r.get("statedTotalRentMonthly") for r in rr_runs]),
        "statedTotalIsMonthly":  bool(rr_runs[0].get("statedTotalIsMonthly", True)) if rr_runs else True,
        "occupiedCount":         int(_median_of([r.get("occupiedCount") for r in rr_runs]) or 0) or None,
        "vacantCount":           int(_median_of([r.get("vacantCount")   for r in rr_runs]) or 0) or None,
    }
    types_by_name = {}
    for r in rr_runs:
        for ut in r.get("unitTypes") or []:
            name = (ut.get("unitType") or "").strip().lower()
            if not name:
                continue
            types_by_name.setdefault(name, []).append(ut)
    merged_types = []
    for name, items in types_by_name.items():
        merged_types.append({
            "unitType":           items[0].get("unitType", name),
            "count":              int(_median_of([i.get("count") for i in items]) or 0),
            "occupiedCount":      int(_median_of([i.get("occupiedCount") for i in items]) or 0),
            "vacantCount":        int(_median_of([i.get("vacantCount") for i in items]) or 0),
            "avgLotRentOccupied": _median_of([i.get("avgLotRentOccupied") for i in items]),
            "hasHomeRentEntries": any(i.get("hasHomeRentEntries") for i in items),
            "avgHomeRent":        _median_of([i.get("avgHomeRent") for i in items]),
        })
    merged_rr["unitTypes"] = merged_types

    # ── documentsSeen + extractionNotes: union / concat ──────────────────
    docs = sorted({d for r in runs for d in (r.get("documentsSeen") or [])})
    notes = " || ".join(sorted({
        r.get("extractionNotes", "") for r in runs if r.get("extractionNotes")
    }))

    # Add a synthetic line so the disagreement is visible downstream
    if inc_disagree or exp_disagree:
        all_disagree = inc_disagree + exp_disagree
        sample = ", ".join(
            f"{d['label']} ${d['range'][0]:,.0f}–${d['range'][1]:,.0f} "
            f"({d['spread']*100:.0f}%)"
            for d in all_disagree[:5]
        )
        notes = (
            (notes + " || " if notes else "")
            + f"Run-to-run disagreement on {len(all_disagree)} line item(s) "
            + f"across {len(runs)} extraction runs. Sample: {sample}"
        )

    return {
        "reportingPeriod": merged_rp,
        "income":          income,
        "expenses":        expenses,
        "rentRoll":        merged_rr,
        "documentsSeen":   docs,
        "extractionNotes": notes,
        "_meta": {
            "merge_runs":   len(runs),
            "disagreements": len(inc_disagree) + len(exp_disagree),
        },
    }


# ═══════════════════════════════════════════════════════════════════════════
# CLAUDE CALL #2 — METHODOLOGY (Opus 4.8, judgment)
# Takes the CLEAN extracted data from Call 1 and applies GGC's categorization
# and underwriting methodology. Because the numbers are already extracted and
# verified, this call focuses purely on analysis instead of fighting raw PDFs.
# ═══════════════════════════════════════════════════════════════════════════
FINANCIAL_PARSE_PROMPT = f"""You are a real estate underwriting analyst at Gary Group Capital (GGC), a private equity firm focused on mobile home parks.

You are given CLEAN, PRE-EXTRACTED financial data that a separate extraction step has already pulled from the seller's documents. The numbers have already been transcribed and the correct reporting period has already been identified for you. Your job is NOT to re-read raw documents — it is to apply GGC's categorization and underwriting methodology to the clean data you are given.

Trust the extracted numbers as your source of truth. The annual totals and monthly arrays in the provided data are what you work from. Map every extracted line item to GGC's standardized categories and apply the methodology below.

CRITICAL OUTPUT RULES:
- Your response MUST start with `{{` (a JSON open brace) — NOTHING before it
- Your response MUST end with `}}` (a JSON close brace) — NOTHING after it
- DO NOT add preamble, explanation, "I'll analyze...", or commentary of any kind
- DO NOT wrap the JSON in markdown code fences
- If the user-provided property info has typos or inconsistencies, use your best interpretation silently — DO NOT explain corrections in the output
- Notes should go in the "notes" field of each line item, NEVER as freestanding text

## GGC Income Categories (use EXACTLY these strings, including punctuation and capitalization — schema validation rejects deviations):
{json.dumps(GGC_INCOME_CATEGORIES)}

## GGC Expense Categories (use EXACTLY these strings — note 'Electrcitiy' typo is intentional, GGC uses it in their model. Also exact: 'Home Rent Expense (MH)' with parenthetical, 'Cap-Ex Reserve' with hyphen):
{json.dumps(GGC_EXPENSE_CATEGORIES)}

## Canonical Unit-Type Taxonomy (use EXACTLY these strings — Unit Mix Summary COUNTIFS keys on these exact labels):
{json.dumps(CANONICAL_UNIT_TYPES)}

Map seller's unit-type labels to one of the four canonical buckets:
- Lot Rent / Site Rent / TOH (tenant-owned home) lots / Pad rent → "TOH MH Site"
- POH Rent / Park-owned Home / Rental Home / Infilled Home → "POH-Infilled units"
- RV Lot / RV Site / Annual RV / Long-Term RV → "Long term RV Site"
- Storage / Retail / Commercial Space / Storefront → "Retail/Commercial"
Preserve the raw seller label in `sellerUnitLabel` for audit traceability.

## Income Categorization — GL Account Mapping (CRITICAL: emit ONE ROW per seller GL account)

Sellers' charts of accounts vary, but the GGC bucketing follows account-number prefixes consistently. Use this mapping. EMIT ONE ROW PER SELLER GL ACCOUNT — do NOT aggregate multiple GL accounts that share a GGC category into a single row. Combining 4101 + 4103 destroys the bifurcated lot/RV NOI the Underwriting tab depends on.

INCOME:
| Seller GL prefix | Example label | GGC ggcCategory |
|---|---|---|
| 4101 | Lot Rent / Site Rent / Pad Rent | "Gross Potential Rent" |
| 4103 | Long Term RV Lot Rent / RV Site | "RV Site Rental Income" |
| 4108 | Storage Unit Rent / Boat Storage | "Storage Income" |
| 4110 | Retail Unit Rent / Commercial Space | "Retail Income" |
| 4131 / 4132 / "Move-in Specials" / "Concessions" / "Discounts" | Move-in Specials | "Less: Concessions" (NEGATIVE) |
| 6120 / "Bad Debt" | Bad Debt | "Less: Bad Debt" (NEGATIVE) |
| 4304 | Damages | "Other Income" |
| 4402 (NEGATIVE: water/sewer non-recurring recovery) | Water & Sewer refund | "Other Income" with flag (NOT Utility Reimbursement — it's a contra) |
| 4403 / 4404 | Electric / Garbage tenant pass-through | "Utility Reimbursement" |
| 4905 | Recovered Legal Fees | "Other Income" |
| 4908 | Payment Processing Fee | "Other Income" |
| 4909 | Cable Revenue Sharing | "Other Income" |
| 4910 | Rental Pool Revenue Sharing | "Other Income" |
| 4913 | Application Fees | "Other Income" |
| 4914 | Late Fees | "Other Income" |
| 4915 | NSF Fees | "Other Income" |
| 4131 / Pet Fees / Damage Fees / Misc Fees | Various | "Other Income" |
| Home Rent / POH Rent / Lease-to-Own income | Home Rent | "Home Rent Income" |

DECISION RULES:
- "Utility Reimbursement" = tenant pass-through of metered/billed utility consumption (water, sewer, electric, gas, trash). If the line item represents a CONSUMPTION pass-through to a tenant, it's Utility Reimbursement.
- "Other Income" = revenue-sharing arrangements (cable, internet, laundry, vending), application/late/NSF/pet fees, damages, legal recoveries. If the line is a fee, fine, or revenue share rather than a utility pass-through, it's Other Income.
- NEGATIVE income amounts: if a line in the income block is negative AND under ~$5k absolute, treat as "Other Income" with a `notes` flag explaining the contra. If material negative, flag and route to "Other Income" with a question.
- 5407 Tenant Cable TV: if a recurring tenant charge in the income block → "Utility Reimbursement"; if a vendor revenue share → "Other Income"; if appearing in the expense block → leave as G&A.
- Concessions: any GL labeled "specials", "move-in", "concessions", "discounts" → "Less: Concessions" (always negative number).

EXPENSE BUCKETING:
| Seller GL | GGC ggcCategory |
|---|---|
| 5301 Property Tax | "RE Taxes" |
| 5050 / 5053 Liability Insurance / 5051 Car Insurance | "Insurance" |
| 5402 Water & Sewer / 5403 Water Testing | "Water and Sewer" |
| 5404 Electric | "Electrcitiy" (sic) |
| 5405 Garbage / Trash | "Trash Removal" |
| 5406 Gas / Propane / 5401 Fuel for Vehicles | "Gas/Fuel" |
| 5102 Tree / 5104 Grounds / 5103 Pest | "Ground Maintenance" |
| 5107 Septic / 5108 Plumbing / 5109 Misc / 5110 Equipment / 5111 Electrical / 5200 Supplies | "Repair and Maintenance" |
| 5000 Management Fees | "Management Fee" (will be OVERRIDDEN by GGC's % of EGI) |
| 5700-5716 (wages, casual labour, taxes, benefits, workers comp) | "Payroll" |
| 5070 Licenses & Permits / 5072 Dues / 5601-5650 Office / 5407 Cable | "General and Administrative" |
| 5061 / 5062 / 5066 Professional | "Professional Fees" |
| 5001 Advertising | "Advertising" |
| 5113 Home Repairs / POH Maintenance / "Home" labels | "Home Rent Expense (MH)" |
| 5300 Cap-Ex | "Cap-Ex Reserve" |

REJECT any deviation from the exact category strings above. The downstream Excel SUMIFS keys on these exact strings; even a trailing space or different capitalization will silently zero out the line.

## GGC Underwriting Methodology

### COLLECTIONS METHODOLOGY (NRI = Net Rental Income)

Follow this exact 4-step sequence. Do not deviate.

DEFINITIONS:
- NRI (Net Rental Income) = GPR − Vacancy − Concessions − Bad Debt. This is the standard industry term; many sellers and brokers refer to "collections" when they mean NRI.
- EGI (Effective Gross Income) = NRI + Other Income.
- Other Income excludes Home Rent (Home Rent is its own income stream and gets bifurcated, see POH section).

STEP 1 — GPR (Gross Potential Rent):
Pull directly from the rent roll.

- For OCCUPIED units: use contracted lot rent from the rent roll.
- For VACANT units WITH a market rent listed: use that market rent.
- For VACANT units WITHOUT a market rent (i.e. $0, blank, or missing): impute the market rent as the AVERAGE of occupied lot rents within the same unit type. Calculate this per unit type — a vacant Premium Lot uses the average of occupied Premium Lots, not the average of all occupied lots.
- If there is only ONE unit type in the rent roll, use the overall average of occupied rents.
- If a unit type has NO occupied units to average from (all vacant), use the average of all occupied lot rents at the property as the fallback.

- Sum the above to get GPR. This represents the 100%-occupied, market-rate ceiling.
- Note in the income line item "notes" field if you imputed market rents for any vacant units, and how many were imputed.

STEP 2 — Physical Vacancy:
 - Set from the rent roll occupancy. This is the actual physical vacancy at the property today (vacant units × market rent, as a negative line item). Do NOT use historical vacancy from the P&L for this line — physical vacancy ties to the rent roll only.

STEP 3 — Concessions:
Tie directly to T12 historical concessions. No annualization adjustment, no trend factor. If concessions are zero in the T12, set to zero.

STEP 4 — Bad Debt (the "what-if" / goal-seek step):
- This is the plug that makes underwritten NRI tie to the trend in historical collections. The logic:
  (a) Calculate underwritten NRI target: if collections are trending UP, target = T3 annualized. If flat/fluctuating, target = T12. If trending DOWN, target = T3 annualized AND flag heavily (ask why collections are declining).
  (b) Solve for Bad Debt so that: GPR − Physical Vacancy − Concessions − Bad Debt = Target NRI.
  (c) The resulting Bad Debt is the underwritten figure. Sanity check it against historical bad debt — if the underwritten bad debt diverges materially from what the seller's P&L shows, flag it. A large divergence usually means either (i) the seller is hiding something, (ii) physical occupancy changed recently, or (iii) the T3 trend is being driven by a one-time event.

- If there is a distorting one-time item in the T3 window (e.g. a one-month bad debt recovery that inflates a single month), use T6 annualized as the NRI target instead of T3, and note it.
- Output a "notes" field on the Bad Debt line item showing: which NRI target was used (T3/T6/T12), the target $ figure, and any flags.

STEP 4.5 — Income Spike Detection (pre-paid rent / lump sum):

Before annualizing T3 or T6, scan the monthly income data for anomalous spikes — single months where rental income or other income is materially above the surrounding months.

A spike is defined as: any month where a single income line is ≥1.5× the average of the other 11 months in the same line item.

Common causes:
- Tenants pre-paying multiple months at once (residents do this to avoid eviction or to lock in current rent)
- A lump-sum bad debt recovery (someone paid back a large past-due balance)
- A one-time lease-up bonus, settlement, or insurance proceeds miscategorized as rental income

If a spike is detected:
- Do NOT include the spike month in any T3 or T6 annualization calculation. Replace it with the average of the surrounding 11 months when computing T3 or T6.
- Flag the spike in the flags array with severity "medium": "Income spike detected in [month] for [line item]: [amount] vs. [avg] average. Likely pre-paid rent or lump-sum recovery. Excluded from T3/T6 annualization but kept in T12 total. Confirm with seller."
- DO include the spike in the T12 total (because T12 is the trailing reality, not the run-rate forecast).

- This is symmetric with how expense spikes are handled — flag, don't auto-strip the historical, but exclude from annualized forecasts.

### CROSS-CHECK OUTPUT (mandatory)

- In addition to the standard JSON output, include a "sourceMapping" object that shows, for each major line item, which source column/cell the dollar value came from. Format:

"sourceMapping": [
  "gpr": "Rent Roll page 8 'Total Possible Rent' $125,695 × 12 = $1,508,940",
  "reTaxes": "T-12 col 'T-12 Ended 5/23' row '6810·TAXES-PROPERTY' = $44,798",
  "insurance": "T-12 col 'T-12 Ended 5/23' row '6450 INSURANCE' = $50,462",
  ...
]

This lets the reviewer trace every number back to its source.

### GPR — MULTI-COLUMN SPREADSHEET DISAMBIGUATION (READ CAREFULLY)

Seller financials often contain MULTIPLE potential GPR figures. You must pick
the right one. Common traps:

1. **Multiple T-12 columns in one spreadsheet:**
   The seller may show several annual periods side-by-side (e.g. "T-12 Ended 9/22",
   "T-12 Ended 5/23", "Oct 22 - May 23"). Use the MOST RECENT T-12 ending date.
   The "Oct 22 - May 23"-style header is a PARTIAL period (8 months) — IGNORE it.
   The "T-12 Ended X/YY" column is the annual figure — USE THAT.

2. **Rent roll totals are MONTHLY, not annual:**
   A rent roll's "Total Possible Rent" or "Totals for [property]" row shows the
   MONTHLY rent roll total. To get annual GPR from the rent roll, multiply by 12.

3. **Sanity check (MANDATORY — apply for every deal):**
   - Expected Annual GPR ≈ Total Units × Average Lot Rent × 12
   - For Las Brisas: 295 units × ~$450/mo × 12 = ~$1.59M annual
   - If your parsed GPR is < 50% of this expected figure, you have grabbed the
     wrong column. Re-check the source and pick the column that matches the
     expected magnitude.
   - Output a note explaining which column/source you used and why.

4. **Sub-line items with the property name:**
   A line labeled "RENTAL INCOME [PROPERTY NAME]" or "4015 · RENTAL INCOME LAS
   BRISAS" is STILL rental income. Bucket it under Gross Potential Rent, not
   Other Income. The property name in the label does not change the category.

### DATA QUALITY CROSS-CHECKS (run these first, before any underwriting math)

The user provides two counts in the property form: Total Units and Park-Owned Home (POH) Count. Both must reconcile against the rent roll. If they don't, you must flag it explicitly — these mismatches almost always indicate missing data, not a clean dataset.

CHECK 1 — Total Units vs. Rent Roll Rows:
Compare the user-entered "Total Units" against the number of rent roll rows you find in the uploaded documents.

- If MATCH: proceed normally.
- If RENT ROLL HAS FEWER ROWS than Total Units: this is the most common case. Sellers often omit vacant sites from the rent roll. Assume the missing rows are vacant lots and include them in GPR at market rent (use the average of occupied lot rents in the same unit type as the imputed market rent). Add a "high" severity flag to the flags array with the exact wording: "Rent roll shows [N rows] but property is [Total Units] units — assumed [difference] additional vacant lots at market rent. Confirm with broker: are vacant sites excluded from the rent roll, or is the property actually smaller than stated?"
- If RENT ROLL HAS MORE ROWS than Total Units: flag as "high" severity. This means the user entered the wrong unit count or the rent roll includes something non-unit (model homes, storage spaces, common-area structures). Do not silently override — ask the user to confirm.

CHECK 2 — POH Count vs. Rent Roll Home Rent Entries:
Compare the user-entered "Park-Owned Home Count" against the number of rent roll rows where Home Rent > 0.

- If MATCH (within ±2): proceed normally, no flag.
- If USER SAYS MORE POH than rent roll shows home rent for: the seller is likely hiding home rent income (or comping the home for the on-site manager). Flag as "medium" severity: "User stated [X] POH but only [Y] rent roll entries show home rent. Possible employee allowance (comped lot for manager) or hidden home rent income. Verify with seller."
- If USER SAYS FEWER POH than rent roll shows home rent for: the user count is probably wrong, or some rent roll "home rent" entries are actually other charges miscategorized. Flag as "medium" severity and use the rent roll count as the authoritative POH number.
- If USER STATED 0 POH but rent roll has home rent entries: flag as "high" severity. The user either didn't know or input wrong. Use rent roll count.

POH PERCENTAGE CALCULATION:
After reconciliation, compute POH % = (POH Count / Total Units) × 100. Use this number throughout the rest of the analysis (affects expense ratio benchmarks, NOI bifurcation, and risk flags).

Output a "dataQualityChecks" object in the JSON response showing what you found:
- totalUnitsStated (integer, user input)
- rentRollRowsFound (integer, what you parsed)
- unitCountReconciliation ("match" | "rent_roll_short" | "rent_roll_long")
- pohCountStated (integer, user input)
- pohRowsFound (integer, what you parsed from rent roll)
- pohReconciliation ("match" | "stated_high" | "stated_low" | "stated_zero_but_found")
- finalPohCount (integer, the number to use going forward)
- finalPohPercent (decimal)

### OTHER INCOME
- All other income items (laundry, application fees, pet fees, month-to-month premiums, storage, etc.): use T12 as-is, no annualization adjustment
- Do not apply the T3 annualization logic to other income — only to lot rent collections

### EXPENSES (general rule)
- Most expenses: T12 * 1.03 (trailing 12 months plus 3% inflation factor for year one)
- The 3% reflects standard annual inflation assumption going into year one of ownership
- If an expense line appears materially above or below market on a per-unit basis, flag it and use your judgment — do not blindly apply T12 * 1.03 if the number is clearly distorted

### TAXES (post-acquisition reassessment)

- Underwritten taxes must reflect post-sale reassessment, not the seller's historical number.

- PRIMARY METHOD:
  - Underwritten Taxes = (Purchase Price × 0.65) × Local Tax Rate

- The new assessed value is typically 60-70% of purchase price; use 65% as the default. Apply the local tax rate (look up the millage rate or effective rate for the county/township).

- FALLBACK RULE (when you cannot find a clean tax rate or when the calculation looks suspect):
  - Underwritten Taxes = Historical T12 Taxes × 1.15

- SANITY CHECK (always apply):
  - If the primary method produces a number LOWER than historical T12 taxes, the primary method is wrong. Override it with: Historical T12 Taxes × 1.15. Reassessment after a sale never reduces taxes — if your math says it does, the millage rate or assessed value ratio is off.

NOTES:
- Cook County IL is aggressive: assesses at 10% of market value, then multiplies by a state equalization factor of ~3x, then applies the tax rate. Cook County also chases sales — flag any Cook County property as a high-reassessment-risk diligence item.
- Other counties may only reassess every 10 years.
- If a parcel number is available from the seller's tax bill, use it to look up current assessed value as a cross-check.
- Always show both methods in the notes field so Michael can sanity-check.

### INSURANCE

- Default: T12 × 1.05 (insurance has been inflating faster than general inflation).

- FLOOD ZONE OVERRIDE:
- If the property is in a FEMA flood zone (anything OTHER than Zone X), trend insurance by an additional 15%:
  - Underwritten Insurance = T12 × 1.05 × 1.15 = T12 × 1.2075

- Zone X = no special flood hazard, no override needed. Zones A, AE, AH, AO, AR, A99, V, VE = special flood hazard, apply the override.

- The user will indicate flood zone status via a yes/no input in the form. If "yes," apply the override. If "no," use the base T12 × 1.05. If unknown, default to no override but flag "verify flood zone status via FEMA flood map" as a diligence item.

- NOTE: GGC may obtain better pricing than the seller through their portfolio umbrella policy — if the seller's per-unit insurance figure is materially above market (rough benchmark: $200-300/unit/year for MHC), flag this as a potential expense reduction opportunity.

### MANAGEMENT FEE
Override the seller's management fee entirely.
- Properties under 200 sites: 5% of EGI
- Properties 200 sites or more: 4% of EGI
- EGI = NRI + Other Income (where NRI = GPR − Vacancy − Concessions − Bad Debt)

DO NOT REASSIGN PAYROLL TO MANAGEMENT FEE:
- If the seller's P&L has a "Payroll" or "Wages" line, that's on-site staff (manager, maintenance). It stays in the Payroll category.
- The Management Fee category is GGC's synthetic 4-5% of EGI fee — it's added ON TOP of payroll, not in place of it.
- If the seller already has a "Management Fee" line, override it with GGC's calculated fee. Do NOT add the seller's mgmt fee number to anything else.
- Result: every deal has BOTH a Payroll line (from seller actuals) AND a Management Fee line (synthetic GGC override). They are different categories serving different purposes.

### REPAIR & MAINTENANCE / GROUND MAINTENANCE
- These are discretionary line items — apply judgment, do not blindly use T12 * 1.03
- Look for one-time items: a single month with a spike (e.g. $4,000 vs $1,500 in all other months) likely indicates a one-time project. FLAG these per the One-Time Item Handling rule above — do not silently exclude them.
- ONE-TIME ITEM HANDLING (flag, do not auto-strip):
    - Scan monthly data for anomalous spikes: any month where a single expense line is ≥2× the average of the other 11 months.
    - DO NOT automatically back out these items. Michael (the reviewer) wants to make the inclusion/exclusion decision himself.
    - Instead, flag each spike in the flags array with severity "medium" and these exact fields:
    - item: the line item name
    - issue: "One-time item detected: [month] showed [amount] vs. [avg] average across other months. Variance of [X]×."
    - severity: "medium"
    - recommendation: "Possible one-time project (e.g. tree work, road repair, plumbing). Confirm scope with seller. If genuinely one-time, exclude from T12 × 1.03 underwriting basis. If recurring, leave in."
    - In the line item's "notes" field, also note the spike month and amount, so the reviewer can see it next to the line.
    - Use the T12 total AS-IS for the underwriting basis (T12 × 1.03). The reviewer will manually adjust if they decide a spike was truly one-time.
        - This rule applies to ALL expense lines, not just R&M — Insurance, Professional Fees, Repairs, Ground Maintenance, anything with a discernible monthly pattern. Apply consistently.
- Use per-unit benchmarks as a sanity check — flag if R&M is materially above or below typical range
- R&M covers: road repairs, pothole patching, cement work, common area repairs, occasionally home repairs that get charged back to tenants
- Ground Maintenance covers: landscaping, lawn care, tree trimming, snow removal, etc.

HOME RENT EXPENSE RE-BUCKETING:
Scan all expense line items for labels containing "home," "POH," "park-owned home," "home repairs," "RM Home," or similar. These are NOT general R&M — they belong in the Home Rent Expense bucket (GGC category: "Home Rent Expense (MH)"), not in the general "Repair and Maintenance" line.

Examples to re-bucket:
- "RM Home Repairs" → Home Rent Expense (MH)
- "POH Maintenance" → Home Rent Expense (MH)
- "Home Renovations" → Home Rent Expense (MH)

Items that stay in general R&M:
- "RM Community" → Repair and Maintenance
- "Road Repairs" → Repair and Maintenance
- "Common Area" → Repair and Maintenance

When in doubt, flag the line item and ask in the questions array.

POH percentage cross-check: If POH > 20% and you see no Home Rent Expense in the seller's financials, assume home expenses are embedded inside general R&M and flag it. Expense ratio benchmarks:
- All-TOH (no park-owned homes): 25-35% expense ratio
- Mixed (some POH): 35-40% expense ratio  
- High POH (40%+): 45-50% expense ratio
- If actual expense ratio is materially above these benchmarks for the POH mix, flag "expenses appear inflated by hidden home rent costs."

### PAYROLL
- Typically one on-site manager per community
- Many managers live on-site and receive a free lot as compensation — this is called an Employee Allowance (comped site) and should be broken out as a separate line item, not buried in payroll
- Employee Allowance = lot rent value of the comped site * 12

### PARK-OWNED HOME (POH) BIFURCATION
- This is critical — do not blend lot rent NOI and home rent NOI into a single valuation
- Identify POH percentage from the rent roll: any unit with a home rent column value is a park-owned home
- Separate income into two streams:
  (1) Lot Rent NOI — ground rent only, excludes all home rent income and home rent expense
  (2) Home Rent NOI — home rent income minus home rent expense only
- The home rent business is valued at a much higher (worse) cap rate than the lot rent business because it carries more risk, more expense, and lower rent growth
- Flag POH percentage: under 20% is acceptable, 40%+ signals a fundamentally different and less desirable operating model
- If the seller does not break out home rent expenses separately, flag this and ask for the breakout

### STABILIZED COLUMN
- In addition to the standard underwritten column, create a stabilized NOI column
- Replace current contracted lot rents with market rents (from rent comps) to show what the property generates at full market occupancy and market rents
- This is the "ceiling" valuation — what is this worth once the value-add plan is fully executed
- Stabilized cap rate = Stabilized NOI / Purchase Price — this tells you the ingoing cap rate on a fully stabilized basis

### STABILIZED YIELD ON COST (critical decision metric)

- In addition to Stabilized NOI, compute Stabilized Yield on Cost:
  - Stabilized Yield on Cost = Stabilized NOI / (Purchase Price + CapEx Reserve)

- Why "on cost" not "on price": the CapEx reserve is real money GGC has to raise and deploy to reach the stabilized state. Total cost basis is the honest denominator.

- Also compute Ingoing Cap Rate:
  - Ingoing Cap Rate = Underwritten (year-1) NOI / Purchase Price

- The decision metric is the SPREAD:
  - Spread = Stabilized Yield on Cost − Ingoing Cap Rate

- GGC's rule: the spread must be at least 200 basis points (2.00%) for the deal to clear. If the spread is below 200 bps, flag it as a hard "DOES NOT MEET INVESTMENT CRITERIA" warning at the top of the output. If the spread is above 200 bps, note how much above (e.g. "+250 bps — meets criteria with 50 bps cushion").

- GGC does NOT pay for value it's creating. The recommended purchase price should be based on the INGOING NOI at market cap rate, NOT the stabilized NOI. The stabilized column is a forward-looking check, not a valuation input.

- Return these in propertyInfo:
  - ingoingCapRate (decimal)
  - stabilizedYieldOnCost (decimal)
  - spreadBps (integer, basis points)
  - meetsInvestmentCriteria (boolean)

### CAPEX RESERVE
- Add a CapEx reserve on top of T12 expenses — this is not in the seller's financials
- Standard: $50/unit/year for older or lower-quality assets
- Higher-quality or larger assets: up to $75/unit/year
- This covers road resurfacing, utility infrastructure, common area improvements, and home addition costs

### PER-UNIT BENCHMARKS
- Always calculate and display expense figures on a per-unit basis alongside total figures
- Per-unit figures allow quick sanity checks against GGC's experience across 100+ deals
- If a per-unit figure looks materially off vs. typical ranges, flag it for review — do not just accept the seller's number

### ONE-TIME ITEMS & VARIANCE FLAGS
- Scan all expense line items for month-over-month spikes — any month that is 2x or more the average of other months is likely a one-time item
- Flag these explicitly: identify the month, the amount, and the variance from average
- Michael will decide whether to exclude them from the underwritten run rate
- High professional fees often correlate with high bad debt — may indicate the seller has been filing frequent evictions
- High R&M in a single month often indicates a capital project (water line, road work) that should not recur

### ALTERNATIVE HOUSING (required output)
- Average single-family home sale price within 5-mile and 10-mile radius
- Year-over-year home price appreciation percentage
- Average 2-bedroom apartment rent in the market
- Average 3-bedroom apartment rent in the market
- MHP all-in monthly cost = lot rent + estimated home payment (assume $50,000 home, $3,000 down, $47,000 loan at current rates)
- Compare MHP all-in cost to 2BR and 3BR apartment rents — GGC targets MHP all-in cost at less than comparable rental alternatives
- GGC targets home sale prices at approximately 30% of the average single-family home price in the area — this is their demand threshold

### RECOMMENDED PURCHASE PRICE OUTPUT
- Derive a recommended purchase price using two methods:
  (1) Cap Rate Method: Underwritten NOI / Market Cap Rate (from sale comps)
  (2) Stabilized Cap Rate Method: Stabilized NOI / Market Cap Rate
- Also show: Price per unit at recommended purchase price
- Compare to asking price and flag the gap
- Note: Lot Rent NOI and Home Rent NOI should be capitalized at different rates — do not blend them
  * Lot Rent NOI: capitalize at market cap rate for the asset class and submarket (typically 5-7%)
  * Home Rent NOI: capitalize at a higher cap rate (typically 12-15%) reflecting the lower quality of that income stream

## Output (JSON only, no prose, no fences)

{{
  "propertyInfo": {{
    "name": "string", "address": "string", "city": "string", "state": "string",
    "county": "string", "totalUnits": integer, "askingPrice": number,
    "propertyType": "MHC|RV|Hybrid",
    "ingoingCapRate": number,
    "stabilizedYieldOnCost": number,
    "spreadBps": integer,
    "meetsInvestmentCriteria": boolean
  }},
  "income": [
    {{
      "ggcCategory": "string", "sellerName": "string",
      "fyPrior": number, "fyCurrent": number, "brokerProforma": number,
      "t12Total": number, "monthly": [12 numbers],
      "ggcUnderwritten": number, "confidence": "high|medium|low", "notes": "string"
    }}
  ],
  "expenses": [{{ same shape as income }}],

  "rentRoll": {{
    "totalUnits": integer, "occupiedUnits": integer, "vacantUnits": integer,
    "occupancyRate": number, "avgLotRent": number, "parkOwnedHomes": integer,
    "pohPercent": number,
    "rentRollRows": [
      {{"unitId": "string (lot number, e.g. 'A05', 'B07', 'EL02A')",
        "unitType": "MUST be one of [TOH MH Site, POH-Infilled units, Long term RV Site, Retail/Commercial]",
        "status": "Occupied | Vacant",
        "tenantName": "string (real tenant name from rent roll, or '' for vacant)",
        "lotRent": number,
        "homeRent": number (0 if no POH rent),
        "sellerUnitLabel": "string (raw seller label, e.g. 'WHA Lot', 'Storage') — for audit trail"}}
    ],
    "unitGroups": [
      {{"unitType": "MUST be one of [TOH MH Site, POH-Infilled units, Long term RV Site, Retail/Commercial]",
        "occupiedCount": integer,
        "vacantCount": integer,
        "lotRent": number,
        "pohRent": number (0 if TOH),
        "ltoPremium": number,
        "tenantNamePattern": "string",
        "sellerUnitLabel": "string (raw seller label, for audit)"}}
    ],
    "unitMixSummary": [
      {{"unitType": "one of [TOH MH Site, POH-Infilled units, Long term RV Site, Retail/Commercial]",
        "count": integer, "occupied": integer, "avgRent": number, "marketRent": number}}
    ]
  }},
  "flags": [{{"item", "issue", "severity", "recommendation"}}],
    "dataQualityChecks": {{
    "totalUnitsStated": integer,
    "rentRollRowsFound": integer,
    "unitCountReconciliation": "match|rent_roll_short|rent_roll_long",
    "pohCountStated": integer,
    "pohRowsFound": integer,
    "pohReconciliation": "match|stated_high|stated_low|stated_zero_but_found",
    "finalPohCount": integer,
    "finalPohPercent": number
  }},
  "questions": ["string"],
  "dataQuality": {{"hasT12", "hasT3", "hasRentRoll", "hasMonthlyBreakdown",
    "t12Period": "string", "missingData": ["string"]}}
}}

RENT ROLL OUTPUT GUIDANCE:
- ALWAYS emit `rentRollRows` with one entry per actual rent roll row (preserves real tenant names and Unit IDs). This is required for audit and downstream delinquency analysis.
- ALSO emit `unitGroups` as the aggregated 4-category summary. The backend reads `rentRollRows` first, falling back to `unitGroups` only if missing.
- For `unitGroups`, aggregate ONLY by the 4 canonical categories above — never by raw seller labels like "WHA Lot" or "Storage". A property with 100 WHA Lots + 5 Storage units + 1 Commercial should emit 2-3 `unitGroups`: one "TOH MH Site" with count=100, one "Retail/Commercial" with count=5-6. The COUNTIFS in the Excel template keys on these exact canonical strings."""


def call_parse_financials(api_key, extracted, property_info):
    """
    Call 2 of the financial pipeline: GGC methodology on clean extracted data.
    Opus 4.8 (adaptive thinking, effort=high) takes the verified extraction
    output and applies categorization + underwriting logic. No raw documents —
    it works from the clean JSON the extraction step produced.
    """
    user_blocks = [{
        "type": "text",
        "text": f"""Property context (user-stated):
- Name: {property_info.get('name', 'N/A')}
- Address: {property_info.get('address', 'N/A')}
- Total Units: {property_info.get('units', 'N/A')}
- Park-Owned Home Count (user-stated): {property_info.get('pohCount', '0')}
- Asking Price: ${property_info.get('askingPrice', 'N/A')}
- Flood Zone Status: {property_info.get('floodZone', 'unknown')}

Below is the CLEAN, PRE-EXTRACTED financial data from the seller's documents. The reporting period has already been identified and the numbers transcribed. Apply GGC's categorization and underwriting methodology to this data and return the structured JSON.

=== EXTRACTED FINANCIAL DATA ===
{json.dumps(extracted, indent=2)[:180000]}
=== END EXTRACTED DATA ===

Apply the GGC methodology and return the structured JSON."""
    }]
    print("[Claude] Stage 2/2 — METHODOLOGY (Opus 4.8, effort=high)...")
    t0 = time.time()
    response = call_claude(api_key, FINANCIAL_PARSE_PROMPT, user_blocks,
                           use_thinking=True, model=MODEL_METHODOLOGY,
                           output_schema=METHODOLOGY_OUTPUT_SCHEMA,
                           max_tokens=MAX_TOKENS_METHODOLOGY)
    elapsed = time.time() - t0
    print(f"[Claude] Methodology returned in {elapsed:.1f}s "
          f"(stop_reason: {response.get('stop_reason', '?')})")
    parsed = extract_json(response)

    # Activate the methodology-line validators. They reject any income
    # row whose ggcCategory isn't in GGC_INCOME_CATEGORIES (likewise for
    # expenses) — the SUMIFS in the Underwriting template will silently
    # zero out anything misspelled, so we want to fail loud here.
    validation_errors = []
    for i, item in enumerate(parsed.get("income") or []):
        try:
            MethodologyIncomeItem.model_validate(item)
        except ValidationError as e:
            validation_errors.append(f"income[{i}]: {e.errors()[0].get('msg', str(e))}")
    for i, item in enumerate(parsed.get("expenses") or []):
        try:
            MethodologyExpenseItem.model_validate(item)
        except ValidationError as e:
            validation_errors.append(f"expenses[{i}]: {e.errors()[0].get('msg', str(e))}")
    if validation_errors:
        # Capture so the Extraction Check tab can render the failures.
        # Don't raise — the rest of the run may still be usable, and the
        # tab is where the reviewer expects to see categorization issues.
        parsed.setdefault("_methodologyValidation", []).extend(validation_errors)
        print(f"[Methodology] {len(validation_errors)} validation warnings "
              f"(see Extraction Check tab)")
    return parsed


# ═══════════════════════════════════════════════════════════════════════════
# CLAUDE CALL #2 — Market Research
# ═══════════════════════════════════════════════════════════════════════════
MARKET_RESEARCH_PROMPT = """You are a CRE analyst at GGC researching a mobile home park acquisition. Be exhaustive - better data means a better deal decision.

Use web_search aggressively (8 searches available) to find:

1. **Rent comps** - 12-20 neighboring MHPs within ~25 miles. For EACH one capture:
   - Name, full address, city, state
   - Distance from subject (miles)
   - Total units (sites)
   - Lot rent (monthly)
   - Occupancy if disclosed
   - Year built / age class
   - Park-owned home %
   - Amenities list (pool, clubhouse, etc.)
   - Star/quality rating if available (1-5)
   - Source URL

2. **Sale comps** — 8-15 recent MHP transactions in the broader region (last 4 years). For EACH:
   - Property name, location, sale date (MM/YY)
   - # sites
   - Sale price
   - $/unit
   - Cap rate at sale (if disclosed)
   - NOI at sale (if known)
   - Buyer / seller (if disclosed)
   - Source URL

3. **Demographics (rich)** — for both county AND MSA level if different:
   - Total population (current)
   - 1-year, 5-year, and 10-year population growth %
   - Population projection (next 5 years if available)
   - Median household income (current)
   - 5-year household income growth %
   - Unemployment rate (current and 1yr ago for trend)
   - Poverty rate
   - % of households earning under $50K (target MHP demographic)
   - Top 5-10 employers by name with employee count if available
   - Major industries / economic drivers
   - Any large planned developments, factory openings, base expansions, etc.

4. **Alternative housing** —
   - Avg single-family home price + 1yr appreciation %
   - Median home price
   - Avg apartment rent: 1BR, 2BR, 3BR
   - Apartment rent growth (1yr)
   - Vacancy rate of local apartments
   - Construction permits / pipeline for apartments and SFR

5. **MHP affordability calc** — compute MHP all-in monthly cost (lot rent + $50K home @ 8% × 20yr ≈ $418/mo + lot rent) vs 2BR apt rent. Output the savings %.

6. **Landmarks** — distance in miles to:
   - 3 nearest major employers (NAMED, not generic)
   - Walmart / Target / big-box
   - Grocery store
   - Hospital / urgent care
   - Elementary school
   - High school
   - Community college or university
   - Major highway / interstate
   - Downtown (nearest city)
   - Airport

7. **Visual reference URLs** — Google Maps satellite/street view URLs for the subject

CRITICAL OUTPUT RULES:
- Your response MUST start with `{` and end with `}`. NOTHING else.
- DO NOT preface, explain, correct typos out loud, or comment.
- DO NOT use markdown code fences.
- All percentages as decimals (0.045 NOT "4.5%" or "4.5").
- All currency as plain numbers (78000 NOT "$78,000").
- Inside string values, escape any newlines as \\n. NO literal line breaks inside strings.

## Output schema (JSON only):
{
  "rentComps": [{"name", "address", "city", "state", "distance" (string e.g. "4.2 mi"), "units" (int), "lotRent" (number), "occupancy" (decimal 0-1), "yearBuilt" (int or null), "pohPercent" (decimal 0-1 or null), "amenities" (string), "qualityRating" (1-5 or null), "source" (URL)}],
  "saleComps": [{"name", "location", "saleDate" (MM/YY), "units" (int), "salePrice" (number), "pricePerUnit" (number), "capRate" (decimal 0-1 or null), "noi" (number or null), "buyer" (string or null), "seller" (string or null), "source" (URL)}],
  "demographics": {
    "countyName" (string),
    "countyPopulation" (int),
    "populationGrowth1yr" (decimal),
    "populationGrowth5yr" (decimal),
    "populationGrowth10yr" (decimal),
    "populationProjection5yr" (int or null),
    "medianHHIncome" (int),
    "incomeGrowth5yr" (decimal),
    "unemploymentRate" (decimal),
    "unemploymentRate1yrAgo" (decimal or null),
    "povertyRate" (decimal or null),
    "pctHHUnder50k" (decimal or null),
    "majorEmployers" (array of strings, ideally with employee counts),
    "majorIndustries" (array of strings),
    "plannedDevelopments" (string summary or null)
  },
  "altHousing": {
    "avgSFHomePrice" (int),
    "medianSFHomePrice" (int or null),
    "homePriceGrowth1yr" (decimal),
    "avgRent1BR" (int),
    "avgRent2BR" (int),
    "avgRent3BR" (int),
    "apartmentRentGrowth1yr" (decimal or null),
    "apartmentVacancyRate" (decimal or null),
    "mhpAllInCost" (int),
    "mhpVsApartmentSavingsPercent" (number, e.g. 35 for 35%)
  },
  "landmarks": [{"type" (use the exact labels: "Nearest Major Employer #1", "Nearest Major Employer #2", "Nearest Major Employer #3", "Nearest Walmart / Big Box", "Nearest Grocery Store", "Nearest Hospital / Urgent Care", "Nearest Elementary School", "Nearest High School", "Nearest Community College / University", "Nearest Major Highway / Interstate", "Downtown (nearest city)", "Nearest Airport"), "name", "distanceMiles" (number)}],
  "visuals": {"aerialView", "streetViewEntrance", "streetViewInterior", "exampleHome1", "exampleHome2", "directions"},
  "marketRentConclusion": "string (3-4 sentences with specific numbers — where subject sits vs comp range, implied upside)",
  "marketCapRateConclusion": "string (3-4 sentences with specific cap rate ranges from comps)",
  "demandSignal": "STRONG|MODERATE|WEAK",
  "demandRationale": "string (4-6 sentences citing specific demographic, employer, and affordability numbers — escape any newlines as \\n)"
}
"""


def call_market_research(api_key, property_info):
    prompt = f"""Property: {property_info.get('name', '')}
Address: {property_info.get('address', '')}
City/State: {property_info.get('city', '')}, {property_info.get('state', '')}
Units: {property_info.get('units', '')}

Pull comps, demographics, alt housing, landmarks, and visual URLs."""
    tools = [{"type": "web_search_20250305", "name": "web_search", "max_uses": 12}]
    print("[Claude] Starting market research call (with web_search)...")
    t0 = time.time()
    response = call_claude(api_key, MARKET_RESEARCH_PROMPT,
                            [{"type": "text", "text": prompt}], tools=tools,
                            model=MODEL_MARKET,
                            max_tokens=MAX_TOKENS_MARKET)
    elapsed = time.time() - t0
    print(f"[Claude] Market research returned in {elapsed:.1f}s "
          f"(stop_reason: {response.get('stop_reason', '?')})")
    return extract_json(response)

def call_market_research_merged(api_key, property_info, n_runs=3):
    """
    Run market research multiple times in parallel and merge the results.
    Dedupes rent comps by name, sale comps by name+date, takes the union of
    landmarks and demographics (most common value wins), and concatenates
    the narrative conclusions.

    Cost: ~n_runs × single-run cost. Use for high-stakes deals.
    """
    print(f"[Claude] Starting {n_runs}× market research merge...")
    t0 = time.time()

    # Fire all runs in parallel — they're independent and the bottleneck is
    # network/web-search latency, not local CPU
    with ThreadPoolExecutor(max_workers=n_runs) as executor:
        futures = [executor.submit(call_market_research, api_key, property_info)
                   for _ in range(n_runs)]
        results = []
        for i, f in enumerate(as_completed(futures)):
            try:
                results.append(f.result())
                print(f"[Claude] Market research run {i+1}/{n_runs} complete")
            except Exception as e:
                print(f"[Claude] Market research run {i+1}/{n_runs} FAILED: {e}")
                # Continue with however many succeeded — don't lose the whole job
                # because one of three calls timed out

    if not results:
        raise RuntimeError(f"All {n_runs} market research runs failed")

    elapsed = time.time() - t0
    print(f"[Claude] Merged {len(results)} runs in {elapsed:.1f}s")

    return _merge_market_research(results)


def _merge_market_research(runs):
    """
    Merge logic:
    - rentComps: union by name (case-insensitive), prefer the run with more fields populated
    - saleComps: union by name+saleDate, prefer the run with more fields
    - landmarks: union by type, prefer the closest-distance entry for each type
    - demographics: take the most common non-null value per field (mode), fall back to first
    - altHousing: average the numeric fields, take most common for strings
    - narrative conclusions: concatenate with " | " separator, deduped
    - demandSignal: majority vote across runs
    """
    def _completeness(d):
        """Score how 'full' a comp dict is — used to break ties between duplicate comps."""
        return sum(1 for v in d.values() if v not in (None, "", 0))

    # ── RENT COMPS: dedupe by lowercased name, keep most complete version ──
    rent_by_name = {}
    for run in runs:
        for comp in run.get("rentComps", []) or []:
            key = (comp.get("name") or "").strip().lower()
            if not key:
                continue
            if key not in rent_by_name or _completeness(comp) > _completeness(rent_by_name[key]):
                rent_by_name[key] = comp
    # Parse distance strings like "4.2 mi" or "12 miles" into a float so
    # we sort numerically (lexicographic puts "10.1 mi" before "4.2 mi").
    def _parse_distance_miles(s):
        if isinstance(s, (int, float)):
            return float(s)
        if not isinstance(s, str):
            return float("inf")
        import re
        m = re.search(r"-?\d+(\.\d+)?", s)
        return float(m.group(0)) if m else float("inf")

    merged_rent_comps = sorted(rent_by_name.values(),
                                key=lambda c: _parse_distance_miles(c.get("distance")))

    # ── SALE COMPS: dedupe by name+saleDate, same completeness rule ──
    sale_by_key = {}
    for run in runs:
        for comp in run.get("saleComps", []) or []:
            key = ((comp.get("name") or "").strip().lower(),
                   (comp.get("saleDate") or "").strip())
            if not key[0]:
                continue
            if key not in sale_by_key or _completeness(comp) > _completeness(sale_by_key[key]):
                sale_by_key[key] = comp

    # Parse sale dates ("MM/YY", "MM/YYYY", "YYYY-MM") so we sort
    # newest-first (most recent transactions are the meaningful cap-rate
    # signal; lexicographic puts "01/26" before "12/22").
    def _parse_sale_year(s):
        if not isinstance(s, str):
            return -1
        import re
        # Capture trailing 2- or 4-digit year, optional month prefix.
        m = re.search(r"(\d{1,2})[/\-](\d{2,4})$", s.strip())
        if m:
            yr = int(m.group(2))
            yr = 2000 + yr if yr < 100 else yr
            return yr * 12 + int(m.group(1))
        # ISO-ish fallback "2024-03"
        m = re.match(r"(\d{4})-(\d{1,2})", s.strip())
        if m:
            return int(m.group(1)) * 12 + int(m.group(2))
        return -1

    merged_sale_comps = sorted(sale_by_key.values(),
                                key=lambda c: _parse_sale_year(c.get("saleDate")),
                                reverse=True)

    # ── LANDMARKS: one entry per landmark type, keep closest ──
    landmark_by_type = {}
    for run in runs:
        for lm in run.get("landmarks", []) or []:
            t = lm.get("type", "")
            if not t:
                continue
            dist = lm.get("distanceMiles") or 999
            if t not in landmark_by_type or dist < (landmark_by_type[t].get("distanceMiles") or 999):
                landmark_by_type[t] = lm
    merged_landmarks = list(landmark_by_type.values())

    # ── DEMOGRAPHICS: mode (most common value) per field, falls back to first non-null ──
    from collections import Counter
    def _consensus(field_name, runs_list):
        values = [r.get("demographics", {}).get(field_name) for r in runs_list]
        non_null = [v for v in values if v not in (None, "")]
        if not non_null:
            return None
        # For numeric fields, average; for strings, take mode
        if all(isinstance(v, (int, float)) for v in non_null):
            return sum(non_null) / len(non_null)
        if all(isinstance(v, list) for v in non_null):
            # Lists like majorEmployers — take union, preserve order
            seen, out = set(), []
            for v in non_null:
                for item in v:
                    if item not in seen:
                        seen.add(item)
                        out.append(item)
            return out
        # Strings — return most common
        c = Counter(str(v) for v in non_null)
        return c.most_common(1)[0][0]

    demo_fields = ["countyName", "countyPopulation", "populationGrowth1yr",
                    "populationGrowth5yr", "populationGrowth10yr", "populationProjection5yr",
                    "medianHHIncome", "incomeGrowth5yr", "unemploymentRate",
                    "unemploymentRate1yrAgo", "povertyRate", "pctHHUnder50k",
                    "majorEmployers", "majorIndustries", "plannedDevelopments"]
    merged_demo = {f: _consensus(f, runs) for f in demo_fields}

    # ── ALT HOUSING: average numerics, mode for strings ──
    alt_fields = ["avgSFHomePrice", "medianSFHomePrice", "homePriceGrowth1yr",
                   "avgRent1BR", "avgRent2BR", "avgRent3BR",
                   "apartmentRentGrowth1yr", "apartmentVacancyRate",
                   "mhpAllInCost", "mhpVsApartmentSavingsPercent"]
    merged_alt = {}
    for f in alt_fields:
        values = [r.get("altHousing", {}).get(f) for r in runs]
        non_null = [v for v in values if isinstance(v, (int, float))]
        if non_null:
            merged_alt[f] = sum(non_null) / len(non_null)

    # ── VISUALS: take first non-empty per field ──
    visual_fields = ["aerialView", "streetViewEntrance", "streetViewInterior",
                      "exampleHome1", "exampleHome2", "directions"]
    merged_visuals = {}
    for f in visual_fields:
        for r in runs:
            v = (r.get("visuals", {}) or {}).get(f)
            if v:
                merged_visuals[f] = v
                break

    # ── NARRATIVE CONCLUSIONS: concat, deduped — different runs caught different angles ──
    def _merge_narrative(field):
        seen, parts = set(), []
        for r in runs:
            text = (r.get(field) or "").strip()
            if text and text.lower() not in seen:
                seen.add(text.lower())
                parts.append(text)
        return "  ||  ".join(parts) if parts else ""

    # ── DEMAND SIGNAL: majority vote ──
    signals = [r.get("demandSignal") for r in runs if r.get("demandSignal")]
    if signals:
        signal_vote = Counter(signals).most_common(1)[0][0]
    else:
        signal_vote = "MODERATE"

    return {
        "rentComps": merged_rent_comps,
        "saleComps": merged_sale_comps,
        "landmarks": merged_landmarks,
        "demographics": merged_demo,
        "altHousing": merged_alt,
        "visuals": merged_visuals,
        "marketRentConclusion": _merge_narrative("marketRentConclusion"),
        "marketCapRateConclusion": _merge_narrative("marketCapRateConclusion"),
        "demandSignal": signal_vote,
        "demandRationale": _merge_narrative("demandRationale"),
        "_meta": {
            "merge_runs": len(runs),
            "merged_rent_comps_count": len(merged_rent_comps),
            "merged_sale_comps_count": len(merged_sale_comps),
        }
    }

# ═══════════════════════════════════════════════════════════════════════════
# TEMPLATE FILLING — uses GGC's actual blank template
# ═══════════════════════════════════════════════════════════════════════════
def fill_template(financials, market, output_path):
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"GGC template not found at {TEMPLATE_PATH}. "
            f"Make sure GGC_Blank_Underwriting_Sizer.xlsx is in the same folder as backend.py."
        )

    wb = load_workbook(TEMPLATE_PATH)

    # ── Data Consolidation ────────────────────────────────────────────────
    # Income rows 3-21, Expense rows 28-58
    # Cols: A=GGC Cat, B=Source Name, D=FY Prior, E=FY Current, F=Broker PF,
    #       G=T12, J-U=monthly (12), H=annualization (formula — don't touch)
    ws = wb["Data Consolidation"]
    income_items = financials.get("income", [])
    expense_items = financials.get("expenses", [])

    for i, item in enumerate(income_items[:19]):
        r = 3 + i
        ws.cell(row=r, column=1, value=item.get("ggcCategory", ""))
        ws.cell(row=r, column=2, value=item.get("sellerName", ""))
        ws.cell(row=r, column=4, value=item.get("fyPrior", 0))
        ws.cell(row=r, column=5, value=item.get("fyCurrent", 0))
        ws.cell(row=r, column=6, value=item.get("brokerProforma", 0))
        ws.cell(row=r, column=7, value=item.get("t12Total", 0))
        monthly = item.get("monthly") or []
        if len(monthly) == 12:
            for m_i, val in enumerate(monthly):
                ws.cell(row=r, column=10 + m_i, value=val)
        elif item.get("t12Total"):
            even = (item["t12Total"] or 0) / 12
            for m_i in range(12):
                ws.cell(row=r, column=10 + m_i, value=even)

    for i, item in enumerate(expense_items[:31]):
        r = 28 + i
        ws.cell(row=r, column=1, value=item.get("ggcCategory", ""))
        ws.cell(row=r, column=2, value=item.get("sellerName", ""))
        ws.cell(row=r, column=4, value=item.get("fyPrior", 0))
        ws.cell(row=r, column=5, value=item.get("fyCurrent", 0))
        ws.cell(row=r, column=6, value=item.get("brokerProforma", 0))
        ws.cell(row=r, column=7, value=item.get("t12Total", 0))
        monthly = item.get("monthly") or []
        if len(monthly) == 12:
            for m_i, val in enumerate(monthly):
                ws.cell(row=r, column=10 + m_i, value=val)
        elif item.get("t12Total"):
            even = (item["t12Total"] or 0) / 12
            for m_i in range(12):
                ws.cell(row=r, column=10 + m_i, value=even)

    # ── Rent Roll Input ────────────────────────────────────────────────────
    # Restructured column layout (matches Unit Mix Summary COUNTIFS/SUMIFS):
    #   A=Count (formula), B=Unit ID, C=Unit Type (canonical), D=Status,
    #   F=Tenant Name, G=Type detail, H=Type code, I=Lot Rent, J=Home Rent,
    #   K=Combined (formula). Data rows 3-1002.
    #
    # Prefer per-row data from rentRoll.rentRollRows (preserves real tenant
    # names + unit IDs). Fall back to unitGroups expansion when only the
    # aggregated form is available.
    ws = wb["Rent Roll Input"]
    rr = financials.get("rentRoll") or {}
    per_row = rr.get("rentRollRows") or []
    unit_groups = rr.get("unitGroups") or []

    individual_units = []
    if per_row:
        for row in per_row:
            individual_units.append({
                "unitId":    row.get("unitId", "") or "",
                "unitType":  row.get("unitType", "") or "",
                "status":    row.get("status", "Occupied") or "Occupied",
                "tenantName": row.get("tenantName", "") or "",
                "lotRent":   row.get("lotRent", 0) or 0,
                "homeRent":  row.get("homeRent", 0) or 0,
            })
    else:
        for grp in unit_groups:
            ut = grp.get("unitType", "Unit")
            lot_rent = grp.get("lotRent", 0) or 0
            home_rent = grp.get("pohRent", 0) or 0
            name_prefix = grp.get("tenantNamePattern", "Tenant")
            occ_count = grp.get("occupiedCount", 0) or 0
            vac_count = grp.get("vacantCount", 0) or 0
            for i in range(occ_count):
                individual_units.append({
                    "unitId": "", "unitType": ut, "status": "Occupied",
                    "tenantName": f"{name_prefix} {len(individual_units) + 1}",
                    "lotRent": lot_rent, "homeRent": home_rent,
                })
            for _ in range(vac_count):
                individual_units.append({
                    "unitId": "", "unitType": ut, "status": "Vacant",
                    "tenantName": "", "lotRent": 0, "homeRent": 0,
                })

    for i, unit in enumerate(individual_units[:1000]):
        r = 3 + i
        ws.cell(row=r, column=2,  value=unit.get("unitId", ""))      # B
        ws.cell(row=r, column=3,  value=unit.get("unitType", ""))    # C
        ws.cell(row=r, column=4,  value=unit.get("status", ""))      # D
        ws.cell(row=r, column=6,  value=unit.get("tenantName", ""))  # F
        ws.cell(row=r, column=9,  value=unit.get("lotRent", 0))      # I
        ws.cell(row=r, column=10, value=unit.get("homeRent", 0))     # J
        # A (Count) and K (Combined) are formulas seeded by fix_template.py

    # ── Add Miscellaneous tab ──────────────────────────────────────────────
    if "Miscellaneous" in wb.sheetnames:
        del wb["Miscellaneous"]
    add_miscellaneous_tab(wb, financials, market)

    # ── Add Comps Analysis tab ─────────────────────────────────────────────
    if "Comps Analysis" in wb.sheetnames:
        del wb["Comps Analysis"]
    add_comps_analysis_tab(wb, financials, market)

    # ── Add Extraction Check tab (source reconciliation) ───────────────────
    # This is the "do the numbers tie out?" tab Michael asked for in the
    # meeting. Lives at the front so it's the first thing the reviewer sees.
    if "Extraction Check" in wb.sheetnames:
        del wb["Extraction Check"]
    add_extraction_check_tab(wb, financials)

    # ── Subject property cells the template formulas key on ───────────────
    # The patched template puts the Subject pricing block at columns O-P
    # of GGC Underwriting. The P4 (Purchase Price) cell is wired to
    # =IFERROR(IF(ISNUMBER(P9),P9,0),0), so it reads from P9 (Asking
    # Price). Without this write, P4 stays at 0 and the entire Sources
    # and Uses / Loan Scenario / Pro Forma Y0 chain collapses.
    underw = wb["GGC Underwriting"]
    prop = financials.get("propertyInfo") or {}
    try:
        ask = float(prop.get("askingPrice") or 0)
    except (TypeError, ValueError):
        ask = 0
    if ask > 0:
        underw["P9"] = ask
    # Also expose the property name / address / county / acreage cells
    # so the M-N subject block renders something meaningful at runtime.
    if prop.get("name"):
        underw["N5"] = prop.get("name")
    if prop.get("address"):
        underw["N6"] = prop.get("address")
    if prop.get("county"):
        underw["N10"] = prop.get("county")

    # Pass the user-supplied county tax rate into P12. The patched
    # template's I22 (RE Taxes stabilized) uses it as the preferred
    # reassessment method when present; otherwise the formula falls
    # back to T12 × 1.15. Accepts a decimal (0.0125) or a percent
    # string ("1.25%").
    rate = to_decimal_pct(prop.get("countyTaxRate"))
    if rate is not None:
        underw["P12"] = rate

    # Atomic save: write to a sibling temp file, then rename. A mid-save
    # crash would otherwise leave the destination half-written, and
    # /api/download would happily ship the corrupt bytes.
    output_path = Path(output_path)
    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    try:
        wb.save(tmp_path)
        os.replace(tmp_path, output_path)
    except Exception:
        # Don't leave a stale .tmp behind on failure.
        try:
            tmp_path.unlink(missing_ok=True)
        except Exception:
            pass
        raise
    return output_path


def add_extraction_check_tab(wb, financials):
    """
    Source-reconciliation tab. Shows what the extraction step pulled from the
    documents and whether it ties out — the verification checkpoint that lets
    the reviewer trust the numbers before reading the underwriting. Placed at
    index 0 so it's the first tab in the workbook.
    """
    ws = wb.create_sheet("Extraction Check", 0)
    ws.sheet_view.showGridLines = False

    NAVY = "1F3864"
    MID_BLUE = "2E5090"
    GREEN = "16A34A"
    AMBER = "D97706"
    RED = "DC2626"
    GRAY = "6B7280"
    LIGHT_YEL = "FFF2CC"
    WHITE = "FFFFFF"
    GREEN_BG = "C6EFCE"
    AMBER_BG = "FFEB9C"
    RED_BG = "FFC7CE"

    def style(cell, value=None, bold=False, color="000000", size=10,
              bg=None, align="left", v_align="center", wrap=False, italic=False):
        if value is not None:
            cell.value = value
        cell.font = Font(name="Arial", bold=bold, color=color, size=size, italic=italic)
        cell.alignment = Alignment(horizontal=align, vertical=v_align, wrap_text=wrap)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)

    for col, w in enumerate([2, 38, 30, 12, 52], 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # Title
    ws.merge_cells("B2:F2")
    style(ws["B2"], "EXTRACTION CHECK — SOURCE RECONCILIATION", bold=True,
          color=WHITE, size=16, bg=NAVY, align="center")
    ws.row_dimensions[2].height = 30
    ws.merge_cells("B3:F3")
    style(ws["B3"], "What the model pulled from the documents, and whether it ties to the source. Review before trusting the underwriting.",
          color=GRAY, size=10, align="center")
    ws.row_dimensions[3].height = 18

    extraction = financials.get("_extraction", {}) or {}
    checks = financials.get("_extractionChecks", []) or []

    # ── Reporting period summary ─────────────────────────────────────────
    rp = extraction.get("reportingPeriod", {}) or {}
    ws.merge_cells("B5:F5")
    style(ws["B5"], "  REPORTING PERIOD USED", bold=True, color=WHITE, size=12, bg=NAVY)
    ws.row_dimensions[5].height = 22

    period_rows = [
        ("Period used", rp.get("periodUsed", "—")),
        ("Date range", rp.get("dateRange", "—")),
        ("Months covered", rp.get("monthsCovered", "—")),
        ("Other periods seen in doc", ", ".join(str(c) for c in (rp.get("candidatePeriodsSeen") or [])) or "—"),
        ("Notes", rp.get("notes", "") or "—"),
    ]
    for i, (label, val) in enumerate(period_rows):
        r = 6 + i
        style(ws.cell(row=r, column=2), label, bold=True, size=10)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)
        style(ws.cell(row=r, column=3), str(val), bg=LIGHT_YEL, color="0000FF", wrap=True)
        ws.row_dimensions[r].height = 18

    # ── Checks table ─────────────────────────────────────────────────────
    checks_start = 6 + len(period_rows) + 1
    ws.merge_cells(start_row=checks_start, start_column=2, end_row=checks_start, end_column=6)
    style(ws.cell(row=checks_start, column=2), "  RECONCILIATION CHECKS", bold=True,
          color=WHITE, size=12, bg=NAVY)
    ws.row_dimensions[checks_start].height = 22

    # Summary counts
    n_ok = sum(1 for c in checks if c.get("status") == "ok")
    n_warn = sum(1 for c in checks if c.get("status") == "warn")
    n_fail = sum(1 for c in checks if c.get("status") == "fail")
    summ_r = checks_start + 1
    ws.merge_cells(start_row=summ_r, start_column=2, end_row=summ_r, end_column=6)
    if n_fail:
        summ_text = f"⚠ {n_fail} FAILED, {n_warn} warnings, {n_ok} passed — review the failures before using these numbers."
        summ_bg = RED
    elif n_warn:
        summ_text = f"{n_warn} warnings, {n_ok} passed — numbers mostly tie out; check the warnings."
        summ_bg = AMBER
    else:
        summ_text = f"✓ All {n_ok} checks passed — extracted numbers tie to the source."
        summ_bg = GREEN
    style(ws.cell(row=summ_r, column=2), summ_text, bold=True, color=WHITE, size=11,
          bg=summ_bg, align="center")
    ws.row_dimensions[summ_r].height = 24

    # Column headers
    hdr_r = summ_r + 1
    for col, label in [(2, "Item"), (3, "Check"), (4, "Status"), (5, "Detail")]:
        style(ws.cell(row=hdr_r, column=col), label, bold=True, color=WHITE,
              size=9, bg=MID_BLUE, align="center")
    ws.row_dimensions[hdr_r].height = 18

    status_map = {
        "ok":   ("✓ OK",   GREEN_BG, GREEN),
        "warn": ("⚠ WARN", AMBER_BG, AMBER),
        "fail": ("✗ FAIL", RED_BG,   RED),
    }
    # Sort so failures float to the top, then warnings, then OK
    order = {"fail": 0, "warn": 1, "ok": 2}
    sorted_checks = sorted(checks, key=lambda c: order.get(c.get("status"), 3))
    for i, c in enumerate(sorted_checks):
        r = hdr_r + 1 + i
        label, bg, fg = status_map.get(c.get("status"), ("?", "FFFFFF", "000000"))
        style(ws.cell(row=r, column=2), c.get("item", ""), size=9, wrap=True)
        style(ws.cell(row=r, column=3), c.get("check", ""), size=9, wrap=True)
        style(ws.cell(row=r, column=4), label, bold=True, color=fg, bg=bg, align="center", size=9)
        style(ws.cell(row=r, column=5), c.get("detail", ""), size=9, wrap=True)
        ws.row_dimensions[r].height = 24

    if not checks:
        r = hdr_r + 1
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        style(ws.cell(row=r, column=2), "No checks were generated for this run.",
              italic=True, color=GRAY, align="center")

    ws.sheet_properties.tabColor = "DC2626"
    return ws


def add_comps_analysis_tab(wb, financials, market):
    """Dedicated comps analysis tab — rent comps, sale comps, statistics,
    subject vs market comparison, and demographics expansion."""
    ws = wb.create_sheet("Comps Analysis", 1)  # Place right after Miscellaneous
    ws.sheet_view.showGridLines = False

    NAVY = "1F3864"
    MID_BLUE = "2E5090"
    GRAY_HDR = "404040"
    LIGHT_YEL = "FFF2CC"
    LIGHT_GRAY = "F2F2F2"
    GREEN_BG = "C6EFCE"
    RED_BG = "FFC7CE"
    WHITE = "FFFFFF"

    def style(cell, value=None, bold=False, color="000000", size=10,
              bg=None, align="left", v_align="center", wrap=False, italic=False, fmt=None):
        if value is not None:
            cell.value = value
        cell.font = Font(name="Arial", bold=bold, color=color, size=size, italic=italic)
        cell.alignment = Alignment(horizontal=align, vertical=v_align, wrap_text=wrap)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)
        if fmt:
            cell.number_format = fmt

    def section_header(row, end_col, title):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=end_col)
        style(ws.cell(row=row, column=2), title, bold=True, color=WHITE, size=12, bg=NAVY)
        ws.row_dimensions[row].height = 22

    def col_header(row, col, label, bg=GRAY_HDR):
        style(ws.cell(row=row, column=col), label, bold=True, color=WHITE,
              size=9, bg=bg, align="center", wrap=True)

    # Column widths — wide enough for all comp data
    widths = [2, 26, 22, 9, 9, 11, 11, 11, 9, 22, 18, 14]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    # Title
    ws.merge_cells("B2:L2")
    style(ws["B2"], "COMPS ANALYSIS — DETAILED COMPETITIVE SET", bold=True,
          color=WHITE, size=16, bg=NAVY, align="center")
    ws.row_dimensions[2].height = 32
    ws.merge_cells("B3:L3")
    style(ws["B3"], "Rent comps, sale comps, statistics, and subject-vs-market positioning",
          color="6B7280", size=10, align="center")
    ws.row_dimensions[3].height = 18

    # ── INVESTMENT CRITERIA CHECK (Michael's 200 bps spread rule) ───────────
    prop_info = financials.get("propertyInfo", {}) or {}
    # Normalize cap-rate-like fields to decimal form (Excel's 0.00% format
    # multiplies by 100). The LLM is inconsistent — a value like 7.5 could
    # mean 7.5% (percent form) or 750% (decimal). to_decimal_pct heuristics
    # divide values >= 1.5 by 100. spread_bps stays in basis points.
    ingoing_cap = to_decimal_pct(prop_info.get("ingoingCapRate"))
    stab_yoc = to_decimal_pct(prop_info.get("stabilizedYieldOnCost"))
    spread_bps = prop_info.get("spreadBps")
    # Recompute spread locally so we don't depend on the LLM's arithmetic
    # being internally consistent with its rate fields.
    if isinstance(ingoing_cap, (int, float)) and isinstance(stab_yoc, (int, float)):
        spread_bps = round((stab_yoc - ingoing_cap) * 10000)
    meets_criteria_local = (
        isinstance(spread_bps, (int, float)) and spread_bps >= 200
    )
    meets_criteria = prop_info.get("meetsInvestmentCriteria")

    crit_start = 5
    section_header(crit_start, 12, "  INVESTMENT CRITERIA CHECK  —  200 bps spread rule")

    # Headers row
    crit_headers = ["Ingoing Cap Rate", "Stabilized Yield on Cost", "Spread (bps)", "Hurdle", "Verdict"]
    for i, h in enumerate(crit_headers):
        col_header(crit_start + 1, 2 + i, h)
    # Span the verdict column wider so the pass/fail box is prominent
    ws.merge_cells(start_row=crit_start + 1, start_column=6, end_row=crit_start + 1, end_column=12)
    col_header(crit_start + 1, 6, "Verdict")
    ws.row_dimensions[crit_start + 1].height = 28

    # Values row
    val_row = crit_start + 2
    style(ws.cell(row=val_row, column=2), ingoing_cap, bg=LIGHT_YEL, color="0000FF",
          align="center", size=14, bold=True, fmt="0.00%")
    style(ws.cell(row=val_row, column=3), stab_yoc, bg=LIGHT_YEL, color="0000FF",
          align="center", size=14, bold=True, fmt="0.00%")
    style(ws.cell(row=val_row, column=4),
          (f"{spread_bps:+,} bps" if isinstance(spread_bps, (int, float)) else "—"),
          bg=LIGHT_YEL, color="0000FF", align="center", size=14, bold=True)
    style(ws.cell(row=val_row, column=5), "≥ 200 bps",
          align="center", size=12, italic=True, color="6B7280")

    # Pass / fail / unknown verdict box — spans cols F through L for visual impact
    ws.merge_cells(start_row=val_row, start_column=6, end_row=val_row, end_column=12)
    # Use the locally recomputed verdict so the displayed status always
    # agrees with the displayed spread, regardless of what the LLM flagged.
    if not isinstance(spread_bps, (int, float)):
        verdict_text = "— INSUFFICIENT DATA TO EVALUATE"
        verdict_bg = "6B7280"
    elif meets_criteria_local:
        verdict_text = "✓ PASSES INVESTMENT CRITERIA"
        verdict_bg = "16A34A"
    else:
        verdict_text = "✗ DOES NOT MEET INVESTMENT CRITERIA"
        verdict_bg = "DC2626"
    style(ws.cell(row=val_row, column=6), verdict_text, bold=True, color=WHITE,
          size=14, bg=verdict_bg, align="center")
    ws.row_dimensions[val_row].height = 42

    # Explanatory subtext row
    explain_row = val_row + 1
    ws.merge_cells(start_row=explain_row, start_column=2, end_row=explain_row, end_column=12)
    if meets_criteria_local and isinstance(spread_bps, (int, float)):
        cushion = spread_bps - 200
        explain_text = (f"Spread is {spread_bps:,} bps — {cushion:+,} bps cushion above the 200 bps hurdle. "
                        f"Deal clears GGC's go/no-go threshold on stabilized yield economics.")
    elif (not meets_criteria_local) and isinstance(spread_bps, (int, float)):
        shortfall = 200 - spread_bps
        explain_text = (f"Spread is {spread_bps:,} bps — {shortfall:,} bps short of the 200 bps hurdle. "
                        f"GGC does not pay for value it's creating; this deal would require either a "
                        f"lower purchase price or a more aggressive stabilized plan to clear.")
    else:
        explain_text = ("One or more inputs missing. Verify the model produced both an ingoing cap rate "
                        "(year-1 underwritten NOI / purchase price) and stabilized yield on cost "
                        "(stabilized NOI / total cost basis incl. CapEx).")
    style(ws.cell(row=explain_row, column=2), explain_text, italic=True, color="374151",
          size=10, wrap=True, v_align="top")
    ws.row_dimensions[explain_row].height = 36

    # Push the "SUBJECT vs MARKET" section down to avoid overlap
    # (it starts at row 5 in the original — shift to row 11)

    rent_comps = market.get("rentComps", []) or []
    sale_comps = market.get("saleComps", []) or []

    # ── SUBJECT vs MARKET POSITIONING ────────────────────────────────────────
    subject_rent = financials.get("rentRoll", {}).get("avgLotRent", 0) or 0
    subject_units = financials.get("propertyInfo", {}).get("totalUnits", 0) or 0
    subject_occ = financials.get("rentRoll", {}).get("occupancyRate", 0) or 0
    asking = financials.get("propertyInfo", {}).get("askingPrice", 0) or 0
    ppu_ask = (asking / subject_units) if subject_units else 0

    # Comp set statistics
    rents = [c.get("lotRent") for c in rent_comps if isinstance(c.get("lotRent"), (int, float))]
    units_list = [c.get("units") for c in rent_comps if isinstance(c.get("units"), (int, float))]
    # Normalize occupancy to decimal (LLM may emit 95 to mean 95%, or 0.95).
    # Same heuristic as to_decimal_pct: values >= 1.5 are presumed percent
    # form and divided by 100. Otherwise % cells show 9500%.
    occ_list = [to_decimal_pct(c.get("occupancy")) for c in rent_comps]
    occ_list = [v for v in occ_list if isinstance(v, (int, float))]

    sale_caps = [to_decimal_pct(c.get("capRate")) for c in sale_comps]
    sale_caps = [v for v in sale_caps if isinstance(v, (int, float))]
    sale_ppu = [c.get("pricePerUnit") for c in sale_comps if isinstance(c.get("pricePerUnit"), (int, float))]

    def safe_min(lst): return min(lst) if lst else None
    def safe_max(lst): return max(lst) if lst else None
    def safe_avg(lst): return (sum(lst) / len(lst)) if lst else None
    def safe_median(lst):
        if not lst: return None
        s = sorted(lst)
        n = len(s)
        return s[n//2] if n % 2 else (s[n//2 - 1] + s[n//2]) / 2

    section_header(11, 12, "  SUBJECT vs MARKET — KEY METRICS")
    headers = ["Metric", "Subject", "Comp Min", "Comp Max", "Comp Avg", "Comp Median", "Subject Position"]
    for i, h in enumerate(headers):
        col_header(12, 2 + i, h)
    ws.row_dimensions[12].height = 28

    def positioning(subj, comp_avg):
        if subj is None or comp_avg is None or comp_avg == 0:
            return ("N/A", "F2F2F2")
        delta = (subj - comp_avg) / comp_avg
        if delta < -0.05:
            return (f"{abs(delta)*100:.1f}% below avg ↑ upside", GREEN_BG)
        elif delta > 0.05:
            return (f"{delta*100:.1f}% above avg", RED_BG)
        else:
            return (f"In line ({delta*100:+.1f}%)", LIGHT_GRAY)

    pos_text, pos_bg = positioning(subject_rent, safe_avg(rents))
    metrics = [
        ("Lot Rent ($/mo)", subject_rent, safe_min(rents), safe_max(rents),
         safe_avg(rents), safe_median(rents), pos_text, pos_bg, "$#,##0"),
        ("# Sites", subject_units, safe_min(units_list), safe_max(units_list),
         safe_avg(units_list), safe_median(units_list), "", LIGHT_GRAY, "#,##0"),
        ("Occupancy", subject_occ, safe_min(occ_list), safe_max(occ_list),
         safe_avg(occ_list), safe_median(occ_list), "", LIGHT_GRAY, "0.0%"),
    ]
    for i, m in enumerate(metrics):
        r = 13 + i
        style(ws.cell(row=r, column=2), m[0], bold=True, size=10)
        for j in range(5):
            val = m[1 + j]
            style(ws.cell(row=r, column=3 + j), val, bg=LIGHT_YEL,
                  color="0000FF", align="right", fmt=m[8])
        style(ws.cell(row=r, column=8), m[6], bold=True, size=10, bg=m[7], align="center")
        ws.row_dimensions[r].height = 18

    # ── RENT COMPS TABLE ─────────────────────────────────────────────────────
    rc_start = 18
    section_header(rc_start, 12, f"  RENT COMPS  ({len(rent_comps)} found)")
    rc_headers = ["Property", "Location", "Distance", "# Sites", "Lot Rent",
                  "Occupancy", "Year Built", "POH %", "Amenities", "Quality", "Source"]
    for i, h in enumerate(rc_headers):
        col_header(rc_start + 1, 2 + i, h)
    ws.row_dimensions[rc_start + 1].height = 28

    for i, c in enumerate(rent_comps[:30]):
        r = rc_start + 2 + i
        loc = f"{c.get('city', '')}, {c.get('state', '')}".strip(", ")
        rating = c.get("qualityRating")
        rating_str = ("★" * int(rating)) if isinstance(rating, (int, float)) and rating else ""
        cells = [
            (c.get("name", ""),         "left",   None),
            (loc,                       "left",   None),
            (c.get("distance", ""),     "center", None),
            (c.get("units", 0),         "center", "#,##0"),
            (c.get("lotRent", 0),       "right",  "$#,##0"),
            (to_decimal_pct(c.get("occupancy")),  "center", "0.0%"),
            (c.get("yearBuilt", ""),    "center", None),
            (to_decimal_pct(c.get("pohPercent")), "center", "0.0%"),
            ((c.get("amenities") or "")[:80], "left", None),
            (rating_str,                "center", None),
            (c.get("source", ""),       "left",   None),
        ]
        for j, (val, align, fmt) in enumerate(cells):
            bg = LIGHT_GRAY if i % 2 == 0 else WHITE
            style(ws.cell(row=r, column=2 + j), val, size=9, bg=bg,
                  align=align, fmt=fmt, wrap=(j == 8))
        ws.row_dimensions[r].height = 30

    # Stats row at the bottom of the rent comps table
    if rent_comps:
        stats_r = rc_start + 2 + len(rent_comps[:30])
        style(ws.cell(row=stats_r, column=2), "STATISTICS", bold=True,
              color=WHITE, bg=MID_BLUE)
        style(ws.cell(row=stats_r, column=3), "", bg=MID_BLUE)
        style(ws.cell(row=stats_r, column=4), "Range:", bold=True, bg=MID_BLUE,
              color=WHITE, align="right")
        if rents:
            style(ws.cell(row=stats_r, column=5),
                  f"{safe_min(units_list) or 0:,.0f}-{safe_max(units_list) or 0:,.0f}",
                  bold=True, color=WHITE, bg=MID_BLUE, align="center")
            style(ws.cell(row=stats_r, column=6),
                  f"${safe_min(rents):,.0f}-${safe_max(rents):,.0f}",
                  bold=True, color=WHITE, bg=MID_BLUE, align="center")
        ws.row_dimensions[stats_r].height = 20

        avg_r = stats_r + 1
        style(ws.cell(row=avg_r, column=4), "Average:", bold=True, color=WHITE,
              bg=MID_BLUE, align="right")
        if units_list:
            style(ws.cell(row=avg_r, column=5),
                  f"{safe_avg(units_list):,.0f}", bold=True, color=WHITE,
                  bg=MID_BLUE, align="center")
        if rents:
            style(ws.cell(row=avg_r, column=6),
                  f"${safe_avg(rents):,.0f}", bold=True, color=WHITE,
                  bg=MID_BLUE, align="center")
        if occ_list:
            style(ws.cell(row=avg_r, column=7),
                  f"{safe_avg(occ_list)*100:.1f}%", bold=True, color=WHITE,
                  bg=MID_BLUE, align="center")

    # ── MARKET RENT CONCLUSION ──────────────────────────────────────────────
    rc_concl_r = rc_start + 4 + len(rent_comps[:30])
    section_header(rc_concl_r, 12, "  MARKET RENT CONCLUSION")
    style(ws.cell(row=rc_concl_r + 1, column=2), market.get("marketRentConclusion", ""),
          size=11, wrap=True, bg=LIGHT_YEL, color="0000FF", v_align="top")
    ws.merge_cells(start_row=rc_concl_r + 1, start_column=2, end_row=rc_concl_r + 1, end_column=12)
    ws.row_dimensions[rc_concl_r + 1].height = 60

    # ── SALE COMPS TABLE ─────────────────────────────────────────────────────
    sc_start = rc_concl_r + 3
    section_header(sc_start, 12, f"  SALE COMPS  ({len(sale_comps)} found)")
    sc_headers = ["Property", "Location", "Sale Date", "# Sites", "Sale Price",
                  "$ / Unit", "Cap Rate", "NOI", "Buyer", "Seller", "Source"]
    for i, h in enumerate(sc_headers):
        col_header(sc_start + 1, 2 + i, h)
    ws.row_dimensions[sc_start + 1].height = 28

    for i, c in enumerate(sale_comps[:30]):
        r = sc_start + 2 + i
        cells = [
            (c.get("name", ""),         "left",   None),
            (c.get("location", ""),     "left",   None),
            (c.get("saleDate", ""),     "center", None),
            (c.get("units", 0),         "center", "#,##0"),
            (c.get("salePrice", 0),     "right",  "$#,##0"),
            (c.get("pricePerUnit", 0),  "right",  "$#,##0"),
            (c.get("capRate"),          "center", "0.00%"),
            (c.get("noi"),              "right",  "$#,##0"),
            (c.get("buyer", ""),        "left",   None),
            (c.get("seller", ""),       "left",   None),
            (c.get("source", ""),       "left",   None),
        ]
        for j, (val, align, fmt) in enumerate(cells):
            bg = LIGHT_GRAY if i % 2 == 0 else WHITE
            style(ws.cell(row=r, column=2 + j), val, size=9, bg=bg, align=align, fmt=fmt)
        ws.row_dimensions[r].height = 18

    # Sale comp statistics
    if sale_comps:
        stats_r = sc_start + 2 + len(sale_comps[:30])
        style(ws.cell(row=stats_r, column=2), "STATISTICS", bold=True,
              color=WHITE, bg=MID_BLUE)
        if sale_ppu:
            style(ws.cell(row=stats_r, column=7),
                  f"${safe_avg(sale_ppu):,.0f}", bold=True, color=WHITE,
                  bg=MID_BLUE, align="center")
            style(ws.cell(row=stats_r, column=6), "Avg $/Unit:", bold=True,
                  color=WHITE, bg=MID_BLUE, align="right")
        if sale_caps:
            style(ws.cell(row=stats_r, column=9),
                  f"{safe_avg(sale_caps)*100:.2f}%", bold=True, color=WHITE,
                  bg=MID_BLUE, align="center")
            style(ws.cell(row=stats_r, column=8), "Avg Cap:", bold=True,
                  color=WHITE, bg=MID_BLUE, align="right")
        ws.row_dimensions[stats_r].height = 20

    # Subject implied valuation row. Requires at least 3 sale comps before
    # surfacing a single "market cap rate" — averaging 1-2 comps is just
    # repackaging one data point as a market signal.
    if sale_caps and subject_units:
        impl_r = sc_start + 4 + len(sale_comps[:30])
        section_header(impl_r, 12, "  IMPLIED VALUATION USING COMP SET")
        if len(sale_caps) < 3:
            note_r = impl_r + 1
            ws.merge_cells(start_row=note_r, start_column=2,
                           end_row=note_r, end_column=12)
            style(ws.cell(row=note_r, column=2),
                  f"INSUFFICIENT DATA — only {len(sale_caps)} sale comp(s) "
                  f"with usable cap rate. GGC requires ≥3 to publish a "
                  f"market cap rate. Sourcing additional comps is "
                  f"recommended before valuation.",
                  bold=True, color="DC2626", bg=LIGHT_YEL, wrap=True, size=10)
            ws.row_dimensions[note_r].height = 36
        else:
            # Trimmed mean: drop the highest and lowest cap rate if N>=5
            # to reduce single-outlier sensitivity.
            sorted_caps = sorted(sale_caps)
            trimmed = sorted_caps[1:-1] if len(sorted_caps) >= 5 else sorted_caps
            avg_cap = safe_avg(trimmed)
            med_cap = safe_median(sale_caps)
            avg_ppu = safe_avg(sale_ppu) if sale_ppu else None
            try:
                import statistics
                std_cap = statistics.pstdev(sale_caps) if len(sale_caps) > 1 else 0
            except Exception:
                std_cap = 0
            labels = [
                ("Asking Price", asking, "$#,##0"),
                ("Asking $ / Unit", ppu_ask, "$#,##0"),
                ("Comp Avg $ / Unit", avg_ppu, "$#,##0"),
                (f"Comp Cap (trimmed mean, N={len(sale_caps)})", avg_cap, "0.00%"),
                ("Comp Cap (median)", med_cap, "0.00%"),
                ("Comp Cap (std dev)", std_cap, "0.00%"),
            ]
            for i, (label, val, fmt) in enumerate(labels):
                r = impl_r + 1 + i
                style(ws.cell(row=r, column=2), label, bold=True, size=10)
                style(ws.cell(row=r, column=3), val, bg=LIGHT_YEL, color="0000FF",
                      align="right", fmt=fmt)
                ws.row_dimensions[r].height = 18

    # ── MARKET CAP RATE CONCLUSION ──────────────────────────────────────────
    cap_concl_r = sc_start + 4 + len(sale_comps[:30]) + (5 if sale_caps else 0)
    section_header(cap_concl_r, 12, "  MARKET CAP RATE CONCLUSION")
    style(ws.cell(row=cap_concl_r + 1, column=2),
          market.get("marketCapRateConclusion", ""),
          size=11, wrap=True, bg=LIGHT_YEL, color="0000FF", v_align="top")
    ws.merge_cells(start_row=cap_concl_r + 1, start_column=2,
                   end_row=cap_concl_r + 1, end_column=12)
    ws.row_dimensions[cap_concl_r + 1].height = 60

    # ── DEMOGRAPHICS DEEP DIVE ───────────────────────────────────────────────
    demo = market.get("demographics", {}) or {}
    dem_start = cap_concl_r + 3
    section_header(dem_start, 12, "  DEMOGRAPHICS — DEEP DIVE")

    demo_rows = [
        ("County", demo.get("countyName", "")),
        ("Total Population", demo.get("countyPopulation")),
        ("1-Year Pop Growth", safe_pct(demo.get("populationGrowth1yr"))),
        ("5-Year Pop Growth", safe_pct(demo.get("populationGrowth5yr"))),
        ("10-Year Pop Growth", safe_pct(demo.get("populationGrowth10yr"))),
        ("5-Yr Population Projection", demo.get("populationProjection5yr")),
        ("Median HH Income", safe_money(demo.get("medianHHIncome"))),
        ("5-Year Income Growth", safe_pct(demo.get("incomeGrowth5yr"))),
        ("Unemployment (Current)", safe_pct(demo.get("unemploymentRate"))),
        ("Unemployment (1yr Ago)", safe_pct(demo.get("unemploymentRate1yrAgo"))),
        ("Poverty Rate", safe_pct(demo.get("povertyRate"))),
        ("% HHs Earning < $50K", safe_pct(demo.get("pctHHUnder50k"))),
        ("Major Industries", ", ".join(demo.get("majorIndustries", []) or [])),
        ("Planned Developments", demo.get("plannedDevelopments", "")),
    ]
    for i, (label, val) in enumerate(demo_rows):
        r = dem_start + 1 + i
        style(ws.cell(row=r, column=2), label, bold=True, size=10)
        style(ws.cell(row=r, column=3), val, bg=LIGHT_YEL, color="0000FF")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=12)
        ws.row_dimensions[r].height = 18

    # Top employers as a separate sub-section
    emp_start = dem_start + 1 + len(demo_rows) + 1
    style(ws.cell(row=emp_start, column=2), "TOP EMPLOYERS",
          bold=True, color=WHITE, bg=MID_BLUE, size=11)
    ws.merge_cells(start_row=emp_start, start_column=2, end_row=emp_start, end_column=12)
    ws.row_dimensions[emp_start].height = 22

    employers = demo.get("majorEmployers", []) or []
    for i, emp in enumerate(employers[:15]):
        r = emp_start + 1 + i
        style(ws.cell(row=r, column=2), f"#{i+1}", bold=True, size=10, align="center")
        style(ws.cell(row=r, column=3), str(emp), bg=LIGHT_YEL, color="0000FF")
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=12)
        ws.row_dimensions[r].height = 18

    ws.sheet_properties.tabColor = "70AD47"  # green tab


def add_miscellaneous_tab(wb, financials, market):
    """Add a Miscellaneous tab at the front with property links, demographics,
    images, landmarks, and amenities — value-add data on top of GGC's template."""
    ws = wb.create_sheet("Miscellaneous", 0)
    ws.sheet_view.showGridLines = False

    NAVY = "1F3864"
    MID_BLUE = "2E5090"
    GRAY_HDR = "404040"
    LIGHT_YEL = "FFF2CC"
    WHITE = "FFFFFF"

    def style_cell(cell, value=None, bold=False, color="000000", size=10,
                   bg=None, align="left", v_align="center", wrap=False, italic=False):
        if value is not None:
            cell.value = value
        cell.font = Font(name="Arial", bold=bold, color=color, size=size, italic=italic)
        cell.alignment = Alignment(horizontal=align, vertical=v_align, wrap_text=wrap)
        if bg:
            cell.fill = PatternFill("solid", fgColor=bg)

    def section_header(row, end_col_letter, title):
        ws.merge_cells(f"B{row}:{end_col_letter}{row}")
        style_cell(ws[f"B{row}"], title, bold=True, color=WHITE, size=12, bg=NAVY)
        ws.row_dimensions[row].height = 22

    for col, w in enumerate([2, 28, 38, 38, 14], 1):
        ws.column_dimensions[get_column_letter(col)].width = w

    ws.merge_cells("B2:E2")
    style_cell(ws["B2"], "PROPERTY OVERVIEW", bold=True, color=WHITE, size=16, bg=NAVY, align="center")
    ws.row_dimensions[2].height = 32
    ws.merge_cells("B3:E3")
    style_cell(ws["B3"], "Location context, visuals, demographics, and qualitative assessment",
               color="6B7280", size=10, align="center")
    ws.row_dimensions[3].height = 18

    prop = financials.get("propertyInfo", {})
    full_addr = f"{prop.get('address','')}, {prop.get('city','')}, {prop.get('state','')}".strip(", ")
    addr_encoded = quote_plus(full_addr) if full_addr else ""

    # PROPERTY LINKS
    section_header(5, "E", "  PROPERTY LINKS")
    links = [
        ("Property Name", prop.get("name", "")),
        ("Address", full_addr),
        ("Property Type", prop.get("propertyType", "")),
        ("Total Units", prop.get("totalUnits", "")),
        ("County", prop.get("county", "")),
        ("Asking Price", prop.get("askingPrice", "")),
        ("CoStar Link", ""),
        ("Property Website", ""),
        ("Broker / OM", ""),
        ("County Assessor", ""),
        ("FEMA Flood Map", ""),
    ]
    for i, (label, val) in enumerate(links):
        r = 6 + i
        style_cell(ws.cell(row=r, column=2), label, bold=True, size=10)
        style_cell(ws.cell(row=r, column=3), val, color="0000FF", bg=LIGHT_YEL)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 18

    # DEMOGRAPHICS
    demo_start = 6 + len(links) + 1
    section_header(demo_start, "E", "  DEMOGRAPHICS")
    demo = market.get("demographics", {}) or {}
    demo_rows = [
        ("County Population", demo.get("countyPopulation", "")),
        ("5-Year Population Growth", safe_pct(demo.get("populationGrowth5yr"))),
        ("Median Household Income", safe_money(demo.get("medianHHIncome"))),
        ("Unemployment Rate", safe_pct(demo.get("unemploymentRate"))),
        ("Top Employers", ", ".join(demo.get("majorEmployers", []) or [])),
    ]
    for i, (label, val) in enumerate(demo_rows):
        r = demo_start + 1 + i
        style_cell(ws.cell(row=r, column=2), label, bold=True, size=10)
        style_cell(ws.cell(row=r, column=3), val, color="0000FF", bg=LIGHT_YEL)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 18

    # ALTERNATIVE HOUSING
    alt_start = demo_start + 1 + len(demo_rows) + 1
    section_header(alt_start, "E", "  ALTERNATIVE HOUSING (affordability)")
    alt = market.get("altHousing", {}) or {}
    alt_rows = [
        ("Avg Single-Family Home Price", safe_money(alt.get("avgSFHomePrice"))),
        ("Avg 1BR Apartment Rent",       safe_money(alt.get("avgRent1BR"), "/mo")),
        ("Avg 2BR Apartment Rent",       safe_money(alt.get("avgRent2BR"), "/mo")),
        ("Avg 3BR Apartment Rent",       safe_money(alt.get("avgRent3BR"), "/mo")),
        ("MHP All-In Monthly Cost",      safe_money(alt.get("mhpAllInCost"), "/mo")),
        ("MHP Savings vs 2BR Apt",       safe_pct(alt.get("mhpVsApartmentSavingsPercent"))),
    ]
    for i, (label, val) in enumerate(alt_rows):
        r = alt_start + 1 + i
        style_cell(ws.cell(row=r, column=2), label, bold=True, size=10)
        style_cell(ws.cell(row=r, column=3), val, color="0000FF", bg=LIGHT_YEL)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 18

    # AERIAL & ROADMAP
    aerial_start = alt_start + 1 + len(alt_rows) + 1
    section_header(aerial_start, "E", "  AERIAL & LOCATION MAP")
    style_cell(ws.cell(row=aerial_start+1, column=2), "Satellite / Aerial",
               bold=True, color=WHITE, size=10, bg=MID_BLUE, align="center")
    ws.merge_cells(start_row=aerial_start+1, start_column=2, end_row=aerial_start+1, end_column=3)
    style_cell(ws.cell(row=aerial_start+1, column=4), "Roadmap (with landmarks)",
               bold=True, color=WHITE, size=10, bg=MID_BLUE, align="center")
    ws.merge_cells(start_row=aerial_start+1, start_column=4, end_row=aerial_start+1, end_column=5)

    IMG_HEIGHT_ROWS = 15
    for r in range(aerial_start + 2, aerial_start + 2 + IMG_HEIGHT_ROWS):
        ws.row_dimensions[r].height = 20

    if full_addr:
        aerial_path = fetch_google_static_map(full_addr, "satellite", 17, "600x400")
        if aerial_path:
            embed_image_in_cell(ws, aerial_path, f"B{aerial_start+2}")
        else:
            style_cell(ws.cell(row=aerial_start+2, column=2),
                       "(set GOOGLE_MAPS_API_KEY to embed images)",
                       italic=True, color="9CA3AF", size=9, align="center")
            ws.merge_cells(start_row=aerial_start+2, start_column=2,
                           end_row=aerial_start+2+IMG_HEIGHT_ROWS-1, end_column=3)

        roadmap_path = fetch_google_static_map(full_addr, "roadmap", 14, "600x400")
        if roadmap_path:
            embed_image_in_cell(ws, roadmap_path, f"D{aerial_start+2}")
        else:
            style_cell(ws.cell(row=aerial_start+2, column=4),
                       "(set GOOGLE_MAPS_API_KEY to embed images)",
                       italic=True, color="9CA3AF", size=9, align="center")
            ws.merge_cells(start_row=aerial_start+2, start_column=4,
                           end_row=aerial_start+2+IMG_HEIGHT_ROWS-1, end_column=5)

    # STREET VIEW (4 images in 2x2)
    sv_start = aerial_start + 2 + IMG_HEIGHT_ROWS + 1
    section_header(sv_start, "E", "  STREET VIEW")
    for col_l, col_r, label in [("B", "C", "Main Entrance"), ("D", "E", "Interior Road")]:
        style_cell(ws.cell(row=sv_start+1, column=ord(col_l)-64), label,
                   bold=True, color=WHITE, size=10, bg=MID_BLUE, align="center")
        ws.merge_cells(start_row=sv_start+1, start_column=ord(col_l)-64,
                       end_row=sv_start+1, end_column=ord(col_r)-64)
    for r in range(sv_start + 2, sv_start + 2 + IMG_HEIGHT_ROWS):
        ws.row_dimensions[r].height = 20
    if full_addr:
        for heading, anchor in [(0, f"B{sv_start+2}"), (90, f"D{sv_start+2}")]:
            sv = fetch_google_streetview(full_addr, heading=heading)
            if sv:
                embed_image_in_cell(ws, sv, anchor)

    sv2_start = sv_start + 2 + IMG_HEIGHT_ROWS + 1
    for col_l, col_r, label in [("B", "C", "Example Home #1"), ("D", "E", "Example Home #2")]:
        style_cell(ws.cell(row=sv2_start, column=ord(col_l)-64), label,
                   bold=True, color=WHITE, size=10, bg=MID_BLUE, align="center")
        ws.merge_cells(start_row=sv2_start, start_column=ord(col_l)-64,
                       end_row=sv2_start, end_column=ord(col_r)-64)
    for r in range(sv2_start + 1, sv2_start + 1 + IMG_HEIGHT_ROWS):
        ws.row_dimensions[r].height = 20
    if full_addr:
        for heading, anchor in [(180, f"B{sv2_start+1}"), (270, f"D{sv2_start+1}")]:
            sv = fetch_google_streetview(full_addr, heading=heading)
            if sv:
                embed_image_in_cell(ws, sv, anchor)

    # CLICKABLE URLs
    url_start = sv2_start + 1 + IMG_HEIGHT_ROWS + 1
    section_header(url_start, "E", "  CLICKABLE URLs")
    visuals = market.get("visuals", {}) or {}
    fb_aerial = f"https://www.google.com/maps/search/?api=1&query={addr_encoded}&t=k" if addr_encoded else ""
    fb_street = f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={addr_encoded}" if addr_encoded else ""
    fb_dir = f"https://www.google.com/maps/dir/?api=1&destination={addr_encoded}" if addr_encoded else ""
    url_rows = [
        ("Aerial / Satellite", visuals.get("aerialView") or fb_aerial),
        ("Roadmap", fb_aerial.replace("&t=k", "") if fb_aerial else ""),
        ("Street View — Main Entrance", visuals.get("streetViewEntrance") or fb_street),
        ("Street View — Interior Road", visuals.get("streetViewInterior") or fb_street),
        ("Street View — Example Home #1", visuals.get("exampleHome1") or fb_street),
        ("Street View — Example Home #2", visuals.get("exampleHome2") or fb_street),
        ("Driving Directions", visuals.get("directions") or fb_dir),
    ]
    for i, (label, url) in enumerate(url_rows):
        r = url_start + 1 + i
        style_cell(ws.cell(row=r, column=2), label, bold=True, size=10)
        style_cell(ws.cell(row=r, column=3), url, color="0000FF", bg=LIGHT_YEL)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 18

    # LANDMARK DISTANCES
    lm_start = url_start + 1 + len(url_rows) + 1
    section_header(lm_start, "E", "  PROXIMITY TO KEY LANDMARKS")
    style_cell(ws.cell(row=lm_start+1, column=2), "Landmark Type",
               bold=True, color=WHITE, size=9, bg=GRAY_HDR, align="center")
    style_cell(ws.cell(row=lm_start+1, column=3), "Name / Address",
               bold=True, color=WHITE, size=9, bg=GRAY_HDR, align="center")
    ws.merge_cells(start_row=lm_start+1, start_column=3, end_row=lm_start+1, end_column=4)
    style_cell(ws.cell(row=lm_start+1, column=5), "Distance (mi)",
               bold=True, color=WHITE, size=9, bg=GRAY_HDR, align="center")

    # Fuzzy match: Claude often returns "Walmart" instead of "Nearest Walmart / Big Box"
    # Match by keyword presence rather than exact string equality
    raw_landmarks = market.get("landmarks") or []
    def find_landmark(target_label):
        target_lower = target_label.lower()
        # Extract key tokens from the target (skip "nearest", articles, etc.)
        target_keywords = [w for w in re.findall(r"\w+", target_lower)
                          if w not in {"nearest", "the", "a", "of", "or", "and", "/"}]
        # Score each AI-returned landmark by keyword overlap
        best_match = None
        best_score = 0
        for lm in raw_landmarks:
            lm_text = (lm.get("type", "") + " " + lm.get("name", "")).lower()
            score = sum(1 for kw in target_keywords if kw in lm_text)
            # Boost score for exact "type" match
            if lm.get("type", "").lower() == target_lower:
                score += 100
            if score > best_score:
                best_score = score
                best_match = lm
        return best_match if best_score > 0 else {}

    landmark_types = [
        "Nearest Major Employer #1", "Nearest Major Employer #2", "Nearest Major Employer #3",
        "Nearest Walmart / Big Box", "Nearest Grocery Store", "Nearest Hospital / Urgent Care",
        "Nearest Elementary School", "Nearest High School",
        "Nearest Community College / University", "Nearest Major Highway / Interstate",
        "Downtown (nearest city)", "Nearest Airport",
    ]
    used_landmarks = set()
    for i, lm_type in enumerate(landmark_types):
        r = lm_start + 2 + i
        lm = find_landmark(lm_type)
        # Avoid using the same landmark twice
        lm_id = id(lm) if lm else None
        if lm_id in used_landmarks:
            lm = {}
        else:
            if lm_id:
                used_landmarks.add(lm_id)
        style_cell(ws.cell(row=r, column=2), lm_type, bold=True, size=10)
        style_cell(ws.cell(row=r, column=3), lm.get("name", ""), color="0000FF", bg=LIGHT_YEL)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        style_cell(ws.cell(row=r, column=5), lm.get("distanceMiles"),
                   color="0000FF", bg=LIGHT_YEL, align="right")
        ws.row_dimensions[r].height = 18

    # DILIGENCE FLAGS
    flags_start = lm_start + 2 + len(landmark_types) + 1
    section_header(flags_start, "E", "  DILIGENCE FLAGS")
    flags = financials.get("flags", []) or []
    for i, f in enumerate(flags[:10]):
        r = flags_start + 1 + i
        sev = (f.get("severity", "") or "").lower()
        sev_color = {"high": "DC2626", "medium": "D97706", "low": "16A34A"}.get(sev, "6B7280")
        style_cell(ws.cell(row=r, column=2), f.get("severity", "").upper(),
                   bold=True, color=WHITE, size=9, bg=sev_color, align="center")
        msg = f"{f.get('item', '')}: {f.get('issue', '')} — Recommendation: {f.get('recommendation', '')}"
        style_cell(ws.cell(row=r, column=3), msg, size=10, wrap=True)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 36

    # QUESTIONS FOR SELLER
    q_start = flags_start + 1 + len(flags[:10]) + 1
    section_header(q_start, "E", "  QUESTIONS FOR SELLER / BROKER")
    questions = financials.get("questions", []) or []
    for i, q in enumerate(questions[:10]):
        r = q_start + 1 + i
        style_cell(ws.cell(row=r, column=2), f"#{i+1}", bold=True, size=10, align="center")
        style_cell(ws.cell(row=r, column=3), q, size=10, wrap=True)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        ws.row_dimensions[r].height = 30

    ws.sheet_properties.tabColor = "F59E0B"


# ═══════════════════════════════════════════════════════════════════════════
# JOB ORCHESTRATION
# ═══════════════════════════════════════════════════════════════════════════
def _set_job(job_id, **fields):
    """Atomically update a job's mutable fields under JOBS_LOCK."""
    with JOBS_LOCK:
        if job_id in JOBS:
            JOBS[job_id].update(fields)


def run_analysis_job(job_id, api_key, file_blocks, property_info):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
    if job is None:
        return
    try:
        _set_job(job_id, status="running", progress="Starting analysis...")

        deep_search = property_info.get("deepSearch", "off") == "on"
        market_fn = (lambda *a, **k: call_market_research_merged(*a, **k, n_runs=3)) if deep_search else call_market_research

        # The financial side is now a 4-step sequence with two opt-in stages:
        #   0. CACHE LOOKUP                   — fingerprint hit returns instantly
        #   1. EXTRACT (Sonnet 4.6, temp=0)   — N=1 default, N=3 when deep_search
        #      → field-level median across N runs
        #   2. VERIFY  (pure Python)          — tie-outs, 2σ rents, POH, cross-doc
        #   3. METHODOLOGY (Opus 4.8)         — GGC categorization + underwriting
        #   4. CACHE WRITE                    — if no hard-fail in verification
        # Market research is independent, so we run it in parallel with the
        # whole financial sequence.
        n_extract_runs = FINANCIAL_PARSE_RUNS_DEEP if deep_search else 1

        def financial_pipeline():
            cache_key = extraction_cache_key(
                file_blocks, property_info, n_extract_runs,
                EXTRACTION_PROMPT, FINANCIAL_PARSE_PROMPT)
            cached = extraction_cache_get(cache_key)
            if cached is not None:
                print(f"[Cache] HIT key={cache_key[:8]}... — returning memoized "
                      "financials (no Claude calls this run)")
                cached.setdefault("_cache", {})
                cached["_cache"].update({"hit": True, "key": cache_key[:8]})
                return cached
            print(f"[Cache] MISS key={cache_key[:8]} "
                  f"(deep_search={deep_search}, n_extract_runs={n_extract_runs})")

            if n_extract_runs > 1:
                extracted = call_extract_financials_merged(
                    api_key, file_blocks, property_info, n_runs=n_extract_runs)
            else:
                extracted = call_extract_financials(
                    api_key, file_blocks, property_info)
            checks = verify_extraction(extracted, property_info)
            n_fail = sum(1 for c in checks if c["status"] == "fail")
            n_warn = sum(1 for c in checks if c["status"] == "warn")
            print(f"[Verify] {len(checks)} checks: {n_fail} fail, {n_warn} warn")
            financials = call_parse_financials(api_key, extracted, property_info)
            # Carry the user-provided county tax rate through into
            # financials.propertyInfo so fill_template can stamp it into
            # the Underwriting tab (P12) — the RE Taxes override formula
            # uses it as the preferred reassessment method.
            if property_info.get("countyTaxRate"):
                financials.setdefault("propertyInfo", {})["countyTaxRate"] = \
                    property_info.get("countyTaxRate")
            # Methodology-side checks run AFTER categorization because they
            # need both the income.ggcCategory tags and the rent roll's
            # canonical unit types. This is where the lot-rent / RV-rent
            # collapse bug would surface.
            checks.extend(verify_methodology(financials))
            # Carry the raw extraction + checks through so they can be rendered
            # on the Extraction Check tab.
            financials["_extraction"] = extracted
            financials["_extractionChecks"] = checks
            financials["_cache"] = {"hit": False, "key": cache_key[:8],
                                     "n_extract_runs": n_extract_runs}

            # Only memoize a clean result. If hard fails remain we want the
            # next run to re-try, not return the broken cached output.
            if n_fail == 0:
                extraction_cache_put(cache_key, financials)
            else:
                print(f"[Cache] Skipping write — {n_fail} hard failures remain")
            return financials

        with ThreadPoolExecutor(max_workers=2) as executor:
            future_market = executor.submit(market_fn, api_key, property_info)
            future_financials = executor.submit(financial_pipeline)

            results = {}
            for future in as_completed([future_financials, future_market]):
                if future is future_financials:
                    results["financials"] = future.result()
                    _set_job(job_id, progress="✓ Financials extracted, verified, and underwritten.")
                else:
                    results["market"] = future.result()
                    _set_job(job_id, progress="✓ Market research complete.")

        _set_job(job_id, progress="Filling GGC template...")
        output_path = JOBS_DIR / f"{job_id}.xlsx"
        fill_template(results["financials"], results["market"], output_path)

        _set_job(job_id,
                 status="complete",
                 progress="Done.",
                 result={
                     "financials": results["financials"],
                     "market": results["market"],
                     "download_url": f"/api/download/{job_id}",
                 })
    except Exception as e:
        # Log full traceback server-side; the /api/status response is
        # sanitized so the client never sees internal error text.
        traceback.print_exc()
        _set_job(job_id, status="error",
                 error=str(e)[:200],
                 progress="Error: analysis failed — check server logs.")


# ═══════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════
@app.route("/")
def root():
    return send_from_directory(".", "index.html")


@app.route("/api/config")
def config():
    """Frontend pulls this on page load to learn what optional integrations
    the server has configured. The server's Anthropic key is NEVER returned
    over the wire — the frontend either provides its own key per request or
    falls back to the server key only inside `run_analysis_job`. Returning
    a key here would let any unauthenticated visitor read GGC's billing
    credential."""
    return jsonify({
        "default_api_key_present": bool(DEFAULT_ANTHROPIC_KEY),
        "google_maps_enabled": bool(GOOGLE_MAPS_API_KEY),
    })


@app.route("/api/analyze", methods=["POST"])
def analyze():
    api_key = (request.form.get("api_key") or "").strip() or DEFAULT_ANTHROPIC_KEY
    if not api_key:
        return jsonify({"error": "API key required"}), 400

    property_info = {
        "name":        request.form.get("property_name", ""),
        "address":     request.form.get("address", ""),
        "city":        request.form.get("city", ""),
        "county":      request.form.get("county", ""),
        "countyTaxRate": request.form.get("county_tax_rate", ""),
        "pohCount":    request.form.get("poh_count", "0"),
        "state":       request.form.get("state", ""),
        "units":       request.form.get("units", ""),
        "askingPrice": request.form.get("asking_price", ""),
        "floodZone":   request.form.get("flood_zone", "unknown"),
        "deepSearch":  request.form.get("deep_search", "off"),
    }

    if not property_info["city"] or not property_info["state"]:
        return jsonify({"error": "City and State are required for market research"}), 400

    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "At least one file required"}), 400

    # Whitelist file extensions; reject anything else before reading bytes.
    for f in files:
        ext = (Path(f.filename or "").suffix or "").lower()
        if ext not in ALLOWED_UPLOAD_EXTS:
            return jsonify({"error": f"Unsupported file type: {ext or '<none>'}. "
                                     f"Allowed: {sorted(ALLOWED_UPLOAD_EXTS)}"}), 400

    file_blocks = [encode_file_for_claude(f) for f in files]
    job_id = _new_job_id()
    with JOBS_LOCK:
        JOBS[job_id] = {"status": "queued", "progress": "Queued", "result": None}
        _evict_old_jobs()

    Thread(target=run_analysis_job, args=(job_id, api_key, file_blocks, property_info),
           daemon=True).start()

    return jsonify({"job_id": job_id})


@app.route("/api/status/<job_id>")
def status(job_id):
    # Validate the path component before any dict / filesystem lookup so
    # /api/status/../../etc/passwd can't even probe state.
    if not _valid_job_id(job_id):
        return jsonify({"error": "Invalid job id"}), 400
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "Job not found"}), 404
        # Return progress fields + the full analysis result (financials,
        # market, download_url) — the frontend's showResults() reads
        # result.financials / result.market to populate the KPI cards.
        # Internal error strings are still scrubbed: a server-side
        # exception text may include prompt snippets or API-response
        # excerpts and isn't useful to the client anyway.
        public = {
            "status":   job.get("status"),
            "progress": job.get("progress"),
            "result":   job.get("result") if isinstance(job.get("result"), dict) else None,
        }
        if job.get("status") == "error":
            public["error"] = "Analysis failed — check server logs."
    return jsonify(public)


@app.route("/api/download/<job_id>")
def download(job_id):
    if not _valid_job_id(job_id):
        return jsonify({"error": "Invalid job id"}), 400
    file_path = JOBS_DIR / f"{job_id}.xlsx"
    # Defense in depth: even with the regex guard, resolve and verify the
    # final path stays inside JOBS_DIR before serving.
    try:
        resolved = file_path.resolve(strict=True)
        if JOBS_DIR.resolve() not in resolved.parents:
            return jsonify({"error": "Invalid job id"}), 400
    except FileNotFoundError:
        return jsonify({"error": "File not ready"}), 404
    with JOBS_LOCK:
        name = (JOBS.get(job_id, {}).get("result", {}) or {}) \
            .get("financials", {}).get("propertyInfo", {}).get("name", "Property")
    safe = re.sub(r"[^A-Za-z0-9_-]", "_", name)[:40] or "Property"
    return send_file(file_path, as_attachment=True, download_name=f"GGC_UW_{safe}.xlsx")


if __name__ == "__main__":
    print(" ╔═════════════════════════════════════════════════════════════════════════════════════╗")
    print(" ║  GGC Deal Engine — Backend Server v6 (playbook hardening)                           ║")
    print(f"║  Extraction:   {MODEL_EXTRACTION:<48s}                                 ║")
    print(f"║  Methodology:  {MODEL_METHODOLOGY:<48s}                                 ║")
    print(f"║  Market:       {MODEL_MARKET:<48s}                                 ║")
    print(f"║  Pipeline: parse → extract (N={FINANCIAL_PARSE_RUNS_DEEP} deep) → verify → methodology → cache → fill ║")
    print(f"║  Parser:       {PARSER_BACKEND:<48s}                                 ║")
    print(f"║  Structured outputs (beta):  {'ENABLED' if USE_STRUCTURED_OUTPUTS else 'DISABLED':<48s}             ║")
    print(f"║  Extraction cache:           {'ENABLED' if EXTRACTION_CACHE_ENABLED else 'DISABLED':<48s}             ║")
    print(f"║  Validator retries:          {MAX_PARSE_RETRIES}                                                      ║")
    print(f"║  Template: GGC_Blank_Underwriting_Sizer_Extended (1000 rows)                        ║")
    print(f"║  Google Maps: {'ENABLED' if GOOGLE_MAPS_API_KEY else 'DISABLED (no key set)':<43s}  ║")
    print(f"║  Document AI: {'ENABLED' if DOC_AI_ENABLED else 'DISABLED (no GCP config)':<43s} ║")
    print(" ║  Open: http://localhost:5001                                                        ║")
    print(" ╚═════════════════════════════════════════════════════════════════════════════════════╝")
    app.run(host="0.0.0.0", port=5001, debug=False, threaded=True)
