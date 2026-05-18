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
import traceback
import re
from pathlib import Path
from dotenv import load_dotenv
from threading import Thread
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from urllib.parse import quote_plus

import requests
from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from dotenv import load_dotenv

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════════════
ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
MODEL_FINANCIAL   = "claude-sonnet-4-6"   # deterministic with temperature=0
MODEL_MARKET      = "claude-sonnet-4-6"     # adaptive thinking for open-ended research
API_VERSION       = "2023-06-01"
MAX_TOKENS        = 32000  # bumped from 16k — large rent rolls + comp lists need headroom
MAX_RETRIES       = 6
BASE_BACKOFF_SEC  = 2

GOOGLE_MAPS_API_KEY = os.environ.get("GOOGLE_MAPS_API_KEY", "***REMOVED_GOOGLE_MAPS_KEY***")
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

# IMPORTANT: GGC's official blank template, extended to 1000 rent roll rows
TEMPLATE_PATH = Path(__file__).parent / "GGC_Blank_Underwriting_Sizer_Extended.xlsx"
JOBS_DIR      = Path(__file__).parent.parent / "jobs"
IMG_CACHE_DIR = Path(__file__).parent.parent / "img_cache"
JOBS_DIR.mkdir(exist_ok=True)
IMG_CACHE_DIR.mkdir(exist_ok=True)

# GGC's exact category strings — must match column A in Data Consolidation
# (these feed the SUMIFS in the GGC Underwriting tab)
GGC_INCOME_CATEGORIES = [
    "Gross Potential Rent", "Less: Vacancy", "Less: Concessions", "Less: Bad Debt",
    "Utility Reimbursement", "Home Rent Income", "LTO income", "SFH",
    "Laundry Income", "Other Income", "Employee Allowance", "Model Units",
]

GGC_EXPENSE_CATEGORIES = [
    "RE Taxes", "Insurance", "Gas/Fuel", "Electrcitiy",  # GGC's spelling
    "Water and Sewer", "Trash Removal", "Repair and Maintenance",
    "Ground Maintenance", "Recreational Amenities", "Management Fee",
    "Payroll", "General and Administrative", "Professional Fees",
    "Advertising", "Home Rent Expense (MH)", "Other", "Cap-Ex Reserve",
]

JOBS = {}

app = Flask(__name__, static_folder=".", static_url_path="")
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
                use_thinking=True, temperature=None, model=None):
    """
    Call Claude with streaming enabled. Streaming keeps the connection alive
    during long-running requests (which can hit 3+ minutes when Claude is doing
    heavy thinking + web search) instead of timing out at the request level.

    Model routing:
    - Defaults to MODEL_MARKET (Opus 4.7) for market research with adaptive thinking
    - Pass model=MODEL_FINANCIAL (Sonnet 4.6) for deterministic financial parsing
      with temperature=0

    NOTE: Opus 4.7 deprecated temperature/top_p/top_k entirely. Only pass
    temperature when targeting Sonnet 4.6 (or other pre-4.7 models).
    """
    headers = {"x-api-key": api_key, "anthropic-version": API_VERSION,
               "content-type": "application/json"}
    body = {"model": model or MODEL_MARKET,
            "max_tokens": MAX_TOKENS,
            "system": system_prompt,
            "messages": [{"role": "user", "content": user_content}],
            "stream": True,}
    if use_thinking:
        body["thinking"] = {"type": "adaptive"}
        body["output_config"] = {"effort": "high"}
    elif temperature is not None:
        body["temperature"] = temperature
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
        return {"type": "document",
                "source": {"type": "base64", "media_type": "application/pdf", "data": b64},
                "title": filename}
    elif ext in (".png", ".jpg", ".jpeg", ".webp", ".gif"):
        media_type = f"image/{'jpeg' if ext == '.jpg' else ext[1:]}"
        return {"type": "image",
                "source": {"type": "base64", "media_type": media_type, "data": b64}}
    elif ext in (".xlsx", ".xls"):
        try:
            wb = load_workbook(BytesIO(data), data_only=True)
            text_parts = [f"[Spreadsheet: {filename}]"]
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                text_parts.append(f"\n## Sheet: {sheet_name}")
                for row in ws.iter_rows(values_only=True):
                    if any(c is not None for c in row):
                        text_parts.append("\t".join(str(c) if c is not None else "" for c in row))
            return {"type": "text", "text": "\n".join(text_parts)[:200000]}
        except Exception as e:
            return {"type": "text", "text": f"[Could not parse {filename}: {e}]"}
    elif ext in (".txt", ".csv", ".md"):
        return {"type": "text",
                "text": f"[File: {filename}]\n{data.decode('utf-8', errors='replace')[:200000]}"}
    else:
        return {"type": "text",
                "text": f"[File: {filename}]\n{data.decode('utf-8', errors='replace')[:100000]}"}


# ═══════════════════════════════════════════════════════════════════════════
# CLAUDE CALL #1 — Financial Parsing
# ═══════════════════════════════════════════════════════════════════════════
FINANCIAL_PARSE_PROMPT = f"""You are a real estate underwriting analyst at Gary Group Capital (GGC), a private equity firm focused on mobile home parks.

Parse the attached seller financials and map every line item to GGC's standardized categories.

CRITICAL OUTPUT RULES:
- Your response MUST start with `{{` (a JSON open brace) — NOTHING before it
- Your response MUST end with `}}` (a JSON close brace) — NOTHING after it
- DO NOT add preamble, explanation, "I'll analyze...", or commentary of any kind
- DO NOT wrap the JSON in markdown code fences
- If the user-provided property info has typos or inconsistencies, use your best interpretation silently — DO NOT explain corrections in the output
- Notes should go in the "notes" field of each line item, NEVER as freestanding text

## GGC Income Categories (use EXACTLY these strings):
{json.dumps(GGC_INCOME_CATEGORIES)}

## GGC Expense Categories (use EXACTLY these strings — note 'Electrcitiy' typo is intentional, GGC uses it in their model):
{json.dumps(GGC_EXPENSE_CATEGORIES)}

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
    "unitGroups": [
      {{"unitType": "string (e.g. Standard Lot, Premium Lot, POH, RV Annual)",
        "occupiedCount": integer,
        "vacantCount": integer,
        "lotRent": number,
        "pohRent": number (0 if TOH),
        "ltoPremium": number (0 if not LTO),
        "tenantNamePattern": "string (e.g. 'Tenant', use this prefix + sequential number for each occupied unit)"}}
    ],
    "unitMixSummary": [
      {{"unitType": "string", "count": integer, "occupied": integer,
        "avgRent": number, "marketRent": number}}
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

CRITICAL: For "unitGroups", aggregate by unit type. Do NOT list every unit individually — group them. Example: instead of listing 150 separate "Standard Lot" entries, return one entry with occupiedCount=140, vacantCount=10. The backend will expand groups into individual rows automatically."""


def call_parse_financials(api_key, file_blocks, property_info):
    user_blocks = file_blocks + [{
        "type": "text",
        "text": f"""Property context:
- Name: {property_info.get('name', 'N/A')}
- Address: {property_info.get('address', 'N/A')}
- Total Units: {property_info.get('units', 'N/A')}
- Park-Owned Home Count (user-stated): {property_info.get('pohCount', '0')}
- Asking Price: ${property_info.get('askingPrice', 'N/A')}
- Flood Zone Status: {property_info.get('floodZone', 'unknown')}

Parse the attached documents and return the structured JSON."""
    }]
    print("[Claude] Starting financial parsing call...")
    t0 = time.time()
    response = call_claude(api_key, FINANCIAL_PARSE_PROMPT, user_blocks, use_thinking=False, temperature=0, model=MODEL_FINANCIAL)
    elapsed = time.time() - t0
    print(f"[Claude] Financial parsing returned in {elapsed:.1f}s "
          f"(stop_reason: {response.get('stop_reason', '?')})")
    return extract_json(response)


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
                            [{"type": "text", "text": prompt}], tools=tools)
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
    merged_rent_comps = sorted(rent_by_name.values(),
                                key=lambda c: c.get("distance", "999"))

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
    merged_sale_comps = sorted(sale_by_key.values(),
                                key=lambda c: c.get("saleDate") or "")

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
        monthly = item.get("monthly", [])
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
        monthly = item.get("monthly", [])
        if len(monthly) == 12:
            for m_i, val in enumerate(monthly):
                ws.cell(row=r, column=10 + m_i, value=val)
        elif item.get("t12Total"):
            even = (item["t12Total"] or 0) / 12
            for m_i in range(12):
                ws.cell(row=r, column=10 + m_i, value=even)

    # ── Rent Roll Input ────────────────────────────────────────────────────
    # Cols: A=Count (skip — has formula), B=Unit Type, C=Status, D=Tenant,
    #       E=Lot Rent, F=POH Rent, G=LTO Premium, H=Combined (formula — skip)
    # Data rows 3-1002 (1000 unit slots in extended template)
    # Expand unit groups into individual rows.
    ws = wb["Rent Roll Input"]
    rr = financials.get("rentRoll", {})
    unit_groups = rr.get("unitGroups", [])

    individual_units = []
    for grp in unit_groups:
        ut = grp.get("unitType", "Unit")
        lot_rent = grp.get("lotRent", 0) or 0
        poh_rent = grp.get("pohRent", 0) or 0
        lto_premium = grp.get("ltoPremium", 0) or 0
        name_prefix = grp.get("tenantNamePattern", "Tenant")
        occ_count = grp.get("occupiedCount", 0) or 0
        vac_count = grp.get("vacantCount", 0) or 0

        for i in range(occ_count):
            individual_units.append({
                "unitType": ut, "status": "Occupied",
                "tenantName": f"{name_prefix} {len(individual_units) + 1}",
                "lotRent": lot_rent, "pohRent": poh_rent, "ltoPremium": lto_premium,
            })
        for _ in range(vac_count):
            individual_units.append({
                "unitType": ut, "status": "Vacant",
                "tenantName": "", "lotRent": 0, "pohRent": 0, "ltoPremium": 0,
            })

    for i, unit in enumerate(individual_units[:1000]):
        r = 3 + i
        ws.cell(row=r, column=2, value=unit.get("unitType", ""))
        ws.cell(row=r, column=3, value=unit.get("status", ""))
        ws.cell(row=r, column=4, value=unit.get("tenantName", ""))
        ws.cell(row=r, column=5, value=unit.get("lotRent", 0))
        ws.cell(row=r, column=6, value=unit.get("pohRent", 0))
        ws.cell(row=r, column=7, value=unit.get("ltoPremium", 0))

    # ── Add Miscellaneous tab ──────────────────────────────────────────────
    if "Miscellaneous" in wb.sheetnames:
        del wb["Miscellaneous"]
    add_miscellaneous_tab(wb, financials, market)

    # ── Add Comps Analysis tab ─────────────────────────────────────────────
    if "Comps Analysis" in wb.sheetnames:
        del wb["Comps Analysis"]
    add_comps_analysis_tab(wb, financials, market)

    wb.save(output_path)
    return output_path


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
    ingoing_cap = prop_info.get("ingoingCapRate")
    stab_yoc = prop_info.get("stabilizedYieldOnCost")
    spread_bps = prop_info.get("spreadBps")
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
    if meets_criteria is True:
        verdict_text = "✓ PASSES INVESTMENT CRITERIA"
        verdict_bg = "16A34A"   # green
    elif meets_criteria is False:
        verdict_text = "✗ DOES NOT MEET INVESTMENT CRITERIA"
        verdict_bg = "DC2626"   # red
    else:
        verdict_text = "— INSUFFICIENT DATA TO EVALUATE"
        verdict_bg = "6B7280"   # gray
    style(ws.cell(row=val_row, column=6), verdict_text, bold=True, color=WHITE,
          size=14, bg=verdict_bg, align="center")
    ws.row_dimensions[val_row].height = 42

    # Explanatory subtext row
    explain_row = val_row + 1
    ws.merge_cells(start_row=explain_row, start_column=2, end_row=explain_row, end_column=12)
    if meets_criteria is True and isinstance(spread_bps, (int, float)):
        cushion = spread_bps - 200
        explain_text = (f"Spread is {spread_bps:,} bps — {cushion:+,} bps cushion above the 200 bps hurdle. "
                        f"Deal clears GGC's go/no-go threshold on stabilized yield economics.")
    elif meets_criteria is False and isinstance(spread_bps, (int, float)):
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
    occ_list = [c.get("occupancy") for c in rent_comps if isinstance(c.get("occupancy"), (int, float))]

    sale_caps = [c.get("capRate") for c in sale_comps if isinstance(c.get("capRate"), (int, float))]
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
            (c.get("occupancy"),        "center", "0.0%"),
            (c.get("yearBuilt", ""),    "center", None),
            (c.get("pohPercent"),       "center", "0.0%"),
            (c.get("amenities", "")[:80], "left", None),
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

    # Subject implied valuation row
    if sale_caps and subject_units:
        impl_r = sc_start + 4 + len(sale_comps[:30])
        section_header(impl_r, 12, "  IMPLIED VALUATION USING COMP SET")
        avg_cap = safe_avg(sale_caps)
        avg_ppu = safe_avg(sale_ppu) if sale_ppu else None
        labels = [
            ("Asking Price", asking, "$#,##0"),
            ("Asking $ / Unit", ppu_ask, "$#,##0"),
            ("Comp Avg $ / Unit", avg_ppu, "$#,##0"),
            ("Comp Avg Cap Rate", avg_cap, "0.00%"),
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
def run_analysis_job(job_id, api_key, file_blocks, property_info):
    job = JOBS[job_id]
    try:
        job["status"] = "running"
        job["progress"] = "Starting parallel analysis..."

        deep_search = property_info.get("deepSearch", "off") == "on"
        market_fn = (lambda *a, **k: call_market_research_merged(*a, **k, n_runs=3)) if deep_search else call_market_research

        with ThreadPoolExecutor(max_workers=2) as executor:
            future_financials = executor.submit(call_parse_financials, api_key, file_blocks, property_info)
            future_market = executor.submit(market_fn, api_key, property_info)


            results = {}
            for future in as_completed([future_financials, future_market]):
                if future is future_financials:
                    results["financials"] = future.result()
                    job["progress"] = "✓ Financials parsed."
                else:
                    results["market"] = future.result()
                    job["progress"] = "✓ Market research complete."

        job["progress"] = "Filling GGC template..."
        output_path = JOBS_DIR / f"{job_id}.xlsx"
        fill_template(results["financials"], results["market"], output_path)

        job["status"] = "complete"
        job["progress"] = "Done."
        job["result"] = {
            "financials": results["financials"],
            "market": results["market"],
            "download_url": f"/api/download/{job_id}",
        }
    except Exception as e:
        traceback.print_exc()
        job["status"] = "error"
        job["error"] = str(e)
        job["progress"] = f"Error: {e}"


# ═══════════════════════════════════════════════════════════════════════════
# ROUTES
# ═══════════════════════════════════════════════════════════════════════════
@app.route("/")
def root():
    return send_from_directory(".", "index.html")


@app.route("/api/config")
def config():
    """Frontend pulls this on page load to get the default API key (if set in backend.py)."""
    return jsonify({
        "default_api_key": DEFAULT_ANTHROPIC_KEY,
        "google_maps_enabled": bool(GOOGLE_MAPS_API_KEY),
    })


@app.route("/api/analyze", methods=["POST"])
def analyze():
    api_key = request.form.get("api_key", "").strip()
    if not api_key:
        return jsonify({"error": "API key required"}), 400

    property_info = {
        "name":        request.form.get("property_name", ""),
        "address":     request.form.get("address", ""),
        "city":        request.form.get("city", ""),
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

    file_blocks = [encode_file_for_claude(f) for f in files]
    job_id = str(int(time.time() * 1000))
    JOBS[job_id] = {"status": "queued", "progress": "Queued", "result": None}

    Thread(target=run_analysis_job, args=(job_id, api_key, file_blocks, property_info),
           daemon=True).start()

    return jsonify({"job_id": job_id})


@app.route("/api/status/<job_id>")
def status(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify(job)


@app.route("/api/download/<job_id>")
def download(job_id):
    file_path = JOBS_DIR / f"{job_id}.xlsx"
    if not file_path.exists():
        return jsonify({"error": "File not ready"}), 404
    name = JOBS.get(job_id, {}).get("result", {}).get("financials", {}).get("propertyInfo", {}).get("name", "Property")
    safe = re.sub(r"[^A-Za-z0-9_-]", "_", name)[:40] or "Property"
    return send_file(file_path, as_attachment=True, download_name=f"GGC_UW_{safe}.xlsx")


if __name__ == "__main__":
    print(" ╔═════════════════════════════════════════════════════════════════════════════════════╗")
    print(" ║  GGC Deal Engine — Backend Server v5                                                ║")
    print(f"║  Models: Financial={MODEL_FINANCIAL}, Market={MODEL_MARKET}                         ║")
    print(f"║  Template: GGC_Blank_Underwriting_Sizer_Extended (1000 rows)                        ║")
    print(f"║  Google Maps: {'ENABLED' if GOOGLE_MAPS_API_KEY else 'DISABLED (no key set)':<43s}  ║")
    print(" ║  Open: http://localhost:5001                                                        ║")
    print(" ╚═════════════════════════════════════════════════════════════════════════════════════╝")
    app.run(host="0.0.0.0", port=5001, debug=False, threaded=True)
