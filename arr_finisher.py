#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import subprocess
import re
import shutil
import time
import json
import logging
import unicodedata
import ctypes
from datetime import datetime
from logging.handlers import RotatingFileHandler
from urllib.parse import quote
from contextlib import contextmanager

import requests

__version__ = "1.0.0"

# Try optional pywin32 for .lnk creation; degrade gracefully if missing
try:
    import win32com.client  # type: ignore
    import pythoncom        # type: ignore
    HAS_WIN32COM = True
except Exception:
    HAS_WIN32COM = False

# Raised when a rating provider has a transient outage (5xx, timeout, network
# error). The caller should keep the folder's existing rating rather than
# falling back to a different provider — otherwise an anime that's normally
# rated by MAL gets re-rated as IMDb while Jikan is down, then incorrectly
# renamed. Returning None silently from get_mal_rating / get_mdl_rating used
# to cause exactly that bug (13 anime folders rebadged during a real outage).
class ProviderUnavailable(Exception):
    pass

# ==========================
# .env loader (no external dependency)
# ==========================
def _load_env_file(path):
    """Load KEY=VALUE lines from a file into os.environ (without overriding existing vars)."""
    try:
        with open(path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#') or '=' not in line:
                    continue
                key, _, val = line.partition('=')
                key = key.strip()
                val = val.strip()
                # Strip ONLY a matching pair of outer quotes — naive .strip('"')
                # would mangle values like KEY="a'b" (becomes a'b after strip).
                if len(val) >= 2 and val[0] == val[-1] and val[0] in ("'", '"'):
                    val = val[1:-1]
                # Real env vars win over .env file
                if key and key not in os.environ:
                    os.environ[key] = val
    except FileNotFoundError:
        pass
    except Exception as e:
        # Can't log yet — logging isn't configured. Print to stderr.
        sys.stderr.write(f"[arr_finisher] Failed to load {path}: {e}\n")

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_load_env_file(os.path.join(SCRIPT_DIR, ".env"))

def _require_env(name):
    """Return env var or empty string if unset."""
    return os.environ.get(name, "")

# ==========================
# Feature Toggles
# ==========================
ENABLE_RENAME_FOLDER        = True
ENABLE_UPDATE_SERVICE_PATH  = True
ENABLE_CREATE_FOLDER_ICON   = True
ENABLE_CREATE_SHORTCUTS     = True
ENABLE_SET_TOOLTIP          = True
ENABLE_GROUPED_LINKS_FOLDER = True
ENABLE_MDL_RATING           = True   # For Korean titles: use MDL rating (falls back to IMDb)
ENABLE_MAL_RATING           = True   # For anime: use MAL rating (falls back to IMDb)
FORCE_REGENERATE_SHORTCUTS  = False  # When True, delete + recreate all shortcuts on every run
DRY_RUN                     = False  # When True, log intended actions but don't touch disk/APIs
RATING_ONLY                 = False  # Sweep mode: refresh rating only, skip icon/shortcuts/tooltip

# Sweep cache TTL (days). Overridable via ARR_FINISHER_RATING_CACHE_TTL_DAYS.
try:
    RATING_CACHE_TTL_DAYS = int(os.environ.get("ARR_FINISHER_RATING_CACHE_TTL_DAYS", "7"))
except (TypeError, ValueError):
    RATING_CACHE_TTL_DAYS = 7

# --- Shortcuts ---
ENABLE_SHORTCUT_IMDB          = True
ENABLE_SHORTCUT_PARENTAL      = True
ENABLE_SHORTCUT_LETTERBOXD    = True
ENABLE_SHORTCUT_TVTIME        = True
ENABLE_SHORTCUT_TWITTER       = True
ENABLE_SHORTCUT_SUBTITLES     = True   # Combined Subtitle.vbs (SubDL + Subsource + OpenSubtitles)
ENABLE_SHORTCUT_MYDRAMALIST   = True   # Auto-detected: only added for Korean titles
ENABLE_SHORTCUT_MYANIMELIST   = True   # Auto-detected: only added for anime

# ==========================
# Paths & Keys
# ==========================
ICONS_DIR = os.path.join(SCRIPT_DIR, "icons")
FOLDER_ICON_EXE = _require_env("FOLDER_ICON_EXE")

RADARR_API_URL = _require_env("RADARR_API_URL") or "http://localhost:7878"
RADARR_API_KEY = _require_env("RADARR_API_KEY")

SONARR_API_URL = _require_env("SONARR_API_URL") or "http://localhost:8989"
SONARR_API_KEY = _require_env("SONARR_API_KEY")

# OMDb is the fallback rating source (IMDb GraphQL is primary). Still
# required for the Explorer hover tooltip (plot summary).
OMDB_API_KEY = _require_env("OMDB_API_KEY")

# Kuryana = unofficial MyDramaList API. Two mirrors for failover.
KURYANA_BASE_URLS = [
    _require_env("KURYANA_BASE_URL") or "https://kuryana.tbdh.app",
    "https://kuryana.vercel.app",
]

# Jikan = unofficial MyAnimeList API.
JIKAN_BASE_URL = _require_env("JIKAN_BASE_URL") or "https://api.jikan.moe/v4"

# Primary search language (ISO 639-1) — used by Twitter hashtag search and
# OpenSubtitles listings. Falls back to "ar" if unset for backward compat.
SEARCH_LANGUAGE = (_require_env("SEARCH_LANGUAGE") or "ar").strip().lower()

# --- SubDL Configuration ---
SUBDL_API_KEY = _require_env("SUBDL_API_KEY")

# --- OpenSubtitles Configuration ---
# API Key: https://www.opensubtitles.com/en/consumers
OPENSUBTITLES_API_KEY = _require_env("OPENSUBTITLES_API_KEY")

# ==========================
# Logging
# ==========================
LOGGER_NAME = "arr_finisher"

def _setup_logging():
    # Named logger (not root) so requests/urllib3 chatter doesn't get captured.
    # RotatingFileHandler is not safe across concurrent processes — concurrent
    # webhook invocations may lose a handful of records during rotation. Fine
    # for this use case (we mostly look at the latest run).
    logger = logging.getLogger(LOGGER_NAME)
    logger.propagate = False
    # Allow `ARR_FINISHER_LOG_LEVEL=DEBUG` for permanent verbose, else INFO default.
    level_name = os.environ.get("ARR_FINISHER_LOG_LEVEL", "INFO").upper()
    logger.setLevel(getattr(logging, level_name, logging.INFO))
    fmt = logging.Formatter('[%(asctime)s] %(levelname)s %(message)s', '%Y-%m-%d %H:%M:%S')

    try:
        log_dir = os.environ.get("ARR_FINISHER_LOG_DIR") or SCRIPT_DIR
        os.makedirs(log_dir, exist_ok=True)
        fh = RotatingFileHandler(os.path.join(log_dir, "arr_finisher.log"),
                                 maxBytes=1_000_000, backupCount=3, encoding='utf-8')
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except Exception:
        try:
            temp_dir = os.environ.get("TEMP") or os.environ.get("TMP") or "/tmp"
            fh = RotatingFileHandler(os.path.join(temp_dir, "arr_finisher.log"),
                                     maxBytes=1_000_000, backupCount=3, encoding='utf-8')
            fh.setFormatter(fmt)
            logger.addHandler(fh)
        except Exception:
            pass

    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(fmt)
    logger.addHandler(sh)

def _excepthook(exc_type, exc, tb):
    try:
        logging.getLogger(LOGGER_NAME).exception("Uncaught exception", exc_info=(exc_type, exc, tb))
    finally:
        import traceback
        traceback.print_exception(exc_type, exc, tb, file=sys.stderr)
sys.excepthook = _excepthook

_setup_logging()
log = logging.getLogger(LOGGER_NAME).info
log_err = logging.getLogger(LOGGER_NAME).error
log_warn = logging.getLogger(LOGGER_NAME).warning
log_debug = logging.getLogger(LOGGER_NAME).debug

# ==========================
# Per-series/movie lock (cross-process, with stale-lock reclaim)
# ==========================
_FS_LOCK_STALE_SECS = 600  # 10 minutes — a real run never approaches this

@contextmanager
def _fs_lock(key: str, retries: int = 6, wait_s: float = 0.5):
    """Best-effort cross-process lock. Retries on contention; reclaims locks
    older than _FS_LOCK_STALE_SECS (presumed crashed). If still unavailable,
    logs a warning and proceeds — better than blocking a webhook forever."""
    base = os.environ.get("TEMP") or os.environ.get("TMP") or "/tmp"
    safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", key or "lock")
    path = os.path.join(base, f"arr_finisher_lock_{safe}")
    acquired = False
    for _ in range(retries):
        try:
            os.makedirs(path, exist_ok=False)
            acquired = True
            break
        except FileExistsError:
            try:
                age = time.time() - os.path.getmtime(path)
            except OSError:
                age = 0
            if age > _FS_LOCK_STALE_SECS:
                log_warn(f"Reclaiming stale lock (age {age:.0f}s): {path}")
                shutil.rmtree(path, ignore_errors=True)
                continue
            time.sleep(wait_s)
        except Exception:
            break
    if not acquired:
        log_warn(f"Could not acquire lock {key!r} after {retries} attempts; proceeding")
    try:
        yield acquired
    finally:
        if acquired:
            try:
                shutil.rmtree(path, ignore_errors=True)
            except Exception:
                pass

# ==========================
# Windows file attributes (ctypes; replaces subprocess attrib)
# ==========================
_FILE_ATTRIBUTE_READONLY = 0x01
_FILE_ATTRIBUTE_HIDDEN   = 0x02
_FILE_ATTRIBUTE_SYSTEM   = 0x04
_INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF

def _set_file_attrs(path, add=0, remove=0):
    """Best-effort: add/remove Windows file-attribute flags. Returns True on success."""
    if os.name != "nt":
        return False
    try:
        kernel32 = ctypes.windll.kernel32
        kernel32.GetFileAttributesW.argtypes = [ctypes.c_wchar_p]
        kernel32.GetFileAttributesW.restype = ctypes.c_uint32
        kernel32.SetFileAttributesW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32]
        kernel32.SetFileAttributesW.restype = ctypes.c_int
        current = kernel32.GetFileAttributesW(str(path))
        if current == _INVALID_FILE_ATTRIBUTES:
            return False
        new = (current | add) & (~remove & 0xFFFFFFFF)
        if new == current:
            return True
        return bool(kernel32.SetFileAttributesW(str(path), new))
    except Exception:
        return False

# ==========================
# URL safety (for content interpolated into VBS / shortcut targets)
# ==========================
def _is_safe_url(url):
    """Reject anything that would break VBS double-quoted strings or shell args.
    URLs from APIs should be plain ASCII https — anything else is suspicious."""
    if not url or not isinstance(url, str):
        return False
    if not url.startswith(("http://", "https://")):
        return False
    if any(c in url for c in ('"', '\r', '\n', '\0')):
        return False
    if len(url) > 2048:
        return False
    return True

# ==========================
# Log redaction
# ==========================
def _redact(s):
    """Return s with known secrets replaced by '<redacted>'."""
    try:
        text = str(s)
    except Exception:
        return s
    for secret in (OMDB_API_KEY, SUBDL_API_KEY, OPENSUBTITLES_API_KEY,
                   SONARR_API_KEY, RADARR_API_KEY):
        if secret and len(secret) > 3:
            text = text.replace(secret, "<redacted>")
    return text

# ==========================
# Shared HTTP session
# ==========================
_session = None
def http():
    global _session
    if _session is not None:
        return _session
    from requests.adapters import HTTPAdapter
    try:
        from urllib3.util import Retry
    except Exception:
        from urllib3.util.retry import Retry

    retry = Retry(
        total=3,
        backoff_factor=0.6,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=frozenset(["GET","POST","PUT","PATCH","DELETE","HEAD","OPTIONS"]),
        raise_on_status=False,
    )
    s = requests.Session()
    s.headers.update({"User-Agent": "arr-finisher/1.0 (Compatible; Python)"})
    s.mount("http://", HTTPAdapter(max_retries=retry))
    s.mount("https://", HTTPAdapter(max_retries=retry))
    _session = s
    return _session

# ==========================
# Slugify Utility
# ==========================
def slugify(value):
    if not value: return "unknown"
    try:
        value = str(value)
        value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode('ascii')
        value = re.sub(r'[^\w\s-]', '', value).strip().lower()
        value = re.sub(r'[-\s]+', '-', value)
        return value
    except Exception:
        return "unknown"

# ==========================
# OpenSubtitles API Logic (URL Resolver)
# ==========================
def get_opensubtitles_web_url(imdb_id, content_type="movie"):
    """
    Uses the API to find the correct OpenSubtitles web page.
    Fallback: Standard search URL.
    """
    default_search = f"https://www.opensubtitles.com/{SEARCH_LANGUAGE}/{SEARCH_LANGUAGE}/search-all/q-{imdb_id}/hearing_impaired-exclude/machine_translated-/trusted_sources-"
    
    if not OPENSUBTITLES_API_KEY:
        return default_search

    # Clean IMDb ID (remove 'tt' and leading zeros)
    clean_id = re.sub(r"\D", "", str(imdb_id))
    try:
        clean_id = int(clean_id)
    except ValueError:
        return default_search

    url = "https://api.opensubtitles.com/api/v1/features"
    headers = {
        "Api-Key": OPENSUBTITLES_API_KEY,
        "Content-Type": "application/json",
        "User-Agent": "arr-finisher/1.0"
    }
    params = {"imdb_id": clean_id}

    try:
        r = http().get(url, params=params, headers=headers, timeout=20)
        r.raise_for_status()
        data = r.json()
        
        # The API returns a list of features matching the ID
        results = data.get("data", [])
        if not results:
            log(f"OpenSubtitles API: No features found for {imdb_id}")
            return default_search

        # Pick the first match
        # Structure: attributes -> url (often provides the direct slug link)
        # OR attributes -> title, year, feature_id
        
        item = results[0]
        attr = item.get("attributes", {})
        
        # OpenSubtitles often returns a direct URL in attributes
        if attr.get("url"):
            return attr.get("url")

        # Fallback: Construct URL from slug if available
        # Web format: https://www.opensubtitles.com/en/movies/{year}-{slug}
        # Web format: https://www.opensubtitles.com/en/tvshows/{year}-{slug}
        
        title = attr.get("title")
        year = attr.get("year")
        slug = slugify(title)
        
        # Determine type prefix
        # API type: 'Movie' or 'Tvshow' usually
        ft_type = attr.get("feature_type", "").lower()
        
        if "movie" in ft_type or content_type == "movie":
            base_section = "movies"
        else:
            base_section = "tvshows"
            
        if year and slug:
            return f"https://www.opensubtitles.com/en/{base_section}/{year}-{slug}"
            
        # Last resort fallback to search if we can't construct a pretty URL
        return default_search

    except Exception as e:
        log_err(f"OpenSubtitles API URL lookup failed: {e}")
        return default_search


# ==========================
# SubDL API Logic
# ==========================
def get_subdl_web_url(imdb_id, content_type="movie"):
    default_search = f"https://subdl.com/search/{imdb_id}"
    if not SUBDL_API_KEY: return default_search

    params = { "api_key": SUBDL_API_KEY, "imdb_id": imdb_id, "type": content_type }
    urls_to_try = ["https://api.subdl.com/api/v1/subtitles", "https://subdl.com/api/v1/subtitles"]
    
    data = None
    for url in urls_to_try:
        try:
            r = http().get(url, params=params, timeout=30)
            json_data = r.json()
            if json_data.get("status"):
                data = json_data
                break
        except Exception:
            continue

    if not data: return default_search

    try:
        sd_id = None
        name = None
        def clean_sd_id(raw_id):
            if raw_id:
                s = str(raw_id).strip()
                if s.isdigit(): return f"sd{s}"
                if s.lower().startswith("sd"): return s
            return None

        for key in ["movie", "show", "tv", "results"]:
            obj = data.get(key)
            if isinstance(obj, list) and obj: obj = obj[0]
            if isinstance(obj, dict):
                found_id = clean_sd_id(obj.get("sd_id"))
                if found_id:
                    sd_id = found_id
                    name = obj.get("name") or obj.get("title") or obj.get("original_name")
                    break

        if not sd_id and data.get("subtitles"):
            first = data["subtitles"][0]
            sd_id = clean_sd_id(first.get("sd_id") or first.get("movie_id"))
            if not name: name = first.get("release_name", "unknown").split(".")[0] 

        if sd_id:
            final_slug = slugify(name) if name else "unknown"
            return f"https://subdl.com/subtitle/{sd_id}/{final_slug}"
    except Exception as e:
        log_err(f"SubDL URL lookup failed: {_redact(e)}")

    return default_search


# ==========================
# IMDb Rating
# ==========================
# Per-process cache of full OMDb responses, keyed by imdb_id. OMDb returns
# rating + plot in one call, but the two callers ask separately — cache so we
# don't double-fetch (cuts a sweep's OMDb traffic in half for normal libraries).
_omdb_response_cache = {}

def _fetch_omdb(imdb_id):
    """Return the full OMDb JSON for imdb_id, cached per process."""
    if not imdb_id or not OMDB_API_KEY:
        return {}
    if imdb_id in _omdb_response_cache:
        return _omdb_response_cache[imdb_id]
    try:
        url = f"https://www.omdbapi.com/?apikey={OMDB_API_KEY}&i={quote(imdb_id)}&plot=short&r=json"
        r = http().get(url, timeout=15)
        if r.status_code != 200:
            _omdb_response_cache[imdb_id] = {}
            return {}
        data = r.json() or {}
    except Exception as e:
        log_debug(f"OMDb fetch failed: {_redact(e)}")
        data = {}
    _omdb_response_cache[imdb_id] = data
    return data

def get_omdb_plot(imdb_id: str) -> str:
    """Return the short plot summary from OMDb, or empty string."""
    plot = (_fetch_omdb(imdb_id) or {}).get("Plot") or ""
    return "" if plot == "N/A" else plot.strip()

def set_folder_tooltip(folder_path: str, tooltip: str) -> None:
    """Add/replace InfoTip in folder's desktop.ini so Explorer shows a hover tooltip."""
    if not ENABLE_SET_TOOLTIP:
        return
    if not tooltip or not os.path.isdir(folder_path):
        return
    ini_path = os.path.join(folder_path, "desktop.ini")
    if DRY_RUN:
        log(f"[DRY RUN] Would set InfoTip on {folder_path!r}")
        return

    # Read any existing desktop.ini (try common encodings).
    # NB: cp1252 accepts almost any byte sequence, so do a sanity check before
    # accepting a candidate decoding — must contain '=' or a section header.
    lines = []
    if os.path.exists(ini_path):
        try:
            with open(ini_path, "rb") as fh:
                raw = fh.read()
        except OSError:
            raw = b""
        for enc in ("utf-16", "utf-8-sig", "cp1252"):
            try:
                decoded = raw.decode(enc)
            except UnicodeDecodeError:
                continue
            if "=" in decoded or "[" in decoded or not raw:
                lines = decoded.splitlines()
                break

    # Idempotency check: if the existing InfoTip already matches, skip the write.
    existing_tip = None
    in_shell = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            in_shell = stripped.lower() == "[.shellclassinfo]"
        elif in_shell and stripped.lower().startswith("infotip="):
            existing_tip = stripped[len("infotip="):]
            break
    if existing_tip == tooltip:
        log_debug(f"Tooltip unchanged on {os.path.basename(folder_path)}")
        return

    # Rebuild lines: ensure [.ShellClassInfo] section and a single InfoTip
    result = []
    in_shell = False
    has_shell = False
    replaced = False
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("[") and stripped.endswith("]"):
            # Leaving previous section; insert InfoTip before the next section if we
            # were in ShellClassInfo and haven't replaced yet.
            if in_shell and not replaced:
                result.append(f"InfoTip={tooltip}")
                replaced = True
            in_shell = stripped.lower() == "[.shellclassinfo]"
            if in_shell:
                has_shell = True
        elif in_shell and stripped.lower().startswith("infotip="):
            result.append(f"InfoTip={tooltip}")
            replaced = True
            continue
        result.append(line)
    # End of file — if we're still inside ShellClassInfo, append
    if in_shell and not replaced:
        result.append(f"InfoTip={tooltip}")
        replaced = True
    if not has_shell:
        result.append("[.ShellClassInfo]")
        result.append(f"InfoTip={tooltip}")
        has_shell = True

    # Atomic write: temp file -> rename. desktop.ini may already be system+hidden,
    # which os.replace handles fine on NTFS even if dest has those attributes set.
    try:
        _set_file_attrs(ini_path, remove=(_FILE_ATTRIBUTE_READONLY
                                          | _FILE_ATTRIBUTE_HIDDEN
                                          | _FILE_ATTRIBUTE_SYSTEM))
        tmp_path = ini_path + ".tmp"
        # UTF-16 with BOM for max Explorer compatibility on non-ASCII tooltips.
        with open(tmp_path, "w", encoding="utf-16") as f:
            f.write("\r\n".join(result) + "\r\n")
        os.replace(tmp_path, ini_path)
        _set_file_attrs(ini_path, add=(_FILE_ATTRIBUTE_SYSTEM | _FILE_ATTRIBUTE_HIDDEN))
        # Folder itself needs System attr for desktop.ini to be consulted.
        _set_file_attrs(folder_path, add=_FILE_ATTRIBUTE_SYSTEM)
        log(f"Tooltip set on {os.path.basename(folder_path)}")
    except Exception as e:
        log_err(f"Failed to set tooltip on {folder_path}: {e}")
        try:
            if os.path.exists(ini_path + ".tmp"):
                os.remove(ini_path + ".tmp")
        except OSError:
            pass

def get_imdb_rating_from_omdb(imdb_id: str) -> str:
    val = (_fetch_omdb(imdb_id) or {}).get("imdbRating") or ""
    if val and val != "N/A":
        m = re.match(r"^(\d+(?:\.\d)?)", str(val))
        return m.group(1) if m else str(val)
    return "N/A"

# IMDb's public GraphQL endpoint — returns live ratings (not OMDb-cached values,
# which lag by days on recent titles). No API key required; used internally by
# imdb.com itself, so it's stable. Non-commercial use is permitted per IMDb's
# data-usage policy (the response includes a disclaimer to that effect).
# Replaces the previous HTML JSON-LD scrape, which stopped working when IMDb
# put the site behind AWS WAF (requests get HTTP 202 + empty body without JS).
# Overridable via IMDB_GRAPHQL_URL env var (matches the pattern of
# KURYANA_BASE_URL / JIKAN_BASE_URL — useful if IMDb moves the endpoint).
_IMDB_GRAPHQL_URL = _require_env("IMDB_GRAPHQL_URL") or "https://caching.graphql.imdb.com/"
_IMDB_RATING_QUERY = (
    "query R($id: ID!) { title(id: $id) { "
    "ratingsSummary { aggregateRating } } }"
)

def get_imdb_rating_from_graphql(imdb_id: str) -> str:
    if not imdb_id:
        return "N/A"
    try:
        r = http().post(
            _IMDB_GRAPHQL_URL,
            headers={
                "Content-Type": "application/json",
                "Accept": "application/graphql+json, application/json",
            },
            json={"query": _IMDB_RATING_QUERY, "variables": {"id": imdb_id}},
            timeout=15,
        )
        if r.status_code != 200:
            return "N/A"
        data = (r.json() or {}).get("data") or {}
        rs = ((data.get("title") or {}).get("ratingsSummary") or {})
        val = rs.get("aggregateRating")
        if val is None:
            return "N/A"
        return f"{float(val):.1f}"
    except (requests.RequestException, ValueError, TypeError) as e:
        log_err(f"IMDb GraphQL request failed: {_redact(e)}")
        return "N/A"

def _title_similarity(a, b):
    """Normalized-edit-distance similarity in [0, 1]."""
    from difflib import SequenceMatcher
    return SequenceMatcher(None, (a or "").lower(), (b or "").lower()).ratio()

# When get_mdl_rating / get_mal_rating finds a match, it caches the direct URL
# here so create_shortcuts can link straight to the show's page instead of a
# search URL. Keyed by (title-without-year, year).
_mdl_url_cache = {}
_mal_url_cache = {}

def _provider_cache_key(title, year):
    clean = re.sub(r'\s*\(\d{4}\)\s*$', '', (title or '')).strip().lower()
    return (clean, str(year or ''))

# ==========================
# Rating freshness cache (sweep mode only)
# ==========================
RATING_CACHE_PATH = os.path.join(SCRIPT_DIR, ".rating_cache.json")
_rating_cache = None    # lazy-loaded
_rating_cache_dirty = False  # flips True on _rating_cache_set

def _load_rating_cache():
    global _rating_cache
    if _rating_cache is not None:
        return _rating_cache
    try:
        with open(RATING_CACHE_PATH, "r", encoding="utf-8") as f:
            _rating_cache = json.load(f) or {}
    except FileNotFoundError:
        _rating_cache = {}
    except Exception as e:
        log_err(f"Could not load rating cache: {e}; starting empty")
        _rating_cache = {}
    return _rating_cache

def _save_rating_cache():
    """No-op if the cache hasn't been modified since last save (or load)."""
    global _rating_cache_dirty
    if _rating_cache is None or not _rating_cache_dirty:
        return
    if DRY_RUN:
        return
    tmp = RATING_CACHE_PATH + ".tmp"
    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(_rating_cache, f, indent=2, sort_keys=True)
        os.replace(tmp, RATING_CACHE_PATH)
        _rating_cache_dirty = False
    except Exception as e:
        log_err(f"Could not save rating cache: {e}")
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass

def _parse_checked_at(value):
    """Accept either an ISO-8601 string (new) or a float epoch (legacy).
    Returns float epoch, or None if unparseable."""
    if value is None:
        return None
    try:
        if isinstance(value, str):
            return datetime.fromisoformat(value).timestamp()
        return float(value)
    except (ValueError, TypeError):
        return None

def _rating_cache_is_fresh(imdb_id):
    """True if the cached rating for imdb_id is younger than RATING_CACHE_TTL_DAYS."""
    if not imdb_id:
        return False
    cache = _load_rating_cache()
    entry = cache.get(imdb_id)
    if not entry:
        return False
    ts = _parse_checked_at(entry.get("checked_at"))
    if ts is None:
        return False
    age_days = (time.time() - ts) / 86400.0
    return age_days < RATING_CACHE_TTL_DAYS

def _rating_cache_set(imdb_id, rating, source):
    global _rating_cache_dirty
    if not imdb_id:
        return
    cache = _load_rating_cache()
    cache[imdb_id] = {
        # ISO 8601 (local time, second precision) — `cat .rating_cache.json`
        # is now readable. Legacy float-epoch entries are still understood.
        "checked_at": datetime.now().isoformat(timespec="seconds"),
        "rating": rating,
        "source": source,
    }
    _rating_cache_dirty = True

def get_mdl_rating(title: str, year=None):
    """
    Query kuryana (unofficial MDL API) and return (rating_str, 'MDL') or None.
    Requires a confident match: year must match OR title similarity >= 0.85.

    Raises ProviderUnavailable if ALL kuryana mirrors fail with transient
    errors (5xx, 429, timeout, network) — so the caller can keep the existing
    rating instead of falling back to IMDb. Returns None when at least one
    mirror responded but no confident match was found.
    """
    if not title:
        return None
    clean = re.sub(r'\s*\(\d{4}\)\s*$', '', title).strip()
    if not clean:
        return None
    query = quote(clean)

    all_unavailable = True  # flips to False if any mirror responds (200 or 4xx)
    for base in KURYANA_BASE_URLS:
        try:
            r = http().get(f"{base}/search/q/{query}", timeout=15)
            if r.status_code in (429, 500, 502, 503, 504):
                continue  # try next mirror, keep all_unavailable=True
            all_unavailable = False
            if r.status_code != 200:
                continue
            data = r.json() or {}
            dramas = (data.get("results") or {}).get("dramas") or []
            if not dramas:
                continue

            # Prefer Korean type (Drama or Movie)
            korean = [d for d in dramas if "korean" in (d.get("type") or "").lower()]
            candidates = korean or dramas

            # Score each candidate by (year match, title similarity)
            def score(d):
                yr_ok = bool(year) and str(d.get("year") or "") == str(year)
                sim = _title_similarity(d.get("title") or "", clean)
                return (yr_ok, sim)

            candidates.sort(key=score, reverse=True)
            best = candidates[0]

            # Reject if no year match AND title similarity is too low
            yr_ok, sim = score(best)
            if not yr_ok and sim < 0.85:
                log(f"MDL match rejected for {title!r}: best='{best.get('title')}' ({best.get('year')}), similarity={sim:.2f}")
                return None

            rating = best.get("rating")
            try:
                val = float(rating)
            except (TypeError, ValueError):
                continue
            if val <= 0:
                continue

            # Cache the direct MDL URL so the shortcut can link to the show page
            slug = best.get("slug")
            if slug:
                _mdl_url_cache[_provider_cache_key(title, year)] = f"https://mydramalist.com/{slug}"

            return (f"{val:.1f}", "MDL")
        except Exception as e:
            # Anything that prevents us from getting a rating is treated as a
            # transient outage (network error, malformed JSON, unexpected shape
            # raising TypeError/AttributeError). Better to keep the existing
            # rating than to silently fall back to IMDb on weird input.
            log_err(f"Kuryana lookup at {base} failed: {e}")
            continue  # keep all_unavailable=True so we raise at the end

    if all_unavailable:
        raise ProviderUnavailable("All kuryana mirrors unavailable")
    return None

def _normalize_for_match(s):
    """Lowercase and strip any trailing (YYYY) so title comparisons are fair."""
    return re.sub(r'\s*\(\d{4}\)\s*$', '', (s or '')).strip().lower()

def _all_candidate_titles(item):
    """Collect every title variant jikan knows about for an anime entry."""
    out = []
    for key in ("title", "title_english", "title_japanese"):
        v = item.get(key)
        if v: out.append(v)
    out.extend(item.get("title_synonyms") or [])
    for t in (item.get("titles") or []):
        v = t.get("title") if isinstance(t, dict) else None
        if v: out.append(v)
    return out

def get_mal_rating(title, year=None):
    """
    Query jikan (unofficial MAL API) and return (rating_str, 'MAL') or None.
    Matches against every known alternative title (English, Japanese, synonyms,
    titles[] array). Requires year match OR title similarity >= 0.85.

    Raises ProviderUnavailable on transient errors (5xx, 429, timeout, network)
    so the caller can preserve the existing rating instead of falling back to
    IMDb. Returns None only when jikan responded successfully but no confident
    match was found.
    """
    if not title:
        return None
    clean = re.sub(r'\s*\(\d{4}\)\s*$', '', title).strip()
    if not clean:
        return None
    clean_norm = clean.lower()
    try:
        # NOTE: do NOT pass order_by=score — it overrides jikan's relevance
        # ranking and returns arbitrary high-rated shows.
        params = {"q": clean, "limit": 10}
        r = http().get(f"{JIKAN_BASE_URL}/anime", params=params, timeout=15)
        if r.status_code in (429, 500, 502, 503, 504):
            raise ProviderUnavailable(f"Jikan returned {r.status_code}")
        if r.status_code != 200:
            return None
        data = r.json() or {}
        results = data.get("data") or []
        if not results:
            return None

        def score(item):
            # Extract year from top-level `year` or `aired.from`
            yr = None
            for key in ("year", "aired"):
                v = item.get(key)
                if isinstance(v, int):
                    yr = v
                elif isinstance(v, dict):
                    try:
                        yr = int((v.get("from") or "")[:4])
                    except (ValueError, TypeError):
                        yr = None
                if yr: break
            yr_ok = bool(year) and yr is not None and abs(int(year) - yr) <= 1
            # Compare against every alternative title, strip year from each
            sim = max(
                (_title_similarity(_normalize_for_match(t), clean_norm)
                 for t in _all_candidate_titles(item)),
                default=0,
            )
            return (yr_ok, sim)

        results.sort(key=score, reverse=True)
        best = results[0]
        yr_ok, sim = score(best)
        if not yr_ok and sim < 0.85:
            log(f"MAL match rejected for {title!r}: best='{best.get('title')}', similarity={sim:.2f}")
            return None

        rating = best.get("score")
        try:
            val = float(rating)
        except (TypeError, ValueError):
            return None
        if val <= 0:
            return None

        # Cache the direct MAL URL (from jikan) or build it from mal_id
        mal_url = best.get("url")
        if not mal_url and best.get("mal_id"):
            mal_url = f"https://myanimelist.net/anime/{best['mal_id']}"
        if mal_url:
            _mal_url_cache[_provider_cache_key(title, year)] = mal_url

        return (f"{val:.1f}", "MAL")
    except ProviderUnavailable:
        raise
    except Exception as e:
        # Anything else (network error, malformed JSON, unexpected shape
        # raising TypeError/AttributeError) → treat as transient. The point of
        # ProviderUnavailable is to fully contain provider weirdness so we
        # never silently fall back to IMDb on bad input from Jikan.
        raise ProviderUnavailable(f"Jikan lookup failed: {e}") from e

def get_rating_for_title(imdb_id, title, year=None, is_korean=False, is_anime=False):
    """
    Return (rating, source) — e.g. ("7.5", "MDL"), ("8.9", "MAL"),
    ("8.6", "IMDb"), or ("N/A", "IMDb").
    Tries MAL for anime, MDL for Korean, falls back to IMDb.
    """
    if is_anime and ENABLE_MAL_RATING:
        result = get_mal_rating(title, year)
        if result:
            return result
    if is_korean and ENABLE_MDL_RATING:
        result = get_mdl_rating(title, year)
        if result:
            return result
    r = get_imdb_rating(imdb_id)
    return (r, "IMDb")

def get_imdb_rating(imdb_id: str) -> str:
    """Return current IMDb rating as a string, or 'N/A'.

    Tries IMDb's GraphQL endpoint first (live ratings, no API key, no WAF
    challenge) and falls back to OMDb only if GraphQL returns nothing. OMDb's
    cached values can lag the live IMDb rating by 0–0.2 stars on recent
    titles, so GraphQL is preferred when reachable.
    """
    if not imdb_id:
        return "N/A"
    rating = get_imdb_rating_from_graphql(imdb_id)
    if rating != "N/A":
        return rating
    # GraphQL miss → OMDb. Logged at INFO so a sustained pattern (e.g. IMDb
    # broke their endpoint) is visible in the rotating log without needing
    # --verbose. Logged once per call site, which is fine for a sweep.
    log("IMDb GraphQL returned no rating; falling back to OMDb")
    return get_imdb_rating_from_omdb(imdb_id)

# ==========================
# Rename Folder
# ==========================
_RATING_SUFFIX_RE = re.compile(r'\s+\[(?:IMDb|MDL|MAL|TMDb|RT)\s*\d+(?:\.\d+)?\]$')

def _strip_rating_suffix(name: str) -> str:
    """Strip any rating suffix like [IMDb 7.2], [MDL 9.1], [MAL 8.4]."""
    return _RATING_SUFFIX_RE.sub('', name or '')

def _has_rating_suffix(name: str) -> bool:
    return bool(_RATING_SUFFIX_RE.search(name or ''))

def rename_folder(old_path, rating, source="IMDb"):
    old_path = os.path.normpath((old_path or '').rstrip(r'\/'))
    base_name = _strip_rating_suffix(os.path.basename(old_path))
    new_name = f"{base_name} [{source} {rating}]"
    new_path = os.path.join(os.path.dirname(old_path), new_name)

    if os.path.abspath(old_path) == os.path.abspath(new_path):
        log_debug("Rename not needed (same name).")
        return old_path
    if not os.path.exists(old_path):
        log_err(f"Rename skipped: source path not found ({old_path})")
        return new_path if os.path.exists(new_path) else old_path

    if os.path.exists(new_path):
        log(f"Destination exists ({new_path}). Attempting merge...")
        try:
            skipped = []
            for item in os.listdir(old_path):
                s = os.path.join(old_path, item)
                d = os.path.join(new_path, item)
                if os.path.exists(d):
                    log_warn(f"File exists in destination, leaving in source: {item}")
                    skipped.append(item)
                    continue
                shutil.move(s, d)
            if skipped:
                # Don't rmtree — would silently destroy the files we just refused
                # to overwrite. Leave the source folder for the user to resolve.
                log_err(
                    f"Merge incomplete: {len(skipped)} item(s) remain in {old_path} "
                    f"because the same name exists in {new_path}. Resolve manually."
                )
                return new_path
            try:
                shutil.rmtree(old_path)
                log(f"Merge complete. Removed source: {old_path}")
                return new_path
            except OSError as e:
                log_err(f"Could not remove source folder: {e}")
                return new_path
        except Exception as e:
            log_err(f"Merge failed: {e}")
            return old_path

    if DRY_RUN:
        log(f"[DRY RUN] Would rename {old_path} -> {new_path}")
        return new_path

    # Use os.rename (atomic on same-drive NTFS) instead of shutil.move.
    # shutil.move falls back to copy+delete when the rename fails, which
    # leaves partial duplicates behind when a child file is locked.
    for attempt in range(1, 6):
        try:
            os.rename(old_path, new_path)
            log(f"Renamed {old_path} -> {new_path}")
            return new_path
        except OSError as e:
            log_err(f"Rename attempt {attempt} failed: {e}")
            time.sleep(1.0)

    log_err(f"Rename failed after retries ({old_path} -> {new_path}).")
    return old_path

_ROLLBACK_MARKER_PATH = os.path.join(SCRIPT_DIR, ".rollbacks.log")

def _append_rollback_marker(line):
    """Append a one-line record to .rollbacks.log so users can grep for issues."""
    if DRY_RUN:
        return
    try:
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(_ROLLBACK_MARKER_PATH, "a", encoding="utf-8") as fh:
            fh.write(f"[{ts}] {line}\n")
    except OSError:
        pass

def rollback_rename(new_path, old_path):
    """Best-effort rename from new_path back to old_path. Logs on failure."""
    if DRY_RUN:
        log(f"[DRY RUN] Would roll back: {new_path} -> {old_path}")
        return True
    try:
        os.rename(new_path, old_path)
        log(f"Rolled back disk rename: {new_path} -> {old_path}")
        _append_rollback_marker(f"OK   {new_path} -> {old_path} (API refused rename)")
        return True
    except OSError as e:
        log_err(
            f"Rollback failed; disk is at {new_path} but service expects {old_path}. "
            f"Manual fix needed: {e}"
        )
        _append_rollback_marker(f"FAIL disk={new_path} service_expects={old_path} err={e}")
        return False

def _desktop_ini_has_icon(folder_path):
    """True if folder's desktop.ini already declares an IconResource= line.
    File existence alone isn't enough — set_folder_tooltip also writes
    desktop.ini, so a tooltip-only run shouldn't lock out future icon runs."""
    ini = os.path.join(folder_path, "desktop.ini")
    if not os.path.isfile(ini):
        return False
    try:
        with open(ini, "rb") as fh:
            raw = fh.read()
    except OSError:
        return False
    for enc in ("utf-16", "utf-8-sig", "cp1252"):
        try:
            text = raw.decode(enc)
        except UnicodeDecodeError:
            continue
        for line in text.splitlines():
            if line.strip().lower().startswith("iconresource="):
                return True
        return False
    return False

def create_folder_icon(path):
    if not ENABLE_CREATE_FOLDER_ICON:
        return
    try:
        if _desktop_ini_has_icon(path):
            log_debug("Folder icon already set; skipping")
            return
        if DRY_RUN:
            log(f"[DRY RUN] Would run FolderIconCreator on {path}")
            return
        subprocess.run([FOLDER_ICON_EXE, "-h", "-f", path], check=False)
        _set_file_attrs(path, add=(_FILE_ATTRIBUTE_SYSTEM | _FILE_ATTRIBUTE_READONLY))
        log("FolderIconCreator OK")
    except Exception as e:
        log_err(f"FolderIconCreator failed: {e}")

# ==========================
# .lnk helper
# ==========================
def _write_lnk(path, target, name):
    if os.path.exists(path):
        if FORCE_REGENERATE_SHORTCUTS:
            try:
                os.remove(path)
                log(f"Regenerating: {os.path.basename(path)}")
            except Exception as e:
                log_err(f"Could not remove existing {path} for regeneration: {e}")
                return
        else:
            log_debug(f"Already exists: {os.path.basename(path)}")
            return
    if not HAS_WIN32COM:
        log_err("win32com is not available; skipping .lnk creation.")
        return
    if DRY_RUN:
        log(f"[DRY RUN] Would create shortcut: {os.path.basename(path)} -> {target}")
        return
    # COM needs to be initialized per thread. Safe to call repeatedly —
    # returns S_FALSE when already initialized.
    com_initialized = False
    try:
        pythoncom.CoInitialize()
        com_initialized = True
    except Exception:
        pass
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(path)
        if str(target).lower().startswith("http"):
            shortcut.TargetPath = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "explorer.exe")
            shortcut.Arguments = f'"{target}"'
            shortcut.WorkingDirectory = os.path.dirname(path)
        else:
            shortcut.TargetPath = target
            shortcut.WorkingDirectory = os.path.dirname(target)
        icon_path = os.path.join(ICONS_DIR, f"{name}.ico")
        if os.path.exists(icon_path):
            shortcut.IconLocation = f"{icon_path},0"
        shortcut.Save()
        _set_file_attrs(path, remove=_FILE_ATTRIBUTE_HIDDEN)
        log(f"Shortcut: {os.path.basename(path)}")
    except Exception as e:
        log_err(f"Failed creating {name} shortcut: {e}")
    finally:
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

# ==========================
# Shortcuts creator
# ==========================
def create_shortcuts(service, folder_path, imdb_id, tmdb_or_tvdb_id, title, is_korean=False, is_anime=False, year=None):
    if not ENABLE_CREATE_SHORTCUTS:
        return

    links_dir = folder_path
    if ENABLE_GROUPED_LINKS_FOLDER:
        links_dir = os.path.join(folder_path, "Links")
        try:
            os.makedirs(links_dir, exist_ok=True)
        except Exception as e:
            log_err(f"Could not create Links folder: {e}")
            links_dir = folder_path

    # Force-regenerate: wipe existing .lnk and .vbs files so they're rebuilt
    # with current URL formats, icons, etc.
    if FORCE_REGENERATE_SHORTCUTS and os.path.isdir(links_dir):
        for fname in os.listdir(links_dir):
            if fname.lower().endswith(('.lnk', '.vbs')):
                try:
                    os.remove(os.path.join(links_dir, fname))
                except Exception as e:
                    log_err(f"Could not remove {fname} for regeneration: {e}")
        log("Regenerating all shortcuts in Links folder")

    def make_link(name, url_or_path):
        _write_lnk(os.path.join(links_dir, f"{name}.lnk"), url_or_path, name)

    if imdb_id and ENABLE_SHORTCUT_IMDB:
        make_link("IMDb", f"https://www.imdb.com/title/{imdb_id}/")

    # Parents guide: a single click opens IMDb parental guide + Common Sense
    # Media + Does The Dog Die in three tabs. IMDb is community-edited and
    # often sparse on newer titles; CSM has curated age/content scores; DTDD
    # covers triggers IMDb skips. Same pattern as the combined Subtitle shortcut.
    if imdb_id and ENABLE_SHORTCUT_PARENTAL:
        try:
            imdb_pg_url = f"https://www.imdb.com/title/{quote(imdb_id)}/parentalguide/"
            clean_title = re.sub(r'\s*\(\d{4}\)$', '', title or '').strip()
            title_q = quote(clean_title) if clean_title else ""
            csm_url = (f"https://www.commonsensemedia.org/search?query={title_q}"
                       if title_q else "https://www.commonsensemedia.org/")
            dtdd_url = (f"https://www.doesthedogdie.com/search?text={title_q}"
                        if title_q else "https://www.doesthedogdie.com/")
            # URLs go into a VBS literal — reject anything with quotes/newlines.
            if not _is_safe_url(imdb_pg_url):
                imdb_pg_url = "https://www.imdb.com/"
            if not _is_safe_url(csm_url):
                csm_url = "https://www.commonsensemedia.org/"
            if not _is_safe_url(dtdd_url):
                dtdd_url = "https://www.doesthedogdie.com/"

            vbs_path = os.path.join(links_dir, "Parents guide.vbs")
            lnk_path = os.path.join(links_dir, "Parents guide.lnk")

            # Migration: old version was a single .lnk pointing straight at the
            # IMDb parental-guide URL. When the .vbs is being written for the
            # first time, delete the stale .lnk so the next _write_lnk call
            # regenerates it pointing at the .vbs.
            if not os.path.exists(vbs_path) and os.path.exists(lnk_path):
                try: os.remove(lnk_path)
                except Exception: pass

            # Rewrite the .vbs if missing OR if URLs drifted (drift-aware,
            # same as Subtitle.vbs).
            content = (
                'Set sh = CreateObject("WScript.Shell")\n'
                f'sh.Run "explorer.exe ""{imdb_pg_url}""", 1, False\n'
                'WScript.Sleep 200\n'
                f'sh.Run "explorer.exe ""{csm_url}""", 1, False\n'
                'WScript.Sleep 200\n'
                f'sh.Run "explorer.exe ""{dtdd_url}""", 1, False\n'
            )
            existing = ""
            if os.path.exists(vbs_path):
                try:
                    with open(vbs_path, "r", encoding="utf-8") as fh:
                        existing = fh.read()
                except OSError:
                    pass
            if existing != content:
                if DRY_RUN:
                    log(f"[DRY RUN] Would write Parents guide.vbs opening IMDb/CSM/DTDD")
                else:
                    _set_file_attrs(vbs_path, remove=_FILE_ATTRIBUTE_HIDDEN)
                    tmp_vbs = vbs_path + ".tmp"
                    with open(tmp_vbs, "w", encoding="utf-8") as fh:
                        fh.write(content)
                    os.replace(tmp_vbs, vbs_path)
                    _set_file_attrs(vbs_path, add=_FILE_ATTRIBUTE_HIDDEN)
            _write_lnk(lnk_path, vbs_path, "Parents guide")
        except Exception as e:
            log_err(f"Failed creating Parents guide shortcut: {e}")

    if ENABLE_SHORTCUT_TWITTER and title:
        try:
            # Twitter allows non-ASCII hashtags — strip year + non-letter chars
            # but preserve Unicode letters (Korean, Japanese, Arabic, …).
            hashtag_tag = re.sub(r'\s*\(\d{4}\)$', '', title).strip()
            hashtag_tag = ''.join(ch for ch in hashtag_tag if ch.isalpha())
            clean_title = re.sub(r'\s*\(\d{4}\)$', '', title).strip()
            if hashtag_tag or clean_title:
                # One-time migration: old versions used a Twitter.vbs wrapper to
                # open two tabs. Delete it + the stale .lnk so the next create
                # points at the new single-URL format.
                old_vbs = os.path.join(links_dir, "Twitter.vbs")
                old_lnk = os.path.join(links_dir, "Twitter.lnk")
                if os.path.exists(old_vbs):
                    for p in (old_vbs, old_lnk):
                        try: os.remove(p)
                        except Exception: pass

                # Single URL uses Twitter's OR operator to combine hashtag +
                # exact-phrase search, filtered to SEARCH_LANGUAGE. The OR
                # clause is parenthesized so `lang:` applies to the whole
                # alternation, not just the last term.
                parts = []
                if hashtag_tag:
                    parts.append(f"#{hashtag_tag}")
                if clean_title:
                    parts.append(f'"{clean_title}"')
                or_clause = f"({' OR '.join(parts)})" if len(parts) > 1 else parts[0]
                twitter_query = quote(f"{or_clause} lang:{SEARCH_LANGUAGE}")
                make_link("Twitter", f"https://x.com/search?q={twitter_query}")
        except Exception as e:
            log_err(f"Failed creating Twitter shortcut: {e}")

    if service == "radarr":
        tmdb_id = tmdb_or_tvdb_id
        if tmdb_id and ENABLE_SHORTCUT_LETTERBOXD:
            make_link("Letterboxd", f"https://letterboxd.com/tmdb/{tmdb_id}/")
        if tmdb_id and ENABLE_SHORTCUT_TVTIME:
            make_link("TVTime", f"https://app.tvtime.com/movie/{tmdb_id}")

    if service == "sonarr":
        tvdb_id = tmdb_or_tvdb_id
        if tvdb_id and ENABLE_SHORTCUT_TVTIME:
            make_link("TVTime", f"https://app.tvtime.com/series/{tvdb_id}/episodes")

    # MyDramaList shortcut (Korean titles only) — applies to both Sonarr and Radarr.
    # Use the direct show-page URL if get_mdl_rating cached one, else fall back to search.
    if is_korean and ENABLE_SHORTCUT_MYDRAMALIST and title:
        direct = _mdl_url_cache.get(_provider_cache_key(title, year))
        if direct:
            make_link("MyDramaList", direct)
        else:
            clean_title = re.sub(r'\s*\(\d{4}\)$', '', title).strip()
            mdl_query = quote(clean_title)
            # co=3 restricts advanced search to Korea
            make_link("MyDramaList", f"https://mydramalist.com/search?adv=titles&co=3&q={mdl_query}")

    # MyAnimeList shortcut (anime only) — direct show page if cached, else search
    if is_anime and ENABLE_SHORTCUT_MYANIMELIST and title:
        direct = _mal_url_cache.get(_provider_cache_key(title, year))
        if direct:
            make_link("MyAnimeList", direct)
        else:
            clean_title = re.sub(r'\s*\(\d{4}\)$', '', title).strip()
            mal_query = quote(clean_title)
            make_link("MyAnimeList", f"https://myanimelist.net/search/all?q={mal_query}&cat=all")

    # === Subtitle shortcut: Subtitle.vbs opens SubDL + Subsource + OpenSubtitles ===
    if ENABLE_SHORTCUT_SUBTITLES and imdb_id:
        try:
            content_type = "tv" if service == "sonarr" else "movie"
            # URLs go into a VBS literal — reject anything that could break out of
            # the double-quoted string or inject script (e.g., compromised upstream
            # API returning a URL with `"` or newline).
            subdl_url = get_subdl_web_url(imdb_id, content_type)
            if not _is_safe_url(subdl_url):
                subdl_url = f"https://subdl.com/search/{quote(imdb_id)}"
            opensub_url = get_opensubtitles_web_url(imdb_id, content_type)
            if not _is_safe_url(opensub_url):
                opensub_url = (f"https://www.opensubtitles.com/{SEARCH_LANGUAGE}/"
                               f"{SEARCH_LANGUAGE}/search-all/q-{quote(imdb_id)}")
            subsource_url = f"https://subsource.net/search?q={quote(imdb_id)}"

            vbs_path = os.path.join(links_dir, "Subtitle.vbs")
            lnk_path = os.path.join(links_dir, "Subtitle.lnk")

            # Rewrite if missing OR if the resolved URLs differ from on-disk
            # content. Catches drift when a SUBDL/OPENSUBTITLES key is added
            # later, or when upstream URL formats change.
            content = (
                'Set sh = CreateObject("WScript.Shell")\n'
                f'sh.Run "explorer.exe ""{subdl_url}""", 1, False\n'
                'WScript.Sleep 200\n'
                f'sh.Run "explorer.exe ""{subsource_url}""", 1, False\n'
                'WScript.Sleep 200\n'
                f'sh.Run "explorer.exe ""{opensub_url}""", 1, False\n'
            )
            existing = ""
            if os.path.exists(vbs_path):
                try:
                    with open(vbs_path, "r", encoding="utf-8") as fh:
                        existing = fh.read()
                except OSError:
                    pass
            if existing != content:
                if DRY_RUN:
                    log(f"[DRY RUN] Would write Subtitle.vbs opening SubDL/Subsource/OpenSubtitles")
                else:
                    # Clear hidden attribute so we can overwrite, write atomically,
                    # then re-apply the attribute.
                    _set_file_attrs(vbs_path, remove=_FILE_ATTRIBUTE_HIDDEN)
                    tmp_vbs = vbs_path + ".tmp"
                    with open(tmp_vbs, "w", encoding="utf-8") as f:
                        f.write(content)
                    os.replace(tmp_vbs, vbs_path)
                    _set_file_attrs(vbs_path, add=_FILE_ATTRIBUTE_HIDDEN)
            _write_lnk(lnk_path, vbs_path, "Subtitle")
        except Exception as e:
            log_err(f"Failed creating combined subtitle script/shortcut: {e}")

# ==========================
# Service Path Logic
# ==========================
# Per-process memo of full Sonarr/Radarr library responses, keyed by service.
# Populated lazily by get_object_by_path; reused so a sweep over N folders
# doesn't fetch the full library N times. Cleared between sweep runs would be
# wasteful (each main() invocation handles one command) so we don't reset.
_library_cache = {}

def get_object_by_path(service, path):
    """Look up a Sonarr series / Radarr movie by its on-disk path.

    The first call per service fetches the full library list and caches it
    for the lifetime of the process. Subsequent calls scan the cached list
    in memory — turns a sweep from O(folders × library_size) network into
    O(folders + library_size).
    """
    if service not in _library_cache:
        headers = {"X-Api-Key": RADARR_API_KEY if service == "radarr" else SONARR_API_KEY}
        url = (f"{RADARR_API_URL}/api/v3/movie" if service == "radarr"
               else f"{SONARR_API_URL}/api/v3/series")
        try:
            response = http().get(url, headers=headers, timeout=20)
            response.raise_for_status()
            _library_cache[service] = response.json() or []
        except Exception as e:
            log_err(f"{service.capitalize()} list error: {e}")
            return None
    data = _library_cache[service]

    def _normalize(p):
        # Normalize path separators and case, drop trailing slash. Windows paths
        # are case-insensitive, and Sonarr/Radarr store forward-slash variants.
        return os.path.normpath(os.path.normcase((p or "").replace("/", "\\"))).rstrip("\\")

    needle = _normalize(path)
    for item in data:
        item_path = _normalize(item.get("path") or "")
        if item_path and item_path == needle:
            return item
    return None

def radarr_update_path_via_put(movie_id, new_path):
    """Return True if Radarr accepted the path change, False otherwise."""
    if DRY_RUN:
        log(f"[DRY RUN] Would update Radarr movie {movie_id} path -> {new_path}")
        return True
    headers = {"X-Api-Key": RADARR_API_KEY}
    try:
        get_url = f"{RADARR_API_URL}/api/v3/movie/{int(movie_id)}"
        r = http().get(get_url, headers=headers, timeout=20)
        if r.status_code != 200:
            log_err(f"Radarr GET /movie/{movie_id} returned {r.status_code}")
            return False
        movie = r.json() or {}
        movie["path"] = new_path

        put_url = f"{RADARR_API_URL}/api/v3/movie/{int(movie_id)}"
        r2 = http().put(put_url, headers=headers, json=movie, timeout=30)
        if r2.status_code in (200, 202):
            log(f"Updated Radarr movie {movie_id} path -> {new_path}")
            return True
        log_err(f"Radarr PUT /movie/{movie_id} returned {r2.status_code}")
        return False
    except Exception as e:
        log_err(f"Radarr path update error: {e}")
        return False

# Korean detection — tiny in-memory cache so repeated imports of the same
# series/movie don't hit the API each time. Persisted across a single run only.
_korean_cache = {}
_anime_cache = {}

def _language_from_env(prefix):
    """Check Sonarr/Radarr's OriginalLanguage env vars to skip the API call.
    Sonarr uses Sonarr_Series_OriginalLanguages (plural); Radarr uses
    Radarr_Movie_OriginalLanguage. We also accept the unprefixed forms our
    own sweep code uses. Returns lower-case language name or None."""
    if prefix == "Sonarr":
        candidates = ("Sonarr_OriginalLanguage", "Sonarr_OriginalLanguages",
                      "Sonarr_Series_OriginalLanguage", "Sonarr_Series_OriginalLanguages")
    else:
        candidates = ("Radarr_OriginalLanguage", "Radarr_OriginalLanguages",
                      "Radarr_Movie_OriginalLanguage")
    for key in candidates:
        val = os.environ.get(key, "").strip()
        if val:
            return val.lower()
    return None

def is_korean_radarr_movie(movie_id):
    """Return True if the Radarr movie's original language is Korean.
    Tries env var first; falls back to Radarr API (cached per-run)."""
    env_lang = _language_from_env("Radarr")
    if env_lang is not None:
        return "korean" in env_lang
    if not movie_id:
        return False
    cache_key = f"radarr:{movie_id}"
    if cache_key in _korean_cache:
        return _korean_cache[cache_key]
    headers = {"X-Api-Key": RADARR_API_KEY}
    try:
        url = f"{RADARR_API_URL}/api/v3/movie/{int(movie_id)}"
        r = http().get(url, headers=headers, timeout=15)
        if r.status_code != 200:
            result = False
        else:
            movie = r.json() or {}
            lang = (movie.get("originalLanguage") or {}).get("name", "") or ""
            result = lang.strip().lower() == "korean"
    except Exception as e:
        log_err(f"Radarr Korean detection error: {e}")
        result = False
    _korean_cache[cache_key] = result
    return result

def is_korean_sonarr_series(series_id):
    """Return True if the Sonarr series' original language is Korean.
    Tries env var first; falls back to Sonarr API (cached per-run)."""
    env_lang = _language_from_env("Sonarr")
    if env_lang is not None:
        return "korean" in env_lang
    if not series_id:
        return False
    cache_key = f"sonarr:{series_id}"
    if cache_key in _korean_cache:
        return _korean_cache[cache_key]
    headers = {"X-Api-Key": SONARR_API_KEY}
    try:
        url = f"{SONARR_API_URL}/api/v3/series/{int(series_id)}"
        r = http().get(url, headers=headers, timeout=15)
        if r.status_code != 200:
            result = False
        else:
            series = r.json() or {}
            lang = (series.get("originalLanguage") or {}).get("name", "") or ""
            result = lang.strip().lower() == "korean"
    except Exception as e:
        log_err(f"Korean detection error: {e}")
        result = False
    _korean_cache[cache_key] = result
    return result

def is_anime_sonarr_series(series_id, path=None):
    """Return True if this is an anime. Checks, in order:
       1. Sonarr_Series_Type env var (== 'anime')
       2. Path contains an '\\anime\\' segment (case-insensitive, user convention)
       3. Sonarr API seriesType == 'anime'
    """
    # 1. Env var from Sonarr hook
    stype = os.environ.get("Sonarr_Series_Type", "").strip().lower()
    if stype:
        return stype == "anime"
    # 2. Path-based heuristic (case-insensitive; Windows paths usually are)
    if path and (os.sep + "anime" + os.sep) in (path + os.sep).lower():
        return True
    # 3. API lookup with cache
    if not series_id:
        return False
    cache_key = f"sonarr:{series_id}"
    if cache_key in _anime_cache:
        return _anime_cache[cache_key]
    headers = {"X-Api-Key": SONARR_API_KEY}
    try:
        url = f"{SONARR_API_URL}/api/v3/series/{int(series_id)}"
        r = http().get(url, headers=headers, timeout=15)
        if r.status_code != 200:
            result = False
        else:
            series = r.json() or {}
            result = (series.get("seriesType", "") or "").strip().lower() == "anime"
    except Exception as e:
        log_err(f"Anime detection error: {e}")
        result = False
    _anime_cache[cache_key] = result
    return result

def is_anime_radarr_movie(movie_id, path=None):
    """Return True if this is an anime movie. Checks:
       1. Path contains an '\\anime\\' segment (case-insensitive)
       2. Radarr API: genres include 'Animation' AND originalLanguage == Japanese
    """
    if path and (os.sep + "anime" + os.sep) in (path + os.sep).lower():
        return True
    if not movie_id:
        return False
    cache_key = f"radarr:{movie_id}"
    if cache_key in _anime_cache:
        return _anime_cache[cache_key]
    headers = {"X-Api-Key": RADARR_API_KEY}
    try:
        url = f"{RADARR_API_URL}/api/v3/movie/{int(movie_id)}"
        r = http().get(url, headers=headers, timeout=15)
        if r.status_code != 200:
            result = False
        else:
            movie = r.json() or {}
            genres = [g.lower() for g in (movie.get("genres") or [])]
            lang = (movie.get("originalLanguage") or {}).get("name", "").lower()
            result = "animation" in genres and lang == "japanese"
    except Exception as e:
        log_err(f"Anime detection error: {e}")
        result = False
    _anime_cache[cache_key] = result
    return result

def sonarr_update_path_via_put(series_id, new_path):
    """Return True if Sonarr accepted the path change, False otherwise."""
    if DRY_RUN:
        log(f"[DRY RUN] Would update Sonarr series {series_id} path -> {new_path}")
        return True
    headers = {"X-Api-Key": SONARR_API_KEY}
    try:
        get_url = f"{SONARR_API_URL}/api/v3/series/{int(series_id)}"
        r = http().get(get_url, headers=headers, timeout=20)
        if r.status_code != 200:
            log_err(f"Sonarr GET /series/{series_id} returned {r.status_code}")
            return False
        series = r.json() or {}
        series["path"] = new_path

        put_url = f"{SONARR_API_URL}/api/v3/series/{int(series_id)}"
        r2 = http().put(put_url, headers=headers, json=series, timeout=30)
        if r2.status_code in (200, 202):
            log(f"Updated Sonarr series {series_id} path -> {new_path}")
            return True
        log_err(f"Sonarr PUT /series/{series_id} returned {r2.status_code}")
        return False
    except Exception as e:
        log_err(f"Sonarr path update error: {e}")
        return False

# ==========================
# Sonarr / Radarr workers
# ==========================
# Per-service config: env-var names + detection/update callbacks.
_SERVICE_ADAPTERS = {
    "sonarr": {
        "label":        "Sonarr",
        "id_env":       "Sonarr_Series_Id",
        "imdb_env":     "Sonarr_Series_ImdbId",
        "other_id_env": "Sonarr_Series_TvdbId",
        "title_env":    "Sonarr_Series_Title",
        "year_env":     "Sonarr_Series_Year",
        "is_anime":     lambda obj_id, p: is_anime_sonarr_series(obj_id, p),
        "is_korean":    lambda obj_id: is_korean_sonarr_series(obj_id),
        "update_path":  lambda obj_id, np: sonarr_update_path_via_put(int(obj_id), np),
    },
    "radarr": {
        "label":        "Radarr",
        "id_env":       "Radarr_Movie_Id",
        "imdb_env":     "Radarr_Movie_ImdbId",
        "other_id_env": "Radarr_Movie_TmdbId",
        "title_env":    "Radarr_Movie_Title",
        "year_env":     "Radarr_Movie_Year",
        "is_anime":     lambda obj_id, p: is_anime_radarr_movie(obj_id, p),
        "is_korean":    lambda obj_id: is_korean_radarr_movie(obj_id),
        "update_path":  lambda obj_id, np: radarr_update_path_via_put(int(obj_id), np),
    },
}

def _process(service, path):
    """Shared post-import flow for Sonarr/Radarr."""
    cfg = _SERVICE_ADAPTERS[service]
    imdb     = os.environ.get(cfg["imdb_env"])
    other_id = os.environ.get(cfg["other_id_env"])
    title    = os.environ.get(cfg["title_env"], "")

    # Sweep mode: skip if rating was checked recently AND the folder still
    # has a rating suffix. If the suffix is missing (no rating yet, or user
    # stripped it manually), re-check regardless of cache age.
    if (RATING_ONLY and imdb
            and _has_rating_suffix(os.path.basename(path))
            and _rating_cache_is_fresh(imdb)):
        log_debug(f"Cache fresh, skipping {title or path}")
        return

    log(f"{cfg['label']} post-import: {title or path}")

    obj_id = os.environ.get(cfg["id_env"])
    if not obj_id:
        obj_pre = get_object_by_path(service, path)
        obj_id = (obj_pre or {}).get("id")

    lock_key = f"{service}_{obj_id or re.sub(r'\\W+', '_', title or 'unknown')}"
    with _fs_lock(lock_key) as lock_acquired:
        if not lock_acquired:
            # Another process is already handling this series/movie. Skip
            # rather than racing — every operation here is idempotent, but
            # concurrent rename + API-update can fight each other.
            log(f"Skipping (lock held by another process): {title or path}")
            return
        # Detect content class once — used for rating dispatch + shortcut selection
        need_korean = ENABLE_MDL_RATING or ENABLE_SHORTCUT_MYDRAMALIST
        need_anime  = ENABLE_MAL_RATING or ENABLE_SHORTCUT_MYANIMELIST
        is_anime    = cfg["is_anime"](obj_id, path) if need_anime else False
        # Anime and Korean are mutually exclusive in practice; anime takes priority
        is_korean   = (not is_anime) and need_korean and cfg["is_korean"](obj_id)

        # Year from service env var (preferred) or parse from title
        year = os.environ.get(cfg["year_env"]) or ""
        if not year:
            m = re.search(r'\((\d{4})\)\s*$', title or '')
            year = m.group(1) if m else ""

        provider_failed = False
        try:
            rating, source = get_rating_for_title(imdb, title, year, is_korean=is_korean, is_anime=is_anime)
        except ProviderUnavailable as e:
            # Provider (MAL/MDL) is having a transient outage. Keep the
            # folder's existing rating — DON'T fall back to IMDb, which
            # would silently rebrand an anime as IMDb and rename the folder.
            log(f"Rating provider unavailable for {title or path}: {e}; keeping existing rating")
            if RATING_ONLY:
                return
            rating, source = "N/A", "IMDb"  # skips rename + cache below
            provider_failed = True          # also skips tooltip rewrite below
        target_suffix = f" [{source} {rating}]"

        if rating != "N/A" and ENABLE_RENAME_FOLDER and not os.path.basename(path).endswith(target_suffix):
            new_path = rename_folder(path, rating, source)
            renamed = (os.path.abspath(new_path) != os.path.abspath(path)
                       and os.path.exists(new_path))
            if renamed and ENABLE_UPDATE_SERVICE_PATH:
                rid = obj_id
                if not rid:
                    obj = get_object_by_path(service, new_path)
                    rid = (obj or {}).get("id")
                api_ok = cfg["update_path"](rid, new_path) if rid else False
                if not api_ok:
                    # Roll back so disk and service stay in sync
                    if rollback_rename(new_path, path):
                        new_path = path  # use old path for subsequent steps
            path = new_path

        # Always remember when we last checked, so the next sweep can skip us
        if rating != "N/A":
            _rating_cache_set(imdb, rating, source)

        # Sweep mode stops here — webhook continues with icon/shortcuts/tooltip
        if RATING_ONLY:
            return

        if not os.path.isdir(path):
            log_err(f"Skipping post-rename steps: folder does not exist on disk ({path})")
            return
        if ENABLE_CREATE_FOLDER_ICON: create_folder_icon(path)
        if ENABLE_CREATE_SHORTCUTS:
            create_shortcuts(service, path, imdb, other_id, title,
                             is_korean=is_korean, is_anime=is_anime, year=year)
        # Tooltip: "Plot... — [SOURCE X.X]" — shown on folder hover in Explorer.
        # Skip on provider outage: the existing tooltip (if any) likely has the
        # correct rating; rewriting it without the rating part would be a
        # downgrade until the next successful run restores it.
        if ENABLE_SET_TOOLTIP and not provider_failed:
            plot = get_omdb_plot(imdb) if imdb else ""
            if plot:
                # Single line — desktop.ini InfoTip doesn't support real newlines
                tip = f"{plot}  [{source} {rating}]" if rating != "N/A" else plot
                set_folder_tooltip(path, tip)

def process_sonarr(path):
    return _process("sonarr", path)

def process_radarr(path):
    return _process("radarr", path)

# ==========================
# --validate
# ==========================
def validate_config():
    """Run a config + connectivity health check. Returns 0 OK, 1 on issues."""
    issues = []
    checks = []

    def ok(msg): checks.append(("OK", msg))
    def fail(msg): checks.append(("FAIL", msg)); issues.append(msg)

    # Env / config
    for key in ("SONARR_API_KEY", "RADARR_API_KEY", "OMDB_API_KEY",
                "SUBDL_API_KEY", "OPENSUBTITLES_API_KEY"):
        if os.environ.get(key, ""):
            ok(f"env  {key} is set")
        else:
            fail(f"env  {key} is missing")

    # pywin32
    if HAS_WIN32COM:
        ok("module  pywin32 available")
    else:
        fail("module  pywin32 missing — shortcut creation will be skipped")

    # FolderIconCreator.exe
    if FOLDER_ICON_EXE and os.path.isfile(FOLDER_ICON_EXE):
        ok(f"file  FolderIconCreator at {FOLDER_ICON_EXE}")
    else:
        fail(f"file  FolderIconCreator missing at {FOLDER_ICON_EXE or '(unset)'}")

    # Icons directory. Only check for icons we actually use as .lnk display
    # icons (SubDL.ico / Subsource.ico in the repo are legacy assets from when
    # those providers had their own .lnk — today they're only opened from
    # within Subtitle.vbs, which uses Subtitle.ico).
    expected_icons = ["IMDb", "Parents guide", "Twitter", "TVTime", "Letterboxd",
                       "Subtitle", "MyDramaList", "MyAnimeList"]
    missing_icons = [i for i in expected_icons if not os.path.isfile(os.path.join(ICONS_DIR, f"{i}.ico"))]
    if not missing_icons:
        ok(f"icons  all {len(expected_icons)} icon files present")
    else:
        fail(f"icons  missing: {', '.join(missing_icons)}")

    # Connectivity probes
    def probe(name, url, extra=None, expect_status=200):
        try:
            r = http().get(url, headers=extra or {}, timeout=10)
            if r.status_code == expect_status:
                ok(f"http  {name} reachable ({r.status_code})")
                return True
            fail(f"http  {name} returned {r.status_code}")
        except Exception as e:
            fail(f"http  {name} unreachable: {_redact(e)}")
        return False

    def probe_any(name, urls, extra=None, expect_status=200):
        """Pass if any URL responds. Used for providers with multiple mirrors."""
        last_err = None
        for url in urls:
            try:
                r = http().get(url, headers=extra or {}, timeout=10)
                if r.status_code == expect_status:
                    ok(f"http  {name} reachable ({r.status_code})")
                    return True
                last_err = f"returned {r.status_code}"
            except Exception as e:
                last_err = _redact(e)
        fail(f"http  {name} unreachable: {last_err}")
        return False

    probe("Sonarr", f"{SONARR_API_URL}/api/v3/system/status", {"X-Api-Key": SONARR_API_KEY})
    probe("Radarr", f"{RADARR_API_URL}/api/v3/system/status", {"X-Api-Key": RADARR_API_KEY})
    # IMDb GraphQL — primary rating source. POST-only endpoint so the generic
    # probe() helper can't cover it; call get_imdb_rating_from_graphql directly.
    if get_imdb_rating_from_graphql("tt0111161") != "N/A":
        ok("http  IMDb GraphQL reachable")
    else:
        fail("http  IMDb GraphQL unreachable or returned no rating")
    # OMDb: status 200 isn't enough — bad API keys return 200 with
    # {"Response":"False","Error":"Invalid API key!"}.
    try:
        r = http().get(f"https://www.omdbapi.com/?apikey={OMDB_API_KEY}&i=tt0111161", timeout=10)
        if r.status_code != 200:
            fail(f"http  OMDb returned {r.status_code}")
        else:
            body = r.json() or {}
            if str(body.get("Response", "")).lower() == "true":
                ok("http  OMDb reachable (200)")
            else:
                fail(f"http  OMDb rejected request: {body.get('Error') or 'unknown error'}")
    except Exception as e:
        fail(f"http  OMDb unreachable: {_redact(e)}")
    probe_any("Kuryana", [f"{u}/search/q/reverse" for u in KURYANA_BASE_URLS])
    probe("Jikan",  f"{JIKAN_BASE_URL}/anime?q=frieren&limit=1")
    if SUBDL_API_KEY:
        probe("SubDL", f"https://api.subdl.com/api/v1/subtitles?api_key={SUBDL_API_KEY}&imdb_id=tt0111161&type=movie")
    if OPENSUBTITLES_API_KEY:
        probe("OpenSubtitles",
              "https://api.opensubtitles.com/api/v1/features?imdb_id=111161",
              {"Api-Key": OPENSUBTITLES_API_KEY, "User-Agent": "arr-finisher/1.0"})

    # Report
    print(f"\n=== arr_finisher {__version__} --validate ===\n")
    for status, msg in checks:
        icon = "OK " if status == "OK" else "FAIL"
        print(f"  [{icon}] {msg}")
    print(f"\n{len(checks) - len(issues)} passed, {len(issues)} failed")
    return 0 if not issues else 1

# ==========================
# --sweep
# ==========================
def _sweep_one(service, path):
    """Populate env vars from service API and invoke process_<service>(path)."""
    obj = get_object_by_path(service, path)
    if not obj:
        return False  # service doesn't know about this folder

    # Env-var stuffing lets us reuse process_sonarr/process_radarr without refactoring.
    saved = {}
    def set_env(k, v):
        saved[k] = os.environ.get(k)
        os.environ[k] = str(v) if v is not None else ""

    try:
        if service == "sonarr":
            lang = (obj.get("originalLanguage") or {}).get("name", "")
            set_env("Sonarr_Series_Id", obj.get("id"))
            set_env("Sonarr_Series_ImdbId", obj.get("imdbId") or "")
            set_env("Sonarr_Series_TvdbId", obj.get("tvdbId") or "")
            set_env("Sonarr_Series_Title", obj.get("title") or "")
            set_env("Sonarr_Series_Year", obj.get("year") or "")
            set_env("Sonarr_OriginalLanguage", lang)
            set_env("Sonarr_Series_Type", obj.get("seriesType") or "")
            process_sonarr(path)
        else:
            lang = (obj.get("originalLanguage") or {}).get("name", "")
            set_env("Radarr_Movie_Id", obj.get("id"))
            set_env("Radarr_Movie_ImdbId", obj.get("imdbId") or "")
            set_env("Radarr_Movie_TmdbId", obj.get("tmdbId") or "")
            set_env("Radarr_Movie_Title", obj.get("title") or "")
            set_env("Radarr_Movie_Year", obj.get("year") or "")
            set_env("Radarr_Movie_OriginalLanguage", lang)
            process_radarr(path)
        return True
    finally:
        for k, v in saved.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v

def _parse_roots_arg(values):
    """Parse a list of 'path:service' strings into [(path, service), ...].
    Uses rsplit so Windows drive letters in paths (e.g. 'D:\\TV Shows:sonarr')
    are handled correctly."""
    out = []
    for raw in values:
        parts = raw.rsplit(":", 1)
        if len(parts) != 2:
            raise ValueError(f"--roots entry must be 'path:service', got {raw!r}")
        p, s = parts[0].strip(), parts[1].strip().lower()
        if s not in ("sonarr", "radarr"):
            raise ValueError(f"--roots service must be sonarr or radarr, got {s!r}")
        out.append((p, s))
    return out

# No hardcoded fallback — too easy to ship the original author's drive layout.
# When neither --roots, ARR_FINISHER_SWEEP_ROOTS, nor service auto-discovery
# yields anything, the sweep will refuse to run and tell the user to configure.
_HARDCODED_FALLBACK_ROOTS = []

# Backward-compat alias (other code/users may reference this constant).
DEFAULT_SWEEP_ROOTS = _HARDCODED_FALLBACK_ROOTS

def _discover_sweep_roots():
    """Query Sonarr and Radarr for their configured root folders, returning
    [(path, service), ...]. Returns [] if neither service is configured or
    reachable. Short timeout so a single down service doesn't stall sweep."""
    out = []
    for service, base, key in (("sonarr", SONARR_API_URL, SONARR_API_KEY),
                               ("radarr", RADARR_API_URL, RADARR_API_KEY)):
        if not key:
            continue
        try:
            r = http().get(f"{base}/api/v3/rootfolder",
                           headers={"X-Api-Key": key}, timeout=5)
            if r.status_code != 200:
                continue
            for entry in (r.json() or []):
                p = entry.get("path") or ""
                if p:
                    out.append((os.path.normpath(p), service))
        except Exception as e:
            log_debug(f"{service} root discovery failed: {_redact(e)}")
    return out

def _default_sweep_roots():
    """Decide which sweep roots to walk. Precedence:
      1. ARR_FINISHER_SWEEP_ROOTS env var (pipe-separated path:service pairs).
      2. Auto-discovery from Sonarr/Radarr /api/v3/rootfolder.
      3. Empty list — caller will refuse to sweep with a clear error.
    """
    env_val = os.environ.get("ARR_FINISHER_SWEEP_ROOTS", "").strip()
    if env_val:
        try:
            return _parse_roots_arg([s for s in env_val.split("|") if s.strip()])
        except ValueError as e:
            log_err(f"Bad ARR_FINISHER_SWEEP_ROOTS: {e}; trying auto-discovery")
    discovered = _discover_sweep_roots()
    if discovered:
        log(f"Auto-discovered sweep roots from services: {discovered}")
        return discovered
    return []

class _EventCounter(logging.Handler):
    """Counts log records matching known event patterns. Used for sweep summary."""
    PATTERNS = {
        "renamed":         re.compile(r"^Renamed "),
        "api_updated":     re.compile(r"^Updated (?:Sonarr|Radarr) "),
        "shortcut_new":    re.compile(r"^Shortcut: "),
        "tooltip":         re.compile(r"^Tooltip set on "),
        "rollbacks":       re.compile(r"^Rolled back disk rename"),
        "mal_rejected":    re.compile(r"^MAL match rejected"),
        "mdl_rejected":    re.compile(r"^MDL match rejected"),
        "provider_outage": re.compile(r"^Rating provider unavailable"),
    }
    def __init__(self):
        super().__init__(level=logging.INFO)
        self.counts = {k: 0 for k in self.PATTERNS}
    def emit(self, record):
        msg = record.getMessage()
        for k, rx in self.PATTERNS.items():
            if rx.search(msg):
                self.counts[k] += 1
                return

def sweep_library(roots=None, force_refresh=False):
    """Walk library roots and refresh ratings only (rename + API path-sync if changed).
    Skips folders whose rating was checked in the last RATING_CACHE_TTL_DAYS days,
    unless force_refresh is True (treats every folder as stale).
    Icons, shortcuts, and tooltips are NOT touched — those are webhook-time work."""
    if roots is None:
        roots = _default_sweep_roots()
    elif isinstance(roots, list) and roots and isinstance(roots[0], str):
        # Accept a list of "path:service" strings from CLI
        try:
            roots = _parse_roots_arg(roots)
        except ValueError as e:
            log_err(str(e))
            return 1

    if not roots:
        log_err("Sweep: no roots configured. Set ARR_FINISHER_SWEEP_ROOTS, pass "
                "--roots, or configure Sonarr/Radarr root folders so auto-discovery works.")
        return 1

    global RATING_ONLY
    prev_rating_only = RATING_ONLY
    RATING_ONLY = True

    # --force-refresh: pretend every cache entry is stale by temporarily
    # replacing the freshness check. Restored in `finally` below.
    fresh_check_orig = None
    if force_refresh:
        global _rating_cache_is_fresh
        fresh_check_orig = _rating_cache_is_fresh
        _rating_cache_is_fresh = lambda imdb_id: False  # noqa: E731
        log("Sweep: --force-refresh — treating every entry as stale")

    # Attach a counter to tally events during the sweep. Attached to our named
    # logger (not root) — propagate=False on that logger means root won't see
    # the records.
    counter = _EventCounter()
    arr_logger = logging.getLogger(LOGGER_NAME)
    arr_logger.addHandler(counter)

    start_ts = time.time()
    processed = skipped = unknown = 0
    try:
        for root, service in roots:
            if not os.path.isdir(root):
                log(f"Sweep: root not found, skipping: {root}")
                continue
            log(f"Sweep: scanning {root} ({service}) — rating-only mode, {RATING_CACHE_TTL_DAYS}d TTL")
            for entry in sorted(os.listdir(root)):
                sub = os.path.join(root, entry)
                if not os.path.isdir(sub) or entry.startswith("."):
                    continue
                try:
                    if _sweep_one(service, sub):
                        processed += 1
                    else:
                        unknown += 1
                        log(f"Sweep: {service} doesn't know {sub}")
                except Exception as e:
                    skipped += 1
                    log_err(f"Sweep: error processing {sub}: {e}")
    finally:
        arr_logger.removeHandler(counter)
        RATING_ONLY = prev_rating_only
        if fresh_check_orig is not None:
            _rating_cache_is_fresh = fresh_check_orig
        _save_rating_cache()

    elapsed = time.time() - start_ts
    c = counter.counts
    log(
        f"Sweep complete: {processed} processed, {unknown} unknown, {skipped} errored "
        f"in {elapsed:.1f}s — {c['renamed']} rating change(s), {c['api_updated']} API-sync, "
        f"{c['mal_rejected']+c['mdl_rejected']} rejected, {c['rollbacks']} rollback(s), "
        f"{c['provider_outage']} provider outage(s)"
    )
    return 0

def regenerate_shortcuts(roots=None):
    """Walk library roots and rebuild every Links/ shortcut from current code.
    Useful after URL formats change (e.g. provider rebrand, new query params).
    Does NOT re-check ratings, icons, or tooltips."""
    if roots is None:
        roots = _default_sweep_roots()
    elif isinstance(roots, list) and roots and isinstance(roots[0], str):
        try:
            roots = _parse_roots_arg(roots)
        except ValueError as e:
            log_err(str(e))
            return 1
    if not roots:
        log_err("--regenerate-shortcuts: no roots configured.")
        return 1

    global FORCE_REGENERATE_SHORTCUTS
    prev_force = FORCE_REGENERATE_SHORTCUTS
    FORCE_REGENERATE_SHORTCUTS = True
    start_ts = time.time()
    processed = unknown = errored = 0
    try:
        for root, service in roots:
            if not os.path.isdir(root):
                log(f"Regen: root not found, skipping: {root}")
                continue
            log(f"Regen: scanning {root} ({service})")
            for entry in sorted(os.listdir(root)):
                sub = os.path.join(root, entry)
                if not os.path.isdir(sub) or entry.startswith("."):
                    continue
                try:
                    obj = get_object_by_path(service, sub)
                    if not obj:
                        unknown += 1
                        log(f"Regen: {service} doesn't know {sub}")
                        continue
                    title = obj.get("title") or ""
                    year = str(obj.get("year") or "")
                    imdb_id = obj.get("imdbId") or ""
                    if service == "sonarr":
                        sid = obj.get("id")
                        tvdb_id = obj.get("tvdbId") or ""
                        is_anime = is_anime_sonarr_series(sid, sub) if (
                            ENABLE_MAL_RATING or ENABLE_SHORTCUT_MYANIMELIST) else False
                        is_korean = (not is_anime) and is_korean_sonarr_series(sid) if (
                            ENABLE_MDL_RATING or ENABLE_SHORTCUT_MYDRAMALIST) else False
                        create_shortcuts("sonarr", sub, imdb_id, tvdb_id, title,
                                         is_korean=is_korean, is_anime=is_anime, year=year)
                    else:
                        mid = obj.get("id")
                        tmdb_id = obj.get("tmdbId") or ""
                        is_anime = is_anime_radarr_movie(mid, sub) if (
                            ENABLE_MAL_RATING or ENABLE_SHORTCUT_MYANIMELIST) else False
                        is_korean = (not is_anime) and is_korean_radarr_movie(mid) if (
                            ENABLE_MDL_RATING or ENABLE_SHORTCUT_MYDRAMALIST) else False
                        create_shortcuts("radarr", sub, imdb_id, tmdb_id, title,
                                         is_korean=is_korean, is_anime=is_anime, year=year)
                    processed += 1
                except Exception as e:
                    errored += 1
                    log_err(f"Regen: error processing {sub}: {e}")
    finally:
        FORCE_REGENERATE_SHORTCUTS = prev_force
    log(f"Regen complete: {processed} processed, {unknown} unknown, "
        f"{errored} errored in {time.time()-start_ts:.1f}s")
    return 0

def clear_rating_cache(imdb_id=None):
    """Delete the on-disk rating cache (or just one entry if imdb_id is given)."""
    if not imdb_id:
        try:
            os.remove(RATING_CACHE_PATH)
            log(f"Cleared rating cache: {RATING_CACHE_PATH}")
        except FileNotFoundError:
            log("Rating cache already absent")
        except OSError as e:
            log_err(f"Could not clear rating cache: {e}")
            return 1
        return 0
    # Single-entry refresh
    global _rating_cache_dirty
    cache = _load_rating_cache()
    if imdb_id in cache:
        del cache[imdb_id]
        _rating_cache_dirty = True
        _save_rating_cache()
        log(f"Removed rating cache entry for {imdb_id}")
    else:
        log(f"No rating cache entry for {imdb_id}")
    return 0

# ==========================
# Setup-help sidecar
# ==========================
_SETUP_HELP_PATH = os.path.join(SCRIPT_DIR, "arr_finisher_setup.txt")

def _check_critical_config(sonarr_event=None, radarr_event=None):
    """Return list of (var_name, help_text) for missing critical config.
    `sonarr_event` / `radarr_event` are the lowercased event names — when set,
    we require the matching service API key."""
    missing = []
    if not OMDB_API_KEY:
        missing.append(("OMDB_API_KEY",
                        "Free key from https://www.omdbapi.com/apikey.aspx — "
                        "required for any rating fetch."))
    if sonarr_event and not SONARR_API_KEY:
        missing.append(("SONARR_API_KEY",
                        "Sonarr → Settings → General → API Key"))
    if radarr_event and not RADARR_API_KEY:
        missing.append(("RADARR_API_KEY",
                        "Radarr → Settings → General → API Key"))
    if ENABLE_CREATE_FOLDER_ICON and (not FOLDER_ICON_EXE
                                       or not os.path.isfile(FOLDER_ICON_EXE)):
        missing.append(("FOLDER_ICON_EXE",
                        "Absolute path to Folder-Icon-Creator's Creator.exe. "
                        "See README install steps."))
    return missing

def _update_setup_help(missing):
    """Write or remove the setup-help sidecar based on `missing` config items."""
    if not missing:
        try:
            os.remove(_SETUP_HELP_PATH)
        except OSError:
            pass
        return
    try:
        with open(_SETUP_HELP_PATH, "w", encoding="utf-8") as fh:
            fh.write("arr_finisher needs configuration before it can do its job.\n")
            fh.write(f"Last checked: {datetime.now().isoformat(timespec='seconds')}\n\n")
            fh.write("Edit `.env` in the repo directory and set:\n\n")
            for name, hint in missing:
                fh.write(f"  - {name}\n      {hint}\n\n")
            fh.write("Then re-trigger the import in Sonarr/Radarr, or run:\n")
            fh.write("  python arr_finisher.py --validate\n")
            fh.write("for a full health check.\n")
    except OSError:
        pass

# ==========================
# Entrypoint
# ==========================
def main():
    import argparse
    parser = argparse.ArgumentParser(
        prog="arr_finisher",
        description="Post-import finisher for Sonarr/Radarr: rating suffix, shortcuts, folder icon.",
    )
    parser.add_argument("--dry-run", action="store_true",
                        help="Log intended actions but don't touch disk, APIs, or external tools.")
    parser.add_argument("--validate", action="store_true",
                        help="Run config + connectivity checks and exit.")
    parser.add_argument("--sweep", action="store_true",
                        help="Walk library roots and (re)process every folder.")
    parser.add_argument("--force-refresh", action="store_true",
                        help="With --sweep: ignore the rating-cache TTL and "
                             "re-fetch every folder's rating.")
    parser.add_argument("--regenerate-shortcuts", action="store_true",
                        help="Walk library roots and rebuild every Links/ shortcut "
                             "from current code (useful when URL formats change). "
                             "Does NOT re-check ratings.")
    parser.add_argument("--roots", nargs="*",
                        help="Override sweep / regen roots as path:service pairs, "
                             "e.g. --roots 'D:\\TV Shows:sonarr' 'D:\\Movies:radarr'")
    parser.add_argument("--service", choices=["sonarr", "radarr"],
                        help="Manual mode: service to use with --path.")
    parser.add_argument("--path", help="Manual mode: folder to process.")
    parser.add_argument("--clear-cache", action="store_true",
                        help="Delete the rating-freshness cache and exit.")
    parser.add_argument("--refresh", metavar="IMDB_ID",
                        help="Remove a single IMDb ID from the rating cache and exit "
                             "(next sweep will re-fetch it).")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="Enable DEBUG-level logging (idempotent 'Already exists' etc.).")
    parser.add_argument("--version", action="version", version=f"arr_finisher {__version__}")
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger(LOGGER_NAME).setLevel(logging.DEBUG)

    # --service and --path must be paired
    if bool(args.service) != bool(args.path):
        parser.error("--service and --path must be used together")

    global DRY_RUN
    DRY_RUN = args.dry_run
    if DRY_RUN:
        log("DRY RUN mode — no disk or API mutations will occur")

    if args.clear_cache:
        return clear_rating_cache()
    if args.refresh:
        return clear_rating_cache(args.refresh)
    if args.validate:
        return validate_config()

    sonarr_event = os.environ.get("Sonarr_EventType", "").lower()
    radarr_event = os.environ.get("Radarr_EventType", "").lower()

    # Sonarr/Radarr "Test" webhook: the user is verifying the script wires up.
    # Acknowledge before any config check, so a fresh user testing connectivity
    # doesn't get a spurious "missing config" sidecar before they've finished setup.
    if sonarr_event == "test" or radarr_event == "test":
        log(f"{(sonarr_event or radarr_event).capitalize()} test event — OK")
        return 0

    # For real work (sweep, manual, webhook): refresh the setup-help sidecar.
    missing = _check_critical_config(sonarr_event=sonarr_event, radarr_event=radarr_event)
    _update_setup_help(missing)
    if missing:
        names = ", ".join(name for name, _ in missing)
        log_err(f"Missing config: {names}. See {_SETUP_HELP_PATH} for setup help.")
        # Webhook callers expect a quiet exit — don't error if the user is
        # running for the first time. They'll see the sidecar.
        return 0 if (sonarr_event or radarr_event) else 1

    if args.sweep:
        return sweep_library(args.roots or None, force_refresh=args.force_refresh)
    if args.regenerate_shortcuts:
        return regenerate_shortcuts(args.roots or None)
    if args.service and args.path:
        # Manual mode: route through _sweep_one so env vars get populated from
        # the service API (title, imdb_id, year, language). Without this the
        # process_*() functions see empty fields and silently do nothing.
        if not _sweep_one(args.service, args.path):
            log_err(f"{args.service} doesn't know about {args.path}")
            return 1
        _save_rating_cache()
        return 0

    # Default: arr webhook (env-driven)
    if sonarr_event == "download":
        sp = os.environ.get("Sonarr_Series_Path")
        if sp: process_sonarr(sp)
    elif radarr_event == "download":
        mp = os.environ.get("Radarr_Movie_Path")
        if mp: process_radarr(mp)
    # Persist any rating-cache updates from this webhook so the next sweep
    # has a warm cache and can skip this folder.
    _save_rating_cache()
    return 0

if __name__ == "__main__":
    sys.exit(main() or 0)