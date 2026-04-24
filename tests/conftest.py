"""Pytest fixtures shared across all test modules."""
import os, sys, shutil, tempfile
import pytest

# Put the script directory on sys.path so `import arr_finisher` works
HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import arr_finisher as f  # noqa


@pytest.fixture
def staging(tmp_path):
    """A fresh temp directory for fake series/movie folders. Auto-cleaned."""
    d = tmp_path / "lib"
    d.mkdir()
    yield str(d)


@pytest.fixture(autouse=True)
def clear_env_vars():
    """Scrub all Sonarr_* / Radarr_* env vars before AND after each test.
    Case-insensitive because Windows reports env keys as uppercase on iteration."""
    def _clear():
        for k in list(os.environ):
            if k.lower().startswith(("sonarr_", "radarr_")):
                del os.environ[k]
    _clear()
    yield
    _clear()


@pytest.fixture
def patch_providers(monkeypatch):
    """
    Replace real rating-provider / API-call functions with deterministic stubs.
    Returns a dict the test can customize (e.g. `patches['mdl_rating'] = ("9.1", "MDL")`).
    """
    patches = {
        "imdb_rating": "N/A",
        "mdl_rating": None,
        "mal_rating": None,
        "subdl_url": "https://subdl.com/fake",
        "opensub_url": "https://opensub.com/fake",
        "omdb_plot": "",
        "sonarr_put_ok": True,
        "radarr_put_ok": True,
    }
    monkeypatch.setattr(f, "get_imdb_rating",       lambda imdb_id: patches["imdb_rating"])
    monkeypatch.setattr(f, "get_mdl_rating",        lambda title, year=None: patches["mdl_rating"])
    monkeypatch.setattr(f, "get_mal_rating",        lambda title, year=None: patches["mal_rating"])
    monkeypatch.setattr(f, "get_subdl_web_url",     lambda *a, **kw: patches["subdl_url"])
    monkeypatch.setattr(f, "get_opensubtitles_web_url", lambda *a, **kw: patches["opensub_url"])
    monkeypatch.setattr(f, "get_omdb_plot",         lambda imdb_id: patches["omdb_plot"])
    monkeypatch.setattr(f, "sonarr_update_path_via_put", lambda *a, **kw: patches["sonarr_put_ok"])
    monkeypatch.setattr(f, "radarr_update_path_via_put", lambda *a, **kw: patches["radarr_put_ok"])

    # Also stub out anime/korean detection so tests never reach real Sonarr/Radarr APIs.
    # Detection falls through to env vars that the test sets; no API call is made.
    def fake_is_korean_sonarr(series_id):
        return "korean" in os.environ.get("Sonarr_OriginalLanguage", "").lower()
    def fake_is_korean_radarr(movie_id):
        return "korean" in os.environ.get("Radarr_Movie_OriginalLanguage", "").lower()
    def fake_is_anime_sonarr(series_id, path=None):
        if os.environ.get("Sonarr_Series_Type", "").lower() == "anime":
            return True
        return bool(path and os.sep + "Anime" + os.sep in path + os.sep)
    def fake_is_anime_radarr(movie_id, path=None):
        return bool(path and os.sep + "Anime" + os.sep in path + os.sep)
    monkeypatch.setattr(f, "is_korean_sonarr_series", fake_is_korean_sonarr)
    monkeypatch.setattr(f, "is_korean_radarr_movie",  fake_is_korean_radarr)
    monkeypatch.setattr(f, "is_anime_sonarr_series",  fake_is_anime_sonarr)
    monkeypatch.setattr(f, "is_anime_radarr_movie",   fake_is_anime_radarr)
    return patches


@pytest.fixture(autouse=True)
def disable_external_tools(monkeypatch):
    """Don't actually call FolderIconCreator.exe during tests."""
    monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", False)


@pytest.fixture(autouse=True)
def clear_module_caches():
    """Korean/anime detection caches are module-level dicts. Reset each test."""
    f._korean_cache.clear()
    f._anime_cache.clear()
    yield
    f._korean_cache.clear()
    f._anime_cache.clear()


@pytest.fixture
def make_series(staging):
    """Factory for creating fake series/movie folders with placeholder files."""
    def _make(name, files=("S01E01.mkv",)):
        p = os.path.join(staging, name)
        os.makedirs(p, exist_ok=True)
        for fname in files:
            open(os.path.join(p, fname), "wb").write(b"fake\n")
        return p
    return _make
