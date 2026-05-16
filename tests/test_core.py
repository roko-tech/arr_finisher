"""Core behavior tests for arr_finisher."""
import os, threading, time
import pytest
import arr_finisher as f


# ---------- Rating dispatch ----------

class TestRatingDispatch:
    def test_korean_uses_mdl(self, patch_providers):
        patch_providers["mdl_rating"] = ("7.5", "MDL")
        patch_providers["imdb_rating"] = "7.2"
        assert f.get_rating_for_title("tt1", "X", "2025", is_korean=True) == ("7.5", "MDL")

    def test_anime_uses_mal(self, patch_providers):
        patch_providers["mal_rating"] = ("9.3", "MAL")
        patch_providers["imdb_rating"] = "8.0"
        assert f.get_rating_for_title("tt1", "X", "2023", is_anime=True) == ("9.3", "MAL")

    def test_non_korean_non_anime_uses_imdb(self, patch_providers):
        patch_providers["imdb_rating"] = "8.6"
        assert f.get_rating_for_title("tt1", "X", "2019") == ("8.6", "IMDb")

    def test_mdl_fallback_to_imdb_when_none(self, patch_providers):
        patch_providers["mdl_rating"] = None
        patch_providers["imdb_rating"] = "7.0"
        assert f.get_rating_for_title("tt1", "X", "2025", is_korean=True) == ("7.0", "IMDb")

    def test_mal_fallback_to_imdb_when_none(self, patch_providers):
        patch_providers["mal_rating"] = None
        patch_providers["imdb_rating"] = "6.5"
        assert f.get_rating_for_title("tt1", "X", "2023", is_anime=True) == ("6.5", "IMDb")

    def test_anime_precedence_over_korean(self, patch_providers):
        patch_providers["mal_rating"] = ("9.0", "MAL")
        patch_providers["mdl_rating"] = ("7.5", "MDL")
        # When both flags are set, anime should win
        assert f.get_rating_for_title("tt1", "X", "2023", is_korean=True, is_anime=True) == ("9.0", "MAL")


# ---------- Suffix handling ----------

class TestSuffixStripping:
    def test_strips_imdb(self):
        assert f._strip_rating_suffix("The Boys [IMDb 8.6]") == "The Boys"

    def test_strips_mdl(self):
        assert f._strip_rating_suffix("Reverse (2026) [MDL 7.5]") == "Reverse (2026)"

    def test_strips_mal(self):
        assert f._strip_rating_suffix("Frieren (2023) [MAL 9.3]") == "Frieren (2023)"

    def test_leaves_unrelated_brackets(self):
        assert f._strip_rating_suffix("Show [Season 1]") == "Show [Season 1]"

    def test_empty_input(self):
        assert f._strip_rating_suffix("") == ""


# ---------- Language detection env var short-circuit ----------

class TestLanguageDetection:
    def test_sonarr_korean_env(self, clear_env_vars):
        os.environ["Sonarr_OriginalLanguage"] = "Korean"
        assert f.is_korean_sonarr_series(None) is True

    def test_sonarr_not_korean_env(self, clear_env_vars):
        os.environ["Sonarr_OriginalLanguage"] = "English"
        assert f.is_korean_sonarr_series(None) is False

    def test_radarr_korean_env(self, clear_env_vars):
        os.environ["Radarr_Movie_OriginalLanguage"] = "Korean"
        assert f.is_korean_radarr_movie(None) is True

    def test_sonarr_anime_from_env(self, clear_env_vars):
        os.environ["Sonarr_Series_Type"] = "anime"
        assert f.is_anime_sonarr_series(None) is True

    def test_sonarr_anime_from_path(self, clear_env_vars):
        path = r"D:\Anime\Some Show"
        assert f.is_anime_sonarr_series(None, path=path) is True

    def test_sonarr_nonanime(self, clear_env_vars):
        os.environ["Sonarr_Series_Type"] = "standard"
        assert f.is_anime_sonarr_series(None, path=r"D:\TV Shows\Show") is False


# ---------- Kuryana match-quality filter (#10) ----------

class TestMatchQuality:
    def test_title_similarity(self):
        assert f._title_similarity("Reverse", "Reverse") == 1.0
        assert f._title_similarity("Reverse", "Reverses") > 0.85
        assert f._title_similarity("Reverse", "Something totally different") < 0.5


# ---------- Path matching (#2) ----------

class TestPathMatching:
    def test_normalization_equal(self):
        n = lambda p: os.path.normpath(os.path.normcase(p.replace("/", "\\"))).rstrip("\\")
        assert n("D:/TV Shows/The Boys") == n("D:\\tv shows\\THE BOYS")

    def test_normalization_different_shows(self):
        n = lambda p: os.path.normpath(os.path.normcase(p.replace("/", "\\"))).rstrip("\\")
        assert n("D:\\TV\\The Boys") != n("D:\\TV\\The Boys of Summer")

    def test_trailing_slash_ignored(self):
        n = lambda p: os.path.normpath(os.path.normcase(p.replace("/", "\\"))).rstrip("\\")
        assert n("D:\\TV\\Foo\\") == n("D:\\TV\\Foo")


# ---------- End-to-end Sonarr flow ----------

class TestSonarrProcess:
    def test_korean_series_renamed_with_mdl(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("Reverse (2026)")
        patch_providers["mdl_rating"] = ("7.5", "MDL")
        os.environ.update({
            "Sonarr_Series_Id": "1", "Sonarr_Series_ImdbId": "tt1",
            "Sonarr_Series_Title": "Reverse (2026)", "Sonarr_Series_Year": "2026",
            "Sonarr_OriginalLanguage": "Korean",
        })
        f.process_sonarr(folder)
        assert os.path.isdir(folder + " [MDL 7.5]")
        assert os.path.isfile(os.path.join(folder + " [MDL 7.5]", "Links", "MyDramaList.lnk"))

    def test_anime_series_gets_mal_rating(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("Frieren (2023)")
        patch_providers["mal_rating"] = ("9.3", "MAL")
        os.environ.update({
            "Sonarr_Series_Id": "2", "Sonarr_Series_ImdbId": "tt2",
            "Sonarr_Series_Title": "Frieren (2023)", "Sonarr_Series_Year": "2023",
            "Sonarr_Series_Type": "anime",
        })
        f.process_sonarr(folder)
        assert os.path.isdir(folder + " [MAL 9.3]")
        assert os.path.isfile(os.path.join(folder + " [MAL 9.3]", "Links", "MyAnimeList.lnk"))

    def test_non_korean_non_anime_uses_imdb(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("The Boys (2019)")
        patch_providers["imdb_rating"] = "8.6"
        os.environ.update({
            "Sonarr_Series_Id": "3", "Sonarr_Series_ImdbId": "tt3",
            "Sonarr_Series_Title": "The Boys (2019)", "Sonarr_Series_Year": "2019",
            "Sonarr_OriginalLanguage": "English",
        })
        f.process_sonarr(folder)
        assert os.path.isdir(folder + " [IMDb 8.6]")
        assert not os.path.isfile(os.path.join(folder + " [IMDb 8.6]", "Links", "MyDramaList.lnk"))
        assert not os.path.isfile(os.path.join(folder + " [IMDb 8.6]", "Links", "MyAnimeList.lnk"))

    def test_rating_na_skips_rename(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("Mystery Show (2024)")
        patch_providers["imdb_rating"] = "N/A"
        os.environ.update({
            "Sonarr_Series_Id": "4", "Sonarr_Series_Title": "Mystery Show (2024)",
            "Sonarr_Series_Year": "2024", "Sonarr_OriginalLanguage": "English",
        })
        f.process_sonarr(folder)
        assert os.path.isdir(folder)  # unchanged
        assert not any(d.endswith("N/A]") for d in os.listdir(staging))

    def test_missing_folder_doesnt_create_phantom(self, staging, patch_providers, clear_env_vars):
        bogus = os.path.join(staging, "Ghost Show")
        patch_providers["imdb_rating"] = "7.0"
        os.environ.update({
            "Sonarr_Series_Id": "5", "Sonarr_Series_ImdbId": "tt5",
            "Sonarr_Series_Title": "Ghost Show", "Sonarr_OriginalLanguage": "English",
        })
        f.process_sonarr(bogus)
        assert not os.path.isdir(bogus)

    def test_rollback_on_api_failure(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("Rollback Test (2025)")
        patch_providers["mdl_rating"] = ("7.5", "MDL")
        patch_providers["sonarr_put_ok"] = False
        os.environ.update({
            "Sonarr_Series_Id": "6", "Sonarr_Series_ImdbId": "tt6",
            "Sonarr_Series_Title": "Rollback Test (2025)", "Sonarr_Series_Year": "2025",
            "Sonarr_OriginalLanguage": "Korean",
        })
        f.process_sonarr(folder)
        assert os.path.isdir(folder)
        assert not os.path.isdir(folder + " [MDL 7.5]")

    def test_provider_migration_imdb_to_mdl(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("Reverse (2026) [IMDb 7.2]")
        patch_providers["mdl_rating"] = ("7.5", "MDL")
        os.environ.update({
            "Sonarr_Series_Id": "7", "Sonarr_Series_ImdbId": "tt7",
            "Sonarr_Series_Title": "Reverse (2026)", "Sonarr_Series_Year": "2026",
            "Sonarr_OriginalLanguage": "Korean",
        })
        f.process_sonarr(folder)
        assert not os.path.isdir(folder)
        assert os.path.isdir(os.path.join(staging, "Reverse (2026) [MDL 7.5]"))


# ---------- Direct-link shortcuts (MDL + MAL use show page, not search) ----------

class TestDirectLinks:
    def test_mdl_shortcut_uses_direct_url_when_cached(self, staging, make_series, patch_providers, monkeypatch):
        folder = make_series("Reverse (2026)")
        patch_providers["mdl_rating"] = ("7.5", "MDL")
        # Simulate get_mdl_rating having cached a direct URL
        f._mdl_url_cache[(("reverse"), "2026")] = "https://mydramalist.com/773635-reverse"
        os.environ.update({
            "Sonarr_Series_Id": "1", "Sonarr_Series_ImdbId": "tt1",
            "Sonarr_Series_Title": "Reverse (2026)", "Sonarr_Series_Year": "2026",
            "Sonarr_OriginalLanguage": "Korean",
        })
        f.process_sonarr(folder)

        # Read the MyDramaList.lnk back and assert its args reference the direct URL
        import win32com.client, pythoncom
        pythoncom.CoInitialize()
        try:
            lnk = win32com.client.Dispatch("WScript.Shell").CreateShortcut(
                os.path.join(folder + " [MDL 7.5]", "Links", "MyDramaList.lnk"))
            assert "773635-reverse" in lnk.Arguments
            assert "search" not in lnk.Arguments.lower()
        finally:
            pythoncom.CoUninitialize()
            f._mdl_url_cache.clear()

    def test_mal_shortcut_uses_direct_url_when_cached(self, staging, make_series, patch_providers):
        folder = make_series("Frieren (2023)")
        patch_providers["mal_rating"] = ("9.3", "MAL")
        f._mal_url_cache[(("frieren"), "2023")] = "https://myanimelist.net/anime/52991/Sousou_no_Frieren"
        os.environ.update({
            "Sonarr_Series_Id": "2", "Sonarr_Series_ImdbId": "tt2",
            "Sonarr_Series_Title": "Frieren (2023)", "Sonarr_Series_Year": "2023",
            "Sonarr_Series_Type": "anime",
        })
        f.process_sonarr(folder)

        import win32com.client, pythoncom
        pythoncom.CoInitialize()
        try:
            lnk = win32com.client.Dispatch("WScript.Shell").CreateShortcut(
                os.path.join(folder + " [MAL 9.3]", "Links", "MyAnimeList.lnk"))
            assert "/anime/52991/" in lnk.Arguments
            assert "search" not in lnk.Arguments.lower()
        finally:
            pythoncom.CoUninitialize()
            f._mal_url_cache.clear()

    def test_falls_back_to_search_when_uncached(self, staging, make_series, patch_providers):
        folder = make_series("Unknown Korean (2025)")
        patch_providers["mdl_rating"] = ("7.0", "MDL")
        # Deliberately no cache entry
        f._mdl_url_cache.clear()
        os.environ.update({
            "Sonarr_Series_Id": "3", "Sonarr_Series_ImdbId": "tt3",
            "Sonarr_Series_Title": "Unknown Korean (2025)", "Sonarr_Series_Year": "2025",
            "Sonarr_OriginalLanguage": "Korean",
        })
        f.process_sonarr(folder)

        import win32com.client, pythoncom
        pythoncom.CoInitialize()
        try:
            lnk = win32com.client.Dispatch("WScript.Shell").CreateShortcut(
                os.path.join(folder + " [MDL 7.0]", "Links", "MyDramaList.lnk"))
            # Falls back to search URL — should include 'search?'
            assert "search?" in lnk.Arguments
        finally:
            pythoncom.CoUninitialize()


# ---------- Dry-run mode (#6) ----------

class TestDryRun:
    def test_dry_run_doesnt_rename(self, staging, make_series, patch_providers, clear_env_vars, monkeypatch):
        folder = make_series("Reverse (2026)")
        monkeypatch.setattr(f, "DRY_RUN", True)
        patch_providers["mdl_rating"] = ("7.5", "MDL")
        os.environ.update({
            "Sonarr_Series_Id": "1", "Sonarr_Series_ImdbId": "tt1",
            "Sonarr_Series_Title": "Reverse (2026)", "Sonarr_Series_Year": "2026",
            "Sonarr_OriginalLanguage": "Korean",
        })
        f.process_sonarr(folder)
        assert os.path.isdir(folder)                      # original still there
        assert not os.path.isdir(folder + " [MDL 7.5]")   # renamed path NOT created


# ---------- --roots CLI parsing (regression for the rsplit fix) ----------

class TestRootsParsing:
    def test_windows_drive_letter_in_path(self):
        # 'D:\TV Shows:sonarr' has TWO colons — must rsplit so service is `sonarr`.
        result = f._parse_roots_arg([r"D:\TV Shows:sonarr"])
        assert result == [(r"D:\TV Shows", "sonarr")]

    def test_multiple_entries(self):
        result = f._parse_roots_arg([r"D:\TV:sonarr", r"E:\Movies:radarr"])
        assert result == [(r"D:\TV", "sonarr"), (r"E:\Movies", "radarr")]

    def test_unknown_service_rejected(self):
        with pytest.raises(ValueError):
            f._parse_roots_arg([r"D:\Foo:notaservice"])

    def test_missing_colon_rejected(self):
        with pytest.raises(ValueError):
            f._parse_roots_arg([r"D:\Foo"])


# ---------- rename_folder merge: must NOT delete conflicting items ----------

class TestRenameMerge:
    def test_merge_with_conflict_leaves_source(self, staging):
        old = os.path.join(staging, "Show (2024)")
        new = os.path.join(staging, "Show (2024) [IMDb 8.0]")
        os.makedirs(old); os.makedirs(new)
        # Same filename exists in BOTH — must not be silently destroyed.
        open(os.path.join(old, "ep1.mkv"), "wb").write(b"old-content-keep-me")
        open(os.path.join(new, "ep1.mkv"), "wb").write(b"new-content")
        # And a non-conflicting file that should successfully move.
        open(os.path.join(old, "ep2.mkv"), "wb").write(b"unique")

        f.rename_folder(old, "8.0", "IMDb")

        # The conflicting file in the SOURCE must still exist (not deleted).
        assert os.path.isfile(os.path.join(old, "ep1.mkv"))
        # The unique file moved successfully.
        assert os.path.isfile(os.path.join(new, "ep2.mkv"))
        # The destination still has its original ep1.mkv (we didn't overwrite).
        with open(os.path.join(new, "ep1.mkv"), "rb") as fh:
            assert fh.read() == b"new-content"

    def test_merge_clean_removes_source(self, staging):
        old = os.path.join(staging, "Show (2024)")
        new = os.path.join(staging, "Show (2024) [IMDb 8.0]")
        os.makedirs(old); os.makedirs(new)
        open(os.path.join(old, "ep1.mkv"), "wb").write(b"moves-cleanly")

        f.rename_folder(old, "8.0", "IMDb")

        assert not os.path.isdir(old)
        assert os.path.isfile(os.path.join(new, "ep1.mkv"))


# ---------- URL safety filter (Subtitle.vbs injection guard) ----------

class TestUrlSafety:
    def test_accepts_plain_https(self):
        assert f._is_safe_url("https://www.opensubtitles.com/en/movies/2024-foo")

    def test_rejects_double_quote(self):
        assert not f._is_safe_url('https://evil.example/" Set fso = ...')

    def test_rejects_newline(self):
        assert not f._is_safe_url("https://x.example/a\nb")

    def test_rejects_non_http(self):
        assert not f._is_safe_url("javascript:alert(1)")
        assert not f._is_safe_url("file:///c:/secret")

    def test_rejects_empty(self):
        assert not f._is_safe_url("")
        assert not f._is_safe_url(None)


# ---------- Log redaction ----------

class TestRedaction:
    def test_redacts_omdb_key(self, monkeypatch):
        monkeypatch.setattr(f, "OMDB_API_KEY", "SECRET_KEY_12345")
        msg = "GET https://www.omdbapi.com/?apikey=SECRET_KEY_12345&i=tt1 failed"
        assert "SECRET_KEY_12345" not in f._redact(msg)
        assert "<redacted>" in f._redact(msg)

    def test_pass_through_when_no_secret(self, monkeypatch):
        monkeypatch.setattr(f, "OMDB_API_KEY", "")
        assert f._redact("nothing to redact") == "nothing to redact"


# ---------- Sweep mode: cache TTL, env stuffing ----------

class TestSweep:
    def test_sweep_one_stuffs_env_and_processes(self, staging, make_series,
                                                 patch_providers, monkeypatch, clear_env_vars):
        folder = make_series("The Boys (2019)")
        patch_providers["imdb_rating"] = "8.6"
        fake_obj = {"id": 42, "imdbId": "tt1190634", "tvdbId": 355567,
                    "title": "The Boys", "year": 2019,
                    "originalLanguage": {"name": "English"}, "seriesType": "standard"}
        monkeypatch.setattr(f, "get_object_by_path", lambda svc, p: fake_obj)
        assert f._sweep_one("sonarr", folder) is True
        # Sweep doesn't create shortcuts/icons/tooltips — only renames.
        assert os.path.isdir(folder + " [IMDb 8.6]")
        # Cache should now have an entry for this IMDb ID.
        assert "tt1190634" in f._load_rating_cache()

    def test_sweep_one_returns_false_for_unknown(self, staging, make_series, monkeypatch):
        folder = make_series("Unknown Show (2024)")
        monkeypatch.setattr(f, "get_object_by_path", lambda svc, p: None)
        assert f._sweep_one("sonarr", folder) is False

    def test_rating_cache_freshness_skips(self, monkeypatch):
        # Force a cache entry that's brand new.
        monkeypatch.setattr(f, "_rating_cache", {"tt9999": {"checked_at": __import__("time").time(),
                                                             "rating": "9.0", "source": "IMDb"}})
        assert f._rating_cache_is_fresh("tt9999") is True

    def test_rating_cache_stale_does_not_skip(self, monkeypatch):
        import time as _t
        monkeypatch.setattr(f, "_rating_cache", {"tt9999": {"checked_at": _t.time() - 999 * 86400,
                                                             "rating": "9.0", "source": "IMDb"}})
        assert f._rating_cache_is_fresh("tt9999") is False


# ---------- --clear-cache / --refresh ----------

class TestCacheCommands:
    def test_clear_cache_deletes_file(self, monkeypatch, tmp_path):
        cache_file = tmp_path / ".rating_cache.json"
        cache_file.write_text('{"tt1":{"checked_at":1,"rating":"8.0","source":"IMDb"}}')
        monkeypatch.setattr(f, "RATING_CACHE_PATH", str(cache_file))
        monkeypatch.setattr(f, "_rating_cache", None)
        assert f.clear_rating_cache() == 0
        assert not cache_file.exists()

    def test_refresh_removes_single_entry(self, monkeypatch, tmp_path):
        cache_file = tmp_path / ".rating_cache.json"
        cache_file.write_text(
            '{"tt1":{"checked_at":1,"rating":"8.0","source":"IMDb"},'
            ' "tt2":{"checked_at":2,"rating":"7.0","source":"IMDb"}}'
        )
        monkeypatch.setattr(f, "RATING_CACHE_PATH", str(cache_file))
        monkeypatch.setattr(f, "_rating_cache", None)
        assert f.clear_rating_cache("tt1") == 0
        # The cache file should still exist with tt2 intact.
        assert cache_file.exists()
        import json as _json
        remaining = _json.loads(cache_file.read_text())
        assert "tt1" not in remaining
        assert "tt2" in remaining


# ---------- Tooltip toggle (ENABLE_SET_TOOLTIP) ----------

class TestTooltipToggle:
    def test_disabled_toggle_short_circuits(self, staging, make_series, monkeypatch):
        folder = make_series("Anything (2024)")
        monkeypatch.setattr(f, "ENABLE_SET_TOOLTIP", False)
        f.set_folder_tooltip(folder, "Some plot")
        # Nothing was written.
        assert not os.path.exists(os.path.join(folder, "desktop.ini"))


# ---------- OMDb response caching (S6) ----------

class TestOmdbCaching:
    def test_single_fetch_serves_both_callers(self, monkeypatch):
        calls = {"n": 0}
        def fake_get(url, timeout=15):
            calls["n"] += 1
            class _R:
                status_code = 200
                def json(self_): return {"imdbRating": "8.6", "Plot": "A pithy summary."}
            return _R()
        monkeypatch.setattr(f, "OMDB_API_KEY", "fake")
        monkeypatch.setattr(f, "_omdb_response_cache", {})
        monkeypatch.setattr(f, "http", lambda: type("S", (), {"get": staticmethod(fake_get)})())
        assert f.get_imdb_rating_from_omdb("tt1234567") == "8.6"
        assert f.get_omdb_plot("tt1234567") == "A pithy summary."
        assert calls["n"] == 1   # cached


# ---------- Version + CHANGELOG ----------

class TestVersion:
    def test_version_is_string(self):
        assert isinstance(f.__version__, str)
        # Sanity check it looks like semver (e.g., "1.0.0")
        parts = f.__version__.split(".")
        assert len(parts) >= 2
        assert all(p.isdigit() for p in parts[:2])


# ---------- ISO timestamps (back-compat: read float OR string) ----------

class TestRatingCacheTimestamps:
    def test_write_uses_iso_string(self, monkeypatch):
        cache = {}
        monkeypatch.setattr(f, "_rating_cache", cache)
        f._rating_cache_set("tt100", "8.0", "IMDb")
        # Stored as an ISO-8601 string (parseable by datetime.fromisoformat).
        ts = cache["tt100"]["checked_at"]
        assert isinstance(ts, str)
        from datetime import datetime as _dt
        _dt.fromisoformat(ts)   # raises if not parseable

    def test_legacy_float_timestamp_still_reads_as_fresh(self, monkeypatch):
        import time as _t
        monkeypatch.setattr(f, "_rating_cache", {"tt200": {
            "checked_at": _t.time(),   # legacy float epoch
            "rating": "8.0", "source": "IMDb"
        }})
        assert f._rating_cache_is_fresh("tt200") is True

    def test_iso_string_reads_as_fresh(self, monkeypatch):
        from datetime import datetime as _dt
        monkeypatch.setattr(f, "_rating_cache", {"tt300": {
            "checked_at": _dt.now().isoformat(timespec="seconds"),
            "rating": "8.0", "source": "IMDb"
        }})
        assert f._rating_cache_is_fresh("tt300") is True

    def test_iso_string_old_reads_as_stale(self, monkeypatch):
        # 100 days ago — past any reasonable TTL.
        from datetime import datetime as _dt, timedelta
        old = (_dt.now() - timedelta(days=100)).isoformat(timespec="seconds")
        monkeypatch.setattr(f, "_rating_cache", {"tt400": {
            "checked_at": old, "rating": "8.0", "source": "IMDb"
        }})
        assert f._rating_cache_is_fresh("tt400") is False

    def test_garbage_timestamp_treated_as_stale(self, monkeypatch):
        monkeypatch.setattr(f, "_rating_cache", {"tt500": {
            "checked_at": "not-a-date", "rating": "8.0", "source": "IMDb"
        }})
        assert f._rating_cache_is_fresh("tt500") is False


# ---------- Auto-discovery of sweep roots ----------

class TestRootDiscovery:
    def test_discover_returns_paths_from_both_services(self, monkeypatch):
        def fake_http():
            class _S:
                @staticmethod
                def get(url, headers=None, timeout=5):
                    class _R:
                        status_code = 200
                        def json(self_):
                            if "/movie" in url or "7878" in url or "radarr" in url.lower():
                                return [{"path": r"E:\Movies"}]
                            return [{"path": r"D:\TV Shows"}, {"path": r"D:\Anime"}]
                    return _R()
            return _S()
        monkeypatch.setattr(f, "SONARR_API_KEY", "fakesonarr")
        monkeypatch.setattr(f, "RADARR_API_KEY", "fakeradarr")
        monkeypatch.setattr(f, "SONARR_API_URL", "http://localhost:8989")
        monkeypatch.setattr(f, "RADARR_API_URL", "http://localhost:7878")
        monkeypatch.setattr(f, "http", fake_http)
        roots = f._discover_sweep_roots()
        # Should have entries from both Sonarr and Radarr.
        services = {svc for _, svc in roots}
        assert "sonarr" in services
        assert "radarr" in services

    def test_discover_skips_unconfigured_service(self, monkeypatch):
        monkeypatch.setattr(f, "SONARR_API_KEY", "")     # not configured
        monkeypatch.setattr(f, "RADARR_API_KEY", "")     # not configured
        roots = f._discover_sweep_roots()
        assert roots == []

    def test_discover_swallows_errors(self, monkeypatch):
        def fake_http():
            class _S:
                @staticmethod
                def get(url, headers=None, timeout=5):
                    raise ConnectionError("nope")
            return _S()
        monkeypatch.setattr(f, "SONARR_API_KEY", "x")
        monkeypatch.setattr(f, "RADARR_API_KEY", "x")
        monkeypatch.setattr(f, "http", fake_http)
        assert f._discover_sweep_roots() == []   # no crash


# ---------- Setup-help sidecar ----------

class TestSetupHelp:
    def test_missing_omdb_writes_file(self, monkeypatch, tmp_path):
        setup_path = tmp_path / "arr_finisher_setup.txt"
        monkeypatch.setattr(f, "_SETUP_HELP_PATH", str(setup_path))
        monkeypatch.setattr(f, "OMDB_API_KEY", "")
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", False)
        missing = f._check_critical_config()
        assert any(name == "OMDB_API_KEY" for name, _ in missing)
        f._update_setup_help(missing)
        assert setup_path.exists()
        content = setup_path.read_text(encoding="utf-8")
        assert "OMDB_API_KEY" in content

    def test_healthy_config_clears_stale_file(self, monkeypatch, tmp_path):
        setup_path = tmp_path / "arr_finisher_setup.txt"
        setup_path.write_text("stale notice")
        monkeypatch.setattr(f, "_SETUP_HELP_PATH", str(setup_path))
        # No missing config — calling update with [] should delete the file.
        f._update_setup_help([])
        assert not setup_path.exists()

    def test_sonarr_event_requires_sonarr_key(self, monkeypatch):
        monkeypatch.setattr(f, "OMDB_API_KEY", "set")
        monkeypatch.setattr(f, "SONARR_API_KEY", "")
        monkeypatch.setattr(f, "RADARR_API_KEY", "set")
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", False)
        # Sonarr webhook -> SONARR_API_KEY required, RADARR_API_KEY not.
        missing = f._check_critical_config(sonarr_event="download")
        names = [n for n, _ in missing]
        assert "SONARR_API_KEY" in names
        assert "RADARR_API_KEY" not in names
