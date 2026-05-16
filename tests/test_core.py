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


# ---------- Provider-error semantics (no silent fallback to IMDb) ----------

class TestProviderUnavailable:
    """Regression: when MAL/MDL has a transient outage (5xx, timeout) the
    dispatcher must propagate ProviderUnavailable instead of silently
    falling back to IMDb. Otherwise anime gets re-rated as IMDb every time
    Jikan goes down, triggering destructive folder renames — a real bug
    that rebadged 13 anime folders during a Jikan outage."""

    def test_mal_provider_unavailable_propagates(self, monkeypatch):
        def boom(title, year=None):
            raise f.ProviderUnavailable("Jikan returned 504")
        monkeypatch.setattr(f, "get_mal_rating", boom)
        with pytest.raises(f.ProviderUnavailable):
            f.get_rating_for_title("tt1", "Frieren", "2023", is_anime=True)

    def test_mdl_provider_unavailable_propagates(self, monkeypatch):
        def boom(title, year=None):
            raise f.ProviderUnavailable("All kuryana mirrors unavailable")
        monkeypatch.setattr(f, "get_mdl_rating", boom)
        with pytest.raises(f.ProviderUnavailable):
            f.get_rating_for_title("tt1", "Reverse", "2026", is_korean=True)

    def test_mal_none_still_falls_back_to_imdb(self, patch_providers):
        # None (legitimate no-match) — fallback is the desired behavior.
        patch_providers["mal_rating"] = None
        patch_providers["imdb_rating"] = "7.0"
        assert f.get_rating_for_title("tt1", "X", "2023", is_anime=True) == ("7.0", "IMDb")

    def test_process_sonarr_keeps_folder_on_provider_outage(self, staging, make_series, patch_providers, monkeypatch, clear_env_vars):
        """End-to-end: anime folder with [MAL X.X] suffix must NOT be renamed
        to [IMDb Y.Y] when Jikan is down."""
        folder = make_series("Some Anime (2024) [MAL 8.5]")
        def boom(title, year=None):
            raise f.ProviderUnavailable("Jikan returned 504")
        monkeypatch.setattr(f, "get_mal_rating", boom)
        patch_providers["imdb_rating"] = "7.0"   # would-be wrong fallback
        os.environ.update({
            "Sonarr_Series_Id": "1", "Sonarr_Series_ImdbId": "tt1",
            "Sonarr_Series_Title": "Some Anime (2024)", "Sonarr_Series_Year": "2024",
            "Sonarr_Series_Type": "anime",
        })
        f.process_sonarr(folder)
        # Folder must keep its MAL suffix; no IMDb-renamed twin must exist.
        assert os.path.isdir(folder), "Original [MAL] folder must remain"
        assert not os.path.isdir(folder.replace("[MAL 8.5]", "[IMDb 7.0]"))


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


# ---------- Parents guide multi-tab .vbs ----------

class TestParentsGuideVbs:
    def test_writes_vbs_with_three_sites(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("The Boys (2019)")
        patch_providers["imdb_rating"] = "8.6"
        os.environ.update({
            "Sonarr_Series_Id": "1", "Sonarr_Series_ImdbId": "tt1190634",
            "Sonarr_Series_Title": "The Boys (2019)", "Sonarr_Series_Year": "2019",
            "Sonarr_OriginalLanguage": "English",
        })
        f.process_sonarr(folder)
        vbs = os.path.join(folder + " [IMDb 8.6]", "Links", "Parents guide.vbs")
        lnk = os.path.join(folder + " [IMDb 8.6]", "Links", "Parents guide.lnk")
        assert os.path.isfile(vbs), "Parents guide.vbs must be created"
        assert os.path.isfile(lnk), "Parents guide.lnk must point at the .vbs"
        body = open(vbs, "r", encoding="utf-8").read()
        # All three sources must be present.
        assert "imdb.com/title/tt1190634/parentalguide" in body
        assert "commonsensemedia.org/search" in body
        assert "doesthedogdie.com/search" in body
        # Title is URL-encoded (year stripped).
        assert "query=The%20Boys" in body
        assert "text=The%20Boys" in body

    def test_lnk_target_is_the_vbs(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("The Boys (2019)")
        patch_providers["imdb_rating"] = "8.6"
        os.environ.update({
            "Sonarr_Series_Id": "1", "Sonarr_Series_ImdbId": "tt1190634",
            "Sonarr_Series_Title": "The Boys (2019)", "Sonarr_Series_Year": "2019",
            "Sonarr_OriginalLanguage": "English",
        })
        f.process_sonarr(folder)
        import win32com.client, pythoncom
        pythoncom.CoInitialize()
        try:
            lnk = win32com.client.Dispatch("WScript.Shell").CreateShortcut(
                os.path.join(folder + " [IMDb 8.6]", "Links", "Parents guide.lnk"))
            # The .lnk should target the .vbs, not a URL directly.
            assert lnk.TargetPath.lower().endswith("parents guide.vbs")
        finally:
            pythoncom.CoUninitialize()

    def test_old_lnk_replaced_when_vbs_first_created(self, staging, make_series, patch_providers, clear_env_vars, monkeypatch):
        # Simulate the migration path: a pre-existing Parents guide.lnk left
        # over from the single-link era exists, .vbs does NOT. After running,
        # the .lnk must be replaced with one pointing at the new .vbs.
        folder = make_series("The Boys (2019) [IMDb 8.6]")
        links_dir = os.path.join(folder, "Links")
        os.makedirs(links_dir, exist_ok=True)
        # Pre-create a stale .lnk (just an empty file; content doesn't matter
        # — the migration shim only checks existence).
        stale = os.path.join(links_dir, "Parents guide.lnk")
        open(stale, "wb").write(b"stale")
        stale_mtime = os.path.getmtime(stale)
        patch_providers["imdb_rating"] = "8.6"
        os.environ.update({
            "Sonarr_Series_Id": "1", "Sonarr_Series_ImdbId": "tt1190634",
            "Sonarr_Series_Title": "The Boys (2019)", "Sonarr_Series_Year": "2019",
            "Sonarr_OriginalLanguage": "English",
        })
        import time as _t; _t.sleep(0.01)
        f.process_sonarr(folder)
        # .vbs must be created.
        assert os.path.isfile(os.path.join(links_dir, "Parents guide.vbs"))
        # .lnk should have been regenerated (mtime advanced; non-stub content).
        new_lnk = os.path.join(links_dir, "Parents guide.lnk")
        assert os.path.isfile(new_lnk)
        assert os.path.getsize(new_lnk) > 4   # the stub was 5 bytes "stale"; real .lnk is larger


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


# ---------- IMDb GraphQL (primary rating source) ----------

class _FakeHttp:
    """Minimal stand-in for f.http() — captures the last POST and returns
    a canned response."""
    def __init__(self, status=200, body=None):
        self.last = None
        self._status = status
        self._body = body or {}
    def __call__(self):
        return self
    def post(self, url, headers=None, json=None, timeout=None):
        self.last = {"url": url, "headers": headers, "json": json}
        class _R:
            status_code = self._status
            _body = self._body
            def json(self_): return self_._body
        # Bind outer state into _R via closure
        outer = self
        class _Resp:
            status_code = outer._status
            def json(self_): return outer._body
        return _Resp()


class TestGraphQLRating:
    def test_returns_formatted_rating(self, monkeypatch):
        body = {"data": {"title": {"ratingsSummary": {"aggregateRating": 8.3}}}}
        monkeypatch.setattr(f, "http", _FakeHttp(status=200, body=body))
        assert f.get_imdb_rating_from_graphql("tt12042730") == "8.3"

    def test_handles_integer_rating(self, monkeypatch):
        # IMDb sometimes returns ints (e.g. 9 for a perfect score)
        body = {"data": {"title": {"ratingsSummary": {"aggregateRating": 9}}}}
        monkeypatch.setattr(f, "http", _FakeHttp(status=200, body=body))
        assert f.get_imdb_rating_from_graphql("tt0111161") == "9.0"

    def test_missing_aggregateRating_returns_na(self, monkeypatch):
        body = {"data": {"title": {"ratingsSummary": {}}}}
        monkeypatch.setattr(f, "http", _FakeHttp(status=200, body=body))
        assert f.get_imdb_rating_from_graphql("tt9999999") == "N/A"

    def test_missing_title_returns_na(self, monkeypatch):
        body = {"data": {"title": None}}
        monkeypatch.setattr(f, "http", _FakeHttp(status=200, body=body))
        assert f.get_imdb_rating_from_graphql("tt0000000") == "N/A"

    def test_http_error_returns_na(self, monkeypatch):
        monkeypatch.setattr(f, "http", _FakeHttp(status=500, body={}))
        assert f.get_imdb_rating_from_graphql("tt1") == "N/A"

    def test_empty_imdb_id_returns_na(self):
        assert f.get_imdb_rating_from_graphql("") == "N/A"
        assert f.get_imdb_rating_from_graphql(None) == "N/A"


class TestGetImdbRating:
    """get_imdb_rating: GraphQL first, OMDb fallback."""

    def test_prefers_graphql_over_omdb(self, monkeypatch):
        monkeypatch.setattr(f, "get_imdb_rating_from_graphql", lambda _id: "8.5")
        monkeypatch.setattr(f, "get_imdb_rating_from_omdb", lambda _id: "7.0")
        assert f.get_imdb_rating("tt1") == "8.5"

    def test_falls_back_to_omdb_when_graphql_misses(self, monkeypatch):
        monkeypatch.setattr(f, "get_imdb_rating_from_graphql", lambda _id: "N/A")
        monkeypatch.setattr(f, "get_imdb_rating_from_omdb", lambda _id: "7.0")
        assert f.get_imdb_rating("tt1") == "7.0"

    def test_both_miss_returns_na(self, monkeypatch):
        monkeypatch.setattr(f, "get_imdb_rating_from_graphql", lambda _id: "N/A")
        monkeypatch.setattr(f, "get_imdb_rating_from_omdb", lambda _id: "N/A")
        assert f.get_imdb_rating("tt1") == "N/A"

    def test_empty_imdb_id_returns_na(self):
        assert f.get_imdb_rating("") == "N/A"


# ---------- _load_env_file (matched-quote stripping) ----------

class TestEnvLoader:
    def _load(self, tmp_path, monkeypatch, lines):
        # Scrub the keys we're about to set so the "real env wins" guard
        # doesn't no-op our test data.
        for line in lines:
            if "=" in line and not line.strip().startswith("#"):
                k = line.split("=", 1)[0].strip()
                monkeypatch.delenv(k, raising=False)
        path = tmp_path / ".env"
        path.write_text("\n".join(lines), encoding="utf-8")
        f._load_env_file(str(path))

    def test_strips_matching_double_quotes(self, tmp_path, monkeypatch):
        self._load(tmp_path, monkeypatch, ['ARRTEST_DQ="hello"'])
        assert os.environ["ARRTEST_DQ"] == "hello"

    def test_strips_matching_single_quotes(self, tmp_path, monkeypatch):
        self._load(tmp_path, monkeypatch, ["ARRTEST_SQ='hello'"])
        assert os.environ["ARRTEST_SQ"] == "hello"

    def test_preserves_mismatched_quotes(self, tmp_path, monkeypatch):
        # KEY="a'b" used to become a'b under the naive strip; should stay a'b
        # (i.e. only the outer quotes are stripped, the inner ' is preserved).
        self._load(tmp_path, monkeypatch, ["""ARRTEST_MIX="a'b" """])
        assert os.environ["ARRTEST_MIX"] == "a'b"

    def test_no_quotes_passes_through(self, tmp_path, monkeypatch):
        self._load(tmp_path, monkeypatch, ["ARRTEST_NQ=plain"])
        assert os.environ["ARRTEST_NQ"] == "plain"

    def test_real_env_wins(self, tmp_path, monkeypatch):
        monkeypatch.setenv("ARRTEST_PRECEDENCE", "from-real-env")
        path = tmp_path / ".env"
        path.write_text("ARRTEST_PRECEDENCE=from-file\n", encoding="utf-8")
        f._load_env_file(str(path))
        assert os.environ["ARRTEST_PRECEDENCE"] == "from-real-env"

    def test_comment_and_blank_ignored(self, tmp_path, monkeypatch):
        monkeypatch.delenv("ARRTEST_AFTER_BLANK", raising=False)
        self._load(tmp_path, monkeypatch, [
            "# comment",
            "",
            "ARRTEST_AFTER_BLANK=ok",
        ])
        assert os.environ["ARRTEST_AFTER_BLANK"] == "ok"


# ---------- set_folder_tooltip desktop.ini parser/rewriter ----------

class TestTooltipWriter:
    def test_creates_new_desktop_ini(self, tmp_path, monkeypatch):
        monkeypatch.setattr(f, "ENABLE_SET_TOOLTIP", True)
        folder = tmp_path / "show"
        folder.mkdir()
        f.set_folder_tooltip(str(folder), "Plot summary  [IMDb 8.0]")
        ini = folder / "desktop.ini"
        assert ini.exists()
        content = ini.read_bytes().decode("utf-16")
        assert "[.ShellClassInfo]" in content
        assert "InfoTip=Plot summary  [IMDb 8.0]" in content

    def test_replaces_existing_infotip(self, tmp_path, monkeypatch):
        monkeypatch.setattr(f, "ENABLE_SET_TOOLTIP", True)
        folder = tmp_path / "show"
        folder.mkdir()
        ini = folder / "desktop.ini"
        ini.write_bytes(
            ("[.ShellClassInfo]\r\nInfoTip=OLD\r\nIconResource=poster.ico,0\r\n"
             ).encode("utf-16")
        )
        f.set_folder_tooltip(str(folder), "NEW tooltip")
        text = ini.read_bytes().decode("utf-16")
        assert "InfoTip=NEW tooltip" in text
        assert "InfoTip=OLD" not in text
        # Existing IconResource line must be preserved.
        assert "IconResource=poster.ico,0" in text

    def test_idempotent_when_tooltip_unchanged(self, tmp_path, monkeypatch):
        monkeypatch.setattr(f, "ENABLE_SET_TOOLTIP", True)
        folder = tmp_path / "show"
        folder.mkdir()
        f.set_folder_tooltip(str(folder), "Same tooltip")
        ini = folder / "desktop.ini"
        first_mtime = ini.stat().st_mtime_ns
        # Force a different timestamp by waiting 1 ms, then re-set with the
        # same value. The file should not be rewritten.
        import time as _t
        _t.sleep(0.01)
        f.set_folder_tooltip(str(folder), "Same tooltip")
        assert ini.stat().st_mtime_ns == first_mtime

    def test_disabled_toggle_short_circuits(self, tmp_path, monkeypatch):
        # Carried-over coverage from TestTooltipToggle, but consolidated here.
        monkeypatch.setattr(f, "ENABLE_SET_TOOLTIP", False)
        folder = tmp_path / "show"
        folder.mkdir()
        f.set_folder_tooltip(str(folder), "ignored")
        assert not (folder / "desktop.ini").exists()


# ---------- create_folder_icon: only skip when icon already declared ----------

class TestFolderIconSkip:
    def test_tooltip_only_desktop_ini_does_not_block_icon(self, tmp_path, monkeypatch):
        # Simulate the scenario: an earlier run wrote a tooltip (so
        # desktop.ini exists) but never wrote an IconResource line. A later
        # run with ENABLE_CREATE_FOLDER_ICON=True must NOT short-circuit.
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "desktop.ini").write_bytes(
            "[.ShellClassInfo]\r\nInfoTip=tooltip without icon\r\n".encode("utf-16")
        )
        called = {"n": 0}
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "FOLDER_ICON_EXE", "stub.exe")
        monkeypatch.setattr(f.subprocess, "run", lambda *a, **kw: called.__setitem__("n", called["n"] + 1))
        f.create_folder_icon(str(folder))
        assert called["n"] == 1   # Creator.exe was invoked

    def test_desktop_ini_with_icon_resource_skips(self, tmp_path, monkeypatch):
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "desktop.ini").write_bytes(
            "[.ShellClassInfo]\r\nIconResource=poster.ico,0\r\n".encode("utf-16")
        )
        called = {"n": 0}
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "FOLDER_ICON_EXE", "stub.exe")
        monkeypatch.setattr(f.subprocess, "run", lambda *a, **kw: called.__setitem__("n", called["n"] + 1))
        f.create_folder_icon(str(folder))
        assert called["n"] == 0   # short-circuited


# ---------- _save_rating_cache dirty flag ----------

class TestRatingCacheDirtyFlag:
    def test_no_write_when_clean(self, monkeypatch, tmp_path):
        cache_path = tmp_path / ".rating_cache.json"
        cache_path.write_text('{"tt1":{"checked_at":1,"rating":"8.0","source":"IMDb"}}')
        before_mtime = cache_path.stat().st_mtime_ns
        monkeypatch.setattr(f, "RATING_CACHE_PATH", str(cache_path))
        monkeypatch.setattr(f, "_rating_cache", None)
        monkeypatch.setattr(f, "_rating_cache_dirty", False)
        # Force a lazy load (mark cache as loaded but not dirty)
        f._load_rating_cache()
        f._save_rating_cache()
        # File should be untouched — no write happened.
        assert cache_path.stat().st_mtime_ns == before_mtime

    def test_writes_after_set(self, monkeypatch, tmp_path):
        cache_path = tmp_path / ".rating_cache.json"
        monkeypatch.setattr(f, "RATING_CACHE_PATH", str(cache_path))
        monkeypatch.setattr(f, "_rating_cache", {})
        monkeypatch.setattr(f, "_rating_cache_dirty", False)
        f._rating_cache_set("tt99", "8.4", "IMDb")
        f._save_rating_cache()
        assert cache_path.exists()
        import json as _json
        loaded = _json.loads(cache_path.read_text())
        assert "tt99" in loaded
