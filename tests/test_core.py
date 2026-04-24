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
