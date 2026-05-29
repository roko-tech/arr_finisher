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


# ---------- Env-vs-API tiebreaker (no silent MDL/MAL → IMDb demotion) ----------

class TestEnvVsApiTiebreaker:
    """Regression: when Sonarr/Radarr sends a stale OriginalLanguage env var
    (mid-metadata-refresh), the fast env-var path used to silently demote a
    [MDL X.X] folder to [IMDb Y.Y]. The fix consults the service API as a
    tiebreaker when the existing folder suffix contradicts the env-var
    detection — same idea for [MAL X.X] and anime."""

    def test_mdl_suffix_preserved_when_api_says_korean(self, staging, make_series, patch_providers, clear_env_vars):
        # Folder is already labeled [MDL 7.7]; env var lies (says English),
        # API tells the truth (Korean). Defensive check must consult API
        # and keep the show on MDL — NOT demote to IMDb.
        folder = make_series("Gold Land (2026) [MDL 7.7]")
        patch_providers["mdl_rating"]      = ("7.6", "MDL")
        patch_providers["imdb_rating"]     = "7.7"      # would be the wrong demotion target
        patch_providers["api_says_korean"] = True       # ← API tiebreaker says Korean
        os.environ.update({
            "Sonarr_Series_Id": "1", "Sonarr_Series_ImdbId": "tt1",
            "Sonarr_Series_Title": "Gold Land (2026)", "Sonarr_Series_Year": "2026",
            "Sonarr_OriginalLanguage": "English",       # ← stale, wrong env var
        })
        f.process_sonarr(folder)
        assert not os.path.isdir(folder)                            # renamed
        assert os.path.isdir(os.path.join(staging, "Gold Land (2026) [MDL 7.6]"))
        assert not os.path.isdir(os.path.join(staging, "Gold Land (2026) [IMDb 7.7]"))

    def test_mal_suffix_preserved_when_api_says_anime(self, staging, make_series, patch_providers, clear_env_vars):
        folder = make_series("Frieren (2023) [MAL 9.3]")
        patch_providers["mal_rating"]     = ("9.2", "MAL")
        patch_providers["imdb_rating"]    = "8.5"
        patch_providers["api_says_anime"] = True        # ← API tiebreaker says anime
        os.environ.update({
            "Sonarr_Series_Id": "2", "Sonarr_Series_ImdbId": "tt2",
            "Sonarr_Series_Title": "Frieren (2023)", "Sonarr_Series_Year": "2023",
            "Sonarr_Series_Type": "standard",           # ← stale, wrong env var
        })
        f.process_sonarr(folder)
        assert not os.path.isdir(folder)
        assert os.path.isdir(os.path.join(staging, "Frieren (2023) [MAL 9.2]"))
        assert not os.path.isdir(os.path.join(staging, "Frieren (2023) [IMDb 8.5]"))

    def test_tiebreaker_does_not_fire_without_prior_suffix(self, staging, make_series, patch_providers, clear_env_vars, monkeypatch):
        # No [MDL]/[MAL] suffix yet: the tiebreaker has no prior evidence to
        # disagree with, so it must NOT fire (would be a wasted API call).
        # We assert this by tracking whether force_api=True was ever used.
        folder = make_series("Some Show (2024)")
        force_api_calls = {"n": 0}
        orig_is_korean = f.is_korean_sonarr_series
        orig_is_anime  = f.is_anime_sonarr_series
        def spy_is_korean(series_id, force_api=False):
            if force_api: force_api_calls["n"] += 1
            return orig_is_korean(series_id, force_api=force_api)
        def spy_is_anime(series_id, path=None, force_api=False):
            if force_api: force_api_calls["n"] += 1
            return orig_is_anime(series_id, path=path, force_api=force_api)
        monkeypatch.setattr(f, "is_korean_sonarr_series", spy_is_korean)
        monkeypatch.setattr(f, "is_anime_sonarr_series",  spy_is_anime)
        patch_providers["imdb_rating"] = "8.0"
        os.environ.update({
            "Sonarr_Series_Id": "3", "Sonarr_Series_ImdbId": "tt3",
            "Sonarr_Series_Title": "Some Show (2024)", "Sonarr_Series_Year": "2024",
            "Sonarr_OriginalLanguage": "English",
        })
        f.process_sonarr(folder)
        assert force_api_calls["n"] == 0, (
            f"Tiebreaker called force_api={force_api_calls['n']} times on a "
            "folder with no prior MDL/MAL suffix — should not have fired."
        )
        assert os.path.isdir(os.path.join(staging, "Some Show (2024) [IMDb 8.0]"))

    def test_tiebreaker_respects_api_disagreement(self, staging, make_series, patch_providers, clear_env_vars):
        # Folder is [MDL 7.0] but API also says non-Korean (e.g. user re-tagged
        # the show in Sonarr). The tiebreaker should NOT override — the
        # demotion to IMDb is intentional.
        folder = make_series("Mislabeled (2024) [MDL 7.0]")
        patch_providers["imdb_rating"]     = "6.5"
        patch_providers["api_says_korean"] = False      # ← API agrees: NOT Korean
        os.environ.update({
            "Sonarr_Series_Id": "4", "Sonarr_Series_ImdbId": "tt4",
            "Sonarr_Series_Title": "Mislabeled (2024)", "Sonarr_Series_Year": "2024",
            "Sonarr_OriginalLanguage": "English",
        })
        f.process_sonarr(folder)
        # Folder was correctly re-rated as IMDb.
        assert os.path.isdir(os.path.join(staging, "Mislabeled (2024) [IMDb 6.5]"))


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

    def test_merge_discards_generated_artifacts_from_source(self, staging):
        # Regression: a webhook-2 race that re-creates the source folder with
        # only folder.jpg used to leave an orphan folder behind ("Merge
        # incomplete: 1 item(s) remain") because the merge refused to delete
        # the conflicting copy. folder.jpg / folder.ico / desktop.ini are
        # arr_finisher-generated; dest's copy is current → safe to discard.
        old = os.path.join(staging, "Show (2024)")
        new = os.path.join(staging, "Show (2024) [IMDb 8.0]")
        os.makedirs(old); os.makedirs(new)
        # Source has only generated artifacts (mimics Sonarr re-creating the
        # path between webhooks for a metadata write).
        open(os.path.join(old, "folder.jpg"), "wb").write(b"stale-poster")
        open(os.path.join(old, "folder.ico"), "wb").write(b"stale-icon")
        open(os.path.join(old, "desktop.ini"), "wb").write(b"stale-ini")
        # Dest already has live copies.
        open(os.path.join(new, "folder.jpg"), "wb").write(b"current-poster")
        open(os.path.join(new, "folder.ico"), "wb").write(b"current-icon")
        open(os.path.join(new, "desktop.ini"), "wb").write(b"current-ini")

        f.rename_folder(old, "8.0", "IMDb")

        # Source folder cleaned up — no orphan left behind.
        assert not os.path.isdir(old), "Source orphan should be removed"
        # Dest's live copies untouched.
        with open(os.path.join(new, "folder.jpg"), "rb") as fh:
            assert fh.read() == b"current-poster"

    def test_merge_discards_generated_keeps_user_data(self, staging):
        # Source has both a generated artifact (conflicts) AND user data
        # (no conflict). Artifact gets discarded, user data moves, source
        # folder cleaned up.
        old = os.path.join(staging, "Show (2024)")
        new = os.path.join(staging, "Show (2024) [IMDb 8.0]")
        os.makedirs(old); os.makedirs(new)
        open(os.path.join(old, "folder.jpg"), "wb").write(b"stale-poster")
        open(os.path.join(old, "S01E02.mkv"), "wb").write(b"new-episode")
        open(os.path.join(new, "folder.jpg"), "wb").write(b"current-poster")

        f.rename_folder(old, "8.0", "IMDb")

        assert not os.path.isdir(old)
        assert os.path.isfile(os.path.join(new, "S01E02.mkv"))
        with open(os.path.join(new, "folder.jpg"), "rb") as fh:
            assert fh.read() == b"current-poster"

    def test_merge_user_data_conflict_still_leaves_source(self, staging):
        # Mixed: generated artifact (safe to discard) + user data (must
        # preserve). The user-data conflict triggers "Merge incomplete" but
        # the generated artifact is still cleaned up so the user only has
        # to deal with the real conflict.
        old = os.path.join(staging, "Show (2024)")
        new = os.path.join(staging, "Show (2024) [IMDb 8.0]")
        os.makedirs(old); os.makedirs(new)
        open(os.path.join(old, "folder.jpg"), "wb").write(b"stale-poster")
        open(os.path.join(old, "S01E01.mkv"), "wb").write(b"source-version")
        open(os.path.join(new, "folder.jpg"), "wb").write(b"current-poster")
        open(os.path.join(new, "S01E01.mkv"), "wb").write(b"dest-version")

        f.rename_folder(old, "8.0", "IMDb")

        # User data left in source for manual resolution.
        assert os.path.isfile(os.path.join(old, "S01E01.mkv"))
        with open(os.path.join(old, "S01E01.mkv"), "rb") as fh:
            assert fh.read() == b"source-version"
        # Generated artifact was discarded from source (no longer there).
        assert not os.path.exists(os.path.join(old, "folder.jpg"))
        # Dest's user-data version untouched.
        with open(os.path.join(new, "S01E01.mkv"), "rb") as fh:
            assert fh.read() == b"dest-version"


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

def _fake_build_writing_ico(record):
    """Return a _build_folder_ico stand-in that records calls and writes a
    placeholder folder.ico so the rest of create_folder_icon can proceed."""
    def _build(poster, ico):
        record["n"] = record.get("n", 0) + 1
        with open(ico, "wb") as fh:
            fh.write(b"ico")
        return True
    return _build


class TestFolderIconSkip:
    def test_tooltip_only_desktop_ini_does_not_block_icon(self, tmp_path, monkeypatch):
        # Simulate the scenario: an earlier run wrote a tooltip (so
        # desktop.ini exists) but never wrote an IconResource line. A later
        # run with ENABLE_CREATE_FOLDER_ICON=True must NOT short-circuit.
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "folder.jpg").write_bytes(b"poster")
        (folder / "desktop.ini").write_bytes(
            "[.ShellClassInfo]\r\nInfoTip=tooltip without icon\r\n".encode("utf-16")
        )
        built = {}
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "_build_folder_ico", _fake_build_writing_ico(built))
        f.create_folder_icon(str(folder))
        assert built.get("n") == 1   # the icon was built (not short-circuited)

    def test_desktop_ini_with_icon_resource_skips(self, tmp_path, monkeypatch):
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "folder.jpg").write_bytes(b"poster")
        (folder / "desktop.ini").write_bytes(
            "[.ShellClassInfo]\r\nIconResource=poster.ico,0\r\n".encode("utf-16")
        )
        built = {}
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "_build_folder_ico", _fake_build_writing_ico(built))
        f.create_folder_icon(str(folder))
        assert built.get("n") is None   # short-circuited; never built

    def test_no_poster_skips(self, tmp_path, monkeypatch):
        # No folder.jpg → nothing to build an icon from; must skip cleanly.
        folder = tmp_path / "show"
        folder.mkdir()
        built = {}
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "_build_folder_ico", _fake_build_writing_ico(built))
        f.create_folder_icon(str(folder))
        assert built.get("n") is None


# ---------- _build_folder_ico: native poster -> multi-size .ico ----------

class TestBuildFolderIco:
    def test_produces_all_sizes_on_transparent_canvas(self, tmp_path):
        from PIL import Image
        jpg = tmp_path / "folder.jpg"
        # 2:3 poster, to exercise the fit-and-center-pad path.
        Image.new("RGB", (600, 900), (200, 30, 30)).save(str(jpg), "JPEG")
        ico = tmp_path / "folder.ico"
        assert f._build_folder_ico(str(jpg), str(ico)) is True
        with Image.open(str(ico)) as im:
            assert set(im.ico.sizes()) == {(s, s) for s in f._ICO_SIZES}
            # 256 frame is a square canvas with transparent padding around a
            # portrait poster — the top-left corner must be fully transparent.
            frame = im.ico.getimage((256, 256)).convert("RGBA")
            assert frame.getpixel((0, 0))[3] == 0

    def test_returns_false_on_unreadable_image(self, tmp_path):
        bad = tmp_path / "folder.jpg"
        bad.write_bytes(b"this is not an image")
        ico = tmp_path / "folder.ico"
        assert f._build_folder_ico(str(bad), str(ico)) is False
        assert not ico.exists()           # no partial/leftover .ico


# ---------- _apply_icon_to_desktop_ini: icon keys, preserve InfoTip ----------

class TestApplyIconToDesktopIni:
    def test_writes_icon_keys_and_foldertype(self, tmp_path):
        folder = tmp_path / "show"
        folder.mkdir()
        assert f._apply_icon_to_desktop_ini(str(folder)) is True
        text = (folder / "desktop.ini").read_bytes().decode("utf-16")
        assert "IconResource=folder.ico,0" in text
        assert "IconFile=folder.ico" in text
        assert "IconIndex=0" in text
        assert "FolderType=Videos" in text

    def test_preserves_existing_infotip(self, tmp_path):
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "desktop.ini").write_bytes(
            "[.ShellClassInfo]\r\nInfoTip=Keep this plot\r\n".encode("utf-16")
        )
        assert f._apply_icon_to_desktop_ini(str(folder)) is True
        text = (folder / "desktop.ini").read_bytes().decode("utf-16")
        assert "InfoTip=Keep this plot" in text
        assert "IconResource=folder.ico,0" in text


# ---------- --force / FORCE_REBUILD: bypass idempotency skips ----------

class TestForceRebuild:
    def test_icon_skip_bypassed_when_forced(self, tmp_path, monkeypatch):
        # An IconResource line normally short-circuits create_folder_icon.
        # With FORCE_REBUILD=True, the icon must be rebuilt anyway.
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "folder.jpg").write_bytes(b"poster")
        (folder / "desktop.ini").write_bytes(
            "[.ShellClassInfo]\r\nIconResource=poster.ico,0\r\n".encode("utf-16")
        )
        built = {}
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "FORCE_REBUILD", True)
        monkeypatch.setattr(f, "_build_folder_ico", _fake_build_writing_ico(built))
        f.create_folder_icon(str(folder))
        assert built.get("n") == 1   # rebuilt despite an existing IconResource

    def test_icon_wipes_existing_files_when_forced(self, tmp_path, monkeypatch):
        # --force must delete folder.ico + desktop.ini BEFORE the rebuild, so
        # Windows drops its cached icon for this folder path.
        folder = tmp_path / "show"
        folder.mkdir()
        ico = folder / "folder.ico"
        ini = folder / "desktop.ini"
        jpg = folder / "folder.jpg"
        ico.write_bytes(b"old-icon-bytes")
        ini.write_bytes("[.ShellClassInfo]\r\nIconResource=folder.ico,0\r\n".encode("utf-16"))
        jpg.write_bytes(b"poster-source")   # must be preserved (icon source)

        present_at_build = {}
        def fake_build(poster, ico_path):
            # Snapshot which files exist when the rebuild starts.
            present_at_build["ico"] = ico.exists()
            present_at_build["ini"] = ini.exists()
            present_at_build["jpg"] = jpg.exists()
            with open(ico_path, "wb") as fh:
                fh.write(b"ico")
            return True
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "FORCE_REBUILD", True)
        monkeypatch.setattr(f, "_build_folder_ico", fake_build)
        f.create_folder_icon(str(folder))
        assert present_at_build["ico"] is False, "folder.ico must be wiped before rebuild"
        assert present_at_build["ini"] is False, "desktop.ini must be wiped before rebuild"
        assert present_at_build["jpg"] is True,  "folder.jpg must be preserved (icon source)"

    def test_icon_does_not_wipe_without_force(self, tmp_path, monkeypatch):
        # Without --force, the wipe must NOT happen — webhook + sweep behavior
        # relies on the idempotency skip, not on pre-deleting files.
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "folder.jpg").write_bytes(b"poster")
        # No existing desktop.ini → create_folder_icon proceeds past the skip,
        # but should not pre-delete anything else.
        ico = folder / "folder.ico"
        ico.write_bytes(b"do-not-delete-me")
        present = {}
        def fake_build(poster, ico_path):
            present["ico"] = ico.exists()
            return True   # don't overwrite — we're checking the wipe didn't run
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "FORCE_REBUILD", False)
        monkeypatch.setattr(f, "_build_folder_ico", fake_build)
        f.create_folder_icon(str(folder))
        assert present["ico"] is True, "folder.ico must NOT be wiped when --force is off"

    def test_tooltip_rewritten_when_forced(self, tmp_path, monkeypatch):
        # Tooltip normally skipped when InfoTip matches. --force rewrites it.
        monkeypatch.setattr(f, "ENABLE_SET_TOOLTIP", True)
        folder = tmp_path / "show"
        folder.mkdir()
        f.set_folder_tooltip(str(folder), "Same tooltip")
        ini = folder / "desktop.ini"
        first_mtime = ini.stat().st_mtime_ns
        import time as _t
        _t.sleep(0.01)
        monkeypatch.setattr(f, "FORCE_REBUILD", True)
        f.set_folder_tooltip(str(folder), "Same tooltip")
        # Force-rebuild should rewrite even though content is identical.
        assert ini.stat().st_mtime_ns != first_mtime

    def test_existing_lnk_recreated_when_forced(self, tmp_path, monkeypatch):
        # _write_lnk normally short-circuits on existing files. --force deletes
        # and recreates them.
        lnk = tmp_path / "Existing.lnk"
        lnk.write_text("placeholder")   # pretend an old shortcut is here
        monkeypatch.setattr(f, "FORCE_REBUILD", True)
        # Stub the COM bits — we only care that the old file gets deleted and
        # _write_lnk proceeds past the existence check.
        monkeypatch.setattr(f, "HAS_WIN32COM", False)   # logs an error after deletion, then returns
        f._write_lnk(str(lnk), "https://example.com", "Existing")
        # File was removed (and not recreated because HAS_WIN32COM=False).
        assert not lnk.exists()

    def test_existing_lnk_kept_when_not_forced(self, tmp_path, monkeypatch):
        # Sanity counterpart: without --force, _write_lnk leaves existing files alone.
        lnk = tmp_path / "Existing.lnk"
        lnk.write_text("placeholder")
        monkeypatch.setattr(f, "FORCE_REBUILD", False)
        monkeypatch.setattr(f, "FORCE_REGENERATE_SHORTCUTS", False)
        f._write_lnk(str(lnk), "https://example.com", "Existing")
        # File untouched.
        assert lnk.read_text() == "placeholder"


# ---------- --force CLI guards ----------

class TestForceCliGuards:
    def _run_main(self, argv, monkeypatch):
        """Invoke f.main() with argv; return SystemExit code (argparse uses exit)."""
        monkeypatch.setattr("sys.argv", ["arr_finisher.py"] + argv)
        try:
            f.main()
        except SystemExit as e:
            return e.code
        return 0

    def test_force_without_service_path_errors(self, monkeypatch, capsys):
        rc = self._run_main(["--force"], monkeypatch)
        # argparse.error exits with code 2 and prints to stderr.
        assert rc == 2
        err = capsys.readouterr().err
        assert "--force" in err and "--service" in err

    def test_force_with_sweep_errors(self, monkeypatch, capsys):
        rc = self._run_main(["--sweep", "--force", "--service", "radarr",
                             "--path", "x"], monkeypatch)
        assert rc == 2
        err = capsys.readouterr().err
        assert "--force" in err and "--sweep" in err


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


# ---------- HIGH-1 / HIGH-2: DRY_RUN respected by force-flag wipe paths ----------

class TestForceDryRun:
    """Regression: --force --dry-run (or --regenerate-shortcuts --dry-run)
    must not mutate disk. Previously, the os.remove in _write_lnk and the
    bulk wipe in create_shortcuts ran before the DRY_RUN check, deleting
    files in what was supposed to be a preview-only run."""

    def test_write_lnk_does_not_delete_under_dry_run(self, tmp_path, monkeypatch):
        lnk = tmp_path / "Existing.lnk"
        lnk.write_text("ORIGINAL")
        monkeypatch.setattr(f, "FORCE_REBUILD", True)
        monkeypatch.setattr(f, "DRY_RUN", True)
        f._write_lnk(str(lnk), "https://example.com", "Existing")
        assert lnk.exists()
        assert lnk.read_text() == "ORIGINAL"

    def test_write_lnk_does_not_delete_under_dry_run_regen_shortcuts(self, tmp_path, monkeypatch):
        # Same guarantee for the FORCE_REGENERATE_SHORTCUTS path (used by
        # --regenerate-shortcuts) — both flags share the wipe codepath.
        lnk = tmp_path / "Existing.lnk"
        lnk.write_text("ORIGINAL")
        monkeypatch.setattr(f, "FORCE_REGENERATE_SHORTCUTS", True)
        monkeypatch.setattr(f, "DRY_RUN", True)
        f._write_lnk(str(lnk), "https://example.com", "Existing")
        assert lnk.read_text() == "ORIGINAL"

    def test_create_shortcuts_bulk_wipe_skipped_under_dry_run(self, tmp_path, monkeypatch):
        folder = tmp_path / "show"
        folder.mkdir()
        links = folder / "Links"
        links.mkdir()
        lnk = links / "IMDb.lnk"
        vbs = links / "Subtitle.vbs"
        lnk.write_text("PRESERVE")
        vbs.write_text("PRESERVE")
        monkeypatch.setattr(f, "FORCE_REBUILD", True)
        monkeypatch.setattr(f, "DRY_RUN", True)
        f.create_shortcuts("radarr", str(folder), "tt1", "123", "Show",
                           is_korean=False, is_anime=False, year="2024")
        assert lnk.exists() and lnk.read_text() == "PRESERVE"
        assert vbs.exists() and vbs.read_text() == "PRESERVE"

    def test_makedirs_skipped_under_dry_run(self, tmp_path, monkeypatch):
        folder = tmp_path / "show"
        folder.mkdir()
        # Links/ doesn't exist yet
        monkeypatch.setattr(f, "DRY_RUN", True)
        f.create_shortcuts("radarr", str(folder), "tt1", "123", "Show",
                           is_korean=False, is_anime=False, year="2024")
        # Under DRY_RUN, no Links/ directory should appear on disk.
        assert not (folder / "Links").exists()


# ---------- HIGH-3: get_mdl_rating malformed-JSON treated as transient ----------

class TestKuryanaMalformedJson:
    """Regression: malformed JSON from a kuryana mirror used to mask outages.
    Mirror 0 returns 200 + garbage, mirror 1 unreachable — function used to
    return None (silently falling back to IMDb), masking the real outage."""

    def _stub_http(self, monkeypatch, responses):
        """`responses` is a list of either ('json', body_dict) | ('raise', exc) | ('status', code)."""
        idx = {"n": 0}
        def fake_get(url, timeout=15):
            i = idx["n"]
            idx["n"] += 1
            if i >= len(responses):
                raise ConnectionError("ran out of stubbed responses")
            kind, payload = responses[i]
            if kind == "raise":
                raise payload
            if kind == "status":
                class _R:
                    status_code = payload
                    def json(self_):
                        return {}
                return _R()
            class _R:
                status_code = 200
                def json(self_):
                    if isinstance(payload, Exception):
                        raise payload
                    return payload
            return _R()
        monkeypatch.setattr(f, "http", lambda: type("S", (), {"get": staticmethod(fake_get)})())

    def test_all_mirrors_5xx_raises(self, monkeypatch):
        self._stub_http(monkeypatch, [("status", 503), ("status", 503)])
        with pytest.raises(f.ProviderUnavailable):
            f.get_mdl_rating("Test", year="2024")

    def test_malformed_then_network_error_raises(self, monkeypatch):
        # The exact bug: mirror 0 returns 200 with body that can't be parsed
        # as JSON; mirror 1 is unreachable. Both mirrors effectively failed;
        # contract says raise, not return None.
        self._stub_http(monkeypatch, [
            ("json", ValueError("malformed")),
            ("raise", ConnectionError("mirror 1 down")),
        ])
        with pytest.raises(f.ProviderUnavailable):
            f.get_mdl_rating("Test", year="2024")

    def test_malformed_then_good_returns_match(self, monkeypatch):
        # If mirror 0 is broken but mirror 1 gives a clean response with a
        # confident match, we should return that match.
        self._stub_http(monkeypatch, [
            ("json", ValueError("malformed")),
            ("json", {"results": {"dramas": [
                {"title": "Test Show", "year": "2024", "type": "Korean Drama",
                 "rating": "8.5", "slug": "12345-test-show"}
            ]}}),
        ])
        result = f.get_mdl_rating("Test Show", year="2024")
        assert result == ("8.5", "MDL")

    def test_404_then_404_returns_none(self, monkeypatch):
        # Both mirrors responsive but no resource — return None (no match),
        # NOT raise. Confirms 404 doesn't get mistaken for "transient".
        self._stub_http(monkeypatch, [("status", 404), ("status", 404)])
        assert f.get_mdl_rating("Nonexistent", year="2024") is None


# ---------- MED-2: regenerate_shortcuts tiebreaker preserves classification ----------

class TestRegenerateShortcutsTiebreaker:
    """Regression: regenerate_shortcuts used to skip the env-vs-API tiebreaker
    that _process applies. A [MAL X.X] folder whose Sonarr seriesType went
    stale would lose its MyAnimeList.lnk on regen, then get a stray
    MyDramaList.lnk if the language env var also lied."""

    def test_mal_folder_keeps_anime_classification_via_tiebreaker(
            self, staging, patch_providers, monkeypatch):
        folder = os.path.join(staging, "Frieren (2023) [MAL 9.3]")
        os.makedirs(folder)
        fake_obj = {
            "id": 1, "imdbId": "tt2", "tvdbId": 999,
            "title": "Frieren (2023)", "year": 2023,
            "seriesType": "standard",   # ← env-fast-path would say non-anime
            "originalLanguage": {"name": "Japanese"},
        }
        monkeypatch.setattr(f, "get_object_by_path", lambda svc, p: fake_obj)
        monkeypatch.setattr(f, "_default_sweep_roots", lambda: [(staging, "sonarr")])
        patch_providers["api_says_anime"] = True   # API tiebreaker confirms anime
        rc = f.regenerate_shortcuts()
        assert rc == 0
        assert os.path.isfile(os.path.join(folder, "Links", "MyAnimeList.lnk")), (
            "Tiebreaker should have preserved anime classification → MyAnimeList.lnk created"
        )


# ---------- MED-4: --force preserves existing tooltip across icon wipe ----------

class TestReadDesktopIniInfoTip:
    def test_reads_existing_infotip(self, tmp_path):
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "desktop.ini").write_bytes(
            "[.ShellClassInfo]\r\nInfoTip=Some plot summary\r\n".encode("utf-16")
        )
        assert f._read_desktop_ini_infotip(str(folder)) == "Some plot summary"

    def test_returns_none_when_no_desktop_ini(self, tmp_path):
        folder = tmp_path / "show"
        folder.mkdir()
        assert f._read_desktop_ini_infotip(str(folder)) is None

    def test_returns_none_when_infotip_in_wrong_section(self, tmp_path):
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "desktop.ini").write_bytes(
            "[OtherSection]\r\nInfoTip=Wrong section\r\n".encode("utf-16")
        )
        assert f._read_desktop_ini_infotip(str(folder)) is None


class TestForceRebuildPreservesTooltip:
    def test_existing_infotip_restored_after_wipe(self, tmp_path, monkeypatch):
        # Real scenario: folder has a working tooltip, --force wipes desktop.ini,
        # the rebuild writes a fresh one with IconResource only (no InfoTip).
        # Without the restore, the tooltip is lost until the next webhook run
        # fetches a fresh plot.
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "folder.jpg").write_bytes(b"poster")
        ini = folder / "desktop.ini"
        ini.write_bytes(
            ("[.ShellClassInfo]\r\nIconResource=folder.ico,0\r\n"
             "InfoTip=User's old tooltip\r\n").encode("utf-16")
        )
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "ENABLE_SET_TOOLTIP", True)
        monkeypatch.setattr(f, "FORCE_REBUILD", True)
        monkeypatch.setattr(f, "_build_folder_ico", _fake_build_writing_ico({}))
        f.create_folder_icon(str(folder))
        text = ini.read_bytes().decode("utf-16")
        assert "InfoTip=User's old tooltip" in text, (
            "Pre-wipe tooltip should be restored after the desktop.ini rewrite"
        )
        assert "IconResource=folder.ico,0" in text, (
            "The IconResource line must still be present"
        )

    def test_no_tooltip_to_preserve_is_safe(self, tmp_path, monkeypatch):
        # Folder has IconResource but no InfoTip. Wipe + regenerate must not
        # invent a tooltip or crash trying to restore a None.
        folder = tmp_path / "show"
        folder.mkdir()
        (folder / "folder.jpg").write_bytes(b"poster")
        ini = folder / "desktop.ini"
        ini.write_bytes(
            "[.ShellClassInfo]\r\nIconResource=folder.ico,0\r\n".encode("utf-16")
        )
        monkeypatch.setattr(f, "ENABLE_CREATE_FOLDER_ICON", True)
        monkeypatch.setattr(f, "ENABLE_SET_TOOLTIP", True)
        monkeypatch.setattr(f, "FORCE_REBUILD", True)
        monkeypatch.setattr(f, "_build_folder_ico", _fake_build_writing_ico({}))
        f.create_folder_icon(str(folder))
        text = ini.read_bytes().decode("utf-16")
        assert "InfoTip=" not in text, "No tooltip should be invented from nothing"
