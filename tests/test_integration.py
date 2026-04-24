"""Integration tests that hit real external APIs. Can be slow or skipped."""
import os
import pytest
import arr_finisher as f


# Skip these if NETWORK_TESTS != "1" — they depend on external services being up.
pytestmark = pytest.mark.skipif(
    os.environ.get("NETWORK_TESTS") != "1",
    reason="set NETWORK_TESTS=1 to enable (hits real APIs)",
)


class TestKuryanaIntegration:
    def test_reverse_2026_returns_mdl(self):
        result = f.get_mdl_rating("Reverse", year="2026")
        assert result is not None
        rating, source = result
        assert source == "MDL"
        assert 0 < float(rating) <= 10

    def test_gibberish_title_rejected(self):
        # Quality filter should reject bad matches
        result = f.get_mdl_rating("XXYYZZdefinitelynotakoreandramaabc", year="2099")
        assert result is None


class TestJikanIntegration:
    def test_frieren_returns_mal(self):
        result = f.get_mal_rating("Frieren Beyond Journeys End", year=2023)
        assert result is not None
        rating, source = result
        assert source == "MAL"
        assert 0 < float(rating) <= 10

    def test_gibberish_rejected(self):
        result = f.get_mal_rating("XXYYZZnotananimeeverrrabc123", year=2099)
        assert result is None


class TestValidate:
    def test_validate_config_runs_without_error(self):
        # Just check it doesn't raise. Real assertion is return code.
        rc = f.validate_config()
        assert rc in (0, 1)   # either is OK — depends on current env
