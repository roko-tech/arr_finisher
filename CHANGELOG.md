# Changelog

All notable changes to arr_finisher are listed here. Versions follow [SemVer](https://semver.org).

## [1.0.0] — 2026-05-16

First versioned release. Marks the post-review baseline: all known critical
and high-priority issues from the project review are addressed, with strong
test coverage and atomic disk operations.

### Added
- `--clear-cache`: delete the rating-freshness cache.
- `--refresh <imdb_id>`: invalidate one cache entry; next sweep re-fetches.
- `--version`: print version and exit.
- `ENABLE_SET_TOOLTIP` feature toggle, for symmetry with other toggles.
- `ARR_FINISHER_RATING_CACHE_TTL_DAYS` env var (was a hardcoded constant).
- `ARR_FINISHER_SWEEP_ROOTS` env var — set sweep roots without editing code.
- Auto-discovery of sweep roots from Sonarr/Radarr `/api/v3/rootfolder`
  when no `--roots` flag or env var is set.
- Setup-help sidecar (`arr_finisher_setup.txt`) is written next to the
  script when critical config is missing, so first-time users see what to fix.
- Rollback events also appended to `.rollbacks.log` for easy triage.
- 21 new tests covering `--roots` parsing, rename-merge conflicts, URL
  safety, log redaction, sweep env-var stuffing, cache TTL, cache commands,
  and OMDb response caching.

### Changed
- Rating-cache `checked_at` is now an ISO-8601 string (legacy float-epoch
  entries still read correctly).
- `process_sonarr` / `process_radarr` collapsed into `_process(service, path)`
  driven by per-service config (`_SERVICE_ADAPTERS`). Public functions unchanged.
- Webhook mode persists the rating cache on exit so the next sweep starts warm.
- OMDb responses cached per process — rating + plot share a single request.
- `--validate` probes all Kuryana mirrors (pass if any responds) and
  optionally probes SubDL + OpenSubtitles when their keys are set.
- All `subprocess attrib` shell calls replaced with `ctypes.SetFileAttributesW`.
- Logger is named (`arr_finisher`), not the root logger.
- `_fs_lock` retries on contention and reclaims stale (>10 min) lock dirs
  instead of silently proceeding without a lock.

### Fixed
- `--sweep --roots "D:\TV Shows:sonarr"` now parses correctly (was a
  `ValueError` because `split(":")` produced three parts).
- `rename_folder` merge branch no longer `rmtree`s the source when items
  conflict with the destination — those items were being silently deleted.
- Manual mode (`--service` + `--path`) looks up the full object from
  Sonarr/Radarr instead of silently no-opping with empty env vars.
- `Subtitle.vbs` rejects unsafe URLs (non-`https://`, embedded quotes or
  newlines) to prevent script injection via compromised upstream APIs.
- OMDb / SubDL API keys redacted from error log messages.
- `desktop.ini` reader uses a context manager (no file-handle leak); the
  `cp1252` fallback decoder is sanity-checked instead of blindly accepted.
- `desktop.ini` and `.rating_cache.json` writes are atomic (tmp + replace).
- Twitter shortcut hashtag preserves Unicode letters.
- `argparse` errors if `--service` and `--path` are used without each other.
- README install step uses `copy` (Windows cmd) instead of `cp` (Unix).

### Removed
- Dead `existing_encoding` variable in `set_folder_tooltip`.
- Stray `Sonarr_Movie_OriginalLanguage` / `Radarr_Series_OriginalLanguage`
  keys from the language env-var probe list (Sonarr has no movies; Radarr
  has no series).
