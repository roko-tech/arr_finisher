# Changelog

All notable changes to arr_finisher are listed here. Versions follow [SemVer](https://semver.org).

## [1.0.0] — 2026-05-16

First versioned release. Marks the post-review baseline: all known critical
and high-priority issues from the project review are addressed, with strong
test coverage and atomic disk operations.

### Added
- `--clear-cache`: delete the rating-freshness cache.
- `--refresh <imdb_id>`: invalidate one cache entry; next sweep re-fetches.
- `--sweep --force-refresh`: bypass the cache TTL and re-rate every folder
  in this run (without permanently clearing the cache).
- `--regenerate-shortcuts`: walk library roots and rebuild every `Links/`
  shortcut from current code. Useful when URL formats change (e.g. provider
  rebrand). Skips ratings/icons/tooltips.
- `--version`: print version and exit.
- `ENABLE_SET_TOOLTIP` feature toggle, for symmetry with other toggles.
- `ARR_FINISHER_RATING_CACHE_TTL_DAYS` env var (was a hardcoded constant).
- `ARR_FINISHER_SWEEP_ROOTS` env var — set sweep roots without editing code.
- `IMDB_GRAPHQL_URL` env var — override the IMDb GraphQL endpoint without
  editing the script. Matches the KURYANA_BASE_URL / JIKAN_BASE_URL pattern.
- Auto-discovery of sweep roots from Sonarr/Radarr `/api/v3/rootfolder`
  when no `--roots` flag or env var is set.
- Setup-help sidecar (`arr_finisher_setup.txt`) is written next to the
  script when critical config is missing, so first-time users see what to fix.
- Rollback events also appended to `.rollbacks.log` for easy triage.
- 33 net-new tests covering `--roots` parsing, rename-merge conflicts, URL
  safety, log redaction, sweep env-var stuffing, cache TTL, cache commands,
  OMDb response caching, ISO timestamps, root discovery, setup-help, and
  the no-silent-fallback-to-IMDb contract for MAL/MDL outages.
- `get_imdb_rating_from_graphql`: IMDb's public GraphQL endpoint (no API
  key) is now the primary IMDb-rating source. Returns live ratings; OMDb
  remains the fallback for when GraphQL is unreachable.
- `ProviderUnavailable` exception raised by `get_mal_rating` and
  `get_mdl_rating` on transient errors (5xx / 429 / network / malformed
  responses) — the rating dispatcher in `_process` catches it and preserves
  the folder's existing rating instead of silently rebadging anime as IMDb
  when Jikan is down. (Real-world bug: 13 anime folders were rebadged
  during a Jikan outage before this fix.)
- `--validate` probes the IMDb GraphQL endpoint and detects invalid
  OMDb API keys (200 response with `Response: False`).
- Sweep summary now reports the number of provider outages encountered.
- Per-process cache of the Sonarr/Radarr library list — turns a sweep's
  service-API traffic from O(folders) into O(1).

### Changed
- Rating-cache `checked_at` is now an ISO-8601 string (legacy float-epoch
  entries still read correctly).
- `process_sonarr` / `process_radarr` collapsed into `_process(service, path)`
  driven by per-service config (`_SERVICE_ADAPTERS`). Public functions unchanged.
- Webhook mode persists the rating cache on exit so the next sweep starts warm.
- OMDb responses cached per process — rating + plot share a single request.
- `--validate` probes all Kuryana mirrors (pass if any responds) and
  optionally probes SubDL + OpenSubtitles when their keys are set. Also
  probes the IMDb GraphQL endpoint (primary rating source).
- All `subprocess attrib` shell calls replaced with `ctypes.SetFileAttributesW`.
- Logger is named (`arr_finisher`), not the root logger.
- `_fs_lock` retries on contention and reclaims stale (>10 min) lock dirs
  instead of silently proceeding without a lock.
- Sweep TTL skip now requires the folder to ALREADY carry a rating suffix.
  Folders that are missing one (newly added, or stripped by the user) are
  always re-checked, regardless of cache age.
- Twitter shortcut OR clause is parenthesized so `lang:<code>` applies to
  the whole alternation rather than only the last term.

### Fixed
- `--sweep --roots "D:\TV Shows:sonarr"` now parses correctly (was a
  `ValueError` because `split(":")` produced three parts).
- `rename_folder` merge branch no longer `rmtree`s the source when items
  conflict with the destination — those items were being silently deleted.
- Manual mode (`--service` + `--path`) looks up the full object from
  Sonarr/Radarr instead of silently no-opping with empty env vars.
- `Subtitle.vbs` rejects unsafe URLs (non-`https://`, embedded quotes or
  newlines) to prevent script injection via compromised upstream APIs.
- `Subtitle.vbs` is now rewritten when the resolved URLs differ from the
  on-disk content (was: only written when missing, so a later SUBDL key
  addition didn't take effect until manual deletion).
- OMDb / SubDL API keys redacted from error log messages.
- `desktop.ini` reader uses a context manager (no file-handle leak); the
  `cp1252` fallback decoder is sanity-checked instead of blindly accepted.
- `desktop.ini` and `.rating_cache.json` writes are atomic (tmp + replace).
- `create_folder_icon` now checks for an `IconResource=` line specifically,
  not just `desktop.ini` existence (which `set_folder_tooltip` also writes).
  A failed-icon + succeeded-tooltip first run no longer permanently locks
  out future icon creation.
- `_fs_lock` now actually skips work when the lock can't be acquired
  (previously yielded `acquired=False` that the only caller ignored).
- `_save_rating_cache` only writes when the cache has been modified.
- Twitter shortcut hashtag preserves Unicode letters.
- Twitter shortcut switched from `twitter.com` to `x.com` (canonical domain
  since 2023).
- Anime / Korean path heuristic is now case-insensitive (`D:\anime\` works,
  not just `D:\Anime\`).
- `_load_env_file` only strips a matching pair of outer quotes (was: naive
  sequential `.strip('"').strip("'")` that mangled mixed-quote values).
- Sonarr/Radarr `Test` event is acknowledged before the config check fires,
  so a fresh install verifying connectivity doesn't get a spurious
  setup-help sidecar.
- `--validate` icon list no longer claims SubDL/Subsource as required —
  those .ico files were legacy and never referenced by any `.lnk`.
- MAL/MDL providers now treat any unexpected exception (malformed JSON,
  TypeError, AttributeError) as a transient outage rather than letting it
  crash the webhook.
- Tooltip is no longer overwritten with a plot-only version during a
  provider outage — would have flickered the rating off until the next
  successful run.
- `argparse` errors if `--service` and `--path` are used without each other.
- README install step uses `copy` (Windows cmd) instead of `cp` (Unix).
- Hardcoded `D:\TV Shows / D:\Anime / E:\Movies` fallback sweep roots
  removed; `sweep_library` refuses to run with no configured roots rather
  than walking the original-author's drive layout on a fresh install.

### Removed
- Dead `existing_encoding` variable in `set_folder_tooltip`.
- Stray `Sonarr_Movie_OriginalLanguage` / `Radarr_Series_OriginalLanguage`
  keys from the language env-var probe list (Sonarr has no movies; Radarr
  has no series).
- `_extract_from_jsonld` — the HTML JSON-LD scrape stopped working when
  IMDb moved behind AWS WAF (every `requests` call returns HTTP 202 with
  an empty body until JavaScript clears the challenge). GraphQL replaces it.
