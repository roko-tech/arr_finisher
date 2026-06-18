# Changelog

All notable changes to arr_finisher are listed here. Versions follow [SemVer](https://semver.org).

## Unreleased

### Added
- `--force` flag for manual mode (`--service` + `--path`): rewrite folder icon,
  shortcuts, and tooltip even when they're already up to date. Bypasses the
  idempotency skips that otherwise make a re-run of an unchanged folder a no-op.
  Single-folder only; rejected with `--sweep` or `--regenerate-shortcuts`.
  Under the hood: wipes `folder.ico` + `desktop.ini` before rebuilding the
  icon (busts Windows' per-path icon cache) and forces every `Links/*.lnk`
  and the tooltip to be rewritten.
- `_apply_content_class_tiebreaker` helper, used by both `_process` and
  `regenerate_shortcuts` so the env-vs-API tiebreaker logic doesn't drift
  between the two call sites.
- `--check-rollbacks` command: scans `.rollbacks.log` for unresolved `FAIL`
  entries (disk/service desync) and exits non-zero if any are found, so the
  rare-but-serious case no longer relies on the user remembering to grep.
- `ENABLE_HIDE_METADATA` feature toggle (default off): also hides `.nfo` +
  extra-artwork sidecars in each folder, restoring the cleaner Explorer view
  the old Creator.exe `-h` produced — opt-in, without the aggressive default.

### Changed
- **Folder icons are now generated natively** — the external Folder-Icon-Creator
  binary (`Creator.exe`) is no longer required or invoked. `folder.ico` is built
  with Pillow (poster resized to fit 256×256, centered on a transparent canvas,
  sizes 256→16) and bound via the same Windows shell APIs the tool used
  (`SHGetSetFolderCustomSettings` + an `ie4uinit.exe` cache refresh). The
  `FOLDER_ICON_EXE` setting and the external download are gone from setup.
  - Add `Pillow` to your environment: `pip install -r requirements.txt`.
  - Metadata hiding is now minimal — only the generated `folder.jpg` and
    `folder.ico` are hidden. The old path passed `-h`, which recursively hid
    every non-media file (`.nfo`, extra art, …); those stay visible now.

### Fixed
- `set_folder_tooltip` no longer writes doubled line endings (`\r\r\n`) into
  `desktop.ini` — text-mode newline translation turned each `\r\n` into
  `\r\r\n`. Now writes with `newline=""` and drops blank lines, so existing
  files with the old artifacts self-heal on the next write. (Harmless to
  Explorer either way, but the files are now well-formed.)
- `rename_folder` merge no longer leaves orphan source folders when the only
  conflicting items are arr_finisher-generated artifacts (`folder.jpg`,
  `folder.ico`, `desktop.ini`). The source's stale copies are discarded
  (dest's are the live ones); user data (`.mkv`, `.srt`, etc.) keeps the
  existing protective behavior unchanged. Fixes the common scenario where
  Sonarr/Radarr re-creates the source path between webhooks for a metadata
  write, then a stale-path webhook 2 leaves a leftover folder containing
  only `folder.jpg`.
- `--regenerate-shortcuts --dry-run` and `--force --dry-run` no longer delete
  existing `.lnk` / `.vbs` files in `Links/`, and no longer create an empty
  `Links/` directory. The wipe paths in `_write_lnk` and `create_shortcuts`
  now check `DRY_RUN` before mutating disk.
- `get_mdl_rating` now raises `ProviderUnavailable` when a kuryana mirror
  returns HTTP 200 with malformed JSON (previously masked the outage by
  returning `None`, causing a Korean drama to be re-rated as IMDb — the same
  failure mode the `ProviderUnavailable` contract was added to prevent in 1.0).
- `regenerate_shortcuts` now applies the `[MAL ...]` / `[MDL ...]` env-vs-API
  tiebreaker. A folder whose Sonarr `seriesType` env var went stale used to
  silently lose its MyAnimeList/MyDramaList shortcut on regen.
- `--force` now preserves any existing folder tooltip across the `desktop.ini`
  wipe — restored as a fallback after the fresh `desktop.ini` is written. If
  OMDb later returns a plot, the regular tooltip path overwrites it; if OMDb
  has nothing, the user's previous tooltip survives.
- README's "Sweep roots" section no longer claims a hardcoded
  `D:\TV Shows / D:\Anime / E:\Movies` fallback (the fallback was removed in
  1.0.0 but the doc was stale).
- `--dry-run` is now honored by `rename_folder`'s merge branch. The merge
  (`shutil.move` / `rmtree`) previously ran *before* the `DRY_RUN` check, so a
  preview run (`--sweep --dry-run`) could move files and delete the source
  folder when the rated-name destination already existed.
- The rating-freshness cache is no longer marked fresh after a rolled-back or
  failed rename. A successful disk rename + rejected API PUT + rollback used to
  still cache the new rating, so the next sweep skipped the folder for the whole
  TTL — defeating the retry the rollback exists to enable.
- A Sonarr/Radarr outage during `--sweep` now aborts the affected root and
  exits non-zero, instead of mislabeling the entire library as "unknown
  folders" with a falsely-green exit 0. `get_object_by_path` raises
  `ServiceUnavailable` (distinct from "folder not in library").
- Concurrent webhook processes no longer clobber each other's rating-cache
  entries: `_save_rating_cache` re-reads the on-disk cache and merges only the
  keys this process changed (burst imports each ran in their own process and
  the last whole-file writer dropped the others' entries).
- `--clear-cache` now respects `--dry-run` (it deleted the cache file even in a
  preview run; `--refresh` already honored dry-run).
- `_fetch_omdb` no longer caches OMDb's `200 + {"Response":"False"}` error
  bodies (bad key / quota / transient miss), which used to pin a title to
  `N/A` for the rest of a long sweep.
- A non-numeric OMDb rating is coerced to `N/A` rather than leaking into the
  folder suffix (e.g. `[IMDb high]`), which the suffix regex couldn't strip.
- `get_mdl_rating` now tolerates a +/-1 year difference (matching
  `get_mal_rating`) — K-dramas listed on MyDramaList one year off no longer
  fall back to IMDb on a low title-similarity match.
- `slugify` returns its documented `"unknown"` fallback for all-non-Latin
  titles (Korean/Japanese/…) instead of an empty string.
- A failed `desktop.ini` write re-applies System/Hidden to the original file,
  so a transient write error can't leave a previously-hidden `desktop.ini`
  visible as a plain file.
- `rename_folder`'s merge-complete path now logs as `Renamed ...` so the sweep
  summary counts merge-path renames (they were previously undercounted).

### Changed
- `_fs_lock` docstring and contention warning corrected: the helper yields
  `acquired=False` for the caller to decide; the only current caller
  (`_process`) skips. Previous wording claimed "proceeding".
- The OpenSubtitles API's `attributes.url` (the one fully upstream-controlled
  string that reaches the generated `Subtitle.vbs`) is now constrained to the
  `opensubtitles.com` domain before it's trusted — a compromised/malicious
  endpoint can no longer plant an arbitrary URL the user opens by clicking the
  shortcut.
- `--force-refresh` is now threaded through `_process` as a parameter instead
  of temporarily rebinding the module-level `_rating_cache_is_fresh` function.
- `_redact` also strips `api_key`/`apikey` query-param values by name, covering
  encoded/transformed keys it doesn't hold a verbatim copy of.

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
