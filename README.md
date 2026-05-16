# arr_finisher

Finishing touches for Sonarr/Radarr libraries on Windows. After every import
(or during a nightly sweep), each series or movie folder gets a **poster-based
folder icon**, a **rating suffix in the folder name**, a **Links/ subfolder**
of handy web shortcuts, and a **hover tooltip** with the plot summary.

![arr_finisher in action — library view with poster folder icons + rating suffixes, a series folder with its Links subfolder, and the Links subfolder full of shortcuts](Screenshots/example.jpg)

---

## Quick start

```
git clone https://github.com/roko-tech/arr_finisher.git
cd arr_finisher
python -m pip install -r requirements.txt
copy .env.example .env
notepad .env                              :: fill in your API keys
python arr_finisher.py --validate
```

Then in Sonarr/Radarr: **Settings → Connect → Custom Script**, point at
`<repo>\arr_finisher.bat`, and trigger on **On Import** + **On Upgrade**.

---

## What it does

| Output | How |
|---|---|
| **Folder icon from the show's poster** | Drives [maforget/Folder-Icon-Creator](https://github.com/maforget/Folder-Icon-Creator) — the original motivation for this project |
| **Rating suffix in folder name** | `[IMDb 8.6]`, `[MDL 7.5]`, or `[MAL 9.3]` — picks the best source automatically (see below) |
| **`Links/` subfolder with shortcuts** | IMDb, Parents Guide, TVTime, Letterboxd, MyDramaList, MyAnimeList, Twitter, combined Subtitle (SubDL + Subsource + OpenSubtitles) |
| **Explorer tooltip** | OMDb plot summary + rating shown on hover, via `desktop.ini` `InfoTip` |
| **Sonarr/Radarr path sync** | Folder rename is mirrored via API, with rollback on API failure. The rare double-failure case (rollback also fails) is appended to `.rollbacks.log` |

### Rating source auto-detection

- **Korean** content → [MyDramaList](https://mydramalist.com) (via the unofficial [kuryana](https://github.com/tbdsux/kuryana) API)
- **Anime** → [MyAnimeList](https://myanimelist.net) (via [jikan](https://jikan.moe))
- **Everything else** → IMDb (via [OMDb](https://www.omdbapi.com); falls back to scraping `imdb.com` JSON-LD if OMDb is unavailable)

A title-similarity + year filter rejects bad fuzzy matches, so you don't
accidentally end up with the wrong rating for "Bones (2005)" vs some other show.

### How a single import flows

```
Sonarr/Radarr fires On Import / On Upgrade
        │  (sets env vars: Sonarr_Series_* / Radarr_Movie_*)
        ▼
arr_finisher.bat  ─►  python arr_finisher.py
        │
        ├─ detect language / type (env vars first, service API as fallback)
        ├─ fetch rating from MAL / MDL / IMDb (in that priority order)
        ├─ rename folder, appending "[SOURCE X.X]"
        ├─ PUT new path back to Sonarr/Radarr (rollback on API rejection)
        ├─ run Folder-Icon-Creator to generate a .ico from the poster
        ├─ create Links/ subfolder with .lnk + Subtitle.vbs shortcuts
        └─ write plot summary into desktop.ini as Explorer InfoTip
```

Other Sonarr/Radarr events (`Grab`, `Rename`, `Test`, …) are acknowledged and
ignored — only `Download` events trigger the work above.

---

## Install

**Prerequisites:**
- Windows (`pywin32` is used to create `.lnk` shortcut files)
- Python 3.8+
- [maforget/Folder-Icon-Creator](https://github.com/maforget/Folder-Icon-Creator) — download a release, extract (e.g. to `D:\Tools\FolderIconCreator\`), and note the path to `Creator.exe`

**Install the script:**

```
git clone https://github.com/roko-tech/arr_finisher.git
cd arr_finisher
python -m pip install -r requirements.txt
copy .env.example .env
```

Then edit `.env` with your API keys (see [Configure](#configure) below) and verify:

```
python arr_finisher.py --validate
```

---

## Configure

All config lives in `.env` (created from `.env.example`). Real OS environment
variables override the file.

### Required

| Key | What it's for |
|---|---|
| `OMDB_API_KEY` | Free key from [omdbapi.com](https://www.omdbapi.com/apikey.aspx) — needed for any rating fetch |
| `FOLDER_ICON_EXE` | Absolute path to Folder-Icon-Creator's `Creator.exe` (only if `ENABLE_CREATE_FOLDER_ICON` is on — it is, by default) |
| `SONARR_API_KEY` | Sonarr → Settings → General → API Key. Required if you'll trigger from Sonarr |
| `RADARR_API_KEY` | Radarr → Settings → General → API Key. Required if you'll trigger from Radarr |

### Optional

| Key | Default | What it's for |
|---|---|---|
| `SONARR_API_URL` | `http://localhost:8989` | Override if Sonarr isn't on the local host |
| `RADARR_API_URL` | `http://localhost:7878` | Same, for Radarr |
| `SUBDL_API_KEY` | — | Resolves direct SubDL URLs for the Subtitle shortcut (without it, the shortcut falls back to a search URL) |
| `OPENSUBTITLES_API_KEY` | — | Same idea, for OpenSubtitles |
| `SEARCH_LANGUAGE` | `ar` | ISO code (`ar`, `en`, `ja`, …) used by the Twitter hashtag filter and OpenSubtitles search URL |
| `KURYANA_BASE_URL` | `https://kuryana.tbdh.app` | Self-host or alternate mirror |
| `JIKAN_BASE_URL` | `https://api.jikan.moe/v4` | Self-host or alternate mirror |
| `ARR_FINISHER_LOG_LEVEL` | `INFO` | Set to `DEBUG` for permanently-verbose logs |
| `ARR_FINISHER_LOG_DIR` | repo dir | Where to write `arr_finisher.log` (falls back to `%TEMP%`) |
| `ARR_FINISHER_RATING_CACHE_TTL_DAYS` | `7` | How long a rating stays fresh before sweep re-fetches it |
| `ARR_FINISHER_SWEEP_ROOTS` | — | Override sweep roots (otherwise auto-discovered from Sonarr/Radarr). Pipe-separated `path:service` pairs |

`SUBDL_API_KEY` and `OPENSUBTITLES_API_KEY` are listed as optional here, but
`--validate` will still report them as "missing" when unset. That's a heads-up,
not a failure — the rest of the script works without them.

---

## Run

### As a Sonarr/Radarr custom script (primary use)

In Sonarr/Radarr go to **Settings → Connect → Custom Script** and set the path
to `<repo>\arr_finisher.bat`. Trigger on **On Import** and **On Upgrade** (both
fire the same `Download` event). The `.bat` auto-locates the `.py` via `%~dp0`,
so you can keep the repo anywhere.

The `Test` event Sonarr/Radarr fires when you click "Test" in the UI is
acknowledged but does nothing. Other events (`Grab`, `Rename`, …) are
deliberately ignored.

**First-run troubleshooting.** The webhook silently swallows stdout/stderr (so
a crash never breaks the import flow). To make missing config visible, the
script writes an `arr_finisher_setup.txt` next to itself when it detects
missing keys. It's auto-deleted once config is healthy. Example contents:

```
arr_finisher needs configuration before it can do its job.
Last checked: 2026-05-16T11:40:15

Edit `.env` in the repo directory and set:

  - OMDB_API_KEY
      Free key from https://www.omdbapi.com/apikey.aspx — required for any rating fetch.

Then re-trigger the import in Sonarr/Radarr, or run:
  python arr_finisher.py --validate
for a full health check.
```

### Nightly sweep (rating-refresh safety net)

The webhook handles every new import in full. The sweep is a **rating-only**
follow-up that keeps ratings fresh as they evolve over time. Folder icons,
shortcuts, and tooltips are **not** re-touched during sweep — the hook already
did that when the content was imported.

**What the sweep does, per folder:**

1. Looks up the IMDb ID via Sonarr/Radarr
2. Skips if the rating was checked recently (default 7 days; tunable via `ARR_FINISHER_RATING_CACHE_TTL_DAYS`)
3. Otherwise fetches a current rating from the right provider
4. If the rating changed, renames the folder and updates the service path via API (rollback on failure)

A typical nightly sweep on a fully-cached medium library finishes in well
under a minute.

**Scheduling.** The simple option runs daily at 3 AM as `LOCAL SYSTEM`:

```powershell
schtasks /Create /TN "arr_finisher_nightly_sweep" `
    /TR "\"<repo>\arr_finisher_sweep.bat\"" `
    /SC DAILY /ST 03:00 /RL HIGHEST /RU SYSTEM /F
```

If your library lives on a network share, `LOCAL SYSTEM` can't reach mapped
drives. Either use a UNC path everywhere, or switch the task to run as your
user account:

```powershell
schtasks /Create /TN "arr_finisher_nightly_sweep" `
    /TR "\"<repo>\arr_finisher_sweep.bat\"" `
    /SC DAILY /ST 03:00 /RU "%USERNAME%" /F
```

**Tuning.** Set `ARR_FINISHER_RATING_CACHE_TTL_DAYS` in `.env` to change the
freshness window. To force a full re-fetch on the next run:

```
python arr_finisher.py --clear-cache              :: wipe the entire cache
python arr_finisher.py --refresh tt7160070        :: invalidate one IMDb ID
```

**Rollbacks.** When Sonarr/Radarr refuses a path change after the disk has
already been renamed, the script renames the folder back and records the event
in `.rollbacks.log` next to the script. Check that file periodically for any
`FAIL` entries (disk and service out of sync — rare, needs manual fixup).

### Manual / one-off commands

```
python arr_finisher.py --validate                                  # config + connectivity check
python arr_finisher.py --sweep                                     # rating refresh across all roots
python arr_finisher.py --sweep --dry-run                           # preview without touching anything
python arr_finisher.py --service sonarr --path "D:\TV Shows\Foo"   # single folder (looks up env from Sonarr/Radarr)
python arr_finisher.py --clear-cache                               # wipe .rating_cache.json
python arr_finisher.py --refresh tt7160070                         # invalidate one IMDb ID so next sweep re-fetches
python arr_finisher.py --version                                   # print version and exit
python arr_finisher.py --verbose                                   # include DEBUG-level logs in this run
```

`--validate` output looks like this when everything is healthy:

```
=== arr_finisher --validate ===

  [OK ] env  SONARR_API_KEY is set
  [OK ] env  RADARR_API_KEY is set
  [OK ] env  OMDB_API_KEY is set
  [OK ] env  SUBDL_API_KEY is set
  [OK ] env  OPENSUBTITLES_API_KEY is set
  [OK ] module  pywin32 available
  [OK ] file  FolderIconCreator at D:\Tools\FolderIconCreator\Creator.exe
  [OK ] icons  all 10 icon files present
  [OK ] http  Sonarr reachable (200)
  [OK ] http  Radarr reachable (200)
  [OK ] http  OMDb reachable (200)
  [OK ] http  Kuryana reachable (200)
  [OK ] http  Jikan reachable (200)
  [OK ] http  SubDL reachable (200)
  [OK ] http  OpenSubtitles reachable (200)

15 passed, 0 failed
```

### Sweep roots

Resolved in this order, first hit wins:

1. `--roots "D:\Shows:sonarr" "F:\Movies:radarr"` on the CLI (the service is matched on the last colon, so Windows drive letters in the path are fine)
2. `ARR_FINISHER_SWEEP_ROOTS` env var — pipe-separated `path:service` pairs
3. **Auto-discovery** — Sonarr `/api/v3/rootfolder` and Radarr `/api/v3/rootfolder` are queried; their configured root folders become the sweep roots
4. Hardcoded fallback (`D:\TV Shows`, `D:\Anime`, `E:\Movies`) — only used when no service is reachable

Auto-discovery means you usually don't need to configure roots at all if your
Sonarr / Radarr instances are healthy when the sweep runs.

---

## Files arr_finisher writes next to itself

| File | What it is |
|---|---|
| `arr_finisher.log` | Rotating log (1 MB × 3 backups). Relocate with `ARR_FINISHER_LOG_DIR` |
| `.rating_cache.json` | Per-IMDb-ID "last checked" timestamps + last known rating. Powers the sweep TTL |
| `.rollbacks.log` | Append-only journal of rename rollbacks. Rare but worth grepping for `FAIL` |
| `arr_finisher_setup.txt` | Sidecar that appears only when critical config is missing. Auto-deleted when fixed |

`.rating_cache.json` and `arr_finisher_setup.txt` are safe to delete by hand;
the script regenerates them as needed.

---

## Feature toggles

At the top of `arr_finisher.py` — every behavior is independently switchable
(rename, icon, shortcuts, tooltip, MDL/MAL, subtitle combo, etc.). Defaults are
sensible; flip what you don't want.

## Tests

```
tests\run_tests.bat       # fast unit tests (no network, ~0.3 s)
tests\run_tests.bat all   # + integration tests (hits live kuryana + jikan; sets NETWORK_TESTS=1)
```

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for the version history.

## Credits

- **[maforget/Folder-Icon-Creator](https://github.com/maforget/Folder-Icon-Creator)** — the folder-icon generator this whole project is built around. None of this exists without it.
- **[kuryana](https://github.com/tbdsux/kuryana)** — unofficial MyDramaList API.
- **[jikan](https://jikan.moe)** — unofficial MyAnimeList API.
- **[OMDb](https://www.omdbapi.com)** — ratings + plot summaries.
- **[SubDL](https://subdl.com)** and **[OpenSubtitles](https://www.opensubtitles.com)** — direct subtitle URL resolution for the combined Subtitle shortcut.
- Sites linked from the generated shortcuts: **[IMDb](https://www.imdb.com)**, **[Letterboxd](https://letterboxd.com)**, **[TVTime](https://www.tvtime.com)**, **[Subsource](https://subsource.net)**.

## License

MIT — see [LICENSE](LICENSE).
