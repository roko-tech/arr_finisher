# arr_finisher

Post-import finishing touches for **Sonarr** and **Radarr** libraries on Windows.
Runs as a custom-script hook on import and as a nightly sweep. For every series
or movie folder it:

- **Appends a rating suffix** to the folder name — `[IMDb 8.6]`, `[MDL 7.5]`, or `[MAL 9.3]`
- **Picks the best rating source automatically**:
  - Korean content → [MyDramaList](https://mydramalist.com) (via the unofficial
    [kuryana](https://github.com/tbdsux/kuryana) API)
  - Anime → [MyAnimeList](https://myanimelist.net) (via [jikan](https://jikan.moe))
  - Everything else → IMDb (via [OMDb](https://www.omdbapi.com) → IMDb JSON-LD scrape as fallback)
- **Creates a `Links/` subfolder** with Windows shortcuts to IMDb, Parents Guide, Twitter,
  TVTime, Letterboxd, MyDramaList, MyAnimeList, and a combined subtitle search (SubDL +
  Subsource + OpenSubtitles). Icons included.
- **Sets the folder tooltip** (`desktop.ini` `InfoTip`) to the OMDb plot summary so the
  description shows on hover in Explorer.
- **Updates the Sonarr/Radarr path via API** to match the renamed folder. Rolls back the
  disk rename if the API refuses the change, so on-disk and in-service paths never drift.

## Quick start

```bash
git clone https://github.com/roko-tech/arr_finisher.git
cd arr_finisher
python -m pip install -r requirements.txt
cp .env.example .env      # then edit .env with your real API keys
python arr_finisher.py --validate
```

### Wire it into Sonarr / Radarr

In **Settings → Connect → Custom Script**, set the path to:

```
<repo>\arr_finisher.bat
```

Trigger on `On Import` (and optionally `On Upgrade`). The .bat auto-locates the Python
script via `%~dp0`, so you can place the repo anywhere.

### Nightly sweep (recommended)

Periodic re-scan that catches anything missed by the hook (webhook failures, manual
file moves, rating changes over time). Schedule via Windows Task Scheduler to run
`arr_finisher_sweep.bat` daily at a quiet hour.

```powershell
schtasks /Create /TN "arr_finisher_nightly_sweep" `
    /TR "\"<repo>\arr_finisher_sweep.bat\"" `
    /SC DAILY /ST 03:00 /RL HIGHEST /RU SYSTEM /F
```

## CLI

```
python arr_finisher.py --validate         # config + connectivity check
python arr_finisher.py --sweep            # process every folder in library roots
python arr_finisher.py --sweep --dry-run  # preview without touching anything
python arr_finisher.py --service sonarr --path "D:\TV Shows\Foo"   # manual single-folder
python arr_finisher.py --verbose          # include DEBUG-level logs
```

Environment variable `ARR_FINISHER_LOG_LEVEL=DEBUG` flips verbose on permanently.

## Configuration (`.env`)

Copy `.env.example` to `.env` and fill in your keys. See the example file for all
available options. Real env vars always win over the `.env` file.

| Key | Purpose | Required |
|---|---|:---:|
| `SONARR_API_KEY` | API key from Sonarr → Settings → General | ✓ |
| `RADARR_API_KEY` | API key from Radarr → Settings → General | ✓ |
| `OMDB_API_KEY` | Free key from [omdbapi.com](https://www.omdbapi.com/apikey.aspx) | ✓ |
| `SUBDL_API_KEY` | [subdl.com/panel/api](https://subdl.com/panel/api) | optional |
| `OPENSUBTITLES_API_KEY` | [opensubtitles.com/consumers](https://www.opensubtitles.com/en/consumers) | optional |
| `SEARCH_LANGUAGE` | ISO code (ar, en, ja, …) for Twitter hashtag + OpenSubtitles listings | optional |
| `FOLDER_ICON_EXE` | Path to [FolderIconCreator.exe](https://github.com/shadoversion/FolderIconCreator) | optional |
| `KURYANA_BASE_URL` | Kuryana mirror override (self-host, etc.) | optional |
| `JIKAN_BASE_URL` | Jikan mirror override | optional |

## Sweep roots

By default the sweep scans:

```python
DEFAULT_SWEEP_ROOTS = [
    (r"D:\TV Shows", "sonarr"),
    (r"D:\Anime",    "sonarr"),
    (r"E:\Movies",   "radarr"),
]
```

Edit the list at the top of the `sweep_library` section of `arr_finisher.py`, or
override per-invocation:

```bash
python arr_finisher.py --sweep --roots "D:\TV Shows:sonarr" "E:\Movies:radarr"
```

## Feature toggles

All at the top of `arr_finisher.py`. Every piece of behavior is independently
switchable — rename, icon, shortcuts, tooltip, rollback, MAL/MDL, etc.

## Running tests

```bash
tests\run_tests.bat       # unit tests (fast, no network)
tests\run_tests.bat all   # + integration tests (hits kuryana and jikan)
```

## Requirements

- Windows (uses `pywin32` for `.lnk` creation and `desktop.ini` attribs)
- Python 3.10+
- Optional: [FolderIconCreator](https://github.com/shadoversion/FolderIconCreator)
  for folder-icon generation from cover art

## License

MIT — see [LICENSE](LICENSE).
