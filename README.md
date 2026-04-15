<div align="center">
  <img src="assets/logo.png" alt="VN Pathfinder Logo" width="120">
  <h1>VN Pathfinder</h1>
  <p>Your complete visual novel library manager — track, organise, launch, and maintain your RenPy collection.</p>

  [![Download](https://img.shields.io/github/v/release/NikoCloud/VN-Pathfinder?label=Download&style=for-the-badge&color=f38ba8)](https://github.com/NikoCloud/VN-Pathfinder/releases/latest)
  [![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue?style=for-the-badge)](https://github.com/NikoCloud/VN-Pathfinder/releases/latest)
  [![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-yellow?style=for-the-badge)](https://www.python.org/)
  [![License: Apache 2.0](https://img.shields.io/badge/License-Apache%202.0-green?style=for-the-badge)](LICENSE)
</div>

---

## Overview

VN Pathfinder is a standalone Windows desktop application for managing large local collections of RenPy visual novels. Whether you have tens or hundreds of titles, VN Pathfinder gives you a clean, fast interface to see what you have, what you've played, where your saves live, and how much disk space your archives are taking up.

Everything runs locally. No accounts. No telemetry. **Zero network access by default** — an optional metadata scraper and update checker are available but locked behind a master switch that is off out of the box.

---

## Features

### Library
- **Card-based game list** with cover artwork auto-loaded from each game's assets — or fetched and stored from online metadata sources
- **Played / Unplayed tracking** via save file detection (`game/saves/` and `%APPDATA%\RenPy\`)
- **Play time counter** tracked automatically when you launch through the app
- **Last played date** and **play count** per game
- **Multi-version grouping** — multiple versions of the same game are grouped under one entry
- **Custom display names** and **notes** per game
- **Tag system** with colour-coded chips, built-in genre/status presets, user-promoted custom presets, and a bulk preset manager
- **Search, sort, and filter** by name, played status, and tags

### Metadata Scraping *(opt-in)*
Fetch rich game information — title, developer, synopsis, cover art, and tags — from multiple sources and pick the best data field by field.

| Source | Login required | What it provides |
|--------|---------------|-----------------|
| **VNDB** | No | Title, developer, synopsis, cover art, tags |
| **F95Zone** | Username + password | Title, developer, synopsis, cover art, tags |
| **LewdCorner** | Username + password | Title, developer, synopsis, cover art, tags |
| **itch.io** | In-app browser | Title, developer, synopsis, cover art |

- **Per-field source picker** — choose which site provides each individual field
- **Cover art** downloaded and stored locally in `.vnpf/` per game — no repeated network calls
- **itch.io browser login** uses an embedded Chromium window so Cloudflare is handled automatically — no copying cookies or opening DevTools

### Archive Management
- **ZIP extraction queue** — non-modal, runs in the background
  - Byte-accurate progress bar with real-time MB/s and ETA
  - Cancel mid-extraction; post-completion **Clear** / **Clear & Delete ZIP** buttons
  - **Clear All & Delete ZIPs** for batch workflows
- **Redundant archive detection** — finds ZIPs that already have an extracted counterpart, shows total wasted space, lets you bulk-delete in one click
- **RAR patch support** — assign any archive as a patch for a specific game, preview files, apply with overwrite warnings

### Maintenance
- **Clean Orphaned Files** — scans library root for unrecognised items, lets you delete with checkboxes and size display
- **Delete Extracted Archives** — checklist of archives that already have an extracted game folder, with cumulative size counter

### Settings
- **Configurable library directory** — point the app at any folder, hot-reloads immediately without a restart
- **LOCKDOWN MODE** — master kill-switch for all network access, **on by default**. One toggle to unlock everything; one toggle to go offline again.
- Per-feature toggles for update checks, metadata scraping, and site logins
- Lockdown indicator in the status bar — click it to open Settings

---

## Screenshots

> Screenshots will be added before v1.0.0 stable release.

---

## Installation

### Option 1 — Installer (Recommended)

1. Download **`VNPathfinder_Setup.exe`** from the [latest release](https://github.com/NikoCloud/VN-Pathfinder/releases/latest)
2. Run the installer — places VN Pathfinder in Program Files with a Start Menu shortcut
3. On first launch, go to **⚙ Settings → General** and set your game library folder

### Option 2 — Portable EXE

1. Download **`VNPathfinder.exe`** from the [latest release](https://github.com/NikoCloud/VN-Pathfinder/releases/latest)
2. Place it **anywhere** — it doesn't need to live inside your games folder
3. On first launch, go to **⚙ Settings → General** and set your game library folder

> **Note:** Windows SmartScreen may warn about an unsigned binary on first run. Click *More info → Run anyway*.

### Option 3 — Run from Source

```bash
git clone https://github.com/NikoCloud/VN-Pathfinder.git
cd VN-Pathfinder
pip install -r requirements.txt
python vn_pathfinder.py
```

---

## Usage

### First launch
Open Settings (⚙ button, top toolbar) → **General** → set your game library folder. The app scans it immediately and populates the library.

### Library tab

| Action | How |
|--------|-----|
| Launch a game | Select it → **▶ Launch** |
| Fetch metadata | Select it → **⬇ Fetch…** in the detail panel |
| Mark as played/unplayed | Detail panel → **Mark Played / Unplayed** |
| Add/edit tags | Detail panel → **Edit Tags** |
| Rename a game | Detail panel → edit the display name field |
| Add notes | Detail panel → Notes text box (auto-saves on focus loss) |
| Hide a game | Detail panel → **Hide** |

### Metadata fetching

1. Select a game → click **⬇ Fetch…** in the detail panel
2. Edit the search query if needed, choose which sources to search, click **Search**
3. In the picker, select the best result from each source using the radio buttons
4. On the right panel, use the dropdowns to choose which source provides each field
5. Click **Save** — cover art is downloaded and stored locally

**Site logins** are managed via the login dialog inside the fetch screen, or directly from **⚙ Settings**:
- **F95Zone / LewdCorner** — enter username and password; a session token is saved locally, your password is never stored
- **itch.io** — click *Log in with browser*, log in normally in the window that opens, it closes automatically when done

### Archives tab

| Action | How |
|--------|-----|
| Extract a ZIP | Select it → **Extract** |
| Delete already-extracted archives | Toolbar → **Delete Extracted (N) — X GB** |
| Assign a RAR patch | Select it → **Assign as Patch for…** |
| Apply a patch | Select it → **Apply Patch** |

### Settings

| Setting | Default | What it does |
|---------|---------|--------------|
| Game library directory | *(set on first launch)* | Folder to scan for games |
| LOCKDOWN MODE | **On** | Master kill-switch — disables all network access |
| App update checks | On | Check GitHub for new releases on startup |
| Metadata fetch | On | Allow VNDB / F95Zone / LewdCorner / itch.io scraping |
| Provider logins | On | Allow site login dialogs |

---

## Requirements

- Windows 10 / 11 (64-bit)
- [Microsoft Edge WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/) — for the itch.io browser login. Already installed on most Windows 11 machines; Windows 10 users may need to install it manually.

---

## Building from Source

Requires Python 3.10+, PyInstaller, and Inno Setup 6 (for the installer).

```bash
pip install -r requirements.txt
build.bat
```

Outputs in `dist/`:
- `VNPathfinder.exe` — portable single-file executable
- `VNPathfinder_Setup.exe` — full Windows installer (if Inno Setup is found)

Releases are built automatically by GitHub Actions on every version tag push.

---

## Roadmap

- [ ] Screenshots in README
- [ ] UI redesign (v2.0)
- [ ] Import/export library data
- [ ] Bulk tag assignment
- [ ] Series / franchise grouping

---

## License

Apache License 2.0 — see [LICENSE](LICENSE)

Copyright 2025 NikoCloud
