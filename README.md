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

VN Pathfinder is a standalone Windows desktop application for managing large local collections of RenPy visual novels. Whether you have tens or hundreds of titles, VN Pathfinder gives you a clean, fast interface to see what you have, what you've played, where your save files live, and how much disk space your archives are taking up.

Everything runs locally. No accounts. No telemetry. No internet required (except the optional update check, which can be disabled).

---

## Features

### Library
- **Card-based game list** with cover artwork auto-loaded from each game's assets
- **Played / Unplayed tracking** via save file detection (both `game/saves/` and `%APPDATA%\RenPy\`)
- **Play time counter** — tracked automatically when you launch a game from VN Pathfinder
- **Last played date** and **play count** per game
- **Multi-version grouping** — multiple versions of the same game are grouped under one entry
- **Custom display names** and **notes** per game
- **Tag system** with built-in genre presets, user-promoted custom presets, and a bulk preset manager
- **Search, sort, and filter** by name, played status, and tags
- **Hide games** you don't want in the main list

### Archive Management
- **ZIP extraction queue** — non-modal, runs in the background, fully responsive main window
  - Byte-accurate progress bar with real-time MB/s speed and ETA
  - Sequential processing (one at a time for best disk performance on large files)
  - Cancel mid-extraction between chunks
  - Post-completion **Clear** / **Clear & Delete ZIP** buttons — no race conditions
  - **Clear All & Delete ZIPs** bulk button for batch workflows
- **Redundant archive detection** — identifies ZIPs that already have an extracted counterpart, shows total wasted space, and lets you bulk-delete with a single click
- **RAR patch support** — assign any archive as a patch for a specific game, preview the files to be applied, and apply with overwrite warnings

### Maintenance
- **Clean Orphaned Files** — scans the library root for unrecognised items (partial downloads, leftover folders) and lets you delete them with checkboxes and size display
- **Delete Extracted Archives** — one-click access to all archives that already have an extracted game folder, with a checklist and cumulative size counter

### App
- **Dark theme** tuned for long sessions
- **Persistent user data** saved to `vn_pathfinder.json` in the library folder (notes, tags, playtime, patch assignments, custom presets)
- **Update checker** with opt-out — checks GitHub Releases on startup, shows a clickable banner if a new version is available

---

## Screenshots

> Screenshots will be added before v1.0.0 stable release.

| Library view | Archive queue | Tag editor |
|---|---|---|
| ![Library](docs/screenshots/library.png) | ![Queue](docs/screenshots/queue.png) | ![Tags](docs/screenshots/tags.png) |

---

## Installation

### Option 1 — Pre-built Installer (Recommended)

1. Download **`VNPathfinder_Setup.exe`** from the [latest release](https://github.com/NikoCloud/VN-Pathfinder/releases/latest)
2. Run the installer — it places VN Pathfinder in `Program Files` and creates a Start Menu shortcut
3. Copy (or move) `vn_pathfinder.py` **into your RenPy game library folder** — the app reads games from wherever it lives

> **Important:** VN Pathfinder must be placed in the root of your game library (the folder that contains your game folders and ZIP archives). The installer puts the EXE where you tell it; just make sure you point it at or near your library.

### Option 2 — Portable EXE

1. Download **`VNPathfinder.exe`** from the [latest release](https://github.com/NikoCloud/VN-Pathfinder/releases/latest)
2. Drop it into your RenPy library folder
3. Double-click to run — no installation needed

### Option 3 — Run from Source

```bash
# Prerequisites: Python 3.10+, pip
git clone https://github.com/NikoCloud/VN-Pathfinder.git
cd VN-Pathfinder

pip install Pillow

# Copy vn_pathfinder.py and the assets/ folder into your library folder
# then run:
python vn_pathfinder.py
```

---

## Usage

### First launch
Place `VNPathfinder.exe` (or `vn_pathfinder.py`) in the root of your RenPy library folder — the same folder that contains your game folders and ZIP/RAR archives. On first launch VN Pathfinder scans the directory and populates the library automatically.

### Library tab
| Action | How |
|---|---|
| Launch a game | Select it → **▶ Launch** |
| Mark as played/unplayed | Detail panel → **Mark Played / Unplayed** |
| Add tags | Detail panel → **Edit Tags** |
| Rename a game | Detail panel → edit the display name field |
| Add notes | Detail panel → Notes text box (auto-saves on focus loss) |
| Hide a game | Detail panel → **Hide** |

### Archives tab
| Action | How |
|---|---|
| Extract a ZIP | Select it → **Extract** (added to queue, runs in background) |
| Delete an archive | Select it → **Delete Archive** |
| Delete all already-extracted archives | Toolbar → **Delete Extracted (N) — X GB** |
| Assign a RAR patch to a game | Select it → **Assign as Patch for...** |
| Apply an assigned patch | Select it → **Apply Patch** |

### Extraction queue
The queue window opens automatically when you queue a ZIP. You can:
- Move it to a second monitor while you continue browsing
- Cancel any job (queued or in-progress)
- After completion: **Clear** (remove row) or **Clear & Delete ZIP** (delete source + remove row)
- **Clear All & Delete ZIPs** — process your whole batch in one click

### Tags
- Built-in genre/status presets are available as checkboxes in the tag editor
- Type a custom tag → tick **Save as preset** to permanently add it to your presets
- **★ button** on existing custom tags promotes them to presets
- **Manage Presets...** button opens a bulk editor for all tags used across your library

### Update checking
VN Pathfinder checks for new releases on startup. A clickable banner appears in the status bar if an update is available. To disable, uncheck **"Check for updates"** in the status bar — this preference is saved permanently.

---

## Building from Source

Requires Python 3.10+, PyInstaller, and Inno Setup 6 (for the installer).

```bash
# Install build dependencies
pip install -r requirements.txt

# Build portable EXE + installer (Windows only)
build.bat
```

The script produces:
- `dist/VNPathfinder.exe` — portable single-file executable
- `dist/VNPathfinder_Setup.exe` — full Windows installer (if Inno Setup is found)

---

## Roadmap

- [ ] UI redesign (v2.0)
- [ ] Game cover art from online sources (with opt-in)
- [ ] Import/export library data
- [ ] Bulk tag assignment
- [ ] Series / franchise grouping

---

## License

Apache License 2.0 — see [LICENSE](LICENSE)

Copyright 2025 NikoCloud
