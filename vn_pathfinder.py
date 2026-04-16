#!/usr/bin/env python3
"""
VN Pathfinder
Visual novel library manager — track, organise, and launch your RenPy collection.
100 % local — zero network access except optional update checks.

Usage:  python vn_pathfinder.py
"""
from __future__ import annotations

import collections
import datetime
import itertools
import json
import multiprocessing
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import urllib.request
import weakref
import webbrowser
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

try:
    from PIL import Image, ImageTk          # type: ignore[import]
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    # curl_cffi impersonates Chrome's TLS fingerprint — bypasses Cloudflare
    from curl_cffi import requests as _req  # type: ignore[import]
    from bs4 import BeautifulSoup as _BS    # type: ignore[import]
    HAS_SCRAPING = True
except ImportError:
    HAS_SCRAPING = False

try:
    import webview as _webview              # type: ignore[import]
    HAS_WEBVIEW = True
except ImportError:
    HAS_WEBVIEW = False


# ── Version ────────────────────────────────────────────────────────────────────

APP_VERSION  = "1.0.0-beta"
GITHUB_REPO  = "NikoCloud/VN-Pathfinder"
UPDATE_URL   = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
RELEASES_URL = f"https://github.com/{GITHUB_REPO}/releases/latest"

# ── Paths ──────────────────────────────────────────────────────────────────────

def _resource_path(rel: str) -> Path:
    """Resolve asset path for both normal Python run and PyInstaller bundle."""
    base = Path(getattr(sys, "_MEIPASS", None) or Path(__file__).parent)
    return base / rel

# APP_DIR  — folder that contains vn_pathfinder.py / VNPathfinder.exe
#            e.g.  F:\RenPy\_VNPathfinder\
# RENPY_DIR — parent of APP_DIR; the actual games library root
#            e.g.  F:\RenPy\
if getattr(sys, "frozen", False):                     # PyInstaller bundle
    APP_DIR = Path(sys.executable).parent.resolve()
else:                                                  # normal Python run
    APP_DIR = Path(__file__).parent.resolve()

RENPY_DIR     = APP_DIR.parent
APPDATA_RENPY = Path(os.environ.get("APPDATA", "")) / "RenPy"
USERDATA_FILE = APP_DIR / "vn_pathfinder.json"
SETTINGS_FILE = Path(os.environ.get("APPDATA", "")) / "VN Pathfinder" / "settings.json"
ASSETS_DIR    = _resource_path("assets")

# ── Data file migrations ───────────────────────────────────────────────────────
# v1: renpy_manager.json → vn_pathfinder.json
for _old in (APP_DIR / "renpy_manager.json", RENPY_DIR / "renpy_manager.json"):
    if _old.exists() and not USERDATA_FILE.exists():
        try:
            _old.rename(USERDATA_FILE)
        except OSError:
            pass

# v2: vn_pathfinder.json in games root → app subfolder  (post-reorganisation)
_old_root_data = RENPY_DIR / "vn_pathfinder.json"
if _old_root_data.exists() and not USERDATA_FILE.exists():
    try:
        _old_root_data.rename(USERDATA_FILE)
    except OSError:
        pass

# ── Thumbnail / artwork sizes ──────────────────────────────────────────────────

CARD_H        = 82
THUMB_W, THUMB_H          = 80, 54
DETAIL_ART_W, DETAIL_ART_H = 330, 190

# ── Parsing helpers ────────────────────────────────────────────────────────────

PLATFORM_SUFFIXES = {
    "pc", "win", "windows", "linux", "mac",
    "standard", "free", "cracked", "official",
    "ultra", "compressed", "public", "release", "market",
}

VERSION_RE = re.compile(
    r"[-_]"
    r"(v\.?\d[\w.]*|\d+\.\d[\w.]*|\d{3,}|Demo|DEMO|demo"
    r"|Chapter[\w_]*|Day\d+[\w_]*|Act\.[\w.]+|Final|FINAL"
    r"|VER_[\w.]+|Vers\.[\w.]+|Episode[\w_-]*)"
    r"(?:[-_].+)?$",
    re.IGNORECASE,
)

ART_CANDIDATES = [
    # .vnpf/ metadata cover comes first so it takes priority over in-game art
    "../.vnpf/cover.jpg",
    "../.vnpf/cover.png",
    "gui/main_menu.png",
    "gui/main_menu.jpg",
    "gui/game_menu.png",
    "gui/game_menu.jpg",
    "gui/window_icon.png",
]

# ── Metadata scraping constants ────────────────────────────────────────────────

VNDB_API_URL    = "https://api.vndb.org/kana/vn"
F95_BASE        = "https://f95zone.to"
LC_BASE         = "https://lewdcorner.com"
ITCHIO_BASE     = "https://itch.io"
METADATA_DIR    = ".vnpf"
METADATA_FILE   = "metadata.json"
SCRAPER_UA      = f"VN-Pathfinder/{APP_VERSION} (https://github.com/NikoCloud/VN-Pathfinder)"
SCRAPER_TIMEOUT = 15  # seconds

# ── Tag catalogue ──────────────────────────────────────────────────────────────

PRESET_TAGS = [
    "Romance", "Comedy", "Drama", "Horror", "Thriller",
    "Fantasy", "Sci-Fi", "Slice of Life", "Mystery",
    "Completed", "In Progress", "Abandoned", "Favorite", "Want to Play",
    "Short", "Long", "Has Walkthrough",
]

TAG_COLORS: dict[str, str] = {
    "Romance":       "#f38ba8",
    "Comedy":        "#f9e2af",
    "Drama":         "#89dceb",
    "Horror":        "#cba6f7",
    "Thriller":      "#cba6f7",
    "Fantasy":       "#a6e3a1",
    "Sci-Fi":        "#74c7ec",
    "Slice of Life": "#fab387",
    "Mystery":       "#89dceb",
    "Completed":     "#a6e3a1",
    "In Progress":   "#f9e2af",
    "Abandoned":     "#6c7086",
    "Favorite":      "#f38ba8",
    "Want to Play":  "#89b4fa",
    "Short":         "#cdd6f4",
    "Long":          "#cdd6f4",
    "Has Walkthrough": "#f9e2af",
}
DEFAULT_TAG_COLOR = "#89b4fa"

# ── Colour palette ─────────────────────────────────────────────────────────────

BG       = "#1e1e2e"
BG2      = "#181825"
BG3      = "#11111b"
CARD_BG  = "#1e1e2e"
CARD_HOV = "#2a2b3d"
CARD_SEL = "#313244"
FG       = "#cdd6f4"
FG_DIM   = "#6c7086"
FG_MUT   = "#45475a"
SEL      = "#313244"
ACCENT   = "#89b4fa"
GREEN    = "#a6e3a1"
YELLOW   = "#f9e2af"
RED      = "#f38ba8"
MAUVE    = "#cba6f7"
TEAL     = "#94e2d5"

# ══════════════════════════════════════════════════════════════════════════════
# ── Settings (update-check opt-out) ───────────────────────────────────────────

def _version_tuple(v: str) -> tuple[int, ...]:
    v = re.split(r"[-+]", v.lstrip("v"))[0]
    try:
        return tuple(int(x) for x in v.split("."))
    except ValueError:
        return (0,)


def load_settings() -> dict:
    try:
        with open(SETTINGS_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(data: dict) -> None:
    SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except OSError:
        pass


def _set_lib_dir(path: Path) -> None:
    """Override the global library directory at runtime (called after settings load)."""
    global RENPY_DIR
    RENPY_DIR = path


# Settings key defaults — used in SettingsDialog and net_ok()
_SETTINGS_DEFAULTS: dict[str, object] = {
    "check_updates":        True,
    "fetch_metadata":       True,
    "allow_provider_login": True,
    "allow_download_links": True,
    "lockdown":             True,   # ON by default; persists once user disables it
}


# ══════════════════════════════════════════════════════════════════════════════
# DATA CLASSES
# ══════════════════════════════════════════════════════════════════════════════

@dataclass
class GameVersion:
    folder_name: str
    folder_path: Path
    base_key: str
    version_str: str
    display_name: str
    exe_path: Path | None
    local_save_dir: Path
    appdata_save_dir: Path | None = None
    metadata: dict             = field(default_factory=dict)
    is_renpy: bool             = True


@dataclass
class MetadataCandidate:
    """A single search result returned by a provider before the user picks one."""
    source:          str          # "vndb" | "f95zone" | "lewdcorner"
    title:           str
    url:             str
    developer:       str       = ""
    synopsis:        str       = ""
    cover_url:       str       = ""
    tags:            list[str] = field(default_factory=list)
    screenshot_urls: list[str] = field(default_factory=list)   # additional images
    extra:           dict      = field(default_factory=dict)   # source-specific raw fields


@dataclass
class MetadataResult:
    """Aggregated result for one source after the user selects a candidate."""
    source:    str
    title:     str
    url:       str
    developer: str = ""
    synopsis:  str = ""
    cover_url: str = ""
    tags:      list[str] = field(default_factory=list)
    extra:     dict      = field(default_factory=dict)


@dataclass
class Archive:
    archive_path: Path
    base_key: str
    version_str: str
    matched_folder: str | None = None


@dataclass
class GameGroup:
    base_key: str
    display_name: str
    versions: list[GameVersion] = field(default_factory=list)
    archives: list[Archive]    = field(default_factory=list)


@dataclass
class UserData:
    # ── v1 fields ──────────────────────────────────────────────────────────
    notes:               dict[str, str]  = field(default_factory=dict)
    hidden:              set[str]        = field(default_factory=set)
    manual_played:       set[str]        = field(default_factory=set)
    manual_unplayed:     set[str]        = field(default_factory=set)
    custom_display_names:dict[str, str]  = field(default_factory=dict)
    # ── v2 fields ──────────────────────────────────────────────────────────
    playtime:  dict[str, int]        = field(default_factory=dict)   # folder_name → s
    last_played: dict[str, str]      = field(default_factory=dict)   # folder_name → ISO
    play_count: dict[str, int]       = field(default_factory=dict)   # folder_name → N
    tags:      dict[str, list[str]]  = field(default_factory=dict)   # base_key → [tag]
    custom_art: dict[str, str]       = field(default_factory=dict)   # base_key → path
    patch_assignments: dict[str, str] = field(default_factory=dict)  # archive_name → base_key
    custom_presets:   list[str]       = field(default_factory=list)   # user-promoted preset tags


# ══════════════════════════════════════════════════════════════════════════════
# NAME PARSING
# ══════════════════════════════════════════════════════════════════════════════

def is_game_dir(path: Path) -> bool:
    """True for RenPy game folders (game/ subdir or *.py launcher)."""
    if not path.is_dir():
        return False
    if (path / "game").is_dir():
        return True
    return any(path.glob("*.py"))


def is_exe_game_dir(path: Path) -> bool:
    """True for non-RenPy folders that contain a .exe at root or one level deep."""
    if not path.is_dir():
        return False
    if is_game_dir(path):
        return False   # already handled as RenPy
    # Direct .exe in the folder
    if any(path.glob("*.exe")):
        return True
    # One subfolder that itself contains a .exe (common bundled structure)
    try:
        subdirs = [d for d in path.iterdir() if d.is_dir()]
        if len(subdirs) == 1:
            return any(subdirs[0].glob("*.exe"))
    except OSError:
        pass
    return False


def _strip_platform(tokens: list[str]) -> list[str]:
    while tokens and tokens[-1].lower() in PLATFORM_SUFFIXES:
        tokens.pop()
    return tokens


def _camel_split(s: str) -> str:
    s = re.sub(r"[-_]+", " ", s)
    s = re.sub(r"([a-z])([A-Z])", r"\1 \2", s)
    return re.sub(r" {2,}", " ", s).strip()


def parse_folder_name(name: str) -> tuple[str, str, str]:
    """Return (base_key, version_str, display_name)."""
    name_clean = re.sub(r"^\[[^\]]+\]\s*", "", name)
    m = VERSION_RE.search(name_clean)
    version_str = ""
    base_raw = name_clean
    if m:
        version_str = m.group(1)
        base_raw = name_clean[: m.start()]
    tokens = [t for t in re.split(r"[-_\s]+", base_raw) if t]
    tokens = _strip_platform(tokens)
    if not version_str:
        tokens = _strip_platform(tokens)
    base_for_display = " ".join(tokens)
    display_name = _camel_split(base_for_display)
    base_key = re.sub(r"[-_\s]+", "", base_for_display).lower()
    return base_key, version_str, display_name.strip() or name


def parse_version_tuple(v: str) -> tuple:
    if not v:
        return (-2,)
    if v.lower() in ("demo", "final"):
        return (-1,)
    nums = re.findall(r"\d+", v)
    return tuple(int(n) for n in nums) if nums else (-1,)


# ══════════════════════════════════════════════════════════════════════════════
# SAVE DETECTION
# ══════════════════════════════════════════════════════════════════════════════

def read_save_dir_from_options(game_path: Path) -> str | None:
    opt = game_path / "game" / "options.rpy"
    if not opt.exists():
        return None
    try:
        text = opt.read_text(encoding="utf-8", errors="ignore")
        m = re.search(r'config\.save_directory\s*=\s*["\']([^"\']+)["\']', text)
        return m.group(1) if m else None
    except OSError:
        return None


def find_appdata_save_dir(base_key: str, appdata_renpy: Path) -> Path | None:
    if not appdata_renpy.exists():
        return None
    try:
        candidates = [d for d in appdata_renpy.iterdir() if d.is_dir()]
    except OSError:
        return None
    for c in candidates:
        stripped = re.sub(r"[-_]\d{6,}$", "", c.name)
        ck = re.sub(r"[-_\s]+", "", stripped).lower()
        if ck == base_key:
            return c
        ml = min(len(base_key), len(ck), 5)
        if ml >= 5 and ck.startswith(base_key[:ml]):
            return c
    return None


def resolve_appdata(game_path: Path, base_key: str) -> Path | None:
    name = read_save_dir_from_options(game_path)
    if name:
        exact = APPDATA_RENPY / name
        if exact.exists():
            return exact
    return find_appdata_save_dir(base_key, APPDATA_RENPY)


def _has_save_files(directory: Path) -> bool:
    try:
        return any(
            f.suffix == ".save" or f.name == "persistent"
            for f in directory.iterdir() if f.is_file()
        )
    except OSError:
        return False


def detect_played(version: GameVersion, ud: UserData) -> bool:
    n = version.folder_name
    if n in ud.manual_unplayed:
        return False
    if n in ud.manual_played:
        return True
    if version.local_save_dir.exists():
        try:
            if any(version.local_save_dir.iterdir()):
                return True
        except OSError:
            pass
    if version.appdata_save_dir and version.appdata_save_dir.exists():
        if _has_save_files(version.appdata_save_dir):
            return True
    return False


def saves_location(version: GameVersion, ud: UserData) -> str:
    locs = []
    if version.local_save_dir.exists():
        try:
            if any(version.local_save_dir.iterdir()):
                locs.append("Local")
        except OSError:
            pass
    if version.appdata_save_dir and version.appdata_save_dir.exists():
        if _has_save_files(version.appdata_save_dir):
            locs.append("AppData")
    return " + ".join(locs) if locs else "—"


# ══════════════════════════════════════════════════════════════════════════════
# SCANNER & GROUPER
# ══════════════════════════════════════════════════════════════════════════════

def scan_game_version(path: Path) -> GameVersion | None:
    renpy = is_game_dir(path)
    non_renpy = (not renpy) and is_exe_game_dir(path)
    if not renpy and not non_renpy:
        return None
    base_key, version_str, display_name = parse_folder_name(path.name)
    if not base_key:
        return None
    exe_path: Path | None = None
    for p in sorted(path.glob("*.exe")):
        if "-32" not in p.stem:
            exe_path = p
            break
    # For non-RenPy games bundled one folder deep, find the exe there
    if not exe_path and not renpy:
        try:
            subdirs = [d for d in path.iterdir() if d.is_dir()]
            if len(subdirs) == 1:
                for p in sorted(subdirs[0].glob("*.exe")):
                    if "-32" not in p.stem:
                        exe_path = p
                        break
        except OSError:
            pass
    return GameVersion(
        folder_name=path.name,
        folder_path=path,
        base_key=base_key,
        version_str=version_str,
        display_name=display_name,
        exe_path=exe_path,
        local_save_dir=path / "game" / "saves",
        appdata_save_dir=resolve_appdata(path, base_key) if renpy else None,
        metadata=load_game_metadata(path),
        is_renpy=renpy,
    )


def scan_archive(path: Path) -> Archive:
    base_key, version_str, _ = parse_folder_name(path.stem)
    return Archive(archive_path=path, base_key=base_key, version_str=version_str)


def scan_all() -> list[GameGroup]:
    versions: list[GameVersion] = []
    archives: list[Archive] = []
    try:
        entries = list(RENPY_DIR.iterdir())
    except OSError:
        return []
    for e in entries:
        if e.is_dir():
            v = scan_game_version(e)
            if v:
                versions.append(v)
        elif e.suffix.lower() in (".zip", ".rar"):
            archives.append(scan_archive(e))
    return build_groups(versions, archives)


def build_groups(
    versions: list[GameVersion], archives: list[Archive]
) -> list[GameGroup]:
    groups: dict[str, GameGroup] = {}
    for v in versions:
        k = v.base_key
        if k not in groups:
            groups[k] = GameGroup(base_key=k, display_name=v.display_name)
        g = groups[k]
        g.versions.append(v)
        if len(v.display_name) > len(g.display_name):
            g.display_name = v.display_name
    for g in groups.values():
        g.versions.sort(key=lambda x: parse_version_tuple(x.version_str))

    for a in archives:
        k = a.base_key
        if k in groups:
            g = groups[k]
            for v in g.versions:
                if v.version_str and v.version_str == a.version_str:
                    a.matched_folder = v.folder_name
                    break
            if not a.matched_folder and g.versions:
                a.matched_folder = g.versions[-1].folder_name
            g.archives.append(a)
        else:
            _, _, disp = parse_folder_name(a.archive_path.stem)
            groups[k] = GameGroup(base_key=k, display_name=disp, archives=[a])

    return sorted(groups.values(), key=lambda g: g.display_name.lower())


def _share_appdata(groups: list[GameGroup]) -> None:
    for g in groups:
        if len(g.versions) < 2:
            continue
        resolved = next((v.appdata_save_dir for v in g.versions
                         if v.appdata_save_dir), None)
        if resolved:
            for v in g.versions:
                if not v.appdata_save_dir:
                    v.appdata_save_dir = resolved


# ══════════════════════════════════════════════════════════════════════════════
# PERSISTENCE
# ══════════════════════════════════════════════════════════════════════════════

def load_userdata() -> UserData:
    if not USERDATA_FILE.exists():
        return UserData()
    try:
        with open(USERDATA_FILE, encoding="utf-8") as f:
            raw = json.load(f)
        return UserData(
            notes=raw.get("notes", {}),
            hidden=set(raw.get("hidden", [])),
            manual_played=set(raw.get("manual_played", [])),
            manual_unplayed=set(raw.get("manual_unplayed", [])),
            custom_display_names=raw.get("custom_display_names", {}),
            playtime=raw.get("playtime", {}),
            last_played=raw.get("last_played", {}),
            play_count=raw.get("play_count", {}),
            tags=raw.get("tags", {}),
            custom_art=raw.get("custom_art", {}),
            patch_assignments=raw.get("patch_assignments", {}),
            custom_presets=raw.get("custom_presets", []),
        )
    except (OSError, json.JSONDecodeError):
        return UserData()


def save_userdata(ud: UserData) -> None:
    try:
        with open(USERDATA_FILE, "w", encoding="utf-8") as f:
            json.dump({
                "version": 2,
                "notes": ud.notes,
                "hidden": sorted(ud.hidden),
                "manual_played": sorted(ud.manual_played),
                "manual_unplayed": sorted(ud.manual_unplayed),
                "custom_display_names": ud.custom_display_names,
                "playtime": ud.playtime,
                "last_played": ud.last_played,
                "play_count": ud.play_count,
                "tags": ud.tags,
                "custom_art": ud.custom_art,
                "patch_assignments": ud.patch_assignments,
                "custom_presets": ud.custom_presets,
            }, f, indent=2)
    except OSError as e:
        messagebox.showerror("Save Error", f"Could not save data:\n{e}")


# ══════════════════════════════════════════════════════════════════════════════
# UTILITY FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def fmt_time(seconds: int) -> str:
    if seconds <= 0:
        return ""
    h, rem = divmod(seconds, 3600)
    m = rem // 60
    if h:
        return f"{h}h {m:02d}m"
    return f"{m}m"


def fmt_date(iso: str) -> str:
    try:
        dt = datetime.datetime.fromisoformat(iso)
        delta = datetime.datetime.now() - dt
        if delta.days == 0:
            return "Today"
        if delta.days == 1:
            return "Yesterday"
        if delta.days < 7:
            return f"{delta.days}d ago"
        if delta.days < 30:
            return f"{delta.days // 7}w ago"
        return dt.strftime("%b %d, %Y")
    except ValueError:
        return iso[:10]


def find_art_path(game_path: Path) -> Path | None:
    game_dir = game_path / "game"
    for rel in ART_CANDIDATES:
        p = game_dir / rel
        if p.exists() and p.stat().st_size > 1000:
            return p
    gui = game_dir / "gui"
    if gui.exists():
        try:
            imgs = [
                p for ext in ("*.png", "*.jpg", "*.jpeg")
                for p in gui.glob(ext)
            ]
            biggest = max(imgs, key=lambda p: p.stat().st_size, default=None)
            if biggest and biggest.stat().st_size > 20_000:
                return biggest
        except OSError:
            pass
    return None


def _group_art_path(g: GameGroup, ud: UserData) -> Path | None:
    ca = ud.custom_art.get(g.base_key)
    if ca:
        p = Path(ca)
        if p.exists():
            return p
    # .vnpf/ cover takes priority over in-game art
    for v in reversed(g.versions):
        for ext in ("cover.jpg", "cover.png"):
            vnpf_cover = _vnpf_dir(v.folder_path) / ext
            if vnpf_cover.exists():
                return vnpf_cover
    for v in reversed(g.versions):
        p = find_art_path(v.folder_path)
        if p:
            return p
    return None


def _group_carousel_paths(g: GameGroup, ud: UserData) -> list[Path]:
    """Return ordered list of image paths for the carousel.
    First entry is the cover (same as _group_art_path); subsequent entries are
    screenshots stored in .vnpf/screenshot_N.* — sorted by filename."""
    cover = _group_art_path(g, ud)
    paths: list[Path] = [cover] if cover else []
    for v in reversed(g.versions):
        vnpf = _vnpf_dir(v.folder_path)
        shots = sorted(vnpf.glob("screenshot_*.*"),
                       key=lambda p: p.stem)
        for s in shots:
            if s not in paths:
                paths.append(s)
    return paths


def _pil_load(path: Path, w: int, h: int):
    """Load + resize with PIL. Must be called from background thread."""
    with Image.open(path) as img:
        img = img.convert("RGB")
        img.thumbnail((w, h), Image.LANCZOS)
        # Centre on solid background
        bg = Image.new("RGB", (w, h), (int(BG2[1:3], 16),
                                       int(BG2[3:5], 16),
                                       int(BG2[5:7], 16)))
        x = (w - img.width) // 2
        y = (h - img.height) // 2
        bg.paste(img, (x, y))
        return bg.copy()


# ══════════════════════════════════════════════════════════════════════════════
# METADATA HELPERS  (.vnpf/ storage)
# ══════════════════════════════════════════════════════════════════════════════

def _vnpf_dir(folder_path: Path) -> Path:
    return folder_path / METADATA_DIR


def load_game_metadata(folder_path: Path) -> dict:
    """Read .vnpf/metadata.json; return empty dict on any error."""
    try:
        with open(_vnpf_dir(folder_path) / METADATA_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, json.JSONDecodeError):
        return {}


def save_game_metadata(folder_path: Path, meta: dict) -> None:
    d = _vnpf_dir(folder_path)
    d.mkdir(exist_ok=True)
    with open(d / METADATA_FILE, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)


def _download_image(url: str, dest: Path, timeout: int = SCRAPER_TIMEOUT) -> bool:
    """Download a single image URL to dest. Returns True on success."""
    try:
        resp = _req.get(url, impersonate="chrome131", timeout=timeout)
        resp.raise_for_status()
        dest.parent.mkdir(exist_ok=True)
        with open(dest, "wb") as fh:
            fh.write(resp.content)
        return True
    except Exception:
        return False


def _download_cover(url: str, folder_path: Path) -> Path | None:
    """Download cover art to .vnpf/cover.jpg (or .png). Returns saved path."""
    if not HAS_SCRAPING or not url:
        return None
    ext = ".png" if url.lower().endswith(".png") else ".jpg"
    dest = _vnpf_dir(folder_path) / f"cover{ext}"
    return dest if _download_image(url, dest) else None


def _download_screenshots(urls: list[str], folder_path: Path) -> list[str]:
    """Download screenshot images to .vnpf/screenshot_N.jpg/.png in parallel.
    Returns list of filenames that were saved (relative to .vnpf/), sorted."""
    if not HAS_SCRAPING or not urls:
        return []
    vnpf = _vnpf_dir(folder_path)
    vnpf.mkdir(exist_ok=True)
    # Remove stale screenshots from a previous fetch
    for old in vnpf.glob("screenshot_*"):
        try:
            old.unlink()
        except OSError:
            pass

    items = list(enumerate(urls[:18], start=1))
    results: dict[int, str] = {}   # idx → filename if saved

    def _dl(i: int, url: str) -> None:
        ext = ".png" if url.lower().endswith(".png") else ".jpg"
        dest = vnpf / f"screenshot_{i}{ext}"
        if _download_image(url, dest, timeout=8):
            results[i] = dest.name

    threads = [threading.Thread(target=_dl, args=(i, u), daemon=True)
               for i, u in items]
    for t in threads:
        t.start()
    for t in threads:
        t.join(timeout=10)   # overall cap regardless of individual timeouts

    return [results[i] for i in sorted(results)]


def _strip_bbcode(text: str) -> str:
    """Remove common BBCode tags from a string."""
    text = re.sub(r"\[/?(?:b|i|u|s|url|img|color|size|spoiler|quote)[^\]]*\]",
                  "", text, flags=re.IGNORECASE)
    return text.strip()


# ══════════════════════════════════════════════════════════════════════════════
# METADATA PROVIDERS
# ══════════════════════════════════════════════════════════════════════════════

# ── Per-site cookie storage ────────────────────────────────────────────────────

COOKIE_FILE = Path(os.environ.get("APPDATA", "")) / "VN Pathfinder" / "cookies.json"

def load_site_cookies(site: str) -> dict:
    """Return saved cookies for a site key, or empty dict."""
    try:
        with open(COOKIE_FILE, encoding="utf-8") as f:
            return json.load(f).get(site, {})
    except (OSError, json.JSONDecodeError):
        return {}

def save_site_cookies(site: str, cookies: dict) -> None:
    COOKIE_FILE.parent.mkdir(parents=True, exist_ok=True)
    try:
        existing: dict = {}
        try:
            with open(COOKIE_FILE, encoding="utf-8") as f:
                existing = json.load(f)
        except (OSError, json.JSONDecodeError):
            pass
        existing[site] = cookies
        with open(COOKIE_FILE, "w", encoding="utf-8") as f:
            json.dump(existing, f, indent=2)
    except OSError:
        pass

def _cffi_get(url: str, cookies: dict | None = None,
              timeout: int = SCRAPER_TIMEOUT) -> "object":
    """GET via curl_cffi impersonating Chrome. Raises on non-2xx.
    NOTE: do NOT pass a custom User-Agent — impersonate= sets the correct
    Chrome UA automatically. Overriding it breaks the Cloudflare bypass.
    """
    return _req.get(
        url,
        impersonate="chrome131",
        cookies=cookies or {},
        timeout=timeout,
    )

def _cffi_post(url: str, json_body: dict,
               timeout: int = SCRAPER_TIMEOUT) -> "object":
    """POST via curl_cffi impersonating Chrome."""
    return _req.post(
        url,
        impersonate="chrome131",
        json=json_body,
        headers={"User-Agent": SCRAPER_UA, "Content-Type": "application/json"},
        timeout=timeout,
    )


# ══════════════════════════════════════════════════════════════════════════════
# METADATA PROVIDERS
# ══════════════════════════════════════════════════════════════════════════════

class VNDBProvider:
    """Query VNDB public API v2 (no auth required)."""

    SOURCE = "vndb"

    def search(self, title: str) -> list[MetadataCandidate]:
        if not HAS_SCRAPING:
            raise RuntimeError("curl_cffi/bs4 not installed")
        payload = {
            "filters": ["search", "=", title],
            "fields": (
                "title,alttitle,released,developers.name,"
                "description,image.url,image.sexual,tags.name,tags.spoiler"
            ),
            "results": 10,
            "sort": "searchrank",
        }
        resp = _cffi_post(VNDB_API_URL, payload)
        resp.raise_for_status()
        data = resp.json()
        results: list[MetadataCandidate] = []
        for item in data.get("results", []):
            vid   = item.get("id", "")
            url   = f"https://vndb.org/{vid}"
            devs  = [d["name"] for d in item.get("developers", []) if d.get("name")]
            # Filter out spoiler tags and keep only non-sexual images
            tags  = [t["name"] for t in item.get("tags", [])
                     if t.get("name") and not t.get("spoiler")]
            raw_synopsis = _strip_bbcode(item.get("description") or "")
            cover = ""
            img = item.get("image") or {}
            if img.get("url"):
                cover = img["url"]
            results.append(MetadataCandidate(
                source=self.SOURCE,
                title=item.get("title", ""),
                url=url,
                developer=", ".join(devs),
                synopsis=raw_synopsis,
                cover_url=cover,
                tags=tags,
                extra={"vndb_id": vid, "released": item.get("released", "")},
            ))
        return results


class F95ZoneProvider:
    """
    Scrape F95Zone using curl_cffi (Chrome TLS impersonation) + session cookies.
    Users must supply their xf_user + xf_session cookies from a logged-in browser
    session for adult content to be visible.
    """

    SOURCE = "f95zone"
    # XenForo search endpoint
    _SEARCH_URL = f"{F95_BASE}/search/search/"

    def _cookies(self) -> dict:
        return load_site_cookies("f95zone")

    def search(self, title: str) -> list[MetadataCandidate]:
        if not HAS_SCRAPING:
            raise RuntimeError("curl_cffi/bs4 not installed")
        cookies = self._cookies()
        if not cookies:
            raise RuntimeError(
                "Not logged in to F95Zone.\n"
                "Use the '⬇ Fetch…' dialog → 'Configure site login' to log in.")
        # XenForo search: POST to /search/search/
        from urllib.parse import urlencode
        params = urlencode({
            "keywords": title,
            "t":        "post",
            "o":        "relevance",
            "c[nodes][0]": "2",          # Games/Comics node
            "c[child_nodes]": "1",
            "c[title_only]": "1",
        })
        url = f"{self._SEARCH_URL}?{params}"
        resp = _cffi_get(url, cookies=cookies)
        resp.raise_for_status()
        soup = _BS(resp.text, "lxml")

        # Detect login wall
        if soup.select_one("form.js-loginForm, .blockMessage--login"):
            raise RuntimeError(
                "F95Zone session expired. Log in again via '⬇ Fetch…' → "
                "'Configure site login'.")

        results: list[MetadataCandidate] = []
        for item in soup.select("li.block-row")[:8]:
            a = item.select_one("h3.contentRow-title a, .contentRow-title a")
            if not a:
                continue
            href = a.get("href", "")
            if not href.startswith("http"):
                href = F95_BASE + href
            snippet_el = item.select_one(".contentRow-snippet")
            excerpt = snippet_el.get_text(" ", strip=True) if snippet_el else ""
            # Strip F95Zone's [Engine][Version][Dev] bracket convention from
            # search-result titles so cards show a clean name
            raw = a.get_text(separator=" ", strip=True)
            clean = re.sub(r"\s*\[[^\]]+\]", "", raw).strip() or raw
            results.append(MetadataCandidate(
                source=self.SOURCE,
                title=clean,
                url=href,
                synopsis=excerpt,
            ))
        return results

    def fetch_thread(self, url: str) -> MetadataCandidate:
        cookies = self._cookies()
        resp = _cffi_get(url, cookies=cookies)
        resp.raise_for_status()
        soup = _BS(resp.text, "lxml")

        # Title — F95Zone format: "[Engine] Title [Version] [Developer]"
        # The H1 contains a child <a class="labelLink"> with the engine/category
        # label ("Ren'Py", "RPGM", …).  get_text() with no separator would jam
        # that label directly onto the game name, giving "Ren'PyGame Name".
        # Fix: collect only the direct NavigableString children of the H1,
        # which contain the actual title text (the <a> child's text is skipped).
        from bs4 import NavigableString as _NST
        raw_title = ""
        title_el = soup.select_one("h1.p-title-value")
        if title_el:
            parts = [str(t).strip() for t in title_el.children
                     if isinstance(t, _NST) and str(t).strip()]
            raw_title = " ".join(parts)
            if not raw_title:           # fallback: separator-joined full text
                raw_title = title_el.get_text(separator=" ", strip=True)
        # Extract developer from last bracket pair
        developer = ""
        dev_m = re.search(r"\[([^\]]+)\]\s*$", raw_title)
        if dev_m:
            developer = dev_m.group(1)
        # Clean title: strip all [brackets]
        clean_title = re.sub(r"\s*\[[^\]]+\]", "", raw_title).strip()

        # First post body — try multiple selectors (F95 structure varies)
        body_el = (
            soup.select_one("article.message--first .bbWrapper") or
            soup.select_one(".message-body .bbWrapper") or
            soup.select_one(".bbWrapper")
        )
        # Scope image extraction to the first post only (prevents picking up
        # images from replies, signature blocks, and user avatars).
        first_post_el = (
            soup.select_one("article.message--first") or
            soup.select_one(".message--first") or
            body_el
        )
        first_post_html = str(first_post_el) if first_post_el else ""

        synopsis = cover_url = ""
        tags: list[str] = []
        screenshot_urls: list[str] = []

        # Images: F95Zone lazy-loads via JS — img[data-src] is injected at runtime
        # and does NOT appear in BeautifulSoup's parsed tree.  Instead, extract
        # attachment URLs directly from the raw HTML with a regex.
        # Full-size attachments never have '/thumb/' in the path; thumbnail copies do.
        _seen: set[str] = set()
        _att_pat = re.compile(r'https://attachments\.f95zone\.to/[^\s"\'<>]+')
        for u in _att_pat.findall(first_post_html):
            if '/thumb/' in u or u in _seen:
                continue
            # Skip common noise (emoticons, awards badges, etc.)
            if any(x in u for x in ('smilies', 'emoji', '/awards/')):
                continue
            _seen.add(u)
            if not cover_url:
                cover_url = u
            else:
                screenshot_urls.append(u)
            if len(screenshot_urls) >= 18:          # 1 cover + 18 = 19 total
                break
        # Fallback: external images (imgur, etc.) referenced via data-src in raw HTML
        if not cover_url:
            _ext_pat = re.compile(
                r'data-src="(https://(?!attachments\.f95zone)[^\s"<>]+\.'
                r'(?:jpg|jpeg|png|webp|gif)(?:\?[^\s"<>]*)?)"',
                re.IGNORECASE,
            )
            for u in _ext_pat.findall(first_post_html):
                if any(x in u for x in ('smilies', 'emoji')) or u in _seen:
                    continue
                _seen.add(u)
                if not cover_url:
                    cover_url = u
                else:
                    screenshot_urls.append(u)
                if len(screenshot_urls) >= 18:
                    break

        if body_el:
            # Synopsis: first paragraph with real text
            for el in body_el.select("p, div"):
                t = el.get_text(" ", strip=True)
                if len(t) > 80 and not t.startswith("["):
                    synopsis = t
                    break
            # Developer: look for "Developer:" label in post body as a more
            # reliable source than the title bracket (title may omit it)
            if not developer:
                for bold in body_el.select("b, strong"):
                    if "developer" in bold.get_text(strip=True).lower():
                        # Text immediately after the label (strip icon links)
                        parent = bold.parent
                        if parent:
                            raw = parent.get_text(" ", strip=True)
                            m = re.search(
                                r"[Dd]eveloper\s*[:\-]\s*([^\n\r|/\\]{2,60})",
                                raw)
                            if m:
                                # Trim trailing link-noise like "Itch.io Instagram…"
                                dev_raw = m.group(1).strip()
                                developer = re.split(
                                    r"\s{2,}|\s(?:Itch|Patreon|Twitter|Discord"
                                    r"|Instagram|Subscribestar|Boosty|Buy\b)",
                                    dev_raw)[0].strip()
                        break

        # Tags — try header tag links first, then Genre field in post body
        # XenForo can put tagItem on the <li> OR the <a> depending on theme
        for sel in (
            ".p-description a.tagItem",
            "li.tagItem a",
            "a.tagItem",
            "ul.tagList li a",
            ".tagList a",
        ):
            found = [el.get_text(strip=True) for el in soup.select(sel)
                     if el.get_text(strip=True)]
            if found:
                tags = found
                break

        # Fallback: Genre field in the post body
        # F95 formats:
        #   <b>Genre</b>: [SPOILER]3dcg, Female protagonist, ...[/SPOILER]
        #   <b>Genre</b>: 3dcg, Female protagonist, ... (no spoiler)
        if not tags and body_el:
            for bold in body_el.select("b, strong"):
                if "genre" in bold.get_text(strip=True).lower():
                    # Strategy 1: Genre in a spoiler block
                    spoiler = bold.find_next(
                        "div", class_=lambda c: c and "bbCodeSpoiler" in c)
                    if spoiler:
                        # Use the inner content div — outer div includes button text ("Spoiler")
                        inner = (
                            spoiler.select_one(".bbCodeSpoiler-content") or
                            spoiler.select_one(".bbCodeBlock-content") or
                            spoiler
                        )
                        raw = inner.get_text(", ", strip=True)
                        tags = [t.strip() for t in re.split(r"[,\n]+", raw) if t.strip()]

                    # Strategy 2: Genre as inline text on the same line (no spoiler)
                    if not tags:
                        parent = bold.parent
                        if parent:
                            full = parent.get_text(" ", strip=True)
                            m = re.search(r"[Gg]enre\s*[:\-]\s*(.+)", full)
                            if m:
                                raw = m.group(1).strip()
                                tags = [t.strip() for t in re.split(r"[,\n]+", raw)
                                        if t.strip()]
                    break

        return MetadataCandidate(
            source=self.SOURCE,
            title=clean_title,
            url=url,
            developer=developer,
            synopsis=synopsis,
            cover_url=cover_url,
            tags=tags,
            screenshot_urls=screenshot_urls,
        )


class LewdCornerProvider:
    """
    Scrape LewdCorner.com — runs XenForo, same structure as F95Zone.
    """

    SOURCE = "lewdcorner"
    _SEARCH_URL = f"{LC_BASE}/search/search/"

    def _cookies(self) -> dict:
        return load_site_cookies("lewdcorner")

    def search(self, title: str) -> list[MetadataCandidate]:
        if not HAS_SCRAPING:
            raise RuntimeError("curl_cffi/bs4 not installed")
        cookies = self._cookies()
        if not cookies:
            raise RuntimeError(
                "Not logged in to Lewd Corner.\n"
                "Use '⬇ Fetch…' → 'Configure site login' to log in.")
        from urllib.parse import urlencode
        params = urlencode({
            "keywords":        title,
            "t":               "post",
            "o":               "relevance",
            "c[nodes][0]":     "6",   # LC "Games" node — parent of all game sub-forums
            "c[child_nodes]":  "1",   # recursively includes Games+, AI Games, Ports, etc.
            "c[title_only]":   "1",
        })
        url = f"{self._SEARCH_URL}?{params}"
        resp = _cffi_get(url, cookies=cookies)
        resp.raise_for_status()
        soup = _BS(resp.text, "lxml")

        if soup.select_one("form.js-loginForm, .blockMessage--login"):
            raise RuntimeError(
                "Lewd Corner session expired. Log in again via '⬇ Fetch…'.")

        results: list[MetadataCandidate] = []
        for item in soup.select("li.block-row")[:8]:
            a = item.select_one("h3.contentRow-title a, .contentRow-title a")
            if not a:
                continue
            href = a.get("href", "")
            if not href.startswith("http"):
                href = LC_BASE + href
            snippet_el = item.select_one(".contentRow-snippet")
            excerpt = snippet_el.get_text(" ", strip=True) if snippet_el else ""
            raw = a.get_text(separator=" ", strip=True)
            clean = re.sub(r"\s*\[[^\]]+\]", "", raw).strip() or raw
            results.append(MetadataCandidate(
                source=self.SOURCE,
                title=clean,
                url=href,
                synopsis=excerpt,
            ))
        return results

    def fetch_thread(self, url: str) -> MetadataCandidate:
        cookies = self._cookies()
        resp = _cffi_get(url, cookies=cookies)
        resp.raise_for_status()
        soup = _BS(resp.text, "lxml")

        # Title — same bracket format as F95Zone: [Engine] Title [Version] [Dev]
        # Same labelLink <a> child issue; use NavigableString children only.
        from bs4 import NavigableString as _NST
        raw_title = ""
        title_el = soup.select_one("h1.p-title-value")
        if title_el:
            parts = [str(t).strip() for t in title_el.children
                     if isinstance(t, _NST) and str(t).strip()]
            raw_title = " ".join(parts)
            if not raw_title:
                raw_title = title_el.get_text(separator=" ", strip=True)
        developer = ""
        dev_m = re.search(r"\[([^\]]+)\]\s*$", raw_title)
        if dev_m:
            developer = dev_m.group(1)
        clean_title = re.sub(r"\s*\[[^\]]+\]", "", raw_title).strip()

        body_el = (
            soup.select_one("article.message--first .bbWrapper") or
            soup.select_one(".message-body .bbWrapper") or
            soup.select_one(".bbWrapper")
        )
        # Scope image extraction to the first post only.
        first_post_el = (
            soup.select_one("article.message--first") or
            soup.select_one(".message--first") or
            body_el
        )
        first_post_html = str(first_post_el) if first_post_el else ""

        synopsis = cover_url = ""
        tags: list[str] = []
        screenshot_urls: list[str] = []

        # Images: LewdCorner also uses XenForo lazy-loading; extract attachment
        # URLs directly from the first post's HTML.  Filter out /thumb/ copies.
        _seen: set[str] = set()
        _att_pat = re.compile(r'https://attachments\.lewdcorner\.com/[^\s"\'<>]+')
        for u in _att_pat.findall(first_post_html):
            if '/thumb/' in u or u in _seen:
                continue
            if any(x in u for x in ('smilies', 'emoji', '/awards/')):
                continue
            _seen.add(u)
            if not cover_url:
                cover_url = u
            else:
                screenshot_urls.append(u)
            if len(screenshot_urls) >= 18:          # 1 cover + 18 = 19 total
                break
        # Fallback: external images via data-src in first post HTML
        if not cover_url:
            _ext_pat = re.compile(
                r'data-src="(https://(?!attachments\.lewdcorner)[^\s"<>]+\.'
                r'(?:jpg|jpeg|png|webp|gif)(?:\?[^\s"<>]*)?)"',
                re.IGNORECASE,
            )
            for u in _ext_pat.findall(first_post_html):
                if any(x in u for x in ('smilies', 'emoji')) or u in _seen:
                    continue
                _seen.add(u)
                if not cover_url:
                    cover_url = u
                else:
                    screenshot_urls.append(u)
                if len(screenshot_urls) >= 18:
                    break

        if body_el:
            for el in body_el.select("p, div"):
                t = el.get_text(" ", strip=True)
                if len(t) > 80 and not t.startswith("["):
                    synopsis = t
                    break
            if not developer:
                for bold in body_el.select("b, strong"):
                    if "developer" in bold.get_text(strip=True).lower():
                        parent = bold.parent
                        if parent:
                            raw = parent.get_text(" ", strip=True)
                            m = re.search(
                                r"[Dd]eveloper\s*[:\-]\s*([^\n\r|/\\]{2,60})", raw)
                            if m:
                                developer = m.group(1).strip().split("  ")[0].strip()
                        break

        for sel in (
            ".p-description a.tagItem",
            "li.tagItem a",
            "a.tagItem",
            "ul.tagList li a",
            ".tagList a",
        ):
            found = [el.get_text(strip=True) for el in soup.select(sel)
                     if el.get_text(strip=True)]
            if found:
                tags = found
                break
        if not tags and body_el:
            for bold in body_el.select("b, strong"):
                if "genre" in bold.get_text(strip=True).lower():
                    spoiler = bold.find_next(
                        "div", class_=lambda c: c and "bbCodeSpoiler" in c)
                    if spoiler:
                        inner = (
                            spoiler.select_one(".bbCodeSpoiler-content") or
                            spoiler.select_one(".bbCodeBlock-content") or
                            spoiler
                        )
                        raw = inner.get_text(", ", strip=True)
                        tags = [t.strip() for t in re.split(r"[,\n]+", raw) if t.strip()]
                    if not tags:
                        parent = bold.parent
                        if parent:
                            full = parent.get_text(" ", strip=True)
                            m = re.search(r"[Gg]enre\s*[:\-]\s*(.+)", full)
                            if m:
                                raw = m.group(1).strip()
                                tags = [t.strip() for t in re.split(r"[,\n]+", raw)
                                        if t.strip()]
                    break

        return MetadataCandidate(
            source=self.SOURCE,
            title=clean_title or raw_title,
            url=url,
            developer=developer,
            synopsis=synopsis,
            cover_url=cover_url,
            tags=tags,
            screenshot_urls=screenshot_urls,
        )


class ItchioProvider:
    """
    Scrape itch.io search and game pages.
    No login needed for general content; cookies enable 18+ results.
    """

    SOURCE = "itchio"

    def _cookies(self) -> dict:
        return load_site_cookies("itchio")

    def search(self, title: str) -> list[MetadataCandidate]:
        if not HAS_SCRAPING:
            raise RuntimeError("curl_cffi/bs4 not installed")
        from urllib.parse import quote_plus
        # Search scoped to games; logged-in cookies surface adult results
        url = f"{ITCHIO_BASE}/search?q={quote_plus(title)}&type=games"
        resp = _cffi_get(url, cookies=self._cookies())
        resp.raise_for_status()
        soup = _BS(resp.text, "lxml")

        results: list[MetadataCandidate] = []
        for cell in soup.select(".game_cell")[:10]:
            link = cell.select_one("a.game_link, a.title")
            if not link:
                continue
            href = link.get("href", "")
            if not href.startswith("http"):
                href = ITCHIO_BASE + href

            title_el  = cell.select_one(".title, .game_title")
            author_el = cell.select_one(".game_author")
            img_el    = cell.select_one("img.lazy_image, img[data-src], img[src]")
            desc_el   = cell.select_one(".game_text, .short_text")

            cover = ""
            if img_el:
                cover = img_el.get("data-src") or img_el.get("src", "")

            results.append(MetadataCandidate(
                source=self.SOURCE,
                title=title_el.get_text(strip=True) if title_el else "",
                url=href,
                developer=author_el.get_text(strip=True).lstrip("by ") if author_el else "",
                synopsis=desc_el.get_text(" ", strip=True) if desc_el else "",
                cover_url=cover,
            ))
        return results

    def fetch_thread(self, url: str) -> MetadataCandidate:
        resp = _cffi_get(url, cookies=self._cookies())
        resp.raise_for_status()
        soup = _BS(resp.text, "lxml")

        title_el = soup.select_one("h1.game_title, h1.title")
        title = title_el.get_text(strip=True) if title_el else ""

        author_el = soup.select_one(".user_name a, .creator_name")
        developer = author_el.get_text(strip=True) if author_el else ""

        # Cover: og:image is the most reliable
        cover_url = ""
        og = soup.select_one('meta[property="og:image"]')
        if og:
            cover_url = og.get("content", "")
        if not cover_url:
            img_el = soup.select_one(".screenshot_list img, .header_image img")
            if img_el:
                cover_url = img_el.get("src", "")

        # Synopsis: first substantial paragraph in the description
        synopsis = ""
        for el in soup.select(".formatted_description p, .game_description p"):
            t = el.get_text(" ", strip=True)
            if len(t) > 60:
                synopsis = t
                break

        # Tags
        tags: list[str] = []
        for tag_el in soup.select(".game_tags a, .tags a"):
            t = tag_el.get_text(strip=True)
            if t and t not in tags:
                tags.append(t)

        # Screenshots: itch.io shows them in a .screenshot_list
        screenshot_urls: list[str] = []
        _seen: set[str] = {cover_url} if cover_url else set()
        for img_el in soup.select(".screenshot_list img, .screenshots img"):
            src = (img_el.get("data-src") or img_el.get("src", "")).strip()
            # itch.io thumbnail URLs end in /original — swap to full res
            src = re.sub(r"/\d+x\d+/", "/original/", src)
            if src and src.startswith("http") and src not in _seen:
                _seen.add(src)
                screenshot_urls.append(src)
            if len(screenshot_urls) >= 18:          # 1 cover + 18 = 19 total
                break

        return MetadataCandidate(
            source=self.SOURCE,
            title=title,
            url=url,
            developer=developer,
            synopsis=synopsis,
            cover_url=cover_url,
            tags=tags,
            screenshot_urls=screenshot_urls,
        )


# ══════════════════════════════════════════════════════════════════════════════
# METADATA FETCHER  (parallel background threads)
# ══════════════════════════════════════════════════════════════════════════════

_PROVIDERS: dict[str, object] = {
    "vndb":       VNDBProvider(),
    "itchio":     ItchioProvider(),
    "f95zone":    F95ZoneProvider(),
    "lewdcorner": LewdCornerProvider(),
}

SOURCE_LABELS = {
    "vndb":       "VNDB",
    "itchio":     "itch.io",
    "f95zone":    "F95Zone",
    "lewdcorner": "Lewd Corner",
}


class MetadataFetcher:
    """Run provider searches in parallel threads; deliver results via callback."""

    def __init__(
        self,
        title: str,
        sources: list[str],
        on_result: "Callable[[str, list[MetadataCandidate] | Exception], None]",
        on_done:   "Callable[[], None]",
    ) -> None:
        self._title     = title
        self._sources   = sources
        self._on_result = on_result
        self._on_done   = on_done
        self._threads: list[threading.Thread] = []
        self._remaining = len(sources)
        self._lock = threading.Lock()

    def start(self) -> None:
        for src in self._sources:
            t = threading.Thread(target=self._fetch, args=(src,), daemon=True)
            self._threads.append(t)
            t.start()

    def _fetch(self, source: str) -> None:
        provider = _PROVIDERS.get(source)
        try:
            if provider is None:
                raise RuntimeError(f"Unknown source: {source}")
            candidates = provider.search(self._title)  # type: ignore[attr-defined]
            # Eagerly enrich the top result so auto-assignment has real data.
            # Other results are enriched lazily when the user selects them.
            if candidates and hasattr(provider, "fetch_thread"):
                try:
                    candidates[0] = provider.fetch_thread(candidates[0].url)
                except Exception:
                    pass  # keep shallow on network/parse error
            self._on_result(source, candidates)
        except Exception as exc:
            self._on_result(source, exc)
        finally:
            with self._lock:
                self._remaining -= 1
                if self._remaining == 0:
                    self._on_done()


# ══════════════════════════════════════════════════════════════════════════════
# THUMBNAIL CACHE  (thread-safe)
# ══════════════════════════════════════════════════════════════════════════════

class ThumbnailCache:
    """Load PIL images in background; convert to PhotoImage on main thread."""

    def __init__(self, root: tk.Tk) -> None:
        self._root = root
        self._cache: dict[str, "ImageTk.PhotoImage"] = {}
        self._pil_cache: dict[str, "Image.Image"] = {}   # for cross-fade blending
        self._pending: set[str] = set()
        self._work_q: queue.Queue = queue.Queue()
        self._done_q: queue.Queue = queue.Queue()
        self._placeholder: "ImageTk.PhotoImage | None" = None
        threading.Thread(target=self._worker, daemon=True).start()
        self._root.after(60, self._drain)

    # ── public ───────────────────────────────────────────────────────────────

    def request(self, key: str, path: Path, w: int, h: int,
                on_ready) -> None:
        """Queue an async load. on_ready(key, photo) called on main thread."""
        if key in self._cache:
            on_ready(key, self._cache[key], self._pil_cache.get(key))
            return
        if key in self._pending:
            return
        self._pending.add(key)
        self._work_q.put((key, path, w, h, on_ready))

    def get(self, key: str) -> "ImageTk.PhotoImage | None":
        return self._cache.get(key)

    def get_pil(self, key: str) -> "Image.Image | None":
        return self._pil_cache.get(key)

    def placeholder(self) -> "ImageTk.PhotoImage":
        if self._placeholder is None and HAS_PIL:
            img = Image.new("RGB", (THUMB_W, THUMB_H),
                            (int(SEL[1:3], 16), int(SEL[3:5], 16),
                             int(SEL[5:7], 16)))
            self._placeholder = ImageTk.PhotoImage(img)
        return self._placeholder  # type: ignore[return-value]

    def detail_placeholder(self) -> "ImageTk.PhotoImage | None":
        if not HAS_PIL:
            return None
        img = Image.new("RGB", (DETAIL_ART_W, DETAIL_ART_H),
                        (int(SEL[1:3], 16), int(SEL[3:5], 16),
                         int(SEL[5:7], 16)))
        return ImageTk.PhotoImage(img)

    # ── background worker ────────────────────────────────────────────────────

    def _worker(self) -> None:
        while True:
            key, path, w, h, cb = self._work_q.get()
            try:
                pil_img = _pil_load(path, w, h)
            except Exception:
                pil_img = None
            self._done_q.put((key, pil_img, cb))

    def _drain(self) -> None:
        try:
            while True:
                key, pil_img, cb = self._done_q.get_nowait()
                self._pending.discard(key)
                if pil_img is not None and HAS_PIL:
                    photo = ImageTk.PhotoImage(pil_img)
                    self._cache[key] = photo
                    self._pil_cache[key] = pil_img
                    cb(key, photo, pil_img)
                else:
                    cb(key, self.placeholder(), None)
        except queue.Empty:
            pass
        self._root.after(60, self._drain)


# ══════════════════════════════════════════════════════════════════════════════
# PLAY TRACKER
# ══════════════════════════════════════════════════════════════════════════════

class PlayTracker:
    def __init__(self) -> None:
        self._active: dict[str, tuple[subprocess.Popen, float]] = {}

    def launch(self, version: GameVersion, app: "LibraryApp") -> bool:
        if not version.exe_path:
            return False
        if version.folder_name in self._active:
            return False
        try:
            proc = subprocess.Popen(
                [str(version.exe_path)], cwd=str(version.folder_path))
        except OSError as e:
            messagebox.showerror("Launch Error", str(e))
            return False
        start = time.time()
        self._active[version.folder_name] = (proc, start)
        threading.Thread(
            target=self._monitor,
            args=(version.folder_name, proc, start, weakref.ref(app)),
            daemon=True,
        ).start()
        return True

    def _monitor(self, folder_name: str, proc: subprocess.Popen,
                 start: float, app_ref) -> None:
        proc.wait()
        elapsed = int(time.time() - start)
        self._active.pop(folder_name, None)
        app = app_ref()
        if app and app.winfo_exists():
            app.after(0, lambda: app.on_game_exit(folder_name, elapsed))

    def is_playing(self, folder_name: str) -> bool:
        return folder_name in self._active


# ══════════════════════════════════════════════════════════════════════════════
# THEME
# ══════════════════════════════════════════════════════════════════════════════

def apply_theme(root: tk.Tk) -> None:
    style = ttk.Style(root)
    style.theme_use("clam")
    style.configure(".", background=BG, foreground=FG,
                    fieldbackground=BG2, font=("Segoe UI", 10), borderwidth=0)
    style.configure("TFrame", background=BG)
    style.configure("TLabel", background=BG, foreground=FG)
    style.configure("TButton",
                    background=SEL, foreground=FG, padding=(8, 4), relief="flat")
    style.map("TButton",
              background=[("active", ACCENT), ("pressed", ACCENT)],
              foreground=[("active", BG), ("pressed", BG)])
    style.configure("TMenubutton",
                    background=SEL, foreground=FG, padding=(8, 4), relief="flat",
                    arrowsize=0, arrowcolor=SEL)
    style.map("TMenubutton",
              background=[("active", ACCENT), ("disabled", BG2)],
              foreground=[("active", BG), ("disabled", FG_MUT)],
              arrowcolor=[("active", ACCENT), ("disabled", BG2)])
    style.configure("Accent.TButton", background=ACCENT, foreground=BG)
    style.map("Accent.TButton",
              background=[("active", "#74c7ec")],
              foreground=[("active", BG)])
    style.configure("Launch.TButton", background=ACCENT, foreground=BG,
                    padding=(8, 10), font=("Segoe UI", 11, "bold"))
    style.map("Launch.TButton",
              background=[("active", "#74c7ec"), ("disabled", FG_MUT)],
              foreground=[("active", BG), ("disabled", BG)])
    style.configure("Danger.TButton", background="#45475a", foreground=RED)
    style.map("Danger.TButton",
              background=[("active", RED)],
              foreground=[("active", BG)])
    style.configure("Filter.TButton",
                    background=SEL, foreground=FG_DIM, padding=(8, 3))
    style.configure("FilterActive.TButton",
                    background=ACCENT, foreground=BG, padding=(8, 3))
    style.map("FilterActive.TButton",
              background=[("active", "#74c7ec")])
    style.configure("TEntry",
                    fieldbackground=BG2, foreground=FG, insertcolor=FG)
    style.configure("TCombobox", fieldbackground=BG2, foreground=FG,
                    selectbackground=SEL)
    style.map("TCombobox", fieldbackground=[("readonly", BG2)],
              foreground=[("readonly", FG)])
    style.configure("Treeview", background=BG2, foreground=FG,
                    fieldbackground=BG2, rowheight=26, borderwidth=0)
    style.configure("Treeview.Heading",
                    background=SEL, foreground=ACCENT, relief="flat",
                    font=("Segoe UI", 10, "bold"))
    style.map("Treeview",
              background=[("selected", SEL)],
              foreground=[("selected", ACCENT)])
    style.configure("TSeparator", background=SEL)
    style.configure("Status.TLabel", background=BG2, foreground=FG_DIM,
                    font=("Segoe UI", 9), padding=(8, 4))
    style.configure("TNotebook", background=BG2, borderwidth=0)
    style.configure("TNotebook.Tab",
                    background=BG2, foreground=FG_DIM, padding=(14, 6),
                    font=("Segoe UI", 10))
    style.map("TNotebook.Tab",
              background=[("selected", BG)],
              foreground=[("selected", ACCENT)])
    root.configure(bg=BG)
    root.option_add("*TCombobox*Listbox.background", BG2)
    root.option_add("*TCombobox*Listbox.foreground", FG)
    root.option_add("*TCombobox*Listbox.selectBackground", SEL)
    root.option_add("*TCombobox*Listbox.selectForeground", ACCENT)


# ══════════════════════════════════════════════════════════════════════════════
# SCROLLABLE CARD LIST
# ══════════════════════════════════════════════════════════════════════════════

class ScrollableCardList(ttk.Frame):
    def __init__(self, parent, on_select_cb, **kw) -> None:
        super().__init__(parent, **kw)
        self._on_select = on_select_cb
        self._cards: dict[str, "GameCard"] = {}  # base_key → card

        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = ttk.Frame(self.canvas)
        self._win_id = self.canvas.create_window(
            (0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>",
                        lambda e: self.canvas.configure(
                            scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",
                         lambda e: self.canvas.itemconfigure(
                             self._win_id, width=e.width))
        self.canvas.bind("<MouseWheel>", self._on_mw)
        self.inner.bind("<MouseWheel>", self._on_mw)

    def _on_mw(self, e: tk.Event) -> None:  # type: ignore[type-arg]
        self.canvas.yview_scroll(-1 * (e.delta // 120), "units")

    def populate(self, groups: list[GameGroup], ud: UserData,
                 tc: ThumbnailCache, selected_key: str | None = None) -> None:
        for c in self._cards.values():
            c.destroy()
        self._cards.clear()

        for g in groups:
            card = GameCard(
                self.inner, g, ud, tc,
                on_click=lambda bk=g.base_key: self._on_select(bk),
                on_mw=self._on_mw,
            )
            card.pack(fill="x", pady=(0, 1))
            self._cards[g.base_key] = card

            art = _group_art_path(g, ud)
            if art:
                tc.request(
                    g.base_key, art, THUMB_W, THUMB_H,
                    on_ready=self._on_thumb_ready,
                )
            else:
                card.set_thumbnail(tc.placeholder())

        if selected_key and selected_key in self._cards:
            self._cards[selected_key].set_selected(True)

        # Always reset scroll to top when the list is rebuilt
        self.canvas.yview_moveto(0.0)

    def _on_thumb_ready(self, base_key: str, photo, _pil=None) -> None:
        card = self._cards.get(base_key)
        if card and card.winfo_exists():
            card.set_thumbnail(photo)

    def update_card(self, base_key: str, ud: UserData) -> None:
        card = self._cards.get(base_key)
        if card and card.winfo_exists():
            card.refresh_data(ud)

    def scroll_to(self, base_key: str) -> None:
        """Scroll only if the card is not already fully visible."""
        card = self._cards.get(base_key)
        if not card:
            return
        self.update_idletasks()
        total = self.inner.winfo_height()
        if total <= 0:
            return
        card_top    = card.winfo_y()
        card_bottom = card_top + card.winfo_height()
        view_top, view_bottom = self.canvas.yview()
        vis_top    = view_top    * total
        vis_bottom = view_bottom * total
        if card_top >= vis_top and card_bottom <= vis_bottom:
            return   # already fully visible — don't move
        self.canvas.yview_moveto(card_top / total)


# ══════════════════════════════════════════════════════════════════════════════
# GAME CARD
# ══════════════════════════════════════════════════════════════════════════════

class GameCard(tk.Frame):
    def __init__(self, parent, group: GameGroup, ud: UserData,
                 tc: ThumbnailCache, on_click, on_mw, **kw) -> None:
        super().__init__(parent, bg=CARD_BG, height=CARD_H, **kw)
        self.pack_propagate(False)
        self.group = group
        self._on_click = on_click
        self._selected = False

        # ── thumbnail ─────────────────────────────────────────────────────
        self._thumb_lbl = tk.Label(
            self, bg=SEL, width=THUMB_W, height=THUMB_H, cursor="hand2")
        self._thumb_lbl.place(x=8, y=(CARD_H - THUMB_H) // 2,
                              width=THUMB_W, height=THUMB_H)

        # ── text frame ────────────────────────────────────────────────────
        tf = tk.Frame(self, bg=CARD_BG)
        tf.place(x=THUMB_W + 18, y=6,
                 width=1000, height=CARD_H - 8)  # width shrinks with card

        self._name_lbl = tk.Label(
            tf, bg=CARD_BG, fg=FG, text=group.display_name,
            font=("Segoe UI", 10, "bold"), anchor="w", cursor="hand2")
        self._name_lbl.pack(fill="x", anchor="w")

        self._sub_lbl = tk.Label(
            tf, bg=CARD_BG, fg=FG_DIM, text="",
            font=("Segoe UI", 8), anchor="w")
        self._sub_lbl.pack(fill="x", anchor="w")

        self._time_lbl = tk.Label(
            tf, bg=CARD_BG, fg=FG_DIM, text="",
            font=("Segoe UI", 8), anchor="w")
        self._time_lbl.pack(fill="x", anchor="w")

        # ── now-playing badge ─────────────────────────────────────────────
        self._playing_badge = tk.Label(
            self, text="▶ Playing", bg=GREEN, fg=BG,
            font=("Segoe UI", 7, "bold"), padx=4, pady=1)

        # ── played dot ────────────────────────────────────────────────────
        self._dot = tk.Label(self, bg=CARD_BG, font=("Segoe UI", 14))
        self._dot.place(relx=1.0, rely=0.5, anchor="e", x=-12)

        self.refresh_data(ud)
        self._bind_events(on_mw)

    def _bind_events(self, on_mw) -> None:
        for w in self.winfo_children() + [self]:
            w.bind("<Button-1>", lambda e: self._on_click())
            w.bind("<Enter>", self._enter)
            w.bind("<Leave>", self._leave)
            w.bind("<MouseWheel>", on_mw)
        # text frame children too
        for w in self._name_lbl.master.winfo_children():
            w.bind("<Button-1>", lambda e: self._on_click())
            w.bind("<Enter>", self._enter)
            w.bind("<Leave>", self._leave)
            w.bind("<MouseWheel>", on_mw)

    def refresh_data(self, ud: UserData) -> None:
        g = self.group
        # Version string
        vers = " · ".join(
            v.version_str for v in g.versions if v.version_str) or "—"
        # Most recent last_played across all versions
        lp_strs = [ud.last_played[v.folder_name]
                   for v in g.versions if v.folder_name in ud.last_played]
        lp_str = max(lp_strs) if lp_strs else ""
        sub = vers
        if lp_str:
            sub += f"  ·  {fmt_date(lp_str)}"
        elif not any(detect_played(v, ud) for v in g.versions):
            sub += "  ·  Never played"
        if g.versions and not g.versions[-1].is_renpy:
            sub += "  ·  [EXE]"
        self._sub_lbl.configure(text=sub)
        # Total playtime
        total = sum(ud.playtime.get(v.folder_name, 0) for v in g.versions)
        self._time_lbl.configure(text=f"  ⏱ {fmt_time(total)}" if total else "")
        # Played dot
        played = any(detect_played(v, ud) for v in g.versions)
        self._dot.configure(text="●" if played else "○",
                             fg=GREEN if played else FG_MUT)

    def set_thumbnail(self, photo) -> None:
        if photo:
            self._thumb_lbl.configure(image=photo)
            self._thumb_lbl.image = photo  # keep ref alive

    def set_selected(self, sel: bool) -> None:
        self._selected = sel
        self._set_bg(CARD_SEL if sel else CARD_BG)

    def set_playing(self, playing: bool) -> None:
        if playing:
            self._playing_badge.place(relx=1.0, y=4, anchor="ne", x=-36)
        else:
            self._playing_badge.place_forget()

    def _enter(self, _e=None) -> None:
        if not self._selected:
            self._set_bg(CARD_HOV)

    def _leave(self, _e=None) -> None:
        if not self._selected:
            self._set_bg(CARD_BG)

    def _set_bg(self, c: str) -> None:
        self.configure(bg=c)
        self._thumb_lbl.configure(bg=c if c != CARD_BG else SEL)
        tf = self._name_lbl.master
        tf.configure(bg=c)
        for w in tf.winfo_children():
            w.configure(bg=c)
        self._dot.configure(bg=c)


# ══════════════════════════════════════════════════════════════════════════════
# METADATA DIALOGS
# ══════════════════════════════════════════════════════════════════════════════

class FetchProgressDialog(tk.Toplevel):
    """
    Non-blocking progress window shown while providers are queried.
    Closes automatically when all sources have responded; the caller
    can also call .destroy() to cancel in-flight requests.
    """

    def __init__(self, parent, title: str, sources: list[str]) -> None:
        super().__init__(parent)
        self.title("Fetching Metadata")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(parent)

        tk.Label(self, bg=BG, fg=FG, text=f'Searching for "{title}"',
                 font=("Segoe UI", 10, "bold"), padx=16, pady=10).pack()

        self._rows: dict[str, tk.Label] = {}
        for src in sources:
            row = tk.Frame(self, bg=BG, padx=16)
            row.pack(fill="x", pady=2)
            tk.Label(row, bg=BG, fg=FG_DIM,
                     text=SOURCE_LABELS.get(src, src) + ":",
                     font=("Segoe UI", 9), width=14, anchor="w").pack(side="left")
            lbl = tk.Label(row, bg=BG, fg=YELLOW, text="…searching",
                           font=("Segoe UI", 9), anchor="w")
            lbl.pack(side="left")
            self._rows[src] = lbl

        ttk.Button(self, text="Cancel", command=self.destroy).pack(pady=(6, 12))

        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width()  // 2 - self.winfo_width()  // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2 - self.winfo_height() // 2
        self.geometry(f"+{px}+{py}")

    def set_status(self, source: str, text: str, color: str = FG) -> None:
        lbl = self._rows.get(source)
        if lbl and lbl.winfo_exists():
            lbl.configure(text=text, fg=color)


class MetadataPickerDialog(tk.Toplevel):
    """
    Full metadata picker: shows candidates per source, lets the user
    choose which source provides each field, then saves to .vnpf/.
    """

    # Fields the user can assign per-source
    _FIELDS = ["title", "developer", "synopsis", "cover_art", "tags"]
    _FIELD_LABELS = {
        "title":     "Title",
        "developer": "Developer",
        "synopsis":  "Synopsis",
        "cover_art": "Cover Art",
        "tags":      "Tags",
    }

    def __init__(self, parent, version: "GameVersion", app: "LibraryApp") -> None:
        super().__init__(parent)
        self.title("Metadata — " + version.display_name)
        self.configure(bg=BG)
        self.resizable(True, True)
        self.transient(parent)
        self.geometry("900x640")

        self._version  = version
        self._app      = app
        self._selected_candidate: dict[str, MetadataCandidate | None] = {}
        self._preview_url: str = ""  # currently loaded cover URL in preview

        # Per-field source: stores SOURCE KEY (e.g. "vndb"), not display label
        self._field_source: dict[str, str] = {f: "" for f in self._FIELDS}
        # Ordered list of available source keys for combo indexing
        self._combo_keys: list[str] = []
        # All images from the cover_art source in user-defined order.
        # _img_order[0] → cover,  [1:] → screenshots.
        self._img_order: list[str] = []

        self._build()
        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width()  // 2 - self.winfo_width()  // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2 - self.winfo_height() // 2
        self.geometry(f"+{max(0,px)}+{max(0,py)}")

    # ── Layout ───────────────────────────────────────────────────────────────

    def _build(self) -> None:
        # ── Top bar ──────────────────────────────────────────────────────────
        top = tk.Frame(self, bg=BG2, padx=10, pady=8)
        top.pack(fill="x")
        tk.Label(top, bg=BG2, fg=ACCENT,
                 text=f"Metadata for: {self._version.display_name}",
                 font=("Segoe UI", 11, "bold")).pack(side="left")

        btn_bar = tk.Frame(top, bg=BG2)
        btn_bar.pack(side="right")
        self._save_btn = ttk.Button(btn_bar, text="Save", style="Accent.TButton",
                                    command=self._save)
        self._save_btn.pack(side="left", padx=(0, 6))
        ttk.Button(btn_bar, text="Cancel",
                   command=self.destroy).pack(side="left")

        # ── Two-column PanedWindow ────────────────────────────────────────────
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=10, pady=8)
        self._paned = paned

        # Left panel: scrollable source columns
        left = ttk.Frame(paned)
        paned.add(left, weight=3)

        src_canvas = tk.Canvas(left, bg=BG, highlightthickness=0)
        src_canvas.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(left, orient="vertical", command=src_canvas.yview)
        vsb.pack(side="right", fill="y")
        src_canvas.configure(yscrollcommand=vsb.set)
        src_inner = tk.Frame(src_canvas, bg=BG)
        src_canvas.create_window((0, 0), window=src_inner, anchor="nw")
        src_inner.bind("<Configure>",
            lambda e: src_canvas.configure(scrollregion=src_canvas.bbox("all")))
        self._src_inner = src_inner
        self._src_canvas = src_canvas

        # Mouse-wheel scrolling on canvas and all its content
        src_canvas.bind("<MouseWheel>", self._on_src_wheel)
        src_inner.bind("<MouseWheel>", self._on_src_wheel)

        # Right panel: field-source picker + preview
        right = tk.Frame(paned, bg=BG2)
        paned.add(right, weight=2)
        self._build_field_picker(right)

        # Force sash to saved/default position after window renders
        default_sash = self._app._settings.get("meta_sash", 560)
        paned.bind(
            "<Map>",
            lambda e, p=paned, d=default_sash: self.after(
                50, lambda: self._apply_sash(p, d)))
        paned.bind(
            "<ButtonRelease-1>",
            lambda e, p=paned: self._persist_sash(p))

    def _apply_sash(self, paned: ttk.PanedWindow, pos: int) -> None:
        try:
            paned.update_idletasks()
            if paned.winfo_width() > pos + 80:
                paned.sashpos(0, pos)
        except Exception:
            pass

    def _persist_sash(self, paned: ttk.PanedWindow) -> None:
        try:
            pos = paned.sashpos(0)
            if pos > 0:
                self._app._settings["meta_sash"] = pos
                save_settings(self._app._settings)
        except Exception:
            pass

    def _on_src_wheel(self, event) -> None:
        self._src_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _bind_wheel_to_tree(self, widget: tk.Widget) -> None:
        """Recursively bind mouse-wheel on widget and all descendants."""
        widget.bind("<MouseWheel>", self._on_src_wheel)
        for child in widget.winfo_children():
            self._bind_wheel_to_tree(child)

    def _build_field_picker(self, parent: tk.Frame) -> None:
        tk.Label(parent, bg=BG2, fg=FG, text="Field Sources",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 4))
        tk.Label(parent, bg=BG2, fg=FG_DIM,
                 text="Pick which source provides each field:",
                 font=("Segoe UI", 8)).pack(anchor="w", padx=10, pady=(0, 8))

        grid_frame = tk.Frame(parent, bg=BG2)
        grid_frame.pack(fill="x", padx=10)
        self._field_combos: dict[str, ttk.Combobox] = {}

        for row_idx, f in enumerate(self._FIELDS):
            tk.Label(grid_frame, bg=BG2, fg=FG_DIM,
                     text=self._FIELD_LABELS[f] + ":",
                     font=("Segoe UI", 9), width=11, anchor="w"
                     ).grid(row=row_idx, column=0, sticky="w", pady=3)
            cb = ttk.Combobox(grid_frame, state="readonly", width=15)
            cb.grid(row=row_idx, column=1, sticky="w", pady=3, padx=(4, 0))
            cb.bind("<<ComboboxSelected>>",
                    lambda e, field=f, combo=cb: self._on_combo_select(field, combo))
            self._field_combos[f] = cb

        ttk.Separator(parent, orient="horizontal").pack(fill="x", padx=10, pady=10)

        # Preview area
        tk.Label(parent, bg=BG2, fg=FG, text="Preview",
                 font=("Segoe UI", 10, "bold")).pack(anchor="w", padx=10, pady=(0, 6))
        self._preview_cover = tk.Label(parent, bg=SEL,
                                       text="No cover", fg=FG_DIM,
                                       font=("Segoe UI", 8))
        self._preview_cover.pack(fill="x", padx=10, pady=(0, 6))

        self._preview_title = tk.Label(parent, bg=BG2, fg=ACCENT, text="",
                                       font=("Segoe UI", 9, "bold"),
                                       wraplength=280, justify="left", anchor="w")
        self._preview_title.pack(anchor="w", padx=10)
        self._preview_dev = tk.Label(parent, bg=BG2, fg=FG_DIM, text="",
                                     font=("Segoe UI", 8), anchor="w")
        self._preview_dev.pack(anchor="w", padx=10)
        self._preview_syn = tk.Text(parent, bg=BG3, fg=FG,
                                    height=6, wrap="word",
                                    font=("Segoe UI", 8),
                                    state="disabled", relief="flat")
        self._preview_syn.pack(fill="x", padx=10, pady=(4, 0))

        tk.Label(parent, bg=BG2, fg=FG_DIM, text="Tags:",
                 font=("Segoe UI", 8), anchor="w").pack(anchor="w", padx=10, pady=(6, 2))
        self._preview_tags = tk.Label(parent, bg=BG2, fg=FG,
                                      font=("Segoe UI", 8),
                                      wraplength=280, justify="left", anchor="w")
        self._preview_tags.pack(anchor="w", padx=10)

        # ── Image order list ──────────────────────────────────────────────────
        ttk.Separator(parent, orient="horizontal").pack(
            fill="x", padx=10, pady=(8, 4))
        tk.Label(parent, bg=BG2, fg=FG_DIM,
                 text="Images  (first = cover):",
                 font=("Segoe UI", 8), anchor="w").pack(
            anchor="w", padx=10, pady=(0, 2))

        img_list_frame = tk.Frame(parent, bg=BG2)
        img_list_frame.pack(fill="x", padx=10)
        self._img_listbox = tk.Listbox(
            img_list_frame, bg=BG3, fg=FG,
            selectmode="single", height=6,
            font=("Segoe UI", 8), activestyle="none",
            selectbackground=SEL, selectforeground=FG,
            exportselection=False, relief="flat", bd=0,
        )
        self._img_listbox.pack(side="left", fill="both", expand=True)
        _img_sb = ttk.Scrollbar(img_list_frame, orient="vertical",
                                 command=self._img_listbox.yview)
        _img_sb.pack(side="left", fill="y")
        self._img_listbox.configure(yscrollcommand=_img_sb.set)

        img_btn_frame = tk.Frame(parent, bg=BG2)
        img_btn_frame.pack(fill="x", padx=10, pady=(2, 6))
        for _lbl, _cmd in [
            ("▲ Up",     self._img_move_up),
            ("▼ Down",   self._img_move_down),
            ("★ Cover",  self._img_set_cover),
            ("✕ Remove", self._img_remove),
        ]:
            tk.Button(img_btn_frame, bg=BG3, fg=FG,
                      text=_lbl, font=("Segoe UI", 8),
                      relief="flat", padx=4,
                      command=_cmd).pack(side="left", padx=(0, 2))

    # ── Source column building ────────────────────────────────────────────────

    def _build_source_column(self, source: str,
                              candidates: list[MetadataCandidate]) -> None:
        """Add or refresh a source column in the left pane."""
        for child in self._src_inner.winfo_children():
            if getattr(child, "_meta_source", None) == source:
                child.destroy()

        col = tk.Frame(self._src_inner, bg=BG,
                       highlightbackground=FG_MUT, highlightthickness=1)
        col._meta_source = source  # type: ignore[attr-defined]
        col.pack(fill="x", padx=4, pady=4)

        hdr = tk.Frame(col, bg=BG2, padx=8, pady=4)
        hdr.pack(fill="x")
        tk.Label(hdr, bg=BG2, fg=ACCENT,
                 text=SOURCE_LABELS.get(source, source),
                 font=("Segoe UI", 9, "bold")).pack(side="left")

        if not candidates:
            tk.Label(col, bg=BG, fg=FG_DIM, text="No results",
                     font=("Segoe UI", 8), padx=8, pady=6).pack(anchor="w")
            self._selected_candidate.setdefault(source, None)
            self._refresh_field_combos()
            return

        self._selected_candidate[source] = candidates[0]
        sel_var = tk.IntVar(value=0)

        for i, cand in enumerate(candidates):
            row = tk.Frame(col, bg=BG, padx=6, pady=4,
                           highlightbackground=FG_MUT, highlightthickness=1,
                           cursor="hand2")
            row.pack(fill="x", padx=4, pady=2)

            rb = tk.Radiobutton(
                row, bg=BG, fg=FG, selectcolor=BG2,
                variable=sel_var, value=i,
                command=lambda c=cand, src=source: self._pick_candidate(src, c),
            )
            rb.pack(side="left")

            info = tk.Frame(row, bg=BG)
            info.pack(side="left", fill="x", expand=True)
            tk.Label(info, bg=BG, fg=FG, text=cand.title,
                     font=("Segoe UI", 8, "bold"),
                     wraplength=340, justify="left", anchor="w").pack(anchor="w")
            if cand.developer:
                tk.Label(info, bg=BG, fg=FG_DIM, text=cand.developer,
                         font=("Segoe UI", 7), anchor="w").pack(anchor="w")
            if cand.synopsis:
                snippet = cand.synopsis[:120] + ("…" if len(cand.synopsis) > 120 else "")
                tk.Label(info, bg=BG, fg=FG_MUT, text=snippet,
                         font=("Segoe UI", 7), wraplength=340,
                         justify="left", anchor="w").pack(anchor="w")

            if cand.cover_url and HAS_PIL and HAS_SCRAPING:
                thumb_lbl = tk.Label(row, bg=BG, text="")
                thumb_lbl.pack(side="right", padx=4)
                threading.Thread(
                    target=self._load_thumb,
                    args=(cand.cover_url, thumb_lbl),
                    daemon=True,
                ).start()

        self._refresh_field_combos()
        self._bind_wheel_to_tree(col)
        # Auto-assign fields from first source that loads with data
        self._auto_assign(source, candidates[0])

    def _load_thumb(self, url: str, label: tk.Label) -> None:
        if not HAS_PIL or not HAS_SCRAPING:
            return
        try:
            resp = _req.get(url, headers={"User-Agent": SCRAPER_UA},
                            timeout=SCRAPER_TIMEOUT)
            resp.raise_for_status()
            from io import BytesIO
            img = Image.open(BytesIO(resp.content)).convert("RGB")
            img.thumbnail((64, 48), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.after(0, lambda lbl=label, ph=photo: self._set_thumb(lbl, ph))
        except Exception:
            pass

    def _set_thumb(self, lbl: tk.Label, photo: "ImageTk.PhotoImage") -> None:
        if lbl.winfo_exists():
            lbl.configure(image=photo, width=64, height=48)
            lbl._photo = photo  # type: ignore[attr-defined]

    # ── Candidate / field-source management ──────────────────────────────────

    def _pick_candidate(self, source: str, cand: MetadataCandidate) -> None:
        """User clicked a radio button to select this candidate."""
        self._selected_candidate[source] = cand
        provider = _PROVIDERS.get(source) if HAS_SCRAPING else None
        # If this candidate is shallow (no developer/cover), enrich it lazily
        if provider and hasattr(provider, "fetch_thread") and not cand.developer and not cand.cover_url:
            threading.Thread(
                target=self._enrich_candidate,
                args=(source, cand),
                daemon=True,
            ).start()
        else:
            self._auto_assign(source, cand)
            self._update_preview()

    def _enrich_candidate(self, source: str, original: MetadataCandidate) -> None:
        provider = _PROVIDERS.get(source)
        if not provider:
            return
        try:
            enriched = provider.fetch_thread(original.url)  # type: ignore[attr-defined]
        except Exception:
            return
        # Only apply if the user hasn't switched to a different candidate
        if self._selected_candidate.get(source) is original:
            self.after(0, lambda e=enriched, s=source: self._on_enriched(s, e))

    def _on_enriched(self, source: str, enriched: MetadataCandidate) -> None:
        """Called on main thread after lazy enrichment completes."""
        self._selected_candidate[source] = enriched
        self._auto_assign(source, enriched)
        self._update_preview()

    def _auto_assign(self, source: str, cand: MetadataCandidate) -> None:
        """Fill any unset fields from this candidate."""
        for f in self._FIELDS:
            if not self._field_source[f]:
                val = self._candidate_value(cand, f)
                if val:
                    self._field_source[f] = source
        self._sync_combos_to_state()
        # If cover_art was just assigned and image list is empty, populate it
        if not self._img_order and self._field_source.get("cover_art") == source:
            self._img_order = self._build_img_order(cand)
            self._populate_img_list()
        self._update_preview()

    def _refresh_field_combos(self) -> None:
        """Rebuild combo value lists from currently available sources."""
        self._combo_keys = [s for s, c in self._selected_candidate.items()
                            if c is not None]
        display = ["(none)"] + [SOURCE_LABELS.get(k, k) for k in self._combo_keys]
        for cb in self._field_combos.values():
            cb["values"] = display
        self._sync_combos_to_state()

    def _sync_combos_to_state(self) -> None:
        """Set each combo's displayed text to match _field_source[f] (a key)."""
        for f, cb in self._field_combos.items():
            key = self._field_source[f]
            if key and key in self._combo_keys:
                cb.set(SOURCE_LABELS.get(key, key))
            else:
                cb.set("(none)")

    def _on_combo_select(self, field: str, combo: ttk.Combobox) -> None:
        """User picked a source from a combo — store the key, update preview."""
        idx = combo.current()
        if idx <= 0:
            self._field_source[field] = ""
            if field == "cover_art":
                self._img_order = []
                self._populate_img_list()
        else:
            key = self._combo_keys[idx - 1]
            self._field_source[field] = key
            if field == "cover_art":
                cand = self._selected_candidate.get(key)
                self._img_order = self._build_img_order(cand)
                self._populate_img_list()
        self._update_preview()

    # ── Image list helpers ────────────────────────────────────────────────────

    @staticmethod
    def _build_img_order(cand: "MetadataCandidate | None") -> list[str]:
        """Return [cover_url] + screenshot_urls from a candidate, deduped."""
        if not cand:
            return []
        urls: list[str] = []
        seen: set[str] = set()
        for u in ([cand.cover_url] if cand.cover_url else []) + list(cand.screenshot_urls):
            if u and u not in seen:
                seen.add(u)
                urls.append(u)
        return urls

    def _populate_img_list(self) -> None:
        """Rebuild the image listbox from self._img_order."""
        self._img_listbox.delete(0, "end")
        for i, url in enumerate(self._img_order):
            name = url.rsplit("/", 1)[-1].split("?")[0][:50]
            prefix = "★  " if i == 0 else f"   {i:2d}. "
            self._img_listbox.insert("end", f"{prefix}{name}")
        if self._img_order:
            self._img_listbox.itemconfigure(0, fg=ACCENT)

    def _img_move_up(self) -> None:
        sel = self._img_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        i = sel[0]
        self._img_order[i - 1], self._img_order[i] = (
            self._img_order[i], self._img_order[i - 1])
        self._populate_img_list()
        self._img_listbox.selection_set(i - 1)
        self._img_listbox.see(i - 1)
        self._refresh_cover_from_order()

    def _img_move_down(self) -> None:
        sel = self._img_listbox.curselection()
        if not sel or sel[0] >= len(self._img_order) - 1:
            return
        i = sel[0]
        self._img_order[i], self._img_order[i + 1] = (
            self._img_order[i + 1], self._img_order[i])
        self._populate_img_list()
        self._img_listbox.selection_set(i + 1)
        self._img_listbox.see(i + 1)
        self._refresh_cover_from_order()

    def _img_set_cover(self) -> None:
        sel = self._img_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        url = self._img_order.pop(sel[0])
        self._img_order.insert(0, url)
        self._populate_img_list()
        self._img_listbox.selection_set(0)
        self._img_listbox.see(0)
        self._refresh_cover_from_order()

    def _img_remove(self) -> None:
        sel = self._img_listbox.curselection()
        if not sel:
            return
        i = sel[0]
        self._img_order.pop(i)
        self._populate_img_list()
        new_sel = min(i, len(self._img_order) - 1)
        if new_sel >= 0:
            self._img_listbox.selection_set(new_sel)
        self._refresh_cover_from_order()

    def _refresh_cover_from_order(self) -> None:
        """Reload the cover preview from the current first entry in _img_order."""
        url = self._img_order[0] if self._img_order else ""
        if url and url != self._preview_url and HAS_PIL and HAS_SCRAPING:
            self._preview_url = url
            threading.Thread(
                target=self._fetch_preview_cover, args=(url,), daemon=True
            ).start()
        elif not url:
            self._preview_url = ""
            self._preview_cover.configure(image="", text="No cover")

    # ── Preview ───────────────────────────────────────────────────────────────

    def _update_preview(self) -> None:
        """Refresh the right-hand preview from current _field_source selections."""
        title = dev = syn = ""
        cover_url = ""
        tags: list[str] = []

        for f in self._FIELDS:
            src = self._field_source[f]
            cand = self._selected_candidate.get(src)
            if not cand:
                continue
            if f == "title":
                title = cand.title
            elif f == "developer":
                dev = cand.developer
            elif f == "synopsis":
                syn = cand.synopsis
            elif f == "cover_art":
                # Use the user's ordered list if populated; else fall back to candidate
                cover_url = self._img_order[0] if self._img_order else cand.cover_url
            elif f == "tags":
                tags = cand.tags

        self._preview_title.configure(text=title)
        self._preview_dev.configure(text=dev)
        self._preview_syn.configure(state="normal")
        self._preview_syn.delete("1.0", "end")
        if syn:
            self._preview_syn.insert("1.0", syn)
        self._preview_syn.configure(state="disabled")
        self._preview_tags.configure(
            text=", ".join(tags) if tags else "(none)")

        if cover_url and cover_url != self._preview_url and HAS_PIL and HAS_SCRAPING:
            self._preview_url = cover_url
            threading.Thread(
                target=self._fetch_preview_cover,
                args=(cover_url,),
                daemon=True,
            ).start()
        elif not cover_url:
            self._preview_url = ""
            self._preview_cover.configure(image="", text="No cover")

    def _fetch_preview_cover(self, url: str) -> None:
        try:
            resp = _req.get(url, impersonate="chrome131",
                            timeout=SCRAPER_TIMEOUT)
            resp.raise_for_status()
            from io import BytesIO
            img = Image.open(BytesIO(resp.content)).convert("RGB")
            img.thumbnail((300, 180), Image.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            self.after(0, lambda ph=photo, u=url: self._set_preview_cover(ph, u))
        except Exception:
            pass

    def _set_preview_cover(self, photo: "ImageTk.PhotoImage", url: str) -> None:
        # Discard if user already switched to a different source
        if not self.winfo_exists() or url != self._preview_url:
            return
        self._preview_cover.configure(image=photo, text="")
        self._preview_cover._photo = photo  # type: ignore[attr-defined]

    @staticmethod
    def _candidate_value(cand: MetadataCandidate, field: str) -> str:
        return {
            "title":     cand.title,
            "developer": cand.developer,
            "synopsis":  cand.synopsis,
            "cover_art": cand.cover_url,
            "tags":      ", ".join(cand.tags),
        }.get(field, "")

    # ── Public API ───────────────────────────────────────────────────────────

    def populate_source(self, source: str,
                        candidates: list[MetadataCandidate] | Exception) -> None:
        """Called from MetadataFetcher callback — safe to call from any thread."""
        if isinstance(candidates, Exception):
            self.after(0, lambda: self._build_error_column(source, str(candidates)))
        else:
            self.after(0, lambda c=candidates: self._build_source_column(source, c))

    def _build_error_column(self, source: str, msg: str) -> None:
        for child in self._src_inner.winfo_children():
            if getattr(child, "_meta_source", None) == source:
                child.destroy()

        col = tk.Frame(self._src_inner, bg=BG,
                       highlightbackground=RED, highlightthickness=1)
        col._meta_source = source  # type: ignore[attr-defined]
        col.pack(fill="x", padx=4, pady=4)
        hdr = tk.Frame(col, bg=BG2, padx=8, pady=4)
        hdr.pack(fill="x")
        tk.Label(hdr, bg=BG2, fg=RED,
                 text=SOURCE_LABELS.get(source, source) + " — Error",
                 font=("Segoe UI", 9, "bold")).pack(side="left")
        tk.Label(col, bg=BG, fg=FG_DIM,
                 text=msg[:160], font=("Segoe UI", 7),
                 wraplength=400, padx=8, pady=4).pack(anchor="w")
        self._selected_candidate[source] = None
        self._refresh_field_combos()
        self._bind_wheel_to_tree(col)

    # ── Save ─────────────────────────────────────────────────────────────────

    def _save(self) -> None:
        """Build metadata dict from selections and save to .vnpf/ synchronously."""
        field_sources: dict[str, str] = {}
        final: dict[str, object] = {}
        cover_url = ""
        screenshot_urls: list[str] = []

        for f in self._FIELDS:
            src = self._field_source[f]
            cand = self._selected_candidate.get(src)
            if not cand:
                continue
            field_sources[f] = src
            if f == "title":
                final["game_title"] = cand.title
            elif f == "developer":
                final["developer"] = cand.developer
            elif f == "synopsis":
                final["synopsis"] = cand.synopsis
            elif f == "cover_art":
                if self._img_order:
                    cover_url      = self._img_order[0]
                    screenshot_urls = list(self._img_order[1:])
                else:
                    cover_url       = cand.cover_url
                    screenshot_urls = list(cand.screenshot_urls)
            elif f == "tags":
                final["fetched_tags"] = cand.tags

        sources: dict[str, dict] = {}
        for src, cand in self._selected_candidate.items():
            if cand:
                sources[src] = {
                    "url":        cand.url,
                    "fetched_at": datetime.datetime.now().date().isoformat(),
                }

        meta: dict = {
            "vnpf_version": 1,
            "fetched_at":   datetime.datetime.now().date().isoformat(),
            "field_sources": field_sources,
            "sources":       sources,
            **final,
        }

        folder_path = self._version.folder_path

        # Disable save button while working
        self._save_btn.configure(text="Saving…", state="disabled")

        # Capture fetched tags now so the bg thread can pass them to the main thread
        fetched_tags: list[str] = list(meta.get("fetched_tags", []))

        n_shots = len(screenshot_urls)

        def _bg_save():
            if cover_url and HAS_SCRAPING:
                self._app.after(0, lambda: self._save_btn.configure(
                    text="Downloading cover…"))
                _download_cover(cover_url, folder_path)
                if n_shots:
                    self._app.after(0, lambda: self._save_btn.configure(
                        text=f"Downloading {n_shots} screenshots…"))
                saved = _download_screenshots(screenshot_urls, folder_path)
                meta["screenshot_files"] = saved          # local filenames
                meta["screenshot_urls"]  = screenshot_urls  # original URLs
            save_game_metadata(folder_path, meta)
            self._version.metadata = meta
            self._app.after(0, lambda: self._finish_save(fetched_tags))

        threading.Thread(target=_bg_save, daemon=True).start()

    def _finish_save(self, fetched_tags: list[str]) -> None:
        base_key = self._version.base_key

        # Merge fetched tags into UserData.tags — ADD to existing, never wipe.
        if fetched_tags:
            ud = self._app.user_data
            existing = ud.tags.get(base_key, [])
            merged = list(existing)
            for t in fetched_tags:
                if t and t not in merged:
                    merged.append(t)
            if merged != existing:
                ud.tags[base_key] = merged
                save_userdata(self._app.user_data)

        # Bust thumbnail cache for this game so the newly downloaded cover
        # is loaded instead of the stale cached thumbnail.
        tc = self._app.thumb_cache
        for k in list(tc._cache):
            if k == base_key or k.startswith(f"detail:{base_key}"):
                tc._cache.pop(k, None)
                tc._pil_cache.pop(k, None)

        self._app.refresh()
        if self.winfo_exists():
            self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# TAG PICKER DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class TagPickerDialog(tk.Toplevel):
    """Modal dialog for adding/removing tags from a game."""

    def __init__(self, parent, current_tags: list[str], ud: "UserData") -> None:
        super().__init__(parent)
        self.title("Edit Tags")
        self.configure(bg=BG)
        self.resizable(True, True)
        self.grab_set()
        self.result: list[str] | None = None
        self._ud = ud
        self._vars: dict[str, tk.BooleanVar] = {}

        # All preset tags = built-in + user-promoted (deduplicated, stable order)
        all_presets = list(PRESET_TAGS) + [
            t for t in ud.custom_presets if t not in PRESET_TAGS]

        ttk.Label(self, text="Preset tags:", font=("Segoe UI", 10, "bold"),
                  padding=(12, 10, 12, 4)).pack(anchor="w")

        self._grid_frame = tk.Frame(self, bg=BG)
        self._grid_frame.pack(padx=12, pady=(0, 8), anchor="w")
        self._render_preset_grid(all_presets, current_tags)

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        # Custom tag entry row
        custom_row = tk.Frame(self, bg=BG)
        custom_row.pack(fill="x", padx=12, pady=6)
        ttk.Label(custom_row, text="Custom tag:").pack(side="left")
        self._custom_var = tk.StringVar()
        entry = ttk.Entry(custom_row, textvariable=self._custom_var, width=18)
        entry.pack(side="left", padx=(6, 4))
        entry.bind("<Return>", lambda _: self._add_custom())
        # "Save as preset" checkbox on the Add row
        self._save_as_preset_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            custom_row, text="Save as preset", variable=self._save_as_preset_var,
            bg=BG, fg=FG_DIM, selectcolor=SEL,
            activebackground=BG, activeforeground=FG,
            font=("Segoe UI", 8)).pack(side="left", padx=(0, 4))
        ttk.Button(custom_row, text="Add",
                   command=self._add_custom).pack(side="left")

        # Existing custom tags (tags on this game that aren't in any preset list)
        self._custom_tags: list[str] = [t for t in current_tags
                                         if t not in all_presets]
        self._custom_frame = tk.Frame(self, bg=BG)
        self._custom_frame.pack(fill="x", padx=12, pady=(0, 4))
        self._render_custom()

        ttk.Separator(self, orient="horizontal").pack(fill="x", padx=12)

        # Buttons
        btn_row = ttk.Frame(self)
        btn_row.pack(fill="x", padx=12, pady=(6, 12))
        ttk.Button(btn_row, text="Save", style="Accent.TButton",
                   command=self._save).pack(side="right", padx=(4, 0))
        ttk.Button(btn_row, text="Cancel",
                   command=self.destroy).pack(side="right")
        ttk.Button(btn_row, text="Manage Presets...",
                   command=self._open_manage_presets).pack(side="left")

    def _render_preset_grid(self, all_presets: list[str],
                             current_tags: list[str]) -> None:
        for w in self._grid_frame.winfo_children():
            w.destroy()
        self._vars.clear()
        cols = 3
        for i, tag in enumerate(all_presets):
            var = tk.BooleanVar(value=tag in current_tags or tag in self._vars)
            # Preserve existing var state if re-rendering
            if tag in self._vars:
                var.set(self._vars[tag].get())
            self._vars[tag] = var
            cb = tk.Checkbutton(self._grid_frame, text=tag, variable=var,
                                bg=BG, fg=FG, selectcolor=SEL,
                                activebackground=BG, activeforeground=FG,
                                font=("Segoe UI", 9))
            cb.grid(row=i // cols, column=i % cols, sticky="w", padx=4, pady=1)

    def _render_custom(self) -> None:
        for w in self._custom_frame.winfo_children():
            w.destroy()
        if not self._custom_tags:
            return
        ttk.Label(self._custom_frame, text="Custom (this game only):",
                  foreground=FG_DIM, font=("Segoe UI", 8)).pack(
                      anchor="w", pady=(4, 2))
        for tag in self._custom_tags:
            row = tk.Frame(self._custom_frame, bg=BG)
            row.pack(anchor="w", pady=1)
            tk.Label(row, text=tag, bg=BG, fg=ACCENT,
                     font=("Segoe UI", 9)).pack(side="left")
            # ★ promote to preset
            tk.Button(row, text="★", bg=BG, fg=YELLOW, relief="flat",
                      command=lambda t=tag: self._promote_to_preset(t),
                      font=("Segoe UI", 9), cursor="hand2",
                      bd=0, padx=3).pack(side="left")
            tk.Button(row, text="×", bg=BG, fg=RED, relief="flat",
                      command=lambda t=tag: self._remove_custom(t),
                      font=("Segoe UI", 9), cursor="hand2",
                      bd=0, padx=3).pack(side="left")

    def _add_custom(self) -> None:
        t = self._custom_var.get().strip()
        if not t:
            return
        all_presets = list(PRESET_TAGS) + self._ud.custom_presets
        if self._save_as_preset_var.get():
            # Promote immediately to preset
            if t not in all_presets:
                self._ud.custom_presets.append(t)
                save_userdata(self._ud)
            # Add as checked preset checkbox
            all_presets = list(PRESET_TAGS) + [
                p for p in self._ud.custom_presets if p not in PRESET_TAGS]
            current_checked = [tag for tag, v in self._vars.items() if v.get()]
            current_checked.append(t)
            self._render_preset_grid(all_presets, current_checked)
            # Remove from custom list if it was there
            self._custom_tags = [c for c in self._custom_tags if c != t]
        else:
            if t not in self._custom_tags and t not in all_presets:
                self._custom_tags.append(t)
        self._render_custom()
        self._custom_var.set("")
        self._save_as_preset_var.set(False)

    def _remove_custom(self, tag: str) -> None:
        self._custom_tags = [t for t in self._custom_tags if t != tag]
        self._render_custom()

    def _promote_to_preset(self, tag: str) -> None:
        """Move a custom tag into the preset grid."""
        if tag not in self._ud.custom_presets and tag not in PRESET_TAGS:
            self._ud.custom_presets.append(tag)
            save_userdata(self._ud)
        self._custom_tags = [t for t in self._custom_tags if t != tag]
        # Rebuild preset grid with the tag checked
        all_presets = list(PRESET_TAGS) + [
            p for p in self._ud.custom_presets if p not in PRESET_TAGS]
        current_checked = [t for t, v in self._vars.items() if v.get()]
        current_checked.append(tag)
        self._render_preset_grid(all_presets, current_checked)
        self._render_custom()

    def _open_manage_presets(self) -> None:
        # Collect all tags used anywhere in the library
        all_used: set[str] = set()
        for tag_list in self._ud.tags.values():
            all_used.update(tag_list)
        # Also include custom tags typed in this session
        all_used.update(self._custom_tags)
        dlg = ManagePresetsDialog(self, self._ud, all_used)
        self.wait_window(dlg)
        # Rebuild preset grid in case presets changed
        all_presets = list(PRESET_TAGS) + [
            p for p in self._ud.custom_presets if p not in PRESET_TAGS]
        current_checked = [t for t, v in self._vars.items() if v.get()]
        # Any demoted presets that were checked become custom tags
        for t in list(current_checked):
            if t not in all_presets and t not in self._custom_tags:
                self._custom_tags.append(t)
        self._render_preset_grid(all_presets, current_checked)
        self._render_custom()

    def _save(self) -> None:
        selected = [t for t, v in self._vars.items() if v.get()]
        self.result = selected + self._custom_tags
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# MANAGE PRESETS DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class ManagePresetsDialog(tk.Toplevel):
    """
    Shows every tag used across the library.
    Built-in presets are shown greyed out (cannot be removed).
    User-promoted presets can be demoted back to custom.
    Any used tag can be promoted to a preset.
    """

    def __init__(self, parent, ud: "UserData", all_used: set[str]) -> None:
        super().__init__(parent)
        self.title("Manage Tag Presets")
        self.configure(bg=BG)
        self.geometry("420x480")
        self.resizable(False, True)
        self.transient(parent)
        self.grab_set()
        self._ud = ud

        ttk.Label(self, text="Manage Tag Presets",
                  font=("Segoe UI", 11, "bold"),
                  padding=(14, 12, 14, 2)).pack(anchor="w")
        ttk.Label(self,
                  text="★ = preset (shows as checkbox in tag editor)\n"
                       "Click ★ to promote a custom tag · click ✕ to demote a preset",
                  foreground=FG_DIM, font=("Segoe UI", 8),
                  padding=(14, 0, 14, 8)).pack(anchor="w")

        ttk.Separator(self, orient="horizontal").pack(fill="x")

        # Scrollable list
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=14, pady=8)
        canvas = tk.Canvas(container, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self._inner = tk.Frame(canvas, bg=BG)
        win_id = canvas.create_window((0, 0), window=self._inner, anchor="nw")
        self._inner.bind("<Configure>",
                         lambda e: canvas.configure(
                             scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(win_id, width=e.width))

        # Sort: built-in presets first, then user presets, then custom-only tags
        def _sort_key(t: str):
            if t in PRESET_TAGS:
                return (0, t.lower())
            if t in ud.custom_presets:
                return (1, t.lower())
            return (2, t.lower())

        # All tags = built-in + user presets + tags used in library
        all_tags = sorted(
            (set(PRESET_TAGS) | set(ud.custom_presets) | all_used),
            key=_sort_key)

        self._rows: list[tk.Frame] = []
        for tag in all_tags:
            self._add_row(tag)

        # "Add new preset" entry at bottom
        ttk.Separator(self, orient="horizontal").pack(fill="x")
        add_row = tk.Frame(self, bg=BG)
        add_row.pack(fill="x", padx=14, pady=8)
        ttk.Label(add_row, text="New preset:").pack(side="left")
        self._new_var = tk.StringVar()
        e = ttk.Entry(add_row, textvariable=self._new_var, width=18)
        e.pack(side="left", padx=(6, 4))
        e.bind("<Return>", lambda _: self._add_new())
        ttk.Button(add_row, text="Add", command=self._add_new).pack(side="left")

        ttk.Button(self, text="Done", style="Accent.TButton",
                   command=self.destroy).pack(padx=14, pady=(0, 12), anchor="e")

    def _add_row(self, tag: str) -> None:
        is_builtin  = tag in PRESET_TAGS
        is_promoted = tag in self._ud.custom_presets
        is_preset   = is_builtin or is_promoted

        row = tk.Frame(self._inner, bg=BG)
        row.pack(fill="x", pady=1)

        # Preset star indicator
        star_color = YELLOW if is_preset else FG_DIM
        star_lbl = tk.Label(row, text="★", bg=BG, fg=star_color,
                            font=("Segoe UI", 10), width=2)
        star_lbl.pack(side="left")

        # Tag name
        fg = FG if is_preset else FG_DIM
        tk.Label(row, text=tag, bg=BG, fg=fg,
                 font=("Segoe UI", 9), width=24, anchor="w").pack(side="left")

        if is_builtin:
            tk.Label(row, text="built-in", bg=BG, fg=FG_DIM,
                     font=("Segoe UI", 8, "italic")).pack(side="left", padx=4)
        elif is_promoted:
            # ✕ demote button
            def _demote(t=tag, r=row, sl=star_lbl):
                self._ud.custom_presets = [
                    x for x in self._ud.custom_presets if x != t]
                save_userdata(self._ud)
                sl.configure(fg=FG_DIM)
                # Update the label colour
                for child in r.winfo_children():
                    if isinstance(child, tk.Label) and child.cget("text") == t:
                        child.configure(fg=FG_DIM)
                # Swap ✕ for ★ promote button
                for child in list(r.winfo_children()):
                    if isinstance(child, tk.Button):
                        child.destroy()
                self._add_promote_btn(r, t, sl)
            btn = tk.Button(row, text="✕ demote", bg=BG, fg=RED,
                            relief="flat", font=("Segoe UI", 8),
                            cursor="hand2", bd=0, command=_demote)
            btn.pack(side="left", padx=4)
        else:
            # ★ promote button
            self._add_promote_btn(row, tag, star_lbl)

    def _add_promote_btn(self, row: tk.Frame, tag: str,
                          star_lbl: tk.Label) -> None:
        def _promote(t=tag, r=row, sl=star_lbl):
            if t not in self._ud.custom_presets:
                self._ud.custom_presets.append(t)
                save_userdata(self._ud)
            sl.configure(fg=YELLOW)
            for child in r.winfo_children():
                if isinstance(child, tk.Label) and child.cget("text") == t:
                    child.configure(fg=FG)
            for child in list(r.winfo_children()):
                if isinstance(child, tk.Button):
                    child.destroy()
            # Swap for ✕ demote button
            def _demote(t2=t, r2=r, sl2=sl):
                self._ud.custom_presets = [
                    x for x in self._ud.custom_presets if x != t2]
                save_userdata(self._ud)
                sl2.configure(fg=FG_DIM)
                for child in r2.winfo_children():
                    if isinstance(child, tk.Label) and child.cget("text") == t2:
                        child.configure(fg=FG_DIM)
                for child in list(r2.winfo_children()):
                    if isinstance(child, tk.Button):
                        child.destroy()
                self._add_promote_btn(r2, t2, sl2)
            tk.Button(r, text="✕ demote", bg=BG, fg=RED,
                      relief="flat", font=("Segoe UI", 8),
                      cursor="hand2", bd=0, command=_demote).pack(
                          side="left", padx=4)
        tk.Button(row, text="★ promote", bg=BG, fg=FG_DIM,
                  relief="flat", font=("Segoe UI", 8),
                  cursor="hand2", bd=0, command=_promote).pack(
                      side="left", padx=4)

    def _add_new(self) -> None:
        t = self._new_var.get().strip()
        if not t:
            return
        if t in PRESET_TAGS or t in self._ud.custom_presets:
            self._new_var.set("")
            return
        self._ud.custom_presets.append(t)
        save_userdata(self._ud)
        self._add_row(t)
        self._new_var.set("")


# ══════════════════════════════════════════════════════════════════════════════
# DETAIL PANEL
# ══════════════════════════════════════════════════════════════════════════════

class DetailPanel(ttk.Frame):
    def __init__(self, parent, app: "LibraryApp", **kw) -> None:
        super().__init__(parent, **kw)
        self._app = app
        self._group: GameGroup | None = None
        self._version: GameVersion | None = None
        self._current_key: str | None = None   # for async art guard
        self._detail_photo = None               # keep PhotoImage alive
        self._carousel_paths: list[Path] = []  # all images for current game
        self._carousel_idx: int = 0            # currently shown index
        self._slideshow_id: str | None = None  # after() handle for auto-advance
        self._current_pil: "Image.Image | None" = None   # PIL image on screen for fade
        self._fade_ids: list[str] = []                    # after() handles for fade frames

        self._build()

    def _build(self) -> None:
        # ── Vertical pane: art section (resizable) | rest of detail ──────────
        art_h = self._app._settings.get("detail_art_h", DETAIL_ART_H)
        self._detail_pane = ttk.PanedWindow(self, orient="vertical")
        self._detail_pane.pack(fill="both", expand=True)

        # Top pane — artwork carousel
        art_outer = tk.Frame(self._detail_pane, bg=BG)
        self._detail_pane.add(art_outer, weight=0)

        # Navigation bar packed FIRST so it always claims its space at the bottom
        # before the image label expands — prevents nav from being pushed off-screen
        nav = tk.Frame(art_outer, bg=BG)
        nav.pack(side="bottom", fill="x")

        self._car_prev = tk.Button(
            nav, text="❮", bg=BG, fg=FG, relief="flat",
            font=("Segoe UI", 9, "bold"), cursor="hand2", bd=0,
            activebackground=BG2, activeforeground=ACCENT,
            command=self._carousel_prev)
        self._car_prev.pack(side="left", padx=(4, 0))

        self._car_counter = tk.Label(
            nav, bg=BG, fg=FG_DIM, font=("Segoe UI", 7), text="")
        self._car_counter.pack(side="left", expand=True)

        self._car_next = tk.Button(
            nav, text="❯", bg=BG, fg=FG, relief="flat",
            font=("Segoe UI", 9, "bold"), cursor="hand2", bd=0,
            activebackground=BG2, activeforeground=ACCENT,
            command=self._carousel_next)
        self._car_next.pack(side="right", padx=(0, 4))

        # ── Art label container — packed after nav so it fills remaining space
        art_frame = tk.Frame(art_outer, bg=BG)
        art_frame.pack(side="top", fill="both", expand=True)

        self._art_lbl = tk.Label(
            art_frame, bg=SEL,
            text="", cursor="hand2")
        self._art_lbl.pack(fill="both", expand=True)
        # Left-click → fullscreen viewer; also grab focus for keyboard nav
        self._art_lbl.bind("<Button-1>",
                           lambda e: (self._art_lbl.focus_set(),
                                      self._open_fullscreen(e)))
        # Keyboard arrows when art label is focused
        self._art_lbl.bind("<Left>",  lambda _: self._carousel_prev())
        self._art_lbl.bind("<Right>", lambda _: self._carousel_next())
        # Hover overlay — show/hide the change button
        self._art_lbl.bind("<Enter>", self._on_art_enter)
        self._art_lbl.bind("<Leave>", self._on_art_leave)

        # ── Hover "✎ Change" button (placed over the art label via place())
        self._change_btn = tk.Button(
            art_frame,
            text="✎ Change Image",
            bg=BG3, fg=FG,
            relief="flat", bd=0,
            font=("Segoe UI", 8),
            cursor="hand2",
            activebackground=SEL, activeforeground=ACCENT,
            command=self._set_custom_art)
        # Bind hover on button itself so it doesn't flicker when cursor moves to it
        self._change_btn.bind("<Enter>", self._on_art_enter)
        self._change_btn.bind("<Leave>", self._on_art_leave)
        # Button starts hidden; shown on <Enter>
        self._change_btn_visible = False

        # Apply saved height after window is drawn; persist on sash drag
        self._detail_pane.bind(
            "<Map>",
            lambda e: self.after(60, lambda: self._apply_art_sash(art_h)))
        self._detail_pane.bind(
            "<ButtonRelease-1>",
            lambda e: self._persist_art_sash())

        # Bottom pane — everything else
        inner_wrap = ttk.Frame(self._detail_pane)
        self._detail_pane.add(inner_wrap, weight=1)

        inner = ttk.Frame(inner_wrap)
        inner.pack(fill="both", expand=True, padx=12, pady=(0, 8))

        # ── Title + version selector ─────────────────────────────────────────
        title_row = tk.Frame(inner, bg=BG)
        title_row.pack(fill="x", pady=(4, 0))
        self._title_lbl = tk.Label(
            title_row, bg=BG, fg=ACCENT, text="Select a game",
            font=("Segoe UI", 13, "bold"), anchor="w", wraplength=280)
        self._title_lbl.pack(side="left", fill="x", expand=True)

        self._ver_var = tk.StringVar()
        self._ver_cb = ttk.Combobox(
            title_row, textvariable=self._ver_var,
            state="readonly", width=10)
        self._ver_cb.pack(side="right")
        self._ver_cb.bind("<<ComboboxSelected>>", self._on_ver_changed)

        # ── Stats row ────────────────────────────────────────────────────────
        stats = tk.Frame(inner, bg=BG)
        stats.pack(fill="x", pady=(6, 0))
        self._last_lbl  = self._stat_lbl(stats, "Last played", "—")
        self._time_lbl  = self._stat_lbl(stats, "Play time", "—")
        self._sess_lbl  = self._stat_lbl(stats, "Sessions", "—")

        ttk.Separator(inner, orient="horizontal").pack(fill="x", pady=6)

        # ── Tags ─────────────────────────────────────────────────────────────
        tag_header = tk.Frame(inner, bg=BG)
        tag_header.pack(fill="x")
        tk.Label(tag_header, bg=BG, fg=FG_DIM, text="Tags",
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        ttk.Button(tag_header, text="＋ Edit",
                   command=self._edit_tags).pack(side="right")

        self._tag_frame = tk.Frame(inner, bg=BG)
        self._tag_frame.pack(fill="x", pady=(3, 4))

        ttk.Separator(inner, orient="horizontal").pack(fill="x", pady=4)

        # ── Notes ────────────────────────────────────────────────────────────
        tk.Label(inner, bg=BG, fg=FG_DIM, text="Notes",
                 font=("Segoe UI", 8, "bold")).pack(anchor="w")
        self._notes = tk.Text(
            inner, height=3, bg=BG2, fg=FG, insertbackground=FG,
            relief="flat", font=("Segoe UI", 9), wrap="word",
            padx=4, pady=3)
        self._notes.pack(fill="x", pady=(3, 6))
        self._notes.bind("<FocusOut>", self._save_notes)

        ttk.Separator(inner, orient="horizontal").pack(fill="x", pady=4)

        # ── Actions ──────────────────────────────────────────────────────────
        # ── Launch — full-width, prominent ───────────────────────────────────
        self._btn_launch = ttk.Button(
            inner, text="▶   Launch", style="Launch.TButton",
            command=self._launch)
        self._btn_launch.pack(fill="x", pady=(0, 6))

        ttk.Separator(inner, orient="horizontal").pack(fill="x", pady=(0, 6))

        # ── Six utility buttons in 3×2 grid ──────────────────────────────────
        row1 = ttk.Frame(inner)
        row1.pack(fill="x", pady=(0, 4))
        self._btn_open = ttk.Button(
            row1, text="Open Folder", command=self._open_folder)
        self._btn_open.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self._btn_played = ttk.Button(
            row1, text="Mark Played", command=self._toggle_played)
        self._btn_played.pack(side="left", expand=True, fill="x")

        row2 = ttk.Frame(inner)
        row2.pack(fill="x", pady=(0, 4))
        self._btn_saves = ttk.Button(
            row2, text="Open Saves", command=self._open_saves)
        self._btn_saves.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self._btn_hide = ttk.Button(
            row2, text="Hide Game", command=self._hide)
        self._btn_hide.pack(side="left", expand=True, fill="x")

        row3 = ttk.Frame(inner)
        row3.pack(fill="x")
        self._btn_del_arc = ttk.Button(
            row3, text="Delete Archive",
            style="Danger.TButton", command=self._delete_archive)
        self._btn_del_arc.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self._download_menu = tk.Menu(row3, tearoff=False,
                                      bg=BG2, fg=FG,
                                      activebackground=SEL,
                                      activeforeground=ACCENT, bd=0)
        self._btn_download = ttk.Menubutton(
            row3, text="↗ Download Page",
            menu=self._download_menu, state="disabled")
        self._btn_download.pack(side="left", expand=True, fill="x")

        # ── Metadata row ─────────────────────────────────────────────────────
        ttk.Separator(inner, orient="horizontal").pack(fill="x", pady=(8, 4))
        meta_hdr = tk.Frame(inner, bg=BG)
        meta_hdr.pack(fill="x")
        tk.Label(meta_hdr, bg=BG, fg=FG_DIM, text="Metadata",
                 font=("Segoe UI", 8, "bold")).pack(side="left")
        self._btn_fetch_meta = ttk.Button(
            meta_hdr, text="⬇ Fetch…", command=self._fetch_metadata)
        self._btn_fetch_meta.pack(side="right")
        if not HAS_SCRAPING:
            self._btn_fetch_meta.configure(state="disabled")
            tk.Label(meta_hdr, bg=BG, fg=FG_MUT,
                     text="(pip install curl-cffi beautifulsoup4 lxml)",
                     font=("Segoe UI", 7)).pack(side="right", padx=(0, 4))

        self._meta_source_lbl = tk.Label(
            inner, bg=BG, fg=FG_MUT, text="", font=("Segoe UI", 7), anchor="w")
        self._meta_source_lbl.pack(fill="x")

        # Synopsis expander
        self._synopsis_frame = tk.Frame(inner, bg=BG)
        self._synopsis_frame.pack(fill="x", pady=(2, 0))
        self._synopsis_text = tk.Text(
            self._synopsis_frame, height=4, bg=BG2, fg=FG,
            insertbackground=FG, relief="flat", font=("Segoe UI", 8),
            wrap="word", padx=4, pady=3, state="disabled")
        self._synopsis_text.pack(fill="x")
        self._synopsis_frame.pack_forget()   # hidden until metadata has synopsis

        self._set_enabled(False)

    @staticmethod
    def _stat_lbl(parent, label: str, value: str) -> tk.Label:
        f = tk.Frame(parent, bg=BG)
        f.pack(side="left", expand=True)
        tk.Label(f, bg=BG, fg=FG_DIM, text=label,
                 font=("Segoe UI", 7)).pack()
        val = tk.Label(f, bg=BG, fg=FG, text=value,
                       font=("Segoe UI", 9, "bold"))
        val.pack()
        return val

    # ── Public ───────────────────────────────────────────────────────────────

    def show_group(self, g: GameGroup, ud: UserData,
                   tc: ThumbnailCache) -> None:
        self._group = g
        self._current_key = g.base_key

        # Default to latest version
        self._version = g.versions[-1] if g.versions else None
        self._ver_cb.configure(values=[v.version_str or "—"
                                        for v in g.versions])
        if g.versions:
            self._ver_var.set(g.versions[-1].version_str or "—")
            self._ver_cb.configure(
                state="readonly" if len(g.versions) > 1 else "disabled")
        else:
            self._ver_cb.configure(state="disabled")

        self._title_lbl.configure(text=g.display_name)
        self._set_enabled(True)
        self._refresh_stats(ud)
        self._render_tags(ud.tags.get(g.base_key, []))
        self._load_notes(ud)
        self._load_art(g, ud, tc)
        self._refresh_metadata_display()

    def show_empty(self) -> None:
        self._slideshow_cancel()
        self._fade_cancel()
        self._current_pil = None
        self._group = None
        self._version = None
        self._current_key    = None
        self._carousel_paths = []
        self._carousel_idx   = 0
        self._title_lbl.configure(text="Select a game")
        self._set_enabled(False)
        self._art_lbl.configure(image="", text="", bg=SEL)
        self._detail_photo = None
        self._car_counter.configure(text="")
        self._car_prev.configure(state="disabled")
        self._car_next.configure(state="disabled")

    # ── Internals ────────────────────────────────────────────────────────────

    def _refresh_stats(self, ud: UserData) -> None:
        v = self._version
        if not v:
            return
        lp = ud.last_played.get(v.folder_name, "")
        self._last_lbl.configure(text=fmt_date(lp) if lp else "Never")
        t = ud.playtime.get(v.folder_name, 0)
        self._time_lbl.configure(text=fmt_time(t) if t else "—")
        pc = ud.play_count.get(v.folder_name, 0)
        self._sess_lbl.configure(text=str(pc) if pc else "—")

        played = detect_played(v, ud)
        self._btn_played.configure(
            text="Mark Unplayed" if played else "Mark Played")

        if v.is_renpy:
            has_saves = bool(v.local_save_dir.exists() or
                             (v.appdata_save_dir and v.appdata_save_dir.exists()))
            self._btn_saves.configure(state="normal" if has_saves else "disabled")
        else:
            self._btn_saves.configure(state="disabled")
        self._btn_launch.configure(
            state="normal" if v.exe_path else "disabled")
        # Show engine note in title for non-RenPy games
        engine_note = "" if v.is_renpy else "  [Non-RenPy EXE]"
        name = (self._group.display_name if self._group else "") + engine_note
        self._title_lbl.configure(text=name)
        has_arc = bool(self._group and self._group.archives)
        self._btn_del_arc.configure(
            state="normal" if has_arc else "disabled")

    def _render_tags(self, tags: list[str]) -> None:
        for w in self._tag_frame.winfo_children():
            w.destroy()
        if not tags:
            tk.Label(self._tag_frame, bg=BG, fg=FG_MUT,
                     text="No tags", font=("Segoe UI", 8)).pack(side="left")
            return
        for tag in tags:
            color = TAG_COLORS.get(tag, DEFAULT_TAG_COLOR)
            chip = tk.Label(self._tag_frame, text=tag, bg=color, fg=BG,
                            font=("Segoe UI", 8, "bold"), padx=6, pady=2,
                            cursor="hand2")
            chip.pack(side="left", padx=(0, 4), pady=2)

    def _load_notes(self, ud: UserData) -> None:
        key = self._group.base_key if self._group else ""
        text = ud.notes.get(key, "")
        self._notes.delete("1.0", "end")
        if text:
            self._notes.insert("1.0", text)
        self._notes._notes_key = key  # type: ignore[attr-defined]

    def _load_art(self, g: GameGroup, ud: UserData,
                  tc: ThumbnailCache) -> None:
        if not HAS_PIL:
            return
        # Build carousel path list; reset to first image
        self._fade_cancel()
        self._current_pil    = None   # don't carry over prev game's image
        self._carousel_paths = _group_carousel_paths(g, ud)
        self._carousel_idx   = 0
        self._tc             = tc   # keep reference for navigation
        self._carousel_show(g.base_key)
        self._slideshow_start()
        # Warm the PIL cache for all remaining images so cross-fades are smooth
        # on the very first cycle. Short delay lets layout settle so winfo_* is valid.
        self.after(200, self._preload_carousel)

    def _preload_carousel(self) -> None:
        """Fire background loads for every carousel image that isn't cached yet."""
        if not self._carousel_paths or not HAS_PIL or not hasattr(self, "_tc"):
            return
        w = max(self._art_lbl.winfo_width(),  DETAIL_ART_W)
        h = max(self._art_lbl.winfo_height(), DETAIL_ART_H)
        for idx, path in enumerate(self._carousel_paths):
            key = f"detail:{self._current_key}:{idx}"
            self._tc.request(key, path, w, h, on_ready=lambda *_: None)

    def _carousel_show(self, base_key: str | None = None) -> None:
        """Load and display the image at _carousel_idx."""
        paths = self._carousel_paths
        total = len(paths)
        if not HAS_PIL or not total:
            placeholder = getattr(self, "_tc", None) and self._tc.detail_placeholder()
            if placeholder:
                self._art_lbl.configure(image=placeholder, text="")
                self._detail_photo = placeholder
            else:
                self._art_lbl.configure(image="", text="No artwork",
                                        bg=SEL, fg=FG_DIM)
            self._car_counter.configure(text="")
            self._car_prev.configure(fg=FG_MUT)
            self._car_next.configure(fg=FG_MUT)
            return

        idx  = self._carousel_idx
        path = paths[idx]
        key  = f"detail:{self._current_key}:{idx}"

        self._fade_cancel()
        self._art_lbl.configure(text="Loading…", image="")
        self._detail_photo = None
        self._car_counter.configure(text=f"{idx + 1} / {total}")
        # Arrows always visible; dim when at the boundary but keep clickable
        self._car_prev.configure(fg=FG     if idx > 0        else FG_MUT)
        self._car_next.configure(fg=FG     if idx < total - 1 else FG_MUT)

        w = max(self._art_lbl.winfo_width(),  DETAIL_ART_W)
        h = max(self._art_lbl.winfo_height(), DETAIL_ART_H)
        self._tc.request(key, path, w, h, on_ready=self._on_art_ready)

    def _carousel_prev(self) -> None:
        total = len(self._carousel_paths)
        if total:
            self._carousel_idx = (self._carousel_idx - 1) % total
            self._carousel_show()
            self._slideshow_start()   # reset timer after manual nav

    def _carousel_next(self) -> None:
        total = len(self._carousel_paths)
        if total:
            self._carousel_idx = (self._carousel_idx + 1) % total
            self._carousel_show()
            self._slideshow_start()   # reset timer after manual nav

    # ── Slideshow auto-advance ────────────────────────────────────────────────

    def _slideshow_start(self) -> None:
        """Schedule the next auto-advance tick. Reads interval from settings."""
        self._slideshow_cancel()
        if len(self._carousel_paths) < 2:
            return
        interval_s = self._app._settings.get("slideshow_interval", 3.5)
        ms = max(500, int(float(interval_s) * 1000))
        self._slideshow_id = self.after(ms, self._slideshow_tick)

    def _slideshow_cancel(self) -> None:
        if self._slideshow_id is not None:
            try:
                self.after_cancel(self._slideshow_id)
            except Exception:
                pass
            self._slideshow_id = None

    def _slideshow_tick(self) -> None:
        self._slideshow_id = None
        if not self._carousel_paths:
            return
        self._carousel_idx = (self._carousel_idx + 1) % len(self._carousel_paths)
        self._carousel_show()
        self._slideshow_start()

    def _on_art_ready(self, key: str, photo, pil_img=None) -> None:
        # Guard: key must match current game + current carousel index
        expected = f"detail:{self._current_key}:{self._carousel_idx}"
        if key != expected:
            return
        if not photo:
            self._art_lbl.configure(image="", text="No artwork",
                                    bg=SEL, fg=FG_DIM)
            self._current_pil = None
            return
        old_pil = self._current_pil
        if HAS_PIL and old_pil is not None and pil_img is not None:
            self._fade_transition(old_pil, pil_img, photo)
        else:
            # No previous image or no PIL — show immediately
            self._detail_photo = photo
            self._art_lbl.configure(image=photo, text="")
            self._current_pil = pil_img

    # ── Cross-fade transition ─────────────────────────────────────────────────

    def _fade_cancel(self) -> None:
        """Cancel any in-progress cross-fade."""
        for fid in self._fade_ids:
            try:
                self.after_cancel(fid)
            except Exception:
                pass
        self._fade_ids.clear()

    def _fade_transition(self, old_pil: "Image.Image", new_pil: "Image.Image",
                         new_photo: "ImageTk.PhotoImage",
                         frames: int = 8, step_ms: int = 25) -> None:
        """Cross-fade from old_pil → new_pil over frames×step_ms ms."""
        self._fade_cancel()

        # Ensure both images are the same size for blending
        target_size = new_pil.size
        if old_pil.size != target_size:
            old_pil = old_pil.resize(target_size, Image.LANCZOS)

        # Convert both to RGBA for clean blending
        old_rgba = old_pil.convert("RGBA")
        new_rgba = new_pil.convert("RGBA")

        # Pre-build all intermediate PhotoImages so the main loop stays light
        blend_photos: list["ImageTk.PhotoImage"] = []
        for i in range(1, frames + 1):
            alpha = i / frames
            blended = Image.blend(old_rgba, new_rgba, alpha)
            blend_photos.append(ImageTk.PhotoImage(blended))

        def _apply_frame(idx: int) -> None:
            if idx >= len(blend_photos):
                # Final frame: commit new_photo and PIL image
                self._detail_photo = new_photo
                self._art_lbl.configure(image=new_photo, text="")
                self._current_pil = new_pil
                self._fade_ids.clear()
                return
            self._detail_photo = blend_photos[idx]
            self._art_lbl.configure(image=blend_photos[idx], text="")
            fid = self.after(step_ms, lambda i=idx + 1: _apply_frame(i))
            self._fade_ids.append(fid)

        _apply_frame(0)

    def _set_enabled(self, en: bool) -> None:
        s = "normal" if en else "disabled"
        for b in (self._btn_launch, self._btn_open, self._btn_played,
                  self._btn_saves, self._btn_hide, self._btn_del_arc):
            b.configure(state=s)
        if HAS_SCRAPING:
            self._btn_fetch_meta.configure(state=s)
        if not en:
            self._btn_download.configure(state="disabled")

    # Sources excluded from the download menu (info-only, not download hosts)
    _DOWNLOAD_EXCLUDE = {"vndb"}

    def _refresh_download_menu(self) -> None:
        """Rebuild the Download Page menubutton from stored metadata sources."""
        self._download_menu.delete(0, "end")
        if not self._app.net_ok("allow_download_links"):
            self._btn_download.configure(state="disabled")
            return
        v = self._version
        if not v or not v.metadata:
            self._btn_download.configure(state="disabled")
            return
        sources = v.metadata.get("sources", {})
        entries = [
            (src, data["url"])
            for src, data in sources.items()
            if src not in self._DOWNLOAD_EXCLUDE and data.get("url")
        ]
        if not entries:
            self._btn_download.configure(state="disabled")
            return
        if len(entries) == 1:
            # Single source — skip the menu, open directly on click
            url = entries[0][1]
            self._btn_download.configure(
                state="normal",
                command=lambda u=url: webbrowser.open(u))
            # Menubutton doesn't natively support command= in ttk; wrap via menu
            self._download_menu.add_command(
                label=SOURCE_LABELS.get(entries[0][0], entries[0][0]),
                command=lambda u=url: webbrowser.open(u))
        else:
            for src, url in entries:
                label = SOURCE_LABELS.get(src, src)
                self._download_menu.add_command(
                    label=label,
                    command=lambda u=url: webbrowser.open(u))
        self._btn_download.configure(state="normal")

    # ── Metadata ─────────────────────────────────────────────────────────────

    def _refresh_metadata_display(self) -> None:
        """Show synopsis and source line if metadata present for this version."""
        self._refresh_download_menu()
        v = self._version
        if not v or not v.metadata:
            self._meta_source_lbl.configure(text="")
            self._synopsis_frame.pack_forget()
            return

        meta = v.metadata
        sources = meta.get("sources", {})
        src_names = [SOURCE_LABELS.get(s, s) for s in sources]
        self._meta_source_lbl.configure(
            text="Sources: " + ", ".join(src_names) if src_names else "")

        synopsis = meta.get("synopsis", "")
        if synopsis:
            self._synopsis_text.configure(state="normal")
            self._synopsis_text.delete("1.0", "end")
            self._synopsis_text.insert("1.0", synopsis)
            self._synopsis_text.configure(state="disabled")
            self._synopsis_frame.pack(fill="x", pady=(2, 0))
        else:
            self._synopsis_frame.pack_forget()

    def _fetch_metadata(self) -> None:
        v = self._version
        if not v:
            return
        if not self._app.net_ok("fetch_metadata"):
            messagebox.showinfo(
                "Disabled",
                "Metadata fetch is disabled.\n\n"
                "Enable it in Settings → Connection.",
                parent=self)
            return
        if not HAS_SCRAPING:
            messagebox.showinfo(
                "Scraping Unavailable",
                "curl_cffi and beautifulsoup4 are not installed.\n\n"
                "Run:  pip install curl-cffi beautifulsoup4 lxml",
                parent=self,
            )
            return

        # ── Search query editor ──────────────────────────────────────────────
        root = self._app
        dlg = tk.Toplevel(root)
        dlg.title("Fetch Metadata")
        dlg.configure(bg=BG)
        dlg.resizable(False, False)
        dlg.transient(root)

        pad = {"padx": 16}

        tk.Label(dlg, bg=BG, fg=FG_DIM,
                 text="Search query  (edit if the title looks wrong):",
                 font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(12, 0))

        # Use the stored title but clean any leftover [bracket] conventions
        # (old saves may contain raw F95Zone thread-title format)
        _stored = v.metadata.get("game_title", "")
        default_title = (
            re.sub(r"\s*\[[^\]]+\]", "", _stored).strip()
            if _stored else v.display_name
        )
        query_var = tk.StringVar(value=default_title)
        entry = ttk.Entry(dlg, textvariable=query_var, width=44)
        entry.pack(anchor="w", padx=16, pady=(4, 0))
        entry.select_range(0, "end")
        entry.focus_set()

        tk.Label(dlg, bg=BG, fg=FG_DIM, text="Sources:",
                 font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=(10, 2))

        src_vars = {
            "vndb":       tk.BooleanVar(value=True),
            "itchio":     tk.BooleanVar(value=True),
            "f95zone":    tk.BooleanVar(value=bool(load_site_cookies("f95zone"))),
            "lewdcorner": tk.BooleanVar(value=bool(load_site_cookies("lewdcorner"))),
        }
        for src, var in src_vars.items():
            label = SOURCE_LABELS.get(src, src)
            if src == "vndb":
                suffix = ""
            elif src == "itchio":
                suffix = "  ✓" if load_site_cookies(src) else "  (log in for adult content)"
            else:
                suffix = "  ✓" if load_site_cookies(src) else "  (not logged in)"
            tk.Checkbutton(dlg, bg=BG, fg=FG, selectcolor=BG2,
                           activebackground=BG, activeforeground=FG,
                           text=label + suffix, variable=var).pack(
                               anchor="w", padx=24)

        login_lbl = tk.Label(dlg, bg=BG, fg=ACCENT, cursor="hand2",
                             text="Manage site logins →",
                             font=("Segoe UI", 8, "underline"))
        login_lbl.pack(anchor="w", padx=16, pady=(6, 0))
        login_lbl.bind("<Button-1>", lambda _: (dlg.destroy(),
                                                 root.open_cookie_settings()))

        btn_row = tk.Frame(dlg, bg=BG)
        btn_row.pack(fill="x", padx=16, pady=12)
        search_clicked = [False]

        def _do_search():
            search_clicked[0] = True
            dlg.destroy()

        ttk.Button(btn_row, text="Search", style="Accent.TButton",
                   command=_do_search).pack(side="left", padx=(0, 8))
        ttk.Button(btn_row, text="Cancel",
                   command=dlg.destroy).pack(side="left")

        dlg.bind("<Return>", lambda _: _do_search())
        dlg.update_idletasks()
        px = root.winfo_rootx() + root.winfo_width()  // 2 - dlg.winfo_width()  // 2
        py = root.winfo_rooty() + root.winfo_height() // 2 - dlg.winfo_height() // 2
        dlg.geometry(f"+{max(0,px)}+{max(0,py)}")
        dlg.grab_set()
        dlg.wait_window()

        if not search_clicked[0]:
            return

        search_title = query_var.get().strip() or default_title
        sources = [s for s, var in src_vars.items() if var.get()]
        if not sources:
            return

        # ── Fire the fetcher ──────────────────────────────────────────────────
        picker   = MetadataPickerDialog(self.winfo_toplevel(), v, self._app)
        progress = FetchProgressDialog(self.winfo_toplevel(), search_title, sources)

        def on_result(source: str,
                      result: "list[MetadataCandidate] | Exception") -> None:
            if progress.winfo_exists():
                if isinstance(result, Exception):
                    msg = str(result)[:60]
                    progress.after(0, lambda s=source, m=msg: progress.set_status(
                        s, f"Error: {m}", RED))
                else:
                    n = len(result)
                    progress.after(0, lambda s=source, n=n: progress.set_status(
                        s, f"{n} result{'s' if n != 1 else ''}", GREEN))
            picker.populate_source(source, result)

        def on_done() -> None:
            if progress.winfo_exists():
                progress.after(0, progress.destroy)

        MetadataFetcher(search_title, sources, on_result, on_done).start()

    # ── Actions ──────────────────────────────────────────────────────────────

    def _on_ver_changed(self, _e=None) -> None:
        if not self._group:
            return
        sel = self._ver_var.get()
        for v in self._group.versions:
            if (v.version_str or "—") == sel:
                self._version = v
                break
        self._refresh_stats(self._app.user_data)
        self._refresh_metadata_display()

    def _launch(self) -> None:
        v = self._version
        g = self._group
        if not v or not g:
            return
        if self._app.play_tracker.launch(v, self._app):
            ud = self._app.user_data
            ud.last_played[v.folder_name] = datetime.datetime.now().isoformat()
            ud.play_count[v.folder_name] = ud.play_count.get(
                v.folder_name, 0) + 1
            save_userdata(ud)
            card = self._app.card_list._cards.get(g.base_key)
            if card:
                card.set_playing(True)
            self._btn_launch.configure(state="disabled", text="▶  Playing…")

    def _open_folder(self) -> None:
        v = self._version
        if v:
            os.startfile(str(v.folder_path))

    def _toggle_played(self) -> None:
        v = self._version
        if not v:
            return
        ud = self._app.user_data
        if detect_played(v, ud):
            ud.manual_played.discard(v.folder_name)
            ud.manual_unplayed.add(v.folder_name)
        else:
            ud.manual_unplayed.discard(v.folder_name)
            ud.manual_played.add(v.folder_name)
        save_userdata(ud)
        self._refresh_stats(ud)
        if self._group:
            self._app.card_list.update_card(self._group.base_key, ud)
        self._app._update_tab_titles()

    def _open_saves(self) -> None:
        v = self._version
        if not v:
            return
        target = (v.appdata_save_dir if v.appdata_save_dir
                  and v.appdata_save_dir.exists() else None) \
            or (v.local_save_dir if v.local_save_dir.exists() else None)
        if target:
            os.startfile(str(target))
        else:
            messagebox.showinfo("No Saves", "No saves folder found.")

    def _hide(self) -> None:
        v = self._version
        if not v:
            return
        ud = self._app.user_data
        name = v.folder_name
        if name in ud.hidden:
            ud.hidden.discard(name)
            self._btn_hide.configure(text="Hide Game")
        else:
            ud.hidden.add(name)
            self._btn_hide.configure(text="Unhide Game")
        save_userdata(ud)
        self._app._apply_filters()

    def _delete_archive(self) -> None:
        g = self._group
        if not g or not g.archives:
            return
        self._app._action_delete_archives(g)

    def _edit_tags(self) -> None:
        g = self._group
        if not g:
            return
        ud = self._app.user_data
        current = ud.tags.get(g.base_key, [])
        dlg = TagPickerDialog(self, current, ud)
        self.wait_window(dlg)
        if dlg.result is not None:
            if dlg.result:
                ud.tags[g.base_key] = dlg.result
            else:
                ud.tags.pop(g.base_key, None)
            save_userdata(ud)
            self._render_tags(ud.tags.get(g.base_key, []))
            # Update tag filter in parent
            self._app._rebuild_tag_filter()

    def _save_notes(self, _e=None) -> None:
        g = self._group
        if not g:
            return
        text = self._notes.get("1.0", "end-1c").strip()
        ud = self._app.user_data
        if text:
            ud.notes[g.base_key] = text
        else:
            ud.notes.pop(g.base_key, None)
        save_userdata(ud)

    # ── Art sash persistence ──────────────────────────────────────────────────

    def _apply_art_sash(self, h: int) -> None:
        try:
            self._detail_pane.update_idletasks()
            total = self._detail_pane.winfo_height()
            if total > h + 60:
                self._detail_pane.sashpos(0, h)
        except Exception:
            pass

    def _persist_art_sash(self) -> None:
        try:
            pos = self._detail_pane.sashpos(0)
            if pos > 0:
                self._app._settings["detail_art_h"] = pos
                save_settings(self._app._settings)
        except Exception:
            pass

    # ── Hover overlay ─────────────────────────────────────────────────────────

    def _on_art_enter(self, _e=None) -> None:
        if not self._change_btn_visible and self._group:
            self._change_btn_visible = True
            self._change_btn.place(relx=1.0, rely=1.0, anchor="se", x=-4, y=-4)
            self._change_btn.lift()

    def _on_art_leave(self, _e=None) -> None:
        # Check pointer is truly outside the art+button area before hiding
        try:
            px, py = self._art_lbl.winfo_pointerxy()
            wx  = self._art_lbl.winfo_rootx()
            wy  = self._art_lbl.winfo_rooty()
            ww  = self._art_lbl.winfo_width()
            wh  = self._art_lbl.winfo_height()
            if wx <= px <= wx + ww and wy <= py <= wy + wh:
                return
            bx  = self._change_btn.winfo_rootx()
            by  = self._change_btn.winfo_rooty()
            bw  = self._change_btn.winfo_width()
            bh  = self._change_btn.winfo_height()
            if bx <= px <= bx + bw and by <= py <= by + bh:
                return
        except Exception:
            pass
        self._change_btn_visible = False
        self._change_btn.place_forget()

    # ── Fullscreen viewer ─────────────────────────────────────────────────────

    def _open_fullscreen(self, _e=None) -> None:
        if not self._carousel_paths or not HAS_PIL:
            return

        sw = self._app.winfo_screenwidth()
        sh = self._app.winfo_screenheight()
        max_w, max_h = sw - 80, sh - 120

        # Mutable index shared across closures
        state = {"idx": self._carousel_idx}

        win = tk.Toplevel(self._app)
        win.configure(bg=BG3)
        win.grab_set()
        win.focus_set()

        # Image label
        img_lbl = tk.Label(win, bg=BG3, cursor="hand2")
        img_lbl.pack(fill="both", expand=True, padx=20, pady=(20, 8))
        img_lbl.bind("<Button-1>", lambda _: win.destroy())

        # Nav bar
        nav = tk.Frame(win, bg=BG3)
        nav.pack(fill="x", pady=(0, 12))

        def _load(idx: int) -> None:
            paths = self._carousel_paths
            total = len(paths)
            if not total:
                return
            state["idx"] = idx % total
            i = state["idx"]
            try:
                raw = Image.open(paths[i]).convert("RGB")
            except Exception:
                img_lbl.configure(image="", text="Could not load image", fg=FG_DIM)
                return
            raw.thumbnail((max_w, max_h), Image.LANCZOS)
            photo = ImageTk.PhotoImage(raw)
            img_lbl.configure(image=photo, text="")
            img_lbl.image = photo  # type: ignore[attr-defined]
            counter.configure(text=f"{i + 1} / {total}")
            prev_btn.configure(fg=FG if total > 1 else FG_MUT)
            next_btn.configure(fg=FG if total > 1 else FG_MUT)
            win.title(paths[i].name)
            # Re-centre after first load
            win.update_idletasks()
            x = (sw - win.winfo_width())  // 2
            y = (sh - win.winfo_height()) // 2
            win.geometry(f"+{x}+{y}")

        prev_btn = tk.Button(
            nav, text="❮", bg=BG3, fg=FG, relief="flat",
            font=("Segoe UI", 14, "bold"), cursor="hand2", bd=0,
            activebackground=BG3, activeforeground=ACCENT,
            command=lambda: _load(state["idx"] - 1))
        prev_btn.pack(side="left", padx=(20, 0))

        counter = tk.Label(nav, bg=BG3, fg=FG_DIM, font=("Segoe UI", 9), text="")
        counter.pack(side="left", expand=True)

        next_btn = tk.Button(
            nav, text="❯", bg=BG3, fg=FG, relief="flat",
            font=("Segoe UI", 14, "bold"), cursor="hand2", bd=0,
            activebackground=BG3, activeforeground=ACCENT,
            command=lambda: _load(state["idx"] + 1))
        next_btn.pack(side="right", padx=(0, 20))

        win.bind("<Left>",  lambda _: _load(state["idx"] - 1))
        win.bind("<Right>", lambda _: _load(state["idx"] + 1))
        win.bind("<Escape>", lambda _: win.destroy())

        _load(state["idx"])

    # ── Image management ──────────────────────────────────────────────────────

    def _set_custom_art(self, _e=None) -> None:
        g = self._group
        if not g:
            return
        paths = filedialog.askopenfilenames(
            title="Select image(s)  —  first selected becomes the cover",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp"),
                       ("All files", "*.*")])
        if not paths:
            return

        ud  = self._app.user_data
        tc  = self._app.thumb_cache

        # First file → cover
        cover_path = paths[0]
        ud.custom_art[g.base_key] = cover_path
        save_userdata(ud)

        # Additional files → save as screenshots in .vnpf/ of the latest version
        if len(paths) > 1 and g.versions:
            vnpf = _vnpf_dir(g.versions[-1].folder_path)
            vnpf.mkdir(parents=True, exist_ok=True)
            # Find next free screenshot index
            existing = sorted(vnpf.glob("screenshot_*.*"), key=lambda p: p.stem)
            next_idx = len(existing) + 1
            for extra in paths[1:]:
                src = Path(extra)
                dest = vnpf / f"screenshot_{next_idx:03d}{src.suffix.lower()}"
                try:
                    shutil.copy2(src, dest)
                    next_idx += 1
                except OSError:
                    pass

        # Bust cache and reload
        for k in list(tc._cache):
            if k == g.base_key or k.startswith(f"detail:{g.base_key}"):
                tc._cache.pop(k, None)
        self._load_art(g, ud, tc)
        art = _group_art_path(g, ud)
        if art:
            tc.request(g.base_key, art, THUMB_W, THUMB_H,
                       on_ready=self._app.card_list._on_thumb_ready)


# ══════════════════════════════════════════════════════════════════════════════
# EXTRACTION QUEUE
# ══════════════════════════════════════════════════════════════════════════════

_job_id_seq = itertools.count()


@dataclass
class ExtractJob:
    archive:      Archive
    # ── mutable state (written by worker, read by UI via .after()) ──
    status:       str   = "queued"   # queued|extracting|done|failed|cancelled
    progress_pct: int   = 0
    bytes_done:   int   = 0
    total_bytes:  int   = 0
    speed_bps:    float = 0.0
    current_file: str   = ""
    error:        str | None       = None
    extracted_path: Path | None    = None
    cancelled:    bool  = False
    job_id:       int   = field(default_factory=lambda: next(_job_id_seq))


def _fmt_bytes(b: int) -> str:
    if b < 1024 ** 2:
        return f"{b / 1024:.0f} KB"
    if b < 1024 ** 3:
        return f"{b / 1024**2:.1f} MB"
    return f"{b / 1024**3:.2f} GB"


class ExtractionQueueWindow(tk.Toplevel):
    """
    Persistent, non-modal extraction queue.
    Lives for the lifetime of the app — hidden rather than destroyed.
    Can be moved to a second monitor while the main window stays responsive.
    """

    CHUNK = 4 * 1024 * 1024   # 4 MB read chunks for byte-accurate progress

    def __init__(self, app: "LibraryApp") -> None:
        super().__init__(app)
        self._app = app
        self.title("Extraction Queue")
        self.configure(bg=BG)
        self.geometry("640x420")
        self.minsize(500, 240)
        # Closing hides rather than destroys so the queue keeps running
        self.protocol("WM_DELETE_WINDOW", self.withdraw)

        self._jobs:         list[ExtractJob]    = []
        self._row_widgets:  dict[int, dict]     = {}  # job_id → widget dict
        self._worker_active = False
        self._lock = threading.Lock()

        self._build()

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build(self) -> None:
        # Header bar
        hdr = ttk.Frame(self, padding=(12, 8))
        hdr.pack(fill="x")
        ttk.Label(hdr, text="Extraction Queue",
                  font=("Segoe UI", 12, "bold"),
                  foreground=ACCENT).pack(side="left")
        self._hdr_lbl = ttk.Label(hdr, text="", foreground=FG_DIM)
        self._hdr_lbl.pack(side="left", padx=10)
        ttk.Button(hdr, text="Clear Completed",
                   command=self._clear_completed).pack(side="right", padx=(4, 0))
        ttk.Button(hdr, text="Clear All & Delete ZIPs",
                   style="Danger.TButton",
                   command=self._clear_all_and_delete).pack(side="right", padx=(0, 4))

        ttk.Separator(self, orient="horizontal").pack(fill="x")

        # Scrollable job list
        c_frame = ttk.Frame(self)
        c_frame.pack(fill="both", expand=True)
        self._canvas = tk.Canvas(c_frame, bg=BG, highlightthickness=0)
        vsb = ttk.Scrollbar(c_frame, orient="vertical",
                            command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        self._inner = tk.Frame(self._canvas, bg=BG)
        self._win_id = self._canvas.create_window(
            (0, 0), window=self._inner, anchor="nw")
        self._inner.bind(
            "<Configure>",
            lambda e: self._canvas.configure(
                scrollregion=self._canvas.bbox("all")))
        self._canvas.bind(
            "<Configure>",
            lambda e: self._canvas.itemconfigure(self._win_id, width=e.width))

        self._empty_lbl = ttk.Label(
            self._inner,
            text="No extractions queued.  Use the Archives tab to add ZIPs.",
            foreground=FG_DIM, padding=(16, 24))
        self._empty_lbl.pack()

    def _add_row_ui(self, job: ExtractJob) -> None:
        self._empty_lbl.pack_forget()

        size_bytes = job.archive.archive_path.stat().st_size \
            if job.archive.archive_path.exists() else 0
        size_str = _fmt_bytes(size_bytes)

        outer = tk.Frame(self._inner, bg=BG2, pady=0)
        outer.pack(fill="x", padx=8, pady=(6, 0))

        # ── Row 1: name + size + action button ──
        r1 = tk.Frame(outer, bg=BG2)
        r1.pack(fill="x", padx=10, pady=(8, 2))
        tk.Label(r1, text=job.archive.archive_path.name,
                 bg=BG2, fg=FG, font=("Segoe UI", 9, "bold"),
                 anchor="w").pack(side="left")
        tk.Label(r1, text=size_str, bg=BG2, fg=FG_DIM,
                 font=("Segoe UI", 8)).pack(side="left", padx=(8, 0))
        action_btn = tk.Button(r1, text="Cancel", bg=BG2, fg=FG_DIM,
                               relief="flat", font=("Segoe UI", 8),
                               cursor="hand2", bd=0,
                               command=lambda j=job: self._cancel_job(j))
        action_btn.pack(side="right")

        # ── Row 2: progress bar ──
        bar = ttk.Progressbar(outer, orient="horizontal",
                              mode="determinate", maximum=100)
        bar.pack(fill="x", padx=10, pady=(2, 0))

        # ── Row 3: status / speed / ETA ──
        status_lbl = tk.Label(outer, text="Queued", bg=BG2, fg=FG_DIM,
                              font=("Segoe UI", 8), anchor="w")
        status_lbl.pack(fill="x", padx=10, pady=(1, 0))

        # ── Row 4: current filename ──
        file_lbl = tk.Label(outer, text="", bg=BG2, fg=FG_DIM,
                            font=("Segoe UI", 7), anchor="w")
        file_lbl.pack(fill="x", padx=10, pady=(0, 8))

        # ── Row 5: post-completion buttons (hidden until done) ──
        btn_row = tk.Frame(outer, bg=BG2)
        # populated dynamically in _update_row_ui when status → done

        tk.Frame(self._inner, bg=BG, height=1).pack(fill="x", padx=8)

        self._row_widgets[job.job_id] = {
            "outer": outer, "bar": bar,
            "status_lbl": status_lbl, "file_lbl": file_lbl,
            "btn_row": btn_row,
            "action_btn": action_btn,
        }
        self._update_header()

    def _update_row_ui(self, job: ExtractJob) -> None:
        w = self._row_widgets.get(job.job_id)
        if not w:
            return

        bar, status_lbl, file_lbl = w["bar"], w["status_lbl"], w["file_lbl"]
        action_btn, btn_row = w["action_btn"], w["btn_row"]

        if job.status == "extracting":
            bar.configure(value=job.progress_pct)
            parts = [f"{job.progress_pct}%",
                     f"{_fmt_bytes(job.bytes_done)} / {_fmt_bytes(job.total_bytes)}"]
            if job.speed_bps > 0:
                parts.append(f"{_fmt_bytes(int(job.speed_bps))}/s")
                remaining = job.total_bytes - job.bytes_done
                eta_s = remaining / job.speed_bps
                parts.append(f"~{eta_s:.0f}s" if eta_s < 60
                              else f"~{eta_s/60:.0f}m remaining")
            status_lbl.configure(text="  ".join(parts), foreground=FG_DIM)
            name = job.current_file
            if len(name) > 72:
                name = "…" + name[-69:]
            file_lbl.configure(text=f"↳ {name}" if name else "")

        elif job.status == "done":
            bar.configure(value=100)
            status_lbl.configure(text="Done ✓", foreground=GREEN)
            file_lbl.configure(text="")
            action_btn.configure(state="disabled", text="")  # hide cancel
            # Show post-completion buttons
            for child in btn_row.winfo_children():
                child.destroy()
            tk.Button(
                btn_row, text="Clear",
                bg=BG2, fg=ACCENT, relief="flat",
                font=("Segoe UI", 8, "bold"), cursor="hand2", bd=0,
                padx=10, pady=4,
                command=lambda j=job: self._remove_row(j)
            ).pack(side="left", padx=(0, 6))
            tk.Button(
                btn_row, text="Clear & Delete ZIP",
                bg=BG2, fg=RED, relief="flat",
                font=("Segoe UI", 8, "bold"), cursor="hand2", bd=0,
                padx=10, pady=4,
                command=lambda j=job: self._clear_and_delete(j)
            ).pack(side="left")
            btn_row.pack(anchor="w", padx=10, pady=(0, 8))
            self._recolor_row(w, BG)

        elif job.status == "failed":
            bar.configure(value=0)
            err = (job.error or "Unknown error")
            if len(err) > 80:
                err = err[:77] + "…"
            status_lbl.configure(text=f"Failed: {err}", foreground=RED)
            file_lbl.configure(text="")
            action_btn.configure(
                text="Remove", fg=FG_DIM,
                command=lambda j=job: self._remove_row(j))

        elif job.status == "cancelled":
            bar.configure(value=0)
            status_lbl.configure(text="Cancelled", foreground=FG_DIM)
            file_lbl.configure(text="")
            action_btn.configure(
                text="Remove", fg=FG_DIM,
                command=lambda j=job: self._remove_row(j))
            self._recolor_row(w, BG)

    @staticmethod
    def _recolor_row(w: dict, color: str) -> None:
        w["outer"].configure(bg=color)
        for child in w["outer"].winfo_children():
            try:
                child.configure(bg=color)
            except tk.TclError:
                pass
            if isinstance(child, tk.Frame):
                for grandchild in child.winfo_children():
                    try:
                        grandchild.configure(bg=color)
                    except tk.TclError:
                        pass

    def _remove_row(self, job: ExtractJob) -> None:
        w = self._row_widgets.pop(job.job_id, None)
        if w:
            w["outer"].destroy()
        with self._lock:
            self._jobs = [j for j in self._jobs if j.job_id != job.job_id]
        if not self._row_widgets:
            self._empty_lbl.pack()
        self._update_header()

    def _clear_and_delete(self, job: ExtractJob) -> None:
        """Delete the source archive then remove the row."""
        path = job.archive.archive_path
        if path.exists():
            try:
                path.unlink()
            except OSError as e:
                messagebox.showerror("Delete Failed", str(e), parent=self)
                return
        self._remove_row(job)
        self._app.refresh()   # update Archives tab (archive no longer listed)

    def _clear_completed(self) -> None:
        for job in list(self._jobs):
            if job.status in ("done", "failed", "cancelled"):
                self._remove_row(job)

    def _clear_all_and_delete(self) -> None:
        """Delete the archive file for every completed job, then clear all rows."""
        done_jobs = [j for j in self._jobs if j.status == "done"]
        if not done_jobs:
            self._clear_completed()   # still clears failed/cancelled rows
            return
        errors: list[str] = []
        for job in done_jobs:
            path = job.archive.archive_path
            if path.exists():
                try:
                    path.unlink()
                except OSError as e:
                    errors.append(f"{path.name}: {e}")
        for job in list(self._jobs):
            if job.status in ("done", "failed", "cancelled"):
                self._remove_row(job)
        if errors:
            messagebox.showerror(
                "Some Deletes Failed",
                "\n".join(errors),
                parent=self)
        self._app.refresh()

    def _update_header(self) -> None:
        with self._lock:
            counts = collections.Counter(j.status for j in self._jobs)
        parts = []
        if counts["extracting"]: parts.append(f"{counts['extracting']} extracting")
        if counts["queued"]:     parts.append(f"{counts['queued']} queued")
        if counts["done"]:       parts.append(f"{counts['done']} done")
        self._hdr_lbl.configure(text=" · ".join(parts) if parts else "")

    # ── Queue management ──────────────────────────────────────────────────────

    def add_job(self, archive: Archive) -> None:
        job = ExtractJob(archive=archive)
        with self._lock:
            self._jobs.append(job)
        self.after(0, lambda j=job: self._add_row_ui(j))
        self._start_worker_if_needed()
        self.deiconify()
        self.lift()

    def is_active(self, archive_path: Path) -> bool:
        """True if this archive is currently queued or extracting."""
        with self._lock:
            return any(
                j.archive.archive_path == archive_path
                and j.status in ("queued", "extracting")
                for j in self._jobs)

    def _cancel_job(self, job: ExtractJob) -> None:
        if job.status == "queued":
            job.status = "cancelled"
            job.cancelled = True
            self.after(0, lambda j=job: self._update_row_ui(j))
            self.after(0, self._update_header)
        elif job.status == "extracting":
            job.cancelled = True  # worker checks this between chunks

    # ── Worker ────────────────────────────────────────────────────────────────

    def _start_worker_if_needed(self) -> None:
        with self._lock:
            if self._worker_active:
                return
            self._worker_active = True
        threading.Thread(target=self._worker_loop, daemon=True).start()

    def _worker_loop(self) -> None:
        while True:
            job = None
            with self._lock:
                for j in self._jobs:
                    if j.status == "queued":
                        j.status = "extracting"
                        job = j
                        break
            if job is None:
                with self._lock:
                    self._worker_active = False
                self.after(0, self._update_header)
                return
            self.after(0, lambda j=job: self._update_row_ui(j))
            self._process_job(job)
            self.after(0, lambda j=job: self._on_job_done(j))

    def _process_job(self, job: ExtractJob) -> None:
        try:
            with zipfile.ZipFile(job.archive.archive_path) as zf:
                all_entries = zf.infolist()
                file_entries = [e for e in all_entries if not e.is_dir()]
                job.total_bytes = sum(e.file_size for e in file_entries)

                # Detect single vs multi root folder
                roots = {
                    e.filename.replace("\\", "/").split("/")[0]
                    for e in all_entries
                    if e.filename.replace("\\", "/").split("/")[0]
                }
                if len(roots) == 1:
                    dest = RENPY_DIR
                    job.extracted_path = RENPY_DIR / next(iter(roots))
                else:
                    job.extracted_path = RENPY_DIR / job.archive.archive_path.stem
                    dest = job.extracted_path
                    dest.mkdir(exist_ok=True)

                speed_samples: collections.deque[tuple[int, float]] = \
                    collections.deque(maxlen=8)
                last_t = time.monotonic()
                bytes_at_last = 0
                last_ui_t = last_t

                for entry in all_entries:
                    if job.cancelled:
                        job.status = "cancelled"
                        return

                    norm = entry.filename.replace("\\", "/")
                    out = dest / norm
                    if entry.is_dir() or norm.endswith("/"):
                        out.mkdir(parents=True, exist_ok=True)
                        continue

                    out.parent.mkdir(parents=True, exist_ok=True)
                    job.current_file = Path(norm).name

                    with zf.open(entry) as src, open(out, "wb") as dst:
                        while True:
                            if job.cancelled:
                                job.status = "cancelled"
                                return
                            chunk = src.read(self.CHUNK)
                            if not chunk:
                                break
                            dst.write(chunk)
                            job.bytes_done += len(chunk)

                            now = time.monotonic()
                            if now - last_ui_t >= 0.20:   # 5 fps UI updates
                                dt = now - last_t
                                db = job.bytes_done - bytes_at_last
                                if dt > 0:
                                    speed_samples.append((db, dt))
                                tot_b = sum(s[0] for s in speed_samples)
                                tot_t = sum(s[1] for s in speed_samples)
                                job.speed_bps = tot_b / tot_t if tot_t > 0 else 0
                                if job.total_bytes > 0:
                                    job.progress_pct = min(
                                        99, int(job.bytes_done * 100
                                                / job.total_bytes))
                                last_t = now
                                bytes_at_last = job.bytes_done
                                last_ui_t = now
                                self.after(0, lambda j=job: self._update_row_ui(j))

            job.progress_pct = 100
            job.status = "done"

        except zipfile.BadZipFile:
            job.status = "failed"
            job.error = "Invalid or corrupted ZIP file."
        except Exception as exc:
            job.status = "failed"
            job.error = str(exc)

    def _on_job_done(self, job: ExtractJob) -> None:
        self._update_row_ui(job)
        self._update_header()
        if job.status == "done":
            self._app.refresh()


# ══════════════════════════════════════════════════════════════════════════════
# ORPHAN SCANNER & PATCH UTILITIES
# ══════════════════════════════════════════════════════════════════════════════

# Root-level items that are always safe to keep
SAFE_ROOT_NAMES = {
    # app own subfolder (all app files live here now)
    "_VNPathfinder",
    # legacy app files at games root (pre-reorganisation)
    "vn_pathfinder.py", "vn_pathfinder.json",
    "renpy_manager.py", "renpy_manager.json",
    "vn_pathfinder.spec",
    "requirements.txt", "build.bat",
    "installer.iss", "LICENSE", "README.md",
    "assets", "build", "dist", "docs",
    "__pycache__", ".github", ".git",
    # common RenPy root clutter
    "lib", "win", "renpy",
    "orpheus_voice.rpy", "orpheus_voice_guide.html",
}


def find_orphans(groups: list[GameGroup]) -> list[Path]:
    """Return root-level items not recognised as a game, archive, or utility."""
    known_folders  = {v.folder_path.name for g in groups for v in g.versions}
    known_archives = {a.archive_path.name for g in groups for a in g.archives}
    orphans: list[Path] = []
    try:
        for item in sorted(RENPY_DIR.iterdir(), key=lambda p: p.name.lower()):
            name = item.name
            if name in SAFE_ROOT_NAMES or name.startswith("."):
                continue
            if item.is_dir() and name in known_folders:
                continue
            if item.is_file() and name in known_archives:
                continue
            orphans.append(item)
    except OSError:
        pass
    return orphans


def _guess_patch_game(archive_path: Path, groups: list[GameGroup]) -> GameGroup | None:
    """Try to identify which game a patch archive belongs to by name similarity."""
    stem = archive_path.stem.lower()
    stem_norm = re.sub(r"[-_\s]+", "", stem)
    best: tuple[int, GameGroup | None] = (0, None)
    for g in groups:
        if not g.versions:
            continue
        # Score = length of longest common prefix between normalised names
        score = 0
        for i, (a, b) in enumerate(zip(stem_norm, g.base_key)):
            if a == b:
                score = i + 1
            else:
                break
        # Also check if game name appears anywhere inside the stem
        if g.base_key in stem_norm and len(g.base_key) > score:
            score = len(g.base_key)
        if score > best[0]:
            best = (score, g)
    return best[1] if best[0] >= 4 else None


def apply_patch_from_folder(source_dir: Path,
                             target_version: GameVersion) -> list[str]:
    """
    Copy patch files from source_dir into the game folder.
    Returns list of relative paths that were written.
    Handles two structures:
      - source has a game/ subfolder → merge into target/game/
      - source has only loose files   → copy to target root
    """
    game_root = target_version.folder_path
    copied: list[str] = []

    source_game_sub = source_dir / "game"
    if source_game_sub.is_dir():
        # Merge game/ contents
        for item in source_game_sub.rglob("*"):
            if item.is_file():
                rel = item.relative_to(source_game_sub)
                dest = game_root / "game" / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(item, dest)
                copied.append(f"game/{rel}")
        # Also copy any root-level loose files from source
        for item in source_dir.iterdir():
            if item.is_file():
                dest = game_root / item.name
                shutil.copy2(item, dest)
                copied.append(item.name)
    else:
        # All files go to game root (typical .py patch)
        for item in source_dir.rglob("*"):
            if item.is_file():
                rel = item.relative_to(source_dir)
                dest = game_root / rel
                dest.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(item, dest)
                copied.append(str(rel))

    return copied


# ══════════════════════════════════════════════════════════════════════════════
# ORPHANED FILES DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class OrphanedFilesDialog(tk.Toplevel):
    """Shows unrecognised root-level items and lets the user delete them."""

    def __init__(self, parent, groups: list[GameGroup], on_done) -> None:
        super().__init__(parent)
        self.title("Clean Orphaned Files")
        self.configure(bg=BG)
        self.grab_set()
        self.resizable(True, True)
        self.geometry("600x460")
        self._on_done = on_done

        orphans = find_orphans(groups)

        ttk.Label(self,
                  text="Items in F:\\RenPy\\ not recognised as a game, archive, or utility:",
                  foreground=FG_DIM, font=("Segoe UI", 9),
                  padding=(12, 10, 12, 4)).pack(anchor="w")

        # Scrollable checkbox list
        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=True, padx=12)

        canvas = tk.Canvas(list_frame, bg=BG2, highlightthickness=0)
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=BG2)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(win_id, width=e.width))
        canvas.bind("<MouseWheel>",
                    lambda e: canvas.yview_scroll(-1 * (e.delta // 120), "units"))

        self._vars: list[tuple[tk.BooleanVar, Path]] = []

        if not orphans:
            tk.Label(inner, text="No orphaned items found — directory is clean.",
                     bg=BG2, fg=GREEN, font=("Segoe UI", 10),
                     pady=20).pack()
        else:
            # Select all / none header
            hdr = tk.Frame(inner, bg=BG2)
            hdr.pack(fill="x", padx=4, pady=(4, 2))
            tk.Button(hdr, text="Select all", bg=BG2, fg=ACCENT, relief="flat",
                      font=("Segoe UI", 8), bd=0, cursor="hand2",
                      command=lambda: [v.set(True) for v, _ in self._vars]
                      ).pack(side="left", padx=(0, 8))
            tk.Button(hdr, text="Select none", bg=BG2, fg=ACCENT, relief="flat",
                      font=("Segoe UI", 8), bd=0, cursor="hand2",
                      command=lambda: [v.set(False) for v, _ in self._vars]
                      ).pack(side="left")

            for path in orphans:
                var = tk.BooleanVar(value=False)
                row = tk.Frame(inner, bg=BG2)
                row.pack(fill="x", padx=4, pady=1)

                cb = tk.Checkbutton(
                    row, variable=var, bg=BG2, fg=FG,
                    selectcolor=SEL, activebackground=BG2,
                    activeforeground=FG)
                cb.pack(side="left")

                # Icon + name
                icon = "📁" if path.is_dir() else "📄"
                tk.Label(row, text=f"{icon} {path.name}",
                         bg=BG2, fg=FG, font=("Segoe UI", 9),
                         anchor="w").pack(side="left", fill="x", expand=True)

                # Size
                try:
                    if path.is_file():
                        sz = path.stat().st_size
                        size_str = (f"{sz / 1_048_576:.1f} MB" if sz > 1_048_576
                                    else f"{sz / 1024:.0f} KB")
                    else:
                        size_str = "folder"
                except OSError:
                    size_str = "?"
                tk.Label(row, text=size_str, bg=BG2, fg=FG_DIM,
                         font=("Segoe UI", 8), width=8,
                         anchor="e").pack(side="right", padx=4)

                self._vars.append((var, path))

        # Buttons
        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=(6, 0))
        btn_row = ttk.Frame(self, padding=(12, 8))
        btn_row.pack(fill="x")
        if orphans:
            ttk.Button(btn_row, text="Delete Selected",
                       style="Danger.TButton",
                       command=self._delete_selected).pack(side="left")
        ttk.Button(btn_row, text="Close",
                   command=self.destroy).pack(side="right")
        self._count_lbl = ttk.Label(btn_row, text="", foreground=FG_DIM,
                                     font=("Segoe UI", 9))
        self._count_lbl.pack(side="right", padx=8)

    def _delete_selected(self) -> None:
        targets = [p for v, p in self._vars if v.get()]
        if not targets:
            messagebox.showinfo("Nothing selected",
                                "Tick items above to delete them.", parent=self)
            return
        names = "\n".join(f"  {p.name}" for p in targets)
        if not messagebox.askyesno(
                "Confirm Delete",
                f"Permanently delete {len(targets)} item(s)?\n\n{names}",
                default="no", parent=self):
            return
        errors: list[str] = []
        deleted = 0
        for p in targets:
            try:
                if p.is_dir():
                    shutil.rmtree(p)
                else:
                    p.unlink()
                deleted += 1
            except OSError as e:
                errors.append(f"{p.name}: {e}")
        if errors:
            messagebox.showerror("Errors", "\n".join(errors), parent=self)
        self._count_lbl.configure(text=f"Deleted {deleted} item(s)")
        self._on_done()
        self.destroy()


# ══════════════════════════════════════════════════════════════════════════════
# PATCH APPLY DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class PatchApplyDialog(tk.Toplevel):
    """Preview and apply patch files to a target game folder."""

    def __init__(self, parent, source_dir: Path,
                 target_version: GameVersion, on_done) -> None:
        super().__init__(parent)
        self.title("Apply Patch")
        self.configure(bg=BG)
        self.grab_set()
        self.resizable(True, True)
        self.geometry("540x460")
        self._source = source_dir
        self._target = target_version
        self._on_done = on_done

        ttk.Label(self,
                  text=f"Target game:  {target_version.folder_name}",
                  font=("Segoe UI", 10, "bold"),
                  padding=(12, 10, 12, 2)).pack(anchor="w")
        ttk.Label(self,
                  text=f"Patch source:  {source_dir}",
                  foreground=FG_DIM, font=("Segoe UI", 8),
                  padding=(12, 0, 12, 6)).pack(anchor="w")

        ttk.Label(self, text="Files that will be copied:",
                  foreground=FG_DIM, font=("Segoe UI", 9),
                  padding=(12, 0)).pack(anchor="w")

        # File preview list
        lf = ttk.Frame(self)
        lf.pack(fill="both", expand=True, padx=12, pady=4)
        lb = tk.Listbox(lf, bg=BG2, fg=FG, selectbackground=SEL,
                        font=("Segoe UI", 9), relief="flat",
                        activestyle="none")
        vsb = ttk.Scrollbar(lf, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        lb.pack(fill="both", expand=True)

        # Build preview
        self._preview: list[tuple[Path, Path]] = []  # (src, dest)
        game_root = target_version.folder_path
        source_game_sub = source_dir / "game"
        if source_game_sub.is_dir():
            for item in sorted(source_game_sub.rglob("*")):
                if item.is_file():
                    rel = item.relative_to(source_game_sub)
                    dest = game_root / "game" / rel
                    self._preview.append((item, dest))
            for item in sorted(source_dir.iterdir()):
                if item.is_file():
                    self._preview.append((item, game_root / item.name))
        else:
            for item in sorted(source_dir.rglob("*")):
                if item.is_file():
                    rel = item.relative_to(source_dir)
                    self._preview.append((item, game_root / rel))

        overwrites = 0
        for src, dest in self._preview:
            exists = dest.exists()
            if exists:
                overwrites += 1
            rel_dest = dest.relative_to(game_root)
            flag = "  ⚠ overwrite" if exists else ""
            lb.insert("end", f"  {rel_dest}{flag}")
            if exists:
                lb.itemconfigure(lb.size() - 1, foreground=YELLOW)

        # Warning
        if overwrites:
            ttk.Label(self,
                      text=f"⚠  {overwrites} file(s) will be overwritten.",
                      foreground=YELLOW, font=("Segoe UI", 9),
                      padding=(12, 0)).pack(anchor="w")

        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=(6, 0))
        btn_row = ttk.Frame(self, padding=(12, 8))
        btn_row.pack(fill="x")
        ttk.Button(btn_row, text=f"Apply ({len(self._preview)} files)",
                   style="Accent.TButton",
                   command=self._apply).pack(side="left")
        ttk.Button(btn_row, text="Cancel",
                   command=self.destroy).pack(side="right")
        self._status = ttk.Label(btn_row, text="", foreground=FG_DIM,
                                  font=("Segoe UI", 9))
        self._status.pack(side="right", padx=8)

    def _apply(self) -> None:
        try:
            copied = apply_patch_from_folder(self._source, self._target)
            self._status.configure(
                text=f"Applied {len(copied)} file(s)", foreground=GREEN)
            messagebox.showinfo(
                "Patch Applied",
                f"Successfully copied {len(copied)} file(s) to:\n"
                f"{self._target.folder_path}",
                parent=self)
            self._on_done()
            self.destroy()
        except OSError as e:
            messagebox.showerror("Error", str(e), parent=self)


# ══════════════════════════════════════════════════════════════════════════════
# ARCHIVES TAB
# ══════════════════════════════════════════════════════════════════════════════

class ArchivesTab(ttk.Frame):
    def __init__(self, parent, app: "LibraryApp", **kw) -> None:
        super().__init__(parent, **kw)
        self._app = app
        self._build()

    def _build(self) -> None:
        # Toolbar
        tb = ttk.Frame(self, padding=(8, 6))
        tb.pack(fill="x")
        ttk.Label(tb, text="ZIP / RAR archives — extract to add games to your library",
                  foreground=FG_DIM, font=("Segoe UI", 9)).pack(side="left")
        ttk.Button(tb, text="⟳ Refresh",
                   command=self._app.refresh).pack(side="right", padx=(4, 0))
        self._btn_del_extracted = ttk.Button(
            tb, text="Delete Extracted (0)",
            style="Danger.TButton",
            command=self._delete_extracted, state="disabled")
        self._btn_del_extracted.pack(side="right", padx=(0, 4))

        ttk.Separator(self, orient="horizontal").pack(fill="x")

        cols = ("size", "type", "status")
        self._tree = ttk.Treeview(
            self, columns=cols, show="tree headings", selectmode="browse")
        self._tree.heading("#0",     text="Name",   anchor="w")
        self._tree.heading("size",   text="Size",   anchor="center")
        self._tree.heading("type",   text="Type",   anchor="center")
        self._tree.heading("status", text="Status / Patch",  anchor="w")
        self._tree.column("#0",     width=340, minwidth=200, stretch=True)
        self._tree.column("size",   width=80,  anchor="center")
        self._tree.column("type",   width=60,  anchor="center")
        self._tree.column("status", width=200, anchor="w")

        self._tree.tag_configure("zip",       foreground=YELLOW)
        self._tree.tag_configure("rar",       foreground=MAUVE)
        self._tree.tag_configure("extracted", foreground=GREEN)
        self._tree.tag_configure("patch",     foreground="#89dceb")

        vsb = ttk.Scrollbar(self, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._tree.pack(fill="both", expand=True)
        self._tree.bind("<<TreeviewSelect>>", self._on_select)

        # Action bar
        act = ttk.Frame(self, padding=(8, 6))
        act.pack(fill="x")
        self._btn_extract = ttk.Button(
            act, text="Extract", style="Accent.TButton",
            command=self._extract, state="disabled")
        self._btn_extract.pack(side="left", padx=(0, 6))
        self._btn_delete = ttk.Button(
            act, text="Delete Archive",
            style="Danger.TButton",
            command=self._delete, state="disabled")
        self._btn_delete.pack(side="left", padx=(0, 6))
        self._btn_open_arc = ttk.Button(
            act, text="Open Folder",
            command=self._open_folder, state="disabled")
        self._btn_open_arc.pack(side="left", padx=(0, 6))

        ttk.Separator(act, orient="vertical").pack(side="left", fill="y", padx=8)

        self._btn_assign_patch = ttk.Button(
            act, text="Assign as Patch for...",
            command=self._assign_patch, state="disabled")
        self._btn_assign_patch.pack(side="left", padx=(0, 6))
        self._btn_apply_patch = ttk.Button(
            act, text="Apply Patch",
            style="Accent.TButton",
            command=self._apply_patch, state="disabled")
        self._btn_apply_patch.pack(side="left", padx=(0, 6))
        self._btn_clear_patch = ttk.Button(
            act, text="Clear Patch Assignment",
            command=self._clear_patch, state="disabled")
        self._btn_clear_patch.pack(side="left")

        self._info_lbl = ttk.Label(
            act, text="", foreground=FG_DIM, font=("Segoe UI", 9))
        self._info_lbl.pack(side="right", padx=8)

    def populate(self, groups: list[GameGroup]) -> None:
        ud = self._app.user_data
        self._tree.delete(*self._tree.get_children())
        self._extracted_archives: list[Archive] = []   # track for bulk delete

        for g in groups:
            for a in g.archives:
                size_mb = a.archive_path.stat().st_size / 1_048_576 \
                    if a.archive_path.exists() else 0
                fmt_size = f"{size_mb:.0f} MB" if size_mb < 1024 \
                    else f"{size_mb / 1024:.1f} GB"
                arc_type = a.archive_path.suffix.upper().lstrip(".")

                # Determine status — check patch assignment first
                arc_name = a.archive_path.name
                patch_for_key = ud.patch_assignments.get(arc_name)
                if patch_for_key:
                    patch_game_name = patch_for_key
                    for pg in groups:
                        if pg.base_key == patch_for_key:
                            patch_game_name = pg.display_name
                            break
                    status = f"Patch for: {patch_game_name}"
                    tag = "patch"
                elif a.matched_folder:
                    status = "Already extracted — safe to delete"
                    tag = "extracted"
                    self._extracted_archives.append(a)
                else:
                    status = "Not extracted"
                    tag = arc_type.lower()

                self._tree.insert(
                    "", "end",
                    text=a.archive_path.stem,
                    values=(fmt_size, arc_type, status),
                    tags=(tag,),
                    iid=str(a.archive_path),
                )

        # Update bulk-delete button
        n = len(self._extracted_archives)
        if n:
            total_bytes = sum(
                a.archive_path.stat().st_size
                for a in self._extracted_archives
                if a.archive_path.exists())
            self._btn_del_extracted.configure(
                text=f"Delete Extracted ({n})  —  {_fmt_bytes(total_bytes)}",
                state="normal")
        else:
            self._btn_del_extracted.configure(
                text="Delete Extracted (0)", state="disabled")

    def _delete_extracted(self) -> None:
        """Bulk-delete all archives whose game folder is already present."""
        candidates = getattr(self, "_extracted_archives", [])
        if not candidates:
            return

        qw = self._app._queue_window
        # Filter out anything currently in the extraction queue
        safe     = [a for a in candidates
                    if not (qw and qw.is_active(a.archive_path))]
        skipped  = len(candidates) - len(safe)

        if not safe:
            messagebox.showinfo(
                "Nothing to Delete",
                "All extracted archives are currently being processed\n"
                "in the Extraction Queue.",
                parent=self)
            return

        total_bytes = sum(
            a.archive_path.stat().st_size
            for a in safe if a.archive_path.exists())

        # Confirmation dialog with scrollable checklist
        dlg = tk.Toplevel(self)
        dlg.title("Delete Extracted Archives")
        dlg.configure(bg=BG)
        dlg.geometry("520x420")
        dlg.resizable(False, True)
        dlg.transient(self)
        dlg.grab_set()

        ttk.Label(dlg,
                  text="These archives already have an extracted game folder.\n"
                       "Deleting them frees space without losing anything.",
                  foreground=FG_DIM, font=("Segoe UI", 9),
                  padding=(14, 12, 14, 4)).pack(anchor="w")

        size_lbl = ttk.Label(
            dlg,
            text=f"Selected: {len(safe)} archives  —  {_fmt_bytes(total_bytes)}",
            font=("Segoe UI", 10, "bold"), foreground=ACCENT,
            padding=(14, 0, 14, 8))
        size_lbl.pack(anchor="w")

        if skipped:
            ttk.Label(dlg,
                      text=f"({skipped} skipped — currently extracting)",
                      foreground=YELLOW, font=("Segoe UI", 8),
                      padding=(14, 0)).pack(anchor="w")

        ttk.Separator(dlg, orient="horizontal").pack(fill="x")

        # Scrollable checklist
        c_frame = ttk.Frame(dlg)
        c_frame.pack(fill="both", expand=True, padx=14, pady=6)
        canvas = tk.Canvas(c_frame, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(c_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        inner = tk.Frame(canvas, bg=BG)
        wid = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(wid, width=e.width))

        bvars: list[tuple[tk.BooleanVar, Archive]] = []
        for a in safe:
            sz = a.archive_path.stat().st_size if a.archive_path.exists() else 0
            bv = tk.BooleanVar(value=True)
            row = tk.Frame(inner, bg=BG)
            row.pack(fill="x", pady=1)
            tk.Checkbutton(
                row, text=a.archive_path.name,
                variable=bv, bg=BG, fg=FG, selectcolor=SEL,
                activebackground=BG, activeforeground=FG,
                font=("Segoe UI", 9),
                command=lambda: _refresh_size()
            ).pack(side="left")
            tk.Label(row, text=_fmt_bytes(sz), bg=BG, fg=FG_DIM,
                     font=("Segoe UI", 8)).pack(side="right", padx=8)
            bvars.append((bv, a))

        def _refresh_size():
            sel = [a for bv, a in bvars if bv.get()]
            total = sum(a.archive_path.stat().st_size
                        for a in sel if a.archive_path.exists())
            size_lbl.configure(
                text=f"Selected: {len(sel)} archives  —  {_fmt_bytes(total)}")

        # Select all / none
        ctrl = tk.Frame(dlg, bg=BG)
        ctrl.pack(fill="x", padx=14, pady=(2, 0))
        tk.Button(ctrl, text="Select all", bg=BG, fg=ACCENT, relief="flat",
                  font=("Segoe UI", 8), cursor="hand2", bd=0,
                  command=lambda: [bv.set(True) for bv, _ in bvars] or _refresh_size()
                  ).pack(side="left")
        tk.Button(ctrl, text="Select none", bg=BG, fg=FG_DIM, relief="flat",
                  font=("Segoe UI", 8), cursor="hand2", bd=0,
                  command=lambda: [bv.set(False) for bv, _ in bvars] or _refresh_size()
                  ).pack(side="left", padx=8)

        ttk.Separator(dlg, orient="horizontal").pack(fill="x", pady=(6, 0))

        bf = ttk.Frame(dlg, padding=(14, 8))
        bf.pack(fill="x")

        def _do_delete():
            errors: list[str] = []
            for bv, a in bvars:
                if bv.get() and a.archive_path.exists():
                    try:
                        a.archive_path.unlink()
                    except OSError as e:
                        errors.append(f"{a.archive_path.name}: {e}")
            dlg.destroy()
            if errors:
                messagebox.showerror("Some Deletes Failed",
                                     "\n".join(errors), parent=self)
            self._app.refresh()

        ttk.Button(bf, text="Delete Selected",
                   style="Danger.TButton",
                   command=_do_delete).pack(side="left", padx=(0, 8))
        ttk.Button(bf, text="Cancel",
                   command=dlg.destroy).pack(side="left")

    def _on_select(self, _e=None) -> None:
        iid = self._tree.focus()
        has = bool(iid)
        self._btn_extract.configure(state="normal" if has else "disabled")
        self._btn_delete.configure(state="normal" if has else "disabled")
        self._btn_open_arc.configure(state="normal" if has else "disabled")
        self._btn_assign_patch.configure(state="normal" if has else "disabled")

        # Apply Patch and Clear only enabled when archive has a patch assignment
        if iid:
            p = Path(iid)
            self._info_lbl.configure(text=str(p))
            ud = self._app.user_data
            has_assignment = p.name in ud.patch_assignments
            self._btn_apply_patch.configure(
                state="normal" if has_assignment else "disabled")
            self._btn_clear_patch.configure(
                state="normal" if has_assignment else "disabled")
        else:
            self._btn_apply_patch.configure(state="disabled")
            self._btn_clear_patch.configure(state="disabled")

    def _selected_archive(self) -> Archive | None:
        iid = self._tree.focus()
        if not iid:
            return None
        target = Path(iid)
        for g in self._app.groups:
            for a in g.archives:
                if a.archive_path == target:
                    return a
        return None

    def _extract(self) -> None:
        a = self._selected_archive()
        if not a:
            return
        if a.archive_path.suffix.lower() == ".rar":
            messagebox.showinfo(
                "RAR Archive",
                "RAR files require 7-Zip to extract.\n\n"
                "1. Extract the RAR manually to F:\\RenPy\\\n"
                "2. Click Refresh to add the game to your library",
                parent=self)
            return
        self._app.queue_extraction(a)

    def _delete(self) -> None:
        a = self._selected_archive()
        if not a:
            return
        qw = self._app._queue_window
        if qw and qw.is_active(a.archive_path):
            messagebox.showwarning(
                "Cannot Delete",
                f"{a.archive_path.name}\n\n"
                "This archive is currently queued or being extracted.\n"
                "Cancel it in the Extraction Queue first.",
                parent=self)
            return
        if not messagebox.askyesno(
                "Confirm Delete",
                f"Permanently delete:\n{a.archive_path.name}?",
                default="no", parent=self):
            return
        try:
            a.archive_path.unlink()
        except OSError as e:
            messagebox.showerror("Error", str(e), parent=self)
            return
        self._app.refresh()

    def _open_folder(self) -> None:
        a = self._selected_archive()
        if a:
            os.startfile(str(a.archive_path.parent))

    # ── Patch management ─────────────────────────────────────────────────────

    def _assign_patch(self) -> None:
        """Open a dialog to assign this archive as a patch for a specific game."""
        a = self._selected_archive()
        if not a:
            return
        groups_with_versions = [g for g in self._app.groups if g.versions]
        if not groups_with_versions:
            messagebox.showinfo("No Games", "No extracted games found.", parent=self)
            return

        # Auto-guess a match to pre-select
        guess = _guess_patch_game(a.archive_path, groups_with_versions)

        dlg = tk.Toplevel(self)
        dlg.title("Assign Patch to Game")
        dlg.configure(bg=BG)
        dlg.geometry("460x340")
        dlg.resizable(False, False)
        dlg.transient(self)
        dlg.grab_set()

        ttk.Label(dlg, text=f"Archive:  {a.archive_path.name}",
                  font=("Segoe UI", 9, "bold")).pack(padx=16, pady=(16, 4), anchor="w")
        ttk.Label(dlg, text="Select the game this patch belongs to:",
                  foreground=FG_DIM).pack(padx=16, anchor="w")

        # Search box
        search_var = tk.StringVar()
        ttk.Entry(dlg, textvariable=search_var, width=40).pack(
            padx=16, pady=(8, 4), anchor="w")

        # Listbox with scrollbar
        lf = ttk.Frame(dlg)
        lf.pack(fill="both", expand=True, padx=16, pady=4)
        lb = tk.Listbox(lf, bg=BG2, fg=FG, selectbackground=SEL,
                        selectforeground=ACCENT, activestyle="none",
                        relief="flat", bd=0, font=("Segoe UI", 9))
        lsb = ttk.Scrollbar(lf, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=lsb.set)
        lsb.pack(side="right", fill="y")
        lb.pack(side="left", fill="both", expand=True)

        # Populate listbox
        all_names = [(g.display_name, g.base_key) for g in groups_with_versions]
        all_names.sort(key=lambda x: x[0].lower())

        def _refresh_list(*_):
            q = search_var.get().lower()
            lb.delete(0, "end")
            for dname, _ in all_names:
                if not q or q in dname.lower():
                    lb.insert("end", dname)
            # Try to select guess or first item
            if lb.size() > 0:
                lb.selection_set(0)
                lb.see(0)

        search_var.trace_add("write", _refresh_list)
        _refresh_list()

        # Pre-select guess if any
        if guess:
            for i in range(lb.size()):
                if lb.get(i) == guess.display_name:
                    lb.selection_clear(0, "end")
                    lb.selection_set(i)
                    lb.see(i)
                    break

        def _confirm():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("No Selection",
                                       "Please select a game.", parent=dlg)
                return
            chosen_name = lb.get(sel[0])
            chosen_key = next(
                (k for n, k in all_names if n == chosen_name), None)
            if chosen_key:
                ud = self._app.user_data
                ud.patch_assignments[a.archive_path.name] = chosen_key
                save_userdata(ud)
            dlg.destroy()
            self.populate(self._app.groups)
            # Re-select the same item
            iid = str(a.archive_path)
            if self._tree.exists(iid):
                self._tree.selection_set(iid)
                self._tree.focus(iid)
                self._on_select()

        bf = ttk.Frame(dlg)
        bf.pack(fill="x", padx=16, pady=(4, 16))
        ttk.Button(bf, text="Assign", style="Accent.TButton",
                   command=_confirm).pack(side="left", padx=(0, 8))
        ttk.Button(bf, text="Cancel",
                   command=dlg.destroy).pack(side="left")
        lb.bind("<Double-1>", lambda _: _confirm())

    def _clear_patch(self) -> None:
        """Remove the patch assignment for the selected archive."""
        a = self._selected_archive()
        if not a:
            return
        ud = self._app.user_data
        ud.patch_assignments.pop(a.archive_path.name, None)
        save_userdata(ud)
        self.populate(self._app.groups)
        iid = str(a.archive_path)
        if self._tree.exists(iid):
            self._tree.selection_set(iid)
            self._tree.focus(iid)
            self._on_select()

    def _apply_patch(self) -> None:
        """Extract/locate patch files then open PatchApplyDialog."""
        a = self._selected_archive()
        if not a:
            return
        ud = self._app.user_data
        patch_key = ud.patch_assignments.get(a.archive_path.name)
        if not patch_key:
            messagebox.showinfo(
                "No Assignment",
                "This archive has not been assigned to a game yet.\n"
                "Use 'Assign as Patch for...' first.",
                parent=self)
            return

        # Find the target game group and pick the newest version
        target_group = next(
            (g for g in self._app.groups if g.base_key == patch_key and g.versions),
            None)
        if not target_group:
            messagebox.showerror(
                "Game Not Found",
                "The assigned game could not be found. Re-assign and try again.",
                parent=self)
            return
        target_version = target_group.versions[-1]  # newest

        suffix = a.archive_path.suffix.lower()
        if suffix == ".zip":
            # Extract to temp dir, then open preview dialog
            try:
                tmp = Path(tempfile.mkdtemp(prefix="renpy_patch_"))
                with zipfile.ZipFile(a.archive_path) as zf:
                    zf.extractall(tmp)
            except Exception as exc:
                messagebox.showerror("Extraction Error", str(exc), parent=self)
                return

            def _cleanup_and_refresh():
                try:
                    shutil.rmtree(tmp, ignore_errors=True)
                except Exception:
                    pass
                self._app.refresh()

            PatchApplyDialog(self, tmp, target_version,
                             on_done=_cleanup_and_refresh)

        elif suffix == ".rar":
            # Can't extract RAR — ask user to browse to the extracted folder
            messagebox.showinfo(
                "RAR Patch",
                "RAR files cannot be extracted automatically.\n\n"
                "Please:\n"
                "1. Extract the RAR file manually with 7-Zip\n"
                "2. Click OK, then browse to the extracted folder",
                parent=self)
            folder = filedialog.askdirectory(
                title="Select Extracted Patch Folder",
                initialdir=str(RENPY_DIR),
                parent=self)
            if not folder:
                return
            PatchApplyDialog(self, Path(folder), target_version,
                             on_done=self._app.refresh)
        else:
            messagebox.showinfo(
                "Unsupported Format",
                f"Cannot handle {suffix} archives for patching.",
                parent=self)


# ══════════════════════════════════════════════════════════════════════════════
# SITE LOGIN  (form-based, no browser dependency)
# ══════════════════════════════════════════════════════════════════════════════

def _form_login(site_key: str, username: str, password: str) -> str:
    """
    Log in to a site using curl_cffi form POST.
    Returns "" on success, or an error message string on failure.
    Saves cookies to disk on success.
    """
    if not HAS_SCRAPING:
        return "curl_cffi not installed."

    def _resp_cookies(resp) -> dict:
        """Extract cookies from a response object — works around curl_cffi
        Session jar inconsistencies by reading the raw response cookies."""
        try:
            return dict(resp.cookies)
        except Exception:
            return {}

    def _merge(*cookie_dicts) -> dict:
        out: dict = {}
        for d in cookie_dicts:
            out.update(d)
        return out

    try:
        session = _req.Session(impersonate="chrome131")

        if site_key == "itchio":
            login_url = f"{ITCHIO_BASE}/login"
            # Warmup: hit the homepage first so Cloudflare issues cf_clearance
            # before we touch the login page (itch.io CF is stricter than others)
            session.get(ITCHIO_BASE, timeout=SCRAPER_TIMEOUT)
            # Step 1: GET login page — session now carries cf_clearance
            r_get = session.get(login_url, timeout=SCRAPER_TIMEOUT)
            get_cookies = _merge(dict(session.cookies), _resp_cookies(r_get))

            # CSRF token = the itchio_token cookie set by the GET response
            csrf = (get_cookies.get("itchio_token", "")
                    or get_cookies.get("csrf_token", ""))
            # Also check HTML form / meta / inline JS as fallbacks
            if not csrf:
                soup = _BS(r_get.text, "lxml")
                el = soup.select_one('input[name="csrf_token"]')
                if el:
                    csrf = el.get("value", "")
            if not csrf:
                m = re.search(r'"csrf_token"\s*[=:]\s*"([^"]{8,})"', r_get.text)
                if m:
                    csrf = m.group(1)

            # tz = local timezone offset in minutes (matches what the browser sends)
            import time as _time
            tz_minutes = -int(_time.timezone / 60)

            # Step 2: POST with all GET-phase cookies so cf_clearance is sent
            r_post = session.post(
                login_url, timeout=SCRAPER_TIMEOUT,
                cookies=get_cookies,
                headers={"Referer": login_url},
                data={
                    "username":   username,
                    "password":   password,
                    "csrf_token": csrf,
                    "tz":         str(tz_minutes),
                    "source":     "login_page",
                })
            post_cookies = _resp_cookies(r_post)
            all_cookies  = _merge(get_cookies, post_cookies, dict(session.cookies))

            # Success = redirected away from /login (302 → /my-feed)
            final_url = getattr(r_post, "url", "")
            if final_url.rstrip("/").endswith("/login"):
                err_el = _BS(r_post.text, "lxml").select_one(
                    ".form_errors li, .error_list li, p.error, .notice--error")
                if err_el:
                    return f"Login failed — {err_el.get_text(strip=True)}"
                return (f"Login failed (csrf={'yes' if csrf else 'no'}; "
                        f"cookies: {list(all_cookies.keys()) or 'none'}; "
                        f"url: {final_url[:80]})")
            save_site_cookies(site_key, all_cookies)
            return ""

        elif site_key == "f95zone":
            r_get = session.get(f"{F95_BASE}/login", timeout=SCRAPER_TIMEOUT)
            get_cookies = _merge(dict(session.cookies), _resp_cookies(r_get))
            soup = _BS(r_get.text, "lxml")
            tok_el = soup.select_one('input[name="_xfToken"]')
            xf_token = tok_el.get("value", "") if tok_el else ""
            r_post = session.post(
                f"{F95_BASE}/login/login", timeout=SCRAPER_TIMEOUT,
                cookies=get_cookies,
                data={
                    "login":    username,
                    "password": password,
                    "_xfToken": xf_token,
                    "remember": "1",
                })
            post_cookies = _resp_cookies(r_post)
            all_cookies  = _merge(get_cookies, post_cookies, dict(session.cookies))
            if "xf_user" not in all_cookies:
                return "Login failed — check your username and password."
            save_site_cookies(site_key, all_cookies)
            return ""

        elif site_key == "lewdcorner":
            # LewdCorner runs XenForo — identical flow to F95Zone
            r_get = session.get(f"{LC_BASE}/login", timeout=SCRAPER_TIMEOUT)
            get_cookies = _merge(dict(session.cookies), _resp_cookies(r_get))
            soup = _BS(r_get.text, "lxml")
            tok_el = soup.select_one('input[name="_xfToken"]')
            xf_token = tok_el.get("value", "") if tok_el else ""
            r_post = session.post(
                f"{LC_BASE}/login/login", timeout=SCRAPER_TIMEOUT,
                cookies=get_cookies,
                headers={"Referer": f"{LC_BASE}/"},
                data={
                    "login":       username,
                    "password":    password,
                    "_xfToken":    xf_token,
                    "remember":    "1",
                    "_xfRedirect": f"{LC_BASE}/",
                })
            post_cookies = _resp_cookies(r_post)
            all_cookies  = _merge(get_cookies, post_cookies, dict(session.cookies))
            if "xf_user" not in all_cookies:
                return "Login failed — check your username and password."
            save_site_cookies(site_key, all_cookies)
            return ""

        else:
            return f"Unknown site: {site_key}"

    except Exception as exc:
        return str(exc)


# ── pywebview subprocess worker ─────────────────────────────────────────────
# Must be a module-level function so multiprocessing can pickle it on Windows.

def _webview_worker(url: str, success_url_fragment: str,
                    out_queue: multiprocessing.Queue) -> None:
    """
    Run in a subprocess: open an embedded Chromium window at *url*.
    Once the current URL contains *success_url_fragment*, scrape cookies
    and put ("ok", {name: value, ...}) on the queue.
    Cancelled / closed → ("cancel", {}).
    Error → ("error", message).
    """
    try:
        import webview
        import threading as _thr

        found: dict = {}
        win = webview.create_window(
            "Log in",
            url,
            width=980,
            height=720,
            on_top=True,
        )

        def _poll_loop() -> None:
            import time as _t
            while True:
                _t.sleep(0.5)
                try:
                    cur = win.get_current_url() or ""
                    if success_url_fragment in cur:
                        # get_cookies() returns a list of SimpleCookie objects
                        # (one SimpleCookie per cookie).  Iterate .items() to get
                        # (name, Morsel) pairs; Morsel.value is the cookie value.
                        for simple_cookie in (win.get_cookies() or []):
                            try:
                                for name, morsel in simple_cookie.items():
                                    if name:
                                        found[name] = morsel.value
                            except Exception:
                                # Fallback: treat as plain dict / object
                                try:
                                    n = getattr(simple_cookie, "name",  None) or simple_cookie.get("name",  "")
                                    v = getattr(simple_cookie, "value", None) or simple_cookie.get("value", "")
                                    if n:
                                        found[n] = v
                                except Exception:
                                    pass
                        out_queue.put(("ok", found))
                        win.destroy()
                        return
                except Exception:
                    pass  # window not ready yet — keep polling

        def _on_shown() -> None:
            _thr.Thread(target=_poll_loop, daemon=True).start()

        win.events.shown += _on_shown
        webview.start(gui="edgechromium", debug=False)

        # If we reach here without found being populated the user closed the window
        if not found:
            out_queue.put(("cancel", {}))

    except Exception as exc:
        out_queue.put(("error", str(exc)))


class SiteCookiesDialog(tk.Toplevel):
    """
    Login dialog for metadata scraping sites.
    - itch.io  → embedded Chromium browser via pywebview (handles Cloudflare)
    - F95Zone  → username + password form
    - LewdCorner → username + password form
    """

    # (site_key, display_label, login_mode)
    # login_mode: "browser" | "form"
    _SITES = [
        ("itchio",     "itch.io",     "browser"),
        ("f95zone",    "F95Zone",     "form"),
        ("lewdcorner", "Lewd Corner", "form"),
    ]

    def __init__(self, parent) -> None:
        super().__init__(parent)
        self.title("Site Login")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(parent)

        # Keep a reference so the GC doesn't collect it mid-poll
        self._webview_proc: multiprocessing.Process | None = None
        self._webview_queue: multiprocessing.Queue | None = None

        tk.Label(self, bg=BG, fg=FG,
                 text="Log in to enable metadata scraping.",
                 font=("Segoe UI", 10, "bold")).pack(padx=20, pady=(16, 4))
        tk.Label(self, bg=BG, fg=FG_DIM,
                 text="Credentials are used once to get a session token.\n"
                      "Your password is never stored.",
                 font=("Segoe UI", 9), justify="center").pack(padx=20, pady=(0, 10))

        for site_key, label, login_mode in self._SITES:
            has_cookies = bool(load_site_cookies(site_key))

            frm = tk.LabelFrame(self, text=label, bg=BG, fg=ACCENT,
                                font=("Segoe UI", 9, "bold"), padx=14, pady=10)
            frm.pack(fill="x", padx=16, pady=(0, 8))

            status_lbl = tk.Label(frm, bg=BG, font=("Segoe UI", 8),
                                  text="✓ Logged in" if has_cookies else "Not logged in",
                                  fg=GREEN if has_cookies else FG_DIM)
            status_lbl.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))

            if login_mode == "browser":
                self._build_browser_section(frm, site_key, status_lbl)
            else:
                self._build_form_section(frm, site_key, status_lbl)

        ttk.Button(self, text="Close", command=self.destroy).pack(pady=14)

        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width()  // 2 - self.winfo_width()  // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2 - self.winfo_height() // 2
        self.geometry(f"+{max(0,px)}+{max(0,py)}")

    # ── section builders ────────────────────────────────────────────────────

    def _build_browser_section(self, frm: tk.LabelFrame,
                               site_key: str, status_lbl: tk.Label) -> None:
        """itch.io: single button that opens an embedded browser window."""
        inner = tk.Frame(frm, bg=BG)
        inner.grid(row=1, column=0, columnspan=2, sticky="ew")

        login_btn = ttk.Button(
            inner, text="Log in with browser", style="Accent.TButton",
            command=lambda: self._browser_login(site_key, status_lbl, login_btn)
        )
        login_btn.pack(side="left", padx=(0, 8))

        if not HAS_WEBVIEW:
            login_btn.configure(state="disabled")
            tk.Label(inner, bg=BG, fg=YELLOW,
                     text="pywebview not installed — run: pip install pywebview",
                     font=("Segoe UI", 8)).pack(side="left")

        ttk.Button(
            inner, text="Log out",
            command=lambda k=site_key, sl=status_lbl: self._logout(k, sl)
        ).pack(side="left")

    def _build_form_section(self, frm: tk.LabelFrame,
                            site_key: str, status_lbl: tk.Label) -> None:
        """F95Zone / LewdCorner: username + password entry fields."""
        tk.Label(frm, bg=BG, fg=FG_DIM, text="Username:",
                 font=("Segoe UI", 9), width=10, anchor="w").grid(
                     row=1, column=0, sticky="w")
        user_var = tk.StringVar()
        ttk.Entry(frm, textvariable=user_var, width=26).grid(
            row=1, column=1, sticky="w", padx=(4, 8))

        tk.Label(frm, bg=BG, fg=FG_DIM, text="Password:",
                 font=("Segoe UI", 9), width=10, anchor="w").grid(
                     row=2, column=0, sticky="w", pady=(4, 0))
        pass_var = tk.StringVar()
        ttk.Entry(frm, textvariable=pass_var, show="●", width=26).grid(
            row=2, column=1, sticky="w", padx=(4, 8), pady=(4, 0))

        btn_frame = tk.Frame(frm, bg=BG)
        btn_frame.grid(row=1, column=2, rowspan=2, sticky="e")

        ttk.Button(
            btn_frame, text="Log in",
            style="Accent.TButton",
            command=lambda k=site_key, u=user_var, p=pass_var,
                           sl=status_lbl: self._login(k, u, p, sl)
        ).pack(pady=(0, 4))
        ttk.Button(
            btn_frame, text="Log out",
            command=lambda k=site_key, sl=status_lbl: self._logout(k, sl)
        ).pack()

    # ── browser login (itch.io) ─────────────────────────────────────────────

    def _browser_login(self, site_key: str, status_lbl: tk.Label,
                       login_btn: ttk.Button) -> None:
        """Spawn a pywebview subprocess and poll for the success cookie."""
        if not HAS_WEBVIEW:
            return

        login_btn.configure(state="disabled")
        status_lbl.configure(text="Browser opening…", fg=YELLOW)
        self.update()

        q: multiprocessing.Queue = multiprocessing.Queue()
        p = multiprocessing.Process(
            target=_webview_worker,
            args=(f"{ITCHIO_BASE}/login", "/my-feed", q),
            daemon=True,
        )
        p.start()

        self._webview_proc  = p
        self._webview_queue = q

        self._poll_webview(site_key, status_lbl, login_btn, q, p)

    def _poll_webview(self, site_key: str, status_lbl: tk.Label,
                      login_btn: ttk.Button,
                      q: multiprocessing.Queue,
                      p: multiprocessing.Process) -> None:
        """Called every 500 ms from the tkinter event loop to check queue."""
        # Dialog may have been closed by the user while the browser was open
        if not self.winfo_exists():
            if p.is_alive():
                p.terminate()
            return

        try:
            kind, payload = q.get_nowait()
        except Exception:
            # Nothing yet — is the process still alive?
            if p.is_alive():
                self.after(500, lambda: self._poll_webview(
                    site_key, status_lbl, login_btn, q, p))
            else:
                # Process died unexpectedly
                status_lbl.configure(text="✗ Browser closed unexpectedly", fg=RED)
                login_btn.configure(state="normal")
            return

        p.join(timeout=2)

        if kind == "ok":
            cookies: dict = payload
            if cookies:
                save_site_cookies(site_key, cookies)
                status_lbl.configure(text="✓ Logged in", fg=GREEN)
            else:
                status_lbl.configure(text="✗ No cookies captured — try again", fg=RED)
        elif kind == "cancel":
            status_lbl.configure(text="Login cancelled", fg=FG_DIM)
        else:
            status_lbl.configure(text=f"✗ {payload}", fg=RED)

        login_btn.configure(state="normal")

    # ── form login (F95 / LC) ───────────────────────────────────────────────

    def _login(self, site_key: str, user_var: tk.StringVar,
               pass_var: tk.StringVar, status_lbl: tk.Label) -> None:
        u = user_var.get().strip()
        p = pass_var.get()
        if not u or not p:
            messagebox.showwarning("Missing fields",
                                   "Enter both username and password.",
                                   parent=self)
            return
        status_lbl.configure(text="Logging in…", fg=YELLOW)
        self.update()

        def _bg():
            err = _form_login(site_key, u, p)
            if err:
                status_lbl.after(0, lambda: status_lbl.configure(
                    text=f"✗ {err}", fg=RED))
            else:
                pass_var.set("")   # clear password from UI
                status_lbl.after(0, lambda: status_lbl.configure(
                    text="✓ Logged in", fg=GREEN))

        threading.Thread(target=_bg, daemon=True).start()

    def _logout(self, site_key: str, status_lbl: tk.Label) -> None:
        save_site_cookies(site_key, {})
        status_lbl.configure(text="Not logged in", fg=FG_DIM)


# ══════════════════════════════════════════════════════════════════════════════
# SETTINGS DIALOG
# ══════════════════════════════════════════════════════════════════════════════

class SettingsDialog(tk.Toplevel):
    """App-wide settings window."""

    def __init__(self, app: "LibraryApp") -> None:
        super().__init__(app)
        self.app = app
        self.title("Settings")
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(app)
        self.grab_set()
        self._build()
        self.update_idletasks()
        x = app.winfo_x() + (app.winfo_width()  - self.winfo_width())  // 2
        y = app.winfo_y() + (app.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build(self) -> None:
        s = self.app._settings

        # ── General section ───────────────────────────────────────────────────
        hdr = tk.Frame(self, bg=BG2)
        hdr.pack(fill="x")
        tk.Label(hdr, text="  General", bg=BG2, fg=ACCENT,
                 font=("Segoe UI", 10, "bold"), pady=6).pack(side="left")
        ttk.Separator(self, orient="horizontal").pack(fill="x")

        row = tk.Frame(self, bg=BG)
        row.pack(fill="x", padx=20, pady=(10, 4))
        tk.Label(row, text="Game library directory", bg=BG, fg=FG,
                 font=("Segoe UI", 9)).pack(side="left")

        dir_row = tk.Frame(self, bg=BG)
        dir_row.pack(fill="x", padx=20, pady=(0, 10))
        self._lib_var = tk.StringVar(value=s.get("library_dir", str(RENPY_DIR)))
        ttk.Entry(dir_row, textvariable=self._lib_var, width=46).pack(
            side="left", fill="x", expand=True)
        ttk.Button(dir_row, text="Browse…",
                   command=self._browse_lib).pack(side="left", padx=(4, 0))
        ttk.Button(dir_row, text="Apply",
                   command=self._apply_lib_dir).pack(side="left", padx=(4, 0))

        # Slideshow interval
        slide_row = tk.Frame(self, bg=BG)
        slide_row.pack(fill="x", padx=20, pady=(4, 10))
        tk.Label(slide_row, text="Slideshow interval", bg=BG, fg=FG,
                 font=("Segoe UI", 9)).pack(side="left")
        self._slide_val_lbl = tk.Label(slide_row, bg=BG, fg=ACCENT,
                                       font=("Segoe UI", 9, "bold"), width=5)
        self._slide_val_lbl.pack(side="right")
        tk.Label(slide_row, text="s", bg=BG, fg=FG_DIM,
                 font=("Segoe UI", 9)).pack(side="right")
        saved_interval = float(s.get("slideshow_interval", 3.5))
        # Store as int*10 internally (5 = 0.5s, 300 = 30s) for Scale resolution
        self._slide_var = tk.IntVar(value=int(round(saved_interval * 10)))
        slide_scale = ttk.Scale(
            slide_row, from_=5, to=300,
            orient="horizontal", variable=self._slide_var,
            command=self._on_slide_change)
        slide_scale.pack(side="right", fill="x", expand=True, padx=(10, 6))
        self._on_slide_change()   # set initial label

        # ── Connection section ────────────────────────────────────────────────
        ttk.Separator(self, orient="horizontal").pack(fill="x", pady=(4, 0))
        hdr2 = tk.Frame(self, bg=BG2)
        hdr2.pack(fill="x")
        tk.Label(hdr2, text="  Connection", bg=BG2, fg=ACCENT,
                 font=("Segoe UI", 10, "bold"), pady=6).pack(side="left")
        ttk.Separator(self, orient="horizontal").pack(fill="x")

        # LOCKDOWN master toggle
        lock_row = tk.Frame(self, bg=BG)
        lock_row.pack(fill="x", padx=20, pady=(10, 2))
        self._lock_var = tk.BooleanVar(value=bool(s.get("lockdown", True)))
        self._lock_cb = tk.Checkbutton(
            lock_row,
            text="  LOCKDOWN MODE  —  block all internet access",
            variable=self._lock_var,
            bg=BG, fg=RED,
            selectcolor=BG3,
            activebackground=BG, activeforeground=RED,
            font=("Segoe UI", 10, "bold"),
            command=self._on_lockdown_toggle)
        self._lock_cb.pack(side="left")

        lock_note = tk.Label(
            self,
            text="Master override — blocks all connectivity regardless of individual\n"
                 "settings. ON by default. Only this toggle can restore access.",
            bg=BG, fg=FG_DIM, font=("Segoe UI", 8), justify="left")
        lock_note.pack(anchor="w", padx=36, pady=(0, 8))

        ttk.Separator(self, orient="horizontal").pack(
            fill="x", padx=20, pady=(0, 6))

        # Individual net toggles — each in its own row for optional extra widgets
        self._net_toggles: list[tuple[str, tk.BooleanVar, tk.Checkbutton]] = []
        # (key, label, extra_widget_builder_or_None)
        toggle_defs = [
            ("check_updates",        "App update checks"),
            ("fetch_metadata",       "Metadata fetch  (VNDB, F95Zone, LewdCorner)"),
            ("allow_provider_login", "Provider logins  (site cookies / webview)"),
            ("allow_download_links", "Download page links  (↗ Download Page button)"),
        ]
        tgl_frame = tk.Frame(self, bg=BG)
        tgl_frame.pack(fill="x", padx=36, pady=(0, 14))
        self._check_now_btn: ttk.Button | None = None
        for key, label in toggle_defs:
            default = bool(_SETTINGS_DEFAULTS.get(key, True))
            var = tk.BooleanVar(value=bool(s.get(key, default)))
            row = tk.Frame(tgl_frame, bg=BG)
            row.pack(fill="x", pady=2)
            cb = tk.Checkbutton(
                row, text=label,
                variable=var,
                bg=BG, fg=FG,
                selectcolor=BG3,
                activebackground=BG, activeforeground=FG,
                font=("Segoe UI", 9),
                command=lambda k=key, v=var: self._on_net_toggle(k, v))
            cb.pack(side="left")
            if key == "check_updates":
                self._check_now_btn = ttk.Button(
                    row, text="Check Now", command=self._check_now)
                self._check_now_btn.pack(side="left", padx=(10, 0))
            self._net_toggles.append((key, var, cb))

        self._sync_lockdown_ui()

        # ── Close + version footer ─────────────────────────────────────────────
        ttk.Separator(self, orient="horizontal").pack(fill="x")
        foot = tk.Frame(self, bg=BG)
        foot.pack(fill="x", padx=16, pady=(6, 0))
        tk.Label(foot, text=f"VN Pathfinder  v{APP_VERSION}",
                 bg=BG, fg=FG_MUT, font=("Segoe UI", 8)).pack(side="right")
        ttk.Button(self, text="Close",
                   command=self.destroy).pack(pady=(4, 10))

    # ── General handlers ──────────────────────────────────────────────────────

    def _on_slide_change(self, _=None) -> None:
        val = self._slide_var.get()
        seconds = val / 10.0
        # Display as integer if whole number, else one decimal place
        display = f"{int(seconds)}" if seconds == int(seconds) else f"{seconds:.1f}"
        self._slide_val_lbl.configure(text=display)
        self.app._settings["slideshow_interval"] = seconds
        save_settings(self.app._settings)

    def _browse_lib(self) -> None:
        d = filedialog.askdirectory(
            title="Select game library directory",
            initialdir=self._lib_var.get() or str(RENPY_DIR),
            parent=self)
        if d:
            self._lib_var.set(d)

    def _apply_lib_dir(self) -> None:
        p_str = self._lib_var.get().strip()
        if not p_str:
            return
        p = Path(p_str)
        if not p.is_dir():
            messagebox.showerror("Invalid path",
                                 f"Directory not found:\n{p_str}", parent=self)
            return
        self.app._settings["library_dir"] = str(p)
        save_settings(self.app._settings)
        _set_lib_dir(p)
        self.app.refresh()

    # ── Connection handlers ───────────────────────────────────────────────────

    def _on_lockdown_toggle(self) -> None:
        locked = self._lock_var.get()
        self.app._settings["lockdown"] = locked
        if locked:
            for key, var, _ in self._net_toggles:
                var.set(False)
                self.app._settings[key] = False
            self.app._upd_note_lbl.configure(text="")
        save_settings(self.app._settings)
        self._sync_lockdown_ui()
        self.app._update_lockdown_indicator()

    def _on_net_toggle(self, key: str, var: tk.BooleanVar) -> None:
        self.app._settings[key] = var.get()
        save_settings(self.app._settings)

    def _check_now(self) -> None:
        if not self.app.net_ok("check_updates"):
            messagebox.showinfo(
                "Blocked by Lockdown",
                "Update checks are blocked while Lockdown Mode is active.\n\n"
                "Disable Lockdown Mode in Settings → Connection first.",
                parent=self)
            return
        self.app._upd_note_lbl.configure(text="Checking…", fg=FG_DIM)
        self.app._start_update_check()

    def _sync_lockdown_ui(self) -> None:
        locked = self._lock_var.get()
        state = "disabled" if locked else "normal"
        for _, _, cb in self._net_toggles:
            cb.configure(state=state, fg=FG_MUT if locked else FG)
        if self._check_now_btn is not None:
            self._check_now_btn.configure(state=state)


# ══════════════════════════════════════════════════════════════════════════════
# MAIN APPLICATION
# ══════════════════════════════════════════════════════════════════════════════

class LibraryApp(tk.Tk):

    def __init__(self) -> None:
        super().__init__()
        self.title(f"VN Pathfinder  v{APP_VERSION}")
        self.geometry("1200x820")
        self.minsize(860, 600)
        apply_theme(self)

        # Window icon
        _ico = ASSETS_DIR / "logo.ico"
        if _ico.exists():
            try:
                self.iconbitmap(str(_ico))
            except tk.TclError:
                pass

        self._settings: dict = load_settings()
        # Apply saved library directory (overrides APP_DIR.parent default)
        _lib = self._settings.get("library_dir", "")
        if _lib:
            _p = Path(_lib)
            if _p.is_dir():
                _set_lib_dir(_p)
        self.user_data: UserData = load_userdata()
        self.groups: list[GameGroup] = []
        self.thumb_cache: ThumbnailCache = ThumbnailCache(self)
        self.play_tracker: PlayTracker = PlayTracker()
        self._show_hidden: bool = False
        self._filter_status: str = "All"    # All / Played / Unplayed
        self._filter_tags: set[str] = set()
        self._filter_tags_mode: str = "any"  # "any" (OR) or "all" (AND)
        self._queue_window: ExtractionQueueWindow | None = None

        self._build_ui()
        self._update_lockdown_indicator()
        self.refresh()
        # Kick off update check 800 ms after window opens
        if self.net_ok("check_updates"):
            self.after(800, self._start_update_check)

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # ── Top toolbar ──────────────────────────────────────────────────────
        tb = ttk.Frame(self, padding=(12, 8))
        tb.pack(fill="x")
        ttk.Label(tb, text="VN Pathfinder",
                  font=("Segoe UI", 15, "bold"),
                  foreground=ACCENT).pack(side="left")

        ttk.Button(tb, text="⚙ Settings",
                   command=self.open_settings).pack(side="right", padx=4)
        ttk.Button(tb, text="⟳ Refresh",
                   command=self.refresh).pack(side="right", padx=4)
        self._btn_hidden = ttk.Button(
            tb, text="Show Hidden", command=self._toggle_hidden)
        self._btn_hidden.pack(side="right", padx=4)
        ttk.Button(tb, text="🧹 Clean Orphans",
                   command=self._open_orphan_cleaner).pack(side="right", padx=4)

        # Sort
        ttk.Label(tb, text="Sort:").pack(side="right", padx=(16, 2))
        self._sort_var = tk.StringVar(value="Name A–Z")
        sort_cb = ttk.Combobox(
            tb, textvariable=self._sort_var,
            values=["Name A–Z", "Name Z–A",
                    "Most Played", "Recently Played",
                    "Newest Version", "Oldest Version"],
            state="readonly", width=14)
        sort_cb.pack(side="right", padx=2)
        sort_cb.bind("<<ComboboxSelected>>", lambda _: self._apply_filters())

        # Tag filter — [Tags ▾] [✕] as a unit
        _tag_frame = tk.Frame(tb, bg=BG)
        _tag_frame.pack(side="right", padx=(16, 0))
        self._tag_btn = ttk.Menubutton(_tag_frame, text="Tags ▾")
        self._tag_menu = tk.Menu(self._tag_btn, tearoff=False,
                                  bg=BG2, fg=FG, activebackground=SEL,
                                  activeforeground=ACCENT, bd=0)
        self._tag_btn.configure(menu=self._tag_menu)
        self._tag_btn.pack(side="left")
        self._tag_clear_btn = tk.Button(
            _tag_frame, text="✕", bg=BG, fg=FG_MUT,
            relief="flat", font=("Segoe UI", 8, "bold"),
            cursor="hand2", bd=0, padx=4,
            command=self._clear_tags)
        self._tag_clear_btn.pack(side="left", padx=(2, 0))
        self._tag_vars: dict[str, tk.BooleanVar] = {}
        self._tag_mode_var = tk.StringVar(value="any")

        # Search
        ttk.Label(tb, text="Search:").pack(side="right", padx=(16, 2))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._apply_filters())
        ttk.Entry(tb, textvariable=self._search_var, width=22).pack(
            side="right", padx=2)

        ttk.Separator(self, orient="horizontal").pack(fill="x")

        # ── Notebook ─────────────────────────────────────────────────────────
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        # Library tab
        self._lib_frame = ttk.Frame(self.notebook)
        self.notebook.add(self._lib_frame, text="  Library  ")

        # Sub-filter bar inside library tab
        sfbar = tk.Frame(self._lib_frame, bg=BG2)
        sfbar.pack(fill="x")
        self._sf_btns: dict[str, tk.Button] = {}
        for label in ("All", "Played", "Unplayed"):
            b = tk.Button(
                sfbar, text=label, bg=ACCENT if label == "All" else SEL,
                fg=BG if label == "All" else FG_DIM,
                relief="flat", font=("Segoe UI", 9, "bold"),
                padx=14, pady=5, cursor="hand2",
                command=lambda l=label: self._set_status_filter(l))
            b.pack(side="left")
            self._sf_btns[label] = b

        ttk.Separator(self._lib_frame, orient="horizontal").pack(fill="x")

        # Horizontal pane: card list | detail panel
        self._lib_pane = ttk.PanedWindow(self._lib_frame, orient="horizontal")
        self._lib_pane.pack(fill="both", expand=True)

        left = ttk.Frame(self._lib_pane)
        self._lib_pane.add(left, weight=2)
        self.card_list = ScrollableCardList(
            left, on_select_cb=self.select_group)
        self.card_list.pack(fill="both", expand=True)

        right = ttk.Frame(self._lib_pane)
        self._lib_pane.add(right, weight=1)
        self.detail_panel = DetailPanel(right, app=self)
        self.detail_panel.pack(fill="both", expand=True)

        # Restore saved sash position after window is drawn
        default_sash = self._settings.get("lib_sash", 620)
        self._lib_pane.bind(
            "<Map>",
            lambda e, p=self._lib_pane, d=default_sash: self.after(
                50, lambda: self._apply_lib_sash(p, d)))
        self._lib_pane.bind(
            "<ButtonRelease-1>",
            lambda e, p=self._lib_pane: self._persist_lib_sash(p))

        # Archives tab
        self._arc_frame = ttk.Frame(self.notebook)
        self.notebook.add(self._arc_frame, text="  Archives  ")
        self.archives_tab = ArchivesTab(self._arc_frame, app=self)
        self.archives_tab.pack(fill="both", expand=True)

        # ── Status bar ───────────────────────────────────────────────────────
        ttk.Separator(self, orient="horizontal").pack(fill="x")
        sb = tk.Frame(self, bg=BG2)
        sb.pack(fill="x")
        self._status_lbl = ttk.Label(sb, style="Status.TLabel", text="")
        self._status_lbl.pack(side="left", fill="x", expand=True)
        # Lockdown indicator (far right)
        self._lockdown_lbl = tk.Label(
            sb, text="", bg=BG2, fg=RED,
            font=("Segoe UI", 8, "bold"), cursor="hand2")
        self._lockdown_lbl.pack(side="right", padx=(0, 10))
        self._lockdown_lbl.bind("<Button-1>", lambda _: self.open_settings())
        # Update notification label
        self._upd_note_lbl = tk.Label(
            sb, text="", bg=BG2, fg=ACCENT,
            font=("Segoe UI", 8), cursor="hand2")
        self._upd_note_lbl.pack(side="right", padx=(0, 8))

    # ── Sash persistence ─────────────────────────────────────────────────────

    def _apply_lib_sash(self, pane: ttk.PanedWindow, pos: int) -> None:
        try:
            pane.update_idletasks()
            if pane.winfo_width() > pos + 100:
                pane.sashpos(0, pos)
        except Exception:
            pass

    def _persist_lib_sash(self, pane: ttk.PanedWindow) -> None:
        try:
            pos = pane.sashpos(0)
            if pos > 0:
                self._settings["lib_sash"] = pos
                save_settings(self._settings)
        except Exception:
            pass

    # ── Refresh & filter ──────────────────────────────────────────────────────

    def refresh(self) -> None:
        self.groups = scan_all()
        _share_appdata(self.groups)
        self._rebuild_tag_filter()
        self._apply_filters()
        self.archives_tab.populate(self.groups)
        self._update_tab_titles()

    def _set_status_filter(self, status: str) -> None:
        self._filter_status = status
        for lbl, btn in self._sf_btns.items():
            active = lbl == status
            btn.configure(bg=ACCENT if active else SEL,
                          fg=BG if active else FG_DIM)
        self._apply_filters()

    def _apply_filters(self) -> None:
        search = self._search_var.get().lower().strip()
        ud = self.user_data
        sort = self._sort_var.get()

        filtered: list[GameGroup] = []
        for g in self.groups:
            # Archive-only groups have no extracted versions — skip in library
            if not g.versions:
                continue
            names = [v.folder_name for v in g.versions]
            # Hidden filter
            if not self._show_hidden and all(n in ud.hidden for n in names):
                continue
            # Search
            if search and search not in g.display_name.lower():
                continue
            # Status
            if self._filter_status != "All":
                any_played = any(detect_played(v, ud) for v in g.versions)
                if self._filter_status == "Played" and not any_played:
                    continue
                if self._filter_status == "Unplayed" and any_played:
                    continue
            # Tags — AND mode: game must have every selected tag
            #         OR mode: game must have at least one selected tag
            if self._filter_tags:
                game_tags = set(ud.tags.get(g.base_key, []))
                if self._filter_tags_mode == "all":
                    if not self._filter_tags.issubset(game_tags):
                        continue
                else:
                    if not (self._filter_tags & game_tags):
                        continue
            filtered.append(g)

        # Sort
        if sort == "Name Z–A":
            filtered.sort(key=lambda g: g.display_name.lower(), reverse=True)
        elif sort == "Most Played":
            def _total_pt(g: GameGroup) -> int:
                return sum(ud.playtime.get(v.folder_name, 0)
                           for v in g.versions)
            filtered.sort(key=_total_pt, reverse=True)
        elif sort == "Recently Played":
            def _last_pt(g: GameGroup) -> str:
                lps = [ud.last_played[v.folder_name]
                       for v in g.versions
                       if v.folder_name in ud.last_played]
                return max(lps) if lps else ""
            filtered.sort(key=_last_pt, reverse=True)
        elif sort == "Newest Version":
            filtered.sort(
                key=lambda g: parse_version_tuple(
                    g.versions[-1].version_str if g.versions else ""),
                reverse=True)
        elif sort == "Oldest Version":
            filtered.sort(
                key=lambda g: parse_version_tuple(
                    g.versions[0].version_str if g.versions else ""))

        sel_key = (self.detail_panel._group.base_key
                   if self.detail_panel._group else None)
        self.card_list.populate(filtered, ud, self.thumb_cache, sel_key)
        self._update_status(len(filtered))

    def _rebuild_tag_filter(self) -> None:
        """Rebuild the Tags dropdown from all tags currently in user_data."""
        ud = self.user_data
        all_tags: set[str] = set()
        for tag_list in ud.tags.values():
            all_tags.update(tag_list)

        self._tag_menu.delete(0, "end")

        # Match-mode toggle at the top
        self._tag_menu.add_radiobutton(
            label="Match: Any tag  (OR)",
            variable=self._tag_mode_var, value="any",
            command=self._on_tag_mode_change,
            foreground=FG, background=BG2,
            activeforeground=ACCENT, activebackground=SEL)
        self._tag_menu.add_radiobutton(
            label="Match: All tags  (AND)",
            variable=self._tag_mode_var, value="all",
            command=self._on_tag_mode_change,
            foreground=FG, background=BG2,
            activeforeground=ACCENT, activebackground=SEL)
        if all_tags:
            self._tag_menu.add_separator()

        self._tag_vars.clear()
        for tag in sorted(all_tags):
            var = tk.BooleanVar(value=tag in self._filter_tags)
            self._tag_vars[tag] = var
            self._tag_menu.add_checkbutton(
                label=tag, variable=var,
                command=self._on_tag_toggle,
                foreground=FG, background=BG2,
                activeforeground=ACCENT, activebackground=SEL)
        if all_tags:
            self._tag_menu.add_separator()
            self._tag_menu.add_command(
                label="✕  Clear all tags",
                command=self._clear_tags,
                foreground=FG_MUT, background=BG2,
                activeforeground=FG, activebackground=SEL)

    def _on_tag_toggle(self) -> None:
        self._filter_tags = {t for t, v in self._tag_vars.items() if v.get()}
        self._update_tag_btn_label()
        self._apply_filters()

    def _on_tag_mode_change(self) -> None:
        self._filter_tags_mode = self._tag_mode_var.get()
        self._update_tag_btn_label()
        self._apply_filters()

    def _update_tag_btn_label(self) -> None:
        n = len(self._filter_tags)
        if n:
            mode = "ALL" if self._filter_tags_mode == "all" else "ANY"
            self._tag_btn.configure(text=f"Tags ({n} {mode}) ▾")
            self._tag_clear_btn.configure(fg=ACCENT)
        else:
            self._tag_btn.configure(text="Tags ▾")
            self._tag_clear_btn.configure(fg=FG_MUT)

    def _clear_tags(self) -> None:
        for v in self._tag_vars.values():
            v.set(False)
        self._filter_tags.clear()
        self._update_tag_btn_label()
        self._apply_filters()

    # ── Selection ─────────────────────────────────────────────────────────────

    def select_group(self, base_key: str,
                     version_folder: str | None = None) -> None:
        ud = self.user_data
        # Deselect old card
        if self.detail_panel._group:
            old = self.card_list._cards.get(
                self.detail_panel._group.base_key)
            if old and old.winfo_exists():
                old.set_selected(False)

        g = next((gg for gg in self.groups
                  if gg.base_key == base_key), None)
        if not g:
            self.detail_panel.show_empty()
            return

        card = self.card_list._cards.get(base_key)
        if card and card.winfo_exists():
            card.set_selected(True)
            self.card_list.scroll_to(base_key)

        self.detail_panel.show_group(g, ud, self.thumb_cache)

    # ── Game exit callback (called from PlayTracker monitor thread via after())

    def on_game_exit(self, folder_name: str, elapsed: int) -> None:
        ud = self.user_data
        ud.playtime[folder_name] = ud.playtime.get(folder_name, 0) + elapsed
        ud.last_played[folder_name] = datetime.datetime.now().isoformat()
        ud.play_count[folder_name] = ud.play_count.get(folder_name, 0) + 1
        # Auto-mark as played if session was meaningful
        if elapsed > 60:
            ud.manual_unplayed.discard(folder_name)
            ud.manual_played.add(folder_name)
        save_userdata(ud)

        base_key = next((g.base_key for g in self.groups
                         for v in g.versions
                         if v.folder_name == folder_name), None)
        if base_key:
            self.card_list.update_card(base_key, ud)
            card = self.card_list._cards.get(base_key)
            if card and card.winfo_exists():
                card.set_playing(False)
            if (self.detail_panel._group and
                    self.detail_panel._group.base_key == base_key):
                self.detail_panel._refresh_stats(ud)
                self.detail_panel._btn_launch.configure(
                    state="normal", text="▶  Launch")
        self._update_tab_titles()

    # ── Archive deletion (also callable from detail panel) ────────────────────

    def _action_delete_archives(self, g: GameGroup) -> None:
        if not g.archives:
            return
        qw = self._queue_window

        def _is_active(a: Archive) -> bool:
            return bool(qw and qw.is_active(a.archive_path))

        if len(g.archives) == 1:
            a = g.archives[0]
            if _is_active(a):
                messagebox.showwarning(
                    "Cannot Delete",
                    f"{a.archive_path.name}\n\n"
                    "This archive is currently being extracted.\n"
                    "Cancel it in the Extraction Queue first.")
                return
            if not messagebox.askyesno(
                    "Confirm Delete",
                    f"Permanently delete:\n{a.archive_path.name}?",
                    default="no"):
                return
            try:
                a.archive_path.unlink()
            except OSError as e:
                messagebox.showerror("Error", str(e))
                return
            self.refresh()
        else:
            win = tk.Toplevel(self)
            win.title("Delete Archives")
            win.configure(bg=BG)
            win.grab_set()
            win.resizable(False, False)
            ttk.Label(win, text="Select archives to delete:",
                      font=("Segoe UI", 10, "bold"),
                      padding=(16, 12, 16, 4)).pack(anchor="w")
            bvars: list[tuple[tk.BooleanVar, Archive]] = []
            for a in g.archives:
                bv = tk.BooleanVar(value=False)
                active = _is_active(a)
                label = a.archive_path.name + (" (extracting — skip)" if active else "")
                tk.Checkbutton(
                    win, text=label,
                    variable=bv, bg=BG,
                    fg=FG_DIM if active else FG,
                    selectcolor=SEL,
                    activebackground=BG, activeforeground=FG,
                    font=("Segoe UI", 9),
                    state="disabled" if active else "normal",
                ).pack(anchor="w", padx=24, pady=1)
                if not active:
                    bvars.append((bv, a))

            def _do():
                for bv, a in bvars:
                    if bv.get():
                        try:
                            a.archive_path.unlink()
                        except OSError as e:
                            messagebox.showerror("Error", str(e))
                win.destroy()
                self.refresh()

            ttk.Button(win, text="Delete Selected",
                       style="Danger.TButton",
                       command=_do).pack(pady=(8, 14))
            self.wait_window(win)

    # ── Misc ──────────────────────────────────────────────────────────────────

    # ── Update checker ────────────────────────────────────────────────────────

    def _start_update_check(self) -> None:
        if not self.net_ok("check_updates"):
            return
        threading.Thread(target=self._fetch_latest_version, daemon=True).start()

    def _fetch_latest_version(self) -> None:
        try:
            req = urllib.request.Request(
                UPDATE_URL,
                headers={"User-Agent": f"VN-Pathfinder/{APP_VERSION}"})
            with urllib.request.urlopen(req, timeout=6) as resp:
                data = json.loads(resp.read())
            tag = data.get("tag_name", "")
            self.after(0, lambda t=tag: self._on_update_result(t))
        except Exception:
            pass

    def _on_update_result(self, tag: str) -> None:
        if not tag:
            return
        if _version_tuple(tag) > _version_tuple(APP_VERSION):
            self._upd_note_lbl.configure(
                text=f"● {tag} available — download ↗",
                fg=GREEN)
            self._upd_note_lbl.bind(
                "<Button-1>",
                lambda _: webbrowser.open(RELEASES_URL))
        else:
            self._upd_note_lbl.configure(
                text="✓ Up to date", fg=FG_DIM)

    def net_ok(self, key: str) -> bool:
        """Return True only if the feature is enabled AND lockdown is off."""
        if self._settings.get("lockdown", True):
            return False
        default = bool(_SETTINGS_DEFAULTS.get(key, True))
        return bool(self._settings.get(key, default))

    def _update_lockdown_indicator(self) -> None:
        locked = self._settings.get("lockdown", True)
        if locked:
            self._lockdown_lbl.configure(text="🔒 LOCKDOWN  (click to open Settings)")
        else:
            self._lockdown_lbl.configure(text="")

    def open_settings(self) -> None:
        SettingsDialog(self)

    def _toggle_hidden(self) -> None:
        self._show_hidden = not self._show_hidden
        self._btn_hidden.configure(
            text="Hide Hidden" if self._show_hidden else "Show Hidden")
        self._apply_filters()

    def _open_orphan_cleaner(self) -> None:
        OrphanedFilesDialog(self, self.groups, on_done=self.refresh)

    def open_cookie_settings(self) -> None:
        if not self.net_ok("allow_provider_login"):
            messagebox.showinfo(
                "Disabled",
                "Provider logins are disabled.\n\n"
                "Enable them in Settings → Connection.",
                parent=self)
            return
        SiteCookiesDialog(self)

    def queue_extraction(self, archive: Archive) -> None:
        """Add an archive to the extraction queue, creating the window if needed."""
        if self._queue_window is None:
            self._queue_window = ExtractionQueueWindow(self)
        self._queue_window.add_job(archive)

    def _update_tab_titles(self) -> None:
        ud = self.user_data
        n_lib = sum(1 for g in self.groups if g.versions)
        n_arc = sum(len(g.archives) for g in self.groups)
        n_played = sum(1 for g in self.groups
                       if g.versions and any(
                           detect_played(v, ud) for v in g.versions))
        self.notebook.tab(0, text=f"  Library ({n_lib})  ")
        self.notebook.tab(1, text=f"  Archives ({n_arc})  ")
        # Update sub-filter button counts
        n_unplayed = sum(1 for g in self.groups
                         if g.versions and not any(
                             detect_played(v, ud) for v in g.versions))
        self._sf_btns["All"].configure(text=f"All ({n_lib})")
        self._sf_btns["Played"].configure(text=f"Played ({n_played})")
        self._sf_btns["Unplayed"].configure(text=f"Unplayed ({n_unplayed})")

    def _update_status(self, visible: int) -> None:
        ud = self.user_data
        total = sum(len(g.versions) for g in self.groups)
        played = sum(1 for g in self.groups
                     for v in g.versions if detect_played(v, ud))
        arcs = sum(len(g.archives) for g in self.groups)
        unext = sum(1 for g in self.groups
                    for a in g.archives if not a.matched_folder)
        self._status_lbl.configure(
            text=(f"  Showing {visible}  ·  {total} versions  ·  "
                  f"{played} played  ·  "
                  f"{arcs} archives ({unext} not extracted)"))


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    multiprocessing.freeze_support()   # required for PyInstaller multiprocessing
    app = LibraryApp()
    app.mainloop()
