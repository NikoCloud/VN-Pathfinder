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

RENPY_DIR     = Path(__file__).parent.resolve()
APPDATA_RENPY = Path(os.environ.get("APPDATA", "")) / "RenPy"
USERDATA_FILE = RENPY_DIR / "vn_pathfinder.json"
SETTINGS_FILE = Path(os.environ.get("APPDATA", "")) / "VN Pathfinder" / "settings.json"
ASSETS_DIR    = _resource_path("assets")

# ── Data file migration (renpy_manager.json → vn_pathfinder.json) ──────────────
_OLD_USERDATA = RENPY_DIR / "renpy_manager.json"
if _OLD_USERDATA.exists() and not USERDATA_FILE.exists():
    try:
        _OLD_USERDATA.rename(USERDATA_FILE)
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
    "gui/main_menu.png",
    "gui/main_menu.jpg",
    "gui/game_menu.png",
    "gui/game_menu.jpg",
    "gui/window_icon.png",
]

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
    if not path.is_dir():
        return False
    if (path / "game").is_dir():
        return True
    return any(path.glob("*.py"))


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
    if not is_game_dir(path):
        return None
    base_key, version_str, display_name = parse_folder_name(path.name)
    if not base_key:
        return None
    exe_path: Path | None = None
    for p in sorted(path.glob("*.exe")):
        if "-32" not in p.stem:
            exe_path = p
            break
    return GameVersion(
        folder_name=path.name,
        folder_path=path,
        base_key=base_key,
        version_str=version_str,
        display_name=display_name,
        exe_path=exe_path,
        local_save_dir=path / "game" / "saves",
        appdata_save_dir=resolve_appdata(path, base_key),
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
    for v in reversed(g.versions):
        p = find_art_path(v.folder_path)
        if p:
            return p
    return None


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
# THUMBNAIL CACHE  (thread-safe)
# ══════════════════════════════════════════════════════════════════════════════

class ThumbnailCache:
    """Load PIL images in background; convert to PhotoImage on main thread."""

    def __init__(self, root: tk.Tk) -> None:
        self._root = root
        self._cache: dict[str, "ImageTk.PhotoImage"] = {}
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
            on_ready(key, self._cache[key])
            return
        if key in self._pending:
            return
        self._pending.add(key)
        self._work_q.put((key, path, w, h, on_ready))

    def get(self, key: str) -> "ImageTk.PhotoImage | None":
        return self._cache.get(key)

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
                    cb(key, photo)
                else:
                    cb(key, self.placeholder())
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
    style.configure("Accent.TButton", background=ACCENT, foreground=BG)
    style.map("Accent.TButton",
              background=[("active", "#74c7ec")],
              foreground=[("active", BG)])
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

    def _on_thumb_ready(self, base_key: str, photo) -> None:
        card = self._cards.get(base_key)
        if card and card.winfo_exists():
            card.set_thumbnail(photo)

    def update_card(self, base_key: str, ud: UserData) -> None:
        card = self._cards.get(base_key)
        if card and card.winfo_exists():
            card.refresh_data(ud)

    def scroll_to(self, base_key: str) -> None:
        card = self._cards.get(base_key)
        if not card:
            return
        self.update_idletasks()
        total = self.inner.winfo_height()
        if total > 0:
            self.canvas.yview_moveto(card.winfo_y() / total)


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
        self._detail_photo = None               # keep art PhotoImage alive

        self._build()

    def _build(self) -> None:
        # ── Artwork ──────────────────────────────────────────────────────────
        self._art_lbl = tk.Label(
            self, bg=SEL, width=DETAIL_ART_W, height=DETAIL_ART_H,
            text="", cursor="hand2")
        self._art_lbl.pack(fill="x", padx=0, pady=0)
        self._art_lbl.bind("<Button-1>", self._set_custom_art)
        ttk.Label(self, text="Click artwork to set custom image",
                  foreground=FG_MUT, font=("Segoe UI", 7),
                  padding=(0, 1, 0, 4)).pack()

        inner = ttk.Frame(self)
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
        row1 = ttk.Frame(inner)
        row1.pack(fill="x", pady=(0, 4))
        self._btn_launch = ttk.Button(
            row1, text="▶  Launch", style="Accent.TButton",
            command=self._launch)
        self._btn_launch.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self._btn_open = ttk.Button(
            row1, text="Open Folder", command=self._open_folder)
        self._btn_open.pack(side="left", expand=True, fill="x")

        row2 = ttk.Frame(inner)
        row2.pack(fill="x", pady=(0, 4))
        self._btn_played = ttk.Button(
            row2, text="Mark Played", command=self._toggle_played)
        self._btn_played.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self._btn_saves = ttk.Button(
            row2, text="Open Saves", command=self._open_saves)
        self._btn_saves.pack(side="left", expand=True, fill="x")

        row3 = ttk.Frame(inner)
        row3.pack(fill="x")
        self._btn_hide = ttk.Button(
            row3, text="Hide Game", command=self._hide)
        self._btn_hide.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self._btn_del_arc = ttk.Button(
            row3, text="Delete Archive",
            style="Danger.TButton", command=self._delete_archive)
        self._btn_del_arc.pack(side="left", expand=True, fill="x")

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

    def show_empty(self) -> None:
        self._group = None
        self._version = None
        self._current_key = None
        self._title_lbl.configure(text="Select a game")
        self._set_enabled(False)
        self._art_lbl.configure(image="", text="", bg=SEL)
        self._detail_photo = None

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

        has_saves = bool(v.local_save_dir.exists() or
                         (v.appdata_save_dir and v.appdata_save_dir.exists()))
        self._btn_saves.configure(
            state="normal" if has_saves else "disabled")
        self._btn_launch.configure(
            state="normal" if v.exe_path else "disabled")
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
        art = _group_art_path(g, ud)
        if not art:
            placeholder = tc.detail_placeholder()
            if placeholder:
                self._art_lbl.configure(image=placeholder, text="")
                self._detail_photo = placeholder
            return
        self._art_lbl.configure(text="Loading…", image="")
        self._detail_photo = None
        key = f"detail:{g.base_key}"
        tc.request(key, art, DETAIL_ART_W, DETAIL_ART_H,
                   on_ready=self._on_art_ready)

    def _on_art_ready(self, key: str, photo) -> None:
        expected = f"detail:{self._current_key}"
        if key != expected:
            return  # user navigated away
        self._detail_photo = photo
        if photo:
            self._art_lbl.configure(image=photo, text="")
        else:
            self._art_lbl.configure(image="", text="No artwork",
                                    bg=SEL, fg=FG_DIM)

    def _set_enabled(self, en: bool) -> None:
        s = "normal" if en else "disabled"
        for b in (self._btn_launch, self._btn_open, self._btn_played,
                  self._btn_saves, self._btn_hide, self._btn_del_arc):
            b.configure(state=s)

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

    def _set_custom_art(self, _e=None) -> None:
        g = self._group
        if not g:
            return
        path = filedialog.askopenfilename(
            title="Select cover image",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.webp"),
                       ("All files", "*.*")])
        if not path:
            return
        ud = self._app.user_data
        ud.custom_art[g.base_key] = path
        save_userdata(ud)
        # Reload artwork
        # Bust cache for this key
        tc = self._app.thumb_cache
        tc._cache.pop(g.base_key, None)
        tc._cache.pop(f"detail:{g.base_key}", None)
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
    "renpy_manager.py", "renpy_manager.json", ".manager_cache",
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
        self.user_data: UserData = load_userdata()
        self.groups: list[GameGroup] = []
        self.thumb_cache: ThumbnailCache = ThumbnailCache(self)
        self.play_tracker: PlayTracker = PlayTracker()
        self._show_hidden: bool = False
        self._filter_status: str = "All"    # All / Played / Unplayed
        self._filter_tags: set[str] = set()
        self._queue_window: ExtractionQueueWindow | None = None

        self._build_ui()
        self.refresh()
        # Kick off update check 800 ms after window opens
        if self._settings.get("check_updates", True):
            self.after(800, self._start_update_check)

    # ── Build ─────────────────────────────────────────────────────────────────

    def _build_ui(self) -> None:
        # ── Top toolbar ──────────────────────────────────────────────────────
        tb = ttk.Frame(self, padding=(12, 8))
        tb.pack(fill="x")
        ttk.Label(tb, text="RenPy Library",
                  font=("Segoe UI", 15, "bold"),
                  foreground=ACCENT).pack(side="left")

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

        # Tag filter
        self._tag_btn = ttk.Menubutton(tb, text="Tags ▾")
        self._tag_menu = tk.Menu(self._tag_btn, tearoff=False,
                                  bg=BG2, fg=FG, activebackground=SEL,
                                  activeforeground=ACCENT, bd=0)
        self._tag_btn.configure(menu=self._tag_menu)
        self._tag_btn.pack(side="right", padx=(16, 2))
        self._tag_vars: dict[str, tk.BooleanVar] = {}

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
        pane = ttk.PanedWindow(self._lib_frame, orient="horizontal")
        pane.pack(fill="both", expand=True)

        left = ttk.Frame(pane)
        pane.add(left, weight=2)
        self.card_list = ScrollableCardList(
            left, on_select_cb=self.select_group)
        self.card_list.pack(fill="both", expand=True)

        right = ttk.Frame(pane)
        pane.add(right, weight=1)
        self.detail_panel = DetailPanel(right, app=self)
        self.detail_panel.pack(fill="both", expand=True)

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
        # Update checker controls (right side of status bar)
        self._upd_note_lbl = tk.Label(
            sb, text="", bg=BG2, fg=ACCENT,
            font=("Segoe UI", 8), cursor="hand2")
        self._upd_note_lbl.pack(side="right", padx=(0, 4))
        self._upd_var = tk.BooleanVar(
            value=self._settings.get("check_updates", True))
        tk.Checkbutton(
            sb, text="Check for updates",
            variable=self._upd_var, bg=BG2, fg=FG_DIM,
            selectcolor=SEL, activebackground=BG2, activeforeground=FG,
            font=("Segoe UI", 8),
            command=self._on_update_toggle
        ).pack(side="right", padx=(0, 8))

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
            # Tags (OR — game must have at least one selected tag)
            if self._filter_tags:
                game_tags = set(ud.tags.get(g.base_key, []))
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
        self._tag_vars.clear()
        for tag in sorted(all_tags):
            var = tk.BooleanVar(value=tag in self._filter_tags)
            self._tag_vars[tag] = var
            self._tag_menu.add_checkbutton(
                label=tag, variable=var,
                command=self._on_tag_toggle)
        if all_tags:
            self._tag_menu.add_separator()
            self._tag_menu.add_command(label="Clear all tags",
                                        command=self._clear_tags)

    def _on_tag_toggle(self) -> None:
        self._filter_tags = {t for t, v in self._tag_vars.items() if v.get()}
        n = len(self._filter_tags)
        self._tag_btn.configure(
            text=f"Tags ({n}) ▾" if n else "Tags ▾")
        self._apply_filters()

    def _clear_tags(self) -> None:
        for v in self._tag_vars.values():
            v.set(False)
        self._filter_tags.clear()
        self._tag_btn.configure(text="Tags ▾")
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

    def _on_update_toggle(self) -> None:
        enabled = self._upd_var.get()
        self._settings["check_updates"] = enabled
        save_settings(self._settings)
        if enabled:
            self._upd_note_lbl.configure(text="")
            self._start_update_check()
        else:
            self._upd_note_lbl.configure(
                text="○  Updates disabled", fg=FG_DIM)
            self._upd_note_lbl.unbind("<Button-1>")

    def _toggle_hidden(self) -> None:
        self._show_hidden = not self._show_hidden
        self._btn_hidden.configure(
            text="Hide Hidden" if self._show_hidden else "Show Hidden")
        self._apply_filters()

    def _open_orphan_cleaner(self) -> None:
        OrphanedFilesDialog(self, self.groups, on_done=self.refresh)

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
    app = LibraryApp()
    app.mainloop()
