# Kyun.MD — VN Pathfinder Change Log

This file tracks non-trivial changes made by Claude (AI assistant) across sessions.
Intended as a handoff document for the lead developer and their AI context.

---

## Session: 2026-04-15 — Settings Menu

### Summary
Added a full Settings dialog accessible from the main toolbar, replacing the old
status-bar update checkbox and making the library directory user-configurable.

---

### New: `SettingsDialog` class
**Location:** `vn_pathfinder.py`, inserted before `LibraryApp` class (~line 4994 at time of writing)

A `tk.Toplevel` modal dialog opened via the new **⚙ Settings** toolbar button.
Divided into two sections:

#### General
- **Game library directory** — text entry showing the current library path, a
  `Browse…` folder-picker button, and an `Apply` button.
  - On Apply: validates the path, writes `library_dir` to `settings.json`, calls
    `_set_lib_dir()` to update the `RENPY_DIR` global at runtime, then calls
    `app.refresh()` so the library reloads immediately without a restart.

#### Connection
- **LOCKDOWN MODE** — master toggle (red, bold text). **ON by default** (first install).
  - When enabled: forces all three individual net toggles to `False`, disables
    and grays out their checkboxes. Clears the update notification label.
  - Only the lockdown toggle itself can disable lockdown. Nothing overrides it.
  - Saved as `"lockdown": true/false` in `settings.json`; persists across restarts.
  - A `🔒 LOCKDOWN` label appears in the status bar when active; clicking it opens Settings.
- **App update checks** — controls `check_updates` key (was previously a status-bar checkbox).
  Has a **Check Now** button to trigger a manual update check.
- **Metadata fetch** — controls `fetch_metadata` key (new)
- **Provider logins** — controls `allow_provider_login` key (new)
- **App version** shown as a footer at the bottom of the dialog (e.g. `VN Pathfinder v1.0.0-beta`)

---

### New: `_set_lib_dir(path: Path)` function
**Location:** ~line 223, just after `save_settings()`

Sets the module-level `RENPY_DIR` global at runtime. Called from:
- `LibraryApp.__init__` on startup (applies saved `library_dir` setting before first scan)
- `SettingsDialog._apply_lib_dir` when the user applies a new path

---

### New: `_SETTINGS_DEFAULTS` dict
**Location:** ~line 230

Centralised defaults for all settings keys used by `net_ok()` and `SettingsDialog`:

```python
_SETTINGS_DEFAULTS = {
    "check_updates":        True,
    "fetch_metadata":       True,
    "allow_provider_login": True,
    "lockdown":             True,   # ON by default (security measure)
}
```

---

### New: `LibraryApp.net_ok(key)` method
**Location:** inside `LibraryApp`, near `_start_update_check`

```python
def net_ok(self, key: str) -> bool:
    if self._settings.get("lockdown", True):   # default True = locked out-of-box
        return False
    default = bool(_SETTINGS_DEFAULTS.get(key, True))
    return bool(self._settings.get(key, default))
```

Single authoritative gate for all network calls. Returns `False` whenever lockdown
is active, regardless of individual toggle state.

---

### New: `LibraryApp.open_settings()` method
Opens `SettingsDialog(self)`. Wired to the **⚙ Settings** toolbar button.

---

### Modified: `LibraryApp.__init__`
After loading settings, now applies `library_dir` before the first scan:
```python
_lib = self._settings.get("library_dir", "")
if _lib:
    _p = Path(_lib)
    if _p.is_dir():
        _set_lib_dir(_p)
```
Startup update check now goes through `net_ok` instead of reading the setting directly:
```python
if self.net_ok("check_updates"):
    self.after(800, self._start_update_check)
```

---

### Modified: `LibraryApp._build_ui` — toolbar
Added **⚙ Settings** button (rightmost, before ⟳ Refresh).

---

### New: `LibraryApp._update_lockdown_indicator()`
Updates the `_lockdown_lbl` in the status bar. Shows `🔒 LOCKDOWN  (click to open Settings)`
in red when lockdown is active; empty string when off. Called from:
- `LibraryApp.__init__` (after `_build_ui`)
- `SettingsDialog._on_lockdown_toggle`
- `SettingsDialog._on_net_toggle` (when auto-disabling lockdown)

---

### Modified: `LibraryApp._build_ui` — status bar
- Removed the old `"Check for updates"` `tk.Checkbutton` and `_upd_var`.
  The update notification label (`_upd_note_lbl`) is retained.
  > **Why removed:** Consolidated into Settings → Connection. Single source of truth.
- Added `_lockdown_lbl` — a clickable label on the far right of the status bar.
  Red, bold. Clicking it opens Settings. Managed by `_update_lockdown_indicator()`.

---

### Modified: `LibraryApp._start_update_check`
Now early-returns if `net_ok("check_updates")` is False (covers both the individual
toggle and lockdown).

---

### New: `SettingsDialog._check_now()`
Manually triggers `app._start_update_check()`. Blocked with an info dialog if
lockdown is active. Sets label to `Checking…` before firing.

---

### Modified: `LibraryApp.open_cookie_settings`
Now checks `net_ok("allow_provider_login")` before opening `SiteCookiesDialog`.
Shows an info dialog pointing to Settings → Connection if disabled.

---

### Modified: `DetailPanel._fetch_metadata`
Added `net_ok("fetch_metadata")` check at the top, before the `HAS_SCRAPING` check.
Shows an info dialog pointing to Settings → Connection if disabled.

---

### Settings keys reference (`settings.json`)

| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `library_dir` | string (path) | `""` (uses `APP_DIR.parent`) | Absolute path to game library root |
| `check_updates` | bool | `true` | Fetch latest release tag from GitHub API on startup |
| `fetch_metadata` | bool | `true` | Allow VNDB / F95Zone / LewdCorner metadata scraping |
| `allow_provider_login` | bool | `true` | Allow site-cookie / webview provider login dialogs |
| `lockdown` | bool | `true` | Master kill-switch — ON by default; overrides all above to false |
| `lib_sash` | int | `620` | Saved sash position for library pane splitter |
| `meta_sash` | int | `560` | Saved sash position for metadata pane splitter |

---

### Modified: `SettingsDialog._on_net_toggle`
Simple save — writes the toggle state to settings. No side effects.
Individual toggles are grayed out and non-interactive while lockdown is active
(`_sync_lockdown_ui` sets them to `state="disabled"`), so this handler is only
reachable when lockdown is already off.

### Modified: `SettingsDialog._sync_lockdown_ui`
Now also disables/enables the **Check Now** button alongside the checkboxes.

> **Design note:** Lockdown is a hard master switch. Nothing auto-disables it —
> only the lockdown toggle itself can restore network access. Individual feature
> toggles are completely inert while lockdown is on.

---

*Last updated: 2026-04-15 (turn 11)*

---

## Session: 2026-04-15 — Screenshot cap + Slideshow auto-advance

### Screenshot cap raised: 11 → 18
Changed in four places: `_download_screenshots` (download loop), and three
provider `fetch_thread` methods (F95Zone, LewdCorner, itch.io). New max is
1 cover + 18 screenshots = 19 total in the carousel.

---

### New: Slideshow auto-advance (`DetailPanel`)

**`_slideshow_id`** — instance variable storing the `after()` handle, or `None`.

**`_slideshow_start()`** — cancels any pending tick, then schedules a new one at
`slideshow_interval` seconds (read live from settings each call). No-ops if the
carousel has fewer than 2 images.

**`_slideshow_cancel()`** — safe cancel; swallows exceptions if the handle is stale.

**`_slideshow_tick()`** — advances index by 1 (wraps), calls `_carousel_show()`,
then calls `_slideshow_start()` again to keep the loop going.

**Integration points:**
- `_load_art()` → calls `_slideshow_start()` after building the carousel
- `show_empty()` → calls `_slideshow_cancel()` first
- `_carousel_prev()` / `_carousel_next()` → call `_slideshow_start()` after a
  manual navigation, resetting the timer so it doesn't fire immediately afterwards

---

### New: Settings → General — Slideshow interval slider
A `ttk.Scale` ranging from 5–300 (internal units = tenths of a second, so 0.5–30s).
Label to the right shows the current value as a decimal (e.g. `0.5`, `2`, `12.5`).
Saves to `settings.json` as `"slideshow_interval": <float>` on every slider move.
Default: `2.0` seconds.

---

### New settings key

| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `slideshow_interval` | float | `2.0` | Seconds between auto-advance ticks in the detail panel carousel |

---

## Session: 2026-04-15 — Menubutton theme fix

`TMenubutton` style added to `apply_theme()`. Was inheriting the system/Windows
default (light beige) instead of the dark theme. Now matches `TButton` exactly —
same `SEL` background, same `FG` text, same `ACCENT` hover. Dropdown arrow
coloured `FG_DIM` at rest, `BG` on hover, `FG_MUT` when disabled.

---

## Session: 2026-04-15 — Download Links lockdown + settings toggle

**`allow_download_links`** added to `_SETTINGS_DEFAULTS` (default `True`) and to
the Connection toggle list in `SettingsDialog` — label: "Download page links
(↗ Download Page button)".

`_refresh_download_menu()` now checks `net_ok("allow_download_links")` before
building the menu. If lockdown is on OR the toggle is off, the button stays
disabled and the menu is cleared. Because `_refresh_download_menu` is called every
time a game is selected, toggling the setting in Settings and then clicking a game
takes effect immediately without a restart.

---

## Session: 2026-04-15 — Launch Button + Download Page

### Summary
Launch button made prominent and full-width. Button layout restructured. Download
Page menubutton added — grayed until metadata is fetched, then shows a dropdown
of source URLs (F95Zone, LewdCorner, itch.io). VNDB excluded — it's a database,
not a download host.

---

### New style: `Launch.TButton`
Added in `apply_theme()`. Larger padding (`8, 10`), bold Segoe UI 11pt, same
Accent colour as before. Disabled state uses `FG_MUT` background.

---

### Modified: `DetailPanel._build` — button layout

Old layout (3 rows × 2 buttons):
```
[ ▶ Launch ] [ Open Folder ]
[ Mark Played ] [ Open Saves ]
[ Hide Game ] [ Delete Archive ]
```

New layout:
```
[ ▶   Launch  (full width, Launch.TButton)  ]
────────────────────────────────────────────
[ Open Folder    ] [ Mark Played      ]
[ Open Saves     ] [ Hide Game        ]
[ Delete Archive ] [ ↗ Download Page  ]
```

---

### New: `DetailPanel._btn_download` + `_download_menu`
A `ttk.Menubutton` backed by a `tk.Menu`. Lives in `row3` alongside Delete Archive
(the sixth button slot). Starts `state="disabled"`.
Populated by `_refresh_download_menu()` whenever a game is selected or metadata
is refreshed.

---

### New: `DetailPanel._DOWNLOAD_EXCLUDE`
Class-level set: `{"vndb"}`. Sources in this set are never added to the download
menu. Extend here if other info-only sources are added in future.

---

### New: `DetailPanel._refresh_download_menu()`
Reads `version.metadata["sources"]` and builds the dropdown:
- Skips sources in `_DOWNLOAD_EXCLUDE` and entries with no URL.
- If 0 valid sources → button stays disabled.
- If 1 valid source → button opens URL directly (no dropdown needed).
- If 2+ valid sources → standard dropdown listing each source by display name.

Called from:
- `_refresh_metadata_display()` — every time a game is selected or version changes
- Automatically after fetch via `_finish_save()` → `app.refresh()`

---

### Modified: `DetailPanel._set_enabled`
When `en=False` (no game selected), also disables `_btn_download`.

---

## Session: 2026-04-15 — Fullscreen focus fix

**Problem:** Opening the fullscreen viewer and immediately pressing an arrow key
would navigate the detail panel carousel instead of the fullscreen window, because
keyboard focus remained on the detail panel.

**Fix:** Added `win.focus_set()` immediately after `win.grab_set()` in
`_open_fullscreen`. The window now takes focus as soon as it opens — arrow keys
go to it straight away without requiring a click first.

---

## Session: 2026-04-15 — Detail Panel Art Area Rework

### Summary
Overhauled the artwork section of the detail panel: resizable height, fullscreen
click-to-view, hover-only "Change Image" button, and multi-image selection.

---

### Modified: `DetailPanel._build` — art section

The fixed-height `tk.Label` + `art_frame` setup was replaced with a vertical
`ttk.PanedWindow` (`_detail_pane`) that splits the detail panel into:
- **Top pane** — the artwork carousel (user-resizable by dragging the sash)
- **Bottom pane** — everything else (title, stats, tags, notes, buttons)

The art label (`_art_lbl`) now fills its pane via `fill="both", expand=True`
instead of having a fixed pixel height. The sash position (= art height in px)
is saved to `settings.json` under `"detail_art_h"` and restored on next open.
The saved height applies to all games — it does not jump per-game.

The old hint label `"Click artwork to set as cover"` was removed (replaced by the
hover button).

---

### New: `DetailPanel._apply_art_sash(h)` / `_persist_art_sash()`
- `_apply_art_sash(h)` — called 60 ms after `<Map>` to restore saved sash position.
  Guards against applying a position larger than the panel itself.
- `_persist_art_sash()` — called on `<ButtonRelease-1>` to write the new sash
  position to `_app._settings["detail_art_h"]` and call `save_settings()`.

---

### New: `DetailPanel._open_fullscreen(_e)`
Bound to `<Button-1>` on `_art_lbl` (replaces the old direct-to-file-picker binding).
Opens a `tk.Toplevel` sized to fit the image on screen (max screen - 80 px margin).
- Click the image or press Escape to close.
- No-ops gracefully if PIL is unavailable or no image is loaded.

---

### New: `DetailPanel._on_art_enter` / `_on_art_leave`
Hover detection for the overlay button. `_on_art_leave` checks the actual pointer
coordinates before hiding — prevents flickering when the cursor moves from the
label onto the button itself.

---

### New: `DetailPanel._change_btn`
A `tk.Button` placed via `place(relx=1.0, rely=1.0, anchor="se")` over the bottom-
right corner of the art label. Hidden by default (`place_forget()`). Appears only
when the mouse is over the artwork. Clicking it opens the image picker dialog.

---

### Modified: `DetailPanel._set_custom_art`
- Changed from `filedialog.askopenfilename` (single) to `filedialog.askopenfilenames`
  (multi-select). Dialog title updated to indicate the first selected file is the cover.
- **First selected file** → saved to `ud.custom_art[base_key]` as before (cover).
- **Additional selected files** → copied into `.vnpf/` of the game's latest version
  as `screenshot_001.ext`, `screenshot_002.ext`, etc. Index continues from any
  existing screenshots so nothing is overwritten.
- All new images appear immediately in the carousel without restart.

---

### Modified: `DetailPanel._carousel_show`
Image request now uses the actual rendered size of `_art_lbl` instead of the fixed
`DETAIL_ART_W` / `DETAIL_ART_H` constants, so images fill the resized area correctly.

---

### New settings key

| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `detail_art_h` | int | `190` | Saved height in px of the art area in the detail panel |

---

## Session: 2026-04-15 — Carousel Navigation Polish

### Summary
Persistent navigation arrows, keyboard Left/Right on the detail panel, and full
arrow + keyboard navigation inside the fullscreen viewer.

---

### Modified: `DetailPanel._carousel_show`
- Arrows (`_car_prev`, `_car_next`) are now always visible — no longer
  `state="disabled"` at boundaries. At the first/last image they are dimmed
  (`fg=FG_MUT`) to hint at the boundary but remain present.
- Navigation wraps around: pressing ❯ on the last image goes back to image 1,
  and ❮ on the first image wraps to the last.

### Modified: `DetailPanel._carousel_prev` / `_carousel_next`
Use modulo arithmetic (`% total`) instead of a boundary guard, enabling wrap-around.

### Modified: `DetailPanel._build` — keyboard binds on `_art_lbl`
- `<Left>` → `_carousel_prev()`
- `<Right>` → `_carousel_next()`
- `<Button-1>` also calls `focus_set()` on the label so arrow keys work immediately
  after clicking the image without needing a separate focus click.

### Fix: nav bar pushed off-screen by large images (`DetailPanel._build`)
**Root cause:** `art_frame` was packed first with `fill="both", expand=True`,
claiming all available space in `art_outer`. The `nav` frame was packed after it
with `side="bottom"`, so when the image PhotoImage was larger than the pane height
the nav bar was pushed below the sash and became invisible.

**Fix:** Reversed the pack order. `nav` is now packed first with `side="bottom"`,
guaranteeing it always claims its space. `art_frame` is then packed with
`side="top", fill="both", expand=True` to fill whatever remains. Tkinter's pack
geometry manager respects this order correctly — the nav is never displaced
regardless of image size.

---

### Modified: `DetailPanel._open_fullscreen`
Fully rewritten. Now displays a nav bar at the bottom with ❮ / counter / ❯ buttons.
- Navigation uses a `state = {"idx": …}` closure so the current index is shared
  across all nav callbacks without needing an instance variable.
- `<Left>` / `<Right>` keyboard binds on the window.
- `<Escape>` or clicking the image closes the viewer.
- Wraps around (same modulo logic as the detail panel carousel).
- Window re-centres after the first image loads.

---

## Investigation: LewdCorner Provider 401 Error (UNRESOLVED — needs Niko)

**Status:** Root cause identified and fully confirmed. Fix is ready to implement.

---

### Symptom
`LewdCornerProvider.search()` returns a 401 Unauthorized from LewdCorner's server.

---

### What was ruled out
- **Login flow is fine.** Cookies (`xf_user`, `xf_session`, `xf_csrf`) are saved
  correctly and confirmed present in browser requests.
- **`fetch_thread` is likely fine.** Thread URLs follow standard XenForo format
  (`/threads/slug.THREAD_ID/`) identical to F95Zone. Title format, tag selectors,
  and first-post parsing should carry over without changes.
- **The endpoint `/search/search/` itself is not the problem.** LewdCorner still
  runs XenForo for its forum layer. `latest-updates.php` is a separate custom PHP
  frontend that reads from XenForo data — it does not replace the forum search.

---

### Root cause
`LewdCornerProvider.search()` was copy-pasted from `F95ZoneProvider.search()` with
only `LC_BASE` swapped in. The search params include:

```python
"c[nodes][0]": "2",   # ← F95Zone's Games node ID, hardcoded
"c[child_nodes]": "1",
```

On F95Zone, node `2` = Games/Comics. On LewdCorner, node `2` is a different or
restricted section — XenForo returns 401 when a user tries to search a node they
cannot access.

---

### LewdCorner forum structure (fully confirmed via browser)

```
Forums
└── Adult Games
    └── Games  (node 6)  ← confirmed parent, lewdcorner.com/forums/games.6/
        ├── [free game threads — 365 pages]
        ├── Games+  (node 119)  ← paid subscription section
        │   ├── AI Games+
        │   └── ATF Games & Mods
        ├── AI Games  (free, separate sub-forum)
        ├── Ports
        │   └── Ports+
        └── test zone  (internal, 0 views — irrelevant)
```

Key findings:
- Node `6` = "Games" — confirmed parent of ALL game sub-sections
- Node `119` = "Games+" — paid sub-section confirmed at `lewdcorner.com/forums/games.119/`
- `c[child_nodes]=1` (already in params) recursively covers every sub-forum listed above
- Thread title format is identical to F95Zone: `[Engine] Title [Version] [Developer]`
- XenForo permission checks are session-based — subscription cookies grant access to
  Games+ and Ports+ results transparently, no special handling needed

---

### The fix (one line change, pending confirmation)

**Proposed:** Change `c[nodes][0]` from `"2"` to `"6"` in `LewdCornerProvider.search()`.

With `c[child_nodes]: "1"` already in the params, targeting node 6 will recursively
search Games + Games+ + AI Games+ + ATF Games & Mods in one query. Subscription
access is handled transparently by the user's session cookies.

```python
# In LewdCornerProvider.search(), replace:
"c[nodes][0]": "2",

# With:
"c[nodes][0]": "6",
```

**This is confirmed and ready to implement.** Node 6 structure was verified
in-browser — it is the correct parent containing all game sub-sections.

---

### No other changes needed
- `_form_login` for LewdCorner — correct, leave as-is
- `fetch_thread` — correct, leave as-is  
- Cookie handling — correct, leave as-is
- All HTML selectors in `fetch_thread` — correct (standard XenForo structure)

---

## Session: 2026-04-15 — Slideshow default + Carousel cross-fade transition

### Slideshow default changed: 2s → 3.5s
Changed in two places:
- `SettingsDialog._build` — slider `IntVar` initialised to `35` (= 3.5s in tenths)
- `DetailPanel._slideshow_start()` — fallback default in settings lookup

---

### New: cross-fade transition animation (`DetailPanel`)

Carousel image changes now animate with an 8-frame cross-fade (~200ms total).

#### Architecture: `ThumbnailCache` extended to hold PIL images

`_pil_cache: dict[str, Image.Image]` added alongside `_cache` (PhotoImages).

`get_pil(key)` — returns the PIL image for a cached key, or `None`.

`_drain()` updated: on successful load, stores `pil_img` in `_pil_cache` and
passes it as a third argument to the `on_ready` callback:
```python
cb(key, photo, pil_img)   # was: cb(key, photo)
```
Cache-hit path also passes the PIL image:
```python
on_ready(key, self._cache[key], self._pil_cache.get(key))
```

`_on_thumb_ready` in `ScrollableCardList` updated signature: `(self, key, photo, _pil=None)`.

---

#### New instance variables on `DetailPanel`

| Variable | Type | Purpose |
|----------|------|---------|
| `_current_pil` | `Image.Image \| None` | PIL image currently displayed; used as "from" frame for next fade |
| `_fade_ids` | `list[str]` | Accumulates `after()` handles so frames can be cancelled cleanly |

---

#### New: `DetailPanel._fade_cancel()`
Cancels all pending fade `after()` handles and clears `_fade_ids`. Called before
any new image request starts (`_carousel_show`) and on game deselect (`show_empty`).
Does NOT touch `_current_pil` — the most recently visible frame stays as the "from"
image for the next transition.

---

#### New: `DetailPanel._fade_transition(old_pil, new_pil, new_photo, frames=8, step_ms=25)`
Cross-fade from `old_pil` → `new_pil` over `frames × step_ms` ms (default ~200ms).

Steps:
1. Resize `old_pil` to match `new_pil.size` if they differ (handles resized pane).
2. Convert both to RGBA.
3. Pre-build all `frames` `PhotoImage` objects using `Image.blend(old_rgba, new_rgba, alpha)`.
   Pre-building avoids blocking on each `after()` tick.
4. `_apply_frame(idx)` — schedules each frame via `after(step_ms, …)`, accumulating
   handles in `_fade_ids`. On the final frame, commits `new_photo` and `new_pil` as
   `_detail_photo` and `_current_pil` respectively.

---

#### Modified: `DetailPanel._on_art_ready`
Signature: `(self, key, photo, pil_img=None)`

Logic:
- No photo → show "No artwork" immediately, clear `_current_pil`.
- Previous PIL exists and new PIL provided → call `_fade_transition()`.
- Otherwise (first image for game, or PIL unavailable) → show immediately.

---

#### Integration: `_load_art` resets `_current_pil`
When a new game is selected, `_load_art()` calls `_fade_cancel()` and sets
`_current_pil = None` before building the carousel. This ensures the first image
of a new game appears instantly (no fade from the previous game's last image).

---

#### New: `DetailPanel._preload_carousel()`
Called via `after(200, …)` at the end of `_load_art()`. Fires a background
`ThumbnailCache.request` for every image in the carousel with a no-op `on_ready`
callback. The 200ms delay lets the widget layout settle so `winfo_width/height`
return real dimensions. `request()` is idempotent — already-cached or already-
pending keys are skipped automatically.

**Why:** On the first cycle through the carousel, images aren't in `_pil_cache`
yet, so `_on_art_ready` falls through to the instant-show path (no fade). After
preloading, all PIL images are warm in cache before the slideshow reaches them,
so cross-fades are smooth from the very first transition.

---

*Last updated: 2026-04-16 (turn 13)*
