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

*Last updated: 2026-04-15 (turn 3)*

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
