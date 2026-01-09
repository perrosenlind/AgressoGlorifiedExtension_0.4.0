# Glorify Agresso (Autosave) — v0.4.0

A small content-script extension that improves Agresso's UI and provides an inline autosave helper with a lightweight floating indicator.

## Quick install (Chrome / Chromium / Edge)

1. Open `chrome://extensions/`.
2. Enable **Developer mode** (top-right).
3. Click **Load unpacked** and choose the folder `AgressoGlorifiedExtension_0.4.0`.
4. Confirm the extension appears and is enabled.

Notes:
- The extension runs on `https://agresso.advania.se/*` (see `manifest.json`).
- It injects `cells.js` and `styles.css` as a content script (configured to run in all frames when possible).

## Temporary install (Firefox)

1. Open `about:debugging#/runtime/this-firefox`.
2. Click **Load Temporary Add-on…** and select the `manifest.json` file.

Warning: Temporary add-ons are removed on restart.

## Main features

- Idle autosave: watches for user activity and triggers a save when the user is idle for a configured period.
- Shortcut-based save: attempts a keyboard shortcut (Alt+S on Windows/Linux, Option+S on macOS) and falls back to clicking detected save buttons.
- Dialog sweep: searches for and dismisses save/confirmation dialogs after a save attempt.
- Floating indicator: small status UI injected into the page (or top-level document when accessible) showing state and an on/off toggle.
- Persisted toggle: the autosave toggle state is stored in `localStorage` under `agresso_autosave_enabled` (default ON).
- Layout tweaks: applies a few field sizing and labeling improvements to improve usability.

## Permissions and behavior notes

- `manifest.json` requests only `activeTab` and a host permission for `https://agresso.advania.se/*`.
- The content script is configured with `all_frames: true`. The indicator prefers to inject into the top-level same-origin document so it remains visible when the script runs inside frames.

## Configuration

Most behavior is controlled by constants at the top of `cells.js`. Common values:

- `IDLE_TIMEOUT_MS` — how long to wait after user inactivity before saving.
- `SAVE_DEBOUNCE_MS`, `SAVE_COOLDOWN_MS` — debounce/cooldown around save attempts.
- `DIALOG_SWEEP_MS` — how long to run dialog dismissal sweeps.

Selectors and keywords in `cells.js`:

- `SAVE_BUTTON_SELECTORS` — DOM selectors used to find fallback save buttons.
- `SAVE_DIALOG_SELECTORS` and `SAVE_DIALOG_KEYWORDS` — used to detect save confirmation dialogs.

Edit `cells.js` and reload the extension to apply changes.

## How autosave works (high level)

1. The script watches activity events (typing, clicks, mouse/pointer movement, touch, scroll) and resets a unified idle timer.
2. When the idle timer elapses, the script simulates the configured save shortcut and then tries fallback save button selectors.
3. After a save attempt a dialog sweep runs for `DIALOG_SWEEP_MS` to find and dismiss confirmation dialogs.

## Floating indicator & toggle

- Indicator id: `agresso-autosave-indicator`.
- Toggle persistence key: `agresso_autosave_enabled` (value `1`/`0`).
- Default: ON. When OFF the script stops timers and will not auto-dismiss dialogs.

Debug controls (hidden by default): set `INDICATOR_DEBUG = true` in `cells.js` or call `window.agresso_setIndicatorDebug(true)` in the Console during the session to reveal extra buttons for development.

## Testing / verification steps

1. Load the extension and open `https://agresso.advania.se/`.
2. Open DevTools Console and watch for messages prefixed with `[Agresso Autosave]`.
3. Edit a field/row, interact to start the timer, then stop interacting. After the idle timeout the extension should attempt a save and log actions.
4. Toggle the on/off control and verify autosave stops when OFF and resumes when ON.

## Troubleshooting

- No saves: ensure `SAVE_BUTTON_SELECTORS` matches your page or manually verify the save shortcut works.
- No dialog dismissal: cross-origin frames or CSP may block access to dialogs; check Console for errors.
- Missing indicator: top-level injection may be blocked by cross-origin frames; the indicator attempts to place itself in the top-level same-origin document when possible.

## Development

- Edit `cells.js` and `styles.css`, then reload the extension at `chrome://extensions/` → **Reload**.
- Console helpers: many internal flags and helpers are accessible via the Console during a session for debugging.

## Key files

- `manifest.json` — extension manifest (version: 0.4.0).
- `cells.js` — content script with autosave logic and UI.
- `styles.css` — styles for the floating indicator and minimal layout tweaks.
- `icons/` — extension icons.

## License

Provided as-is. Modify and adapt for your environment; contributions welcome via forks.
