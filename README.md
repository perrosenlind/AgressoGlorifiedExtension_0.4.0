# Glorify Agresso (Autosave)

This small extension improves Agresso's UI and adds an autosave helper that watches for edits, performs saves via keyboard shortcut, and dismisses save dialogs automatically.

## Installation (Chrome / Chromium / Edge)

1. Open Chrome (or a Chromium-based browser) and go to `chrome://extensions/`.
2. Enable **Developer mode** (top-right).
3. Click **Load unpacked** and choose the folder:

   - `AgressoGlorifiedExtension_0.4.0`

4. Confirm the extension appears in the list and is enabled.

Notes:
- The extension is configured to run on `https://agresso.advania.se/*` (see `manifest.json`).
- The extension uses a content script (`cells.js`) and a stylesheet (`styles.css`).

## Installation (Firefox, temporary)

1. Open `about:debugging#/runtime/this-firefox`.
2. Click **Load Temporary Add-on…** and select the `manifest.json` file inside the extension folder.

Warning: Temporary add-ons are removed when the browser restarts.

## What it does

- Watches input fields and table rows for edits. When the user stops interacting for a short idle period, it triggers a save.
- Triggers the save via simulated keyboard shortcut (Alt+S or Option+S), and falls back to clicking any detected save button.
- Shows a small floating status indicator with save/pending/error states at the right of the page.
- Detects and dismisses save confirmation dialogs automatically (including in background/unfocused windows when possible).
- Applies minor layout/column sizing improvements for better usability.

## Recent changes

- 2026-01-09: Removed the acknowledgement feature (ack button and stored acknowledgement state). This change removes the extra acknowledge controls and storage; other reminder and period-highlight behavior remains unchanged.

## New: On/Off Toggle (Autosave control)

- A clickable on/off toggle is embedded into the floating indicator (left side). It persists its state in `localStorage` under the key `agresso_autosave_enabled`.
- Default behavior: the extension now forces the toggle to ON at startup so autosave is enabled by default.
- When toggled OFF:
   - Autosave timers are stopped and automatic saves are skipped.
   - Dialog sweeping/dismissal is disabled (the script will not auto-dismiss save dialogs).
   - The indicator shows **Autosave disabled** and the main label turns red; the switch appears grey.
- When toggled ON:
   - Autosave resumes (timer restarts) and dialog sweeping is enabled again.

The toggle is intentionally placed where the previous green status dot appeared so it's visible and easy to reach.

## Key files

- `manifest.json` — extension manifest and host permissions.
- `cells.js` — main content script implementing autosave, timers, dialog sweeping, and UI indicator.
- `styles.css` — indicator and layout styles.

## Configuration (quick)

Most behavior is controlled by constants near the top of `cells.js` (e.g., `IDLE_TIMEOUT_MS`, `SAVE_DEBOUNCE_MS`, `DIALOG_SWEEP_MS`). Edit and re-load the extension to change values.

Important selectors and keywords are defined in `cells.js`:
- `SAVE_BUTTON_SELECTORS` — selectors to find save buttons.
- `SAVE_DIALOG_SELECTORS` / `SAVE_DIALOG_KEYWORDS` — what the script treats as save dialogs.

## How it decides to save

1. On any activity (typing, clicks, pointer/mouse movement, touch, scroll), the timer restarts.
2. When the timer completes (no activity for `IDLE_TIMEOUT_MS`), the script triggers the keyboard shortcut and then attempts a fallback save button click.
3. A dialog sweep runs for a short period after a save to find and dismiss save confirmation dialogs.

## Testing / Verification

1. Open `https://agresso.advania.se/` and load the extension.
2. Open DevTools (Console) to watch logs — look for messages prefixed with `[Agresso Autosave]`.
3. Edit a row or field. Interact (type, click, scroll) — you should see "Timer started" or other timer logs in the console.
4. Stop interacting; after the idle timeout the extension should attempt to save and you should see logs like `Saving via shortcut` and `Dialog dismissed` if a dialog appeared.
5. While the browser/tab is unfocused, try triggering a save; the script will attempt to dismiss background dialogs as well.

### Verifying the toggle

1. Toggle the switch to OFF and confirm the indicator label reads **Autosave disabled** in red.
2. Make an edit and wait past the idle timeout — no automatic save should be attempted (check Console for lack of `Saving via shortcut`).
3. Also confirm that save confirmation dialogs are not auto-dismissed while disabled.
4. Toggle back ON to restore normal autosave behavior.

If you don't see expected behavior:

- Ensure the extension is enabled and matches the Agresso URL.
- Check console for errors (cross-origin frames can prevent some behavior).
- Verify the constants at the top of `cells.js` haven't been changed to very large values.

## Troubleshooting

- No saves: make sure `SAVE_BUTTON_SELECTORS` matches your page or that keyboard shortcut works manually.
- No dialog dismissal: cross-origin frames or strict CSP may prevent access to dialog elements; check Console for cross-origin errors.
- Too many warnings: the script logs warnings for missing buttons/dialogs, but only once per session for each type to reduce noise.

## Development / Iteration

Edit `cells.js` and then reload the extension via `chrome://extensions/` → **Reload** for the unpacked extension.

## License & Credits

Provided as-is. You can modify and adapt for your environment. Report issues or push improvements in your fork.
