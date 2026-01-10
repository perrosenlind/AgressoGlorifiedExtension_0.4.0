(() => {
  const SAVE_DEBOUNCE_MS = 1800;
  const SAVE_COOLDOWN_MS = 4500;
  const LAYOUT_REFRESH_MS = 200;
  const DIALOG_SWEEP_MS = 8000;
  const DIALOG_SWEEP_INTERVAL_MS = 120;
  const DELETE_KEYWORDS = ['ta bort', 'delete', 'remove'];
  const ADD_KEYWORDS = ['lägg till', 'add', 'new row'];
  const IDLE_TIMEOUT_MS = 15000;
  const IDLE_RETRY_MS = 1200;
  const NO_CHANGES_POLL_MS = 600;
  const IS_MAC = /mac/i.test(navigator.platform);
  const SHORTCUT_LABEL = IS_MAC ? 'Option+S' : 'Alt+S';
  const LOG_PREFIX = '[Agresso Autosave]';
  try { console.info(LOG_PREFIX, 'cells.js loaded'); } catch (e) {}
  const NO_CHANGES_TEXT = 'inga ändringar gjorda!';
  // Pages sometimes show an alternate no-data banner text. Detect that too.
  const NO_CHANGES_TEXT_ALT = 'tidrapporten är tom. inga data har sparats.';
  const SAVE_DIALOG_SELECTORS = [
    '[id^="u4_messageoverlay_success"]',
    '.u4-messageoverlay-success-header',
    '.u4-messageoverlay-success-body',
    '.u4-messageoverlay-success-footer'
  ];
  const SAVE_DIALOG_KEYWORDS = [
    'spara',
    'sparade',
    'sparats som utkast',
    'tidrapport',
    'utkast',
    'genomfört',
    'save',
    'saved',
    'uppdatera'
  ];
  const SAVE_BUTTON_SELECTORS = [
    'button[data-cmd="save"]',
    'button[data-action="save"]',
    'button[id*="save"]',
    'button[name*="save"]',
    'input[type="submit"][value*="Save"]',
    'input[type="button"][value*="Save"]',
    'input[type="submit"][value*="Spara"]',
    'input[type="button"][value*="Spara"]',
    'input[type="submit"][value*="Uppdatera"]',
    'input[type="button"][value*="Uppdatera"]',
    'button[title*="Save"]',
    'button[title*="Spara"]',
    'button[title*="Uppdatera"]',
    'a[data-cmd="save"]',
    'a[menu_id="TS294"]',
    'a[data-menu-id="TS294"]',
    'a[href*="menu_id=TS294"]',
    'a[href*="type=topgen"][href*="TS294"]'
  ];
  const SHORTCUT_COMBOS = [
    { altKey: true, metaKey: false },
    { altKey: false, metaKey: true },
    { altKey: true, metaKey: true }
  ];

  // Unified timer state
  let unifiedTimer = null;
  let timerStartedAt = 0;
  let timerDuration = IDLE_TIMEOUT_MS;
  let timerReason = 'idle';
  
  let layoutTimer = null;
  let lastSaveAt = 0;
  let pendingRow = null;
  let dialogSweepTimer = null;
  let dialogSweepEndAt = 0;
  let periodStatusRefreshTimer = null;
  let periodHighlightEnforcer = null;
  let dropdownActive = false;
  let dropdownRow = null;
  let lastActivityAt = Date.now();
  let dialogMissLogged = false;
  let noChangesBannerVisible = false;
  let noChangesPollTimer = null;
  let timerBar = null;
  const trackedWindows = new Set();
  const trackedActivityDocs = new Set();

  function getAllDocuments() {
    const seen = new Set();
    const docs = [];

    const enqueue = (doc) => {
      if (!doc || seen.has(doc)) {
        return;
      }
      seen.add(doc);
      docs.push(doc);

      try {
        const frames = doc.querySelectorAll('iframe, frame');
        frames.forEach((f) => {
          try {
            enqueue(f.contentDocument);
          } catch (e) {
            // ignore cross-origin frames
          }
        });
      } catch (e) {
        // ignore
      }
    };

    enqueue(document);
    try {
      if (window.top && window.top.document) {
        enqueue(window.top.document);
      }
    } catch (e) {
      // ignore cross-origin access to top
    }

    return docs;
  }

  const FIELD_STYLE_RULES = [
    {
      selector: '[data-fieldname="timecode"]',
      style: 'width: 35px !important;max-width: 50px !important;'
    },
    {
      selector: '[data-fieldname="work_order"]',
      style: 'width: 75px !important;max-width: 100px !important;min-width: 50px !important'
    },
    {
      selector: '[data-fieldname="activity"]',
      style: 'width: 46px !important;max-width: 50px !important;'
    },
    {
      selector: '[data-fieldname="description"]',
      style: 'width: 500px !important; white-space: break-spaces !important;'
    },
    {
      selector: '[data-fieldname="ace_code"]',
      style: 'width: 0px !important;max-width: 0px !important;'
    },
    {
      selector: '[data-fieldname="work_type"]',
      style: 'width: 0px !important;max-width: 0px !important;'
    },
    {
      selector: '[data-fieldname="reg_unit"]',
      style: 'width: 46px !important;max-width: 50px !important;'
    }
  ];

  // Apply configured field style rules to matching elements.
  function applyFieldSizing() {
    try {
      FIELD_STYLE_RULES.forEach((r) => {
        try {
          const els = document.querySelectorAll(r.selector);
          els.forEach((el) => {
            try {
              el.style.cssText = (el.style.cssText || '') + ';' + r.style;
            } catch (e) {}
          });
        } catch (e) {}
      });
    } catch (e) {
      // ignore
    }
  }

  // Lightweight project-label augmentation. Kept minimal to avoid layout thrash
  // — this is intentionally a no-op if the page doesn't contain target fields.
  function addProjectLabels() {
    try {
      // Example: add a subtle label next to work_order fields for clarity
      const nodes = document.querySelectorAll('[data-fieldname="work_order"]');
      nodes.forEach((n) => {
        try {
          if (n && !n.dataset.agressoLabelAdded) {
            n.dataset.agressoLabelAdded = '1';
            // don't mutate heavy markup; just set a title attribute as gentle augmentation
            try { n.setAttribute('title', (n.getAttribute('title') || '') + ' (Work order)'); } catch (e) {}
          }
        } catch (e) {}
      });
    } catch (e) {}
  }

  // Central layout enhancer used during init and on layout refresh.
  function enhanceLayout() {
    try { applyFieldSizing(); } catch (e) {}
    try { positionIndicatorNearSaveButton(); } catch (e) {}
    try { addProjectLabels(); } catch (e) {}
  }

  const INDICATOR_ID = 'agresso-autosave-indicator';
  // When true, show extra debug controls (bell, debug, override/settings).
  // Toggle in code during development by setting to `true` and reloading the page.
  // You can also call `window.agresso_setIndicatorDebug(true)` in the console
  // after reloading to persist the flag for the session (no reload required).
  let INDICATOR_DEBUG = false;
  const OK_LABELS = ['ok', 'stäng', 'close', 'oké'];
  const CLOSE_SELECTORS = ['[aria-label="Close"]', '.close', '.k-i-close', '.modal-close'];
  const ACTIVITY_MESSAGE = 'agresso-autosave-activity';

  // Toggle persistence key
  const TOGGLE_KEY = 'agresso_autosave_enabled';

  function getToggleEnabled() {
    try {
      const v = localStorage.getItem(TOGGLE_KEY);
      if (v === null) return true;
      return v === '1' || v === 'true';
    } catch (e) {
      return true;
    }
  }

  function setToggleEnabled(enabled) {
    try {
      localStorage.setItem(TOGGLE_KEY, enabled ? '1' : '0');
    } catch (e) {
      // ignore
    }
    applyToggleState(enabled);
    // When toggled off, stop timers and prevent saves. When toggled on, resume watching.
    try {
      if (!enabled) {
        stopTimer();
        setIndicator('pending', 'Autosave disabled', 'Paused');
      } else {
        setIndicator('saved', 'Autosave ready', 'Watching for edits');
        // start idle timer anew
        try { startTimer(IDLE_TIMEOUT_MS, 'idle'); } catch (e) { /* ignore */ }
      }
    } catch (e) {
      // ignore
    }
  }

  function applyToggleState(enabled) {
    try {
      const doc = getIndicatorDocument();
      const ind = doc.getElementById(INDICATOR_ID);
      if (ind) {
        if (!enabled) {
          ind.classList.add('agresso-disabled');
        } else {
          ind.classList.remove('agresso-disabled');
        }
        if (enabled) {
          ind.classList.add('agresso-enabled');
        } else {
          ind.classList.remove('agresso-enabled');
        }
      }
      try {
        document.documentElement.setAttribute('data-agresso-autosave-enabled', enabled ? '1' : '0');
      } catch (e) {
        // ignore
      }
    } catch (e) {
      // ignore
    }
  }

  function getIndicatorDocument() {
    // Return the document where the indicator should be injected.
    // Prefer the top-level same-origin document when available so the
    // indicator is visible even when the script runs in a frame. Do not
    // call `ensureIndicator()` here to avoid recursion.
    try {
      try {
        if (window.top && window.top.document && window.top !== window) {
          return window.top.document;
        }
      } catch (e) {
        // access to top may be cross-origin
      }
    } catch (e) {}
    return document;
  }

  function scheduleLayoutRefresh() {
    if (layoutTimer) {
      clearTimeout(layoutTimer);
    }
    layoutTimer = window.setTimeout(() => {
      layoutTimer = null;
      enhanceLayout();
    }, LAYOUT_REFRESH_MS);
  }

  function ensureIndicator() {
    const indicatorDoc = getIndicatorDocument();
    let indicator = indicatorDoc.getElementById(INDICATOR_ID);
    if (indicator) {
      return indicator;
    }

    indicator = indicatorDoc.createElement('div');
    indicator.id = INDICATOR_ID;

    // Create a small on/off toggle inside the indicator (we'll place it first)
    let toggle = null;
    try {
      toggle = indicatorDoc.createElement('button');
      toggle.className = 'agresso-toggle';
      toggle.setAttribute('type', 'button');
      toggle.setAttribute('aria-pressed', String(getToggleEnabled()));
      toggle.title = 'Toggle autosave on / off';

      const sw = indicatorDoc.createElement('span');
      sw.className = 'switch';
      const knob = indicatorDoc.createElement('span');
      knob.className = 'knob';
      sw.appendChild(knob);
      toggle.appendChild(sw);

      toggle.addEventListener('click', (ev) => {
        ev.stopPropagation();
        ev.preventDefault();
        const cur = getToggleEnabled();
        const next = !cur;
        setToggleEnabled(next);
        try { toggle.setAttribute('aria-pressed', String(next)); } catch (e) {}
      }, true);
    } catch (e) {
      toggle = null;
    }

    // Small reminder bell button: click toggles reminder on/off, Shift+click cycles language
    let reminderBtn = null;
    if (INDICATOR_DEBUG) {
      try {
        reminderBtn = indicatorDoc.createElement('button');
        reminderBtn.className = 'agresso-reminder-btn';
        reminderBtn.setAttribute('type', 'button');
        reminderBtn.style.marginLeft = '6px';
        reminderBtn.style.fontSize = '14px';
        reminderBtn.style.lineHeight = '1';
        reminderBtn.style.padding = '2px 6px';
        reminderBtn.style.borderRadius = '4px';
        reminderBtn.style.border = 'none';
        reminderBtn.style.background = 'transparent';
        reminderBtn.textContent = '🔔';
        reminderBtn.addEventListener('click', (ev) => {
          ev.stopPropagation();
          ev.preventDefault();
          if (ev.shiftKey) {
            const next = cycleReminderLang();
            updateReminderButtonState(reminderBtn);
            try { setIndicator('pending', next === 'en' ? 'Reminder (en)' : 'Påminnelse (sv)', ''); } catch (e) {}
            return;
          }
          const cur = getReminderEnabled();
          setReminderEnabled(!cur);
          updateReminderButtonState(reminderBtn);
        }, true);
      } catch (e) {
        reminderBtn = null;
      }
    }

    // Debug button to log detection report to console (avoids CSP issues)
    let debugBtn = null;
    if (INDICATOR_DEBUG) {
      try {
        debugBtn = indicatorDoc.createElement('button');
        debugBtn.className = 'agresso-debug-btn';
        debugBtn.setAttribute('type', 'button');
        debugBtn.style.marginLeft = '6px';
        debugBtn.style.fontSize = '12px';
        debugBtn.style.lineHeight = '1';
        debugBtn.style.padding = '2px 6px';
        debugBtn.style.borderRadius = '4px';
        debugBtn.style.border = 'none';
        debugBtn.style.background = 'transparent';
        debugBtn.textContent = '🐞';
        debugBtn.title = 'Debug: print period-detection report to console';
        debugBtn.addEventListener('click', (ev) => {
          ev.stopPropagation(); ev.preventDefault();
          try {
            const report = buildDebugReport();
            console.log(LOG_PREFIX, 'Period detection report', report);
          } catch (e) {
            console.error(LOG_PREFIX, 'Debug report failed', e);
          }
        }, true);
      } catch (e) {
        debugBtn = null;
      }
    }

    

    // Settings / override panel (only in debug mode)
    let settingsBtn = null;
    let settingsPanel = null;
    if (INDICATOR_DEBUG) {
      try {
        settingsBtn = indicatorDoc.createElement('button');
        settingsBtn.className = 'agresso-settings-btn';
        settingsBtn.setAttribute('type', 'button');
        settingsBtn.style.marginLeft = '6px';
        settingsBtn.style.fontSize = '12px';
        settingsBtn.style.lineHeight = '1';
        settingsBtn.style.padding = '2px 6px';
        settingsBtn.style.borderRadius = '4px';
        settingsBtn.style.border = 'none';
        settingsBtn.style.background = 'transparent';
        settingsBtn.textContent = '⚙️';
        settingsBtn.title = 'Settings: set manual period end override';

        settingsPanel = indicatorDoc.createElement('div');
        settingsPanel.className = 'agresso-settings-panel';
        settingsPanel.style.position = 'fixed';
        settingsPanel.style.zIndex = '999999';
        settingsPanel.style.padding = '8px';
        settingsPanel.style.background = '#fff';
        settingsPanel.style.border = '1px solid #ccc';
        settingsPanel.style.borderRadius = '6px';
        settingsPanel.style.boxShadow = '0 2px 8px rgba(0,0,0,0.2)';
        settingsPanel.style.display = 'none';
        settingsPanel.innerHTML = '<div style="font-size:12px;margin-bottom:6px;">Manual period end (YYYY-MM-DD or DD/MM):</div>';

        const inp = indicatorDoc.createElement('input');
        inp.type = 'text';
        inp.placeholder = '2026-01-11 or 11/01';
        inp.style.width = '150px';
        inp.style.marginRight = '6px';
        settingsPanel.appendChild(inp);

        const saveBtn = indicatorDoc.createElement('button');
        saveBtn.textContent = 'Save';
        saveBtn.style.marginRight = '4px';
        settingsPanel.appendChild(saveBtn);

        const clearBtn = indicatorDoc.createElement('button');
        clearBtn.textContent = 'Clear';
        settingsPanel.appendChild(clearBtn);

        saveBtn.addEventListener('click', (ev) => {
          ev.stopPropagation(); ev.preventDefault();
          const v = inp.value && inp.value.trim();
          if (!v) return;
          setOverrideDate(v);
          try { setIndicator('pending', 'Override saved', v); } catch (e) {}
          // Request notification permission and trigger notification if override is today
          try {
            const parsed = parseDateFlexible(v);
            const today = new Date();
            const isToday = parsed && parsed.getFullYear && parsed.getFullYear() === today.getFullYear() && parsed.getMonth() === today.getMonth() && parsed.getDate() === today.getDate();
            if (typeof Notification !== 'undefined' && Notification.permission !== 'granted' && Notification.permission !== 'denied') {
              Notification.requestPermission().then((perm) => {
                if (perm === 'granted' && isToday) {
                  try { showPeriodNotification(parsed); } catch (e) {}
                }
              }).catch(() => {
                if (isToday) try { setIndicator('pending', 'Sista dagen i perioden', 'Skicka in tidrapport idag'); } catch (e) {}
              });
            } else {
              if (isToday) {
                try { showPeriodNotification(parsed); } catch (e) { try { setIndicator('pending', 'Sista dagen i perioden', 'Skicka in tidrapport idag'); } catch (e2) {} }
              }
            }
          } catch (e) {
            // ignore
          }
          settingsPanel.style.display = 'none';
        }, true);

        clearBtn.addEventListener('click', (ev) => {
          ev.stopPropagation(); ev.preventDefault();
          clearOverrideDate();
          try { setIndicator('saved', 'Override cleared', ''); } catch (e) {}
          settingsPanel.style.display = 'none';
        }, true);

        settingsBtn.addEventListener('click', (ev) => {
          ev.stopPropagation(); ev.preventDefault();
          if (settingsPanel.style.display === 'none') {
            // prefill with existing override if present
            try { const o = localStorage.getItem(PERIOD_OVERRIDE_KEY); if (o) inp.value = o; else inp.value = ''; } catch (e) { inp.value = ''; }
            const rect = settingsBtn.getBoundingClientRect();
            settingsPanel.style.left = `${Math.max(8, rect.left)}px`;
            settingsPanel.style.top = `${Math.max(8, rect.top - 80)}px`;
            settingsPanel.style.display = 'block';
          } else {
            settingsPanel.style.display = 'none';
          }
        }, true);

      } catch (e) {
        settingsBtn = null; settingsPanel = null;
      }
    }

    const dot = indicatorDoc.createElement('span');
    dot.className = 'agresso-autosave-dot';

    const label = indicatorDoc.createElement('span');
    label.className = 'agresso-autosave-label';

    const sub = indicatorDoc.createElement('span');
    sub.className = 'agresso-autosave-sub';

    // Append toggle first so it replaces the left-side dot visually
    if (toggle) indicator.appendChild(toggle);
    if (reminderBtn) indicator.appendChild(reminderBtn);
    if (debugBtn) indicator.appendChild(debugBtn);
    if (settingsBtn) indicator.appendChild(settingsBtn);
    indicator.appendChild(dot);
    indicator.appendChild(label);
    indicator.appendChild(sub);

    // Append indicator to document and apply saved toggle state
    indicatorDoc.body.appendChild(indicator);
    try { applyToggleState(getToggleEnabled()); } catch (e) {}
    try { updateReminderButtonState(reminderBtn); } catch (e) {}
    try { if (settingsPanel) indicatorDoc.body.appendChild(settingsPanel); } catch (e) {}
    return indicator;
  }

  function findPrimarySaveButton() {
    const docs = getAllDocuments();
    for (const doc of docs) {
      const candidates = Array.from(
        doc.querySelectorAll('button, input[type="button"], input[type="submit"], a')
      );
      const match = candidates.find((el) => {
        if (!isVisible(el)) {
          return false;
        }
        const text = (el.innerText || el.value || '').toLowerCase();
        const id = (el.id || '').toLowerCase();
        const title = (el.getAttribute('title') || '').toLowerCase();
        return text.includes('spara') || id === 'b$tblsyssave'.toLowerCase() || title.includes('alt+s');
      });
      if (match) {
        return match;
      }
    }
    return null;
  }

  function getViewportRect(el) {
    if (!el) {
      return null;
    }

    try {
      const rect = el.getBoundingClientRect();
      let top = rect.top;
      let left = rect.left;
      let width = rect.width;
      let height = rect.height;
      let win = el.ownerDocument?.defaultView;

      while (win && win.parent && win !== win.parent) {
        const frame = win.frameElement;
        if (!frame) {
          break;
        }
        const frameRect = frame.getBoundingClientRect();
        top += frameRect.top;
        left += frameRect.left;
        win = win.parent;
      }

      return { top, left, width, height, right: left + width, bottom: top + height };
    } catch (e) {
      return null;
    }
  }

  function positionIndicatorNearSaveButton() {
    const indicator = ensureIndicator();
    // Use the document that actually contains the indicator (may be top-level)
    const doc = (indicator && indicator.ownerDocument) || document;
    // Prefer a save button inside the same document as the indicator to
    // compute a stable anchor; fall back to global search if none found.
    let btn = null;
    try {
      for (const sel of SAVE_BUTTON_SELECTORS) {
        try {
          const el = doc.querySelector(sel);
          if (el && !el.disabled && isVisible(el)) { btn = el; break; }
        } catch (e) {}
      }
    } catch (e) {}
    if (!btn) btn = findPrimarySaveButton();
    if (!btn) {
      indicator.style.top = 'auto';
      indicator.style.bottom = '20px';
      indicator.style.right = '16px';
      indicator.style.left = 'auto';
      return;
    }
    const rect = getViewportRect(btn);
    if (!rect) {
      indicator.style.top = 'auto';
      indicator.style.bottom = '20px';
      indicator.style.right = '16px';
      indicator.style.left = 'auto';
      return;
    }
    const indHeight = indicator.offsetHeight || 34;
    const indWidth = indicator.offsetWidth || 180;
    const top = rect.top + (rect.height - indHeight) / 2;
    const viewportWidth = (doc.defaultView && doc.defaultView.innerWidth) || window.innerWidth || 0;
    const viewportHeight = (doc.defaultView && doc.defaultView.innerHeight) || window.innerHeight || 0;

    // Anchor vertically aligned with the save button, but place the indicator
    // on the right side of the page (inside the main content area if possible).
    try {
      const clampedTop = Math.max(8, Math.min(top, Math.max(8, viewportHeight - indHeight - 8)));
      // find a suitable content container to anchor inside
      let rightPos = 16; // default distance from viewport right edge
      try {
        const contentSelectors = ['main', '#content', '.container', '.page', '.u4-main', '.u4-content', '.k-grid', '.u4-body'];
        let chosen = null;
        let chosenW = 0;
        for (const sel of contentSelectors) {
          try {
            const el = doc.querySelector(sel);
            if (!el) continue;
            const r = el.getBoundingClientRect();
            if (r.width > chosenW && r.width < viewportWidth - 40) { chosen = r; chosenW = r.width; }
          } catch (e) {}
        }
        if (chosen) {
          const paddingFromContent = 12;
          const extraInset = 6; // give a few extra pixels from the content edge
          rightPos = Math.max(8, Math.min(viewportWidth - indWidth - 8, Math.round(viewportWidth - chosen.right + paddingFromContent + extraInset)));
        } else {
          // If no content container, try mirroring the save button X position
          const margin = 12;
          const extraInset = 6;
            try {
              if (rect && typeof rect.left === 'number') {
                // Mirror the button's horizontal center across the viewport center
                const btnCenter = rect.left + (rect.width || 0) / 2;
                const mirroredCenter = Math.round(viewportWidth - btnCenter);
                let desiredLeft = Math.round(mirroredCenter - indWidth / 2);
                // Clamp inside viewport with small margins
                desiredLeft = Math.max(8, Math.min(viewportWidth - indWidth - 8, desiredLeft));
                // Compute a right offset so the indicator's right edge is fixed
                let rightPosFromLeft = Math.round(viewportWidth - desiredLeft - indWidth);
                // Shift a few pixels to the left so it's not flush with the border
                const mirrorExtraShift = 30;
                let rightPos = Math.max(8, Math.min(viewportWidth - indWidth - 8, rightPosFromLeft + mirrorExtraShift));
                // Position using `right` so expansion grows to the left (right edge fixed)
                indicator.style.position = 'fixed';
                indicator.style.top = `${clampedTop}px`;
                indicator.style.right = `${rightPos}px`;
                indicator.style.left = 'auto';
                indicator.style.bottom = 'auto';
                // Ensure transforms/origins anchor to the right side
                try { indicator.style.transformOrigin = 'right center'; } catch (e) {}
                return;
              } else {
                rightPos = 16 + extraInset;
              }
            } catch (e) {
              rightPos = 16 + extraInset;
            }
        }
      } catch (e) {
        rightPos = 16;
      }

      indicator.style.position = 'fixed';
      indicator.style.top = `${clampedTop}px`;
      indicator.style.right = `${rightPos}px`;
      indicator.style.left = 'auto';
      indicator.style.bottom = 'auto';
    } catch (e) {
      // fallback: bottom-right corner
      indicator.style.position = 'fixed';
      indicator.style.top = 'auto';
      indicator.style.bottom = '20px';
      indicator.style.right = '16px';
      indicator.style.left = 'auto';
    }
  }

  function bindIndicatorTracking() {
    const attach = (win) => {
      if (!win || trackedWindows.has(win)) {
        return;
      }
      trackedWindows.add(win);
      try {
        win.addEventListener('resize', positionIndicatorNearSaveButton, true);
        win.addEventListener('scroll', positionIndicatorNearSaveButton, true);
      } catch (e) {
        // ignore cross-origin listeners
      }
    };

    attach(window);
    try {
      attach(window.top);
    } catch (e) {
      // ignore cross-origin top access
    }

    getAllDocuments().forEach((doc) => {
      try {
        attach(doc.defaultView);
      } catch (e) {
        // ignore frames we cannot access
      }
    });
  }

  function bindActivityListeners() {
    const attachListeners = (target) => {
      if (!target || !target.addEventListener) return;
      try {
        // Broad set of events to catch typing, clicks, touch, scroll and pointer movement
        const events = [
          'keydown',
          'keyup',
          'keypress',
          'input',
          'click',
          'pointerdown',
          'pointerup',
          'touchstart',
          'wheel',
          'scroll',
          'mousemove',
          'focusin',
          'focusout'
        ];
        events.forEach((evt) => target.addEventListener(evt, markActivity, true));
      } catch (e) {
        // ignore frames we cannot access
      }
    };

    const attachDoc = (doc) => {
      if (!doc || trackedActivityDocs.has(doc)) return;
      trackedActivityDocs.add(doc);
      // Attach to the Document itself
      attachListeners(doc);
      // Also attach to the Window if available
      try {
        if (doc.defaultView) attachListeners(doc.defaultView);
      } catch (e) {
        // ignore cross-origin
      }
    };

    // Attach to all reachable documents and the main document
    getAllDocuments().forEach(attachDoc);
    attachDoc(document);
  }

  function ensureTimerBar() {
    const indicator = ensureIndicator();
    let bar = indicator.querySelector('.agresso-autosave-timer');
    if (!bar) {
      bar = indicator.ownerDocument.createElement('div');
      bar.className = 'agresso-autosave-timer';
      // Ensure the bar has sensible sizing so transform-based animation is visible
      try {
        bar.style.display = 'block';
        // initialize at 0 width so JS-driven transitions animate reliably
        bar.style.width = '0px';
        bar.style.height = '6px';
        bar.style.background = '#22c55e';
        bar.style.transformOrigin = 'left';
      } catch (e) {}
      indicator.appendChild(bar);
    }
    timerBar = bar;
    return bar;
  }

  function getTimerRemainingMs() {
    try {
      if (timerStartedAt && timerDuration) {
        const elapsed = Date.now() - timerStartedAt;
        return Math.max(800, timerDuration - elapsed);
      }
    } catch (e) {}
    return timerDuration || IDLE_TIMEOUT_MS;
  }

  function resetTimerBar(durationMs) {
    const bar = ensureTimerBar();
    // If indicator is paused, keep the bar full and don't animate it
    try {
      const parentIndicator = bar.closest && bar.closest('#' + INDICATOR_ID);
      if (parentIndicator && parentIndicator.classList.contains('agresso-paused')) {
        try { bar.style.transition = 'none'; } catch (e) {}
        try { bar.style.width = '100%'; } catch (e) {}
        return;
      }
    } catch (e) {}
    // Disable any CSS animations/transforms that may conflict
    try { bar.style.animation = 'none'; } catch (e) {}
    try { bar.style.transform = 'none'; } catch (e) {}
    bar.style.transition = 'none';
    // initialize to 0px (not percent) so the pixel delta is definite
    bar.style.width = '0px';
    // force reflow
    // eslint-disable-next-line no-unused-expressions
    bar.offsetWidth;
    // Determine target width in pixels from the indicator/container so
    // percent-based computed widths don't interfere with transition.
    let targetPx = null;
    try {
      const parent = bar.parentElement || bar.ownerDocument.body;
      const pw = (parent && parent.clientWidth) || bar.offsetWidth || 0;
      targetPx = `${pw}px`;
    } catch (e) {
      targetPx = '100%';
    }
    // Use inline transition so it takes precedence and animates from 0px->targetPx
    try { bar.style.willChange = 'width'; } catch (e) {}
    bar.style.transition = `width ${durationMs}ms linear`;
    // Trigger the width change to start the animation
    bar.style.width = targetPx;
  }

  function stopTimerBar() {
    const bar = timerBar || ensureTimerBar();
    try { bar.style.animation = 'none'; } catch (e) {}
    // If indicator is paused, keep it full instead of collapsing
    try {
      const parentIndicator = bar.closest && bar.closest('#' + INDICATOR_ID);
      if (parentIndicator && parentIndicator.classList.contains('agresso-paused')) {
        try { bar.style.transition = 'none'; } catch (e) {}
        try { bar.style.width = '100%'; } catch (e) {}
        return;
      }
    } catch (e) {}
    bar.style.transition = 'none';
    bar.style.width = '0px';
  }

  function findSaveButton() {
    const docs = getAllDocuments();
    for (const doc of docs) {
      for (const selector of SAVE_BUTTON_SELECTORS) {
        const el = doc.querySelector(selector);
        if (el && !el.disabled && isVisible(el)) {
          return el;
        }
      }
    }
    return null;
  }

  function setIndicator(state, labelText, subText) {
    const indicator = ensureIndicator();
    // Preserve whether the indicator currently has the period-end marker
    const hadPeriodMarker = indicator.classList.contains('agresso-period-end');

    indicator.classList.remove(
      'agresso-saving',
      'agresso-saved',
      'agresso-pending',
      'agresso-error'
    );
    indicator.classList.add(`agresso-${state}`);

    const label = indicator.querySelector('.agresso-autosave-label');
    const sub = indicator.querySelector('.agresso-autosave-sub');
    if (label) {
      label.textContent = labelText;
    }
    if (sub) {
      sub.textContent = subText;
    }

    // If label/subtext indicate autosave is paused, mark indicator so
    // timer-bar logic can keep the bar full and ignore activity.
    try {
      const hint = ((labelText || '') + ' ' + (subText || '')).toLowerCase();
      if (hint.indexOf('autosave paused') >= 0 || hint.indexOf('inga ändringar gjorda') >= 0 || hint.indexOf('autosave disabled') >= 0) {
        try { indicator.classList.add('agresso-paused'); } catch (e) {}
        try {
          const bar = indicator.querySelector('.agresso-autosave-timer');
          if (bar) { bar.style.transition = 'none'; bar.style.width = '100%'; }
        } catch (e) {}
      } else {
        try { indicator.classList.remove('agresso-paused'); } catch (e) {}
      }
    } catch (e) {}

    positionIndicatorNearSaveButton();
    // If we just reached a saved state, clear any period-end highlight
    try {
      if (state === 'saved') {
        // Only clear the period-end reminder when the report Status is set to 'Klar'
        try {
          const isKlar = isReportStatusKlar();
          if (isKlar) {
            try { indicator.classList.remove('agresso-period-end'); } catch (e) {}
            try {
              const bar = indicator.querySelector('.agresso-autosave-timer');
              if (bar) {
                try { bar.classList.remove('agresso-period-moving'); } catch (e) {}
                bar.style.background = '#22c55e';
                bar.style.boxShadow = 'none';
              }
            } catch (e) {}
            try { localStorage.removeItem(PERIOD_NOTIFY_KEY); } catch (e) {}
            try { if (periodStatusRefreshTimer) { clearInterval(periodStatusRefreshTimer); periodStatusRefreshTimer = null; } } catch (e) {}
              try { if (periodHighlightEnforcer) { clearInterval(periodHighlightEnforcer); periodHighlightEnforcer = null; } } catch (e) {}
            try { highlightStatusField(false); } catch (e) {}
            // Remove any persistent submit banners
            try { const b = document.getElementById('agresso-period-banner'); if (b && b.parentNode) b.parentNode.removeChild(b); } catch (e) {}
            try { if (window.top && window.top.document && window.top !== window) { const bt = window.top.document.getElementById('agresso-period-banner'); if (bt && bt.parentNode) bt.parentNode.removeChild(bt); } } catch (e) {}
          } else {
            // If we were displaying the period marker before this state change,
            // reapply it unless the report is confirmed 'Klar'. This avoids the
            // page briefly removing our visual reminder during normal UI updates.
            try {
              if (hadPeriodMarker) {
                try { indicator.classList.add('agresso-period-end'); } catch (e) {}
                try {
                  const bar = indicator.querySelector('.agresso-autosave-timer');
                  if (bar) {
                    try { bar.classList.add('agresso-period-moving'); try { resetTimerBar(getTimerRemainingMs()); } catch (e2) {} } catch (e) {}
                    bar.style.background = '#d9534f';
                    bar.style.boxShadow = '0 0 6px rgba(217,83,79,0.6)';
                  }
                } catch (e) {}
                try { indicator.style.border = '2px solid rgba(217,83,79,0.9)'; } catch (e) {}
                try { indicator.style.background = 'linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95))'; } catch (e) {}
              }
            } catch (e) {
              // ignore
            }
          }
        } catch (e) {
          // ignore
        }
      }
    } catch (e) {
      // ignore
    }

    // If the indicator is currently marked as period-end, ensure the subtext
    // clearly instructs the user to submit the timereport.
    try {
      const subEl = indicator.querySelector('.agresso-autosave-sub');
      if (indicator.classList.contains('agresso-period-end')) {
        if (subEl) subEl.textContent = '- Submit time report!';
      }
    } catch (e) {
      // ignore
    }
  }

  function isReportStatusKlar() {
    try {
      const docs = getAllDocuments();
      for (const doc of docs) {
        try {
          // First try: hidden RowDescription / RowValue inputs which often
          // accompany datalist controls. RowDescription typically contains
          // the human-readable text like 'Klar'.
          try {
            const desc = doc.querySelector('input[id$="RowDescription"], input[name$="RowDescription"], input[id*="RowDescription"], input[name*="RowDescription"]');
            if (desc && (desc.value || desc.getAttribute('value'))) {
              const v = (desc.value || desc.getAttribute('value') || '').toString().trim().toLowerCase();
              if (v.indexOf('klar') >= 0) return true;
            }
          } catch (e) {}

          try {
            const valIn = doc.querySelector('input[id$="RowValue"], input[name$="RowValue"], input[id*="RowValue"], input[name*="RowValue"]');
            if (valIn && (valIn.value || valIn.getAttribute('value'))) {
              const vv = (valIn.value || valIn.getAttribute('value') || '').toString().trim().toLowerCase();
              // Some implementations use 'N' to indicate the selected item
              if (vv === 'n' || vv.indexOf('klar') >= 0) return true;
            }
          } catch (e) {}

          // Next, try datalistcontrol elements that have title='Status' and
          // contain an input with the human-readable description.
          try {
            const dl = Array.from(doc.querySelectorAll('datalistcontrol')).find(d => {
              try { return (d.getAttribute('title') || '').toLowerCase().indexOf('status') >= 0; } catch (e) { return false; }
            });
            if (dl) {
              try {
                const inner = dl.querySelector('input');
                if (inner && (inner.value || inner.getAttribute('value'))) {
                  const v = (inner.value || inner.getAttribute('value') || '').toString().trim().toLowerCase();
                  if (v.indexOf('klar') >= 0) return true;
                }
              } catch (e) {}
            }
          } catch (e) {}

          // Fallback: look for visible label/text mentioning 'Status' and check
          // nearby select/input values.
          try {
            const labelNode = Array.from(doc.querySelectorAll('label, th, td, div, span'))
              .find(n => /\bstatus\b/i.test((n.textContent||'').trim()));
            if (labelNode) {
              const container = labelNode.closest('tr') || labelNode.parentElement || doc;
              const input = container.querySelector('select, input[type="text"], input');
              if (input) {
                const val = (input.value || (input.selectedOptions && input.selectedOptions[0] && input.selectedOptions[0].text) || '').toString().trim().toLowerCase();
                if (val.indexOf('klar') >= 0) return true;
              }
            }
          } catch (e) {}

          // Final fallback: check any select's selected option text for 'Klar'
          const selects = Array.from(doc.querySelectorAll('select'));
          for (const s of selects) {
            try {
              const selText = (s.selectedOptions && s.selectedOptions[0] && s.selectedOptions[0].text) || (s.options && s.options[s.selectedIndex] && s.options[s.selectedIndex].text) || '';
              if ((selText||'').toLowerCase().indexOf('klar') >= 0) return true;
            } catch (e) {}
          }
        } catch (e) {
          // ignore per-document errors
        }
      }
    } catch (e) {
      // ignore
    }
    return false;
  }

  function getReportStatusText() {
    try {
      const docs = getAllDocuments();
      for (const doc of docs) {
        try {
          const labelNode = Array.from(doc.querySelectorAll('label, th, td, div, span'))
            .find(n => /\bstatus\b/i.test((n.textContent||'').trim()));
          if (labelNode) {
            const container = labelNode.closest('tr') || labelNode.parentElement || doc;
            const input = container.querySelector('select, input[type="text"], input');
            if (input) {
              // selected option text or input value
              const selText = (input.selectedOptions && input.selectedOptions[0] && input.selectedOptions[0].text) || input.value || '';
              if (selText) return (selText || '').trim();
            }
          }

          const selects = Array.from(doc.querySelectorAll('select'));
          for (const s of selects) {
            try {
              const selText = (s.selectedOptions && s.selectedOptions[0] && s.selectedOptions[0].text) || (s.options && s.options[s.selectedIndex] && s.options[s.selectedIndex].text) || '';
              if (selText) return (selText || '').trim();
            } catch (e) {}
          }
        } catch (e) {
          // ignore per-document errors
        }
      }
    } catch (e) {
      // ignore
    }
    return null;
  }

  function refreshPeriodIndicatorStatus() {
    try {
      const indicator = ensureIndicator();
      if (!indicator.classList.contains('agresso-period-end')) return;
      const lang = getReminderLang();
      const base = lang === 'en' ? 'Today is the last day of the period — submit your time report.' : 'Idag är sista dagen för perioden – skicka in din tidrapport.';
      const statusText = getReportStatusText();
      const sub = statusText ? `${base} • Status: ${statusText}` : base;
      try {
        const subEl = indicator.querySelector('.agresso-autosave-sub');
        if (subEl) subEl.textContent = sub;
      } catch (e) {}
    } catch (e) {
      // ignore
    }
  }

  function highlightStatusField(highlight) {
    try {
      const docs = getAllDocuments();
      for (const doc of docs) {
        try {
          const labelNode = Array.from(doc.querySelectorAll('label, th, td, div, span'))
            .find(n => /\bstatus\b/i.test((n.textContent||'').trim()));
          if (!labelNode) continue;
          const container = labelNode.closest('tr') || labelNode.parentElement || doc;
          // Try to find a status input or readable text in the same container
          const input = container.querySelector('select, input[type="text"], input');
          let val = null;
          try {
            if (input) {
              const tag = (input.tagName || '').toLowerCase();
              if (tag === 'select') {
                val = (input.selectedOptions && input.selectedOptions[0] && input.selectedOptions[0].text) || null;
              } else {
                val = (input.value || input.textContent || null) || null;
              }
            }
          } catch (e) {
            val = null;
          }

          // If no form control found, try to read textual status from nearby cells
          if (val === null || val === '') {
            try {
              // look for a span/div/td inside the container that contains status text
              const textNode = container.querySelector('span, div, td, strong, b, em');
              if (textNode) {
                const t = (textNode.textContent || '').trim();
                if (t) val = t;
              }
            } catch (e) {
              // ignore
            }
          }

          // If we still couldn't determine a value, do not apply highlight
          if (val === null || (typeof val === 'string' && val.trim() === '')) {
            if (!highlight) {
              try { if (input) { input.style.boxShadow = ''; input.style.background = ''; } } catch (e) {}
            }
            return false;
          }

          const isKlar = ('' + val).toLowerCase().indexOf('klar') >= 0;
          if (highlight) {
            if (!isKlar) {
              try {
                if (input) { input.style.boxShadow = '0 0 8px rgba(217,83,79,0.65)'; input.style.background = '#fff7f7'; }
              } catch (e) {}
            } else {
              try { if (input) { input.style.boxShadow = ''; input.style.background = ''; } } catch (e) {}
            }
          } else {
            try { if (input) { input.style.boxShadow = ''; input.style.background = ''; } } catch (e) {}
          }
          return true;
        } catch (e) {
          // ignore per-doc
        }
      }
    } catch (e) {}
    return false;
  }

  function isDeletionButton(target) {
    if (!(target instanceof HTMLElement)) {
      return false;
    }
    const text = (target.textContent || target.innerText || target.getAttribute('value') || '').toLowerCase();
    return DELETE_KEYWORDS.some((kw) => text.includes(kw));
  }

  function isAddRowButton(target) {
    if (!(target instanceof HTMLElement)) {
      return false;
    }
    const text = (target.textContent || target.innerText || target.getAttribute('value') || '').toLowerCase();
    return ADD_KEYWORDS.some((kw) => text.includes(kw));
  }

  function isNoChangesBannerVisible() {
    const docs = getAllDocuments();
    return docs.some((doc) => {
      try {
        const body = doc.body;
        if (!body) {
          return false;
        }
        const text = (body.innerText || '').toLowerCase();
        return text.includes(NO_CHANGES_TEXT) || text.includes(NO_CHANGES_TEXT_ALT);
      } catch (e) {
        return false;
      }
    });
  }

  function refreshNoChangesBannerState(reason) {
    const visible = isNoChangesBannerVisible();
    if (visible === noChangesBannerVisible) {
      return;
    }

    noChangesBannerVisible = visible;
    console.info(LOG_PREFIX, 'No-changes banner state', { visible, reason });

    if (visible) {
      stopTimer();
      setIndicator('pending', 'Autosave paused', '');
    } else {
      const row = pendingRow || getDirtyRow();
      if (row) {
        setIndicator('pending', 'Autosave ready', 'Watching for edits');
        startTimer(IDLE_TIMEOUT_MS, 'idle');
      } else {
        setIndicator('saved', 'Autosave ready', 'Watching for edits');
        startTimer(IDLE_TIMEOUT_MS, 'idle');
      }
    }
  }

  function stopTimer() {
    if (unifiedTimer) {
      clearTimeout(unifiedTimer);
      unifiedTimer = null;
    }
    stopTimerBar();
  }

  function startTimer(durationMs, reason) {
    stopTimer();
    
    timerStartedAt = Date.now();
    timerDuration = durationMs;
    timerReason = reason;
    
    console.debug(LOG_PREFIX, 'Timer started', { durationMs, reason });
    resetTimerBar(durationMs);
    
    unifiedTimer = window.setTimeout(() => {
      unifiedTimer = null;
      onTimerComplete();
    }, durationMs);
  }

  function onTimerComplete() {
    console.debug(LOG_PREFIX, 'Timer completed', { reason: timerReason });
    // If autosave is disabled, stop and don't proceed with save
    try {
      if (!getToggleEnabled()) {
        console.debug(LOG_PREFIX, 'Timer completed but autosave disabled');
        stopTimer();
        setIndicator('pending', 'Autosave disabled', 'Paused');
        return;
      }
    } catch (e) {
      // ignore
    }
    
    // Check if we should actually save
    if (dropdownActive) {
      console.debug(LOG_PREFIX, 'Save skipped - dropdown active');
      startTimer(IDLE_RETRY_MS, 'retry-dropdown');
      return;
    }
    
    if (noChangesBannerVisible) {
      console.debug(LOG_PREFIX, 'Save skipped - no changes banner visible');
      stopTimer();
      setIndicator('pending', 'Autosave paused', '');
      return;
    }
    
    // Perform the save - let the save function handle the outcome
    performSave(timerReason);
  }

  function markActivity() {
    // If we're in a frame (not the top window), forward activity to the top window
    try {
      if (window.top && window.top !== window) {
        // If the top-level indicator is paused, don't forward or restart
        try {
          const topDoc = window.top.document;
          const topInd = topDoc && topDoc.getElementById && topDoc.getElementById(INDICATOR_ID);
          if (topInd && topInd.classList && topInd.classList.contains('agresso-paused')) {
            return;
          }
        } catch (e) {
          // cross-origin may throw; ignore
        }
        window.top.postMessage({ type: ACTIVITY_MESSAGE, ts: Date.now() }, '*');
        return;
      }
    } catch (e) {
      // ignore cross-origin access; fall through to local handling
    }

    // If local/top indicator is marked paused, don't restart timer on movement/clicks
    try {
      const doc = getIndicatorDocument();
      const ind = doc && doc.getElementById && doc.getElementById(INDICATOR_ID);
      if (ind && ind.classList && ind.classList.contains('agresso-paused')) return;
    } catch (e) {
      // ignore
    }

    lastActivityAt = Date.now();
    // Restart idle countdown on activity
    startTimer(IDLE_TIMEOUT_MS, 'idle');
  }

  function performSave(trigger) {
    // Respect toggle: skip saves when disabled
    try {
      if (!getToggleEnabled()) {
        console.info(LOG_PREFIX, 'Autosave disabled, skipping save');
        setIndicator('pending', 'Autosave disabled', 'Paused');
        return;
      }
    } catch (e) {
      // ignore
    }
    // Only perform save from the top frame
    try {
      if (window.top && window.top !== window) return;
    } catch (e) {
      // if cross-origin error, bail
      return;
    }

    console.info(LOG_PREFIX, 'Saving via shortcut', { trigger, shortcut: SHORTCUT_LABEL });
    setIndicator('saving', 'Saving…', `Using ${SHORTCUT_LABEL}`);
    triggerShortcutSave();

    // Note: removed fallback button-click logic to avoid CSP/page errors.

    lastSaveAt = Date.now();
    startDialogSweep('autosave');

    window.setTimeout(() => {
      const timestamp = new Date().toLocaleTimeString(undefined, {
        hour12: false,
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      });
      try {
        // If the indicator (or a no-changes banner) indicates paused state
        // after save, preserve the paused state and do not restart timer.
        const doc = getIndicatorDocument();
        const ind = doc && doc.getElementById ? doc.getElementById(INDICATOR_ID) : null;
        const pausedNow = (ind && ind.classList && ind.classList.contains('agresso-paused')) || noChangesBannerVisible || isNoChangesBannerVisible();
        if (pausedNow) {
          try { setIndicator('pending', 'Autosave paused', ''); } catch (e) {}
          try { stopTimer(); } catch (e) {}
        } else {
          setIndicator('saved', 'Saved', `at ${timestamp}`);
          // Restart idle countdown after a save completes
          lastActivityAt = Date.now();
          startTimer(IDLE_TIMEOUT_MS, 'idle');
        }
      } catch (e) {
        // Fallback: behave as saved
        try { setIndicator('saved', 'Saved', `at ${timestamp}`); } catch (e2) {}
        try { lastActivityAt = Date.now(); startTimer(IDLE_TIMEOUT_MS, 'idle'); } catch (e2) {}
      }
      if (pendingRow) {
        pendingRow.dataset.agressoDirty = '0';
        pendingRow = null;
      }
    }, 1500);
    // After save completes, give the page a short moment and then re-check
    // for the "no changes" banner. If present, follow existing procedure
    // (refreshNoChangesBannerState will pause/stop timers as needed).
    window.setTimeout(() => {
      try { refreshNoChangesBannerState('post-save-check'); } catch (e) {}
    }, 800);
  }

  function getDirtyRow() {
    return document.querySelector('tr[data-agresso-dirty="1"]');
  }

  function triggerShortcutSave() {
    const docs = getAllDocuments();
    const targets = new Set();

    docs.forEach((doc) => {
      if (!doc) {
        return;
      }
      targets.add(doc.activeElement);
      targets.add(doc.body);
      targets.add(doc);
      try {
        if (doc.defaultView) {
          targets.add(doc.defaultView);
        }
      } catch (e) {
        // ignore
      }
    });

    SHORTCUT_COMBOS.forEach((combo) => {
      const base = {
        key: 's',
        code: 'KeyS',
        altKey: combo.altKey,
        metaKey: combo.metaKey,
        bubbles: true,
        cancelable: true,
        composed: true,
        keyCode: 83,
        which: 83
      };

      const events = [
        new KeyboardEvent('keydown', base),
        new KeyboardEvent('keypress', base),
        new KeyboardEvent('keyup', base)
      ];

      targets.forEach((el) => {
        if (el && typeof el.dispatchEvent === 'function') {
          events.forEach((evt) => el.dispatchEvent(evt));
        }
      });
    });

    console.info(LOG_PREFIX, 'Shortcut events dispatched', { shortcut: SHORTCUT_LABEL, targets: targets.size, combos: SHORTCUT_COMBOS.length });
  }

  // Fallback save click logic removed (caused errors on some pages/CSP).

  function isVisible(el) {
    if (!el) {
      return false;
    }
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  function findDialogButton(allowHidden = false) {
    const docs = getAllDocuments();
    for (const doc of docs) {
      try {
        const buttons = Array.from(
          doc.querySelectorAll('button, input[type="button"], input[type="submit"], a')
        );

        const okButton = buttons.find((btn) => {
          const text = (btn.textContent || btn.value || '').toLowerCase().trim();
          return OK_LABELS.some((label) => text === label || text === `${label}.` || text === `[${label}]`);
        });
        if (okButton && (allowHidden || isVisible(okButton))) {
          return okButton;
        }

        for (const selector of CLOSE_SELECTORS) {
          const el = doc.querySelector(selector);
          if (el && (allowHidden || isVisible(el))) {
            return el;
          }
        }
      } catch (e) {
        // ignore cross-origin issues
      }
    }

    return null;
  }

  function findDialogContainer(el) {
    if (!(el instanceof HTMLElement)) {
      return null;
    }
    const selectors = [
      '[role="dialog"]',
      '.modal',
      '.k-window',
      '.notification',
      '.alert',
      '.k-dialog',
      ...SAVE_DIALOG_SELECTORS
    ];
    if (el.matches(selectors.join(','))) {
      return el;
    }
    return el.closest(selectors.join(','));
  }

  function isSaveDialog(dialog) {
    if (!dialog) {
      return false;
    }
    if (SAVE_DIALOG_SELECTORS.some((sel) => dialog.matches(sel))) {
      return true;
    }
    const text = (dialog.innerText || dialog.textContent || '').toLowerCase();
    return SAVE_DIALOG_KEYWORDS.some((kw) => text.includes(kw));
  }

  function sweepDialogs(reason) {
    // If autosave is disabled, do not attempt to sweep/dismiss dialogs
    try {
      if (!getToggleEnabled()) {
        return false;
      }
    } catch (e) {
      // ignore
    }
    // Only sweep dialogs from the top frame
    try {
      if (window.top && window.top !== window) return false;
    } catch (e) {
      return false;
    }

    // First ensure there's any dialog-like element that looks like a save/confirmation
    const docs = getAllDocuments();
    let sawSaveDialogCandidate = false;
    const dialogSelectors = ['[role="dialog"]', '.modal', '.k-window', '.notification', '.alert', '.k-dialog', ...SAVE_DIALOG_SELECTORS];
    for (const doc of docs) {
      try {
        const candidate = doc.querySelector(dialogSelectors.join(','));
        if (!candidate) {
          continue;
        }

        // If candidate explicitly looks like a save dialog, proceed
        if (isSaveDialog(candidate)) {
          sawSaveDialogCandidate = true;
          break;
        }

        // Otherwise scan candidate text for save keywords
        const txt = (candidate.innerText || candidate.textContent || '').toLowerCase();
        if (SAVE_DIALOG_KEYWORDS.some((kw) => txt.includes(kw))) {
          sawSaveDialogCandidate = true;
          break;
        }
      } catch (e) {
        // ignore cross-origin issues
      }
    }

    // If nothing looks like a save dialog, silently skip without logging.
    if (!sawSaveDialogCandidate) {
      return false;
    }

    // We found a dialog-like candidate; look for a dismiss/OK button.
    // If the user has left the window (document hidden or not focused) allow
    // finding buttons even if they are not visible so background popups get dismissed.
    let allowHidden = false;
    try {
      const docsVisHidden = docs.some((d) => {
        try {
          return d.visibilityState === 'hidden' || d.hidden;
        } catch (e) {
          return false;
        }
      });
      allowHidden = (typeof document.hidden !== 'undefined' && document.hidden) || !document.hasFocus() || docsVisHidden;
    } catch (e) {
      // ignore
    }

    const button = findDialogButton(allowHidden);
    if (button) {
      const dialog = findDialogContainer(button);
      if (!dialog || !isSaveDialog(dialog)) {
        console.info(LOG_PREFIX, 'Dialog ignored (non-save)', { reason });
        return false;
      }

      // Hide dialog and any backdrop/overlay immediately
      try {
        dialog.style.display = 'none';
        dialog.style.visibility = 'hidden';
        dialog.style.opacity = '0';
        dialog.style.pointerEvents = 'none';
      } catch (e) {
        // ignore
      }

      // Also hide any backdrop/overlay elements
      docs.forEach((doc) => {
        try {
          const overlays = doc.querySelectorAll('.k-overlay, .modal-backdrop, [class*="overlay"], [class*="backdrop"]');
          overlays.forEach((overlay) => {
            try {
              overlay.style.display = 'none';
              overlay.style.visibility = 'hidden';
              overlay.style.opacity = '0';
            } catch (e) {
              // ignore
            }
          });
        } catch (e) {
          // ignore
        }
      });

      // Click button to dismiss
      try {
        button.click();
      } catch (e) {
        // ignore
      }
      console.info(LOG_PREFIX, 'Dialog dismissed', { reason });
      return true;
    }

    // Only log a warning if we haven't already when a save-like dialog was present
    if (!dialogMissLogged) {
      console.warn(LOG_PREFIX, 'Dialog button not found', { reason });
      dialogMissLogged = true;
    }
    return false;
  }

  function startDialogSweep(reason) {
    // Respect toggle: don't start sweeping dialogs when autosave is disabled
    try {
      if (!getToggleEnabled()) return;
    } catch (e) {
      // ignore
    }
    // Extend sweep time if already running
    dialogSweepEndAt = Date.now() + DIALOG_SWEEP_MS;
    
    if (dialogSweepTimer) {
      console.debug(LOG_PREFIX, 'Extending dialog sweep', { reason });
      return;
    }

    dialogMissLogged = false;

    dialogSweepTimer = window.setInterval(() => {
      if (Date.now() > dialogSweepEndAt) {
        window.clearInterval(dialogSweepTimer);
        dialogSweepTimer = null;
        return;
      }

      sweepDialogs(reason);
    }, DIALOG_SWEEP_INTERVAL_MS);
    console.debug(LOG_PREFIX, 'Started dialog sweep', { reason, durationMs: DIALOG_SWEEP_MS });
  }

  // --- Period end detection and notification ---
  const PERIOD_NOTIFY_KEY = 'agresso_period_notify_date';
  const REMINDER_ENABLED_KEY = 'agresso_period_notify_enabled';
  const REMINDER_LANG_KEY = 'agresso_period_notify_lang';
  const PERIOD_OVERRIDE_KEY = 'agresso_period_override';

  function getReminderEnabled() {
    try {
      const v = localStorage.getItem(REMINDER_ENABLED_KEY);
      if (v === null) return true;
      return v === '1' || v === 'true';
    } catch (e) {
      return true;
    }
  }

  function setReminderEnabled(enabled) {
    try { localStorage.setItem(REMINDER_ENABLED_KEY, enabled ? '1' : '0'); } catch (e) {}
  }

  function getReminderLang() {
    try { return localStorage.getItem(REMINDER_LANG_KEY) || 'sv'; } catch (e) { return 'sv'; }
  }

  function setReminderLang(lang) {
    try { localStorage.setItem(REMINDER_LANG_KEY, String(lang)); } catch (e) {}
  }

  function cycleReminderLang() {
    const cur = getReminderLang();
    const next = cur === 'en' ? 'sv' : 'en';
    setReminderLang(next);
    return next;
  }

  function getOverrideDate() {
    try {
      const v = localStorage.getItem(PERIOD_OVERRIDE_KEY);
      if (!v) return null;
      const d = parseDateFlexible(v);
      try { console.info(LOG_PREFIX, 'getOverrideDate: raw override', v, 'parsed', d); } catch (e) {}
      return d;
    } catch (e) { return null; }
  }

  function setOverrideDate(v) {
    try { localStorage.setItem(PERIOD_OVERRIDE_KEY, String(v)); } catch (e) {}
  }

  function clearOverrideDate() {
    try { localStorage.removeItem(PERIOD_OVERRIDE_KEY); } catch (e) {}
  }

  function updateReminderButtonState(btn) {
    try {
      if (!btn) return;
      const enabled = getReminderEnabled();
      const lang = getReminderLang();
      try { btn.setAttribute('aria-pressed', String(enabled)); } catch (e) {}
      btn.title = enabled ? (lang === 'en' ? 'Reminder: On (en) - Shift+click to switch language' : 'Påminnelse: På (sv) - Shift+click för språk') : (lang === 'en' ? 'Reminder: Off - click to enable' : 'Påminnelse: Av - klicka för att aktivera');
      btn.style.opacity = enabled ? '1' : '0.45';
    } catch (e) {
      // ignore
    }
  }

  function parseDateFlexible(s) {
    if (!s) return null;
    s = s.trim();
    // ISO yyyy-mm-dd
    const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (iso) {
      return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    }

    // dd/mm/yyyy or dd.mm.yyyy or dd mm yyyy
    const parts = /^(\d{1,2})[\/\.\s](\d{1,2})[\/\.\s](\d{2,4})$/.exec(s);
    if (parts) {
      let day = Number(parts[1]);
      let month = Number(parts[2]);
      let year = Number(parts[3]);
      if (year < 100) year += 2000;
      return new Date(year, month - 1, day);
    }

    // dd/mm or dd.mm (no year) -> assume current year
    const twoPart = /^(\d{1,2})[\/\.\s](\d{1,2})$/.exec(s);
    if (twoPart) {
      const day = Number(twoPart[1]);
      const month = Number(twoPart[2]);
      const year = (new Date()).getFullYear();
      return new Date(year, month - 1, day);
    }

    // Try Date.parse fallback (e.g., "1 January 2025")
    const parsed = Date.parse(s);
    if (!Number.isNaN(parsed)) {
      return new Date(parsed);
    }
    return null;
  }

  function findPeriodEndDate() {
    try {
      console.info(LOG_PREFIX, 'findPeriodEndDate: start scan');
      // (Removed deterministic "direct div match" scan to avoid static div identification)
      const indicatorEl = document.getElementById('agresso-autosave-indicator');
      let nodes = Array.from(document.querySelectorAll('h1,h2,h3,p,div,span,label,td,th'));
      if (indicatorEl) {
        nodes = nodes.filter(n => !indicatorEl.contains(n));
      }
      const dateRangeIso = /(\d{4}-\d{2}-\d{2})\s*[–—-]\s*(\d{4}-\d{2}-\d{2})/;
      const dateRangeSlashed = /(\d{1,2}[\/\.\s]\d{1,2}[\/\.\s]\d{2,4})\s*[–—-]\s*(\d{1,2}[\/\.\s]\d{1,2}[\/\.\s]\d{2,4})/;
      const monthNameRange = /(\d{1,2}\s+[A-Za-zåäöÅÄÖ]+\s+\d{4})\s*[–—-]\s*(\d{1,2}\s+[A-Za-zåäöÅÄÖ]+\s+\d{4})/;

      // Prefer elements that mention 'period' or similar
      const priority = nodes.filter((n) => /period|perioden|vecka|veckor|tidrapport/i.test((n.textContent||'')));
      const searchList = priority.length ? priority : nodes;

      for (const el of searchList) {
        if (indicatorEl && indicatorEl.contains(el)) continue;
        const txt = (el.textContent || '').trim();
        let m = dateRangeIso.exec(txt);
        if (m) {
          console.info(LOG_PREFIX, 'findPeriodEndDate: matched iso range in element', txt.slice(0,200));
          try { console.info(LOG_PREFIX, 'findPeriodEndDate: matched element outer', (el.outerHTML||'').slice(0,200)); } catch (e) {}
          return parseDateFlexible(m[2]);
        }
        m = dateRangeSlashed.exec(txt);
        if (m) {
          console.info(LOG_PREFIX, 'findPeriodEndDate: matched slashed range in element', txt.slice(0,200));
          try { console.info(LOG_PREFIX, 'findPeriodEndDate: matched element outer', (el.outerHTML||'').slice(0,200)); } catch (e) {}
          return parseDateFlexible(m[2]);
        }
        m = monthNameRange.exec(txt);
        if (m) {
          console.info(LOG_PREFIX, 'findPeriodEndDate: matched month name range in element', txt.slice(0,200));
          try { console.info(LOG_PREFIX, 'findPeriodEndDate: matched element outer', (el.outerHTML||'').slice(0,200)); } catch (e) {}
          return parseDateFlexible(m[2]);
        }
      }

      // Deterministic header `th` scan: look for th elements containing
      // DivOverflowNoWrap header DIVs with date tokens (e.g. "Fre<br>09/01").
      try {
        const ths = Array.from(document.querySelectorAll('th'));
        if (ths.length) {
          const headerDates = [];
          const dateTokenSimple = /\b(\d{1,2})[\/\.](\d{1,2})\b/;
          const isRedColorLocal = (colorStr) => {
            if (!colorStr) return false;
            const rgb = /rgb\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)/i.exec(colorStr);
            if (rgb) {
              const r = Number(rgb[1]), g = Number(rgb[2]), b = Number(rgb[3]);
              return r > 140 && r > g + 30 && r > b + 30;
            }
            const hex = /#([0-9a-f]{6}|[0-9a-f]{3})/i.exec(colorStr);
            if (hex) {
              let h = hex[1]; if (h.length === 3) h = h.split('').map(c=>c+c).join('');
              const r = parseInt(h.slice(0,2),16), g = parseInt(h.slice(2,4),16), b = parseInt(h.slice(4,6),16);
              return r > 140 && r > g + 30 && r > b + 30;
            }
            return false;
          };

          ths.forEach((th) => {
            try {
              const div = th.querySelector && th.querySelector('.DivOverflowNoWrap, .Ellipsis, .Separator');
              if (!div) return;
              const raw = ((div.dataset && div.dataset.originaltext) || div.getAttribute('title') || div.textContent || div.innerHTML || '').toString().replace(/<br\s*\/?>(\s*)/gi,' ').trim();
              const m = dateTokenSimple.exec(raw);
              if (m) {
                // read computed color if possible
                let color = '';
                try { color = (window.getComputedStyle && window.getComputedStyle(div).color) || div.style && div.style.color || ''; } catch (e) {}
                headerDates.push({ th, day: Number(m[1]), month: Number(m[2]), raw, color, idx: th.cellIndex });
              }
            } catch (e) {}
          });

          if (headerDates.length) {
            // choose the rightmost non-Sum header (closest to Sum on the left)
            headerDates.sort((a,b) => (a.idx || 0) - (b.idx || 0));
            // find index of Sum header if present
            const sumTh = Array.from(document.querySelectorAll('th')).find(t => /\b(sum|summa|\u03a3)\b/i.test((t.textContent||t.getAttribute('title')||'').toLowerCase()));
            const sumIdx = sumTh ? sumTh.cellIndex : null;
            // iterate left-to-right up to sumIdx or choose rightmost
            let candidate = null;
            if (sumIdx !== null) {
              for (let i = headerDates.length - 1; i >= 0; i--) {
                const h = headerDates[i];
                if (h.idx >= sumIdx) continue; // skip anything at/after sum
                if (isRedColorLocal(h.color)) continue;
                candidate = h; break;
              }
            } else {
              // no sum found: pick rightmost non-red
              for (let i = headerDates.length - 1; i >= 0; i--) {
                const h = headerDates[i]; if (isRedColorLocal(h.color)) continue; candidate = h; break;
              }
            }
            if (!candidate) candidate = headerDates[headerDates.length - 1];
            if (candidate) {
              const inferredYear = (function(){ try { const explicit = Array.from(document.querySelectorAll('input,span,div')).map(n=>(n.value||n.textContent||'').trim()).find(v=>/^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v)); if (explicit){ const d=parseDateFlexible(explicit); if (d) return d.getFullYear(); } } catch(e){} return (new Date()).getFullYear(); })();
              const dt = new Date(inferredYear, candidate.month - 1, candidate.day);
              console.info(LOG_PREFIX, 'findPeriodEndDate: detected end by deterministic th-scan', { raw: candidate.raw, columnIndex: candidate.idx, date: dt });
              return dt;
            }
          }
        }
      } catch (e) {}

        // New: scan the entire table for day/date tokens (handles layouts where first row isn't the date row)
      try {
          try { console.info(LOG_PREFIX, 'findPeriodEndDate: table chosen for Sum-left scan', tbl2 ? { tag: tbl2.tagName, rows: (tbl2.querySelectorAll && tbl2.querySelectorAll('tr').length) || 0 } : null); } catch (e) {}
        // Prefer the table inside the 'Daglig tidregistrering' section if present
        const heading = Array.from(document.querySelectorAll('h1,h2,h3,legend,div,span,th'))
          .find(el => /Arbetstimmar|Daglig tidregistrering|Tidrapport/i.test(el.textContent||''));
        const tableRoot = heading ? (heading.closest('section') || heading.closest('fieldset') || heading.closest('table') || document.body) : document.body;
        try { console.info(LOG_PREFIX, 'findPeriodEndDate: heading found', !!heading, heading && (heading.textContent||'').slice(0,120)); } catch (e) {}
        try { console.info(LOG_PREFIX, 'findPeriodEndDate: tableRoot selected', tableRoot && (tableRoot.tagName || 'body')); } catch (e) {}
        // Choose the most likely table that contains date tokens
        const candidateTables = Array.from(tableRoot.querySelectorAll('table'));
        const dateTokenRe = /\b(\d{1,2})[\/\.](\d{1,2})\b/;
        let tbl2 = null;
        const matches = [];
        // Direct scan: look for floating header DIVs (DivOverflowNoWrap etc.) that contain date tokens
        try {
          const docs = getAllDocuments();
          const floating = [];
          for (const d of docs) {
            try {
              const found = Array.from(d.querySelectorAll('.DivOverflowNoWrap, .Ellipsis, .Separator'))
                .filter(n => {
                  try {
                    const t = (n.textContent || '') + '|' + (n.getAttribute && n.getAttribute('title') || '') + '|' + (n.dataset && n.dataset.originaltext || '') + '|' + (n.innerHTML || '');
                    return dateTokenRe.test(t);
                  } catch (e) { return false; }
                });
              floating.push(...found);
            } catch (e) {}
          }
            if (floating.length) {
              try { console.info(LOG_PREFIX, 'findPeriodEndDate: floating headers count', floating.length); } catch (e) {}
              const candidates = [];
              const floatInfo = [];
              for (const n of floating) {
                try {
                  const hdrCell = n.closest && n.closest('th,td');
                  const raw = (n.getAttribute && n.getAttribute('title') || n.dataset && n.dataset.originaltext || n.textContent || n.innerHTML || '').replace(/<br\s*\/?>(\s*)/gi, ' ').trim();
                  const m = dateTokenRe.exec(raw);
                  if (!m) continue;
                  const inner = n.querySelector && n.querySelector('.DivOverflowNoWrap, .Ellipsis, .Separator');
                  const color = (inner && (window.getComputedStyle ? window.getComputedStyle(inner).color : inner.style && inner.style.color)) || (hdrCell && (window.getComputedStyle ? window.getComputedStyle(hdrCell).color : hdrCell.style && hdrCell.style.color)) || '';
                  const bgColor = (inner && (window.getComputedStyle ? window.getComputedStyle(inner).backgroundColor : inner.style && inner.style.backgroundColor)) || (hdrCell && (window.getComputedStyle ? window.getComputedStyle(hdrCell).backgroundColor : hdrCell.style && hdrCell.style.backgroundColor)) || '';
                  let classes = '';
                  try { classes = hdrCell ? Array.from(hdrCell.classList || []).slice(0,8).join(' ') : (n.className || '').toString(); } catch(e){}
                  let nearestBg = bgColor;
                  try {
                    let elp = inner || hdrCell || n;
                    while (elp && elp.parentElement) {
                      try {
                        const cs = window.getComputedStyle ? window.getComputedStyle(elp) : null;
                        const bg = cs && cs.backgroundColor ? cs.backgroundColor : '';
                        if (bg && !/^(rgba\(0,\s*0,\s*0,\s*0\)|transparent)$/i.test(bg)) { nearestBg = bg; break; }
                      } catch(e){}
                      elp = elp.parentElement;
                    }
                  } catch(e){}
                  let outer = '';
                  try { outer = (n.outerHTML || '').slice(0,200); } catch(e){}
                  // extract inline style color if present (helps when computedStyle differs)
                  let inlineColor = '';
                  try {
                    const s = n.getAttribute && n.getAttribute('style');
                    if (s) {
                      const m = /color\s*:\s*([^;\s]+)/i.exec(s);
                      if (m) inlineColor = m[1];
                    }
                    // also check element.style.color
                    if (!inlineColor && n.style && n.style.color) inlineColor = n.style.color;
                    // normalize hex shorthand to full hex
                    if (inlineColor && /^#[0-9a-f]{3}$/i.test(inlineColor)) inlineColor = inlineColor.split('').map(c=>c+c).join('');
                  } catch(e){}
                  // sample rendered element at the floating node center (helps detect styles applied to other layers)
                  let renderColor = '';
                  let renderBg = '';
                  let renderBefore = '';
                  let renderAfter = '';
                  try {
                    const win = (n.ownerDocument && n.ownerDocument.defaultView) || window;
                    const rect = n.getBoundingClientRect && n.getBoundingClientRect();
                    if (rect && win && typeof win.elementFromPoint === 'function') {
                      const cx = rect.left + (rect.width||0)/2;
                      const cy = rect.top + (rect.height||0)/2;
                      try {
                        const elAt = win.elementFromPoint(cx, cy) || n;
                        const cs = win.getComputedStyle ? win.getComputedStyle(elAt) : (elAt.style||{});
                        renderColor = cs && cs.color ? cs.color : '';
                        renderBg = cs && cs.backgroundColor ? cs.backgroundColor : '';
                        try { renderBefore = win.getComputedStyle(elAt, '::before').color || ''; } catch(e){}
                        try { renderAfter = win.getComputedStyle(elAt, '::after').color || ''; } catch(e){}
                      } catch(e){}
                    }
                  } catch(e){}

                  floatInfo.push({ raw: String(raw).slice(0,120), columnIndex: hdrCell ? hdrCell.cellIndex : null, color: (color||'').toString(), bgColor: (bgColor||'').toString(), classes: classes, nearestBg: (nearestBg||'').toString(), outer: outer, inlineColor: inlineColor, renderColor: renderColor, renderBg: renderBg, renderBefore: renderBefore, renderAfter: renderAfter });
                  candidates.push({ hdrCell, raw, day: Number(m[1]), month: Number(m[2]), idx: hdrCell ? hdrCell.cellIndex : null, color, bgColor, classes, nearestBg, outer, inlineColor, renderColor, renderBg, renderBefore, renderAfter });
                } catch (e) {}
              }
              try { console.info(LOG_PREFIX, 'findPeriodEndDate: floating headers info', floatInfo); } catch (e) {}
              try { console.info(LOG_PREFIX, 'findPeriodEndDate: floating headers info json', JSON.stringify(floatInfo)); } catch (e) {}

              if (candidates.length) {
                // infer year from visible explicit date strings across reachable docs (used to compare dates)
                let inferredYear = (new Date()).getFullYear();
                try {
                  const allTextNodes = [];
                  for (const d2 of docs) {
                    try { allTextNodes.push(...Array.from(d2.querySelectorAll('input,span,div')).map(x => (x.value||x.textContent||'').trim())); } catch(e){}
                  }
                  const explicit = allTextNodes.find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                  if (explicit) {
                    const dd = parseDateFlexible(explicit);
                    if (dd) inferredYear = dd.getFullYear();
                  }
                } catch(e){}

                candidates.sort((a,b) => (b.idx || 0) - (a.idx || 0));
                const sumTh = Array.from(document.querySelectorAll('th')).find(t => /\b(sum|summa|\u03a3)\b/i.test((t.textContent||t.getAttribute('title')||'').toLowerCase()));
                const sumIdx = sumTh ? sumTh.cellIndex : null;
                let chosen = null;

                // helper: detect dark/black-ish colors
                const isBlackish = (colorStr) => {
                  if (!colorStr) return false;
                  try {
                    // rgba or rgb
                    const rgbMatch = /rgba?\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})/i.exec(colorStr);
                    if (rgbMatch) {
                      const r = Number(rgbMatch[1]), g = Number(rgbMatch[2]), b = Number(rgbMatch[3]);
                      const mx = Math.max(r,g,b), mn = Math.min(r,g,b);
                      return mx <= 100 && (mx - mn) <= 40; // dark and low chroma
                    }
                    // hex formats #rrggbb or #rgb
                    const hex = /^#([0-9a-f]{6}|[0-9a-f]{3})/i.exec(colorStr.trim());
                    if (hex) {
                      let h = hex[1];
                      if (h.length === 3) h = h.split('').map(c=>c+c).join('');
                      const r = parseInt(h.slice(0,2),16), g = parseInt(h.slice(2,4),16), b = parseInt(h.slice(4,6),16);
                      const mx = Math.max(r,g,b), mn = Math.min(r,g,b);
                      return mx <= 100 && (mx - mn) <= 40;
                    }
                    // common named black/grays
                    const lowered = (colorStr||'').toString().toLowerCase();
                    if (['black','#000','grey','gray','darkgray','darkgrey'].includes(lowered)) return true;
                  } catch(e){}
                  return false;
                };

                // prefer latest date among candidates that are black/dark gray (not red)
                try {
                  // Gather only candidates that have an inline color and where that inline color is black/dark
                  const blackCandidates = candidates.filter(c => {
                    try {
                      if (columnLooksRed(c)) return false;
                      if (!c.inlineColor) return false;
                      return isBlackish(c.inlineColor.toString());
                    } catch(e){ return false; }
                  });
                  if (blackCandidates.length) {
                    // pick the latest calendar date among blackCandidates
                    let best = null; let bestTime = -Infinity;
                    for (const c of blackCandidates) {
                      try {
                        const dt = new Date(inferredYear, (c.month||1)-1, c.day||1).getTime();
                        if (dt > bestTime) { bestTime = dt; best = c; }
                      } catch(e){}
                    }
                    if (best) {
                      chosen = best;
                      try { console.info(LOG_PREFIX, 'findPeriodEndDate: chose latest black candidate', { raw: chosen.raw, columnIndex: chosen.idx, inferredYear }); } catch(e){}
                    }
                  }
                } catch(e){}
                const isRedColorLocal = (colorStr) => {
                  if (!colorStr) return false;
                  const rgb = /rgb\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)/i.exec(colorStr);
                  if (rgb) {
                    const r = Number(rgb[1]), g = Number(rgb[2]), b = Number(rgb[3]);
                    return r > 140 && r > g + 30 && r > b + 30;
                  }
                  const hex = /#([0-9a-f]{6}|[0-9a-f]{3})/i.exec(colorStr);
                  if (hex) {
                    let h = hex[1]; if (h.length === 3) h = h.split('').map(c=>c+c).join('');
                    const r = parseInt(h.slice(0,2),16), g = parseInt(h.slice(2,4),16), b = parseInt(h.slice(4,6),16);
                    return r > 140 && r > g + 30 && r > b + 30;
                  }
                  return false;
                };

                const columnLooksRed = (candidate) => {
                  try {
                    if (isRedColorLocal(candidate.color) || isRedColorLocal(candidate.bgColor)) return true;
                    let tbl = candidate.hdrCell && candidate.hdrCell.closest && candidate.hdrCell.closest('table');
                    if (!tbl) {
                      try {
                        const docsAll = getAllDocuments();
                        const token = (candidate.raw || '').replace(/\s+-\s+Sidhuvud$/i, '').split(/\s+/).slice(0,3).join(' ');
                        for (const d of docsAll) {
                          try {
                            const tables = Array.from(d.querySelectorAll('table'));
                            for (const t of tables) {
                              try {
                                const hdr = t.querySelector('thead tr') || t.querySelector('tr');
                                if (!hdr) continue;
                                const cells = Array.from(hdr.querySelectorAll('th,td'));
                                const match = cells.find(c => {
                                  try {
                                    const inner = c.querySelector && c.querySelector('.DivOverflowNoWrap, .Ellipsis, .Separator');
                                    const txt = (inner && (inner.getAttribute && inner.getAttribute('title') || inner.dataset && inner.dataset.originaltext || inner.textContent || inner.innerHTML)) || (c.getAttribute && c.getAttribute('title') || c.textContent || c.innerHTML) || '';
                                    return (txt || '').indexOf(token) >= 0 || (candidate.idx !== undefined && (c.cellIndex === candidate.idx));
                                  } catch (e) { return false; }
                                });
                                if (match) { tbl = t; break; }
                              } catch (e) {}
                            }
                            if (tbl) break;
                          } catch (e) {}
                        }
                      } catch (e) {}
                    }
                    if (!tbl) {
                      try { console.info(LOG_PREFIX, 'findPeriodEndDate: no owning table found for floating candidate', { raw: candidate.raw, columnIndex: candidate.idx }); } catch (e) {}
                      return false;
                    }
                    const rows = Array.from(tbl.querySelectorAll('tr'));
                    let checked = 0;
                    for (let r = 1; r < rows.length && checked < 6; r++) {
                      try {
                        const cell = rows[r].cells && rows[r].cells[candidate.idx];
                        if (!cell) continue;
                        const txt = (cell.textContent || '').trim();
                        if (!txt) continue;
                        checked++;
                        const innerCell = cell.querySelector && cell.querySelector('.DivOverflowNoWrap, .Ellipsis, .Separator');
                        const cellColor = (innerCell && (window.getComputedStyle ? window.getComputedStyle(innerCell).color : innerCell.style && innerCell.style.color)) || (window.getComputedStyle ? window.getComputedStyle(cell).color : cell.style && cell.style.color) || '';
                        const cellBg = (innerCell && (window.getComputedStyle ? window.getComputedStyle(innerCell).backgroundColor : innerCell.style && innerCell.style.backgroundColor)) || (window.getComputedStyle ? window.getComputedStyle(cell).backgroundColor : cell.style && cell.style.backgroundColor) || '';
                        if (isRedColorLocal(cellColor) || isRedColorLocal(cellBg)) return true;
                      } catch (e) { /* ignore row errors */ }
                    }
                  } catch (e) {}
                  return false;
                };

                if (!chosen) {
                  if (sumIdx !== null) {
                    for (const c of candidates) {
                      if (c.idx >= sumIdx) continue;
                      if (columnLooksRed(c)) {
                        try { console.info(LOG_PREFIX, 'findPeriodEndDate: skipping floating candidate because column looks red', { raw: c.raw, columnIndex: c.idx, color: c.color, bgColor: c.bgColor }); } catch (e) {}
                        continue;
                      }
                      chosen = c; break;
                    }
                  }
                }
                if (!chosen) {
                  for (const c of candidates) {
                    if (columnLooksRed(c)) {
                      try { console.info(LOG_PREFIX, 'findPeriodEndDate: skipping floating candidate because column looks red', { raw: c.raw, columnIndex: c.idx, color: c.color, bgColor: c.bgColor }); } catch (e) {}
                      continue;
                    }
                    chosen = c; break;
                  }
                }
                if (!chosen) chosen = candidates[0];

                try {
                  let inferredYear = (new Date()).getFullYear();
                  try {
                    const allTextNodes = [];
                    for (const d2 of docs) {
                      try { allTextNodes.push(...Array.from(d2.querySelectorAll('input,span,div')).map(x => (x.value||x.textContent||'').trim())); } catch (e) {}
                    }
                    const explicit = allTextNodes.find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                    if (explicit) {
                      const d = parseDateFlexible(explicit);
                      if (d) inferredYear = d.getFullYear();
                    }
                  } catch (e) {}

                  // Enforce: if any header candidates have inlineColor that is blackish,
                  // choose the latest date among those headers and override `chosen`.
                  try {
                    const inlineBlack = candidates.filter(c => c && c.inlineColor && isBlackish(c.inlineColor.toString()));
                    if (inlineBlack.length) {
                      let bestH = null; let bestT = -Infinity;
                      for (const c of inlineBlack) {
                        try {
                          if (!c.day || !c.month) continue;
                          const dtVal = new Date(inferredYear, (c.month||1)-1, c.day||1).getTime();
                          if (dtVal > bestT) { bestT = dtVal; bestH = c; }
                        } catch(e){}
                      }
                      if (bestH) {
                        chosen = bestH;
                        try { console.info(LOG_PREFIX, 'findPeriodEndDate: overriding chosen with latest inline-black header', { raw: chosen.raw, columnIndex: chosen.idx }); } catch(e){}
                      }
                    }
                  } catch(e){}
                  const dt = new Date(inferredYear, (chosen.month || 1) - 1, chosen.day || 1);
                  console.info(LOG_PREFIX, 'findPeriodEndDate: detected end by floating header (chosen)', { raw: chosen.raw, columnIndex: chosen.idx, date: dt });
                  try { console.info(LOG_PREFIX, 'findPeriodEndDate: floating chosen json', JSON.stringify({ raw: chosen.raw, columnIndex: chosen.idx, date: dt.toISOString(), color: chosen.color, bgColor: chosen.bgColor, classes: chosen.classes, nearestBg: chosen.nearestBg, outer: chosen.outer })); } catch (e) {}
                  return dt;
                } catch (e) {}
              }
            }
        } catch (e) {}
        if (candidateTables.length) {
          let best = null;
          let bestScore = 0;
          candidateTables.forEach((t) => {
            try {
              const txt = (t.innerText || '').trim();
              let score = 0;
              if (dateTokenRe.test(txt)) score += 10;
              // count day tokens
              const dayMatches = txt.match(/\b\d{1,2}[\/\.]\d{1,2}\b/g) || [];
              score += dayMatches.length;
              // prefer tables with multiple columns/rows
              const cols = t.querySelectorAll('tr:first-child th, tr:first-child td').length || 0;
              const rows = t.querySelectorAll('tr').length || 0;
              score += Math.min(10, cols) + Math.min(5, rows);
              if (score > bestScore) { bestScore = score; best = t; }
            } catch (e) {}
          });
          try { console.info(LOG_PREFIX, 'findPeriodEndDate: candidateTables count', candidateTables.length, 'bestScore', bestScore, 'bestTable', !!best); } catch (e) {}
          tbl2 = best || candidateTables[0];
        } else {
          try { console.info(LOG_PREFIX, 'findPeriodEndDate: no candidateTables found under tableRoot'); } catch (e) {}
          // If none found under the chosen tableRoot, fall back to searching all reachable documents
          try {
            const docs = getAllDocuments();
            const allTables = [];
            for (const d of docs) {
              try { allTables.push(...Array.from(d.querySelectorAll('table'))); } catch (e) {}
            }
            if (allTables.length) {
              // Score tables across documents similarly to the candidateTables path
              let bestGlobal = null;
              let bestGlobalScore = 0;
              allTables.forEach((t) => {
                try {
                  const txt = (t.innerText || '').trim();
                  let score = 0;
                  if (dateTokenRe.test(txt)) score += 10;
                  const dayMatches = txt.match(/\b\d{1,2}[\/\.]\d{1,2}\b/g) || [];
                  score += dayMatches.length;
                  const cols = t.querySelectorAll('tr:first-child th, tr:first-child td').length || 0;
                  const rows = t.querySelectorAll('tr').length || 0;
                  score += Math.min(10, cols) + Math.min(5, rows);
                  if (score > bestGlobalScore) { bestGlobalScore = score; bestGlobal = t; }
                } catch (e) {}
              });
              tbl2 = bestGlobal || allTables[0];
            } else {
              tbl2 = tableRoot.querySelector('table');
            }
          } catch (e) {
            tbl2 = tableRoot.querySelector('table');
          }
        }

        // New heuristic: locate a header cell labelled 'Sum' (or variants) and scan left
        // from that column. If a column's displayed text is styled red, skip it;
        // the first non-red column to the left is assumed to be the last workday.
        try {
          if (tbl2) {
            let headerRow = tbl2.querySelector('thead tr') || tbl2.querySelector('tr');
            // If the chosen header row looks too small, try a few top rows to find a better header
            try {
              if (headerRow) {
                const topRows = Array.from(tbl2.querySelectorAll('tr'));
                if ((headerRow.querySelectorAll('th,td').length || 0) < 2) {
                  for (let ri = 0; ri < Math.min(6, topRows.length); ri++) {
                    const r = topRows[ri];
                    if ((r.querySelectorAll('th,td').length || 0) > 2) { headerRow = r; break; }
                  }
                }
              }
            } catch (e) {}
            if (headerRow) {
              try {
                // Collect header text and computed color for debugging
                const hdrCells = Array.from(headerRow.querySelectorAll('th,td'));
                const headerInfo = hdrCells.map((h, ci) => {
                  try {
                    const inner = h.querySelector && h.querySelector('.DivOverflowNoWrap, .Ellipsis, .Separator');
                    const rawText = (inner && ((inner.dataset && inner.dataset.originaltext) || inner.getAttribute && inner.getAttribute('title') || inner.textContent || inner.innerHTML)) || (h && (h.getAttribute && h.getAttribute('title') || h.textContent || h.innerHTML)) || '';
                    let color = '';
                    try { color = (window.getComputedStyle ? window.getComputedStyle(inner || h).color : (inner && inner.style && inner.style.color) || (h && h.style && h.style.color)) || ''; } catch (e) { color = ''; }
                    return { idx: (h.cellIndex || ci), text: String(rawText).replace(/<br\s*\/?>(\s*)/gi, ' ').trim().slice(0,120), color };
                  } catch (e) { return { idx: ci, text: '', color: '' }; }
                });
                try { console.info(LOG_PREFIX, 'findPeriodEndDate: header info', headerInfo); } catch (e) {}
              } catch (e) {}
              const headers = Array.from(headerRow.querySelectorAll('th,td'));
              const sumIdx = headers.findIndex(h => /\b(sum|summa|\u03a3)\b/i.test((h.textContent||'').trim()));
                if (sumIdx > 0) {
                // helper to detect red-ish colors
                const isRedColor = (colorStr) => {
                  if (!colorStr) return false;
                  const rgb = /rgb\s*\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)/i.exec(colorStr);
                  if (rgb) {
                    const r = Number(rgb[1]), g = Number(rgb[2]), b = Number(rgb[3]);
                    return r > 140 && r > g + 30 && r > b + 30;
                  }
                  const hex = /#([0-9a-f]{6}|[0-9a-f]{3})/i.exec(colorStr);
                  if (hex) {
                    let h = hex[1]; if (h.length === 3) h = h.split('').map(c=>c+c).join('');
                    const r = parseInt(h.slice(0,2),16), g = parseInt(h.slice(2,4),16), b = parseInt(h.slice(4,6),16);
                    return r > 140 && r > g + 30 && r > b + 30;
                  }
                  return false;
                };

                // iterate left from sumIdx - 1, look for a non-red column
                for (let ci = sumIdx - 1; ci >= 0; ci--) {
                  try {
                    const headerCell = headers[ci];
                    // prefer inner DivOverflowNoWrap / Ellipsis / Separator when reading color/text
                    let headerInner = null;
                    try { headerInner = headerCell && headerCell.querySelector && headerCell.querySelector('.DivOverflowNoWrap, .Ellipsis, .Separator'); } catch (e) { headerInner = null; }
                    // quick check on header color (prefer inner element if present)
                    const headerColor = (headerInner && (window.getComputedStyle ? window.getComputedStyle(headerInner).color : headerInner.style && headerInner.style.color)) || (headerCell && (window.getComputedStyle ? window.getComputedStyle(headerCell).color : headerCell.style && headerCell.style.color));
                    let colIsRed = isRedColor(headerColor);

                    // If header not red, inspect up to first 6 data rows in that column
                    if (!colIsRed) {
                      const rows = Array.from(tbl2.querySelectorAll('tr'));
                      let checked = 0;
                      for (let r = 1; r < rows.length && checked < 6; r++) {
                        const cell = rows[r].cells && rows[r].cells[ci];
                        if (!cell) continue;
                        const txt = (cell.textContent || '').trim();
                        if (!txt) continue;
                        checked++;
                        // prefer inner styled element inside the data cell
                        let cellInner = null;
                        try { cellInner = cell.querySelector && cell.querySelector('.DivOverflowNoWrap, .Ellipsis, .Separator'); } catch (e) { cellInner = null; }
                        const cellColor = (cellInner && (window.getComputedStyle ? window.getComputedStyle(cellInner).color : cellInner.style && cellInner.style.color)) || (window.getComputedStyle ? window.getComputedStyle(cell).color : cell.style && cell.style.color);
                        if (isRedColor(cellColor)) {
                          colIsRed = true;
                          break;
                        }
                      }
                    }

                    if (!colIsRed) {
                      // Try to parse a date token from the header cell, its attributes, or inner HTML
                      let txt = '';
                      try {
                        const sourceEl = (headerInner && headerInner.nodeType) ? headerInner : headerCell;
                        txt = (sourceEl && (sourceEl.textContent || sourceEl.getAttribute && sourceEl.getAttribute('title') || sourceEl.dataset && sourceEl.dataset.originaltext || sourceEl.innerHTML)) || '';
                        // Replace any <br> tags with spaces for parsing
                        txt = String(txt).replace(/<br\s*\/?>(\s*)/gi, ' ');
                        txt = txt.trim();
                      } catch (e) { txt = (headerCell && headerCell.textContent) || ''; }
                      const m = dateTokenRe.exec(txt);
                      if (m) {
                        // infer year similar to other heuristics
                        let inferredYear = (new Date()).getFullYear();
                        try {
                          const explicit = Array.from(document.querySelectorAll('input,span,div')).map(n => (n.value||n.textContent||'').trim()).find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                          if (explicit) {
                            const d = parseDateFlexible(explicit);
                            if (d) inferredYear = d.getFullYear();
                          } else {
                            const bodyMatch = (document.body && document.body.innerText) || '';
                            const bm = /(\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4})/.exec(bodyMatch);
                            if (bm) {
                              const d = parseDateFlexible(bm[1]);
                              if (d) inferredYear = d.getFullYear();
                            }
                          }
                        } catch (e) {}
                        const dt = new Date(inferredYear, Number(m[2]) - 1, Number(m[1]));
                        console.info(LOG_PREFIX, 'findPeriodEndDate: detected end by Sum-left red-scan', { headerText: txt, columnIndex: ci, date: dt });
                        return dt;
                      }
                      // if header doesn't include explicit date token, attempt to find any day token inside header or first data cell
                      try {
                        const maybe = (headerCell && (headerCell.textContent || headerCell.getAttribute('title') || headerCell.dataset && headerCell.dataset.originaltext || headerCell.innerHTML)) || '';
                        const maybeClean = String(maybe).replace(/<br\s*\/?>(\s*)/gi, ' ').trim();
                        const dm = /\b(\d{1,2})\b/.exec(maybeClean);
                        if (dm) {
                          let inferredYear = (new Date()).getFullYear();
                          try {
                            const explicit = Array.from(document.querySelectorAll('input,span,div')).map(n => (n.value||n.textContent||'').trim()).find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                            if (explicit) {
                              const d = parseDateFlexible(explicit);
                              if (d) inferredYear = d.getFullYear();
                            }
                          } catch (e) {}
                          const day = Number(dm[1]);
                          // attempt to infer month by looking rightmost date-like header before sum
                          let month = null;
                          for (let k = ci; k >= Math.max(0, ci - 8); k--) {
                            try {
                              const hhRaw = headers[k] && (headers[k].textContent || headers[k].getAttribute('title') || headers[k].dataset && headers[k].dataset.originaltext || headers[k].innerHTML) || '';
                              const hh = String(hhRaw).replace(/<br\s*\/?>(\s*)/gi, ' ').trim();
                              const mm = /\b(\d{1,2})[\/\.](\d{1,2})\b/.exec(hh);
                              if (mm) { month = Number(mm[2]); break; }
                            } catch (e) {}
                          }
                          if (!month) month = (new Date()).getMonth() + 1;
                          const dt = new Date(inferredYear, month - 1, day);
                          console.info(LOG_PREFIX, 'findPeriodEndDate: detected end by Sum-left day-scan', { headerText: maybe, columnIndex: ci, date: dt });
                          return dt;
                        }
                      } catch (e) {}
                    }
                  } catch (e) {
                    // continue to next column on any error
                  }
                }
              }
            }
          }
        } catch (e) {
          // ignore Sum-left heuristic errors
        }

        // Fallback: if Sum-left didn't return, try scanning header cells for date tokens
        try {
          if (tbl2) {
            const headerRowCandidates = Array.from(tbl2.querySelectorAll('tr')).slice(0, 6);
            let headerCells = [];
            for (const r of headerRowCandidates) {
              const cols = Array.from(r.querySelectorAll('th,td'));
              if (cols.length > headerCells.length) headerCells = cols;
            }
            if (headerCells.length) {
              // collect date-like header cells with their column index
              const hdrs = headerCells.map((h, idx) => {
                let txt = '';
                try { txt = (h.textContent || h.getAttribute('title') || h.dataset && h.dataset.originaltext || h.innerHTML) || ''; txt = String(txt).replace(/<br\s*\/?>(\s*)/gi, ' ').trim(); } catch (e) { txt = (h.textContent||'').trim(); }
                const m = dateTokenRe.exec(txt);
                return { el: h, idx, txt, hasDate: !!m };
              }).filter(x => x.hasDate);

              // iterate right-to-left across detected date headers and return first non-red column
              for (let i = hdrs.length - 1; i >= 0; i--) {
                const info = hdrs[i];
                const ci = info.idx;
                let colIsRed = false;
                try {
                  const headerColor = info.el && (window.getComputedStyle ? window.getComputedStyle(info.el).color : info.el.style && info.el.style.color);
                  colIsRed = isRedColor(headerColor);
                } catch (e) {}
                if (!colIsRed) {
                  // inspect a few rows for red text
                  try {
                    const rows = Array.from(tbl2.querySelectorAll('tr'));
                    let checked = 0;
                    for (let r = 1; r < rows.length && checked < 6; r++) {
                      const cell = rows[r].cells && rows[r].cells[ci];
                      if (!cell) continue;
                      const txt = (cell.textContent || '').trim();
                      if (!txt) continue;
                      checked++;
                      const cellColor = window.getComputedStyle ? window.getComputedStyle(cell).color : cell.style && cell.style.color;
                      if (isRedColor(cellColor)) { colIsRed = true; break; }
                    }
                  } catch (e) {}
                }
                if (!colIsRed) {
                  // parse date from header text
                  const m2 = dateTokenRe.exec(info.txt);
                  if (m2) {
                    let inferredYear = (new Date()).getFullYear();
                    try {
                      const explicit = Array.from(document.querySelectorAll('input,span,div')).map(n => (n.value||n.textContent||'').trim()).find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                      if (explicit) {
                        const d = parseDateFlexible(explicit);
                        if (d) inferredYear = d.getFullYear();
                      }
                    } catch (e) {}
                    const dt = new Date(inferredYear, Number(m2[2]) - 1, Number(m2[1]));
                    console.info(LOG_PREFIX, 'findPeriodEndDate: detected end by header-scan fallback', { headerText: info.txt, columnIndex: ci, date: dt });
                    return dt;
                  }
                }
              }
            }
          }
        } catch (e) {
          // ignore fallback errors
        }
        if (matches.length) {
          // Infer year from any explicit dd/mm/yyyy found on page or in nearby 'Datum i perioden' input
          let inferredYear = (new Date()).getFullYear();
          try {
            const explicit = Array.from(document.querySelectorAll('input,span,div')).map(n => (n.value||n.textContent||'').trim()).find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
            if (explicit) {
              const d = parseDateFlexible(explicit);
              if (d) inferredYear = d.getFullYear();
            } else {
              // try previous heuristic: look for any dd/mm/yyyy anywhere
              const bodyMatch = (document.body && document.body.innerText) || '';
              const m = /(\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4})/.exec(bodyMatch);
              if (m) {
                const d = parseDateFlexible(m[1]);
                if (d) inferredYear = d.getFullYear();
              }
            }
          } catch (e) {}
          // Build candidate dates for all matches and choose the latest local date
          try {
            const dts = matches.map(ch => new Date(inferredYear, ch.month - 1, ch.day));
            // Normalize by local date value (strip time)
            const maxDt = dts.reduce((a,b) => (a > b ? a : b));
            console.info(LOG_PREFIX, 'findPeriodEndDate: parsed table matches', { matchesCount: matches.length, inferredYear, maxDt });
            return maxDt;
          } catch (e) {
            const chosen = matches[matches.length - 1];
            const dt = new Date(inferredYear, chosen.month - 1, chosen.day);
            console.info(LOG_PREFIX, 'findPeriodEndDate: fallback parsed table match', { text: chosen.text, day: chosen.day, month: chosen.month, year: inferredYear, dt, matchesCount: matches.length });
            return dt;
          }
        }
      } catch (e) {
        // ignore
      }

      // Fallback: search body text
      const body = (document.body && document.body.innerText) || '';
      let m = dateRangeIso.exec(body);
      if (m) {
        console.info(LOG_PREFIX, 'findPeriodEndDate: matched iso range in body', m[0].slice(0,200));
        return parseDateFlexible(m[2]);
      }
      m = dateRangeSlashed.exec(body);
      if (m) {
        console.info(LOG_PREFIX, 'findPeriodEndDate: matched slashed range in body', m[0].slice(0,200));
        return parseDateFlexible(m[2]);
      }
      m = monthNameRange.exec(body);
      if (m) {
        console.info(LOG_PREFIX, 'findPeriodEndDate: matched month name range in body', m[0].slice(0,200));
        return parseDateFlexible(m[2]);
      }
      // Same-origin frames: attempt the same heuristics inside each frame (handles Agresso iframe nesting)
      try {
        for (let fi = 0; fi < (window.frames && window.frames.length || 0); fi++) {
          try {
            const fr = window.frames[fi];
            const fd = fr.document;
            if (!fd) continue;
            // Look for DivOverflowNoWrap-style date headers first
            const candidates = Array.from(fd.querySelectorAll('.DivOverflowNoWrap, .Ellipsis, .Separator'));
            const dateNodes = candidates.filter((n) => {
              try {
                const t = (n.textContent || '').trim();
                const title = n.getAttribute && (n.getAttribute('title') || '');
                const data = n.dataset && n.dataset.originaltext ? n.dataset.originaltext : '';
                const inner = n.innerHTML || '';
                return dateTokenRe.test(t) || dateTokenRe.test(title) || dateTokenRe.test(data) || dateTokenRe.test(inner);
              } catch (e) { return false; }
            });

            if (dateNodes.length) {
              try { console.info(LOG_PREFIX, 'findPeriodEndDate: frame dateNodes count', fi, dateNodes.length); } catch (e) {}
              // Group by table
              const tablesMap = new Map();
              dateNodes.forEach((n) => {
                try {
                  const tbl = n.closest && n.closest('table');
                  if (!tbl) return;
                  if (!tablesMap.has(tbl)) tablesMap.set(tbl, []);
                  tablesMap.get(tbl).push(n);
                } catch (e) {}
              });

              // If no tables found from these nodes, try a broader search for nodes
              if (tablesMap.size === 0) {
                try {
                  const extra = Array.from(fd.querySelectorAll('[data-originaltext], [title]'))
                    .filter(el => {
                      try {
                        const v = (el.dataset && el.dataset.originaltext) || el.getAttribute('title') || el.innerHTML || '';
                        return dateTokenRe.test(String(v));
                      } catch (e) { return false; }
                    });
                  try { console.info(LOG_PREFIX, 'findPeriodEndDate: frame extra candidate count', fi, extra.length); } catch (e) {}
                  extra.forEach((n) => {
                    try {
                      const tbl = n.closest && n.closest('table');
                      if (!tbl) return;
                      if (!tablesMap.has(tbl)) tablesMap.set(tbl, []);
                      tablesMap.get(tbl).push(n);
                    } catch (e) {}
                  });
                  // Spatial fallback: if still no tables mapped, try matching each date node
                  // to any table in the frame by comparing the node's horizontal center
                  // with header cell bounding rects.
                  if (tablesMap.size === 0) {
                    try {
                      const allTables = Array.from(fd.querySelectorAll('table'));
                      if (allTables.length) {
                        for (const n of extra.length ? extra : nodes) {
                          try {
                            const srcRect = n.getBoundingClientRect ? n.getBoundingClientRect() : null;
                            if (!srcRect) continue;
                            const srcX = srcRect.left + (srcRect.width || 0) / 2;
                            for (const t of allTables) {
                              try {
                                const hdr = t.querySelector('thead tr') || t.querySelector('tr');
                                if (!hdr) continue;
                                const hdrs = Array.from(hdr.querySelectorAll('th,td'));
                                for (let i = 0; i < hdrs.length; i++) {
                                  try {
                                    const r = hdrs[i].getBoundingClientRect();
                                    if (srcX >= (r.left - 2) && srcX <= (r.right + 2)) {
                                      if (!tablesMap.has(t)) tablesMap.set(t, []);
                                      tablesMap.get(t).push(n);
                                      throw 'mapped';
                                    }
                                  } catch (e) {
                                    if (e === 'mapped') break;
                                  }
                                }
                                // if mapped, move to next node
                                if (tablesMap.has(t) && tablesMap.get(t).indexOf(n) >= 0) break;
                              } catch (e) {}
                            }
                          } catch (e) {}
                        }
                      }
                    } catch (e) {}
                  }
                } catch (e) {}
              }

              for (const [tbl, nodes] of tablesMap.entries()) {
                try {
                  // New: try mapping date DIV nodes directly to their nearest th/td
                  for (const n of nodes) {
                    try {
                      const headerCellDirect = n.closest && n.closest('th,td');
                      if (headerCellDirect && headerCellDirect.cellIndex >= 0) {
                        // prefer date token on the node (handles <div class="DivOverflowNoWrap" elements)
                        let txtRaw = '';
                        try { txtRaw = (n.getAttribute && (n.getAttribute('title') || '')) || n.dataset && n.dataset.originaltext || (n.textContent || n.innerHTML || ''); } catch (e) { txtRaw = (n.textContent||n.innerHTML||''); }
                        txtRaw = String(txtRaw).replace(/<br\s*\/?>(\s*)/gi, ' ').trim();
                        const m = dateTokenRe.exec(txtRaw);
                        if (m) {
                          let inferredYear = (new Date()).getFullYear();
                          try {
                            const explicit = Array.from(fd.querySelectorAll('input,span,div')).map(x => (x.value||x.textContent||'').trim()).find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                            if (explicit) {
                              const d = parseDateFlexible(explicit);
                              if (d) inferredYear = d.getFullYear();
                            }
                          } catch (e) {}
                          const dt = new Date(inferredYear, Number(m[2]) - 1, Number(m[1]));
                          console.info(LOG_PREFIX, 'findPeriodEndDate: mapped date-node to header cell in frame', { frame: fi, nodeText: txtRaw, columnIndex: headerCellDirect.cellIndex, date: dt });
                          return dt;
                        }
                      }
                    } catch (e) {}
                  }
                  // find header row/cells
                  const headerRow = tbl.querySelector('thead tr') || tbl.querySelector('tr');
                  const headers = headerRow ? Array.from(headerRow.querySelectorAll('th,td')) : [];
                  const sumIdx = headers.findIndex(h => /\b(sum|summa|\u03a3)\b/i.test((h.textContent||'').trim()));
                  const frGetStyle = (el) => { try { return fr.getComputedStyle ? fr.getComputedStyle(el).color : (el.style && el.style.color) || ''; } catch (e) { return ''; } };
                  const isRedInFrame = (colorStr) => {
                    return isRedColor(colorStr);
                  };

                  // Spatial mapping: find which header column horizontally matches a floating node
                  const getColumnIndexByX = (tblNode, srcNode) => {
                    try {
                      const srcRect = srcNode.getBoundingClientRect();
                      const srcX = srcRect.left + srcRect.width / 2;
                      const hdrRow = tblNode.querySelector('thead tr') || tblNode.querySelector('tr');
                      if (!hdrRow) return -1;
                      const hdrs = Array.from(hdrRow.querySelectorAll('th,td'));
                      for (let i = 0; i < hdrs.length; i++) {
                        try {
                          const r = hdrs[i].getBoundingClientRect();
                          if (srcX >= (r.left - 2) && srcX <= (r.right + 2)) return i;
                        } catch (e) { continue; }
                      }
                    } catch (e) {}
                    return -1;
                  };

                  if (sumIdx > 0) {
                    for (let ci = sumIdx - 1; ci >= 0; ci--) {
                      try {
                        const headerCell = headers[ci];
                        const headerColor = headerCell && frGetStyle(headerCell);
                        let colIsRed = isRedInFrame(headerColor);
                        if (!colIsRed) {
                          const rows = Array.from(tbl.querySelectorAll('tr'));
                          let checked = 0;
                          for (let r = 1; r < rows.length && checked < 6; r++) {
                            const cell = rows[r].cells && rows[r].cells[ci];
                            if (!cell) continue;
                            const txt = (cell.textContent || '').trim();
                            if (!txt) continue;
                            checked++;
                            const cellColor = frGetStyle(cell);
                            if (isRedInFrame(cellColor)) { colIsRed = true; break; }
                          }
                        }
                        if (!colIsRed) {
                          // parse date from header cell (title, data-originaltext, innerHTML)
                          let txt = '';
                          try { txt = (headerCell && (headerCell.textContent || headerCell.getAttribute('title') || headerCell.dataset && headerCell.dataset.originaltext || headerCell.innerHTML)) || ''; txt = String(txt).replace(/<br\s*\/?>(\s*)/gi, ' ').trim(); } catch (e) { txt = (headerCell && headerCell.textContent) || ''; }
                          const m = dateTokenRe.exec(txt);
                          if (m) {
                            let inferredYear = (new Date()).getFullYear();
                            try {
                              const explicit = Array.from(fd.querySelectorAll('input,span,div')).map(n => (n.value||n.textContent||'').trim()).find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                              if (explicit) {
                                const d = parseDateFlexible(explicit);
                                if (d) inferredYear = d.getFullYear();
                              }
                            } catch (e) {}
                            const dt = new Date(inferredYear, Number(m[2]) - 1, Number(m[1]));
                            console.info(LOG_PREFIX, 'findPeriodEndDate: detected end in frame by Sum-left', { frame: fi, headerText: txt, columnIndex: ci, date: dt });
                            return dt;
                          }
                        }
                      } catch (e) {}
                    }
                  }

                  // fallback: examine date-like nodes in this table, choose rightmost non-red
                  // Also attempt spatial mapping: map floating date nodes to columns by X coordinate
                  for (const dn of nodes) {
                    try {
                      let raw = '';
                      try { raw = (dn.getAttribute && (dn.getAttribute('title') || '')) || dn.dataset && dn.dataset.originaltext || (dn.textContent || dn.innerHTML || ''); } catch (e) { raw = (dn.textContent||dn.innerHTML||''); }
                      raw = String(raw).replace(/<br\s*\/?>(\s*)/gi, ' ').trim();
                      const m = dateTokenRe.exec(raw);
                      if (!m) continue;
                      let colIdx = -1;
                      try {
                        const maybeCell = dn.closest && dn.closest('th,td');
                        if (maybeCell && typeof maybeCell.cellIndex === 'number') colIdx = maybeCell.cellIndex;
                        if (colIdx < 0) colIdx = getColumnIndexByX(tbl, dn);
                      } catch (e) { colIdx = getColumnIndexByX(tbl, dn); }
                      if (colIdx >= 0) {
                        let inferredYear = (new Date()).getFullYear();
                        try {
                          const explicit = Array.from(fd.querySelectorAll('input,span,div')).map(n => (n.value||n.textContent||'').trim()).find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                          if (explicit) {
                            const d = parseDateFlexible(explicit);
                            if (d) inferredYear = d.getFullYear();
                          }
                        } catch (e) {}
                        const dt = new Date(inferredYear, Number(m[2]) - 1, Number(m[1]));
                        console.info(LOG_PREFIX, 'findPeriodEndDate: mapped node->column by geometry in frame', { frame: fi, raw, columnIndex: colIdx, date: dt });
                        return dt;
                      }
                    } catch (e) {}
                  }
                  const headerCells = headers.length ? headers : Array.from(tbl.querySelectorAll('th,td'));
                  const hdrs = [];
                  headerCells.forEach((h, idx) => {
                    try {
                      let txt = (h.textContent || h.getAttribute('title') || h.dataset && h.dataset.originaltext || h.innerHTML) || '';
                      txt = String(txt).replace(/<br\s*\/?>(\s*)/gi, ' ').trim();
                      if (dateTokenRe.test(txt)) hdrs.push({ el: h, idx, txt });
                    } catch (e) {}
                  });
                  for (let i = hdrs.length - 1; i >= 0; i--) {
                    const info = hdrs[i];
                    const ci = info.idx;
                    let colIsRed = false;
                    try { colIsRed = isRedInFrame(frGetStyle(info.el)); } catch (e) {}
                    if (!colIsRed) {
                      try {
                        const rows = Array.from(tbl.querySelectorAll('tr'));
                        let checked = 0;
                        for (let r = 1; r < rows.length && checked < 6; r++) {
                          const cell = rows[r].cells && rows[r].cells[ci];
                          if (!cell) continue;
                          const txt = (cell.textContent || '').trim();
                          if (!txt) continue;
                          checked++;
                          const cellColor = frGetStyle(cell);
                          if (isRedInFrame(cellColor)) { colIsRed = true; break; }
                        }
                      } catch (e) {}
                    }
                    if (!colIsRed) {
                      const m2 = dateTokenRe.exec(info.txt);
                      if (m2) {
                        let inferredYear = (new Date()).getFullYear();
                        try {
                          const explicit = Array.from(fd.querySelectorAll('input,span,div')).map(n => (n.value||n.textContent||'').trim()).find(v => /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(v));
                          if (explicit) {
                            const d = parseDateFlexible(explicit);
                            if (d) inferredYear = d.getFullYear();
                          }
                        } catch (e) {}
                        const dt = new Date(inferredYear, Number(m2[2]) - 1, Number(m2[1]));
                        console.info(LOG_PREFIX, 'findPeriodEndDate: detected end in frame by header fallback', { frame: fi, headerText: info.txt, columnIndex: ci, date: dt });
                        return dt;
                      }
                    }
                  }
                } catch (e) {}
              }
            }
          } catch (e) {
            // cross-origin or access error - ignore
          }
        }
      } catch (e) {}
      // If not found in this document, try same-origin iframes (Agresso may render inside frames)
      try {
        for (let i = 0; i < (window.frames && window.frames.length || 0); i++) {
          try {
            const fr = window.frames[i];
            const fd = fr.document;
            if (!fd) continue;
            const fbody = (fd.body && fd.body.innerText) || '';
            let fm = dateRangeIso.exec(fbody);
            if (fm) {
              console.info(LOG_PREFIX, 'findPeriodEndDate: matched iso range in iframe body', fm[0].slice(0,200));
              return parseDateFlexible(fm[2]);
            }
            fm = dateRangeSlashed.exec(fbody);
            if (fm) {
              console.info(LOG_PREFIX, 'findPeriodEndDate: matched slashed range in iframe body', fm[0].slice(0,200));
              return parseDateFlexible(fm[2]);
            }
            fm = monthNameRange.exec(fbody);
            if (fm) {
              console.info(LOG_PREFIX, 'findPeriodEndDate: matched month name range in iframe body', fm[0].slice(0,200));
              return parseDateFlexible(fm[2]);
            }
          } catch (e) {
            // cross-origin or other access errors - ignore
          }
        }
      } catch (e) {}

      // Extra heuristics: look for labelled 'Datum' inputs or nearby tokens
      try {
        const datumNode = Array.from(document.querySelectorAll('label,div,span,th,td')).find(n => /Datum i perioden|Datum i period|Datum i perioden|Datum i perioder|Datum/i.test(n.textContent || ''));
        if (datumNode) {
          // try to locate an input within the same row or nearby
          const input = datumNode.closest('tr')?.querySelector('input') || datumNode.querySelector('input') || datumNode.nextElementSibling?.querySelector('input') || document.querySelector('input[name*="datum"], input[id*="datum"], input[name*="date"], input[id*="date"], input[type="date"]');
          const val = input && (input.value || input.getAttribute('value') || '').trim();
          if (val) {
            const pd = parseDateFlexible(val);
            if (pd) {
              console.info(LOG_PREFIX, 'findPeriodEndDate: parsed date from datum input', val, pd);
              return pd;
            }
          }
          const nearbyText = (datumNode.textContent || '') + ' ' + (datumNode.nextElementSibling && datumNode.nextElementSibling.textContent || '');
          const tokenMatch = /(\d{1,2}[\/\.]\\\d{1,2}(?:[\/\.]\d{2,4})?)/.exec(nearbyText) || /(\d{1,2}\s+[A-Za-zåäöÅÄÖ]{3,}\s*\d{0,4})/.exec(nearbyText);
          if (tokenMatch) {
            const pd = parseDateFlexible(tokenMatch[1]);
            if (pd) {
              console.info(LOG_PREFIX, 'findPeriodEndDate: parsed date from nearby datum text', tokenMatch[1], pd);
              return pd;
            }
          }
        }
      } catch (e) {}
    } catch (e) {
      // ignore
    }
    return null;
  }

  function buildDebugReport() {
    try {
      const heading = Array.from(document.querySelectorAll('h1,h2,h3,legend,div,span,th'))
        .find(el => /Arbetstimmar|Daglig tidregistrering|Tidrapport/i.test(el.textContent||''));
      const tbl = heading ? (heading.closest('section') || heading.closest('fieldset') || heading.closest('table') || document.body).querySelector('table') : null;
      const headerRow = tbl ? (tbl.querySelector('tr') || tbl.querySelector('thead tr')) : null;
      const cells = headerRow ? Array.from(headerRow.querySelectorAll('th,td')) : [];
      const headerCells = cells.map(c => ({ text: (c.innerText||'').trim(), html: (c.innerHTML||'').trim(), outer: (c.outerHTML||'').slice(0,500), attrs: Array.from(c.attributes||[]).map(a=>({name:a.name,value:a.value})) }));

      const attrMatches = [];
      if (tbl) {
        tbl.querySelectorAll('*').forEach(el=>{
          Array.from(el.attributes||[]).forEach(a=>{
            if (/\d{1,2}[\/\.]\d{1,2}/.test(a.value)) {
              attrMatches.push({ tag: el.tagName, attr: a.name, value: a.value, outer: (el.outerHTML||'').slice(0,300) });
            }
          });
        });
      }

      const textMatches = [];
      if (tbl) {
        tbl.querySelectorAll('*').forEach(el=>{
          const t = (el.textContent||'').trim();
          if (/\b\d{1,2}[\/\.]\d{1,2}\b/.test(t)) textMatches.push({ tag: el.tagName, text: t.slice(0,200), outer: (el.outerHTML||'').slice(0,200) });
        });
      }

      const datumInput = (() => {
        const node = Array.from(document.querySelectorAll('label,div,span,th,td')).find(n => /Datum i perioden/i.test(n.textContent || ''));
        if (!node) return null;
        const input = node.closest('tr')?.querySelector('input') || node.querySelector('input') || node.nextElementSibling?.querySelector('input') || document.querySelector('input[type="text"], input[type="date"]');
        return input ? { value: input.value || input.getAttribute('value') || null, outer: input.outerHTML } : null;
      })();

      return { heading: heading ? (heading.textContent||'').trim() : null, tableFound: !!tbl, headerCells, attrMatches: attrMatches.slice(0,50), textMatches: textMatches.slice(0,50), datumInput };
    } catch (e) {
      return { error: String(e) };
    }
  }

  function isSameDay(a, b) {
    return a && b && a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
  }

  function localIsoDate(d) {
    try {
      if (!d || !d.getFullYear) return null;
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const da = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${da}`;
    } catch (e) { return null; }
  }

  function showPeriodNotification(endDate) {
    const today = new Date();
    // Respect user preference for reminder
    try {
      const enabled = (function() { try { const v = localStorage.getItem(REMINDER_ENABLED_KEY); return v === null ? true : v === '1' || v === 'true'; } catch (e) { return true; } })();
      console.info(LOG_PREFIX, 'showPeriodNotification: reminder enabled?', enabled);
      if (!enabled) return;
    } catch (e) {}

    const lastNotified = (() => {
      try {
        try {
          if (window.top && window.top.localStorage) return window.top.localStorage.getItem(PERIOD_NOTIFY_KEY);
        } catch (e) {}
        return localStorage.getItem(PERIOD_NOTIFY_KEY);
      } catch (e) { try { return localStorage.getItem(PERIOD_NOTIFY_KEY); } catch (e2) { return null; } }
    })();
    const endIso = localIsoDate(endDate) || endDate.toISOString().slice(0,10);
    console.info(LOG_PREFIX, 'showPeriodNotification: endIso', endIso, 'lastNotified', lastNotified);
    // If the report status is already 'Klar', skip showing any notification/UI.
    try {
      if (isReportStatusKlar()) {
        try { console.info(LOG_PREFIX, 'showPeriodNotification: status is Klar, skipping notification entirely'); } catch (e) {}
        return;
      }
    } catch (e) {}
    if (lastNotified === endIso) {
      // Already stored as notified — ensure the visual indicator is applied
      try { console.info(LOG_PREFIX, 'showPeriodNotification: already notified, enforcing UI highlight'); } catch (e) {}
      // If the report status is already 'Klar', do not enforce the highlight
      try {
        if (isReportStatusKlar()) {
          try { console.info(LOG_PREFIX, 'showPeriodNotification: status is Klar, skipping UI enforcement'); } catch (e) {}
          return;
        }
      } catch (e) {}
      try { /* force UI-only highlight without updating storage */ notifyNow(true); } catch (e) {}
      try {
        // Ask the top-level frame to enforce the highlight as well (works via postMessage
        // even across origins; the top frame's content-script will listen and apply).
        if (window.top && window.top !== window) {
          try { window.top.postMessage({ type: 'agresso_period_enforce', endIso: endIso }, '*'); } catch (e) {}
        }
      } catch (e) {}
      return;
    }
    // Instead of using browser notifications (which may be blocked), highlight the
    // autosave timer bar in red and display a clear reminder in the indicator.
    const notifyNow = (forceUIOnly) => {
      const lang = (function() { try { return localStorage.getItem(REMINDER_LANG_KEY) || 'sv'; } catch (e) { return 'sv'; } })();
      const title = lang === 'en' ? 'Time report reminder' : 'Tidrapport påminnelse';
      const body = lang === 'en' ? 'Today is the last day of the period — submit your time report.' : 'Idag är sista dagen för perioden – skicka in din tidrapport.';

        try {
          // Update the indicator label/subtext and add an explicit explanation
          try {
            const explanation = lang === 'en' ? 'Red = today is the last day — submit your time report.' : 'Röd = idag är sista dagen — skicka in din tidrapport.';
            try { setIndicator('pending', title, `${body} • ${explanation}`); } catch (e) {}
          } catch (e) {}

          // Ensure the timer bar exists and set it to a red color to indicate urgency.
          // Also set the animation duration to match the current autosave timer remaining time.
          try {
            const bar = ensureTimerBar();
            bar.style.backgroundColor = '#d9534f';
            bar.style.boxShadow = '0 0 6px rgba(217,83,79,0.6)';
            try {
              // Use the same width-based timer as normal mode so behavior matches
              resetTimerBar(getTimerRemainingMs());
            } catch (e) {}
          } catch (e) {}

          // Add a persistent visual marker on the indicator element
          try {
            const indicator = ensureIndicator();
            indicator.classList.add('agresso-period-end');
            // Apply inline styles to ensure visual highlight even if CSS didn't load
            try {
              indicator.style.border = '2px solid rgba(217,83,79,0.9)';
              indicator.style.background = 'linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95))';
              const labelEl = indicator.querySelector('.agresso-autosave-label'); if (labelEl) labelEl.style.color = '#fff';
              const subEl = indicator.querySelector('.agresso-autosave-sub'); if (subEl) subEl.style.color = '#ffecec';
            } catch (e) {}
            try { highlightStatusField(true); } catch (e) {}
            
          } catch (e) {}

            // Also attempt to update the top-level document's indicator so the
            // visual highlight is visible when this script runs inside a frame.
            try {
              if (window.top && window.top.document) {
                try {
                  const topDoc = window.top.document;
                  const topInd = topDoc.getElementById(INDICATOR_ID);
                  if (topInd) {
                    topInd.classList.add('agresso-period-end');
                    try { topInd.classList.remove('agresso-saving'); } catch (e) {}
                    try { topInd.classList.remove('agresso-saved'); } catch (e) {}
                    try { topInd.classList.add('agresso-pending'); } catch (e) {}
                    try {
                      const lbl = topInd.querySelector('.agresso-autosave-label');
                      if (lbl) lbl.textContent = title;
                      const subEl = topInd.querySelector('.agresso-autosave-sub');
                      if (subEl) subEl.textContent = body;
                    } catch (e) {}
                    try {
                      const bar = topInd.querySelector('.agresso-autosave-timer');
                      if (bar) {
                        try { bar.classList.add('agresso-period-moving'); try { resetTimerBar(getTimerRemainingMs()); } catch (e2) {} } catch (e) {}
                        bar.style.backgroundColor = '#d9534f';
                        bar.style.boxShadow = '0 0 6px rgba(217,83,79,0.6)';
                      }
                    } catch (e) {}
                    try {
                      topInd.style.border = '2px solid rgba(217,83,79,0.9)';
                      topInd.style.background = 'linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95))';
                    } catch (e) {}
                    
                  }
                } catch (e) {}
              }
            } catch (e) {}

          // Immediately refresh subtext with current Status and start periodic refresh
          try { refreshPeriodIndicatorStatus(); } catch (e) {}
          try {
            if (periodStatusRefreshTimer) clearInterval(periodStatusRefreshTimer);
            periodStatusRefreshTimer = window.setInterval(refreshPeriodIndicatorStatus, 2000);
          } catch (e) {}
          // No native system notification — keep visual indicator coloring only.
        } catch (e) {
          // ignore
        }
        try {
          // Inject a forcing CSS override into both current and top documents
              const injectStyle = (doc) => {
            try {
              if (!doc) return;
              const existing = doc.getElementById('agresso-period-end-style');
              if (existing) return;
              const s = doc.createElement('style');
              s.id = 'agresso-period-end-style';
                  s.textContent = `#${INDICATOR_ID} { border: 2px solid rgba(217,83,79,0.9) !important; background: linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95)) !important; }
                #${INDICATOR_ID} .agresso-autosave-timer { height: 6px !important; display: block !important; background-color: #d9534f !important; box-shadow: 0 0 6px rgba(217,83,79,0.6) !important; transform-origin: left !important; }
                #${INDICATOR_ID} .agresso-autosave-label, #${INDICATOR_ID} .agresso-autosave-sub { color: #fff !important; }
                /* No CSS animation here; progress is driven via inline width transition from JS */
                #${INDICATOR_ID} .agresso-autosave-timer.agresso-period-moving { /* uses JS width transition */ }
              `;
              (doc.head || doc.body || doc.documentElement).appendChild(s);
            } catch (e) {}
          };

          try { injectStyle(document); } catch (e) {}
          try { if (window.top && window.top.document && window.top !== window) injectStyle(window.top.document); } catch (e) {}

          // Enforce visual highlight on both current and top documents. Do this
          // immediately and at a few short delays to beat any UI updates that
          // would otherwise remove the styling/classes.
          const enforceHighlight = (doc) => {
            try {
              if (!doc) return;
              const ind = doc.getElementById && doc.getElementById(INDICATOR_ID);
              if (!ind) return;
              try { ind.classList.add('agresso-period-end'); } catch (e) {}
              try { ind.classList.remove('agresso-saving'); } catch (e) {}
              try { ind.classList.remove('agresso-saved'); } catch (e) {}
              try { ind.classList.add('agresso-pending'); } catch (e) {}
              try { ind.style.border = '2px solid rgba(217,83,79,0.9)'; } catch (e) {}
              try { ind.style.background = 'linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95))'; } catch (e) {}
                  try {
                const bar = ind.querySelector && ind.querySelector('.agresso-autosave-timer');
                if (bar) {
                  try { bar.classList.add('agresso-period-moving'); try { resetTimerBar(getTimerRemainingMs()); } catch (e2) {} } catch (e) {}
                  try { bar.style.backgroundColor = '#d9534f'; } catch (e) {}
                  try { bar.style.boxShadow = '0 0 6px rgba(217,83,79,0.6)'; } catch (e) {}
                }
              } catch (e) {}
            } catch (e) {}
          };

          try { enforceHighlight(document); } catch (e) {}
          try { if (window.top && window.top.document && window.top !== window) enforceHighlight(window.top.document); } catch (e) {}
          [100, 500, 1500, 3000].forEach((ms) => {
            try { window.setTimeout(() => { try { enforceHighlight(document); } catch (e) {} try { if (window.top && window.top.document && window.top !== window) enforceHighlight(window.top.document); } catch (e) {} }, ms); } catch (e) {}
          });

          // Start a persistent enforcer that reapplies highlight until the
          // stored notify key changes or the user acknowledges the period.
          try {
            // Helper to show a small persistent banner prompting submission
            const createSubmitBanner = (doc, titleText, bodyText) => {
              try {
                if (!doc || !doc.body) return;
                if (doc.getElementById('agresso-period-banner')) return;
                const ban = doc.createElement('div');
                ban.id = 'agresso-period-banner';
                ban.style.position = 'fixed';
                ban.style.right = '16px';
                ban.style.bottom = '20px';
                ban.style.zIndex = '9999999';
                ban.style.padding = '18px 20px';
                ban.style.background = 'linear-gradient(180deg, rgba(217,83,79,0.95), rgba(181,62,62,0.95))';
                ban.style.color = '#fff';
                ban.style.borderRadius = '10px';
                ban.style.boxShadow = '0 6px 18px rgba(0,0,0,0.35)';
                ban.style.fontSize = '15px';
                ban.style.display = 'flex';
                ban.style.flexDirection = 'column';
                ban.style.alignItems = 'flex-start';
                ban.style.gap = '10px';
                ban.style.maxWidth = '420px';
                // Explanation line on top (bold)
                const expl = doc.createElement('div');
                expl.style.fontWeight = '700';
                expl.style.fontSize = '15px';
                expl.textContent = titleText || 'SISTA DAGEN I PERIODEN';
                // Regular message below
                const txt = doc.createElement('div');
                txt.style.maxWidth = '360px';
                txt.style.fontSize = '13px';
                txt.textContent = bodyText || 'Submit your time report today.';
                const controls = doc.createElement('div');
                controls.style.display = 'flex';
                controls.style.gap = '8px';
                // Primary action
                const btn = doc.createElement('button');
                btn.textContent = (titleText || 'Open report');
                btn.style.background = '#fff';
                btn.style.color = '#b02a2a';
                btn.style.border = 'none';
                btn.style.padding = '8px 10px';
                btn.style.borderRadius = '6px';
                btn.style.cursor = 'pointer';
                btn.addEventListener('click', (ev) => {
                  ev.stopPropagation(); ev.preventDefault();
                  try {
                    const sb = findSaveButton();
                    if (sb) {
                      try { sb.scrollIntoView({ behavior: 'smooth' }); } catch (e) {}
                      try { sb.focus(); } catch (e) {}
                    }
                  } catch (e) {}
                }, true);
                controls.appendChild(btn);
                ban.appendChild(expl);
                ban.appendChild(txt);
                ban.appendChild(controls);
                (doc.body || doc.documentElement).appendChild(ban);
              } catch (e) {}
            };

            const removeSubmitBanner = (doc) => {
              try {
                if (!doc) return;
                const ex = doc.getElementById('agresso-period-banner');
                if (ex && ex.parentNode) ex.parentNode.removeChild(ex);
              } catch (e) {}
            };

            // Diagnostic: compare timer bar behavior in normal vs highlighted mode.
            const compareTimerModes = () => {
              return new Promise((resolve) => {
                try {
                  const doc = document;
                  const bar = doc.querySelector('.agresso-autosave-timer');
                  if (!bar) return resolve({ error: 'no-timer-bar' });

                  const sample = () => {
                    const cs = window.getComputedStyle(bar);
                    return {
                      classList: Array.from(bar.classList),
                      animationName: cs.animationName,
                      animationDuration: cs.animationDuration,
                      transform: cs.transform,
                      width: cs.width,
                      inlineWidth: bar.style.width || null,
                      inlineTransform: bar.style.transform || null,
                      inlineTransition: bar.style.transition || null
                    };
                  };

                  // baseline
                  const baseline = sample();
                  const hadMoving = bar.classList.contains('agresso-period-moving');
                  const hadMarker = (bar.closest('.agresso-period-end') != null) || (document.querySelector('.agresso-period-end') != null);

                  // apply highlighted mode
                  bar.classList.add('agresso-period-moving');
                  // also ensure parent indicator has period-end marker
                  const parentIndicator = bar.closest('.agresso-enabled, .agresso-disabled, .agresso-indicator') || document.body;
                  parentIndicator.classList.add('agresso-period-end');

                  // allow styles to settle
                  setTimeout(() => {
                    const highlighted = sample();
                    // revert to previous state
                    if (!hadMoving) bar.classList.remove('agresso-period-moving');
                    if (!hadMarker) parentIndicator.classList.remove('agresso-period-end');
                    resolve({ baseline, highlighted });
                  }, 110);
                } catch (e) { resolve({ error: String(e) }); }
              });
            };
            try { if (periodHighlightEnforcer) { clearInterval(periodHighlightEnforcer); periodHighlightEnforcer = null; } } catch (e) {}
            periodHighlightEnforcer = window.setInterval(() => {
              try { enforceHighlight(document); } catch (e) {}
              try { if (window.top && window.top.document && window.top !== window) enforceHighlight(window.top.document); } catch (e) {}
              // Stop if notify key no longer matches or ack equals endIso
              let topNotify = null; try { topNotify = (window.top && window.top.localStorage) ? window.top.localStorage.getItem(PERIOD_NOTIFY_KEY) : null; } catch (e) { topNotify = null; }
              let localNotify = null; try { localNotify = localStorage.getItem(PERIOD_NOTIFY_KEY); } catch (e) { localNotify = null; }
              if ((topNotify !== endIso && localNotify !== endIso)) {
                try { clearInterval(periodHighlightEnforcer); periodHighlightEnforcer = null; } catch (e) {}
                try { removeSubmitBanner(document); } catch (e) {}
                try { if (window.top && window.top.document && window.top !== window) removeSubmitBanner(window.top.document); } catch (e) {}
              }
            }, 1000);
          } catch (e) {}

          if (!forceUIOnly) {
            try {
              if (window.top && window.top.localStorage) {
                window.top.localStorage.setItem(PERIOD_NOTIFY_KEY, endIso);
              } else {
                localStorage.setItem(PERIOD_NOTIFY_KEY, endIso);
              }
            } catch (e) {
              try { localStorage.setItem(PERIOD_NOTIFY_KEY, endIso); } catch (e2) {}
            }
            // Create a persistent banner prompting submission (current + top)
            try { createSubmitBanner(document, explanation, `${title} — ${body}`); } catch (e) {}
            try { if (window.top && window.top.document && window.top !== window) createSubmitBanner(window.top.document, explanation, `${title} — ${body}`); } catch (e) {}
            console.info(LOG_PREFIX, 'showPeriodNotification: stored PERIOD_NOTIFY_KEY', endIso);
          } else {
            try { console.info(LOG_PREFIX, 'showPeriodNotification: UI-only enforcement, not storing key'); } catch (e) {}
          }
          // Always attempt to create the persistent submit banner so users
          // get a visible prompt even when we only enforce UI (already-notified path).
          try { createSubmitBanner(document, explanation, `${title} — ${body}`); } catch (e) {}
          try { if (window.top && window.top.document && window.top !== window) createSubmitBanner(window.top.document, explanation, `${title} — ${body}`); } catch (e) {}
        } catch (e) {}
      };

    // Fire after a short timeout so init tasks finish first
    window.setTimeout(notifyNow, 200);
  }

  function checkPeriodAndNotify(context) {
    try {
      // Helper: normalize to start of day for comparisons
      const startOfDay = (d) => { const x = new Date(d); x.setHours(0,0,0,0); return x; };
      const todayStart = startOfDay(new Date());

      // Respect manual override first
      const override = getOverrideDate();
      if (override) {
        console.info(LOG_PREFIX, 'checkPeriodAndNotify: using override', override);
        try {
          const overrideStart = startOfDay(override);
          // Notify if the override date is today or earlier (period due/overdue).
          if (overrideStart.getTime() <= todayStart.getTime()) {
                if (!isReportStatusKlar()) {
                  showPeriodNotification(override);
                  console.info(LOG_PREFIX, 'checkPeriodAndNotify: override is today-or-earlier, notified');
                  return true;
                } else {
                  console.info(LOG_PREFIX, 'checkPeriodAndNotify: override is today-or-earlier but status Klar, not notifying');
                  return false;
                }
              }
              console.info(LOG_PREFIX, 'checkPeriodAndNotify: override is after today');
        } catch (e) {
          console.info(LOG_PREFIX, 'checkPeriodAndNotify: override parsing error', e);
        }
        return false;
      }

      const end = findPeriodEndDate();
      if (!end) {
        console.info(LOG_PREFIX, 'checkPeriodAndNotify: no end date found');
        return false;
      }
      const endStart = startOfDay(end);
      const today = new Date();
      console.info(LOG_PREFIX, 'checkPeriodAndNotify: found end date', localIsoDate(end) || end.toISOString().slice(0,10), 'today', localIsoDate(today) || today.toISOString().slice(0,10));

      // Notify when the period end is today or earlier (due/overdue), but only until Status becomes 'Klar'
      if (endStart.getTime() <= todayStart.getTime()) {
        if (!isReportStatusKlar()) {
          showPeriodNotification(end);
          // Force the GUI timer bar to red when not 'Klar'
          try {
            const applyRedUI = (doc) => {
              try {
                const d = doc || document;
                const ind = d.getElementById && d.getElementById(INDICATOR_ID);
                if (ind) {
                  try { ind.classList.add('agresso-period-end'); } catch (e) {}
                  try { ind.classList.remove('agresso-saving'); } catch (e) {}
                  try { ind.classList.remove('agresso-saved'); } catch (e) {}
                  try { ind.classList.add('agresso-pending'); } catch (e) {}
                  try { ind.style.border = '2px solid rgba(217,83,79,0.9)'; } catch (e) {}
                  try { ind.style.background = 'linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95))'; } catch (e) {}
                  try {
                    const bar = ind.querySelector && ind.querySelector('.agresso-autosave-timer');
                    if (bar) {
                      try { bar.classList.add('agresso-period-moving'); try { resetTimerBar(getTimerRemainingMs()); } catch (e2) {} } catch (e) {}
                      bar.style.backgroundColor = '#d9534f';
                      bar.style.boxShadow = '0 0 6px rgba(217,83,79,0.6)';
                    }
                  } catch (e) {}
                }
                try {
                    const existing = (d.getElementById && d.getElementById('agresso-period-end-style')) || null;
                  if (!existing) {
                    const s = d.createElement('style');
                    s.id = 'agresso-period-end-style';
                    s.textContent = `#${INDICATOR_ID} { border: 2px solid rgba(217,83,79,0.9) !important; background: linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95)) !important; } #${INDICATOR_ID} .agresso-autosave-timer { height: 6px !important; display: block !important; background-color: #d9534f !important; box-shadow: 0 0 6px rgba(217,83,79,0.6) !important; transform-origin: left !important; } #${INDICATOR_ID} .agresso-autosave-label, #${INDICATOR_ID} .agresso-autosave-sub { color: #fff !important; } /* No CSS animation; JS width transition controls progress */ #${INDICATOR_ID} .agresso-autosave-timer.agresso-period-moving { }
`;
                    (d.head || d.body || d.documentElement).appendChild(s);
                  }
                } catch (e) {}
              } catch (e) {}
            };
            try { applyRedUI(document); } catch (e) {}
            try { if (window.top && window.top.document && window.top !== window) applyRedUI(window.top.document); } catch (e) {}
          } catch (e) {}

          console.info(LOG_PREFIX, 'checkPeriodAndNotify: end date is today-or-earlier, notified');
          return true;
        }
        console.info(LOG_PREFIX, 'checkPeriodAndNotify: end date is today-or-earlier but status Klar, not notifying');
        return false;
      }
      console.info(LOG_PREFIX, 'checkPeriodAndNotify: end date is after today');
    } catch (e) {
      // ignore
    }
    return false;
  }
  // Page-context helper injection removed due to site Content Security Policy (CSP).
  // Use the content-script debug button or dispatch the DOM event `agresso_check_period`
  // (e.g. `document.dispatchEvent(new Event('agresso_check_period'))`) to trigger checks safely.

  function onFieldInput(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (!target.matches('input, textarea, select')) {
      return;
    }

    if (target.matches('input[type="checkbox"]') && !target.hasAttribute('data-fieldname')) {
      return; // ignore row select checkboxes
    }

    const row = target.closest('tr');
    if (!row) {
      return;
    }

    row.dataset.agressoDirty = '1';
    pendingRow = row;
    markActivity();
  }

  function onFieldBlur(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (!target.matches('input, textarea, select')) {
      return;
    }

    const row = target.closest('tr');
    if (!row) {
      return;
    }

    const next = event.relatedTarget;
    if (next && row.contains(next)) {
      return; // still inside the same row
    }

    if (row.dataset.agressoDirty === '1') {
      pendingRow = row;
    }

    markActivity();
  }

  function onDropdownOpen(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (target.matches('select')) {
      dropdownActive = true;
      dropdownRow = target.closest('tr') || dropdownRow;
    }

    markActivity();
  }

  function onDropdownClose(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (target.matches('select')) {
      dropdownActive = false;
      const row = target.closest('tr') || dropdownRow;
      const isDirty = row && row.dataset.agressoDirty === '1';
      dropdownRow = null;
      if (row && (isDirty || pendingRow === row)) {
        pendingRow = row;
      }

      markActivity();
    }
  }

  function onDeleteClick(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) {
      return;
    }

    if (isDeletionButton(target)) {
      const row = target.closest('tr');
      pendingRow = row || pendingRow;
      markActivity();
      return;
    }

    if (isAddRowButton(target)) {
      markActivity();
    }
  }

  function initObservers() {
    const observer = new MutationObserver((records) => {
      let rowsRemoved = false;
      let dialogAdded = false;
      for (const rec of records) {
        if (rec.removedNodes && rec.removedNodes.length) {
          rowsRemoved = rowsRemoved || Array.from(rec.removedNodes).some((n) => n.nodeName === 'TR');
        }
        if (rec.addedNodes && rec.addedNodes.length) {
          dialogAdded = dialogAdded || Array.from(rec.addedNodes).some((n) => {
            if (!(n instanceof HTMLElement)) {
              return false;
            }
            const roleDialog = n.getAttribute('role') === 'dialog';
            const modalClass = n.classList.contains('modal') || n.classList.contains('k-window');
            const alertClass = n.classList.contains('alert') || n.classList.contains('notification');
            return roleDialog || modalClass || alertClass;
          });
        }
      }
      if (rowsRemoved) {
        markActivity();
      }
      if (dialogAdded) {
        startDialogSweep('dialog added');
      }
      refreshNoChangesBannerState('mutation');
      scheduleLayoutRefresh();
        // Check whether today is the last day in the currently shown period and notify once.
        // Run this after a short debounce so transient DOM swaps during navigation
        // don't cause false negatives/positives.
        try {
          if (periodStatusRefreshTimer) clearTimeout(periodStatusRefreshTimer);
          periodStatusRefreshTimer = window.setTimeout(() => {
            try { checkPeriodAndNotify('mutation'); } catch (e) { /* ignore */ }
            periodStatusRefreshTimer = null;
          }, 300);
        } catch (e) { /* ignore */ }
      bindIndicatorTracking();
      bindActivityListeners();
    });
    // Observe the document root (documentElement) rather than `body` so the
    // observer stays active if the page replaces or re-creates <body> during
    // client-side navigation (common on SPA-like pages).
    const rootNode = document.documentElement || document.body;
    try {
      observer.observe(rootNode, { childList: true, subtree: true });
    } catch (e) {
      // fallback to observing body if documentElement isn't available for some reason
      try { observer.observe(document.body, { childList: true, subtree: true }); } catch (err) { /* ignore */ }
    }
  }

  function init() {
    enhanceLayout();
    console.info(LOG_PREFIX, 'Init', { isMac: IS_MAC, shortcut: SHORTCUT_LABEL });
    // Ensure autosave toggle always starts enabled
    try {
      setToggleEnabled(true);
    } catch (e) {
      // ignore
    }

    // If this is not the top-level frame, do not start timers or perform saves here.
    // Non-top frames will forward activity events to the top frame via postMessage.
    const isTop = (() => {
      try { return window.top === window; } catch (e) { return false; }
    })();

    if (!isTop) {
      // Minimal init for frames: layout tweaks and activity listeners only.
      document.addEventListener('input', onFieldInput, true);
      document.addEventListener('blur', onFieldBlur, true);
      document.addEventListener('click', onDeleteClick, true);
      document.addEventListener('focusin', onDropdownOpen, true);
      document.addEventListener('change', onDropdownClose, true);
      document.addEventListener('focusout', onDropdownClose, true);
      bindActivityListeners();
      scheduleLayoutRefresh();
      return;
    }

    // Top-level frame: full behavior
    setIndicator('saved', 'Autosave ready', 'Watching for edits');
    // Check whether today is the last day in the currently shown period and notify once
    try { checkPeriodAndNotify('init'); } catch (e) { /* ignore */ }
    

    // Attach field event listeners
    document.addEventListener('input', onFieldInput, true);
    document.addEventListener('blur', onFieldBlur, true);
    document.addEventListener('click', onDeleteClick, true);
    document.addEventListener('focusin', onDropdownOpen, true);
    document.addEventListener('change', onDropdownClose, true);
    document.addEventListener('focusout', onDropdownClose, true);

    bindActivityListeners();

    // Start idle timer on init
    lastActivityAt = Date.now();
    startTimer(IDLE_TIMEOUT_MS, 'idle');

    // Kick off a short sweep at load in case a dialog is already present.
    startDialogSweep('init');

    refreshNoChangesBannerState('init');

    positionIndicatorNearSaveButton();
    bindIndicatorTracking();

    if (!noChangesPollTimer) {
      noChangesPollTimer = window.setInterval(() => {
        refreshNoChangesBannerState('poll');
      }, NO_CHANGES_POLL_MS);
    }
    
    // Initialize mutation observer
    initObservers();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  // Allow page to request a manual check by dispatching a DOM event (works despite CSP)
  try {
    document.addEventListener('agresso_check_period', (ev) => {
      try {
        const res = checkPeriodAndNotify('page-event');
        console.info(LOG_PREFIX, 'agresso_check_period handler result', res);
        try {
          const indicator = ensureIndicator();
          if (indicator) {
            indicator.dataset.lastPeriodCheck = JSON.stringify({ result: !!res, ts: new Date().toISOString() });
          }
        } catch (e) {}
      } catch (e) {
        // ignore
      }
    }, false);
  } catch (e) {
    // ignore
  }

  // Listen for activity messages from child frames and treat them as activity
  try {
    if (window.top === window) {
      window.addEventListener('message', (ev) => {
        try {
          if (ev && ev.data && ev.data.type === ACTIVITY_MESSAGE) {
            // Mark activity in the top frame
            markActivity();
              return;
            }
            // Handle requests from child frames to enforce a period highlight
            if (ev && ev.data && ev.data.type === 'agresso_period_enforce') {
              try {
                const endIso = ev.data && ev.data.endIso;
                // Apply enforcement locally in the top frame
                try {
                  const indicator = ensureIndicator();
                  if (indicator) {
                    indicator.classList.add('agresso-period-end');
                    try { indicator.classList.remove('agresso-saving'); } catch (e) {}
                    try { indicator.classList.remove('agresso-saved'); } catch (e) {}
                    try { indicator.classList.add('agresso-pending'); } catch (e) {}
                    try { indicator.style.border = '2px solid rgba(217,83,79,0.9)'; } catch (e) {}
                    try { indicator.style.background = 'linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95))'; } catch (e) {}
                    try {
                      const bar = indicator.querySelector('.agresso-autosave-timer');
                      if (bar) {
                        try { bar.classList.add('agresso-period-moving'); try { resetTimerBar(getTimerRemainingMs()); } catch (e2) {} } catch (e) {}
                        bar.style.backgroundColor = '#d9534f';
                        bar.style.boxShadow = '0 0 6px rgba(217,83,79,0.6)';
                      }
                    } catch (e) {}
                  }
                } catch (e) {}

                // Inject forcing style into top doc if missing
                try {
                  const doc = document;
                  const existing = doc.getElementById('agresso-period-end-style');
                  if (!existing) {
                    const s = doc.createElement('style');
                    s.id = 'agresso-period-end-style';
                    s.textContent = `#${INDICATOR_ID} { border: 2px solid rgba(217,83,79,0.9) !important; background: linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95)) !important; } #${INDICATOR_ID} .agresso-autosave-timer { height: 6px !important; display: block !important; background-color: #d9534f !important; box-shadow: 0 0 6px rgba(217,83,79,0.6) !important; transform-origin: left !important; } #${INDICATOR_ID} .agresso-autosave-label, #${INDICATOR_ID} .agresso-autosave-sub { color: #fff !important; } /* No CSS animation; JS width transition controls progress */ #${INDICATOR_ID} .agresso-autosave-timer.agresso-period-moving { }
`;
                    (doc.head || doc.body || doc.documentElement).appendChild(s);
                  }
                } catch (e) {}

                // Start a persistent enforcer in top if not already running
                try {
                  if (periodHighlightEnforcer) { clearInterval(periodHighlightEnforcer); periodHighlightEnforcer = null; }
                  periodHighlightEnforcer = window.setInterval(() => {
                    try {
                      const ind2 = ensureIndicator();
                      if (ind2) {
                        ind2.classList.add('agresso-period-end');
                        ind2.classList.remove('agresso-saving');
                        ind2.classList.remove('agresso-saved');
                        ind2.classList.add('agresso-pending');
                        try { ind2.style.border = '2px solid rgba(217,83,79,0.9)'; } catch (e) {}
                        try { ind2.style.background = 'linear-gradient(180deg, rgba(36,41,50,0.95), rgba(30,25,28,0.95))'; } catch (e) {}
                        try { const bar = ind2.querySelector('.agresso-autosave-timer'); if (bar) { try { bar.classList.add('agresso-period-moving'); try { resetTimerBar(getTimerRemainingMs()); } catch (e2) {} } catch (e) {} bar.style.backgroundColor = '#d9534f'; bar.style.boxShadow = '0 0 6px rgba(217,83,79,0.6)'; } } catch (e) {}
                      }
                    } catch (e) {}
                    // Stop if notify key changed or ack set
                    try {
                      const topN = (window.top && window.top.localStorage) ? window.top.localStorage.getItem(PERIOD_NOTIFY_KEY) : null;
                      const locN = localStorage.getItem(PERIOD_NOTIFY_KEY);
                      if ((topN !== endIso && locN !== endIso)) {
                        try { clearInterval(periodHighlightEnforcer); periodHighlightEnforcer = null; } catch (e) {}
                      }
                    } catch (e) {}
                  }, 1000);
                } catch (e) {}
              } catch (e) {}
              return;
          }
        } catch (e) {
          // ignore malformed messages
        }
      }, false);
    }
  } catch (e) {
    // ignore cross-origin
  }
  try {
    try { window.agresso_buildDebugReport = buildDebugReport; } catch (e) {}
    try { window.agresso_compareTimerModes = compareTimerModes; } catch (e) {}
    try {
      window.agresso_setIndicatorDebug = function(enabled) {
        try { INDICATOR_DEBUG = !!enabled; } catch (e) {}
        try { console.info(LOG_PREFIX, 'agresso_setIndicatorDebug =>', INDICATOR_DEBUG); } catch (e) {}
      };
    } catch (e) {}
  } catch (e) {}
})();
