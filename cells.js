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
  const NO_CHANGES_TEXT = 'inga ändringar gjorda!';
  const SAVE_DIALOG_SELECTORS = [
    '[id^="u4_messageoverlay_success"]',
    '.u4-messageoverlay-success-header',
    '.u4-messageoverlay-success-body',
    '.u4-messageoverlay-success-footer'
  ];
  const SAVE_DIALOG_KEYWORDS = [
    'spara',
    'sparats',
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
  let dropdownActive = false;
  let dropdownRow = null;
  let lastActivityAt = Date.now();
  let dialogMissLogged = false;
  let fallbackMissLogged = false;
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

  const INDICATOR_ID = 'agresso-autosave-indicator';
  const OK_LABELS = ['ok', 'stäng', 'close', 'oké'];
  const CLOSE_SELECTORS = ['[aria-label="Close"]', '.close', '.k-i-close', '.modal-close'];
  const ACTIVITY_MESSAGE = 'agresso-autosave-activity';

  function getIndicatorDocument() {
    try {
      if (window.top && window.top.document && window.top.document.body) {
        return window.top.document;
      }
    } catch (e) {
      // Ignore cross-origin access errors and fall back to current document.
    }
    return document;
  }

  function applyFieldSizing() {
    FIELD_STYLE_RULES.forEach(({ selector, style }) => {
      document.querySelectorAll(selector).forEach((node) => {
        node.setAttribute('style', style);
      });
    });
  }

  function addProjectLabels() {
    document.querySelectorAll('td').forEach((cell) => {
      const title = cell.getAttribute('title');
      if (!title || cell.querySelector('.agresso-project-label')) {
        return;
      }

      const headerRow = cell.parentNode?.parentNode?.parentNode;
      const cellIndex = cell.cellIndex;
      if (!headerRow) {
        return;
      }

      const headers = headerRow.getElementsByTagName('th');
      const header = headers[cellIndex];
      const fieldName = header?.getAttribute('data-fieldname');

      if (fieldName === 'work_order' || fieldName === 'project') {
        const projectName = title.split('-')[0]?.trim();
        if (!projectName) {
          return;
        }

        const para = document.createElement('p');
        para.className = 'agresso-project-label';
        para.setAttribute(
          'style',
          'font-size: 10px; margin-top: 1px; overflow: hidden; text-overflow: ellipsis;'
        );
        para.textContent = projectName;
        cell.appendChild(para);
      }
    });
  }

  function enhanceLayout() {
    applyFieldSizing();
    addProjectLabels();
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

    const dot = indicatorDoc.createElement('span');
    dot.className = 'agresso-autosave-dot';

    const label = indicatorDoc.createElement('span');
    label.className = 'agresso-autosave-label';

    const sub = indicatorDoc.createElement('span');
    sub.className = 'agresso-autosave-sub';

    indicator.appendChild(dot);
    indicator.appendChild(label);
    indicator.appendChild(sub);

    indicatorDoc.body.appendChild(indicator);
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
    const btn = findPrimarySaveButton();
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
    const viewportWidth = indicator.ownerDocument?.defaultView?.innerWidth || window.innerWidth || 0;
    const mirroredLeft = viewportWidth ? viewportWidth - rect.left - indWidth : rect.right + 12;

    indicator.style.position = 'fixed';
    indicator.style.top = `${Math.max(8, top)}px`;
    indicator.style.left = `${Math.max(8, mirroredLeft)}px`;
    indicator.style.right = 'auto';
    indicator.style.bottom = 'auto';
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
      indicator.appendChild(bar);
    }
    timerBar = bar;
    return bar;
  }

  function resetTimerBar(durationMs) {
    const bar = ensureTimerBar();
    bar.style.transition = 'none';
    bar.style.width = '0%';
    // force reflow
    // eslint-disable-next-line no-unused-expressions
    bar.offsetWidth;
    bar.style.transition = `width ${durationMs}ms linear`;
    bar.style.width = '100%';
  }

  function stopTimerBar() {
    const bar = timerBar || ensureTimerBar();
    bar.style.transition = 'none';
    bar.style.width = '0%';
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

    positionIndicatorNearSaveButton();
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
        return text.includes(NO_CHANGES_TEXT);
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
        window.top.postMessage({ type: ACTIVITY_MESSAGE, ts: Date.now() }, '*');
        return;
      }
    } catch (e) {
      // ignore cross-origin access; fall through to local handling
    }

    lastActivityAt = Date.now();
    // Always restart timer on activity
    startTimer(IDLE_TIMEOUT_MS, 'idle');
  }

  function performSave(trigger) {
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

    // Silent fallback: if the page ignores the shortcut, attempt a button click shortly after.
    window.setTimeout(() => {
      clickSaveButtonsFallback();
    }, 250);

    lastSaveAt = Date.now();
    startDialogSweep('autosave');

    window.setTimeout(() => {
      const timestamp = new Date().toLocaleTimeString(undefined, {
        hour12: false,
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      });
      setIndicator('saved', 'Saved', `at ${timestamp}`);
      // Restart idle countdown after a save completes
      lastActivityAt = Date.now();
      startTimer(IDLE_TIMEOUT_MS, 'idle');
      if (pendingRow) {
        pendingRow.dataset.agressoDirty = '0';
        pendingRow = null;
      }
    }, 1500);
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

  function clickSaveButtonsFallback() {
    const btn = findSaveButton();
    if (btn) {
      // reset miss log counter when we actually find a button
      fallbackMissLogged = false;
      console.info(LOG_PREFIX, 'Fallback clicking save button');
      try {
        btn.click();
      } catch (e) {
        // ignore
      }
      return true;
    }

    // Before warning, check whether anything actually looks save-related (dialog, save keywords, or dirty row).
    const docs = getAllDocuments();
    let sawSaveCandidate = false;
    try {
      for (const doc of docs) {
        if (!doc) continue;
        // Any explicit save button candidate in DOM?
        try {
          if (doc.querySelector(SAVE_BUTTON_SELECTORS.join(','))) {
            sawSaveCandidate = true;
            break;
          }
        } catch (e) {
          // ignore selector errors
        }

        // Any dialog-like element with save keywords?
        try {
          const dialog = doc.querySelector('[role="dialog"], .modal, .k-window, .notification, .alert, .k-dialog');
          if (dialog && isSaveDialog(dialog)) {
            sawSaveCandidate = true;
            break;
          }
        } catch (e) {
          // ignore
        }

        // Body text contains save-like keywords?
        try {
          const bodyText = (doc.body && (doc.body.innerText || doc.body.textContent || '')).toLowerCase();
          if (bodyText && SAVE_DIALOG_KEYWORDS.some((kw) => bodyText.includes(kw))) {
            sawSaveCandidate = true;
            break;
          }
        } catch (e) {
          // ignore
        }
      }
    } catch (e) {
      // ignore
    }

    // Also consider a dirty row present as a valid trigger for attempting a fallback click
    if (!sawSaveCandidate && getDirtyRow()) {
      sawSaveCandidate = true;
    }

    if (!sawSaveCandidate) {
      // Nothing indicating a save is pending — skip noisy warning.
      return false;
    }

    if (!fallbackMissLogged) {
      console.warn(LOG_PREFIX, 'Fallback save button not found');
      fallbackMissLogged = true;
    }
    return false;
  }

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
      bindIndicatorTracking();
      bindActivityListeners();
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  function init() {
    enhanceLayout();
    console.info(LOG_PREFIX, 'Init', { isMac: IS_MAC, shortcut: SHORTCUT_LABEL });

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

  // Listen for activity messages from child frames and treat them as activity
  try {
    if (window.top === window) {
      window.addEventListener('message', (ev) => {
        try {
          if (ev && ev.data && ev.data.type === ACTIVITY_MESSAGE) {
            // Mark activity in the top frame
            markActivity();
          }
        } catch (e) {
          // ignore malformed messages
        }
      }, false);
    }
  } catch (e) {
    // ignore cross-origin
  }
})();
