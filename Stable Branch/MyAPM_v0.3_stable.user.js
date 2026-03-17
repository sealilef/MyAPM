// ==UserScript==
// @name         MyAPM
// @namespace    https://w.amazon.com/bin/view/MLB1-RME/MyAPM/
// @version      0.3.99_stable
// @description  APM Customizer and feature enhancer
// @author       sealilef
// @match        https://us1.eam.hxgnsmartcloud.com/*
// @match        https://eu1.eam.hxgnsmartcloud.com/*
// @match        https://*.apm-es.gps.amazon.dev/*
// @match        https://*.insights.amazon.dev/*
// @homepageURL  https://github.com/sealilef/MyAPM/blob/main/Stable%20Branch/MyAPM_v0.3_stable.user.js
// @supportURL   https://github.com/sealilef/MyAPM/blob/main/Stable%20Branch/MyAPM_v0.3_stable.user.js
// @updateURL    https://raw.githubusercontent.com/sealilef/MyAPM/main/Stable%20Branch/MyAPM_v0.3_stable.user.js
// @downloadURL  https://raw.githubusercontent.com/sealilef/MyAPM/main/Stable%20Branch/MyAPM_v0.3_stable.user.js
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        unsafeWindow
// @run-at       document-end
// ==/UserScript==

(function () {
    'use strict';

    const SCRIPT_ID = 'myapm-audit-navigator-standalone';
    const TRACE = '[MyAPM][nav]';
    const NAV_DEBUG = false;
    const PAGE_WINDOW = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;
    const CURRENT_VERSION = '0.3.99_stable';
    const UPDATE_URL = 'https://raw.githubusercontent.com/sealilef/MyAPM/main/Stable%20Branch/MyAPM_v0.3_stable.user.js';
    const DOWNLOAD_URL = 'https://raw.githubusercontent.com/sealilef/MyAPM/main/Stable%20Branch/MyAPM_v0.3_stable.user.js';
    const SCRIPT_PAGE_URL = 'https://github.com/sealilef/MyAPM/raw/refs/heads/main/Stable%20Branch/MyAPM_v0.3_stable.user.js';

    const POLL_MS = 100;
    const NAV_TIMEOUT_MS = 15000;
    const READY_TIMEOUT_MS = 12000;
    const STORE_TIMEOUT_MS = 15000;
    const LINKIFY_INTERVAL_MS = 1500;
    const UPDATE_CHECK_RETRY_MS = 4000;
    const UPDATE_CHECK_MAX_ATTEMPTS = 3;

    const GRID_RESIZE_REQUEST_KEY = 'myapm_force_grid_resize_request';
    const GRID_RESIZE_RETRY_MS = 700;
    const GRID_RESIZE_RETRY_COUNT = 12;
    const CURRENT_USER_STORAGE_KEY = 'myapmCurrentUser';
    const MODAL_HOST_ID = 'myapm-results-modal-host';
    const CONTROL_Z_INDEX = '1000';
    const REORDER_PANEL_Z_INDEX = '1000';
    const SETTINGS_PANEL_Z_INDEX = '1001';
    const MODAL_Z_INDEX = '999999';
    const SETTINGS_PANEL_ID = 'myapm-settings-panel';
    const CONTROL_BAR_ID = 'myapm-control-bar';
    const REORDER_PANEL_ID = 'myapm-reorder-panel';
    const REORDER_STORAGE_KEY = 'myapmLayoutOrderV1';
    const REORDER_BUTTON_ID = 'myapm-layout-reorder-button';
    const WORKORDER_REGEX = /\b\d{11}\b/g;
    const WORKORDER_PLAIN_REGEX = /\b\d{11}\b/;
    const PTP_HISTORY_STORAGE_KEY = 'apm_ptp_history';
    const PTP_LINK_BASE = 'https://user.sparsy.insights.amazon.dev/ptp';
    const PTP_HISTORY_EVENT_NAME = 'MYAPM_PTP_HISTORY_UPDATED';
    const PTP_COMPLETE_VALID_MS = 12 * 60 * 60 * 1000;

    const DUE_WINDOW_CONFIG = {
        compliance: { modeKey: 'apmComplianceDueMode', customKey: 'apmComplianceDueCustomDays', migrationKey: 'apmComplianceDefaultsToTodayMigratedV1', label: 'Compliance PMs' },
        pms: { modeKey: 'apmPmDueMode', customKey: 'apmPmDueCustomDays', migrationKey: 'apmPmDefaultsToTodayMigratedV1', label: 'PMs' },
        fwos: { modeKey: 'apmFwoDueMode', customKey: 'apmFwoDueCustomDays', migrationKey: 'apmFwoDefaultsToTodayMigratedV1', label: 'WOs' },
        audits: { modeKey: 'apmAuditDueMode', customKey: 'apmAuditDueCustomDays', migrationKey: 'apmAuditDefaultsToTodayMigratedV1', label: 'Audits' }
    };

    const DEFAULT_COLUMN_PRIORITY = [
        { aliases: ['Priority Icon', 'priority', 'priorityicon'] },
        { aliases: ['Work Order', 'workordernum', 'workorder'] },
        { aliases: ['Description', 'description'] },
        { aliases: ['Equipment Description', 'equipmentdescription', 'equipmentdesc'] },
        { aliases: ['Equipment', 'equipment'] },
        { aliases: ['Status', 'workorderstatus', 'status'] },
        { aliases: ['Sched Start', 'Sched. Start Date', 'schedstartdate', 'schedstart'] },
        { aliases: ['Sched End', 'Sched. End Date', 'schedenddate', 'schedend'] },
        { aliases: ['Original Due Date', 'Original PM Due Date', 'duedate', 'originalduedate', 'originalpmduedate', 'audit date'] },
        { aliases: ['Type', 'workordertype', 'type'] },
        { aliases: ['Assigned To', 'assignedto'] },
        { aliases: ['Department', 'department'] }
    ];

    const DEFAULT_RECORD_TAB_PRIORITY = ['Details', 'Activities', 'Checklist', 'Comments', 'Documents', 'Attachments', 'Scheduling', 'Costs'];

    const REORDER_CONTEXTS = {
        wsjobs: { key: 'wsjobs', label: 'Work Orders Grid', kind: 'columns', emptyText: 'No Work Orders grid found. Open Work Orders.' },
        compliance: { key: 'compliance', label: 'Compliance Grid', kind: 'columns', emptyText: 'No Compliance grid found. Open Compliance.' },
        audits: { key: 'audits', label: 'Audits Grid', kind: 'columns', emptyText: 'No Audits grid found. Open Audits.' },
        recordTabs: { key: 'recordTabs', label: 'Record Tabs', kind: 'tabs', emptyText: 'No record tabs found. Open a record view.' }
    };

    const reorderState = {
        contexts: {}
    };

    function startOfDay(date) {
        const d = new Date(date);
        d.setHours(0, 0, 0, 0);
        return d;
    }

    function addDays(date, days) {
        const d = new Date(date);
        d.setDate(d.getDate() + Number(days || 0));
        return d;
    }

    function clampDueDays(value) {
        const parsed = parseInt(String(value || '').trim(), 10);
        if (Number.isNaN(parsed)) return 7;
        return Math.max(0, Math.min(365, parsed));
    }

    function ensureDueWindowDefaults() {
        Object.values(DUE_WINDOW_CONFIG).forEach((config) => {
            if (!localStorage.getItem(config.modeKey)) localStorage.setItem(config.modeKey, 'today');
            if (!localStorage.getItem(config.customKey)) localStorage.setItem(config.customKey, '7');
            if (!localStorage.getItem(config.migrationKey)) {
                const currentMode = (localStorage.getItem(config.modeKey) || '').toLowerCase();
                if (!currentMode || currentMode === '7') localStorage.setItem(config.modeKey, 'today');
                localStorage.setItem(config.migrationKey, 'true');
            }
        });
    }

    function getDueWindowState(flowKey) {
        const config = DUE_WINDOW_CONFIG[flowKey];
        if (!config) return { mode: 'today', customDays: 7, daysAhead: 0, label: 'Today' };
        ensureDueWindowDefaults();
        const mode = (localStorage.getItem(config.modeKey) || 'today').toLowerCase();
        const customDays = clampDueDays(localStorage.getItem(config.customKey) || '7');
        const daysAhead = mode === 'today' ? 0 : mode === 'custom' ? customDays : 7;
        const label = mode === 'today' ? 'Today' : mode === 'custom' ? `Custom (${customDays} days)` : '7 Days';
        return { mode, customDays, daysAhead, label, config };
    }

    function filterRowsByDueWindow(rows, flowKey) {
        const state = getDueWindowState(flowKey);
        const cutoff = startOfDay(addDays(new Date(), state.daysAhead));
        const filtered = (Array.isArray(rows) ? rows : []).filter((row) => {
            const due = row && row.dueDate instanceof Date ? startOfDay(row.dueDate) : null;
            return !!due && due <= cutoff;
        });
        return { rows: filtered, dueWindow: state };
    }

    const FLOWS = {
        audits: {
            key: 'audits',
            buttonId: 'myapm-audit-run-btn',
            buttonLabel: 'Check Audits',
            runningLabel: 'Check Audits (Running...)',
            systemFunction: 'WSJOBS',
            userFunction: 'ADJOBS',
            launchTarget: 'WSJOBS?USER_FUNCTION_NAME=ADJOBS',
            titleHints: ['rme audit', 'audits'],
            gridStoreIncludes: ['audit', 'adjobs'],
            gridMarkers: ['rme audit', 'audit date'],
            modalTitle: 'Audit Results',
            dueFilterLabels: ['Audit Date']
        },
        compliance: {
            key: 'compliance',
            buttonId: 'myapm-compliance-run-btn',
            buttonLabel: 'Check Compliance PMs',
            runningLabel: 'Check Compliance PMs (Running...)',
            systemFunction: 'WSJOBS',
            userFunction: 'CTJOBS',
            launchTarget: 'WSJOBS?USER_FUNCTION_NAME=CTJOBS',
            titleHints: ['compliance', 'work orders - compliance'],
            gridStoreIncludes: ['ctjobs', 'compliance'],
            gridMarkers: ['work orders - compliance', 'original pm due date'],
            modalTitle: 'Compliance Results'
        },
        fwos: {
            key: 'fwos',
            buttonId: 'myapm-fwo-run-btn',
            buttonLabel: 'Check WOs',
            runningLabel: 'Check WOs (Running...)',
            systemFunction: 'WSJOBS',
            userFunction: 'WSJOBS',
            launchTarget: 'WSJOBS?USER_FUNCTION_NAME=WSJOBS',
            titleHints: ['work order', 'work orders'],
            gridStoreIncludes: ['wsjobs'],
            gridMarkers: ['work order', 'sched. start date'],
            modalTitle: 'WO Results',
            dataspyLabel: 'Open Work Orders',
            dueFilterLabels: ['Sched. End Date', 'Scheduled End Date', 'End Date']
        },
        pms: {
            key: 'pms',
            buttonId: 'myapm-pm-run-btn',
            buttonLabel: 'Check PMs',
            runningLabel: 'Check PMs (Running...)',
            systemFunction: 'WSJOBS',
            userFunction: 'WSJOBS',
            launchTarget: 'WSJOBS?USER_FUNCTION_NAME=WSJOBS',
            titleHints: ['work order', 'work orders'],
            gridStoreIncludes: ['wsjobs'],
            gridMarkers: ['work order', 'sched. start date'],
            modalTitle: 'PM Results',
            dueFilterLabels: ['Original PM Due Date', 'Original PM Due', 'Original Due Date', 'Due Date']
        }
    };

    const MY_SHIFT = {
        buttonId: 'myapm-my-shift-btn',
        buttonLabel: 'My Shift',
        runningLabel: 'My Shift (Running...)'
    };

    const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

    function cleanText(raw) {
        return String(raw || '')
            .replace(/<[^>]*>/g, ' ')
            .replace(/&nbsp;/gi, ' ')
            .replace(/\s+/g, ' ')
            .trim();
    }

    function escapeHtml(raw) {
        return String(raw ?? '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    function textIncludes(raw, target) {
        return cleanText(raw).toLowerCase().includes(cleanText(target).toLowerCase());
    }

    function log(message, data) {
        if (!NAV_DEBUG) return;
        if (typeof data === 'undefined') {
            console.log(`${TRACE} ${message}`);
        } else {
            console.log(`${TRACE} ${message}`, data);
        }
    }

    function showToast(message, kind = 'info') {
        try {
            const doc = window.top.document;
            let toast = doc.getElementById('myapm-inline-toast');
            if (!toast) {
                toast = doc.createElement('div');
                toast.id = 'myapm-inline-toast';
                Object.assign(toast.style, {
                    position: 'fixed',
                    bottom: '24px',
                    left: '50%',
                    transform: 'translateX(-50%)',
                    padding: '10px 14px',
                    borderRadius: '8px',
                    fontFamily: 'Arial, sans-serif',
                    fontSize: '13px',
                    fontWeight: '700',
                    zIndex: CONTROL_Z_INDEX,
                    boxShadow: '0 6px 24px rgba(0,0,0,0.35)',
                    transition: 'opacity 0.2s ease',
                    opacity: '0',
                    pointerEvents: 'none'
                });
                doc.body.appendChild(toast);
            }
            const palette = kind === 'error'
                ? { background: '#b3261e', color: '#fff' }
                : kind === 'success'
                    ? { background: '#1f7a1f', color: '#fff' }
                    : { background: '#223247', color: '#fff' };
            toast.textContent = message;
            toast.style.background = palette.background;
            toast.style.color = palette.color;
            toast.style.opacity = '1';
            clearTimeout(showToast._timer);
            showToast._timer = setTimeout(() => {
                toast.style.opacity = '0';
            }, 1800);
        } catch (_) {}
    }

    const updateListeners = [];

    function normalizeVersionParts(version) {
        return String(version || '')
            .split(/[^0-9]+/)
            .map((part) => Number(part))
            .filter((part) => Number.isFinite(part));
    }

    function isNewerVersion(currentVersion, remoteVersion) {
        const currentParts = normalizeVersionParts(currentVersion);
        const remoteParts = normalizeVersionParts(remoteVersion);
        const length = Math.max(currentParts.length, remoteParts.length);
        for (let i = 0; i < length; i += 1) {
            const current = currentParts[i] || 0;
            const remote = remoteParts[i] || 0;
            if (remote > current) return true;
            if (remote < current) return false;
        }
        return false;
    }

    function subscribeToUpdates(callback) {
        if (typeof callback !== 'function') return;
        updateListeners.push(callback);
        if (window.__myapmUpdateAvailable) callback(window.__myapmRemoteVersion || '');
    }

    function notifyUpdateAvailable(remoteVersion) {
        window.__myapmUpdateAvailable = true;
        window.__myapmRemoteVersion = remoteVersion || '';
        updateListeners.forEach((callback) => {
            try {
                callback(window.__myapmRemoteVersion);
            } catch (_) {}
        });
    }

    function buildUpdateCheckUrl() {
        const separator = UPDATE_URL.includes('?') ? '&' : '?';
        return `${UPDATE_URL}${separator}myapm_update_check=${Date.now()}`;
    }

    function buildInstallScriptUrl() {
        const separator = SCRIPT_PAGE_URL.includes('?') ? '&' : '?';
        return `${SCRIPT_PAGE_URL}${separator}myapm_install_refresh=${Date.now()}`;
    }

    function scheduleScriptUpdateRetry(nextAttempt) {
        if (nextAttempt > UPDATE_CHECK_MAX_ATTEMPTS) return;
        clearTimeout(window.__myapmUpdateRetryTimer);
        window.__myapmUpdateRetryTimer = window.setTimeout(() => {
            checkForScriptUpdates(nextAttempt);
        }, UPDATE_CHECK_RETRY_MS);
    }

    function checkForScriptUpdates(attempt = 1) {
        if (window.__myapmUpdateAvailable || window.__myapmUpdateCheckInFlight) return;
        if (attempt === 1 && window.__myapmUpdateChecked) return;
        window.__myapmUpdateChecked = true;
        window.__myapmUpdateCheckInFlight = true;
        fetch(buildUpdateCheckUrl(), { cache: 'no-store' })
            .then((response) => {
                if (!response.ok) throw new Error(`Update check failed with ${response.status}`);
                return response.text();
            })
            .then((text) => {
                const match = String(text || '').match(/\/\/\s*@version\s+([^\s]+)/i);
                const remoteVersion = match && match[1] ? String(match[1]).trim() : '';
                if (!remoteVersion) throw new Error('Remote version missing');
                if (isNewerVersion(CURRENT_VERSION, remoteVersion)) {
                    notifyUpdateAvailable(remoteVersion);
                }
                clearTimeout(window.__myapmUpdateRetryTimer);
            })
            .catch(() => {
                scheduleScriptUpdateRetry(attempt + 1);
            })
            .finally(() => {
                window.__myapmUpdateCheckInFlight = false;
            });
    }

    function createUpdateBanner() {
        const wrap = document.createElement('div');
        wrap.id = 'myapm-settings-update-container';
        Object.assign(wrap.style, {
            flex: '0 0 auto',
            display: 'block'
        });

        const link = document.createElement('a');
        link.id = 'myapm-settings-update-link';
        link.href = buildInstallScriptUrl();
        link.target = '_blank';
        link.rel = 'noopener noreferrer';
        link.textContent = 'Check for Updates';
        Object.assign(link.style, {
            display: 'inline-flex',
            alignItems: 'center',
            gap: '6px',
            padding: '6px 10px',
            borderRadius: '999px',
            background: '#f39c12',
            color: '#fff',
            fontSize: '12px',
            fontWeight: '700',
            textDecoration: 'none',
            boxShadow: '0 4px 14px rgba(0,0,0,0.22)'
        });
        link.addEventListener('click', () => {
            link.href = buildInstallScriptUrl();
        });
        wrap.appendChild(link);

        subscribeToUpdates((remoteVersion) => {
            link.textContent = remoteVersion ? `Update Available: ${remoteVersion}` : 'Update Available';
        });

        return wrap;
    }

    function isCmpVisible(cmp) {
        try {
            if (!cmp || cmp.destroyed || cmp.isDestroyed || cmp.hidden) return false;
            if (typeof cmp.isVisible === 'function') return !!cmp.isVisible(true);
            const el = typeof cmp.getEl === 'function' ? cmp.getEl() : cmp.el;
            if (el && el.dom) {
                const style = el.dom.ownerDocument.defaultView.getComputedStyle(el.dom);
                return style.display !== 'none' && style.visibility !== 'hidden';
            }
            return true;
        } catch (_) {
            return false;
        }
    }

    function getTopEAM() {
        return PAGE_WINDOW.top && PAGE_WINDOW.top.EAM ? PAGE_WINDOW.top.EAM : null;
    }

    function getTopExt() {
        return PAGE_WINDOW.top && PAGE_WINDOW.top.Ext ? PAGE_WINDOW.top.Ext : null;
    }

    function getTopOrigin() {
        return (PAGE_WINDOW.top.location.origin && PAGE_WINDOW.top.location.origin.includes('hxgnsmartcloud.com'))
            ? PAGE_WINDOW.top.location.origin
            : 'https://us1.eam.hxgnsmartcloud.com';
    }

    function getThemeParam() {
        try {
            return PAGE_WINDOW.top.localStorage.getItem('apmTheme') || 'theme-dark';
        } catch (_) {
            return 'theme-dark';
        }
    }

    function buildWorkOrderUrl(workOrderNum, userFunction = 'WSJOBS') {
        const value = String(workOrderNum || '').trim();
        if (!value) return '';
        return `${getTopOrigin()}/web/base/logindisp?tenant=AMAZONRMENA_PRD&FROMEMAIL=YES&uitheme=${encodeURIComponent(getThemeParam())}&SYSTEM_FUNCTION_NAME=WSJOBS&USER_FUNCTION_NAME=${encodeURIComponent(userFunction)}&workordernum=${encodeURIComponent(value)}`;
    }

    function buildAuditUrl(auditNum, workOrderNum = '') {
        const auditValue = String(auditNum || '').trim();
        if (!auditValue) return '';
        const woValue = String(workOrderNum || auditValue).trim();
        return `${getTopOrigin()}/web/base/logindisp?tenant=AMAZONRMENA_PRD&FROMEMAIL=YES&uitheme=${encodeURIComponent(getThemeParam())}&SYSTEM_FUNCTION_NAME=WSJOBS&USER_FUNCTION_NAME=ADJOBS&workordernum=${encodeURIComponent(woValue)}&apmAudit=${encodeURIComponent(auditValue)}#apmAudit=${encodeURIComponent(auditValue)}`;
    }


    function normalizeUserText(value) {
        return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
    }

    function extractUserNameOnly(value) {
        const raw = String(value || '').trim();
        if (!raw) return '';
        const emailMatch = raw.match(/([A-Z0-9._%+-]+)@amazon\.com/i);
        if (emailMatch && emailMatch[1]) return emailMatch[1].trim();
        const genericEmailMatch = raw.match(/([^\s@]+)@[^\s@]+/);
        if (genericEmailMatch && genericEmailMatch[1]) return genericEmailMatch[1].trim();
        return raw;
    }

    function getSavedCurrentUserName() {
        try {
            return extractUserNameOnly(window.top.localStorage.getItem(CURRENT_USER_STORAGE_KEY) || '');
        } catch (_) {
            return '';
        }
    }

    function saveCurrentUserName(value) {
        const next = extractUserNameOnly(value);
        if (!next) return '';
        try {
            window.top.localStorage.setItem(CURRENT_USER_STORAGE_KEY, next);
        } catch (_) {}
        return next;
    }

    function detectCurrentUserName() {
        const topWin = window.top;
        const doc = topWin.document;

        try {
            const getter = topWin.EAM && topWin.EAM.UserData && typeof topWin.EAM.UserData.get === 'function'
                ? topWin.EAM.UserData.get.bind(topWin.EAM.UserData)
                : null;
            if (getter) {
                const keys = ['username', 'userName', 'login', 'loginName', 'email', 'person', 'employee'];
                for (const key of keys) {
                    const raw = getter(key);
                    const value = String(raw || '').trim();
                    if (value && value.length >= 2 && value.length <= 120) return saveCurrentUserName(value);
                }
            }
        } catch (_) {}

        const selectors = [
            '[data-testid*="user" i]',
            '[aria-label*="user" i]',
            '[title*="user" i]',
            '[id*="user" i]',
            '[class*="user" i]'
        ];
        for (const selector of selectors) {
            const el = doc.querySelector(selector);
            const text = String(el && el.textContent ? el.textContent : '').trim();
            if (text && text.length >= 2 && text.length <= 120) return saveCurrentUserName(text);
        }

        try {
            const headerText = Array.from(doc.querySelectorAll('header, .x-toolbar, .x-tab-bar, .x-box-inner')).slice(0, 12)
                .map((el) => String(el.textContent || '').trim())
                .filter(Boolean)
                .join(' ');
            const emailMatch = headerText.match(/\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i);
            if (emailMatch && emailMatch[0]) return saveCurrentUserName(emailMatch[0]);
        } catch (_) {}

        return getSavedCurrentUserName();
    }

    function getActiveModuleHeaderInner(ctx) {
        const scopedCtx = ctx || getActiveAPMContext();
        try {
            const screen = scopedCtx && scopedCtx.screen;
            if (screen && typeof screen.getModuleHeader === 'function') {
                const headerCmp = screen.getModuleHeader();
                const el = headerCmp && typeof headerCmp.getEl === 'function' ? headerCmp.getEl() : headerCmp && headerCmp.el;
                if (el && el.dom) {
                    const inner = el.dom.querySelector('div.module-header-inner') || el.dom;
                    if (inner) return inner;
                }
            }
        } catch (_) {}
        try {
            const doc = scopedCtx && scopedCtx.appWin && scopedCtx.appWin.document ? scopedCtx.appWin.document : document;
            return doc.querySelector('div.module-header-inner');
        } catch (_) {
            return null;
        }
    }

    function ensureActiveRecordHeaderUi() {
        const ctx = getActiveAPMContext();
        if (!ctx) return;
        const headerInner = getActiveModuleHeaderInner(ctx);
        if (!headerInner) return;
        const code = headerInner.querySelector('span.recordcode');
        const desc = headerInner.querySelector('span.recorddesc');
        if (!code || !desc || !code.parentNode || code.parentNode !== desc.parentNode) return;
        if (desc.querySelector && desc.querySelector('.apm-wo-inline-wrap, a.better-apm-workorder, button.copy-btn')) {
            desc.textContent = cleanText(desc.textContent || '');
            delete desc.dataset.workorderLinked;
            delete desc.dataset.workorderLinkedKey;
        }

        const parent = code.parentNode;
        const codeText = String(code.textContent || '').trim();
        const prefix = parent.querySelector('span.myapm-record-prefix');
        const separator = parent.querySelector('span.myapm-record-separator');
        if (!prefix || prefix.nextSibling !== code) {
            if (prefix) prefix.remove();
            const node = headerInner.ownerDocument.createElement('span');
            node.className = 'myapm-record-prefix';
            node.textContent = ' - ';
            node.setAttribute('aria-hidden', 'true');
            parent.insertBefore(node, code);
        }
        if (!separator || separator.previousSibling !== code || separator.nextSibling !== desc) {
            if (separator) separator.remove();
            const node = headerInner.ownerDocument.createElement('span');
            node.className = 'myapm-record-separator';
            node.textContent = ' - ';
            node.setAttribute('aria-hidden', 'true');
            parent.insertBefore(node, desc);
        }
        const ptpSeparator = parent.querySelector('span.myapm-record-ptp-separator');
        const badge = parent.querySelector('a.myapm-ptp-status-badge');
        if (badge) {
            let insertBeforeNode = badge.nextSibling;
            while (insertBeforeNode && (
                insertBeforeNode === ptpSeparator ||
                (insertBeforeNode.nodeType === 3 && !String(insertBeforeNode.textContent || '').trim())
            )) {
                insertBeforeNode = insertBeforeNode.nextSibling;
            }
            if (!insertBeforeNode) insertBeforeNode = prefix || code;
            if (!ptpSeparator || ptpSeparator.previousSibling !== badge || ptpSeparator.nextSibling !== insertBeforeNode) {
                if (ptpSeparator) ptpSeparator.remove();
                const node = headerInner.ownerDocument.createElement('span');
                node.className = 'myapm-record-ptp-separator';
                node.textContent = ' - ';
                node.setAttribute('aria-hidden', 'true');
                parent.insertBefore(node, insertBeforeNode);
            }
        } else if (ptpSeparator) {
            ptpSeparator.remove();
        }

        if (!codeText) return;
        const auditMode = flowMatchesContext(FLOWS.audits, ctx);
        const copyUrl = auditMode
            ? buildAuditUrl(codeText, codeText)
            : buildWorkOrderUrl(codeText, (ctx.screen && ctx.screen.getUserFunction && ctx.screen.getUserFunction()) || 'WSJOBS');
        let codeLink = code.querySelector('a.myapm-header-wo-link');
        if (!codeLink) {
            code.textContent = '';
            codeLink = headerInner.ownerDocument.createElement('a');
            codeLink.className = 'myapm-header-wo-link';
            codeLink.target = '_blank';
            codeLink.rel = 'noopener noreferrer';
            Object.assign(codeLink.style, {
                color: 'inherit',
                textDecoration: 'underline',
                textUnderlineOffset: '2px',
                cursor: 'pointer'
            });
            code.appendChild(codeLink);
        }
        codeLink.href = copyUrl;
        codeLink.textContent = codeText;
        codeLink.title = auditMode ? `Open Audit ${codeText} in a New Tab` : `Open Work Order ${codeText} in a New Tab`;

        let btn = parent.querySelector('button.myapm-header-copy-btn');
        if (!btn) {
            btn = headerInner.ownerDocument.createElement('button');
            btn.className = 'myapm-header-copy-btn';
            btn.type = 'button';
            btn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 384 512" width="15" height="15" aria-hidden="true" focusable="false"><path fill="currentColor" d="M192 0c-41.8 0-77.4 26.7-90.5 64L64 64C28.7 64 0 92.7 0 128V448c0 35.3 28.7 64 64 64h256c35.3 0 64-28.7 64-64V128c0-35.3-28.7-64-64-64h-37.5C269.4 26.7 233.8 0 192 0zm0 64a32 32 0 1 1 0 64 32 32 0 1 1 0-64zM305 273L177 401c-9.4 9.4-24.6 9.4-33.9 0L79 337c-9.4-9.4-9.4-24.6 0-33.9s24.6-9.4 33.9 0l47 47 111-111c9.4-9.4 24.6-9.4 33.9 0s9.4 24.6 0 33.9z"></path></svg>';
            Object.assign(btn.style, {
                marginLeft: '6px',
                display: 'inline-flex',
                alignItems: 'center',
                justifyContent: 'center',
                verticalAlign: 'middle',
                background: 'rgba(228, 121, 17, 0.26)',
                border: '1px solid rgba(255, 179, 71, 0.7)',
                borderRadius: '4px',
                cursor: 'pointer',
                width: '18px',
                height: '18px',
                padding: '1px',
                color: '#ffb347',
                opacity: '1',
                boxShadow: '0 0 0 1px rgba(0,0,0,0.12)',
                transition: 'background 0.15s, border-color 0.15s, color 0.15s, box-shadow 0.15s'
            });
            btn.addEventListener('mouseenter', () => {
                btn.style.background = 'rgba(228, 121, 17, 0.45)';
                btn.style.borderColor = 'rgba(255, 179, 71, 0.95)';
                btn.style.color = '#ffd18a';
                btn.style.boxShadow = '0 0 0 1px rgba(255,179,71,0.35)';
            });
            btn.addEventListener('mouseleave', () => {
                btn.style.background = 'rgba(228, 121, 17, 0.26)';
                btn.style.borderColor = 'rgba(255, 179, 71, 0.7)';
                btn.style.color = '#ffb347';
                btn.style.boxShadow = '0 0 0 1px rgba(0,0,0,0.12)';
            });
            btn.addEventListener('click', async (event) => {
                event.preventDefault();
                event.stopPropagation();
                try {
                    await navigator.clipboard.writeText(btn.dataset.copyUrl || '');
                    showToast('Link copied.', 'success');
                } catch (_) {
                    showToast('Failed to copy link.', 'error');
                }
            });
        }
        btn.dataset.copyUrl = copyUrl;
        btn.title = auditMode ? 'Copy Audit Link' : 'Copy Work Order Link';
        const desiredPrev = desc || code;
        if (btn.previousSibling !== desiredPrev) {
            parent.insertBefore(btn, desiredPrev.nextSibling);
        }
    }


    function ptpStatusTrackingEnabled() {
        return localStorage.getItem('myapmPtpStatusTracking') !== 'false';
    }

    function getPtpHistorySnapshot() {
        try {
            if (typeof window.__myApmGetPtpHistory === 'function') {
                return window.__myApmGetPtpHistory() || {};
            }
        } catch (_) {}
        try {
            return JSON.parse(localStorage.getItem(PTP_HISTORY_STORAGE_KEY) || '{}') || {};
        } catch (_) {
            return {};
        }
    }

    function getPtpSiteCandidate() {
        const candidates = [];

        try {
            const topWin = window.top;
            const getter = topWin && topWin.EAM && topWin.EAM.UserData && typeof topWin.EAM.UserData.get === 'function'
                ? topWin.EAM.UserData.get.bind(topWin.EAM.UserData)
                : null;
            if (getter) {
                ['organization', 'org', 'site', 'sitecode', 'siteCode', 'location'].forEach((key) => {
                    const value = cleanText(getter(key));
                    if (value) candidates.push(value);
                });
            }
        } catch (_) {}

        try {
            const topDoc = window.top && window.top.document ? window.top.document : document;
            const headerText = Array.from(topDoc.querySelectorAll('header, .x-toolbar, .x-box-inner, .x-panel-header, .x-tab-bar')).slice(0, 20)
                .map((el) => cleanText(el.textContent || ''))
                .filter(Boolean)
                .join(' ');
            const headerOrgMatch = headerText.match(/\bOrganization\s*\(([^)]+)\)/i);
            if (headerOrgMatch && headerOrgMatch[1]) candidates.push(cleanText(headerOrgMatch[1]));
        } catch (_) {}

        try {
            const search = new URLSearchParams(String(window.top.location.search || ''));
            ['organization', 'org', 'site', 'sitecode'].forEach((key) => {
                const value = cleanText(search.get(key));
                if (value) candidates.push(value);
            });
        } catch (_) {}

        const selectors = [
            'input[name*=organization]', 'input[name*=site]', 'input[name*=dept]',
            '[name*=organization]', '[name*=site]', '[data-field*=organization]', '[data-field*=site]'
        ];
        selectors.forEach((selector) => {
            document.querySelectorAll(selector).forEach((node) => {
                const value = cleanText(node.value || node.textContent || node.getAttribute('value') || '');
                if (value) candidates.push(value);
            });
        });

        const metaText = cleanText((window.top && window.top.document && window.top.document.body
            ? window.top.document.body.innerText
            : (document.body ? document.body.innerText : '')).slice(0, 6000));
        const bodyOrgMatch = metaText.match(/\bOrganization\s*\(([^)]+)\)/i);
        if (bodyOrgMatch && bodyOrgMatch[1]) candidates.push(cleanText(bodyOrgMatch[1]));
        const siteTokenMatch = metaText.match(/\b([A-Z]{2,6}\d{1,4}[A-Z]?)\b/);
        if (siteTokenMatch && siteTokenMatch[1]) candidates.push(siteTokenMatch[1]);

        return candidates.find((value) => value && value.length <= 32) || '';
    }

    function getPtpDescriptionForHost(host, fallback = '') {
        const fallbackText = cleanText(fallback);
        if (!host) return fallbackText;
        const explicit = cleanText(host.getAttribute('data-wo-desc') || '');
        if (explicit) return explicit;
        const row = host.closest ? host.closest('.x-grid-item') : null;
        if (row) {
            const cells = Array.from(row.querySelectorAll('.x-grid-cell'));
            const hostCell = host.closest ? host.closest('.x-grid-cell') : null;
            const hostIndex = hostCell ? cells.indexOf(hostCell) : -1;
            const preferredCell = hostIndex > -1 ? (cells[hostIndex + 1] || null) : null;
            const preferredText = cleanText(preferredCell && preferredCell.innerText || '');
            if (preferredText && !/^\d{11}$/.test(preferredText) && preferredText.toUpperCase() !== 'PTP') return preferredText;

            const cellTexts = cells.map((node) => cleanText(node.innerText || node.textContent || ''));
            const description = cellTexts.find((value, index) => {
                if (!value || index === hostIndex) return false;
                if (/^\d{1,4}$/.test(value)) return false;
                if (/^\d{11}$/.test(value)) return false;
                if (value.toUpperCase() === 'PTP') return false;
                return true;
            });
            if (description) return description;
        }
        return fallbackText;
    }

    function buildPtpUrlForWorkOrder(woNumber, description = '') {
        const workOrder = cleanText(woNumber);
        if (!workOrder) return '';
        const site = cleanText(getPtpSiteCandidate());
        const simTitle = cleanText(description);
        const params = new URLSearchParams();
        params.set('workordernum', workOrder);
        if (site || simTitle) {
            params.set('organization', simTitle ? `${site || ''}, ${simTitle}` : site);
        }
        return `${PTP_LINK_BASE}?${params.toString()}`;
    }

    function isPtpCompletionStillValid(record) {
        if (!record) return false;
        const normalizedStatus = String(record.status || '').toUpperCase();
        if (normalizedStatus !== 'COMPLETE') return false;
        const completedAt = Number(record.time || 0);
        if (!Number.isFinite(completedAt) || completedAt <= 0) return false;
        return (Date.now() - completedAt) <= PTP_COMPLETE_VALID_MS;
    }

    function upsertPtpBadge(host, woNumber, record, options = {}) {
        if (!host) return;
        let badge = host.querySelector('.myapm-ptp-status-badge');
        const showPending = !!options.showPending;
        if (!ptpStatusTrackingEnabled() || !woNumber || (!record && !showPending)) {
            if (badge) badge.remove();
            return;
        }
        const normalizedStatus = String(record && record.status || '').toUpperCase();
        const isComplete = isPtpCompletionStillValid(record);
        const isPending = !record || (normalizedStatus === 'COMPLETE' && !isComplete);
        const icon = isPending ? '⌛' : (isComplete ? '✅' : '❌');
        const statusText = isPending
            ? (record && normalizedStatus === 'COMPLETE' ? 'Completed PTP expired' : 'Not completed yet')
            : (normalizedStatus === 'INCOMPLETE' ? 'Incomplete' : (isComplete ? 'Completed' : 'Cancelled/Incomplete'));
        const dateText = record && record.time ? new Date(record.time).toLocaleString() : '';
        const description = getPtpDescriptionForHost(host, options.description || '');
        const ptpUrl = buildPtpUrlForWorkOrder(woNumber, description);
        if (!badge) {
            badge = host.ownerDocument.createElement('a');
            badge.className = 'myapm-ptp-status-badge';
            badge.target = '_blank';
            badge.rel = '';
            Object.assign(badge.style, {
                display: 'inline-flex',
                alignItems: 'center',
                gap: '4px',
                marginRight: '6px',
                padding: '1px 6px',
                borderRadius: '999px',
                fontSize: '11px',
                fontWeight: '700',
                lineHeight: '16px',
                whiteSpace: 'nowrap',
                verticalAlign: 'middle',
                cursor: 'pointer',
                flexShrink: '0',
                textDecoration: 'none'
            });
            badge.addEventListener('click', (event) => {
                event.stopPropagation();
            });
        }
        Object.assign(badge.style, isPending ? {
            border: '1px solid rgba(255, 179, 71, 0.7)',
            background: 'rgba(64, 38, 8, 0.92)',
            color: '#ffd18a'
        } : (isComplete ? {
            border: '1px solid rgba(76, 175, 80, 0.6)',
            background: 'rgba(15, 48, 23, 0.9)',
            color: '#d7f7dd'
        } : {
            border: '1px solid rgba(220, 92, 92, 0.62)',
            background: 'rgba(66, 20, 20, 0.9)',
            color: '#ffd6d6'
        }));
        badge.dataset.woNum = String(woNumber);
        badge.href = ptpUrl || '#';
        badge.title = ((isPending && !(record && normalizedStatus === 'COMPLETE'))
            ? 'PTP not completed yet'
            : `${statusText} on ${dateText || 'unknown date'}`) + (ptpUrl ? ' — Click to open PTP page' : '');
        badge.textContent = `${icon} PTP`;
        const desiredFirst = host.firstChild;
        if (desiredFirst !== badge) {
            host.insertBefore(badge, desiredFirst || null);
        }
    }

    function refreshPtpGridBadges() {
        const ctx = getActiveAPMContext();
        const doc = ctx && ctx.appWin && ctx.appWin.document ? ctx.appWin.document : document;
        const history = getPtpHistorySnapshot();
        doc.querySelectorAll('.apm-wo-inline-wrap').forEach((wrap) => {
            const woNumber = String(wrap.getAttribute('data-wo-num') || '').trim();
            const description = String(wrap.getAttribute('data-wo-desc') || '').trim();
            upsertPtpBadge(wrap, woNumber, history[woNumber], { showPending: true, description });
        });
    }

    function refreshPtpHeaderBadge() {
        const ctx = getActiveAPMContext();
        const headerInner = getActiveModuleHeaderInner(ctx);
        if (!headerInner) return;
        const code = headerInner.querySelector('span.recordcode');
        const desc = headerInner.querySelector('span.recorddesc');
        const parent = code && desc && code.parentNode === desc.parentNode ? code.parentNode : null;
        if (!parent || !code) return;
        const woNumber = String(code.textContent || '').match(/\b\d{11}\b/);
        const history = getPtpHistorySnapshot();
        upsertPtpBadge(parent, woNumber ? woNumber[0] : '', woNumber ? history[woNumber[0]] : null, { showPending: true, description: cleanText(desc && desc.textContent || '') });
    }

    function refreshPtpDecorations() {
        refreshPtpGridBadges();
        refreshPtpHeaderBadge();
    }

    function getVisibleGridComponents() {
        const extCandidates = [];
        const topExt = window.top && window.top.Ext ? window.top.Ext : null;
        if (topExt) extCandidates.push(topExt);
        try {
            const ctx = getActiveAPMContext();
            if (ctx && ctx.Ext && !extCandidates.includes(ctx.Ext)) extCandidates.push(ctx.Ext);
            if (ctx && ctx.appWin && ctx.appWin.Ext && !extCandidates.includes(ctx.appWin.Ext)) extCandidates.push(ctx.appWin.Ext);
        } catch (_) {}
        const grids = [];
        const seen = new Set();
        extCandidates.forEach((extRef) => {
            if (!extRef || !extRef.ComponentQuery || typeof extRef.ComponentQuery.query !== 'function') return;
            try {
                extRef.ComponentQuery.query('gridpanel').forEach((grid) => {
                    if (!grid || grid.hidden || grid.destroyed) return;
                    const key = grid.id || grid.itemId || String(grid);
                    if (seen.has(key)) return;
                    const el = typeof grid.getEl === 'function' ? grid.getEl() : null;
                    const visible = !el || typeof el.isVisible !== 'function' ? true : (() => {
                        try { return el.isVisible(true); } catch (_) { return true; }
                    })();
                    if (!visible) return;
                    seen.add(key);
                    grids.push(grid);
                });
            } catch (_) {}
        });
        return grids;
    }

    function isWorkOrderColumn(column) {
        if (!column) return false;
        const bits = [column.text, column.header, column.menuText, column.dataIndex, column.itemId, column.name]
            .map((v) => String(v || '').toLowerCase().trim())
            .filter(Boolean);
        return bits.some((v) => v === 'work order' || v === 'workorder' || v === 'wonum' || v.includes('work order'));
    }

    function isRmeAuditColumn(column) {
        if (!column) return false;
        const bits = [column.text, column.header, column.menuText, column.dataIndex, column.itemId, column.name]
            .map((v) => String(v || '').toLowerCase().trim())
            .filter(Boolean);
        if (!bits.length) return false;
        return bits.some((v) => v === 'rme audits' || v === 'rme audit' || v === 'audit' || v === 'audits' || v.includes('rme audit'))
            && !bits.some((v) => v.includes('audit date'));
    }


    function requestGridResize(reason) {
        try {
            const payload = { reason: String(reason || 'generic'), ts: Date.now() };
            localStorage.setItem(GRID_RESIZE_REQUEST_KEY, JSON.stringify(payload));
        } catch (_) {}
    }

    function peekGridResizeRequest(maxAgeMs) {
        try {
            const raw = localStorage.getItem(GRID_RESIZE_REQUEST_KEY);
            if (!raw) return null;
            const parsed = JSON.parse(raw);
            if (!parsed || !parsed.ts) {
                try { localStorage.removeItem(GRID_RESIZE_REQUEST_KEY); } catch (_) {}
                return null;
            }
            const age = Math.abs(Date.now() - Number(parsed.ts || 0));
            if (age > (Number(maxAgeMs) || 10000)) {
                try { localStorage.removeItem(GRID_RESIZE_REQUEST_KEY); } catch (_) {}
                return null;
            }
            return parsed;
        } catch (_) {
            try { localStorage.removeItem(GRID_RESIZE_REQUEST_KEY); } catch (_) {}
            return null;
        }
    }

    function clearGridResizeRequest() {
        try { localStorage.removeItem(GRID_RESIZE_REQUEST_KEY); } catch (_) {}
    }

    function scheduleGridResizeRetries(reason, attempts, delayMs) {
        const count = Math.max(1, Number(attempts) || GRID_RESIZE_RETRY_COUNT);
        const delay = Math.max(100, Number(delayMs) || GRID_RESIZE_RETRY_MS);
        for (let i = 0; i < count; i += 1) {
            setTimeout(() => {
                try {
                    ensureReasonableWorkOrderColumnWidth();
                } catch (_) {}
            }, i * delay);
        }
    }

    function ensureReasonableWorkOrderColumnWidth() {
        const COLUMN_RULES = [
            { match: isWorkOrderColumn, minWidth: 190 },
            { match: isRmeAuditColumn, minWidth: 190 }
        ];
        let foundTargetColumn = false;
        getVisibleGridComponents().forEach((grid) => {
            const columns = getGridColumns(grid);
            let gridTouched = false;
            columns.forEach((column) => {
                const rule = COLUMN_RULES.find((entry) => entry.match(column));
                if (!rule) return;
                foundTargetColumn = true;
                try {
                    const currentWidth = Number(typeof column.getWidth === 'function' ? column.getWidth() : column.width || 0) || 0;
                    if ('minWidth' in column) column.minWidth = Math.max(Number(column.minWidth) || 0, rule.minWidth);
                    if ('flex' in column) column.flex = null;
                    if (currentWidth >= rule.minWidth && Number(column.minWidth) >= rule.minWidth) return;
                    if (typeof column.setWidth === 'function') column.setWidth(rule.minWidth);
                    else column.width = rule.minWidth;
                    gridTouched = true;
                } catch (_) {}
            });
            if (!gridTouched) return;
            try {
                if (typeof grid.updateLayout === 'function') grid.updateLayout();
            } catch (_) {}
            try {
                if (grid.view && typeof grid.view.refresh === 'function') grid.view.refresh();
            } catch (_) {}
        });
        return foundTargetColumn;
    }


    function isAssignedToCurrentUser(assignedValue, currentUserRaw) {
        const assigned = normalizeUserText(assignedValue);
        const current = normalizeUserText(currentUserRaw);
        if (!assigned || !current) return false;
        if (assigned.includes(current) || current.includes(assigned)) return true;
        const localPart = current.split('@')[0];
        return !!(localPart && (assigned.includes(localPart) || localPart.includes(assigned)));
    }

    function getActiveAPMContext() {
        try {
            const topWin = window.top;
            const pageTopWin = PAGE_WINDOW.top || PAGE_WINDOW;
            const topEAM = getTopEAM();
            const topExt = (PAGE_WINDOW.top && PAGE_WINDOW.top.Ext) ? PAGE_WINDOW.top.Ext : null;
            if (!topWin || !topEAM || !topExt) return null;

            const main = topEAM.Utils && typeof topEAM.Utils.getMainContentPanel === 'function'
                ? topEAM.Utils.getMainContentPanel()
                : null;
            if (!main || typeof main.getActiveTab !== 'function') return null;

            const activeTab = main.getActiveTab();
            if (!activeTab) return null;

            const cached = !!activeTab.isCachedFrame;
            let appWin = pageTopWin;
            if (cached && typeof activeTab.getWin === 'function') {
                try {
                    appWin = activeTab.getWin() || pageTopWin;
                } catch (_) {
                    appWin = pageTopWin;
                }
            }

            const EAMCtx = (appWin && appWin.EAM) || topEAM;
            const ExtCtx = (appWin && appWin.Ext) || topExt;
            if (!EAMCtx || !ExtCtx) return null;

            let screen = null;
            if (cached && typeof activeTab.getScreen === 'function') {
                screen = activeTab.getScreen() || null;
            }
            if (!screen && EAMCtx.Utils && typeof EAMCtx.Utils.getScreen === 'function') {
                screen = EAMCtx.Utils.getScreen() || null;
            }

            const currentTab = screen && typeof screen.getCurrentTab === 'function'
                ? screen.getCurrentTab()
                : (typeof activeTab.getCurrentTab === 'function' ? activeTab.getCurrentTab() : null);

            return { topWin: pageTopWin, appWin, EAM: EAMCtx, Ext: ExtCtx, main, activeTab, cached, screen, currentTab };
        } catch (_) {
            return null;
        }
    }

    function getScreenIdentity(ctx) {
        const screen = ctx && ctx.screen;
        const currentTab = ctx && ctx.currentTab;
        let moduleHeaderText = '';

        try {
            if (screen && typeof screen.getModuleHeader === 'function') {
                const headerCmp = screen.getModuleHeader();
                const el = headerCmp && typeof headerCmp.getEl === 'function' ? headerCmp.getEl() : headerCmp && headerCmp.el;
                if (el && el.dom) moduleHeaderText = cleanText(el.dom.textContent);
            }
        } catch (_) {}

        if (!moduleHeaderText) {
            try {
                const doc = ctx && ctx.appWin && ctx.appWin.document;
                const headerEl = doc && doc.querySelector('.module_header, [class*="module_header"]');
                moduleHeaderText = cleanText(headerEl ? headerEl.textContent : '');
            } catch (_) {}
        }

        return {
            systemFunction: screen && typeof screen.getSystemFunction === 'function' ? cleanText(screen.getSystemFunction()) : '',
            userFunction: screen && typeof screen.getUserFunction === 'function' ? cleanText(screen.getUserFunction()) : '',
            currentTabName: currentTab && typeof currentTab.getTabName === 'function' ? cleanText(currentTab.getTabName()) : '',
            moduleHeaderText,
            activeTabTitle: (() => {
                try {
            const tabEl = document.querySelector('.x-tab-active .x-tab-inner');
            return cleanText(tabEl ? tabEl.textContent : '');
        } catch (_) {
            return '';
                }
            })()
        };
    }

    function flowMatchesContext(flow, ctx) {
        if (!flow || !ctx || !ctx.screen) return false;
        const id = getScreenIdentity(ctx);
        if (flow.systemFunction && id.systemFunction && id.systemFunction !== flow.systemFunction) return false;
        if (flow.userFunction && id.userFunction && id.userFunction !== flow.userFunction) return false;

        if (flow.titleHints && flow.titleHints.length) {
            const haystack = `${id.moduleHeaderText} ${id.activeTabTitle}`.toLowerCase();
            const anyHint = flow.titleHints.some((hint) => haystack.includes(String(hint).toLowerCase()));
            if (!id.userFunction && !anyHint) return false;
        }

        return true;
    }

    function isRunnableFlowContext(flow, ctx) {
        if (!flowMatchesContext(flow, ctx)) return false;
        try {
            const id = getScreenIdentity(ctx);
            const tabName = cleanText(id.currentTabName || '').toUpperCase();
            const runBtn = getRunButton(ctx);
            const grid = getActiveGrid(ctx, flow);
            if (runBtn && grid) return true;
            if (tabName === 'HDR' && !runBtn) return false;
            return !!(runBtn || grid);
        } catch (_) {
            return false;
        }
    }

    function isAppMasked(ctx) {
        const wins = [];
        if (ctx && ctx.topWin) wins.push(ctx.topWin);
        if (ctx && ctx.appWin && ctx.appWin !== ctx.topWin) wins.push(ctx.appWin);

        for (const win of wins) {
            try {
                const doc = win.document;
                if (!doc) continue;
                const selectors = [
                    '.x-mask[style*="display: block"]',
                    '.x-mask-msg[style*="display: block"]',
                    '.x-mask:not([style*="display: none"])',
                    '[id^="loadmask-"]'
                ];
                if (selectors.some((selector) => doc.querySelector(selector))) return true;
            } catch (_) {}
        }

        try {
            const app = ctx && ctx.EAM && typeof ctx.EAM.getApplication === 'function' ? ctx.EAM.getApplication() : null;
            if (app && app.isMasked === true) return true;
        } catch (_) {}

        return false;
    }

    async function waitFor(predicate, timeoutMs, label) {
        const start = Date.now();
        let lastError = null;

        while ((Date.now() - start) < timeoutMs) {
            try {
                const value = predicate();
                if (value) return value;
            } catch (error) {
                lastError = error;
            }
            await delay(POLL_MS);
        }

        const suffix = label ? `: ${label}` : '';
        const errorMessage = lastError ? `${suffix} (${lastError.message || String(lastError)})` : suffix;
        throw new Error(`Timed out waiting${errorMessage}`);
    }

    async function waitForContext(flow, timeoutMs) {
        return waitFor(() => {
            const ctx = getActiveAPMContext();
            return flowMatchesContext(flow, ctx) ? ctx : null;
        }, timeoutMs, `active context ${flow.systemFunction}/${flow.userFunction}`);
    }

    async function waitForUnmasked(flow, timeoutMs) {
        return waitFor(() => {
            const ctx = getActiveAPMContext();
            if (!flowMatchesContext(flow, ctx)) return null;
            return isAppMasked(ctx) ? null : ctx;
        }, timeoutMs, `screen ready ${flow.key}`);
    }

    function navigateToFlow(flow) {
        const topEAM = getTopEAM();
        if (!topEAM || !topEAM.Nav || typeof topEAM.Nav.launchScreen !== 'function') {
            throw new Error('EAM.Nav.launchScreen is unavailable');
        }

        log(`launching ${flow.key}`, { target: flow.launchTarget });
        topEAM.Nav.launchScreen(flow.launchTarget, null, {
            fromNav: true,
            smartCache: false,
            skipCacheCheck: true,
            disableAutoLoadMask: false
        });
    }

    function getOwnedGridCandidates(rootCmp) {
        if (!rootCmp || typeof rootCmp.query !== 'function') return [];
        const selectors = [
            'gridpanel[displayDataspy=true]',
            'gridpanel',
            'readonlygrid',
            'editablegrid'
        ];
        const seen = new Set();
        const grids = [];

        for (const selector of selectors) {
            let matches = [];
            try {
                matches = rootCmp.query(selector) || [];
            } catch (_) {
                matches = [];
            }
            for (const cmp of matches) {
                if (!cmp || seen.has(cmp.id) || !isCmpVisible(cmp)) continue;
                seen.add(cmp.id);
                grids.push(cmp);
            }
        }

        return grids;
    }

    function scoreGrid(flow, grid) {
        if (!grid) return -1;
        let score = 0;
        try {
            const store = typeof grid.getStore === 'function' ? grid.getStore() : null;
            const storeId = String((store && (store.storeId || (typeof store.getStoreId === 'function' ? store.getStoreId() : ''))) || '').toLowerCase();
            const xtype = String(grid.xtype || '').toLowerCase();
            const itemId = String(grid.itemId || '').toLowerCase();
            const gridText = (() => {
                try {
                    const el = typeof grid.getEl === 'function' ? grid.getEl() : grid.el;
                    return cleanText(el && el.dom ? el.dom.textContent : '').toLowerCase();
                } catch (_) {
                    return '';
                }
            })();

            if (grid.displayDataspy === true) score += 8;
            if (xtype.includes('grid')) score += 2;
            if (itemId.includes('grid')) score += 1;
            if ((flow.gridStoreIncludes || []).some((token) => storeId.includes(String(token).toLowerCase()))) score += 10;
            if ((flow.gridMarkers || []).some((marker) => gridText.includes(String(marker).toLowerCase()))) score += 6;

            if (store && typeof store.getCount === 'function' && store.getCount() >= 0) score += 1;
        } catch (_) {}
        return score;
    }

    function getActiveGrid(ctx, flow) {
        if (!ctx || !ctx.screen) return null;

        const candidates = [];
        const push = (cmp, source) => {
            if (!cmp || !isCmpVisible(cmp)) return;
            candidates.push({ cmp, source, score: scoreGrid(flow, cmp) });
        };

        try {
            if (typeof ctx.screen.getListView === 'function') {
                push(ctx.screen.getListView(), 'screen.getListView');
            }
        } catch (_) {}

        if (ctx.currentTab) {
            const owned = getOwnedGridCandidates(ctx.currentTab);
            owned.forEach((cmp) => push(cmp, 'currentTab.query'));
        }

        const screenOwned = getOwnedGridCandidates(ctx.screen);
        screenOwned.forEach((cmp) => push(cmp, 'screen.query'));

        candidates.sort((a, b) => b.score - a.score);
        const winner = candidates[0] || null;
        if (winner) {
            log(`grid selected for ${flow.key}`, {
                source: winner.source,
                score: winner.score,
                xtype: winner.cmp.xtype,
                itemId: winner.cmp.itemId || null
            });
            return winner.cmp;
        }

        return null;
    }

    function getRunButton(ctx) {
        const roots = [ctx && ctx.currentTab, ctx && ctx.screen].filter(Boolean);
        const selectors = [
            'button[action=run]',
            'button[uftId=run]',
            'dataspy button[action=run]',
            'button[itemId=run]',
            'button'
        ];

        for (const root of roots) {
            if (!root || typeof root.down !== 'function') continue;
            for (const selector of selectors) {
                try {
                    const cmp = selector === 'button' ? null : root.down(selector);
                    if (cmp && isCmpVisible(cmp)) return cmp;
                } catch (_) {}
            }

            try {
                const buttons = root.query ? root.query('button') : [];
                const byText = buttons.find((cmp) => {
                    if (!isCmpVisible(cmp)) return false;
                    const label = cleanText(cmp.text || cmp.ariaLabel || cmp.title || '');
                    return label.toLowerCase() === 'run';
                });
                if (byText) return byText;
            } catch (_) {}
        }

        return null;
    }


    function getContextRoots(ctx) {
        return [ctx && ctx.currentTab, ctx && ctx.screen].filter(Boolean);
    }

    function queryVisibleOwnedComponents(ctx, selector) {
        const roots = getContextRoots(ctx);
        const found = [];
        const seen = new Set();
        roots.forEach((root) => {
            if (!root || typeof root.query !== 'function') return;
            let matches = [];
            try {
                matches = root.query(selector) || [];
            } catch (_) {
                matches = [];
            }
            matches.forEach((cmp) => {
                if (!cmp || seen.has(cmp) || !isCmpVisible(cmp)) return;
                seen.add(cmp);
                found.push(cmp);
            });
        });
        return found;
    }

    async function ensureComponentStoreLoaded(store, timeoutMs = 5000) {
        if (!store) return;
        if (typeof store.isLoading === 'function' && store.isLoading()) {
            await waitFor(() => !(store.isLoading && store.isLoading()), timeoutMs, 'component store load');
            return;
        }
        const count = typeof store.getCount === 'function' ? store.getCount() : 0;
        if (count > 0) return;
        try {
            if (typeof store.load === 'function') store.load();
        } catch (_) {}
        await delay(120);
    }

    function getRecordStrings(record) {
        if (!record) return [];
        let data = {};
        try {
            data = typeof record.getData === 'function' ? record.getData() : (record.data || {});
        } catch (_) {
            data = record && record.data ? record.data : {};
        }
        const values = [];
        Object.keys(data || {}).forEach((key) => {
            const value = data[key];
            if (value !== null && typeof value !== 'undefined') values.push(cleanText(value));
        });
        return values.filter(Boolean);
    }

    function findComboRecordByLabel(combo, labels) {
        const wanted = (Array.isArray(labels) ? labels : [labels])
            .map((label) => cleanText(label).toLowerCase())
            .filter(Boolean);
        if (!combo || !wanted.length || typeof combo.getStore !== 'function') return null;
        const store = combo.getStore();
        if (!store || typeof store.each !== 'function') return null;
        let winner = null;
        store.each((record) => {
            if (winner) return;
            const haystack = getRecordStrings(record).join(' | ').toLowerCase();
            if (wanted.some((label) => haystack === label || haystack.includes(label))) winner = record;
        });
        return winner;
    }

    async function setComboByLabel(combo, labels) {
        if (!combo || typeof combo.getStore !== 'function') return false;
        await ensureComponentStoreLoaded(combo.getStore(), 5000);
        const record = findComboRecordByLabel(combo, labels);
        if (!record) return false;
        const valueField = combo.valueField || 'field1';
        let nextValue = null;
        try {
            nextValue = typeof record.get === 'function' ? (record.get(valueField) ?? record.get('field1')) : null;
        } catch (_) {
            nextValue = null;
        }
        if (nextValue === null || typeof nextValue === 'undefined') return false;
        try {
            combo.setValue(nextValue);
            if (typeof combo.fireEvent === 'function') {
                combo.fireEvent('select', combo, record);
                combo.fireEvent('change', combo, nextValue);
                combo.fireEvent('blur', combo);
            }
            return true;
        } catch (_) {
            return false;
        }
    }

    function setFieldValue(cmp, value) {
        if (!cmp) return false;
        try {
            if (typeof cmp.setValue === 'function') cmp.setValue(value);
            if (typeof cmp.setRawValue === 'function') cmp.setRawValue(value);
            if (typeof cmp.fireEvent === 'function') {
                cmp.fireEvent('change', cmp, value);
                cmp.fireEvent('blur', cmp);
            }
            const el = typeof cmp.getEl === 'function' ? cmp.getEl() : cmp.el;
            const input = el && el.dom ? el.dom.querySelector('input, textarea') : null;
            if (input) {
                input.value = value;
                input.dispatchEvent(new Event('input', { bubbles: true }));
                input.dispatchEvent(new Event('change', { bubbles: true }));
            }
            return true;
        } catch (_) {
            return false;
        }
    }

    function findDataspyCombo(ctx) {
        const selectors = ['combobox[name=dataspylist]', 'uxcombobox[name=dataspylist]'];
        for (const selector of selectors) {
            const cmp = queryVisibleOwnedComponents(ctx, selector)[0];
            if (cmp) return cmp;
        }
        return null;
    }

    function findPreferredField(ctx, selectors) {
        for (const selector of selectors) {
            const cmp = queryVisibleOwnedComponents(ctx, selector)[0];
            if (cmp) return cmp;
        }
        return null;
    }

    function clearWorkOrderField(ctx) {
        const field = findPreferredField(ctx, [
            'textfield[name=ff_workordernum]',
            'textfield[name*=workordernum]',
            'textfield[name*=wonum]'
        ]);
        return setFieldValue(field, '');
    }

    function findOwnedFieldBySelectors(ctx, selectors) {
        for (const selector of selectors) {
            const cmp = queryVisibleOwnedComponents(ctx, selector)[0];
            if (cmp) return cmp;
        }
        return null;
    }

    function getCmpDom(cmp) {
        try {
            const el = typeof cmp.getEl === 'function' ? cmp.getEl() : cmp && cmp.el;
            return el && el.dom ? el.dom : null;
        } catch (_) {
            return null;
        }
    }

    function getOperatorMenuItems(button) {
        const menu = button && (button.menu || (typeof button.getMenu === 'function' ? button.getMenu() : null));
        return menu && menu.items && Array.isArray(menu.items.items) ? menu.items.items : [];
    }

    function getOperatorItemText(item) {
        return cleanText(item && (item.text || item.itemId || item.tooltip || item.iconCls || ''))
            .toLowerCase();
    }

    function scoreOperatorItem(item, wanted) {
        const text = getOperatorItemText(item);
        if (!text) return Number.NEGATIVE_INFINITY;
        let score = Number.NEGATIVE_INFINITY;
        wanted.forEach((label) => {
            if (!label) return;
            if (text === label) score = Math.max(score, 100);
            else if (text.includes(label)) score = Math.max(score, 80);
            if (label === '<=' && /less than/.test(text) && /equal/.test(text)) score = Math.max(score, 120);
            if ((label === 'on or before' || label === 'before') && /before/.test(text)) score = Math.max(score, 90);
            if ((label === 'begins with' || label === 'starts with') && (/begin/.test(text) || /start/.test(text))) score = Math.max(score, 120);
        });
        return score;
    }

    function invokeMenuItem(item, button) {
        if (!item) return false;
        try {
            if (typeof item.handler === 'function') item.handler.call(item.scope || item, item);
            else if (typeof item.fireEvent === 'function') item.fireEvent('click', item);
            if (button && typeof button.setText === 'function' && item.text) button.setText(item.text);
            return true;
        } catch (_) {
            return false;
        }
    }

    function findOperatorButtonForField(ctx, fieldCmp) {
        if (!fieldCmp) return null;
        const fieldEl = getCmpDom(fieldCmp);
        const allButtons = queryVisibleOwnedComponents(ctx, 'button');
        const candidates = [];
        const push = (button, score) => {
            if (!button || candidates.some((entry) => entry.button === button)) return;
            const items = getOperatorMenuItems(button);
            if (!items.length) return;
            candidates.push({ button, score });
        };

        if (fieldEl) {
            const wrappers = [
                fieldEl.closest('.field-filter-column'),
                fieldEl.closest('.x-form-item'),
                fieldEl.parentElement,
                fieldEl.closest('.x-container')
            ].filter(Boolean);
            wrappers.forEach((wrapper, index) => {
                allButtons.forEach((button) => {
                    const btnEl = getCmpDom(button);
                    if (btnEl && wrapper.contains(btnEl)) push(button, 200 - index * 20);
                });
            });
        }

        const ownerCt = fieldCmp.ownerCt || (typeof fieldCmp.up === 'function' ? fieldCmp.up() : null);
        if (ownerCt && typeof ownerCt.query === 'function') {
            try {
                (ownerCt.query('button') || []).forEach((button) => push(button, 150));
            } catch (_) {}
        }

        allButtons.forEach((button) => {
            const itemId = cleanText(button.itemId || button.name || button.tooltip || '').toLowerCase();
            if (itemId.includes('operator') || itemId.includes('filter')) push(button, 20);
        });

        candidates.sort((a, b) => b.score - a.score);
        return candidates[0] ? candidates[0].button : null;
    }

    function setFieldOperatorByLabel(ctx, fieldCmp, preferredLabels) {
        const button = findOperatorButtonForField(ctx, fieldCmp);
        if (!button) return false;
        const wanted = (Array.isArray(preferredLabels) ? preferredLabels : [preferredLabels])
            .map((label) => cleanText(label).toLowerCase())
            .filter(Boolean);
        const items = getOperatorMenuItems(button);
        if (!items.length || !wanted.length) return false;
        let best = null;
        let bestScore = Number.NEGATIVE_INFINITY;
        items.forEach((item) => {
            const score = scoreOperatorItem(item, wanted);
            if (score > bestScore) {
                best = item;
                bestScore = score;
            }
        });
        if (!best || bestScore === Number.NEGATIVE_INFINITY) return false;
        return invokeMenuItem(best, button);
    }

    async function applyFieldFilter(ctx, spec) {
        if (!ctx || !spec || !Array.isArray(spec.selectors) || !spec.selectors.length) return false;
        const field = findOwnedFieldBySelectors(ctx, spec.selectors);
        if (!field) return false;
        if (spec.clearFirst) setFieldValue(field, '');
        if (spec.operatorLabels && spec.operatorLabels.length) setFieldOperatorByLabel(ctx, field, spec.operatorLabels);
        return setFieldValue(field, spec.value);
    }

    function findFilterFieldCombo(ctx, candidateLabels) {
        const combos = queryVisibleOwnedComponents(ctx, 'combobox').concat(queryVisibleOwnedComponents(ctx, 'uxcombobox'));
        const unique = [];
        const seen = new Set();
        combos.forEach((cmp) => {
            if (!cmp || seen.has(cmp)) return;
            seen.add(cmp);
            const name = String(cmp.name || cmp.itemId || '').toLowerCase();
            if (name.includes('dataspylist')) return;
            unique.push(cmp);
        });
        const wanted = (candidateLabels || []).map((label) => cleanText(label).toLowerCase()).filter(Boolean);
        return unique.find((combo) => {
            try {
                const store = typeof combo.getStore === 'function' ? combo.getStore() : null;
                if (!store || typeof store.each !== 'function') return false;
                let matched = false;
                store.each((record) => {
                    if (matched) return;
                    const haystack = getRecordStrings(record).join(' | ').toLowerCase();
                    if (wanted.some((label) => haystack.includes(label))) matched = true;
                });
                return matched;
            } catch (_) {
                return false;
            }
        }) || null;
    }

    function findOperatorButton(ctx) {
        const buttons = queryVisibleOwnedComponents(ctx, 'button');
        const operatorLabels = ['contains', 'begins with', 'starts with', 'is', '=', '<=', '>=', '<', '>'];
        return buttons.find((cmp) => {
            const text = cleanText(cmp.text || cmp.ariaLabel || cmp.tooltip || cmp.itemId || '').toLowerCase();
            const itemId = cleanText(cmp.itemId || cmp.name || '').toLowerCase();
            return operatorLabels.includes(text) || itemId.includes('operator') || itemId.includes('filter');
        }) || null;
    }

    function setOperatorByLabel(ctx, preferredLabels) {
        const button = findOperatorButton(ctx);
        if (!button) return false;
        const wanted = (Array.isArray(preferredLabels) ? preferredLabels : [preferredLabels]).map((label) => cleanText(label).toLowerCase()).filter(Boolean);
        const menu = button.menu || (typeof button.getMenu === 'function' ? button.getMenu() : null);
        const items = menu && menu.items && Array.isArray(menu.items.items) ? menu.items.items : [];
        const match = items.find((item) => wanted.some((label) => cleanText(item.text || item.itemId || '').toLowerCase() === label || cleanText(item.text || item.itemId || '').toLowerCase().includes(label)));
        if (!match) return false;
        try {
            if (typeof match.handler === 'function') {
                match.handler.call(match.scope || match, match);
            } else if (typeof match.fireEvent === 'function') {
                match.fireEvent('click', match);
            }
            if (typeof button.setText === 'function' && match.text) button.setText(match.text);
            return true;
        } catch (_) {
            return false;
        }
    }

    function findGenericFilterValueField(ctx) {
        const selectors = [
            'datefield',
            'uxdate',
            'uxdatetime',
            'textfield',
            'triggerfield',
            'lovfield'
        ];
        for (const selector of selectors) {
            const matches = queryVisibleOwnedComponents(ctx, selector);
            const winner = matches.find((cmp) => {
                const name = String(cmp.name || cmp.itemId || '').toLowerCase();
                if (!name) return true;
                return !name.includes('dataspylist') && !name.includes('workordernum') && !name.includes('description');
            });
            if (winner) return winner;
        }
        return null;
    }

    async function setDataspyLabel(ctx, label) {
        const combo = findDataspyCombo(ctx);
        return combo ? setComboByLabel(combo, [label]) : false;
    }

    async function applyGenericToolbarFilter(ctx, spec) {
        if (!spec || !Array.isArray(spec.fieldLabels) || !spec.fieldLabels.length) return false;
        const fieldCombo = findFilterFieldCombo(ctx, spec.fieldLabels);
        if (!fieldCombo) return false;
        const fieldSet = await setComboByLabel(fieldCombo, spec.fieldLabels);
        if (!fieldSet) return false;
        if (spec.operatorLabels && spec.operatorLabels.length) setOperatorByLabel(ctx, spec.operatorLabels);
        const valueField = findGenericFilterValueField(ctx);
        if (!valueField) return false;
        return setFieldValue(valueField, spec.value);
    }

    async function applyPmDueDateEmptyFilter(ctx) {
        const fieldLabels = ['Original PM Due Date', 'Original PM Due', 'Original Due Date', 'Due Date'];

        const selectorHit = await applyFieldFilter(ctx, {
            selectors: [
                'datefield[name=ff_duedate]',
                'uxdate[name=ff_duedate]',
                'datefield[name*=duedate]',
                'uxdate[name*=duedate]'
            ],
            operatorLabels: ['is empty', 'empty'],
            value: ''
        });
        if (selectorHit) return true;

        return applyGenericToolbarFilter(ctx, {
            fieldLabels,
            operatorLabels: ['is empty', 'empty'],
            value: ''
        });
    }

    async function applyDueWindowFilter(ctx, flow) {
        if (!flow) return false;
        const dueState = getDueWindowState(flow.key);
        const cutoff = startOfDay(addDays(new Date(), dueState.daysAhead));
        const value = formatUSDate(cutoff);

        const selectorMap = {
            audits: [
                'datefield[name=ff_auditdate]',
                'uxdate[name=ff_auditdate]',
                'datefield[name*=auditdate]',
                'uxdate[name*=auditdate]',
                'datefield[name=ff_duedate]',
                'uxdate[name=ff_duedate]',
                'datefield[name*=duedate]',
                'uxdate[name*=duedate]'
            ],
            compliance: [
                'datefield[name=ff_duedate]',
                'uxdate[name=ff_duedate]',
                'datefield[name*=duedate]',
                'uxdate[name*=duedate]'
            ],
            pms: [
                'datefield[name=ff_duedate]',
                'uxdate[name=ff_duedate]',
                'datefield[name*=duedate]',
                'uxdate[name*=duedate]'
            ],
            fwos: [
                'datefield[name=ff_schedenddate]',
                'uxdate[name=ff_schedenddate]',
                'datefield[name*=schedenddate]',
                'uxdate[name*=schedenddate]',
                'datefield[name*=enddate]',
                'uxdate[name*=enddate]'
            ]
        };

        const selectorHit = await applyFieldFilter(ctx, {
            selectors: selectorMap[flow.key] || [],
            operatorLabels: ['<=', 'less than or equals', 'on or before', 'before'],
            value
        });
        if (selectorHit) return true;

        if (!Array.isArray(flow.dueFilterLabels) || !flow.dueFilterLabels.length) return false;
        return applyGenericToolbarFilter(ctx, {
            fieldLabels: flow.dueFilterLabels,
            operatorLabels: ['<=', 'less than or equals', 'on or before', 'before'],
            value
        });
    }

    async function clearFieldFilters(ctx, selectors) {
        if (!ctx || !Array.isArray(selectors) || !selectors.length) return false;
        let cleared = false;
        for (const selector of selectors) {
            const fields = queryVisibleOwnedComponents(ctx, selector) || [];
            for (const field of fields) {
                if (setFieldValue(field, '')) cleared = true;
            }
        }
        return cleared;
    }

    async function clearGridFilterRow(ctx) {
        if (!ctx) return false;
        const button = findPreferredField(ctx, [
            'button[itemId=clearfilter]',
            'button[action=clearfilter]',
            'button[iconCls=clearFilter]'
        ]);
        if (!button) return false;
        return clickCmp(button);
    }

    async function clearGenericToolbarFilter(ctx) {
        if (!ctx) return false;
        let cleared = false;
        const valueField = findGenericFilterValueField(ctx);
        if (valueField && setFieldValue(valueField, '')) cleared = true;

        const filterCombo = findFilterFieldCombo(ctx, ['work order', 'description', 'due date', 'original pm due date', 'sched. end date']);
        if (filterCombo) {
            try {
                if (typeof filterCombo.clearValue === 'function') filterCombo.clearValue();
                else if (typeof filterCombo.setValue === 'function') filterCombo.setValue('');
                if (typeof filterCombo.fireEvent === 'function') {
                    filterCombo.fireEvent('change', filterCombo, '');
                    filterCombo.fireEvent('blur', filterCombo);
                }
                cleared = true;
            } catch (_) {}
        }
        return cleared;
    }

    async function clearFlowSpecificFilters(ctx) {
        if (!ctx) return [];
        const notes = [];
        if (await clearGridFilterRow(ctx)) notes.push('clear-filter-row');
        if (await clearGenericToolbarFilter(ctx)) notes.push('clear-toolbar-filter');
        if (clearWorkOrderField(ctx)) notes.push('clear-wo');
        const descCleared = await clearFieldFilters(ctx, [
            'textfield[name=ff_description]',
            'triggerfield[name=ff_description]',
            'textfield[name*=description]',
            'triggerfield[name*=description]',
            'textfield[name=ff_desc]',
            'triggerfield[name=ff_desc]'
        ]);
        if (descCleared) notes.push('clear-description');
        const dueCleared = await clearFieldFilters(ctx, [
            'datefield[name=ff_duedate]',
            'uxdate[name=ff_duedate]',
            'datefield[name*=duedate]',
            'uxdate[name*=duedate]',
            'datefield[name=ff_schedenddate]',
            'uxdate[name=ff_schedenddate]',
            'datefield[name*=schedenddate]',
            'uxdate[name*=schedenddate]',
            'datefield[name*=enddate]',
            'uxdate[name*=enddate]'
        ]);
        if (dueCleared) notes.push('clear-due-fields');
        return notes;
    }

    async function applyPreRunFlowFilters(flow, ctx) {
        const notes = [];
        if (!flow || !ctx) return notes;

        const cleared = await clearFlowSpecificFilters(ctx);
        if (cleared.length) notes.push(...cleared);
        if (flow.key === 'fwos') {
            const nonPmApplied = await applyPmDueDateEmptyFilter(ctx);
            if (nonPmApplied) notes.push('wo-pm-due-empty');
            const dueApplied = await applyDueWindowFilter(ctx, flow);
            if (dueApplied) notes.push('wo-end-date-lte');
            return notes;
        }

        if (flow.key === 'audits' || flow.key === 'pms' || flow.key === 'compliance') {
            const dueApplied = await applyDueWindowFilter(ctx, flow);
            if (dueApplied) notes.push(`${flow.key}-due-date-lte`);
        }
        return notes;
    }

    function clickCmp(cmp) {
        if (!cmp) return false;
        try {
            if (typeof cmp.fireEvent === 'function' && cmp.fireEvent('click', cmp) === false) {
                return false;
            }
        } catch (_) {}

        try {
            if (typeof cmp.handler === 'function') {
                cmp.handler.call(cmp.scope || cmp, cmp);
                return true;
            }
        } catch (_) {}

        try {
            if (typeof cmp.fireEvent === 'function') {
                cmp.fireEvent('tap', cmp);
                cmp.fireEvent('click', cmp);
                return true;
            }
        } catch (_) {}

        try {
            const el = typeof cmp.getEl === 'function' ? cmp.getEl() : cmp.el;
            if (el && el.dom && typeof el.dom.click === 'function') {
                el.dom.click();
                return true;
            }
        } catch (_) {}

        return false;
    }

    async function waitForGridReady(ctx, flow, timeoutMs) {
        return waitFor(() => {
            const currentCtx = getActiveAPMContext();
            if (!flowMatchesContext(flow, currentCtx) || isAppMasked(currentCtx)) return null;
            const grid = getActiveGrid(currentCtx, flow);
            const runBtn = getRunButton(currentCtx);
            return grid && runBtn ? { ctx: currentCtx, grid, runBtn } : null;
        }, timeoutMs, `grid and Run button for ${flow.key}`);
    }

    async function waitForStoreCycle(grid, timeoutMs) {
        const store = grid && typeof grid.getStore === 'function' ? grid.getStore() : null;
        if (!store) return;

        if (store.isLoading && store.isLoading()) {
            await waitFor(() => !(store.isLoading && store.isLoading()), timeoutMs, 'store load complete');
            return;
        }

        await new Promise((resolve) => {
            let done = false;
            const finish = () => {
                if (done) return;
                done = true;
                resolve();
            };
            const timer = setTimeout(finish, timeoutMs);
            const onLoad = () => {
                clearTimeout(timer);
                try {
                    store.un('load', onLoad);
                    store.un('exception', onLoad);
                    store.un('refresh', onLoad);
                } catch (_) {}
                finish();
            };

            try {
                if (typeof store.on === 'function') {
                    store.on('load', onLoad, null, { single: true });
                    store.on('exception', onLoad, null, { single: true });
                    store.on('refresh', onLoad, null, { single: true });
                } else {
                    clearTimeout(timer);
                    finish();
                }
            } catch (_) {
                clearTimeout(timer);
                finish();
            }
        });
    }

    function normalizeKey(raw) {
        return String(raw || '')
            .replace(/[^a-z0-9]+/gi, '_')
            .replace(/^_+|_+$/g, '')
            .toLowerCase();
    }

    function isProbablyWorkOrder(value) {
        const s = String(value ?? '').trim();
        return /^\d{6,}$/.test(s) || (/^[A-Za-z0-9-]{6,}$/.test(s) && /\d/.test(s));
    }

    function parseUSDate(value) {
        if (value instanceof Date && !Number.isNaN(value.getTime())) {
            return new Date(value.getFullYear(), value.getMonth(), value.getDate());
        }
        const raw = String(value || '').trim();
        if (!raw) return null;
        const usMatch = raw.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})(?:\D.*)?$/);
        if (usMatch) {
            let year = Number(usMatch[3]);
            if (year < 100) year += 2000;
            const month = Number(usMatch[1]) - 1;
            const day = Number(usMatch[2]);
            const date = new Date(year, month, day);
            return Number.isNaN(date.getTime()) ? null : date;
        }

        const isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
        if (isoMatch) {
            const year = Number(isoMatch[1]);
            const month = Number(isoMatch[2]) - 1;
            const day = Number(isoMatch[3]);
            const date = new Date(year, month, day);
            return Number.isNaN(date.getTime()) ? null : date;
        }

        const native = new Date(raw);
        if (!Number.isNaN(native.getTime())) {
            return new Date(native.getFullYear(), native.getMonth(), native.getDate());
        }

        return null;
    }

    function formatUSDate(date) {
        if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
        const mm = String(date.getMonth() + 1).padStart(2, '0');
        const dd = String(date.getDate()).padStart(2, '0');
        const yyyy = String(date.getFullYear());
        return `${mm}/${dd}/${yyyy}`;
    }

    function scoreKeys(keys, re) {
        return keys.filter((key) => re.test(String(key)));
    }

    function scoreDateField(records, key, sample = 60) {
        let score = 0;
        const count = Math.min(sample, records.length);
        for (let i = 0; i < count; i += 1) {
            if (parseUSDate(records[i] && records[i][key])) score += 1;
        }
        return score;
    }

    function detectDateField(records, keys, regexes) {
        const prioritized = keys.filter((key) => regexes.some((re) => re.test(key)));
        let best = '';
        let bestScore = -1;
        for (const key of prioritized) {
            const score = scoreDateField(records, key);
            if (score > bestScore) {
                best = key;
                bestScore = score;
            }
        }
        const minScore = Math.min(2, Math.max(1, records.length || 0));
        return bestScore >= minScore ? best : '';
    }

    function resolveColumnKeyByLabels(records, columns, labels) {
        if (!Array.isArray(labels) || !labels.length || !Array.isArray(columns) || !columns.length) return '';
        const normalizedLabels = labels.map((label) => normalizeKey(label)).filter(Boolean);
        let best = '';
        let bestScore = -1;
        for (const col of columns) {
            if (!col) continue;
            const candidates = [col.header, col.key, col.originalDataIndex].map((value) => normalizeKey(value)).filter(Boolean);
            if (!candidates.some((candidate) => normalizedLabels.includes(candidate))) continue;
            const scoreKey = col.key || normalizeKey(col.originalDataIndex || '');
            if (!scoreKey) continue;
            const score = scoreDateField(records, scoreKey);
            if (score > bestScore) {
                best = scoreKey;
                bestScore = score;
            }
        }
        const minScore = Math.min(2, Math.max(1, records.length || 0));
        return bestScore >= minScore ? best : '';
    }

    function detectTextField(records, keys, regexes, predicate) {
        const prioritized = keys.filter((key) => regexes.some((re) => re.test(key)));
        const inspect = prioritized.length ? prioritized : keys;
        let best = '';
        let bestScore = -1;
        const count = Math.min(60, records.length);
        for (const key of inspect) {
            let score = 0;
            for (let i = 0; i < count; i += 1) {
                const value = records[i] && records[i][key];
                if (predicate(value)) score += 1;
            }
            if (score > bestScore) {
                best = key;
                bestScore = score;
            }
        }
        const minScore = Math.min(2, Math.max(1, records.length || 0));
        return bestScore >= minScore ? best : '';
    }

    function getGridColumns(grid) {
        try {
            if (grid && Array.isArray(grid.columns)) return grid.columns;
            if (grid && typeof grid.getVisibleColumns === 'function') return grid.getVisibleColumns();
            if (grid && grid.headerCt && typeof grid.headerCt.getGridColumns === 'function') return grid.headerCt.getGridColumns();
        } catch (_) {}
        return [];
    }

    function collectColumnMeta(grid) {
        const columns = getGridColumns(grid);
        const meta = [];
        const visit = (col) => {
            if (!col || col.hidden) return;
            if (Array.isArray(col.items) && col.items.length && !col.dataIndex) {
                col.items.forEach(visit);
                return;
            }
            const dataIndex = normalizeKey(col.dataIndex || col.itemId || col.name || col.id || '');
            const header = cleanText(col.text || col.header || col.menuText || dataIndex);
            if (!dataIndex) return;
            meta.push({ key: dataIndex, header, originalDataIndex: col.dataIndex || '', column: col });
        };
        columns.forEach(visit);
        return meta;
    }

    function extractStoreRecords(grid) {
        const store = grid && typeof grid.getStore === 'function' ? grid.getStore() : null;
        if (!store || typeof store.getRange !== 'function') return [];
        return store.getRange().map((record) => {
            let data = {};
            try {
                data = typeof record.getData === 'function' ? record.getData() : (record.data || {});
            } catch (_) {
                data = record && record.data ? record.data : {};
            }
            const normalized = {};
            Object.keys(data || {}).forEach((key) => {
                normalized[normalizeKey(key)] = data[key];
            });
            return normalized;
        });
    }

    function detectFieldMap(flow, records, columns) {
        const keys = Array.from(new Set([
            ...Object.keys(records[0] || {}),
            ...columns.map((col) => col.key)
        ])).filter(Boolean);

        const duePatternsByFlow = {
            audits: [/audit.*date/i],
            fwos: [/sched.*end/i, /end.*date/i],
            pms: [/orig.*pm.*due/i, /pm.*due/i, /orig.*due/i],
            compliance: [/orig.*pm.*due/i, /pm.*due/i, /orig.*due/i, /due.*date/i]
        };
        const dueDateField = resolveColumnKeyByLabels(records, columns, flow.dueFilterLabels)
            || detectDateField(records, keys, duePatternsByFlow[flow.key] || [/due/i, /orig.*due/i, /audit.*date/i, /sched.*end/i, /end.*date/i]);
        const pmDueDateField = flow.key === 'fwos'
            ? (resolveColumnKeyByLabels(records, columns, ['Original PM Due Date', 'Original PM Due', 'Original Due Date', 'Due Date'])
                || detectDateField(records, keys, [/orig.*pm.*due/i, /pm.*due/i, /orig.*due/i]))
            : '';
        const workOrderField = detectTextField(records, keys, [/work.?order/i, /wonum/i, /wo_number/i], isProbablyWorkOrder);
        const descriptionField = detectTextField(records, keys, [/desc/i, /description/i], (v) => String(v || '').trim().length >= 3);
        const equipmentField = detectTextField(records, keys, [/equip/i, /asset/i, /unit/i], (v) => String(v || '').trim().length >= 2);
        const assignedField = detectTextField(records, keys, [/assigned/i, /owner/i, /planner/i, /technician/i], (v) => /[A-Za-z]/.test(String(v || '')));
        const startDateField = detectDateField(records, keys, [/start/i, /sched.*start/i]);
        const endDateField = detectDateField(records, keys, [/end/i, /sched.*end/i, /finish/i]);
        const statusField = detectTextField(records, keys, [/status/i], (v) => /[A-Za-z]/.test(String(v || '')));
        const auditField = flow.key === 'audits'
            ? detectTextField(records, keys, [/audit/i, /rme/i, /apm.?audit/i], (v) => String(v || '').trim().length >= 3)
            : '';

        return {
            dueDateField,
            pmDueDateField,
            workOrderField,
            descriptionField,
            equipmentField,
            assignedField,
            startDateField,
            endDateField,
            statusField,
            auditField
        };
    }

    function buildDisplayRows(flow, records, fieldMap) {
        const rows = records.map((record) => {
            const dueDate = parseUSDate(fieldMap.dueDateField ? record[fieldMap.dueDateField] : '');
            const pmDueDate = parseUSDate(fieldMap.pmDueDateField ? record[fieldMap.pmDueDateField] : '');
            const startDate = parseUSDate(fieldMap.startDateField ? record[fieldMap.startDateField] : '');
            const endDate = parseUSDate(fieldMap.endDateField ? record[fieldMap.endDateField] : '');
            const auditValue = fieldMap.auditField ? cleanText(record[fieldMap.auditField]) : '';
            const workOrderValue = fieldMap.workOrderField ? cleanText(record[fieldMap.workOrderField]) : '';

            return {
                dueDate,
                pmDueDate,
                endDate,
                dueText: dueDate ? formatUSDate(dueDate) : cleanText(fieldMap.dueDateField ? record[fieldMap.dueDateField] : ''),
                workOrder: workOrderValue,
                audit: auditValue,
                description: cleanText(fieldMap.descriptionField ? record[fieldMap.descriptionField] : ''),
                equipment: cleanText(fieldMap.equipmentField ? record[fieldMap.equipmentField] : ''),
                assignedTo: cleanText(fieldMap.assignedField ? record[fieldMap.assignedField] : ''),
                startText: startDate ? formatUSDate(startDate) : cleanText(fieldMap.startDateField ? record[fieldMap.startDateField] : ''),
                endText: endDate ? formatUSDate(endDate) : cleanText(fieldMap.endDateField ? record[fieldMap.endDateField] : ''),
                endAfterDue: !!(dueDate instanceof Date && endDate instanceof Date && endDate > dueDate),
                status: cleanText(fieldMap.statusField ? record[fieldMap.statusField] : ''),
                raw: record
            };
        }).filter((row) => {
            if (flow.key === 'audits') {
                return !!(row.audit || row.workOrder || row.description || row.dueText);
            }
            if (flow.key === 'fwos') {
                return !row.pmDueDate && !!(row.workOrder || row.description || row.dueText);
            }
            return !!(row.workOrder || row.description || row.dueText);
        });

        rows.sort((a, b) => {
            const aTime = a.dueDate instanceof Date ? a.dueDate.getTime() : Number.MAX_SAFE_INTEGER;
            const bTime = b.dueDate instanceof Date ? b.dueDate.getTime() : Number.MAX_SAFE_INTEGER;
            if (aTime !== bTime) return aTime - bTime;
            const aKey = flow.key === 'audits' ? (a.audit || a.workOrder) : a.workOrder;
            const bKey = flow.key === 'audits' ? (b.audit || b.workOrder) : b.workOrder;
            return String(aKey).localeCompare(String(bKey));
        });

        return rows;
    }

    function buildFlowWarnings(flow, allRows) {
        if (!flow || !Array.isArray(allRows) || !allRows.length) return [];
        if (flow.key !== 'pms') return [];
        const seen = new Set();
        const warnings = [];
        allRows.forEach((row) => {
            if (!row || !row.workOrder || !row.endAfterDue || !(row.dueDate instanceof Date) || !(row.endDate instanceof Date)) return;
            const key = `${row.workOrder}|${row.dueText}|${row.endText}`;
            if (seen.has(key)) return;
            seen.add(key);
            warnings.push({
                workOrder: row.workOrder,
                description: row.description,
                dueText: row.dueText,
                endText: row.endText
            });
        });
        warnings.sort((a, b) => String(a.dueText).localeCompare(String(b.dueText)) || String(a.workOrder).localeCompare(String(b.workOrder)));
        return warnings;
    }

    function buildOverdueWarnings(flow, allRows) {
        if (!flow || !Array.isArray(allRows) || !allRows.length) return [];
        if (!['pms', 'fwos', 'audits', 'compliance'].includes(flow.key)) return [];
        const today = startOfDay(new Date()).getTime();
        const seen = new Set();
        const warnings = [];
        allRows.forEach((row) => {
            if (!row || !(row.dueDate instanceof Date)) return;
            if (flow.key === 'pms' && row.endAfterDue) return;
            if (startOfDay(row.dueDate).getTime() >= today) return;
            const identity = flow.key === 'audits' ? (row.audit || row.workOrder) : row.workOrder;
            if (!identity) return;
            const key = `${identity}|${row.dueText}`;
            if (seen.has(key)) return;
            seen.add(key);
            warnings.push({
                workOrder: row.workOrder,
                audit: row.audit,
                description: row.description,
                dueText: row.dueText,
                equipment: row.equipment,
                assignedTo: row.assignedTo
            });
        });
        warnings.sort((a, b) => String(a.dueText).localeCompare(String(b.dueText)) || String(a.workOrder || a.audit).localeCompare(String(b.workOrder || b.audit)));
        return warnings;
    }

    function filterPrimaryRows(flow, allRows, warnings, overdueWarnings) {
        if (!flow || !Array.isArray(allRows) || !allRows.length) return [];
        const warningKeys = new Set((Array.isArray(warnings) ? warnings : []).map((row) => `${row.workOrder || row.audit || ''}|${row.dueText || ''}`));
        const overdueKeys = new Set((Array.isArray(overdueWarnings) ? overdueWarnings : []).map((row) => `${row.workOrder || row.audit || ''}|${row.dueText || ''}`));
        return allRows.filter((row) => {
            if (!row) return false;
            const key = `${row.workOrder || row.audit || ''}|${row.dueText || ''}`;
            if (warningKeys.has(key)) return false;
            if (overdueKeys.has(key)) return false;
            return true;
        });
    }

    function extractGridData(grid, flow) {
        const columns = collectColumnMeta(grid);
        const records = extractStoreRecords(grid);
        const fieldMap = detectFieldMap(flow, records, columns);
        const allRows = buildDisplayRows(flow, records, fieldMap);
        const dueWindow = getDueWindowState(flow.key);
        const warnings = buildFlowWarnings(flow, allRows);
        const overdueWarnings = buildOverdueWarnings(flow, allRows);
        const rows = filterPrimaryRows(flow, allRows, warnings, overdueWarnings);
        return { columns, records, fieldMap, rows, allRows, warnings, overdueWarnings, dueWindow };
    }

    function getModalHost() {
        const doc = window.top.document;
        let host = doc.getElementById(MODAL_HOST_ID);
        if (!host) {
            host = doc.createElement('div');
            host.id = MODAL_HOST_ID;
            doc.body.appendChild(host);
            host.attachShadow({ mode: 'open' });
        }
        host.dataset.modalHidden = host.dataset.modalHidden || 'false';
        host.dataset.modalTabLabel = host.dataset.modalTabLabel || 'Show Results';
        return host;
    }

    function getResultsModalRoot() {
        const host = window.top.document.getElementById(MODAL_HOST_ID);
        return host && host.shadowRoot ? host.shadowRoot : null;
    }

    function setResultsModalVisibility(hidden) {
        const host = window.top.document.getElementById(MODAL_HOST_ID);
        if (!host || !host.shadowRoot) return;
        const isHidden = !!hidden;
        host.dataset.modalHidden = isHidden ? 'true' : 'false';
        const overlay = host.shadowRoot.querySelector('.overlay');
        const launcher = host.shadowRoot.querySelector('.showTab');
        if (overlay) {
            overlay.hidden = isHidden;
            overlay.style.display = isHidden ? 'none' : 'flex';
            overlay.style.pointerEvents = isHidden ? 'none' : 'auto';
            overlay.setAttribute('aria-hidden', isHidden ? 'true' : 'false');
        }
        if (launcher) {
            launcher.hidden = !isHidden;
            launcher.style.display = isHidden ? 'block' : 'none';
            launcher.style.pointerEvents = isHidden ? 'auto' : 'none';
            launcher.setAttribute('aria-hidden', isHidden ? 'false' : 'true');
        }
    }

    function hideResultsModal() {
        setResultsModalVisibility(true);
    }

    function isAnyCheckerRunning() {
        const ids = [
            MY_SHIFT && MY_SHIFT.buttonId,
            FLOWS && FLOWS.audits && FLOWS.audits.buttonId,
            FLOWS && FLOWS.fwos && FLOWS.fwos.buttonId,
            FLOWS && FLOWS.compliance && FLOWS.compliance.buttonId,
            FLOWS && FLOWS.pms && FLOWS.pms.buttonId
        ].filter(Boolean);
        return ids.some((id) => {
            const button = document.getElementById(id);
            return !!(button && button.dataset && button.dataset.busy === 'true');
        });
    }

    function reopenResultsModal() {
        if (isAnyCheckerRunning()) {
            showToast('Wait for the current checker to finish before reopening saved results.', 'error');
            return;
        }
        setResultsModalVisibility(false);
    }

    function destroyResultsModal() {
        const host = window.top.document.getElementById(MODAL_HOST_ID);
        if (host) host.remove();
    }

    function closeResultsModal() {
        hideResultsModal();
    }

    function buildModalPtpBadgeHtml(row) {
        const woNumber = cleanText(row && row.workOrder || '');
        if (!woNumber || !ptpStatusTrackingEnabled()) return '';
        const history = getPtpHistorySnapshot();
        const record = history[woNumber] || null;
        const normalizedStatus = String(record && record.status || '').toUpperCase();
        const isComplete = isPtpCompletionStillValid(record);
        const isPending = !record || (normalizedStatus === 'COMPLETE' && !isComplete);
        const icon = isPending ? '⌛' : (isComplete ? '✅' : '❌');
        const statusText = isPending
            ? (record && normalizedStatus === 'COMPLETE' ? 'Completed PTP expired' : 'PTP not completed yet')
            : (normalizedStatus === 'INCOMPLETE' ? 'Incomplete PTP' : (isComplete ? 'Completed PTP' : 'Cancelled/Incomplete PTP'));
        const ptpUrl = buildPtpUrlForWorkOrder(woNumber, row && row.description || '');
        if (!ptpUrl) return '';
        const style = isPending
            ? 'display:inline-flex;align-items:center;gap:4px;margin-right:6px;padding:1px 6px;border-radius:999px;font-size:11px;font-weight:700;line-height:16px;white-space:nowrap;vertical-align:middle;cursor:pointer;flex-shrink:0;text-decoration:none;border:1px solid rgba(255, 179, 71, 0.7);background:rgba(64, 38, 8, 0.92);color:#ffd18a;'
            : (isComplete
                ? 'display:inline-flex;align-items:center;gap:4px;margin-right:6px;padding:1px 6px;border-radius:999px;font-size:11px;font-weight:700;line-height:16px;white-space:nowrap;vertical-align:middle;cursor:pointer;flex-shrink:0;text-decoration:none;border:1px solid rgba(76, 175, 80, 0.6);background:rgba(15, 48, 23, 0.9);color:#d7f7dd;'
                : 'display:inline-flex;align-items:center;gap:4px;margin-right:6px;padding:1px 6px;border-radius:999px;font-size:11px;font-weight:700;line-height:16px;white-space:nowrap;vertical-align:middle;cursor:pointer;flex-shrink:0;text-decoration:none;border:1px solid rgba(220, 92, 92, 0.62);background:rgba(66, 20, 20, 0.9);color:#ffd6d6;');
        return `<a href="${escapeHtml(ptpUrl)}" target="_blank" class="myapm-ptp-status-badge" title="${escapeHtml(statusText + ' — Click to open PTP page')}" style="${style}">${icon} PTP</a>`;
    }


    function getSummaryDateText(flow, row) {
        return cleanText(row && (row.dueText || row.auditDateText || row.auditDate || row.dueDateText || ''));
    }

    function getSummaryPtpUrl(flow, row) {
        if (!flow) return '';
        const woNumber = cleanText(row && row.workOrder || '');
        if (!woNumber || !ptpStatusTrackingEnabled()) return '';
        return buildPtpUrlForWorkOrder(woNumber, row && row.description || '');
    }

    function getSummaryWorkOrderLabel(flow, row) {
        return cleanText(row && (row.workOrder || row.audit || ''));
    }

    function getSummaryWorkOrderUrl(flow, row) {
        if (!flow || !row) return '';
        const woNumber = cleanText(row.workOrder || '');
        if (woNumber) return buildWorkOrderUrl(woNumber, flow.userFunction || 'WSJOBS');
        const auditNumber = cleanText(row.audit || '');
        return auditNumber ? buildAuditUrl(auditNumber, row.workOrder || auditNumber) : '';
    }

    function buildSummaryEntry(flow, row) {
        return {
            dateText: getSummaryDateText(flow, row),
            ptpUrl: getSummaryPtpUrl(flow, row),
            workOrderLabel: getSummaryWorkOrderLabel(flow, row),
            workOrderUrl: getSummaryWorkOrderUrl(flow, row),
            description: cleanText(row && row.description || ''),
            assignedTo: cleanText(row && row.assignedTo || '')
        };
    }

    function getSummarySectionLabel(flowOrKey) {
        const flowKey = typeof flowOrKey === 'string'
            ? flowOrKey
            : cleanText(flowOrKey && flowOrKey.key || '').toLowerCase();
        switch (flowKey) {
            case 'audits': return 'Audits';
            case 'compliance': return 'Compliance PMs';
            case 'fwos': return 'WOs';
            case 'pms': return 'PMs';
            default: return cleanText(flowOrKey && flowOrKey.modalTitle || flowKey || 'Results');
        }
    }

    function normalizeSummaryEntry(entry) {
        return {
            dateText: cleanText(entry && entry.dateText || ''),
            ptpUrl: cleanText(entry && entry.ptpUrl || ''),
            workOrderLabel: cleanText(entry && entry.workOrderLabel || ''),
            workOrderUrl: cleanText(entry && entry.workOrderUrl || ''),
            description: cleanText(entry && entry.description || ''),
            assignedTo: cleanText(entry && entry.assignedTo || '').toLowerCase()
        };
    }

    function buildSummaryClipboardPayload(entries) {
        const normalizedEntries = (Array.isArray(entries) ? entries : []).map(normalizeSummaryEntry).filter((entry) => entry.dateText || entry.workOrderLabel || entry.description || entry.assignedTo);

        const toTextLine = (entry) => {
            const parts = [entry.dateText];
            if (entry.ptpUrl) parts.push('PTP');
            parts.push(entry.workOrderLabel, entry.description, entry.assignedTo);
            return parts.filter(Boolean).join(' - ');
        };

        const toHtmlLine = (entry) => {
            const parts = [];
            if (entry.dateText) parts.push(escapeHtml(entry.dateText));
            if (entry.ptpUrl) parts.push(`<a href="${escapeHtml(entry.ptpUrl)}">PTP</a>`);
            if (entry.workOrderLabel) {
                const labelHtml = escapeHtml(entry.workOrderLabel);
                parts.push(entry.workOrderUrl ? `<a href="${escapeHtml(entry.workOrderUrl)}">${labelHtml}</a>` : labelHtml);
            }
            if (entry.description) parts.push(escapeHtml(entry.description));
            if (entry.assignedTo) parts.push(escapeHtml(entry.assignedTo));
            return parts.join(' - ');
        };

        return {
            text: normalizedEntries.map(toTextLine).join('\n'),
            html: normalizedEntries.length ? `<div>${normalizedEntries.map((entry) => `<div>${toHtmlLine(entry)}</div>`).join('')}</div>` : ''
        };
    }

    function buildSectionedSummaryClipboardPayload(groups) {
        const normalizedGroups = (Array.isArray(groups) ? groups : []).map((group) => ({
            title: cleanText(group && group.title || ''),
            entries: (Array.isArray(group && group.entries) ? group.entries : []).map(normalizeSummaryEntry).filter((entry) => entry.dateText || entry.workOrderLabel || entry.description || entry.assignedTo)
        })).filter((group) => group.title && group.entries.length);

        const htmlParts = normalizedGroups.map((group) => {
            const innerHtml = buildSummaryClipboardPayload(group.entries).html.replace(/^<div>|<\/div>$/g, '');
            return `<p><strong>${escapeHtml(group.title)}:</strong></p>${innerHtml}<p><br></p>`;
        });

        return {
            text: normalizedGroups.map((group) => `${group.title}:\n${buildSummaryClipboardPayload(group.entries).text}`).join('\n\n'),
            html: normalizedGroups.length ? htmlParts.join('').replace(/<p><br><\/p>$/, '') : ''
        };
    }

    async function writeSummaryClipboard(entriesOrGroups, options) {
        const payload = options && options.sectioned
            ? buildSectionedSummaryClipboardPayload(entriesOrGroups)
            : buildSummaryClipboardPayload(entriesOrGroups);
        if (!payload.text && !payload.html) return false;
        if (window.ClipboardItem && navigator.clipboard && typeof navigator.clipboard.write === 'function') {
            try {
                await navigator.clipboard.write([
                    new ClipboardItem({
                        'text/plain': new Blob([payload.text], { type: 'text/plain' }),
                        'text/html': new Blob([payload.html || payload.text], { type: 'text/html' })
                    })
                ]);
                return true;
            } catch (_) {}
        }
        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            await navigator.clipboard.writeText(payload.text);
            return true;
        }
        return false;
    }

    function buildLinkCellHtml(flow, row) {
        const label = flow.key === 'audits' ? (row.audit || row.workOrder) : row.workOrder;
        if (!label) return '';
        const href = flow.key === 'audits'
            ? buildAuditUrl(row.audit || row.workOrder, row.workOrder || row.audit)
            : buildWorkOrderUrl(row.workOrder, flow.userFunction || 'WSJOBS');
        const title = flow.key === 'audits' ? 'Copy Audit Link' : 'Copy Work Order Link';
        const openTitle = flow.key === 'audits'
            ? `Open Audit ${escapeHtml(label)} in a New Tab`
            : `Open Work Order ${escapeHtml(label)} in a New Tab`;
        const ptpBadgeHtml = buildModalPtpBadgeHtml(row);
        return `
            <span class="myapm-wo-inline-wrap">
                ${ptpBadgeHtml}
                <a href="${escapeHtml(href)}" target="_blank" rel="noopener noreferrer" class="myapm-wo-link" title="${openTitle}">${escapeHtml(label)}</a>
                <button type="button" class="myapm-copy-btn" data-copy-url="${escapeHtml(href)}" title="${escapeHtml(title)}" aria-label="${escapeHtml(title)}">
                    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 384 512" width="12" height="12" aria-hidden="true" focusable="false">
                        <path fill="currentColor" d="M192 0c-41.8 0-77.4 26.7-90.5 64L64 64C28.7 64 0 92.7 0 128V448c0 35.3 28.7 64 64 64h256c35.3 0 64-28.7 64-64V128c0-35.3-28.7-64-64-64h-37.5C269.4 26.7 233.8 0 192 0zm0 64a32 32 0 1 1 0 64 32 32 0 1 1 0-64zM305 273L177 401c-9.4 9.4-24.6 9.4-33.9 0L79 337c-9.4-9.4-9.4-24.6 0-33.9s24.6-9.4 33.9 0l47 47 111-111c9.4-9.4 24.6-9.4 33.9 0s9.4 24.6 0 33.9z"></path>
                    </svg>
                </button>
            </span>`;
    }

    function bindModalTextInput(inputEl, onEscape) {
        if (!inputEl) return;
        ['mousedown', 'pointerdown', 'click'].forEach((eventName) => {
            inputEl.addEventListener(eventName, (event) => {
                event.stopPropagation();
            });
        });
        inputEl.addEventListener('keydown', (event) => {
            event.stopPropagation();
            if (event.key === 'Escape') {
                event.preventDefault();
                if (typeof onEscape === 'function') onEscape();
            }
        });
        inputEl.addEventListener('keyup', (event) => {
            event.stopPropagation();
        });
    }

    function makeModalDraggable(root, modalEl, handleEl) {
        if (!root || !modalEl || !handleEl) return;
        let dragging = false;
        let startX = 0;
        let startY = 0;
        let startLeft = 0;
        let startTop = 0;

        const onMouseMove = (event) => {
            if (!dragging) return;
            const dx = event.clientX - startX;
            const dy = event.clientY - startY;
            const maxLeft = Math.max(0, window.top.innerWidth - modalEl.offsetWidth);
            const maxTop = Math.max(0, window.top.innerHeight - modalEl.offsetHeight);
            const nextLeft = Math.min(Math.max(0, startLeft + dx), maxLeft);
            const nextTop = Math.min(Math.max(0, startTop + dy), maxTop);
            modalEl.style.left = `${nextLeft}px`;
            modalEl.style.top = `${nextTop}px`;
        };

        const onMouseUp = () => {
            if (!dragging) return;
            dragging = false;
            window.top.document.removeEventListener('mousemove', onMouseMove, true);
            window.top.document.removeEventListener('mouseup', onMouseUp, true);
        };

        handleEl.addEventListener('mousedown', (event) => {
            if (event.button !== 0) return;
            const tag = event.target && event.target.tagName ? event.target.tagName.toLowerCase() : '';
            if (['button', 'input', 'select', 'textarea', 'a', 'svg', 'path'].includes(tag)) return;
            if (event.target && event.target.closest && event.target.closest('[data-resize-handle]')) return;
            event.preventDefault();
            dragging = true;
            const rect = modalEl.getBoundingClientRect();
            startX = event.clientX;
            startY = event.clientY;
            startLeft = rect.left;
            startTop = rect.top;
            modalEl.style.transform = 'none';
            modalEl.style.left = `${startLeft}px`;
            modalEl.style.top = `${startTop}px`;
            window.top.document.addEventListener('mousemove', onMouseMove, true);
            window.top.document.addEventListener('mouseup', onMouseUp, true);
        }, true);
    }

    function makeModalResizable(root, modalEl, options = {}) {
        if (!root || !modalEl) return;
        if (modalEl.dataset.resizableBound === 'true') return;
        modalEl.dataset.resizableBound = 'true';
        modalEl.style.boxSizing = 'border-box';
        modalEl.style.position = 'fixed';
        const minWidth = Number(options.minWidth || 760);
        const minHeight = Number(options.minHeight || 420);
        modalEl.style.minWidth = `${minWidth}px`;
        modalEl.style.minHeight = `${minHeight}px`;

        const handles = [
            { dir: 'n', cursor: 'ns-resize', style: { top: '0', left: '12px', right: '12px', height: '8px' } },
            { dir: 's', cursor: 'ns-resize', style: { bottom: '0', left: '12px', right: '12px', height: '8px' } },
            { dir: 'e', cursor: 'ew-resize', style: { top: '12px', right: '0', bottom: '12px', width: '8px' } },
            { dir: 'w', cursor: 'ew-resize', style: { top: '12px', left: '0', bottom: '12px', width: '8px' } },
            { dir: 'ne', cursor: 'nesw-resize', style: { top: '0', right: '0', width: '14px', height: '14px' } },
            { dir: 'nw', cursor: 'nwse-resize', style: { top: '0', left: '0', width: '14px', height: '14px' } },
            { dir: 'se', cursor: 'nwse-resize', style: { bottom: '0', right: '0', width: '16px', height: '16px' } },
            { dir: 'sw', cursor: 'nesw-resize', style: { bottom: '0', left: '0', width: '14px', height: '14px' } }
        ];

        const ensurePositioned = () => {
            const rect = modalEl.getBoundingClientRect();
            modalEl.style.transform = 'none';
            modalEl.style.left = `${rect.left}px`;
            modalEl.style.top = `${rect.top}px`;
            modalEl.style.width = `${rect.width}px`;
            modalEl.style.height = `${rect.height}px`;
        };

        const clampRect = (nextLeft, nextTop, nextWidth, nextHeight) => {
            const viewportWidth = window.top.innerWidth;
            const viewportHeight = window.top.innerHeight;
            const width = Math.max(minWidth, Math.min(nextWidth, viewportWidth - Math.max(0, nextLeft)));
            const height = Math.max(minHeight, Math.min(nextHeight, viewportHeight - Math.max(0, nextTop)));
            const left = Math.min(Math.max(0, nextLeft), Math.max(0, viewportWidth - width));
            const top = Math.min(Math.max(0, nextTop), Math.max(0, viewportHeight - height));
            return { left, top, width, height };
        };

        handles.forEach((handleConfig) => {
            const handle = root.ownerDocument.createElement('div');
            handle.setAttribute('data-resize-handle', handleConfig.dir);
            handle.style.position = 'absolute';
            handle.style.zIndex = '3';
            handle.style.userSelect = 'none';
            handle.style.touchAction = 'none';
            handle.style.cursor = handleConfig.cursor;
            handle.style.background = 'transparent';
            Object.entries(handleConfig.style).forEach(([key, value]) => {
                handle.style[key] = value;
            });
            modalEl.appendChild(handle);

            handle.addEventListener('mousedown', (event) => {
                if (event.button !== 0) return;
                event.preventDefault();
                event.stopPropagation();
                ensurePositioned();
                const startRect = modalEl.getBoundingClientRect();
                const startX = event.clientX;
                const startY = event.clientY;
                const direction = handleConfig.dir;

                const onMouseMove = (moveEvent) => {
                    const dx = moveEvent.clientX - startX;
                    const dy = moveEvent.clientY - startY;
                    let nextLeft = startRect.left;
                    let nextTop = startRect.top;
                    let nextWidth = startRect.width;
                    let nextHeight = startRect.height;

                    if (direction.includes('e')) nextWidth = startRect.width + dx;
                    if (direction.includes('s')) nextHeight = startRect.height + dy;
                    if (direction.includes('w')) {
                        nextWidth = startRect.width - dx;
                        nextLeft = startRect.left + dx;
                    }
                    if (direction.includes('n')) {
                        nextHeight = startRect.height - dy;
                        nextTop = startRect.top + dy;
                    }

                    if (nextWidth < minWidth && direction.includes('w')) nextLeft -= (minWidth - nextWidth);
                    if (nextHeight < minHeight && direction.includes('n')) nextTop -= (minHeight - nextHeight);
                    const clamped = clampRect(nextLeft, nextTop, nextWidth, nextHeight);
                    modalEl.style.left = `${clamped.left}px`;
                    modalEl.style.top = `${clamped.top}px`;
                    modalEl.style.width = `${clamped.width}px`;
                    modalEl.style.height = `${clamped.height}px`;
                };

                const onMouseUp = () => {
                    window.top.document.removeEventListener('mousemove', onMouseMove, true);
                    window.top.document.removeEventListener('mouseup', onMouseUp, true);
                };

                window.top.document.addEventListener('mousemove', onMouseMove, true);
                window.top.document.addEventListener('mouseup', onMouseUp, true);
            }, true);
        });
    }

    function showResultsModal(flow, extraction) {
        destroyResultsModal();
        const host = getModalHost();
        const root = host.shadowRoot;
        const rows = extraction.rows || [];
        const count = rows.length;
        const totalCount = Array.isArray(extraction.allRows) ? extraction.allRows.length : count;
        const dueWindowLabel = extraction.dueWindow && extraction.dueWindow.label ? extraction.dueWindow.label : 'Today';
        const warnings = Array.isArray(extraction.warnings) ? extraction.warnings : [];
        const overdueWarnings = Array.isArray(extraction.overdueWarnings) ? extraction.overdueWarnings : [];
        const subtitleLabels = {
            audits: 'Audits Due',
            compliance: 'Compliance PMs Due',
            fwos: 'WOs Due',
            pms: 'PMs Due'
        };
        const subtitleLabel = subtitleLabels[flow.key] || 'Results';
        const primaryDueHeader = flow.key === 'fwos' ? 'Due Date' : 'Original Due Date';
        const headers = flow.key === 'audits'
            ? ['Audit Date', 'WO', 'Description', 'Equipment', 'Assigned To']
            : [primaryDueHeader, 'WO', 'Description', 'Equipment', 'Start Date', 'End Date', 'Assigned To'];

        const currentUserRaw = detectCurrentUserName();
        const bodyRows = rows.map((row, rowIndex) => {
            const cells = flow.key === 'audits'
                ? [
                    row.dueText,
                    buildLinkCellHtml(flow, row),
                    escapeHtml(row.description),
                    escapeHtml(row.equipment),
                    escapeHtml(row.assignedTo)
                ]
                : [
                    row.dueText,
                    buildLinkCellHtml(flow, row),
                    escapeHtml(row.description),
                    escapeHtml(row.equipment),
                    escapeHtml(row.startText),
                    escapeHtml(row.endText),
                    escapeHtml(row.assignedTo)
                ];
            return `<tr data-row-index="${rowIndex}" data-assigned="${escapeHtml(row.assignedTo || '')}">${cells.map((cell, index) => `<td class="col-${index === 1 ? 'wo' : 'std'}">${cell || '&nbsp;'}</td>`).join('')}</tr>`;
        }).join('');

        host.dataset.modalTabLabel = 'My Shift Summary';
        root.innerHTML = `
            <style>
                .overlay { position: fixed; inset: 0; z-index: ${MODAL_Z_INDEX}; background: rgba(5,10,18,0.58); display: flex; align-items: center; justify-content: center; padding: 18px; }
                .showTab { position: fixed; right: 0; top: 50%; transform: translateY(-50%); z-index: ${MODAL_Z_INDEX}; border: 1px solid rgba(255,255,255,0.22); border-right: none; border-radius: 10px 0 0 10px; background: rgba(10, 18, 30, 0.96); color: #eaf0ff; padding: 12px 10px; font: 700 12px/1.1 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; writing-mode: vertical-rl; text-orientation: mixed; letter-spacing: 0.04em; cursor: pointer; box-shadow: 0 10px 24px rgba(0,0,0,0.35); }
                .showTab:hover { background: rgba(20, 34, 56, 0.98); }
                .modal { position: fixed; left: 50%; top: 56%; transform: translate(-50%, -50%); width: min(1420px, calc(100vw - 50px)); max-height: calc(100vh - 90px); display: flex; flex-direction: column; background: rgb(10, 18, 30); color: #eaf0ff; border: 1px solid rgba(255,255,255,0.18); border-radius: 10px; box-shadow: 0 10px 35px rgba(0,0,0,0.55); font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; overflow: hidden; }
                .header { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 10px 14px; border-bottom: 1px solid rgba(255,255,255,0.12); background: transparent; cursor: move; user-select: none; }
                .titleWrap { display: flex; flex-direction: column; gap: 2px; }
                .title { font-size: 15px; font-weight: 800; color: #eaf0ff; }
                .subtitle { font-size: 12px; color: #b8c7df; }
                .windowPill { display: inline-flex; align-items: center; gap: 6px; margin-top: 6px; width: fit-content; padding: 3px 8px; border-radius: 999px; border: 1px solid rgba(255,255,255,0.18); background: rgba(255,255,255,0.08); color: #eaf0ff; font-size: 11px; font-weight: 700; }
                .actions { display: flex; align-items: center; gap: 8px; }
                .actionBtn { cursor: pointer; border: 1px solid rgba(255,255,255,0.25); background: rgba(255,255,255,0.10); color: #fff; border-radius: 8px; padding: 4px 10px; font-size: 13px; font-weight: 700; }
                .actionBtn:hover { background: rgba(255,255,255,0.16); box-shadow: 0 2px 8px rgba(0,0,0,0.35); }
                .actionBtn.secondary { padding: 4px 8px; font-size: 11px; }
                .assignedUserInput { width: 240px; padding: 4px 8px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.25); background: rgba(255,255,255,0.08); color: #fff; font-size: 12px; box-sizing: border-box; }
                .assignedUserInput::placeholder { color: #b8c7df; }
                .body { padding: 10px 14px 14px; overflow: auto; max-height: calc(100vh - 170px); }
                .empty { color: #d6deee; font-size: 13px; padding: 14px 0; }
                .warnBox { margin: 0 0 14px; padding: 10px 12px; border-radius: 10px; border: 1px solid rgba(255,77,77,0.55); background: rgba(255,77,77,0.12); }
                .warnHeader { font-weight: 900; color: #ff7373; margin-bottom: 8px; }
                .warnTableWrap { margin-top: 8px; border: 1px solid rgba(255,77,77,0.35); border-radius: 8px; overflow: auto; }
                .warnTable thead th { background: rgba(75,18,18,0.72); color: #ffd6d6; }
                .warnTable tbody td { color: #ffe3e3; }
                .overdueBox { margin: 0 0 14px; padding: 10px 12px; border-radius: 10px; border: 1px solid rgba(255,179,71,0.65); background: rgba(255,179,71,0.13); }
                .overdueHeader { font-weight: 900; color: #ffb347; margin-bottom: 8px; }
                .overdueTableWrap { margin-top: 8px; border: 1px solid rgba(255,179,71,0.4); border-radius: 8px; overflow: auto; }
                .overdueTable thead th { background: rgba(93,56,8,0.72); color: #ffe1ad; }
                .overdueTable tbody td { color: #ffe7c2; }
                table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 12px; }
                .shift-col-date { width: 120px; }
                .shift-col-wo { width: 180px; }
                .shift-col-description { width: 340px; }
                .shift-col-equipment { width: 180px; }
                .shift-col-start { width: 120px; }
                .shift-col-end { width: 120px; }
                .shift-col-assigned { width: 150px; }
                thead th { position: sticky; top: 0; background: #0a1524; color: #eaf0ff; text-align: left; padding: 9px 8px; border-bottom: 1px solid #304258; white-space: nowrap; }
                tbody td { padding: 8px; border-bottom: 1px solid rgba(61,82,109,0.6); color: #d6deee; vertical-align: top; overflow: hidden; text-overflow: ellipsis; }
                tbody tr:hover td { background: rgba(53,80,115,0.18); }
                .col-wo { white-space: nowrap; overflow: visible; text-overflow: clip; }
                .warnTable th:empty, .warnTable td:empty, .overdueTable th:empty, .overdueTable td:empty { color: transparent; }
                .myapm-wo-inline-wrap { display: inline-flex; align-items: center; gap: 4px; white-space: nowrap; }
                .myapm-wo-link { color: #7fb7ff; text-decoration: underline; font-weight: 700; }
                .myapm-copy-btn { display: inline-flex; align-items: center; justify-content: center; width: 18px; height: 18px; padding: 1px; border-radius: 4px; border: 1px solid rgba(255,179,71,0.75); background: rgba(228,121,17,0.18); color: #ffb347; cursor: pointer; }
                .myapm-copy-btn:hover { background: rgba(228,121,17,0.34); color: #ffd18a; }
                .actionBtn[data-active="true"] { background: rgba(20,110,180,0.35); border-color: rgba(20,110,180,0.65); color: #fff; }
                .assignedEmpty { margin-top: 10px; color: #d6deee; font-size: 12px; }
            </style>
            <button type="button" class="showTab" data-action="show" hidden>${escapeHtml(flow.modalTitle || 'Show Results')}</button>
            <div class="overlay">
                <div class="modal" role="dialog" aria-modal="true" aria-label="${escapeHtml(flow.modalTitle)}">
                    <div class="header">
                        <div class="titleWrap">
                            <div class="title">${escapeHtml(flow.modalTitle)}</div>
                            <div class="subtitle">${count} ${count === 1 ? 'Record Found' : 'Records Found'}</div>
                            <div class="windowPill">Due window: ${escapeHtml(dueWindowLabel)}</div>
                        </div>
                        <div class="actions">
                            <button type="button" class="actionBtn" data-action="assigned-to-me">My PMs</button>
                            <input type="text" class="assignedUserInput" data-role="assigned-user-input" name="assignee-filter" placeholder="Assigned to:" aria-label="Assigned to filter" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" data-lpignore="true">
                            <button type="button" class="actionBtn" data-action="copy-summary">Copy Summary</button>
                            <button type="button" class="actionBtn" data-action="hide">Hide</button>
                        </div>
                    </div>
                    <div class="body">
                        ${warnings.length ? `
                            <div class="warnBox">
                                <div class="warnHeader">WARNING: The following workorders are scheduled past the original due date!</div>
                                <div class="warnTableWrap">
                                    <table class="warnTable">
                                        <thead><tr><th>Original Due Date</th><th>WO</th><th>Description</th><th>End Date</th></tr></thead>
                                        <tbody>${warnings.map((row) => `<tr><td>${escapeHtml(row.dueText)}</td><td class="col-wo">${buildLinkCellHtml(flow, { workOrder: row.workOrder })}</td><td>${escapeHtml(row.description)}</td><td>${escapeHtml(row.endText)}</td></tr>`).join('')}</tbody>
                                    </table>
                                </div>
                            </div>` : ''}
                        ${overdueWarnings.length ? `
                            <div class="overdueBox">
                                <div class="overdueHeader">OVERDUE: The following ${flow.key === 'audits' ? 'audits' : (flow.key === 'compliance' ? 'compliance PMs' : 'work orders')} are overdue!</div>
                                <div class="overdueTableWrap">
                                    <table class="overdueTable">
                                        <thead><tr><th>${escapeHtml(flow.key === 'audits' ? 'Audit Date' : 'Due Date')}</th><th>${escapeHtml(flow.key === 'audits' ? 'Audit' : 'WO')}</th><th>Description</th>${flow.key === 'audits' ? '' : '<th>Equipment</th>'}</tr></thead>
                                        <tbody>${overdueWarnings.map((row) => `<tr><td>${escapeHtml(row.dueText)}</td><td class="col-wo">${buildLinkCellHtml(flow, row)}</td><td>${escapeHtml(row.description)}</td>${flow.key === 'audits' ? '' : `<td>${escapeHtml(row.equipment)}</td>`}</tr>`).join('')}</tbody>
                                    </table>
                                </div>
                            </div>` : ''}
                        ${count ? `<table class="resultsTable"><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead><tbody>${bodyRows}</tbody></table><div class="assignedEmpty" hidden>No rows are assigned to you in this result set.</div>` : `<div class="empty">No rows matched the current due-window setting.</div>`}
                    </div>
                </div>
            </div>`;

        const modalEl = root.querySelector('.modal');
        const headerEl = root.querySelector('.header');
        const subtitleEl = root.querySelector('.subtitle');
        const assignedBtn = root.querySelector('[data-action="assigned-to-me"]');
        const assignedUserInput = root.querySelector('[data-role="assigned-user-input"]');
        const resultsRows = Array.from(root.querySelectorAll('.resultsTable tbody tr'));
        const assignedEmptyEl = root.querySelector('.assignedEmpty');
        let customAssignedUser = '';
        const formatSubtitle = (visibleCount) => `${visibleCount} ${visibleCount === 1 ? 'Record Found' : 'Records Found'}`;
        const syncAssignedUserUi = () => {
            if (assignedUserInput && assignedUserInput.value !== (customAssignedUser || '')) {
                assignedUserInput.value = customAssignedUser || '';
            }
        };
        const getAssignedFilterValue = () => {
            if (customAssignedUser) return customAssignedUser;
            if (assignedBtn && assignedBtn.dataset.active === 'true') return currentUserRaw;
            return '';
        };
        const applyAssignedFilter = () => {
            const activeAssignedValue = getAssignedFilterValue();
            let visibleCount = 0;
            resultsRows.forEach((rowEl) => {
                const assignedValue = rowEl.getAttribute('data-assigned') || '';
                const match = !activeAssignedValue || isAssignedToCurrentUser(assignedValue, activeAssignedValue);
                rowEl.style.display = match ? '' : 'none';
                if (match) visibleCount += 1;
            });
            if (subtitleEl) subtitleEl.textContent = `${visibleCount} ${visibleCount === 1 ? 'Record Found' : 'Records Found'}`;
            if (assignedEmptyEl) assignedEmptyEl.hidden = !(activeAssignedValue && visibleCount === 0 && resultsRows.length);
            return visibleCount;
        };
        const commitAssignedUserFilter = () => {
            customAssignedUser = extractUserNameOnly(assignedUserInput ? assignedUserInput.value : '');
            if (assignedBtn) {
                assignedBtn.dataset.active = 'false';
                assignedBtn.textContent = 'Assigned to me';
            }
            syncAssignedUserUi();
            applyAssignedFilter();
        };
        if (assignedBtn) {
            assignedBtn.title = currentUserRaw ? `Filter to rows assigned to ${currentUserRaw}` : 'Filter to rows assigned to the current user';
            assignedBtn.addEventListener('click', () => {
                if (!currentUserRaw) {
                    showToast('Unable to detect current user.', 'error');
                    return;
                }
                const nextActive = assignedBtn.dataset.active === 'true' ? 'false' : 'true';
                assignedBtn.dataset.active = nextActive;
                assignedBtn.textContent = nextActive === 'true' ? 'Show All' : 'Assigned to me';
                if (nextActive === 'true') {
                    customAssignedUser = '';
                    syncAssignedUserUi();
                }
                applyAssignedFilter();
            });
        }
        if (assignedUserInput) {
            bindModalTextInput(assignedUserInput, () => {
                customAssignedUser = '';
                syncAssignedUserUi();
                applyAssignedFilter();
            });
            assignedUserInput.addEventListener('input', commitAssignedUserFilter);
        }
        syncAssignedUserUi();
        applyAssignedFilter();
        makeModalDraggable(root, modalEl, headerEl);
        makeModalResizable(root, modalEl, { minWidth: 960, minHeight: 520 });
        makeModalResizable(root, modalEl, { minWidth: 900, minHeight: 480 });

        const overlayEl = root.querySelector('.overlay');
        const showTabEl = root.querySelector('[data-action="show"]');
        root.querySelector('[data-action="hide"]').addEventListener('click', hideResultsModal);
        if (showTabEl) showTabEl.addEventListener('click', reopenResultsModal);
        if (overlayEl) {
            overlayEl.addEventListener('click', (event) => {
                if (event.target === overlayEl) hideResultsModal();
            });
        }
        setResultsModalVisibility(false);
        root.querySelector('[data-action="copy-summary"]').addEventListener('click', async () => {
            const visibleIndexes = new Set(Array.from(root.querySelectorAll('.resultsTable tbody tr')).filter((el) => el.style.display !== 'none').map((el) => Number(el.getAttribute('data-row-index'))));
            const entries = rows.filter((_, index) => visibleIndexes.has(index)).map((row) => buildSummaryEntry(flow, row));
            const groups = entries.length ? [{ title: getSummarySectionLabel(flow), entries }] : [];
            try {
                const copied = await writeSummaryClipboard(groups, { sectioned: true });
                if (!copied) throw new Error('clipboard unavailable');
                showToast('Summary copied.', 'success');
            } catch (_) {
                showToast('Failed to copy summary.', 'error');
            }
        });
        root.querySelectorAll('.myapm-copy-btn').forEach((button) => {
            button.addEventListener('click', async () => {
                try {
                    await navigator.clipboard.writeText(button.getAttribute('data-copy-url') || '');
                    showToast('Link copied.', 'success');
                } catch (_) {
                    showToast('Failed to copy link.', 'error');
                }
            });
        });
    }

    function linkifyWorkorderNumbers() {
        const ctx = getActiveAPMContext();
        const doc = ctx && ctx.appWin && ctx.appWin.document ? ctx.appWin.document : document;
        const auditMode = !!(ctx && flowMatchesContext(FLOWS.audits, ctx));
        const selector = 'div.x-grid-cell-inner';
        doc.querySelectorAll(selector).forEach((el) => {
            if (el.closest && el.closest(`#${MODAL_HOST_ID}`)) return;
            const existingWraps = Array.from(el.querySelectorAll('.apm-wo-inline-wrap'));
            const baseText = existingWraps.length
                ? cleanText(Array.from(el.childNodes).map((node) => {
                    if (node.nodeType === 3) return node.textContent || '';
                    if (node.nodeType === 1 && node.matches && node.matches('.apm-wo-inline-wrap')) return node.getAttribute('data-wo-num') || node.textContent || '';
                    if (node.nodeType === 1) return node.textContent || '';
                    return '';
                }).join(' '))
                : cleanText(el.textContent || '');
            const signature = baseText;
            const hasInlineDecorations = !!el.querySelector('.apm-wo-inline-wrap, a.better-apm-workorder, .copy-btn');
            const currentCell = el.closest ? el.closest('.x-grid-cell') : null;
            const hasWorkOrderText = WORKORDER_PLAIN_REGEX.test(signature) && /^\d{11}$/.test(signature);
            if (el.dataset.workorderLinked === 'true'
                && el.dataset.workorderLinkedKey === signature
                && (!hasWorkOrderText || hasInlineDecorations)) {
                return;
            }
            if (!hasWorkOrderText && !existingWraps.length) {
                el.dataset.workorderLinked = 'true';
                el.dataset.workorderLinkedKey = signature;
                return;
            }
            if (existingWraps.length) {
                existingWraps.forEach((wrap) => {
                    const rawWo = wrap.getAttribute('data-wo-num') || cleanText(wrap.textContent || '');
                    wrap.replaceWith(doc.createTextNode(rawWo));
                });
            }
            Array.from(el.querySelectorAll('a.better-apm-workorder, .copy-btn, .ptp-inline-badge')).forEach((node) => {
                if (!node.closest || !node.closest('.apm-wo-inline-wrap')) node.remove();
            });
            el.dataset.workorderLinked = 'true';
            el.dataset.workorderLinkedKey = signature;

            const walker = doc.createTreeWalker(el, NodeFilter.SHOW_TEXT, null, false);
            const textNodes = [];
            while (walker.nextNode()) textNodes.push(walker.currentNode);

            textNodes.forEach((textNode) => {
                const text = textNode.textContent || '';
                const matches = [...text.matchAll(WORKORDER_REGEX)];
                if (!matches.length) return;
                const fragment = doc.createDocumentFragment();
                let lastIndex = 0;
                for (const match of matches) {
                    const workOrder = match[0];
                    const index = match.index || 0;
                    fragment.appendChild(doc.createTextNode(text.slice(lastIndex, index)));
                    const wrap = doc.createElement('span');
                    wrap.className = 'apm-wo-inline-wrap';
                    wrap.setAttribute('data-wo-num', workOrder);
                    const row = el.closest ? el.closest('.x-grid-item') : null;
                    if (row) {
                        const rowCells = Array.from(row.querySelectorAll('.x-grid-cell'));
                        const currentIndex = currentCell ? rowCells.indexOf(currentCell) : -1;
                        const nextCellText = cleanText((currentIndex > -1 && rowCells[currentIndex + 1]) ? (rowCells[currentIndex + 1].innerText || rowCells[currentIndex + 1].textContent || '') : '');
                        const rowCellTexts = rowCells.map((node) => cleanText(node.innerText || node.textContent || ''));
                        const rowDescription = (nextCellText && !/^\d{1,4}$/.test(nextCellText) && !/^\d{11}$/.test(nextCellText) && nextCellText.toUpperCase() !== 'PTP')
                            ? nextCellText
                            : rowCellTexts.find((value, index) => value && index !== currentIndex && !/^\d{1,4}$/.test(value) && !/^\d{11}$/.test(value) && value.toUpperCase() !== 'PTP');
                        if (rowDescription) wrap.setAttribute('data-wo-desc', rowDescription);
                    }
                    Object.assign(wrap.style, { display: 'inline-flex', alignItems: 'center', whiteSpace: 'nowrap', maxWidth: '100%' });

                    const link = doc.createElement('a');
                    link.className = 'better-apm-workorder';
                    link.href = auditMode ? buildAuditUrl(workOrder, workOrder) : buildWorkOrderUrl(workOrder, (ctx && ctx.screen && ctx.screen.getUserFunction && ctx.screen.getUserFunction()) || 'WSJOBS');
                    link.target = '_blank';
                    link.rel = 'noopener noreferrer';
                    link.textContent = workOrder;
                    link.title = auditMode ? `Open Audit ${workOrder} in a New Tab` : `Open Work Order ${workOrder} in a New Tab`;
                    Object.assign(link.style, { color: '#2196F3', textDecoration: 'underline', fontWeight: '600', flexShrink: '0' });
                    link.addEventListener('click', () => {
                        requestGridResize(auditMode ? 'audit-link-open' : 'workorder-link-open');
                        scheduleGridResizeRetries(auditMode ? 'audit-link-open-current-tab' : 'workorder-link-open-current-tab', 3, 250);
                    });

                    const copyBtn = doc.createElement('button');
                    copyBtn.className = 'copy-btn';
                    copyBtn.type = 'button';
                    copyBtn.title = auditMode ? 'Copy Audit Link' : 'Copy Work Order Link';
                    copyBtn.innerHTML = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 384 512" width="14" height="14" aria-hidden="true" focusable="false"><path fill="currentColor" d="M192 0c-41.8 0-77.4 26.7-90.5 64L64 64C28.7 64 0 92.7 0 128V448c0 35.3 28.7 64 64 64h256c35.3 0 64-28.7 64-64V128c0-35.3-28.7-64-64-64h-37.5C269.4 26.7 233.8 0 192 0zm0 64a32 32 0 1 1 0 64 32 32 0 1 1 0-64zM305 273L177 401c-9.4 9.4-24.6 9.4-33.9 0L79 337c-9.4-9.4-9.4-24.6 0-33.9s24.6-9.4 33.9 0l47 47 111-111c9.4-9.4 24.6-9.4 33.9 0s9.4 24.6 0 33.9z"></path></svg>';
                    Object.assign(copyBtn.style, {
                        marginLeft: '4px', display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                        background: 'rgba(228, 121, 17, 0.26)', border: '1px solid rgba(255, 179, 71, 0.7)',
                        borderRadius: '4px', cursor: 'pointer', width: '18px', height: '18px', padding: '1px',
                        color: '#ffb347', opacity: '1', boxShadow: '0 0 0 1px rgba(0,0,0,0.12)', transition: 'background 0.15s, border-color 0.15s, color 0.15s, box-shadow 0.15s', flexShrink: '0'
                    });
                    copyBtn.addEventListener('click', async (event) => {
                        event.preventDefault();
                        event.stopPropagation();
                        try {
                            await navigator.clipboard.writeText(link.href);
                            showToast('Link copied.', 'success');
                        } catch (_) {
                            showToast('Failed to copy link.', 'error');
                        }
                    });

                    wrap.appendChild(link);
                    wrap.appendChild(copyBtn);
                    fragment.appendChild(wrap);
                    lastIndex = index + workOrder.length;
                }
                fragment.appendChild(doc.createTextNode(text.slice(lastIndex)));
                textNode.parentNode.replaceChild(fragment, textNode);
            });
        });
    }

    async function navigateAndRun(flow) {
        log(`${flow.key} run started`);

        let ctx = getActiveAPMContext();
        if (!isRunnableFlowContext(flow, ctx)) {
            const identity = getScreenIdentity(ctx || {});
            log(`${flow.key} navigating because current context is not runnable`, {
                currentSystemFunction: identity.systemFunction || null,
                currentUserFunction: identity.userFunction || null,
                currentTabName: identity.currentTabName || null,
                activeTabTitle: identity.activeTabTitle || null
            });
            navigateToFlow(flow);
            ctx = await waitForContext(flow, NAV_TIMEOUT_MS);
        }

        ctx = await waitForUnmasked(flow, READY_TIMEOUT_MS);
        let ready = await waitForGridReady(ctx, flow, READY_TIMEOUT_MS);
        const preRunNotes = await applyPreRunFlowFilters(flow, ready.ctx);
        if (preRunNotes.length) {
            log(`${flow.key} pre-run filters applied`, { steps: preRunNotes });
            await delay(180);
            ctx = await waitForUnmasked(flow, READY_TIMEOUT_MS);
            ready = await waitForGridReady(ctx, flow, READY_TIMEOUT_MS);
        }

        const store = ready.grid && typeof ready.grid.getStore === 'function' ? ready.grid.getStore() : null;
        const beforeCount = store && typeof store.getCount === 'function' ? store.getCount() : null;
        log(`${flow.key} executing run`, { beforeCount });

        let invoked = false;
        try {
            if (ready.grid && typeof ready.grid.runDataspy === 'function') {
                ready.grid.runDataspy();
                invoked = true;
            }
        } catch (_) {}

        if (!invoked && !clickCmp(ready.runBtn)) {
            throw new Error(`Run button was found but could not be invoked for ${flow.key}`);
        }

        await waitForStoreCycle(ready.grid, STORE_TIMEOUT_MS);
        await waitForUnmasked(flow, READY_TIMEOUT_MS);

        const afterCount = store && typeof store.getCount === 'function' ? store.getCount() : null;
        log(`${flow.key} run completed`, { afterCount });
        return { ctx: ready.ctx, grid: ready.grid, afterCount };
    }


    function showCombinedResultsModal(sections) {
        destroyResultsModal();
        const host = getModalHost();
        const root = host.shadowRoot;
        const currentUserRaw = detectCurrentUserName();
        const normalizedSections = (Array.isArray(sections) ? sections : []).map((section) => ({
            flow: section.flow,
            extraction: section.extraction || {},
            rows: Array.isArray(section.extraction && section.extraction.rows) ? section.extraction.rows : [],
            warnings: Array.isArray(section.extraction && section.extraction.warnings) ? section.extraction.warnings : [],
            overdueWarnings: Array.isArray(section.extraction && section.extraction.overdueWarnings) ? section.extraction.overdueWarnings : []
        }));
        const totalCount = normalizedSections.reduce((sum, section) => sum + section.rows.length, 0);

        const buildShiftColGroup = () => '<colgroup><col class="shift-col-date"><col class="shift-col-wo"><col class="shift-col-description"><col class="shift-col-equipment"><col class="shift-col-start"><col class="shift-col-end"><col class="shift-col-assigned"></colgroup>';

        const buildShiftCells = (flow, row, mode) => {
            const dueHeader = mode === 'warning' ? 'Original Due Date' : (flow.key === 'fwos' ? 'Due Date' : (flow.key === 'audits' ? 'Audit Date' : 'Original Due Date'));
            const woCell = buildLinkCellHtml(flow, {
                workOrder: row.workOrder,
                audit: row.audit,
                description: row.description
            });
            if (flow.key === 'audits') {
                return {
                    headers: [dueHeader, 'WO', 'Description', 'Equipment', 'Start Date', 'End Date', 'Assigned To'],
                    cells: [row.dueText, woCell, escapeHtml(row.description), escapeHtml(row.equipment), '', '', escapeHtml(row.assignedTo)]
                };
            }
            if (mode === 'warning') {
                return {
                    headers: ['Original Due Date', 'WO', 'Description', 'Equipment', 'Start Date', 'End Date', 'Assigned To'],
                    cells: [row.dueText, woCell, escapeHtml(row.description), escapeHtml(row.equipment), escapeHtml(row.startText), escapeHtml(row.endText), escapeHtml(row.assignedTo)]
                };
            }
            if (mode === 'overdue') {
                return {
                    headers: ['Due Date', 'WO', 'Description', 'Equipment', 'Start Date', 'End Date', 'Assigned To'],
                    cells: [row.dueText, woCell, escapeHtml(row.description), escapeHtml(row.equipment), escapeHtml(row.startText), escapeHtml(row.endText), escapeHtml(row.assignedTo)]
                };
            }
            return {
                headers: [flow.key === 'fwos' ? 'Due Date' : 'Original Due Date', 'WO', 'Description', 'Equipment', 'Start Date', 'End Date', 'Assigned To'],
                cells: [row.dueText, woCell, escapeHtml(row.description), escapeHtml(row.equipment), escapeHtml(row.startText), escapeHtml(row.endText), escapeHtml(row.assignedTo)]
            };
        };

        const renderShiftRows = (flow, rows, mode, sectionKey) => {
            return rows.map((row, rowIndex) => {
                const config = buildShiftCells(flow, row, mode);
                return `<tr data-section="${escapeHtml(sectionKey || flow.key)}" data-row-index="${rowIndex}" data-assigned="${escapeHtml(row.assignedTo || '')}">${config.cells.map((cell, index) => `<td class="col-${index === 1 ? 'wo' : 'std'}">${cell || '&nbsp;'}</td>`).join('')}</tr>`;
            }).join('');
        };

        const renderSectionTable = (section) => {
            const flow = section.flow;
            const rows = section.rows;
            const warnings = section.warnings;
            const overdueWarnings = section.overdueWarnings;
            const primaryDueHeader = flow.key === 'fwos' ? 'Due Date' : (flow.key === 'audits' ? 'Audit Date' : 'Original Due Date');
            const headers = ['Audit Date', 'WO', 'Description', 'Equipment', 'Start Date', 'End Date', 'Assigned To'];
            if (flow.key !== 'audits') {
                headers[0] = primaryDueHeader;
            }
            const body = renderShiftRows(flow, rows, 'results', flow.key);
            const dueWindowLabel = section.extraction && section.extraction.dueWindow && section.extraction.dueWindow.label ? section.extraction.dueWindow.label : 'Today';
            return `
                <section class="combinedSection" data-flow="${escapeHtml(flow.key)}">
                    <div class="combinedSectionHeader">
                        <div class="combinedSectionTitle">${escapeHtml(flow.modalTitle)}</div>
                        <div class="combinedSectionMeta">${rows.length} ${rows.length === 1 ? 'Record' : 'Records'} • Due window: ${escapeHtml(dueWindowLabel)}</div>
                    </div>
                    ${warnings.length ? `
                        <div class="warnBox">
                            <div class="warnHeader">WARNING: The following ${flow.key === 'audits' ? 'records' : 'workorders'} are scheduled past the original due date!</div>
                            <div class="warnTableWrap">
                                <table class="warnTable">
                                    ${buildShiftColGroup()}
                                    <thead><tr>${buildShiftCells(flow, warnings[0], 'warning').headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead>
                                    <tbody>${renderShiftRows(flow, warnings, 'warning', `${flow.key}-warning`)}</tbody>
                                </table>
                            </div>
                        </div>` : ''}
                    ${overdueWarnings.length ? `
                        <div class="overdueBox">
                            <div class="overdueHeader">OVERDUE: The following ${flow.key === 'audits' ? 'audits' : (flow.key === 'compliance' ? 'compliance PMs' : 'work orders')} are overdue!</div>
                            <div class="overdueTableWrap">
                                <table class="overdueTable">
                                    ${buildShiftColGroup()}
                                    <thead><tr>${buildShiftCells(flow, overdueWarnings[0], 'overdue').headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead>
                                    <tbody>${renderShiftRows(flow, overdueWarnings, 'overdue', `${flow.key}-overdue`)}</tbody>
                                </table>
                            </div>
                        </div>` : ''}
                    ${rows.length ? `<table class="resultsTable">${buildShiftColGroup()}<thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead><tbody>${body}</tbody></table>` : `<div class="empty">No rows matched the current due-window setting.</div>`}
                </section>`;
        };

        host.dataset.modalTabLabel = 'My Shift Summary';
        root.innerHTML = `
            <style>
                .overlay { position: fixed; inset: 0; z-index: ${MODAL_Z_INDEX}; background: rgba(5,10,18,0.58); display: flex; align-items: center; justify-content: center; padding: 18px; }
                .showTab { position: fixed; right: 0; top: 50%; transform: translateY(-50%); z-index: ${MODAL_Z_INDEX}; border: 1px solid rgba(255,255,255,0.22); border-right: none; border-radius: 10px 0 0 10px; background: rgba(10, 18, 30, 0.96); color: #eaf0ff; padding: 12px 10px; font: 700 12px/1.1 system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; writing-mode: vertical-rl; text-orientation: mixed; letter-spacing: 0.04em; cursor: pointer; box-shadow: 0 10px 24px rgba(0,0,0,0.35); }
                .showTab:hover { background: rgba(20, 34, 56, 0.98); }
                .modal { position: fixed; left: 50%; top: 56%; transform: translate(-50%, -50%); width: min(1320px, 96vw); max-height: 88vh; display: flex; flex-direction: column; background: linear-gradient(to bottom, #1a2534, #111926); color: #d6deee; border: 1px solid #304258; border-radius: 8px; box-shadow: 0 10px 28px rgba(0,0,0,0.45); font: 13px/1.4 Arial, Helvetica, sans-serif; overflow: hidden; }
                .header { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 10px 14px; border-bottom: 1px solid rgba(255,255,255,0.12); background: transparent; cursor: move; user-select: none; }
                .titleWrap { display: flex; flex-direction: column; gap: 2px; }
                .title { font-size: 15px; font-weight: 800; color: #eaf0ff; }
                .subtitle { font-size: 12px; color: #b8c7df; }
                .actions { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; justify-content: flex-end; }
                .actionBtn { cursor: pointer; border: 1px solid rgba(255,255,255,0.25); background: rgba(255,255,255,0.10); color: #fff; border-radius: 8px; padding: 4px 10px; font-size: 13px; font-weight: 700; }
                .actionBtn:hover { background: rgba(255,255,255,0.16); box-shadow: 0 2px 8px rgba(0,0,0,0.35); }
                .actionBtn.secondary { padding: 4px 8px; font-size: 11px; }
                .actionBtn[data-active="true"] { background: rgba(20,110,180,0.35); border-color: rgba(20,110,180,0.65); color: #fff; }
                .assignedUserInput { width: 240px; padding: 4px 8px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.25); background: rgba(255,255,255,0.08); color: #fff; font-size: 12px; box-sizing: border-box; }
                .body { padding: 10px 14px 14px; overflow: auto; max-height: calc(100vh - 170px); }
                .combinedSection { margin-bottom: 16px; border: 1px solid rgba(61,82,109,0.65); border-radius: 8px; overflow: hidden; }
                .combinedSectionHeader { display: flex; justify-content: space-between; gap: 10px; padding: 10px 12px; background: rgba(39,54,75,0.65); border-bottom: 1px solid rgba(61,82,109,0.65); }
                .combinedSectionTitle { font-size: 14px; font-weight: 700; color: #eaf0ff; }
                .combinedSectionMeta { font-size: 11px; color: #b8c7df; }
                .empty { padding: 12px; color: #d6deee; }
                .warnBox { margin: 0 0 14px; padding: 10px 12px; border-radius: 10px; border: 1px solid rgba(255,77,77,0.55); background: rgba(255,77,77,0.12); }
                .warnHeader { font-weight: 900; color: #ff7373; margin-bottom: 8px; }
                .warnTableWrap { margin-top: 8px; border: 1px solid rgba(255,77,77,0.35); border-radius: 8px; overflow: auto; }
                .warnTable thead th { background: rgba(75,18,18,0.72); color: #ffd6d6; }
                .warnTable tbody td { color: #ffe3e3; }
                .overdueBox { margin: 0 0 14px; padding: 10px 12px; border-radius: 10px; border: 1px solid rgba(255,179,71,0.65); background: rgba(255,179,71,0.13); }
                .overdueHeader { font-weight: 900; color: #ffb347; margin-bottom: 8px; }
                .overdueTableWrap { margin-top: 8px; border: 1px solid rgba(255,179,71,0.4); border-radius: 8px; overflow: auto; }
                .overdueTable thead th { background: rgba(93,56,8,0.72); color: #ffe1ad; }
                .overdueTable tbody td { color: #ffe7c2; }
                table { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 12px; }
                .shift-col-date { width: 120px; }
                .shift-col-wo { width: 180px; }
                .shift-col-description { width: 340px; }
                .shift-col-equipment { width: 180px; }
                .shift-col-start { width: 120px; }
                .shift-col-end { width: 120px; }
                .shift-col-assigned { width: 150px; }
                thead th { position: sticky; top: 0; background: #0a1524; color: #eaf0ff; text-align: left; padding: 9px 8px; border-bottom: 1px solid #304258; white-space: nowrap; }
                tbody td { padding: 8px; border-bottom: 1px solid rgba(61,82,109,0.6); color: #d6deee; vertical-align: top; overflow: hidden; text-overflow: ellipsis; }
                tbody tr:hover td { background: rgba(53,80,115,0.18); }
                .col-wo { white-space: nowrap; overflow: visible; text-overflow: clip; }
                .warnTable th:empty, .warnTable td:empty, .overdueTable th:empty, .overdueTable td:empty { color: transparent; }
                .myapm-wo-inline-wrap { display: inline-flex; align-items: center; gap: 4px; white-space: nowrap; }
                .myapm-wo-link { color: #7fb7ff; text-decoration: underline; font-weight: 700; }
                .myapm-copy-btn { display: inline-flex; align-items: center; justify-content: center; width: 18px; height: 18px; padding: 1px; border-radius: 4px; border: 1px solid rgba(255,179,71,0.75); background: rgba(228,121,17,0.18); color: #ffb347; cursor: pointer; }
                .assignedEmpty { margin-top: 10px; color: #d6deee; font-size: 12px; }
            </style>
            <button type="button" class="showTab" data-action="show" hidden>My Shift Summary</button>
            <div class="overlay">
                <div class="modal" role="dialog" aria-modal="true" aria-label="My Shift Summary">
                    <div class="header">
                        <div class="titleWrap">
                            <div class="title">My Shift Summary</div>
                            <div class="subtitle">${totalCount} ${totalCount === 1 ? 'Record Found' : 'Records Found'}</div>
                        </div>
                        <div class="actions">
                            <button type="button" class="actionBtn" data-action="assigned-to-me">My PMs</button>
                            <input type="text" class="assignedUserInput" data-role="assigned-user-input" name="assignee-filter" placeholder="Assigned to:" aria-label="Assigned to filter" autocomplete="off" autocapitalize="off" autocorrect="off" spellcheck="false" data-lpignore="true">
                            <button type="button" class="actionBtn" data-action="copy-summary">Copy Summary</button>
                            <button type="button" class="actionBtn" data-action="hide">Hide</button>
                        </div>
                    </div>
                    <div class="body">${normalizedSections.map(renderSectionTable).join('')}<div class="assignedEmpty" hidden>No rows are assigned to this filter in the current result set.</div></div>
                </div>
            </div>`;

        const modalEl = root.querySelector('.modal');
        const headerEl = root.querySelector('.header');
        const subtitleEl = root.querySelector('.subtitle');
        const assignedBtn = root.querySelector('[data-action="assigned-to-me"]');
        const assignedUserInput = root.querySelector('[data-role="assigned-user-input"]');
        const resultsRows = Array.from(root.querySelectorAll('.resultsTable tbody tr'));
        const sectionEls = Array.from(root.querySelectorAll('.combinedSection'));
        const assignedEmptyEl = root.querySelector('.assignedEmpty');
        let customAssignedUser = '';
        const syncAssignedUserUi = () => {
            if (assignedUserInput && assignedUserInput.value !== (customAssignedUser || '')) {
                assignedUserInput.value = customAssignedUser || '';
            }
        };
        const getAssignedFilterValue = () => {
            if (customAssignedUser) return customAssignedUser;
            if (assignedBtn && assignedBtn.dataset.active === 'true') return currentUserRaw;
            return '';
        };
        const applyAssignedFilter = () => {
            const activeAssignedValue = getAssignedFilterValue();
            let visibleCount = 0;
            resultsRows.forEach((rowEl) => {
                const assignedValue = rowEl.getAttribute('data-assigned') || '';
                const match = !activeAssignedValue || isAssignedToCurrentUser(assignedValue, activeAssignedValue);
                rowEl.style.display = match ? '' : 'none';
                if (match) visibleCount += 1;
            });
            sectionEls.forEach((sectionEl) => {
                const visibleRows = Array.from(sectionEl.querySelectorAll('tbody tr')).filter((rowEl) => rowEl.style.display !== 'none');
                const emptyEl = sectionEl.querySelector('.empty');
                const tableEl = sectionEl.querySelector('.resultsTable');
                if (tableEl) tableEl.style.display = visibleRows.length ? '' : 'none';
                if (!emptyEl) {
                    const placeholder = document.createElement('div');
                    placeholder.className = 'empty';
                    placeholder.textContent = 'No rows match the current filter.';
                    sectionEl.appendChild(placeholder);
                    placeholder.hidden = visibleRows.length > 0;
                } else {
                    emptyEl.hidden = visibleRows.length > 0;
                }
                const metaEl = sectionEl.querySelector('.combinedSectionMeta');
                const flowKey = sectionEl.getAttribute('data-flow');
                const sectionData = normalizedSections.find((item) => item.flow.key === flowKey);
                const dueWindowLabel = sectionData && sectionData.extraction && sectionData.extraction.dueWindow && sectionData.extraction.dueWindow.label ? sectionData.extraction.dueWindow.label : 'Today';
                if (metaEl) metaEl.textContent = `${visibleRows.length} ${visibleRows.length === 1 ? 'Record' : 'Records'} • Due window: ${dueWindowLabel}`;
            });
            if (subtitleEl) subtitleEl.textContent = `${visibleCount} ${visibleCount === 1 ? 'Record Found' : 'Records Found'}`;
            if (assignedEmptyEl) assignedEmptyEl.hidden = !(activeAssignedValue && visibleCount === 0 && resultsRows.length);
        };
        const commitAssignedUserFilter = () => {
            customAssignedUser = extractUserNameOnly(assignedUserInput ? assignedUserInput.value : '');
            if (assignedBtn) {
                assignedBtn.dataset.active = 'false';
                assignedBtn.textContent = 'Assigned to me';
            }
            syncAssignedUserUi();
            applyAssignedFilter();
        };
        if (assignedBtn) {
            assignedBtn.title = currentUserRaw ? `Filter to rows assigned to ${currentUserRaw}` : 'Filter to rows assigned to the current user';
            assignedBtn.addEventListener('click', () => {
                if (!currentUserRaw) {
                    showToast('Unable to detect current user.', 'error');
                    return;
                }
                const nextActive = assignedBtn.dataset.active === 'true' ? 'false' : 'true';
                assignedBtn.dataset.active = nextActive;
                assignedBtn.textContent = nextActive === 'true' ? 'Show All' : 'Assigned to me';
                if (nextActive === 'true') {
                    customAssignedUser = '';
                    syncAssignedUserUi();
                }
                applyAssignedFilter();
            });
        }

        if (assignedUserInput) {
            bindModalTextInput(assignedUserInput, () => {
                customAssignedUser = '';
                syncAssignedUserUi();
                applyAssignedFilter();
            });
            assignedUserInput.addEventListener('input', commitAssignedUserFilter);
        }
        syncAssignedUserUi();
        applyAssignedFilter();
        makeModalDraggable(root, modalEl, headerEl);
        makeModalResizable(root, modalEl, { minWidth: 960, minHeight: 520 });
        const overlayEl = root.querySelector('.overlay');
        const showTabEl = root.querySelector('[data-action="show"]');
        root.querySelector('[data-action="hide"]').addEventListener('click', hideResultsModal);
        if (showTabEl) showTabEl.addEventListener('click', reopenResultsModal);
        if (overlayEl) {
            overlayEl.addEventListener('click', (event) => {
                if (event.target === overlayEl) hideResultsModal();
            });
        }
        setResultsModalVisibility(false);
        root.querySelector('[data-action="copy-summary"]').addEventListener('click', async () => {
            const groups = normalizedSections.map((sectionData) => {
                const sectionEl = Array.from(root.querySelectorAll('.combinedSection')).find((el) => el.getAttribute('data-flow') === sectionData.flow.key);
                const visibleRows = Array.from(sectionEl ? sectionEl.querySelectorAll('.resultsTable tbody tr') : []).filter((el) => el.style.display !== 'none');
                const entries = visibleRows.map((rowEl) => {
                    const rowIndex = Number(rowEl.getAttribute('data-row-index'));
                    const row = Number.isFinite(rowIndex) ? sectionData.rows[rowIndex] : null;
                    return row ? buildSummaryEntry(sectionData.flow, row) : null;
                }).filter(Boolean);
                return { title: getSummarySectionLabel(sectionData.flow), entries };
            }).filter((group) => group.entries.length);
            try {
                const copied = await writeSummaryClipboard(groups, { sectioned: true });
                if (!copied) throw new Error('clipboard unavailable');
                showToast('Summary copied.', 'success');
            } catch (_) {
                showToast('Failed to copy summary.', 'error');
            }
        });
        root.querySelectorAll('.myapm-copy-btn').forEach((button) => {
            button.addEventListener('click', async () => {
                try {
                    await navigator.clipboard.writeText(button.getAttribute('data-copy-url') || '');
                    showToast('Link copied.', 'success');
                } catch (_) {
                    showToast('Failed to copy link.', 'error');
                }
            });
        });
    }

    async function onMyShiftClick() {
        const button = document.getElementById(MY_SHIFT.buttonId);
        if (!button || button.dataset.busy === 'true') return;
        button.dataset.busy = 'true';
        button.textContent = MY_SHIFT.runningLabel;
        try {
            const sections = [];
            for (const flow of [FLOWS.audits, FLOWS.fwos, FLOWS.compliance, FLOWS.pms]) {
                const result = await navigateAndRun(flow);
                sections.push({ flow, extraction: extractGridData(result.grid, flow) });
            }
            showCombinedResultsModal(sections);
        } catch (error) {
            console.error(TRACE, error);
            log('myShift error', error && error.message ? error.message : String(error));
            showToast(`${MY_SHIFT.buttonLabel} failed.`, 'error');
        } finally {
            button.dataset.busy = 'false';
            button.textContent = MY_SHIFT.buttonLabel;
        }
    }

    async function onButtonClick(flow) {
        const button = document.getElementById(flow.buttonId);
        if (!button || button.dataset.busy === 'true') return;
        button.dataset.busy = 'true';
        button.textContent = flow.runningLabel;

        try {
            const result = await navigateAndRun(flow);
            const extraction = extractGridData(result.grid, flow);
            showResultsModal(flow, extraction);
        } catch (error) {
            console.error(TRACE, error);
            log(`${flow.key} error`, error && error.message ? error.message : String(error));
            showToast(`${flow.buttonLabel} failed.`, 'error');
        } finally {
            button.dataset.busy = 'false';
            button.textContent = flow.buttonLabel;
        }
    }

    function createToolbarButton(label, onClick, options = {}) {
        const button = document.createElement('button');
        button.type = 'button';
        button.textContent = label;
        if (options.id) button.id = options.id;
        Object.assign(button.style, {
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'center',
            gap: '4px',
            minHeight: '26px',
            padding: '2px 10px',
            font: '700 11px/22px Arial, Helvetica, sans-serif',
            background: 'linear-gradient(to bottom, #27364b, #1a2534)',
            color: '#eaf0ff',
            border: '1px solid #3d526d',
            borderRadius: '5px',
            boxShadow: '0 4px 14px rgba(0,0,0,0.35)',
            cursor: 'pointer',
            userSelect: 'none'
        });
        button.addEventListener('mouseenter', () => {
            button.style.background = 'linear-gradient(to bottom, #355073, #253a55)';
        });
        button.addEventListener('mouseleave', () => {
            if (button.dataset.active === 'true') return;
            button.style.background = 'linear-gradient(to bottom, #27364b, #1a2534)';
        });
        button.addEventListener('click', onClick);
        return button;
    }

    function getControlBar() {
        let bar = document.getElementById(CONTROL_BAR_ID);
        if (bar) return bar;
        bar = document.createElement('div');
        bar.id = CONTROL_BAR_ID;
        Object.assign(bar.style, {
            position: 'fixed',
            top: '10px',
            right: '12px',
            zIndex: CONTROL_Z_INDEX,
            display: 'flex',
            flexDirection: 'row',
            alignItems: 'center',
            justifyContent: 'flex-end',
            gap: '8px',
            flexWrap: 'wrap'
        });
        document.body.appendChild(bar);
        return bar;
    }

    function getSettingsIconMarkup() {
        return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="14" height="14" aria-hidden="true" focusable="false"><path fill="currentColor" d="M0 416c0 17.7 14.3 32 32 32l54.7 0c12.3 28.3 40.5 48 73.3 48s61-19.7 73.3-48L480 448c17.7 0 32-14.3 32-32s-14.3-32-32-32l-246.7 0c-12.3-28.3-40.5-48-73.3-48s-61 19.7-73.3 48L32 384c-17.7 0-32 14.3-32 32zm128 0a32 32 0 1 1 64 0 32 32 0 1 1 -64 0zM320 256a32 32 0 1 1 64 0 32 32 0 1 1 -64 0zm32-80c-32.8 0-61 19.7-73.3 48L32 224c-17.7 0-32 14.3-32 32s14.3 32 32 32l246.7 0c12.3 28.3 40.5 48 73.3 48s61-19.7 73.3-48l54.7 0c17.7 0 32-14.3 32-32s-14.3-32-32-32l-54.7 0c-12.3-28.3-40.5-48-73.3-48zM192 128a32 32 0 1 1 0-64 32 32 0 1 1 0 64zm73.3-64C253 35.7 224.8 16 192 16s-61 19.7-73.3 48L32 64C14.3 64 0 78.3 0 96s14.3 32 32 32l86.7 0c12.3 28.3 40.5 48 73.3 48s61-19.7 73.3-48L480 128c17.7 0 32-14.3 32-32s-14.3-32-32-32L265.3 64z"/></svg>`;
    }

    function isApmControlHost() {
        const host = String(location.hostname || '').toLowerCase();
        return host.includes('hxgnsmartcloud.com') || host.endsWith('.apm-es.gps.amazon.dev');
    }


    async function navigateDirectToWorkOrder(workOrderNum) {
        const target = String(workOrderNum || '').trim();
        if (!target) {
            showToast('Enter a work order first.', 'error');
            return;
        }
        try {
            const topEAM = PAGE_WINDOW.top && PAGE_WINDOW.top.EAM;
            if (!topEAM || !topEAM.Nav || typeof topEAM.Nav.launchScreen !== 'function') {
                throw new Error('EAM.Nav.launchScreen unavailable');
            }

            const woFlow = {
                key: 'woSearch',
                systemFunction: 'WSJOBS',
                userFunction: 'WSJOBS',
                titleHints: ['work order', 'work orders', 'jobs']
            };

            topEAM.Nav.launchScreen('WSJOBS?USER_FUNCTION_NAME=WSJOBS', null, {
                fromNav: true,
                smartCache: false,
                skipCacheCheck: true,
                disableAutoLoadMask: false
            });

            const ctx = await waitForUnmasked(woFlow, 10000);
            const field = findPreferredField(ctx, [
                'textfield[name=ff_workordernum]',
                'textfield[name*=workordernum]',
                'textfield[name*=wonum]'
            ]);
            if (!field || !setFieldValue(field, target)) {
                throw new Error('WO filter field unavailable');
            }

            await delay(120);
            const runButton = getRunButton(ctx);
            if (!runButton) {
                throw new Error('Run button unavailable');
            }

            if (typeof runButton.handler === 'function') {
                runButton.handler.call(runButton.scope || runButton, runButton);
            } else if (typeof runButton.fireEvent === 'function') {
                runButton.fireEvent('click', runButton);
            } else if (typeof runButton.el?.dom?.click === 'function') {
                runButton.el.dom.click();
            } else {
                throw new Error('Run action unavailable');
            }
        } catch (error) {
            console.error('[MyAPM] native WO navigation failed', error);
            showToast('Unable to open work order search.', 'error');
        }
    }

    function createHeaderWorkOrderSearch() {
        const wrapper = document.createElement('div');
        wrapper.id = 'myapm-header-wo-search';
        Object.assign(wrapper.style, {
            display: 'inline-flex',
            alignItems: 'center',
            gap: '6px',
            padding: '0',
            flex: '0 0 auto'
        });

        const input = document.createElement('input');
        input.id = 'myapm-header-wo-input';
        input.type = 'text';
        input.placeholder = 'WO Number...';
        input.setAttribute('aria-label', 'Enter WO and press Enter');
        Object.assign(input.style, {
            width: '140px',
            minHeight: '26px',
            padding: '2px 8px',
            border: '1px solid #71839a',
            borderRadius: '5px',
            background: '#f7f9fc',
            color: '#18212b',
            font: '700 11px/22px Consolas, Menlo, Monaco, monospace',
            boxSizing: 'border-box',
            outline: 'none'
        });
        input.addEventListener('focus', () => {
            input.style.borderColor = '#ffb347';
            input.style.boxShadow = '0 0 0 2px rgba(255, 179, 71, 0.2)';
        });
        input.addEventListener('blur', () => {
            input.style.borderColor = '#71839a';
            input.style.boxShadow = 'none';
        });
        input.addEventListener('keydown', (event) => {
            if (event.key !== 'Enter') return;
            event.preventDefault();
            navigateDirectToWorkOrder(input.value);
        });

        const button = document.createElement('button');
        button.type = 'button';
        button.id = 'myapm-header-wo-button';
        button.setAttribute('aria-label', 'Search work order');
        button.innerHTML = '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>';
        Object.assign(button.style, {
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'center',
            minWidth: '28px',
            minHeight: '26px',
            padding: '0 8px',
            background: 'linear-gradient(to bottom, #27364b, #1a2534)',
            color: '#eaf0ff',
            border: '1px solid #3d526d',
            borderRadius: '5px',
            boxShadow: '0 4px 14px rgba(0,0,0,0.35)',
            cursor: 'pointer'
        });
        button.addEventListener('mouseenter', () => {
            button.style.background = 'linear-gradient(to bottom, #355073, #253a55)';
        });
        button.addEventListener('mouseleave', () => {
            button.style.background = 'linear-gradient(to bottom, #27364b, #1a2534)';
        });
        button.addEventListener('click', (event) => {
            event.preventDefault();
            navigateDirectToWorkOrder(input.value);
        });

        wrapper.append(input, button);
        return wrapper;
    }

    function ensureHeaderWorkOrderSearch() {
        const bar = getControlBar();
        if (!bar) return;
        let wrapper = document.getElementById('myapm-header-wo-search');
        if (!wrapper) {
            wrapper = createHeaderWorkOrderSearch();
        }
        if (!bar.contains(wrapper)) {
            bar.insertBefore(wrapper, bar.firstChild || null);
        }
    }

    function createDueWindowRow(flowKey) {
        const state = getDueWindowState(flowKey);
        const config = DUE_WINDOW_CONFIG[flowKey];
        const card = document.createElement('section');
        card.className = 'myapm-settings-card';
        Object.assign(card.style, {
            display: 'flex',
            flexDirection: 'column',
            gap: '12px',
            padding: '16px',
            background: 'linear-gradient(180deg, rgba(28,43,64,0.96), rgba(17,27,40,0.96))',
            border: '1px solid rgba(88,118,156,0.45)',
            borderRadius: '14px',
            boxShadow: '0 10px 24px rgba(0,0,0,0.22)'
        });

        const head = document.createElement('div');
        Object.assign(head.style, {
            display: 'flex',
            flexDirection: 'column',
            gap: '4px'
        });

        const title = document.createElement('div');
        title.textContent = config.label;
        Object.assign(title.style, {
            margin: '0',
            fontSize: '15px',
            fontWeight: '700',
            color: '#f3f7ff',
            letterSpacing: '0.2px'
        });

        head.append(title);

        const controlWrap = document.createElement('div');
        Object.assign(controlWrap.style, {
            display: 'flex',
            flexDirection: 'column',
            gap: '10px'
        });

        const segmented = document.createElement('div');
        Object.assign(segmented.style, {
            display: 'inline-flex',
            flexWrap: 'wrap',
            alignItems: 'center',
            gap: '8px'
        });

        const baseStyle = {
            background: 'linear-gradient(180deg, #f6f8fb, #e8edf5)',
            color: '#223146',
            fontSize: '12px',
            padding: '7px 12px',
            border: '1px solid #b5c3d6',
            borderRadius: '999px',
            cursor: 'pointer',
            userSelect: 'none',
            fontWeight: '700',
            boxShadow: 'inset 0 1px 0 rgba(255,255,255,0.9)'
        };

        const makeModeBtn = (textLabel, modeValue) => {
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.textContent = textLabel;
            Object.assign(btn.style, baseStyle);
            btn.dataset.mode = modeValue;
            return btn;
        };

        const btnToday = makeModeBtn('Today', 'today');
        const btn7 = makeModeBtn('7 Days', '7');
        const btnCustom = makeModeBtn('Custom', 'custom');

        const customRow = document.createElement('div');
        Object.assign(customRow.style, {
            display: state.mode === 'custom' ? 'flex' : 'none',
            alignItems: 'center',
            flexWrap: 'wrap',
            gap: '8px',
            padding: '10px 12px',
            background: 'rgba(7, 15, 26, 0.35)',
            border: '1px solid rgba(88,118,156,0.35)',
            borderRadius: '10px'
        });

        const customLabel = document.createElement('span');
        customLabel.textContent = 'Show items due within';
        Object.assign(customLabel.style, {
            fontSize: '12px',
            color: '#c6d4ea'
        });

        const customDays = document.createElement('input');
        customDays.type = 'text';
        customDays.value = String(state.customDays);
        customDays.placeholder = 'Days';
        Object.assign(customDays.style, {
            width: '72px',
            minHeight: '30px',
            padding: '4px 10px',
            borderRadius: '8px',
            border: '1px solid #6e86a6',
            background: '#f7f9fc',
            color: '#18212b',
            font: '700 12px/20px Arial, Helvetica, sans-serif',
            boxSizing: 'border-box',
            outline: 'none'
        });
        customDays.addEventListener('focus', () => {
            customDays.style.borderColor = '#6fb3ff';
            customDays.style.boxShadow = '0 0 0 3px rgba(111,179,255,0.18)';
        });
        customDays.addEventListener('blur', () => {
            customDays.style.borderColor = '#6e86a6';
            customDays.style.boxShadow = 'none';
        });

        const customHint = document.createElement('span');
        customHint.textContent = 'days (0-365)';
        Object.assign(customHint.style, {
            fontSize: '12px',
            color: '#8fa7c7'
        });

        const applyModeVisuals = (mode) => {
            [btnToday, btn7, btnCustom].forEach((btn) => {
                const active = btn.dataset.mode === mode;
                btn.style.background = active ? 'linear-gradient(180deg, #2d8cff, #1b63d0)' : baseStyle.background;
                btn.style.color = active ? '#ffffff' : baseStyle.color;
                btn.style.border = active ? '1px solid #2d8cff' : baseStyle.border;
                btn.style.boxShadow = active ? '0 8px 18px rgba(23, 104, 214, 0.28)' : baseStyle.boxShadow;
            });
            customRow.style.display = mode === 'custom' ? 'flex' : 'none';
        };

        const commitCustom = () => {
            const value = clampDueDays(customDays.value);
            customDays.value = String(value);
            localStorage.setItem(config.customKey, String(value));
        };

        const setMode = (mode) => {
            localStorage.setItem(config.modeKey, mode);
            applyModeVisuals(mode);
        };

        btnToday.addEventListener('click', () => setMode('today'));
        btn7.addEventListener('click', () => setMode('7'));
        btnCustom.addEventListener('click', () => setMode('custom'));
        customDays.addEventListener('blur', commitCustom);
        customDays.addEventListener('keydown', (event) => {
            if (event.key === 'Enter') {
                event.preventDefault();
                commitCustom();
                customDays.blur();
            }
        });

        segmented.append(btnToday, btn7, btnCustom);
        customRow.append(customLabel, customDays, customHint);
        controlWrap.append(segmented, customRow);
        applyModeVisuals(state.mode);
        card.append(head, controlWrap);
        return card;
    }



    function normalizeReorderLabel(value) {
        return String(value || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    function loadReorderState() {
        try {
            const parsed = JSON.parse(localStorage.getItem(REORDER_STORAGE_KEY) || '{}');
            reorderState.contexts = parsed && parsed.contexts && typeof parsed.contexts === 'object' ? parsed.contexts : {};
        } catch (_) {
            reorderState.contexts = {};
        }
    }

    function saveReorderState() {
        localStorage.setItem(REORDER_STORAGE_KEY, JSON.stringify(reorderState));
    }

    function getReorderConfig(contextKey) {
        if (!reorderState.contexts[contextKey] || typeof reorderState.contexts[contextKey] !== 'object') {
            reorderState.contexts[contextKey] = { order: '', customized: false };
        }
        return reorderState.contexts[contextKey];
    }

    function orderItemsBySavedOrder(items, orderStr) {
        const savedOrder = String(orderStr || '').split(',').map((s) => s.trim()).filter(Boolean);
        if (!savedOrder.length || !Array.isArray(items) || !items.length) return Array.isArray(items) ? items.slice() : [];
        const rank = new Map(savedOrder.map((key, index) => [key, index]));
        return items.slice().sort((a, b) => {
            const aRank = rank.has(a.index) ? rank.get(a.index) : Number.MAX_SAFE_INTEGER;
            const bRank = rank.has(b.index) ? rank.get(b.index) : Number.MAX_SAFE_INTEGER;
            return aRank - bRank;
        });
    }

    function buildDefaultColumnOrder(items) {
        const used = new Set();
        const ordered = [];
        DEFAULT_COLUMN_PRIORITY.forEach((entry) => {
            const aliases = (entry.aliases || []).map(normalizeReorderLabel).filter(Boolean);
            let match = items.find((item) => {
                if (used.has(item.index)) return false;
                const textValue = normalizeReorderLabel(item.text);
                const indexValue = normalizeReorderLabel(item.index);
                return aliases.some((alias) => alias === textValue || alias === indexValue);
            });
            if (!match) {
                match = items.find((item) => {
                    if (used.has(item.index)) return false;
                    const textValue = normalizeReorderLabel(item.text);
                    const indexValue = normalizeReorderLabel(item.index);
                    return aliases.some((alias) => textValue.includes(alias) || alias.includes(textValue) || indexValue.includes(alias) || alias.includes(indexValue));
                });
            }
            if (match) {
                used.add(match.index);
                ordered.push(match.index);
            }
        });
        items.forEach((item) => {
            if (!used.has(item.index)) ordered.push(item.index);
        });
        return ordered.join(', ');
    }

    function buildDefaultRecordTabOrder(items) {
        const used = new Set();
        const ordered = [];
        DEFAULT_RECORD_TAB_PRIORITY.forEach((preferred) => {
            const match = items.find((item) => !used.has(item.index) && normalizeReorderLabel(item.text) === normalizeReorderLabel(preferred));
            if (match) {
                used.add(match.index);
                ordered.push(match.index);
            }
        });
        items.forEach((item) => {
            if (!used.has(item.index)) ordered.push(item.index);
        });
        return ordered.join(', ');
    }

    function ensureDefaultReorderForContext(contextKey, items) {
        if (!Array.isArray(items) || !items.length) return '';
        const config = getReorderConfig(contextKey);
        if (config.customized && String(config.order || '').trim()) return config.order;
        if (!String(config.order || '').trim()) {
            config.order = contextKey === 'recordTabs' ? buildDefaultRecordTabOrder(items) : buildDefaultColumnOrder(items);
            saveReorderState();
        }
        return config.order;
    }

    function getExtWindows() {
        const wins = [window.top, window];
        try {
            window.top.document.querySelectorAll('iframe').forEach((frame) => {
                try {
                    if (frame.contentWindow && frame.contentWindow.Ext) wins.push(frame.contentWindow);
                } catch (_) {}
            });
        } catch (_) {}
        return Array.from(new Set(wins.filter(Boolean)));
    }

    function getRenderedColumnsFromGrid(grid) {
        if (!grid || !grid.headerCt || !grid.headerCt.items) return [];
        const items = grid.headerCt.items.items || [];
        return items.map((col) => {
            if (!col || !col.dataIndex) return null;
            if (col.isCheckerHd || col.locked || col.xtype === 'rownumberer') return null;
            if (typeof col.isHidden === 'function' && col.isHidden()) return null;
            const textValue = cleanText(col.text || col.dataIndex);
            if (!textValue || textValue === '&#160;') return null;
            return { index: col.dataIndex, text: textValue, ref: col };
        }).filter(Boolean);
    }

    function gridColumnsMatchAliases(items, aliases) {
        const wanted = (Array.isArray(aliases) ? aliases : [aliases]).map(normalizeReorderLabel).filter(Boolean);
        if (!wanted.length || !Array.isArray(items) || !items.length) return false;
        return items.some((item) => {
            const textValue = normalizeReorderLabel(item && item.text);
            const indexValue = normalizeReorderLabel(item && item.index);
            return wanted.some((alias) => alias === textValue || alias === indexValue || textValue.includes(alias) || indexValue.includes(alias));
        });
    }

    function isExpectedGridForReorder(contextKey, items) {
        if (!Array.isArray(items) || !items.length) return false;
        if (contextKey === 'wsjobs') {
            return gridColumnsMatchAliases(items, ['Work Order', 'workordernum', 'workorder']);
        }
        if (contextKey === 'compliance') {
            return gridColumnsMatchAliases(items, ['Work Order', 'workordernum', 'workorder', 'Original PM Due Date', 'Original Due Date']);
        }
        if (contextKey === 'audits') {
            return gridColumnsMatchAliases(items, ['RME Audit', 'RME Audits', 'Audit', 'Audits', 'Work Order', 'workordernum', 'workorder']);
        }
        return true;
    }

    function isListGridContextForReorder(ctx, grid) {
        if (!ctx || !grid) return false;
        if (grid.displayDataspy === true) return true;
        if (getRunButton(ctx)) return true;
        const store = typeof grid.getStore === 'function' ? grid.getStore() : null;
        const storeId = String((store && (store.storeId || (typeof store.getStoreId === 'function' ? store.getStoreId() : ''))) || '').toLowerCase();
        if (storeId.includes('dataspy') || storeId.includes('wsjobs') || storeId.includes('ctjobs') || storeId.includes('adjobs') || storeId.includes('audit')) {
            return true;
        }
        return false;
    }

    function probeGridColumnsForFlow(flow) {
        for (const win of getExtWindows()) {
            try {
                const ctx = win === window.top ? getActiveAPMContext() : null;
                if (!ctx || !flowMatchesContext(flow, ctx)) continue;
                const grid = getActiveGrid(ctx, flow);
                const cols = getRenderedColumnsFromGrid(grid);
                if (cols.length) return { win, ctx, grid, items: cols };
            } catch (_) {}
        }
        return null;
    }

    function probeRecordTabs() {
        const ctx = getActiveAPMContext();
        if (!ctx || !ctx.screen) return null;
        const roots = [ctx.currentTab, ctx.screen].filter(Boolean);
        const seen = new Set();
        for (const root of roots) {
            let panels = [];
            try {
                panels = typeof root.query === 'function' ? (root.query('tabpanel') || []) : [];
            } catch (_) {
                panels = [];
            }
            for (const panel of panels) {
                if (!panel || seen.has(panel.id) || !isCmpVisible(panel) || !panel.items || !panel.items.items || panel.items.items.length < 2) continue;
                seen.add(panel.id);
                const items = panel.items.items.map((item) => {
                    if (!item || item.isDestroyed) return null;
                    if (item.tab && typeof item.tab.isHidden === 'function' && item.tab.isHidden()) return null;
                    const textValue = cleanText(item.title || item.text || item.itemId || '');
                    if (!textValue) return null;
                    return { index: textValue, text: textValue, ref: item };
                }).filter(Boolean);
                if (items.length >= 2) return { ctx, panel, items };
            }
        }
        return null;
    }

    function applyGridOrderForFlow(flow, contextKey) {
        const probe = probeGridColumnsForFlow(flow);
        if (!probe || !probe.grid || !probe.items.length) return;
        if (!isListGridContextForReorder(probe.ctx, probe.grid)) return;
        if (!isExpectedGridForReorder(contextKey, probe.items)) return;
        const orderStr = ensureDefaultReorderForContext(contextKey, probe.items);
        const preferredOrder = String(orderStr || '').split(',').map((s) => s.trim()).filter(Boolean);
        if (!preferredOrder.length) return;
        const headerCt = probe.grid.headerCt;
        const visibleCols = getRenderedColumnsFromGrid(probe.grid).map((item) => item.ref).filter(Boolean);
        const activePreferred = preferredOrder.map((dataIndex) => visibleCols.find((col) => col.dataIndex === dataIndex)).filter(Boolean);
        let needsMove = false;
        for (let i = 0; i < activePreferred.length; i += 1) {
            if (visibleCols[i] !== activePreferred[i]) {
                needsMove = true;
                break;
            }
        }
        if (!needsMove) return;
        let layoutsSuspended = false;
        try {
            const Ext = getTopExt();
            if (Ext && typeof Ext.suspendLayouts === 'function') {
                Ext.suspendLayouts();
                layoutsSuspended = true;
            }
            let targetIndex = 0;
            activePreferred.forEach((targetCol) => {
                while (targetIndex < headerCt.items.length) {
                    const current = headerCt.items.getAt(targetIndex);
                    if (current && (current.isCheckerHd || current.locked || current.xtype === 'rownumberer' || (typeof current.isHidden === 'function' && current.isHidden()))) {
                        targetIndex += 1;
                    } else {
                        break;
                    }
                }
                const currentIndex = headerCt.items.indexOf(targetCol);
                if (currentIndex !== -1 && currentIndex !== targetIndex) headerCt.move(currentIndex, targetIndex);
                targetIndex += 1;
            });
        } catch (error) {
            console.warn(TRACE, 'grid reorder failed', contextKey, error);
        } finally {
            try {
                const Ext = getTopExt();
                if (layoutsSuspended && Ext && typeof Ext.resumeLayouts === 'function') Ext.resumeLayouts(true);
                if (probe.grid.getView && probe.grid.getView() && typeof probe.grid.getView().refresh === 'function') probe.grid.getView().refresh();
            } catch (_) {}
        }
    }

    function applyRecordTabOrder() {
        const probe = probeRecordTabs();
        if (!probe || !probe.panel || !probe.items.length) return;
        const orderStr = ensureDefaultReorderForContext('recordTabs', probe.items);
        const preferredOrder = String(orderStr || '').split(',').map((s) => s.trim()).filter(Boolean);
        if (!preferredOrder.length) return;
        let needsMove = false;
        preferredOrder.forEach((tabName, targetIndex) => {
            const currentIndex = probe.panel.items.findIndexBy((item) => cleanText(item && (item.title || item.text || item.itemId || '')) === tabName);
            if (currentIndex !== -1 && currentIndex !== targetIndex) needsMove = true;
        });
        if (!needsMove) return;
        let layoutsSuspended = false;
        try {
            const Ext = getTopExt();
            if (Ext && typeof Ext.suspendLayouts === 'function') {
                Ext.suspendLayouts();
                layoutsSuspended = true;
            }
            preferredOrder.forEach((tabName, targetIndex) => {
                const currentIndex = probe.panel.items.findIndexBy((item) => cleanText(item && (item.title || item.text || item.itemId || '')) === tabName);
                if (currentIndex !== -1 && currentIndex !== targetIndex) probe.panel.move(currentIndex, targetIndex);
            });
        } catch (error) {
            console.warn(TRACE, 'record tab reorder failed', error);
        } finally {
            try {
                const Ext = getTopExt();
                if (layoutsSuspended && Ext && typeof Ext.resumeLayouts === 'function') Ext.resumeLayouts(true);
                if (typeof probe.panel.updateLayout === 'function') probe.panel.updateLayout();
            } catch (_) {}
        }
    }

    function applySavedLayoutOrders() {
        applyGridOrderForFlow(FLOWS.fwos, 'wsjobs');
        applyGridOrderForFlow(FLOWS.compliance, 'compliance');
        applyGridOrderForFlow(FLOWS.audits, 'audits');
        applyRecordTabOrder();
    }

    function closeReorderPanel() {
        const panel = document.getElementById(REORDER_PANEL_ID);
        if (panel) panel.style.display = 'none';
    }

    function positionReorderPanel(panel, anchorEl) {
        if (!panel) return;
        const anchor = anchorEl || document.getElementById(REORDER_BUTTON_ID);
        if (!anchor) {
            panel.style.top = '60px';
            panel.style.right = '20px';
            panel.style.left = 'auto';
            return;
        }
        const rect = anchor.getBoundingClientRect();
        const panelWidth = panel.offsetWidth || 380;
        const maxLeft = Math.max(10, window.innerWidth - panelWidth - 20);
        const desiredLeft = Math.min(maxLeft, Math.max(10, rect.left - 200));
        panel.style.top = `${Math.max(10, rect.bottom + 6)}px`;
        panel.style.left = `${desiredLeft}px`;
        panel.style.right = 'auto';
    }

    function toggleReorderPanel() {
        const panel = document.getElementById(REORDER_PANEL_ID);
        const button = document.getElementById(REORDER_BUTTON_ID);
        if (!panel) return;
        const shouldShow = panel.style.display === 'none' || panel.style.display === '';
        if (shouldShow) {
            if (typeof panel.refreshReorderPanel === 'function') panel.refreshReorderPanel();
            panel.style.display = 'flex';
            positionReorderPanel(panel, button);
            return;
        }
        panel.style.display = 'none';
    }

    function createReorderPanel() {
        if (document.getElementById(REORDER_PANEL_ID)) return;
        loadReorderState();
        const panel = document.createElement('div');
        panel.id = REORDER_PANEL_ID;
        panel.style.display = 'none';
        Object.assign(panel.style, {
            position: 'fixed',
            top: '60px',
            right: '20px',
            zIndex: REORDER_PANEL_Z_INDEX,
            width: '380px',
            maxWidth: '92vw',
            maxHeight: '78vh',
            background: '#35404a',
            border: '1px solid #2c353c',
            borderRadius: '8px',
            boxShadow: '0 8px 25px rgba(0,0,0,0.6)',
            padding: '15px',
            font: '13px/1.4 Arial, Helvetica, sans-serif',
            color: '#ffffff',
            flexDirection: 'column',
            gap: '10px'
        });

        panel.innerHTML = `
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0;">
                <h4 style="margin:0; font-size:14px; color:#ffffff; font-weight:700;">Record Layout Reorder</h4>
            </div>
            <div data-role="tabs" style="display:flex; margin-bottom:0; background:#22292f; border-radius:4px; overflow:hidden;"></div>
            <div data-role="hint" style="font-size:11px; color:#aaa; margin-bottom:0;">Drag and drop to reorder, then click save.</div>
            <div data-role="list" style="background:#22292f; border:1px solid #45535e; border-radius:4px; padding:5px; min-height:60px; max-height:260px; overflow-y:auto; margin-bottom:0;"></div>
            <div style="display:flex; gap:8px; justify-content:flex-end;">
                <button type="button" data-action="reset" style="border:1px solid #45535e; background:#34495e; color:#ffffff; border-radius:6px; padding:10px 12px; font-size:12px; font-weight:700; cursor:pointer;">Reset to Default</button>
                <button type="button" data-action="save" style="border:none; background:#2ecc71; color:#ffffff; border-radius:6px; padding:12px 14px; font-size:14px; font-weight:700; cursor:pointer;">Save Layout Order</button>
            </div>`;
        document.body.appendChild(panel);

        const tabWrap = panel.querySelector('[data-role="tabs"]');
        const listEl = panel.querySelector('[data-role="list"]');
        let activeContextKey = 'wsjobs';

        const probeForContext = (contextKey) => {
            if (contextKey === 'wsjobs') return probeGridColumnsForFlow(FLOWS.fwos);
            if (contextKey === 'compliance') return probeGridColumnsForFlow(FLOWS.compliance);
            if (contextKey === 'audits') return probeGridColumnsForFlow(FLOWS.audits);
            return probeRecordTabs();
        };

        const renderTabs = () => {
            tabWrap.innerHTML = '';
            Object.values(REORDER_CONTEXTS).forEach((context) => {
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.textContent = context.label;
                btn.dataset.context = context.key;
                Object.assign(btn.style, {
                    flex: '1',
                    textAlign: 'center',
                    padding: '8px',
                    cursor: 'pointer',
                    fontSize: '12px',
                    fontWeight: '700',
                    border: 'none',
                    background: activeContextKey === context.key ? '#3498db' : 'transparent',
                    color: activeContextKey === context.key ? '#fff' : '#7f8c8d'
                });
                btn.addEventListener('click', (event) => {
                    event.preventDefault();
                    event.stopPropagation();
                    activeContextKey = context.key;
                    renderTabs();
                    renderList();
                });
                tabWrap.appendChild(btn);
            });
        };

        const renderList = () => {
            listEl.innerHTML = '';
            const context = REORDER_CONTEXTS[activeContextKey];
            const probe = probeForContext(activeContextKey);
            const items = probe && Array.isArray(probe.items) ? probe.items.map((item) => ({ index: item.index, text: item.text })) : [];
            const orderStr = ensureDefaultReorderForContext(activeContextKey, items);
            const ordered = orderItemsBySavedOrder(items, orderStr);
            if (!ordered.length) {
                const emptyEl = document.createElement('div');
                emptyEl.textContent = context.emptyText;
                emptyEl.style.color = '#7f8c8d';
                emptyEl.style.textAlign = 'center';
                emptyEl.style.padding = '10px';
                listEl.appendChild(emptyEl);
                return;
            }

            ordered.forEach((item) => {
                const row = document.createElement('div');
                row.draggable = true;
                row.dataset.index = item.index;
                row.className = 'myapm-reorder-item';
                Object.assign(row.style, {
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: '8px',
                    padding: '8px',
                    marginBottom: '4px',
                    background: '#34495e',
                    color: '#fff',
                    borderRadius: '4px',
                    cursor: 'grab',
                    fontSize: '12px',
                    border: '1px solid #2c3e50',
                    userSelect: 'none'
                });
                row.innerHTML = `<span><b style="color:#3498db;">::</b> ${escapeHtml(item.text)}</span><span style="color:#7f8c8d; font-size:10px;">[${escapeHtml(String(item.index))}]</span>`;
                row.addEventListener('dragstart', (event) => {
                    event.dataTransfer.setData('text/plain', item.index);
                    row.classList.add('dragging');
                    row.style.opacity = '0.5';
                    row.style.background = '#f39c12';
                    row.style.borderColor = '#e67e22';
                });
                row.addEventListener('dragend', () => {
                    row.classList.remove('dragging');
                    row.style.opacity = '1';
                    row.style.background = '#34495e';
                    row.style.borderColor = '#2c3e50';
                });
                listEl.appendChild(row);
            });
        };

        const prewarmAvailableReorderContexts = () => {
            Object.keys(REORDER_CONTEXTS).forEach((contextKey) => {
                try {
                    const probe = probeForContext(contextKey);
                    const items = probe && Array.isArray(probe.items) ? probe.items.map((item) => ({ index: item.index, text: item.text })) : [];
                    if (items.length) ensureDefaultReorderForContext(contextKey, items);
                } catch (_) {}
            });
        };

        const refreshReorderPanel = () => {
            prewarmAvailableReorderContexts();
            renderTabs();
            renderList();
        };

        listEl.addEventListener('dragover', (event) => {
            event.preventDefault();
            const dragging = Array.from(listEl.children).find((child) => child.classList && child.classList.contains('dragging'));
            if (!dragging) return;
            const edgeThreshold = 36;
            const rect = listEl.getBoundingClientRect();
            if (event.clientY < rect.top + edgeThreshold) listEl.scrollTop -= 14;
            else if (event.clientY > rect.bottom - edgeThreshold) listEl.scrollTop += 14;
            const siblings = Array.from(listEl.children).filter((child) => child !== dragging);
            const nextSibling = siblings.find((sibling) => {
                const box = sibling.getBoundingClientRect();
                return event.clientY <= box.top + (box.height / 2);
            });
            if (nextSibling) listEl.insertBefore(dragging, nextSibling);
            else listEl.appendChild(dragging);
        });

        panel.querySelector('[data-action="save"]').addEventListener('click', (event) => {
            event.preventDefault();
            event.stopPropagation();
            const items = Array.from(listEl.querySelectorAll('[data-index]'));
            if (!items.length) {
                showToast('No layout items to save.', 'error');
                return;
            }
            const config = getReorderConfig(activeContextKey);
            config.order = items.map((item) => item.dataset.index).join(', ');
            config.customized = true;
            saveReorderState();
            applySavedLayoutOrders();
            showToast('Layout order saved.', 'success');
        });

        panel.querySelector('[data-action="reset"]').addEventListener('click', (event) => {
            event.preventDefault();
            event.stopPropagation();
            const config = getReorderConfig(activeContextKey);
            config.order = '';
            config.customized = false;
            saveReorderState();
            renderList();
            applySavedLayoutOrders();
            showToast('Layout reset to default.', 'success');
        });

        if (!window._myapmReorderPanelCloseListener) {
            window._myapmReorderPanelCloseListener = function(event) {
                const openPanel = document.getElementById(REORDER_PANEL_ID);
                if (!openPanel || openPanel.style.display === 'none') return;
                const reorderBtn = document.getElementById(REORDER_BUTTON_ID);
                const path = event && typeof event.composedPath === 'function' ? event.composedPath() : [];
                if (path.includes(openPanel) || (reorderBtn && path.includes(reorderBtn))) return;
                if (event && event.target && typeof event.target.closest === 'function') {
                    if (event.target.closest(`#${REORDER_PANEL_ID}, #${REORDER_BUTTON_ID}`)) return;
                }
                if (reorderBtn && reorderBtn.contains(event.target)) return;
                openPanel.style.display = 'none';
            };
            document.addEventListener('pointerdown', window._myapmReorderPanelCloseListener, true);
        }

        panel.refreshReorderPanel = refreshReorderPanel;
        refreshReorderPanel();
    }


    function createPtpTimerCard() {
        const card = document.createElement('section');
        card.className = 'myapm-settings-card myapm-settings-card-ptp';
        Object.assign(card.style, {
            display: 'flex',
            flexDirection: 'column',
            gap: '12px',
            padding: '16px',
            background: 'linear-gradient(180deg, rgba(28,43,64,0.96), rgba(17,27,40,0.96))',
            border: '1px solid rgba(88,118,156,0.45)',
            borderRadius: '14px',
            boxShadow: '0 10px 24px rgba(0,0,0,0.22)'
        });

        const title = document.createElement('div');
        title.textContent = 'PTP Timer';
        Object.assign(title.style, {
            margin: '0',
            fontSize: '15px',
            fontWeight: '700',
            color: '#f3f7ff',
            letterSpacing: '0.2px'
        });

        const toggleRow = document.createElement('label');
        Object.assign(toggleRow.style, {
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: '12px',
            padding: '10px 12px',
            borderRadius: '12px',
            background: 'rgba(7, 15, 26, 0.35)',
            border: '1px solid rgba(88,118,156,0.35)',
            cursor: 'pointer'
        });

        const left = document.createElement('span');
        left.textContent = 'Enabled';
        Object.assign(left.style, {
            fontSize: '13px',
            fontWeight: '700',
            color: '#dce7f7'
        });

        const input = document.createElement('input');
        input.type = 'checkbox';
        input.checked = localStorage.getItem('apmPTPTimer') !== 'false';
        Object.assign(input.style, {
            width: '18px',
            height: '18px',
            cursor: 'pointer'
        });

        input.addEventListener('change', () => {
            const enabled = !!input.checked;
            localStorage.setItem('apmPTPTimer', enabled ? 'true' : 'false');
            if (enabled) {
                showToast('PTP Timer enabled.', 'success');
                window.postMessage({ type: 'MYAPM_PTP_SETTINGS_CHANGED', enabled: true, source: 'settings' }, location.origin);
                window.postMessage({ __myApmPtpMsgV1: true, ptpTimer: 'start', source: 'settings' }, location.origin);
            } else {
                showToast('PTP Timer disabled.', 'success');
                window.postMessage({ type: 'MYAPM_PTP_SETTINGS_CHANGED', enabled: false, source: 'settings' }, location.origin);
                window.postMessage({ __myApmPtpMsgV1: true, ptpTimer: 'reset', source: 'settings' }, location.origin);
                const existing = document.getElementById('ptp-timer');
                if (existing) existing.remove();
            }
        });

        toggleRow.append(left, input);

        const statusRow = document.createElement('label');
        Object.assign(statusRow.style, {
            display: 'inline-flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: '12px',
            padding: '10px 12px',
            borderRadius: '12px',
            background: 'rgba(7, 15, 26, 0.35)',
            border: '1px solid rgba(88,118,156,0.35)',
            cursor: 'pointer'
        });

        const statusLeft = document.createElement('span');
        statusLeft.textContent = 'PTP Status Icons';
        Object.assign(statusLeft.style, {
            fontSize: '13px',
            fontWeight: '700',
            color: '#dce7f7'
        });

        const statusInput = document.createElement('input');
        statusInput.type = 'checkbox';
        statusInput.checked = localStorage.getItem('myapmPtpStatusTracking') !== 'false';
        Object.assign(statusInput.style, {
            width: '18px',
            height: '18px',
            cursor: 'pointer'
        });
        statusInput.addEventListener('change', () => {
            const enabled = !!statusInput.checked;
            localStorage.setItem('myapmPtpStatusTracking', enabled ? 'true' : 'false');
            window.dispatchEvent(new CustomEvent('MYAPM_PTP_HISTORY_UPDATED'));
            showToast(enabled ? 'PTP status icons enabled.' : 'PTP status icons disabled.', 'success');
        });

        statusRow.append(statusLeft, statusInput);
        card.append(title, toggleRow, statusRow);
        return card;
    }

    function closeSettingsPanel() {
        const panel = document.getElementById(SETTINGS_PANEL_ID);
        if (panel) panel.style.display = 'none';
    }

    function toggleSettingsPanel() {
        const panel = document.getElementById(SETTINGS_PANEL_ID);
        if (!panel) return;
        panel.style.display = panel.style.display === 'none' ? 'flex' : 'none';
    }

    function createSettingsPanel() {
        if (document.getElementById(SETTINGS_PANEL_ID)) return;
        ensureDueWindowDefaults();
        const panel = document.createElement('div');
        panel.id = SETTINGS_PANEL_ID;
        panel.style.display = 'none';
        Object.assign(panel.style, {
            position: 'fixed',
            top: '42px',
            right: '12px',
            zIndex: SETTINGS_PANEL_Z_INDEX,
            width: '760px',
            maxWidth: '92vw',
            maxHeight: '78vh',
            overflow: 'auto',
            background: 'linear-gradient(180deg, #172334, #0f1724)',
            border: '1px solid rgba(76, 102, 135, 0.9)',
            borderRadius: '18px',
            boxShadow: '0 24px 60px rgba(0,0,0,0.46)',
            padding: '20px',
            font: '13px/1.4 Arial, Helvetica, sans-serif',
            color: '#d6deee',
            flexDirection: 'column',
            gap: '18px'
        });
        panel.innerHTML = '';

        const topRow = document.createElement('div');
        Object.assign(topRow.style, {
            display: 'flex',
            alignItems: 'flex-start',
            justifyContent: 'space-between',
            gap: '16px',
            flexWrap: 'wrap'
        });

        const titleWrap = document.createElement('div');
        Object.assign(titleWrap.style, {
            display: 'flex',
            flexDirection: 'column',
            gap: '6px',
            minWidth: '260px',
            flex: '1 1 320px'
        });

        const intro = document.createElement('div');
        Object.assign(intro.style, {
            margin: '0',
            display: 'flex',
            alignItems: 'baseline',
            gap: '10px',
            flexWrap: 'wrap',
            fontSize: '20px',
            lineHeight: '1.1',
            fontWeight: '800',
            color: '#f4f8ff'
        });

        const introTitle = document.createElement('span');
        introTitle.textContent = 'My APM Settings';

        const introVersion = document.createElement('span');
        introVersion.textContent = `version ${CURRENT_VERSION}`;
        Object.assign(introVersion.style, {
            fontSize: '12px',
            fontWeight: '600',
            color: '#9fb0c4'
        });

        intro.append(introTitle, introVersion);

        const headingRow = document.createElement('div');
        Object.assign(headingRow.style, {
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: '12px',
            flexWrap: 'wrap'
        });

        headingRow.append(intro, createUpdateBanner());
        titleWrap.append(headingRow);
        topRow.append(titleWrap);

        const topActions = document.createElement('div');
        Object.assign(topActions.style, {
            display: 'flex',
            alignItems: 'flex-start',
            justifyContent: 'flex-end',
            flex: '0 0 auto',
            minWidth: 'fit-content'
        });

        const saveBtn = createToolbarButton('Save Settings', () => {
            closeSettingsPanel();
            showToast('Settings Saved!', 'success');
        });
        saveBtn.style.font = '700 12px/22px Arial, Helvetica, sans-serif';
        saveBtn.style.padding = '4px 14px';
        saveBtn.style.minHeight = '30px';
        saveBtn.style.whiteSpace = 'nowrap';

        topActions.append(saveBtn);
        topRow.append(topActions);

        const grid = document.createElement('div');
        Object.assign(grid.style, {
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(280px, 1fr))',
            gap: '14px',
            alignItems: 'stretch'
        });

        grid.appendChild(createPtpTimerCard());
        const bookedLaborSlot = document.createElement('div');
        bookedLaborSlot.id = 'myapm-booked-labor-settings-slot';
        grid.appendChild(bookedLaborSlot);
        setTimeout(() => {
            const slot = document.getElementById('myapm-booked-labor-settings-slot');
            if (!slot || slot.childElementCount > 0) return;
            if (typeof window.createBookedLaborSettingsCard === 'function') {
                slot.appendChild(window.createBookedLaborSettingsCard());
            }
        }, 0);
        Object.keys(DUE_WINDOW_CONFIG).forEach((flowKey) => {
            grid.appendChild(createDueWindowRow(flowKey));
        });

        panel.append(topRow, grid);
        document.body.appendChild(panel);

        createReorderPanel();

        document.addEventListener('click', (event) => {
            const settingsBtn = document.getElementById('myapm-settings-button');
            if (!panel.contains(event.target) && !(settingsBtn && settingsBtn.contains(event.target))) {
                panel.style.display = 'none';
            }
        });
    }


    function makeFlowButton(flow) {
        const button = createToolbarButton(flow.buttonLabel, () => onButtonClick(flow), { id: flow.buttonId });
        log('button injected', { id: flow.buttonId, script: SCRIPT_ID });
        return button;
    }

    function getLayoutReorderAnchorRect() {
        const forecastBtn = document.getElementById('apm-forecast-ext-btn') || document.getElementById('eam-forecast-toggle');
        if (forecastBtn && forecastBtn.getBoundingClientRect().width > 0) {
            return forecastBtn.getBoundingClientRect();
        }
        const rawBtns = Array.from(document.querySelectorAll('.x-btn-mainmenuButton-toolbar-small'));
        const visibleBtns = rawBtns.filter((btn) => btn.getBoundingClientRect().width > 0);
        if (visibleBtns.length) return visibleBtns[visibleBtns.length - 1].getBoundingClientRect();
        return null;
    }

    function ensureLayoutReorderButton() {
        const anchorRect = getLayoutReorderAnchorRect();
        const button = document.getElementById(REORDER_BUTTON_ID);
        if (!anchorRect) {
            if (button) button.style.display = 'none';
            return;
        }
        if (button) {
            button.style.display = 'flex';
            button.style.left = `${anchorRect.right + 12}px`;
            button.style.top = `${anchorRect.top}px`;
            button.style.height = `${anchorRect.height || 42}px`;
            return;
        }
        const toggleBtn = document.createElement('div');
        toggleBtn.id = REORDER_BUTTON_ID;
        toggleBtn.style.cssText = `
            position: fixed;
            left: ${anchorRect.right + 12}px;
            top: ${anchorRect.top}px;
            height: ${anchorRect.height || 42}px;
            display: flex;
            align-items: center;
            cursor: pointer;
            padding: 0 10px;
            color: #d1d1d1;
            font-family: sans-serif;
            font-size: 13px;
            font-weight: 600;
            z-index: 1000;
            transition: color 0.15s;
            user-select: none;
        `;
        toggleBtn.innerHTML = 'Layout Reorder <span style="color:#e74c3c; margin-left:4px; font-weight:bold;">+</span>';
        toggleBtn.addEventListener('mouseenter', () => { toggleBtn.style.color = '#fff'; });
        toggleBtn.addEventListener('mouseleave', () => { toggleBtn.style.color = '#d1d1d1'; });
        toggleBtn.addEventListener('click', (event) => {
            event.preventDefault();
            event.stopPropagation();
            toggleReorderPanel();
        });
        document.body.appendChild(toggleBtn);
    }

    function injectButtons() {
        const bar = getControlBar();
        if (!document.getElementById('myapm-settings-button')) {
            if (!document.getElementById(MY_SHIFT.buttonId)) bar.appendChild(createToolbarButton(MY_SHIFT.buttonLabel, () => onMyShiftClick(), { id: MY_SHIFT.buttonId }));
            [FLOWS.audits, FLOWS.compliance, FLOWS.fwos, FLOWS.pms].forEach((flow) => {
                if (!document.getElementById(flow.buttonId)) bar.appendChild(makeFlowButton(flow));
            });
            const settingsBtn = createToolbarButton('My APM', () => toggleSettingsPanel(), { id: 'myapm-settings-button' });
            settingsBtn.innerHTML = `My&nbsp;APM&nbsp;${getSettingsIconMarkup()}`;
            bar.appendChild(settingsBtn);
            createSettingsPanel();
        }
        ensureLayoutReorderButton();
    }

    function bootstrap() {
        if (window.top !== window.self) return;
        if (!isApmControlHost()) return;
        ensureDueWindowDefaults();
        loadReorderState();
        injectButtons();
        ensureHeaderWorkOrderSearch();
        checkForScriptUpdates();
        applySavedLayoutOrders();
        detectCurrentUserName();
        linkifyWorkorderNumbers();
        ensureActiveRecordHeaderUi();
        refreshPtpDecorations();
        const resizedAtBootstrap = ensureReasonableWorkOrderColumnWidth();
        ensureGridResizeObserver();
        const pendingGridResize = peekGridResizeRequest(45000);
        if (pendingGridResize) {
            scheduleGridResizeRetries(pendingGridResize.reason || 'bootstrap', GRID_RESIZE_RETRY_COUNT + 8, GRID_RESIZE_RETRY_MS);
        }
        clearInterval(bootstrap._linkifyTimer);
        bootstrap._linkifyTimer = setInterval(() => {
            detectCurrentUserName();
            linkifyWorkorderNumbers();
            ensureActiveRecordHeaderUi();
            refreshPtpDecorations();
            ensureReasonableWorkOrderColumnWidth();
            const pendingResize = peekGridResizeRequest(45000);
            if (pendingResize) scheduleGridResizeRetries(pendingResize.reason || 'interval', 2, 500);
            applySavedLayoutOrders();
            ensureLayoutReorderButton();
            ensureHeaderWorkOrderSearch();
        }, LINKIFY_INTERVAL_MS);
    }

    window.addEventListener('pageshow', () => {
        const pendingResize = peekGridResizeRequest(45000);
        if (pendingResize) scheduleGridResizeRetries(pendingResize.reason || 'pageshow', GRID_RESIZE_RETRY_COUNT + 4, 500);
        ensureReasonableWorkOrderColumnWidth();
    });

    window.addEventListener('focus', () => {
        const pendingResize = peekGridResizeRequest(45000);
        if (pendingResize) scheduleGridResizeRetries(pendingResize.reason || 'focus', 6, 400);
        ensureReasonableWorkOrderColumnWidth();
    });

    let gridResizeObserver = null;
    function ensureGridResizeObserver() {
        if (gridResizeObserver || !document.body || typeof MutationObserver !== 'function') return;
        let timer = null;
        gridResizeObserver = new MutationObserver(() => {
            clearTimeout(timer);
            timer = setTimeout(() => {
                ensureReasonableWorkOrderColumnWidth();
            }, 120);
        });
        try {
            gridResizeObserver.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['style', 'class'] });
        } catch (_) {}
    }

    window.addEventListener(PTP_HISTORY_EVENT_NAME, () => {
        refreshPtpDecorations();
        ensureReasonableWorkOrderColumnWidth();
    });

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', bootstrap, { once: true });
    } else {
        bootstrap();
    }
})();


/* ==========================================================
   - Portable labor tally adapted from APM Forecast
   - Defaults to current user, supports manual user search
   ========================================================== */
(function () {
  'use strict';

  if (window.self !== window.top) return;
  const PAGE_WINDOW = typeof unsafeWindow !== 'undefined' ? unsafeWindow : window;

  const IDS = {
    style: 'apm-my-labor-style',
    root: 'apm-my-labor-settings-card',
    sum: 'apm-my-labor-sum',
    list: 'apm-my-labor-list',
    refresh: 'apm-my-labor-refresh',
    searchInput: 'apm-my-labor-search-input',
    searchStatus: 'apm-my-labor-search-status'
  };

  let laborCache = { data: [], lastFetch: 0 };
  let activeTab = 1;
  let isFetching = false;
  let laborTargetOverride = '';

  function ensureBookedLaborStyles() {
    if (document.getElementById(IDS.style)) return;
    const style = document.createElement('style');
    style.id = IDS.style;
    style.textContent = `
      #${IDS.root} {
        display: flex;
        flex-direction: column;
        min-height: 0;
      }
      #${IDS.root} .my-labor-tabs {
        display: flex;
        gap: 2px;
        background: #0f1723;
        border: 1px solid #3c516d;
        border-radius: 6px;
        overflow: hidden;
        margin-bottom: 12px;
      }
      #${IDS.root} .my-labor-tab {
        flex: 1;
        padding: 8px;
        text-align: center;
        font-size: 11px;
        cursor: pointer;
        color: #d6deee;
        font-weight: 700;
        transition: 0.2s;
        user-select: none;
      }
      #${IDS.root} .my-labor-tab.active {
        background: linear-gradient(to bottom, #146eb4, #0d4f8b);
        color: #fff;
      }
      #${IDS.root} .my-labor-total {
        font-size: 32px;
        font-weight: 700;
        text-align: center;
        margin: 10px 0;
        color: #eaf0ff;
      }
      #${IDS.root} .my-labor-row {
        display: flex;
        justify-content: space-between;
        padding: 6px 10px;
        border-bottom: 1px solid #304258;
        font-size: 12px;
        color: #d6deee;
      }
    `;
    document.head.appendChild(style);
  }

  function extractEamIdAggressive() {
    const uuidPattern = /([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i;
    if (PAGE_WINDOW.EAM && PAGE_WINDOW.EAM.AppData && PAGE_WINDOW.EAM.AppData.eamid) return PAGE_WINDOW.EAM.AppData.eamid;
    const cookieMatch = document.cookie.match(new RegExp(`eamid=${uuidPattern.source}`, 'i'));
    if (cookieMatch) return cookieMatch[1];

    try {
      const pageHTML = document.documentElement.innerHTML;
      const htmlMatch = pageHTML.match(new RegExp(`eamid["'=:%&?]+${uuidPattern.source}`, 'i'));
      if (htmlMatch) return htmlMatch[1];
    } catch (_) {}

    const frames = document.querySelectorAll('iframe');
    for (const f of frames) {
      const src = String(f.src || '');
      const srcMatch = src.match(new RegExp(`eamid=${uuidPattern.source}`, 'i'));
      if (srcMatch) return srcMatch[1];
    }
    return null;
  }

  function extractEmployeeId() {
    if (PAGE_WINDOW.EAM && PAGE_WINDOW.EAM.Context && PAGE_WINDOW.EAM.Context.employee) {
      return String(PAGE_WINDOW.EAM.Context.employee).toUpperCase();
    }
    if (PAGE_WINDOW.EAM && PAGE_WINDOW.EAM.AppData && PAGE_WINDOW.EAM.AppData.employee) {
      return String(PAGE_WINDOW.EAM.AppData.employee).toUpperCase();
    }
    const saved = localStorage.getItem('apmLogin') || '';
    return String(saved).trim().toUpperCase();
  }

  function getDefaultLaborEmployeeId() {
    const employeeId = extractEmployeeId();
    if (employeeId) return employeeId;
    return extractUserNameOnly(detectCurrentUserName()).toUpperCase();
  }

  function getEffectiveEmployeeId() {
    const override = String(laborTargetOverride || '').trim().toUpperCase();
    if (override) return override;
    return getDefaultLaborEmployeeId();
  }

  async function fetchLaborData(force = false) {
    if (!force && Date.now() - laborCache.lastFetch < 900000 && laborCache.data.length > 0) {
      updateBookedLaborUI();
      return;
    }

    isFetching = true;
    updateBookedLaborUI('Loading...');

    const currentEamId = extractEamIdAggressive();
    if (!currentEamId) {
      isFetching = false;
      updateBookedLaborUI('Session Error');
      return;
    }

    const targetEmployee = getEffectiveEmployeeId();
    if (!targetEmployee) {
      isFetching = false;
      updateBookedLaborUI('Set login in MyAPM');
      return;
    }

    const url = 'https://us1.eam.hxgnsmartcloud.com/web/base/WSBOOK.HDR.xmlhttp';
    const currentTenant = PAGE_WINDOW.EAM?.AppData?.tenant || 'AMAZONRMENA_PRD';

    const payload = new URLSearchParams({
      GRID_ID: '1742',
      GRID_NAME: 'WSBOOK_HDR',
      DATASPY_ID: '100696',
      USER_FUNCTION_NAME: 'WSBOOK',
      SYSTEM_FUNCTION_NAME: 'WSBOOK',
      CURRENT_TAB_NAME: 'HDR',
      COMPONENT_INFO_TYPE: 'DATA_ONLY',
      employee: targetEmployee,
      tenant: currentTenant,
      eamid: currentEamId,
      NUMBER_OF_ROWS_FIRST_RETURNED: '5000'
    });

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
          'X-Requested-With': 'XMLHttpRequest'
        },
        body: payload.toString()
      });

      const text = await response.text();
      if (text.trim().startsWith('<') || text.includes('System Error')) {
        updateBookedLaborUI('Server Rejected');
        isFetching = false;
        return;
      }

      const jsonStart = text.indexOf('{');
      if (jsonStart === -1) throw new Error('No JSON found');

      const dataObj = JSON.parse(text.substring(jsonStart));
      laborCache.data = dataObj?.pageData?.grid?.GRIDRESULT?.GRID?.DATA || [];
      laborCache.lastFetch = Date.now();
    } catch (_) {
      updateBookedLaborUI('Data Error');
    } finally {
      isFetching = false;
      updateBookedLaborUI();
    }
  }

  function calculateLabor(daysParam) {
    const now = new Date();
    now.setHours(0, 0, 0, 0);
    let total = 0;
    const breakdown = {};

    laborCache.data.forEach((r) => {
      const rDate = new Date(r.datework);
      rDate.setHours(0, 0, 0, 0);
      const diffDays = Math.floor((now - rDate) / (1000 * 3600 * 24));
      const maxDaysAgo = daysParam === 7 ? 7 : daysParam - 1;

      if (diffDays <= maxDaysAgo && diffDays >= 0) {
        const hrs = parseFloat(r.hrswork);
        if (!Number.isNaN(hrs)) {
          total += hrs;
          breakdown[r.datework] = (breakdown[r.datework] || 0) + hrs;
        }
      }
    });

    return { total, breakdown };
  }

  function bindBookedLaborCard(card) {
    if (!card || card.dataset.myApmBound === 'true') return;
    card.dataset.myApmBound = 'true';

    card.querySelectorAll('.my-labor-tab').forEach((tabEl) => {
      tabEl.addEventListener('click', (e) => {
        card.querySelectorAll('.my-labor-tab').forEach((x) => x.classList.remove('active'));
        e.currentTarget.classList.add('active');
        activeTab = parseInt(e.currentTarget.getAttribute('data-d'), 10) || 1;
        if (!isFetching) updateBookedLaborUI();
      });
    });

    const refreshBtn = card.querySelector(`#${IDS.refresh}`);
    if (refreshBtn) {
      refreshBtn.addEventListener('click', () => {
        laborCache.lastFetch = 0;
        fetchLaborData(true);
      });
    }

    const searchInput = card.querySelector(`#${IDS.searchInput}`);
    if (searchInput) {
      searchInput.addEventListener('keydown', (e) => {
        if (e.key !== 'Enter') return;
        e.preventDefault();
        laborTargetOverride = String(searchInput.value || '').trim().toUpperCase();
        laborCache.lastFetch = 0;
        fetchLaborData(true);
      });
    }
  }

  function updateBookedLaborUI(errorMsg = null) {
    const root = document.getElementById(IDS.root);
    if (!root) return;
    const sumBox = root.querySelector(`#${IDS.sum}`);
    const list = root.querySelector(`#${IDS.list}`);
    const searchStatus = root.querySelector(`#${IDS.searchStatus}`);
    const searchInput = root.querySelector(`#${IDS.searchInput}`);
    if (!sumBox || !list) return;

    const activeEmployee = getEffectiveEmployeeId();
    const defaultEmployee = getDefaultLaborEmployeeId();
    if (searchStatus) {
      const showingSelf = !String(laborTargetOverride || '').trim();
      searchStatus.textContent = showingSelf
        ? `Showing: ${defaultEmployee || 'Self'} (default)`
        : `Showing: ${activeEmployee}`;
    }
    if (searchInput && document.activeElement !== searchInput) {
      searchInput.value = String(laborTargetOverride || '');
    }

    if (errorMsg) {
      let color = '#ff6b6b';
      if (errorMsg === 'Loading...') color = '#ffb347';
      if (errorMsg === 'Set login in MyAPM') color = '#ffd18a';
      sumBox.innerHTML = `<span style="font-size:15px; color:${color};">${errorMsg}</span>`;
      list.innerHTML = '';
      return;
    }

    const { total, breakdown } = calculateLabor(activeTab);
    sumBox.innerHTML = `${total.toFixed(2)} <span style="font-size:14px; color:#9fb0c4;">hrs</span>`;

    list.innerHTML = '';
    const sortedDates = Object.keys(breakdown).sort((a, b) => new Date(b) - new Date(a));
    if (sortedDates.length === 0) {
      list.innerHTML = '<div style="text-align:center; padding:10px; color:#9fb0c4; font-size:12px;">No labor records found.</div>';
      return;
    }

    sortedDates.forEach((d) => {
      const row = document.createElement('div');
      row.className = 'my-labor-row';
      row.innerHTML = `<span>${d}</span><strong>${breakdown[d].toFixed(2)}</strong>`;
      list.appendChild(row);
    });
  }

  window.createBookedLaborSettingsCard = function createBookedLaborSettingsCard() {
    ensureBookedLaborStyles();

    const card = document.createElement('section');
    card.id = IDS.root;
    Object.assign(card.style, {
      background: 'linear-gradient(180deg, rgba(27, 39, 58, 0.94), rgba(16, 24, 38, 0.94))',
      border: '1px solid rgba(69, 95, 130, 0.9)',
      borderRadius: '16px',
      padding: '16px',
      boxShadow: 'inset 0 1px 0 rgba(255,255,255,0.04)',
      minWidth: '0'
    });

    card.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; gap:8px;">
        <span style="font-size:13px; color:#eaf0ff; font-weight:700;">Booked Labor</span>
      </div>
      <div class="my-labor-tabs">
        <div class="my-labor-tab active" data-d="1">Today</div>
        <div class="my-labor-tab" data-d="2">2-Day</div>
        <div class="my-labor-tab" data-d="7">7-Day</div>
      </div>
      <input id="${IDS.searchInput}" type="text" placeholder="Search username and press Enter" style="width:100%; box-sizing:border-box; margin-bottom:6px; padding:7px 8px; border:1px solid #3c516d; border-radius:4px; background:#0f1723; color:#eaf0ff; font-size:12px;" />
      <div id="${IDS.searchStatus}" style="font-size:11px; color:#9fb0c4; margin-bottom:8px;"></div>
      <div id="${IDS.sum}" class="my-labor-total">0.00 <span style="font-size:14px; color:#9fb0c4;">hrs</span></div>
      <div id="${IDS.list}" style="max-height:220px; overflow-y:auto;"></div>
      <button id="${IDS.refresh}" style="margin-top:14px; border:1px solid #3c516d; background:linear-gradient(to bottom, #2a3a50, #1b2738); color:#eaf0ff; padding:8px; border-radius:4px; cursor:pointer; font-size:11px; font-weight:700;">Refresh from Server</button>
    `;

    bindBookedLaborCard(card);
    setTimeout(() => {
      updateBookedLaborUI();
      fetchLaborData(false);
    }, 0);
    return card;
  };
})();


/* --------------------------------------------------
   PTP TIMER + CLOSING COMMENTS COUNTER PORT
--------------------------------------------------*/
(function() {
  const PTP_MSG_TAG = '__myApmPtpMsgV1';
  const PTP_PARENT_TIMER_ID = 'ptp-timer';
  const PTP_PARENT_STYLE_ID = 'myapm-ptp-pulse-style';
  const PTP_COUNTDOWN_SECONDS = 120;
  const PTP_SHARED_HISTORY_KEY = 'myapm_shared_ptp_history_v1';
  const PTP_LOCAL_HISTORY_KEY = 'apm_ptp_history';

  function ptpEnabled() {
    return localStorage.getItem('apmPTPTimer') !== 'false';
  }

  let _ptpHistoryCleaned = false;

  function readSharedValue(key, fallbackValue) {
    try {
      if (typeof GM_getValue === 'function') {
        return GM_getValue(key, fallbackValue);
      }
    } catch (_) {}
    try {
      const localValue = localStorage.getItem(key);
      return localValue === null ? fallbackValue : localValue;
    } catch (_) {
      return fallbackValue;
    }
  }

  function writeSharedValue(key, value) {
    try {
      if (typeof GM_setValue === 'function') {
        GM_setValue(key, value);
      }
    } catch (_) {}
    try {
      localStorage.setItem(key, value);
    } catch (_) {}
  }

  function parsePtpHistoryValue(raw) {
    if (!raw) return {};
    if (typeof raw === 'object') return raw || {};
    try {
      return JSON.parse(String(raw)) || {};
    } catch (_) {
      return {};
    }
  }

  function dispatchPtpHistoryUpdated(detail) {
    try { window.dispatchEvent(new CustomEvent('MYAPM_PTP_HISTORY_UPDATED', { detail })); } catch (_) {}
    try {
      if (window.top && window.top !== window) {
        window.top.dispatchEvent(new CustomEvent('MYAPM_PTP_HISTORY_UPDATED', { detail }));
      }
    } catch (_) {}
  }

  function getPtpHistory() {
    let history = parsePtpHistoryValue(readSharedValue(PTP_SHARED_HISTORY_KEY, ''));
    if (!history || !Object.keys(history).length) {
      history = parsePtpHistoryValue(readSharedValue(PTP_LOCAL_HISTORY_KEY, '{}'));
    }
    if (!_ptpHistoryCleaned) {
      _ptpHistoryCleaned = true;
      const now = Date.now();
      const maxAgeMs = 7 * 24 * 60 * 60 * 1000;
      let changed = false;
      Object.keys(history).forEach((key) => {
        const entry = history[key];
        const timestamp = typeof entry === 'object' ? Number(entry.time || 0) : Number(entry || 0);
        if (!timestamp || (now - timestamp) > maxAgeMs) {
          delete history[key];
          changed = true;
        } else if (typeof entry !== 'object') {
          history[key] = { status: 'COMPLETE', time: timestamp };
          changed = true;
        }
      });
      if (changed) {
        const serialized = JSON.stringify(history);
        writeSharedValue(PTP_SHARED_HISTORY_KEY, serialized);
        try { localStorage.setItem(PTP_LOCAL_HISTORY_KEY, serialized); } catch (_) {}
      }
    }
    try { localStorage.setItem(PTP_LOCAL_HISTORY_KEY, JSON.stringify(history)); } catch (_) {}
    return history;
  }

  function updatePtpHistory(woNumber, status) {
    const wo = String(woNumber || '').trim();
    if (!wo) return;
    const history = getPtpHistory();
    history[wo] = { status: String(status || 'COMPLETE').toUpperCase(), time: Date.now() };
    const serialized = JSON.stringify(history);
    writeSharedValue(PTP_SHARED_HISTORY_KEY, serialized);
    try { localStorage.setItem(PTP_LOCAL_HISTORY_KEY, serialized); } catch (_) {}
    dispatchPtpHistoryUpdated({ wo, data: history[wo] });
  }

  window.__myApmGetPtpHistory = getPtpHistory;

  function isParentApmHost() {
    return location.hostname.endsWith('.hxgnsmartcloud.com');
  }

  function isPtpIframeHost() {
    return location.hostname.endsWith('.apm-es.gps.amazon.dev') || location.hostname.endsWith('.insights.amazon.dev');
  }

  function relayPtpMessage(type, extra) {
    const payload = Object.assign({ type }, extra || {});
    const targets = [];
    try { if (window.top && window.top !== window) targets.push(window.top); } catch (_) {}
    try { if (window.parent && window.parent !== window.top && window.parent !== window) targets.push(window.parent); } catch (_) {}
    try { if (window.opener && window.opener !== window) targets.push(window.opener); } catch (_) {}
    const posted = new Set();
    targets.forEach((target) => {
      if (!target || posted.has(target)) return;
      posted.add(target);
      try { target.postMessage(payload, '*'); } catch (_) {}
    });
  }

  function createStandalonePtpTimer() {
    if (!isPtpIframeHost()) return { start() {}, reset() {} };
    let isTopLevel = false;
    try {
      isTopLevel = window.top === window.self;
    } catch (_) {
      isTopLevel = true;
    }
    if (!isTopLevel) return { start() {}, reset() {} };

    const BOX_ID = 'myapm-ptp-page-timer';
    const TEXT_ID = 'myapm-ptp-page-timer-text';
    const BAR_ID = 'myapm-ptp-page-timer-progress';
    let box = null;
    let secs = PTP_COUNTDOWN_SECONDS;
    let tick = null;

    function clearTick() {
      if (tick) clearInterval(tick);
      tick = null;
    }

    function removeBox() {
      const node = box || document.getElementById(BOX_ID);
      box = null;
      if (node) node.remove();
    }

    function ensureBox() {
      box = document.getElementById(BOX_ID);
      if (box) return box;
      box = document.createElement('div');
      box.id = BOX_ID;
      Object.assign(box.style, {
        position: 'fixed',
        top: '72px',
        right: '20px',
        width: '196px',
        minHeight: '56px',
        padding: '8px 14px 12px',
        zIndex: '2147483647',
        background: 'linear-gradient(to bottom,#146eb4,#0d4f8b)',
        color: '#fff',
        font: '600 16px/20px "Segoe UI",sans-serif',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        borderRadius: '10px',
        boxShadow: '0 8px 20px rgba(0,0,0,.28)',
        border: '1px solid #146eb4'
      });

      const txt = document.createElement('span');
      txt.id = TEXT_ID;
      const bar = document.createElement('div');
      bar.id = BAR_ID;
      Object.assign(bar.style, {
        width: '100%',
        height: '4px',
        marginTop: '8px',
        borderRadius: '999px',
        background: '#2196F3',
        transition: 'width 1s linear,background-color .3s'
      });

      const close = document.createElement('div');
      close.innerHTML = '&times;';
      Object.assign(close.style, {
        position: 'absolute',
        top: '4px',
        right: '8px',
        width: '20px',
        height: '20px',
        fontSize: '16px',
        fontWeight: 'bold',
        lineHeight: '14px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        borderRadius: '50%',
        color: '#ffffffcc',
        cursor: 'pointer'
      });
      close.onclick = () => reset();

      box.append(txt, bar, close);
      document.body.appendChild(box);
      return box;
    }

    function draw() {
      const txt = document.getElementById(TEXT_ID);
      const bar = document.getElementById(BAR_ID);
      if (!txt || !bar) return;
      txt.textContent = `PTP Timer: ${Math.floor(secs / 60)}:${String(secs % 60).padStart(2, '0')}`;
      const pct = Math.max(0, (secs / PTP_COUNTDOWN_SECONDS) * 100);
      bar.style.width = `${pct}%`;
    }

    function finish() {
      const txt = document.getElementById(TEXT_ID);
      const bar = document.getElementById(BAR_ID);
      if (!box || !txt || !bar) return;
      txt.textContent = 'READY';
      bar.style.width = '100%';
      bar.style.background = '#2ecc71';
      box.style.background = 'linear-gradient(135deg,#4CAF50 0%,#81C784 100%)';
      box.style.color = '#1B5E20';
      box.style.borderColor = '#2ecc71';
    }

    function start() {
      if (!ptpEnabled()) return;
      clearTick();
      ensureBox();
      secs = PTP_COUNTDOWN_SECONDS;
      draw();
      tick = setInterval(() => {
        secs -= 1;
        if (secs <= 0) {
          clearTick();
          secs = 0;
          finish();
          return;
        }
        draw();
      }, 1000);
    }

    function reset() {
      clearTick();
      removeBox();
    }

    return { start, reset };
  }

  function installParentPtpTimer() {
    if (!isParentApmHost() || window.__myApmPtpParentInstalled) return;
    window.__myApmPtpParentInstalled = true;

    let box = null;
    let secs = PTP_COUNTDOWN_SECONDS;
    let tick = null;
    let statusWatch = null;

    const state = window.__myApmPtpParentState = window.__myApmPtpParentState || {
      running: false,
      dismissed: false,
      lastHeartbeat: 0
    };

    const isAllowedPtpIframeOrigin = (origin) => {
      try {
        const host = String(new URL(origin).hostname || '').toLowerCase();
        return host.endsWith('.apm-es.gps.amazon.dev') || host.endsWith('.insights.amazon.dev');
      } catch (_) {
        return false;
      }
    };

    function logParent() {}

    function ensurePulseStyle() {
      if (document.getElementById(PTP_PARENT_STYLE_ID)) return;
      const style = document.createElement('style');
      style.id = PTP_PARENT_STYLE_ID;
      style.textContent = '@keyframes ptpPulse { from { transform: translateY(0) scale(1); } to { transform: translateY(0) scale(1.02); } }';
      document.head.appendChild(style);
    }

    function clearTimers() {
      if (tick) clearInterval(tick);
      if (statusWatch) clearInterval(statusWatch);
      tick = null;
      statusWatch = null;
    }

    function removeBox() {
      const node = box || document.getElementById(PTP_PARENT_TIMER_ID);
      box = null;
      if (node) node.remove();
    }

    function finishVisual() {
      if (!box) return;
      const txt = document.getElementById('ptp-timer-text');
      const bar = document.getElementById('ptp-progress');
      if (txt) txt.textContent = 'READY';
      if (bar) {
        bar.style.width = '100%';
        bar.style.background = '#2ecc71';
      }
      box.style.background = 'linear-gradient(135deg,#4CAF50 0%,#81C784 100%)';
      box.style.color = '#1B5E20';
      box.style.borderColor = '#2ecc71';
      box.style.animation = 'ptpPulse 1.5s ease-in-out infinite alternate';
    }

    function draw() {
      const txt = document.getElementById('ptp-timer-text');
      const bar = document.getElementById('ptp-progress');
      if (!txt || !bar) return;
      txt.textContent = `PTP Timer: ${Math.floor(secs / 60)}:${String(secs % 60).padStart(2, '0')}`;
      const pct = Math.max(0, (secs / PTP_COUNTDOWN_SECONDS) * 100);
      bar.style.width = `${pct}%`;
    }

    function hide(markDismissed = true) {
      clearTimers();
      state.running = false;
      if (markDismissed) state.dismissed = true;
      if (!box) {
        removeBox();
        return;
      }
      const node = box;
      box = null;
      node.style.opacity = '0';
      node.style.transform = 'translateY(-10px)';
      setTimeout(() => node.remove(), 220);
    }

    function resetTimer(markDismissed = false) {
      hide(markDismissed);
    }

    function createBox() {
      removeBox();
      ensurePulseStyle();
      box = document.createElement('div');
      box.id = PTP_PARENT_TIMER_ID;
      Object.assign(box.style, {
        position: 'fixed',
        top: '300px',
        right: '100px',
        width: '180px',
        minHeight: '50px',
        padding: '6px 14px 10px',
        zIndex: '10000',
        background: 'linear-gradient(to bottom,#146eb4,#0d4f8b)',
        color: '#fff',
        font: '600 16px/20px "Segoe UI",sans-serif',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        borderRadius: '8px',
        boxShadow: '0 4px 8px rgba(0,0,0,.2)',
        border: '1px solid #146eb4',
        opacity: '0',
        transform: 'translateY(-10px)',
        transition: 'all .25s ease, transform .25s ease'
      });

      const txt = document.createElement('span');
      txt.id = 'ptp-timer-text';
      const bar = document.createElement('div');
      bar.id = 'ptp-progress';
      Object.assign(bar.style, {
        width: '100%',
        height: '4px',
        marginTop: '6px',
        borderRadius: '2px',
        background: '#2196F3',
        transition: 'width 1s linear,background-color .3s'
      });

      const close = document.createElement('div');
      close.innerHTML = '&times;';
      Object.assign(close.style, {
        position: 'absolute',
        top: '4px',
        right: '8px',
        width: '20px',
        height: '20px',
        fontSize: '16px',
        fontWeight: 'bold',
        lineHeight: '14px',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        borderRadius: '50%',
        color: '#ffffffcc',
        cursor: 'pointer'
      });
      close.onclick = () => hide(true);

      box.append(txt, bar, close);
      document.body.appendChild(box);
      requestAnimationFrame(() => {
        if (!box) return;
        box.style.opacity = '1';
        box.style.transform = 'translateY(0)';
      });
    }

    function start() {
      logParent('start', {
        enabled: ptpEnabled(),
        dismissed: state.dismissed,
        href: location.href,
        lastHeartbeat: state.lastHeartbeat
      });
      if (!ptpEnabled() || state.dismissed) return;
      clearTimers();
      createBox();
      state.running = true;
      secs = PTP_COUNTDOWN_SECONDS;
      draw();
      tick = setInterval(() => {
        secs -= 1;
        if (secs <= 0) {
          clearInterval(tick);
          tick = null;
          secs = 0;
          finishVisual();
          state.running = false;
          return;
        }
        draw();
      }, 1000);
    }

    function checkPtpStatus(hasHeartbeat = false) {
      logParent('checkPtpStatus', {
        hasHeartbeat,
        enabled: ptpEnabled(),
        lastHeartbeat: state.lastHeartbeat,
        running: state.running,
        dismissed: state.dismissed
      });
      if (!ptpEnabled()) {
        resetTimer(false);
        return;
      }
      if (hasHeartbeat) {
        state.lastHeartbeat = Date.now();
      }
    }

    statusWatch = setInterval(() => checkPtpStatus(false), 3000);

    window.addEventListener('message', (e) => {
      const data = e && e.data ? e.data : null;
      logParent('message received', {
        origin: e && e.origin ? e.origin : '',
        type: data && data.type ? data.type : '',
        data,
        isAllowedPtpIframeOrigin: isAllowedPtpIframeOrigin(e && e.origin ? e.origin : '')
      });
      if (!data) return;

      if (data.type === 'MYAPM_PTP_SETTINGS_CHANGED' && e.origin === location.origin) {
        logParent('message settings changed', { enabled: data.enabled, origin: e.origin });
        if (data.enabled === false) {
          resetTimer(false);
        } else if (data.enabled === true) {
          state.dismissed = false;
          checkPtpStatus(false);
        }
        return;
      }

      if (data.type === 'MYAPM_PTP_HEARTBEAT' && isAllowedPtpIframeOrigin(e.origin)) {
        logParent('message heartbeat accepted', { origin: e.origin, url: data.url || '', visible: data.visible });
        state.lastHeartbeat = Date.now();
        checkPtpStatus(true);
        return;
      }

      if (data.type === 'MYAPM_PTP_START' && isAllowedPtpIframeOrigin(e.origin)) {
        logParent('message start accepted', { origin: e.origin, data });
        state.dismissed = false;
        state.lastHeartbeat = Date.now();
        start();
        return;
      }

      if ((data.type === 'MYAPM_PTP_CANCELLED' || data.type === 'MYAPM_PTP_COMPLETED' || data.type === 'MYAPM_PTP_INCOMPLETE') && isAllowedPtpIframeOrigin(e.origin)) {
        logParent('message terminal accepted', {
          origin: e.origin,
          type: data.type,
          wo: data.wo || '',
          data
        });
        if (data.type === 'MYAPM_PTP_COMPLETED' && data.wo) updatePtpHistory(data.wo, 'COMPLETE');
        if (data.type === 'MYAPM_PTP_CANCELLED' && data.wo) updatePtpHistory(data.wo, 'CANCELLED');
        if (data.type === 'MYAPM_PTP_INCOMPLETE' && data.wo) updatePtpHistory(data.wo, 'INCOMPLETE');
        if (data.type !== 'MYAPM_PTP_INCOMPLETE') resetTimer(false);
        return;
      }

      if (data.type && /^MYAPM_PTP_/.test(String(data.type))) {
        logParent('message rejected', {
          origin: e.origin,
          type: data.type,
          allowedOrigin: isAllowedPtpIframeOrigin(e.origin),
          locationOrigin: location.origin
        });
      }

      if (data[PTP_MSG_TAG] !== true) return;
      const action = String(data.ptpTimer || '').toLowerCase();
      const source = String(data.source || '').toLowerCase();
      const sameOriginSettingsMessage = source === 'settings' && e.origin === location.origin;
      const trustedIframeMessage = source === 'iframe-ptp' && isAllowedPtpIframeOrigin(e.origin);
      logParent('legacy message path', {
        action,
        source,
        sameOriginSettingsMessage,
        trustedIframeMessage,
        origin: e.origin
      });
      if (action === 'start') {
        state.dismissed = false;
        if (trustedIframeMessage) {
          state.lastHeartbeat = Date.now();
          start();
        }
        if (sameOriginSettingsMessage && state.lastHeartbeat) {
          start();
        }
      }
      if (action === 'reset' && (trustedIframeMessage || sameOriginSettingsMessage)) resetTimer(false);
    });

    window.ptpStart = start;
    window.ptpHide = hide;
  }

  function installIframePtpBridge() {
    if (!isPtpIframeHost() || window.__myApmPtpIframeBridgeBound) return;

    let completionFired = false;
    let currentPtpWo = '';
    const standaloneTimer = createStandalonePtpTimer();
    function logBridge() {}

    function post(type, extra) {
      logBridge('postMessage', { type, extra: extra || null, href: location.href, currentPtpWo, completionFired });
      relayPtpMessage(type, extra);
    }

    function triggerStart() {
      logBridge('triggerStart', { href: location.href, currentPtpWo, completionFired });
      standaloneTimer.start();
      post('MYAPM_PTP_START');
      post('MYAPM_PTP_HEARTBEAT', { visible: true, url: location.href });
    }

    function triggerCancel(woNumber) {
      const payload = woNumber ? { wo: String(woNumber) } : {};
      logBridge('triggerCancel', { woNumber: woNumber || '', payload, currentPtpWo, completionFired });
      if (payload.wo) updatePtpHistory(payload.wo, 'CANCELLED');
      standaloneTimer.reset();
      post('MYAPM_PTP_HEARTBEAT', { visible: false, url: location.href });
    }

    function triggerCompletion(woNumber) {
      if (completionFired) {
        logBridge('triggerCompletion skipped', { reason: 'completion already fired', woNumber: woNumber || '', currentPtpWo });
        return;
      }
      logBridge('triggerCompletion', { woNumber: woNumber || '', currentPtpWo, completionFired });
      completionFired = true;
      if (woNumber) updatePtpHistory(woNumber, 'COMPLETE');
      standaloneTimer.reset();
    }

    function triggerIncomplete(woNumber) {
      logBridge('triggerIncomplete', { woNumber: woNumber || '', currentPtpWo, completionFired });
      if (woNumber) updatePtpHistory(woNumber, 'INCOMPLETE');
    }

    function resolveWoNumber(url, requestBody, responseObj) {
      const requestUrl = String(url || '');
      const locationUrl = String(location.href || '');
      const urlMatch = requestUrl.match(/workordernum=(\d{6,})/i);
      const locationMatch = locationUrl.match(/workordernum=(\d{6,})/i);
      const resolved = urlMatch && urlMatch[1]
        ? String(urlMatch[1])
        : (locationMatch && locationMatch[1] ? String(locationMatch[1]) : '');
      logBridge('resolveWoNumber', {
        requestUrl,
        locationUrl,
        resolved,
        hadRequestBody: typeof requestBody !== 'undefined' && requestBody !== null,
        responseKeys: responseObj && typeof responseObj === 'object' ? Object.keys(responseObj).slice(0, 8) : []
      });
      return resolved;
    }

    function handleAssessmentResponse(url, text, status, requestBody) {
      try {
        const previousWo = currentPtpWo;
        const resolvedFromUrl = resolveWoNumber(url, requestBody, null) || '';
        currentPtpWo = resolvedFromUrl || currentPtpWo;
        logBridge('handleAssessmentResponse:start', {
          url,
          status,
          previousWo,
          resolvedFromUrl,
          currentPtpWo,
          completionFired,
          textLength: String(text || '').length,
          bodyPreview: typeof requestBody === 'string' ? String(requestBody).slice(0, 300) : requestBody
        });
        if (url.includes('submit_assessment') && status >= 200 && status < 300) {
          logBridge('handleAssessmentResponse:submit_assessment success', {
            url,
            status,
            currentPtpWo,
            willTriggerCompletion: !!currentPtpWo
          });
          if (currentPtpWo) triggerCompletion(currentPtpWo);
          else logBridge('handleAssessmentResponse:submit_assessment missing WO', { url, status });
          return;
        }
        if (url.includes('create_assessment') && status === 200) {
          completionFired = false;
          currentPtpWo = resolveWoNumber(url, requestBody, null) || currentPtpWo;
          logBridge('handleAssessmentResponse:create_assessment success', {
            url,
            status,
            currentPtpWo,
            completionFired
          });
          if (currentPtpWo) triggerIncomplete(currentPtpWo);
          triggerStart();
          return;
        }
        if (!text || !text.includes('100')) {
          logBridge('handleAssessmentResponse:ignored non-json-ish response', {
            url,
            status,
            hasText: !!text,
            textPreview: String(text || '').slice(0, 300)
          });
          return;
        }
        const res = JSON.parse(text);
        currentPtpWo = resolveWoNumber(url, requestBody, res) || currentPtpWo;
        logBridge('handleAssessmentResponse:parsed', {
          url,
          status,
          currentPtpWo,
          completionFired,
          hasBody: !!(res && res.body),
          assessmentStatus: res && res.body && res.body.assessment ? String(res.body.assessment.AssessmentStatus || '') : '',
          finalStatus: res && res.body && res.body.response ? String(res.body.response.final_status || '') : '',
          responseWorkOrderId: res && res.body && res.body.response ? String(res.body.response.workorder_id || '') : ''
        });

        if (res && res.body && res.body.assessment && res.body.assessment.AssessmentStatus === 'INCOMPLETE') {
          logBridge('handleAssessmentResponse:assessment incomplete', { url, currentPtpWo });
          completionFired = false;
          if (currentPtpWo) triggerIncomplete(currentPtpWo);
          triggerStart();
        } else if (res && res.body && res.body.response && res.body.response.workorder_id) {
          const finalStatus = String(res.body.response.final_status || '');
          logBridge('handleAssessmentResponse:response workorder_id branch', {
            url,
            finalStatus,
            responseWorkOrderId: String(res.body.response.workorder_id || ''),
            currentPtpWo
          });
          if (finalStatus === 'COMPLETE') {
            triggerCompletion(res.body.response.workorder_id);
          } else if (finalStatus === 'CANCELLED') {
            triggerCancel(currentPtpWo);
          } else {
            logBridge('handleAssessmentResponse:response branch no terminal status', { url, finalStatus, currentPtpWo });
          }
        } else if (res && res.body && res.body.assessment && res.body.assessment.AssessmentStatus) {
          const assessmentStatus = String(res.body.assessment.AssessmentStatus || '');
          logBridge('handleAssessmentResponse:assessment status branch', {
            url,
            assessmentStatus,
            currentPtpWo
          });
          if (assessmentStatus === 'COMPLETE') {
            const woNumber = resolveWoNumber(url, requestBody, res) || currentPtpWo;
            if (woNumber) triggerCompletion(woNumber);
            else logBridge('handleAssessmentResponse:assessment complete without WO', { url, assessmentStatus });
          } else if (assessmentStatus === 'CANCELLED') {
            triggerCancel(currentPtpWo);
          } else {
            logBridge('handleAssessmentResponse:assessment status non-terminal', { url, assessmentStatus });
          }
        } else if (url.includes('get_revisions') && res && res.body && Array.isArray(res.body.revisions)) {
          const inactiveRevision = res.body.revisions.find((r) => r && r.status === 'inactive');
          logBridge('handleAssessmentResponse:get_revisions branch', {
            url,
            revisions: res.body.revisions.length,
            foundInactiveRevision: !!inactiveRevision,
            currentPtpWo
          });
          if (inactiveRevision) triggerCancel(currentPtpWo);
        } else if (url.includes('submit_assessment')) {
          logBridge('handleAssessmentResponse:submit_assessment text fallback branch', {
            url,
            currentPtpWo,
            hasCompleteToken: String(text || '').includes('OMPLETE'),
            hasCancelledToken: String(text || '').includes('CANCELLED')
          });
          if (text.includes('OMPLETE')) {
            const woNumber = resolveWoNumber(url, requestBody, res) || currentPtpWo;
            if (woNumber) triggerCompletion(woNumber);
            else logBridge('handleAssessmentResponse:submit fallback complete without WO', { url });
          } else if (text.includes('CANCELLED')) {
            triggerCancel(currentPtpWo);
          } else {
            logBridge('handleAssessmentResponse:submit fallback no terminal token', { url });
          }
        } else {
          logBridge('handleAssessmentResponse:no matching branch', { url, status, currentPtpWo });
        }
      } catch (error) {
        logBridge('handleAssessmentResponse:error', {
          url,
          status,
          message: error && error.message ? error.message : String(error),
          textPreview: String(text || '').slice(0, 300)
        });
      }
    }

    const origOpen = XMLHttpRequest.prototype.open;
    const origSend = XMLHttpRequest.prototype.send;
    XMLHttpRequest.prototype.open = function(method, url) {
      this._myApmPtpUrl = (url || '').toString();
      logBridge('xhr.open', { method: method || '', url: this._myApmPtpUrl || '' });
      return origOpen.apply(this, arguments);
    };
    XMLHttpRequest.prototype.send = function(body) {
      if (this._myApmPtpUrl && /(submit_assessment|get_assessment|create_assessment)/i.test(this._myApmPtpUrl)) {
        logBridge('xhr.send matched', {
          url: this._myApmPtpUrl,
          bodyPreview: typeof body === 'string' ? String(body).slice(0, 300) : body
        });
        this.addEventListener('load', function() {
          logBridge('xhr.load matched', {
            url: this._myApmPtpUrl,
            status: this.status,
            responsePreview: String(this.responseText || '').slice(0, 300)
          });
          handleAssessmentResponse(this._myApmPtpUrl, this.responseText, this.status, body);
        });
      }
      return origSend.apply(this, arguments);
    };

    const origFetch = window.fetch;
    if (typeof origFetch === 'function') {
      window.fetch = async function(...args) {
        const response = await origFetch.apply(this, args);
        try {
          const url = args[0] instanceof Request ? args[0].url : typeof args[0] === 'string' ? args[0] : '';
          logBridge('fetch observed', {
            url,
            status: response && typeof response.status !== 'undefined' ? response.status : null
          });
          if (url && /(submit_assessment|get_assessment|create_assessment|get_revisions)/i.test(url)) {
            const clone = response.clone();
            const text = await clone.text();
            const reqObj = args[0] instanceof Request ? args[0] : (args[1] || {});
            logBridge('fetch matched', {
              url,
              status: response.status,
              bodyPreview: typeof reqObj.body === 'string' ? String(reqObj.body).slice(0, 300) : reqObj.body,
              responsePreview: String(text || '').slice(0, 300)
            });
            handleAssessmentResponse(url, text, response.status, reqObj.body);
          }
        } catch (error) {
          logBridge('fetch hook error', {
            message: error && error.message ? error.message : String(error)
          });
        }
        return response;
      };
    }

    function heartbeat() {
      const hasPtpHeader = !!document.querySelector('.ptp-header, .permit-details, #ptp-main-content, [class*="awsui_root_"]');
      const hasWorkOrder = /workorder/i.test(location.href);
      if (hasPtpHeader || hasWorkOrder) {
        logBridge('heartbeat', { hasPtpHeader, hasWorkOrder, href: location.href, currentPtpWo, completionFired });
        post('MYAPM_PTP_HEARTBEAT', { visible: true, url: location.href });
      }
    }

    document.addEventListener('click', (ev) => {
      const path = typeof ev.composedPath === 'function' ? ev.composedPath() : [];
      const btn = path.find((n) => ['MWC-BUTTON', 'BUTTON'].includes(n && n.tagName) && ['Create New Assessment', 'Create Assessment'].includes((n.textContent || '').trim()));
      if (btn) {
        logBridge('click matched create assessment', {
          tagName: btn && btn.tagName ? btn.tagName : '',
          text: btn && btn.textContent ? String(btn.textContent).trim() : ''
        });
        completionFired = false;
        triggerStart();
      }
    }, true);

    setInterval(heartbeat, 8000);
    currentPtpWo = resolveWoNumber(location.href, null, null) || currentPtpWo;
    logBridge('bridge initialized', { href: location.href, currentPtpWo, completionFired });
    heartbeat();

    window.__myApmPtpIframeBridgeBound = true;
  }

  function waitForCondition(checkFn, options = {}) {
    const timeoutMs = Number(options.timeoutMs || 8000);
    const intervalMs = Number(options.intervalMs || 120);
    const observeMutations = options.observeMutations === true;

    return new Promise((resolve) => {
      let settled = false;
      let intervalId = null;
      let timeoutId = null;
      let observer = null;

      const cleanup = () => {
        if (intervalId) clearInterval(intervalId);
        if (timeoutId) clearTimeout(timeoutId);
        if (observer) observer.disconnect();
      };

      const finish = (value) => {
        if (settled) return;
        settled = true;
        cleanup();
        resolve(value || null);
      };

      const probe = () => {
        let result = null;
        try { result = checkFn(); } catch (_) {}
        if (result) finish(result);
      };

      timeoutId = setTimeout(() => finish(null), timeoutMs);
      intervalId = setInterval(probe, intervalMs);

      if (observeMutations && document.body) {
        observer = new MutationObserver(probe);
        observer.observe(document.body, { childList: true, subtree: true });
      }

      probe();
    });
  }

  function fireTextInput(el, value) {
    if (!el) return false;
    el.focus();
    el.value = value;
    ['input', 'change', 'keyup', 'blur'].forEach((evtName) => {
      try {
        el.dispatchEvent(new Event(evtName, { bubbles: true }));
      } catch (_) {}
    });
    return true;
  }



  function clickElementSafe(el) {
    if (!el) return false;
    try { el.click(); return true; } catch (_) {}
    try {
      el.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
      return true;
    } catch (_) {}
    return false;
  }

  function openRecordTabByText(tabText) {
    const wanted = String(tabText || '').trim().toLowerCase();
    if (!wanted) return false;

    const domTab = Array.from(document.querySelectorAll('.x-tab .x-tab-inner, span.x-tab-inner, a.x-tab'))
      .find((el) => isElementVisible(el) && String(el.textContent || '').trim().toLowerCase() === wanted);
    if (domTab && clickElementSafe(domTab.closest('a.x-tab') || domTab)) return true;

    const topExt = (window.top && window.top.Ext) ? window.top.Ext : null;
    if (topExt && topExt.ComponentQuery) {
      try {
        const panels = topExt.ComponentQuery.query('tabpanel');
        for (const panel of panels) {
          if (!panel || panel.isDestroyed || !panel.items || !panel.items.items) continue;
          const match = panel.items.items.find((item) => {
            const title = String((item && (item.title || item.text)) || '').trim().toLowerCase();
            return title === wanted;
          });
          if (match) {
            if (typeof panel.setActiveTab === 'function') panel.setActiveTab(match);
            else if (match.tab && match.tab.el && match.tab.el.dom) match.tab.el.dom.click();
            return true;
          }
        }
      } catch (_) {}
    }

    return false;
  }

  function isExtButtonEnabled(btn) {
    if (!btn) return false;
    const cls = String(btn.className || '').toLowerCase();
    const ariaDisabled = String(btn.getAttribute('aria-disabled') || '').toLowerCase();
    return !cls.includes('x-item-disabled') && !cls.includes('x-btn-disabled') && ariaDisabled !== 'true';
  }

  function findPopupSaveButton(popupTextarea) {
    const root = (popupTextarea && popupTextarea.closest && popupTextarea.closest('div.x-window')) || document;
    const prioritized = [
      'a.uft-id-save.x-btn-popupfooter-small',
      'a.x-btn.uft-id-save.x-btn-popupfooter-small',
      'a.uft-id-save',
      'a.x-btn.uft-id-save',
      'button.uft-id-save',
      'a[data-qtip="Save"]',
      'button[data-qtip="Save"]'
    ];
    for (const sel of prioritized) {
      const found = Array.from(root.querySelectorAll(sel)).find((btn) => isElementVisible(btn) && isExtButtonEnabled(btn));
      if (found) return found;
    }
    return null;
  }

  function isPopupCommentSaveButton(target) {
    const node = target && target.closest ? target.closest('a, button, span, div') : null;
    if (!node) return false;
    const popupWindow = node.closest ? node.closest('div.x-window') : null;
    if (!popupWindow) return false;
    return !!popupWindow.querySelector('textarea[name="bsccomment"], div.x-html-editor-input, iframe.x-htmleditor-iframe');
  }

  function clickPopupFooterSaveButton(btn) {
    if (!btn) return false;

    const topExt = (window.top && window.top.Ext) ? window.top.Ext : null;
    const cmpId = String(btn.getAttribute('data-componentid') || '').trim();
    if (cmpId && topExt && typeof topExt.getCmp === 'function') {
      try {
        const cmp = topExt.getCmp(cmpId);
        if (cmp && !cmp.destroyed && !cmp.isDestroyed) {
          try { if (typeof cmp.focus === 'function') cmp.focus(); } catch (_) {}
          try {
            if (typeof cmp.fireHandler === 'function') {
              cmp.fireHandler();
              return true;
            }
          } catch (_) {}
          try {
            if (typeof cmp.handler === 'function') {
              cmp.handler.call(cmp.scope || cmp, cmp);
              return true;
            }
          } catch (_) {}
          try {
            if (typeof cmp.fireEvent === 'function') {
              cmp.fireEvent('click', cmp);
              return true;
            }
          } catch (_) {}
          try {
            const cmpEl = typeof cmp.getEl === 'function' ? cmp.getEl() : cmp.el;
            if (cmpEl && cmpEl.dom && typeof cmpEl.dom.click === 'function') {
              cmpEl.dom.click();
              return true;
            }
          } catch (_) {}
        }
      } catch (_) {}
    }

    return clickElementSafe(btn);
  }

  function setPopupCommentEditorValue(editorTextarea, rawText) {
    if (!editorTextarea) return false;

    const value = String(rawText || '');
    const root = editorTextarea.closest('div.x-html-editor-input') || editorTextarea.parentElement || document;
    const normalizedHtml = value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/\r?\n/g, '<br>');
    let wroteValue = false;

    editorTextarea.value = value;
    editorTextarea.dispatchEvent(new Event('input', { bubbles: true }));
    editorTextarea.dispatchEvent(new Event('change', { bubbles: true }));
    wroteValue = true;

    const iframe = root.querySelector('iframe.x-htmleditor-iframe');
    let editorCmp = null;
    if (iframe) {
      const cmpId = String(iframe.getAttribute('data-componentid') || '').trim();
      const topExt = (window.top && window.top.Ext) ? window.top.Ext : null;
      if (cmpId && topExt && typeof topExt.getCmp === 'function') {
        try {
          editorCmp = topExt.getCmp(cmpId);
          if (editorCmp && typeof editorCmp.setValue === 'function') {
            if (typeof editorCmp.focus === 'function') editorCmp.focus();
            editorCmp.setValue(normalizedHtml);
            if (typeof editorCmp.syncValue === 'function') editorCmp.syncValue();
            if (typeof editorCmp.pushValue === 'function') editorCmp.pushValue();
            wroteValue = true;
          }
        } catch (_) {}
      }

      try {
        const doc = iframe.contentDocument || (iframe.contentWindow && iframe.contentWindow.document);
        if (doc && doc.body) {
          doc.body.innerText = value;
          doc.body.innerHTML = normalizedHtml;
          doc.body.focus();
          doc.body.dispatchEvent(new Event('input', { bubbles: true }));
          doc.body.dispatchEvent(new Event('change', { bubbles: true }));
          doc.body.dispatchEvent(new Event('keyup', { bubbles: true }));
          doc.body.dispatchEvent(new Event('blur', { bubbles: true }));
          wroteValue = true;
        }
      } catch (_) {}
    }

    if (editorCmp) {
      try {
        if (typeof editorCmp.syncValue === 'function') editorCmp.syncValue();
        if (typeof editorCmp.pushValue === 'function') editorCmp.pushValue();
      } catch (_) {}
    }

    editorTextarea.dispatchEvent(new Event('blur', { bubbles: true }));
    return wroteValue;
  }

  function getPopupCommentEditorValue(editorTextarea) {
    if (!editorTextarea) return '';
    const root = editorTextarea.closest('div.x-html-editor-input') || editorTextarea.parentElement || document;
    const iframe = root.querySelector('iframe.x-htmleditor-iframe');
    if (!iframe) return String(editorTextarea.value || '').trim();

    const cmpId = String(iframe.getAttribute('data-componentid') || '').trim();
    const topExt = (window.top && window.top.Ext) ? window.top.Ext : null;
    if (cmpId && topExt && typeof topExt.getCmp === 'function') {
      try {
        const editorCmp = topExt.getCmp(cmpId);
        if (editorCmp && typeof editorCmp.getValue === 'function') {
          const cmpVal = String(editorCmp.getValue() || '')
            .replace(/<br\s*\/?>/gi, '\n')
            .replace(/<[^>]+>/g, ' ')
            .trim();
          if (cmpVal) return cmpVal;
        }
      } catch (_) {}
    }

    try {
      const doc = iframe.contentDocument || (iframe.contentWindow && iframe.contentWindow.document);
      const bodyText = String((doc && doc.body && doc.body.innerText) || '').trim();
      if (bodyText) return bodyText;
    } catch (_) {}
    return String(editorTextarea.value || '').trim();
  }

  function normalizeComparableText(text) {
    return String(text || '').replace(/\s+/g, ' ').trim();
  }

  function showMiniToast(message, kind) {
    const id = 'myapm-mini-toast';
    let toast = document.getElementById(id);
    if (!toast) {
      toast = document.createElement('div');
      toast.id = id;
      Object.assign(toast.style, {
        position: 'fixed',
        right: '16px',
        bottom: '18px',
        zIndex: '100000',
        padding: '8px 12px',
        borderRadius: '8px',
        font: '700 12px/1.3 Arial, Helvetica, sans-serif',
        boxShadow: '0 10px 24px rgba(0,0,0,0.28)',
        opacity: '0',
        transform: 'translateY(8px)',
        transition: 'all .18s ease'
      });
      document.body.appendChild(toast);
    }
    const themes = {
      success: { bg: '#1f8f4d', fg: '#fff' },
      error: { bg: '#b63b3b', fg: '#fff' },
      info: { bg: '#29486b', fg: '#e8f1ff' }
    };
    const theme = themes[kind] || themes.info;
    toast.textContent = message;
    toast.style.background = theme.bg;
    toast.style.color = theme.fg;
    toast.style.opacity = '1';
    toast.style.transform = 'translateY(0)';
    clearTimeout(window.__myApmMiniToastTimer);
    window.__myApmMiniToastTimer = setTimeout(() => {
      const node = document.getElementById(id);
      if (!node) return;
      node.style.opacity = '0';
      node.style.transform = 'translateY(8px)';
    }, 1800);
  }

  async function copyClosingCommentsToCommentsTab(sourceTextarea, copyBtn) {
    const rawText = String((sourceTextarea && sourceTextarea.value) || '');
    if (!rawText.trim()) {
      showMiniToast('Enter closing comments first.', 'error');
      return;
    }

    if (copyBtn) {
      copyBtn.disabled = true;
      copyBtn.style.opacity = '0.65';
    }

    try {
      openRecordTabByText('Comments');

      const addCommentBtn = await waitForCondition(() => {
        const byClass = Array.from(document.querySelectorAll('a.uft-id-newcomment, a.x-btn.uft-id-newcomment'));
        const byQtip = Array.from(document.querySelectorAll('a[data-qtip], button[data-qtip]'))
          .filter((el) => /add comment/i.test(String(el.getAttribute('data-qtip') || '')));
        const candidates = byClass.concat(byQtip);
        return candidates.find(isElementVisible) || null;
      }, { timeoutMs: 6000, intervalMs: 120, observeMutations: true });

      if (!addCommentBtn) {
        showMiniToast('Add Comment button not found.', 'error');
        return;
      }

      const addCommentClicked = clickElementSafe(addCommentBtn);
      if (!addCommentClicked) {
        showMiniToast('Add Comment click failed.', 'error');
        return;
      }

      const popupTextarea = await waitForCondition(() => {
        const explicit = document.querySelector('textarea[name="bsccomment"]');
        if (explicit && explicit.closest('div.x-window')) return explicit;

        const candidates = Array.from(document.querySelectorAll('div.x-window textarea'))
          .filter((ta) =>
            ta !== sourceTextarea &&
            String(ta.name || '').toLowerCase() !== 'udfnote01' &&
            !ta.disabled &&
            !ta.readOnly
          );
        return candidates[0] || null;
      }, { timeoutMs: 7000, intervalMs: 120, observeMutations: true });

      if (!popupTextarea) {
        showMiniToast('Comment editor not found.', 'error');
        return;
      }

      const wroteEditor = setPopupCommentEditorValue(popupTextarea, rawText);
      if (!wroteEditor) {
        showMiniToast('Comment text did not populate.', 'error');
        return;
      }

      const expected = normalizeComparableText(rawText);
      const expectedSnippet = expected.slice(0, 48);
      await new Promise((resolve) => setTimeout(resolve, 120));

      const editorFilled = await waitForCondition(() => {
        const actual = normalizeComparableText(getPopupCommentEditorValue(popupTextarea));
        if (!actual) return null;
        if (!expected) return true;
        if (actual === expected) return true;
        if (expectedSnippet && actual.includes(expectedSnippet)) return true;
        return null;
      }, { timeoutMs: 3000, intervalMs: 80, observeMutations: false });

      if (!editorFilled) {
        showMiniToast('Comment text did not populate.', 'error');
        return;
      }

      const saveBtn = await waitForCondition(() => findPopupSaveButton(popupTextarea), {
        timeoutMs: 5000,
        intervalMs: 120,
        observeMutations: true
      });

      if (!saveBtn) {
        showMiniToast('Save button not found.', 'error');
        return;
      }

      await new Promise((resolve) => setTimeout(resolve, 180));
      const saveClicked = clickPopupFooterSaveButton(saveBtn);
      if (!saveClicked) {
        showMiniToast('Save click failed.', 'error');
        return;
      }

      showMiniToast('Closing comments copied to Comments tab.', 'success');
    } catch (_) {
      showMiniToast('Failed to copy to Comments tab.', 'error');
    } finally {
      if (copyBtn) {
        copyBtn.disabled = false;
        copyBtn.style.opacity = '1';
      }
    }
  }

  function isElementVisible(el) {
    if (!el || !el.isConnected) return false;
    if (el.hidden) return false;
    if (el.getClientRects().length === 0) return false;
    const style = window.getComputedStyle(el);
    return style.display !== 'none' && style.visibility !== 'hidden';
  }

  let allowFollowUpWoClickOnce = false;
  let followUpWarningOpen = false;

  function normalizeActionText(text) {
    return String(text || '')
      .toLowerCase()
      .replace(/[‐‑‒–—−]/g, '-')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function getSelectedChecklistActivityText() {
    const labels = Array.from(document.querySelectorAll('label.x-form-item-label, span.x-form-item-label-text, label, span'));
    for (const lbl of labels) {
      const labelText = String(lbl.textContent || '').trim().toLowerCase();
      if (!/^activity:?$/.test(labelText)) continue;
      const formItem = lbl.closest('div.x-form-item') || lbl.parentElement;
      const input = formItem && formItem.querySelector
        ? formItem.querySelector('input.x-form-field, input')
        : null;
      if (input && isElementVisible(input)) {
        const value = String(input.value || '').trim();
        if (value) return value;
      }
    }

    const selectedOption = Array.from(document.querySelectorAll('li.x-boundlist-item.x-boundlist-selected')).find(isElementVisible);
    if (selectedOption) {
      const value = String(selectedOption.textContent || '').trim();
      if (value) return value;
    }

    const candidates = Array.from(document.querySelectorAll('input[name*="activity" i], input[id*="activity" i], input[aria-label*="activity" i]')).filter(isElementVisible);
    for (const candidate of candidates) {
      const value = String(candidate.value || '').trim();
      if (value) return value;
    }
    return '';
  }

  function isRiskyChecklistActivity(activityText) {
    return /^\s*1\s*-\s*technician\b/i.test(String(activityText || ''));
  }

  function showFollowUpWoWarning(onContinue) {
    if (followUpWarningOpen) return;
    followUpWarningOpen = true;

    const overlay = document.createElement('div');
    overlay.id = 'apm-fwo-warning-overlay';
    Object.assign(overlay.style, {
      position: 'fixed',
      inset: '0',
      background: 'rgba(0,0,0,0.55)',
      zIndex: '200000',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      padding: '16px'
    });

    const modal = document.createElement('div');
    Object.assign(modal.style, {
      width: 'min(620px, 95vw)',
      background: '#fff',
      color: '#111',
      borderRadius: '8px',
      boxShadow: '0 12px 30px rgba(0,0,0,0.45)',
      border: '1px solid #d5d9d9',
      padding: '18px'
    });

    const title = document.createElement('div');
    title.textContent = 'WARNING';
    Object.assign(title.style, {
      color: '#b12704',
      fontSize: '18px',
      fontWeight: '700',
      marginBottom: '10px'
    });

    const body = document.createElement('div');
    body.textContent =
      'WARNING: You are about to create a FWO from the Safety Task List.\n' +
      'This will increase the Equipment Risk Index for the site, and should only be done with caution.\n' +
      'An example of when this action should be done is if a LOTO Point has been damaged and needs repair.';
    Object.assign(body.style, {
      whiteSpace: 'pre-line',
      lineHeight: '1.45',
      fontSize: '14px'
    });

    const actions = document.createElement('div');
    Object.assign(actions.style, {
      display: 'flex',
      justifyContent: 'flex-end',
      gap: '10px',
      marginTop: '16px'
    });

    const cancelBtn = document.createElement('button');
    cancelBtn.textContent = 'Cancel';
    Object.assign(cancelBtn.style, {
      padding: '6px 14px',
      border: '1px solid #c7c7c7',
      borderRadius: '4px',
      background: '#f7f7f7',
      cursor: 'pointer'
    });

    const continueBtn = document.createElement('button');
    continueBtn.textContent = 'Continue';
    Object.assign(continueBtn.style, {
      padding: '6px 14px',
      border: '1px solid #146eb4',
      borderRadius: '4px',
      background: '#146eb4',
      color: '#fff',
      cursor: 'pointer',
      fontWeight: '700'
    });

    function close() {
      overlay.remove();
      followUpWarningOpen = false;
    }

    cancelBtn.addEventListener('click', close);
    continueBtn.addEventListener('click', () => {
      close();
      if (typeof onContinue === 'function') onContinue();
    });

    actions.append(cancelBtn, continueBtn);
    modal.append(title, body, actions);
    overlay.appendChild(modal);
    document.body.appendChild(overlay);
  }

  function installFollowUpWoSafetyGuard() {
    if (window.__apmFollowUpWoGuardInstalled) return;
    window.__apmFollowUpWoGuardInstalled = true;

    document.addEventListener('click', (event) => {
      if (allowFollowUpWoClickOnce) return;

      const clickedMenuNode = event.target?.closest?.('.x-menu-item, a.x-menu-item-link, .x-menu-item-text');
      if (!clickedMenuNode) return;

      const textNode = clickedMenuNode.matches('.x-menu-item-text')
        ? clickedMenuNode
        : clickedMenuNode.querySelector('.x-menu-item-text');
      const actionText = normalizeActionText(textNode?.textContent || '');
      if (actionText !== 'create follow-up wo' && actionText !== 'create follow up wo') return;

      const activity = getSelectedChecklistActivityText();
      if (!isRiskyChecklistActivity(activity)) return;

      const actionTarget = clickedMenuNode.closest('a.x-menu-item-link') || clickedMenuNode;

      event.preventDefault();
      event.stopPropagation();
      event.stopImmediatePropagation();

      showFollowUpWoWarning(() => {
        allowFollowUpWoClickOnce = true;
        requestAnimationFrame(() => {
          try {
            if (typeof actionTarget.click === 'function') {
              actionTarget.click();
            } else {
              actionTarget.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true, view: window }));
            }
          } finally {
            setTimeout(() => { allowFollowUpWoClickOnce = false; }, 0);
          }
        });
      });
    }, true);
  }

  function normalizeDateValue(value) {
    if (!value) return null;
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      const date = new Date(value.getTime());
      date.setHours(0, 0, 0, 0);
      return date;
    }

    const raw = String(value).trim();
    const usMatch = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (usMatch) {
      let year = Number(usMatch[3]);
      if (year < 100) year += 2000;
      const date = new Date(year, Number(usMatch[1]) - 1, Number(usMatch[2]));
      if (!Number.isNaN(date.getTime())) {
        date.setHours(0, 0, 0, 0);
        return date;
      }
    }

    const isoMatch = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (isoMatch) {
      const date = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
      if (!Number.isNaN(date.getTime())) {
        date.setHours(0, 0, 0, 0);
        return date;
      }
    }

    const asDate = new Date(raw);
    if (!Number.isNaN(asDate.getTime())) {
      asDate.setHours(0, 0, 0, 0);
      return asDate;
    }
    return null;
  }

  function getExtComponentQuery() {
    const ext = window.Ext;
    if (!ext || !ext.ComponentQuery || typeof ext.ComponentQuery.query !== 'function') return null;
    return ext.ComponentQuery;
  }

  function getFirstVisibleFieldValue(selectors) {
    const componentQuery = getExtComponentQuery();
    if (!componentQuery || !Array.isArray(selectors)) return null;
    for (const selector of selectors) {
      const matches = componentQuery.query(selector) || [];
      for (const cmp of matches) {
        if (!cmp) continue;
        const hidden = cmp.hidden === true || (typeof cmp.isHidden === 'function' && cmp.isHidden());
        const rendered = cmp.rendered !== false;
        if (hidden || !rendered) continue;
        const el = typeof cmp.getEl === 'function' ? cmp.getEl() : cmp.el;
        if (el && el.dom && !isElementVisible(el.dom)) continue;
        const value = typeof cmp.getValue === 'function' ? cmp.getValue() : cmp.value;
        if (value !== null && typeof value !== 'undefined' && String(value).trim() !== '') return value;
      }
    }
    return null;
  }

  function getFieldValueFromVisibleLabel(labelPatterns) {
    const labels = Array.from(document.querySelectorAll('label.x-form-item-label, span.x-form-item-label-text, label, span'));
    for (const lbl of labels) {
      const text = String(lbl.textContent || '').replace(/\s+/g, ' ').trim();
      if (!text) continue;
      if (!labelPatterns.some((pattern) => pattern.test(text))) continue;
      const formItem = lbl.closest('div.x-form-item') || lbl.parentElement;
      if (!formItem) continue;
      const input = formItem.querySelector('input.x-form-field, input, textarea');
      if (!input || !isElementVisible(input)) continue;
      const value = String(input.value || '').trim();
      if (value) return value;
    }
    return null;
  }

  function getSaveDateGuardState() {
    const schedEndRaw = getFirstVisibleFieldValue([
      'datefield[name=ff_schedenddate]',
      'datetimefield[name=ff_schedenddate]',
      'textfield[name=ff_schedenddate]',
      'field[name=ff_schedenddate]',
      'datefield[name*=schedenddate]',
      'datetimefield[name*=schedenddate]',
      'textfield[name*=schedenddate]',
      'field[name*=schedenddate]'
    ]) || getFieldValueFromVisibleLabel([/^sched\.?\s*end\s*date:?$/i, /^scheduled\s*end\s*date:?$/i, /^sched\s*end:?$/i]);

    const originalDueRaw = getFirstVisibleFieldValue([
      'datefield[name=ff_duedate]',
      'datetimefield[name=ff_duedate]',
      'textfield[name=ff_duedate]',
      'field[name=ff_duedate]',
      'datefield[name=ff_originalduedate]',
      'datetimefield[name=ff_originalduedate]',
      'textfield[name=ff_originalduedate]',
      'field[name=ff_originalduedate]',
      'datefield[name*=duedate]',
      'datetimefield[name*=duedate]',
      'textfield[name*=duedate]',
      'field[name*=duedate]'
    ]) || getFieldValueFromVisibleLabel([/^original\s*pm\s*due\s*date:?$/i, /^original\s*due\s*date:?$/i, /^due\s*date:?$/i]);

    const schedEndDate = normalizeDateValue(schedEndRaw);
    const originalDueDate = normalizeDateValue(originalDueRaw);
    return {
      schedEndDate,
      originalDueDate,
      isBlocked: !!(schedEndDate && originalDueDate && schedEndDate > originalDueDate)
    };
  }

  function looksLikeSaveButton(target) {
    const node = target && target.closest ? target.closest('a, button, span, div') : null;
    if (!node) return false;
    const text = normalizeActionText(node.textContent || node.getAttribute('aria-label') || node.getAttribute('data-qtip') || node.getAttribute('title') || '');
    const cls = String(node.className || '').toLowerCase();
    const qtip = normalizeActionText(node.getAttribute('data-qtip') || '');
    return text === 'save' || qtip === 'save' || cls.includes('toolbarsave') || cls.includes('uft-id-saverec') || cls.includes('uft-id-save');
  }

  function blockScheduledEndSave(event) {
    if (!looksLikeSaveButton(event.target)) return;
    if (isPopupCommentSaveButton(event.target)) return;
    const state = getSaveDateGuardState();
    if (!state.isBlocked) return;
    event.preventDefault();
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === 'function') event.stopImmediatePropagation();
    showMiniToast('Cannot save: Scheduled End Date is past the Original PM Due Date.', 'error');
  }

  function isSaveShortcut(event) {
    if (!event) return false;
    const key = String(event.key || '').toLowerCase();
    const code = String(event.code || '').toLowerCase();
    return (event.ctrlKey || event.metaKey) && !event.altKey && (key === 's' || code === 'keys');
  }

  function blockScheduledEndSaveShortcut(event) {
    if (!isSaveShortcut(event)) return;
    const state = getSaveDateGuardState();
    if (!state.isBlocked) return;
    event.preventDefault();
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === 'function') event.stopImmediatePropagation();
    showMiniToast('Cannot save: Scheduled End Date is past the Original PM Due Date.', 'error');
  }

  function installScheduledEndSaveGuard() {
    if (window.__apmScheduledEndSaveGuardInstalled) return;
    window.__apmScheduledEndSaveGuardInstalled = true;

    document.addEventListener('mousedown', blockScheduledEndSave, true);
    document.addEventListener('click', blockScheduledEndSave, true);
    document.addEventListener('keydown', blockScheduledEndSaveShortcut, true);
  }

  function installClosingCommentsCounter() {
    const textareas = Array.from(document.querySelectorAll('textarea[name="udfnote01"]')).filter(isElementVisible);
    if (!textareas.length) return;

    textareas.forEach((textarea) => {
      const formItem = textarea.closest('div.x-form-item') || textarea.parentElement;
      let labelSpan = null;

      if (formItem) {
        labelSpan = Array.from(formItem.querySelectorAll('span.x-form-item-label-text, label, span'))
          .find((el) => /closing comments:?/i.test(String(el.textContent || '').trim()));
      }

      if (!labelSpan) {
        labelSpan = Array.from(document.querySelectorAll('span.x-form-item-label-text, label, span'))
          .find((el) => /closing comments:?/i.test(String(el.textContent || '').trim()));
      }

      if (!labelSpan) return;

      const host = labelSpan.parentElement || labelSpan;
      const buttonHost = formItem || textarea.parentElement || host;
      if (!buttonHost) return;

      let copyBtn = buttonHost.querySelector('.apm-closing-comments-copy-btn');
      if (!copyBtn) {
        copyBtn = document.createElement('button');
        copyBtn.type = 'button';
        copyBtn.className = 'apm-closing-comments-copy-btn';
        copyBtn.textContent = 'Copy to Comments';
        Object.assign(copyBtn.style, {
          display: 'block',
          padding: '2px 6px',
          fontSize: '11px',
          fontFamily: 'inherit',
          lineHeight: '1.3',
          border: '1px solid #3c516d',
          borderRadius: '3px',
          background: 'linear-gradient(to bottom, #2a3a50, #1b2738)',
          color: '#d6deee',
          cursor: 'pointer',
          whiteSpace: 'nowrap'
        });
        buttonHost.appendChild(copyBtn);
      }

      let counterEl = buttonHost.querySelector('.apm-closing-comments-counter');
      if (!counterEl) {
        counterEl = document.createElement('div');
        counterEl.className = 'apm-closing-comments-counter';
        Object.assign(counterEl.style, {
          display: 'block',
          fontSize: '11px',
          fontWeight: 'normal',
          lineHeight: '1.2',
          opacity: '0.95',
          position: 'absolute',
          zIndex: '5',
          whiteSpace: 'nowrap',
          pointerEvents: 'none'
        });
        buttonHost.appendChild(counterEl);
      }

      const hostStyle = window.getComputedStyle(buttonHost);
      if (hostStyle.position === 'static') buttonHost.style.position = 'relative';

      const hostRect = buttonHost.getBoundingClientRect();
      const textRect = textarea.getBoundingClientRect();
      const left = Math.max(0, Math.round(textRect.right - hostRect.left + 8));
      const top = Math.max(0, Math.round(textRect.top - hostRect.top));
      Object.assign(copyBtn.style, {
        position: 'absolute',
        left: `${left}px`,
        top: `${top}px`,
        marginTop: '0',
        zIndex: '5'
      });
      Object.assign(counterEl.style, {
        left: `${left}px`,
        top: `${top + copyBtn.offsetHeight + 4}px`,
        textAlign: 'left'
      });

      const updateCounter = () => {
        const len = String(textarea.value || '').length;
        const reachedMin = len >= 160;
        counterEl.textContent = `${len}/160${reachedMin ? ' ✓' : ''}`;
        counterEl.style.color = reachedMin ? '#2ecc71' : '';
        counterEl.style.fontWeight = reachedMin ? '700' : 'normal';
      };

      if (!textarea.dataset.apmClosingCommentsCounterBound) {
        textarea.addEventListener('input', updateCounter);
        textarea.addEventListener('keyup', updateCounter);
        textarea.dataset.apmClosingCommentsCounterBound = 'true';
      }
      if (!copyBtn.dataset.apmCopyToCommentsBound) {
        copyBtn.addEventListener('click', () => copyClosingCommentsToCommentsTab(textarea, copyBtn));
        copyBtn.dataset.apmCopyToCommentsBound = 'true';
      }
      updateCounter();
    });
  }

  function initPorts() {
    installParentPtpTimer();
    installIframePtpBridge();
    installClosingCommentsCounter();
    installFollowUpWoSafetyGuard();
    installScheduledEndSaveGuard();
  }

  let portsInitTimer = null;
  function schedulePortsInit(delay = 120) {
    if (portsInitTimer) clearTimeout(portsInitTimer);
    portsInitTimer = setTimeout(() => {
      portsInitTimer = null;
      initPorts();
    }, delay);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => schedulePortsInit(0), { once: true });
  } else {
    schedulePortsInit(0);
  }
  window.addEventListener('load', () => schedulePortsInit(0), { once: true });
  window.addEventListener('focus', () => schedulePortsInit(80));
  window.addEventListener('hashchange', () => schedulePortsInit(120));
  document.addEventListener('visibilitychange', () => {
    if (!document.hidden) schedulePortsInit(80);
  });
  setInterval(() => schedulePortsInit(0), 1500);
})();
