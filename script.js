// Initialize theme and placeholders; user must upload a file now (no default data.json)
const THEME_STORAGE_KEY = 'copilotMetricsTheme';

window.addEventListener('load', () => {
    syncThemeDocumentState(getPreferredTheme());
    setupHighchartsTheme();
    updateThemeToggleState(getCurrentTheme());
    setStatus('Ready for a metrics export. Upload a JSON or JSON Lines file to populate the dashboard.');
    renderPlaceholders();
    warnIfFileOrigin();
});

const CHAT_FEATURES = new Set([
    'chat_inline',
    'chat_panel_agent_mode',
    'chat_panel_ask_mode',
    'chat_panel_custom_mode',
    'chat_panel_edit_mode',
    'chat_panel_plan_mode',
    'chat_panel_unknown_mode'
]);

const CHAT_MODE_FEATURES = [
    'chat_panel_ask_mode',
    'chat_panel_edit_mode',
    'chat_panel_plan_mode',
    'chat_panel_agent_mode',
    'chat_panel_custom_mode',
    'chat_panel_unknown_mode'
];

function warnIfFileOrigin() {
    try {
        if (location.protocol === 'file:') {
            const msg = 'Running from file:// origin. Some report downloads may fail due to CORS. Serve locally (e.g. python3 -m http.server) to avoid CORS issues.';
            console.warn('[cors]', msg);
            const status = document.getElementById('statusMessage');
            if (status && !status.textContent.includes('CORS')) setStatus(msg);
        }
    } catch (_) { /* ignore */ }
}

function prefersReducedMotion() {
    return !!window.matchMedia?.('(prefers-reduced-motion: reduce)').matches;
}

function preferredScrollBehavior() {
    return prefersReducedMotion() ? 'auto' : 'smooth';
}

function waitForPaint(frames = 1) {
    const runFrame = () => new Promise(resolve => {
        if (typeof window.requestAnimationFrame === 'function') {
            window.requestAnimationFrame(() => resolve());
            return;
        }
        setTimeout(resolve, 16);
    });
    return Array.from({ length: Math.max(1, frames) }).reduce(
        promise => promise.then(runFrame),
        Promise.resolve()
    );
}

function getStoredTheme() {
    const storedTheme = localStorage.getItem(THEME_STORAGE_KEY);
    return storedTheme === 'dark' || storedTheme === 'light' ? storedTheme : null;
}

function getSystemTheme() {
    return window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function getPreferredTheme() {
    return getStoredTheme() || getSystemTheme();
}

function getCurrentTheme() {
    return document.documentElement.dataset.theme === 'dark' ? 'dark' : 'light';
}

function syncThemeDocumentState(theme) {
    const resolvedTheme = theme === 'dark' ? 'dark' : 'light';
    document.documentElement.dataset.theme = resolvedTheme;
    if (document.body) {
        document.body.classList.remove('light', 'dark');
        document.body.classList.add(resolvedTheme);
    }
    return resolvedTheme;
}

function getThemeTokens() {
    const styles = getComputedStyle(document.documentElement);
    const read = name => styles.getPropertyValue(name).trim();
    return {
        text: read('--text'),
        textDim: read('--text-dim'),
        accent: read('--accent'),
        accentStrong: read('--accent-strong'),
        panel: read('--panel'),
        panelBorder: read('--panel-border'),
        chartGrid: read('--chart-grid'),
        exportSurface: read('--export-surface'),
        heatStops: [
            read('--heatmap-1'),
            read('--heatmap-2'),
            read('--heatmap-3'),
            read('--heatmap-4')
        ],
        heatBorder: read('--heatmap-border'),
        chartColors: Array.from({ length: 10 }, (_, index) => read(`--chart-${index + 1}`)).filter(Boolean)
    };
}

function updateThemeToggleState(theme = getCurrentTheme()) {
    const toggle = document.getElementById('themeToggle');
    if (!toggle) return;
    const isDark = theme === 'dark';
    toggle.setAttribute('aria-checked', isDark ? 'true' : 'false');
    toggle.setAttribute('aria-label', isDark ? 'Switch to light mode' : 'Switch to dark mode');
}

function refreshThemeCharts() {
    setupHighchartsTheme();
    const activeData = Array.isArray(window.__currentFilteredData) ? window.__currentFilteredData : [];
    if (activeData.length && window.__dashboardModel) {
        renderCharts(window.__dashboardModel);
    }
}

function applyTheme(theme, { persist = true, refreshCharts = true } = {}) {
    const resolvedTheme = syncThemeDocumentState(theme);
    if (persist) {
        localStorage.setItem(THEME_STORAGE_KEY, resolvedTheme);
    }
    updateThemeToggleState(resolvedTheme);
    if (refreshCharts) refreshThemeCharts();
    return resolvedTheme;
}

function initializeThemeToggle() {
    const toggle = document.getElementById('themeToggle');
    const resolvedTheme = syncThemeDocumentState(getPreferredTheme());
    updateThemeToggleState(resolvedTheme);
    if (!toggle) return;

    toggle.addEventListener('click', () => {
        const nextTheme = getCurrentTheme() === 'dark' ? 'light' : 'dark';
        applyTheme(nextTheme);
    });

    const systemThemeQuery = window.matchMedia ? window.matchMedia('(prefers-color-scheme: dark)') : null;
    if (systemThemeQuery && typeof systemThemeQuery.addEventListener === 'function') {
        systemThemeQuery.addEventListener('change', event => {
            if (getStoredTheme()) return;
            applyTheme(event.matches ? 'dark' : 'light', { persist: false });
        });
    }
}

function setupHighchartsTheme() {
    if (typeof Highcharts === 'undefined') return;
    const tokens = getThemeTokens();
    Highcharts.setOptions({
        chart: {
            backgroundColor: 'transparent',
            style: { fontFamily: 'ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif' }
        },
        colors: tokens.chartColors.length ? tokens.chartColors : ['#0b5bd3', '#0f766e', '#7c3aed', '#b45309', '#2563eb', '#be185d', '#1d4ed8', '#15803d', '#9333ea', '#c2410c'],
        title: { style: { color: tokens.text, fontWeight: '700', fontSize: '15px' } },
        subtitle: { style: { color: tokens.textDim, fontSize: '12px' } },
        xAxis: {
            lineColor: tokens.panelBorder,
            tickColor: tokens.panelBorder,
            gridLineColor: tokens.chartGrid,
            labels: { style: { color: tokens.textDim, fontSize: '12px' } },
            title: { style: { color: tokens.textDim, fontSize: '12px' } }
        },
        yAxis: {
            lineColor: tokens.panelBorder,
            tickColor: tokens.panelBorder,
            gridLineColor: tokens.chartGrid,
            labels: { style: { color: tokens.textDim, fontSize: '12px' } },
            title: { style: { color: tokens.textDim, fontSize: '12px' } }
        },
        legend: {
            backgroundColor: 'transparent',
            itemStyle: { color: tokens.text, fontSize: '12px' },
            itemHoverStyle: { color: tokens.accentStrong, fontSize: '12px' }
        },
        tooltip: {
            backgroundColor: tokens.panel,
            borderColor: tokens.panelBorder,
            style: { color: tokens.text, fontSize: '12px' },
            valueDecimals: 0,
            // Show the specific section (category / slice / point) name instead of the chart title
            formatter: function() {
                const point = this.point || {};
                const val = (typeof point.y !== 'undefined') ? point.y : this.y;
                // Determine the most descriptive label available
                let label = point.name || point.category || this.key;
                // Fallback to series name only if we still don't have a label
                if (!label && this.series) label = this.series.name;
                // Escape if Highcharts provides helper
                if (Highcharts.escapeHTML) label = Highcharts.escapeHTML(label);
                return `<span style="font-weight:600">${label}</span><br/>${Highcharts.numberFormat(val, 0, '.', ',')}`;
            }
        },
        plotOptions: {
            column: { borderRadius: 2, borderWidth: 0 },
            bar: { borderRadius: 2, borderWidth: 0 },
            pie: { dataLabels: { style: { fontSize: '12px', color: tokens.text } } }
        },
        credits: { enabled: false }
    });
}

// Enhanced manual file selection & parsing (supports JSON array or JSONL)
let fileInputEl, analyzeBtnEl, fetchApiBtnEl;
document.addEventListener('DOMContentLoaded', () => {
    fileInputEl = document.getElementById('jsonFileInput');
    analyzeBtnEl = document.getElementById('analyzeBtn');
    fetchApiBtnEl = document.getElementById('fetchApiBtn');
    const fileTriggerEl = document.getElementById('jsonFileTrigger');

    initializeThemeToggle();
    
    // API members toggle handling
    const apiToggle = document.getElementById('apiMembersToggle');
    if (apiToggle) {
        apiToggle.addEventListener('change', handleApiToggleChange);
        handleApiToggleChange();
    }
    
    if (analyzeBtnEl) {
        analyzeBtnEl.addEventListener('click', () => handleFileSelection(fileInputEl && fileInputEl.files[0]));
    }
    if (fileTriggerEl && fileInputEl) {
        fileTriggerEl.addEventListener('click', () => fileInputEl.click());
    }
    if (fetchApiBtnEl) {
        fetchApiBtnEl.addEventListener('click', handleApiDataFetch);
    }
    // Prevent implicit form submission via Enter key
    const filtersForm = document.getElementById('filtersForm');
    if (filtersForm) {
        filtersForm.addEventListener('submit', e => e.preventDefault());
        filtersForm.addEventListener('keydown', e => {
            if (e.key === 'Enter' && e.target && (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA')) {
                e.preventDefault();
            }
        });
    }
    const downloadBtn = document.getElementById('downloadPdfBtn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', () => {
            if (!window.__rawData || !window.__rawData.length) return;
            generatePdfReport();
        });
    }
    if (fileInputEl) {
        syncSelectedFileName(fileInputEl, 'jsonFileName');
        fileInputEl.addEventListener('change', () => {
            syncSelectedFileName(fileInputEl, 'jsonFileName');
            if (fileInputEl.files && fileInputEl.files[0]) handleFileSelection(fileInputEl.files[0]);
        });
    }
    const membersInputEl = document.getElementById('membersFileInput');
    if (membersInputEl) {
        membersInputEl.addEventListener('change', () => membersInputEl.files && handleMembersFile(membersInputEl.files[0]));
    }

    // Per-user usage view navigation & export
    const userUsageBtn = document.getElementById('userUsageBtn');
    const backBtn = document.getElementById('backToDashboardBtn');
    const exportUsersCsvBtn = document.getElementById('exportUsersCsvBtn');
    if (userUsageBtn) {
        userUsageBtn.addEventListener('click', () => {
            buildUserUsageTable(window.__currentFilteredData || window.__rawData || []);
            toggleUserUsage(true);
        });
    }
    if (backBtn) {
        backBtn.addEventListener('click', () => toggleUserUsage(false));
    }
    if (exportUsersCsvBtn) {
        exportUsersCsvBtn.addEventListener('click', () => exportUserUsageCsv());
    }
});

function syncSelectedFileName(inputEl, outputId, emptyText = 'No file chosen') {
    const outputEl = document.getElementById(outputId);
    if (!inputEl || !outputEl) return;
    const fileName = inputEl.files && inputEl.files[0] ? inputEl.files[0].name : emptyText;
    outputEl.textContent = fileName;
    outputEl.title = fileName;
}

function looksLikeUsageRecordObject(obj) {
    if (!obj || typeof obj !== 'object') return false;
    return [
        'day',
        'day_totals',
        'enterprise_id',
        'organization_id',
        'pull_requests',
        'report_end_day',
        'report_start_day',
        'totals_by_cli',
        'totals_by_feature',
        'totals_by_ide',
        'user_id',
        'user_login'
    ].some(key => Object.prototype.hasOwnProperty.call(obj, key));
}

function normalizeUsageRecords(records) {
    const normalized = [];
    (records || []).forEach(record => flattenUsageRecord(record, normalized));
    return normalized;
}

function flattenUsageRecord(record, output) {
    if (!record || typeof record !== 'object') return;
    if (Array.isArray(record.day_totals)) {
        const sharedMeta = {
            enterprise_id: record.enterprise_id,
            organization_id: record.organization_id,
            report_start_day: record.report_start_day,
            report_end_day: record.report_end_day,
            etl_id: record.etl_id,
            day_partition: record.day_partition,
            entity_id_partition: record.entity_id_partition
        };
        record.day_totals.forEach(dayTotal => {
            if (!dayTotal || typeof dayTotal !== 'object') return;
            output.push({
                ...sharedMeta,
                ...dayTotal,
                __record_scope: 'aggregate'
            });
        });
        return;
    }
    output.push({
        ...record,
        __record_scope: (record.user_id || record.user_login) ? 'user' : 'daily'
    });
}

function parseUploadedText(text) {
    // Try full JSON parse first (array or object with records?)
    try {
        const preliminary = JSON.parse(text);
        if (Array.isArray(preliminary)) { console.info('[upload] Parsed as JSON array'); return preliminary; }
        if (looksLikeUsageRecordObject(preliminary)) {
            console.info('[upload] Parsed as usage record object');
            return [preliminary];
        }
        // If object, attempt to find an array property with objects containing user_id or day
        const candidateKey = Object.keys(preliminary).find(k => Array.isArray(preliminary[k]) && preliminary[k].length && typeof preliminary[k][0] === 'object');
        if (candidateKey) { console.info('[upload] Parsed as object wrapper key=' + candidateKey); return preliminary[candidateKey]; }
        console.info('[upload] Parsed as single object');
        return [preliminary];
    } catch (_) { /* fallthrough to JSONL */ }
    // JSON Lines (skip blank/comment lines)
    const lines = text.split(/\r?\n/).filter(l => l.trim() && !l.trim().startsWith('#'));
    const data = [];
    for (let i = 0; i < lines.length; i++) {
        try {
            data.push(JSON.parse(lines[i]));
        } catch (e) {
            throw new Error(`Line ${i+1}: ${e.message}`);
        }
    }
    return data;
}

function handleFileSelection(file) {
    if (!file) { setStatus('No file selected. Please choose a JSON / JSONL export file.'); return; }
    showLoading(true);
    setStatus(`Reading ${file.name} …`);
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const text = e.target.result;
            const data = parseUploadedText(text);
            const normalized = normalizeUsageRecords(data);
            if (!Array.isArray(normalized) || !normalized.length) {
                throw new Error('Parsed result is empty.');
            }
            window.__sourceData = data;
            window.__rawData = normalized;
            initializeFilters(normalized);
            analyzeData(normalized);
            setStatus(`Loaded ${normalized.length} usable records from ${file.name}`);
            enableDownloadButton();
        } catch (err) {
            console.error(err);
            setStatus(`Upload parse error: ${err.message}`, true);
        } finally {
            showLoading(false);
        }
    };
    reader.onerror = () => { setStatus('File read error', true); showLoading(false); };
    reader.readAsText(file);
}

// --- Data source toggle handling ---
function handleApiToggleChange() {
    const enabled = document.getElementById('apiMembersToggle')?.checked;
    const apiControls = document.querySelectorAll('.api-input');
    const fetchBtn = document.getElementById('fetchApiBtn');
    apiControls.forEach(el => el.style.display = enabled ? 'flex' : 'none');
    if (fetchBtn) fetchBtn.style.display = enabled ? 'inline-flex' : 'none';
    const membersFileWrapper = document.getElementById('membersFileWrapper');
    if (membersFileWrapper) membersFileWrapper.style.display = enabled ? 'none' : 'flex';
}

// --- GitHub API (members only) ---
async function handleApiDataFetch() {
    const pat = document.getElementById('githubPat').value.trim();
    const org = document.getElementById('orgNameApi').value.trim();
    if (!pat) { setStatus('PAT required to fetch members.', true); return; }
    if (!org) { setStatus('Organization name required.', true); return; }
    showLoading(true);
    setStatus('Fetching organization members…');
    try {
        const members = await fetchOrganizationMembers(pat, org);
        if (!members.length) throw new Error('No members returned.');
        const logins = new Set(); members.forEach(m => { if (m.login) logins.add(m.login.toLowerCase()); });
        window.__membersSet = logins;
        updateMembersStatus();
        setStatus(`Loaded ${logins.size} members. Reapply the current filters or load a metrics file if you have not uploaded one yet.`);
        if (document.getElementById('membersOnlyChk')?.checked && window.__rawData) applyFilters();
    } catch (err) {
        console.error('Members fetch error:', err);
        setStatus(`Members fetch failed: ${err.message}`, true);
    } finally { showLoading(false); }
}

async function fetchOrganizationMembers(pat, orgName) {
    const response = await fetch(`https://api.github.com/orgs/${orgName}/members`, {
        headers: {
            'Authorization': `Bearer ${pat}`,
            'Accept': 'application/vnd.github.v3+json',
            'X-GitHub-Api-Version': '2022-11-28'
        }
    });
    
    if (!response.ok) {
        if (response.status === 401) {
            throw new Error('Invalid GitHub token or insufficient permissions. Token needs "read:org" scope.');
        } else if (response.status === 404) {
            throw new Error(`Organization "${orgName}" not found or token lacks permission to view members.`);
        } else {
            throw new Error(`GitHub API error: ${response.status} ${response.statusText}`);
        }
    }
    
    return await response.json();
}

// Removed legacy fetchCopilotMetrics / normalizeApiData in favor of report-based ingestion.

// --- Members file handling (org members export) ---
function handleMembersFile(file) {
    if (!file) return;
    setStatus(`Reading members file ${file.name} …`);
    const reader = new FileReader();
    reader.onload = e => {
        try {
            const text = e.target.result;
            let membersData = parseUploadedText(text);
            if (!Array.isArray(membersData) || !membersData.length) throw new Error('Members file is empty');
            // Accept objects with login / user_login / name; build set of logins
            const logins = new Set();
            membersData.forEach(m => {
                if (m && typeof m === 'object') {
                    const login = m.login || m.user_login || m.user || m.name; // fallback guesses
                    if (login) logins.add(String(login).toLowerCase());
                } else if (typeof m === 'string') {
                    logins.add(m.toLowerCase());
                }
            });
            if (!logins.size) throw new Error('No recognizable login fields in members file');
            window.__membersSet = logins;
            updateMembersStatus();
            setStatus(`Loaded ${logins.size} members from ${file.name}`);
            // If user already checked members-only, reapply filters
            if (document.getElementById('membersOnlyChk')?.checked && window.__rawData) {
                applyFilters();
            }
        } catch (err) {
            console.error(err);
            setStatus(`Members parse error: ${err.message}`, true);
        }
    };
    reader.onerror = () => setStatus('Members file read error', true);
    reader.readAsText(file);
}

function updateMembersStatus() {
    const statusEl = document.getElementById('membersStatus');
    const chk = document.getElementById('membersOnlyChk');
    if (!statusEl || !chk) return;
    const size = window.__membersSet ? window.__membersSet.size : 0;
    statusEl.textContent = size ? `${size} loaded` : '(none loaded yet)';
    chk.disabled = size === 0;
    if (size === 0) chk.checked = false;
}


function loadDataFile(filename) {
    setStatus(`Loading ${filename} …`);
    showLoading(true);
    fetch(filename)
        .then(response => {
            if (!response.ok) {
                throw new Error(`Failed to load ${filename}: ${response.statusText}`);
            }
            return response.text();
        })
        .then(text => {
            try {
                const data = parseUploadedText(text);
                const normalized = normalizeUsageRecords(data);
                window.__sourceData = data;
                window.__rawData = normalized;
                initializeFilters(normalized);
                analyzeData(normalized);
                enableDownloadButton();
            } catch (error) {
                setStatus(`Parse error: ${error.message}`, true);
            }
        })
        .catch(error => {
            setStatus(`Load error: ${error.message}`, true);
        })
        .finally(() => {
            showLoading(false);
        });
}

function analyzeData(data) {
    const chartsContainer = document.getElementById('chartsContainer');
    if (chartsContainer) chartsContainer.innerHTML = '';

    if (!data || data.length === 0) {
        const hasLoadedData = Boolean(window.__rawData && window.__rawData.length);
        setStatus(
            hasLoadedData
                ? 'No records match the current filters. Widen the date range, clear user search, or turn off the members-only filter.'
                : 'Ready for a metrics export. Upload a JSON or JSON Lines file to populate the dashboard.'
        );
        const tablesSection = document.getElementById('tablesSection');
        if (tablesSection) tablesSection.hidden = true;
        renderPlaceholders(hasLoadedData ? 'filters' : 'upload');
        const downloadBtn = document.getElementById('downloadPdfBtn');
        if (downloadBtn) downloadBtn.disabled = true;
        const userUsageBtn = document.getElementById('userUsageBtn');
        if (userUsageBtn) userUsageBtn.disabled = true;
        return;
    }

    const model = buildDashboardModel(data);
    window.__dashboardModel = model;
    window.__currentFilteredData = data;

    computeSummaryMetrics(model);
    renderCharts(model);
    renderReferenceTables(model);
    enableDownloadButton();

    const userUsageBtn = document.getElementById('userUsageBtn');
    if (userUsageBtn) userUsageBtn.disabled = !model.meta.hasUserRecords;
    if (!model.meta.hasUserRecords) {
        const userUsageSection = document.getElementById('userUsageSection');
        if (userUsageSection && !userUsageSection.hidden) toggleUserUsage(false);
    }

    const scopeLabel = model.meta.hasAggregateRecords
        ? (model.meta.hasUserRecords ? 'mixed export' : 'aggregate export')
        : 'user export';
    setStatus(`Displaying ${model.meta.recordCount.toLocaleString()} records across ${model.meta.days.length.toLocaleString()} days (${scopeLabel})`);

    const main = document.getElementById('mainContent');
    if (main) {
        setTimeout(() => { main.focus(); }, 0);
    }
}

function buildDashboardModel(data) {
    const userRecords = data.filter(isUserLevelRecord);
    const aggregateRecords = data.filter(record => !isUserLevelRecord(record));
    const overallRecords = aggregateRecords.length ? aggregateRecords : data;

    const featureTotals = new Map();
    const ideTotals = new Map();
    const languageTotals = new Map();
    const modelTotals = new Map();
    const modelFeatureMatrix = new Map();
    const languageModelMatrix = new Map();
    const dayLanguageMap = new Map();
    const dayModelMap = new Map();
    const dayOverviewMap = new Map();
    const cliDayMap = new Map();
    const pullRequestDayMap = new Map();
    const dayUserSets = new Map();
    const dayCliUserSets = new Map();
    const dayReviewActiveSets = new Map();
    const dayReviewPassiveSets = new Map();
    const weekUserSets = new Map();

    const uniqueUsers = new Set();
    const chatUsers = new Set();
    const agentUsers = new Set();
    const cliUsers = new Set();
    const reviewActiveUsers = new Set();
    const reviewPassiveUsers = new Set();

    let totalInteractions = 0;
    let totalGenerations = 0;
    let totalAcceptances = 0;
    let hasCli = false;
    let hasPullRequests = false;
    let hasCodeReview = false;

    overallRecords.forEach(record => {
        const day = record.day || '';
        if (!day) return;

        const dayRow = getOrCreateMapValue(dayOverviewMap, day, () => createDailyOverviewRow(day));
        totalInteractions += toNum(record.user_initiated_interaction_count);
        totalGenerations += toNum(record.code_generation_activity_count);
        totalAcceptances += toNum(record.code_acceptance_activity_count);

        dayRow.interactions += toNum(record.user_initiated_interaction_count);
        dayRow.code_generations += toNum(record.code_generation_activity_count);
        dayRow.acceptances += toNum(record.code_acceptance_activity_count);
        dayRow.loc_suggested_add += toNum(record.loc_suggested_to_add_sum);
        dayRow.loc_suggested_delete += toNum(record.loc_suggested_to_delete_sum);
        dayRow.loc_added += toNum(record.loc_added_sum);
        dayRow.loc_deleted += toNum(record.loc_deleted_sum);
        dayRow.daily_active_users += toNum(record.daily_active_users);
        dayRow.weekly_active_users += toNum(record.weekly_active_users);
        dayRow.monthly_active_users += toNum(record.monthly_active_users);
        dayRow.monthly_active_chat_users += toNum(record.monthly_active_chat_users);
        dayRow.monthly_active_agent_users += toNum(record.monthly_active_agent_users);
        dayRow.daily_active_cli_users += toNum(record.daily_active_cli_users);

        if (Array.isArray(record.totals_by_feature)) {
            record.totals_by_feature.forEach(featureBucket => {
                const feature = featureBucket.feature || 'unknown';
                accumulateMetricBucket(featureTotals, feature, featureBucket);
            });
        }

        if (Array.isArray(record.totals_by_ide)) {
            record.totals_by_ide.forEach(ideBucket => {
                const ide = ideBucket.ide || 'unknown';
                const row = accumulateMetricBucket(ideTotals, ide, ideBucket);
                updateLatestVersion(row, 'latest_ide_version', ideBucket.last_known_ide_version, info => info.ide_version || '');
                updateLatestVersion(row, 'latest_plugin_version', ideBucket.last_known_plugin_version, info => {
                    const parts = [info.plugin, info.plugin_version].filter(Boolean);
                    return parts.join(' ') || info.plugin_version || '';
                });
            });
        }

        if (Array.isArray(record.totals_by_language_feature)) {
            record.totals_by_language_feature.forEach(languageBucket => {
                const language = languageBucket.language || 'unknown';
                accumulateMetricBucket(languageTotals, language, languageBucket);
                accumulateNestedMetric(dayLanguageMap, day, language, pickPrimaryUsageValue(languageBucket));
            });
        }

        if (Array.isArray(record.totals_by_model_feature)) {
            record.totals_by_model_feature.forEach(modelBucket => {
                const model = modelBucket.model || 'unknown';
                const feature = modelBucket.feature || 'unknown';
                accumulateMetricBucket(modelTotals, model, modelBucket);
                accumulateNestedMetric(dayModelMap, day, model, pickPrimaryUsageValue(modelBucket));
                accumulateNestedMetric(modelFeatureMatrix, feature, model, pickPrimaryUsageValue(modelBucket));
            });
        }

        if (Array.isArray(record.totals_by_language_model)) {
            record.totals_by_language_model.forEach(languageModelBucket => {
                const language = languageModelBucket.language || 'unknown';
                const model = languageModelBucket.model || 'unknown';
                accumulateNestedMetric(languageModelMatrix, language, model, pickPrimaryUsageValue(languageModelBucket));
            });
        }

        if (record.totals_by_cli) {
            hasCli = true;
            const cliRow = getOrCreateMapValue(cliDayMap, day, () => createCliDayRow(day));
            accumulateCliRow(cliRow, record.totals_by_cli);
        }

        if (record.pull_requests) {
            hasPullRequests = true;
            const prRow = getOrCreateMapValue(pullRequestDayMap, day, () => createPullRequestDayRow(day));
            accumulatePullRequestRow(prRow, record.pull_requests);
        }
    });

    userRecords.forEach(record => {
        const userKey = record.user_id || record.user_login;
        const day = record.day || '';
        if (!userKey || !day) return;

        uniqueUsers.add(userKey);
        getOrCreateMapValue(dayUserSets, day, () => new Set()).add(userKey);
        getOrCreateMapValue(weekUserSets, getWeekStart(day), () => new Set()).add(userKey);

        if (record.used_chat || recordContainsChatActivity(record)) {
            chatUsers.add(userKey);
        }
        if (record.used_agent || recordContainsAgentActivity(record)) {
            agentUsers.add(userKey);
        }
        if (record.used_cli || record.totals_by_cli) {
            cliUsers.add(userKey);
            getOrCreateMapValue(dayCliUserSets, day, () => new Set()).add(userKey);
            hasCli = true;
        }
        if (record.used_copilot_code_review_active) {
            reviewActiveUsers.add(userKey);
            getOrCreateMapValue(dayReviewActiveSets, day, () => new Set()).add(userKey);
            hasCodeReview = true;
        }
        if (record.used_copilot_code_review_passive) {
            reviewPassiveUsers.add(userKey);
            getOrCreateMapValue(dayReviewPassiveSets, day, () => new Set()).add(userKey);
            hasCodeReview = true;
        }
    });

    const daySet = new Set([
        ...Array.from(dayOverviewMap.keys()),
        ...Array.from(dayUserSets.keys()),
        ...Array.from(dayCliUserSets.keys()),
        ...Array.from(dayReviewActiveSets.keys()),
        ...Array.from(dayReviewPassiveSets.keys())
    ]);
    const days = Array.from(daySet).filter(Boolean).sort();

    days.forEach(day => {
        const row = getOrCreateMapValue(dayOverviewMap, day, () => createDailyOverviewRow(day));
        if (!row.daily_active_users) row.daily_active_users = (dayUserSets.get(day) || new Set()).size;
        if (!row.daily_active_cli_users) row.daily_active_cli_users = (dayCliUserSets.get(day) || new Set()).size;
        row.code_review_active_users = (dayReviewActiveSets.get(day) || new Set()).size;
        row.code_review_passive_users = (dayReviewPassiveSets.get(day) || new Set()).size;
        if (row.code_review_active_users || row.code_review_passive_users) hasCodeReview = true;
    });

    const dailyRows = days.map(day => dayOverviewMap.get(day));
    const weeklyRows = dailyRows.some(row => toNum(row.weekly_active_users) > 0)
        ? dailyRows.filter(row => toNum(row.weekly_active_users) > 0).map(row => ({ label: row.day, value: toNum(row.weekly_active_users) }))
        : Array.from(weekUserSets.entries())
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([weekStart, users]) => ({ label: weekStart, value: users.size }));

    const featureRows = sortMetricRows(featureTotals);
    const ideRows = sortMetricRows(ideTotals);
    const languageRows = sortMetricRows(languageTotals);
    const modelRows = sortMetricRows(modelTotals);
    const cliDayRows = Array.from(cliDayMap.values()).sort((a, b) => a.day.localeCompare(b.day));
    const pullRequestDayRows = Array.from(pullRequestDayMap.values()).sort((a, b) => a.day.localeCompare(b.day));
    const codeReviewDayRows = dailyRows
        .filter(row => row.code_review_active_users || row.code_review_passive_users)
        .map(row => ({
            day: row.day,
            active_users: row.code_review_active_users,
            passive_users: row.code_review_passive_users
        }));

    const latestDailyRow = dailyRows[dailyRows.length - 1] || createDailyOverviewRow('');
    const latestDailyActiveUsers = toNum(latestDailyRow.daily_active_users);
    const latestWeeklyActiveUsers = toNum(latestDailyRow.weekly_active_users) || (weeklyRows[weeklyRows.length - 1]?.value || 0);
    const latestMonthlyActiveUsers = toNum(latestDailyRow.monthly_active_users) || uniqueUsers.size || latestDailyActiveUsers;
    const latestMonthlyChatUsers = toNum(latestDailyRow.monthly_active_chat_users) || chatUsers.size;
    const latestMonthlyAgentUsers = toNum(latestDailyRow.monthly_active_agent_users) || agentUsers.size;
    const latestDailyCliUsers = toNum(latestDailyRow.daily_active_cli_users);
    const acceptanceRate = totalGenerations ? (totalAcceptances / totalGenerations) * 100 : 0;

    const loc = computeLocAggregations(overallRecords);
    const agentLocBucket = loc?.byFeature.get('agent_edit') || { added: 0, deleted: 0 };
    const totalLinesChanged = (loc?.totalAdded || 0) + (loc?.totalDeleted || 0);
    const agentLinesChanged = (agentLocBucket.added || 0) + (agentLocBucket.deleted || 0);
    const avgChatRequestsPerActiveUser = latestMonthlyActiveUsers ? (totalInteractions / latestMonthlyActiveUsers) : 0;
    const agentContributionPct = totalLinesChanged ? (agentLinesChanged / totalLinesChanged) * 100 : 0;
    const averageAgentDeletedPerActiveUser = latestMonthlyActiveUsers ? ((agentLocBucket.deleted || 0) / latestMonthlyActiveUsers) : 0;
    const mostUsedChatModel = modelRows[0]?.key || 'n/a';
    const pullRequestTotals = pullRequestDayRows.reduce((acc, row) => ({
        total_created: acc.total_created + toNum(row.total_created),
        total_reviewed: acc.total_reviewed + toNum(row.total_reviewed),
        total_merged: acc.total_merged + toNum(row.total_merged),
        total_suggestions: acc.total_suggestions + toNum(row.total_suggestions),
        total_applied_suggestions: acc.total_applied_suggestions + toNum(row.total_applied_suggestions),
        total_copilot_suggestions: acc.total_copilot_suggestions + toNum(row.total_copilot_suggestions),
        total_copilot_applied_suggestions: acc.total_copilot_applied_suggestions + toNum(row.total_copilot_applied_suggestions)
    }), {
        total_created: 0,
        total_reviewed: 0,
        total_merged: 0,
        total_suggestions: 0,
        total_applied_suggestions: 0,
        total_copilot_suggestions: 0,
        total_copilot_applied_suggestions: 0
    });
    const cliTotals = cliDayRows.reduce((acc, row) => ({
        request_count: acc.request_count + toNum(row.request_count),
        prompt_count: acc.prompt_count + toNum(row.prompt_count),
        session_count: acc.session_count + toNum(row.session_count),
        prompt_tokens_sum: acc.prompt_tokens_sum + toNum(row.prompt_tokens_sum),
        output_tokens_sum: acc.output_tokens_sum + toNum(row.output_tokens_sum)
    }), {
        request_count: 0,
        prompt_count: 0,
        session_count: 0,
        prompt_tokens_sum: 0,
        output_tokens_sum: 0
    });

    return {
        data,
        userRecords,
        overallRecords,
        userRows: aggregateUserUsage(userRecords),
        meta: {
            recordCount: data.length,
            userRecordCount: userRecords.length,
            aggregateRecordCount: aggregateRecords.length,
            hasUserRecords: userRecords.length > 0,
            hasAggregateRecords: aggregateRecords.length > 0,
            hasCli,
            hasPullRequests,
            hasCodeReview,
            days,
            latestDay: days[days.length - 1] || '',
            earliestDay: days[0] || ''
        },
        totals: {
            totalInteractions,
            totalGenerations,
            totalAcceptances,
            acceptanceRate,
            latestDailyActiveUsers,
            latestWeeklyActiveUsers,
            latestMonthlyActiveUsers,
            latestMonthlyChatUsers,
            latestMonthlyAgentUsers,
            latestDailyCliUsers,
            avgChatRequestsPerActiveUser,
            mostUsedChatModel,
            totalLinesChanged,
            agentLinesChanged,
            agentContributionPct,
            averageAgentDeletedPerActiveUser,
            uniqueUsers: uniqueUsers.size,
            chatUsers: chatUsers.size,
            agentUsers: agentUsers.size,
            cliUsers: cliUsers.size,
            reviewActiveUsers: reviewActiveUsers.size,
            reviewPassiveUsers: reviewPassiveUsers.size,
            pullRequests: pullRequestTotals,
            cli: cliTotals
        },
        breakdowns: {
            dailyRows,
            weeklyRows,
            featureRows,
            ideRows,
            languageRows,
            modelRows,
            cliDayRows,
            pullRequestDayRows,
            codeReviewDayRows
        },
        matrices: {
            dayLanguageMap,
            dayModelMap,
            modelFeatureMatrix,
            languageModelMatrix
        },
        loc
    };
}

function renderCharts(model) {
    const chartsContainer = document.getElementById('chartsContainer');
    if (!chartsContainer) return;
    chartsContainer.innerHTML = '';

    const usageSectionHasCharts = Boolean(
        model.breakdowns.dailyRows.length ||
        model.breakdowns.weeklyRows.length ||
        model.breakdowns.featureRows.length ||
        model.breakdowns.languageRows.length ||
        model.breakdowns.modelRows.length ||
        model.breakdowns.ideRows.length ||
        model.userRows.length
    );
    if (usageSectionHasCharts) {
        appendChartSectionHeading('Adoption, usage, and chat', 'Current usage views follow the latest Copilot dashboard metrics and filter with the loaded export.');
    }

    const chatModeRows = buildTopRows(
        model.breakdowns.featureRows.filter(row => CHAT_MODE_FEATURES.includes(row.key)),
        row => row.interactions || row.generations,
        8
    );
    if (chatModeRows.length) {
        createChart('Requests by Chat Mode', 'column', chatModeRows.map(row => formatFeatureName(row.key)), chatModeRows.map(row => row.interactions || row.generations));
    }

    const modelUsageRows = buildTopRows(model.breakdowns.modelRows, row => metricPrimaryValue(row), 8);
    if (modelUsageRows.length) {
        createChart('Model Usage', 'doughnut', modelUsageRows.map(row => row.key), modelUsageRows.map(row => metricPrimaryValue(row)));
    }

    const modelDaySeries = buildSeriesFromDayMatrix(model.matrices.dayModelMap, 8);
    if (modelDaySeries.categories.length) {
        createStackedChart('Model Usage per Day', modelDaySeries.categories, modelDaySeries.series, 'area');
    }

    const modelFeatureRows = buildMatrixRows(model.matrices.modelFeatureMatrix, 8);
    if (modelFeatureRows.outerKeys.length && modelFeatureRows.innerKeys.length) {
        const series = modelFeatureRows.innerKeys.map(modelKey => ({
            name: modelKey,
            data: modelFeatureRows.outerKeys.map(featureKey => model.matrices.modelFeatureMatrix.get(featureKey)?.get(modelKey) || 0)
        }));
        createStackedChart('Model Usage per Chat Mode', modelFeatureRows.outerKeys.map(formatFeatureName), series, 'column');
    }

    const languageUsageRows = buildTopRows(model.breakdowns.languageRows, row => metricPrimaryValue(row), 8);
    if (languageUsageRows.length) {
        createChart('Language Usage', 'pie', languageUsageRows.map(row => row.key), languageUsageRows.map(row => metricPrimaryValue(row)));
    }

    const languageDaySeries = buildSeriesFromDayMatrix(model.matrices.dayLanguageMap, 8);
    if (languageDaySeries.categories.length) {
        createStackedChart('Language Usage per Day', languageDaySeries.categories, languageDaySeries.series, 'area');
    }

    const languageModelRows = buildHeatmapRows(model.matrices.languageModelMatrix, 24);
    if (languageModelRows.xCategories.length && languageModelRows.yCategories.length) {
        createHeatmap('Language Usage by Model', languageModelRows.xCategories, languageModelRows.yCategories, languageModelRows.points, 'Usage');
    }

    if (model.breakdowns.ideRows.length) {
        const ideRows = buildTopRows(model.breakdowns.ideRows, row => metricPrimaryValue(row), 8);
        createChart('IDE Usage', 'doughnut', ideRows.map(row => row.key), ideRows.map(row => metricPrimaryValue(row)));
    }

    if (model.breakdowns.dailyRows.length) {
        createChart('Daily Active Users', 'line', model.breakdowns.dailyRows.map(row => row.day), model.breakdowns.dailyRows.map(row => row.daily_active_users));
    }

    if (model.breakdowns.weeklyRows.length) {
        createChart('Weekly Active Users', 'line', model.breakdowns.weeklyRows.map(row => row.label), model.breakdowns.weeklyRows.map(row => row.value));
    }

    if (model.userRows.length) {
        const topInteractionUsers = buildTopRows(model.userRows, row => row.interactions, 10);
        if (topInteractionUsers.length) {
            createChart('Top Users by Interaction Count', 'bar', topInteractionUsers.map(row => row.user_login), topInteractionUsers.map(row => row.interactions));
        }

        const topCompletionUsers = buildTopRows(model.userRows, row => row.completions, 10);
        if (topCompletionUsers.length) {
            createGroupedBarChart('Completions vs Acceptances (Top Users)', topCompletionUsers.map(row => row.user_login), [
                { label: 'Completions', data: topCompletionUsers.map(row => row.completions) },
                { label: 'Acceptances', data: topCompletionUsers.map(row => row.acceptances) }
            ]);
        }

        const topRateUsers = buildTopRows(model.userRows.filter(row => row.completions > 0), row => row.acceptance_rate, 10);
        if (topRateUsers.length) {
            createChart('Acceptance Rate % (Top Users)', 'bar', topRateUsers.map(row => row.user_login), topRateUsers.map(row => +row.acceptance_rate.toFixed(1)));
        }
    }

    const codeGenerationSectionHasCharts = Boolean(
        model.loc ||
        model.breakdowns.languageRows.some(row => row.user_loc_changed || row.agent_loc_changed) ||
        model.breakdowns.modelRows.some(row => row.user_loc_changed || row.agent_loc_changed)
    );
    if (codeGenerationSectionHasCharts) {
        appendChartSectionHeading('Code generation', 'These views separate suggestion metrics from actual code changes, including user-initiated versus agent-initiated edits.');
    }

    if (model.breakdowns.dailyRows.length) {
        createStackedChart('Daily Lines Added and Deleted', model.breakdowns.dailyRows.map(row => row.day), [
            { name: 'LoC Added', data: model.breakdowns.dailyRows.map(row => row.loc_added) },
            { name: 'LoC Deleted', data: model.breakdowns.dailyRows.map(row => row.loc_deleted) }
        ], 'area');
    }

    if (model.loc) {
        const featureKeys = Array.from(model.loc.byFeature.keys());
        if (featureKeys.length) {
            createStackedChart('LoC by Feature (Suggested vs Changed)', featureKeys.map(formatFeatureName), [
                { name: 'LoC Suggested to Add', data: featureKeys.map(key => model.loc.byFeature.get(key).suggestAdd || 0) },
                { name: 'LoC Added', data: featureKeys.map(key => model.loc.byFeature.get(key).added || 0) },
                { name: 'LoC Deleted', data: featureKeys.map(key => model.loc.byFeature.get(key).deleted || 0) }
            ], 'column');

            const suggestedDelete = featureKeys.map(key => model.loc.byFeature.get(key).suggestDel || 0);
            if (suggestedDelete.some(value => value > 0)) {
                createChart('LoC Suggested to Delete by Feature', 'column', featureKeys.map(formatFeatureName), suggestedDelete);
            }
        }
    }

    if (model.loc && model.totals.totalLinesChanged > 0) {
        const agentBucket = model.loc.byFeature.get('agent_edit') || { added: 0, deleted: 0 };
        const userAdded = Math.max(0, (model.loc.totalAdded || 0) - (agentBucket.added || 0));
        const userDeleted = Math.max(0, (model.loc.totalDeleted || 0) - (agentBucket.deleted || 0));
        createGroupedBarChart('User vs Agent Code Changes', ['LoC Added', 'LoC Deleted'], [
            { label: 'User initiated', data: [userAdded, userDeleted] },
            { label: 'Agent initiated', data: [agentBucket.added || 0, agentBucket.deleted || 0] }
        ]);
    }

    const userLocByModel = buildTopRows(model.breakdowns.modelRows.filter(row => row.user_loc_changed > 0), row => row.user_loc_changed, 10);
    if (userLocByModel.length) {
        createChart('User-Initiated Code Changes per Model', 'bar', userLocByModel.map(row => row.key), userLocByModel.map(row => row.user_loc_changed));
    }

    const agentLocByModel = buildTopRows(model.breakdowns.modelRows.filter(row => row.agent_loc_changed > 0), row => row.agent_loc_changed, 10);
    if (agentLocByModel.length) {
        createChart('Agent-Initiated Code Changes per Model', 'bar', agentLocByModel.map(row => row.key), agentLocByModel.map(row => row.agent_loc_changed));
    }

    const userLocByLanguage = buildTopRows(model.breakdowns.languageRows.filter(row => row.user_loc_changed > 0), row => row.user_loc_changed, 10);
    if (userLocByLanguage.length) {
        createChart('User-Initiated Code Changes per Language', 'bar', userLocByLanguage.map(row => row.key), userLocByLanguage.map(row => row.user_loc_changed));
    }

    const agentLocByLanguage = buildTopRows(model.breakdowns.languageRows.filter(row => row.agent_loc_changed > 0), row => row.agent_loc_changed, 10);
    if (agentLocByLanguage.length) {
        createChart('Agent-Initiated Code Changes per Language', 'bar', agentLocByLanguage.map(row => row.key), agentLocByLanguage.map(row => row.agent_loc_changed));
    }

    const opsSectionHasCharts = Boolean(
        model.breakdowns.cliDayRows.length ||
        model.breakdowns.codeReviewDayRows.length ||
        model.breakdowns.pullRequestDayRows.length ||
        model.totals.latestDailyCliUsers
    );
    if (opsSectionHasCharts) {
        appendChartSectionHeading('CLI, code review, and pull requests', 'New operational views surface CLI activity, code review adoption, and pull request lifecycle metrics when present in the export.');
    }

    if (model.breakdowns.dailyRows.some(row => row.daily_active_cli_users > 0)) {
        createChart('Daily Active CLI Users', 'line', model.breakdowns.dailyRows.map(row => row.day), model.breakdowns.dailyRows.map(row => row.daily_active_cli_users));
    }

    if (model.breakdowns.cliDayRows.length) {
        createStackedChart('Copilot CLI Requests, Prompts, and Sessions', model.breakdowns.cliDayRows.map(row => row.day), [
            { name: 'Requests', data: model.breakdowns.cliDayRows.map(row => row.request_count) },
            { name: 'Prompts', data: model.breakdowns.cliDayRows.map(row => row.prompt_count) },
            { name: 'Sessions', data: model.breakdowns.cliDayRows.map(row => row.session_count) }
        ], 'column', null);
    }

    if (model.breakdowns.codeReviewDayRows.length) {
        createStackedChart('Copilot Code Review Activity', model.breakdowns.codeReviewDayRows.map(row => row.day), [
            { name: 'Active users', data: model.breakdowns.codeReviewDayRows.map(row => row.active_users) },
            { name: 'Passive users', data: model.breakdowns.codeReviewDayRows.map(row => row.passive_users) }
        ], 'column', null);
    }

    if (model.breakdowns.pullRequestDayRows.length) {
        createStackedChart('Pull Request Activity', model.breakdowns.pullRequestDayRows.map(row => row.day), [
            { name: 'Created', data: model.breakdowns.pullRequestDayRows.map(row => row.total_created) },
            { name: 'Reviewed', data: model.breakdowns.pullRequestDayRows.map(row => row.total_reviewed) },
            { name: 'Merged', data: model.breakdowns.pullRequestDayRows.map(row => row.total_merged) }
        ], 'column', null);

        const suggestionRows = model.breakdowns.pullRequestDayRows.filter(row =>
            row.total_suggestions ||
            row.total_applied_suggestions ||
            row.total_copilot_suggestions ||
            row.total_copilot_applied_suggestions
        );
        if (suggestionRows.length) {
            createStackedChart('Pull Request Suggestions', suggestionRows.map(row => row.day), [
                { name: 'Suggestions', data: suggestionRows.map(row => row.total_suggestions) },
                { name: 'Applied suggestions', data: suggestionRows.map(row => row.total_applied_suggestions) },
                { name: 'Copilot suggestions', data: suggestionRows.map(row => row.total_copilot_suggestions) },
                { name: 'Copilot applied suggestions', data: suggestionRows.map(row => row.total_copilot_applied_suggestions) }
            ], 'column', null);
        }
    }
}

function renderReferenceTables(model) {
    const section = document.getElementById('tablesSection');
    const container = document.getElementById('tablesContainer');
    if (!section || !container) return;

    const blocks = [];
    const dailyRows = [...model.breakdowns.dailyRows].reverse();
    if (dailyRows.length) {
        blocks.push(renderTableBlock(
            'Daily overview',
            'Latest daily, weekly, and monthly active-user values plus code generation and LoC totals for each day.',
            [
                { key: 'day', label: 'Day' },
                { key: 'daily_active_users', label: 'Daily active', type: 'number' },
                { key: 'weekly_active_users', label: 'Weekly active', type: 'number' },
                { key: 'monthly_active_users', label: 'Monthly active', type: 'number' },
                { key: 'monthly_active_chat_users', label: 'Monthly chat', type: 'number' },
                { key: 'monthly_active_agent_users', label: 'Monthly agent', type: 'number' },
                { key: 'daily_active_cli_users', label: 'Daily CLI active', type: 'number' },
                { key: 'interactions', label: 'Interactions', type: 'number' },
                { key: 'code_generations', label: 'Generations', type: 'number' },
                { key: 'acceptances', label: 'Acceptances', type: 'number' },
                { key: 'loc_suggested_add', label: 'LoC suggested add', type: 'number' },
                { key: 'loc_added', label: 'LoC added', type: 'number' },
                { key: 'loc_deleted', label: 'LoC deleted', type: 'number' }
            ],
            dailyRows,
            true
        ));
    }

    if (model.breakdowns.featureRows.length) {
        blocks.push(renderTableBlock(
            'Feature breakdown',
            'Copilot feature totals, including chat modes and the agent-edit feature used for direct file changes.',
            [
                { key: 'key', label: 'Feature', render: value => escapeHtml(formatFeatureName(value)) },
                { key: 'interactions', label: 'Interactions', type: 'number' },
                { key: 'generations', label: 'Generations', type: 'number' },
                { key: 'acceptances', label: 'Acceptances', type: 'number' },
                { key: 'loc_suggested_add', label: 'LoC suggested add', type: 'number' },
                { key: 'loc_suggested_delete', label: 'LoC suggested delete', type: 'number' },
                { key: 'loc_added', label: 'LoC added', type: 'number' },
                { key: 'loc_deleted', label: 'LoC deleted', type: 'number' }
            ],
            model.breakdowns.featureRows
        ));
    }

    if (model.breakdowns.languageRows.length) {
        blocks.push(renderTableBlock(
            'Language breakdown',
            'Usage, generation, acceptance, and user-versus-agent code change totals grouped by language.',
            [
                { key: 'key', label: 'Language' },
                { key: 'interactions', label: 'Interactions', type: 'number' },
                { key: 'generations', label: 'Generations', type: 'number' },
                { key: 'acceptances', label: 'Acceptances', type: 'number' },
                { key: 'user_loc_changed', label: 'User-changed LoC', type: 'number' },
                { key: 'agent_loc_changed', label: 'Agent-changed LoC', type: 'number' },
                { key: 'loc_added', label: 'LoC added', type: 'number' },
                { key: 'loc_deleted', label: 'LoC deleted', type: 'number' }
            ],
            model.breakdowns.languageRows
        ));
    }

    if (model.breakdowns.modelRows.length) {
        blocks.push(renderTableBlock(
            'Model breakdown',
            'Usage, code generation, and code change totals grouped by model.',
            [
                { key: 'key', label: 'Model' },
                { key: 'interactions', label: 'Interactions', type: 'number' },
                { key: 'generations', label: 'Generations', type: 'number' },
                { key: 'acceptances', label: 'Acceptances', type: 'number' },
                { key: 'user_loc_changed', label: 'User-changed LoC', type: 'number' },
                { key: 'agent_loc_changed', label: 'Agent-changed LoC', type: 'number' },
                { key: 'loc_added', label: 'LoC added', type: 'number' },
                { key: 'loc_deleted', label: 'LoC deleted', type: 'number' }
            ],
            model.breakdowns.modelRows
        ));
    }

    if (model.breakdowns.ideRows.length) {
        blocks.push(renderTableBlock(
            'IDE breakdown',
            'IDE usage with the latest observed IDE and plugin versions available in the export.',
            [
                { key: 'key', label: 'IDE' },
                { key: 'interactions', label: 'Interactions', type: 'number' },
                { key: 'generations', label: 'Generations', type: 'number' },
                { key: 'acceptances', label: 'Acceptances', type: 'number' },
                { key: 'latest_ide_version', label: 'Latest IDE version' },
                { key: 'latest_plugin_version', label: 'Latest plugin version' }
            ],
            model.breakdowns.ideRows
        ));
    }

    if (model.breakdowns.cliDayRows.length) {
        blocks.push(renderTableBlock(
            'CLI activity',
            'Daily Copilot CLI usage, requests, sessions, tokens, and the latest detected CLI version.',
            [
                { key: 'day', label: 'Day' },
                { key: 'request_count', label: 'Requests', type: 'number' },
                { key: 'prompt_count', label: 'Prompts', type: 'number' },
                { key: 'session_count', label: 'Sessions', type: 'number' },
                { key: 'prompt_tokens_sum', label: 'Prompt tokens', type: 'number' },
                { key: 'output_tokens_sum', label: 'Output tokens', type: 'number' },
                { key: 'avg_tokens_per_request', label: 'Avg tokens/request', type: 'decimal' },
                { key: 'latest_cli_version', label: 'Latest CLI version' }
            ],
            [...model.breakdowns.cliDayRows].reverse()
        ));
    }

    if (model.breakdowns.codeReviewDayRows.length) {
        blocks.push(renderTableBlock(
            'Code review activity',
            'Daily counts of users who actively engaged with Copilot code review or had passive Copilot review activity.',
            [
                { key: 'day', label: 'Day' },
                { key: 'active_users', label: 'Active users', type: 'number' },
                { key: 'passive_users', label: 'Passive users', type: 'number' }
            ],
            [...model.breakdowns.codeReviewDayRows].reverse()
        ));
    }

    if (model.breakdowns.pullRequestDayRows.length) {
        blocks.push(renderTableBlock(
            'Pull request activity',
            'Daily pull request lifecycle metrics, including Copilot-authored and Copilot-reviewed activity when present.',
            [
                { key: 'day', label: 'Day' },
                { key: 'total_created', label: 'Created', type: 'number' },
                { key: 'total_reviewed', label: 'Reviewed', type: 'number' },
                { key: 'total_merged', label: 'Merged', type: 'number' },
                { key: 'median_minutes_to_merge', label: 'Median minutes to merge', type: 'decimal' },
                { key: 'total_suggestions', label: 'Suggestions', type: 'number' },
                { key: 'total_applied_suggestions', label: 'Applied suggestions', type: 'number' },
                { key: 'total_created_by_copilot', label: 'Created by Copilot', type: 'number' },
                { key: 'total_reviewed_by_copilot', label: 'Reviewed by Copilot', type: 'number' },
                { key: 'total_merged_created_by_copilot', label: 'Merged created by Copilot', type: 'number' },
                { key: 'total_merged_reviewed_by_copilot', label: 'Merged reviewed by Copilot', type: 'number' },
                { key: 'total_copilot_suggestions', label: 'Copilot suggestions', type: 'number' },
                { key: 'total_copilot_applied_suggestions', label: 'Copilot applied suggestions', type: 'number' }
            ],
            [...model.breakdowns.pullRequestDayRows].reverse()
        ));
    }

    if (!blocks.length) {
        section.hidden = true;
        container.innerHTML = '';
        return;
    }

    section.hidden = false;
    container.innerHTML = blocks.join('');
}

function renderTableBlock(title, note, columns, rows, open = false) {
    return `<details class="data-table-block"${open ? ' open' : ''}>
        <summary>
            <div>
                <h3 class="data-table-title">${escapeHtml(title)}</h3>
                <p class="data-table-note">${escapeHtml(note)}</p>
            </div>
        </summary>
        <div class="data-table-content">
            <div class="data-table-scroll">
                ${renderStaticTable(title, note, columns, rows)}
            </div>
        </div>
    </details>`;
}

function renderStaticTable(title, note, columns, rows) {
    const caption = `<caption class="visually-hidden">${escapeHtml(title)}. ${escapeHtml(note)}</caption>`;
    const head = columns.map(column => `<th scope="col">${escapeHtml(column.label)}</th>`).join('');
    const body = rows.map(row => `<tr>${columns.map((column, index) => index === 0
        ? `<th scope="row">${formatTableCell(row[column.key], column, row)}</th>`
        : `<td>${formatTableCell(row[column.key], column, row)}</td>`).join('')}</tr>`).join('');
    return `<table class="usage-table">${caption}<thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

function formatTableCell(value, column, row) {
    if (column.render) return column.render(value, row);
    if (value === null || value === undefined || value === '') return '&mdash;';
    if (column.type === 'number') return escapeHtml(Number(value).toLocaleString());
    if (column.type === 'decimal') return escapeHtml(Number(value).toFixed(1));
    return escapeHtml(value);
}

function appendChartSectionHeading(title, note) {
    const chartsContainer = document.getElementById('chartsContainer');
    if (!chartsContainer) return;
    const wrapper = document.createElement('div');
    wrapper.className = 'charts-section-heading';
    wrapper.innerHTML = `<h3>${escapeHtml(title)}</h3><p>${escapeHtml(note)}</p>`;
    chartsContainer.appendChild(wrapper);
}

function buildTopRows(rows, getValue, limit = 10) {
    return [...rows]
        .filter(row => getValue(row) > 0)
        .sort((a, b) => getValue(b) - getValue(a))
        .slice(0, limit);
}

function buildSeriesFromDayMatrix(dayMap, limit = 8) {
    const categories = Array.from(dayMap.keys()).sort();
    const totals = {};
    dayMap.forEach(dayEntries => {
        dayEntries.forEach((value, key) => {
            totals[key] = (totals[key] || 0) + value;
        });
    });
    const topKeys = Object.entries(totals).sort((a, b) => b[1] - a[1]).slice(0, limit).map(entry => entry[0]);
    if (!categories.length || !topKeys.length) return { categories: [], series: [] };
    const series = topKeys.map(key => ({
        name: key,
        data: categories.map(day => dayMap.get(day)?.get(key) || 0)
    }));
    const otherData = categories.map(day => {
        let sum = 0;
        (dayMap.get(day) || new Map()).forEach((value, key) => {
            if (!topKeys.includes(key)) sum += value;
        });
        return sum;
    });
    if (otherData.some(value => value > 0)) {
        series.push({ name: 'Other', data: otherData });
    }
    return { categories, series };
}

function buildMatrixRows(matrixMap, limitRows = 8) {
    const outerKeys = Array.from(matrixMap.keys())
        .sort((a, b) => sumNestedMap(matrixMap.get(b)) - sumNestedMap(matrixMap.get(a)))
        .slice(0, limitRows);
    const innerKeySet = new Set();
    outerKeys.forEach(key => {
        (matrixMap.get(key) || new Map()).forEach((_, innerKey) => innerKeySet.add(innerKey));
    });
    const innerKeys = Array.from(innerKeySet);
    return { outerKeys, innerKeys };
}

function buildHeatmapRows(matrixMap, limitRows = 24) {
    const yKeys = Array.from(matrixMap.keys())
        .sort((a, b) => sumNestedMap(matrixMap.get(b)) - sumNestedMap(matrixMap.get(a)))
        .slice(0, limitRows);
    const xKeySet = new Set();
    yKeys.forEach(key => {
        (matrixMap.get(key) || new Map()).forEach((_, innerKey) => xKeySet.add(innerKey));
    });
    const xCategories = Array.from(xKeySet);
    const points = [];
    yKeys.forEach((yKey, yIndex) => {
        xCategories.forEach((xKey, xIndex) => {
            points.push([xIndex, yIndex, matrixMap.get(yKey)?.get(xKey) || 0]);
        });
    });
    return { xCategories, yCategories: yKeys, points };
}

function isUserLevelRecord(record) {
    return Boolean(record && (record.__record_scope === 'user' || record.user_id || record.user_login));
}

function createDailyOverviewRow(day) {
    return {
        day,
        daily_active_users: 0,
        weekly_active_users: 0,
        monthly_active_users: 0,
        monthly_active_chat_users: 0,
        monthly_active_agent_users: 0,
        daily_active_cli_users: 0,
        code_review_active_users: 0,
        code_review_passive_users: 0,
        interactions: 0,
        code_generations: 0,
        acceptances: 0,
        loc_suggested_add: 0,
        loc_suggested_delete: 0,
        loc_added: 0,
        loc_deleted: 0
    };
}

function createCliDayRow(day) {
    return {
        day,
        request_count: 0,
        prompt_count: 0,
        session_count: 0,
        prompt_tokens_sum: 0,
        output_tokens_sum: 0,
        avg_tokens_per_request: 0,
        latest_cli_version: ''
    };
}

function createPullRequestDayRow(day) {
    return {
        day,
        total_created: 0,
        total_reviewed: 0,
        total_merged: 0,
        median_minutes_to_merge: 0,
        total_suggestions: 0,
        total_applied_suggestions: 0,
        total_created_by_copilot: 0,
        total_reviewed_by_copilot: 0,
        total_merged_created_by_copilot: 0,
        total_merged_reviewed_by_copilot: 0,
        median_minutes_to_merge_copilot_authored: 0,
        total_copilot_suggestions: 0,
        total_copilot_applied_suggestions: 0
    };
}

function accumulateMetricBucket(map, key, bucket) {
    const row = getOrCreateMapValue(map, key, () => ({
        key,
        interactions: 0,
        generations: 0,
        acceptances: 0,
        loc_suggested_add: 0,
        loc_suggested_delete: 0,
        loc_added: 0,
        loc_deleted: 0,
        user_loc_changed: 0,
        agent_loc_changed: 0,
        latest_ide_version: '',
        latest_plugin_version: ''
    }));

    row.interactions += toNum(bucket.user_initiated_interaction_count);
    row.generations += toNum(bucket.code_generation_activity_count);
    row.acceptances += toNum(bucket.code_acceptance_activity_count);
    row.loc_suggested_add += toNum(bucket.loc_suggested_to_add_sum);
    row.loc_suggested_delete += toNum(bucket.loc_suggested_to_delete_sum);
    row.loc_added += toNum(bucket.loc_added_sum);
    row.loc_deleted += toNum(bucket.loc_deleted_sum);

    const changedTotal = toNum(bucket.loc_added_sum) + toNum(bucket.loc_deleted_sum);
    if (isAgentFeature(bucket.feature)) {
        row.agent_loc_changed += changedTotal;
    } else {
        row.user_loc_changed += changedTotal;
    }
    return row;
}

function accumulateNestedMetric(map, outerKey, innerKey, value) {
    if (!outerKey || !innerKey || !value) return;
    const row = getOrCreateMapValue(map, outerKey, () => new Map());
    row.set(innerKey, (row.get(innerKey) || 0) + value);
}

function accumulateCliRow(row, cliBucket) {
    row.request_count += toNum(cliBucket.request_count);
    row.prompt_count += toNum(cliBucket.prompt_count);
    row.session_count += toNum(cliBucket.session_count);
    row.prompt_tokens_sum += toNum(cliBucket.token_usage?.prompt_tokens_sum);
    row.output_tokens_sum += toNum(cliBucket.token_usage?.output_tokens_sum);
    row.avg_tokens_per_request = row.request_count
        ? (row.prompt_tokens_sum + row.output_tokens_sum) / row.request_count
        : 0;
    updateLatestVersion(row, 'latest_cli_version', cliBucket.last_known_cli_version, info => info.cli_version || '');
}

function accumulatePullRequestRow(row, pullRequests) {
    Object.keys(row).forEach(key => {
        if (key === 'day') return;
        row[key] += toNum(pullRequests[key]);
    });
}

function updateLatestVersion(target, key, versionInfo, formatter) {
    if (!versionInfo || typeof versionInfo !== 'object') return;
    const sampleKey = `__${key}_sampled_at`;
    const sampledAt = versionInfo.sampled_at || '';
    if (!target[sampleKey] || sampledAt >= target[sampleKey]) {
        target[sampleKey] = sampledAt;
        target[key] = formatter(versionInfo);
    }
}

function metricPrimaryValue(row) {
    return row.interactions || row.generations || row.acceptances || row.user_loc_changed || row.agent_loc_changed || row.loc_added || 0;
}

function sortMetricRows(map) {
    return Array.from(map.values()).sort((a, b) => metricPrimaryValue(b) - metricPrimaryValue(a));
}

function pickPrimaryUsageValue(bucket) {
    return toNum(bucket.user_initiated_interaction_count) || toNum(bucket.code_generation_activity_count) || toNum(bucket.code_acceptance_activity_count);
}

function recordContainsChatActivity(record) {
    return Array.isArray(record.totals_by_feature)
        && record.totals_by_feature.some(featureBucket => CHAT_FEATURES.has(featureBucket.feature) && pickPrimaryUsageValue(featureBucket) > 0);
}

function recordContainsAgentActivity(record) {
    return Array.isArray(record.totals_by_feature)
        && record.totals_by_feature.some(featureBucket => featureBucket.feature === 'chat_panel_agent_mode' || featureBucket.feature === 'agent_edit');
}

function isAgentFeature(feature) {
    return feature === 'agent_edit';
}

function getOrCreateMapValue(map, key, createValue) {
    if (!map.has(key)) map.set(key, createValue());
    return map.get(key);
}

function getWeekStart(day) {
    const date = new Date(`${day}T00:00:00Z`);
    const offset = (date.getUTCDay() + 6) % 7;
    date.setUTCDate(date.getUTCDate() - offset);
    return date.toISOString().slice(0, 10);
}

function sumNestedMap(map) {
    if (!map) return 0;
    let total = 0;
    map.forEach(value => { total += value; });
    return total;
}

function createChart(title, type, categories, seriesData) {
    const chartsContainer = document.getElementById('chartsContainer');
    const chartContainer = document.createElement('div');
    chartContainer.classList.add('chart-container');
    chartContainer.setAttribute('role','group');
    chartContainer.setAttribute('aria-label', title);
    const div = document.createElement('div');
    // Automatically trim very large categorical bar/column charts to first 40 entries for readability
    if ((type === 'bar' || type === 'column') && categories.length > 60) {
        categories = categories.slice(0, 40);
        seriesData = seriesData.slice(0, 40);
    }
    chartContainer.appendChild(div);
    chartsContainer.appendChild(chartContainer);
    function buildOptions(cats, data) {
        const tokens = getThemeTokens();
        const opts = {
            chart: { type: type === 'doughnut' ? 'pie' : type, backgroundColor: 'transparent', height: 420 },
            title: { text: title, style: { fontSize: '15px' } },
            xAxis: { categories: cats, labels: { style: { fontSize: '12px' } } },
            yAxis: { title: { text: null }, gridLineColor: tokens.chartGrid },
            legend: { itemStyle: { fontSize: '12px' } },
            accessibility: { enabled: true },
            credits: { enabled: false },
            tooltip: { shared: true },
            series: [{ name: title, data: data }]
        };
        if (type === 'pie' || type === 'doughnut') {
            opts.series = [{
                type: 'pie',
                name: title,
                innerSize: type === 'doughnut' ? '55%' : undefined,
                data: cats.map((c,i) => ({ name: c, y: data[i] })),
                dataLabels: {
                    enabled: true,
                    // Show slice name and percentage with one decimal
                    format: '{point.name}: {point.percentage:.1f}%',
                    style: { fontSize: '12px', fontWeight: '500', textOutline: 'none', color: tokens.text }
                }
            }];
            delete opts.xAxis; delete opts.yAxis; opts.tooltip.shared = false;
        } else if (type === 'line') {
            opts.chart.type = 'line';
        }
        return opts;
    }
    Highcharts.chart(div, buildOptions(categories, seriesData));
}

function createGroupedBarChart(title, categories, datasets) {
    const chartsContainer = document.getElementById('chartsContainer');
    const chartContainer = document.createElement('div');
    chartContainer.classList.add('chart-container');
    chartContainer.setAttribute('role', 'group');
    chartContainer.setAttribute('aria-label', title);
    const div = document.createElement('div');
    chartContainer.appendChild(div);
    chartsContainer.appendChild(chartContainer);

    Highcharts.chart(div, {
    chart: { type: 'column', backgroundColor: 'transparent', height: 420 },
        title: { text: title, style: { fontSize: '15px' } },
        xAxis: { categories: categories, crosshair: true },
        yAxis: { min: 0, title: { text: null } },
        tooltip: { shared: true },
        legend: { itemStyle: { fontSize: '12px' } },
        accessibility: { enabled: true },
        credits: { enabled: false },
        plotOptions: { column: { pointPadding: 0.08, borderWidth: 0, groupPadding: 0.12 } },
        series: datasets.map(d => ({ name: d.label, data: d.data }))
    });
}

function createHeatmap(title, xCategories, yCategories, dataPoints, colorAxisTitle) {
    const chartsContainer = document.getElementById('chartsContainer');
    const chartContainer = document.createElement('div');
    chartContainer.classList.add('chart-container');
    chartContainer.setAttribute('role', 'group');
    chartContainer.setAttribute('aria-label', title);
    const div = document.createElement('div');
    chartContainer.appendChild(div);
    chartsContainer.appendChild(chartContainer);
    const maxVal = dataPoints.reduce((m,p)=> Math.max(m,p[2]),0) || 0;
    const tokens = getThemeTokens();
    Highcharts.chart(div, {
    chart: { type: 'heatmap', backgroundColor: 'transparent', height: 480 },
        title: { text: title, style: { fontSize: '15px' } },
        xAxis: { categories: xCategories, labels: { style: { fontSize: '12px' }, rotation: 40 } },
        yAxis: { categories: yCategories, title: null, labels: { style: { fontSize: '12px' } }, reversed: true },
        accessibility: { enabled: true },
        legend: { align: 'right', layout: 'vertical', verticalAlign: 'middle' },
        colorAxis: {
            min: 0,
            max: maxVal,
            stops: [
                [0, tokens.heatStops[0]],
                [0.4, tokens.heatStops[1]],
                [0.7, tokens.heatStops[2]],
                [1, tokens.heatStops[3]]
            ]
        },
        tooltip: { 
            formatter: function() { return `<b>${this.series.name}</b><br/>${Highcharts.numberFormat(this.point.value,0,'.',',')} ${colorAxisTitle}<br/>`;} 
        },
        series: [{
            name: colorAxisTitle,
            borderWidth: 1,
            borderColor: tokens.heatBorder,
            colsize: 1,
            rowsize: 1,
            data: dataPoints,
            dataLabels: { enabled: false }
        }]
    });
}

function createStackedChart(title, categories, series, type='column', stacking='normal') {
    const chartsContainer = document.getElementById('chartsContainer');
    const chartContainer = document.createElement('div');
    chartContainer.classList.add('chart-container');
    chartContainer.setAttribute('role','group');
    chartContainer.setAttribute('aria-label', title);
    const div = document.createElement('div');
    chartContainer.appendChild(div);
    chartsContainer.appendChild(chartContainer);
    Highcharts.chart(div, {
    chart: { type: type === 'area' ? 'area' : 'column', backgroundColor: 'transparent', height: (type === 'area' ? 420 : 440) },
        title: { text: title, style: { fontSize: '15px' } },
        xAxis: { categories, labels: { style: { fontSize: '12px' } } },
        yAxis: { min: 0, title: { text: null } },
        legend: { itemStyle: { fontSize: '12px' } },
        tooltip: { shared: true },
        plotOptions: { 
            series: stacking ? { stacking } : {},
            area: stacking ? { stacking, marker: { enabled: false }, lineWidth: 1 } : { marker: { enabled: false }, lineWidth: 1 }
        },
        accessibility: { enabled: true },
        series
    });
}

// -------- Added: Filters, metrics, theme toggle, status helpers -------- //

function initializeFilters(data) {
    if (!document.getElementById('applyFiltersBtn')) return; // already enhanced
    // Set date bounds
    const days = [...new Set(data.map(r => r.day).filter(Boolean))].sort();
    window.__allDays = days;
    const dateFrom = document.getElementById('dateFrom');
    const dateTo = document.getElementById('dateTo');
    if (days.length) {
        dateFrom.min = days[0];
        dateFrom.max = days[days.length - 1];
        dateTo.min = days[0];
        dateTo.max = days[days.length - 1];
        dateFrom.value = days[0];
        dateTo.value = days[days.length - 1];
    }

    document.getElementById('applyFiltersBtn').onclick = () => {
        applyFilters();
    };
    document.getElementById('resetFiltersBtn').onclick = () => {
    // Reset text search
    const searchEl = document.getElementById('userSearch');
    if (searchEl) searchEl.value = '';
    // Reset date range to full span
    if (days.length) { dateFrom.value = days[0]; dateTo.value = days[days.length - 1]; }
    // Reset members-only filter
    const membersChk = document.getElementById('membersOnlyChk');
    if (membersChk) membersChk.checked = false;
    // Clear active quick-range buttons
    document.querySelectorAll('.range-btn.active').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.range-btn').forEach(btn => btn.setAttribute('aria-pressed','false'));
    // Re-run with original unfiltered dataset
    analyzeData(window.__rawData || []);
    };

    // Auto-apply when members-only checkbox toggled
    const membersChk = document.getElementById('membersOnlyChk');
    if (membersChk) {
        membersChk.addEventListener('change', () => {
            applyFilters();
        });
    }

    enableApplyButton();
    setupQuickRangeButtons();
    // Cache initial dataset for user usage table reuse
    window.__currentFilteredData = data;
}

function applyFilters() {
    const search = document.getElementById('userSearch').value.trim().toLowerCase();
    const from = document.getElementById('dateFrom').value;
    const to = document.getElementById('dateTo').value;
    const membersOnly = document.getElementById('membersOnlyChk')?.checked;
    let filtered = window.__rawData || [];
    if (search) {
        filtered = filtered.filter(r => (r.user_login || '').toLowerCase().includes(search));
    }
    if (from) {
        filtered = filtered.filter(r => !r.day || r.day >= from);
    }
    if (to) {
        filtered = filtered.filter(r => !r.day || r.day <= to);
    }
    if (membersOnly && window.__membersSet) {
        filtered = filtered.filter(r => window.__membersSet.has((r.user_login || '').toLowerCase()));
    }
    analyzeData(filtered);
    // Cache for user usage table & update if visible
    window.__currentFilteredData = filtered;
    const userUsageSection = document.getElementById('userUsageSection');
    if (userUsageSection && !userUsageSection.hidden) {
        buildUserUsageTable(filtered);
    }
}

function enableApplyButton() {
    const btn = document.getElementById('applyFiltersBtn');
    if (btn) btn.disabled = false;
}

function setupQuickRangeButtons() {
    const buttons = document.querySelectorAll('.range-btn');
    if (!buttons.length || !window.__allDays || !window.__allDays.length) return;
    const days = window.__allDays;
    const latest = days[days.length - 1];
    const latestDate = new Date(latest + 'T00:00:00Z');
    buttons.forEach(btn => {
        btn.onclick = () => {
            buttons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            // Update aria-pressed for toggle group
            buttons.forEach(b => b.setAttribute('aria-pressed', b.classList.contains('active') ? 'true' : 'false'));
            const range = btn.getAttribute('data-range');
            const fromEl = document.getElementById('dateFrom');
            const toEl = document.getElementById('dateTo');
            if (!fromEl || !toEl) return;
            if (range === 'all') {
                fromEl.value = days[0];
                toEl.value = days[days.length - 1];
            } else {
                const n = parseInt(range, 10) || 0;
                const fromDate = new Date(latestDate);
                fromDate.setUTCDate(latestDate.getUTCDate() - (n - 1));
                const iso = d => d.toISOString().substring(0,10);
                let fromStr = iso(fromDate);
                if (fromStr < days[0]) fromStr = days[0];
                fromEl.value = fromStr;
                toEl.value = latest;
            }
            applyFilters();
        };
    });
}

function computeSummaryMetrics(model) {
    const container = document.getElementById('summaryMetrics');
    if (!container) return;
    container.innerHTML = '';
    if (!model || !model.meta.recordCount) { container.innerHTML = buildMetricPlaceholders(); return; }

    const totalActiveUsers = model.totals.latestMonthlyActiveUsers || model.totals.uniqueUsers || model.totals.latestDailyActiveUsers;
    const agentAdoptionPct = totalActiveUsers ? ((model.totals.latestMonthlyAgentUsers || model.totals.agentUsers) / totalActiveUsers) * 100 : 0;
    const cards = [
        { label: 'Total Active Users', value: totalActiveUsers.toLocaleString() },
        { label: 'Daily Active Users (Latest)', value: model.totals.latestDailyActiveUsers.toLocaleString() },
        { label: 'Weekly Active Users (Latest)', value: model.totals.latestWeeklyActiveUsers.toLocaleString() },
        { label: 'Interactions', value: model.totals.totalInteractions.toLocaleString() },
        { label: 'Code Generations', value: model.totals.totalGenerations.toLocaleString() },
        { label: 'Acceptances', value: model.totals.totalAcceptances.toLocaleString() },
        { label: 'Acceptance Rate %', value: model.totals.acceptanceRate.toFixed(1) },
        { label: 'Avg Chat Requests / Active User', value: model.totals.avgChatRequestsPerActiveUser.toFixed(1) },
        { label: 'Agent Adoption %', value: agentAdoptionPct.toFixed(1) },
        { label: 'Most Used Chat Model', value: model.totals.mostUsedChatModel },
        { label: 'Distinct Days', value: model.meta.days.length.toLocaleString() }
    ];

    if (model.loc && (model.loc.totalSuggestedAdd > 0 || model.loc.totalAdded > 0 || model.loc.totalDeleted > 0)) {
        cards.splice(7, 0,
            { label: 'AI Lines Changed', value: model.totals.totalLinesChanged.toLocaleString() },
            { label: 'Agent Contribution %', value: model.totals.agentContributionPct.toFixed(1) },
            { label: 'Avg Agent Lines Deleted / Active User', value: model.totals.averageAgentDeletedPerActiveUser.toFixed(1) },
            { label: 'LoC Suggested (Add)', value: model.loc.totalSuggestedAdd.toLocaleString() },
            { label: 'LoC Added', value: model.loc.totalAdded.toLocaleString() },
            { label: 'LoC Deleted', value: model.loc.totalDeleted.toLocaleString() }
        );
    }

    if (model.meta.hasCli) {
        cards.push(
            { label: 'Daily Active CLI Users (Latest)', value: model.totals.latestDailyCliUsers.toLocaleString() },
            { label: 'CLI Requests', value: model.totals.cli.request_count.toLocaleString() }
        );
    }

    if (model.meta.hasCodeReview) {
        cards.push(
            { label: 'Code Review Active Users', value: model.totals.reviewActiveUsers.toLocaleString() },
            { label: 'Code Review Passive Users', value: model.totals.reviewPassiveUsers.toLocaleString() }
        );
    }

    if (model.meta.hasPullRequests) {
        cards.push(
            { label: 'Pull Requests Created', value: model.totals.pullRequests.total_created.toLocaleString() },
            { label: 'Pull Requests Merged', value: model.totals.pullRequests.total_merged.toLocaleString() }
        );
    }
    const cardsHtml = cards.map(c => metricCard(c.label, c.value)).join('');
    let locNoteHtml = '';
    if (model.loc && model.loc.boundary && (model.loc.boundary.hasBefore || model.loc.boundary.hasNulls)) {
        locNoteHtml = `<div class="small-note" style="margin-top:8px;">LoC metrics: reports before 2025-09-01 may show partial/null values. Agent edits are counted in <code>agent_edit</code> as added/deleted; suggestions come from chat panel only.</div>`;
    }
    container.innerHTML = cardsHtml + locNoteHtml;
}

function metricCard(label, value) {
    return `<div class="metric-card"><div class="metric-label">${escapeHtml(label)}</div><div class="metric-value">${escapeHtml(value)}</div></div>`;
}

function sum(arr, fn) { return arr.reduce((acc, x) => acc + (fn(x) || 0), 0); }

function setStatus(msg, isError=false) {
    const el = document.getElementById('statusMessage');
    if (!el) return;
    el.textContent = msg;
    el.dataset.state = isError ? 'error' : 'info';
    el.setAttribute('role', isError ? 'alert' : 'status');
    el.setAttribute('aria-live', isError ? 'assertive' : 'polite');
    el.setAttribute('aria-atomic', 'true');
    window.__statusMessage = { text: msg, error: !!isError };
}

function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.hidden = !show;
    ['mainContent', 'summaryMetrics', 'chartsContainer', 'tablesContainer', 'userUsageTable'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.setAttribute('aria-busy', show ? 'true' : 'false');
    });
}

function enableDownloadButton() {
    const btn = document.getElementById('downloadPdfBtn');
    if (btn) btn.disabled = false;
    const uBtn = document.getElementById('userUsageBtn');
    if (uBtn) uBtn.disabled = !(window.__dashboardModel && window.__dashboardModel.meta.hasUserRecords);
}

function escapeHtml(str) {
    return String(str).replace(/[&<>"]+/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[s]));
}

// Numeric helper: coerce to number or 0
function toNum(v) {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
}

// Map raw feature code names to friendly display names
function formatFeatureName(raw) {
    if (!raw) return 'Unknown';
    let name = raw;
    // Remove chat_panel_ prefix entirely
    name = name.replace(/^chat_panel_/, '');
    // Replace known tokens
    name = name.replace(/_mode$/, '');
    // Replace underscores with spaces
    name = name.split('_').map(part => part ? (part.charAt(0).toUpperCase() + part.slice(1)) : part).join(' ');
    // Specific canonical overrides
    const overrides = {
        'Agent Edit': 'Agent Edit',
        'Code Completion': 'Code Completion',
        'Chat Inline': 'Inline Chat',
        'Ask': 'Ask',
        'Agent': 'Agent',
        'Custom': 'Custom',
        'Edit': 'Edit',
        'Plan': 'Plan',
        'Unknown': 'Unknown'
    };
    return overrides[name] || name;
}

// -------- Placeholder / skeleton rendering -------- //
function buildMetricPlaceholders(mode = 'upload', count = 4) {
    const isFilterEmpty = mode === 'filters';
    const title = isFilterEmpty
        ? 'No metrics match these filters'
        : 'Load a metrics export to see the dashboard';
    const description = isFilterEmpty
        ? 'Clear a filter or widen the date range.'
        : 'Cards, charts, and tables appear after upload.';
    const checklist = isFilterEmpty
        ? [
            'Clear or relax the current filters.',
            'Turn off the members-only filter if no members source is loaded.'
        ]
        : [
            'Start with the Copilot metrics export.',
            'Optionally load organization members for members-only filtering.'
        ];
    const guidance = `<div class="empty-state" role="note">
        <div class="empty-state-copy">
            <h3>${escapeHtml(title)}</h3>
            <p>${escapeHtml(description)}</p>
            <ul class="empty-state-list">${checklist.map(item => `<li>${escapeHtml(item)}</li>`).join('')}</ul>
        </div>
        <div class="empty-preview" aria-hidden="true">
            <span class="empty-chip">Local only</span>
            <div class="empty-preview-card"></div>
        </div>
    </div>`;
    const cards = Array.from({ length: count }).map(() => '<div class="metric-card placeholder skeleton" aria-hidden="true"></div>').join('');
    return guidance + cards;
}

function buildChartPlaceholders(mode = 'upload', count = 1) {
    const isFilterEmpty = mode === 'filters';
    const title = isFilterEmpty
        ? 'No charts match these filters'
        : 'Charts appear here after upload';
    const description = isFilterEmpty
        ? 'Adjust the filters and apply them again.'
        : 'Adoption, usage, and code-generation views populate here after upload.';
    const guidance = `<div class="empty-state" role="note">
        <div class="empty-state-copy">
            <h3>${escapeHtml(title)}</h3>
            <p>${escapeHtml(description)}</p>
            <ul class="empty-state-list">
                <li>Use quick ranges to jump to the latest 7, 14, or 28 days.</li>
                <li>Open the reference tables for exact values.</li>
            </ul>
        </div>
        <div class="empty-preview" aria-hidden="true">
            <span class="empty-chip">Readable charts</span>
            <div class="empty-preview-card"></div>
        </div>
    </div>`;
    const charts = Array.from({ length: count }).map(() => `<div class="chart-container skeleton" aria-hidden="true">
            <div class="placeholder-chart">
                <div class="placeholder-bar lg skeleton"></div>
                ${Array.from({ length: 6 }).map(() => '<div class="placeholder-bar md skeleton"></div>').join('')}
            </div>
        </div>`).join('');
    return guidance + charts;
}

function renderPlaceholders(mode = 'upload') {
    const metrics = document.getElementById('summaryMetrics');
    if (metrics) metrics.innerHTML = buildMetricPlaceholders(mode);
    const charts = document.getElementById('chartsContainer');
    if (charts) charts.innerHTML = buildChartPlaceholders(mode);
}

// --- Enhanced PDF generation (multi-page) ---
// ================= LoC Aggregations & Charts ================= //
function computeLocAggregations(data) {
    if (!Array.isArray(data) || !data.length) return null;
    const byFeature = new Map(); // feature -> {suggestAdd, suggestDel, added, deleted}
    const byLanguage = new Map(); // language -> {added, deleted}
    let totalSuggestedAdd = 0, totalSuggestedDel = 0, totalAdded = 0, totalDeleted = 0;
    let hasNulls = false;
    let hasBefore = false, hasOnOrAfter = false;
    for (const r of data) {
        const day = r.day;
        if (day) {
            if (day < '2025-09-01') hasBefore = true; else hasOnOrAfter = true;
        }
        const rootLocFields = [
            r.loc_suggested_to_add_sum,
            r.loc_suggested_to_delete_sum,
            r.loc_added_sum,
            r.loc_deleted_sum
        ];
        const recordHasRootLocTotals = rootLocFields.some(v => v !== undefined && v !== null);
        if (recordHasRootLocTotals) {
            totalSuggestedAdd += toNum(r.loc_suggested_to_add_sum);
            totalSuggestedDel += toNum(r.loc_suggested_to_delete_sum);
            totalAdded += toNum(r.loc_added_sum);
            totalDeleted += toNum(r.loc_deleted_sum);
            if (rootLocFields.some(v => v === null)) hasNulls = true;
        }
        if (Array.isArray(r.totals_by_feature)) {
            for (const f of r.totals_by_feature) {
                const feat = f.feature || 'unknown';
                const sAdd = f.loc_suggested_to_add_sum;
                const sDel = f.loc_suggested_to_delete_sum;
                const add = f.loc_added_sum;
                const del = f.loc_deleted_sum;
                if (sAdd === null || sDel === null || add === null || del === null) hasNulls = true;
                const v = byFeature.get(feat) || { suggestAdd: 0, suggestDel: 0, added: 0, deleted: 0 };
                if (Number.isFinite(Number(sAdd))) {
                    v.suggestAdd += Number(sAdd);
                    if (!recordHasRootLocTotals) totalSuggestedAdd += Number(sAdd);
                }
                if (Number.isFinite(Number(sDel))) {
                    v.suggestDel += Number(sDel);
                    if (!recordHasRootLocTotals) totalSuggestedDel += Number(sDel);
                }
                if (Number.isFinite(Number(add))) {
                    v.added += Number(add);
                    if (!recordHasRootLocTotals) totalAdded += Number(add);
                }
                if (Number.isFinite(Number(del))) {
                    v.deleted += Number(del);
                    if (!recordHasRootLocTotals) totalDeleted += Number(del);
                }
                byFeature.set(feat, v);
            }
        }
        if (Array.isArray(r.totals_by_language_feature)) {
            for (const lf of r.totals_by_language_feature) {
                const lang = lf.language || 'unknown';
                const add = lf.loc_added_sum;
                const del = lf.loc_deleted_sum;
                if (add === null || del === null) hasNulls = true;
                const v = byLanguage.get(lang) || { added: 0, deleted: 0 };
                if (Number.isFinite(Number(add))) { v.added += Number(add); }
                if (Number.isFinite(Number(del))) { v.deleted += Number(del); }
                byLanguage.set(lang, v);
            }
        }
    }
    return { totalSuggestedAdd, totalSuggestedDel, totalAdded, totalDeleted, byFeature, byLanguage, boundary: { hasBefore, hasOnOrAfter, hasNulls } };
}
async function generatePdfReport() {
    const { jsPDF } = window.jspdf || {};
    if (!jsPDF || !window.html2canvas) {
        setStatus('PDF export is unavailable because the PDF libraries did not load.', true);
        return;
    }
    const downloadBtn = document.getElementById('downloadPdfBtn');
    const originalTheme = getCurrentTheme();
    const shouldRestoreTheme = originalTheme !== 'light';
    const loadingTextEl = document.querySelector('#loadingOverlay .loading-text');
    const previousLoadingText = loadingTextEl ? loadingTextEl.textContent : '';
    if (downloadBtn) downloadBtn.disabled = true;
    if (loadingTextEl) loadingTextEl.textContent = 'Building PDF…';
    showLoading(true);
    setStatus('Building PDF report…');
    try {
        if (shouldRestoreTheme) {
            applyTheme('light', { persist: false });
            await waitForPaint(2);
        }
        const doc = new jsPDF({ orientation: 'p', unit: 'pt', format: 'a4' });
        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        const margin = 36; // 0.5in
        let cursorY = margin;
        const lineHeight = 14;
    // Pull optional enterprise/org names
    const enterpriseName = (document.getElementById('enterpriseName')?.value || '').trim();
    const orgName = (document.getElementById('orgName')?.value || '').trim();
        const addHeader = (title, subtitle) => {
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(16);
            doc.text(title, margin, cursorY);
            cursorY += 20;
            doc.setFont('helvetica', 'normal');
            doc.setFontSize(9);
            doc.setTextColor(90);
            doc.text(subtitle, margin, cursorY);
            doc.setTextColor(0);
            cursorY += 18;
        };
        const ensurePage = (neededHeight) => {
            if (cursorY + neededHeight + margin > pageHeight) {
                doc.addPage();
                cursorY = margin;
            }
        };
        const timestamp = new Date().toLocaleString();
    let headerTitle = 'Copilot Metrics Report';
    if (enterpriseName && orgName) headerTitle = `${enterpriseName} – ${orgName} Copilot Metrics Report`;
    else if (enterpriseName) headerTitle = `${enterpriseName} Copilot Metrics Report`;
    else if (orgName) headerTitle = `${orgName} Copilot Metrics Report`;
    const headerSubtitleParts = [`Generated ${timestamp}`];
    if (enterpriseName && !headerTitle.includes(enterpriseName)) headerSubtitleParts.push(enterpriseName);
    if (orgName && !headerTitle.includes(orgName)) headerSubtitleParts.push(orgName);
    addHeader(headerTitle, headerSubtitleParts.join('  •  '));
        // Summary metrics: render as text grid (3 columns)
        const metrics = Array.from(document.querySelectorAll('#summaryMetrics .metric-card'))
            .map(card => ({ label: card.querySelector('.metric-label')?.textContent?.trim(), value: card.querySelector('.metric-value')?.textContent?.trim() }));
        const colCount = 3;
        const colWidth = (pageWidth - margin * 2) / colCount;
        doc.setFontSize(10);
        // Render metrics in aligned rows (each row contains up to colCount metrics)
        const rows = [];
        for (let i = 0; i < metrics.length; i += colCount) {
            rows.push(metrics.slice(i, i + colCount));
        }
        const rowLabelOffset = 0;
        const rowValueOffset = lineHeight; // value below label
        const rowHeight = lineHeight * 2 + 6; // label + value + padding
        rows.forEach((row, rIdx) => {
            ensurePage(rowHeight);
            const rowY = cursorY + rowLabelOffset;
            row.forEach((m, cIdx) => {
                const x = margin + cIdx * colWidth;
                doc.setFont('helvetica', 'bold');
                doc.text(m.label || '', x, rowY);
                doc.setFont('helvetica', 'normal');
                doc.text(String(m.value || ''), x, rowY + rowValueOffset);
            });
            cursorY += rowHeight;
            // Optional subtle separator except after last row
            if (rIdx !== rows.length - 1) {
                doc.setDrawColor(235);
                doc.setLineWidth(0.4);
                doc.line(margin, cursorY - 4, pageWidth - margin, cursorY - 4);
                doc.setDrawColor(0);
            }
        });
        cursorY += 4; // extra spacing before charts
        // Capture each chart sequentially (to reduce memory)
        const chartDivs = Array.from(document.querySelectorAll('.chart-container'));
    const chartImageScale = 1.5; // upscale factor for higher DPI in PDF
        for (let i = 0; i < chartDivs.length; i++) {
            const chartEl = chartDivs[i];
            const title = chartEl.querySelector('.highcharts-title')?.textContent?.trim() || chartEl.getAttribute('aria-label') || `Chart ${i+1}`;
            // Use Highcharts built-in export to get high-res data URL if available
            let dataUrl; let tmpSVG;
            const hcChart = Highcharts.charts.find(c => c && c.renderTo && chartEl.contains(c.renderTo));
            if (hcChart) {
                try {
                    // Export to PNG via built-in toDataURL fallback using SVG
                    tmpSVG = hcChart.getSVG();
                    // Convert SVG to canvas
                    const svgBlob = new Blob([tmpSVG], { type: 'image/svg+xml;charset=utf-8' });
                    const url = URL.createObjectURL(svgBlob);
                    dataUrl = await new Promise(resolve => {
                        const img = new Image();
                        img.onload = () => {
                const canvas = document.createElement('canvas');
                canvas.width = img.width * chartImageScale; 
                canvas.height = img.height * chartImageScale;
                            const ctx = canvas.getContext('2d');
                ctx.imageSmoothingQuality = 'high';
                ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
                            resolve(canvas.toDataURL('image/png'));
                            URL.revokeObjectURL(url);
                        };
                        img.src = url;
                    });
                } catch (_) { /* fallback below */ }
            }
            if (!dataUrl) {
                // Fallback: rasterize container
                const exportSource = hcChart?.renderTo || chartEl.firstElementChild || chartEl;
        const canvas = await html2canvas(exportSource, { backgroundColor: '#ffffff', scale: chartImageScale, useCORS: true });
                dataUrl = canvas.toDataURL('image/png');
            }
            // Scale image to fit width
            const imgProps = doc.getImageProperties(dataUrl);
            const maxImgWidth = pageWidth - margin * 2;
            const scale = maxImgWidth / imgProps.width;
            const imgHeight = imgProps.height * scale;
            ensurePage(imgHeight + 34);
            // Title
            doc.setFont('helvetica', 'bold');
            doc.setFontSize(12);
            doc.text(title, margin, cursorY);
            cursorY += 14;
            doc.addImage(dataUrl, 'PNG', margin, cursorY, maxImgWidth, imgHeight);
            cursorY += imgHeight + 20;
        }
        // Footer
        doc.setFontSize(8);
        const pageCount = doc.getNumberOfPages();
        for (let p = 1; p <= pageCount; p++) {
            doc.setPage(p);
            doc.setFont('helvetica', 'normal');
            doc.setTextColor(120);
            doc.text(`Page ${p} / ${pageCount}`, pageWidth - margin - 60, pageHeight - 20);
            doc.text('Generated locally - Copilot Metrics Dashboard', margin, pageHeight - 20);
            doc.setTextColor(0);
        }
    const dateStr = new Date().toISOString().substring(0,10);
    const safe = s => s.replace(/[^a-z0-9]+/gi,'-').replace(/^-+|-+$/g,'').toLowerCase();
    let nameParts = [];
    if (enterpriseName) nameParts.push(safe(enterpriseName));
    if (orgName) nameParts.push(safe(orgName));
    nameParts.push('copilot-metrics-report', dateStr);
    const fileName = nameParts.join('-') + '.pdf';
        doc.save(fileName);
        setStatus('PDF ready.');
    } catch (err) {
        console.error('PDF generation error', err);
        setStatus(`PDF generation failed: ${err.message}`, true);
    } finally {
        if (shouldRestoreTheme) {
            applyTheme(originalTheme, { persist: false });
            await waitForPaint(2);
        }
        if (loadingTextEl) loadingTextEl.textContent = previousLoadingText || 'Analyzing…';
        showLoading(false);
        if (downloadBtn) downloadBtn.disabled = false;
    }
}

// ================= Per-User Usage Table & CSV Export ================= //
let userUsageSort = { key: 'user_login', dir: 'asc' };

function toggleUserUsage(show) {
    const section = document.getElementById('userUsageSection');
    const chartsSec = document.getElementById('chartsSection');
    if (!section || !chartsSec) return;
    section.hidden = !show;
    chartsSec.hidden = show;
    if (show) {
        section.scrollIntoView({ behavior: preferredScrollBehavior(), block: 'start' });
        document.getElementById('exportUsersCsvBtn')?.removeAttribute('disabled');
        section.focus({ preventScroll: true });
    } else {
        chartsSec.scrollIntoView({ behavior: preferredScrollBehavior(), block: 'start' });
        chartsSec.focus({ preventScroll: true });
    }
}

function aggregateUserUsage(data) {
    const map = new Map();
    data.forEach(r => {
        const login = r.user_login || (r.user_id ? `user-${r.user_id}` : '');
        if (!login) return;
        if (!map.has(login)) {
            map.set(login, {
                user_login: login,
                user_id: r.user_id,
                interactions: 0,
                completions: 0,
                acceptances: 0,
                acceptance_rate: 0,
                days_active: new Set(),
                chat_days: new Set(),
                agent_days: new Set(),
                cli_days: new Set(),
                review_active_days: new Set(),
                review_passive_days: new Set(),
                models: {},
                languages: {},
                features: {},
                ides: {},
                chat_modes: {},
                cli_requests: 0,
                cli_sessions: 0,
                loc_suggested_add: 0,
                loc_suggested_delete: 0,
                loc_added: 0,
                loc_deleted: 0
            });
        }
        const row = map.get(login);
        row.interactions += (r.user_initiated_interaction_count || 0);
        row.completions += (r.code_generation_activity_count || 0);
        row.acceptances += (r.code_acceptance_activity_count || 0);
        if (r.day) row.days_active.add(r.day);
        if (r.day && (r.used_chat || recordContainsChatActivity(r))) row.chat_days.add(r.day);
        if (r.day && (r.used_agent || recordContainsAgentActivity(r))) row.agent_days.add(r.day);
        if (r.day && (r.used_cli || r.totals_by_cli)) row.cli_days.add(r.day);
        if (r.day && r.used_copilot_code_review_active) row.review_active_days.add(r.day);
        if (r.day && r.used_copilot_code_review_passive) row.review_passive_days.add(r.day);
        if (r.totals_by_model_feature) {
            r.totals_by_model_feature.forEach(mf => {
                const m = mf.model || 'unknown';
                row.models[m] = (row.models[m] || 0) + pickPrimaryUsageValue(mf);
            });
        }
        if (r.totals_by_language_feature) {
            r.totals_by_language_feature.forEach(lf => {
                const lang = lf.language || 'unknown';
                const val = pickPrimaryUsageValue(lf);
                row.languages[lang] = (row.languages[lang] || 0) + val;
            });
        }
        if (r.totals_by_ide) {
            r.totals_by_ide.forEach(ideBucket => {
                const ide = ideBucket.ide || 'unknown';
                row.ides[ide] = (row.ides[ide] || 0) + pickPrimaryUsageValue(ideBucket);
            });
        }
        if (r.totals_by_feature) {
            r.totals_by_feature.forEach(f => {
                const feat = f.feature || 'unknown';
                row.features[feat] = (row.features[feat] || 0) + pickPrimaryUsageValue(f);
                if (CHAT_MODE_FEATURES.includes(feat)) {
                    row.chat_modes[feat] = (row.chat_modes[feat] || 0) + (f.user_initiated_interaction_count || 0);
                }
                // LoC per-user: sum from feature bucket to avoid double-counting
                row.loc_suggested_add += toNum(f.loc_suggested_to_add_sum);
                row.loc_suggested_delete += toNum(f.loc_suggested_to_delete_sum);
                row.loc_added += toNum(f.loc_added_sum);
                row.loc_deleted += toNum(f.loc_deleted_sum);
            });
        }
        if (r.totals_by_cli) {
            row.cli_requests += toNum(r.totals_by_cli.request_count);
            row.cli_sessions += toNum(r.totals_by_cli.session_count);
        }
    });
    const rows = Array.from(map.values()).map(r => {
        r.acceptance_rate = r.completions ? (r.acceptances / r.completions * 100) : 0;
        r.days_active_count = r.days_active.size;
        r.chat_days_count = r.chat_days.size;
        r.agent_days_count = r.agent_days.size;
        r.cli_days_count = r.cli_days.size;
        r.review_active_days_count = r.review_active_days.size;
        r.review_passive_days_count = r.review_passive_days.size;
        r.top_model = topKey(r.models);
        r.top_language = topKey(r.languages);
        r.top_feature = formatFeatureName(topKey(r.features));
        r.top_ide = topKey(r.ides);
        r.top_chat_mode = formatFeatureName(topKey(r.chat_modes));
        return r;
    });
    return rows;
}

function topKey(obj) {
    const entries = Object.entries(obj || {});
    if (!entries.length) return '';
    entries.sort((a,b)=>b[1]-a[1]);
    return entries[0][0];
}

function buildUserUsageTable(data) {
    const table = document.getElementById('userUsageTable');
    if (!table) return;
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const rows = aggregateUserUsage(data);
    window.__userUsageRows = rows;
    const exportBtn = document.getElementById('exportUsersCsvBtn');
    // Build head
    thead.innerHTML = '';
    if (!rows.length) {
        if (exportBtn) exportBtn.disabled = true;
        tbody.innerHTML = '<tr><td colspan="22">No user-level records are available for the current filters.</td></tr>';
        return;
    }
    if (exportBtn) exportBtn.disabled = false;
    const headers = [
        { key: 'user_login', label: 'User' },
        { key: 'interactions', label: 'Interactions' },
        { key: 'completions', label: 'Completions' },
        { key: 'acceptances', label: 'Acceptances' },
        { key: 'acceptance_rate', label: 'Acceptance %' },
        { key: 'days_active_count', label: 'Days Active' },
        { key: 'chat_days_count', label: 'Chat Days' },
        { key: 'agent_days_count', label: 'Agent Days' },
        { key: 'cli_days_count', label: 'CLI Days' },
        { key: 'review_active_days_count', label: 'Review Active Days' },
        { key: 'review_passive_days_count', label: 'Review Passive Days' },
        { key: 'cli_requests', label: 'CLI Requests' },
        { key: 'cli_sessions', label: 'CLI Sessions' },
        { key: 'loc_suggested_add', label: 'LoC Suggested' },
        { key: 'loc_suggested_delete', label: 'LoC Suggested Del' },
        { key: 'loc_added', label: 'LoC Added' },
        { key: 'loc_deleted', label: 'LoC Deleted' },
        { key: 'top_model', label: 'Top Model' },
        { key: 'top_language', label: 'Top Language' },
        { key: 'top_feature', label: 'Top Feature' },
        { key: 'top_ide', label: 'Top IDE' },
        { key: 'top_chat_mode', label: 'Top Chat Mode' }
    ];
    const tr = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.dataset.key = h.key;
        th.dataset.label = h.label;
        th.scope = 'col';
        th.classList.add('sortable');
        th.setAttribute('aria-sort', h.key === userUsageSort.key ? (userUsageSort.dir === 'asc' ? 'ascending' : 'descending') : 'none');
        const sortBtn = document.createElement('button');
        sortBtn.type = 'button';
        sortBtn.className = 'sort-button';
        sortBtn.textContent = h.label;
        sortBtn.setAttribute('aria-label', `Sort by ${h.label}`);
        sortBtn.addEventListener('click', () => {
            if (userUsageSort.key === h.key) {
                userUsageSort.dir = userUsageSort.dir === 'asc' ? 'desc' : 'asc';
            } else {
                userUsageSort.key = h.key; userUsageSort.dir = 'asc';
            }
            renderUserUsageRows();
        });
        th.appendChild(sortBtn);
        tr.appendChild(th);
    });
    thead.appendChild(tr);
    renderUserUsageRows();

    function renderUserUsageRows() {
        const key = userUsageSort.key; const dir = userUsageSort.dir === 'asc' ? 1 : -1;
        rows.sort((a,b) => {
            const va = a[key]; const vb = b[key];
            if (typeof va === 'number' && typeof vb === 'number') return (va - vb) * dir;
            return String(va).localeCompare(String(vb), undefined, { sensitivity: 'base' }) * dir;
        });
        tbody.innerHTML = rows.map(r => `<tr>
            <th scope="row">${escapeHtml(r.user_login)}</th>
            <td>${r.interactions}</td>
            <td>${r.completions}</td>
            <td>${r.acceptances}</td>
            <td>${r.acceptance_rate.toFixed(1)}</td>
            <td>${r.days_active_count}</td>
            <td>${r.chat_days_count}</td>
            <td>${r.agent_days_count}</td>
            <td>${r.cli_days_count}</td>
            <td>${r.review_active_days_count}</td>
            <td>${r.review_passive_days_count}</td>
            <td>${r.cli_requests}</td>
            <td>${r.cli_sessions}</td>
            <td>${r.loc_suggested_add}</td>
            <td>${r.loc_suggested_delete}</td>
            <td>${r.loc_added}</td>
            <td>${r.loc_deleted}</td>
            <td>${escapeHtml(r.top_model)}</td>
            <td>${escapeHtml(r.top_language)}</td>
            <td>${escapeHtml(r.top_feature)}</td>
            <td>${escapeHtml(r.top_ide)}</td>
            <td>${escapeHtml(r.top_chat_mode)}</td>
        </tr>`).join('');
        thead.querySelectorAll('th.sortable').forEach(th => {
            const isActive = th.dataset.key === key;
            const label = th.dataset.label || th.dataset.key;
            th.setAttribute('aria-sort', isActive ? (dir === 1 ? 'ascending' : 'descending') : 'none');
            th.classList.toggle('sort-asc', isActive && dir === 1);
            th.classList.toggle('sort-desc', isActive && dir !== 1);
            const button = th.querySelector('.sort-button');
            if (button) {
                button.setAttribute(
                    'aria-label',
                    isActive
                        ? `${label}, sorted ${dir === 1 ? 'ascending' : 'descending'}. Activate to sort ${dir === 1 ? 'descending' : 'ascending'}.`
                        : `Sort by ${label}`
                );
            }
        });
    }
}

function exportUserUsageCsv() {
    const rows = window.__userUsageRows || aggregateUserUsage(window.__currentFilteredData || window.__rawData || []);
    if (!rows.length) { setStatus('No per-user rows are available to export for the current filters.', true); return; }
    const header = [
        'user_login',
        'user_id',
        'interactions',
        'completions',
        'acceptances',
        'acceptance_rate',
        'days_active',
        'chat_days',
        'agent_days',
        'cli_days',
        'review_active_days',
        'review_passive_days',
        'cli_requests',
        'cli_sessions',
        'loc_suggested_add',
        'loc_suggested_delete',
        'loc_added',
        'loc_deleted',
        'top_model',
        'top_language',
        'top_feature',
        'top_ide',
        'top_chat_mode'
    ];
    const lines = [header.join(',')];
    rows.forEach(r => {
        const vals = [
            r.user_login,
            r.user_id,
            r.interactions,
            r.completions,
            r.acceptances,
            r.acceptance_rate.toFixed(2),
            r.days_active_count,
            r.chat_days_count,
            r.agent_days_count,
            r.cli_days_count,
            r.review_active_days_count,
            r.review_passive_days_count,
            r.cli_requests,
            r.cli_sessions,
            r.loc_suggested_add,
            r.loc_suggested_delete,
            r.loc_added,
            r.loc_deleted,
            r.top_model,
            r.top_language,
            r.top_feature,
            r.top_ide,
            r.top_chat_mode
        ].map(v => csvEscape(v));
        lines.push(vals.join(','));
    });
    const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    const dateStr = new Date().toISOString().substring(0,10);
    a.href = url; a.download = `copilot-user-usage-${dateStr}.csv`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function csvEscape(val) {
    if (val === null || val === undefined) return '';
    const s = String(val);
    if (/[",\n]/.test(s)) return '"' + s.replace(/"/g,'""') + '"';
    return s;
}
