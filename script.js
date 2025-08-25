// Apply theme only; user must upload a file now (no default data.json)
window.addEventListener('load', () => {
    setupHighchartsTheme();
    setStatus('Awaiting upload…');
    renderPlaceholders();
});

function setupHighchartsTheme() {
    if (typeof Highcharts === 'undefined') return;
    Highcharts.setOptions({
        chart: {
            backgroundColor: 'transparent',
            style: { fontFamily: 'system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,sans-serif' }
        },
        colors: ['#2563eb', '#059669', '#7c3aed', '#dc2626', '#f59e0b', '#0d9488', '#9333ea', '#ea580c', '#1d4ed8', '#16a34a'],
        title: { style: { color: '#1f2328', fontWeight: '600'} },
        subtitle: { style: { color: '#57606a' } },
        xAxis: {
            lineColor: '#d0d7de',
            tickColor: '#d0d7de',
            gridLineColor: '#e5e9ef',
            labels: { style: { color: '#57606a', fontSize: '11px' } },
            title: { style: { color: '#57606a' } }
        },
        yAxis: {
            lineColor: '#d0d7de',
            tickColor: '#d0d7de',
            gridLineColor: '#e5e9ef',
            labels: { style: { color: '#57606a', fontSize: '11px' } },
            title: { style: { color: '#57606a' } }
        },
        legend: {
            backgroundColor: 'transparent',
            itemStyle: { color: '#1f2328', fontSize: '11px' },
            itemHoverStyle: { color: '#2563eb' }
        },
        tooltip: {
            backgroundColor: '#ffffff',
            borderColor: '#d0d7de',
            style: { color: '#1f2328' },
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
            pie: { dataLabels: { style: { fontSize: '11px', color: '#1f2328' } } }
        },
        credits: { enabled: false }
    });
}

// Enhanced manual file selection & parsing (supports JSON array or JSONL)
let fileInputEl, analyzeBtnEl;
document.addEventListener('DOMContentLoaded', () => {
    fileInputEl = document.getElementById('jsonFileInput');
    analyzeBtnEl = document.getElementById('analyzeBtn');
    if (analyzeBtnEl) {
        analyzeBtnEl.addEventListener('click', () => handleFileSelection(fileInputEl && fileInputEl.files[0]));
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
        fileInputEl.addEventListener('change', () => fileInputEl.files && handleFileSelection(fileInputEl.files[0]));
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

function parseUploadedText(text) {
    // Try full JSON parse first (array or object with records?)
    try {
        const preliminary = JSON.parse(text);
        if (Array.isArray(preliminary)) { console.info('[upload] Parsed as JSON array'); return preliminary; }
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
            if (!Array.isArray(data) || !data.length) {
                throw new Error('Parsed result is empty.');
            }
            window.__rawData = data;
            initializeFilters(data);
            analyzeData(data);
            setStatus(`Loaded ${data.length} records from ${file.name}`);
            enableDownloadButton();
        } catch (err) {
            console.error(err);
            setStatus(`Upload parse error: ${err.message}`, true);
            alert('Error parsing file: ' + err.message);
        } finally {
            showLoading(false);
        }
    };
    reader.onerror = () => { setStatus('File read error', true); showLoading(false); };
    reader.readAsText(file);
}

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
            alert('Error parsing members file: ' + err.message);
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
    statusEl.textContent = size ? `${size}` : '(none)';
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
                const jsonLines = text.split('\n').filter(line => line.trim() !== '');
                const data = jsonLines.map(line => JSON.parse(line));
                window.__rawData = data; // store original dataset
                initializeFilters(data);
                analyzeData(data);
                enableDownloadButton();
            } catch (error) {
                alert('Error parsing JSON file: ' + error.message);
                setStatus('Parse error.', true);
            }
        })
        .catch(error => {
            alert('Error loading data.json: ' + error.message);
            setStatus('Load error.', true);
        })
        .finally(() => {
            showLoading(false);
        });
}

function analyzeData(data) {
    const chartsContainer = document.getElementById('chartsContainer');
    chartsContainer.innerHTML = ''; // Clear previous charts

    if (!data || data.length === 0) {
    setStatus('No data. Upload a file to begin.', true);
    renderPlaceholders();
        return;
    }

    setStatus(`Displaying ${data.length} records`);
    // Accessibility: shift focus to main content after rendering new results
    const main = document.getElementById('mainContent');
    if (main) {
        // Delay slightly to ensure DOM updates (charts) are in place for assistive tech
        setTimeout(() => { main.focus(); }, 0);
    }

    // User interaction count by user (Top 10)
    const userInteractions = {};
    data.forEach(record => {
        const user = record.user_login;
        if (!userInteractions[user]) userInteractions[user] = 0;
        userInteractions[user] += record.user_initiated_interaction_count;
    });
    const topUserInteractions = Object.entries(userInteractions)
        .sort((a,b)=> b[1]-a[1])
        .slice(0,10);
    createChart('User Interaction Count (Top 10)', 'bar', topUserInteractions.map(x=>x[0]), topUserInteractions.map(x=>x[1]));

    // Code generations by language
    const langGenerations = {};
    data.forEach(record => {
        if (record.totals_by_language_feature) {
            record.totals_by_language_feature.forEach(langFeature => {
                const lang = langFeature.language || 'unknown';
                if (!langGenerations[lang]) {
                    langGenerations[lang] = 0;
                }
                langGenerations[lang] += langFeature.code_generation_activity_count;
            });
        }
    });
    createChart('Code Generations by Language', 'pie', Object.keys(langGenerations), Object.values(langGenerations));

    // Acceptances by IDE
    const ideAcceptances = {};
    data.forEach(record => {
        if (record.totals_by_ide) {
            record.totals_by_ide.forEach(ide => {
                const ideName = ide.ide;
                if (!ideAcceptances[ideName]) {
                    ideAcceptances[ideName] = 0;
                }
                ideAcceptances[ideName] += ide.code_acceptance_activity_count;
            });
        }
    });
    createChart('Code Acceptances by IDE', 'doughnut', Object.keys(ideAcceptances), Object.values(ideAcceptances));

    // Completions vs Acceptances (Top 10 by completions)
    const completions = {};
    const acceptances = {};
    data.forEach(record => {
        const user = record.user_login;
        if (!completions[user]) completions[user] = 0;
        if (!acceptances[user]) acceptances[user] = 0;
        completions[user] += record.code_generation_activity_count;
        acceptances[user] += record.code_acceptance_activity_count;
    });
    const topCompUsers = Object.entries(completions)
        .sort((a,b)=> b[1]-a[1])
        .slice(0,10)
        .map(x=>x[0]);
    createGroupedBarChart('Completions vs. Acceptances (Top 10)', topCompUsers,
        [{ label: 'Completions', data: topCompUsers.map(u=>completions[u]) }, { label: 'Acceptances', data: topCompUsers.map(u=>acceptances[u]) }]
    );

    // Daily Active Users
    const dailyUsers = {};
    data.forEach(record => {
        const day = record.day;
        if (!dailyUsers[day]) {
            dailyUsers[day] = new Set();
        }
        dailyUsers[day].add(record.user_id);
    });
    const dailyUserCounts = Object.keys(dailyUsers).map(day => dailyUsers[day].size);
    createChart('Daily Active Users', 'line', Object.keys(dailyUsers), dailyUserCounts);

    // Model Usage
    const modelUsage = {};
    data.forEach(record => {
        if (record.totals_by_model_feature) {
            record.totals_by_model_feature.forEach(modelFeature => {
                const model = modelFeature.model || 'unknown';
                if (!modelUsage[model]) {
                    modelUsage[model] = 0;
                }
                modelUsage[model] += modelFeature.user_initiated_interaction_count;
            });
        }
    });
    createChart('Model Usage', 'pie', Object.keys(modelUsage), Object.values(modelUsage));

    // Feature Usage (with pretty names)
    const featureUsage = {};
    data.forEach(record => {
        if (record.totals_by_feature) {
            record.totals_by_feature.forEach(feature => {
                const featureName = feature.feature;
                if (!featureUsage[featureName]) {
                    featureUsage[featureName] = 0;
                }
                featureUsage[featureName] += feature.user_initiated_interaction_count;
            });
        }
    });
    createChart('Feature Usage', 'bar', Object.keys(featureUsage).map(formatFeatureName), Object.values(featureUsage));

    // Acceptance Rate by User (top N to keep chart readable)
    const acceptanceRateByUser = [];
    Object.keys(completions).forEach(u => {
        const gen = completions[u] || 0;
        const acc = acceptances[u] || 0;
        const rate = gen ? (acc / gen) * 100 : 0;
        acceptanceRateByUser.push({ user: u, rate });
    });
    acceptanceRateByUser.sort((a,b) => b.rate - a.rate);
    const topRateUsers = acceptanceRateByUser.slice(0, 10); // top 10
    createChart('Acceptance Rate % (Top 10 Users)', 'bar', topRateUsers.map(r => r.user), topRateUsers.map(r => +r.rate.toFixed(1)));

    // Heatmap: Language vs Model (using totals_by_language_model if present)
    const langModelMatrix = {}; // lang -> model -> generations
    data.forEach(record => {
        if (record.totals_by_language_model) {
            record.totals_by_language_model.forEach(lm => {
                const lang = lm.language || 'unknown';
                const model = lm.model || 'unknown';
                if (!langModelMatrix[lang]) langModelMatrix[lang] = {};
                langModelMatrix[lang][model] = (langModelMatrix[lang][model] || 0) + (lm.code_generation_activity_count || 0);
            });
        }
    });
    const langCategories = Object.keys(langModelMatrix).slice(0, 40); // limit to 40 languages
    const modelSet = new Set();
    langCategories.forEach(lang => Object.keys(langModelMatrix[lang]).forEach(m => modelSet.add(m)));
    const modelCategories = Array.from(modelSet);
    const heatmapData = [];
    langCategories.forEach((lang, i) => {
        modelCategories.forEach((model, j) => {
            const val = langModelMatrix[lang][model] || 0;
            heatmapData.push([j, i, val]);
        });
    });
    createHeatmap('Code Generations: Language vs Model', modelCategories, langCategories, heatmapData, 'Generations');

    // Heatmap: Feature vs Model (using totals_by_model_feature or totals_by_feature with model unknown)
    const featureModelMatrix = {}; // feature -> model -> interactions
    data.forEach(record => {
        if (record.totals_by_model_feature) {
            record.totals_by_model_feature.forEach(mf => {
                const model = mf.model || 'unknown';
                const feature = mf.feature || 'unknown';
                if (!featureModelMatrix[feature]) featureModelMatrix[feature] = {};
                featureModelMatrix[feature][model] = (featureModelMatrix[feature][model] || 0) + (mf.user_initiated_interaction_count || 0);
            });
        } else if (record.totals_by_feature) {
            // fallback: aggregate without model detail
            record.totals_by_feature.forEach(f => {
                const feature = f.feature || 'unknown';
                if (!featureModelMatrix[feature]) featureModelMatrix[feature] = {};
                featureModelMatrix[feature]['(all models)'] = (featureModelMatrix[feature]['(all models)'] || 0) + (f.user_initiated_interaction_count || 0);
            });
        }
    });
    const featureCategories = Object.keys(featureModelMatrix); // raw keys
    const modelSet2 = new Set();
    featureCategories.forEach(f => Object.keys(featureModelMatrix[f]).forEach(m => modelSet2.add(m)));
    const modelCategories2 = Array.from(modelSet2);
    const heatmapData2 = [];
    featureCategories.forEach((f, i) => {
        modelCategories2.forEach((m, j) => {
            const val = featureModelMatrix[f][m] || 0;
            heatmapData2.push([j, i, val]);
        });
    });
    createHeatmap('Interactions: Feature vs Model', modelCategories2, featureCategories.map(formatFeatureName), heatmapData2, 'Interactions');

    // ===== Additional Metrics & Visualizations =====
    // Overall language usage (interactions) leveraging totals_by_language_feature
    const languageInteractions = {};
    data.forEach(r => {
        if (r.totals_by_language_feature) {
            r.totals_by_language_feature.forEach(lf => {
                const lang = lf.language || 'unknown';
                const val = (lf.user_initiated_interaction_count || lf.code_generation_activity_count || 0);
                languageInteractions[lang] = (languageInteractions[lang] || 0) + val;
            });
        }
    });
    if (Object.keys(languageInteractions).length) {
        createChart('Language Usage (Interactions)', 'pie', Object.keys(languageInteractions), Object.values(languageInteractions));
    }

    // Language usage per day (top 8 + other) stacked area
    const langDayMap = {}; // day -> lang -> count
    const langTotals = {};
    data.forEach(r => {
        const day = r.day || 'unknown';
        if (!langDayMap[day]) langDayMap[day] = {};
        if (r.totals_by_language_feature) {
            r.totals_by_language_feature.forEach(lf => {
                const lang = lf.language || 'unknown';
                const val = (lf.code_generation_activity_count || lf.user_initiated_interaction_count || 0);
                langDayMap[day][lang] = (langDayMap[day][lang] || 0) + val;
                langTotals[lang] = (langTotals[lang] || 0) + val;
            });
        }
    });
    const topLangs = Object.entries(langTotals).sort((a,b)=>b[1]-a[1]).slice(0,8).map(x=>x[0]);
    const langDays = Object.keys(langDayMap).sort();
    if (langDays.length) {
        const langSeries = topLangs.map(lang => ({ name: lang, data: langDays.map(d => langDayMap[d][lang] || 0) }));
        const otherData = langDays.map(d => { let sum=0; Object.keys(langDayMap[d]).forEach(l => { if(!topLangs.includes(l)) sum += langDayMap[d][l]; }); return sum; });
        langSeries.push({ name: 'Other', data: otherData });
        createStackedChart('Language Usage Per Day (Top 8 + Other)', langDays, langSeries, 'area');
    }

    // Model usage per day (chat requests)
    const modelDayMap = {}; // day -> model -> count
    const modelTotalsDay = {};
    data.forEach(r => {
        const day = r.day || 'unknown';
        if (!modelDayMap[day]) modelDayMap[day] = {};
        if (r.totals_by_model_feature) {
            r.totals_by_model_feature.forEach(mf => {
                const model = mf.model || 'unknown';
                const val = mf.user_initiated_interaction_count || 0;
                modelDayMap[day][model] = (modelDayMap[day][model] || 0) + val;
                modelTotalsDay[model] = (modelTotalsDay[model] || 0) + val;
            });
        }
    });
    const topModels = Object.entries(modelTotalsDay).sort((a,b)=>b[1]-a[1]).slice(0,8).map(x=>x[0]);
    const modelDays = Object.keys(modelDayMap).sort();
    if (modelDays.length) {
        const modelSeries = topModels.map(m => ({ name: m, data: modelDays.map(d => modelDayMap[d][m] || 0) }));
        const otherModelData = modelDays.map(d => { let sum=0; Object.keys(modelDayMap[d]).forEach(mm => { if(!topModels.includes(mm)) sum += modelDayMap[d][mm]; }); return sum; });
        modelSeries.push({ name: 'Other', data: otherModelData });
        createStackedChart('Model Usage Per Day (Chat Requests)', modelDays, modelSeries, 'area');
    }

    // Model usage per feature (stacked column)
    if (featureCategories.length && modelCategories2.length) {
        const stackedSeries = modelCategories2.map(model => ({ name: model, data: featureCategories.map(f => featureModelMatrix[f][model] || 0) }));
        createStackedChart('Model Usage per Feature', featureCategories.map(formatFeatureName), stackedSeries, 'column');
    }

    // Weekly active users (unique user IDs per ISO week start Monday)
    const weekMap = {}; // weekStart -> Set(user)
    data.forEach(r => {
        if (!r.day) return; const date = new Date(r.day + 'T00:00:00Z');
        const dayNum = date.getUTCDay();
        const diffToMonday = (dayNum + 6) % 7; // Monday baseline
        const monday = new Date(date); monday.setUTCDate(date.getUTCDate() - diffToMonday);
        const key = monday.toISOString().substring(0,10);
        if (!weekMap[key]) weekMap[key] = new Set();
        weekMap[key].add(r.user_id);
    });
    const weekKeys = Object.keys(weekMap).sort();
    if (weekKeys.length) {
        const weekCounts = weekKeys.map(k => weekMap[k].size);
        createChart('Weekly Active Users', 'line', weekKeys, weekCounts);
    }
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
        const opts = {
            chart: { type: type === 'doughnut' ? 'pie' : (type === 'bar' ? 'column' : type), backgroundColor: 'transparent', height: 420 },
            title: { text: title, style: { fontSize: '13px' } },
            xAxis: { categories: cats, labels: { style: { fontSize: '11px' } } },
            yAxis: { title: { text: null }, gridLineColor: 'rgba(255,255,255,0.07)' },
            legend: { itemStyle: { fontSize: '11px' } },
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
                    style: { fontSize: '11px', fontWeight: '500', textOutline: 'none', color: '#1f2328' }
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
    const div = document.createElement('div');
    chartContainer.appendChild(div);
    chartsContainer.appendChild(chartContainer);

    Highcharts.chart(div, {
    chart: { type: 'column', backgroundColor: 'transparent', height: 420 },
        title: { text: title, style: { fontSize: '13px' } },
        xAxis: { categories: categories, crosshair: true },
        yAxis: { min: 0, title: { text: null } },
        tooltip: { shared: true },
        legend: { itemStyle: { fontSize: '11px' } },
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
    const div = document.createElement('div');
    chartContainer.appendChild(div);
    chartsContainer.appendChild(chartContainer);
    const maxVal = dataPoints.reduce((m,p)=> Math.max(m,p[2]),0) || 0;
    Highcharts.chart(div, {
    chart: { type: 'heatmap', backgroundColor: 'transparent', height: 480 },
        title: { text: title, style: { fontSize: '13px' } },
        xAxis: { categories: xCategories, labels: { style: { fontSize: '10px' }, rotation: 40 } },
        yAxis: { categories: yCategories, title: null, labels: { style: { fontSize: '10px' } }, reversed: true },
        accessibility: { enabled: true },
        legend: { align: 'right', layout: 'vertical', verticalAlign: 'middle' },
        colorAxis: {
            min: 0,
            max: maxVal,
            stops: [
                [0, '#f0f6ff'],
                [0.4, '#93c5fd'],
                [0.7, '#3b82f6'],
                [1, '#1d4ed8']
            ]
        },
        tooltip: { 
            formatter: function() { return `<b>${this.series.name}</b><br/>${Highcharts.numberFormat(this.point.value,0,'.',',')} ${colorAxisTitle}<br/>`;} 
        },
        series: [{
            name: colorAxisTitle,
            borderWidth: 1,
            borderColor: '#ffffff',
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
        title: { text: title, style: { fontSize: '13px' } },
        xAxis: { categories, labels: { style: { fontSize: '10px' } } },
        yAxis: { min: 0, title: { text: null } },
        legend: { itemStyle: { fontSize: '11px' } },
        tooltip: { shared: true },
        plotOptions: { 
            series: { stacking },
            area: { stacking, marker: { enabled: false }, lineWidth: 1 }
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
    computeSummaryMetrics(window.__rawData || []);
    analyzeData(window.__rawData || []);
    };

    // Auto-apply when members-only checkbox toggled
    const membersChk = document.getElementById('membersOnlyChk');
    if (membersChk) {
        membersChk.addEventListener('change', () => {
            applyFilters();
        });
    }

    // Theme toggle removed (light mode default)

    computeSummaryMetrics(data);

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
    computeSummaryMetrics(filtered);
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

function computeSummaryMetrics(data) {
    const container = document.getElementById('summaryMetrics');
    if (!container) return;
    container.innerHTML = '';
    if (!data.length) { container.innerHTML = buildMetricPlaceholders(); return; }

    const uniqueUsers = new Set(data.map(r => r.user_id)).size;
    const totalInteractions = sum(data, r => r.user_initiated_interaction_count);
    const totalGenerations = sum(data, r => r.code_generation_activity_count);
    const totalAcceptances = sum(data, r => r.code_acceptance_activity_count);
    const acceptanceRate = totalGenerations ? ((totalAcceptances / totalGenerations) * 100).toFixed(1) : '0.0';
    const days = new Set(data.map(r => r.day)).size;
    const avgChatPerUser = uniqueUsers ? (totalInteractions / uniqueUsers).toFixed(1) : '0.0';

    // Chat features & agent adoption
    const chatFeatureNames = ['chat_panel_agent_mode','chat_panel_unknown_mode','chat_panel_ask_mode','chat_inline','chat_panel_custom_mode','chat_panel_edit_mode'];
    const chatUsers = new Set();
    const agentUsers = new Set();
    let chatRequestsTotal = 0;
    data.forEach(r => {
        if (r.totals_by_feature) {
            let userHadChat = false, userHadAgent = false;
            r.totals_by_feature.forEach(f => {
                if (chatFeatureNames.includes(f.feature)) {
                    userHadChat = true;
                    chatRequestsTotal += (f.user_initiated_interaction_count || 0);
                    if (f.feature === 'chat_panel_agent_mode') userHadAgent = true;
                }
            });
            if (userHadChat) chatUsers.add(r.user_id);
            if (userHadAgent) agentUsers.add(r.user_id);
        }
    });
    const avgChatPerChatUser = chatUsers.size ? (chatRequestsTotal / chatUsers.size).toFixed(1) : '0.0';
    const agentAdoptionPct = uniqueUsers ? ((agentUsers.size / uniqueUsers) * 100).toFixed(1) : '0.0';

    // Most used chat model
    const modelTotals = {};
    data.forEach(r => { if (r.totals_by_model_feature) { r.totals_by_model_feature.forEach(mf => { const m = mf.model || 'unknown'; modelTotals[m] = (modelTotals[m] || 0) + (mf.user_initiated_interaction_count || 0); }); }});
    const mostUsedChatModel = Object.entries(modelTotals).sort((a,b)=>b[1]-a[1])[0]?.[0] || 'n/a';

    // Weekly active users (latest week)
    const weekUserMap = {};
    data.forEach(r => { if (!r.day) return; const d=new Date(r.day+'T00:00:00Z'); const dow=d.getUTCDay(); const diff=(dow+6)%7; const monday=new Date(d); monday.setUTCDate(d.getUTCDate()-diff); const key=monday.toISOString().substring(0,10); if(!weekUserMap[key]) weekUserMap[key]=new Set(); weekUserMap[key].add(r.user_id); });
    const latestWeekCount = Object.keys(weekUserMap).sort().slice(-1).map(k => weekUserMap[k].size)[0] || 0;

    const cards = [
        { label: 'Total Active Users', value: uniqueUsers },
        { label: 'Total Interactions', value: totalInteractions },
        { label: 'Code Completions', value: totalGenerations },
        { label: 'Completions Accepted', value: totalAcceptances },
        { label: 'Completion Acceptance Rate %', value: acceptanceRate },
        { label: 'Avg Interactions / User', value: avgChatPerUser },
        { label: 'Avg Chat Requests / Chat User', value: avgChatPerChatUser },
        { label: 'Agent Adoption %', value: agentAdoptionPct },
        { label: 'Most Used Chat Model', value: mostUsedChatModel },
        { label: 'Weekly Active Users (Latest)', value: latestWeekCount },
        { label: 'Distinct Days', value: days }
    ];
    container.innerHTML = cards.map(c => metricCard(c.label, c.value)).join('');
}

function metricCard(label, value) {
    return `<div class="metric-card"><div class="metric-label">${escapeHtml(label)}</div><div class="metric-value">${escapeHtml(value)}</div></div>`;
}

function sum(arr, fn) { return arr.reduce((acc, x) => acc + (fn(x) || 0), 0); }

function setStatus(msg, isError=false) {
    const el = document.getElementById('statusMessage');
    if (!el) return; el.textContent = msg; el.style.color = isError ? 'var(--danger)' : 'var(--text-dim)';
    window.__statusMessage = { text: msg, error: !!isError };
}

function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (overlay) overlay.hidden = !show;
}

function enableDownloadButton() {
    const btn = document.getElementById('downloadPdfBtn');
    if (btn) btn.disabled = false;
    const uBtn = document.getElementById('userUsageBtn');
    if (uBtn) uBtn.disabled = false;
}

function escapeHtml(str) {
    return String(str).replace(/[&<>"]+/g, s => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[s]));
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
        'Code Completion': 'Code Completion',
        'Chat Inline': 'Inline Chat',
        'Ask': 'Ask',
        'Agent': 'Agent',
        'Custom': 'Custom',
        'Edit': 'Edit'
    };
    return overrides[name] || name;
}

// -------- Placeholder / skeleton rendering -------- //
function buildMetricPlaceholders(count=8) {
    return Array.from({length: count}).map(()=> '<div class="metric-card placeholder skeleton"></div>').join('');
}

function renderPlaceholders() {
    const metrics = document.getElementById('summaryMetrics');
    if (metrics && !metrics.children.length) {
        metrics.innerHTML = buildMetricPlaceholders();
    }
    const charts = document.getElementById('chartsContainer');
    if (charts && !charts.children.length) {
    // Reduced placeholder preview charts from 4 to 2 for less initial clutter
    for (let i=0;i<2;i++) {
            const c = document.createElement('div');
            c.className = 'chart-container skeleton';
            c.innerHTML = '<div class="placeholder-chart">'
                + '<div class="placeholder-bar lg skeleton"></div>'
                + Array.from({length:8}).map(()=>'<div class="placeholder-bar md skeleton"></div>').join('')
                + '</div>';
            charts.appendChild(c);
        }
    }
}

// --- Enhanced PDF generation (multi-page) ---
async function generatePdfReport() {
    const { jsPDF } = window.jspdf || {};
    if (!jsPDF || !window.html2canvas) {
        alert('PDF libraries not loaded.');
        return;
    }
    const downloadBtn = document.getElementById('downloadPdfBtn');
    if (downloadBtn) downloadBtn.disabled = true;
    setStatus('Building PDF report…');
    try {
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
        const canvas = await html2canvas(chartEl, { backgroundColor: '#ffffff', scale: chartImageScale, useCORS: true });
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
        setStatus('PDF generation failed', true);
        alert('PDF generation failed: ' + err.message);
    } finally {
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
        section.scrollIntoView({ behavior: 'smooth', block: 'start' });
        document.getElementById('exportUsersCsvBtn')?.removeAttribute('disabled');
    } else {
        chartsSec.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
}

function aggregateUserUsage(data) {
    const map = new Map();
    data.forEach(r => {
        const login = r.user_login || '(unknown)';
        if (!map.has(login)) {
            map.set(login, {
                user_login: login,
                user_id: r.user_id,
                interactions: 0,
                completions: 0,
                acceptances: 0,
                acceptance_rate: 0,
                days_active: new Set(),
                models: {},
                languages: {},
                features: {}
            });
        }
        const row = map.get(login);
        row.interactions += (r.user_initiated_interaction_count || 0);
        row.completions += (r.code_generation_activity_count || 0);
        row.acceptances += (r.code_acceptance_activity_count || 0);
        if (r.day) row.days_active.add(r.day);
        if (r.totals_by_model_feature) {
            r.totals_by_model_feature.forEach(mf => {
                const m = mf.model || 'unknown';
                row.models[m] = (row.models[m] || 0) + (mf.user_initiated_interaction_count || 0);
            });
        }
        if (r.totals_by_language_feature) {
            r.totals_by_language_feature.forEach(lf => {
                const lang = lf.language || 'unknown';
                const val = (lf.user_initiated_interaction_count || lf.code_generation_activity_count || 0);
                row.languages[lang] = (row.languages[lang] || 0) + val;
            });
        }
        if (r.totals_by_feature) {
            r.totals_by_feature.forEach(f => {
                const feat = f.feature || 'unknown';
                row.features[feat] = (row.features[feat] || 0) + (f.user_initiated_interaction_count || 0);
            });
        }
    });
    const rows = Array.from(map.values()).map(r => {
        r.acceptance_rate = r.completions ? (r.acceptances / r.completions * 100) : 0;
        r.days_active_count = r.days_active.size;
        r.top_model = topKey(r.models);
        r.top_language = topKey(r.languages);
        r.top_feature = formatFeatureName(topKey(r.features));
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
    // Build head
    thead.innerHTML = '';
    const headers = [
        { key: 'user_login', label: 'User' },
        { key: 'interactions', label: 'Interactions' },
        { key: 'completions', label: 'Completions' },
        { key: 'acceptances', label: 'Acceptances' },
        { key: 'acceptance_rate', label: 'Acceptance %' },
        { key: 'days_active_count', label: 'Days Active' },
        { key: 'top_model', label: 'Top Model' },
        { key: 'top_language', label: 'Top Language' },
        { key: 'top_feature', label: 'Top Feature' }
    ];
    const tr = document.createElement('tr');
    headers.forEach(h => {
        const th = document.createElement('th');
        th.textContent = h.label;
        th.dataset.key = h.key;
        th.scope = 'col';
        th.classList.add('sortable');
        if (h.key === userUsageSort.key) th.classList.add(userUsageSort.dir === 'asc' ? 'sort-asc' : 'sort-desc');
        th.addEventListener('click', () => {
            if (userUsageSort.key === h.key) {
                userUsageSort.dir = userUsageSort.dir === 'asc' ? 'desc' : 'asc';
            } else {
                userUsageSort.key = h.key; userUsageSort.dir = 'asc';
            }
            renderUserUsageRows();
        });
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
            <td>${escapeHtml(r.user_login)}</td>
            <td>${r.interactions}</td>
            <td>${r.completions}</td>
            <td>${r.acceptances}</td>
            <td>${r.acceptance_rate.toFixed(1)}</td>
            <td>${r.days_active_count}</td>
            <td>${escapeHtml(r.top_model)}</td>
            <td>${escapeHtml(r.top_language)}</td>
            <td>${escapeHtml(r.top_feature)}</td>
        </tr>`).join('');
        table.querySelectorAll('th').forEach(th => { th.classList.remove('sort-asc','sort-desc'); if (th.dataset.key === key) th.classList.add(dir === 1 ? 'sort-asc' : 'sort-desc'); });
    }
}

function exportUserUsageCsv() {
    const rows = window.__userUsageRows || aggregateUserUsage(window.__currentFilteredData || window.__rawData || []);
    if (!rows.length) { alert('No rows to export'); return; }
    const header = ['user_login','user_id','interactions','completions','acceptances','acceptance_rate','days_active','top_model','top_language','top_feature'];
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
            r.top_model,
            r.top_language,
            r.top_feature
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
