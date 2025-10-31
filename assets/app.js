const { jsPDF } = window.jspdf;

// --- State Management ---
let parsedData = [];
let headers = [];
let plotTabs = [];
let activeTabId = null;
let savedReportViews = [];
let currentTheme = 'dark'; // Dark by default

// NEW: Tom Select Instances
let tomSelects = {
    xAxis: null,
    leftYAxis: null,
    rightYAxis: null,
    grouping: null
};

// --- DOM Element Cache ---
const dom = {
    notificationContainer: document.getElementById('notification-container'),
    dataInput: document.getElementById('dataInput'),
    fileInput: document.getElementById('fileInput'),
    loadDataBtn: document.getElementById('loadDataBtn'),
    exportPlotBtn: document.getElementById('exportPlotBtn'),
    exportPlotSVGBtn: document.getElementById('exportPlotSVGBtn'), // NEW
    exportPlotHTMLBtn: document.getElementById('exportPlotHTMLBtn'), // NEW
    exportDataBtn: document.getElementById('exportDataBtn'),
    saveToReportBtn: document.getElementById('saveToReportBtn'),
    generateReportBtn: document.getElementById('generateReportBtn'),
    saveSessionBtn: document.getElementById('saveSessionBtn'), // NEW
    loadSessionInput: document.getElementById('loadSessionInput'), // NEW
    addPlotTabBtn: document.getElementById('addPlotTabBtn'),
    xAxisSelect: document.getElementById('xAxisSelect'),
    leftYAxisSelect: document.getElementById('leftYAxisSelect'),
    rightYAxisSelect: document.getElementById('rightYAxisSelect'),
    groupingSelect: document.getElementById('groupingSelect'),
    chartTypeSelect: document.getElementById('chartTypeSelect'),
    outlierToggle: document.getElementById('outlierToggle'),
    showDataPointsToggle: document.getElementById('showDataPointsToggle'),
    filtersContainer: document.getElementById('filtersContainer'),
    plotDiv: document.getElementById('plot'),
    savedViewsList: document.getElementById('savedViewsList'),
    tabsContainer: document.getElementById('tabsContainer'),
    slopeSection: document.getElementById('slopeSection'),
    slopeResults: document.getElementById('slopeResults'),
    dataPreviewContainer: document.getElementById('dataPreviewContainer'),
    dataStatsContainer: document.getElementById('dataStatsContainer'),
    customizationTabsContainer: document.getElementById('customizationTabsContainer'),
    titlesPanel: document.getElementById('titlesPanel'),
    legendsPanel: document.getElementById('legendsPanel'),
    annotationsPanel: document.getElementById('annotationsPanel'),
    customTitle: document.getElementById('customTitle'),
    axisTitlesContainer: document.getElementById('axisTitlesContainer'),
    legendEditorContainer: document.getElementById('legendEditorContainer'),
    annotationsListContainer: document.getElementById('annotationsListContainer'),
    addAnnotationBtn: document.getElementById('addAnnotationBtn'),
    annotationAxis: document.getElementById('annotationAxis'),
    annotationValue: document.getElementById('annotationValue'),
    annotationText: document.getElementById('annotationText'),
    annotationStyle: document.getElementById('annotationStyle'),
    annotationColor: document.getElementById('annotationColor'),
};

// --- Utility Functions ---
const isNumericArray = (arr) => arr.every(val => val === null || val === undefined || val === '' || !isNaN(val));
const capitalize = (str) => str.charAt(0).toUpperCase() + str.slice(1);

function showNotification(title, message, type = 'success') {
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <div class="notification-icon">
          ${type === 'success' ? `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-green-600"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>` : `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-red-600"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>`}
        </div>
        <div class="notification-content">
            <p class="notification-title">${title}</p>
            <p class="notification-message">${message}</p>
        </div>
    `;
    dom.notificationContainer.appendChild(notification);
    setTimeout(() => {
        notification.remove();
    }, 5000);
}

// --- Core Application Logic ---

function parseTabDelimited(text) {
    const lines = text.trim().split('\n').filter(line => line.trim());
    if (lines.length < 2) throw new Error('[E_PARSE_002]: At least one header row and one data row are required.');
    
    const firstLine = lines[0];
    const delimiter = firstLine.includes('\t') ? '\t' : (firstLine.includes(',') ? ',' : ' ');
    
    headers = lines[0].split(delimiter).map(h => h.trim());
    const dataRows = lines.slice(1);
    
    return dataRows.map(line => {
        const parts = line.split(delimiter);
        const obj = {};
        headers.forEach((h, i) => {
            let val = parts[i] ? parts[i].trim() : '';
            const numVal = Number(val);
            obj[h] = (val === '' || isNaN(numVal)) ? val : numVal;
        });
        return obj;
    });
}

// NEW: Initialize Tom Select controls
function initializeTomSelects() {
    const multiSelectSettings = {
        create: false,
        plugins: ['remove_button']
    };
    
    tomSelects.xAxis = new TomSelect(dom.xAxisSelect, { create: false, maxItems: 1 });
    tomSelects.leftYAxis = new TomSelect(dom.leftYAxisSelect, multiSelectSettings);
    tomSelects.rightYAxis = new TomSelect(dom.rightYAxisSelect, multiSelectSettings);
    tomSelects.grouping = new TomSelect(dom.groupingSelect, multiSelectSettings);

    // Add change listeners
    const onChange = () => saveTabStateAndDraw(activeTabId);
    tomSelects.xAxis.on('change', onChange);
    tomSelects.leftYAxis.on('change', onChange);
    tomSelects.rightYAxis.on('change', onChange);
    tomSelects.grouping.on('change', onChange);
}

// NEW: Populate Tom Select controls
function populateTomSelects(headers) {
    const options = headers.map(h => ({ value: h, text: h }));
    
    Object.values(tomSelects).forEach(select => {
        const value = select.getValue(); // Save current value
        select.clearOptions();
        select.addOptions(options);
        select.setValue(value, true); // Restore value silently
    });
}

function updateDataPreview() {
  if (!parsedData.length) {
    dom.dataPreviewContainer.innerHTML = '<div class="empty-state">Load data to see a preview.</div>';
    return;
  }

  const previewData = parsedData.slice(0, 10);
  let html = `
    <div class="preview-table-wrapper">
      <table class="preview-table">
        <thead>
          <tr>
            ${headers.map(h => `<th>${h}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
          ${previewData.map(row => `
            <tr>
              ${headers.map(h => `<td>${row[h] || ''}</td>`).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
    <p class="text-xs text-slate-400 text-center mt-2">
      Showing first ${previewData.length} of ${parsedData.length} total rows
    </p>
  `;
  dom.dataPreviewContainer.innerHTML = html;
}

function updateDataStats() {
  if (!parsedData.length) {
    dom.dataStatsContainer.innerHTML = '<div class="empty-state">Load data to see statistics.</div>';
    return;
  }
  
  let statsHtml = '';
  headers.forEach(header => {
    const colData = parsedData.map(d => d[header]).filter(v => v !== '' && v !== null && v !== undefined);
    if (!colData.length) return;
    
    const isNumeric = isNumericArray(colData);
    
    if (isNumeric) {
      const numbers = colData.map(v => Number(v));
      const mean = numbers.reduce((a, b) => a + b, 0) / numbers.length;
      const min = Math.min(...numbers);
      const max = Math.max(...numbers);
      
      statsHtml += `
        <div class="stat-card">
          <h4 class="font-semibold text-sm text-slate-200">${header}</h4>
          <p class="text-xs text-slate-400 mb-1">Numeric</p>
          <div class="text-sm">
            <div>Mean: <strong>${mean.toFixed(2)}</strong></div>
            <div>Min: <strong>${min}</strong></div>
            <div>Max: <strong>${max}</strong></div>
          </div>
        </div>
      `;
    } else {
      const unique = [...new Set(colData)];
      statsHtml += `
        <div class="stat-card">
          <h4 class="font-semibold text-sm text-slate-200">${header}</h4>
          <p class="text-xs text-slate-400 mb-1">Categorical</p>
          <div class="text-sm">
            <div>Unique: <strong>${unique.length}</strong></div>
            <div class="truncate">e.g.: <strong>${unique[0]}</strong></div>
          </div>
        </div>
      `;
    }
  });
  
  dom.dataStatsContainer.innerHTML = `<div class="stats-grid">${statsHtml}</div>`;
}

function buildFilters() {
    dom.filtersContainer.innerHTML = '';
    if (!parsedData.length) {
        dom.filtersContainer.innerHTML = '<div class="empty-state">Load data to see available filters.</div>';
        return;
    }

    headers.forEach(h => {
        const colData = parsedData.map(d => d[h]).filter(v => v !== '' && v !== null && v !== undefined);
        if (colData.length === 0) return;

        const groupDiv = document.createElement('div');
        groupDiv.className = 'mb-3';

        if (isNumericArray(colData)) {
            const min = Math.min(...colData);
            const max = Math.max(...colData);
            groupDiv.innerHTML = `
                <label class="font-semibold text-sm">${h} (min/max):</label>
                <div class="flex gap-2 mt-1">
                    <input type="number" step="any" id="filter_min_${h}" placeholder="Min" class="form-input w-full p-1 border rounded-md" value="${min}" />
                    <input type="number" step="any" id="filter_max_${h}" placeholder="Max" class="form-input w-full p-1 border rounded-md" value="${max}" />
                </div>`;
        } else {
            const uniqueValues = [...new Set(colData)].sort();
            groupDiv.innerHTML = `
                <label class="font-semibold text-sm">${h} (select categories):</label>
                <select multiple id="filter_cat_${h}" class="form-select w-full p-1 mt-1 border rounded-md" size="3">
                    ${uniqueValues.map(val => `<option value="${val}" selected>${val}</option>`).join('')}
                </select>`;
        }
        dom.filtersContainer.appendChild(groupDiv);
    });

    dom.filtersContainer.querySelectorAll('input, select').forEach(input => {
        input.addEventListener('change', () => saveTabStateAndDraw(activeTabId));
    });
}

function filterData(data) {
    let filtered = [...data];
    headers.forEach(h => {
        const colData = data.map(d => d[h]).filter(v => v !== '' && v !== null && v !== undefined);
        if (!colData.length) return;

        if (isNumericArray(colData)) {
            const minInput = document.getElementById(`filter_min_${h}`);
            const maxInput = document.getElementById(`filter_max_${h}`);
            if (!minInput || !maxInput || minInput.value === '' || maxInput.value === '') return;
            const minVal = parseFloat(minInput.value);
            const maxVal = parseFloat(maxInput.value);
            filtered = filtered.filter(d => {
                const val = d[h];
                return val >= minVal && val <= maxVal;
            });
        } else {
            const sel = document.getElementById(`filter_cat_${h}`);
            if (!sel) return;
            const selectedOptions = Array.from(sel.selectedOptions).map(o => o.value);
            if (selectedOptions.length > 0 && selectedOptions.length < sel.options.length) {
                filtered = filtered.filter(d => selectedOptions.includes(d[h]));
            }
        }
    });
    return filtered;
}

function getRandomColor(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) hash = str.charCodeAt(i) + ((hash << 5) - hash);
    const c = (hash & 0x00FFFFFF).toString(16).toUpperCase();
    return `#${"00000".substring(0, 6 - c.length)}${c}`;
}

// --- Tab Management ---

function createNewTab(name) {
    return {
        id: Date.now() + Math.random(),
        name: name || `Plot ${plotTabs.length + 1}`,
        xAxis: null, leftYAxes: [], rightYAxes: [], grouping: [],
        chartType: 'box', showOutliers: true, showDataPoints: false,
        filters: {}, plotTraces: [], plotLayout: {},
        legendStyles: {},
        annotations: [],
        customTitles: { title: '', xaxis: '' }
    };
}

function renderTabs() {
    dom.tabsContainer.innerHTML = '';
    plotTabs.forEach(tab => {
        const tabEl = document.createElement('div');
        tabEl.className = 'tab' + (tab.id === activeTabId ? ' active' : '');
        tabEl.textContent = tab.name;
        tabEl.title = 'Click to activate. Right-click to rename. Double-click to delete.';
        
        tabEl.onclick = () => activateTab(tab.id);
        tabEl.oncontextmenu = e => {
            e.preventDefault();
            const newName = prompt('Rename tab:', tab.name);
            if (newName && newName.trim()) {
                tab.name = newName.trim();
                renderTabs();
            }
        };
        tabEl.ondblclick = e => {
            e.preventDefault();
            if (plotTabs.length <= 1) {
                showNotification('Action Denied', '[E_STATE_003]: At least one plot tab is required.', 'error');
                return;
            }
            if (confirm(`Are you sure you want to delete the tab "${tab.name}"?`)) {
                const idxToRemove = plotTabs.findIndex(t => t.id === tab.id);
                if (idxToRemove !== -1) {
                    plotTabs.splice(idxToRemove, 1);
                    if (activeTabId === tab.id) {
                        activateTab(plotTabs[0].id);
                    } else {
                        renderTabs();
                    }
                }
            }
        };
        dom.tabsContainer.appendChild(tabEl);
    });
}

function activateTab(tabId) {
    const tab = plotTabs.find(t => t.id === tabId);
    if (!tab) return;
    activeTabId = tabId;

    // Smartly set X-axis if it's null
    if (!tab.xAxis && headers.length > 0) {
        const categoricalHeaders = headers.filter(h => !isNumericArray(parsedData.map(d => d[h])));
        tab.xAxis = categoricalHeaders[0] || headers[0];
    }
    
    // Smartly set Y-axis if it's null
    if (tab.leftYAxes.length === 0 && headers.length > 0) {
        const numericHeaders = headers.filter(h => isNumericArray(parsedData.map(d => d[h])));
        if (numericHeaders.length > 0) {
            tab.leftYAxes = [numericHeaders[0]];
        }
    }
    
    // UPDATED: Set values using Tom Select instances (silently)
    tomSelects.xAxis.setValue(tab.xAxis, true);
    tomSelects.leftYAxis.setValue(tab.leftYAxes, true);
    tomSelects.rightYAxis.setValue(tab.rightYAxes, true);
    tomSelects.grouping.setValue(tab.grouping, true);

    dom.chartTypeSelect.value = tab.chartType;
    dom.outlierToggle.checked = tab.showOutliers;
    dom.showDataPointsToggle.checked = tab.showDataPoints;

    buildFilters(); 
    for (const h in tab.filters) {
        const val = tab.filters[h];
        if (val === null || val === undefined) continue;
        if (typeof val === 'object' && 'min' in val) {
            const minInput = document.getElementById(`filter_min_${h}`);
            const maxInput = document.getElementById(`filter_max_${h}`);
            if (minInput && maxInput) {
                minInput.value = val.min;
                maxInput.value = val.max;
            }
        } else if (Array.isArray(val)) {
            const sel = document.getElementById(`filter_cat_${h}`);
            if (sel) {
                Array.from(sel.options).forEach(opt => { opt.selected = val.includes(opt.value); });
            }
        }
    }
    
    drawPlot();
    renderTabs();
}

function saveTabStateAndDraw(tabId) {
    const tab = plotTabs.find(t => t.id === tabId);
    if (!tab) return;

    // UPDATED: Get values from Tom Select instances
    tab.xAxis = tomSelects.xAxis.getValue();
    tab.leftYAxes = tomSelects.leftYAxis.getValue();
    tab.rightYAxes = tomSelects.rightYAxis.getValue();
    tab.grouping = tomSelects.grouping.getValue();

    tab.chartType = dom.chartTypeSelect.value;
    tab.showOutliers = dom.outlierToggle.checked;
    tab.showDataPoints = dom.showDataPointsToggle.checked;

    
    tab.filters = {};
    headers.forEach(h => {
        const minInput = document.getElementById(`filter_min_${h}`);
        const maxInput = document.getElementById(`filter_max_${h}`);
        const sel = document.getElementById(`filter_cat_${h}`);
        if (minInput && maxInput && minInput.value !== '' && maxInput.value !== '') {
            tab.filters[h] = { min: parseFloat(minInput.value), max: parseFloat(maxInput.value) };
        } else if (sel) {
            tab.filters[h] = Array.from(sel.selectedOptions).map(o => o.value);
        }
    });
    
    tab.customTitles.title = dom.customTitle.value;
    tab.customTitles.xaxis = document.getElementById('customXAxisTitle')?.value || '';
    
    let yAxisCounter = 1;
    tab.leftYAxes.forEach(() => {
        const axisId = yAxisCounter === 1 ? 'y' : `y${yAxisCounter}`;
        const axisKey = `${axisId}axis`;
        tab.customTitles[axisKey] = document.getElementById(`custom_${axisKey}_Title`)?.value || '';
        yAxisCounter += 2;
    });
    yAxisCounter = 2;
    tab.rightYAxes.forEach(() => {
        const axisId = `y${yAxisCounter}`;
        const axisKey = `${axisId}axis`;
        tab.customTitles[axisKey] = document.getElementById(`custom_${axisKey}_Title`)?.value || '';
        yAxisCounter += 2;
    });

    drawPlot(tab);
}

// NEW: Function to toggle UI elements based on chart type
function updateUIForChartType(chartType) {
    const isBoxPlot = chartType === 'box';
    const isBoxLike = isBoxPlot || chartType === 'violin';
    const isHistogram = chartType === 'histogram';

    document.getElementById('outlierToggleContainer').style.display = isBoxPlot ? 'flex' : 'none';
    document.getElementById('dataPointsToggleContainer').style.display = isBoxLike ? 'flex' : 'none';

    // --- NEW Histogram Logic ---
    if (isHistogram) {
        tomSelects.xAxis.disable();
        tomSelects.xAxis.clear(true); // silent clear
        tomSelects.grouping.disable();
        tomSelects.grouping.clear(true); // silent clear
        tomSelects.rightYAxis.disable();
        tomSelects.rightYAxis.clear(true); // silent clear
        
        // Modify Left Y-Axis to only allow 1 selection
        tomSelects.leftYAxis.settings.maxItems = 1;
        tomSelects.leftYAxis.refreshItems();
        if (tomSelects.leftYAxis.getValue().length > 1) {
            tomSelects.leftYAxis.setValue(tomSelects.leftYAxis.getValue().slice(0, 1), true);
        }
        
    } else {
        // Re-enable for other plot types
        tomSelects.xAxis.enable();
        tomSelects.grouping.enable();
        tomSelects.rightYAxis.enable();
        
        // Restore multi-select for Left Y-Axis
        tomSelects.leftYAxis.settings.maxItems = null; // null means unlimited
        tomSelects.leftYAxis.refreshItems();
    }
}
    
// --- Plotting ---

function buildTracesAndLayout(data, xKey, leftYKeys, rightYKeys, groupingKeys, chartType, showOutliers, showDataPoints, legendStyles) {
    const traces = [];
    const layout = {};
    const defaultSymbols = ['circle', 'square', 'diamond', 'cross', 'x', 'triangle-up', 'star', 'hexagon', 'pentagon'];
    
    // --- NEW: Histogram Logic ---
    if (chartType === 'histogram') {
        const yKey = leftYKeys[0]; // We enforced only one
        if (!yKey) return { traces: [], layoutYAxes: {} };
        
        traces.push({
            x: data.map(d => d[yKey]),
            type: 'histogram',
            name: yKey,
            marker: { color: 'var(--text-accent)' }
        });
        
        layout.barmode = 'overlay';
        layout.xaxis = { title: { text: yKey } };
        layout.yaxis = { title: { text: 'Count' } };
        return { traces, layoutYAxes: layout };
    }
    
    // --- Existing logic for other charts ---
    const xIsCategorical = !isNumericArray(data.map(d => d[xKey]));

    const groups = {};
    if (groupingKeys.length === 0) {
        groups['All Data'] = data;
    } else {
        data.forEach(d => {
            const groupName = groupingKeys.map(k => (d[k] ?? 'Unknown')).join(', ');
            if (!groups[groupName]) groups[groupName] = [];
            groups[groupName].push(d);
        });
    }
    
    const groupNames = Object.keys(groups);
    const numGroups = groupNames.length;
    const xCategories = xIsCategorical ? [...new Set(data.map(d => d[xKey]))].sort() : [];
    const xMap = {};
    xCategories.forEach((cat, i) => xMap[cat] = i);

    const processYKey = (yKey, axisId, axisSide, axisPosition) => {
        layout[axisId === 'y' ? 'yaxis' : `yaxis${axisId.substring(1)}`] = {
            title: { text: yKey },
            overlaying: axisId === 'y' ? undefined : 'y',
            side: axisSide,
            position: axisPosition,
            zeroline: false
        };
        
        groupNames.forEach((groupName, groupIdx) => {
            const groupData = groups[groupName];
            let legendKey = `${groupName} - ${yKey}`;
            const customStyle = legendStyles[legendKey] || {};
            const symbol = customStyle.symbol || defaultSymbols[groupIdx % defaultSymbols.length];
            
            if ((chartType === 'box' || chartType === 'violin') && showDataPoints) {
                const validYData = groupData.map(d => d[yKey]).filter(y => y !== null && y !== undefined && y !== '' && !isNaN(y));
                const pointCount = validYData.length;
                legendKey += ` (n=${pointCount})`;
            }
            
            const traceBase = {
                name: legendKey,
                yaxis: axisId,
                marker: { 
                    color: customStyle.color || getRandomColor(legendKey),
                    symbol: symbol,
                }
            };
            
            if (chartType === 'box' || chartType === 'violin') {
                traces.push({ 
                    ...traceBase, 
                    x: groupData.map(d => d[xKey]), 
                    y: groupData.map(d => d[yKey]), 
                    type: chartType, 
                    boxpoints: chartType === 'box' ? (showOutliers ? 'outliers' : false) : (showOutliers ? 'all' : 'outliers'), // Show all points for violin
                    pointpos: 0 // Center points for violin
                });

            } else if (chartType === 'bar') {
                const xGrouped = {};
                groupData.forEach(d => {
                    const xVal = d[xKey] ?? 'Unknown';
                    if (!xGrouped[xVal]) xGrouped[xVal] = [];
                    if(!isNaN(d[yKey])) xGrouped[xVal].push(d[yKey]);
                });
                const xs = Object.keys(xGrouped).sort();
                const ys = xs.map(x => xGrouped[x].length > 0 ? xGrouped[x].reduce((a,b) => a+b, 0) / xGrouped[x].length : 0);
                traces.push({ ...traceBase, x: xs, y: ys, type: 'bar' });
            } else if (chartType === 'scatter' || chartType === 'line') {
                let xVals = groupData.map(d => d[xKey]);
                let yVals = groupData.map(d => d[yKey]);
                
                if (chartType === 'line' && !xIsCategorical) { // Sort line plot data by X if X is numeric
                    const sortedPairs = groupData.map(d => [d[xKey], d[yKey]])
                        .filter(p => p[0] !== null && p[1] !== null && !isNaN(p[0]))
                        .sort((a, b) => a[0] - b[0]);
                    xVals = sortedPairs.map(p => p[0]);
                    yVals = sortedPairs.map(p => p[1]);
                }

                if (chartType === 'scatter' && numGroups > 1 && xIsCategorical) {
                    const jitterWidth = 0.8; 
                    const groupWidth = jitterWidth / numGroups;
                    const offset = (groupIdx - (numGroups - 1) / 2) * groupWidth;
                    xVals = groupData.map(d => xMap[d[xKey]] + offset);
                }
                
                traces.push({ ...traceBase, x: xVals, y: yVals, mode: chartType === 'scatter' ? 'markers' : 'lines+markers', type: 'scatter' });
            }
        });
    };

    // Left axes: y, y3, y5, ...
    leftYKeys.forEach((yKey, i) => {
        const axisNum = i * 2 + 1;
        const axisId = axisNum === 1 ? 'y' : `y${axisNum}`;
        const position = 0 + (i * 0.08); // Spread axes
        processYKey(yKey, axisId, 'left', position);
    });

    // Right axes: y2, y4, y6, ...
    rightYKeys.forEach((yKey, i) => {
        const axisNum = (i + 1) * 2;
        const axisId = `y${axisNum}`;
        const position = 1 - (i * 0.08); // Spread axes
        processYKey(yKey, axisId, 'right', position);
    });
    
    if(leftYKeys.length === 0 && rightYKeys.length > 0) {
        // Promote first right axis to be the main Y axis
        layout.yaxis = { ...layout.yaxis2, title: { text: '' }, showticklabels: false, side: 'left' };
        delete layout.yaxis2;
        traces.forEach(t => { if(t.yaxis === 'y2') t.yaxis = 'y'; });
    }
    
    if (chartType === 'scatter' && numGroups > 1 && xIsCategorical) {
        layout.xaxis = {
            tickvals: xCategories.map((_, i) => i),
            ticktext: xCategories
        };
    }

    return { traces, layoutYAxes: layout };
}

function drawPlot(tab) {
    if (!parsedData.length) {
        Plotly.purge(dom.plotDiv);
        return;
    }
    tab = tab || plotTabs.find(t => t.id === activeTabId);
    if (!tab) return;
    
    // NEW: Update UI based on chart type
    updateUIForChartType(tab.chartType);
    
    const xKey = tab.xAxis;
    const leftYKeys = tab.leftYAxes;
    const rightYKeys = tab.rightYAxes;

    if (!xKey && tab.chartType !== 'histogram') {
        Plotly.purge(dom.plotDiv);
        dom.plotDiv.innerHTML = `<div class="empty-state">Please select at least one X and one Y axis to generate a plot.</div>`;
        return;
    }
    if (leftYKeys.length === 0 && rightYKeys.length === 0) {
         Plotly.purge(dom.plotDiv);
        dom.plotDiv.innerHTML = `<div class="empty-state">Please select at least one Y axis to generate a plot.</div>`;
        return;
    }

    const filteredData = filterData(parsedData);
    if (!filteredData.length) {
        Plotly.purge(dom.plotDiv);
        dom.plotDiv.innerHTML = `<div class="empty-state">Your current filters result in no data to display. Please adjust your filters.</div>`;
        return;
    }

    const { traces, layoutYAxes } = buildTracesAndLayout(filteredData, xKey, leftYKeys, rightYKeys, tab.grouping, tab.chartType, tab.showOutliers, tab.showDataPoints, tab.legendStyles);
    
    const isDark = currentTheme === 'dark';
    const bgColor = isDark ? '#0f172a' : '#ffffff';
    const textColor = isDark ? '#e2e8f0' : '#1e293b';
    const gridColor = isDark ? '#334155' : '#e2e8f0';

    const layout = {
        title: { text: `Title`, font: { size: 18, family: 'Inter, sans-serif', color: textColor } },
        xaxis: { title: { text: xKey }, tickangle: -45, domain: [0.1 * leftYKeys.length, 1 - 0.1 * rightYKeys.length], color: textColor, gridcolor: gridColor, linecolor: gridColor },
        margin: { t: 50, b: 120, l: 60, r: 60 },
        height: 600,
        boxmode: ['box', 'violin', 'bar'].includes(tab.chartType) ? 'group' : undefined,
        legend: { orientation: 'h', y: -0.3, x: 0.5, xanchor: 'center', font: { color: textColor } },
        plot_bgcolor: bgColor,
        paper_bgcolor: bgColor
    };
    Object.assign(layout, layoutYAxes);
    
    // Apply dark mode to all axes
    Object.keys(layout).filter(k => k.startsWith('yaxis') || k.startsWith('xaxis')).forEach(axisKey => {
        layout[axisKey].color = textColor;
        layout[axisKey].gridcolor = gridColor;
        layout[axisKey].linecolor = gridColor;
    });
    
    const allYKeys = [...leftYKeys, ...rightYKeys];
    
    let defaultTitle = '';
    if (tab.chartType === 'histogram') {
        defaultTitle = `Histogram of ${leftYKeys[0]}`;
    } else {
        defaultTitle = `${capitalize(tab.chartType)} Plot: ${allYKeys.join(', ')} vs ${xKey}`;
    }

    layout.title.text = tab.customTitles.title || defaultTitle;
    
    if (layout.xaxis.title) {
        layout.xaxis.title.text = tab.customTitles.xaxis || layout.xaxis.title.text || xKey;
    }
    
    Object.keys(layout).filter(k => k.startsWith('yaxis')).forEach(axisKey => {
        if (!layout[axisKey].title) return;
        const axisName = axisKey.replace('yaxis', 'y');
        const customTitleKey = `${axisName}axis`;
        const originalTitle = layout[axisKey].title.text;
        if (tab.customTitles[customTitleKey] !== undefined) {
            layout[axisKey].title.text = tab.customTitles[customTitleKey] || originalTitle;
        }
    });
    
    layout.shapes = tab.annotations.map(a => ({
        type: 'line',
        x0: a.axis === 'y' ? 0 : a.value, y0: a.axis === 'x' ? 0 : a.value,
        x1: a.axis === 'y' ? 1 : a.value, y1: a.axis === 'x' ? 1 : a.value,
        xref: a.axis === 'y' ? 'paper' : 'x', yref: a.axis === 'x' ? 'paper' : 'y',
        line: { color: a.color, width: 2, dash: a.style }
    }));

    const manualAnnotations = tab.annotations.filter(a => a.text).map(a => ({
        x: a.axis === 'y' ? 0.98 : a.value, y: a.axis === 'x' ? 0.98 : a.value,
        xref: a.axis === 'y' ? 'paper' : 'x', yref: a.axis === 'x' ? 'paper' : 'y',
        text: a.text, showarrow: false, xanchor: 'right', yanchor: 'top', font: { color: a.color }
    }));
    
    layout.annotations = manualAnnotations;

    const xIsNum = isNumericArray(filteredData.map(d => d[xKey]));
    const yIsNum = allYKeys.every(yKey => isNumericArray(filteredData.map(d => d[yKey])));
    const isSlopeEligible = (tab.chartType === 'line' || tab.chartType === 'scatter') && xIsNum && yIsNum;

    layout.dragmode = isSlopeEligible ? 'select' : 'zoom';
    dom.slopeSection.classList.toggle('hidden', !isSlopeEligible);
    if(isSlopeEligible) dom.slopeResults.innerHTML = '<div class="empty-state">No range selected yet.</div>';

    tab.plotTraces = traces;
    tab.plotLayout = layout;

    try {
        Plotly.newPlot(dom.plotDiv, traces, layout, { responsive: true, displaylogo: false });
        renderCustomizationPanel(tab);
    } catch (e) {
        const errorCode = '[E_PLOT_001]';
        console.error(`${errorCode}: Plotly failed to render.`, e.stack, traces, layout);
        showNotification('Plotting Error', `${errorCode}: ${e.message}`, 'error');
        dom.plotDiv.innerHTML = `<div class="empty-state">${errorCode}: Plotly failed to render. Check console for details.</div>`;
    }
    
    dom.plotDiv.removeAllListeners && dom.plotDiv.removeAllListeners('plotly_selected');
    dom.plotDiv.removeAllListeners && dom.plotDiv.removeAllListeners('plotly_deselect');
    
    if (isSlopeEligible) {
        dom.plotDiv.on('plotly_selected', handleSelection);
        dom.plotDiv.on('plotly_deselect', () => {
            dom.slopeResults.innerHTML = '<div class="empty-state">Selection cleared.</div>';
        });
    }
}
    
// --- Customization Panel Logic ---
function renderCustomizationPanel(tab) {
    renderTitlesEditor(tab);
    renderLegendEditor(tab);
    renderAnnotationsEditor(tab);
}

function renderTitlesEditor(tab) {
    dom.customTitle.value = tab.customTitles.title || '';
    dom.axisTitlesContainer.innerHTML = '';
    
    // X-Axis
    const xAxisDiv = document.createElement('div');
    const xKey = (tab.chartType === 'histogram') ? (tab.leftYAxes[0] || 'Value') : tab.xAxis;
    xAxisDiv.innerHTML = `<label for="customXAxisTitle" class="block font-medium mb-1 text-sm">X-Axis Title:</label>
                          <input type="text" id="customXAxisTitle" class="form-input" value="${tab.customTitles.xaxis || ''}" placeholder="${xKey}">`;
    dom.axisTitlesContainer.appendChild(xAxisDiv);
    
    // Y-Axis
    if (tab.chartType === 'histogram') {
        const yAxisDiv = document.createElement('div');
        yAxisDiv.innerHTML = `<label for="custom_yaxis_Title" class="block font-medium mb-1 text-sm">Y-Axis Title:</label>
                              <input type="text" id="custom_yaxis_Title" data-axis-key="yaxis" class="form-input" value="${tab.customTitles.yaxis || ''}" placeholder="Count">`;
        dom.axisTitlesContainer.appendChild(yAxisDiv);
    } else {
        let yAxisCounter = 1;
        tab.leftYAxes.forEach((y) => {
            const axisId = yAxisCounter === 1 ? 'y' : `y${yAxisCounter}`;
            const axisKey = `${axisId}axis`;
            const yAxisDiv = document.createElement('div');
            yAxisDiv.innerHTML = `<label for="custom_${axisKey}_Title" class="block font-medium mb-1 text-sm">Left Y-Axis Title (${y}):</label>
                                  <input type="text" id="custom_${axisKey}_Title" data-axis-key="${axisKey}" class="form-input" value="${tab.customTitles[axisKey] || ''}" placeholder="${y}">`;
            dom.axisTitlesContainer.appendChild(yAxisDiv);
            yAxisCounter += 2;
        });

        yAxisCounter = 2;
        tab.rightYAxes.forEach((y) => {
            const axisId = `y${yAxisCounter}`;
            const axisKey = `${axisId}axis`;
            const yAxisDiv = document.createElement('div');
            yAxisDiv.innerHTML = `<label for="custom_${axisKey}_Title" class="block font-medium mb-1 text-sm">Right Y-Axis Title (${y}):</label>
                                  <input type="text" id="custom_${axisKey}_Title" data-axis-key="${axisKey}" class="form-input" value="${tab.customTitles[axisKey] || ''}" placeholder="${y}">`;
            dom.axisTitlesContainer.appendChild(yAxisDiv);
            yAxisCounter += 2;
        });
    }
    
    // Add listeners to new inputs
    dom.axisTitlesContainer.querySelectorAll('input').forEach(input => {
        input.addEventListener('input', () => saveTabStateAndDraw(activeTabId));
    });
}

function renderLegendEditor(tab) {
    dom.legendEditorContainer.innerHTML = '';
    if (!tab.plotTraces || tab.plotTraces.length === 0) {
        dom.legendEditorContainer.innerHTML = '<div class="empty-state col-span-full">Plot must be generated to see legends.</div>';
        return;
    }

    const symbolOptions = ['circle', 'square', 'diamond', 'cross', 'x', 'triangle-up', 'star', 'hexagon', 'pentagon'].map(s => `<option value="${s}">${capitalize(s)}</option>`).join('');

    tab.plotTraces.forEach((trace, index) => {
        const legendKey = trace.name;
        if (!legendKey) return; // Skip traces without names (like histograms)
        
        const style = tab.legendStyles[legendKey] || {};
        const currentColor = style.color || (trace.marker ? trace.marker.color : '#000000');
        const currentSymbol = style.symbol || (trace.marker ? trace.marker.symbol : 'circle');

        const editorDiv = document.createElement('div');
        editorDiv.className = 'flex items-center gap-2 text-sm';
        editorDiv.innerHTML = `
            <input type="color" value="${(currentColor || '#000000').slice(0, 7)}" data-trace-index="${index}" class="p-0 h-6 w-6 block bg-white border border-slate-300 dark:border-slate-600 cursor-pointer rounded-md">
            <select data-trace-index="${index}" class="form-select !py-1 text-xs">
              ${symbolOptions}
            </select>
            <span class="truncate" title="${legendKey}">${legendKey}</span>
        `;
        editorDiv.querySelector('select').value = currentSymbol;
        dom.legendEditorContainer.appendChild(editorDiv);
    });
}
    
function renderAnnotationsEditor(tab) {
    dom.annotationsListContainer.innerHTML = '';
    if (tab.annotations.length === 0) {
        dom.annotationsListContainer.innerHTML = '<div class="empty-state">No annotations yet.</div>';
        return;
    }
    tab.annotations.forEach((a, index) => {
        const annoDiv = document.createElement('div');
        annoDiv.className = 'flex items-center justify-between text-sm p-1 rounded hover:bg-slate-100 dark:hover:bg-slate-700';
        annoDiv.innerHTML = `
            <span>${a.axis.toUpperCase()}-axis line at ${a.value} ${a.text ? `(${a.text})` : ''}</span>
            <button data-index="${index}" class="font-bold text-red-500 hover:text-red-700 px-2">&times;</button>
        `;
        dom.annotationsListContainer.appendChild(annoDiv);
    });
}

// --- Slope Calculation ---

function calculateLinReg(xData, yData) {
    const n = xData.length;
    if (n < 2) return { slope: NaN, intercept: NaN, r2: NaN };
    const sumX = xData.reduce((a, b) => a + b, 0);
    const sumY = yData.reduce((a, b) => a + b, 0);
    const sumXY = xData.map((x, i) => x * yData[i]).reduce((a, b) => a + b, 0);
    const sumX2 = xData.map(x => x * x).reduce((a, b) => a + b, 0);
    const sumY2 = yData.map(y => y * y).reduce((a, b) => a + b, 0);
    const denominator = (n * sumX2) - (sumX * sumX);
    if (denominator === 0) return { slope: NaN, intercept: NaN, r2: NaN };
    const slope = ((n * sumXY) - (sumX * sumY)) / denominator;
    const intercept = (sumY / n) - (slope * sumX / n);
    const r2_num = Math.pow((n * sumXY) - (sumX * sumY), 2);
    const r2_den = ((n * sumX2) - Math.pow(sumX, 2)) * ((n * sumY2) - Math.pow(sumY, 2));
    const r2 = r2_den === 0 ? NaN : r2_num / r2_den;
    return { slope, intercept, r2 };
}

function handleSelection(eventData) {
    if (!eventData || !eventData.points || eventData.points.length === 0) {
        dom.slopeResults.innerHTML = '<div class="empty-state">Invalid selection.</div>';
        return;
    }
    dom.slopeResults.innerHTML = '';
    const activeTab = plotTabs.find(t => t.id === activeTabId);
    if (!activeTab) return;

    const pointsByTrace = {};
    eventData.points.forEach(p => {
        if (!pointsByTrace[p.curveNumber]) pointsByTrace[p.curveNumber] = { x: [], y: [] };
        pointsByTrace[p.curveNumber].x.push(p.x);
        pointsByTrace[p.curveNumber].y.push(p.y);
    });

    for (const curveIdx in pointsByTrace) {
        const traceData = pointsByTrace[curveIdx];
        const traceName = activeTab.plotTraces[curveIdx].name || `Trace ${curveIdx}`;
        if (traceData.x.length < 2) continue;

        const { slope, intercept, r2 } = calculateLinReg(traceData.x, traceData.y);
        const resultDiv = document.createElement('div');
        resultDiv.className = 'mb-2 p-2 bg-slate-100 dark:bg-slate-700 rounded shadow-sm';
        resultDiv.innerHTML = `
            <p class="font-semibold text-indigo-800 dark:text-indigo-300">${traceName}:</p>
            <p class="ml-4"><strong>Slope:</strong> ${slope.toFixed(4)}</p>
            <p class="ml-4"><strong>Y-Intercept:</strong> ${intercept.toFixed(4)}</p>
            <p class="ml-4"><strong>R-squared:</strong> ${r2.toFixed(4)}</p>
            <p class="ml-4 text-xs text-slate-500 dark:text-slate-400">(${traceData.x.length} points selected)</p>`;
        dom.slopeResults.appendChild(resultDiv);
    }
    if (dom.slopeResults.innerHTML === '') {
        dom.slopeResults.innerHTML = '<div class="empty-state">Not enough points selected per trace for calculation.</div>';
    }
}

// --- Data Loading and Initialization ---

function loadCurrentData() {
    try {
        parsedData = parseTabDelimited(dom.dataInput.value);
        if (!parsedData.length) throw new Error('[E_PARSE_001]: No data rows found.');
        
        headers = Object.keys(parsedData[0]); // Get headers from first data object
        
        populateTomSelects(headers); // UPDATED
        
        enableControls(true);

        updateDataPreview();
        updateDataStats();

        if (plotTabs.length === 0) {
            const firstTab = createNewTab('Plot 1');
            plotTabs.push(firstTab);
            activateTab(firstTab.id);
        } else {
            // Data has been reloaded, just refresh the current tab
            activateTab(activeTabId);
        }
        
        showNotification('Data Loaded', `${parsedData.length} rows parsed successfully.`, 'success');

    } catch (err) {
        const errorCode = err.message.startsWith('[E_') ? err.message.split(']:')[0] : '[E_PARSE_001]';
        const displayMessage = err.message.startsWith('[E_') ? err.message : `[E_PARSE_001]: Failed to parse data. ${err.message}`;
        console.error(`${errorCode}: ${err.message}`, err.stack);
        showNotification('Parsing Error', displayMessage, 'error');
        enableControls(false);
        Plotly.purge(dom.plotDiv);
    }
}

function enableControls(enabled) {
    const elements = [
        dom.chartTypeSelect, dom.outlierToggle, dom.showDataPointsToggle, 
        dom.exportPlotBtn, dom.exportPlotSVGBtn, dom.exportPlotHTMLBtn, 
        dom.exportDataBtn, dom.saveToReportBtn, dom.addPlotTabBtn,
        dom.saveSessionBtn
    ];
    elements.forEach(el => el.disabled = !enabled);
    
    // UPDATED: Enable/Disable Tom Select
    Object.values(tomSelects).forEach(select => {
        if (enabled) select.enable();
        else select.disable();
    });

    dom.generateReportBtn.disabled = savedReportViews.length === 0;
}

function exportFilteredData() {
  if (!parsedData.length) {
    showNotification('Export Error', '[E_EXPORT_004]: No data to export.', 'error');
    return;
  }
  
  const filtered = filterData(parsedData);
  
  const csvHeader = headers.join(',');
  const csvRows = filtered.map(row => {
    return headers.map(header => {
      let val = row[header];
      if (typeof val === 'string' && val.includes(',')) {
        return `"${val}"`;
      }
      return val;
    }).join(',');
  });
  
  const csvContent = [csvHeader, ...csvRows].join('\n');
  
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  link.setAttribute('href', url);
  link.setAttribute('download', 'filtered_data.csv');
  link.style.visibility = 'hidden';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  
  showNotification('Export Successful', 'Filtered data has been downloaded as a CSV.', 'success');
}

// --- Event Listeners ---

dom.loadDataBtn.addEventListener('click', loadCurrentData);
dom.exportDataBtn.addEventListener('click', exportFilteredData); 

// UPDATED: File input listener for TXT, CSV, TSV, and XLSX
dom.fileInput.addEventListener('change', evt => {
    if (!evt.target.files.length) return;
    const file = evt.target.files[0];
    const reader = new FileReader();
    const fileName = file.name.toLowerCase();

    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        // --- NEW: Handle Excel ---
        reader.onload = e => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                // Convert to CSV. Our parser auto-detects delimiters.
                let csvData = XLSX.utils.sheet_to_csv(worksheet);
                
                dom.dataInput.value = csvData;
                loadCurrentData();
                showNotification('File Loaded', `Imported "${file.name}" successfully.`, 'success');
            } catch (err) {
                const errorCode = '[E_FILE_002]';
                console.error(`${errorCode}: Failed to parse Excel file.`, err.stack);
                showNotification('Excel Parse Error', `${errorCode}: ${err.message}`, 'error');
            }
        };
        reader.onerror = () => {
            const errorCode = '[E_FILE_001]';
            console.error(`${errorCode}: Could not read the selected file.`);
            showNotification('File Error', `${errorCode}: Could not read the selected file.`, 'error');
        };
        reader.readAsArrayBuffer(file); // Read as ArrayBuffer for SheetJS
    } else {
        // --- Original: Handle TXT/CSV/TSV ---
        reader.onload = e => {
            dom.dataInput.value = e.target.result;
            loadCurrentData();
        };
        reader.onerror = () => {
            const errorCode = '[E_FILE_001]';
            console.error(`${errorCode}: Could not read the selected file.`);
            showNotification('File Error', `${errorCode}: Could not read the selected file.`, 'error');
        };
        reader.readAsText(file); // Read as text for plain files
    }
    evt.target.value = '';
});


// Note: Tom Select listeners are added in initializeTomSelects
[dom.chartTypeSelect, dom.outlierToggle, dom.showDataPointsToggle, dom.customTitle].forEach(el => {
    el.addEventListener('change', () => {
        if (el.id === 'chartTypeSelect') {
            updateUIForChartType(el.value); // Update UI immediately
        }
        saveTabStateAndDraw(activeTabId);
    });
});

dom.addPlotTabBtn.addEventListener('click', () => {
  const newTab = createNewTab();
  const activeTab = plotTabs.find(t => t.id === activeTabId);
  if(activeTab) {
    // Deep copy the active tab config
    Object.assign(newTab, JSON.parse(JSON.stringify(activeTab)));
    newTab.id = Date.now() + Math.random(); // Ensure new ID
    newTab.name = `Plot ${plotTabs.length + 1}`; // Ensure new name
    newTab.customTitles.title = ''; // Clear title
  }
  plotTabs.push(newTab);
  activateTab(newTab.id);
});

dom.customizationTabsContainer.addEventListener('click', e => {
    if (e.target.tagName === 'BUTTON') {
        const tabName = e.target.dataset.tab;
        dom.customizationTabsContainer.querySelectorAll('.custom-tab').forEach(t => t.classList.remove('active'));
        e.target.classList.add('active');
        ['titlesPanel', 'legendsPanel', 'annotationsPanel'].forEach(id => {
            dom[id].classList.toggle('hidden', id !== `${tabName}Panel`);
        });
    }
});

dom.titlesPanel.addEventListener('input', e => {
    if (e.target.tagName === 'INPUT') {
        saveTabStateAndDraw(activeTabId);
    }
});

dom.legendEditorContainer.addEventListener('change', e => {
    const activeTab = plotTabs.find(t => t.id === activeTabId);
    if (!activeTab) return;
    const traceIndex = parseInt(e.target.dataset.traceIndex, 10);
    if (isNaN(traceIndex) || !activeTab.plotTraces[traceIndex]) return;
    const legendKey = activeTab.plotTraces[traceIndex].name;
    
    if (!activeTab.legendStyles[legendKey]) activeTab.legendStyles[legendKey] = {};

    const update = {};
    if (e.target.type === 'color') {
        activeTab.legendStyles[legendKey].color = e.target.value;
        update['marker.color'] = e.target.value;
    } else if (e.target.tagName === 'SELECT') {
         activeTab.legendStyles[legendKey].symbol = e.target.value;
         update['marker.symbol'] = e.target.value;
    }
    Plotly.restyle(dom.plotDiv, update, [traceIndex]);
});

dom.addAnnotationBtn.addEventListener('click', () => {
    const activeTab = plotTabs.find(t => t.id === activeTabId);
    if (!activeTab) return;
    const value = parseFloat(dom.annotationValue.value);
    if (isNaN(value)) {
        return showNotification('Invalid Input', '[E_PLOT_002]: Annotation value must be a number.', 'error');
    }
    
    activeTab.annotations.push({
        axis: dom.annotationAxis.value,
        value: value,
        text: dom.annotationText.value,
        style: dom.annotationStyle.value,
        color: dom.annotationColor.value,
    });

    dom.annotationValue.value = '';
    dom.annotationText.value = '';
    drawPlot(activeTab); 
});

dom.annotationsListContainer.addEventListener('click', e => {
    if (e.target.tagName === 'BUTTON') {
         const activeTab = plotTabs.find(t => t.id === activeTabId);
         if (!activeTab) return;
         const index = parseInt(e.target.dataset.index, 10);
         activeTab.annotations.splice(index, 1);
         drawPlot(activeTab);
    }
});

// --- PDF Reporting & Exporting ---
dom.exportPlotBtn.addEventListener('click', () => {
    try {
        const activeTab = plotTabs.find(t => t.id === activeTabId);
        if (!activeTab || !activeTab.plotLayout.title) throw new Error("No active plot to export.");
        const filename = (activeTab.plotLayout.title.text || 'plot').replace(/ /g, '_');
        Plotly.downloadImage(dom.plotDiv, {format: 'png', width: 1200, height: 800, filename: filename});
    } catch (e) {
        const errorCode = '[E_EXPORT_005]';
        console.error(`${errorCode}: PNG export failed.`, e.stack);
        showNotification('PNG Export Error', `${errorCode}: ${e.message}`, 'error');
    }
});

// NEW: SVG Export
dom.exportPlotSVGBtn.addEventListener('click', () => {
    try {
        const activeTab = plotTabs.find(t => t.id === activeTabId);
        if (!activeTab || !activeTab.plotLayout.title) throw new Error("No active plot to export.");
        const filename = (activeTab.plotLayout.title.text || 'plot').replace(/ /g, '_');
        Plotly.downloadImage(dom.plotDiv, {format: 'svg', width: 1200, height: 800, filename: filename});
    } catch (e) {
        const errorCode = '[E_EXPORT_002]';
        console.error(`${errorCode}: SVG export failed.`, e.stack);
        showNotification('SVG Export Error', `${errorCode}: ${e.message}`, 'error');
    }
});

// NEW: HTML Export
dom.exportPlotHTMLBtn.addEventListener('click', () => {
    try {
        const activeTab = plotTabs.find(t => t.id === activeTabId);
        if (!activeTab) throw new Error("No active tab found.");

        const traces = activeTab.plotTraces;
        const layout = activeTab.plotLayout;

        if (!traces || !layout) throw new Error("Plot has not been generated yet.");
        
        const title = layout.title.text || 'Exported Plot';

        // 3. Generate HTML string
        const htmlContent = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Exported Plot: ${title}</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
</head>
<body style="font-family: sans-serif; background-color: ${layout.paper_bgcolor || '#fff'}; color: ${layout.xaxis.color || '#000'}; padding: 1rem;">
    <h1>${title}</h1>
    <div id="plot" style="width:90vw; height:80vh;"></div>
    <script>
        const traces = ${JSON.stringify(traces)};
        const layout = ${JSON.stringify(layout)};
        
        // Make layout responsive for standalone file
        delete layout.height; 
        delete layout.width;

        window.onload = () => {
            Plotly.newPlot('plot', traces, layout, {responsive: true});
        };
    </script>
</body>
</html>
        `;
        
        // 4. Offer as download
        const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', 'exported_plot.html');
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showNotification('Export Successful', 'Standalone HTML file has been generated.', 'success');

    } catch (e) {
        const errorCode = '[E_EXPORT_003]';
        console.error(`${errorCode}: HTML export failed.`, e.stack);
        showNotification('HTML Export Error', `${errorCode}: ${e.message}`, 'error');
    }
});


dom.saveToReportBtn.addEventListener('click', async () => {
  if (!activeTabId) return showNotification('Error', '[E_EXPORT_001]: No active plot tab to save.', 'error');
  const tab = plotTabs.find(t => t.id === activeTabId);
  if (!tab) return;

  try {
    const imgData = await Plotly.toImage(dom.plotDiv, {format:'png', width: 900, height: 600});
    savedReportViews.push({ ...tab, imgData });
    renderSavedViews();
    dom.generateReportBtn.disabled = false;
    showNotification('View Saved', `"${tab.name}" has been added to the report.`, 'success');
  } catch(e) {
    const errorCode = '[E_EXPORT_001]';
    console.error(`${errorCode}: Could not capture plot image.`, e.stack);
    showNotification('Image Error', `${errorCode}: Could not capture plot image: ${e.message}`, 'error');
  }
});

function renderSavedViews() {
    dom.savedViewsList.innerHTML = '';
    if (savedReportViews.length === 0) {
        dom.savedViewsList.innerHTML = '<div class="empty-state">No saved views yet.</div>';
        dom.generateReportBtn.disabled = true;
        return;
    }
    savedReportViews.forEach((view, idx) => {
        const div = document.createElement('div');
        div.className = 'flex items-center justify-between mb-2 gap-2 p-2 rounded-md hover:bg-slate-100 dark:hover:bg-slate-700';
        div.innerHTML = `
            <div class="flex items-center gap-3 flex-grow">
                <input type="checkbox" checked id="viewCheckbox_${idx}" class="h-4 w-4 rounded border-slate-300 dark:border-slate-600 text-indigo-600 focus:ring-indigo-500">
                <label for="viewCheckbox_${idx}" class="text-sm font-medium cursor-pointer">${view.name} - ${view.chartType}</label>
            </div>
            <button title="Remove" class="text-slate-400 hover:text-red-600 transition-colors" data-index="${idx}">&times;</button>
        `;
        div.querySelector('button').onclick = (e) => {
            const indexToRemove = parseInt(e.currentTarget.dataset.index, 10);
            savedReportViews.splice(indexToRemove, 1);
            renderSavedViews();
        };
        dom.savedViewsList.appendChild(div);
    });
}

dom.generateReportBtn.addEventListener('click', () => {
    const selectedViews = savedReportViews.filter((_, idx) => document.getElementById(`viewCheckbox_${idx}`).checked);
    if (selectedViews.length === 0) return showNotification('No Selection', '[E_EXPORT_006]: Please select at least one saved view for the report.', 'error');
    
    try {
        const pdf = new jsPDF({ orientation: 'landscape', unit: 'pt', format: 'a4' });
        const margin = 40;
        const imgWidth = pdf.internal.pageSize.getWidth() - margin * 2;
        const imgHeight = imgWidth * (600 / 900);

        selectedViews.forEach((view, idx) => {
            if (idx > 0) pdf.addPage();
            pdf.setFontSize(16).text(view.plotLayout.title.text, margin, margin - 15);
            pdf.setFontSize(10);

            let axesInfo = [`X-Axis: ${view.xAxis}`];
            if (view.leftYAxes.length > 0) axesInfo.push(`Left Y: ${view.leftYAxes.join(', ')}`);
            if (view.rightYAxes.length > 0) axesInfo.push(`Right Y: ${view.rightYAxes.join(', ')}`);
            pdf.text(axesInfo.join(' | '), margin, margin + 5);

            if (view.grouping.length > 0) pdf.text(`Grouped By: ${view.grouping.join(', ')}`, margin, margin + 20);
            pdf.addImage(view.imgData, 'PNG', margin, margin + 40, imgWidth, imgHeight);
        });
        pdf.save(`DataPlotter_Report_${new Date().toISOString().slice(0, 10)}.pdf`);
    } catch (e) {
        const errorCode = '[E_EXPORT_007]';
        console.error(`${errorCode}: PDF generation failed.`, e.stack);
        showNotification('PDF Error', `${errorCode}: ${e.message}`, 'error');
    }
});

// --- NEW: Session Management ---
dom.saveSessionBtn.addEventListener('click', () => {
    try {
        // Create the state object
        const sessionState = {
            version: 'AIDP_v1.2.0', // Versioning is good practice
            dataInputValue: dom.dataInput.value, // Save the raw input data
            headers: headers,
            plotTabs: plotTabs,
            activeTabId: activeTabId,
            savedReportViews: savedReportViews.map(view => {
                // Don't save imgData, it's huge
                const { imgData, ...rest } = view;
                return rest;
            })
        };
        
        const jsonString = JSON.stringify(sessionState, null, 2);
        const blob = new Blob([jsonString], { type: 'application/json;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', 'plotter_session.json');
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showNotification('Session Saved', 'Your session has been downloaded.', 'success');
        
    } catch (e) {
        const errorCode = '[E_STATE_001]';
        console.error(`${errorCode}: Session save failed.`, e.stack);
        showNotification('Session Save Error', `${errorCode}: ${e.message}`, 'error');
    }
});

dom.loadSessionInput.addEventListener('change', evt => {
    if (!evt.target.files.length) return;
    const file = evt.target.files[0];
    const reader = new FileReader();
    
    reader.onload = e => {
        try {
            const state = JSON.parse(e.target.result);
            
            // --- State Validation ---
            if (state.version !== 'AIDP_v1.2.0' || !state.plotTabs || !state.dataInputValue) {
                throw new Error('Invalid or incompatible session file.');
            }
            
            // --- Restore State ---
            dom.dataInput.value = state.dataInputValue;
            plotTabs = state.plotTabs; // Set tabs *before* loading
            activeTabId = state.activeTabId;
            savedReportViews = state.savedReportViews || [];
            
            // --- Re-initialize ---
            try {
                parsedData = parseTabDelimited(dom.dataInput.value);
                if (!parsedData.length) throw new Error('[E_PARSE_001]: No data rows found in session.');
                
                headers = Object.keys(parsedData[0]); // Re-get headers
                
                populateTomSelects(headers); // Re-populate selects
                enableControls(true);
                updateDataPreview();
                updateDataStats();
                renderSavedViews();
                
                if (!plotTabs.find(t => t.id === activeTabId)) {
                    activeTabId = plotTabs[0].id; // Fallback
                }
                
                activateTab(activeTabId); // This will build filters for the tab and draw plot
                
                showNotification('Session Loaded', 'Your session has been restored.', 'success');

            } catch (err) {
                // This catches parsing error from the session data
                throw new Error(`Failed to re-process data from session. ${err.message}`);
            }

        } catch (err) {
            const errorCode = '[E_STATE_002]';
            console.error(`${errorCode}: Failed to load session.`, err.stack);
            showNotification('Session Load Error', `${errorCode}: ${err.message}`, 'error');
        }
    };
    
    reader.onerror = () => {
        const errorCode = '[E_FILE_003]';
        console.error(`${errorCode}: Could not read session file.`);
        showNotification('File Error', `${errorCode}: Could not read session file.`, 'error');
    };
    reader.readAsText(file);
    evt.target.value = ''; // Clear input
});


// --- Theme Toggle Logic ---
function setTheme(theme) {
    currentTheme = theme;
    // Dark is default, no class toggling needed for now
    // Redraw the active plot with the new theme
    if (activeTabId) {
        drawPlot(plotTabs.find(t => t.id === activeTabId));
    }
}

// --- Initial Load ---
function initialize() {
  const sampleData = `Displacement,Pull Load (N),Temperature (C),Adhesive,Test Case
0.1,150,25,Fuller,RT
0.2,310,25,Fuller,RT
0.3,450,25,Fuller,RT
0.4,620,25,Fuller,RT
0.5,780,25,Fuller,RT
0.6,900,25,Fullf,RT
0.1,120,-40,Sika,Cold
0.2,250,-40,Sika,Cold
0.3,380,-40,Sika,Cold
0.4,510,-40,Sika,Cold
0.5,630,-40,Sika,Cold
0.6,740,-40,Sika,Cold
0.1,180,80,Dow,Hot
0.2,350,80,Dow,Hot
0.3,530,80,Dow,Hot
0.4,710,80,Dow,Hot
0.5,880,80,Dow,Hot
0.6,1050,80,Dow,Hot`;
  dom.dataInput.value = sampleData;
  
  // NEW: Initialize Tom Select
  initializeTomSelects();
  
  // Set theme to dark by default
  setTheme('dark');
  
  // Auto-load the sample data
  loadCurrentData();
}

window.onload = initialize;
