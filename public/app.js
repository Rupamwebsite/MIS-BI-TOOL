/**
 * MIS ANALYTICS PRO - Client Application Logic v3.0
 * Advanced ERP Intelligence Suite
 * Features:
 * - 6 Advanced Interactive Charts (Dual-Axis Combo, Pareto 80/20, Share Donut, Ranked Bars, Channels, Cumulative Growth)
 * - Chart Image Download (PNG) & Type Switching
 * - 100% Reactive Multi-Column & Multi-Operator Advanced Filter Engine
 * - Multi-Sheet Executive Excel Report Download (.xlsx)
 * - Formatted Executive PDF Report Download (.pdf via html2pdf)
 * - Quick Date Presets & Dynamic Metric / Dimension Switchers
 * - Deep Data Grid Explorer with real-time filtered sync
 */

// Application State
const state = {
  currentDatasetId: null,
  currentDataset: null,
  datasetsList: [],
  analytics: null,
  rawGridPage: 1,
  rawGridLimit: 25,
  rawGridTotal: 0,
  rawGridRows: [],
  gridSearch: '',
  charts: {},
  theme: localStorage.getItem('mis_theme') || 'dark',

  // Active Filter & Configuration State
  selectedMetricCol: null,
  selectedDimCol: null,
  quickCategoryFilter: 'All',
  datePreset: 'all',
  dateRange: { start: '', end: '' },
  multiFilters: [], // Array of { col, op, val }

  // Power BI Features State
  selectedVisualType: 'bar',
  currentAskBiResult: null,
  currentBuilderResult: null,
  askBiChartType: 'bar',

  // Doctor MIS Analytics State
  doctorSearchTerm: '',
  doctorDeptFilter: ''
};

// Initialize on DOM load or immediately if already loaded
async function initApp() {
  initTheme();
  checkSystemStatus();
  fetchDatasetsList();
  setupEventListeners();
  setupDropzones();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initApp);
} else {
  initApp();
}

// ==========================================
// THEME MANAGEMENT
// ==========================================
function initTheme() {
  document.documentElement.setAttribute('data-theme', state.theme);
  const themeBtn = document.getElementById('themeToggle');
  if (themeBtn) {
    themeBtn.addEventListener('click', () => {
      state.theme = state.theme === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', state.theme);
      localStorage.setItem('mis_theme', state.theme);
      if (state.analytics) {
        renderCharts(state.analytics);
      }
    });
  }
}

// ==========================================
// SYSTEM STATUS & DATASETS API
// ==========================================
async function checkSystemStatus() {
  try {
    const res = await fetch('/api/system/status');
    const data = await res.json();
    if (data.success) {
      const meta = document.getElementById('dbMeta');
      const badge = document.getElementById('dbStatusBadge');
      if (meta) {
        meta.textContent = `Online • ${data.totalDatasets} set${data.totalDatasets !== 1 ? 's' : ''}`;
      }
      if (badge) {
        badge.title = `MySQL: MIS_DB connected with ${data.totalDatasets} datasets and ${Number(data.totalRowsStored).toLocaleString()} records stored`;
      }
    }
  } catch (err) {
    console.warn('[Status Check]', err);
    const meta = document.getElementById('dbMeta');
    if (meta) meta.textContent = 'Connected';
  }
}

async function fetchDatasetsList() {
  try {
    const res = await fetch('/api/datasets');
    const data = await res.json();
    if (data.success) {
      state.datasetsList = data.datasets;
      renderDatasetDropdown();

      // Check for saved dataset in localStorage for multi-page persistence
      const savedId = localStorage.getItem('mis_active_dataset_id');
      const targetDataset = (savedId && data.datasets.find(d => String(d.id) === String(savedId)))
        ? savedId
        : (data.datasets[0] ? data.datasets[0].id : null);

      if (targetDataset && !state.currentDatasetId) {
        loadDataset(targetDataset);
      } else if (data.datasets.length === 0) {
        showHero(true);
      }
    }
  } catch (err) {
    showToast('Failed to load datasets from MIS_DB: ' + err.message);
  }
}

function renderDatasetDropdown() {
  const select = document.getElementById('activeDatasetSelect');
  if (select) {
    select.innerHTML = '<option value="">Select a Dataset...</option>';
    state.datasetsList.forEach(ds => {
      const opt = document.createElement('option');
      opt.value = ds.id;
      opt.textContent = `${ds.name} (${Number(ds.total_rows).toLocaleString()} rows)`;
      if (state.currentDatasetId == ds.id) opt.selected = true;
      select.appendChild(opt);
    });
  }

  // Populate Corporate Home Datasets List
  const homeContainer = document.getElementById('corpDatasetsListContainer');
  if (homeContainer) {
    if (!state.datasetsList || state.datasetsList.length === 0) {
      homeContainer.innerHTML = '<div style="font-size: 0.8rem; color: var(--text-muted); padding: 0.85rem; text-align: center;">No datasets ingested yet in MIS_DB.<br><small style="color:var(--primary);">Use the Ingestion Console on the left to upload.</small></div>';
    } else {
      homeContainer.innerHTML = state.datasetsList.map(ds => {
        const isActive = state.currentDatasetId == ds.id;
        return `
          <div class="corp-dataset-item ${isActive ? 'active' : ''}">
            <div class="corp-ds-meta-col">
              <span class="corp-ds-title" title="${escapeHtml(ds.name)}">${escapeHtml(ds.name)}</span>
              <span class="corp-ds-sub"><strong>${Number(ds.total_rows).toLocaleString()}</strong> rows • ${ds.total_columns || 0} cols</span>
            </div>
            <div style="display: flex; gap: 0.35rem; align-items: center;">
              <button type="button" class="corp-ds-action-btn" onclick="loadDataset(${ds.id})" title="Load into Active Workspace">
                ${isActive ? '✓ Active' : 'Load'}
              </button>
              <a href="overview.html" class="corp-ds-action-btn" title="View Executive Dashboard" style="text-decoration: none;">
                📊 Open
              </a>
            </div>
          </div>
        `;
      }).join('');
    }
  }

  // Update Corporate Home KPI Badge
  const homeDsCount = document.getElementById('corpHomeDsCount');
  if (homeDsCount) {
    homeDsCount.textContent = `${state.datasetsList ? state.datasetsList.length : 0} Datasets Active`;
  }
}

async function loadDataset(id) {
  if (!id) return;
  try {
    showToast('Loading MIS report from MIS_DB...');
    const res = await fetch(`/api/datasets/${id}`);
    const data = await res.json();

    if (!data.success) {
      throw new Error(data.error || 'Failed to fetch dataset');
    }

    state.currentDatasetId = id;
    state.currentDataset = data.dataset;
    state.analytics = data.dataset.metrics_summary;
    localStorage.setItem('mis_active_dataset_id', id);

    // Reset filters
    state.quickCategoryFilter = 'All';
    state.datePreset = 'all';
    state.dateRange = { start: '', end: '' };
    state.multiFilters = [];

    // Set active metric and dimension defaults
    state.selectedMetricCol = state.analytics.summaryKPIs?.metricColumn || null;
    state.selectedDimCol = state.analytics.summaryKPIs?.dimensionColumn || null;

    // Check for drilldown passed via sessionStorage across pages
    const storedDim = sessionStorage.getItem('mis_drilldown_dim');
    const storedVal = sessionStorage.getItem('mis_drilldown_val');
    if (storedDim && storedVal) {
      sessionStorage.removeItem('mis_drilldown_dim');
      sessionStorage.removeItem('mis_drilldown_val');
      state.selectedDimCol = storedDim;
      state.quickCategoryFilter = storedVal;
    }

    showHero(false);
    renderDatasetDropdown();
    renderMetadata(data.dataset);
    populateMetricAndDimSelectors(state.analytics);
    renderKPIs(state.analytics);
    renderInsights(state.analytics);
    renderCharts(state.analytics);
    renderCategoryTable(state.analytics);
    renderAnomalies(state.analytics);
    populatePivotSelectors(state.analytics);
    populateQuickCategoryFilter(state.analytics);
    renderActiveFilterChips();
    populateVisualBuilderControls();
    renderDimensionPills(state.analytics);
    renderDoctorAnalytics(state.analytics);

    // Trigger tab-specific renders for standalone pages
    if (document.getElementById('pivotContainer') && typeof runPivot === 'function') {
      runPivot();
    }
    if (document.getElementById('visualBuilderStudio') && typeof renderCustomVisualStudio === 'function') {
      if (!state.currentBuilderResult) renderCustomVisualStudio();
    }

    // Reset and load grid
    state.rawGridPage = 1;
    loadRawDataGrid();
    showToast(`Loaded "${data.dataset.name}" successfully!`);
  } catch (err) {
    showToast('Error loading dataset: ' + err.message);
  }
}

function showHero(show) {
  const hero = document.getElementById('welcomeHero');
  const ws = document.getElementById('dashboardWorkspace');
  if (hero && ws) {
    if (show) {
      hero.style.display = 'flex';
      ws.style.display = 'none';
    } else {
      hero.style.display = 'none';
      ws.style.display = 'block';
    }
  }
}

// ==========================================
// METRIC & DIMENSION SWITCHERS
// ==========================================
function populateMetricAndDimSelectors(analytics) {
  const metricSelect = document.getElementById('activeMetricSelect');
  const dimSelect = document.getElementById('activeDimensionSelect');

  if (!metricSelect || !dimSelect || !analytics.columns) return;

  metricSelect.innerHTML = '';
  dimSelect.innerHTML = '';

  const numericCols = analytics.columns.filter(c => c.type === 'numeric');
  const dimCols = analytics.columns.filter(c => c.type === 'category' || c.type === 'text');

  numericCols.forEach(col => {
    const opt = document.createElement('option');
    opt.value = col.name;
    opt.textContent = `${col.name} (${col.stats?.sum ? formatCompactNumber(col.stats.sum) : 'Numeric'})`;
    if (col.name === state.selectedMetricCol) opt.selected = true;
    metricSelect.appendChild(opt);
  });

  dimCols.forEach(col => {
    const opt = document.createElement('option');
    opt.value = col.name;
    opt.textContent = `${col.name} (${col.stats?.uniqueCount || 0} unique)`;
    if (col.name === state.selectedDimCol) opt.selected = true;
    dimSelect.appendChild(opt);
  });
}

// ==========================================
// SAMPLE DATASET LOADER
// ==========================================
async function loadSampleDataset(type) {
  closeModals();
  const title = type === 'sales' ? 'Commercial Sales ERP' : 'Hospital OPD Billing';
  showImportProgress(`${title} (Sample ERP)`, 250000);
  clearImportLogs();

  addImportLog(`Generating pre-configured ${title} dataset...`, 'active');
  updateImportProgress(20, 'Generating realistic ERP transactions...', '280+ bills', 'Initializing...');
  setImportStep(1, 'active');

  let simPct = 20;
  const simTimer = setInterval(() => {
    if (simPct < 85) {
      simPct += 14;
      if (simPct >= 35 && simPct < 60) {
        setImportStep(1, 'completed');
        setImportStep(2, 'active');
        addImportLog('Validating column data types & financial schema...', 'active');
        updateImportProgress(simPct, 'Smart parsing headers & data types...', 'Schema detection', 'In progress...');
      } else if (simPct >= 60) {
        setImportStep(2, 'completed');
        setImportStep(3, 'active');
        addImportLog('Writing dataset transactions to MySQL MIS_DB...', 'active');
        updateImportProgress(simPct, 'Writing dataset to MySQL MIS_DB...', 'Batch insert', 'Saving...');
      }
    }
  }, 220);

  try {
    const res = await fetch('/api/load-sample', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ type })
    });
    const data = await res.json();
    clearInterval(simTimer);

    if (data.success) {
      setImportStep(3, 'completed');
      setImportStep(4, 'completed');
      updateImportProgress(100, `✓ Sample loaded into MIS_DB!`, 'All records stored', 'Ready!');

      const iconWrap = document.getElementById('importIconWrap');
      const fileSvg = iconWrap?.querySelector('.import-file-svg');
      const succSvg = iconWrap?.querySelector('.import-success-svg');
      if (iconWrap) iconWrap.classList.add('success');
      if (fileSvg) fileSvg.style.display = 'none';
      if (succSvg) succSvg.style.display = 'block';

      addImportLog(`✓ Sample "${title}" (${data.analytics?.summaryKPIs?.totalRecords} rows) saved in MySQL MIS_DB.`, 'success');

      // Show Completion Card
      const compCard = document.getElementById('importCompleteCard');
      if (compCard) {
        compCard.style.display = 'block';
        document.getElementById('compTotalRows').textContent = data.analytics?.summaryKPIs?.totalRecords || 0;
        document.getElementById('compTotalCols').textContent = data.analytics?.summaryKPIs?.activeColumns || 0;
        document.getElementById('compDomain').textContent = data.analytics?.detectedDomain || 'Sample';
        document.getElementById('compTime').textContent = 'Instant';
      }

      activeImportDatasetId = data.datasetId;

      const btnGo = document.getElementById('btnGoToDashboard');
      if (btnGo) {
        btnGo.onclick = async () => {
          closeModals();
          localStorage.setItem('mis_active_dataset_id', data.datasetId);
          if (window.location.pathname.endsWith('home.html')) {
            window.location.href = 'overview.html';
          } else {
            await fetchDatasetsList();
            await checkSystemStatus();
            await loadDataset(data.datasetId);
            showToast(`✓ Loaded ${title} into MIS_DB!`);
          }
        };
      }

      setTimeout(async () => {
        if (document.getElementById('importProgressModal')?.style.display === 'flex' && activeImportDatasetId === data.datasetId) {
          closeModals();
          localStorage.setItem('mis_active_dataset_id', data.datasetId);
          if (window.location.pathname.endsWith('home.html')) {
            window.location.href = 'overview.html';
          } else {
            await fetchDatasetsList();
            await checkSystemStatus();
            await loadDataset(data.datasetId);
            showToast(`✓ Loaded ${title} into MIS_DB!`);
          }
        }
      }, 2500);
    } else {
      throw new Error(data.error);
    }
  } catch (err) {
    clearInterval(simTimer);
    handleImportFailure(err.message);
  }
}

// ==========================================
// METADATA & KPIS RENDERING
// ==========================================
function renderMetadata(ds) {
  const title = document.getElementById('datasetTitle');
  const domain = document.getElementById('domainTag');
  const metaRec = document.getElementById('metaRecords');
  const metaCol = document.getElementById('metaColumns');
  const metaSheet = document.getElementById('metaSheetName');
  const metaDate = document.getElementById('metaDateRange');

  if (title) title.textContent = ds.name || ds.filename;
  if (domain) domain.textContent = ds.detected_type || 'Enterprise MIS';
  if (metaRec) metaRec.textContent = `${Number(ds.total_rows).toLocaleString()} Records`;
  if (metaCol) metaCol.textContent = `${ds.total_columns} Columns`;
  if (metaSheet) metaSheet.textContent = `Sheet: ${ds.primary_sheet || 'Sheet1'}`;

  const dateCol = ds.metrics_summary?.columns?.find(c => c.type === 'date');
  if (metaDate) {
    if (dateCol && dateCol.stats && dateCol.stats.startDate) {
      metaDate.textContent = `${dateCol.stats.startDate} to ${dateCol.stats.endDate}`;
    } else {
      metaDate.textContent = 'All Records Period';
    }
  }
}

function renderKPIs(analytics) {
  if (!analytics || !analytics.summaryKPIs) return;
  const kpis = analytics.summaryKPIs;

  const totalValEl = document.getElementById('kpiTotalValue');
  const totalRecEl = document.getElementById('kpiTotalRecords');
  const avgValEl = document.getElementById('kpiAvgValue');
  const maxValEl = document.getElementById('kpiMaxValue');
  const kpiLabel1 = document.getElementById('kpiLabel1');
  const kpiColName1 = document.getElementById('kpiColName1');

  if (kpis.metricColumn) {
    if (kpiLabel1) kpiLabel1.textContent = `Total ${formatColName(kpis.metricColumn)}`;
    if (kpiColName1) kpiColName1.textContent = `Cumulative sum of ${kpis.metricColumn}`;
  }

  if (totalValEl) totalValEl.textContent = formatCurrency(kpis.totalValue);
  if (totalRecEl) totalRecEl.textContent = Number(kpis.totalRecords).toLocaleString();
  if (avgValEl) avgValEl.textContent = formatCurrency(kpis.averageTicket);
  if (maxValEl) maxValEl.textContent = formatCurrency(kpis.highestTransaction);

  const badge = document.getElementById('anomalyCountBadge');
  const anomCount = (analytics.anomalies || []).length;
  if (badge) {
    badge.textContent = anomCount;
    badge.style.display = anomCount > 0 ? 'inline-block' : 'none';
  }
}

function renderInsights(analytics) {
  const container = document.getElementById('insightsList');
  if (!container) return;

  const insights = analytics.insights || [];
  if (insights.length === 0) {
    container.innerHTML = '<p class="kpi-subtext">No distinct anomalies or outliers detected.</p>';
    return;
  }

  container.innerHTML = insights.map(text => {
    const formatted = text.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    return `
      <div class="insight-item">
        <span class="insight-bullet">✦</span>
        <span>${formatted}</span>
      </div>
    `;
  }).join('');
}

// ==========================================
// ADVANCED CHARTS SUITE (Chart.js 4.4)
// ==========================================
function renderCharts(analytics) {
  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const gridColor = isDark ? 'rgba(0, 149, 199, 0.08)' : 'rgba(0, 80, 130, 0.06)';
  const textColor = isDark ? '#8bb5d0' : '#336b8a';

  // 1. Dual-Axis Combo Chart: Financial Volume vs Transaction Count
  destroyChart('overviewTrendChart');
  const trendCanvas = document.getElementById('overviewTrendChart');
  if (trendCanvas && analytics.trendData && analytics.trendData.length > 0) {
    const ctx = trendCanvas.getContext('2d');
    const gradient = ctx.createLinearGradient(0, 0, 0, 300);
    gradient.addColorStop(0, 'rgba(0, 149, 199, 0.80)');
    gradient.addColorStop(1, 'rgba(0, 90, 142, 0.15)');

    state.charts.overviewTrendChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: analytics.trendData.map(d => d.date),
        datasets: [
          {
            type: 'bar',
            label: analytics.summaryKPIs.metricColumn || 'Financial Volume',
            data: analytics.trendData.map(d => d.totalValue),
            backgroundColor: gradient,
            borderRadius: 5,
            yAxisID: 'y',
            order: 2
          },
          {
            type: 'line',
            label: 'Transaction Count',
            data: analytics.trendData.map(d => d.volume),
            borderColor: '#10b981',
            backgroundColor: 'rgba(16, 185, 129, 0.15)',
            borderWidth: 2.5,
            tension: 0.35,
            pointRadius: analytics.trendData.length > 30 ? 0 : 3,
            yAxisID: 'y1',
            order: 1
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: true,
            position: 'top',
            labels: { color: textColor, boxWidth: 12, padding: 12 }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                if (ctx.dataset.yAxisID === 'y') {
                  return ` Value: ${formatCurrency(ctx.parsed.y)}`;
                } else {
                  return ` Volume: ${ctx.parsed.y} transactions`;
                }
              }
            }
          }
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: { color: textColor, maxTicksLimit: 12 }
          },
          y: {
            type: 'linear',
            position: 'left',
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              callback: (v) => formatCompactNumber(v)
            }
          },
          y1: {
            type: 'linear',
            position: 'right',
            grid: { drawOnChartArea: false },
            ticks: { color: '#10b981' }
          }
        }
      }
    });
  }

  // 2. Overview Donut Chart with Center Share
  destroyChart('overviewDonutChart');
  const donutCanvas = document.getElementById('overviewDonutChart');
  if (donutCanvas && analytics.categoryBreakdown && analytics.categoryBreakdown.length > 0) {
    const topCats = analytics.categoryBreakdown.slice(0, 6);
    const otherVal = analytics.categoryBreakdown.slice(6).reduce((acc, c) => acc + c.totalValue, 0);

    const labels = topCats.map(c => c.category);
    const values = topCats.map(c => c.totalValue);
    if (otherVal > 0) {
      labels.push('Others');
      values.push(otherVal);
    }

    const donutColors = ['#0095c7', '#00c896', '#0ea5e9', '#f59e0b', '#38bdf8', '#005a8e', '#64748b'];

    state.charts.overviewDonutChart = new Chart(donutCanvas.getContext('2d'), {
      type: 'doughnut',
      data: {
        labels,
        datasets: [{
          data: values,
          backgroundColor: donutColors,
          borderWidth: 0,
          hoverOffset: 8
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        cutout: '68%',
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const index = elements[0].index;
            const clickedLabel = labels[index];
            if (clickedLabel && clickedLabel !== 'Others') {
              applyQuickCategoryDrilldown(clickedLabel);
            }
          }
        },
        plugins: {
          legend: {
            position: 'bottom',
            labels: { color: textColor, boxWidth: 12, padding: 10 }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => ` ${ctx.label}: ${formatCurrency(ctx.parsed)}`
            }
          }
        }
      }
    });

    const donutTitle = document.getElementById('overviewDonutTitle');
    if (donutTitle && state.selectedDimCol) {
      donutTitle.textContent = `${formatColName(state.selectedDimCol)} Share`;
    }
  }

  // 3. Pareto 80/20 Distribution Chart
  destroyChart('paretoChart');
  const paretoCanvas = document.getElementById('paretoChart');
  if (paretoCanvas && analytics.paretoData && analytics.paretoData.length > 0) {
    const labels = analytics.paretoData.map(p => p.category);
    const revValues = analytics.paretoData.map(p => p.totalValue);
    const cumPctValues = analytics.paretoData.map(p => p.cumPct);

    state.charts.paretoChart = new Chart(paretoCanvas.getContext('2d'), {
      type: 'bar',
      data: {
        labels,
        datasets: [
          {
            type: 'bar',
            label: 'Category Revenue',
            data: revValues,
            backgroundColor: 'rgba(0, 149, 199, 0.75)',
            borderRadius: 6,
            yAxisID: 'y',
            order: 2
          },
          {
            type: 'line',
            label: 'Cumulative %',
            data: cumPctValues,
            borderColor: '#f59e0b',
            borderWidth: 2.5,
            pointBackgroundColor: '#f59e0b',
            pointRadius: 4,
            tension: 0.25,
            yAxisID: 'y1',
            order: 1
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const index = elements[0].index;
            applyQuickCategoryDrilldown(labels[index]);
          }
        },
        plugins: {
          legend: {
            display: true,
            position: 'top',
            labels: { color: textColor, boxWidth: 12 }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                if (ctx.dataset.yAxisID === 'y') {
                  return ` Revenue: ${formatCurrency(ctx.parsed.y)}`;
                } else {
                  return ` Cumulative Share: ${ctx.parsed.y}%`;
                }
              }
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { color: textColor }
          },
          y: {
            type: 'linear',
            position: 'left',
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              callback: (v) => formatCompactNumber(v)
            }
          },
          y1: {
            type: 'linear',
            position: 'right',
            min: 0,
            max: 100,
            grid: { drawOnChartArea: false },
            ticks: {
              color: '#f59e0b',
              callback: (v) => `${v}%`
            }
          }
        }
      }
    });
  }

  // 4. Category Ranked Performance Horizontal Bars
  destroyChart('categoryBarChart');
  const barCanvas = document.getElementById('categoryBarChart');
  if (barCanvas && analytics.categoryBreakdown && analytics.categoryBreakdown.length > 0) {
    const topRanked = analytics.categoryBreakdown.slice(0, 10);
    const labels = topRanked.map(c => c.category);

    state.charts.categoryBarChart = new Chart(barCanvas.getContext('2d'), {
      type: 'bar',
      data: {
        labels,
        datasets: [{
          label: 'Total Value',
          data: topRanked.map(c => c.totalValue),
          backgroundColor: 'rgba(0, 149, 199, 0.75)',
          borderRadius: 6,
          hoverBackgroundColor: '#0095c7'
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        onClick: (evt, elements) => {
          if (elements.length > 0) {
            const index = elements[0].index;
            applyQuickCategoryDrilldown(labels[index]);
          }
        },
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => ` Value: ${formatCurrency(ctx.parsed.x)}`
            }
          }
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              callback: (v) => formatCompactNumber(v)
            }
          },
          y: {
            grid: { display: false },
            ticks: { color: textColor }
          }
        }
      }
    });

    const catBarTitle = document.getElementById('categoryBarTitle');
    if (catBarTitle && state.selectedDimCol) {
      catBarTitle.textContent = `Ranked ${formatColName(state.selectedDimCol)} Performance`;
    }
  }

  // 5. Payment Channel / Status Distribution
  destroyChart('paymentChart');
  const payCanvas = document.getElementById('paymentChart');
  if (payCanvas && analytics.paymentBreakdown && analytics.paymentBreakdown.length > 0) {
    state.charts.paymentChart = new Chart(payCanvas.getContext('2d'), {
      type: 'bar',
      data: {
        labels: analytics.paymentBreakdown.map(p => p.method),
        datasets: [{
          label: 'Revenue / Value',
          data: analytics.paymentBreakdown.map(p => p.totalAmount),
          backgroundColor: 'rgba(16, 185, 129, 0.75)',
          borderRadius: 6,
          hoverBackgroundColor: '#10b981'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => ` Total: ${formatCurrency(ctx.parsed.y)} (${ctx.raw ? Number(ctx.raw).toLocaleString() : ''})`
            }
          }
        },
        scales: {
          x: {
            grid: { display: false },
            ticks: { color: textColor }
          },
          y: {
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              callback: (v) => formatCompactNumber(v)
            }
          }
        }
      }
    });
  }

  // 6. Cumulative Revenue Growth Curve (Run-Rate)
  destroyChart('growthTrendChart');
  const growthCanvas = document.getElementById('growthTrendChart');
  if (growthCanvas && analytics.trendData && analytics.trendData.length > 0) {
    let runningTotal = 0;
    const growthLabels = [];
    const growthValues = [];

    analytics.trendData.forEach(d => {
      runningTotal += d.totalValue;
      growthLabels.push(d.date);
      growthValues.push(Math.round(runningTotal * 100) / 100);
    });

    const ctx = growthCanvas.getContext('2d');
    const grad = ctx.createLinearGradient(0, 0, 0, 300);
    grad.addColorStop(0, 'rgba(14, 165, 233, 0.45)');
    grad.addColorStop(1, 'rgba(14, 165, 233, 0.0)');

    state.charts.growthTrendChart = new Chart(ctx, {
      type: 'line',
      data: {
        labels: growthLabels,
        datasets: [{
          label: 'Cumulative Total',
          data: growthValues,
          borderColor: '#0ea5e9',
          backgroundColor: grad,
          fill: true,
          tension: 0.35,
          borderWidth: 2.5,
          pointRadius: growthLabels.length > 30 ? 0 : 3
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => ` Cumulative: ${formatCurrency(ctx.parsed.y)}`
            }
          }
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: { color: textColor, maxTicksLimit: 10 }
          },
          y: {
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              callback: (v) => formatCompactNumber(v)
            }
          }
        }
      }
    });
  }
}

function destroyChart(name) {
  if (state.charts[name]) {
    state.charts[name].destroy();
    delete state.charts[name];
  }
}

// Download high-resolution PNG image of any chart
function downloadChartImage(canvasId, filename = 'chart.png') {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const link = document.createElement('a');
  link.download = filename;
  link.href = canvas.toDataURL('image/png', 1.0);
  link.click();
  showToast(`Downloaded chart as ${filename}`);
}

// Drill-down clicking helper
function applyQuickCategoryDrilldown(categoryName) {
  state.quickCategoryFilter = categoryName;
  const dimSelect = document.getElementById('filterDimension');
  if (dimSelect) dimSelect.value = categoryName;
  triggerAdvancedFilterRecompute();
}

// ==========================================
// TABLES (Category Breakdown & Anomalies)
// ==========================================
function renderCategoryTable(analytics) {
  const tbody = document.getElementById('categoryTableBody');
  if (!tbody) return;

  const cats = analytics.categoryBreakdown || [];
  if (cats.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center">No dimension data available.</td></tr>';
    return;
  }

  tbody.innerHTML = cats.map(c => `
    <tr style="cursor: pointer;" onclick="applyQuickCategoryDrilldown('${escapeHtml(c.category)}')">
      <td><strong>${c.category}</strong> <span style="font-size: 0.72rem; color: var(--primary);">🔍</span></td>
      <td class="text-right">${Number(c.volume).toLocaleString()}</td>
      <td class="text-right">${formatCurrency(c.totalValue)}</td>
      <td class="text-right">${formatCurrency(c.avgValue)}</td>
      <td class="text-right"><span class="badge-tag">${c.sharePct}%</span></td>
    </tr>
  `).join('');
}

function renderAnomalies(analytics) {
  const tbody = document.getElementById('anomalyTableBody');
  const banner = document.getElementById('anomalyAlertBanner');
  if (!tbody) return;

  const anoms = analytics.anomalies || [];
  if (anoms.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5" class="text-center" style="padding: 2rem; color: #10b981;">✓ Internal Audit passed: No discrepancies or outlier flags detected.</td></tr>';
    if (banner) banner.style.display = 'none';
    return;
  }

  if (banner) banner.style.display = 'flex';

  tbody.innerHTML = anoms.map(a => {
    let sevClass = 'info';
    if (a.severity === 'Critical') sevClass = 'critical';
    if (a.severity === 'Warning') sevClass = 'warning';

    const refSnippet = a.data ? Object.entries(a.data).slice(0, 3).map(([k, v]) => `${k}: ${v}`).join(' | ') : `Row ${a.rowIndex}`;

    return `
      <tr>
        <td><strong>#${a.rowIndex}</strong></td>
        <td>${a.type}</td>
        <td><span class="badge-severity ${sevClass}">${a.severity}</span></td>
        <td>${a.description}</td>
        <td style="font-family: monospace; font-size: 0.78rem; color: var(--text-muted);">${refSnippet}</td>
      </tr>
    `;
  }).join('');
}

// ==========================================
// DYNAMIC PIVOT MATRIX
// ==========================================
function populatePivotSelectors(analytics) {
  const rowSelect = document.getElementById('pivotRowDim');
  const colSelect = document.getElementById('pivotColDim');
  const valSelect = document.getElementById('pivotValCol');

  if (!rowSelect || !colSelect || !valSelect || !analytics.columns) return;

  rowSelect.innerHTML = '';
  colSelect.innerHTML = '<option value="">(None - Single Column Total)</option>';
  valSelect.innerHTML = '';

  analytics.columns.forEach(col => {
    if (col.type === 'category' || col.type === 'text') {
      rowSelect.innerHTML += `<option value="${col.name}">${col.name}</option>`;
      colSelect.innerHTML += `<option value="${col.name}">${col.name}</option>`;
    }
    if (col.type === 'numeric') {
      valSelect.innerHTML += `<option value="${col.name}">${col.name}</option>`;
    }
  });

  if (state.selectedDimCol) rowSelect.value = state.selectedDimCol;
  if (state.selectedMetricCol) valSelect.value = state.selectedMetricCol;
}

async function runPivot() {
  const rowDim = document.getElementById('pivotRowDim').value;
  const colDim = document.getElementById('pivotColDim').value;
  const valCol = document.getElementById('pivotValCol').value;
  const aggMethod = document.getElementById('pivotAggMethod').value;

  if (!rowDim || !valCol) {
    showToast('Please select both a Row Group and a Value Column.');
    return;
  }

  showToast('Generating Pivot Matrix...');
  try {
    const res = await fetch(`/api/datasets/${state.currentDatasetId}/pivot`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ rowDim, colDim, valCol, aggMethod })
    });
    const data = await res.json();
    if (data.success) {
      renderPivotTable(data.pivot);
    } else {
      throw new Error(data.error);
    }
  } catch (err) {
    showToast('Pivot Error: ' + err.message);
  }
}

function renderPivotTable(pivot) {
  const thead = document.getElementById('pivotThead');
  const tbody = document.getElementById('pivotTbody');
  const title = document.getElementById('pivotResultTitle');

  if (title) {
    title.textContent = `Pivot Matrix: ${pivot.rowDim} × ${pivot.colDim || 'Total'} (${pivot.aggMethod.toUpperCase()} of ${pivot.valCol})`;
  }

  let headerHtml = `<tr><th>${pivot.rowDim} \\ ${pivot.colDim || 'Metrics'}</th>`;
  pivot.colKeys.forEach(col => {
    headerHtml += `<th class="text-right pivot-subhead">${col}</th>`;
  });
  headerHtml += `<th class="text-right">Total (${pivot.aggMethod.toUpperCase()})</th></tr>`;
  thead.innerHTML = headerHtml;

  let bodyHtml = '';
  pivot.rowKeys.forEach(rowKey => {
    bodyHtml += `<tr><td><strong>${rowKey}</strong></td>`;
    pivot.colKeys.forEach(colKey => {
      const val = pivot.matrix[rowKey][colKey] || 0;
      bodyHtml += `<td class="text-right">${formatCurrency(val)}</td>`;
    });
    bodyHtml += `<td class="text-right"><strong>${formatCurrency(pivot.rowTotals[rowKey])}</strong></td></tr>`;
  });

  bodyHtml += `<tr class="pivot-total-row"><td><strong>Grand Total</strong></td>`;
  pivot.colKeys.forEach(colKey => {
    bodyHtml += `<td class="text-right"><strong>${formatCurrency(pivot.colTotals[colKey])}</strong></td>`;
  });
  bodyHtml += `<td class="text-right"><strong>${formatCurrency(pivot.grandTotal)}</strong></td></tr>`;

  tbody.innerHTML = bodyHtml;
}

// ==========================================
// 100% BULLETPROOF ADVANCED FILTER ENGINE
// ==========================================
function populateQuickCategoryFilter(analytics) {
  const dimSelect = document.getElementById('filterDimension');
  const label = document.getElementById('quickCatLabel');
  if (!dimSelect) return;

  const activeDim = state.selectedDimCol || analytics.summaryKPIs?.dimensionColumn;
  if (label && activeDim) {
    label.textContent = `Filter By ${formatColName(activeDim)}:`;
  }

  dimSelect.innerHTML = '<option value="All">All Items</option>';
  if (analytics.categoryBreakdown) {
    analytics.categoryBreakdown.forEach(c => {
      dimSelect.innerHTML += `<option value="${escapeHtml(c.category)}">${c.category} (${c.volume})</option>`;
    });
  }

  dimSelect.value = state.quickCategoryFilter || 'All';
}

function handleDatePresetClick(preset) {
  state.datePreset = preset;
  document.querySelectorAll('.pill-btn').forEach(b => {
    b.classList.toggle('active', b.getAttribute('data-range') === preset);
  });

  const now = new Date();
  let startStr = '';
  let endStr = '';

  if (preset === 'all') {
    startStr = '';
    endStr = '';
  } else if (preset === 'this_month') {
    const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
    startStr = firstDay.toISOString().split('T')[0];
    endStr = now.toISOString().split('T')[0];
  } else if (preset === 'last_30') {
    const past30 = new Date(now.getTime() - 30 * 86400000);
    startStr = past30.toISOString().split('T')[0];
    endStr = now.toISOString().split('T')[0];
  }

  state.dateRange.start = startStr;
  state.dateRange.end = endStr;
  document.getElementById('filterDateStart').value = startStr;
  document.getElementById('filterDateEnd').value = endStr;

  triggerAdvancedFilterRecompute();
}

function addMultiRuleRow() {
  const container = document.getElementById('rulesContainer');
  if (!container || !state.analytics.columns) return;

  const ruleId = 'rule_' + Date.now();
  const row = document.createElement('div');
  row.className = 'rule-row';
  row.id = ruleId;

  const colOptions = state.analytics.columns.map(c => `<option value="${c.name}">${c.name}</option>`).join('');

  row.innerHTML = `
    <select class="rule-col">${colOptions}</select>
    <select class="rule-op">
      <option value="=">=</option>
      <option value="!=">!=</option>
      <option value="contains">contains</option>
      <option value=">">&gt;</option>
      <option value="<">&lt;</option>
    </select>
    <input type="text" class="rule-val" placeholder="Value...">
    <button type="button" class="btn-remove-rule" onclick="document.getElementById('${ruleId}').remove()">&times;</button>
  `;

  container.appendChild(row);
}

function collectActiveRules() {
  const rules = [];
  document.querySelectorAll('.rule-row').forEach(row => {
    const col = row.querySelector('.rule-col')?.value;
    const op = row.querySelector('.rule-op')?.value;
    const val = row.querySelector('.rule-val')?.value;
    if (col && val !== undefined && String(val).trim() !== '') {
      rules.push({ col, op, val: String(val).trim() });
    }
  });
  return rules;
}

// Master Filter Recomputation (Directly reads all UI inputs)
async function triggerAdvancedFilterRecompute() {
  if (!state.currentDatasetId) return;

  // Direct DOM reads
  state.selectedMetricCol = document.getElementById('activeMetricSelect')?.value || state.selectedMetricCol;
  state.selectedDimCol = document.getElementById('activeDimensionSelect')?.value || state.selectedDimCol;

  const quickCat = document.getElementById('filterDimension')?.value || 'All';
  state.quickCategoryFilter = quickCat;

  const startVal = document.getElementById('filterDateStart')?.value || '';
  const endVal = document.getElementById('filterDateEnd')?.value || '';
  state.dateRange.start = startVal;
  state.dateRange.end = endVal;

  const customRules = collectActiveRules();
  const filters = {};
  if (quickCat && quickCat !== 'All') {
    filters[state.selectedDimCol] = quickCat;
  }

  const dateCol = state.analytics.summaryKPIs?.dateColumn;
  const dateRange = (startVal || endVal) ? {
    col: dateCol,
    start: startVal,
    end: endVal
  } : null;

  showToast('Applying live filters & recomputing MIS metrics...');
  try {
    const res = await fetch(`/api/datasets/${state.currentDatasetId}/filter`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        filters,
        multiFilters: customRules,
        dateRange,
        metricCol: state.selectedMetricCol,
        dimCol: state.selectedDimCol
      })
    });
    const data = await res.json();
    if (data.success) {
      state.analytics = data.analytics;
      renderKPIs(data.analytics);
      renderInsights(data.analytics);
      renderCharts(data.analytics);
      renderCategoryTable(data.analytics);
      renderAnomalies(data.analytics);
      populatePivotSelectors(data.analytics);
      renderActiveFilterChips();
      populateQuickCategoryFilter(data.analytics);
      renderDimensionPills(data.analytics);
      renderDoctorAnalytics(data.analytics);

      // UPDATE DATA GRID WITH FILTERED ROWS!
      if (data.filteredRows) {
        state.rawGridRows = data.filteredRows;
        state.rawGridTotal = data.filteredCount;
        renderRawDataGrid(data.filteredRows);
        const counter = document.getElementById('gridTotalCount');
        if (counter) counter.textContent = `Showing ${data.filteredRows.length} of ${data.filteredCount} filtered records`;
      }

      showToast(`✓ Filter applied: ${data.filteredCount} active records.`);
    } else {
      throw new Error(data.error);
    }
  } catch (err) {
    showToast('Filter error: ' + err.message);
  }
}

function renderActiveFilterChips() {
  const container = document.getElementById('chipsList');
  const wrap = document.getElementById('activeFilterChips');
  if (!container || !wrap) return;

  const chips = [];

  if (state.quickCategoryFilter && state.quickCategoryFilter !== 'All') {
    chips.push({
      type: 'category',
      label: `${formatColName(state.selectedDimCol)}: ${state.quickCategoryFilter}`
    });
  }

  if (state.dateRange.start || state.dateRange.end) {
    chips.push({
      type: 'date',
      label: `Date: ${state.dateRange.start || '...'} to ${state.dateRange.end || '...'}`
    });
  }

  collectActiveRules().forEach((r, idx) => {
    chips.push({
      type: 'rule',
      index: idx,
      label: `${r.col} ${r.op} ${r.val}`
    });
  });

  if (chips.length === 0) {
    wrap.style.display = 'none';
    return;
  }

  wrap.style.display = 'flex';
  container.innerHTML = chips.map(c => `
    <div class="chip">
      <span>${c.label}</span>
      <span class="chip-close" onclick="removeActiveFilter('${c.type}', ${c.index})">&times;</span>
    </div>
  `).join('');
}

function removeActiveFilter(type, index) {
  if (type === 'category') {
    state.quickCategoryFilter = 'All';
    const sel = document.getElementById('filterDimension');
    if (sel) sel.value = 'All';
  } else if (type === 'date') {
    state.datePreset = 'all';
    state.dateRange = { start: '', end: '' };
    document.getElementById('filterDateStart').value = '';
    document.getElementById('filterDateEnd').value = '';
    document.querySelectorAll('.pill-btn').forEach(b => b.classList.toggle('active', b.getAttribute('data-range') === 'all'));
  } else if (type === 'rule') {
    const rows = document.querySelectorAll('.rule-row');
    if (rows[index]) rows[index].remove();
  }
  triggerAdvancedFilterRecompute();
}

function resetAllFilters() {
  state.quickCategoryFilter = 'All';
  state.datePreset = 'all';
  state.dateRange = { start: '', end: '' };
  document.getElementById('filterDimension').value = 'All';
  document.getElementById('filterDateStart').value = '';
  document.getElementById('filterDateEnd').value = '';
  document.getElementById('rulesContainer').innerHTML = '';
  document.querySelectorAll('.pill-btn').forEach(b => b.classList.toggle('active', b.getAttribute('data-range') === 'all'));
  if (state.currentDatasetId) {
    loadDataset(state.currentDatasetId);
  }
}

// ==========================================
// COLUMN INTELLIGENCE & INSPECTOR MODAL
// ==========================================
function openColumnInspector() {
  const modal = document.getElementById('colInspectorModal');
  const tbody = document.getElementById('colInspectorTbody');
  if (!modal || !tbody || !state.analytics.columns) return;

  tbody.innerHTML = state.analytics.columns.map(col => {
    let typeBadge = `<span class="badge-tag">${col.type}</span>`;
    if (col.type === 'numeric') typeBadge = `<span class="badge-tag" style="background: rgba(16, 185, 129, 0.2); color: #34d399;">Numeric Metric</span>`;
    if (col.type === 'identifier') typeBadge = `<span class="badge-tag" style="background: rgba(245, 158, 11, 0.2); color: #fbbf24;">ID / Contact</span>`;
    if (col.type === 'date') typeBadge = `<span class="badge-tag" style="background: rgba(14, 165, 233, 0.2); color: #38bdf8;">Date</span>`;

    const sampleText = (col.sampleValues || []).join(', ');
    let statsSummary = '-';
    if (col.type === 'numeric' && col.stats) {
      statsSummary = `Sum: ${formatCompactNumber(col.stats.sum)} | Avg: ${formatCurrency(col.stats.avg)}`;
    } else if (col.stats && col.stats.uniqueCount) {
      statsSummary = `${col.stats.uniqueCount} unique categories`;
    }

    return `
      <tr>
        <td><strong>${col.name}</strong></td>
        <td>${typeBadge}</td>
        <td style="max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 0.8rem; color: var(--text-muted);">${escapeHtml(sampleText)}</td>
        <td style="font-size: 0.8rem;">${statsSummary}</td>
      </tr>
    `;
  }).join('');

  modal.style.display = 'flex';
}

// ==========================================
// DATA GRID EXPLORER
// ==========================================
async function loadRawDataGrid() {
  if (!state.currentDatasetId) return;

  const url = `/api/datasets/${state.currentDatasetId}/data?page=${state.rawGridPage}&limit=${state.rawGridLimit}&search=${encodeURIComponent(state.gridSearch)}`;
  try {
    const res = await fetch(url);
    const data = await res.json();
    if (data.success) {
      state.rawGridRows = data.rows;
      state.rawGridTotal = data.totalRows;

      const counter = document.getElementById('gridTotalCount');
      if (counter) counter.textContent = `Showing ${data.rows.length} of ${data.totalRows} records`;

      const curPage = document.getElementById('curPageNum');
      const totalPage = document.getElementById('totalPageNum');
      if (curPage) curPage.textContent = data.page;
      if (totalPage) totalPage.textContent = data.totalPages || 1;

      renderRawDataGrid(data.rows);
    }
  } catch (err) {
    console.error('[Grid Load Error]', err);
  }
}

function renderRawDataGrid(rows) {
  const thead = document.getElementById('rawDataThead');
  const tbody = document.getElementById('rawDataTbody');

  if (!rows || rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10" class="text-center">No records match your query.</td></tr>';
    return;
  }

  const columns = Object.keys(rows[0]);
  thead.innerHTML = `<tr>${columns.map(c => `<th>${c}</th>`).join('')}</tr>`;
  tbody.innerHTML = rows.map(r => {
    return `<tr>${columns.map(c => {
      const val = r[c];
      const displayVal = (val === null || val === undefined) ? '-' : String(val);
      return `<td>${displayVal}</td>`;
    }).join('')}</tr>`;
  }).join('');
}

// ==========================================
// UPLOAD & IMPORT PROGRESS ENGINE
// ==========================================
let selectedUploadFile = null;
let parsedWorkbookData = null;
let activeImportDatasetId = null;

function formatFileSize(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function clearImportLogs() {
  ['heroImportLogContent', 'modalImportLogContent', 'importLogContent'].forEach(id => {
    const box = document.getElementById(id);
    if (box) box.innerHTML = '';
  });
}

function addImportLog(text, type = 'normal') {
  const now = new Date();
  const timeStr = `${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}.${String(Math.floor(now.getMilliseconds() / 100))}`;
  const logHtml = `[${timeStr}] ${text}`;

  ['heroImportLogContent', 'modalImportLogContent', 'importLogContent'].forEach(id => {
    const box = document.getElementById(id);
    if (!box) return;
    const entry = document.createElement('div');
    entry.className = `log-entry ${type}`;
    entry.textContent = logHtml;
    box.appendChild(entry);
    box.scrollTop = box.scrollHeight;
  });
}

// Client-side Excel Inspection & Live Preview (using SheetJS)
async function inspectAndPreviewFile(file) {
  if (!file) return;

  const validExts = ['.xlsx', '.xls', '.csv'];
  const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
  if (!validExts.includes(ext)) {
    showToast('Invalid file format. Please upload an Excel (.xlsx, .xls) or CSV file.');
    return;
  }

  selectedUploadFile = file;

  const uploadModal = document.getElementById('uploadModal');
  const isModalOpen = uploadModal && uploadModal.style.display === 'flex';

  // Toggle display depending on where user picked file
  if (isModalOpen) {
    const dropzone = document.getElementById('modalDropzone');
    const previewPanel = document.getElementById('filePreviewPanel');
    const inModalProg = document.getElementById('inModalProgress');
    const uploadForm = document.getElementById('modalUploadForm');
    if (inModalProg) inModalProg.style.display = 'none';
    if (uploadForm) uploadForm.style.display = 'block';
    if (dropzone) dropzone.style.display = 'none';
    if (previewPanel) previewPanel.style.display = 'block';
  } else {
    const mainDz = document.getElementById('mainDropzone');
    const heroPrev = document.getElementById('heroFilePreview');
    const heroProg = document.getElementById('heroProgressSection');
    if (heroProg) heroProg.style.display = 'none';
    if (mainDz) mainDz.style.display = 'none';
    if (heroPrev) heroPrev.style.display = 'block';
  }

  // Populate file info on BOTH Hero and Modal previews
  const fileSzStr = formatFileSize(file.size);
  const cleanTitle = file.name.replace(/\.[^/.]+$/, '').replace(/[_-]/g, ' ');

  // Update Hero elements
  const hName = document.getElementById('heroFileName');
  const hSize = document.getElementById('heroFileSize');
  const hRows = document.getElementById('heroRowCount');
  const hCols = document.getElementById('heroColCount');
  const hStat = document.getElementById('heroStatus');
  const hTitleInp = document.getElementById('heroDatasetName');

  if (hName) hName.textContent = file.name;
  if (hSize) hSize.textContent = fileSzStr;
  if (hRows) hRows.textContent = 'Reading sheets...';
  if (hCols) hCols.textContent = 'Analyzing...';
  if (hStat) hStat.textContent = 'Parsing File...';
  if (hTitleInp && !hTitleInp.value) hTitleInp.value = cleanTitle;

  // Update Modal elements
  const mName = document.getElementById('previewFileName');
  const mSize = document.getElementById('previewFileSize');
  const mRows = document.getElementById('previewRowCount');
  const mCols = document.getElementById('previewColCount');
  const mStat = document.getElementById('previewStatus');
  const mTitleInp = document.getElementById('modalDatasetName');

  if (mName) mName.textContent = file.name;
  if (mSize) mSize.textContent = fileSzStr;
  if (mRows) mRows.textContent = 'Reading sheets...';
  if (mCols) mCols.textContent = 'Analyzing...';
  if (mStat) mStat.textContent = 'Parsing File...';
  if (mTitleInp && !mTitleInp.value) mTitleInp.value = cleanTitle;

  const mBtnSubmit = document.getElementById('btnSubmitUpload');
  if (mBtnSubmit) mBtnSubmit.disabled = false;

  try {
    const arrayBuffer = await file.arrayBuffer();
    if (typeof XLSX === 'undefined') {
      [hStat, mStat].forEach(el => el && (el.textContent = 'Ready for Server Import'));
      [hRows, mRows].forEach(el => el && (el.textContent = 'Verified File'));
      [hCols, mCols].forEach(el => el && (el.textContent = `${ext.toUpperCase()} format`));
      return;
    }

    const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    parsedWorkbookData = workbook;

    // Populate Sheet Selectors
    ['heroSheetSelect', 'modalSheetSelect'].forEach(id => {
      const select = document.getElementById(id);
      if (select && workbook.SheetNames.length > 0) {
        select.innerHTML = '';
        workbook.SheetNames.forEach((sName, idx) => {
          const opt = document.createElement('option');
          opt.value = sName;
          opt.textContent = `${sName}${idx === 0 ? ' (Default)' : ''}`;
          select.appendChild(opt);
        });
      }
    });

    renderSheetPreview(workbook, workbook.SheetNames[0]);
  } catch (err) {
    console.warn('[Preview Notice]', err);
    [hStat, mStat].forEach(el => el && (el.textContent = '✓ Ready to Import'));
    [hRows, mRows].forEach(el => el && (el.textContent = 'Verified File'));
    [hCols, mCols].forEach(el => el && (el.textContent = `${ext.toUpperCase()}`));
  }
}

function renderSheetPreview(workbook, targetSheetName) {
  if (!workbook || !workbook.Sheets) return;
  const worksheet = workbook.Sheets[targetSheetName];
  if (!worksheet) return;

  const rawMatrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  if (rawMatrix.length === 0) {
    ['heroRowCount', 'previewRowCount'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = '0 Rows';
    });
    ['heroColCount', 'previewColCount'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = '0 Columns';
    });
    return;
  }

  // Find header row in first 10 rows
  let headerRowIdx = 0;
  let maxCols = 0;
  for (let i = 0; i < Math.min(10, rawMatrix.length); i++) {
    const row = rawMatrix[i];
    const stringCols = row.filter(cell => cell !== '' && cell !== null && cell !== undefined).length;
    if (stringCols > maxCols) {
      maxCols = stringCols;
      headerRowIdx = i;
    }
  }

  const rawHeaders = rawMatrix[headerRowIdx] || [];
  const headers = rawHeaders.map((h, i) => (h ? String(h).trim() : `Col_${i + 1}`));
  const dataRows = rawMatrix.slice(headerRowIdx + 1).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));

  ['heroRowCount', 'previewRowCount'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = `${dataRows.length} Data Rows`;
  });
  ['heroColCount', 'previewColCount'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = `${headers.length} Columns`;
  });
  ['heroStatus', 'previewStatus'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = '✓ Verified & Ready';
  });

  // Render column chips with auto-detected types
  ['heroColChips', 'previewColChips'].forEach(boxId => {
    const chipsContainer = document.getElementById(boxId);
    if (!chipsContainer) return;
    chipsContainer.innerHTML = '';

    headers.forEach((h, colIdx) => {
      let isNumeric = true;
      let isDate = false;
      let isId = /id|uhid|no|code|phone|pin|sl/i.test(h);
      let samplesCount = 0;

      for (let r = 0; r < Math.min(15, dataRows.length); r++) {
        const val = dataRows[r][colIdx];
        if (val !== '' && val !== null && val !== undefined) {
          samplesCount++;
          if (val instanceof Date || (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}/.test(val))) {
            isDate = true;
            isNumeric = false;
            break;
          }
          if (isNaN(Number(String(val).replace(/[$,₹]/g, '').trim()))) {
            isNumeric = false;
          }
        }
      }

      let typeClass = 'category';
      let typeText = 'TXT';
      if (isId) {
        typeClass = 'identifier';
        typeText = 'ID';
      } else if (isDate) {
        typeClass = 'date';
        typeText = 'DATE';
      } else if (isNumeric && samplesCount > 0) {
        typeClass = 'numeric';
        typeText = 'NUM';
      }

      const chip = document.createElement('span');
      chip.className = 'col-chip';
      chip.innerHTML = `<strong>${escapeHtml(h)}</strong> <span class="col-chip-type ${typeClass}">${typeText}</span>`;
      chipsContainer.appendChild(chip);
    });
  });

  // Render Table Preview (first 5 rows) on both Hero and Modal
  const previewSubset = dataRows.slice(0, 5);
  const theadHtml = `<tr>${headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}</tr>`;
  const tbodyHtml = previewSubset.map(row => {
    return `<tr>${headers.map((_, idx) => {
      const val = row[idx];
      const displayVal = val instanceof Date ? val.toLocaleDateString() : (val === null || val === undefined ? '' : String(val));
      return `<td>${escapeHtml(displayVal)}</td>`;
    }).join('')}</tr>`;
  }).join('');

  ['heroThead', 'previewThead'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = theadHtml;
  });
  ['heroTbody', 'previewTbody'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = tbodyHtml;
  });
}

function clearSelectedFilePreview() {
  selectedUploadFile = null;
  parsedWorkbookData = null;

  // Reset Hero view
  const mainDz = document.getElementById('mainDropzone');
  const heroPrev = document.getElementById('heroFilePreview');
  const heroProg = document.getElementById('heroProgressSection');
  const heroInp = document.getElementById('fileInput');

  if (mainDz) mainDz.style.display = 'block';
  if (heroPrev) heroPrev.style.display = 'none';
  if (heroProg) heroProg.style.display = 'none';
  if (heroInp) heroInp.value = '';

  // Reset Modal view
  const modalDz = document.getElementById('modalDropzone');
  const modalPrev = document.getElementById('filePreviewPanel');
  const inModalProg = document.getElementById('inModalProgress');
  const uploadForm = document.getElementById('modalUploadForm');
  const modalInp = document.getElementById('modalFileInput');
  const btnSubmit = document.getElementById('btnSubmitUpload');
  const titleEl = document.getElementById('uploadModalTitle');
  const subEl = document.getElementById('uploadModalSubtitle');

  if (titleEl) titleEl.textContent = 'Upload & Import ERP Excel File';
  if (subEl) subEl.textContent = 'Smart header detection, column profiling & live data preview';
  if (uploadForm) uploadForm.style.display = 'block';
  if (inModalProg) inModalProg.style.display = 'none';
  if (modalDz) modalDz.style.display = 'block';
  if (modalPrev) modalPrev.style.display = 'none';
  if (btnSubmit) btnSubmit.disabled = true;
  if (modalInp) modalInp.value = '';
}

// Unified Progress UI Manager (updates Hero, In-Modal, and Standalone Modal)
function showImportProgress(filename, fileSize, location = 'hero') {
  const fileSzStr = formatFileSize(fileSize);

  if (location === 'modal') {
    const uploadForm = document.getElementById('modalUploadForm');
    const inModalProg = document.getElementById('inModalProgress');
    const titleEl = document.getElementById('uploadModalTitle');
    const subEl = document.getElementById('uploadModalSubtitle');

    if (uploadForm) uploadForm.style.display = 'none';
    if (inModalProg) inModalProg.style.display = 'block';
    if (titleEl) titleEl.textContent = 'Importing ERP Dataset into MySQL MIS_DB';
    if (subEl) subEl.textContent = 'Multi-stage processing: Smart Header Parsing → Batch Insert → Executive Analytics';

    setImportStep(1, 'active');
  } else if (location === 'hero') {
    const mainDz = document.getElementById('mainDropzone');
    const heroPrev = document.getElementById('heroFilePreview');
    const heroProg = document.getElementById('heroProgressSection');

    if (mainDz) mainDz.style.display = 'none';
    if (heroPrev) heroPrev.style.display = 'none';
    if (heroProg) heroProg.style.display = 'block';

    setImportStep(1, 'active');
  } else {
    const modal = document.getElementById('importProgressModal');
    if (modal) modal.style.display = 'flex';
    setImportStep(1, 'active');
  }

  // Update labels across all containers
  ['heroImportFileName', 'modalImportFileName', 'importFileName'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = filename || 'Dataset.xlsx';
  });

  ['heroImportFileSize', 'modalImportFileSize', 'importFileSize'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = fileSzStr;
  });

  ['heroImportStatusText', 'modalImportStatusText', 'importStatusText'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.textContent = 'Uploading dataset to server...';
      el.style.color = 'var(--primary)';
    }
  });

  ['heroImportPercent', 'modalImportPercent', 'importPercent'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = '0%';
  });

  ['heroImportBarFill', 'modalImportBarFill', 'importBarFill'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.style.width = '0%';
      el.style.background = 'linear-gradient(90deg, #6366f1 0%, #3b82f6 45%, #10b981 100%)';
    }
  });

  ['heroImportBytes', 'modalImportBytes', 'importBytes'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = `0 B of ${fileSzStr}`;
  });

  ['heroImportEta', 'modalImportEta', 'importEta'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = 'Connecting...';
  });

  ['heroImportCompleteCard', 'modalImportCompleteCard', 'importCompleteCard'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });

  ['modalImportActions', 'importActions'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
}

function updateImportProgress(pct, statusText, bytesText = '', etaText = '') {
  const clamped = Math.min(100, Math.max(0, Math.round(pct)));

  ['heroImportPercent', 'modalImportPercent', 'importPercent'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = `${clamped}%`;
  });

  ['heroImportBarFill', 'modalImportBarFill', 'importBarFill'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.width = `${clamped}%`;
  });

  if (statusText) {
    ['heroImportStatusText', 'modalImportStatusText', 'importStatusText'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = statusText;
    });
  }

  if (bytesText) {
    ['heroImportBytes', 'modalImportBytes', 'importBytes'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = bytesText;
    });
  }

  if (etaText) {
    ['heroImportEta', 'modalImportEta', 'importEta'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.textContent = etaText;
    });
  }
}

function setImportStep(stepNum, status = 'active') {
  const prefixes = ['h', 'm', ''];

  prefixes.forEach(prefix => {
    const stepIds = [`${prefix}step-upload`, `${prefix}step-parse`, `${prefix}step-db`, `${prefix}step-analytics`];
    const lineIds = [`${prefix}line-1`, `${prefix}line-2`, `${prefix}line-3`];

    stepIds.forEach((id, idx) => {
      const el = document.getElementById(id);
      if (!el) return;
      const num = idx + 1;
      el.classList.remove('active', 'completed');
      if (num < stepNum) {
        el.classList.add('completed');
      } else if (num === stepNum) {
        if (status === 'completed') {
          el.classList.add('completed');
        } else {
          el.classList.add('active');
        }
      }
    });

    lineIds.forEach((id, idx) => {
      const line = document.getElementById(id);
      if (!line) return;
      if (idx + 1 < stepNum) {
        line.classList.add('completed');
      } else {
        line.classList.remove('completed');
      }
    });
  });
}

function handleImportFailure(errorMsg) {
  ['heroImportIconWrap', 'modalImportIconWrap', 'importIconWrap'].forEach(id => {
    const wrap = document.getElementById(id);
    if (!wrap) return;
    wrap.classList.add('error');
    const fileSvg = wrap.querySelector('.import-file-svg');
    const errSvg = wrap.querySelector('.import-error-svg');
    if (fileSvg) fileSvg.style.display = 'none';
    if (errSvg) errSvg.style.display = 'block';
  });

  ['heroImportBarFill', 'modalImportBarFill', 'importBarFill'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.background = '#f43f5e';
  });

  ['heroImportStatusText', 'modalImportStatusText', 'importStatusText'].forEach(id => {
    const el = document.getElementById(id);
    if (el) {
      el.textContent = 'Import Failed: ' + errorMsg;
      el.style.color = '#f43f5e';
    }
  });

  ['heroImportEta', 'modalImportEta', 'importEta'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.textContent = 'Failed';
  });

  ['modalImportActions', 'importActions'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = 'flex';
  });

  addImportLog(`❌ Error: ${errorMsg}`, 'error');
  showToast('Error: ' + errorMsg);
}

function setupDropzones() {
  const mainDz = document.getElementById('mainDropzone');
  const fileInp = document.getElementById('fileInput');
  const modalDz = document.getElementById('modalDropzone');
  const modalFileInp = document.getElementById('modalFileInput');
  const heroSheetSelect = document.getElementById('heroSheetSelect');
  const modalSheetSelect = document.getElementById('modalSheetSelect');

  // Hero Dropzone Drag/Drop & Change
  if (mainDz && fileInp) {
    mainDz.addEventListener('dragover', (e) => {
      e.preventDefault();
      mainDz.classList.add('dragover');
    });
    mainDz.addEventListener('dragleave', () => mainDz.classList.remove('dragover'));
    mainDz.addEventListener('drop', (e) => {
      e.preventDefault();
      mainDz.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) {
        handleFileUpload(e.dataTransfer.files[0]);
      }
    });

    fileInp.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleFileUpload(e.target.files[0]);
      }
    });
  }

  // Modal Dropzone Drag/Drop & Change
  if (modalDz && modalFileInp) {
    modalDz.addEventListener('dragover', (e) => {
      e.preventDefault();
      modalDz.classList.add('dragover');
    });
    modalDz.addEventListener('dragleave', () => modalDz.classList.remove('dragover'));
    modalDz.addEventListener('drop', (e) => {
      e.preventDefault();
      modalDz.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) {
        inspectAndPreviewFile(e.dataTransfer.files[0]);
      }
    });

    modalFileInp.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        inspectAndPreviewFile(e.target.files[0]);
      }
    });
  }

  // Hero Submit Upload Button
  document.getElementById('btnHeroSubmitUpload')?.addEventListener('click', () => {
    const file = selectedUploadFile || document.getElementById('fileInput')?.files[0];
    const nameInp = document.getElementById('heroDatasetName');
    const sheetSelect = document.getElementById('heroSheetSelect');
    if (file) {
      handleFileUpload(file, nameInp ? nameInp.value : '', sheetSelect ? sheetSelect.value : '');
    } else {
      showToast('Please select an Excel or CSV file first.');
    }
  });

  // Sheet Select Change in Previews
  [heroSheetSelect, modalSheetSelect].forEach(select => {
    if (select) {
      select.addEventListener('change', (e) => {
        if (parsedWorkbookData && e.target.value) {
          renderSheetPreview(parsedWorkbookData, e.target.value);
        }
      });
    }
  });
}

function handleFileUpload(file, customName = '', targetSheet = '') {
  if (!file) return;

  const validExts = ['.xlsx', '.xls', '.csv'];
  const fileExt = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
  if (!validExts.includes(fileExt)) {
    showToast('Invalid file format. Please upload an Excel (.xlsx, .xls) or CSV file.');
    return;
  }

  const isModalOpen = document.getElementById('uploadModal')?.style.display === 'flex';
  const location = isModalOpen ? 'modal' : 'hero';

  showImportProgress(file.name, file.size, location);
  clearImportLogs();
  addImportLog(`Starting import pipeline for "${file.name}" (${formatFileSize(file.size)})...`, 'active');

  const formData = new FormData();
  formData.append('file', file);
  if (customName) formData.append('datasetName', customName);
  if (targetSheet) formData.append('sheetName', targetSheet);

  const xhr = new XMLHttpRequest();
  let serverProcessingInterval = null;
  let simulatedPct = 20;

  updateImportProgress(15, 'Uploading dataset to server...', `${formatFileSize(file.size * 0.15)} of ${formatFileSize(file.size)}`, 'Transferring bytes...');
  setImportStep(1, 'active');

  xhr.upload.addEventListener('progress', (e) => {
    if (e.lengthComputable) {
      const uploadPct = Math.min(30, Math.round((e.loaded / e.total) * 30));
      updateImportProgress(uploadPct, `Uploading file (${uploadPct}%)...`, `${formatFileSize(e.loaded)} of ${formatFileSize(e.total)}`, 'Transferring bytes...');
    }
  });

  xhr.upload.addEventListener('load', () => {
    setImportStep(1, 'completed');
    setImportStep(2, 'active');
    addImportLog('Byte transfer complete. Server reading Excel workbook & extracting headers...', 'active');
    updateImportProgress(35, 'Smart parsing headers, sheets & schema...', 'File uploaded to server', 'Analyzing structure...');

    // Smooth stage pacing so user sees every phase clearly
    serverProcessingInterval = setInterval(() => {
      if (simulatedPct < 90) {
        simulatedPct += 8;
        if (simulatedPct >= 45 && simulatedPct < 65) {
          setImportStep(2, 'completed');
          setImportStep(3, 'active');
          addImportLog('Auto-detected table schema. Writing records into MySQL MIS_DB in batches...', 'active');
          updateImportProgress(simulatedPct, 'Bulk inserting records into MySQL MIS_DB...', 'Writing to dataset_rows...', 'High-speed batch insert');
        } else if (simulatedPct >= 65) {
          setImportStep(3, 'completed');
          setImportStep(4, 'active');
          addImportLog('Generating Pareto 80/20 distributions, KPIs & anomaly watch...', 'active');
          updateImportProgress(simulatedPct, 'Computing ERP analytics, Pareto & KPIs...', 'Calculating aggregates...', 'Finalizing charts...');
        } else {
          updateImportProgress(simulatedPct, 'Smart parsing sheets & auto-detecting types...', 'Reading rows...', 'In progress...');
        }
      }
    }, 280);
  });

  xhr.addEventListener('load', async () => {
    clearInterval(serverProcessingInterval);

    if (xhr.status >= 200 && xhr.status < 300) {
      try {
        const data = JSON.parse(xhr.responseText);
        if (data.success) {
          // Smooth finish to 100%
          setTimeout(() => {
            setImportStep(4, 'completed');
            updateImportProgress(100, `✓ Successfully imported in ${data.processingTimeSec}s!`, `${Number(data.analytics?.summaryKPIs?.totalRecords || 0).toLocaleString()} rows stored`, 'Ready!');

            // Success Icons
            ['heroImportIconWrap', 'modalImportIconWrap', 'importIconWrap'].forEach(id => {
              const wrap = document.getElementById(id);
              if (!wrap) return;
              wrap.classList.add('success');
              const fileSvg = wrap.querySelector('.import-file-svg');
              const succSvg = wrap.querySelector('.import-success-svg');
              if (fileSvg) fileSvg.style.display = 'none';
              if (succSvg) succSvg.style.display = 'block';
            });

            addImportLog(`✓ [Success] "${data.datasetName}" (${data.analytics?.summaryKPIs?.totalRecords} rows, ${data.analytics?.summaryKPIs?.activeColumns} columns) ready in MIS_DB.`, 'success');

            // Show Completion Cards with stats
            const totalRowsFormatted = Number(data.analytics?.summaryKPIs?.totalRecords || 0).toLocaleString();
            const totalCols = data.analytics?.summaryKPIs?.activeColumns || 0;
            const detectedDomain = data.analytics?.detectedDomain || 'MIS Analytics';
            const procTime = `${data.processingTimeSec}s`;

            // Update all complete cards
            ['heroCompTotalRows', 'modalCompTotalRows', 'compTotalRows'].forEach(id => {
              const el = document.getElementById(id);
              if (el) el.textContent = totalRowsFormatted;
            });
            ['heroCompTotalCols', 'modalCompTotalCols', 'compTotalCols'].forEach(id => {
              const el = document.getElementById(id);
              if (el) el.textContent = totalCols;
            });
            ['heroCompDomain', 'modalCompDomain', 'compDomain'].forEach(id => {
              const el = document.getElementById(id);
              if (el) el.textContent = detectedDomain;
            });
            ['heroCompTime', 'modalCompTime', 'compTime'].forEach(id => {
              const el = document.getElementById(id);
              if (el) el.textContent = procTime;
            });

            ['heroImportCompleteCard', 'modalImportCompleteCard', 'importCompleteCard'].forEach(id => {
              const el = document.getElementById(id);
              if (el) el.style.display = 'block';
            });

            activeImportDatasetId = data.datasetId;

            // Wire Go to Dashboard buttons
            const handleGo = async () => {
              closeModals();
              clearSelectedFilePreview();
              localStorage.setItem('mis_active_dataset_id', data.datasetId);
              if (window.location.pathname.endsWith('home.html')) {
                window.location.href = 'overview.html';
              } else {
                await fetchDatasetsList();
                await checkSystemStatus();
                await loadDataset(data.datasetId);
                showToast(`✓ Loaded "${data.datasetName}"!`);
              }
            };

            ['heroBtnGoToDashboard', 'modalBtnGoToDashboard', 'btnGoToDashboard'].forEach(id => {
              const btn = document.getElementById(id);
              if (btn) btn.onclick = handleGo;
            });

            // Auto-redirect after 3.2s so user has ample time to see 100% and stats
            setTimeout(async () => {
              if (activeImportDatasetId === data.datasetId) {
                handleGo();
              }
            }, 3200);
          }, 350);
        } else {
          handleImportFailure(data.error || 'Server processing error');
        }
      } catch (e) {
        handleImportFailure('Failed to parse server response: ' + e.message);
      }
    } else {
      let errMsg = 'Upload failed';
      try {
        const errData = JSON.parse(xhr.responseText);
        errMsg = errData.error || errMsg;
      } catch (_) {
        errMsg = `Server error ${xhr.status}: ${xhr.statusText}`;
      }
      handleImportFailure(errMsg);
    }
  });

  xhr.addEventListener('error', () => {
    clearInterval(serverProcessingInterval);
    handleImportFailure('Network connection error while uploading file.');
  });

  xhr.addEventListener('abort', () => {
    clearInterval(serverProcessingInterval);
    handleImportFailure('Upload was cancelled.');
  });

  xhr.open('POST', '/api/upload', true);
  xhr.send(formData);
}

// ==========================================
// ASK BI COPILOT (CONVERSATIONAL BI)
// ==========================================
async function handleAskBIQuery(query) {
  if (!state.currentDatasetId) {
    showToast('Please select or upload a dataset first.');
    return;
  }

  const inputEl = document.getElementById('copilotQueryInput');
  const cleanQ = (query || inputEl?.value || '').trim();
  if (!cleanQ) {
    showToast('Please type a question for Ask BI Copilot.');
    return;
  }

  if (inputEl) inputEl.value = cleanQ;

  const resultWrap = document.getElementById('copilotResultWrap');
  const narrative = document.getElementById('copilotNarrative');
  const metaTags = document.getElementById('copilotMetaTags');
  const title = document.getElementById('askBiChartTitle');
  const caption = document.getElementById('askBiChartCaption');

  showToast('Analyzing dataset with Ask BI Copilot...');

  try {
    const res = await fetch(`/api/datasets/${state.currentDatasetId}/ask`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ query: cleanQ })
    });
    const data = await res.json();
    if (!data.success) {
      throw new Error(data.error || 'Failed to process BI query');
    }

    state.currentAskBiResult = data;
    state.askBiChartType = data.chartType || 'bar';

    if (resultWrap) resultWrap.style.display = 'flex';
    if (narrative) {
      narrative.innerHTML = (data.answer || '').replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    }
    if (metaTags) {
      metaTags.innerHTML = `
        <span class="copilot-meta-tag">Metric: <strong>${escapeHtml(data.yAxis)}</strong></span>
        <span class="copilot-meta-tag">Dimension: <strong>${escapeHtml(data.xAxis)}</strong></span>
        <span class="copilot-meta-tag">Aggregation: <strong>${data.agg.toUpperCase()}</strong></span>
        <span class="copilot-meta-tag">Top: <strong>${data.topN}</strong></span>
      `;
    }
    if (title) title.textContent = `${data.agg.toUpperCase()} of ${data.yAxis} by ${data.xAxis}`;
    if (caption) caption.textContent = `Auto-parsed question: "${cleanQ}"`;

    // Render chart and table
    renderAskBiChart(data.visual, state.askBiChartType);
    renderAskBiTable(data.visual);

    // Update active button in type selector
    document.querySelectorAll('#askBiTypeSelector .btn-tool-type').forEach(b => {
      b.classList.toggle('active', b.getAttribute('data-type') === state.askBiChartType);
    });

  } catch (err) {
    showToast('Ask BI Error: ' + err.message);
  }
}

function renderAskBiChart(visual, chartType = 'bar') {
  destroyChart('askBiChart');
  const canvas = document.getElementById('askBiChart');
  if (!canvas || !visual || !visual.labels) return;

  const isDark = state.theme === 'dark';
  const textColor = isDark ? '#94a3b8' : '#475569';
  const gridColor = isDark ? 'rgba(255, 255, 255, 0.07)' : 'rgba(0, 0, 0, 0.07)';

  const colors = [
    '#0095c7', '#00c896', '#38bdf8', '#f59e0b', '#005a8e',
    '#14b8a6', '#f97316', '#06b6d4', '#84cc16', '#a855f7'
  ];

  let effectiveType = chartType;
  let indexAxis = 'x';
  if (chartType === 'horizontalBar') {
    effectiveType = 'bar';
    indexAxis = 'y';
  }

  const ds = visual.datasets[0] || { data: [] };
  let bgColors;
  if (chartType === 'doughnut') {
    bgColors = visual.labels.map((_, idx) => colors[idx % colors.length]);
  } else {
    bgColors = 'rgba(0, 149, 199, 0.75)';
  }

  state.charts['askBiChart'] = new Chart(canvas.getContext('2d'), {
    type: effectiveType,
    data: {
      labels: visual.labels,
      datasets: [{
        label: ds.label || 'Value',
        data: ds.data,
        backgroundColor: bgColors,
        borderColor: chartType === 'line' ? '#0095c7' : undefined,
        borderWidth: chartType === 'line' ? 2.5 : 1,
        fill: chartType === 'line',
        tension: 0.35,
        borderRadius: chartType === 'doughnut' ? 4 : 6
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: indexAxis,
      plugins: {
        legend: {
          display: chartType === 'doughnut',
          position: 'right',
          labels: { color: textColor, boxWidth: 12 }
        },
        tooltip: {
          callbacks: {
            label: (ctx) => ` Value: ${formatCurrency(ctx.parsed[indexAxis === 'y' ? 'x' : 'y'] !== undefined ? ctx.parsed[indexAxis === 'y' ? 'x' : 'y'] : ctx.raw)}`
          }
        }
      },
      scales: chartType === 'doughnut' ? {} : {
        x: {
          grid: { color: gridColor },
          ticks: { color: textColor }
        },
        y: {
          grid: { color: gridColor },
          ticks: {
            color: textColor,
            callback: (v) => formatCompactNumber(v)
          }
        }
      }
    }
  });
}

function renderAskBiTable(visual) {
  const tbody = document.getElementById('askBiTableBody');
  if (!tbody || !visual || !visual.labels) return;

  const total = visual.datasets[0]?.data.reduce((a, b) => a + (Number(b) || 0), 0) || 1;
  tbody.innerHTML = visual.labels.map((lbl, idx) => {
    const val = visual.datasets[0]?.data[idx] || 0;
    const share = ((val / total) * 100).toFixed(1);
    return `
      <tr>
        <td><strong>${escapeHtml(lbl)}</strong></td>
        <td class="text-right">${formatCurrency(val)}</td>
        <td class="text-right">${share}%</td>
      </tr>
    `;
  }).join('');
}

function exportAskBiCsv() {
  if (!state.currentAskBiResult || !state.currentAskBiResult.visual) {
    showToast('No Ask BI visual data to export.');
    return;
  }
  const visual = state.currentAskBiResult.visual;
  const rows = [['Category', 'Value']];
  visual.labels.forEach((lbl, idx) => {
    rows.push([lbl, visual.datasets[0]?.data[idx] || 0]);
  });
  const csvContent = 'data:text/csv;charset=utf-8,' + rows.map(e => e.join(',')).join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', `AskBI_${(state.currentAskBiResult.xAxis || 'visual').replace(/\s+/g, '_')}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('✓ Ask BI data exported as CSV!');
}

// ==========================================
// VISUAL BUILDER STUDIO (POWER BI DESIGNER)
// ==========================================
function populateVisualBuilderControls() {
  const xAxisSelect = document.getElementById('builderXAxis');
  const yAxisSelect = document.getElementById('builderYAxis');
  const groupBySelect = document.getElementById('builderGroupBy');

  if (!xAxisSelect || !yAxisSelect || !groupBySelect || !state.analytics || !state.analytics.columns) return;

  xAxisSelect.innerHTML = '';
  yAxisSelect.innerHTML = '';
  groupBySelect.innerHTML = '<option value="">None (Single Series)</option>';

  const cols = state.analytics.columns;
  const numCols = cols.filter(c => c.type === 'numeric');
  const dimCols = cols.filter(c => c.type === 'category' || c.type === 'text' || c.type === 'date');

  // Populate X-Axis
  cols.forEach(col => {
    const opt = document.createElement('option');
    opt.value = col.name;
    opt.textContent = `${col.name} [${col.type}]`;
    if (col.name === state.selectedDimCol || (!state.selectedDimCol && (col.type === 'category' || col.type === 'text'))) {
      if (!xAxisSelect.value) opt.selected = true;
    }
    xAxisSelect.appendChild(opt);
  });

  // Populate Y-Axis
  numCols.forEach(col => {
    const opt = document.createElement('option');
    opt.value = col.name;
    opt.textContent = `${col.name} (${col.stats?.sum ? formatCompactNumber(col.stats.sum) : 'Numeric'})`;
    if (col.name === state.selectedMetricCol || (!state.selectedMetricCol && numCols[0]?.name === col.name)) {
      opt.selected = true;
    }
    yAxisSelect.appendChild(opt);
  });

  // Populate Group By (Legend)
  dimCols.forEach(col => {
    const opt = document.createElement('option');
    opt.value = col.name;
    opt.textContent = col.name;
    groupBySelect.appendChild(opt);
  });
}

async function renderCustomVisualStudio() {
  if (!state.currentDatasetId) {
    showToast('Please select or upload a dataset first.');
    return;
  }

  const xAxis = document.getElementById('builderXAxis')?.value;
  const yAxis = document.getElementById('builderYAxis')?.value;
  const agg = document.getElementById('builderAgg')?.value || 'sum';
  const groupBy = document.getElementById('builderGroupBy')?.value || '';
  const topN = parseInt(document.getElementById('builderTopN')?.value || '0', 10);
  const sort = document.getElementById('builderSort')?.value || 'desc';

  if (!xAxis || !yAxis) {
    showToast('Please select both Axis Dimension and Metric Value.');
    return;
  }

  showToast('Rendering custom visual...');
  try {
    const res = await fetch(`/api/datasets/${state.currentDatasetId}/visual`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ xAxis, yAxis, agg, groupBy, topN, sort })
    });
    const data = await res.json();
    if (!data.success) {
      throw new Error(data.error || 'Failed to build visual');
    }

    state.currentBuilderResult = { ...data.visual, xAxis, yAxis, agg, groupBy };

    // Update Header
    const title = document.getElementById('builderVisualTitle');
    const caption = document.getElementById('builderVisualCaption');
    if (title) title.textContent = `${agg.toUpperCase()} of ${yAxis} by ${xAxis} ${groupBy ? '(Legend: ' + groupBy + ')' : ''}`;
    if (caption) caption.textContent = `Visualized as ${state.selectedVisualType.toUpperCase()} • ${data.visual.labels.length} categories`;

    // Compute Stats
    let totalVal = 0;
    let peakItem = '-';
    let peakVal = -Infinity;
    let dataPoints = 0;

    if (data.visual.items) {
      dataPoints = data.visual.items.length;
      data.visual.items.forEach(it => {
        totalVal += it.value;
        if (it.value > peakVal) {
          peakVal = it.value;
          peakItem = `${it.key} (${formatCompactNumber(it.value)})`;
        }
      });
    } else if (data.visual.datasets) {
      data.visual.datasets.forEach(ds => {
        ds.data.forEach(v => totalVal += Number(v) || 0);
      });
      dataPoints = data.visual.labels.length;
    }

    const avgVal = dataPoints > 0 ? totalVal / dataPoints : 0;

    const bTotal = document.getElementById('bkpiTotal');
    const bAvg = document.getElementById('bkpiAvg');
    const bPeak = document.getElementById('bkpiPeak');
    const bCount = document.getElementById('bkpiCount');

    if (bTotal) bTotal.textContent = formatCurrency(totalVal);
    if (bAvg) bAvg.textContent = formatCurrency(avgVal);
    if (bPeak) bPeak.textContent = peakItem !== '-' ? peakItem : formatCurrency(peakVal > -Infinity ? peakVal : 0);
    if (bCount) bCount.textContent = dataPoints.toLocaleString();

    // Render Canvas Chart
    renderBuilderChart(data.visual, state.selectedVisualType);

    // Render Data Breakdown Table
    renderBuilderTable(data.visual, totalVal);

  } catch (err) {
    showToast('Visual Build Error: ' + err.message);
  }
}

function renderBuilderChart(visual, vType = 'bar') {
  destroyChart('builderCanvasChart');
  const canvas = document.getElementById('builderCanvasChart');
  if (!canvas || !visual || !visual.labels) return;

  const isDark = state.theme === 'dark';
  const textColor = isDark ? '#94a3b8' : '#475569';
  const gridColor = isDark ? 'rgba(255, 255, 255, 0.07)' : 'rgba(0, 0, 0, 0.07)';

  const colors = [
    '#0095c7', '#00c896', '#38bdf8', '#f59e0b', '#005a8e',
    '#14b8a6', '#f97316', '#06b6d4', '#84cc16', '#a855f7'
  ];

  let chartType = vType;
  let indexAxis = 'x';
  if (vType === 'horizontalBar') {
    chartType = 'bar';
    indexAxis = 'y';
  } else if (vType === 'area') {
    chartType = 'line';
  }

  // Multi-series or Single-series
  const datasets = visual.datasets.map((ds, idx) => {
    const col = colors[idx % colors.length];
    return {
      label: ds.label,
      data: ds.data,
      backgroundColor: vType === 'doughnut' ? visual.labels.map((_, i) => colors[i % colors.length]) : (vType === 'area' ? col + '33' : (vType === 'line' || vType === 'radar' ? col : col + 'bb')),
      borderColor: col,
      borderWidth: vType === 'line' || vType === 'area' || vType === 'radar' ? 2.5 : 1,
      fill: vType === 'area',
      tension: 0.35,
      borderRadius: vType === 'doughnut' ? 4 : 6
    };
  });

  state.charts['builderCanvasChart'] = new Chart(canvas.getContext('2d'), {
    type: chartType,
    data: {
      labels: visual.labels,
      datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: indexAxis,
      plugins: {
        legend: {
          display: visual.datasets.length > 1 || vType === 'doughnut' || vType === 'radar',
          position: 'top',
          labels: { color: textColor, boxWidth: 12 }
        },
        tooltip: {
          callbacks: {
            label: (ctx) => ` ${ctx.dataset.label}: ${formatCurrency(ctx.parsed[indexAxis === 'y' ? 'x' : 'y'] !== undefined ? ctx.parsed[indexAxis === 'y' ? 'x' : 'y'] : ctx.raw)}`
          }
        }
      },
      scales: (vType === 'doughnut' || vType === 'radar') ? {} : {
        x: {
          grid: { color: gridColor },
          ticks: { color: textColor }
        },
        y: {
          grid: { color: gridColor },
          ticks: {
            color: textColor,
            callback: (v) => formatCompactNumber(v)
          }
        }
      }
    }
  });
}

function renderBuilderTable(visual, totalVal) {
  const tbody = document.getElementById('builderTbody');
  if (!tbody || !visual) return;

  if (visual.items && visual.items.length > 0) {
    tbody.innerHTML = visual.items.map(it => {
      const share = totalVal > 0 ? ((it.value / totalVal) * 100).toFixed(1) : 0;
      return `
        <tr>
          <td><strong>${escapeHtml(it.key)}</strong></td>
          <td class="text-right">${formatCurrency(it.value)}</td>
          <td class="text-right">${Number(it.count).toLocaleString()}</td>
          <td class="text-right">${share}%</td>
        </tr>
      `;
    }).join('');
  } else if (visual.labels) {
    tbody.innerHTML = visual.labels.map((lbl, idx) => {
      const sum = visual.datasets.reduce((acc, ds) => acc + (Number(ds.data[idx]) || 0), 0);
      const share = totalVal > 0 ? ((sum / totalVal) * 100).toFixed(1) : 0;
      return `
        <tr>
          <td><strong>${escapeHtml(lbl)}</strong></td>
          <td class="text-right">${formatCurrency(sum)}</td>
          <td class="text-right">-</td>
          <td class="text-right">${share}%</td>
        </tr>
      `;
    }).join('');
  }
}

function exportBuilderCsv() {
  if (!state.currentBuilderResult) {
    showToast('No Custom Visual data to export.');
    return;
  }
  const visual = state.currentBuilderResult;
  const rows = [];
  if (visual.items && visual.items.length > 0) {
    rows.push(['Category', 'Aggregated_Value', 'Row_Count']);
    visual.items.forEach(it => {
      rows.push([`"${it.key}"`, it.value, it.count]);
    });
  } else if (visual.labels && visual.datasets) {
    const headers = ['Category', ...visual.datasets.map(d => d.label)];
    rows.push(headers);
    visual.labels.forEach((lbl, idx) => {
      const row = [`"${lbl}"`, ...visual.datasets.map(d => d.data[idx] || 0)];
      rows.push(row);
    });
  }

  const csvContent = 'data:text/csv;charset=utf-8,' + rows.map(e => e.join(',')).join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', `Visual_${(visual.xAxis || 'custom').replace(/\s+/g, '_')}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  showToast('✓ Custom Visual data exported as CSV!');
}

// ==========================================
// REPORT DOWNLOADS: MULTI-SHEET EXCEL & PDF
// ==========================================

// 1. Comprehensive Multi-Sheet Excel Report (.xlsx)
function downloadFullExcelReport() {
  if (!state.currentDataset || !state.analytics) {
    showToast('No active report to download.');
    return;
  }

  showToast('Generating Multi-Sheet Executive Excel Workbook...');
  const wb = XLSX.utils.book_new();

  // Sheet 1: Executive Summary
  const kpis = state.analytics.summaryKPIs || {};
  const summaryData = [
    ['MIS ANALYTICS PRO - EXECUTIVE REPORT'],
    ['Report Title', state.currentDataset.name],
    ['Original File', state.currentDataset.filename],
    ['Domain Classification', state.currentDataset.detected_type],
    ['Generated At', new Date().toLocaleString()],
    ['Active Financial Metric', kpis.metricColumn || 'Default'],
    ['Primary Dimension', state.selectedDimCol || kpis.dimensionColumn || 'Default'],
    [''],
    ['EXECUTIVE SUMMARY METRICS', 'VALUE'],
    ['Total Financial Volume', kpis.totalValue],
    ['Total Transactions', kpis.totalRecords],
    ['Average Ticket Size', kpis.averageTicket],
    ['Max Single Transaction', kpis.highestTransaction],
    ['Lowest Transaction', kpis.lowestTransaction],
    ['Zero / Negative Transactions Flagged', kpis.zeroOrNegativeCount],
    [''],
    ['AI EXECUTIVE FINDINGS']
  ];

  (state.analytics.insights || []).forEach(ins => {
    summaryData.push([ins.replace(/\*\*/g, '')]);
  });

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Executive_Summary');

  // Sheet 2: Dimension Breakdown
  if (state.analytics.categoryBreakdown && state.analytics.categoryBreakdown.length > 0) {
    const wsCategory = XLSX.utils.json_to_sheet(state.analytics.categoryBreakdown.map(c => ({
      Category: c.category,
      'Transaction Count': c.volume,
      'Total Value': c.totalValue,
      'Average Value': c.avgValue,
      'Market Share (%)': c.sharePct
    })));
    XLSX.utils.book_append_sheet(wb, wsCategory, 'Category_Breakdown');
  }

  // Sheet 3: Filtered Raw Data Records
  if (state.rawGridRows && state.rawGridRows.length > 0) {
    const wsRaw = XLSX.utils.json_to_sheet(state.rawGridRows);
    XLSX.utils.book_append_sheet(wb, wsRaw, 'Filtered_Records');
  }

  // Sheet 4: Internal Audit & Exceptions
  if (state.analytics.anomalies && state.analytics.anomalies.length > 0) {
    const wsAudit = XLSX.utils.json_to_sheet(state.analytics.anomalies.map(a => ({
      'Row #': a.rowIndex,
      'Type': a.type,
      'Severity': a.severity,
      'Description': a.description
    })));
    XLSX.utils.book_append_sheet(wb, wsAudit, 'Audit_Exceptions');
  }

  // Save Workbook
  const safeName = (state.currentDataset.name || 'MIS_Report').replace(/[^a-zA-Z0-9_-]/g, '_');
  XLSX.writeFile(wb, `${safeName}_Executive_Report.xlsx`);
  showToast(`✓ Downloaded ${safeName}_Executive_Report.xlsx!`);
}

// 2. Formatted Executive PDF Report (.pdf)
async function downloadExecutivePdfReport() {
  if (!state.currentDataset || !state.analytics) {
    showToast('No active report to download.');
    return;
  }

  showToast('Compiling Executive PDF Report...');

  // Create printable container
  const reportWrap = document.createElement('div');
  reportWrap.id = 'tempPdfReportWrap';
  reportWrap.style.padding = '24px';
  reportWrap.style.fontFamily = 'Inter, sans-serif';
  reportWrap.style.background = '#ffffff';
  reportWrap.style.color = '#0f172a';
  reportWrap.style.maxWidth = '800px';
  reportWrap.style.margin = '0 auto';

  const kpis = state.analytics.summaryKPIs || {};
  const cats = state.analytics.categoryBreakdown || [];

  reportWrap.innerHTML = `
    <div style="border-bottom: 2px solid #0095c7; padding-bottom: 12px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: flex-end;">
      <div>
        <h1 style="font-size: 20px; font-weight: 800; color: #071422; margin: 0;">🏥 PEERLESS HOSPITAL & B.K. ROY RESEARCH CENTRE</h1>
        <p style="font-size: 13px; color: #0095c7; font-weight: 700; margin: 2px 0 0 0;">Clinical Intelligence Platform — Executive MIS Analytics Report</p>
        <p style="font-size: 11px; color: #64748b; margin: 3px 0 0 0;">Dataset: ${state.currentDataset.name} • Domain: ${state.currentDataset.detected_type}</p>
      </div>
      <div style="text-align: right; font-size: 11px; color: #64748b;">
        <p style="margin: 0;">Database: <strong>MIS_DB</strong></p>
        <p style="margin: 2px 0 0 0;">Generated: <strong>${new Date().toLocaleDateString('en-IN', {day:'2-digit', month:'short', year:'numeric'})}</strong></p>
      </div>
    </div>

    <!-- KPI Box Grid -->
    <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 12px; margin-bottom: 20px;">
      <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px;">
        <span style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Total Volume</span>
        <h2 style="font-size: 18px; font-weight: 800; color: #0095c7; margin: 4px 0 0 0;">${formatCurrency(kpis.totalValue)}</h2>
      </div>
      <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px;">
        <span style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Total Transactions</span>
        <h2 style="font-size: 18px; font-weight: 800; color: #00c896; margin: 4px 0 0 0;">${Number(kpis.totalRecords).toLocaleString()}</h2>
      </div>
      <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px;">
        <span style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Avg Ticket Size</span>
        <h2 style="font-size: 18px; font-weight: 800; color: #0284c7; margin: 4px 0 0 0;">${formatCurrency(kpis.averageTicket)}</h2>
      </div>
      <div style="background: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 12px;">
        <span style="font-size: 10px; font-weight: 700; color: #64748b; text-transform: uppercase;">Peak Transaction</span>
        <h2 style="font-size: 18px; font-weight: 800; color: #b45309; margin: 4px 0 0 0;">${formatCurrency(kpis.highestTransaction)}</h2>
      </div>
    </div>

    <!-- AI Observations -->
    <div style="background: #f0fdf4; border-left: 4px solid #00c896; border-radius: 6px; padding: 12px 16px; margin-bottom: 20px;">
      <h3 style="font-size: 12px; font-weight: 700; color: #065f46; text-transform: uppercase; margin: 0 0 6px 0;">Executive MIS Insights & Clinical Findings</h3>
      <ul style="margin: 0; padding-left: 18px; font-size: 11px; line-height: 1.6; color: #334155;">
        ${(state.analytics.insights || []).map(i => `<li>${i.replace(/\*\*/g, '')}</li>`).join('')}
      </ul>
    </div>

    <!-- Performance Table -->
    <div style="margin-bottom: 24px;">
      <h3 style="font-size: 13px; font-weight: 700; color: #0f172a; margin: 0 0 8px 0;">Top Category / Dimension Breakdown</h3>
      <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
        <thead>
          <tr style="background: #e2e8f0; color: #334155;">
            <th style="padding: 6px 10px; text-align: left; border: 1px solid #cbd5e1;">Category</th>
            <th style="padding: 6px 10px; text-align: right; border: 1px solid #cbd5e1;">Count</th>
            <th style="padding: 6px 10px; text-align: right; border: 1px solid #cbd5e1;">Total Value</th>
            <th style="padding: 6px 10px; text-align: right; border: 1px solid #cbd5e1;">Average</th>
            <th style="padding: 6px 10px; text-align: right; border: 1px solid #cbd5e1;">Share</th>
          </tr>
        </thead>
        <tbody>
          ${cats.slice(0, 10).map(c => `
            <tr>
              <td style="padding: 6px 10px; border: 1px solid #e2e8f0;"><strong>${c.category}</strong></td>
              <td style="padding: 6px 10px; text-align: right; border: 1px solid #e2e8f0;">${Number(c.volume).toLocaleString()}</td>
              <td style="padding: 6px 10px; text-align: right; border: 1px solid #e2e8f0;">${formatCurrency(c.totalValue)}</td>
              <td style="padding: 6px 10px; text-align: right; border: 1px solid #e2e8f0;">${formatCurrency(c.avgValue)}</td>
              <td style="padding: 6px 10px; text-align: right; border: 1px solid #e2e8f0;">${c.sharePct}%</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>

    <div style="display: flex; justify-content: space-between; border-top: 1px solid #e2e8f0; padding-top: 16px; font-size: 10px; color: #94a3b8;">
      <span>Generated by MIS Analytics Pro • MySQL Database MIS_DB</span>
      <span>Management Sign-off: _____________________</span>
    </div>
  `;

  document.body.appendChild(reportWrap);

  // Use html2pdf if available, otherwise window.print
  if (window.html2pdf) {
    const opt = {
      margin: 10,
      filename: `${state.currentDataset.name || 'MIS_Report'}_Executive.pdf`,
      image: { type: 'jpeg', quality: 0.98 },
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };
    await html2pdf().set(opt).from(reportWrap).save();
    document.body.removeChild(reportWrap);
    showToast('✓ Executive PDF Report downloaded!');
  } else {
    document.body.removeChild(reportWrap);
    window.print();
  }
}

async function deleteCurrentDataset() {
  if (!state.currentDatasetId) return;
  if (!confirm(`Are you sure you want to permanently delete "${state.currentDataset?.name}" from MySQL MIS_DB?`)) {
    return;
  }

  try {
    const res = await fetch(`/api/datasets/${state.currentDatasetId}`, { method: 'DELETE' });
    const data = await res.json();
    if (data.success) {
      showToast('Dataset deleted from MIS_DB.');
      state.currentDatasetId = null;
      await fetchDatasetsList();
      await checkSystemStatus();
    }
  } catch (err) {
    showToast('Delete Error: ' + err.message);
  }
}

function exportPivotToExcel() {
  const table = document.getElementById('pivotMatrixTable');
  if (!table) {
    showToast('No pivot matrix generated yet.');
    return;
  }
  try {
    const wb = XLSX.utils.table_to_book(table, { sheet: 'Pivot Matrix' });
    XLSX.writeFile(wb, `Pivot_Matrix_${state.currentDataset?.name || 'Hospital'}.xlsx`);
    showToast('✓ Pivot table exported to Excel successfully!');
  } catch (err) {
    console.error('[Pivot Export Error]', err);
    showToast('Failed to export pivot table: ' + err.message);
  }
}

function exportGridToExcel() {
  if (!state.rawGridRows || state.rawGridRows.length === 0) {
    showToast('No raw data loaded to export.');
    return;
  }
  try {
    const ws = XLSX.utils.json_to_sheet(state.rawGridRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data Explorer');
    XLSX.writeFile(wb, `Data_Explorer_${state.currentDataset?.name || 'Hospital'}.xlsx`);
    showToast(`✓ Exported ${state.rawGridRows.length} rows to Excel!`);
  } catch (err) {
    console.error('[Grid Export Error]', err);
    showToast('Failed to export data grid: ' + err.message);
  }
}

// ==========================================
// EVENT LISTENERS & SETUP
// ==========================================
function setupEventListeners() {
  // Download Action Buttons
  document.getElementById('btnDownloadExcelReport')?.addEventListener('click', downloadFullExcelReport);
  document.getElementById('btnDownloadPdfReport')?.addEventListener('click', downloadExecutivePdfReport);
  document.getElementById('btnPrintReport')?.addEventListener('click', () => window.print());

  // Metric Switcher
  document.getElementById('activeMetricSelect')?.addEventListener('change', (e) => {
    state.selectedMetricCol = e.target.value;
    triggerAdvancedFilterRecompute();
  });

  // Dimension Switcher
  document.getElementById('activeDimensionSelect')?.addEventListener('change', (e) => {
    state.selectedDimCol = e.target.value;
    state.quickCategoryFilter = 'All';
    triggerAdvancedFilterRecompute();
  });

  // Date Preset Buttons
  document.querySelectorAll('.pill-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      handleDatePresetClick(btn.getAttribute('data-range'));
    });
  });

  // Quick Category Filter
  document.getElementById('filterDimension')?.addEventListener('change', (e) => {
    state.quickCategoryFilter = e.target.value;
    triggerAdvancedFilterRecompute();
  });

  // Custom Date Pickers
  document.getElementById('filterDateStart')?.addEventListener('change', (e) => {
    state.dateRange.start = e.target.value;
  });
  document.getElementById('filterDateEnd')?.addEventListener('change', (e) => {
    state.dateRange.end = e.target.value;
  });

  // Multi-rules Drawer Toggle
  document.getElementById('btnToggleAdvancedFilters')?.addEventListener('click', () => {
    const drawer = document.getElementById('multiRulesDrawer');
    if (drawer) {
      drawer.style.display = drawer.style.display === 'none' ? 'block' : 'none';
      if (drawer.style.display === 'block' && document.querySelectorAll('.rule-row').length === 0) {
        addMultiRuleRow();
      }
    }
  });

  document.getElementById('btnAddRuleRow')?.addEventListener('click', addMultiRuleRow);
  document.getElementById('btnApplyFilter')?.addEventListener('click', triggerAdvancedFilterRecompute);
  document.getElementById('btnResetFilter')?.addEventListener('click', resetAllFilters);

  // Column Inspector Modal
  document.getElementById('btnToggleColInspector')?.addEventListener('click', openColumnInspector);
  document.getElementById('btnCloseColInspector')?.addEventListener('click', closeModals);

  // Dataset Switcher
  document.getElementById('activeDatasetSelect')?.addEventListener('change', (e) => {
    if (e.target.value) {
      localStorage.setItem('mis_active_dataset_id', e.target.value);
      if (window.location.pathname.endsWith('home.html')) {
        window.location.href = 'overview.html';
      } else {
        loadDataset(e.target.value);
      }
    }
  });

  // Quick Sample Modal
  document.getElementById('btnQuickSample')?.addEventListener('click', () => {
    document.getElementById('sampleModal').style.display = 'flex';
  });
  document.getElementById('btnCloseSampleModal')?.addEventListener('click', closeModals);

  // Upload Modal
  document.getElementById('btnOpenUploadModal')?.addEventListener('click', () => {
    document.getElementById('uploadModal').style.display = 'flex';
  });
  document.getElementById('btnCloseUploadModal')?.addEventListener('click', closeModals);
  document.getElementById('btnCancelUpload')?.addEventListener('click', closeModals);
  document.getElementById('btnCloseImportError')?.addEventListener('click', closeModals);

  // Modal Upload Form
  document.getElementById('modalUploadForm')?.addEventListener('submit', (e) => {
    e.preventDefault();
    const file = selectedUploadFile || document.getElementById('modalFileInput')?.files[0];
    const nameInp = document.getElementById('modalDatasetName');
    const sheetSelect = document.getElementById('modalSheetSelect');
    if (file) {
      handleFileUpload(file, nameInp ? nameInp.value : '', sheetSelect ? sheetSelect.value : '');
    } else {
      showToast('Please select an Excel or CSV file first.');
    }
  });

  // Tabs Navigation
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

      btn.classList.add('active');
      const targetTab = btn.getAttribute('data-tab');
      const content = document.getElementById(targetTab);
      if (content) content.classList.add('active');

      if (targetTab === 'tab-pivot' && !document.getElementById('pivotTbody').hasChildNodes()) {
        runPivot();
      }
      if (targetTab === 'tab-builder') {
        populateVisualBuilderControls();
        if (!state.currentBuilderResult) {
          renderCustomVisualStudio();
        }
      }
      if (targetTab === 'tab-doctor') {
        if (state.analytics) renderDoctorAnalytics(state.analytics);
      }
    });
  });

  // Ask BI Copilot Events
  document.getElementById('btnAskBiSubmit')?.addEventListener('click', () => handleAskBIQuery());
  document.getElementById('copilotQueryInput')?.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      handleAskBIQuery();
    }
  });

  document.querySelectorAll('.copilot-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      const q = chip.getAttribute('data-query');
      if (q) {
        handleAskBIQuery(q);
      }
    });
  });

  document.querySelectorAll('#askBiTypeSelector .btn-tool-type').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#askBiTypeSelector .btn-tool-type').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      const chartType = btn.getAttribute('data-type');
      state.askBiChartType = chartType;
      if (state.currentAskBiResult && state.currentAskBiResult.visual) {
        renderAskBiChart(state.currentAskBiResult.visual, chartType);
      }
    });
  });

  document.getElementById('btnExportAskBiCsv')?.addEventListener('click', exportAskBiCsv);

  // Visual Builder Studio Events
  document.querySelectorAll('#visualTypeSelector .vtype-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('#visualTypeSelector .vtype-btn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      state.selectedVisualType = btn.getAttribute('data-vtype') || 'bar';
      if (state.currentBuilderResult) {
        renderBuilderChart(state.currentBuilderResult, state.selectedVisualType);
      }
    });
  });

  document.getElementById('btnBuildVisual')?.addEventListener('click', renderCustomVisualStudio);
  document.getElementById('btnExportBuilderData')?.addEventListener('click', exportBuilderCsv);

  // Pivot & Export
  document.getElementById('btnRunPivot')?.addEventListener('click', runPivot);
  document.getElementById('btnExportPivotExcel')?.addEventListener('click', exportPivotToExcel);

  // Data Grid Search & Pagination
  let searchTimer;
  document.getElementById('gridSearchInput')?.addEventListener('input', (e) => {
    clearTimeout(searchTimer);
    searchTimer = setTimeout(() => {
      state.gridSearch = e.target.value;
      state.rawGridPage = 1;
      loadRawDataGrid();
    }, 300);
  });

  document.getElementById('btnPrevPage')?.addEventListener('click', () => {
    if (state.rawGridPage > 1) {
      state.rawGridPage--;
      loadRawDataGrid();
    }
  });

  document.getElementById('btnNextPage')?.addEventListener('click', () => {
    state.rawGridPage++;
    loadRawDataGrid();
  });

  document.getElementById('btnExportGrid')?.addEventListener('click', exportGridToExcel);
  document.getElementById('btnDeleteDataset')?.addEventListener('click', deleteCurrentDataset);

  // Doctor Analytics Search, Filter & Excel Export Events
  let docSearchTimer;
  document.getElementById('docSearchInput')?.addEventListener('input', (e) => {
    clearTimeout(docSearchTimer);
    docSearchTimer = setTimeout(() => {
      state.doctorSearchTerm = e.target.value;
      renderDoctorTableRows();
    }, 200);
  });

  document.getElementById('docDeptFilterSelect')?.addEventListener('change', (e) => {
    state.doctorDeptFilter = e.target.value;
    renderDoctorTableRows();
  });

  document.getElementById('btnExportDoctorExcel')?.addEventListener('click', exportDoctorExcel);
}

function closeModals() {
  document.querySelectorAll('.modal-backdrop').forEach(m => m.style.display = 'none');
}

// ==========================================
// UTILITY FUNCTIONS
// ==========================================
function formatCurrency(val) {
  if (val === null || val === undefined || isNaN(val)) return '₹ 0.00';
  return '₹ ' + Number(val).toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatCompactNumber(num) {
  if (!num || isNaN(num)) return '0';
  if (num >= 10000000) return (num / 10000000).toFixed(1) + ' Cr';
  if (num >= 100000) return (num / 100000).toFixed(1) + ' L';
  if (num >= 1000) return (num / 1000).toFixed(1) + 'k';
  return Number(num).toLocaleString();
}

function formatColName(name) {
  if (!name) return '';
  return name.replace(/_/g, ' ').replace(/\b\w/g, l => l.toUpperCase());
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function showToast(msg) {
  const toast = document.getElementById('toast');
  if (!toast) return;
  toast.textContent = msg;
  toast.style.display = 'block';
  setTimeout(() => {
    toast.style.display = 'none';
  }, 3500);
}

// ==========================================
// DIMENSION QUICK SWITCHER PILLS BAR
// ==========================================
function renderDimensionPills(analytics) {
  const container = document.getElementById('dimensionPillsContainer');
  const bar = document.getElementById('dimensionPillsBar');
  if (!container) return;

  const dims = (analytics && analytics.availableDimensions && analytics.availableDimensions.length > 0)
    ? analytics.availableDimensions
    : (analytics && analytics.columns)
      ? analytics.columns.filter(c => c.type === 'category' || c.type === 'text').map(c => ({
          col: c.name,
          label: c.name,
          icon: '🏷️',
          uniqueCount: c.stats?.uniqueCount || 0
        }))
      : [];

  if (dims.length === 0) {
    if (bar) bar.style.display = 'none';
    return;
  }
  if (bar) bar.style.display = 'flex';

  const currentDim = state.selectedDimCol || analytics.summaryKPIs?.dimensionColumn;

  container.innerHTML = dims.map(d => {
    const isActive = d.col === currentDim;
    return `
      <button type="button" class="dim-pill ${isActive ? 'active' : ''}" data-col="${escapeHtml(d.col)}" title="Group MIS by ${escapeHtml(d.label)} (${d.uniqueCount} items)">
        <span class="dim-pill-icon">${d.icon || '🏷️'}</span>
        <span class="dim-pill-name">${escapeHtml(d.label)}</span>
        <span class="dim-pill-count">${Number(d.uniqueCount).toLocaleString()}</span>
      </button>
    `;
  }).join('');

  container.querySelectorAll('.dim-pill').forEach(pill => {
    pill.addEventListener('click', () => {
      const col = pill.getAttribute('data-col');
      if (!col) return;
      state.selectedDimCol = col;
      state.quickCategoryFilter = 'All';

      // Keep activeDimensionSelect dropdown in sync
      const dimSelect = document.getElementById('activeDimensionSelect');
      if (dimSelect) dimSelect.value = col;

      container.querySelectorAll('.dim-pill').forEach(p => p.classList.toggle('active', p.getAttribute('data-col') === col));
      triggerAdvancedFilterRecompute();
    });
  });
}

// ==========================================
// DOCTOR MIS ANALYTICS HUB
// ==========================================
function renderDoctorAnalytics(analytics) {
  if (!analytics) return;

  const docs = analytics.doctorPerformance || [];
  const kpis = analytics.doctorKPIs || {};

  // Update tab badge
  const badge = document.getElementById('doctorCountBadge');
  if (badge) {
    badge.textContent = docs.length;
    badge.style.display = docs.length > 0 ? 'inline-block' : 'none';
  }

  // Populate Hero KPIs
  const topRev = kpis.topRevenueDoctor;
  const topVol = kpis.topVolumeDoctor;
  const topAvg = kpis.topAvgTicketDoctor;

  const topRevVal = document.getElementById('kpiDocTopRevenueVal');
  const topRevName = document.getElementById('kpiDocTopRevenueName');
  if (topRevVal) topRevVal.textContent = topRev ? formatCurrency(topRev.net) : '₹ 0.00';
  if (topRevName) topRevName.textContent = topRev ? `${topRev.name} • ${topRev.department || 'General'}` : 'Analyzing doctors...';

  const topVolVal = document.getElementById('kpiDocTopVolumeVal');
  const topVolName = document.getElementById('kpiDocTopVolumeName');
  if (topVolVal) topVolVal.textContent = topVol ? `${Number(topVol.patients).toLocaleString()} Patients` : '0 Patients';
  if (topVolName) topVolName.textContent = topVol ? `${topVol.name} • ${topVol.department || 'General'}` : 'Top by footfall';

  const topAvgVal = document.getElementById('kpiDocTopAvgVal');
  const topAvgName = document.getElementById('kpiDocTopAvgName');
  if (topAvgVal) topAvgVal.textContent = topAvg ? formatCurrency(topAvg.avgTicket) : '₹ 0.00';
  if (topAvgName) topAvgName.textContent = topAvg ? `${topAvg.name} • ${topAvg.department || 'General'}` : 'Avg per patient';

  const totalDocEl = document.getElementById('kpiDocTotalCount');
  const totalDeptEl = document.getElementById('kpiDocTotalDept');
  const depts = Array.from(new Set(docs.map(d => d.department).filter(Boolean))).sort();
  if (totalDocEl) totalDocEl.textContent = `${docs.length} Specialists`;
  if (totalDeptEl) totalDeptEl.textContent = `Across ${depts.length} Medical Departments`;

  // Populate Department Filter Dropdown if not already populated
  const deptSelect = document.getElementById('docDeptFilterSelect');
  if (deptSelect && deptSelect.options.length <= 1) {
    const curVal = deptSelect.value;
    deptSelect.innerHTML = '<option value="">All Departments</option>' + depts.map(d => `<option value="${escapeHtml(d)}">${escapeHtml(d)}</option>`).join('');
    if (curVal) deptSelect.value = curVal;
  }

  // Update Subtitle
  const subEl = document.getElementById('docTableSubtitle');
  if (subEl) {
    const totalRev = docs.reduce((sum, d) => sum + (d.totalNet || 0), 0);
    subEl.textContent = `Analyzing ${docs.length} doctors across ${depts.length} departments • Total Net Revenue: ${formatCurrency(totalRev)}`;
  }

  // Render Charts
  renderDoctorCharts(analytics);

  // Render Doctor Table
  renderDoctorTableRows();
}

function renderDoctorCharts(analytics) {
  const docs = analytics.doctorPerformance || [];
  if (docs.length === 0) return;

  const isDark = document.documentElement.getAttribute('data-theme') === 'dark';
  const gridColor = isDark ? 'rgba(255, 255, 255, 0.08)' : 'rgba(0, 0, 0, 0.06)';
  const textColor = isDark ? '#94a3b8' : '#475569';

  // 1. Top 15 Doctors by Net Revenue (Horizontal Bar Chart)
  destroyChart('docRevenueChart');
  const revCanvas = document.getElementById('docRevenueChart');
  if (revCanvas) {
    const top15Rev = docs.slice(0, 15);
    const ctx = revCanvas.getContext('2d');
    const grad = ctx.createLinearGradient(0, 0, 400, 0);
    grad.addColorStop(0, 'rgba(0, 149, 199, 0.9)');
    grad.addColorStop(1, 'rgba(0, 200, 150, 0.85)');

    state.charts.docRevenueChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: top15Rev.map(d => d.doctor),
        datasets: [{
          label: 'Total Net Revenue (₹)',
          data: top15Rev.map(d => d.totalNet),
          backgroundColor: grad,
          borderRadius: 6,
          borderSkipped: false
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                const item = top15Rev[ctx.dataIndex];
                return [
                  ` Net Revenue: ${formatCurrency(item.totalNet)}`,
                  ` Patients: ${Number(item.patientVolume).toLocaleString()}`,
                  ` Avg Ticket: ${formatCurrency(item.avgTicket)}`,
                  ` Department: ${item.department || 'General'}`
                ];
              }
            }
          }
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              callback: val => formatCompactNumber(val)
            }
          },
          y: {
            grid: { display: false },
            ticks: {
              color: textColor,
              font: { weight: '600', size: 11 },
              callback: function(val, index) {
                const name = this.getLabelForValue(index);
                return name.length > 22 ? name.substring(0, 20) + '…' : name;
              }
            }
          }
        }
      }
    });
  }

  // 2. Doctor Patient Volume vs Revenue Contribution (Top 15 by Volume)
  destroyChart('docVolumeChart');
  const volCanvas = document.getElementById('docVolumeChart');
  if (volCanvas) {
    const top15Vol = [...docs].sort((a, b) => b.patientVolume - a.patientVolume).slice(0, 15);
    const ctx = volCanvas.getContext('2d');

    state.charts.docVolumeChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: top15Vol.map(d => d.doctor),
        datasets: [
          {
            type: 'bar',
            label: 'Patient Volume (Count)',
            data: top15Vol.map(d => d.patientVolume),
            backgroundColor: 'rgba(16, 185, 129, 0.75)',
            borderRadius: 5,
            yAxisID: 'y',
            order: 2
          },
          {
            type: 'line',
            label: 'Net Revenue (₹)',
            data: top15Vol.map(d => d.totalNet),
            borderColor: '#0095c7',
            backgroundColor: 'rgba(0, 149, 199, 0.2)',
            borderWidth: 2.5,
            pointRadius: 4,
            pointHoverRadius: 6,
            tension: 0.35,
            yAxisID: 'y1',
            order: 1
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: true,
            position: 'top',
            labels: { color: textColor, boxWidth: 12 }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => {
                if (ctx.dataset.yAxisID === 'y') {
                  return ` Patients: ${Number(ctx.raw).toLocaleString()}`;
                }
                return ` Net Revenue: ${formatCurrency(ctx.raw)}`;
              }
            }
          }
        },
        scales: {
          x: {
            grid: { color: gridColor },
            ticks: {
              color: textColor,
              maxRotation: 35,
              callback: function(val, index) {
                const name = this.getLabelForValue(index);
                return name.length > 14 ? name.substring(0, 12) + '…' : name;
              }
            }
          },
          y: {
            type: 'linear',
            position: 'left',
            grid: { color: gridColor },
            title: { display: true, text: 'Patients', color: textColor },
            ticks: { color: textColor }
          },
          y1: {
            type: 'linear',
            position: 'right',
            grid: { drawOnChartArea: false },
            title: { display: true, text: 'Revenue (₹)', color: textColor },
            ticks: {
              color: textColor,
              callback: val => formatCompactNumber(val)
            }
          }
        }
      }
    });
  }
}

function renderDoctorTableRows() {
  const tbody = document.getElementById('doctorPerformanceTbody');
  if (!tbody || !state.analytics) return;

  const allDocs = state.analytics.doctorPerformance || [];
  if (allDocs.length === 0) {
    tbody.innerHTML = '<tr><td colspan="12" class="text-center" style="padding: 2.5rem; color: var(--text-muted);">No doctor performance metrics found in this dataset.</td></tr>';
    return;
  }

  const search = (state.doctorSearchTerm || '').toLowerCase().trim();
  const deptFilter = state.doctorDeptFilter || '';

  const filtered = allDocs.filter(d => {
    if (deptFilter && d.department !== deptFilter) return false;
    if (search) {
      const matchDoc = String(d.doctor).toLowerCase().includes(search);
      const matchDept = String(d.department).toLowerCase().includes(search);
      if (!matchDoc && !matchDept) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="12" class="text-center" style="padding: 2.5rem; color: var(--text-muted);">No doctors match your search or department filter.</td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map(doc => {
    let rankBadge = `<span class="doc-rank-pill">#${doc.rank}</span>`;
    if (doc.rank === 1) rankBadge = `<span class="doc-rank-pill top-1">🥇 #1</span>`;
    else if (doc.rank === 2) rankBadge = `<span class="doc-rank-pill top-2">🥈 #2</span>`;
    else if (doc.rank === 3) rankBadge = `<span class="doc-rank-pill top-3">🥉 #3</span>`;

    const cleanName = (doc.doctor || 'Dr').replace(/^dr\.?\s*/i, '');
    const initials = cleanName.slice(0, 2).toUpperCase() || 'DR';
    const pendingColor = (doc.totalPending && doc.totalPending > 0) ? '#f43f5e' : 'inherit';

    return `
      <tr>
        <td>${rankBadge}</td>
        <td>
          <span class="doc-avatar-badge">${initials}</span>
          <strong style="color: var(--text-main);">${escapeHtml(doc.doctor)}</strong>
        </td>
        <td><span class="badge-tag" style="background: rgba(14, 165, 233, 0.15); color: #38bdf8;">${escapeHtml(doc.department || 'General')}</span></td>
        <td class="text-right" style="font-weight: 600;">${Number(doc.patientVolume).toLocaleString()}</td>
        <td class="text-right"><strong style="color: #0095c7;">${formatCurrency(doc.totalNet)}</strong></td>
        <td class="text-right">${formatCurrency(doc.totalBilled)}</td>
        <td class="text-right" style="color: #f59e0b;">${formatCurrency(doc.totalDiscount)}</td>
        <td class="text-right" style="color: #10b981;">${formatCurrency(doc.totalPaid)}</td>
        <td class="text-right" style="color: ${pendingColor}; font-weight: ${doc.totalPending > 0 ? '700' : 'normal'};">${formatCurrency(doc.totalPending)}</td>
        <td class="text-right" style="font-weight: 600;">${formatCurrency(doc.avgTicket)}</td>
        <td>
          <div class="doc-share-cell">
            <span class="doc-share-val">${doc.sharePct}%</span>
            <div class="doc-share-track">
              <div class="doc-share-fill" style="width: ${Math.min(100, Math.max(3, doc.sharePct * 6))}%;"></div>
            </div>
          </div>
        </td>
        <td class="text-center">
          <button type="button" class="btn btn-secondary btn-sm" onclick="drilldownToDoctor('${escapeHtml(doc.doctor)}')" title="Filter entire MIS dashboard by ${escapeHtml(doc.doctor)}" style="padding: 0.25rem 0.6rem; font-size: 0.76rem; font-weight: 600;">
            🔍 Drilldown
          </button>
        </td>
      </tr>
    `;
  }).join('');
}

function drilldownToDoctor(doctorName) {
  if (!doctorName) return;

  const doctorCol = state.analytics?.doctorKPIs?.doctorColumnName || 'PRIMARY DOCTOR';
  state.selectedDimCol = doctorCol;
  state.quickCategoryFilter = doctorName;

  // Sync Dimension Dropdown
  const dimSelect = document.getElementById('activeDimensionSelect');
  if (dimSelect) dimSelect.value = doctorCol;

  // Switch tab to Executive Summary
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));

  const overviewBtn = document.querySelector('.tab-btn[data-tab="tab-overview"]');
  const overviewContent = document.getElementById('tab-overview');
  if (overviewBtn) overviewBtn.classList.add('active');
  if (overviewContent) overviewContent.classList.add('active');

  showToast(`Drilling down entire dashboard for: ${doctorName}`);
  triggerAdvancedFilterRecompute();
}

function exportDoctorExcel() {
  if (!state.analytics || !state.analytics.doctorPerformance || state.analytics.doctorPerformance.length === 0) {
    showToast('No doctor performance data available to export.');
    return;
  }

  const docs = state.analytics.doctorPerformance;
  const rows = docs.map(d => ({
    'Rank': d.rank,
    'Doctor Name': d.doctor,
    'Department': d.department || 'General',
    'Patient Volume': d.patientVolume,
    'Total Net Amount (₹)': d.totalNet,
    'Billed Amount (₹)': d.totalBilled,
    'Discount Amount (₹)': d.totalDiscount,
    'Paid Amount (₹)': d.totalPaid,
    'Pending Amount (₹)': d.totalPending,
    'Avg Ticket (₹)': d.avgTicket,
    'Revenue Share (%)': d.sharePct
  }));

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);

  ws['!cols'] = [
    { wch: 8 },
    { wch: 30 },
    { wch: 22 },
    { wch: 15 },
    { wch: 20 },
    { wch: 18 },
    { wch: 16 },
    { wch: 18 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 }
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Doctor MIS Performance');
  const filename = `Doctor_MIS_Analytics_${state.currentDataset?.name || 'Hospital'}.xlsx`;
  XLSX.writeFile(wb, filename);
  showToast(`✓ Exported ${docs.length} doctors to ${filename}!`);
}

// Global window bindings for inline HTML event handlers
window.handleFileUpload = handleFileUpload;
window.inspectAndPreviewFile = inspectAndPreviewFile;
window.clearSelectedFilePreview = clearSelectedFilePreview;
window.loadSampleDataset = loadSampleDataset;
window.drilldownToDoctor = drilldownToDoctor;
window.exportDoctorExcel = exportDoctorExcel;

// ==========================================
// PEERLESS HOSPITAL SIDEBAR NAVIGATION
// ==========================================

(function initSidebarNav() {
  // Wait for DOM
  function setupSidebar() {
    const sidebar = document.getElementById('appSidebar');
    const toggleBtn = document.getElementById('sidebarToggle');

    // Sidebar collapse / expand controls
    const appLayout = document.querySelector('.app-layout');
    const topToggleBtn = document.getElementById('topbarSidebarToggle');
    const brandLogo = document.getElementById('sidebarBrandLogo');

    function toggleSidebar() {
      if (!sidebar) return;
      const isCollapsed = sidebar.classList.toggle('collapsed');
      appLayout?.classList.toggle('sidebar-collapsed', isCollapsed);
      localStorage.setItem('sidebar_collapsed', isCollapsed);
      if (toggleBtn) toggleBtn.title = isCollapsed ? 'Expand sidebar (Boro korun)' : 'Collapse sidebar (Chhoto korun)';
      if (topToggleBtn) topToggleBtn.title = isCollapsed ? 'Expand sidebar (Boro korun)' : 'Collapse sidebar (Chhoto korun)';
    }

    // Check saved state
    const savedCollapsed = localStorage.getItem('sidebar_collapsed') === 'true';
    if (savedCollapsed && sidebar) {
      sidebar.classList.add('collapsed');
      appLayout?.classList.add('sidebar-collapsed');
    }

    if (toggleBtn) {
      toggleBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleSidebar();
      });
    }

    if (topToggleBtn) {
      topToggleBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        toggleSidebar();
      });
    }

    if (brandLogo) {
      brandLogo.addEventListener('click', (e) => {
        if (sidebar && sidebar.classList.contains('collapsed')) {
          e.stopPropagation();
          toggleSidebar();
        }
      });
    }

    // Mobile Drawer toggle
    const mobileToggle = document.getElementById('mobileMenuToggle');
    const backdrop = document.getElementById('sidebarBackdrop');
    if (mobileToggle && sidebar && backdrop) {
      mobileToggle.addEventListener('click', () => {
        sidebar.classList.toggle('mobile-open');
        backdrop.classList.toggle('active', sidebar.classList.contains('mobile-open'));
      });
      backdrop.addEventListener('click', () => {
        sidebar.classList.remove('mobile-open');
        backdrop.classList.remove('active');
      });
    }

    // Nav item tab switching
    const navItems = document.querySelectorAll('.sidebar-nav-item[data-tab]');
    navItems.forEach(item => {
      item.addEventListener('click', () => {
        const targetTab = item.getAttribute('data-tab');
        if (sidebar && sidebar.classList.contains('mobile-open')) {
          sidebar.classList.remove('mobile-open');
          backdrop?.classList.remove('active');
        }
        switchToSidebarTab(targetTab, item);
      });
    });

    // Wire top-bar buttons to original button handlers
    const topExcelBtn = document.getElementById('btnDownloadExcelReportTop');
    const topPdfBtn = document.getElementById('btnDownloadPdfReportTop');
    const topDeleteBtn = document.getElementById('btnDeleteDatasetTop');
    const topColBtn = document.getElementById('btnToggleColInspectorTop');

    // Get the original buttons (in workspace-header, hidden)
    if (topExcelBtn) topExcelBtn.addEventListener('click', () => {
      document.getElementById('btnDownloadExcelReport')?.click();
    });
    if (topPdfBtn) topPdfBtn.addEventListener('click', () => {
      document.getElementById('btnDownloadPdfReport')?.click();
    });
    if (topDeleteBtn) topDeleteBtn.addEventListener('click', () => {
      document.getElementById('btnDeleteDataset')?.click();
    });
    if (topColBtn) topColBtn.addEventListener('click', () => {
      document.getElementById('btnToggleColInspector')?.click();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', setupSidebar);
  } else {
    setupSidebar();
  }
})();

// Switch sidebar tabs
function switchToSidebarTab(targetTab, clickedItem) {
  // Update nav item active states
  document.querySelectorAll('.sidebar-nav-item').forEach(i => i.classList.remove('active'));
  if (clickedItem) {
    clickedItem.classList.add('active');
  } else {
    const item = document.querySelector(`.sidebar-nav-item[data-tab="${targetTab}"]`);
    if (item) item.classList.add('active');
  }

  if (targetTab === 'home') {
    // Show hero / upload section
    const hero = document.getElementById('welcomeHero');
    const ws = document.getElementById('dashboardWorkspace');
    if (hero) hero.style.display = 'flex';
    if (ws) ws.style.display = 'none';
    updateTopBarTitle('🏥 Peerless Hospital', 'Upload ERP Data to MIS_DB');
    toggleTopBarDashboardBtns(false);
    return;
  }

  // Ensure dashboard workspace is visible and hero hidden
  const hero = document.getElementById('welcomeHero');
  const ws = document.getElementById('dashboardWorkspace');
  if (hero) hero.style.display = 'none';
  if (ws) ws.style.display = 'block';
  toggleTopBarDashboardBtns(true);

  // Switch tab content
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));

  const tabContent = document.getElementById(targetTab);
  if (tabContent) {
    tabContent.classList.add('active');
    // Scroll smoothly to top of workspace
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // Also update old .tab-btn if it exists (for backward compat with drilldown)
  const oldTabBtn = document.querySelector(`.tab-btn[data-tab="${targetTab}"]`);
  if (oldTabBtn) oldTabBtn.classList.add('active');

  // Trigger tab-specific initialization & rendering
  try {
    if (targetTab === 'tab-doctor') {
      if (state.analytics) renderDoctorAnalytics(state.analytics);
    } else if (targetTab === 'tab-charts') {
      if (state.analytics) renderCharts(state.analytics);
    } else if (targetTab === 'tab-pivot') {
      runPivot();
    } else if (targetTab === 'tab-builder') {
      populateVisualBuilderControls();
      if (!state.currentBuilderResult) {
        renderCustomVisualStudio();
      }
    } else if (targetTab === 'tab-datagrid') {
      loadRawDataGrid();
    } else if (targetTab === 'tab-anomalies') {
      if (state.analytics) renderAnomalies(state.analytics);
    } else if (targetTab === 'tab-overview') {
      if (state.analytics) {
        renderKPIs(state.analytics);
        renderInsights(state.analytics);
        renderCharts(state.analytics);
      }
    }
  } catch (renderErr) {
    console.warn(`[Tab Switch Render Error on ${targetTab}]`, renderErr);
  }

  // Update top bar title
  const tabTitles = {
    'tab-overview': ['📊 Executive Summary', 'KPIs, Charts & Insights'],
    'tab-doctor': ['👨‍⚕️ Doctor MIS Analytics', 'Revenue, Volume & Performance'],
    'tab-charts': ['📈 Visual Analytics Gallery', 'Charts & Performance Breakdown'],
    'tab-pivot': ['🔢 Dynamic Pivot Matrix', 'Multi-dimensional Analysis'],
    'tab-ask-bi': ['⚡ Ask BI Copilot', 'Natural Language Analytics'],
    'tab-builder': ['🎨 Visual Studio Builder', 'Custom Chart Designer'],
    'tab-anomalies': ['⚠️ Audit & Anomaly Watch', 'Flagged Exceptions & Discrepancies'],
    'tab-datagrid': ['🗄️ Data Explorer', 'Raw Data & Search'],
  };
  const [title, subtitle] = tabTitles[targetTab] || ['🏥 Peerless Hospital', 'Clinical MIS Intelligence'];
  updateTopBarTitle(title, subtitle);
}

window.switchToSidebarTab = switchToSidebarTab;

function updateTopBarTitle(title, subtitle) {
  const titleEl = document.getElementById('topBarTitle');
  if (titleEl) titleEl.textContent = title;

  const subEl = document.querySelector('.topbar-subtitle');
  if (subEl) subEl.textContent = subtitle || 'B.K. Roy Research Centre — Clinical MIS Intelligence';
}

function toggleTopBarDashboardBtns(show) {
  document.querySelectorAll('.dashboard-only-btn').forEach(btn => {
    btn.style.display = show ? 'inline-flex' : 'none';
  });
}

// Override showHero to also update sidebar active state and top bar
const _originalShowHero = showHero;
window.showHero = function(show) {
  _originalShowHero(show);
  toggleTopBarDashboardBtns(!show);

  if (show) {
    // Activate home nav item
    document.querySelectorAll('.sidebar-nav-item').forEach(i => i.classList.remove('active'));
    const homeItem = document.getElementById('navHome');
    if (homeItem) homeItem.classList.add('active');
    updateTopBarTitle('🏥 Peerless Hospital', 'Upload ERP Excel Data to MIS_DB');
  } else {
    // Activate overview nav item if nothing else is active
    const hasActive = [...document.querySelectorAll('.sidebar-nav-item')].some(
      i => i.classList.contains('active') && i.getAttribute('data-tab') !== 'home'
    );
    if (!hasActive) {
      document.querySelectorAll('.sidebar-nav-item').forEach(i => i.classList.remove('active'));
      const overviewItem = document.getElementById('navOverview');
      if (overviewItem) overviewItem.classList.add('active');
      updateTopBarTitle('📊 Executive Summary', 'KPIs, Charts & Insights');
    }
  }
};

// Override drilldown to use sidebar tab switching
window.drilldownToDoctor = function(doctorName) {
  if (!doctorName) return;

  const doctorCol = state.analytics?.doctorKPIs?.doctorColumnName || 'PRIMARY DOCTOR';
  state.selectedDimCol = doctorCol;
  state.quickCategoryFilter = doctorName;

  sessionStorage.setItem('mis_drilldown_dim', doctorCol);
  sessionStorage.setItem('mis_drilldown_val', doctorName);

  if (!window.location.pathname.endsWith('overview.html') && window.location.pathname !== '/') {
    window.location.href = 'overview.html';
    return;
  }

  const dimSelect = document.getElementById('activeDimensionSelect');
  if (dimSelect) dimSelect.value = doctorCol;

  showToast(`Drilling down entire dashboard for: ${doctorName}`);
  triggerAdvancedFilterRecompute();
};

