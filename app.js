/**
 * Fund Portfolio Analyzer - Web Application
 * Handles file upload, Excel parsing, calculations, and visualization
 */

// Global state
let portfolioData = null;
let privateFundData = null;
let analysisResults = null;
let chartInstances = {};
let publicFundDataCache = {};
let simulationNoticeShown = false;

// DOM Elements
const portfolioInput = document.getElementById('portfolioInput');
const privateInput = document.getElementById('privateInput');
const portfolioDropzone = document.getElementById('portfolioDropzone');
const privateDropzone = document.getElementById('privateDropzone');
const generateBtn = document.getElementById('generateBtn');
const statusText = document.getElementById('statusText');
const loadingSection = document.getElementById('loadingSection');
const resultsSection = document.getElementById('resultsSection');
const progressBar = document.getElementById('progressBar');
const loadingText = document.getElementById('loadingText');

// ============================================================================
// Utility Functions
// ============================================================================

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // Excel date serial number
    const epoch = new Date(1899, 11, 30);
    return new Date(epoch.getTime() + val * 24 * 60 * 60 * 1000);
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

function parseNumber(val) {
  if (val === undefined || val === null || val === '') return 0;
  if (typeof val === 'number') return val;
  const n = parseFloat(String(val).replace(/[,%]/g, ''));
  return isNaN(n) ? 0 : n;
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function formatMoney(val) {
  if (val === null || val === undefined) return 'N/A';
  if (val >= 10000) {
    return (val / 10000).toFixed(2) + '万';
  }
  return val.toFixed(2);
}

function formatPercent(val) {
  if (val === null || val === undefined) return 'N/A';
  return (val >= 0 ? '+' : '') + (val * 100).toFixed(2) + '%';
}

function getCategoryStyle(category) {
  const styles = {
    '私募': 'bg-sky-500/20 text-sky-400',
    '公募': 'bg-purple-500/20 text-purple-400',
    '专户': 'bg-emerald-500/20 text-emerald-400',
  };
  return styles[category] || 'bg-slate-500/20 text-slate-400';
}

/**
 * Show a toast notification
 */
function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  const bgColor = type === 'error' ? 'bg-red-500/90' : type === 'warning' ? 'bg-amber-500/90' : 'bg-sky-500/90';
  toast.className = `fixed top-4 right-4 ${bgColor} text-white px-4 py-3 rounded-lg shadow-lg z-50 text-sm font-medium transition-opacity duration-300`;
  toast.textContent = message;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

// ============================================================================
// File Upload Handlers
// ============================================================================

function setupUpload(input, dropzone, placeholder, fileInfo, fileName, fileSize, onUpload) {
  // Click to upload
  dropzone.addEventListener('click', () => input.click());

  // File selected
  input.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
      handleFile(e.target.files[0], placeholder, fileInfo, fileName, fileSize, onUpload);
    }
  });

  // Drag and drop
  dropzone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropzone.classList.add('dragover');
  });

  dropzone.addEventListener('dragleave', () => {
    dropzone.classList.remove('dragover');
  });

  dropzone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropzone.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0];
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        input.files = e.dataTransfer.files;
        handleFile(file, placeholder, fileInfo, fileName, fileSize, onUpload);
      } else {
        showToast('请上传 Excel 文件 (.xlsx 或 .xls)', 'error');
      }
    }
  });
}

function handleFile(file, placeholder, fileInfo, fileNameEl, fileSizeEl, onUpload) {
  try {
    placeholder.classList.add('hidden');
    fileInfo.classList.remove('hidden');
    fileNameEl.textContent = file.name;
    fileSizeEl.textContent = formatFileSize(file.size);

    const reader = new FileReader();
    reader.onerror = () => {
      showToast('文件读取失败，请重试', 'error');
      placeholder.classList.remove('hidden');
      fileInfo.classList.add('hidden');
    };
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        onUpload(workbook);
      } catch (err) {
        showToast('Excel 解析失败: ' + err.message, 'error');
        placeholder.classList.remove('hidden');
        fileInfo.classList.add('hidden');
      }
    };
    reader.readAsArrayBuffer(file);
  } catch (err) {
    showToast('文件处理出错: ' + err.message, 'error');
  }
}

// Setup uploads
setupUpload(
  portfolioInput, portfolioDropzone,
  document.getElementById('portfolioPlaceholder'),
  document.getElementById('portfolioFileInfo'),
  document.getElementById('portfolioFileName'),
  document.getElementById('portfolioFileSize'),
  (workbook) => {
    try {
      portfolioData = parsePortfolioWorkbook(workbook);
      if (!portfolioData || portfolioData.length === 0) {
        showToast('持仓文件中没有找到有效的基金数据', 'warning');
        portfolioData = null;
        return;
      }
      showToast(`成功读取 ${portfolioData.length} 只基金`);
      checkReady();
    } catch (err) {
      showToast('解析持仓文件失败: ' + err.message, 'error');
    }
  }
);

setupUpload(
  privateInput, privateDropzone,
  document.getElementById('privatePlaceholder'),
  document.getElementById('privateFileInfo'),
  document.getElementById('privateFileName'),
  document.getElementById('privateFileSize'),
  (workbook) => {
    try {
      privateFundData = parsePrivateFundWorkbook(workbook);
      const fundCount = Object.keys(privateFundData).length;
      if (fundCount === 0) {
        showToast('私募基金文件中没有找到业绩数据', 'warning');
      } else {
        showToast(`成功读取 ${fundCount} 只私募基金业绩数据`);
      }
      checkReady();
    } catch (err) {
      showToast('解析私募基金文件失败: ' + err.message, 'error');
    }
  }
);

function checkReady() {
  if (portfolioData && portfolioData.length > 0) {
    generateBtn.disabled = false;
    statusText.textContent = privateFundData
      ? '两个文件已上传，点击生成报告'
      : '持仓文件已上传，可点击生成报告（私募基金数据可选）';
    statusText.classList.add('text-emerald-400');
  }
}

// ============================================================================
// Workbook Parsers
// ============================================================================

function parsePortfolioWorkbook(workbook) {
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new Error('Excel 文件中没有工作表');
  }

  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  if (!jsonData || jsonData.length === 0) {
    throw new Error('工作表为空');
  }

  // Find header row
  let headerRow = 0;
  for (let i = 0; i < Math.min(10, jsonData.length); i++) {
    const row = jsonData[i];
    if (row && (row.includes('基金代码') || row.includes('产品名称') || row.includes('Fund Code'))) {
      headerRow = i;
      break;
    }
  }

  const headers = jsonData[headerRow];
  const funds = [];

  for (let i = headerRow + 1; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (!row || row.length < 3) continue;

    const fund = {};
    headers.forEach((h, idx) => {
      if (h) fund[h] = row[idx];
    });

    const code = String(fund['基金代码'] || fund['Fund Code'] || '').trim();
    const name = String(fund['产品名称'] || fund['Fund Name'] || '').trim();

    if (code && name) {
      funds.push({
        code: code,
        name: name,
        category: String(fund['产品分类'] || fund['Category'] || ''),
        purchaseDate: parseDate(fund['购买时间'] || fund['Purchase Date']),
        costPrice: parseNumber(fund['成本价'] || fund['Cost Price']),
        shares: parseNumber(fund['当前份额'] || fund['Shares'] || fund['份额']),
        nav: parseNumber(fund['T-1日净值'] || fund['Current NAV'] || fund['净值']),
        marketValue: parseNumber(fund['证券市值'] || fund['Market Value'] || fund['市值']),
      });
    }
  }

  return funds.filter(f => f.code && f.name);
}

function parsePrivateFundWorkbook(workbook) {
  const returns = {};

  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    if (!jsonData || jsonData.length === 0) return;

    // Step 1: Find code column (产品代码/基金代码) - search first 15 rows
    let codeCol = -1;
    for (let i = 0; i < Math.min(15, jsonData.length); i++) {
      const row = jsonData[i];
      if (!row) continue;
      row.forEach((cell, idx) => {
        if (cell === '产品代码' || cell === '基金代码' || cell === 'Code') {
          codeCol = idx;
        }
      });
      if (codeCol >= 0) break;
    }

    if (codeCol < 0) return; // No code column found

    // Step 2: Find year columns (support two-level headers)
    // Year headers like "2025", "2024" may be in a sub-header row
    let yearCols = {};
    let dataStartRow = -1;

    for (let i = 0; i < Math.min(20, jsonData.length); i++) {
      const row = jsonData[i];
      if (!row) continue;

      let foundYearsInRow = false;
      row.forEach((cell, idx) => {
          // Support both string ("2025") and number (2025) year formats
          const cellStr = String(cell).trim();
          if (/^20\d{2}$/.test(cellStr)) {
            yearCols[cellStr] = idx;
            foundYearsInRow = true;
          }
        });

      if (foundYearsInRow) {
        // The data starts after this header row
        dataStartRow = i + 1;
      }
    }

    if (dataStartRow < 0 || Object.keys(yearCols).length === 0) return;

    // Step 3: Parse data rows
    for (let i = dataStartRow; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || !row[codeCol]) continue;

      const code = String(row[codeCol]).trim();
      if (!code) continue;

      // Initialize if not exists
      if (!returns[code]) {
        returns[code] = {};
      }

      Object.entries(yearCols).forEach(([year, col]) => {
        const val = row[col];
        if (val !== undefined && val !== null && val !== '') {
          // The return values in the Excel are already in decimal format (e.g., 0.3722 = 37.22%)
          // No need to divide by 100 again
          returns[code][year] = parseNumber(val);
        }
      });
    }
  });

  return returns;
}

// ============================================================================
// Public Fund Data Fetching (Simulated)
// ============================================================================

/**
 * Generate a deterministic pseudo-random number from a string seed
 */
function seededRandom(seed) {
  let hash = 0;
  for (let i = 0; i < seed.length; i++) {
    hash = ((hash << 5) - hash) + seed.charCodeAt(i);
    hash = hash & hash;
  }
  const x = Math.sin(hash) * 10000;
  return x - Math.floor(x);
}

/**
 * Generate simulated yearly returns for a public fund based on its code and name.
 * The data is pseudo-random but deterministic (same code+name always gives same data).
 */
function generateSimulatedReturns(fundCode, fundName) {
  const years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  const returns = {};

  // Use fund code and name as seed for deterministic results
  const baseSeed = fundCode + fundName;

  // Determine fund style based on name keywords
  const nameLower = fundName.toLowerCase();
  let style = 'balanced'; // balanced, aggressive, conservative
  if (nameLower.includes('债') || nameLower.includes('货币') || nameLower.includes('理财')) {
    style = 'conservative';
  } else if (nameLower.includes('指数') || nameLower.includes('股票') || nameLower.includes('混合')) {
    style = 'aggressive';
  }

  years.forEach((year, idx) => {
    const seed = baseSeed + year;
    const rand = seededRandom(seed);

    let baseReturn;
    let volatility;

    // Market context for each year
    const marketContext = {
      '2018': { base: -0.15, vol: 0.12 },
      '2019': { base: 0.25, vol: 0.10 },
      '2020': { base: 0.20, vol: 0.14 },
      '2021': { base: 0.05, vol: 0.11 },
      '2022': { base: -0.12, vol: 0.13 },
      '2023': { base: -0.05, vol: 0.10 },
      '2024': { base: 0.10, vol: 0.09 },
      '2025': { base: 0.08, vol: 0.10 },
    };

    const ctx = marketContext[year] || { base: 0.05, vol: 0.10 };

    if (style === 'conservative') {
      baseReturn = ctx.base * 0.3 + 0.03;
      volatility = ctx.vol * 0.4;
    } else if (style === 'aggressive') {
      baseReturn = ctx.base * 1.2;
      volatility = ctx.vol * 1.1;
    } else {
      baseReturn = ctx.base * 0.7 + 0.02;
      volatility = ctx.vol * 0.8;
    }

    // Generate return with some randomness
    const ret = baseReturn + (rand - 0.5) * volatility * 2;
    returns[year] = Math.max(-0.35, Math.min(0.80, ret));
  });

  return returns;
}

/**
 * Fetch public fund data. Since browsers cannot make cross-origin requests to fund websites,
 * this function first checks privateFundData, then falls back to simulated data.
 */
function fetchPublicFundData(fundCode, fundName) {
  // Validate inputs
  if (!fundCode || !fundName) {
    console.warn('fetchPublicFundData: missing fundCode or fundName');
    return null;
  }

  // Check cache first
  if (publicFundDataCache[fundCode]) {
    return publicFundDataCache[fundCode];
  }

  // Check if data exists in privateFundData
  if (privateFundData && privateFundData[fundCode]) {
    publicFundDataCache[fundCode] = { ...privateFundData[fundCode] };
    return publicFundDataCache[fundCode];
  }

  // Generate simulated data
  const simulated = generateSimulatedReturns(fundCode, fundName);
  publicFundDataCache[fundCode] = simulated;

  // Show notice once
  if (!simulationNoticeShown) {
    showToast('公募基金数据为模拟数据，仅供参考', 'warning');
    simulationNoticeShown = true;
  }

  return simulated;
}

/**
 * Check if a fund is a public fund (公募)
 */
function isPublicFund(fund) {
  const category = (fund.category || '').trim();
  return category === '公募' || category.includes('公募');
}

// ============================================================================
// Report Generation
// ============================================================================

generateBtn.addEventListener('click', async () => {
  if (!portfolioData || portfolioData.length === 0) {
    showToast('请先上传基金持仓文件', 'error');
    return;
  }

  // Reset simulation notice for new analysis
  simulationNoticeShown = false;
  publicFundDataCache = {};

  // Show loading
  loadingSection.classList.remove('hidden');
  resultsSection.classList.add('hidden');
  generateBtn.disabled = true;

  // Dispose old chart instances
  disposeAllCharts();

  // Progress steps
  const steps = [
    { text: '读取基金持仓文件', progress: 10 },
    { text: '解析私募基金业绩数据', progress: 25 },
    { text: '查询公募基金历史数据', progress: 40 },
    { text: '计算各基金年化收益率', progress: 55 },
    { text: '计算组合加权收益', progress: 70 },
    { text: '生成可视化图表', progress: 90 },
    { text: '报告生成完成', progress: 100 },
  ];

  try {
    for (const step of steps) {
      loadingText.textContent = step.text;
      progressBar.style.width = step.progress + '%';
      await sleep(300);
    }

    // Perform calculations
    analysisResults = calculateAnalysis();

    // Render results
    renderResults();

    // Show results
    loadingSection.classList.add('hidden');
    resultsSection.classList.remove('hidden');
    generateBtn.disabled = false;

    // Scroll to results
    resultsSection.scrollIntoView({ behavior: 'smooth' });

    // Initialize charts after container becomes visible
    // Use requestAnimationFrame for reliable timing
    requestAnimationFrame(() => {
      setTimeout(() => {
        initAllCharts();
      }, 50);
    });
  } catch (err) {
    loadingSection.classList.add('hidden');
    generateBtn.disabled = false;
    showToast('报告生成失败: ' + err.message, 'error');
    console.error('Report generation error:', err);
  }
});

// ============================================================================
// Chart Management
// ============================================================================

function disposeAllCharts() {
  Object.values(chartInstances).forEach(chart => {
    if (chart && !chart.isDisposed()) {
      chart.dispose();
    }
  });
  chartInstances = {};
}

function initAllCharts() {
  if (!analysisResults) return;

  try {
    renderYearlyChart(analysisResults.weightedReturns);
  } catch (e) {
    console.error('Yearly chart error:', e);
  }

  try {
    renderCategoryChart(analysisResults.categories);
  } catch (e) {
    console.error('Category chart error:', e);
  }

  try {
    renderAnnualizedChart(analysisResults.funds);
  } catch (e) {
    console.error('Annualized chart error:', e);
  }

  try {
    renderCumulativeChart(analysisResults.funds, analysisResults.weightedReturns);
  } catch (e) {
    console.error('Cumulative chart error:', e);
  }

  try {
    renderHeatmap(analysisResults.funds);
  } catch (e) {
    console.error('Heatmap error:', e);
  }
}

function getChart(domId) {
  const dom = document.getElementById(domId);
  if (!dom) return null;

  // Check if already initialized
  if (chartInstances[domId] && !chartInstances[domId].isDisposed()) {
    return chartInstances[domId];
  }

  // Create new instance
  const chart = echarts.init(dom);
  chartInstances[domId] = chart;
  return chart;
}

// ============================================================================
// Analysis Calculation
// ============================================================================

function calculateAnalysis() {
  const years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];

  const funds = portfolioData.map(fund => {
    // Validate required fields
    if (!fund.code || !fund.name) {
      console.warn('Skipping fund with missing code or name');
      return null;
    }

    // Calculate total return
    const totalReturn = fund.costPrice > 0 ? (fund.nav - fund.costPrice) / fund.costPrice : 0;

    // Calculate annualized return
    let annualizedReturn = null;
    if (fund.purchaseDate && fund.costPrice > 0) {
      const days = (new Date('2026-06-12') - fund.purchaseDate) / (1000 * 60 * 60 * 24);
      const years = days / 365.25;
      if (years > 0) {
        annualizedReturn = Math.pow(1 + totalReturn, 1 / years) - 1;
      }
    }

    // Get yearly returns
    const yearlyReturns = {};

    // For public funds, try to fetch data
    if (isPublicFund(fund)) {
      const publicData = fetchPublicFundData(fund.code, fund.name);
      if (publicData) {
        Object.entries(publicData).forEach(([year, ret]) => {
          yearlyReturns[year] = ret;
        });
      }
    }

    // Override with private fund data if available (takes precedence)
    if (privateFundData && privateFundData[fund.code]) {
      Object.entries(privateFundData[fund.code]).forEach(([year, ret]) => {
        yearlyReturns[year] = ret;
      });
    }

    return {
      ...fund,
      totalReturn,
      annualizedReturn,
      yearlyReturns,
    };
  }).filter(f => f !== null);

  // Calculate portfolio totals
  const totalCost = funds.reduce((sum, f) => sum + (f.costPrice * f.shares), 0);
  const totalValue = funds.reduce((sum, f) => sum + f.marketValue, 0);
  const portfolioReturn = totalCost > 0 ? (totalValue - totalCost) / totalCost : 0;

  // Calculate weighted yearly returns
  const weightedReturns = {};

  years.forEach(year => {
    let totalWeight = 0;
    let weightedSum = 0;

    funds.forEach(fund => {
      const ret = fund.yearlyReturns[year];
      if (ret !== undefined && ret !== null && fund.marketValue > 0) {
        weightedSum += ret * fund.marketValue;
        totalWeight += fund.marketValue;
      }
    });

    weightedReturns[year] = totalWeight > 0 ? weightedSum / totalWeight : null;
  });

  // Category breakdown
  const categories = {};
  funds.forEach(fund => {
    const cat = fund.category || '其他';
    if (!categories[cat]) {
      categories[cat] = { count: 0, value: 0 };
    }
    categories[cat].count++;
    categories[cat].value += fund.marketValue;
  });

  return {
    funds,
    totalCost,
    totalValue,
    portfolioReturn,
    weightedReturns,
    categories,
  };
}

// ============================================================================
// Results Rendering
// ============================================================================

function renderResults() {
  const r = analysisResults;

  // Report date
  const reportDateEl = document.getElementById('reportDate');
  if (reportDateEl) {
    reportDateEl.textContent = new Date().toLocaleDateString('zh-CN');
  }

  // KPI Cards
  document.getElementById('kpiCost').textContent = formatMoney(r.totalCost);
  document.getElementById('kpiValue').textContent = formatMoney(r.totalValue);

  const kpiReturnEl = document.getElementById('kpiReturn');
  kpiReturnEl.textContent = formatPercent(r.portfolioReturn);
  kpiReturnEl.className = 'text-3xl font-bold tracking-tight ' + (r.portfolioReturn >= 0 ? 'text-emerald-400' : 'text-red-400');

  const kpiReturnArrow = document.getElementById('kpiReturnArrow');
  if (kpiReturnArrow) {
    kpiReturnArrow.innerHTML = r.portfolioReturn >= 0
      ? '<svg class="w-5 h-5 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"/></svg>'
      : '<svg class="w-5 h-5 text-red-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 17h8m0 0V9m0 8l-8-8-4 4-6-6"/></svg>';
  }

  const ret2025 = r.weightedReturns['2025'];
  const kpi2025El = document.getElementById('kpi2025');
  kpi2025El.textContent = ret2025 !== null ? formatPercent(ret2025) : 'N/A';
  kpi2025El.className = 'text-3xl font-bold tracking-tight ' + (ret2025 >= 0 ? 'text-emerald-400' : 'text-red-400');

  const kpi2025Arrow = document.getElementById('kpi2025Arrow');
  if (kpi2025Arrow && ret2025 !== null) {
    kpi2025Arrow.innerHTML = ret2025 >= 0
      ? '<svg class="w-5 h-5 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 7h8m0 0v8m0-8l-8 8-4-4-6 6"/></svg>'
      : '<svg class="w-5 h-5 text-red-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 17h8m0 0V9m0 8l-8-8-4 4-6-6"/></svg>';
  }

  // Investment Overview
  renderInvestmentOverview(r.funds);

  // Table
  renderTable(r.funds);

  // Table info
  const tableInfoEl = document.getElementById('tableInfo');
  if (tableInfoEl) {
    tableInfoEl.textContent = `共 ${r.funds.length} 只基金`;
  }
}

function renderInvestmentOverview(funds) {
  const overviewFundCount = document.getElementById('overviewFundCount');
  const overviewFirstPurchase = document.getElementById('overviewFirstPurchase');
  const overviewMaxHolding = document.getElementById('overviewMaxHolding');
  const overviewAvgHolding = document.getElementById('overviewAvgHolding');

  if (overviewFundCount) overviewFundCount.textContent = funds.length;

  const fundsWithDate = funds.filter(f => f.purchaseDate);
  if (fundsWithDate.length > 0) {
    const dates = fundsWithDate.map(f => f.purchaseDate);
    const earliest = new Date(Math.min(...dates));
    const now = new Date('2026-06-12');

    if (overviewFirstPurchase) {
      overviewFirstPurchase.textContent = earliest.toLocaleDateString('zh-CN', { year: 'numeric', month: '2-digit' });
    }

    const holdingPeriods = fundsWithDate.map(f => (now - f.purchaseDate) / (1000 * 60 * 60 * 24));
    const maxDays = Math.max(...holdingPeriods);
    const avgDays = holdingPeriods.reduce((a, b) => a + b, 0) / holdingPeriods.length;

    if (overviewMaxHolding) {
      const maxYears = (maxDays / 365.25).toFixed(1);
      overviewMaxHolding.textContent = maxYears + '年';
    }
    if (overviewAvgHolding) {
      const avgYears = (avgDays / 365.25).toFixed(1);
      overviewAvgHolding.textContent = avgYears + '年';
    }
  } else {
    if (overviewFirstPurchase) overviewFirstPurchase.textContent = '-';
    if (overviewMaxHolding) overviewMaxHolding.textContent = '-';
    if (overviewAvgHolding) overviewAvgHolding.textContent = '-';
  }
}

// Table state for sorting and pagination
let tableState = {
  funds: [],
  sortKey: null,
  sortDesc: true,
  page: 1,
  pageSize: 10
};

function renderTable(funds) {
  tableState.funds = funds;
  tableState.page = 1;
  renderTablePage();
  setupTableSorting();
  setupTablePagination();
}

function renderTablePage() {
  const tbody = document.getElementById('fundTableBody');
  tbody.innerHTML = '';

  let displayFunds = [...tableState.funds];

  // Sort
  if (tableState.sortKey) {
    const key = tableState.sortKey;
    displayFunds.sort((a, b) => {
      let av = a[key];
      let bv = b[key];
      if (av === null || av === undefined) av = -Infinity;
      if (bv === null || bv === undefined) bv = -Infinity;
      if (typeof av === 'string') av = av.toLowerCase();
      if (typeof bv === 'string') bv = bv.toLowerCase();
      if (av < bv) return tableState.sortDesc ? 1 : -1;
      if (av > bv) return tableState.sortDesc ? -1 : 1;
      return 0;
    });
  }

  // Pagination
  const totalPages = Math.max(1, Math.ceil(displayFunds.length / tableState.pageSize));
  if (tableState.page > totalPages) tableState.page = totalPages;
  const start = (tableState.page - 1) * tableState.pageSize;
  const end = start + tableState.pageSize;
  const pageFunds = displayFunds.slice(start, end);

  const years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];

  pageFunds.forEach(fund => {
    const tr = document.createElement('tr');
    tr.className = 'border-b border-slate-800/50 table-row-hover transition-colors';

    const totalReturn = fund.totalReturn;
    const annReturn = fund.annualizedReturn;

    let yearlyCells = '';
    years.forEach(y => {
      const ret = fund.yearlyReturns[y];
      if (ret !== undefined && ret !== null) {
        const c = ret >= 0 ? 'text-emerald-400' : 'text-red-400';
        yearlyCells += `<td class="px-4 py-3 text-right ${c} text-xs">${(ret * 100).toFixed(2)}%</td>`;
      } else {
        yearlyCells += `<td class="px-4 py-3 text-right text-slate-600 text-xs">-</td>`;
      }
    });

    tr.innerHTML = `
      <td class="px-4 py-3 text-slate-200 font-medium">${fund.name}</td>
      <td class="px-4 py-3">
        <span class="px-2 py-0.5 rounded text-xs ${getCategoryStyle(fund.category)}">${fund.category || '其他'}</span>
      </td>
      <td class="px-4 py-3 text-right text-slate-300">${fund.costPrice.toFixed(2)}</td>
      <td class="px-4 py-3 text-right text-slate-300">${fund.nav.toFixed(3)}</td>
      <td class="px-4 py-3 text-right text-slate-300">${formatMoney(fund.marketValue)}</td>
      <td class="px-4 py-3 text-right ${totalReturn >= 0 ? 'text-emerald-400' : 'text-red-400'} font-medium">${formatPercent(totalReturn)}</td>
      <td class="px-4 py-3 text-right ${annReturn !== null ? (annReturn >= 0 ? 'text-emerald-400' : 'text-red-400') : 'text-slate-500'} font-medium">${annReturn !== null ? formatPercent(annReturn) : 'N/A'}</td>
      ${yearlyCells}
    `;

    tbody.appendChild(tr);
  });

  // Update pagination
  updatePagination(displayFunds.length, totalPages);
}

function updatePagination(totalItems, totalPages) {
  const paginationEl = document.getElementById('tablePagination');
  const infoEl = document.getElementById('paginationInfo');
  const prevBtn = document.getElementById('pagePrev');
  const nextBtn = document.getElementById('pageNext');
  const pageNumbers = document.getElementById('pageNumbers');

  if (!paginationEl) return;

  if (totalItems <= tableState.pageSize) {
    paginationEl.classList.add('hidden');
    return;
  }

  paginationEl.classList.remove('hidden');
  if (infoEl) {
    infoEl.textContent = `共 ${totalItems} 只基金，第 ${tableState.page} / ${totalPages} 页`;
  }
  if (prevBtn) prevBtn.disabled = tableState.page <= 1;
  if (nextBtn) nextBtn.disabled = tableState.page >= totalPages;

  if (pageNumbers) {
    let html = '';
    for (let i = 1; i <= totalPages; i++) {
      if (i === tableState.page) {
        html += `<span class="px-2.5 py-1 rounded-lg bg-sky-500/20 text-sky-400 text-xs font-medium">${i}</span>`;
      } else {
        html += `<button class="page-num px-2.5 py-1 rounded-lg bg-slate-800 text-slate-400 text-xs hover:bg-slate-700 transition" data-page="${i}">${i}</button>`;
      }
    }
    pageNumbers.innerHTML = html;

    pageNumbers.querySelectorAll('.page-num').forEach(btn => {
      btn.addEventListener('click', () => {
        tableState.page = parseInt(btn.dataset.page);
        renderTablePage();
      });
    });
  }
}

function setupTablePagination() {
  const prevBtn = document.getElementById('pagePrev');
  const nextBtn = document.getElementById('pageNext');

  if (prevBtn) {
    prevBtn.onclick = () => {
      if (tableState.page > 1) {
        tableState.page--;
        renderTablePage();
      }
    };
  }
  if (nextBtn) {
    nextBtn.onclick = () => {
      const totalPages = Math.ceil(tableState.funds.length / tableState.pageSize);
      if (tableState.page < totalPages) {
        tableState.page++;
        renderTablePage();
      }
    };
  }
}

function setupTableSorting() {
  document.querySelectorAll('.sortable-header').forEach(th => {
    th.addEventListener('click', () => {
      const key = th.dataset.sort;
      if (!key) return;

      if (tableState.sortKey === key) {
        tableState.sortDesc = !tableState.sortDesc;
      } else {
        tableState.sortKey = key;
        tableState.sortDesc = true;
      }

      // Update sort icons
      document.querySelectorAll('.sortable-header .sort-icon').forEach(icon => {
        icon.style.opacity = '0.3';
        icon.textContent = '\u21D5';
      });
      const activeIcon = th.querySelector('.sort-icon');
      if (activeIcon) {
        activeIcon.style.opacity = '1';
        activeIcon.textContent = tableState.sortDesc ? '\u21D3' : '\u21D1';
      }

      renderTablePage();
    });
  });
}

// ============================================================================
// Chart Rendering
// ============================================================================

function renderYearlyChart(weightedReturns) {
  const chart = getChart('yearlyChart');
  if (!chart) return;

  const years = Object.keys(weightedReturns);
  const values = Object.values(weightedReturns);

  chart.setOption({
    animation: false,
    tooltip: {
      trigger: 'axis',
      formatter: (params) => {
        const p = params[0];
        const val = p.value;
        return p.name + '年: ' + (val !== null ? (val >= 0 ? '+' : '') + (val * 100).toFixed(2) + '%' : 'N/A');
      }
    },
    grid: { left: '3%', right: '4%', bottom: '3%', top: '10%', containLabel: true },
    xAxis: {
      type: 'category',
      data: years,
      axisLine: { lineStyle: { color: '#334155' } },
      axisLabel: { color: '#94a3b8' }
    },
    yAxis: {
      type: 'value',
      axisLine: { lineStyle: { color: '#334155' } },
      axisLabel: {
        color: '#94a3b8',
        formatter: (v) => (v * 100).toFixed(0) + '%'
      },
      splitLine: { lineStyle: { color: '#1e293b', type: 'dashed' } }
    },
    series: [{
      type: 'bar',
      data: values.map(v => v !== null ? v : 0),
      barWidth: '50%',
      itemStyle: {
        color: (params) => {
          const val = values[params.dataIndex];
          return val >= 0 ? '#34d399' : '#f87171';
        },
        borderRadius: [4, 4, 0, 0]
      },
      label: {
        show: true,
        position: 'top',
        color: '#e2e8f0',
        formatter: (p) => {
          const val = values[p.dataIndex];
          return val !== null ? (val * 100).toFixed(1) + '%' : 'N/A';
        }
      }
    }]
  });
}

function renderCategoryChart(categories) {
  const chart = getChart('categoryChart');
  if (!chart) return;

  const data = Object.entries(categories).map(([name, info]) => ({
    name,
    value: info.value
  }));

  const colors = ['#38bdf8', '#a78bfa', '#34d399', '#fbbf24'];

  chart.setOption({
    animation: false,
    tooltip: {
      trigger: 'item',
      formatter: (p) => p.name + ': ' + formatMoney(p.value) + ' (' + p.percent.toFixed(1) + '%)'
    },
    legend: {
      orient: 'vertical',
      right: '5%',
      top: 'center',
      textStyle: { color: '#e2e8f0' }
    },
    series: [{
      type: 'pie',
      radius: ['40%', '70%'],
      center: ['40%', '50%'],
      avoidLabelOverlap: true,
      itemStyle: { borderRadius: 6, borderColor: '#0f172a', borderWidth: 2 },
      label: { show: false },
      emphasis: {
        label: { show: true, fontSize: 14, fontWeight: 'bold', color: '#e2e8f0' }
      },
      data: data.map((d, i) => ({
        ...d,
        itemStyle: { color: colors[i % colors.length] }
      }))
    }]
  });
}

function renderAnnualizedChart(funds) {
  const chart = getChart('annualizedChart');
  if (!chart) return;

  const sortedFunds = [...funds].sort((a, b) => {
    const av = a.annualizedReturn !== null ? a.annualizedReturn : -999;
    const bv = b.annualizedReturn !== null ? b.annualizedReturn : -999;
    return bv - av;
  });

  chart.setOption({
    animation: false,
    tooltip: {
      trigger: 'axis',
      formatter: (params) => {
        const p = params[0];
        const fund = sortedFunds[p.dataIndex];
        return fund.name + '<br/>年化收益: ' + (fund.annualizedReturn !== null ? (fund.annualizedReturn * 100).toFixed(1) + '%' : 'N/A');
      }
    },
    grid: { left: '3%', right: '8%', bottom: '3%', top: '3%', containLabel: true },
    xAxis: {
      type: 'value',
      axisLine: { lineStyle: { color: '#334155' } },
      axisLabel: {
        color: '#94a3b8',
        formatter: (v) => (v * 100).toFixed(0) + '%'
      },
      splitLine: { lineStyle: { color: '#1e293b', type: 'dashed' } }
    },
    yAxis: {
      type: 'category',
      data: sortedFunds.map(f => f.name.length > 10 ? f.name.substring(0, 10) + '...' : f.name),
      axisLine: { lineStyle: { color: '#334155' } },
      axisLabel: { color: '#e2e8f0', fontSize: 11 }
    },
    series: [{
      type: 'bar',
      data: sortedFunds.map(f => f.annualizedReturn !== null ? f.annualizedReturn : 0),
      barWidth: '60%',
      itemStyle: {
        color: (params) => {
          const fund = sortedFunds[params.dataIndex];
          if (fund.annualizedReturn === null) return '#64748b';
          return fund.annualizedReturn >= 0 ? '#34d399' : '#f87171';
        },
        borderRadius: [0, 4, 4, 0]
      },
      label: {
        show: true,
        position: 'right',
        color: '#e2e8f0',
        formatter: (p) => {
          const fund = sortedFunds[p.dataIndex];
          return fund.annualizedReturn !== null ? (fund.annualizedReturn * 100).toFixed(1) + '%' : 'N/A';
        }
      }
    }]
  });
}

function renderCumulativeChart(funds, weightedReturns) {
  const chart = getChart('cumulativeChart');
  if (!chart) return;

  const years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];

  // Calculate cumulative portfolio return
  let cumulative = 1;
  const cumulativeData = years.map(year => {
    const ret = weightedReturns[year];
    if (ret !== null && ret !== undefined) {
      cumulative = cumulative * (1 + ret);
    }
    return (cumulative - 1) * 100;
  });

  // Calculate benchmark (假设年化5%)
  let benchmark = 1;
  const benchmarkData = years.map(() => {
    benchmark = benchmark * 1.05;
    return (benchmark - 1) * 100;
  });

  chart.setOption({
    animation: true,
    tooltip: {
      trigger: 'axis',
      backgroundColor: 'rgba(15, 23, 42, 0.95)',
      borderColor: 'rgba(56, 189, 248, 0.2)',
      textStyle: { color: '#e2e8f0' },
      formatter: (params) => {
        let result = '<strong>' + params[0].name + '年</strong><br/>';
        params.forEach(p => {
          const color = p.color;
          result += '<span style="display:inline-block;width:10px;height:10px;border-radius:50%;background:' + color + ';margin-right:6px;"></span>';
          result += p.seriesName + ': ' + (p.value >= 0 ? '+' : '') + p.value.toFixed(2) + '%<br/>';
        });
        return result;
      }
    },
    legend: {
      data: ['组合累计收益', '基准收益(5%/年)'],
      textStyle: { color: '#94a3b8' },
      top: 0
    },
    grid: { left: '3%', right: '4%', bottom: '3%', top: '12%', containLabel: true },
    xAxis: {
      type: 'category',
      data: years,
      axisLine: { lineStyle: { color: '#334155' } },
      axisLabel: { color: '#94a3b8' }
    },
    yAxis: {
      type: 'value',
      axisLine: { lineStyle: { color: '#334155' } },
      axisLabel: {
        color: '#94a3b8',
        formatter: (v) => v.toFixed(0) + '%'
      },
      splitLine: { lineStyle: { color: '#1e293b', type: 'dashed' } }
    },
    series: [
      {
        name: '组合累计收益',
        type: 'line',
        data: cumulativeData,
        smooth: true,
        symbol: 'circle',
        symbolSize: 8,
        lineStyle: { width: 3, color: '#38bdf8' },
        itemStyle: { color: '#38bdf8', borderWidth: 2, borderColor: '#0f172a' },
        areaStyle: {
          color: {
            type: 'linear',
            x: 0, y: 0, x2: 0, y2: 1,
            colorStops: [
              { offset: 0, color: 'rgba(56, 189, 248, 0.3)' },
              { offset: 1, color: 'rgba(56, 189, 248, 0.02)' }
            ]
          }
        },
        label: {
          show: true,
          position: 'top',
          color: '#e2e8f0',
          formatter: (p) => (p.value >= 0 ? '+' : '') + p.value.toFixed(1) + '%'
        }
      },
      {
        name: '基准收益(5%/年)',
        type: 'line',
        data: benchmarkData,
        smooth: true,
        symbol: 'none',
        lineStyle: { width: 2, color: '#64748b', type: 'dashed' },
        itemStyle: { color: '#64748b' }
      }
    ]
  });
}

function renderHeatmap(funds) {
  const chart = getChart('heatmapChart');
  if (!chart) return;

  const years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];

  // Build full data matrix including N/A cells
  const data = [];
  funds.forEach((fund, fundIdx) => {
    years.forEach((year, yearIdx) => {
      const ret = fund.yearlyReturns[year];
      if (ret !== undefined && ret !== null) {
        data.push([yearIdx, fundIdx, ret]);
      } else {
        data.push([yearIdx, fundIdx, '-']);
      }
    });
  });

  // Dynamic height based on fund count
  const container = document.getElementById('heatmapChart');
  const dynamicHeight = Math.max(400, funds.length * 36 + 140);
  container.style.height = dynamicHeight + 'px';
  chart.resize();

  chart.setOption({
    animation: false,
    tooltip: {
      position: 'top',
      formatter: (params) => {
        const fund = funds[params.value[1]];
        const year = years[params.value[0]];
        const val = params.value[2];
        if (val === '-') return fund.name + ' (' + year + '): N/A';
        return fund.name + ' (' + year + '): ' + (val >= 0 ? '+' : '') + (val * 100).toFixed(2) + '%';
      }
    },
    grid: { left: '25%', right: '5%', bottom: '10%', top: '5%' },
    xAxis: {
      type: 'category',
      data: years,
      axisLine: { lineStyle: { color: '#334155' } },
      axisLabel: { color: '#94a3b8' },
      splitArea: { show: false }
    },
    yAxis: {
      type: 'category',
      data: funds.map(f => f.name.length > 14 ? f.name.substring(0, 14) + '...' : f.name),
      axisLine: { lineStyle: { color: '#334155' } },
      axisLabel: { color: '#e2e8f0', fontSize: 11 },
      splitArea: { show: false }
    },
    visualMap: {
      min: -0.35,
      max: 0.8,
      calculable: true,
      orient: 'horizontal',
      left: 'center',
      bottom: '0%',
      textStyle: { color: '#94a3b8' },
      inRange: {
        color: ['#f87171', '#1e293b', '#34d399']
      },
      outOfRange: { color: 'transparent' }
    },
    series: [{
      type: 'heatmap',
      data: data,
      label: {
        show: true,
        formatter: (p) => {
          if (p.value[2] === '-') return '';
          return (p.value[2] * 100).toFixed(0) + '%';
        },
        fontSize: 10,
        color: '#e2e8f0'
      },
      emphasis: {
        itemStyle: { borderColor: '#38bdf8', borderWidth: 2 }
      }
    }]
  });
}

// ============================================================================
// Export Functions
// ============================================================================

function exportToExcel() {
  if (!analysisResults) {
    showToast('请先生成分析报告', 'warning');
    return;
  }

  try {
    const wb = XLSX.utils.book_new();
    const years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];

    // Sheet 1: Fund Details with yearly returns
    const fundData = analysisResults.funds.map(f => {
      const row = {
        '基金代码': f.code,
        '产品名称': f.name,
        '产品分类': f.category,
        '购买时间': f.purchaseDate ? f.purchaseDate.toISOString().split('T')[0] : '',
        '成本价': f.costPrice,
        '当前份额': f.shares,
        '当前净值': f.nav,
        '证券市值': f.marketValue,
        '总收益率': f.totalReturn,
        '持有期年化收益率': f.annualizedReturn,
      };
      // Add yearly returns
      years.forEach(y => {
        const ret = f.yearlyReturns[y];
        row[y + '年收益'] = (ret !== undefined && ret !== null) ? ret : '';
      });
      return row;
    });

    const ws1 = XLSX.utils.json_to_sheet(fundData);
    XLSX.utils.book_append_sheet(wb, ws1, '基金年度收益明细');

    // Sheet 2: Per-fund yearly return since purchase
    const yearlyChangeData = [];
    analysisResults.funds.forEach(f => {
      if (!f.purchaseDate) return;
      const purchaseYear = f.purchaseDate.getFullYear();

      // For each year from purchase year to 2025, calculate cumulative return
      let cumulativeReturn = 0;
      const row = {
        '基金代码': f.code,
        '产品名称': f.name,
        '购买时间': f.purchaseDate.toISOString().split('T')[0],
      };

      for (let y = purchaseYear; y <= 2025; y++) {
        const yearRet = f.yearlyReturns[String(y)];
        if (yearRet !== undefined && yearRet !== null) {
          cumulativeReturn = (1 + cumulativeReturn) * (1 + yearRet) - 1;
        }
        const yearsHeld = y - purchaseYear + (y === purchaseYear ? 0.5 : 1);
        const annualized = yearsHeld > 0 ? (Math.pow(1 + cumulativeReturn, 1 / yearsHeld) - 1) : 0;

        row[y + '年累计收益'] = cumulativeReturn;
        row[y + '年持有年化'] = annualized;
      }

      yearlyChangeData.push(row);
    });

    const ws2 = XLSX.utils.json_to_sheet(yearlyChangeData);
    XLSX.utils.book_append_sheet(wb, ws2, '购买后逐年收益变化');

    // Sheet 3: Portfolio Summary
    const summaryData = [{
      '组合总成本': analysisResults.totalCost,
      '组合总市值': analysisResults.totalValue,
      '组合总收益率': analysisResults.portfolioReturn,
    }];

    Object.entries(analysisResults.weightedReturns).forEach(([year, ret]) => {
      summaryData[0][year + '年加权收益'] = ret !== null ? ret : 'N/A';
    });

    const ws3 = XLSX.utils.json_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, ws3, '组合概览');

    XLSX.writeFile(wb, '基金组合分析.xlsx');
    showToast('Excel 报告已导出');
  } catch (err) {
    showToast('导出 Excel 失败: ' + err.message, 'error');
    console.error('Excel export error:', err);
  }
}

function exportToHtml() {
  if (!analysisResults) {
    showToast('请先生成分析报告', 'warning');
    return;
  }

  try {
    // Build a standalone HTML report
    const r = analysisResults;
    const years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];

    let yearlyRows = '';
    years.forEach(y => {
      const ret = r.weightedReturns[y];
      const color = ret >= 0 ? '#34d399' : '#f87171';
      const display = ret !== null ? (ret >= 0 ? '+' : '') + (ret * 100).toFixed(2) + '%' : 'N/A';
      yearlyRows += `<tr><td style="padding:8px 12px;color:#94a3b8">${y}</td><td style="padding:8px 12px;text-align:right;color:${color};font-weight:600">${display}</td></tr>`;
    });

    let fundRows = '';
    r.funds.forEach(f => {
      const trClass = f.totalReturn >= 0 ? 'color:#34d399' : 'color:#f87171';
      const annClass = f.annualizedReturn !== null ? (f.annualizedReturn >= 0 ? 'color:#34d399' : 'color:#f87171') : 'color:#64748b';
      const annDisplay = f.annualizedReturn !== null ? (f.annualizedReturn >= 0 ? '+' : '') + (f.annualizedReturn * 100).toFixed(2) + '%' : 'N/A';

      let yearlyCells = '';
      years.forEach(y => {
        const ret = f.yearlyReturns[y];
        if (ret !== undefined && ret !== null) {
          const c = ret >= 0 ? '#34d399' : '#f87171';
          yearlyCells += `<td style="padding:6px 8px;text-align:right;color:${c}">${(ret * 100).toFixed(2)}%</td>`;
        } else {
          yearlyCells += `<td style="padding:6px 8px;text-align:right;color:#475569">-</td>`;
        }
      });

      fundRows += `
        <tr style="border-bottom:1px solid #1e293b">
          <td style="padding:8px 12px;color:#e2e8f0">${f.name}</td>
          <td style="padding:8px 12px">${f.category || '其他'}</td>
          <td style="padding:8px 12px;text-align:right;color:#94a3b8">${f.purchaseDate ? f.purchaseDate.toISOString().split('T')[0] : '-'}</td>
          <td style="padding:8px 12px;text-align:right">${f.costPrice.toFixed(2)}</td>
          <td style="padding:8px 12px;text-align:right">${f.nav.toFixed(3)}</td>
          <td style="padding:8px 12px;text-align:right">${formatMoney(f.marketValue)}</td>
          <td style="padding:8px 12px;text-align:right;${trClass};font-weight:600">${(f.totalReturn >= 0 ? '+' : '') + (f.totalReturn * 100).toFixed(2)}%</td>
          <td style="padding:8px 12px;text-align:right;${annClass}">${annDisplay}</td>
          ${yearlyCells}
        </tr>`;
    });

    const html = `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>基金投资组合分析报告</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#0a0e1a;color:#e2e8f0;line-height:1.6;padding:40px 20px}
.container{max-width:1200px;margin:0 auto}
h1{font-size:2rem;margin-bottom:8px;background:linear-gradient(135deg,#38bdf8,#818cf8,#c084fc);-webkit-background-clip:text;-webkit-text-fill-color:transparent}
h2{font-size:1.3rem;margin:32px 0 16px;color:#e2e8f0;border-bottom:1px solid #1e293b;padding-bottom:8px}
.subtitle{color:#94a3b8;margin-bottom:32px}
.kpi{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:32px}
.kpi-card{background:#1e293b;border:1px solid #334155;border-radius:12px;padding:20px}
.kpi-label{font-size:0.75rem;color:#94a3b8;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:6px}
.kpi-value{font-size:1.6rem;font-weight:700}
.pos{color:#34d399}.neg{color:#f87171}
table{width:100%;border-collapse:collapse;font-size:0.85rem;margin-bottom:24px}
th{background:#1e293b;color:#94a3b8;font-weight:600;text-align:left;padding:10px 12px;border-bottom:2px solid #334155;position:sticky;top:0}
td{padding:8px 12px;border-bottom:1px solid #1e293b}
tr:hover td{background:rgba(56,189,248,0.05)}
.table-wrap{overflow-x:auto;border:1px solid #334155;border-radius:12px;margin-bottom:32px}
.note{color:#64748b;font-size:0.85rem;margin-top:32px;border-top:1px solid #1e293b;padding-top:16px}
.warning{color:#fbbf24;font-size:0.85rem;margin-top:16px;padding:12px;background:rgba(251,191,36,0.1);border:1px solid rgba(251,191,36,0.2);border-radius:8px}
@media(max-width:768px){.kpi{grid-template-columns:repeat(2,1fr)}}
</style>
</head>
<body>
<div class="container">
  <h1>基金投资组合分析报告</h1>
  <p class="subtitle">报告日期：${new Date().toISOString().split('T')[0]}</p>

  <div class="kpi">
    <div class="kpi-card">
      <div class="kpi-label">组合总成本</div>
      <div class="kpi-value">${formatMoney(r.totalCost)}</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">组合总市值</div>
      <div class="kpi-value">${formatMoney(r.totalValue)}</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">总收益率</div>
      <div class="kpi-value ${r.portfolioReturn >= 0 ? 'pos' : 'neg'}">${(r.portfolioReturn >= 0 ? '+' : '') + (r.portfolioReturn * 100).toFixed(2)}%</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-label">2025年加权收益</div>
      <div class="kpi-value ${r.weightedReturns['2025'] >= 0 ? 'pos' : 'neg'}">${r.weightedReturns['2025'] !== null ? (r.weightedReturns['2025'] >= 0 ? '+' : '') + (r.weightedReturns['2025'] * 100).toFixed(2) + '%' : 'N/A'}</div>
    </div>
  </div>

  <h2>组合加权年度收益率</h2>
  <div class="table-wrap"><table><thead><tr><th>年份</th><th style="text-align:right">加权收益率</th></tr></thead><tbody>${yearlyRows}</tbody></table></div>

  <h2>基金业绩明细（含各年度收益）</h2>
  <div class="table-wrap"><table>
    <thead><tr>
      <th>基金名称</th><th>分类</th><th>购买时间</th><th style="text-align:right">成本价</th>
      <th style="text-align:right">当前净值</th><th style="text-align:right">市值</th>
      <th style="text-align:right">总收益</th><th style="text-align:right">年化收益</th>
      ${years.map(y => '<th style="text-align:right">' + y + '年</th>').join('')}
    </tr></thead>
    <tbody>${fundRows}</tbody>
  </table></div>

  <div class="warning">
    <strong>注意：</strong>公募基金数据为模拟数据，仅供参考。私募基金数据来源于用户上传文件。
  </div>

  <p class="note">数据仅供参考，投资有风险。</p>
</div>
</body></html>`;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '基金分析报告.html';
    a.click();
    URL.revokeObjectURL(url);
    showToast('HTML 报告已导出');
  } catch (err) {
    showToast('导出 HTML 失败: ' + err.message, 'error');
    console.error('HTML export error:', err);
  }
}

// ============================================================================
// Event Listeners
// ============================================================================

document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
document.getElementById('exportHtmlBtn').addEventListener('click', exportToHtml);

// Refresh button
const refreshBtn = document.getElementById('refreshBtn');
if (refreshBtn) {
  refreshBtn.addEventListener('click', () => {
    if (!analysisResults) {
      showToast('请先生成分析报告', 'warning');
      return;
    }
    showToast('正在刷新数据...', 'info');
    // Re-render all charts
    disposeAllCharts();
    requestAnimationFrame(() => {
      setTimeout(() => {
        initAllCharts();
        showToast('数据已刷新');
      }, 50);
    });
  });
}

// Resize charts on window resize with debounce
let resizeTimeout;
window.addEventListener('resize', () => {
  clearTimeout(resizeTimeout);
  resizeTimeout = setTimeout(() => {
    Object.values(chartInstances).forEach(chart => {
      if (chart && !chart.isDisposed()) {
        chart.resize();
      }
    });
  }, 150);
});

// Also handle visibility changes (tab switching)
document.addEventListener('visibilitychange', () => {
  if (!document.hidden) {
    setTimeout(() => {
      Object.values(chartInstances).forEach(chart => {
        if (chart && !chart.isDisposed()) {
          chart.resize();
        }
      });
    }, 100);
  }
});
