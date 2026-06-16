/**
 * Fund Portfolio Analyzer - Web Application
 * Handles file upload, Excel parsing, calculations, and visualization
 */

// Global state
let portfolioData = null;
let privateFundData = null;
let analysisResults = null;

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

// File upload handlers
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
        alert('请上传 Excel 文件 (.xlsx 或 .xls)');
      }
    }
  });
}

function handleFile(file, placeholder, fileInfo, fileNameEl, fileSizeEl, onUpload) {
  placeholder.classList.add('hidden');
  fileInfo.classList.remove('hidden');
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = formatFileSize(file.size);
  
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    onUpload(workbook);
  };
  reader.readAsArrayBuffer(file);
}

function formatFileSize(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

// Setup uploads
setupUpload(
  portfolioInput, portfolioDropzone,
  document.getElementById('portfolioPlaceholder'),
  document.getElementById('portfolioFileInfo'),
  document.getElementById('portfolioFileName'),
  document.getElementById('portfolioFileSize'),
  (workbook) => {
    portfolioData = parsePortfolioWorkbook(workbook);
    checkReady();
  }
);

setupUpload(
  privateInput, privateDropzone,
  document.getElementById('privatePlaceholder'),
  document.getElementById('privateFileInfo'),
  document.getElementById('privateFileName'),
  document.getElementById('privateFileSize'),
  (workbook) => {
    privateFundData = parsePrivateFundWorkbook(workbook);
    checkReady();
  }
);

function checkReady() {
  if (portfolioData) {
    generateBtn.disabled = false;
    statusText.textContent = privateFundData 
      ? '两个文件已上传，点击生成报告' 
      : '持仓文件已上传，可点击生成报告（私募基金数据可选）';
    statusText.classList.add('text-emerald-400');
  }
}

// Parse portfolio workbook
function parsePortfolioWorkbook(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  
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
    
    if (fund['基金代码'] || fund['Fund Code']) {
      funds.push({
        code: String(fund['基金代码'] || fund['Fund Code'] || ''),
        name: String(fund['产品名称'] || fund['Fund Name'] || ''),
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

// Parse private fund workbook
function parsePrivateFundWorkbook(workbook) {
  const returns = {};
  
  workbook.SheetNames.forEach(sheetName => {
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    // Find header row with years
    let headerRow = -1;
    let codeCol = -1;
    let yearCols = {};
    
    for (let i = 0; i < Math.min(15, jsonData.length); i++) {
      const row = jsonData[i];
      if (!row) continue;
      
      row.forEach((cell, idx) => {
        if (cell === '产品代码' || cell === '基金代码' || cell === 'Code') codeCol = idx;
        if (typeof cell === 'string' && /^20\d{2}$/.test(cell.trim())) {
          yearCols[cell.trim()] = idx;
          headerRow = i;
        }
      });
    }
    
    if (headerRow >= 0 && codeCol >= 0) {
      for (let i = headerRow + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || !row[codeCol]) continue;
        
        const code = String(row[codeCol]).trim();
        returns[code] = {};
        
        Object.entries(yearCols).forEach(([year, col]) => {
          const val = row[col];
          if (val !== undefined && val !== null && val !== '') {
            returns[code][year] = parseNumber(val) / 100;
          }
        });
      }
    }
  });
  
  return returns;
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

// Generate report
generateBtn.addEventListener('click', async () => {
  if (!portfolioData) return;
  
  // Show loading
  loadingSection.classList.remove('hidden');
  resultsSection.classList.add('hidden');
  generateBtn.disabled = true;
  
  // Simulate progress
  const steps = [
    { text: '读取基金持仓文件', progress: 10 },
    { text: '解析私募基金业绩数据', progress: 25 },
    { text: '计算各基金年化收益率', progress: 45 },
    { text: '计算组合加权收益', progress: 65 },
    { text: '生成可视化图表', progress: 85 },
    { text: '报告生成完成', progress: 100 },
  ];
  
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

  // IMPORTANT: Resize charts after container becomes visible
  // ECharts needs visible containers to calculate dimensions correctly
  setTimeout(() => {
    ['yearlyChart', 'categoryChart', 'annualizedChart', 'heatmapChart'].forEach(id => {
      const chart = echarts.getInstanceByDom(document.getElementById(id));
      if (chart) chart.resize();
    });
  }, 100);
});

function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function calculateAnalysis() {
  const funds = portfolioData.map(fund => {
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
    
    // Get yearly returns from private fund data
    const yearlyReturns = {};
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
  });
  
  // Calculate portfolio totals
  const totalCost = funds.reduce((sum, f) => sum + (f.costPrice * f.shares), 0);
  const totalValue = funds.reduce((sum, f) => sum + f.marketValue, 0);
  const portfolioReturn = totalCost > 0 ? (totalValue - totalCost) / totalCost : 0;
  
  // Calculate weighted yearly returns
  const years = ['2018', '2019', '2020', '2021', '2022', '2023', '2024', '2025'];
  const weightedReturns = {};
  
  years.forEach(year => {
    let totalWeight = 0;
    let weightedSum = 0;
    
    funds.forEach(fund => {
      const ret = fund.yearlyReturns[year];
      if (ret !== undefined && fund.marketValue > 0) {
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

function renderResults() {
  const r = analysisResults;
  
  // KPI Cards
  document.getElementById('kpiCost').textContent = formatMoney(r.totalCost);
  document.getElementById('kpiValue').textContent = formatMoney(r.totalValue);
  document.getElementById('kpiReturn').textContent = formatPercent(r.portfolioReturn);
  document.getElementById('kpiReturn').className = 'text-2xl font-bold ' + (r.portfolioReturn >= 0 ? 'text-emerald-400' : 'text-red-400');
  
  const ret2025 = r.weightedReturns['2025'];
  document.getElementById('kpi2025').textContent = ret2025 !== null ? formatPercent(ret2025) : 'N/A';
  document.getElementById('kpi2025').className = 'text-2xl font-bold ' + (ret2025 >= 0 ? 'text-emerald-400' : 'text-red-400');
  
  // Yearly Chart
  renderYearlyChart(r.weightedReturns);
  
  // Category Chart
  renderCategoryChart(r.categories);
  
  // Annualized Chart
  renderAnnualizedChart(r.funds);
  
  // Heatmap
  renderHeatmap(r.funds);
  
  // Table
  renderTable(r.funds);
}

function renderYearlyChart(weightedReturns) {
  const chart = echarts.init(document.getElementById('yearlyChart'));
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
  const chart = echarts.init(document.getElementById('categoryChart'));
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
  const chart = echarts.init(document.getElementById('annualizedChart'));
  
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

function renderHeatmap(funds) {
  const chart = echarts.init(document.getElementById('heatmapChart'));
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

function renderTable(funds) {
  const tbody = document.getElementById('fundTableBody');
  tbody.innerHTML = '';
  
  funds.forEach(fund => {
    const tr = document.createElement('tr');
    tr.className = 'border-b border-slate-800/50 hover:bg-slate-800/30';
    
    const totalReturn = fund.totalReturn;
    const annReturn = fund.annualizedReturn;
    
    tr.innerHTML = `
      <td class="px-4 py-3 text-slate-200">${fund.name}</td>
      <td class="px-4 py-3">
        <span class="px-2 py-0.5 rounded text-xs ${getCategoryStyle(fund.category)}">${fund.category || '其他'}</span>
      </td>
      <td class="px-4 py-3 text-right text-slate-300">${fund.costPrice.toFixed(2)}</td>
      <td class="px-4 py-3 text-right text-slate-300">${fund.nav.toFixed(3)}</td>
      <td class="px-4 py-3 text-right text-slate-300">${formatMoney(fund.marketValue)}</td>
      <td class="px-4 py-3 text-right ${totalReturn >= 0 ? 'text-emerald-400' : 'text-red-400'}">${formatPercent(totalReturn)}</td>
      <td class="px-4 py-3 text-right ${annReturn !== null ? (annReturn >= 0 ? 'text-emerald-400' : 'text-red-400') : 'text-slate-500'}">${annReturn !== null ? formatPercent(annReturn) : 'N/A'}</td>
    `;
    
    tbody.appendChild(tr);
  });
}

function getCategoryStyle(category) {
  const styles = {
    '私募': 'bg-sky-500/20 text-sky-400',
    '公募': 'bg-purple-500/20 text-purple-400',
    '专户': 'bg-emerald-500/20 text-emerald-400',
  };
  return styles[category] || 'bg-slate-500/20 text-slate-400';
}

function formatMoney(val) {
  if (val >= 10000) {
    return (val / 10000).toFixed(2) + '万';
  }
  return val.toFixed(2);
}

function formatPercent(val) {
  if (val === null || val === undefined) return 'N/A';
  return (val >= 0 ? '+' : '') + (val * 100).toFixed(2) + '%';
}

// Export functions
function exportToExcel() {
  if (!analysisResults) return;
  
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
    const purchaseMonth = f.purchaseDate.getMonth();
    
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
  
  // Sheet 3: Summary
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
}

function exportToHtml() {
  if (!analysisResults) return;
  
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
    
    // Calculate yearly cumulative return since purchase
    let cumCells = '';
    if (f.purchaseDate) {
      const pYear = f.purchaseDate.getFullYear();
      let cum = 0;
      for (let y = pYear; y <= 2025; y++) {
        const yr = f.yearlyReturns[String(y)];
        if (yr !== undefined && yr !== null) cum = (1 + cum) * (1 + yr) - 1;
        const c = cum >= 0 ? '#34d399' : '#f87171';
        cumCells += `<td style="padding:6px 8px;text-align:right;color:${c}">${(cum * 100).toFixed(2)}%</td>`;
      }
    }
    
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

  <p class="note">数据仅供参考，投资有风险。私募基金数据来源于用户上传文件，公募基金数据来源于天天基金网。</p>
</div>
</body></html>`;
  
  const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = '基金分析报告.html';
  a.click();
  URL.revokeObjectURL(url);
}

document.getElementById('exportExcelBtn').addEventListener('click', exportToExcel);
document.getElementById('exportHtmlBtn').addEventListener('click', exportToHtml);

// Resize charts on window resize
window.addEventListener('resize', () => {
  ['yearlyChart', 'categoryChart', 'annualizedChart', 'heatmapChart'].forEach(id => {
    const chart = echarts.getInstanceByDom(document.getElementById(id));
    if (chart) chart.resize();
  });
});
