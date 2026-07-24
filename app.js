// State Management
let allData = window.RAW_DATA || [];
let filteredData = [...allData];

let currentFilters = {
  year: 'ALL',
  category: 'ALL',
  warehouse: 'ALL',
  search: ''
};

let tableState = {
  page: 1,
  pageSize: 12,
  sortCol: 'thang',
  sortDir: 'asc'
};

// Chart instances
let chartCategoryYear = null;
let chartDonutCategory = null;
let chartMonthlyTrend = null;
let chartTopWarehouse = null;

// Formatters
const formatVND = (amount) => {
  if (amount === undefined || amount === null || isNaN(amount)) return '0 ₫';
  return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(amount);
};

const formatShortVND = (amount) => {
  if (amount >= 1e9) return (amount / 1e9).toFixed(2) + ' Tỷ';
  if (amount >= 1e6) return (amount / 1e6).toFixed(1) + ' Tr';
  if (amount >= 1e3) return (amount / 1e3).toFixed(0) + ' K';
  return amount.toLocaleString('vi-VN') + ' ₫';
};

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
  initFilterDropdowns();
  setupEventListeners();
  updateSyncBadge();
  applyFilters();
});

function updateSyncBadge() {
  const badge = document.getElementById('sync-badge');
  if (!badge) return;
  if (window.SYNC_INFO && window.SYNC_INFO.last_updated) {
    badge.innerHTML = `<i class="fas fa-check-circle" style="color:#10b981; margin-right:4px;"></i> Đã đồng bộ lúc: ${window.SYNC_INFO.last_updated}`;
  } else {
    badge.innerHTML = `<i class="fas fa-check-circle" style="color:#10b981; margin-right:4px;"></i> Dữ liệu sẵn sàng`;
  }
}

// Populate Filter Options dynamically
function initFilterDropdowns() {
  const years = setOfValues('nam').sort((a, b) => b - a);
  const selYear = document.getElementById('filter-year');
  selYear.innerHTML = '<option value="ALL">-- Tất Cả Các Năm --</option>';

  years.forEach(y => {
    if (y && y !== 'N/A' && Number(y) > 1900) {
      const opt = document.createElement('option');
      opt.value = y;
      opt.textContent = `Năm ${y}`;
      selYear.appendChild(opt);
    }
  });

  updateDependentDropdowns();
}

function updateDependentDropdowns() {
  const selCat = document.getElementById('filter-category');
  const selWh = document.getElementById('filter-warehouse');

  const currYear = currentFilters.year;
  const currCat = currentFilters.category;
  const currWh = currentFilters.warehouse;

  // 1. Data subset for Category dropdown (filtered by active Year and Search)
  let catData = allData;
  if (currYear !== 'ALL') {
    catData = catData.filter(item => String(item.nam) === String(currYear));
  }
  if (currentFilters.search) {
    const s = currentFilters.search;
    catData = catData.filter(item =>
      (item.loai_cp && item.loai_cp.toLowerCase().includes(s)) ||
      (item.tieu_muc && item.tieu_muc.toLowerCase().includes(s)) ||
      (item.so_hd && item.so_hd.toLowerCase().includes(s)) ||
      (item.ly_do && item.ly_do.toLowerCase().includes(s)) ||
      (item.chi_tiet && item.chi_tiet.toLowerCase().includes(s)) ||
      (item.kho && item.kho.toLowerCase().includes(s)) ||
      (item.nguoi_thu_huong && item.nguoi_thu_huong.toLowerCase().includes(s))
    );
  }

  const validCategories = new Set(catData.map(i => i.loai_cp).filter(Boolean));
  const sortedCategories = Array.from(validCategories).sort();

  selCat.innerHTML = '<option value="ALL">-- Tất Cả Loại CP --</option>';
  sortedCategories.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c;
    opt.textContent = c;
    if (c === currCat) opt.selected = true;
    selCat.appendChild(opt);
  });

  if (currCat !== 'ALL' && !validCategories.has(currCat)) {
    currentFilters.category = 'ALL';
    selCat.value = 'ALL';
  }

  // 2. Data subset for Warehouse dropdown (filtered by active Year, Category, and Search)
  let whData = catData;
  if (currentFilters.category !== 'ALL') {
    whData = whData.filter(item => item.loai_cp === currentFilters.category);
  }

  const validWarehouses = new Set(whData.map(i => i.kho).filter(Boolean));
  const sortedWarehouses = Array.from(validWarehouses).sort();

  selWh.innerHTML = '<option value="ALL">-- Tất Cả Kho --</option>';
  sortedWarehouses.forEach(w => {
    const opt = document.createElement('option');
    opt.value = w;
    opt.textContent = w;
    if (w === currWh) opt.selected = true;
    selWh.appendChild(opt);
  });

  if (currWh !== 'ALL' && !validWarehouses.has(currWh)) {
    currentFilters.warehouse = 'ALL';
    selWh.value = 'ALL';
  }
}

function setOfValues(key) {
  const set = new Set();
  allData.forEach(item => {
    if (item[key] !== undefined && item[key] !== null && item[key] !== '') {
      set.add(item[key]);
    }
  });
  return Array.from(set);
}

// Event Listeners
function setupEventListeners() {
  document.getElementById('filter-year').addEventListener('change', (e) => {
    currentFilters.year = e.target.value;
    applyFilters();
  });

  document.getElementById('filter-category').addEventListener('change', (e) => {
    currentFilters.category = e.target.value;
    applyFilters();
  });

  document.getElementById('filter-warehouse').addEventListener('change', (e) => {
    currentFilters.warehouse = e.target.value;
    applyFilters();
  });

  document.getElementById('filter-search').addEventListener('input', (e) => {
    currentFilters.search = e.target.value.toLowerCase().trim();
    applyFilters();
  });

  document.getElementById('btn-reset').addEventListener('click', () => {
    document.getElementById('filter-year').value = 'ALL';
    document.getElementById('filter-category').value = 'ALL';
    document.getElementById('filter-warehouse').value = 'ALL';
    document.getElementById('filter-search').value = '';
    currentFilters = { year: 'ALL', category: 'ALL', warehouse: 'ALL', search: '' };
    applyFilters();
  });

  const btnRefresh = document.getElementById('btn-refresh');
  if (btnRefresh) {
    btnRefresh.addEventListener('click', async () => {
      const icon = document.getElementById('icon-refresh');
      if (icon) icon.classList.add('fa-spin');
      try {
        await fetch('/api/refresh?t=' + Date.now(), { cache: 'no-store' });
      } catch (e) {}
      window.location.href = window.location.pathname + '?t=' + Date.now();
    });
  }

  document.getElementById('btn-export').addEventListener('click', exportToCSV);

  // Table header sorting
  document.querySelectorAll('#detail-table th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.getAttribute('data-sort');
      if (tableState.sortCol === col) {
        tableState.sortDir = tableState.sortDir === 'asc' ? 'desc' : 'asc';
      } else {
        tableState.sortCol = col;
        tableState.sortDir = 'asc';
      }
      renderDataTable();
    });
  });
}

// Apply Filters
function applyFilters() {
  updateDependentDropdowns();

  filteredData = allData.filter(item => {
    if (currentFilters.year !== 'ALL' && String(item.nam) !== String(currentFilters.year)) return false;
    if (currentFilters.category !== 'ALL' && item.loai_cp !== currentFilters.category) return false;
    if (currentFilters.warehouse !== 'ALL' && item.kho !== currentFilters.warehouse) return false;

    if (currentFilters.search) {
      const s = currentFilters.search;
      const match = 
        (item.loai_cp && item.loai_cp.toLowerCase().includes(s)) ||
        (item.tieu_muc && item.tieu_muc.toLowerCase().includes(s)) ||
        (item.so_hd && item.so_hd.toLowerCase().includes(s)) ||
        (item.ly_do && item.ly_do.toLowerCase().includes(s)) ||
        (item.chi_tiet && item.chi_tiet.toLowerCase().includes(s)) ||
        (item.kho && item.kho.toLowerCase().includes(s)) ||
        (item.nguoi_thu_huong && item.nguoi_thu_huong.toLowerCase().includes(s));
      if (!match) return false;
    }
    return true;
  });

  tableState.page = 1;

  updateKPIs();
  renderMatrixPivot();
  renderCharts();
  renderDataTable();
}

// Update KPI Metric Cards
function updateKPIs() {
  const totalAmount = filteredData.reduce((sum, item) => sum + (item.st_vat || 0), 0);
  const totalItems = filteredData.length;
  const avgAmount = totalItems > 0 ? totalAmount / totalItems : 0;

  // Category with highest spending in filtered data
  const catMap = {};
  filteredData.forEach(item => {
    catMap[item.loai_cp] = (catMap[item.loai_cp] || 0) + (item.st_vat || 0);
  });
  let topCat = 'N/A';
  let topCatAmt = 0;
  Object.entries(catMap).forEach(([cat, val]) => {
    if (val > topCatAmt) {
      topCatAmt = val;
      topCat = cat;
    }
  });

  document.getElementById('kpi-total-amount').textContent = formatVND(totalAmount);
  document.getElementById('kpi-total-items').textContent = totalItems.toLocaleString('vi-VN') + ' mục';
  document.getElementById('kpi-avg-amount').textContent = formatVND(avgAmount);
  document.getElementById('kpi-top-category').textContent = topCat;
  document.getElementById('kpi-top-category-sub').textContent = topCatAmt > 0 ? `Tổng: ${formatShortVND(topCatAmt)}` : '';
}

// Render Matrix Pivot Table (Loại CP x Năm)
function renderMatrixPivot() {
  const years = [2023, 2024, 2025, 2026].filter(y => {
    if (currentFilters.year !== 'ALL') return String(y) === String(currentFilters.year);
    return true;
  });

  // Group filtered data by Category and Year
  const matrix = {};
  const catTotals = {};
  const yearTotals = {};
  years.forEach(y => yearTotals[y] = 0);
  let grandTotal = 0;

  filteredData.forEach(item => {
    const cat = item.loai_cp;
    const yr = item.nam;

    if (!matrix[cat]) {
      matrix[cat] = {};
      years.forEach(y => matrix[cat][y] = 0);
      catTotals[cat] = 0;
    }

    if (matrix[cat][yr] !== undefined) {
      matrix[cat][yr] += (item.st_vat || 0);
      catTotals[cat] += (item.st_vat || 0);
      yearTotals[yr] += (item.st_vat || 0);
      grandTotal += (item.st_vat || 0);
    }
  });

  // Build Table HTML
  const thead = document.getElementById('matrix-thead');
  const tbody = document.getElementById('matrix-tbody');

  let headHTML = `<tr><th>Loại Chi Phí (Cột A)</th>`;
  years.forEach(y => headHTML += `<th style="text-align:right">Năm ${y}</th>`);
  headHTML += `<th style="text-align:right">TỔNG CỘNG</th><th style="text-align:right">% TỶ TRỌNG</th></tr>`;
  thead.innerHTML = headHTML;

  let bodyHTML = '';
  const sortedCategories = Object.keys(matrix).sort((a, b) => catTotals[b] - catTotals[a]);

  sortedCategories.forEach(cat => {
    const tot = catTotals[cat];
    const pct = grandTotal > 0 ? ((tot / grandTotal) * 100).toFixed(1) + '%' : '0%';

    bodyHTML += `<tr>`;
    bodyHTML += `<td><strong>${cat}</strong></td>`;
    years.forEach(y => {
      const val = matrix[cat][y];
      bodyHTML += `<td>${val > 0 ? formatShortVND(val) : '-'}</td>`;
    });
    bodyHTML += `<td class="highlight">${formatVND(tot)}</td>`;
    bodyHTML += `<td style="color:var(--text-muted); font-size:0.82rem">${pct}</td>`;
    bodyHTML += `</tr>`;
  });

  // Total Row
  bodyHTML += `<tr class="total-row"><td>TỔNG CỘNG BỘ PHẬN</td>`;
  years.forEach(y => {
    bodyHTML += `<td>${formatShortVND(yearTotals[y])}</td>`;
  });
  bodyHTML += `<td>${formatVND(grandTotal)}</td><td>100%</td></tr>`;

  tbody.innerHTML = bodyHTML;
}

// Render Chart.js Visualizations
function renderCharts() {
  const textColor = '#94a3b8';
  const gridColor = 'rgba(255, 255, 255, 0.07)';

  // Chart 1: Category Expenditure across Years (Bar Chart)
  const categories = setOfValues('loai_cp').sort();
  const years = [2023, 2024, 2025, 2026];
  const yearColors = ['#06b6d4', '#3b82f6', '#8b5cf6', '#10b981'];

  const datasets = years.map((yr, idx) => {
    const data = categories.map(cat => {
      return filteredData
        .filter(item => item.loai_cp === cat && item.nam === yr)
        .reduce((sum, item) => sum + (item.st_vat || 0), 0) / 1e6; // In Millions
    });
    return {
      label: `Năm ${yr}`,
      data: data,
      backgroundColor: yearColors[idx % yearColors.length],
      borderRadius: 6,
      borderWidth: 0
    };
  });

  const ctxBar = document.getElementById('chartCategoryYear').getContext('2d');
  if (chartCategoryYear) chartCategoryYear.destroy();

  chartCategoryYear = new Chart(ctxBar, {
    type: 'bar',
    data: { labels: categories, datasets: datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: textColor, font: { family: 'Plus Jakarta Sans', weight: '600' } } },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          titleColor: '#06b6d4',
          borderColor: 'rgba(255, 255, 255, 0.15)',
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.raw.toLocaleString('vi-VN')} Triệu VNĐ`
          }
        }
      },
      scales: {
        x: { ticks: { color: textColor, font: { size: 10 } }, grid: { color: gridColor } },
        y: { 
          ticks: { 
            color: textColor,
            callback: (val) => val >= 1000 ? (val/1000).toFixed(1) + ' Tỷ' : val + ' Tr'
          }, 
          grid: { color: gridColor } 
        }
      }
    }
  });

  // Chart 2: Category % Share (Donut Chart)
  const catMap = {};
  filteredData.forEach(item => {
    catMap[item.loai_cp] = (catMap[item.loai_cp] || 0) + (item.st_vat || 0);
  });
  const donutLabels = Object.keys(catMap);
  const donutValues = Object.values(catMap).map(v => (v / 1e6).toFixed(1));

  const ctxDonut = document.getElementById('chartDonutCategory').getContext('2d');
  if (chartDonutCategory) chartDonutCategory.destroy();

  chartDonutCategory = new Chart(ctxDonut, {
    type: 'doughnut',
    data: {
      labels: donutLabels,
      datasets: [{
        data: donutValues,
        backgroundColor: [
          '#06b6d4', '#3b82f6', '#8b5cf6', '#10b981', '#f59e0b', 
          '#ec4899', '#f43f5e', '#a855f7', '#64748b', '#14b8a6'
        ],
        borderWidth: 2,
        borderColor: '#0f172a'
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { color: textColor, font: { size: 11, family: 'Plus Jakarta Sans' } } },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          borderColor: 'rgba(255, 255, 255, 0.15)',
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label: (ctx) => ` ${ctx.label}: ${parseFloat(ctx.raw).toLocaleString('vi-VN')} Triệu VNĐ`
          }
        }
      },
      cutout: '68%'
    }
  });

  // Chart 3: Top Warehouses (Horizontal Bar)
  const whMap = {};
  filteredData.forEach(item => {
    const k = item.kho || 'Khác';
    whMap[k] = (whMap[k] || 0) + (item.st_vat || 0);
  });
  const sortedWh = Object.entries(whMap).sort((a, b) => b[1] - a[1]).slice(0, 8);

  const ctxWh = document.getElementById('chartTopWarehouse').getContext('2d');
  if (chartTopWarehouse) chartTopWarehouse.destroy();

  chartTopWarehouse = new Chart(ctxWh, {
    type: 'bar',
    data: {
      labels: sortedWh.map(x => x[0]),
      datasets: [{
        label: 'Tổng Chi Phí',
        data: sortedWh.map(x => (x[1] / 1e6).toFixed(1)),
        backgroundColor: 'rgba(6, 182, 212, 0.85)',
        borderColor: '#06b6d4',
        borderWidth: 1,
        borderRadius: 6
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          borderColor: 'rgba(255, 255, 255, 0.15)',
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label: (ctx) => ` Chi phí: ${parseFloat(ctx.raw).toLocaleString('vi-VN')} Triệu VNĐ`
          }
        }
      },
      scales: {
        x: { ticks: { color: textColor }, grid: { color: gridColor } },
        y: { ticks: { color: textColor }, grid: { display: false } }
      }
    }
  });

  // Chart 4: Monthly Trend Line Chart
  const monthMap = {};
  filteredData.forEach(item => {
    if (item.thang) {
      monthMap[item.thang] = (monthMap[item.thang] || 0) + (item.st_vat || 0);
    }
  });
  const sortedMonths = Object.keys(monthMap).sort();

  const ctxTrend = document.getElementById('chartMonthlyTrend').getContext('2d');
  if (chartMonthlyTrend) chartMonthlyTrend.destroy();

  chartMonthlyTrend = new Chart(ctxTrend, {
    type: 'line',
    data: {
      labels: sortedMonths,
      datasets: [{
        label: 'Chi phí theo Tháng',
        data: sortedMonths.map(m => (monthMap[m] / 1e6).toFixed(1)),
        borderColor: '#10b981',
        backgroundColor: 'rgba(16, 185, 129, 0.18)',
        fill: true,
        tension: 0.4,
        pointRadius: 5,
        pointBackgroundColor: '#10b981',
        pointBorderColor: '#ffffff',
        pointBorderWidth: 2,
        pointHoverRadius: 8
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: textColor, font: { family: 'Plus Jakarta Sans' } } },
        tooltip: {
          backgroundColor: 'rgba(15, 23, 42, 0.95)',
          borderColor: 'rgba(16, 185, 129, 0.4)',
          borderWidth: 1,
          padding: 12,
          callbacks: {
            label: (ctx) => ` ${ctx.raw} Triệu VNĐ`
          }
        }
      },
      scales: {
        x: { ticks: { color: textColor }, grid: { color: gridColor } },
        y: { ticks: { color: textColor }, grid: { color: gridColor } }
      }
    }
  });
}

// Render Data Table with Pagination & Sorting
function renderDataTable() {
  const sorted = [...filteredData].sort((a, b) => {
    let valA = a[tableState.sortCol];
    let valB = b[tableState.sortCol];
    if (typeof valA === 'string') valA = valA.toLowerCase();
    if (typeof valB === 'string') valB = valB.toLowerCase();

    if (valA < valB) return tableState.sortDir === 'asc' ? -1 : 1;
    if (valA > valB) return tableState.sortDir === 'asc' ? 1 : -1;
    return 0;
  });

  const totalPages = Math.ceil(sorted.length / tableState.pageSize) || 1;
  if (tableState.page > totalPages) tableState.page = totalPages;

  const startIdx = (tableState.page - 1) * tableState.pageSize;
  const pageItems = sorted.slice(startIdx, startIdx + tableState.pageSize);

  const tbody = document.getElementById('detail-tbody');
  if (pageItems.length === 0) {
    tbody.innerHTML = `<tr><td colspan="9" style="text-align:center; padding: 2rem; color: var(--text-muted)">Không tìm thấy dữ liệu phù hợp</td></tr>`;
  } else {
    tbody.innerHTML = pageItems.map(item => {
      const chiTietHD = item.chi_tiet_hd || item.chi_tiet || '-';
      const lyDo = item.ly_do || '-';
      return `
        <tr>
          <td><span class="badge badge-emerald">${item.thang || '-'}</span></td>
          <td>${item.tieu_muc || '-'}</td>
          <td><strong>${item.so_hd || '-'}</strong></td>
          <td>${item.ngay_hd || '-'}</td>
          <td title="${lyDo}">${lyDo.length > 30 ? lyDo.substring(0, 30) + '...' : lyDo}</td>
          <td title="${chiTietHD}">${chiTietHD.length > 30 ? chiTietHD.substring(0, 30) + '...' : chiTietHD}</td>
          <td style="font-weight:600; text-align:right; color:#38bdf8">${formatVND(item.st_no_vat)}</td>
          <td style="font-weight:600; text-align:right; color:#f59e0b">${formatVND(item.vat)}</td>
          <td style="font-weight:700; text-align:right; color:#34d399">${formatVND(item.st_vat)}</td>
        </tr>
      `;
    }).join('');
  }

  // Calculate Grand Totals for Footer Row across all filtered items
  const totalStNoVat = filteredData.reduce((sum, item) => sum + (item.st_no_vat || 0), 0);
  const totalVat = filteredData.reduce((sum, item) => sum + (item.vat || 0), 0);
  const totalStVat = filteredData.reduce((sum, item) => sum + (item.st_vat || 0), 0);

  const tfoot = document.getElementById('detail-tfoot');
  if (tfoot) {
    tfoot.innerHTML = `
      <tr class="table-total-row">
        <td colspan="6" style="text-align:right; font-weight:800; color:var(--accent-cyan);">TỔNG CỘNG (${filteredData.length.toLocaleString('vi-VN')} MỤC):</td>
        <td style="text-align:right; font-weight:800; color:#38bdf8;">${formatVND(totalStNoVat)}</td>
        <td style="text-align:right; font-weight:800; color:#f59e0b;">${formatVND(totalVat)}</td>
        <td style="text-align:right; font-weight:800; color:#34d399;">${formatVND(totalStVat)}</td>
      </tr>
    `;
  }

  // Update Pagination Info
  document.getElementById('table-info').textContent = 
    `Hiển thị ${sorted.length > 0 ? startIdx + 1 : 0} - ${Math.min(startIdx + tableState.pageSize, sorted.length)} trên tổng số ${sorted.length} mục`;

  renderPaginationControls(totalPages);
}

function renderPaginationControls(totalPages) {
  const container = document.getElementById('pagination-btns');
  let html = '';

  html += `<button class="page-btn" ${tableState.page === 1 ? 'disabled' : ''} onclick="changePage(${tableState.page - 1})"><i class="fas fa-chevron-left"></i></button>`;

  for (let p = 1; p <= totalPages; p++) {
    if (p === 1 || p === totalPages || (p >= tableState.page - 1 && p <= tableState.page + 1)) {
      html += `<button class="page-btn ${p === tableState.page ? 'active' : ''}" onclick="changePage(${p})">${p}</button>`;
    } else if (p === tableState.page - 2 || p === tableState.page + 2) {
      html += `<span style="color:var(--text-muted)">...</span>`;
    }
  }

  html += `<button class="page-btn" ${tableState.page === totalPages ? 'disabled' : ''} onclick="changePage(${tableState.page + 1})"><i class="fas fa-chevron-right"></i></button>`;

  container.innerHTML = html;
}

function changePage(page) {
  tableState.page = page;
  renderDataTable();
}

// Export Filtered Data to CSV
function exportToCSV() {
  if (filteredData.length === 0) {
    alert('Không có dữ liệu để xuất!');
    return;
  }

  const headers = ['STT', 'Loại CP', 'Tiểu Mục', 'Số HĐ', 'Ngày HĐ', 'Tháng', 'Lý do TT', 'Chi Tiết', 'Kho', 'Số tiền (chưa VAT)', 'VAT', 'Số tiền (VAT)', 'Ngày đề nghị TT', 'Người thụ hưởng', 'Ngân hàng', 'Năm'];
  
  const csvRows = [headers.join(',')];

  filteredData.forEach(item => {
    const row = [
      item.id,
      `"${item.loai_cp.replace(/"/g, '""')}"`,
      `"${item.tieu_muc.replace(/"/g, '""')}"`,
      `"${item.so_hd.replace(/"/g, '""')}"`,
      `"${item.ngay_hd}"`,
      `"${item.thang}"`,
      `"${item.ly_do.replace(/"/g, '""')}"`,
      `"${item.chi_tiet.replace(/"/g, '""')}"`,
      `"${item.kho.replace(/"/g, '""')}"`,
      item.st_no_vat,
      item.vat,
      item.st_vat,
      `"${item.ngay_tt}"`,
      `"${item.nguoi_thu_huong.replace(/"/g, '""')}"`,
      `"${item.ngan_hang.replace(/"/g, '""')}"`,
      item.nam
    ];
    csvRows.push(row.join(','));
  });

  const blob = new Blob(['\ufeff' + csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.setAttribute('href', url);
  link.setAttribute('download', `Bao_Cao_Chi_Phi_${currentFilters.year}_${new Date().toISOString().slice(0,10)}.csv`);
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
