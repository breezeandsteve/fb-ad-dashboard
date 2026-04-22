const DEFAULT_SHEET_NAME = 'Sheet1';
const SHEET_SOURCES = {
  CG: {
    code: 'CG',
    label: 'CG FB SPY',
    sheetId: '1NzSHaQe6puchCA1B-tU2-4VLR1_gHlOQCiCuV9DIltk',
    sheetName: DEFAULT_SHEET_NAME,
  },
  CH: {
    code: 'CH',
    label: 'CH FB SPY',
    sheetId: '1_Ni_mQ4xVJRZ86q5y75KTgtFBhxi45tHM2YeDYGf2dA',
    sheetName: DEFAULT_SHEET_NAME,
  },
  KENNY: {
    code: 'KENNY',
    label: 'KENNY FB SPY',
    sheetId: '1tJbCPvzak9eJjvh7qWHoX0akPMK71PpM__CousYkmwY',
    sheetName: DEFAULT_SHEET_NAME,
  },
};
const ADS_PER_PAGE = 12;
const SIDEBAR_STORAGE_KEY = 'ads-war-room-sidebar';
const SIDEBAR_COLLAPSE_BREAKPOINT = 1120;

const escapeMap = {
  '&': '&amp;',
  '<': '&lt;',
  '>': '&gt;',
  '"': '&quot;',
  "'": '&#39;',
};

export function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (character) => escapeMap[character]);
}

function normalizeSourceCode(value) {
  return String(value ?? '').trim().toUpperCase();
}

export function resolveSheetSource(code) {
  return SHEET_SOURCES[normalizeSourceCode(code)] || null;
}

export function resolveInitialSource() {
  return null;
}

export function buildSheetUrl(source = resolveSheetSource('CG')) {
  return `https://docs.google.com/spreadsheets/d/${source.sheetId}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(source.sheetName)}`;
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function getCellValue(row, index) {
  return row?.c?.[index]?.v ?? '';
}

function buildColumnIndex(columns) {
  return columns.reduce((indexMap, column, index) => {
    const label = normalizeText(column?.label || column?.id).toLowerCase();

    if (label && indexMap[label] === undefined) {
      indexMap[label] = index;
    }

    return indexMap;
  }, {});
}

function getColumnValue(row, columnIndex, label) {
  const index = columnIndex[label];
  return index === undefined ? '' : getCellValue(row, index);
}

function isSafeMediaUrl(url) {
  if (!url) {
    return false;
  }

  try {
    const parsed = new URL(url, 'https://example.com');
    return parsed.protocol === 'http:' || parsed.protocol === 'https:';
  } catch {
    return false;
  }
}

function normalizeText(value) {
  if (Array.isArray(value)) {
    return value.map((item) => String(item ?? '').trim()).filter(Boolean).join(', ');
  }

  return String(value ?? '').trim();
}

export function extractSheetPayload(text) {
  const raw = String(text ?? '').trim();

  if (!raw) {
    throw new Error('數據為空');
  }

  if (raw.startsWith('{')) {
    return JSON.parse(raw);
  }

  const wrappedMatch = raw.match(/setResponse\(([\s\S]+)\);?$/);
  if (wrappedMatch) {
    return JSON.parse(wrappedMatch[1]);
  }

  throw new Error('無法解析資料格式');
}

export function formatAdsFromSheetRows(rows, columns = []) {
  const columnIndex = buildColumnIndex(columns);

  return rows.map((row) => {
    const originalMediaUrl = normalizeText(getColumnValue(row, columnIndex, 'original_media_url'));
    const finalMediaUrl = normalizeText(getColumnValue(row, columnIndex, 'final_media_url'));
    const mediaUrl = isSafeMediaUrl(finalMediaUrl)
      ? finalMediaUrl
      : isSafeMediaUrl(originalMediaUrl)
        ? originalMediaUrl
        : '';

    return {
      ad_archive_id: normalizeText(getColumnValue(row, columnIndex, 'ad_archive_id')),
      type: normalizeText(getColumnValue(row, columnIndex, 'type')) || 'image',
      media_url: mediaUrl,
      start_date: normalizeText(getColumnValue(row, columnIndex, 'start_date')),
      page_name: normalizeText(getColumnValue(row, columnIndex, 'page_name')) || 'Unknown',
      original_feed: normalizeText(getColumnValue(row, columnIndex, 'original_feed')),
      platforms: normalizeText(getColumnValue(row, columnIndex, 'platforms')),
      cta_text: normalizeText(getColumnValue(row, columnIndex, 'cta_text')),
      original_url: normalizeText(getColumnValue(row, columnIndex, 'original_url')),
    };
  });
}

export function highlightText(text, term) {
  const content = String(text ?? '');
  const keyword = String(term ?? '').trim();

  if (!keyword) {
    return escapeHtml(content);
  }

  const regex = new RegExp(`(${escapeRegExp(keyword)})`, 'gi');
  return content
    .split(regex)
    .map((segment) => (
      segment.toLowerCase() === keyword.toLowerCase()
        ? `<span class="highlight">${escapeHtml(segment)}</span>`
        : escapeHtml(segment)
    ))
    .join('');
}

function formatDate(dateStr) {
  if (!dateStr) {
    return 'N/A';
  }

  const googleMatch = String(dateStr).match(/Date\((\d+),(\d+),(\d+)/);
  if (googleMatch) {
    const year = Number.parseInt(googleMatch[1], 10);
    const month = Number.parseInt(googleMatch[2], 10);
    const day = Number.parseInt(googleMatch[3], 10);
    return new Date(year, month, day).toLocaleDateString('zh-HK').replace(/\//g, '-');
  }

  const parsed = new Date(dateStr);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toLocaleDateString('zh-HK').replace(/\//g, '-');
  }

  return String(dateStr);
}

function truncateText(text, maxLength = 140) {
  const content = String(text ?? '').trim();
  if (content.length <= maxLength) {
    return content;
  }

  return `${content.slice(0, maxLength).trimEnd()}…`;
}

export function computeStats(ads) {
  const stats = {
    totalAds: ads.length,
    totalBrands: 0,
    imageAds: 0,
    videoAds: 0,
    brands: {},
    ctas: {},
  };

  ads.forEach((ad) => {
    if (ad.page_name) {
      stats.brands[ad.page_name] = (stats.brands[ad.page_name] || 0) + 1;
    }

    const ctaLabel = ad.cta_text || 'No CTA';
    stats.ctas[ctaLabel] = (stats.ctas[ctaLabel] || 0) + 1;

    if (ad.type === 'video') {
      stats.videoAds += 1;
      return;
    }

    stats.imageAds += 1;
  });

  stats.totalBrands = Object.keys(stats.brands).length;
  return stats;
}

export function shouldCollapseSidebar(storedState, viewportWidth) {
  return storedState === 'collapsed' && viewportWidth > SIDEBAR_COLLAPSE_BREAKPOINT;
}

function createDetailsMarkup(ad, term) {
  return [
    ['原始文案', ad.original_feed],
    ['投放平台', ad.platforms],
    ['行動呼籲', ad.cta_text],
    ['原始連結', ad.original_url],
  ]
    .filter(([, value]) => value)
    .map(([label, value]) => `
      <section class="detail-panel">
        <p class="detail-label">${escapeHtml(label)}</p>
        ${
          label === '原始連結' && isSafeMediaUrl(value)
            ? `<a class="detail-copy detail-link" href="${escapeHtml(value)}" target="_blank" rel="noreferrer">${highlightText(value, term)}</a>`
            : `<div class="detail-copy">${highlightText(value, term)}</div>`
        }
      </section>
    `)
    .join('');
}

function createMediaNode(ad) {
  const frame = document.createElement('div');
  frame.className = 'ad-media-frame';

  if (!ad.media_url) {
    const placeholder = document.createElement('div');
    placeholder.className = 'ad-media-empty';
    placeholder.textContent = 'No media';
    frame.appendChild(placeholder);
    return frame;
  }

  if (ad.type === 'video') {
    const video = document.createElement('video');
    video.className = 'ad-video';
    video.controls = true;
    video.preload = 'metadata';
    video.src = ad.media_url;
    frame.appendChild(video);
    return frame;
  }

  const image = document.createElement('img');
  image.className = 'ad-image';
  image.loading = 'lazy';
  image.alt = ad.page_name || 'Ad preview';
  image.src = ad.media_url;
  frame.appendChild(image);
  return frame;
}

function syncToggleButton(button, expanded) {
  button.setAttribute('aria-expanded', String(expanded));
  button.innerHTML = expanded ? '收起情報層' : '展開情報層';
}

function createAdCard(ad, index, term, expanded) {
  const card = document.createElement('article');
  const detailId = `details-${index}`;
  const detailMarkup = createDetailsMarkup(ad, term);
  const previewSource = ad.original_feed || ad.cta_text || ad.original_url || '沒有文字內容';
  const previewText = truncateText(previewSource);
  const mediaBadge = ad.type === 'video' ? 'VIDEO' : 'IMAGE';
  const metaItems = [formatDate(ad.start_date), ad.platforms].filter(Boolean);

  card.className = 'ad-card';
  card.innerHTML = `
    <div class="ad-head">
      <div>
        <p class="ad-kicker">${escapeHtml(mediaBadge)}</p>
        <h3 class="ad-title">${highlightText(ad.page_name, term)}</h3>
      </div>
      <div class="ad-meta">
        ${metaItems.map((item) => `<span>${escapeHtml(item)}</span>`).join('')}
      </div>
    </div>
    <div class="ad-stage"></div>
    <div class="ad-body">
      <div class="ad-preview">${highlightText(previewText, term)}</div>
      ${
        detailMarkup
          ? `
            <div class="ad-details-wrap${expanded ? ' expanded' : ''}" id="${detailId}">
              <div class="ad-details-grid">${detailMarkup}</div>
            </div>
            <button class="detail-toggle" data-target="${detailId}" type="button"></button>
          `
          : ''
      }
    </div>
  `;

  card.querySelector('.ad-stage').appendChild(createMediaNode(ad));
  const toggleButton = card.querySelector('.detail-toggle');
  if (toggleButton) {
    syncToggleButton(toggleButton, expanded);
  }

  return card;
}

function updateMetricText(element, value) {
  if (element) {
    element.textContent = String(value);
  }
}

function updateCharts(stats, chartState) {
  if (typeof Chart === 'undefined') {
    return;
  }

  const brandChartNode = document.getElementById('brandChart');
  const typeChartNode = document.getElementById('typeChart');
  const ctaChartNode = document.getElementById('ctaChart');

  if (!brandChartNode || !typeChartNode || !ctaChartNode) {
    return;
  }

  if (chartState.brand) {
    chartState.brand.destroy();
  }

  if (chartState.type) {
    chartState.type.destroy();
  }

  if (chartState.cta) {
    chartState.cta.destroy();
  }

  const brandEntries = Object.entries(stats.brands)
    .sort((left, right) => right[1] - left[1])
    .slice(0, 8);
  const ctaEntries = Object.entries(stats.ctas)
    .sort((left, right) => right[1] - left[1])
    .slice(0, 6);
  const chartTextColor = '#52525b';
  const chartGridColor = '#e4e4e7';

  chartState.brand = new Chart(brandChartNode, {
    type: 'bar',
    data: {
      labels: brandEntries.map(([label]) => label),
      datasets: [
        {
          data: brandEntries.map(([, value]) => value),
          backgroundColor: '#18181b',
          borderRadius: 10,
          maxBarThickness: 18,
        },
      ],
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
      },
      scales: {
        x: {
          beginAtZero: true,
          grid: {
            color: chartGridColor,
          },
          ticks: {
            precision: 0,
            color: chartTextColor,
          },
        },
        y: {
          grid: {
            display: false,
          },
          ticks: {
            color: chartTextColor,
          },
        },
      },
    },
  });

  chartState.type = new Chart(typeChartNode, {
    type: 'doughnut',
    data: {
      labels: ['圖片', '影片'],
      datasets: [
        {
          data: [stats.imageAds, stats.videoAds],
          backgroundColor: ['#18181b', '#d4d4d8'],
          borderWidth: 0,
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: '72%',
      plugins: {
        legend: {
          position: 'bottom',
          labels: {
            usePointStyle: true,
            padding: 18,
            color: chartTextColor,
          },
        },
      },
    },
  });

  chartState.cta = new Chart(ctaChartNode, {
    type: 'bar',
    data: {
      labels: ctaEntries.map(([label]) => label),
      datasets: [
        {
          data: ctaEntries.map(([, value]) => value),
          backgroundColor: '#71717a',
          borderRadius: 10,
          maxBarThickness: 18,
        },
      ],
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
      },
      scales: {
        x: {
          beginAtZero: true,
          grid: {
            color: chartGridColor,
          },
          ticks: {
            precision: 0,
            color: chartTextColor,
          },
        },
        y: {
          grid: {
            display: false,
          },
          ticks: {
            color: chartTextColor,
          },
        },
      },
    },
  });
}

function showOnly(stateName, elements) {
  Object.entries(elements).forEach(([name, element]) => {
    if (!element) {
      return;
    }

    element.hidden = name !== stateName;
  });
}

function updateShellLabel(currentBrandKey, source) {
  if (!source) {
    document.title = currentBrandKey ? `${currentBrandKey} · Facebook Ads Spy` : 'Facebook Ads Spy';
    return;
  }

  document.title = currentBrandKey ? `${currentBrandKey} · Facebook Ads Spy` : `${source.label} · Facebook Ads Spy`;
}

function setupDashboard() {
  const elements = {
    empty: document.getElementById('emptyState'),
    loading: document.getElementById('loadingState'),
    error: document.getElementById('errorState'),
    content: document.getElementById('mainContent'),
  };

  const refreshButton = document.getElementById('refreshBtn');
  const mainGrid = document.getElementById('mainGrid');
  const commandPanel = document.getElementById('commandPanel');
  const sidebarToggleButton = document.getElementById('sidebarToggleBtn');
  const sidebarToggleIcon = document.getElementById('sidebarToggleIcon');
  const sidebarToggleLabel = document.getElementById('sidebarToggleLabel');
  const sourceInput = document.getElementById('sourceInput');
  const sourceApplyButton = document.getElementById('sourceApplyBtn');
  const sourceStatus = document.getElementById('sourceStatus');
  const brandFilter = document.getElementById('brandFilter');
  const typeFilter = document.getElementById('typeFilter');
  const searchInput = document.getElementById('searchInput');
  const adsContainer = document.getElementById('adsContainer');
  const noResults = document.getElementById('noResults');
  const loadMoreButton = document.getElementById('loadMoreBtn');
  const loadMoreContainer = document.getElementById('loadMoreContainer');
  const expandAllButton = document.getElementById('expandAllBtn');
  const resultCount = document.getElementById('resultCount');
  const errorMessage = document.getElementById('errorMessage');
  const lastUpdate = document.getElementById('lastUpdate');
  const chartState = { brand: null, type: null, cta: null };

  let allAds = [];
  let filteredAds = [];
  let currentPage = 1;
  let allExpanded = false;
  let currentBrandKey = new URLSearchParams(window.location.search).get('brand')?.trim() || '';
  let activeSource = resolveInitialSource();
  let sidebarState = window.localStorage.getItem(SIDEBAR_STORAGE_KEY) || 'expanded';

  function setSourceStatus(message = '', isError = false) {
    if (!sourceStatus) {
      return;
    }

    sourceStatus.hidden = !message;
    sourceStatus.textContent = message;
    sourceStatus.classList.toggle('is-error', isError);
  }

  function renderSidebarState() {
    const collapsed = shouldCollapseSidebar(sidebarState, window.innerWidth);
    const expanded = !collapsed;

    if (mainGrid) {
      mainGrid.classList.toggle('is-collapsed', collapsed);
    }

    if (commandPanel) {
      commandPanel.classList.toggle('is-collapsed', collapsed);
    }

    if (sidebarToggleButton) {
      sidebarToggleButton.setAttribute('aria-expanded', String(expanded));
      sidebarToggleButton.title = expanded ? '收起篩選側欄' : '展開篩選側欄';
    }

    if (sidebarToggleIcon) {
      sidebarToggleIcon.textContent = expanded ? '‹' : '›';
    }

    if (sidebarToggleLabel) {
      sidebarToggleLabel.textContent = expanded ? '收起篩選側欄' : '展開篩選側欄';
    }
  }

  function renderStats() {
    const stats = computeStats(filteredAds);

    updateMetricText(document.getElementById('totalAds'), stats.totalAds);
    updateMetricText(document.getElementById('totalBrands'), stats.totalBrands);
    updateMetricText(document.getElementById('imageAds'), stats.imageAds);
    updateMetricText(document.getElementById('videoAds'), stats.videoAds);

    updateCharts(stats, chartState);
  }

  function renderAds(append = false) {
    const term = searchInput.value.trim();
    const start = (currentPage - 1) * ADS_PER_PAGE;
    const end = Math.min(start + ADS_PER_PAGE, filteredAds.length);
    const items = append ? filteredAds.slice(start, end) : filteredAds.slice(0, end);

    if (!append) {
      adsContainer.innerHTML = '';
    }

    items.forEach((ad, index) => {
      const realIndex = append ? start + index : index;
      adsContainer.appendChild(createAdCard(ad, realIndex, term, allExpanded));
    });

    resultCount.textContent = `${Math.min(end, filteredAds.length)} / ${filteredAds.length} creatives in view`;
    noResults.hidden = filteredAds.length !== 0;
    loadMoreContainer.hidden = end >= filteredAds.length || filteredAds.length === 0;
    expandAllButton.disabled = filteredAds.length === 0;
    expandAllButton.textContent = allExpanded ? '全部收起' : '全部展開';
  }

  function applyFilters() {
    const selectedBrand = brandFilter.value;
    const selectedType = typeFilter.value;
    const term = searchInput.value.trim().toLowerCase();

    filteredAds = allAds.filter((ad) => {
      const matchesBrand = !selectedBrand || ad.page_name === selectedBrand;
      const matchesType = !selectedType || ad.type === selectedType;
      const haystack = [
        ad.page_name,
        ad.original_feed,
        ad.platforms,
        ad.cta_text,
        ad.original_url,
      ].join(' ').toLowerCase();
      const matchesSearch = !term || haystack.includes(term);

      return matchesBrand && matchesType && matchesSearch;
    });

    currentPage = 1;
    allExpanded = false;
    renderStats();
    renderAds();
  }

  function populateFilters() {
    const brands = [...new Set(allAds.map((ad) => ad.page_name).filter(Boolean))].sort((left, right) => left.localeCompare(right));
    brandFilter.innerHTML = '<option value="">全部廣告主</option>';
    brands.forEach((brand) => {
      brandFilter.add(new Option(brand, brand));
    });

    typeFilter.innerHTML = `
      <option value="">所有素材</option>
      <option value="image">圖片</option>
      <option value="video">影片</option>
    `;
  }

  function showErrorState(message) {
    errorMessage.textContent = message;
    showOnly('error', elements);
  }

  async function loadData() {
    if (!activeSource) {
      setSourceStatus('請先輸入代號', true);
      showOnly('empty', elements);
      return;
    }

    refreshButton.disabled = true;
    if (sourceApplyButton) {
      sourceApplyButton.disabled = true;
    }
    showOnly('loading', elements);

    try {
      const response = await fetch(buildSheetUrl(activeSource));
      const bodyText = await response.text();

      if (!response.ok) {
        let message = 'Google Sheet 目前無法讀取';
        try {
          const errorPayload = JSON.parse(bodyText);
          message = errorPayload.error || message;
        } catch {
          message = bodyText || message;
        }
        throw new Error(message);
      }

      const payload = extractSheetPayload(bodyText);
      const columns = payload?.table?.cols;
      const rows = payload?.table?.rows;

      if (!Array.isArray(rows) || rows.length === 0) {
        throw new Error('工作表沒有可顯示的廣告資料');
      }

      allAds = formatAdsFromSheetRows(rows, Array.isArray(columns) ? columns : []);
      filteredAds = [...allAds];
      currentPage = 1;
      allExpanded = false;
      populateFilters();

      if (currentBrandKey) {
        brandFilter.value = [...brandFilter.options].some((option) => option.value === currentBrandKey)
          ? currentBrandKey
          : '';
      }

      applyFilters();
      showOnly('content', elements);
      if (sourceInput) {
        sourceInput.value = activeSource.code;
      }
      setSourceStatus('');
      lastUpdate.textContent = new Date().toLocaleString('zh-HK', {
        hour12: false,
      });
    } catch (error) {
      showErrorState(error.message || '數據載入失敗');
    } finally {
      refreshButton.disabled = false;
      if (sourceApplyButton) {
        sourceApplyButton.disabled = false;
      }
    }
  }

  function applySourceSelection() {
    const nextSource = resolveSheetSource(sourceInput?.value);

    if (!nextSource) {
      setSourceStatus('找不到代號', true);
      return;
    }

    activeSource = nextSource;
    updateShellLabel(currentBrandKey, activeSource);
    if (sourceInput) {
      sourceInput.value = activeSource.code;
    }
    setSourceStatus('');
    loadData();
  }

  function toggleAllAds() {
    allExpanded = !allExpanded;
    document.querySelectorAll('.ad-details-wrap').forEach((panel) => {
      panel.classList.toggle('expanded', allExpanded);
    });
    document.querySelectorAll('.detail-toggle').forEach((button) => {
      syncToggleButton(button, allExpanded);
    });
    expandAllButton.textContent = allExpanded ? '全部收起' : '全部展開';
  }

  updateShellLabel(currentBrandKey, activeSource);
  renderSidebarState();

  if (sourceInput) {
    sourceInput.value = '';
  }

  refreshButton.addEventListener('click', () => {
    loadData();
  });
  sourceApplyButton.addEventListener('click', applySourceSelection);
  sidebarToggleButton.addEventListener('click', () => {
    sidebarState = sidebarState === 'collapsed' ? 'expanded' : 'collapsed';
    window.localStorage.setItem(SIDEBAR_STORAGE_KEY, sidebarState);
    renderSidebarState();
  });
  brandFilter.addEventListener('change', applyFilters);
  typeFilter.addEventListener('change', applyFilters);
  searchInput.addEventListener('input', applyFilters);
  loadMoreButton.addEventListener('click', () => {
    currentPage += 1;
    renderAds(true);
  });
  expandAllButton.addEventListener('click', toggleAllAds);
  adsContainer.addEventListener('click', (event) => {
    const button = event.target.closest('.detail-toggle');
    if (!button) {
      return;
    }

    const panel = document.getElementById(button.dataset.target);
    const expanded = !panel.classList.contains('expanded');
    panel.classList.toggle('expanded', expanded);
    syncToggleButton(button, expanded);
  });
  window.addEventListener('resize', renderSidebarState);

  showOnly('empty', elements);
}

if (typeof document !== 'undefined') {
  setupDashboard();
}
