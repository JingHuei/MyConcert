/* =========================================================
   Lucy 的演唱會紀錄網站
   功能：
   1) 讀取公開 Google Sheet JSON (OpenSheet)
   2) Google 登入後寫入 Sheets API
   3) 搜尋 / 篩選 / 排序 / 分頁
   4) Leaflet 地圖聚合城市標記
   5) 新增表單草稿自動保存
   ========================================================= */

/* =========================
   0. 設定區
   ========================= */
const CONFIG = {
  GOOGLE_CLIENT_ID: 'PASTE_YOUR_GOOGLE_OAUTH_CLIENT_ID',
  SPREADSHEET_ID: 'PASTE_YOUR_SPREADSHEET_ID',
  RECORD_SHEET_NAME: '演唱會紀錄',
  FIELDS_SHEET_NAME: '欄位表',
  OPEN_SHEET_BASE: 'https://opensheet.elk.sh',
  SHEETS_API_BASE: 'https://sheets.googleapis.com/v4/spreadsheets',
  GOOGLE_SCOPES: 'https://www.googleapis.com/auth/spreadsheets',
  PAGE_SIZE: 10,
  DEFAULT_CENTER: [25.0330, 121.5654], // 雙北
  DEFAULT_ZOOM: 10,
};

const STORAGE_KEYS = {
  AUTH: 'lucy_concert_auth_v1',
  DRAFT: 'lucy_concert_draft_v1',
  GEO: 'lucy_concert_geo_cache_v1',
};

const CITY_QUERY_ALIASES = {
  '台灣台北': 'Taipei, Taiwan',
  '台北': 'Taipei, Taiwan',
  '台北市': 'Taipei, Taiwan',
  '新北': 'New Taipei City, Taiwan',
  '新北市': 'New Taipei City, Taiwan',
  '台灣桃園': 'Taoyuan, Taiwan',
  '桃園': 'Taoyuan, Taiwan',
  '桃園市': 'Taoyuan, Taiwan',
  '台灣高雄': 'Kaohsiung, Taiwan',
  '高雄': 'Kaohsiung, Taiwan',
  '高雄市': 'Kaohsiung, Taiwan',
  '韓國首爾': 'Seoul, South Korea',
  '首爾': 'Seoul, South Korea',
  '東京': 'Tokyo, Japan',
  '東京都': 'Tokyo, Japan',
};

const CITY_FALLBACK_COORDS = {
  '台灣台北': [25.0330, 121.5654],
  '台北': [25.0330, 121.5654],
  '台北市': [25.0330, 121.5654],
  '新北': [25.0169, 121.4628],
  '新北市': [25.0169, 121.4628],
  '台灣桃園': [24.9936, 121.3010],
  '桃園': [24.9936, 121.3010],
  '桃園市': [24.9936, 121.3010],
  '台灣高雄': [22.6273, 120.3014],
  '高雄': [22.6273, 120.3014],
  '高雄市': [22.6273, 120.3014],
  '韓國首爾': [37.5665, 126.9780],
  '首爾': [37.5665, 126.9780],
  '東京': [35.6762, 139.6503],
  '東京都': [35.6762, 139.6503],
};

const state = {
  records: [],
  fieldRows: [],
  filteredRecords: [],
  currentPage: 1,
  pageSize: CONFIG.PAGE_SIZE,
  map: null,
  markersLayer: null,
  cityMarkers: new Map(),
  accessToken: null,
  tokenExpiry: 0,
  tokenClient: null,
  currentCityFilter: '',
  currentYearFilter: '',
  initialized: false,
};

const els = {};

/* =========================
   1. DOM / 初始化
   ========================= */
document.addEventListener('DOMContentLoaded', init);

function init() {
  cacheElements();
  bindEvents();
  restoreAuth();
  restoreDraft();
  initModalState();
  initMap();
  loadAllData();
  setupScrollFadeIn();
}

function cacheElements() {
  const ids = [
    'authStatus', 'loginBtn', 'searchInput', 'sortSelect', 'filterType', 'filterCity', 'filterArtist',
    'filterStartDate', 'filterEndDate', 'yearSelect', 'resetFiltersBtn', 'resultInfo', 'metricTotal',
    'metricTopArtist', 'metricTopCity', 'statTopArtist', 'statTopArtistCount', 'statTopCity',
    'statTopCityCount', 'statYearCount', 'statFavoriteCount', 'loadingState', 'mapHint', 'cardGrid',
    'emptyState', 'pageIndicator', 'prevPageBtn', 'nextPageBtn', 'openFormBtn', 'recordModal',
    'closeFormBtn', 'recordForm', 'clearFormBtn', 'formMessage', 'formType', 'formConcertName',
    'formArtist', 'formDate', 'formCountry', 'formLocation', 'formPrice', 'formSeat', 'formPartner',
    'formImgUrlS', 'formImgUrlM', 'formNote', 'formFavorite', 'submitFormBtn'
  ];
  ids.forEach(id => els[id] = document.getElementById(id));
}

/* =========================
   2. 事件綁定
   ========================= */
function bindEvents() {
  els.loginBtn.addEventListener('click', handleLoginClick);
  els.searchInput.addEventListener('input', handleFilterChange);
  els.sortSelect.addEventListener('change', handleFilterChange);
  els.filterType.addEventListener('change', handleFilterChange);
  els.filterCity.addEventListener('change', handleCityFilterChange);
  els.filterArtist.addEventListener('change', handleFilterChange);
  els.filterStartDate.addEventListener('change', handleFilterChange);
  els.filterEndDate.addEventListener('change', handleFilterChange);
  els.yearSelect.addEventListener('change', handleYearChange);

  els.resetFiltersBtn.addEventListener('click', resetFilters);
  els.prevPageBtn.addEventListener('click', () => changePage(-1));
  els.nextPageBtn.addEventListener('click', () => changePage(1));

  els.openFormBtn.addEventListener('click', handleOpenForm);
  els.closeFormBtn.addEventListener('click', closeModal);
  els.recordModal.addEventListener('click', (e) => {
    if (e.target.classList.contains('record-modal__backdrop')) closeModal();
  });

  els.recordForm.addEventListener('input', saveDraft);
  els.recordForm.addEventListener('change', saveDraft);
  els.recordForm.addEventListener('submit', submitForm);
  els.clearFormBtn.addEventListener('click', clearForm);
}

/* =========================
   3. 資料讀取
   ========================= */
async function loadAllData() {
  setLoading(true, '正在讀取公開資料…');
  try {
    const [recordsRes, fieldsRes] = await Promise.all([
      axios.get(buildOpenSheetUrl(CONFIG.RECORD_SHEET_NAME)),
      axios.get(buildOpenSheetUrl(CONFIG.FIELDS_SHEET_NAME)),
    ]);

    state.records = normalizeRecords(recordsRes.data || []);
    state.fieldRows = Array.isArray(fieldsRes.data) ? fieldsRes.data : [];

    initFiltersFromData();
    updateStats();
    await buildMapMarkers();
    applyAll();
    els.loadingState.classList.add('d-none');
    state.initialized = true;
  } catch (error) {
    console.error(error);
    els.loadingState.innerHTML = `
      <div class="fw-bold mb-2">載入失敗</div>
      <div class="text-secondary small">請確認 Spreadsheet ID、工作表名稱與公開權限是否正確。</div>
    `;
    els.loadingState.classList.remove('d-none');
    setStatus('讀取失敗', 'danger');
  }
}

function buildOpenSheetUrl(sheetName) {
  const sid = CONFIG.SPREADSHEET_ID.trim();
  const tab = encodeURIComponent(sheetName);
  return `${CONFIG.OPEN_SHEET_BASE}/${sid}/${tab}`;
}

function normalizeRecords(rows) {
  return rows.map((row, index) => {
    const id = row.id || row.ID || row.Id || `${Date.now()}-${index}`;
    const date = parseDate(row.Date);
    return {
      id,
      type: (row.type || '').trim(),
      ConcertName: (row.ConcertName || '').trim(),
      Artist: (row.Artist || '').trim(),
      Country: (row.Country || '').trim(),
      Location: (row.Location || '').trim(),
      Date: date,
      Price: row.Price ?? '',
      Seat: (row.Seat || '').trim(),
      imgUrlS: (row.imgUrlS || '').trim(),
      imgUrlM: (row.imgUrlM || '').trim(),
      note: (row.note || '').trim(),
      partner: (row.partner || '').trim(),
      favorite: normalizeBoolean(row.favorite),
      _rawDate: row.Date || '',
    };
  }).filter(row => row.ConcertName || row.Artist || row.Country || row.Location || row.Date);
}

function parseDate(value) {
  if (!value) return null;
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeBoolean(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'number') return value === 1;
  if (typeof value === 'string') {
    const v = value.toLowerCase().trim();
    return ['true', 'yes', 'y', '1', '是', '愛', '❤️'].includes(v);
  }
  return false;
}

/* =========================
   4. 篩選 / 排序 / 分頁
   ========================= */
function initFiltersFromData() {
  const types = uniqFromData(state.fieldRows.map(r => r.Type || r.type || '').filter(Boolean));
  const cities = uniqFromData(state.fieldRows.map(r => r.Country || r.country || '').filter(Boolean));
  const artists = uniqFromData(state.records.map(r => r.Artist).filter(Boolean));

  fillSelect(els.filterType, types);
  fillSelect(els.filterCity, cities);
  fillSelect(els.filterArtist, artists, true);
  fillSelect(els.formType, types);
  fillSelect(els.formCountry, cities);
  fillSelect(els.formLocation, uniqFromData(state.fieldRows.map(r => r.Location || r.location || '').filter(Boolean)));

  fillYearSelect();
}

function uniqFromData(arr) {
  return [...new Set(arr.map(s => String(s).trim()).filter(Boolean))];
}

function fillSelect(selectEl, values, sortAlpha = false) {
  const current = selectEl.value;
  const items = sortAlpha ? [...values].sort((a, b) => a.localeCompare(b, 'zh-Hant-TW')) : values;
  const baseOption = selectEl.querySelector('option[value=""]');
  selectEl.innerHTML = '';
  if (baseOption) selectEl.appendChild(baseOption);

  items.forEach(value => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = value;
    selectEl.appendChild(opt);
  });

  if (items.includes(current)) selectEl.value = current;
}

function fillYearSelect() {
  const years = uniqFromData(state.records.map(r => r.Date && String(r.Date.getFullYear())).filter(Boolean))
    .sort((a, b) => Number(b) - Number(a));

  const current = els.yearSelect.value;
  els.yearSelect.innerHTML = '<option value="">全部年份</option>';
  years.forEach(year => {
    const opt = document.createElement('option');
    opt.value = year;
    opt.textContent = year;
    els.yearSelect.appendChild(opt);
  });
  if (years.includes(current)) els.yearSelect.value = current;
}

function getFilteredRecords() {
  const keyword = els.searchInput.value.trim().toLowerCase();
  const type = els.filterType.value;
  const city = els.filterCity.value;
  const artist = els.filterArtist.value;
  const startDate = els.filterStartDate.value ? new Date(`${els.filterStartDate.value}T00:00:00`) : null;
  const endDate = els.filterEndDate.value ? new Date(`${els.filterEndDate.value}T23:59:59`) : null;
  const year = els.yearSelect.value;

  let result = [...state.records].filter(record => {
    const textBlob = [
      record.ConcertName,
      record.Artist,
      record.Country,
      record.Location,
      record.Seat,
      record.note,
      record.partner,
      record.type,
    ].join(' ').toLowerCase();

    const dateOK = (!startDate || !record.Date || record.Date >= startDate)
      && (!endDate || !record.Date || record.Date <= endDate)
      && (!year || !record.Date || String(record.Date.getFullYear()) === String(year));

    return (!keyword || textBlob.includes(keyword))
      && (!type || record.type === type)
      && (!city || record.Country === city)
      && (!artist || record.Artist === artist)
      && dateOK;
  });

  result = sortRecords(result, els.sortSelect.value);
  return result;
}

function sortRecords(records, mode) {
  const out = [...records];
  switch (mode) {
    case 'date-asc':
      out.sort((a, b) => (a.Date?.getTime() || 0) - (b.Date?.getTime() || 0));
      break;
    case 'date-desc':
      out.sort((a, b) => (b.Date?.getTime() || 0) - (a.Date?.getTime() || 0));
      break;
    case 'artist-stroke-asc':
      out.sort(byStrokeThenDate('Artist', false));
      break;
    case 'artist-stroke-desc':
      out.sort(byStrokeThenDate('Artist', true));
      break;
    case 'city-stroke-asc':
      out.sort(byStrokeThenDate('Country', false));
      break;
    case 'city-stroke-desc':
      out.sort(byStrokeThenDate('Country', true));
      break;
    default:
      out.sort((a, b) => (b.Date?.getTime() || 0) - (a.Date?.getTime() || 0));
  }
  return out;
}

function byStrokeThenDate(field, desc = false) {
  const items = [...new Set(state.records.map(r => r[field]).filter(Boolean))];
  const ordered = sortByStroke(items, desc);
  const rank = new Map(ordered.map((v, i) => [v, i]));
  return (a, b) => {
    const ra = rank.get(a[field]) ?? Number.MAX_SAFE_INTEGER;
    const rb = rank.get(b[field]) ?? Number.MAX_SAFE_INTEGER;
    if (ra !== rb) return ra - rb;
    return (b.Date?.getTime() || 0) - (a.Date?.getTime() || 0);
  };
}

function sortByStroke(strings, desc = false) {
  let sorted = [...strings];
  try {
    if (window.cnchar && typeof window.cnchar.sortStroke === 'function') {
      sorted = window.cnchar.sortStroke([...strings]);
    } else if (typeof ''.sortStroke === 'function') {
      sorted = [...strings].sortStroke();
    } else {
      sorted = [...strings].sort((a, b) => a.localeCompare(b, 'zh-Hant-TW'));
    }
  } catch (err) {
    sorted = [...strings].sort((a, b) => a.localeCompare(b, 'zh-Hant-TW'));
  }
  return desc ? sorted.reverse() : sorted;
}

function applyAll() {
  state.filteredRecords = getFilteredRecords();
  state.currentPage = 1;
  renderEverything();
}

function renderEverything() {
  renderStats();
  renderCards();
  renderPagination();
  renderMapHighlight();
  updateResultInfo();
}

function renderCards() {
  const records = getPagedRecords();
  els.cardGrid.innerHTML = '';

  if (!records.length) {
    els.emptyState.classList.remove('d-none');
    return;
  }
  els.emptyState.classList.add('d-none');

  records.forEach(record => {
    const card = document.createElement('article');
    card.className = 'concert-card fade-in';
    card.dataset.city = record.Country || '';
    card.innerHTML = `
      <div class="position-relative">
        ${record.favorite ? '<div class="heart"><i class="fa-solid fa-heart"></i></div>' : ''}
        ${renderPicture(record)}
      </div>
      <div class="card-body">
        <div class="card-badges">
          <span class="pill"><i class="fa-solid fa-calendar-day"></i>${formatDate(record.Date)}</span>
          <span class="pill"><i class="fa-solid fa-location-dot"></i>${escapeHtml(record.Country || '-')}</span>
        </div>
        <div class="card-title">${escapeHtml(record.ConcertName || '-')}</div>
        <div class="card-artist">${escapeHtml(record.Artist || '-')}</div>
        <div class="card-meta">
          <div>${escapeHtml(record.Country || '-')} · ${escapeHtml(record.Location || '-')}</div>
          <div class="card-small">座位：${escapeHtml(record.Seat || '-')}｜票價：${formatPrice(record.Price)}</div>
          <div class="card-small">${daysSinceText(record.Date)}</div>
          ${record.partner ? `<div class="card-small">夥伴：${escapeHtml(record.partner)}</div>` : ''}
          ${record.note ? `<div class="card-small text-truncate-2">心得：${escapeHtml(record.note)}</div>` : ''}
        </div>
      </div>
    `;
    els.cardGrid.appendChild(card);
  });

  observeFadeIn();
}

function renderPicture(record) {
  const fallback = 'https://images.unsplash.com/photo-1516280440614-37939bbacd81?auto=format&fit=crop&w=1200&q=80';
  const srcM = record.imgUrlM || record.imgUrlS || fallback;
  const srcS = record.imgUrlS || record.imgUrlM || fallback;
  return `
    <picture>
      <source media="(min-width: 768px)" srcset="${escapeAttr(srcM)}">
      <img class="card-image" src="${escapeAttr(srcS)}" alt="${escapeAttr(record.ConcertName || '演唱會圖片')}" loading="lazy">
    </picture>
  `;
}

function getPagedRecords() {
  const start = (state.currentPage - 1) * state.pageSize;
  return state.filteredRecords.slice(start, start + state.pageSize);
}

function renderPagination() {
  const totalPages = Math.max(1, Math.ceil(state.filteredRecords.length / state.pageSize));
  state.currentPage = Math.min(state.currentPage, totalPages);
  els.pageIndicator.textContent = `第 ${state.currentPage} 頁 / 共 ${totalPages} 頁`;
  els.prevPageBtn.disabled = state.currentPage <= 1;
  els.nextPageBtn.disabled = state.currentPage >= totalPages;
}

function changePage(delta) {
  const totalPages = Math.max(1, Math.ceil(state.filteredRecords.length / state.pageSize));
  state.currentPage = Math.min(totalPages, Math.max(1, state.currentPage + delta));
  renderEverything();
  window.scrollTo({ top: document.querySelector('.card-grid').offsetTop - 120, behavior: 'smooth' });
}

function updateResultInfo() {
  const total = state.filteredRecords.length;
  const pageStart = total ? ((state.currentPage - 1) * state.pageSize + 1) : 0;
  const pageEnd = Math.min(total, state.currentPage * state.pageSize);
  els.resultInfo.textContent = total
    ? `共 ${total} 場，顯示第 ${pageStart}–${pageEnd} 場`
    : '沒有符合條件的資料';
}

function resetFilters() {
  els.searchInput.value = '';
  els.sortSelect.value = 'date-desc';
  els.filterType.value = '';
  els.filterCity.value = '';
  els.filterArtist.value = '';
  els.filterStartDate.value = '';
  els.filterEndDate.value = '';
  els.yearSelect.value = '';
  state.currentCityFilter = '';
  state.currentYearFilter = '';
  applyAll();
  renderMapHighlight();
}

function handleFilterChange() {
  state.currentCityFilter = '';
  applyAll();
  renderMapHighlight();
}

function handleCityFilterChange() {
  state.currentCityFilter = els.filterCity.value;
  applyAll();
  renderMapHighlight();
}

function handleYearChange() {
  state.currentYearFilter = els.yearSelect.value;
  applyAll();
}

/* =========================
   5. 統計
   ========================= */
function renderStats() {
  const artistCounts = countBy(state.records, 'Artist');
  const cityCounts = countBy(state.records, 'Country');
  const favoriteCount = state.records.filter(r => r.favorite).length;
  const year = els.yearSelect.value;
  const yearCount = year ? state.records.filter(r => r.Date && String(r.Date.getFullYear()) === String(year)).length : state.records.length;

  const topArtist = topEntry(artistCounts);
  const topCity = topEntry(cityCounts);

  els.metricTotal.textContent = state.records.length;
  els.metricTopArtist.textContent = topArtist.key || '-';
  els.metricTopCity.textContent = topCity.key || '-';

  els.statTopArtist.textContent = topArtist.key || '-';
  els.statTopArtistCount.textContent = `${topArtist.count || 0} 場`;
  els.statTopCity.textContent = topCity.key || '-';
  els.statTopCityCount.textContent = `${topCity.count || 0} 場`;
  els.statYearCount.textContent = yearCount;
  els.statFavoriteCount.textContent = favoriteCount;
}

function updateStats() {
  renderStats();
}

function countBy(arr, field) {
  return arr.reduce((acc, item) => {
    const key = item[field] || '-';
    acc[key] = (acc[key] || 0) + 1;
    return acc;
  }, {});
}

function topEntry(countMap) {
  const entries = Object.entries(countMap);
  if (!entries.length) return { key: '-', count: 0 };
  entries.sort((a, b) => b[1] - a[1]);
  return { key: entries[0][0], count: entries[0][1] };
}

/* =========================
   6. 地圖
   ========================= */
function initMap() {
  state.map = L.map('map', {
    zoomControl: true,
    scrollWheelZoom: false,
  }).setView(CONFIG.DEFAULT_CENTER, CONFIG.DEFAULT_ZOOM);

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap contributors',
  }).addTo(state.map);

  state.markersLayer = L.layerGroup().addTo(state.map);

  els.mapHint.textContent = '預設視野：雙北';
}

async function buildMapMarkers() {
  state.cityMarkers.clear();
  state.markersLayer.clearLayers();

  const cityGroups = groupByCity(state.records);

  const bounds = [];
  for (const [city, records] of cityGroups.entries()) {
    const coords = await geocodeCity(city);
    if (!coords) continue;

    const marker = L.marker(coords, {
      icon: L.divIcon({
        className: '',
        html: `
          <div class="city-badge">
            <strong>${records.length}</strong>
            <span>場</span>
          </div>
        `,
        iconSize: [44, 44],
        iconAnchor: [22, 22],
      }),
      title: city,
    }).addTo(state.markersLayer);

    marker.bindPopup(buildMarkerPopup(city, records), { maxWidth: 280 });
    marker.on('click', () => {
      els.filterCity.value = city;
      state.currentCityFilter = city;
      applyAll();
    });

    state.cityMarkers.set(city, { marker, records, coords });
    bounds.push(coords);
  }

  if (bounds.length > 0) {
    state.map.fitBounds(bounds, { padding: [30, 30] });
    els.mapHint.textContent = `已標出 ${bounds.length} 個城市`;
  } else {
    state.map.setView(CONFIG.DEFAULT_CENTER, CONFIG.DEFAULT_ZOOM);
    els.mapHint.textContent = '尚未取得城市座標，顯示雙北預設視野';
  }
}

function renderMapHighlight() {
  if (!state.map) return;

  const activeCity = els.filterCity.value || state.currentCityFilter;
  state.cityMarkers.forEach(({ marker, records }, city) => {
    const el = marker.getElement();
    if (!el) return;
    const inner = el.querySelector('.city-badge');
    if (!inner) return;
    inner.style.filter = activeCity && activeCity !== city ? 'grayscale(.45) opacity(.55)' : 'none';
    inner.style.transform = activeCity === city ? 'scale(1.08)' : 'scale(1)';
  });
}

function groupByCity(records) {
  const map = new Map();
  records.forEach(record => {
    if (!record.Country) return;
    if (!map.has(record.Country)) map.set(record.Country, []);
    map.get(record.Country).push(record);
  });
  return map;
}

function buildMarkerPopup(city, records) {
  const list = records.slice(0, 5).map(r => `
    <li>
      <strong>${escapeHtml(r.ConcertName || '-')}</strong><br>
      <span class="text-secondary">${escapeHtml(formatDate(r.Date))} · ${escapeHtml(r.Location || '-')}</span>
    </li>
  `).join('');
  return `
    <div class="popup-content">
      <div class="fw-bold mb-1">${escapeHtml(city)}</div>
      <div class="small text-secondary mb-2">共 ${records.length} 場</div>
      <ul class="mb-0 ps-3">${list}</ul>
    </div>
  `;
}

async function geocodeCity(city) {
  if (!city) return null;

  const cache = readJsonStorage(STORAGE_KEYS.GEO, {});
  if (cache[city]) return cache[city];

  if (CITY_FALLBACK_COORDS[city]) {
    cache[city] = CITY_FALLBACK_COORDS[city];
    writeJsonStorage(STORAGE_KEYS.GEO, cache);
    return cache[city];
  }

  const query = CITY_QUERY_ALIASES[city] || city;
  try {
    const res = await axios.get('https://nominatim.openstreetmap.org/search', {
      params: {
        q: query,
        format: 'jsonv2',
        limit: 1,
      },
      headers: {
        'Accept-Language': 'zh-TW',
      },
    });

    if (Array.isArray(res.data) && res.data[0]) {
      const coords = [Number(res.data[0].lat), Number(res.data[0].lon)];
      cache[city] = coords;
      writeJsonStorage(STORAGE_KEYS.GEO, cache);
      return coords;
    }
  } catch (err) {
    console.warn('Geocoding failed:', city, err);
  }

  return CITY_FALLBACK_COORDS[city] || null;
}

/* =========================
   7. Google 登入 / 寫入
   ========================= */
function restoreAuth() {
  const saved = readJsonStorage(STORAGE_KEYS.AUTH, null);
  if (saved && saved.accessToken && Date.now() < saved.expiresAt - 60_000) {
    state.accessToken = saved.accessToken;
    state.tokenExpiry = saved.expiresAt;
    setStatus('已登入', 'success');
  } else {
    clearAuth();
  }
}

function saveAuth(accessToken, expiresInSeconds) {
  const expiresAt = Date.now() + (Number(expiresInSeconds || 3600) * 1000);
  state.accessToken = accessToken;
  state.tokenExpiry = expiresAt;
  writeJsonStorage(STORAGE_KEYS.AUTH, { accessToken, expiresAt });
  setStatus('已登入', 'success');
}

function clearAuth() {
  state.accessToken = null;
  state.tokenExpiry = 0;
  localStorage.removeItem(STORAGE_KEYS.AUTH);
  setStatus('未登入', 'secondary');
}

function isTokenValid() {
  return Boolean(state.accessToken) && Date.now() < state.tokenExpiry - 60_000;
}

function handleLoginClick() {
  if (isTokenValid()) {
    showToast('目前已登入，可直接新增資料。', 'success');
    return;
  }
  requestLogin();
}

function requestLogin() {
  if (!window.google || !google.accounts || !google.accounts.oauth2) {
    showToast('Google 登入模組尚未載入。', 'danger');
    return;
  }

  if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.includes('PASTE_')) {
    showToast('請先填入 Google OAuth Client ID。', 'warning');
    return;
  }

  if (!state.tokenClient) {
    state.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.GOOGLE_CLIENT_ID,
      scope: CONFIG.GOOGLE_SCOPES,
      callback: (response) => {
        if (response.error) {
          console.error(response);
          showToast('登入失敗，請再試一次。', 'danger');
          return;
        }
        saveAuth(response.access_token, response.expires_in || 3600);
        showToast('登入成功。', 'success');
      },
    });
  }

  state.tokenClient.callback = (response) => {
    if (response.error) {
      console.error(response);
      showToast('登入失敗，請再試一次。', 'danger');
      return;
    }
    saveAuth(response.access_token, response.expires_in || 3600);
    showToast('登入成功。', 'success');
  };

  state.tokenClient.requestAccessToken({ prompt: 'consent' });
}

async function handleOpenForm() {
  if (!isTokenValid()) {
    requestLogin();
    const waitForToken = setInterval(() => {
      if (isTokenValid()) {
        clearInterval(waitForToken);
        openModal();
      }
    }, 400);
    setTimeout(() => clearInterval(waitForToken), 120000);
    return;
  }

  openModal();
}

function openModal() {
  els.recordModal.classList.remove('d-none');
  els.recordModal.setAttribute('aria-hidden', 'false');
  restoreDraft();
  setFormMessage('可直接填寫資料，關閉時會保留草稿。', 'secondary');
}

function closeModal() {
  els.recordModal.classList.add('d-none');
  els.recordModal.setAttribute('aria-hidden', 'true');
  saveDraft();
}

function initModalState() {
  clearFormFields(false);
}

function submitForm(event) {
  event.preventDefault();

  if (!isTokenValid()) {
    showToast('請先登入 Google。', 'warning');
    requestLogin();
    return;
  }

  const payload = collectFormData();
  const validation = validatePayload(payload);
  if (!validation.ok) {
    setFormMessage(validation.message, 'danger');
    return;
  }

  setFormMessage('送出中…', 'secondary');
  setFormLoading(true);

  axios.post(
    `${CONFIG.SHEETS_API_BASE}/${encodeURIComponent(CONFIG.SPREADSHEET_ID)}/values/${encodeURIComponent(CONFIG.RECORD_SHEET_NAME)}!A:N:append`,
    {
      values: [[
        payload.id,
        payload.type,
        payload.ConcertName,
        payload.Artist,
        payload.Country,
        payload.Location,
        payload.Date,
        payload.Price,
        payload.Seat,
        payload.imgUrlS,
        payload.imgUrlM,
        payload.note,
        payload.partner,
        payload.favorite,
      ]],
    },
    {
      params: {
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        includeValuesInResponse: 'true',
      },
      headers: {
        Authorization: `Bearer ${state.accessToken}`,
        'Content-Type': 'application/json',
      },
    }
  ).then(async () => {
    clearDraft();
    clearFormFields(true);
    closeModal();
    showToast('已成功新增演唱會紀錄。', 'success');
    await loadAllData();
  }).catch((error) => {
    console.error(error);
    setFormMessage('送出失敗，請確認試算表權限與欄位名稱。', 'danger');
    showToast('送出失敗。', 'danger');
  }).finally(() => {
    setFormLoading(false);
  });
}

function collectFormData() {
  const formData = new FormData(els.recordForm);
  return {
    id: String(Date.now()),
    type: String(formData.get('type') || '').trim(),
    ConcertName: String(formData.get('ConcertName') || '').trim(),
    Artist: String(formData.get('Artist') || '').trim(),
    Country: String(formData.get('Country') || '').trim(),
    Location: String(formData.get('Location') || '').trim(),
    Date: String(formData.get('Date') || '').trim(),
    Price: String(formData.get('Price') || '').trim(),
    Seat: String(formData.get('Seat') || '').trim(),
    imgUrlS: String(formData.get('imgUrlS') || '').trim(),
    imgUrlM: String(formData.get('imgUrlM') || '').trim(),
    note: String(formData.get('note') || '').trim(),
    partner: String(formData.get('partner') || '').trim(),
    favorite: els.formFavorite.checked ? 'TRUE' : 'FALSE',
  };
}

function validatePayload(payload) {
  if (!payload.type) return { ok: false, message: '請選擇類型。' };
  if (!payload.ConcertName) return { ok: false, message: '請輸入演唱會名稱。' };
  if (!payload.Artist) return { ok: false, message: '請輸入歌手 / 團體。' };
  if (!payload.Country) return { ok: false, message: '請選擇城市。' };
  if (!payload.Location) return { ok: false, message: '請選擇場館。' };
  if (!payload.Date) return { ok: false, message: '請選擇日期。' };
  return { ok: true };
}

function clearFormFields(keepDraft = false) {
  els.recordForm.reset();
  els.formFavorite.checked = false;
  els.formMessage.textContent = '尚未送出。';
  if (!keepDraft) clearDraft();
  saveDraft();
}

function clearForm() {
  clearFormFields(false);
  setFormMessage('已清空所有輸入內容。', 'secondary');
}

function setFormLoading(isLoading) {
  els.submitFormBtn.disabled = isLoading;
  els.clearFormBtn.disabled = isLoading;
  els.submitFormBtn.innerHTML = isLoading
    ? '<span class="spinner-border spinner-border-sm me-2" aria-hidden="true"></span>送出中'
    : '<i class="fa-solid fa-paper-plane me-1"></i>送出';
}

function setFormMessage(message, tone = 'secondary') {
  els.formMessage.className = `small text-${tone}`;
  els.formMessage.textContent = message;
}

/* =========================
   8. 草稿保存
   ========================= */
function saveDraft() {
  const draft = {
    type: els.formType.value,
    ConcertName: els.formConcertName.value,
    Artist: els.formArtist.value,
    Date: els.formDate.value,
    Country: els.formCountry.value,
    Location: els.formLocation.value,
    Price: els.formPrice.value,
    Seat: els.formSeat.value,
    partner: els.formPartner.value,
    imgUrlS: els.formImgUrlS.value,
    imgUrlM: els.formImgUrlM.value,
    note: els.formNote.value,
    favorite: els.formFavorite.checked,
  };
  writeJsonStorage(STORAGE_KEYS.DRAFT, draft);
}

function restoreDraft() {
  const draft = readJsonStorage(STORAGE_KEYS.DRAFT, null);
  if (!draft) return;

  els.formType.value = draft.type || '';
  els.formConcertName.value = draft.ConcertName || '';
  els.formArtist.value = draft.Artist || '';
  els.formDate.value = draft.Date || '';
  els.formCountry.value = draft.Country || '';
  els.formLocation.value = draft.Location || '';
  els.formPrice.value = draft.Price || '';
  els.formSeat.value = draft.Seat || '';
  els.formPartner.value = draft.partner || '';
  els.formImgUrlS.value = draft.imgUrlS || '';
  els.formImgUrlM.value = draft.imgUrlM || '';
  els.formNote.value = draft.note || '';
  els.formFavorite.checked = Boolean(draft.favorite);
}

function clearDraft() {
  localStorage.removeItem(STORAGE_KEYS.DRAFT);
}

/* =========================
   9. UI / 小工具
   ========================= */
function setLoading(isLoading, text = '載入中…') {
  els.loadingState.classList.toggle('d-none', !isLoading);
  if (isLoading) {
    els.loadingState.querySelector('.fw-semibold').textContent = text;
  }
}

function setStatus(text, tone = 'secondary') {
  els.authStatus.textContent = text;
  els.authStatus.className = `badge rounded-pill text-bg-${tone} auth-badge`;
}

function showToast(message, tone = 'secondary') {
  setFormMessage(message, tone);
}

function formatDate(date) {
  if (!date) return '-';
  try {
    return new Intl.DateTimeFormat('zh-TW', {
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    }).format(date);
  } catch {
    return date.toISOString().slice(0, 10);
  }
}

function formatPrice(value) {
  if (value === '' || value === null || value === undefined) return '-';
  const num = Number(value);
  if (Number.isNaN(num)) return String(value);
  return new Intl.NumberFormat('zh-TW').format(num);
}

function daysSinceText(date) {
  if (!date) return '日期未填';
  const diff = Math.floor((new Date().setHours(0,0,0,0) - new Date(date).setHours(0,0,0,0)) / 86400000);
  if (diff >= 0) return `距離今天已過 ${diff} 天`;
  return `距離今天還有 ${Math.abs(diff)} 天`;
}

function handleYearChange() {
  applyAll();
  renderMapHighlight();
}

function escapeHtml(str) {
  return String(str ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function escapeAttr(str) {
  return escapeHtml(str).replaceAll('\n', ' ');
}

function writeJsonStorage(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function readJsonStorage(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) return fallback;
    return JSON.parse(raw);
  } catch {
    return fallback;
  }
}

/* =========================
   10. 動畫
   ========================= */
function setupScrollFadeIn() {
  observeFadeIn();
}

function observeFadeIn() {
  const items = document.querySelectorAll('.concert-card.fade-in');
  if (!('IntersectionObserver' in window)) {
    items.forEach(item => item.classList.add('in-view'));
    return;
  }

  const io = new IntersectionObserver((entries, observer) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('in-view');
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.15 });

  items.forEach(item => io.observe(item));
}

/* =========================
   11. 防呆
   ========================= */
window.addEventListener('beforeunload', () => {
  saveDraft();
});
