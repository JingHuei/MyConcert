/* =========================================================
   Lucy Concert Diary — script.js
   - 公開 OpenSheet JSON 讀取
   - Google GIS OAuth 登入 + Sheets API append
   - Leaflet 地圖 / 統計 / 篩選 / 分頁 / 草稿保存
   - 地圖 Overlay：清單 → 詳細資訊
   ========================================================= */

/* =========================
   1) 設定集中管理
   ========================= */
const CONFIG = {
  GOOGLE_CLIENT_ID: '562464022417-8c2sckejaft6de7ch8kejqomm1fbi0ga.apps.googleusercontent.com',
  SPREADSHEET_ID: '1tfNjim8BbKmv3KVNIQrRvx2T6W8YtfjyJCr1_wXDcrA',
  SHEET_NAME: '演唱會紀錄',
  PAGE_SIZE: 10,
  STORAGE_KEYS: {
    draft: 'lucyConcertDraftV1',
    token: 'lucyConcertTokenV1',
    geocode: 'lucyConcertGeocodeCacheV1',
  },
  DEFAULT_MAP_CENTER: [25.033964, 121.564468],
  DEFAULT_MAP_ZOOM: 10.5,
};

CONFIG.OPEN_SHEET_URL = `https://opensheet.elk.sh/${CONFIG.SPREADSHEET_ID}/${encodeURIComponent(CONFIG.SHEET_NAME)}`;
CONFIG.APPEND_RANGE = `${CONFIG.SHEET_NAME}!A:N`;

const DEFAULT_COUNTRY_OPTIONS = [
  '台北', '新北', '桃園', '台中', '台南', '高雄',
  '首爾', '東京', '大阪', '名古屋', '香港', '新加坡', '曼谷', '吉隆坡'
];

const DEFAULT_TYPE_OPTIONS = ['演唱會', '拼盤', 'FanMeeting', 'FanConcert'];

const DEFAULT_CITY_COORDS = {
  '台北': [25.033964, 121.564468],
  '新北': [25.0143, 121.4672],
  '桃園': [24.9936, 121.3010],
  '台中': [24.1477, 120.6736],
  '台南': [22.9999, 120.2270],
  '高雄': [22.6273, 120.3014],
  '首爾': [37.5665, 126.9780],
  '東京': [35.6762, 139.6503],
  '大阪': [34.6937, 135.5023],
  '名古屋': [35.1815, 136.9066],
  '香港': [22.3193, 114.1694],
  '新加坡': [1.3521, 103.8198],
  '曼谷': [13.7563, 100.5018],
  '吉隆坡': [3.1390, 101.6869],
};

const DEFAULT_LOCATION_OPTIONS_BY_COUNTRY = {
  '台北': ['台北大巨蛋','台北小巨蛋', '台北流行音樂中心', '台大體育館', 'Legacy Taipei', '台北國際會議中心', '華山1914文化創意產業園區'],
  '新北': ['新莊體育館', '林口體育館'],
  '桃園': ['桃園會展中心', '桃園巨蛋'],
  '高雄': ['高雄巨蛋', '高雄流行音樂中心' , '夢時代對面廣場'],
  '首爾': ['KSPO DOME', 'Olympic Hall', 'Jamsil Arena'],
  '東京': ['Tokyo Dome', 'Buddokan', 'Yoyogi National Gymnasium'],
  '大阪': ['Osaka Jo Hall', 'Kyocera Dome Osaka'],
  '名古屋': ['Aichi Sky Expo', 'Nagoya Dome'],
  '新加坡': ['Singapore Indoor Stadium', 'The Star Theatre'],
  '曼谷': ['Impact Arena', 'Thunder Dome'],
  '吉隆坡': ['Axiata Arena', 'Zepp Kuala Lumpur'],
};

const TW_STROKE_ORDER = '一乙二十丁厂七卜八人入九几儿了乃刀力又三干于亏士土工才下寸丈大與萬上小口巾山千乞川億个么久及夕女子孑孓孚孛孜宀己已巳巾干幵并廿弋弓彐彡彳心戈戶手支攴文斗斤方火爪父爻爿片牙牛犬王玨玉瓜瓦甘生用田疋疒癶白皮皿目矛矢石示禸禾穴立竹米缶羊羽老而耒耳聿肉臣自至臼舌舛舟艮色艸虍虫血行衣襾見角言谷豆豕豸貝赤走足身車辛辰辵邑酉釆里金長門阜隶隹雨靑非面革韋韭音頁風飛食首香馬骨高髟鬥鬯鬲鬼魚鳥鹵鹿麥麻黃黍黑黹黽鼎鼓鼠鼻齊齒龍龜龠';

const FALLBACK_IMAGE_HTML = '<div class="concert-img-fallback"><div class="concert-img-fallback-bubble" aria-hidden="true">🎤</div></div>';

/* =========================
   2) DOM 快取
   ========================= */
const $ = (id) => document.getElementById(id);

const els = {
  userAvatarWrap: $('userAvatarWrap'),
  userAvatar: $('userAvatar'),
  userLabel: $('userLabel'),
  tokenState: $('tokenState'),
  btnLogin: $('btnLogin'),
  btnLogout: $('btnLogout'),

  totalCountPill: $('totalCountPill'),
  filteredCountPill: $('filteredCountPill'),
  searchInput: $('searchInput'),
  filterType: $('filterType'),
  filterCountry: $('filterCountry'),
  filterDate: $('filterDate'),
  sortSelect: $('sortSelect'),

  loadingState: $('loadingState'),
  emptyState: $('emptyState'),
  cardGrid: $('cardGrid'),
  pageInfo: $('pageInfo'),
  btnPrevPage: $('btnPrevPage'),
  btnNextPage: $('btnNextPage'),
  btnReload: $('btnReload'),

  topArtist: $('topArtist'),
  topArtistMeta: $('topArtistMeta'),
  topCountry: $('topCountry'),
  topCountryMeta: $('topCountryMeta'),
  yearLabel: $('yearLabel'),
  yearInput: $('yearInput'),
  yearCount: $('yearCount'),
  syncStatus: $('syncStatus'),
  syncHint: $('syncHint'),

  mapSummary: $('mapSummary'),
  btnResetMap: $('btnResetMap'),
  leafletMap: $('leafletMap'),
  cityLegend: $('cityLegend'),

  summaryTotal: $('summaryTotal'),
  summaryFavorite: $('summaryFavorite'),
  summaryCities: $('summaryCities'),
  displayCount: $('displayCount'),
  displayPageRate: $('displayPageRate'),

  btnOpenForm: $('btnOpenForm'),
  formBackdrop: $('formBackdrop'),
  btnCloseForm: $('btnCloseForm'),
  entryForm: $('entryForm'),
  entryType: $('entryType'),
  entryConcertName: $('entryConcertName'),
  entryArtist: $('entryArtist'),
  entryDate: $('entryDate'),
  entryCountry: $('entryCountry'),
  entryLocation: $('entryLocation'),
  entryPrice: $('entryPrice'),
  entrySeat: $('entrySeat'),
  entryFavorite: $('entryFavorite'),
  entryImgS: $('entryImgS'),
  entryImgM: $('entryImgM'),
  entryPartner: $('entryPartner'),
  entryNote: $('entryNote'),
  btnClearDraft: $('btnClearDraft'),
  btnSubmitEntry: $('btnSubmitEntry'),

  mapOverlay: $('mapOverlay'),
  mapOverlayBackdrop: $('mapOverlayBackdrop'),
  mapOverlayTitle: $('mapOverlayTitle'),
  mapOverlayMeta: $('mapOverlayMeta'),
  mapOverlayBody: $('mapOverlayBody'),
  btnCloseMapOverlay: $('btnCloseMapOverlay'),

  toastHost: $('toastHost'),
};

/* =========================
   3) 狀態
   ========================= */
const state = {
  raw: [],
  filtered: [],
  currentPage: 1,
  loading: true,
  map: null,
  markersLayer: null,
  cityMarkers: [],
  geocodeCache: loadJson(CONFIG.STORAGE_KEYS.geocode, {}),
  geocodeRequests: new Map(),
  mapReady: false,
  mapRenderVersion: 0,
  tokenClient: null,
  auth: {
    accessToken: '',
    expiresAt: 0,
    profile: null,
  },
  overlay: {
    open: false,
    city: '',
    mode: 'list', // list | detail
    selectedConcertId: null,
  },
  ui: {
    formOpen: false,
  },
  filters: {
    search: '',
    type: '',
    country: '',
    date: '',
    sort: 'date-desc',
  },
};

/* =========================
   4) 工具函式
   ========================= */
function saveJson(key, value) {
  localStorage.setItem(key, JSON.stringify(value));
}

function loadJson(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch (error) {
    return fallback;
  }
}

function pad2(n) {
  return String(n).padStart(2, '0');
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
  return escapeHtml(str).replaceAll('`', '&#96;');
}

function normalizeText(value) {
  return String(value ?? '').trim().toLowerCase();
}

function toBool(value) {
  if (typeof value === 'boolean') return value;
  if (value === 1 || value === '1') return true;
  if (typeof value === 'string') {
    const v = value.trim().toLowerCase();
    return v === 'true' || v === 'yes' || v === 'y' || v === '是';
  }
  return false;
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const text = String(value).trim();
  if (!text) return null;

  const normalized = text.replace(/\./g, '/').replace(/-/g, '/');
  const direct = new Date(normalized);
  if (!Number.isNaN(direct.getTime())) return direct;

  const match = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (match) {
    const d = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    if (!Number.isNaN(d.getTime())) return d;
  }

  return null;
}

function formatDateYMD(value) {
  const d = parseDate(value);
  if (!d) return String(value || '—');
  return `${d.getFullYear()}/${pad2(d.getMonth() + 1)}/${pad2(d.getDate())}`;
}

function formatDateForInput(value) {
  const d = parseDate(value);
  if (!d) return '';
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function getStrokeRank(text) {
  const chars = String(text || '').replace(/\s+/g, '').split('');
  if (!chars.length) return 9999;
  let score = 0;
  for (const char of chars) {
    const idx = TW_STROKE_ORDER.indexOf(char);
    score += idx >= 0 ? idx + 1 : 200;
  }
  return score / chars.length;
}

function compareByStroke(a, b, key, desc = false) {
  const av = getStrokeRank(a[key]);
  const bv = getStrokeRank(b[key]);
  if (av === bv) return String(a[key] || '').localeCompare(String(b[key] || ''), 'zh-Hant');
  return desc ? bv - av : av - bv;
}

function compareByDate(a, b, desc = true) {
  const ad = parseDate(a.Date)?.getTime() || 0;
  const bd = parseDate(b.Date)?.getTime() || 0;
  return desc ? bd - ad : ad - bd;
}

function uniqueValues(arr) {
  return [...new Set(arr.filter(Boolean).map(v => String(v).trim()).filter(Boolean))];
}

function normalizeRow(row) {
  const concertName = row.ConcertName || '';
  const date = row.Date || '';
  const country = row.Country || '';
  const location = row.Location || '';
  return {
    id: String(row.id || `${concertName}__${date}__${country}__${location}__${Math.random().toString(36).slice(2, 8)}`),
    type: row.type || '',
    ConcertName: concertName,
    Artist: row.Artist || '',
    Country: country,
    Location: location,
    Date: date,
    Price: row.Price || '',
    Seat: row.Seat || '',
    imgUrlS: row.imgUrlS || '',
    imgUrlM: row.imgUrlM || '',
    note: row.note || '',
    partner: row.partner || '',
    favorite: toBool(row.favorite),
  };
}

function showBodyLock() {
  document.body.classList.add('modal-open');
}

function updateBodyLock() {
  const shouldLock = state.overlay.open || state.ui.formOpen;
  document.body.classList.toggle('modal-open', shouldLock);
}

function isLoggedIn() {
  return Boolean(state.auth.accessToken && Date.now() < state.auth.expiresAt - 60_000);
}

function setSyncStatus(text, hint = '') {
  els.syncStatus.textContent = text;
  if (hint) els.syncHint.textContent = hint;
}

function updateAuthUi() {
  const loggedIn = isLoggedIn();
  els.btnLogout.disabled = !loggedIn;
  els.userLabel.textContent = loggedIn ? (state.auth.profile?.name || 'Google 已登入') : '尚未登入';
  els.tokenState.textContent = loggedIn ? '可新增資料，資料同步已就緒' : '請登入後新增資料';
  els.syncStatus.textContent = loggedIn ? '已登入' : '待登入';
  els.syncHint.textContent = loggedIn ? '可直接新增到 Google Sheet' : '登入後會自動抓取試算表資料';

  if (loggedIn && state.auth.profile?.picture) {
    els.userAvatar.src = state.auth.profile.picture;
    els.userAvatar.classList.remove('d-none');
    els.userAvatarWrap.querySelector('.avatar-placeholder')?.classList.add('d-none');
  } else {
    els.userAvatar.src = '';
    els.userAvatar.classList.add('d-none');
    const placeholder = els.userAvatarWrap.querySelector('.avatar-placeholder');
    if (placeholder) placeholder.classList.remove('d-none');
  }
}

function showToast(message, title = '提示', variant = 'info') {
  const toast = document.createElement('div');
  toast.className = 'concert-toast';
  const icon = variant === 'success' ? '✓' : variant === 'error' ? '!' : '•';
  toast.innerHTML = `
    <div class="concert-toast__body">
      <p class="concert-toast__title">${escapeHtml(title)} ${icon}</p>
      <p class="concert-toast__message">${escapeHtml(message)}</p>
    </div>
    <button type="button" class="concert-toast__close" aria-label="關閉">×</button>
  `;
  const closeBtn = toast.querySelector('.concert-toast__close');
  const removeToast = () => {
    toast.style.opacity = '0';
    toast.style.transform = 'translateY(10px)';
    setTimeout(() => toast.remove(), 180);
  };
  closeBtn.addEventListener('click', removeToast);
  els.toastHost.appendChild(toast);
  setTimeout(removeToast, 3000);
}

function renderFadeSections() {
  document.querySelectorAll('.fade-up').forEach((el) => el.classList.add('is-visible'));
}

/* =========================
   5) 資料處理
   ========================= */
function getVisibleRows() {
  let rows = [...state.raw];
  const filters = state.filters;

  if (filters.search) {
    const q = normalizeText(filters.search);
    rows = rows.filter((item) => {
      const hay = [
        item.ConcertName,
        item.Artist,
        item.Country,
        item.Location,
        item.Seat,
        item.note,
        item.partner,
        item.type,
      ].map(normalizeText).join(' | ');
      return hay.includes(q);
    });
  }

  if (filters.type) {
    rows = rows.filter((item) => item.type === filters.type);
  }

  if (filters.country) {
    rows = rows.filter((item) => item.Country === filters.country);
  }

  if (filters.date) {
    rows = rows.filter((item) => formatDateForInput(item.Date) === filters.date);
  }

  switch (filters.sort) {
    case 'date-asc':
      rows.sort((a, b) => compareByDate(a, b, false));
      break;
    case 'artist-asc':
      rows.sort((a, b) => compareByStroke(a, b, 'Artist', false) || compareByDate(a, b, true));
      break;
    case 'artist-desc':
      rows.sort((a, b) => compareByStroke(a, b, 'Artist', true) || compareByDate(a, b, true));
      break;
    case 'country-asc':
      rows.sort((a, b) => compareByStroke(a, b, 'Country', false) || compareByDate(a, b, true));
      break;
    case 'country-desc':
      rows.sort((a, b) => compareByStroke(a, b, 'Country', true) || compareByDate(a, b, true));
      break;
    case 'date-desc':
    default:
      rows.sort((a, b) => compareByDate(a, b, true));
      break;
  }

  return rows;
}

function getTopItems(rows, key) {
  const map = new Map();
  rows.forEach((item) => {
    const value = item[key];
    if (!value) return;
    map.set(value, (map.get(value) || 0) + 1);
  });
  const maxCount = Math.max(0, ...map.values());
  const winners = [...map.entries()]
    .filter(([, count]) => count === maxCount)
    .map(([value]) => value);
  return { winners, maxCount };
}

function groupByCity(rows) {
  return rows.reduce((acc, item) => {
    const city = item.Country || '未知城市';
    if (!acc[city]) acc[city] = [];
    acc[city].push(item);
    return acc;
  }, {});
}

function getConcertById(id) {
  return state.raw.find((item) => item.id === id) || null;
}

function getCityConcerts(city) {
  return state.raw
    .filter((item) => item.Country === city)
    .sort((a, b) => compareByDate(a, b, true));
}

/* =========================
   6) 選單 / 草稿
   ========================= */
function populateSelect(selectEl, values, placeholderValue = '') {
  const current = selectEl.value;
  selectEl.innerHTML = '';

  if (placeholderValue !== null) {
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = values.length ? '全部' : '請選擇';
    selectEl.appendChild(placeholder);
  }

  values.forEach((value) => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  });

  if ([...selectEl.options].some((opt) => opt.value === current)) {
    selectEl.value = current;
  } else if (placeholderValue !== null) {
    selectEl.value = '';
  }
}

function hydrateFormSelects() {
  const countries = uniqueValues([...DEFAULT_COUNTRY_OPTIONS, ...state.raw.map((item) => item.Country)]);
  const types = uniqueValues([...DEFAULT_TYPE_OPTIONS, ...state.raw.map((item) => item.type)]);

  const currentCountry = els.entryCountry.value;
  const currentType = els.entryType.value;

  els.entryType.innerHTML = '<option value="">請選擇</option>';
  types.forEach((value) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = value;
    els.entryType.appendChild(opt);
  });

  els.entryCountry.innerHTML = '<option value="">請選擇</option>';
  countries.forEach((value) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = value;
    els.entryCountry.appendChild(opt);
  });

  if ([...els.entryType.options].some((opt) => opt.value === currentType)) els.entryType.value = currentType;
  if ([...els.entryCountry.options].some((opt) => opt.value === currentCountry)) els.entryCountry.value = currentCountry;

  syncLocationOptions();
}

function hydrateFilterSelects() {
  const types = uniqueValues(state.raw.map((item) => item.type));
  const countries = uniqueValues(state.raw.map((item) => item.Country));

  const currentType = els.filterType.value;
  const currentCountry = els.filterCountry.value;

  els.filterType.innerHTML = '<option value="">全部</option>';
  types.forEach((value) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = value;
    els.filterType.appendChild(opt);
  });

  els.filterCountry.innerHTML = '<option value="">全部</option>';
  countries.forEach((value) => {
    const opt = document.createElement('option');
    opt.value = value;
    opt.textContent = value;
    els.filterCountry.appendChild(opt);
  });

  if ([...els.filterType.options].some((opt) => opt.value === currentType)) els.filterType.value = currentType;
  if ([...els.filterCountry.options].some((opt) => opt.value === currentCountry)) els.filterCountry.value = currentCountry;
}

function syncLocationOptions() {
  const selectedCountry = els.entryCountry.value;
  const allLocations = uniqueValues(state.raw.map((item) => item.Location));
  const countryDefaults = DEFAULT_LOCATION_OPTIONS_BY_COUNTRY[selectedCountry] || [];
  const filteredLocations = selectedCountry
    ? uniqueValues([
        ...state.raw.filter((item) => item.Country === selectedCountry).map((item) => item.Location),
        ...countryDefaults,
      ])
    : uniqueValues([...allLocations, ...Object.values(DEFAULT_LOCATION_OPTIONS_BY_COUNTRY).flat()]);
  const options = filteredLocations.length ? filteredLocations : allLocations;
  const current = els.entryLocation.value;

  els.entryLocation.innerHTML = '';

  if (!options.length) {
    const empty = document.createElement('option');
    empty.value = '';
    empty.textContent = selectedCountry ? '目前沒有可用場館' : '請先選 Country';
    els.entryLocation.appendChild(empty);
  } else {
    const placeholder = document.createElement('option');
    placeholder.value = '';
    placeholder.textContent = '請選擇';
    els.entryLocation.appendChild(placeholder);

    options.forEach((value) => {
      const opt = document.createElement('option');
      opt.value = value;
      opt.textContent = value;
      els.entryLocation.appendChild(opt);
    });
  }

  if ([...els.entryLocation.options].some((opt) => opt.value === current)) {
    els.entryLocation.value = current;
  } else {
    els.entryLocation.value = '';
  }
}

function loadDraft() {
  return loadJson(CONFIG.STORAGE_KEYS.draft, null);
}

function saveDraft() {
  const draft = {
    type: els.entryType.value,
    concertName: els.entryConcertName.value,
    artist: els.entryArtist.value,
    date: els.entryDate.value,
    country: els.entryCountry.value,
    location: els.entryLocation.value,
    price: els.entryPrice.value,
    seat: els.entrySeat.value,
    favorite: els.entryFavorite.value,
    imgUrlS: els.entryImgS.value,
    imgUrlM: els.entryImgM.value,
    partner: els.entryPartner.value,
    note: els.entryNote.value,
  };
  saveJson(CONFIG.STORAGE_KEYS.draft, draft);
}

function clearDraft() {
  localStorage.removeItem(CONFIG.STORAGE_KEYS.draft);
}

function fillFormFromDraft() {
  const draft = loadDraft();
  if (!draft) return;

  els.entryType.value = draft.type || '';
  els.entryConcertName.value = draft.concertName || '';
  els.entryArtist.value = draft.artist || '';
  els.entryDate.value = draft.date || '';
  els.entryCountry.value = draft.country || '';
  syncLocationOptions();
  els.entryLocation.value = draft.location || '';
  els.entryPrice.value = draft.price || '';
  els.entrySeat.value = draft.seat || '';
  els.entryFavorite.value = draft.favorite || 'FALSE';
  els.entryImgS.value = draft.imgUrlS || '';
  els.entryImgM.value = draft.imgUrlM || '';
  els.entryPartner.value = draft.partner || '';
  els.entryNote.value = draft.note || '';
}

/* =========================
   7) 主頁面渲染
   ========================= */
function getTotalPages() {
  return Math.max(1, Math.ceil(state.filtered.length / CONFIG.PAGE_SIZE));
}

function updateCounters() {
  const total = state.raw.length;
  const filtered = state.filtered.length;

  els.totalCountPill.textContent = `${total} 場`;
  els.filteredCountPill.textContent = `${filtered} 場`;
  els.displayCount.textContent = `${filtered} 場`;
  els.displayPageRate.textContent = `第 ${Math.min(state.currentPage, getTotalPages())} / ${getTotalPages()} 頁`;
  els.pageInfo.textContent = total
    ? `第 ${Math.min(state.currentPage, getTotalPages())} 頁 / 共 ${getTotalPages()} 頁`
    : '第 1 頁 / 共 1 頁';
}

function renderStats() {
  const topArtist = getTopItems(state.raw, 'Artist');
  const topCountry = getTopItems(state.raw, 'Country');

  els.topArtist.textContent = topArtist.winners.length ? topArtist.winners.join('、') : '—';
  els.topArtistMeta.textContent = topArtist.winners.length ? `${topArtist.maxCount} 場` : '尚無資料';

  els.topCountry.textContent = topCountry.winners.length ? topCountry.winners.join('、') : '—';
  els.topCountryMeta.textContent = topCountry.winners.length ? `${topCountry.maxCount} 場` : '尚無資料';

  els.summaryTotal.textContent = String(state.raw.length);
  els.summaryFavorite.textContent = String(state.raw.filter((item) => item.favorite).length);
  els.summaryCities.textContent = String(uniqueValues(state.raw.map((item) => item.Country)).length);

  const selectedYear = Number(els.yearInput.value || new Date().getFullYear());
  const yearCount = state.raw.filter((item) => {
    const d = parseDate(item.Date);
    return d && d.getFullYear() === selectedYear;
  }).length;
  els.yearLabel.textContent = String(selectedYear);
  els.yearCount.textContent = String(yearCount);
}

function renderFiltersAndFormControls() {
  hydrateFilterSelects();
  hydrateFormSelects();
  updateCounters();
  renderStats();
}

function renderCards() {
  const total = state.filtered.length;
  const totalPages = getTotalPages();
  const currentPage = Math.min(state.currentPage, totalPages);
  const start = (currentPage - 1) * CONFIG.PAGE_SIZE;
  const pageRows = state.filtered.slice(start, start + CONFIG.PAGE_SIZE);

  state.currentPage = currentPage;
  updateCounters();

  if (!state.raw.length) {
    els.loadingState.classList.add('d-none');
    els.emptyState.classList.remove('d-none');
    els.cardGrid.innerHTML = '';
    return;
  }

  els.emptyState.classList.add('d-none');

  if (!total) {
    els.cardGrid.innerHTML = `
      <div class="col-12">
        <div class="empty-state glass-card p-4">
          <div class="empty-icon">💗</div>
          <h3 class="empty-title">找不到符合條件的演唱會</h3>
          <p class="empty-text">可以試著放寬搜尋、類型、城市、日期或排序條件。</p>
        </div>
      </div>
    `;
    return;
  }

  els.cardGrid.innerHTML = pageRows.map((item) => {
    const img = item.imgUrlM || item.imgUrlS || '';
    const hasImage = Boolean(img.trim());
    return `
      <div class="col-12 col-md-6 col-xl-4 card-grid-item">
        <article class="concert-card is-ready" data-concert-id="${escapeAttr(item.id)}">
          <div class="concert-img-wrap">
            ${hasImage ? `<img class="concert-image" src="${escapeAttr(img)}" alt="${escapeAttr(item.ConcertName)}" loading="lazy" data-fallback="concert">` : FALLBACK_IMAGE_HTML}
            ${item.favorite ? '<div class="concert-badges"><span class="badge-soft badge-favorite">最愛 ❤️</span></div>' : ''}
          </div>
          <div class="concert-body">
            <h3 class="concert-title">${escapeHtml(item.ConcertName || '未命名演唱會')}</h3>
            <p class="concert-artist">${escapeHtml(item.Artist || '—')}</p>
            <div class="concert-meta">
              <div>${escapeHtml(formatDateYMD(item.Date))}</div>
              <div>${escapeHtml(item.Country || '—')}  |  ${escapeHtml(item.Location || '—')}</div>
              ${item.Price ? `<div>票價：${escapeHtml(String(item.Price))}</div>` : ''}
              ${item.Seat ? `<div>座位：${escapeHtml(item.Seat)}</div>` : ''}
            </div>
          </div>
        </article>
      </div>
    `;
  }).join('');

  bindCardFallbackImages();
  window.requestAnimationFrame(() => {
    document.querySelectorAll('.concert-card').forEach((card) => card.classList.add('is-ready'));
  });
}

function bindCardFallbackImages() {
  document.querySelectorAll('img[data-fallback="concert"]').forEach((img) => {
    img.addEventListener('error', () => {
      const wrap = img.closest('.concert-img-wrap');
      if (wrap) wrap.innerHTML = FALLBACK_IMAGE_HTML;
    }, { once: true });
  });
}

function renderMainView() {
  state.filtered = getVisibleRows();
  if (state.currentPage > getTotalPages()) state.currentPage = getTotalPages();
  updateCounters();
  renderStats();
  renderCards();
  if (state.overlay.open) renderMapOverlay();
}

/* =========================
   8) 地圖
   ========================= */
async function ensureMap() {
  if (state.map) return;
  state.map = L.map('leafletMap', { scrollWheelZoom: false, zoomControl: true }).setView(CONFIG.DEFAULT_MAP_CENTER, CONFIG.DEFAULT_MAP_ZOOM);
  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '&copy; OpenStreetMap contributors'
  }).addTo(state.map);
  state.markersLayer = L.layerGroup().addTo(state.map);
  state.mapReady = true;
}

async function geocodeCity(city) {
  if (!city) return null;
  if (DEFAULT_CITY_COORDS[city]) return { lat: DEFAULT_CITY_COORDS[city][0], lng: DEFAULT_CITY_COORDS[city][1] };
  if (state.geocodeCache[city]) return state.geocodeCache[city];
  if (state.geocodeRequests.has(city)) return state.geocodeRequests.get(city);

  const request = (async () => {
    try {
      const query = encodeURIComponent(city);
      const url = `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&q=${query}`;
      const response = await axios.get(url, { timeout: 12000, headers: { 'Accept': 'application/json' } });
      const item = Array.isArray(response.data) ? response.data[0] : null;
      if (!item) return null;
      const result = { lat: Number(item.lat), lng: Number(item.lon) };
      if (Number.isFinite(result.lat) && Number.isFinite(result.lng)) {
        state.geocodeCache[city] = result;
        saveJson(CONFIG.STORAGE_KEYS.geocode, state.geocodeCache);
        return result;
      }
    } catch (error) {
      console.warn('Geocode failed:', city, error);
    }
    return null;
  })();

  state.geocodeRequests.set(city, request);
  const resolved = await request;
  state.geocodeRequests.delete(city);
  return resolved;
}

async function renderMap() {
  await ensureMap();
  const version = ++state.mapRenderVersion;
  state.markersLayer.clearLayers();
  state.cityMarkers = [];

  const grouped = groupByCity(state.raw);
  const cities = Object.keys(grouped);
  const markerPoints = [];

  if (!cities.length) {
    els.mapSummary.textContent = '目前沒有可顯示的城市資料。';
    els.cityLegend.innerHTML = '';
    if (state.map) {
      state.map.setView(CONFIG.DEFAULT_MAP_CENTER, CONFIG.DEFAULT_MAP_ZOOM);
    }
    return;
  }

  els.mapSummary.textContent = `共 ${cities.length} 個城市`;
  els.cityLegend.innerHTML = cities
    .sort((a, b) => grouped[b].length - grouped[a].length)
    .map((city) => `<button type="button" class="city-chip" data-city="${escapeAttr(city)}">${escapeHtml(city)} <span class="count">${grouped[city].length}</span></button>`)
    .join('');

  els.cityLegend.querySelectorAll('[data-city]').forEach((btn) => {
    btn.addEventListener('click', () => {
      openMapOverlay(btn.getAttribute('data-city') || '');
    });
  });

  for (const city of cities) {
    if (version !== state.mapRenderVersion) return;
    const items = grouped[city];
    const geo = await geocodeCity(city);
    if (!geo) continue;

    const marker = L.marker([geo.lat, geo.lng], {
      icon: L.divIcon({
        className: '',
        html: `<div class="city-marker"><div class="city-marker-bubble">${items.length}</div></div>`,
        iconSize: [34, 34],
        iconAnchor: [17, 17],
      })
    });

    marker.on('click', () => openMapOverlay(city));
    marker.bindPopup(`
      <div>
        <div class="map-popup-title">${escapeHtml(city)}</div>
        <div class="map-popup-meta">共 ${items.length} 場</div>
      </div>
    `);

    marker.addTo(state.markersLayer);
    state.cityMarkers.push({ city, count: items.length, lat: geo.lat, lng: geo.lng });
    markerPoints.push([geo.lat, geo.lng]);
  }

  if (markerPoints.length) {
    const bounds = L.latLngBounds(markerPoints);
    if (bounds.isValid()) state.map.fitBounds(bounds.pad(0.2), { animate: true });
  } else {
    state.map.setView(CONFIG.DEFAULT_MAP_CENTER, CONFIG.DEFAULT_MAP_ZOOM);
  }
}

function resetMapView() {
  if (!state.map) return;
  state.map.setView(CONFIG.DEFAULT_MAP_CENTER, CONFIG.DEFAULT_MAP_ZOOM, { animate: true });
}

/* =========================
   9) 地圖 Overlay
   ========================= */
function openMapOverlay(city) {
  if (!city) return;
  state.overlay.open = true;
  state.overlay.city = city;
  state.overlay.mode = 'list';
  state.overlay.selectedConcertId = null;
  els.mapOverlay.classList.remove('d-none');
  els.mapOverlay.setAttribute('aria-hidden', 'false');
  updateBodyLock();
  renderMapOverlay();
}

function closeMapOverlay() {
  state.overlay.open = false;
  state.overlay.city = '';
  state.overlay.mode = 'list';
  state.overlay.selectedConcertId = null;
  els.mapOverlay.classList.add('d-none');
  els.mapOverlay.setAttribute('aria-hidden', 'true');
  updateBodyLock();
}

function renderMapOverlay() {
  const city = state.overlay.city;
  const rows = city ? getCityConcerts(city) : [];
  const title = city ? `${city}（${rows.length}）` : '—';
  els.mapOverlayTitle.textContent = title;

  if (!city) {
    els.mapOverlayMeta.textContent = '點選地圖標記即可查看城市清單';
    els.mapOverlayBody.innerHTML = `<div class="overlay-empty">尚未選擇城市</div>`;
    return;
  }

  if (state.overlay.mode === 'detail') {
    const selected = rows.find((item) => item.id === state.overlay.selectedConcertId) || getConcertById(state.overlay.selectedConcertId);
    if (!selected) {
      state.overlay.mode = 'list';
      state.overlay.selectedConcertId = null;
      renderMapOverlay();
      return;
    }

    els.mapOverlayMeta.textContent = '詳細資訊';
    const detailRows = [];
    detailRows.push(`<div class="detail-row"><span class="detail-label">日期</span><span class="detail-value">${escapeHtml(formatDateYMD(selected.Date))}</span></div>`);
    detailRows.push(`<div class="detail-row"><span class="detail-label">地點</span><span class="detail-value">${escapeHtml(selected.Country || '—')} ／ ${escapeHtml(selected.Location || '—')}</span></div>`);
    detailRows.push(`<div class="detail-row"><span class="detail-label">名稱</span><span class="detail-value">${escapeHtml(selected.ConcertName || '—')}</span></div>`);
    detailRows.push(`<div class="detail-row"><span class="detail-label">歌手</span><span class="detail-value">${escapeHtml(selected.Artist || '—')}</span></div>`);
    if (selected.favorite) {
      detailRows.push(`<div class="detail-row"><span class="detail-label">收藏</span><span class="detail-value favorite-chip">♥ 最愛場次</span></div>`);
    }

    els.mapOverlayBody.innerHTML = `
      <div class="overlay-stage">
        <button type="button" class="back-btn" data-back-to-list>← 返回清單</button>
        <article class="overlay-detail-card">
          <h3 class="detail-title">${escapeHtml(selected.ConcertName || '未命名演唱會')}</h3>
          ${selected.Artist ? `<p class="detail-artist">${escapeHtml(selected.Artist)}</p>` : ''}
          <div class="detail-meta-list">
            ${detailRows.join('')}
          </div>
        </article>
      </div>
    `;

    els.mapOverlayBody.querySelector('[data-back-to-list]')?.addEventListener('click', () => {
      state.overlay.mode = 'list';
      state.overlay.selectedConcertId = null;
      renderMapOverlay();
    });

    return;
  }

  els.mapOverlayMeta.textContent = '點選名稱查看詳細資訊';
  if (!rows.length) {
    els.mapOverlayBody.innerHTML = `<div class="overlay-empty">${escapeHtml(city)} 目前沒有演唱會資料</div>`;
    return;
  }

  els.mapOverlayBody.innerHTML = `
    <div class="overlay-stage">
      <div class="overlay-list">
        ${rows.map((item) => `
          <button type="button" class="overlay-list-item" data-concert-id="${escapeAttr(item.id)}">${escapeHtml(item.ConcertName || '未命名演唱會')}</button>
        `).join('')}
      </div>
    </div>
  `;
}

/* =========================
   10) 表單 / 寫入 Google Sheet
   ========================= */
function openFormOverlay() {
  state.ui.formOpen = true;
  els.formBackdrop.classList.remove('d-none');
  els.formBackdrop.setAttribute('aria-hidden', 'false');
  updateBodyLock();
  fillFormFromDraft();
  syncLocationOptions();
  setTimeout(() => els.entryConcertName.focus(), 0);
}

function closeFormOverlay() {
  state.ui.formOpen = false;
  els.formBackdrop.classList.add('d-none');
  els.formBackdrop.setAttribute('aria-hidden', 'true');
  updateBodyLock();
}

function collectFormData() {
  return {
    id: String(Date.now()),
    type: els.entryType.value.trim(),
    ConcertName: els.entryConcertName.value.trim(),
    Artist: els.entryArtist.value.trim(),
    Country: els.entryCountry.value.trim(),
    Location: els.entryLocation.value.trim(),
    Date: els.entryDate.value.trim(),
    Price: els.entryPrice.value.trim(),
    Seat: els.entrySeat.value.trim(),
    imgUrlS: els.entryImgS.value.trim(),
    imgUrlM: els.entryImgM.value.trim(),
    note: els.entryNote.value.trim(),
    partner: els.entryPartner.value.trim(),
    favorite: els.entryFavorite.value === 'TRUE' ? 'TRUE' : 'FALSE',
  };
}

function validateForm(payload) {
  const requiredFields = ['type', 'ConcertName', 'Artist', 'Country', 'Location', 'Date'];
  for (const key of requiredFields) {
    if (!payload[key]) return `欄位 ${key} 不可空白。`;
  }
  return '';
}

function setSubmitLoading(loading) {
  els.btnSubmitEntry.disabled = loading;
  els.btnClearDraft.disabled = loading;
  els.btnLogin.disabled = loading;
  els.btnLogout.disabled = loading;
  els.btnSubmitEntry.textContent = loading ? '送出中...' : '送出';
}

async function appendRowToSheet(payload) {
  const headers = {
    Authorization: `Bearer ${state.auth.accessToken}`,
    'Content-Type': 'application/json',
  };

  const body = {
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
    ]]
  };

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SPREADSHEET_ID}/values/${encodeURIComponent(CONFIG.APPEND_RANGE)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
  await axios.post(url, body, { headers, timeout: 15000 });
}

function insertLocalConcert(payload) {
  const record = normalizeRow({
    ...payload,
    favorite: payload.favorite === 'TRUE',
  });
  state.raw.unshift(record);
  state.currentPage = 1;
  hydrateFilterSelects();
  hydrateFormSelects();
  renderMainView();
  renderMap();
  if (state.overlay.open) renderMapOverlay();
}

async function handleSubmit(event) {
  event.preventDefault();
  const payload = collectFormData();
  const error = validateForm(payload);
  if (error) {
    showToast(error, '表單驗證失敗', 'error');
    return;
  }

  if (!isLoggedIn()) {
    showToast('登入已過期，請先重新登入。', '登入失敗', 'error');
    updateAuthUi();
    return;
  }

  try {
    setSubmitLoading(true);
    await appendRowToSheet(payload);
    clearDraft();
    els.entryForm.reset();
    hydrateFormSelects();
    insertLocalConcert(payload);
    closeFormOverlay();
    showToast('已成功寫入 Google Sheet！', '送出成功', 'success');
  } catch (error) {
    console.error(error);
    const status = Number(error?.response?.status || 0);
    if (status === 401) {
      localStorage.removeItem(CONFIG.STORAGE_KEYS.token);
      state.auth.accessToken = '';
      state.auth.expiresAt = 0;
      state.auth.profile = null;
      updateAuthUi();
      showToast('登入已過期，請重新登入後再送出。', '送出失敗', 'error');
    } else {
      showToast('送出失敗，請確認試算表權限、Spreadsheet ID 與欄位是否正確。', '送出失敗', 'error');
    }
  } finally {
    setSubmitLoading(false);
  }
}

/* =========================
   11) Google OAuth
   ========================= */
function persistToken(tokenResponse) {
  const expiresAt = Date.now() + ((tokenResponse.expires_in || 3600) * 1000);
  const payload = {
    accessToken: tokenResponse.access_token,
    expiresAt,
  };
  saveJson(CONFIG.STORAGE_KEYS.token, payload);
  state.auth.accessToken = payload.accessToken;
  state.auth.expiresAt = payload.expiresAt;
}

function restoreToken() {
  const saved = loadJson(CONFIG.STORAGE_KEYS.token, null);
  if (!saved || !saved.accessToken || !saved.expiresAt) return false;
  if (Date.now() >= saved.expiresAt - 60_000) {
    localStorage.removeItem(CONFIG.STORAGE_KEYS.token);
    return false;
  }
  state.auth.accessToken = saved.accessToken;
  state.auth.expiresAt = saved.expiresAt;
  return true;
}

async function fetchUserProfile() {
  try {
    const res = await axios.get('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${state.auth.accessToken}` },
      timeout: 12000,
    });
    state.auth.profile = res.data || null;
  } catch (error) {
    state.auth.profile = null;
  }
}

function initTokenClient() {
  if (!window.google?.accounts?.oauth2) return false;
  state.tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.GOOGLE_CLIENT_ID,
    scope: 'openid email profile https://www.googleapis.com/auth/spreadsheets',
    callback: async (tokenResponse) => {
      if (tokenResponse.error) {
        showToast(`登入失敗：${tokenResponse.error}`, 'Google 登入', 'error');
        return;
      }
      persistToken(tokenResponse);
      await fetchUserProfile();
      updateAuthUi();
      showToast('登入成功，可以開始新增資料。', 'Google 登入', 'success');
    },
  });
  return true;
}

function requestLogin() {
  if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.startsWith('YOUR_')) {
    showToast('請先在 script.js 裡設定 GOOGLE_CLIENT_ID。', '設定未完成', 'error');
    return;
  }
  if (!state.tokenClient && !initTokenClient()) {
    showToast('Google 登入元件尚未載入完成，請稍後再試。', '登入中', 'error');
    return;
  }
  state.tokenClient.requestAccessToken({ prompt: 'consent' });
}

function logout() {
  localStorage.removeItem(CONFIG.STORAGE_KEYS.token);
  state.auth.accessToken = '';
  state.auth.expiresAt = 0;
  state.auth.profile = null;
  updateAuthUi();
  showToast('已登出。', 'Google 登入', 'success');
}

async function restoreAuth() {
  if (restoreToken()) {
    await fetchUserProfile();
  }
  updateAuthUi();
}

async function waitForGoogleReady() {
  const started = Date.now();
  while (!window.google?.accounts?.oauth2) {
    if (Date.now() - started > 8000) return;
    await new Promise((resolve) => setTimeout(resolve, 120));
  }
  initTokenClient();
}

/* =========================
   12) 事件綁定
   ========================= */
function bindEvents() {
  els.searchInput.addEventListener('input', () => {
    state.filters.search = els.searchInput.value.trim();
    state.currentPage = 1;
    renderMainView();
  });

  els.filterType.addEventListener('change', () => {
    state.filters.type = els.filterType.value;
    state.currentPage = 1;
    renderMainView();
  });

  els.filterCountry.addEventListener('change', () => {
    state.filters.country = els.filterCountry.value;
    state.currentPage = 1;
    renderMainView();
  });

  els.filterDate.addEventListener('change', () => {
    state.filters.date = els.filterDate.value;
    state.currentPage = 1;
    renderMainView();
  });

  els.sortSelect.addEventListener('change', () => {
    state.filters.sort = els.sortSelect.value;
    state.currentPage = 1;
    renderMainView();
  });

  els.btnPrevPage.addEventListener('click', () => {
    state.currentPage = Math.max(1, state.currentPage - 1);
    renderCards();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  els.btnNextPage.addEventListener('click', () => {
    state.currentPage = Math.min(getTotalPages(), state.currentPage + 1);
    renderCards();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  els.btnReload.addEventListener('click', fetchConcerts);
  els.btnResetMap.addEventListener('click', resetMapView);

  els.btnOpenForm.addEventListener('click', () => openFormOverlay());
  els.btnCloseForm.addEventListener('click', () => closeFormOverlay());
  els.formBackdrop.addEventListener('click', (event) => {
    if (event.target === els.formBackdrop) closeFormOverlay();
  });

  els.btnCloseMapOverlay.addEventListener('click', closeMapOverlay);
  els.mapOverlayBackdrop.addEventListener('click', closeMapOverlay);
  els.mapOverlay.addEventListener('click', (event) => {
    if (event.target === els.mapOverlay) closeMapOverlay();
  });

  els.btnLogin.addEventListener('click', requestLogin);
  els.btnLogout.addEventListener('click', logout);

  els.btnClearDraft.addEventListener('click', () => {
    els.entryForm.reset();
    hydrateFormSelects();
    clearDraft();
    showToast('已清空表單。', '草稿', 'success');
  });

  els.entryCountry.addEventListener('change', () => {
    syncLocationOptions();
    saveDraft();
  });

  [
    els.entryType, els.entryConcertName, els.entryArtist, els.entryDate,
    els.entryCountry, els.entryLocation, els.entryPrice, els.entrySeat,
    els.entryFavorite, els.entryImgS, els.entryImgM, els.entryPartner, els.entryNote,
  ].forEach((input) => {
    input.addEventListener('input', saveDraft);
    input.addEventListener('change', saveDraft);
  });

  els.entryForm.addEventListener('submit', handleSubmit);

  els.yearInput.addEventListener('input', () => {
    renderStats();
  });

  document.addEventListener('keydown', (event) => {
    if (event.key === 'Escape') {
      if (state.ui.formOpen) {
        closeFormOverlay();
      } else if (state.overlay.open) {
        closeMapOverlay();
      }
    }
  });
  els.mapOverlayBody.addEventListener('click', (event) => {
    const target = event.target.closest('[data-concert-id]');
    if (!target) return;
    state.overlay.mode = 'detail';
    state.overlay.selectedConcertId = target.getAttribute('data-concert-id');
    renderMapOverlay();
  });
}

/* =========================
   13) 資料載入
   ========================= */
async function fetchConcerts() {
  state.loading = true;
  els.loadingState.classList.remove('d-none');
  els.emptyState.classList.add('d-none');
  els.cardGrid.innerHTML = '';
  setSyncStatus('載入中', '正在抓取公開資料與建立視覺內容...');

  try {
    const response = await axios.get(CONFIG.OPEN_SHEET_URL, {
      timeout: 15000,
      headers: { Accept: 'application/json' },
    });

    const rows = Array.isArray(response.data) ? response.data.map(normalizeRow) : [];
    state.raw = rows.filter((item) => item.ConcertName || item.Artist || item.Date);
    state.loading = false;

    renderFiltersAndFormControls();
    state.filtered = getVisibleRows();
    renderMainView();
    await renderMap();

    els.loadingState.classList.add('d-none');
    if (!state.raw.length) {
      els.emptyState.classList.remove('d-none');
      setSyncStatus('無資料', '目前試算表內尚未有資料');
    } else {
      setSyncStatus('已同步', '資料已載入完成');
    }
  } catch (error) {
    console.error(error);
    state.loading = false;
    els.loadingState.classList.add('d-none');
    els.emptyState.classList.remove('d-none');
    setSyncStatus('載入失敗', '請確認公開 Sheet JSON URL 與分享權限是否正確');
    showToast('資料載入失敗，請確認公開 Sheet JSON URL 與分享權限是否正確。', '載入失敗', 'error');
  }
}

/* =========================
   14) 初始化
   ========================= */
async function init() {
  const currentYear = new Date().getFullYear();
  els.yearInput.value = String(currentYear);
  els.yearLabel.textContent = String(currentYear);
  els.sortSelect.value = 'date-desc';
  state.filters.sort = 'date-desc';

  bindEvents();
  renderFadeSections();
  await restoreAuth();
  await waitForGoogleReady();
  await fetchConcerts();
  if (state.overlay.open) renderMapOverlay();
}

document.addEventListener('DOMContentLoaded', init);
 

