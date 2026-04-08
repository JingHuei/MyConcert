/* =========================================================
   Lucy Concert Diary — script.js
   - 公開 OpenSheet JSON 讀取
   - Google GIS OAuth 登入 + Sheets API append
   - Leaflet 地圖 / 統計 / 篩選 / 分頁 / 草稿保存
   ========================================================= */

/* =========================
   1) 請在這裡填入你的資料
   ========================= */
const CONFIG = {
  GOOGLE_CLIENT_ID: '562464022417-8c2sckejaft6de7ch8kejqomm1fbi0ga.apps.googleusercontent.com',
  SPREADSHEET_ID: '1tfNjim8BbKmv3KVNIQrRvx2T6W8YtfjyJCr1_wXDcrA',
  SHEET_NAME: '演唱會紀錄',
  OPEN_SHEET_URL: 'https://opensheet.elk.sh/1tfNjim8BbKmv3KVNIQrRvx2T6W8YtfjyJCr1_wXDcrA/%E6%BC%94%E5%94%B1%E6%9C%83%E7%B4%80%E9%8C%84',
  APPEND_RANGE: '演唱會紀錄!A:N',
  PAGE_SIZE: 10,
  STORAGE_KEYS: {
    draft: 'lucyConcertDraftV1',
    token: 'lucyConcertTokenV1',
  },
  DEFAULT_MAP_CENTER: [25.033964, 121.564468],
  DEFAULT_MAP_ZOOM: 10.5,
  GEOCODE_CACHE_KEY: 'lucyConcertGeocodeCacheV1'
};

const TW_STROKE_ORDER = '一乙二十丁厂七卜八人入九几儿了乃刀力又三干于亏士土工才下寸丈大與萬上小口巾山千乞川億个么久及夕女子孑孓孚孛孜宀己已巳巾干幵并廿弋弓彐彡彳心戈戶手支攴文斗斤方火爪父爻爿片牙牛犬王玨玉瓜瓦甘生用田疋疒癶白皮皿目矛矢石示禸禾穴立竹米缶羊羽老而耒耳聿肉臣自至臼舌舛舟艮色艸虍虫血行衣襾見角言谷豆豕豸貝赤走足身車辛辰辵邑酉釆里金長門阜隶隹雨靑非面革韋韭音頁風飛食首香馬骨高髟鬥鬯鬲鬼魚鳥鹵鹿麥麻黃黍黑黹黽鼎鼓鼠鼻齊齒龍龜龠';

/* =========================
   2) DOM 快取
   ========================= */
const els = {
  statusPanel: document.getElementById('statusPanel'),
  statusText: document.getElementById('statusText'),
  concertGrid: document.getElementById('concertGrid'),
  resultsMeta: document.getElementById('resultsMeta'),
  pageInfo: document.getElementById('pageInfo'),
  prevPageBtn: document.getElementById('prevPageBtn'),
  nextPageBtn: document.getElementById('nextPageBtn'),
  searchInput: document.getElementById('searchInput'),
  dateFromInput: document.getElementById('dateFromInput'),
  dateToInput: document.getElementById('dateToInput'),
  cityFilter: document.getElementById('cityFilter'),
  artistFilter: document.getElementById('artistFilter'),
  typeFilter: document.getElementById('typeFilter'),
  yearFilter: document.getElementById('yearFilter'),
  statYearFilter: document.getElementById('statYearFilter'),
  sortSelect: document.getElementById('sortSelect'),
  resetFiltersBtn: document.getElementById('resetFiltersBtn'),
  activeChips: document.getElementById('activeChips'),
  citySummary: document.getElementById('citySummary'),
  mapHint: document.getElementById('mapHint'),
  topArtistText: document.getElementById('topArtistText'),
  topCityText: document.getElementById('topCityText'),
  yearCountText: document.getElementById('yearCountText'),
  heroTotalConcerts: document.getElementById('heroTotalConcerts'),
  heroTotalCities: document.getElementById('heroTotalCities'),
  heroFavoriteCount: document.getElementById('heroFavoriteCount'),
  addConcertBtn: document.getElementById('addConcertBtn'),
  modalOverlay: document.getElementById('modalOverlay'),
  closeModalBtn: document.getElementById('closeModalBtn'),
  authPrompt: document.getElementById('authPrompt'),
  loginBtn: document.getElementById('loginBtn'),
  concertForm: document.getElementById('concertForm'),
  clearDraftBtn: document.getElementById('clearDraftBtn'),
  submitFormBtn: document.getElementById('submitFormBtn'),
  formMessage: document.getElementById('formMessage'),
  formType: document.getElementById('formType'),
  formConcertName: document.getElementById('formConcertName'),
  formArtist: document.getElementById('formArtist'),
  formCountry: document.getElementById('formCountry'),
  formLocation: document.getElementById('formLocation'),
  formDate: document.getElementById('formDate'),
  formPrice: document.getElementById('formPrice'),
  formSeat: document.getElementById('formSeat'),
  formImgUrlS: document.getElementById('formImgUrlS'),
  formImgUrlM: document.getElementById('formImgUrlM'),
  formFavorite: document.getElementById('formFavorite'),
  formNote: document.getElementById('formNote'),
  formPartner: document.getElementById('formPartner'),
  map: document.getElementById('map'),
  mapSelectedCityTitle: document.getElementById('mapSelectedCityTitle'),
  mapSelectedCityMeta: document.getElementById('mapSelectedCityMeta'),
  mapCityConcertGrid: document.getElementById('mapCityConcertGrid'),
};

/* =========================
   3) 狀態
   ========================= */
const state = {
  raw: [],
  filtered: [],
  currentPage: 1,
  map: null,
  markersLayer: null,
  cityMarkers: [],
  geocodeCache: loadJson(CONFIG.GEOCODE_CACHE_KEY, {}),
  loading: true,
  selectedCityFromMap: '',
  auth: {
    accessToken: '',
    expiresAt: 0,
    profile: null,
  },
  tokenClient: null,
  filters: {
    search: '',
    dateFrom: '',
    dateTo: '',
    city: '',
    artist: '',
    type: '',
    year: '',
    sort: 'dateDesc',
  }
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

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const text = String(value).trim();
  if (!text) return null;

  const normalized = text.replace(/\./g, '/').replace(/-/g, '/');
  const d = new Date(normalized);
  if (!Number.isNaN(d.getTime())) return d;

  const m = text.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (m) {
    const dd = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    if (!Number.isNaN(dd.getTime())) return dd;
  }
  return null;
}

function formatDatePretty(value) {
  const d = parseDate(value);
  if (!d) return value || '—';
  return new Intl.DateTimeFormat('zh-TW', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    weekday: 'short'
  }).format(d);
}

function daysDiffText(dateValue) {
  const d = parseDate(dateValue);
  if (!d) return '日期未知';
  const today = new Date();
  const startToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const startEvent = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const diff = Math.floor((startToday - startEvent) / 86400000);
  if (diff >= 0) return `已過 ${diff} 天`;
  return `距今天還有 ${Math.abs(diff)} 天`;
}

function uniqueValues(arr) {
  return [...new Set(arr.filter(Boolean).map(v => String(v).trim()).filter(Boolean))];
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

function normalizeRow(row) {
  return {
    id: row.id || '',
    type: row.type || '',
    ConcertName: row.ConcertName || '',
    Artist: row.Artist || '',
    Country: row.Country || '',
    Location: row.Location || '',
    Date: row.Date || '',
    Price: row.Price || '',
    Seat: row.Seat || '',
    imgUrlS: row.imgUrlS || '',
    imgUrlM: row.imgUrlM || '',
    note: row.note || '',
    partner: row.partner || '',
    favorite: toBool(row.favorite),
  };
}

function showStatus(message, show = true) {
  els.statusText.textContent = message;
  els.statusPanel.hidden = !show;
}

function hideStatus() {
  els.statusPanel.hidden = true;
}

function normalizeText(value) {
  return String(value || '').trim().toLowerCase();
}

function loadDraft() {
  return loadJson(CONFIG.STORAGE_KEYS.draft, null);
}

function saveDraft() {
  const draft = {
    type: els.formType.value,
    concertName: els.formConcertName.value,
    artist: els.formArtist.value,
    country: els.formCountry.value,
    location: els.formLocation.value,
    date: els.formDate.value,
    price: els.formPrice.value,
    seat: els.formSeat.value,
    imgUrlS: els.formImgUrlS.value,
    imgUrlM: els.formImgUrlM.value,
    favorite: els.formFavorite.checked,
    note: els.formNote.value,
    partner: els.formPartner.value,
  };
  saveJson(CONFIG.STORAGE_KEYS.draft, draft);
}

function clearDraft() {
  localStorage.removeItem(CONFIG.STORAGE_KEYS.draft);
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

/* =========================
   5) 資料過濾 / 統計基礎
   ========================= */
function getVisibleRows() {
  let rows = [...state.raw];
  const filters = state.filters;

  if (filters.search) {
    const q = normalizeText(filters.search);
    rows = rows.filter(item => {
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

  if (filters.dateFrom) {
    const from = new Date(filters.dateFrom);
    rows = rows.filter(item => {
      const d = parseDate(item.Date);
      return d && d >= new Date(from.getFullYear(), from.getMonth(), from.getDate());
    });
  }

  if (filters.dateTo) {
    const to = new Date(filters.dateTo);
    rows = rows.filter(item => {
      const d = parseDate(item.Date);
      return d && d <= new Date(to.getFullYear(), to.getMonth(), to.getDate(), 23, 59, 59);
    });
  }

  if (filters.city) {
    rows = rows.filter(item => item.Country === filters.city);
  }

  if (filters.artist) {
    rows = rows.filter(item => item.Artist === filters.artist);
  }

  if (filters.type) {
    rows = rows.filter(item => item.type === filters.type);
  }

  if (filters.year) {
    rows = rows.filter(item => {
      const d = parseDate(item.Date);
      return d && String(d.getFullYear()) === String(filters.year);
    });
  }

  switch (filters.sort) {
    case 'dateAsc':
      rows.sort((a, b) => compareByDate(a, b, false));
      break;
    case 'artistStrokeAsc':
      rows.sort((a, b) => compareByStroke(a, b, 'Artist', false) || compareByDate(a, b, true));
      break;
    case 'artistStrokeDesc':
      rows.sort((a, b) => compareByStroke(a, b, 'Artist', true) || compareByDate(a, b, true));
      break;
    case 'cityStrokeAsc':
      rows.sort((a, b) => compareByStroke(a, b, 'Country', false) || compareByDate(a, b, true));
      break;
    case 'cityStrokeDesc':
      rows.sort((a, b) => compareByStroke(a, b, 'Country', true) || compareByDate(a, b, true));
      break;
    case 'dateDesc':
    default:
      rows.sort((a, b) => compareByDate(a, b, true));
      break;
  }

  return rows;
}

function groupByCity(rows) {
  return rows.reduce((acc, item) => {
    const city = item.Country || '未知城市';
    if (!acc[city]) acc[city] = [];
    acc[city].push(item);
    return acc;
  }, {});
}

function getTopItems(rows, key) {
  const map = new Map();
  rows.forEach(item => {
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

/* =========================
   6) 資料讀取
   ========================= */
async function fetchConcerts() {
  showStatus('正在抓取公開資料與建立視覺內容...');
  try {
    const response = await axios.get(CONFIG.OPEN_SHEET_URL, {
      timeout: 15000,
      headers: { 'Accept': 'application/json' },
    });
    const rows = Array.isArray(response.data) ? response.data.map(normalizeRow) : [];
    state.raw = rows.filter(item => item.ConcertName || item.Artist || item.Date);
    state.loading = false;

    hydrateFilters();

  function setLoading(isLoading) {
    els.loadingState.classList.toggle('hidden', !isLoading);
  }

  function setEmpty(isEmpty) {
    els.emptyState.classList.toggle('hidden', !isEmpty);
  }

  function parseDate(value) {
    if (!value) return null;
    const text = String(value).trim().replaceAll('/', '-');
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function formatDate(value) {
    const d = parseDate(value);
    return d ? dateFmt.format(d) : (value || '-');
  }

  function daysSince(value) {
    const d = parseDate(value);
    if (!d) return '-';
    const diff = Math.floor((Date.now() - d.getTime()) / (1000 * 60 * 60 * 24));
    if (diff >= 0) return `距今 ${numberFmt.format(diff)} 天`;
    return `尚有 ${numberFmt.format(Math.abs(diff))} 天`;
  }

  function toNumber(value) {
    if (value === null || value === undefined || value === '') return null;
    const n = Number(String(value).replaceAll(',', '').trim());
    return Number.isFinite(n) ? n : null;
  }

  function typeLabel(value) {
    const map = {
      '演唱會': '演唱會',
      '拼盤': '拼盤',
      'fanmeeting': 'FanMeeting',
      'FanMeeting': 'FanMeeting',
      'FanConcert': 'FanConcert',
    };
    return map[value] || value || '-';
  }

  function normalizeTypeForFilter(value) {
    const text = String(value || '').trim();
    if (text === 'FanConcert') return 'FanConcert';
    if (text === 'FanMeeting' || text === 'fanmeeting') return 'FanMeeting';
    return text || '';
  }

  function cityKey(value) {
    return String(value || '').trim();
  }

  function safeStrokeCompare(a, b) {
    try {
      return strokeCollator.compare(a, b);
    } catch {
      return String(a).localeCompare(String(b), 'zh-Hant');
    }
  }

  function groupBy(items, keyFn) {
    return items.reduce((acc, item) => {
      const key = keyFn(item);
      if (!acc[key]) acc[key] = [];
      acc[key].push(item);
      return acc;
    }, {});
  }

  // ========== 初始資料載入 ==========
  async function loadInitialData() {
    setLoading(true);
    try {
      const [recordsRes, metaRes] = await Promise.all([
        axios.get(`${RECORDS_URL}?_=${Date.now()}`),
        axios.get(`${META_URL}?_=${Date.now()}`),
      ]);

      state.allRecords = normalizeRecords(recordsRes.data);
      state.metaRows = Array.isArray(metaRes.data) ? metaRes.data : [];
      prepareFilters();
      await warmCityGeocodes();
      renderAll();
      initMapIfNeeded();
      renderCityMarkers();
    } catch (error) {
      console.error(error);
      showToast('資料載入失敗，請確認公開試算表與工作表名稱是否正確。');
      els.cardsGrid.innerHTML = '';
      setEmpty(true);
    } finally {
      setLoading(false);
    }
  }

  function normalizeRecords(rows) {
    if (!Array.isArray(rows)) return [];
    return rows
      .filter(Boolean)
      .map((row) => {
        const record = {
          row: row,
          id: String(row.id ?? row.ID ?? row.timestamp ?? Date.now()),
          type: normalizeTypeForFilter(row.type),
          ConcertName: row.ConcertName ?? row.concertName ?? '',
          Artist: row.Artist ?? row.artist ?? '',
          Country: row.Country ?? row.country ?? '',
          Location: row.Location ?? row.location ?? '',
          Date: row.Date ?? row.date ?? '',
          Price: row.Price ?? row.price ?? '',
          Seat: row.Seat ?? row.seat ?? '',
          imgUrlS: row.imgUrlS ?? row.imgUrlS1 ?? row.imgS ?? '',
          imgUrlM: row.imgUrlM ?? row.imgUrlM1 ?? row.imgM ?? '',
          note: row.note ?? '',
          partner: row.partner ?? '',
          favorite: String(row.favorite ?? '').toLowerCase() === 'true' || row.favorite === 'TRUE' || row.favorite === 'yes' || row.favorite === '1',
        };
        record.dateObj = parseDate(record.Date);
        record.year = record.dateObj ? String(record.dateObj.getFullYear()) : '';
        record.priceNumber = toNumber(record.Price);
        record.daysText = daysSince(record.Date);
        return record;
      })
      .sort((a, b) => (b.dateObj?.getTime() || 0) - (a.dateObj?.getTime() || 0));
  }

  function prepareFilters() {
    const types = [...new Set(state.allRecords.map((r) => r.type).filter(Boolean))];
    const cities = [...new Set(state.allRecords.map((r) => r.Country).filter(Boolean))];
    const artists = [...new Set(state.allRecords.map((r) => r.Artist).filter(Boolean))];
    const years = [...new Set(state.allRecords.map((r) => r.year).filter(Boolean))]
      .sort((a, b) => Number(b) - Number(a));

    fillSelect(els.typeFilter, [
      { value: '', label: '全部類型' },
      ...types.map((t) => ({ value: t, label: typeLabel(t) })),
    ]);
    fillSelect(els.cityFilter, [
      { value: '', label: '全部城市' },
      ...cities.map((c) => ({ value: c, label: c })),
    ]);
    fillSelect(els.artistFilter, [
      { value: '', label: '全部藝人' },
      ...artists.map((a) => ({ value: a, label: a })),
    ]);
    fillSelect(els.yearSelect, years.length
      ? years.map((y) => ({ value: y, label: y }))
      : [{ value: '', label: '無年份資料' }]
    );

    fillSelect(els.sortSelect, [
      { value: 'date_desc', label: '日期：新 → 舊' },
      { value: 'date_asc', label: '日期：舊 → 新' },
      { value: 'artist_asc', label: '藝人：筆畫少 → 多' },
      { value: 'artist_desc', label: '藝人：筆畫多 → 少' },
      { value: 'city_asc', label: '城市：筆畫少 → 多' },
      { value: 'city_desc', label: '城市：筆畫多 → 少' },
    ]);

    fillSelect(els.formType, [
      { value: '', label: '請選擇類型' },
      ...types.map((t) => ({ value: t, label: typeLabel(t) })),
    ]);
    fillSelect(els.formCountry, [
      { value: '', label: '請選擇城市' },
      ...cities.map((c) => ({ value: c, label: c })),
    ]);
    fillSelect(els.formLocation, [
      { value: '', label: '請選擇場館' },
      ...getMetaUnique('Location').map((v) => ({ value: v, label: v })),
    ]);

    if (years.length) els.yearSelect.value = years[0];
  }

  

function fillSelect(selectEl, entries) {
  if (!selectEl) return;

  selectEl.innerHTML = entries.map((entry) => {
    
    if (Array.isArray(entry)) {
      return `<option value="${entry[0]}">${entry[1]}</option>`;
    }

    const value = entry?.value ?? '';
    const label = entry?.label ?? value;

    return `<option value="${value}">${label}</option>`;
  }).join('');
}

  function getMetaUnique(column) {
    const values = state.metaRows
      .map((row) => row?.[column])
      .filter((v) => String(v || '').trim() !== '')
      .map((v) => String(v).trim());
    return [...new Set(values)];
  }

  // ========== 篩選 / 排序 ==========
  function applyFilters() {
    const { keyword, type, city, artist, dateFrom, dateTo } = state.filters;
    let result = [...state.allRecords];

    if (keyword) {
      const q = keyword.toLowerCase();
      result = result.filter((r) => [
        r.ConcertName, r.Artist, r.Country, r.Location, r.note, r.partner, r.type
      ].some((field) => String(field || '').toLowerCase().includes(q)));
    }

    renderAll();
  } catch (error) {
    console.error(error);
    showStatus('資料載入失敗，請確認公開 Sheet JSON URL 與分享權限是否正確。');
  }
}

/* =========================
   7) 篩選選單與統計
   ========================= */
function populateSelect(selectEl, values) {
  const current = selectEl.value;
  selectEl.innerHTML = '';

  values.forEach((value, index) => {
    const option = document.createElement('option');
    option.value = index === 0 ? '' : value;
    option.textContent = value;
    selectEl.appendChild(option);
  });

  if ([...selectEl.options].some(opt => opt.value === current)) {
    selectEl.value = current;
  }
}

function populateFormSelect(selectEl, placeholder, values) {
  const current = selectEl.value;
  selectEl.innerHTML = '';

  const placeholderOption = document.createElement('option');
  placeholderOption.value = '';
  placeholderOption.textContent = placeholder;
  selectEl.appendChild(placeholderOption);

  values.forEach(value => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    selectEl.appendChild(option);
  });

  if ([...selectEl.options].some(opt => opt.value === current)) {
    selectEl.value = current;
  } else {
    selectEl.value = '';
  }
}

function populateYearSelects(years) {
  const currentFilterYear = els.yearFilter.value;
  const currentStatYear = els.statYearFilter.value;
  const options = ['全部年份', ...years];

  populateSelect(els.yearFilter, options);

  els.statYearFilter.innerHTML = '';
  const placeholder = document.createElement('option');
  placeholder.value = '';
  placeholder.textContent = '這一';
  els.statYearFilter.appendChild(placeholder);
  years.forEach(year => {
    const option = document.createElement('option');
    option.value = year;
    option.textContent = year;
    els.statYearFilter.appendChild(option);
  });

  if ([...els.statYearFilter.options].some(opt => opt.value === currentStatYear)) {
    els.statYearFilter.value = currentStatYear;
  } else {
    els.statYearFilter.value = '';
  }

  if ([...els.yearFilter.options].some(opt => opt.value === currentFilterYear)) {
    els.yearFilter.value = currentFilterYear;
  } else {
    els.yearFilter.value = '';
  }
}

function syncLocationOptions() {
  const allLocations = uniqueValues(state.raw.map(item => item.Location));
  const selectedCountry = els.formCountry.value;
  const options = allLocations.filter(loc => {
    if (!selectedCountry) return true;
    return state.raw.some(row => row.Country === selectedCountry && row.Location === loc);
  });

  const list = options.length ? options : allLocations;
  const current = els.formLocation.value;
  els.formLocation.innerHTML = '';

  list.forEach(loc => {
    const option = document.createElement('option');
    option.value = loc;
    option.textContent = loc;
    els.formLocation.appendChild(option);
  });

  if ([...els.formLocation.options].some(opt => opt.value === current)) {
    els.formLocation.value = current;
  } else if (els.formLocation.options.length) {
    els.formLocation.selectedIndex = 0;
  }
}

function hydrateFilters() {
  const cities = uniqueValues(state.raw.map(item => item.Country));
  const artists = uniqueValues(state.raw.map(item => item.Artist));
  const types = uniqueValues(state.raw.map(item => item.type));
  const years = uniqueValues(state.raw.map(item => {
    const d = parseDate(item.Date);
    return d ? String(d.getFullYear()) : '';
  })).sort((a, b) => Number(b) - Number(a));

  populateSelect(els.cityFilter, ['全部城市', ...cities]);
  populateSelect(els.artistFilter, ['全部藝人', ...artists]);
  populateSelect(els.typeFilter, ['全部類型', ...types]);
  populateYearSelects(years);

  populateFormSelect(els.formType, '請選擇類型', ['演唱會', '拼盤', 'FanMeeting', 'FanConcert']);
  populateFormSelect(els.formCountry, '請選擇城市', ['台灣台北', '台灣新北', '台灣桃園', '台灣高雄', '韓國首爾']);

  syncLocationOptions();
  updateSummaryText();
}

function updateSummaryText() {
  const topArtist = getTopItems(state.raw, 'Artist');
  const topCity = getTopItems(state.raw, 'Country');

  els.topArtistText.textContent = topArtist.winners.length
    ? `${topArtist.winners.join('、')}（${topArtist.maxCount} 場）`
    : '—';

  els.topCityText.textContent = topCity.winners.length
    ? `${topCity.winners.join('、')}（${topCity.maxCount} 場）`
    : '—';

  els.heroTotalConcerts.textContent = String(state.raw.length);
  els.heroTotalCities.textContent = String(uniqueValues(state.raw.map(r => r.Country)).length);
  els.heroFavoriteCount.textContent = String(state.raw.filter(r => r.favorite).length);
}

function renderStats() {
  updateSummaryText();

  const year = els.statYearFilter.value;
  if (year) {
    const count = state.raw.filter(item => String(parseDate(item.Date)?.getFullYear() || '') === String(year)).length;
    els.yearCountText.textContent = `${year} 年共 ${count} 場`;
  } else {
    els.yearCountText.textContent = '請先選擇年份';
  }
}

/* =========================
   8) 卡片與分頁
   ========================= */
function getTotalPages() {
  return Math.max(1, Math.ceil(state.filtered.length / CONFIG.PAGE_SIZE));
}

function renderResults() {
  const total = state.filtered.length;
  const page = state.currentPage;
  const start = total ? (page - 1) * CONFIG.PAGE_SIZE + 1 : 0;
  const end = Math.min(page * CONFIG.PAGE_SIZE, total);

  els.resultsMeta.textContent = total
    ? `目前顯示 ${start}–${end} / ${total} 場`
    : '目前沒有符合條件的資料';
}

function renderCards() {
  const start = (state.currentPage - 1) * CONFIG.PAGE_SIZE;
  const pageRows = state.filtered.slice(start, start + CONFIG.PAGE_SIZE);

  if (!pageRows.length) {
    els.concertGrid.innerHTML = `
      <div class="concert-empty" style="grid-column: 1 / -1;">
        <strong>找不到符合條件的演唱會</strong>
        <p>可以試著放寬日期、城市、藝人或排序條件。</p>
      </div>
    `;
    return;
  }

  els.concertGrid.innerHTML = pageRows.map((item, index) => {
    const img = item.imgUrlS || item.imgUrlM || '';
    return `
      <article class="concert-card fade-in" style="animation-delay:${index * 40}ms">
        <picture>
          <source media="(min-width: 768px)" srcset="${escapeAttr(item.imgUrlM || img)}">
          <img class="concert-image" src="${escapeAttr(item.imgUrlS || img)}" alt="${escapeAttr(item.ConcertName)}" loading="lazy" />
        </picture>
        <div class="concert-body">
          <h3 class="concert-title">${escapeHtml(item.ConcertName || '未命名演唱會')}</h3>
          <p class="concert-artist">${escapeHtml(item.Artist || '—')}</p>
          <div class="concert-meta">
            ${escapeHtml(formatDatePretty(item.Date))} · ${escapeHtml(item.Country || '—')} / ${escapeHtml(item.Location || '—')}
          </div>
          <div class="concert-submeta">
            <span class="badge">座位：${escapeHtml(item.Seat || '—')}</span>
            <span class="badge">票價：${escapeHtml(String(item.Price || '—'))}</span>
            <span class="badge">${escapeHtml(daysDiffText(item.Date))}</span>
            ${item.favorite ? '<span class="badge badge--favorite">最愛 ❤️</span>' : ''}
          </div>
        </div>
      </article>
    `;
  }).join('');
}

function renderPagination() {
  const totalPages = getTotalPages();
  els.pageInfo.textContent = `${state.currentPage} / ${totalPages}`;
  els.prevPageBtn.disabled = state.currentPage <= 1;
  els.nextPageBtn.disabled = state.currentPage >= totalPages;
}

function renderActiveChips() {
  const chips = [];
  const f = state.filters;

  if (f.search) chips.push(`搜尋：${f.search}`);
  if (f.dateFrom) chips.push(`起：${f.dateFrom}`);
  if (f.dateTo) chips.push(`迄：${f.dateTo}`);
  if (f.city) chips.push(`城市：${f.city}`);
  if (f.artist) chips.push(`藝人：${f.artist}`);
  if (f.type) chips.push(`類型：${f.type}`);
  if (f.year) chips.push(`年份：${f.year}`);
  if (state.selectedCityFromMap) chips.push(`地圖選擇：${state.selectedCityFromMap}`);

  els.activeChips.innerHTML = chips.length
    ? chips.map(text => `<span class="chip is-active">${escapeHtml(text)}</span>`).join('')
    : '<span class="chip">目前沒有套用篩選</span>';
}

function renderAll() {
  state.filtered = getVisibleRows();
  state.currentPage = Math.min(state.currentPage, getTotalPages());
  renderActiveChips();
  renderResults();
  renderCards();
  renderPagination();
  renderStats();
  renderMap();
}

/* =========================
   9) 地圖（Leaflet）
   ========================= */
async function ensureMap() {
  if (state.map) return;

  state.map = L.map('map', {
    scrollWheelZoom: false,
    zoomControl: true
  }).setView(CONFIG.DEFAULT_MAP_CENTER, CONFIG.DEFAULT_MAP_ZOOM);

  L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    attribution: '&copy; OpenStreetMap contributors'
  }).addTo(state.map);

  state.markersLayer = L.layerGroup().addTo(state.map);
}

async function geocodeCity(city) {
  if (!city) return null;
  if (state.geocodeCache[city]) return state.geocodeCache[city];

  try {
    const query = encodeURIComponent(city);
    const url = `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&q=${query}`;
    const response = await axios.get(url, { timeout: 12000 });
    const item = Array.isArray(response.data) ? response.data[0] : null;
    if (!item) return null;

    const result = { lat: Number(item.lat), lng: Number(item.lon) };
    if (Number.isFinite(result.lat) && Number.isFinite(result.lng)) {
      state.geocodeCache[city] = result;
      saveJson(CONFIG.GEOCODE_CACHE_KEY, state.geocodeCache);
      return result;
    }
  } catch (error) {
    console.warn('Geocode failed:', city, error);
  }

  return null;
}

function setSelectedCityFromMap(city) {
  state.selectedCityFromMap = city || '';
  renderCitySummary(state.cityMarkers);
  renderSelectedCityPanel();
  renderActiveChips();
}

function renderCitySummary(markers) {
  const selected = state.selectedCityFromMap;
  els.citySummary.innerHTML = markers
    .sort((a, b) => b.count - a.count)
    .map(m => `
      <button type="button" class="city-card ${m.city === selected ? 'is-active' : ''}" data-city="${escapeAttr(m.city)}">
        ${escapeHtml(m.city)}
        <small>${m.count} 場</small>
      </button>
    `)
    .join('');

  els.citySummary.querySelectorAll('[data-city]').forEach(btn => {
    btn.addEventListener('click', () => {
      const city = btn.getAttribute('data-city') || '';
      setSelectedCityFromMap(city);
    });
  });
}

function renderSelectedCityPanel() {
  const city = state.selectedCityFromMap;
  const rows = city ? state.raw.filter(item => item.Country === city) : [];

  if (!city) {
    els.mapSelectedCityTitle.textContent = '尚未選擇城市';
    els.mapSelectedCityMeta.textContent = '點選地圖上的城市圓圈，右方會顯示該城市所有演唱會卡片。';
    els.mapCityConcertGrid.innerHTML = `
      <div class="map-empty">
        <strong>右側城市卡片區</strong>
        <p>請先點地圖上的圓圈，或直接點上方的城市按鈕。</p>
      </div>
    `;
    return;
  }

  els.mapSelectedCityTitle.textContent = city;
  els.mapSelectedCityMeta.textContent = `共 ${rows.length} 場演唱會`;

  if (!rows.length) {
    els.mapCityConcertGrid.innerHTML = `
      <div class="map-empty">
        <strong>沒有找到這個城市的資料</strong>
        <p>可能是資料目前尚未載入完成，或試算表中的城市名稱尚未對應。</p>
      </div>
    `;
    return;
  }

  const sortedRows = [...rows].sort((a, b) => compareByDate(a, b, true));

  els.mapCityConcertGrid.innerHTML = sortedRows.map(item => {
    const img = item.imgUrlS || item.imgUrlM || '';
    return `
      <article class="city-concert-card">
        <picture>
          <source media="(min-width: 768px)" srcset="${escapeAttr(item.imgUrlM || img)}">
          <img class="city-concert-image" src="${escapeAttr(item.imgUrlS || img)}" alt="${escapeAttr(item.ConcertName)}" loading="lazy" />
        </picture>
        <div class="city-concert-body">
          <h4 class="city-concert-title">${escapeHtml(item.ConcertName || '未命名演唱會')}</h4>
          <p class="city-concert-artist">${escapeHtml(item.Artist || '—')}</p>
          <div class="city-concert-meta">${escapeHtml(formatDatePretty(item.Date))} · ${escapeHtml(item.Location || '—')}</div>
          <div class="city-concert-submeta">
            <span class="badge">座位：${escapeHtml(item.Seat || '—')}</span>
            <span class="badge">票價：${escapeHtml(String(item.Price || '—'))}</span>
            ${item.favorite ? '<span class="badge badge--favorite">最愛 ❤️</span>' : ''}
          </div>
        </div>
      </article>
    `;
  }).join('');
}

async function renderMap() {
  await ensureMap();

  state.markersLayer.clearLayers();
  state.cityMarkers = [];

  const grouped = groupByCity(state.raw);
  const cities = Object.keys(grouped);

  if (!cities.length) {
    els.mapHint.textContent = '目前沒有可顯示的城市資料。';
    els.mapCityConcertGrid.innerHTML = `
      <div class="map-empty">
        <strong>目前沒有可顯示的城市資料</strong>
        <p>請先確認試算表有資料。</p>
      </div>
    `;
    return;
  }

  const markerBounds = [];

  for (const [index, city] of cities.entries()) {
    const items = grouped[city];
    const geo = await geocodeCity(city);
    if (!geo) continue;

    const marker = L.marker([geo.lat, geo.lng], {
      icon: L.divIcon({
        className: '',
        html: `<div class="city-marker"><span>${items.length}</span></div>`,
        iconSize: [46, 46],
        iconAnchor: [23, 23]
      })
    });

    const popupBtnId = `popupFilter-${index}`;
    const popupHtml = `
      <div style="min-width:220px">
        <strong style="display:block;margin-bottom:6px">${escapeHtml(city)}</strong>
        <div style="margin-bottom:8px;color:#8c6d7e">共 ${items.length} 場</div>
        <div style="display:grid;gap:4px">
          ${items.slice(0, 5).map(item => `<div>• ${escapeHtml(item.ConcertName)}</div>`).join('')}
        </div>
        <button id="${popupBtnId}" type="button" style="margin-top:8px;border:0;background:#f48fb1;color:#fff;padding:8px 12px;border-radius:999px;font-weight:700;cursor:pointer">在右側查看</button>
      </div>
    `;

    marker.bindPopup(popupHtml);
    marker.on('click', () => {
      setSelectedCityFromMap(city);
    });

    marker.on('popupopen', () => {
      const btn = document.getElementById(popupBtnId);
      if (btn) {
        btn.addEventListener('click', () => {
          setSelectedCityFromMap(city);
          marker.closePopup();
        });
      }
    });

    marker.addTo(state.markersLayer);
    markerBounds.push([geo.lat, geo.lng]);
    state.cityMarkers.push({ city, count: items.length, lat: geo.lat, lng: geo.lng });
  }

  renderCitySummary(state.cityMarkers);
  renderSelectedCityPanel();

  if (markerBounds.length) {
    const bounds = L.latLngBounds(markerBounds);
    if (bounds.isValid()) {
      state.map.fitBounds(bounds.pad(0.2), { animate: true });
    }
  }

  els.mapHint.textContent = '提示：點任一城市標記，就會在右方顯示該城市的所有場次，完全不影響上方卡片列表。';
}

/* =========================
   10) OAuth / 寫入 Sheet
   ========================= */
function restoreToken() {
  const saved = loadJson(CONFIG.STORAGE_KEYS.token, null);
  if (!saved) return false;
  if (!saved.accessToken || !saved.expiresAt) return false;
  if (Date.now() >= saved.expiresAt - 60_000) {
    localStorage.removeItem(CONFIG.STORAGE_KEYS.token);
    return false;
  }
  state.auth.accessToken = saved.accessToken;
  state.auth.expiresAt = saved.expiresAt;
  return true;
}

function persistToken(tokenResponse) {
  const expiresAt = Date.now() + ((tokenResponse.expires_in || 3600) * 1000);
  const payload = {
    accessToken: tokenResponse.access_token,
    expiresAt
  };
  saveJson(CONFIG.STORAGE_KEYS.token, payload);
  state.auth.accessToken = payload.accessToken;
  state.auth.expiresAt = payload.expiresAt;
}

function isLoggedIn() {
  return Boolean(state.auth.accessToken && Date.now() < state.auth.expiresAt - 60_000);
}

function initTokenClient() {
  if (!window.google?.accounts?.oauth2) return;
  state.tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: CONFIG.GOOGLE_CLIENT_ID,
    scope: 'https://www.googleapis.com/auth/spreadsheets',
    callback: (tokenResponse) => {
      if (tokenResponse.error) {
        setFormMessage(`登入失敗：${tokenResponse.error}`, true);
        return;
      }
      persistToken(tokenResponse);
      setFormMessage('登入成功，可以開始新增資料。', false);
      showAuthState();
      openModalToForm();
    },
  });
}

function requestLogin() {
  if (!CONFIG.GOOGLE_CLIENT_ID || CONFIG.GOOGLE_CLIENT_ID.startsWith('YOUR_')) {
    setFormMessage('請先把 script.js 裡的 GOOGLE_CLIENT_ID 填好。', true);
    return;
  }
  if (!state.tokenClient) initTokenClient();
  if (!state.tokenClient) {
    setFormMessage('Google 登入元件尚未載入完成，請稍後再試。', true);
    return;
  }
  state.tokenClient.requestAccessToken({ prompt: 'consent' });
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
  const response = await axios.post(url, body, { headers, timeout: 15000 });
  return response.data;
}

function collectFormData() {
  return {
    id: String(Date.now()),
    type: els.formType.value.trim(),
    ConcertName: els.formConcertName.value.trim(),
    Artist: els.formArtist.value.trim(),
    Country: els.formCountry.value.trim(),
    Location: els.formLocation.value.trim(),
    Date: els.formDate.value.trim(),
    Price: String(els.formPrice.value).trim(),
    Seat: els.formSeat.value.trim(),
    imgUrlS: els.formImgUrlS.value.trim(),
    imgUrlM: els.formImgUrlM.value.trim(),
    note: els.formNote.value.trim(),
    partner: els.formPartner.value.trim(),
    favorite: els.formFavorite.checked ? 'TRUE' : 'FALSE',
  };
}

function validateForm(payload) {
  const requiredFields = ['type', 'ConcertName', 'Artist', 'Country', 'Location', 'Date', 'Price', 'Seat', 'imgUrlS', 'imgUrlM'];
  for (const key of requiredFields) {
    if (!payload[key]) {
      return `欄位 ${key} 不可空白。`;
    }
  }
  return '';
}

function setSubmitLoading(loading, text) {
  els.submitFormBtn.disabled = loading;
  els.submitFormBtn.textContent = loading ? text : '送出';
  els.clearDraftBtn.disabled = loading;
  els.loginBtn.disabled = loading;
}

function setFormMessage(text, isError = false) {
  els.formMessage.textContent = text;
  els.formMessage.style.color = isError ? '#cf4d79' : 'var(--primary-dark)';
}

function resetForm() {
  els.concertForm.reset();
  syncLocationOptions();
}

function fillFormFromDraft() {
  const draft = loadDraft();
  if (!draft) return;

  els.formType.value = draft.type || '';
  els.formConcertName.value = draft.concertName || '';
  els.formArtist.value = draft.artist || '';
  els.formCountry.value = draft.country || '';
  syncLocationOptions();
  els.formLocation.value = draft.location || '';
  els.formDate.value = draft.date || '';
  els.formPrice.value = draft.price || '';
  els.formSeat.value = draft.seat || '';
  els.formImgUrlS.value = draft.imgUrlS || '';
  els.formImgUrlM.value = draft.imgUrlM || '';
  els.formFavorite.checked = Boolean(draft.favorite);
  els.formNote.value = draft.note || '';
  els.formPartner.value = draft.partner || '';
}

function showAuthState() {
  const loggedIn = isLoggedIn();
  els.authPrompt.hidden = loggedIn;
  if (!loggedIn) {
    setFormMessage('尚未登入，請先按 Google Login。', false);
  } else {
    setFormMessage('已登入，可直接新增資料。', false);
  }
}

function openModalToForm() {
  els.modalOverlay.hidden = false;
  els.modalOverlay.setAttribute('aria-hidden', 'false');
  document.body.style.overflow = 'hidden';
  showAuthState();
  fillFormFromDraft();
  els.formConcertName.focus();
}

function closeModal() {
  els.modalOverlay.hidden = true;
  els.modalOverlay.setAttribute('aria-hidden', 'true');
  document.body.style.overflow = '';
}

function openAddFlow() {
  if (isLoggedIn()) {
    openModalToForm();
  } else {
    openModalToForm();
    showAuthState();
  }
}

function maybeRestoreAuth() {
  restoreToken();
  showAuthState();
}

async function handleSubmit(event) {
  event.preventDefault();
  const payload = collectFormData();
  const error = validateForm(payload);
  if (error) {
    setFormMessage(error, true);
    return;
  }

  try {
    setSubmitLoading(true, '送出中...');
    await appendRowToSheet(payload);
    clearDraft();
    resetForm();
    setFormMessage('已成功寫入 Google Sheet！', false);
    await fetchConcerts();
  } catch (error) {
    console.error(error);
    if (String(error?.response?.status) === '401') {
      localStorage.removeItem(CONFIG.STORAGE_KEYS.token);
      state.auth.accessToken = '';
      state.auth.expiresAt = 0;
      showAuthState();
      setFormMessage('登入已過期，請重新登入後再送出。', true);
    } else {
      setFormMessage('送出失敗，請確認試算表權限、Spreadsheet ID 與表單欄位是否正確。', true);
    }
  } finally {
    setSubmitLoading(false, '送出');
  }
}

/* =========================
   11) 事件綁定
   ========================= */
function bindEvents() {
  els.searchInput.addEventListener('input', () => {
    state.filters.search = els.searchInput.value.trim();
    state.currentPage = 1;
    renderAll();
  });

  els.dateFromInput.addEventListener('change', () => {
    state.filters.dateFrom = els.dateFromInput.value;
    state.currentPage = 1;
    renderAll();
  });

  els.dateToInput.addEventListener('change', () => {
    state.filters.dateTo = els.dateToInput.value;
    state.currentPage = 1;
    renderAll();
  });

  els.cityFilter.addEventListener('change', () => {
    state.filters.city = els.cityFilter.value;
    state.currentPage = 1;
    renderAll();
  });

  els.artistFilter.addEventListener('change', () => {
    state.filters.artist = els.artistFilter.value;
    state.currentPage = 1;
    renderAll();
  });

  els.typeFilter.addEventListener('change', () => {
    state.filters.type = els.typeFilter.value;
    state.currentPage = 1;
    renderAll();
  });

  els.yearFilter.addEventListener('change', () => {
    state.filters.year = els.yearFilter.value;
    state.currentPage = 1;
    renderAll();
  });

  els.statYearFilter.addEventListener('change', () => {
    renderStats();
  });

  els.sortSelect.addEventListener('change', () => {
    state.filters.sort = els.sortSelect.value;
    state.currentPage = 1;
    renderAll();
  });

  els.resetFiltersBtn.addEventListener('click', () => {
    state.filters = {
      search: '',
      dateFrom: '',
      dateTo: '',
      city: '',
      artist: '',
      type: '',
      year: '',
      sort: 'dateDesc',
    };
    state.selectedCityFromMap = '';
    els.searchInput.value = '';
    els.dateFromInput.value = '';
    els.dateToInput.value = '';
    els.cityFilter.value = '';
    els.artistFilter.value = '';
    els.typeFilter.value = '';
    els.yearFilter.value = '';
    els.sortSelect.value = 'dateDesc';
    renderAll();
  });

  els.prevPageBtn.addEventListener('click', () => {
    state.currentPage = Math.max(1, state.currentPage - 1);
    renderAll();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  els.nextPageBtn.addEventListener('click', () => {
    state.currentPage = Math.min(getTotalPages(), state.currentPage + 1);
    renderAll();
    window.scrollTo({ top: 0, behavior: 'smooth' });
  });

  els.addConcertBtn.addEventListener('click', openAddFlow);
  els.closeModalBtn.addEventListener('click', closeModal);
  els.modalOverlay.addEventListener('click', (event) => {
    if (event.target === els.modalOverlay) closeModal();
  });

  els.loginBtn.addEventListener('click', requestLogin);
  els.clearDraftBtn.addEventListener('click', () => {
    resetForm();
    clearDraft();
    setFormMessage('已清空表單。', false);
  });

  els.concertForm.addEventListener('submit', handleSubmit);

  const draftInputs = [
    els.formType, els.formConcertName, els.formArtist, els.formCountry, els.formLocation,
    els.formDate, els.formPrice, els.formSeat, els.formImgUrlS, els.formImgUrlM,
    els.formFavorite, els.formNote, els.formPartner
  ];
  draftInputs.forEach(input => {
    input.addEventListener('input', saveDraft);
    input.addEventListener('change', saveDraft);
  });

    state.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_CLIENT_ID,
      scope: SHEETS_SCOPE,
      callback: (tokenResponse) => {
        if (tokenResponse?.access_token) {
          state.accessToken = tokenResponse.access_token;
          state.tokenReady = true;
          if (window.__authResolve) {
            window.__authResolve(tokenResponse.access_token);
            window.__authResolve = null;
          }
          showToast('Google 授權成功');
        } else if (window.__authReject) {
          window.__authReject(new Error('Google 授權失敗'));
          window.__authReject = null;
        }
      },
    });
  }

  function ensureToken() {
    if (state.accessToken) return Promise.resolve(state.accessToken);
    if (!state.tokenClient) {
      return Promise.reject(new Error('Google OAuth 尚未初始化'));
    }

    return new Promise((resolve, reject) => {
      window.__authResolve = resolve;
      window.__authReject = reject;
      state.tokenClient.requestAccessToken({ prompt: 'consent' });
    });
  }

  function openModal() {
    els.formModal.classList.remove('hidden');
    els.formModal.setAttribute('aria-hidden', 'false');
    restoreDraft();
    els.formConcertName.focus();
  }

  function closeModal() {
    els.formModal.classList.add('hidden');
    els.formModal.setAttribute('aria-hidden', 'true');
  }

  function getFormData() {
    return {
      type: els.formType.value,
      date: els.formDate.value,
      concertName: els.formConcertName.value,
      artist: els.formArtist.value,
      country: els.formCountry.value,
      location: els.formLocation.value,
      price: els.formPrice.value,
      seat: els.formSeat.value,
      imgS: els.formImgS.value,
      imgM: els.formImgM.value,
      note: els.formNote.value,
      partner: els.formPartner.value,
      favorite: els.formFavorite.checked,
    };
  }

  function setFormData(data = {}) {
    els.formType.value = data.type || '';
    els.formDate.value = data.date || '';
    els.formConcertName.value = data.concertName || '';
    els.formArtist.value = data.artist || '';
    els.formCountry.value = data.country || '';
    populateLocationsForCountry(data.country || '');
    els.formLocation.value = data.location || '';
    els.formPrice.value = data.price || '';
    els.formSeat.value = data.seat || '';
    els.formImgS.value = data.imgS || '';
    els.formImgM.value = data.imgM || '';
    els.formNote.value = data.note || '';
    els.formPartner.value = data.partner || '';
    els.formFavorite.checked = !!data.favorite;
  }

  function restoreDraft() {
    const draft = readJson(DRAFT_KEY, null);
    if (draft && Object.keys(draft).length) {
      setFormData(draft);
      showToast('已恢復未送出的內容');
    }
  }

  function clearForm() {
    els.concertForm.reset();
    els.formLocation.innerHTML = '<option value="">請先選擇城市</option>';
    localStorage.removeItem(DRAFT_KEY);
  }

  function populateLocationsForCountry(country) {
  const locations = state.metaRows
    .filter((row) => String(row.Country || '').trim() === String(country || '').trim())
    .map((row) => String(row.Location || '').trim())
    .filter(Boolean);

  const uniqueLocations = [...new Set(locations)];

  if (!uniqueLocations.length) {
    els.formLocation.innerHTML = '<option value="">請先選擇城市</option>';
    return;
  }

  
  fillSelect(els.formLocation, [
    { value: '', label: '請選擇場館' },
    ...uniqueLocations.map(v => ({ value: v, label: v }))
  ]);
  }
  function syncLocationOptions() {
    els.formCountry.addEventListener('change', (e) => {
      populateLocationsForCountry(e.target.value);
      els.formLocation.value = '';
      writeJson(DRAFT_KEY, getFormData());
    });
  }

   async function loadInitialData() {
  setLoading(true);
  try {
    const recordsRes = await axios.get(RECORDS_URL);
    state.allRecords = normalizeRecords(recordsRes.data);

    try {
      const metaRes = await axios.get(META_URL);
      state.metaRows = Array.isArray(metaRes.data) ? metaRes.data : [];
    } catch {
      console.warn('欄位表失敗，但不影響主功能');
      state.metaRows = [];
    }

    prepareFilters();
    renderAll();

  } catch (error) {
    console.error(error);
    showToast('演唱會資料載入失敗');
  } finally {
    setLoading(false);
  }
   }
   
  async function submitForm(e) {
    e.preventDefault();

    const data = getFormData();
    if (!data.type || !data.date || !data.concertName || !data.artist || !data.country || !data.location) {
      showToast('請先填完必填欄位');
      return;
    }

    els.submitFormBtn.disabled = true;
    els.submitFormBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> 送出中';

    try {
      const token = await ensureToken();
      const timestamp = String(Date.now());
      const row = [
        timestamp,
        data.type,
        data.concertName,
        data.artist,
        data.country,
        data.location,
        data.date,
        data.price || '',
        data.seat || '',
        data.imgS || '',
        data.imgM || '',
        data.note || '',
        data.partner || '',
        data.favorite ? 'TRUE' : 'FALSE',
      ];

      await axios.post(
        `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${encodeURIComponent(APPEND_RANGE)}:append`,
        {
          values: [row],
        },
        {
          params: {
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
          },
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-Type': 'application/json',
          },
        }
      );

      localStorage.removeItem(DRAFT_KEY);
      clearForm();
      closeModal();
      showToast('已成功寫入 Google Sheet');

      await loadInitialData();
    } catch (error) {
      console.error(error);
      showToast('送出失敗，請確認 OAuth 與 Sheets API 設定。');
    } finally {
      els.submitFormBtn.disabled = false;
      els.submitFormBtn.innerHTML = '<i class="fa-solid fa-paper-plane"></i> 送出';
    }
  }

  function syncFormFromLocation() {
    els.formCountry.addEventListener('change', (e) => {
      populateLocationsForCountry(e.target.value);
    });
  }

  // ========== 啟動 ==========
  function init() {
    bindEvents();
    syncLocationOptions();
    saveDraft();
  });

  els.formLocation.addEventListener('change', saveDraft);

  window.addEventListener('keydown', (event) => {
    if (event.key === 'Escape' && !els.modalOverlay.hidden) {
      closeModal();
    }
  });
}

/* =========================
   12) 初始化
   ========================= */
async function init() {
  bindEvents();
  maybeRestoreAuth();
  await fetchConcerts();
}

document.addEventListener('DOMContentLoaded', init);
