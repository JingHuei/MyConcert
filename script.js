/* =========================================================
   Lucy 的演唱會紀錄網站
   - 首頁：公開 Google Sheet JSON
   - 新增：Google OAuth + Sheets API append
   - 地圖：Leaflet
   ========================================================= */

(() => {
  'use strict';

  // ========== 基本設定 ==========
  const GOOGLE_CLIENT_ID = '562464022417-8c2sckejaft6de7ch8kejqomm1fbi0ga.apps.googleusercontent.com';
  const SPREADSHEET_ID = '1tfNjim8BbKmv3KVNIQrRvx2T6W8YtfjyJCr1_wXDcrA';
  const RECORD_SHEET_NAME = '演唱會紀錄';
  const META_SHEET_NAME = '欄位表';
  const OPEN_SHEET_BASE = `https://opensheet.elk.sh/${SPREADSHEET_ID}`;
  const RECORDS_URL = `${OPEN_SHEET_BASE}/${encodeURIComponent(RECORD_SHEET_NAME)}`;
  const META_URL = `${OPEN_SHEET_BASE}/${encodeURIComponent(META_SHEET_NAME)}`;
  const APPEND_RANGE = `${RECORD_SHEET_NAME}!A:N`;
  const SHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets';
  const PAGE_SIZE = 10;
  const DRAFT_KEY = 'lucy-concert-draft-v1';
  const CITY_CACHE_KEY = 'lucy-city-geocode-cache-v1';

  // ========== DOM 取得 ==========
  const els = {
    loadingState: document.getElementById('loadingState'),
    emptyState: document.getElementById('emptyState'),
    cardsGrid: document.getElementById('cardsGrid'),
    pageInfo: document.getElementById('pageInfo'),
    prevPageBtn: document.getElementById('prevPageBtn'),
    nextPageBtn: document.getElementById('nextPageBtn'),
    searchInput: document.getElementById('searchInput'),
    sortSelect: document.getElementById('sortSelect'),
    typeFilter: document.getElementById('typeFilter'),
    cityFilter: document.getElementById('cityFilter'),
    artistFilter: document.getElementById('artistFilter'),
    dateFromFilter: document.getElementById('dateFromFilter'),
    dateToFilter: document.getElementById('dateToFilter'),
    clearFiltersBtn: document.getElementById('clearFiltersBtn'),
    statTotal: document.getElementById('stat-total'),
    statCityCount: document.getElementById('stat-city-count'),
    statTopArtist: document.getElementById('stat-top-artist'),
    topArtistBox: document.getElementById('topArtistBox'),
    topCityBox: document.getElementById('topCityBox'),
    yearSelect: document.getElementById('yearSelect'),
    yearCountBox: document.getElementById('yearCountBox'),
    cityLegend: document.getElementById('cityLegend'),
    mapContainer: document.getElementById('cityMap'),
    openFormBtn: document.getElementById('openFormBtn'),
    closeFormBtn: document.getElementById('closeFormBtn'),
    formModal: document.getElementById('formModal'),
    concertForm: document.getElementById('concertForm'),
    clearFormBtn: document.getElementById('clearFormBtn'),
    submitFormBtn: document.getElementById('submitFormBtn'),
    formHint: document.getElementById('formHint'),
    toast: document.getElementById('toast'),
    formType: document.getElementById('formType'),
    formDate: document.getElementById('formDate'),
    formConcertName: document.getElementById('formConcertName'),
    formArtist: document.getElementById('formArtist'),
    formCountry: document.getElementById('formCountry'),
    formLocation: document.getElementById('formLocation'),
    formPrice: document.getElementById('formPrice'),
    formSeat: document.getElementById('formSeat'),
    formImgS: document.getElementById('formImgS'),
    formImgM: document.getElementById('formImgM'),
    formNote: document.getElementById('formNote'),
    formPartner: document.getElementById('formPartner'),
    formFavorite: document.getElementById('formFavorite'),
    statTopArtistBox: document.getElementById('stat-top-artist'),
  };

  // ========== 狀態 ==========
  const state = {
    allRecords: [],
    metaRows: [],
    filteredRecords: [],
    currentPage: 1,
    pageSize: PAGE_SIZE,
    sortMode: 'date_desc',
    filters: {
      keyword: '',
      type: '',
      city: '',
      artist: '',
      dateFrom: '',
      dateTo: '',
    },
    map: null,
    cityMarkers: [],
    cityGeoCache: readJson(CITY_CACHE_KEY, {}),
    tokenClient: null,
    accessToken: '',
    tokenReady: false,
    loadingCities: false,
  };

  const strokeCollator = new Intl.Collator('zh-Hant-u-co-stroke');
  const numberFmt = new Intl.NumberFormat('zh-TW');
  const dateFmt = new Intl.DateTimeFormat('zh-TW', {
    year: 'numeric', month: 'long', day: 'numeric'
  });

  // ========== 工具函式 ==========
  function readJson(key, fallback) {
    try {
      const raw = localStorage.getItem(key);
      return raw ? JSON.parse(raw) : fallback;
    } catch {
      return fallback;
    }
  }

  function writeJson(key, value) {
    try {
      localStorage.setItem(key, JSON.stringify(value));
    } catch {
      // ignore
    }
  }

  function escapeHtml(input = '') {
    return String(input)
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#39;');
  }

  function showToast(message) {
    els.toast.textContent = message;
    els.toast.classList.remove('hidden');
    clearTimeout(showToast._timer);
    showToast._timer = setTimeout(() => {
      els.toast.classList.add('hidden');
    }, 2600);
  }

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
      const value = entry?.value ?? '';
      const label = entry?.label ?? value;
      return `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`;
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

    if (type) result = result.filter((r) => r.type === type);
    if (city) result = result.filter((r) => r.Country === city);
    if (artist) result = result.filter((r) => r.Artist === artist);

    if (dateFrom) {
      const from = parseDate(dateFrom);
      result = result.filter((r) => r.dateObj && from && r.dateObj >= from);
    }

    if (dateTo) {
      const to = parseDate(dateTo);
      if (to) {
        to.setHours(23, 59, 59, 999);
        result = result.filter((r) => r.dateObj && r.dateObj <= to);
      }
    }

    result = sortRecords(result, state.sortMode);
    state.filteredRecords = result;
    const maxPage = Math.max(1, Math.ceil(result.length / state.pageSize));
    if (state.currentPage > maxPage) state.currentPage = maxPage;
  }

  function sortRecords(records, mode) {
    const items = [...records];
    switch (mode) {
      case 'date_asc':
        return items.sort((a, b) => (a.dateObj?.getTime() || 0) - (b.dateObj?.getTime() || 0));
      case 'artist_asc':
        return items.sort((a, b) => safeStrokeCompare(a.Artist, b.Artist));
      case 'artist_desc':
        return items.sort((a, b) => safeStrokeCompare(b.Artist, a.Artist));
      case 'city_asc':
        return items.sort((a, b) => safeStrokeCompare(a.Country, b.Country));
      case 'city_desc':
        return items.sort((a, b) => safeStrokeCompare(b.Country, a.Country));
      case 'date_desc':
      default:
        return items.sort((a, b) => (b.dateObj?.getTime() || 0) - (a.dateObj?.getTime() || 0));
    }
  }

  // ========== 渲染 ==========
  function renderAll() {
    applyFilters();
    renderCards();
    renderPagination();
    renderStats();
    renderCityLegend();
    renderYearStats();
    renderCityMarkers();
  }

  function renderCards() {
    const start = (state.currentPage - 1) * state.pageSize;
    const pageItems = state.filteredRecords.slice(start, start + state.pageSize);

    setEmpty(state.filteredRecords.length === 0);
    els.cardsGrid.innerHTML = pageItems.map(renderCard).join('');
    applyRevealObservers();
  }

  function renderCard(record) {
    const imageHtml = `
      <picture>
        <source media="(min-width: 768px)" srcset="${escapeHtml(record.imgUrlM || record.imgUrlS || '')}">
        <img src="${escapeHtml(record.imgUrlS || record.imgUrlM || '')}" alt="${escapeHtml(record.ConcertName || '演唱會照片')}" loading="lazy" onerror="this.onerror=null;this.src='data:image/svg+xml;charset=UTF-8,${encodeURIComponent(fallbackImage())}'">
      </picture>
    `;

    return `
      <article class="concert-card reveal">
        <div class="concert-media">
          ${record.favorite ? `<div class="favorite-badge"><i class="fa-solid fa-heart"></i> 最愛</div>` : ''}
          ${imageHtml}
        </div>
        <div class="concert-body">
          <h3 class="concert-title">${escapeHtml(record.ConcertName || '-')}</h3>
          <p class="concert-artist">${escapeHtml(record.Artist || '-')}</p>

          <div class="meta-grid">
            <div class="meta-line">
              <i class="fa-regular fa-calendar"></i>
              <span>${escapeHtml(formatDate(record.Date))}</span>
            </div>
            <div class="meta-line">
              <i class="fa-solid fa-location-dot"></i>
              <span>${escapeHtml(record.Country || '-')}｜${escapeHtml(record.Location || '-')}</span>
            </div>
            <div class="meta-line">
              <i class="fa-solid fa-chair"></i>
              <span>${escapeHtml(record.Seat || '-')}${record.Seat ? ' 座位' : ''}</span>
            </div>
            <div class="meta-line">
              <i class="fa-solid fa-tag"></i>
              <span>${record.Price ? `NT$ ${escapeHtml(numberFmt.format(record.Price))}` : '-'}</span>
            </div>
            <div class="meta-line">
              <i class="fa-solid fa-hourglass-half"></i>
              <span>${escapeHtml(record.daysText)}</span>
            </div>
          </div>

          <div class="concert-foot">
            <span class="tag-pill">${escapeHtml(typeLabel(record.type))}</span>
            ${record.partner ? `<span class="tag-pill">夥伴：${escapeHtml(record.partner)}</span>` : ''}
          </div>
        </div>
      </article>
    `;
  }

  function fallbackImage() {
    return `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 800">
        <defs>
          <linearGradient id="g" x1="0" x2="1" y1="0" y2="1">
            <stop offset="0%" stop-color="#ffd0df"/>
            <stop offset="100%" stop-color="#fff7fb"/>
          </linearGradient>
        </defs>
        <rect width="1200" height="800" fill="url(#g)"/>
        <circle cx="980" cy="120" r="90" fill="rgba(255,255,255,0.45)"/>
        <circle cx="140" cy="680" r="160" fill="rgba(255,255,255,0.36)"/>
        <text x="60" y="410" font-size="42" font-family="Arial" fill="#b94d7a">Lucy Concert Memories</text>
      </svg>
    `;
  }

  function renderPagination() {
    const total = state.filteredRecords.length;
    const totalPages = Math.max(1, Math.ceil(total / state.pageSize));
    els.pageInfo.textContent = `第 ${total ? state.currentPage : 0} / ${totalPages} 頁`;
    els.prevPageBtn.disabled = state.currentPage <= 1;
    els.nextPageBtn.disabled = state.currentPage >= totalPages || total === 0;
  }

  function renderStats() {
    const total = state.allRecords.length;
    const cityGroups = groupBy(state.allRecords, (r) => cityKey(r.Country));
    const artistGroups = groupBy(state.allRecords, (r) => String(r.Artist || '').trim());

    const topArtist = topEntry(artistGroups);
    const topCity = topEntry(cityGroups);

    els.statTotal.textContent = numberFmt.format(total);
    els.statCityCount.textContent = numberFmt.format(Object.keys(cityGroups).filter(Boolean).length);
    els.statTopArtist.textContent = topArtist ? `${topArtist.key} (${topArtist.count})` : '-';
    els.topArtistBox.textContent = topArtist ? `${topArtist.key} · ${topArtist.count} 場` : '-';
    els.topCityBox.textContent = topCity ? `${topCity.key} · ${topCity.count} 場` : '-';
  }

  function topEntry(grouped) {
    const entries = Object.entries(grouped)
      .filter(([key]) => String(key).trim() !== '')
      .map(([key, arr]) => ({ key, count: arr.length }))
      .sort((a, b) => b.count - a.count || safeStrokeCompare(a.key, b.key));
    return entries[0] || null;
  }

  function renderYearStats() {
    const selectedYear = els.yearSelect.value;
    const total = state.allRecords.filter((r) => r.year === selectedYear).length;
    els.yearCountBox.textContent = selectedYear ? `${selectedYear} 年共 ${numberFmt.format(total)} 場` : '-';
  }

  function renderCityLegend() {
    const cityGroups = groupBy(state.allRecords, (r) => cityKey(r.Country));
    const entries = Object.entries(cityGroups)
      .filter(([city]) => city)
      .sort((a, b) => b[1].length - a[1].length);

    els.cityLegend.innerHTML = entries.length
      ? `目前共記錄 <strong>${entries.length}</strong> 個城市。`
      : '目前尚無城市資料。';
  }

  // ========== 地圖 ==========
  function initMapIfNeeded() {
    if (state.map) return;

    state.map = L.map('cityMap', {
      scrollWheelZoom: false,
      zoomControl: true,
    }).setView([25.033, 121.5654], 10); // 預設雙北

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      attribution: '&copy; OpenStreetMap contributors',
      maxZoom: 19,
    }).addTo(state.map);
  }

  async function warmCityGeocodes() {
    if (state.loadingCities) return;
    state.loadingCities = true;
    const cities = [...new Set(state.allRecords.map((r) => cityKey(r.Country)).filter(Boolean))];
    const tasks = cities.map((city) => ensureCityCoords(city));
    await Promise.allSettled(tasks);
    state.loadingCities = false;
  }

  async function ensureCityCoords(city) {
    if (!city) return null;
    if (state.cityGeoCache[city]) return state.cityGeoCache[city];

    const query = buildCityQuery(city);
    try {
      const res = await axios.get('https://nominatim.openstreetmap.org/search', {
        params: {
          format: 'jsonv2',
          limit: 1,
          q: query,
        },
      });
      const hit = Array.isArray(res.data) ? res.data[0] : null;
      if (hit?.lat && hit?.lon) {
        const coords = [Number(hit.lat), Number(hit.lon)];
        state.cityGeoCache[city] = coords;
        writeJson(CITY_CACHE_KEY, state.cityGeoCache);
        return coords;
      }
    } catch (error) {
      console.warn('geocode fail', city, error);
    }

    const fallback = fallbackCityCoords(city);
    if (fallback) {
      state.cityGeoCache[city] = fallback;
      writeJson(CITY_CACHE_KEY, state.cityGeoCache);
      return fallback;
    }
    return null;
  }

  function buildCityQuery(city) {
    const text = String(city || '').trim();
    if (text.includes('台北') || text.includes('臺北')) return 'Taipei, Taiwan';
    if (text.includes('桃園')) return 'Taoyuan, Taiwan';
    if (text.includes('高雄')) return 'Kaohsiung, Taiwan';
    if (text.includes('首爾')) return 'Seoul, South Korea';
    if (text.includes('東京')) return 'Tokyo, Japan';
    return text;
  }

  function fallbackCityCoords(city) {
    const map = {
      '台灣台北': [25.0330, 121.5654],
      '台北': [25.0330, 121.5654],
      '台灣桃園': [24.9936, 121.3010],
      '桃園': [24.9936, 121.3010],
      '台灣高雄': [22.6273, 120.3014],
      '高雄': [22.6273, 120.3014],
      '韓國首爾': [37.5665, 126.9780],
      '首爾': [37.5665, 126.9780],
      '東京': [35.6762, 139.6503],
    };
    return map[city] || null;
  }

  function renderCityMarkers() {
    if (!state.map) return;

    state.cityMarkers.forEach((marker) => state.map.removeLayer(marker));
    state.cityMarkers = [];

    const cityGroups = groupBy(state.allRecords, (r) => cityKey(r.Country));
    const cities = Object.entries(cityGroups).filter(([city]) => city);

    const bounds = [];
    cities.forEach(([city, records]) => {
      const coords = state.cityGeoCache[city];
      if (!coords) return;

      const marker = L.marker(coords).addTo(state.map);
      const html = `
        <div>
          <strong>${escapeHtml(city)}</strong><br>
          <span>${records.length} 場演唱會</span><br>
          <button type="button" class="popup-filter-btn" data-city="${escapeHtml(city)}">
            只看這個城市
          </button>
        </div>
      `;
      marker.bindPopup(html);
      marker.on('popupopen', () => {
        const popupEl = marker.getPopup()?.getElement();
        const btn = popupEl?.querySelector('.popup-filter-btn');
        if (btn && !btn.dataset.bound) {
          btn.dataset.bound = '1';
          btn.addEventListener('click', () => {
            state.filters.city = city;
            els.cityFilter.value = city;
            state.currentPage = 1;
            renderAll();
            showToast(`已篩選：${city}`);
          });
        }
      });

      state.cityMarkers.push(marker);
      bounds.push(coords);
    });

    if (bounds.length) {
      if (bounds.length === 1) {
        state.map.setView(bounds[0], 10);
      } else {
        state.map.fitBounds(bounds, { padding: [28, 28] });
      }
    } else {
      state.map.setView([25.033, 121.5654], 10);
    }
  }

  // ========== 互動 ==========
  function bindEvents() {
    els.searchInput.addEventListener('input', debounce((e) => {
      state.filters.keyword = e.target.value.trim();
      state.currentPage = 1;
      renderAll();
    }, 180));

    els.sortSelect.addEventListener('change', (e) => {
      state.sortMode = e.target.value;
      state.currentPage = 1;
      renderAll();
    });

    els.typeFilter.addEventListener('change', (e) => {
      state.filters.type = e.target.value;
      state.currentPage = 1;
      renderAll();
    });

    els.cityFilter.addEventListener('change', (e) => {
      state.filters.city = e.target.value;
      state.currentPage = 1;
      renderAll();
    });

    els.artistFilter.addEventListener('change', (e) => {
      state.filters.artist = e.target.value;
      state.currentPage = 1;
      renderAll();
    });

    els.dateFromFilter.addEventListener('change', (e) => {
      state.filters.dateFrom = e.target.value;
      state.currentPage = 1;
      renderAll();
    });

    els.dateToFilter.addEventListener('change', (e) => {
      state.filters.dateTo = e.target.value;
      state.currentPage = 1;
      renderAll();
    });

    els.clearFiltersBtn.addEventListener('click', () => {
      state.filters = { keyword: '', type: '', city: '', artist: '', dateFrom: '', dateTo: '' };
      state.sortMode = 'date_desc';
      state.currentPage = 1;
      els.searchInput.value = '';
      els.sortSelect.value = 'date_desc';
      els.typeFilter.value = '';
      els.cityFilter.value = '';
      els.artistFilter.value = '';
      els.dateFromFilter.value = '';
      els.dateToFilter.value = '';
      renderAll();
      showToast('篩選已清除');
    });

    els.prevPageBtn.addEventListener('click', () => {
      if (state.currentPage > 1) {
        state.currentPage -= 1;
        renderCards();
        renderPagination();
      }
    });

    els.nextPageBtn.addEventListener('click', () => {
      const totalPages = Math.max(1, Math.ceil(state.filteredRecords.length / state.pageSize));
      if (state.currentPage < totalPages) {
        state.currentPage += 1;
        renderCards();
        renderPagination();
      }
    });

    els.yearSelect.addEventListener('change', renderYearStats);

    els.openFormBtn.addEventListener('click', async () => {
      try {
        await ensureToken();
        openModal();
      } catch (error) {
        console.error(error);
      }
    });

    els.closeFormBtn.addEventListener('click', closeModal);

    els.formModal.addEventListener('click', (e) => {
      if (e.target === els.formModal) closeModal();
    });

    els.clearFormBtn.addEventListener('click', () => {
      clearForm();
      showToast('表單已清空');
    });

    els.concertForm.addEventListener('submit', submitForm);

    bindDraftEvents();
    bindRevealObserver();
  }

  function bindDraftEvents() {
    const draftFields = [
      els.formType, els.formDate, els.formConcertName, els.formArtist, els.formCountry,
      els.formLocation, els.formPrice, els.formSeat, els.formImgS, els.formImgM,
      els.formNote, els.formPartner, els.formFavorite,
    ];

    const saveDraft = () => {
      const draft = getFormData();
      writeJson(DRAFT_KEY, draft);
    };

    draftFields.forEach((el) => {
      const evt = el.type === 'checkbox' ? 'change' : 'input';
      el.addEventListener(evt, saveDraft);
    });
  }

  function bindRevealObserver() {
    applyRevealObservers();
  }

  function applyRevealObservers() {
    const items = document.querySelectorAll('.reveal:not(.in-view)');
    if (!('IntersectionObserver' in window)) {
      items.forEach((el) => el.classList.add('in-view'));
      return;
    }
    const observer = new IntersectionObserver((entries, obs) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add('in-view');
          obs.unobserve(entry.target);
        }
      });
    }, { threshold: 0.12 });

    items.forEach((el) => observer.observe(el));
  }

  function debounce(fn, wait = 200) {
    let timer = null;
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => fn(...args), wait);
    };
  }

  // ========== Google OAuth / Sheets 寫入 ==========
  function initGoogleAuth() {
    if (!window.google?.accounts?.oauth2) {
      setTimeout(initGoogleAuth, 250);
      return;
    }

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
    fillSelect(els.formLocation, ['','請選擇場館', ...uniqueLocations.map((v) => [v, v])]);
  }

  function syncLocationOptions() {
    els.formCountry.addEventListener('change', (e) => {
      populateLocationsForCountry(e.target.value);
      els.formLocation.value = '';
      writeJson(DRAFT_KEY, getFormData());
    });
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
    initGoogleAuth();
    loadInitialData();
    renderYearStats();
  }

  document.addEventListener('DOMContentLoaded', init);
})();
