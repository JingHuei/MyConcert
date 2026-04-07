/* global axios, google, L */
(() => {
  'use strict';

  const CONFIG = {
    GOOGLE_CLIENT_ID: '562464022417-8c2sckejaft6de7ch8kejqomm1fbi0ga.apps.googleusercontent.com',
    SPREADSHEET_ID: '1tfNjim8BbKmv3KVNIQrRvx2T6W8YtfjyJCr1_wXDcrA',
    SHEETS: {
      RECORDS: '演唱會紀錄',
      FIELDS: '欄位表',
    },
    RANGES: {
      RECORDS: '演唱會紀錄!A1:N',
      FIELDS: '欄位表!A1:C',
    },
    PAGE_SIZE: 10,
    DRAFT_KEY: 'concert_review_draft_v2',
    AUTH_KEY: 'concert_review_auth_v2',
    CITY_CACHE_KEY: 'concert_review_city_cache_v2',
    ENABLE_GEOCODING: true,
    SCOPES: [
      'openid',
      'email',
      'profile',
      'https://www.googleapis.com/auth/spreadsheets'
    ].join(' ')
  };

  const DEFAULT_CITY_GEO = {
    '台北': [25.033964, 121.564468],
    '台北市': [25.033964, 121.564468],
    '台灣台北': [25.033964, 121.564468],
    'Taipei': [25.033964, 121.564468],
    '高雄': [22.627278, 120.301435],
    '高雄市': [22.627278, 120.301435],
    '台灣高雄': [22.627278, 120.301435],
    'Kaohsiung': [22.627278, 120.301435],
    '新北': [25.012001, 121.465447],
    '新北市': [25.012001, 121.465447],
    '桃園': [24.993628, 121.300979],
    '桃園市': [24.993628, 121.300979],
    '首爾': [37.566535, 126.977969],
    '首爾市': [37.566535, 126.977969],
    '韓國首爾': [37.566535, 126.977969],
    'Seoul': [37.566535, 126.977969],
    '서울': [37.566535, 126.977969],
    '東京': [35.6762, 139.6503],
    '東京都': [35.6762, 139.6503],
    '日本東京': [35.6762, 139.6503],
    'Tokyo': [35.6762, 139.6503],
    '大阪': [34.693738, 135.502165],
    '大阪市': [34.693738, 135.502165],
    'Osaka': [34.693738, 135.502165],
    '新加坡': [1.352083, 103.819836],
    'Singapore': [1.352083, 103.819836],
  };

  const CITY_ALIASES = {
    '台北市': '台北',
    '台灣台北': '台北',
    '高雄市': '高雄',
    '台灣高雄': '高雄',
    '新北市': '新北',
    '桃園市': '桃園',
    '首爾市': '首爾',
    '韓國首爾': '首爾',
    '東京都': '東京',
    '日本東京': '東京',
    '大阪市': '大阪',
  };

  const els = {};
  const state = {
    token: '',
    user: null,
    tokenClient: null,
    records: [],
    fieldRows: [],
    filters: {
      search: '',
      type: '',
      country: '',
      date: ''
    },
    sortMode: 'date-desc',
    year: new Date().getFullYear(),
    page: 1,
    loading: false,
    draft: loadDraft(),
    map: null,
    markerLayer: null,
    cityMarkerIndex: new Map(),
    cityGeoCache: loadCityGeoCache(),
  };

  const collator = new Intl.Collator('zh-Hant-u-co-stroke', { numeric: true, sensitivity: 'base' });

  document.addEventListener('DOMContentLoaded', init);

  async function init() {
    cacheElements();
    bindEvents();
    bindDraftInputs();
    renderStaticYear();
    restoreUIState();
    initIntersectionObserver();
    initMap();
    await initGoogleIdentity();
    syncAuthUI();
    if (state.token) {
      await loadEverything();
    } else {
      renderAll();
      showStatus('待登入', '請先登入 Google 才能讀取資料');
    }
  }

  function cacheElements() {
    const ids = [
      'userAvatarWrap','userAvatar','userLabel','tokenState','btnLogin','btnLogout',
      'searchInput','filterType','filterCountry','filterDate','sortSelect','yearInput','yearLabel',
      'topArtist','topArtistMeta','topCountry','topCountryMeta','yearCount','syncStatus','syncHint',
      'totalCountPill','filteredCountPill','summaryTotal','summaryFavorite','summaryCities',
      'cityLegend','loadingState','emptyState','cardGrid','pageInfo','displayCount','displayPageRate',
      'btnPrevPage','btnNextPage','btnReload','btnOpenForm','formBackdrop','btnCloseForm',
      'entryForm','btnClearDraft','entryType','entryConcertName','entryArtist','entryCountry',
      'entryLocation','entryDate','entryPrice','entrySeat','entryImgS','entryImgM','entryNote',
      'entryPartner','entryFavorite','btnSubmitEntry','toastHost','mapSummary','btnResetMap',
      'leafletMap'
    ];
    ids.forEach((id) => { els[id] = document.getElementById(id); });
  }

  function bindEvents() {
    els.btnLogin.addEventListener('click', login);
    els.btnLogout.addEventListener('click', logout);

    els.searchInput.addEventListener('input', onFilterChange);
    els.filterType.addEventListener('change', onFilterChange);
    els.filterCountry.addEventListener('change', onFilterChange);
    els.filterDate.addEventListener('change', onFilterChange);
    els.sortSelect.addEventListener('change', onSortChange);
    els.yearInput.addEventListener('input', onYearChange);

    els.btnPrevPage.addEventListener('click', () => changePage(-1));
    els.btnNextPage.addEventListener('click', () => changePage(1));
    els.btnReload.addEventListener('click', async () => {
      if (!state.token) {
        toast('請先登入後再重新整理');
        return;
      }
      await loadEverything();
    });

    els.btnOpenForm.addEventListener('click', openForm);
    els.btnCloseForm.addEventListener('click', () => closeForm(true));
    els.formBackdrop.addEventListener('click', (e) => {
      if (e.target === els.formBackdrop) closeForm(true);
    });

    els.entryForm.addEventListener('submit', submitEntry);
    els.btnClearDraft.addEventListener('click', clearDraftAndForm);
    els.entryCountry.addEventListener('change', syncLocationsByCountry);
    els.btnResetMap.addEventListener('click', resetMapView);

    window.addEventListener('beforeunload', saveDraftFromForm);
    window.addEventListener('resize', debounce(() => {
      if (state.map) state.map.invalidateSize();
    }, 150));

    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape' && !els.formBackdrop.classList.contains('d-none')) {
        closeForm(true);
      }
    });
  }

  function bindDraftInputs() {
    [
      els.entryType, els.entryConcertName, els.entryArtist, els.entryCountry, els.entryLocation,
      els.entryDate, els.entryPrice, els.entrySeat, els.entryImgS, els.entryImgM,
      els.entryNote, els.entryPartner, els.entryFavorite
    ].forEach((el) => {
      el.addEventListener('input', saveDraftFromForm);
      el.addEventListener('change', saveDraftFromForm);
    });
  }

  function restoreUIState() {
    els.sortSelect.value = localStorage.getItem('concert_review_sort') || 'date-desc';
    els.searchInput.value = localStorage.getItem('concert_review_search') || '';
    els.filterType.value = localStorage.getItem('concert_review_filter_type') || '';
    els.filterCountry.value = localStorage.getItem('concert_review_filter_country') || '';
    els.filterDate.value = localStorage.getItem('concert_review_filter_date') || '';
    const savedYear = localStorage.getItem('concert_review_year');
    if (savedYear) {
      state.year = parseInt(savedYear, 10) || state.year;
    }
    els.yearInput.value = String(state.year);
    els.yearLabel.textContent = String(state.year);
    state.sortMode = els.sortSelect.value;
    state.filters.search = els.searchInput.value.trim();
    state.filters.type = els.filterType.value;
    state.filters.country = els.filterCountry.value;
    state.filters.date = els.filterDate.value;
    if (state.draft) {
      fillFormFromDraft(state.draft);
    }
  }

  function renderStaticYear() {
    els.yearInput.value = String(state.year);
    els.yearLabel.textContent = String(state.year);
  }

  function initIntersectionObserver() {
    const observer = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add('is-visible');
        }
      });
    }, { threshold: 0.12 });

    document.querySelectorAll('.fade-up').forEach((el) => observer.observe(el));
  }

  function initMap() {
    if (!window.L || !els.leafletMap) return;

    state.map = L.map('leafletMap', {
      zoomControl: true,
      scrollWheelZoom: false,
      worldCopyJump: true,
    }).setView([25.033964, 121.564468], 3);

    L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
      maxZoom: 19,
      attribution: '&copy; OpenStreetMap contributors'
    }).addTo(state.map);

    state.markerLayer = L.layerGroup().addTo(state.map);

    setTimeout(() => {
      if (state.map) state.map.invalidateSize();
    }, 200);
  }

  function resetMapView() {
    if (!state.map) return;
    if (state.cityMarkerIndex.size) {
      const latlngs = [...state.cityMarkerIndex.values()].map((item) => item.marker.getLatLng());
      const bounds = L.latLngBounds(latlngs);
      state.map.fitBounds(bounds, { padding: [30, 30] });
    } else {
      state.map.setView([25.033964, 121.564468], 3);
    }
  }

  async function initGoogleIdentity() {
    await waitForGoogle();
    state.tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: CONFIG.GOOGLE_CLIENT_ID,
      scope: CONFIG.SCOPES,
      callback: async (response) => {
        if (response.error) {
          showStatus('登入失敗', response.error);
          toast('Google 登入失敗');
          return;
        }
        state.token = response.access_token;
        localStorage.setItem(CONFIG.AUTH_KEY, state.token);
        await loadUserProfile();
        syncAuthUI();
        await loadEverything();
      }
    });

    const saved = localStorage.getItem(CONFIG.AUTH_KEY);
    if (saved) {
      state.token = saved;
      await loadUserProfile().catch(() => {});
    }
  }

  function waitForGoogle() {
    return new Promise((resolve) => {
      const timer = setInterval(() => {
        if (window.google?.accounts?.oauth2) {
          clearInterval(timer);
          resolve();
        }
      }, 50);
    });
  }

  async function login() {
    if (!state.tokenClient) return;
    state.tokenClient.requestAccessToken({ prompt: state.token ? '' : 'consent' });
  }

  async function logout() {
    if (!state.token) return;
    try {
      await google.accounts.oauth2.revoke(state.token, () => {});
    } catch (error) {
      console.warn(error);
    }
    state.token = '';
    state.user = null;
    localStorage.removeItem(CONFIG.AUTH_KEY);
    syncAuthUI();
    showStatus('已登出', '請重新登入以載入資料');
    state.records = [];
    state.fieldRows = [];
    renderAll();
  }

  function syncAuthUI() {
    const loggedIn = Boolean(state.token);
    els.btnLogin.disabled = loggedIn;
    els.btnLogout.disabled = !loggedIn;
    els.tokenState.textContent = loggedIn ? '已授權，可讀寫試算表' : '請登入後載入資料';
    els.userLabel.textContent = state.user?.name || (loggedIn ? '已登入 Google' : '尚未登入');
    if (state.user?.picture) {
      els.userAvatar.src = state.user.picture;
      els.userAvatar.classList.remove('d-none');
      els.userAvatarWrap.querySelector('.avatar-placeholder')?.classList.add('d-none');
    } else {
      els.userAvatar.classList.add('d-none');
      const placeholder = els.userAvatarWrap.querySelector('.avatar-placeholder');
      if (placeholder) placeholder.classList.remove('d-none');
    }
  }

  async function loadUserProfile() {
    const { data } = await axios.get('https://www.googleapis.com/oauth2/v3/userinfo', {
      headers: { Authorization: `Bearer ${state.token}` }
    });
    state.user = data;
  }

  async function loadEverything() {
    try {
      setLoading(true);
      showStatus('同步中', '正在讀取 Google Sheet');
      await Promise.all([loadFieldRows(), loadRecords()]);
      await renderAll();
      showStatus('已同步', '資料已成功載入');
    } catch (error) {
      console.error(error);
      const msg = error?.response?.data?.error?.message || error.message || '讀取失敗';
      showStatus('同步失敗', msg);
      toast(`載入失敗：${msg}`);
      if (error?.response?.status === 401) {
        localStorage.removeItem(CONFIG.AUTH_KEY);
        state.token = '';
        state.user = null;
        syncAuthUI();
      }
    } finally {
      setLoading(false);
    }
  }

  async function loadRecords() {
    const rows = await fetchSheetValues(CONFIG.RANGES.RECORDS);
    state.records = rowsToObjects(rows, true)
      .map(normalizeRecord)
      .filter((row) => row.id || row.ConcertName || row.Artist);
  }

  async function loadFieldRows() {
    const rows = await fetchSheetValues(CONFIG.RANGES.FIELDS);
    const objects = rowsToObjects(rows, false);
    state.fieldRows = objects.map((row) => ({
      type: row.Type || row.type || '',
      country: row.Country || row.country || '',
      location: row.Location || row.location || ''
    })).filter((row) => row.type || row.country || row.location);
    populateFilterOptions();
    populateFormOptions();
  }

  async function fetchSheetValues(range) {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(CONFIG.SPREADSHEET_ID)}/values/${encodeURIComponent(range)}?majorDimension=ROWS`;
    const { data } = await axios.get(url, { headers: { Authorization: `Bearer ${state.token}` } });
    return data.values || [];
  }

  function rowsToObjects(rows, hasHeader = true) {
    if (!rows?.length) return [];
    if (!hasHeader) {
      return rows.map((row) => ({
        Type: row[0] || '',
        Country: row[1] || '',
        Location: row[2] || ''
      }));
    }
    const headers = rows[0].map((h) => String(h || '').trim());
    return rows.slice(1).map((row) => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = row[index] ?? '';
      });
      return obj;
    });
  }

  function normalizeRecord(row) {
    return {
      id: String(row.id || row.ID || row.Id || ''),
      type: String(row.type || row.Type || '').trim(),
      ConcertName: String(row.ConcertName || row.concertName || '').trim(),
      Artist: String(row.Artist || row.artist || '').trim(),
      Country: String(row.Country || row.country || '').trim(),
      Location: String(row.Location || row.location || '').trim(),
      Date: String(row.Date || row.date || '').trim(),
      Price: String(row.Price || row.price || '').trim(),
      Seat: String(row.Seat || row.seat || '').trim(),
      imgUrlS: String(row.imgUrlS || row.imgurls || '').trim(),
      imgUrlM: String(row.imgUrlM || row.imgurlm || '').trim(),
      note: String(row.note || row.Note || '').trim(),
      partner: String(row.partner || row.Partner || '').trim(),
      favorite: toBoolean(row.favorite || row.Favorite)
    };
  }

  function toBoolean(value) {
    const v = String(value ?? '').trim().toLowerCase();
    return ['true', '1', 'yes', 'y', '是', '❤️', '❤', '最愛'].includes(v);
  }

  function populateFilterOptions() {
    const types = uniqueValues(state.fieldRows.map((r) => r.type));
    const countries = uniqueValues(state.fieldRows.map((r) => r.country));
    fillOptions(els.filterType, types, true);
    fillOptions(els.filterCountry, countries, true);
    els.filterType.value = state.filters.type || '';
    els.filterCountry.value = state.filters.country || '';
  }

  function populateFormOptions() {
    const types = uniqueValues(state.fieldRows.map((r) => r.type));
    const countries = uniqueValues(state.fieldRows.map((r) => r.country));
    fillOptions(els.entryType, types, true);
    fillOptions(els.entryCountry, countries, true);

    if (state.draft) {
      els.entryType.value = state.draft.type || '';
      els.entryCountry.value = state.draft.country || '';
    }
    syncLocationsByCountry();
    if (state.draft) {
      els.entryLocation.value = state.draft.location || '';
    }
  }

  function uniqueValues(list) {
    return [...new Set(list.map((v) => String(v || '').trim()).filter(Boolean))];
  }

  function fillOptions(selectEl, values, keepFirst = true) {
    const first = keepFirst ? selectEl.querySelector('option')?.outerHTML || '<option value="">全部</option>' : '';
    selectEl.innerHTML = first + values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join('');
  }

  function syncLocationsByCountry() {
    const country = els.entryCountry.value.trim();
    const matches = state.fieldRows
      .filter((r) => !country || r.country === country)
      .map((r) => r.location)
      .filter(Boolean);
    const uniqueLocations = uniqueValues(matches);
    if (!uniqueLocations.length) {
      els.entryLocation.innerHTML = '<option value="">請先選 Country</option>';
      return;
    }
    els.entryLocation.innerHTML = '<option value="">請選擇</option>' + uniqueLocations.map((loc) => `<option value="${escapeHtml(loc)}">${escapeHtml(loc)}</option>`).join('');
    if (state.draft?.location) {
      els.entryLocation.value = state.draft.location;
    }
  }

  function onFilterChange() {
    state.filters.search = els.searchInput.value.trim();
    state.filters.type = els.filterType.value;
    state.filters.country = els.filterCountry.value;
    state.filters.date = els.filterDate.value;
    localStorage.setItem('concert_review_search', state.filters.search);
    localStorage.setItem('concert_review_filter_type', state.filters.type);
    localStorage.setItem('concert_review_filter_country', state.filters.country);
    localStorage.setItem('concert_review_filter_date', state.filters.date);
    state.page = 1;
    renderAll();
  }

  function onSortChange() {
    state.sortMode = els.sortSelect.value;
    localStorage.setItem('concert_review_sort', state.sortMode);
    state.page = 1;
    renderAll();
  }

  function onYearChange() {
    const val = parseInt(els.yearInput.value, 10);
    state.year = Number.isFinite(val) ? val : new Date().getFullYear();
    localStorage.setItem('concert_review_year', String(state.year));
    els.yearLabel.textContent = String(state.year);
    renderAll();
  }

  function changePage(delta) {
    const totalPages = getTotalPages();
    const next = state.page + delta;
    if (next < 1 || next > totalPages) return;
    state.page = next;
    renderCards();
  }

  function getFilteredRecords() {
    let list = [...state.records];
    const search = state.filters.search.toLowerCase();
    if (search) {
      list = list.filter((r) => {
        const hay = [r.ConcertName, r.Artist, r.Country, r.Location, r.note, r.partner, r.Seat, r.Price, r.type].join(' ').toLowerCase();
        return hay.includes(search);
      });
    }
    if (state.filters.type) list = list.filter((r) => r.type === state.filters.type);
    if (state.filters.country) list = list.filter((r) => r.Country === state.filters.country);
    if (state.filters.date) list = list.filter((r) => normalizeDateOnly(r.Date) === state.filters.date);
    return list.sort(sortRecords);
  }

  function sortRecords(a, b) {
    const dateA = parseDate(a.Date)?.getTime() || 0;
    const dateB = parseDate(b.Date)?.getTime() || 0;
    const artistCmp = collator.compare(a.Artist || '', b.Artist || '');
    const countryCmp = collator.compare(a.Country || '', b.Country || '');

    switch (state.sortMode) {
      case 'date-asc':
        return dateA - dateB;
      case 'artist-asc':
        return artistCmp;
      case 'artist-desc':
        return -artistCmp;
      case 'country-asc':
        return countryCmp;
      case 'country-desc':
        return -countryCmp;
      case 'date-desc':
      default:
        return dateB - dateA;
    }
  }

  async function renderAll() {
    renderStats();
    renderCards();
    await renderMap();
  }

  function renderStats() {
    const total = state.records.length;
    const visible = getFilteredRecords().length;
    const favoriteCount = state.records.filter((r) => r.favorite).length;
    const cities = uniqueValues(state.records.map((r) => r.Country));
    const yearCount = state.records.filter((r) => String(r.Date || '').startsWith(String(state.year))).length;

    els.totalCountPill.textContent = `${total} 場`;
    els.filteredCountPill.textContent = `${visible} 場`;
    els.summaryTotal.textContent = String(total);
    els.summaryFavorite.textContent = String(favoriteCount);
    els.summaryCities.textContent = String(cities.length);
    els.yearCount.textContent = String(yearCount);
    els.displayCount.textContent = `${visible} 場`;
    els.displayPageRate.textContent = `第 ${state.page} / ${getTotalPages()} 頁`;

    const artistStats = topBy(state.records, (r) => r.Artist);
    const countryStats = topBy(state.records, (r) => r.Country);

    els.topArtist.textContent = artistStats.value || '—';
    els.topArtistMeta.textContent = artistStats.value ? `共 ${artistStats.count} 場` : '尚無資料';
    els.topCountry.textContent = countryStats.value || '—';
    els.topCountryMeta.textContent = countryStats.value ? `共 ${countryStats.count} 場` : '尚無資料';
  }

  function topBy(list, getter) {
    const map = new Map();
    list.forEach((item) => {
      const value = String(getter(item) || '').trim();
      if (!value) return;
      map.set(value, (map.get(value) || 0) + 1);
    });
    let topValue = '';
    let topCount = 0;
    for (const [value, count] of map.entries()) {
      if (count > topCount) {
        topValue = value;
        topCount = count;
      }
    }
    return { value: topValue, count: topCount };
  }

  async function renderMap() {
    if (!state.map || !state.markerLayer) return;
    const records = getFilteredRecords();
    const groups = groupByCity(records);

    state.markerLayer.clearLayers();
    state.cityMarkerIndex.clear();
    els.cityLegend.innerHTML = '';

    const resolved = [];
    for (const group of groups) {
      const geo = await resolveGeo(group.cityLabel);
      resolved.push({ ...group, geo });
    }

    const validPoints = resolved.filter((item) => item.geo);
    if (!validPoints.length) {
      els.mapSummary.textContent = records.length ? '目前沒有可定位的城市資料' : '尚未載入地圖資料';
      els.cityLegend.innerHTML = '<div class="city-chip is-empty"><span class="muted">尚未有可標點城市</span></div>';
      return;
    }

    const latlngs = [];
    validPoints.forEach((group) => {
      const [lat, lng] = group.geo;
      const marker = L.marker([lat, lng], {
        icon: L.divIcon({
          className: 'city-marker',
          html: `<div class="city-marker-bubble${group.count > 9 ? ' small' : ''}">${group.count}</div>`,
          iconSize: [34, 34],
          iconAnchor: [17, 17],
        }),
        riseOnHover: true,
      }).addTo(state.markerLayer);

      marker.bindPopup(buildPopupHtml(group));
      state.cityMarkerIndex.set(group.cityLabel, { marker, group });
      latlngs.push([lat, lng]);
    });

    if (latlngs.length === 1) {
      state.map.setView(latlngs[0], 5);
    } else if (latlngs.length > 1) {
      state.map.fitBounds(latlngs, { padding: [40, 40] });
    }

    els.mapSummary.textContent = `目前顯示 ${validPoints.length} 個城市標點、共 ${records.length} 場紀錄`;
    renderCityLegend(resolved);
    setTimeout(() => state.map.invalidateSize(), 60);
  }

  function groupByCity(records) {
    const map = new Map();
    records.forEach((record) => {
      const cityLabel = pickCityLabel(record);
      if (!cityLabel) return;
      const key = normalizeKey(cityLabel);
      const current = map.get(key) || { cityLabel, records: [], count: 0 };
      current.records.push(record);
      current.count += 1;
      current.cityLabel = current.cityLabel || cityLabel;
      map.set(key, current);
    });
    return [...map.values()].sort((a, b) => b.count - a.count || collator.compare(a.cityLabel, b.cityLabel));
  }

  function pickCityLabel(record) {
    const raw = String(record.Country || '').trim() || String(record.Location || '').trim();
    if (!raw) return '';
    const normalized = resolveCityCanonical(raw) || raw;
    return normalized;
  }

  function resolveCityCanonical(value) {
    const text = normalizeKey(value);
    if (!text) return '';
    for (const [alias, canonical] of Object.entries(CITY_ALIASES)) {
      if (normalizeKey(alias) === text) return canonical;
    }
    for (const key of Object.keys(DEFAULT_CITY_GEO)) {
      if (normalizeKey(key) === text) return key;
    }
    for (const [alias, canonical] of Object.entries(CITY_ALIASES)) {
      if (text.includes(normalizeKey(alias))) return canonical;
    }
    for (const key of Object.keys(DEFAULT_CITY_GEO)) {
      if (text.includes(normalizeKey(key))) return key;
    }
    return value.trim();
  }

  function normalizeKey(input) {
    return String(input ?? '')
      .toLowerCase()
      .replace(/\s+/g, '')
      .replace(/[()（）,【】\[\]·•、，。.!?~\-_/\\]/g, '')
      .trim();
  }

  async function resolveGeo(cityLabel) {
    const canonical = resolveCityCanonical(cityLabel);
    const cacheKey = normalizeKey(canonical || cityLabel);
    if (state.cityGeoCache[cacheKey]) return state.cityGeoCache[cacheKey];
    if (DEFAULT_CITY_GEO[canonical]) {
      const geo = DEFAULT_CITY_GEO[canonical];
      state.cityGeoCache[cacheKey] = geo;
      saveCityGeoCache();
      return geo;
    }

    if (!CONFIG.ENABLE_GEOCODING) return null;

    try {
      const query = canonical || cityLabel;
      const url = `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&q=${encodeURIComponent(query)}`;
      const response = await fetch(url, {
        headers: {
          'Accept': 'application/json'
        }
      });
      if (!response.ok) return null;
      const data = await response.json();
      if (!Array.isArray(data) || !data.length) return null;
      const item = data[0];
      const geo = [Number(item.lat), Number(item.lon)];
      if (!Number.isFinite(geo[0]) || !Number.isFinite(geo[1])) return null;
      state.cityGeoCache[cacheKey] = geo;
      saveCityGeoCache();
      return geo;
    } catch (error) {
      console.warn('Geocoding failed:', error);
      return null;
    }
  }

  function renderCityLegend(groups) {
    const topGroups = groups.slice(0, 12);
    els.cityLegend.innerHTML = topGroups.map((group) => {
      const markerData = state.cityMarkerIndex.get(group.cityLabel) || [...state.cityMarkerIndex.values()].find((item) => normalizeKey(item.group.cityLabel) === normalizeKey(group.cityLabel));
      const hasMarker = Boolean(markerData);
      const extra = group.geo ? '' : '<span class="muted">未定位</span>';
      return `
        <button type="button" class="city-chip ${hasMarker ? '' : 'is-empty'}" data-city="${escapeAttr(group.cityLabel)}" ${hasMarker ? '' : 'disabled'}>
          <span>${escapeHtml(group.cityLabel)}</span>
          <span class="count">${group.count}</span>
          ${extra}
        </button>
      `;
    }).join('');

    els.cityLegend.querySelectorAll('.city-chip[data-city]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const city = btn.getAttribute('data-city') || '';
        const markerData = state.cityMarkerIndex.get(city) || [...state.cityMarkerIndex.values()].find((item) => normalizeKey(item.group.cityLabel) === normalizeKey(city));
        if (!markerData) return;
        state.map.setView(markerData.marker.getLatLng(), Math.max(state.map.getZoom(), 5), { animate: true });
        markerData.marker.openPopup();
      });
    });
  }

  function buildPopupHtml(group) {
    const records = group.records.slice(0, 8);
    const list = records.map((record) => `<li>${escapeHtml(record.ConcertName || '未命名演唱會')}｜${escapeHtml(record.Artist || '—')}</li>`).join('');
    return `
      <div>
        <div class="map-popup-title">${escapeHtml(group.cityLabel)}</div>
        <div class="map-popup-meta">共 ${group.count} 場演唱會紀錄</div>
        <ul class="map-popup-list">
          ${list}
        </ul>
      </div>
    `;
  }

  function renderCards() {
    const filtered = getFilteredRecords();
    const totalPages = getTotalPages();
    state.page = Math.min(state.page, totalPages);
    const start = (state.page - 1) * CONFIG.PAGE_SIZE;
    const pageItems = filtered.slice(start, start + CONFIG.PAGE_SIZE);

    els.pageInfo.textContent = `第 ${state.page} 頁 / 共 ${totalPages} 頁`;
    els.displayPageRate.textContent = `第 ${state.page} / ${totalPages} 頁`;

    els.btnPrevPage.disabled = state.page <= 1;
    els.btnNextPage.disabled = state.page >= totalPages;

    if (filtered.length === 0) {
      els.emptyState.classList.remove('d-none');
      els.cardGrid.innerHTML = '';
    } else {
      els.emptyState.classList.add('d-none');
      els.cardGrid.innerHTML = pageItems.map(renderCard).join('');
      requestAnimationFrame(() => {
        document.querySelectorAll('.concert-card').forEach((card) => card.classList.add('is-ready'));
      });
    }
  }

  function renderCard(record) {
    const image = getCardImage(record);
    const dateText = formatDate(record.Date);
    const daysText = daysSinceText(record.Date);
    const favoriteBadge = record.favorite ? '<span class="badge-soft badge-favorite">❤️ 最愛</span>' : '';
    const typeBadge = record.type ? `<span class="badge-soft">${escapeHtml(record.type)}</span>` : '';
    const partnerBadge = record.partner ? `<span class="badge-soft">${escapeHtml(record.partner)}</span>` : '';
    const imageAlt = record.ConcertName || '演唱會照片';
    const fallback = `data:image/svg+xml;utf8,${fallbackImageSvg()}`;

    return `
      <div class="col-12 col-md-6 col-xl-4 card-grid-item">
        <article class="concert-card">
          <div class="concert-img-wrap">
            <picture>
              <source media="(min-width: 992px)" srcset="${escapeAttr(record.imgUrlM || image)}" />
              <source media="(max-width: 991px)" srcset="${escapeAttr(record.imgUrlS || image)}" />
              <img src="${escapeAttr(record.imgUrlS || image)}" alt="${escapeAttr(imageAlt)}" loading="lazy" onerror="this.onerror=null;this.src='${escapeAttr(fallback)}';" />
            </picture>
            <div class="concert-badges">
              ${typeBadge}
              ${favoriteBadge}
              ${partnerBadge}
            </div>
          </div>
          <div class="concert-body">
            <h3 class="concert-title">${escapeHtml(record.ConcertName || '未命名演唱會')}</h3>
            <div class="concert-artist">${escapeHtml(record.Artist || '—')}</div>
            <div class="concert-meta">
              <div>📅 ${escapeHtml(dateText)}</div>
              <div>📍 ${escapeHtml(record.Country || '—')}｜${escapeHtml(record.Location || '—')}</div>
              <div>💺 座位：${escapeHtml(record.Seat || '—')}</div>
              <div>💰 票價：${escapeHtml(formatPrice(record.Price))}</div>
              <div>⏳ ${escapeHtml(daysText)}</div>
            </div>
          </div>
        </article>
      </div>
    `;
  }

  function getCardImage(record) {
    return record.imgUrlM || record.imgUrlS || `data:image/svg+xml;utf8,${fallbackImageSvg()}`;
  }

  function fallbackImageSvg() {
    const svg = `
      <svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 800 600'>
        <defs>
          <linearGradient id='g' x1='0' y1='0' x2='1' y2='1'>
            <stop offset='0%' stop-color='#ffe4ef'/>
            <stop offset='100%' stop-color='#fff'/>
          </linearGradient>
        </defs>
        <rect width='800' height='600' rx='30' fill='url(#g)'/>
        <circle cx='400' cy='220' r='120' fill='#ffd1e2'/>
        <circle cx='320' cy='190' r='18' fill='#fff'/>
        <circle cx='480' cy='190' r='18' fill='#fff'/>
        <path d='M330 270 Q400 330 470 270' stroke='#fff' stroke-width='18' fill='none' stroke-linecap='round'/>
        <text x='400' y='420' font-size='42' text-anchor='middle' fill='#e96ea6' font-family='Arial'>Concert Memory</text>
      </svg>`;
    return encodeURIComponent(svg.replace(/\n\s+/g, ' ').trim());
  }

  function formatDate(value) {
    const date = parseDate(value);
    if (!date) return value || '—';
    return new Intl.DateTimeFormat('zh-TW', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      weekday: 'short',
    }).format(date);
  }

  function normalizeDateOnly(value) {
    const d = parseDate(value);
    if (!d) return '';
    return d.toISOString().slice(0, 10);
  }

  function parseDate(value) {
    if (!value) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(value)) return new Date(`${value}T00:00:00`);
    const parsed = new Date(value);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function daysSinceText(value) {
    const date = parseDate(value);
    if (!date) return '日期未設定';
    const today = new Date();
    const now = new Date(today.getFullYear(), today.getMonth(), today.getDate()).getTime();
    const target = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
    const diff = Math.floor((now - target) / 86400000);
    if (diff >= 0) return `已過 ${diff} 天`;
    return `距今 ${Math.abs(diff)} 天`;
  }

  function formatPrice(value) {
    const num = Number(String(value).replace(/[^\d.-]/g, ''));
    if (!Number.isFinite(num) || value === '') return '—';
    return `NT$ ${new Intl.NumberFormat('zh-TW').format(num)}`;
  }

  function getTotalPages() {
    return Math.max(1, Math.ceil(getFilteredRecords().length / CONFIG.PAGE_SIZE));
  }

  function openForm() {
    els.formBackdrop.classList.remove('d-none');
    document.body.classList.add('modal-open');
    if (state.draft) fillFormFromDraft(state.draft);
  }

  function closeForm(keepDraft = true) {
    if (keepDraft) saveDraftFromForm();
    els.formBackdrop.classList.add('d-none');
    document.body.classList.remove('modal-open');
  }

  function clearDraftAndForm() {
    clearForm();
    localStorage.removeItem(CONFIG.DRAFT_KEY);
    state.draft = null;
    toast('已清空草稿');
  }

  function clearForm() {
    els.entryForm.reset();
    els.entryLocation.innerHTML = '<option value="">請先選 Country</option>';
  }

  function saveDraftFromForm() {
    const draft = getFormValues();
    localStorage.setItem(CONFIG.DRAFT_KEY, JSON.stringify(draft));
    state.draft = draft;
  }

  function loadDraft() {
    try {
      return JSON.parse(localStorage.getItem(CONFIG.DRAFT_KEY) || 'null');
    } catch {
      return null;
    }
  }

  function fillFormFromDraft(draft) {
    if (!draft) return;
    els.entryType.value = draft.type || '';
    els.entryConcertName.value = draft.concertName || '';
    els.entryArtist.value = draft.artist || '';
    els.entryCountry.value = draft.country || '';
    syncLocationsByCountry();
    els.entryLocation.value = draft.location || '';
    els.entryDate.value = draft.date || '';
    els.entryPrice.value = draft.price || '';
    els.entrySeat.value = draft.seat || '';
    els.entryImgS.value = draft.imgUrlS || '';
    els.entryImgM.value = draft.imgUrlM || '';
    els.entryNote.value = draft.note || '';
    els.entryPartner.value = draft.partner || '';
    els.entryFavorite.value = draft.favorite || 'FALSE';
  }

  function getFormValues() {
    return {
      type: els.entryType.value.trim(),
      concertName: els.entryConcertName.value.trim(),
      artist: els.entryArtist.value.trim(),
      country: els.entryCountry.value.trim(),
      location: els.entryLocation.value.trim(),
      date: els.entryDate.value,
      price: els.entryPrice.value.trim(),
      seat: els.entrySeat.value.trim(),
      imgUrlS: els.entryImgS.value.trim(),
      imgUrlM: els.entryImgM.value.trim(),
      note: els.entryNote.value.trim(),
      partner: els.entryPartner.value.trim(),
      favorite: els.entryFavorite.value
    };
  }

  async function submitEntry(event) {
    event.preventDefault();
    if (!state.token) {
      toast('請先登入 Google');
      return;
    }

    const data = getFormValues();
    const missing = [];
    ['type', 'concertName', 'artist', 'country', 'location', 'date'].forEach((key) => {
      if (!data[key]) missing.push(key);
    });
    if (missing.length) {
      toast('請先補齊必填欄位');
      return;
    }

    const payload = [
      String(Date.now()),
      data.type,
      data.concertName,
      data.artist,
      data.country,
      data.location,
      data.date,
      data.price ? Number(data.price) : '',
      data.seat,
      data.imgUrlS,
      data.imgUrlM,
      data.note,
      data.partner,
      data.favorite === 'TRUE' ? 'TRUE' : 'FALSE'
    ];

    try {
      setSubmitting(true);
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(CONFIG.SPREADSHEET_ID)}/values/${encodeURIComponent(CONFIG.SHEETS.RECORDS + '!A:N')}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
      await axios.post(url, { values: [payload] }, {
        headers: { Authorization: `Bearer ${state.token}` }
      });

      toast('已成功送出並寫入 Google Sheet');
      clearForm();
      localStorage.removeItem(CONFIG.DRAFT_KEY);
      state.draft = null;
      await loadRecords();
      await renderAll();
      closeForm(false);
    } catch (error) {
      console.error(error);
      const msg = error?.response?.data?.error?.message || error.message || '送出失敗';
      toast(`送出失敗：${msg}`);
    } finally {
      setSubmitting(false);
    }
  }

  function setSubmitting(isSubmitting) {
    const btn = els.btnSubmitEntry;
    if (btn) {
      btn.disabled = isSubmitting;
      btn.textContent = isSubmitting ? '送出中...' : '送出';
    }
    els.btnClearDraft.disabled = isSubmitting;
    els.btnCloseForm.disabled = isSubmitting;
  }

  function setLoading(isLoading) {
    state.loading = isLoading;
    els.loadingState.classList.toggle('d-none', !isLoading);
    if (isLoading) {
      els.emptyState.classList.add('d-none');
    }
  }

  function showStatus(title, hint) {
    els.syncStatus.textContent = title;
    els.syncHint.textContent = hint;
  }

  function loadCityGeoCache() {
    try {
      return JSON.parse(localStorage.getItem(CONFIG.CITY_CACHE_KEY) || '{}');
    } catch {
      return {};
    }
  }

  function saveCityGeoCache() {
    localStorage.setItem(CONFIG.CITY_CACHE_KEY, JSON.stringify(state.cityGeoCache));
  }

  function toast(message) {
    const id = `toast-${Date.now()}`;
    const item = document.createElement('div');
    item.className = 'toast align-items-center text-bg-light border-0 show mb-2';
    item.id = id;
    item.setAttribute('role', 'alert');
    item.setAttribute('aria-live', 'assertive');
    item.setAttribute('aria-atomic', 'true');
    item.innerHTML = `
      <div class="d-flex">
        <div class="toast-body">${escapeHtml(message)}</div>
        <button type="button" class="btn-close me-2 m-auto" aria-label="Close"></button>
      </div>
    `;
    item.querySelector('.btn-close').addEventListener('click', () => item.remove());
    els.toastHost.appendChild(item);
    setTimeout(() => item.remove(), 2800);
  }

  function escapeHtml(input) {
    return String(input ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/\"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function escapeAttr(input) {
    return escapeHtml(input).replace(/`/g, '&#96;');
  }

  function debounce(fn, wait) {
    let timer = null;
    return (...args) => {
      clearTimeout(timer);
      timer = setTimeout(() => fn(...args), wait);
    };
  }
})();
