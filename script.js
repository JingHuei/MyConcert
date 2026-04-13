/* Lucy 的演唱會記錄網站
   - 公開資料：opensheet JSON
   - 寫入資料：Google OAuth + Sheets API append
   - 狀態管理：單一 state
*/

const CONFIG = {
  GCP_ID: "562464022417-8c2sckejaft6de7ch8kejqomm1fbi0ga.apps.googleusercontent.com",
  SHEET_ID: "1tfNjim8BbKmv3KVNIQrRvx2T6W8YtfjyJCr1_wXDcrA",
};

const SHEET_NAME = "演唱會紀錄";
const FIELD_SHEET_NAME = "欄位表";
const STORAGE_KEYS = {
  draft: "lucy_concert_draft_v1",
  token: "lucy_google_token_v1",
  mapCity: "lucy_map_city_v1",
};

const state = {
  loading: true,
  error: "",
  allConcerts: [],
  filteredConcerts: [],
  sheetFields: {
    types: [],
    countries: [],
    locations: [],
  },
  options: {
    years: [],
    artists: [],
    cities: [],
    types: [],
  },
  filters: {
    query: "",
    type: "",
    city: "",
    date: "",
    artist: "",
    favorite: "",
    year: "",
  },
  sort: "date-desc",
  pagination: {
    page: 1,
    perPage: 10,
  },
  ui: {
    mapOverlayOpen: false,
    mapLayer: "list",
    mapCity: "",
    mapDetail: null,
    addOverlayOpen: false,
    authMessage: "",
    submitting: false,
  },
  auth: {
    accessToken: "",
    tokenClient: null,
  },
  map: {
    instance: null,
    markers: [],
    geocodeCache: loadJSON("lucy_geocode_cache_v1", {}),
  },
  draft: loadJSON(STORAGE_KEYS.draft, {}),
};

const els = {};
let cardObserver = null;

function loadJSON(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function saveJSON(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {}
}

function syncBodyScroll() {
  document.body.style.overflow = (state.ui.mapOverlayOpen || state.ui.addOverlayOpen) ? "hidden" : "";
}

function $(id) {
  return document.getElementById(id);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatDate(dateStr) {
  if (!dateStr) return "—";
  const date = parseDate(dateStr);
  if (!date) return dateStr;
  return new Intl.DateTimeFormat("zh-Hant", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    weekday: "short",
  }).format(date);
}

function parseDate(dateStr) {
  if (!dateStr) return null;
  const normalized = String(dateStr).trim().replaceAll("-", "/");
  const date = new Date(normalized);
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function isoDateValue(dateStr) {
  const date = parseDate(dateStr);
  if (!date) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function daysSince(dateStr) {
  const date = parseDate(dateStr);
  if (!date) return "—";
  const diff = Date.now() - date.getTime();
  const days = Math.max(0, Math.floor(diff / 86400000));
  return `${days} 天前`;
}

function splitArtists(text) {
  return String(text || "")
    .split(/[、,，\/|]/g)
    .map((v) => v.trim())
    .filter(Boolean);
}

function unique(values) {
  return [...new Set(values.filter(Boolean))];
}

function normalizeConcert(row) {
  const artistList = splitArtists(row.Artist);
  return {
    id: String(row.id || row.ID || row.timestamp || Date.now()),
    type: row.type || "",
    ConcertName: row.ConcertName || "",
    Artist: row.Artist || "",
    artistList,
    Country: row.Country || "",
    Location: row.Location || "",
    Date: row.Date || "",
    Price: row.Price || "",
    Seat: row.Seat || "",
    imgUrlS: row.imgUrlS || "",
    imgUrlM: row.imgUrlM || "",
    note: row.note || "",
    partner: row.partner || "",
    favorite: String(row.favorite).toLowerCase() === "true" || row.favorite === true,
    dateMs: parseDate(row.Date)?.getTime() || 0,
    year: parseDate(row.Date)?.getFullYear() || "",
  };
}

async function fetchSheet(sheetName) {
  const url = `https://opensheet.elk.sh/${encodeURIComponent(CONFIG.SHEET_ID)}/${encodeURIComponent(sheetName)}`;
  const res = await axios.get(url, { timeout: 15000 });
  return Array.isArray(res.data) ? res.data : [];
}

async function initData() {
  try {
    state.loading = true;
    renderLoading();

    const [concertRows, fieldRows] = await Promise.all([
      fetchSheet(SHEET_NAME),
      fetchSheet(FIELD_SHEET_NAME).catch(() => []),
    ]);

    state.allConcerts = concertRows.map(normalizeConcert).filter((item) => item.ConcertName);
    state.sheetFields = {
      types: unique(fieldRows.map((r) => r.type).filter(Boolean)),
      countries: unique(fieldRows.map((r) => r.Country).filter(Boolean)),
      locations: unique(fieldRows.map((r) => r.Location).filter(Boolean)),
    };

    buildOptionsFromData();
    bindSelectOptions();
    renderAll();
    initMap();
    await renderMap();
  } catch (error) {
    console.error(error);
    state.error = "資料載入失敗，請確認 SHEET_ID 與公開權限設定。";
    renderError();
  } finally {
    state.loading = false;
  }
}

function buildOptionsFromData() {
  const allYears = unique(
    state.allConcerts
      .map((item) => item.year)
      .filter(Boolean)
      .sort((a, b) => b - a)
  );
  const allArtists = unique(
    state.allConcerts
      .flatMap((item) => item.artistList.length ? item.artistList : splitArtists(item.Artist))
      .sort((a, b) => a.localeCompare(b, "zh-Hant"))
  );
  const allCities = unique(state.allConcerts.map((item) => item.Country).filter(Boolean).sort((a, b) => a.localeCompare(b, "zh-Hant")));

  state.options = {
    years: allYears,
    artists: allArtists,
    cities: allCities,
    types: unique([
      ...state.sheetFields.types,
      ...state.allConcerts.map((item) => item.type).filter(Boolean),
    ]),
  };
}

function bindSelectOptions() {
  const optionSets = {
    typeFilter: state.options.types,
    cityFilter: state.options.cities,
    artistFilter: state.options.artists,
    yearFilter: state.options.years,
    statsYearSelect: state.options.years,
    addType: state.options.types,
    addCountry: state.options.cities.length ? state.options.cities : state.sheetFields.countries,
    addLocation: state.sheetFields.locations.length ? state.sheetFields.locations : unique(state.allConcerts.map((item) => item.Location)),
  };

  Object.entries(optionSets).forEach(([id, items]) => {
    const select = $(id);
    if (!select) return;
    const keepFirst = select.querySelector("option");
    select.innerHTML = "";
    if (keepFirst) select.appendChild(keepFirst);
    if (id === "statsYearSelect") {
      select.innerHTML = '<option value="">請先選擇年份</option>';
    }
    items.forEach((item) => {
      if (item === undefined || item === null || String(item).trim() === "") return;
      const option = document.createElement("option");
      option.value = String(item);
      option.textContent = String(item);
      select.appendChild(option);
    });
  });

  if ($("addFavorite")) {
    $("addFavorite").value = String(state.draft.favorite ?? false);
  }
}

function renderLoading() {
  $("loadingState")?.classList.remove("hidden");
  $("emptyState")?.classList.add("hidden");
  $("concertList").innerHTML = "";
  $("paginationControls").innerHTML = "";
  $("paginationInfo").textContent = "";
}

function renderError() {
  $("loadingState")?.classList.add("hidden");
  $("emptyState")?.classList.remove("hidden");
  $("emptyState").innerHTML = `
    <p class="state-title">${escapeHtml(state.error)}</p>
    <p class="state-subtitle">請檢查 Google Sheets 是否為公開可讀，或稍後再試。</p>
  `;
}

function applyFiltersAndSort() {
  let list = [...state.allConcerts];

  const { query, type, city, date, artist, favorite, year } = state.filters;
  const q = query.trim().toLowerCase();

  if (q) {
    list = list.filter((item) => {
      const haystack = [
        item.ConcertName,
        item.Artist,
        item.Country,
        item.Location,
        item.Seat,
        item.note,
        item.partner,
      ].join(" ").toLowerCase();
      return haystack.includes(q);
    });
  }

  if (type) list = list.filter((item) => item.type === type);
  if (city) list = list.filter((item) => item.Country === city);
  if (date) list = list.filter((item) => isoDateValue(item.Date) === date);
  if (year) list = list.filter((item) => String(item.year) === String(year));
  if (favorite) list = list.filter((item) => String(item.favorite) === favorite);
  if (artist) {
    list = list.filter((item) => item.artistList.some((name) => name === artist || name.includes(artist) || artist.includes(name)));
  }

  const collator = new Intl.Collator("zh-Hant", { numeric: true, sensitivity: "base" });
  const sorters = {
    "date-desc": (a, b) => b.dateMs - a.dateMs,
    "date-asc": (a, b) => a.dateMs - b.dateMs,
    "artist-asc": (a, b) => a.Artist.length - b.Artist.length || collator.compare(a.Artist, b.Artist),
    "artist-desc": (a, b) => b.Artist.length - a.Artist.length || collator.compare(b.Artist, a.Artist),
    "city-asc": (a, b) => a.Country.length - b.Country.length || collator.compare(a.Country, b.Country),
    "city-desc": (a, b) => b.Country.length - a.Country.length || collator.compare(b.Country, a.Country),
  };
  list.sort(sorters[state.sort] || sorters["date-desc"]);

  state.filteredConcerts = list;
  state.pagination.page = 1;
}

function renderAll() {
  applyFiltersAndSort();
  renderFiltersSummary();
  renderList();
  renderPagination();
  void renderMap();
  renderChips();
  renderStats();
}

function renderFiltersSummary() {
  const yearFilter = $("yearFilter");
  if (yearFilter && !Array.from(yearFilter.options).some((opt) => opt.value === state.filters.year && state.filters.year)) {
    yearFilter.value = "";
  }
}

function renderList() {
  const listEl = $("concertList");
  const loadingEl = $("loadingState");
  const emptyEl = $("emptyState");
  loadingEl?.classList.add("hidden");

  const total = state.filteredConcerts.length;
  if (!total) {
    emptyEl?.classList.remove("hidden");
    listEl.innerHTML = "";
    $("paginationInfo").textContent = "";
    return;
  }
  emptyEl?.classList.add("hidden");

  const start = (state.pagination.page - 1) * state.pagination.perPage;
  const slice = state.filteredConcerts.slice(start, start + state.pagination.perPage);

  listEl.innerHTML = slice.map(renderConcertCard).join("");

  requestAnimationFrame(() => {
    if (cardObserver) cardObserver.disconnect();
    cardObserver = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) entry.target.classList.add("is-visible");
      });
    }, { threshold: 0.12 });
    document.querySelectorAll(".concert-card").forEach((el) => cardObserver.observe(el));
  });

  slice.forEach((item) => {
    const img = document.querySelector(`[data-img-id="${item.id}"]`);
    if (img) {
      img.addEventListener("error", () => swapToFallback(img, item));
    }
  });
}

function renderConcertCard(item) {
  const primaryImg = getResponsiveImage(item);
  const artistText = escapeHtml(item.Artist || "—");
  const cityText = escapeHtml(item.Country || "—");
  const locationText = escapeHtml(item.Location || "—");
  const seatText = escapeHtml(item.Seat || "—");
  const priceText = item.Price ? `NT$ ${escapeHtml(item.Price)}` : "—";
  const favorite = item.favorite ? '<span class="favorite-badge" title="最愛">❤️</span>' : '<span class="favorite-badge" title="一般">🤍</span>';

  return `
    <article class="concert-card panel" data-card-id="${escapeHtml(item.id)}">
      <div class="concert-media">
        ${primaryImg}
      </div>
      <div class="card-body">
        <div class="card-top">
          <h3 class="card-title">${escapeHtml(item.ConcertName)}</h3>
          ${favorite}
        </div>
        <p class="card-artist">${artistText}</p>
        <p class="card-meta">${formatDate(item.Date)}</p>
        <p class="card-meta">${cityText} · ${locationText}</p>
        <div class="card-submeta">
          <span>座位：${seatText}</span>
          <span>票價：${priceText}</span>
          <span>距離今天：${daysSince(item.Date)}</span>
        </div>
        <div class="card-actions">
          <span class="card-chip"><i class="fa-regular fa-calendar"></i> ${item.type || "—"}</span>
          <button class="link-btn" type="button" data-open-detail="${escapeHtml(item.id)}">查看詳情</button>
        </div>
      </div>
    </article>
  `;
}

function getResponsiveImage(item) {
  const mobileUrl = item.imgUrlS?.trim();
  const desktopUrl = item.imgUrlM?.trim();
  const src = mobileUrl || desktopUrl || "";
  if (!src) {
    return `<div class="fallback-art" aria-label="圖片缺失">🎤</div>`;
  }
  return `
    <picture>
      <source media="(min-width: 768px)" srcset="${escapeHtml(desktopUrl || src)}">
      <img data-img-id="${escapeHtml(item.id)}" src="${escapeHtml(src)}" alt="${escapeHtml(item.ConcertName)}" loading="lazy" />
    </picture>
  `;
}

function swapToFallback(imgEl, item) {
  const wrapper = imgEl.closest(".concert-media");
  if (!wrapper) return;
  wrapper.innerHTML = `<div class="fallback-art" aria-label="圖片載入失敗">🎤</div>`;
}

function renderPagination() {
  const info = $("paginationInfo");
  const controls = $("paginationControls");
  const total = state.filteredConcerts.length;
  const totalPages = Math.max(1, Math.ceil(total / state.pagination.perPage));
  const page = Math.min(state.pagination.page, totalPages);

  if (!total) {
    info.textContent = "";
    controls.innerHTML = "";
    return;
  }

  info.textContent = `顯示 ${Math.min((page - 1) * state.pagination.perPage + 1, total)} - ${Math.min(page * state.pagination.perPage, total)} / ${total} 筆`;
  controls.innerHTML = "";

  const makeBtn = (label, targetPage, active = false, disabled = false) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `page-btn${active ? " active" : ""}`;
    btn.textContent = label;
    btn.disabled = disabled;
    btn.addEventListener("click", () => {
      state.pagination.page = targetPage;
      renderList();
      renderPagination();
      window.scrollTo({ top: document.querySelector("#listSection").offsetTop - 24, behavior: "smooth" });
    });
    return btn;
  };

  controls.appendChild(makeBtn("‹", Math.max(1, page - 1), false, page === 1));

  const maxButtons = 5;
  const start = Math.max(1, Math.min(page - 2, totalPages - maxButtons + 1));
  const end = Math.min(totalPages, start + maxButtons - 1);

  for (let p = start; p <= end; p++) {
    controls.appendChild(makeBtn(String(p), p, p === page));
  }

  controls.appendChild(makeBtn("›", Math.min(totalPages, page + 1), false, page === totalPages));
}

function geocodeCity(city) {
  if (!city) return Promise.resolve(null);
  if (state.map.geocodeCache[city]) return Promise.resolve(state.map.geocodeCache[city]);

  const queries = [
    `${city}`,
    `${city}, Taiwan`,
    `${city}, South Korea`,
  ];

  return (async () => {
    for (const q of queries) {
      try {
        const url = `https://nominatim.openstreetmap.org/search?format=jsonv2&limit=1&q=${encodeURIComponent(q)}`;
        const res = await axios.get(url, { timeout: 12000 });
        const hit = Array.isArray(res.data) && res.data[0];
        if (hit?.lat && hit?.lon) {
          const result = { lat: Number(hit.lat), lon: Number(hit.lon) };
          state.map.geocodeCache[city] = result;
          saveJSON("lucy_geocode_cache_v1", state.map.geocodeCache);
          return result;
        }
      } catch (error) {
        console.warn("geocode error", city, error);
      }
    }
    return null;
  })();
}

async function renderMap() {
  if (!state.map.instance) return;
  const list = state.allConcerts;
  const groupMap = groupByCity(list);

  state.map.markers.forEach((marker) => marker.remove());
  state.map.markers = [];

  const entries = Object.entries(groupMap);
  const boundsPoints = [];

  for (const [city, concerts] of entries) {
    const coord = await geocodeCity(city);
    if (!coord) continue;
    const marker = L.marker([coord.lat, coord.lon], {
      title: `${city}（${concerts.length}）`,
    }).addTo(state.map.instance);

    marker.bindPopup(`
      <div style="min-width: 160px">
        <strong>${escapeHtml(city)}</strong><br />
        <span>${concerts.length} 場</span>
      </div>
    `);

    marker.on("click", () => openMapOverlay(city));
    state.map.markers.push(marker);
    boundsPoints.push([coord.lat, coord.lon]);
  }

  if (boundsPoints.length) {
    const bounds = L.latLngBounds(boundsPoints);
    state.map.instance.fitBounds(bounds.pad(0.28));
  } else {
    state.map.instance.setView([25.04, 121.56], 11);
  }

  const isEmpty = !entries.length;
  if (isEmpty) {
    $("cityChips").innerHTML = `<div class="state-card"><p class="state-title">目前沒有地圖資料</p></div>`;
  }
}

function groupByCity(list) {
  return list.reduce((acc, item) => {
    const city = item.Country || "未命名城市";
    (acc[city] ||= []).push(item);
    return acc;
  }, {});
}

function renderChips() {
  const el = $("cityChips");
  const entries = Object.entries(groupByCity(state.allConcerts));
  if (!entries.length) {
    el.innerHTML = "";
    return;
  }

  el.innerHTML = entries
    .sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0], "zh-Hant"))
    .map(([city, concerts]) => `
      <button type="button" class="city-chip" data-city-chip="${escapeHtml(city)}">
        ${escapeHtml(city)}（${concerts.length}）
      </button>
    `).join("");

  el.querySelectorAll("[data-city-chip]").forEach((btn) => {
    btn.addEventListener("click", () => openMapOverlay(btn.getAttribute("data-city-chip")));
  });
}

function renderStats() {
  const years = state.options.years;
  const yearSelect = $("statsYearSelect");
  if (yearSelect && !yearSelect.value && years.length) {
    yearSelect.value = years[0];
  }
  updateStatsYearCount(yearSelect?.value || "");
  const topArtists = calcTopLabel("artist");
  const topCities = calcTopLabel("city");
  $("topArtists").textContent = topArtists || "—";
  $("topCities").textContent = topCities || "—";
}

function calcTopLabel(type) {
  const source = type === "artist"
    ? state.allConcerts.flatMap((item) => item.artistList.length ? item.artistList : splitArtists(item.Artist))
    : state.allConcerts.map((item) => item.Country).filter(Boolean);

  const counts = source.reduce((acc, value) => {
    acc[value] = (acc[value] || 0) + 1;
    return acc;
  }, {});

  const values = Object.entries(counts);
  if (!values.length) return "";
  const max = Math.max(...values.map(([, count]) => count));
  return values.filter(([, count]) => count === max).map(([value]) => value).join("、");
}

function updateStatsYearCount(year) {
  const count = year ? state.allConcerts.filter((item) => String(item.year) === String(year)).length : 0;
  $("statsYearCount").textContent = year ? `${year} 年共 ${count} 場` : "—";
}

function initMap() {
  if (state.map.instance) return;

  state.map.instance = L.map("map", {
    zoomControl: true,
    scrollWheelZoom: true,
  }).setView([25.04, 121.56], 11);

  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    maxZoom: 19,
    attribution: '&copy; OpenStreetMap contributors',
  }).addTo(state.map.instance);

  state.map.instance.on("click", () => {
    if (state.ui.mapOverlayOpen) closeMapOverlay();
  });
}

function openMapOverlay(city) {
  state.ui.mapOverlayOpen = true;
  state.ui.mapCity = city || "";
  state.ui.mapLayer = "list";
  state.ui.mapDetail = null;
  updateOverlayBody();
  $("mapOverlay").classList.remove("hidden");
  $("mapOverlay").setAttribute("aria-hidden", "false");
  syncBodyScroll();
}

function closeMapOverlay() {
  state.ui.mapOverlayOpen = false;
  state.ui.mapCity = "";
  state.ui.mapLayer = "list";
  state.ui.mapDetail = null;
  $("mapOverlay").classList.add("hidden");
  $("mapOverlay").setAttribute("aria-hidden", "true");
  syncBodyScroll();
}

function updateOverlayBody() {
  const body = $("mapOverlayBody");
  if (state.ui.mapLayer === "detail" && state.ui.mapDetail) {
    body.innerHTML = renderDetailView(state.ui.mapDetail);
  } else {
    body.innerHTML = renderMapListView();
  }
}

function renderMapListView() {
  const grouped = groupByCity(state.allConcerts);
  const entries = Object.entries(grouped)
    .filter(([city]) => !state.ui.mapCity || city === state.ui.mapCity)
    .sort((a, b) => b[1].length - a[1].length || a[0].localeCompare(b[0], "zh-Hant"));

  const title = state.ui.mapCity
    ? `${state.ui.mapCity}（${grouped[state.ui.mapCity]?.length || 0}）`
    : "城市清單";

  const itemsHtml = entries.length
    ? entries.map(([city, concerts]) => `
      <div class="space-y-3">
        <h3 class="overlay-title group-title">${escapeHtml(city)}（${concerts.length}）</h3>
        <div class="overlay-list">
          ${concerts
            .sort((a, b) => b.dateMs - a.dateMs)
            .map((item) => `
              <button type="button" class="list-item text-left" data-map-item="${escapeHtml(item.id)}" aria-label="查看 ${escapeHtml(item.ConcertName)}">
                <p class="list-item-title">${escapeHtml(item.ConcertName)}</p>
                <p class="list-item-date">${formatDate(item.Date)}</p>
              </button>
            `).join("")}
        </div>
      </div>
    `).join("")
    : `<div class="state-card"><p class="state-title">這個城市目前沒有紀錄</p></div>`;

  return `
    <div class="space-y-5">
      <div class="space-y-2">
        <p class="overlay-eyebrow">Map Overview</p>
        <h2 id="mapOverlayTitle" class="overlay-title">${escapeHtml(title)}</h2>
      </div>
      ${itemsHtml}
    </div>
  `;
}

function renderDetailView(item) {
  return `
    <div class="overlay-detail">
      <div class="detail-media">
        ${getDetailImage(item)}
      </div>
      <div class="detail-copy">
        <p class="overlay-eyebrow">Concert Detail</p>
        <h2 class="detail-title">${escapeHtml(item.ConcertName)}</h2>
        <p class="detail-meta">${formatDate(item.Date)}</p>
        <p class="detail-meta">${escapeHtml(item.Country)} · ${escapeHtml(item.Location)}</p>
        <p class="detail-meta"><strong>歌手：</strong>${escapeHtml(item.Artist)}</p>
        <p class="detail-meta"><strong>類型：</strong>${escapeHtml(item.type || "—")}</p>
        <p class="detail-meta"><strong>座位：</strong>${escapeHtml(item.Seat || "—")}</p>
        <p class="detail-meta"><strong>票價：</strong>${item.Price ? `NT$ ${escapeHtml(item.Price)}` : "—"}</p>
      </div>
    </div>
  `;
}

function getDetailImage(item) {
  const mobileUrl = item.imgUrlS?.trim();
  const desktopUrl = item.imgUrlM?.trim();
  const src = desktopUrl || mobileUrl;
  if (!src) return `<div class="fallback-art">🎤</div>`;
  return `<img src="${escapeHtml(src)}" alt="${escapeHtml(item.ConcertName)}" loading="lazy" />`;
}

function showMapDetailById(id) {
  const item = state.allConcerts.find((concert) => concert.id === id);
  if (!item) return;
  state.ui.mapLayer = "detail";
  state.ui.mapDetail = item;
  updateOverlayBody();
}

function openAddOverlay() {
  state.ui.addOverlayOpen = true;
  $("addOverlay").classList.remove("hidden");
  $("addOverlay").setAttribute("aria-hidden", "false");
  syncBodyScroll();
  restoreDraftToForm();
  showAuthNotice(state.auth.accessToken ? "已登入，可直接新增資料。" : "請先使用 Google OAuth 登入，才能新增資料。");
}

function closeAddOverlay() {
  state.ui.addOverlayOpen = false;
  $("addOverlay").classList.add("hidden");
  $("addOverlay").setAttribute("aria-hidden", "true");
  syncBodyScroll();
}

function showAuthNotice(message) {
  const notice = $("authNotice");
  if (!notice) return;
  if (!message) {
    notice.classList.add("hidden");
    notice.textContent = "";
    return;
  }
  notice.textContent = message;
  notice.classList.remove("hidden");
}

function restoreDraftToForm() {
  const draft = loadJSON(STORAGE_KEYS.draft, {});
  const map = {
    addType: draft.type,
    addConcertName: draft.ConcertName,
    addArtist: draft.Artist,
    addCountry: draft.Country,
    addLocation: draft.Location,
    addDate: draft.Date,
    addPrice: draft.Price,
    addSeat: draft.Seat,
    addImgUrlS: draft.imgUrlS,
    addImgUrlM: draft.imgUrlM,
    addNote: draft.note,
    addPartner: draft.partner,
    addFavorite: String(draft.favorite ?? false),
  };

  Object.entries(map).forEach(([id, value]) => {
    const el = $(id);
    if (el && value !== undefined && value !== null) {
      el.value = value;
    }
  });
}

function collectFormData() {
  return {
    type: $("addType").value.trim(),
    ConcertName: $("addConcertName").value.trim(),
    Artist: $("addArtist").value.trim(),
    Country: $("addCountry").value.trim(),
    Location: $("addLocation").value.trim(),
    Date: $("addDate").value,
    Price: $("addPrice").value.trim(),
    Seat: $("addSeat").value.trim(),
    imgUrlS: $("addImgUrlS").value.trim(),
    imgUrlM: $("addImgUrlM").value.trim(),
    note: $("addNote").value.trim(),
    partner: $("addPartner").value.trim(),
    favorite: $("addFavorite").value === "true",
  };
}

function saveDraftFromForm() {
  const data = collectFormData();
  saveJSON(STORAGE_KEYS.draft, data);
}

function clearDraft() {
  localStorage.removeItem(STORAGE_KEYS.draft);
  $("addForm").reset();
  $("addFavorite").value = "false";
  showAuthNotice(state.auth.accessToken ? "已清空。" : "已清空，尚未登入。");
}

async function ensureToken() {
  if (state.auth.accessToken) return state.auth.accessToken;
  if (!window.google?.accounts?.oauth2) {
    throw new Error("Google Identity Services 尚未載入完成。");
  }

  return await new Promise((resolve, reject) => {
    if (!state.auth.tokenClient) {
      state.auth.tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: CONFIG.GCP_ID,
        scope: "https://www.googleapis.com/auth/spreadsheets",
        callback: (response) => {
          if (response?.access_token) {
            state.auth.accessToken = response.access_token;
            saveJSON(STORAGE_KEYS.token, { accessToken: response.access_token });
            resolve(response.access_token);
          } else {
            reject(new Error("OAuth 登入失敗。"));
          }
        },
        error_callback: (err) => reject(err),
      });
    }

    state.auth.tokenClient.requestAccessToken({
      prompt: state.auth.accessToken ? "" : "consent",
    });
  });
}

async function submitForm(event) {
  event.preventDefault();
  if (state.ui.submitting) return;

  try {
    state.ui.submitting = true;
    $("submitConcertBtn").disabled = true;
    showAuthNotice("資料送出中，請稍候...");

    await ensureToken();

    const payload = collectFormData();
    const required = ["type", "ConcertName", "Artist", "Country", "Location", "Date", "Price"];
    const missing = required.filter((key) => !payload[key]);
    if (missing.length) {
      throw new Error("請先完成必填欄位。");
    }

    const newRow = [
      String(Date.now()),
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
      String(payload.favorite),
    ];

    const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.SHEET_ID}/values/${encodeURIComponent(SHEET_NAME + "!A:N")}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`;
    await axios.post(url, { values: [newRow] }, {
      headers: {
        Authorization: `Bearer ${state.auth.accessToken}`,
        "Content-Type": "application/json",
      },
      timeout: 15000,
    });

    const newRecord = normalizeConcert({
      id: newRow[0],
      type: newRow[1],
      ConcertName: newRow[2],
      Artist: newRow[3],
      Country: newRow[4],
      Location: newRow[5],
      Date: newRow[6],
      Price: newRow[7],
      Seat: newRow[8],
      imgUrlS: newRow[9],
      imgUrlM: newRow[10],
      note: newRow[11],
      partner: newRow[12],
      favorite: newRow[13],
    });

    state.allConcerts.unshift(newRecord);
    buildOptionsFromData();
    bindSelectOptions();
    renderAll();

    $("addForm").reset();
    $("addFavorite").value = "false";
    localStorage.removeItem(STORAGE_KEYS.draft);

    showAuthNotice("新增成功，已同步到 Google Sheet。");
  } catch (error) {
    console.error(error);
    showAuthNotice(error.message || "新增失敗，請稍後再試。");
  } finally {
    state.ui.submitting = false;
    $("submitConcertBtn").disabled = false;
  }
}

function attachEvents() {
  $("searchInput").addEventListener("input", debounce((e) => {
    state.filters.query = e.target.value;
    renderAll();
  }, 180));

  $("typeFilter").addEventListener("change", (e) => {
    state.filters.type = e.target.value;
    renderAll();
  });
  $("cityFilter").addEventListener("change", (e) => {
    state.filters.city = e.target.value;
    renderAll();
  });
  $("dateFilter").addEventListener("change", (e) => {
    state.filters.date = e.target.value;
    renderAll();
  });
  $("artistFilter").addEventListener("change", (e) => {
    state.filters.artist = e.target.value;
    renderAll();
  });
  $("favoriteFilter").addEventListener("change", (e) => {
    state.filters.favorite = e.target.value;
    renderAll();
  });
  $("yearFilter").addEventListener("change", (e) => {
    state.filters.year = e.target.value;
    renderAll();
  });
  $("sortField").addEventListener("change", (e) => {
    state.sort = e.target.value;
    renderAll();
  });

  $("statsYearSelect").addEventListener("change", (e) => {
    updateStatsYearCount(e.target.value);
  });

  $("fabAdd").addEventListener("click", async () => {
    try {
      await ensureToken();
      showAuthNotice("已登入，可直接新增資料。");
      openAddOverlay();
    } catch (error) {
      console.error(error);
      showAuthNotice(error.message || "登入失敗，請稍後再試。");
      openAddOverlay();
    }
  });

  $("addCloseBtn").addEventListener("click", closeAddOverlay);
  $("mapCloseBtn").addEventListener("click", closeMapOverlay);
  $("mapBackBtn").addEventListener("click", () => {
    state.ui.mapLayer = "list";
    state.ui.mapDetail = null;
    updateOverlayBody();
  });

  document.querySelectorAll("[data-close]").forEach((el) => {
    el.addEventListener("click", (event) => {
      if (event.target?.dataset?.close === "map") closeMapOverlay();
      if (event.target?.dataset?.close === "add") closeAddOverlay();
    });
  });

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") {
      if (state.ui.addOverlayOpen) closeAddOverlay();
      else if (state.ui.mapOverlayOpen) closeMapOverlay();
    }
  });

  $("clearDraftBtn").addEventListener("click", clearDraft);
  $("addForm").addEventListener("submit", submitForm);

  const draftFields = [
    "addType", "addConcertName", "addArtist", "addCountry", "addLocation", "addDate",
    "addPrice", "addSeat", "addImgUrlS", "addImgUrlM", "addNote", "addPartner", "addFavorite",
  ];
  draftFields.forEach((id) => {
    $(id).addEventListener("input", debounce(saveDraftFromForm, 120));
    $(id).addEventListener("change", debounce(saveDraftFromForm, 120));
  });

  document.addEventListener("click", (event) => {
    const detailId = event.target.closest("[data-open-detail]")?.getAttribute("data-open-detail");
    if (detailId) {
      state.ui.mapCity = "";
      state.ui.mapLayer = "detail";
      state.ui.mapDetail = state.allConcerts.find((item) => item.id === detailId) || null;
      openMapOverlay();
      showMapDetailById(detailId);
    }
    const mapItemId = event.target.closest("[data-map-item]")?.getAttribute("data-map-item");
    if (mapItemId) {
      showMapDetailById(mapItemId);
    }
  });
}

function debounce(fn, wait = 150) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), wait);
  };
}

function bootstrapAuth() {
  const stored = loadJSON(STORAGE_KEYS.token, {});
  if (stored.accessToken) state.auth.accessToken = stored.accessToken;
}

function setInitialValues() {
  state.filters = {
    query: "",
    type: "",
    city: "",
    date: "",
    artist: "",
    favorite: "",
    year: "",
  };

  ["searchInput", "dateFilter"].forEach((id) => {
    $(id).value = "";
  });

  ["typeFilter", "cityFilter", "artistFilter", "favoriteFilter", "yearFilter", "sortField"].forEach((id) => {
    const el = $(id);
    if (!el) return;
  });
  $("sortField").value = "date-desc";
  $("favoriteFilter").value = "";
  $("statsYearSelect").value = "";
}

function updateMapOverlayHeight() {
  const map = $("map");
  if (state.map.instance) {
    setTimeout(() => state.map.instance.invalidateSize(), 50);
  }
  if (map && window.innerWidth < 768) {
    map.style.minHeight = "54vh";
  }
}

function bindAll() {
  Object.assign(els, {
    searchInput: $("searchInput"),
    typeFilter: $("typeFilter"),
    cityFilter: $("cityFilter"),
    dateFilter: $("dateFilter"),
    artistFilter: $("artistFilter"),
    favoriteFilter: $("favoriteFilter"),
    yearFilter: $("yearFilter"),
    sortField: $("sortField"),
  });
  attachEvents();
  window.addEventListener("resize", debounce(updateMapOverlayHeight, 180));
}

function syncAddDefaults() {
  if ($("addFavorite") && !$("addFavorite").value) {
    $("addFavorite").value = String(state.draft.favorite ?? false);
  }
}

async function main() {
  bootstrapAuth();
  setInitialValues();
  bindAll();
  syncAddDefaults();
  await initData();
  updateMapOverlayHeight();
}

document.addEventListener("DOMContentLoaded", main);
