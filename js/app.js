// ── Data Layer Config ─────────────────────────────────────────────────
const API_BASE = 'https://weygand-team.pages.dev';
const MONDAY_BOARD_URL = 'https://weygand.monday.com/boards/5765387029';

const CACHE_KEY_JOBS  = 'wjc_job_data_v4';
const CACHE_KEY_CLOCK = 'wjc_clock_data_v4';

// ── Costable Statuses ─────────────────────────────────────────────────
const COSTABLE_STATUSES = new Set([
    'Draft Provided', 'Ready for Signatures', 'Sign & Send',
    'Submitted with Municipality', 'Jeff Approved',
    'AR (Accounts Receivable)', 'Final Invoice Sent',
    'Collections', 'Uncollectable',
]);

function isCostable(status) { return COSTABLE_STATUSES.has(status); }

// ── Status Colors ─────────────────────────────────────────────────────
const STATUS_COLORS = {
    'Draft Provided': '#9d50dd',
    'Ready for Signatures': '#ff6d3b',
    'Sign & Send': '#175a63',
    'Submitted with Municipality': '#cab641',
    'Jeff Approved': '#757575',
    'AR (Accounts Receivable)': '#401694',
    'Final Invoice Sent': '#ffadad',
    'Collections': '#7f5347',
    'Uncollectable': '#225091',
    'Field': '#00c875',
    'Set Pins': '#579bfc',
    'Contract Drafting': '#66ccff',
    'Drafting': '#037f4c',
    'Field Review': '#c4c4c4',
    'Drafting Review': '#fdab3d',
    'Feed Back Provided': '#e2445c',
    'Waiting on Client': '#ff158a',
    'Approved': '#00c875',
    'Cancel': '#333333',
};

// ── Person color palette ──────────────────────────────────────────────
const PERSON_COLORS = [
    '#C99E50', '#7FA8C9', '#4A9B6F', '#EBCA9C',
    '#A07C3C', '#5B8FAD', '#D4AE6A', '#3D7A5C',
    '#B87333', '#6CA0BE', '#8FBD9A', '#C8A060',
    '#9F79C8', '#D4876A', '#5AABB8',
];

const personColorMap = {};
let nextPersonColorIdx = 0;

function colorForPerson(name) {
    if (personColorMap[name]) return personColorMap[name];
    personColorMap[name] = PERSON_COLORS[nextPersonColorIdx % PERSON_COLORS.length];
    nextPersonColorIdx++;
    return personColorMap[name];
}

// ── State ─────────────────────────────────────────────────────────────
let activeTab = 'jobs';
let jobDataCache   = null;  // Monday data: jobNum → { price, status, jobTypes, subitems, name }
let clockDataCache = null;  // Clockify: { projects, entries }
let mergedJobs = [];        // computed merged list
let allPersons = [];
let allJobTypes = [];
let allStatuses = [];

// Filters
let searchQuery = '';
let activeDateFrom = null;
let activeDateTo   = null;
let activePreset = 'mtd';
let activeJobTypes = new Set();
let activeStatuses = new Set();
let activePersons  = new Set();
let activeFilter   = 'all';
let activeSort     = 'hours-desc';

// ── Utility ───────────────────────────────────────────────────────────
function fmt$(n) {
    if (!n && n !== 0) return '—';
    return '$' + Math.round(n).toLocaleString('en-US');
}

function fmtHrs(h) {
    if (!h && h !== 0) return '—';
    return h.toFixed(1) + 'h';
}

function fmtRate(price, hours) {
    if (!price || !hours) return '—';
    return '$' + Math.round(price / hours) + '/hr';
}

function jobNum(projectName) {
    const m = projectName.match(/^(\d{4}-\d+)/);
    return m ? m[1] : null;
}

function durationHours(interval) {
    if (!interval || !interval.start || !interval.end) return 0;
    return (new Date(interval.end) - new Date(interval.start)) / 3600000;
}

function firstName(fullName) { return (fullName || '').split(' ')[0]; }

function statusColor(status) { return STATUS_COLORS[status] || '#888'; }

function el(tag, cls, text) {
    const e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
}

// ── DOM refs ──────────────────────────────────────────────────────────
const $loadingState = document.getElementById('loading-state');
const $errorState   = document.getElementById('error-state');
const $jobsList     = document.getElementById('jobs-list');
const $peopleList   = document.getElementById('people-list');
const $summaryCharts = document.getElementById('summary-charts');
const $statRevenue  = document.getElementById('stat-revenue');
const $statHours    = document.getElementById('stat-hours');
const $statRate     = document.getElementById('stat-rate');
const $statJobs     = document.getElementById('stat-jobs');
const $typeFilters  = document.getElementById('type-filters');
const $statusFilters = document.getElementById('status-filters');
const $personFilters = document.getElementById('person-filters');
const $cacheInfo    = document.getElementById('cache-info');
const $jobSearch    = document.getElementById('job-search');
const $dateFrom     = document.getElementById('date-from');
const $dateTo       = document.getElementById('date-to');
const $refreshBtn    = document.getElementById('refresh-btn');
const $sortSelect    = document.getElementById('sort-select');
const $filterSelect  = document.getElementById('filter-select');

// ── Date helpers ──────────────────────────────────────────────────────
function getDateRange(preset) {
    const now = new Date();
    const end = new Date();
    end.setHours(23, 59, 59, 999);
    let start;

    switch (preset) {
        case 'mtd':
            start = new Date(now.getFullYear(), now.getMonth(), 1);
            break;
        case 'last-month': {
            const lm = new Date(now.getFullYear(), now.getMonth() - 1, 1);
            start = lm;
            end.setTime(new Date(now.getFullYear(), now.getMonth(), 0, 23, 59, 59, 999).getTime());
            break;
        }
        case 'qtd': {
            const qMonth = Math.floor(now.getMonth() / 3) * 3;
            start = new Date(now.getFullYear(), qMonth, 1);
            break;
        }
        case 'last-quarter': {
            const curQ = Math.floor(now.getMonth() / 3);
            const prevQ = curQ === 0 ? 3 : curQ - 1;
            const yr = curQ === 0 ? now.getFullYear() - 1 : now.getFullYear();
            start = new Date(yr, prevQ * 3, 1);
            end.setTime(new Date(yr, prevQ * 3 + 3, 0, 23, 59, 59, 999).getTime());
            break;
        }
        case 'ytd':
            start = new Date(now.getFullYear(), 0, 1);
            break;
        default:
            start = new Date(now.getFullYear(), now.getMonth(), 1);
    }

    start.setHours(0, 0, 0, 0);
    return { start, end };
}

function toDateStr(d) {
    return d.toISOString().slice(0, 10);
}

function initDateRange() {
    const { start, end } = getDateRange(activePreset);
    activeDateFrom = start;
    activeDateTo   = end;
    $dateFrom.value = toDateStr(start);
    $dateTo.value   = toDateStr(end);
}

// ── API helpers ───────────────────────────────────────────────────────
async function apiFetch(path) {
    const res = await fetch(API_BASE + path);
    if (!res.ok) {
        if (res.status === 401 || res.status === 403) {
            throw new Error('Data layer offline — authentication required');
        }
        throw new Error(`API ${res.status}: ${path}`);
    }
    return res.json();
}

async function fetchAllJobs() {
    const jobs = [];
    let offset = 0;
    const limit = 500;
    while (true) {
        const data = await apiFetch(`/api/jobs?limit=${limit}&offset=${offset}`);
        jobs.push(...(data.data || []));
        if (jobs.length >= (data.meta?.total || 0)) break;
        offset += limit;
    }
    return jobs;
}

async function fetchAllTimeEntries(dateFrom, dateTo) {
    const entries = [];
    let offset = 0;
    const limit = 1000;
    const from = toDateStr(dateFrom);
    const to = toDateStr(dateTo);
    while (true) {
        const data = await apiFetch(`/api/time?date_from=${from}&date_to=${to}&limit=${limit}&offset=${offset}`);
        entries.push(...(data.data || []));
        if (entries.length >= (data.meta?.total || 0)) break;
        offset += limit;
    }
    return entries;
}

async function fetchApiStatus() {
    try {
        const data = await apiFetch('/api/status');
        return data.data || null;
    } catch(e) { return null; }
}

async function triggerSync() {
    const res = await fetch(API_BASE + '/api/sync', { method: 'POST' });
    if (!res.ok) throw new Error('Sync trigger failed: ' + res.status);
    return res.json();
}

// ── Merge API jobs + time entries ─────────────────────────────────────
function mergeData(jobsData, timeEntries) {
    // Build jobs lookup map
    const jobsMap = {};
    for (const job of jobsData) {
        if (!job.job_num) continue;
        jobsMap[job.job_num] = {
            mondayName: job.name || '',
            price: job.price || 0,
            status: job.status || '',
            jobTypes: job.job_types || [],
            subitems: [],
            costable: isCostable(job.status || ''),
            mondayUrl: job.monday_url || null,
        };
    }

    // Group time entries by job_num
    const byProject = {};
    for (const e of timeEntries) {
        let jn = e.job_num;
        if (!jn && e.project_name) {
            const m = e.project_name.match(/^(\d{4}-\d+)/);
            jn = m ? m[1] : null;
        }
        if (!jn) continue;

        const person  = e.person_name || 'Unknown';
        const hrs     = e.hours || 0;
        const dateStr = (e.date || '').slice(0, 7); // YYYY-MM

        if (!byProject[jn]) {
            byProject[jn] = {
                jobNum: jn,
                projectName: e.project_name || jn,
                people: {},
                totalHours: 0,
                byMonth: {},
            };
        }
        if (!byProject[jn].people[person]) {
            byProject[jn].people[person] = { hours: 0, days: new Set() };
        }
        byProject[jn].people[person].hours += hrs;
        if (e.date) byProject[jn].people[person].days.add(e.date);
        byProject[jn].totalHours += hrs;
        if (dateStr) {
            if (!byProject[jn].byMonth[dateStr]) byProject[jn].byMonth[dateStr] = 0;
            byProject[jn].byMonth[dateStr] += hrs;
        }
    }

    // Merge with job data
    const result = [];
    for (const [jn, clockData] of Object.entries(byProject)) {
        const mon      = jobsMap[jn] || null;
        const price    = mon?.price || 0;
        const status   = mon?.status || '';
        const jobTypes = mon?.jobTypes || [];
        const subitems = mon?.subitems || [];
        const costable = mon?.costable || false;
        const mondayName = mon?.mondayName || '';
        const mondayUrl  = mon?.mondayUrl || null;

        const name = mondayName && mondayName.trim() ? mondayName.trim() : clockData.projectName;

        const effectiveRate = price > 0 && clockData.totalHours > 0
            ? price / clockData.totalHours
            : null;

        const people = Object.entries(clockData.people).map(([pName, pd]) => ({
            name: pName,
            hours: pd.hours,
            days: pd.days.size,
            pct: clockData.totalHours > 0 ? pd.hours / clockData.totalHours : 0,
        })).sort((a, b) => b.hours - a.hours);

        result.push({
            jobNum: jn,
            name,
            status,
            jobTypes,
            subitems,
            costable,
            price,
            mondayUrl,
            totalHours: clockData.totalHours,
            effectiveRate,
            people,
            byMonth: clockData.byMonth,
        });
    }

    return result.sort((a, b) => b.totalHours - a.totalHours);
}

// ── LocalStorage cache ────────────────────────────────────────────────
function saveCache(key, data) {
    try {
        localStorage.setItem(key, JSON.stringify({ ts: Date.now(), data }));
    } catch(e) {
        console.warn('Cache save failed:', e);
    }
}

function loadCache(key) {
    try {
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        return JSON.parse(raw);
    } catch(e) { return null; }
}

function getCacheAge(key) {
    const c = loadCache(key);
    if (!c) return null;
    return Math.round((Date.now() - c.ts) / 60000); // minutes
}

function updateCacheInfo(apiStatus) {
    if (apiStatus?.last_syncs?.[0]?.completed_at) {
        const lastSync = new Date(apiStatus.last_syncs[0].completed_at);
        const ageMin = Math.round((Date.now() - lastSync) / 60000);
        $cacheInfo.textContent = `Last synced: ${ageMin < 60 ? ageMin + 'm' : Math.round(ageMin/60) + 'h'} ago`;
    } else {
        const mAge = getCacheAge(CACHE_KEY_JOBS);
        const cAge = getCacheAge(CACHE_KEY_CLOCK);
        if (mAge !== null && cAge !== null) {
            const age = Math.max(mAge, cAge);
            $cacheInfo.textContent = `Cached ${age < 60 ? age + 'm' : Math.round(age/60) + 'h'} ago`;
        } else {
            $cacheInfo.textContent = '';
        }
    }
}

// ── Main load ─────────────────────────────────────────────────────────
async function loadData(forceRefresh = false) {
    showLoading(true);
    clearError();

    try {
        // Get API status (for last sync time + staleness check)
        const apiStatus = await fetchApiStatus();

        // Jobs data
        if (forceRefresh || !jobDataCache) {
            const cached = loadCache(CACHE_KEY_JOBS);
            const apiLastSync = apiStatus?.last_syncs?.[0]?.completed_at;
            const cacheStale = !cached || (apiLastSync && new Date(apiLastSync) > new Date(cached.ts));

            if (!forceRefresh && cached && !cacheStale) {
                jobDataCache = cached.data;
            } else {
                jobDataCache = await fetchAllJobs();
                saveCache(CACHE_KEY_JOBS, jobDataCache);
            }
        }

        // Time entries (cached with date range key)
        const rangeKey = `${toDateStr(activeDateFrom)}_${toDateStr(activeDateTo)}`;
        const clockCached = loadCache(CACHE_KEY_CLOCK);
        const needTimeFetch = forceRefresh ||
            !clockDataCache ||
            !clockCached ||
            clockCached.data?.rangeKey !== rangeKey;

        if (needTimeFetch) {
            const entries = await fetchAllTimeEntries(activeDateFrom, activeDateTo);
            clockDataCache = { entries, rangeKey };
            saveCache(CACHE_KEY_CLOCK, clockDataCache);
        } else if (!clockDataCache) {
            clockDataCache = clockCached.data;
        }

        // Merge
        mergedJobs = mergeData(jobDataCache, clockDataCache.entries || []);

        // Build filter options
        buildFilterOptions();

        // Render
        applyFiltersAndRender();
        updateCacheInfo(apiStatus);

    } catch(err) {
        console.error(err);
        // Offline fallback — use cached data if available
        const cachedJobs  = loadCache(CACHE_KEY_JOBS);
        const cachedClock = loadCache(CACHE_KEY_CLOCK);
        if (cachedJobs && cachedClock) {
            jobDataCache   = cachedJobs.data;
            clockDataCache = cachedClock.data;
            mergedJobs     = mergeData(jobDataCache, clockDataCache.entries || []);
            buildFilterOptions();
            applyFiltersAndRender();
            updateCacheInfo(null);
            showError('Data layer offline — showing cached data. Click Refresh to retry.');
        } else {
            showError('Failed to load data: ' + err.message + ' — No cached data available.');
        }
    }

    showLoading(false);
}

// ── Build filter options ──────────────────────────────────────────────
function buildFilterOptions() {
    // Collect all job types
    const types = new Set();
    const statuses = new Set();
    const persons = new Set();

    for (const j of mergedJobs) {
        for (const t of j.jobTypes) types.add(t);
        if (j.status) statuses.add(j.status);
        for (const p of j.people) persons.add(p.name);
    }

    allJobTypes = [...types].sort();
    allStatuses = [...statuses].sort();
    allPersons  = [...persons].sort((a, b) => firstName(a).localeCompare(firstName(b)));

    renderFilterSection($typeFilters, allJobTypes, activeJobTypes);
    renderFilterSection($statusFilters, allStatuses, activeStatuses);
    renderPersonFilters();
}

function renderFilterSection($el, items, activeSet) {
    $el.innerHTML = '';
    if (items.length === 0) {
        $el.innerHTML = '<div style="font-size:11px;color:#888;padding:4px 0">No data</div>';
        return;
    }

    const toggleAll = el('span', 'filter-toggle-all', 'Toggle All');
    toggleAll.addEventListener('click', () => {
        const allOn = items.every(i => activeSet.has(i));
        if (allOn) { activeSet.clear(); }
        else { items.forEach(i => activeSet.add(i)); }
        renderFilterSection($el, items, activeSet);
        applyFiltersAndRender();
    });
    $el.appendChild(toggleAll);

    for (const item of items) {
        const label = el('label', 'filter-item');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = activeSet.size === 0 || activeSet.has(item);
        cb.addEventListener('change', () => {
            if (cb.checked) activeSet.add(item);
            else activeSet.delete(item);
            applyFiltersAndRender();
        });
        label.appendChild(cb);
        label.appendChild(document.createTextNode(item));
        $el.appendChild(label);
    }
}

function renderPersonFilters() {
    $personFilters.innerHTML = '';
    if (allPersons.length === 0) return;

    const toggleAll = el('span', 'filter-toggle-all', 'Toggle All');
    toggleAll.addEventListener('click', () => {
        const allOn = allPersons.every(p => activePersons.has(p));
        if (allOn) activePersons.clear();
        else allPersons.forEach(p => activePersons.add(p));
        renderPersonFilters();
        applyFiltersAndRender();
    });
    $personFilters.appendChild(toggleAll);

    for (const person of allPersons) {
        const label = el('label', 'filter-item');
        const dot = el('div', 'detail-dot');
        dot.style.background = colorForPerson(person);
        label.appendChild(dot);
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = activePersons.size === 0 || activePersons.has(person);
        cb.addEventListener('change', () => {
            if (cb.checked) activePersons.add(person);
            else activePersons.delete(person);
            applyFiltersAndRender();
        });
        label.appendChild(cb);
        label.appendChild(document.createTextNode(firstName(person)));
        $personFilters.appendChild(label);
    }
}

// ── Apply filters ─────────────────────────────────────────────────────
function applyFiltersAndRender() {
    let filtered = mergedJobs.filter(j => {
        // Search
        if (searchQuery) {
            const q = searchQuery.toLowerCase();
            if (!j.jobNum.toLowerCase().includes(q) && !j.name.toLowerCase().includes(q)) return false;
        }

        // Show filter: costable
        if (activeFilter === 'costable' && !j.costable) return false;

        // Job type filter (empty = all)
        if (activeJobTypes.size > 0) {
            const hasType = j.jobTypes.some(t => activeJobTypes.has(t));
            if (!hasType && j.jobTypes.length > 0) return false;
            if (!hasType && j.jobTypes.length === 0) return false; // no types = exclude if filter active
        }

        // Status filter
        if (activeStatuses.size > 0 && j.status) {
            if (!activeStatuses.has(j.status)) return false;
        }

        // Person filter
        if (activePersons.size > 0) {
            const hasPerson = j.people.some(p => activePersons.has(p.name));
            if (!hasPerson) return false;
        }

        return true;
    });

    // Show filter: outliers — apply after other filters so threshold is relative to visible set
    if (activeFilter === 'outliers') {
        const ceil = computeOutlierThreshold(filtered);
        filtered = filtered.filter(j => j.effectiveRate != null && j.effectiveRate > ceil);
    }

    updateStats(filtered);
    renderCurrentTab(filtered);
}

// ── Update summary stats ──────────────────────────────────────────────
function updateStats(jobs) {
    const costableJobs = jobs.filter(j => j.costable && j.price > 0);
    const totalRevenue = costableJobs.reduce((s, j) => s + j.price, 0);
    const totalHours   = jobs.reduce((s, j) => s + j.totalHours, 0);
    const costHours    = costableJobs.reduce((s, j) => s + j.totalHours, 0);
    const avgRate      = totalRevenue > 0 && costHours > 0 ? totalRevenue / costHours : 0;

    $statRevenue.textContent = totalRevenue > 0 ? fmt$(totalRevenue) : '—';
    $statHours.textContent   = fmtHrs(totalHours);
    $statRate.textContent    = avgRate > 0 ? '$' + Math.round(avgRate) + '/hr' : '—';
    $statJobs.textContent    = jobs.length;
}

// ── Tab rendering dispatch ────────────────────────────────────────────
function renderCurrentTab(filtered) {
    if (activeTab === 'jobs') renderJobs(filtered);
    else if (activeTab === 'people') renderPeople(filtered);
    else if (activeTab === 'summary') renderSummary(filtered);
}

// ── Sort helpers ─────────────────────────────────────────────────────
function sortJobs(jobs, sortKey) {
    const sorted = [...jobs];
    switch (sortKey) {
        case 'rate-asc':
            sorted.sort((a, b) => {
                const ra = a.effectiveRate ?? Infinity;
                const rb = b.effectiveRate ?? Infinity;
                return ra - rb;
            });
            break;
        case 'rate-desc':
            sorted.sort((a, b) => {
                const ra = a.effectiveRate ?? -Infinity;
                const rb = b.effectiveRate ?? -Infinity;
                return rb - ra;
            });
            break;
        case 'price-desc':
            sorted.sort((a, b) => b.price - a.price);
            break;
        case 'job-num':
            sorted.sort((a, b) => a.jobNum.localeCompare(b.jobNum));
            break;
        default: // hours-desc
            sorted.sort((a, b) => b.totalHours - a.totalHours);
    }
    return sorted;
}

function computeOutlierThreshold(jobs) {
    const rates = jobs
        .filter(j => j.effectiveRate != null && j.effectiveRate > 0)
        .map(j => j.effectiveRate);
    if (rates.length < 3) return Infinity;
    rates.sort((a, b) => a - b);
    const q1 = rates[Math.floor(rates.length * 0.25)];
    const q3 = rates[Math.floor(rates.length * 0.75)];
    const iqr = q3 - q1;
    return q3 + 1.5 * iqr;
}

// ── Render: Jobs Tab ──────────────────────────────────────────────────
function renderJobs(jobs) {
    $jobsList.innerHTML = '';
    if (jobs.length === 0) {
        $jobsList.appendChild(emptyState('No jobs match the current filters.'));
        return;
    }

    const sorted = sortJobs(jobs, activeSort);
    const outlierCeil = computeOutlierThreshold(jobs);

    for (const job of sorted) {
        const wrapper = el('div', 'job-card-wrapper');

        // Card
        const card = el('div', 'job-card');

        // Header row
        const header = el('div', 'job-card-header');
        header.appendChild(el('span', 'job-num', job.jobNum));
        const nameEl = el('span', 'job-name', job.name !== job.jobNum ? job.name.replace(/^\d{4}-\d+\s*[-–]?\s*/,'') || job.name : '');
        header.appendChild(nameEl);
        if (job.status) {
            const badge = el('span', 'status-badge', job.status);
            badge.style.background = statusColor(job.status) + '22';
            badge.style.color = statusColor(job.status);
            badge.style.border = `1px solid ${statusColor(job.status)}44`;
            header.appendChild(badge);
        }
        if (job.effectiveRate != null && job.effectiveRate > outlierCeil) {
            const outlier = el('span', 'outlier-badge', 'High $/hr');
            header.appendChild(outlier);
        }
        card.appendChild(header);

        // Metrics row
        const metrics = el('div', 'job-card-metrics');

        const mPrice = el('div', 'metric');
        const priceVal = el('div', 'metric-value', job.price > 0 ? fmt$(job.price) : '—');
        if (!job.price) priceVal.classList.add('no-price');
        mPrice.appendChild(priceVal);
        mPrice.appendChild(el('div', 'metric-label', 'Total Price'));
        metrics.appendChild(mPrice);

        const mHours = el('div', 'metric');
        mHours.appendChild(el('div', 'metric-value', fmtHrs(job.totalHours)));
        mHours.appendChild(el('div', 'metric-label', 'Hours'));
        metrics.appendChild(mHours);

        const mRate = el('div', 'metric');
        const rateVal = el('div', 'metric-value', fmtRate(job.price, job.totalHours));
        if (job.effectiveRate) {
            rateVal.classList.add(job.effectiveRate >= 100 ? 'good' : 'warn');
        } else {
            rateVal.classList.add('no-price');
        }
        mRate.appendChild(rateVal);
        mRate.appendChild(el('div', 'metric-label', '$/hr'));
        metrics.appendChild(mRate);

        if (job.jobTypes.length > 0) {
            const mType = el('div', 'metric');
            mType.appendChild(el('div', 'metric-value', job.jobTypes.join(', ')));
            mType.appendChild(el('div', 'metric-label', 'Type'));
            metrics.appendChild(mType);
        }

        card.appendChild(metrics);

        // Person bar
        if (job.people.length > 0) {
            const barWrap = el('div', 'person-bar-wrap');
            for (const p of job.people) {
                const seg = el('div', 'person-bar-seg');
                seg.style.width = (p.pct * 100).toFixed(1) + '%';
                seg.style.background = colorForPerson(p.name);
                seg.title = `${firstName(p.name)}: ${fmtHrs(p.hours)} (${(p.pct*100).toFixed(0)}%)`;
                barWrap.appendChild(seg);
            }
            card.appendChild(barWrap);
        }

        card.addEventListener('click', () => wrapper.classList.toggle('expanded'));
        wrapper.appendChild(card);

        // Detail panel
        const detail = el('div', 'job-detail');

        // People breakdown
        if (job.people.length > 0) {
            const sec = el('div', 'detail-section');
            sec.appendChild(el('div', 'detail-section-title', 'By Person'));
            for (const p of job.people) {
                const row = el('div', 'detail-row');
                const dot = el('div', 'detail-dot');
                dot.style.background = colorForPerson(p.name);
                row.appendChild(dot);
                row.appendChild(el('span', 'detail-name', firstName(p.name)));
                row.appendChild(el('span', 'detail-val', p.days + 'd'));
                row.appendChild(el('span', 'detail-hrs', fmtHrs(p.hours)));
                row.appendChild(el('span', 'detail-pct', (p.pct*100).toFixed(0) + '%'));
                // Revenue share
                if (job.price > 0) {
                    row.appendChild(el('span', 'detail-val-em', fmt$(job.price * p.pct)));
                }
                sec.appendChild(row);
            }
            detail.appendChild(sec);
        }

        // Line items from subitems
        if (job.subitems.length > 0) {
            const sec = el('div', 'detail-section');
            sec.appendChild(el('div', 'detail-section-title', 'Line Items'));
            for (const si of job.subitems) {
                const row = el('div', 'detail-row');
                row.appendChild(el('span', 'detail-name', si.name));
                row.appendChild(el('span', 'detail-val-em', si.price > 0 ? fmt$(si.price) : '—'));
                sec.appendChild(row);
            }
            detail.appendChild(sec);
        }

        // Monday link
        const linkSec = el('div', 'detail-section');
        const link = document.createElement('a');
        link.href = job.mondayUrl || (MONDAY_BOARD_URL + '?term=' + encodeURIComponent(job.jobNum));
        link.target = '_blank';
        link.rel = 'noopener';
        link.style.cssText = 'color:#7FA8C9;font-size:12px;text-decoration:none;';
        link.textContent = 'View on Monday.com →';
        linkSec.appendChild(link);
        detail.appendChild(linkSec);

        wrapper.appendChild(detail);
        $jobsList.appendChild(wrapper);
    }
}

// ── Render: People Tab ────────────────────────────────────────────────
function renderPeople(jobs) {
    $peopleList.innerHTML = '';

    // Aggregate by person
    const personMap = {};
    for (const job of jobs) {
        for (const p of job.people) {
            if (!personMap[p.name]) {
                personMap[p.name] = {
                    name: p.name,
                    totalHours: 0,
                    costableHours: 0,
                    jobs: [],
                    revenue: 0,
                };
            }
            const pm = personMap[p.name];
            pm.totalHours += p.hours;
            if (job.costable) {
                pm.costableHours += p.hours;
                if (job.price > 0) pm.revenue += job.price * p.pct;
            }
            pm.jobs.push({
                jobNum: job.jobNum,
                name: job.name,
                hours: p.hours,
                pct: p.pct,
                status: job.status,
                price: job.price,
                effectiveRate: job.effectiveRate,
            });
        }
    }

    const persons = Object.values(personMap)
        .sort((a, b) => b.totalHours - a.totalHours);

    if (persons.length === 0) {
        $peopleList.appendChild(emptyState('No people data for current filters.'));
        return;
    }

    for (const person of persons) {
        const wrapper = el('div', 'person-card-wrapper');

        const card = el('div', 'person-card');

        // Header
        const header = el('div', 'person-card-header');

        const avatar = el('div', 'person-avatar', firstName(person.name)[0]);
        avatar.style.background = colorForPerson(person.name);
        header.appendChild(avatar);

        const info = el('div', 'person-card-info');
        info.appendChild(el('div', 'person-card-name', firstName(person.name)));
        const utilPct = person.totalHours > 0
            ? Math.round((person.costableHours / person.totalHours) * 100)
            : 0;
        info.appendChild(el('div', 'person-card-sub',
            `${person.jobs.length} jobs · ${utilPct}% costable`));
        header.appendChild(info);

        card.appendChild(header);

        // Metrics
        const metrics = el('div', 'person-card-metrics');

        const m1 = el('div', 'metric');
        m1.appendChild(el('div', 'metric-value', fmtHrs(person.totalHours)));
        m1.appendChild(el('div', 'metric-label', 'Total Hours'));
        metrics.appendChild(m1);

        const m2 = el('div', 'metric');
        m2.appendChild(el('div', 'metric-value', person.jobs.length.toString()));
        m2.appendChild(el('div', 'metric-label', 'Jobs'));
        metrics.appendChild(m2);

        if (person.revenue > 0) {
            const m3 = el('div', 'metric');
            m3.appendChild(el('div', 'metric-value', fmt$(person.revenue)));
            m3.appendChild(el('div', 'metric-label', 'Revenue Share'));
            metrics.appendChild(m3);

            const m4 = el('div', 'metric');
            const avgRate = person.costableHours > 0 ? person.revenue / person.costableHours : 0;
            m4.appendChild(el('div', 'metric-value', avgRate > 0 ? '$' + Math.round(avgRate) + '/hr' : '—'));
            m4.appendChild(el('div', 'metric-label', 'Avg Rate'));
            metrics.appendChild(m4);
        }

        card.appendChild(metrics);
        card.addEventListener('click', () => wrapper.classList.toggle('expanded'));
        wrapper.appendChild(card);

        // Detail
        const detail = el('div', 'person-detail');
        const sec = el('div', 'detail-section');
        sec.appendChild(el('div', 'detail-section-title', 'Jobs'));

        const sortedJobs = [...person.jobs].sort((a, b) => b.hours - a.hours);
        for (const j of sortedJobs) {
            const row = el('div', 'detail-row');
            row.appendChild(el('span', 'detail-name', `${j.jobNum}`));
            row.appendChild(el('span', 'detail-hrs', fmtHrs(j.hours)));
            if (j.effectiveRate) {
                row.appendChild(el('span', 'detail-val', '$' + Math.round(j.effectiveRate) + '/hr'));
            }
            if (j.status) {
                const badge = el('span', 'status-badge', j.status);
                badge.style.cssText = `background:${statusColor(j.status)}22;color:${statusColor(j.status)};border:1px solid ${statusColor(j.status)}44;font-size:9px;`;
                row.appendChild(badge);
            }
            sec.appendChild(row);
        }
        detail.appendChild(sec);
        wrapper.appendChild(detail);
        $peopleList.appendChild(wrapper);
    }
}

// ── Render: Summary Tab ───────────────────────────────────────────────
function renderSummary(jobs) {
    $summaryCharts.innerHTML = '';

    // 1. Revenue vs Hours by Month
    renderMonthlyChart(jobs);

    // 2. Avg $/hr by job type
    renderJobTypeRate(jobs);

    // 3. Hours by person
    renderHoursByPerson(jobs);

    // 4. Top 10 by revenue
    renderTopTable(jobs, 'revenue', 'Top 10 Jobs by Revenue', j => j.price);

    // 5. Top 10 by hours
    renderTopTable(jobs, 'hours', 'Top 10 Jobs by Hours', j => j.totalHours);

    // 6. Costable vs non-costable
    renderCostableBreakdown(jobs);
}

function renderMonthlyChart(jobs) {
    const sec = el('div', 'chart-section');
    sec.appendChild(el('div', 'chart-title', 'Revenue & Hours by Month'));

    // Aggregate by month
    const months = {};
    for (const job of jobs) {
        for (const [m, hrs] of Object.entries(job.byMonth || {})) {
            if (!months[m]) months[m] = { revenue: 0, hours: 0 };
            months[m].hours += hrs;
            if (job.costable) months[m].revenue += job.price
                ? (hrs / job.totalHours) * job.price
                : 0;
        }
    }

    const sortedMonths = Object.keys(months).sort();
    if (sortedMonths.length === 0) {
        sec.appendChild(el('p', '', 'No monthly data available.'));
        $summaryCharts.appendChild(sec);
        return;
    }

    const maxRevenue = Math.max(...sortedMonths.map(m => months[m].revenue));
    const maxHours   = Math.max(...sortedMonths.map(m => months[m].hours));

    const chartWrap = el('div', 'month-chart');

    for (const m of sortedMonths) {
        const d = months[m];
        const col = el('div', 'month-col');
        const bars = el('div', 'month-col-bars');

        // Revenue bar
        const revH = maxRevenue > 0 ? Math.round((d.revenue / maxRevenue) * 90) : 0;
        const revBar = el('div', 'month-bar');
        revBar.style.cssText = `height:${revH}px;background:#C99E50;`;
        revBar.title = `${m}: ${fmt$(d.revenue)} revenue`;
        bars.appendChild(revBar);

        // Hours bar (scale separately)
        const hrsH = maxHours > 0 ? Math.round((d.hours / maxHours) * 90) : 0;
        const hrsBar = el('div', 'month-bar');
        hrsBar.style.cssText = `height:${hrsH}px;background:#7FA8C9;`;
        hrsBar.title = `${m}: ${fmtHrs(d.hours)} hours`;
        bars.appendChild(hrsBar);

        col.appendChild(bars);
        col.appendChild(el('div', 'month-label', m.slice(5))); // MM
        chartWrap.appendChild(col);
    }

    sec.appendChild(chartWrap);

    // Legend
    const legend = el('div', 'chart-legend');
    legend.innerHTML = `
      <span class="legend-item"><span class="legend-dot" style="background:#C99E50"></span>Revenue (costable)</span>
      <span class="legend-item"><span class="legend-dot" style="background:#7FA8C9"></span>Hours</span>
    `;
    sec.appendChild(legend);
    $summaryCharts.appendChild(sec);
}

function renderJobTypeRate(jobs) {
    const sec = el('div', 'chart-section');
    sec.appendChild(el('div', 'chart-title', 'Avg $/hr by Job Type'));

    const typeData = {};
    for (const job of jobs) {
        if (!job.price || !job.totalHours) continue;
        for (const t of job.jobTypes) {
            if (!typeData[t]) typeData[t] = { totalRev: 0, totalHrs: 0 };
            typeData[t].totalRev += job.price;
            typeData[t].totalHrs += job.totalHours;
        }
    }

    const entries = Object.entries(typeData)
        .map(([t, d]) => ({ type: t, rate: d.totalHrs > 0 ? d.totalRev / d.totalHrs : 0 }))
        .filter(e => e.rate > 0)
        .sort((a, b) => b.rate - a.rate);

    if (entries.length === 0) {
        sec.appendChild(el('p', '', 'No job type rate data available.'));
        $summaryCharts.appendChild(sec);
        return;
    }

    const maxRate = entries[0].rate;
    for (const e of entries) {
        const row = el('div', 'h-bar-row');
        const label = el('div', 'h-bar-label', e.type);
        const track = el('div', 'h-bar-track');
        const fill  = el('div', 'h-bar-fill');
        fill.style.cssText = `width:${(e.rate/maxRate*100).toFixed(1)}%;background:#C99E50;`;
        track.appendChild(fill);
        row.appendChild(label);
        row.appendChild(track);
        row.appendChild(el('div', 'h-bar-val', '$' + Math.round(e.rate) + '/hr'));
        sec.appendChild(row);
    }

    $summaryCharts.appendChild(sec);
}

function renderHoursByPerson(jobs) {
    const sec = el('div', 'chart-section');
    sec.appendChild(el('div', 'chart-title', 'Hours by Person'));

    const personHours = {};
    for (const job of jobs) {
        for (const p of job.people) {
            if (!personHours[p.name]) personHours[p.name] = 0;
            personHours[p.name] += p.hours;
        }
    }

    const entries = Object.entries(personHours)
        .sort((a, b) => b[1] - a[1]);

    if (entries.length === 0) {
        sec.appendChild(el('p', '', 'No person data available.'));
        $summaryCharts.appendChild(sec);
        return;
    }

    const maxHrs = entries[0][1];
    for (const [name, hrs] of entries) {
        const row = el('div', 'h-bar-row');
        const label = el('div', 'h-bar-label', firstName(name));
        const track = el('div', 'h-bar-track');
        const fill  = el('div', 'h-bar-fill');
        fill.style.cssText = `width:${(hrs/maxHrs*100).toFixed(1)}%;background:${colorForPerson(name)};`;
        track.appendChild(fill);
        row.appendChild(label);
        row.appendChild(track);
        row.appendChild(el('div', 'h-bar-val', fmtHrs(hrs)));
        sec.appendChild(row);
    }

    $summaryCharts.appendChild(sec);
}

function renderTopTable(jobs, _key, title, getValue) {
    const sec = el('div', 'chart-section');
    sec.appendChild(el('div', 'chart-title', title));

    const sorted = [...jobs]
        .filter(j => getValue(j) > 0)
        .sort((a, b) => getValue(b) - getValue(a))
        .slice(0, 10);

    if (sorted.length === 0) {
        sec.appendChild(el('p', '', 'No data available.'));
        $summaryCharts.appendChild(sec);
        return;
    }

    const table = el('table', 'top-table');
    const thead = el('thead');
    const hrow  = el('tr');
    hrow.appendChild(el('th', '', '#'));
    hrow.appendChild(el('th', '', 'Job'));
    hrow.appendChild(el('th', '', _key === 'revenue' ? 'Revenue' : 'Hours'));
    hrow.appendChild(el('th', '', 'Rate'));
    thead.appendChild(hrow);
    table.appendChild(thead);

    const tbody = el('tbody');
    sorted.forEach((j, i) => {
        const row = el('tr');
        const td0 = el('td', 'num', (i+1).toString());
        const td1 = el('td', 'emph', j.jobNum);
        const td2 = el('td', 'num', _key === 'revenue' ? fmt$(j.price) : fmtHrs(j.totalHours));
        const td3 = el('td', 'num', fmtRate(j.price, j.totalHours));
        row.appendChild(td0);
        row.appendChild(td1);
        row.appendChild(td2);
        row.appendChild(td3);
        tbody.appendChild(row);
    });
    table.appendChild(tbody);
    sec.appendChild(table);
    $summaryCharts.appendChild(sec);
}

function renderCostableBreakdown(jobs) {
    const sec = el('div', 'chart-section');
    sec.appendChild(el('div', 'chart-title', 'Costable vs Non-Costable'));

    const costable    = jobs.filter(j => j.costable);
    const nonCostable = jobs.filter(j => !j.costable);

    const cRevenue = costable.reduce((s, j) => s + j.price, 0);
    const cHours   = costable.reduce((s, j) => s + j.totalHours, 0);
    const ncHours  = nonCostable.reduce((s, j) => s + j.totalHours, 0);

    const row = el('div', 'costable-row');

    const c1 = el('div', 'costable-card');
    c1.innerHTML = `<div class="costable-card-val">${costable.length}</div><div class="costable-card-sub">Costable Jobs</div>`;
    row.appendChild(c1);

    const c2 = el('div', 'costable-card');
    c2.innerHTML = `<div class="costable-card-val">${fmt$(cRevenue)}</div><div class="costable-card-sub">Costable Revenue</div>`;
    row.appendChild(c2);

    const c3 = el('div', 'costable-card');
    c3.innerHTML = `<div class="costable-card-val">${fmtHrs(cHours)}</div><div class="costable-card-sub">Costable Hours</div>`;
    row.appendChild(c3);

    const c4 = el('div', 'costable-card');
    c4.innerHTML = `<div class="costable-card-val">${fmtHrs(ncHours)}</div><div class="costable-card-sub">Non-Costable Hours</div>`;
    row.appendChild(c4);

    const c5 = el('div', 'costable-card');
    const totalH = cHours + ncHours;
    const utilPct = totalH > 0 ? Math.round((cHours / totalH) * 100) : 0;
    c5.innerHTML = `<div class="costable-card-val">${utilPct}%</div><div class="costable-card-sub">Utilization Rate</div>`;
    row.appendChild(c5);

    sec.appendChild(row);
    $summaryCharts.appendChild(sec);
}

// ── UI helpers ────────────────────────────────────────────────────────
function showLoading(on) {
    $loadingState.style.display = on ? 'flex' : 'none';
    ['tab-jobs', 'tab-people', 'tab-summary'].forEach(id => {
        document.getElementById(id).style.display = on ? 'none' : '';
    });
    if (!on) {
        document.getElementById('tab-' + activeTab).style.display = 'block';
    }
}

function showError(msg) {
    $errorState.textContent = msg;
    $errorState.hidden = !msg;
}

function clearError() { showError(''); }

function emptyState(msg) {
    const div = el('div', 'empty-state');
    div.innerHTML = `<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1" stroke-linecap="round"><path d="M9 19c-5 1.5-5-2.5-7-3m14 6v-3.87a3.37 3.37 0 0 0-.94-2.61c3.14-.35 6.44-1.54 6.44-7A5.44 5.44 0 0 0 20 4.77 5.07 5.07 0 0 0 19.91 1S18.73.65 16 2.48a13.38 13.38 0 0 0-7 0C6.27.65 5.09 1 5.09 1A5.07 5.07 0 0 0 5 4.77a5.44 5.44 0 0 0-1.5 3.78c0 5.42 3.3 6.61 6.44 7A3.37 3.37 0 0 0 9 18.13V22"/></svg><p>${msg}</p>`;
    return div;
}

// ── Collapsible sections ──────────────────────────────────────────────
function initCollapsible() {
    document.querySelectorAll('h3.collapsible').forEach(h3 => {
        const content = h3.nextElementSibling;
        if (!content || !content.classList.contains('collapsible-content')) return;

        const isCollapsed = h3.classList.contains('collapsed');
        if (!isCollapsed) {
            content.style.maxHeight = content.scrollHeight + 'px';
        } else {
            content.style.maxHeight = '0';
        }

        h3.addEventListener('click', () => {
            const collapsed = h3.classList.toggle('collapsed');
            content.classList.toggle('collapsed', collapsed);
            if (!collapsed) {
                content.style.maxHeight = content.scrollHeight + 'px';
            } else {
                content.style.maxHeight = '0';
            }
        });
    });
}

// ── Mobile sidebar ────────────────────────────────────────────────────
function initMobileSidebar() {
    const sidebar  = document.getElementById('sidebar');
    const overlay  = document.getElementById('sidebar-overlay');
    const menuBtn  = document.getElementById('mobile-menu-btn');
    const closeBtn = document.getElementById('sidebar-close');

    function open() {
        sidebar.classList.add('open');
        overlay.classList.add('visible');
    }
    function close() {
        sidebar.classList.remove('open');
        overlay.classList.remove('visible');
    }

    menuBtn?.addEventListener('click', open);
    closeBtn?.addEventListener('click', close);
    overlay?.addEventListener('click', close);
}

// ── Event wiring ──────────────────────────────────────────────────────
function initEvents() {
    // Tab switching
    document.querySelectorAll('.tab[data-tab]').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.tab[data-tab]').forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-panel').forEach(p => { p.style.display = 'none'; p.classList.remove('active'); });
            btn.classList.add('active');
            activeTab = btn.dataset.tab;
            const panel = document.getElementById('tab-' + activeTab);
            if (panel) { panel.style.display = 'block'; panel.classList.add('active'); }
            applyFiltersAndRender();
        });
    });

    // Search
    $jobSearch.addEventListener('input', () => {
        searchQuery = $jobSearch.value.trim();
        applyFiltersAndRender();
    });

    // Sort control
    $sortSelect.addEventListener('change', () => {
        activeSort = $sortSelect.value;
        applyFiltersAndRender();
    });

    // Show filter (All / Costable / Outliers)
    $filterSelect.addEventListener('change', () => {
        activeFilter = $filterSelect.value;
        applyFiltersAndRender();
    });

    // Date inputs
    $dateFrom.addEventListener('change', () => {
        activeDateFrom = new Date($dateFrom.value + 'T00:00:00');
        // deactivate preset buttons
        document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('active'));
        activePreset = null;
        loadData();
    });

    $dateTo.addEventListener('change', () => {
        activeDateTo = new Date($dateTo.value + 'T23:59:59');
        document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('active'));
        activePreset = null;
        loadData();
    });

    // Preset buttons
    document.querySelectorAll('.preset-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.preset-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            activePreset = btn.dataset.preset;
            const { start, end } = getDateRange(activePreset);
            activeDateFrom = start;
            activeDateTo   = end;
            $dateFrom.value = toDateStr(start);
            $dateTo.value   = toDateStr(end);
            loadData();
        });
    });

    // Refresh — trigger server-side sync then re-fetch
    $refreshBtn.addEventListener('click', async () => {
        $refreshBtn.classList.add('loading');
        $refreshBtn.textContent = 'Syncing…';
        try { await triggerSync(); } catch(e) { console.warn('Sync trigger failed:', e); }
        localStorage.removeItem(CACHE_KEY_JOBS);
        localStorage.removeItem(CACHE_KEY_CLOCK);
        jobDataCache   = null;
        clockDataCache = null;
        loadData(true).finally(() => {
            $refreshBtn.classList.remove('loading');
            $refreshBtn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg> Refresh Data`;
        });
    });
}

// ── Boot ──────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
    initCollapsible();
    initMobileSidebar();
    initEvents();
    initDateRange();
    loadData();
});
