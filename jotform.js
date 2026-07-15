/**
 * CheckFire Marketing Hub — Showroom Calendar + Booking
 * ────────────────────────────────────────────────────────
 * Two jobs:
 *  1. Booking modal — embeds the public Jotform booking form (iframe).
 *  2. Showroom calendar + "upcoming visits" — reads bookings from a
 *     SharePoint list ("Showroom Bookings") via fetchListItems() in
 *     graph.js, using the signed-in user's token.
 *
 * SECURITY: this file uses NO Jotform API key. The booking form is a
 * public iframe (no key needed). The "who's coming in" data is read
 * from SharePoint with the user's own login — exactly how the hub
 * reads Campaigns/Events. Nothing secret lives in the browser.
 *
 * Bookings get INTO the SharePoint list automatically via a Power
 * Automate flow on each Jotform submission — see
 * SHOWROOM-CALENDAR-SETUP.md. Jotform stays the booking system of
 * record; SharePoint is just the safe read-back path for the hub.
 */

// Immediate proof-of-life — runs synchronously as scripts load at bottom of <body>
(function () {
  const vc = document.getElementById('sd-visits-container');
  if (vc) {
    vc.innerHTML = '<div class="sd-eyebrow">Upcoming visits</div><p class="sd-empty">Loading showroom…</p>';
  }
})();

// ─── State ───────────────────────────────────────────────────

const JF = {
  bookedDates:  new Set(),   // "YYYY-MM-DD" showroom-booked days (mini calendar)
  marks:        new Map(),   // "YYYY-MM-DD" → { types:Set, labels:[] } (main calendar)
  visits:       [],          // parsed upcoming visits
  calendarYear: new Date().getFullYear(),
  calendarMonth: new Date().getMonth(), // 0-indexed
  loaded: false,
};

// ─── Read bookings from SharePoint ───────────────────────────
// SharePoint list items come back as their "fields" object. Column
// internal names are read defensively so small naming differences
// (e.g. Title vs CompanyName) still work. Create the list using the
// names in SHOWROOM-CALENDAR-SETUP.md for a clean match.

function _firstField(obj, names) {
  for (const n of names) {
    if (obj[n] !== undefined && obj[n] !== null && obj[n] !== '') return obj[n];
  }
  return '';
}


// Accepts a SharePoint Date column value (ISO) OR a text value written
// by Jotform's SharePoint integration (e.g. "2026-07-15 10:30" or UK
// "15/07/2026 10:30") and returns "YYYY-MM-DD", or '' if unreadable.
function _toIsoDate(raw) {
  const s = String(raw).trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);               // ISO first
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/); // UK day/month/year
  if (m) return m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0');
  const d = new Date(s);                                     // last resort
  return isNaN(d) ? '' : d.toISOString().slice(0, 10);
}

function parseSpBookings(items) {
  const dates  = new Set();
  const parsed = [];

  for (const it of (items || [])) {
    const rawDate = _firstField(it, ['BookingDate', 'Booking_x0020_Date', 'Date', 'VisitDate']);
    if (!rawDate) continue;
    const dateStr = _toIsoDate(rawDate);
    if (!dateStr) continue;

    dates.add(dateStr);
    parsed.push({
      bookingDate:   dateStr,
      companyName:   _firstField(it, ['Title', 'CompanyName', 'Company']) || 'Showroom visit',
      accountMgr:    _firstField(it, ['AccountManager', 'AccountManagerEmail', 'AM', 'Owner']),
      customerNames: _firstField(it, ['CustomerNames', 'Customers', 'Visitors']),
      numVisitors:   _firstField(it, ['NumberOfVisitors', 'NumVisitors', 'VisitorCount']),
      arrivalTime:   _firstField(it, ['ArrivalTime', 'Time']),
    });
  }
  return { dates, parsed };
}

// ─── Calendar Rendering ──────────────────────────────────────

const MONTH_NAMES = ['January','February','March','April','May','June',
                     'July','August','September','October','November','December'];
const DOW_SHORT   = ['M','T','W','T','F','S','S'];

// marks: Map "YYYY-MM-DD" → { types:Set('showroom'|'launch'|'campaign'), labels:[] }
function renderShowroomCalendar(containerId, year, month, marks) {
  const container = document.getElementById(containerId);
  if (!container) return;
  marks = marks || new Map();

  const today    = new Date();
  const firstDay = new Date(year, month, 1);
  let startDow = firstDay.getDay() - 1; // JS Sunday=0, we want Mon=0
  if (startDow < 0) startDow = 6;

  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const daysInPrev  = new Date(year, month, 0).getDate();

  let cells = '';

  // Previous month overflow
  for (let i = startDow - 1; i >= 0; i--) {
    cells += `<div class="sd-d muted">${daysInPrev - i}</div>`;
  }

  // Current month days. A day may carry showroom / launch / campaign
  // markers; showroom (booked) takes visual priority, then campaign,
  // then launch. The title shows what's on that day.
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dow     = new Date(year, month, d).getDay(); // 0=Sun
    const isWknd  = dow === 0 || dow === 6;
    const isToday = (d === today.getDate() && month === today.getMonth() && year === today.getFullYear());
    const mk      = marks.get(dateStr);
    const isFuture = new Date(year, month, d) > today;

    let cls = 'sd-d';
    if (isToday)                              cls += ' today';
    else if (mk && mk.types.has('showroom'))  cls += ' booked';
    else if (mk && mk.types.has('campaign'))  cls += ' camp';
    else if (mk && mk.types.has('launch'))    cls += ' launch';
    else if (isWknd)                          cls += ' wknd';
    else if (isFuture)                        cls += ' free';

    const title = mk ? escHtml(mk.labels.join(' · ')) : '';
    cells += `<div class="${cls}"${title ? ` title="${title}"` : ''}>${d}</div>`;
  }

  // Next month overflow (fill to complete last row)
  const totalCells = startDow + daysInMonth;
  const remainder  = totalCells % 7 === 0 ? 0 : 7 - (totalCells % 7);
  for (let d = 1; d <= remainder; d++) {
    cells += `<div class="sd-d muted wknd">${d}</div>`;
  }

  // Count free weekdays this month (no showroom booking, not weekend/past)
  let freeDays = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month, d).getDay();
    const isWknd = dow === 0 || dow === 6;
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const mk = marks.get(dateStr);
    const isFuture = new Date(year, month, d) > today;
    if (!isWknd && !(mk && mk.types.has('showroom')) && isFuture) freeDays++;
  }

  container.innerHTML = `
    <div class="sd-cal-head">
      <div>
        <div class="sd-eyebrow">CHECKFIRE MARKETING</div>
        <div class="sd-cal-title">${MONTH_NAMES[month]} ${year}</div>
      </div>
      <div class="sd-nav">
        <button class="cal-btn" aria-label="Previous" onclick="shiftShowroomMonth(-1)">‹</button>
        <button class="cal-btn" aria-label="Next"     onclick="shiftShowroomMonth(1)">›</button>
      </div>
    </div>
    <div class="sd-grid">
      ${DOW_SHORT.map(d => `<div class="sd-dow">${d}</div>`).join('')}
      ${cells}
    </div>
    <div class="sd-legend">
      <span><span class="sw red"></span>Showroom</span>
      <span><span class="sw blue"></span>Launch</span>
      <span><span class="sw amber"></span>Campaign</span>
      <span><span class="sw today-sw"></span>Today</span>
    </div>
    <div class="sd-free-count">${freeDays} showroom day${freeDays !== 1 ? 's' : ''} free this month</div>
  `;
}

function shiftShowroomMonth(delta) {
  JF.calendarMonth += delta;
  if (JF.calendarMonth > 11) { JF.calendarMonth = 0; JF.calendarYear++; }
  if (JF.calendarMonth < 0)  { JF.calendarMonth = 11; JF.calendarYear--; }
  renderShowroomCalendar('sd-cal-container', JF.calendarYear, JF.calendarMonth, JF.marks);
}

// ─── Upcoming Visits ─────────────────────────────────────────

function renderUpcomingVisits(containerId, visits) {
  const container = document.getElementById(containerId);
  if (!container) return;

  const today    = new Date(); today.setHours(0, 0, 0, 0);
  const upcoming = (visits || [])
    .filter(b => b.bookingDate && new Date(b.bookingDate) >= today)
    .sort((a, b) => a.bookingDate.localeCompare(b.bookingDate))
    .slice(0, 5);

  if (!upcoming.length) {
    container.innerHTML = '<div class="sd-eyebrow">Upcoming visits</div><p class="sd-empty">No upcoming visits — calendar is clear.</p>';
    return;
  }

  // All values come from SharePoint (staff-entered) but are still escaped.
  const html = upcoming.map(b => {
    const d    = new Date(b.bookingDate + 'T00:00:00');
    const day  = d.getDate();
    const mon  = MONTH_NAMES[d.getMonth()].slice(0, 3);
    const name = b.companyName || 'Showroom visit';
    const amName = personName(b.accountMgr);
    const n = parseInt(b.numVisitors, 10);
    const visitors = n > 0 ? `${n} visitor${n > 1 ? 's' : ''}` : '';
    const time = b.arrivalTime || '';

    return `
      <li class="sd-visit">
        <div class="sd-date"><span class="sd-day">${day}</span><span class="sd-mon">${escHtml(mon)}</span></div>
        <div class="sd-info">
          <div class="sd-name">${escHtml(name)}</div>
          <div class="sd-meta">${[time, amName, visitors].filter(Boolean).map(escHtml).join(' · ')}</div>
        </div>
        <div class="sd-avatars">
          ${amName ? `<span class="avatar small bg-red">${escHtml(personInitials(b.accountMgr))}</span>` : ''}
        </div>
      </li>`;
  }).join('');

  container.innerHTML = `
    <div class="sd-eyebrow">Upcoming visits</div>
    <ul class="sd-list">${html}</ul>
  `;
}

// Account manager may be an email ("jane.doe@checkfire.co.uk") or a name.
function personName(v) {
  if (!v) return '';
  const s = String(v).includes('@') ? String(v).split('@')[0].replace(/[._]/g, ' ') : String(v);
  return s.split(' ').filter(Boolean).map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
}

function personInitials(v) {
  const name = personName(v);
  if (!name) return '';
  return name.split(' ').filter(Boolean).map(p => p[0]).join('').slice(0, 2).toUpperCase();
}

// ─── Booking Modal ───────────────────────────────────────────

// Close on Escape key
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') {
    const modal = document.getElementById('booking-modal');
    if (modal && !modal.classList.contains('hidden')) closeBookingModal();
  }
});

function openBookingModal() {
  const modal = document.getElementById('booking-modal');
  if (!modal) return;

  // Build iframe src — public booking form, no API key involved.
  const { formId } = HUB_CONFIG.jotform || {};
  let iframeSrc = `https://eu.jotform.com/${encodeURIComponent(formId)}?isIframe=1`;

  // Pre-fill the account manager email from the signed-in user
  const userEmail = (window.AUTH && window.AUTH.account && window.AUTH.account.mail)
    ? window.AUTH.account.mail.toLowerCase()
    : null;
  if (userEmail) {
    iframeSrc += `&q4_pleaseConfirm=${encodeURIComponent(userEmail)}`;
  }

  const iframe = document.getElementById('booking-iframe');
  if (iframe && iframe.src !== iframeSrc) iframe.src = iframeSrc;

  modal.classList.remove('hidden');
  document.body.classList.add('modal-open');

  // Listen for Jotform submission message
  window.addEventListener('message', onJotformMessage);
}

function closeBookingModal() {
  const modal = document.getElementById('booking-modal');
  if (modal) modal.classList.add('hidden');
  document.body.classList.remove('modal-open');
  window.removeEventListener('message', onJotformMessage);

  // Reset success state
  const success = document.getElementById('booking-success');
  const iframe  = document.getElementById('booking-iframe');
  if (success) success.classList.add('hidden');
  if (iframe)  iframe.classList.remove('hidden');
}

function onJotformMessage(e) {
  // Only trust messages from Jotform's domain
  if (typeof e.origin === 'string' && !/^https:\/\/([a-z0-9-]+\.)?jotform\.com$/.test(e.origin)) return;
  if (typeof e.data !== 'string') return;
  try {
    const data = JSON.parse(e.data);
    if (data.action === 'submission-completed' || data.type === 'form-submit') {
      onBookingSubmitted();
    }
  } catch (_) {
    // Not JSON — ignore
  }
}

function onBookingSubmitted() {
  const iframe  = document.getElementById('booking-iframe');
  const success = document.getElementById('booking-success');
  if (iframe)  iframe.classList.add('hidden');
  if (success) success.classList.remove('hidden');

  // Refresh from SharePoint shortly — gives the Power Automate flow a
  // moment to write the new booking into the list. (Clear cache first.)
  setTimeout(() => {
    try { sessionStorage.removeItem('hubcache_list_' + (HUB_CONFIG.showroom?.list || 'Showroom Bookings')); } catch (_) {}
    loadShowroomData();
  }, 8000);
}

// ─── Main Load ───────────────────────────────────────────────

// Pull live submissions from the Azure Function proxy (if configured).
// The proxy holds the Jotform API key server-side; this call is
// anonymous and works even before the user signs in. Returns an array
// of booking objects in the same shape as parseSpBookings().
async function fetchProxyBookings(url) {
  const res = await fetch(url, { headers: { Accept: 'application/json' } });
  if (!res.ok) throw new Error('Proxy returned ' + res.status);
  const data = await res.json();
  const rows = Array.isArray(data) ? data : (data.bookings || data.value || []);
  return rows.map(r => ({
    bookingDate:   _toIsoDate(r.bookingDate || r.date || r.BookingDate || ''),
    companyName:   r.companyName || r.company || r.Title || 'Showroom visit',
    accountMgr:    r.accountMgr || r.accountManager || r.AccountManager || '',
    customerNames: r.customerNames || r.customers || '',
    numVisitors:   r.numVisitors || r.visitors || '',
    arrivalTime:   r.arrivalTime || r.time || '',
  })).filter(b => b.bookingDate);
}

function _dedupeBookings(list) {
  const seen = new Set();
  const out = [];
  for (const b of list) {
    const key = b.bookingDate + '|' + String(b.companyName).toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(b);
  }
  return out;
}

async function loadShowroomData() {
  // Two sources, merged: (1) the SharePoint "Showroom Bookings" list
  // (needs sign-in) and (2) the Jotform proxy (anonymous, if a proxyUrl
  // is set in config.js). Either can be absent without breaking the UI.
  let dates  = new Set();
  let parsed = [];
  let loaded = false;

  // 1) SharePoint list — read with the signed-in user's token.
  try {
    const listName = (HUB_CONFIG.showroom && HUB_CONFIG.showroom.list) || 'Showroom Bookings';
    const items = await fetchListItems(listName);
    const sp = parseSpBookings(items);
    sp.dates.forEach(d => dates.add(d));
    parsed.push(...sp.parsed);
    loaded = true;
  } catch (e) {
    console.info('[Showroom] SharePoint bookings not loaded yet:', e.message);
  }

  // 2) Jotform proxy — live submissions straight from Jotform.
  const proxyUrl = HUB_CONFIG.jotform && HUB_CONFIG.jotform.proxyUrl;
  if (proxyUrl) {
    try {
      const extra = await fetchProxyBookings(proxyUrl);
      extra.forEach(b => { if (b.bookingDate) dates.add(b.bookingDate); });
      parsed = _dedupeBookings(parsed.concat(extra));
      loaded = true;
    } catch (e) {
      console.info('[Showroom] Jotform proxy unavailable:', e.message);
    }
  }

  JF.bookedDates = dates;
  JF.visits = parsed;
  JF.loaded = loaded;

  // Build the combined marker map: showroom (red) + launches (blue) +
  // campaign runs (amber). List reads are cached (5 min) and any that
  // fail (e.g. before sign-in) are skipped without breaking the calendar.
  const marks = new Map();
  const addMark = (dateStr, type, label) => {
    if (!dateStr) return;
    const m = marks.get(dateStr) || { types: new Set(), labels: [] };
    m.types.add(type);
    if (label) m.labels.push(label);
    marks.set(dateStr, m);
  };
  parsed.forEach(b => addMark(b.bookingDate, 'showroom', b.companyName || 'Showroom visit'));
  try {
    const launches = await fetchListItems(HUB_CONFIG.lists.launches);
    launches.forEach(f => { const ds = _toIsoDate(f.LaunchDate); if (ds) addMark(ds, 'launch', 'Launch: ' + (f.Title || 'Untitled')); });
  } catch (e) { console.info('[Calendar] launches not marked:', e.message); }
  try {
    const camps = await fetchListItems(HUB_CONFIG.lists.campaigns);
    camps.forEach(f => {
      const s = _toIsoDate(f.StartDate), e = _toIsoDate(f.EndDate);
      if (s) addMark(s, 'campaign', 'Campaign starts: ' + (f.Title || 'Untitled'));
      if (e && e !== s) addMark(e, 'campaign', 'Campaign ends: ' + (f.Title || 'Untitled'));
    });
  } catch (e) { console.info('[Calendar] campaigns not marked:', e.message); }
  JF.marks = marks;

  if (document.getElementById('sd-cal-container')) {
    renderShowroomCalendar('sd-cal-container', JF.calendarYear, JF.calendarMonth, JF.marks);
  }

  const vc = document.getElementById('sd-visits-container');
  if (vc) {
    if (JF.loaded) {
      renderUpcomingVisits('sd-visits-container', JF.visits);
    } else {
      vc.innerHTML =
        '<div class="sd-eyebrow">Upcoming visits</div>' +
        '<p class="sd-empty">Sign in to see upcoming showroom visits.<br>' +
        '<span style="font-size:11px;color:#AAA">Use “Book the showroom” to request a date.</span></p>';
    }
  }

  renderMiniShowroomCalendar(JF.bookedDates);
  updateFreeCount();
}

function updateFreeCount() {
  const el = document.getElementById('sd-free-stat');
  if (!el) return;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const year  = JF.calendarYear;
  const month = JF.calendarMonth;
  let free = 0;
  let weekdays = 0;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month, d).getDay();
    if (dow === 0 || dow === 6) continue;
    weekdays++;
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    if (!JF.bookedDates.has(dateStr) && new Date(year, month, d) >= today) free++;
  }
  el.innerHTML = `<span class="sd-stat-num">${free}<span class="sd-stat-of">/ ${weekdays}</span></span><div class="sd-stat-lbl">days free this month</div>`;
}

// ─── Trade Page Mini Calendar ─────────────────────────────────

function renderMiniShowroomCalendar(bookedDates) {
  const weekContainer = document.getElementById('trade-showroom-week');
  const titleEl       = document.getElementById('trade-showroom-title');
  const badgeEl       = document.getElementById('trade-showroom-badge');
  if (!weekContainer) return;

  const today    = new Date();
  const year     = today.getFullYear();
  const month    = today.getMonth();
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  let firstDow = new Date(year, month, 1).getDay() - 1;
  if (firstDow < 0) firstDow = 6;

  let cells = '';

  // Prev month overflow
  const prevDays = new Date(year, month, 0).getDate();
  for (let i = firstDow - 1; i >= 0; i--) {
    cells += `<div class="day muted">${prevDays - i}</div>`;
  }

  let freeDays = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month, d).getDay();
    const isWknd  = dow === 0 || dow === 6;
    const isToday = d === today.getDate();
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const isBooked = bookedDates.has(dateStr);
    const isFuture = new Date(year, month, d) >= today;

    let cls = 'day';
    if (isToday)       cls += ' today';
    else if (isBooked) cls += ' booked';
    else if (isWknd)   cls += ' muted';

    if (!isWknd && !isBooked && isFuture) freeDays++;

    cells += `<div class="${cls}">${d}</div>`;
  }

  // Next month overflow
  const total = firstDow + daysInMonth;
  const remainder = total % 7 === 0 ? 0 : 7 - (total % 7);
  for (let d = 1; d <= remainder; d++) {
    cells += `<div class="day muted">${d}</div>`;
  }

  const headerRow = '<div class="dow">M</div><div class="dow">T</div><div class="dow">W</div><div class="dow">T</div><div class="dow">F</div><div class="dow">S</div><div class="dow">S</div>';
  weekContainer.innerHTML = headerRow + cells;

  if (titleEl) {
    titleEl.textContent = `Bookings · ${MONTH_NAMES[month]} ${year}`;
  }
  if (badgeEl) {
    badgeEl.innerHTML = `<span class="status-dot ${freeDays > 0 ? 'green' : 'red'}"></span>${freeDays} free`;
    badgeEl.className = `badge ${freeDays > 0 ? 'green' : 'red'}`;
  }
}

// ─── Calendar subscribe / export ─────────────────────────────
// If a live feed URL is configured (published Outlook/SharePoint
// calendar), Subscribe opens it (a true auto-updating subscription).
// Otherwise it builds and downloads an .ics of the current marked
// dates, which imports into Outlook / Google / Apple Calendar.
function _icsDate(dateStr) {
  return String(dateStr).replace(/-/g, '');           // YYYYMMDD
}
function _icsNextDay(dateStr) {
  const d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + 1);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}
function _icsEscape(s) {
  return String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n');
}

function buildCalendarICS() {
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d+Z$/, 'Z');
  const lines = [
    'BEGIN:VCALENDAR', 'VERSION:2.0',
    'PRODID:-//CheckFire//Marketing Hub//EN',
    'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
    'X-WR-CALNAME:CheckFire Marketing',
  ];
  let n = 0;
  for (const [dateStr, mk] of JF.marks.entries()) {
    const summary = mk.labels.length ? mk.labels.join(' · ')
                  : (mk.types.has('showroom') ? 'Showroom visit' : mk.types.has('campaign') ? 'Campaign' : 'Product launch');
    lines.push(
      'BEGIN:VEVENT',
      `UID:cf-${_icsDate(dateStr)}-${n++}@checkfire-hub`,
      `DTSTAMP:${stamp}`,
      `DTSTART;VALUE=DATE:${_icsDate(dateStr)}`,
      `DTEND;VALUE=DATE:${_icsNextDay(dateStr)}`,
      `SUMMARY:${_icsEscape(summary)}`,
      'END:VEVENT'
    );
  }
  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}

function subscribeCalendar() {
  const feed = HUB_CONFIG.calendar && HUB_CONFIG.calendar.feedUrl;
  if (feed) {
    // Live subscription: hand the webcal/https feed to the OS calendar.
    window.open(feed, '_blank', 'noopener');
    if (typeof showToast === 'function') showToast('Opening your calendar app to subscribe…');
    return;
  }
  if (!JF.marks || JF.marks.size === 0) {
    if (typeof showToast === 'function') showToast('No calendar dates to export yet');
    return;
  }
  const blob = new Blob([buildCalendarICS()], { type: 'text/calendar;charset=utf-8' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = 'checkfire-marketing-calendar.ics';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  if (typeof showToast === 'function') showToast('Calendar downloaded — open it to add to Outlook');
}

// ─── Self-boot ────────────────────────────────────────────────
// Render an immediate (empty) calendar so the UI isn't blank; the
// real bookings load is triggered from app.js once sign-in completes.
setTimeout(loadShowroomData, 500);
