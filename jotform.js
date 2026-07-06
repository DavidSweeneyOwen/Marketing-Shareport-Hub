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
  bookedDates:  new Set(),   // "YYYY-MM-DD" strings
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

function parseSpBookings(items) {
  const dates  = new Set();
  const parsed = [];

  for (const it of (items || [])) {
    const rawDate = _firstField(it, ['BookingDate', 'Booking_x0020_Date', 'Date', 'VisitDate']);
    if (!rawDate) continue;
    const dateStr = String(rawDate).slice(0, 10);       // ISO → "YYYY-MM-DD"
    if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) continue;

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

function renderShowroomCalendar(containerId, year, month, bookedDates) {
  const container = document.getElementById(containerId);
  if (!container) return;

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

  // Current month days
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const dow     = new Date(year, month, d).getDay(); // 0=Sun
    const isWknd  = dow === 0 || dow === 6;
    const isToday = (d === today.getDate() && month === today.getMonth() && year === today.getFullYear());
    const isBooked = bookedDates.has(dateStr);
    const isFuture = new Date(year, month, d) > today;

    let cls = 'sd-d';
    if (isToday)       cls += ' today';
    else if (isBooked) cls += ' booked';
    else if (isWknd)   cls += ' wknd';
    else if (isFuture) cls += ' free';

    cells += `<div class="${cls}">${d}</div>`;
  }

  // Next month overflow (fill to complete last row)
  const totalCells = startDow + daysInMonth;
  const remainder  = totalCells % 7 === 0 ? 0 : 7 - (totalCells % 7);
  for (let d = 1; d <= remainder; d++) {
    cells += `<div class="sd-d muted wknd">${d}</div>`;
  }

  // Count free weekdays this month (not booked, not weekend, not past)
  let freeDays = 0;
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month, d).getDay();
    const isWknd = dow === 0 || dow === 6;
    const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
    const isFuture = new Date(year, month, d) > today;
    if (!isWknd && !bookedDates.has(dateStr) && isFuture) freeDays++;
  }

  container.innerHTML = `
    <div class="sd-cal-head">
      <div>
        <div class="sd-eyebrow">CHECKFIRE SHOWROOM</div>
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
      <span><span class="sw red"></span>Visit booked</span>
      <span><span class="sw green"></span>Available</span>
      <span><span class="sw today-sw"></span>Today</span>
    </div>
    <div class="sd-free-count">${freeDays} day${freeDays !== 1 ? 's' : ''} available this month</div>
  `;
}

function shiftShowroomMonth(delta) {
  JF.calendarMonth += delta;
  if (JF.calendarMonth > 11) { JF.calendarMonth = 0; JF.calendarYear++; }
  if (JF.calendarMonth < 0)  { JF.calendarMonth = 11; JF.calendarYear--; }
  renderShowroomCalendar('sd-cal-container', JF.calendarYear, JF.calendarMonth, JF.bookedDates);
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

async function loadShowroomData() {
  // Pull bookings from SharePoint using the signed-in user's token.
  // If not signed in yet, or the list isn't created, fall back to a
  // clean availability view (no scary errors).
  try {
    const listName = (HUB_CONFIG.showroom && HUB_CONFIG.showroom.list) || 'Showroom Bookings';
    const items = await fetchListItems(listName);
    const { dates, parsed } = parseSpBookings(items);
    JF.bookedDates = dates;
    JF.visits = parsed;
    JF.loaded = true;
  } catch (e) {
    console.info('[Showroom] bookings not loaded yet:', e.message);
    JF.loaded = false;
  }

  if (document.getElementById('sd-cal-container')) {
    renderShowroomCalendar('sd-cal-container', JF.calendarYear, JF.calendarMonth, JF.bookedDates);
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

// ─── Self-boot ────────────────────────────────────────────────
// Render an immediate (empty) calendar so the UI isn't blank; the
// real bookings load is triggered from app.js once sign-in completes.
setTimeout(loadShowroomData, 500);
