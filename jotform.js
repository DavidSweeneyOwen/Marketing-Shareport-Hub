/**
 * CheckFire Marketing Hub — Jotform Showroom Integration
 * ────────────────────────────────────────────────────────
 * Fetches showroom booking submissions from Jotform EU API and
 * renders live calendar data + upcoming visits.
 * Also handles the booking modal with Jotform iframe embed.
 */

// ─── State ───────────────────────────────────────────────────

const JF = {
  submissions:  [],
  bookedDates:  new Set(),   // "YYYY-MM-DD" strings
  calendarYear: new Date().getFullYear(),
  calendarMonth: new Date().getMonth(), // 0-indexed
  loaded: false,
};

// ─── API Fetch ───────────────────────────────────────────────

async function fetchShowroomSubmissions() {
  const { apiKey, formId, apiBase } = HUB_CONFIG.jotform || {};
  if (!apiKey || apiKey === 'YOUR_JOTFORM_API_KEY') {
    console.info('[Jotform] API key not set — using demo data');
    return null;
  }
  try {
    const url = `${apiBase}/form/${formId}/submissions?apiKey=${apiKey}&limit=100&orderby=created_at&direction=DESC`;
    const res  = await fetch(url);
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    return data.content || [];
  } catch (e) {
    console.warn('[Jotform] Failed to fetch submissions:', e.message);
    // Surface error to visits panel
    const vc = document.getElementById('sd-visits-container');
    if (vc) vc.innerHTML = `<div class="sd-eyebrow">Upcoming visits</div><p class="sd-empty" style="color:#D1242B">API error: ${e.message}</p>`;
    return null;
  }
}

// ─── Parse Submissions ───────────────────────────────────────

/**
 * Extracts booking date and key info from a Jotform submission.
 * The Appointment field answer comes back as an object with a "date" key
 * (e.g. { date: "2026-05-20", time: "10:00 AM" }) or as a plain date string.
 */
function parseSubmission(sub) {
  const answers = sub.answers || {};
  let bookingDate  = null;
  let companyName  = '';
  let amEmail      = '';
  let customerNames = '';
  let numCustomers  = '';
  let arrivalTime   = '';

  for (const key of Object.keys(answers)) {
    const a = answers[key];
    const text = (a.text || '').toLowerCase();

    if (text.includes('appointment') || text.includes('date')) {
      const ans = a.answer;
      if (ans && typeof ans === 'object' && ans.date) {
        bookingDate = ans.date; // "YYYY-MM-DD"
      } else if (typeof ans === 'string' && /\d{4}-\d{2}-\d{2}/.test(ans)) {
        bookingDate = ans.match(/\d{4}-\d{2}-\d{2}/)[0];
      }
    }
    if (text.includes('company')) companyName = a.answer || '';
    if (text.includes('account manager')) amEmail = a.answer || '';
    if (text.includes('customer name')) customerNames = a.answer || '';
    if (text.includes('number of customer')) numCustomers = a.answer || '';
    if (text.includes('arrival')) arrivalTime = a.answer || '';
  }

  return { bookingDate, companyName, amEmail, customerNames, numCustomers, arrivalTime, id: sub.id };
}

function parseBookedDates(submissions) {
  const dates = new Set();
  const parsed = [];
  for (const sub of submissions) {
    const p = parseSubmission(sub);
    if (p.bookingDate) {
      dates.add(p.bookingDate);
      parsed.push(p);
    }
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
  // Jotform day: Monday = 0
  let startDow = firstDay.getDay() - 1; // JS Sunday=0, we want Mon=0
  if (startDow < 0) startDow = 6;

  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const daysInPrev  = new Date(year, month, 0).getDate();

  // Build day cells
  let cells = '';

  // Previous month overflow
  for (let i = startDow - 1; i >= 0; i--) {
    cells += `<div class="sd-d muted">${daysInPrev - i}</div>`;
  }

  // Current month days
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${year}-${String(month+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
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
    const dateStr = `${year}-${String(month+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
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

function renderUpcomingVisits(containerId, parsedBookings) {
  const container = document.getElementById(containerId);
  if (!container) return;

  const today    = new Date(); today.setHours(0,0,0,0);
  const upcoming = parsedBookings
    .filter(b => b.bookingDate && new Date(b.bookingDate) >= today)
    .sort((a, b) => a.bookingDate.localeCompare(b.bookingDate))
    .slice(0, 5);

  if (!upcoming.length) {
    container.innerHTML = '<p class="sd-empty">No upcoming visits — calendar is clear.</p>';
    return;
  }

  const html = upcoming.map(b => {
    const d    = new Date(b.bookingDate + 'T00:00:00');
    const day  = d.getDate();
    const mon  = MONTH_NAMES[d.getMonth()].slice(0,3);
    const name = b.companyName || 'Showroom visit';
    const amInitials = amEmailToInitials(b.amEmail);
    const visitors = b.numCustomers ? `${b.numCustomers} visitor${b.numCustomers > 1 ? 's' : ''}` : '';
    const time = b.arrivalTime || '';

    return `
      <li class="sd-visit">
        <div class="sd-date"><span class="sd-day">${day}</span><span class="sd-mon">${mon}</span></div>
        <div class="sd-info">
          <div class="sd-name">${escHtml(name)}</div>
          <div class="sd-meta">${[time, amInitials, visitors].filter(Boolean).join(' · ')}</div>
        </div>
        <div class="sd-avatars">
          ${amInitials ? `<span class="avatar small bg-red">${amEmailToInitials(b.amEmail, true)}</span>` : ''}
        </div>
      </li>`;
  }).join('');

  container.innerHTML = `
    <div class="sd-eyebrow">Upcoming visits</div>
    <ul class="sd-list">${html}</ul>
  `;
}

function amEmailToInitials(email, short = false) {
  if (!email) return '';
  const local = email.split('@')[0].replace(/[._]/g, ' ');
  const parts = local.split(' ').filter(Boolean);
  if (short) return parts.map(p => p[0]).join('').slice(0,2).toUpperCase();
  return parts.map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(' ');
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

  // Build iframe src — pre-fill account manager email if user is signed in
  const { formId } = HUB_CONFIG.jotform || {};
  let iframeSrc = `https://eu.jotform.com/${formId}?isIframe=1`;

  // Pre-fill the account manager email from logged-in user
  // Field name from Jotform: "accountManagersEmail" (dropdown)
  // Jotform pre-fill format: ?q{fieldOrder}_{fieldName}={value}
  // We match against the dropdown options by passing the email value
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
  // Jotform posts a message when the form is submitted
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

  // Refresh booking data after a short delay
  setTimeout(() => loadShowroomData(), 3000);
}

// ─── Main Load ───────────────────────────────────────────────

async function loadShowroomData() {
  const submissions = await fetchShowroomSubmissions();

  if (submissions === null) {
    // API unavailable — show a soft message in visits panel
    const vc = document.getElementById('sd-visits-container');
    if (vc) vc.innerHTML = '<div class="sd-eyebrow">Upcoming visits</div><p class="sd-empty" style="color:#D1242B">Could not connect to booking system.<br><span style="color:#AAA;font-size:11px">Check your Jotform API key in config.js</span></p>';
  } else if (submissions.length === 0) {
    const vc = document.getElementById('sd-visits-container');
    if (vc) vc.innerHTML = '<div class="sd-eyebrow">Upcoming visits</div><p class="sd-empty">No bookings found in Jotform yet.</p>';
  }

  if (submissions && submissions.length > 0) {
    const { dates, parsed } = parseBookedDates(submissions);
    JF.bookedDates  = dates;
    JF.submissions  = parsed;
    JF.loaded = true;
  }
  // If API not configured (demo mode), bookedDates stays empty — calendar still renders cleanly

  // Render home page showroom calendar
  if (document.getElementById('sd-cal-container')) {
    renderShowroomCalendar('sd-cal-container', JF.calendarYear, JF.calendarMonth, JF.bookedDates);
  }
  // Render home page upcoming visits
  if (document.getElementById('sd-visits-container')) {
    renderUpcomingVisits('sd-visits-container', JF.submissions);
  }
  // Render Trade page mini calendar sidebar
  renderMiniShowroomCalendar(JF.bookedDates);
  // Update quick-action free count
  updateFreeCount();
}

function updateFreeCount() {
  const el = document.getElementById('sd-free-stat');
  if (!el) return;
  const today = new Date(); today.setHours(0,0,0,0);
  const year  = JF.calendarYear;
  const month = JF.calendarMonth;
  let free = 0;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month, d).getDay();
    if (dow === 0 || dow === 6) continue;
    const dateStr = `${year}-${String(month+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    if (!JF.bookedDates.has(dateStr) && new Date(year, month, d) >= today) free++;
  }
  el.innerHTML = `<span class="sd-stat-num">${free}<span class="sd-stat-of">/ 22</span></span><div class="sd-stat-lbl">days free this month</div>`;
}

// ─── Utility ─────────────────────────────────────────────────

function escHtml(str) {
  return String(str)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;')
    .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

// ─── Trade Page Mini Calendar ─────────────────────────────────

/**
 * Renders the compact 4-row week-view in the Trade & Events sidebar card.
 * Shows the current month with today highlighted and booked days marked.
 */
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
    const dateStr = `${year}-${String(month+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
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
// Jotform data doesn't need Microsoft auth — load it as soon as
// the page is ready, regardless of sign-in state.
document.addEventListener('DOMContentLoaded', () => {
  // Small delay so the rest of the page scripts have initialised
  setTimeout(loadShowroomData, 800);
});
