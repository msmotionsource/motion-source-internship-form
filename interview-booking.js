/**
 * Interview booking — reads slots from SharePoint via Power Automate.
 * Replace FLOW URLs after creating flows (see INTERVIEW-BOOKING-SETUP.md).
 */
const INTERVIEW_BOOKING_CONFIG = {
  /** GET or POST — returns { slots: Slot[], booking: Booking | null } */
  listSlotsFlowUrl: "https://defaultdf0fc509acb44023b500fdf382dde4.30.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/555caa1460e845628fa5914c1914c272/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=03B9QTZ3PEzSgP3ogLinaGbROxWGGYqlVbOP7PZhM6E",
  /** POST — body: { email, slotId } — returns { success, booking } */
  bookSlotFlowUrl: "https://defaultdf0fc509acb44023b500fdf382dde4.30.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/1f3321816255458591d838b451ae7748/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=8z1uUxR7o02jmtn0veuRhLWqQcpcEoKzkwef5OwqK-E"
};

const EMAIL_PARAM_NAMES = ["email", "e"];

function getEmailFromUrl() {
  const params = new URLSearchParams(window.location.search);
  for (const name of EMAIL_PARAM_NAMES) {
    const value = params.get(name);
    if (value && value.trim()) {
      return decodeURIComponent(value.trim()).toLowerCase();
    }
  }
  return "";
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function showScreen(screenId) {
  const screens = [
    "missingEmailScreen",
    "loadingScreen",
    "errorScreen",
    "alreadyBookedScreen",
    "noSlotsScreen",
    "interviewBookingForm",
    "thankYouScreen"
  ];
  screens.forEach((id) => {
    const el = document.getElementById(id);
    if (el) {
      el.hidden = id !== screenId;
    }
  });
}

const DEFAULT_TIME_ZONE = "Europe/Skopje";

/** SharePoint / flow often sends labels like "Skopje CET" — not valid for Intl. */
function resolveTimeZone(timeZone) {
  if (!timeZone || typeof timeZone !== "string") {
    return DEFAULT_TIME_ZONE;
  }
  const trimmed = timeZone.trim();
  if (!trimmed) {
    return DEFAULT_TIME_ZONE;
  }
  try {
    Intl.DateTimeFormat(undefined, { timeZone: trimmed });
    return trimmed;
  } catch {
    const lower = trimmed.toLowerCase();
    if (lower.includes("skopje") || lower.includes("cet") || lower.includes("belgrade")) {
      return "Europe/Skopje";
    }
    return DEFAULT_TIME_ZONE;
  }
}

function parseInstant(isoString) {
  const date = new Date(isoString);
  return Number.isNaN(date.getTime()) ? null : date;
}

function getBrowserTimeZone() {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || DEFAULT_TIME_ZONE;
  } catch {
    return DEFAULT_TIME_ZONE;
  }
}

/** Offset from UTC in minutes for an IANA zone at a given instant (handles DST). */
function getTimeZoneOffsetMinutes(timeZone, date) {
  const d = date instanceof Date ? date : new Date(date);
  const formatter = new Intl.DateTimeFormat("en-US", {
    timeZone: resolveTimeZone(timeZone),
    timeZoneName: "shortOffset"
  });
  const tzPart = formatter.formatToParts(d).find((p) => p.type === "timeZoneName")?.value || "";
  const match = tzPart.match(/([+-])(\d{1,2})(?::?(\d{2}))?/);
  if (!match) {
    return 0;
  }
  const sign = match[1] === "+" ? 1 : -1;
  const hours = Number.parseInt(match[2], 10);
  const minutes = match[3] ? Number.parseInt(match[3], 10) : 0;
  return sign * (hours * 60 + minutes);
}

function formatUtcOffsetLabel(offsetMinutes) {
  if (offsetMinutes === 0) {
    return "UTC";
  }
  const sign = offsetMinutes > 0 ? "+" : "-";
  const abs = Math.abs(offsetMinutes);
  const h = Math.floor(abs / 60);
  const m = abs % 60;
  return `UTC${sign}${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}`;
}

function formatTimeZoneContext(timeZone, date) {
  const tz = resolveTimeZone(timeZone);
  const d = date instanceof Date ? date : new Date(date);
  const offset = getTimeZoneOffsetMinutes(tz, d);
  let longName = tz;
  try {
    longName =
      new Intl.DateTimeFormat("en-GB", { timeZone: tz, timeZoneName: "long" })
        .formatToParts(d)
        .find((p) => p.type === "timeZoneName")?.value || tz;
  } catch {
    /* use tz id */
  }
  return `${longName} (${formatUtcOffsetLabel(offset)})`;
}

function describeTimeZoneDifference(interviewTimeZone, date) {
  const interviewTz = resolveTimeZone(interviewTimeZone);
  const browserTz = getBrowserTimeZone();
  if (interviewTz === browserTz) {
    return null;
  }

  const diffMinutes = getTimeZoneOffsetMinutes(interviewTz, date) - getTimeZoneOffsetMinutes(browserTz, date);
  if (diffMinutes === 0) {
    return null;
  }

  const abs = Math.abs(diffMinutes);
  const h = Math.floor(abs / 60);
  const m = abs % 60;
  const parts = [];
  if (h) {
    parts.push(`${h} h`);
  }
  if (m) {
    parts.push(`${m} min`);
  }
  const span = parts.join(" ") || "0 min";

  if (diffMinutes > 0) {
    return `Interview time is ${span} ahead of your local time (${browserTz})`;
  }
  return `Interview time is ${span} behind your local time (${browserTz})`;
}

function formatWithTimeZone(date, options, timeZone) {
  const tz = resolveTimeZone(timeZone);
  try {
    return new Intl.DateTimeFormat("en-GB", { ...options, timeZone: tz }).format(date);
  } catch {
    return new Intl.DateTimeFormat("en-GB", { ...options, timeZone: DEFAULT_TIME_ZONE }).format(date);
  }
}

function formatSlotDateTime(isoString, timeZone, endIso) {
  const details = formatSlotTimeDetails(isoString, endIso, timeZone);
  const date = parseInstant(isoString);
  if (!date) {
    return details.primary;
  }
  const when = formatWithTimeZone(
    date,
    {
      weekday: "long",
      day: "numeric",
      month: "long",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
      timeZoneName: "short"
    },
    timeZone
  );
  if (!details.lines.length) {
    return when;
  }
  return `${when}\n${details.lines.join("\n")}`;
}

function formatDateGroupKey(isoString, timeZone) {
  const date = new Date(isoString);
  if (Number.isNaN(date.getTime())) {
    return "Unknown date";
  }
  return formatWithTimeZone(
    date,
    { weekday: "long", day: "numeric", month: "long", year: "numeric" },
    timeZone
  );
}

function formatTimeRange(startIso, endIso, timeZone) {
  const start = parseInstant(startIso);
  if (!start) {
    return String(startIso || "");
  }
  const opts = { hour: "2-digit", minute: "2-digit", timeZoneName: "short" };
  const startStr = formatWithTimeZone(start, opts, timeZone);
  const end = endIso ? parseInstant(endIso) : null;
  if (!end) {
    return startStr;
  }
  const endStr = formatWithTimeZone(end, opts, timeZone);
  return `${startStr} – ${endStr}`;
}

/** Build display lines from UTC instants + interview time zone. */
function formatSlotTimeDetails(startIso, endIso, timeZone) {
  const start = parseInstant(startIso);
  if (!start) {
    return { primary: String(startIso || ""), lines: [] };
  }

  const tz = resolveTimeZone(timeZone);
  const primary = formatTimeRange(startIso, endIso, tz);
  const lines = [formatTimeZoneContext(tz, start)];

  const browserTz = getBrowserTimeZone();
  if (browserTz !== tz) {
    const localRange = formatTimeRange(startIso, endIso, browserTz);
    lines.push(`Your local time: ${localRange}`);
    const diff = describeTimeZoneDifference(tz, start);
    if (diff) {
      lines.push(diff);
    }
  }

  return { primary, lines };
}

function appendSlotMetaLines(container, details) {
  const meta = document.createElement("div");
  meta.className = "ms-slot-option-meta";
  meta.textContent = details.primary;
  container.appendChild(meta);

  details.lines.forEach((line) => {
    const sub = document.createElement("div");
    sub.className = "ms-slot-option-meta-sub";
    sub.textContent = line;
    container.appendChild(sub);
  });
}

function setSlotTimeZoneIntro(slots) {
  const intro = document.getElementById("slotTimeZoneIntro");
  if (!intro || !slots.length) {
    return;
  }
  const tz = resolveTimeZone(slots[0].timeZone);
  const sample = parseInstant(slots[0].startDateTime);
  intro.textContent = `All times are shown in interview time (${formatTimeZoneContext(tz, sample || new Date())}).`;
}

function formatBookingTimeDisplay(startIso, endIso, timeZone) {
  const details = formatSlotTimeDetails(startIso, endIso, timeZone);
  const parts = [details.primary, ...details.lines];
  return parts.join("\n");
}

/** Map Power Automate / SharePoint PascalCase payload to what the UI expects. */
function normalizeSlot(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const id = raw.id ?? raw.ID ?? raw.ItemInternalId;
  const startDateTime = raw.startDateTime ?? raw.StartDateTime;
  const endDateTime = raw.endDateTime ?? raw.EndDateTime;
  const timeZone = resolveTimeZone(raw.timeZone ?? raw.TimeZone);
  const title =
    raw.title ?? raw.Title ?? raw.label ?? raw.Label ?? raw["{Name}"] ?? `Interview slot ${id}`;
  const maxBookings = raw.maxBookings ?? raw.MaxBookings ?? 1;
  const remaining = raw.remaining ?? raw.Remaining;
  const active = raw.isActive ?? raw.IsActive ?? raw.Active ?? raw.active;

  if (id == null || !startDateTime) {
    return null;
  }

  if (active === false) {
    return null;
  }

  const computedRemaining =
    remaining != null ? Number(remaining) : maxBookings != null ? Number(maxBookings) : 1;

  return {
    id: Number(id),
    title: String(title).trim() || `Interview slot ${id}`,
    label: String(title).trim() || `Interview slot ${id}`,
    startDateTime,
    endDateTime: endDateTime || null,
    timeZone,
    maxBookings: Number(maxBookings),
    remaining: computedRemaining,
    available: raw.available !== false && computedRemaining > 0
  };
}

function normalizeBooking(raw) {
  if (!raw || typeof raw !== "object") {
    return null;
  }

  const startDateTime = raw.startDateTime ?? raw.StartDateTime ?? raw.SlotStart;
  const endDateTime = raw.endDateTime ?? raw.EndDateTime ?? raw.SlotEnd ?? null;
  const slotId = raw.slotId ?? raw.SlotId ?? raw.ID;
  const title = raw.title ?? raw.Title ?? raw.label ?? raw.Label ?? "Interview";

  if (!startDateTime) {
    return null;
  }

  return {
    slotId: slotId != null ? Number(slotId) : undefined,
    title,
    label: raw.label ?? raw.Label ?? title,
    startDateTime,
    endDateTime,
    timeZone: resolveTimeZone(raw.timeZone ?? raw.TimeZone)
  };
}

function normalizeListSlotsResponse(data) {
  const rawSlots = Array.isArray(data?.slots) ? data.slots : Array.isArray(data?.value) ? data.value : [];
  const slots = rawSlots.map(normalizeSlot).filter(Boolean);
  const booking = data?.booking ? normalizeBooking(data.booking) : null;
  return { slots, booking };
}

function parseFlowJson(raw) {
  if (raw == null || raw === "") {
    return null;
  }
  if (typeof raw === "object") {
    return raw;
  }
  if (typeof raw === "string") {
    try {
      return JSON.parse(raw);
    } catch {
      return { message: raw };
    }
  }
  return null;
}

function getFlowErrorMessage(data, fallback) {
  const payload = parseFlowJson(data) ?? data;
  if (!payload || typeof payload !== "object") {
    return fallback;
  }
  if (typeof payload.message === "string" && payload.message) {
    return payload.message;
  }
  if (typeof payload.error === "string" && payload.error) {
    return payload.error;
  }
  if (payload.error && typeof payload.error.message === "string") {
    return payload.error.message;
  }
  return fallback;
}

/** Normalize book-slot flow output (many Power Automate / SharePoint shapes). */
function normalizeBookSlotResponse(data) {
  let payload = parseFlowJson(data) ?? data;

  if (!payload || typeof payload !== "object") {
    return { success: false, booking: null, message: "Empty response from booking service." };
  }

  // Some flows wrap the HTTP body again
  if (payload.body != null) {
    payload = parseFlowJson(payload.body) ?? payload.body;
  }

  if (payload.success === false) {
    return {
      success: false,
      booking: null,
      message: getFlowErrorMessage(payload, "Booking failed.")
    };
  }

  let rawBooking =
    payload.booking ??
    payload.Booking ??
    payload.data?.booking ??
    payload.data?.Booking ??
    null;

  // SharePoint "Create item" often returns the list row at the root
  const looksLikeBookingRow =
    rawBooking == null &&
    (payload.SlotStart != null ||
      payload.SlotId != null ||
      payload.startDateTime != null ||
      payload.CandidateEmail != null);

  if (looksLikeBookingRow) {
    rawBooking = payload;
  }

  // Sometimes the slot is returned instead of the booking row
  if (rawBooking == null && (payload.StartDateTime != null || payload.startDateTime != null) && payload.MaxBookings != null) {
    rawBooking = {
      slotId: payload.ID ?? payload.id,
      title: payload.Title ?? payload.title,
      startDateTime: payload.StartDateTime ?? payload.startDateTime,
      endDateTime: payload.EndDateTime ?? payload.endDateTime,
      timeZone: payload.TimeZone ?? payload.timeZone
    };
  }

  const booking = normalizeBooking(rawBooking);
  const explicitSuccess = payload.success === true || payload.Success === true;

  if (!booking && !explicitSuccess) {
    return {
      success: false,
      booking: null,
      message: getFlowErrorMessage(payload, "Booking could not be confirmed.")
    };
  }

  return {
    success: true,
    booking,
    message: payload.message ?? null
  };
}

function bookingFromSelectedInput(selected) {
  return {
    title: selected.dataset.label || "Interview",
    label: selected.dataset.label || "Interview",
    startDateTime: selected.dataset.start,
    endDateTime: selected.dataset.end || null,
    timeZone: resolveTimeZone(selected.dataset.timezone),
    slotId: selected.value ? Number(selected.value) : undefined
  };
}

function setBookingSpinner(show) {
  const spinner = document.getElementById("bookingSpinner");
  const btn = document.getElementById("confirmBookingBtn");
  spinner.hidden = !show;
  btn.disabled = show;
}

function setSlotSelectionError(show) {
  const list = document.getElementById("slotList");
  const error = document.getElementById("slotSelectionError");
  list.classList.toggle("is-invalid", show);
  error.hidden = !show;
}

function setMultilineText(elementId, text) {
  const el = document.getElementById(elementId);
  el.textContent = text;
  el.style.whiteSpace = "pre-line";
}

function renderExistingBooking(booking, email) {
  document.getElementById("bookedEmailDisplay").textContent = email;
  document.getElementById("bookedSlotLabel").textContent = booking.label || booking.title || "Interview";
  setMultilineText(
    "bookedSlotTime",
    formatSlotDateTime(booking.startDateTime, booking.timeZone, booking.endDateTime)
  );
  showScreen("alreadyBookedScreen");
}

function renderThankYou(booking, email) {
  document.getElementById("confirmedEmailDisplay").textContent = email;
  document.getElementById("confirmedSlotLabel").textContent = booking.label || booking.title || "Interview";
  setMultilineText(
    "confirmedSlotTime",
    formatSlotDateTime(booking.startDateTime, booking.timeZone, booking.endDateTime)
  );
  showScreen("thankYouScreen");
}

function renderSlots(slots) {
  const container = document.getElementById("slotList");
  container.innerHTML = "";

  const available = slots.filter((s) => s.available !== false && (s.remaining == null || s.remaining > 0));
  if (available.length === 0) {
    showScreen("noSlotsScreen");
    return;
  }

  const byDate = new Map();
  available.forEach((slot) => {
    const key = formatDateGroupKey(slot.startDateTime, slot.timeZone);
    if (!byDate.has(key)) {
      byDate.set(key, []);
    }
    byDate.get(key).push(slot);
  });

  byDate.forEach((dateSlots, dateLabel) => {
    const group = document.createElement("div");
    group.className = "ms-slot-date-group";

    const heading = document.createElement("div");
    heading.className = "ms-slot-date-heading";
    heading.textContent = dateLabel;
    group.appendChild(heading);

    dateSlots
      .sort((a, b) => new Date(a.startDateTime) - new Date(b.startDateTime))
      .forEach((slot) => {
        const id = `slot-${slot.id}`;
        const label = document.createElement("label");
        label.className = "ms-slot-option";
        label.setAttribute("for", id);

        const input = document.createElement("input");
        input.type = "radio";
        input.className = "form-check-input";
        input.name = "selectedSlot";
        input.id = id;
        input.value = String(slot.id);
        input.required = true;
        input.dataset.label = slot.title || slot.label || "";
        input.dataset.start = slot.startDateTime;
        input.dataset.timezone = slot.timeZone || DEFAULT_TIME_ZONE;
        input.dataset.end = slot.endDateTime || "";

        input.addEventListener("change", () => setSlotSelectionError(false));

        const body = document.createElement("div");
        body.className = "ms-slot-option-body";

        const title = document.createElement("div");
        title.className = "ms-slot-option-title";
        title.textContent = slot.title || slot.label || formatTimeRange(slot.startDateTime, slot.endDateTime, slot.timeZone);

        body.appendChild(title);
        appendSlotMetaLines(body, formatSlotTimeDetails(slot.startDateTime, slot.endDateTime, slot.timeZone));

        label.appendChild(input);
        label.appendChild(body);
        group.appendChild(label);
      });

    container.appendChild(group);
  });

  setSlotTimeZoneIntro(available);
  showScreen("interviewBookingForm");
}

async function fetchSlots(email) {
  if (
    INTERVIEW_BOOKING_CONFIG.listSlotsFlowUrl.includes("REPLACE_WITH") ||
    !INTERVIEW_BOOKING_CONFIG.listSlotsFlowUrl
  ) {
    throw new Error("Configure listSlotsFlowUrl in interview-booking.js (see INTERVIEW-BOOKING-SETUP.md).");
  }

  const response = await fetch(INTERVIEW_BOOKING_CONFIG.listSlotsFlowUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json", Accept: "application/json" },
    body: JSON.stringify({ email })
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(text || `Failed to load slots (${response.status})`);
  }

  const data = await response.json();
  return normalizeListSlotsResponse(data);
}

async function bookSlot(email, slotId) {
  if (
    INTERVIEW_BOOKING_CONFIG.bookSlotFlowUrl.includes("REPLACE_WITH") ||
    !INTERVIEW_BOOKING_CONFIG.bookSlotFlowUrl
  ) {
    throw new Error("Configure bookSlotFlowUrl in interview-booking.js (see INTERVIEW-BOOKING-SETUP.md).");
  }

  const response = await fetch(INTERVIEW_BOOKING_CONFIG.bookSlotFlowUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email, slotId: Number(slotId) })
  });

  const rawText = await response.text();
  const data = parseFlowJson(rawText) ?? {};

  if (!response.ok) {
    throw new Error(getFlowErrorMessage(data, `Booking failed (${response.status})`));
  }

  const result = normalizeBookSlotResponse(data);
  if (!result.success) {
    throw new Error(result.message || "Booking could not be confirmed.");
  }

  return result;
}

async function loadBookingPage() {
  const email = getEmailFromUrl();

  if (!email || !isValidEmail(email)) {
    showScreen("missingEmailScreen");
    return;
  }

  document.getElementById("candidateEmailDisplay").textContent = email;
  showScreen("loadingScreen");

  try {
    const data = await fetchSlots(email);

    if (data.booking) {
      renderExistingBooking(data.booking, email);
      return;
    }

    renderSlots(data.slots);
  } catch (err) {
    console.error(err);
    document.getElementById("errorMessage").textContent =
      err.message || "Something went wrong while loading time slots.";
    showScreen("errorScreen");
  }
}

document.getElementById("retryLoadBtn").addEventListener("click", loadBookingPage);

document.getElementById("interviewBookingForm").addEventListener("submit", async (e) => {
  e.preventDefault();

  const email = getEmailFromUrl();
  const selected = document.querySelector('input[name="selectedSlot"]:checked');

  if (!selected) {
    setSlotSelectionError(true);
    document.getElementById("slotList").scrollIntoView({ behavior: "smooth", block: "center" });
    return;
  }

  setSlotSelectionError(false);
  setBookingSpinner(true);

  try {
    const result = await bookSlot(email, selected.value);
    const booking = result.booking || bookingFromSelectedInput(selected);
    renderThankYou(booking, email);
  } catch (err) {
    console.error(err);
    alert(err.message || "Could not save your booking. The slot may have been taken. Please refresh and try again.");
    setBookingSpinner(false);
    await loadBookingPage();
  }
});

loadBookingPage();
