/**
 * MRTH Jobs calendar sync.
 *
 * One rule: a job has a calendar event exactly while its Status is
 * "Booked" and it has a date. Any other status — Delivered, Closed, back
 * to Quoted, whatever — means no event, so the calendar only ever shows
 * what's still upcoming, not a running history of everything that's
 * happened. Called from updateJob() after every field edit, so a status
 * change, a date change, or a materials swap all keep the event honest
 * without needing separate wiring for each case.
 *
 * Trailer-hire jobs use Start/End Date as a multi-day event; everything
 * else uses the single Scheduled date column.
 */

const JOB_CALENDAR_NAME = 'MRTH Jobs';

function jobCalendar_() {
  const existing = CalendarApp.getCalendarsByName(JOB_CALENDAR_NAME);
  return existing.length ? existing[0] : CalendarApp.createCalendar(JOB_CALENDAR_NAME);
}

function toCalDate_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function safeGetEvent_(cal, id) {
  try { return cal.getEventById(id); } catch (e) { return null; }
}

function syncJobCalendarEvent_(row) {
  const sh = sheet_();
  const map = headerMap_(sh);
  const lastCol = sh.getLastColumn();
  const values = sh.getRange(row, 1, 1, lastCol).getValues()[0];
  const get = function (field) {
    const i = map[field];
    if (i === undefined) return '';
    const v = values[i];
    return v === null || v === undefined ? '' : v;
  };

  const status = String(get('Status') || '');
  const eventId = String(get('Calendar Event ID') || '').trim();
  const isTrailer = String(get('Type') || '').toLowerCase().indexOf('trailer') !== -1;

  let start = null, end = null;
  if (isTrailer && toCalDate_(get('Start Date'))) {
    start = toCalDate_(get('Start Date'));
    end = toCalDate_(get('End Date')) || start;
  } else if (toCalDate_(get('Scheduled date'))) {
    start = toCalDate_(get('Scheduled date'));
    end = start;
  }

  const shouldExist = status === 'Booked' && !!start;
  const cal = jobCalendar_();

  if (!shouldExist) {
    if (eventId) {
      const existing = safeGetEvent_(cal, eventId);
      if (existing) existing.deleteEvent();
      sh.getRange(row, map['Calendar Event ID'] + 1).setValue('');
    }
    return;
  }

  const material = String(get('Material') || '').trim();
  const qty = String(get('Quantity') || '').trim();
  const suburb = String(get('Suburb') || '').trim();
  const load = qty && material ? (qty + ' ' + material) : (material || String(get('Type') || 'Job'));
  const title = load + (suburb ? ' — ' + suburb : '');

  const descLines = [];
  if (get('Name')) descLines.push('Customer: ' + get('Name'));
  if (get('Phone')) descLines.push('Phone: ' + String(get('Phone')).replace(/^'/, ''));
  if (get('Quoted $')) descLines.push('Quoted: $' + get('Quoted $'));
  if (get('Notes')) descLines.push('Notes: ' + get('Notes'));
  const description = descLines.join('\n');

  // All-day events are exclusive of their end date, so a same-day job
  // needs end == start + 1 day to actually show as one day.
  const endExclusive = new Date(end.getTime());
  endExclusive.setDate(endExclusive.getDate() + 1);

  const existing = eventId ? safeGetEvent_(cal, eventId) : null;
  if (existing) {
    existing.setTitle(title);
    existing.setDescription(description);
    existing.setAllDayDates(start, endExclusive);
  } else {
    const created = cal.createAllDayEvent(title, start, endExclusive, { description: description });
    sh.getRange(row, map['Calendar Event ID'] + 1).setValue(created.getId());
  }
}

/** Run once from the editor after pasting this in, to trigger the
 *  Calendar permissions prompt before deploying — same reason testSetup
 *  exists for Sheets. Creates and immediately deletes a throwaway event. */
function testCalendarSetup() {
  const cal = jobCalendar_();
  const ev = cal.createAllDayEvent('MRTH calendar test — safe to ignore', new Date());
  Logger.log('Created test event: ' + ev.getId());
  ev.deleteEvent();
  Logger.log('Deleted test event. Calendar "' + JOB_CALENDAR_NAME + '" is ready.');
  return 'OK';
}
