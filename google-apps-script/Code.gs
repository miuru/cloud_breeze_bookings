/**
 * Cloud Breeze booking receiver for Google Apps Script.
 *
 * Setup:
 * 1. Create a Google Sheet and copy its ID from the URL.
 * 2. Paste this file into Extensions > Apps Script for that sheet, or into a
 *    standalone Apps Script project that has access to the sheet.
 * 3. Set SPREADSHEET_ID below, then deploy as a Web app with access set to
 *    "Anyone". Copy the deployment URL into the website if it changes.
 */

const SPREADSHEET_ID = 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE';
const SHEET_NAME = 'Bookings';
const API_KEY = 'miuru97@7124';
const BOOKINGS_CACHE_KEY = 'cloud_breeze_bookings_v1';
const BOOKINGS_CACHE_SECONDS = 300;

const HEADERS = [
  'ID',
  'Guest Name',
  'Country',
  'Check-in Date',
  'Nights',
  'Booking Date',
  'Number of Guests',
  'Total Price (LKR)',
  'Booking Source',
  'Date Added',
  'Last Modified'
];

function doGet(e) {
  const data = (e && e.parameter) || {};
  const suppliedKey = data.key || data.apiKey;

  if (suppliedKey !== API_KEY) {
    return jsonResponse({ status: 'error', message: 'Unauthorized request.' });
  }

  if (data.action === 'getBookings') {
    return jsonResponse({ status: 'ok', bookings: getAllBookings() });
  }

  return jsonResponse({ status: 'ok', message: 'Cloud Breeze booking endpoint is ready.' });
}

function doPost(e) {
  try {
    const data = (e && e.parameter) || {};
    const suppliedKey = data.key || data.apiKey;

    if (suppliedKey !== API_KEY) {
      return jsonResponse({ status: 'error', message: 'Unauthorized request.' });
    }

    if (data.action === 'replaceAllBookings') {
      const bookings = JSON.parse(data.bookings || '[]');
      if (!Array.isArray(bookings)) throw new Error('Bookings must be an array.');
      replaceAllBookings(bookings);
      invalidateBookingsCache();
      return jsonResponse({ status: 'ok', message: 'All bookings replaced.', count: bookings.length });
    }

    if (data.action === 'clearAllBookings') {
      replaceAllBookings([]);
      invalidateBookingsCache();
      return jsonResponse({ status: 'ok', message: 'All bookings cleared.' });
    }

    if (data.action === 'deleteBooking') {
      const sheet = getBookingsSheet();
      const existingRow = findBookingRow(sheet, String(data.id || ''));
      if (!existingRow) {
        return jsonResponse({ status: 'ok', message: 'Booking was already absent.' });
      }
      sheet.deleteRow(existingRow);
      invalidateBookingsCache();
      return jsonResponse({ status: 'ok', message: 'Booking deleted.' });
    }

    if (data.action !== 'addBooking' && data.action !== 'updateBooking') {
      return jsonResponse({ status: 'error', message: 'Unsupported action.' });
    }

    const booking = data.booking ? JSON.parse(data.booking) : data;
    validateBooking(booking);

    const sheet = getBookingsSheet();
    const row = [
      String(booking.id || ''),
      booking.guestName,
      booking.country,
      booking.checkIn,
      Number(booking.nights),
      booking.bookingDate || '',
      Number(booking.numberOfGuests || 0),
      Number(booking.totalPrice),
      booking.bookingSource || '',
      booking.dateAdded || new Date().toISOString(),
      booking.lastModified || ''
    ];

    const existingRow = findBookingRow(sheet, String(booking.id || ''));
    if (existingRow) {
      sheet.getRange(existingRow, 1, 1, HEADERS.length).setValues([row]);
      invalidateBookingsCache();
      return jsonResponse({ status: 'ok', message: 'Booking updated.', row: existingRow });
    }

    sheet.appendRow(row);
    invalidateBookingsCache();
    return jsonResponse({ status: 'ok', message: 'Booking added.', row: sheet.getLastRow() });
  } catch (error) {
    console.error(error);
    return jsonResponse({ status: 'error', message: error.message });
  }
}

function getBookingsSheet() {
  if (SPREADSHEET_ID === 'PASTE_YOUR_GOOGLE_SHEET_ID_HERE') {
    throw new Error('Set SPREADSHEET_ID in Code.gs before deploying.');
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function findBookingRow(sheet, bookingId) {
  if (!bookingId || sheet.getLastRow() < 2) return null;

  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const index = ids.findIndex(id => String(id) === bookingId);
  return index === -1 ? null : index + 2;
}

function getAllBookings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(BOOKINGS_CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const sheet = getBookingsSheet();
  if (sheet.getLastRow() < 2) return [];

  const bookings = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues()
    .filter(row => row[0] !== '')
    .map(row => ({
      id: Number(row[0]) || row[0],
      guestName: String(row[1] || ''),
      country: String(row[2] || ''),
      checkIn: formatSheetDate(row[3]),
      nights: Number(row[4]) || 0,
      bookingDate: formatSheetDate(row[5]),
      numberOfGuests: Number(row[6]) || 0,
      totalPrice: Number(row[7]) || 0,
      bookingSource: String(row[8] || ''),
      dateAdded: formatSheetTimestamp(row[9]),
      lastModified: formatSheetTimestamp(row[10])
    }));

  const serialized = JSON.stringify(bookings);
  // Apps Script cache entries are limited to 100 KB, so skip caching very large lists.
  if (serialized.length <= 100000) cache.put(BOOKINGS_CACHE_KEY, serialized, BOOKINGS_CACHE_SECONDS);
  return bookings;
}

function invalidateBookingsCache() {
  CacheService.getScriptCache().remove(BOOKINGS_CACHE_KEY);
}

function replaceAllBookings(bookings) {
  const sheet = getBookingsSheet();
  bookings.forEach(validateBooking);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).clearContent();
  }

  if (!bookings.length) return;
  const rows = bookings.map(booking => [
    String(booking.id || ''),
    booking.guestName,
    booking.country,
    booking.checkIn,
    Number(booking.nights),
    booking.bookingDate || '',
    Number(booking.numberOfGuests || 0),
    Number(booking.totalPrice),
    booking.bookingSource || '',
    booking.dateAdded || new Date().toISOString(),
    booking.lastModified || ''
  ]);
  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
}

function formatSheetDate(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

function formatSheetTimestamp(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value.toISOString();
  }
  return String(value);
}

function validateBooking(booking) {
  const required = ['id', 'guestName', 'country', 'checkIn', 'nights', 'totalPrice'];
  const missing = required.filter(field => booking[field] === undefined || booking[field] === null || booking[field] === '');
  if (missing.length) throw new Error(`Missing required fields: ${missing.join(', ')}`);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
