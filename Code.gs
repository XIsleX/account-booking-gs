function doGet(e) {
  if (e.parameter.page === 'app') {
    // If the user has logged in, show the main application.
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('KodeKiddo Account Booking v1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  } else {
    // By default, show the login page.
    const template = HtmlService.createTemplateFromFile('Login');
    return template.evaluate()
        .setTitle('KodeKiddo Booking Hub - Login')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
  }
}


function checkPassword(password) {
  const correctPassword = PropertiesService.getScriptProperties().getProperty('appPassword');

  if (password === correctPassword) {
    // On success, return the URL with the special parameter
    const url = ScriptApp.getService().getUrl() + '?page=app';
    return { ok: true, url: url };
  } else {
    return { ok: false, error: 'Invalid password. Please try again.' };
  }
}


function getSessionsForDate(dateStr) {
  const date = new Date(dateStr + 'T12:00:00');
  const day = date.getDay();
  return [
    { label: "Sesi 1", startISO: dateStr + "T09:00:00", endISO: dateStr + "T10:30:00" },
    { label: "Sesi 2", startISO: dateStr + "T10:30:00", endISO: dateStr + "T12:00:00" },
    { label: "Sesi 3", startISO: dateStr + "T13:00:00", endISO: dateStr + "T14:30:00" },
    { label: "Sesi 4", startISO: dateStr + "T14:30:00", endISO: dateStr + "T16:00:00" },
    { label: "Sesi 5", startISO: dateStr + "T16:00:00", endISO: dateStr + "T17:30:00" }
  ];
}

function getSpreadsheet() {
  const ssId = '1QgvC2AXkOGGM16PyoYZ5fEh6tKsrowQ2dZCQ2S0QQa8';
  return SpreadsheetApp.openById(ssId);
}

function getPlatforms() {
  const sheet = getSpreadsheet().getSheetByName('Accounts');
  if (!sheet) return ["No Accounts sheet found"];
  const values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
  const platforms = [...new Set(values.flat().filter(v => v))];
  return platforms.length > 0 ? platforms : ["No platforms available"];
}

function getAvailableAccounts(platform, dateStr) {
  const ss = getSpreadsheet();
  const accSheet = ss.getSheetByName('Accounts');
  const bookSheet = ss.getSheetByName('Bookings');

  if (!accSheet) return [];
  const allAccountsForPlatform = accSheet.getRange(2, 1, accSheet.getLastRow() - 1, 7).getValues()
    .filter(r => r[1] === platform)
    .map(r => ({ id: String(r[0]), name: String(r[2]), note: r[6] || "" }));

  if (!bookSheet || bookSheet.getLastRow() < 2) return allAccountsForPlatform;

  const bookingsData = bookSheet.getRange(2, 1, bookSheet.getLastRow() - 1, 4).getValues();
  const bookedAccountIdsForDay = new Set(
    bookingsData
      .filter(row => {
        if (!row[2]) return false;
        const bookingDate = new Date(row[2]);
        const bookingDateStr = Utilities.formatDate(bookingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        return bookingDateStr === dateStr;
      })
      .map(row => String(row[3]))
  );
  return allAccountsForPlatform.filter(acc => !bookedAccountIdsForDay.has(acc.id));
}

function bookAccount(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = getSpreadsheet();
    let bookSheet = ss.getSheetByName('Bookings');
    if (!bookSheet) bookSheet = ss.insertSheet('Bookings');

    if (bookSheet.getLastRow() > 1) {
      const bookingsData = bookSheet.getRange(2, 3, bookSheet.getLastRow() - 1, 2).getValues();
      const isAlreadyBooked = bookingsData.some(row => {
        const bookingDate = new Date(row[0]);
        const bookingDateStr = Utilities.formatDate(bookingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        return bookingDateStr === data.date && String(row[1]) === String(data.accountId);
      });
      if (isAlreadyBooked) return { ok: false, error: "This account was just booked. Please select another." };
    }

    const accSheet = ss.getSheetByName('Accounts');
    if (!accSheet) return { ok: false, error: "Accounts sheet missing" };
    if (bookSheet.getLastRow() === 0) bookSheet.appendRow(['Booker', 'Branch', 'Date', 'AccountID', 'Platform', 'StartISO', 'EndISO', 'Username', 'Password']);

    const accValues = accSheet.getRange(2, 1, accSheet.getLastRow() - 1, 7).getValues();
    const rowIndex = accValues.findIndex(r => String(r[0]) === String(data.accountId));
    if (rowIndex === -1) return { ok: false, error: "Account not found" };

    const row = accValues[rowIndex];
    const username = row[3], password = row[4], note = row[6] || '';
    const formattedDate = Utilities.formatDate(new Date(data.date), Session.getScriptTimeZone(), 'yyyy-MM-dd');

    bookSheet.appendRow([data.bookerName, data.branch, formattedDate, data.accountId, data.platform, data.sessionStartISO, data.sessionEndISO, username, password]);
    SpreadsheetApp.flush();
    return { ok: true, username, password, note };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function deleteBooking(accountId, date) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const bookSheet = getSpreadsheet().getSheetByName('Bookings');
    if (!bookSheet || bookSheet.getLastRow() < 2) return { ok: false, error: "No bookings found to delete." };
    const data = bookSheet.getRange(2, 1, bookSheet.getLastRow() - 1, 4).getValues();
    const rowIndex = data.findIndex(row => {
      const bookingDateStr = Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      return bookingDateStr === date && String(row[3]) === String(accountId);
    });
    if (rowIndex !== -1) {
      bookSheet.deleteRow(rowIndex + 2);
      SpreadsheetApp.flush();
      return { ok: true, message: "Booking successfully cancelled." };
    } else {
      return { ok: false, error: "Booking not found." };
    }
  } catch (e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getBookings(dateStr) {
  try {
    const bookSheet = getSpreadsheet().getSheetByName('Bookings');
    if (!bookSheet || bookSheet.getLastRow() < 2) return { ok: true, bookings: [] };
    const values = bookSheet.getRange(2, 1, bookSheet.getLastRow() - 1, 8).getValues();
    let bookings = values.map((row, index) => {
      const dateObject = new Date(row[2]);
      const formattedDate = !isNaN(dateObject.getTime()) ? Utilities.formatDate(dateObject, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'Invalid Date';
      return { rowNum: index + 2, booker: row[0], branch: row[1], date: formattedDate, platform: row[4], username: row[7] };
    });
    if (dateStr) bookings = bookings.filter(booking => booking.date === dateStr);
    bookings.sort((a, b) => new Date(b.date) - new Date(a.date));
    return { ok: true, bookings };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function deleteBookingByRow(rowNum) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const bookSheet = getSpreadsheet().getSheetByName('Bookings');
    if (!bookSheet) return { ok: false, error: "Bookings sheet not found." };
    if (rowNum < 2 || rowNum > bookSheet.getMaxRows()) return { ok: false, error: "Invalid booking." };
    bookSheet.deleteRow(rowNum);
    SpreadsheetApp.flush();
    return { ok: true, message: "Booking successfully cancelled." };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

