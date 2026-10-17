const SHEET_ID = '1AQC0pw7pclbCbXmcBeh-1hOkcJ8ujaaXhNrAJKdVu40';
const SHEET_TAB = 'FoodData';

function doGet(e) {
  const action = (e.parameter.action || '').toLowerCase();

  if (action === 'read') {
    return readData();
  }

  if (action === 'add') {
    return addEntry(
      e.parameter.restaurant,
      e.parameter.food,
      e.parameter.kid,
      e.parameter.lat,
      e.parameter.lng,
      e.parameter.address
    );
  }

  return json({ error: 'Unknown action' });
}

function readData() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);

    if (!sheet) {
      return json({ error: 'Sheet not found' });
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return json({ ok: true, entries: [] });
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 7).getDisplayValues();

    const entries = data.map(row => ({
      date: row[0],
      restaurant: row[1],
      food: row[2],
      kid: row[3],
      lat: row[4] ? parseFloat(row[4]) : null,
      lng: row[5] ? parseFloat(row[5]) : null,
      address: row[6] || ''
    })).filter(e => e.restaurant);

    return json({ ok: true, entries: entries });
  } catch (error) {
    return json({ error: error.toString() });
  }
}

function addEntry(restaurant, food, kid, lat, lng, address) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_TAB);

    if (!sheet) {
      return json({ error: 'Sheet not found' });
    }

    sheet.appendRow([
      new Date(),
      restaurant,
      food,
      kid,
      lat || '',
      lng || '',
      address || ''
    ]);

    return json({ ok: true });
  } catch (error) {
    return json({ error: error.toString() });
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
