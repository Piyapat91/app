/**
 * Google Apps Script API for AI Calorie & Weight Tracker (v2026)
 * Supports structured data storage for better readability in Google Sheets
 */

function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  let result = { success: false, message: "Invalid action" };

  if (action === 'read') {
    result = loadDataStructured();
  }

  const json = JSON.stringify(result);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    return ContentService.createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result = { success: false, message: "Invalid action" };

    if (action === 'sync') {
      result = syncDataStructured(data.payload);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Sync data into structured sheets
 */
function syncDataStructured(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Sync Entries
    if (payload.entries) {
      let entrySheet = ss.getSheetByName('Entries');
      if (!entrySheet) {
        entrySheet = ss.insertSheet('Entries');
        entrySheet.appendRow(['ID', 'Date', 'Meal', 'Name', 'Calories', 'Protein', 'Carbs', 'Fat']);
        entrySheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f3f4f6');
      }
      entrySheet.clearContents();
      entrySheet.appendRow(['ID', 'Date', 'Meal', 'Name', 'Calories', 'Protein', 'Carbs', 'Fat']);
      
      // Force Column B (Date) to be Plain Text to prevent Google from changing the format
      entrySheet.getRange("B:B").setNumberFormat("@");
      
      const rows = payload.entries.map(e => [e.id, String(e.date), e.meal, e.name, e.calories, e.protein, e.carbs, e.fat]);
      if (rows.length > 0) {
        entrySheet.getRange(2, 1, rows.length, 8).setValues(rows);
      }
    }

    // 2. Sync WeightData
    if (payload.weightData) {
      let weightSheet = ss.getSheetByName('Weights');
      if (!weightSheet) {
        weightSheet = ss.insertSheet('Weights');
        weightSheet.appendRow(['Date', 'Morning', 'Evening']);
        weightSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f3f4f6');
      }
      weightSheet.clearContents();
      weightSheet.appendRow(['Date', 'Morning', 'Evening']);
      const dates = Object.keys(payload.weightData).sort();
      const rows = dates.map(d => [d, payload.weightData[d].morning || "", payload.weightData[d].evening || ""]);
      if (rows.length > 0) {
        weightSheet.getRange(2, 1, rows.length, 3).setValues(rows);
      }
    }

    // 3. Sync Goals/Settings
    if (payload.goals) {
      let settingsSheet = ss.getSheetByName('Settings');
      if (!settingsSheet) {
        settingsSheet = ss.insertSheet('Settings');
        settingsSheet.appendRow(['Key', 'Value']);
      }
      settingsSheet.clearContents();
      settingsSheet.appendRow(['Key', 'Value']);
      const goalRows = Object.keys(payload.goals).map(k => [k, payload.goals[k]]);
      if (goalRows.length > 0) {
        settingsSheet.getRange(2, 1, goalRows.length, 2).setValues(goalRows);
      }
    }

    return { success: true, timestamp: new Date().toLocaleString() };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Load data from structured sheets back to JSON for the frontend
 */
function loadDataStructured() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = { entries: [], weightData: {}, goals: {} };

    // 1. Load Entries
    const entrySheet = ss.getSheetByName('Entries');
    if (entrySheet) {
      const data = entrySheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        result.entries.push({
          id: data[i][0],
          date: data[i][1] instanceof Date ? data[i][1].toISOString().split('T')[0] : data[i][1],
          meal: data[i][2],
          name: data[i][3],
          calories: data[i][4],
          protein: data[i][5],
          carbs: data[i][6],
          fat: data[i][7]
        });
      }
    }

    // 2. Load Weights
    const weightSheet = ss.getSheetByName('Weights');
    if (weightSheet) {
      const data = weightSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const dateStr = data[i][0] instanceof Date ? data[i][0].toISOString().split('T')[0] : data[i][0];
        result.weightData[dateStr] = {
          morning: data[i][1],
          evening: data[i][2]
        };
      }
    }

    // 3. Load Settings
    const settingsSheet = ss.getSheetByName('Settings');
    if (settingsSheet) {
      const data = settingsSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        result.goals[data[i][0]] = data[i][1];
      }
    }

    return result;
  } catch (e) {
    console.error("Load Error: " + e);
    return { entries: [], weightData: {}, goals: {} };
  }
}
