/**
 * מעקב פיננסי - Backend (Google Apps Script)
 * -------------------------------------------
 * התקנה:
 * 1. פתחו את קובץ הגיליון ב-Google Sheets (אחרי שתעלו את money_tracker_google_sheets.xlsx
 *    ותמירו אותו ל-Google Sheets: קובץ > שמירה כ-Google Sheets).
 * 2. תפריט Extensions / תוספים > Apps Script.
 * 3. מחקו את כל הקוד שיש שם, והדביקו את כל הקובץ הזה במקומו.
 * 4. שמרו (Ctrl+S), ותנו שם לפרויקט, למשל "money-tracker-api".
 * 5. לחצו על Deploy > New deployment.
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone (או Anyone with the link)
 * 6. אשרו הרשאות (Google יזהיר שזה סקריפט לא מאומת - זה תקין, זה הסקריפט שלכם).
 * 7. העתיקו את ה-Web app URL שמתקבל - זה הכתובת שתדביקו בהגדרות ה-HTML.
 *
 * חשוב: כל שינוי בקוד מחייב Deploy > Manage deployments > עריכה (עיפרון) > גרסה חדשה > Deploy,
 * אחרת ה-URL הישן ימשיך להריץ את הגרסה הישנה.
 */

const SHEET_NAME = 'תנועות';
const CATEGORIES_SHEET_NAME = 'קטגוריות';
const BUDGETS_SHEET_NAME = 'תקציבים';
const RECURRING_SHEET_NAME = 'תנועות קבועות';
const SETTINGS_SHEET_NAME = 'הגדרות';

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}
function getNamedSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateCell_(dateVal) {
  if (dateVal && typeof dateVal.getTime === 'function') {
    return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(dateVal);
}

// בודק אם כבר קיימת שורה עם אותו clientId בעמודה F, בטווח השורות האחרונות (עד 60 אחורה).
// זה מה שהופך "הוספה" לבטוחה לניסיון חוזר: אם הבקשה כבר הצליחה בעבר אבל התשובה לא הגיעה
// ללקוח בזמן, ניסיון נוסף לא ייצור שורה כפולה - הוא פשוט ימצא את השורה הקיימת ויחזיר אותה.
function findExistingByClientId_(sheet, clientId) {
  if (!clientId) return -1;
  const lastRow = sheet.getLastRow();
  const scanStart = Math.max(2, lastRow - 60 + 1);
  if (lastRow < scanStart) return -1;
  const ids = sheet.getRange(scanStart, 6, lastRow - scanStart + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === clientId) return scanStart + i;
  }
  return -1;
}

// GET - מחזיר את כל התנועות, הקטגוריות, התקציבים, התנועות הקבועות וההגדרות
function doGet(e) {
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    const rows = [];

    if (lastRow >= 2) {
      const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
      data.forEach(function (r, idx) {
        const dateVal = r[0];
        if (!dateVal) return; // דלג על שורות ריקות
        rows.push({
          row: idx + 2,
          date: formatDateCell_(dateVal),
          type: r[1],
          category: r[2],
          amount: Number(r[3]) || 0,
          note: r[4] || ''
        });
      });
    }

    const catSheet = getNamedSheet_(CATEGORIES_SHEET_NAME);
    let categories = [];
    if (catSheet) {
      const catLast = catSheet.getLastRow();
      if (catLast >= 2) {
        categories = catSheet.getRange(2, 1, catLast - 1, 2).getValues()
          .filter(function (r) { return r[0] && r[1]; })
          .map(function (r) { return { type: r[0], category: r[1] }; });
      }
    }

    const budgetSheet = getNamedSheet_(BUDGETS_SHEET_NAME);
    let budgets = [];
    if (budgetSheet) {
      const bLast = budgetSheet.getLastRow();
      if (bLast >= 2) {
        budgets = budgetSheet.getRange(2, 1, bLast - 1, 3).getValues()
          .filter(function (r) { return r[0] && r[1] && r[2]; })
          .map(function (r) { return { type: r[0], category: r[1], budget: Number(r[2]) || 0 }; });
      }
    }

    const recSheet = getNamedSheet_(RECURRING_SHEET_NAME);
    let recurring = [];
    if (recSheet) {
      const rLast = recSheet.getLastRow();
      if (rLast >= 2) {
        recurring = recSheet.getRange(2, 1, rLast - 1, 5).getValues()
          .map(function (r, idx) { return { row: idx + 2, type: r[0], category: r[1], amount: Number(r[2]) || 0, note: r[3] || '', day: Number(r[4]) || 0 }; })
          .filter(function (r) { return r.type && r.category && r.amount; });
      }
    }

    const setSheet = getNamedSheet_(SETTINGS_SHEET_NAME);
    let settings = {};
    if (setSheet) {
      const sLast = setSheet.getLastRow();
      if (sLast >= 2) {
        setSheet.getRange(2, 1, sLast - 1, 2).getValues().forEach(function (r) {
          if (r[0]) settings[r[0]] = r[1];
        });
      }
    }

    return jsonResponse_({ ok: true, transactions: rows, categories: categories, budgets: budgets, recurring: recurring, settings: settings });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

// POST - ראו action לפירוט הפעולות הנתמכות
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const sheet = getSheet_();

    if (body.action === 'delete') {
      const rowToDelete = Number(body.row);
      if (!rowToDelete || rowToDelete < 2) {
        return jsonResponse_({ ok: false, error: 'invalid row' });
      }
      sheet.deleteRow(rowToDelete);
      return jsonResponse_({ ok: true });
    }

    if (body.action === 'edit') {
      const rowToEdit = Number(body.row);
      if (!rowToEdit || rowToEdit < 2) {
        return jsonResponse_({ ok: false, error: 'invalid row' });
      }
      const date = body.date || '';
      const type = body.type || '';
      const category = body.category || '';
      const amount = Number(body.amount) || 0;
      const note = body.note || '';
      sheet.getRange(rowToEdit, 1, 1, 5).setValues([[date, type, category, amount, note]]);
      return jsonResponse_({ ok: true });
    }

    if (body.action === 'addCategory') {
      const catSheet = getNamedSheet_(CATEGORIES_SHEET_NAME);
      const type = (body.type || '').trim();
      const category = (body.category || '').trim();
      if (!type || !category) {
        return jsonResponse_({ ok: false, error: 'missing type or category' });
      }
      const lastRow = catSheet.getLastRow();
      if (lastRow >= 2) {
        const existing = catSheet.getRange(2, 1, lastRow - 1, 2).getValues();
        const dup = existing.some(function (r) { return r[0] === type && r[1] === category; });
        if (dup) return jsonResponse_({ ok: true, duplicate: true });
      }
      catSheet.appendRow([type, category]);
      return jsonResponse_({ ok: true });
    }

    if (body.action === 'setBudget') {
      const budgetSheet = getNamedSheet_(BUDGETS_SHEET_NAME);
      const type = (body.type || '').trim();
      const category = (body.category || '').trim();
      const budget = Number(body.budget) || 0;
      if (!type || !category) {
        return jsonResponse_({ ok: false, error: 'missing type or category' });
      }
      const lastRow = budgetSheet.getLastRow();
      let foundRow = -1;
      if (lastRow >= 2) {
        const existing = budgetSheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (let i = 0; i < existing.length; i++) {
          if (existing[i][0] === type && existing[i][1] === category) { foundRow = i + 2; break; }
        }
      }
      if (budget <= 0) {
        if (foundRow > 0) budgetSheet.deleteRow(foundRow);
        return jsonResponse_({ ok: true });
      }
      if (foundRow > 0) {
        budgetSheet.getRange(foundRow, 3).setValue(budget);
      } else {
        budgetSheet.appendRow([type, category, budget]);
      }
      return jsonResponse_({ ok: true });
    }

    if (body.action === 'addRecurring') {
      const recSheet = getNamedSheet_(RECURRING_SHEET_NAME);
      const type = (body.type || '').trim();
      const category = (body.category || '').trim();
      const amount = Number(body.amount) || 0;
      const note = body.note || '';
      const day = Number(body.day) || 0;
      const clientId = body.clientId || '';
      if (!type || !category || !amount) {
        return jsonResponse_({ ok: false, error: 'missing fields' });
      }
      const existingRow = findExistingByClientId_(recSheet, clientId);
      if (existingRow > 0) {
        return jsonResponse_({ ok: true, row: existingRow, duplicate: true });
      }
      recSheet.appendRow([type, category, amount, note, day || '', clientId]);
      return jsonResponse_({ ok: true, row: recSheet.getLastRow() });
    }

    if (body.action === 'deleteRecurring') {
      const recSheet = getNamedSheet_(RECURRING_SHEET_NAME);
      const rowToDelete = Number(body.row);
      if (!rowToDelete || rowToDelete < 2) {
        return jsonResponse_({ ok: false, error: 'invalid row' });
      }
      recSheet.deleteRow(rowToDelete);
      return jsonResponse_({ ok: true });
    }

    if (body.action === 'editRecurring') {
      const recSheet = getNamedSheet_(RECURRING_SHEET_NAME);
      const rowToEdit = Number(body.row);
      if (!rowToEdit || rowToEdit < 2) {
        return jsonResponse_({ ok: false, error: 'invalid row' });
      }
      const type = (body.type || '').trim();
      const category = (body.category || '').trim();
      const amount = Number(body.amount) || 0;
      const note = body.note || '';
      const day = Number(body.day) || 0;
      if (!type || !category || !amount) {
        return jsonResponse_({ ok: false, error: 'missing fields' });
      }
      recSheet.getRange(rowToEdit, 1, 1, 5).setValues([[type, category, amount, note, day || '']]);
      return jsonResponse_({ ok: true });
    }

    if (body.action === 'setSetting') {
      const setSheet = getNamedSheet_(SETTINGS_SHEET_NAME);
      const key = (body.key || '').trim();
      const value = body.value;
      if (!key) {
        return jsonResponse_({ ok: false, error: 'missing key' });
      }
      const lastRow = setSheet.getLastRow();
      let foundRow = -1;
      if (lastRow >= 2) {
        const existing = setSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < existing.length; i++) {
          if (existing[i][0] === key) { foundRow = i + 2; break; }
        }
      }
      if (foundRow > 0) {
        setSheet.getRange(foundRow, 2).setValue(value);
      } else {
        setSheet.appendRow([key, value]);
      }
      return jsonResponse_({ ok: true });
    }

    // הוספת תנועה (ברירת מחדל)
    const date = body.date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const type = body.type || '';
    const category = body.category || '';
    const amount = Number(body.amount) || 0;
    const note = body.note || '';
    const clientId = body.clientId || '';

    const existingRow = findExistingByClientId_(sheet, clientId);
    if (existingRow > 0) {
      return jsonResponse_({ ok: true, row: existingRow, duplicate: true });
    }

    sheet.appendRow([date, type, category, amount, note, clientId]);
    return jsonResponse_({ ok: true, row: sheet.getLastRow() });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}
