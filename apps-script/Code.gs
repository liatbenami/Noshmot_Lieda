/**
 * נושמות לידה — מקבל הרשמות מהטופס באתר ושומר אותן בגיליון.
 *
 * התקנה - שתי דרכים, שתיהן עובדות:
 *  א. מתוך הגיליון: תוספים (Extensions) -> Apps Script
 *  ב. או ישירות: script.google.com/home -> New project -> מדביקים כאן
 *
 *  1. פותחים את עורך הסקריפטים באחת משתי הדרכים
 *  2. מדביקים את הקוד הזה במקום כל מה שיש שם, ושומרים
 *  3. Deploy -> New deployment -> type: Web app
 *     Execute as: Me  |  Who has access: Anyone
 *  4. מעתיקים את ה-Web app URL ומדביקים אותו ב-index.html בקבוע SCRIPT_URL
 */

var NOTIFY_EMAIL = 'doriternest@gmail.com';  // לאן נשלחת התראה על כל הרשמה
var RETREAT_CAPACITY = 15;                   // מספר המקומות בריטריט

/* מזהה הגיליון. openById עובד גם מסקריפט עצמאי וגם מסקריפט שמחובר לגיליון,
   בניגוד ל-getActiveSpreadsheet שעובד רק במחובר. */
var SHEET_ID = '1Of5lSDXxCA4EI_Ddb55XpUpgDnMcdSUyJY_J1-9C3Ik';

function doPost(e) {
  try {
    var p = (e && e.parameter) || {};

    // שדות חובה — מגן גם אם מישהו עוקף את הוולידציה בדפדפן
    if (!p.full_name || !p.phone || !p.email) {
      return json({ ok: false, error: 'missing_fields' });
    }

    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    var status = evaluateRegistration(sheet, p);

    if (status.reject) {
      return json({ ok: false, error: status.reject, message: status.message });
    }

    sheet.appendRow([
      Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm'),
      p.full_name,
      "'" + String(p.phone).trim(),   // גרש מונע מגוגל לאכול את ה-0 המוביל
      p.email,
      p.pregnancy_week || '',
      'טרם שולם',
      status.note || ''
    ]);

    notify(p, status);
    return json({ ok: true, message: status.message || '' });

  } catch (err) {
    // נכשל? רושמים ביומן כדי שלא נאבד את הפנייה בשקט
    console.error('registration failed: ' + err + ' | data: ' + JSON.stringify(e && e.parameter));
    return json({ ok: false, error: 'server_error' });
  }
}

/**
 * מחליט מה לעשות עם הרשמה חדשה: לקבל, לקבל עם הערה, או לדחות.
 *
 * מקבל:  sheet - הגיליון (שורה 1 = כותרות, עמודה B = שם, D = אימייל, F = סטטוס)
 *        p     - פרטי הטופס: full_name, phone, email, pregnancy_week
 *
 * מחזיר: { reject: false|'מזהה-שגיאה', message: 'טקסט לנרשמת', note: 'הערה לגיליון' }
 *
 * TODO (ליאת): כאן ההחלטה העסקית. כרגע מוגדרת התנהגות ברירת־מחדל שמרנית
 * שמקבלת הכל ורק מסמנת בהערה. שווה לשקול:
 *   - מה קורה כשנרשמת כבר קיימת באותו אימייל? לדחות ("כבר נרשמת")
 *     או לקבל כהרשמה נוספת (אולי היא רושמת חברה)?
 *   - מה קורה כשעוברים את RETREAT_CAPACITY? לדחות, או לקבל לרשימת
 *     המתנה ולהגיד לה את זה מראש כדי שלא תשלם בביט לחינם?
 * השורה התחתונה: דחייה מונעת תשלום מבולבל, קבלה לא מפספסת לקוחה.
 */
function evaluateRegistration(sheet, p) {
  var rows = sheet.getLastRow() - 1;  // בלי שורת הכותרות
  var existing = rows > 0 ? sheet.getRange(2, 4, rows, 1).getValues() : [];
  var isDuplicate = existing.some(function (r) {
    return String(r[0]).trim().toLowerCase() === String(p.email).trim().toLowerCase();
  });
  var isOverCapacity = rows >= RETREAT_CAPACITY;

  return {
    reject: false,
    message: '',
    note: [isDuplicate ? 'אימייל כפול' : '', isOverCapacity ? 'מעל התפוסה' : ''].filter(String).join(' / ')
  };
}

function notify(p, status) {
  if (!NOTIFY_EMAIL) return;
  try {
    MailApp.sendEmail({
      to: NOTIFY_EMAIL,
      subject: 'הרשמה חדשה לריטריט: ' + p.full_name,
      htmlBody:
        '<div dir="rtl" style="font-family:Arial,sans-serif;font-size:15px">' +
        '<h2 style="color:#731c38">הרשמה חדשה לנושמות לידה</h2>' +
        '<p><b>שם:</b> ' + p.full_name + '</p>' +
        '<p><b>טלפון:</b> <a href="https://wa.me/972' + String(p.phone).replace(/\D/g, '').replace(/^0/, '') + '">' + p.phone + '</a></p>' +
        '<p><b>אימייל:</b> ' + p.email + '</p>' +
        '<p><b>שבוע הריון:</b> ' + (p.pregnancy_week || '—') + '</p>' +
        (status.note ? '<p style="color:#b00"><b>לתשומת לבך:</b> ' + status.note + '</p>' : '') +
        '<p style="color:#666">התשלום בביט מסומן כ"טרם שולם" — לעדכן בגיליון כשהכסף מגיע.</p>' +
        '</div>'
    });
  } catch (err) {
    console.error('notify failed: ' + err);  // התראה שנכשלה לא מפילה את ההרשמה
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
