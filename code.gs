/**
 * SRI SRINIVASA KALYANAM - Registration & Attendance System
 * -----------------------------------------------------------
 * ARCHITECTURE (v2):
 *  - Register.html  -> served INSIDE Apps Script (no camera needed, google.script.run works fine)
 *  - Dashboard.html -> served INSIDE Apps Script (no camera needed)
 *  - Scanner.html   -> hosted EXTERNALLY on GitHub Pages (needs camera, so it must be a normal
 *                      top-level HTTPS page, NOT inside Apps Script's sandboxed iframe).
 *                      It talks to this script as a plain JSON API using fetch().
 *
 * SETUP:
 * 1. Create a new Google Sheet. Open it -> Extensions -> Apps Script.
 * 2. Paste this file as Code.gs. Add Register.html and Dashboard.html (provided separately)
 *    as HTML files in the same Apps Script project. (Scanner.html is NOT added here - it goes
 *    to GitHub Pages, see the separate instructions.)
 * 3. Run `setup()` once from the Apps Script editor (authorize permissions when asked).
 * 4. Deploy -> New deployment -> Web app.
 *      Execute as: Me
 *      Who has access: Anyone
 *    Copy the deployed /exec URL.
 * 5. That URL is your REGISTRATION page (default).
 *    That URL + "?page=dashboard" is your Dashboard page.
 *    That URL itself (with ?action=... params) is the JSON API the external Scanner page calls.
 * 6. Put the same /exec URL into Scanner.html's SCRIPT_URL constant before publishing to GitHub Pages.
 */

// ====================== CONFIG ======================
const CONFIG = {
  EVENT_NAME: "Sri Srinivasa Kalyanam",
  EVENT_DATE: "Saturday, 11 July 2026",
  EVENT_TIME: "11:00 AM",
  VENUE: "Prema Ravi Auditorium, TCE Campus, Karumathampatti",
  MAPS_LINK: "https://maps.app.goo.gl/hHtmKYiv4pw1DKvR8?g_st=ic",
  REG_PREFIX: "SK2026-",       // Registration No prefix
  CODE_PREFIX: "SK-",          // Entry pass code prefix
  ORG_NAME: "PARK INSTITUTIONS",
  SENDER_NAME: "Park Institutions - Sri Srinivasa Kalyanam",
  SHEET_NAME: "Registrations"
};

// Column indices (1-based) in the "Registrations" sheet.
const COL = {
  TIMESTAMP: 1,
  NAME: 2,
  MOBILE: 3,
  EMAIL: 4,
  ADDRESS: 5,
  ORG: 6,
  DESIGNATION: 7,
  FAMILY_COUNT: 8,      // declared at registration time
  FAMILY_NAMES: 9,
  CITY: 10,
  SPECIAL_ASSIST: 11,
  REG_NO: 12,
  ENTRY_CODE: 13,
  QR_URL: 14,
  ATTENDANCE_STATUS: 15,
  CHECKIN_TIME: 16,
  ACTUAL_COUNT: 17      // <-- filled in BY THE HOST at check-in (actual people who showed up)
};

const HEADERS = [
  "Timestamp", "Full Name", "Mobile Number", "Email ID", "Home Address",
  "Working Place / Organization", "Designation", "Number of Family Members Attending",
  "Names of Family Members", "City / District", "Special Assistance Required",
  "Registration No", "Entry Pass Code", "QR Code URL", "Attendance Status",
  "Check-in Time", "Actual Members Attended"
];

// ====================== ONE-TIME SETUP ======================
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  SpreadsheetApp.getUi().alert("Setup complete! 'Registrations' sheet is ready.");
}

// ====================== WEB APP ROUTING ======================
// Handles THREE kinds of requests:
//  1. Normal page loads from inside Apps Script:      ?page=register (default) | ?page=dashboard
//  2. JSON API calls from the external Scanner page:  ?action=lookup&code=... | ?action=checkin&code=...&count=...
function doGet(e) {
  const params = e.parameter || {};

  // ---- JSON API branch (called by the external GitHub-hosted Scanner page) ----
  if (params.action) {
    return handleApiRequest(params);
  }

  // ---- Normal HTML page branch ----
  const page = (params.page || "register").toLowerCase();
  let template;
  if (page === "dashboard") {
    template = HtmlService.createTemplateFromFile("Dashboard");
  } else {
    template = HtmlService.createTemplateFromFile("Register");
  }
  return template.evaluate()
    .setTitle(CONFIG.EVENT_NAME + " - " + page)
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApiRequest(params) {
  try {
    if (params.action === "lookup") {
      return jsonOut(lookupByCode(params.code || ""));
    }
    if (params.action === "checkin") {
      return jsonOut(checkInByCode(params.code || "", params.count));
    }
    if (params.action === "stats") {
      return jsonOut(getEventStats());
    }
    return jsonOut({ found: false, error: "Unknown action" });
  } catch (err) {
    return jsonOut({ found: false, error: String(err) });
  }
}

// ====================== REGISTRATION (called from Register.html via google.script.run) ======================
function registerMember(form) {
  // form: {name, mobile, email, address, org, designation, familyCount, familyNames, city, specialAssist}
  if (!form.name || !form.mobile || !form.address || !form.org || !form.familyCount || !form.city) {
    throw new Error("Please fill all required fields.");

  }
  // Mobile validation
form.mobile = String(form.mobile).trim();

if (!/^[6-9]\d{9}$/.test(form.mobile)) {
  throw new Error("Invalid mobile number.");
}

// Family count validation
form.familyCount = Number(form.familyCount);

if (!Number.isInteger(form.familyCount) || form.familyCount < 1 || form.familyCount > 25) {
  throw new Error("Invalid family member count.");
}

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const row = sheet.getLastRow() + 1;

  const regNo = CONFIG.REG_PREFIX + String(row - 1).padStart(5, "0");
  const entryCode = generateUniqueEntryCode(sheet);
  const qrUrl = "https://quickchart.io/qr?text=" + encodeURIComponent(entryCode) + "&size=300";

  sheet.getRange(row, COL.TIMESTAMP).setValue(new Date());
  sheet.getRange(row, COL.NAME).setValue(form.name);
  sheet.getRange(row, COL.MOBILE).setValue(form.mobile);
  sheet.getRange(row, COL.EMAIL).setValue(form.email || "");
  sheet.getRange(row, COL.ADDRESS).setValue(form.address);
  sheet.getRange(row, COL.ORG).setValue(form.org);
  sheet.getRange(row, COL.DESIGNATION).setValue(form.designation || "");
  sheet.getRange(row, COL.FAMILY_COUNT).setValue(form.familyCount);
  sheet.getRange(row, COL.FAMILY_NAMES).setValue(form.familyNames || "");
  sheet.getRange(row, COL.CITY).setValue(form.city);
  sheet.getRange(row, COL.SPECIAL_ASSIST).setValue(form.specialAssist || "");
  sheet.getRange(row, COL.REG_NO).setValue(regNo);
  sheet.getRange(row, COL.ENTRY_CODE).setValue(entryCode);
  sheet.getRange(row, COL.QR_URL).setValue(qrUrl);
  sheet.getRange(row, COL.ATTENDANCE_STATUS).setValue("Not Checked In");

  // Confirmation email (optional - only if email given)
  if (form.email) {
    try {
      sendConfirmationEmail(form.email, form.name, regNo, entryCode, qrUrl);
    } catch (err) {
      Logger.log("Email failed: " + err);
    }
  }

  // WhatsApp (optional - only if Twilio configured below)
  try {
    sendWhatsAppMessage(form.mobile, form.name, regNo, entryCode, qrUrl);
  } catch (err) {
    Logger.log("WhatsApp failed: " + err);
  }

  return {
    success: true,
    name: form.name,
    regNo: regNo,
    entryCode: entryCode,
    qrUrl: qrUrl,
    familyCount: form.familyCount
  };
}

function generateUniqueEntryCode(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return CONFIG.CODE_PREFIX + generateRandomCode(6);
  }

  const existing = sheet
    .getRange(2, COL.ENTRY_CODE, lastRow - 1, 1)
    .getValues()
    .flat();

  let code;

  do {
    code = CONFIG.CODE_PREFIX + generateRandomCode(6);
  } while (existing.includes(code));

  return code;
}
function generateRandomCode(length) {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; // no confusing chars like 0/O, 1/I
  let code = "";
  for (let i = 0; i < length; i++) {
    code += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return code;
}

// ====================== WHATSAPP (via Twilio) - OPTIONAL ======================
// 1. Sign up at twilio.com, get a WhatsApp-enabled sender.
// 2. Fill in ACCOUNT_SID / AUTH_TOKEN below. Leave blank to skip WhatsApp sending entirely.
const TWILIO_CONFIG = {
  ACCOUNT_SID: "",   // leave blank to disable WhatsApp sending
  AUTH_TOKEN: "",
  FROM_WHATSAPP: "whatsapp:+14155238886",
  CONTENT_SID: ""
};

function sendWhatsAppMessage(mobile, name, regNo, entryCode, qrUrl) {
  if (!TWILIO_CONFIG.ACCOUNT_SID) return; // disabled

  let toNumber = mobile.toString().replace(/\D/g, "");
  if (toNumber.length === 10) toNumber = "91" + toNumber;
  const to = "whatsapp:+" + toNumber;

  const url = `https://api.twilio.com/2010-04-01/Accounts/${TWILIO_CONFIG.ACCOUNT_SID}/Messages.json`;
  const payload = { From: TWILIO_CONFIG.FROM_WHATSAPP, To: to };

  if (TWILIO_CONFIG.CONTENT_SID) {
    payload.ContentSid = TWILIO_CONFIG.CONTENT_SID;
    payload.ContentVariables = JSON.stringify({ "1": name, "2": regNo, "3": entryCode, "4": qrUrl });
  } else {
    payload.Body = `Namaste ${name}! Your registration for Sri Srinivasa Kalyanam is confirmed.\nReg No: ${regNo}\nEntry Code: ${entryCode}\nQR Code: ${qrUrl}\nDate: ${CONFIG.EVENT_DATE}, ${CONFIG.EVENT_TIME}\nVenue: ${CONFIG.VENUE}`;
  }

  const options = {
    method: "post",
    payload: payload,
    headers: { Authorization: "Basic " + Utilities.base64Encode(TWILIO_CONFIG.ACCOUNT_SID + ":" + TWILIO_CONFIG.AUTH_TOKEN) },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log(response.getContentText());
}

// ====================== EMAIL ======================
function sendConfirmationEmail(email, name, regNo, entryCode, qrUrl) {
  const qrBlob = UrlFetchApp.fetch(qrUrl).getBlob().setName("entry_pass_qr.png");

  const subject = `Registration Confirmed - ${CONFIG.EVENT_NAME}`;
  const htmlBody = `
    <div style="font-family:Georgia,serif;max-width:600px;margin:auto;border:1px solid #eee;padding:24px;">
      <h2 style="color:#b0006e;text-align:center;">🎉 Registration Successful!</h2>
      <p>🙏 Thank you for registering.</p>
      <p>Dear <b>${name}</b>,</p>
      <p>Your registration for <b>${CONFIG.EVENT_NAME}</b> has been successfully received.</p>
      <h3>Registration Details</h3>
      <ul>
        <li><b>Registration No.:</b> ${regNo}</li>
        <li><b>Entry Pass Code:</b> ${entryCode}</li>
        <li><b>Date:</b> ${CONFIG.EVENT_DATE}</li>
        <li><b>Time:</b> ${CONFIG.EVENT_TIME}</li>
        <li><b>Venue:</b> ${CONFIG.VENUE}</li>
      </ul>
      <p><a href="${CONFIG.MAPS_LINK}">📍 View on Google Maps</a></p>
      <p style="text-align:center;">
        <img src="cid:qrimage" style="width:220px;height:220px;" />
        <br/><i>Your Entry Pass QR Code</i>
      </p>
      <p>Kindly present this Entry Pass/QR Code at the registration counter for a smooth entry.</p>
      <p>We look forward to welcoming you and your family.</p>
      <p style="text-align:center;color:#b0006e;"><i>With the divine blessings of Lord Venkateswara & Goddess Padmavathi</i><br/>— ${CONFIG.ORG_NAME}</p>
    </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    name: CONFIG.SENDER_NAME,
    inlineImages: { qrimage: qrBlob }
  });
}

// ====================== ATTENDANCE / SCANNER API ======================

// Read-only lookup - used the instant a QR is scanned, BEFORE the host confirms.
function lookupByCode(scannedCode) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const code = scannedCode.trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL.ENTRY_CODE - 1]).trim() === code) {
      return { found: true, member: rowToMember(data[i]) };
    }
  }
  return { found: false };
}

// Called when the host taps "Confirm Check-in" after (optionally) editing the actual head-count.
function checkInByCode(scannedCode, actualCount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const code = scannedCode.trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][COL.ENTRY_CODE - 1]).trim() === code) {
      const rowNum = i + 1;
      const wasAlreadyIn = data[i][COL.ATTENDANCE_STATUS - 1] === "Checked In";
      const count = (actualCount !== undefined && actualCount !== null && actualCount !== "")
    ? parseInt(actualCount, 10)
    : Number(data[i][COL.FAMILY_COUNT - 1]) || 1;

    if (isNaN(count) || count < 0 || count > 25) {
      throw new Error("Invalid attendance count.");
      }

      sheet.getRange(rowNum, COL.ATTENDANCE_STATUS).setValue("Checked In");
      if (!wasAlreadyIn) {
        sheet.getRange(rowNum, COL.CHECKIN_TIME).setValue(new Date());
      }
      sheet.getRange(rowNum, COL.ACTUAL_COUNT).setValue(count);

      const updatedRow = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
      const member = rowToMember(updatedRow);
      member.alreadyCheckedIn = wasAlreadyIn;
      return { found: true, member: member };
    }
  }
  return { found: false };
}

function rowToMember(r) {
  return {
    name: r[COL.NAME - 1],
    mobile: r[COL.MOBILE - 1],
    regNo: r[COL.REG_NO - 1],
    entryCode: r[COL.ENTRY_CODE - 1],
    familyCount: r[COL.FAMILY_COUNT - 1],
    city: r[COL.CITY - 1],
    status: r[COL.ATTENDANCE_STATUS - 1] || "Not Checked In",
    actualCount: r[COL.ACTUAL_COUNT - 1] || "",
    alreadyCheckedIn: r[COL.ATTENDANCE_STATUS - 1] === "Checked In"
  };
}

// Returns full list for the dashboard (called via google.script.run since Dashboard.html
// stays inside Apps Script).
function getAttendanceList() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    if (!data[i][COL.NAME - 1]) continue;
    list.push({
      regNo: data[i][COL.REG_NO - 1],
      entryCode: data[i][COL.ENTRY_CODE - 1],
      name: data[i][COL.NAME - 1],
      mobile: data[i][COL.MOBILE - 1],
      familyCount: data[i][COL.FAMILY_COUNT - 1],
      actualCount: data[i][COL.ACTUAL_COUNT - 1] || "",
      city: data[i][COL.CITY - 1],
      status: data[i][COL.ATTENDANCE_STATUS - 1] || "Not Checked In",
      checkinTime: data[i][COL.CHECKIN_TIME - 1] ? data[i][COL.CHECKIN_TIME - 1].toString() : ""
    });
  }
  return list;
}
function getEventStats() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
      .getSheetByName(CONFIG.SHEET_NAME);

  const data = sheet.getDataRange().getValues();

  let totalFamilies = 0;
  let totalMembers = 0;
  let checkedInFamilies = 0;
  let attendedMembers = 0;

  for (let i = 1; i < data.length; i++) {

    if (!data[i][COL.NAME - 1]) continue;

    totalFamilies++;

    totalMembers += Number(data[i][COL.FAMILY_COUNT - 1]) || 0;

    if (data[i][COL.ATTENDANCE_STATUS - 1] === "Checked In") {

      checkedInFamilies++;

      attendedMembers += Number(data[i][COL.ACTUAL_COUNT - 1]) || 0;

    }

  }

  return {
    totalFamilies: totalFamilies,
    totalMembers: totalMembers,
    checkedInFamilies: checkedInFamilies,
    attendedMembers: attendedMembers
  };

}
