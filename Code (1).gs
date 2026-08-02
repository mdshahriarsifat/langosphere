/**
 * LangoSphere — Google Apps Script backend
 * 100% free — no billing account, no credit card, ever.
 */

// ============================= CONFIG =======================================

const COURSES = {
  japanese: {
    name: "Japanese",
    paidFolderId: "REPLACE_WITH_JAPANESE_PAID_FOLDER_ID",
    priceBDT: 10000,
    priceUSD: 95,
    payUrl: "https://langosphere.paymently.io/paymentlink/pay/c0vSfSDrAfiMev9zrp7GW7ELtaIw4WtLiu8hffV4",
    paidColumn: "Japanese_Paid",
  },
  chinese: {
    name: "Chinese",
    paidFolderId: "REPLACE_WITH_CHINESE_PAID_FOLDER_ID",
    priceBDT: 10000,
    priceUSD: 95,
    payUrl: "https://langosphere.paymently.io/paymentlink/pay/QgoE8rn2GCLoAja0UBkDwDusPJcwV1KQpgVExtk6",
    paidColumn: "Chinese_Paid",
  },
};

// =============================================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

function getSheet_(name) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) throw new Error(`Sheet tab "${name}" not found. See README setup steps.`);
  return sheet;
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("LangoSphere Academy")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

// ---------------------------------------------------------------------------
// OTP flow — shared for register & login
// ---------------------------------------------------------------------------
function sendOtp(email) {
  email = String(email || "").trim().toLowerCase();
  if (!email.includes("@")) throw new Error("Enter a valid email address.");

  const otpSheet = getSheet_("OTPs");
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const expiry = new Date(Date.now() + 10 * 60 * 1000).toISOString();

  const data = otpSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).toLowerCase() === email) otpSheet.deleteRow(i + 1);
  }
  otpSheet.appendRow([email, code, expiry]);

  MailApp.sendEmail({
    to: email,
    subject: "Your LangoSphere verification code",
    body: `Your verification code is ${code}. It expires in 10 minutes.\n\nIf you didn't request this, you can ignore this email.\n\n— LangoSphere Academy`,
  });

  return { ok: true };
}

function verifyOtp(email, code, name) {
  email = String(email || "").trim().toLowerCase();
  code = String(code || "").trim();

  const otpSheet = getSheet_("OTPs");
  const otpData = otpSheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < otpData.length; i++) {
    if (String(otpData[i][0]).toLowerCase() === email) { rowIndex = i; break; }
  }
  if (rowIndex === -1) throw new Error("No code was requested for this email. Request a new one.");

  const [, storedCode, expiryIso] = otpData[rowIndex];
  if (new Date() > new Date(expiryIso)) throw new Error("Code expired. Request a new one.");
  if (String(storedCode) !== code) throw new Error("Incorrect code.");

  otpSheet.deleteRow(rowIndex + 1);

  const studentsSheet = getSheet_("Students");
  const headers = studentsSheet.getRange(1, 1, 1, studentsSheet.getLastColumn()).getValues()[0];
  const emailCol = headers.indexOf("Email");
  const data = studentsSheet.getDataRange().getValues();

  let studentRow = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][emailCol]).toLowerCase() === email) { studentRow = i; break; }
  }

  const token = Utilities.getUuid();
  const tokenExpiry = new Date(Date.now() + 30 * 24 * 60 * 60 * 1000).toISOString();
  const tokenCol = headers.indexOf("Token");
  const tokenExpiryCol = headers.indexOf("TokenExpiry");

  if (studentRow === -1) {
    const newRow = new Array(headers.length).fill("");
    newRow[headers.indexOf("Timestamp")] = new Date();
    newRow[headers.indexOf("Name")] = name || "";
    newRow[emailCol] = email;
    newRow[headers.indexOf("Verified")] = true;
    newRow[tokenCol] = token;
    newRow[tokenExpiryCol] = tokenExpiry;
    studentsSheet.appendRow(newRow);
  } else {
    studentsSheet.getRange(studentRow + 1, tokenCol + 1).setValue(token);
    studentsSheet.getRange(studentRow + 1, tokenExpiryCol + 1).setValue(tokenExpiry);
  }

  return { token };
}

// ---------------------------------------------------------------------------
// Dashboard data
// ---------------------------------------------------------------------------
function getDashboardData(token) {
  const studentsSheet = getSheet_("Students");
  const headers = studentsSheet.getRange(1, 1, 1, studentsSheet.getLastColumn()).getValues()[0];
  const data = studentsSheet.getDataRange().getValues();
  const tokenCol = headers.indexOf("Token");
  const tokenExpiryCol = headers.indexOf("TokenExpiry");

  let row = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][tokenCol] === token) { row = data[i]; break; }
  }
  if (!row) throw new Error("Session expired or invalid. Please log in again.");
  if (new Date() > new Date(row[tokenExpiryCol])) throw new Error("Session expired. Please log in again.");

  const name = row[headers.indexOf("Name")];
  const courses = Object.entries(COURSES).map(([id, course]) => {
    const paidColIndex = headers.indexOf(course.paidColumn);
    const owned = paidColIndex !== -1 && row[paidColIndex] === true;
    return {
      id,
      name: course.name,
      priceBDT: course.priceBDT,
      priceUSD: course.priceUSD,
      payUrl: course.payUrl,
      owned,
    };
  });

  return { name, courses };
}

// ---------------------------------------------------------------------------
// Auto-grant Drive access when admin ticks a *_Paid checkbox
// SETUP: attach as installable onEdit trigger
// ---------------------------------------------------------------------------
function handleEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== "Students") return;
  if (e.value !== "TRUE") return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const editedCol = e.range.getColumn();
  const editedHeader = headers[editedCol - 1];

  const course = Object.values(COURSES).find(c => c.paidColumn === editedHeader);
  if (!course) return;

  const emailCol = headers.indexOf("Email") + 1;
  const email = sheet.getRange(e.range.getRow(), emailCol).getValue();
  if (!email) return;

  try {
    DriveApp.getFolderById(course.paidFolderId).addViewer(email);
  } catch (err) {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "LangoSphere: Drive access grant FAILED",
      body: `Could not grant ${email} access to ${course.name}'s paid folder.\nError: ${err.message}`,
    });
    return;
  }

  MailApp.sendEmail({
    to: email,
    subject: `You now have access to ${course.name} — LangoSphere Academy`,
    body: `Thanks for your payment! 🎉\n\nYou now have lifetime access to the full ${course.name} flagship course.\n\nIf you don't see it in Google Drive under "Shared with me", check the email notification Drive just sent you.\n\nHappy learning!\n— LangoSphere Academy`,
  });
}
