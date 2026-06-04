const SPREADSHEET_ID = "1yutNepv1S21sBHZMu5pjaPKHdQ2hItkLmkaYLTcWD1U";
const SHEET_NAME = "Service Requests";
const SIGNATURE_FOLDER_NAME = "Tele Space Service Signatures";

const HEADERS = [
  "Timestamp",
  "Job No",
  "Store",
  "Customer Name",
  "Contact Number",
  "Email",
  "Device Model",
  "IMEI or Serial",
  "Reported Faults",
  "Passcode",
  "Quoted Price",
  "Deposit",
  "Payment Method",
  "Heavy Damage",
  "Swollen Battery",
  "Water Damage",
  "Frame Bent",
  "Scratch Dent",
  "Repair History",
  "Terms Accepted",
  "Signature",
  "Notes",
  "Pre Issues",
  "Post Issues"
];

function doPost(e) {
  const data = e.parameter || {};
  const requestId = data.requestId || "";

  try {
    requireFields(data, [
      "store",
      "customerName",
      "contactNumber",
      "email",
      "deviceModel",
      "reportedFaults",
      "signature"
    ]);

    if (data.termsAccepted !== "YES") {
      throw new Error("Terms must be accepted.");
    }

    const sheet = getSheet();
    ensureHeaders(sheet);

    const timestamp = new Date();
    const jobNo = makeUniqueJobNo(sheet, timestamp);
    const signatureUrl = saveSignature(data.signature, jobNo + "-signature");

    sheet.appendRow([
      timestamp,
      jobNo,
      data.store || "",
      data.customerName || "",
      data.contactNumber || "",
      data.email || "",
      data.deviceModel || "",
      data.imei || "",
      data.reportedFaults || "",
      data.passcode || "",
      data.quotedPrice || "",
      data.deposit || "",
      data.paymentMethod || "",
      yesNo(data.heavyDamage),
      yesNo(data.swollenBattery),
      yesNo(data.waterDamage),
      yesNo(data.frameBent),
      yesNo(data.scratchDent),
      yesNo(data.repairHistory),
      "YES",
      signatureUrl,
      data.notes || "",
      data.preIssues || "",
      data.postIssues || ""
    ]);

    sendConfirmationEmail(data, jobNo);

    return createResponse({
      ok: true,
      requestId: requestId,
      jobNo: jobNo
    });
  } catch (error) {
    console.error(error);
    return createResponse({
      ok: false,
      requestId: requestId,
      error: error.message || String(error)
    });
  }
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders(sheet) {
  if (sheet.getMaxColumns() < HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  }

  const current = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];

  if (current.every(function (value) { return !value; })) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  HEADERS.forEach(function (header, index) {
    if (current[index] !== header) {
      throw new Error("Sheet headers do not match. Expected column " + (index + 1) + " to be: " + header);
    }
  });
}

function saveSignature(dataUrl, fileName) {
  const match = String(dataUrl).match(/^data:(image\/(?:png|jpeg));base64,(.+)$/);
  if (!match) throw new Error("Invalid signature image.");

  const extension = match[1] === "image/jpeg" ? "jpg" : "png";
  const blob = Utilities.newBlob(
    Utilities.base64Decode(match[2]),
    match[1],
    fileName + "." + extension
  );

  return getSignatureFolder().createFile(blob).getUrl();
}

function getSignatureFolder() {
  const folders = DriveApp.getFoldersByName(SIGNATURE_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(SIGNATURE_FOLDER_NAME);
}

function makeUniqueJobNo(sheet, timestamp) {
  const base = "TS-" + Utilities.formatDate(timestamp, "Australia/Perth", "yyyyMMdd-HHmmss");
  const existing = sheet.getDataRange().getValues().map(function (row) {
    return String(row[1]);
  });
  let jobNo = base;
  let suffix = 2;

  while (existing.indexOf(jobNo) !== -1) {
    jobNo = base + "-" + suffix++;
  }

  return jobNo;
}

function sendConfirmationEmail(data, jobNo) {
  if (!data.email) return;

  const issueLines = [];
  if (data.preIssues) issueLines.push("Pre Repair Issues: " + data.preIssues);
  if (data.postIssues) issueLines.push("Post Repair Issues: " + data.postIssues);
  if (!issueLines.length) issueLines.push("No function issues found.");

  const body = [
    "Hi " + (data.customerName || "Customer") + ",",
    "",
    "Thank you for choosing Tele Space.",
    "Your service request has been received.",
    "",
    "Job No: " + jobNo,
    "Store: " + (data.store || ""),
    "Device: " + (data.deviceModel || ""),
    "IMEI / Serial: " + (data.imei || ""),
    "Reported Faults: " + (data.reportedFaults || ""),
    "Quoted Price: " + (data.quotedPrice || ""),
    "Deposit: " + (data.deposit || ""),
    "Payment Method: " + (data.paymentMethod || ""),
    "",
    "Function Issues:",
    issueLines.join("\n"),
    "Notes: " + (data.notes || "None"),
    "",
    "Please keep this email for your record.",
    "",
    "Kind regards,",
    "Tele Space"
  ].join("\n");

  MailApp.sendEmail(data.email, "Tele Space Service Request - " + jobNo, body);
}

function requireFields(data, fields) {
  fields.forEach(function (field) {
    if (!String(data[field] || "").trim()) {
      throw new Error("Missing required field: " + field);
    }
  });
}

function yesNo(value) {
  return value === "YES" ? "YES" : "NO";
}

function createResponse(result) {
  const json = JSON.stringify(result).replace(/</g, "\\u003c");
  return HtmlService
    .createHtmlOutput("<script>window.top.postMessage(" + json + ', "*");</script>')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function authorizeServices() {
  const sheet = getSheet();
  ensureHeaders(sheet);
  getSignatureFolder();
  MailApp.getRemainingDailyQuota();
}
