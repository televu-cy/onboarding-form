function doPost(e) {
  try {
    const payload = parsePayload(e);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Form Responses";
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    const headers = getCanonicalHeaders();

    const lock = LockService.getScriptLock();
    let lockAcquired = false;

    try {
      lock.waitLock(10000);
      lockAcquired = true;

      ensureHeaders(sheet, headers);

      const timestamp = new Date().toISOString();
      const startingId = getNextSubmissionId(sheet);
      const rows = buildRows(payload, timestamp, headers, startingId);

      if (rows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
        sendFormEmailNotification(rows, headers);
      }

      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, message: "Saved to sheet", rows: rows.length }))
        .setMimeType(ContentService.MimeType.JSON);
    } finally {
      if (lockAcquired) {
        lock.releaseLock();
      }
    }
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getCanonicalHeaders() {
  return [
    "ID",
    "Timestamp",
    "Source",
    "Form For",
    "Portal Name",
    "First Name",
    "Last Name",
    "Name",
    "Email",
    "Phone",
    "Role",
    "User Type",
    "Devices",
    "Shipping Address",
    "Additional Info",
    "Required Delivery Date"
  ];
}

function parsePayload(e) {
  const raw = (e && e.postData && e.postData.contents) || "{}";

  if (typeof raw === "object" && raw !== null) {
    return raw;
  }

  try {
    return JSON.parse(raw);
  } catch (error) {
    return {};
  }
}

function getNextSubmissionId(sheet) {
  const lastRow = sheet.getLastRow();
  return lastRow <= 1 ? 1 : lastRow;
}

function getSubmissionDate(timestamp) {
  return new Date(timestamp).toISOString().slice(0, 10);
}

function buildSubmissionId(timestamp, sequence) {
  return `OF-${getSubmissionDate(timestamp)}-${sequence}`;
}

function ensureHeaders(sheet, headers) {
  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.appendRow(headers);
    return;
  }

  const lastColumn = Math.max(headers.length, sheet.getLastColumn());
  const topRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

  if (isHeaderLikeRow(topRow)) {
    normalizeTopHeaderRow(sheet, topRow, headers);
  } else {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  removeDuplicateHeaderRows(sheet, headers);
}

function normalizeTopHeaderRow(sheet, topRow, headers) {
  const existing = topRow.slice(0, headers.length);
  const matches = headers.every(function(header, index) {
    return normalizeCell(existing[index]) === header;
  });

  if (!matches) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  if (topRow.length > headers.length) {
    sheet.getRange(1, headers.length + 1, 1, topRow.length - headers.length).clearContent();
  }
}

function removeDuplicateHeaderRows(sheet, headers) {
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return;
  }

  const lastColumn = Math.max(headers.length, sheet.getLastColumn());
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  const rowsToDelete = [];

  values.forEach(function(row, index) {
    if (isHeaderLikeRow(row)) {
      rowsToDelete.push(index + 2);
    }
  });

  rowsToDelete.reverse().forEach(function(rowNumber) {
    sheet.deleteRow(rowNumber);
  });
}

function isHeaderLikeRow(row) {
  if (!row || row.length < 3) {
    return false;
  }

  const firstCell = normalizeCell(row[0]);
  const secondCell = normalizeCell(row[1]);
  const thirdCell = normalizeCell(row[2]);

  const hasNewHeader = firstCell === "ID"
    && secondCell === "Timestamp"
    && thirdCell === "Source";

  const hasLegacyHeader = firstCell === "Timestamp"
    && secondCell === "Source"
    && thirdCell === "Form For";

  return hasNewHeader || hasLegacyHeader;
}

function normalizeCell(value) {
  return String(value || "").trim();
}

function buildRows(payload, timestamp, headers, startingId) {
  const rows = [];
  const source = payload.source || "";
  let nextId = startingId;

  if (source === "main-form") {
    rows.push(buildRowFromObject(headers, {
      ID: buildSubmissionId(timestamp, nextId++),
      Timestamp: timestamp,
      Source: source,
      "Form For": payload.formFor || "",
      "Portal Name": payload.portalName || "",
      "First Name": payload.firstName || "",
      "Last Name": payload.lastName || "",
      Name: "",
      Email: payload.email || "",
      Phone: payload.phone || "",
      Role: payload.role || "",
      "User Type": payload.userType || "",
      Devices: joinArray(payload.devices),
      "Shipping Address": payload.shippingAddress || "",
      "Additional Info": payload.additionalInfo || "",
      "Required Delivery Date": payload.requiredDeliveryDate || ""
    }, headers));

    return rows;
  }

  if (source === "user-details-form" && Array.isArray(payload.users)) {
    const mainForm = payload.mainForm || {};

    payload.users.forEach(function(user) {
      rows.push(buildRowFromObject(headers, {
        ID: buildSubmissionId(timestamp, nextId++),
        Timestamp: timestamp,
        Source: source,
        "Form For": mainForm.formFor || "",
        "Portal Name": mainForm.portalName || "",
        "First Name": mainForm.firstName || "",
        "Last Name": mainForm.lastName || "",
        Name: user.name || "",
        Email: user.email || mainForm.email || "",
        Phone: user.phone || mainForm.phone || "",
        Role: user.role || "",
        "User Type": user.userType || "",
        Devices: joinArray(user.devices),
        "Shipping Address": user.shippingAddress || "",
        "Additional Info": user.additionalInfo || "",
        "Required Delivery Date": ""
      }, headers));
    });
  }

  return rows;
}

function buildRowFromObject(headers, values) {
  return headers.map(function(header) {
    return values[header] || "";
  });
}

function sendFormEmailNotification(rows, headers) {
  const recipient = "support@televu.ca";
  const sheetUrl = "https://docs.google.com/spreadsheets/d/1E8ABVdsWoXQen3QeISVlWMYAaunBup7kNgP3Rp0WWbk/edit?usp=sharing";
  const subject = `New Form Submission Received - ${rows.length} row${rows.length === 1 ? "" : "s"}`;

  let plainBody = `A new Onboarding form submission has been received with ${rows.length} row${rows.length === 1 ? "" : "s"}:\n\n`;
  let htmlBody = `<h3>New Onboarding Form Submission Received</h3>`;

  rows.forEach(function(row, rowIndex) {
    const rowLabel = `Row ${rowIndex + 1}`;
    plainBody += `${rowLabel}:\n`;
    htmlBody += `<h4>${rowLabel}</h4>`;
    htmlBody += `<table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif;">`;

    headers.forEach(function(header, index) {
      const value = row[index] || "N/A";
      plainBody += `${header}: ${value}\n`;
      htmlBody += `<tr><td style="background-color: #f2f2f2; font-weight: bold;">${header}</td><td>${value}</td></tr>`;
    });

    plainBody += `\n`;
    htmlBody += `</table><br/>`;
  });

  plainBody += `----------------------------------------\n`;
  plainBody += `Spreadsheet Link:\n${sheetUrl}`;
  htmlBody += `<p><strong>Spreadsheet Link:</strong><br><a href="${sheetUrl}">${sheetUrl}</a></p>`;

  MailApp.sendEmail({
    to: recipient,
    subject: subject,
    body: plainBody,
    htmlBody: htmlBody
  });
}

function joinArray(value) {
  return Array.isArray(value) ? value.join(", ") : "";
}
