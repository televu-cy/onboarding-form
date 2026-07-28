function doPost(e) {
  try {
    const payload = parsePayload(e);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Form Responses";
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    const headers = getCanonicalHeaders();

    ensureHeaders(sheet, headers);

    const timestamp = new Date().toISOString();
    const rows = buildRows(payload, timestamp, headers);

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, message: "Saved to sheet", rows: rows.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getCanonicalHeaders() {
  return [
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

  return normalizeCell(row[0]) === "Timestamp"
    && normalizeCell(row[1]) === "Source"
    && normalizeCell(row[2]) === "Form For";
}

function normalizeCell(value) {
  return String(value || "").trim();
}

function buildRows(payload, timestamp, headers) {
  const rows = [];
  const source = payload.source || "";

  if (source === "main-form") {
    rows.push(buildRowFromObject(headers, {
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
    payload.users.forEach(function(user) {
      rows.push(buildRowFromObject(headers, {
        Timestamp: timestamp,
        Source: source,
        "Form For": "",
        "Portal Name": "",
        "First Name": "",
        "Last Name": "",
        Name: user.name || "",
        Email: user.email || "",
        Phone: user.phone || "",
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

function joinArray(value) {
  return Array.isArray(value) ? value.join(", ") : "";
}
