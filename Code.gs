function doPost(e) {
  try {
    const payload = parsePayload(e);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = "Form Responses";
    const sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);

    const headers = [
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
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    return;
  }

  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const matches = headers.every((h, i) => String(existing[i] || "").trim() === h);

  if (!matches) {
    sheet.insertRowBefore(1);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
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