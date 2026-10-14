/**
 * NBDY WARRANTY TRACKER - GOOGLE SHEETS BACKEND
 * ------------------------------------------------
 * Simpanan tambahan:
 * - customerAddress
 * - platform (Website / TikTok)
 *
 * Selepas mengubah kod ini:
 * Deploy > Manage deployments > Edit > New version > Deploy.
 */

const SHEET_NAME = "Claims";

const HEADERS = [
  "id",
  "customerName",
  "customerPhone",
  "customerAddress",
  "source",
  "platform",
  "orderNumber",
  "itemsJSON",
  "status",
  "returnStatus",
  "trackingReturn",
  "trackingNew",
  "notes",
  "createdAt",
  "updatedAt",
  "reportDate",
  "issueSummary",
  "photosJSON"
];

function getHeaders_(sheet) {
  if (sheet.getLastColumn() === 0) return [];
  return sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0]
    .map(function (header) {
      return String(header || "").trim();
    });
}

function getHeaderIndex_(headers) {
  const index = {};
  headers.forEach(function (header, columnIndex) {
    if (header) index[header] = columnIndex;
  });
  return index;
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // Tambah kolum baharu di hujung tanpa mengganggu data atau susunan lama.
    const existingHeaders = getHeaders_(sheet);
    const missingHeaders = HEADERS.filter(function (header) {
      return existingHeaders.indexOf(header) === -1;
    });

    if (missingHeaders.length > 0) {
      sheet
        .getRange(1, existingHeaders.length + 1, 1, missingHeaders.length)
        .setValues([missingHeaders]);
    }
  }

  // Cari reportDate berdasarkan header sebenar kerana kolum baharu mungkin
  // ditambah di hujung Sheet lama.
  const currentHeaders = getHeaders_(sheet);
  const reportDateColumn = currentHeaders.indexOf("reportDate") + 1;

  if (reportDateColumn > 0) {
    const maxRows = Math.max(sheet.getMaxRows(), 1000);
    sheet.getRange(1, reportDateColumn, maxRows, 1).setNumberFormat("@");
  }

  return sheet;
}

function rowToClaim_(row, headerIndex) {
  const claim = {};

  HEADERS.forEach(function (header) {
    const columnIndex = headerIndex[header];
    claim[header] = columnIndex === undefined ? "" : row[columnIndex];
  });

  // Betulkan rekod lama jika reportDate pernah ditukar kepada Date oleh Sheets.
  if (claim.reportDate instanceof Date) {
    claim.reportDate = Utilities.formatDate(
      claim.reportDate,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );
  }

  try {
    claim.items = claim.itemsJSON ? JSON.parse(claim.itemsJSON) : [];
  } catch (err) {
    claim.items = [];
  }
  delete claim.itemsJSON;

  try {
    claim.photos = claim.photosJSON ? JSON.parse(claim.photosJSON) : [];
  } catch (err) {
    claim.photos = [];
  }
  delete claim.photosJSON;

  return claim;
}

function getAllClaims_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  const headers = values[0].map(function (header) {
    return String(header || "").trim();
  });
  const headerIndex = getHeaderIndex_(headers);
  const idIndex = headerIndex.id;

  if (idIndex === undefined) {
    throw new Error('Header "id" tidak dijumpai.');
  }

  const claims = [];

  for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
    const row = values[rowIndex];
    if (!row[idIndex]) continue;
    claims.push(rowToClaim_(row, headerIndex));
  }

  claims.sort(function (a, b) {
    return new Date(b.createdAt) - new Date(a.createdAt);
  });

  return claims;
}

function findRowById_(sheet, id) {
  const headers = getHeaders_(sheet);
  const idColumnIndex = headers.indexOf("id");

  if (idColumnIndex === -1) {
    throw new Error('Header "id" tidak dijumpai.');
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;

  const result = sheet
    .getRange(2, idColumnIndex + 1, lastRow - 1, 1)
    .createTextFinder(String(id))
    .matchEntireCell(true)
    .findNext();

  return result ? result.getRow() : -1;
}

function claimValues_(claim, existingCreatedAt) {
  const now = new Date().toISOString();

  return {
    id: claim.id || "",
    customerName: claim.customerName || "",
    customerPhone: claim.customerPhone || "",
    customerAddress: claim.customerAddress || "",
    source: claim.source || "",
    platform: claim.platform || "",
    orderNumber: claim.orderNumber || "",
    itemsJSON: JSON.stringify(claim.items || []),
    status: claim.status || "Pending",
    returnStatus: claim.returnStatus || "",
    trackingReturn: claim.trackingReturn || "",
    trackingNew: claim.trackingNew || "",
    notes: claim.notes || "",
    createdAt: claim.createdAt || existingCreatedAt || now,
    updatedAt: now,
    reportDate: claim.reportDate || "",
    issueSummary: claim.issueSummary || "",
    photosJSON: JSON.stringify(claim.photos || [])
  };
}

function valuesForHeaders_(headers, values, existingRow) {
  return headers.map(function (header, columnIndex) {
    if (Object.prototype.hasOwnProperty.call(values, header)) {
      return values[header];
    }

    // Kekalkan kolum tambahan yang mungkin dibuat secara manual.
    return existingRow ? existingRow[columnIndex] : "";
  });
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET: pulangkan semua tuntutan.
function doGet(e) {
  try {
    return jsonOut_({
      success: true,
      claims: getAllClaims_()
    });
  } catch (err) {
    console.error("Warranty API error: " + (err.stack || err.message || err));
    return jsonOut_({
      success: false,
      error: err.message
    });
  }
}

// POST: add / update / delete / toggleStatus / clearAll / uploadPhoto.
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Tiada data diterima.");
    }

    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    // Upload foto tidak mengubah Sheet, jadi tidak perlu mengambil lock.
    if (action === "uploadPhoto") {
      const url = uploadPhotoToDrive_(
        body.dataUrl,
        body.filename || ("photo-" + Date.now() + ".jpg")
      );
      return jsonOut_({ success: true, url: url });
    }

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);

    try {
      const sheet = getSheet_();
      const headers = getHeaders_(sheet);

      if (action === "add") {
        const claim = body.claim || {};
        if (!claim.id) throw new Error("Claim ID diperlukan.");

        const values = claimValues_(claim, "");
        sheet.appendRow(valuesForHeaders_(headers, values, null));
      } else if (action === "update") {
        const claim = body.claim || {};
        const rowNumber = findRowById_(sheet, claim.id);

        if (rowNumber === -1) {
          throw new Error("Claim not found: " + claim.id);
        }

        const existingRow = sheet
          .getRange(rowNumber, 1, 1, headers.length)
          .getValues()[0];
        const createdAtIndex = headers.indexOf("createdAt");
        const existingCreatedAt =
          createdAtIndex === -1 ? "" : existingRow[createdAtIndex];
        const values = claimValues_(claim, existingCreatedAt);

        sheet
          .getRange(rowNumber, 1, 1, headers.length)
          .setValues([
            valuesForHeaders_(headers, values, existingRow)
          ]);
      } else if (action === "toggleStatus") {
        const rowNumber = findRowById_(sheet, body.id);
        if (rowNumber === -1) {
          throw new Error("Claim not found: " + body.id);
        }

        const statusColumn = headers.indexOf("status") + 1;
        const updatedAtColumn = headers.indexOf("updatedAt") + 1;

        if (statusColumn < 1) {
          throw new Error('Header "status" tidak dijumpai.');
        }

        const current = sheet.getRange(rowNumber, statusColumn).getValue();
        const next = current === "Done" ? "Pending" : "Done";
        sheet.getRange(rowNumber, statusColumn).setValue(next);

        if (updatedAtColumn > 0) {
          sheet
            .getRange(rowNumber, updatedAtColumn)
            .setValue(new Date().toISOString());
        }
      } else if (action === "delete") {
        const rowNumber = findRowById_(sheet, body.id);
        if (rowNumber === -1) {
          throw new Error("Claim not found: " + body.id);
        }
        sheet.deleteRow(rowNumber);
      } else if (action === "clearAll") {
        const lastRow = sheet.getLastRow();
        const lastColumn = sheet.getLastColumn();

        if (lastRow > 1 && lastColumn > 0) {
          sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
        }
      } else {
        throw new Error("Unknown action: " + action);
      }

      return jsonOut_({
        success: true,
        claims: getAllClaims_()
      });
    } finally {
      if (lock.hasLock()) {
        lock.releaseLock();
      }
    }
  } catch (err) {
    console.error("Warranty API error: " + (err.stack || err.message || err));
    return jsonOut_({
      success: false,
      error: err.message
    });
  }
}

// Simpan foto dan video ke Google Drive.
function getPhotosFolder_() {
  const FOLDER_NAME = "NBDY Warranty Claim Photos";
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);

  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(FOLDER_NAME);
}

// Jalankan fungsi ini sekali dari editor Apps Script selepas memasang fungsi
// gambar. Ini meminta kebenaran Google Drive dan memastikan folder tersedia.
function authorizePhotoStorage() {
  const folder = getPhotosFolder_();
  return "Photo storage ready: " + folder.getName();
}

function uploadPhotoToDrive_(dataUrl, filename) {
  const matches = String(dataUrl || "").match(
    /^data:((?:image|video)\/[a-zA-Z0-9.+-]+);base64,(.+)$/
  );

  if (!matches) throw new Error("Invalid photo or video data");

  const mimeType = matches[1];
  const allowedVideoTypes = ["video/mp4", "video/quicktime", "video/webm"];

  if (
    mimeType.indexOf("image/") !== 0 &&
    allowedVideoTypes.indexOf(mimeType) === -1
  ) {
    throw new Error("Unsupported media type: " + mimeType);
  }

  const base64Data = matches[2];
  const bytes = Utilities.base64Decode(base64Data);

  if (mimeType.indexOf("video/") === 0 && bytes.length > 10 * 1024 * 1024) {
    throw new Error("Video must not exceed 10 MB");
  }

  const blob = Utilities.newBlob(bytes, mimeType, filename);
  const folder = getPhotosFolder_();
  const file = folder.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return "https://drive.google.com/uc?export=view&id=" + file.getId();
}
