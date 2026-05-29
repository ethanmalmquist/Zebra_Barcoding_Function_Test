// server.js
const express = require("express");
const path = require("path");
const fs = require("fs");
const PDFDocument = require("pdfkit");
const XLSX = require("xlsx");
const { printer } = require("./config");
const printerClient = require("./lib/printer");

const labelModules = {
  product: require("./labels/product"),
  location: require("./labels/location"),
  equipment: require("./labels/equipment"),
};

const app = express();
const PORT = Number(process.env.PORT || 3000);

let lastCompletedPrintJob = null;
let lastCompletedPrintSummary = null;
let printJobHistory = [];
const MAX_PRINT_HISTORY_ITEMS = 25;
const FIXED_COMPANY_NAME = "Bulldog Forklifts";
const PRODUCT_BARCODE_LENGTH = 12;

app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));
app.use(express.static(path.join(__dirname, "public")));

function getLabelModule(type) {
  const mod = labelModules[type];

  if (!mod) {
    throw new Error(
      `Unknown label type: ${type}. Use product, location, or equipment.`
    );
  }

  return mod;
}

function getMode(req) {
  return req.body?.mode || req.query?.mode || "test";
}

function getPrinterHost() {
  return printer.host || printer.ip || "192.168.1.126";
}

function getPrinterPort() {
  return printer.port || 9100;
}

function getLabelSize(type) {
  const sizes = {
    product: {
      width: 3,
      height: 1,
    },
    location: {
      width: 3,
      height: 1,
    },
    equipment: {
      width: 3,
      height: 1,
    },
  };

  return sizes[type] || sizes.product;
}

function splitZplLabels(zpl) {
  if (!zpl || typeof zpl !== "string") {
    return [];
  }

  const matches = zpl.match(/\^XA[\s\S]*?\^XZ/g);
  return matches || [];
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function parseBoolean(value, defaultValue = false) {
  if (value === undefined || value === null) {
    return defaultValue;
  }

  if (typeof value === "boolean") {
    return value;
  }

  const text = String(value).toLowerCase().trim();
  return ["true", "1", "yes", "on"].includes(text);
}

function groupUniqueZplLabels(labels) {
  const map = new Map();

  for (let i = 0; i < labels.length; i++) {
    const zpl = labels[i];

    if (!map.has(zpl)) {
      map.set(zpl, {
        zpl,
        firstOriginalIndex: i + 1,
        quantity: 0,
      });
    }

    map.get(zpl).quantity += 1;
  }

  return Array.from(map.values()).map((item, index) => ({
    ...item,
    displayIndex: index + 1,
  }));
}

const excelDataFiles = {
  product: path.join(__dirname, "data", "product-labels.xlsx"),
  location: path.join(__dirname, "data", "location-labels.xlsx"),
  equipment: path.join(__dirname, "data", "equipment-labels.xlsx"),
};

function getExcelDataFile(type) {
  getLabelModule(type);

  const filePath = excelDataFiles[type];

  if (!filePath) {
    throw new Error(
      `Unknown Excel data type: ${type}. Use product, location, or equipment.`
    );
  }

  if (!fs.existsSync(filePath)) {
    throw new Error(`Excel file not found: ${filePath}`);
  }

  return filePath;
}

function readExcelWorkbook(type) {
  const filePath = getExcelDataFile(type);

  const workbook = XLSX.readFile(filePath, {
    cellDates: true,
    cellNF: false,
    cellText: false,
  });

  const sheetName = workbook.SheetNames[0];

  if (!sheetName) {
    throw new Error(`No worksheet found in ${filePath}.`);
  }

  const worksheet = workbook.Sheets[sheetName];

  if (!worksheet) {
    throw new Error(`Worksheet ${sheetName} could not be read.`);
  }

  const rowsAsArrays = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    blankrows: false,
  });

  const headerRow = rowsAsArrays[0] || [];
  const headers = headerRow.map((header) => String(header || "").trimEnd());
  const usableHeaders = headers.filter((header) => header.trim() !== "");

  if (!usableHeaders.length) {
    throw new Error(
      `No column headers found in ${filePath}. The first row must contain headers.`
    );
  }

  return {
    filePath,
    workbook,
    worksheet,
    sheetName,
    headers,
    usableHeaders,
    rowsAsArrays,
  };
}

function createExcelBackup(filePath) {
  const backupDir = path.join(__dirname, "data", "backups");

  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, {
      recursive: true,
    });
  }

  const parsed = path.parse(filePath);
  const timestamp = new Date()
    .toISOString()
    .replace(/[:.]/g, "-")
    .replace("T", "_")
    .replace("Z", "");

  const backupPath = path.join(
    backupDir,
    `${parsed.name}_${timestamp}${parsed.ext}`
  );

  fs.copyFileSync(filePath, backupPath);

  return backupPath;
}

function normalizeHeaderName(header) {
  return String(header || "")
    .trim()
    .toLowerCase()
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ");
}

function getIncomingValueByHeader(rowData, header) {
  if (!rowData || typeof rowData !== "object") {
    return "";
  }

  if (Object.prototype.hasOwnProperty.call(rowData, header)) {
    return rowData[header];
  }

  const normalizedTarget = normalizeHeaderName(header);

  for (const [key, value] of Object.entries(rowData)) {
    if (normalizeHeaderName(key) === normalizedTarget) {
      return value;
    }
  }

  return "";
}

function isNumericExcelColumn(header) {
  const cleanHeader = normalizeHeaderName(header);

  return ["bay", "level", "bin", "quantity", "qty", "count"].includes(
    cleanHeader
  );
}

function shouldKeepExcelColumnAsText(header) {
  const cleanHeader = normalizeHeaderName(header);

  return [
    "aisle",
    "location barcode",
    "barcode",
    "barcode 128",
    "part number",
    "product name",
    "unit number",
    "unit type",
    "company name",
    "part",
    "sku",
    "description",
    "equipment",
    "equipment id",
    "serial",
    "serial number",
  ].includes(cleanHeader);
}

function normalizeExcelInputValue(header, value) {
  if (value === undefined || value === null) {
    return "";
  }

  const trimmedValue = typeof value === "string" ? value.trim() : value;

  if (trimmedValue === "") {
    return "";
  }

  if (shouldKeepExcelColumnAsText(header)) {
    return String(trimmedValue).trim();
  }

  if (isNumericExcelColumn(header)) {
    const numericValue = Number(trimmedValue);

    if (!Number.isFinite(numericValue)) {
      throw new Error(`${header} must be a valid number.`);
    }

    return numericValue;
  }

  return trimmedValue;
}

function findHeader(headers, targetHeader) {
  const normalizedTarget = normalizeHeaderName(targetHeader);

  return headers.find((header) => {
    return normalizeHeaderName(header) === normalizedTarget;
  });
}

function getRequiredExcelHeaders(type, headers) {
  if (type === "location") {
    return ["Aisle", "Bay", "Level", "Bin", "Quantity"].filter(
      (requiredHeader) => {
        return Boolean(findHeader(headers, requiredHeader));
      }
    );
  }

  if (type === "product") {
    return ["Product Name", "Part Number", "Quantity"].filter(
      (requiredHeader) => {
        return Boolean(findHeader(headers, requiredHeader));
      }
    );
  }

  if (type === "equipment") {
    return ["Unit Number", "Unit type", "Quantity"].filter(
      (requiredHeader) => {
        return Boolean(findHeader(headers, requiredHeader));
      }
    );
  }

  return [];
}

function validateRequiredExcelFields(type, headers, rowData) {
  const requiredHeaders = getRequiredExcelHeaders(type, headers);

  for (const requiredHeader of requiredHeaders) {
    const actualHeader = findHeader(headers, requiredHeader);
    const value = getIncomingValueByHeader(rowData, actualHeader);

    if (String(value ?? "").trim() === "") {
      throw new Error(`${actualHeader} is required.`);
    }
  }
}

function padTwoDigits(value) {
  const trimmedValue = String(value ?? "").trim();

  if (trimmedValue === "") {
    return "";
  }

  const numberValue = Number(trimmedValue);

  if (!Number.isFinite(numberValue)) {
    return trimmedValue.toUpperCase();
  }

  return String(numberValue).padStart(2, "0");
}

function buildLocationBarcodeValueFromFields(
  rowData,
  aisleHeader,
  bayHeader,
  levelHeader,
  binHeader
) {
  const aisle = String(getIncomingValueByHeader(rowData, aisleHeader) ?? "")
    .trim()
    .toUpperCase();

  const bay = getIncomingValueByHeader(rowData, bayHeader);
  const level = getIncomingValueByHeader(rowData, levelHeader);
  const bin = getIncomingValueByHeader(rowData, binHeader);

  if (
    !aisle ||
    String(bay ?? "").trim() === "" ||
    String(level ?? "").trim() === "" ||
    String(bin ?? "").trim() === ""
  ) {
    return "";
  }

  return `${aisle}${padTwoDigits(bay)}${padTwoDigits(level)}${padTwoDigits(
    bin
  )}`;
}

function applyLocationBarcodeAutoFill(type, headers, rowData) {
  if (type !== "location") {
    return rowData;
  }

  const aisleHeader = findHeader(headers, "Aisle");
  const bayHeader = findHeader(headers, "Bay");
  const levelHeader = findHeader(headers, "Level");
  const binHeader = findHeader(headers, "Bin");
  const barcodeHeader = findHeader(headers, "Location Barcode");

  if (
    !aisleHeader ||
    !bayHeader ||
    !levelHeader ||
    !binHeader ||
    !barcodeHeader
  ) {
    return rowData;
  }

  const generatedBarcode = buildLocationBarcodeValueFromFields(
    rowData,
    aisleHeader,
    bayHeader,
    levelHeader,
    binHeader
  );

  rowData[barcodeHeader] = generatedBarcode || "";

  return rowData;
}

function getHeaderIndex(headers, targetHeader) {
  const header = findHeader(headers, targetHeader);
  return header ? headers.indexOf(header) : -1;
}

function isValidProductBarcode(value) {
  return /^\d{12}$/.test(String(value ?? "").trim());
}

function createGeneratedBarcodeReserve(type, headers, rowsAsArrays = []) {
  const reserve = new Set();

  if (type !== "product") {
    return reserve;
  }

  const barcodeIndex = getHeaderIndex(headers, "Barcode 128");

  if (barcodeIndex < 0) {
    return reserve;
  }

  rowsAsArrays.slice(1).forEach((rowArray) => {
    const value = String(rowArray[barcodeIndex] ?? "").trim();

    if (value) {
      reserve.add(value);
    }
  });

  return reserve;
}

function createCleanGeneratedBarcodeReserve(type) {
  return type === "product" ? new Set() : null;
}

function generateRandomProductBarcode() {
  let value = "";

  for (let i = 0; i < PRODUCT_BARCODE_LENGTH; i++) {
    value += String(Math.floor(Math.random() * 10));
  }

  return value;
}

function generateUniqueProductBarcode(reservedGeneratedBarcodes) {
  const reserved = reservedGeneratedBarcodes || new Set();

  for (let attempt = 0; attempt < 10000; attempt++) {
    const value = generateRandomProductBarcode();

    if (!reserved.has(value)) {
      reserved.add(value);
      return value;
    }
  }

  throw new Error("Unable to generate a unique product Barcode 128 value.");
}

function applyProductGeneratedValues(type, headers, rowData, options = {}) {
  if (type !== "product") {
    return rowData;
  }

  const barcodeHeader = findHeader(headers, "Barcode 128");
  const companyHeader = findHeader(headers, "Company Name");
  const reservedGeneratedBarcodes =
    options.reservedGeneratedBarcodes || new Set();

  if (companyHeader) {
    rowData[companyHeader] = FIXED_COMPANY_NAME;
  }

  if (!barcodeHeader) {
    return rowData;
  }

  const existingBarcode = String(
    getIncomingValueByHeader(rowData, barcodeHeader) ?? ""
  ).trim();

  if (options.preserveExistingProductBarcode && existingBarcode) {
    reservedGeneratedBarcodes.add(existingBarcode);
    rowData[barcodeHeader] = existingBarcode;
    return rowData;
  }

  if (
    options.preserveExistingGeneratedValues &&
    isValidProductBarcode(existingBarcode)
  ) {
    reservedGeneratedBarcodes.add(existingBarcode);
    rowData[barcodeHeader] = existingBarcode;
    return rowData;
  }

  rowData[barcodeHeader] = generateUniqueProductBarcode(
    reservedGeneratedBarcodes
  );

  return rowData;
}

function applyEquipmentGeneratedValues(type, headers, rowData) {
  if (type !== "equipment") {
    return rowData;
  }

  const unitNumberHeader = findHeader(headers, "Unit Number");
  const barcodeHeader = findHeader(headers, "Barcode 128");
  const companyHeader = findHeader(headers, "Company Name");

  if (companyHeader) {
    rowData[companyHeader] = FIXED_COMPANY_NAME;
  }

  if (unitNumberHeader && barcodeHeader) {
    rowData[barcodeHeader] = String(
      getIncomingValueByHeader(rowData, unitNumberHeader) ?? ""
    ).trim();
  }

  return rowData;
}

function applyGeneratedExcelValues(type, headers, rowData, options = {}) {
  applyLocationBarcodeAutoFill(type, headers, rowData);
  applyProductGeneratedValues(type, headers, rowData, options);
  applyEquipmentGeneratedValues(type, headers, rowData);

  return rowData;
}

function buildExcelRowFromRequest(type, headers, incomingRow, options = {}) {
  const rowData =
    incomingRow && typeof incomingRow === "object" ? { ...incomingRow } : {};

  applyGeneratedExcelValues(type, headers, rowData, options);
  validateRequiredExcelFields(type, headers, rowData);

  const values = headers.map((header) => {
    if (!header || String(header).trim() === "") {
      return "";
    }

    const incomingValue = getIncomingValueByHeader(rowData, header);

    return normalizeExcelInputValue(header, incomingValue);
  });

  const hasAtLeastOneValue = values.some((value) => {
    return String(value || "").trim() !== "";
  });

  if (!hasAtLeastOneValue) {
    throw new Error("Fill out at least one field before adding a row.");
  }

  return values;
}

function writeWorkbookToDisk(workbook, filePath) {
  const outputBuffer = XLSX.write(workbook, {
    type: "buffer",
    bookType: "xlsx",
    cellDates: true,
  });

  fs.writeFileSync(filePath, outputBuffer);

  const stats = fs.statSync(filePath);

  return {
    savedTo: filePath,
    savedToRelative: path.relative(__dirname, filePath),
    bytes: stats.size,
    modifiedAt: stats.mtime.toISOString(),
  };
}

function appendExcelRow(type, incomingRow) {
  const {
    filePath,
    workbook,
    worksheet,
    sheetName,
    headers,
    usableHeaders,
    rowsAsArrays,
  } = readExcelWorkbook(type);

  const reservedGeneratedBarcodes = createGeneratedBarcodeReserve(
    type,
    headers,
    rowsAsArrays
  );

  const values = buildExcelRowFromRequest(type, headers, incomingRow, {
    reservedGeneratedBarcodes,
  });
  const backupPath = createExcelBackup(filePath);

  const nextExcelRowNumber = rowsAsArrays.length + 1;

  XLSX.utils.sheet_add_aoa(worksheet, [values], {
    origin: -1,
  });

  workbook.Sheets[sheetName] = worksheet;

  const saveResult = writeWorkbookToDisk(workbook, filePath);

  const addedRow = {};

  usableHeaders.forEach((header) => {
    const columnIndex = headers.indexOf(header);
    addedRow[header] = values[columnIndex] ?? "";
  });

  return {
    type,
    sheetName,
    excelRowNumber: nextExcelRowNumber,
    columns: usableHeaders,
    row: addedRow,
    savedTo: saveResult.savedTo,
    savedToRelative: saveResult.savedToRelative,
    bytes: saveResult.bytes,
    modifiedAt: saveResult.modifiedAt,
    backupFile: path.relative(__dirname, backupPath),
  };
}

function appendBulkExcelRows(type, incomingRows) {
  if (!Array.isArray(incomingRows)) {
    throw new Error("Bulk upload requires an array of rows.");
  }

  if (!incomingRows.length) {
    throw new Error("No rows were provided for bulk upload.");
  }

  const {
    filePath,
    workbook,
    worksheet,
    sheetName,
    headers,
    usableHeaders,
    rowsAsArrays,
  } = readExcelWorkbook(type);

  const backupPath = createExcelBackup(filePath);

  const cleanedRows = [];
  const addedRows = [];
  const skippedRows = [];
  const reservedGeneratedBarcodes = createGeneratedBarcodeReserve(
    type,
    headers,
    rowsAsArrays
  );

  incomingRows.forEach((incomingRow, index) => {
    const sourceRowNumber = index + 1;

    const hasAnyValue = Object.values(incomingRow || {}).some((value) => {
      return String(value ?? "").trim() !== "";
    });

    if (!hasAnyValue) {
      skippedRows.push({
        sourceRowNumber,
        reason: "Blank pasted row skipped.",
      });
      return;
    }

    try {
      const values = buildExcelRowFromRequest(type, headers, incomingRow, {
        preserveExistingProductBarcode: type === "product",
        reservedGeneratedBarcodes,
      });

      cleanedRows.push(values);

      const addedRow = {};

      usableHeaders.forEach((header) => {
        const columnIndex = headers.indexOf(header);
        addedRow[header] = values[columnIndex] ?? "";
      });

      addedRows.push({
        sourceRowNumber,
        excelRowNumber: rowsAsArrays.length + cleanedRows.length,
        row: addedRow,
      });
    } catch (error) {
      skippedRows.push({
        sourceRowNumber,
        reason: error.message,
        row: incomingRow,
      });
    }
  });

  if (!cleanedRows.length) {
    throw new Error(
      `No valid rows to upload. ${skippedRows.length} row(s) were skipped.`
    );
  }

  XLSX.utils.sheet_add_aoa(worksheet, cleanedRows, {
    origin: -1,
  });

  workbook.Sheets[sheetName] = worksheet;

  const saveResult = writeWorkbookToDisk(workbook, filePath);

  return {
    type,
    sheetName,
    columns: usableHeaders,
    uploadedRowCount: cleanedRows.length,
    skippedRowCount: skippedRows.length,
    addedRows,
    skippedRows,
    savedTo: saveResult.savedTo,
    savedToRelative: saveResult.savedToRelative,
    bytes: saveResult.bytes,
    modifiedAt: saveResult.modifiedAt,
    backupFile: path.relative(__dirname, backupPath),
  };
}

function cleanExcelRows(type) {
  const {
    filePath,
    workbook,
    sheetName,
    headers,
    usableHeaders,
    rowsAsArrays,
  } = readExcelWorkbook(type);

  const backupPath = createExcelBackup(filePath);
  const originalDataRows = rowsAsArrays.slice(1);

  const cleanedRows = [];
  const removedRows = [];
  const errorRows = [];
  const reservedGeneratedBarcodes = createCleanGeneratedBarcodeReserve(type);

  originalDataRows.forEach((rowArray, index) => {
    const excelRowNumber = index + 2;

    const rowObject = {};

    headers.forEach((header, columnIndex) => {
      if (!header || String(header).trim() === "") {
        return;
      }

      rowObject[header] = rowArray[columnIndex] ?? "";
    });

    const hasAnyValue = Object.values(rowObject).some((value) => {
      return String(value ?? "").trim() !== "";
    });

    if (!hasAnyValue) {
      removedRows.push({
        excelRowNumber,
        reason: "Fully blank row removed.",
      });
      return;
    }

    try {
      if (type === "location") {
        const requiredHeaders = getRequiredExcelHeaders(type, headers);

        const missingRequired = requiredHeaders.filter((requiredHeader) => {
          const actualHeader = findHeader(headers, requiredHeader);
          const value = getIncomingValueByHeader(rowObject, actualHeader);

          return String(value ?? "").trim() === "";
        });

        if (missingRequired.length) {
          removedRows.push({
            excelRowNumber,
            reason: `Incomplete location row removed. Missing: ${missingRequired.join(
              ", "
            )}`,
            row: rowObject,
          });
          return;
        }
      }

      const cleanedValues = buildExcelRowFromRequest(type, headers, rowObject, {
        preserveExistingGeneratedValues: true,
        reservedGeneratedBarcodes,
      });
      cleanedRows.push(cleanedValues);
    } catch (error) {
      errorRows.push({
        excelRowNumber,
        reason: error.message,
        row: rowObject,
      });
    }
  });

  if (errorRows.length) {
    throw new Error(
      `Clean failed. ${errorRows.length} row(s) could not be cleaned. First error: row ${errorRows[0].excelRowNumber}: ${errorRows[0].reason}`
    );
  }

  const rebuiltRows = [headers, ...cleanedRows];
  const newWorksheet = XLSX.utils.aoa_to_sheet(rebuiltRows);

  workbook.Sheets[sheetName] = newWorksheet;

  const saveResult = writeWorkbookToDisk(workbook, filePath);

  return {
    type,
    sheetName,
    originalDataRowCount: originalDataRows.length,
    cleanedDataRowCount: cleanedRows.length,
    removedRowCount: removedRows.length,
    removedRows,
    columns: usableHeaders,
    savedTo: saveResult.savedTo,
    savedToRelative: saveResult.savedToRelative,
    bytes: saveResult.bytes,
    modifiedAt: saveResult.modifiedAt,
    backupFile: path.relative(__dirname, backupPath),
  };
}

function clearExcelTable(type) {
  const {
    filePath,
    workbook,
    sheetName,
    headers,
    usableHeaders,
    rowsAsArrays,
  } = readExcelWorkbook(type);

  const backupPath = createExcelBackup(filePath);

  const originalDataRows = rowsAsArrays.slice(1);
  const rebuiltRows = [headers];

  const newWorksheet = XLSX.utils.aoa_to_sheet(rebuiltRows);

  workbook.Sheets[sheetName] = newWorksheet;

  const saveResult = writeWorkbookToDisk(workbook, filePath);

  return {
    type,
    sheetName,
    clearedRowCount: originalDataRows.length,
    remainingDataRowCount: 0,
    columns: usableHeaders,
    savedTo: saveResult.savedTo,
    savedToRelative: saveResult.savedToRelative,
    bytes: saveResult.bytes,
    modifiedAt: saveResult.modifiedAt,
    backupFile: path.relative(__dirname, backupPath),
  };
}

function deleteExcelRows(type, excelRowNumbers) {
  if (!Array.isArray(excelRowNumbers) || !excelRowNumbers.length) {
    throw new Error("Select at least one saved Excel row to delete.");
  }

  const {
    filePath,
    workbook,
    sheetName,
    headers,
    usableHeaders,
    rowsAsArrays,
  } = readExcelWorkbook(type);

  const rowNumberSet = new Set(
    excelRowNumbers
      .map((rowNumber) => Number(rowNumber))
      .filter((rowNumber) => Number.isInteger(rowNumber) && rowNumber >= 2)
  );

  if (!rowNumberSet.size) {
    throw new Error("No valid Excel row numbers were selected for deletion.");
  }

  const backupPath = createExcelBackup(filePath);
  const originalDataRows = rowsAsArrays.slice(1);
  const deletedRows = [];
  const keptRows = [];

  originalDataRows.forEach((rowArray, index) => {
    const excelRowNumber = index + 2;

    if (rowNumberSet.has(excelRowNumber)) {
      const deletedRow = {};

      usableHeaders.forEach((header) => {
        const columnIndex = headers.indexOf(header);
        deletedRow[header] = rowArray[columnIndex] ?? "";
      });

      deletedRows.push({
        excelRowNumber,
        row: deletedRow,
      });
      return;
    }

    keptRows.push(rowArray);
  });

  if (!deletedRows.length) {
    throw new Error("None of the selected rows were found in the Excel sheet.");
  }

  const rebuiltRows = [headers, ...keptRows];
  const newWorksheet = XLSX.utils.aoa_to_sheet(rebuiltRows);

  workbook.Sheets[sheetName] = newWorksheet;

  const saveResult = writeWorkbookToDisk(workbook, filePath);

  return {
    type,
    sheetName,
    requestedRowNumbers: Array.from(rowNumberSet).sort((a, b) => a - b),
    deletedRowCount: deletedRows.length,
    remainingDataRowCount: keptRows.length,
    deletedRows,
    columns: usableHeaders,
    savedTo: saveResult.savedTo,
    savedToRelative: saveResult.savedToRelative,
    bytes: saveResult.bytes,
    modifiedAt: saveResult.modifiedAt,
    backupFile: path.relative(__dirname, backupPath),
  };
}

function buildFeedOneBlankLabelCommand() {
  return "~PH\n";
}

function buildTerminatePrintCommand() {
  return "~JA\n";
}

function buildPausePrintCommand() {
  return "~PP\n";
}

function buildResumePrintCommand() {
  return "~PS\n";
}

async function renderSingleLabelWithLabelary(
  zpl,
  labelWidthInches,
  labelHeightInches
) {
  const imageBuffer = await renderSingleLabelBufferWithLabelary(
    zpl,
    labelWidthInches,
    labelHeightInches
  );

  return `data:image/png;base64,${imageBuffer.toString("base64")}`;
}

async function renderSingleLabelBufferWithLabelary(
  zpl,
  labelWidthInches,
  labelHeightInches
) {
  const dpmm = "8dpmm";
  const labelIndex = 0;

  const url = `http://api.labelary.com/v1/printers/${dpmm}/labels/${labelWidthInches}x${labelHeightInches}/${labelIndex}/`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      Accept: "image/png",
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body: zpl,
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Labelary render failed: ${response.status} ${errorText}`);
  }

  return Buffer.from(await response.arrayBuffer());
}

async function renderUniqueLabelImagesForPdf(labelItems, labelSize) {
  const imageMap = new Map();

  for (let i = 0; i < labelItems.length; i++) {
    const item = labelItems[i];

    if (imageMap.has(item.zpl)) {
      continue;
    }

    if (imageMap.size > 0) {
      await sleep(750);
    }

    const imageBuffer = await renderSingleLabelBufferWithLabelary(
      item.zpl,
      labelSize.width,
      labelSize.height
    );

    imageMap.set(item.zpl, imageBuffer);
  }

  return imageMap;
}

function createLabelsPdfBuffer(labelItems, imageMap, labelSize) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "LETTER",
      margin: 0,
      autoFirstPage: false,
    });

    const chunks = [];

    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);

    const pageWidth = 612;
    const pageHeight = 792;

    const columns = 2;
    const rows = 9;
    const labelsPerPage = columns * rows;

    const labelWidth = labelSize.width * 72;
    const labelHeight = labelSize.height * 72;

    const horizontalGap = (pageWidth - columns * labelWidth) / (columns + 1);
    const verticalGap = (pageHeight - rows * labelHeight) / (rows + 1);

    const borderWidth = 0.75;

    let labelIndex = 0;

    for (const item of labelItems) {
      const imageBuffer = imageMap.get(item.zpl);

      if (!imageBuffer) {
        continue;
      }

      if (labelIndex % labelsPerPage === 0) {
        doc.addPage({
          size: "LETTER",
          margin: 0,
        });
      }

      const indexOnPage = labelIndex % labelsPerPage;
      const row = Math.floor(indexOnPage / columns);
      const col = indexOnPage % columns;

      const x = horizontalGap + col * (labelWidth + horizontalGap);
      const y = verticalGap + row * (labelHeight + verticalGap);

      doc.lineWidth(borderWidth).rect(x, y, labelWidth, labelHeight).stroke();

      doc.image(imageBuffer, x, y, {
        width: labelWidth,
        height: labelHeight,
      });

      labelIndex += 1;
    }

    if (labelIndex === 0) {
      doc.addPage({
        size: "LETTER",
        margin: 0,
      });

      doc.fontSize(14).text("No labels generated.", 72, 72);
    }

    doc.end();
  });
}

function buildStatusSummary() {
  const status = {};

  for (const [type, mod] of Object.entries(labelModules)) {
    const job = mod.buildPrintJob({
      mode: "preview",
    });

    status[type] = job.summary;
  }

  return status;
}


function getDuplicateKeyField(type) {
  if (type === "location") {
    return "Generated Location Barcode";
  }

  if (type === "product" || type === "equipment") {
    return "Barcode 128";
  }

  return "Barcode";
}

function getDuplicateKeyValue(type, mod, row) {
  if (type === "location" && typeof mod.buildLocationBarcode === "function") {
    return mod.buildLocationBarcode(row);
  }

  if (type === "equipment" && typeof mod.buildEquipmentBarcode === "function") {
    return mod.buildEquipmentBarcode(row);
  }

  if (type === "product" && typeof mod.buildProductBarcode === "function") {
    return mod.buildProductBarcode(row);
  }

  if (type === "product" || type === "equipment") {
    return mod.sanitizeZplText(row["Barcode 128"]);
  }

  return "";
}

function findDuplicateRows(type) {
  const mod = getLabelModule(type);
  const rows = mod.readExcelRows();
  const seen = new Map();
  const duplicates = [];

  rows.forEach((row) => {
    if (!mod.hasRequiredData(row)) {
      return;
    }

    const key = getDuplicateKeyValue(type, mod, row);

    if (!key) {
      return;
    }

    if (!seen.has(key)) {
      seen.set(key, []);
    }

    seen.get(key).push(row.__excelRowNumber);
  });

  for (const [value, excelRowNumbers] of seen.entries()) {
    if (excelRowNumbers.length > 1) {
      duplicates.push({
        field: getDuplicateKeyField(type),
        value,
        excelRowNumbers,
        count: excelRowNumbers.length,
      });
    }
  }

  return duplicates;
}

function buildPrintValidation(type, mode) {
  const mod = getLabelModule(type);
  const job = mod.buildPrintJob({ mode });
  const duplicates = findDuplicateRows(type);
  const warnings = [];

  if (duplicates.length) {
    warnings.push(
      `${duplicates.length} duplicate ${getDuplicateKeyField(type)} value(s) found.`
    );
  }

  if (Number(job.summary?.totalLabels || 0) >= 100) {
    warnings.push(
      `Large print job: ${job.summary.totalLabels} labels will be generated.`
    );
  }

  if (Number(job.summary?.skippedRows || 0) > 0) {
    warnings.push(
      `${job.summary.skippedRows} incomplete row(s) will be skipped.`
    );
  }

  return {
    type,
    mode,
    okToPrint: Boolean(job.zpl),
    summary: job.summary,
    duplicates,
    warnings,
  };
}

function addPrintHistoryEntry(entry) {
  const historyEntry = {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    createdAt: new Date().toISOString(),
    ...entry,
  };

  printJobHistory.unshift(historyEntry);
  printJobHistory = printJobHistory.slice(0, MAX_PRINT_HISTORY_ITEMS);

  return historyEntry;
}

app.get("/api/config", (req, res) => {
  res.json({
    printer: {
      ...printer,
      host: getPrinterHost(),
      port: getPrinterPort(),
    },
    labelTypes: Object.keys(labelModules),
  });
});

app.get("/api/status", (req, res) => {
  try {
    const status = buildStatusSummary();

    res.json({
      ok: true,
      printer: {
        ...printer,
        host: getPrinterHost(),
        port: getPrinterPort(),
      },
      lastCompletedPrintSummary,
      status,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.get("/api/validate/:type", (req, res) => {
  try {
    const type = req.params.type;
    const mode = getMode(req);
    const validation = buildPrintValidation(type, mode);

    res.json({
      ok: true,
      ...validation,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.get("/api/history", (req, res) => {
  res.json({
    ok: true,
    maxItems: MAX_PRINT_HISTORY_ITEMS,
    count: printJobHistory.length,
    history: printJobHistory,
  });
});

app.post("/api/sync-excel", async (req, res) => {
  try {
    const status = buildStatusSummary();

    res.json({
      ok: true,
      action: "sync-excel",
      message: "Excel sheets synced successfully.",
      syncedAt: new Date().toISOString(),
      status,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.get("/api/excel/columns/:type", (req, res) => {
  try {
    const type = req.params.type;
    getLabelModule(type);

    const { filePath, sheetName, usableHeaders } = readExcelWorkbook(type);

    res.json({
      ok: true,
      type,
      sheetName,
      columns: usableHeaders,
      filePath,
      relativeFilePath: path.relative(__dirname, filePath),
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.get("/api/excel/rows/:type", (req, res) => {
  try {
    const type = req.params.type;
    getLabelModule(type);

    const { filePath, sheetName, worksheet } = readExcelWorkbook(type);

    const rows = XLSX.utils.sheet_to_json(worksheet, {
      defval: "",
      blankrows: false,
    }).map((row, index) => ({
      __excelRowNumber: index + 2,
      ...row,
    }));

    res.json({
      ok: true,
      type,
      sheetName,
      count: rows.length,
      rows,
      filePath,
      relativeFilePath: path.relative(__dirname, filePath),
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/excel/add-row/:type", (req, res) => {
  try {
    const type = req.params.type;
    getLabelModule(type);

    const incomingRow = req.body?.row || req.body || {};
    const result = appendExcelRow(type, incomingRow);
    const status = buildStatusSummary();

    res.json({
      ok: true,
      action: "add-excel-row",
      message: "Excel row added successfully and saved to the Excel file.",
      addedAt: new Date().toISOString(),
      result,
      status,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/excel/bulk-add/:type", (req, res) => {
  try {
    const type = req.params.type;
    getLabelModule(type);

    const rows = Array.isArray(req.body?.rows) ? req.body.rows : [];

    const result = appendBulkExcelRows(type, rows);
    const status = buildStatusSummary();

    res.json({
      ok: true,
      action: "bulk-add-excel-rows",
      message: "Bulk labels uploaded successfully and saved to the Excel file.",
      uploadedAt: new Date().toISOString(),
      result,
      status,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/excel/clean/:type", (req, res) => {
  try {
    const type = req.params.type;
    getLabelModule(type);

    const result = cleanExcelRows(type);
    const status = buildStatusSummary();

    res.json({
      ok: true,
      action: "clean-excel-rows",
      message: "Excel rows cleaned successfully and saved to the Excel file.",
      cleanedAt: new Date().toISOString(),
      result,
      status,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/excel/clear-table/:type", (req, res) => {
  try {
    const type = req.params.type;
    getLabelModule(type);

    const result = clearExcelTable(type);
    const status = buildStatusSummary();

    res.json({
      ok: true,
      action: "clear-excel-table",
      message: "Excel table cleared successfully. Header row was kept.",
      clearedAt: new Date().toISOString(),
      result,
      status,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/excel/delete-rows/:type", (req, res) => {
  try {
    const type = req.params.type;
    getLabelModule(type);

    const excelRowNumbers = Array.isArray(req.body?.excelRowNumbers)
      ? req.body.excelRowNumbers
      : [];

    const result = deleteExcelRows(type, excelRowNumbers);
    const status = buildStatusSummary();

    res.json({
      ok: true,
      action: "delete-excel-rows",
      message: "Selected Excel rows deleted successfully. Header row was kept.",
      deletedAt: new Date().toISOString(),
      result,
      status,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/printer/test-connection", async (req, res) => {
  try {
    const result = await printerClient.testConnection(req.body || {});

    res.json({
      ok: true,
      result,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/printer/feed", async (req, res) => {
  try {
    const count = Number(req.body?.count || 1);
    const safeCount = Math.max(1, Math.min(count, 10));

    const zpl = Array.from({ length: safeCount }, () =>
      buildFeedOneBlankLabelCommand()
    ).join("");

    const result = await printerClient.sendToPrinter(zpl, req.body || {});

    res.json({
      ok: true,
      action: "feed",
      count: safeCount,
      command: "~PH",
      result,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/printer/terminate", async (req, res) => {
  try {
    const cancelResult = await printerClient.sendToPrinter(
      buildTerminatePrintCommand(),
      req.body || {}
    );

    res.json({
      ok: true,
      action: "terminate",
      message:
        "Cancel command sent. Current label may finish, but queued labels should stop.",
      cancelResult,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/printer/pause", async (req, res) => {
  try {
    const pauseResult = await printerClient.sendToPrinter(
      buildPausePrintCommand(),
      req.body || {}
    );

    res.json({
      ok: true,
      action: "pause",
      message:
        "Pause command sent. The printer may finish the label currently in motion before pausing.",
      pauseResult,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/printer/resume", async (req, res) => {
  try {
    const resumeResult = await printerClient.sendToPrinter(
      buildResumePrintCommand(),
      req.body || {}
    );

    res.json({
      ok: true,
      action: "resume",
      message: "Resume command sent.",
      resumeResult,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/printer/reprint-last", async (req, res) => {
  try {
    if (!lastCompletedPrintJob) {
      return res.status(400).json({
        ok: false,
        error: "No previous completed print job is stored yet.",
      });
    }

    const printResult = await printerClient.sendToPrinter(
      lastCompletedPrintJob,
      req.body || {}
    );

    const historyEntry = addPrintHistoryEntry({
      action: "reprint-last",
      type: lastCompletedPrintSummary?.type || "unknown",
      mode: lastCompletedPrintSummary?.mode || "previous",
      totalLabels: lastCompletedPrintSummary?.summary?.totalLabels || 0,
      printedRows: lastCompletedPrintSummary?.summary?.printedRows || 0,
      skippedRows: lastCompletedPrintSummary?.summary?.skippedRows || 0,
      duplicateCount: lastCompletedPrintSummary?.validation?.duplicates?.length || 0,
      status: "sent",
      printer: {
        host: getPrinterHost(),
        port: getPrinterPort(),
      },
    });

    res.json({
      ok: true,
      action: "reprint-last",
      summary: lastCompletedPrintSummary,
      printResult,
      historyEntry,
      history: printJobHistory,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/printer/calibrate", async (req, res) => {
  try {
    const zpl = printerClient.buildCalibrateCommand();
    const result = await printerClient.sendToPrinter(zpl, req.body || {});

    res.json({
      ok: true,
      result,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.get("/api/preview/:type", (req, res) => {
  try {
    const type = req.params.type;
    const mod = getLabelModule(type);
    const mode = getMode(req);

    const job = mod.buildPrintJob({
      mode,
    });

    res.json({
      ok: true,
      type,
      mode,
      summary: job.summary,
      zpl: job.zpl,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/preview/:type/image", async (req, res) => {
  try {
    const type = req.params.type;
    const mod = getLabelModule(type);
    const mode = getMode(req);
    const uniqueOnly = parseBoolean(req.body?.uniqueOnly, true);

    const job = mod.buildPrintJob({
      mode,
    });

    if (!job.zpl) {
      return res.status(400).json({
        ok: false,
        summary: job.summary,
        error: job.summary?.message || "No ZPL generated.",
      });
    }

    const labelSize = getLabelSize(type);
    const labels = splitZplLabels(job.zpl);

    if (!labels.length) {
      return res.status(400).json({
        ok: false,
        summary: job.summary,
        error: "No individual ^XA ... ^XZ labels found in generated ZPL.",
      });
    }

    const groupedUniqueLabels = groupUniqueZplLabels(labels);

    let labelItems;

    if (uniqueOnly) {
      labelItems = groupedUniqueLabels;
    } else {
      labelItems = labels.map((zpl, index) => ({
        zpl,
        displayIndex: index + 1,
        firstOriginalIndex: index + 1,
        quantity: 1,
      }));
    }

    const requestedLimit =
      req.body?.maxPreviewLabels || req.query?.maxPreviewLabels || "10";

    let maxPreviewLabels = labelItems.length;

    if (requestedLimit !== "all") {
      const parsed = Number(requestedLimit);

      if (Number.isFinite(parsed) && parsed > 0) {
        maxPreviewLabels = parsed;
      }
    }

    const labelsToPreview = labelItems.slice(0, maxPreviewLabels);
    const images = [];

    for (let i = 0; i < labelsToPreview.length; i++) {
      const item = labelsToPreview[i];

      try {
        if (i > 0) {
          await sleep(750);
        }

        const imageDataUrl = await renderSingleLabelWithLabelary(
          item.zpl,
          labelSize.width,
          labelSize.height
        );

        images.push({
          index: item.displayIndex,
          originalIndex: item.firstOriginalIndex,
          quantity: item.quantity,
          ok: true,
          imageDataUrl,
          zpl: item.zpl,
        });
      } catch (error) {
        images.push({
          index: item.displayIndex,
          originalIndex: item.firstOriginalIndex,
          quantity: item.quantity,
          ok: false,
          error: error.message,
          zpl: item.zpl,
        });
      }
    }

    res.json({
      ok: true,
      type,
      mode,
      uniqueOnly,
      labelSize,
      totalLabels: labels.length,
      uniqueLabels: groupedUniqueLabels.length,
      previewedLabels: images.length,
      requestDelayMs: 750,
      images,
      summary: job.summary,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/pdf/:type", async (req, res) => {
  try {
    const type = req.params.type;
    const mod = getLabelModule(type);
    const mode = getMode(req);

    const job = mod.buildPrintJob({
      mode,
    });

    if (!job.zpl) {
      return res.status(400).json({
        ok: false,
        summary: job.summary,
        error: job.summary?.message || "No ZPL generated.",
      });
    }

    const labelSize = getLabelSize(type);
    const labels = splitZplLabels(job.zpl);

    if (!labels.length) {
      return res.status(400).json({
        ok: false,
        summary: job.summary,
        error: "No individual ^XA ... ^XZ labels found in generated ZPL.",
      });
    }

    const labelItems = labels.map((zpl, index) => ({
      zpl,
      displayIndex: index + 1,
      firstOriginalIndex: index + 1,
      quantity: 1,
    }));

    const imageMap = await renderUniqueLabelImagesForPdf(labelItems, labelSize);
    const pdfBuffer = await createLabelsPdfBuffer(
      labelItems,
      imageMap,
      labelSize
    );

    const safeDate = new Date().toISOString().slice(0, 10);
    const filename = `${type}-labels-${mode}-${safeDate}.pdf`;

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.setHeader("Content-Length", pdfBuffer.length);

    res.send(pdfBuffer);
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/print/:type", async (req, res) => {
  try {
    const type = req.params.type;
    const mod = getLabelModule(type);
    const mode = getMode(req);

    const job = mod.buildPrintJob({
      mode,
    });

    if (!job.zpl) {
      return res.status(400).json({
        ok: false,
        summary: job.summary,
        error: job.summary?.message || "No ZPL generated.",
      });
    }

    const printResult = await printerClient.sendToPrinter(
      job.zpl,
      req.body || {}
    );

    const validation = buildPrintValidation(type, mode);
    const historyEntry = addPrintHistoryEntry({
      action: "print",
      type,
      mode,
      totalLabels: job.summary.totalLabels || 0,
      printedRows: job.summary.printedRows || 0,
      skippedRows: job.summary.skippedRows || 0,
      duplicateCount: validation.duplicates.length,
      status: "sent",
      printer: {
        host: getPrinterHost(),
        port: getPrinterPort(),
      },
    });

    lastCompletedPrintJob = job.zpl;
    lastCompletedPrintSummary = {
      type,
      mode,
      printedAt: historyEntry.createdAt,
      summary: job.summary,
      printResult,
      validation,
      historyEntry,
    };

    res.json({
      ok: true,
      type,
      mode,
      summary: job.summary,
      validation,
      printResult,
      historyEntry,
      history: printJobHistory,
      savedAsLastCompletedPrintJob: true,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.post("/api/raw-zpl", async (req, res) => {
  try {
    const zpl = req.body?.zpl;

    if (!zpl) {
      return res.status(400).json({
        ok: false,
        error: "Missing zpl in request body.",
      });
    }

    const printResult = await printerClient.sendToPrinter(zpl, req.body || {});

    lastCompletedPrintJob = zpl;
    lastCompletedPrintSummary = {
      type: "raw-zpl",
      mode: "manual",
      printedAt: new Date().toISOString(),
      summary: {
        zplLength: zpl.length,
      },
      printResult,
    };

    res.json({
      ok: true,
      printResult,
      savedAsLastCompletedPrintJob: true,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      error: error.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Zebra ZD421 label dashboard running at http://localhost:${PORT}`);
  console.log(`Printer target: ${getPrinterHost()}:${getPrinterPort()}`);
  console.log("Excel files used by dashboard:");
  console.log(`Product:   ${excelDataFiles.product}`);
  console.log(`Location:  ${excelDataFiles.location}`);
  console.log(`Equipment: ${excelDataFiles.equipment}`);
});