const net = require("net");
const path = require("path");
const XLSX = require("xlsx");

/**
 * Zebra ZD421 Product Label Printer
 * Label size: 3 x 1 inch
 * Printer DPI: 203
 * Connection: Ethernet using port 9100
 *
 * Excel columns required:
 * - Product Name
 * - Part Number
 * - Barcode 128
 * - Company Name
 * - Quantity
 */

const PRINTER_IP = "192.168.1.126";
const PRINTER_PORT = 9100;

const EXCEL_FILE = path.join(__dirname, "product-labels.xlsx");
const SHEET_NAME = "Sheet1";

// true = only print first valid completed row
// false = print every valid completed row
const PRINT_FIRST_VALID_ROW_ONLY = false;

const LABEL_WIDTH_DOTS = 609;
const LABEL_HEIGHT_DOTS = 203;

/**
 * Barcode layout
 */
const BARCODE_Y = 82;
const BARCODE_MODULE_WIDTH = 2;
const BARCODE_HEIGHT = 58;

// Negative = move barcode left
// Positive = move barcode right
const BARCODE_X_OFFSET = 0;

/**
 * Text layout
 */
const PRODUCT_NAME_Y = 16;
const PART_NUMBER_Y = 48;
const COMPANY_Y = 166;

/**
 * Font settings
 *
 * Using Zebra built-in Font 0:
 * ^A0N = Font 0, normal orientation
 *
 * Material Name and Part Number were reduced again.
 * Forced letter spacing was removed.
 */
const PRODUCT_NAME_FONT_H = 24;
const PRODUCT_NAME_FONT_W = 24;

const PART_NUMBER_FONT_H = 24;
const PART_NUMBER_FONT_W = 24;

const COMPANY_FONT_H = 30;
const COMPANY_FONT_W = 30;

/**
 * eeeX logo placeholder
 */
const LOGO_EEE_X = 500;
const LOGO_EEE_Y = 160;
const LOGO_EEE_FONT_H = 52;
const LOGO_EEE_FONT_W = 52;

const LOGO_X_X = 575;
const LOGO_X_Y = 160;
const LOGO_X_FONT_H = 52;
const LOGO_X_FONT_W = 52;

const REQUIRED_COLUMNS = [
  "Product Name",
  "Part Number",
  "Barcode 128",
  "Company Name",
  "Quantity",
];

function readExcelRows() {
  const workbook = XLSX.readFile(EXCEL_FILE);
  const worksheet = workbook.Sheets[SHEET_NAME];

  if (!worksheet) {
    throw new Error(`Sheet not found: ${SHEET_NAME}`);
  }

  const rawRows = XLSX.utils.sheet_to_json(worksheet, {
    defval: "",
  });

  return rawRows.map((row, index) => {
    const cleanedRow = {
      __excelRowNumber: index + 2,
    };

    for (const key of Object.keys(row)) {
      cleanedRow[key.trim()] = row[key];
    }

    return cleanedRow;
  });
}

function sanitizeZplText(value) {
  return String(value ?? "")
    .replace(/\^/g, "")
    .replace(/~/g, "")
    .trim();
}

function parseQuantity(value) {
  const quantity = Number(String(value ?? "").trim());

  if (!Number.isInteger(quantity) || quantity < 1) {
    return 0;
  }

  return quantity;
}

function hasRequiredData(row) {
  return REQUIRED_COLUMNS.every((columnName) => {
    if (columnName === "Quantity") {
      return parseQuantity(row[columnName]) > 0;
    }

    return sanitizeZplText(row[columnName]) !== "";
  });
}

function getMissingColumns(row) {
  return REQUIRED_COLUMNS.filter((columnName) => {
    if (columnName === "Quantity") {
      return parseQuantity(row[columnName]) < 1;
    }

    return sanitizeZplText(row[columnName]) === "";
  });
}

function calculateCenteredBarcodeX(barcodeValue) {
  const value = String(barcodeValue);

  const estimatedModules = 11 + value.length * 11 + 11 + 13;
  const estimatedWidthDots = estimatedModules * BARCODE_MODULE_WIDTH;

  const calculatedX = Math.round((LABEL_WIDTH_DOTS - estimatedWidthDots) / 2);
  const adjustedX = calculatedX + BARCODE_X_OFFSET;

  return Math.max(20, adjustedX);
}

function buildLabel(row) {
  const productName = sanitizeZplText(row["Product Name"]).toUpperCase();
  const partNumber = sanitizeZplText(row["Part Number"]);
  const barcodeValue = sanitizeZplText(row["Barcode 128"]);
  const companyName = sanitizeZplText(row["Company Name"]);

  const barcodeX = calculateCenteredBarcodeX(barcodeValue);

  return `
^XA
^CI28
^PW${LABEL_WIDTH_DOTS}
^LL${LABEL_HEIGHT_DOTS}
^LH0,0
^PON
^MTD
^MNY
^PR1
^MD10

^FX Product name centered top - normal letter spacing
^FO0,${PRODUCT_NAME_Y}
^FB609,1,0,C,0
^A0N,${PRODUCT_NAME_FONT_H},${PRODUCT_NAME_FONT_W}
^FD${productName}^FS

^FX Part number centered - normal letter spacing
^FO0,${PART_NUMBER_Y}
^FB609,1,0,C,0
^A0N,${PART_NUMBER_FONT_H},${PART_NUMBER_FONT_W}
^FDPart No: ${partNumber}^FS

^FX Code 128 barcode centered
^FO${barcodeX},${BARCODE_Y}
^BY${BARCODE_MODULE_WIDTH},2.5,${BARCODE_HEIGHT}
^BCN,${BARCODE_HEIGHT},N,N,N
^FD${barcodeValue}^FS

^FX Company name bottom left
^FO23,${COMPANY_Y}
^A0N,${COMPANY_FONT_H},${COMPANY_FONT_W}
^FD${companyName}^FS

^FX eee portion of eeeX placeholder
^FO${LOGO_EEE_X},${LOGO_EEE_Y}
^A0N,${LOGO_EEE_FONT_H},${LOGO_EEE_FONT_W}
^FDeee^FS

^FX X portion of eeeX placeholder
^FO${LOGO_X_X},${LOGO_X_Y}
^A0N,${LOGO_X_FONT_H},${LOGO_X_FONT_W}
^FDX^FS

^XZ
`;
}

function sendToPrinter(zplData) {
  return new Promise((resolve, reject) => {
    console.log(`Connecting to Zebra printer at ${PRINTER_IP}:${PRINTER_PORT}`);

    const client = new net.Socket();

    client.setTimeout(10000);

    client.connect(PRINTER_PORT, PRINTER_IP, () => {
      console.log("Connected to Zebra printer.");
      console.log("Sending product label data...");

      client.write(zplData, () => {
        console.log("Product label data sent.");
        client.end();
      });
    });

    client.on("timeout", () => {
      client.destroy();
      reject(new Error("Connection timed out after 10 seconds."));
    });

    client.on("close", () => {
      console.log("Connection closed.");
      resolve();
    });

    client.on("error", (error) => {
      reject(error);
    });
  });
}

async function main() {
  console.log("Reading product Excel file...");
  console.log(EXCEL_FILE);

  const rows = readExcelRows();

  const validRows = [];
  const skippedRows = [];

  for (const row of rows) {
    if (hasRequiredData(row)) {
      validRows.push(row);
    } else {
      skippedRows.push({
        excelRowNumber: row.__excelRowNumber,
        missingColumns: getMissingColumns(row),
      });
    }
  }

  if (skippedRows.length > 0) {
    console.log("");
    console.log("Skipped rows with missing data:");

    for (const skippedRow of skippedRows) {
      console.log(
        `Excel Row ${skippedRow.excelRowNumber} skipped. Missing or invalid: ${skippedRow.missingColumns.join(", ")}`
      );
    }

    console.log("");
  }

  if (validRows.length === 0) {
    console.log("No valid completed product rows found.");
    return;
  }

  const rowsToPrint = PRINT_FIRST_VALID_ROW_ONLY
    ? validRows.slice(0, 1)
    : validRows;

  let totalLabels = 0;

  const zplBatch = rowsToPrint
    .flatMap((row) => {
      const quantity = parseQuantity(row["Quantity"]);
      totalLabels += quantity;

      console.log(
        `Excel Row ${row.__excelRowNumber}: printing ${quantity} product label(s)`
      );

      return Array.from({ length: quantity }, () => buildLabel(row));
    })
    .join("");

  console.log(`Completed valid rows found: ${validRows.length}`);
  console.log(`Rows skipped for missing data: ${skippedRows.length}`);
  console.log(`Total product labels to print now: ${totalLabels}`);

  await sendToPrinter(zplBatch);

  console.log("Product print job complete.");
}

main().catch((error) => {
  console.error("Product print failed:");
  console.error(error.message);
});