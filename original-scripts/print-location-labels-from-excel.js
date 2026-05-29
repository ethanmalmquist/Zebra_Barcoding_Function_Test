const net = require("net");
const path = require("path");
const XLSX = require("xlsx");

/**
 * Zebra ZD421 Location Label Printer - TEST VERSION
 * Purpose:
 * - Keep final location label alignment
 * - Test location-style eeeX orientation with product-label font size
 *
 * Label size: 3 x 1 inch
 * Printer DPI: 203
 * Connection: Ethernet using port 9100
 *
 * Excel columns required:
 * - Aisle
 * - Bay
 * - Level
 * - Bin
 * - Location Barcode
 * - Quantity
 */

const PRINTER_IP = "192.168.1.126";
const PRINTER_PORT = 9100;

const EXCEL_FILE = path.join(__dirname, "location-labels.xlsx");
const SHEET_NAME = "Sheet1";

const PRINT_FIRST_VALID_ROW_ONLY = false;

/**
 * true  = show alignment boxes and center lines
 * false = production print
 */
const PRINT_ALIGNMENT_GUIDES = false;

const LABEL_WIDTH_DOTS = 609;
const LABEL_HEIGHT_DOTS = 203;

/**
 * First 4 sections are the usable field area.
 * Right side remains for eeeX.
 */
const FIELD_AREA_WIDTH = 500;
const FIELD_SECTION_WIDTH = Math.floor(FIELD_AREA_WIDTH / 4);

const SECTION_1_X = 0;
const SECTION_2_X = FIELD_SECTION_WIDTH;
const SECTION_3_X = FIELD_SECTION_WIDTH * 2;
const SECTION_4_X = FIELD_SECTION_WIDTH * 3;

const SECTION_1_CENTER_X = SECTION_1_X + Math.round(FIELD_SECTION_WIDTH / 2);
const SECTION_2_CENTER_X = SECTION_2_X + Math.round(FIELD_SECTION_WIDTH / 2);
const SECTION_3_CENTER_X = SECTION_3_X + Math.round(FIELD_SECTION_WIDTH / 2);
const SECTION_4_CENTER_X = SECTION_4_X + Math.round(FIELD_SECTION_WIDTH / 2);

/**
 * Barcode settings
 * Keep barcode aligned to Level section center.
 */
const BARCODE_Y = 18;
const BARCODE_MODULE_WIDTH = 2;
const BARCODE_HEIGHT = 48;
const BARCODE_X_OFFSET = 0;

/**
 * Vertical layout
 */
const AISLE_VALUE_Y = 74;
const VALUE_Y = 98;
const LABEL_Y = 170;

/**
 * Font settings
 */
const AISLE_FONT_H = 92;
const AISLE_FONT_W = 92;

const VALUE_FONT_H = 64;
const VALUE_FONT_W = 64;

const LABEL_FONT_H = 26;
const LABEL_FONT_W = 26;

/**
 * Field widths
 */
const AISLE_VALUE_FIELD_W = 90;
const VALUE_FIELD_W = 80;
const LABEL_FIELD_W = 90;

/**
 * Guide box heights
 * Kept only so guides can be turned back on if needed.
 */
const VALUE_BOX_H = 58;
const AISLE_VALUE_BOX_H = 72;
const LABEL_BOX_H = 24;

/**
 * Final working visual offsets
 */
const AISLE_VALUE_OFFSET_X = 15;
const AISLE_LABEL_OFFSET_X = 5;

const BAY_VALUE_OFFSET_X = 2;
const BAY_LABEL_OFFSET_X = 2;

const LEVEL_VALUE_OFFSET_X = 2;
const LEVEL_LABEL_OFFSET_X = 3;

const BIN_VALUE_OFFSET_X = 2;
const BIN_LABEL_OFFSET_X = 4;

/**
 * Section guide box settings
 */
const SECTION_GUIDE_TOP_Y = 72;
const SECTION_GUIDE_HEIGHT = 122;
const GUIDE_LINE_THICKNESS = 2;

/**
 * TEST eeeX settings
 *
 * Location label design:
 * - rotated orientation with ^A0B
 * - location label placement
 * - product label font size
 */
const LOGO_X = 545;
const LOGO_Y = 55;
const LOGO_FONT_H = 52;
const LOGO_FONT_W = 52;

const REQUIRED_COLUMNS = [
  "Aisle",
  "Bay",
  "Level",
  "Bin",
  "Location Barcode",
  "Quantity",
];

function fieldX(centerX, fieldWidth, offsetX = 0) {
  return Math.round(centerX - fieldWidth / 2 + offsetX);
}

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

function estimateBarcodeWidthDots(barcodeValue) {
  const value = String(barcodeValue);
  const estimatedModules = 11 + value.length * 11 + 11 + 13;
  return estimatedModules * BARCODE_MODULE_WIDTH;
}

function calculateBarcodeXForSectionCenter(barcodeValue, desiredCenterX) {
  const estimatedWidthDots = estimateBarcodeWidthDots(barcodeValue);

  const calculatedX = Math.round(
    desiredCenterX - estimatedWidthDots / 2 + BARCODE_X_OFFSET
  );

  return Math.max(10, calculatedX);
}

function buildSectionGuides() {
  if (!PRINT_ALIGNMENT_GUIDES) {
    return "";
  }

  return `
^FX Section boxes
^FO${SECTION_1_X},${SECTION_GUIDE_TOP_Y}
^GB${FIELD_SECTION_WIDTH},${SECTION_GUIDE_HEIGHT},${GUIDE_LINE_THICKNESS}^FS

^FO${SECTION_2_X},${SECTION_GUIDE_TOP_Y}
^GB${FIELD_SECTION_WIDTH},${SECTION_GUIDE_HEIGHT},${GUIDE_LINE_THICKNESS}^FS

^FO${SECTION_3_X},${SECTION_GUIDE_TOP_Y}
^GB${FIELD_SECTION_WIDTH},${SECTION_GUIDE_HEIGHT},${GUIDE_LINE_THICKNESS}^FS

^FO${SECTION_4_X},${SECTION_GUIDE_TOP_Y}
^GB${FIELD_SECTION_WIDTH},${SECTION_GUIDE_HEIGHT},${GUIDE_LINE_THICKNESS}^FS

^FX Section center lines
^FO${SECTION_1_CENTER_X},${SECTION_GUIDE_TOP_Y}
^GB${GUIDE_LINE_THICKNESS},${SECTION_GUIDE_HEIGHT},${GUIDE_LINE_THICKNESS}^FS

^FO${SECTION_2_CENTER_X},${SECTION_GUIDE_TOP_Y}
^GB${GUIDE_LINE_THICKNESS},${SECTION_GUIDE_HEIGHT},${GUIDE_LINE_THICKNESS}^FS

^FO${SECTION_3_CENTER_X},${SECTION_GUIDE_TOP_Y}
^GB${GUIDE_LINE_THICKNESS},${SECTION_GUIDE_HEIGHT},${GUIDE_LINE_THICKNESS}^FS

^FO${SECTION_4_CENTER_X},${SECTION_GUIDE_TOP_Y}
^GB${GUIDE_LINE_THICKNESS},${SECTION_GUIDE_HEIGHT},${GUIDE_LINE_THICKNESS}^FS
`;
}

function buildTextFieldGuide(x, y, width, height, centerX) {
  return `
^FO${x},${y}
^GB${width},${height},${GUIDE_LINE_THICKNESS}^FS

^FO${centerX},${y}
^GB${GUIDE_LINE_THICKNESS},${height},${GUIDE_LINE_THICKNESS}^FS
`;
}

function buildTextFieldGuides(coords) {
  if (!PRINT_ALIGNMENT_GUIDES) {
    return "";
  }

  return `
^FX Aisle value field box and center line
${buildTextFieldGuide(
    coords.aisleValueX,
    AISLE_VALUE_Y,
    AISLE_VALUE_FIELD_W,
    AISLE_VALUE_BOX_H,
    SECTION_1_CENTER_X
  )}

^FX Aisle label field box and center line
${buildTextFieldGuide(
    coords.aisleLabelX,
    LABEL_Y,
    LABEL_FIELD_W,
    LABEL_BOX_H,
    SECTION_1_CENTER_X
  )}

^FX Bay value field box and center line
${buildTextFieldGuide(
    coords.bayValueX,
    VALUE_Y,
    VALUE_FIELD_W,
    VALUE_BOX_H,
    SECTION_2_CENTER_X
  )}

^FX Bay label field box and center line
${buildTextFieldGuide(
    coords.bayLabelX,
    LABEL_Y,
    LABEL_FIELD_W,
    LABEL_BOX_H,
    SECTION_2_CENTER_X
  )}

^FX Level value field box and center line
${buildTextFieldGuide(
    coords.levelValueX,
    VALUE_Y,
    VALUE_FIELD_W,
    VALUE_BOX_H,
    SECTION_3_CENTER_X
  )}

^FX Level label field box and center line
${buildTextFieldGuide(
    coords.levelLabelX,
    LABEL_Y,
    LABEL_FIELD_W,
    LABEL_BOX_H,
    SECTION_3_CENTER_X
  )}

^FX Bin value field box and center line
${buildTextFieldGuide(
    coords.binValueX,
    VALUE_Y,
    VALUE_FIELD_W,
    VALUE_BOX_H,
    SECTION_4_CENTER_X
  )}

^FX Bin label field box and center line
${buildTextFieldGuide(
    coords.binLabelX,
    LABEL_Y,
    LABEL_FIELD_W,
    LABEL_BOX_H,
    SECTION_4_CENTER_X
  )}
`;
}

function buildLabel(row) {
  const aisle = sanitizeZplText(row["Aisle"]);
  const bay = sanitizeZplText(row["Bay"]);
  const level = sanitizeZplText(row["Level"]);
  const bin = sanitizeZplText(row["Bin"]);
  const barcodeValue = sanitizeZplText(row["Location Barcode"]);

  const barcodeX = calculateBarcodeXForSectionCenter(
    barcodeValue,
    SECTION_3_CENTER_X
  );

  const coords = {
    aisleValueX: fieldX(
      SECTION_1_CENTER_X,
      AISLE_VALUE_FIELD_W,
      AISLE_VALUE_OFFSET_X
    ),
    aisleLabelX: fieldX(
      SECTION_1_CENTER_X,
      LABEL_FIELD_W,
      AISLE_LABEL_OFFSET_X
    ),

    bayValueX: fieldX(
      SECTION_2_CENTER_X,
      VALUE_FIELD_W,
      BAY_VALUE_OFFSET_X
    ),
    bayLabelX: fieldX(
      SECTION_2_CENTER_X,
      LABEL_FIELD_W,
      BAY_LABEL_OFFSET_X
    ),

    levelValueX: fieldX(
      SECTION_3_CENTER_X,
      VALUE_FIELD_W,
      LEVEL_VALUE_OFFSET_X
    ),
    levelLabelX: fieldX(
      SECTION_3_CENTER_X,
      LABEL_FIELD_W,
      LEVEL_LABEL_OFFSET_X
    ),

    binValueX: fieldX(
      SECTION_4_CENTER_X,
      VALUE_FIELD_W,
      BIN_VALUE_OFFSET_X
    ),
    binLabelX: fieldX(
      SECTION_4_CENTER_X,
      LABEL_FIELD_W,
      BIN_LABEL_OFFSET_X
    ),
  };

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

${buildSectionGuides()}
${buildTextFieldGuides(coords)}

^FX Barcode centered on Level section - no readable text below
^FO${barcodeX},${BARCODE_Y}
^BY${BARCODE_MODULE_WIDTH},2.5,${BARCODE_HEIGHT}
^BCN,${BARCODE_HEIGHT},N,N,N
^FD${barcodeValue}^FS

^FX Aisle value
^FO${coords.aisleValueX},${AISLE_VALUE_Y}
^FB${AISLE_VALUE_FIELD_W},1,0,C,0
^A0N,${AISLE_FONT_H},${AISLE_FONT_W}
^FD${aisle}^FS

^FX Bay value
^FO${coords.bayValueX},${VALUE_Y}
^FB${VALUE_FIELD_W},1,0,C,0
^A0N,${VALUE_FONT_H},${VALUE_FONT_W}
^FD${bay}^FS

^FX Level value
^FO${coords.levelValueX},${VALUE_Y}
^FB${VALUE_FIELD_W},1,0,C,0
^A0N,${VALUE_FONT_H},${VALUE_FONT_W}
^FD${level}^FS

^FX Bin value
^FO${coords.binValueX},${VALUE_Y}
^FB${VALUE_FIELD_W},1,0,C,0
^A0N,${VALUE_FONT_H},${VALUE_FONT_W}
^FD${bin}^FS

^FX Aisle label
^FO${coords.aisleLabelX},${LABEL_Y}
^FB${LABEL_FIELD_W},1,0,C,0
^A0N,${LABEL_FONT_H},${LABEL_FONT_W}
^FDAisle^FS

^FX Bay label
^FO${coords.bayLabelX},${LABEL_Y}
^FB${LABEL_FIELD_W},1,0,C,0
^A0N,${LABEL_FONT_H},${LABEL_FONT_W}
^FDBay^FS

^FX Level label
^FO${coords.levelLabelX},${LABEL_Y}
^FB${LABEL_FIELD_W},1,0,C,0
^A0N,${LABEL_FONT_H},${LABEL_FONT_W}
^FDLevel^FS

^FX Bin label
^FO${coords.binLabelX},${LABEL_Y}
^FB${LABEL_FIELD_W},1,0,C,0
^A0N,${LABEL_FONT_H},${LABEL_FONT_W}
^FDBin^FS

^FX TEST eeeX using location design with product font size
^FO${LOGO_X},${LOGO_Y}
^A0B,${LOGO_FONT_H},${LOGO_FONT_W}
^FDeeeX^FS

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
      console.log("Sending location product-logo TEST label data...");

      client.write(zplData, () => {
        console.log("Location product-logo TEST label data sent.");
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
  console.log("Reading location Excel file for product-logo TEST...");
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
    console.log("No valid completed location rows found.");
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
        `Excel Row ${row.__excelRowNumber}: printing ${quantity} location product-logo TEST label(s)`
      );

      return Array.from({ length: quantity }, () => buildLabel(row));
    })
    .join("");

  console.log(`Completed valid rows found: ${validRows.length}`);
  console.log(`Rows skipped for missing data: ${skippedRows.length}`);
  console.log(`Total location product-logo TEST labels to print now: ${totalLabels}`);

  await sendToPrinter(zplBatch);

  console.log("Location product-logo TEST print job complete.");
}

main().catch((error) => {
  console.error("Location product-logo TEST print failed:");
  console.error(error.message);
});