const PAGE = document.body.dataset.page || "home";
const DEFAULT_DATA = window.SALES_DATA;

if (!DEFAULT_DATA || !DEFAULT_DATA.rows || !DEFAULT_DATA.meta) {
  throw new Error("Sales data was not loaded.");
}

const STORAGE_KEYS = {
  dataset: "daily-sales-dashboard.dataset.v5",
  state: "daily-sales-dashboard.state.v5",
  outletMaster: "daily-sales-dashboard.outlet-master.v1",
  deletedOutlets: "daily-sales-dashboard.deleted-outlets.v1",
  historicalMaster: "daily-sales-dashboard.historical-master.v1",
  historicalImport: "daily-sales-dashboard.historical-import.v1",
  areaTargets: "daily-sales-dashboard.area-targets.v1",
  deletedAreaTargets: "daily-sales-dashboard.deleted-area-targets.v1",
};

const KNOWN_PRODUCT_ORDER = ["MS", "HSD", "Power", "Lube", "DEF"];
const HISTORICAL_COMPARISON_PRODUCTS = ["MS", "HSD", "Power"];
const TARGET_PRODUCTS = ["MS", "HSD", "Power", "Lube", "DEF"];
const DEPOT_FILTER_OPTIONS = ["JABALPUR DEPOT", "SAGAR DEPOT"];
const DEPOT_ALIASES = {
  1424: "JABALPUR DEPOT",
  JABALPUR: "JABALPUR DEPOT",
  "JABALPUR DEPOT": "JABALPUR DEPOT",
  1436: "SAGAR DEPOT",
  SAGAR: "SAGAR DEPOT",
  "SAGAR DEPOT": "SAGAR DEPOT",
};
const DEFAULT_PROJECT_CODES =
  DEFAULT_DATA.meta?.projectAbhuyadayCodes ||
  (DEFAULT_DATA.filters?.outlets || [])
    .filter((outlet) => outlet.isProjectAbhuyaday)
    .map((outlet) => String(outlet.sapCode));
const DEFAULT_HISTORICAL_LOOKUP = {
  MS: { ...(DEFAULT_DATA.historicalLookup?.MS || {}) },
  HSD: { ...(DEFAULT_DATA.historicalLookup?.HSD || {}) },
  Power: { ...(DEFAULT_DATA.historicalLookup?.Power || {}) },
};

const formatNumber = new Intl.NumberFormat("en-IN", {
  maximumFractionDigits: 2,
});
const formatWhole = new Intl.NumberFormat("en-IN", {
  maximumFractionDigits: 0,
});
const formatPercent = new Intl.NumberFormat("en-IN", {
  maximumFractionDigits: 1,
});
const formatDateLabel = new Intl.DateTimeFormat("en-GB", {
  day: "2-digit",
  month: "short",
  year: "numeric",
});
const formatDateShort = new Intl.DateTimeFormat("en-GB", {
  day: "2-digit",
  month: "short",
});
const formatMonthLabel = new Intl.DateTimeFormat("en-GB", {
  month: "short",
  year: "numeric",
});

const elements = {
  sourceFile: document.querySelector("#source-file"),
  sourcePeriod: document.querySelector("#source-period"),
  generatedAt: document.querySelector("#generated-at"),
  csvUpload: document.querySelector("#csv-upload"),
  restoreBuiltIn: document.querySelector("#restore-built-in"),
  uploadStatus: document.querySelector("#upload-status"),
  startDate: document.querySelector("#start-date"),
  endDate: document.querySelector("#end-date"),
  salesArea: document.querySelector("#sales-area"),
  plantFilter: document.querySelector("#plant-filter"),
  outletSelect: document.querySelector("#outlet-select"),
  outletSearch: document.querySelector("#outlet-search"),
  paFilter: document.querySelector("#pa-filter"),
  productPills: document.querySelector("#product-pills"),
  activeSummary: document.querySelector("#active-summary"),
  metricsGrid: document.querySelector("#metrics-grid"),
  productCards: document.querySelector("#product-cards"),
  emptyProductsNote: document.querySelector("#empty-products-note"),
  areaBars: document.querySelector("#area-bars"),
  latestDay: document.querySelector("#latest-day"),
  fullPeriod: document.querySelector("#full-period"),
  resetFilters: document.querySelector("#reset-filters"),
  registerBody: document.querySelector("#register-body"),
  monthlyRegisterBody: document.querySelector("#monthly-register-body"),
  historicalBody: document.querySelector("#historical-body"),
  historicalAreaBody: document.querySelector("#historical-area-body"),
  historicalMetrics: document.querySelector("#historical-metrics"),
  historicalCountNote: document.querySelector("#historical-count-note"),
  historicalAreaNote: document.querySelector("#historical-area-note"),
  outletBody: document.querySelector("#outlet-body"),
  nilSaleBody: document.querySelector("#nil-sale-body"),
  nilSaleCount: document.querySelector("#nil-sale-count"),
  exportRegister: document.querySelector("#export-register"),
  exportRegisterExcel: document.querySelector("#export-register-excel"),
  exportMonthlyRegister: document.querySelector("#export-monthly-register"),
  exportMonthlyRegisterExcel: document.querySelector("#export-monthly-register-excel"),
  exportHistoricalExcel: document.querySelector("#export-historical-excel"),
  exportSummary: document.querySelector("#export-summary"),
  exportSummaryExcel: document.querySelector("#export-summary-excel"),
  printView: document.querySelector("#print-view"),
  masterBack: document.querySelector("#master-back"),
  backupMasterData: document.querySelector("#backup-master-data"),
  restoreMasterBackup: document.querySelector("#restore-master-backup"),
  restoreMasterData: document.querySelector("#restore-master-data"),
  masterStatus: document.querySelector("#master-status"),
  masterOutletCount: document.querySelector("#master-outlet-count"),
  masterHistoricalCount: document.querySelector("#master-historical-count"),
  masterTargetCount: document.querySelector("#master-target-count"),
  masterTabButtons: document.querySelectorAll("[data-master-tab]"),
  masterPanels: document.querySelectorAll("[data-master-panel]"),
  historicalImportForm: document.querySelector("#historical-import-form"),
  historicalExcelUpload: document.querySelector("#historical-excel-upload"),
  historicalImportProduct: document.querySelector("#historical-import-product"),
  historicalImportList: document.querySelector("#historical-import-list"),
  outletMasterForm: document.querySelector("#outlet-master-form"),
  outletMasterBody: document.querySelector("#outlet-master-body"),
  cancelOutletEdit: document.querySelector("#cancel-outlet-edit"),
  exportOutletMasterExcel: document.querySelector("#export-outlet-master-excel"),
  historicalMasterForm: document.querySelector("#historical-master-form"),
  historicalMasterBody: document.querySelector("#historical-master-body"),
  historicalOutletSelect: document.querySelector("#historical-outlet-select"),
  targetMasterForm: document.querySelector("#target-master-form"),
  targetMasterBody: document.querySelector("#target-master-body"),
  targetAreaSelect: document.querySelector("#target-area-select"),
  targetMonth: document.querySelector("#target-month"),
  cancelTargetEdit: document.querySelector("#cancel-target-edit"),
  exportTargetMasterExcel: document.querySelector("#export-target-master-excel"),
  targetCharts: document.querySelector("#target-charts"),
  shareDashboardEmail: document.querySelector("#share-dashboard-email"),
  toggleNilSale: document.querySelector("#toggle-nil-sale"),
  nilSaleListWrap: document.querySelector("#nil-sale-list-wrap"),
  toggleMonthlyRegister: document.querySelector("#toggle-monthly-register"),
  monthlyRegisterListWrap: document.querySelector("#monthly-register-list-wrap"),
  toggleOutletSummary: document.querySelector("#toggle-outlet-summary"),
  outletSummaryListWrap: document.querySelector("#outlet-summary-list-wrap"),
  downloadAreaPng: document.querySelector("#download-area-png"),
  downloadAreaJpeg: document.querySelector("#download-area-jpeg"),
  exportAreaExcel: document.querySelector("#export-area-excel"),
  downloadTargetPng: document.querySelector("#download-target-png"),
  downloadTargetJpeg: document.querySelector("#download-target-jpeg"),
  exportTargetExcel: document.querySelector("#export-target-excel"),
};

let builtInData = null;
let baseData = null;
let currentData = null;
let currentState = null;
let lastRegisterRows = [];
let lastMonthlyRegisterRows = [];
let lastHistoricalRows = [];
let lastOutletRows = [];

function parseDate(value) {
  return new Date(`${value}T00:00:00`);
}

function displayDate(value) {
  return formatDateLabel.format(parseDate(value));
}

function displayDateShort(value) {
  return formatDateShort.format(parseDate(value));
}

function displayMonth(value) {
  return formatMonthLabel.format(parseDate(`${value}-01`));
}

function displayMonthKey(value) {
  return value ? displayMonth(value) : "All months";
}

function safeDivide(numerator, denominator) {
  if (!denominator) {
    return 0;
  }
  return numerator / denominator;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function sortByName(items) {
  return [...items].sort((left, right) => {
    const areaCompare = left.salesArea.localeCompare(right.salesArea);
    if (areaCompare !== 0) {
      return areaCompare;
    }
    return left.outletName.localeCompare(right.outletName);
  });
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) {
    return value;
  }
  const text = String(value ?? "")
    .trim()
    .replaceAll(",", "");
  if (!text) {
    return 0;
  }
  const number = Number(text);
  return Number.isFinite(number) ? number : 0;
}

function cleanText(value, fallback = "") {
  const text = String(value ?? "").trim();
  return text || fallback;
}

function plantDescription(value) {
  const text = cleanText(value).toUpperCase().replace(/\s+/g, " ").trim();
  if (!text) {
    return "";
  }
  return DEPOT_ALIASES[text] || text;
}

function collectDepotOptions(rows, outlets) {
  const presentDepots = new Set(
    rows
      .map((row) => plantDescription(row.plant))
      .concat(outlets.map((outlet) => plantDescription(outlet.plant)))
      .filter(Boolean),
  );
  const knownDepots = DEPOT_FILTER_OPTIONS.filter((depot) => presentDepots.has(depot));
  return knownDepots.length ? knownDepots : [...DEPOT_FILTER_OPTIONS];
}

function normalizeCode(value) {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(Math.trunc(value));
  }
  const text = String(value).trim();
  if (!text) {
    return "";
  }
  const numeric = Number(text);
  if (Number.isFinite(numeric) && text.includes(".")) {
    return String(Math.trunc(numeric));
  }
  return text;
}

function slugify(value) {
  return cleanText(value, "UNKNOWN")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function normalizeArea(value) {
  const text = cleanText(value, "UNKNOWN").toUpperCase();
  return text.replace(/\s+RETAIL\s+SA$/i, "").trim();
}

function normalizeDate(value) {
  if (!value) {
    return null;
  }

  if (typeof value === "number" && Number.isFinite(value)) {
    return excelSerialToIsoDate(value);
  }

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  const text = String(value).trim();
  if (!text) {
    return null;
  }

  const isoMatch = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  const dayFirstMatch = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})$/);
  if (dayFirstMatch) {
    let [, day, month, year] = dayFirstMatch;
    if (year.length === 2) {
      year = `20${year}`;
    }
    return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
  }

  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }

  return null;
}

function excelSerialToIsoDate(serial) {
  const numeric = Number(serial);
  if (!Number.isFinite(numeric)) {
    return null;
  }

  const wholeDays = Math.round(numeric);
  const excelEpochUtc = Date.UTC(1899, 11, 30);
  return new Date(excelEpochUtc + wholeDays * 86400000).toISOString().slice(0, 10);
}

function columnIndexFromCellReference(reference) {
  const letters = String(reference).match(/[A-Z]+/i)?.[0];
  if (!letters) {
    return -1;
  }

  return letters
    .toUpperCase()
    .split("")
    .reduce((total, letter) => total * 26 + (letter.charCodeAt(0) - 64), 0) - 1;
}

function parseXmlDocument(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function readInlineCellText(node) {
  return Array.from(node.getElementsByTagNameNS("*", "t"))
    .map((item) => item.textContent || "")
    .join("");
}

function parseSharedStrings(xmlText) {
  if (!xmlText) {
    return [];
  }

  const xml = parseXmlDocument(xmlText);
  return Array.from(xml.getElementsByTagNameNS("*", "si")).map((node) => readInlineCellText(node));
}

function readWorksheetCellValue(cell, sharedStrings) {
  const type = cell.getAttribute("t");
  if (type === "inlineStr") {
    return readInlineCellText(cell);
  }

  const valueNode = cell.getElementsByTagNameNS("*", "v")[0];
  if (!valueNode) {
    return "";
  }

  const rawValue = valueNode.textContent || "";
  if (type === "s") {
    return sharedStrings[Number(rawValue)] || "";
  }
  if (type === "b") {
    return rawValue === "1";
  }

  const numeric = Number(rawValue);
  return Number.isFinite(numeric) ? numeric : rawValue;
}

function readWorksheetRowValues(row, sharedStrings) {
  const valuesByColumn = new Map();
  Array.from(row.getElementsByTagNameNS("*", "c")).forEach((cell) => {
    valuesByColumn.set(
      columnIndexFromCellReference(cell.getAttribute("r") || ""),
      readWorksheetCellValue(cell, sharedStrings),
    );
  });
  return valuesByColumn;
}

function expandRowValues(valuesByColumn) {
  const maxColumn = Math.max(...valuesByColumn.keys(), -1);
  return Array.from({ length: maxColumn + 1 }, (_, index) => valuesByColumn.get(index) ?? "");
}

function detectHistoricalWorksheetFormat(headerValues) {
  const normalizedHeaders = headerValues.map((value) => normalizeHeader(value));
  if (
    normalizedHeaders.includes("shiptoparty") &&
    normalizedHeaders.includes("materialdescription") &&
    normalizedHeaders.includes("billingdate")
  ) {
    return "export";
  }

  const hasDateColumns = headerValues.some((value, index) => {
    if (index < 4) {
      return false;
    }
    const normalizedValue =
      typeof value === "number" ? excelSerialToIsoDate(value) : normalizeDate(cleanText(value));
    return Boolean(normalizedValue);
  });
  if (hasDateColumns) {
    return "matrix";
  }

  return "unknown";
}

function detectHistoricalUploadProduct(fileName, selectedProduct) {
  const override = cleanText(selectedProduct, "AUTO").toUpperCase();
  if (override === "MS" || override === "HSD" || override === "POWER") {
    return override;
  }

  const signal = cleanText(fileName).toUpperCase();
  if (/(^|[^A-Z])HSD([^A-Z]|$)/.test(signal)) {
    return "HSD";
  }
  if (/(^|[^A-Z])MS([^A-Z]|$)/.test(signal)) {
    return "MS";
  }

  throw new Error(`Could not detect whether ${fileName} is an MS, HSD, or Power backup. Choose the product manually and upload again.`);
}

function parseHistoricalMatrixFile(file, rows, sharedStrings, selectedProduct) {
  const product = detectHistoricalUploadProduct(file.name, selectedProduct);
  const headerMap = new Map();
  const headerValuesByColumn = readWorksheetRowValues(rows[0], sharedStrings);

  headerValuesByColumn.forEach((cellValue, columnIndex) => {
    if (columnIndex < 4) {
      return;
    }
    const normalizedDate =
      typeof cellValue === "number" ? excelSerialToIsoDate(cellValue) : normalizeDate(cleanText(cellValue));
    if (!normalizedDate) {
      return;
    }
    headerMap.set(columnIndex, normalizedDate.slice(5));
  });

  if (!headerMap.size) {
    throw new Error(`${file.name} must have date columns in the first row.`);
  }

  const lookup = {};
  let entryCount = 0;

  rows.slice(1).forEach((row) => {
    const valuesByColumn = readWorksheetRowValues(row, sharedStrings);
    if (!valuesByColumn.size) {
      return;
    }

    const sapCode = normalizeCode(valuesByColumn.get(0));
    if (!sapCode) {
      return;
    }

    headerMap.forEach((monthDay, columnIndex) => {
      const historicalKl = Number(toNumber(valuesByColumn.get(columnIndex)).toFixed(3));
      if (historicalKl <= 0) {
        return;
      }
      lookup[`${sapCode}|${monthDay}`] = historicalKl;
      entryCount += 1;
    });
  });

  if (!entryCount) {
    throw new Error(`${file.name} did not contain any positive Hist. KL values.`);
  }

  return [
    normalizeHistoricalImportFile({
      id: `${product}-${slugify(file.name)}`,
      name: file.name,
      product,
      importedAt: new Date().toISOString(),
      entryCount,
      lookup,
    }),
  ];
}

function findHeaderIndex(headerValues, aliases) {
  const normalizedHeaders = headerValues.map((value) => normalizeHeader(value));
  return aliases
    .map((alias) => normalizedHeaders.indexOf(alias))
    .find((index) => index !== -1);
}

function deriveHistoricalSalesKl(value, unit) {
  const numericValue = toNumber(value);
  if (!numericValue) {
    return 0;
  }

  const normalizedUnit = cleanText(unit).toUpperCase();
  if (["L", "LTR", "LITRE", "LITRES"].includes(normalizedUnit)) {
    return numericValue / 1000;
  }
  if (normalizedUnit === "KL") {
    return numericValue;
  }
  return numericValue >= 1000 ? numericValue / 1000 : numericValue;
}

function parseHistoricalExportFile(file, rows, sharedStrings) {
  const headerValues = expandRowValues(readWorksheetRowValues(rows[0], sharedStrings));
  const shipToIndex = findHeaderIndex(headerValues, ["shiptoparty"]);
  const materialIndex = findHeaderIndex(headerValues, ["material"]);
  const materialDescIndex = findHeaderIndex(headerValues, ["materialdescription", "itemdescription"]);
  const billingDateIndex = findHeaderIndex(headerValues, ["billingdate"]);
  const volumeIndex = findHeaderIndex(headerValues, ["volume"]);
  const quantityIndex = findHeaderIndex(headerValues, ["quantity"]);
  const salesUnitIndex = findHeaderIndex(headerValues, ["salesunit", "volumeunit"]);

  if (
    shipToIndex === undefined ||
    materialDescIndex === undefined ||
    billingDateIndex === undefined ||
    (volumeIndex === undefined && quantityIndex === undefined)
  ) {
    throw new Error(`${file.name} is missing one or more required export columns.`);
  }

  const bucketsByProduct = {
    MS: new Map(),
    HSD: new Map(),
    Power: new Map(),
  };

  rows.slice(1).forEach((row) => {
    const valuesByColumn = readWorksheetRowValues(row, sharedStrings);
    if (!valuesByColumn.size) {
      return;
    }

    const sapCode = normalizeCode(valuesByColumn.get(shipToIndex));
    const billingDate = normalizeDate(valuesByColumn.get(billingDateIndex));
    if (!sapCode || !billingDate) {
      return;
    }

    const materialDescription = cleanText(valuesByColumn.get(materialDescIndex));
    const material = cleanText(valuesByColumn.get(materialIndex));
    const product = canonicalizeProduct("", materialDescription, material);
    if (!["MS", "HSD", "Power"].includes(product)) {
      return;
    }

    const unit = valuesByColumn.get(salesUnitIndex ?? volumeIndex ?? quantityIndex);
    const historicalKl = Number(deriveHistoricalSalesKl(valuesByColumn.get(volumeIndex ?? quantityIndex), unit || "L").toFixed(3));
    if (historicalKl <= 0) {
      return;
    }

    const lookupKey = `${sapCode}|${billingDate.slice(5)}`;
    const datedBucket = bucketsByProduct[product].get(lookupKey) || new Map();
    datedBucket.set(billingDate, Number(((datedBucket.get(billingDate) || 0) + historicalKl).toFixed(3)));
    bucketsByProduct[product].set(lookupKey, datedBucket);
  });

  const lookupByProduct = {
    MS: {},
    HSD: {},
    Power: {},
  };
  Object.entries(bucketsByProduct).forEach(([product, productBuckets]) => {
    productBuckets.forEach((datedBucket, lookupKey) => {
      const firstHistoricalDate = [...datedBucket.keys()].sort()[0];
      lookupByProduct[product][lookupKey] = datedBucket.get(firstHistoricalDate);
    });
  });

  const importedAt = new Date().toISOString();
  const entries = ["MS", "HSD", "Power"]
    .map((product) => {
      const lookup = lookupByProduct[product];
      const entryCount = Object.keys(lookup).length;
      if (!entryCount) {
        return null;
      }
      return normalizeHistoricalImportFile({
        id: `${slugify(file.name)}-${product}`,
        name: `${file.name} (${product})`,
        product,
        importedAt,
        entryCount,
        lookup,
      });
    })
    .filter(Boolean);

  if (!entries.length) {
    throw new Error(`${file.name} did not contain usable MS, HSD, or Power Hist. rows.`);
  }

  return entries;
}

async function readFirstExcelWorksheet(file) {
  if (!window.JSZip) {
    throw new Error("Excel import support is not available on this page right now.");
  }

  const archive = await window.JSZip.loadAsync(await file.arrayBuffer());
  const worksheetPath = Object.keys(archive.files)
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((left, right) => left.localeCompare(right))[0];

  if (!worksheetPath) {
    throw new Error(`${file.name} does not contain a readable worksheet.`);
  }

  const [sheetXmlText, sharedStringsXmlText] = await Promise.all([
    archive.file(worksheetPath)?.async("string"),
    archive.file("xl/sharedStrings.xml")?.async("string") || Promise.resolve(""),
  ]);

  if (!sheetXmlText) {
    throw new Error(`${file.name} could not be read.`);
  }

  const sharedStrings = parseSharedStrings(sharedStringsXmlText);
  const worksheetXml = parseXmlDocument(sheetXmlText);
  const rows = Array.from(worksheetXml.getElementsByTagNameNS("*", "row"));
  if (!rows.length) {
    throw new Error(`${file.name} does not contain worksheet rows.`);
  }

  return { rows, sharedStrings };
}

async function parseHistoricalExcelFile(file, selectedProduct) {
  const { rows, sharedStrings } = await readFirstExcelWorksheet(file);
  const headerValues = expandRowValues(readWorksheetRowValues(rows[0], sharedStrings));
  const format = detectHistoricalWorksheetFormat(headerValues);

  if (format === "export") {
    return parseHistoricalExportFile(file, rows, sharedStrings);
  }
  if (format === "matrix") {
    return parseHistoricalMatrixFile(file, rows, sharedStrings, selectedProduct);
  }

  throw new Error(`${file.name} is not a supported Hist. workbook format.`);
}

function canonicalizeProduct(productRaw, materialDescriptionRaw, materialRaw) {
  const explicit = cleanText(productRaw);
  const materialDescription = cleanText(materialDescriptionRaw).toUpperCase().replace(/\s+/g, " ");
  const material = cleanText(materialRaw).toUpperCase().replace(/\s+/g, " ");
  const signal = `${explicit} ${materialDescription} ${material}`.toUpperCase();

  // Raw export material mapping: P95 is Power, normal EBMS is MS, and HSD BS VI is HSD.
  if (/^EBMS\s+P95-HSN-27101242\b/.test(materialDescription) || /\bP95\b/.test(signal) || /\bPOWER\b/.test(signal)) {
    return "Power";
  }
  if (/^HSD\s*-\s*BS\s*VI\b/.test(materialDescription) || /\bHSD\b/.test(signal)) {
    return "HSD";
  }
  if (signal.includes("LUBE") || signal.includes("LUBRICANT") || /\bLUB\b/.test(signal)) {
    return "Lube";
  }
  if (signal.includes("DEF") || signal.includes("AUS32") || signal.includes("UREA")) {
    return "DEF";
  }
  if (signal.includes("E-20") || /\bE20\b/.test(signal)) {
    return "E20";
  }
  if (
    /^EBMS-HSN-2710124[23]\b/.test(materialDescription) ||
    /\bMS\b/.test(signal) ||
    /\bPETROL\b/.test(signal) ||
    /\bEBMS\b/.test(signal)
  ) {
    return "MS";
  }

  return explicit;
}

function parseProjectFlag(rawFlag, outletType, sapCode, projectCodes) {
  if (typeof rawFlag === "boolean") {
    return rawFlag;
  }
  const signal = `${cleanText(outletType)} ${cleanText(rawFlag)}`.toUpperCase();
  if (signal.includes("PROJECT") || signal.includes("ABHUYADAY") || signal === "YES" || signal === "TRUE") {
    return true;
  }
  if (signal.includes("NON-PROJECT") || signal.includes("STANDARD") || signal === "NO" || signal === "FALSE") {
    return false;
  }
  return projectCodes.has(sapCode);
}

function buildProductList(detectedProducts) {
  const extras = detectedProducts
    .filter((product) => product && !KNOWN_PRODUCT_ORDER.includes(product))
    .sort((left, right) => left.localeCompare(right));
  return [...KNOWN_PRODUCT_ORDER, ...extras];
}

function normalizeRow(record, projectCodes) {
  const date = normalizeDate(record.date);
  const salesArea = normalizeArea(record.salesArea || record.salesAreaFull);
  const outletName = cleanText(record.outletName, "UNKNOWN OUTLET");
  const rawSapCode = normalizeCode(record.sapCode);
  const sapCode = rawSapCode || slugify(outletName);
  const product = canonicalizeProduct(record.product, record.materialDescription, record.material);
  let salesUnits = toNumber(record.salesUnits);
  let netVolume = toNumber(record.netVolume);
  const billingDocument = cleanText(record.billingDocument);
  const documents = Math.max(1, Math.round(toNumber(record.documents || (billingDocument ? 1 : 1))));

  if (!salesUnits && netVolume) {
    salesUnits = record.netVolumeIsLiter ? deriveHistoricalSalesKl(netVolume, "L") : deriveHistoricalSalesKl(netVolume, record.volumeUnit);
  }
  if (!netVolume && salesUnits) {
    netVolume = salesUnits * 1000;
  }

  if (!date || !product || !outletName) {
    return null;
  }
  if (salesUnits === 0 && netVolume === 0 && !billingDocument) {
    return null;
  }

  return {
    product,
    date,
    salesArea,
    salesAreaFull: cleanText(record.salesAreaFull || record.salesArea, salesArea),
    sapCode,
    outletName,
    salesUnits: Number(salesUnits.toFixed(3)),
    netVolume: Number(netVolume.toFixed(3)),
    billingDocument,
    material: cleanText(record.material),
    materialDescription: cleanText(record.materialDescription),
    plant: plantDescription(record.plant),
    isProjectAbhuyaday: parseProjectFlag(
      record.isProjectAbhuyaday,
      record.outletType,
      sapCode,
      projectCodes,
    ),
    documents,
  };
}

function normalizeHistoricalRow(record, projectCodes) {
  const date = normalizeDate(record.date || record.referenceDate);
  const salesArea = normalizeArea(record.salesArea || record.salesAreaFull);
  const outletName = cleanText(record.outletName, "UNKNOWN OUTLET");
  const rawSapCode = normalizeCode(record.sapCode);
  const sapCode = rawSapCode || slugify(outletName);
  const product = canonicalizeProduct(record.product, record.materialDescription, record.material);
  const historicalKl = toNumber(record.historicalKl ?? record.salesUnits ?? record.volume);

  if (!date || !sapCode || !product || historicalKl <= 0 || !HISTORICAL_COMPARISON_PRODUCTS.includes(product)) {
    return null;
  }

  return {
    date,
    salesArea,
    salesAreaFull: cleanText(record.salesAreaFull || record.salesArea, salesArea),
    sapCode,
    outletName,
    product,
    historicalKl: Number(historicalKl.toFixed(3)),
    plant: plantDescription(record.plant),
    isProjectAbhuyaday: parseProjectFlag(
      record.isProjectAbhuyaday,
      record.outletType,
      sapCode,
      projectCodes,
    ),
  };
}

function buildDatasetFromRows(rawRows, options) {
  const projectCodes = new Set(options.projectCodes || []);
  const rows = rawRows
    .map((row) => normalizeRow(row, projectCodes))
    .filter(Boolean)
    .sort((left, right) => {
      const dateCompare = left.date.localeCompare(right.date);
      if (dateCompare !== 0) {
        return dateCompare;
      }
      const areaCompare = left.salesArea.localeCompare(right.salesArea);
      if (areaCompare !== 0) {
        return areaCompare;
      }
      const outletCompare = left.outletName.localeCompare(right.outletName);
      if (outletCompare !== 0) {
        return outletCompare;
      }
      return left.product.localeCompare(right.product);
    });

  if (!rows.length) {
    throw new Error("No usable sales rows were found in that file.");
  }

  const detectedProducts = [...new Set(rows.map((row) => row.product))];
  const products = buildProductList(detectedProducts);
  const areas = [...new Set(rows.map((row) => row.salesArea))].sort((left, right) => left.localeCompare(right));
  const dates = [...new Set(rows.map((row) => row.date))].sort((left, right) => left.localeCompare(right));
  const outletMap = new Map();
  const recordsPerDate = {};
  const productStats = {};

  products.forEach((product) => {
    productStats[product] = {
      records: 0,
      outlets: 0,
      areas: [],
      dateMin: null,
      dateMax: null,
      paOutlets: 0,
      salesUnits: 0,
      netVolume: 0,
    };
  });

  rows.forEach((row) => {
    recordsPerDate[row.date] = (recordsPerDate[row.date] || 0) + 1;
    const existingOutlet = outletMap.get(row.sapCode) || {
      sapCode: row.sapCode,
      outletName: row.outletName,
      salesArea: row.salesArea,
      plant: row.plant,
      isProjectAbhuyaday: row.isProjectAbhuyaday,
    };
    existingOutlet.plant = existingOutlet.plant || row.plant;
    outletMap.set(row.sapCode, existingOutlet);

    const stats = productStats[row.product];
    stats.records += 1;
    stats.salesUnits += row.salesUnits;
    stats.netVolume += row.netVolume;
    stats.dateMin = stats.dateMin ? (row.date < stats.dateMin ? row.date : stats.dateMin) : row.date;
    stats.dateMax = stats.dateMax ? (row.date > stats.dateMax ? row.date : stats.dateMax) : row.date;
  });

  products.forEach((product) => {
    const productRows = rows.filter((row) => row.product === product);
    const outletCodes = new Set(productRows.map((row) => row.sapCode));
    const paOutletCodes = new Set(productRows.filter((row) => row.isProjectAbhuyaday).map((row) => row.sapCode));
    const areaList = [...new Set(productRows.map((row) => row.salesArea))].sort((left, right) =>
      left.localeCompare(right),
    );

    productStats[product].outlets = outletCodes.size;
    productStats[product].paOutlets = paOutletCodes.size;
    productStats[product].areas = areaList;
    productStats[product].salesUnits = Number(productStats[product].salesUnits.toFixed(3));
    productStats[product].netVolume = Number(productStats[product].netVolume.toFixed(3));
  });

  const outlets = sortByName([...outletMap.values()]);
  const plants = collectDepotOptions(rows, outlets);
  const activePaOutlets = new Set(
    outlets.filter((outlet) => outlet.isProjectAbhuyaday).map((outlet) => outlet.sapCode),
  );

  return {
    rows,
    filters: {
      dates,
      areas,
      plants,
      outlets,
    },
    meta: {
      title: "Jabalpur Retail Region Sales Dashboard",
      sourceFile: options.sourceFile,
      sourceMode: options.sourceMode,
      generatedAt: options.generatedAt,
      dateMin: dates[0],
      dateMax: dates[dates.length - 1],
      defaultDate: dates[dates.length - 1],
      dateCount: dates.length,
      areas,
      plants,
      products,
      emptyProducts: products.filter((product) => productStats[product].records === 0),
      paOutletCount: projectCodes.size || activePaOutlets.size,
      projectAbhuyadayCodes: [...projectCodes].sort((left, right) => left.localeCompare(right)),
      activeOutletCount: outlets.length,
      recordCount: rows.length,
      productStats,
      recordsPerDate,
    },
  };
}

function buildDatasetFromSnapshot(rawData) {
  const projectCodes = rawData.meta?.projectAbhuyadayCodes || DEFAULT_PROJECT_CODES;
  const dataset = buildDatasetFromRows(rawData.rows || [], {
    sourceFile: rawData.meta?.sourceFile || "Built-in workbook snapshot",
    sourceMode: rawData.meta?.sourceMode || "Built-in workbook snapshot",
    generatedAt: rawData.meta?.generatedAt || new Date().toISOString(),
    projectCodes,
  });
  const historicalRows = (rawData.historicalRows || [])
    .map((row) => normalizeHistoricalRow(row, new Set(projectCodes)))
    .filter(Boolean)
    .sort((left, right) => {
      const dateCompare = left.date.localeCompare(right.date);
      if (dateCompare !== 0) {
        return dateCompare;
      }
      const areaCompare = left.salesArea.localeCompare(right.salesArea);
      if (areaCompare !== 0) {
        return areaCompare;
      }
      const outletCompare = left.outletName.localeCompare(right.outletName);
      if (outletCompare !== 0) {
        return outletCompare;
      }
      return KNOWN_PRODUCT_ORDER.indexOf(left.product) - KNOWN_PRODUCT_ORDER.indexOf(right.product);
    });
  const historicalDates = [...new Set(historicalRows.map((row) => row.date))].sort((left, right) => left.localeCompare(right));
  dataset.historicalLookup = rawData.historicalLookup || DEFAULT_HISTORICAL_LOOKUP;
  dataset.historicalRows = historicalRows;
  dataset.areaTargets = Array.isArray(rawData.areaTargets) ? rawData.areaTargets : [];
  dataset.projectAbhuyadayReference = rawData.projectAbhuyadayReference || null;
  dataset.filters.historicalDates = historicalDates;
  dataset.meta = {
    ...dataset.meta,
    historicalSourceFile: rawData.meta?.historicalSourceFile || null,
    historicalDateMin: rawData.meta?.historicalDateMin || historicalDates[0] || null,
    historicalDateMax: rawData.meta?.historicalDateMax || historicalDates[historicalDates.length - 1] || null,
    historicalRecordCount: historicalRows.length,
    projectAbhuyadaySourceFile: rawData.meta?.projectAbhuyadaySourceFile || null,
    projectAbhuyadayReferenceCount: rawData.meta?.projectAbhuyadayReferenceCount || 0,
  };
  return dataset;
}

function loadMasterOutlets() {
  const raw = storageGet(STORAGE_KEYS.outletMaster);
  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveMasterOutlets(entries) {
  storageSet(STORAGE_KEYS.outletMaster, JSON.stringify(entries));
}

function loadDeletedOutletCodes() {
  const raw = storageGet(STORAGE_KEYS.deletedOutlets);
  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.map(normalizeCode).filter(Boolean) : [];
  } catch {
    return [];
  }
}

function saveDeletedOutletCodes(codes) {
  return storageSet(STORAGE_KEYS.deletedOutlets, JSON.stringify([...new Set(codes.map(normalizeCode).filter(Boolean))]));
}

function loadMasterHistoricalEntries() {
  const raw = storageGet(STORAGE_KEYS.historicalMaster);
  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveMasterHistoricalEntries(entries) {
  storageSet(STORAGE_KEYS.historicalMaster, JSON.stringify(entries));
}

function loadHistoricalImportFiles() {
  const raw = storageGet(STORAGE_KEYS.historicalImport);
  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveHistoricalImportFiles(entries) {
  return storageSet(STORAGE_KEYS.historicalImport, JSON.stringify(entries));
}

function loadAreaTargets() {
  const raw = storageGet(STORAGE_KEYS.areaTargets);
  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveAreaTargets(entries) {
  return storageSet(STORAGE_KEYS.areaTargets, JSON.stringify(entries));
}

function loadDeletedAreaTargetKeys() {
  const raw = storageGet(STORAGE_KEYS.deletedAreaTargets);
  if (!raw) {
    return [];
  }

  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed.map(cleanText).filter(Boolean) : [];
  } catch {
    return [];
  }
}

function saveDeletedAreaTargetKeys(keys) {
  return storageSet(STORAGE_KEYS.deletedAreaTargets, JSON.stringify([...new Set(keys.map(cleanText).filter(Boolean))]));
}

function normalizeHistoricalImportFile(entry) {
  const normalizedProduct = cleanText(entry.product).toUpperCase();
  const product = normalizedProduct === "HSD" ? "HSD" : normalizedProduct === "POWER" ? "Power" : "MS";
  const lookupEntries =
    entry.lookup && typeof entry.lookup === "object"
      ? Object.entries(entry.lookup)
          .map(([key, value]) => [cleanText(key), Number(toNumber(value).toFixed(3))])
          .filter(([key, value]) => key && Number.isFinite(value) && value > 0)
      : [];

  return {
    id: cleanText(entry.id, `${product}-${Date.now()}`),
    name: cleanText(entry.name, `${product} HIST.xlsx`),
    product,
    importedAt: cleanText(entry.importedAt, new Date().toISOString()),
    entryCount: Math.max(0, Math.trunc(toNumber(entry.entryCount) || lookupEntries.length)),
    lookup: Object.fromEntries(lookupEntries),
  };
}

function normalizeMasterOutlet(entry) {
  return {
    sapCode: normalizeCode(entry.sapCode),
    outletName: cleanText(entry.outletName),
    salesArea: normalizeArea(entry.salesArea),
    appointedOn: normalizeDate(entry.appointedOn),
    plant: cleanText(entry.plant),
    notes: cleanText(entry.notes),
    isProjectAbhuyaday: parseProjectFlag(entry.isProjectAbhuyaday, entry.outletType, normalizeCode(entry.sapCode), new Set()),
    outletType: parseProjectFlag(entry.isProjectAbhuyaday, entry.outletType, normalizeCode(entry.sapCode), new Set())
      ? "Project Abhuyaday"
      : "Standard",
  };
}

function normalizeMonthKey(value) {
  const text = cleanText(value);
  if (!text) {
    return "";
  }
  const monthMatch = text.match(/^(\d{4})-(\d{1,2})$/);
  if (monthMatch) {
    const [, year, month] = monthMatch;
    return `${year}-${String(month).padStart(2, "0")}`;
  }
  const normalizedDate = normalizeDate(text);
  return normalizedDate ? normalizedDate.slice(0, 7) : "";
}

function normalizeMasterHistoricalEntry(entry) {
  return {
    sapCode: normalizeCode(entry.sapCode),
    product: canonicalizeProduct(entry.product, "", ""),
    referenceDate: normalizeDate(entry.referenceDate),
    historicalKl: Number(toNumber(entry.historicalKl).toFixed(3)),
    note: cleanText(entry.note),
  };
}

function normalizeAreaTarget(entry) {
  const targets = entry.targets && typeof entry.targets === "object" ? entry.targets : entry;
  const monthKey = normalizeMonthKey(entry.monthKey ?? entry.targetMonth ?? entry.month);
  return {
    monthKey,
    salesArea: normalizeArea(entry.salesArea),
    targets: {
      MS: Number(toNumber(targets.MS ?? targets.msTarget).toFixed(3)),
      HSD: Number(toNumber(targets.HSD ?? targets.hsdTarget).toFixed(3)),
      Power: Number(toNumber(targets.Power ?? targets.powerTarget).toFixed(3)),
      Lube: Number(toNumber(targets.Lube ?? targets.lubeTarget).toFixed(3)),
      DEF: Number(toNumber(targets.DEF ?? targets.defTarget).toFixed(3)),
    },
    note: cleanText(entry.note),
    updatedAt: cleanText(entry.updatedAt, new Date().toISOString()),
  };
}

function isUsableAreaTarget(entry) {
  return entry.salesArea && TARGET_PRODUCTS.some((product) => entry.targets[product] > 0);
}

function areaTargetKey(entry) {
  return `${entry.monthKey || "ALL"}|${entry.salesArea}`;
}

function getBuiltInAreaTargetEntries() {
  const sourceTargets = baseData?.areaTargets?.length ? baseData.areaTargets : DEFAULT_DATA.areaTargets || [];
  return sourceTargets.map(normalizeAreaTarget).filter(isUsableAreaTarget);
}

function getCustomAreaTargetEntries() {
  return loadAreaTargets().map(normalizeAreaTarget).filter(isUsableAreaTarget);
}

function getAreaTargetEntries() {
  const deletedKeys = new Set(loadDeletedAreaTargetKeys());
  const merged = new Map();
  getBuiltInAreaTargetEntries().forEach((entry) => {
    const key = areaTargetKey(entry);
    if (!deletedKeys.has(key)) {
      merged.set(key, entry);
    }
  });
  getCustomAreaTargetEntries().forEach((entry) => {
    merged.set(areaTargetKey(entry), entry);
  });
  return [...merged.values()].filter(isUsableAreaTarget);
}

function getAreaTargetMap() {
  const targetMonth = getTargetMonthKey();
  const map = new Map();
  getAreaTargetEntries().forEach((entry) => {
    if (!entry.monthKey) {
      map.set(entry.salesArea, entry);
    }
  });
  getAreaTargetEntries().forEach((entry) => {
    if (entry.monthKey === targetMonth) {
      map.set(entry.salesArea, entry);
    }
  });
  return map;
}

function buildImportedHistoricalLookup(files) {
  const merged = {
    MS: {},
    HSD: {},
    Power: {},
  };

  files
    .map(normalizeHistoricalImportFile)
    .forEach((file) => {
      if (!merged[file.product]) {
        return;
      }
      Object.assign(merged[file.product], file.lookup);
    });

  return merged;
}

function buildHistoricalLookup(baseLookup, importedLookup, entries) {
  const merged = {
    MS: { ...(baseLookup?.MS || {}), ...(importedLookup?.MS || {}) },
    HSD: { ...(baseLookup?.HSD || {}), ...(importedLookup?.HSD || {}) },
    Power: { ...(baseLookup?.Power || {}), ...(importedLookup?.Power || {}) },
  };

  entries.forEach((entry) => {
    if (!entry.sapCode || !entry.referenceDate || !entry.product || entry.historicalKl <= 0 || !merged[entry.product]) {
      return;
    }
    merged[entry.product][`${entry.sapCode}|${entry.referenceDate.slice(5)}`] = entry.historicalKl;
  });

  return merged;
}

function createEffectiveData(dataset) {
  const includeHistoricalFilters = PAGE === "historical";
  const deletedOutletCodes = new Set(loadDeletedOutletCodes());
  const builtInMasterOutlets = (DEFAULT_DATA.filters?.outlets || [])
    .map(normalizeMasterOutlet)
    .filter((entry) => entry.sapCode && entry.outletName && !deletedOutletCodes.has(entry.sapCode));
  const masterOutlets = loadMasterOutlets()
    .map(normalizeMasterOutlet)
    .filter((entry) => entry.sapCode && entry.outletName && !deletedOutletCodes.has(entry.sapCode));
  const historicalImportFiles = loadHistoricalImportFiles().map(normalizeHistoricalImportFile);
  const masterHistoricalEntries = loadMasterHistoricalEntries()
    .map(normalizeMasterHistoricalEntry)
    .filter((entry) => entry.sapCode && entry.referenceDate && entry.product && entry.historicalKl > 0);
  const outletMasterMap = new Map([...builtInMasterOutlets, ...masterOutlets].map((entry) => [entry.sapCode, entry]));

  const rows = dataset.rows.filter((row) => !deletedOutletCodes.has(row.sapCode)).map((row) => {
    const master = outletMasterMap.get(row.sapCode);
    if (!master) {
      return row;
    }
    return {
      ...row,
      outletName: master.outletName || row.outletName,
      salesArea: master.salesArea || row.salesArea,
      salesAreaFull: master.salesArea ? `${master.salesArea} RETAIL SA` : row.salesAreaFull,
      plant: master.plant || row.plant,
      isProjectAbhuyaday: Boolean(master.isProjectAbhuyaday || row.isProjectAbhuyaday),
    };
  });

  const historicalRows = (dataset.historicalRows || []).filter((row) => !deletedOutletCodes.has(row.sapCode)).map((row) => {
    const master = outletMasterMap.get(row.sapCode);
    if (!master) {
      return row;
    }
    return {
      ...row,
      outletName: master.outletName || row.outletName,
      salesArea: master.salesArea || row.salesArea,
      salesAreaFull: master.salesArea ? `${master.salesArea} RETAIL SA` : row.salesAreaFull,
      plant: master.plant || row.plant,
      isProjectAbhuyaday: Boolean(master.isProjectAbhuyaday || row.isProjectAbhuyaday),
    };
  });

  const outletMap = new Map();
  const seedOutlets = []
    .concat(DEFAULT_DATA.filters?.outlets || [])
    .concat(dataset.filters?.outlets || [])
    .concat(
      rows.map((row) => ({
        sapCode: row.sapCode,
        outletName: row.outletName,
        salesArea: row.salesArea,
        plant: row.plant,
        isProjectAbhuyaday: row.isProjectAbhuyaday,
      })),
    )
    .concat(
      includeHistoricalFilters
        ? historicalRows.map((row) => ({
            sapCode: row.sapCode,
            outletName: row.outletName,
            salesArea: row.salesArea,
            plant: row.plant,
            isProjectAbhuyaday: row.isProjectAbhuyaday,
          }))
        : [],
    )
    .concat(
      masterOutlets.map((entry) => ({
        sapCode: entry.sapCode,
        outletName: entry.outletName,
        salesArea: entry.salesArea,
        isProjectAbhuyaday: entry.isProjectAbhuyaday,
        appointedOn: entry.appointedOn,
        plant: entry.plant,
        notes: entry.notes,
        outletType: entry.outletType,
      })),
    );

  seedOutlets.forEach((outlet) => {
    if (!outlet.sapCode) {
      return;
    }
    if (deletedOutletCodes.has(outlet.sapCode)) {
      return;
    }
    const existing = outletMap.get(outlet.sapCode) || {};
    outletMap.set(outlet.sapCode, {
      sapCode: outlet.sapCode,
      outletName: cleanText(outlet.outletName, existing.outletName || "UNKNOWN OUTLET"),
      salesArea: normalizeArea(outlet.salesArea || existing.salesArea),
      isProjectAbhuyaday: Boolean(outlet.isProjectAbhuyaday || existing.isProjectAbhuyaday),
      appointedOn: normalizeDate(outlet.appointedOn) || existing.appointedOn || "",
      plant: cleanText(outlet.plant, existing.plant || ""),
      notes: cleanText(outlet.notes, existing.notes || ""),
      outletType: cleanText(outlet.outletType, outlet.isProjectAbhuyaday || existing.isProjectAbhuyaday ? "Project Abhuyaday" : "Standard"),
    });
  });

  const outlets = sortByName([...outletMap.values()]);
  const areas = [
    ...new Set(
      outlets
        .map((outlet) => outlet.salesArea)
        .concat(rows.map((row) => row.salesArea))
        .concat(includeHistoricalFilters ? historicalRows.map((row) => row.salesArea) : []),
    ),
  ]
    .filter(Boolean)
    .sort((left, right) => left.localeCompare(right));
  const plants = collectDepotOptions(includeHistoricalFilters ? rows.concat(historicalRows) : rows, outlets);
  const historicalDates = [...new Set(historicalRows.map((row) => row.date))].sort((left, right) => left.localeCompare(right));
  const projectCodes = outlets.filter((outlet) => outlet.isProjectAbhuyaday).map((outlet) => outlet.sapCode);
  const importedHistoricalLookup = buildImportedHistoricalLookup(historicalImportFiles);
  const historicalLookup = buildHistoricalLookup(
    dataset.historicalLookup || DEFAULT_HISTORICAL_LOOKUP,
    importedHistoricalLookup,
    masterHistoricalEntries,
  );

  return {
    ...dataset,
    rows,
    historicalRows,
    filters: {
      ...dataset.filters,
      areas,
      plants,
      outlets,
      historicalDates,
    },
    meta: {
      ...dataset.meta,
      areas,
      plants,
      historicalDateMin: historicalDates[0] || dataset.meta.historicalDateMin || null,
      historicalDateMax: historicalDates[historicalDates.length - 1] || dataset.meta.historicalDateMax || null,
      historicalRecordCount: historicalRows.length,
      paOutletCount: projectCodes.length,
      projectAbhuyadayCodes: projectCodes,
      activeOutletCount: outlets.length,
    },
    historicalLookup,
  };
}

function normalizeHeader(value) {
  return cleanText(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

const CSV_FIELD_ALIASES = {
  date: ["date", "billingdate"],
  salesArea: ["salesarea", "salesgroupdesc", "salesgroup", "salesareafull"],
  outletName: ["outletname", "shiptopartyname", "shiptoparty", "dealername", "nameofdealer"],
  sapCode: ["sapcode", "shiptoparty", "outletcode", "dealercode"],
  product: ["product", "productname", "fuel", "item"],
  salesUnits: ["saleskl", "salesunits", "salesunit", "sumofsalesunit", "actual", "kl"],
  netVolume: ["netvolume", "quantity", "volume", "netquantity"],
  volumeUnit: ["salesunit", "volumeunit"],
  billingDocument: ["billingdocumentno", "billingdocument", "documentno", "invoiceno", "invoice"],
  material: ["material", "materialcode"],
  materialDescription: ["materialdescription", "materialdesc", "description"],
  plant: ["plantdescription", "plant"],
  outletType: ["outlettype", "projectabhuyaday", "paoutlet"],
  documents: ["documents", "documentcount"],
};

const WORKSHEET_SALES_KL_ALIASES = ["saleskl", "salesunits", "sumofsalesunit", "actual", "kl"];

function detectDelimiter(text) {
  const sample = text
    .split(/\r?\n/)
    .filter((line) => line.trim())
    .slice(0, 5)
    .join("\n");
  const candidates = [",", ";", "\t", "|"];
  let best = ",";
  let bestScore = -1;

  candidates.forEach((candidate) => {
    const score = sample.split(candidate).length - 1;
    if (score > bestScore) {
      best = candidate;
      bestScore = score;
    }
  });

  return best;
}

function parseCsvMatrix(text, delimiter) {
  const rows = [];
  let row = [];
  let value = "";
  let inQuotes = false;
  const source = text.replace(/^\uFEFF/, "");

  for (let index = 0; index < source.length; index += 1) {
    const character = source[index];
    const nextCharacter = source[index + 1];

    if (character === '"') {
      if (inQuotes && nextCharacter === '"') {
        value += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (character === delimiter && !inQuotes) {
      row.push(value);
      value = "";
      continue;
    }

    if ((character === "\n" || character === "\r") && !inQuotes) {
      if (character === "\r" && nextCharacter === "\n") {
        index += 1;
      }
      row.push(value);
      if (row.some((cell) => cleanText(cell))) {
        rows.push(row);
      }
      row = [];
      value = "";
      continue;
    }

    value += character;
  }

  row.push(value);
  if (row.some((cell) => cleanText(cell))) {
    rows.push(row);
  }

  return rows;
}

function resolveCsvValue(record, aliasList) {
  for (const alias of aliasList) {
    if (alias in record && cleanText(record[alias])) {
      return record[alias];
    }
  }
  return "";
}

function buildDatasetFromCsv(text, fileName) {
  const delimiter = detectDelimiter(text);
  const matrix = parseCsvMatrix(text, delimiter);
  if (matrix.length < 2) {
    throw new Error("The uploaded CSV is empty or does not contain any sales rows.");
  }

  const headers = matrix[0];
  const normalizedHeaders = headers.map((header) => normalizeHeader(header));
  const netVolumeIndex = findHeaderIndex(headers, CSV_FIELD_ALIASES.netVolume);
  const netVolumeHeader = netVolumeIndex === undefined ? "" : normalizedHeaders[netVolumeIndex];
  const rawRows = matrix.slice(1).map((cells) => {
    const record = {};
    normalizedHeaders.forEach((header, index) => {
      record[header] = cleanText(cells[index]);
    });

    return {
      date: resolveCsvValue(record, CSV_FIELD_ALIASES.date),
      salesArea: resolveCsvValue(record, CSV_FIELD_ALIASES.salesArea),
      outletName: resolveCsvValue(record, CSV_FIELD_ALIASES.outletName),
      sapCode: resolveCsvValue(record, CSV_FIELD_ALIASES.sapCode),
      product: resolveCsvValue(record, CSV_FIELD_ALIASES.product),
      salesUnits: resolveCsvValue(record, CSV_FIELD_ALIASES.salesUnits),
      netVolume: resolveCsvValue(record, CSV_FIELD_ALIASES.netVolume),
      volumeUnit: resolveCsvValue(record, CSV_FIELD_ALIASES.volumeUnit),
      netVolumeIsLiter: netVolumeHeader === "volume",
      billingDocument: resolveCsvValue(record, CSV_FIELD_ALIASES.billingDocument),
      material: resolveCsvValue(record, CSV_FIELD_ALIASES.material),
      materialDescription: resolveCsvValue(record, CSV_FIELD_ALIASES.materialDescription),
      plant: resolveCsvValue(record, CSV_FIELD_ALIASES.plant),
      outletType: resolveCsvValue(record, CSV_FIELD_ALIASES.outletType),
      documents: resolveCsvValue(record, CSV_FIELD_ALIASES.documents),
    };
  });

  return buildDatasetFromRows(rawRows, {
    sourceFile: fileName,
    sourceMode: "Uploaded CSV",
    generatedAt: new Date().toISOString(),
    projectCodes: builtInData.meta.projectAbhuyadayCodes,
  });
}

function resolveWorksheetValue(valuesByColumn, headerValues, aliasList) {
  const index = findHeaderIndex(headerValues, aliasList);
  if (index === undefined) {
    return "";
  }
  return valuesByColumn.get(index);
}

function buildSalesRecordFromWorksheetRow(valuesByColumn, headerValues) {
  const volume = resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.netVolume);
  const directSalesKl = resolveWorksheetValue(valuesByColumn, headerValues, WORKSHEET_SALES_KL_ALIASES);
  const unit = resolveWorksheetValue(valuesByColumn, headerValues, ["salesunit", "volumeunit"]);
  const materialDescription = resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.materialDescription);
  const material = resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.material);
  const salesUnits = cleanText(volume) ? deriveHistoricalSalesKl(volume, unit || "L") : toNumber(directSalesKl);

  return {
    date: resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.date),
    salesArea: resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.salesArea),
    outletName: resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.outletName),
    sapCode: resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.sapCode),
    product: canonicalizeProduct(
      resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.product),
      materialDescription,
      material,
    ),
    salesUnits,
    netVolume: cleanText(volume) ? volume : salesUnits * 1000,
    billingDocument: resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.billingDocument),
    material,
    materialDescription,
    plant: resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.plant),
    outletType: resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.outletType),
    documents: resolveWorksheetValue(valuesByColumn, headerValues, CSV_FIELD_ALIASES.documents),
  };
}

async function buildDatasetFromExcelFile(file) {
  const { rows, sharedStrings } = await readFirstExcelWorksheet(file);
  if (!rows.length) {
    throw new Error("The uploaded Excel file does not contain any readable rows.");
  }
  const headerValues = expandRowValues(readWorksheetRowValues(rows[0], sharedStrings));
  const rawRows = rows.slice(1).map((row) => buildSalesRecordFromWorksheetRow(readWorksheetRowValues(row, sharedStrings), headerValues));

  return buildDatasetFromRows(rawRows, {
    sourceFile: file.name,
    sourceMode: "Uploaded Excel",
    generatedAt: new Date().toISOString(),
    projectCodes: builtInData.meta.projectAbhuyadayCodes,
  });
}

function isExcelFile(file) {
  return /\.(xlsx|xlsm)$/i.test(file.name);
}

function storageGet(key) {
  try {
    return window.localStorage.getItem(key);
  } catch {
    return null;
  }
}

function storageSet(key, value) {
  try {
    window.localStorage.setItem(key, value);
    return true;
  } catch {
    // Ignore storage failures on restricted browser contexts.
    return false;
  }
}

function storageRemove(key) {
  try {
    window.localStorage.removeItem(key);
  } catch {
    // Ignore storage failures on restricted browser contexts.
  }
}

function loadPersistedDataset() {
  const raw = storageGet(STORAGE_KEYS.dataset);
  if (!raw) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw);
    if (!parsed || !parsed.rows || !parsed.meta || !parsed.filters) {
      return null;
    }
    return parsed;
  } catch {
    return null;
  }
}

function savePersistedDataset(dataset) {
  return storageSet(STORAGE_KEYS.dataset, JSON.stringify(dataset));
}

function loadPersistedState() {
  const raw = storageGet(STORAGE_KEYS.state);
  if (!raw) {
    return null;
  }

  try {
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : null;
  } catch {
    return null;
  }
}

function savePersistedState(state) {
  storageSet(STORAGE_KEYS.state, JSON.stringify(state));
}

function getPageDateRange(data) {
  if (PAGE === "historical") {
    const historicalDates = data.filters?.historicalDates || [];
    const min = data.meta?.historicalDateMin || historicalDates[0] || data.meta.dateMin;
    const max = data.meta?.historicalDateMax || historicalDates[historicalDates.length - 1] || data.meta.dateMax;
    return { min, max, defaultDate: max };
  }
  return {
    min: data.meta.dateMin,
    max: data.meta.dateMax,
    defaultDate: data.meta.defaultDate,
  };
}

function buildDefaultState(data) {
  const pageRange = getPageDateRange(data);
  if (PAGE === "historical") {
    return {
      startDate: pageRange.min,
      endDate: pageRange.max,
      salesArea: "ALL",
      plantDescription: "ALL",
      outletCode: "ALL",
      outletSearch: "",
      paFilter: "all",
      product: "ALL",
    };
  }
  const latestDate = pageRange.defaultDate;
  const latestMonthStart = latestDate ? `${latestDate.slice(0, 7)}-01` : pageRange.min;
  return {
    startDate: clampDateToDataRange(latestMonthStart, data, pageRange.min),
    endDate: pageRange.defaultDate,
    salesArea: "ALL",
    plantDescription: "ALL",
    outletCode: "ALL",
    outletSearch: "",
    paFilter: "all",
    product: "ALL",
  };
}

function isDateInsideDataRange(value, data) {
  const pageRange = getPageDateRange(data);
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "")) && value >= pageRange.min && value <= pageRange.max;
}

function clampDateToDataRange(value, data, fallback = data.meta.defaultDate) {
  const pageRange = getPageDateRange(data);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(value || ""))) {
    return fallback;
  }
  if (value < pageRange.min) {
    return pageRange.min;
  }
  if (value > pageRange.max) {
    return pageRange.max;
  }
  return value;
}

function clampStateToData(state, data) {
  const next = {
    ...buildDefaultState(data),
    ...(state || {}),
  };

  const availableAreas = new Set(["ALL", ...data.filters.areas]);
  const availablePlants = new Set(["ALL", ...(data.filters.plants || [])]);
  const availableProducts = new Set(["ALL", ...data.meta.products]);
  const availableOutlets = data.filters.outlets;

  if (!availableAreas.has(next.salesArea)) {
    next.salesArea = "ALL";
  }
  if (!availablePlants.has(next.plantDescription)) {
    next.plantDescription = "ALL";
  }
  if (!availableProducts.has(next.product)) {
    next.product = "ALL";
  }
  if (!["all", "project", "non-project"].includes(next.paFilter)) {
    next.paFilter = "all";
  }
  next.outletSearch = cleanText(next.outletSearch);

  next.startDate = isDateInsideDataRange(next.startDate, data)
    ? next.startDate
    : clampDateToDataRange(next.startDate, data, data.meta.dateMin);
  next.endDate = isDateInsideDataRange(next.endDate, data)
    ? next.endDate
    : clampDateToDataRange(next.endDate, data, data.meta.dateMax);
  if (next.startDate > next.endDate) {
    next.endDate = next.startDate;
  }

  const scopedOutlets = availableOutlets.filter(
    (outlet) =>
      (next.salesArea === "ALL" || outlet.salesArea === next.salesArea) &&
      (!elements.plantFilter || next.plantDescription === "ALL" || plantDescription(outlet.plant) === next.plantDescription),
  );
  const validOutletCodes = new Set(["ALL", ...scopedOutlets.map((outlet) => outlet.sapCode)]);
  if (!validOutletCodes.has(next.outletCode)) {
    next.outletCode = "ALL";
  }

  return next;
}

function setUploadStatus(message, tone = "info") {
  if (!elements.uploadStatus) {
    return;
  }
  elements.uploadStatus.textContent = message;
  elements.uploadStatus.classList.remove("is-success", "is-error");
  if (tone === "success") {
    elements.uploadStatus.classList.add("is-success");
  }
  if (tone === "error") {
    elements.uploadStatus.classList.add("is-error");
  }
}

function initializeMeta() {
  const pageRange = getPageDateRange(currentData);
  if (elements.sourceFile) {
    elements.sourceFile.textContent =
      PAGE === "historical" ? currentData.meta.historicalSourceFile || currentData.meta.sourceFile : currentData.meta.sourceFile;
  }
  if (elements.sourcePeriod) {
    elements.sourcePeriod.textContent = `${displayDate(pageRange.min)} to ${displayDate(pageRange.max)}`;
  }
  if (elements.generatedAt) {
    elements.generatedAt.textContent = new Date(currentData.meta.generatedAt).toLocaleString("en-IN");
  }
}

function syncControls() {
  const pageRange = getPageDateRange(currentData);
  if (elements.startDate) {
    elements.startDate.min = pageRange.min;
    elements.startDate.max = pageRange.max;
    elements.startDate.value = currentState.startDate;
  }
  if (elements.endDate) {
    elements.endDate.min = pageRange.min;
    elements.endDate.max = pageRange.max;
    elements.endDate.value = currentState.endDate;
  }
  if (elements.salesArea) {
    elements.salesArea.value = currentState.salesArea;
  }
  if (elements.plantFilter) {
    elements.plantFilter.value = currentState.plantDescription;
  }
  if (elements.outletSearch) {
    elements.outletSearch.value = currentState.outletSearch;
  }
  if (elements.paFilter) {
    elements.paFilter.value = currentState.paFilter;
  }
}

function populateAreaSelect() {
  if (!elements.salesArea) {
    return;
  }
  const options = ['<option value="ALL">All sales areas</option>'].concat(
    currentData.filters.areas.map((area) => `<option value="${escapeHtml(area)}">${escapeHtml(area)}</option>`),
  );
  elements.salesArea.innerHTML = options.join("");
  elements.salesArea.value = currentState.salesArea;
}

function populatePlantSelect() {
  if (!elements.plantFilter) {
    return;
  }
  const options = ['<option value="ALL">All depot</option>'].concat(
    (currentData.filters.plants || []).map((plant) => `<option value="${escapeHtml(plant)}">${escapeHtml(plant)}</option>`),
  );
  elements.plantFilter.innerHTML = options.join("");
  elements.plantFilter.value = currentState.plantDescription;
}

function populateOutletSelect() {
  if (!elements.outletSelect) {
    return;
  }

  const relevantOutlets = sortByName(
    currentData.filters.outlets.filter(
      (outlet) =>
        (currentState.salesArea === "ALL" || outlet.salesArea === currentState.salesArea) &&
        (!elements.plantFilter ||
          currentState.plantDescription === "ALL" ||
          plantDescription(outlet.plant) === currentState.plantDescription),
    ),
  );

  const options = ['<option value="ALL">All outlets</option>'].concat(
    relevantOutlets.map(
      (outlet) =>
        `<option value="${escapeHtml(outlet.sapCode)}">${escapeHtml(outlet.outletName)} (${escapeHtml(outlet.salesArea)})${outlet.isProjectAbhuyaday ? " - PA" : ""}</option>`,
    ),
  );

  elements.outletSelect.innerHTML = options.join("");
  const outletStillExists = relevantOutlets.some((outlet) => outlet.sapCode === currentState.outletCode);
  if (!outletStillExists) {
    currentState.outletCode = "ALL";
  }
  elements.outletSelect.value = currentState.outletCode;
}

function renderProductPills() {
  if (!elements.productPills) {
    return;
  }

  const availablePillProducts = PAGE === "historical" ? HISTORICAL_COMPARISON_PRODUCTS : currentData.meta.products;
  if (PAGE === "historical" && !["ALL", ...availablePillProducts].includes(currentState.product)) {
    currentState.product = "ALL";
  }

  const pillData = [{ key: "ALL", label: PAGE === "historical" ? "All Hist. products" : "All products", empty: false }].concat(
    availablePillProducts.map((product) => ({
      key: product,
      label: product,
      empty: currentData.meta.productStats[product].records === 0,
    })),
  );

  elements.productPills.innerHTML = pillData
    .map(
      (item) => `
        <button
          class="product-pill ${currentState.product === item.key ? "active" : ""} ${item.empty ? "is-empty" : ""}"
          data-product="${escapeHtml(item.key)}"
          type="button"
        >
          ${escapeHtml(item.label)}
        </button>
      `,
    )
    .join("");

  elements.productPills.querySelectorAll("[data-product]").forEach((button) => {
    button.addEventListener("click", () => {
      currentState.product = button.dataset.product;
      render();
    });
  });
}

function plantFilterMatches(value) {
  return !elements.plantFilter || currentState.plantDescription === "ALL" || plantDescription(value) === currentState.plantDescription;
}

function matchesFilters(row) {
  if (row.date < currentState.startDate || row.date > currentState.endDate) {
    return false;
  }
  if (currentState.salesArea !== "ALL" && row.salesArea !== currentState.salesArea) {
    return false;
  }
  if (!plantFilterMatches(row.plant)) {
    return false;
  }
  if (currentState.outletCode !== "ALL" && row.sapCode !== currentState.outletCode) {
    return false;
  }
  if (currentState.product !== "ALL" && row.product !== currentState.product) {
    return false;
  }
  if (currentState.paFilter === "project" && !row.isProjectAbhuyaday) {
    return false;
  }
  if (currentState.paFilter === "non-project" && row.isProjectAbhuyaday) {
    return false;
  }
  if (currentState.outletSearch) {
    const searchValue = currentState.outletSearch.toLowerCase();
    const haystack = `${row.outletName} ${row.sapCode}`.toLowerCase();
    if (!haystack.includes(searchValue)) {
      return false;
    }
  }
  return true;
}

function outletMatchesFilters(outlet) {
  if (currentState.salesArea !== "ALL" && outlet.salesArea !== currentState.salesArea) {
    return false;
  }
  if (!plantFilterMatches(outlet.plant)) {
    return false;
  }
  if (currentState.outletCode !== "ALL" && outlet.sapCode !== currentState.outletCode) {
    return false;
  }
  if (currentState.paFilter === "project" && !outlet.isProjectAbhuyaday) {
    return false;
  }
  if (currentState.paFilter === "non-project" && outlet.isProjectAbhuyaday) {
    return false;
  }
  if (currentState.outletSearch) {
    const searchValue = currentState.outletSearch.toLowerCase();
    const haystack = `${outlet.outletName} ${outlet.sapCode}`.toLowerCase();
    if (!haystack.includes(searchValue)) {
      return false;
    }
  }
  return true;
}

function getFilteredRows() {
  return currentData.rows.filter(matchesFilters);
}

function getFilteredOutlets() {
  return currentData.filters.outlets.filter(outletMatchesFilters);
}

function selectedDatesInRange() {
  if (PAGE === "historical" && currentData.filters?.historicalDates?.length) {
    return currentData.filters.historicalDates.filter((date) => date >= currentState.startDate && date <= currentState.endDate);
  }
  return (currentData.filters.dates || []).filter((date) => date >= currentState.startDate && date <= currentState.endDate);
}

function getTargetAnchorDate() {
  return currentState?.endDate || currentData?.meta?.defaultDate || currentData?.meta?.dateMax || "";
}

function getTargetMonthKey() {
  return getTargetAnchorDate().slice(0, 7);
}

function selectedTargetMonthDates() {
  const anchorDate = getTargetAnchorDate();
  const monthKey = getTargetMonthKey();
  if (!monthKey) {
    return [];
  }
  return (currentData.filters.dates || []).filter((date) => date.startsWith(monthKey) && date <= anchorDate);
}

function getTargetMonthLabel() {
  const anchorDate = getTargetAnchorDate();
  if (!anchorDate) {
    return "Monthly target calculation";
  }
  return `${displayMonth(anchorDate.slice(0, 7))} MTD up to ${displayDate(anchorDate)}`;
}

function matchesTargetComparisonFilters(row) {
  const targetDates = selectedTargetMonthDates();
  if (!targetDates.includes(row.date)) {
    return false;
  }
  if (currentState.salesArea !== "ALL" && row.salesArea !== currentState.salesArea) {
    return false;
  }
  if (!plantFilterMatches(row.plant)) {
    return false;
  }
  if (currentState.outletCode !== "ALL" && row.sapCode !== currentState.outletCode) {
    return false;
  }
  if (currentState.product !== "ALL" && row.product !== currentState.product) {
    return false;
  }
  if (currentState.paFilter === "project" && !row.isProjectAbhuyaday) {
    return false;
  }
  if (currentState.paFilter === "non-project" && row.isProjectAbhuyaday) {
    return false;
  }
  if (currentState.outletSearch) {
    const searchValue = currentState.outletSearch.toLowerCase();
    const haystack = `${row.outletName} ${row.sapCode}`.toLowerCase();
    if (!haystack.includes(searchValue)) {
      return false;
    }
  }
  return true;
}

function getTargetComparisonRows() {
  return currentData.rows.filter(matchesTargetComparisonFilters);
}

function summarize(filteredRows) {
  const totalSalesUnits = filteredRows.reduce((sum, row) => sum + row.salesUnits, 0);
  const totalDocuments = filteredRows.reduce((sum, row) => sum + row.documents, 0);
  const activeOutlets = new Set(filteredRows.map((row) => row.sapCode)).size;
  const activeAreas = new Set(filteredRows.map((row) => row.salesArea)).size;
  const projectRows = filteredRows.filter((row) => row.isProjectAbhuyaday);
  const projectUnits = projectRows.reduce((sum, row) => sum + row.salesUnits, 0);
  const productSales = {
    MS: filteredRows.filter((row) => row.product === "MS").reduce((sum, row) => sum + row.salesUnits, 0),
    HSD: filteredRows.filter((row) => row.product === "HSD").reduce((sum, row) => sum + row.salesUnits, 0),
    Power: filteredRows.filter((row) => row.product === "Power").reduce((sum, row) => sum + row.salesUnits, 0),
  };

  return {
    totalSalesUnits,
    totalDocuments,
    activeOutlets,
    activeAreas,
    projectUnits,
    projectShare: safeDivide(projectUnits * 100, totalSalesUnits),
    productSales,
  };
}

function buildProductSummary(filteredRows) {
  const productMap = new Map();

  currentData.meta.products.forEach((product) => {
    productMap.set(product, {
      product,
      salesUnits: 0,
      netVolume: 0,
      documents: 0,
      outlets: new Set(),
    });
  });

  filteredRows.forEach((row) => {
    if (!productMap.has(row.product)) {
      productMap.set(row.product, {
        product: row.product,
        salesUnits: 0,
        netVolume: 0,
        documents: 0,
        outlets: new Set(),
      });
    }

    const target = productMap.get(row.product);
    target.salesUnits += row.salesUnits;
    target.netVolume += row.netVolume;
    target.documents += row.documents;
    target.outlets.add(row.sapCode);
  });

  return currentData.meta.products.map((product) => {
    const item = productMap.get(product);
    return {
      product,
      salesUnits: item.salesUnits,
      netVolume: item.netVolume,
      documents: item.documents,
      outlets: item.outlets.size,
    };
  });
}

function getHistoricalSales(product, sapCode, date) {
  const monthDay = date.slice(5);
  return toNumber(currentData?.historicalLookup?.[product]?.[`${sapCode}|${monthDay}`]);
}

function getVisibleHistoricalProducts() {
  if (HISTORICAL_COMPARISON_PRODUCTS.includes(currentState.product)) {
    return [currentState.product];
  }
  return HISTORICAL_COMPARISON_PRODUCTS;
}

function getVisibleTargetProducts() {
  if (HISTORICAL_COMPARISON_PRODUCTS.includes(currentState.product)) {
    return [currentState.product];
  }
  if (TARGET_PRODUCTS.includes(currentState.product)) {
    return [currentState.product];
  }
  return HISTORICAL_COMPARISON_PRODUCTS;
}

function buildAreaProductSummary(filteredRows, filteredOutlets) {
  const selectedDates = selectedDatesInRange();
  const products = getVisibleHistoricalProducts();
  const productMaps = new Map(products.map((product) => [product, new Map()]));

  filteredOutlets.forEach((outlet) => {
    products.forEach((product) => {
      const areaMap = productMaps.get(product);
      if (!areaMap.has(outlet.salesArea)) {
        areaMap.set(outlet.salesArea, {
          salesArea: outlet.salesArea,
          product,
          salesUnits: 0,
          historicalSales: 0,
        });
      }
      const target = areaMap.get(outlet.salesArea);
      selectedDates.forEach((date) => {
        target.historicalSales += getHistoricalSales(product, outlet.sapCode, date);
      });
    });
  });

  filteredRows.forEach((row) => {
    if (!productMaps.has(row.product)) {
      return;
    }
    const areaMap = productMaps.get(row.product);
    const target =
      areaMap.get(row.salesArea) ||
      {
        salesArea: row.salesArea,
        product: row.product,
        salesUnits: 0,
        historicalSales: 0,
      };
    target.salesUnits += row.salesUnits;
    areaMap.set(row.salesArea, target);
  });

  return products
    .map((product) => ({
      product,
      areas: [...productMaps.get(product).values()]
        .map((item) => ({
          ...item,
          historicalSales: Number(item.historicalSales.toFixed(3)),
          salesUnits: Number(item.salesUnits.toFixed(3)),
          growthValue: formatGrowthValue(item.salesUnits, item.historicalSales),
        }))
        .filter((item) => item.salesUnits > 0 || item.historicalSales > 0)
        .sort((left, right) => right.salesUnits - left.salesUnits),
    }))
    .filter((group) => group.areas.length);
}

function buildAreaTargetSummary(filteredRows) {
  const targetMap = getAreaTargetMap();
  const products = getVisibleTargetProducts();
  const productMaps = new Map(products.map((product) => [product, new Map()]));
  const allowedTargetAreas = elements.plantFilter
    ? new Set(getFilteredOutlets().map((outlet) => outlet.salesArea).concat(filteredRows.map((row) => row.salesArea)))
    : null;

  targetMap.forEach((entry) => {
    if (currentState.salesArea !== "ALL" && entry.salesArea !== currentState.salesArea) {
      return;
    }
    if (allowedTargetAreas && !allowedTargetAreas.has(entry.salesArea)) {
      return;
    }
    products.forEach((product) => {
      const targetKl = entry.targets[product] || 0;
      if (!targetKl) {
        return;
      }
      productMaps.get(product).set(entry.salesArea, {
        salesArea: entry.salesArea,
        product,
        salesUnits: 0,
        targetSales: targetKl,
      });
    });
  });

  filteredRows.forEach((row) => {
    if (!productMaps.has(row.product)) {
      return;
    }
    const areaMap = productMaps.get(row.product);
    const targetSales = targetMap.get(row.salesArea)?.targets?.[row.product] || 0;
    if (!areaMap.has(row.salesArea) && !targetSales) {
      return;
    }
    const target =
      areaMap.get(row.salesArea) ||
      {
        salesArea: row.salesArea,
        product: row.product,
        salesUnits: 0,
        targetSales,
      };
    target.salesUnits += row.salesUnits;
    areaMap.set(row.salesArea, target);
  });

  return products
    .map((product) => ({
      product,
      areas: [...productMaps.get(product).values()]
        .map((item) => ({
          ...item,
          targetSales: Number(item.targetSales.toFixed(3)),
          salesUnits: Number(item.salesUnits.toFixed(3)),
          achievementValue: item.targetSales ? `${formatPercent.format(safeDivide(item.salesUnits * 100, item.targetSales))}%` : "--",
        }))
        .filter((item) => item.targetSales > 0 || item.salesUnits > 0)
        .sort((left, right) => right.targetSales - left.targetSales),
    }))
    .filter((group) => group.areas.length);
}

function formatGrowthValue(currentSales, historicalSales) {
  if (!historicalSales) {
    return currentSales > 0 ? "New" : "--";
  }
  const growth = ((currentSales - historicalSales) / historicalSales) * 100;
  return `${growth >= 0 ? "+" : ""}${formatPercent.format(growth)}%`;
}

function formatPowerConversion(msSales, powerSales) {
  if (!msSales) {
    return "--";
  }
  return `${formatPercent.format((powerSales / msSales) * 100)}%`;
}

function buildRegisterRows(filteredRows, filteredOutlets) {
  const registerMap = new Map();
  const selectedDates = selectedDatesInRange();

  selectedDates.forEach((date) => {
    filteredOutlets.forEach((outlet) => {
      const key = [date, outlet.salesArea, outlet.sapCode].join("|");
      registerMap.set(key, {
        date,
        salesArea: outlet.salesArea,
        sapCode: outlet.sapCode,
        outletName: outlet.outletName,
        msSales: 0,
        hsdSales: 0,
        powerSales: 0,
        isProjectAbhuyaday: outlet.isProjectAbhuyaday,
        isNilSale: true,
      });
    });
  });

  filteredRows.forEach((row) => {
    const key = [row.date, row.salesArea, row.sapCode].join("|");
    const existing = registerMap.get(key) || {
      date: row.date,
      salesArea: row.salesArea,
      sapCode: row.sapCode,
      outletName: row.outletName,
      msSales: 0,
      hsdSales: 0,
      powerSales: 0,
      isProjectAbhuyaday: row.isProjectAbhuyaday,
      isNilSale: true,
    };

    existing.isNilSale = false;
    if (row.product === "MS") {
      existing.msSales += row.salesUnits;
    }
    if (row.product === "HSD") {
      existing.hsdSales += row.salesUnits;
    }
    if (row.product === "Power") {
      existing.powerSales += row.salesUnits;
    }
    registerMap.set(key, existing);
  });

  return [...registerMap.values()]
    .map((row) => {
      const historicalMs = getHistoricalSales("MS", row.sapCode, row.date);
      const historicalHsd = getHistoricalSales("HSD", row.sapCode, row.date);
      const historicalPower = getHistoricalSales("Power", row.sapCode, row.date);
      return {
        ...row,
        historicalMs,
        growthMs: formatGrowthValue(row.msSales, historicalMs),
        historicalHsd,
        growthHsd: formatGrowthValue(row.hsdSales, historicalHsd),
        historicalPower,
        growthPower: formatGrowthValue(row.powerSales, historicalPower),
        powerConversion: formatPowerConversion(row.msSales, row.powerSales),
      };
    })
    .sort((left, right) => {
      const dateCompare = right.date.localeCompare(left.date);
      if (dateCompare !== 0) {
        return dateCompare;
      }
      const areaCompare = left.salesArea.localeCompare(right.salesArea);
      if (areaCompare !== 0) {
        return areaCompare;
      }
      return right.hsdSales + right.msSales + right.powerSales - (left.hsdSales + left.msSales + left.powerSales);
    });
}

function selectedDatesForMonth(monthKey) {
  return (currentData.filters.dates || []).filter(
    (date) => date.startsWith(monthKey) && date >= currentState.startDate && date <= currentState.endDate,
  );
}

function buildMonthlyRegisterRows(filteredRows, filteredOutlets) {
  const monthlyMap = new Map();
  const selectedMonths = [...new Set(selectedDatesInRange().map((date) => date.slice(0, 7)))];

  selectedMonths.forEach((month) => {
    filteredOutlets.forEach((outlet) => {
      const key = [month, outlet.salesArea, outlet.sapCode].join("|");
      monthlyMap.set(key, {
        month,
        salesArea: outlet.salesArea,
        sapCode: outlet.sapCode,
        outletName: outlet.outletName,
        msSales: 0,
        hsdSales: 0,
        powerSales: 0,
        lubeSales: 0,
        defSales: 0,
        activeDates: new Set(),
        isProjectAbhuyaday: outlet.isProjectAbhuyaday,
        isNilSale: true,
      });
    });
  });

  filteredRows.forEach((row) => {
    const month = row.date.slice(0, 7);
    const key = [month, row.salesArea, row.sapCode].join("|");
    const existing = monthlyMap.get(key) || {
      month,
      salesArea: row.salesArea,
      sapCode: row.sapCode,
      outletName: row.outletName,
      msSales: 0,
      hsdSales: 0,
      powerSales: 0,
      lubeSales: 0,
      defSales: 0,
      activeDates: new Set(),
      isProjectAbhuyaday: row.isProjectAbhuyaday,
      isNilSale: true,
    };

    existing.activeDates.add(row.date);
    existing.isNilSale = false;
    if (row.product === "MS") {
      existing.msSales += row.salesUnits;
    }
    if (row.product === "HSD") {
      existing.hsdSales += row.salesUnits;
    }
    if (row.product === "Power") {
      existing.powerSales += row.salesUnits;
    }
    if (row.product === "Lube") {
      existing.lubeSales += row.salesUnits;
    }
    if (row.product === "DEF") {
      existing.defSales += row.salesUnits;
    }
    monthlyMap.set(key, existing);
  });

  return [...monthlyMap.values()]
    .map((row) => {
      const monthDates = selectedDatesForMonth(row.month);
      const historicalMs = monthDates.reduce((sum, date) => sum + getHistoricalSales("MS", row.sapCode, date), 0);
      const historicalHsd = monthDates.reduce((sum, date) => sum + getHistoricalSales("HSD", row.sapCode, date), 0);
      const historicalPower = monthDates.reduce((sum, date) => sum + getHistoricalSales("Power", row.sapCode, date), 0);

      return {
        ...row,
        activeDays: row.activeDates.size,
        historicalMs,
        growthMs: formatGrowthValue(row.msSales, historicalMs),
        historicalHsd,
        growthHsd: formatGrowthValue(row.hsdSales, historicalHsd),
        historicalPower,
        growthPower: formatGrowthValue(row.powerSales, historicalPower),
        powerConversion: formatPowerConversion(row.msSales, row.powerSales),
      };
    })
    .sort((left, right) => {
      const monthCompare = right.month.localeCompare(left.month);
      if (monthCompare !== 0) {
        return monthCompare;
      }
      const areaCompare = left.salesArea.localeCompare(right.salesArea);
      if (areaCompare !== 0) {
        return areaCompare;
      }
      return (
        right.hsdSales +
        right.msSales +
        right.powerSales +
        right.lubeSales +
        right.defSales -
        (left.hsdSales + left.msSales + left.powerSales + left.lubeSales + left.defSales)
      );
    });
}

function buildHistoricalRows(filteredRows, filteredOutlets) {
  const products = getVisibleHistoricalProducts();
  const presentLookup = new Map();

  filteredRows.forEach((row) => {
    if (!products.includes(row.product)) {
      return;
    }
    const key = [row.date, row.sapCode, row.product].join("|");
    presentLookup.set(key, Number(((presentLookup.get(key) || 0) + row.salesUnits).toFixed(3)));
  });

  if (currentData.historicalRows?.length) {
    return currentData.historicalRows
      .filter((row) => {
        if (row.date < currentState.startDate || row.date > currentState.endDate) {
          return false;
        }
        if (!products.includes(row.product)) {
          return false;
        }
        if (currentState.salesArea !== "ALL" && row.salesArea !== currentState.salesArea) {
          return false;
        }
        if (!plantFilterMatches(row.plant)) {
          return false;
        }
        if (currentState.outletCode !== "ALL" && row.sapCode !== currentState.outletCode) {
          return false;
        }
        if (currentState.paFilter === "project" && !row.isProjectAbhuyaday) {
          return false;
        }
        if (currentState.paFilter === "non-project" && row.isProjectAbhuyaday) {
          return false;
        }
        if (currentState.outletSearch) {
          const searchValue = currentState.outletSearch.toLowerCase();
          const haystack = `${row.outletName} ${row.sapCode}`.toLowerCase();
          if (!haystack.includes(searchValue)) {
            return false;
          }
        }
        return true;
      })
      .map((row) => {
        const presentKl = presentLookup.get([row.date, row.sapCode, row.product].join("|")) || 0;
        return {
          ...row,
          presentKl,
          growthVsHistorical: formatGrowthValue(presentKl, row.historicalKl),
        };
      })
      .sort((left, right) => {
        const dateCompare = right.date.localeCompare(left.date);
        if (dateCompare !== 0) {
          return dateCompare;
        }
        const areaCompare = left.salesArea.localeCompare(right.salesArea);
        if (areaCompare !== 0) {
          return areaCompare;
        }
        const outletCompare = left.outletName.localeCompare(right.outletName);
        if (outletCompare !== 0) {
          return outletCompare;
        }
        return KNOWN_PRODUCT_ORDER.indexOf(left.product) - KNOWN_PRODUCT_ORDER.indexOf(right.product);
      });
  }

  const selectedDates = selectedDatesInRange();
  const rows = [];
  filteredOutlets.forEach((outlet) => {
    selectedDates.forEach((date) => {
      products.forEach((product) => {
        const historicalKl = getHistoricalSales(product, outlet.sapCode, date);
        if (!historicalKl) {
          return;
        }
        const presentKl = presentLookup.get([date, outlet.sapCode, product].join("|")) || 0;
        rows.push({
          date,
          salesArea: outlet.salesArea,
          sapCode: outlet.sapCode,
          outletName: outlet.outletName,
          product,
          historicalKl,
          presentKl,
          growthVsHistorical: formatGrowthValue(presentKl, historicalKl),
          isProjectAbhuyaday: outlet.isProjectAbhuyaday,
        });
      });
    });
  });

  return rows.sort((left, right) => {
    const dateCompare = right.date.localeCompare(left.date);
    if (dateCompare !== 0) {
      return dateCompare;
    }
    const areaCompare = left.salesArea.localeCompare(right.salesArea);
    if (areaCompare !== 0) {
      return areaCompare;
    }
    const outletCompare = left.outletName.localeCompare(right.outletName);
    if (outletCompare !== 0) {
      return outletCompare;
    }
    return KNOWN_PRODUCT_ORDER.indexOf(left.product) - KNOWN_PRODUCT_ORDER.indexOf(right.product);
  });
}

function buildOutletRows(filteredRows, filteredOutlets) {
  const outletMap = new Map();

  filteredOutlets.forEach((outlet) => {
    outletMap.set(outlet.sapCode, {
      sapCode: outlet.sapCode,
      outletName: outlet.outletName,
      salesArea: outlet.salesArea,
      products: new Set(),
      salesUnits: 0,
      netVolume: 0,
      dates: new Set(),
      documents: 0,
      isProjectAbhuyaday: outlet.isProjectAbhuyaday,
      isNilSale: true,
    });
  });

  filteredRows.forEach((row) => {
    const target = outletMap.get(row.sapCode) || {
      sapCode: row.sapCode,
      outletName: row.outletName,
      salesArea: row.salesArea,
      products: new Set(),
      salesUnits: 0,
      netVolume: 0,
      dates: new Set(),
      documents: 0,
      isProjectAbhuyaday: row.isProjectAbhuyaday,
      isNilSale: true,
    };

    target.products.add(row.product);
    target.salesUnits += row.salesUnits;
    target.netVolume += row.netVolume;
    target.dates.add(row.date);
    target.documents += row.documents;
    target.isNilSale = false;
    outletMap.set(row.sapCode, target);
  });

  return [...outletMap.values()]
    .map((row) => ({
      ...row,
      productList: row.products.size ? [...row.products].sort().join(", ") : "Nil sale",
      activeDays: row.dates.size,
    }))
    .sort((left, right) => {
      const salesCompare = right.salesUnits - left.salesUnits;
      if (salesCompare !== 0) {
        return salesCompare;
      }
      return left.outletName.localeCompare(right.outletName);
    });
}

function renderSummaryText(filteredRows) {
  if (!elements.activeSummary) {
    return;
  }

  const rangeText =
    currentState.startDate === currentState.endDate
      ? `for ${displayDate(currentState.startDate)}`
      : `from ${displayDate(currentState.startDate)} to ${displayDate(currentState.endDate)}`;
  const productText = currentState.product === "ALL" ? "all products" : currentState.product;
  const areaText = currentState.salesArea === "ALL" ? "all sales areas" : currentState.salesArea;
  const plantText =
    elements.plantFilter && currentState.plantDescription !== "ALL" ? `, depot ${currentState.plantDescription}` : "";
  const outletText = currentState.outletCode === "ALL" ? "all outlets" : "one selected outlet";
  if (PAGE === "historical") {
    elements.activeSummary.textContent = `${formatWhole.format(
      lastHistoricalRows.length,
    )} Hist. rows matched ${rangeText} across ${areaText}${plantText} and ${outletText}.`;
    return;
  }
  elements.activeSummary.textContent = `${filteredRows.length} transaction rows matched ${rangeText} across ${areaText}${plantText}, ${productText}, and ${outletText}.`;
}

function renderMetrics(summary) {
  if (!elements.metricsGrid) {
    return;
  }

  const cards = [
    {
      label: "Total KL",
      value: formatNumber.format(summary.totalSalesUnits),
      detail: `${formatWhole.format(summary.totalDocuments)} billing lines in the current filtered register.`,
    },
    {
      label: "MS sales",
      value: formatNumber.format(summary.productSales.MS),
      detail: "Motor spirit sales in KL for the selected view.",
    },
    {
      label: "HSD sales",
      value: formatNumber.format(summary.productSales.HSD),
      detail: "High speed diesel sales in KL for the selected view.",
    },
    {
      label: "Power sales",
      value: formatNumber.format(summary.productSales.Power),
      detail: "Power sales in KL for the selected view.",
    },
    {
      label: "Active outlets",
      value: formatWhole.format(summary.activeOutlets),
      detail: `${formatWhole.format(summary.activeAreas)} sales areas contributed to this view.`,
    },
    {
      label: "Project share",
      value: `${formatPercent.format(summary.projectShare)}%`,
      detail: `${formatNumber.format(summary.projectUnits)} KL came from Project Abhuyaday outlets.`,
    },
  ];

  elements.metricsGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="panel metric-card">
          <span class="metric-label">${escapeHtml(card.label)}</span>
          <strong class="metric-value">${escapeHtml(card.value)}</strong>
          <p class="metric-detail">${escapeHtml(card.detail)}</p>
        </article>
      `,
    )
    .join("");
}

function renderProductCards(productSummary) {
  if (!elements.productCards) {
    return;
  }

  elements.productCards.innerHTML = productSummary
    .map(
      (item) => `
        <article class="product-card">
          <h3>${escapeHtml(item.product)}</h3>
          <strong>${escapeHtml(formatNumber.format(item.salesUnits))}</strong>
          <p>${escapeHtml(
            `${formatWhole.format(item.documents)} billing lines, ${formatWhole.format(item.outlets)} outlets, ${formatNumber.format(item.salesUnits)} KL sales.`,
          )}</p>
        </article>
      `,
    )
    .join("");

  if (elements.emptyProductsNote) {
    if (currentData.meta.emptyProducts.length) {
      elements.emptyProductsNote.textContent = `No transaction rows were available in the current source for ${currentData.meta.emptyProducts.join(", ")}. They remain visible in the dashboard so future data can drop in without changing the layout.`;
    } else {
      elements.emptyProductsNote.textContent = "";
    }
  }
}

const AREA_CHART_COLORS = ["#4dbddd", "#f2ae49", "#92d9a2", "#e77474", "#8fb7ff", "#d9b36d", "#7ed4c5", "#b7a4ff"];

function getGrowthTone(item) {
  if (!item.historicalSales) {
    return item.salesUnits > 0 ? "new" : "neutral";
  }
  return item.salesUnits >= item.historicalSales ? "positive" : "negative";
}

function compactAreaSummary(areaSummary) {
  if (areaSummary.length <= 7) {
    return areaSummary;
  }

  const visible = areaSummary.slice(0, 6);
  const other = areaSummary.slice(6).reduce(
    (target, item) => {
      target.salesUnits += item.salesUnits;
      target.historicalSales += item.historicalSales;
      target.netVolume += item.netVolume;
      return target;
    },
    { salesArea: "Other areas", salesUnits: 0, historicalSales: 0, netVolume: 0 },
  );
  other.growthValue = formatGrowthValue(other.salesUnits, other.historicalSales);
  return visible.concat(other);
}

function renderDonutRing(items, valueKey, radius, strokeWidth) {
  const total = items.reduce((sum, item) => sum + item[valueKey], 0);
  if (!total) {
    return `<circle class="donut-empty-ring" cx="120" cy="120" r="${radius}" fill="none" stroke-width="${strokeWidth}"></circle>`;
  }

  let offset = 0;
  return items
    .map((item, index) => {
      const share = safeDivide(item[valueKey] * 100, total);
      const segment = `
        <circle
          class="donut-segment"
          cx="120"
          cy="120"
          r="${radius}"
          fill="none"
          pathLength="100"
          stroke="${AREA_CHART_COLORS[index % AREA_CHART_COLORS.length]}"
          stroke-width="${strokeWidth}"
          stroke-dasharray="${share} ${100 - share}"
          stroke-dashoffset="${-offset}"
        ></circle>
      `;
      offset += share;
      return segment;
    })
    .join("");
}

function renderComparisonDonutSegments(presentValue, comparisonValue, comparisonColor = "var(--accent-warm)") {
  const total = presentValue + comparisonValue;
  if (!total) {
    return `<circle class="donut-empty-ring" cx="120" cy="120" r="72" fill="none" stroke-width="20"></circle>`;
  }

  const presentShare = safeDivide(presentValue * 100, total);
  const comparisonShare = 100 - presentShare;
  return `
    <circle
      class="donut-segment"
      cx="120"
      cy="120"
      r="72"
      fill="none"
      pathLength="100"
      stroke="var(--accent)"
      stroke-width="20"
      stroke-dasharray="${presentShare} ${100 - presentShare}"
      stroke-dashoffset="0"
    ></circle>
    <circle
      class="donut-segment"
      cx="120"
      cy="120"
      r="72"
      fill="none"
      pathLength="100"
      stroke="${comparisonColor}"
      stroke-width="20"
      stroke-dasharray="${comparisonShare} ${100 - comparisonShare}"
      stroke-dashoffset="${-presentShare}"
    ></circle>
  `;
}

function combineAreaProductSummary(areaProductSummary) {
  const products = getVisibleHistoricalProducts();
  const areaMap = new Map();

  areaProductSummary.forEach((group) => {
    group.areas.forEach((item) => {
      const area =
        areaMap.get(item.salesArea) ||
        {
          salesArea: item.salesArea,
          products: Object.fromEntries(
            products.map((product) => [
              product,
              {
                product,
                salesUnits: 0,
                historicalSales: 0,
                growthValue: "--",
              },
            ]),
          ),
          totalPresent: 0,
          totalHistorical: 0,
        };

      area.products[group.product] = item;
      area.totalPresent += item.salesUnits;
      area.totalHistorical += item.historicalSales;
      areaMap.set(item.salesArea, area);
    });
  });

  return [...areaMap.values()]
    .map((area) => ({
      ...area,
      growthValue: formatGrowthValue(area.totalPresent, area.totalHistorical),
    }))
    .sort((left, right) => right.totalPresent - left.totalPresent);
}

function renderGroupedProductBars(productRows) {
  return productRows
    .map((item) => {
      const maxValue = Math.max(item.salesUnits, item.historicalSales, 1);
      return `
        <div class="area-product-row">
          <div class="area-product-row-head">
            <strong>${escapeHtml(item.product)}</strong>
            <span class="growth-pill ${getGrowthTone(item)}">${escapeHtml(item.growthValue)}</span>
          </div>
          <div class="area-product-bars">
            <div class="area-product-bar-line">
              <span>Present</span>
              <div class="area-product-track">
                <div class="area-product-fill present" style="width: ${safeDivide(item.salesUnits * 100, maxValue)}%"></div>
              </div>
              <b>${escapeHtml(formatNumber.format(item.salesUnits))}</b>
            </div>
            <div class="area-product-bar-line">
              <span>Hist.</span>
              <div class="area-product-track">
                <div class="area-product-fill historical" style="width: ${safeDivide(item.historicalSales * 100, maxValue)}%"></div>
              </div>
              <b>${escapeHtml(formatNumber.format(item.historicalSales))}</b>
            </div>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderHistoricalGauge(productRows) {
  return productRows
    .map((item) => {
      const achievement = item.historicalSales ? safeDivide(item.salesUnits * 100, item.historicalSales) : item.salesUnits > 0 ? 140 : 0;
      const gaugeValue = Math.min(Math.max(achievement, 0), 140);
      const gaugeFill = safeDivide(gaugeValue * 100, 140);
      return `
        <div class="target-gauge historical-gauge">
          <div class="target-gauge-title">
            <strong>${escapeHtml(item.product)}</strong>
            <span class="${getGrowthTone(item)}">${escapeHtml(item.growthValue)}</span>
          </div>
          <svg class="target-gauge-svg" viewBox="0 0 180 104" role="img" aria-label="${escapeHtml(
            `${item.product} present versus Hist. gauge`,
          )}">
            <path class="target-gauge-bg" d="M 24 86 A 66 66 0 0 1 156 86" pathLength="100"></path>
            <path class="target-gauge-fill historical" d="M 24 86 A 66 66 0 0 1 156 86" pathLength="100" style="stroke-dasharray: ${gaugeFill} ${
              100 - gaugeFill
            }"></path>
            <text class="target-gauge-value" x="90" y="68" text-anchor="middle">${escapeHtml(
              item.historicalSales ? `${formatPercent.format(achievement)}%` : item.salesUnits > 0 ? "New" : "--",
            )}</text>
            <text class="target-gauge-label" x="90" y="90" text-anchor="middle">vs Hist.</text>
          </svg>
          <div class="target-gauge-meta">
            <span>P ${escapeHtml(formatNumber.format(item.salesUnits))}</span>
            <span>H ${escapeHtml(formatNumber.format(item.historicalSales))}</span>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderAreaMixChart(areaProductSummary) {
  const areaSummary = combineAreaProductSummary(areaProductSummary);
  const products = getVisibleHistoricalProducts();
  const productLabel = products.join(", ");
  elements.areaBars.innerHTML = `
    <div class="area-mix-legend">
      <span><i class="ring-dot present"></i>Present KL</span>
      <span><i class="ring-dot historical"></i>Hist. KL</span>
      <span>${escapeHtml(productLabel)} ${products.length > 1 ? "are" : "is"} grouped inside each sales area.</span>
    </div>
    <div class="chart-context-strip">
      <span>Selected date: ${escapeHtml(getSelectionDateLabel())}</span>
      <span>Comparison: Present vs Hist.</span>
    </div>
    <div class="area-mix-card-grid">
      ${areaSummary
        .map((area) => {
          const productRows = products.map((product) => area.products[product]);
          return `
            <article class="area-combo-card">
              <div class="area-donut-card-head">
                <strong>${escapeHtml(area.salesArea)}</strong>
                <span class="growth-pill ${getGrowthTone({
                  salesUnits: area.totalPresent,
                  historicalSales: area.totalHistorical,
                })}">${escapeHtml(area.growthValue)}</span>
              </div>
              <p class="area-combo-total">
                Present ${escapeHtml(formatNumber.format(area.totalPresent))} KL | Hist. ${escapeHtml(
                  formatNumber.format(area.totalHistorical),
                )} KL
              </p>
              <div class="target-gauge-group historical-gauge-group">
                ${renderHistoricalGauge(productRows)}
              </div>
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

function getTargetTone(item) {
  if (!item.targetSales) {
    return item.salesUnits > 0 ? "new" : "neutral";
  }
  return item.salesUnits >= item.targetSales ? "positive" : "negative";
}

function combineAreaTargetSummary(targetSummary) {
  const products = getVisibleTargetProducts();
  const areaMap = new Map();

  targetSummary.forEach((group) => {
    group.areas.forEach((item) => {
      const area =
        areaMap.get(item.salesArea) ||
        {
          salesArea: item.salesArea,
          products: Object.fromEntries(
            products.map((product) => [
              product,
              {
                product,
                salesUnits: 0,
                targetSales: 0,
                achievementValue: "--",
              },
            ]),
          ),
          totalPresent: 0,
          totalTarget: 0,
        };

      area.products[group.product] = item;
      area.totalPresent += item.salesUnits;
      area.totalTarget += item.targetSales;
      areaMap.set(item.salesArea, area);
    });
  });

  return [...areaMap.values()]
    .map((area) => ({
      ...area,
      achievementValue: area.totalTarget ? `${formatPercent.format(safeDivide(area.totalPresent * 100, area.totalTarget))}%` : "--",
    }))
    .sort((left, right) => right.totalTarget - left.totalTarget);
}

function getTargetAchievementNumber(item) {
  return item.targetSales ? safeDivide(item.salesUnits * 100, item.targetSales) : 0;
}

function renderTargetGauge(productRows) {
  return productRows
    .map((item) => {
      const achievement = getTargetAchievementNumber(item);
      const gaugeValue = Math.min(Math.max(achievement, 0), 140);
      const gaugeFill = safeDivide(gaugeValue * 100, 140);
      return `
        <div class="target-gauge">
          <div class="target-gauge-title">
            <strong>${escapeHtml(item.product)}</strong>
            <span class="${getTargetTone(item)}">${escapeHtml(item.achievementValue)}</span>
          </div>
          <svg class="target-gauge-svg" viewBox="0 0 180 104" role="img" aria-label="${escapeHtml(
            `${item.product} present versus target gauge`,
          )}">
            <path class="target-gauge-bg" d="M 24 86 A 66 66 0 0 1 156 86" pathLength="100"></path>
            <path class="target-gauge-fill" d="M 24 86 A 66 66 0 0 1 156 86" pathLength="100" style="stroke-dasharray: ${gaugeFill} ${
              100 - gaugeFill
            }"></path>
            <text class="target-gauge-value" x="90" y="68" text-anchor="middle">${escapeHtml(item.achievementValue)}</text>
            <text class="target-gauge-label" x="90" y="90" text-anchor="middle">Achieved</text>
          </svg>
          <div class="target-gauge-meta">
            <span>P ${escapeHtml(formatNumber.format(item.salesUnits))}</span>
            <span>T ${escapeHtml(formatNumber.format(item.targetSales))}</span>
          </div>
        </div>
      `;
    })
    .join("");
}

function renderTargetCharts(targetSummary) {
  if (!elements.targetCharts) {
    return;
  }

  if (!targetSummary.length) {
    elements.targetCharts.innerHTML =
      '<div class="empty-state">No sales-area targets saved yet. Open the Master page and add MS, HSD, Power, Lube, or DEF targets.</div>';
    return;
  }

  const areaSummary = combineAreaTargetSummary(targetSummary);
  const products = getVisibleTargetProducts();
  const productLabel = products.join(", ");
  const cardsHtml = areaSummary
    .map((area) => {
      const productRows = products.map((product) => area.products[product]);
      return `
        <article class="area-combo-card target-card">
          <div class="area-donut-card-head">
            <strong>${escapeHtml(area.salesArea)}</strong>
            <span class="growth-pill ${getTargetTone({
              salesUnits: area.totalPresent,
              targetSales: area.totalTarget,
            })}">${escapeHtml(area.achievementValue)}</span>
          </div>
          <p class="area-combo-total">
            Present ${escapeHtml(formatNumber.format(area.totalPresent))} KL | Target ${escapeHtml(
              formatNumber.format(area.totalTarget),
            )} KL
          </p>
          <div class="target-gauge-group">
            ${renderTargetGauge(productRows)}
          </div>
        </article>
      `;
    })
    .join("");
  const shouldSlideTargets = PAGE === "home" && areaSummary.length > 1;
  const targetCardsHtml = shouldSlideTargets
    ? `
      <div class="target-slider-shell" aria-label="Sliding present versus target cards">
        <div class="target-gauge-card-grid target-slider-track is-sliding">
          ${cardsHtml}
          <div class="target-slider-duplicate" aria-hidden="true">${cardsHtml}</div>
        </div>
      </div>
    `
    : `<div class="target-gauge-card-grid">${cardsHtml}</div>`;

  elements.targetCharts.innerHTML = `
    <div class="area-mix-legend">
      <span><i class="ring-dot present"></i>Present KL</span>
      <span><i class="ring-dot target"></i>Target KL</span>
      <span>${escapeHtml(productLabel)} ${products.length > 1 ? "are" : "is"} grouped inside each sales area.</span>
    </div>
    <div class="chart-context-strip target-context-strip">
      <span>Target month: ${escapeHtml(getTargetMonthLabel())}</span>
      <span>Calculation: Monthly present sales vs saved area target</span>
    </div>
    ${targetCardsHtml}
  `;
}

function getSelectionDateLabel() {
  if (!currentState?.startDate || !currentState?.endDate) {
    return "Current selection";
  }
  if (currentState.startDate === currentState.endDate) {
    return displayDate(currentState.startDate);
  }
  return `${displayDate(currentState.startDate)} to ${displayDate(currentState.endDate)}`;
}

function formatEmailKl(value) {
  return `${formatNumber.format(value)} KL`;
}

function buildAreaMixEmailLines(areaProductSummary) {
  const areaSummary = combineAreaProductSummary(areaProductSummary);
  const products = getVisibleHistoricalProducts();

  if (!areaSummary.length) {
    return ["Sales area mix / Present vs Hist. by area: No matching sales area data."];
  }

  return [
    "Sales area mix / Present vs Hist. by area:",
    ...areaSummary.map((area) => {
      const productText = products
        .map((product) => {
          const item = area.products[product] || {
            product,
            salesUnits: 0,
            historicalSales: 0,
            growthValue: "--",
          };
          return `${product} P ${formatEmailKl(item.salesUnits)}, Hist. ${formatEmailKl(item.historicalSales)}, Growth ${item.growthValue}`;
        })
        .join("; ");
      return `${area.salesArea}: Present ${formatEmailKl(area.totalPresent)}, Hist. ${formatEmailKl(
        area.totalHistorical,
      )}, Growth ${area.growthValue}. ${productText}`;
    }),
  ];
}

function buildTargetEmailLines(areaTargetSummary) {
  if (!areaTargetSummary.length) {
    return ["Present vs target by area: No saved area targets found."];
  }

  const areaSummary = combineAreaTargetSummary(areaTargetSummary);
  const products = getVisibleTargetProducts();

  return [
    "Present vs target by area:",
    ...areaSummary.map((area) => {
      const productText = products
        .map((product) => {
          const item = area.products[product] || {
            product,
            salesUnits: 0,
            targetSales: 0,
            achievementValue: "--",
          };
          return `${product} P ${formatEmailKl(item.salesUnits)}, Target ${formatEmailKl(item.targetSales)}, Achieved ${
            item.achievementValue
          }`;
        })
        .join("; ");
      return `${area.salesArea}: Present ${formatEmailKl(area.totalPresent)}, Target ${formatEmailKl(
        area.totalTarget,
      )}, Achieved ${area.achievementValue}. ${productText}`;
    }),
  ];
}

function buildDashboardEmailBody() {
  const payload = getDashboardSharePayload();
  const { summary, areaProductSummary, areaTargetSummary, areaLabel, outletLabel } = payload;

  return [
    "Jabalpur Retail Region Sales Dashboard",
    `Period: ${getSelectionDateLabel()}`,
    `Target period: ${getTargetMonthLabel()}`,
    `Area: ${areaLabel}`,
    `Outlet: ${outletLabel}`,
    `Total: ${formatEmailKl(summary.totalSalesUnits)} | MS ${formatEmailKl(summary.productSales.MS)} | HSD ${formatEmailKl(
      summary.productSales.HSD,
    )} | Power ${formatEmailKl(summary.productSales.Power)}`,
    "",
    ...buildAreaMixEmailLines(areaProductSummary),
    "",
    ...buildTargetEmailLines(areaTargetSummary),
    "",
    `Generated: ${new Date().toLocaleString("en-IN")}`,
  ].join("\n");
}

function getDashboardSharePayload() {
  const filteredRows = getFilteredRows();
  const filteredOutlets = getFilteredOutlets();
  const targetRows = getTargetComparisonRows();
  const summary = summarize(filteredRows);
  const areaProductSummary = buildAreaProductSummary(filteredRows, filteredOutlets);
  const areaTargetSummary = buildAreaTargetSummary(targetRows);
  const areaLabel = currentState.salesArea === "ALL" ? "All sales areas" : currentState.salesArea;
  const outletLabel = currentState.outletCode === "ALL" ? "All outlets" : currentState.outletCode;

  return {
    filteredRows,
    filteredOutlets,
    targetRows,
    summary,
    areaProductSummary,
    areaTargetSummary,
    areaLabel,
    outletLabel,
  };
}

function reportToneClass(value) {
  if (value === "New") {
    return "new";
  }
  if (String(value).startsWith("+")) {
    return "positive";
  }
  if (String(value).startsWith("-")) {
    return "negative";
  }
  return "neutral";
}

function buildReportBar(label, value, maxValue, className) {
  return `
    <div class="report-bar-row">
      <span>${escapeHtml(label)}</span>
      <div class="report-track"><div class="report-fill ${className}" style="width:${safeDivide(value * 100, maxValue)}%"></div></div>
      <b>${escapeHtml(formatEmailKl(value))}</b>
    </div>
  `;
}

function buildReportAreaMixHtml(areaProductSummary) {
  const areaSummary = combineAreaProductSummary(areaProductSummary);
  const products = getVisibleHistoricalProducts();

  if (!areaSummary.length) {
    return '<p class="report-empty">No Sales area mix / Present vs Hist. data for this selection.</p>';
  }

  return `
    <div class="report-card-grid">
      ${areaSummary
        .map((area) => {
          const productCards = products
            .map((product) => {
              const item = area.products[product] || {
                product,
                salesUnits: 0,
                historicalSales: 0,
                growthValue: "--",
              };
              const maxValue = Math.max(item.salesUnits, item.historicalSales, 1);
              return `
                <div class="report-product-card">
                  <div class="report-product-head">
                    <strong>${escapeHtml(product)}</strong>
                    <span class="report-pill ${reportToneClass(item.growthValue)}">${escapeHtml(item.growthValue)}</span>
                  </div>
                  ${buildReportBar("Present", item.salesUnits, maxValue, "present")}
                  ${buildReportBar("Hist.", item.historicalSales, maxValue, "historical")}
                </div>
              `;
            })
            .join("");
          return `
            <article class="report-area-card">
              <div class="report-area-head">
                <strong>${escapeHtml(area.salesArea)}</strong>
                <span class="report-pill ${reportToneClass(area.growthValue)}">${escapeHtml(area.growthValue)}</span>
              </div>
              <p>Present ${escapeHtml(formatEmailKl(area.totalPresent))} | Hist. ${escapeHtml(formatEmailKl(area.totalHistorical))}</p>
              <div class="report-product-grid">${productCards}</div>
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

function buildReportTargetHtml(areaTargetSummary) {
  const areaSummary = combineAreaTargetSummary(areaTargetSummary);
  const products = getVisibleTargetProducts();

  if (!areaSummary.length) {
    return '<p class="report-empty">No Present vs target data saved for this selection.</p>';
  }

  return `
    <div class="report-card-grid">
      ${areaSummary
        .map((area) => {
          const productCards = products
            .map((product) => {
              const item = area.products[product] || {
                product,
                salesUnits: 0,
                targetSales: 0,
                achievementValue: "--",
              };
              const maxValue = Math.max(item.salesUnits, item.targetSales, 1);
              return `
                <div class="report-product-card target">
                  <div class="report-product-head">
                    <strong>${escapeHtml(product)}</strong>
                    <span class="report-pill ${getTargetTone(item)}">${escapeHtml(item.achievementValue)}</span>
                  </div>
                  ${buildReportBar("Present", item.salesUnits, maxValue, "present")}
                  ${buildReportBar("Target", item.targetSales, maxValue, "target")}
                </div>
              `;
            })
            .join("");
          return `
            <article class="report-area-card target">
              <div class="report-area-head">
                <strong>${escapeHtml(area.salesArea)}</strong>
                <span class="report-pill ${getTargetTone({
                  salesUnits: area.totalPresent,
                  targetSales: area.totalTarget,
                })}">${escapeHtml(area.achievementValue)}</span>
              </div>
              <p>Present ${escapeHtml(formatEmailKl(area.totalPresent))} | Target ${escapeHtml(formatEmailKl(area.totalTarget))}</p>
              <div class="report-product-grid">${productCards}</div>
            </article>
          `;
        })
        .join("")}
    </div>
  `;
}

function buildDashboardEmailHtml(payload, subject, textBody) {
  const { summary, areaProductSummary, areaTargetSummary, areaLabel, outletLabel } = payload;
  const mailHref = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(textBody)}`;

  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>${escapeHtml(subject)}</title>
  <style>
    body { margin: 0; background: #07101d; color: #edf4fb; font-family: Aptos, Segoe UI, Arial, sans-serif; }
    .report-controls { position: sticky; top: 0; z-index: 2; display: flex; flex-wrap: wrap; gap: 10px; padding: 12px 18px; background: rgba(7, 16, 29, 0.96); border-bottom: 1px solid #253751; }
    .report-controls button, .report-controls a { border: 1px solid #3e5476; border-radius: 999px; background: #101d2e; color: #edf4fb; cursor: pointer; padding: 10px 14px; text-decoration: none; }
    .report-controls .primary { background: linear-gradient(135deg, #36d399, #60a5fa); color: #061018; font-weight: 800; }
    .report-page { max-width: 1160px; margin: 0 auto; padding: 22px; }
    .report-hero, .report-section, .report-metric { background: #101a2a; border: 1px solid #293b57; border-radius: 18px; box-shadow: 0 12px 32px rgba(0, 0, 0, 0.24); }
    .report-hero { padding: 22px; margin-bottom: 16px; }
    .report-hero p, .report-area-card p { color: #9fb0c7; }
    h1, h2, p { margin-top: 0; }
    h1 { font-size: 30px; margin-bottom: 8px; }
    h2 { font-size: 20px; margin-bottom: 12px; }
    .report-metrics { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin: 16px 0; }
    .report-metric { padding: 14px; }
    .report-metric span { color: #9fb0c7; display: block; font-size: 12px; text-transform: uppercase; letter-spacing: 0.12em; }
    .report-metric strong { display: block; font-size: 22px; margin-top: 6px; }
    .report-section { padding: 18px; margin: 16px 0; }
    .report-section-note { color: #9fb0c7; margin-bottom: 14px; }
    .report-card-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
    .report-area-card { background: #0b1422; border: 1px solid #263953; border-radius: 16px; padding: 14px; }
    .report-area-card.target { border-color: rgba(251, 191, 36, 0.42); }
    .report-area-head, .report-product-head, .report-bar-row { display: flex; align-items: center; justify-content: space-between; gap: 10px; }
    .report-area-head strong { font-size: 18px; letter-spacing: 0.04em; }
    .report-product-grid { display: grid; gap: 10px; }
    .report-product-card { background: #07101d; border: 1px solid #21334d; border-radius: 13px; padding: 12px; }
    .report-product-card.target { border-color: rgba(251, 191, 36, 0.26); }
    .report-product-head strong { letter-spacing: 0.12em; }
    .report-pill { border-radius: 999px; padding: 6px 9px; background: #1a2a3f; color: #edf4fb; font-weight: 800; font-size: 12px; }
    .report-pill.positive { background: rgba(54, 211, 153, 0.18); color: #b7e9c2; }
    .report-pill.negative { background: rgba(231, 116, 116, 0.18); color: #f2b2b2; }
    .report-pill.new { background: rgba(96, 165, 250, 0.18); color: #bfdbfe; }
    .report-bar-row { display: grid; grid-template-columns: 72px minmax(0, 1fr) 92px; margin-top: 8px; color: #9fb0c7; font-size: 13px; }
    .report-bar-row b { color: #edf4fb; text-align: right; }
    .report-track { height: 10px; overflow: hidden; border-radius: 999px; background: rgba(143, 160, 184, 0.16); }
    .report-fill { height: 100%; min-width: 2px; border-radius: inherit; }
    .report-fill.present { background: linear-gradient(90deg, #36d399, #83e6bd); }
    .report-fill.historical { background: linear-gradient(90deg, #60a5fa, #93c5fd); }
    .report-fill.target { background: linear-gradient(90deg, #fbbf24, #fcd34d); }
    .report-empty { color: #9fb0c7; }
    @media (max-width: 780px) {
      .report-card-grid, .report-metrics { grid-template-columns: 1fr; }
      .report-page { padding: 12px; }
    }
    @media print {
      body { background: #ffffff; color: #111827; }
      .report-controls { display: none; }
      .report-hero, .report-section, .report-metric, .report-area-card, .report-product-card { box-shadow: none; background: #ffffff; color: #111827; border-color: #d1d5db; }
      .report-hero p, .report-area-card p, .report-section-note, .report-bar-row { color: #4b5563; }
      .report-bar-row b { color: #111827; }
    }
  </style>
</head>
<body>
  <div class="report-controls">
    <button class="primary" id="copy-report" type="button">Copy graphical report</button>
    <button type="button" onclick="window.print()">Print / PDF</button>
    <a href="${mailHref}">Open email text</a>
  </div>
  <main class="report-page">
    <div class="report-copy-surface" id="email-report">
      <section class="report-hero">
        <h1>Jabalpur Retail Region Sales Dashboard</h1>
        <p><strong>Period:</strong> ${escapeHtml(getSelectionDateLabel())} | <strong>Target period:</strong> ${escapeHtml(
          getTargetMonthLabel(),
        )}</p>
        <p><strong>Area:</strong> ${escapeHtml(areaLabel)} | <strong>Outlet:</strong> ${escapeHtml(outletLabel)}</p>
      </section>
      <section class="report-metrics">
        <div class="report-metric"><span>Total</span><strong>${escapeHtml(formatEmailKl(summary.totalSalesUnits))}</strong></div>
        <div class="report-metric"><span>MS</span><strong>${escapeHtml(formatEmailKl(summary.productSales.MS))}</strong></div>
        <div class="report-metric"><span>HSD</span><strong>${escapeHtml(formatEmailKl(summary.productSales.HSD))}</strong></div>
        <div class="report-metric"><span>Power</span><strong>${escapeHtml(formatEmailKl(summary.productSales.Power))}</strong></div>
      </section>
      <section class="report-section">
        <h2>Sales area mix / Present vs Hist. by area</h2>
        <p class="report-section-note">Present and Hist. values shown in KL with MS, HSD, and Power grouped by sales area.</p>
        ${buildReportAreaMixHtml(areaProductSummary)}
      </section>
      <section class="report-section">
        <h2>Present vs target by area</h2>
        <p class="report-section-note">Target calculation uses monthly present sales: ${escapeHtml(getTargetMonthLabel())}.</p>
        ${buildReportTargetHtml(areaTargetSummary)}
      </section>
    </div>
  </main>
</body>
</html>`;
}

function copyGraphicalReport(reportWindow) {
  const reportElement = reportWindow.document.querySelector("#email-report");
  if (!reportElement) {
    return;
  }

  const html = reportElement.outerHTML;
  const text = reportElement.innerText;
  if (reportWindow.navigator.clipboard && reportWindow.ClipboardItem) {
    reportWindow.navigator.clipboard
      .write([
        new reportWindow.ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([text], { type: "text/plain" }),
        }),
      ])
      .then(() => reportWindow.alert("Graphical report copied. Paste it into your email body."))
      .catch(() => reportWindow.alert("Copy failed. You can still select the report and copy it manually."));
    return;
  }

  const range = reportWindow.document.createRange();
  range.selectNode(reportElement);
  const selection = reportWindow.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
  reportWindow.document.execCommand("copy");
  selection.removeAllRanges();
  reportWindow.alert("Graphical report copied. Paste it into your email body.");
}

function openGraphicalEmailReport(html) {
  const reportWindow = window.open("", "_blank");
  if (!reportWindow) {
    return false;
  }

  reportWindow.document.open();
  reportWindow.document.write(html);
  reportWindow.document.close();
  const attachCopyButton = () => {
    const copyButton = reportWindow.document.querySelector("#copy-report");
    if (copyButton && !copyButton.dataset.bound) {
      copyButton.dataset.bound = "true";
      copyButton.addEventListener("click", () => copyGraphicalReport(reportWindow));
    }
  };
  attachCopyButton();
  reportWindow.addEventListener("load", attachCopyButton, { once: true });
  return true;
}

function shareDashboardByEmail() {
  const payload = getDashboardSharePayload();
  const subject = `Jabalpur Sales Dashboard - ${getSelectionDateLabel()}`;
  const textBody = buildDashboardEmailBody();
  const html = buildDashboardEmailHtml(payload, subject, textBody);

  if (!openGraphicalEmailReport(html)) {
    window.location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(textBody)}`;
  }
}

function renderAreaBarComparison(areaSummary) {
  const maxValue = Math.max(
    ...areaSummary.flatMap((item) => [item.salesUnits, item.historicalSales]),
    1,
  );

  const legend = `
    <div class="bar-legend">
      <span class="legend-chip present">Present KL</span>
      <span class="legend-chip historical">Hist. KL</span>
      <span class="legend-chip growth">Growth %</span>
    </div>
  `;

  const rows = areaSummary
    .map(
      (item) => `
        <div class="bar-row comparison-row">
          <div class="bar-topline">
            <div class="bar-title-block">
              <strong>${escapeHtml(item.salesArea)}</strong>
              <span class="bar-caption">Present ${escapeHtml(formatNumber.format(item.salesUnits))} KL | Hist. ${escapeHtml(
                formatNumber.format(item.historicalSales),
              )} KL</span>
            </div>
            <span class="growth-pill ${getGrowthTone(item)}">${escapeHtml(item.growthValue)}</span>
          </div>
          <div class="bar-series">
            <div class="bar-series-head">
              <span>Present</span>
              <span>${escapeHtml(formatNumber.format(item.salesUnits))} KL</span>
            </div>
            <div class="bar-track">
              <div class="bar-fill bar-fill-present" style="width: ${safeDivide(item.salesUnits * 100, maxValue)}%"></div>
            </div>
          </div>
          <div class="bar-series">
            <div class="bar-series-head">
              <span>Hist.</span>
              <span>${escapeHtml(formatNumber.format(item.historicalSales))} KL</span>
            </div>
            <div class="bar-track">
              <div class="bar-fill bar-fill-historical" style="width: ${safeDivide(item.historicalSales * 100, maxValue)}%"></div>
            </div>
          </div>
        </div>
      `,
    )
    .join("");

  elements.areaBars.innerHTML = `${legend}${rows}`;
}

function renderAreaGrowthList(areaSummary) {
  elements.areaBars.innerHTML = `
    <div class="growth-list">
      ${areaSummary
        .map(
          (item) => `
            <div class="growth-list-row">
              <strong>${escapeHtml(item.salesArea)}</strong>
              <span>${escapeHtml(formatNumber.format(item.salesUnits))} KL present</span>
              <span>${escapeHtml(formatNumber.format(item.historicalSales))} KL Hist.</span>
              <span class="growth-pill ${getGrowthTone(item)}">${escapeHtml(item.growthValue)}</span>
            </div>
          `,
        )
        .join("")}
    </div>
  `;
}

function renderAreaBars(areaProductSummary) {
  if (!elements.areaBars) {
    return;
  }

  if (!areaProductSummary.length) {
    elements.areaBars.innerHTML = '<div class="empty-state">No area totals are available for this filter selection.</div>';
    return;
  }

  renderAreaMixChart(areaProductSummary);
}

function typeBadge(isProjectAbhuyaday) {
  if (isProjectAbhuyaday) {
    return '<span class="badge project">Project Abhuyaday</span>';
  }
  return '<span class="badge standard">Standard</span>';
}

function renderRegisterTable(registerRows) {
  if (!elements.registerBody) {
    return;
  }

  if (!registerRows.length) {
    elements.registerBody.innerHTML = '<tr><td colspan="14" class="empty-state">No register rows match the current filters.</td></tr>';
    return;
  }

  elements.registerBody.innerHTML = registerRows
    .map(
      (row) => `
        <tr>
          <td><span class="table-value">${escapeHtml(displayDate(row.date))}</span></td>
          <td><span class="table-value">${escapeHtml(row.salesArea)}</span></td>
          <td>
            <span class="table-value">${escapeHtml(row.outletName)}</span>
            <span class="table-sub">SAP ${escapeHtml(row.sapCode)}${row.isNilSale ? " | Nil sale" : ""}</span>
          </td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.msSales))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.historicalMs))}</span></td>
          <td><span class="table-value">${escapeHtml(row.growthMs)}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.hsdSales))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.historicalHsd))}</span></td>
          <td><span class="table-value">${escapeHtml(row.growthHsd)}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.powerSales))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.historicalPower))}</span></td>
          <td><span class="table-value">${escapeHtml(row.growthPower)}</span></td>
          <td><span class="table-value">${escapeHtml(row.powerConversion)}</span></td>
          <td>${typeBadge(row.isProjectAbhuyaday)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderMonthlyRegisterTable(monthlyRows) {
  if (!elements.monthlyRegisterBody) {
    return;
  }

  if (!monthlyRows.length) {
    elements.monthlyRegisterBody.innerHTML = '<tr><td colspan="16" class="empty-state">No monthly register rows match the current filters.</td></tr>';
    return;
  }

  elements.monthlyRegisterBody.innerHTML = monthlyRows
    .map(
      (row) => `
        <tr>
          <td><span class="table-value">${escapeHtml(displayMonth(row.month))}</span></td>
          <td><span class="table-value">${escapeHtml(row.salesArea)}</span></td>
          <td>
            <span class="table-value">${escapeHtml(row.outletName)}</span>
            <span class="table-sub">SAP ${escapeHtml(row.sapCode)} | ${escapeHtml(formatWhole.format(row.activeDays))} active days${row.isNilSale ? " | Nil sale" : ""}</span>
          </td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.msSales))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.historicalMs))}</span></td>
          <td><span class="table-value">${escapeHtml(row.growthMs)}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.hsdSales))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.historicalHsd))}</span></td>
          <td><span class="table-value">${escapeHtml(row.growthHsd)}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.powerSales))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.historicalPower))}</span></td>
          <td><span class="table-value">${escapeHtml(row.growthPower)}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.lubeSales))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.defSales))}</span></td>
          <td><span class="table-value">${escapeHtml(row.powerConversion)}</span></td>
          <td>${typeBadge(row.isProjectAbhuyaday)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderHistoricalMetrics(historicalRows) {
  if (!elements.historicalMetrics) {
    return;
  }

  const totals = historicalRows.reduce(
    (summary, row) => {
      summary.historicalKl += row.historicalKl;
      summary.products[row.product] = (summary.products[row.product] || 0) + row.historicalKl;
      summary.outlets.add(row.sapCode);
      summary.dates.add(row.date);
      return summary;
    },
    {
      historicalKl: 0,
      products: { MS: 0, HSD: 0, Power: 0 },
      outlets: new Set(),
      dates: new Set(),
    },
  );

  const cards = [
    {
      label: "Hist. KL",
      value: formatNumber.format(totals.historicalKl),
      detail: "Total historical reference KL for the current filters.",
    },
    {
      label: "MS Hist.",
      value: formatNumber.format(totals.products.MS || 0),
      detail: "Motor spirit historical reference in KL.",
    },
    {
      label: "HSD Hist.",
      value: formatNumber.format(totals.products.HSD || 0),
      detail: "High speed diesel historical reference in KL.",
    },
    {
      label: "Power Hist.",
      value: formatNumber.format(totals.products.Power || 0),
      detail: "Power historical reference in KL.",
    },
    {
      label: "Hist. outlets",
      value: formatWhole.format(totals.outlets.size),
      detail: "Outlets with historical reference values in this view.",
    },
    {
      label: "Hist. dates",
      value: formatWhole.format(totals.dates.size),
      detail: "Historical billing dates available in the current view.",
    },
  ];

  elements.historicalMetrics.innerHTML = cards
    .map(
      (card) => `
        <article class="panel metric-card">
          <span class="metric-label">${escapeHtml(card.label)}</span>
          <strong class="metric-value">${escapeHtml(card.value)}</strong>
          <p class="metric-detail">${escapeHtml(card.detail)}</p>
        </article>
      `,
    )
    .join("");
}

function renderHistoricalTable(historicalRows) {
  if (!elements.historicalBody) {
    return;
  }

  if (elements.historicalCountNote) {
    elements.historicalCountNote.textContent = `${formatWhole.format(historicalRows.length)} Hist. rows`;
  }

  if (!historicalRows.length) {
    elements.historicalBody.innerHTML = '<tr><td colspan="6" class="empty-state">No Hist. rows match the current filters.</td></tr>';
    return;
  }

  const visibleRows = historicalRows.slice(0, 1500);
  if (elements.historicalCountNote && historicalRows.length > visibleRows.length) {
    elements.historicalCountNote.textContent = `Showing first ${formatWhole.format(visibleRows.length)} of ${formatWhole.format(
      historicalRows.length,
    )} Hist. rows. Export Excel for the full list.`;
  }

  elements.historicalBody.innerHTML = visibleRows
    .map(
      (row) => `
        <tr>
          <td><span class="table-value">${escapeHtml(displayDate(row.date))}</span></td>
          <td><span class="table-value">${escapeHtml(row.salesArea)}</span></td>
          <td>
            <span class="table-value">${escapeHtml(row.outletName)}</span>
            <span class="table-sub">SAP ${escapeHtml(row.sapCode)}</span>
          </td>
          <td><span class="table-value">${escapeHtml(row.product)}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.historicalKl))}</span></td>
          <td>${typeBadge(row.isProjectAbhuyaday)}</td>
        </tr>
      `,
    )
    .join("");
}

function buildHistoricalAreaRows(historicalRows) {
  const areaMap = new Map();

  historicalRows.forEach((row) => {
    const target =
      areaMap.get(row.salesArea) ||
      {
        salesArea: row.salesArea,
        msHistorical: 0,
        hsdHistorical: 0,
        powerHistorical: 0,
        totalHistorical: 0,
        outlets: new Set(),
        dates: new Set(),
      };

    if (row.product === "MS") {
      target.msHistorical += row.historicalKl;
    }
    if (row.product === "HSD") {
      target.hsdHistorical += row.historicalKl;
    }
    if (row.product === "Power") {
      target.powerHistorical += row.historicalKl;
    }
    target.totalHistorical += row.historicalKl;
    target.outlets.add(row.sapCode);
    target.dates.add(row.date);
    areaMap.set(row.salesArea, target);
  });

  return [...areaMap.values()]
    .map((row) => ({
      ...row,
      msHistorical: Number(row.msHistorical.toFixed(3)),
      hsdHistorical: Number(row.hsdHistorical.toFixed(3)),
      powerHistorical: Number(row.powerHistorical.toFixed(3)),
      totalHistorical: Number(row.totalHistorical.toFixed(3)),
      outletCount: row.outlets.size,
      activeDateCount: row.dates.size,
    }))
    .sort((left, right) => right.totalHistorical - left.totalHistorical);
}

function renderHistoricalAreaTable(historicalRows) {
  if (!elements.historicalAreaBody) {
    return;
  }

  const areaRows = buildHistoricalAreaRows(historicalRows);
  if (elements.historicalAreaNote) {
    elements.historicalAreaNote.textContent = `${formatWhole.format(areaRows.length)} sales areas`;
  }

  if (!areaRows.length) {
    elements.historicalAreaBody.innerHTML = '<tr><td colspan="7" class="empty-state">No sales-area wise Hist. sales match the current filters.</td></tr>';
    return;
  }

  elements.historicalAreaBody.innerHTML = areaRows
    .map(
      (row) => `
        <tr>
          <td><span class="table-value">${escapeHtml(row.salesArea)}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.msHistorical))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.hsdHistorical))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.powerHistorical))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.totalHistorical))}</span></td>
          <td><span class="table-value">${escapeHtml(formatWhole.format(row.outletCount))}</span></td>
          <td><span class="table-value">${escapeHtml(formatWhole.format(row.activeDateCount))}</span></td>
        </tr>
      `,
    )
    .join("");
}

function renderOutletTable(outletRows) {
  if (!elements.outletBody) {
    return;
  }

  if (!outletRows.length) {
    elements.outletBody.innerHTML = '<tr><td colspan="8" class="empty-state">No outlet summary rows match the current filters.</td></tr>';
    return;
  }

  elements.outletBody.innerHTML = outletRows
    .slice(0, 25)
    .map(
      (row) => `
        <tr>
          <td>
            <span class="table-value">${escapeHtml(row.outletName)}</span>
            <span class="table-sub">SAP ${escapeHtml(row.sapCode)}${row.isNilSale ? " | Nil sale" : ""}</span>
          </td>
          <td><span class="table-value">${escapeHtml(row.salesArea)}</span></td>
          <td><span class="table-value">${escapeHtml(row.productList)}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.salesUnits))}</span></td>
          <td><span class="table-value">${escapeHtml(formatNumber.format(row.netVolume))}</span></td>
          <td><span class="table-value">${escapeHtml(formatWhole.format(row.activeDays))}</span></td>
          <td><span class="table-value">${escapeHtml(formatWhole.format(row.documents))}</span></td>
          <td>${typeBadge(row.isProjectAbhuyaday)}</td>
        </tr>
      `,
    )
    .join("");
}

function renderNilSaleHomeList(outletRows) {
  if (!elements.nilSaleBody) {
    return;
  }

  const nilRows = outletRows
    .filter((row) => row.isNilSale)
    .sort((left, right) => {
      const areaCompare = left.salesArea.localeCompare(right.salesArea);
      if (areaCompare !== 0) {
        return areaCompare;
      }
      return left.outletName.localeCompare(right.outletName);
    });

  if (elements.nilSaleCount) {
    elements.nilSaleCount.textContent = `${formatWhole.format(nilRows.length)} outlet${nilRows.length === 1 ? "" : "s"}`;
  }

  if (!nilRows.length) {
    elements.nilSaleBody.innerHTML = '<tr><td colspan="5" class="empty-state">No nil-sale outlets match the current filters.</td></tr>';
    return;
  }

  elements.nilSaleBody.innerHTML = nilRows
    .map(
      (row) => `
        <tr>
          <td><span class="table-value">${escapeHtml(row.outletName)}</span></td>
          <td><span class="table-value">${escapeHtml(row.sapCode)}</span></td>
          <td><span class="table-value">${escapeHtml(row.salesArea)}</span></td>
          <td><span class="table-value">${escapeHtml(formatWhole.format(row.activeDays))}</span></td>
          <td>${typeBadge(row.isProjectAbhuyaday)}</td>
        </tr>
      `,
    )
    .join("");
}

function createCsv(rowsToExport, headers) {
  const escape = (value) => `"${String(value ?? "").replaceAll('"', '""')}"`;
  const output = [headers.map((item) => escape(item.label)).join(",")];
  rowsToExport.forEach((row) => {
    output.push(headers.map((item) => escape(row[item.key])).join(","));
  });
  return output.join("\n");
}

function downloadTextFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function excelSheetName(value) {
  return cleanText(value, "Sheet1").replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || "Sheet1";
}

function downloadRowsAsExcel(filename, sheetName, headers, rowsToExport) {
  const headerRow = headers.map((header) => `<Cell><Data ss:Type="String">${escapeXml(header.label)}</Data></Cell>`).join("");
  const dataRows = rowsToExport
    .map((row) => {
      const cells = headers
        .map((header) => {
          const value = row[header.key];
          const isNumber = typeof value === "number" && Number.isFinite(value);
          return `<Cell><Data ss:Type="${isNumber ? "Number" : "String"}">${escapeXml(value)}</Data></Cell>`;
        })
        .join("");
      return `<Row>${cells}</Row>`;
    })
    .join("");
  const workbook = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Worksheet ss:Name="${escapeXml(excelSheetName(sheetName))}">
    <Table>
      <Row>${headerRow}</Row>
      ${dataRows}
    </Table>
  </Worksheet>
</Workbook>`;
  downloadTextFile(filename, workbook, "application/vnd.ms-excel;charset=utf-8");
}

function excelCell(cell) {
  const descriptor = cell && typeof cell === "object" && Object.hasOwn(cell, "value") ? cell : { value: cell };
  const value = descriptor.value ?? "";
  const isNumber = typeof value === "number" && Number.isFinite(value);
  const type = descriptor.type || (isNumber ? "Number" : "String");
  const style = descriptor.styleId ? ` ss:StyleID="${escapeXml(descriptor.styleId)}"` : "";
  const merge = descriptor.mergeAcross ? ` ss:MergeAcross="${descriptor.mergeAcross}"` : "";
  return `<Cell${style}${merge}><Data ss:Type="${type}">${escapeXml(value)}</Data></Cell>`;
}

function downloadStyledWorkbook(filename, worksheets) {
  const worksheetXml = worksheets
    .map((worksheet) => {
      const headers = worksheet.headers || [];
      const columns = headers.map((header) => `<Column ss:Width="${header.width || 110}"/>`).join("");
      const headerRow = `<Row ss:Height="24">${headers
        .map((header) => excelCell({ value: header.label, styleId: "Header" }))
        .join("")}</Row>`;
      const bodyRows = (worksheet.rows || [])
        .map((row) => `<Row>${row.map((cell) => excelCell(cell)).join("")}</Row>`)
        .join("");
      const titleMerge = Math.max(headers.length - 1, 0);
      return `
  <Worksheet ss:Name="${escapeXml(excelSheetName(worksheet.name))}">
    <Table>
      ${columns}
      <Row ss:Height="30">${excelCell({ value: worksheet.title, styleId: "Title", mergeAcross: titleMerge })}</Row>
      <Row ss:Height="22">${excelCell({ value: worksheet.subtitle || "", styleId: "Subtitle", mergeAcross: titleMerge })}</Row>
      ${worksheet.note ? `<Row>${excelCell({ value: worksheet.note, styleId: "Note", mergeAcross: titleMerge })}</Row>` : ""}
      <Row></Row>
      ${headerRow}
      ${bodyRows}
    </Table>
    <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
      <FreezePanes/>
      <FrozenNoSplit/>
      <SplitHorizontal>5</SplitHorizontal>
      <TopRowBottomPane>5</TopRowBottomPane>
      <ActivePane>2</ActivePane>
    </WorksheetOptions>
  </Worksheet>`;
    })
    .join("");

  const workbook = `<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Styles>
    <Style ss:ID="Default" ss:Name="Normal"><Font ss:FontName="Calibri" ss:Size="11"/></Style>
    <Style ss:ID="Title"><Font ss:Bold="1" ss:Size="16" ss:Color="#FFFFFF"/><Interior ss:Color="#0B1220" ss:Pattern="Solid"/><Alignment ss:Vertical="Center"/></Style>
    <Style ss:ID="Subtitle"><Font ss:Color="#D8E2F1"/><Interior ss:Color="#162338" ss:Pattern="Solid"/></Style>
    <Style ss:ID="Note"><Font ss:Italic="1" ss:Color="#5B6B82"/></Style>
    <Style ss:ID="Header"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#1F5F8B" ss:Pattern="Solid"/><Borders><Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#8DB9D8"/></Borders></Style>
    <Style ss:ID="Total"><Font ss:Bold="1" ss:Color="#FFFFFF"/><Interior ss:Color="#26465F" ss:Pattern="Solid"/></Style>
    <Style ss:ID="Number"><NumberFormat ss:Format="0.000"/></Style>
    <Style ss:ID="Positive"><Font ss:Bold="1" ss:Color="#096B3B"/><Interior ss:Color="#DDF7EA" ss:Pattern="Solid"/></Style>
    <Style ss:ID="Negative"><Font ss:Bold="1" ss:Color="#9F1239"/><Interior ss:Color="#FFE4E6" ss:Pattern="Solid"/></Style>
    <Style ss:ID="Neutral"><Font ss:Bold="1" ss:Color="#475569"/><Interior ss:Color="#E2E8F0" ss:Pattern="Solid"/></Style>
  </Styles>
  ${worksheetXml}
</Workbook>`;
  downloadTextFile(filename, workbook, "application/vnd.ms-excel;charset=utf-8");
}

function downloadTableAsExcel(tableSelector, filename, sheetName) {
  const target = document.querySelector(tableSelector);
  const table = target?.tagName === "TABLE" ? target : target?.closest("table");
  if (!table) {
    return;
  }
  const rows = Array.from(table.querySelectorAll("tr")).map((row) =>
    Array.from(row.children).map((cell) => cell.innerText.replace(/\s+/g, " ").trim()),
  );
  const headers = rows[0].map((label, index) => ({ key: `c${index}`, label }));
  const bodyRows = rows.slice(1).map((row) => Object.fromEntries(headers.map((header, index) => [header.key, row[index] || ""])));
  downloadRowsAsExcel(filename, sheetName, headers, bodyRows);
}

function roundedKl(value) {
  return Number((Number(value) || 0).toFixed(3));
}

function comparisonCellStyle(present, reference) {
  if (!reference) {
    return present > 0 ? "Positive" : "Neutral";
  }
  return present >= reference ? "Positive" : "Negative";
}

function workbookContextSubtitle(extra = "") {
  const parts = [`Period: ${getSelectionDateLabel()}`];
  if (extra) {
    parts.push(extra);
  }
  parts.push(`Area: ${currentState.salesArea === "ALL" ? "All sales areas" : currentState.salesArea}`);
  if (elements.plantFilter) {
    parts.push(`Depot: ${currentState.plantDescription === "ALL" ? "All depot" : currentState.plantDescription}`);
  }
  parts.push(`Outlet: ${currentState.outletCode === "ALL" ? "All outlets" : currentState.outletCode}`);
  parts.push(`Generated: ${new Date().toLocaleString("en-IN")}`);
  return parts.join(" | ");
}

function chartExportPrefix() {
  return PAGE === "monthly" ? "monthly-" : "";
}

function buildAreaMixWorkbookSheet() {
  const areaSummary = combineAreaProductSummary(buildAreaProductSummary(getFilteredRows(), getFilteredOutlets()));
  const products = getVisibleHistoricalProducts();
  const headers = [
    { key: "salesArea", label: "Sales Area", width: 150 },
    { key: "product", label: "Product", width: 95 },
    { key: "present", label: "Present KL", width: 105 },
    { key: "historical", label: "Hist. KL", width: 105 },
    { key: "growth", label: "Growth vs Hist.", width: 120 },
    { key: "remark", label: "Remark", width: 160 },
  ];
  const rows = [];

  if (!areaSummary.length) {
    rows.push([{ value: "No Present vs Hist. data is available for the current filters.", styleId: "Neutral", mergeAcross: headers.length - 1 }]);
  }

  areaSummary.forEach((area) => {
    rows.push([
      { value: area.salesArea, styleId: "Total" },
      { value: "TOTAL", styleId: "Total" },
      { value: roundedKl(area.totalPresent), styleId: "Total" },
      { value: roundedKl(area.totalHistorical), styleId: "Total" },
      { value: area.growthValue, styleId: comparisonCellStyle(area.totalPresent, area.totalHistorical) },
      { value: "Area total", styleId: "Total" },
    ]);

    products.forEach((product) => {
      const item = area.products[product] || {
        product,
        salesUnits: 0,
        historicalSales: 0,
        growthValue: "--",
      };
      rows.push([
        { value: area.salesArea },
        { value: product },
        { value: roundedKl(item.salesUnits), styleId: "Number" },
        { value: roundedKl(item.historicalSales), styleId: "Number" },
        { value: item.growthValue, styleId: comparisonCellStyle(item.salesUnits, item.historicalSales) },
        { value: item.historicalSales ? "Compared with Hist." : item.salesUnits > 0 ? "New sale vs Hist." : "No sale" },
      ]);
    });
  });

  return {
    name: "Present vs Hist",
    title: "Sales area mix - Present vs Hist. by area",
    subtitle: workbookContextSubtitle(),
    note: "Present and Hist. values are shown in KL. The TOTAL row gives each sales area's combined MS, HSD, and Power view.",
    headers,
    rows,
  };
}

function buildTargetWorkbookSheet() {
  const targetSummary = combineAreaTargetSummary(buildAreaTargetSummary(getTargetComparisonRows()));
  const products = getVisibleTargetProducts();
  const headers = [
    { key: "salesArea", label: "Sales Area", width: 150 },
    { key: "product", label: "Product", width: 95 },
    { key: "present", label: "Present KL", width: 105 },
    { key: "target", label: "Target KL", width: 105 },
    { key: "achievement", label: "Achievement", width: 120 },
    { key: "remark", label: "Remark", width: 170 },
  ];
  const rows = [];

  if (!targetSummary.length) {
    rows.push([{ value: "No Present vs target data is available. Add month-wise targets in the Master page.", styleId: "Neutral", mergeAcross: headers.length - 1 }]);
  }

  targetSummary.forEach((area) => {
    rows.push([
      { value: area.salesArea, styleId: "Total" },
      { value: "TOTAL", styleId: "Total" },
      { value: roundedKl(area.totalPresent), styleId: "Total" },
      { value: roundedKl(area.totalTarget), styleId: "Total" },
      { value: area.achievementValue, styleId: comparisonCellStyle(area.totalPresent, area.totalTarget) },
      { value: "Area target total", styleId: "Total" },
    ]);

    products.forEach((product) => {
      const item = area.products[product] || {
        product,
        salesUnits: 0,
        targetSales: 0,
        achievementValue: "--",
      };
      rows.push([
        { value: area.salesArea },
        { value: product },
        { value: roundedKl(item.salesUnits), styleId: "Number" },
        { value: roundedKl(item.targetSales), styleId: "Number" },
        { value: item.achievementValue, styleId: comparisonCellStyle(item.salesUnits, item.targetSales) },
        { value: item.targetSales ? "Compared with target" : item.salesUnits > 0 ? "No saved target" : "No sale/target" },
      ]);
    });
  });

  return {
    name: "Present vs Target",
    title: "Sales targets - Present vs target by area",
    subtitle: workbookContextSubtitle(`Target period: ${getTargetMonthLabel()}`),
    note: "Present values use the monthly target calculation period. Target values come from the Master page month-wise targets.",
    headers,
    rows,
  };
}

function exportChartDataExcel(kind) {
  if (kind === "target") {
    downloadStyledWorkbook(`${chartExportPrefix()}present-vs-target-by-area.xls`, [buildTargetWorkbookSheet()]);
    return;
  }
  downloadStyledWorkbook(`${chartExportPrefix()}present-vs-hist-by-area.xls`, [buildAreaMixWorkbookSheet()]);
}

function collectExportCss() {
  return Array.from(document.styleSheets)
    .map((sheet) => {
      try {
        return Array.from(sheet.cssRules)
          .map((rule) => rule.cssText)
          .join("\n");
      } catch {
        return "";
      }
    })
    .join("\n");
}

async function downloadElementImage(selector, filename, mimeType) {
  const source = document.querySelector(selector);
  if (!source) {
    return;
  }

  const rect = source.getBoundingClientRect();
  const width = Math.max(900, Math.ceil(rect.width));
  const height = Math.max(420, Math.ceil(rect.height));
  const clone = source.cloneNode(true);
  clone.style.width = `${width}px`;
  clone.style.margin = "0";
  clone.style.background = "#070b13";

  const wrapper = document.createElement("div");
  wrapper.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
  wrapper.style.width = `${width}px`;
  wrapper.style.minHeight = `${height}px`;
  wrapper.style.padding = "18px";
  wrapper.style.background = "#070b13";

  const style = document.createElement("style");
  style.textContent = collectExportCss();
  wrapper.append(style, clone);

  const serialized = new XMLSerializer().serializeToString(wrapper);
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="${width + 36}" height="${height + 36}">
      <foreignObject width="100%" height="100%">${serialized}</foreignObject>
    </svg>
  `;
  const svgUrl = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml;charset=utf-8" }));
  const image = new Image();
  const loaded = new Promise((resolve, reject) => {
    image.onload = resolve;
    image.onerror = reject;
  });
  image.src = svgUrl;
  await loaded;

  const canvas = document.createElement("canvas");
  const scale = 2;
  canvas.width = (width + 36) * scale;
  canvas.height = (height + 36) * scale;
  const context = canvas.getContext("2d");
  context.fillStyle = "#070b13";
  context.fillRect(0, 0, canvas.width, canvas.height);
  context.scale(scale, scale);
  context.drawImage(image, 0, 0);
  URL.revokeObjectURL(svgUrl);

  const extension = mimeType === "image/jpeg" ? "jpg" : "png";
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, mimeType, 0.92));
  if (!blob) {
    throw new Error("Chart image could not be generated.");
  }
  const imageUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = imageUrl;
  anchor.download = `${filename}.${extension}`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(imageUrl), 1000);
}

function drawRoundedRect(context, x, y, width, height, radius) {
  context.beginPath();
  context.moveTo(x + radius, y);
  context.lineTo(x + width - radius, y);
  context.quadraticCurveTo(x + width, y, x + width, y + radius);
  context.lineTo(x + width, y + height - radius);
  context.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
  context.lineTo(x + radius, y + height);
  context.quadraticCurveTo(x, y + height, x, y + height - radius);
  context.lineTo(x, y + radius);
  context.quadraticCurveTo(x, y, x + radius, y);
  context.closePath();
}

function drawCanvasGauge(context, x, y, width, item, mode) {
  const centerX = x + width / 2;
  const centerY = y + 92;
  const radius = 52;
  const isTarget = mode === "target";
  const denominator = isTarget ? item.targetSales : item.historicalSales;
  const numerator = item.salesUnits;
  const percent = denominator ? safeDivide(numerator * 100, denominator) : numerator > 0 ? 140 : 0;
  const capped = Math.min(Math.max(percent, 0), 140);
  const endAngle = Math.PI + (Math.PI * capped) / 140;
  const valueText = isTarget
    ? item.achievementValue
    : denominator
      ? `${formatPercent.format(percent)}%`
      : numerator > 0
        ? "New"
        : "--";
  const accent = isTarget ? "#fbbf24" : "#60a5fa";

  context.fillStyle = "#edf4fb";
  context.font = "700 18px Segoe UI";
  context.fillText(item.product, x, y + 18);
  context.fillStyle = isTarget ? (getTargetTone(item) === "positive" ? "#b7e9c2" : "#f2b2b2") : getGrowthTone(item) === "positive" ? "#b7e9c2" : "#f2b2b2";
  context.textAlign = "right";
  context.fillText(isTarget ? item.achievementValue : item.growthValue, x + width, y + 18);
  context.textAlign = "left";

  context.lineWidth = 16;
  context.lineCap = "round";
  context.strokeStyle = "rgba(142, 160, 184, 0.22)";
  context.beginPath();
  context.arc(centerX, centerY, radius, Math.PI, 0);
  context.stroke();
  context.strokeStyle = accent;
  context.beginPath();
  context.arc(centerX, centerY, radius, Math.PI, endAngle);
  context.stroke();

  context.fillStyle = "#edf4fb";
  context.font = "800 25px Segoe UI";
  context.textAlign = "center";
  context.fillText(valueText, centerX, centerY - 12);
  context.fillStyle = "#8ea0b8";
  context.font = "12px Segoe UI";
  context.fillText(isTarget ? "Achieved" : "vs Hist.", centerX, centerY + 12);
  context.textAlign = "left";

  context.fillStyle = "#8ea0b8";
  context.font = "13px Segoe UI";
  context.fillText(`P ${formatNumber.format(numerator)}`, x, y + 160);
  context.textAlign = "right";
  context.fillText(`${isTarget ? "T" : "H"} ${formatNumber.format(denominator || 0)}`, x + width, y + 160);
  context.textAlign = "left";
}

function downloadCanvasImage(canvas, filename, mimeType) {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) {
          reject(new Error("Chart image could not be generated."));
          return;
        }
        const extension = mimeType === "image/jpeg" ? "jpg" : "png";
        const url = URL.createObjectURL(blob);
        const anchor = document.createElement("a");
        anchor.href = url;
        anchor.download = `${filename}.${extension}`;
        document.body.append(anchor);
        anchor.click();
        anchor.remove();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
        resolve();
      },
      mimeType,
      0.92,
    );
  });
}

async function downloadChartImage(kind, mimeType) {
  const isTarget = kind === "target";
  const filteredRows = getFilteredRows();
  const filteredOutlets = getFilteredOutlets();
  const sourceSummary = isTarget
    ? combineAreaTargetSummary(buildAreaTargetSummary(getTargetComparisonRows()))
    : combineAreaProductSummary(buildAreaProductSummary(filteredRows, filteredOutlets));
  const products = isTarget ? getVisibleTargetProducts() : getVisibleHistoricalProducts();
  const title = isTarget ? "Present vs target by area" : "Present vs Hist. by area";
  const subtitle = isTarget ? `Target calculation: ${getTargetMonthLabel()}` : `Period: ${getSelectionDateLabel()}`;

  const width = 1600;
  const padding = 48;
  const gap = 24;
  const columns = 2;
  const cardWidth = (width - padding * 2 - gap) / columns;
  const cardHeight = 286;
  const rows = Math.max(1, Math.ceil(sourceSummary.length / columns));
  const height = padding + 112 + rows * cardHeight + Math.max(0, rows - 1) * gap + padding;
  const canvas = document.createElement("canvas");
  const scale = 2;
  canvas.width = width * scale;
  canvas.height = height * scale;
  const context = canvas.getContext("2d");
  context.scale(scale, scale);

  context.fillStyle = "#070b13";
  context.fillRect(0, 0, width, height);
  context.fillStyle = "#edf4fb";
  context.font = "800 38px Segoe UI";
  context.fillText(title, padding, padding + 38);
  context.fillStyle = "#8ea0b8";
  context.font = "18px Segoe UI";
  context.fillText(subtitle, padding, padding + 72);
  context.fillText("Jabalpur Retail Region Sales Dashboard", padding, padding + 98);

  if (!sourceSummary.length) {
    context.fillStyle = "#8ea0b8";
    context.font = "22px Segoe UI";
    context.fillText("No chart data available for the current selection.", padding, padding + 150);
    await downloadCanvasImage(canvas, isTarget ? "present-vs-target-by-area" : "present-vs-hist-by-area", mimeType);
    return;
  }

  sourceSummary.forEach((area, index) => {
    const column = index % columns;
    const row = Math.floor(index / columns);
    const x = padding + column * (cardWidth + gap);
    const y = padding + 126 + row * (cardHeight + gap);
    drawRoundedRect(context, x, y, cardWidth, cardHeight, 20);
    context.fillStyle = isTarget ? "rgba(251, 191, 36, 0.08)" : "rgba(54, 211, 153, 0.08)";
    context.fill();
    context.strokeStyle = "rgba(122, 162, 194, 0.24)";
    context.lineWidth = 1;
    context.stroke();

    context.fillStyle = "#edf4fb";
    context.font = "800 24px Segoe UI";
    context.fillText(area.salesArea, x + 22, y + 34);
    context.textAlign = "right";
    context.fillStyle = "#fbbf24";
    context.fillText(isTarget ? area.achievementValue : area.growthValue, x + cardWidth - 22, y + 34);
    context.textAlign = "left";
    context.fillStyle = "#8ea0b8";
    context.font = "15px Segoe UI";
    const comparisonLabel = isTarget
      ? `Present ${formatNumber.format(area.totalPresent)} KL | Target ${formatNumber.format(area.totalTarget)} KL`
      : `Present ${formatNumber.format(area.totalPresent)} KL | Hist. ${formatNumber.format(area.totalHistorical)} KL`;
    context.fillText(comparisonLabel, x + 22, y + 62);

    const gaugeGap = 18;
    const gaugeWidth = (cardWidth - 44 - gaugeGap * (products.length - 1)) / products.length;
    products.forEach((product, productIndex) => {
      const item = area.products[product] || {
        product,
        salesUnits: 0,
        historicalSales: 0,
        targetSales: 0,
        growthValue: "--",
        achievementValue: "--",
      };
      drawCanvasGauge(context, x + 22 + productIndex * (gaugeWidth + gaugeGap), y + 86, gaugeWidth, item, kind);
    });
  });

  await downloadCanvasImage(canvas, isTarget ? "present-vs-target-by-area" : "present-vs-hist-by-area", mimeType);
}

function wireExports() {
  const monthlyRegisterExportRows = () =>
    lastMonthlyRegisterRows.map((row) => ({
      month: row.month,
      salesArea: row.salesArea,
      outletName: row.outletName,
      sapCode: row.sapCode,
      msKl: row.msSales,
      msHistoricalKl: row.historicalMs,
      msGrowthVsHistorical: row.growthMs,
      hsdKl: row.hsdSales,
      hsdHistoricalKl: row.historicalHsd,
      hsdGrowthVsHistorical: row.growthHsd,
      powerKl: row.powerSales,
      powerHistoricalKl: row.historicalPower,
      powerGrowthVsHistorical: row.growthPower,
      lubeKl: row.lubeSales,
      defKl: row.defSales,
      powerConversion: row.powerConversion,
      activeDays: row.activeDays,
      saleStatus: row.isNilSale ? "Nil sale" : "Sale",
      outletType: row.isProjectAbhuyaday ? "Project Abhuyaday" : "Standard",
    }));
  const monthlyRegisterExportColumns = [
    { key: "month", label: "Month" },
    { key: "salesArea", label: "Sales Area" },
    { key: "outletName", label: "Outlet Name" },
    { key: "sapCode", label: "SAP Code" },
    { key: "msKl", label: "MS KL" },
    { key: "msHistoricalKl", label: "MS Hist." },
    { key: "msGrowthVsHistorical", label: "MS Growth vs Hist." },
    { key: "hsdKl", label: "HSD KL" },
    { key: "hsdHistoricalKl", label: "HSD Hist." },
    { key: "hsdGrowthVsHistorical", label: "HSD Growth vs Hist." },
    { key: "powerKl", label: "Power KL" },
    { key: "powerHistoricalKl", label: "Power Hist." },
    { key: "powerGrowthVsHistorical", label: "Power Growth vs Hist." },
    { key: "lubeKl", label: "Lube KL" },
    { key: "defKl", label: "DEF KL" },
    { key: "powerConversion", label: "Power Conversion" },
    { key: "activeDays", label: "Active Days" },
    { key: "saleStatus", label: "Sale Status" },
    { key: "outletType", label: "Outlet Type" },
  ];
  const historicalExportRows = () =>
    lastHistoricalRows.map((row) => ({
      date: row.date,
      salesArea: row.salesArea,
      outletName: row.outletName,
      sapCode: row.sapCode,
      product: row.product,
      historicalKl: row.historicalKl,
      outletType: row.isProjectAbhuyaday ? "Project Abhuyaday" : "Standard",
    }));
  const historicalExportColumns = [
    { key: "date", label: "Reference Date" },
    { key: "salesArea", label: "Sales Area" },
    { key: "outletName", label: "Outlet Name" },
    { key: "sapCode", label: "SAP Code" },
    { key: "product", label: "Product" },
    { key: "historicalKl", label: "Hist. KL" },
    { key: "outletType", label: "Outlet Type" },
  ];
  const outletSummaryExportRows = () =>
    lastOutletRows.map((row) => ({
      outletName: row.outletName,
      sapCode: row.sapCode,
      salesArea: row.salesArea,
      products: row.productList,
      salesKl: row.salesUnits,
      netVolume: row.netVolume,
      activeDays: row.activeDays,
      documents: row.documents,
      saleStatus: row.isNilSale ? "Nil sale" : "Sale",
      outletType: row.isProjectAbhuyaday ? "Project Abhuyaday" : "Standard",
    }));
  const outletSummaryExportColumns = [
    { key: "outletName", label: "Outlet Name" },
    { key: "sapCode", label: "SAP Code" },
    { key: "salesArea", label: "Sales Area" },
    { key: "products", label: "Products" },
    { key: "salesKl", label: "Total KL" },
    { key: "netVolume", label: "Net Volume" },
    { key: "activeDays", label: "Active Days" },
    { key: "documents", label: "Documents" },
    { key: "saleStatus", label: "Sale Status" },
    { key: "outletType", label: "Outlet Type" },
  ];

  if (elements.exportRegister) {
    elements.exportRegister.onclick = () => {
      const csvRows = lastRegisterRows.map((row) => ({
        date: row.date,
        salesArea: row.salesArea,
        outletName: row.outletName,
        sapCode: row.sapCode,
        msKl: row.msSales,
        msHistoricalKl: row.historicalMs,
        msGrowthVsHistorical: row.growthMs,
        hsdKl: row.hsdSales,
        hsdHistoricalKl: row.historicalHsd,
        hsdGrowthVsHistorical: row.growthHsd,
        powerKl: row.powerSales,
        powerHistoricalKl: row.historicalPower,
        powerGrowthVsHistorical: row.growthPower,
        powerConversion: row.powerConversion,
        saleStatus: row.isNilSale ? "Nil sale" : "Sale",
        outletType: row.isProjectAbhuyaday ? "Project Abhuyaday" : "Standard",
      }));
      const csv = createCsv(csvRows, [
        { key: "date", label: "Date" },
        { key: "salesArea", label: "Sales Area" },
        { key: "outletName", label: "Outlet Name" },
        { key: "sapCode", label: "SAP Code" },
        { key: "msKl", label: "MS KL" },
        { key: "msHistoricalKl", label: "MS Hist." },
        { key: "msGrowthVsHistorical", label: "MS Growth vs Hist." },
        { key: "hsdKl", label: "HSD KL" },
        { key: "hsdHistoricalKl", label: "HSD Hist." },
        { key: "hsdGrowthVsHistorical", label: "HSD Growth vs Hist." },
        { key: "powerKl", label: "Power KL" },
        { key: "powerHistoricalKl", label: "Power Hist." },
        { key: "powerGrowthVsHistorical", label: "Power Growth vs Hist." },
        { key: "powerConversion", label: "Power Conversion" },
        { key: "saleStatus", label: "Sale Status" },
        { key: "outletType", label: "Outlet Type" },
      ]);
      downloadTextFile("daily-sales-register.csv", csv, "text/csv;charset=utf-8");
    };
  }

  if (elements.exportRegisterExcel) {
    elements.exportRegisterExcel.onclick = () => {
      downloadTableAsExcel("#register-body", "daily-sales-outlet.xls", "Daily Sales Outlet");
    };
  }

  if (elements.exportMonthlyRegister) {
    elements.exportMonthlyRegister.onclick = () => {
      const csv = createCsv(monthlyRegisterExportRows(), monthlyRegisterExportColumns);
      downloadTextFile("monthly-sales-register.csv", csv, "text/csv;charset=utf-8");
    };
  }

  if (elements.exportMonthlyRegisterExcel) {
    elements.exportMonthlyRegisterExcel.onclick = () => {
      downloadRowsAsExcel(
        "monthly-sales-register.xls",
        "Monthly Register",
        monthlyRegisterExportColumns,
        monthlyRegisterExportRows(),
      );
    };
  }

  if (elements.exportHistoricalExcel) {
    elements.exportHistoricalExcel.onclick = () => {
      downloadRowsAsExcel("historical-sales-data.xls", "Hist Data", historicalExportColumns, historicalExportRows());
    };
  }

  if (elements.exportSummary) {
    elements.exportSummary.onclick = () => {
      const filename = PAGE === "monthly" ? "monthly-best-performance-outlets.csv" : "daily-best-performance-outlets.csv";
      const csv = createCsv(outletSummaryExportRows(), outletSummaryExportColumns);
      downloadTextFile(filename, csv, "text/csv;charset=utf-8");
    };
  }

  if (elements.exportSummaryExcel) {
    elements.exportSummaryExcel.onclick = () => {
      const filename = PAGE === "monthly" ? "monthly-best-performance-outlets.xls" : "daily-best-performance-outlets.xls";
      downloadRowsAsExcel(filename, "Best Performance", outletSummaryExportColumns, outletSummaryExportRows());
    };
  }

  if (elements.printView) {
    elements.printView.onclick = () => {
      window.print();
    };
  }

  if (elements.exportOutletMasterExcel) {
    elements.exportOutletMasterExcel.onclick = () => {
      const rows = sortByName(currentData.filters.outlets || []).map((outlet) => ({
        sapCode: outlet.sapCode,
        outletName: outlet.outletName,
        salesArea: outlet.salesArea,
        appointedOn: outlet.appointedOn ? displayDate(outlet.appointedOn) : "",
        outletType: outlet.outletType || (outlet.isProjectAbhuyaday ? "Project Abhuyaday" : "Standard"),
        plant: outlet.plant || "",
        notes: outlet.notes || "",
      }));
      downloadRowsAsExcel(
        "retail-outlet-master.xls",
        "Outlet Master",
        [
          { key: "sapCode", label: "SAP Code" },
          { key: "outletName", label: "Outlet" },
          { key: "salesArea", label: "Sales Area" },
          { key: "appointedOn", label: "Appointed On" },
          { key: "outletType", label: "Type" },
          { key: "plant", label: "Plant" },
          { key: "notes", label: "Notes" },
        ],
        rows,
      );
    };
  }

  if (elements.exportTargetMasterExcel) {
    elements.exportTargetMasterExcel.onclick = () => {
      const rows = getAreaTargetEntries().map((entry) => ({
        month: displayMonthKey(entry.monthKey),
        salesArea: entry.salesArea,
        ms: entry.targets.MS,
        hsd: entry.targets.HSD,
        power: entry.targets.Power,
        lube: entry.targets.Lube,
        def: entry.targets.DEF,
        note: entry.note || "",
      }));
      downloadRowsAsExcel(
        "month-wise-sales-targets.xls",
        "Month Targets",
        [
          { key: "month", label: "Month" },
          { key: "salesArea", label: "Sales Area" },
          { key: "ms", label: "MS Target KL" },
          { key: "hsd", label: "HSD Target KL" },
          { key: "power", label: "Power Target KL" },
          { key: "lube", label: "Lube Target KL" },
          { key: "def", label: "DEF Target KL" },
          { key: "note", label: "Note" },
        ],
        rows,
      );
    };
  }

  if (elements.downloadAreaPng) {
    elements.downloadAreaPng.onclick = () =>
      downloadChartImage("area", "image/png").catch((error) =>
        window.alert(error instanceof Error ? error.message : "Chart download failed."),
      );
  }
  if (elements.downloadAreaJpeg) {
    elements.downloadAreaJpeg.onclick = () =>
      downloadChartImage("area", "image/jpeg").catch((error) =>
        window.alert(error instanceof Error ? error.message : "Chart download failed."),
      );
  }
  if (elements.exportAreaExcel) {
    elements.exportAreaExcel.onclick = () => {
      exportChartDataExcel("area");
    };
  }
  if (elements.downloadTargetPng) {
    elements.downloadTargetPng.onclick = () =>
      downloadChartImage("target", "image/png").catch((error) =>
        window.alert(error instanceof Error ? error.message : "Chart download failed."),
      );
  }
  if (elements.downloadTargetJpeg) {
    elements.downloadTargetJpeg.onclick = () =>
      downloadChartImage("target", "image/jpeg").catch((error) =>
        window.alert(error instanceof Error ? error.message : "Chart download failed."),
      );
  }
  if (elements.exportTargetExcel) {
    elements.exportTargetExcel.onclick = () => {
      exportChartDataExcel("target");
    };
  }
}

function setMasterStatus(message, tone = "info") {
  if (!elements.masterStatus) {
    return;
  }
  elements.masterStatus.textContent = message;
  elements.masterStatus.classList.remove("is-success", "is-error");
  if (tone === "success") {
    elements.masterStatus.classList.add("is-success");
  }
  if (tone === "error") {
    elements.masterStatus.classList.add("is-error");
  }
}

function restoreMasterData() {
  storageRemove(STORAGE_KEYS.outletMaster);
  storageRemove(STORAGE_KEYS.deletedOutlets);
  storageRemove(STORAGE_KEYS.historicalMaster);
  storageRemove(STORAGE_KEYS.historicalImport);
  storageRemove(STORAGE_KEYS.areaTargets);
  storageRemove(STORAGE_KEYS.deletedAreaTargets);
  currentData = createEffectiveData(baseData);
  currentState = clampStateToData(currentState || buildDefaultState(currentData), currentData);

  if (elements.outletMasterForm) {
    elements.outletMasterForm.reset();
  }
  if (elements.historicalImportForm) {
    elements.historicalImportForm.reset();
  }
  if (elements.historicalMasterForm) {
    elements.historicalMasterForm.reset();
  }
  if (elements.targetMasterForm) {
    elements.targetMasterForm.reset();
  }

  renderMasterPage();
  setMasterStatus("Master page was restored to the built-in outlet, Hist., and target state.", "success");
}

function buildMasterBackupPayload() {
  return {
    app: "daily-sales-dashboard",
    version: 3,
    exportedAt: new Date().toISOString(),
    sourceMeta: {
      sourceFile: baseData?.meta?.sourceFile,
      sourceMode: baseData?.meta?.sourceMode,
      dateMin: baseData?.meta?.dateMin,
      dateMax: baseData?.meta?.dateMax,
    },
    dataset: baseData,
    state: currentState,
    master: {
      outlets: loadMasterOutlets().map(normalizeMasterOutlet).filter((entry) => entry.sapCode && entry.outletName),
      deletedOutlets: loadDeletedOutletCodes(),
      historicalEntries: loadMasterHistoricalEntries()
        .map(normalizeMasterHistoricalEntry)
        .filter((entry) => entry.sapCode && entry.referenceDate && entry.product && entry.historicalKl > 0),
      historicalImports: loadHistoricalImportFiles().map(normalizeHistoricalImportFile),
      areaTargets: getAreaTargetEntries(),
      deletedAreaTargets: loadDeletedAreaTargetKeys(),
    },
  };
}

function backupMasterData() {
  const payload = buildMasterBackupPayload();
  const stamp = new Date().toISOString().slice(0, 10);
  downloadTextFile(
    `daily-sales-master-backup-${stamp}.json`,
    JSON.stringify(payload, null, 2),
    "application/json;charset=utf-8",
  );
  setMasterStatus("Master backup file downloaded successfully.", "success");
}

function restoreMasterBackupPayload(payload) {
  const master = payload?.master || payload;
  const outlets = master?.outlets || master?.outletMaster || [];
  const deletedOutlets = master?.deletedOutlets || [];
  const historicalEntries = master?.historicalEntries || master?.historicalMaster || [];
  const historicalImports = master?.historicalImports || master?.historicalImport || [];
  const areaTargets = master?.areaTargets || master?.targets || [];
  const deletedAreaTargets = master?.deletedAreaTargets || [];

  if (
    !Array.isArray(outlets) ||
    !Array.isArray(deletedOutlets) ||
    !Array.isArray(historicalEntries) ||
    !Array.isArray(historicalImports) ||
    !Array.isArray(areaTargets) ||
    !Array.isArray(deletedAreaTargets)
  ) {
    throw new Error("Backup file is not a valid master backup.");
  }

  if (payload?.dataset?.rows && payload?.dataset?.meta) {
    baseData = buildDatasetFromSnapshot(payload.dataset);
    if (!savePersistedDataset(baseData)) {
      throw new Error("Browser storage is full while restoring the backup dataset.");
    }
  }

  const normalizedOutlets = outlets.map(normalizeMasterOutlet).filter((entry) => entry.sapCode && entry.outletName);
  const normalizedDeletedOutlets = [...new Set(deletedOutlets.map(normalizeCode).filter(Boolean))];
  const normalizedHistoricalEntries = historicalEntries
    .map(normalizeMasterHistoricalEntry)
    .filter((entry) => entry.sapCode && entry.referenceDate && entry.product && entry.historicalKl > 0);
  const normalizedHistoricalImports = historicalImports.map(normalizeHistoricalImportFile);
  const normalizedAreaTargets = areaTargets
    .map(normalizeAreaTarget)
    .filter((entry) => entry.salesArea && TARGET_PRODUCTS.some((product) => entry.targets[product] > 0));
  const normalizedDeletedAreaTargets = [...new Set(deletedAreaTargets.map(cleanText).filter(Boolean))];

  if (!storageSet(STORAGE_KEYS.outletMaster, JSON.stringify(normalizedOutlets))) {
    throw new Error("Browser storage is full while restoring outlet master data.");
  }
  if (!storageSet(STORAGE_KEYS.deletedOutlets, JSON.stringify(normalizedDeletedOutlets))) {
    throw new Error("Browser storage is full while restoring deleted outlet data.");
  }
  if (!storageSet(STORAGE_KEYS.historicalMaster, JSON.stringify(normalizedHistoricalEntries))) {
    throw new Error("Browser storage is full while restoring Hist. master data.");
  }
  if (!storageSet(STORAGE_KEYS.historicalImport, JSON.stringify(normalizedHistoricalImports))) {
    throw new Error("Browser storage is full while restoring imported Hist. backups.");
  }
  if (!storageSet(STORAGE_KEYS.areaTargets, JSON.stringify(normalizedAreaTargets))) {
    throw new Error("Browser storage is full while restoring area target data.");
  }
  if (!storageSet(STORAGE_KEYS.deletedAreaTargets, JSON.stringify(normalizedDeletedAreaTargets))) {
    throw new Error("Browser storage is full while restoring deleted target data.");
  }

  currentData = createEffectiveData(baseData);
  currentState = clampStateToData(payload?.state || currentState || buildDefaultState(currentData), currentData);
  savePersistedState(currentState);
  renderMasterPage();
  setMasterStatus(
    `Backup restored: ${formatWhole.format(normalizedOutlets.length)} outlets, ${formatWhole.format(
      normalizedHistoricalEntries.length,
    )} manual Hist. rows, ${formatWhole.format(normalizedHistoricalImports.length)} imported Hist. file(s), and ${formatWhole.format(
      normalizedAreaTargets.length,
    )} target row(s).`,
    "success",
  );
}

async function handleMasterBackupRestore(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  try {
    const payload = JSON.parse(await file.text());
    const shouldRestore = window.confirm(
      "Restore this master backup? This will replace custom outlets, Hist. entries, imported Hist. backups, sales-area targets, and the saved dashboard data in this browser.",
    );
    if (!shouldRestore) {
      return;
    }
    restoreMasterBackupPayload(payload);
  } catch (error) {
    setMasterStatus(error instanceof Error ? error.message : "Backup restore failed.", "error");
  } finally {
    event.target.value = "";
  }
}

function renderMasterTab(activeTab) {
  elements.masterTabButtons.forEach((button) => {
    button.classList.toggle("active", button.dataset.masterTab === activeTab);
  });
  elements.masterPanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.masterPanel === activeTab);
  });
}

function renderOutletMasterTable() {
  if (!elements.outletMasterBody) {
    return;
  }

  const entries = sortByName(currentData.filters.outlets || []);
  if (elements.masterOutletCount) {
    elements.masterOutletCount.textContent = String(entries.length);
  }

  if (!entries.length) {
    elements.outletMasterBody.innerHTML = '<tr><td colspan="7" class="empty-state">No outlet master records are available yet.</td></tr>';
    return;
  }

  elements.outletMasterBody.innerHTML = entries
    .map(
      (entry) => `
        <tr>
          <td>${escapeHtml(entry.sapCode)}</td>
          <td>${escapeHtml(entry.outletName)}</td>
          <td>${escapeHtml(entry.salesArea)}</td>
          <td>${escapeHtml(entry.appointedOn ? displayDate(entry.appointedOn) : "--")}</td>
          <td>${escapeHtml(entry.outletType || (entry.isProjectAbhuyaday ? "Project Abhuyaday" : "Standard"))}</td>
          <td>${escapeHtml(entry.plant || "--")}</td>
          <td>
            <button class="ghost-button row-action-button" type="button" data-edit-outlet="${escapeHtml(entry.sapCode)}">Edit</button>
            <button class="ghost-button row-action-button" type="button" data-remove-outlet="${escapeHtml(entry.sapCode)}">Delete</button>
          </td>
        </tr>
      `,
    )
    .join("");
}

function renderHistoricalOutletOptions() {
  if (!elements.historicalOutletSelect) {
    return;
  }

  const options = ['<option value="">Choose outlet</option>'].concat(
    sortByName(currentData.filters.outlets).map(
      (outlet) =>
        `<option value="${escapeHtml(outlet.sapCode)}">${escapeHtml(outlet.outletName)} (${escapeHtml(outlet.salesArea)})</option>`,
    ),
  );
  elements.historicalOutletSelect.innerHTML = options.join("");
}

function renderTargetAreaOptions() {
  if (!elements.targetAreaSelect) {
    return;
  }

  const savedAreas = getAreaTargetEntries().map((entry) => entry.salesArea);
  const areas = [...new Set([...(currentData.filters.areas || []), ...savedAreas])]
    .filter(Boolean)
    .sort((left, right) => left.localeCompare(right));
  elements.targetAreaSelect.innerHTML = ['<option value="">Choose sales area</option>']
    .concat(areas.map((area) => `<option value="${escapeHtml(area)}">${escapeHtml(area)}</option>`))
    .join("");
}

function renderHistoricalImportSummary() {
  if (!elements.historicalImportList) {
    return;
  }

  const importFiles = loadHistoricalImportFiles().map(normalizeHistoricalImportFile);
  if (!importFiles.length) {
    elements.historicalImportList.innerHTML =
      '<div class="empty-state">No Hist. Excel backups imported yet. Upload your MS, HSD, or Power Hist. file here.</div>';
    return;
  }

  elements.historicalImportList.innerHTML = importFiles
    .sort((left, right) => left.product.localeCompare(right.product))
    .map(
      (file) => `
        <div class="master-import-item">
          <div class="master-import-copy">
            <strong>${escapeHtml(file.name)}</strong>
            <span>${escapeHtml(file.product)} | ${escapeHtml(formatWhole.format(file.entryCount))} Hist. values</span>
            <span>Imported on ${escapeHtml(displayDate(normalizeDate(file.importedAt) || new Date().toISOString().slice(0, 10)))}</span>
          </div>
          <button class="ghost-button row-action-button" type="button" data-remove-import="${escapeHtml(file.id)}">Remove</button>
        </div>
      `,
    )
    .join("");
}

function renderHistoricalMasterTable() {
  if (!elements.historicalMasterBody) {
    return;
  }

  const importFiles = loadHistoricalImportFiles().map(normalizeHistoricalImportFile);
  const entries = loadMasterHistoricalEntries()
    .map(normalizeMasterHistoricalEntry)
    .filter((entry) => entry.sapCode && entry.referenceDate && entry.historicalKl > 0);
  if (elements.masterHistoricalCount) {
    const importedCount = importFiles.reduce((sum, file) => sum + file.entryCount, 0);
    elements.masterHistoricalCount.textContent = formatWhole.format(entries.length + importedCount);
  }

  if (!entries.length) {
    elements.historicalMasterBody.innerHTML = '<tr><td colspan="7" class="empty-state">No custom Hist. records added yet.</td></tr>';
    return;
  }

  const outletMap = new Map(currentData.filters.outlets.map((outlet) => [outlet.sapCode, outlet]));
  elements.historicalMasterBody.innerHTML = entries
    .sort((left, right) => right.referenceDate.localeCompare(left.referenceDate))
    .map((entry) => {
      const outlet = outletMap.get(entry.sapCode);
      return `
        <tr>
          <td>${escapeHtml(outlet?.outletName || "--")}</td>
          <td>${escapeHtml(entry.sapCode)}</td>
          <td>${escapeHtml(entry.product)}</td>
          <td>${escapeHtml(displayDate(entry.referenceDate))}</td>
          <td>${escapeHtml(formatNumber.format(entry.historicalKl))}</td>
          <td>${escapeHtml(entry.note || "--")}</td>
          <td><button class="ghost-button row-action-button" type="button" data-remove-history="${escapeHtml(
            `${entry.sapCode}|${entry.product}|${entry.referenceDate}`,
          )}">Remove</button></td>
        </tr>
      `;
    })
    .join("");
}

function renderTargetMasterTable() {
  if (!elements.targetMasterBody) {
    return;
  }

  const entries = getAreaTargetEntries();
  if (elements.masterTargetCount) {
    elements.masterTargetCount.textContent = formatWhole.format(entries.length);
  }

  if (!entries.length) {
    elements.targetMasterBody.innerHTML = '<tr><td colspan="9" class="empty-state">No month-wise sales-area target rows saved yet.</td></tr>';
    return;
  }

  elements.targetMasterBody.innerHTML = entries
    .sort((left, right) => (right.monthKey || "").localeCompare(left.monthKey || "") || left.salesArea.localeCompare(right.salesArea))
    .map(
      (entry) => `
        <tr>
          <td>${escapeHtml(displayMonthKey(entry.monthKey))}</td>
          <td>${escapeHtml(entry.salesArea)}</td>
          <td>${escapeHtml(formatNumber.format(entry.targets.MS))}</td>
          <td>${escapeHtml(formatNumber.format(entry.targets.HSD))}</td>
          <td>${escapeHtml(formatNumber.format(entry.targets.Power))}</td>
          <td>${escapeHtml(formatNumber.format(entry.targets.Lube))}</td>
          <td>${escapeHtml(formatNumber.format(entry.targets.DEF))}</td>
          <td>${escapeHtml(entry.note || "--")}</td>
          <td>
            <button class="ghost-button row-action-button" type="button" data-edit-target="${escapeHtml(areaTargetKey(entry))}">Edit</button>
            <button class="ghost-button row-action-button" type="button" data-remove-target="${escapeHtml(areaTargetKey(entry))}">Remove</button>
          </td>
        </tr>
      `,
    )
    .join("");
}

function renderMasterPage() {
  currentData = createEffectiveData(baseData);
  initializeMeta();
  renderOutletMasterTable();
  renderHistoricalOutletOptions();
  renderTargetAreaOptions();
  if (elements.targetMonth && !elements.targetMonth.value) {
    elements.targetMonth.value = getTargetMonthKey();
  }
  renderHistoricalImportSummary();
  renderHistoricalMasterTable();
  renderTargetMasterTable();
  wireExports();
  setMasterStatus(
    "Use the outlet tab for new outlets, Hist. tab for comparison data, and Target tab for area-wise targets.",
  );
}

function attachMasterEvents() {
  attachUploadEvents();

  if (elements.masterBack) {
    elements.masterBack.addEventListener("click", () => {
      if (window.history.length > 1) {
        window.history.back();
        return;
      }
      window.location.href = "./index.html";
    });
  }

  if (elements.restoreMasterData) {
    elements.restoreMasterData.addEventListener("click", () => {
      const hasCustomData =
        loadMasterOutlets().length > 0 ||
        loadDeletedOutletCodes().length > 0 ||
        loadMasterHistoricalEntries().length > 0 ||
        loadHistoricalImportFiles().length > 0 ||
        loadAreaTargets().length > 0 ||
        loadDeletedAreaTargetKeys().length > 0;
      if (!hasCustomData) {
        setMasterStatus("Master page is already using the built-in outlet and Hist. data.");
        return;
      }

      const shouldRestore = window.confirm(
        "Restore the master page to the built-in state? This will remove all custom outlet, Hist., and sales-area target entries saved on this browser.",
      );
      if (!shouldRestore) {
        return;
      }

      restoreMasterData();
    });
  }

  if (elements.backupMasterData) {
    elements.backupMasterData.addEventListener("click", backupMasterData);
  }

  if (elements.restoreMasterBackup) {
    elements.restoreMasterBackup.addEventListener("change", handleMasterBackupRestore);
  }

  if (elements.historicalImportForm) {
    elements.historicalImportForm.addEventListener("submit", async (event) => {
      event.preventDefault();

      const files = Array.from(elements.historicalExcelUpload?.files || []);
      const selectedProduct = elements.historicalImportProduct?.value || "AUTO";

      if (!files.length) {
        setMasterStatus("Choose at least one Hist. Excel file before importing.", "error");
        return;
      }

      if (selectedProduct !== "AUTO" && files.length > 1) {
        setMasterStatus("When product mapping is fixed to MS, HSD, or Power, import one Excel file at a time.", "error");
        return;
      }

      setMasterStatus(`Importing ${files.map((file) => file.name).join(", ")}...`);

      try {
        const importedFiles = [];
        for (const file of files) {
          importedFiles.push(...(await parseHistoricalExcelFile(file, selectedProduct)));
        }

        const productsToReplace = new Set(importedFiles.map((file) => file.product));
        const nextFiles = loadHistoricalImportFiles()
          .map(normalizeHistoricalImportFile)
          .filter((file) => !productsToReplace.has(file.product))
          .concat(importedFiles);

        if (!saveHistoricalImportFiles(nextFiles)) {
          throw new Error("Browser storage is full. Please restore unused master data and try again.");
        }

        currentData = createEffectiveData(baseData);
        currentState = clampStateToData(currentState || buildDefaultState(currentData), currentData);
        elements.historicalImportForm.reset();
        renderMasterPage();
        renderMasterTab("historical");
        setMasterStatus(
          `${importedFiles.map((file) => `${file.product}: ${file.name}`).join(" | ")} imported successfully.`,
          "success",
        );
      } catch (error) {
        setMasterStatus(error instanceof Error ? error.message : "Hist. Excel import failed.", "error");
      }
    });
  }

  elements.masterTabButtons.forEach((button) => {
    button.addEventListener("click", () => {
      renderMasterTab(button.dataset.masterTab);
    });
  });

  if (elements.outletMasterForm) {
    elements.outletMasterForm.addEventListener("submit", (event) => {
      event.preventDefault();
      const formData = new FormData(elements.outletMasterForm);
      const entry = normalizeMasterOutlet({
        sapCode: formData.get("sapCode"),
        outletName: formData.get("outletName"),
        salesArea: formData.get("salesArea"),
        appointedOn: formData.get("appointedOn"),
        plant: formData.get("plant"),
        notes: formData.get("notes"),
        outletType: formData.get("outletType"),
      });

      if (!entry.sapCode || !entry.outletName || !entry.salesArea) {
        setMasterStatus("SAP code, outlet name, and retail region are required for a new outlet.", "error");
        return;
      }

      const entries = loadMasterOutlets().map(normalizeMasterOutlet).filter((item) => item.sapCode !== entry.sapCode);
      entries.push(entry);
      saveMasterOutlets(entries);
      saveDeletedOutletCodes(loadDeletedOutletCodes().filter((code) => code !== entry.sapCode));
      currentData = createEffectiveData(baseData);
      currentState = clampStateToData(currentState || buildDefaultState(currentData), currentData);
      elements.outletMasterForm.reset();
      renderMasterPage();
      setMasterStatus(`${entry.outletName} was saved to the outlet master list.`, "success");
    });
  }

  if (elements.cancelOutletEdit && elements.outletMasterForm) {
    elements.cancelOutletEdit.addEventListener("click", () => {
      elements.outletMasterForm.reset();
      setMasterStatus("Outlet form cleared.");
    });
  }

  if (elements.historicalMasterForm) {
    elements.historicalMasterForm.addEventListener("submit", (event) => {
      event.preventDefault();
      const formData = new FormData(elements.historicalMasterForm);
      const entry = normalizeMasterHistoricalEntry({
        sapCode: formData.get("sapCode"),
        product: formData.get("product"),
        referenceDate: formData.get("referenceDate"),
        historicalKl: formData.get("historicalKl"),
        note: formData.get("note"),
      });

      if (!entry.sapCode || !entry.referenceDate || !entry.product || entry.historicalKl <= 0) {
        setMasterStatus("Outlet, product, reference date, and Hist. KL are required for Hist. data.", "error");
        return;
      }

      const entryKey = `${entry.sapCode}|${entry.product}|${entry.referenceDate}`;
      const entries = loadMasterHistoricalEntries()
        .map(normalizeMasterHistoricalEntry)
        .filter((item) => item.historicalKl > 0 && `${item.sapCode}|${item.product}|${item.referenceDate}` !== entryKey);
      entries.push(entry);
      saveMasterHistoricalEntries(entries);
      currentData = createEffectiveData(baseData);
      currentState = clampStateToData(currentState || buildDefaultState(currentData), currentData);
      const preservedValues = {
        sapCode: entry.sapCode,
        referenceDate: entry.referenceDate,
      };
      elements.historicalMasterForm.reset();
      renderMasterPage();
      renderMasterTab("historical");
      if (elements.historicalOutletSelect) {
        elements.historicalOutletSelect.value = preservedValues.sapCode;
      }
      const referenceDateInput = elements.historicalMasterForm.elements.namedItem("referenceDate");
      if (referenceDateInput instanceof HTMLInputElement) {
        referenceDateInput.value = preservedValues.referenceDate;
      }
      setMasterStatus(`Hist. data for ${entry.product} on ${displayDate(entry.referenceDate)} was saved.`, "success");
    });
  }

  if (elements.targetMasterForm) {
    elements.targetMasterForm.addEventListener("submit", (event) => {
      event.preventDefault();
      const formData = new FormData(elements.targetMasterForm);
      const entry = normalizeAreaTarget({
        targetMonth: formData.get("targetMonth"),
        salesArea: formData.get("salesArea"),
        targets: {
          MS: formData.get("msTarget"),
          HSD: formData.get("hsdTarget"),
          Power: formData.get("powerTarget"),
          Lube: formData.get("lubeTarget"),
          DEF: formData.get("defTarget"),
        },
        note: formData.get("note"),
        updatedAt: new Date().toISOString(),
      });

      if (!entry.salesArea) {
        setMasterStatus("Sales area is required before saving target data.", "error");
        return;
      }
      if (!entry.monthKey) {
        setMasterStatus("Target month is required before saving target data.", "error");
        return;
      }

      const hasAnyTarget = TARGET_PRODUCTS.some((product) => entry.targets[product] > 0);
      if (!hasAnyTarget) {
        setMasterStatus("Enter at least one target KL value for MS, HSD, Power, Lube, or DEF.", "error");
        return;
      }

      const entryKey = areaTargetKey(entry);
      const entries = getCustomAreaTargetEntries().filter((item) => areaTargetKey(item) !== entryKey);
      entries.push(entry);
      if (!saveAreaTargets(entries)) {
        setMasterStatus("Browser storage is full. Please backup and restore unused data, then try again.", "error");
        return;
      }
      saveDeletedAreaTargetKeys(loadDeletedAreaTargetKeys().filter((key) => key !== entryKey));

      elements.targetMasterForm.reset();
      if (elements.targetMonth) {
        elements.targetMonth.value = getTargetMonthKey();
      }
      renderMasterPage();
      renderMasterTab("targets");
      setMasterStatus(`${entry.salesArea} target values for ${displayMonthKey(entry.monthKey)} were saved.`, "success");
    });
  }

  if (elements.cancelTargetEdit && elements.targetMasterForm) {
    elements.cancelTargetEdit.addEventListener("click", () => {
      elements.targetMasterForm.reset();
      if (elements.targetMonth) {
        elements.targetMonth.value = getTargetMonthKey();
      }
      setMasterStatus("Target form cleared.");
    });
  }

  if (elements.outletMasterBody) {
    elements.outletMasterBody.addEventListener("click", (event) => {
      const editTarget = event.target.closest("[data-edit-outlet]");
      if (editTarget && elements.outletMasterForm) {
        const sapCode = editTarget.dataset.editOutlet;
        const entry = currentData.filters.outlets.find((outlet) => outlet.sapCode === sapCode);
        if (!entry) {
          return;
        }
        elements.outletMasterForm.elements.namedItem("sapCode").value = entry.sapCode;
        elements.outletMasterForm.elements.namedItem("outletName").value = entry.outletName;
        elements.outletMasterForm.elements.namedItem("salesArea").value = entry.salesArea;
        elements.outletMasterForm.elements.namedItem("appointedOn").value = normalizeDate(entry.appointedOn) || "";
        elements.outletMasterForm.elements.namedItem("outletType").value = entry.outletType || (entry.isProjectAbhuyaday ? "Project Abhuyaday" : "Standard");
        elements.outletMasterForm.elements.namedItem("plant").value = entry.plant || "";
        elements.outletMasterForm.elements.namedItem("notes").value = entry.notes || "";
        setMasterStatus(`Editing ${entry.outletName}. Save outlet to update this master row.`);
        return;
      }

      const target = event.target.closest("[data-remove-outlet]");
      if (!target) {
        return;
      }
      const sapCode = target.dataset.removeOutlet;
      const entry = currentData.filters.outlets.find((outlet) => outlet.sapCode === sapCode);
      const shouldDelete = window.confirm(`Delete ${entry?.outletName || sapCode} from the active outlet master list?`);
      if (!shouldDelete) {
        return;
      }
      const entries = loadMasterOutlets().map(normalizeMasterOutlet).filter((entry) => entry.sapCode !== sapCode);
      saveMasterOutlets(entries);
      saveDeletedOutletCodes([...loadDeletedOutletCodes(), sapCode]);
      currentData = createEffectiveData(baseData);
      currentState = clampStateToData(currentState || buildDefaultState(currentData), currentData);
      renderMasterPage();
      setMasterStatus(`Outlet ${entry?.outletName || sapCode} was deleted from the active outlet master list.`, "success");
    });
  }

  if (elements.historicalMasterBody) {
    elements.historicalMasterBody.addEventListener("click", (event) => {
      const target = event.target.closest("[data-remove-history]");
      if (!target) {
        return;
      }
      const entryKey = target.dataset.removeHistory;
      const entries = loadMasterHistoricalEntries()
        .map(normalizeMasterHistoricalEntry)
        .filter((entry) => `${entry.sapCode}|${entry.product}|${entry.referenceDate}` !== entryKey);
      saveMasterHistoricalEntries(entries);
      currentData = createEffectiveData(baseData);
      renderMasterPage();
      renderMasterTab("historical");
      setMasterStatus("Hist. entry removed.", "success");
    });
  }

  if (elements.historicalImportList) {
    elements.historicalImportList.addEventListener("click", (event) => {
      const target = event.target.closest("[data-remove-import]");
      if (!target) {
        return;
      }

      const importId = target.dataset.removeImport;
      const remainingFiles = loadHistoricalImportFiles()
        .map(normalizeHistoricalImportFile)
        .filter((file) => file.id !== importId);

      if (!saveHistoricalImportFiles(remainingFiles)) {
        setMasterStatus("Imported Hist. backup could not be updated. Please try again.", "error");
        return;
      }

      currentData = createEffectiveData(baseData);
      renderMasterPage();
      renderMasterTab("historical");
      setMasterStatus("Imported Hist. backup removed.", "success");
    });
  }

  if (elements.targetMasterBody) {
    elements.targetMasterBody.addEventListener("click", (event) => {
      const editTarget = event.target.closest("[data-edit-target]");
      if (editTarget && elements.targetMasterForm) {
        const entry = getAreaTargetEntries().find((item) => areaTargetKey(item) === editTarget.dataset.editTarget);
        if (!entry) {
          return;
        }
        elements.targetMasterForm.elements.namedItem("targetMonth").value = entry.monthKey || getTargetMonthKey();
        elements.targetMasterForm.elements.namedItem("salesArea").value = entry.salesArea;
        elements.targetMasterForm.elements.namedItem("msTarget").value = entry.targets.MS || "";
        elements.targetMasterForm.elements.namedItem("hsdTarget").value = entry.targets.HSD || "";
        elements.targetMasterForm.elements.namedItem("powerTarget").value = entry.targets.Power || "";
        elements.targetMasterForm.elements.namedItem("lubeTarget").value = entry.targets.Lube || "";
        elements.targetMasterForm.elements.namedItem("defTarget").value = entry.targets.DEF || "";
        elements.targetMasterForm.elements.namedItem("note").value = entry.note || "";
        setMasterStatus(`Editing ${entry.salesArea} target for ${displayMonthKey(entry.monthKey)}.`);
        return;
      }

      const removeTarget = event.target.closest("[data-remove-target]");
      if (!removeTarget) {
        return;
      }

      const targetKey = removeTarget.dataset.removeTarget;
      const entry = getAreaTargetEntries().find((item) => areaTargetKey(item) === targetKey);
      const entries = getCustomAreaTargetEntries().filter((item) => areaTargetKey(item) !== targetKey);
      if (!saveAreaTargets(entries)) {
        setMasterStatus("Target row could not be removed. Please try again.", "error");
        return;
      }
      const builtInKeys = new Set(getBuiltInAreaTargetEntries().map(areaTargetKey));
      const deletedKeys = loadDeletedAreaTargetKeys().filter((key) => key !== targetKey);
      if (builtInKeys.has(targetKey)) {
        deletedKeys.push(targetKey);
      }
      saveDeletedAreaTargetKeys(deletedKeys);

      renderMasterPage();
      renderMasterTab("targets");
      setMasterStatus(`${entry?.salesArea || "Selected"} target row removed.`, "success");
    });
  }
}

function render() {
  if (PAGE === "register" || PAGE === "monthly") {
    currentState.product = "ALL";
  }
  currentState = clampStateToData(currentState, currentData);
  initializeMeta();
  syncControls();
  populateAreaSelect();
  populatePlantSelect();
  populateOutletSelect();
  renderProductPills();

  const filteredRows = getFilteredRows();
  const filteredOutlets = getFilteredOutlets();
  const summary = summarize(filteredRows);
  const productSummary = elements.productCards ? buildProductSummary(filteredRows) : [];
  const areaProductSummary =
    elements.areaBars || elements.exportAreaExcel || elements.downloadAreaPng || elements.downloadAreaJpeg
      ? buildAreaProductSummary(filteredRows, filteredOutlets)
      : [];
  const targetRows =
    elements.targetCharts || elements.exportTargetExcel || elements.downloadTargetPng || elements.downloadTargetJpeg
      ? getTargetComparisonRows()
      : [];
  const areaTargetSummary = targetRows.length ? buildAreaTargetSummary(targetRows) : [];
  lastRegisterRows = elements.registerBody || elements.exportRegister || elements.exportRegisterExcel ? buildRegisterRows(filteredRows, filteredOutlets) : [];
  lastMonthlyRegisterRows =
    elements.monthlyRegisterBody || elements.exportMonthlyRegister || elements.exportMonthlyRegisterExcel
      ? buildMonthlyRegisterRows(filteredRows, filteredOutlets)
      : [];
  lastHistoricalRows = elements.historicalBody || elements.historicalMetrics || elements.exportHistoricalExcel
    ? buildHistoricalRows(filteredRows, filteredOutlets)
    : [];
  lastOutletRows =
    elements.outletBody || elements.nilSaleBody || elements.exportSummary || elements.exportSummaryExcel
      ? buildOutletRows(filteredRows, filteredOutlets)
      : [];

  renderSummaryText(filteredRows);
  renderMetrics(summary);
  renderProductCards(productSummary);
  renderAreaBars(areaProductSummary);
  renderTargetCharts(areaTargetSummary);
  renderRegisterTable(lastRegisterRows);
  renderMonthlyRegisterTable(lastMonthlyRegisterRows);
  renderHistoricalMetrics(lastHistoricalRows);
  renderHistoricalAreaTable(lastHistoricalRows);
  renderHistoricalTable(lastHistoricalRows);
  renderOutletTable(lastOutletRows);
  renderNilSaleHomeList(lastOutletRows);
  wireExports();

  if (PAGE !== "home") {
    savePersistedState(currentState);
  }
}

function applyDataset(dataset, options = {}) {
  baseData = dataset;
  currentData = createEffectiveData(baseData);
  if (options.keepState) {
    currentState = clampStateToData(currentState, currentData);
  } else {
    currentState = buildDefaultState(currentData);
  }

  savePersistedDataset(baseData);
  savePersistedState(currentState);
  if (PAGE === "master") {
    renderMasterPage();
    return;
  }
  render();
}

async function handleCsvUpload(event) {
  const file = event.target.files?.[0];
  if (!file) {
    return;
  }

  setUploadStatus(`Reading ${file.name}...`);

  try {
    const dataset = isExcelFile(file) ? await buildDatasetFromExcelFile(file) : buildDatasetFromCsv(await file.text(), file.name);
    applyDataset(dataset);
    setUploadStatus(`${file.name} loaded successfully. The dashboard is now using your uploaded daily sales file.`, "success");
  } catch (error) {
    setUploadStatus(
      `${file.name} could not be loaded. ${error instanceof Error ? error.message : "Please check the file format and try again."}`,
      "error",
    );
  } finally {
    if (elements.csvUpload) {
      elements.csvUpload.value = "";
    }
  }
}

function attachUploadEvents() {
  if (elements.csvUpload) {
    elements.csvUpload.addEventListener("change", handleCsvUpload);
  }

  if (elements.restoreBuiltIn) {
    elements.restoreBuiltIn.addEventListener("click", () => {
      applyDataset(builtInData);
      setUploadStatus("Using the built-in workbook snapshot again. Upload another daily sales file anytime to replace it.", "info");
    });
  }
}

function attachPanelToggle(button, panelBody, showLabel, hideLabel) {
  if (!button || !panelBody) {
    return;
  }
  button.addEventListener("click", () => {
    const isHidden = panelBody.classList.toggle("hidden");
    button.textContent = isHidden ? showLabel : hideLabel;
  });
}

function attachEvents() {
  attachUploadEvents();

  if (elements.startDate) {
    elements.startDate.addEventListener("input", (event) => {
      currentState.startDate = event.target.value;
      render();
    });
  }

  if (elements.endDate) {
    elements.endDate.addEventListener("input", (event) => {
      currentState.endDate = event.target.value;
      render();
    });
  }

  if (elements.salesArea) {
    elements.salesArea.addEventListener("change", (event) => {
      currentState.salesArea = event.target.value;
      currentState.outletCode = "ALL";
      render();
    });
  }

  if (elements.plantFilter) {
    elements.plantFilter.addEventListener("change", (event) => {
      currentState.plantDescription = event.target.value;
      currentState.outletCode = "ALL";
      render();
    });
  }

  if (elements.outletSelect) {
    elements.outletSelect.addEventListener("change", (event) => {
      currentState.outletCode = event.target.value;
      render();
    });
  }

  if (elements.outletSearch) {
    elements.outletSearch.addEventListener("input", (event) => {
      currentState.outletSearch = event.target.value.trim();
      render();
    });
  }

  if (elements.paFilter) {
    elements.paFilter.addEventListener("change", (event) => {
      currentState.paFilter = event.target.value;
      render();
    });
  }

  if (elements.latestDay) {
    elements.latestDay.addEventListener("click", () => {
      const pageRange = getPageDateRange(currentData);
      currentState.startDate = pageRange.defaultDate;
      currentState.endDate = pageRange.defaultDate;
      render();
    });
  }

  if (elements.fullPeriod) {
    elements.fullPeriod.addEventListener("click", () => {
      const pageRange = getPageDateRange(currentData);
      currentState.startDate = pageRange.min;
      currentState.endDate = pageRange.max;
      render();
    });
  }

  if (elements.resetFilters) {
    elements.resetFilters.addEventListener("click", () => {
      currentState = buildDefaultState(currentData);
      render();
    });
  }

  if (elements.shareDashboardEmail) {
    elements.shareDashboardEmail.addEventListener("click", shareDashboardByEmail);
  }

  attachPanelToggle(elements.toggleNilSale, elements.nilSaleListWrap, "View nil sale outlets", "Hide nil sale outlets");
  attachPanelToggle(
    elements.toggleMonthlyRegister,
    elements.monthlyRegisterListWrap,
    "View monthly register",
    "Hide monthly register",
  );
  attachPanelToggle(
    elements.toggleOutletSummary,
    elements.outletSummaryListWrap,
    "View outlet summary",
    "Hide outlet summary",
  );

  wireExports();
}

function expandMonthlyDefaultRange() {
  return;
}

function boot() {
  builtInData = buildDatasetFromSnapshot(DEFAULT_DATA);
  baseData = loadPersistedDataset() || builtInData;
  currentData = createEffectiveData(baseData);
  currentState = buildDefaultState(currentData);
  expandMonthlyDefaultRange();

  if (PAGE === "master") {
    attachMasterEvents();
    renderMasterPage();
    setUploadStatus(
      currentData.meta.sourceMode === "Uploaded CSV" || currentData.meta.sourceMode === "Uploaded Excel"
        ? `${currentData.meta.sourceFile} is active right now. Upload another daily sales file anytime to replace it.`
        : "Using the built-in workbook snapshot right now. Upload your next daily sales Excel or CSV file here to refresh the dashboard.",
      currentData.meta.sourceMode === "Uploaded CSV" || currentData.meta.sourceMode === "Uploaded Excel" ? "success" : "info",
    );
    return;
  }

  attachEvents();
  render();

  if (elements.uploadStatus) {
    setUploadStatus(
      currentData.meta.sourceMode === "Uploaded CSV" || currentData.meta.sourceMode === "Uploaded Excel"
        ? `${currentData.meta.sourceFile} is active right now. Upload another daily sales file anytime to replace it.`
        : "Using the built-in workbook snapshot right now. Upload your next daily sales Excel or CSV file to refresh the dashboard instantly.",
      currentData.meta.sourceMode === "Uploaded CSV" || currentData.meta.sourceMode === "Uploaded Excel" ? "success" : "info",
    );
  }
}

boot();
