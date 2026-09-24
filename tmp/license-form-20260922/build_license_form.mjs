import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile, Workbook } from '@oai/artifact-tool';

const inputPath = 'C:/Users/Fernclinic/Downloads/แจ้งเตือนการชำระเงิน_ตาราง_เส้นตาราง.xlsx';
const outputDir = 'D:/All Ai/แดชบอร์ด/outputs/license-form-20260922';
const outputPath = `${outputDir}/License_Import_App_Format.xlsx`;

const inputBlob = await FileBlob.load(inputPath);
const sourceWorkbook = await SpreadsheetFile.importXlsx(inputBlob);
const sourceSummary = await sourceWorkbook.inspect({
  kind: 'workbook,sheet,table',
  maxChars: 5000,
  tableMaxRows: 5,
  tableMaxCols: 12,
});
console.log('SOURCE_SUMMARY');
console.log(sourceSummary.ndjson);

const sourceSheet = sourceWorkbook.worksheets.getItemAt(0);
const sourceValues = sourceSheet.getUsedRange(true).values;
const sourceRows = sourceValues.slice(1).filter((row) => String(row[0] ?? '').trim());

const toNumber = (value) => {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const normalized = String(value ?? '').replace(/,/g, '').replace(/[^0-9.-]/g, '');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
};

const toDate = (value) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    const excelEpoch = Date.UTC(1899, 11, 30);
    return new Date(excelEpoch + value * 86400000);
  }
  if (!value) return null;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const workbook = Workbook.create();
const sheet = workbook.worksheets.add('License Import');
sheet.showGridLines = false;

const headers = [
  'ชื่อซอฟต์แวร์/โปรแกรม',
  'Owner',
  'ใช้งาน',
  'ว่าง',
  'รวม',
  'ราคา',
  'ช่องทางชำระ',
  'วันที่ชำระ',
  'วันหมดสัญญา',
  'อีเมลสมัคร',
  'ผู้ใช้งานปัจจุบัน',
];

sheet.getRange('A1:K1').values = [headers];

const outputRows = sourceRows.map((row) => [
  String(row[0] ?? '').trim(),
  String(row[2] ?? '').trim(),
  0,
  0,
  null,
  toNumber(row[3]),
  String(row[6] ?? '').trim(),
  String(row[4] ?? '').trim(),
  toDate(row[5]),
  String(row[1] ?? '').trim(),
  String(row[8] ?? '').trim()
    || (String(row[7] ?? '').trim() ? `แผนก: ${String(row[7]).trim()}` : 'ไม่ระบุในไฟล์ต้นฉบับ'),
]);

if (outputRows.length > 0) {
  sheet.getRange(`A2:K${outputRows.length + 1}`).values = outputRows;
  sheet.getRange('E2').formulas = [['=C2+D2']];
  sheet.getRange(`E2:E${outputRows.length + 1}`).fillDown();
}

const lastRow = outputRows.length + 1;
const fullRange = sheet.getRange(`A1:K${lastRow}`);
fullRange.format.font = { name: 'Arial', size: 10, color: '#172033' };
fullRange.format.verticalAlignment = 'center';
fullRange.format.borders = { preset: 'all', style: 'thin', color: '#C7D2E5' };

const headerRange = sheet.getRange('A1:K1');
headerRange.format.fill = '#1E3A6D';
headerRange.format.font = { name: 'Arial', size: 10, bold: true, color: '#FFFFFF' };
headerRange.format.horizontalAlignment = 'center';
headerRange.format.verticalAlignment = 'center';
headerRange.format.rowHeight = 30;

if (lastRow >= 2) {
  sheet.getRange(`A2:K${lastRow}`).format.rowHeight = 42;
  sheet.getRange(`C2:E${lastRow}`).format.horizontalAlignment = 'center';
  sheet.getRange(`F2:F${lastRow}`).format.numberFormat = '#,##0.00';
  sheet.getRange(`I2:I${lastRow}`).format.numberFormat = 'yyyy-mm-dd';
  sheet.getRange(`A2:B${lastRow}`).format.fill = '#F8FAFC';
  sheet.getRange(`F2:K${lastRow}`).format.fill = '#F8FAFC';
  sheet.getRange(`C2:E${lastRow}`).format.fill = '#EFF6FF';
  sheet.getRange(`B2:B${lastRow}`).format.wrapText = true;
  sheet.getRange(`K2:K${lastRow}`).format.wrapText = true;
}

const widths = [28, 24, 10, 10, 10, 14, 23, 15, 16, 30, 38];
widths.forEach((width, index) => {
  sheet.getRangeByIndexes(0, index, Math.max(lastRow, 1), 1).format.columnWidth = width;
});

sheet.getRange('A1:K1').format.wrapText = true;
sheet.freezePanes.freezeRows(1);

if (lastRow >= 2) {
  const table = sheet.tables.add(`A1:K${lastRow}`, true, 'LicenseImportTable');
  table.style = 'TableStyleMedium2';
  table.showFilterButton = true;
  table.showBandedRows = true;
}

workbook.recalculate();

const check = await workbook.inspect({
  kind: 'table',
  range: `License Import!A1:K${Math.min(lastRow, 8)}`,
  include: 'values,formulas',
  tableMaxRows: 8,
  tableMaxCols: 11,
  maxChars: 10000,
});
console.log('OUTPUT_CHECK');
console.log(check.ndjson);

const errors = await workbook.inspect({
  kind: 'match',
  searchTerm: '#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!',
  options: { useRegex: true, maxResults: 100 },
  summary: 'final formula error scan',
});
console.log('ERROR_SCAN');
console.log(errors.ndjson);

await fs.mkdir(outputDir, { recursive: true });
const preview = await workbook.render({ sheetName: 'License Import', range: `A1:K${Math.min(lastRow, 16)}`, scale: 1.4, format: 'png' });
await fs.writeFile(`${outputDir}/preview.png`, new Uint8Array(await preview.arrayBuffer()));

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outputPath);
console.log(JSON.stringify({ outputPath, rows: outputRows.length, lastRow }));
