import fs from 'node:fs/promises';
import XLSX from 'xlsx';

const sourcePath = 'C:/Users/Fernclinic/Desktop/Inventory_Inventory_Results (2).xlsx';
const outputPath = 'D:/แดชบอร์ด/outputs/inventory-results-2/assets.json';
const workbook = XLSX.readFile(sourcePath, { cellDates: false });
const rows = XLSX.utils.sheet_to_json(workbook.Sheets.Inventory, { defval: '', raw: true });

const formatDate = (value) => {
  if (!value) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`;
  }
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
    }
  }
  return String(value).trim();
};

const assets = rows.map((row, index) => ({
  sn: Number(row.Nember) || index + 1,
  submittedOn: formatDate(row['Submitted on']),
  respondent: String(row.Respondents || '').trim(),
  date: formatDate(row['วันที่เบิกใช้งาน']),
  user: String(row['บุคคลเบิกใช้อุปกรณ์'] || '').trim(),
  position: String(row['ตำแหน่ง'] || '').trim(),
  itemType: String(row['รายการอุปกรณ์หลัก'] || '').trim(),
  additionalEquipment: String(row['อุปกรณ์เพิ่มเติมที่ต้องการเบิก'] || '').trim(),
  softwareApp: String(row['ซอต์ฟแวร์/ App'] || '').trim(),
  registeredEmail: String(row['เมลที่ลงทะเบียน'] || '').trim(),
  deviceSerial: String(row['หมายเลขอุปกรณ์ (เช่น  Ipad 016)'] || '').trim(),
  additionalSerial: String(row['หมายเลขอุปกรณ์ เพิ่มเติม  (เช่น  สาย อะเเดปเตอร์ ipad-011))'] || '').trim(),
  returnDueDate: formatDate(row['กำหนดคืนอุปกรณ์']),
  status: String(row['สถานะ'] || '').trim(),
  notes: String(row['หมายเหตุ'] || '').trim(),
  inspectionDate: formatDate(row['วันที่ตรวจสอบ']),
  purchaseDate: formatDate(row['วันที่ซื้อ']),
  warrantyEndDate: formatDate(row['วันหมดประกัน']),
  expense: Number(row['ค่าใช้จ่าย']) || 0,
}));

const duplicateIds = assets
  .map((asset) => asset.sn)
  .filter((sn, index, all) => all.indexOf(sn) !== index);

await fs.writeFile(outputPath, JSON.stringify(assets, null, 2), 'utf8');
console.log(JSON.stringify({
  rows: assets.length,
  first: assets[0],
  last: assets.at(-1),
  duplicateIds: [...new Set(duplicateIds)],
  blankStatus: assets.filter((asset) => !asset.status).length,
  blankItemType: assets.filter((asset) => !asset.itemType).length,
  outputPath,
}, null, 2));
