import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';

const src = 'C:/Users/Fernclinic/Desktop/Inventory_Inventory_Results.xlsx';
const out = 'D:/แดชบอร์ด/outputs/019f8875-e783-7a30-b5a4-4e36bff3248d';
const wb = await SpreadsheetFile.importXlsx(await FileBlob.load(src));
const summary = await wb.inspect({kind:'workbook,sheet,table,region', maxChars:12000, tableMaxRows:20, tableMaxCols:20, tableMaxCellChars:120});
console.log(summary.ndjson);
const sheets = await wb.inspect({kind:'sheet', include:'id,name', maxChars:3000});
console.log(sheets.ndjson);
for (const sh of wb.worksheets.items) {
  const used = sh.getUsedRange();
  console.log('SHEET', sh.name, 'USED', used?.address ?? 'none');
  const preview = await wb.render({sheetName:sh.name, autoCrop:'all', scale:1, format:'png'});
  const safe = sh.name.replace(/[\\/:*?"<>|]/g,'_');
  await fs.writeFile(`${out}/before_${safe}.png`, new Uint8Array(await preview.arrayBuffer()));
}
