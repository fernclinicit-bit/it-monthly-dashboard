import fs from 'node:fs/promises';
import { FileBlob, SpreadsheetFile } from '@oai/artifact-tool';
const src='C:/Users/Fernclinic/Desktop/Inventory_Inventory_Results (2).xlsx';
const out='D:/แดชบอร์ด/outputs/inventory-results-2';
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(src));
await fs.mkdir(out,{recursive:true});
console.log((await wb.inspect({kind:'workbook,sheet,table,region',maxChars:24000,tableMaxRows:30,tableMaxCols:25,tableMaxCellChars:150})).ndjson);
for(const sh of wb.worksheets.items){const used=sh.getUsedRange();console.log('SHEET',sh.name,'USED',used?.address??'none');const preview=await wb.render({sheetName:sh.name,autoCrop:'all',scale:1,format:'png'});await fs.writeFile(`${out}/before_${sh.name.replace(/[\\/:*?\"<>|]/g,'_')}.png`,new Uint8Array(await preview.arrayBuffer()));}
