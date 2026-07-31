import fs from 'node:fs/promises';

const apiBase = 'https://it-monthly-dashboard-new.onrender.com';
const assets = JSON.parse(await fs.readFile('D:/แดชบอร์ด/outputs/inventory-results-2/assets.json', 'utf8'));
const beforeResponse = await fetch(`${apiBase}/api/db-state`);
if (!beforeResponse.ok) throw new Error(`Cannot read dashboard state: ${beforeResponse.status}`);
const before = await beforeResponse.json();

const assetMetrics = {
  totalAssets: assets.length,
  assetsBroken: assets.filter((asset) => asset.status === 'รอซ่อม').length,
  assetsLost: assets.filter((asset) => asset.status === 'สูญหาย').length,
  assetsVacant: assets.filter((asset) => asset.status === 'ว่าง').length,
};

const data = Object.fromEntries(Object.entries(before.data).map(([monthKey, monthData]) => [
  monthKey,
  { ...monthData, ...assetMetrics },
]));

const syncResponse = await fetch(`${apiBase}/api/sync-all`, {
  method: 'POST',
  headers: { 'Content-Type': 'application/json' },
  body: JSON.stringify({ data, assetsList: assets }),
});
const syncResult = await syncResponse.json();
if (!syncResponse.ok) throw new Error(syncResult.error || `Sync failed: ${syncResponse.status}`);

const afterResponse = await fetch(`${apiBase}/api/db-state`);
if (!afterResponse.ok) throw new Error(`Cannot verify dashboard state: ${afterResponse.status}`);
const after = await afterResponse.json();

const preservedKeys = [
  'monthName', 'assetValue', 'assetsExpiring', 'ticketsCount', 'slaPercent',
  'responseTime', 'resolutionTime', 'csat', 'totalSoftware', 'licensesInUse',
  'licensesVacant', 'softwareCost', 'softwareExpiring', 'backupSuccess',
  'securityIncidents', 'antivirusCoverage', 'mfaCoverage', 'repairCount',
  'repairCost', 'softwareExpiringDetails', 'ongoingProjects', 'recommendations',
];
const preservationErrors = [];
for (const [monthKey, monthBefore] of Object.entries(before.data)) {
  for (const key of preservedKeys) {
    if (JSON.stringify(monthBefore[key]) !== JSON.stringify(after.data?.[monthKey]?.[key])) {
      preservationErrors.push(`${monthKey}.${key}`);
    }
  }
}

const expectedIds = assets.map((asset) => asset.sn).sort((a, b) => a - b);
const actualIds = after.assetsList.map((asset) => asset.sn).sort((a, b) => a - b);
const result = {
  syncSuccess: Boolean(syncResult.success),
  imported: after.assetsList.length,
  firstNumber: actualIds[0],
  lastNumber: actualIds.at(-1),
  idsExact: JSON.stringify(expectedIds) === JSON.stringify(actualIds),
  metrics: assetMetrics,
  preservationErrors,
  sample: after.assetsList.find((asset) => asset.sn === 23),
};
console.log(JSON.stringify(result, null, 2));
if (!result.idsExact || result.imported !== 136 || preservationErrors.length) process.exitCode = 2;
