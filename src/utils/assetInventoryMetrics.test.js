import test from 'node:test';
import assert from 'node:assert/strict';

import { buildVacantStockBreakdown } from './assetInventoryMetrics.js';

test('shows vacant monitors first with an explicit vacant label', () => {
  const assets = [
    { status: 'ว่าง', itemType: 'Keyboard' },
    { status: 'ใช้งาน', itemType: 'Monitor' },
    { status: 'ว่าง', itemType: 'Monitor' },
    { status: 'ว่าง', itemType: 'Monitor' }
  ];

  assert.deepEqual(buildVacantStockBreakdown(assets, (itemType) => itemType), [
    { label: 'Monitor', displayLabel: 'Monitor ว่าง', count: 2 },
    { label: 'Keyboard', displayLabel: 'Keyboard', count: 1 }
  ]);
});

