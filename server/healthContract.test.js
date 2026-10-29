import test from 'node:test';
import assert from 'node:assert/strict';
import { createHealthReport } from './healthContract.js';

test('returns the standard healthy contract', () => {
  const report = createHealthReport({ service: 'dashboard', version: '1.0.0', checks: { database: { status: 'ok' } }, now: new Date('2026-10-21T00:00:00Z') });
  assert.deepEqual(Object.keys(report), ['status', 'service', 'version', 'timestamp', 'checks']);
  assert.equal(report.status, 'ok');
});

test('marks a failed dependency unavailable without serializing secrets', () => {
  const report = createHealthReport({ service: 'dashboard', version: '1.0.0', checks: { database: { status: 'error', token: 'secret' } } });
  assert.equal(report.status, 'error');
  assert.doesNotMatch(JSON.stringify(report), /secret/);
});
