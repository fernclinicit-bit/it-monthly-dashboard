import test from 'node:test';
import assert from 'node:assert/strict';
import { classifyCommand, validateApprovalDecision, redactAuditDetails } from './commandGovernance.js';

test('separates read commands from commands that change state', () => {
  assert.equal(classifyCommand('GET'), 'read');
  assert.equal(classifyCommand('HEAD'), 'read');
  assert.equal(classifyCommand('POST'), 'change');
  assert.equal(classifyCommand('DELETE'), 'change');
});

test('prevents self approval', () => {
  assert.throws(() => validateApprovalDecision({ requester: 'alice', approver: 'alice', status: 'pending', expiresAt: Date.now() + 1000 }), /own command/i);
});

test('prevents expired or replayed approvals', () => {
  assert.throws(() => validateApprovalDecision({ requester: 'alice', approver: 'bob', status: 'approved', expiresAt: Date.now() + 1000 }), /pending/i);
  assert.throws(() => validateApprovalDecision({ requester: 'alice', approver: 'bob', status: 'pending', expiresAt: Date.now() - 1 }), /expired/i);
});

test('redacts nested credentials from audit details', () => {
  const value = redactAuditDetails({ target: 'ios', token: 'secret', nested: { password: 'hidden', safe: 'kept' } });
  assert.deepEqual(value, { target: 'ios', token: '[REDACTED]', nested: { password: '[REDACTED]', safe: 'kept' } });
});
