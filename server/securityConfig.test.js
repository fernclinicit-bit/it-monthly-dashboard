import test from 'node:test';
import assert from 'node:assert/strict';
import { loadSecurityConfig } from './securityConfig.js';

test('rejects missing token secret and embedded fallback users', () => {
  assert.throws(() => loadSecurityConfig({}), /AUTH_TOKEN_SECRET/);
});

test('accepts explicit secret and users only', () => {
  const config = loadSecurityConfig({
    AUTH_TOKEN_SECRET: 'x'.repeat(32),
    AUTH_USERS_JSON: JSON.stringify([{ username: 'admin', name: 'Admin', role: 'admin', passwordHash: 'a'.repeat(64) }]),
  });
  assert.equal(config.users.length, 1);
  assert.equal(config.users[0].username, 'admin');
});

