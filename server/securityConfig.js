import crypto from 'crypto';
import { ROLE_LEVEL } from './roleAccess.js';

function hashPassword(password) {
  return crypto.createHash('sha256').update(String(password || '')).digest('hex');
}

export function loadSecurityConfig(env = process.env) {
  const tokenSecret = String(env.AUTH_TOKEN_SECRET || '').trim();
  if (tokenSecret.length < 32) throw new Error('AUTH_TOKEN_SECRET must be configured with at least 32 characters');
  let parsed = [];
  if (env.AUTH_USERS_JSON) {
    try { parsed = JSON.parse(env.AUTH_USERS_JSON); }
    catch { throw new Error('AUTH_USERS_JSON must contain valid JSON'); }
    if (!Array.isArray(parsed)) throw new Error('AUTH_USERS_JSON must be an array');
  }
  const users = parsed.map((user) => ({
    username: String(user.username || '').trim().toLocaleLowerCase('en-US'),
    name: String(user.name || user.username || '').trim(),
    role: ROLE_LEVEL[user.role] ? user.role : 'viewer',
    passwordHash: String(user.passwordHash || hashPassword(user.password || '')),
  })).filter((user) => user.username && /^[a-f0-9]{64}$/i.test(user.passwordHash));
  if (parsed.length !== users.length) throw new Error('AUTH_USERS_JSON contains an invalid user');
  return Object.freeze({ tokenSecret, users });
}
