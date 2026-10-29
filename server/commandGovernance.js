const SENSITIVE_KEY = /password|secret|token|pin|api.?key|authorization|cookie/i;

export function classifyCommand(method) {
  return ['GET', 'HEAD', 'OPTIONS'].includes(String(method || '').toUpperCase()) ? 'read' : 'change';
}

export function validateApprovalDecision({ requester, approver, status, expiresAt }, now = Date.now()) {
  if (String(requester) === String(approver)) throw new Error('A requester cannot approve their own command');
  if (status !== 'pending') throw new Error('Only pending commands can be approved');
  if (new Date(expiresAt).getTime() <= now) throw new Error('The approval request has expired');
  return true;
}

export function redactAuditDetails(value) {
  if (Array.isArray(value)) return value.map(redactAuditDetails);
  if (!value || typeof value !== 'object') return value;
  return Object.fromEntries(Object.entries(value).map(([key, item]) => [
    key,
    SENSITIVE_KEY.test(key) ? '[REDACTED]' : redactAuditDetails(item),
  ]));
}
