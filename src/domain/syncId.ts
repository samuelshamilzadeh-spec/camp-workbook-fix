import { randomBytes } from 'node:crypto';

/**
 * Row identity.
 *
 * Short, random, and not sequential — two clients stamping simultaneously must
 * never collide, and a sequential counter shared across concurrent writers is
 * exactly the thing that does collide. 12 characters of Crockford base32 is ~60
 * bits, which makes a collision across a workbook of a few hundred thousand rows
 * negligible while staying short enough to read out loud over the phone.
 *
 * Stamped once, never changed.
 */

// Crockford base32: no I, L, O or U, so a staff member reading an ID off a
// screen cannot confuse it with 1/0, and it cannot accidentally spell anything.
const ALPHABET = '0123456789ABCDEFGHJKMNPQRSTVWXYZ';
const ID_LENGTH = 12;
const PREFIX = 'S';

export function newSyncId(): string {
  const bytes = randomBytes(ID_LENGTH);
  let out = PREFIX;
  for (let i = 0; i < ID_LENGTH; i++) {
    out += ALPHABET[bytes[i]! % ALPHABET.length];
  }
  return out;
}

const SYNC_ID_PATTERN = new RegExp(`^${PREFIX}[${ALPHABET}]{${ID_LENGTH}}$`);

export function isSyncId(value: unknown): value is string {
  return typeof value === 'string' && SYNC_ID_PATTERN.test(value.trim());
}

export function normalizeSyncId(value: unknown): string | undefined {
  if (value === null || value === undefined) return undefined;
  const text = String(value).trim().toUpperCase();
  return SYNC_ID_PATTERN.test(text) ? text : undefined;
}
