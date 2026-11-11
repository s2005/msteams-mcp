/**
 * Tests for account type detection.
 */

import { describe, it, expect } from 'vitest';
import { getAccountType, isPersonalSessionState } from './account-type.js';
import type { SessionState } from './session-store.js';
import { personalSessionState, workSessionState } from '../__fixtures__/api-responses.js';

describe('getAccountType', () => {
  it('returns personal for a consumers-tenant teams.live.com session', () => {
    expect(getAccountType(personalSessionState as SessionState)).toBe('personal');
  });

  it('returns work for a work session', () => {
    expect(getAccountType(workSessionState as SessionState)).toBe('work');
  });

  it('returns work for an empty session', () => {
    expect(getAccountType({ cookies: [], origins: [] })).toBe('work');
  });

  it('returns work when teams.live.com holds no consumers-tenant entries', () => {
    const state: SessionState = {
      cookies: [],
      origins: [{ origin: 'https://teams.live.com', localStorage: [{ name: 'x', value: 'not json' }] }],
    };
    expect(isPersonalSessionState(state)).toBe(false);
  });
});
