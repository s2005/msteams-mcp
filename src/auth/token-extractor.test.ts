/**
 * Tests for personal-account (Teams free) extraction and work-session parity.
 */

import { describe, it, expect } from 'vitest';
import {
  bestOpaqueToken,
  extractRegionConfig,
  extractMessageAuth,
  extractMriFromSkypeToken,
  extractPersonalMtToken,
  extractPersonalAuthToken,
} from './token-extractor.js';
import type { SessionState } from './session-store.js';
import { personalSessionState, workSessionState } from '../__fixtures__/api-responses.js';

const personal = personalSessionState as SessionState;
const work = workSessionState as SessionState;

function fakeJwt(payload: Record<string, unknown>): string {
  const enc = (o: unknown) => Buffer.from(JSON.stringify(o)).toString('base64url');
  return `${enc({ alg: 'none' })}.${enc(payload)}.sig`;
}

describe('extractRegionConfig', () => {
  it('returns consumer values for a personal session', () => {
    expect(extractRegionConfig(personal)).toEqual({
      region: 'consumer',
      partition: '',
      regionPartition: '',
      hasPartition: false,
      middleTierUrl: 'https://teams.live.com/api/mt',
      chatServiceUrl: 'https://teams.live.com/api/chatsvc/consumer',
      csaServiceUrl: 'https://teams.live.com/api/csa',
      teamsBaseUrl: 'https://teams.live.com',
    });
  });

  it('leaves work output unchanged', () => {
    expect(extractRegionConfig(work)).toEqual({
      region: 'amer',
      partition: '02',
      regionPartition: 'amer-02',
      hasPartition: true,
      middleTierUrl: 'https://teams.microsoft.com/api/mt/part/amer-02',
      chatServiceUrl: 'https://teams.microsoft.com/api/chatsvc/amer',
      csaServiceUrl: 'https://teams.microsoft.com/api/csa/amer',
      teamsBaseUrl: 'https://teams.microsoft.com',
    });
  });
});

describe('opaque token extraction', () => {
  it('picks the latest-expiring unexpired token for the mt scope', () => {
    const info = extractPersonalMtToken(personal);
    expect(info?.token).toBe('fake-opaque-mt-token');
    expect(info?.expiry.getUTCFullYear()).toBe(2100);
  });

  it('selects by target', () => {
    expect(extractPersonalAuthToken(personal)?.token).toBe('fake-opaque-auth-token');
  });

  it('skips expired entries', () => {
    const state: SessionState = {
      cookies: [],
      origins: [{
        origin: 'https://teams.live.com',
        localStorage: personal.origins[0].localStorage.filter(i => i.name === 'msal-auth-token-expired'),
      }],
    };
    expect(extractPersonalAuthToken(state)).toBeNull();
  });

  it('returns null when no target matches', () => {
    expect(bestOpaqueToken(personal, 'no-such-scope')).toBeNull();
  });
});

describe('extractMessageAuth', () => {
  it('reads the teams.live.com cookies for a personal session', () => {
    const state: SessionState = {
      ...personal,
      cookies: personal.cookies.map(c =>
        c.name === 'skypetoken_asm' && c.domain?.includes('live')
          ? { ...c, value: fakeJwt({ skypeid: 'live:fake.user' }) }
          : c),
    };
    const auth = extractMessageAuth(state);
    expect(auth?.authToken).toBe('fake-personal-authtoken');
    expect(auth?.userMri).toBe('8:live:fake.user');
  });
});

describe('extractMriFromSkypeToken', () => {
  it('maps live: ids to 8:live:', () => {
    expect(extractMriFromSkypeToken(fakeJwt({ skypeid: 'live:abc' }))).toBe('8:live:abc');
  });

  it('keeps existing 8: and orgid: handling', () => {
    expect(extractMriFromSkypeToken(fakeJwt({ skypeid: '8:orgid:g' }))).toBe('8:orgid:g');
    expect(extractMriFromSkypeToken(fakeJwt({ skypeid: 'orgid:g' }))).toBe('8:orgid:g');
  });
});
