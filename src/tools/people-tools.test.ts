/**
 * Unit tests for people tools on personal (Teams free) sessions.
 */

import { describe, it, expect, vi } from 'vitest';
import { ok } from '../types/result.js';

const searchConsumerPeople = vi.fn();
const searchPeople = vi.fn();
const resolveProfiles = vi.fn();

vi.mock('../api/consumer-people-api.js', () => ({ searchConsumerPeople: (...a: unknown[]) => searchConsumerPeople(...a) }));
vi.mock('../api/substrate-api.js', () => ({
  searchPeople: (...a: unknown[]) => searchPeople(...a),
  getFrequentContacts: vi.fn(),
}));
vi.mock('../api/profile-api.js', () => ({ resolveProfiles: (...a: unknown[]) => resolveProfiles(...a) }));
vi.mock('../auth/token-extractor.js', () => ({
  getUserProfile: () => null,
  extractMessageAuth: () => ({ skypeToken: 'fake-skype', authToken: 'fake-auth', userMri: '8:live:.cid.fake-me' }),
}));
vi.mock('../utils/auth-guards.js', () => ({ getCachedAccountType: () => 'personal' }));

// Load the tools barrel first so the registry/people-tools import cycle resolves in its normal order
import './index.js';
import { searchPeopleTool, getMeTool } from './people-tools.js';

describe('people tools on a personal session', () => {
  it('teams_search_people uses the consumer people search', async () => {
    const results = [{ id: 'x', mri: '8:live:.cid.fake', displayName: 'Fake Person' }];
    searchConsumerPeople.mockResolvedValue(ok(results));

    const out = await searchPeopleTool.handler({ query: 'fake', limit: 5 });

    expect(searchConsumerPeople).toHaveBeenCalledWith('fake', 5);
    expect(searchPeople).not.toHaveBeenCalled();
    if (!out.success) throw new Error('expected success');
    expect(out.data).toEqual({ query: 'fake', returned: 1, results });
  });

  it('teams_get_me builds the profile from the saved MRI and the short profile', async () => {
    resolveProfiles.mockResolvedValue(ok({
      profiles: new Map([['8:live:.cid.fake-me', { mri: '8:live:.cid.fake-me', displayName: 'Fake Me', email: 'me@example.invalid' }]]),
      unresolved: 0,
    }));

    const out = await getMeTool.handler({});

    expect(resolveProfiles).toHaveBeenCalledWith(['8:live:.cid.fake-me']);
    if (!out.success) throw new Error('expected success');
    expect(out.data).toMatchObject({
      profile: { mri: '8:live:.cid.fake-me', id: 'live:.cid.fake-me', displayName: 'Fake Me', email: 'me@example.invalid' },
    });
  });
});
