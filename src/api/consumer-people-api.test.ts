/**
 * Unit tests for personal-account people search.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';
import { consumerSearchUsersResponse } from '../__fixtures__/api-responses.js';

const httpRequest = vi.fn();
vi.mock('../utils/http.js', () => ({ httpRequest: (...args: unknown[]) => httpRequest(...args) }));
vi.mock('../utils/auth-guards.js', () => ({
  requireSkypeSpacesAuthWithConfig: () => ok({
    skypeToken: 'fake-skype',
    spacesToken: 'fake-mt',
    personal: true,
    regionConfig: { middleTierUrl: 'https://teams.live.com/api/mt' },
  }),
}));

import { parseConsumerPeopleResponse, searchConsumerPeople } from './consumer-people-api.js';

describe('parseConsumerPeopleResponse', () => {
  it('maps profile fields into the person result shape', () => {
    const results = parseConsumerPeopleResponse(consumerSearchUsersResponse, 10);
    expect(results[0]).toEqual({
      id: '00000000-0000-0000-0000-00000000aaaa',
      mri: '8:live:.cid.fake-julia',
      displayName: 'Julia Fake',
      email: 'julia.fake@example.invalid',
      givenName: 'Julia',
      surname: 'Fake',
    });
  });

  it('reads every response key regardless of case and skips profiles without an MRI', () => {
    const results = parseConsumerPeopleResponse(consumerSearchUsersResponse, 10);
    expect(results.map(r => r.displayName)).toEqual(['Julia Fake', 'Julia Second']);
  });

  it('honours the limit', () => {
    expect(parseConsumerPeopleResponse(consumerSearchUsersResponse, 1)).toHaveLength(1);
  });

  it('returns an empty list for unexpected bodies', () => {
    expect(parseConsumerPeopleResponse(null, 5)).toEqual([]);
    expect(parseConsumerPeopleResponse({ x: { status: 'Ok' } }, 5)).toEqual([]);
  });
});

describe('searchConsumerPeople', () => {
  beforeEach(() => httpRequest.mockReset());

  it('posts the search body with personal mt headers', async () => {
    httpRequest.mockResolvedValue(ok({ status: 200, headers: new Headers(), data: consumerSearchUsersResponse }));

    const result = await searchConsumerPeople('julia', 5);

    expect(result.ok).toBe(true);
    const [url, options] = httpRequest.mock.calls[0];
    expect(url).toBe('https://teams.live.com/api/mt/beta/users/searchUsers?ggEnabled=true&resultCount=5');
    expect(options.method).toBe('POST');
    expect(JSON.parse(options.body)).toEqual({ emails: [], phones: [], searchKeyWord: 'julia' });
    expect(options.headers['Authorization']).toBe('Bearer fake-mt');
    expect(options.headers['X-Skypetoken']).toBe('fake-skype');
    expect(options.headers['Authentication']).toBeUndefined();
  });
});
