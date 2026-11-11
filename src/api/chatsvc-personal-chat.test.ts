/**
 * Tests for finding an existing 1:1 chat on a personal (Teams free) session.
 */

import { describe, it, expect, vi, beforeEach } from 'vitest';
import { ok } from '../types/result.js';
import { ErrorCode } from '../types/errors.js';

const httpRequest = vi.fn();

vi.mock('../utils/http.js', () => ({ httpRequest: (...args: unknown[]) => httpRequest(...args) }));
vi.mock('../utils/auth-guards.js', () => ({
  requireMessageAuth: vi.fn(),
  requireMessageAuthWithConfig: () => ({
    ok: true,
    value: {
      auth: { skypeToken: 'fake-skype', authToken: 'fake-auth', userMri: '8:live:me-fake' },
      region: 'consumer',
      baseUrl: 'https://teams.live.com',
    },
  }),
  getTeamsBaseUrl: () => 'https://teams.live.com',
  getTenantId: () => 'fake-tenant',
}));

import { findPersonalOneOnOneChat, normalisePersonalMri } from './chatsvc-messaging.js';

const conv = (id: string, type: string) => ({ id, threadProperties: { productThreadType: type } });

/** Routes mocked HTTP calls: the list endpoint and per-thread member lookups. */
function mockChatsvc(conversations: unknown[], members: Record<string, string[]>): void {
  httpRequest.mockImplementation(async (url: string) => {
    if (url.includes('/users/ME/conversations')) {
      return ok({ data: { conversations } });
    }
    const key = decodeURIComponent(url.split('/threads/')[1].split('?')[0]);
    return ok({ data: { members: (members[key] ?? []).map(id => ({ id })) } });
  });
}

describe('normalisePersonalMri', () => {
  it('keeps 8: identifiers as is', () => {
    expect(normalisePersonalMri('8:live:abc')).toBe('8:live:abc');
    expect(normalisePersonalMri('8:skypename')).toBe('8:skypename');
  });

  it('prefixes live: identifiers', () => {
    expect(normalisePersonalMri('live:abc')).toBe('8:live:abc');
  });

  it('prefixes bare values', () => {
    expect(normalisePersonalMri('skypename')).toBe('8:skypename');
  });
});

describe('findPersonalOneOnOneChat', () => {
  beforeEach(() => {
    httpRequest.mockReset();
  });

  it('finds the matching chat among several OneToOneChat threads, case-insensitively', async () => {
    mockChatsvc(
      [conv('19:uni01_a@thread.v2', 'OneToOneChat'), conv('19:uni01_b@thread.v2', 'OneToOneChat'), conv('19:uni01_c@thread.v2', 'OneToOneChat')],
      {
        '19:uni01_a@thread.v2': ['8:live:me-fake', '8:live:other-1'],
        '19:uni01_b@thread.v2': ['8:live:me-fake', '8:live:Target-Fake'],
        '19:uni01_c@thread.v2': ['8:live:me-fake'],
      }
    );

    const result = await findPersonalOneOnOneChat('live:target-fake');

    expect(result.ok).toBe(true);
    if (result.ok) {
      expect(result.value.conversationId).toBe('19:uni01_b@thread.v2');
      expect(result.value.otherUserId).toBe('8:live:target-fake');
      expect(result.value.currentUserId).toBe('8:live:me-fake');
    }
  });

  it('skips conversations that are not OneToOneChat', async () => {
    mockChatsvc(
      [conv('19:group@thread.v2', 'Chat'), conv('19:meet@thread.v2', 'Meeting')],
      { '19:group@thread.v2': ['8:live:target-fake'], '19:meet@thread.v2': ['8:live:target-fake'] }
    );

    const result = await findPersonalOneOnOneChat('8:live:target-fake');

    expect(result.ok).toBe(false);
    // Only the list call was made; no thread lookups.
    expect(httpRequest).toHaveBeenCalledTimes(1);
  });

  it('returns NOT_FOUND when no thread has the member', async () => {
    mockChatsvc([conv('19:uni01_a@thread.v2', 'OneToOneChat')], { '19:uni01_a@thread.v2': ['8:live:me-fake'] });

    const result = await findPersonalOneOnOneChat('nobody-fake');

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe(ErrorCode.NOT_FOUND);
      expect(result.error.retryable).toBe(false);
      expect(result.error.message).toContain('No existing 1:1 chat');
      expect(result.error.suggestions.join(' ')).toContain('teams_list_chats');
    }
  });
});
