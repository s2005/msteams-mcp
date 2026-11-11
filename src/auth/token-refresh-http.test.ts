/**
 * Tests for browserless HTTP token refresh.
 * 
 * Tests the MSAL cache extraction and session state update logic.
 * Network calls are mocked to avoid hitting real Azure AD endpoints.
 */

import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';

// We need to mock the session-store module
// before importing the module under test.
vi.mock('./session-store.js', () => ({
  readSessionState: vi.fn(),
  writeSessionState: vi.fn(),
  clearTokenCache: vi.fn(),
  getTeamsOrigin: vi.fn(),
}));

import { refreshTokensViaHttp } from './token-refresh-http.js';
import { readSessionState, writeSessionState, clearTokenCache, getTeamsOrigin } from './session-store.js';
import type { SessionState } from './session-store.js';
import { personalSessionState } from '../__fixtures__/api-responses.js';
import { TEAMS_CLIENT_VERSION } from '../utils/api-config.js';
import { PERSONAL_MT_SCOPE, PERSONAL_AUTH_SCOPE, PERSONAL_SKYPEID_KEY } from './account-type.js';

// ============================================================================
// Test Fixtures
// ============================================================================

const MOCK_HOME_ACCOUNT_ID = 'user-guid.tenant-guid';
const MOCK_CLIENT_ID = '5e3ce6c0-2b1f-4285-8d4b-75ee78787346';
const MOCK_TENANT_ID = 'tenant-guid';
const MOCK_ENVIRONMENT = 'login.windows.net';

/** Creates a minimal MSAL refresh token localStorage entry. */
function makeRefreshTokenEntry() {
  return {
    name: `${MOCK_HOME_ACCOUNT_ID}-${MOCK_ENVIRONMENT}-refreshtoken-${MOCK_CLIENT_ID}----`,
    value: JSON.stringify({
      credentialType: 'RefreshToken',
      homeAccountId: MOCK_HOME_ACCOUNT_ID,
      environment: MOCK_ENVIRONMENT,
      clientId: MOCK_CLIENT_ID,
      secret: 'mock-refresh-token-secret',
      expiresOn: String(Math.floor(Date.now() / 1000) + 86400),
      lastUpdatedAt: String(Date.now()),
    }),
  };
}

/** Creates a minimal MSAL access token localStorage entry. */
function makeAccessTokenEntry(resource: string, target: string, expiresInSeconds = 3600) {
  const now = Math.floor(Date.now() / 1000);
  return {
    name: `${MOCK_HOME_ACCOUNT_ID}-${MOCK_ENVIRONMENT}-accesstoken-${MOCK_CLIENT_ID}-${MOCK_TENANT_ID}-${target.toLowerCase()}`,
    value: JSON.stringify({
      credentialType: 'AccessToken',
      homeAccountId: MOCK_HOME_ACCOUNT_ID,
      environment: MOCK_ENVIRONMENT,
      clientId: MOCK_CLIENT_ID,
      realm: MOCK_TENANT_ID,
      target,
      tokenType: 'Bearer',
      secret: 'old-access-token',
      expiresOn: String(now + expiresInSeconds),
      extendedExpiresOn: String(now + expiresInSeconds + 3600),
      cachedAt: String(now),
    }),
  };
}

/** Creates a mock session state with MSAL cache entries. */
function makeMockSessionState(): SessionState {
  return {
    cookies: [
      {
        name: 'skypetoken_asm',
        value: 'old-skype-token',
        domain: '.asyncgw.teams.microsoft.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
      {
        name: 'skypetoken_asm',
        value: 'old-skype-token',
        domain: '.asm.skype.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: true,
        secure: true,
        sameSite: 'None',
      },
      {
        name: 'authtoken',
        value: 'Bearer%3Dold-auth-token',
        domain: 'teams.microsoft.com',
        path: '/',
        expires: Date.now() / 1000 + 3600,
        httpOnly: false,
        secure: true,
        sameSite: 'None',
      },
    ],
    origins: [
      {
        origin: 'https://teams.microsoft.com',
        localStorage: [
          makeRefreshTokenEntry(),
          makeAccessTokenEntry('substrate.office.com', 'https://substrate.office.com/.default'),
          makeAccessTokenEntry('api.spaces.skype.com', 'https://api.spaces.skype.com/.default'),
          makeAccessTokenEntry('chatsvcagg.teams.microsoft.com', 'https://chatsvcagg.teams.microsoft.com/.default'),
        ],
      },
    ],
  };
}

/** Creates a mock Azure AD token response. */
/** Access token entries in the last written session state. */
function savedAccessTokens() {
  const saved = vi.mocked(writeSessionState).mock.calls.at(-1)![0];
  return saved.origins[0].localStorage.map(item => JSON.parse(item.value))
    .filter(entry => entry.credentialType === 'AccessToken');
}

/** Whether the last written session state carries the new skype token. */
function savedNewSkypeToken(): boolean {
  const saved = vi.mocked(writeSessionState).mock.calls.at(-1)![0];
  return saved.cookies.some(cookie => cookie.name === 'skypetoken_asm' && cookie.value === 'new-skype-token');
}

function makeTokenResponse(scope: string, expiresIn = 3600) {
  return {
    access_token: `new-access-token-for-${scope}`,
    refresh_token: 'new-refresh-token',
    token_type: 'Bearer',
    expires_in: expiresIn,
    scope,
    ext_expires_in: expiresIn + 3600,
  };
}

// ============================================================================
// Tests
// ============================================================================

describe('refreshTokensViaHttp', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('fetch', vi.fn());
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('returns AUTH_REQUIRED when no session state exists', async () => {
    vi.mocked(readSessionState).mockReturnValue(null);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_REQUIRED');
      expect(result.error.message).toContain('No session state found');
    }
  });

  it('returns AUTH_REQUIRED when no refresh token in session', async () => {
    const state = makeMockSessionState();
    // Remove the refresh token entry
    const origin = state.origins[0];
    origin.localStorage = origin.localStorage.filter(item => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType !== 'RefreshToken';
      } catch { return true; }
    });

    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_REQUIRED');
      expect(result.error.message).toContain('No MSAL refresh token');
    }
  });

  it('refreshes all tokens successfully', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    // Mock Azure AD token endpoint responses
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('mocked-scope')), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('unexpected URL', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    expect(savedAccessTokens().filter(entry => entry.secret.startsWith('new-access-token'))).toHaveLength(3);
    expect(savedNewSkypeToken()).toBe(true);

    // Verify session state was written back
    expect(writeSessionState).toHaveBeenCalledOnce();

    // Verify token cache was cleared
    expect(clearTokenCache).toHaveBeenCalledOnce();

    // Verify fetch was called 4 times (3 token refreshes + 1 skype exchange)
    expect(vi.mocked(fetch)).toHaveBeenCalledTimes(4);
  });

  it('updates skypetoken_asm cookies in session state', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'brand-new-skype-token', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    // Check the session state that was written
    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const skypeCookies = writtenState.cookies.filter(
      (c: { name: string }) => c.name === 'skypetoken_asm'
    );
    expect(skypeCookies.length).toBe(2);
    for (const cookie of skypeCookies) {
      expect(cookie.value).toBe('brand-new-skype-token');
    }
  });

  it('handles Azure AD error (expired refresh token)', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async () => {
      return new Response(JSON.stringify({
        error: 'invalid_grant',
        error_description: 'AADSTS700082: The refresh token has expired.',
      }), { status: 400 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).toBe('AUTH_EXPIRED');
      expect(result.error.message).toContain('refresh token has expired');
    }
  });

  it.each([400, 401])('core refresh never requests Assignments even when it would fail with HTTP %s', async (status) => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL, init?: RequestInit) => {
      if (String(url).includes('login.microsoftonline.com')) {
        const scope = new URLSearchParams(String(init?.body)).get('scope')!;
        if (scope.includes('8f348934-64be-4bb2-bc16-c54c96789f43')) {
          return new Response(JSON.stringify({
            error: 'invalid_grant',
            error_description: 'AADSTS65001: Consent required for Assignments.',
          }), { status });
        }
        return new Response(JSON.stringify(makeTokenResponse(scope)), { status: 200 });
      }
      if (String(url).includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), { status: 200 });
      }
      throw new Error(`Unexpected URL: ${String(url)}`);
    });

    const result = await refreshTokensViaHttp();

    expect(result).toEqual({ ok: true, value: undefined });
    expect(vi.mocked(fetch).mock.calls.some(([, init]) => String(init?.body).includes('8f348934'))).toBe(false);
    expect(writeSessionState).toHaveBeenCalledOnce();
    const saved = vi.mocked(writeSessionState).mock.calls[0][0];
    const entries = saved.origins[0].localStorage.map(item => JSON.parse(item.value));
    const accessTokens = entries.filter(entry => entry.credentialType === 'AccessToken');
    expect(accessTokens).toHaveLength(3);
    for (const entry of accessTokens) {
      expect(entry.secret).toBe(`new-access-token-for-${entry.target}`);
    }
    expect(entries.find(entry => entry.credentialType === 'RefreshToken').secret).toBe('new-refresh-token');
    expect(saved.cookies.filter(cookie => cookie.name === 'skypetoken_asm')
      .every(cookie => cookie.value === 'new-skype-token')).toBe(true);
    expect(clearTokenCache).toHaveBeenCalledOnce();
  });

  it('continues with remaining scopes if one fails with network error', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    let callCount = 0;
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        callCount++;
        if (callCount === 1) {
          // First scope fails with network error
          throw new Error('ECONNRESET');
        }
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'new-skype-token', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    expect(result.ok).toBe(true);
    // 2 of 3 scopes succeeded (first one failed with network error)
    expect(savedAccessTokens().filter(entry => entry.secret.startsWith('new-access-token'))).toHaveLength(2);
  });

  it('handles skype token exchange failure gracefully', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify(makeTokenResponse('scope')), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response('Forbidden', { status: 403 });
      }
      return new Response('', { status: 500 });
    });

    const result = await refreshTokensViaHttp();

    // Should still succeed — skype token failure is non-fatal
    expect(result.ok).toBe(true);
    expect(savedAccessTokens().filter(entry => entry.secret.startsWith('new-access-token'))).toHaveLength(3);
    expect(savedNewSkypeToken()).toBe(false);
  });

  it('updates MSAL access token cache entries in localStorage', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify({
          access_token: 'fresh-substrate-token',
          refresh_token: 'new-rt',
          token_type: 'Bearer',
          expires_in: 7200,
          scope: 'https://substrate.office.com/.default',
          ext_expires_in: 10800,
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'st', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    // Check that the Substrate access token was updated in localStorage
    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const origin = writtenState.origins[0];
    const substrateEntry = origin.localStorage.find((item: { name: string; value: string }) => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType === 'AccessToken' && val.target?.includes('substrate.office.com');
      } catch { return false; }
    });

    expect(substrateEntry).toBeDefined();
    const parsed = JSON.parse(substrateEntry!.value);
    expect(parsed.secret).toBe('fresh-substrate-token');
  });

  it('updates refresh token when Azure AD rotates it', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);

    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL) => {
      const urlStr = String(url);
      if (urlStr.includes('login.microsoftonline.com')) {
        return new Response(JSON.stringify({
          access_token: 'new-at',
          refresh_token: 'rotated-refresh-token',
          token_type: 'Bearer',
          expires_in: 3600,
          scope: 'scope',
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      if (urlStr.includes('authsvc.teams.microsoft.com')) {
        return new Response(JSON.stringify({
          tokens: { skypeToken: 'st', expiresIn: 86400 },
        }), {
          status: 200, headers: { 'Content-Type': 'application/json' },
        });
      }
      return new Response('', { status: 500 });
    });

    await refreshTokensViaHttp();

    const writtenState = vi.mocked(writeSessionState).mock.calls[0][0];
    const origin = writtenState.origins[0];
    const rtEntry = origin.localStorage.find((item: { name: string; value: string }) => {
      try {
        const val = JSON.parse(item.value);
        return val.credentialType === 'RefreshToken';
      } catch { return false; }
    });

    expect(rtEntry).toBeDefined();
    const parsed = JSON.parse(rtEntry!.value);
    expect(parsed.secret).toBe('rotated-refresh-token');
  });
  it.each([
    ['AADSTS65001: Consent required.', 'ACCESS_DENIED'],
    ['AADSTS500011: Resource principal missing.', 'ACCESS_DENIED'],
    ['AADSTS50105: User not assigned to a role for the application.', 'ACCESS_DENIED'],
    ['AADSTS53003: Access has been blocked by Conditional Access policies.', 'ACCESS_DENIED'],
    ['AADSTS90094: Admin consent is required.', 'ACCESS_DENIED'],
    ['AADSTS650057: Invalid resource.', 'ACCESS_DENIED'],
    ['AADSTS700082: Refresh token expired.', 'AUTH_EXPIRED'],
    ['AADSTS50076: MFA required.', 'AUTH_EXPIRED'],
  ])('keeps Assignments refusal distinct from expired auth: %s', async (description, code) => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockResolvedValue(new Response(JSON.stringify({ error_description: description }), { status: 400 }));
    expect(await refreshTokensViaHttp('assignments')).toMatchObject({ ok: false, error: { code } });
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(writeSessionState).not.toHaveBeenCalled();
  });

  it('refreshes only Assignments and preserves Graph credentials with the same scopes', async () => {
    const state = makeMockSessionState();
    const graphEntry = makeAccessTokenEntry('graph', 'EduAssignments.ReadWrite');
    state.origins[0].localStorage.push(graphEntry);
    graphEntry.value = JSON.stringify({ ...JSON.parse(graphEntry.value), secret: `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ aud: '00000003-0000-0000-c000-000000000000' })).toString('base64url')}.sig` });
    const originalGraph = graphEntry.value;
    const accessToken = `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify({ aud: '8f348934-64be-4bb2-bc16-c54c96789f43', exp: Date.now() / 1000 + 3600 })).toString('base64url')}.sig`;
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockResolvedValue(new Response(JSON.stringify({
      ...makeTokenResponse('EduAssignments.ReadWrite'), access_token: accessToken,
    }), { status: 200 }));
    expect(await refreshTokensViaHttp('assignments')).toEqual({ ok: true, value: undefined });
    expect(fetch).toHaveBeenCalledTimes(1);
    expect(graphEntry.value).toBe(originalGraph);
    const saved = vi.mocked(writeSessionState).mock.calls[0][0];
    expect(saved.origins[0].localStorage.some(item => JSON.parse(item.value).secret === accessToken)).toBe(true);
    expect(saved.cookies[0].value).toBe('old-skype-token');
    expect(clearTokenCache).not.toHaveBeenCalled();
  });

  it('refreshes only Graph on demand, updating the Graph entry in place', async () => {
    const state = makeMockSessionState();
    const jwt = (claims: object) => `eyJhbGciOiJub25lIn0.${Buffer.from(JSON.stringify(claims)).toString('base64url')}.sig`;
    const graphEntry = makeAccessTokenEntry('graph', 'Files.ReadWrite.All');
    graphEntry.value = JSON.stringify({ ...JSON.parse(graphEntry.value), secret: jwt({ aud: 'https://graph.microsoft.com', exp: 1 }) });
    state.origins[0].localStorage.push(graphEntry);
    const assignmentsEntry = makeAccessTokenEntry('assignments', 'EduAssignments.ReadWrite');
    assignmentsEntry.value = JSON.stringify({ ...JSON.parse(assignmentsEntry.value), secret: jwt({ aud: '8f348934-64be-4bb2-bc16-c54c96789f43' }) });
    state.origins[0].localStorage.push(assignmentsEntry);
    const originalAssignments = assignmentsEntry.value;
    const accessToken = jwt({ aud: 'https://graph.microsoft.com', exp: Date.now() / 1000 + 3600 });
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockResolvedValue(new Response(JSON.stringify({
      ...makeTokenResponse('Files.ReadWrite.All'), access_token: accessToken,
    }), { status: 200 }));

    expect(await refreshTokensViaHttp('graph')).toEqual({ ok: true, value: undefined });

    expect(fetch).toHaveBeenCalledTimes(1);
    expect(new URLSearchParams(vi.mocked(fetch).mock.calls[0][1]!.body as string).get('scope')).toBe('https://graph.microsoft.com/.default offline_access');
    expect(JSON.parse(graphEntry.value).secret).toBe(accessToken);
    expect(assignmentsEntry.value).toBe(originalAssignments);
  });

  it('reports a Graph consent refusal as access denied, not an expired login', async () => {
    const state = makeMockSessionState();
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    vi.mocked(fetch).mockResolvedValue(new Response(JSON.stringify({ error_description: 'AADSTS65001: Consent required.' }), { status: 400 }));
    expect(await refreshTokensViaHttp('graph')).toMatchObject({ ok: false, error: { code: 'ACCESS_DENIED', message: expect.stringContaining('Microsoft Graph') } });
  });

});

describe('refreshTokensViaHttp for a personal (Teams free) session', () => {
  const PERSONAL_CLIENT_ID = '4b3e8f46-56d3-427f-b1e2-d239b2ea6bca';

  /** Installs a deep copy of the personal fixture as the saved session. */
  function installPersonalSession(): SessionState {
    const state = JSON.parse(JSON.stringify(personalSessionState)) as SessionState;
    vi.mocked(readSessionState).mockReturnValue(state);
    vi.mocked(getTeamsOrigin).mockReturnValue(state.origins[0]);
    return state;
  }

  /** Mocks the token endpoint and the consumer authz endpoint. */
  function mockPersonalEndpoints() {
    vi.mocked(fetch).mockImplementation(async (url: RequestInfo | URL, init?: RequestInit) => {
      const urlStr = String(url);
      if (urlStr.includes('/oauth2/v2.0/token')) {
        const scope = new URLSearchParams(String(init?.body)).get('scope') ?? '';
        return new Response(JSON.stringify(makeTokenResponse(scope, 1800)), { status: 200 });
      }
      if (urlStr.includes('/authz/consumer')) {
        return new Response(JSON.stringify({
          skypeToken: { skypetoken: 'new-consumer-skype-token', expiresIn: 86400, skypeid: 'live:fake.user' },
          regionGtms: {},
        }), { status: 200 });
      }
      return new Response('{}', { status: 404 });
    });
  }

  beforeEach(() => {
    vi.clearAllMocks();
    vi.stubGlobal('fetch', vi.fn());
  });

  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it('calls the consumers token endpoint with the personal client, scopes and origin', async () => {
    installPersonalSession();
    mockPersonalEndpoints();

    const result = await refreshTokensViaHttp();
    expect(result.ok).toBe(true);

    const tokenCalls = vi.mocked(fetch).mock.calls.filter(([url]) => String(url).includes('/oauth2/v2.0/token'));
    expect(tokenCalls).toHaveLength(2);
    const scopes: string[] = [];
    for (const [url, init] of tokenCalls) {
      expect(String(url)).toBe('https://login.microsoftonline.com/consumers/oauth2/v2.0/token');
      expect((init!.headers as Record<string, string>)['Origin']).toBe('https://teams.live.com');
      const body = new URLSearchParams(String(init!.body));
      expect(body.get('client_id')).toBe(PERSONAL_CLIENT_ID);
      expect(body.get('grant_type')).toBe('refresh_token');
      scopes.push(body.get('scope') ?? '');
    }
    expect(scopes.some(s => s.split(' ').includes(PERSONAL_MT_SCOPE))).toBe(true);
    expect(scopes.some(s => s.split(' ').includes(PERSONAL_AUTH_SCOPE))).toBe(true);
    for (const scope of scopes) {
      expect(scope.split(' ')).toEqual(expect.arrayContaining(['openid', 'profile', 'offline_access']));
    }
  });

  it('exchanges the auth token at the consumer authz endpoint with the client headers', async () => {
    installPersonalSession();
    mockPersonalEndpoints();

    await refreshTokensViaHttp();

    const authzCalls = vi.mocked(fetch).mock.calls.filter(([url]) => String(url).includes('/authz/consumer'));
    expect(authzCalls).toHaveLength(1);
    const [url, init] = authzCalls[0];
    expect(String(url)).toBe('https://teams.live.com/api/auth/v2.0/authz/consumer');
    const headers = init!.headers as Record<string, string>;
    expect(headers['Authorization']).toMatch(/^Bearer new-access-token-for-/);
    expect(headers['Authorization']).toContain(PERSONAL_AUTH_SCOPE);
    expect(headers['ms-teams-authz-type']).toBe('TokenRefresh');
    expect(headers['x-ms-client-type']).toBe('web');
    expect(headers['x-ms-client-version']).toBe(TEAMS_CLIENT_VERSION);
  });

  it('writes cookies on teams.live.com, new access tokens with expiresOn and the Skype ID', async () => {
    installPersonalSession();
    mockPersonalEndpoints();

    const result = await refreshTokensViaHttp();
    expect(result.ok).toBe(true);

    const saved = vi.mocked(writeSessionState).mock.calls.at(-1)![0];
    const live = (name: string) => saved.cookies.filter(c =>
      c.name === name && c.domain?.replace(/^\./, '') === 'teams.live.com');
    expect(live('skypetoken_asm')).toHaveLength(1);
    expect(live('skypetoken_asm')[0].value).toBe('new-consumer-skype-token');
    expect(live('authtoken')).toHaveLength(1);
    expect(decodeURIComponent(live('authtoken')[0].value)).toContain('Bearer=');
    // The work-domain cookie is untouched.
    expect(saved.cookies.find(c => c.domain === '.teams.microsoft.com')?.value).toBe('fake-work-skypetoken');

    const entries = saved.origins[0].localStorage
      .map(item => { try { return JSON.parse(item.value); } catch { return null; } })
      .filter(entry => entry?.credentialType === 'AccessToken');
    const refreshed = entries.filter(entry => String(entry.secret).startsWith('new-access-token-for-'));
    expect(refreshed.length).toBeGreaterThanOrEqual(2);
    const nowSec = Math.floor(Date.now() / 1000);
    for (const entry of refreshed) {
      expect(entry.realm).toBe('9188040d-6c67-4c5b-b112-36a304b66dad');
      expect(typeof entry.expiresOn).toBe('string');
      expect(Number(entry.expiresOn)).toBeGreaterThan(nowSec);
      expect(Number(entry.expiresOn)).toBeLessThanOrEqual(nowSec + 1800 + 5);
    }
    expect(refreshed.some(entry => String(entry.target).includes(PERSONAL_MT_SCOPE))).toBe(true);
    expect(refreshed.some(entry => String(entry.target).includes(PERSONAL_AUTH_SCOPE))).toBe(true);

    const skypeId = saved.origins[0].localStorage.find(item => item.name === PERSONAL_SKYPEID_KEY);
    expect(skypeId?.value).toBe('live:fake.user');
  });

  it.each(['assignments', 'graph'] as const)('does not call fetch for the %s resource', async (resource) => {
    installPersonalSession();

    const result = await refreshTokensViaHttp(resource);

    expect(fetch).not.toHaveBeenCalled();
    expect(result.ok).toBe(false);
    if (!result.ok) {
      expect(result.error.code).not.toBe('AUTH_REQUIRED');
      expect(result.error.code).not.toBe('AUTH_EXPIRED');
      expect(result.error.retryable).toBe(false);
    }
  });
});
