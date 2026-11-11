import { beforeEach, describe, expect, it, vi } from 'vitest';

vi.mock('../auth/session-store.js', () => ({
  hasSessionState: vi.fn(() => true),
  isSessionLikelyExpired: vi.fn(() => false),
  clearAllLocalData: vi.fn(),
}));

const tokenMocks = vi.hoisted(() => ({
  getSubstrateTokenStatus: vi.fn(),
}));

vi.mock('../auth/account-type.js', () => ({
  getAccountType: vi.fn(() => 'work'),
  PERSONAL_LOGIN_URL: 'https://teams.live.com/v2/',
}));

vi.mock('../auth/token-extractor.js', () => ({
  getSubstrateTokenStatus: tokenMocks.getSubstrateTokenStatus,
  // Work sessions delegate to the Substrate status, as the real helper does.
  getPrimaryTokenStatus: vi.fn(() => tokenMocks.getSubstrateTokenStatus()),
  extractPersonalMtToken: vi.fn(),
  extractPersonalAuthToken: vi.fn(),
  getMessageAuthStatus: vi.fn(),
  extractMessageAuth: vi.fn(),
  extractCsaToken: vi.fn(),
  extractSubstrateToken: vi.fn(),
  extractSkypeSpacesToken: vi.fn(),
}));

vi.mock('../auth/token-refresh-http.js', () => ({ refreshTokensViaHttp: vi.fn() }));
vi.mock('../browser/context.js', () => ({ createBrowserContext: vi.fn(), closeBrowser: vi.fn() }));
vi.mock('../browser/auth.js', () => ({ ensureAuthenticated: vi.fn() }));
vi.mock('../utils/auth-guards.js', () => ({ resetAssignmentsAvailability: vi.fn(), clearRegionCache: vi.fn() }));

import { loginTool, statusTool, LoginInputSchema, authTools } from './auth-tools.js';
import { getAccountType } from '../auth/account-type.js';
import { clearAllLocalData, hasSessionState } from '../auth/session-store.js';
import {
  getSubstrateTokenStatus,
  getMessageAuthStatus,
  extractMessageAuth,
  extractCsaToken,
  extractSubstrateToken,
  extractSkypeSpacesToken,
} from '../auth/token-extractor.js';
import { refreshTokensViaHttp } from '../auth/token-refresh-http.js';
import { createBrowserContext, closeBrowser } from '../browser/context.js';
import { ensureAuthenticated } from '../browser/auth.js';

/** A session holding every core credential except the CSA (favourites) token. */
function sessionMissingCsaToken(): void {
  vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, minutesRemaining: 60 });
  vi.mocked(extractSubstrateToken).mockReturnValue({ token: 'substrate', expiry: new Date(Date.now() + 3.6e6) });
  vi.mocked(extractSkypeSpacesToken).mockReturnValue('spaces');
  vi.mocked(extractMessageAuth).mockReturnValue({ skypeToken: 's', authToken: 'a', userMri: '8:orgid:u' });
  vi.mocked(extractCsaToken).mockReturnValue(null);
}

describe('teams_login core credential top-up', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
  });

  it('mints the missing CSA token when the session is otherwise valid', async () => {
    sessionMissingCsaToken();

    const result = await loginTool.handler({ forceNew: false, accountType: 'work' });

    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
    expect(createBrowserContext).not.toHaveBeenCalled();
  });

  it('skips the refresh when every core credential is present', async () => {
    sessionMissingCsaToken();
    vi.mocked(extractCsaToken).mockReturnValue('csa');

    await loginTool.handler({ forceNew: false, accountType: 'work' });

    expect(refreshTokensViaHttp).not.toHaveBeenCalled();
  });

  it('tops up after a forced browser login leaves the CSA token behind', async () => {
    sessionMissingCsaToken();
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    const result = await loginTool.handler({ forceNew: true, accountType: 'work' });

    expect(result.success).toBe(true);
    expect(refreshTokensViaHttp).toHaveBeenCalledOnce();
  });

  it('forceNew wipes all previous login data and signs in from scratch in a visible browser', async () => {
    sessionMissingCsaToken();
    vi.mocked(extractCsaToken).mockReturnValue('csa');
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);

    await loginTool.handler({ forceNew: true, accountType: 'work' });

    expect(clearAllLocalData).toHaveBeenCalledOnce();
    expect(vi.mocked(clearAllLocalData).mock.invocationCallOrder[0])
      .toBeLessThan(vi.mocked(createBrowserContext).mock.invocationCallOrder[0]);
    // Valid-looking tokens from the old session must not short-circuit the login,
    // and a wiped profile has nothing for headless SSO to use.
    expect(createBrowserContext).toHaveBeenCalledOnce();
    expect(createBrowserContext).toHaveBeenCalledWith({ headless: false });
  });

  it('still reports success when the top-up fails', async () => {
    sessionMissingCsaToken();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({
      ok: false,
      error: { code: 'NETWORK_ERROR', message: 'offline' },
    } as never);

    const result = await loginTool.handler({ forceNew: false, accountType: 'work' });

    expect(result.success).toBe(true);
  });
});

describe('teams_login accountType input', () => {
  it('accepts work and personal, rejects anything else, leaves an omitted value unset', () => {
    expect(LoginInputSchema.parse({ accountType: 'personal' }).accountType).toBe('personal');
    expect(LoginInputSchema.parse({ accountType: 'work' }).accountType).toBe('work');
    expect(LoginInputSchema.safeParse({ accountType: 'other' }).success).toBe(false);
    expect(LoginInputSchema.parse({}).accountType).toBeUndefined();
  });

  it('keeps the saved personal account type when accountType is omitted', async () => {
    vi.clearAllMocks();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: false });
    vi.mocked(hasSessionState).mockReturnValue(true);
    vi.mocked(getAccountType).mockReturnValue('personal');

    await loginTool.handler(LoginInputSchema.parse({}));

    expect(vi.mocked(ensureAuthenticated).mock.calls[0][5]).toBe('https://teams.live.com/v2/');
    vi.mocked(getAccountType).mockReturnValue('work');
  });

  it('uses the work login URL when accountType is omitted and no session exists', async () => {
    vi.clearAllMocks();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);
    vi.mocked(hasSessionState).mockReturnValue(false);

    await loginTool.handler(LoginInputSchema.parse({ forceNew: true }));

    expect(vi.mocked(ensureAuthenticated).mock.calls[0][5]).toBe('https://teams.microsoft.com');
    vi.mocked(hasSessionState).mockReturnValue(true);
  });

  it('declares the enum in the JSON schema', () => {
    const def = authTools.find(t => t.definition.name === 'teams_login')!.definition;
    const prop = (def.inputSchema.properties as Record<string, { type: string; enum?: string[]; description: string }>).accountType;
    expect(prop.type).toBe('string');
    expect(prop.enum).toEqual(['work', 'personal']);
    expect(prop.description).toContain('teams.live.com');
  });

  it('opens the personal login URL for a personal visible login', async () => {
    vi.clearAllMocks();
    vi.mocked(refreshTokensViaHttp).mockResolvedValue({ ok: true, value: undefined });
    vi.mocked(createBrowserContext).mockResolvedValue({ page: {}, context: {} } as never);
    vi.mocked(closeBrowser).mockResolvedValue(undefined as never);
    await loginTool.handler({ forceNew: true, accountType: 'personal' });
    expect(vi.mocked(ensureAuthenticated).mock.calls[0][5]).toBe('https://teams.live.com/v2/');
  });
});

describe('teams_status accountType', () => {
  it('reports the saved session account type', async () => {
    vi.mocked(getMessageAuthStatus).mockReturnValue({ hasToken: true });
    vi.mocked(getSubstrateTokenStatus).mockReturnValue({ hasToken: true, minutesRemaining: 60 });
    vi.mocked(getAccountType).mockReturnValue('personal');
    const result = await statusTool.handler({});
    expect(result.success && (result.data as { accountType: string }).accountType).toBe('personal');
    vi.mocked(getAccountType).mockReturnValue('work');
    const work = await statusTool.handler({});
    expect(work.success && (work.data as { accountType: string }).accountType).toBe('work');
  });
});
