/**
 * Account type detection.
 *
 * Teams runs in two flavours: work or school accounts (Entra ID tenants) and
 * Teams free / personal Microsoft accounts (the consumers tenant, served from
 * teams.live.com). The two use different hosts, client IDs, scopes and token
 * formats, so every account-dependent code path asks this module first.
 *
 * Note: this module imports session-store at runtime (to read the saved
 * session) and session-store imports the pure detector below. The cycle is
 * safe because both sides only use the other's exports inside functions.
 */

import { readSessionState, type SessionState } from './session-store.js';

/** Kind of Microsoft account behind the saved session. */
export type AccountType = 'work' | 'personal';

/** Tenant ID Microsoft uses for all personal (consumer) accounts. */
export const CONSUMERS_TENANT_ID = '9188040d-6c67-4c5b-b112-36a304b66dad';

/** OAuth client ID of the Teams free web app. */
export const PERSONAL_CLIENT_ID = '4b3e8f46-56d3-427f-b1e2-d239b2ea6bca';

/** Teams base URL for personal accounts. */
export const PERSONAL_TEAMS_BASE_URL = 'https://teams.live.com';

/** Login page for personal accounts. */
export const PERSONAL_LOGIN_URL = 'https://teams.live.com/v2/';

/** MSAL scope of the opaque middle-tier token used by personal accounts. */
export const PERSONAL_MT_SCOPE = 'https://mtsvc.fl.teams.microsoft.com/teams.mt.readwrite';

/** MSAL scope of the opaque auth-service token used by personal accounts. */
export const PERSONAL_AUTH_SCOPE = 'https://auth.fl.teams.microsoft.com/teams.auth.readwrite';

/**
 * localStorage key (teams.live.com origin) under which token refresh persists the
 * Skype ID returned by the consumer authz exchange. The consumer Skype token may
 * be opaque, so the user MRI cannot always be read from it.
 */
export const PERSONAL_SKYPEID_KEY = 'msteams-mcp:personal-skypeid';

/**
 * Pure check: does this session hold a consumers-tenant MSAL entry on the
 * teams.live.com origin?
 */
export function isPersonalSessionState(state: SessionState | null | undefined): boolean {
  const origin = state?.origins?.find(o => o.origin === PERSONAL_TEAMS_BASE_URL);
  if (!origin) return false;

  for (const item of origin.localStorage ?? []) {
    try {
      const entry = JSON.parse(item.value) as { realm?: unknown; homeAccountId?: unknown };
      if (
        (typeof entry.realm === 'string' && entry.realm.includes(CONSUMERS_TENANT_ID)) ||
        (typeof entry.homeAccountId === 'string' && entry.homeAccountId.includes(CONSUMERS_TENANT_ID))
      ) {
        return true;
      }
    } catch {
      // ignore non-JSON localStorage values
    }
  }
  return false;
}

/**
 * Determines the account type of a session.
 * Reads the saved session when `state` is not passed; no session means 'work'.
 */
export function getAccountType(state?: SessionState): AccountType {
  const sessionState = state ?? readSessionState();
  return isPersonalSessionState(sessionState) ? 'personal' : 'work';
}
