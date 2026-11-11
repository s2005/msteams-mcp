/**
 * People search for personal (Teams free) accounts.
 *
 * Personal accounts have no Substrate token. The Teams free web client searches
 * people through the middle tier (`POST /api/mt/beta/users/searchUsers`), which
 * accepts the opaque mt token plus the Skype token.
 */

import { httpRequest } from '../utils/http.js';
import { PERSONAL_MT_API, getPersonalMtHeaders } from '../utils/api-config.js';
import { type Result, ok } from '../types/result.js';
import type { McpError } from '../types/errors.js';
import { requireSkypeSpacesAuthWithConfig } from '../utils/auth-guards.js';
import type { PersonSearchResult } from '../utils/parsers.js';

/** Reads a string field, returning undefined for anything else. */
function str(value: unknown): string | undefined {
  return typeof value === 'string' && value !== '' ? value : undefined;
}

/**
 * Maps a searchUsers response to person results.
 *
 * The response is keyed by the search term (`{ "<term>": { userProfiles: [...] } }`),
 * and the key may differ in case from the typed query, so every value is read.
 * Profiles without an MRI are skipped, and duplicate MRIs are returned once.
 *
 * @param data - Parsed response body
 * @param limit - Maximum number of results to return
 */
export function parseConsumerPeopleResponse(data: unknown, limit: number): PersonSearchResult[] {
  const results: PersonSearchResult[] = [];
  if (!data || typeof data !== 'object') return results;

  const seen = new Set<string>();
  for (const group of Object.values(data as Record<string, unknown>)) {
    const profiles = (group as { userProfiles?: unknown } | null)?.userProfiles;
    if (!Array.isArray(profiles)) continue;

    for (const raw of profiles) {
      const p = (raw ?? {}) as Record<string, unknown>;
      const mri = str(p.mri);
      if (!mri || seen.has(mri)) continue;
      seen.add(mri);

      results.push({
        id: str(p.objectId) ?? mri,
        mri,
        displayName: str(p.displayName) ?? '',
        email: str(p.userPrincipalName),
        givenName: str(p.givenName),
        surname: str(p.surname),
      });
      if (results.length >= limit) return results;
    }
  }
  return results;
}

/**
 * Searches people on a personal (Teams free) account.
 *
 * @param query - Name, email or phone fragment to search for
 * @param limit - Maximum number of results
 */
export async function searchConsumerPeople(
  query: string,
  limit: number
): Promise<Result<PersonSearchResult[], McpError>> {
  const authResult = requireSkypeSpacesAuthWithConfig();
  if (!authResult.ok) {
    return authResult;
  }
  const { skypeToken, spacesToken: mtToken, regionConfig } = authResult.value;

  const response = await httpRequest<Record<string, unknown>>(
    PERSONAL_MT_API.searchUsers(regionConfig.middleTierUrl, limit),
    {
      method: 'POST',
      headers: getPersonalMtHeaders(skypeToken, mtToken),
      body: JSON.stringify({ emails: [], phones: [], searchKeyWord: query }),
    }
  );

  if (!response.ok) {
    return response;
  }

  return ok(parseConsumerPeopleResponse(response.value.data, limit));
}
