/**
 * Checks the personal-account allowlist against the real tool definitions.
 */

import { describe, it, expect } from 'vitest';
import { PERSONAL_ACCOUNT_TOOLS, getToolDefinitions } from './registry.js';

describe('PERSONAL_ACCOUNT_TOOLS', () => {
  it('only names tools that exist in the registry', () => {
    const names = new Set(getToolDefinitions().map(t => t.name));
    const missing = [...PERSONAL_ACCOUNT_TOOLS].filter(n => !names.has(n));

    expect(missing).toEqual([]);
  });
});
