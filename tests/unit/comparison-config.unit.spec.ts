/**
 * comparison-config.unit.spec.ts
 *
 * Regression coverage for identityIsConfigured(): a blank ReportIdentity
 * field (tenantId: '' etc, as opposed to the REPLACE_WITH_ placeholder text)
 * used to pass this check silently, letting report-parity.spec.ts proceed
 * into a confusing failure deep in the embed/API layer instead of a clear
 * "CONFIG ERROR" message up front. Run with: npm run test:unit
 */

import { test, expect } from '@playwright/test';
import { identityIsConfigured, type ReportIdentity } from '../helpers/comparison-config.helpers';

const filled: ReportIdentity = {
  tenantId: 'tid', groupId: 'gid', reportId: 'rid', datasetId: 'did',
};

test.describe('identityIsConfigured', () => {
  test('accepts a fully filled-in identity', () => {
    expect(identityIsConfigured(filled)).toBe(true);
  });

  test('rejects an identity with a blank field', () => {
    expect(identityIsConfigured({ ...filled, tenantId: '' })).toBe(false);
    expect(identityIsConfigured({ ...filled, groupId: '' })).toBe(false);
    expect(identityIsConfigured({ ...filled, reportId: '' })).toBe(false);
    expect(identityIsConfigured({ ...filled, datasetId: '' })).toBe(false);
  });

  test('rejects an identity with a whitespace-only field', () => {
    expect(identityIsConfigured({ ...filled, tenantId: '   ' })).toBe(false);
  });

  test('rejects the REPLACE_WITH_ placeholder text', () => {
    expect(identityIsConfigured({ ...filled, tenantId: 'REPLACE_WITH_TENANT_ID' })).toBe(false);
  });

  test('rejects the placeholder case-insensitively', () => {
    expect(identityIsConfigured({ ...filled, tenantId: 'replace_with_tenant_id' })).toBe(false);
  });

  test('rlsRole being absent does not affect the result (it is optional)', () => {
    expect(identityIsConfigured(filled)).toBe(true);
  });
});
