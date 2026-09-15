import { readFileSync } from 'node:fs';
import { describe, expect, test } from 'bun:test';

import { openWorkbook } from './index.js';

const fixture = readFileSync(
  new URL('../../crates/ooxml-opc/tests/fixtures/sample.xlsx', import.meta.url)
);

describe('@betteroffice/xlsx-node', () => {
  test('opens, inspects, edits, renders, and saves a workbook', async () => {
    const workbook = await openWorkbook(fixture);

    expect(workbook.sheetCount).toBeGreaterThan(0);
    const mutation = workbook.set(0, 'A1', 'BetterOffice');
    expect(mutation.applied).toBe(true);
    expect(workbook.cell(0, 'A1').input).toBe('BetterOffice');
    workbook.set(0, 'B1', '=1+2');
    expect(workbook.formula(0, 'B1')).toBe('1+2');
    expect(workbook.value(0, 'B1')).toMatchObject({ kind: 'number', number: 3 });
    expect(workbook.setStyle(0, 'A1:B1', { bold: true }).applied).toBe(true);

    const proposal = workbook.propose({
      agentId: 'test-agent',
      edits: [{ sheet: 0, address: 'A2', input: 'proposed' }]
    });
    expect(workbook.proposals).toHaveLength(1);
    expect(workbook.acceptProposal(proposal.id).applied).toBe(true);
    expect(workbook.cell(0, 'A2').input).toBe('proposed');
    const rendered = await workbook.renderSheet({ sheet: 0, range: 'A1:B3' });

    expect(rendered.data.subarray(0, 8)).toEqual(Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]));
    expect((await workbook.save()).subarray(0, 2)).toEqual(Buffer.from('PK'));
  });

  test('rejects client IDs outside the JavaScript safe integer range', async () => {
    await expect(openWorkbook(fixture, { clientId: Number.MAX_SAFE_INTEGER + 1 })).rejects.toThrow(
      'clientId must be a positive safe integer'
    );
  });
});
