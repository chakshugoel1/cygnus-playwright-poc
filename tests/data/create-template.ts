/**
 * Creates the cygnus-expected.xlsx template with all fields from the screenshot.
 * Run once: npx ts-node tests/data/create-template.ts
 */
import ExcelJS from 'exceljs';
import * as path from 'path';
import * as fs from 'fs';

const OUTPUT_FILE = path.join(__dirname, 'cygnus-expected.xlsx');

async function main() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Employee');

  // ── Header row ─────────────────────────────────────────────────────────────
  const headerRow = ws.addRow(['Tab', 'Section', 'Field Label', 'Expected Value', 'Required']);
  headerRow.eachCell((cell) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '1F3864' } };
    cell.font = { bold: true, color: { argb: 'FFFFFF' } };
    cell.alignment = { horizontal: 'center' };
  });

  ws.columns = [
    { key: 'tab',      width: 12 },
    { key: 'section',  width: 30 },
    { key: 'label',    width: 32 },
    { key: 'expected', width: 50 },
    { key: 'required', width: 10 },
  ];

  // ── Data rows ──────────────────────────────────────────────────────────────
  // Fill in the Expected Value column with real values for your test account.
  // The values below match the screenshot. Update them for your test user.
  const rows: [string, string, string, string, string][] = [
    // ── My Organization ──────────────────────────────────────────────────────
    ['Employee', 'My Organization', 'Manager (N+1)',      'FILL_IN',  'true'],
    ['Employee', 'My Organization', 'Manager (N+2)',      'FILL_IN',  'true'],
    ['Employee', 'My Organization', 'Dept. Head',         'FILL_IN',  'true'],
    ['Employee', 'My Organization', 'DU Head',            'FILL_IN',  'true'],
    ['Employee', 'My Organization', 'Delivery Manager',   'FILL_IN',  'true'],
    ['Employee', 'My Organization', 'HRBP',               'FILL_IN',  'true'],
    ['Employee', 'My Organization', 'DU',                 'FILL_IN',  'true'],
    ['Employee', 'My Organization', 'Department',         'FILL_IN',  'true'],
    // ── Diamond Club ─────────────────────────────────────────────────────────
    ['Employee', 'DIAMOND CLUB',    'Current Project',    'FILL_IN',  'true'],
    ['Employee', 'DIAMOND CLUB',    'Total Exp',          'FILL_IN',  'false'],
    ['Employee', 'DIAMOND CLUB',    'SSI Exp',            'FILL_IN',  'false'],
    // ── My Qualifications & Skills ────────────────────────────────────────────
    // List items: label = the skill/qualification text to look for in the list
    ['Employee', 'My Qualifications & Skills', '.Net Azure (Primary Skill)',     '', 'true'],
    ['Employee', 'My Qualifications & Skills', '.Net_Angular (iReflect Skill)',  '', 'true'],
    // ── My Asset ─────────────────────────────────────────────────────────────
    // List items: label = asset identifier to look for in the section
    ['Employee', 'My Asset',        'FILL_IN_ASSET_1',    '',         'false'],
  ];

  for (const [tab, section, label, expected, required] of rows) {
    const row = ws.addRow([tab, section, label, expected, required]);
    const sectionColor = section === 'My Organization' ? 'E2EFDA'
      : section === 'DIAMOND CLUB' ? 'FCE4D6'
      : section === 'My Qualifications & Skills' ? 'FFF2CC'
      : 'EBF3FB';

    row.eachCell((cell) => {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: sectionColor } };
      cell.border = {
        top: { style: 'thin' }, bottom: { style: 'thin' },
        left: { style: 'thin' }, right: { style: 'thin' },
      };
    });
  }

  fs.mkdirSync(path.dirname(OUTPUT_FILE), { recursive: true });
  await wb.xlsx.writeFile(OUTPUT_FILE);
  console.log(`✅ Template created: ${OUTPUT_FILE}`);
  console.log('Fill in the "Expected Value" column with your test account\'s known values.');
}

main().catch(console.error);
