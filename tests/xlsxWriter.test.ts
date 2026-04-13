/**
 * Phase 2 – XLSX Writer Tests
 *
 * Tests for SharedStringTable, CellValueHandler, XmlSaver, and Workbook.save/saveToBuffer.
 */

import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import JSZip from 'jszip';

import {
  Workbook,
  SaveFormat,
  SharedStringTable,
  CellValueHandler,
  CELL_TYPE_STRING,
  CELL_TYPE_NUMBER,
  CELL_TYPE_BOOLEAN,
  XmlSaver,
  Style,
} from '../src';

// ─── Helpers ────────────────────────────────────────────────────────────────

/** Create a temporary directory for test output. */
function tmpDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'xlsx-writer-test-'));
}

/** Unzip a buffer and return the JSZip instance. */
async function unzipBuffer(buf: Buffer): Promise<JSZip> {
  return JSZip.loadAsync(buf);
}

/** Extract a text file from a JSZip instance. */
async function readEntry(zip: JSZip, entryPath: string): Promise<string> {
  const entry = zip.file(entryPath);
  if (!entry) throw new Error(`Missing ZIP entry: ${entryPath}`);
  return entry.async('string');
}

// ─── 1. SharedStringTable ──────────────────────────────────────────────────

describe('SharedStringTable', () => {
  it('deduplicates strings and returns correct indices', () => {
    const sst = new SharedStringTable();

    const i0 = sst.addString('Hello');
    const i1 = sst.addString('World');
    const i2 = sst.addString('Hello'); // duplicate

    expect(i0).toBe(0);
    expect(i1).toBe(1);
    expect(i2).toBe(0); // same as first "Hello"

    expect(sst.getString(0)).toBe('Hello');
    expect(sst.getString(1)).toBe('World');
    expect(sst.count).toBe(3); // total references
    expect(sst.uniqueCount).toBe(2); // unique strings
  });

  it('generates valid SST XML', () => {
    const sst = new SharedStringTable();
    sst.addString('Alpha');
    sst.addString('Beta');
    sst.addString('Alpha'); // duplicate

    const xml = sst.toXml();
    expect(xml).toContain('<?xml version="1.0"');
    expect(xml).toContain('count="3"');
    expect(xml).toContain('uniqueCount="2"');
    expect(xml).toContain('<t>Alpha</t>');
    expect(xml).toContain('<t>Beta</t>');
  });
});

// ─── 2. CellValueHandler ──────────────────────────────────────────────────

describe('CellValueHandler', () => {
  it('detects cell types correctly', () => {
    expect(CellValueHandler.getCellType('hello')).toBe(CELL_TYPE_STRING);
    expect(CellValueHandler.getCellType(42)).toBe(CELL_TYPE_NUMBER);
    expect(CellValueHandler.getCellType(3.14)).toBe(CELL_TYPE_NUMBER);
    expect(CellValueHandler.getCellType(true)).toBe(CELL_TYPE_BOOLEAN);
    expect(CellValueHandler.getCellType(false)).toBe(CELL_TYPE_BOOLEAN);
    expect(CellValueHandler.getCellType(null)).toBeNull(); // null value → null type
    expect(CellValueHandler.getCellType(new Date())).toBe(CELL_TYPE_NUMBER);
  });

  it('converts dates to Excel serial numbers', () => {
    // 2024-01-01 00:00:00 → serial number should be around 45292
    const d = new Date(2024, 0, 1); // Jan 1, 2024
    const serial = CellValueHandler.dateToExcelSerial(d);
    expect(serial).toBeCloseTo(45292, 0);

    // Excel epoch: 1900-01-01 → serial 2 (Lotus 1-2-3 leap year bug: day 1 = Jan 0 1900)
    const epoch = new Date(1900, 0, 1);
    const epochSerial = CellValueHandler.dateToExcelSerial(epoch);
    expect(epochSerial).toBe(2);
  });

  it('formats values for XML output', () => {
    // formatValueForXml returns [formattedValue, cellType] tuple
    expect(CellValueHandler.formatValueForXml(42)).toEqual(['42', CELL_TYPE_NUMBER]);
    expect(CellValueHandler.formatValueForXml(true)).toEqual(['1', CELL_TYPE_BOOLEAN]);
    expect(CellValueHandler.formatValueForXml(false)).toEqual(['0', CELL_TYPE_BOOLEAN]);
    expect(CellValueHandler.formatValueForXml('hello')).toEqual(['hello', CELL_TYPE_STRING]);
  });
});

// ─── 3. XmlSaver – buffer output ──────────────────────────────────────────

describe('XmlSaver', () => {
  it('produces a valid ZIP with required OOXML entries', async () => {
    const wb = new Workbook();
    const ws = wb.worksheets[0];
    ws.cells.get('A1').value = 'Hello';
    ws.cells.get('B1').value = 42;

    const saver = new XmlSaver(wb);
    const buf = await saver.saveToBuffer();

    expect(buf).toBeInstanceOf(Buffer);
    expect(buf.length).toBeGreaterThan(0);

    const zip = await unzipBuffer(buf);

    // Required OOXML parts
    expect(zip.file('[Content_Types].xml')).not.toBeNull();
    expect(zip.file('_rels/.rels')).not.toBeNull();
    expect(zip.file('xl/workbook.xml')).not.toBeNull();
    expect(zip.file('xl/_rels/workbook.xml.rels')).not.toBeNull();
    expect(zip.file('xl/worksheets/sheet1.xml')).not.toBeNull();
    expect(zip.file('xl/styles.xml')).not.toBeNull();
    expect(zip.file('xl/sharedStrings.xml')).not.toBeNull();
    expect(zip.file('xl/theme/theme1.xml')).not.toBeNull();
    expect(zip.file('docProps/core.xml')).not.toBeNull();
    expect(zip.file('docProps/app.xml')).not.toBeNull();
  });

  it('writes correct worksheet XML with cell data', async () => {
    const wb = new Workbook();
    const ws = wb.worksheets[0];
    ws.cells.get('A1').value = 'Text';
    ws.cells.get('B1').value = 100;
    ws.cells.get('C1').value = true;

    const saver = new XmlSaver(wb);
    const buf = await saver.saveToBuffer();
    const zip = await unzipBuffer(buf);

    const sheetXml = await readEntry(zip, 'xl/worksheets/sheet1.xml');

    // Should contain sheetData with row/cell elements
    expect(sheetXml).toContain('<sheetData>');
    expect(sheetXml).toContain('r="A1"');
    expect(sheetXml).toContain('r="B1"');
    expect(sheetXml).toContain('r="C1"');

    // String cell should use shared string type t="s"
    expect(sheetXml).toContain('t="s"');

    // Number cell: value 100
    expect(sheetXml).toContain('<v>100</v>');

    // Boolean cell: value 1 (true) with t="b"
    expect(sheetXml).toContain('t="b"');
    expect(sheetXml).toContain('<v>1</v>');
  });

  it('handles multiple worksheets', async () => {
    const wb = new Workbook();
    wb.worksheets[0].cells.get('A1').value = 'Sheet1Data';
    const ws2 = wb.addWorksheet('Second');
    ws2.cells.get('A1').value = 'Sheet2Data';

    const saver = new XmlSaver(wb);
    const buf = await saver.saveToBuffer();
    const zip = await unzipBuffer(buf);

    // Both sheets present
    expect(zip.file('xl/worksheets/sheet1.xml')).not.toBeNull();
    expect(zip.file('xl/worksheets/sheet2.xml')).not.toBeNull();

    // Workbook XML should reference both sheets
    const wbXml = await readEntry(zip, 'xl/workbook.xml');
    expect(wbXml).toContain('name="Sheet1"');
    expect(wbXml).toContain('name="Second"');

    // Content types should have overrides for both
    const ctXml = await readEntry(zip, '[Content_Types].xml');
    expect(ctXml).toContain('/xl/worksheets/sheet1.xml');
    expect(ctXml).toContain('/xl/worksheets/sheet2.xml');
  });
});

// ─── 4. Workbook.saveToBuffer() ───────────────────────────────────────────

describe('Workbook.saveToBuffer()', () => {
  it('produces a valid XLSX buffer', async () => {
    const wb = new Workbook();
    wb.worksheets[0].cells.get('A1').value = 'via saveToBuffer';
    wb.worksheets[0].cells.get('A2').value = 12345;

    const buf = await wb.saveToBuffer();
    expect(buf).toBeInstanceOf(Buffer);

    // Verify it's a valid ZIP (PK magic bytes)
    expect(buf[0]).toBe(0x50); // 'P'
    expect(buf[1]).toBe(0x4B); // 'K'

    // Verify shared strings contain our text
    const zip = await unzipBuffer(buf);
    const sstXml = await readEntry(zip, 'xl/sharedStrings.xml');
    expect(sstXml).toContain('via saveToBuffer');
  });

  it('rejects unsupported save formats', async () => {
    const wb = new Workbook();
    await expect(wb.saveToBuffer(SaveFormat.CSV)).rejects.toThrow(
      'not yet supported',
    );
  });
});

// ─── 5. Workbook.save() – file output ─────────────────────────────────────

describe('Workbook.save() to file', () => {
  it('writes a valid .xlsx file to disk', async () => {
    const dir = tmpDir();
    const filePath = path.join(dir, 'output.xlsx');

    try {
      const wb = new Workbook();
      const ws = wb.worksheets[0];
      ws.cells.get('A1').value = 'File test';
      ws.cells.get('B1').value = 99.9;

      await wb.save(filePath);

      // File should exist and have non-zero size
      const stat = fs.statSync(filePath);
      expect(stat.size).toBeGreaterThan(0);

      // Read it back and verify ZIP structure
      const data = fs.readFileSync(filePath);
      const zip = await unzipBuffer(data);
      expect(zip.file('xl/workbook.xml')).not.toBeNull();
      expect(zip.file('xl/worksheets/sheet1.xml')).not.toBeNull();

      const sheetXml = await readEntry(zip, 'xl/worksheets/sheet1.xml');
      expect(sheetXml).toContain('<v>99.9</v>');
    } finally {
      // Cleanup
      if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
      fs.rmdirSync(dir);
    }
  });
});

// ─── 6. Styled cells ─────────────────────────────────────────────────────

describe('Styled cells in XLSX output', () => {
  it('writes styles XML with custom font and fill', async () => {
    const wb = new Workbook();
    const ws = wb.worksheets[0];

    const cell = ws.cells.get('A1');
    cell.value = 'Styled';
    cell.style.font.bold = true;
    cell.style.font.size = 14;
    cell.style.font.color = 'FFFF0000'; // red
    cell.style.fill.setSolidFill('FF00FF00'); // green background

    const buf = await wb.saveToBuffer();
    const zip = await unzipBuffer(buf);
    const stylesXml = await readEntry(zip, 'xl/styles.xml');

    // Should contain our custom font
    expect(stylesXml).toContain('<b/>');
    expect(stylesXml).toContain('val="14"'); // font size
    expect(stylesXml).toContain('FFFF0000'); // font color

    // Should contain our fill color
    expect(stylesXml).toContain('FF00FF00'); // fill foreground

    // Should have cellXfs with more than just the default style
    expect(stylesXml).toContain('<cellXfs');

    // Worksheet should reference style index on the cell
    const sheetXml = await readEntry(zip, 'xl/worksheets/sheet1.xml');
    expect(sheetXml).toMatch(/r="A1"[^>]*s="\d+"/);
  });
});

// ─── 7. Merged cells ─────────────────────────────────────────────────────

describe('Merged cells in XLSX output', () => {
  it('writes mergeCell elements', async () => {
    const wb = new Workbook();
    const ws = wb.worksheets[0];
    ws.cells.get('A1').value = 'Merged';
    ws.cells.merge(0, 0, 2, 3); // merge A1:C2 (0-based: row=0, col=0, totalRows=2, totalCols=3)

    const buf = await wb.saveToBuffer();
    const zip = await unzipBuffer(buf);
    const sheetXml = await readEntry(zip, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('<mergeCells');
    expect(sheetXml).toContain('ref="A1:C2"');
  });
});

// ─── 8. Freeze panes ─────────────────────────────────────────────────────

describe('Freeze panes in XLSX output', () => {
  it('writes pane element in sheetViews', async () => {
    const wb = new Workbook();
    const ws = wb.worksheets[0];
    ws.cells.get('A1').value = 'Header';
    ws.setFreezePane(1, 0); // freeze top row

    const buf = await wb.saveToBuffer();
    const zip = await unzipBuffer(buf);
    const sheetXml = await readEntry(zip, 'xl/worksheets/sheet1.xml');

    expect(sheetXml).toContain('<pane');
    expect(sheetXml).toContain('ySplit="1"');
    expect(sheetXml).toContain('state="frozen"');
  });
});
