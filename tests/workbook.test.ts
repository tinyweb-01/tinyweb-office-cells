/**
 * Unit tests for Workbook, SaveFormat, WorkbookProtection,
 * WorkbookProperties, DocumentProperties.
 */
import {
  Workbook,
  SaveFormat,
  saveFormatFromExtension,
  save_format_from_extension,
  WorkbookProtection,
  WorkbookProperties,
  DocumentProperties,
} from '../src/core/Workbook';
import { Worksheet } from '../src/core/Worksheet';

// ─── SaveFormat ──────────────────────────────────────────────────────────

describe('SaveFormat', () => {
  it('has expected string values', () => {
    expect(SaveFormat.AUTO).toBe('auto');
    expect(SaveFormat.XLSX).toBe('xlsx');
    expect(SaveFormat.CSV).toBe('csv');
    expect(SaveFormat.TSV).toBe('tsv');
    expect(SaveFormat.MARKDOWN).toBe('markdown');
    expect(SaveFormat.JSON).toBe('json');
  });
});

describe('saveFormatFromExtension', () => {
  it('maps known extensions', () => {
    expect(saveFormatFromExtension('file.xlsx')).toBe(SaveFormat.XLSX);
    expect(saveFormatFromExtension('data.csv')).toBe(SaveFormat.CSV);
    expect(saveFormatFromExtension('out.tsv')).toBe(SaveFormat.TSV);
    expect(saveFormatFromExtension('readme.md')).toBe(SaveFormat.MARKDOWN);
    expect(saveFormatFromExtension('readme.markdown')).toBe(SaveFormat.MARKDOWN);
    expect(saveFormatFromExtension('data.json')).toBe(SaveFormat.JSON);
    expect(saveFormatFromExtension('book.xlsm')).toBe(SaveFormat.XLSX);
  });

  it('is case-insensitive', () => {
    expect(saveFormatFromExtension('FILE.XLSX')).toBe(SaveFormat.XLSX);
    expect(saveFormatFromExtension('data.CSV')).toBe(SaveFormat.CSV);
  });

  it('throws on unsupported extension', () => {
    expect(() => saveFormatFromExtension('file.pdf')).toThrow('Unsupported file extension');
  });

  it('throws on no extension', () => {
    expect(() => saveFormatFromExtension('noext')).toThrow('Unsupported file extension');
  });

  it('snake_case alias works', () => {
    expect(save_format_from_extension('test.xlsx')).toBe(SaveFormat.XLSX);
  });
});

// ─── WorkbookProtection ──────────────────────────────────────────────────

describe('WorkbookProtection', () => {
  it('has correct defaults', () => {
    const p = new WorkbookProtection();
    expect(p.lockStructure).toBe(false);
    expect(p.lockWindows).toBe(false);
    expect(p.workbookPassword).toBeNull();
  });

  it('supports snake_case aliases', () => {
    const p = new WorkbookProtection();
    p.lock_structure = true;
    expect(p.lockStructure).toBe(true);
    expect(p.lock_structure).toBe(true);

    p.lock_windows = true;
    expect(p.lockWindows).toBe(true);

    p.workbook_password = 'secret';
    expect(p.workbookPassword).toBe('secret');
    expect(p.workbook_password).toBe('secret');
  });
});

// ─── WorkbookProperties ──────────────────────────────────────────────────

describe('WorkbookProperties', () => {
  it('has correct defaults', () => {
    const props = new WorkbookProperties();
    expect(props.calcMode).toBe('auto');
    expect(props.calc_mode).toBe('auto');
    expect(props.definedNames.count).toBe(0);
    expect(props.defined_names.count).toBe(0);
    expect(props.protection).toBeInstanceOf(WorkbookProtection);
  });

  it('view has activeTab defaulting to 0', () => {
    const props = new WorkbookProperties();
    expect(props.view.activeTab).toBe(0);
    expect(props.view.active_tab).toBe(0);
  });

  it('view.active_tab alias works for setting', () => {
    const props = new WorkbookProperties();
    props.view.active_tab = 3;
    expect(props.view.activeTab).toBe(3);
  });
});

// ─── DocumentProperties ──────────────────────────────────────────────────

describe('DocumentProperties', () => {
  it('has correct defaults', () => {
    const dp = new DocumentProperties();
    expect(dp.title).toBe('');
    expect(dp.subject).toBe('');
    expect(dp.creator).toBe('');
    expect(dp.keywords).toBe('');
    expect(dp.description).toBe('');
    expect(dp.lastModifiedBy).toBe('');
    expect(dp.category).toBe('');
    expect(dp._coreXml).toBeNull();
    expect(dp._appXml).toBeNull();
  });

  it('supports snake_case aliases', () => {
    const dp = new DocumentProperties();
    dp.last_modified_by = 'Alice';
    expect(dp.lastModifiedBy).toBe('Alice');
    expect(dp.last_modified_by).toBe('Alice');
  });
});

// ─── Workbook ────────────────────────────────────────────────────────────

describe('Workbook', () => {
  // ── Construction ─────────────────────────────────────────────────────
  it('creates with one default worksheet', () => {
    const wb = new Workbook();
    expect(wb.worksheets).toHaveLength(1);
    expect(wb.worksheets[0].name).toBe('Sheet1');
    expect(wb.filePath).toBeNull();
    expect(wb.file_path).toBeNull();
    expect(wb.styles).toHaveLength(1); // default style
    expect(wb.sharedStrings).toEqual([]);
    expect(wb.shared_strings).toEqual([]);
  });

  it('stores file path from constructor', () => {
    const wb = new Workbook('test.xlsx');
    expect(wb.filePath).toBe('test.xlsx');
  });

  it('has workbook properties', () => {
    const wb = new Workbook();
    expect(wb.properties).toBeInstanceOf(WorkbookProperties);
    expect(wb.properties.view.activeTab).toBe(0);
  });

  it('has document properties (lazily created)', () => {
    const wb = new Workbook();
    expect(wb.documentProperties).toBeInstanceOf(DocumentProperties);
    expect(wb.document_properties).toBe(wb.documentProperties); // same instance
  });

  // ── Worksheet management ────────────────────────────────────────────
  describe('addWorksheet', () => {
    it('adds a named worksheet', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('Data');
      expect(ws.name).toBe('Data');
      expect(wb.worksheets).toHaveLength(2);
      expect(wb.worksheets[1]).toBe(ws);
    });

    it('auto-generates unique name', () => {
      const wb = new Workbook(); // has Sheet1
      const ws2 = wb.addWorksheet(); // Sheet2
      expect(ws2.name).toBe('Sheet2');
      const ws3 = wb.add_worksheet(); // Sheet3
      expect(ws3.name).toBe('Sheet3');
    });

    it('skips existing sheet names when auto-generating', () => {
      const wb = new Workbook(); // Sheet1
      wb.addWorksheet('Sheet2');
      const ws = wb.addWorksheet(); // should be Sheet3
      expect(ws.name).toBe('Sheet3');
    });

    it('createWorksheet alias works', () => {
      const wb = new Workbook();
      const ws = wb.createWorksheet('Alt');
      expect(ws.name).toBe('Alt');
    });

    it('create_worksheet alias works', () => {
      const wb = new Workbook();
      const ws = wb.create_worksheet('Alt2');
      expect(ws.name).toBe('Alt2');
    });
  });

  describe('getWorksheet', () => {
    it('gets by index', () => {
      const wb = new Workbook();
      wb.addWorksheet('Extra');
      expect(wb.getWorksheet(0).name).toBe('Sheet1');
      expect(wb.getWorksheet(1).name).toBe('Extra');
    });

    it('gets by name', () => {
      const wb = new Workbook();
      wb.addWorksheet('Sales');
      expect(wb.getWorksheet('Sales').name).toBe('Sales');
    });

    it('throws RangeError for invalid index', () => {
      const wb = new Workbook();
      expect(() => wb.getWorksheet(5)).toThrow(RangeError);
      expect(() => wb.getWorksheet(-1)).toThrow(RangeError);
    });

    it('throws Error for unknown name', () => {
      const wb = new Workbook();
      expect(() => wb.getWorksheet('NotExist')).toThrow("Worksheet 'NotExist' not found");
    });

    it('get_worksheet alias works', () => {
      const wb = new Workbook();
      expect(wb.get_worksheet(0).name).toBe('Sheet1');
    });
  });

  describe('getWorksheetByName / getWorksheetByIndex', () => {
    it('getWorksheetByName returns worksheet or null', () => {
      const wb = new Workbook();
      expect(wb.getWorksheetByName('Sheet1')?.name).toBe('Sheet1');
      expect(wb.getWorksheetByName('X')).toBeNull();
      expect(wb.get_worksheet_by_name('Sheet1')?.name).toBe('Sheet1');
    });

    it('getWorksheetByIndex returns worksheet or null', () => {
      const wb = new Workbook();
      expect(wb.getWorksheetByIndex(0)?.name).toBe('Sheet1');
      expect(wb.getWorksheetByIndex(10)).toBeNull();
      expect(wb.getWorksheetByIndex(-1)).toBeNull();
      expect(wb.get_worksheet_by_index(0)?.name).toBe('Sheet1');
    });
  });

  describe('removeWorksheet', () => {
    it('removes by index', () => {
      const wb = new Workbook();
      wb.addWorksheet('Extra');
      wb.removeWorksheet(1);
      expect(wb.worksheets).toHaveLength(1);
      expect(wb.worksheets[0].name).toBe('Sheet1');
    });

    it('removes by name', () => {
      const wb = new Workbook();
      wb.addWorksheet('Extra');
      wb.removeWorksheet('Extra');
      expect(wb.worksheets).toHaveLength(1);
    });

    it('removes by reference', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('Extra');
      wb.removeWorksheet(ws);
      expect(wb.worksheets).toHaveLength(1);
    });

    it('throws RangeError for invalid index', () => {
      const wb = new Workbook();
      expect(() => wb.removeWorksheet(5)).toThrow(RangeError);
    });

    it('throws Error for unknown name', () => {
      const wb = new Workbook();
      expect(() => wb.removeWorksheet('Nope')).toThrow("Worksheet 'Nope' not found");
    });

    it('remove_worksheet alias works', () => {
      const wb = new Workbook();
      wb.addWorksheet('Extra');
      wb.remove_worksheet('Extra');
      expect(wb.worksheets).toHaveLength(1);
    });
  });

  describe('active worksheet', () => {
    it('returns first worksheet by default', () => {
      const wb = new Workbook();
      const active = wb.getActiveWorksheet();
      expect(active?.name).toBe('Sheet1');
    });

    it('setActiveWorksheet by index', () => {
      const wb = new Workbook();
      wb.addWorksheet('Second');
      wb.setActiveWorksheet(1);
      expect(wb.getActiveWorksheet()?.name).toBe('Second');
    });

    it('setActiveWorksheet by name', () => {
      const wb = new Workbook();
      wb.addWorksheet('Target');
      wb.setActiveWorksheet('Target');
      expect(wb.getActiveWorksheet()?.name).toBe('Target');
    });

    it('setActiveWorksheet by Worksheet reference', () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('Ref');
      wb.setActiveWorksheet(ws);
      expect(wb.getActiveWorksheet()?.name).toBe('Ref');
    });

    it('snake_case aliases work', () => {
      const wb = new Workbook();
      wb.addWorksheet('Alt');
      wb.set_active_worksheet('Alt');
      expect(wb.get_active_worksheet()?.name).toBe('Alt');
    });

    it('returns null for empty workbook', () => {
      const wb = new Workbook();
      wb.removeWorksheet(0);
      expect(wb.getActiveWorksheet()).toBeNull();
    });
  });

  describe('copyWorksheet', () => {
    it('copies by name', () => {
      const wb = new Workbook();
      wb.worksheets[0].cells.get('A1').value = 'data';
      const copy = wb.copyWorksheet('Sheet1');
      expect(copy).not.toBeNull();
      expect(copy!.name).toBe('Sheet1 (copy)');
      expect(copy!.cells.get('A1').value).toBe('data');
      expect(wb.worksheets).toHaveLength(2);
    });

    it('copies by index', () => {
      const wb = new Workbook();
      const copy = wb.copyWorksheet(0);
      expect(copy).not.toBeNull();
      expect(copy!.name).toBe('Sheet1 (copy)');
    });

    it('copies by Worksheet reference', () => {
      const wb = new Workbook();
      const ws = wb.worksheets[0];
      const copy = wb.copyWorksheet(ws);
      expect(copy).not.toBeNull();
    });

    it('generates unique copy names', () => {
      const wb = new Workbook();
      wb.copyWorksheet(0); // Sheet1 (copy)
      const copy2 = wb.copyWorksheet(0); // Sheet1 (copy1)
      expect(copy2!.name).toBe('Sheet1 (copy1)');
    });

    it('returns null for invalid source', () => {
      const wb = new Workbook();
      expect(wb.copyWorksheet(10)).toBeNull();
      expect(wb.copyWorksheet('NonExistent')).toBeNull();
    });

    it('copy_worksheet alias works', () => {
      const wb = new Workbook();
      const copy = wb.copy_worksheet(0);
      expect(copy).not.toBeNull();
    });
  });

  // ── Protection ──────────────────────────────────────────────────────
  describe('protection', () => {
    it('protect() enables structure lock by default', () => {
      const wb = new Workbook();
      expect(wb.isProtected()).toBe(false);
      expect(wb.is_protected()).toBe(false);

      wb.protect('pass');
      expect(wb.isProtected()).toBe(true);
      expect(wb.protection.lockStructure).toBe(true);
      expect(wb.protection.lockWindows).toBe(false);
      expect(wb.protection.password).toBe('pass');
    });

    it('protect() with custom flags', () => {
      const wb = new Workbook();
      wb.protect('pw', false, true);
      expect(wb.protection.lockStructure).toBe(false);
      expect(wb.protection.lockWindows).toBe(true);
    });

    it('unprotect() clears protection', () => {
      const wb = new Workbook();
      wb.protect('pass', true, true);
      wb.unprotect();
      expect(wb.isProtected()).toBe(false);
      expect(wb.protection.lockStructure).toBe(false);
      expect(wb.protection.lockWindows).toBe(false);
      expect(wb.protection.password).toBeNull();
    });
  });

  // ── File I/O ──────────────────────────────────────────────────────
  describe('save / load', () => {
    it('save() rejects unsupported format', async () => {
      const wb = new Workbook();
      await expect(wb.save('out.csv', SaveFormat.CSV)).rejects.toThrow('not yet supported');
    });

    it('Workbook.load() rejects for non-existent file', async () => {
      await expect(Workbook.load('non_existent_file.xlsx')).rejects.toThrow();
    });
  });

  // ── toString ────────────────────────────────────────────────────────
  it('toString() reflects worksheet count', () => {
    const wb = new Workbook();
    expect(wb.toString()).toBe('Workbook(worksheets=1)');
    wb.addWorksheet('Extra');
    expect(wb.toString()).toBe('Workbook(worksheets=2)');
  });

  // ── Integration: worksheet cells ────────────────────────────────────
  it('can write/read cell data via worksheet', () => {
    const wb = new Workbook();
    const ws = wb.worksheets[0];
    ws.cells.get('A1').value = 'Hello';
    ws.cells.get('B2').value = 42;
    expect(ws.cells.get('A1').value).toBe('Hello');
    expect(ws.cells.get('B2').value).toBe(42);
  });
});
