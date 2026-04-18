/**
 * DrawingsLoader — parses xl/drawings/*.xml + xl/media/* and attaches
 * structured drawing metadata (pictures + shapes) to each Worksheet
 * via the `_drawings` field. Used by the HTML/PNG renderer to overlay
 * embedded images and DrawingML shapes on top of the cell grid.
 */

import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';
import { Worksheet } from '../core/Worksheet';

// Internal parser — keeps NS prefixes since drawing XML uses xdr:/a:/r: heavily
function createDrawingParser(): XMLParser {
  return new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    textNodeName: '#text',
    removeNSPrefix: false,
    parseTagValue: false,
    trimValues: false,
    processEntities: true,
  });
}

// ─── Public types ───────────────────────────────────────────────────────

export interface DrawingAnchor {
  fromCol: number;        // 0-based
  fromRow: number;        // 0-based
  fromColOff: number;     // EMU
  fromRowOff: number;     // EMU
  toCol?: number;
  toRow?: number;
  toColOff?: number;
  toRowOff?: number;
  cx?: number;            // EMU (oneCellAnchor / absoluteAnchor explicit extent)
  cy?: number;
}

export interface DrawingPicture {
  type: 'pic';
  anchor: DrawingAnchor;
  imageDataUri: string;   // data:image/...;base64,...
  rotation?: number;      // degrees (clockwise)
  flipH?: boolean;
  flipV?: boolean;
}

export interface ShapeRun {
  text: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  size?: number;          // pt
  color?: string;         // #rrggbb
  font?: string;
}

export interface ShapeParagraph {
  align?: 'left' | 'center' | 'right' | 'justify';
  runs: ShapeRun[];
}

export interface DrawingShape {
  type: 'sp';
  anchor: DrawingAnchor;
  geom: string;           // 'rect' | 'roundRect' | 'ellipse' | ...
  fillColor?: string;     // #rrggbb (solid only)
  strokeColor?: string;   // #rrggbb
  strokeWidth?: number;   // px
  rotation?: number;      // degrees
  paragraphs: ShapeParagraph[];
  vertAlign?: 'top' | 'middle' | 'bottom';
}

export type DrawingItem = DrawingPicture | DrawingShape;

// ─── Helpers ────────────────────────────────────────────────────────────

const EMU_PER_PX = 9525;

function ensureArray<T>(v: T | T[] | undefined | null): T[] {
  if (v == null) return [];
  return Array.isArray(v) ? v : [v];
}

function attr(el: any, name: string, def = ''): string {
  if (el == null) return def;
  const v = el[`@_${name}`];
  return v == null ? def : String(v);
}

function attrNum(el: any, name: string, def = 0): number {
  const v = attr(el, name, '');
  if (v === '') return def;
  const n = parseFloat(v);
  return isNaN(n) ? def : n;
}

function txt(node: any): string {
  if (node == null) return '';
  if (typeof node === 'string') return node;
  if (typeof node === 'number') return String(node);
  if (typeof node === 'object') {
    if (typeof node['#text'] !== 'undefined') return String(node['#text']);
  }
  return '';
}

function normalizePath(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  // Resolve relative against base directory
  const baseSegs = base.split('/').filter(Boolean);
  const tgtSegs = target.split('/');
  for (const s of tgtSegs) {
    if (s === '..') baseSegs.pop();
    else if (s === '.' || s === '') continue;
    else baseSegs.push(s);
  }
  return baseSegs.join('/');
}

function dirname(p: string): string {
  const i = p.lastIndexOf('/');
  return i >= 0 ? p.slice(0, i) : '';
}

function colorFromSolidFill(node: any): string | undefined {
  // <a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>
  if (!node) return undefined;
  if (node['a:srgbClr']) {
    const v = attr(node['a:srgbClr'], 'val', '');
    if (v) return '#' + v.toUpperCase();
  }
  if (node['a:schemeClr']) {
    // Best-effort: leave undefined (theme resolution out of scope here)
    return undefined;
  }
  if (node['a:sysClr']) {
    const last = attr(node['a:sysClr'], 'lastClr', '');
    if (last) return '#' + last.toUpperCase();
    const v = attr(node['a:sysClr'], 'val', '');
    if (v === 'window') return '#FFFFFF';
    if (v === 'windowText') return '#000000';
  }
  return undefined;
}

function mimeFromExt(filename: string): string {
  const ext = filename.toLowerCase().split('.').pop() || '';
  switch (ext) {
    case 'png': return 'image/png';
    case 'jpg':
    case 'jpeg': return 'image/jpeg';
    case 'gif': return 'image/gif';
    case 'bmp': return 'image/bmp';
    case 'svg': return 'image/svg+xml';
    case 'webp': return 'image/webp';
    case 'tif':
    case 'tiff': return 'image/tiff';
    case 'emf': return 'image/x-emf';
    case 'wmf': return 'image/x-wmf';
    default: return 'application/octet-stream';
  }
}

// ─── Anchor parsing ─────────────────────────────────────────────────────

function parseAnchorFrom(node: any): { col: number; row: number; colOff: number; rowOff: number } {
  return {
    col: parseInt(txt(node?.['xdr:col']) || '0', 10),
    row: parseInt(txt(node?.['xdr:row']) || '0', 10),
    colOff: parseInt(txt(node?.['xdr:colOff']) || '0', 10),
    rowOff: parseInt(txt(node?.['xdr:rowOff']) || '0', 10),
  };
}

function parseAnchorNode(anchor: any, kind: 'one' | 'two' | 'abs'): DrawingAnchor {
  if (kind === 'two') {
    const f = parseAnchorFrom(anchor['xdr:from']);
    const t = parseAnchorFrom(anchor['xdr:to']);
    return {
      fromCol: f.col, fromRow: f.row,
      fromColOff: f.colOff, fromRowOff: f.rowOff,
      toCol: t.col, toRow: t.row,
      toColOff: t.colOff, toRowOff: t.rowOff,
    };
  }
  if (kind === 'one') {
    const f = parseAnchorFrom(anchor['xdr:from']);
    const ext = anchor['xdr:ext'];
    return {
      fromCol: f.col, fromRow: f.row,
      fromColOff: f.colOff, fromRowOff: f.rowOff,
      cx: ext ? attrNum(ext, 'cx', 0) : undefined,
      cy: ext ? attrNum(ext, 'cy', 0) : undefined,
    };
  }
  // absolute
  const pos = anchor['xdr:pos'];
  const ext = anchor['xdr:ext'];
  return {
    fromCol: 0, fromRow: 0,
    fromColOff: pos ? attrNum(pos, 'x', 0) : 0,
    fromRowOff: pos ? attrNum(pos, 'y', 0) : 0,
    cx: ext ? attrNum(ext, 'cx', 0) : undefined,
    cy: ext ? attrNum(ext, 'cy', 0) : undefined,
  };
}

// ─── Shape parsing ──────────────────────────────────────────────────────

function parseTxBody(txBody: any): { paragraphs: ShapeParagraph[]; vertAlign?: DrawingShape['vertAlign'] } {
  const paragraphs: ShapeParagraph[] = [];
  if (!txBody) return { paragraphs };

  const bodyPr = txBody['a:bodyPr'];
  let vertAlign: DrawingShape['vertAlign'] | undefined;
  if (bodyPr) {
    const anchor = attr(bodyPr, 'anchor', '');
    if (anchor === 't') vertAlign = 'top';
    else if (anchor === 'ctr') vertAlign = 'middle';
    else if (anchor === 'b') vertAlign = 'bottom';
  }

  const ps = ensureArray(txBody['a:p']);
  for (const p of ps) {
    const pPr = p['a:pPr'];
    let align: ShapeParagraph['align'] | undefined;
    if (pPr) {
      const a = attr(pPr, 'algn', '');
      if (a === 'l') align = 'left';
      else if (a === 'ctr') align = 'center';
      else if (a === 'r') align = 'right';
      else if (a === 'just') align = 'justify';
    }

    const runs: ShapeRun[] = [];
    const rs = ensureArray(p['a:r']);
    for (const r of rs) {
      const rPr = r['a:rPr'];
      const tNode = r['a:t'];
      const text = typeof tNode === 'string' ? tNode : (tNode != null ? txt(tNode) : '');
      const run: ShapeRun = { text };
      if (rPr) {
        if (attr(rPr, 'b', '') === '1') run.bold = true;
        if (attr(rPr, 'i', '') === '1') run.italic = true;
        if (attr(rPr, 'strike', '') === 'sngStrike') run.strike = true;
        const u = attr(rPr, 'u', '');
        if (u && u !== 'none') run.underline = true;
        const sz = attrNum(rPr, 'sz', 0);
        if (sz > 0) run.size = sz / 100;  // OOXML stores pt*100
        const c = colorFromSolidFill(rPr['a:solidFill']);
        if (c) run.color = c;
        const latin = rPr['a:latin'];
        const ea = rPr['a:ea'];
        const fontTypeface = (latin && attr(latin, 'typeface', '')) || (ea && attr(ea, 'typeface', '')) || '';
        if (fontTypeface) run.font = fontTypeface;
      }
      if (run.text) runs.push(run);
    }
    // Plain <a:fld> (field) or empty paragraphs => still preserve as empty line
    if (runs.length === 0 && p['a:fld']) {
      const flds = ensureArray(p['a:fld']);
      for (const f of flds) {
        const t = f['a:t'];
        const tx = typeof t === 'string' ? t : txt(t);
        if (tx) runs.push({ text: tx });
      }
    }
    paragraphs.push({ align, runs });
  }
  return { paragraphs, vertAlign };
}

function parseShape(sp: any, anchor: DrawingAnchor): DrawingShape {
  const spPr = sp['xdr:spPr'] || sp.spPr || {};
  const xfrm = spPr['a:xfrm'];
  let rotation = 0;
  if (xfrm) {
    const rot = attrNum(xfrm, 'rot', 0);
    if (rot) rotation = rot / 60000; // OOXML rotations: 60000ths of a degree
  }

  const prstGeom = spPr['a:prstGeom'];
  const geom = prstGeom ? attr(prstGeom, 'prst', 'rect') : 'rect';

  const fillColor = colorFromSolidFill(spPr['a:solidFill']);

  let strokeColor: string | undefined;
  let strokeWidth: number | undefined;
  const ln = spPr['a:ln'];
  if (ln) {
    const wEmu = attrNum(ln, 'w', 0);
    if (wEmu > 0) strokeWidth = Math.max(1, Math.round(wEmu / EMU_PER_PX));
    const lnFill = colorFromSolidFill(ln['a:solidFill']);
    if (lnFill) strokeColor = lnFill;
    if (ln['a:noFill']) strokeColor = undefined;
  }

  const txBodyNode = sp['xdr:txBody'] || sp.txBody;
  const { paragraphs, vertAlign } = parseTxBody(txBodyNode);

  return {
    type: 'sp',
    anchor,
    geom,
    fillColor,
    strokeColor,
    strokeWidth,
    rotation: rotation || undefined,
    paragraphs,
    vertAlign,
  };
}

function parsePicture(pic: any, anchor: DrawingAnchor, mediaMap: Map<string, string>): DrawingPicture | null {
  const blipFill = pic['xdr:blipFill'] || pic.blipFill;
  if (!blipFill) return null;
  const blip = blipFill['a:blip'];
  if (!blip) return null;
  const embedId = attr(blip, 'r:embed', '') || attr(blip, 'embed', '');
  if (!embedId) return null;
  const dataUri = mediaMap.get(embedId);
  if (!dataUri) return null;

  const spPr = pic['xdr:spPr'] || pic.spPr || {};
  const xfrm = spPr['a:xfrm'];
  let rotation = 0;
  let flipH = false;
  let flipV = false;
  if (xfrm) {
    const rot = attrNum(xfrm, 'rot', 0);
    if (rot) rotation = rot / 60000;
    if (attr(xfrm, 'flipH', '') === '1') flipH = true;
    if (attr(xfrm, 'flipV', '') === '1') flipV = true;
  }

  return {
    type: 'pic',
    anchor,
    imageDataUri: dataUri,
    rotation: rotation || undefined,
    flipH: flipH || undefined,
    flipV: flipV || undefined,
  };
}

// ─── Main loader ────────────────────────────────────────────────────────

const DRAWING_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

async function readRels(zip: JSZip, parser: XMLParser, relsPath: string): Promise<Map<string, { type: string; target: string }>> {
  const out = new Map<string, { type: string; target: string }>();
  const file = zip.file(relsPath);
  if (!file) return out;
  try {
    const buf = await file.async('nodebuffer');
    const root = parser.parse(buf.toString('utf-8'));
    const rels = ensureArray(root?.Relationships?.Relationship ?? root?.Relationship);
    for (const r of rels) {
      const id = attr(r, 'Id', '');
      const type = attr(r, 'Type', '');
      const target = attr(r, 'Target', '');
      if (id) out.set(id, { type, target });
    }
  } catch {
    // ignore
  }
  return out;
}

export async function loadWorksheetDrawings(
  zip: JSZip,
  _parserUnused: XMLParser | null,
  worksheet: Worksheet,
  sheetNum: number,
): Promise<void> {
  // Use our own parser to preserve xdr:/a:/r: namespace prefixes
  // (the main XmlLoader strips namespaces, which would lose drawing element identities)
  const parser = createDrawingParser();
  // 1. Find drawing rel for sheet
  const sheetRelsPath = `xl/worksheets/_rels/sheet${sheetNum}.xml.rels`;
  const sheetRels = await readRels(zip, parser, sheetRelsPath);

  const drawingRels: { id: string; target: string }[] = [];
  for (const [id, rel] of sheetRels) {
    if (rel.type === DRAWING_REL_TYPE) {
      drawingRels.push({ id, target: rel.target });
    }
  }
  if (drawingRels.length === 0) return;

  const drawings: DrawingItem[] = [];

  for (const dr of drawingRels) {
    const drawingPath = normalizePath('xl/worksheets/', dr.target);
    const drawingFile = zip.file(drawingPath);
    if (!drawingFile) continue;

    // Load drawing rels (image targets)
    const drDir = dirname(drawingPath);
    const drRelsPath = `${drDir}/_rels/${drawingPath.split('/').pop()}.rels`;
    const drRels = await readRels(zip, parser, drRelsPath);

    // Build map relId -> dataURI
    const mediaMap = new Map<string, string>();
    for (const [id, rel] of drRels) {
      if (rel.type !== IMAGE_REL_TYPE) continue;
      const mediaPath = normalizePath(drDir + '/', rel.target);
      const mediaFile = zip.file(mediaPath);
      if (!mediaFile) continue;
      try {
        const bytes = await mediaFile.async('nodebuffer');
        const mime = mimeFromExt(mediaPath);
        const b64 = bytes.toString('base64');
        mediaMap.set(id, `data:${mime};base64,${b64}`);
      } catch {
        // skip broken media
      }
    }

    // Parse drawing XML
    let drawingXml: any;
    try {
      const buf = await drawingFile.async('nodebuffer');
      drawingXml = parser.parse(buf.toString('utf-8'));
    } catch {
      continue;
    }

    const wsDr = drawingXml['xdr:wsDr'] || drawingXml.wsDr;
    if (!wsDr) continue;

    const anchorKinds: { key: string; kind: 'one' | 'two' | 'abs' }[] = [
      { key: 'xdr:twoCellAnchor', kind: 'two' },
      { key: 'xdr:oneCellAnchor', kind: 'one' },
      { key: 'xdr:absoluteAnchor', kind: 'abs' },
    ];

    for (const { key, kind } of anchorKinds) {
      const anchors = ensureArray(wsDr[key]);
      for (const a of anchors) {
        const anchor = parseAnchorNode(a, kind);

        // Picture child
        const picNode = a['xdr:pic'];
        if (picNode) {
          const pic = parsePicture(picNode, anchor, mediaMap);
          if (pic) drawings.push(pic);
          continue;
        }

        // Shape child
        const spNode = a['xdr:sp'];
        if (spNode) {
          drawings.push(parseShape(spNode, anchor));
          continue;
        }

        // Group shapes (xdr:grpSp): flatten one level
        const grp = a['xdr:grpSp'];
        if (grp) {
          const groupPics = ensureArray(grp['xdr:pic']);
          for (const gp of groupPics) {
            const gpItem = parsePicture(gp, anchor, mediaMap);
            if (gpItem) drawings.push(gpItem);
          }
          const groupSps = ensureArray(grp['xdr:sp']);
          for (const gs of groupSps) {
            drawings.push(parseShape(gs, anchor));
          }
        }
      }
    }
  }

  if (drawings.length > 0) {
    (worksheet as any)._drawings = drawings;
  }
}
