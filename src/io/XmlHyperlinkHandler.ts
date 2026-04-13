/**
 * XML Hyperlink Handler
 *
 * Loads and saves hyperlinks from/to XLSX worksheet XML and .rels files.
 */

import { Hyperlink, HyperlinkCollection } from '../features/Hyperlink';

function ensureArray<T>(val: T | T[] | undefined | null): T[] {
  if (val == null) return [];
  return Array.isArray(val) ? val : [val];
}

function attr(elem: any, name: string, defaultVal = ''): string {
  const v = elem?.['@_' + name];
  return v != null ? String(v) : defaultVal;
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// ==================== Loader ====================

export class HyperlinkXmlLoader {
  loadHyperlinks(
    wsRoot: any,
    collection: HyperlinkCollection,
    relationships: Map<string, string>,
  ): void {
    const hlElems = ensureArray(wsRoot?.hyperlinks?.hyperlink);

    for (const hlElem of hlElems) {
      this._loadHyperlink(hlElem, collection, relationships);
    }
  }

  loadRelationships(relsXml: any): Map<string, string> {
    const rels = new Map<string, string>();
    if (!relsXml) return rels;

    const relationships = ensureArray(relsXml?.Relationships?.Relationship);
    for (const rel of relationships) {
      const id = attr(rel, 'Id', '');
      const target = attr(rel, 'Target', '');
      const type = attr(rel, 'Type', '');
      if (id && type.includes('hyperlink')) {
        rels.set(id, target);
      }
    }
    return rels;
  }

  private _loadHyperlink(
    elem: any,
    collection: HyperlinkCollection,
    relationships: Map<string, string>,
  ): void {
    const ref = attr(elem, 'ref', '');
    if (!ref) return;

    let address = '';
    let subAddress = '';
    const display = attr(elem, 'display', '');
    const tooltip = attr(elem, 'tooltip', '');

    // Check for relationship-based hyperlink (external)
    const rId = attr(elem, 'r:id', '') || attr(elem, 'id', '');
    if (rId && relationships.has(rId)) {
      address = relationships.get(rId)!;
    }

    // Internal location
    const location = attr(elem, 'location', '');
    if (location) {
      subAddress = location;
    }

    const hl = new Hyperlink(ref, address, subAddress, display, tooltip);
    collection.addHyperlink(hl);
  }
}

// ==================== Saver ====================

export interface HyperlinkRelationship {
  rId: string;
  target: string;
}

export class HyperlinkXmlSaver {
  private _relIdCounter = 1;

  formatHyperlinksXml(collection: HyperlinkCollection): string {
    if (collection.count === 0) return '';

    const lines: string[] = ['<hyperlinks>'];

    for (const hl of collection) {
      if (hl.isDeleted) continue;
      lines.push(this._formatHyperlinkXml(hl));
    }

    lines.push('</hyperlinks>');
    return lines.join('');
  }

  private _formatHyperlinkXml(hl: Hyperlink): string {
    const attrs: string[] = [`ref="${escapeXml(hl.range)}"`];

    if (hl.type === 'external' || hl.type === 'email') {
      const rId = `rId${this._relIdCounter++}`;
      attrs.push(`r:id="${rId}"`);
    }

    if (hl.subAddress) {
      attrs.push(`location="${escapeXml(hl.subAddress)}"`);
    }

    if (hl.textToDisplay) {
      attrs.push(`display="${escapeXml(hl.textToDisplay)}"`);
    }

    if (hl.screenTip) {
      attrs.push(`tooltip="${escapeXml(hl.screenTip)}"`);
    }

    return `<hyperlink ${attrs.join(' ')}/>`;
  }

  getHyperlinkRelationships(collection: HyperlinkCollection): HyperlinkRelationship[] {
    const rels: HyperlinkRelationship[] = [];
    // Re-use the same starting counter that was set via resetRelationshipCounter
    // so that rIds in .rels match those written in the worksheet XML
    let relId = this._relIdCounter;

    // Reset counter for next use since formatHyperlinksXml already advanced it
    // We need to rebuild from the same starting point
    // Actually, _relIdCounter was already advanced by formatHyperlinksXml,
    // so we compute the starting point from the current value minus the count of external links
    let externalCount = 0;
    for (const hl of collection) {
      if (hl.isDeleted) continue;
      if (hl.type === 'external' || hl.type === 'email') externalCount++;
    }
    const startRelId = this._relIdCounter - externalCount;

    let rid = startRelId;
    for (const hl of collection) {
      if (hl.isDeleted) continue;
      if (hl.type === 'external' || hl.type === 'email') {
        rels.push({
          rId: `rId${rid++}`,
          target: hl.address,
        });
      }
    }

    return rels;
  }

  resetRelationshipCounter(startRelId = 1): void {
    this._relIdCounter = startRelId;
  }
}

// ==================== Relationship Writer ====================

export class HyperlinkRelationshipWriter {
  static formatRelationshipsXml(
    relationships: HyperlinkRelationship[],
    existingRels: string[] = [],
  ): string {
    const lines: string[] = [];

    // Include existing relationships
    for (const rel of existingRels) {
      lines.push(rel);
    }

    // Add hyperlink relationships
    for (const rel of relationships) {
      lines.push(
        `<Relationship Id="${rel.rId}" ` +
        `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" ` +
        `Target="${escapeXml(rel.target)}" TargetMode="External"/>`,
      );
    }

    return lines.join('');
  }
}
