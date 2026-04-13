/**
 * Hyperlink Module
 *
 * Provides classes for Excel hyperlinks according to ECMA-376 specification.
 */

export type HyperlinkType = 'external' | 'internal' | 'email' | 'unknown';

export class Hyperlink {
  private _range: string;
  private _address: string;
  private _subAddress: string;
  private _textToDisplay: string;
  private _screenTip: string;
  private _deleted = false;

  constructor(
    rangeAddress: string,
    address = '',
    subAddress = '',
    textToDisplay = '',
    screenTip = '',
  ) {
    this._range = rangeAddress;
    this._address = address;
    this._subAddress = subAddress;
    this._textToDisplay = textToDisplay;
    this._screenTip = screenTip;
  }

  get range(): string { return this._range; }
  set range(value: string) { this._range = value; }

  get address(): string { return this._address; }
  set address(value: string) { this._address = value; }

  get subAddress(): string { return this._subAddress; }
  set subAddress(value: string) { this._subAddress = value; }

  get textToDisplay(): string { return this._textToDisplay; }
  set textToDisplay(value: string) { this._textToDisplay = value; }

  get screenTip(): string { return this._screenTip; }
  set screenTip(value: string) { this._screenTip = value; }

  get type(): HyperlinkType {
    if (this._subAddress && !this._address) return 'internal';
    if (this._address) {
      if (this._address.startsWith('mailto:')) return 'email';
      return 'external';
    }
    return 'unknown';
  }

  delete(): void {
    this._deleted = true;
  }

  get isDeleted(): boolean {
    return this._deleted;
  }
}

export class HyperlinkCollection {
  private _hyperlinks: Hyperlink[] = [];

  add(
    rangeAddress: string,
    address = '',
    subAddress = '',
    textToDisplay = '',
    screenTip = '',
  ): Hyperlink {
    // Validate range
    if (!rangeAddress || !rangeAddress.match(/^[A-Z]+\d+(:[A-Z]+\d+)?$/i)) {
      // Allow but don't strictly validate
    }

    const hl = new Hyperlink(rangeAddress, address, subAddress, textToDisplay, screenTip);
    this._hyperlinks.push(hl);
    return hl;
  }

  /** Add a pre-built hyperlink object */
  addHyperlink(hl: Hyperlink): void {
    this._hyperlinks.push(hl);
  }

  delete(indexOrHyperlink?: number | Hyperlink): void {
    if (typeof indexOrHyperlink === 'number') {
      if (indexOrHyperlink >= 0 && indexOrHyperlink < this._hyperlinks.length) {
        this._hyperlinks.splice(indexOrHyperlink, 1);
      }
    } else if (indexOrHyperlink instanceof Hyperlink) {
      const idx = this._hyperlinks.indexOf(indexOrHyperlink);
      if (idx >= 0) this._hyperlinks.splice(idx, 1);
    }
  }

  clear(): void {
    this._hyperlinks = [];
  }

  get count(): number {
    return this._hyperlinks.length;
  }

  get(index: number): Hyperlink {
    if (index < 0 || index >= this._hyperlinks.length) {
      throw new Error(`Index ${index} out of range`);
    }
    return this._hyperlinks[index];
  }

  [Symbol.iterator](): Iterator<Hyperlink> {
    let i = 0;
    const items = this._hyperlinks;
    return {
      next(): IteratorResult<Hyperlink> {
        if (i < items.length) {
          return { value: items[i++], done: false };
        }
        return { value: undefined as any, done: true };
      },
    };
  }
}
