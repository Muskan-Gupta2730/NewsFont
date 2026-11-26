import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

export interface INewsFontApplicationCustomizerProperties {
  headingColor: string;
}

interface IRgba {
  r: number;
  g: number;
  b: number;
  a: number;
}

// The out-of-box News web part
const NEWS_WEBPART: string = '[data-sp-web-part-id="8c88f208-6c77-4bdb-86a0-0c47b4316588"]';
const WHITE: string = '#ffffff';

export default class NewsFontApplicationCustomizer
  extends BaseApplicationCustomizer<INewsFontApplicationCustomizerProperties> {

  private _color: string = '#120052';
  private _pending: boolean = false;
  private _observer: MutationObserver | undefined;

  public onInit(): Promise<void> {
    this._color = this.properties.headingColor || '#120052';

    this._applyColors();

    // News loads late and pages change without reload, so keep watching
    this._observer = new MutationObserver(() => this._schedule());
    this._observer.observe(document.body, { childList: true, subtree: true });

    return Promise.resolve();
  }

  public onDispose(): void {
    if (this._observer) {
      this._observer.disconnect();
    }
  }

  // Wait 200ms so we don't run hundreds of times while the page loads
  private _schedule(): void {
    if (this._pending) {
      return;
    }
    this._pending = true;
    window.setTimeout(() => {
      this._pending = false;
      this._applyColors();
    }, 200);
  }

  private _applyColors(): void {
    const webParts: NodeListOf<HTMLElement> = document.querySelectorAll<HTMLElement>(NEWS_WEBPART);

    for (let i: number = 0; i < webParts.length; i++) {
      const wp: HTMLElement = webParts[i];

      // A) Web part title ("Latest Sirva Life stories")
      const titles: NodeListOf<HTMLElement> =
        wp.querySelectorAll<HTMLElement>('[data-automation-id="webPartTitleReadMode"]');
      for (let t: number = 0; t < titles.length; t++) {
        const color: string = this._pickColor(titles[t]);
        this._paint(titles[t], color);
        // The title has its own colour on inner <span>s, so paint those too
        const inner: NodeListOf<HTMLElement> = titles[t].querySelectorAll<HTMLElement>('*');
        for (let k: number = 0; k < inner.length; k++) {
          this._paint(inner[k], color);
        }
      }

      // B) Post headlines
      const heads: NodeListOf<HTMLElement> =
        wp.querySelectorAll<HTMLElement>('[data-automation-id="newsItemTitle"]');
      for (let h: number = 0; h < heads.length; h++) {
        this._paint(heads[h], this._pickColor(heads[h]));
      }
    }
  }

  // Decide: white or brand colour?
  private _pickColor(el: HTMLElement): string {
    // 1. Remember SharePoint's ORIGINAL text colour (before we change it)
    if (!el.hasAttribute('data-newsfont-orig')) {
      const orig: IRgba | undefined = this._parse(window.getComputedStyle(el).color);
      const isLight: boolean = !!orig && this._brightness(orig) > 0.6;
      el.setAttribute('data-newsfont-orig', isLight ? 'light' : 'dark');
    }

    // SharePoint made it white (text on picture / dark section) -> keep white
    if (el.getAttribute('data-newsfont-orig') === 'light') {
      return WHITE;
    }

    // 2. Check the real background behind the text
    const bg: IRgba | undefined = this._findBackground(el);
    if (bg && this._brightness(bg) < 0.5) {
      return WHITE; // dark background
    }

    // 3. Light background
    return this._color;
  }

  // Walk up the page until we find an element with a solid background colour
  private _findBackground(el: HTMLElement): IRgba | undefined {
    let node: HTMLElement | null = el.parentElement;
    while (node) {
      const cs: CSSStyleDeclaration = window.getComputedStyle(node);
      if (cs.backgroundImage && cs.backgroundImage !== 'none') {
        return undefined; // background is a picture - can't measure
      }
      const c: IRgba | undefined = this._parse(cs.backgroundColor);
      if (c && c.a > 0.5) {
        return c;
      }
      node = node.parentElement;
    }
    return undefined; // nothing found - treat as light page
  }

  private _paint(el: HTMLElement, color: string): void {
    el.style.setProperty('color', color, 'important');
  }

  // "rgb(18, 0, 82)" or "rgba(0, 0, 0, 0)" -> numbers
  private _parse(value: string): IRgba | undefined {
    const parts: RegExpMatchArray | null = value.match(/\d+(\.\d+)?/g);
    if (!parts || parts.length < 3) {
      return undefined;
    }
    return {
      r: parseFloat(parts[0]),
      g: parseFloat(parts[1]),
      b: parseFloat(parts[2]),
      a: parts.length > 3 ? parseFloat(parts[3]) : 1
    };
  }

  // 0 = black, 1 = white
  private _brightness(c: IRgba): number {
    return (0.299 * c.r + 0.587 * c.g + 0.114 * c.b) / 255;
  }
}