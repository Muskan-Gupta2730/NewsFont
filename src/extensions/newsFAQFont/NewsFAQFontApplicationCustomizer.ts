import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

export interface INewsFAQFontApplicationCustomizerProperties {
  headingColor: string;
}

interface IRgba {
  r: number;
  g: number;
  b: number;
  a: number;
}

const NEWS_HEADLINE: string = '[data-automation-id="newsItemTitle"]';
const WHITE: string = '#ffffff';

export default class NewsFAQFontApplicationCustomizer
  extends BaseApplicationCustomizer<INewsFAQFontApplicationCustomizerProperties> {

  private _color: string = '#120052';
  private _pending: boolean = false;
  private _observer: MutationObserver | undefined;
  private _editModeInterval: number | undefined;

  public onInit(): Promise<void> {
    this._color = this.properties.headingColor || '#120052';

    this._applyColors();

    // News and FAQ load late; pages change without reload, so keep watching
    this._observer = new MutationObserver(() => this._schedule());
    this._observer.observe(document.body, { childList: true, subtree: true });

    // In edit mode SharePoint renders web parts later than in published view.
    // Run extra passes at 1 s and 3 s to catch elements that weren't in the
    // DOM yet when the MutationObserver first fired.
    window.setTimeout(() => this._applyColors(), 1000);
    window.setTimeout(() => this._applyColors(), 3000);

    // Edit mode: the DnD border overlay (data-dnd-flex-border) appears when
    // SharePoint finishes rendering the canvas. Use it as a signal to repaint.
    this._waitForEditMode();

    return Promise.resolve();
  }

  // Poll until the DnD overlay exists, then do a final color pass.
  // This guarantees we paint AFTER edit mode chrome is fully rendered.
  private _waitForEditMode(): void {
    const dndBorder: Element | null = document.querySelector('[data-dnd-flex-border]');
    if (dndBorder) {
      // Already present — apply immediately
      this._applyColors();
      return;
    }
    // Not present yet — check again in 500 ms (up to 10 s total)
    let attempts: number = 0;
    this._editModeInterval = window.setInterval(() => {
      attempts++;
      if (document.querySelector('[data-dnd-flex-border]')) {
        window.clearInterval(this._editModeInterval);
        this._editModeInterval = undefined;
        this._applyColors();
      } else if (attempts >= 20) {
        // Not edit mode (published view) — stop polling
        window.clearInterval(this._editModeInterval);
        this._editModeInterval = undefined;
      }
    }, 500);
  }
  public onDispose(): void {
    if (this._observer) {
      this._observer.disconnect();
    }
    if (this._editModeInterval !== undefined) {
      window.clearInterval(this._editModeInterval);
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
    }, 300);
  }

  // Colour every news headline and every FAQ accordion heading on the page
  private _applyColors(): void {
    // News headings
    const newsHeads: NodeListOf<HTMLElement> = document.querySelectorAll<HTMLElement>(NEWS_HEADLINE);
    for (let h: number = 0; h < newsHeads.length; h++) {
      this._paint(newsHeads[h], this._pickColor(newsHeads[h]));
    }

    // FAQ accordion question headings (Fluent UI AccordionHeader buttons)
    // In published view: text is inside .fui-AccordionHeader__button
    // In edit mode: text is inside a sibling button (___izlrul0) with a <text> child
    const faqSelectors: string = [
      '.fui-AccordionHeader__button',
      '.fui-AccordionHeader__button text',
      '.___1ot4xyu button',
      '.___1ot4xyu button text'
    ].join(', ');
    const faqButtons: NodeListOf<HTMLElement> = document.querySelectorAll<HTMLElement>(faqSelectors);
    for (let f: number = 0; f < faqButtons.length; f++) {
      this._paint(faqButtons[f], this._pickFaqColor(faqButtons[f]));
    }
  }

  // Decide color for FAQ buttons: check background, no cached-origin logic
  private _pickFaqColor(el: HTMLElement): string {
    const bg: IRgba | undefined = this._findBackground(el);
    if (bg && this._brightness(bg) < 0.5) {
      return WHITE; // dark section background -> white text
    }
    return this._color; // light background -> brand color
  }

  // Decide: white or brand colour — same logic for both news and FAQ
  private _pickColor(el: HTMLElement): string {
    // 1. Remember SharePoint's ORIGINAL text colour (before we change it)
    if (!el.hasAttribute('data-newsfaqfont-orig')) {
      const orig: IRgba | undefined = this._parse(window.getComputedStyle(el).color);
      const isLight: boolean = !!orig && this._brightness(orig) > 0.6;
      el.setAttribute('data-newsfaqfont-orig', isLight ? 'light' : 'dark');
    }

    // SharePoint made it white (text on picture / dark section) -> keep white
    if (el.getAttribute('data-newsfaqfont-orig') === 'light') {
      return WHITE;
    }

    // 2. Check the real background behind the element
    const bg: IRgba | undefined = this._findBackground(el);
    if (bg && this._brightness(bg) < 0.5) {
      return WHITE; // dark background
    }

    // 3. Light background -> brand colour
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