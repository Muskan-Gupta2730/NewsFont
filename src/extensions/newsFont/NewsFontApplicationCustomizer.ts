import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';

export interface INewsFontApplicationCustomizerProperties {
  headingColor: string;
}

// The out-of-box News web part
const NEWS_WEBPART: string = '[data-sp-web-part-id="8c88f208-6c77-4bdb-86a0-0c47b4316588"]';

export default class NewsFontApplicationCustomizer
  extends BaseApplicationCustomizer<INewsFontApplicationCustomizerProperties> {

  private _color: string = '#120052';
  private _pending: boolean = false;
  private _observer: MutationObserver | undefined;

  public onInit(): Promise<void> {
    this._color = this.properties.headingColor || '#120052';

    // A) Web part title ("Latest Sirva Life stories") - always brand colour
    this._addTitleStyle();

    // B) Post headlines - colour the dark ones, leave the white ones
    this._colorHeadlines();

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

  private _addTitleStyle(): void {
    const styleId: string = 'newsfont-heading-style';
    if (document.getElementById(styleId)) {
      return;
    }
    const style: HTMLStyleElement = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      ${NEWS_WEBPART} [data-automation-id="webPartTitleReadMode"],
      ${NEWS_WEBPART} [data-automation-id="webPartTitleReadMode"] * {
        color: ${this._color} !important;
      }
    `;
    document.head.appendChild(style);
  }

  // Wait 200ms so we don't run hundreds of times while the page loads
  private _schedule(): void {
    if (this._pending) {
      return;
    }
    this._pending = true;
    window.setTimeout(() => {
      this._pending = false;
      this._colorHeadlines();
    }, 200);
  }

  private _colorHeadlines(): void {
    const titles: NodeListOf<HTMLElement> = document.querySelectorAll<HTMLElement>(
      `${NEWS_WEBPART} [data-automation-id="newsItemTitle"]`
    );

    for (let i: number = 0; i < titles.length; i++) {
      const el: HTMLElement = titles[i];

      // Already coloured by us - skip
      if (el.getAttribute('data-newsfont') === 'done') {
        continue;
      }

      // White/light headline (Carousel, Tiles, Filmstrip) - leave it alone
      if (this._isLight(window.getComputedStyle(el).color)) {
        continue;
      }

      // Dark headline - make it brand colour
      el.style.setProperty('color', this._color, 'important');
      el.setAttribute('data-newsfont', 'done');
    }
  }

  // Turns "rgb(255, 255, 255)" into a brightness check
  private _isLight(rgb: string): boolean {
    const parts: RegExpMatchArray | null = rgb.match(/\d+(\.\d+)?/g);
    if (!parts || parts.length < 3) {
      return false;
    }
    const r: number = parseFloat(parts[0]);
    const g: number = parseFloat(parts[1]);
    const b: number = parseFloat(parts[2]);
    const brightness: number = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    return brightness > 0.6;
  }
}