import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import * as strings from 'GlobalCssApplicationCustomizerStrings';

const LOG_SOURCE: string = 'GlobalCssApplicationCustomizer';
const EXACT_TARGET_URL: string = 'https://goway0.sharepoint.com/sites/Intranet';
const OVERLAY_KEY: string = `overlayShown_${btoa(EXACT_TARGET_URL)}`;

export default class GlobalCssApplicationCustomizer
  extends BaseApplicationCustomizer<{ cssurl: string }> {

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    const userDisplayName = this.context.pageContext.user.displayName;
    const DEFAULT_MESSAGE: string = userDisplayName ? `Welcome, ${userDisplayName}!` : 'Welcome!';

    // 1. ALWAYS inject CSS immediately (first thing)
    const cssUrl = this.properties.cssurl ||
      `${this.context.pageContext.web.absoluteUrl}/SiteAssets/custom.css`;
    if (cssUrl) {
      this._injectCss(cssUrl);
    }

    // 2. Only proceed with overlay check if we're on exact target URL
    if (window.location.href === EXACT_TARGET_URL) {
      await this._handleOverlayLogic(DEFAULT_MESSAGE);
    }

    return Promise.resolve();
  }

  private async _handleOverlayLogic(DEFAULT_MESSAGE: string): Promise<void> {
    // Check if overlay was already shown or this is a reload
    if (sessionStorage.getItem(OVERLAY_KEY) || performance.navigation.type !== 0) {
      return;
    }

    // Create and show overlay after CSS loads
    await this._showOverlayAfterCssLoad(DEFAULT_MESSAGE);
    this._markOverlayShown();
  }

  private _injectCss(url: string): void {
    const link = document.createElement('link');
    link.href = url;
    link.rel = 'stylesheet';
    document.head.appendChild(link);
  }

  private async _showOverlayAfterCssLoad(message: string): Promise<void> {
    return new Promise((resolve) => {
      // Create overlay (hidden initially)
      const overlay = this._createOverlay(message);
      overlay.style.display = 'none';
      document.body.appendChild(overlay);

      // Check when CSS is loaded
      const checkCssLoaded = () => {
        // Simple check if any styles from your CSS are applied
        const testElement = document.createElement('div');
        testElement.style.position = 'absolute';
        document.body.appendChild(testElement);
        const isLoaded = window.getComputedStyle(testElement).position === 'absolute';
        document.body.removeChild(testElement);

        if (isLoaded) {
          // Show overlay only after confirming CSS is loaded
          overlay.style.display = 'flex';
          this._setupOverlayDismissal(overlay, resolve);
        } else {
          setTimeout(checkCssLoaded, 100);
        }
      };

      // Start checking
      checkCssLoaded();
    });
  }

  private _createOverlay(message: string): HTMLElement {
    const overlay = document.createElement('div');
    overlay.innerHTML = `
     <div style="text-align: center; font-family: 'Segoe UI', SegoeUI, 'Segoe WP', Tahoma, Arial, sans-serif;">
          <div style="font-size: 3em; margin-bottom: 20px; font-weight: 300;">${message}</div>
          <div style="font-size: 1.5em; font-weight: 300;">Loading your personalized experience...</div>
        </div>`;

    overlay.style.position = 'fixed';
    overlay.style.top = '0';
    overlay.style.left = '0';
    overlay.style.width = '100%';
    overlay.style.height = '100%';
    overlay.style.backgroundColor = 'rgba(0, 0, 0, 0.9)';
    overlay.style.zIndex = '9999';
    overlay.style.display = 'flex';
    overlay.style.justifyContent = 'center';
    overlay.style.alignItems = 'center';
    overlay.style.color = 'white';
    overlay.style.fontFamily = "'Segoe UI', SegoeUI, 'Segoe WP', Tahoma, Arial, sans-serif";
    overlay.style.transition = 'opacity 0.5s ease-out';
    document.body.appendChild(overlay);


    return overlay;
  }

  private _setupOverlayDismissal(overlay: HTMLElement, callback: () => void): void {
    setTimeout(() => {
      overlay.style.opacity = '0';
      setTimeout(() => {
        overlay.remove();
        callback();
      }, 500);
    }, 2000);
  }

  private _markOverlayShown(): void {
    sessionStorage.setItem(OVERLAY_KEY, 'true');
  }
}