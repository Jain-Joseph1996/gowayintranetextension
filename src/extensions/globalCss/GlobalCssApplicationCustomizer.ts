import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';

import * as strings from 'GlobalCssApplicationCustomizerStrings';

const LOG_SOURCE: string = 'GlobalCssApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGlobalCssApplicationCustomizerProperties {
  // This is an example; replace with your own property
  cssurl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GlobalCssApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalCssApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    const cssUrl = this.context.pageContext.web.absoluteUrl + "/SiteAssets/custom.css";
    if (cssUrl) {
      // inject the style sheet
      const head: any = document.getElementsByTagName('head')[0] || document.documentElement;
      let customStyle: HTMLLinkElement = document.createElement('link');
      customStyle.href = cssUrl;
      customStyle.rel = 'stylesheet';
      customStyle.type = 'text/css';
      head.insertBefore(customStyle, head.firstChild);
    }
    return Promise.resolve();
  }
}
