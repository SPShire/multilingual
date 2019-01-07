import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { override } from '@microsoft/decorators';
import { BaseApplicationCustomizer, PlaceholderName } from '@microsoft/sp-application-base';

import * as strings from 'MultilingualExtensionApplicationCustomizerStrings';

import { MultilingualExt } from './components/MultilingualExt';
import "@pnp/polyfill-ie11";
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";
export interface IMultilingualExtensionApplicationCustomizerProperties { }

/** A Custom Action which can be run during execution of a Client Side Application */
export default class MultilingualExtensionApplicationCustomizer
  extends BaseApplicationCustomizer<IMultilingualExtensionApplicationCustomizerProperties> {

  private LOG_SOURCE: string = "MultilingualExtensionApplicationCustomizer";
  private elementId: string = "MultilingualApplicationCustomizer";

  @override
  public onInit(): Promise<void> {
    return new Promise(() => {
      Logger.subscribe(new ConsoleListener());
      Logger.activeLogLevel = LogLevel.Info;
      if (document.getElementById(this.elementId)) {
        Logger.write(`ERROR - ${strings.Title} already initialized! - ${this.LOG_SOURCE} (onInit)`, LogLevel.Error);
        document.getElementById(this.elementId).remove();
      }
      Logger.write(`Initialized ${strings.Title} - ${this.LOG_SOURCE}`, LogLevel.Info);
      this.context.placeholderProvider.changedEvent.add(this, this.renderMultilingual);
      return;
    });
  }

  private disableMultilingual() {
    let multiContainer = document.getElementById(this.elementId);
    multiContainer.innerHTML = "";
  }

  protected onDispose(): void {
    ReactDOM.unmountComponentAtNode(document.getElementById(this.elementId));
    document.getElementById(this.elementId).remove();
  }

  private renderMultilingual() {
    try {
      if (document.getElementById(this.elementId)) {
        Logger.write(`ERROR - ${strings.Title} already initialized! - ${this.LOG_SOURCE} (renderMultilingual)`, LogLevel.Error);
        return;
      }
      //let topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {});
      let bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Bottom, { onDispose: this.onDispose });
      let multiContainer = document.createElement("DIV");
      multiContainer.setAttribute("id", this.elementId);
      //topPlaceholder.domElement.appendChild(multiContainer);
      bottomPlaceholder.domElement.appendChild(multiContainer);
      //Placeholder for elements to be added to dom
      let element = React.createElement(MultilingualExt, { context: this.context, disable: this.disableMultilingual.bind(this), topPlaceholder: bottomPlaceholder.domElement });
      //let element = React.createElement();
      let elements: any[] = [];
      elements.push(element);
      ReactDOM.render(elements, multiContainer);
    } catch (err) {
      Logger.write(`${err} - ${this.LOG_SOURCE}`, LogLevel.Error);
    }
  }
}
