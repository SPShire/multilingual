import * as React from 'react';
import * as lodash from 'lodash';
import "@pnp/polyfill-ie11";
import * as strings from 'MultilingualExtensionApplicationCustomizerStrings';

import { MultilingualContainer } from './MultilingualContainer';
import { ILanguage, IPageProperties, IMap } from '../../../common/models/Models';

import InitService from '../services/InitService';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { Logger, LogLevel } from "@pnp/logging";

import * as rs from 'css-element-queries';
import { KeyCodes } from '@uifabric/utilities/lib';

export interface IMultilingualExtProps {
  context: ApplicationCustomizerContext;
  topPlaceholder: HTMLDivElement;
  disable: () => void;
}

export interface IMultilingualExtState {
  location: number[];
  languages: ILanguage[];
  pages: IPageProperties[];
  url: string;
  usersLanguage: string;
  homepage: boolean;
  editMode: boolean;
  rootFolder: boolean;
  userCanEdit: boolean;
}

export class MultilingualExtState implements IMultilingualExtState {
  constructor(
    public location: number[] = [],
    public languages: ILanguage[] = null,
    public pages: IPageProperties[] = null,
    public url: string = null,
    public usersLanguage: string = null,
    public homepage: boolean = false,
    public editMode: boolean = document.location.href.indexOf("Mode=Edit") !== -1,
    public rootFolder: boolean = false,
    public userCanEdit: boolean = false
  ) { }
}

export class MultilingualExt extends React.Component<IMultilingualExtProps, IMultilingualExtState> {
  private LOG_SOURCE: string = 'MultilingualExt';
  private _init: InitService;
  //private _resizeSensorTop;
  //private _topElement;
  //private _rightElement;
  //private _resizeSensorRight;
  private _reinit: boolean = false;

  constructor(props) {
    super(props);
    this.state = new MultilingualExtState();
    this._init = new InitService(props.context, "{E842C4DA-371B-410E-A3CA-B890D4342564}");
    try {
      // this._resizeSensorTop = new rs.ResizeSensor(this.props.topPlaceholder.parentElement, () => {
      //   this.calculateLocation();
      // });
      this.loadingComponent();
    } catch (err) {
      Logger.write(`${err} - ${this.LOG_SOURCE} (constructor)`, LogLevel.Error);
    }
  }

  public async componentWillUnmount() {
    //this._resizeSensorTop.detach();
    //this._resizeSensorRight.detach();
  }

  private async loadingComponent(): Promise<void> {
    //let pattern = null;
    let homepage: boolean = false;
    let usersLanguage: string = null;
    let inLanguageFolder = false;
    let url: string = "";
    let languages: ILanguage[] = null;
    let pages: IPageProperties[] = null;

    try {
      usersLanguage = sessionStorage.getItem('menuLanguage');
      if (!usersLanguage)
        usersLanguage = this.props.context.pageContext.cultureInfo.currentUICultureName;
      url = document.location.href.toLowerCase();
      url = (url.indexOf('?') > 0) ? url.split('?')[0] : url;
      url = decodeURIComponent(url);
      let rootFolder: boolean = false;
      if (url.indexOf('.aspx') == -1) {
        //Assume home page and get full url
        homepage = true;
        url = `${document.location.origin}${this.props.context.pageContext.legacyPageContext["serverRequestPath"]}`.toLowerCase();
      }

      //If not in SitePages don't continue.
      let spIdx = url.toLowerCase().indexOf('sitepages');
      if (spIdx < 0) {
        Logger.write(`User is not in site pages library - ${this.LOG_SOURCE}`, LogLevel.Info);
        return;
      } else {
        rootFolder = ((url.substr(spIdx).split("/").length - 1) == 1);
      }

      //Get Langugages
      languages = await this._init.getLanguages();
      if (languages && languages.length > 0) {
        //Validate Configuration
        var valid: boolean = await this._init.validateConfig(languages);
        //Check if current page is in language folder
        if (url.lastIndexOf('-') > spIdx) {
          let spUrl = url.substr(spIdx);
          var match = spUrl.match("(.{2}-.{2,4})\/");
          if (match != null)
            inLanguageFolder = (lodash.find(languages, (lang) => { return (lang.code.toLowerCase() == match[1].toLowerCase()); }) != null);
        }

        //In Language folder -- configure and render app customizer
        if (rootFolder || inLanguageFolder) {
          if (valid) {
            let userCanEdit = await this._init.pageEdit();
            pages = await this._init.getAllPages(languages, url);
            this.setState({
              usersLanguage: usersLanguage,
              homepage: homepage,
              url: url,
              languages: languages,
              pages: pages,
              rootFolder: rootFolder,
              userCanEdit: userCanEdit
            });
          } else {
            Logger.write(`${strings.Title}: Could not validate configuration of Site Page content type. - ${this.LOG_SOURCE}`, LogLevel.Warning);
          }
        }
      } else {
        Logger.write(`${strings.Title}: No languages were available. Please check the languages.config files in the root sites, root web, Site Assets folder. - ${this.LOG_SOURCE}`, LogLevel.Warning);
      }
      return;
    } catch (err) {
      Logger.write(`${err} - ${this.LOG_SOURCE} (loadingComponent)`, LogLevel.Error);
    }
    return;
  }

  public async componentDidMount(): Promise<void> {
    //this.calculateLocation();
    
    // Thanks to Elio Stuyf - https://www.eliostruyf.com/check-page-mode-from-within-spfx-extensions/
    // Binding to page mode changes
    const _pushState = () => {
      const _defaultPushState = history.pushState;
      // We need the current this context to update the component its state
      const _self = this;
      return function (data: any, title: string, url?: string | null) {
        try {
          let rootFolder: boolean = false;
          let homePage: boolean = false;
          let languageFolder: boolean = false;
          let workingUrl = url.toLowerCase();
          if (workingUrl.indexOf('.aspx') == -1) {
            //Assume home page and get full url
            workingUrl = `${document.location.origin}${_self.props.context.pageContext.legacyPageContext["serverRequestPath"]}`.toLowerCase();
          }
          let pageUrl = workingUrl.split("?")[0];
          let spIdx: number = pageUrl.toLowerCase().indexOf('sitepages');
          if (spIdx < 0) {
            //If not in sites pages folder -- disable multilingual
            Logger.write(`User is not in site pages library - ${_self.LOG_SOURCE} (componentDidMount-pushState)`, LogLevel.Info);
            _self.props.disable();
          } else {
            rootFolder = ((pageUrl.substr(spIdx).split("/").length - 1) == 1);
            homePage = (url.indexOf('.aspx') < 0);
            if (pageUrl.lastIndexOf('-') > spIdx) {
              //Validate if new page in language folder
              let spUrl = pageUrl.substr(spIdx);
              let match = spUrl.match("(.{2}-.{2,4})\/");
              if (match != null) {
                languageFolder = (lodash.find(_self.state.languages, (lang) => { return (lang.code.toLowerCase() == match[1].toLowerCase()); }) != null);
              }
            }
            //If not the home page, root folder, or language folder -- disable multilingual
            if (!homePage && !rootFolder && !languageFolder) {
              Logger.write(`Not homepage, root folder, or language folder - ${_self.LOG_SOURCE} (componentDidMount-pushState)`, LogLevel.Info);
              _self.props.disable();
            } else {
              //Check if state needs to be updated
              let editMode = url.indexOf('Mode=Edit') !== -1;
              if (editMode !== _self.state.editMode ||
                pageUrl !== _self.state.url ||
                rootFolder !== _self.state.rootFolder) {
                _self.setState({
                  url: pageUrl,
                  editMode: editMode,
                  rootFolder: rootFolder
                });
              }
            }
          }
        } catch (err) {
          Logger.write(`${err} - ${_self.LOG_SOURCE} (componentDidMount-pushState)`, LogLevel.Error);
        }
        // Call the original function with the provided arguments
        // This context is necessary for the context of the history change
        return _defaultPushState.apply(this, [data, title, url]);
      };
    };
    history.pushState = _pushState();
  }

  public shouldComponentUpdate(nextProps: Readonly<IMultilingualExtProps>, nextState: Readonly<IMultilingualExtState>) {
    if ((lodash.isEqual(nextState, this.state) && lodash.isEqual(nextProps, this.props)))
      return false;
    //url changed but pages not yet refreshed
    if (!lodash.isEqual(nextState.url, this.state.url) && lodash.isEqual(nextState.pages, this.state.pages))
      this._reinit = true;
    return true;
  }

  public componentDidUpdate() {
    if (this._reinit && this.state.url.length > 0) {
      this._reinit = false;
      this.reloadPages();
    }
  }

  // private getLocation(): boolean {
  //   if (this._topElement == null || this._rightElement == null) {
  //     this._topElement = !document.getElementsByClassName('commandBarWrapper')[0] ? document.getElementsByClassName('od-TopBar-commandBar')[0] : document.getElementsByClassName('commandBarWrapper')[0];
  //     if (this._topElement != null) {
  //       this._rightElement = !this._topElement.getElementsByClassName('ms-CommandBar-primaryCommands')[0] ? this._topElement.getElementsByClassName('ms-OverflowSet')[0] : this._topElement.getElementsByClassName('ms-CommandBar-primaryCommands')[0];
  //       try {
  //         if (this._rightElement) {
  //           this._resizeSensorRight = new rs.ResizeSensor(this._rightElement, () => {
  //             this.calculateLocation();
  //           });
  //         } else {
  //           Logger.write(`${this.LOG_SOURCE} (getLocation) - rightElement was null`, LogLevel.Warning);
  //           let n: any = setTimeout(() => {
  //             this.calculateLocation();
  //           }, 500);
  //           return false;
  //         }
  //       } catch (err) {
  //         Logger.write(`${err} - ${this.LOG_SOURCE} (getLocation)`, LogLevel.Error);
  //         return false;
  //       }
  //     } else {
  //       Logger.write(`${this.LOG_SOURCE} (getLocation) - topElement was null`, LogLevel.Warning);
  //       return false;
  //     }
  //   }
  //   return true;
  // }

  // private calculateLocation() {
  //   try {
  //     //Calculate position of command bar to position language selector/details
  //     let top: number = 0;
  //     let right: number = 0;
  //     if (this.getLocation()) {
  //       if (!this._topElement != null) {
  //         top = this._topElement.getBoundingClientRect().top + 5;
  //         if (this._rightElement != null)
  //           right = window.innerWidth - this._rightElement.getBoundingClientRect().right;
  //       }
  //       Logger.write(`Location: ${top}-${right} - ${this.LOG_SOURCE}`, LogLevel.Info);
  //       this.setState({
  //         location: [top, right]
  //       });
  //     }
  //   } catch (err) {
  //     Logger.write(`${err} - ${this.LOG_SOURCE} (calculateLocation)`, LogLevel.Error);
  //   }
  // }

  private async reloadPages(): Promise<void> {
    let url = this.state.url;
    let pages = this.state.pages;
    try {
      if (url.indexOf(document.location.pathname.toLowerCase()) === -1 && document.location.pathname.toLowerCase().indexOf('.aspx') > -1)
        url = (document.location.href.indexOf('?') > 0) ? document.location.href.split('?')[0] : document.location.href;
      url = url.toLowerCase();
      pages = await this._init.getAllPages(this.state.languages, url);
      this.setState({
        url: url,
        pages: pages
      });
    } catch (err) {
      Logger.write(`${err} - ${this.LOG_SOURCE} (reloadPages)`, LogLevel.Error);
    }
  }

  private async setEditMode(): Promise<void> {
    if (this.state.editMode) return;
    let checkedOut = await this._init.checkoutPage();
    if (checkedOut) {
      document.location.href = `${document.location.href}${(document.location.href.indexOf("?") === -1 ? "?" : "&")}Mode=Edit`;
    }
    return;
  }

  private savePage(): void {
    this._init.savePage();
  }

  private async manageRedirectorPage(redirectorUrl: string, mapping: IMap): Promise<boolean> {
    return await this._init.manageRedirectorPage(redirectorUrl, mapping);
  }

  public render() {
    if (!this.state.languages || !this.state.pages || this.state.pages.length < 1) return null;
    return (
      <MultilingualContainer
        rootFolder={this.state.rootFolder}
        editMode={this.state.editMode}
        //location={this.state.location}
        pages={this.state.pages}
        languages={this.state.languages}
        url={this.state.url}
        reloadPages={this.reloadPages.bind(this)}
        savePage={this.savePage.bind(this)}
        disable={this.props.disable}
        //calculateLocation={this.calculateLocation.bind(this)}
        setEditMode={this.setEditMode.bind(this)}
        userCanEdit={this.state.userCanEdit}
        manageRedirectorPage={this.manageRedirectorPage.bind(this)} />
    );
  }
}