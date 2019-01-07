import * as React from 'react';
import "@pnp/polyfill-ie11";

import styles from "./MultilingualExtension.module.scss";
import * as lodash from 'lodash';
import { IPageProperties, ILanguage } from '../../../common/models/Models';
import { IComboBoxOption } from 'office-ui-fabric-react/lib/ComboBox';
import { DefaultButton, IconButton } from 'office-ui-fabric-react/lib/Button';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import {Logger,LogLevel} from "@pnp/logging";
import { Label } from 'office-ui-fabric-react/lib/Label';
import { sp } from '@pnp/sp';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';

export interface IPageVariants {
  variant: string;
  url: string;
}

export interface IMultilingualButtonProps {
  //location: number[];
  languages: ILanguage[];
  setShowPanel: () => void;
  editMode: boolean;
  rootFolder: boolean;
  currentPage: IPageProperties;
  variantPages: IPageProperties[];
  setEditMode: () => Promise<void>;
  savePage: () => void;
  userCanEdit: boolean;
}

export interface IMultilingualButtonState {
  buttonLabel: string;
  pageLanguages: IPageVariants[];
  pageLanguagesOptions: IComboBoxOption[];
  defaultOption: string;
  currentPageInit: boolean;
  movePageLocation: string;
  movingPage: boolean;
}

export class MultilingualButtonState implements IMultilingualButtonState {
  constructor(
    public buttonLabel: string = "Language Details",
    public pageLanguages: IPageVariants[] = [],
    public pageLanguagesOptions: IComboBoxOption[] = [],
    public defaultOption: string = "",
    public currentPageInit: boolean = false,
    public movePageLocation: string = "",
    public movingPage: boolean = false
  ) { }
}

export class MultilingualButton extends React.Component<IMultilingualButtonProps, IMultilingualButtonState> {
  private LOG_SOURCE: string = 'MultilingualButton';
  private variantsChanged: boolean = false;

  constructor(props) {
    super(props);
    this.state = new MultilingualButtonState("Language Details", [], [], "", false, this.props.languages[0].code);
  }

  public componentDidUpdate() {
    if (this.variantsChanged || (this.props.currentPage.Id != "" && this.state.defaultOption == "")) {
      this.variantsChanged = false;
      this.init();
    }
  }

  public shouldComponentUpdate(nextProps: Readonly<IMultilingualButtonProps>, nextState: Readonly<IMultilingualButtonState>) {
    if (!nextProps.currentPage || nextProps.currentPage.Id == "" || (lodash.isEqual(nextState, this.state) && lodash.isEqual(nextProps, this.props)))
      return false;
    if (!lodash.isEqual(nextProps.variantPages, this.props.variantPages) || !lodash.isEqual(nextProps.currentPage, this.props.currentPage))
      this.variantsChanged = true;
    return true;
  }

  private init(): void {
    let pageLanguages: IPageVariants[] = [];
    let pageLanguagesOptions: IComboBoxOption[] = [];
    let defaultOption: string = "";
    try {
      let currentPageInit: boolean = !(this.props.currentPage.MasterTranslationPage == "");
      if (this.props.currentPage.LanguageVariant.length == 0) {
        pageLanguages.push({ variant: this.props.currentPage.LanguageFolder, url: this.props.currentPage.Url });
        defaultOption = this.props.currentPage.LanguageFolder;
      } else {
        if (this.props.currentPage.MasterTranslationPage == this.props.currentPage.Id) {
          let langDesc = lodash.find(this.props.languages, {code: this.props.currentPage.LanguageFolder}).description;
          pageLanguages.push({ variant: `Default - ${langDesc}`, url: this.props.currentPage.Url });
          defaultOption = `Default - ${langDesc}`;
        } else {
          let variants = this.props.currentPage.LanguageVariant.join(", ");
          if (variants.lastIndexOf(",") > -1)
            variants = variants.slice(0, variants.lastIndexOf(","));
          pageLanguages.push({ variant: variants, url: this.props.currentPage.Url });
          defaultOption = variants;
        }
      }
      if (this.props.variantPages.length > 0) {
        this.props.variantPages.forEach((variant: IPageProperties) => {
          if (variant.MasterTranslationPage == variant.Id) {
            let langDesc = lodash.find(this.props.languages, {code: variant.LanguageFolder}).description;
            pageLanguages.push({ variant: `Default - ${langDesc}`, url: variant.Url });
          } else {
            if (variant.LanguageVariant.length > 0) {
              let variants = variant.LanguageVariant.join(", ");
              if (variants.lastIndexOf(",") > -1)
                variants = variants.slice(0, variants.lastIndexOf(","));
              pageLanguages.push({ variant: variants, url: variant.Url });
            }
          }
        });
      }
      pageLanguages.forEach((item: IPageVariants) => {
        let text = item.variant;
        let language = lodash.find(this.props.languages, { code: item.variant });
        if (language)
          text = language.description;
        pageLanguagesOptions.push({ key: item.variant, text: text });
      });
      this.setState({
        currentPageInit: currentPageInit,
        pageLanguages: pageLanguages,
        pageLanguagesOptions: pageLanguagesOptions,
        defaultOption: defaultOption
      });
    } catch (err) {
      Logger.write(`${err} - ${this.LOG_SOURCE}`, LogLevel.Error);
    }
    return;
  }

  @autobind
  private changePage(event: React.ChangeEvent<HTMLSelectElement>) {
    try {
      let value = event.target.value;
      let variantLocation = lodash.find(this.state.pageLanguages, ["variant", value]);
      if (variantLocation != null) {
        document.location.href = variantLocation.url;
      }
    } catch (err) {
      Logger.write(`${err} - ${this.LOG_SOURCE}`, LogLevel.Error);
    }
  }

  @autobind
  private movePageLocation(event: React.ChangeEvent<HTMLSelectElement>) {
    let value = event.target.value;
    this.setState({
      movePageLocation: value
    });
  }

  private async movePage() {
    try{
      let url = this.props.currentPage.Url;
      let currentUrl = document.location.pathname.toLowerCase();
      if(currentUrl != url)
        url = currentUrl;
      let redirectUrl = url.replace("sitepages", `sitepages/${this.state.movePageLocation}`);
      let destUrl = url.replace("sitepages", `sitepages/${this.state.movePageLocation}`).replace(/'/g, "''");
      let sourceUrl = url.replace(/'/g, "''");
      await sp.web.getFileByServerRelativeUrl(sourceUrl).moveTo(destUrl);
      document.location.href = `${redirectUrl}${document.location.search}`;
    }catch(err){
      Logger.write(`${err} - ${this.LOG_SOURCE}`, LogLevel.Error);
      this.setState({
        movingPage: false
      });
    }
  }

  @autobind
  private async doMovePage(){
    if(this.state.movePageLocation == "") return;
    this.setState({movingPage: true}, () => {
      this.movePage();
    });
  }

  public render() {
    // let buttonStyle = {};
    // if (this.props.location[0] > 0 && this.props.location[1] > 0)
    //   buttonStyle = {
    //     top: this.props.location[0],
    //     left: 'inherit',
    //     right: this.props.location[1],
    //     bottom: 'inherit'
    //   };
      //style={buttonStyle}
    return (
      <div className={styles.multilingualButton} >
        {this.state.currentPageInit && !this.props.rootFolder && !this.props.editMode &&
          <select className={styles.languageCombo} onChange={this.changePage} value={this.state.defaultOption}>
            {this.state.pageLanguagesOptions && this.state.pageLanguagesOptions.map((o) => {
              return (
                <option key={o.key} value={o.key}>{o.text}</option>
              );
            })}
          </select>
        }
        {this.props.userCanEdit && !this.state.currentPageInit && !this.props.rootFolder && !this.props.editMode &&
          <IconButton className={styles.languageButton + " ms-fontColor-redDark"}
          iconProps={{ iconName: 'Error' }}
          title="Multilingual Not Configured"
          onClick={this.props.setEditMode}
        />
        }
        {!this.props.rootFolder && this.props.editMode &&
          <DefaultButton className={styles.languageButton}
            primary={true}
            text={this.state.buttonLabel}
            onClick={this.props.setShowPanel}
          />
        }
        {this.props.userCanEdit && this.props.rootFolder &&
          <div className={styles.buttonMovePage}>
            <Label className={styles.label}>Folder: </Label>
            <select className={styles.languageCombo + " " + styles.label} onChange={this.movePageLocation} value={this.state.movePageLocation}>
              {this.props.languages && this.props.languages.map((o) => {
                return (
                  <option key={o.code} value={o.code}>{o.description}</option>
                );
              })}
            </select>
            {!this.state.movingPage &&
              <DefaultButton className={styles.languageButton + " ms-bgColor-redDark ms-fontColor-white"}
                text="Move"
                onClick={this.doMovePage}
              />
            }
            {this.state.movingPage &&
              <Spinner className={styles.spinner} size={SpinnerSize.medium} />
            }
          </div>
        }
      </div>
    );
  }
}




