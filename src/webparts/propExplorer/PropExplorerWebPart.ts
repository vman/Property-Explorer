import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneCheckbox
} from '@microsoft/sp-client-preview';

import styles from './PropExplorer.module.scss';
import * as strings from 'mystrings';
import { IPropExplorerWebPartProps } from './IPropExplorerWebPartProps';
import { EnvironmentType } from '@microsoft/sp-client-base';

export default class PropExplorerWebPart extends BaseClientSideWebPart<IPropExplorerWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.propExplorer}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">${this.properties.title}</span>
              <p class="ms-font-l ms-fontColor-white">${this.properties.description}</p>
              <p class="ms-font-l ms-fontColor-white"></p>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneCheckbox('reactive', {
                  text: strings.ReactiveFieldLabel
                }),
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel,
                  onGetErrorMessage: this._validateTitleAsync.bind(this), // validation function
                  deferredValidationTime: 500, // delay after which the validation function will be run
                  placeholder: "enter webpart title..."
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel,
                  multiline: true,
                  resizable: true,
                  onGetErrorMessage: this._validateDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private _validateTitleAsync(value: string): Promise<string> | string {
    const currentEnvType: EnvironmentType = this.context.environment.type;
    if (currentEnvType === EnvironmentType.SharePoint || currentEnvType == EnvironmentType.ClassicSharePoint) {

      return this.context.httpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/title`)
        .then((response: Response) => {
          return response.json().then((responseJSON) => {
            if (responseJSON.value.toLowerCase() === value.toLowerCase()) {
              return "Title cannot be same as the SharePoint site title";
            }
            else {
              return "";
            }
          });
        });
    }
    else{
      return "";
    }
  }

  private _validateDescription(value: string): string {
    if (value.length < 10) {
      return "At least 10 characters required";
    }
    else {
      return "";
    }
  }

  protected get disableReactivePropertyChanges(): boolean {
    return !this.properties.reactive;
  }
}
