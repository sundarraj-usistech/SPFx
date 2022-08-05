import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ReadSitePropertiesWebPart.module.scss';
import * as strings from 'ReadSitePropertiesWebPartStrings';

import{

  Environment,
  EnvironmentType

} from '@microsoft/sp-core-library';

import{

  SPHttpClient,
  SPHttpClientResponse

} from '@microsoft/sp-http';

export interface IReadSitePropertiesWebPartProps {
  description: string;
  environmentTitle: string;
}

export interface ISharePointList{

  Title: string;
  Id: string;

}

export interface IListValues{

  value:ISharePointList[];

}

export default class ReadSitePropertiesWebPart extends BaseClientSideWebPart<IReadSitePropertiesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  private findEnvironment (): void{

    if (Environment.type === EnvironmentType.Local){

        this.properties.environmentTitle = "Local SharePoint Environment";

    }
    else if ((Environment.type === EnvironmentType.SharePoint) || (Environment.type === EnvironmentType.ClassicSharePoint)){

        this.properties.environmentTitle = "Online SharePoint Environment";

    } 
  }

  private getList(): Promise<IListValues>{

    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {

      return response.json(); 

    })

  }

  private renderList(): void{

    if(Environment.type === EnvironmentType.Local){

    }
    else if((Environment.type === EnvironmentType.SharePoint) || (Environment.type === EnvironmentType.ClassicSharePoint)){

      this.getList().then((response)=>{

        this.displayList(response.value);

      });

    }

  }

  private displayList(items : ISharePointList[]): void{

    let html : string = ``;

    items.forEach((item:ISharePointList) => {

      html+= `

      <ul class="${styles.list}">
        <li class ="${styles.listItem}">
          <span class="ms-font-1">${item.Title}</span>
        </li>
        <li class="${styles.listItem}">
          <span class="ms-font-1">${item.Id}</span>
        </li>
      </ul>

      `;

    });

    const listPlaceholder : Element = this.domElement.querySelector('#ListPlaceHolder');
    listPlaceholder.innerHTML = html;

  }

  public render(): void {

    this.findEnvironment();
    this.renderList();

    this.domElement.innerHTML = `
    <section class="${styles.readSiteProperties} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
      </div>
      <div>
        
        <p class = "${styles.welcome}">Absolute URL : <b>${escape(this.context.pageContext.web.absoluteUrl)}</b> </p>
        <p class = "${styles.welcome}">Title : <b>${escape(this.context.pageContext.web.title)}</b> </p>
        <p class = "${styles.welcome}">Relative URL : <b>${escape(this.context.pageContext.web.serverRelativeUrl)}</b> </p>
        <p class = "${styles.welcome}">User Name : <b>${escape(this.context.pageContext.user.displayName)}</b> </p>
        <p class = "${styles.welcome}">Environment Type : <b>${escape(this.properties.environmentTitle)}</b> </p>
        <p class = "${styles.welcome}">Culture Name : <b>${escape(this.context.pageContext.cultureInfo.currentCultureName)}</b> </p>
        <p class = "${styles.welcome}">UI Culture Name : <b>${escape(this.context.pageContext.cultureInfo.currentUICultureName)}</b> </p>
        <p class = "${styles.welcome}">is Right to Left ? : <b>${(this.context.pageContext.cultureInfo.isRightToLeft)}</b> </p>

      </div>

    	<div id="ListPlaceHolder">

      </div>

    </section>`;

  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
