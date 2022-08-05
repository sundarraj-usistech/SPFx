import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ListCreationWebPart.module.scss';
import * as strings from 'ListCreationWebPartStrings';

import {

  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions

} from '@microsoft/sp-http';

export interface IListCreationWebPartProps {
  description: string;
}

export default class ListCreationWebPart extends BaseClientSideWebPart<IListCreationWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.listCreation} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        
        <h4>Creating a New List Dynamically</h4><br>

        <p>Please fill the following details</p><br>

        <table>

        <tr><td>List Name:</td>
        <td><input type = "text" id = "listname"></td></tr>

        <tr><td>List Description:</td>
        <td><input type = "text" id = "listdescription"><td></tr>

        </table>
        <input type = "button" id = "btncreateNewList" value="CREATE">

      </div>
    </section>`;

    this.bindEvent();

  }

  private bindEvent(): void {

    this.domElement.querySelector('#btncreateNewList').addEventListener( 'click', () => {

      this.createNewList();
      // alert("Clicked");

    });

  }

  private createNewList(): void {

    var ListName = document.getElementById("listname")["value"];
    var ListDescription = document.getElementById("listdescription")["value"];
    const listurl: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('"+ListName+"')";

    this.context.spHttpClient.get(listurl, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {

      if(response.status === 200){

        alert("List Name already exists");
        return;

      }
      if(response.status === 404){

        const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";
        const ListDefinition: any = {

          "Title": ListName,
          "Description": ListDescription,
          "AllowContentTypes": true,
          "BaseTemplate": 105,
          "BaseContentTypeEnabled": true,
          
        };

        const spHttpClientOptions: ISPHttpClientOptions = {

          "body": JSON.stringify(ListDefinition)

        };

        this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {

          if(response.status === 201){

            alert("List created Successfully");

          }
          else{

            alert("Error Message: "+response.status+ " - " +response.statusText);

          }

        });  

      }

      else{

        alert("Error Message: "+response.status+ " - " +response.statusText);

      }

    });

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
