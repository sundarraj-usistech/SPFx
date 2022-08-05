import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudWebPart.module.scss';
import * as strings from 'CrudWebPartStrings';

import { IProductCatalog } from './IProductCatalog';

import {

  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions

} from '@microsoft/sp-http';

export interface ICrudWebPartProps {
  description: string;
}

export default class CrudWebPart extends BaseClientSideWebPart<ICrudWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.crud} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}">
        
      <table border = '5' bgcolor = 'aqua' style = "display: flex;">

       <tr>
       
       <td>Please Enter the Product ID to Search</td>
       <td><input type = "text" id = "txtID"></td>
       <td><input type = "submit" value = "SEARCH" id = "btnRead"></td>

       </tr>

       <tr>
       
       <td>Title </td>
       <td><input type = "text" id = "txtTitle"></td>

       </tr>

       <tr>
       
       <td>Product ID</td>
       <td><input type = "text" id = "txtProductID"></td>

       </tr>

       <tr>
       
       <td>Description</td>
       <td><input type = "text" id = "txtDescription"></td>

       </tr>

       <tr>
       
       <td>Manufacturing Date</td>
       <td><input type = "text" id = "txtManufacturingDate"></td>

       </tr>

       <tr>
       
       <td>Expiry Date</td>
       <td><input type = "text" id = "txtExpiryDate"></td>

       </tr>

       <tr>
       
       <td colspan = '2' align = 'center'>

       <input type = "submit" value = "INSERT" id = "btnInsert">
       <input type = "submit" value = "UPDATE" id = "btnUpdate">
       <input type = "submit" value = "DELETE" id = "btnDelete">
       <input type = "submit" value = "DISPLAY ALL" id = "btnReadAll">

       </td>

       </tr>

      </table>

      </div>

      <div id = "divStatus">
      </div>

    </section>`;

    this.bindEvents();

  }

  private bindEvents(): void {

    this.domElement.querySelector('#btnInsert').addEventListener ('click', () => {

      this.addListItem();

    }); 

    this.domElement.querySelector('#btnRead').addEventListener ('click', () =>{

      this.readListItem();

    });

    this.domElement.querySelector('#btnUpdate').addEventListener ('click', () => {

      this.updateListItem();

    });

    this.domElement.querySelector('#btnDelete').addEventListener ('click', () => {

      this.deleteListItem();

    });

    this.domElement.querySelector('#btnReadAll').addEventListener ('click', () => {

      this.readAllItems();

    });

  }

  private addListItem(): void{

    var title: string = document.getElementById("txtTitle")["value"];
    var productid: string = document.getElementById("txtProductID")["value"];
    var description: string = document.getElementById("txtDescription")["value"];
    var mfgdate: number = document.getElementById("txtManufacturingDate")["value"];
    var expdate: number = document.getElementById("txtExpiryDate")["value"];

    const url:string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductCatalog')/items";

    const itemBody: any = {

      "Title": title,
      "ProductID": productid,
      "Description": description,
      "ManufacturingDate": mfgdate,
      "ExpiryDate": expdate,

    };

    const spHttpClientOptions: ISPHttpClientOptions = {

      "body": JSON.stringify(itemBody)

    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {

      if(response.status ===201){

        let statusmessage: Element = this.domElement.querySelector('#divStatus');
        statusmessage.innerHTML = "List Item has been created Successfully";
        this.clear();

      }
      else{

        let statusmessage: Element = this.domElement.querySelector('#divStatus');
        statusmessage.innerHTML = "An Error has Occurred " + response.status + " - " + response.statusText;

      }

    });

  }

  private readListItem(): void {

    let id: string = document.getElementById("txtID")["value"];
    this.getListItemByID(id).then(listItem => {

      document.getElementById("txtTitle")["value"] = listItem.Title ;
      document.getElementById("txtProductID")["value"] = listItem.ProductID ;
      document.getElementById("txtDescription")["value"] = listItem.Description ;
      document.getElementById("txtManufacturingDate")["value"] = listItem.ManufacturingDate ;
      document.getElementById("txtExpiryDate")["value"] = listItem.ExpiryDate ;

    })  
    .catch(error => {

      let message: Element = this.domElement.querySelector('#divStatus');
      message.innerHTML = "Could not Fetch Details" + error.message;

    });

  }
 
  private getListItemByID(id: string): Promise<IProductCatalog> {

    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductCatalog')/items?$filter=Id eq "+id;
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then((response: SPHttpClientResponse) => {

      return response.json(); 

    })
    .then((listItems: any) => {

      const untypedItem: any = listItems.value[0];
      const listItem: IProductCatalog = untypedItem as IProductCatalog;
      return listItem;

    }) as Promise <IProductCatalog>;

  } 

  private updateListItem(): void {

    var title: string = document.getElementById("txtTitle")["value"];
    var productid: string = document.getElementById("txtProductID")["value"];
    var description: string = document.getElementById("txtDescription")["value"];
    var mfgdate: number = document.getElementById("txtManufacturingDate")["value"];
    var expdate: number = document.getElementById("txtExpiryDate")["value"];

    let id: string = document.getElementById("txtID")["value"];

    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductCatalog')/items ("+ id +")";

    const itemBody: any = {

      "Title": title,
      "ProductID": productid,
      "Description": description,
      "ManufacturingDate": mfgdate,
      "ExpiryDate": expdate,

    };

    const headers: any  = {

      "X-HTTP-Method": "MERGE",
      "IF-MATCH": "*",

    };

    const spHttpClientOptions: ISPHttpClientOptions = {

      "headers": headers,
      "body":JSON.stringify(itemBody)

    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {

      if(response.status === 204){

        let message: Element = this.domElement.querySelector('#divStatus');
        message.innerHTML = "List Item has been Updated Successfully";

      }
      else{

        let message: Element = this.domElement.querySelector('#divStatus');
        message.innerHTML = "Could not Update Details " + response.status + " - " +response.statusText; 

      }

    });

  }

  private deleteListItem(): void{

    let id: string = document.getElementById("txtID")["value"];
    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductCatalog')/items ("+ id +")";

    const headers: any = {

      "X-HTTP-Method": "DELETE",
      "IF-MATCH": "*"

    };

    const spHttpClientOptions: ISPHttpClientOptions = {

      "headers": headers

    }

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
    .then((response: SPHttpClientResponse) => {

      if(response.status === 204){

        let message: Element = this.domElement.querySelector('#divStatus');
        message.innerHTML = "List Item Deleted Successfully";

      }
      else{

        let message: Element = this.domElement.querySelector('#divStatus');
        message.innerHTML = "Could not delete List Item" + response.status + " - " + response.statusText;

      }
      
    });

  }

  private readAllItems(): void {

    this.getListItems().then(listItems => {

      let html: string = '<table border = 1 width = 100% style = "border-collapse: collapse;">';
      html += '<th>Title</th><th>Product ID</th><th>Description</th><th>Manufacturing Date</th><th>Expiry Date</th>';

      listItems.forEach(listItem => {

        html +=`<tr>

        <td>${listItem.Title}</td>
        <td>${listItem.ProductID}</td>
        <td>${listItem.Description}</td>
        <td>${listItem.ManufacturingDate}</td>
        <td>${listItem.ExpiryDate}</td>

        </tr>`;

      });

      html += '</table>';

      const listContainer: Element = this.domElement.querySelector('#divStatus');
      listContainer.innerHTML = html;

    });

  }

  private getListItems(): Promise<IProductCatalog[]>{

    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductCatalog')/items";
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
    .then(response => {

      return response.json();

    })
    .then(json => {

      return json.value;

    }) as Promise<IProductCatalog[]>;

  }

  private clear(): void{

    document.getElementById("txtTitle")["value"] = '' ;
    document.getElementById("txtProductID")["value"] = '' ;
    document.getElementById("txtDescription")["value"] = '' ;
    document.getElementById("txtManufacturingDate")["value"] = '' ;
    document.getElementById("txtExpiryDate")["value"] = '' ;

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
