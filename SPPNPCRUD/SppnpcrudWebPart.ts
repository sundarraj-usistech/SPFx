import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SppnpcrudWebPart.module.scss';
import * as strings from 'SppnpcrudWebPartStrings';

import * as pnp from 'sp-pnp-js';


export interface ISppnpcrudWebPartProps {
  description: string;
}

export default class SppnpcrudWebPart extends BaseClientSideWebPart<ISppnpcrudWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  // protected onInit(): Promise<void> {
  //   this._environmentMessage = this._getEnvironmentMessage();

  //   return super.onInit();
  // }

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      pnp.setup({
        spfxContext: this.context
      });
      
    });
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.sppnpcrud} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
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

      <div id = "divStatus"/>

      <div id = "listData">

        <h5>All List Items</h5>

      </div>

      </div>
    </section>`;

    this.bindEvents();

  }

  private bindEvents(): void {

    this.domElement.querySelector('#btnInsert').addEventListener('click', () => {

      this.addListItem();

    });

    this.domElement.querySelector('#btnRead').addEventListener('click', () => {

      this.readListItem();

    });

    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => {

      this.updateListItem();

    });

    this.domElement.querySelector('#btnReadAll').addEventListener('click', () => {

      this.readAllListItems();

    });

    this.domElement.querySelector('#btnDelete').addEventListener('click', () => {

      this.deleteListItem();

    })

  }

  private addListItem(): void {

    var title: string = document.getElementById('#txtTitle')["value"];
    var productid: string = document.getElementById('#txtProductID')["value"];
    var description: string = document.getElementById('#txtDescription')["vaue"];
    var manufacturingdate: number = document.getElementById('#txtManufacturingDate')["value"];
    var expirydate: number = document.getElementById('#txtExpiryDate')["value"];

    const url: string = this.context.pageContext.site.absoluteUrl + "/_api/web/lists/getbytitle('ProductCatlog')/items";

    pnp.sp.web.lists.getByTitle("ProductCatalog").items.add ({

      Title: title,
      ProductID: productid,
      Description: description,
      ManufacturingDate: manufacturingdate,
      ExpiryDate: expirydate,

    })
    .then(response => {

      alert("List Item Inserted Successfully");

    });

  }

  private readListItem(): void {

    const id: number = document.getElementById('#txtID')["value"];

    pnp.sp.web.lists.getByTitle("ProductCatalog").items.getById(id).get().then((item: any) => {

      item["Title"] = document.getElementById('#txtTitle')["value"];
      item["ProductID"] = document.getElementById('#txtProductID')["value"];
      item["Description"] = document.getElementById('#txtDescription')["value"];
      item["ManufacturingDate"] = document.getElementById('#txtManufacturingDate')["value"];
      item["ExpiryDate"] = document.getElementById('#txtExpiryDate')["value"];

    });

  }

  private updateListItem(): void{

    var title: string = document.getElementById('#txtTitle')["value"];
    var productid: string = document.getElementById('#txtProductID')["value"];
    var description: string = document.getElementById('#txtDescription')["vaue"];
    var manufacturingdate: number = document.getElementById('#txtManufacturingDate')["value"];
    var expirydate: number = document.getElementById('#txtExpiryDate')["value"];

    let id: number = document.getElementById("#txtID")["value"];

    pnp.sp.web.lists.getByTitle("ProductCatalog").items.getById(id).update ({

      Title: title,
      ProductID: productid,
      Description: description,
      ManufacturingDate: manufacturingdate,
      ExpiryDate: expirydate,

    })
    .then(response => {

      alert("List Item Updated Successfully");

    });


  }

  private readAllListItems(): void {

    let html: string = '<table border = 1 width = 100% style = "bordercollapse: collapse;">';
    html += '<th>Title</th><th>Product ID</th><th>Description</th><th>Manufacturing Date</th><th>Expiry Date</th>';

    pnp.sp.web.lists.getByTitle("Product Catalog").items.get().then ((items: any[]) => {

      items.forEach(function (item) {
        
        html += `
        
        <tr>

          <td>${item["Title"]}</td>
          <td>${item["ProductID"]}</td>
          <td>${item["Description"]}</td>
          <td>${item["ManufacturingDate"]}</td>
          <td>${item["ExpiryDate"]}</td>

        </tr>

        `;

      });

      html += '</table>';

      const allItems: Element = this.domElement.querySelector('#listData');
      allItems.innerHTML = html;

    });

  }

  private deleteListItem(): void {

    const id = document.getElementById('#txtID')["value"];
    pnp.sp.web.lists.getByTitle("ProductCatalog").items.getById(id).delete();
    alert("List Item was Deleted Successfully");
  
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
