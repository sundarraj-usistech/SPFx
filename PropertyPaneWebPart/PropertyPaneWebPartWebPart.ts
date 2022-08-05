import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneLink
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart} from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PropertyPaneWebPartWebPart.module.scss';
import * as strings from 'PropertyPaneWebPartWebPartStrings';
import { BasicGroupName } from 'PropertyPaneWebPartWebPartStrings';

export interface IPropertyPaneWebPartWebPartProps {

  description: string;

  productname: string;
  productdescription: string;
  productcost: number;
  productquantity: number;
  billamount: number;
  discountamount: number;
  netamount: number;

  currenttime: Date;
  iscertified: boolean;

  range: number;

  choice: string;

  review: string;

  coupon: boolean;

}

export default class PropertyPaneWebPartWebPart extends BaseClientSideWebPart<IPropertyPaneWebPartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

  // protected onInit(): Promise<void>{

  //   return new Promise<void>((resolve, _reject) => {

  //     this.properties.productname = "Enter the Product Name";
  //     this.properties.productdescription = "Enter the Product Description";
  //     this.properties.productcost = +'Enter the Cost of the Product';
  //     this.properties.productquantity = +'Enter the Quantity of the Product';      
  //     this.properties.billamount = 0;
  //     this.properties.discountamount = +'10 %';
  //     this.properties.netamount = 0;

  //     resolve(undefined);

  //   });

  // }

protected get disableReactivePropertyChanges(): boolean {
    
    return false;

}

  public render(): void {

    let Certified : string = '';
  
    this.domElement.innerHTML = `
    <section class="${styles.propertyPaneWebPart} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="">

        <table>

          <tr>

          <td><b>Current Date and Time : </b></td>
          <td><b>${this.properties.currenttime = new Date()}</b></td>

          </tr>
      
          <tr>

          <td><b>Product Name : </b></td>
          <td><b>${this.properties.productname}</b></td>

          </tr>

          <tr>

          <td><b>Product Description : </b></td>
          <td><b>${this.properties.productdescription}</b></td>

          </tr>

          <tr>

          <td><b>Product Cost : </b></td>
          <td><b>${this.properties.productcost}</b></td>

          </tr>

          <tr>

          <td><b>Product Quantity : </b></td>
          <td><b>${this.properties.productquantity}</b></td>

          </tr>

          <tr>

          <td><b>Bill Amount : </b></td>
          <td><b>${this.properties.billamount = this.properties.productquantity * this.properties.productcost}</b></td>

          </tr>

          <tr>

          <td><b>Discount Amount : </b></td>
          <td><b>${this.properties.discountamount = this.properties.billamount * 0.10}</b></td>

          </tr>

          <tr>

          <td><b>Net Amount : </b></td>
          <td><b>${this.properties.netamount = this.properties.billamount - this.properties.discountamount}</b></td>

          </tr>

          <tr>

          <td><b>Certified : </b></td>
          <td><b>${Certified = (this.properties.iscertified === true) ? "Yes" : "No"}</b></td>

          </tr>

          <tr>

          <td><b>Rating : </b></td>
          <td><b>${this.properties.range}</b></td>

          </tr>

          <tr>

          <td><b>Recommend : </b></td>
          <td><b>${this.properties.choice}</b></td>

          </tr>

          <tr>

          <td><b>Review : </b></td>
          <td><b>${this.properties.review}</b></td>

          </tr>

          <tr>

          <td><b>Discount Coupon : </b></td>
          <td><b>${this.properties.coupon}</b></td>

          </tr>       

        </table>

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

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           // description: strings.PropertyPaneDescription
  //           description: "Description"
  //         },
  //         groups: [
  //           {
  //             groupName: strings.BasicGroupName,
  //             groupFields: [
  //               PropertyPaneTextField('description', {
  //                 // label: strings.DescriptionFieldLabel
  //                 label: "Description Label"
  //               })
  //             ]
  //           }
  //         ]
  //       }
  //     ]
  //   };
  // }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    return{

      pages: [

        {

          header: {

            description: "Product Details"

          },

          groups: [

            {

              groupName: "Product Details",
              groupFields: [

                PropertyPaneTextField ('productname',{

                  label: "Product Name",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 3500,
                  placeholder: "Please Enter the Product Name", "description": "Name property field"

                }),

                PropertyPaneTextField ('productdescription',{

                  label: "Product Description",
                  multiline: true,
                  resizable: false,
                  deferredValidationTime: 3500,
                  placeholder: "Please Enter the Product Description", "description": "Name property field"

                }),

                PropertyPaneTextField ('productcost',{

                  label: "Product Cost",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 3500,
                  placeholder: "Please Enter the Product Cost", "description": "Number property field"

                }),

                PropertyPaneToggle ('iscertified', {

                  key: "Certified",
                  label: "Certified",
                  onText: "Yes",
                  offText: "No"
              
                })            

              ]

            },

            {
              
              groupName: "Customer Needs",
              groupFields: [

                PropertyPaneTextField ('productquantity',{

                  label: "Product Quantity",
                  multiline: false,
                  resizable: false,
                  deferredValidationTime: 3500,
                  placeholder: "Please Enter the Product Quantity", "description": "Number property field"

                }),

                PropertyPaneCheckbox ('coupon',{

                  text: "Have a Discount Coupon",
                  checked: false,
                  disabled: false

                })

              ]

            }

          ],

          displayGroupsAsAccordion: true

        },

        {

          header: {

            description: "Customer Reviews"

          },

          groups: [{

            groupName: "Customer Reviews",
            groupFields: [

              PropertyPaneSlider ('range', {

                label: "Rating",
                min: 1,
                max: 10,
                showValue: true,
                step: 1,
                value: (isNaN(this.properties.range) ? 0 : this.properties.range)
              }),

              PropertyPaneChoiceGroup ('choice', {

                label: "Recommend",
                options:[

                  {key: 'Yes', text: 'Yes', checked: true},
                  {key: 'Later', text: 'Not Now'},
                  {key: 'No', text: 'Never'}

                ]
              }),

              PropertyPaneDropdown ('review', {

                label: "Review",
                options:[

                  {key: 'Good', text: 'Good'},
                  {key: 'Excellent', text: 'Excellent'},
                  {key: 'Poor', text: 'Poor'}

                ],
                selectedKey: 'Good'
              }),


            ]

          }]

        },

        {

          header: {

            description: "Amazon"

          },

          groups: [{

            groupName: "Amazon",
            groupFields: [

              PropertyPaneLink ('',{

                href: 'https://www.amazon.in/',
                text: 'Buy the Product from Amazon',
                target: '_blank',
                popupWindowProps: {

                  height: 500,
                  width: 500,
                  positionWindowPosition: 2,
                  title: 'Amazon'

                }

              })

            ]

          }]

        }

      ]

    };

  }

}
