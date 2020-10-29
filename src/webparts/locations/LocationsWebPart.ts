import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {IPropertyPaneConfiguration,PropertyPaneTextField} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {SPHttpClient,SPHttpClientResponse} from '@microsoft/sp-http';
import {Environment,EnvironmentType} from '@microsoft/sp-core-library';
import * as strings from 'LocationsWebPartStrings';
import styles from './locations.module.scss';

// #region [interfaces]

  export interface ILocationSPFxWebPartProps {
    title: string; 
    description: string;
    siteURL: string; 
    list: string; 
  }


  export interface ISPLists {
    value: ISPList[];
  }

  export interface ISPList {
    ID: number;
    Title: string;
    State: string;
    Address: string;
    City: string;
    Zip: string;
    Phone: string;
    Email: string;
  }

//#endregion


export default class LocationsWebPart extends BaseClientSideWebPart<ILocationSPFxWebPartProps> {

  // #region [RenderCode]

    public render(): void {
      this.domElement.innerHTML = `
      <div class="${styles.location}">
        <div class="${styles.container}">
          <span class="${styles.head}">${this.properties.title}</span>
          <span class="${styles.subhead}">${this.properties.description}</span>
          <div id="locations"></div>
        </div>
      </div>`;
      this._renderListAsync();
    }

    private _renderList(items: ISPList[]): void {
      let htmlout: string = "";     
      items.forEach((item: ISPList) => {
        htmlout += `<div class="${styles.location}">`;
        htmlout += `  <img src="${this.properties.siteURL}/SiteAssets/MapIcon.jpg" class="${styles.image}" />`;       
        htmlout += `  <span class="${styles.header}">${item.Title}</span><br>`;
        htmlout += `  <span class="${styles.data}">${item.Address}</span><br>`; 
        htmlout += `  <span class="${styles.data}">${item.City}, ${item.State} ${item.Zip} </span><br>`;  
        htmlout += `  <span class="${styles.data}">${item.Phone}</span><br>`;     
        htmlout += `  <span class="${styles.data}"><a href="mailto:${item.Email}">${item.Email}</a></span><br>`; 
        htmlout += `</div>`;
      });
      const listContainer: Element = this.domElement.querySelector('#locations');
      listContainer.innerHTML = htmlout;
    }
    
  //#endregion

 // #region [AsyncCode]

  private _renderListAsync(): void {
    //alert("_renderListAsync!");
    if (Environment.type == EnvironmentType.SharePoint ||  Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    } 
  }

  // #endregion

  // #region [SharePointQueriesCode] 
  
    private _getListData(): Promise<ISPLists> {
      let today : Date = new Date();
      let restQuery : string = `${this.properties.siteURL}/_api/Web/Lists/GetByTitle('${this.properties.list}')/items?`;
      restQuery += "&$select=Id,Title,State,Address,City,Zip,Phone,Email"; 
      return this.context.spHttpClient.get( restQuery ,SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
    }

  // #endregion

  // #region [GenericCode]

    protected onDispose(): void {
      ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
      return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      return {
        pages: [
          {
            header: {
              description: "Locations SPFx Webpart Setup"
            },
            groups: [
              {
                groupName: "Display Settings",
                groupFields: [
                  PropertyPaneTextField('title', {
                    label: "WebPart Title"
                  }),
                  PropertyPaneTextField('description', {
                    label: "WebPart Description"
                  })
                ]
              },
              {
                groupName: "Data Source Configuration",
                groupFields: [
                  PropertyPaneTextField('siteURL', {
                    label: "Site URL"
                  }),
                  PropertyPaneTextField('list', {
                    label: "List Name"
                  })
                ]
              }
            ]
          }
        ]
      };
    }

  //#endregion

}
