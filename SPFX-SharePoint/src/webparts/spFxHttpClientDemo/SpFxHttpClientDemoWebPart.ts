import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "SpFxHttpClientDemoWebPartStrings";
import SpFxHttpClientDemo from "./components/SpFxHttpClientDemo";
import { ISpFxHttpClientDemoProps } from "./components/ISpFxHttpClientDemoProps";
import { ISpFxHttpClientDemoState } from "./components/ISpFxHttpClientDemoState";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { ICountryListItem } from "../../models";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";

export default class SpFxHttpClientDemoWebPart extends BaseClientSideWebPart<
  ISpFxHttpClientDemoProps
> {
  private _countries: ICountryListItem[] = [];
  private get _isSharePoint(): boolean {
    return (
      Environment.type === EnvironmentType.SharePoint ||
      Environment.type === EnvironmentType.ClassicSharePoint
    );
  }
  private _getListItems(): Promise<ICountryListItem[]> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Average Performance')/items?$select=Id,Title,Cost,Performance_x0020_Category&$Filter=Cost le '10'`,
        SPHttpClient.configurations.v1
      )
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.value;
      }) as Promise<ICountryListItem[]>;
  }

  private _onGetListItems = (): void => {
    if (!this._isSharePoint) {
      if (this._countries.length === 0) {
        this._countries = [
          {
            Id: "1",
            Title: "Test Title 1",
            Cost: "200",
            Performance_x0020_Category: "Category 1"
          },
          {
            Id: "2",
            Title: "Test Title 2",
            Cost: "300",
            Performance_x0020_Category: "Category 2"
          },
          {
            Id: "3",
            Title: "Test Title 3",
            Cost: "400",
            Performance_x0020_Category: "Category 3"
          },
          {
            Id: "4",
            Title: "Test Title 4 ",
            Cost: "500",
            Performance_x0020_Category: "Category 4"
          }
        ];
      }
    } else {
      this._getListItems().then(response => {
        this._countries = response;
      });
    }
    this.render();
  };
  public render(): void {
    const element: React.ReactElement<ISpFxHttpClientDemoProps> = React.createElement(
      SpFxHttpClientDemo,
      {
        spListItems: this._countries,
        onGetListItems: this._onGetListItems,
        onAddListItem: this._onAddListItem,
        title: "",
        cost: "",
        category: ""
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private _onAddListItem = (): void => {
    if (!this._isSharePoint) {
      var newItem = {
        Id: "6",
        Title: "Test Title 1",
        Cost: "200",
        Performance_x0020_Category: "Category 1"
      };
      this._countries.push(newItem);
      this.render();
    } else {
      this._addListItem().then(() => {
        this._getListItems().then(response => {
          this._countries = response;
          this.render();
        });
      });
    }
  };

  private _getItemEntityType(): Promise<string> {
    return this.context.spHttpClient
      .get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Average Performance')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1
      )
      .then(response => {
        return response.json();
      })
      .then(jsonResponse => {
        return jsonResponse.ListItemEntityTypeFullName;
      }) as Promise<string>;
  }

  private _addListItem(): Promise<SPHttpClientResponse> {
    return this._getItemEntityType().then(spEntityType => {
      const request: any = {};
      request.body = JSON.stringify({
        Title: new Date().toUTCString(),
        Cost: 20000,
        "@odata.type": spEntityType
      });

      return this.context.spHttpClient.post(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getbytitle('Average Performance')/items`,
        SPHttpClient.configurations.v1,
        request
      );
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
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
                PropertyPaneTextField("description", {
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
