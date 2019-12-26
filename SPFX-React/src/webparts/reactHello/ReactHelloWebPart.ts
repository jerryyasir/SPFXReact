import * as React from "react";
import * as ReactDom from "react-dom";
import {
  Version,
  DisplayMode,
  Environment,
  EnvironmentType
} from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import * as strings from "ReactHelloWebPartStrings";
import ReactHello from "./components/ReactHello";
import { IReactHelloProps } from "./components/IReactHelloProps";
import styles from "./components/ReactHello.module.scss";

export interface IReactHelloWebPartProps {
  description: string;
  areYouGoodWithReact: string;
  myContinent: string;
  numContinentsVisited: number;
}

export default class ReactHelloWebPart extends BaseClientSideWebPart<
  IReactHelloWebPartProps
> {
  public render(): void {
    const element: React.ReactElement<IReactHelloProps> = React.createElement(
      ReactHello,
      {
        description: this.properties.description,
        areYouGoodWithReact: this.properties.areYouGoodWithReact
      }
    );

    ReactDom.render(element, this.domElement);
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
            },
            {
              groupName: strings.MyGroupName,
              groupFields: [
                PropertyPaneTextField("areYouGoodWithReact", {
                  label: strings.MyPropertyDescriptionLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
