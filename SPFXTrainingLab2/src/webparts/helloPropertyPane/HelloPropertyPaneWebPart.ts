import { Version } from "@microsoft/sp-core-library";
import { PropertyPaneSlider } from "@microsoft/sp-webpart-base";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";

import styles from "./HelloPropertyPaneWebPart.module.scss";
import * as strings from "HelloPropertyPaneWebPartStrings";
import {
  PropertyFieldCollectionData,
  CustomCollectionFieldType
} from "@pnp/spfx-property-controls/lib/PropertyFieldCollectionData";
import {
  IPropertyFieldGroupOrPerson,
  PropertyFieldPeoplePicker,
  PrincipalType
} from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

export interface IHelloPropertyPaneWebPartProps {
  description: string;
  myContinent: string;
  numContinentsVisited: number;
  people: IPropertyFieldGroupOrPerson[];
  expansionOptions: any[];
}

export default class HelloPropertyPaneWebPart extends BaseClientSideWebPart<
  IHelloPropertyPaneWebPartProps
> {
  public render(): void {
    console.log(this.properties.people);
    if (this.properties.people && this.properties.people.length > 0) {
      let peopleList: string = "";
      this.properties.people.forEach(person => {
        peopleList = peopleList + `<li>${person.fullName} ${person.email}</li>`;
        console.log(`<li>${person.fullName} ${person.email}</li>`);
      });

      this.domElement.getElementsByClassName(
        "selectedPeople"
      )[0].innerHTML = `<ul>${peopleList}</ul>`;
    }
    this.domElement.innerHTML = `
  <div class="${styles.helloPropertyPane}">
    <div class="${styles.container}">
      <div class="${styles.row}">
        <div class="${styles.column}">
          <span class="${styles.title}">Welcome to SharePoint!</span>
          <p class="${
            styles.subTitle
          }">Customize SharePoint experiences using Web Parts.</p>
          <p class="${styles.description}">${escape(
      this.properties.description
    )}</p>
          <p class="${styles.description}">Continents where I reside: ${escape(
      this.properties.myContinent
    )}</p>
          <p class="${styles.description}">Number of Contents I've visted: ${
      this.properties.numContinentsVisited
    }</p>
          <div class="${styles.selectedPeople}"></div>
          <div class=${styles.expansionElement}"></div>
        </div>
      </div>
    </div>
  </div>`;

    if (
      this.properties.expansionOptions &&
      this.properties.expansionOptions.length > 0
    ) {
      let expansionOptions: string = "";
      this.properties.expansionOptions.forEach(option => {
        expansionOptions =
          expansionOptions +
          `<li>${option["Region"]} :${option["Comment"]}</li>`;
      });
      if (expansionOptions.length > 0) {
        this.domElement.getElementsByClassName(
          "expansionElement"
        )[0].innerHTML = `<ul>${expansionOptions}</ul>`;
      }
    }
  }

  private validateContinents(textboxValue: string): string {
    const validContinentOptions: string[] = [
      "africa",
      "antartica",
      "asia",
      "australia",
      "europe",
      "north america",
      "south america"
    ];
    const inputToValidate: string = textboxValue.toLowerCase();

    return validContinentOptions.indexOf(inputToValidate) === -1
      ? "Invalid content entry; valid options are afraica etc"
      : "";
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
                }),
                PropertyPaneTextField("myContinent", {
                  label: "Continent where I currently ,resie",
                  onGetErrorMessage: this.validateContinents.bind(this)
                }),
                PropertyPaneSlider("numContinentsVisited", {
                  label: "Nmber of Contents I've visited",
                  min: 1,
                  max: 7,
                  showValue: true
                }),
                PropertyFieldPeoplePicker("people", {
                  label: "Select Project Managers",
                  initialData: this.properties.people,
                  allowDuplicate: false,
                  principalType: [
                    PrincipalType.Users,
                    PrincipalType.SharePoint,
                    PrincipalType.Security
                  ],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  context: this.context,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "peopleFieldId"
                }),
                PropertyFieldCollectionData("expansionOptions", {
                  key: "collectionData",
                  label: "Possible expansion options",
                  panelHeader: "Possible expansion options",
                  manageBtnLabel: "Manage expansion options",
                  value: this.properties.expansionOptions,
                  fields: [
                    {
                      id: "Region",
                      title: "Region",
                      required: true,
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key: "Northeast",
                          text: "Northeast"
                        },
                        {
                          key: "Northwest",
                          text: "Northwest"
                        },
                        {
                          key: "Southeast",
                          text: "Southeast"
                        },
                        {
                          key: "Southwest",
                          text: "Southwest"
                        },
                        {
                          key: "North",
                          text: "North"
                        },
                        {
                          key: "South",
                          text: "South"
                        }
                      ]
                    },
                    {
                      id: "Comment",
                      title: "Comment",
                      type: CustomCollectionFieldType.string
                    }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
