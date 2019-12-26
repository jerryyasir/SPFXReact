declare interface IReactHelloWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  MyGroupName: string;
  DescriptionFieldLabel: string;
  MyPropertyDescriptionLabel: string;
}

declare module "ReactHelloWebPartStrings" {
  const strings: IReactHelloWebPartStrings;
  export = strings;
}
