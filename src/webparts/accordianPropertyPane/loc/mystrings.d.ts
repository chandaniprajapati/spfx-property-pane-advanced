declare interface IAccordianPropertyPaneWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'AccordianPropertyPaneWebPartStrings' {
  const strings: IAccordianPropertyPaneWebPartStrings;
  export = strings;
}
