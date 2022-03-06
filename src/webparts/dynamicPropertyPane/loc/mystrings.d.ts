declare interface IDynamicPropertyPaneWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'DynamicPropertyPaneWebPartStrings' {
  const strings: IDynamicPropertyPaneWebPartStrings;
  export = strings;
}
