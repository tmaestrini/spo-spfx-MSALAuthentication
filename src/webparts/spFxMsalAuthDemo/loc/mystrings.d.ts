declare interface ISpFxMsalAuthDemoWebPartStrings {
  ApplicationIDFieldLabel: string;
  RedirectUriFieldLabel: string;
  TenantUrlFieldLabel: string;
  ScopesFieldLabel: string;
  ApiCallFieldLabel: string;

  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'SpFxMsalAuthDemoWebPartStrings' {
  const strings: ISpFxMsalAuthDemoWebPartStrings;
  export = strings;
}
