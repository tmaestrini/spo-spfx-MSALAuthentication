import { HttpClient } from "@microsoft/sp-http";

export interface ISpFxMsalAuthDemoProps {
  applicationID: string;
  redirectUri: string;
  tenantIdentifier: string;
  scopes: string;
  apiCall: string;
  httpClient: HttpClient;
  userMail: string;
}
