import { HttpClient } from "@microsoft/sp-http";

export interface ISpFxMsalAuthDemoProps {
  applicationID: string;
  redirectUri: string;
  tenantUrl: string;
  httpClient: HttpClient;
  userMail: string;
}
