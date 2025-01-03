import * as React from 'react';
import type { ISpFxMsalAuthDemoProps } from './ISpFxMsalAuthDemoProps';

import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { AuthenticationContextProvider } from '../../context/AuthenticationContext';
import { GraphDataDisplay } from './GraphDataDisplay';


export const SpFxMsalAuthDemo: React.FC<ISpFxMsalAuthDemoProps> = (props) => {

  const { applicationID, tenantIdentifier, scopes, redirectUri } = props;

  return (
    <AuthenticationContextProvider clientId={applicationID} tenantId={tenantIdentifier} scopes={scopes} redirectUri={redirectUri}>
      <GraphDataDisplay httpClient={props.httpClient}
        applicationID={applicationID} tenantIdentifier={tenantIdentifier} redirectUri={redirectUri} apiCall={props.apiCall} />
    </AuthenticationContextProvider>
  )
}
