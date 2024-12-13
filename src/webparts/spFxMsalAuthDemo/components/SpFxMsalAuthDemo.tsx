import * as React from 'react';
import styles from './SpFxMsalAuthDemo.module.scss';
import type { ISpFxMsalAuthDemoProps } from './ISpFxMsalAuthDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import type { MSALOptions } from "@pnp/msaljsclient";
import { spfi, SPBrowser } from "@pnp/sp";
import { MSAL } from "@pnp/msaljsclient";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { PublicClientApplication } from "@azure/msal-browser";
import { HttpClient } from '@microsoft/sp-http';


export const SpFxMsalAuthDemo: React.FC<ISpFxMsalAuthDemoProps> = (props) => {
  // clientId: '7d41c187-16f0-4fbb-b5e8-be4b282801df',
  // authority: 'https://login.microsoftonline.com/tmaestrini.onmicrosoft.com',

  const { applicationID, tenantIdentifier, scopes, redirectUri, apiCall } = props;
  const { httpClient } = props;

  const [userMail, setUserMail] = React.useState<string>();
  const [userScopes, setUserScopes] = React.useState<string>();
  const [accessToken, setAccessToken] = React.useState<string>();
  const [apiCallData, setApiCallData] = React.useState<string>();

  async function loginForSPOAccessTokenByMSAL(): Promise<void> {
    const spoOptions: MSALOptions = {
      configuration: {
        auth: {
          authority: "https://login.microsoftonline.com/tmaestrini.onmicrosoft.com/",
          clientId: "7d41c187-16f0-4fbb-b5e8-be4b282801df",
        },
        cache: {
          claimsBasedCachingEnabled: true // in order to avoid network call to refresh a token every time claims are requested
        }
      },
      authParams: {
        forceRefresh: false,
        scopes: ["https://tmaestrini.sharepoint.com/.default"],
      }
    };

    try {
      const sp = spfi("https://tmaestrini.sharepoint.com/").using(SPBrowser(), MSAL(spoOptions));

      const user = await sp.web.currentUser();
      setUserMail(user.Email);
    } catch (error) {
      console.error('SPO error:', error);
    }
  }

  async function loginForGraphAccessTokenByMSAL(): Promise<void> {
    const config = {
      auth: {
        clientId: applicationID,
        authority: `https://login.microsoftonline.com/${tenantIdentifier}`,
        // redirectUri: 'https://tmaestrini.sharepoint.com/_layouts/15/workbench.aspx',
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false,
      },
    };

    const msalInstance = new PublicClientApplication(config);
    await msalInstance.initialize();

    try {
      const result = await msalInstance.acquireTokenSilent({
        scopes: [...scopes.split(',')],
        account: msalInstance.getAllAccounts()[0]
      });
      console.log('Silent token result:', result);
      setUserScopes(result.scopes.join(', '));
      setAccessToken(result.accessToken);
    } catch (error) {
      console.error("Error acquiring token silently:", error);
    }
  }

  async function getUserInfoFromGraph(): Promise<void> {
    console.log(`calling graph with access token: ${accessToken}`);

    const response = await httpClient.get(apiCall, HttpClient.configurations.v1, {
    headers: {
        'Authorization': `Bearer ${accessToken}`
      }
    });

    try {
      if (response.ok) {
        const responseData = await response.json();
        console.log('response data:', responseData);
        setApiCallData(JSON.stringify(responseData, null, 2));
      } else {
        setApiCallData(`Error fetching data: ${response.status}`);
        console.error('Error fetching data:', response.statusText);
      }
    } catch (error) {
      console.error('HTTP request error:', error);
    }
  }

  React.useEffect(() => {
    loginForSPOAccessTokenByMSAL().catch(console.error);
    loginForGraphAccessTokenByMSAL().catch(console.error);
  }, []);

  React.useEffect(() => {
    loginForSPOAccessTokenByMSAL().catch(console.error);
    loginForGraphAccessTokenByMSAL().catch(console.error);
  }, [props, scopes]);

  React.useEffect(() => {
    getUserInfoFromGraph().catch(console.error);
  }, [accessToken, props]);

  return (
    <section className={`${styles.spFxMsalAuthDemo}`}>
      <div className={styles.welcome}>
        <img alt="" src={require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Demo</h2>
        <div>Application ID: <strong>{escape(applicationID)}</strong></div>
        <div>Redirect URI: <strong>{escape(redirectUri)}</strong></div>
        <div>Tenant URL: <strong>{escape(tenantIdentifier)}</strong></div>
      </div>

      {userMail && (
        <div className={styles.welcome}>
          <h2>Current User Info (From Entra ID)</h2>
          <div>User Mail: <strong>{userMail}</strong></div>
        </div>
      )}

      {userScopes && (
        <div className={styles.welcome}>
          <h2>Current User Scopes (From Entra ID)</h2>
          <div><pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{userScopes}</pre></div>
        </div>
      )}

      {apiCallData && (
        <div className={styles.welcome}>
          <h2>Return value from API call</h2>
          <div><pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{apiCallData}</pre></div>
        </div>
      )}
    </section>
  )
}
