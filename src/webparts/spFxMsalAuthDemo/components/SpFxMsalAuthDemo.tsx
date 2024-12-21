import * as React from 'react';
import styles from './SpFxMsalAuthDemo.module.scss';
import type { ISpFxMsalAuthDemoProps } from './ISpFxMsalAuthDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import type { MSALOptions } from "@pnp/msaljsclient";
import { spfi, SPBrowser } from "@pnp/sp";
import { MSAL } from "@pnp/msaljsclient";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import { Configuration, PublicClientApplication } from "@azure/msal-browser";
import { HttpClient } from '@microsoft/sp-http';
import { PrimaryButton } from '@fluentui/react';


export const SpFxMsalAuthDemo: React.FC<ISpFxMsalAuthDemoProps> = (props) => {
  const { applicationID, tenantIdentifier, scopes, redirectUri, apiCall } = props;
  const { httpClient } = props;

  const [msalInstance, setMsalInstance] = React.useState<PublicClientApplication>();
  const [userMail, setUserMail] = React.useState<string>();
  const [userScopes, setUserScopes] = React.useState<string>();
  const [accessToken, setAccessToken] = React.useState<string>();
  const [graphApiCallData, setGraphApiCallData] = React.useState<string>();
  const [isLoggedIn, setIsLoggedIn] = React.useState<boolean>(false);

  async function loginForSPOAccessTokenByMSAL(): Promise<void> {
    const spoOptions: MSALOptions = {
      configuration: {
        auth: {
          clientId: applicationID,
          authority: `https://login.microsoftonline.com/${tenantIdentifier}`,
        },
        cache: {
          claimsBasedCachingEnabled: true // in order to avoid network call to refresh a token every time claims are requested
        }
      },
      authParams: {
        forceRefresh: false,
        scopes: [...scopes.split(',')],
        // scopes: ["https://tmaestrini.sharepoint.com/.default"],
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
    const config: Configuration = {
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
    setIsLoggedIn(false);

    const msalInstance = new PublicClientApplication(config);
    await msalInstance.initialize();
    try {
      setMsalInstance(msalInstance);
      const result = await msalInstance.acquireTokenSilent({
        scopes: scopes ? [...scopes.split(',')] : [],
        account: msalInstance.getAllAccounts()[0]
      });

      console.log('Silent token result:', result);
      setUserScopes(result.scopes.join(', '));
      setAccessToken(result.accessToken);
      setIsLoggedIn(true);
    } catch (error) {
      console.error("Error acquiring token silently:", error);
    }
  }

  async function logout(): Promise<void> {
    if (msalInstance) {
      await msalInstance.logoutPopup();
      setIsLoggedIn(false);
      window.location.reload();
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
        setGraphApiCallData(JSON.stringify(responseData, null, 2));
      } else {
        setGraphApiCallData(`Error fetching data: ${response.status}`);
        console.error('Error fetching data:', response.statusText);
      }
    } catch (error) {
      console.error('HTTP request error:', error);
    }
  }

  React.useMemo(() => {
    console.log('reloading');
    loginForSPOAccessTokenByMSAL().catch(console.error);
    loginForGraphAccessTokenByMSAL().catch(console.error);
  }, [props, scopes]);

  React.useMemo(() => {
    console.log('reloading user info from graph');
    getUserInfoFromGraph().catch(console.error);
  }, [accessToken, isLoggedIn]);

  return (
    <section className={`${styles.spFxMsalAuthDemo}`}>
      <div className={styles.welcome}>
        <img alt="" src={require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Authentication Info (webpart properties)</h2>
        <div>Application ID: <strong>{escape(applicationID)}</strong></div>
        <div>Redirect URI: <strong>{escape(redirectUri)}</strong></div>
        <div>Tenant URL: <strong>{escape(tenantIdentifier)}</strong></div>
      </div>

      {userMail && (
        <div className={styles.welcome}>
          <h2>① Current User Info From SPO</h2>
          <div>User Mail: <strong>{userMail}</strong></div>
        </div>
      )}

      {userScopes && (
        <div className={styles.welcome}>
          <h2>② Microsoft Graph</h2>
          <h3>Current User Scopes (From Entra ID)</h3>
          <div><pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{userScopes}</pre></div>
          {isLoggedIn &&
            <>
              <div><PrimaryButton text='Akquire new token' onClick={loginForGraphAccessTokenByMSAL} /></div>
              <div><PrimaryButton text='Logout' onClick={logout} /></div>
            </>
          }
          {graphApiCallData && (
            <><h3>③ Microsoft Graph: Return value from API call</h3>
              <>{apiCall}</>
              <div><pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{graphApiCallData}</pre></div>
            </>
          )}
        </div>
      )}


    </section>
  )
}
