import * as React from "react"
import { PrimaryButton} from "@fluentui/react"
import { Textarea } from "@fluentui/react-components";
import styles from "./SpFxMsalAuthDemo.module.scss"
import { escape } from '@microsoft/sp-lodash-subset';
import { AuthenticationContext } from "../../context/AuthenticationContext";
import { HttpClient } from '@microsoft/sp-http';

type GraphDataDisplayProps = {
  applicationID: string;
  tenantIdentifier: string;
  redirectUri: string;
  apiCall: string;
  httpClient: HttpClient;
}

export const GraphDataDisplay: React.FC<GraphDataDisplayProps> = (props) => {
  const { isAuthenticated, userScopes, accessToken, reauthenticate } = React.useContext(AuthenticationContext);
  const { applicationID, tenantIdentifier, redirectUri, httpClient, apiCall } = props;

  const [graphApiCallData, setGraphApiCallData] = React.useState<string>();

  async function loadUserInfoFromGraph(): Promise<void> {
    console.log(`calling graph with access token: ${accessToken?.substring(0, 100)}...`);

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

  React.useEffect(() => {
    console.log('reloading user info from graph');
    loadUserInfoFromGraph().catch(console.error);
  }, [accessToken]);

  React.useEffect(() => {
    console.log('reloading user info from graph');
    loadUserInfoFromGraph().catch(console.error);
  }, [props]);

  function akquireNewToken(): void { 
    reauthenticate();
  } 

  return <section className={`${styles.spFxMsalAuthDemo}`}>
    <div className={styles.welcome}>
      <img alt="" src={require('../assets/welcome-light.png')} className={styles.welcomeImage} />
      <h2>Authentication Info (webpart properties)</h2>
      <div>Application ID: <strong>{escape(applicationID)}</strong></div>
      <div>Redirect URI: <strong>{escape(redirectUri)}</strong></div>
      <div>Tenant ID: <strong>{escape(tenantIdentifier)}</strong></div>
    </div>

    <div className={styles.welcome}>
      <h2>① Authentication Infos</h2>
      <div>Authenticated: {isAuthenticated ? 'Yes' : 'No'}</div>
    </div>

    {isAuthenticated && (
      <div className={styles.welcome}>
        <h2>② Entra ID (App registration)</h2>
        <h3>Current User Scopes (From Entra ID)</h3>
        <div><pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{userScopes}</pre></div>
        {isAuthenticated &&
          <>
            <Textarea value={accessToken} size="small" resize="both" />
            <div><PrimaryButton text='Akquire new token' onClick={() => { akquireNewToken() }} /></div>
            {/* <div><PrimaryButton text='Logout' onClick={() => { }} /></div> */}
          </>
        }
        {graphApiCallData && (
          <><h2>③ Graph API</h2>
          <h3>Return value(s) from API call</h3>
            <>{apiCall}</>
            <div><pre style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word' }}>{graphApiCallData}</pre></div>
          </>
        )}
      </div>
    )}
  </section>
}