import * as React from 'react';
import styles from './SpFxMsalAuthDemo.module.scss';
import type { ISpFxMsalAuthDemoProps } from './ISpFxMsalAuthDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import type { MSALOptions } from "@pnp/msaljsclient";
import { spfi, SPBrowser } from "@pnp/sp";
import { MSAL } from "@pnp/msaljsclient";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";

export const SpFxMsalAuthDemo: React.FC<ISpFxMsalAuthDemoProps> = (props) => {
  const { applicationID, redirectUri, tenantUrl } = props;
  const [userMail, setUserMail] = React.useState<string>();

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


  React.useEffect(() => {
    loginForSPOAccessTokenByMSAL().catch(console.error);
  }, []);

  return (
    <section className={`${styles.spFxMsalAuthDemo}`}>
      <div className={styles.welcome}>
        <img alt="" src={require('../assets/welcome-light.png')} className={styles.welcomeImage} />
        <h2>Demo</h2>
        <div>Application ID: <strong>{escape(applicationID)}</strong></div>
        <div>Redirect URI: <strong>{escape(redirectUri)}</strong></div>
        <div>Tenant URL: <strong>{escape(tenantUrl)}</strong></div>
      </div>

      {userMail && (
        <div className={styles.welcome}>
          <h2>Current User Info</h2>
          <div>User Mail: <strong>{userMail}</strong></div>
        </div>
      )}
    </section>
  )
}
