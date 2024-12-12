import * as React from 'react';
import styles from './SpFxMsalAuthDemo.module.scss';
import type { ISpFxMsalAuthDemoProps } from './ISpFxMsalAuthDemoProps';
import { escape } from '@microsoft/sp-lodash-subset';

export const SpFxMsalAuthDemo: React.FC<ISpFxMsalAuthDemoProps> = (props) => {
  const { applicationID, redirectUri, tenantUrl } = props;
  
    return (
      <section className={`${styles.spFxMsalAuthDemo}`}>
        <div className={styles.welcome}>
          <img alt="" src={require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Demo</h2>
          <div>Application ID: <strong>{escape(applicationID)}</strong></div>
          <div>Redirect URI: <strong>{escape(redirectUri)}</strong></div>
          <div>Tenant URL: <strong>{escape(tenantUrl)}</strong></div>
        </div>
      </section>
    )
}
