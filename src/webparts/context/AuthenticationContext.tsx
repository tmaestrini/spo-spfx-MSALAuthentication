import { Configuration, PublicClientApplication } from "@azure/msal-browser";
import * as React from "react";


interface AuthenticationContextProps {
  isAuthenticated: boolean;
  userScopes?: string;
  accessToken?: string;
  reauthenticate: () => void;
}

interface AuthenticationContextProviderProps extends React.PropsWithChildren<{}> {
  clientId: string;
  tenantId: string;
  scopes?: string;
  redirectUri?: string;
}

export const AuthenticationContext = React.createContext<AuthenticationContextProps>({
  isAuthenticated: false,
  reauthenticate: () => { },
});

export const AuthenticationContextProvider = (props: AuthenticationContextProviderProps): JSX.Element => {
  const [isAuthenticated, setIsAuthenticated] = React.useState<boolean>(false);
  const [msalObj, setMsalInstance] = React.useState<PublicClientApplication | undefined>(undefined);
  const [userScopes, setUserScopes] = React.useState<string>('');
  const [accessToken, setAccessToken] = React.useState<string>('');

  const config: Configuration = {
    auth: {
      clientId: props.clientId,
      authority: `https://login.microsoftonline.com/${props.tenantId}`,
      redirectUri: props.redirectUri ?? window.location.origin,
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: false,
    },
  };

  async function initializeMsal(): Promise<void> {
    setIsAuthenticated(false);
    try {
      const msalObj = new PublicClientApplication(config);
      await msalObj.initialize();
      setMsalInstance(msalObj);
    } catch (error) {
      console.error("Error creating MSAL auth object:", error);
    }
  }

  async function login(): Promise<void> {
    try {
      if (msalObj) {
        const result = await msalObj.acquireTokenSilent({
          account: msalObj.getAllAccounts()[0],
          scopes: props.scopes ? [...props.scopes.split(',')] : [],
        });

        console.log('Silent token result:', result);

        if (msalObj && result.accessToken) {
          const accounts = msalObj.getAllAccounts();
          setIsAuthenticated(accounts.length > 0);
          setUserScopes(result.scopes.join(', '));
          setAccessToken(result.accessToken);
        }
      }
    } catch (error) {
      console.error("Error acquiring token silently:", error);
    }
  }

  function reauthenticate(): void {
    console.log('reauthenticating');
    initializeMsal().then(() => {
      console.log('MSAL initialized');
    }).catch(err => {
      console.error('Error initializing MSAL:', err);
    });
  }
  
  React.useEffect(() => {
    initializeMsal().catch(err => {
      console.error('Error initializing MSAL:', err);
    }
    );
  }, [props.scopes, props.clientId, props.tenantId, props.redirectUri]);

  React.useEffect(() => {
    login().catch(console.error);
  }, [msalObj]);

  return (
    <AuthenticationContext.Provider value={{
      isAuthenticated, userScopes, accessToken,
      reauthenticate
    }}>
      {props.children}
    </AuthenticationContext.Provider>
  );

}