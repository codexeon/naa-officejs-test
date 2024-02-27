import React, { useEffect, useRef, useState } from 'react';
import './App.css';
import {
  AccountInfo,
  AuthError,
  AuthenticationResult,
  Configuration,
  IPublicClientApplication,
  InteractionRequiredAuthError,
  PublicClientNext,
  ServerError,
} from '@azure/msal-browser';
import { act } from 'react-dom/test-utils';

interface AuthContext {
  tenantId: string;
  userPrincipalName: string;
  authorityType: 'aad' | 'msa' | 'other';
}

const clientId = 'be4080bc-cd4b-42ae-b970-5e9ac7064f63';
function getConfig(authority?: string): Configuration {
  return {
    auth: {
      clientId,
      authority: authority ?? 'https://login.microsoftonline.com/common',
      redirectUri: window.location.origin,
      supportsNestedAppAuth: true,
    },
    cache: {
      cacheLocation: 'localStorage',
    },
  };
}

function App() {
  const [scope, setScope] = useState('user.read');
  const [claims, setClaims] = useState('');
  const [isPopup, setIsPopup] = useState(false);
  const [output, setOutput] = useState<any>(null);
  const [timestamp, setTimestamp] = useState('');
  const [accessToken, setAccessToken] = useState('');
  const msalInstance = useRef<IPublicClientApplication | undefined>(undefined);
  const account = useRef<AccountInfo | undefined>(undefined);
  const authContext = useRef<AuthContext | undefined>(undefined);

  useEffect(() => {
    (async () => {
      const result = await Office.onReady();
      console.log('Office initialized', result);
    })();
  }, []);

  const clearOutput = () => {
    setOutput(null);
    setTimestamp('');
  };

  const tryGetAuthContext = async () => {
    if (!authContext.current) {
      try {
        authContext.current = (await (
          Office.auth as any
        ).getAuthContext()) as AuthContext;
      } catch (ex) {
        console.error('Error getting auth context', ex);
      }
    }
    return authContext.current;
  };

  const acquireTokenMsalJs = async () => {
    const startTime = new Date();
    const authContext = await tryGetAuthContext();
    let loginHint = authContext?.userPrincipalName;

    if (msalInstance.current == null) {
      // Initialize MSAL.js
      let authority: string | undefined;
      if (authContext?.authorityType === 'msa') {
        authority = 'https://login.microsoftonline.com/consumers';
      } else if (authContext?.tenantId) {
        authority = `https://login.microsoftonline.com/${authContext.tenantId}`;
      }

      msalInstance.current =
        await PublicClientNext.createPublicClientApplication(
          getConfig(authority)
        );
      // Try to find the account in cache
      if (account.current == null) {
        if (authContext) {
          const accounts = msalInstance.current.getAllAccounts();
          account.current = accounts.find(
            (acc) =>
              acc.tenantId.toLowerCase() ===
                authContext.tenantId.toLowerCase() &&
              acc.username.toLowerCase() ===
                authContext.userPrincipalName.toLowerCase()
          );
        }
        if (account.current == null) {
          const activeAccount = msalInstance.current.getActiveAccount();
          if (activeAccount) {
            account.current = activeAccount;
          }
        }
      }
    }

    const pca = msalInstance.current;

    const requestParams = {
      scopes: scope.split(' '),
      claims,
      account: account.current,
      loginHint: loginHint,
    };

    let response: AuthenticationResult | null = null;
    try {
      if (isPopup) {
        response = await pca.acquireTokenPopup(requestParams);
      } else {
        if (account.current == null) {
          response = await pca.ssoSilent(requestParams);
        } else {
          response = await pca.acquireTokenSilent(requestParams);
        }
      }
      account.current = response.account;
      setOutput(response);
      setAccessToken(response.accessToken);
    } catch (ex) {
      let authError = ex as AuthError;
      let authErrorType = 'unknown';
      if (authError instanceof InteractionRequiredAuthError) {
        authErrorType = 'interaction_required';
      } else if (authError instanceof ServerError) {
        authErrorType = 'server_error';
      }
      setOutput({
        error: authError,
        code: authError.errorCode,
        message: authError.errorMessage,
        subError: authError.subError,
        authErrorType,
      });
    }

    const endTime = new Date();
    setTimestamp(`${endTime.getTime() - startTime.getTime()}ms`);
  };

  const prepopulateClaims = () => {
    const claims = {
      access_token: {
        nbf: {
          essential: true,
          value: Math.floor(new Date().getTime() / 1000 - 300).toString(),
        },
      },
    };
    setClaims(JSON.stringify(claims));
  };

  const makeGraphCall = async (endpointUrl: string) => {
    let files = await fetch(endpointUrl, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${accessToken}`,
        'Content-Type': 'application/json',
      },
    });
    let responseJson = await files.json();
    setOutput(responseJson);
  };
  const profileGraphCall = () => {
    makeGraphCall('https://graph.microsoft.com/v1.0/me');
  };

  return (
    <div className="App">
      <div className="App-content">
        <h1>NAA Test App</h1>
        <div className="input-group">
          <label htmlFor="scope-input" className="input-label">
            Scope:
          </label>
          <input
            type="text"
            id="scope-input"
            placeholder="Enter scope"
            value={scope}
            onChange={(e) => setScope(e.target.value)}
          />
        </div>
        <div className="input-group">
          <label htmlFor="claims-input" className="input-label">
            Claims:
          </label>
          <input
            type="text"
            id="claims-input"
            placeholder="Enter claims"
            value={claims}
            onChange={(e) => setClaims(e.target.value)}
          />
          <button onClick={prepopulateClaims}>Pre populate</button>
        </div>
        <div className="input-group checkbox-group">
          <label htmlFor="isPopup" className="input-label">
            Is Popup:
          </label>
          <input
            type="checkbox"
            id="isPopup"
            checked={isPopup}
            onChange={(e) => setIsPopup(e.target.checked)}
          />
        </div>
        <button onClick={acquireTokenMsalJs}>Acquire Token MSAL.js</button>
        <button onClick={clearOutput}>Clear output</button>
        <button onClick={profileGraphCall}>Profile Graph Call</button>
        <button
          onClick={() => window.location.reload()}
          className="reload-button"
        >
          Reload
        </button>
        {timestamp && <p>Time for request: {timestamp}</p>}
        {output && (
          <div className="token-info">
            <h2>Token Information:</h2>
            <pre>{JSON.stringify(output, null, 2)}</pre>
          </div>
        )}
        <div>Running at: {window.location.origin}</div>
      </div>
    </div>
  );
}

export default App;
