import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { PublicClientApplication } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import { msalConfig } from './authConfig';
import './index.css';
import App from './App.tsx';

/**
 * Initialize MSAL instance
 * 
 * This creates a PublicClientApplication with the configuration defined in authConfig.ts
 * The instance handles all authentication operations including token acquisition and caching
 */
const msalInstance = new PublicClientApplication(msalConfig);

// Initialize MSAL and handle redirect promise before rendering
msalInstance.initialize().then(() => {
  msalInstance.handleRedirectPromise().then(() => {
    /**
     * Render the application
     * 
     * The MsalProvider wraps the entire application, making MSAL context available
     * to all child components. This enables authentication throughout the app.
     */
    createRoot(document.getElementById('root')!).render(
      <StrictMode>
        <MsalProvider instance={msalInstance}>
          <App />
        </MsalProvider>
      </StrictMode>,
    );
  }).catch((error) => {
    console.error('Redirect handling error:', error);
  });
}).catch((error) => {
  console.error('MSAL initialization error:', error);
});
