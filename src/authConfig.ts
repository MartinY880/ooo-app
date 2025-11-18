import type { Configuration, PopupRequest } from "@azure/msal-browser";

/**
 * MSAL Configuration
 * 
 * Before deploying this application, you need to:
 * 1. Register your app in Azure Portal (Azure Active Directory > App registrations)
 * 2. Replace the placeholders below with your actual values:
 *    - clientId: Your Application (client) ID
 *    - authority: Your tenant ID (format: https://login.microsoftonline.com/{tenant-id})
 *    - redirectUri: Your application's redirect URI (e.g., http://localhost:5173)
 */
export const msalConfig: Configuration = {
  auth: {
    clientId: "4f61fb3c-386f-4fff-822e-34f909c4f5b4", // Replace with your Azure AD app's client ID
    authority: "https://login.microsoftonline.com/758227bf-0cdd-4ab2-88fa-71bda15be6f1", // Replace with your tenant ID
    redirectUri: "http://localhost:5174", // Update for production deployment
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "sessionStorage", // This configures where your cache will be stored
    storeAuthStateInCookie: false, // Set to true if you are having issues on IE11 or Edge
  },
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 */
export const loginRequest: PopupRequest = {
  scopes: [
    "User.Read",
    "User.Read.All",
    "MailboxSettings.ReadWrite",
    "Calendars.ReadWrite"
  ],
};

/**
 * Scopes for Microsoft Graph API
 * User.Read.All allows reading all users in the directory (required for user search)
 * You may need admin consent for this scope depending on your tenant configuration
 */
export const graphConfig = {
  graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
  graphUsersEndpoint: "https://graph.microsoft.com/v1.0/users",
  scopes: ["User.Read.All"], // Required for searching users across the organization
};
