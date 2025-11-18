import { AuthenticatedTemplate, UnauthenticatedTemplate, useMsal } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import { OofForm } from './components/OofForm';

/**
 * LoginButton Component
 * 
 * Provides a button to sign in using Microsoft Authentication.
 * Uses MSAL's useMsal hook to trigger the login redirect.
 */
function LoginButton() {
  const { instance } = useMsal();

  const handleLogin = async () => {
    try {
      await instance.loginRedirect(loginRequest);
    } catch (error) {
      console.error('Login error:', error);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 flex items-center justify-center px-4">
      <div className="bg-white rounded-2xl shadow-xl p-8 max-w-md w-full text-center">
        <div className="mb-6">
          <div className="mx-auto mb-6">
            <img 
              src="/MTGProsLogoTransparent-Large.png" 
              alt="MTG Pros Logo" 
              className="mx-auto h-24 w-auto"
            />
          </div>
          <h1 className="text-3xl font-bold text-gray-900 mb-2">Out of Office</h1>
          <p className="text-gray-600">Sign in to configure your automatic replies</p>
        </div>

        <button
          onClick={handleLogin}
          className="w-full bg-blue-600 hover:bg-blue-700 text-white font-medium py-3 px-6 rounded-lg transition-colors flex items-center justify-center space-x-2"
        >
          <svg 
            className="w-5 h-5" 
            viewBox="0 0 21 21" 
            fill="currentColor"
          >
            <path d="M0 0h9.996v9.996H0zm11.004 0H21v9.996H11.004zM0 11.004h9.996V21H0zm11.004 0H21V21H11.004z"/>
          </svg>
          <span>Sign in with Microsoft</span>
        </button>

        <p className="mt-4 text-xs text-gray-500">
          You'll be redirected to Microsoft login
        </p>
      </div>
    </div>
  );
}

/**
 * UserInfo Component
 * 
 * Displays the signed-in user's information and provides a logout button.
 */
function UserInfo() {
  const { instance, accounts } = useMsal();

  const handleLogout = () => {
    instance.logoutRedirect({
      postLogoutRedirectUri: window.location.origin,
    });
  };

  const account = accounts[0];

  return (
    <div className="bg-white border-b border-gray-200 shadow-sm">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="w-10 h-10 bg-blue-100 rounded-full flex items-center justify-center">
              <span className="text-blue-600 font-semibold text-lg">
                {account?.name?.charAt(0).toUpperCase() || '?'}
              </span>
            </div>
            <div>
              <p className="text-sm font-medium text-gray-900">{account?.name}</p>
              <p className="text-xs text-gray-500">{account?.username}</p>
            </div>
          </div>
          <div className="absolute left-1/2 transform -translate-x-1/2">
            <img 
              src="/MTGProsLogoTransparent-Large.png" 
              alt="MTG Pros Logo" 
              className="h-24 w-auto"
            />
          </div>
          <button
            onClick={handleLogout}
            className="px-4 py-2 text-sm font-medium text-gray-700 hover:text-gray-900 hover:bg-gray-100 rounded-lg transition-colors"
          >
            Sign Out
          </button>
        </div>
      </div>
    </div>
  );
}

/**
 * App Component
 * 
 * Main application component that handles authentication state.
 * Shows the login button for unauthenticated users and the OOF form
 * for authenticated users.
 */
function App() {
  return (
    <>
      {/* Show content for authenticated users */}
      <AuthenticatedTemplate>
        <UserInfo />
        <OofForm />
      </AuthenticatedTemplate>

      {/* Show login button for unauthenticated users */}
      <UnauthenticatedTemplate>
        <LoginButton />
      </UnauthenticatedTemplate>
    </>
  );
}

export default App;
