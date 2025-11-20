# Out of Office Web Application

A modern React + TypeScript web application that allows internal employees to set their Out of Office automatic replies and configure email forwarding using Microsoft Graph API.

## 🚀 Features

- **Microsoft Authentication** - Secure sign-in using MSAL.js
- **Time-based OOF Settings** - Set start and end times for automatic replies
- **Dual Message Support** - Different messages for internal and external recipients
- **User Search** - Real-time user search with debouncing via Microsoft Graph API
- **Email Forwarding** - Optional email forwarding to another user
- **Modern UI** - Clean, responsive interface built with Tailwind CSS
- **Form Validation** - Comprehensive client-side validation
- **Power Automate Integration** - Submit settings to a webhook for processing

## 📋 Prerequisites

Before you begin, ensure you have the following:

1. **Node.js** (v18 or higher)
2. **Azure AD App Registration** with the following:
   - Client ID
   - Tenant ID
   - API Permissions: `User.Read`, `User.Read.All`
   - Redirect URI configured for your app
3. **Power Automate Webhook URL** (for form submission)

## 🛠️ Setup Instructions

### 1. Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** > **App registrations** > **New registration**
3. Configure your app:
   - **Name**: Out of Office App
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Web - `http://localhost:5173`
4. After registration, note your **Application (client) ID**
5. Go to **Authentication** > Enable **Access tokens** and **ID tokens**
6. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**
   - Add `User.Read` (should be added by default)
   - Add `User.Read.All` (requires admin consent)
7. Click **Grant admin consent** for your organization

### 2. Configure the Application

1. Open `src/authConfig.ts`
2. Replace the placeholder values:

```typescript
export const msalConfig: Configuration = {
  auth: {
    clientId: "YOUR_CLIENT_ID_HERE", // Replace with your Application (client) ID
    authority: "https://login.microsoftonline.com/YOUR_TENANT_ID_HERE", // Replace with your Tenant ID
    redirectUri: "http://localhost:5173", // Update for production
  },
  // ...
};
```

### 3. Configure Power Automate Webhook

1. Open `src/components/OofForm.tsx`
2. Find the `handleSubmit` function
3. Replace `YOUR_POWER_AUTOMATE_WEBHOOK_URL` with your actual webhook URL:

```typescript
const response = await fetch('YOUR_POWER_AUTOMATE_WEBHOOK_URL', {
  method: 'POST',
  headers: {
    'Content-Type': 'application/json',
  },
  body: JSON.stringify(formData),
});
```

### 4. Install Dependencies

```powershell
npm install
```

### 5. Run the Development Server

```powershell
npm run dev
```

The application will be available at `http://localhost:5173`

## 📁 Project Structure

```
OutofOffice/
├── src/
│   ├── components/
│   │   ├── OofForm.tsx          # Main form component
│   │   └── UserSearch.tsx       # User search with autocomplete
│   ├── hooks/
│   │   └── useDebounce.ts       # Debounce custom hook
│   ├── authConfig.ts            # MSAL configuration
│   ├── App.tsx                  # Main app with auth templates
│   ├── main.tsx                 # Entry point with MSAL provider
│   └── index.css                # Tailwind CSS imports
├── tailwind.config.js           # Tailwind configuration
├── postcss.config.js            # PostCSS configuration
├── package.json
└── README.md
```

## 🔑 Key Components

### Authentication Flow

1. **main.tsx**: Initializes MSAL instance and wraps app with `MsalProvider`
2. **App.tsx**: Manages authentication state using `AuthenticatedTemplate` and `UnauthenticatedTemplate`
3. **authConfig.ts**: Contains all MSAL configuration and Graph API settings

### User Search Component

The `UserSearch` component implements:
- **Debounced Search**: Uses custom `useDebounce` hook (400ms delay)
- **Graph API Integration**: Queries users with `$search` parameter
- **Required Headers**: Includes `ConsistencyLevel: eventual` for search queries
- **Token Acquisition**: Uses `acquireTokenSilent` to get access tokens
- **Real-time Results**: Shows results as user types (minimum 2 characters)

### Form Component

The `OofForm` component includes:
- Start/End time inputs
- Internal/External message textareas
- Conditional forwarding section
- Form validation
- Status messages (success/error)
- Power Automate webhook integration

## 📤 Form Data Structure

The form submits the following JSON structure:

```json
{
  "startTime": "2024-12-01T09:00",
  "endTime": "2024-12-15T17:00",
  "internalMessage": "I am currently out of the office...",
  "externalMessage": "Thank you for your email...",
  "enableForwarding": true,
  "forwardToEmail": "user@example.com",
  "forwardToName": "John Doe"
}
```

## 🎨 Styling

The application uses Tailwind CSS with:
- Responsive design (mobile-first approach)
- Gradient backgrounds
- Custom color scheme (blue/indigo theme)
- Smooth transitions and hover effects
- Professional form styling

## 🔒 Security Considerations

1. **Token Storage**: Session storage is used for MSAL tokens
2. **Scopes**: Minimal required scopes (`User.Read`, `User.Read.All`)
3. **HTTPS**: Use HTTPS in production
4. **Admin Consent**: `User.Read.All` requires admin consent
5. **Redirect URI**: Configure proper redirect URIs for production

## 🚢 Production Deployment

1. Update `redirectUri` in `authConfig.ts` to your production URL
2. Add production URL to Azure AD redirect URIs
3. Build the application:
   ```powershell
   npm run build
   ```
4. Deploy the `dist` folder to your hosting service
5. Ensure HTTPS is enabled

## 🐳 Docker Image

Build the production image (no Azure AD values needed at build time):

```powershell
docker build -t ooo-app .
```

Run the container locally and inject your Azure AD configuration at runtime:

```powershell
docker run `
  -e VITE_AZURE_CLIENT_ID=<client-id> `
  -e VITE_AZURE_TENANT_ID=<tenant-id> `
  -e VITE_REDIRECT_URI=https://your-host-name `
  -p 8080:80 ooo-app
```

The site will be available at `http://localhost:8080`. Update or rotate the environment variables without rebuilding the image; the entrypoint regenerates the runtime config on each container start.

## 🐛 Troubleshooting

### Login Issues
- Verify your Client ID and Tenant ID are correct
- Check that redirect URI matches exactly (including trailing slashes)
- Ensure popup blockers are disabled

### User Search Not Working
- Verify `User.Read.All` permission is granted
- Check that admin consent has been provided
- Ensure the `ConsistencyLevel: eventual` header is included

### API Errors
- Check browser console for detailed error messages
- Verify access tokens are being acquired successfully
- Test Graph API calls in [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)

## 📚 Additional Resources

- [MSAL.js Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-overview)
- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview)
- [React Documentation](https://react.dev)
- [Tailwind CSS](https://tailwindcss.com)
- [Vite Documentation](https://vitejs.dev)

## 📄 License

This project is for internal use only.

## 👥 Support

For issues or questions, contact your IT department or project administrator.

