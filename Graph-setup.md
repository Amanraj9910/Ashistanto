# Microsoft Graph API Setup Guide

This guide will help you register an app in Azure and configure Microsoft Graph API access for your voice assistant.

## 📋 Prerequisites

- Azure subscription
- Microsoft 365 account (for testing with real data)
- Administrator access to Azure AD

## 🔧 Step 1: Register Application in Azure Portal

### 1.1 Go to Azure Portal
1. Navigate to https://portal.azure.com
2. Sign in with your Microsoft account
3. Search for "Azure Active Directory" or "Microsoft Entra ID"

### 1.2 Register New Application
1. In the left menu, click **App registrations**
2. Click **+ New registration**
3. Fill in the details:
   - **Name**: `Voice AI Assistant` (or any name you prefer)
   - **Supported account types**: Select one of:
     - "Accounts in this organizational directory only" (Single tenant)
     - "Accounts in any organizational directory" (Multi-tenant)
   - **Redirect URI**: Leave blank for now (we'll add later if needed)
4. Click **Register**

### 1.3 Note Your Application IDs
After registration, you'll see the **Overview** page. Copy these values:
- **Application (client) ID** → This is your `MICROSOFT_CLIENT_ID`
- **Directory (tenant) ID** → This is your `MICROSOFT_TENANT_ID`

Example:
```
Application (client) ID: 12345678-1234-1234-1234-123456789012
Directory (tenant) ID: 87654321-4321-4321-4321-210987654321
```

## 🔑 Step 2: Create Client Secret

### 2.1 Generate Secret
1. In your app registration, click **Certificates & secrets** in the left menu
2. Click **+ New client secret**
3. Add a description: `Voice Assistant Secret`
4. Choose expiration: **24 months** (or as per your security policy)
5. Click **Add**
6. **IMPORTANT**: Copy the **Value** immediately (not the Secret ID)
   - This is your `MICROSOFT_CLIENT_SECRET`
   - You can only see this once!

Example:
```
Client Secret Value: abc123~XyZ.456-qrs_789TUV
```

## 🔐 Step 3: Configure API Permissions

### 3.1 Add Microsoft Graph Permissions
1. Click **API permissions** in the left menu
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions** (for user context)

### 3.2 Select Required Permissions
Add these permissions based on what features you need:

#### For Email (Outlook):
- ✅ `Mail.Read` - Read user's emails
- ✅ `Mail.ReadWrite` - Read and write emails
- ✅ `Mail.Send` - Send emails as the user

#### For Calendar:
- ✅ `Calendars.Read` - Read user's calendar
- ✅ `Calendars.ReadWrite` - Read and write calendar events

#### For OneDrive/SharePoint:
- ✅ `Files.Read` - Read user's files
- ✅ `Files.ReadWrite` - Read and write files
- ✅ `Sites.Read.All` - Read SharePoint sites

#### For Microsoft Teams:
- ✅ `Team.ReadBasic.All` - Read basic team info
- ✅ `Channel.ReadBasic.All` - Read basic channel info

#### For User Profile:
- ✅ `User.Read` - Read user profile
- ✅ `User.ReadBasic.All` - Read basic profiles

### 3.3 Grant Admin Consent
1. After adding all permissions, click **Grant admin consent for [Your Organization]**
2. Confirm by clicking **Yes**
3. Wait for the status to show green checkmarks

## 📝 Step 4: Update Your .env File

Add these values to your `.env` file:

```env
# Microsoft Graph API
MICROSOFT_CLIENT_ID=your_application_client_id_here
MICROSOFT_CLIENT_SECRET=your_client_secret_value_here
MICROSOFT_TENANT_ID=your_directory_tenant_id_here
```

Example (with fake values):
```env
MICROSOFT_CLIENT_ID=12345678-1234-1234-1234-123456789012
MICROSOFT_CLIENT_SECRET=abc123~XyZ.456-qrs_789TUV
MICROSOFT_TENANT_ID=87654321-4321-4321-4321-210987654321
```

## 🔄 Step 5: Authentication Flow Options

You have two options for authentication:

### Option A: Application Permissions (Server-to-Server)
Best for: Background tasks, accessing data without user interaction

1. In Azure Portal → App Permissions
2. Choose **Application permissions** instead of Delegated
3. Add permissions like `Mail.Read`, `Calendars.Read`, etc.
4. Grant admin consent
5. Your app will access data on behalf of the organization

### Option B: Delegated Permissions (User Context) - RECOMMENDED
Best for: User-specific tasks, personal assistant

This requires user login. You have 3 approaches:

#### B1: Device Code Flow (Easiest for Testing)
User sees a code, goes to microsoft.com/devicelogin, enters code

#### B2: Authorization Code Flow (Web App)
Redirects user to Microsoft login page, returns with token

#### B3: Pre-authorize Token (Quick Testing)
Get token manually and add to .env

**To get a test token manually:**
1. Go to: https://developer.microsoft.com/en-us/graph/graph-explorer
2. Sign in with your Microsoft 365 account
3. Run any query (like "GET my profile")
4. Click "Access token" tab in the response
5. Copy the token
6. Add to `.env`:
```env
MICROSOFT_ACCESS_TOKEN=eyJ0eXAiOiJKV1QiLCJub25jZSI6...
```

⚠️ **Note**: Manual tokens expire in 1 hour. For production, implement proper OAuth flow.

## 🧪 Step 6: Test Your Setup

### 6.1 Install Dependencies
```bash
npm install @microsoft/microsoft-graph-client @azure/msal-node isomorphic-fetch
```

### 6.2 Test Graph Connection
Create a test file `test-graph.js`:

```javascript
require('dotenv').config();
const { getUserProfile } = require('./graph-tools');

async function test() {
  try {
    console.log('Testing Microsoft Graph connection...');
    const profile = await getUserProfile();
    console.log('✓ Success! User profile:', profile);
  } catch (error) {
    console.error('✗ Error:', error.message);
  }
}

test();
```

Run:
```bash
node test-graph.js
```

## ✅ Step 7: Verify Permissions

After setup, your app can now:

### 📧 Email Operations
```
"Check my recent emails"
"Search for emails from John"
"Send an email to john@example.com"
```

### 📅 Calendar Operations
```
"What's on my calendar today?"
"Show me my meetings this week"
"Schedule a meeting tomorrow at 2 PM"
```

### 📁 File Operations
```
"Show my recent files"
"Find the Q4 report document"
"What files did I access today?"
```

### 👥 Teams Operations
```
"What teams am I in?"
"Show my team channels"
```

### 👤 Profile Operations
```
"What's my job title?"
"Show my contact information"
```

## 🔒 Security Best Practices

1. **Never commit secrets to Git**
   - Add `.env` to `.gitignore`
   - Use Azure Key Vault in production

2. **Use least privilege**
   - Only request permissions you actually need
   - Review permissions regularly

3. **Rotate secrets regularly**
   - Client secrets expire - set reminders
   - Update secrets before expiration

4. **Monitor API usage**
   - Check Azure AD sign-in logs
   - Monitor for unusual activity

5. **Implement token refresh**
   - For delegated permissions, implement proper OAuth flow
   - Don't rely on manual tokens in production

## 🐛 Troubleshooting

### Error: "Unauthorized" or "Access denied"
- Check if admin consent was granted
- Verify API permissions are correct
- Ensure token has required scopes

### Error: "Invalid client"
- Check CLIENT_ID and TENANT_ID are correct
- Verify the app registration exists

### Error: "Invalid client secret"
- CLIENT_SECRET may be expired
- Create a new secret in Azure Portal
- Update .env file

### Error: "Token expired"
- Manual tokens expire in 1 hour
- Implement proper OAuth flow
- Or get a new token from Graph Explorer

### Error: "Insufficient privileges"
- User doesn't have permissions for the resource
- Grant additional API permissions
- Get admin consent again

## 📚 Additional Resources

- [Microsoft Graph Documentation](https://docs.microsoft.com/graph/)
- [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)
- [Permission Reference](https://docs.microsoft.com/graph/permissions-reference)
- [MSAL Node Documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js/tree/dev/lib/msal-node)

## 🎯 Next Steps

1. Complete the Azure app registration
2. Add credentials to `.env` file
3. Test the connection with `test-graph.js`
4. Start your voice assistant: `npm start`
5. Try voice commands like "Check my emails" or "What's on my calendar?"

Your voice assistant can now access Microsoft 365 data! 🎉