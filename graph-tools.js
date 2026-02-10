require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ConfidentialClientApplication } = require('@azure/msal-node');
const formatters = require('./formatters');
const timezoneHelper = require('./timezone-helper');

// Get user email from environment variable
const userEmail = process.env.MICROSOFT_USER_EMAIL;

// Initialize MSAL as a SINGLETON so the internal token cache persists
// This is CRITICAL for token refresh - MSAL stores refresh tokens in its cache
let _msalClientInstance = null;

function getMsalClient() {
  if (!_msalClientInstance) {
    const config = {
      auth: {
        clientId: process.env.MICROSOFT_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}`,
        clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
      }
    };
    _msalClientInstance = new ConfidentialClientApplication(config);
    console.log('✅ MSAL client initialized (singleton)');
  }
  return _msalClientInstance;
}

// Get authorization URL for user login
async function getAuthUrl() {
  const scopes = [
    'Mail.ReadWrite',
    'Mail.Send',
    'Calendars.ReadWrite',
    'Files.ReadWrite',
    'Sites.Read.All',
    'User.Read',
    'Contacts.Read',
    'OnlineMeetings.ReadWrite',
    'Chat.ReadWrite',
    'offline_access'
  ];

  let redirectUri;
  if (process.env.REDIRECT_URI) {
    redirectUri = process.env.REDIRECT_URI;
  } else if (process.env.NODE_ENV === 'production') {
    redirectUri = 'https://ashistanto-bhc0fpeugkd9fqft.canadacentral-01.azurewebsites.net/auth/callback';
  } else {
    redirectUri = `http://localhost:${process.env.PORT || 3000}/auth/callback`;
  }

  console.log('🔐 Auth URL redirect_uri:', redirectUri);
  console.log('🔐 NODE_ENV:', process.env.NODE_ENV);

  const params = new URLSearchParams({
    client_id: process.env.MICROSOFT_CLIENT_ID,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: 'query',
    scope: scopes.join(' '),
    prompt: 'select_account'  // Force account selection screen
  });

  return `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize?${params.toString()}`;
}

// Get access token using authorization code (delegated flow)
// Uses singleton MSAL client so the refresh token is cached internally
async function getAccessTokenByAuthCode(code) {
  try {
    const msalClient = getMsalClient();

    let redirectUri;
    if (process.env.REDIRECT_URI) {
      redirectUri = process.env.REDIRECT_URI;
    } else if (process.env.NODE_ENV === 'production') {
      redirectUri = 'https://microsoft-agent-aubbhefsbzagdhha.eastus-01.azurewebsites.net/auth/callback';
    } else {
      redirectUri = `http://localhost:${process.env.PORT || 3000}/auth/callback`;
    }

    console.log('🔐 Token exchange redirect_uri:', redirectUri);

    const tokenRequest = {
      code: code,
      scopes: [
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.ReadWrite',
        'Files.ReadWrite',
        'Sites.Read.All',
        'User.Read',
        'Contacts.Read',
        'OnlineMeetings.ReadWrite',
        'Chat.ReadWrite',
        'offline_access'  // ✅ CRITICAL: Required to get refresh token
      ],
      redirectUri: redirectUri,
      codeVerifier: undefined
    };

    const response = await msalClient.acquireTokenByCode(tokenRequest);
    console.log('✅ Token acquired via auth code. Account:', response.account?.username);
    console.log('🔑 MSAL token cache now contains refresh token for silent renewal');
    return response;
  } catch (error) {
    console.error('Error getting access token by auth code:', error);
    throw new Error('Failed to exchange code for token');
  }
}

// Silently refresh access token using MSAL's internal token cache
// MSAL stores the refresh token internally after acquireTokenByCode
// acquireTokenSilent uses that cached refresh token automatically
async function refreshTokenSilently(account) {
  try {
    const msalClient = getMsalClient();
    const tokenRequest = {
      account: account,
      scopes: [
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.ReadWrite',
        'Files.ReadWrite',
        'Sites.Read.All',
        'User.Read',
        'Contacts.Read',
        'OnlineMeetings.ReadWrite',
        'Chat.ReadWrite',
        'offline_access'
      ],
      forceRefresh: false  // Only refresh if needed
    };

    console.log('🔄 Attempting silent token refresh for:', account?.username);
    const response = await msalClient.acquireTokenSilent(tokenRequest);
    console.log('✅ Silent token refresh succeeded');
    return response;
  } catch (error) {
    console.error('❌ Silent token refresh failed:', error.message);
    throw new Error('Token refresh failed. Please log in again.');
  }
}

// Legacy function kept for compatibility but uses silent refresh internally
async function getAccessTokenByRefreshToken(refreshToken, account) {
  if (account) {
    return refreshTokenSilently(account);
  }
  // Fallback: try using acquireTokenByRefreshToken directly
  try {
    const msalClient = getMsalClient();
    const tokenRequest = {
      refreshToken: refreshToken,
      scopes: [
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.ReadWrite',
        'Files.ReadWrite',
        'Sites.Read.All',
        'User.Read',
        'Contacts.Read',
        'OnlineMeetings.ReadWrite',
        'Chat.ReadWrite'
      ]
    };

    const response = await msalClient.acquireTokenByRefreshToken(tokenRequest);
    return response;
  } catch (error) {
    console.error('Error refreshing access token:', error);
    throw new Error('Failed to refresh access token');
  }
}

// Get access token using client credentials flow (fallback)
async function getAccessTokenAppOnly() {
  try {
    const msalClient = getMsalClient();
    const tokenRequest = {
      scopes: ['https://graph.microsoft.com/.default']
    };

    const response = await msalClient.acquireTokenByClientCredential(tokenRequest);
    return response.accessToken;
  } catch (error) {
    console.error('Error getting app-only access token:', error);
    throw new Error('Failed to authenticate with Microsoft Graph');
  }
}

// Initialize Graph client with automatic token refresh via MSAL silent flow
async function getGraphClient(userAccessToken = null, sessionId = null) {
  let accessToken;

  // If sessionId is provided, use MSAL's silent token refresh
  if (sessionId) {
    const { userTokenStore } = require('./auth');
    const tokenData = userTokenStore.get(sessionId);

    if (!tokenData) {
      throw new Error('Session not found. Please log in again.');
    }

    // Check if token needs refresh (expires in less than 5 minutes)
    const timeUntilExpiry = (tokenData.expiresAt || 0) - Date.now();
    const REFRESH_THRESHOLD = 5 * 60 * 1000; // 5 minutes

    if (timeUntilExpiry < REFRESH_THRESHOLD) {
      try {
        // Use MSAL's silent refresh with stored account object
        if (tokenData.account) {
          const newTokenResponse = await refreshTokenSilently(tokenData.account);
          const updatedTokenData = {
            ...tokenData,
            accessToken: newTokenResponse.accessToken,
            account: newTokenResponse.account || tokenData.account,
            expiresAt: Date.now() + ((newTokenResponse.expiresIn || 3600) * 1000),
          };
          userTokenStore.set(sessionId, updatedTokenData);
          accessToken = updatedTokenData.accessToken;
          console.log(`✅ Token refreshed silently for session: ${sessionId}`);
        } else {
          // No account object stored — can't silently refresh, use expired token as last resort
          console.warn(`⚠️ No MSAL account stored for session ${sessionId}, token may be expired`);
          accessToken = tokenData.accessToken;
        }
      } catch (error) {
        console.error(`❌ Silent token refresh failed for session ${sessionId}:`, error.message);
        // Don't delete session — let the user see a proper error
        accessToken = tokenData.accessToken;
      }
    } else {
      accessToken = tokenData.accessToken;
      console.log(`✓ Using valid token for session: ${sessionId}`);
    }
  } else if (userAccessToken) {
    accessToken = userAccessToken;
  } else if (process.env.MICROSOFT_ACCESS_TOKEN) {
    accessToken = process.env.MICROSOFT_ACCESS_TOKEN;
  } else {
    try {
      accessToken = await getAccessTokenAppOnly();
    } catch (error) {
      console.error('Failed to get any access token:', error);
      throw new Error('No valid access token available. Please login first.');
    }
  }

  return Client.init({
    authProvider: (done) => {
      done(null, accessToken);
    }
  });
}

// Get sender's profile information
async function getSenderProfile(userToken = null, sessionId = null) {
  try {
    console.log('🔍 [getSenderProfile] Starting profile fetch...');
    console.log('🔍 [getSenderProfile] Input params:', {
      hasUserToken: !!userToken,
      sessionId: sessionId,
      userTokenType: typeof userToken
    });

    const client = await getGraphClient(userToken, sessionId);
    console.log('✅ [getSenderProfile] Graph client initialized successfully');

    console.log('🔍 [getSenderProfile] Calling /me endpoint...');
    const user = await client.api('/me').get();

    console.log('✅ [getSenderProfile] Profile fetched successfully:', {
      displayName: user.displayName,
      email: user.mail || user.userPrincipalName,
      hasJobTitle: !!user.jobTitle,
      hasDepartment: !!user.department
    });

    return {
      displayName: user.displayName,
      email: user.mail || user.userPrincipalName,
      jobTitle: user.jobTitle || '',
      department: user.department || '',
      officeLocation: user.officeLocation || ''
    };
  } catch (error) {
    console.error('❌ [getSenderProfile] ERROR getting sender profile:');
    console.error('❌ [getSenderProfile] Error type:', error.constructor.name);
    console.error('❌ [getSenderProfile] Error message:', error.message);
    console.error('❌ [getSenderProfile] Error stack:', error.stack);

    // Log detailed error information
    if (error.statusCode) {
      console.error('❌ [getSenderProfile] HTTP Status Code:', error.statusCode);
    }
    if (error.code) {
      console.error('❌ [getSenderProfile] Error Code:', error.code);
    }
    if (error.body) {
      console.error('❌ [getSenderProfile] Error Body:', JSON.stringify(error.body, null, 2));
    }

    // Check token store status
    if (sessionId) {
      const { userTokenStore } = require('./auth');
      const hasSession = userTokenStore.has(sessionId);
      console.error('❌ [getSenderProfile] Session exists in token store:', hasSession);
      if (hasSession) {
        const tokenData = userTokenStore.get(sessionId);
        console.error('❌ [getSenderProfile] Token data available:', {
          hasAccessToken: !!tokenData?.accessToken,
          hasAccount: !!tokenData?.account,
          tokenLength: tokenData?.accessToken?.length || 0
        });
      }
    }

    console.error('⚠️ [getSenderProfile] Falling back to default "User" displayName');

    return {
      displayName: 'User',
      email: userEmail || 'sender@hoshodigital.com',
      jobTitle: '',
      department: '',
      officeLocation: ''
    };
  }
}

/**
 * 🔍 Search for contact email by name from Graph API
 * Enhanced with comprehensive diagnostic logging
 */
async function searchContactEmail(name, userToken = null, sessionId = null) {
  try {
    console.log(`\n${'='.repeat(60)}`);
    console.log(`🔍 CONTACT SEARCH DIAGNOSTIC`);
    console.log(`${'='.repeat(60)}`);
    console.log(`📝 Search term: "${name}"`);
    console.log(`📝 Trimmed: "${name.trim()}"`);

    const client = await getGraphClient(userToken, sessionId);
    const searchedName = name.trim();
    const searchResults = {
      personalContacts: { attempted: false, found: 0, error: null },
      peopleApi: { attempted: false, found: 0, error: null },
      orgDirectory: { attempted: false, found: 0, error: null }
    };

    // Step 1: Search in user's personal contacts
    try {
      searchResults.personalContacts.attempted = true;
      console.log(`\n📇 STEP 1: Personal Contacts`);
      console.log(`  → Filter: startswith(displayName,'${searchedName}') or startswith(givenName,'${searchedName}') or startswith(surname,'${searchedName}')`);

      const contacts = await client
        .api('/me/contacts')
        .filter(`startswith(displayName,'${searchedName}') or startswith(givenName,'${searchedName}') or startswith(surname,'${searchedName}')`)
        .select('displayName,emailAddresses,givenName,surname')
        .top(10)
        .get();

      console.log(`  → Raw results: ${contacts.value?.length || 0} contacts`);
      searchResults.personalContacts.found = contacts.value?.length || 0;

      if (contacts.value && contacts.value.length > 0) {
        contacts.value.forEach((contact, idx) => {
          console.log(`  ${idx + 1}. ${contact.displayName} (${contact.givenName} ${contact.surname})`);
          console.log(`     Emails: ${contact.emailAddresses?.length || 0}`);
        });

        const results = contacts.value
          .filter(contact => contact.emailAddresses && contact.emailAddresses.length > 0)
          .map(contact => ({
            name: contact.displayName,
            email: contact.emailAddresses[0].address,
            source: 'personal_contacts'
          }));

        if (results.length > 0) {
          console.log(`  ✅ SUCCESS: Found ${results.length} valid contact(s)`);
          console.log(`  📧 Selected: ${results[0].name} <${results[0].email}>`);
          console.log(`${'='.repeat(60)}\n`);
          return {
            found: true,
            results: results,
            searchedName: searchedName
          };
        }
      }
      console.log(`  ⚠️ No valid contacts with email addresses`);
    } catch (err) {
      searchResults.personalContacts.error = err.message;
      console.log(`  ❌ ERROR: ${err.message}`);
      if (err.statusCode) console.log(`  📊 Status Code: ${err.statusCode}`);
      if (err.code) console.log(`  📊 Error Code: ${err.code}`);
    }

    // Step 2: Search in People API
    try {
      searchResults.peopleApi.attempted = true;
      console.log(`\n👥 STEP 2: People API`);
      console.log(`  → Search query: "${searchedName}"`);

      const people = await client
        .api('/me/people')
        .search(`"${searchedName}"`)
        .select('displayName,emailAddresses,givenName,surname')
        .top(10)
        .get();

      console.log(`  → Raw results: ${people.value?.length || 0} people`);
      searchResults.peopleApi.found = people.value?.length || 0;

      if (people.value && people.value.length > 0) {
        people.value.forEach((person, idx) => {
          console.log(`  ${idx + 1}. ${person.displayName} (${person.givenName} ${person.surname})`);
          console.log(`     Emails: ${person.emailAddresses?.length || 0}`);
        });

        const results = people.value
          .filter(person => person.emailAddresses && person.emailAddresses.length > 0)
          .map(person => ({
            name: person.displayName,
            email: person.emailAddresses[0].address,
            source: 'people_api'
          }));

        if (results.length > 0) {
          console.log(`  ✅ SUCCESS: Found ${results.length} valid person(s)`);
          console.log(`  📧 Selected: ${results[0].name} <${results[0].email}>`);
          console.log(`${'='.repeat(60)}\n`);
          return {
            found: true,
            results: results,
            searchedName: searchedName
          };
        }
      }
      console.log(`  ⚠️ No valid people with email addresses`);
    } catch (err) {
      searchResults.peopleApi.error = err.message;
      console.log(`  ❌ ERROR: ${err.message}`);
      if (err.statusCode) console.log(`  📊 Status Code: ${err.statusCode}`);
      if (err.code) console.log(`  📊 Error Code: ${err.code}`);
    }

    // Step 3: Search in organization directory (with case-insensitive fallback)
    try {
      searchResults.orgDirectory.attempted = true;
      console.log(`\n🏢 STEP 3: Organization Directory`);
      console.log(`  → Filter: startswith(displayName,'${searchedName}') or startswith(givenName,'${searchedName}') or startswith(surname,'${searchedName}')`);

      const users = await client
        .api('/users')
        .filter(`startswith(displayName,'${searchedName}') or startswith(givenName,'${searchedName}') or startswith(surname,'${searchedName}')`)
        .select('displayName,mail,userPrincipalName,givenName,surname,id')
        .top(10)
        .get();

      console.log(`  → Raw results: ${users.value?.length || 0} users`);
      searchResults.orgDirectory.found = users.value?.length || 0;

      if (users.value && users.value.length > 0) {
        users.value.forEach((user, idx) => {
          console.log(`  ${idx + 1}. ${user.displayName} (${user.givenName} ${user.surname})`);
          console.log(`     Mail: ${user.mail || 'N/A'}`);
          console.log(`     UPN: ${user.userPrincipalName || 'N/A'}`);
          console.log(`     ID: ${user.id}`);
        });

        const results = users.value.map(user => ({
          name: user.displayName,
          email: user.mail || user.userPrincipalName,
          source: 'organization_directory'
        }));

        console.log(`  ✅ SUCCESS: Found ${results.length} user(s)`);
        console.log(`  📧 Selected: ${results[0].name} <${results[0].email}>`);
        console.log(`${'='.repeat(60)}\n`);
        return {
          found: true,
          results: results,
          searchedName: searchedName
        };
      }
      console.log(`  ⚠️ No users found with exact startswith match`);

      // FALLBACK: Try case-insensitive contains search
      console.log(`\n  🔄 FALLBACK: Trying case-insensitive contains search...`);
      const allUsers = await client
        .api('/users')
        .select('displayName,mail,userPrincipalName,givenName,surname,id')
        .top(100)
        .get();

      console.log(`  → Retrieved ${allUsers.value?.length || 0} users for client-side filtering`);

      const searchLower = searchedName.toLowerCase();
      const matchedUsers = allUsers.value?.filter(user => {
        const displayName = (user.displayName || '').toLowerCase();
        const givenName = (user.givenName || '').toLowerCase();
        const surname = (user.surname || '').toLowerCase();
        return displayName.includes(searchLower) ||
          givenName.includes(searchLower) ||
          surname.includes(searchLower);
      }) || [];

      console.log(`  → Matched ${matchedUsers.length} users with contains filter`);

      if (matchedUsers.length > 0) {
        matchedUsers.forEach((user, idx) => {
          console.log(`  ${idx + 1}. ${user.displayName} (${user.givenName} ${user.surname})`);
          console.log(`     Mail: ${user.mail || 'N/A'}`);
        });

        const results = matchedUsers.map(user => ({
          name: user.displayName,
          email: user.mail || user.userPrincipalName,
          source: 'organization_directory_fallback'
        }));

        console.log(`  ✅ FALLBACK SUCCESS: Found ${results.length} user(s)`);
        console.log(`  📧 Selected: ${results[0].name} <${results[0].email}>`);
        console.log(`${'='.repeat(60)}\n`);
        return {
          found: true,
          results: results,
          searchedName: searchedName
        };
      }

    } catch (err) {
      searchResults.orgDirectory.error = err.message;
      console.log(`  ❌ ERROR: ${err.message}`);
      if (err.statusCode) console.log(`  📊 Status Code: ${err.statusCode}`);
      if (err.code) console.log(`  📊 Error Code: ${err.code}`);
    }

    // Final summary
    console.log(`\n❌ SEARCH FAILED - SUMMARY`);
    console.log(`${'='.repeat(60)}`);
    console.log(`📝 Search term: "${searchedName}"`);
    console.log(`\n📊 Results by source:`);
    console.log(`  Personal Contacts: ${searchResults.personalContacts.attempted ? `${searchResults.personalContacts.found} found` : 'Not attempted'}`);
    if (searchResults.personalContacts.error) console.log(`    Error: ${searchResults.personalContacts.error}`);
    console.log(`  People API: ${searchResults.peopleApi.attempted ? `${searchResults.peopleApi.found} found` : 'Not attempted'}`);
    if (searchResults.peopleApi.error) console.log(`    Error: ${searchResults.peopleApi.error}`);
    console.log(`  Org Directory: ${searchResults.orgDirectory.attempted ? `${searchResults.orgDirectory.found} found` : 'Not attempted'}`);
    if (searchResults.orgDirectory.error) console.log(`    Error: ${searchResults.orgDirectory.error}`);

    console.log(`\n💡 Troubleshooting suggestions:`);
    console.log(`  1. Verify the exact spelling of the name`);
    console.log(`  2. Try using the full name (e.g., "Jatin Kumar")`);
    console.log(`  3. Check if the user exists in your organization`);
    console.log(`  4. Verify app permissions: User.Read.All, People.Read, Contacts.Read`);
    console.log(`  5. Try providing the email address directly`);
    console.log(`${'='.repeat(60)}\n`);

    return {
      found: false,
      searchedName: searchedName,
      message: `No user found with name "${searchedName}". Please verify the spelling or provide the email address directly.`,
      diagnostics: searchResults
    };

  } catch (error) {
    console.error(`\n❌ CRITICAL ERROR in searchContactEmail`);
    console.error(`${'='.repeat(60)}`);
    console.error(`Error: ${error.message}`);
    console.error(`Stack: ${error.stack}`);
    console.error(`${'='.repeat(60)}\n`);
    throw new Error(`Failed to search for contact "${name}": ${error.message}`);
  }
}

// ============== EMAIL FUNCTIONS ==============

async function getRecentEmails(count = 5, userToken = null, sessionId = null) {
  try {
    const client = await getGraphClient(userToken);

    const messages = await client
      .api('/me/messages')
      .select('id,subject,from,receivedDateTime,bodyPreview,isRead,hasAttachments,attachmentCount')
      .top(count)
      .orderby('receivedDateTime DESC')
      .get();

    // Get user timezone for formatting
    let userTimeZone = 'UTC';
    if (sessionId) {
      try {
        const tz = await timezoneHelper.getUserTimeZone(sessionId, userToken);
        if (tz) {
          userTimeZone = tz;
          console.log(`✓ Using user timezone: ${userTimeZone}`);
        }
      } catch (err) {
        console.warn('⚠️ Could not retrieve user timezone, using UTC:', err.message);
        userTimeZone = 'UTC';
      }
    } else {
      console.warn('⚠️ No sessionId provided, using UTC timezone');
    }

    const formattedEmails = messages.value.map(msg => ({
      id: msg.id,
      from: msg.from.emailAddress.name || msg.from.emailAddress.address,
      subject: msg.subject,
      receivedDate: msg.receivedDateTime,
      preview: msg.bodyPreview.substring(0, 100),
      hasAttachments: msg.hasAttachments,
      attachmentCount: msg.attachmentCount || 0
    }));

    // Always use formatter with timezone (never fallback)
    console.log(`📧 Formatting ${formattedEmails.length} emails with timezone: ${userTimeZone}`);
    return formatters.formatEmails(formattedEmails, userTimeZone);
  } catch (error) {
    console.error('Error getting emails:', error);
    throw new Error('Failed to retrieve emails');
  }
}

async function searchEmails(query, userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const messages = await client
      .api('/me/messages')
      .filter(`contains(subject,'${query}') or contains(from/emailAddress/address,'${query}')`)
      .select('subject,from,receivedDateTime,bodyPreview')
      .top(5)
      .get();

    return messages.value.map(msg => ({
      subject: msg.subject,
      from: msg.from.emailAddress.name || msg.from.emailAddress.address,
      date: new Date(msg.receivedDateTime).toLocaleString(),
      preview: msg.bodyPreview.substring(0, 100)
    }));
  } catch (error) {
    console.error('Error searching emails:', error);
    throw new Error('Failed to search emails');
  }
}

async function sendEmail(recipient_name, subject, body, ccRecipients = [], userToken = null, validatedRecipientData = null) {
  try {
    console.log(`📧 Sending email to: ${recipient_name}`);

    const senderProfile = await getSenderProfile(userToken);

    let recipient;

    // ✅ OPTIMIZATION: Use cached validated data if available (skip API call)
    if (validatedRecipientData && validatedRecipientData.recipientEmail) {
      console.log(`  ⚡ Using cached recipient data (fast path)`);
      recipient = {
        name: validatedRecipientData.recipientName,
        email: validatedRecipientData.recipientEmail,
        source: validatedRecipientData.source || 'cached'
      };
    } else {
      // Fallback: Search for recipient (slow path)
      console.log(`  🔍 Searching for recipient (slow path)`);
      const searchResult = await searchContactEmail(recipient_name, userToken);

      if (!searchResult.found) {
        console.log(`  ❌ Recipient not found: ${recipient_name}`);
        return {
          success: false,
          notFound: true,
          searchedName: searchResult.searchedName,
          message: searchResult.message
        };
      }

      recipient = searchResult.results[0];
    }

    console.log(`  ✅ Recipient: ${recipient.email} (source: ${recipient.source})`);

    const nameParts = recipient_name.trim().split(/\s+/);
    const firstName = nameParts[0];

    const ccEmailAddresses = [];
    if (ccRecipients && ccRecipients.length > 0) {
      console.log(`  📎 Processing ${ccRecipients.length} CC recipient(s)...`);
      for (const ccName of ccRecipients) {
        try {
          const ccResult = await searchContactEmail(ccName, userToken);
          if (ccResult.found && ccResult.results.length > 0) {
            ccEmailAddresses.push({
              emailAddress: {
                address: ccResult.results[0].email,
                name: ccResult.results[0].name
              }
            });
            console.log(`    ✅ CC: ${ccResult.results[0].email}`);
          } else {
            console.log(`    ⚠ CC recipient not found: ${ccName}`);
          }
        } catch (err) {
          console.log(`    ⚠ Could not find CC recipient: ${ccName}`);
        }
      }
    }

    let cleanBody = body
      .replace(/<[^>]*>/g, '')
      .replace(/^Hi\s+\w+,?\s*/gi, '')
      .replace(/^Dear\s+\w+,?\s*/gi, '')
      .replace(/Best\s+regards,?.*/gi, '')
      .replace(/Best\s+wishes,?.*/gi, '')
      .replace(/Regards,?.*/gi, '')
      .replace(/^--+.*$/gm, '')
      .replace(/\s+/g, ' ')
      .trim();

    const plainTextBody = `Hi ${firstName},

${cleanBody}

Best regards,
${senderProfile.displayName || 'User'}`;

    const client = await getGraphClient(userToken);
    const message = {
      message: {
        subject: subject,
        body: {
          contentType: 'Text',
          content: plainTextBody
        },
        toRecipients: [
          {
            emailAddress: {
              address: recipient.email,
              name: recipient.name
            }
          }
        ],
        ccRecipients: ccEmailAddresses
      },
      saveToSentItems: true
    };

    await client.api('/me/sendMail').post(message);

    const result = {
      success: true,
      message: `Email sent successfully to ${recipient.name}`,
      recipientEmail: recipient.email,
      recipientName: recipient.name,
      subject: subject,
      source: recipient.source
    };

    if (ccEmailAddresses.length > 0) {
      result.ccRecipients = ccEmailAddresses.map(cc => cc.emailAddress.address).join(', ');
    }

    console.log(`  ✅ Email sent successfully`);
    return result;

  } catch (error) {
    console.error('❌ Error sending email:', error);
    throw new Error(`Failed to send email: ${error.message}`);
  }
}

// ============== CALENDAR FUNCTIONS ==============

async function getCalendarEvents(days = 7, userToken = null, sessionId = null) {
  try {
    const client = await getGraphClient(userToken);
    const startDate = new Date();
    const endDate = new Date();
    endDate.setDate(endDate.getDate() + days);

    const events = await client
      .api('/me/calendar/events')
      .filter(`start/dateTime ge '${startDate.toISOString()}' and start/dateTime le '${endDate.toISOString()}'`)
      .select('id,subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeeting,bodyPreview')
      .orderby('start/dateTime')
      .top(50)
      .get();

    const eventsList = events.value.map(event => ({
      id: event.id,
      subject: event.subject,
      startDateTime: event.start.dateTime,
      endDateTime: event.end.dateTime,
      start: new Date(event.start.dateTime),
      end: new Date(event.end.dateTime),
      location: event.location?.displayName || 'No location',
      organizer: event.organizer?.emailAddress?.name || 'Unknown',
      attendees: event.attendees?.map(a => a.emailAddress.name || a.emailAddress.address) || [],
      attendeeCount: event.attendees?.length || 0,
      isTeamsMeeting: event.isOnlineMeeting || false,
      joinUrl: event.onlineMeeting?.joinUrl || null,
      bodyPreview: event.bodyPreview || ''
    }));

    // Get user timezone for formatting
    let userTimeZone = 'UTC';
    if (sessionId) {
      try {
        const tz = await timezoneHelper.getUserTimeZone(sessionId, userToken);
        if (tz) {
          userTimeZone = tz;
          console.log(`✓ Using user timezone: ${userTimeZone}`);
        }
      } catch (err) {
        console.warn('⚠️ Could not retrieve user timezone, using UTC:', err.message);
        userTimeZone = 'UTC';
      }
    } else {
      console.warn('⚠️ No sessionId provided, using UTC timezone');
    }

    // Always use formatter with timezone (never fallback to unformatted)
    console.log(`📋 Formatting ${eventsList.length} calendar events with timezone: ${userTimeZone}`);
    return formatters.formatCalendarEvents(eventsList, userTimeZone);
  } catch (error) {
    console.error('Error getting calendar events:', error);
    throw new Error('Failed to retrieve calendar events');
  }
}

/**
 * 📅 Creates a Calendar Event (Teams or Regular Meeting)
 *
 * 🟢 Automatically:
 *  - Resolves attendee names to emails
 *  - Creates Teams meeting (if requested)
 *  - Fetches join link
 *  - Sends join link to all attendees via Teams Chat
 *
 * @returns Result object including Teams join URL
 */
async function createCalendarEvent(
  subject,
  start,
  end,
  location = '',
  attendeeNames = [],
  isTeamsMeeting = false,
  userToken = null
) {
  try {
    console.log(`📅 Creating calendar event: "${subject}"`);
    console.log(`   Teams meeting: ${isTeamsMeeting}`);

    if (!userToken) throw new Error('Missing user token.');

    const client = await getGraphClient(userToken);

    //------------------------------------------------------
    // 🧑‍🤝‍🧑 Resolve attendees (get Outlook email addresses)
    //------------------------------------------------------
    //------------------------------------------------------
    // 🧑‍🤝‍🧑 Resolve attendees (get Outlook email addresses)
    //------------------------------------------------------
    const attendeeEmails = [];
    const notFoundAttendees = [];

    if (attendeeNames && attendeeNames.length > 0) {
      console.log(`   Processing ${attendeeNames.length} attendee(s)...`);
      for (const name of attendeeNames) {
        try {
          const searchResult = await searchContactEmail(name, userToken);
          if (searchResult.found && searchResult.results.length > 0) {
            attendeeEmails.push({
              emailAddress: {
                address: searchResult.results[0].email,
                name: searchResult.results[0].name
              },
              type: 'required'
            });
            console.log(`     ✅ Attendee: ${searchResult.results[0].email}`);
          } else {
            notFoundAttendees.push(name);
            console.log(`     ⚠ Attendee not found: ${name}`);
          }
        } catch (err) {
          notFoundAttendees.push(name);
          console.log(`     ⚠ Could not find attendee: ${name}`);
        }
      }
    }

    //------------------------------------------------------
    // ❗ Validate attendees before creating meeting
    //------------------------------------------------------
    if (attendeeNames.length > 0 && notFoundAttendees.length > 0) {
      console.log("❌ Cannot create meeting. Attendee(s) not found:", notFoundAttendees);

      return {
        success: false,
        notFound: true,
        message: `Cannot create meeting. I couldn't find: ${notFoundAttendees.join(', ')}. Please verify their name(s).`,
        missingAttendees: notFoundAttendees
      };
    }


    //------------------------------------------------------
    // 📝 Event payload for Graph API
    //------------------------------------------------------
    const event = {
      subject,
      start: { dateTime: start, timeZone: 'Asia/Kolkata' },
      end: { dateTime: end, timeZone: 'Asia/Kolkata' },
      attendees: attendeeEmails
    };

    if (location) {
      event.location = { displayName: location };
    }

    //------------------------------------------------------
    // 🎥 Teams meeting enabled?
    //------------------------------------------------------
    if (isTeamsMeeting) {
      event.isOnlineMeeting = true;
      event.onlineMeetingProvider = 'teamsForBusiness';
    }

    //------------------------------------------------------
    // 🚀 Create meeting
    //------------------------------------------------------
    console.log('   → Creating event with Graph API...');
    const createdEvent = await client.api('/me/events').post(event);

    //------------------------------------------------------
    // 🔁 Wait for Teams join link (sometimes takes 2–3s)
    //------------------------------------------------------
    if (isTeamsMeeting && !createdEvent.onlineMeeting?.joinUrl) {
      console.log('   → Waiting for Teams link generation...');
      await new Promise(resolve => setTimeout(resolve, 2000));

      try {
        const refreshedEvent = await client
          .api(`/me/events/${createdEvent.id}`)
          .select('id,subject,onlineMeeting')
          .get();

        if (refreshedEvent.onlineMeeting?.joinUrl) {
          createdEvent.onlineMeeting = refreshedEvent.onlineMeeting;
          console.log('   ✅ Teams link retrieved after refresh');
        }
      } catch (e) {
        console.log('   ⚠ Could not fetch Teams link after refresh');
      }
    }

    //------------------------------------------------------
    // 🎁 Final response with link
    //------------------------------------------------------
    const result = {
      success: true,
      eventId: createdEvent.id,
      subject: createdEvent.subject,
      attendees: attendeeEmails.map(a => a.emailAddress.name).join(', '),
      attendeeCount: attendeeEmails.length,
      startTime: new Date(start).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      endTime: new Date(end).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      isTeamsMeeting: isTeamsMeeting,
      joinUrl: createdEvent.onlineMeeting?.joinUrl || null
    };

    //------------------------------------------------------
    // 📤 AUTO-SEND JOIN LINK TO EVERY ATTENDEE IN TEAMS CHAT
    //------------------------------------------------------
    if (isTeamsMeeting && result.joinUrl && attendeeNames.length > 0) {
      console.log("📤 Sending Teams meeting link to attendees...");

      for (const attendee of attendeeNames) {
        try {
          await sendTeamsMessage(
            attendee,
            `You are invited to join the meeting:\n"${subject}"\n\n📅 Time: ${result.startTime}\n\n🔗 Join Link: ${result.joinUrl}`,
            userToken
          );
          console.log(`   🚀 Link sent to: ${attendee}`);
        } catch (err) {
          console.log(`   ⚠ Could not send link to: ${attendee}`);
        }
      }
    }

    //------------------------------------------------------
    // 🗨 Response summary
    //------------------------------------------------------
    if (isTeamsMeeting && result.joinUrl) {
      result.message = `Teams meeting created. Link shared with participants.`;
    } else if (isTeamsMeeting) {
      result.message = `Teams meeting created. Link will appear shortly.`;
    } else {
      result.message = `Calendar event created successfully.`;
    }

    return result;
  } catch (error) {
    console.error('❌ Error creating calendar event:', error);
    throw new Error('Failed to create calendar event: ' + error.message);
  }
}


async function updateCalendarEvent(
  eventId,
  newAttendeeNames = [],
  newSubject = null,
  newStart = null,
  newEnd = null,
  userToken = null
) {
  try {
    console.log(`📅 Updating calendar event: ${eventId}`);

    if (!userToken) throw new Error('Missing user token.');

    const client = await getGraphClient(userToken);

    const existingEvent = await client
      .api(`/me/events/${eventId}`)
      .select('subject,start,end,attendees,isOnlineMeeting,onlineMeeting')
      .get();

    console.log(`   → Current event: "${existingEvent.subject}"`);
    console.log(`   → Current attendees: ${existingEvent.attendees?.length || 0}`);

    const updateData = {};

    if (newSubject) {
      updateData.subject = newSubject;
    }

    if (newStart) {
      updateData.start = { dateTime: newStart, timeZone: 'Asia/Kolkata' };
    }
    if (newEnd) {
      updateData.end = { dateTime: newEnd, timeZone: 'Asia/Kolkata' };
    }

    if (newAttendeeNames && newAttendeeNames.length > 0) {
      console.log(`   → Adding ${newAttendeeNames.length} new attendee(s)...`);

      const existingAttendees = existingEvent.attendees || [];
      const newAttendees = [];

      for (const name of newAttendeeNames) {
        try {
          const searchResult = await searchContactEmail(name, userToken);
          if (searchResult.found && searchResult.results.length > 0) {
            const email = searchResult.results[0].email;

            const alreadyExists = existingAttendees.some(a =>
              a.emailAddress.address.toLowerCase() === email.toLowerCase()
            );

            if (!alreadyExists) {
              newAttendees.push({
                emailAddress: {
                  address: email,
                  name: searchResult.results[0].name
                },
                type: 'required'
              });
              console.log(`     ✅ Adding: ${email}`);
            } else {
              console.log(`     ⚠ Already attending: ${email}`);
            }
          } else {
            console.log(`     ⚠ Not found: ${name}`);
          }
        } catch (err) {
          console.log(`     ⚠ Error finding: ${name}`);
        }
      }

      updateData.attendees = [...existingAttendees, ...newAttendees];
      console.log(`   → Total attendees after update: ${updateData.attendees.length}`);
    }

    const updatedEvent = await client
      .api(`/me/events/${eventId}`)
      .patch(updateData);

    console.log('   ✅ Event updated successfully');

    return {
      success: true,
      eventId: updatedEvent.id,
      subject: updatedEvent.subject,
      attendeeCount: updatedEvent.attendees?.length || 0,
      attendees: updatedEvent.attendees?.map(a => a.emailAddress.name || a.emailAddress.address).join(', '),
      message: `Event "${updatedEvent.subject}" updated successfully. Total attendees: ${updatedEvent.attendees?.length || 0}`,
      isTeamsMeeting: existingEvent.isOnlineMeeting,
      joinUrl: existingEvent.onlineMeeting?.joinUrl || null
    };

  } catch (error) {
    console.error('❌ Error updating calendar event:', error);
    throw new Error('Failed to update calendar event: ' + error.message);
  }
}

async function deleteCalendarEvents(subject = null, attendeeName = null, date = null, userToken = null) {
  try {
    console.log(`🗑️ Searching for calendar event(s) to delete...`);

    if (!userToken) throw new Error('Missing user token.');

    const client = await getGraphClient(userToken);

    let startDate, endDate;
    if (date) {
      if (date.toLowerCase() === 'today') {
        startDate = new Date();
        startDate.setHours(0, 0, 0, 0);
        endDate = new Date();
        endDate.setHours(23, 59, 59, 999);
      } else if (date.toLowerCase() === 'tomorrow') {
        startDate = new Date();
        startDate.setDate(startDate.getDate() + 1);
        startDate.setHours(0, 0, 0, 0);
        endDate = new Date(startDate);
        endDate.setHours(23, 59, 59, 999);
      } else {
        startDate = new Date(date);
        startDate.setHours(0, 0, 0, 0);
        endDate = new Date(date);
        endDate.setHours(23, 59, 59, 999);
      }
      console.log(`   → Date filter: ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}`);
    } else {
      startDate = new Date();
      endDate = new Date();
      endDate.setDate(endDate.getDate() + 30);
    }

    let filterQuery = `start/dateTime ge '${startDate.toISOString()}' and start/dateTime le '${endDate.toISOString()}'`;

    const events = await client
      .api('/me/calendar/events')
      .filter(filterQuery)
      .select('id,subject,start,attendees')
      .top(100)
      .orderby('start/dateTime')
      .get();

    if (!events.value || events.value.length === 0) {
      return {
        success: false,
        notFound: true,
        message: 'No events found in the specified time range'
      };
    }

    console.log(`   📅 Found ${events.value.length} events in date range`);

    let matchingEvents = events.value;
    if (subject) {
      const subjectLower = subject.toLowerCase();
      matchingEvents = matchingEvents.filter(e =>
        e.subject && e.subject.toLowerCase().includes(subjectLower)
      );
      console.log(`   🔍 After subject filter ("${subject}"): ${matchingEvents.length} matches`);
    }

    if (attendeeName) {
      const attendeeLower = attendeeName.toLowerCase();
      matchingEvents = matchingEvents.filter(e =>
        e.attendees && e.attendees.some(a =>
          a.emailAddress.name?.toLowerCase().includes(attendeeLower) ||
          a.emailAddress.address?.toLowerCase().includes(attendeeLower)
        )
      );
      console.log(`   🔍 After attendee filter ("${attendeeName}"): ${matchingEvents.length} matches`);
    }

    if (matchingEvents.length === 0) {
      const criteria = [];
      if (subject) criteria.push(`subject containing "${subject}"`);
      if (attendeeName) criteria.push(`attendee "${attendeeName}"`);
      if (date) criteria.push(`on ${date}`);

      return {
        success: false,
        notFound: true,
        message: `No events found with ${criteria.join(' and ')}`
      };
    }

    console.log(`   🗑️ Deleting ${matchingEvents.length} event(s)...`);
    const deletedEvents = [];

    for (const event of matchingEvents) {
      try {
        await client.api(`/me/events/${event.id}`).delete();
        deletedEvents.push({
          subject: event.subject,
          start: new Date(event.start.dateTime).toLocaleString()
        });
        console.log(`     ✅ Deleted: "${event.subject}"`);
      } catch (deleteError) {
        console.log(`     ❌ Failed to delete: "${event.subject}"`);
      }
    }

    if (deletedEvents.length === 0) {
      return {
        success: false,
        message: 'Failed to delete any events'
      };
    }

    return {
      success: true,
      deletedCount: deletedEvents.length,
      deletedEvents: deletedEvents,
      message: `Successfully deleted ${deletedEvents.length} event(s)`
    };

  } catch (error) {
    console.error('❌ Error deleting calendar events:', error);
    throw new Error('Failed to delete calendar events: ' + error.message);
  }
}

// ============== TEAMS FUNCTIONS ==============

async function getTeams(userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const teams = await client
      .api('/me/joinedTeams')
      .get();

    return teams.value.map(team => ({
      name: team.displayName,
      description: team.description || 'No description',
      id: team.id
    }));
  } catch (error) {
    console.error('Error getting teams:', error);
    throw new Error('Failed to retrieve teams');
  }
}

async function getTeamChannels(teamId, userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const channels = await client
      .api(`/teams/${teamId}/channels`)
      .get();

    return channels.value.map(channel => ({
      name: channel.displayName,
      description: channel.description || 'No description',
      id: channel.id
    }));
  } catch (error) {
    console.error('Error getting team channels:', error);
    throw new Error('Failed to retrieve team channels');
  }
}

async function sendTeamsMessage(recipientName, message, userToken = null, validatedRecipientData = null) {
  try {
    console.log(`💬 Sending Teams message to: ${recipientName}`);

    if (!userToken) throw new Error('User token required for Teams messaging');

    const client = await getGraphClient(userToken);

    let recipientEmail;
    let recipientDisplayName;

    // ✅ OPTIMIZATION: Use cached validated data if available (skip API call)
    if (validatedRecipientData && validatedRecipientData.recipientEmail) {
      console.log(`  ⚡ Using cached recipient data (fast path)`);
      recipientEmail = validatedRecipientData.recipientEmail;
      recipientDisplayName = validatedRecipientData.recipientName;
    } else {
      // Fallback: Search for recipient (slow path)
      console.log(`  🔍 Searching for recipient (slow path)`);
      const searchResult = await searchContactEmail(recipientName, userToken);

      if (!searchResult.found) {
        console.log(`   ❌ Recipient not found: ${recipientName}`);
        return {
          success: false,
          notFound: true,
          searchedName: searchResult.searchedName,
          message: searchResult.message
        };
      }

      recipientEmail = searchResult.results[0].email;
      recipientDisplayName = searchResult.results[0].name;
    }

    console.log(`   ✅ Recipient email: ${recipientEmail}`);

    const users = await client
      .api('/users')
      .filter(`mail eq '${recipientEmail}' or userPrincipalName eq '${recipientEmail}'`)
      .select('id,displayName,mail,userPrincipalName')
      .get();

    if (!users.value || users.value.length === 0) {
      throw new Error(`Could not find user ID for: ${recipientEmail}`);
    }

    const recipientUserId = users.value[0].id;
    console.log(`   ✅ Found recipient ID: ${recipientUserId}`);

    console.log('   → Creating/finding chat...');
    const chatBody = {
      chatType: 'oneOnOne',
      members: [
        {
          '@odata.type': '#microsoft.graph.aadUserConversationMember',
          roles: ['owner'],
          'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${recipientUserId}')`
        }
      ]
    };

    let chatId;
    try {
      const chat = await client.api('/chats').post(chatBody);
      chatId = chat.id;
      console.log(`   ✅ Chat created/found: ${chatId}`);
    } catch (err) {
      console.log('   → Chat may already exist, searching...');
      const chats = await client
        .api('/me/chats')
        .filter(`chatType eq 'oneOnOne'`)
        .expand('members')
        .get();

      const existingChat = chats.value.find(chat =>
        chat.members.some(member => member.userId === recipientUserId)
      );

      if (existingChat) {
        chatId = existingChat.id;
        console.log(`   ✅ Found existing chat: ${chatId}`);
      } else {
        throw new Error('Could not create or find chat');
      }
    }

    console.log('   → Sending message...');
    const messageBody = {
      body: {
        contentType: 'text',
        content: message
      }
    };

    const sentMessage = await client
      .api(`/chats/${chatId}/messages`)
      .post(messageBody);

    console.log('   ✅ Message sent successfully');

    return {
      success: true,
      message: `Teams message sent to ${recipientDisplayName}`,
      recipientName: recipientDisplayName,
      recipientEmail: recipientEmail,
      chatId: chatId,
      messageId: sentMessage.id
    };

  } catch (error) {
    console.error('❌ Error sending Teams message:', error);
    throw new Error(`Failed to send Teams message: ${error.message}`);
  }
}

async function getTeamsMessages(chatId = null, count = 10, userToken = null) {
  try {
    console.log(`📋 Getting recent Teams messages...`);

    if (!userToken) {
      throw new Error('User token required for Teams operations');
    }

    const client = await getGraphClient(userToken);

    if (!chatId) {
      console.log('   → No chatId provided, finding most recent chat...');
      const chats = await client
        .api('/me/chats')
        .top(5)
        .orderby('lastMessagePreview/createdDateTime DESC')
        .get();

      if (!chats.value || chats.value.length === 0) {
        return {
          success: true,
          messages: [],
          message: 'No chats found'
        };
      }

      chatId = chats.value[0].id;
      console.log(`   ✅ Using most recent chat: ${chatId}`);
    }

    console.log(`   → Fetching ${count} messages from chat...`);
    const messages = await client
      .api(`/chats/${chatId}/messages`)
      .top(count)
      .orderby('createdDateTime DESC')
      .get();

    console.log(`   ✅ Retrieved ${messages.value.length} messages`);

    return {
      success: true,
      chatId: chatId,
      messages: messages.value.map(msg => ({
        id: msg.id,
        content: msg.body?.content?.substring(0, 100) || '',
        sender: msg.from?.user?.displayName || 'Unknown',
        senderId: msg.from?.user?.id || null,
        sentDate: new Date(msg.createdDateTime).toLocaleString(),
        createdDateTime: msg.createdDateTime,
        deletedDateTime: msg.deletedDateTime || null,
        isDeleted: !!msg.deletedDateTime
      }))
    };

  } catch (error) {
    console.error('❌ Error getting Teams messages:', error);
    throw new Error('Failed to retrieve Teams messages: ' + error.message);
  }
}

async function deleteTeamsMessage(chatId = null, messageId = null, messageContent = null, userToken = null, previewMode = false) {
  try {
    console.log(`🗑️ Attempting to delete Teams message...`);

    if (!userToken) {
      throw new Error('User token required for Teams operations');
    }

    const client = await getGraphClient(userToken);

    const me = await client.api('/me').select('id,displayName').get();
    const currentUserId = me.id;
    console.log(`   → Current user: ${me.displayName} (${currentUserId})`);

    if (!chatId || !messageId) {
      console.log('   → Searching for message to delete...');

      const chats = await client
        .api('/me/chats')
        .top(10)
        .orderby('lastMessagePreview/createdDateTime DESC')
        .get();

      if (!chats.value || chats.value.length === 0) {
        return {
          success: false,
          notFound: true,
          message: 'No chats found'
        };
      }

      for (const chat of chats.value) {
        try {
          const messages = await client
            .api(`/chats/${chat.id}/messages`)
            .top(20)
            .orderby('createdDateTime DESC')
            .get();

          let targetMessage;
          if (messageContent) {
            targetMessage = messages.value.find(msg =>
              msg.from?.user?.id === currentUserId &&
              !msg.deletedDateTime &&
              msg.body?.content?.toLowerCase().includes(messageContent.toLowerCase())
            );
          } else {
            targetMessage = messages.value.find(msg =>
              msg.from?.user?.id === currentUserId &&
              !msg.deletedDateTime
            );
          }

          if (targetMessage) {
            chatId = chat.id;
            messageId = targetMessage.id;
            console.log(`   ✅ Found message: "${targetMessage.body?.content?.substring(0, 50)}..."`);

            // ✅ PREVIEW MODE: Return message details without deleting
            if (previewMode) {
              console.log('   👁️ Preview mode - not deleting yet');
              return {
                success: true,
                messageToDelete: {
                  chatId: chatId,
                  messageId: messageId,
                  content: targetMessage.body?.content || 'Message',
                  sentDate: new Date(targetMessage.createdDateTime).toLocaleString()
                }
              };
            }

            break;
          }
        } catch (chatError) {
          console.log(`   ⚠ Error searching chat:`, chatError.message);
          continue;
        }
      }

      if (!chatId || !messageId) {
        return {
          success: false,
          notFound: true,
          message: messageContent
            ? `No message found containing "${messageContent}"`
            : 'No recent message found to delete'
        };
      }
    }

    console.log(`   → Attempting to delete message...`);

    try {
      await client
        .api(`/chats/${chatId}/messages/${messageId}/softDelete`)
        .post({});

      console.log('   ✅ Teams message deleted successfully');

      return {
        success: true,
        message: 'Teams message deleted successfully',
        deletedMessageId: messageId,
        chatId: chatId
      };

    } catch (deleteError) {
      console.error('   ❌ Soft delete failed:', deleteError);

      if (deleteError.statusCode === 404) {
        return {
          success: false,
          notFound: true,
          message: 'Message not found or already deleted'
        };
      }

      if (deleteError.statusCode === 403 || deleteError.code === 'Forbidden') {
        return {
          success: false,
          message: 'Cannot delete this message. You can only delete messages you sent recently.'
        };
      }

      try {
        console.log('   → Trying alternative: editing message content...');
        await client
          .api(`/chats/${chatId}/messages/${messageId}`)
          .patch({
            body: {
              contentType: 'text',
              content: '[Message deleted]'
            }
          });

        console.log('   ✅ Message content cleared');
        return {
          success: true,
          message: 'Message content cleared (deletion not fully supported)',
          deletedMessageId: messageId,
          chatId: chatId
        };
      } catch (patchError) {
        console.error('   ❌ Message edit also failed:', patchError);
        throw deleteError;
      }
    }

  } catch (error) {
    console.error('❌ Error deleting Teams message:', error);
    throw new Error('Failed to delete Teams message: ' + error.message);
  }
}

// ============== USER FUNCTIONS ==============

async function getUserProfile(userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const user = await client.api('/me').get();

    return {
      name: user.displayName,
      email: user.mail || user.userPrincipalName,
      jobTitle: user.jobTitle || 'Not specified',
      officeLocation: user.officeLocation || 'Not specified',
      mobilePhone: user.mobilePhone || 'Not specified'
    };
  } catch (error) {
    console.error('Error getting user profile:', error);
    throw new Error('Failed to retrieve user profile');
  }
}

// ============== SHAREPOINT FUNCTIONS ==============

async function getRecentFiles(count = 10, userToken = null, sessionId = null) {
  try {
    const client = await getGraphClient(userToken);
    const files = await client
      .api('/me/drive/recent')
      .select('id,name,webUrl,lastModifiedDateTime,size,file,lastModifiedBy')
      .top(count)
      .get();

    const filesList = files.value.map(file => ({
      id: file.id,
      name: file.name,
      type: file.file?.mimeType || 'folder',
      modifiedDate: file.lastModifiedDateTime,
      size: file.size,
      modifiedBy: file.lastModifiedBy?.user?.displayName || 'Unknown',
      webUrl: file.webUrl
    }));

    // Get user timezone for formatting
    let userTimeZone = 'UTC';
    if (sessionId) {
      try {
        const tz = await timezoneHelper.getUserTimeZone(sessionId, userToken);
        if (tz) {
          userTimeZone = tz;
          console.log(`✓ Using user timezone: ${userTimeZone}`);
        }
      } catch (err) {
        console.warn('⚠️ Could not retrieve user timezone, using UTC:', err.message);
        userTimeZone = 'UTC';
      }
    } else {
      console.warn('⚠️ No sessionId provided, using UTC timezone');
    }

    // Always use formatter with timezone (never fallback)
    console.log(`📁 Formatting ${filesList.length} files with timezone: ${userTimeZone}`);
    return formatters.formatFiles(filesList, userTimeZone);
  } catch (error) {
    console.error('Error getting recent files:', error);
    throw new Error('Failed to retrieve recent files');
  }
}

async function searchFiles(query, userToken = null) {
  try {
    console.log(`🔍 Searching files for: "${query}"`);
    const client = await getGraphClient(userToken);

    // Request additional fields including parentReference for folder path
    const files = await client
      .api(`/me/drive/root/search(q='${query}')`)
      .select('id,name,webUrl,lastModifiedDateTime,size,file,parentReference,createdDateTime')
      .top(10)
      .get();

    console.log(`   ✅ Found ${files.value.length} files`);

    const formattedFiles = files.value.map(file => {
      // Extract folder path from parentReference
      let folderPath = '';
      let breadcrumb = '';

      if (file.parentReference && file.parentReference.path) {
        // Path format: /drive/root:/Folder1/Folder2
        const pathParts = file.parentReference.path.split(':');
        if (pathParts.length > 1 && pathParts[1]) {
          folderPath = pathParts[1];
          // Create breadcrumb: Folder1 → Folder2 → FileName
          const folders = folderPath.split('/').filter(f => f);
          folders.push(file.name);
          breadcrumb = folders.join(' → ');
        }
      }

      if (!breadcrumb) {
        breadcrumb = `Root → ${file.name}`;
      }

      return {
        name: file.name,
        webUrl: file.webUrl,
        lastModified: new Date(file.lastModifiedDateTime).toLocaleString(),
        size: formatFileSize(file.size),
        folderPath: folderPath || '/Root',
        breadcrumb: breadcrumb,
        type: file.file?.mimeType || 'folder',
        location: `You can find this file at: ${breadcrumb}`
      };
    });

    // Return formatted response with clear structure
    return {
      success: true,
      count: formattedFiles.length,
      query: query,
      files: formattedFiles,
      summary: `Found ${formattedFiles.length} file(s) matching "${query}"`
    };
  } catch (error) {
    console.error('Error searching files:', error);
    throw new Error('Failed to search files: ' + error.message);
  }
}

function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// ============== DELETION FUNCTIONS ==============

async function getRecentSentEmails(count = 10, userToken = null, sessionId = null) {
  try {
    console.log(`📬 Getting ${count} recent sent emails...`);
    const client = await getGraphClient(userToken);

    const messages = await client
      .api('/me/mailFolders/sentItems/messages')
      .select('id,subject,toRecipients,sentDateTime,bodyPreview,hasAttachments,attachmentCount')
      .top(count)
      .orderby('sentDateTime DESC')
      .get();

    const emailsList = messages.value.map(msg => ({
      id: msg.id,
      from: 'You',
      subject: msg.subject,
      receivedDate: msg.sentDateTime,
      preview: msg.bodyPreview?.substring(0, 100) || '',
      hasAttachments: msg.hasAttachments,
      attachmentCount: msg.attachmentCount || 0
    }));

    // Get user timezone for formatting
    let userTimeZone = 'UTC';
    if (sessionId) {
      try {
        const tz = await timezoneHelper.getUserTimeZone(sessionId, userToken);
        if (tz) {
          userTimeZone = tz;
          console.log(`✓ Using user timezone: ${userTimeZone}`);
        }
      } catch (err) {
        console.warn('⚠️ Could not retrieve user timezone, using UTC:', err.message);
        userTimeZone = 'UTC';
      }
    } else {
      console.warn('⚠️ No sessionId provided, using UTC timezone');
    }

    // Always use formatter with timezone (never fallback)
    console.log(`📬 Formatting ${emailsList.length} sent emails with timezone: ${userTimeZone}`);
    return {
      success: true,
      emails: formatters.formatEmails(emailsList, userTimeZone)
    };
  } catch (error) {
    console.error('Error getting sent emails:', error);
    throw new Error('Failed to retrieve sent emails: ' + error.message);
  }
}

async function deleteEmail(messageId, userToken = null) {
  try {
    console.log(`🗑️ Deleting email: ${messageId}`);
    const client = await getGraphClient(userToken);

    await client.api(`/me/messages/${messageId}`).delete();

    console.log('   ✅ Email deleted successfully');
    return {
      success: true,
      message: 'Email deleted successfully',
      deletedId: messageId
    };
  } catch (error) {
    console.error('Error deleting email:', error);
    throw new Error('Failed to delete email: ' + error.message);
  }
}

async function deleteSentEmail(subject = null, recipientEmail = null, userToken = null, previewMode = false) {
  try {
    console.log(`🗑️ Searching for sent email to delete...`);
    console.log(`   Subject filter: ${subject || 'none'}, Recipient filter: ${recipientEmail || 'none'}`);

    const client = await getGraphClient(userToken);

    console.log('   📥 Fetching recent sent emails...');

    const messages = await client
      .api('/me/mailFolders/sentItems/messages')
      .select('id,subject,toRecipients,sentDateTime')
      .top(50)
      .orderby('sentDateTime DESC')
      .get();

    if (!messages.value || messages.value.length === 0) {
      console.log('   ❌ No sent emails found');
      return {
        success: false,
        notFound: true,
        message: 'No sent emails found in your Sent Items folder'
      };
    }

    console.log(`   📧 Retrieved ${messages.value.length} sent emails`);

    let candidates = messages.value;

    if (subject) {
      const subjectLower = subject.toLowerCase();
      candidates = candidates.filter(msg => {
        const msgSubject = msg.subject || '';
        return msgSubject.toLowerCase().includes(subjectLower);
      });
      console.log(`   🔍 After subject filter ("${subject}"): ${candidates.length} matches`);
    }

    if (recipientEmail) {
      const recipientLower = recipientEmail.toLowerCase();
      candidates = candidates.filter(msg => {
        if (!msg.toRecipients || msg.toRecipients.length === 0) {
          return false;
        }
        return msg.toRecipients.some(r => {
          // Match by email address OR display name
          const address = r.emailAddress?.address || '';
          const name = r.emailAddress?.name || '';
          return address.toLowerCase().includes(recipientLower) ||
            name.toLowerCase().includes(recipientLower);
        });
      });
      console.log(`   🔍 After recipient filter ("${recipientEmail}"): ${candidates.length} matches`);
    }

    if (candidates.length === 0) {
      const criteria = [];
      if (subject) criteria.push(`subject containing "${subject}"`);
      if (recipientEmail) criteria.push(`recipient containing "${recipientEmail}"`);

      console.log(`   ❌ No emails matched the criteria`);
      return {
        success: false,
        notFound: true,
        message: `No sent email found with ${criteria.join(' and ')}`
      };
    }

    const emailToDelete = candidates[0];
    const recipientsList = emailToDelete.toRecipients?.map(r => r.emailAddress.address).join(', ') || 'unknown';

    console.log(`   🎯 Selected email to delete:`);
    console.log(`      Subject: "${emailToDelete.subject}"`);
    console.log(`      To: ${recipientsList}`);
    console.log(`      Sent: ${new Date(emailToDelete.sentDateTime).toLocaleString()}`);
    console.log(`      ID: ${emailToDelete.id}`);

    // ✅ PREVIEW MODE: Return email details without deleting
    if (previewMode) {
      console.log('   👁️ Preview mode - not deleting yet');
      return {
        success: true,
        emailToDelete: {
          id: emailToDelete.id,
          subject: emailToDelete.subject,
          recipient: recipientsList,
          sentDate: new Date(emailToDelete.sentDateTime).toLocaleString()
        }
      };
    }

    console.log(`   🗑️ Deleting email...`);

    await client.api(`/me/messages/${emailToDelete.id}`).delete();

    console.log('   ✅ Email deleted successfully!');

    return {
      success: true,
      message: `Successfully deleted email: "${emailToDelete.subject}"`,
      deletedSubject: emailToDelete.subject,
      deletedId: emailToDelete.id,
      deletedTo: recipientsList,
      sentDate: new Date(emailToDelete.sentDateTime).toLocaleString()
    };

  } catch (error) {
    console.error('❌ Error in deleteSentEmail:', error);

    if (error.statusCode) {
      throw new Error(`Graph API error (${error.code || error.statusCode}): ${error.message}`);
    }

    throw new Error('Failed to delete sent email: ' + error.message);
  }
}

// Get user profile photo
async function getUserProfilePhoto(userToken = null, sessionId = null) {
  try {
    const client = await getGraphClient(userToken, sessionId);
    console.log('🖼️ Calling Graph API for /me/photo/$value');

    const photoResponse = await client.api('/me/photo/$value').get();

    console.log('📦 Response received, type:', typeof photoResponse);
    console.log('📦 Is Buffer?:', Buffer.isBuffer(photoResponse));
    console.log('📦 Is Uint8Array?:', photoResponse instanceof Uint8Array);
    console.log('📦 Constructor name:', photoResponse?.constructor?.name);

    let buffer;

    // Handle Blob (from browser fetch)
    if (typeof Blob !== 'undefined' && photoResponse instanceof Blob) {
      console.log('📦 Response is a Blob - converting to Buffer...');
      // Convert Blob to Buffer
      const arrayBuffer = await photoResponse.arrayBuffer();
      buffer = Buffer.from(arrayBuffer);
      console.log('✅ Blob converted to Buffer, size:', buffer.length);
      return buffer;
    }

    // Handle Buffer
    if (Buffer.isBuffer(photoResponse)) {
      buffer = photoResponse;
      console.log('✅ Response is a Buffer');
      return buffer;
    }

    // Handle Uint8Array
    if (photoResponse instanceof Uint8Array) {
      buffer = Buffer.from(photoResponse);
      console.log('✅ Converted Uint8Array to Buffer');
      return buffer;
    }

    // Handle Stream
    if (photoResponse && photoResponse._readableState) {
      console.log('📦 Response is a Stream - converting to Buffer...');
      const chunks = [];

      return new Promise((resolve, reject) => {
        photoResponse.on('data', (chunk) => {
          chunks.push(chunk);
        });
        photoResponse.on('end', () => {
          buffer = Buffer.concat(chunks);
          console.log('✅ Stream converted to Buffer, size:', buffer.length);
          resolve(buffer);
        });
        photoResponse.on('error', (err) => {
          console.error('❌ Stream error:', err);
          reject(err);
        });
      });
    }

    // Handle other objects
    if (typeof photoResponse === 'object' && photoResponse !== null) {
      if (photoResponse.data) {
        buffer = Buffer.from(photoResponse.data);
        console.log('✅ Extracted data from object wrapper');
        return buffer;
      }
    }

    // Fallback
    console.warn('⚠️ Unknown response type:', photoResponse?.constructor?.name);
    return null;
  } catch (error) {
    console.error('❌ Error getting user profile photo:', error.message);
    console.error('Error code:', error.code);
    return null;
  }
}

module.exports = {
  getAuthUrl,
  getAccessTokenByAuthCode,
  getAccessTokenByRefreshToken,
  refreshTokenSilently,
  getAccessTokenAppOnly,
  getGraphClient,
  getRecentEmails,
  searchEmails,
  sendEmail,
  getCalendarEvents,
  createCalendarEvent,
  updateCalendarEvent,
  deleteCalendarEvents,
  getRecentFiles,
  searchFiles,
  getTeams,
  getTeamChannels,
  getUserProfile,
  searchContactEmail,
  sendTeamsMessage,
  getTeamsMessages,
  deleteTeamsMessage,
  getSenderProfile,
  getRecentSentEmails,
  deleteEmail,
  deleteSentEmail,
  getUserProfilePhoto
};