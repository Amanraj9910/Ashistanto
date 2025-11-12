require('isomorphic-fetch');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ConfidentialClientApplication } = require('@azure/msal-node');

// Get user email from environment variable
const userEmail = process.env.MICROSOFT_USER_EMAIL;

// Initialize MSAL for authentication
function initMsalClient() {
  const config = {
    auth: {
      clientId: process.env.MICROSOFT_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}`,
      clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
    }
  };
  return new ConfidentialClientApplication(config);
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
    'offline_access'
  ];
  
  // Get redirect URI - use environment variable or construct from request
  const redirectUri = process.env.REDIRECT_URI || 'https://microsoft-agent-aubbhefsbzagdhha.eastus-01.azurewebsites.net/auth/callback';
  
  console.log('🔐 Auth URL being generated with redirect_uri:', redirectUri);
  
  const params = new URLSearchParams({
    client_id: process.env.MICROSOFT_CLIENT_ID,
    response_type: 'code',
    redirect_uri: redirectUri,
    response_mode: 'query',
    scope: scopes.join(' ')
  });
  
  return `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize?${params.toString()}`;
}

// Get access token using authorization code (delegated flow)
async function getAccessTokenByAuthCode(code) {
  try {
    const msalClient = initMsalClient();
    const redirectUri = process.env.REDIRECT_URI || 'https://microsoft-agent-aubbhefsbzagdhha.eastus-01.azurewebsites.net/auth/callback';
    
    const tokenRequest = {
      code: code,
      scopes: [
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.ReadWrite',
        'Files.ReadWrite',
        'Sites.Read.All',
        'User.Read'
      ],
      redirectUri: redirectUri,
      codeVerifier: undefined
    };
    
    const response = await msalClient.acquireTokenByCode(tokenRequest);
    return response;
  } catch (error) {
    console.error('Error getting access token by auth code:', error);
    throw new Error('Failed to exchange code for token');
  }
}

// Get access token using refresh token (delegated flow)
async function getAccessTokenByRefreshToken(refreshToken) {
  try {
    const msalClient = initMsalClient();
    const tokenRequest = {
      refreshToken: refreshToken,
      scopes: [
        'Mail.ReadWrite',
        'Mail.Send',
        'Calendars.ReadWrite',
        'Files.ReadWrite',
        'Sites.Read.All',
        'User.Read'
      ]
    };
    
    const response = await msalClient.acquireTokenByRefreshToken(tokenRequest);
    return response.accessToken;
  } catch (error) {
    console.error('Error refreshing access token:', error);
    throw new Error('Failed to refresh access token');
  }
}

// Get access token using client credentials flow (fallback for app-only operations)
async function getAccessTokenAppOnly() {
  try {
    const msalClient = initMsalClient();
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

// Initialize Graph client with proper error handling
async function getGraphClient(userAccessToken = null) {
  let accessToken;
  
  if (userAccessToken) {
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
async function getSenderProfile(userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const user = await client.api('/me').get();
    
    return {
      displayName: user.displayName,
      email: user.mail || user.userPrincipalName,
      jobTitle: user.jobTitle || '',
      department: user.department || '',
      officeLocation: user.officeLocation || ''
    };
  } catch (error) {
    console.error('Error getting sender profile:', error);
    // Return default values if profile fetch fails
    return {
      displayName: 'User',
      email: userEmail || 'sender@hoshodigital.com',
      jobTitle: '',
      department: '',
      officeLocation: ''
    };
  }
}

// ============== EMAIL FUNCTIONS ==============

// Get recent emails from inbox
async function getRecentEmails(count = 5, userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    
    const messages = await client
      .api('/me/messages')
      .select('subject,from,receivedDateTime,bodyPreview,isRead')
      .top(count)
      .orderby('receivedDateTime DESC')
      .get();
    
    return messages.value.map(msg => ({
      subject: msg.subject,
      from: msg.from.emailAddress.name || msg.from.emailAddress.address,
      date: new Date(msg.receivedDateTime).toLocaleString(),
      preview: msg.bodyPreview.substring(0, 100),
      isRead: msg.isRead
    }));
  } catch (error) {
    console.error('Error getting emails:', error);
    throw new Error('Failed to retrieve emails');
  }
}

// Search emails
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

// Send email with proper formatting
async function sendEmail(recipient_name, subject, body, userToken = null) {
  try {
    // Get sender's profile
    const senderProfile = await getSenderProfile(userToken);
    
    // Parse recipient name
    const nameParts = recipient_name.trim().split(/\s+/);
    let firstName, lastName;
    
    if (nameParts.length < 2) {
      firstName = nameParts[0];
      lastName = nameParts[0];
    } else {
      firstName = nameParts[0];
      lastName = nameParts.slice(1).join(' ');
    }
    
    // Generate recipient email
    const recipientEmail = generateEmailFromName(firstName, lastName);

    // 🧹 Clean the body: remove HTML tags
    const cleanBody = body
      .replace(/<[^>]*>/g, '') // remove all HTML tags
      .replace(/\s+/g, ' ')    // normalize spacing
      .trim();

    // 📝 Build the final plain-text message
    const plainTextBody = `
Hi ${firstName},

${cleanBody}

Best regards,
${senderProfile.displayName || 'Aman Raj'}
`;

    // Send the email
    const client = await getGraphClient(userToken);
    const message = {
      message: {
        subject: subject,
        body: {
          contentType: 'Text', // ✅ plain text only
          content: plainTextBody.trim()
        },
        toRecipients: [
          {
            emailAddress: {
              address: recipientEmail,
              name: `${firstName} ${lastName}`
            }
          }
        ]
      }
    };

    await client.api('/me/sendMail').post(message);

    return { 
      success: true, 
      message: `Email sent successfully to ${firstName} ${lastName}`,
      recipientEmail: recipientEmail,
      recipientName: `${firstName} ${lastName}`,
      subject: subject
    };

  } catch (error) {
    console.error('Error sending email:', error);
    throw new Error(`Failed to send email: ${error.message}`);
  }
}


// ============== CALENDAR FUNCTIONS ==============

// Get upcoming calendar events
async function getCalendarEvents(days = 7, userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const startDate = new Date();
    const endDate = new Date();
    endDate.setDate(endDate.getDate() + days);
    
    const events = await client
      .api('/me/calendar/events')
      .filter(`start/dateTime ge '${startDate.toISOString()}' and start/dateTime le '${endDate.toISOString()}'`)
      .select('subject,start,end,location,attendees,organizer')
      .orderby('start/dateTime')
      .top(10)
      .get();
    
    return events.value.map(event => ({
      subject: event.subject,
      start: new Date(event.start.dateTime).toLocaleString(),
      end: new Date(event.end.dateTime).toLocaleString(),
      location: event.location?.displayName || 'No location',
      organizer: event.organizer?.emailAddress?.name || 'Unknown',
      attendees: event.attendees?.length || 0
    }));
  } catch (error) {
    console.error('Error getting calendar events:', error);
    throw new Error('Failed to retrieve calendar events');
  }
}

// Create calendar event with multiple attendees and Teams meeting support
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
    if (!userToken) throw new Error('Missing user token.');

    const client = await getGraphClient(userToken);

    // 🧠 Decode token payload safely
    let decoded;
    try {
      decoded = JSON.parse(Buffer.from(userToken.split('.')[1], 'base64').toString());
    } catch (err) {
      throw new Error('Invalid or malformed access token.');
    }

    // 🧩 Detect if it's an app token (no user identity info)
    const isAppToken = !decoded.upn && !decoded.preferred_username;

    // 👤 Identify which user calendar to use
    const userEmail =
      decoded.upn ||
      decoded.preferred_username ||
      decoded.email ||
      decoded.unique_name ||
      'Amanr@hoshodigital.com'; // ✅ default fallback for app token

    // 🧾 Convert attendee names → emails
    const attendeeEmails = [];
    if (attendeeNames && attendeeNames.length > 0) {
      for (const name of attendeeNames) {
        const nameParts = name.trim().split(/\s+/);
        const firstName = nameParts[0];
        const lastName = nameParts.length > 1 ? nameParts.slice(1).join(' ') : nameParts[0];

        const email = generateEmailFromName(firstName, lastName);
        attendeeEmails.push({
          emailAddress: {
            address: email,
            name: `${firstName.charAt(0).toUpperCase() + firstName.slice(1)} ${lastName
              .charAt(0)
              .toUpperCase() + lastName.slice(1)}`
          },
          type: 'required'
        });
      }
    }

    // 🗓️ Build the calendar event object
    const event = {
      subject,
      start: { dateTime: start, timeZone: 'Asia/Kolkata' },
      end: { dateTime: end, timeZone: 'Asia/Kolkata' },
      location: location ? { displayName: location } : undefined,
      isOnlineMeeting: isTeamsMeeting,
      onlineMeetingProvider: isTeamsMeeting ? 'teamsForBusiness' : undefined,
      attendees: attendeeEmails
    };

    // Remove undefined properties (clean event object)
    Object.keys(event).forEach((key) => event[key] === undefined && delete event[key]);

    // ✅ Choose endpoint based on token type
    const endpoint = isAppToken
      ? `/users/${userEmail}/events`
      : `/me/events`;

    console.log(`📅 Creating calendar event via: ${endpoint}`);

    // 🔐 Create the event using Graph API
    const createdEvent = await client.api(endpoint).post(event);

    console.log('✅ Event created successfully:', createdEvent.id);

    return {
      success: true,
      eventId: createdEvent.id,
      subject: createdEvent.subject,
      attendees: attendeeEmails.map((a) => a.emailAddress.name).join(', '),
      attendeeCount: attendeeEmails.length,
      startTime: new Date(start).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      endTime: new Date(end).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      joinUrl: createdEvent.onlineMeeting?.joinUrl || null,
      message: isTeamsMeeting
        ? `Teams meeting "${subject}" created successfully with ${attendeeEmails.length} attendee(s)`
        : `Calendar event "${subject}" created successfully with ${attendeeEmails.length} attendee(s)`
    };
  } catch (error) {
    console.error('❌ Error creating calendar event:', error);
    throw new Error('Failed to create calendar event: ' + error.message);
  }
}

// ============== SHAREPOINT FUNCTIONS ==============

// Get recent files from OneDrive/SharePoint
async function getRecentFiles(count = 10, userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const files = await client
      .api('/me/drive/recent')
      .top(count)
      .get();
    
    return files.value.map(file => ({
      name: file.name,
      webUrl: file.webUrl,
      lastModified: new Date(file.lastModifiedDateTime).toLocaleString(),
      size: formatFileSize(file.size),
      type: file.file?.mimeType || 'folder'
    }));
  } catch (error) {
    console.error('Error getting recent files:', error);
    throw new Error('Failed to retrieve recent files');
  }
}

// Search files in OneDrive/SharePoint
async function searchFiles(query, userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const files = await client
      .api(`/me/drive/root/search(q='${query}')`)
      .top(10)
      .get();
    
    return files.value.map(file => ({
      name: file.name,
      webUrl: file.webUrl,
      lastModified: new Date(file.lastModifiedDateTime).toLocaleString(),
      size: formatFileSize(file.size)
    }));
  } catch (error) {
    console.error('Error searching files:', error);
    throw new Error('Failed to search files');
  }
}

// ============== TEAMS FUNCTIONS ==============

// Get user's Teams
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

// Get team channels
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

// ============== USER FUNCTIONS ==============

// Search for a contact's email by name
async function searchContactEmail(name, userToken = null) {
  try {
    const nameParts = name.trim().split(/\s+/);
    let firstName, lastName;
    
    if (nameParts.length < 2) {
      firstName = nameParts[0];
      lastName = nameParts[0];
    } else {
      firstName = nameParts[0];
      lastName = nameParts.slice(1).join(' ');
    }
    
    const generatedEmail = generateEmailFromName(firstName, lastName);
    
    return [{
      name: `${firstName} ${lastName}`,
      email: generatedEmail,
      source: 'generated'
    }];
  } catch (error) {
    console.error('Error searching contact:', error);
    throw new Error(`Failed to find contact email for "${name}"`);
  }
}

// Get user profile
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

// ============== HELPER FUNCTIONS ==============

// Generate email from first name and last name
function generateEmailFromName(firstName, lastName) {
  const firstNameClean = firstName.trim().toLowerCase();
  const lastNameClean = lastName.trim().toLowerCase();
  const emailUsername = firstNameClean + lastNameClean.charAt(0);
  
  const userEmailEnv = process.env.MICROSOFT_USER_EMAIL || 'amanr@hoshodigital.com';
  const domain = userEmailEnv.split('@')[1] || 'hoshodigital.com';
  
  return `${emailUsername}@${domain}`;
}

// Search for user by first and last name
async function searchUserByName(firstName, lastName, userToken = null) {
  try {
    const generatedEmail = generateEmailFromName(firstName, lastName);
    
    return {
      success: true,
      displayName: `${firstName} ${lastName}`,
      email: generatedEmail,
      firstName: firstName,
      lastName: lastName,
      source: 'generated'
    };
  } catch (error) {
    console.error('Error generating user email:', error);
    throw new Error(`Failed to generate email for ${firstName} ${lastName}`);
  }
}

function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// Format email body with HTML styling
function formatEmailBodyHTML(messageContent, recipientFirstName, senderProfile) {
  const now = new Date();
  const dateStr = now.toLocaleDateString('en-US', { 
    weekday: 'long', 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  // Clean and format the message content
  let cleanedMessage = messageContent.trim();
  
  // Convert line breaks to HTML
  cleanedMessage = cleanedMessage.replace(/\n/g, '<br>');
  
  // Build HTML email
  const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      line-height: 1.6;
      color: #333;
      max-width: 600px;
      margin: 0 auto;
      padding: 20px;
    }
    .email-container {
      background-color: #ffffff;
      padding: 30px;
      border-radius: 8px;
    }
    .greeting {
      font-size: 16px;
      margin-bottom: 20px;
      color: #2c3e50;
    }
    .content {
      font-size: 15px;
      margin-bottom: 25px;
      color: #34495e;
      line-height: 1.8;
    }
    .signature {
      margin-top: 30px;
      padding-top: 20px;
      border-top: 2px solid #e0e0e0;
    }
    .signature-name {
      font-weight: 600;
      font-size: 16px;
      color: #2c3e50;
      margin-bottom: 5px;
    }
    .signature-title {
      font-size: 14px;
      color: #7f8c8d;
      margin-bottom: 3px;
    }
    .signature-contact {
      font-size: 13px;
      color: #95a5a6;
      margin-top: 10px;
    }
    .footer {
      margin-top: 30px;
      padding-top: 15px;
      border-top: 1px solid #ecf0f1;
      font-size: 12px;
      color: #95a5a6;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="email-container">
    <div class="greeting">
      Dear ${recipientFirstName},
    </div>
    
    <div class="content">
      ${cleanedMessage}
    </div>
    
    <div class="signature">
      <div class="signature-name">${senderProfile.displayName}</div>
      ${senderProfile.jobTitle ? `<div class="signature-title">${senderProfile.jobTitle}</div>` : ''}
      ${senderProfile.department ? `<div class="signature-title">${senderProfile.department}</div>` : ''}
      <div class="signature-contact">
        ${senderProfile.email}
        ${senderProfile.officeLocation ? ` | ${senderProfile.officeLocation}` : ''}
      </div>
    </div>
    
    <div class="footer">
      Sent on ${dateStr}
    </div>
  </div>
</body>
</html>
  `.trim();
  
  return htmlBody;
}

module.exports = {
  getAuthUrl,
  getAccessTokenByAuthCode,
  getAccessTokenByRefreshToken,
  getAccessTokenAppOnly,
  getRecentEmails,
  searchEmails,
  sendEmail,
  getCalendarEvents,
  createCalendarEvent,
  getRecentFiles,
  searchFiles,
  getTeams,
  getTeamChannels,
  getUserProfile,
  searchContactEmail,
  generateEmailFromName,
  searchUserByName,
  formatEmailBodyHTML,
  getSenderProfile
};