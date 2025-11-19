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
    'Contacts.Read',
    'OnlineMeetings.ReadWrite',
    'Chat.ReadWrite',
    'offline_access'
  ];
  
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
        'User.Read',
        'Contacts.Read',
        'OnlineMeetings.ReadWrite',
        'Chat.ReadWrite'
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
        'User.Read',
        'Contacts.Read',
        'OnlineMeetings.ReadWrite',
        'Chat.ReadWrite'
      ]
    };
    
    const response = await msalClient.acquireTokenByRefreshToken(tokenRequest);
    return response.accessToken;
  } catch (error) {
    console.error('Error refreshing access token:', error);
    throw new Error('Failed to refresh access token');
  }
}

// Get access token using client credentials flow (fallback)
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

// Initialize Graph client
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
    return {
      displayName: 'User',
      email: userEmail || 'sender@hoshodigital.com',
      jobTitle: '',
      department: '',
      officeLocation: ''
    };
  }
}

// ============== CONTACT SEARCH FUNCTIONS ==============

/**
 * 🔍 Search for contact email by name from Graph API
 * First tries user's contacts, then organization directory
 * Falls back to email generation if not found
 */
async function searchContactEmail(name, userToken = null) {
  try {
    console.log(`🔍 Searching for contact: "${name}"`);
    const client = await getGraphClient(userToken);
    
    // Step 1: Search in user's personal contacts
    try {
      console.log('  → Searching personal contacts...');
      const contacts = await client
        .api('/me/contacts')
        .filter(`startswith(displayName,'${name}') or startswith(givenName,'${name}') or startswith(surname,'${name}')`)
        .select('displayName,emailAddresses,givenName,surname')
        .top(5)
        .get();
      
      if (contacts.value && contacts.value.length > 0) {
        console.log(`  ✅ Found ${contacts.value.length} contact(s) in personal contacts`);
        return contacts.value.map(contact => ({
          name: contact.displayName,
          email: contact.emailAddresses?.[0]?.address || 'No email',
          source: 'personal_contacts'
        }));
      }
    } catch (err) {
      console.log('  ⚠ Personal contacts search failed:', err.message);
    }
    
    // Step 2: Search in People API (combines contacts, directory, and frequent contacts)
    try {
      console.log('  → Searching People API...');
      const people = await client
        .api('/me/people')
        .search(`"${name}"`)
        .select('displayName,emailAddresses,givenName,surname')
        .top(5)
        .get();
      
      if (people.value && people.value.length > 0) {
        console.log(`  ✅ Found ${people.value.length} person(s) in People API`);
        return people.value
          .filter(person => person.emailAddresses && person.emailAddresses.length > 0)
          .map(person => ({
            name: person.displayName,
            email: person.emailAddresses[0].address,
            source: 'people_api'
          }));
      }
    } catch (err) {
      console.log('  ⚠ People API search failed:', err.message);
    }
    
    // Step 3: Search in organization directory
    try {
      console.log('  → Searching organization directory...');
      const users = await client
        .api('/users')
        .filter(`startswith(displayName,'${name}') or startswith(givenName,'${name}') or startswith(surname,'${name}')`)
        .select('displayName,mail,userPrincipalName,givenName,surname')
        .top(5)
        .get();
      
      if (users.value && users.value.length > 0) {
        console.log(`  ✅ Found ${users.value.length} user(s) in organization`);
        return users.value.map(user => ({
          name: user.displayName,
          email: user.mail || user.userPrincipalName,
          source: 'organization_directory'
        }));
      }
    } catch (err) {
      console.log('  ⚠ Organization directory search failed:', err.message);
    }
    
    // Step 4: Fallback - Generate email from name
    console.log('  → No contacts found, generating email from name...');
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
    console.log(`  ✅ Generated email: ${generatedEmail}`);
    
    return [{
      name: `${firstName} ${lastName}`,
      email: generatedEmail,
      source: 'generated'
    }];
  } catch (error) {
    console.error('❌ Error in searchContactEmail:', error);
    throw new Error(`Failed to find contact email for "${name}"`);
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

/**
 * 📧 Send email with contact search, CC support, and proper formatting
 */
async function sendEmail(recipient_name, subject, body, ccRecipients = [], userToken = null) {
  try {
    console.log(`📧 Sending email to: ${recipient_name}`);
    
    // Get sender's profile
    const senderProfile = await getSenderProfile(userToken);
    
    // Search for recipient email using contact search
    const recipientResults = await searchContactEmail(recipient_name, userToken);
    
    if (!recipientResults || recipientResults.length === 0) {
      throw new Error(`Could not find email for recipient: ${recipient_name}`);
    }
    
    const recipient = recipientResults[0];
    console.log(`  ✅ Found recipient: ${recipient.email} (source: ${recipient.source})`);
    
    // Parse recipient name for greeting
    const nameParts = recipient_name.trim().split(/\s+/);
    const firstName = nameParts[0];
    
    // Process CC recipients
    const ccEmailAddresses = [];
    if (ccRecipients && ccRecipients.length > 0) {
      console.log(`  📎 Processing ${ccRecipients.length} CC recipient(s)...`);
      for (const ccName of ccRecipients) {
        try {
          const ccResults = await searchContactEmail(ccName, userToken);
          if (ccResults && ccResults.length > 0) {
            ccEmailAddresses.push({
              emailAddress: {
                address: ccResults[0].email,
                name: ccResults[0].name
              }
            });
            console.log(`    ✅ CC: ${ccResults[0].email}`);
          }
        } catch (err) {
          console.log(`    ⚠ Could not find CC recipient: ${ccName}`);
        }
      }
    }
    
    // Clean the body
    const cleanBody = body
      .replace(/<[^>]*>/g, '')
      .replace(/\s+/g, ' ')
      .trim();

    // Build plain text email
    const plainTextBody = `
Hi ${firstName},

${cleanBody}

Best regards,
${senderProfile.displayName || 'User'}
`;

    // Send the email
    const client = await getGraphClient(userToken);
    const message = {
      message: {
        subject: subject,
        body: {
          contentType: 'Text',
          content: plainTextBody.trim()
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
      }
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
      .select('subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeeting')
      .orderby('start/dateTime')
      .top(10)
      .get();
    
    return events.value.map(event => ({
      subject: event.subject,
      start: new Date(event.start.dateTime).toLocaleString(),
      end: new Date(event.end.dateTime).toLocaleString(),
      location: event.location?.displayName || 'No location',
      organizer: event.organizer?.emailAddress?.name || 'Unknown',
      attendees: event.attendees?.length || 0,
      isTeamsMeeting: event.isOnlineMeeting || false,
      joinUrl: event.onlineMeeting?.joinUrl || null
    }));
  } catch (error) {
    console.error('Error getting calendar events:', error);
    throw new Error('Failed to retrieve calendar events');
  }
}

/**
 * 📅 Create calendar event with Teams meeting support and attendee search
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
    console.log(`   Teams meeting requested: ${isTeamsMeeting}`);
    
    if (!userToken) throw new Error('Missing user token.');

    const client = await getGraphClient(userToken);

    // Process attendees - search for their emails
    const attendeeEmails = [];
    if (attendeeNames && attendeeNames.length > 0) {
      console.log(`   Processing ${attendeeNames.length} attendee(s)...`);
      for (const name of attendeeNames) {
        try {
          const results = await searchContactEmail(name, userToken);
          if (results && results.length > 0) {
            attendeeEmails.push({
              emailAddress: {
                address: results[0].email,
                name: results[0].name
              },
              type: 'required'
            });
            console.log(`     ✅ Attendee: ${results[0].email} (${results[0].source})`);
          }
        } catch (err) {
          console.log(`     ⚠ Could not find attendee: ${name}`);
        }
      }
    }

    // Build the calendar event object
    const event = {
      subject,
      start: { dateTime: start, timeZone: 'Asia/Kolkata' },
      end: { dateTime: end, timeZone: 'Asia/Kolkata' },
      location: location ? { displayName: location } : undefined,
      attendees: attendeeEmails
    };

    // Try to create Teams meeting if requested
    if (isTeamsMeeting) {
      try {
        console.log('   → Attempting to create Teams meeting...');
        event.isOnlineMeeting = true;
        event.onlineMeetingProvider = 'teamsForBusiness';
        
        const createdEvent = await client.api('/me/events').post(event);
        
        console.log('   ✅ Teams meeting created successfully');
        
        return {
          success: true,
          eventId: createdEvent.id,
          subject: createdEvent.subject,
          attendees: attendeeEmails.map((a) => a.emailAddress.name).join(', '),
          attendeeCount: attendeeEmails.length,
          startTime: new Date(start).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
          endTime: new Date(end).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
          isTeamsMeeting: true,
          joinUrl: createdEvent.onlineMeeting?.joinUrl || null,
          message: `Teams meeting "${subject}" created successfully with ${attendeeEmails.length} attendee(s). Join URL: ${createdEvent.onlineMeeting?.joinUrl || 'Pending'}`
        };
      } catch (teamsError) {
        console.log('   ⚠ Teams meeting creation failed, creating regular calendar event:', teamsError.message);
        // Fall through to create regular event
      }
    }

    // Create regular calendar event (no Teams)
    event.isOnlineMeeting = false;
    delete event.onlineMeetingProvider;
    
    const createdEvent = await client.api('/me/events').post(event);
    
    console.log('   ✅ Calendar event created successfully');

    return {
      success: true,
      eventId: createdEvent.id,
      subject: createdEvent.subject,
      attendees: attendeeEmails.map((a) => a.emailAddress.name).join(', '),
      attendeeCount: attendeeEmails.length,
      startTime: new Date(start).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      endTime: new Date(end).toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
      isTeamsMeeting: false,
      joinUrl: null,
      message: `Calendar event "${subject}" created successfully with ${attendeeEmails.length} attendee(s)`
    };
  } catch (error) {
    console.error('❌ Error creating calendar event:', error);
    throw new Error('Failed to create calendar event: ' + error.message);
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

/**
 * 💬 Send Teams chat message to a user
 */
async function sendTeamsMessage(recipientName, message, userToken = null) {
  try {
    console.log(`💬 Sending Teams message to: ${recipientName}`);
    
    if (!userToken) throw new Error('User token required for Teams messaging');
    
    const client = await getGraphClient(userToken);
    
    // Step 1: Find recipient's user ID
    console.log('   → Searching for recipient...');
    const recipientResults = await searchContactEmail(recipientName, userToken);
    
    if (!recipientResults || recipientResults.length === 0) {
      throw new Error(`Could not find user: ${recipientName}`);
    }
    
    const recipientEmail = recipientResults[0].email;
    console.log(`   ✅ Found recipient email: ${recipientEmail}`);
    
    // Step 2: Get recipient's user ID
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
    
    // Step 3: Create or get existing chat
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
      // If chat already exists, find it
      console.log('   → Chat may already exist, searching...');
      const chats = await client
        .api('/me/chats')
        .filter(`chatType eq 'oneOnOne'`)
        .expand('members')
        .get();
      
      // Find chat with this user
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
    
    // Step 4: Send message
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
      message: `Teams message sent to ${recipientResults[0].name}`,
      recipientName: recipientResults[0].name,
      recipientEmail: recipientEmail,
      chatId: chatId,
      messageId: sentMessage.id
    };
    
  } catch (error) {
    console.error('❌ Error sending Teams message:', error);
    throw new Error(`Failed to send Teams message: ${error.message}`);
  }
}

// ============== USER FUNCTIONS ==============

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

// ============== SHAREPOINT FUNCTIONS ==============

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

// ============== HELPER FUNCTIONS ==============

// Generate email from first name and last name (fallback)
function generateEmailFromName(firstName, lastName) {
  const firstNameClean = firstName.trim().toLowerCase();
  const lastNameClean = lastName.trim().toLowerCase();
  const emailUsername = firstNameClean + lastNameClean.charAt(0);
  
  const userEmailEnv = process.env.MICROSOFT_USER_EMAIL || 'amanr@hoshodigital.com';
  const domain = userEmailEnv.split('@')[1] || 'hoshodigital.com';
  
  return `${emailUsername}@${domain}`;
}

function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
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
  sendTeamsMessage,
  generateEmailFromName,
  getSenderProfile
};