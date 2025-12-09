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

  let redirectUri;
  if (process.env.REDIRECT_URI) {
    redirectUri = process.env.REDIRECT_URI;
  } else if (process.env.NODE_ENV === 'production') {
    redirectUri = 'https://microsoft-agent-aubbhefsbzagdhha.eastus-01.azurewebsites.net/auth/callback';
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
    scope: scopes.join(' ')
  });

  return `https://login.microsoftonline.com/${process.env.MICROSOFT_TENANT_ID}/oauth2/v2.0/authorize?${params.toString()}`;
}

// Get access token using authorization code (delegated flow)
async function getAccessTokenByAuthCode(code) {
  try {
    const msalClient = initMsalClient();

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

/**
 * 🔍 Search for contact email by name from Graph API
 */
async function searchContactEmail(name, userToken = null) {
  try {
    console.log(`🔍 Searching for contact: "${name}"`);
    const client = await getGraphClient(userToken);
    const searchedName = name.trim();

    // Step 1: Search in user's personal contacts
    try {
      console.log('  → Searching personal contacts...');
      const contacts = await client
        .api('/me/contacts')
        .filter(`startswith(displayName,'${searchedName}') or startswith(givenName,'${searchedName}') or startswith(surname,'${searchedName}')`)
        .select('displayName,emailAddresses,givenName,surname')
        .top(5)
        .get();

      if (contacts.value && contacts.value.length > 0) {
        console.log(`  ✅ Found ${contacts.value.length} contact(s) in personal contacts`);
        const results = contacts.value
          .filter(contact => contact.emailAddresses && contact.emailAddresses.length > 0)
          .map(contact => ({
            name: contact.displayName,
            email: contact.emailAddresses[0].address,
            source: 'personal_contacts'
          }));

        if (results.length > 0) {
          return {
            found: true,
            results: results,
            searchedName: searchedName
          };
        }
      }
    } catch (err) {
      console.log('  ⚠ Personal contacts search failed:', err.message);
    }

    // Step 2: Search in People API
    try {
      console.log('  → Searching People API...');
      const people = await client
        .api('/me/people')
        .search(`"${searchedName}"`)
        .select('displayName,emailAddresses,givenName,surname')
        .top(5)
        .get();

      if (people.value && people.value.length > 0) {
        const results = people.value
          .filter(person => person.emailAddresses && person.emailAddresses.length > 0)
          .map(person => ({
            name: person.displayName,
            email: person.emailAddresses[0].address,
            source: 'people_api'
          }));

        if (results.length > 0) {
          console.log(`  ✅ Found ${results.length} person(s) in People API`);
          return {
            found: true,
            results: results,
            searchedName: searchedName
          };
        }
      }
    } catch (err) {
      console.log('  ⚠ People API search failed:', err.message);
    }

    // Step 3: Search in organization directory
    try {
      console.log('  → Searching organization directory...');
      const users = await client
        .api('/users')
        .filter(`startswith(displayName,'${searchedName}') or startswith(givenName,'${searchedName}') or startswith(surname,'${searchedName}')`)
        .select('displayName,mail,userPrincipalName,givenName,surname')
        .top(5)
        .get();

      if (users.value && users.value.length > 0) {
        console.log(`  ✅ Found ${users.value.length} user(s) in organization`);
        const results = users.value.map(user => ({
          name: user.displayName,
          email: user.mail || user.userPrincipalName,
          source: 'organization_directory'
        }));

        return {
          found: true,
          results: results,
          searchedName: searchedName
        };
      }
    } catch (err) {
      console.log('  ⚠ Organization directory search failed:', err.message);
    }

    console.log(`  ❌ No contact found for: "${searchedName}"`);
    return {
      found: false,
      searchedName: searchedName,
      message: `No user found with name "${searchedName}". Please verify the name or provide the email address directly.`
    };

  } catch (error) {
    console.error('❌ Error in searchContactEmail:', error);
    throw new Error(`Failed to search for contact "${name}": ${error.message}`);
  }
}

// ============== EMAIL FUNCTIONS ==============

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

async function sendEmail(recipient_name, subject, body, ccRecipients = [], userToken = null) {
  try {
    console.log(`📧 Sending email to: ${recipient_name}`);

    const senderProfile = await getSenderProfile(userToken);
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

    const recipient = searchResult.results[0];
    console.log(`  ✅ Found recipient: ${recipient.email} (source: ${recipient.source})`);

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

async function getCalendarEvents(days = 7, userToken = null) {
  try {
    const client = await getGraphClient(userToken);
    const startDate = new Date();
    const endDate = new Date();
    endDate.setDate(endDate.getDate() + days);

    const events = await client
      .api('/me/calendar/events')
      .filter(`start/dateTime ge '${startDate.toISOString()}' and start/dateTime le '${endDate.toISOString()}'`)
      .select('id,subject,start,end,location,attendees,organizer,isOnlineMeeting,onlineMeeting')
      .orderby('start/dateTime')
      .top(50)
      .get();

    return events.value.map(event => ({
      id: event.id,
      subject: event.subject,
      start: new Date(event.start.dateTime).toLocaleString(),
      end: new Date(event.end.dateTime).toLocaleString(),
      startDateTime: event.start.dateTime,
      endDateTime: event.end.dateTime,
      location: event.location?.displayName || 'No location',
      organizer: event.organizer?.emailAddress?.name || 'Unknown',
      attendees: event.attendees?.map(a => a.emailAddress.name || a.emailAddress.address) || [],
      attendeeCount: event.attendees?.length || 0,
      isTeamsMeeting: event.isOnlineMeeting || false,
      joinUrl: event.onlineMeeting?.joinUrl || null
    }));
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

async function sendTeamsMessage(recipientName, message, userToken = null) {
  try {
    console.log(`💬 Sending Teams message to: ${recipientName}`);

    if (!userToken) throw new Error('User token required for Teams messaging');

    const client = await getGraphClient(userToken);
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

    const recipientEmail = searchResult.results[0].email;
    const recipientDisplayName = searchResult.results[0].name;
    console.log(`   ✅ Found recipient email: ${recipientEmail}`);

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

async function deleteTeamsMessage(chatId = null, messageId = null, messageContent = null, userToken = null) {
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

function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// ============== DELETION FUNCTIONS ==============

async function getRecentSentEmails(count = 10, userToken = null) {
  try {
    console.log(`📬 Getting ${count} recent sent emails...`);
    const client = await getGraphClient(userToken);

    const messages = await client
      .api('/me/mailFolders/sentItems/messages')
      .select('id,subject,toRecipients,sentDateTime,bodyPreview')
      .top(count)
      .orderby('sentDateTime DESC')
      .get();

    return {
      success: true,
      emails: messages.value.map(msg => ({
        id: msg.id,
        subject: msg.subject,
        to: msg.toRecipients?.map(r => r.emailAddress.name || r.emailAddress.address).join(', ') || 'Unknown',
        sentDate: new Date(msg.sentDateTime).toLocaleString(),
        preview: msg.bodyPreview?.substring(0, 100) || ''
      }))
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

async function deleteSentEmail(subject = null, recipientEmail = null, userToken = null) {
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
          const address = r.emailAddress?.address || '';
          return address.toLowerCase().includes(recipientLower);
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
  deleteSentEmail
};