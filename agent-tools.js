const graphTools = require('./graph-tools');
const actionPreview = require('./action-preview');

// =========================
// 🔧 Define available tools
// =========================
const tools = [
  {
    type: 'function',
    function: {
      name: 'search_contact_email',
      description: 'Search for a person\'s email address by their name.',
      parameters: {
        type: 'object',
        properties: {
          name: { type: 'string' }
        },
        required: ['name']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_recent_emails',
      description: 'Get recent emails.',
      parameters: {
        type: 'object',
        properties: {
          count: { type: 'number', default: 5 }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'search_emails',
      description: 'Search inbox emails.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string' }
        },
        required: ['query']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'send_email',
      description: 'Send formatted email.',
      parameters: {
        type: 'object',
        properties: {
          recipient_name: { type: 'string' },
          subject: { type: 'string' },
          body: { type: 'string' },
          cc_recipients: { type: 'array', items: { type: 'string' }, default: [] }
        },
        required: ['recipient_name', 'subject', 'body']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_calendar_events',
      description: 'Get calendar events.',
      parameters: {
        type: 'object',
        properties: {
          days: { type: 'number', default: 7 }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'create_calendar_event',
      description: 'Create Teams/normal meeting.',
      parameters: {
        type: 'object',
        properties: {
          subject: { type: 'string' },
          start: { type: 'string' },
          end: { type: 'string' },
          location: { type: 'string', default: '' },
          attendeeNames: {
            type: 'array',
            items: { type: 'string' },
            default: []
          },
          isTeamsMeeting: { type: 'boolean', default: false }
        },
        required: ['subject', 'start', 'end']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'update_calendar_event',
      description: 'Update an existing calendar event/meeting. Finds it by subject, then edits it.',
      parameters: {
        type: 'object',
        properties: {
          subject: {
            type: 'string',
            description: 'The current subject/title of the meeting to find it'
          },
          new_subject: { type: 'string', description: 'New subject for the meeting (optional)' },
          new_start_time: { type: 'string', description: 'New start time in ISO format (optional)' },
          new_end_time: { type: 'string', description: 'New end time in ISO format (optional)' },
          add_attendees: {
            type: 'array',
            items: { type: 'string' },
            description: 'List of names or emails to ADD as attendees (optional)'
          }
        },
        required: ['subject']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'send_teams_message',
      description: 'send Teams message.',
      parameters: {
        type: 'object',
        properties: {
          recipient_name: { type: 'string' },
          message: { type: 'string' }
        },
        required: ['recipient_name', 'message']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_recent_files',
      description: 'get recent files.',
      parameters: {
        type: 'object',
        properties: {
          count: { type: 'number', default: 10 }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'search_files',
      description: 'search files.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string' }
        },
        required: ['query']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_teams',
      description: 'get list of teams.',
      parameters: {
        type: 'object',
        properties: {}
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'update_calendar_event',
      description: 'Update an existing calendar event/meeting. Finds it by subject, then edits it.',
      parameters: {
        type: 'object',
        properties: {
          subject: {
            type: 'string',
            description: 'The current subject/title of the meeting to find it'
          },
          new_subject: { type: 'string', description: 'New subject for the meeting (optional)' },
          new_start_time: { type: 'string', description: 'New start time in ISO format (optional)' },
          new_end_time: { type: 'string', description: 'New end time in ISO format (optional)' },
          add_attendees: {
            type: 'array',
            items: { type: 'string' },
            description: 'List of names or emails to ADD as attendees (optional)'
          }
        },
        required: ['subject']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_user_profile',
      description: 'get user profile.',
      parameters: {
        type: 'object',
        properties: {}
      }
    }
  },

  // ============== DELETION TOOLS =================
  {
    type: 'function',
    function: {
      name: 'get_sent_emails',
      description: 'get recent sent emails.',
      parameters: {
        type: 'object',
        properties: {
          count: { type: 'number', default: 10 }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'delete_sent_email',
      description: 'Delete a sent email from the Sent Items folder. Can delete by subject or recipient name. If no filters given, deletes the most recent sent email.',
      parameters: {
        type: 'object',
        properties: {
          subject: {
            type: 'string',
            description: 'Part of the email subject to match (optional)'
          },
          recipient_email: {
            type: 'string',
            description: 'Recipient name or email to match (optional)'
          }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'delete_calendar_event',
      description: 'Delete a calendar event/meeting by its subject.',
      parameters: {
        type: 'object',
        properties: {
          subject: {
            type: 'string',
            description: 'The meeting/event subject to delete'
          }
        },
        required: ['subject']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'delete_teams_message',
      description: 'Delete a Teams chat message. Can delete by message content or the most recent message you sent. Note: Only messages you sent can be deleted.',
      parameters: {
        type: 'object',
        properties: {
          chat_id: {
            type: 'string',
            description: 'Chat ID (optional - will search recent chats if not provided)'
          },
          message_id: {
            type: 'string',
            description: 'Message ID (optional - will find your most recent message if not provided)'
          },
          message_content: {
            type: 'string',
            description: 'Part of the message content to match (optional)'
          }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_teams_messages',
      description: 'Get recent Teams chat messages to see message IDs for deletion.',
      parameters: {
        type: 'object',
        properties: {
          chat_id: { type: 'string' },
          count: { type: 'number', default: 10 }
        }
      }
    }
  }
];

// =================================================
// 🔗 Map tool names to actual functions
// =================================================
const functionMap = {
  get_recent_emails: graphTools.getRecentEmails,
  search_emails: graphTools.searchEmails,
  send_email: graphTools.sendEmail,
  get_calendar_events: graphTools.getCalendarEvents,
  create_calendar_event: graphTools.createCalendarEvent,
  update_calendar_event: graphTools.updateCalendarEvent,
  get_recent_files: graphTools.getRecentFiles,
  search_files: graphTools.searchFiles,
  get_teams: graphTools.getTeams,
  get_user_profile: graphTools.getUserProfile,
  search_contact_email: graphTools.searchContactEmail,
  send_teams_message: graphTools.sendTeamsMessage,

  // deletion tools
  get_sent_emails: graphTools.getRecentSentEmails,
  delete_sent_email: graphTools.deleteSentEmail,
  delete_inbox_email: graphTools.deleteInboxEmail,

  delete_calendar_event: graphTools.deleteCalendarEvents,

  delete_teams_message: graphTools.deleteTeamsMessage,

  get_teams_messages: graphTools.getTeamsMessages,
};

// =================================================
// 🚀 Execute a tool with proper parameter order
// =================================================
// @param {string} functionName - Name of the tool to execute
// @param {object} args - Arguments for the tool
// @param {string} userToken - User's access token
// @param {string} sessionId - Session ID for the user
// @param {boolean} skipConfirmation - If true, skip confirmation flow (used when action already confirmed)
async function executeTool(functionName, args = {}, userToken = null, sessionId = null, skipConfirmation = false) {
  const func = functionMap[functionName];
  if (!func) throw new Error(`Unknown function: ${functionName}`);

  // Actions that require user confirmation
  const confirmationRequiredActions = ['send_email', 'send_teams_message', 'delete_sent_email', 'delete_inbox_email', 'delete_teams_message', 'update_calendar_event'];

  // If action needs confirmation AND we're not skipping it, validate user and return preview
  if (confirmationRequiredActions.includes(functionName) && sessionId && !skipConfirmation) {
    try {
      let actionData = {};
      let validatedRecipientData = null;

      // ✅ EARLY VALIDATION: Validate recipient BEFORE creating preview
      if (functionName === 'send_email') {
        actionData = {
          recipientName: args.recipient_name,
          subject: args.subject,
          body: args.body,
          ccRecipients: args.cc_recipients || []
        };

        // Validate recipient exists
        console.log(`🔍 Validating recipient: ${args.recipient_name}`);
        const searchResult = await graphTools.searchContactEmail(args.recipient_name, userToken, sessionId);

        if (!searchResult.found) {
          // ❌ User not found - return error immediately (no preview)
          console.log(`  ❌ Recipient not found: ${args.recipient_name}`);
          return {
            success: false,
            notFound: true,
            searchedName: searchResult.searchedName,
            message: searchResult.message || `I couldn't find anyone named "${args.recipient_name}" in the organization. Please verify the name or provide their email address.`
          };
        }

        // ✅ User found - cache the validated data
        validatedRecipientData = {
          recipientName: searchResult.results[0].name,
          recipientEmail: searchResult.results[0].email,
          source: searchResult.results[0].source
        };
        console.log(`  ✅ Recipient validated: ${validatedRecipientData.recipientEmail}`);

      } else if (functionName === 'send_teams_message') {
        actionData = {
          recipientName: args.recipient_name,
          message: args.message
        };

        // Validate recipient exists
        console.log(`🔍 Validating recipient: ${args.recipient_name}`);
        const searchResult = await graphTools.searchContactEmail(args.recipient_name, userToken, sessionId);

        if (!searchResult.found) {
          // ❌ User not found - return error immediately (no preview)
          console.log(`  ❌ Recipient not found: ${args.recipient_name}`);
          return {
            success: false,
            notFound: true,
            searchedName: searchResult.searchedName,
            message: searchResult.message || `I couldn't find anyone named "${args.recipient_name}" in the organization. Please verify the name or provide their email address.`
          };
        }

        // ✅ User found - cache the validated data
        validatedRecipientData = {
          recipientName: searchResult.results[0].name,
          recipientEmail: searchResult.results[0].email,
          source: searchResult.results[0].source
        };
        console.log(`  ✅ Recipient validated: ${validatedRecipientData.recipientEmail}`);

      } else if (functionName === 'delete_sent_email') {
        // Find the email to delete first
        console.log(`🔍 Finding email to delete...`);
        const searchResult = await graphTools.deleteSentEmail(args.subject || null, args.recipient_email || null, userToken, true); // true = preview mode

        if (!searchResult.success || searchResult.notFound) {
          return {
            success: false,
            notFound: true,
            message: searchResult.message || 'No matching email found to delete'
          };
        }

        actionData = {
          subject: searchResult.emailToDelete.subject,
          recipient: searchResult.emailToDelete.recipient,
          sentDate: searchResult.emailToDelete.sentDate,
          messageId: searchResult.emailToDelete.id
        };
        console.log(`  ✅ Found email to delete: "${actionData.subject}"`);

      } else if (functionName === 'delete_inbox_email') {
        // Find the email to delete first
        console.log(`🔍 Finding inbox email to delete...`);
        const searchResult = await graphTools.deleteInboxEmail(args.subject || null, args.sender_email || null, userToken, true); // true = preview mode

        if (!searchResult.success || searchResult.notFound) {
          return {
            success: false,
            notFound: true,
            message: searchResult.message || 'No matching inbox email found to delete'
          };
        }

        actionData = {
          subject: searchResult.emailToDelete.subject,
          sender: searchResult.emailToDelete.sender,
          receivedDate: searchResult.emailToDelete.receivedDate,
          messageId: searchResult.emailToDelete.id
        };
        console.log(`  ✅ Found inbox email to delete: "${actionData.subject}"`);

      } else if (functionName === 'delete_teams_message') {
        // Find the Teams message to delete first
        console.log(`🔍 Finding Teams message to delete...`);
        const searchResult = await graphTools.deleteTeamsMessage(args.chat_id || null, args.message_id || null, args.message_content || null, userToken, true); // true = preview mode

        if (!searchResult.success || searchResult.notFound) {
          return {
            success: false,
            notFound: true,
            message: searchResult.message || 'No matching Teams message found to delete'
          };
        }

        actionData = {
          messageContent: searchResult.messageToDelete.content,
          sentDate: searchResult.messageToDelete.sentDate,
          chatId: searchResult.messageToDelete.chatId,
          messageId: searchResult.messageToDelete.messageId
        };
        console.log(`  ✅ Found Teams message to delete`);

      } else if (functionName === 'update_calendar_event') {
        console.log(`🔍 Finding calendar event to update...`);
        const searchResult = await graphTools.updateCalendarEvent(
          args.subject,
          args.add_attendees || [],
          args.new_subject || null,
          args.new_start_time || null,
          args.new_end_time || null,
          userToken,
          true // preview mode flag
        );

        if (!searchResult.success) {
          return {
            success: false,
            notFound: true,
            message: searchResult.message || 'No matching calendar event found'
          };
        }

        const previewData = searchResult.previewData;

        actionData = {
          subject: previewData.originalSubject,
          newSubject: previewData.newSubject,
          newStart: previewData.newStart ? new Date(previewData.newStart).toLocaleString() : 'No change',
          newEnd: previewData.newEnd ? new Date(previewData.newEnd).toLocaleString() : 'No change',
          attendees: previewData.attendees
        };
        console.log(`  ✅ Found calendar event to update: "${actionData.subject}"`);
      }

      // Create preview with cached validated data
      const preview = await actionPreview.createActionPreview(functionName, actionData, validatedRecipientData);
      return {
        type: 'action_preview',
        preview: preview,
        message: 'Action requires confirmation. Review the preview and confirm to proceed.'
      };
    } catch (error) {
      console.error('❌ Error creating action preview:', error);
      // Return error instead of falling through to execution
      return {
        success: false,
        error: error.message || 'Failed to validate recipient'
      };
    }
  }

  let params = [];

  switch (functionName) {

    case 'get_recent_emails':
      params = [args.count || 5, userToken, sessionId];
      break;

    case 'search_emails':
      params = [args.query, userToken];
      break;

    case 'send_email':
      params = [args.recipient_name, args.subject, args.body, args.cc_recipients || [], userToken, null];
      break;

    case 'get_calendar_events':
      params = [args.days || 7, userToken, sessionId];
      break;

    case 'create_calendar_event':

      console.log(`⚡ Teams meeting flag: ${args.isTeamsMeeting}`);

      params = [
        args.subject,
        args.start,
        args.end,
        args.location || '',
        args.attendeeNames || [],
        args.isTeamsMeeting !== undefined ? args.isTeamsMeeting : true,
        userToken
      ];
      break;

    case 'update_calendar_event':
      params = [
        args.subject,
        args.add_attendees || [],
        args.new_subject || null,
        args.new_start_time || null,
        args.new_end_time || null,
        userToken
      ];
      break;

    case 'send_teams_message':
      params = [args.recipient_name, args.message, userToken, null];
      break;

    case 'get_recent_files':
      params = [args.count || 10, userToken, sessionId];
      break;

    case 'search_files':
      params = [args.query, userToken];
      break;

    case 'get_teams':
    case 'get_user_profile':
      params = [userToken];
      break;

    case 'search_contact_email':
      params = [args.name, userToken, sessionId];
      break;

    case 'get_sent_emails':
      params = [args.count || 10, userToken, sessionId];
      break;

    case 'delete_sent_email':
      params = [args.subject || null, args.recipient_email || null, userToken];
      break;

    case 'delete_inbox_email':
      params = [args.subject || null, args.sender_email || null, userToken];
      break;

    case 'delete_calendar_event':
      params = [args.subject || null, null, null, userToken];
      break;

    case 'delete_teams_message':
      // Pass: chatId, messageId, messageContent, userToken
      params = [args.chat_id || null, args.message_id || null, args.message_content || null, userToken];
      break;

    case 'get_teams_messages':
      params = [args.chat_id || null, args.count || 10, userToken];
      break;
  }

  try {
    const result = await func(...params);
    return result;
  } catch (error) {
    console.error(`❌ Tool execution failed for ${functionName}:`, error.message);
    return {
      success: false,
      error: `Tool execution failed: ${error.message}. Please inform the user.`
    };
  }
}

// ======================================
// 📦 Export
// ======================================
module.exports = {
  tools,
  executeTool,
  actionPreview // Export action preview module for server.js to use
};
