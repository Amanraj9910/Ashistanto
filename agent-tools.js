const graphTools = require('./graph-tools');

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
      description: 'delete a sent email.',
      parameters: {
        type: 'object',
        properties: {
          subject: { type: 'string' },
          recipient_email: { type: 'string' }
        }
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'delete_calendar_event',
      description: 'delete calendar event.',
      parameters: {
        type: 'object',
        properties: {
          subject: { type: 'string' }
        },
        required: ['subject']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'delete_teams_message',
      description: 'delete a Teams message.',
      parameters: {
        type: 'object',
        properties: {
          chat_id: { type: 'string' },
          message_id: { type: 'string' }
        },
        required: ['chat_id', 'message_id']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_teams_messages',
      description: 'Get Teams messages.',
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
  get_recent_files: graphTools.getRecentFiles,
  search_files: graphTools.searchFiles,
  get_teams: graphTools.getTeams,
  get_user_profile: graphTools.getUserProfile,
  search_contact_email: graphTools.searchContactEmail,
  send_teams_message: graphTools.sendTeamsMessage,

  // deletion tools
  get_sent_emails: graphTools.getRecentSentEmails,
  delete_sent_email: graphTools.deleteSentEmail,

  delete_calendar_event: graphTools.deleteCalendarEvents,

  delete_teams_message: graphTools.deleteTeamsMessage,

  get_teams_messages: graphTools.getTeamsMessages,
};

// =================================================
// 🚀 Execute a tool with proper parameter order
// =================================================
async function executeTool(functionName, args = {}, userToken = null) {
  const func = functionMap[functionName];
  if (!func) throw new Error(`Unknown function: ${functionName}`);

  let params = [];

  switch (functionName) {

    case 'get_recent_emails':
      params = [args.count || 5, userToken];
      break;

    case 'search_emails':
      params = [args.query, userToken];
      break;

    case 'send_email':
      params = [args.recipient_name, args.subject, args.body, args.cc_recipients || [], userToken];
      break;

    case 'get_calendar_events':
      params = [args.days || 7, userToken];
      break;

    case 'create_calendar_event':

      // 🔥 FORCE TEAMS MEETING ALWAYS
      console.log("⚡ Teams meeting forced ON");

      params = [
        args.subject,
        args.start,
        args.end,
        args.location || '',
        args.attendeeNames || [],
        true,        // ALWAYS TRUE (Teams enabled)
        userToken
      ];
      break;

    case 'send_teams_message':
      params = [args.recipient_name, args.message, userToken];
      break;

    case 'get_recent_files':
      params = [args.count || 10, userToken];
      break;

    case 'search_files':
      params = [args.query, userToken];
      break;

    case 'get_teams':
    case 'get_user_profile':
      params = [userToken];
      break;

    case 'search_contact_email':
      params = [args.name, userToken];
      break;

    case 'get_sent_emails':
      params = [args.count || 10, userToken];
      break;

    case 'delete_sent_email':
      params = [args.subject || null, args.recipient_email || null, userToken];
      break;

    case 'delete_calendar_event':
      params = [args.subject || null, null, null, userToken];
      break;

    case 'delete_teams_message':
      params = [args.chat_id, args.message_id, null, userToken];
      break;

    case 'get_teams_messages':
      params = [args.chat_id || null, args.count || 10, userToken];
      break;
  }

  const result = await func(...params);
  return result;
}

// ======================================
// 📦 Export
// ======================================
module.exports = {
  tools,
  executeTool
};
