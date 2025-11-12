const graphTools = require('./graph-tools');

// =========================
// 🔧 Define available tools
// =========================
const tools = [
  {
    type: 'function',
    function: {
      name: 'search_contact_email',
      description: 'Search for a person\'s email address by their name. Use this when user provides a name but not an email address.',
      parameters: {
        type: 'object',
        properties: {
          name: {
            type: 'string',
            description: 'Person\'s name to search for'
          }
        },
        required: ['name']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_recent_emails',
      description: 'Get recent emails from the user\'s inbox.',
      parameters: {
        type: 'object',
        properties: {
          count: {
            type: 'number',
            description: 'Number of emails to retrieve (default 5, max 20)',
            default: 5
          }
        },
        required: []
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'search_emails',
      description: 'Search for emails by subject or sender.',
      parameters: {
        type: 'object',
        properties: {
          query: {
            type: 'string',
            description: 'Search query (subject or sender name/email)'
          }
        },
        required: ['query']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'send_email',
      description: 'Send a plain text email with proper greeting and signature.',
      parameters: {
        type: 'object',
        properties: {
          recipient_name: {
            type: 'string',
            description: 'Recipient\'s full name (e.g., "Vansh Jain")'
          },
          subject: {
            type: 'string',
            description: 'Email subject line'
          },
          body: {
            type: 'string',
            description: 'Plain text email body'
          }
        },
        required: ['recipient_name', 'subject', 'body']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_calendar_events',
      description: 'Get upcoming calendar events and meetings.',
      parameters: {
        type: 'object',
        properties: {
          days: {
            type: 'number',
            description: 'Number of days to look ahead (default 7)',
            default: 7
          }
        },
        required: []
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'create_calendar_event',
      description: 'Create a new calendar event or meeting with attendees.',
      parameters: {
        type: 'object',
        properties: {
          subject: { type: 'string', description: 'Meeting subject' },
          start: { type: 'string', description: 'Start date/time (ISO)' },
          end: { type: 'string', description: 'End date/time (ISO)' },
          location: { type: 'string', description: 'Meeting location', default: '' },
          attendeeNames: {
            type: 'array',
            items: { type: 'string' },
            description: 'Array of attendee names',
            default: []
          },
          isTeamsMeeting: {
            type: 'boolean',
            description: 'Set true for Teams meeting',
            default: false
          }
        },
        required: ['subject', 'start', 'end']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_recent_files',
      description: 'Get recently accessed files from OneDrive/SharePoint.',
      parameters: {
        type: 'object',
        properties: {
          count: { type: 'number', default: 10 }
        },
        required: []
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'search_files',
      description: 'Search for files in OneDrive/SharePoint.',
      parameters: {
        type: 'object',
        properties: {
          query: { type: 'string', description: 'Search query' }
        },
        required: ['query']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_teams',
      description: 'Get list of Microsoft Teams the user is part of.',
      parameters: { type: 'object', properties: {}, required: [] }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_user_profile',
      description: 'Get user profile information.',
      parameters: { type: 'object', properties: {}, required: [] }
    }
  }
];

// ======================================
// 🔗 Map tool names to actual functions
// ======================================
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
  search_contact_email: graphTools.searchContactEmail
};

// ======================================
// 🚀 Execute a tool with proper userToken
// ======================================
async function executeTool(functionName, args = {}, userToken = null) {
  const func = functionMap[functionName];
  if (!func) throw new Error(`Unknown function: ${functionName}`);

  console.log(`🧩 Executing tool: ${functionName}`);
  console.log(`   Args:`, JSON.stringify(args, null, 2));
  console.log(`   Has userToken:`, !!userToken);

  try {
    // Extract parameters from args object
    const params = [];
    
    // Different functions have different parameter signatures
    // We need to match the order defined in graph-tools.js
    switch(functionName) {
      case 'get_recent_emails':
        params.push(args.count || 5);
        params.push(userToken);
        break;
        
      case 'search_emails':
        params.push(args.query);
        params.push(userToken);
        break;
        
      case 'send_email':
        params.push(args.recipient_name);
        params.push(args.subject);
        params.push(args.body);
        params.push(userToken);
        break;
        
      case 'get_calendar_events':
        params.push(args.days || 7);
        params.push(userToken);
        break;
        
      case 'create_calendar_event':
        params.push(args.subject);
        params.push(args.start);
        params.push(args.end);
        params.push(args.location || '');
        params.push(args.attendeeNames || []);
        params.push(args.isTeamsMeeting || false);
        params.push(userToken);
        break;
        
      case 'get_recent_files':
        params.push(args.count || 10);
        params.push(userToken);
        break;
        
      case 'search_files':
        params.push(args.query);
        params.push(userToken);
        break;
        
      case 'get_teams':
        params.push(userToken);
        break;
        
      case 'get_user_profile':
        params.push(userToken);
        break;
        
      case 'search_contact_email':
        params.push(args.name);
        params.push(userToken);
        break;
        
      default:
        throw new Error(`Unhandled function: ${functionName}`);
    }

    console.log(`   Calling function with ${params.length} parameters`);
    const result = await func(...params);
    console.log(`   ✅ Result:`, JSON.stringify(result, null, 2));
    return result;
  } catch (error) {
    console.error(`   ❌ Error executing ${functionName}:`, error.message);
    throw error;
  }
}

// ======================================
// 📦 Export for server.js
// ======================================
module.exports = {
  tools,
  executeTool
};