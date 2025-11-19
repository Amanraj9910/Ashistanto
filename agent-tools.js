const graphTools = require('./graph-tools');

// =========================
// 🔧 Define available tools
// =========================
const tools = [
  {
    type: 'function',
    function: {
      name: 'search_contact_email',
      description: 'Search for a person\'s email address by their name. Searches personal contacts, People API, organization directory, and generates email as fallback.',
      parameters: {
        type: 'object',
        properties: {
          name: {
            type: 'string',
            description: 'Person\'s name to search for (first name, last name, or full name)'
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
      description: 'Send an email with proper greeting and signature. Automatically finds recipient email from contacts/directory. Supports CC recipients.',
      parameters: {
        type: 'object',
        properties: {
          recipient_name: {
            type: 'string',
            description: 'Primary recipient\'s full name (e.g., "John Doe")'
          },
          subject: {
            type: 'string',
            description: 'Email subject line'
          },
          body: {
            type: 'string',
            description: 'Plain text email body content'
          },
          cc_recipients: {
            type: 'array',
            items: { type: 'string' },
            description: 'Optional array of names to CC on the email (e.g., ["Jane Smith", "Bob Johnson"])',
            default: []
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
      description: 'Create a calendar event or Microsoft Teams meeting with join link. Automatically finds attendee emails from contacts/directory.',
      parameters: {
        type: 'object',
        properties: {
          subject: { 
            type: 'string', 
            description: 'Meeting subject/title' 
          },
          start: { 
            type: 'string', 
            description: 'Start date and time in ISO format (e.g., "2024-12-25T10:00:00")' 
          },
          end: { 
            type: 'string', 
            description: 'End date and time in ISO format (e.g., "2024-12-25T11:00:00")' 
          },
          location: { 
            type: 'string', 
            description: 'Meeting location (optional, not needed for Teams meetings)', 
            default: '' 
          },
          attendeeNames: {
            type: 'array',
            items: { type: 'string' },
            description: 'Array of attendee names (will be automatically searched in contacts)',
            default: []
          },
          isTeamsMeeting: {
            type: 'boolean',
            description: 'Set to true to create a Microsoft Teams meeting with join link. Falls back to regular meeting if Teams meeting creation fails.',
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
      name: 'send_teams_message',
      description: 'Send a direct chat message to someone on Microsoft Teams. Automatically finds the person in contacts/directory and creates or finds the chat.',
      parameters: {
        type: 'object',
        properties: {
          recipient_name: {
            type: 'string',
            description: 'Name of the person to message on Teams (e.g., "John Doe")'
          },
          message: {
            type: 'string',
            description: 'The message content to send'
          }
        },
        required: ['recipient_name', 'message']
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
          count: { 
            type: 'number', 
            description: 'Number of files to retrieve (default 10)',
            default: 10 
          }
        },
        required: []
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'search_files',
      description: 'Search for files in OneDrive/SharePoint by name or content.',
      parameters: {
        type: 'object',
        properties: {
          query: { 
            type: 'string', 
            description: 'Search query (file name or content keywords)' 
          }
        },
        required: ['query']
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_teams',
      description: 'Get list of Microsoft Teams the user is a member of.',
      parameters: { 
        type: 'object', 
        properties: {}, 
        required: [] 
      }
    }
  },
  {
    type: 'function',
    function: {
      name: 'get_user_profile',
      description: 'Get the current user\'s profile information including name, email, job title, and location.',
      parameters: { 
        type: 'object', 
        properties: {}, 
        required: [] 
      }
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
  search_contact_email: graphTools.searchContactEmail,
  send_teams_message: graphTools.sendTeamsMessage  // ✅ NEW: Teams messaging
};

// ======================================
// 🚀 Execute a tool with proper userToken
// ======================================
async function executeTool(functionName, args = {}, userToken = null) {
  const func = functionMap[functionName];
  if (!func) throw new Error(`Unknown function: ${functionName}`);

  console.log(`\n🧩 Executing tool: ${functionName}`);
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
        params.push(args.cc_recipients || []);  // ✅ NEW: CC recipients support
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
        params.push(args.isTeamsMeeting || false);  // ✅ ENHANCED: Teams meeting with link
        params.push(userToken);
        break;
        
      case 'send_teams_message':  // ✅ NEW: Teams direct messaging
        params.push(args.recipient_name);
        params.push(args.message);
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
    
    // ✅ IMPROVED: Pretty print result with truncation for large responses
    const resultStr = JSON.stringify(result, null, 2);
    if (resultStr.length > 500) {
      console.log(`   ✅ Result (truncated):`, resultStr.substring(0, 500) + '...');
    } else {
      console.log(`   ✅ Result:`, resultStr);
    }
    
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