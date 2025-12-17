/**
 * ============================================================
 * 📊 DATA FORMATTERS MODULE
 * ============================================================
 * 
 * Centralized formatting for all data presented to users:
 * - Calendar Events (meetings)
 * - Files (SharePoint/OneDrive)
 * - Emails
 * - Teams Messages
 * 
 * Ensures consistent, clean, structured output
 * Only shows: name/subject, time, topic, overview
 * 
 * ============================================================
 */

/**
 * Format calendar events for user display
 * Shows: time, subject, attendees, join link
 * 
 * @param {Array} events - Raw events from Graph API
 * @param {String} userTimeZone - User's timezone (e.g., 'Asia/Singapore')
 * @returns {Array} Formatted events
 */
function formatCalendarEvents(events, userTimeZone = 'UTC') {
  if (!events || !Array.isArray(events)) return [];

  return events.map(event => {
    try {
      const startTime = new Date(event.start?.dateTime || event.startDateTime);
      const endTime = new Date(event.end?.dateTime || event.endDateTime);

      // Format times in user's timezone
      const formattedStart = formatDateTime(startTime, userTimeZone);
      const formattedEnd = formatDateTime(endTime, userTimeZone);
      const duration = calculateDuration(startTime, endTime);

      return {
        id: event.id,
        subject: event.subject,
        time: `${formattedStart} - ${formattedEnd}`,
        duration: duration,
        attendees: (event.attendees || []).map(a => a.emailAddress?.name || 'Unknown').join(', '),
        attendeeCount: event.attendees?.length || 0,
        location: event.location?.displayName || 'No location',
        isTeamsMeeting: event.isOnlineMeeting || false,
        joinUrl: event.onlineMeeting?.joinUrl || null,
        overview: event.bodyPreview || event.body?.content?.substring(0, 100) || 'No description'
      };
    } catch (err) {
      console.error('Error formatting calendar event:', err);
      return null;
    }
  }).filter(e => e !== null);
}

/**
 * Format files for user display
 * Shows: name, modified date, size, preview link
 * 
 * @param {Array} files - Raw files from Graph API
 * @param {String} userTimeZone - User's timezone
 * @returns {Array} Formatted files
 */
function formatFiles(files, userTimeZone = 'UTC') {
  if (!files || !Array.isArray(files)) return [];

  return files.map(file => {
    try {
      const modifiedTime = new Date(file.lastModifiedDateTime);
      const formattedModified = formatDateTime(modifiedTime, userTimeZone);
      const fileSize = formatFileSize(file.size || 0);

      return {
        id: file.id,
        name: file.name,
        type: getFileType(file.name),
        modifiedDate: formattedModified,
        size: fileSize,
        modifiedBy: file.lastModifiedBy?.user?.displayName || 'Unknown',
        webUrl: file.webUrl || null,
        overview: `${fileSize} • Modified by ${file.lastModifiedBy?.user?.displayName || 'Unknown'}`
      };
    } catch (err) {
      console.error('Error formatting file:', err);
      return null;
    }
  }).filter(f => f !== null);
}

/**
 * Format emails for user display
 * Shows: sender, subject, received time, preview
 * 
 * @param {Array} emails - Raw emails from Graph API
 * @param {String} userTimeZone - User's timezone
 * @returns {Array} Formatted emails
 */
function formatEmails(emails, userTimeZone = 'UTC') {
  if (!emails || !Array.isArray(emails)) return [];

  return emails.map(email => {
    try {
      const receivedTime = new Date(email.receivedDateTime);
      const formattedReceived = formatDateTime(receivedTime, userTimeZone);

      return {
        id: email.id,
        from: email.from?.emailAddress?.name || email.from?.emailAddress?.address || 'Unknown',
        subject: email.subject,
        receivedDate: formattedReceived,
        preview: email.bodyPreview || email.body?.content?.substring(0, 100) || 'No content',
        hasAttachments: email.hasAttachments,
        attachmentCount: email.hasAttachments ? 'Has attachments' : 'No attachments',
        overview: `From ${email.from?.emailAddress?.name || 'Unknown'} • ${formattedReceived}`
      };
    } catch (err) {
      console.error('Error formatting email:', err);
      return null;
    }
  }).filter(e => e !== null);
}

/**
 * Format Teams messages for user display
 * Shows: sender, time, message preview
 * 
 * @param {Array} messages - Raw messages from Graph API
 * @param {String} userTimeZone - User's timezone
 * @returns {Array} Formatted messages
 */
function formatTeamsMessages(messages, userTimeZone = 'UTC') {
  if (!messages || !Array.isArray(messages)) return [];

  return messages.map(msg => {
    try {
      const sentTime = new Date(msg.createdDateTime);
      const formattedTime = formatDateTime(sentTime, userTimeZone);

      return {
        id: msg.id,
        sender: msg.from?.user?.displayName || 'Unknown',
        time: formattedTime,
        content: msg.body?.content?.substring(0, 100) || 'Empty message',
        overview: `${msg.from?.user?.displayName || 'Unknown'} • ${formattedTime}`
      };
    } catch (err) {
      console.error('Error formatting Teams message:', err);
      return null;
    }
  }).filter(m => m !== null);
}

/**
 * Format action preview for confirmation
 * Used for: email preview, Teams message preview
 * 
 * @param {Object} action - Action object {type, recipient/recipientName, subject, body}
 * @returns {Object} Formatted preview
 */
function formatActionPreview(action) {
  if (!action || typeof action !== 'object') return null;

  switch (action.type) {
    case 'send_email':
      return {
        type: 'email',
        title: 'Email Preview',
        recipient: action.recipient || action.recipientName,
        subject: action.subject,
        body: action.body,
        cc: action.cc_recipients?.join(', ') || 'None',
        summary: `Send email to ${action.recipient || action.recipientName} with subject: "${action.subject}"`
      };

    case 'send_teams_message':
      return {
        type: 'teams_message',
        title: 'Teams Message Preview',
        recipient: action.recipient || action.recipientName,
        message: action.message,
        summary: `Send Teams message to ${action.recipient || action.recipientName}`
      };

    default:
      return null;
  }
}

/**
 * ============================================================
 * HELPER FUNCTIONS (Internal)
 * ============================================================
 */

/**
 * Format date and time in user's timezone
 * @param {Date} date - Date to format
 * @param {String} timeZone - Timezone (e.g., 'Asia/Singapore')
 * @returns {String} Formatted time like "2:30 PM, Dec 17"
 */
function formatDateTime(date, timeZone = 'UTC') {
  if (!(date instanceof Date) || isNaN(date)) {
    return 'Invalid date';
  }

  try {
    return date.toLocaleString('en-US', {
      timeZone: timeZone,
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });
  } catch (err) {
    console.warn(`Invalid timezone: ${timeZone}, falling back to UTC`);
    return date.toLocaleString('en-US', {
      month: 'short',
      day: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    });
  }
}

/**
 * Calculate duration between two times
 * @param {Date} startTime 
 * @param {Date} endTime 
 * @returns {String} Duration like "1h 30m"
 */
function calculateDuration(startTime, endTime) {
  if (!(startTime instanceof Date) || !(endTime instanceof Date)) {
    return 'Unknown duration';
  }

  const diffMs = endTime - startTime;
  const diffMins = Math.round(diffMs / 60000);
  const hours = Math.floor(diffMins / 60);
  const mins = diffMins % 60;

  if (hours === 0) return `${mins}m`;
  if (mins === 0) return `${hours}h`;
  return `${hours}h ${mins}m`;
}

/**
 * Format file size in human-readable format
 * @param {Number} bytes 
 * @returns {String} Formatted size like "2.5 MB"
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + ' ' + sizes[i];
}

/**
 * Get file type icon/label
 * @param {String} filename 
 * @returns {String} File type
 */
function getFileType(filename) {
  if (!filename) return 'File';
  const ext = filename.split('.').pop().toLowerCase();
  const types = {
    'docx': 'Word Document',
    'doc': 'Word Document',
    'xlsx': 'Excel Sheet',
    'xls': 'Excel Sheet',
    'pptx': 'PowerPoint',
    'ppt': 'PowerPoint',
    'pdf': 'PDF Document',
    'txt': 'Text File',
    'jpg': 'Image',
    'jpeg': 'Image',
    'png': 'Image',
    'gif': 'Image'
  };
  return types[ext] || 'File';
}

// ============================================================
// EXPORTS
// ============================================================
module.exports = {
  formatCalendarEvents,
  formatFiles,
  formatEmails,
  formatTeamsMessages,
  formatActionPreview,
  formatDateTime,
  calculateDuration,
  formatFileSize,
  getFileType
};
