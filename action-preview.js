/**
 * ============================================================
 * 🔍 ACTION PREVIEW & CONFIRMATION MODULE
 * ============================================================
 * 
 * Handles user confirmation workflow for actions:
 * - Email sending
 * - Teams messages
 * - Calendar invites
 * 
 * Flow:
 * 1. Agent determines action needed
 * 2. System creates preview
 * 3. User gets preview + confirmation request
 * 4. User can: confirm (proceed), edit (get edit options), or cancel
 * 5. System executes based on user choice
 * 
 * No hardcoding - configuration-driven
 * ============================================================
 */

const formatters = require('./formatters');

const { pendingActionsStore } = require('./storage');

// Configuration for confirmation workflows
// IMPORTANT: Field names must match the data passed from agent-tools.js
const CONFIRMATION_CONFIG = {
  send_email: {
    title: '📧 Email Preview',
    requiresConfirmation: true,
    editableFields: ['recipientName', 'subject', 'body', 'ccRecipients'],
    displayFields: ['recipientName', 'subject', 'body', 'ccRecipients']
  },
  send_teams_message: {
    title: '💬 Teams Message Preview',
    requiresConfirmation: true,
    editableFields: ['recipientName', 'message'],
    displayFields: ['recipientName', 'message']
  },
  create_calendar_event: {
    title: '📅 Meeting Preview',
    requiresConfirmation: true,
    editableFields: ['subject', 'attendeeNames', 'startTime', 'endTime'],
    displayFields: ['subject', 'attendeeNames', 'startTime', 'endTime', 'isTeamsMeeting']
  },
  update_calendar_event: {
    title: '📅 Update Meeting Preview',
    requiresConfirmation: true,
    editableFields: ['newSubject', 'newStart', 'newEnd'],
    displayFields: ['subject', 'newSubject', 'newStart', 'newEnd', 'attendees']
  },
  delete_sent_email: {
    title: '🗑️ Delete Email Confirmation',
    requiresConfirmation: true,
    editableFields: [],
    displayFields: ['subject', 'recipient', 'sentDate', 'action']
  },
  delete_inbox_email: {
    title: '🗑️ Delete Inbox Email Confirmation',
    requiresConfirmation: true,
    editableFields: [],
    displayFields: ['subject', 'sender', 'receivedDate', 'action']
  },
  delete_teams_message: {
    title: '🗑️ Delete Teams Message Confirmation',
    requiresConfirmation: true,
    editableFields: [],
    displayFields: ['messageContent', 'sentDate', 'action']
  }
};

/**
 * Create an action preview for user confirmation
 * 
 * @param {String} actionType - Type of action (send_email, send_teams_message, etc.)
 * @param {Object} actionData - Data for the action
 * @param {Object} validatedRecipientData - Optional pre-validated recipient data to cache
 * @returns {Object} Preview object with unique actionId
 */
async function createActionPreview(actionType, actionData, validatedRecipientData = null) {
  // Validate action type
  if (!CONFIRMATION_CONFIG[actionType]) {
    throw new Error(`Unknown action type: ${actionType}`);
  }

  const config = CONFIRMATION_CONFIG[actionType];
  const actionId = generateActionId();

  // Create preview object
  const preview = {
    actionId: actionId,
    actionType: actionType,
    title: config.title,
    timestamp: new Date().toISOString(),
    requiresConfirmation: config.requiresConfirmation,
    data: filterActionData(actionData, config.displayFields),
    editableFields: config.editableFields,
    status: 'pending' // pending, confirmed, edited, cancelled
  };

  // Store in pending actions
  await pendingActionsStore.set(actionId, {
    ...preview,
    originalData: actionData, // Keep original for reference
    editedData: null,
    validatedRecipientData: validatedRecipientData // Cache validated recipient data for fast execution
  });

  console.log(`✓ Action preview created: ${actionId} (${actionType})`);
  if (validatedRecipientData) {
    console.log(`  ✓ Cached validated recipient data for fast execution`);
  }

  return formatPreviewForDisplay(preview);
}

/**
 * Get a pending action by ID
 * @param {String} actionId 
 * @returns {Object} Pending action or null
 */
async function getPendingAction(actionId) {
  return await pendingActionsStore.get(actionId) || null;
}

/**
 * Confirm a pending action
 * Moves action to ready-for-execution status
 * 
 * @param {String} actionId 
 * @param {Object} userConfirmation - {confirmed: true/false}
 * @returns {Object} Confirmation result
 */
async function confirmAction(actionId, userConfirmation = {}) {
  const action = await getPendingAction(actionId);

  if (!action) {
    return {
      success: false,
      error: 'Action not found or already processed'
    };
  }

  if (userConfirmation.confirmed === false) {
    // User cancelled
    await pendingActionsStore.delete(actionId);
    console.log(`✗ Action cancelled: ${actionId}`);
    return {
      success: true,
      status: 'cancelled',
      message: 'Action cancelled by user'
    };
  }

  // User confirmed
  action.status = 'confirmed';
  action.confirmedAt = new Date().toISOString();

  console.log(`✓ Action confirmed: ${actionId}`);

  // Save updated action back to storage
  await pendingActionsStore.set(actionId, action);

  return {
    success: true,
    status: 'confirmed',
    actionId: actionId,
    actionType: action.actionType,
    data: action.editedData || action.originalData
  };
}

/**
 * Edit a pending action before confirmation
 * 
 * @param {String} actionId 
 * @param {Object} edits - Fields to edit
 * @returns {Object} Updated action preview
 */
async function editPendingAction(actionId, edits = {}) {
  const action = await getPendingAction(actionId);

  if (!action) {
    return {
      success: false,
      error: 'Action not found'
    };
  }

  const config = CONFIRMATION_CONFIG[action.actionType];

  // Validate that only editable fields are being changed
  for (const field of Object.keys(edits)) {
    if (!config.editableFields.includes(field)) {
      return {
        success: false,
        error: `Field "${field}" cannot be edited for this action`
      };
    }
  }

  // Apply edits
  action.editedData = action.editedData || { ...action.originalData };
  Object.assign(action.editedData, edits);
  action.status = 'edited';
  action.editedAt = new Date().toISOString();

  console.log(`✓ Action edited: ${actionId}`, edits);

  // Save updated action back to storage
  await pendingActionsStore.set(actionId, action);

  return {
    success: true,
    actionId: actionId,
    changes: edits,
    updatedPreview: formatPreviewForDisplay({
      ...action,
      data: filterActionData(action.editedData, config.displayFields)
    })
  };
}

/**
 * Get action for execution
 * Returns either edited data (if edited) or original data
 * 
 * @param {String} actionId 
 * @returns {Object} Data ready for execution
 */
async function getActionForExecution(actionId) {
  const action = await getPendingAction(actionId);

  if (!action) return null;
  if (action.status !== 'confirmed') return null;

  return action.editedData || action.originalData;
}

/**
 * Clear/remove a processed action
 * @param {String} actionId 
 */
async function clearAction(actionId) {
  if (await pendingActionsStore.has(actionId)) {
    await pendingActionsStore.delete(actionId);
    console.log(`✓ Action cleared: ${actionId}`);
  }
}

/**
 * Format preview for display to user (clean, readable format)
 * @param {Object} preview 
 * @returns {Object} Display-friendly preview
 */
function formatPreviewForDisplay(preview) {
  switch (preview.actionType) {
    case 'send_email':
      return {
        actionId: preview.actionId,
        title: preview.title,
        type: 'email',
        details: {
          to: preview.data.recipientName || preview.data.recipient,
          subject: preview.data.subject,
          body: preview.data.body,
          cc: (preview.data.cc_recipients || preview.data.ccRecipients)?.length > 0 ? (preview.data.cc_recipients || preview.data.ccRecipients).join(', ') : 'None',
          preview: `Email to ${preview.data.recipientName || preview.data.recipient}`
        },
        editable: preview.editableFields,
        status: preview.status
      };

    case 'send_teams_message':
      return {
        actionId: preview.actionId,
        title: preview.title,
        type: 'teams',
        details: {
          recipient: preview.data.recipientName || 'Unknown recipient',
          message: preview.data.message,
          preview: `Teams message to ${preview.data.recipientName || 'Unknown'}`
        },
        editable: preview.editableFields,
        status: preview.status
      };

    case 'create_calendar_event':
      return {
        actionId: preview.actionId,
        title: preview.title,
        type: 'meeting',
        details: {
          subject: preview.data.subject,
          attendees: preview.data.attendeeNames?.join(', ') || 'No attendees',
          startTime: preview.data.startTime || preview.data.start,
          endTime: preview.data.endTime || preview.data.end,
          isTeams: preview.data.isTeamsMeeting || false,
          preview: `Meeting: ${preview.data.subject}`
        },
        editable: preview.editableFields,
        status: preview.status
      };

    case 'update_calendar_event':
      return {
        actionId: preview.actionId,
        title: preview.title,
        type: 'meeting',
        details: {
          originalSubject: preview.data.subject,
          subject: preview.data.newSubject || preview.data.subject,
          attendees: preview.data.attendees || 'No attendees',
          startTime: preview.data.newStart,
          endTime: preview.data.newEnd,
          preview: `Update Meeting: ${preview.data.subject}`
        },
        editable: preview.editableFields,
        status: preview.status
      };

    case 'delete_sent_email':
      return {
        actionId: preview.actionId,
        title: preview.title,
        type: 'delete_email',
        details: {
          subject: preview.data.subject || 'Unknown subject',
          recipient: preview.data.recipient || 'Unknown recipient',
          sentDate: preview.data.sentDate || 'Unknown date',
          action: 'Delete this email',
          preview: `Delete sent email: "${preview.data.subject}"`
        },
        editable: preview.editableFields,
        status: preview.status,
        isDeletion: true
      };

    case 'delete_inbox_email':
      return {
        actionId: preview.actionId,
        title: preview.title,
        type: 'delete_email',
        details: {
          subject: preview.data.subject || 'Unknown subject',
          sender: preview.data.sender || 'Unknown sender',
          receivedDate: preview.data.receivedDate || 'Unknown date',
          action: 'Delete this email',
          preview: `Delete inbox email: "${preview.data.subject}"`
        },
        editable: preview.editableFields,
        status: preview.status,
        isDeletion: true
      };

    case 'delete_teams_message':
      return {
        actionId: preview.actionId,
        title: preview.title,
        type: 'delete_teams',
        details: {
          messageContent: preview.data.messageContent || 'Recent message',
          sentDate: preview.data.sentDate || 'Recently',
          action: 'Delete this Teams message',
          preview: `Delete Teams message: "${preview.data.messageContent?.substring(0, 50)}..."`
        },
        editable: preview.editableFields,
        status: preview.status,
        isDeletion: true
      };

    default:
      return preview;
  }
}

/**
 * ============================================================
 * HELPER FUNCTIONS (Internal)
 * ============================================================
 */

/**
 * Generate unique action ID
 * @returns {String}
 */
function generateActionId() {
  return `action_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
}

/**
 * Filter action data to show only specified fields
 * @param {Object} data 
 * @param {Array} fields 
 * @returns {Object} Filtered data
 */
function filterActionData(data, fields) {
  const filtered = {};
  for (const field of fields) {
    if (field in data) {
      filtered[field] = data[field];
    }
  }
  return filtered;
}

/**
 * Get all pending actions (for debugging/UI)
 * @returns {Array} List of pending actions
 */
async function getAllPendingActions() {
  const actions = await pendingActionsStore.values();
  return actions.map(action => ({
    actionId: action.actionId,
    actionType: action.actionType,
    status: action.status,
    createdAt: action.timestamp
  }));
}

/**
 * Clean up expired actions (older than 1 hour)
 * Call periodically to prevent memory leaks
 */
async function cleanupExpiredActions() {
  const oneHourAgo = Date.now() - (60 * 60 * 1000);
  let cleaned = 0;

  for (const [actionId, action] of await pendingActionsStore.entries()) {
    const createdTime = new Date(action.timestamp).getTime();
    if (createdTime < oneHourAgo) {
      await pendingActionsStore.delete(actionId);
      cleaned++;
    }
  }

  if (cleaned > 0) {
    console.log(`✓ Cleaned up ${cleaned} expired actions`);
  }

  return cleaned;
}

// ============================================================
// EXPORTS
// ============================================================
module.exports = {
  // Main functions
  createActionPreview,
  getPendingAction,
  confirmAction,
  editPendingAction,
  getActionForExecution,
  clearAction,
  formatPreviewForDisplay,

  // Utilities
  getAllPendingActions,
  cleanupExpiredActions,

  // Config
  CONFIRMATION_CONFIG
};
