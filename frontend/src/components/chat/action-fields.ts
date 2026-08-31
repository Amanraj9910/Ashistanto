// Shared field logic for action previews, so the card UI and the edit payload can't drift.
//
// The wire shape comes from formatPreviewForDisplay() in ../../../action-preview.js:
//   { actionId, title, type, details, editable, status }
// `details` keys are DISPLAY names; `editable` holds BACKEND field names. editKeyFor() bridges
// them, and the backend rejects any key that isn't in its editableFields list.

/** `details` carries a synthetic one-line summary that must not render as a field row. */
export const HIDDEN_FIELDS = new Set(['preview']);

export const displayLabels: Record<string, string> = {
  to: 'To',
  recipient: 'Recipient',
  subject: 'Subject',
  originalSubject: 'Original Subject',
  body: 'Body',
  message: 'Message',
  cc: 'CC',
  attendees: 'Attendees',
  startTime: 'Start',
  endTime: 'End',
  isTeams: 'Teams Meeting',
  employee: 'Employee',
  leaveType: 'Leave Type',
  startDate: 'Start Date',
  status: 'Status'
};

/** Maps a display key in `details` to the backend field name listed in `editable`. */
export function editKeyFor(field: string) {
  const mapping: Record<string, string> = {
    to: 'recipientName',
    recipient: 'recipientName',
    // Without this, editing CC posts key "cc" and editPendingAction() rejects the whole request
    // with 'Field "cc" cannot be edited'.
    cc: 'ccRecipients',
    attendees: 'attendeeNames',
    isTeams: 'isTeamsMeeting'
  };
  return mapping[field] || field;
}

export function isFieldEditable(field: string, editable: string[]) {
  return editable.includes(editKeyFor(field)) || editable.includes(field);
}

export function isMultiline(field: string) {
  return field === 'body' || field === 'message';
}

export function normalizeValue(value: unknown) {
  if (Array.isArray(value)) return value.join(', ');
  if (typeof value === 'boolean') return value ? 'Yes' : 'No';
  if (value === null || value === undefined || value === '') return '-';
  return String(value);
}

/**
 * Converts an edited display string into the type the backend expects.
 * The list fields arrive comma-joined (or as the literal 'None' when empty) but must be sent
 * back as arrays, otherwise the Graph call receives a single bogus recipient named "a, b".
 */
export function normalizeEditValue(editKey: string, raw: string): unknown {
  if (editKey === 'ccRecipients' || editKey === 'attendeeNames') {
    const trimmed = raw.trim();
    if (!trimmed || trimmed === 'None' || trimmed === '-') return [];
    return trimmed.split(',').map((part) => part.trim()).filter(Boolean);
  }
  if (editKey === 'isTeamsMeeting') return /^(yes|true|on)$/i.test(raw.trim());
  return raw;
}

/** Builds the request payload, dropping untouched fields and converting list types. */
export function buildEditPayload(edits: Record<string, string>): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(edits).map(([editKey, raw]) => [editKey, normalizeEditValue(editKey, raw)])
  );
}
