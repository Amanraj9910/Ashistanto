export function cn(...classes: Array<string | false | null | undefined>) {
  return classes.filter(Boolean).join(' ');
}

export function createId(prefix = 'id') {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

export function formatClock(date = new Date()) {
  return date.toLocaleTimeString('en-US', {
    hour: '2-digit',
    minute: '2-digit'
  });
}

export function greeting(date = new Date()) {
  const hour = date.getHours();
  if (hour < 12) return 'Good morning';
  if (hour < 18) return 'Good afternoon';
  return 'Good evening';
}

// Mirrors the pre-migration formatAIResponse(): unwraps the envelope objects the agent
// returns before stripping markdown emphasis, so a tool result renders as a sentence rather
// than a JSON blob.
function unwrapEnvelope(value: unknown): unknown {
  if (typeof value === 'object' && value !== null) {
    const record = value as Record<string, unknown>;
    if (record.success !== undefined) {
      return record.success
        ? `✅ ${record.message || 'Action completed successfully'}`
        : `❌ ${record.error || 'Action failed'}`;
    }
    if (record.type === 'action_preview') {
      return record.message || 'Please review and confirm the action above.';
    }
    return JSON.stringify(value, null, 2);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    const looksLikeJson = trimmed.startsWith('{') || trimmed.startsWith('[');
    if (looksLikeJson && (trimmed.includes('"type"') || trimmed.includes('"success"'))) {
      try {
        return unwrapEnvelope(JSON.parse(trimmed));
      } catch {
        // Not valid JSON after all - fall through and treat it as plain text.
      }
    }
  }

  return value;
}

export function cleanAssistantText(value: unknown) {
  if (value === null || value === undefined) return '';

  const unwrapped = unwrapEnvelope(value);
  if (typeof unwrapped === 'object') {
    return JSON.stringify(unwrapped, null, 2);
  }

  return String(unwrapped)
    .replace(/\*\*([^*]+)\*\*/g, '$1')
    .replace(/\*([^*]+)\*/g, '$1')
    .replace(/__([^_]+)__/g, '$1')
    .replace(/_([^_]+)_/g, '$1')
    .replace(/\n{3,}/g, '\n\n');
}
