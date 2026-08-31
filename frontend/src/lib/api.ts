import type {
  ConfigStatus,
  ConfirmActionPayload,
  ConfirmActionResponse,
  ConversationSummary,
  TextMessagePayload,
  TextMessageResponse,
  UserProfile,
  VoiceMessagePayload,
  VoiceMessageResponse
} from '@/types';
import { mockConfig, mockProcessVoice, mockSendTextMessage, mockUser } from './mock-data';

// Empty base = same origin. The UI is a static export served by the Express app itself, so
// every path below is a plain relative request: no CORS, no proxy, no rewrites.
const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL?.replace(/\/$/, '') || '';
// Frontend deployments are self-contained by default. Set this explicitly to false only when a public backend is ready.
const enableMocks = process.env.NEXT_PUBLIC_ENABLE_MOCKS !== 'false';

export function shouldUseMocks() {
  return enableMocks && !apiBaseUrl;
}

export function getAuthLoginUrl() {
  return shouldUseMocks() ? null : `${apiBaseUrl}/auth/login`;
}

/** Raised when the backend reports the Microsoft session has expired. */
export class SessionExpiredError extends Error {
  constructor(message = 'Your session expired. Please sign in again.') {
    super(message);
    this.name = 'SessionExpiredError';
  }
}

type ExpiryListener = () => void;
const expiryListeners = new Set<ExpiryListener>();

/**
 * Registers a callback invoked once when any request reports an expired session. The chat page
 * uses this to clear local state and bounce to /login.
 */
export function onSessionExpired(listener: ExpiryListener) {
  expiryListeners.add(listener);
  return () => expiryListeners.delete(listener);
}

function notifySessionExpired() {
  expiryListeners.forEach((listener) => {
    try { listener(); } catch { /* a failing listener must not mask the original error */ }
  });
}

/**
 * Reads the response body exactly once, then decides.
 *
 * The pre-migration UI had a separate checkTokenExpired() that consumed the body before the
 * caller's own response.json(), so any 401 that was NOT a token expiry surfaced as
 * "body stream already read" instead of the real error.
 */
async function handle<T>(response: Response): Promise<T> {
  const contentType = response.headers.get('content-type') || '';
  const isJson = contentType.includes('application/json');
  const body: unknown = isJson ? await response.json().catch(() => null) : await response.text().catch(() => '');

  if (response.ok) return body as T;

  const record = (body && typeof body === 'object' ? body : {}) as Record<string, unknown>;
  const message = typeof record.error === 'string' ? record.error : response.statusText;

  // Any 401 reaching here means the session is unusable, so force a re-login.
  //
  // Checking only the isTokenExpired/requiresLogin flags (as the previous UI did) misses the
  // most common case: /api/text-message answers a dead session with a bare
  // {error:'Invalid or expired session'} and no flags, which would surface as an error banner
  // on a session that can never recover.
  //
  // /api/user-photo is deliberately NOT routed through here -- it 401s when the photo simply
  // isn't available, and a missing avatar must never sign anyone out.
  if (response.status === 401) {
    notifySessionExpired();
    throw new SessionExpiredError(message || undefined);
  }

  throw new Error(message || `Request failed with ${response.status}`);
}

async function requestJson<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${apiBaseUrl}${path}`, {
    ...init,
    headers: { 'Content-Type': 'application/json', ...(init?.headers || {}) }
  });
  return handle<T>(response);
}

export async function getConfig(): Promise<ConfigStatus> {
  if (shouldUseMocks()) return mockConfig;
  return requestJson<ConfigStatus>('/api/config');
}

export async function validateSession(sessionId: string): Promise<boolean> {
  if (shouldUseMocks()) return true;

  const response = await fetch(`${apiBaseUrl}/auth/session-token/${encodeURIComponent(sessionId)}`);
  return response.ok;
}

export async function getConversations(): Promise<ConversationSummary[]> {
  if (shouldUseMocks()) return [];
  try {
    return await requestJson<ConversationSummary[]>('/api/conversations');
  } catch {
    // The RECENTS list is decorative; never let it break the page.
    return [];
  }
}

export async function getUserProfile(sessionId: string | null): Promise<UserProfile> {
  if (shouldUseMocks() || !sessionId) return mockUser;

  const profile = await requestJson<Partial<UserProfile>>(`/api/user-profile?sessionId=${encodeURIComponent(sessionId)}`);

  return {
    displayName: profile.displayName || profile.firstName || 'User',
    firstName: profile.firstName || profile.displayName?.split(' ')[0] || 'User',
    email: profile.email || 'user@company.onmicrosoft.com',
    role: profile.role || 'Director of Ops'
  };
}

export async function getUserPhotoUrl(sessionId: string | null): Promise<string | null> {
  if (shouldUseMocks() || !sessionId) return null;

  const response = await fetch(`${apiBaseUrl}/api/user-photo?sessionId=${encodeURIComponent(sessionId)}`);
  if (!response.ok) return null;

  const blob = await response.blob();
  if (!blob.size) return null;

  return URL.createObjectURL(blob);
}

export async function sendTextMessage(payload: TextMessagePayload): Promise<TextMessageResponse> {
  if (shouldUseMocks()) return mockSendTextMessage(payload.text, payload.sessionId);

  return requestJson<TextMessageResponse>('/api/text-message', {
    method: 'POST',
    body: JSON.stringify(payload)
  });
}

export async function processVoiceMessage(payload: VoiceMessagePayload): Promise<VoiceMessageResponse> {
  if (shouldUseMocks()) return mockProcessVoice(payload.sessionId);

  const formData = new FormData();
  formData.append('audio', payload.audio, 'recording.webm');
  formData.append('sessionId', payload.sessionId);
  formData.append('language', payload.language);
  formData.append('accent', payload.accent);

  // No Content-Type header: the browser must set the multipart boundary itself.
  const response = await fetch(`${apiBaseUrl}/api/process-voice`, { method: 'POST', body: formData });
  return handle<VoiceMessageResponse>(response);
}

export async function confirmAction(payload: ConfirmActionPayload): Promise<ConfirmActionResponse> {
  if (shouldUseMocks()) {
    await new Promise((resolve) => setTimeout(resolve, 450));
    return {
      success: true,
      message: payload.userChoice === 'confirm' ? 'Action completed successfully.' : 'Action cancelled.'
    };
  }

  return requestJson<ConfirmActionResponse>('/api/confirm-action', {
    method: 'POST',
    body: JSON.stringify(payload)
  });
}

export async function clearConversation(sessionId: string | null): Promise<void> {
  if (shouldUseMocks() || !sessionId) return;

  await requestJson('/api/clear-session', {
    method: 'POST',
    body: JSON.stringify({ sessionId })
  });
}

export async function logoutSession(sessionId: string | null): Promise<void> {
  if (shouldUseMocks() || !sessionId) return;

  await requestJson('/auth/logout', {
    method: 'POST',
    body: JSON.stringify({ sessionId })
  });
}
