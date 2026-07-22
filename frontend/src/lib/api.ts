import type {
  ConfigStatus,
  ConfirmActionPayload,
  ConfirmActionResponse,
  TextMessagePayload,
  TextMessageResponse,
  UserProfile,
  VoiceMessagePayload,
  VoiceMessageResponse
} from '@/types';
import { mockConfig, mockProcessVoice, mockSendTextMessage, mockUser } from './mock-data';

const apiBaseUrl = process.env.NEXT_PUBLIC_API_URL?.replace(/\/$/, '') || '';
const enableMocks = process.env.NEXT_PUBLIC_ENABLE_MOCKS === 'true';

export function shouldUseMocks() {
  return enableMocks && !apiBaseUrl;
}

export function getAuthLoginUrl() {
  return shouldUseMocks() ? null : `${apiBaseUrl}/auth/login`;
}

async function requestJson<T>(path: string, init?: RequestInit): Promise<T> {
  const response = await fetch(`${apiBaseUrl}${path}`, {
    ...init,
    headers: {
      'Content-Type': 'application/json',
      ...(init?.headers || {})
    }
  });

  const data = await response.json();
  if (!response.ok) {
    throw new Error(data.error || response.statusText);
  }

  return data;
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

  const response = await fetch(`${apiBaseUrl}/api/process-voice`, {
    method: 'POST',
    body: formData
  });

  const data = await response.json();
  if (!response.ok) throw new Error(data.error || response.statusText);

  return data;
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
