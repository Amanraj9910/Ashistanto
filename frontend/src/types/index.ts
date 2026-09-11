export type ThemeMode = 'light' | 'dark';

export type VoiceAccent = 'american' | 'british' | 'japanese';

export type UserProfile = {
  displayName: string;
  firstName: string;
  email: string;
  role: string;
  photoUrl?: string;
};

export type ConfigStatus = {
  configured: boolean;
};

export type ChatMessage = {
  id: string;
  user: string;
  assistant: string;
  timestamp: string;
  confirmed?: boolean;
};

export type ActionKind =
  | 'email'
  | 'teams'
  | 'meeting'
  | 'leave'
  | 'expense'
  | 'delete_email'
  | 'delete_teams';

export type ActionPreview = {
  actionId: string;
  title: string;
  type: ActionKind;
  status: 'pending' | 'confirmed' | 'cancelled';
  details: Record<string, string | number | boolean | string[] | null | undefined>;
  editable: string[];
};

export type TextMessagePayload = {
  text: string;
  sessionId: string;
  language: string;
  accent: VoiceAccent;
};

export type TextMessageResponse = {
  success: boolean;
  response: string;
  sessionId: string;
};

export type VoiceMessagePayload = {
  audio: Blob;
  sessionId: string;
  language: string;
  accent: VoiceAccent;
};

export type VoiceMessageResponse = {
  transcript: string;
  agentResponse: string;
  audioData?: string;
  sessionId: string;
};

export type ConfirmActionPayload = {
  sessionId: string;
  actionId: string;
  userChoice: 'confirm' | 'cancel';
  edits?: Record<string, unknown> | null;
};

export type ConfirmActionResponse = {
  success: boolean;
  message: string;
  result?: unknown;
};
