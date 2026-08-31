'use client';

import { useMutation } from '@tanstack/react-query';
import { BriefcaseBusiness, CheckCircle2, FileText, Keyboard, LogIn, Mic, ReceiptText, Send, Square, StopCircle } from 'lucide-react';
import { FormEvent, useEffect, useRef, useState } from 'react';
import { clearConversation, confirmAction, processVoiceMessage, sendTextMessage } from '@/lib/api';
import { cleanAssistantText, cn } from '@/lib/utils';
import { useAppStore } from '@/store/use-app-store';
import type { ActionPreview, VoiceAccent } from '@/types';
import { ActionPreviewCard } from './action-preview-card';

const professionalGreetings = [
  'How may I assist you today?',
  'What can I help you accomplish?',
  'How can I support your work today?',
  'What would you like to take care of?',
  'How can I make your day more efficient?',
  'Which task should we handle first?'
];

const accentOptions: Array<{ value: VoiceAccent; label: string; language: string }> = [
  { value: 'american', label: 'American', language: 'en-US' },
  { value: 'british', label: 'British', language: 'en-GB' },
  { value: 'japanese', label: 'Japanese', language: 'en-US' }
];

function parseAssistantResponse(response: string): { message: string; action?: ActionPreview } {
  try {
    const parsed = JSON.parse(response);
    if (parsed?.type === 'action_preview') {
      return {
        message: parsed.message || 'Please review and confirm this action.',
        action: parsed.preview
      };
    }
  } catch {
    // Plain responses are expected for normal chat turns.
  }

  return { message: cleanAssistantText(response) };
}

function TypingDots() {
  return (
    <div className="flex items-center gap-1 px-4 py-3">
      {[0, 1, 2].map((i) => (
        <span
          key={i}
          className="h-1.5 w-1.5 rounded-full bg-slate-400 dark:bg-slate-500"
          style={{ animation: `dot-pulse 1.2s ease-in-out ${i * 0.15}s infinite` }}
        />
      ))}
    </div>
  );
}

function VoiceRecordingOverlay({ onStop, onSwitchToTyping, accent, accentOptions, onAccentChange }: {
  onStop: () => void;
  onSwitchToTyping: () => void;
  accent: VoiceAccent;
  accentOptions: Array<{ value: VoiceAccent; label: string }>;
  onAccentChange: (v: VoiceAccent) => void;
}) {
  return (
    <div className="voice-overlay">
      <div className="voice-mic-container">
        <span className="voice-pulse-ring" />
        <span className="voice-pulse-ring" />
        <span className="voice-pulse-ring" />
        <button type="button" onClick={onStop} className="voice-mic-btn" aria-label="Stop recording">
          <Mic className="h-8 w-8" />
        </button>
      </div>
      <div className="text-center">
        <p className="voice-status-label">Listening...</p>
      </div>
      <div className="voice-footer-controls">
        <select
          value={accent}
          onChange={(e) => onAccentChange(e.target.value as VoiceAccent)}
          className="rounded-none border border-slate-200 bg-white px-2.5 py-1.5 text-[11px] font-medium text-slate-500 outline-none dark:border-slate-800 dark:bg-slate-950 dark:text-slate-400"
        >
          {accentOptions.map((opt) => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
        </select>
        <button type="button" onClick={onSwitchToTyping} className="voice-switch-btn">
          Switch to typing <Keyboard className="h-3.5 w-3.5" />
        </button>
      </div>
    </div>
  );
}

function VoiceProcessingOverlay({ accent, accentOptions, onAccentChange }: {
  accent: VoiceAccent;
  accentOptions: Array<{ value: VoiceAccent; label: string }>;
  onAccentChange: (v: VoiceAccent) => void;
}) {
  return (
    <div className="voice-overlay">
      <div className="voice-mic-container">
        <span className="voice-pulse-ring" />
        <span className="voice-pulse-ring" />
        <span className="voice-pulse-ring" />
        <div className="voice-mic-btn processing" aria-label="Processing voice">
          <Mic className="h-8 w-8" />
        </div>
      </div>
      <div className="text-center">
        <p className="voice-status-label">Preparing your email...</p>
        <p className="voice-status-sub">This may take a few seconds.</p>
      </div>
      <div className="voice-processing-dots">
        <span /><span /><span />
      </div>
      <div className="voice-footer-controls">
        <select
          value={accent}
          onChange={(e) => onAccentChange(e.target.value as VoiceAccent)}
          className="rounded-none border border-slate-200 bg-white px-2.5 py-1.5 text-[11px] font-medium text-slate-500 outline-none dark:border-slate-800 dark:bg-slate-950 dark:text-slate-400"
        >
          {accentOptions.map((opt) => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
        </select>
      </div>
    </div>
  );
}

type ChatPanelProps = {
  configured: boolean;
  isSignedIn: boolean;
  onLogin: () => void;
};

export function ChatPanel({ configured, isSignedIn, onLogin }: ChatPanelProps) {
  const sessionId = useAppStore((state) => state.sessionId);
  const user = useAppStore((state) => state.user);
  const messages = useAppStore((state) => state.messages);
  const pendingAction = useAppStore((state) => state.pendingAction);
  const accent = useAppStore((state) => state.accent);
  const language = useAppStore((state) => state.language);
  const setAccent = useAppStore((state) => state.setAccent);
  const addMessage = useAppStore((state) => state.addMessage);
  const clearMessages = useAppStore((state) => state.clearMessages);
  const setPendingAction = useAppStore((state) => state.setPendingAction);

  const [input, setInput] = useState('');
  const [error, setError] = useState('');
  const [isRecording, setIsRecording] = useState(false);
  const [isSpeaking, setIsSpeaking] = useState(false);
  const [emptyGreeting, setEmptyGreeting] = useState(professionalGreetings[0]);

  const recorderRef = useRef<MediaRecorder | null>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const chunksRef = useRef<Blob[]>([]);
  const audioRef = useRef<HTMLAudioElement | null>(null);
  const messagesEndRef = useRef<HTMLDivElement | null>(null);

  const textMutation = useMutation({
    mutationFn: sendTextMessage,
    onSuccess: (data, variables) => {
      const parsed = parseAssistantResponse(data.response);
      addMessage({ user: variables.text, assistant: parsed.message });
      setPendingAction(parsed.action || null);
    },
    onError: (mutationError) => {
      setError(mutationError instanceof Error ? mutationError.message : 'Failed to process message.');
    }
  });

  const voiceMutation = useMutation({
    mutationFn: processVoiceMessage,
    onSuccess: async (data) => {
      const parsed = parseAssistantResponse(data.agentResponse);
      addMessage({ user: data.transcript || 'Voice message', assistant: parsed.message });
      setPendingAction(parsed.action || null);

      if (data.audioData && audioRef.current) {
        const audioBlob = new Blob([Uint8Array.from(atob(data.audioData), (char) => char.charCodeAt(0))], {
          type: 'audio/mpeg'
        });
        audioRef.current.src = URL.createObjectURL(audioBlob);
        setIsSpeaking(true);
        await audioRef.current.play();
      }
    },
    onError: (mutationError) => {
      setError(mutationError instanceof Error ? mutationError.message : 'Failed to process voice.');
    }
  });

  const actionMutation = useMutation({
    mutationFn: confirmAction,
    onSuccess: (data, variables) => {
      addMessage({
        user: variables.userChoice === 'confirm' ? '✅ Action confirmed' : 'Action cancelled',
        assistant: data.message || (variables.userChoice === 'confirm' ? 'Action completed successfully.' : 'Action cancelled.'),
        confirmed: variables.userChoice === 'confirm'
      });
      setPendingAction(null);
    },
    onError: (mutationError) => {
      setError(mutationError instanceof Error ? mutationError.message : 'Failed to update action.');
    }
  });

  const isProcessing = textMutation.isPending || voiceMutation.isPending || actionMutation.isPending;  useEffect(() => {
    const storageKey = 'ashistanto-greeting-index';
    const previous = Number.parseInt(localStorage.getItem(storageKey) || '-1', 10);
    const next = Number.isFinite(previous) ? (previous + 1) % professionalGreetings.length : 0;
    localStorage.setItem(storageKey, String(next));
    setEmptyGreeting(professionalGreetings[next]);
  }, []);


  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, pendingAction, isProcessing]);

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const text = input.trim();

    if (!isSignedIn) {
      onLogin();
      return;
    }

    if (!text || !sessionId || isProcessing) return;

    if (pendingAction && /^(confirm|yes|send|proceed)$/i.test(text)) {
      setInput('');
      actionMutation.mutate({
        sessionId,
        actionId: pendingAction.actionId,
        userChoice: 'confirm'
      });
      return;
    }

    if (pendingAction && /^(cancel|no|stop)$/i.test(text)) {
      setInput('');
      actionMutation.mutate({
        sessionId,
        actionId: pendingAction.actionId,
        userChoice: 'cancel'
      });
      return;
    }

    setError('');
    setInput('');
    textMutation.mutate({
      text,
      sessionId,
      language,
      accent
    });
  }

  async function startRecording() {
    if (!isSignedIn) {
      setError('Sign in to use voice recording.');
      return;
    }

    if (!sessionId || isProcessing || !navigator.mediaDevices?.getUserMedia) {
      setError('Microphone recording is not available in this browser.');
      return;
    }

    try {
      setError('');
      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      streamRef.current = stream;

      let mimeType = 'audio/webm;codecs=opus';
      if (!MediaRecorder.isTypeSupported(mimeType)) mimeType = 'audio/webm';

      const recorder = new MediaRecorder(stream, {
        mimeType,
        audioBitsPerSecond: 128000
      });

      recorderRef.current = recorder;
      chunksRef.current = [];

      recorder.ondataavailable = (event) => {
        if (event.data.size > 0) chunksRef.current.push(event.data);
      };

      recorder.onstop = () => {
        const audio = new Blob(chunksRef.current, { type: mimeType });
        streamRef.current?.getTracks().forEach((track) => track.stop());
        streamRef.current = null;
        setIsRecording(false);

        if (audio.size < 1000) {
          setError('Recording is too short. Please speak for at least one second.');
          return;
        }

        voiceMutation.mutate({
          audio,
          sessionId,
          language,
          accent
        });
      };

      recorder.start(100);
      setIsRecording(true);
    } catch {
      setError('Microphone access was denied. Please allow microphone access and try again.');
    }
  }

  function stopRecording() {
    if (recorderRef.current?.state === 'recording') {
      recorderRef.current.stop();
    }
  }

  function stopSpeaking() {
    if (!audioRef.current) return;
    audioRef.current.pause();
    audioRef.current.currentTime = 0;
    setIsSpeaking(false);
  }

  async function handleClear() {
    clearMessages();
    await clearConversation(sessionId);
  }

  return (
    <aside className="dashboard-chat-panel flex min-h-[640px] flex-col overflow-hidden rounded-none border border-slate-200 bg-white shadow-sm dark:border-slate-800 dark:bg-slate-950 xl:h-[calc(100vh-156px)]">
      {/* ─── Header ─── */}
      <div className="chat-panel-header flex items-center justify-between border-b border-slate-100 bg-white px-4 py-3 dark:border-slate-800 dark:bg-slate-950">
        <div className="flex items-center gap-2.5">
          <div className="flex h-8 w-8 items-center justify-center">
            <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" className="h-full w-full object-contain" />
          </div>
          <div>
            <div className="text-sm font-semibold text-slate-900 dark:text-white">Ashistanto</div>
            <div className={cn(
              'text-[11px] font-medium',
              !isSignedIn ? 'text-slate-400' :
              isProcessing ? 'text-amber-500' :
              isRecording ? 'text-red-500' :
              'text-emerald-500'
            )}>
              {!isSignedIn ? 'Preview' : isProcessing ? 'Thinking…' : isRecording ? 'Listening…' : 'Online'}
            </div>
          </div>
        </div>

        <div className="flex items-center gap-1.5">
          {messages.length > 0 && (
            <button
              type="button"
              onClick={handleClear}
              className="rounded-none border border-slate-200 px-2.5 py-1.5 text-[11px] font-medium text-slate-500 hover:bg-slate-50 dark:border-slate-800 dark:text-slate-400 dark:hover:bg-slate-900"
            >
              Clear
            </button>
          )}
          <select
            value={accent}
            onChange={(event) => setAccent(event.target.value as VoiceAccent)}
            className="chat-accent-select rounded-none border border-slate-200 bg-white px-2.5 py-1.5 text-[11px] font-medium text-slate-500 outline-none dark:border-slate-800 dark:bg-slate-950 dark:text-slate-400"
          >
            {accentOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>
      </div>

      {/* ─── Messages ─── */}
      <div className={cn('relative flex-1 space-y-4 bg-slate-50 px-4 py-4 dark:bg-slate-900/50', messages.length === 0 ? 'overflow-hidden' : 'overflow-y-auto') }>
        {/* Voice Recording Overlay */}
        {isRecording && (
          <VoiceRecordingOverlay
            onStop={stopRecording}
            onSwitchToTyping={() => { stopRecording(); }}
            accent={accent}
            accentOptions={accentOptions}
            onAccentChange={setAccent}
          />
        )}
        {/* Voice Processing Overlay */}
        {voiceMutation.isPending && (
          <VoiceProcessingOverlay
            accent={accent}
            accentOptions={accentOptions}
            onAccentChange={setAccent}
          />
        )}
        {messages.length === 0 ? (
          <div className="chat-empty-state">
            <div className="chat-empty-logo"><img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" /></div>
            <p className="chat-empty-greeting">Hi, there <span>👋</span></p>
            <h1>{emptyGreeting}</h1>
            <div className="chat-suggestions">
              {[
  { prompt: 'Help me prepare a leave request', icon: BriefcaseBusiness, tone: 'blue' },
  { prompt: 'Write an expense report', icon: ReceiptText, tone: 'pink' },
  { prompt: 'Draft a meeting summary', icon: FileText, tone: 'amber' }
].map(({ prompt, icon: Icon, tone }) => (
  <button key={prompt} type="button" onClick={() => setInput(prompt)} className={'chat-suggestion-card ' + tone}>
    <span className="chat-suggestion-icon"><Icon className="h-4 w-4" /></span>
    <span>{prompt}</span>
  </button>
))}
            </div>
          </div>
        ) : (
          <div className="flex justify-center"><span className="rounded-none bg-slate-200/70 px-3 py-1 text-[11px] font-medium text-slate-500">Today</span></div>
        )}

        {messages.map((message) => (
          <div key={message.id} className="space-y-3">
            {/* User bubble */}
            <div className="flex items-start justify-end gap-2.5">
              {message.confirmed ? (
                <span className="action-confirmed-badge">
                  <CheckCircle2 className="h-3.5 w-3.5" /> Action confirmed
                </span>
              ) : (
                <div className="max-w-[260px] whitespace-pre-wrap rounded-none bg-slate-200 px-3.5 py-2.5 text-sm leading-relaxed text-slate-700 shadow-sm dark:bg-slate-700 dark:text-slate-100">
                  {message.user}
                </div>
              )}
              {user.photoUrl ? (
                <img
                  src={user.photoUrl}
                  alt={`${user.firstName} profile`}
                  className="h-7 w-7 shrink-0 rounded-full border border-slate-200 object-cover dark:border-slate-700"
                />
              ) : (
                <div className="flex h-7 w-7 shrink-0 items-center justify-center rounded-full border border-slate-200 bg-white text-xs font-semibold text-slate-600 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-300">
                  {user.firstName.charAt(0).toUpperCase()}
                </div>
              )}
            </div>

            {/* Assistant bubble */}
            <div className="flex items-start gap-2.5">
              <div className="flex h-7 w-7 shrink-0 items-center justify-center">
                <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" className="h-full w-full object-contain" />
              </div>
              <div className="max-w-[280px] whitespace-pre-wrap rounded-none border border-slate-100 bg-white px-3.5 py-2.5 text-sm leading-relaxed text-slate-700 shadow-sm dark:border-slate-800 dark:bg-slate-950 dark:text-slate-200">
                {cleanAssistantText(message.assistant)}
              </div>
            </div>
          </div>
        ))}

        {pendingAction && (
          <div className="flex items-start gap-2.5">
            <div className="flex h-7 w-7 shrink-0 items-center justify-center">
              <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" className="h-full w-full object-contain" />
            </div>
            <div className="w-full max-w-[300px]">
              <ActionPreviewCard
                action={pendingAction}
                isSubmitting={actionMutation.isPending}
                onConfirm={(edits) => {
                  if (!sessionId) return;
                  actionMutation.mutate({
                    sessionId,
                    actionId: pendingAction.actionId,
                    userChoice: 'confirm',
                    edits: Object.keys(edits).length > 0 ? edits : null
                  });
                }}
                onCancel={() => {
                  if (!sessionId) return;
                  actionMutation.mutate({
                    sessionId,
                    actionId: pendingAction.actionId,
                    userChoice: 'cancel'
                  });
                }}
              />
            </div>
          </div>
        )}

        {isProcessing && (
          <div className="flex items-start gap-2.5">
            <div className="flex h-7 w-7 shrink-0 items-center justify-center">
              <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" className="h-full w-full object-contain" />
            </div>
            <div className="rounded-none border border-slate-100 bg-white shadow-sm dark:border-slate-800 dark:bg-slate-950">
              <TypingDots />
            </div>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {error && <div className="border-t border-red-100 bg-red-50 px-4 py-2 text-xs font-medium text-red-700 dark:border-red-950 dark:bg-red-950/40 dark:text-red-300">{error}</div>}

      {/* ─── Input area ─── */}
      {!isSignedIn ? (
        <div className="border-t border-slate-100 bg-white p-4 dark:border-slate-800 dark:bg-slate-950">
          <div className="chat-voice-preview">
            <button type="button" onClick={onLogin} className="chat-preview-mic" aria-label="Sign in to use microphone"><Mic className="h-5 w-5" /></button>
            <div><strong>AI conversation</strong><span>Voice input ready</span></div>
            <select value={accent} onChange={(event) => setAccent(event.target.value as VoiceAccent)} aria-label="Voice language">
              {accentOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
            </select>
          </div>
          <button
            type="button"
            onClick={onLogin}
            className="inline-flex w-full items-center justify-center gap-2 rounded-none bg-red-600 px-4 py-2.5 text-sm font-semibold text-white shadow-sm transition hover:bg-red-700"
          >
            <LogIn className="h-4 w-4" />
            Sign in to use Ashistanto
          </button>
          <p className="mt-2.5 text-center text-[11px] leading-4 text-slate-400 dark:text-slate-500">
            Chat and voice are available after signing in with Microsoft.
          </p>
        </div>
      ) : (
        <form onSubmit={handleSubmit} className="border-t border-slate-100 bg-white p-3 dark:border-slate-800 dark:bg-slate-950">
          <div className="flex items-end gap-1.5 rounded-none border border-slate-200 bg-slate-50 px-2.5 py-1.5 dark:border-slate-800 dark:bg-slate-900">
            <button
              type="button"
              onClick={isRecording ? stopRecording : startRecording}
              disabled={isProcessing || !configured}
              className={cn(
                'rounded-none p-2 transition disabled:cursor-not-allowed disabled:opacity-50',
                isRecording ? 'bg-red-600 text-white' : 'text-slate-400 hover:bg-white dark:text-slate-500 dark:hover:bg-slate-800'
              )}
              aria-label={isRecording ? 'Stop recording' : 'Start recording'}
            >
              {isRecording ? <Square className="h-4 w-4 fill-current" /> : <Mic className="h-4 w-4" />}
            </button>

            <textarea
              value={input}
              onChange={(event) => setInput(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === 'Enter' && !event.shiftKey) {
                  event.preventDefault();
                  event.currentTarget.form?.requestSubmit();
                }
              }}
              placeholder={isRecording ? 'Listening…' : 'Type a message…'}
              disabled={isProcessing || !configured}
              rows={1}
              className="max-h-28 min-h-[38px] min-w-0 flex-1 resize-none bg-transparent py-2 text-sm leading-relaxed text-slate-800 outline-none placeholder:text-slate-400 disabled:cursor-not-allowed dark:text-slate-100 dark:placeholder:text-slate-500"
            />

            {isSpeaking && (
              <button
                type="button"
                onClick={stopSpeaking}
                className="inline-flex items-center gap-1 rounded-none border border-amber-200 bg-amber-50 px-2.5 py-1.5 text-xs font-medium text-amber-700 dark:border-amber-900 dark:bg-amber-950/40 dark:text-amber-300"
              >
                <StopCircle className="h-3.5 w-3.5" />
                Stop
              </button>
            )}

            <button
              type="submit"
              disabled={isProcessing || !configured || !input.trim()}
              className="rounded-none bg-red-600 p-2.5 text-white shadow-sm transition hover:bg-red-700 disabled:cursor-not-allowed disabled:opacity-50"
              aria-label="Send message"
            >
              <Send className="h-4 w-4" />
            </button>
          </div>
          <div className="composer-hints">
            <span>
              {accentOptions.find((opt) => opt.value === accent)?.label || 'English (US)'} ⌄
            </span>
            <button type="button" onClick={() => startRecording()} disabled={isProcessing || !configured}>
              Switch to typing <Keyboard className="h-3.5 w-3.5" />
            </button>
          </div>
        </form>
      )}

      <audio ref={audioRef} className="hidden" onEnded={() => setIsSpeaking(false)} />
    </aside>
  );
}
