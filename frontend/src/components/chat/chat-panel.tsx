'use client';

import { useMutation } from '@tanstack/react-query';
import { Mic, Send, Square, StopCircle } from 'lucide-react';
import { FormEvent, useEffect, useRef, useState } from 'react';
import { clearConversation, confirmAction, processVoiceMessage, sendTextMessage } from '@/lib/api';
import { cleanAssistantText, cn } from '@/lib/utils';
import { useAppStore } from '@/store/use-app-store';
import type { ActionPreview, VoiceAccent } from '@/types';
import { ActionPreviewCard } from './action-preview-card';

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

export function ChatPanel({ configured }: { configured: boolean }) {
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
        user: variables.userChoice === 'confirm' ? 'Action confirmed' : 'Action cancelled',
        assistant: data.message || (variables.userChoice === 'confirm' ? 'Action completed successfully.' : 'Action cancelled.')
      });
      setPendingAction(null);
    },
    onError: (mutationError) => {
      setError(mutationError instanceof Error ? mutationError.message : 'Failed to update action.');
    }
  });

  const isProcessing = textMutation.isPending || voiceMutation.isPending || actionMutation.isPending;

  useEffect(() => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' });
  }, [messages, pendingAction, isProcessing]);

  function handleSubmit(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();
    const text = input.trim();

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
    <aside className="flex min-h-[640px] flex-col overflow-hidden rounded-lg border border-slate-200 bg-white shadow-sm dark:border-slate-800 dark:bg-slate-950 xl:h-[calc(100vh-156px)]">
      <div className="flex items-center justify-between border-b border-slate-200 bg-white px-5 py-4 dark:border-slate-800 dark:bg-slate-950">
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center overflow-hidden rounded-md border border-slate-200 bg-white dark:border-slate-800">
            <img src="/img/Hosho DIgital-Logo.jpg" alt="Ashistanto AI" className="h-full w-full object-cover" />
          </div>
          <div>
            <div className="text-base font-extrabold text-slate-950 dark:text-white">ASHISTANTO AI</div>
            <div className="text-xs font-extrabold uppercase tracking-wide text-orange-500">
              {isProcessing ? 'Processing' : isRecording ? 'Listening' : 'Online'}
            </div>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {messages.length > 0 && (
            <button
              type="button"
              onClick={handleClear}
              className="rounded-full border border-slate-200 px-3 py-2 text-xs font-bold text-slate-600 hover:bg-slate-50 dark:border-slate-800 dark:text-slate-300 dark:hover:bg-slate-900"
            >
              Clear
            </button>
          )}
          <select
            value={accent}
            onChange={(event) => setAccent(event.target.value as VoiceAccent)}
            className="rounded-full border border-slate-200 bg-slate-50 px-3 py-2 text-xs font-bold text-slate-600 outline-none dark:border-slate-800 dark:bg-slate-900 dark:text-slate-300"
          >
            {accentOptions.map((option) => (
              <option key={option.value} value={option.value}>
                {option.label}
              </option>
            ))}
          </select>
        </div>
      </div>

      <div className="flex-1 space-y-5 overflow-y-auto bg-slate-50 px-4 py-5 dark:bg-slate-900/70">
        <div className="flex justify-center">
          <span className="rounded-full bg-slate-200 px-3 py-1 text-xs font-bold text-slate-600 dark:bg-slate-800 dark:text-slate-300">
            Today
          </span>
        </div>

        {messages.length === 0 && (
          <div className="flex items-start gap-3">
            <div className="flex h-8 w-8 shrink-0 items-center justify-center overflow-hidden rounded-full border border-slate-200 bg-white dark:border-slate-800">
              <img src="/img/Hosho DIgital-Logo.jpg" alt="Ashistanto AI" className="h-full w-full object-cover" />
            </div>
            <div className="max-w-[280px] rounded-br-2xl rounded-tr-2xl rounded-bl-2xl border border-slate-100 bg-white px-4 py-3 text-sm leading-6 text-slate-800 shadow-sm dark:border-slate-800 dark:bg-slate-950 dark:text-slate-100">
              Good morning. I can help with Microsoft 365 workflows, scheduling, email, Teams messages, files, and operational requests.
            </div>
          </div>
        )}

        {messages.map((message) => (
          <div key={message.id} className="space-y-4">
            <div className="flex items-start justify-end gap-3">
              <div className="max-w-[280px] whitespace-pre-wrap rounded-bl-2xl rounded-tl-2xl rounded-tr-2xl bg-[#0d2740] px-4 py-3 text-sm leading-6 text-white shadow-sm dark:bg-red-700">
                {message.user}
              </div>
              {user.photoUrl ? (
                <img
                  src={user.photoUrl}
                  alt={`${user.firstName} profile`}
                  className="h-8 w-8 shrink-0 rounded-full border border-slate-200 object-cover dark:border-slate-700"
                />
              ) : (
                <div className="flex h-8 w-8 shrink-0 items-center justify-center rounded-full border border-slate-200 bg-white text-xs font-extrabold text-slate-700 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200">
                  {user.firstName.charAt(0).toUpperCase()}
                </div>
              )}
            </div>

            <div className="flex items-start gap-3">
              <div className="flex h-8 w-8 shrink-0 items-center justify-center overflow-hidden rounded-full border border-slate-200 bg-white dark:border-slate-800">
                <img src="/img/Hosho DIgital-Logo.jpg" alt="Ashistanto AI" className="h-full w-full object-cover" />
              </div>
              <div className="max-w-[300px] whitespace-pre-wrap rounded-br-2xl rounded-tr-2xl rounded-bl-2xl border border-slate-100 bg-white px-4 py-3 text-sm leading-6 text-slate-800 shadow-sm dark:border-slate-800 dark:bg-slate-950 dark:text-slate-100">
                {cleanAssistantText(message.assistant)}
              </div>
            </div>
          </div>
        ))}

        {pendingAction && (
          <div className="flex items-start gap-3">
            <div className="flex h-8 w-8 shrink-0 items-center justify-center overflow-hidden rounded-full border border-slate-200 bg-white dark:border-slate-800">
              <img src="/img/Hosho DIgital-Logo.jpg" alt="Ashistanto AI" className="h-full w-full object-cover" />
            </div>
            <div className="w-full max-w-[320px]">
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
          <div className="flex items-start gap-3">
            <div className="flex h-8 w-8 shrink-0 items-center justify-center overflow-hidden rounded-full border border-slate-200 bg-white dark:border-slate-800">
              <img src="/img/Hosho DIgital-Logo.jpg" alt="Ashistanto AI" className="h-full w-full object-cover" />
            </div>
            <div className="rounded-br-2xl rounded-tr-2xl rounded-bl-2xl bg-white px-4 py-3 text-sm font-semibold text-slate-500 shadow-sm dark:bg-slate-950 dark:text-slate-300">
              Processing request...
            </div>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {error && <div className="border-t border-red-100 bg-red-50 px-4 py-2 text-xs font-semibold text-red-700 dark:border-red-950 dark:bg-red-950/40 dark:text-red-300">{error}</div>}

      <form onSubmit={handleSubmit} className="border-t border-slate-200 bg-white p-4 dark:border-slate-800 dark:bg-slate-950">
        <div className="flex items-end gap-2 rounded-lg border border-slate-200 bg-slate-50 px-3 py-2 shadow-[inset_0_2px_4px_rgba(15,23,42,.05)] dark:border-slate-800 dark:bg-slate-900">
          <button
            type="button"
            onClick={isRecording ? stopRecording : startRecording}
            disabled={isProcessing || !configured}
            className={cn(
              'rounded-md p-2 transition disabled:cursor-not-allowed disabled:opacity-50',
              isRecording ? 'bg-red-600 text-white' : 'text-slate-500 hover:bg-white dark:text-slate-300 dark:hover:bg-slate-800'
            )}
            aria-label={isRecording ? 'Stop recording' : 'Start recording'}
          >
            {isRecording ? <Square className="h-5 w-5 fill-current" /> : <Mic className="h-5 w-5" />}
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
            placeholder={isRecording ? 'Listening...' : 'Type a command or ask a question...'}
            disabled={isProcessing || !configured}
            rows={1}
            className="max-h-28 min-h-[42px] min-w-0 flex-1 resize-none bg-transparent py-2 text-sm leading-6 text-slate-800 outline-none placeholder:text-slate-500 disabled:cursor-not-allowed dark:text-slate-100 dark:placeholder:text-slate-500"
          />

          {isSpeaking && (
            <button
              type="button"
              onClick={stopSpeaking}
              className="inline-flex items-center gap-1 rounded-md border border-orange-100 bg-orange-50 px-3 py-2 text-xs font-bold text-orange-700 dark:border-orange-950 dark:bg-orange-950/40 dark:text-orange-300"
            >
              <StopCircle className="h-4 w-4" />
              Stop
            </button>
          )}

          <button
            type="submit"
            disabled={isProcessing || !configured || !input.trim()}
            className="rounded-md bg-red-600 p-3 text-white shadow-sm transition hover:bg-red-700 disabled:cursor-not-allowed disabled:opacity-50"
            aria-label="Send message"
          >
            <Send className="h-5 w-5" />
          </button>
        </div>
      </form>

      <audio ref={audioRef} className="hidden" onEnded={() => setIsSpeaking(false)} />
    </aside>
  );
}
