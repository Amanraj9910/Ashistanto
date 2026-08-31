'use client';

import { useMutation, useQuery } from '@tanstack/react-query';
import { FileText, LogOut, Mail, Mic, PanelLeft, Send, Sparkles, Square, X } from 'lucide-react';
import { useRouter } from 'next/navigation';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ActionCard } from '@/components/chat/action-card';
import { RenderMessage } from '@/components/chat/render-message';
import {
  SessionExpiredError,
  clearConversation,
  confirmAction,
  getConfig,
  getConversations,
  getUserPhotoUrl,
  getUserProfile,
  logoutSession,
  onSessionExpired,
  processVoiceMessage,
  sendTextMessage,
  validateSession
} from '@/lib/api';
import { cleanAssistantText, formatClock } from '@/lib/utils';
import { useTtsPlayback } from '@/hooks/use-tts-playback';
import { useVoiceRecorder } from '@/hooks/use-voice-recorder';
import { useAppStore } from '@/store/use-app-store';
import type { ActionPreview, ConversationSummary, VoiceAccent } from '@/types';

type UiLanguage = 'us' | 'uk' | 'ja';
type InputMode = 'typing' | 'voice';
type ChatMessage = { role: 'user' | 'assistant'; text: string; time: string };

const translations = {
  us: { greeting: 'Hi there 👋', heading: 'How may I assist you today?', placeholder: 'Type a message...', recent: 'RECENTS', newChat: 'New conversation', connected: 'CONNECTED AS', switch: 'Switch to typing', listening: 'Listening...', speaking: 'Speaking...', processing: 'Working on it...', wait: 'This may take a few seconds.', transcript: 'TRANSCRIPT', to: 'TO', subject: 'SUBJECT', cancel: 'Cancel', send: 'Send', signOut: 'Sign out', stopSpeaking: 'Stop', autoStop: 'Auto-stop in', noConversations: 'No conversations yet' },
  uk: { greeting: 'Hi there 👋', heading: 'How may I assist you today?', placeholder: 'Type a message...', recent: 'RECENTS', newChat: 'New conversation', connected: 'CONNECTED AS', switch: 'Switch to typing', listening: 'Listening...', speaking: 'Speaking...', processing: 'Working on it...', wait: 'This may take a few seconds.', transcript: 'TRANSCRIPT', to: 'TO', subject: 'SUBJECT', cancel: 'Cancel', send: 'Send', signOut: 'Sign out', stopSpeaking: 'Stop', autoStop: 'Auto-stop in', noConversations: 'No conversations yet' },
  ja: { greeting: 'こんにちは 👋', heading: '今日はどのようにお手伝いできますか？', placeholder: 'メッセージを入力...', recent: '最近の会話', newChat: '新しい会話', connected: '接続中', switch: '入力に切り替え', listening: '聞いています...', speaking: '話しています...', processing: '処理しています...', wait: '数秒かかる場合があります。', transcript: '文字起こし', to: '宛先', subject: '件名', cancel: 'キャンセル', send: '送信', signOut: 'ログアウト', stopSpeaking: '停止', autoStop: '自動停止', noConversations: '会話はまだありません' }
} as const;

const ACCENT_BY_UI: Record<UiLanguage, VoiceAccent> = { us: 'american', uk: 'british', ja: 'japanese' };
const UI_BY_ACCENT: Record<VoiceAccent, UiLanguage> = { american: 'us', british: 'uk', japanese: 'ja' };

/** Agent replies are either plain text or an action_preview envelope. */
function parseAssistantResponse(response: string): { message: string; action?: ActionPreview } {
  try {
    const parsed = JSON.parse(response);
    if (parsed?.type === 'action_preview') {
      return { message: parsed.message || 'Please review and confirm this action.', action: parsed.preview };
    }
  } catch {
    // Plain replies are the normal case.
  }
  return { message: cleanAssistantText(response) };
}

export default function ChatPage() {
  const router = useRouter();

  const sessionId = useAppStore((state) => state.sessionId);
  const user = useAppStore((state) => state.user);
  const accent = useAppStore((state) => state.accent);
  const language = useAppStore((state) => state.language);
  const setAccent = useAppStore((state) => state.setAccent);
  const setSession = useAppStore((state) => state.setSession);
  const setUser = useAppStore((state) => state.setUser);

  const [ready, setReady] = useState(false);
  const [message, setMessage] = useState('');
  const [error, setError] = useState('');
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [inputMode, setInputMode] = useState<InputMode>('typing');
  const [chatHistory, setChatHistory] = useState<ChatMessage[]>([]);
  const [transcript, setTranscript] = useState('');
  const [pendingAction, setPendingAction] = useState<ActionPreview | null>(null);
  const conversationRef = useRef<HTMLDivElement>(null);

  const uiLanguage = UI_BY_ACCENT[accent];
  const copy = translations[uiLanguage];

  const { isSpeaking, play, stop: stopSpeaking, audioRef, onEnded } = useTtsPlayback();

  // ─── Session bootstrap ───
  useEffect(() => {
    let cancelled = false;
    const stored = typeof window !== 'undefined' ? localStorage.getItem('userSessionId') : null;
    const active = stored || useAppStore.getState().sessionId;

    if (!active) { router.replace('/login'); return; }

    validateSession(active)
      .then((valid) => {
        if (cancelled) return;
        if (!valid) { setSession(null); router.replace('/login'); return; }
        setSession(active);
        setReady(true);
      })
      .catch(() => { if (!cancelled) router.replace('/login'); });

    return () => { cancelled = true; };
    // Runs once on mount: setSession would otherwise retrigger it on every session change.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const forceLogout = useCallback(() => {
    stopSpeaking();
    setSession(null);
    setChatHistory([]);
    setPendingAction(null);
    if (typeof window !== 'undefined') localStorage.removeItem('userSessionId');
    router.replace('/login');
  }, [router, setSession, stopSpeaking]);

  // A 401 on any request means the Microsoft session died; bounce to login.
  useEffect(() => {
    const unsubscribe = onSessionExpired(forceLogout);
    return () => { unsubscribe(); };
  }, [forceLogout]);

  const configQuery = useQuery({ queryKey: ['config'], queryFn: getConfig, enabled: ready });
  const profileQuery = useQuery({ queryKey: ['profile', sessionId], queryFn: () => getUserProfile(sessionId), enabled: ready && Boolean(sessionId) });
  const photoQuery = useQuery({ queryKey: ['photo', sessionId], queryFn: () => getUserPhotoUrl(sessionId), enabled: ready && Boolean(sessionId), retry: false });
  const recentsQuery = useQuery({ queryKey: ['conversations', sessionId], queryFn: getConversations, enabled: ready });

  useEffect(() => { if (profileQuery.data) setUser(profileQuery.data); }, [profileQuery.data, setUser]);
  useEffect(() => { if (photoQuery.data) setUser({ photoUrl: photoQuery.data }); }, [photoQuery.data, setUser]);

  // getUserPhotoUrl mints a blob: URL per fetch; release it when replaced or on unmount.
  useEffect(() => {
    const url = user.photoUrl;
    return () => { if (url?.startsWith('blob:')) URL.revokeObjectURL(url); };
  }, [user.photoUrl]);

  const reportError = useCallback((mutationError: unknown) => {
    // Already navigating to /login on expiry; a banner would only flash.
    if (mutationError instanceof SessionExpiredError) return;
    setError(mutationError instanceof Error ? mutationError.message : 'Something went wrong.');
  }, []);

  const applyReply = useCallback((raw: string, spoken?: string | null) => {
    const parsed = parseAssistantResponse(raw);
    setChatHistory((items) => [...items, { role: 'assistant', text: parsed.message, time: formatClock() }]);
    setPendingAction(parsed.action || null);
    if (spoken) void play(spoken);
  }, [play]);

  const textMutation = useMutation({
    mutationFn: sendTextMessage,
    onSuccess: (data) => applyReply(data.response),
    onError: reportError
  });

  const voiceMutation = useMutation({
    mutationFn: processVoiceMessage,
    onSuccess: (data) => {
      setTranscript(data.transcript || '');
      if (data.transcript) setChatHistory((items) => [...items, { role: 'user', text: data.transcript, time: formatClock() }]);
      applyReply(data.agentResponse, data.audioData);
    },
    onError: reportError
  });

  const actionMutation = useMutation({
    mutationFn: confirmAction,
    onSuccess: (data, variables) => {
      setChatHistory((items) => [
        ...items,
        { role: 'user', text: variables.userChoice === 'confirm' ? '✅ Action confirmed' : 'Action cancelled', time: formatClock() },
        { role: 'assistant', text: cleanAssistantText(data.message || data.result || 'Done.'), time: formatClock() }
      ]);
      setPendingAction(null);
    },
    onError: reportError
  });

  const isProcessing = textMutation.isPending || voiceMutation.isPending;

  // ─── Voice ───
  const handleRecorded = useCallback((audio: Blob) => {
    if (!sessionId) { setError('Your session expired. Please sign in again.'); return; }
    setInputMode('voice');
    voiceMutation.mutate({ audio, sessionId, language, accent });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accent, language, sessionId]);

  const { isRecording, silenceCountdown, start: startRecording, stop: stopRecording } =
    useVoiceRecorder({ onRecorded: handleRecorded, onError: setError });

  const submitText = useCallback(() => {
    const text = message.trim();
    if (!text || !sessionId || isProcessing) return;
    setError('');
    setMessage('');
    setInputMode('typing');
    setChatHistory((items) => [...items, { role: 'user', text, time: formatClock() }]);
    textMutation.mutate({ text, sessionId, language, accent });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [accent, isProcessing, language, message, sessionId]);

  async function handleNewConversation() {
    stopSpeaking();
    if (isRecording) stopRecording();
    setChatHistory([]);
    setPendingAction(null);
    setTranscript('');
    setMessage('');
    setError('');
    try { await clearConversation(sessionId); } catch { /* resetting the UI must not depend on this */ }
    void recentsQuery.refetch();
  }

  async function handleLogout() {
    const active = sessionId;
    stopSpeaking();
    if (isRecording) stopRecording();
    setSession(null);
    setChatHistory([]);
    setPendingAction(null);
    if (typeof window !== 'undefined') localStorage.removeItem('userSessionId');
    router.replace('/login');
    try { await logoutSession(active); } catch { /* local session is already gone */ }
  }

  useEffect(() => {
    const node = conversationRef.current;
    if (node) requestAnimationFrame(() => { node.scrollTop = node.scrollHeight; });
  }, [chatHistory.length, pendingAction, isProcessing]);

  // ─── Derived view ───
  // The screen is derived, never stored. Storing it is what let the mockup drift out of sync
  // with the real pipeline, and it is a stale-closure hazard.
  const view = isRecording ? 'listening'
    : voiceMutation.isPending ? 'processing'
    : (isSpeaking && inputMode === 'voice' && !pendingAction) ? 'speaking'
    : chatHistory.length ? 'chat'
    : 'idle';

  const banner = useMemo(() => {
    if (error) return { text: error, tone: 'error' as const, dismissible: true };
    if (configQuery.isError) return { text: 'Unable to reach the server.', tone: 'warn' as const, dismissible: false };
    if (configQuery.data && !configQuery.data.configured) {
      return { text: 'Server not configured — voice and AI replies are unavailable.', tone: 'warn' as const, dismissible: false };
    }
    return null;
  }, [configQuery.data, configQuery.isError, error]);

  const configured = configQuery.data?.configured ?? true;
  const isVoiceView = view === 'listening' || view === 'processing' || view === 'speaking';

  if (!ready) {
    return (
      <main className="pdf-chat-boot">
        <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" />
        <p>Loading workspace…</p>
      </main>
    );
  }

  const conversation = (
    <div ref={conversationRef} className="pdf-conversation">
      <div className="pdf-today">Today</div>
      {chatHistory.map((item, index) => (
        <div className={'pdf-chat-row ' + item.role} key={`${item.role}-${index}`}>
          {item.role === 'assistant' && <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" />}
          <div className={'pdf-message ' + item.role}>
            {item.role === 'assistant' ? <RenderMessage text={item.text} /> : item.text}
            <time className="pdf-msg-time">{item.time}</time>
          </div>
          {item.role === 'user' && <span className="pdf-user-dot">{(user.firstName?.[0] || 'U').toUpperCase()}</span>}
        </div>
      ))}

      {pendingAction && (
        <ActionCard
          action={pendingAction}
          isSubmitting={actionMutation.isPending}
          labels={{ to: copy.to, subject: copy.subject, cancel: copy.cancel, send: copy.send }}
          onConfirm={(edits) => sessionId && actionMutation.mutate({
            sessionId,
            actionId: pendingAction.actionId,
            userChoice: 'confirm',
            edits: Object.keys(edits).length ? edits : null
          })}
          onCancel={() => sessionId && actionMutation.mutate({
            sessionId,
            actionId: pendingAction.actionId,
            userChoice: 'cancel'
          })}
        />
      )}

      {(textMutation.isPending || actionMutation.isPending) && (
        <div className="pdf-processing-assistant">
          <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" />
          <span className="pdf-chat-typing"><i /><i /><i /></span>
        </div>
      )}
    </div>
  );

  return (
    <main className="pdf-chat-shell">
      <aside className={sidebarOpen ? 'pdf-chat-sidebar' : 'pdf-chat-sidebar collapsed'}>
        <div className="pdf-side-brand">
          <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" />
          <button onClick={() => setSidebarOpen(false)} aria-label="Collapse sidebar"><PanelLeft size={15} /></button>
        </div>
        {sidebarOpen && (
          <>
            <button className="pdf-new-chat" onClick={handleNewConversation}><span>＋</span> {copy.newChat}</button>
            <div className="pdf-recent-label">{copy.recent}</div>
            <div className="pdf-recent-list">
              {(recentsQuery.data || []).length
                ? (recentsQuery.data as ConversationSummary[]).map((item) => (
                    <button className="pdf-recent-item" key={item.id}>
                      <span>{item.title}</span><small>{item.time}</small>
                    </button>
                  ))
                : <span className="pdf-empty-recent">{copy.noConversations}</span>}
            </div>
            <div className="pdf-side-user">
              <span className="pdf-user-avatar">
                {user.photoUrl
                  ? <img src={user.photoUrl} alt="" />
                  : (user.firstName?.[0] || 'U').toUpperCase()}
              </span>
              <span>{copy.connected}<small>{user.email}</small></span>
              <button className="pdf-logout" onClick={handleLogout} title={copy.signOut} aria-label={copy.signOut}>
                <LogOut size={13} />
              </button>
            </div>
          </>
        )}
      </aside>

      {!sidebarOpen && (
        <button className="pdf-open-sidebar" onClick={() => setSidebarOpen(true)} aria-label="Open sidebar">
          <PanelLeft size={16} />
        </button>
      )}

      <section className="pdf-chat-main">
        <header className="pdf-chat-topbar">
          <span />
          <select value={uiLanguage} onChange={(e) => setAccent(ACCENT_BY_UI[e.target.value as UiLanguage])} aria-label="Voice language">
            <option value="us">English (US)</option>
            <option value="uk">English (UK)</option>
            <option value="ja">日本語</option>
          </select>
        </header>

        {view === 'idle' && (
          <div className="pdf-idle-view">
            <div className="pdf-bot-mark"><img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" /></div>
            <p>{copy.greeting}</p>
            <h1>{copy.heading}</h1>
            <div className="pdf-prompt-grid">
              <button onClick={() => setMessage('Help me prepare a leave request')}><span className="blue"><FileText size={15} /></span>{uiLanguage === 'ja' ? '休暇申請を準備' : 'Help me prepare a leave request'}</button>
              <button onClick={() => setMessage('Write an expense report')}><span className="pink"><Mail size={15} /></span>{uiLanguage === 'ja' ? '経費報告書を書く' : 'Write an expense report'}</button>
              <button onClick={() => setMessage('Draft a meeting summary')}><span className="yellow"><Sparkles size={15} /></span>{uiLanguage === 'ja' ? '会議概要を作成' : 'Draft a meeting summary'}</button>
            </div>
          </div>
        )}

        {view === 'chat' && conversation}

        {(view === 'listening' || view === 'speaking') && (
          <div className="pdf-voice-view">
            <div className={'pdf-voice-orb ' + view}>
              <span /><span />
              <button
                onClick={() => (view === 'listening' ? stopRecording() : stopSpeaking())}
                aria-label={view === 'listening' ? 'Stop recording' : copy.stopSpeaking}
              >
                {view === 'listening' ? <Square size={26} fill="currentColor" /> : <Mic size={30} />}
              </button>
            </div>
            <h2>{view === 'listening' ? copy.listening : copy.speaking}</h2>
            {view === 'listening' && silenceCountdown > 0 && (
              <p className="pdf-countdown" aria-live="polite">⏱️ {copy.autoStop}: {silenceCountdown}s</p>
            )}
            {view === 'speaking' && transcript && (
              <div className="pdf-transcript-note"><span>{copy.transcript}</span><p>{transcript}</p></div>
            )}
          </div>
        )}

        {view === 'processing' && (
          <div className="pdf-voice-view">
            <div className="pdf-thinking-card">
              <div className="pdf-thinking-logo"><span /><span /><Sparkles size={25} /></div>
              <h2>{copy.processing}</h2>
              <p>{copy.wait}</p>
              <div className="pdf-loader-dots"><i /><i /><i /></div>
            </div>
          </div>
        )}

        {banner && (
          <div className={'pdf-error-bar' + (isVoiceView ? ' is-voice' : '') + (banner.tone === 'warn' ? ' is-warn' : '')} role="alert">
            <span>{banner.text}</span>
            {banner.dismissible && <button onClick={() => setError('')} aria-label="Dismiss"><X size={12} /></button>}
          </div>
        )}

        {(view === 'idle' || view === 'chat') && (
          <div className="pdf-chat-composer">
            <div className="pdf-composer-inner">
              <textarea
                rows={1}
                value={message}
                disabled={isProcessing}
                onChange={(e) => setMessage(e.target.value)}
                onKeyDown={(e) => {
                  // Enter sends, Shift+Enter inserts a newline - matching the previous UI.
                  if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); submitText(); }
                }}
                placeholder={copy.placeholder}
              />
              {isSpeaking && (
                <button className="pdf-stop-speaking" onClick={stopSpeaking} aria-label={copy.stopSpeaking}>
                  <Square size={12} fill="currentColor" /> {copy.stopSpeaking}
                </button>
              )}
              <button
                className="pdf-composer-mic"
                disabled={isProcessing || (!configured && !message.trim())}
                onClick={() => (message.trim() ? submitText() : void startRecording())}
                aria-label={message.trim() ? copy.send : 'Start voice input'}
              >
                {message.trim() ? <Send size={17} /> : <Mic size={17} />}
              </button>
            </div>
          </div>
        )}

        {isVoiceView && (
          <div className="pdf-voice-controls">
            <select value={uiLanguage} onChange={(e) => setAccent(ACCENT_BY_UI[e.target.value as UiLanguage])} aria-label="Voice language">
              <option value="us">English (US)</option>
              <option value="uk">English (UK)</option>
              <option value="ja">日本語</option>
            </select>
            <button onClick={() => { if (isRecording) stopRecording(); stopSpeaking(); setInputMode('typing'); }}>
              {copy.switch}
            </button>
          </div>
        )}
      </section>

      <audio ref={audioRef} onEnded={onEnded} className="hidden" />
    </main>
  );
}
