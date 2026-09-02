'use client';

import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import { FileText, LogOut, Mail, Mic, PanelLeft, Send, Sparkles, Square, X } from 'lucide-react';
import { useRouter } from 'next/navigation';

type ChatState = 'idle' | 'chat' | 'listening' | 'speaking' | 'processing' | 'preview' | 'sent';
type UiLanguage = 'us' | 'uk' | 'ja';
type InputMode = 'typing' | 'voice';
type ChatMessage = { role: 'user' | 'assistant'; text: string; email?: EmailDraft; sent?: boolean };
type EmailDraft = { body: string; to: string; subject: string };
type Recent = { title: string; time: string; history?: ChatMessage[]; emails?: EmailDraft[] };

const translations = {
  us: { greeting: 'Hi there 👋', heading: 'How may I assist you today?', placeholder: 'Type a message...', recent: 'RECENTS', newChat: 'New conversation', connected: 'CONNECTED AS', switch: 'Switch to typing', listening: 'Listening...', speaking: 'Speak naturally...', processing: 'Preparing your email...', wait: 'This may take a few seconds.', transcript: 'TRANSCRIPT', preview: 'I prepared a preview. Please review and edit the details before I proceed.', to: 'TO', subject: 'SUBJECT', cancel: 'Cancel', send: 'Send', signOut: 'Sign out' },
  uk: { greeting: 'Hi there 👋', heading: 'How may I assist you today?', placeholder: 'Type a message...', recent: 'RECENTS', newChat: 'New conversation', connected: 'CONNECTED AS', switch: 'Switch to typing', listening: 'Listening...', speaking: 'Speak naturally...', processing: 'Preparing your email...', wait: 'This may take a few seconds.', transcript: 'TRANSCRIPT', preview: 'I prepared a preview. Please review and edit the details before I proceed.', to: 'TO', subject: 'SUBJECT', cancel: 'Cancel', send: 'Send', signOut: 'Sign out' },
  ja: { greeting: 'こんにちは 👋', heading: '今日はどのようにお手伝いできますか？', placeholder: 'メッセージを入力...', recent: '最近の会話', newChat: '新しい会話', connected: '接続中', switch: '入力に切り替え', listening: '聞いています...', speaking: '自然に話してください...', processing: 'メールを準備しています...', wait: '数秒かかる場合があります。', transcript: '文字起こし', preview: 'プレビューを作成しました。内容を確認して編集してください。', to: '宛先', subject: '件名', cancel: 'キャンセル', send: '送信', signOut: 'ログアウト' }
} as const;

export default function ChatPage() {
  const router = useRouter();
  const [state, setStateValue] = useState<ChatState>('idle');
  const setState = (next: ChatState) => { if (state === 'preview' && next === 'idle') { setChatHistory((items) => [...items, { role: 'assistant', text: 'No problem — the email was not sent. What would you like to do next?' }]); setStateValue('chat'); return; } setStateValue(next); };
  const [uiLanguage, setUiLanguage] = useState<UiLanguage>('us');
  const [message, setMessage] = useState('');
  const [emailBody, setEmailBody] = useState('Hi Sarah, the project schedule for this week has changed. Please review the updated timeline and let me know if you have any questions.');
  const [sidebarOpen, setSidebarOpen] = useState(true);
  const [inputMode, setInputMode] = useState<InputMode>('typing');
  const [recents, setRecents] = useState<Recent[]>([]);
  const [chatHistory, setChatHistory] = useState<ChatMessage[]>([]);
  const [sentEmails, setSentEmails] = useState<EmailDraft[]>([]);
  const [account, setAccount] = useState({ displayName: 'Microsoft account', email: '' });
  const conversationRef = useRef<HTMLDivElement>(null);
  const copy = translations[uiLanguage];
  const getDemoSessionId = () => { if (typeof window === 'undefined') return 'demo-session'; const existing = localStorage.getItem('ashistantoDemoSessionId'); if (existing) return existing; const created = `demo-${globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random().toString(36).slice(2)}`}`; localStorage.setItem('ashistantoDemoSessionId', created); return created; };
  const transcript = 'Hi Ashistanto, please prepare an email about the updated project schedule.';
  const beginProcessing = async (request: string, mode: InputMode = 'typing') => { const resolvedMode = state === 'speaking' ? 'voice' : mode; setMessage(request); setChatHistory((items) => [...items, { role: 'user', text: request }]); setInputMode(resolvedMode); setState('processing'); const apiBase = process.env.NEXT_PUBLIC_API_URL || 'http://127.0.0.1:3000'; const sessionId = typeof window !== 'undefined' ? (localStorage.getItem('userSessionId') || 'demo-session') : 'demo-session'; try { const response = await fetch(`${apiBase}/api/text-message`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ text: request, sessionId, language: uiLanguage === 'ja' ? 'ja-JP' : uiLanguage === 'uk' ? 'en-GB' : 'en-US', accent: uiLanguage === 'ja' ? 'japanese' : uiLanguage === 'uk' ? 'british' : 'american' }) }); if (response.ok) { const result = await response.json(); const parsed = typeof result.response === 'string' && result.response.startsWith('{') ? JSON.parse(result.response) : null; if (parsed?.type === 'action_preview') { setEmailBody(parsed.preview?.body || emailBody); setChatHistory((items) => [...items, { role: 'assistant', text: parsed.message || copy.preview }]); } else setChatHistory((items) => [...items, { role: 'assistant', text: result.response || copy.preview }]); setState('preview'); return; } } catch { /* fall through to the local visual fallback */ } window.setTimeout(() => { setChatHistory((items) => [...items, { role: 'assistant', text: copy.preview }]); setState('preview'); }, 1200); };
  const confirmEmail = () => { setSentEmails((items) => [...items, { body: emailBody, to: 'Sarah Wihbow', subject: 'Schedule Changes' }]); setChatHistory((items) => [...items, { role: 'user', text: 'Action confirmed' }, { role: 'assistant', text: 'Action completed successfully.' }]); setState('sent'); setMessage(''); };
  const renderEmailCard = (email: EmailDraft, sent: boolean, key: string) => <div key={key} className={'pdf-email-card '+(sent ? 'is-sent' : '')}><div className="pdf-email-head"><span><Mail size={13}/> {sent ? 'SENT EMAIL' : 'PREVIEW EMAIL'}</span>{!sent && <button onClick={() => setState('idle')}><X size={13}/></button>}</div><p><b>{copy.to}</b><input className="pdf-email-field" defaultValue={email.to} readOnly={sent}/></p><p><b>{copy.subject}</b><input className="pdf-email-field" defaultValue={email.subject} readOnly={sent}/></p><textarea className="pdf-email-body pdf-email-editor" value={sent ? email.body : emailBody} onChange={e => !sent && setEmailBody(e.target.value)} readOnly={sent}/>{!sent ? <div className="pdf-email-actions"><button onClick={() => setState('idle')}>{copy.cancel}</button><button onClick={confirmEmail}><Send size={12}/> {copy.send}</button></div> : <div className="pdf-email-sent-status">✓ Email sent successfully</div>}</div>;
  const conversationTitle = (request: string) => { const clean = request.replace(/^hi\s+ashistanto[,\s]*/i, '').replace(/[.!?]+$/, '').trim(); const lower = clean.toLowerCase(); if (lower.includes('email') || lower.includes('mail')) return `Email: ${clean.replace(/\b(email|mail)\b/gi, '').replace(/\s+/g, ' ').trim()}`.slice(0, 38); if (lower.includes('meeting') || lower.includes('calendar')) return `Meeting: ${clean.replace(/\s+/g, ' ')}`.slice(0, 38); if (lower.includes('expense')) return 'Expense report'; if (lower.includes('leave request') || lower.includes('holiday')) return 'Leave request'; return clean.length > 38 ? `${clean.slice(0, 38)}...` : clean; };
  const startNewConversation = () => { const firstRequest = chatHistory.find((item) => item.role === 'user')?.text || message.trim(); if (firstRequest) { const title = conversationTitle(firstRequest); setRecents((items) => [{ title, time: 'Just now', history: chatHistory, emails: sentEmails }, ...items.filter((item) => item.title !== title)]); } setChatHistory([]); setSentEmails([]); setEmailBody('Hi Sarah, the project schedule for this week has changed. Please review the updated timeline and let me know if you have any questions.'); setMessage(''); setStateValue('idle'); };
  const openRecent = (item: Recent) => { if (!item.history?.length) return; setChatHistory(item.history); setSentEmails(item.emails || []); setMessage(''); setState(item.emails?.length ? 'sent' : 'preview'); };

  useEffect(() => { const apiBase = process.env.NEXT_PUBLIC_API_URL || 'http://127.0.0.1:3000'; fetch(`${apiBase}/api/conversations`).then((response) => response.ok ? response.json() : []).then((items: Recent[]) => setRecents(items.filter((item) => item.title && item.title !== 'New conversation'))).catch(() => setRecents([])); }, []);
  useEffect(() => {
    const sessionId = typeof window !== 'undefined' ? (localStorage.getItem('userSessionId') || getDemoSessionId()) : null;
    if (!sessionId) return;
    const apiBase = process.env.NEXT_PUBLIC_API_URL || 'http://127.0.0.1:3000';
    fetch(`${apiBase}/api/user-profile?sessionId=${encodeURIComponent(sessionId)}`)
      .then((response) => response.ok ? response.json() : null)
      .then((profile) => profile && setAccount({ displayName: profile.displayName || profile.firstName || 'Microsoft account', email: profile.email || '' }))
      .catch(() => undefined);
  }, []);
  useEffect(() => {
    const accountNode = document.querySelector('.pdf-side-user small');
    if (accountNode) accountNode.textContent = account.email || copy.connected;
    const accountText = accountNode?.parentElement?.firstChild;
    if (accountText) accountText.textContent = account.displayName || 'Microsoft account';
    const avatarNode = document.querySelector('.pdf-side-user .pdf-user-avatar');
    if (avatarNode) avatarNode.textContent = (account.displayName || 'M').charAt(0).toUpperCase();
  }, [account, copy.connected]);
  useEffect(() => { if (typeof window !== 'undefined' && !localStorage.getItem('userSessionId')) localStorage.setItem('userSessionId', getDemoSessionId()); }, []);
  useEffect(() => { if (state === 'processing' || state === 'preview' || state === 'sent' || state === 'chat') { requestAnimationFrame(() => { const node = conversationRef.current; if (node) node.scrollTop = node.scrollHeight; }); } }, [state, chatHistory.length]);
  useLayoutEffect(() => {
    requestAnimationFrame(() => {
      document.querySelectorAll<HTMLElement>('.pdf-conversation').forEach((conversation) => {
        const cards = Array.from(conversation.children).filter((node) => node.classList.contains('pdf-email-card')) as HTMLElement[];
        const assistants = Array.from(conversation.querySelectorAll<HTMLElement>('.pdf-chat-row.assistant'));
        if (!cards.length || assistants.length < cards.length) return;
        const targets = assistants.slice(-cards.length);
        cards.forEach((card, index) => targets[index].after(card));
      });
    });
  }, [state, chatHistory.length, sentEmails.length, emailBody]);

  return <main className="pdf-chat-shell">
    <aside className={sidebarOpen ? 'pdf-chat-sidebar' : 'pdf-chat-sidebar collapsed'}><div className="pdf-side-brand"><img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto"/><button onClick={() => setSidebarOpen(false)} aria-label="Collapse sidebar"><PanelLeft size={15}/></button></div>{sidebarOpen && <><button className="pdf-new-chat" onClick={startNewConversation}><span>＋</span> {copy.newChat}</button><div className="pdf-recent-label">{copy.recent}</div><div className="pdf-recent-list">{recents.length ? recents.map((item) => <button className="pdf-recent-item" key={item.title} onClick={() => openRecent(item)}><span>{item.title}</span><small>{item.time}</small></button>) : <span className="pdf-empty-recent">{uiLanguage === 'ja' ? '会話はまだありません' : 'No conversations yet'}</span>}</div><div className="pdf-side-user"><span className="pdf-user-avatar">U</span><span>{copy.connected}<small>user@company.com</small></span><button className="pdf-logout" onClick={() => router.push('/login')} title={copy.signOut} aria-label={copy.signOut}><LogOut size={13}/></button></div></>}</aside>
    {!sidebarOpen && <button className="pdf-open-sidebar" onClick={() => setSidebarOpen(true)} aria-label="Open sidebar"><PanelLeft size={16}/></button>}
    <section className="pdf-chat-main"><header className="pdf-chat-topbar"><span/><select value={uiLanguage} onChange={e => setUiLanguage(e.target.value as UiLanguage)} aria-label="Language"><option value="us">English (US)</option><option value="uk">English (UK)</option><option value="ja">日本語</option></select></header>
      {state === 'idle' && <div className="pdf-idle-view"><div className="pdf-bot-mark"><img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto"/></div><p>{copy.greeting}</p><h1>{copy.heading}</h1><div className="pdf-prompt-grid"><button onClick={() => setMessage('Help me prepare a leave request')}><span className="blue"><FileText size={15}/></span>{uiLanguage === 'ja' ? '休暇申請を準備' : 'Help me prepare a leave request'}</button><button onClick={() => setMessage('Write an expense report')}><span className="pink"><Mail size={15}/></span>{uiLanguage === 'ja' ? '経費報告書を書く' : 'Write an expense report'}</button><button onClick={() => setMessage('Draft a meeting summary')}><span className="yellow"><Sparkles size={15}/></span>{uiLanguage === 'ja' ? '会議概要を作成' : 'Draft a meeting summary'}</button></div></div>}
      {(state === 'preview' || state === 'sent' || state === 'chat') && <div ref={conversationRef} className="pdf-conversation"><div className="pdf-today">Today</div>{chatHistory.map((item, index) => <div className={'pdf-chat-row '+item.role} key={`${item.role}-${index}`}>{item.role === 'assistant' && <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto"/>}<div className={'pdf-message '+item.role}>{item.text}</div>{item.role === 'user' && <span className="pdf-user-dot">U</span>}</div>)}{sentEmails.map((email, index) => renderEmailCard(email, true, `sent-${index}`))}{state === 'preview' && renderEmailCard({ body: emailBody, to: 'Sarah Wihbow', subject: 'Schedule Changes' }, false, 'preview')}</div>}
      {(state === 'listening' || state === 'speaking') && <div className="pdf-voice-view"><div className={'pdf-voice-orb '+state}><span/><span/><button onClick={() => state === 'speaking' ? beginProcessing(transcript) : setState('speaking')} aria-label="Voice control">{state === 'speaking' ? <Square size={26} fill="currentColor"/> : <Mic size={30}/>}</button></div><h2>{state === 'listening' ? copy.listening : copy.speaking}</h2><p>{copy.switch}</p>{state === 'speaking' && <div className="pdf-transcript-note"><span>{copy.transcript}</span><p>{transcript}</p><i>▋</i></div>}</div>}
      {state === 'processing' && (inputMode === 'voice' ? <div className="pdf-voice-view"><div className="pdf-thinking-card"><div className="pdf-thinking-logo"><span/><span/><Sparkles size={25}/></div><h2>{copy.processing}</h2><p>{copy.wait}</p><div className="pdf-loader-dots"><i/><i/><i/></div></div></div> : <div ref={conversationRef} className="pdf-conversation pdf-processing-conversation"><div className="pdf-today">Today</div>{chatHistory.map((item, index) => <div className={'pdf-chat-row '+item.role} key={`${item.role}-${index}`}>{item.role === 'assistant' && <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto"/>}<div className={'pdf-message '+item.role}>{item.text}</div>{item.role === 'user' && <span className="pdf-user-dot">U</span>}</div>)}{sentEmails.map((email, index) => renderEmailCard(email, true, `sent-processing-${index}`))}<div className="pdf-processing-assistant"><img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto"/><span className="pdf-chat-typing"><i/><i/><i/></span></div></div>)}
      {(state === 'idle' || state === 'preview' || state === 'sent' || state === 'chat') && <div className="pdf-chat-composer"><div className="pdf-composer-inner"><input value={message} onChange={e => setMessage(e.target.value)} onKeyDown={e => e.key === 'Enter' && message.trim() && beginProcessing(message, 'typing')} placeholder={copy.placeholder}/><button className="pdf-composer-mic" onClick={() => message.trim() ? beginProcessing(message, 'typing') : setState('listening')} aria-label={message.trim() ? copy.send : 'Start voice input'}>{message.trim() ? <Send size={17}/> : <Mic size={17}/>}</button></div></div>}
      {(state === 'listening' || state === 'speaking' || state === 'processing') && <div className="pdf-voice-controls"><select value={uiLanguage} onChange={e => setUiLanguage(e.target.value as UiLanguage)} aria-label="Voice language"><option value="us">English (US)</option><option value="uk">English (UK)</option><option value="ja">日本語</option></select><button onClick={() => setState('idle')}>{copy.switch}</button></div>}
    </section>
  </main>;
}
