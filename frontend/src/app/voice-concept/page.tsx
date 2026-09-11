'use client';

import { useState } from 'react';
import {
  ArrowUpRight,
  AudioWaveform,
  Bell,
  CalendarDays,
  Check,
  ChevronDown,
  CircleStop,
  Clock3,
  FileText,
  Mail,
  Mic,
  MoreHorizontal,
  Paperclip,
  Play,
  Search,
  Send,
  Settings2,
  ShieldCheck,
  SquarePen,
  Volume2,
} from 'lucide-react';

type VoiceState = 'idle' | 'listening' | 'processing' | 'speaking';

const statusCopy: Record<VoiceState, { label: string; detail: string }> = {
  idle: { label: 'Ready when you are', detail: 'Tap the microphone and tell Ashistanto what you need.' },
  listening: { label: 'Listening…', detail: 'Speak naturally. Ashistanto is transcribing securely.' },
  processing: { label: 'Planning your request…', detail: 'Checking your calendar and preparing the next step.' },
  speaking: { label: 'Ashistanto is speaking', detail: 'Your Teams meeting preview is ready to review.' }
};

export default function VoiceConceptPage() {
  const [voiceState, setVoiceState] = useState<VoiceState>('idle');
  const [showPreview, setShowPreview] = useState(true);
  const [confirmed, setConfirmed] = useState(false);

  function startVoice() {
    if (voiceState !== 'idle') {
      setVoiceState('idle');
      return;
    }

    setShowPreview(false);
    setConfirmed(false);
    setVoiceState('listening');
    window.setTimeout(() => setVoiceState('processing'), 1300);
    window.setTimeout(() => {
      setVoiceState('speaking');
      setShowPreview(true);
    }, 2800);
  }

  function confirmMeeting() {
    setConfirmed(true);
    setVoiceState('idle');
  }

  const isActive = voiceState !== 'idle';
  const isListening = voiceState === 'listening';

  return (
    <main className="min-h-screen overflow-hidden bg-[#f5f7f8] text-[#173653]">
      <div className="mx-auto flex min-h-screen max-w-[1600px]">
        <aside className="hidden w-[248px] shrink-0 flex-col border-r border-[#e5eaed] bg-white px-4 py-5 lg:flex">
          <div className="flex items-center gap-3 px-2">
            <div className="grid h-9 w-9 place-items-center rounded-xl bg-[#e41c23] text-sm font-black text-white shadow-[0_8px_20px_rgba(228,28,35,.23)]">A</div>
            <div>
              <p className="text-sm font-extrabold tracking-[-.03em]">Ashistanto</p>
              <p className="text-[10px] font-medium text-[#8d9aa4]">Work companion</p>
            </div>
          </div>

          <button className="mt-8 flex items-center justify-between rounded-xl bg-[#173653] px-3.5 py-3 text-sm font-bold text-white shadow-lg shadow-[#173653]/10">
            <span className="flex items-center gap-2"><SquarePen className="h-4 w-4" /> New conversation</span>
            <span className="rounded-md bg-white/15 px-1.5 py-0.5 text-[10px]">⌘ K</span>
          </button>

          <nav className="mt-7 space-y-1">
            {[
              ['Conversations', AudioWaveform],
              ['Calendar', CalendarDays],
              ['Files', FileText],
              ['Settings', Settings2]
            ].map(([label, Icon]) => {
              const NavIcon = Icon as typeof AudioWaveform;
              return (
                <button key={label as string} className={`flex w-full items-center gap-3 rounded-lg px-3 py-2.5 text-left text-sm font-semibold transition ${label === 'Conversations' ? 'bg-[#eaf2f5] text-[#173653]' : 'text-[#75838e] hover:bg-[#f6f8f9]'}`}>
                  <NavIcon className="h-4 w-4" />
                  {label as string}
                </button>
              );
            })}
          </nav>

          <div className="mt-8">
            <p className="px-3 text-[10px] font-bold uppercase tracking-[.14em] text-[#a2adb4]">Recent</p>
            <div className="mt-2 space-y-1">
              {['Weekly project check-in', 'Travel approval for Bangkok', 'June budget review'].map((item, index) => (
                <button key={item} className={`w-full truncate rounded-lg px-3 py-2 text-left text-xs ${index === 0 ? 'bg-[#fff3f3] font-semibold text-[#c71c22]' : 'text-[#75838e] hover:bg-[#f6f8f9]'}`}>{item}</button>
              ))}
            </div>
          </div>

          <div className="mt-auto rounded-xl border border-[#e8edef] bg-[#f9fbfb] p-3">
            <div className="flex items-center gap-2.5">
              <div className="grid h-8 w-8 place-items-center rounded-full bg-[#173653] text-xs font-bold text-white">JS</div>
              <div className="min-w-0 flex-1"><p className="truncate text-xs font-bold">Jordan Smith</p><p className="truncate text-[10px] text-[#8d9aa4]">Hosho Digital</p></div>
              <ChevronDown className="h-3.5 w-3.5 text-[#8d9aa4]" />
            </div>
          </div>
        </aside>

        <section className="relative flex min-w-0 flex-1 flex-col">
          <header className="flex items-center justify-between border-b border-[#e5eaed] bg-white/80 px-5 py-4 backdrop-blur md:px-8">
            <div className="flex items-center gap-3">
              <div className="grid h-8 w-8 place-items-center rounded-lg bg-[#e41c23] text-xs font-black text-white lg:hidden">A</div>
              <div><p className="text-sm font-extrabold tracking-[-.02em]">Weekly project check-in</p><p className="text-[11px] text-[#8d9aa4]">Today, 10:42 AM</p></div>
            </div>
            <div className="flex items-center gap-2">
              <button className="hidden items-center gap-2 rounded-lg border border-[#e2e8eb] bg-white px-3 py-2 text-xs font-semibold text-[#58707e] sm:flex"><ShieldCheck className="h-3.5 w-3.5 text-emerald-600" /> Enterprise protected</button>
              <button className="grid h-9 w-9 place-items-center rounded-lg border border-[#e2e8eb] bg-white text-[#58707e]"><Bell className="h-4 w-4" /></button>
            </div>
          </header>

          <div className="mx-auto flex w-full max-w-[1030px] flex-1 flex-col px-5 pb-6 pt-8 md:px-10">
            <div className="mb-7 flex items-center justify-between">
              <div><p className="text-xs font-bold uppercase tracking-[.16em] text-[#df2830]">Voice workspace</p><h1 className="mt-1 text-2xl font-extrabold tracking-[-.045em] text-[#173653] md:text-3xl">How can I help you move work forward?</h1></div>
              <button className="hidden rounded-lg border border-[#e2e8eb] bg-white p-2 text-[#71828e] sm:block"><MoreHorizontal className="h-5 w-5" /></button>
            </div>

            <div className="space-y-4">
              <div className="flex max-w-[740px] items-start gap-3">
                <div className="mt-1 grid h-8 w-8 shrink-0 place-items-center rounded-lg bg-[#173653] text-[10px] font-black text-white">A</div>
                <div className="rounded-2xl rounded-tl-sm border border-[#e6ecee] bg-white px-4 py-3.5 shadow-sm">
                  <p className="text-sm leading-6 text-[#47606f]">Good morning, Jordan. You have a client call at 11:30 and three emails that may need a reply. What would you like to handle first?</p>
                  <p className="mt-2 text-[10px] font-medium text-[#a2adb4]">10:42 AM</p>
                </div>
              </div>

              {isActive && (
                <div className="ml-auto flex max-w-[650px] items-end gap-3">
                  <div className="rounded-2xl rounded-br-sm bg-[#173653] px-4 py-3.5 text-white shadow-lg shadow-[#173653]/10">
                    <p className="text-sm leading-6">Schedule a project review with Sarah tomorrow at 3 PM for 30 minutes.</p>
                    <p className="mt-2 text-[10px] text-white/55">Voice transcript</p>
                  </div>
                  <div className="grid h-8 w-8 shrink-0 place-items-center rounded-full bg-[#dce5e9] text-[10px] font-bold text-[#49606d]">JS</div>
                </div>
              )}

              {voiceState === 'processing' && (
                <div className="ml-11 flex max-w-[470px] items-center gap-3 rounded-xl border border-[#dce9ee] bg-[#f2f9fb] px-4 py-3 text-xs font-semibold text-[#47717f]">
                  <span className="flex gap-1"><span className="h-1.5 w-1.5 animate-bounce rounded-full bg-[#38a0bd] [animation-delay:-.2s]" /><span className="h-1.5 w-1.5 animate-bounce rounded-full bg-[#38a0bd] [animation-delay:-.1s]" /><span className="h-1.5 w-1.5 animate-bounce rounded-full bg-[#38a0bd]" /></span>
                  Checking Sarah’s availability and preparing a Teams meeting
                </div>
              )}

              {showPreview && (
                <div className="ml-0 max-w-[740px] rounded-2xl border border-[#dce7ea] bg-white p-4 shadow-[0_14px_38px_rgba(20,54,83,.07)] sm:ml-11 sm:p-5">
                  <div className="flex items-start justify-between gap-4">
                    <div className="flex gap-3"><div className="grid h-9 w-9 place-items-center rounded-xl bg-[#e9f5f8] text-[#238aa5]"><CalendarDays className="h-4.5 w-4.5" /></div><div><p className="text-[11px] font-bold uppercase tracking-[.12em] text-[#238aa5]">Ready for your review</p><h2 className="mt-0.5 font-extrabold tracking-[-.025em]">Project review with Sarah</h2></div></div>
                    <span className="flex items-center gap-1 rounded-full bg-[#eef8f0] px-2.5 py-1 text-[10px] font-bold text-emerald-700"><ShieldCheck className="h-3 w-3" /> Protected</span>
                  </div>
                  <div className="mt-5 grid gap-3 border-y border-[#edf1f2] py-4 text-sm sm:grid-cols-2">
                    <p className="flex items-center gap-2 text-[#5a707e]"><Clock3 className="h-4 w-4 text-[#9aabb4]" /> Tomorrow, 3:00–3:30 PM</p>
                    <p className="flex items-center gap-2 text-[#5a707e]"><AudioWaveform className="h-4 w-4 text-[#9aabb4]" /> Microsoft Teams meeting</p>
                    <p className="flex items-center gap-2 text-[#5a707e]"><Mail className="h-4 w-4 text-[#9aabb4]" /> Sarah Jenkins</p>
                    <p className="flex items-center gap-2 text-[#5a707e]"><CalendarDays className="h-4 w-4 text-[#9aabb4]" /> Calendar invitation ready</p>
                  </div>
                  {confirmed ? (
                    <div className="mt-4 flex items-center gap-2 rounded-lg bg-[#edf8f0] px-3 py-2.5 text-xs font-bold text-emerald-700"><Check className="h-4 w-4" /> Meeting created and invitation sent to Sarah.</div>
                  ) : (
                    <div className="mt-4 flex flex-wrap items-center justify-between gap-3"><p className="text-xs text-[#81919a]">Review the details before sending the invitation.</p><div className="flex gap-2"><button onClick={() => setShowPreview(false)} className="rounded-lg border border-[#dde6e9] px-3.5 py-2 text-xs font-bold text-[#617783]">Edit details</button><button onClick={confirmMeeting} className="rounded-lg bg-[#e41c23] px-3.5 py-2 text-xs font-bold text-white shadow-md shadow-[#e41c23]/20">Confirm meeting</button></div></div>
                  )}
                </div>
              )}
            </div>

            <div className="mt-auto pt-8">
              <div className={`relative overflow-hidden rounded-2xl border bg-white px-4 py-4 shadow-[0_18px_50px_rgba(18,54,83,.10)] transition ${isListening ? 'border-[#e41c23] ring-4 ring-[#e41c23]/10' : 'border-[#dfe7ea]'}`}>
                <div className="flex items-center gap-4">
                  <button onClick={startVoice} className={`relative grid h-14 w-14 shrink-0 place-items-center rounded-2xl text-white transition ${isListening ? 'bg-[#e41c23] shadow-[0_0_0_10px_rgba(228,28,35,.12)]' : isActive ? 'bg-[#173653]' : 'bg-[#e41c23] shadow-lg shadow-[#e41c23]/25 hover:scale-[1.03]'}`} aria-label={isActive ? 'Stop voice conversation' : 'Start voice conversation'}>
                    {isListening && <span className="absolute inset-0 animate-ping rounded-2xl bg-[#e41c23]/35" />}
                    {isActive ? <CircleStop className="relative h-5 w-5" /> : <Mic className="h-5 w-5" />}
                  </button>
                  <div className="min-w-0 flex-1"><p className="text-sm font-extrabold text-[#23445d]">{statusCopy[voiceState].label}</p><p className="mt-0.5 truncate text-xs text-[#83939c]">{statusCopy[voiceState].detail}</p>{isListening && <div className="mt-2 flex h-4 items-center gap-1">{[5, 12, 8, 16, 7, 14, 10, 5, 12, 8, 15, 6].map((height, index) => <span key={index} className="w-1 animate-pulse rounded-full bg-[#e41c23]" style={{ height: `${height}px`, animationDelay: `${index * 70}ms` }} />)}</div>}</div>
                  <div className="hidden items-center gap-2 sm:flex"><button className="grid h-9 w-9 place-items-center rounded-lg text-[#81919a] hover:bg-[#f4f7f8]"><Paperclip className="h-4 w-4" /></button><button className="grid h-9 w-9 place-items-center rounded-lg text-[#81919a] hover:bg-[#f4f7f8]"><Volume2 className="h-4 w-4" /></button></div>
                </div>
              </div>
              <div className="mt-3 flex flex-wrap items-center justify-between gap-3 px-1"><div className="flex flex-wrap gap-2">{['Check my schedule', 'Find a file', 'Draft an email'].map((prompt) => <button key={prompt} className="rounded-full border border-[#dce5e8] bg-white px-3 py-1.5 text-[11px] font-semibold text-[#66808d] hover:border-[#a9c2ca] hover:text-[#173653]">{prompt}</button>)}</div><p className="flex items-center gap-1 text-[10px] font-medium text-[#95a4ac]"><ShieldCheck className="h-3 w-3 text-emerald-600" /> Your conversation is protected</p></div>
            </div>
          </div>
        </section>

        <aside className="hidden w-[274px] shrink-0 border-l border-[#e5eaed] bg-white px-5 py-6 xl:block">
          <div className="flex items-center justify-between"><p className="text-sm font-extrabold">Today’s focus</p><button className="text-[#87969f]"><ArrowUpRight className="h-4 w-4" /></button></div>
          <div className="mt-5 space-y-3">
            <div className="rounded-xl border border-[#e5ecee] p-3.5"><div className="flex items-center justify-between"><p className="text-xs font-bold text-[#34536a]">Client call</p><span className="rounded-full bg-[#fff0f0] px-2 py-0.5 text-[9px] font-bold text-[#d32128]">In 42 min</span></div><p className="mt-2 text-xs text-[#82939c]">11:30 AM · Horizon Labs</p><button className="mt-3 flex items-center gap-1 text-[11px] font-bold text-[#d32128]">Open brief <ArrowUpRight className="h-3 w-3" /></button></div>
            <div className="rounded-xl bg-[#173653] p-3.5 text-white"><p className="text-[10px] font-bold uppercase tracking-[.12em] text-white/50">Suggested next step</p><p className="mt-2 text-sm font-bold leading-5">Reply to the three emails waiting for you.</p><button className="mt-3 flex items-center gap-1 text-[11px] font-bold text-[#95d8e9]">Review emails <Send className="h-3 w-3" /></button></div>
          </div>
          <div className="mt-8"><div className="flex items-center justify-between"><p className="text-sm font-extrabold">Quick search</p><Search className="h-4 w-4 text-[#9aabb3]" /></div><p className="mt-2 text-xs leading-5 text-[#83929b]">Search your email, files, meetings, and Teams conversations from one place.</p></div>
          <div className="mt-auto pt-10 text-center"><button className="inline-flex items-center gap-1.5 text-[11px] font-bold text-[#78909a]"><Play className="h-3.5 w-3.5 fill-current" /> How voice assistance works</button></div>
        </aside>
      </div>
    </main>
  );
}
