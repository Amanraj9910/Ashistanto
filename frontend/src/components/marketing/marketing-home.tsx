'use client';

import { useState } from 'react';
import Link from 'next/link';
import { ArrowRight, Check, ChevronRight, Eye, FileText, LockKeyhole, LogIn, Mic, Play, Send, ShieldCheck, Volume2, Zap } from 'lucide-react';

const logos = [
  ['Microsoft 365', '/img/microsoft-365-copilot-logo-png_seeklogo-501781.png'],
  ['Outlook', '/img/Outlook-Logo---Colored---68x64---zonalogo.com.png'],
  ['Teams', '/img/Microsoft-Teams-Logo.png'],
  ['SharePoint', '/img/Microsoft-SharePoint-Logo---Colored---58x64---zonalogo.com.png'],
  ['OneDrive', '/img/microsoft-onedrive-2025-logo-png_seeklogo-644157.png'],
  ['Excel', '/img/Excel-logo-png-large-size.png'],
  ['Word', '/img/png-transparent-microsoft-word-logo.png']
];

export function MarketingHome() {
  const [capability, setCapability] = useState('Email');
  const capabilityCopy: Record<string, { title: string; description: string; bullets: string[] }> = {
    Email: { title: 'Intelligent Email Management', description: 'Compose, send, read, and delete emails without ever opening Outlook. Ashistanto drafts professional HTML emails with proper formatting and signatures automatically.', bullets: ['Send emails by recipient name — no address needed', 'Auto-formatted with professional HTML styling', 'Read recent emails and summarize key points', 'Delete sent emails with a single command'] },
    Calendar: { title: 'Effortless Calendar Control', description: 'Create, update, and review meetings by voice while Ashistanto checks availability and keeps your schedule organized.', bullets: ['Find available time slots instantly', 'Create meetings with attendees and agenda', 'Reschedule or cancel with one command', 'Receive concise daily schedule summaries'] },
    Teams: { title: 'Connected Teams Communication', description: 'Stay connected with your team without switching applications. Send messages and get important updates through natural conversation.', bullets: ['Send messages to people or channels', 'Summarize unread conversations', 'Share meeting notes with your team', 'Find messages by topic or person'] },
    Files: { title: 'Smart File Access', description: 'Find, summarize, and organize your Microsoft 365 files using a simple voice command.', bullets: ['Search OneDrive and SharePoint', 'Summarize long documents', 'Find the latest version of a file', 'Share documents with the right people'] }
  };
  const activeCapability = capabilityCopy[capability];
  return (
    <main className="marketing-page">
      <header className="marketing-nav">
        <Link href="/" className="marketing-logo"><img src="/img/cropped-logo.png" alt="Ashistanto" /></Link>
        <nav><a href="#features">Features</a><a href="#how">How It Works</a><a href="#capabilities">Capabilities</a><a href="#testimonials">Testimonials</a></nav>
        <Link href="/login" className="marketing-nav-cta">Get Started <ArrowRight size={13} /></Link>
      </header>

      <section className="marketing-hero" style={{ backgroundImage: "url('/img/background-hero.png')" }}>
        <div className="hero-copy">
          <p className="eyebrow light">YOUR AI VOICE</p>
          <h1>Your AI <em>Voice</em><br /><em>Assistant</em> for<br />Microsoft 365</h1>
          <p className="hero-sub">Say goodbye to manual tasks. Ashistanto lets you manage your entire digital workspace on Microsoft 365—Teams, and Outlook through natural voice commands.</p>
          <div className="hero-actions"><Link href="/login" className="button button-red">Start Free <ArrowRight size={14} /></Link><a href="#how" className="button button-white"><Play size={12} fill="currentColor" /> See How It Works</a></div>
        </div>
        <div className="hero-preview"><div className="preview-top"><span></span><span></span><span></span></div><div className="preview-line"></div><div className="preview-block"></div></div>
      </section>

      <section className="stats-row">{[['6+', 'M365 INTEGRATIONS'], ['99.9%', 'UPTIME RELIABILITY'], ['10x', 'FASTER THAN MANUAL'], ['24/7', 'ALWAYS AVAILABLE']].map(([value, label]) => <div key={label}><strong>{value}</strong><small>{label}</small></div>)}</section>
      <section className="ecosystem"><p>POWERED BY THE MICROSOFT ECOSYSTEM</p><div>{[...logos, ...logos, ...logos].map(([name, src], index) => <span key={`${name}-${index}`} aria-hidden={index >= logos.length}><img src={src} alt={index >= logos.length ? '' : name} />{name}</span>)}</div></section>

      <section id="how" className="section section-muted"><div className="section-heading left"><p className="eyebrow">SIMPLE SETUP</p><h2>Up and Running in Three Steps</h2><p>No complex configurations. Connect your Microsoft account and start talking to Ashistanto in under a minute.</p></div><div className="step-grid">{[['01','Sign in with Microsoft','Securely connect your Microsoft 365 account with a single click. Enterprise-grade OAuth authentication keeps your data protected.'],['02','Speak Naturally','No rigid commands to memorize. Just talk like you would to a colleague—“send an email to Sarah about tomorrow’s meeting.”'],['03','Watch It Execute','Ashistanto processes your request in real-time—drafting emails, creating meetings, and more. Results are confirmed before anything executes.']].map(([num,title,copy]) => <article className="step-card" key={num}><b>{num}</b><div className="icon-tile">{num === '01' ? <LogIn size={17}/> : num === '02' ? <Mic size={17}/> : <Check size={17}/>}</div><h3>{title}</h3><p>{copy}</p></article>)}</div></section>

      <section id="capabilities" className="section capabilities"><div className="section-heading center"><p className="eyebrow">CAPABILITIES</p><h2>Everything You Need, One Voice Command Away</h2><p>Ashistanto integrates deeply with Microsoft 365 to handle your most common workplace tasks.</p></div><div className="cap-tabs">{Object.keys(capabilityCopy).map((tab) => <button key={tab} className={capability === tab ? 'active' : ''} onClick={() => setCapability(tab)}>{tab}</button>)}</div><div className="cap-content capability-transition" key={capability}><div><h2>{activeCapability.title}</h2><p>{activeCapability.description}</p>{activeCapability.bullets.map((item) => <p className="check-line" key={item}><Check size={14}/>{item}</p>)}</div><div className="email-mock"><div className="mock-window"><span className="mock-dot red"></span><span className="mock-dot yellow"></span><span className="mock-dot green"></span><div className="mock-input">{capability === 'Email' ? 'Send an email to John about the project...' : `Ask Ashistanto about ${capability.toLowerCase()}...`}</div><div className="mock-success">{capability} action prepared — ready for your review</div></div></div></div></section>

      <section id="testimonials" className="quote-band"><div><span className="quote-mark">“</span><blockquote>The future of work is not about replacing people with machines. It is about giving people intelligent tools that amplify what they do best.</blockquote><small>SATYA NADELLA<br />CEO of Microsoft</small></div><img className="quote-person-image" src="/img/satya-nadella.png" alt="Satya Nadella" /></section>

      <section id="features" className="section section-muted"><div className="section-heading center"><p className="eyebrow">VOICE-FIRST DESIGN</p><h2>Built Around Your Voice</h2><p>Ashistanto is designed with the microphone as the primary interface. Speak naturally and watch real-time transcription in your assistant work.</p></div><div className="voice-grid">{[['Listening','Waveform animation shows active recording, while real-time transcription lets you see exactly what Ashistanto hears as you speak.'],['Processing','AI understands your intent and context, turning natural voice commands into relevant actions across your Microsoft environment.'],['Previewing','Generated responses, content, and actions are presented for review, giving you full control before anything is sent, changed, or executed.'],['Speaking','Natural text-to-speech responses with multi-accent support make conversations feel seamless, accessible, and closer to interacting with a real assistant.']].map(([title,copy],i) => <article key={title}><div className="voice-icon">{i===0?<Mic size={16}/>:i===1?<Zap size={16}/>:i===2?<Eye size={16}/>:<Volume2 size={16}/>}</div><h3>{title}</h3><p>{copy}</p></article>)}</div></section>

      <section className="security-band"><div className="section-heading center"><p className="eyebrow light">ENTERPRISE SECURITY</p><h2>You Are Always in Control</h2><p>Every action requires your explicit approval. Your data stays within Microsoft’s secure ecosystem.</p></div><div className="security-grid">{([{title:'Microsoft OAuth',copy:'Enterprise-grade authentication via Azure AD. Your credentials never leave our servers.',Icon:LockKeyhole},{title:'Action Preview',copy:'Every interaction shows a preview before execution. Edit, confirm, or cancel at any step.',Icon:ShieldCheck},{title:'Scoped Access',copy:'Ashistanto only accesses data your Microsoft account has permission for. No extra privileges.',Icon:FileText}]).map(({title,copy,Icon}) => <article key={title}><Icon size={19}/><h3>{title}</h3><p>{copy}</p></article>)}</div></section>

      <section className="section section-muted use-cases"><div className="section-heading left"><p className="eyebrow">DAILY WORKFLOWS</p><h2>How Teams Use Ashistanto</h2><p>Real scenarios that save time every day.</p></div><div className="use-grid">{['Morning Routine','Meeting Prep','Follow-Up','Document Search','Quick Reply','Research'].map((title,i) => <article className={i%2===0?'pink':''} key={title}><small>{i<3?'START YOUR DAY':'FIND IT FAST'}</small><h3>{title}</h3><p>“{i===0?'What’s my schedule today and which emails need attention?':'Give me the latest updates and prepare the next action.'}”</p></article>)}</div></section>

      <section className="marketing-cta" style={{ backgroundImage: "url('/img/get-started-banner.png')" }}><h2>Ready to Transform How You Work?</h2><p>Join the next generation of workplace productivity. Start using Ashistanto today.</p><Link href="/login" className="button button-white">Get Started — It’s Free <ArrowRight size={14}/></Link></section>
      <footer className="marketing-footer"><div><Link href="/" className="marketing-logo"><img src="/img/cropped-logo.png" alt="Ashistanto" /></Link><p>Your intelligent voice-powered AI assistant for Microsoft 365. Manage your entire digital workspace through natural conversation.</p></div><div><b>Product</b><a href="#features">Features</a><a href="#capabilities">Capabilities</a><a href="/chat">Try it Now</a></div><div><b>Integrations</b><a>Microsoft Outlook</a><a>Microsoft Teams</a><a>OneDrive</a><a>SharePoint</a></div><div><b>Company</b><a>About Hosho Digital</a><a>Privacy Policy</a><a>Terms of Service</a><a>Contact Us</a></div></footer>
    </main>
  );
}
