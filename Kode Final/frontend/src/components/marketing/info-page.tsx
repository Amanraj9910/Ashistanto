import Link from 'next/link';
import {
  ArrowRight,
  ArrowUpRight,
  CalendarDays,
  Check,
  FileText,
  Mail,
  Layers3,
  Mic,
  ShieldCheck,
  Sparkles,
  Users,
} from 'lucide-react';

type InfoKind = 'about' | 'solutions';

function MarketingNav({ kind }: { kind: InfoKind }) {
  return <header className="marketing-nav info-nav">
    <Link href="/" className="marketing-logo"><img src="/img/cropped-logo.png" alt="Ashistanto" /></Link>
    <nav aria-label="Primary navigation">
      <Link href="/">Home</Link>
      <Link href="/about" aria-current={kind === 'about' ? 'page' : undefined}>About</Link>
      <Link href="/solutions" aria-current={kind === 'solutions' ? 'page' : undefined}>Solutions</Link>
      <Link href="/contact">Contact Us</Link>
    </nav>
    <Link href="/login" className="marketing-nav-cta">Get Started <ArrowRight size={13} /></Link>
  </header>;
}

function MarketingFooter() {
  return <footer className="marketing-footer info-footer">
    <div className="footer-main">
      <div className="footer-brand">
        <Link href="/" className="marketing-logo"><img src="/img/cropped-logo.png" alt="Ashistanto" /></Link>
        <small className="footer-byline">An Ashistanto solution by Hosho Digital.</small>
        <div className="info-socials" aria-label="Social links">
          <a href="https://www.linkedin.com/company/hoshodigital/posts/?feedView=all" target="_blank" rel="noreferrer" aria-label="LinkedIn"><svg viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M20.45 20.45h-3.55v-5.57c0-1.33-.03-3.04-1.85-3.04-1.85 0-2.14 1.45-2.14 2.94v5.67H9.35V9h3.41v1.56h.05c.48-.9 1.64-1.85 3.37-1.85 3.6 0 4.27 2.37 4.27 5.46v6.28zM5.34 7.43a2.06 2.06 0 1 1 0-4.12 2.06 2.06 0 0 1 0 4.12zM3.56 20.45h3.56V9H3.56v11.45z"/></svg></a>
          <a href="https://x.com/HoshoDigital" target="_blank" rel="noreferrer" aria-label="X"><img src="https://cdn.simpleicons.org/x/ffffff" alt="" /></a>
          <a href="https://www.instagram.com/hoshodigital/" target="_blank" rel="noreferrer" aria-label="Instagram"><img src="https://cdn.simpleicons.org/instagram/ffffff" alt="" /></a>
          <a href="https://www.youtube.com/@HoshoDigital" target="_blank" rel="noreferrer" aria-label="YouTube"><img src="https://cdn.simpleicons.org/youtube/ffffff" alt="" /></a>
        </div>
      </div>
      <div><b>Company</b><a href="https://hoshodigital.com" target="_blank" rel="noreferrer">Hosho Digital</a><Link href="/">Home</Link><Link href="/about">About</Link></div>
      <div><b>Solutions</b><Link href="/solutions">Ashistanto</Link></div>
      <div><b>Contact Us</b><Link href="/contact">Contact Us</Link></div>
    </div>
    <div className="footer-bottom"><span>© 2026 HOSHO DIGITAL Pte. Ltd. ALL RIGHTS RESERVED.</span><nav><Link href="/privacy">Privacy Policy</Link><Link href="/accessibility">Accessibility Statement</Link><Link href="/terms">Terms of Use</Link><Link href="/cookies">Cookies Policy</Link></nav></div>
  </footer>;
}

function AboutContent() {
  const principles = [
    ['Voice first', 'Start with the way people already work: say what you need, then let Ashistanto turn intent into a clear next step.', Mic],
    ['Preview before action', 'Drafts, summaries, and proposed changes stay visible so people can review the details before anything is sent or changed.', Check],
    ['Microsoft native', 'The assistant is designed around Microsoft 365, connecting Outlook, Teams, OneDrive, and SharePoint in one familiar workspace.', ShieldCheck],
  ] as const;
  return <>
    <section className="info-hero info-hero-about"><div className="info-hero-inner"><p className="eyebrow light">THE ASHISTANTO WAY</p><h1>A voice-first way to move work forward.</h1><p>One assistant for the everyday work that connects your people, conversations, and Microsoft 365 tools.</p></div></section>
    <section className="info-brief"><div><p className="eyebrow">THE ASSISTANT</p><p className="info-brief-copy">Created for Microsoft 365 teams who want less manual work and more time for decisions that matter.</p></div><p className="info-brief-index">A HOSHO DIGITAL SOLUTION<br />BUILT FOR PRODUCTIVE WORK</p></section>
    <section className="info-story" id="the-idea"><div className="info-story-mark" aria-hidden="true"><img src="/img/ashistanto-japanese.png" alt="" /></div><div className="info-story-copy"><p className="eyebrow">THE IDEA</p><h2>Useful intelligence should feel like a natural conversation.</h2><p>Traditional workplace software makes people navigate menus, remember where information lives, and repeat the same steps every day. Ashistanto starts with the outcome instead.</p><p>Speak or type the request in plain language. The assistant keeps the context close, prepares the next action, and shows the result for approval.</p></div></section>
    <section className="info-dark-section"><div className="info-section-head"><div><p className="eyebrow light">THE PRINCIPLES</p><h2>Built to earn independence.</h2></div><p>Automation should stay connected to a real business need while becoming more useful, governable, and self-sustaining.</p></div><div className="info-principles">{principles.map(([title, copy, Icon]) => <article key={title}><span className="info-principle-icon"><Icon size={18} /></span><div><h3>{title}</h3></div><p>{copy}</p></article>)}</div></section>
    <section className="info-long-view"><div><p className="eyebrow">THE LONG VIEW</p><h2>From request to result. From result to capability.</h2><p>As Ashistanto learns the way your team works, recurring tasks become easier to complete and easier to trust.</p><Link href="/solutions" className="button button-red">Explore solutions <ArrowRight size={14} /></Link></div><div className="info-long-view-panel"><span>01</span><strong>Understand the request</strong><span>02</span><strong>Prepare a useful result</strong><span>03</span><strong>Keep the decision with you</strong></div></section>
  </>;
}

function SolutionsContent() {
  const capabilities = [
    ['Email', 'Draft, summarise, reply to, and organise messages without opening Outlook for every task.', Mail],
    ['Calendar', 'Find availability, prepare meeting notes, and turn follow-ups into clear next actions.', CalendarDays],
    ['Teams', 'Keep conversations moving with concise updates, summaries, and decisions ready to share.', Users],
    ['Files', 'Find the right document, surface key information, and work with SharePoint and OneDrive context.', FileText],
  ] as const;
  const flow = [['SPEAK', 'Describe the outcome in your own words.'], ['PREVIEW', 'Review the proposed email, summary, or change.'], ['APPROVE', 'Confirm the action and keep a record of what happened.']];
  return <>
    <section className="info-hero info-hero-solutions"><div className="info-hero-inner"><p className="eyebrow light">SOLUTIONS</p><h1>Solve the business. Then apply the technology.</h1><p>Ashistanto brings the right Microsoft 365 action closer to the person who needs it.</p></div></section>
    <section className="info-brief"><div><p className="eyebrow">THE STARTING POINT</p><p className="info-brief-copy">The business need comes first. Ashistanto helps your team turn that need into a clear, reviewable action.</p></div><p className="info-brief-index">VOICE · CONTEXT · ACTION</p></section>
    <section className="solution-proof"><p className="eyebrow">WHY A VOICE-FIRST LAYER</p><h2>Everyday work is full of small hand-offs.</h2><p>Finding the message, preparing the reply, checking the details, and confirming the action all take time. Ashistanto connects those steps without hiding the decision.</p></section>
    <section className="solution-intelligence"><div className="info-section-head"><div><p className="eyebrow light">CONNECTED INTELLIGENCE</p><h2>Transforming business through connected intelligence.</h2></div><p>One consistent experience across the tools where your team already communicates, plans, and stores knowledge.</p></div><div className="solution-pillars">{['Operational intelligence','Customer intelligence','Workforce intelligence'].map((title, index) => <article key={title}><span>0{index + 1}</span><h3>{title}</h3><p>Make the next decision easier to find, review, and act on.</p></article>)}</div></section>
    <section className="info-capabilities" id="capabilities"><div className="info-section-head"><div><p className="eyebrow">SOLUTIONS</p><h2>From one request to many useful outcomes.</h2></div><p>Choose the Microsoft 365 surface where the work starts. Ashistanto carries the context through the rest of the flow.</p></div><div className="solution-grid">{capabilities.map(([title, copy, Icon]) => <article key={title}><div className="solution-card-top"><span className="info-icon"><Icon size={19} /></span><span className="solution-card-arrow"><ArrowRight size={16} /></span></div><h3>{title}</h3><p>{copy}</p><span className="solution-example">Try: “Prepare a {title.toLowerCase()} update.”</span></article>)}</div></section>
    <section className="info-flow-section"><div className="info-section-head"><div><p className="eyebrow">THE ASHISTANTO FLOW</p><h2>Accelerate from intent to action.</h2></div><p>The experience is deliberately short: say what you need, check the details, and decide what happens next.</p></div><div className="info-flow">{flow.map(([stage, copy], index) => <article key={stage}><span className="flow-number">0{index + 1}</span><p className="stage-name">{stage}</p><h3>{copy}</h3>{index < flow.length - 1 && <ArrowRight className="flow-arrow" size={20} />}</article>)}</div></section>
    <section className="info-security"><div className="info-security-intro"><p className="eyebrow light">ENTERPRISE SECURITY</p><h2>Automation should never feel like a black box.</h2><p>Designed to help teams move faster while keeping permissions, review, and accountability visible.</p></div><div className="info-security-cards"><article><ArrowUpRight size={22}/><h3>End-to-end action previews</h3></article><article><Layers3 size={22}/><h3>Scoped Microsoft 365 permissions</h3></article><article><Sparkles size={22}/><h3>Context that stays with the conversation</h3></article><article><ShieldCheck size={22}/><h3>Review before send or execute</h3></article></div></section>
  </>;
}

export function InfoPage({ kind }: { kind: InfoKind }) {
  return <main className={`info-page info-${kind}`}><MarketingNav kind={kind} />{kind === 'about' ? <AboutContent /> : <SolutionsContent />}<section className="info-proof"><p className="eyebrow">READY WHEN YOU ARE</p><h2>Make the next action clear.</h2><p>See how Ashistanto can fit the way your team already works.</p><Link href="/contact" className="button button-red">Talk to our team <ArrowRight size={14} /></Link></section><MarketingFooter /></main>;
}
