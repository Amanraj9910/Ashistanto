import Link from 'next/link';

type LegalKind = 'contact' | 'privacy' | 'terms' | 'cookies' | 'accessibility';

const content: Record<LegalKind, { label: string; title: string; intro: string; sections: { heading: string; body: string }[] }> = {
  contact: {
    label: 'CONTACT',
    title: 'Let’s identify where Ashistanto can help.',
    intro: 'Tell us what your team is trying to improve across Microsoft 365. The Hosho Digital team will respond with a practical next step.',
    sections: [
      { heading: 'Start a conversation', body: 'Share your organisation, the workflow you want to improve, and the Microsoft 365 tools your team uses today.' },
      { heading: 'Email', body: 'hello@hoshodigital.com' },
      { heading: 'Company', body: 'Hosho Digital Pte. Ltd.\nSingapore' },
    ],
  },
  privacy: {
    label: 'PRIVACY POLICY',
    title: 'Your workspace stays under your control.',
    intro: 'This policy explains how Ashistanto handles account, conversation, and Microsoft 365 data when the service is connected to your organisation.',
    sections: [
      { heading: 'Data we use', body: 'We use account details, conversation context, and action data required to provide the requested assistant experience.' },
      { heading: 'Microsoft 365 permissions', body: 'Ashistanto only uses permissions granted by your Microsoft account. Actions that change or send data require explicit confirmation.' },
      { heading: 'Retention and requests', body: 'Contact Hosho Digital to request access, correction, or deletion of personal data associated with your account.' },
    ],
  },
  terms: {
    label: 'TERMS OF USE',
    title: 'A clear agreement for using Ashistanto.',
    intro: 'These terms describe the responsible use of Ashistanto and the boundaries around voice commands, previews, and Microsoft 365 actions.',
    sections: [
      { heading: 'Use of the service', body: 'Use Ashistanto only with accounts and data you are authorised to access. Review generated content before approving an action.' },
      { heading: 'Your responsibility', body: 'You remain responsible for recipients, instructions, and approvals made through your Microsoft 365 account.' },
      { heading: 'Service changes', body: 'The service may evolve as integrations, security controls, and Microsoft platform requirements change.' },
    ],
  },
  accessibility: {
    label: 'ACCESSIBILITY STATEMENT',
    title: 'Designed to be usable by everyone.',
    intro: 'Ashistanto is committed to making its website and assistant experience accessible, inclusive, and easy to use.',
    sections: [
      { heading: 'Our commitment', body: 'We work toward conformance with recognised accessibility standards and continuously improve the experience across devices, browsers, and input methods.' },
      { heading: 'Accessible experience', body: 'We support keyboard navigation, readable contrast, descriptive labels, responsive layouts, and reduced-motion preferences wherever possible.' },
      { heading: 'Feedback', body: 'If you encounter an accessibility barrier, please contact us with the page, issue, and assistive technology you are using so we can improve it.' },
    ],
  },
  cookies: {
    label: 'COOKIES POLICY',
    title: 'A transparent approach to cookies.',
    intro: 'This policy explains how Ashistanto uses cookies and similar technologies to keep the experience secure, reliable, and useful.',
    sections: [
      { heading: 'Essential cookies', body: 'Essential cookies support sign-in, secure sessions, language preferences, and core product functionality.' },
      { heading: 'Analytics', body: 'Where enabled, privacy-conscious analytics help us understand product usage and improve the service. We do not sell personal data.' },
      { heading: 'Your choices', body: 'You can manage non-essential cookies through your browser settings. Disabling essential cookies may affect the service.' },
    ],
  },
};

export function LegalPage({ kind }: { kind: LegalKind }) {
  const page = content[kind];
  return <main className={`legal-page legal-${kind}`}>
    <header className="legal-nav"><Link href="/" className="marketing-logo"><img src="/img/cropped-logo.png" alt="Ashistanto" /></Link><nav><Link href="/">Home</Link><Link href="/about">About</Link><Link href="/solutions">Solutions</Link><Link href="/contact">Contact Us</Link></nav><Link href="/login" className="marketing-nav-cta">Get Started <span aria-hidden="true">→</span></Link></header>
    <section className="legal-hero" aria-hidden="true"></section>
    <section className="legal-content"><h1>{page.title}</h1><p className="legal-intro">{page.intro}</p>{page.sections.map((section) => <article key={section.heading}><h2>{section.heading}</h2><p>{section.body}</p></article>)}<article className="legal-contact-callout"><h2>Contact Us</h2><p>If you have any questions or concerns regarding this policy, please contact us at <a href="mailto:privacy@hoshodigital.com">privacy@hoshodigital.com</a>.</p></article></section>
    <footer className="marketing-footer legal-site-footer"><div className="footer-main"><div className="footer-brand"><Link href="/" className="marketing-logo"><img src="/img/cropped-logo.png" alt="Ashistanto" /></Link><small className="footer-byline">An Ashistanto solution by Hosho Digital.</small><div className="footer-socials"><a href="https://www.linkedin.com/company/hoshodigital" target="_blank" rel="noreferrer" aria-label="LinkedIn">in</a><a href="https://x.com/HoshoDigital" target="_blank" rel="noreferrer" aria-label="X">𝕏</a><a href="https://www.instagram.com/hoshodigital/" target="_blank" rel="noreferrer" aria-label="Instagram">◎</a><a href="https://www.youtube.com/@HoshoDigital" target="_blank" rel="noreferrer" aria-label="YouTube">▶</a></div></div><div><b>Solutions</b><Link href="/solutions">Ashistanto</Link></div><div><b>Company</b><a href="https://hoshodigital.com" target="_blank" rel="noreferrer">Hosho Digital</a><Link href="/">Home</Link><Link href="/about">About</Link></div></div><div className="footer-bottom"><span>© 2026 HOSHO DIGITAL Pte. Ltd. ALL RIGHTS RESERVED.</span><nav><Link href="/privacy">Privacy Policy</Link><Link href="/accessibility">Accessibility Statement</Link><Link href="/terms">Terms of Use</Link><Link href="/cookies">Cookies Policy</Link></nav></div></footer>
  </main>;
}
