'use client';

import { useRouter } from 'next/navigation';
import { useEffect, useState } from 'react';
import { MarketingHome } from '@/components/marketing/marketing-home';
import { validateSession } from '@/lib/api';

/**
 * Landing page, with an entry-point redirect for signed-in users.
 *
 * Before the UI migration, `/` WAS the app: it loaded the workspace and bounced to the login
 * page without a session. Bookmarks pointing here must keep working, so a visitor who already
 * has a session is forwarded to /chat.
 *
 * This has to happen client-side: the session lives in localStorage, not a cookie, so the
 * Express server serving this static export cannot see it and cannot issue a 302.
 *
 * The check is staged so it costs new visitors nothing:
 *  - No stored session (the common case, and every crawler) renders the marketing page on the
 *    first paint with no delay, so SEO and perceived speed are unaffected.
 *  - A stored session shows a brief splash while it is validated, then redirects.
 */
export default function HomePage() {
  const router = useRouter();
  // 'marketing' until we know otherwise: only a synchronously-found session flips this.
  const [phase, setPhase] = useState<'marketing' | 'checking'>('marketing');

  useEffect(() => {
    const stored = localStorage.getItem('userSessionId');
    if (!stored) return;

    let cancelled = false;
    setPhase('checking');

    validateSession(stored)
      .then((valid) => {
        if (cancelled) return;
        if (valid) { router.replace('/chat'); return; }
        localStorage.removeItem('userSessionId');
        setPhase('marketing');
      })
      .catch(() => { if (!cancelled) setPhase('marketing'); });

    return () => { cancelled = true; };
  }, [router]);

  if (phase === 'checking') {
    return (
      <main className="pdf-chat-boot">
        <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" />
        <p>Loading workspace…</p>
      </main>
    );
  }

  return <MarketingHome />;
}
