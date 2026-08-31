'use client';

import { useEffect, useState } from 'react';
import { getAuthLoginUrl } from '@/lib/api';

export default function LoginPage() {
  const [loading, setLoading] = useState(false);
  const [expired, setExpired] = useState(false);

  // Landing here always invalidates the local session, matching the previous login page.
  // Without this a dead sessionId lingers and /chat bounces straight back here.
  useEffect(() => {
    localStorage.removeItem('userSessionId');
    setExpired(new URLSearchParams(window.location.search).has('expired'));
  }, []);

  function handleLogin() {
    setLoading(true);
    const authUrl = getAuthLoginUrl();
    if (authUrl) { window.location.href = authUrl; return; }
    // Only reachable with mocks explicitly enabled for local UI work.
    setLoading(false);
  }

  return (
    <main className="portal-login">
      <section className="portal-auth" aria-labelledby="login-title">
        <div className="portal-auth-content">
          <a className="portal-wordmark" href="/" aria-label="Ashistanto home">
            <img src="/img/cropped-logo.png" alt="Ashistanto" />
          </a>
          <div className="portal-auth-copy">
            <h1 id="login-title">Secure Portal</h1>
            <p>Enterprise Access Only.<br />Use your corporate account to continue.</p>
            {expired && <p className="portal-expired" role="status">Your session expired. Please sign in again.</p>}
            <button className="portal-ms-button" type="button" onClick={handleLogin} disabled={loading} aria-busy={loading}>
              {loading ? <span className="portal-spinner" aria-hidden="true" /> : (
                <svg width="17" height="17" viewBox="0 0 21 21" fill="none" aria-hidden="true">
                  <rect width="10" height="10" fill="#F25022" /><rect x="11" width="10" height="10" fill="#7FBA00" />
                  <rect y="11" width="10" height="10" fill="#00A4EF" /><rect x="11" y="11" width="10" height="10" fill="#FFB900" />
                </svg>
              )}
              <span>{loading ? 'Redirecting…' : 'Sign in with Microsoft'}</span>
            </button>
            <div className="portal-encrypted" role="status">
              <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" aria-hidden="true">
                <path d="M12 3 19 6v5c0 4.5-2.8 7.8-7 9.5C7.8 18.8 5 15.5 5 11V6l7-3Z" /><path d="m9.5 12 1.6 1.6 3.5-3.6" />
              </svg>
              <strong>ENCRYPTED SESSION</strong>
            </div>
          </div>
        </div>
        <footer className="portal-footer">System monitoring is active. Unauthorized access is prohibited and logged.</footer>
      </section>
      <section className="portal-visual" aria-label="Ashistanto security platform">
        <div className="portal-card">
          <div className="portal-card-rule" aria-hidden="true" />
          <h2>Empowering<br />Digital Security</h2>
          <p>Access the Hosho Digital ecosystem. Advanced threat protection and unified identity management, engineered for enterprise scale.</p>
          <div className="portal-dots" aria-hidden="true"><span className="active" /><span /></div>
        </div>
      </section>
    </main>
  );
}
