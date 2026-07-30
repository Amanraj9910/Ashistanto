'use client';

import { useQuery } from '@tanstack/react-query';
import { useRouter } from 'next/navigation';

import { useEffect, useState } from 'react';
import { ChatPanel } from '@/components/chat/chat-panel';
import { WorkspaceView } from '@/components/workspace-view';

import { AppSidebar, type WorkspaceSection } from '@/components/layout/app-sidebar';
import { MobileHeader } from '@/components/layout/mobile-header';
import { getConfig, getUserPhotoUrl, getUserProfile, logoutSession, validateSession } from '@/lib/api';
import { createId } from '@/lib/utils';
import type { VoiceAccent } from '@/types';
import { useAppStore } from '@/store/use-app-store';

const previewUser = {
  displayName: 'Guest Preview',
  firstName: 'Guest',
  email: '',
  role: 'Preview Mode',
  photoUrl: undefined
};

export default function DashboardPage() {
  const router = useRouter();
  const sessionId = useAppStore((state) => state.sessionId);
  const user = useAppStore((state) => state.user);
  const messages = useAppStore((state) => state.messages);
  const pendingAction = useAppStore((state) => state.pendingAction);
  const setSession = useAppStore((state) => state.setSession);
  const setUser = useAppStore((state) => state.setUser);
  const clearMessages = useAppStore((state) => state.clearMessages);

  const [ready, setReady] = useState(false);
const [sidebarCollapsed, setSidebarCollapsed] = useState(true);
  const [activeSection, setActiveSection] = useState<WorkspaceSection>('Conversation');
  const accent = useAppStore((state) => state.accent);
  const setAccent = useAppStore((state) => state.setAccent);
  const isSignedIn = Boolean(sessionId);

  useEffect(() => {
    let cancelled = false;

    const checkSession = async () => {
      const storedSessionId = localStorage.getItem('userSessionId');
      const activeSession = storedSessionId || sessionId;

      if (!activeSession) {
        setSession(null);
        setUser(previewUser);
        clearMessages();
        if (!cancelled) setReady(true);
        return;
      }

      const isValid = await validateSession(activeSession);
      if (!isValid) {
        setSession(null);
        setUser(previewUser);
        clearMessages();
        if (!cancelled) setReady(true);
        return;
      }

      if (!cancelled) {
        setSession(activeSession);
        setReady(true);
      }
    };

    checkSession().catch(() => {
      setSession(null);
      setUser(previewUser);
      clearMessages();
      if (!cancelled) setReady(true);
    });

    return () => {
      cancelled = true;
    };
  }, [clearMessages, sessionId, setSession, setUser]);

  const configQuery = useQuery({
    queryKey: ['config'],
    queryFn: getConfig,
    enabled: ready
  });

  const profileQuery = useQuery({
    queryKey: ['profile', sessionId],
    queryFn: () => getUserProfile(sessionId),
    enabled: ready && isSignedIn
  });

  const photoQuery = useQuery({
    queryKey: ['profile-photo', sessionId],
    queryFn: () => getUserPhotoUrl(sessionId),
    enabled: ready && isSignedIn,
    retry: false
  });

  useEffect(() => {
    if (profileQuery.data) setUser(profileQuery.data);
  }, [profileQuery.data, setUser]);

  useEffect(() => {
    if (photoQuery.data) setUser({ photoUrl: photoQuery.data });
  }, [photoQuery.data, setUser]);

  function handleLogin() {
    setSession(createId('mock_session'));
    setUser({
      displayName: 'User',
      firstName: 'User',
      email: 'user@company.onmicrosoft.com',
      role: 'Director of Ops',
      photoUrl: undefined
    });
  }

  async function handleLogout() {
    await logoutSession(sessionId);
    setSession(null);
    clearMessages();
    setUser(previewUser);
    router.push('/login');
  }

  function handleNewSession() {
    clearMessages();
  }

  if (!ready) {
    return (
      <main className="flex min-h-screen items-center justify-center bg-white dark:bg-slate-950">
        <div className="text-center animate-fade-in">
          <img src="/img/Hosho-Digital-Logo.png" alt="Ashistanto" className="mx-auto h-12 w-auto object-contain" />
          <p className="mt-4 text-sm text-slate-400 dark:text-slate-500">Loading workspace…</p>
        </div>
      </main>
    );
  }

  const configured = configQuery.data?.configured ?? true;

  return (
    <div className="min-h-screen bg-white dark:bg-slate-950">
      <div className="min-h-screen w-full">
        <AppSidebar
          isSignedIn={isSignedIn}
          onLogin={handleLogin}
          onNewSession={handleNewSession}
          onLogout={handleLogout}
          collapsed={sidebarCollapsed}
          onToggleCollapse={() => setSidebarCollapsed((value) => !value)}
          activeSection={activeSection}
          onNavigate={setActiveSection}
        />

        <main className="dashboard-main min-w-0 bg-slate-50 dark:bg-slate-950">
          <header className="conversation-topbar">
  <MobileHeader isSignedIn={isSignedIn} onLogin={handleLogin} onLogout={handleLogout} />
  <div className="conversation-controls">
  <select value={accent} onChange={(event) => setAccent(event.target.value as VoiceAccent)} className="language-select" aria-label="Voice language">
    <option value="american">English (US)</option>
    <option value="british">English (UK)</option>
    <option value="japanese">日本語</option>
  </select>
</div>
</header>
<section className={activeSection === 'Conversation' ? 'conversation-stage' : activeSection === 'Settings' ? 'conversation-stage settings-stage' : 'workspace-stage'}>
  {activeSection === 'Conversation' ? <>{configQuery.isError && <div className="mb-3 rounded-lg border border-red-100 bg-red-50 px-4 py-2 text-xs font-medium text-red-700">Unable to reach the backend. Preview mode is available.</div>}<ChatPanel configured={configured} isSignedIn={isSignedIn} onLogin={handleLogin} /></> : <WorkspaceView section={activeSection} user={user} accent={accent} onAccentChange={setAccent} />}
</section>
        </main>
      </div>
    </div>
  );
}
