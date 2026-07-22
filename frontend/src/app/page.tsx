'use client';

import { useQuery } from '@tanstack/react-query';
import { Search } from 'lucide-react';
import { useRouter } from 'next/navigation';
import { useEffect, useState } from 'react';
import { ChatPanel } from '@/components/chat/chat-panel';
import { DashboardOverview } from '@/components/dashboard/dashboard-overview';
import { AppSidebar } from '@/components/layout/app-sidebar';
import { MobileHeader } from '@/components/layout/mobile-header';
import { getConfig, getUserPhotoUrl, getUserProfile, logoutSession, validateSession } from '@/lib/api';
import { greeting } from '@/lib/utils';
import { useAppStore } from '@/store/use-app-store';

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
  const [search, setSearch] = useState('');

  useEffect(() => {
    let cancelled = false;

    const checkSession = async () => {
      const storedSessionId = localStorage.getItem('userSessionId');
      const activeSession = storedSessionId || sessionId;

      if (!activeSession) {
        router.replace('/login');
        return;
      }

      const isValid = await validateSession(activeSession);
      if (!isValid) {
        setSession(null);
        router.replace('/login');
        return;
      }

      if (!cancelled) {
        setSession(activeSession);
        setReady(true);
      }
    };

    checkSession().catch(() => {
      setSession(null);
      router.replace('/login');
    });

    return () => {
      cancelled = true;
    };
  }, [router, sessionId, setSession]);

  const configQuery = useQuery({
    queryKey: ['config'],
    queryFn: getConfig,
    enabled: ready
  });

  const profileQuery = useQuery({
    queryKey: ['profile', sessionId],
    queryFn: () => getUserProfile(sessionId),
    enabled: ready && !!sessionId
  });

  const photoQuery = useQuery({
    queryKey: ['profile-photo', sessionId],
    queryFn: () => getUserPhotoUrl(sessionId),
    enabled: ready && !!sessionId,
    retry: false
  });

  useEffect(() => {
    if (profileQuery.data) setUser(profileQuery.data);
  }, [profileQuery.data, setUser]);

  useEffect(() => {
    if (photoQuery.data) setUser({ photoUrl: photoQuery.data });
  }, [photoQuery.data, setUser]);

  async function handleLogout() {
    await logoutSession(sessionId);
    setSession(null);
    clearMessages();
    router.replace('/login');
  }

  function handleNewSession() {
    clearMessages();
  }

  if (!ready) {
    return (
      <main className="flex min-h-screen items-center justify-center bg-white dark:bg-slate-950">
        <div className="text-center">
          <img src="/img/Hosho DIgital-Logo.jpg" alt="Ashistanto" className="mx-auto h-14 w-auto object-contain" />
          <p className="mt-5 text-sm font-semibold text-slate-500 dark:text-slate-400">Loading secure workspace...</p>
        </div>
      </main>
    );
  }

  const configured = configQuery.data?.configured ?? true;

  return (
    <div className="min-h-screen bg-white dark:bg-slate-950">
      <div className="grid min-h-screen w-full grid-cols-1 lg:grid-cols-[280px_minmax(0,1fr)]">
        <AppSidebar onNewSession={handleNewSession} onLogout={handleLogout} />

        <main className="min-w-0 bg-gradient-to-br from-slate-50 to-slate-100 dark:from-slate-950 dark:to-slate-900">
          <header className="border-b border-slate-200 px-5 py-5 dark:border-slate-800 sm:px-7 lg:px-10">
            <MobileHeader onLogout={handleLogout} />

            <div className="flex flex-col gap-5 xl:flex-row xl:items-end xl:justify-between">
              <div>
                <h1 className="text-4xl font-extrabold tracking-tight text-[#0d2740] dark:text-white sm:text-5xl">
                  {greeting()}, {user.firstName || 'User'}.
                </h1>
                <p className="mt-3 max-w-3xl text-base text-slate-600 dark:text-slate-300 sm:text-lg">
                  Your enterprise assistant is ready to streamline today's operations.
                </p>
              </div>

              <label className="relative w-full xl:max-w-[320px]">
                <span className="pointer-events-none absolute inset-y-0 left-4 flex items-center">
                  <Search className="h-4 w-4 text-slate-500" />
                </span>
                <input
                  value={search}
                  onChange={(event) => setSearch(event.target.value)}
                  placeholder="Search enterprise data..."
                  className="w-full rounded-full border border-slate-200 bg-white py-3 pl-11 pr-4 text-sm text-slate-700 shadow-sm outline-none transition focus:border-slate-300 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-200 dark:placeholder:text-slate-500"
                />
              </label>
            </div>
          </header>

          <section className="px-5 py-6 sm:px-7 lg:px-10">
            {configQuery.isError && (
              <div className="mb-5 rounded-lg border border-red-100 bg-red-50 px-4 py-3 text-sm font-semibold text-red-700 dark:border-red-950 dark:bg-red-950/40 dark:text-red-300">
                Unable to reach the backend. Mock mode is available until FastAPI is ready.
              </div>
            )}

            <div className="grid gap-6 xl:grid-cols-[minmax(0,1fr)_370px] 2xl:grid-cols-[minmax(0,1fr)_400px]">
              <DashboardOverview
                config={configQuery.data}
                pendingActions={pendingAction ? 1 : 0}
                messageCount={messages.length}
              />
              <ChatPanel configured={configured} />
            </div>
          </section>
        </main>
      </div>
    </div>
  );
}
