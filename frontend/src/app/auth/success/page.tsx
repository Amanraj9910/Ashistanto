'use client';

import { Check } from 'lucide-react';
import { useRouter, useSearchParams } from 'next/navigation';
import { Suspense, useEffect } from 'react';
import { createId } from '@/lib/utils';
import { useAppStore } from '@/store/use-app-store';

function SuccessContent() {
  const router = useRouter();
  const params = useSearchParams();
  const setSession = useAppStore((state) => state.setSession);
  const user = useAppStore((state) => state.user);

  useEffect(() => {
    setSession(params.get('sessionId') || createId('frontend_session'));
    const timer = setTimeout(() => router.replace('/'), 2000);
    return () => clearTimeout(timer);
  }, [params, router, setSession]);

  return (
    <main className="flex min-h-screen items-center justify-center bg-white px-6 text-slate-900">
      <section className="w-full max-w-md text-center animate-fade-in">
        <div className="mx-auto flex h-20 w-20 items-center justify-center rounded-full bg-emerald-50">
          <Check className="h-10 w-10 text-emerald-600" strokeWidth={2.5} />
        </div>

        <h1 className="mt-8 text-2xl font-bold tracking-tight text-slate-900 sm:text-3xl">
          You&apos;re signed in
        </h1>
        <p className="mt-3 text-base text-slate-500">
          Welcome back, <span className="font-semibold text-slate-700">{user.email || 'User'}</span>
        </p>

        <div className="mx-auto mt-10 flex items-center justify-center gap-3 text-sm text-slate-400">
          <span className="h-4 w-4 animate-spin rounded-full border-2 border-slate-300 border-t-transparent" />
          <span>Taking you to your workspace…</span>
        </div>

        <div className="mx-auto mt-6 h-1 w-full max-w-xs overflow-hidden rounded-full bg-slate-100">
          <div className="h-full animate-[progress_2s_ease-out_forwards] rounded-full bg-red-600" />
        </div>
      </section>
    </main>
  );
}

export default function AuthSuccessPage() {
  return (
    <Suspense fallback={null}>
      <SuccessContent />
    </Suspense>
  );
}
