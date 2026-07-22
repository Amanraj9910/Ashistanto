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
    <main className="flex min-h-screen items-center justify-center bg-white px-6 text-slate-950">
      <section className="w-full max-w-xl text-center">
        <div className="mx-auto flex h-28 w-28 items-center justify-center rounded-full border border-slate-100 bg-slate-100 shadow-[inset_0_0_0_8px_rgba(255,255,255,.7)]">
          <Check className="h-14 w-14 text-slate-700" />
        </div>

        <h1 className="mt-12 text-5xl font-extrabold tracking-tight text-slate-950">Login Successful!</h1>
        <p className="mt-7 text-xl text-slate-500">Welcome back,</p>
        <p className="mt-3 text-xl font-extrabold text-slate-950">{user.email}</p>

        <div className="mt-16 flex items-center justify-center gap-4 text-lg text-slate-500">
          <span className="h-6 w-6 animate-spin rounded-full border-2 border-red-600 border-t-transparent" />
          <span>Redirecting to your dashboard...</span>
        </div>

        <div className="mx-auto mt-10 h-2 w-full max-w-[460px] overflow-hidden rounded-full bg-slate-100">
          <div className="h-full animate-[progress_2s_ease-out_forwards] rounded-full bg-red-700" />
        </div>

        <p className="fixed bottom-12 left-0 right-0 text-center text-sm font-semibold text-slate-500">
          Secure session established
        </p>
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
