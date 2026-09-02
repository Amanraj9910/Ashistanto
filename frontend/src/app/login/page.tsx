'use client';

import { ShieldCheck } from 'lucide-react';
import { BrandLogo } from '@/components/layout/brand-logo';

export default function LoginPage() {
  function handleLogin() {
    window.location.href = '/';
  }

  return (
    <main className="grid min-h-screen grid-cols-1 bg-white text-slate-950 lg:grid-cols-[420px_minmax(0,1fr)] xl:grid-cols-[460px_minmax(0,1fr)]">
      <section className="relative flex min-h-screen flex-col border-r border-slate-200 bg-white">
        <div className="h-2 w-full bg-red-600" />

        <div className="flex flex-1 flex-col justify-center px-8 py-10 sm:px-14">
          <div className="mt-5">
            <BrandLogo className="mb-10" />
          </div>

          <div className="max-w-sm">
            <h1 className="text-4xl font-extrabold tracking-tight text-slate-950">Secure Portal</h1>
            <p className="mt-5 text-lg leading-8 text-slate-700">
              Enterprise Access Only.
              <br />
              Use your corporate account to continue.
            </p>

            <button
              type="button"
              onClick={handleLogin}
              className="mt-10 flex w-full items-center justify-center gap-4 rounded-lg border border-slate-200 bg-white px-5 py-4 text-base font-semibold text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-slate-50 disabled:cursor-wait disabled:opacity-75"
            >
              <span className="grid h-6 w-6 grid-cols-2 gap-0.5">
                <span className="bg-[#f25022]" />
                <span className="bg-[#7fba00]" />
                <span className="bg-[#00a4ef]" />
                <span className="bg-[#ffb900]" />
              </span>
              <span>Sign in with Microsoft</span>
            </button>

            <div className="mt-10 rounded-lg border border-slate-200 bg-slate-50 px-5 py-4">
              <div className="flex items-center gap-3">
                <ShieldCheck className="h-5 w-5 text-red-600" />
                <span className="text-sm font-extrabold uppercase tracking-[0.18em] text-slate-900">Encrypted Session</span>
              </div>
            </div>
          </div>
        </div>

        <footer className="border-t border-slate-200 bg-slate-50 px-8 py-8 sm:px-14">
          <p className="max-w-sm text-sm leading-6 text-slate-600">
            System monitoring is active. Unauthorized access is prohibited and logged.
          </p>
        </footer>
      </section>

      <section className="relative hidden min-h-screen items-center justify-center overflow-hidden bg-[linear-gradient(120deg,rgba(127,29,29,.94),rgba(49,5,13,.98)_48%,rgba(25,3,8,1))] p-12 lg:flex">
        <div className="absolute inset-0 opacity-60 [background:repeating-linear-gradient(108deg,rgba(248,113,113,.17)_0_2px,transparent_2px_28px),radial-gradient(ellipse_at_78%_76%,rgba(244,63,94,.26),transparent_38%)]" />
        <div className="absolute left-[18%] right-0 top-[46%] h-28 -rotate-6 bg-[linear-gradient(90deg,transparent,rgba(248,113,113,.34),transparent)]" />

        <div className="relative z-10 w-full max-w-xl rounded-2xl border border-white/20 bg-white/35 p-12 shadow-2xl backdrop-blur-md">
          <div className="mb-12 h-1.5 w-20 rounded-full bg-red-600" />
          <h2 className="text-5xl font-extrabold leading-tight tracking-tight text-black">
            Empowering
            <br />
            Digital Security
          </h2>
          <p className="mt-9 max-w-lg text-xl leading-8 text-black/75">
            Access the Hosho Digital ecosystem. Advanced threat protection and unified identity management, engineered for
            enterprise scale.
          </p>
          <div className="mt-14 flex items-center gap-3">
            <span className="h-2.5 w-2.5 rounded-full bg-red-600" />
            <span className="h-2.5 w-2.5 rounded-full bg-white/50" />
          </div>
        </div>
      </section>
    </main>
  );
}
