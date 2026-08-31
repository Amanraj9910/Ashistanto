'use client';

import { LogIn, LogOut, Menu } from 'lucide-react';
import { useAppStore } from '@/store/use-app-store';
import { BrandLogo } from './brand-logo';
import { ThemeToggle } from './theme-toggle';

type MobileHeaderProps = {
  isSignedIn: boolean;
  onLogin: () => void;
  onLogout: () => void;
};

export function MobileHeader({ isSignedIn, onLogin, onLogout }: MobileHeaderProps) {
  const setSidebarOpen = useAppStore((state) => state.setSidebarOpen);

  return (
    <div className="mb-5 flex items-center justify-between gap-3 lg:hidden">
      <button
        type="button"
        onClick={() => setSidebarOpen(true)}
        className="inline-flex h-10 w-10 items-center justify-center rounded-none border border-slate-200 bg-white text-slate-600 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-300"
        aria-label="Open navigation"
      >
        <Menu className="h-5 w-5" />
      </button>
      <BrandLogo className="min-w-0 flex-1 justify-center" imageClassName="h-9" />
      <ThemeToggle />
      <button
        type="button"
        onClick={isSignedIn ? onLogout : onLogin}
        className="inline-flex h-10 w-10 items-center justify-center rounded-none border border-slate-200 bg-white text-slate-600 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-300"
        aria-label={isSignedIn ? 'Logout' : 'Login'}
      >
        {isSignedIn ? <LogOut className="h-5 w-5" /> : <LogIn className="h-5 w-5" />}
      </button>
    </div>
  );
}
