'use client';

import { BarChart3, Bot, HelpCircle, LayoutDashboard, LogOut, Plus, Zap } from 'lucide-react';
import { useRouter } from 'next/navigation';
import { useAppStore } from '@/store/use-app-store';
import { cn } from '@/lib/utils';
import { BrandLogo } from './brand-logo';
import { ThemeToggle } from './theme-toggle';

type AppSidebarProps = {
  onNewSession: () => void;
  onLogout: () => void;
};

const navItems = [
  { icon: LayoutDashboard, label: 'Dashboard', active: true },
];

export function AppSidebar({ onNewSession, onLogout }: AppSidebarProps) {
  const router = useRouter();
  const user = useAppStore((state) => state.user);
  const sidebarOpen = useAppStore((state) => state.sidebarOpen);
  const setSidebarOpen = useAppStore((state) => state.setSidebarOpen);

  return (
    <>
      <div
        className={cn(
          'fixed inset-0 z-40 bg-slate-950/40 backdrop-blur-sm lg:hidden',
          sidebarOpen ? 'block' : 'hidden'
        )}
        onClick={() => setSidebarOpen(false)}
      />
      <aside
        className={cn(
          'fixed inset-y-0 left-0 z-50 flex w-[280px] flex-col border-r border-slate-200 bg-white px-5 py-7 transition-transform dark:border-slate-800 dark:bg-slate-950 lg:static lg:translate-x-0',
          sidebarOpen ? 'translate-x-0' : '-translate-x-full'
        )}
      >
        <BrandLogo />
        
        <div className="mt-8 grid gap-3">
          <button
            type="button"
            onClick={() => {
              router.push('/login');
              setSidebarOpen(false);
            }}
            className="inline-flex items-center justify-center gap-3 rounded-xl border border-slate-200 bg-white px-4 py-3 text-base font-semibold text-slate-700 shadow-sm transition hover:border-slate-300 hover:bg-slate-50 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-200 dark:hover:bg-slate-900"
          >
            <LogOut className="h-5 w-5" />
            Login
          </button>

          <button
            type="button"
            onClick={() => {
              onNewSession();
              setSidebarOpen(false);
            }}
            className="inline-flex items-center justify-center gap-3 rounded-xl bg-red-600 px-4 py-3 text-base font-bold text-white shadow-sm transition hover:bg-red-700"
          >
            <Plus className="h-5 w-5" />
            New Session
          </button>
        </div>

        <nav className="mt-6 flex flex-1 flex-col gap-2">
          {navItems.map(({ icon: Icon, label, active }) => (
            <button
              key={label}
              type="button"
              className={cn(
                'flex items-center gap-4 rounded-lg px-4 py-3 text-left text-base transition',
                active
                  ? 'bg-slate-200 font-extrabold text-slate-950 dark:bg-slate-800 dark:text-white'
                  : 'font-semibold text-slate-500 hover:bg-slate-50 dark:text-slate-400 dark:hover:bg-slate-900'
              )}
            >
              <Icon className="h-5 w-5" />
              <span>{label}</span>
            </button>
          ))}
        </nav>

        <div className="mt-auto space-y-4 pt-8">
          <div className="flex items-center justify-between">
            <button
              type="button"
              className="flex items-center gap-3 rounded-lg px-3 py-3 text-base font-semibold text-slate-500 transition hover:bg-slate-50 dark:text-slate-400 dark:hover:bg-slate-900"
            >
              <HelpCircle className="h-5 w-5" />
              <span>Help &amp; Support</span>
            </button>
            <ThemeToggle />
          </div>

          <div className="border-t border-slate-200 dark:border-slate-800" />

          <div className="flex items-center gap-3 px-2 pt-2">
            {user.photoUrl ? (
              <img
                src={user.photoUrl}
                alt={`${user.firstName} profile`}
                className="h-11 w-11 shrink-0 rounded-full border border-slate-200 object-cover dark:border-slate-700"
              />
            ) : (
              <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-full border border-slate-200 bg-slate-100 text-sm font-extrabold text-slate-700 dark:border-slate-700 dark:bg-slate-900 dark:text-slate-200">
                {user.firstName.charAt(0).toUpperCase()}
              </div>
            )}
            <div className="min-w-0 flex-1">
              <div className="truncate text-base font-extrabold text-slate-950 dark:text-white">{user.firstName}</div>
              <div className="truncate text-sm font-semibold text-slate-500 dark:text-slate-400">{user.role}</div>
            </div>
            <button
              type="button"
              onClick={onLogout}
              className="rounded-md p-2 text-slate-500 transition hover:bg-slate-100 hover:text-slate-900 dark:hover:bg-slate-900 dark:hover:text-white"
              aria-label="Logout"
            >
              <LogOut className="h-5 w-5" />
            </button>
          </div>
        </div>
      </aside>
    </>
  );
}
