'use client';

import { useState } from 'react';

export type WorkspaceSection = 'Conversation' | 'Dashboard' | 'Files' | 'Settings';
import { ChevronLeft, ChevronRight, LogIn, LogOut, MessageSquare, Plus, Settings } from 'lucide-react';
import { useAppStore } from '@/store/use-app-store';
import { cn } from '@/lib/utils';
import { ThemeToggle } from './theme-toggle';

type AppSidebarProps = {
  isSignedIn: boolean;
  onLogin: () => void;
  onNewSession: () => void;
  onLogout: () => void;
  collapsed: boolean;
  onToggleCollapse: () => void;
  activeSection: WorkspaceSection;
  onNavigate: (section: WorkspaceSection) => void;
};

const railItems = [
  { icon: MessageSquare, label: 'Conversation', active: true },
  { icon: Settings, label: 'Settings' }
];

const recent = ['Debugging Code Errors', 'High Converting Financial Landing', 'Freelancer Payment Tracking', 'Banking Solutions'];

export function AppSidebar({ isSignedIn, onLogin, onNewSession, onLogout, collapsed, onToggleCollapse, activeSection, onNavigate }: AppSidebarProps) {
  const user = useAppStore((state) => state.user);
  const sidebarOpen = useAppStore((state) => state.sidebarOpen);
  const setSidebarOpen = useAppStore((state) => state.setSidebarOpen);

  const [searchOpen, setSearchOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');

  return (
    <>
      <div className={cn('fixed inset-0 z-40 bg-slate-950/40 backdrop-blur-sm lg:hidden', sidebarOpen ? 'block' : 'hidden')} onClick={() => setSidebarOpen(false)} />
      <aside className={cn('conversation-sidebar fixed inset-y-0 left-0 z-50 flex w-[300px] border-r border-slate-200 bg-white transition-all dark:border-slate-800 dark:bg-slate-950 lg:translate-x-0', sidebarOpen ? 'translate-x-0' : '-translate-x-full', collapsed && 'lg:w-[64px]')}>
        <div className="conversation-rail flex w-16 shrink-0 flex-col items-center border-r border-slate-100 py-5 dark:border-slate-800">
          <div className="mb-8 flex h-11 w-11 items-center justify-center">
            <img src="/img/Ashistanto-Red-Logo-1-transparent.png" alt="Ashistanto" className="h-7 w-7 -translate-x-0.5 object-contain" />
          </div>
          <div className="flex flex-1 flex-col items-center gap-3">
            {railItems.map(({ icon: Icon, label, active }) => (
              <button key={label} type="button" aria-label={label} aria-pressed={activeSection === label} onClick={() => { onNavigate(label as WorkspaceSection); if (label === 'Conversation' || collapsed) onToggleCollapse(); }} className={cn('flex h-10 w-10 items-center justify-center rounded-none text-slate-400 transition hover:bg-slate-100 hover:text-slate-700', activeSection === label && 'bg-sky-50 text-sky-500')}>
                <Icon className="h-5 w-5" />
              </button>
            ))}
          </div>
          <button type="button" onClick={onToggleCollapse} className="hidden h-9 w-9 items-center justify-center rounded-none text-slate-400 hover:bg-slate-100 lg:flex" aria-label={collapsed ? 'Expand navigation' : 'Collapse navigation'}>
            {collapsed ? <ChevronRight className="h-4 w-4" /> : <ChevronLeft className="h-4 w-4" />}
          </button>
        </div>

        {!collapsed && <div className="conversation-list flex min-w-0 flex-1 flex-col">
          <div className="flex items-center justify-between border-b border-slate-100 px-4 py-6 dark:border-slate-800">
            <h2 className="text-lg font-semibold text-slate-800">Conversation</h2>
            <button type="button" onClick={() => setSearchOpen((value) => !value)} className="text-slate-400 hover:text-slate-700" aria-label="Search conversations"><span className="text-xl">⌕</span></button>
          </div>
          <div className="flex-1 overflow-y-auto px-4 py-5">{searchOpen && <input autoFocus value={searchQuery} onChange={(event) => setSearchQuery(event.target.value)} placeholder="Search conversations..." className="mb-4 w-full rounded-none border border-slate-200 px-2.5 py-2 text-xs outline-none focus:border-sky-400" />}
            <p className="mb-4 text-xs font-medium text-slate-400">Recent</p>
            <div className="space-y-5">
              {recent.filter((item) => item.toLowerCase().includes(searchQuery.toLowerCase())).map((item, index) => (
                <button key={item} type="button" onClick={onNewSession} className="flex w-full items-start gap-3 text-left">
                  <MessageSquare className={cn('mt-0.5 h-4 w-4 shrink-0', index > 3 ? 'text-slate-300' : 'text-slate-400')} />
                  <span className="min-w-0"><span className="block truncate text-xs font-medium text-slate-700">{item}</span><span className="mt-1 block text-[10px] text-slate-400">{index === 0 ? '10:42 AM' : 'Yesterday'}</span></span>
                </button>
              ))}
            </div>
          </div>
          <div className="border-t border-slate-100 p-4 dark:border-slate-800">
            <div className="mb-4 flex items-center gap-2.5" style={{ minHeight: '56px' }}>
              <div className="flex items-center justify-center rounded-full bg-slate-200" style={{ width: '36px', height: '36px', minWidth: '36px', fontSize: '16px', fontWeight: 700, color: '#4c5964' }}>{user.firstName.charAt(0)}</div>
              <div className="min-w-0 flex-1">
                <div style={{ fontSize: '9px', fontWeight: 600, letterSpacing: '0.01em', color: '#6b7a85', lineHeight: 1.35, textTransform: 'uppercase' }}>Connected as</div>
                <div className="truncate" style={{ fontSize: '13px', fontWeight: 700, color: '#344353', letterSpacing: '-0.01em', lineHeight: 1.3, marginTop: '3px' }}>{isSignedIn ? (user.email || 'Microsoft account') : 'Preview mode'}</div>
              </div>
              <button type="button" onClick={isSignedIn ? onLogout : onLogin} aria-label={isSignedIn ? 'Logout' : 'Login'} className="text-slate-400 hover:text-slate-700">{isSignedIn ? <LogOut className="h-4 w-4" /> : <LogIn className="h-4 w-4" />}</button>
            </div>
            <button type="button" onClick={onNewSession} className="flex w-full items-center gap-2 rounded-none px-1 text-xs text-slate-500 hover:text-slate-800"><Plus className="h-4 w-4" /> New conversation</button>
          </div>
        </div>}
      </aside>
    </>
  );
}