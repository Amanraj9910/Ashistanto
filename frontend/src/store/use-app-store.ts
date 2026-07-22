'use client';

import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import type { ActionPreview, ChatMessage, ThemeMode, UserProfile, VoiceAccent } from '@/types';
import { createId, formatClock } from '@/lib/utils';

type AppStore = {
  sessionId: string | null;
  user: UserProfile;
  theme: ThemeMode;
  accent: VoiceAccent;
  language: string;
  messages: ChatMessage[];
  pendingAction: ActionPreview | null;
  sidebarOpen: boolean;
  setSession: (sessionId: string | null) => void;
  setUser: (user: Partial<UserProfile>) => void;
  setTheme: (theme: ThemeMode) => void;
  toggleTheme: () => void;
  setAccent: (accent: VoiceAccent) => void;
  addMessage: (message: Omit<ChatMessage, 'id' | 'timestamp'>) => void;
  clearMessages: () => void;
  setPendingAction: (action: ActionPreview | null) => void;
  setSidebarOpen: (open: boolean) => void;
};

const defaultUser: UserProfile = {
  displayName: 'User',
  firstName: 'User',
  email: 'user@company.onmicrosoft.com',
  role: 'Director of Ops'
};

const languageByAccent: Record<VoiceAccent, string> = {
  american: 'en-US',
  british: 'en-GB',
  japanese: 'en-US'
};

export const useAppStore = create<AppStore>()(
  persist(
    (set, get) => ({
      sessionId: null,
      user: defaultUser,
      theme: 'light',
      accent: 'american',
      language: 'en-US',
      messages: [],
      pendingAction: null,
      sidebarOpen: false,
      setSession: (sessionId) => {
        if (sessionId) localStorage.setItem('userSessionId', sessionId);
        else localStorage.removeItem('userSessionId');
        set({ sessionId });
      },
      setUser: (user) => set((state) => ({ user: { ...state.user, ...user } })),
      setTheme: (theme) => set({ theme }),
      toggleTheme: () => set({ theme: get().theme === 'dark' ? 'light' : 'dark' }),
      setAccent: (accent) => set({ accent, language: languageByAccent[accent] }),
      addMessage: (message) =>
        set((state) => ({
          messages: [
            ...state.messages,
            {
              ...message,
              id: createId('message'),
              timestamp: formatClock()
            }
          ]
        })),
      clearMessages: () => set({ messages: [], pendingAction: null }),
      setPendingAction: (pendingAction) => set({ pendingAction }),
      setSidebarOpen: (sidebarOpen) => set({ sidebarOpen })
    }),
    {
      name: 'ashistanto-frontend-state',
      partialize: (state) => ({
        sessionId: state.sessionId,
        user: state.user,
        theme: state.theme,
        accent: state.accent,
        language: state.language,
        messages: state.messages
      })
    }
  )
);
