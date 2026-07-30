'use client';

import { CalendarDays, FileText, ReceiptText } from 'lucide-react';
import type { ConfigStatus } from '@/types';

type DashboardOverviewProps = {
  config?: ConfigStatus;
  pendingActions: number;
  messageCount: number;
  userName: string;
  greetingText: string;
  isSignedIn: boolean;
  onLogin: () => void;
};

const capabilities = [
  { icon: CalendarDays, title: 'Leave Management', desc: 'Apply for 3 days of annual leave starting next Monday.' },
  { icon: ReceiptText, title: 'Expense Submission', desc: 'Log a $45 client lunch with Acme Corp yesterday.' },
  { icon: CalendarDays, title: 'Calendar Operations', desc: 'Schedule a 30m sync with the design team tomorrow.' },
  { icon: FileText, title: 'Document Drafting', desc: 'Draft a Q3 summary report based on last week’s data.' }
];

export function DashboardOverview({ userName, greetingText, isSignedIn, onLogin }: DashboardOverviewProps) {
  return (
    <div className="dashboard-overview space-y-5">
      <section className="dashboard-welcome">
        <div>
          <h1>{greetingText}, {userName || 'User'}</h1>
          <p>Your enterprise assistant is ready to streamline today&apos;s operations.</p>
        </div>
      </section>

      <section className="capabilities-card">
        <div className="capabilities-header">
          <div>
            <h2>What you can do</h2>
            <p>Try these natural-language workflows with your enterprise assistant.</p>
          </div>
          <button type="button" className="capabilities-link">View all capabilities <span>→</span></button>
        </div>
        <div className="capabilities-grid">
          {capabilities.map(({ icon: Icon, title, desc }) => (
            <button key={title} type="button" className="capability-card">
              <span className="capability-icon"><Icon className="h-4 w-4" /></span>
              <span className="capability-title">{title}</span>
              <span className="capability-desc">“{desc}”</span>
            </button>
          ))}
        </div>
      </section>
    </div>
  );
}