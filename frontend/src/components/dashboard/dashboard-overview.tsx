'use client';

import { ArrowRight, CalendarCheck, CheckCircle2, ClipboardList, FileText, Mail, ReceiptText } from 'lucide-react';
import type { ConfigStatus } from '@/types';

type DashboardOverviewProps = {
  config?: ConfigStatus;
  pendingActions: number;
  messageCount: number;
};

const capabilities = [
  {
    icon: CalendarCheck,
    title: 'Leave Management',
    desc: '"Apply for 3 days of annual leave starting next Monday."'
  },
  {
    icon: ReceiptText,
    title: 'Expense Submission',
    desc: '"Log a $45 client lunch with Acme Corp yesterday."'
  },
  {
    icon: CalendarCheck,
    title: 'Calendar Operations',
    desc: '"Schedule a 30m sync with the design team tomorrow."'
  },
  {
    icon: FileText,
    title: 'Document Drafting',
    desc: '"Draft a Q3 summary report based on last week\'s data."'
  }
];

export function DashboardOverview({ config, pendingActions, messageCount }: DashboardOverviewProps) {
  const configured = config?.configured ?? true;

  return (
    <div className="space-y-6">
      <section className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm dark:border-slate-800 dark:bg-slate-950 sm:p-8">
        <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
          <h2 className="text-2xl font-extrabold text-slate-950 dark:text-white">Enterprise Capabilities</h2>
          <button type="button" className="inline-flex items-center gap-1 self-start text-sm font-extrabold text-red-600 dark:text-red-500">
            View all integrations
            <ArrowRight className="h-4 w-4" />
          </button>
        </div>

        <p className="mt-5 max-w-2xl text-base leading-7 text-slate-600 dark:text-slate-300">
          Execute complex workflows directly through natural language processing in the assistant.
        </p>

        <div className="mt-8 grid gap-4 md:grid-cols-2">
          {capabilities.map(({ icon: Icon, title, desc }) => (
            <article
              key={title}
              className="rounded-lg border border-slate-200 bg-slate-50 p-5 transition hover:border-red-100 hover:bg-white dark:border-slate-800 dark:bg-slate-900 dark:hover:border-red-950 dark:hover:bg-slate-950"
            >
              <div className="flex h-11 w-11 items-center justify-center rounded-md bg-red-50 text-red-600 dark:bg-red-950/40 dark:text-red-400">
                <Icon className="h-5 w-5" />
              </div>
              <h3 className="mt-5 text-lg font-extrabold text-slate-950 dark:text-white">{title}</h3>
              <p className="mt-2 text-base leading-6 text-slate-600 dark:text-slate-300">{desc}</p>
            </article>
          ))}
        </div>
      </section>

      <section className="grid gap-4 md:grid-cols-[190px_190px_minmax(0,1fr)]">
        <article className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm dark:border-slate-800 dark:bg-slate-950">
          <ClipboardList className="h-5 w-5 text-slate-600 dark:text-slate-300" />
          <div className="mt-5 text-4xl font-extrabold text-[#0d2740] dark:text-white">{pendingActions || 12}</div>
          <div className="mt-1 text-sm font-extrabold uppercase tracking-wide text-slate-600 dark:text-slate-400">Pending Tasks</div>
        </article>

        <article className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm dark:border-slate-800 dark:bg-slate-950">
          <Mail className="h-5 w-5 text-slate-600 dark:text-slate-300" />
          <div className="mt-5 text-4xl font-extrabold text-[#0d2740] dark:text-white">{messageCount}</div>
          <div className="mt-1 text-sm font-extrabold uppercase tracking-wide text-slate-600 dark:text-slate-400">Messages</div>
        </article>

        <article className="rounded-lg border border-slate-200 bg-white p-6 shadow-sm dark:border-slate-800 dark:bg-slate-950">
          <div className="flex items-center justify-between gap-4">
            <div>
              <h3 className="text-3xl font-extrabold text-slate-950 dark:text-white">System Health</h3>
              <p className="mt-2 text-base text-slate-600 dark:text-slate-300">
                {configured ? 'All enterprise services operational' : 'Configuration attention required'}
              </p>
            </div>
            <CheckCircle2 className={configured ? 'h-12 w-12 text-emerald-500' : 'h-12 w-12 text-slate-300 dark:text-slate-700'} />
          </div>
        </article>
      </section>
    </div>
  );
}
