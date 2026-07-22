'use client';

import { useState } from 'react';
import type { ActionPreview } from '@/types';

type ActionPreviewCardProps = {
  action: ActionPreview;
  isSubmitting: boolean;
  onConfirm: (edits: Record<string, unknown>) => void;
  onCancel: () => void;
};

const displayLabels: Record<string, string> = {
  to: 'To',
  recipient: 'Recipient',
  subject: 'Subject',
  body: 'Body',
  cc: 'CC',
  attendees: 'Attendees',
  startTime: 'Start',
  endTime: 'End',
  isTeams: 'Teams Meeting',
  employee: 'Employee',
  leaveType: 'Leave Type',
  startDate: 'Start Date',
  status: 'Status'
};

function editKeyFor(field: string) {
  const mapping: Record<string, string> = {
    to: 'recipientName',
    attendees: 'attendeeNames',
    isTeams: 'isTeamsMeeting'
  };

  return mapping[field] || field;
}

function normalizeValue(value: unknown) {
  if (Array.isArray(value)) return value.join(', ');
  if (typeof value === 'boolean') return value ? 'Yes' : 'No';
  if (value === null || value === undefined || value === '') return '-';
  return String(value);
}

export function ActionPreviewCard({ action, isSubmitting, onConfirm, onCancel }: ActionPreviewCardProps) {
  const [editing, setEditing] = useState(false);
  const [edits, setEdits] = useState<Record<string, string>>({});

  const fields = Object.entries(action.details);
  const canEdit = fields.some(([key]) => action.editable.includes(editKeyFor(key)) || action.editable.includes(key));

  return (
    <div className="rounded-lg border border-red-100 bg-white p-4 shadow-sm dark:border-red-950/70 dark:bg-slate-950">
      <div className="flex items-start justify-between gap-3 border-b border-slate-100 pb-3 dark:border-slate-800">
        <div>
          <div className="text-xs font-bold uppercase tracking-[0.18em] text-red-600 dark:text-red-500">Review Required</div>
          <h4 className="mt-1 text-base font-extrabold text-slate-950 dark:text-white">{action.title}</h4>
        </div>
        <span className="rounded-full bg-red-50 px-3 py-1 text-xs font-bold uppercase text-red-700 dark:bg-red-950/50 dark:text-red-300">
          Pending
        </span>
      </div>

      <div className="mt-4 space-y-3">
        {fields.map(([field, value]) => {
          const editKey = editKeyFor(field);
          const editable = action.editable.includes(editKey) || action.editable.includes(field);
          const renderedValue = edits[editKey] ?? normalizeValue(value);
          const multiline = field === 'body' || field === 'message';

          return (
            <label key={field} className="block">
              <span className="text-xs font-bold uppercase tracking-wide text-slate-500 dark:text-slate-400">
                {displayLabels[field] || field.replace(/([A-Z])/g, ' $1')}
              </span>
              {editing && editable ? (
                multiline ? (
                  <textarea
                    value={renderedValue}
                    onChange={(event) => setEdits((current) => ({ ...current, [editKey]: event.target.value }))}
                    className="mt-1 min-h-[76px] w-full rounded-md border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-slate-800 outline-none focus:border-red-300 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-100"
                  />
                ) : (
                  <input
                    value={renderedValue}
                    onChange={(event) => setEdits((current) => ({ ...current, [editKey]: event.target.value }))}
                    className="mt-1 w-full rounded-md border border-slate-200 bg-slate-50 px-3 py-2 text-sm text-slate-800 outline-none focus:border-red-300 dark:border-slate-800 dark:bg-slate-900 dark:text-slate-100"
                  />
                )
              ) : (
                <span className="mt-1 block rounded-md bg-slate-50 px-3 py-2 text-sm leading-5 text-slate-700 dark:bg-slate-900 dark:text-slate-200">
                  {renderedValue}
                </span>
              )}
            </label>
          );
        })}
      </div>

      <div className="mt-4 flex flex-wrap justify-end gap-2 border-t border-slate-100 pt-3 dark:border-slate-800">
        <button
          type="button"
          onClick={onCancel}
          disabled={isSubmitting}
          className="rounded-md border border-slate-200 bg-white px-3 py-2 text-sm font-semibold text-slate-700 transition hover:bg-slate-50 disabled:opacity-50 dark:border-slate-800 dark:bg-slate-950 dark:text-slate-200 dark:hover:bg-slate-900"
        >
          Cancel
        </button>
        {canEdit && (
          <button
            type="button"
            onClick={() => setEditing((current) => !current)}
            disabled={isSubmitting}
            className="rounded-md border border-red-100 bg-red-50 px-3 py-2 text-sm font-semibold text-red-700 transition hover:bg-red-100 disabled:opacity-50 dark:border-red-950 dark:bg-red-950/40 dark:text-red-300"
          >
            {editing ? 'Done Editing' : 'Edit'}
          </button>
        )}
        <button
          type="button"
          onClick={() => onConfirm(edits)}
          disabled={isSubmitting}
          className="rounded-md bg-red-600 px-3 py-2 text-sm font-semibold text-white transition hover:bg-red-700 disabled:opacity-50"
        >
          {isSubmitting ? 'Processing...' : 'Confirm'}
        </button>
      </div>
    </div>
  );
}
