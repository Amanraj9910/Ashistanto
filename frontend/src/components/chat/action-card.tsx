'use client';

import { Mail, Send, X } from 'lucide-react';
import { useState } from 'react';
import type { ActionPreview } from '@/types';
import {
  HIDDEN_FIELDS,
  buildEditPayload,
  displayLabels,
  editKeyFor,
  isFieldEditable,
  isMultiline,
  normalizeValue
} from './action-fields';

type ActionCardProps = {
  action: ActionPreview;
  isSubmitting: boolean;
  labels: { to: string; subject: string; cancel: string; send: string };
  onConfirm: (edits: Record<string, unknown>) => void;
  onCancel: () => void;
};

/**
 * Renders any pending action for review, in the chat's own pdf-email-* visual language.
 *
 * Generic over `action.details` on purpose: the same card must serve all seven action types
 * the backend configures (email, teams, meeting, meeting update, and three delete
 * confirmations). The previous version hardcoded To/Subject/Body, so a meeting or delete
 * preview rendered as an empty box.
 */
export function ActionCard({ action, isSubmitting, labels, onConfirm, onCancel }: ActionCardProps) {
  const [edits, setEdits] = useState<Record<string, string>>({});

  const fields = Object.entries(action.details).filter(([field]) => !HIDDEN_FIELDS.has(field));
  const readOnly = action.editable.length === 0;
  const isDeletion = readOnly || /delete/i.test(action.title);

  return (
    <div className={'pdf-email-card' + (isDeletion ? ' is-deletion' : '')}>
      <div className="pdf-email-head">
        <span><Mail size={13} /> {action.title || 'REVIEW REQUIRED'}</span>
        <button onClick={onCancel} disabled={isSubmitting} aria-label={labels.cancel}><X size={13} /></button>
      </div>

      {fields.map(([field, value]) => {
        const editKey = editKeyFor(field);
        const editable = !readOnly && isFieldEditable(field, action.editable);
        const shown = edits[editKey] ?? normalizeValue(value);
        const label = displayLabels[field] || field.replace(/([A-Z])/g, ' $1');

        if (isMultiline(field)) {
          return (
            <textarea
              key={field}
              className="pdf-email-body pdf-email-editor"
              aria-label={label}
              value={shown}
              readOnly={!editable}
              onChange={(event) => setEdits((current) => ({ ...current, [editKey]: event.target.value }))}
            />
          );
        }

        return (
          <p key={field}>
            <b>{label.toUpperCase()}</b>
            <input
              className="pdf-email-field"
              aria-label={label}
              value={shown}
              readOnly={!editable}
              onChange={(event) => setEdits((current) => ({ ...current, [editKey]: event.target.value }))}
            />
          </p>
        );
      })}

      <div className="pdf-email-actions">
        <button onClick={onCancel} disabled={isSubmitting}>{labels.cancel}</button>
        <button onClick={() => onConfirm(buildEditPayload(edits))} disabled={isSubmitting}>
          <Send size={12} /> {isSubmitting ? '...' : labels.send}
        </button>
      </div>
    </div>
  );
}
