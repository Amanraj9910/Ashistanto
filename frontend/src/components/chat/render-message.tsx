'use client';

import { Fragment } from 'react';
import { cleanAssistantText } from '@/lib/utils';

// Matches a markdown link, or a bare URL. Same pattern the pre-migration UI used, so links
// that were clickable before stay clickable.
const LINK_PATTERN = /\[([^\]]+)\]\((https?:\/\/[^\s)]+)\)|(https?:\/\/[^\s]+)/g;

/**
 * Renders an assistant reply: strips markdown emphasis, then turns markdown links and bare
 * URLs into real anchors. Without this, `cleanAssistantText` alone leaves URLs as dead text.
 */
export function RenderMessage({ text }: { text: unknown }) {
  const cleaned = cleanAssistantText(text);
  const parts: Array<string | { label: string; href: string }> = [];
  let cursor = 0;

  for (const match of cleaned.matchAll(LINK_PATTERN)) {
    const index = match.index ?? 0;
    if (index > cursor) parts.push(cleaned.slice(cursor, index));

    const [full, mdLabel, mdHref, bareUrl] = match;
    if (bareUrl) parts.push({ label: bareUrl, href: bareUrl });
    else parts.push({ label: mdLabel, href: mdHref });

    cursor = index + full.length;
  }
  if (cursor < cleaned.length) parts.push(cleaned.slice(cursor));

  return (
    <span className="whitespace-pre-wrap break-words">
      {parts.map((part, index) =>
        typeof part === 'string' ? (
          <Fragment key={index}>{part}</Fragment>
        ) : (
          <a
            key={index}
            href={part.href}
            target="_blank"
            rel="noopener noreferrer"
            className="break-all text-sky-600 underline-offset-2 hover:underline"
          >
            {part.label}
          </a>
        )
      )}
    </span>
  );
}
