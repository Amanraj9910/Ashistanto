'use client';

import { useCallback, useEffect, useRef, useState } from 'react';

function base64ToBlob(base64: string, type: string) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return new Blob([bytes], { type });
}

/**
 * Plays the base64 MP3 that /api/process-voice returns, and exposes an interruptible stop.
 *
 * Unlike the pre-migration UI this revokes each object URL when it is replaced or the component
 * unmounts; that version leaked one blob URL per assistant turn for the life of the page.
 */
export function useTtsPlayback() {
  const [isSpeaking, setIsSpeaking] = useState(false);
  const audioRef = useRef<HTMLAudioElement | null>(null);
  const urlRef = useRef<string | null>(null);

  const revoke = useCallback(() => {
    if (urlRef.current) { URL.revokeObjectURL(urlRef.current); urlRef.current = null; }
  }, []);

  const stop = useCallback(() => {
    const audio = audioRef.current;
    if (audio) { audio.pause(); audio.currentTime = 0; }
    setIsSpeaking(false);
  }, []);

  const play = useCallback(async (base64Audio?: string | null) => {
    const audio = audioRef.current;
    if (!base64Audio || !audio) return;

    revoke();
    urlRef.current = URL.createObjectURL(base64ToBlob(base64Audio, 'audio/mpeg'));
    audio.src = urlRef.current;
    setIsSpeaking(true);
    try {
      await audio.play();
    } catch {
      // Autoplay can be blocked before the first user gesture; not worth an error banner.
      setIsSpeaking(false);
    }
  }, [revoke]);

  useEffect(() => () => { stop(); revoke(); }, [revoke, stop]);

  const onEnded = useCallback(() => setIsSpeaking(false), []);

  return { isSpeaking, play, stop, audioRef, onEnded };
}
