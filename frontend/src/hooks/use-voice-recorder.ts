'use client';

import { useCallback, useEffect, useRef, useState } from 'react';

// Tuning lifted verbatim from the pre-migration UI (public/legacy/index.html) so recording
// behaviour is unchanged. THRESHOLD was raised from 15 to 30 there to cope with background noise.
const SILENCE_THRESHOLD = 30;
const SILENCE_DURATION_MS = 2000;
const MIN_BLOB_BYTES = 1000;
const ANALYSER_FFT_SIZE = 512;

type UseVoiceRecorderOptions = {
  onRecorded: (audio: Blob) => void;
  onError?: (message: string) => void;
};

/**
 * Microphone capture with automatic stop after a period of silence.
 *
 * The silence detector runs a requestAnimationFrame loop over an AnalyserNode and stops the
 * recorder once the average frequency magnitude stays below SILENCE_THRESHOLD for
 * SILENCE_DURATION_MS. A separate 100ms interval drives `silenceCountdown` for the UI.
 */
export function useVoiceRecorder({ onRecorded, onError }: UseVoiceRecorderOptions) {
  const [isRecording, setIsRecording] = useState(false);
  const [silenceCountdown, setSilenceCountdown] = useState(0);

  const recorderRef = useRef<MediaRecorder | null>(null);
  const streamRef = useRef<MediaStream | null>(null);
  const chunksRef = useRef<Blob[]>([]);
  const audioContextRef = useRef<AudioContext | null>(null);
  const rafRef = useRef<number | null>(null);
  const countdownRef = useRef<ReturnType<typeof setInterval> | null>(null);
  const silenceStartRef = useRef<number>(Date.now());

  // Callbacks live in refs so the rAF loop and recorder events never read a stale closure.
  const onRecordedRef = useRef(onRecorded);
  const onErrorRef = useRef(onError);
  useEffect(() => { onRecordedRef.current = onRecorded; }, [onRecorded]);
  useEffect(() => { onErrorRef.current = onError; }, [onError]);

  const clearTimers = useCallback(() => {
    if (rafRef.current !== null) { cancelAnimationFrame(rafRef.current); rafRef.current = null; }
    if (countdownRef.current !== null) { clearInterval(countdownRef.current); countdownRef.current = null; }
    setSilenceCountdown(0);
  }, []);

  const releaseStream = useCallback(() => {
    streamRef.current?.getTracks().forEach((track) => track.stop());
    streamRef.current = null;

    const context = audioContextRef.current;
    audioContextRef.current = null;
    if (context && context.state !== 'closed') {
      // Deferred: closing while the rAF loop may still hold a reference throws in some browsers.
      setTimeout(() => { context.close().catch(() => { /* already closed - safe to ignore */ }); }, 100);
    }
  }, []);

  const stop = useCallback(() => {
    // Read the recorder's own state rather than `isRecording`; React state is stale inside
    // the rAF loop and the recorder event handlers.
    if (recorderRef.current?.state === 'recording') recorderRef.current.stop();
  }, []);

  const start = useCallback(async () => {
    if (recorderRef.current?.state === 'recording') return;

    let stream: MediaStream;
    try {
      stream = await navigator.mediaDevices.getUserMedia({
        audio: { echoCancellation: true, noiseSuppression: true, autoGainControl: true, sampleRate: 16000 }
      });
    } catch {
      onErrorRef.current?.('Microphone access was denied. Please allow microphone access and try again.');
      return;
    }
    streamRef.current = stream;

    let mimeType = 'audio/webm;codecs=opus';
    if (typeof MediaRecorder === 'undefined' || !MediaRecorder.isTypeSupported(mimeType)) mimeType = 'audio/webm';

    const recorder = new MediaRecorder(stream, { mimeType, audioBitsPerSecond: 128000 });
    recorderRef.current = recorder;
    chunksRef.current = [];

    // ─── Silence detection rig ───
    const AudioCtor: typeof AudioContext | undefined =
      window.AudioContext || (window as unknown as { webkitAudioContext?: typeof AudioContext }).webkitAudioContext;
    let analyser: AnalyserNode | null = null;
    let bins: Uint8Array<ArrayBuffer> | null = null;

    if (AudioCtor) {
      const context = new AudioCtor();
      audioContextRef.current = context;
      analyser = context.createAnalyser();
      analyser.fftSize = ANALYSER_FFT_SIZE;
      context.createMediaStreamSource(stream).connect(analyser);
      bins = new Uint8Array(new ArrayBuffer(analyser.frequencyBinCount));
    }

    const checkSilence = () => {
      if (!analyser || !bins || recorderRef.current?.state !== 'recording') { clearTimers(); return; }

      analyser.getByteFrequencyData(bins);
      let total = 0;
      for (let i = 0; i < bins.length; i += 1) total += bins[i];
      const average = total / bins.length;

      if (average < SILENCE_THRESHOLD) {
        if (Date.now() - silenceStartRef.current > SILENCE_DURATION_MS) { clearTimers(); stop(); return; }
      } else {
        silenceStartRef.current = Date.now();
      }
      rafRef.current = requestAnimationFrame(checkSilence);
    };

    recorder.ondataavailable = (event) => { if (event.data.size > 0) chunksRef.current.push(event.data); };

    recorder.onstart = () => {
      silenceStartRef.current = Date.now();
      if (analyser) {
        rafRef.current = requestAnimationFrame(checkSilence);
        countdownRef.current = setInterval(() => {
          const remaining = SILENCE_DURATION_MS - (Date.now() - silenceStartRef.current);
          setSilenceCountdown(remaining > 0 ? Math.ceil(remaining / 1000) : 0);
        }, 100);
      }
    };

    recorder.onstop = () => {
      clearTimers();
      const audio = new Blob(chunksRef.current, { type: mimeType });
      chunksRef.current = [];
      releaseStream();
      setIsRecording(false);

      if (audio.size < MIN_BLOB_BYTES) {
        onErrorRef.current?.('Recording is too short. Please speak for at least one second.');
        return;
      }
      onRecordedRef.current(audio);
    };

    recorder.start(100);
    setIsRecording(true);
  }, [clearTimers, releaseStream, stop]);

  useEffect(() => () => {
    if (recorderRef.current?.state === 'recording') recorderRef.current.stop();
    clearTimers();
    releaseStream();
  }, [clearTimers, releaseStream]);

  return { isRecording, silenceCountdown, start, stop };
}
