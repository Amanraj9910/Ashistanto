/**
 * Text-to-Speech Service with Multi-Voice Support
 * Supports multiple accents and voices using Azure Speech Services
 */

const sdk = require('microsoft-cognitiveservices-speech-sdk');
const fs = require('fs');
const path = require('path');

// Voice mapping for different accents and languages
const VOICE_MAPPING = {
  american: {
    name: 'en-US-JennyNeural',
    language: 'en-US',
    displayName: 'English American Accent',
    style: 'default'
  },
  british: {
    name: 'en-GB-SoniaNeural',
    language: 'en-GB',
    displayName: 'English British Accent',
    style: 'default'
  },
  japanese: {
    name: 'en-US-NancyNeural',
    language: 'en-US',
    displayName: 'English Japanese Accent',
    style: 'default'
  }
};

/**
 * Generate speech from text with specified accent/voice
 * @param {string} text - Text to synthesize
 * @param {string} accent - Accent type: 'american', 'british', 'japanese'
 * @returns {Promise<Buffer>} Audio data as MP3 buffer
 */
async function synthesizeText(text, accent = 'american') {
  return new Promise((resolve, reject) => {
    try {
      const speechKey = process.env.AZURE_SPEECH_KEY;
      const speechRegion = process.env.AZURE_SPEECH_REGION;

      if (!speechKey || !speechRegion) {
        reject(new Error('Azure Speech credentials not configured in .env file'));
        return;
      }

      // Get voice configuration for specified accent
      const voiceConfig = VOICE_MAPPING[accent] || VOICE_MAPPING.american;
      console.log(`  → Initializing TTS with accent: ${voiceConfig.displayName}`);
      console.log(`  → Voice name: ${voiceConfig.name}`);

      // Create speech configuration
      const speechConfig = sdk.SpeechConfig.fromSubscription(speechKey, speechRegion);
      
      // Set voice and language
      speechConfig.speechSynthesisVoiceName = voiceConfig.name;
      speechConfig.speechSynthesisLanguage = voiceConfig.language;
      
      // Output format: MP3 at 16kHz 32-bit rate
      speechConfig.speechSynthesisOutputFormat = 
        sdk.SpeechSynthesisOutputFormat.Audio16Khz32KBitRateMonoMp3;

      // Use null for default speaker (no audio output)
      const synthesizer = new sdk.SpeechSynthesizer(speechConfig, null);

      console.log(`  → Synthesizing text: "${text.substring(0, 50)}..."`);

      synthesizer.speakTextAsync(
        text,
        (result) => {
          if (result.reason === sdk.ResultReason.SynthesizingAudioCompleted) {
            console.log(`  ✓ Speech synthesis completed (${result.audioData.byteLength} bytes)`);
            const audioData = Buffer.from(result.audioData);
            synthesizer.close();
            resolve(audioData);
          } else {
            console.error('  ❌ Speech synthesis failed:', result.errorDetails);
            synthesizer.close();
            reject(new Error('Speech synthesis failed: ' + (result.errorDetails || 'Unknown error')));
          }
        },
        (error) => {
          console.error('  ❌ Speech synthesis error:', error);
          synthesizer.close();
          reject(new Error('Speech synthesis error: ' + error.message));
        }
      );
    } catch (error) {
      console.error('  ❌ Exception during TTS initialization:', error);
      reject(new Error('Failed to initialize speech synthesis: ' + error.message));
    }
  });
}

/**
 * Validate accent selection
 * @param {string} accent - Accent type to validate
 * @returns {boolean} True if valid
 */
function isValidAccent(accent) {
  return accent in VOICE_MAPPING;
}

/**
 * Get all available voice options
 * @returns {Object} Voice options mapping
 */
function getAvailableVoices() {
  return VOICE_MAPPING;
}

/**
 * Get voice info for a specific accent
 * @param {string} accent - Accent type
 * @returns {Object} Voice configuration
 */
function getVoiceInfo(accent) {
  return VOICE_MAPPING[accent] || VOICE_MAPPING.american;
}

module.exports = {
  synthesizeText,
  isValidAccent,
  getAvailableVoices,
  getVoiceInfo,
  VOICE_MAPPING
};
