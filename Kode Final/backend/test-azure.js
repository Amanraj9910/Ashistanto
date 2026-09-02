require('dotenv').config();
const sdk = require('microsoft-cognitiveservices-speech-sdk');

console.log('\n╔═══════════════════════════════════════════╗');
console.log('║   Azure Services Diagnostic Test          ║');
console.log('╚═══════════════════════════════════════════╝\n');

// Test 1: Check Environment Variables
console.log('1️⃣  Checking Environment Variables...');
console.log('  AZURE_SPEECH_KEY:', process.env.AZURE_SPEECH_KEY ? '✓ Set' : '✗ Not set');
console.log('  AZURE_SPEECH_REGION:', process.env.AZURE_SPEECH_REGION || '✗ Not set');
console.log('  AZURE_OPENAI_ENDPOINT:', process.env.AZURE_OPENAI_ENDPOINT ? '✓ Set' : '✗ Not set');
console.log('  AZURE_OPENAI_KEY:', process.env.AZURE_OPENAI_KEY ? '✓ Set' : '✗ Not set');
console.log('  AZURE_OPENAI_DEPLOYMENT:', process.env.AZURE_OPENAI_DEPLOYMENT || 'gpt-4o-mini (default)');

if (!process.env.AZURE_SPEECH_KEY || !process.env.AZURE_SPEECH_REGION) {
  console.log('\n❌ Azure Speech credentials missing!');
  console.log('Please add them to your .env file:\n');
  console.log('AZURE_SPEECH_KEY=your_key_here');
  console.log('AZURE_SPEECH_REGION=your_region_here\n');
  process.exit(1);
}

// Test 2: Test Speech Service Connection
console.log('\n2️⃣  Testing Azure Speech Service Connection...');
try {
  const speechConfig = sdk.SpeechConfig.fromSubscription(
    process.env.AZURE_SPEECH_KEY,
    process.env.AZURE_SPEECH_REGION
  );
  console.log('  ✓ Speech config created successfully');
  console.log('  Region:', process.env.AZURE_SPEECH_REGION);
} catch (error) {
  console.log('  ❌ Failed to create speech config:', error.message);
  process.exit(1);
}

// Test 3: Test Text-to-Speech
console.log('\n3️⃣  Testing Text-to-Speech...');
const speechConfig = sdk.SpeechConfig.fromSubscription(
  process.env.AZURE_SPEECH_KEY,
  process.env.AZURE_SPEECH_REGION
);
speechConfig.speechSynthesisVoiceName = 'en-US-JennyNeural';

const synthesizer = new sdk.SpeechSynthesizer(speechConfig, null);

synthesizer.speakTextAsync(
  'Testing Azure Speech Services',
  (result) => {
    if (result.reason === sdk.ResultReason.SynthesizingAudioCompleted) {
      console.log('  ✓ Text-to-Speech working! Generated', result.audioData.byteLength, 'bytes');
    } else {
      console.log('  ❌ Text-to-Speech failed:', result.errorDetails);
    }
    synthesizer.close();
    
    console.log('\n4️⃣  Speech Recognition Test');
    console.log('  Note: Speech recognition requires actual audio input.');
    console.log('  The web app will test this when you record audio.\n');
    
    console.log('╔═══════════════════════════════════════════╗');
    console.log('║   Diagnostic Complete                      ║');
    console.log('╚═══════════════════════════════════════════╝');
    console.log('\nIf all tests passed, your Azure Speech Service is configured correctly!');
    console.log('Start the server with: npm start\n');
  },
  (error) => {
    console.log('  ❌ Text-to-Speech error:', error);
    synthesizer.close();
    process.exit(1);
  }
);