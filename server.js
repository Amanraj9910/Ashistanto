require('dotenv').config();
const express = require('express');
const multer = require('multer');
const sdk = require('microsoft-cognitiveservices-speech-sdk');
const { AzureOpenAI } = require('openai');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const ffmpeg = require('fluent-ffmpeg');
const wav = require('wav');
const graphTools = require('./graph-tools');
const ttsService = require('./tts-service');

// Set ffmpeg path based on environment
let ffmpegPath;
if (process.env.NODE_ENV === 'production' || process.env.DOCKER_ENV) {
  // In Docker and Azure Linux, ffmpeg is in PATH
  ffmpeg.setFfmpegPath('ffmpeg');
  console.log('✓ Using system FFmpeg (Docker/Linux)');
} else {
  // For local Windows development
  ffmpegPath = process.env.FFMPEG_PATH || 'C:\\ffmpeg\\bin\\ffmpeg.exe';
  ffmpeg.setFfmpegPath(ffmpegPath);
  console.log(`✓ Using custom FFmpeg path: ${ffmpegPath}`);
}

// Import authentication router
const { router: authRouter, userTokenStore } = require('./auth');

const app = express();
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 } // 50MB limit
});

app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Mount auth routes
app.use('/auth', authRouter);

// Get config endpoint
app.get('/api/config', (req, res) => {
  res.json({
    configured: !!(process.env.AZURE_SPEECH_KEY && process.env.AZURE_OPENAI_KEY)
  });
});

// Get available voice accents/languages
app.get('/api/voices', (req, res) => {
  const voices = ttsService.getAvailableVoices();
  const formattedVoices = Object.entries(voices).map(([key, value]) => ({
    id: key,
    name: value.displayName,
    voiceName: value.name,
    language: value.language
  }));

  res.json({
    voices: formattedVoices,
    default: 'american'
  });
});

// Clear conversation history endpoint
app.post('/api/clear-session', express.json(), (req, res) => {
  const sessionId = req.body.sessionId;
  if (sessionId && conversationSessions.has(sessionId)) {
    conversationSessions.delete(sessionId);
    console.log(`✓ Session ${sessionId} cleared`);
  }
  res.json({ success: true });
});

// Debug endpoint - view conversation history
app.get('/api/debug/sessions', (req, res) => {
  const sessions = {};
  for (const [sessionId, history] of conversationSessions.entries()) {
    sessions[sessionId] = {
      messageCount: history.length,
      messages: history
    };
  }
  res.json(sessions);
});

// Get user profile info (name, email, etc)
app.get('/api/user-profile', async (req, res) => {
  try {
    console.log('\n🔍 [/api/user-profile] Request received');

    const sessionId = req.query.sessionId;
    console.log('🔍 [/api/user-profile] Session ID:', sessionId);

    if (!sessionId || !userTokenStore.has(sessionId)) {
      console.error('❌ [/api/user-profile] No valid session found');
      console.error('❌ [/api/user-profile] Session ID provided:', sessionId);
      console.error('❌ [/api/user-profile] Session exists in store:', userTokenStore.has(sessionId));
      return res.status(401).json({ error: 'No valid session' });
    }

    console.log('✅ [/api/user-profile] Valid session found, fetching profile...');

    // Pass sessionId for automatic token refresh
    const profileInfo = await graphTools.getSenderProfile(null, sessionId);

    console.log('✅ [/api/user-profile] Profile info received:', {
      displayName: profileInfo.displayName,
      email: profileInfo.email,
      firstName: profileInfo.displayName.split(' ')[0]
    });

    const response = {
      displayName: profileInfo.displayName,
      email: profileInfo.email,
      firstName: profileInfo.displayName.split(' ')[0]
    };

    console.log('✅ [/api/user-profile] Sending response to frontend:', response);
    res.json(response);
  } catch (err) {
    console.error('❌ [/api/user-profile] Error fetching user profile:', err);
    console.error('❌ [/api/user-profile] Error stack:', err.stack);
    res.status(500).json({ error: 'Failed to fetch user profile' });
  }
});

// Get user profile photo
app.get('/api/user-photo', async (req, res) => {
  try {
    const sessionId = req.query.sessionId;
    console.log('📷 Photo request for session:', sessionId);

    if (!sessionId || !userTokenStore.has(sessionId)) {
      console.warn('❌ Invalid session for photo request');
      return res.status(401).json({ error: 'No valid session' });
    }

    console.log('📷 Fetching photo with sessionId...');

    // Pass sessionId for automatic token refresh
    const photoBuffer = await graphTools.getUserProfilePhoto(null, sessionId);
    console.log('📷 Photo buffer returned, type:', typeof photoBuffer, 'length:', photoBuffer ? photoBuffer.length : 'null');

    if (!photoBuffer) {
      console.warn('⚠️ No photo buffer returned');
      return res.status(404).json({ error: 'No profile photo found' });
    }

    if (Buffer.isBuffer(photoBuffer)) {
      console.log('✅ Photo is a proper Buffer, size:', photoBuffer.length);
    } else {
      console.warn('⚠️ Photo is not a Buffer, type:', typeof photoBuffer);
    }

    res.set('Content-Type', 'image/jpeg');
    res.send(photoBuffer);
  } catch (err) {
    console.error('❌ Error fetching user photo:', err);
    res.status(500).json({ error: 'Failed to fetch user photo' });
  }
});





// Store conversation history per session (in production, use database)
const conversationSessions = new Map();

// Endpoint to process voice interaction
app.post('/api/process-voice', upload.single('audio'), async (req, res) => {
  const tempWebm = path.join(__dirname, `temp_${Date.now()}.webm`);
  const tempWav = path.join(__dirname, `temp_${Date.now()}.wav`);

  try {
    console.log('\n=== Voice Processing Started ===');

    if (!req.file) {
      throw new Error('No audio file uploaded');
    }

    // Get or create session ID
    const sessionId = req.body.sessionId || 'default';
    const accent = req.body.accent || 'american';
    const language = req.body.language || 'en-US';

    if (!conversationSessions.has(sessionId)) {
      conversationSessions.set(sessionId, []);
    }

    const audioBuffer = req.file.buffer;
    console.log('✓ Audio received:', {
      size: audioBuffer.length,
      type: req.file.mimetype,
      sessionId: sessionId,
      accent: accent,
      language: language
    });

    // Validate accent
    if (!ttsService.isValidAccent(accent)) {
      throw new Error(`Invalid accent: ${accent}. Valid options: american, british, japanese`);
    }

    // Save WebM file
    fs.writeFileSync(tempWebm, audioBuffer);
    console.log('✓ Audio saved to temp file:', tempWebm);

    // Convert WebM to WAV using ffmpeg
    console.log('🔄 Converting audio to WAV format...');
    await convertToWav(tempWebm, tempWav);
    console.log('✓ Audio converted to WAV:', tempWav);

    // Read WAV file
    const wavBuffer = fs.readFileSync(tempWav);
    console.log('✓ WAV file loaded, size:', wavBuffer.length);

    // Step 1: Speech-to-Text
    console.log('🎤 Starting speech-to-text...');
    const transcript = await speechToText(wavBuffer);
    console.log('✓ Transcript:', transcript);

    if (!transcript || transcript.trim() === '') {
      throw new Error('No speech detected in the audio. Please speak louder and try again.');
    }

    // Step 2: Query Azure OpenAI Agent
    console.log('🤖 Querying AI agent...');
    const conversationHistory = conversationSessions.get(sessionId);

    // Don't pass token object - pass sessionId for automatic token refresh
    const agentResponse = await queryAgent(transcript, conversationHistory, sessionId, null);
    console.log('✓ Agent Response:', agentResponse);

    // Check if response is an action_preview (skip TTS for confirmations)
    let isActionPreview = false;
    let parsedPreview = null;
    try {
      if (typeof agentResponse === 'string' && agentResponse.startsWith('{')) {
        parsedPreview = JSON.parse(agentResponse);
        if (parsedPreview.type === 'action_preview') {
          isActionPreview = true;
          console.log('🔔 Action preview detected - skipping TTS');
        }
      }
    } catch (e) {
      // Not JSON, continue normally
    }

    // Step 3: Text-to-Speech with selected accent (skip for action previews)
    let audioData = null;
    if (!isActionPreview) {
      console.log(`🔊 Generating speech with ${ttsService.getVoiceInfo(accent).displayName}...`);
      audioData = await ttsService.synthesizeText(agentResponse, accent);
      console.log('✓ Audio generated, size:', audioData.length);
    } else {
      // For action previews, speak a brief confirmation message
      const confirmMessage = 'I need your confirmation before proceeding. Please check the preview.';
      console.log(`🔊 Generating confirmation speech...`);
      audioData = await ttsService.synthesizeText(confirmMessage, accent);
      console.log('✓ Confirmation audio generated, size:', audioData.length);
    }

    // Clean up temp files
    [tempWebm, tempWav].forEach(file => {
      if (fs.existsSync(file)) {
        fs.unlinkSync(file);
      }
    });

    res.json({
      transcript,
      agentResponse,
      audioData: audioData.toString('base64'),
      sessionId: sessionId
    });
  } catch (error) {
    console.error('❌ Error:', error.message);

    // Clean up temp files on error
    [tempWebm, tempWav].forEach(file => {
      if (fs.existsSync(file)) {
        try { fs.unlinkSync(file); } catch (e) { }
      }
    });

    res.status(500).json({
      error: error.message || 'Unknown error occurred'
    });
  }
});

// Convert audio to WAV format using ffmpeg
function convertToWav(inputPath, outputPath) {
  return new Promise((resolve, reject) => {
    ffmpeg(inputPath)
      .toFormat('wav')
      .audioFrequency(16000)
      .audioChannels(1)
      .audioCodec('pcm_s16le')
      .on('end', () => {
        console.log('  ✓ FFmpeg conversion completed');
        resolve();
      })
      .on('error', (err) => {
        console.error('  ❌ FFmpeg error:', err.message);
        reject(new Error('Audio conversion failed: ' + err.message));
      })
      .save(outputPath);
  });
}

// Speech-to-Text using Azure Speech Services (Single-shot recognition)
async function speechToText(wavBuffer) {
  return new Promise((resolve, reject) => {
    let recognizer = null;

    try {
      const speechKey = process.env.AZURE_SPEECH_KEY;
      const speechRegion = process.env.AZURE_SPEECH_REGION;

      if (!speechKey || !speechRegion) {
        reject(new Error('Azure Speech credentials not configured in .env file'));
        return;
      }

      console.log('  → Initializing Azure Speech SDK...');

      // 🔍 DIAGNOSTICS: Analyze WAV file
      console.log('  📊 Audio Diagnostics:');
      console.log(`     Total size: ${wavBuffer.length} bytes`);
      console.log(`     Header size: 44 bytes`);
      console.log(`     PCM data size: ${wavBuffer.length - 44} bytes`);

      // Read WAV header info
      const sampleRate = wavBuffer.readUInt32LE(24);
      const bitsPerSample = wavBuffer.readUInt16LE(34);
      const numChannels = wavBuffer.readUInt16LE(22);
      const duration = (wavBuffer.length - 44) / (sampleRate * numChannels * (bitsPerSample / 8));

      console.log(`     Sample rate: ${sampleRate} Hz`);
      console.log(`     Channels: ${numChannels}`);
      console.log(`     Bits per sample: ${bitsPerSample}`);
      console.log(`     Duration: ${duration.toFixed(2)} seconds`);

      // Check if audio is too short
      if (duration < 0.5) {
        console.log('  ⚠️ WARNING: Audio is very short (< 0.5s)');
      }

      // Analyze audio levels
      const pcmData = wavBuffer.slice(44);
      let maxAmplitude = 0;
      for (let i = 0; i < Math.min(pcmData.length, 10000); i += 2) {
        const sample = Math.abs(pcmData.readInt16LE(i));
        if (sample > maxAmplitude) maxAmplitude = sample;
      }
      const volumePercent = (maxAmplitude / 32768 * 100).toFixed(1);
      console.log(`     Max volume: ${volumePercent}%`);

      if (maxAmplitude < 1000) {
        console.log('  ⚠️ WARNING: Audio level is very low - speak louder!');
      }

      const speechConfig = sdk.SpeechConfig.fromSubscription(speechKey, speechRegion);
      speechConfig.speechRecognitionLanguage = 'en-US';

      console.log('  → Creating audio config from WAV buffer...');

      // Create audio stream from WAV buffer
      const pushStream = sdk.AudioInputStream.createPushStream();

      // Skip WAV header (first 44 bytes) and push the raw PCM data
      pushStream.write(pcmData);
      pushStream.close();

      const audioConfig = sdk.AudioConfig.fromStreamInput(pushStream);
      recognizer = new sdk.SpeechRecognizer(speechConfig, audioConfig);

      console.log('  → Starting single-shot recognition...');

      // Single-shot recognition - recognizes once and returns
      recognizer.recognizeOnceAsync(
        (result) => {
          recognizer.close();

          if (result.reason === sdk.ResultReason.RecognizedSpeech) {
            console.log(`  ✓ Recognized: "${result.text}"`);
            resolve(result.text);
          } else if (result.reason === sdk.ResultReason.NoMatch) {
            console.log('  ⚠ No speech could be recognized');
            console.log(`  💡 Possible reasons:`);
            console.log(`     - Audio too short (${duration.toFixed(2)}s)`);
            console.log(`     - Volume too low (${volumePercent}%)`);
            console.log(`     - Background noise masking speech`);
            console.log(`     - Speaking too fast or unclear`);
            reject(new Error('No speech was detected. Please speak clearly and try again.'));
          } else if (result.reason === sdk.ResultReason.Canceled) {
            const cancellation = sdk.CancellationDetails.fromResult(result);
            console.error(`  ❌ Canceled: ${cancellation.reason}`);
            if (cancellation.reason === sdk.CancellationReason.Error) {
              reject(new Error(`Speech recognition error: ${cancellation.errorDetails}`));
            } else {
              reject(new Error('Speech recognition was canceled.'));
            }
          }
        },
        (err) => {
          recognizer.close();
          console.error('  ❌ Recognition error:', err);
          reject(new Error('Speech recognition failed: ' + err));
        }
      );

    } catch (error) {
      console.error('  ❌ Exception:', error.message);
      if (recognizer) recognizer.close();
      reject(new Error('Speech recognition initialization failed: ' + error.message));
    }
  });
}

// Query Azure OpenAI Agent
// Replace the queryAgent function in your server.js with this updated version
// Query Azure OpenAI Agent
// Replace the queryAgent function in your server.js with this updated version

async function queryAgent(text, conversationHistory = [], sessionId = 'default', userToken = null) {
  try {
    const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
    const apiKey = process.env.AZURE_OPENAI_KEY;
    const deployment = process.env.AZURE_OPENAI_DEPLOYMENT || 'gpt-4o-mini';

    if (!endpoint || !apiKey) {
      throw new Error('Azure OpenAI credentials not configured in .env file');
    }

    console.log('  → Sending request to Azure OpenAI...');
    console.log('  → Conversation history length:', conversationHistory.length);
    console.log('  → User token available:', !!userToken);

    const client = new AzureOpenAI({
      endpoint: endpoint,
      apiKey: apiKey,
      apiVersion: '2024-08-01-preview',
      deployment: deployment
    });

    // Load agent tools
    const { tools, executeTool } = require('./agent-tools');

    // Get current date for context
    const now = new Date();
    const currentDate = now.toISOString().split('T')[0];
    const currentTime = now.toTimeString().split(' ')[0];

    // Build messages array with conversation history
    const messages = [
      {
        role: 'system',
        content: `You are a helpful AI voice assistant with access to Microsoft 365 services.

CURRENT DATE & TIME: ${currentDate} ${currentTime}
Today is: ${now.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}

================================================================================
📅 MEETING RULE — DEFAULT TO TEAMS MEETING
================================================================================
When the user schedules ANY meeting, ALWAYS set:
isTeamsMeeting = true

UNLESS the user clearly says:
- "offline meeting"
- "in person"
- "physical room"
- "not on teams"
- "not online"

So by default, meetings must be created as Teams meetings.

================================================================================
📤 MEETING LINK SHARING RULE
================================================================================
After creating a Teams meeting, tell the user:
"The Teams meeting link will be shared automatically with all participants."

================================================================================
📧 EMAIL SENDING - PROFESSIONAL FORMATTING:
================================================================================
When user wants to send email, ALWAYS use the send_email tool.

Examples:
- "send mail to jatin raj about app issue" → send_email(recipient_name="jatin raj", subject="Application Issue", body="...")
- "email vansh about meeting" → send_email(recipient_name="vansh", subject="Meeting", body="...")

The system will automatically:
✓ Add professional greeting with recipient's name
✓ Format body with HTML styling
✓ Include your signature with name, title, contact
✓ Add footer with date

YOU MUST:
✓ Call send_email tool for ANY email request
✓ Use recipient's name (not email address)
✓ Create clear subject
✓ Write natural body text

⚠️ CONTACT NOT FOUND HANDLING:
If send_email, send_teams_message, or search_contact_email returns notFound=true or found=false:
1. Tell the user the contact was not found
2. Ask them to verify the spelling of the name
3. OR ask them to provide the email address directly
Example response: "I couldn't find anyone named 'aman' in the directory. Could you check the spelling or provide their email address?"

================================================================================
📅 MEETING SCHEDULING - WITH MULTIPLE ATTENDEES:
================================================================================
When user wants to schedule meeting, ALWAYS use the create_calendar_event tool.

DATE/TIME CALCULATIONS:
- Today: ${currentDate}
- "3 PM today" → ${currentDate}T15:00:00
- "tomorrow 2 PM" → ${new Date(now.getTime() + 86400000).toISOString().split('T')[0]}T14:00:00
- "10 AM" → current_date + T10:00:00
- "2:30 PM" → current_date + T14:30:00

DURATION DEFAULTS:
- "30 min" → add 30 minutes to start time
- "1 hour" → add 60 minutes to start time
- No duration specified → default 1 hour

📌 DEFAULT TEAM MEETING (MOST IMPORTANT UPDATE)
SET:
isTeamsMeeting = true
UNLESS user explicitly denies.

Examples:
User: "schedule meet with jatin raj 3 PM today for 30 min"
YOU MUST CALL: create_calendar_event(
  subject="Meeting with Jatin Raj",
  start="${currentDate}T15:00:00",
  end="${currentDate}T15:30:00",
  attendeeNames=["jatin raj"],
  isTeamsMeeting=true
)

User: "set up teams call with john and sarah tomorrow 10 AM"
YOU MUST CALL: create_calendar_event(
  subject="Teams Meeting",
  start="${new Date(now.getTime() + 86400000).toISOString().split('T')[0]}T10:00:00",
  end="${new Date(now.getTime() + 86400000).toISOString().split('T')[0]}T11:00:00",
  attendeeNames=["john", "sarah"],
  isTeamsMeeting=true
)

================================================================================
🗑️ DELETION FEATURES:
================================================================================
DELETE EMAIL:
- "delete the email I just sent" → delete_sent_email()
- "delete the email about meeting" → delete_sent_email(subject="meeting")
- "delete the email to priyanshu" → delete_sent_email(recipient_email="priyanshu")
  ⚠️ IMPORTANT: Use the recipient NAME directly (not email address). The system will match by name OR email.

DELETE CALENDAR EVENT:
- "delete the meeting with raj" → delete_calendar_event(subject="raj")
- "cancel the standup meeting" → delete_calendar_event(subject="standup")

DELETE TEAMS MESSAGE:
- "delete the teams message I just sent" → delete_teams_message()
- "delete the hello message" → delete_teams_message(message_content="hello")
- "delete the message to priyanshu" → delete_teams_message()
  ⚠️ CRITICAL: Call delete_teams_message() DIRECTLY. Do NOT just call get_teams_messages to find it.
  The delete function will automatically find and delete your most recent message.

================================================================================
📁 FILE SEARCH FORMATTING:
================================================================================
When searching for files using search_files:
1. Show the file NAME, LOCATION, SIZE, and LAST MODIFIED DATE
2. ALWAYS show the folder path/location using the "breadcrumb" or "location" field
3. Format like: "You can find this file at: Documents → Projects → FileName.docx"
4. DO NOT use markdown bold (**text**) - use plain text only
5. Include the Open link for user to access

Example response format (plain text, no markdown):
"I found the file 'Report.docx':
- Location: Documents → Work → Reports → Report.docx
- Size: 1.5 MB
- Last Modified: December 17, 2025
- Click here to open: [link]"

================================================================================
🚨 CRITICAL: YOU MUST USE TOOLS!
================================================================================
- For email requests → CALL send_email tool
- For meeting requests → CALL create_calendar_event tool
- For calendar questions → CALL get_calendar_events tool
- For email questions → CALL get_recent_emails tool
- For deleting emails → CALL delete_sent_email tool
- For deleting meetings → CALL delete_calendar_event tool
- For Teams messages → CALL send_teams_message or delete_teams_message tool
- For file search → CALL search_files tool

DO NOT just respond with text. ALWAYS call the appropriate tool when user requests an action.

⚠️ RESPONSE FORMATTING:
- DO NOT use markdown bold (**text**) in responses
- DO NOT use asterisks for emphasis
- Use plain text and line breaks for formatting
- Keep voice responses short (1-2 sentences) after tool execution.
`
      }
    ];

    // Add conversation history
    conversationHistory.forEach(msg => {
      messages.push(msg);
    });

    // Add current user message
    messages.push({
      role: 'user',
      content: text
    });

    console.log('  → Tools available:', tools.length);
    console.log('  → Tool names:', tools.map(t => t.function.name).join(', '));

    // First call - AI decides if it needs to use tools
    let result = await client.chat.completions.create({
      model: deployment,
      messages: messages,
      tools: tools,
      tool_choice: 'auto', // Let AI decide when to use tools
      max_tokens: 500,
      temperature: 0.7
    });

    let responseMessage = result.choices[0].message;
    console.log('  → AI response:', {
      hasToolCalls: !!responseMessage.tool_calls,
      toolCallCount: responseMessage.tool_calls?.length || 0,
      content: responseMessage.content || '(no content)'
    });

    // Check if AI wants to use tools
    if (responseMessage.tool_calls && responseMessage.tool_calls.length > 0) {
      console.log('  → AI requesting tool execution:', responseMessage.tool_calls.map(tc => tc.function.name).join(', '));

      // Add AI's response to messages
      messages.push(responseMessage);

      // Execute all requested tools
      for (const toolCall of responseMessage.tool_calls) {
        const functionName = toolCall.function.name;
        const functionArgs = JSON.parse(toolCall.function.arguments);

        console.log(`  → Executing ${functionName} with args: `, JSON.stringify(functionArgs, null, 2));

        try {
          // ✅ FIXED: Pass userToken and sessionId to executeTool
          const toolResult = await executeTool(functionName, functionArgs, userToken, sessionId);
          console.log(`  ✓ ${functionName} completed: `, toolResult);

          // 🔍 Check if this is an action preview - if so, return immediately without AI processing
          if (toolResult && typeof toolResult === 'object' && toolResult.type === 'action_preview') {
            console.log('  🔔 Action preview detected - returning to user without further AI processing');
            return JSON.stringify({
              type: 'action_preview',
              preview: toolResult.preview,
              message: toolResult.message
            });
          }

          // Add tool result to messages
          messages.push({
            role: 'tool',
            tool_call_id: toolCall.id,
            content: JSON.stringify(toolResult)
          });
        } catch (error) {
          console.error(`  ✗ ${functionName} failed: `, error.message);
          messages.push({
            role: 'tool',
            tool_call_id: toolCall.id,
            content: JSON.stringify({ error: error.message })
          });
        }
      }

      // Second call - AI formulates final response with tool results
      console.log('  → Getting final response from AI...');
      result = await client.chat.completions.create({
        model: deployment,
        messages: messages,
        max_tokens: 300,
        temperature: 0.7
      });

      responseMessage = result.choices[0].message;
      console.log('  ✓ Final response:', responseMessage.content);
    } else {
      console.log('  ℹ No tools were called by AI');
    }

    // Update conversation history
    conversationHistory.push({
      role: 'user',
      content: text
    });
    conversationHistory.push({
      role: 'assistant',
      content: responseMessage.content
    });

    // Keep only last 10 exchanges (20 messages)
    if (conversationHistory.length > 20) {
      conversationHistory.splice(0, conversationHistory.length - 20);
    }

    console.log('  ✓ Received response from AI');
    return responseMessage.content;
  } catch (error) {
    console.error('  ✗ OpenAI error:', error.message);
    if (error.response) {
      console.error('  ✗ Error details:', error.response.data);
    }
    throw new Error('Failed to get AI response: ' + error.message);
  }
}

// Text-to-Speech using Azure Speech Services
async function textToSpeech(text) {
  // Deprecated: Use ttsService.synthesizeText() instead
  // This function is kept for backwards compatibility only
  console.warn('⚠️  textToSpeech() is deprecated, use ttsService.synthesizeText() instead');
  return ttsService.synthesizeText(text, 'american');
}

// Endpoint to process text messages
app.post('/api/text-message', express.json(), async (req, res) => {
  try {
    const { text, sessionId, language, accent } = req.body;

    if (!text || !text.trim()) {
      return res.status(400).json({ error: 'Text message is required' });
    }

    if (!sessionId) {
      return res.status(400).json({ error: 'Session ID is required' });
    }

    const selectedAccent = accent || 'american';

    // Validate accent
    if (!ttsService.isValidAccent(selectedAccent)) {
      return res.status(400).json({ error: `Invalid accent: ${selectedAccent} ` });
    }

    console.log('\n=== Text Message Received ===');
    console.log(`✓ Message: "${text}"`);
    console.log(`✓ Session ID: ${sessionId} `);
    console.log(`✓ Accent: ${selectedAccent} `);

    // Verify session exists
    if (!userTokenStore.has(sessionId)) {
      return res.status(401).json({ error: 'Invalid or expired session' });
    }

    // Query the AI agent with the text message (pass null for userToken, sessionId for automatic refresh)
    const response = await queryAgent(text, conversationSessions.get(sessionId) || [], sessionId, null);

    // Get or create conversation history for this session
    if (!conversationSessions.has(sessionId)) {
      conversationSessions.set(sessionId, []);
    }
    const conversationHistory = conversationSessions.get(sessionId);

    // Add to conversation history
    conversationHistory.push({
      role: 'user',
      content: text
    });
    conversationHistory.push({
      role: 'assistant',
      content: response
    });

    // Keep only last 20 messages
    if (conversationHistory.length > 20) {
      conversationHistory.splice(0, conversationHistory.length - 20);
    }

    console.log('✓ Response generated successfully');

    res.json({
      success: true,
      response: response,
      sessionId: sessionId
    });

  } catch (error) {
    console.error('❌ Error processing text message:', error.message);
    res.status(500).json({
      error: error.message || 'Failed to process text message'
    });
  }
});

// ============================================
// 🔍 ACTION PREVIEW ENDPOINT
// ============================================
app.post('/api/preview-action', async (req, res) => {
  try {
    const { sessionId, actionType, actionData } = req.body;

    if (!sessionId || !actionType || !actionData) {
      return res.status(400).json({
        error: 'Missing required parameters: sessionId, actionType, actionData'
      });
    }

    // Check session
    const userToken = userTokenStore.get(sessionId);
    if (!userToken) {
      return res.status(401).json({ error: 'Invalid or expired session' });
    }

    // Import action preview module
    const { actionPreview } = require('./agent-tools');

    // Create preview
    const preview = actionPreview.createActionPreview(actionType, actionData, sessionId);

    console.log(`✓ Preview created for ${actionType}: `, preview.actionId);

    res.json({
      success: true,
      preview: preview
    });
  } catch (error) {
    console.error('❌ Error creating action preview:', error.message);
    res.status(500).json({
      error: error.message || 'Failed to create action preview'
    });
  }
});

// ============================================
// ✅ ACTION CONFIRMATION ENDPOINT
// ============================================
app.post('/api/confirm-action', async (req, res) => {
  try {
    const { sessionId, actionId, userChoice, edits } = req.body;

    if (!sessionId || !actionId || !userChoice) {
      return res.status(400).json({
        error: 'Missing required parameters: sessionId, actionId, userChoice'
      });
    }

    // Check session
    const userToken = userTokenStore.get(sessionId);
    if (!userToken) {
      return res.status(401).json({ error: 'Invalid or expired session' });
    }

    // Import action preview and agent tools modules
    const { actionPreview, executeTool } = require('./agent-tools');

    // Handle user choice
    if (userChoice === 'edit') {
      // Apply edits to pending action
      if (edits) {
        actionPreview.editPendingAction(actionId, edits);
      }
      const updatedAction = actionPreview.getActionForExecution(actionId);
      return res.json({
        success: true,
        message: 'Action edited successfully',
        action: updatedAction
      });
    }

    if (userChoice === 'confirm') {
      // First confirm the action in the store
      const confirmResult = actionPreview.confirmAction(actionId, { confirmed: true });
      if (!confirmResult.success) {
        return res.status(404).json({
          error: confirmResult.error || 'Action not found or already processed'
        });
      }

      // Get the confirmed action data
      const pendingActionData = actionPreview.getPendingAction(actionId);
      if (!pendingActionData) {
        return res.status(404).json({
          error: 'Action not found or expired'
        });
      }

      // Use edited data if available, otherwise use original data
      const actionData = pendingActionData.editedData || pendingActionData.originalData;
      const actionType = pendingActionData.actionType;

      // ✅ OPTIMIZATION: Get cached validated recipient data
      const validatedRecipientData = pendingActionData.validatedRecipientData || null;
      if (validatedRecipientData) {
        console.log(`  ⚡ Using cached recipient data for fast execution`);
      }

      // Execute the action with skipConfirmation=true to avoid infinite loop
      try {
        let result;
        if (actionType === 'send_email') {
          result = await executeTool('send_email', {
            recipient_name: actionData.recipientName,
            subject: actionData.subject,
            body: actionData.body,
            cc_recipients: actionData.ccRecipients || []
          }, userToken, sessionId, true);  // skipConfirmation = true

          // ✅ OPTIMIZATION: Pass cached data directly to sendEmail
          if (validatedRecipientData) {
            const graphTools = require('./graph-tools');
            result = await graphTools.sendEmail(
              actionData.recipientName,
              actionData.subject,
              actionData.body,
              actionData.ccRecipients || [],
              userToken,
              validatedRecipientData  // Pass cached data
            );
          }
        } else if (actionType === 'send_teams_message') {
          result = await executeTool('send_teams_message', {
            recipient_name: actionData.recipientName,
            message: actionData.message
          }, userToken, sessionId, true);  // skipConfirmation = true

          // ✅ OPTIMIZATION: Pass cached data directly to sendTeamsMessage
          if (validatedRecipientData) {
            const graphTools = require('./graph-tools');
            result = await graphTools.sendTeamsMessage(
              actionData.recipientName,
              actionData.message,
              userToken,
              validatedRecipientData  // Pass cached data
            );
          }
        } else if (actionType === 'delete_sent_email') {
          // Execute deletion with cached message ID
          const graphTools = require('./graph-tools');
          result = await graphTools.deleteEmail(actionData.messageId, userToken);
        } else if (actionType === 'delete_teams_message') {
          // Execute deletion with cached chat and message IDs
          const graphTools = require('./graph-tools');
          const client = await graphTools.getGraphClient(userToken);
          await client
            .api(`/chats/${actionData.chatId}/messages/${actionData.messageId}/softDelete`)
            .post({});
          result = {
            success: true,
            message: 'Teams message deleted successfully',
            deletedMessageId: actionData.messageId
          };
        }

        // Clear the action after successful execution
        actionPreview.clearAction(actionId);

        console.log(`✓ Action executed: ${actionType} `);

        res.json({
          success: true,
          message: actionType === 'send_email'
            ? `Email sent successfully to ${actionData.recipientName} `
            : actionType === 'send_teams_message'
              ? `Teams message sent to ${actionData.recipientName} `
              : actionType === 'delete_sent_email'
                ? `Email deleted successfully`
                : `Teams message deleted successfully`,
          result: result
        });
      } catch (executionError) {
        console.error('❌ Error executing action:', executionError.message);
        res.status(500).json({
          error: 'Failed to execute action: ' + executionError.message
        });
      }
    } else if (userChoice === 'cancel') {
      console.log(`✓ Action cancelled: ${actionId} `);
      res.json({
        success: true,
        message: 'Action cancelled'
      });
    } else {
      return res.status(400).json({
        error: 'Invalid userChoice. Must be "confirm", "edit", or "cancel"'
      });
    }
  } catch (error) {
    console.error('❌ Error confirming action:', error.message);
    res.status(500).json({
      error: error.message || 'Failed to process action confirmation'
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`
╔═══════════════════════════════════════════╗
║   Azure Voice AI Agent Server Running     ║
╚═══════════════════════════════════════════╝
  
  🌐 URL: http://localhost:${PORT}
  
  Configuration Status:
  ${process.env.AZURE_SPEECH_KEY ? '✓' : '✗'} Azure Speech Service ${process.env.AZURE_SPEECH_REGION ? `(${process.env.AZURE_SPEECH_REGION})` : ''}
  ${process.env.AZURE_OPENAI_KEY ? '✓' : '✗'} Azure OpenAI ${process.env.AZURE_OPENAI_DEPLOYMENT ? `(${process.env.AZURE_OPENAI_DEPLOYMENT})` : ''}
  
  ${!process.env.AZURE_SPEECH_KEY || !process.env.AZURE_OPENAI_KEY ?
      '⚠️  Please configure your .env file with Azure credentials\n' : '✓ All services configured - Ready to use!\n'
    }
    `);
});