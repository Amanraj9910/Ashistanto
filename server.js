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
    if (!conversationSessions.has(sessionId)) {
      conversationSessions.set(sessionId, []);
    }
    
    const audioBuffer = req.file.buffer;
    console.log('✓ Audio received:', {
      size: audioBuffer.length,
      type: req.file.mimetype,
      sessionId: sessionId
    });

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
    const userToken = userTokenStore.get(sessionId);
    
    if (!userToken) {
      console.warn('⚠️  No user token found for session:', sessionId);
    }
    
    const agentResponse = await queryAgent(transcript, conversationHistory, sessionId, userToken);
    console.log('✓ Agent Response:', agentResponse);

    // Step 3: Text-to-Speech
    console.log('🔊 Generating speech...');
    const audioData = await textToSpeech(agentResponse);
    console.log('✓ Audio generated, size:', audioData.length);

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
        try { fs.unlinkSync(file); } catch (e) {}
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

// Speech-to-Text using Azure Speech Services
async function speechToText(wavBuffer) {
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      if (!isResolved) {
        isResolved = true;
        if (recognizer) recognizer.close();
        reject(new Error('Speech recognition timeout. Please try speaking again.'));
      }
    }, 20000);

    let isResolved = false;
    let recognizer = null;
    
    try {
      const speechKey = process.env.AZURE_SPEECH_KEY;
      const speechRegion = process.env.AZURE_SPEECH_REGION;

      if (!speechKey || !speechRegion) {
        clearTimeout(timeout);
        reject(new Error('Azure Speech credentials not configured in .env file'));
        return;
      }

      console.log('  → Initializing Azure Speech SDK...');
      const speechConfig = sdk.SpeechConfig.fromSubscription(speechKey, speechRegion);
      speechConfig.speechRecognitionLanguage = 'en-US';
      speechConfig.outputFormat = sdk.OutputFormat.Detailed;

      console.log('  → Creating audio config from WAV buffer...');
      
      // Create audio stream from WAV buffer
      const pushStream = sdk.AudioInputStream.createPushStream();
      
      // Skip WAV header (first 44 bytes) and push the raw PCM data
      const pcmData = wavBuffer.slice(44);
      pushStream.write(pcmData);
      pushStream.close();

      const audioConfig = sdk.AudioConfig.fromStreamInput(pushStream);
      recognizer = new sdk.SpeechRecognizer(speechConfig, audioConfig);

      console.log('  → Starting recognition...');

      let fullText = '';

      recognizer.recognizing = (s, e) => {
        if (e.result.text) {
          console.log(`  → Recognizing: "${e.result.text}"`);
        }
      };

      recognizer.recognized = (s, e) => {
        if (e.result.reason === sdk.ResultReason.RecognizedSpeech) {
          console.log(`  ✓ Recognized: "${e.result.text}"`);
          fullText += e.result.text + ' ';
        } else if (e.result.reason === sdk.ResultReason.NoMatch) {
          console.log('  ⚠ No speech could be recognized');
        }
      };

      recognizer.canceled = (s, e) => {
        console.log(`  ❌ Canceled: ${e.reason}`);
        
        if (e.reason === sdk.CancellationReason.Error) {
          console.error(`  ❌ Error details: ${e.errorDetails}`);
          if (!isResolved) {
            clearTimeout(timeout);
            isResolved = true;
            recognizer.close();
            reject(new Error(`Speech recognition failed: ${e.errorDetails}`));
          }
        }
      };

      recognizer.sessionStopped = (s, e) => {
        console.log('  → Session stopped');
        if (!isResolved) {
          clearTimeout(timeout);
          isResolved = true;
          recognizer.close();
          
          const trimmedText = fullText.trim();
          if (trimmedText) {
            resolve(trimmedText);
          } else {
            reject(new Error('No speech was detected. Please check your microphone and speak clearly.'));
          }
        }
      };

      // Start continuous recognition
      recognizer.startContinuousRecognitionAsync(
        () => {
          console.log('  ✓ Recognition started successfully');
          
          // Stop after 15 seconds
          setTimeout(() => {
            if (!isResolved) {
              console.log('  → Stopping recognition...');
              recognizer.stopContinuousRecognitionAsync(
                () => console.log('  ✓ Recognition stopped'),
                (err) => console.error('  ❌ Error stopping recognition:', err)
              );
            }
          }, 15000);
        },
        (err) => {
          console.error('  ❌ Failed to start recognition:', err);
          if (!isResolved) {
            clearTimeout(timeout);
            isResolved = true;
            reject(new Error('Failed to start speech recognition: ' + err));
          }
        }
      );

    } catch (error) {
      clearTimeout(timeout);
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

📧 EMAIL SENDING - PROFESSIONAL FORMATTING:
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

📅 MEETING SCHEDULING - WITH MULTIPLE ATTENDEES:
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

Examples:
User: "schedule meet with jatin raj 3 PM today for 30 min"
YOU MUST CALL: create_calendar_event(
  subject="Meeting with Jatin Raj",
  start="${currentDate}T15:00:00",
  end="${currentDate}T15:30:00",
  attendeeNames=["jatin raj"],
  isTeamsMeeting=false
)

User: "set up teams call with john and sarah tomorrow 10 AM"
YOU MUST CALL: create_calendar_event(
  subject="Teams Meeting",
  start="${new Date(now.getTime() + 86400000).toISOString().split('T')[0]}T10:00:00",
  end="${new Date(now.getTime() + 86400000).toISOString().split('T')[0]}T11:00:00",
  attendeeNames=["john", "sarah"],
  isTeamsMeeting=true
)

🚨 CRITICAL: YOU MUST USE TOOLS!
- For email requests → CALL send_email tool
- For meeting requests → CALL create_calendar_event tool
- For calendar questions → CALL get_calendar_events tool
- For email questions → CALL get_recent_emails tool

DO NOT just respond with text. ALWAYS call the appropriate tool when user requests an action.

Keep voice responses short (1-2 sentences) after tool execution.`
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
        
        console.log(`  → Executing ${functionName} with args:`, JSON.stringify(functionArgs, null, 2));
        
        try {
          // ✅ FIXED: Pass userToken as third parameter to executeTool
          const toolResult = await executeTool(functionName, functionArgs, userToken);
          console.log(`  ✓ ${functionName} completed:`, toolResult);
          
          // Add tool result to messages
          messages.push({
            role: 'tool',
            tool_call_id: toolCall.id,
            content: JSON.stringify(toolResult)
          });
        } catch (error) {
          console.error(`  ✗ ${functionName} failed:`, error.message);
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
  return new Promise((resolve, reject) => {
    try {
      const speechKey = process.env.AZURE_SPEECH_KEY;
      const speechRegion = process.env.AZURE_SPEECH_REGION;

      if (!speechKey || !speechRegion) {
        reject(new Error('Azure Speech credentials not configured in .env file'));
        return;
      }

      console.log('  → Initializing text-to-speech...');

      const speechConfig = sdk.SpeechConfig.fromSubscription(speechKey, speechRegion);
      speechConfig.speechSynthesisVoiceName = 'en-US-JennyNeural';
      speechConfig.speechSynthesisOutputFormat = 
        sdk.SpeechSynthesisOutputFormat.Audio16Khz32KBitRateMonoMp3;

      const synthesizer = new sdk.SpeechSynthesizer(speechConfig, null);

      synthesizer.speakTextAsync(
        text,
        (result) => {
          if (result.reason === sdk.ResultReason.SynthesizingAudioCompleted) {
            console.log('  ✓ Speech synthesis completed');
            const audioData = Buffer.from(result.audioData);
            synthesizer.close();
            resolve(audioData);
          } else {
            console.error('  ❌ Speech synthesis failed:', result.errorDetails);
            synthesizer.close();
            reject(new Error('Speech synthesis failed: ' + result.errorDetails));
          }
        },
        (error) => {
          console.error('  ❌ Speech synthesis error:', error);
          synthesizer.close();
          reject(new Error('Speech synthesis error: ' + error));
        }
      );
    } catch (error) {
      reject(new Error('Failed to initialize speech synthesis: ' + error.message));
    }
  });
}

// Endpoint to process text messages
app.post('/api/text-message', express.json(), async (req, res) => {
  try {
    const { text, sessionId } = req.body;

    if (!text || !text.trim()) {
      return res.status(400).json({ error: 'Text message is required' });
    }

    if (!sessionId) {
      return res.status(400).json({ error: 'Session ID is required' });
    }

    console.log('\n=== Text Message Received ===');
    console.log(`✓ Message: "${text}"`);
    console.log(`✓ Session ID: ${sessionId}`);

    // Retrieve user token from session store
    const userToken = userTokenStore.get(sessionId);
    if (!userToken) {
      return res.status(401).json({ error: 'Invalid or expired session' });
    }

    // Query the AI agent with the text message
    const response = await queryAgent(text, conversationSessions.get(sessionId) || [], sessionId, userToken);

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
    '⚠️  Please configure your .env file with Azure credentials\n' : '✓ All services configured - Ready to use!\n'}
  `);
});