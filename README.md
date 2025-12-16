# Azure Voice AI Agent - Complete Setup & Deployment Guide

## 🎯 Project Overview

Azure Voice AI Agent is a production-ready voice-based AI assistant that integrates with Microsoft 365 and Azure cognitive services. Users can interact via voice or text with customizable voice accents, sending emails, scheduling meetings, managing calendars, and more.

### Key Features
- 🎤 **Voice Interaction**: Natural speech-to-text and text-to-speech
- 🌐 **Multi-Accent Support**: American, British, and Japanese English accents
- 📧 **Microsoft 365 Integration**: Email, Calendar, Teams, OneDrive
- 🤖 **AI-Powered**: Azure OpenAI with tool-use capabilities
- 🐳 **Docker Ready**: Production-grade containerization
- ☸️ **Kubernetes Compatible**: Full K8s deployment manifests
- 📊 **Session Management**: User-specific conversation history
- 🔒 **Secure**: Environment-based credential management

## 📋 Project Structure

```
azure-voice-ai/
├── server.js                    # Main Express server
├── auth.js                      # Azure AD authentication
├── agent-tools.js               # Microsoft 365 integration tools
├── graph-tools.js               # Microsoft Graph API wrapper
├── tts-service.js               # Multi-voice TTS service (NEW)
├── package.json                 # Dependencies
├── Dockerfile                   # Docker image definition
├── docker-compose.yml           # Docker Compose setup (NEW)
├── deployment.yaml              # Kubernetes manifests (NEW)
├── public/
│   ├── index.html               # React frontend with voice UI
│   └── img/                     # Images and logos
├── backup/                      # Backup files
├── .env.example                 # Environment template (NEW)
├── PRODUCTION_DEPLOYMENT.md     # Deployment guide (NEW)
├── VOICE_FEATURE_DOCUMENTATION.md # Voice feature details (NEW)
└── README.md                    # This file
```

## 🚀 Quick Start (5 Minutes)

### Prerequisites
- Node.js 18+
- Docker & Docker Compose (for containerized deployment)
- Azure subscription with credentials
- FFmpeg (automatically installed in Docker)

### Local Development

1. **Clone & Setup**
```bash
cd azure-voice-ai
cp .env.example .env
# Edit .env with your Azure credentials
```

2. **Install Dependencies**
```bash
npm install
```

3. **Start Development Server**
```bash
npm start
```

4. **Access Application**
```
http://localhost:3000
```

### Docker Deployment (Recommended)

1. **Configure Environment**
```bash
cp .env.example .env
# Edit .env with Azure credentials
```

2. **Build & Run**
```bash
docker-compose build
docker-compose up -d
```

3. **Check Status**
```bash
docker-compose ps
docker-compose logs -f
```

4. **Access Application**
```
http://localhost:3000
```

## 🔑 Azure Credentials Setup

### 1. Azure Speech Services (STT/TTS)

**Steps:**
1. Go to [Azure Portal](https://portal.azure.com)
2. Create "Speech" resource
3. Copy **Key** and **Region**

**Provided Credentials:**
```

```

### 2. Azure OpenAI (Language Model)

**Steps:**
1. Create "Azure OpenAI" resource
2. Deploy model (e.g., gpt-4o-mini)
3. Get endpoint and key

**.env Configuration:**
```env
AZURE_OPENAI_KEY=your-key
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_DEPLOYMENT=gpt-4o-mini
```

### 3. Azure Active Directory (Microsoft 365)

**Steps:**
1. Go to Azure Active Directory > App Registrations
2. Register new application
3. Add credentials (Client Secret)
4. Configure API Permissions:
   - Mail.Send
   - Calendar.Create
   - Chat.ReadWrite
   - User.Read

**.env Configuration:**
```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-secret
APP_REDIRECT_URI=http://localhost:3000/auth/callback
```

## 🎙️ Voice & Accent Selection (NEW)

### Available Voices
1. **American English** (Default)
   - Voice: en-US-JennyNeural
   - Use Case: General, professional

2. **British English**
   - Voice: en-GB-SoniaNeural
   - Use Case: UK/European audience

3. **Japanese Accented English**
   - Voice: en-US-NancyNeural
   - Use Case: Japanese English learners

### How to Use
1. Click 🌐 globe icon in navigation bar
2. Select preferred accent
3. Send voice or text message
4. AI responds with selected voice accent

## 📱 API Documentation

### Get Available Voices
```bash
GET /api/voices
```

### Process Voice Message
```bash
POST /api/process-voice
Content-Type: multipart/form-data

audio: [audio file]
sessionId: [session-uuid]
accent: american|british|japanese
```

### Send Text Message
```bash
POST /api/text-message
Content-Type: application/json

{
  "text": "Send email to John",
  "sessionId": "session-uuid",
  "accent": "british"
}
```

## 🐳 Docker Deployment

### Build Docker Image
```bash
docker build -t azure-voice-ai:latest .
```

### Run with Docker Compose
```bash
docker-compose up -d

# View logs
docker-compose logs -f

# Stop services
docker-compose down
```

### Run with Docker CLI
```bash
docker run -it \
  -e AZURE_SPEECH_KEY=your-key \
  -e AZURE_SPEECH_REGION=eastus \
  -e AZURE_OPENAI_KEY=your-key \
  -e AZURE_OPENAI_ENDPOINT=your-endpoint \
  -e AZURE_OPENAI_DEPLOYMENT=gpt-4o-mini \
  -e AZURE_TENANT_ID=your-tenant \
  -e AZURE_CLIENT_ID=your-client \
  -e AZURE_CLIENT_SECRET=your-secret \
  -p 3000:3000 \
  azure-voice-ai:latest
```

## ☸️ Kubernetes Deployment

### Prerequisites
- AKS cluster running
- kubectl configured
- ACR (Azure Container Registry) or Docker Hub

### Deploy to AKS
```bash
# Create namespace
kubectl create namespace azure-voice-ai

# Update deployment.yaml with your credentials
nano deployment.yaml

# Deploy
kubectl apply -f deployment.yaml

# Check status
kubectl get deployment -n azure-voice-ai
kubectl get pods -n azure-voice-ai
kubectl logs -n azure-voice-ai deployment/azure-voice-ai

# Port forward
kubectl port-forward -n azure-voice-ai svc/azure-voice-ai-service 3000:80
```

## 🔧 Configuration Options

### Environment Variables

**Core Settings:**
```env
NODE_ENV=production
PORT=3000
DOCKER_ENV=true
DEFAULT_ACCENT=american
```

**Azure Speech Services:**
```env
AZURE_SPEECH_KEY=your-speech-key
AZURE_SPEECH_REGION=eastus
```

**Azure OpenAI:**
```env
AZURE_OPENAI_KEY=your-openai-key
AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
AZURE_OPENAI_DEPLOYMENT=gpt-4o-mini
```

**Microsoft 365:**
```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-secret
APP_REDIRECT_URI=http://localhost:3000/auth/callback
```

## 🧪 Testing

### Test Configuration
```bash
curl http://localhost:3000/api/config
```

### Test Available Voices
```bash
curl http://localhost:3000/api/voices
```

### Test Text-to-Speech
```bash
curl -X POST http://localhost:3000/api/text-message \
  -H "Content-Type: application/json" \
  -d '{
    "text": "Hello world",
    "sessionId": "test-123",
    "accent": "british"
  }'
```

## 📊 Performance Metrics

### Processing Times
- Speech-to-Text: 500ms - 1s
- Text-to-Speech: 300-800ms per 100 characters
- OpenAI Response: 1-3s
- Total round-trip: 2-5s

### System Requirements
- **CPU**: 0.5-1 core recommended
- **Memory**: 256-512MB
- **Storage**: 100MB for application
- **Network**: Minimum 1Mbps

### Scaling
- Horizontal: Load balance across multiple instances
- Vertical: Increase CPU/Memory allocation
- Sessions: Use Redis/Database for persistence

## 🔒 Security Best Practices

1. **Environment Variables**
   - Never commit .env file
   - Use Azure Key Vault in production
   - Rotate keys regularly

2. **Docker Security**
   - Run as non-root user
   - Use read-only root filesystem
   - Scan images for vulnerabilities

3. **HTTPS**
   - Enable SSL/TLS certificates
   - Use ACME/Let's Encrypt
   - Redirect HTTP to HTTPS

4. **Authentication**
   - OAuth 2.0 with Azure AD
   - Session tokens validated
   - Timeout after 24 hours

5. **API Security**
   - Rate limiting enabled
   - CORS restrictions
   - Request size limits (50MB)

## 📚 Documentation Files

- **[PRODUCTION_DEPLOYMENT.md](./PRODUCTION_DEPLOYMENT.md)** - Comprehensive deployment guide
- **[VOICE_FEATURE_DOCUMENTATION.md](./VOICE_FEATURE_DOCUMENTATION.md)** - Voice feature details
- **[QUICK_REFERENCE.md](./QUICK_REFERENCE.md)** - Quick command reference
- **[Graph-setup.md](./Graph-setup.md)** - Microsoft Graph setup

## 🐛 Troubleshooting

### Common Issues

**Issue: "Azure credentials not configured"**
```bash
# Check .env file
cat .env

# Verify variables are set
echo $AZURE_SPEECH_KEY
```

**Issue: "FFmpeg not found"**
```bash
# Check FFmpeg installation
which ffmpeg

# Install FFmpeg (Ubuntu/Debian)
apt-get install ffmpeg

# Install FFmpeg (macOS)
brew install ffmpeg

# Windows: Download from https://ffmpeg.org/download.html
```

**Issue: "No speech detected"**
- Speak louder and more clearly
- Check microphone permissions
- Ensure proper audio input device

**Issue: "Invalid session"**
- Clear browser cookies
- Logout and login again
- Check sessionId is valid

## 🚀 Deployment Checklist

- [ ] Copy .env.example to .env
- [ ] Fill in all Azure credentials
- [ ] Test locally: `npm start`
- [ ] Build Docker image: `docker build -t azure-voice-ai:latest .`
- [ ] Test Docker: `docker run -it -p 3000:3000 azure-voice-ai:latest`
- [ ] Configure docker-compose.yml
- [ ] Deploy to Docker Compose: `docker-compose up -d`
- [ ] Verify health check: `curl http://localhost:3000/api/config`
- [ ] Test voice features via UI
- [ ] Monitor logs: `docker-compose logs -f`
- [ ] Set up monitoring/alerts (optional)
- [ ] Configure backup/restore procedures

## 📞 Support & Resources

- **Azure Speech Services**: https://docs.microsoft.com/azure/cognitive-services/speech-service/
- **Azure OpenAI**: https://learn.microsoft.com/azure/ai-services/openai/
- **Microsoft Graph API**: https://docs.microsoft.com/graph/
- **Docker**: https://docs.docker.com/
- **Kubernetes**: https://kubernetes.io/docs/

## 🤝 Contributing

To add new voice accents or features:

1. **Add Voice to tts-service.js**:
```javascript
const VOICE_MAPPING = {
  spanish: {
    name: 'es-ES-HelenaNeural',
    language: 'es-ES',
    displayName: 'Spanish Accent',
    style: 'default'
  }
};
```

2. **Update Frontend** (index.html):
```javascript
const voiceOptions = {
  spanish: {
    label: 'Spanish Accent',
    voice: 'es-ES-HelenaNeural',
    language: 'es-ES'
  }
};
```

3. **Test and Deploy**:
```bash
npm test
docker-compose up -d
```

## 📝 License

Proprietary - Hosho Digital

## 👥 Authors

- **Hosho Digital** - Main Development
- **Azure Team** - Infrastructure & Services

## 🎉 Version History

### v1.1.0 (Current)
- ✅ Multi-voice accent support
- ✅ Production Docker setup
- ✅ Kubernetes manifests
- ✅ Enhanced documentation
- ✅ Voice feature UI

### v1.0.0
- Initial release
- Single voice support
- Basic functionality

---

**Last Updated:** December 16, 2024
**Status:** Production Ready ✅
**Maintained By:** Hosho Digital

## 🎯 Next Steps

1. **Immediate**: Configure .env with your Azure credentials
2. **Short-term**: Deploy locally via Docker Compose
3. **Medium-term**: Test with real users and collect feedback
4. **Long-term**: Deploy to Azure (ACI, App Service, or AKS)

Good luck! 🚀
