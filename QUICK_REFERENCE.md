# 🚀 Docker & ACR - Quick Reference Card

## Your Registry Details
```
Name: ashistanto
Server: ashistanto.azurecr.io
Image Path: ashistanto.azurecr.io/azure-voice-ai:latest
```

---

## One-Line Commands

### Build & Push (Automated)
```powershell
node azure-push.js
```

### Deploy to Azure (Automated)
```powershell
.\deploy-azure.ps1 -ResourceGroup "your-rg-name"
```

### Quick Setup (Interactive)
```powershell
.\setup-docker.ps1 -ImageTag "latest"
```

---

## Manual Commands

### Build Docker Image
```powershell
docker build -t ashistanto.azurecr.io/azure-voice-ai:latest .
```

### Test Locally
```powershell
docker run -it -p 3000:3000 `
  -e AZURE_SPEECH_KEY=key `
  ashistanto.azurecr.io/azure-voice-ai:latest
```

### Push to ACR
```powershell
az acr login -n ashistanto
docker push ashistanto.azurecr.io/azure-voice-ai:latest
```

### Deploy to Container Apps
```powershell
az containerapp create `
  --resource-group my-rg `
  --name azure-voice-ai `
  --image ashistanto.azurecr.io/azure-voice-ai:latest `
  --registry-server ashistanto.azurecr.io `
  --registry-username username `
  --registry-password password `
  --target-port 3000 `
  --ingress external
```

### Update Environment Variables
```powershell
az containerapp update -g my-rg -n azure-voice-ai `
  --set-env-vars `
    AZURE_SPEECH_KEY=value `
    AZURE_OPENAI_KEY=value `
    NODE_ENV=production
```

---

## Verification Commands

### Check Image Built
```powershell
docker images | grep azure-voice-ai
```

### Check Image in ACR
```powershell
az acr repository show-tags -n ashistanto --repository azure-voice-ai
```

### Check App Running
```powershell
az containerapp show -g my-rg -n azure-voice-ai
```

### View Logs
```powershell
az containerapp logs show -g my-rg -n azure-voice-ai --follow
```

---

## Common Issues & Fixes

| Issue | Fix |
|-------|-----|
| Docker not found | `choco install docker-desktop` |
| ACR login failed | `az logout` → `az login` → `az acr login -n ashistanto` |
| Build stuck | Press Ctrl+C, then: `docker build --no-cache ...` |
| Container won't start | Check logs: `az containerapp logs show ...` |
| FFmpeg error | Verify image built: `docker run ashistanto.azurecr.io/azure-voice-ai:latest ffmpeg -version` |

---

## Environment Variables Needed

For production deployment, set in Azure:
```
AZURE_SPEECH_KEY=your-key
AZURE_SPEECH_REGION=eastus
AZURE_OPENAI_KEY=your-key
AZURE_OPENAI_ENDPOINT=your-endpoint
AZURE_OPENAI_DEPLOYMENT_ID=your-deployment
NODE_ENV=production
```

---

## File Descriptions

| File | Purpose | Command |
|------|---------|---------|
| **Dockerfile** | Image definition | `docker build -f Dockerfile ...` |
| **azure-push.js** | Build & push | `node azure-push.js` |
| **setup-docker.ps1** | One-command setup | `.\setup-docker.ps1` |
| **deploy-azure.ps1** | Azure deploy | `.\deploy-azure.ps1 -ResourceGroup rg` |

---

## Documentation Files

| File | Contents | Read Time |
|------|----------|-----------|
| **README_DOCKER.md** | Navigation & index | 5 min |
| **DOCKER_QUICK_START.md** | Quick reference | 5 min |
| **SETUP_INSTRUCTIONS.md** | Step-by-step guide | 15 min |
| **DOCKER_SETUP.md** | Detailed docs | 20 min |
| **IMPLEMENTATION_SUMMARY.md** | What was done | 10 min |

---

## Workflow Overview

```
1. Write Code
   ↓
2. docker build (builds Dockerfile)
   ↓
3. docker push (sends to ACR)
   ↓
4. az containerapp update (redeploys)
   ↓
5. App is Live with FFmpeg! ✅
```

---

## Key Points

✅ FFmpeg is included in Docker image  
✅ No more App Service issues  
✅ ACR stores your images  
✅ Container Apps runs your app  
✅ All scripted for automation  

---

## Getting Help

- **Setup Issues**: See `SETUP_INSTRUCTIONS.md` → Troubleshooting
- **Docker Questions**: See `DOCKER_SETUP.md` → Docker Overview
- **Commands Reference**: See `SETUP_INSTRUCTIONS.md` → Commands
- **All Options**: See `DOCKER_SETUP.md` → Detailed guide

---

## Quick Start

```powershell
# Step 1: Build & Push (2 minutes)
node azure-push.js

# Step 2: Deploy (2 minutes)
.\deploy-azure.ps1 -ResourceGroup "your-rg"

# Step 3: Set Variables (1 minute)
az containerapp update -g your-rg -n azure-voice-ai `
  --set-env-vars AZURE_SPEECH_KEY=key AZURE_OPENAI_KEY=key

# Step 4: Get URL (1 minute)
az containerapp show -g your-rg -n azure-voice-ai `
  --query properties.configuration.ingress.fqdn

# Step 5: Test (1 minute)
curl https://your-app-url/api/config
```

**Total Time: ~7 minutes** ⏱️

---

## Useful Links

- Docker: https://docs.docker.com/
- Azure: https://learn.microsoft.com/azure/
- Container Registry: https://learn.microsoft.com/azure/container-registry/
- Container Apps: https://learn.microsoft.com/azure/container-apps/
- FFmpeg: https://ffmpeg.org/

---

## Support Matrix

| Task | Command | Docs |
|------|---------|------|
| Build | `docker build` | DOCKER_SETUP.md |
| Push | `docker push` | DOCKER_SETUP.md |
| Deploy | `az containerapp create` | SETUP_INSTRUCTIONS.md |
| Monitor | `az containerapp logs show` | DOCKER_SETUP.md |
| Update | `az containerapp update` | SETUP_INSTRUCTIONS.md |

---

## Remember

✨ Your ACR: **ashistanto.azurecr.io**  
🐳 Your Image: **azure-voice-ai**  
🚀 Your Apps will have: **FFmpeg included**  

**You're ready to go!** 🎉
