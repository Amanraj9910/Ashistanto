#!/usr/bin/env powershell
# Quick setup script for Docker and ACR push
# Usage: .\setup-docker.ps1 -ImageTag "v1.0.0"

param(
    [string]$ImageTag = "latest",
    [switch]$SkipLogin = $false,
    [switch]$LocalOnly = $false
)

$ErrorActionPreference = "Stop"

Write-Host "🐳 Azure Voice AI - Docker Setup" -ForegroundColor Cyan
Write-Host "=================================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$ACR_SERVER = "ashistanto.azurecr.io"
$ACR_NAME = "ashistanto"
$IMAGE_NAME = "azure-voice-ai"
$FULL_IMAGE = "$ACR_SERVER/$IMAGE_NAME`:$ImageTag"

Write-Host "📋 Configuration:" -ForegroundColor Yellow
Write-Host "   Registry: $ACR_NAME ($ACR_SERVER)"
Write-Host "   Image: $FULL_IMAGE"
Write-Host ""

# Step 1: Check Docker
Write-Host "1️⃣  Checking Docker installation..." -ForegroundColor Green
try {
    $dockerVersion = docker --version
    Write-Host "   ✅ $dockerVersion" -ForegroundColor Green
}
catch {
    Write-Host "   ❌ Docker not found. Please install Docker Desktop." -ForegroundColor Red
    exit 1
}

# Step 2: Check Azure CLI
Write-Host "2️⃣  Checking Azure CLI installation..." -ForegroundColor Green
try {
    $azVersion = az --version | Select-Object -First 1
    Write-Host "   ✅ $azVersion" -ForegroundColor Green
}
catch {
    Write-Host "   ⚠️  Azure CLI not found. Some features will be limited." -ForegroundColor Yellow
}

# Step 3: Build Docker image
Write-Host "3️⃣  Building Docker image..." -ForegroundColor Green
Write-Host "   Running: docker build -t $FULL_IMAGE ." -ForegroundColor Gray
try {
    docker build -t $FULL_IMAGE .
    Write-Host "   ✅ Image built successfully" -ForegroundColor Green
}
catch {
    Write-Host "   ❌ Failed to build image: $_" -ForegroundColor Red
    exit 1
}

# Step 4: Test locally
Write-Host "4️⃣  Image Information:" -ForegroundColor Green
try {
    $imageInfo = docker image ls --filter "reference=$FULL_IMAGE" --format "{{.Size}}"
    Write-Host "   ✅ Image size: $imageInfo" -ForegroundColor Green
}
catch {
    Write-Host "   ⚠️  Could not retrieve image info" -ForegroundColor Yellow
}

# Step 5: Push to ACR (optional)
if (-not $LocalOnly) {
    Write-Host "5️⃣  Logging into ACR..." -ForegroundColor Green
    
    if (-not $SkipLogin) {
        try {
            az acr login -n $ACR_NAME
            Write-Host "   ✅ Logged into ACR" -ForegroundColor Green
        }
        catch {
            Write-Host "   ❌ Failed to login to ACR: $_" -ForegroundColor Red
            Write-Host "   💡 Try running: az login" -ForegroundColor Yellow
            exit 1
        }
    }
    else {
        Write-Host "   ⏭️  Skipping login (using existing credentials)" -ForegroundColor Yellow
    }
    
    Write-Host "6️⃣  Pushing image to ACR..." -ForegroundColor Green
    Write-Host "   Running: docker push $FULL_IMAGE" -ForegroundColor Gray
    try {
        docker push $FULL_IMAGE
        Write-Host "   ✅ Image pushed successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "   ❌ Failed to push image: $_" -ForegroundColor Red
        exit 1
    }
}

# Success message
Write-Host ""
Write-Host "✨ Setup Complete!" -ForegroundColor Cyan
Write-Host ""

if ($LocalOnly) {
    Write-Host "📝 Local image created: $FULL_IMAGE" -ForegroundColor Green
    Write-Host "   Run locally: docker run -it -p 3000:3000 $FULL_IMAGE" -ForegroundColor Gray
}
else {
    Write-Host "📝 Image pushed to ACR: $FULL_IMAGE" -ForegroundColor Green
}

Write-Host ""
Write-Host "🚀 Next steps:" -ForegroundColor Yellow
Write-Host "   1. Deploy to Azure:"
Write-Host "      az containerapp create \" -ForegroundColor Gray
Write-Host "        --resource-group <your-rg> \" -ForegroundColor Gray
Write-Host "        --name azure-voice-ai \" -ForegroundColor Gray
Write-Host "        --image $FULL_IMAGE \" -ForegroundColor Gray
Write-Host "        --registry-server $ACR_SERVER" -ForegroundColor Gray
Write-Host ""
Write-Host "   2. Or run locally to test:"
Write-Host "      docker run -it -p 3000:3000 \" -ForegroundColor Gray
Write-Host "        -e AZURE_SPEECH_KEY=your-key \" -ForegroundColor Gray
Write-Host "        -e AZURE_OPENAI_KEY=your-key \" -ForegroundColor Gray
Write-Host "        $FULL_IMAGE" -ForegroundColor Gray
Write-Host ""
Write-Host "📚 See DOCKER_SETUP.md for more details" -ForegroundColor Cyan
