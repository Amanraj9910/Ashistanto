#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Build and push Azure Voice AI Docker image to Azure Container Registry.

.DESCRIPTION
    This script builds the Docker image with FFmpeg support and pushes it to
    Azure Container Registry (ACR) for deployment to Azure App Service.

.PARAMETER Tag
    Image tag to use (default: latest)

.PARAMETER SkipPush
    Build only, don't push to ACR

.EXAMPLE
    .\push-to-acr.ps1
    .\push-to-acr.ps1 -Tag "v1.0.0"
    .\push-to-acr.ps1 -SkipPush

.NOTES
    Prerequisites:
    - Docker Desktop installed and running
    - Azure CLI installed (for ACR login)
    - Run 'az login' if not already authenticated
#>

param(
    [string]$Tag = "latest",
    [switch]$SkipPush = $false
)

# Configuration
$ACR_NAME = "Ashistanto"
$ACR_SERVER = "ashistanto.azurecr.io"
$IMAGE_NAME = "azure-voice-ai"
$FULL_IMAGE = "${ACR_SERVER}/${IMAGE_NAME}:${Tag}"

$ErrorActionPreference = "Stop"

# ============================================================================
# Helper Functions
# ============================================================================

function Write-Header {
    param([string]$Text)
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " $Text" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}

function Write-Step {
    param([int]$Number, [string]$Text)
    Write-Host ""
    Write-Host "[$Number] $Text" -ForegroundColor Yellow
}

function Write-Success {
    param([string]$Text)
    Write-Host "    ✅ $Text" -ForegroundColor Green
}

function Write-Info {
    param([string]$Text)
    Write-Host "    ℹ️  $Text" -ForegroundColor Gray
}

function Write-Fail {
    param([string]$Text)
    Write-Host "    ❌ $Text" -ForegroundColor Red
}

# ============================================================================
# Main Script
# ============================================================================

Write-Header "Azure Voice AI - Docker Build & Push"
Write-Host ""
Write-Host "Configuration:" -ForegroundColor White
Write-Host "  Registry:    $ACR_NAME ($ACR_SERVER)"
Write-Host "  Image:       $FULL_IMAGE"
Write-Host "  Skip Push:   $SkipPush"

# Step 1: Check Docker
Write-Step 1 "Checking Docker installation..."
try {
    $dockerVersion = docker --version 2>&1
    if ($LASTEXITCODE -ne 0) { throw "Docker not found" }
    Write-Success "Docker installed: $dockerVersion"
    
    # Check if Docker daemon is running
    docker info 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Docker daemon not running" }
    Write-Success "Docker daemon is running"
}
catch {
    Write-Fail "Docker is not available. Please install Docker Desktop and ensure it's running."
    exit 1
}

# Step 2: Build Docker image
Write-Step 2 "Building Docker image..."
Write-Info "This may take a few minutes on first build..."

try {
    docker build -t $FULL_IMAGE .
    if ($LASTEXITCODE -ne 0) { throw "Docker build failed" }
    Write-Success "Image built successfully: $FULL_IMAGE"
}
catch {
    Write-Fail "Failed to build Docker image: $_"
    exit 1
}

# Step 3: Verify FFmpeg in container
Write-Step 3 "Verifying FFmpeg installation in container..."
try {
    $ffmpegVersion = docker run --rm $FULL_IMAGE ffmpeg -version 2>&1 | Select-Object -First 1
    if ($LASTEXITCODE -ne 0) { throw "FFmpeg not found in container" }
    Write-Success "FFmpeg verified: $ffmpegVersion"
}
catch {
    Write-Fail "FFmpeg verification failed: $_"
    exit 1
}

# Step 4: Get image size
Write-Step 4 "Image information..."
try {
    $imageInfo = docker images $FULL_IMAGE --format "{{.Size}}"
    Write-Success "Image size: $imageInfo"
}
catch {
    Write-Info "Could not retrieve image info"
}

# If SkipPush, exit here
if ($SkipPush) {
    Write-Header "Build Complete (Push Skipped)"
    Write-Host ""
    Write-Host "To run the container locally:" -ForegroundColor Yellow
    Write-Host "  docker run -p 3000:3000 --env-file .env $FULL_IMAGE" -ForegroundColor Gray
    Write-Host ""
    exit 0
}

# Step 5: Login to Azure Container Registry
Write-Step 5 "Logging into Azure Container Registry..."
try {
    az acr login --name $ACR_NAME 2>&1 | Out-Null
    if ($LASTEXITCODE -ne 0) { 
        Write-Info "ACR login failed. Trying 'az login' first..."
        az login
        az acr login --name $ACR_NAME
        if ($LASTEXITCODE -ne 0) { throw "ACR login failed" }
    }
    Write-Success "Logged into ACR: $ACR_NAME"
}
catch {
    Write-Fail "Failed to login to ACR. Please run 'az login' and try again."
    Write-Host ""
    Write-Host "Troubleshooting:" -ForegroundColor Yellow
    Write-Host "  1. Run: az login"
    Write-Host "  2. Run: az acr login --name $ACR_NAME"
    Write-Host "  3. Re-run this script"
    exit 1
}

# Step 6: Push to ACR
Write-Step 6 "Pushing image to Azure Container Registry..."
Write-Info "This may take several minutes depending on your connection..."

try {
    docker push $FULL_IMAGE
    if ($LASTEXITCODE -ne 0) { throw "Docker push failed" }
    Write-Success "Image pushed successfully!"
}
catch {
    Write-Fail "Failed to push image: $_"
    exit 1
}

# Step 7: Verify image in ACR
Write-Step 7 "Verifying image in ACR..."
try {
    $tags = az acr repository show-tags --name $ACR_NAME --repository $IMAGE_NAME --output tsv 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Success "Image verified in ACR"
        Write-Host "    Available tags: $($tags -join ', ')" -ForegroundColor Gray
    }
}
catch {
    Write-Info "Could not verify image in ACR (this is normal if repository is new)"
}

# Success summary
Write-Header "Deployment Complete!"
Write-Host ""
Write-Host "Image successfully pushed to:" -ForegroundColor Green
Write-Host "  $FULL_IMAGE" -ForegroundColor White
Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "  1. Update your Azure App Service container settings:"
Write-Host "     - Image: $FULL_IMAGE"
Write-Host "     - Registry: $ACR_SERVER"
Write-Host ""
Write-Host "  2. Or deploy via Azure CLI:" -ForegroundColor Yellow
Write-Host "     az webapp config container set \" -ForegroundColor Gray
Write-Host "       --name <your-app-name> \" -ForegroundColor Gray
Write-Host "       --resource-group <your-rg> \" -ForegroundColor Gray
Write-Host "       --docker-custom-image-name $FULL_IMAGE \" -ForegroundColor Gray
Write-Host "       --docker-registry-server-url https://$ACR_SERVER" -ForegroundColor Gray
Write-Host ""
