#!/usr/bin/env powershell
# Deploy to Azure Container Instances or Container Apps
# Usage: .\deploy-azure.ps1 -Service "aci" -ResourceGroup "my-rg"

param(
    [ValidateSet("aci", "containerapp")]
    [string]$Service = "containerapp",
    [string]$ResourceGroup,
    [string]$Location = "eastus",
    [string]$ImageTag = "latest",
    [string]$AppName = "azure-voice-ai"
)

$ErrorActionPreference = "Stop"

Write-Host "🚀 Azure Voice AI - Deployment Script" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Configuration
$ACR_SERVER = "ashistanto.azurecr.io"
$ACR_NAME = "ashistanto"
$IMAGE_NAME = "azure-voice-ai"
$FULL_IMAGE = "$ACR_SERVER/$IMAGE_NAME`:$ImageTag"

# Validate resource group
if (-not $ResourceGroup) {
    Write-Host "❌ Resource group is required" -ForegroundColor Red
    Write-Host "   Usage: .\deploy-azure.ps1 -ResourceGroup 'my-rg' -Service 'containerapp'" -ForegroundColor Yellow
    exit 1
}

Write-Host "📋 Deployment Configuration:" -ForegroundColor Yellow
Write-Host "   Service: $Service (ACI/Container Apps)" -ForegroundColor Gray
Write-Host "   Resource Group: $ResourceGroup" -ForegroundColor Gray
Write-Host "   Location: $Location" -ForegroundColor Gray
Write-Host "   App Name: $AppName" -ForegroundColor Gray
Write-Host "   Image: $FULL_IMAGE" -ForegroundColor Gray
Write-Host ""

# Get ACR credentials
Write-Host "1️⃣  Getting ACR credentials..." -ForegroundColor Green
try {
    $credentials = az acr credential show -n $ACR_NAME | ConvertFrom-Json
    $acrUsername = $credentials.username
    $acrPassword = $credentials.passwords[0].value
    Write-Host "   ✅ Retrieved ACR credentials" -ForegroundColor Green
}
catch {
    Write-Host "   ❌ Failed to get ACR credentials: $_" -ForegroundColor Red
    exit 1
}

# Deployment based on service type
if ($Service -eq "aci") {
    Write-Host "2️⃣  Deploying to Azure Container Instances..." -ForegroundColor Green
    
    try {
        az container create `
            --resource-group $ResourceGroup `
            --name $AppName `
            --image $FULL_IMAGE `
            --registry-login-server $ACR_SERVER `
            --registry-username $acrUsername `
            --registry-password $acrPassword `
            --environment-variables `
                NODE_ENV=production `
                PORT=3000 `
            --ports 3000 `
            --cpu 2 --memory 3.5 `
            --restart-policy Always
        
        Write-Host "   ✅ Container created successfully" -ForegroundColor Green
        
        # Get container info
        $container = az container show --resource-group $ResourceGroup --name $AppName | ConvertFrom-Json
        Write-Host ""
        Write-Host "📊 Container Information:" -ForegroundColor Green
        Write-Host "   URL: http://$($container.containers[0].instanceView.currentState.state)" -ForegroundColor Gray
        Write-Host "   IP Address: $($container.ipAddress.ip)" -ForegroundColor Gray
        Write-Host "   Port: $($container.ipAddress.ports[0].port)" -ForegroundColor Gray
    }
    catch {
        Write-Host "   ❌ Deployment failed: $_" -ForegroundColor Red
        exit 1
    }
}
else {
    Write-Host "2️⃣  Deploying to Azure Container Apps..." -ForegroundColor Green
    
    try {
        # Create environment first if it doesn't exist
        $envName = "$($AppName)-env"
        
        try {
            az containerapp env show -g $ResourceGroup -n $envName > $null
            Write-Host "   ℹ️  Using existing environment: $envName" -ForegroundColor Gray
        }
        catch {
            Write-Host "   Creating Container Apps environment..." -ForegroundColor Gray
            az containerapp env create `
                --resource-group $ResourceGroup `
                --name $envName `
                --location $Location
        }
        
        # Create container app
        az containerapp create `
            --resource-group $ResourceGroup `
            --name $AppName `
            --environment $envName `
            --image $FULL_IMAGE `
            --registry-server $ACR_SERVER `
            --registry-username $acrUsername `
            --registry-password $acrPassword `
            --target-port 3000 `
            --ingress external `
            --cpu 0.5 --memory 1.0
        
        Write-Host "   ✅ Container App created successfully" -ForegroundColor Green
        
        # Get app URL
        $app = az containerapp show --resource-group $ResourceGroup --name $AppName | ConvertFrom-Json
        $appUrl = $app.properties.configuration.ingress.fqdn
        
        Write-Host ""
        Write-Host "📊 Container App Information:" -ForegroundColor Green
        Write-Host "   Name: $AppName" -ForegroundColor Gray
        Write-Host "   URL: https://$appUrl" -ForegroundColor Green
        Write-Host "   Resource Group: $ResourceGroup" -ForegroundColor Gray
    }
    catch {
        Write-Host "   ❌ Deployment failed: $_" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "✨ Deployment Complete!" -ForegroundColor Cyan
Write-Host ""
Write-Host "📝 Important - Set Environment Variables:" -ForegroundColor Yellow

if ($Service -eq "aci") {
    Write-Host "   az container create --update-environment-variables `
        AZURE_SPEECH_KEY=your-key `
        AZURE_OPENAI_KEY=your-key `
        ..." -ForegroundColor Gray
}
else {
    Write-Host "   az containerapp update `
        --resource-group $ResourceGroup `
        --name $AppName `
        --set-env-vars `
          AZURE_SPEECH_KEY=your-key `
          AZURE_OPENAI_KEY=your-key" -ForegroundColor Gray
}

Write-Host ""
Write-Host "🔍 View logs:" -ForegroundColor Cyan
if ($Service -eq "aci") {
    Write-Host "   az container logs --resource-group $ResourceGroup --name $AppName" -ForegroundColor Gray
}
else {
    Write-Host "   az containerapp logs show --resource-group $ResourceGroup --name $AppName" -ForegroundColor Gray
}

Write-Host ""
Write-Host "📚 More commands:" -ForegroundColor Cyan
Write-Host "   View details: az containerapp show -g $ResourceGroup -n $AppName" -ForegroundColor Gray
Write-Host "   Update image: az containerapp update -g $ResourceGroup -n $AppName --image $FULL_IMAGE" -ForegroundColor Gray
Write-Host "   Delete app: az containerapp delete -g $ResourceGroup -n $AppName" -ForegroundColor Gray
