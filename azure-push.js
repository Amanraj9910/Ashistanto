#!/usr/bin/env node

/**
 * Script to build Docker image and push to Azure Container Registry
 * Usage: node azure-push.js [image-tag]
 * 
 * Set these environment variables:
 * - ACR_NAME: Your Azure Container Registry name (e.g., ashistanto)
 * - ACR_SERVER: Your ACR server URL (e.g., ashistanto.azurecr.io)
 * - ACR_USERNAME: Your ACR username
 * - ACR_PASSWORD: Your ACR password
 * 
 * Or run: az acr login -n ashistanto
 */

const { exec } = require('child_process');
const { promisify } = require('util');
const path = require('path');
const os = require('os');

const execAsync = promisify(exec);

// Configuration
const ACR_SERVER = process.env.ACR_SERVER || 'ashistanto.azurecr.io';
const ACR_NAME = process.env.ACR_NAME || 'ashistanto';
const IMAGE_NAME = 'azure-voice-ai';
const IMAGE_TAG = process.argv[2] || 'latest';
const FULL_IMAGE_NAME = `${ACR_SERVER}/${IMAGE_NAME}:${IMAGE_TAG}`;

console.log('🐳 Azure Voice AI - Docker Push Script');
console.log('=====================================\n');
console.log(`📋 Configuration:`);
console.log(`   ACR Server: ${ACR_SERVER}`);
console.log(`   Image: ${FULL_IMAGE_NAME}`);
console.log(`   Dockerfile: ${path.join(__dirname, 'Dockerfile')}`);
console.log('');

async function main() {
  try {
    // Step 1: Validate environment
    console.log('📝 Step 1: Validating environment...');
    
    if (!process.env.ACR_USERNAME && !process.env.ACR_PASSWORD) {
      console.log('⚠️  No ACR credentials found in environment variables.');
      console.log('    Please ensure you\'re logged in with: az acr login -n ' + ACR_NAME);
      console.log('    OR set ACR_USERNAME and ACR_PASSWORD environment variables\n');
    }

    // Step 2: Build Docker image
    console.log('🔨 Step 2: Building Docker image...');
    console.log(`   Running: docker build -t ${FULL_IMAGE_NAME} .`);
    
    try {
      const { stdout, stderr } = await execAsync(
        `docker build -t ${FULL_IMAGE_NAME} .`,
        { cwd: __dirname, maxBuffer: 10 * 1024 * 1024 }
      );
      if (stderr) console.log(stderr);
      console.log('✅ Docker image built successfully!\n');
    } catch (error) {
      console.error('❌ Failed to build Docker image:');
      console.error(error.message);
      process.exit(1);
    }

    // Step 3: Login to ACR (if credentials provided)
    if (process.env.ACR_USERNAME && process.env.ACR_PASSWORD) {
      console.log('🔐 Step 3: Logging into Azure Container Registry...');
      try {
        const loginCmd = `docker login -u ${process.env.ACR_USERNAME} -p ${process.env.ACR_PASSWORD} ${ACR_SERVER}`;
        await execAsync(loginCmd, { cwd: __dirname });
        console.log('✅ Successfully logged into ACR\n');
      } catch (error) {
        console.error('❌ Failed to login to ACR:');
        console.error(error.message);
        process.exit(1);
      }
    } else {
      console.log('⏭️  Step 3: Skipping ACR login (using existing Docker daemon credentials)\n');
    }

    // Step 4: Push image to ACR
    console.log('📤 Step 4: Pushing image to Azure Container Registry...');
    console.log(`   Running: docker push ${FULL_IMAGE_NAME}`);
    
    try {
      const { stdout, stderr } = await execAsync(
        `docker push ${FULL_IMAGE_NAME}`,
        { cwd: __dirname, maxBuffer: 10 * 1024 * 1024 }
      );
      if (stderr) console.log(stderr);
      console.log('✅ Image pushed successfully!\n');
    } catch (error) {
      console.error('❌ Failed to push image to ACR:');
      console.error(error.message);
      process.exit(1);
    }

    // Step 5: Display image info
    console.log('📊 Step 5: Image Information');
    try {
      const { stdout } = await execAsync(
        `docker inspect ${FULL_IMAGE_NAME}`,
        { cwd: __dirname }
      );
      const imageInfo = JSON.parse(stdout)[0];
      console.log(`   Image ID: ${imageInfo.Id.substring(7, 19)}...`);
      console.log(`   Size: ${(imageInfo.Size / 1024 / 1024).toFixed(2)} MB`);
      console.log(`   Created: ${imageInfo.Created}`);
    } catch (error) {
      console.log('   (Could not retrieve image details)');
    }

    console.log('\n✨ Docker image successfully built and pushed to ACR!');
    console.log('\n📝 Next steps:');
    console.log(`   1. Deploy to Azure Container Instances (ACI):`);
    console.log(`      az container create \\`);
    console.log(`        --resource-group <your-rg> \\`);
    console.log(`        --name azure-voice-ai \\`);
    console.log(`        --image ${FULL_IMAGE_NAME} \\`);
    console.log(`        --registry-login-server ${ACR_SERVER} \\`);
    console.log(`        --registry-username <acr-username> \\`);
    console.log(`        --registry-password <acr-password> \\`);
    console.log(`        --environment-variables AZURE_SPEECH_KEY=<key> AZURE_OPENAI_KEY=<key> \\`);
    console.log(`        --ports 3000 \\`);
    console.log(`        --cpu 2 --memory 3.5`);
    console.log(`\n   2. Or deploy to Azure Container Apps:`);
    console.log(`      az containerapp create \\`);
    console.log(`        --resource-group <your-rg> \\`);
    console.log(`        --name azure-voice-ai \\`);
    console.log(`        --image ${FULL_IMAGE_NAME} \\`);
    console.log(`        --registry-server ${ACR_SERVER} \\`);
    console.log(`        --registry-username <acr-username> \\`);
    console.log(`        --registry-password <acr-password>`);
    console.log(`\n   3. Or deploy to Azure Kubernetes Service (AKS)`);
    
  } catch (error) {
    console.error('❌ An unexpected error occurred:');
    console.error(error.message);
    process.exit(1);
  }
}

main();
