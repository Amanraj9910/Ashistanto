# Build stage
FROM node:18-bullseye AS builder

WORKDIR /app

# Install ffmpeg and other required system dependencies
RUN apt-get update && apt-get install -y \
    ffmpeg \
    python3 \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy package files
COPY package*.json ./

# Install dependencies (production only)
RUN npm install --omit=dev

# Production stage
FROM node:18-bullseye

WORKDIR /app

# Install ffmpeg, curl (for health check), and other runtime dependencies
RUN apt-get update && apt-get install -y \
    ffmpeg \
    curl \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Copy node modules from builder
COPY --from=builder /app/node_modules ./node_modules

# Copy application files
COPY package*.json ./
COPY server.js .
COPY auth.js .
COPY tts-service.js .
COPY agent-tools.js .
COPY graph-tools.js .
COPY formatters.js .
COPY timezone-helper.js .
COPY action-preview.js .
COPY public/ ./public/

# Create working directory with proper ownership for node user
RUN chown -R node:node /app
USER node

# Expose port
EXPOSE 3000

# Build arguments for versioning
ARG BUILD_DATE
ARG VCS_REF
ARG VERSION

# Labels for image metadata
LABEL org.opencontainers.image.created="${BUILD_DATE}" \
      org.opencontainers.image.revision="${VCS_REF}" \
      org.opencontainers.image.version="${VERSION}" \
      org.opencontainers.image.title="Azure Voice AI Agent" \
      org.opencontainers.image.description="Voice-based AI agent using Azure Speech Services and Azure OpenAI"

# Set environment variables
ENV NODE_ENV=production
ENV PORT=3000
ENV DOCKER_ENV=true

# Health check
HEALTHCHECK --interval=30s --timeout=10s --start-period=5s --retries=3 \
    CMD curl -f http://localhost:3000/api/config || exit 1

# Start application
CMD ["node", "server.js"]
