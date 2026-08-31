# Build stage
FROM node:20-bookworm AS builder

WORKDIR /app

# Install ffmpeg and other required system dependencies
RUN apt-get update --fix-missing && apt-get install -y \
    ffmpeg \
    python3 \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Copy package files
COPY package*.json ./

# Install dependencies (production only)
RUN npm install --omit=dev

# Frontend build stage
# The UI is a Next.js static export served by Express from the same origin. NEXT_PUBLIC_* values
# are INLINED INTO THE BUNDLE AT BUILD TIME -- passing them via `docker run -e` has no effect.
FROM node:20-bookworm AS ui
WORKDIR /app/frontend
COPY frontend/package*.json ./
RUN npm ci
COPY frontend/ ./
# Must be "false": lib/api.ts treats anything else as mocks-on and would ship fake data.
ENV NEXT_PUBLIC_ENABLE_MOCKS=false
# Must be empty so every fetch is same-origin relative.
ENV NEXT_PUBLIC_API_URL=""
RUN npm run build

# Production stage
FROM node:20-bookworm

WORKDIR /app

# Install ffmpeg, curl (for health check), and other runtime dependencies
RUN apt-get update --fix-missing && apt-get install -y \
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
COPY storage.js .
COPY public/ ./public/
# Built UI (index.html, login/, chat/, auth/success/, _next/, img/, official/, 404.html)
COPY --from=ui /app/frontend/out ./frontend/out

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
    CMD curl -f http://localhost:3000/api/config && curl -f -o /dev/null http://localhost:3000/ || exit 1

# Start application
CMD ["node", "server.js"]
