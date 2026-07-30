# Ashistanto Frontend

Frontend layer for Ashistanto using Next.js 15, Tailwind CSS v4, Zustand, and TanStack Query.

## Run Locally

```bash
npm install
npm run dev
```

Open `http://localhost:3001`.

## Backend Integration

The frontend is wired to the existing backend routes used by `public/index.html`.
Run the existing Node/Express backend on `http://127.0.0.1:3000`, then run this frontend on `http://127.0.0.1:3001`.

Default integration:

```env
BACKEND_URL=http://127.0.0.1:3000
NEXT_PUBLIC_ENABLE_MOCKS=false
```

For frontend-only development without the backend:

```env
NEXT_PUBLIC_ENABLE_MOCKS=true
```

## Deploy to Vercel (frontend preview)

When deploying this folder as a standalone frontend, set the Vercel project
Root Directory to `frontend`, then configure this Environment Variable for
Preview and Production:

```env
NEXT_PUBLIC_ENABLE_MOCKS=true
```

This keeps the preview self-contained and uses the built-in mock responses.
When the backend is deployed, switch mocks off and add its public HTTPS URL:

```env
NEXT_PUBLIC_ENABLE_MOCKS=false
BACKEND_URL=https://your-backend.example.com
```

After changing environment variables, redeploy so Next.js can include them in
the client bundle. The Vercel build command is `npm run build`; no custom output
directory is required.