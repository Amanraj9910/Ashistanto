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
