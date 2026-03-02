# Expert Front-End (SP Sir UI)

React + Vite frontend for SP Sir.  
It is designed to work with backend APIs under `/avatar-service`.

## Prerequisites

- Node.js 20+
- npm
- Optional: Docker

## Local Development

1. Install dependencies:

```bash
npm ci
```

2. Create env file (optional):

```bash
cp .env.example .env
```

3. Run dev server:

```bash
npm run dev
```

Default dev URL: `http://127.0.0.1:5173`

## Backend API Configuration

Environment variable:

- `VITE_API_BASE_URL` (default: `/avatar-service`)

Examples:

- Same origin reverse proxy: `/avatar-service`
- Direct backend URL: `http://localhost:8000/avatar-service`

## Production Build

```bash
npm run build
```

Build output is generated in `dist/`.

## Docker

Build image:

```bash
docker build -t expert-front-end:latest .
```

Run container:

```bash
docker run --rm -p 8080:80 expert-front-end:latest
```

Open UI at:

`http://localhost:8080`

## Nginx Proxy

`nginx.conf` proxies `/avatar-service/*` to backend service `backend:8000`.

If you deploy frontend separately, update this proxy target accordingly.

## Useful Scripts

- `npm run dev` - start local dev server
- `npm run build` - production build
- `npm run preview` - preview build locally

