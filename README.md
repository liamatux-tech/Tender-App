# Tender App

A small Flask web app that fetches tender information from a tender portal by event ID or link, and exports the details into a formatted Excel (.xlsx) report.

## Features
- Fetch tender data via `/api/tender`
- Export tender details to a styled Excel file via `/api/export`
- Simple single-page web UI

## Requirements
- Docker & Docker Compose

## Run locally

```bash
docker compose up -d --build
```

The app starts on `http://localhost:5000` (internal, exposed only inside the `edge` Docker network).

## Deployment

- Runs as an independent Docker Compose project on Hetzner, connected to a shared external network `edge`
- Reverse proxy handled by a separate `nginx-proxy` project (also on `edge`), terminating SSL with a Cloudflare Origin Certificate for the `liamatux.com` zone
- Publicly available at [tender.liamatux.com](https://tender.liamatux.com)
- **CI**: GitHub Actions (`.github/workflows/ci.yml`) builds the Docker image on every push/PR
- **CD**: GitHub Actions (`.github/workflows/deploy.yml`) auto-deploys to the server via SSH on every push to `main` — `git pull` + `docker compose up -d --build`, no manual steps required

## Architecture

GitHub (main) → GitHub Actions (CI + CD)
↓ SSH
Hetzner Server
│
┌──────────────┼──────────────┐
nginx-proxy portflix-app tender-app
(SSL, routing) (this project)
└───────── edge network ───────┘


## Project structure
- `tender_app.py` — main application
- `Dockerfile` / `docker-compose.yml` — containerization
- `requirements.txt` — Python dependencies
- `.github/workflows/` — CI and CD pipelines
