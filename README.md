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

The app starts on `http://localhost:5000`.

## Deployment

- Runs as a Docker container on Ubuntu Server via `docker-compose.yml` (`restart: unless-stopped` — survives reboots and crashes)
- Publicly exposed via **Cloudflare Tunnel** at [tender.liamatux.com](https://tender.liamatux.com) — no port forwarding needed, tunnel connects outbound to Cloudflare and routes to `localhost:5000`
- Cloudflare Tunnel runs as its own systemd service (`cloudflared`)

## CI

GitHub Actions (`.github/workflows/ci.yml`) builds the Docker image on every push/PR to catch build errors before deployment.

## Project structure
- `tender_app.py` — main application
- `Dockerfile` / `docker-compose.yml` — containerization
- `requirements.txt` — Python dependencies
