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

The app starts on `http://localhost:5000` (internal only, exposed via the shared `edge` Docker network — not published to the host directly).

## Deployment

Runs as an independent Docker Compose project on a Hetzner VPS, alongside two other independent projects (`nginx-proxy` and `portflix-deploy`), connected only through a shared external Docker network called `edge`. Nothing is nested — each project lives in its own folder with its own docker-compose.yml.

- **nginx-proxy** — separate project, terminates SSL with a Cloudflare Origin Certificate (wildcard for `*.liamatux.com`), routes by domain to the right container by service name
- **Public URL**: [tender.liamatux.com](https://tender.liamatux.com)
- **CI**: `.github/workflows/ci.yml` — on every push/PR: Python syntax check (`py_compile`), then builds the Docker image
- **CD**: `.github/workflows/deploy.yml` — on every push to `main`: SSH into the server, `git pull`, `docker compose up -d --build`. Fully automatic, no manual deploy steps
- **Image**: multi-stage Dockerfile (build stage installs deps, final stage copies only the app + installed packages) — cut image size from 259MB to 203MB
- **Healthcheck**: Docker healthcheck pings the app every 30s; combined with `restart: unless-stopped`, the container self-heals on crash (tested live — verified it actually restarts, not just configured)

- **nginx-proxy** — separate project, terminates SSL with a Cloudflare Origin Certificate (wildcard for `*.liamatux.com`), routes by domain to the right container by service name
- **Public URL**: [tender.liamatux.com](https://tender.liamatux.com)
- **CI**: `.github/workflows/ci.yml` — on every push/PR: Python syntax check (`py_compile`), then builds the Docker image
- **CD**: `.github/workflows/deploy.yml` — on every push to `main`: SSH into the server, `git pull`, `docker compose up -d --build`. Fully automatic, no manual deploy steps
- **Image**: multi-stage Dockerfile (build stage installs deps, final stage copies only the app + installed packages) — cut image size from 259MB to 203MB
- **Healthcheck**: Docker healthcheck pings the app every 30s; combined with `restart: unless-stopped`, the container self-heals on crash (tested live — verified it actually restarts, not just configured)

## Monitoring

A separate `monitoring` project on the same server (also on the `edge` network) runs:
- **Prometheus** — scrapes metrics from itself, `node-exporter` (host CPU/RAM/disk), and `cadvisor` (per-container metrics)
- **Grafana** — [grafana.liamatux.com](https://grafana.liamatux.com), dashboards for host resources and per-container CPU/memory/network, including a custom panel built from scratch with PromQL

## Infrastructure as Code

- **Terraform**: current server described as code and imported (`terraform import`) — config not yet committed to a repo, since the server is planned to be replaced soon with a new VPS
- **Ansible**: playbook written to provision a fresh server from scratch (Docker install, firewall rules, `edge` network, project folders) — not yet run against a live target, pending the new VPS

## Project structure
- `tender_app.py` — main application
- `Dockerfile` / `docker-compose.yml` — containerization (multi-stage build)
- `requirements.txt` — Python dependencies
- `.github/workflows/` — CI and CD pipelines

## Still on the roadmap
- Kubernetes (k3s) — planned on a separate dedicated server, pending access
- `git rebase -i` practice on a batch of small commits
