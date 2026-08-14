# Tender App

A small Flask web app that fetches tender information from a tender portal by event ID or link, and exports the details into a formatted Excel (.xlsx) report.

## Features
- Fetch tender data via `/api/tender`
- Export tender details to a styled Excel file via `/api/export`
- Simple single-page web UI

## Requirements
- Python 3.9+
- See `requirements.txt` for dependencies

## Setup

```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Run

```bash
python3 tender_app.py
```

The app starts on `http://0.0.0.0:5000`.

## Deployment
Running as a systemd service on Ubuntu Server (`tender-app.service`), auto-restarts on failure and starts on boot.
