# ---- Stage 1: build dependencies ----
FROM python:3.11-slim AS builder

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir --user -r requirements.txt

# ---- Stage 2: final slim image ----
FROM python:3.11-slim

WORKDIR /app
COPY --from=builder /root/.local /root/.local
COPY tender_app.py .

ENV PATH=/root/.local/bin:$PATH

EXPOSE 5000
CMD ["python3", "tender_app.py"]
