# Multi-stage build for slide-engine
FROM python:3.12-slim-bookworm AS builder

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PIP_DISABLE_PIP_VERSION_CHECK=1

WORKDIR /build

RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

COPY pyproject.toml /build/
COPY src/ /build/src/

RUN pip install --no-cache-dir hatchling \
    && pip install --no-cache-dir .

# Runtime stage
FROM python:3.12-slim-bookworm AS runtime

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1

RUN groupadd -g 1000 appuser && \
    useradd -u 1000 -g appuser -s /bin/bash -m appuser

WORKDIR /app

# Create assets and output directories with proper permissions
RUN mkdir -p /app/assets /app/output /app/examples && \
    chown -R appuser:appuser /app

COPY --from=builder /usr/local/lib/python3.12/site-packages /usr/local/lib/python3.12/site-packages
COPY --from=builder /usr/local/bin/slide-engine /usr/local/bin/slide-engine

COPY --chown=appuser:appuser assets/ /app/assets/
COPY --chown=appuser:appuser examples/ /app/examples/

USER appuser

VOLUME ["/app/assets", "/app/output", "/app/examples"]

ENTRYPOINT ["slide-engine"]
CMD ["--help"]