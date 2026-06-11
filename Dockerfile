# GGC Deal Engine — analysis engine container (deploy target: Cloud Run).
#
# The pipeline runs analyses on background threads inside ONE process, and job
# state lives in that process's memory. Two hard consequences:
#   1. gunicorn must run a SINGLE worker (threads are fine),
#   2. the Cloud Run service must be deployed with --max-instances 1 and
#      CPU always allocated (--no-cpu-throttling) so background analysis
#      keeps running between status polls.
# See DEPLOYMENT.md for the exact deploy command.
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    # Writable per-instance scratch space (Cloud Run's filesystem is an
    # in-memory overlay). Durable copies of finished models go to Firebase
    # Storage when FIREBASE_STORAGE_BUCKET is set.
    JOBS_DIR=/tmp/ggc/jobs \
    IMG_CACHE_DIR=/tmp/ggc/img_cache \
    EXTRACTION_CACHE_DIR=/tmp/ggc/extraction_cache

WORKDIR /app

COPY requirements.txt ./
RUN pip install -r requirements.txt

# The engine is one Python file + GGC's official 16-tab template. index.html
# keeps the legacy localhost UI reachable at / for debugging; the hosted UI is
# the Next.js app in web/ (deployed to Vercel, not part of this image).
COPY backend.py index.html GGC_Blank_Underwriting_Sizer_Extended.xlsx ./

# Cloud Run injects PORT (8080 by default).
#   --workers 1  REQUIRED — job state is in-process; a second worker would
#                see none of the first worker's jobs.
#   --threads 16 concurrent request handling + background analysis threads.
#   --timeout 0  analyses legitimately run for minutes; never kill the worker.
CMD exec gunicorn --bind :${PORT:-8080} --workers 1 --threads 16 --timeout 0 backend:app
