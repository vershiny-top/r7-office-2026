# Debian slim + R7 Document Builder (documentconstructor) + Python (for codegen).
FROM debian:bookworm-slim

ENV DEBIAN_FRONTEND=noninteractive \
    PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

RUN apt-get update && apt-get install -y --no-install-recommends \
        ca-certificates wget gnupg python3 \
        fonts-dejavu fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# --- R7 apt repo ----------------------------------------------------------
# VERIFY: repo URL, key URL, and package name against the support page.
# Typical pattern (adjust to actual R7 docs):
RUN wget -qO- "https://download.r7-office.ru/repo/r7-office.key" \
        | gpg --dearmor -o /usr/share/keyrings/r7-office.gpg \
 && echo "deb [signed-by=/usr/share/keyrings/r7-office.gpg] https://download.r7-office.ru/repo/deb stable main" \
        > /etc/apt/sources.list.d/r7-office.list \
 && apt-get update \
 && apt-get install -y --no-install-recommends r7-office-documentconstructor \
 && rm -rf /var/lib/apt/lists/*
# --------------------------------------------------------------------------

# Fallback binary names to try at runtime: documentconstructor / docbuilder
RUN (command -v documentconstructor || command -v docbuilder) >/dev/null \
    || echo "WARN: neither documentconstructor nor docbuilder in PATH"

WORKDIR /work
CMD ["python3", "build_via_builder.py"]
