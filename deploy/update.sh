#!/usr/bin/env bash
set -euo pipefail

cd -- "$(dirname -- "${BASH_SOURCE[0]}")/.."
if [[ -n "$(git status --porcelain)" ]]; then
    echo "Commit or stash local changes before updating EtoGradeMo." >&2
    exit 1
fi

docker compose config --quiet
git pull --ff-only
docker compose build --pull
docker compose up -d --wait --wait-timeout 60
echo "EtoGradeMo is running and responding."
