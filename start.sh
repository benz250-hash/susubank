#!/usr/bin/env bash
set -e

mkdir -p .streamlit

if [ -z "$STREAMLIT_SECRETS_TOML" ]; then
  echo "ERROR: STREAMLIT_SECRETS_TOML is not set in Railway Variables"
  exit 1
fi

printf "%s" "$STREAMLIT_SECRETS_TOML" > .streamlit/secrets.toml

streamlit run asu_money2.py --server.port "${PORT:-8080}" --server.address 0.0.0.0
