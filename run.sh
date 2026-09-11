#!/usr/bin/env bash
# Cron wrapper — cron inherits almost no environment, so set everything here.
cd "$(dirname "$0")" || exit 1

# --- email alerts (Gmail needs an App Password, not your normal password) ---
export EMAIL_TO="timotej.volavsek@gmail.com"
export SMTP_HOST="smtp.gmail.com"
export SMTP_PORT="587"
export SMTP_USER="timotej.volavsek@gmail.com"
export SMTP_PASS="xxxx xxxx xxxx xxxx"      # <-- Google App Password

# --- optional phone push; delete if you only want email ---
# export NTFY_TOPIC="pick-something-unguessable"

# --- optional: ignore list from your published Google Sheet ---
# export IGNORE_URL="https://docs.google.com/spreadsheets/d/e/..../pub?output=csv"

exec /usr/bin/python3 watch.py >> watch.log 2>&1
