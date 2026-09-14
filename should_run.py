"""Decide whether a scheduled run should actually do work.

GitHub's `schedule` trigger is best-effort and has been observed to silently
drop a target firing entirely (not just lag it). The workaround is to add
extra retry firings shortly after each real target time, but that alone would
email once per retry that fires. This script makes retries a no-op once one
attempt for the current slot has already succeeded, so you still get at most
one email per slot no matter how many retries fire.

A "slot" is one of the three daily target times (22:07, 03:07, 13:07 UTC).
The current slot is the most recent target time at or before now. If
state.json's checked_at is already at or after that slot's start, this slot
has already been handled and the run should be skipped.

workflow_dispatch runs are subject to the same slot dedup UNLESS the "force"
input is true — so an external cron hitting the dispatches API can poll
often (as a backup for dropped schedule firings) without causing duplicate
emails, while the manual "Run workflow" button in the UI defaults force=true
for quick testing.
"""

import datetime
import json
import os
import sys

TARGET_HOURS = (22, 3, 13)
TARGET_MINUTE = 7


def slot_starts_near(now):
    starts = []
    for day_offset in (-1, 0):
        day = (now + datetime.timedelta(days=day_offset)).date()
        for hour in TARGET_HOURS:
            starts.append(
                datetime.datetime(
                    day.year, day.month, day.day, hour, TARGET_MINUTE,
                    tzinfo=datetime.timezone.utc,
                )
            )
    return starts


def current_slot_start(now):
    return max(s for s in slot_starts_near(now) if s <= now)


def already_done(slot_start):
    try:
        with open("state.json") as f:
            checked_at = datetime.datetime.fromisoformat(json.load(f)["checked_at"])
    except (FileNotFoundError, KeyError, ValueError):
        return False
    return checked_at >= slot_start


def main():
    event = os.environ.get("GITHUB_EVENT_NAME", "schedule")
    force = os.environ.get("FORCE", "").lower() == "true"

    if event not in ("schedule", "workflow_dispatch"):
        print("run")
        return
    if event == "workflow_dispatch" and force:
        print("run")
        return

    now = datetime.datetime.now(datetime.timezone.utc)
    if already_done(current_slot_start(now)):
        print("skip")
    else:
        print("run")


if __name__ == "__main__":
    main()
