# zzzs-watch

Daily check of the official ZZZS spreadsheet for personal doctors accepting new
patients around Ljubljana. Alerts only on doctors that are **new since last run**
and not on your ignore list.

```bash
pip install openpyxl

python watch.py --list                    # what's open right now
python watch.py --file SA_ZO_07_09_2026.xlsx --list   # parse a file you downloaded
python watch.py                           # the real daily run
python watch.py --all                     # ignore the ignore list
```

First real run records a silent baseline so you don't get the whole backlog.
Delete `state.json` to reset.

## Ignore list

Two sources. If `IGNORE_URL` is set it wins; `ignore.txt` is the offline fallback.

**Google Sheet (recommended — editable from your phone):**
In the sheet: File → Share → Publish to web → pick the tab → CSV → Publish.
Then set the `/pub?output=csv` URL it gives you as `IGNORE_URL` in `run.sh`.
The script reads the column headed `Name`, or the first column if there's no
such header. Publishing makes that tab readable by anyone with the link.

If the URL returns a login page instead of CSV, the script says so and falls
back to `ignore.txt` rather than treating the list as empty and re-alerting you
about every doctor you've already called.

**Local file:** `ignore.txt`, one name per line. Matching is accent-insensitive and
order-insensitive, so all of these hit the same doctor:

```
ŽAGAR  DRAŽEN
Drazen Zagar
zagar drazen
```

## Alerts

Email and phone-push are independent — set either, or both. Every configured
channel fires, and one failing does not silence the other. With neither set,
alerts print to stdout.

**Email.** Set `EMAIL_TO`, `SMTP_HOST`, `SMTP_PORT`, `SMTP_USER`, `SMTP_PASS`.
For Gmail you need a **Google App Password** (myaccount.google.com → Security →
2-Step Verification → App passwords). Your normal password will be rejected.
Port 587 uses STARTTLS; port 465 switches to implicit SSL automatically.

**GitHub issue (zero config, CI only).** Set `ALERT_FILE=alert.txt` and the
workflow turns it into an issue, which GitHub emails you. This is the default
in `.github/workflows/watch.yml`.

**Phone push.** Set `NTFY_TOPIC` to an unguessable string, install the ntfy app,
subscribe to that topic.

Test both at once:

```bash
python watch.py --test-notify
```

## Running it with your computer off

Push this to a **private** GitHub repo. `.github/workflows/watch.yml` then runs
it on GitHub's machines at 22:00, 03:00 and 13:00 UTC (midnight, 05:00 and
15:00 Ljubljana in summer, 23:00/04:00/14:00 winter), with nothing of yours
switched on.

Add these under Settings → Secrets and variables → Actions:

**Email needs no setup.** When something new turns up, the workflow opens a
GitHub issue and GitHub emails it to your account address. No SMTP, no app
password, no secrets. Check Settings → Notifications if the mail doesn't arrive.

Optional secrets, all of which the job runs fine without:

| Secret | Value |
|---|---|
| `IGNORE_URL` | published-CSV link to your sheet |
| `NTFY_TOPIC` | phone-push topic |
| `EMAIL_TO` / `SMTP_USER` / `SMTP_PASS` / `SMTP_HOST` | only if you want direct SMTP as well |

Then Actions tab → zzzs-watch → **Run workflow** to test it immediately rather
than waiting for the schedule.

`state.json` is committed back to the repo after each run — that's how the job
remembers what it already told you about. It must stay out of `.gitignore`.

Private repos get 2,000 free Actions minutes a month; three runs a day uses
well under 100. GitHub's cron is best-effort and can lag 5-30 minutes, which is
irrelevant for a file that changes once a day. GitHub also disables schedules on
repos with no activity for 60 days — it emails you first, and any commit or a
manual run resets the clock.

## Schedule (only if running on your own machine)

ZZZS refreshes this file about once a day, overnight on business days only
(it's still Friday's file on Saturday evening), so once or twice daily is
plenty — three below is already generous. Times are local (Ljubljana):

```cron
0 0,5,15 * * * /path/to/zzzs-watch/run.sh
```

Put your `NTFY_TOPIC` in `run.sh` — cron inherits almost no environment, which
is the usual reason a job works by hand and silently stops alerting under cron.

A stale-feed guard warns if the published data date is more than 4 days old, so
a frozen upstream shows up instead of looking like "no new doctors".

## How the file is located

`stdatoteke=14314` is a TYPO3 file ID and `cHash` is a hash over the query
params — both can rotate when ZZZS republishes. So the script scrapes the
listing page and matches the link by its **label** ("Število opredeljenih pri
splošnih zdravnikih"), falling back to the hardcoded URL only if that fails.
Downloads are checked for the `PK` zip magic, so an HTML error page raises
instead of being parsed as an empty spreadsheet.

Verify the fetch on first run:

```bash
python watch.py --check-fetch
```

It prints the resolved URL, content-type, the filename ZZZS served, byte count,
and the parsed data date. If the label match fails it falls back to the
hardcoded URL and says so on stderr — that warning means the page changed and
the label needs updating.

The download path was tested against a local mock (gzip listing page, HTML
entities, a decoy link, and an error-page response), not against the live host —
the sandbox this was built in couldn't reach `partner.zzzs.si`. The parser is
tested against the real 07.09.2026 file.

## Sheet layout (as of 07.09.2026)

One sheet, `Splošna dejavnost`. Row 2 holds the data date, row 4 the headers,
data from row 5. 1442 rows.

| Col | Field | Notes |
|---|---|---|
| A | Območna enota | admin unit; can disagree with column E — see below |
| E | Kraj | **the city filter runs on this** |
| G | Priimek in ime | surname first, often double-spaced |
| I | Naziv dejavnosti | adult `splošna ambulanta` vs `otroški in šolski dispanzer` |
| J | Še sprejema | **`DA`/`NE` — the whole signal** |
| M | Št. opredeljenih / tim | patient count per full team |
| N | Glavarinski količniki / tim | above ~1895 a doctor may refuse |

Identity key is `šifra zdravnika + šifra izvajalca + šifra dejavnosti`, not the
name — names get respelled, and one doctor can appear at several practices.

## Tuning

Edit at the top of `watch.py`:

- `CITIES` (env var) — practice city must contain one of these. Defaults to
  `LJUBLJANA`, which covers 1000 LJUBLJANA plus LJUBLJANA - POLJE / - ČRNUČE.
  Widen with `CITIES="LJUBLJANA,DOMZALE,KAMNIK"`. There are no coordinates in
  this file, so filtering is by city text, not radius.
- `WANTED_ACTIVITY` — currently adult family medicine only. Add `"otroski"` to
  include paediatric/school clinics.

## Column A vs column E

They can disagree. Two doctors at Zdravje Fit (Triglavska 6, Ljubljana) are
filed under `IZPOSTAVA GROSUPLJE`, which suggests column E is the company's
registered address rather than the room you'd sit in. Column E is used because
it's the more specific field, but confirm the actual location by phone.

## Caveat

`DA` means the doctor is still *obliged* to accept under their quota, not that a
slot is free today. Phone first.
