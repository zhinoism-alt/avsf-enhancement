# Daily AVSF digest

Emails a list of requests that have been open for N+ days (default 7), split into
"in our queue" and "waiting on others". Runs on a schedule via GitHub Actions
(`.github/workflows/daily-digest.yml`, weekdays 13:00 UTC). The workflow only runs
on the default branch.

## One-time setup (repo → Settings → Secrets and variables → Actions)

Secrets:

| Name | Value |
|---|---|
| `FIREBASE_SERVICE_ACCOUNT` | Full JSON of a Firebase service-account key (Firebase console → Project settings → Service accounts → Generate new private key). Needed because the Firestore rules block anonymous reads. |
| `SMTP_HOST` / `SMTP_PORT` | Your mail server, e.g. `smtp.office365.com` / `587` |
| `SMTP_USER` / `SMTP_PASS` | Mailbox login (use an app password where required) |
| `DIGEST_FROM` | Optional sender address (defaults to `SMTP_USER`) |
| `DIGEST_TO` | Recipients, comma-separated |

Variables (optional): `STUCK_DAYS` (default `7`), `DASHBOARD_URL` (link in the email).

Test it: Actions → "Daily AVSF digest" → Run workflow. The default is a dry run that
prints the digest in the job log. Untick "dry run" to send a real email.

## Local

    cd digest && npm install
    FIREBASE_SERVICE_ACCOUNT="$(cat key.json)" node send-digest.js --dry-run

If your company blocks SMTP basic auth, use a mailbox or relay that allows it, or
use the in-app "Draft digest" button (bell panel) to open the same list in Outlook.
