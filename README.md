# Injector Insights: sign-up site

Members sign up on the site. Their details are saved to **members.xlsx** and any certificate they upload is stored as a file.

## Run it on your computer
```
pip install -r requirements.txt
set ADMIN_PASSWORD=choose-a-strong-password      (Mac/Linux: export ADMIN_PASSWORD=...)
python app.py
```
Open http://localhost:8000

## Get your data (any time)
Sign in with any username and your ADMIN_PASSWORD:
- `/admin/members.xlsx` downloads the Excel file with every member.
- `/admin/certificates.zip` downloads all certificates. Each file is named with the member's ID, which is in column A of the sheet (column M shows the file name).

## Publish it
Use any host that runs Python and has a **persistent disk** (Render, Railway, Fly.io, a VPS). Without a persistent disk, the Excel file and certificates are lost at each redeploy.
1. Upload this folder (`app.py`, `requirements.txt`, `Procfile`, `static/`).
2. Start command: `gunicorn app:app --workers 1 --threads 4 --bind 0.0.0.0:$PORT` (already in `Procfile`). Keep it to **one worker** because the Excel file is written by one process.
3. Set environment variables: `ADMIN_PASSWORD` (required for admin downloads) and `DATA_DIR` (path to the persistent disk, for example `/data`).
4. Point your domain at the host and make sure HTTPS is on. The host usually does this for you.

## Before you go live
- Replace the bracketed placeholders (contact emails, governing jurisdiction) in the legal pop-ups in `static/index.html`, and have the documents reviewed by a lawyer.
- Certificates are sensitive personal data. Use disk encryption, back up the data folder, and delete certificates after verification, since the Privacy Policy says they are deleted once verified.
- The public form has no CAPTCHA or rate limit. Add one (for example Cloudflare Turnstile or your host's rate limiting) if you see spam.
