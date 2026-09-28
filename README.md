# F5 Report Generator - web version

The weekly MTN F5 report from a website: log in, upload the source
spreadsheet, pick the week from a list, download the PowerPoint and Excel.

**This folder is self-contained - it is the git repository.** The report
code, both blank templates, `requirements.txt` and `render.yaml` are all
in here. The `.bat` version in the folder above runs this same code
(`f5-report-generator\run_weekly_reports.py`), so there's only ever one copy.

## Certificate list
The site uses a saved default `Certificate.xlsx`, so nobody has to upload
it each week. Anyone can still upload a different one for a single report
("Use a different certificate file") - the saved default doesn't change.
The upload page shows when the saved list was last checked.

- **Online:** the default is stored in Render as a Secret File (step 4
  below), so the MTN certificate list never goes into GitHub.
- **On your PC:** it uses the `Certificate.xlsx` in the folder above.

## Try it on your PC
In Command Prompt, from this `f5-report-generator` folder (first time only: the
`venv` and `pip` lines):
```
python -m venv venv
venv\Scripts\pip install -r requirements.txt
set APP_PASSWORD=test
venv\Scripts\python app.py
```
Then open http://127.0.0.1:5000 and log in with `test`. Ctrl+C stops it.

## Put it online (Render, free plan)
1. Create a **private** GitHub repository from this `f5-report-generator` folder.
   `.gitignore` keeps every spreadsheet and deck out except the two
   blank templates - never commit MTN data.
2. On render.com: **New -> Blueprint**, pick the repository. It reads
   `render.yaml` by itself.
3. When asked for `APP_PASSWORD`, enter the team password.
4. Add the default certificate list. Render's secret files are text, so
   turn the spreadsheet into text first - in PowerShell, from the folder
   that has `Certificate.xlsx`:
   ```
   [Convert]::ToBase64String([IO.File]::ReadAllBytes("$PWD\Certificate.xlsx")) | Set-Clipboard
   ```
   That copies it. In Render, open the service -> **Environment ->
   Secret Files -> Add Secret File**, name it `Certificate.xlsx.b64`,
   paste, and save (Render redeploys). Do the same whenever the
   certificate list is updated.
5. Share the web address Render gives you, plus the password.

On the free plan the site sleeps after 15 minutes with no visitors, so
the first visit after a break takes about a minute (Render's own loading
page). Then everyone gets the "Hi Guys" screen before the login - once
per visit, not on reloads.

## Privacy
Uploaded files live in a temporary folder that's deleted when you click
"Done", or automatically 15 minutes after uploading. Nothing is stored.
