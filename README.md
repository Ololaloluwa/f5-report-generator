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

- **Online:** the default is stored in Render as a Secret File (Part D
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
Takes about 20 minutes the first time. You need: a GitHub account, Git for
Windows (git-scm.com - install with the defaults), and this folder.

### Part A - Clean the folder
1. Make sure this folder only has what the site needs. It should contain:
   `app.py`, `generate_weekly_report.py`, `generate_weekly_excel.py`,
   `run_weekly_reports.py`, `validate_outputs.py`, the two templates
   (`MTN_Security_Metrics_Report_SOURCE.pptx`,
   `MTN_F5_WAF_Weekly_Report_Temp.xlsx`), `requirements.txt`,
   `render.yaml`, `.gitignore`, `README.md`, and the `templates` and
   `static` folders. Delete anything else (old helper files, `venv`,
   `__pycache__`).

### Part B - Put the code on GitHub (private)
2. Go to github.com, sign in, click the **+** (top right) -> **New
   repository**.
3. Repository name: `f5-report-generator`. Choose **Private**. Leave
   "Add a README", ".gitignore" and "license" all **unticked** (this
   folder already has them). Click **Create repository**. Keep that page
   open - it shows the repository's address, like
   `https://github.com/<your-username>/f5-report-generator.git`.
4. Open this folder in File Explorer, click the address bar, type `cmd`
   and press Enter - a Command Prompt opens in this folder.
5. Run these one at a time (put your own address in the `remote` line):
   ```
   git init
   git add .
   git status
   ```
   **Check the `git status` list before going on.** It must show only the
   files from step 1 - the only `.xlsx` / `.pptx` files allowed are the
   two templates. If you see `MTN_F5_ATTACKS_SOURCE.xlsx`,
   `Certificate.xlsx` or anything with MTN data, stop - don't commit.
   ```
   git commit -m "First version of the F5 report generator"
   git branch -M main
   git remote add origin https://github.com/<your-username>/f5-report-generator.git
   git push -u origin main
   ```
   The first push opens a browser window asking you to sign in to
   GitHub - approve it. (If git says "Please tell me who you are", run
   the two `git config --global user.name / user.email` commands it
   shows, then `git commit` again.)
6. Refresh the GitHub page - your files should be there, marked Private.

### Part C - Create the site on Render
7. Go to render.com and click **Get Started** -> sign up **with GitHub**
   (same account). This links the two.
8. In the Render dashboard click **New** (top right) -> **Blueprint**.
9. Render lists your GitHub repositories. If `f5-report-generator` isn't
   there, click **Configure account** / **Connect GitHub**, allow Render
   to see that repository, and come back. Click **Connect** next to it.
10. Give the Blueprint a name (e.g. `f5-report-generator`), leave the
    branch as `main` and the Blueprint path as `render.yaml`.
11. Render reads `render.yaml` and shows what it will create: one web
    service called `f5-report-generator`, on the **Free** plan. It asks
    for a value for **APP_PASSWORD** - type the team password here
    (this is the password everyone will log in with).
12. Click **Deploy Blueprint**.
13. Click into the `f5-report-generator` service and open **Logs**. The
    first build installs the packages and takes 3-5 minutes. It's done
    when the logs show gunicorn "Listening at ..." and the status at the
    top says **Live**.
14. The site's address is at the top of the service page, something like
    `https://f5-report-generator.onrender.com` (Render adds a few extra
    characters if that name is taken). Open it - you should get the "Hi
    Guys" screen, then the login.

### Part D - Add the saved certificate list
Render stores extra files as text, so the spreadsheet goes in as text:
15. Open the folder that has your `Certificate.xlsx` (the folder above
    this one). Click the address bar, type `powershell`, press Enter.
16. Paste this and press Enter - it copies the certificate list, as
    text, to your clipboard (nothing is shown; that's normal):
    ```
    [Convert]::ToBase64String([IO.File]::ReadAllBytes("$PWD\Certificate.xlsx")) | Set-Clipboard
    ```
17. In Render, open the service -> **Environment** (left menu) -> under
    **Secret Files** click **+ Add Secret File**.
18. Filename: `Certificate.xlsx.b64` (exactly). Contents: paste
    (Ctrl+V) - it's one very long line of letters, that's right.
19. Click **Save Changes**. Render redeploys by itself (about a minute).
20. Log in to the site: the upload page should now say **Saved default
    - checked <date> by <engineer>**. If it says "No saved default",
    check the filename in step 18.

### Part E - Share it
21. Send your colleagues the site address and the team password. That's
    all they need - no Python, no `.bat` file.

### Keeping it up to date
- **Changed the code?** In this folder: `git add .`, then `git status`
  (check for MTN data again), `git commit -m "what changed"`,
  `git push`. Render sees the push and redeploys by itself in a few
  minutes.
- **New certificate list?** Redo steps 15-16, then in Render ->
  Environment -> Secret Files, edit `Certificate.xlsx.b64`, replace the
  contents, **Save Changes**.
- **New password?** Render -> Environment -> edit `APP_PASSWORD` ->
  **Save and deploy**. Everyone will need the new one.

### Good to know (free plan)
- The site sleeps after 15 minutes with no visitors. The next visit shows
  Render's loading page for about a minute, then the "Hi Guys" screen
  (once per visit, not on reloads), then the login.
- The free plan gives 750 hours a month - more than enough for one site
  running all month.
- Nothing is stored between visits anyway, so sleeping loses nothing.
- If something breaks, the **Logs** tab of the service shows the error.

## Privacy
Uploaded files live in a temporary folder that's deleted when you click
"Done", or automatically 15 minutes after uploading. Nothing is stored.
