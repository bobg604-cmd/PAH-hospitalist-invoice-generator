# PAH Hospitalist Invoice Generator

This tool creates a hospitalist invoice workbook from the master schedule.

You do not need to know Python to use the normal version of this app.

## The Easy Way

For most people, using this app should look like this:

1. Open this folder on your computer.
2. Double-click `Start Hospitalist Invoice App.bat`.
3. A black window will open.
4. Open `http://127.0.0.1:8765` in your web browser.
5. Fill in the form.
6. Click the button to generate the invoice.
7. Download the finished Excel file from the page.

That is the main way this project is meant to be used.

## What You Need Before You Start

You need:

- this project folder
- your invoice template Excel file
- access to the master Google Sheet
- your physician billing details

## What the App Asks You For

The form will ask for:

- physician name
- first name
- last name
- MSP number
- site, such as `PAH`
- the name used on the master schedule
- the month and half of the month
- whether to add 15 minutes of clinical admin after certain shift types

You can also add schedule aliases if your name appears in more than one way on the schedule.

Example:

- `Smith`
- `Smith, J`
- `J Smith`

## Where the Finished File Goes

The app saves the finished invoice into the `outputs` folder inside this project.

That folder is ignored by Git, so generated invoices are not pushed to GitHub in normal use.

If you host the app online, generated files are still created on the server, but they are meant to be downloaded right away rather than stored permanently.

## Special Shift Rules

The app uses these rules:

- regular day/team shifts are `07:00-17:00`
- if a day-shift entry has `*C` after the name, such as `Grewal*C`, that day is billed as `06:30-17:30`
- `VIRTUAL` shifts are billed as `17:00-21:00` and are `Off-Site`
- clinical admin is `Off-Site` only when it follows a `VIRTUAL` shift
- clinical admin is `On-Site` for the other supported shift types

## If Double-Clicking Works

If the app starts and the browser page opens, you do not need to worry about Python or installing anything else.

## If Double-Clicking Does Not Work

If the `.bat` file does not start the app, there are usually 2 possible reasons:

1. The bundled Python runtime is missing on that computer.
2. A required Python package is missing.

If that happens, ask whoever is setting up the computer to help with the install, or use the steps below.

## Simple Repair Steps

If Python is already installed on the computer, open PowerShell in this folder and run:

```powershell
python -m pip install openpyxl
python .\hospitalist_invoice_generator.py --serve
```

Then open:

```text
http://127.0.0.1:8765
```

## If Python Is Not Installed

If the computer does not have Python at all, the easiest fix is:

1. Install Python from the official Python website.
2. During install, make sure Python is added to PATH if that option appears.
3. Open PowerShell in this folder.
4. Run:

```powershell
python -m pip install openpyxl
python .\hospitalist_invoice_generator.py --serve
```

If the person using this tool is not comfortable doing that, it is completely reasonable to have someone else do the one-time setup.

## Advanced Use

Most users can ignore this section.

You can also run the generator from the command line instead of the web form.

Example:

```powershell
python .\hospitalist_invoice_generator.py `
  --physician-name "Dr Jane Smith" `
  --first-name "Jane" `
  --last-name "Smith" `
  --msp "12345" `
  --site "PAH" `
  --schedule-name "Smith" `
  --year 2026 `
  --month 4 `
  --period first `
  --clinical-admin-type scheduled `
  --clinical-admin-type virtual
```

## Hosting On Railway

This project can be hosted on Railway so users can open a website instead of installing anything locally.

### What Changes For Website Users

On Railway:

- users open a website instead of double-clicking the `.bat` file
- the bundled invoice template is used automatically
- users then download the finished workbook from the results page

They do not need Python installed on their computer.

### Railway Setup

1. Push this project to GitHub.
2. Create a new Railway project.
3. Choose `Deploy from GitHub repo`.
4. Select this repository.
5. Railway should detect the Python app automatically because this repo includes:
   - `requirements.txt`
   - `Procfile`
6. After deploy, open the Railway-generated public URL.
7. Test the site by uploading a template workbook and generating an invoice.

### Custom Domain

After the Railway app is working, you can add your own domain inside Railway.

Typical setup:

1. Buy a domain from a registrar such as Cloudflare Registrar or Namecheap.
2. In Railway, open your service settings and add a custom domain.
3. Copy the DNS record Railway gives you.
4. Add that DNS record at your domain registrar.
5. Wait for the domain to verify.

### Notes For Hosting

- Railway storage is not permanent, so users should download files right away.
- The app includes a bundled default template for normal use.
- If you want login protection later, that can be added as a separate improvement.

## Files In This Project

- `hospitalist_invoice_generator.py`: the main app
- `Start Hospitalist Invoice App.bat`: the easiest way to launch it
- `templates/`: bundled invoice template files used by default
- `requirements.txt`: Python dependencies for deployment
- `Procfile`: tells Railway how to start the web app
- `outputs`: where finished invoice files are saved
