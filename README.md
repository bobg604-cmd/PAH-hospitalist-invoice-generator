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
- whether to add handover time around evening and night shifts, team handovers
  around a run of team days, and day call prep time
- whether to add 30 minutes of clinical admin after certain shift types

You can also add schedule aliases if your name appears in more than one way on the schedule.

The website also has a `Templates` tab where you can:

- see which invoice template is currently active
- download either bundled PAH template
- switch to another saved template
- upload a new template and make it the active one

Most people never need to open that tab. See `Which Invoice Template Is Used` below.

Example:

- `Smith`
- `Smith, J`
- `J Smith`

## Where the Finished File Goes

The app saves the finished invoice into the `outputs` folder inside this project.

That folder is ignored by Git, so generated invoices are not pushed to GitHub in normal use.

If you host the app online, generated files are still created on the server, but they are meant to be downloaded right away rather than stored permanently.

## Older Months And Archive Tabs

The master Google Sheet keeps the current schedule on one tab and older months on
archive tabs such as `2026 Archive`.

You do not need to pick the right tab yourself. If the month you asked for is not
on the tab the master schedule URL points at, the app searches the other tabs in
the same workbook, archives included, and tells you which tab it used.

Tabs that are not schedules, such as contact lists and surge lines, are ignored.

## Which Invoice Template Is Used

The department reissued the invoice workbook, with higher after-hours rates, for
reporting periods from October 2026 onward.

The app keeps both workbooks and picks the right one from the month you are
invoicing:

- October 2026 and later: `FHA Invoice - Hospitalist V3.xlsx`
- September 2026 and earlier: `FHA Invoice 2025 Hospitalist All Hours Final - 02172026.xlsx`

You do not need to choose. Billing an earlier month still uses the rates that
applied then, which is what you want when a late or corrected invoice goes in.

The results page tells you which workbook the invoice was built from.

If you do pick one workbook by hand in the `Templates` tab, or upload your own,
that choice applies to every reporting period until you set the tab back to
`Match the reporting period`.

## Special Shift Rules

The app uses these rules:

- regular day/team shifts are `07:00-17:00`
- `ADMIT` shifts are billed as `11:00-17:00`
- if a day-shift entry has `*C` after the name, such as `Grewal*C`, that day is billed as `06:30-17:30`
- if you also worked the overnight into that same morning, the `*C` day shift starts when the overnight ends, so it is billed as `07:00-17:30` instead, because the same half hour cannot be billed twice
- `VIRTUAL` shifts are billed as `17:00-21:00` and are `Off-Site`
- clinical admin is `Off-Site` only when it follows a `VIRTUAL` shift
- clinical admin is `On-Site` for the other supported shift types

## Optional Clinical Admin

The Clinical Admin section of the form has some optional extras. Each one is
billed as `Non-Scheduled` `Clinical Admin`, and never overlaps time you are
already billing.

- **Evening shift handover** (ticked to start with): 30 minutes just before
  the shift, `16:30-17:00`
- **Night shift handover** (ticked to start with): 30 minutes just after the
  shift, `07:00-07:30`
- **Team day stretch handover**: one hour, `18:00-19:00`, on the day before you
  start a run of team day shifts, and again on the last day of that run. If you
  are on an evening shift at that time it moves to `15:30-16:30`, leaving time
  for the evening handover; if you are on a `VIRTUAL` shift it moves to
  `22:00-23:00`. If the new time is taken too, it is left off
- **Day call prep**: `06:00-06:30` before each `*C` day shift, noted as
  `printing day lists and billing sheets`

A handover is left off when you are already on shift at that time. For
example, an evening shift straight after your own day or ADMIT shift gets no
handover before it, and a night shift followed by your own day shift gets no
handover at 07:00. The results page lists every handover that was left off
and why.

After a night shift handover, the clinical admin follows the handover and is
15 minutes of `filling invoices`, such as `07:30-07:45`, because the handover
already has its own row. If you typed your own clinical admin note, that note
is kept. With no night handover, the admin is the usual 30 minutes straight
after the shift.

On the last day of a run of team days, the clinical admin stays right after the
shift, `17:00-17:30`, and the team handover follows at `18:00-19:00`.

The app also reads the day either side of the half-month you are invoicing. That
is how it knows whether a run of team days carries on past the 15th or the end
of the month, so each team handover is billed once, on the invoice for the date
it happens. If the next or previous month is not on the same sheet tab, check
handovers that fall on the first or last day of the period.

Every clinical admin row is dated by the day it starts, so admin after an
evening shift is dated the next morning.

## Split Shifts

Some shifts are shared between two people. The schedule writes them in one cell,
either with a slash or with one name per line, such as `Beegan/Gu`.

A shared shift is not a full shift, so the app does not bill it as one:

- if a time is written beside your name, such as `Gu 1200-1700` or
  `Chang (0700-1200)`, that time is used
- if no time is written, the row is still created, but its start and end are left
  blank and the row is highlighted in the workbook for you to fill in

Times can be written as `0700-1200`, `07:00-12:00`, `7am-12pm` or `7-12pm`.

When a row is left blank, the results page lists the date and shift so you know
exactly what to look up, and any clinical admin that follows that shift is left
blank and highlighted too.

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
  --clinical-admin-type virtual `
  --evening-handover 30 `
  --night-handover 30 `
  --team-stretch-handover `
  --day-call-prep
```

## Hosting On Railway

This project can be hosted on Railway so users can open a website instead of installing anything locally.

### What Changes For Website Users

On Railway:

- users open a website instead of double-clicking the `.bat` file
- the bundled invoice template for that reporting period is used automatically
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
7. Test the site by generating an invoice.

If you ever need to change the workbook used by the site, open the `Templates` tab in the app.

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
- The app includes both bundled templates, so past and current months both work after a fresh deploy.
- The app can track anonymous aggregate counts such as visits and invoice generations without storing physician names, MSP numbers, or schedule details.
- If you want login protection later, that can be added as a separate improvement.

## Files In This Project

- `hospitalist_invoice_generator.py`: the main app
- `Start Hospitalist Invoice App.bat`: the easiest way to launch it
- `templates/`: bundled invoice template files, one per rate period
- `requirements.txt`: Python dependencies for deployment
- `Procfile`: tells Railway how to start the web app
- `outputs`: where finished invoice files are saved
