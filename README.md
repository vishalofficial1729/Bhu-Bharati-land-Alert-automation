# Bhu Bharati Land Monitor

Checks your land records on [Bhu Bharati](https://bhubharati.telangana.gov.in/knowLandStatus)
(Land Details Search → Survey No.), compares them with your own Excel list, and tells you when
something changes or doesn't match: a new pending transaction, prohibition, missing digital
signature, NALA land, a different owner or a different extent.

It runs on your own computer. Your sheets, results and history stay in this folder.

---

## Set up on a new computer

You only do this once. It takes about 10 minutes, mostly downloading.

### Mac

1. **Install Python 3.12**
   - Go to https://www.python.org/downloads/macos/ and download the **Python 3.12** macOS installer (the latest 3.12.x).
   - Open it and click through the installer.
   - When it finishes, a Finder window opens. Double-click **Install Certificates.command** in it. (Or open Applications → Python 3.12 → Install Certificates.command.) Without this, the Google Sheet connection fails.

2. **Download this app**
   - Open https://github.com/vishalofficial1729/Bhu-Bharati-land-Alert-automation
   - Click the green **Code** button → **Download ZIP**.
   - Double-click the downloaded ZIP to unzip it, and move the folder somewhere easy, such as Documents.

3. **Run the setup**
   - Open **Terminal** (press Cmd+Space, type *Terminal*, press Enter).
   - Type `cd ` (with a space after it), drag the app folder from Finder into the Terminal window, and press Enter.
   - Paste this and press Enter:
     ```bash
     bash setup.command
     ```
   - Wait until it says **Setup finished**.

### Windows

1. **Install Python 3.12**
   - Go to https://www.python.org/downloads/windows/ and download the **Python 3.12** Windows installer (64-bit).
   - Open it. On the first screen, **tick "Add python.exe to PATH"**, then click **Install Now**.

2. **Download this app**
   - Open https://github.com/vishalofficial1729/Bhu-Bharati-land-Alert-automation
   - Click the green **Code** button → **Download ZIP**.
   - Right-click the ZIP → **Extract All**, and put the folder somewhere easy, such as Documents.

3. **Run the setup**
   - Open the app folder and double-click **setup.bat**.
   - If Windows shows "Windows protected your PC", click **More info** → **Run anyway**.
   - Wait until it says **Setup finished**, then press any key.

---

## Start the app

| | How |
|---|---|
| **Mac** | In Terminal, go to the app folder (`cd ` + drag the folder in) and run `bash start.command` |
| **Windows** | Double-click **start.bat** |

Your browser opens the app at **http://127.0.0.1:8765**. Keep the Terminal or black window open while you use it. To stop the app, press **Ctrl+C** in that window, or close it.

---

## First use

1. **Add your land list.** Go to **Land lists** and drop your Excel sheet onto the drop zone. A check on Bhu Bharati starts right away, and each parcel fills in on the **Parcels** page as it is fetched (about 2 seconds per parcel).
   - Use **Blank template** on that page to get the right columns.
   - To try the app first, drop `sample_data/bhubharati_sample.xlsx`. It has 5 real survey numbers from Moinabad mandal.
2. **Look at the results.**
   - **Overview:** the dashboard.
   - **Parcels:** one row per survey number. Click a row to see the full Bhu Bharati record, how it compares with your sheet, and its history.
3. **Set up email alerts** in **Settings → Email alerts** (see below), then click **Send test email**.
4. **Optional:** turn on **Settings → Automatic checks** (for example 09:00, 14:00, 19:00). They run while the app is open.
5. **Optional:** connect a **Google Sheet** (see below).

The first check saves a starting point and reports any mismatches or warnings. After that you are only told about **changes**.

---

## Your Excel sheet

One row per survey number. Only the first sheet in the file is read.

| Column | Required | What to enter |
|---|---|---|
| District, Mandal, Village | yes | English name exactly as in the Bhu Bharati dropdowns |
| Survey No | yes | Exactly as on the Pahani/ROR, including Telugu letters, e.g. `22/ఇ`, `4/15అ/2/1` |
| Expected Owner Name | no | From your documents. English or Telugu |
| Expected Extent (Acres), Expected Extent (Guntas) | no | 1 acre = 40 guntas |
| Alert On | no | Fields to watch for this parcel, e.g. `Transaction Status, Land Status`. Blank means any change. Warnings and owner/extent mismatches always alert |

The khata number is not needed; the app picks it. If a name is misspelled, the error suggests the closest match (e.g. *"did you mean: aziznagar?"*). Open a sheet on the **Land lists** page to see problems in it: duplicates, blank villages, guntas over 39.

---

## What the app shows

| Page | What's there |
|---|---|
| **Overview** | Totals, parcel health, warnings found, alerts over the last 30 days, recent checks, parcels that need a look, summary by village |
| **Parcels** | One row per survey number, grouped by village or mandal, with search and status filters. Click a row for details or **Check this parcel**. **Re-check failed** retries only failed ones |
| **Alerts** | Everything noticed, newest first, with filters. Mark alerts as seen to clear the sidebar count |
| **Land lists** | Every sheet you have added, grouped by file name as versions. Open one to see its rows, problems in it, and how it differs from the list in use |
| **Settings** | Email, automatic checks, speed, Google Sheet, and where files are stored |

At the top right: **Check now** checks everything (it turns into **Stop** while running) and **Report** downloads an Excel report.

### What gets checked and flagged

- **Pending transaction:** the Transaction Status field is not empty
- **Prohibited:** any field mentions prohibition
- **Not digitally signed:** Land Status is not "Digital Signed"
- **NALA / non-agricultural:** any field mentions NALA or non-agricultural
- **Owner differs:** compared with Expected Owner Name (approximate match, since the site shows Telugu)
- **Extent differs:** compared with Expected Extent
- **Any other change** to the record since the last check

---

## Email alerts

You get one email after a check, only if something new was found. A survey number that fails 3 checks in a row also sends an email.

**Gmail:**
1. Turn on 2-Step Verification for the Google account.
2. Go to https://myaccount.google.com/apppasswords and create an app password named "Bhu Bharati". Copy the 16 letters.
3. In **Settings → Email alerts**, click **Gmail**. Enter the Gmail address, paste the app password, add who should receive alerts, and click **Save email settings**.
4. Click **Send test email**.

**Outlook / Office 365:** click **Outlook** in the same place. **Zoho:** click **Zoho**.

---

## Google Sheet (optional)

Copies the results into a Google Sheet you own after every check. It has a **Current** tab (one row per parcel with Pending Transaction / Prohibited / Digitally Signed / NALA columns) and a **History** tab. No Google Cloud project is needed.

1. Create a new Google Sheet.
2. Go to **Extensions → Apps Script**. Delete the sample code and paste in everything from `google_sheet_receiver.gs` (in this folder).
3. Change `const TOKEN = 'change-me'` to a secret phrase of your own, then click **Save**.
4. Click **Deploy → New deployment**. Click the gear icon, choose **Web app**, and set:
   - *Execute as:* **Me**
   - *Who has access:* **Anyone** (the link is unguessable, and the secret phrase blocks anyone else)
5. Click **Deploy** and allow the permissions Google asks for. Copy the **Web app URL**.
6. In the app, go to **Settings → Google Sheet**, paste the URL and your secret phrase, and click **Save**.
7. Click **Check now**. The Current and History tabs appear in your Sheet.

---

## Where your files are kept

Everything stays inside the app folder:

| What | Where |
|---|---|
| Every sheet you upload (each one kept, with date and time) | `data/uploads/` |
| Check history (records, alerts, past checks) | `data/monitor.db` |
| Excel report | `data/reports/land_status_report.xlsx` |
| Settings (including your email app password) | `config.json` |

**Settings → Files & storage** shows these with sizes and an **Open folder** button. To move the app to another computer with its history, copy the whole folder, including `data/` and `config.json`.

`data/` and `config.json` are never uploaded to GitHub.

---

## Speed and settings

**Settings → Speed:**
- **Browsers at once** (default 3). Measured on the live site: 30 survey numbers in 53 seconds, about 30 minutes per 1,000 parcels.
- **Pause after each parcel:** keeps the load on the government site low.
- **2Captcha key:** optional paid fallback if the built-in captcha reader fails 5 times.
- **Show the browsers while checking:** opens the browser windows so you can watch the app use the real Bhu Bharati site.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `python3: command not found` / `'python' is not recognized` | Install Python 3.12 (step 1). On Windows, re-run the installer and tick **Add python.exe to PATH** |
| Google Sheet: `CERTIFICATE_VERIFY_FAILED` (Mac) | Run **Install Certificates.command** from Applications → Python 3.12 |
| `Address already in use` | The app is already running in another window. Use that one, or close it first |
| Page looks broken after updating the app | Stop the app (Ctrl+C) and start it again |
| A survey number keeps failing | Open it on Parcels to see the reason. Usually the village or survey number in the sheet doesn't match the site exactly |

---

## Files in this folder

| File | Purpose |
|---|---|
| `setup.command` / `setup.bat` | One-time setup (Mac / Windows) |
| `start.command` / `start.bat` | Start the app (Mac / Windows) |
| `app.py` | Local web server, settings, uploads, scheduler |
| `ui.html` | The app's screens |
| `monitor.py` | Compares records, stores history, sends email, writes the report. `python monitor.py` runs a self-check |
| `scraper.py` | Drives the Bhu Bharati website. `python scraper.py Rangareddy Moinabad Chilkoor 122 --show` checks one survey number |
| `google_sheet_receiver.gs` | Apps Script for the optional Google Sheet |
| `templates/` | Blank Excel template |
| `sample_data/` | Sample sheet with 5 real survey numbers |

## Known limits

- If a survey number has more than one khata, it is reported as an error instead of guessing which khata to use.
- Owner-name matching is approximate. If a correct name keeps showing as different, enter it exactly as on the Pahani (Telugu works).
- Automatic checks only run while the app is open.
- If Bhu Bharati changes its page layout, `scraper.py` needs an update.
