# MB Ireland Wages Automation

Runs every Monday morning, pulls the previous week's staff shift hours from Square,
fills them into the monthly Excel wages file, and emails it to you.

---

## How it works

1. **Monday 8am Ireland time** — scheduler fires automatically
2. Pulls all closed shifts for **Mon–Sat and Sunday** from Square for every Irish location
3. Opens the current month's wages file (stored on the Railway volume)
4. Finds the correct week columns by matching the Sunday date in the header row
5. Writes each staff member's Mon-Sat hours and Sunday hours into the right cells
6. Saves the file and emails it to you with a summary

Lital and Karmika's hours are tracked but their final payment uses the fixed salary
already in the sheet — no changes needed there.

---

## Setup

### 1. Get your Square Location IDs

Go to: **Square Dashboard → Account & Settings → Locations**

Or run this in a terminal to list them:

```bash
curl -H "Authorization: Bearer YOUR_SQUARE_TOKEN" \
     -H "Square-Version: 2024-10-17" \
     https://connect.squareup.com/v2/locations
```

Copy the `id` value for each Irish store.

### 2. Deploy to Railway

```bash
# Install Railway CLI
npm install -g @railway/cli

# Login and create project
railway login
railway new mb-ireland-wages

# Link this folder
cd mb-ireland-wages
railway link

# Deploy
railway up
```

### 3. Add a Volume (persistent file storage)

In Railway Dashboard:
1. Open your service → **Volumes** tab
2. Add volume, mount path: `/data`

### 4. Set Environment Variables

In Railway Dashboard → **Variables**, add all values from `.env.example`:

| Variable | Value |
|---|---|
| `SQUARE_ACCESS_TOKEN` | Your Square API token |
| `SQUARE_LOC_BLANCHARDSTOWN` | Square location ID |
| `SQUARE_LOC_CORK` | Square location ID |
| `SQUARE_LOC_LIFFEY_VALLEY` | Square location ID |
| `SQUARE_LOC_NUTGROVE` | Square location ID |
| `SQUARE_LOC_WHITEWATER` | Square location ID |
| `SENDGRID_API_KEY` | SendGrid API key |
| `EMAIL_TO` | yuvi@memoryblock.com.au |
| `EMAIL_FROM` | wages@memoryblock.com.au (verified in SendGrid) |
| `API_KEY` | A strong secret key you choose |
| `DATA_DIR` | `/data` |

### 5. Upload the Monthly Wages File

At the start of each month, upload the blank monthly template:

```bash
curl -X POST "https://your-app.railway.app/upload?key=YOUR_API_KEY" \
     -F "file=@MB_Ireland_Jul_2026.xlsx"
```

Or use a tool like **Postman** / **Insomnia** if you prefer a UI.

---

## Endpoints

| Endpoint | Method | What it does |
|---|---|---|
| `/health` | GET | Check the app is running and what file is loaded |
| `/trigger?key=YOUR_KEY` | GET/POST | Manually run wages right now |
| `/upload?key=YOUR_KEY` | POST | Upload a new monthly wages file |
| `/download?key=YOUR_KEY` | GET | Download the current wages file |
| `/list-files?key=YOUR_KEY` | GET | List all uploaded files |

### Manual trigger (useful for testing)

```bash
curl -X POST "https://your-app.railway.app/trigger?key=YOUR_API_KEY"
```

---

## Monthly workflow

| When | What to do |
|---|---|
| Start of month | Upload the new blank monthly Excel template via `/upload` |
| Every Monday (auto) | Script runs at 8am, fills in last week, emails you the file |
| End of month | Download the completed file via `/download`, archive it |

---

## Staff name matching

Square employee names (e.g. "Anron Smith") are matched to Excel sheet names (e.g. "Anron")
using first-name matching. If a name doesn't match, it will appear in the email summary
as **unmatched** — you can update it manually in the sheet.

If Square has a staff member with an unusual name format, just make sure the first name
in Square matches what's in the Excel sheet.

---

## Troubleshooting

**"No column found for Sunday X"**
The monthly file doesn't have a column for that Sunday date. This happens if the file
for the wrong month is uploaded, or at the very start/end of a month (partial weeks).
Upload the correct monthly file.

**Staff showing as unmatched**
The name in Square doesn't match the name in the Excel sheet. Check both places and
align the first names.

**Email not arriving**
Check the `EMAIL_FROM` address is verified as a sender in your SendGrid account.

---

## Notes

- The scheduler runs inside the app (single-worker gunicorn). If the dyno restarts on Monday
  before 8am, the job will still fire at 8am once the app is back up.
- For extra reliability, you can also set up a Railway Cron service to hit `/trigger?key=YOUR_KEY`
  every Monday at `0 8 * * 1` (UTC+1 in summer = `0 7 * * 1` UTC).
- The wages file is named `MB_Ireland_*.xlsx` — the script always uses the most recently
  uploaded file (alphabetically last).
