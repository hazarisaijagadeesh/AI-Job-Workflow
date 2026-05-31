# H1B Job Hunt Automation
### Sai Jagadeesh Hazari — Senior SDET

Runs every 12 hours via GitHub Actions (free).
Scrapes → H1B filters → ATS scores vs your resume → saves Excel → emails top matches.

---

## ⚡ Quick Setup (15 minutes)

### Step 1 — Create a private GitHub repo

1. Go to https://github.com/new
2. Name it `job-hunt-automation` (set to **Private**)
3. Click **Create repository**

### Step 2 — Upload these files

Upload the following to your repo root:
```
job_hunter.py
requirements.txt
.github/workflows/job_hunt.yml
```

To upload: go to your repo → **Add file** → **Upload files**

### Step 3 — Add your secrets

Go to: **Repo → Settings → Secrets and variables → Actions → New repository secret**

Add these 5 secrets:

| Secret name | Value |
|---|---|
| `YOUR_NAME` | `Sai Jagadeesh Hazari` |
| `YOUR_EMAIL` | `hazarisaijagadeesh@gmail.com` |
| `GMAIL_APP_PASS` | Your Gmail App Password (see below) |
| `ANTHROPIC_API_KEY` | Your Anthropic API key (see below) |
| `RESUME_PDF_B64` | Your resume as base64 (see below) |

---

## 🔑 Getting Each Credential

### Gmail App Password (not your real password)
1. Go to https://myaccount.google.com/security
2. Enable **2-Step Verification** if not already on
3. Go to https://myaccount.google.com/apppasswords
4. App: **Mail** | Device: **Other** → type "Job Hunt Bot"
5. Copy the 16-character password → paste as `GMAIL_APP_PASS`

### Anthropic API Key
1. Go to https://console.anthropic.com/
2. Click **API Keys** → **Create Key**
3. Copy it → paste as `ANTHROPIC_API_KEY`
4. Add ~$5 credit (enough for hundreds of runs)

### Resume as Base64
Run this command on your Mac/Linux terminal:
```bash
base64 -i Sai_Jagadeesh_Hazari_Resume.pdf | tr -d '\n'
```
Copy the output → paste as `RESUME_PDF_B64`

On Windows (PowerShell):
```powershell
[Convert]::ToBase64String([IO.File]::ReadAllBytes("Sai_Jagadeesh_Hazari_Resume.pdf"))
```

---

## ▶️ Running It

### Automatic
The workflow runs at **6 AM and 6 PM UTC** (= 2 AM / 2 PM ET) every day automatically.

### Manual trigger
1. Go to your repo → **Actions** tab
2. Click **H1B Job Hunt — Every 12 Hours**
3. Click **Run workflow** → **Run workflow**

---

## 📊 Viewing Results

### Excel tracker
After each run:
1. Go to **Actions** → click the latest run
2. Scroll to **Artifacts** at the bottom
3. Download `job-tracker-N.xlsx`

### Email digest
You'll receive two emails after each run:
- **Digest**: all shortlisted jobs with scores
- **Individual application emails**: sent for top 5 matches (score ≥ 70, H1B not blocked)

---

## ⚙️ Customization

Open `job_hunter.py` and edit the top section:

```python
MIN_ATS_SCORE  = 70   # raise to 80 for higher quality only
MAX_APPLY_PER_RUN = 5  # max emails sent per run

COMPANIES = [          # add/remove companies here
    ("Google", "https://careers.google.com/...", None),
    ...
]
```

### Adding more companies
```python
("Your Target Company", "https://company.com/careers?search=QA+automation", None),
```

Find the right URL by:
1. Going to the company's careers page
2. Searching "QA automation" or "SDET"
3. Copying the results URL

---

## 🛡️ H1B Filtering Logic

Jobs are filtered in two ways:

**Blocked (H1B = NO):** skipped automatically if description contains:
- "no sponsorship", "must be authorized", "no visa", "citizens only", etc.

**Confirmed (H1B = YES):** flagged as promising if description contains:
- "visa sponsorship", "will sponsor", "h1b", "sponsorship available", etc.

**Unknown:** included with a caution flag (many companies don't mention it in JD)

---

## 📋 Excel Columns Explained

| Column | What it means |
|---|---|
| Date Found | When the job was discovered |
| Company | Employer name |
| Job Title | Exact title from listing |
| ATS Score | 0–100 match vs your resume (≥80 = green, ≥65 = yellow, <65 = red) |
| H1B Sponsor | YES / NO / UNKNOWN based on JD text |
| Match Skills | Skills that matched your resume |
| Gaps | Skills mentioned in JD you don't have |
| AI Summary | 1-sentence Claude analysis |
| Apply URL | Direct link to job posting |
| Status | New / Applied / Interview / Rejected (update manually) |
| Notes | Your personal notes |

---

## 💰 Cost Estimate

| Service | Cost |
|---|---|
| GitHub Actions | **Free** (2,000 min/month for private repos) |
| Anthropic API | ~$0.01 per job scored · ~$0.50–1.00/month |
| Gmail | **Free** |
| **Total** | **< $1/month** |

---

## 🔧 Troubleshooting

**No jobs found**: Many career sites block scrapers. If a company returns 0 jobs, manually find their RSS feed or use their API. LinkedIn, Greenhouse, Lever, and Workday all have public APIs.

**Emails not sending**: Double-check your Gmail App Password. Make sure 2FA is enabled and you're using the App Password (not your login password).

**ATS scores all 70**: The Anthropic API key is missing or has no credit. Check your console.anthropic.com balance.

**GitHub Actions not running**: Check that the workflow file is at exactly `.github/workflows/job_hunt.yml` in your repo.
