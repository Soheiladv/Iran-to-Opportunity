<div dir="rtl">

# 🎯 MigrationHunter

### International Job Hunting Engine for Iranian Professionals

<div align="center">

![Version](https://img.shields.io/badge/version-3.0-blue)
![Python](https://img.shields.io/badge/Python-3.10+-green)
![License](https://img.shields.io/badge/License-MIT-yellow)
![Status](https://img.shields.io/badge/Status-Active-brightgreen)

</div>

---

## 🔗 Quick Links

| Link | Purpose |
|------|---------|
| **https://myaccount.google.com/apppasswords** | **Create Gmail App Password** |
| https://myaccount.google.com/security | Enable 2FA |
| https://mail.google.com/mail/u/0/#settings/fwdandpop | Enable IMAP in Gmail |

---

## 📖 About

**MigrationHunter** is an intelligent international job hunting system that:

- 🔍 Automatically searches global job boards
- 🤖 Uses AI API (OpenAI/Gemini) for intelligent analysis
- 🎯 Ranks opportunities based on applicant profiles
- 📊 Generates Excel dashboard with charts & conditional formatting
- 📧 Creates personalized cover letters & application emails
- 📤 Sends emails with user confirmation
- 🧠 Learns from previous searches & ranks successful sources
- 📝 Generates daily Persian reports

**End Goal:**

```
Real Job → Real Employer → Job Offer
→ Work Visa → Legal Relocation → Family Migration
```

---

## 🏗️ Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│  Collection     │ →  │  AI Analysis     │ →  │  Output         │
│  (Web Scraper)  │    │  (Job Analyzer)  │    │  (Excel + MD)   │
└─────────────────┘    └──────────────────┘    └─────────────────┘
        ↓                       ↓                       ↓
   ┌─────────┐           ┌──────────┐           ┌──────────┐
   │Requests │           │Scoring   │           │openpyxl  │
   │BeautifulSoup│      │Matching  │           │Markdown  │
   │HTML Parse│         │Learning  │           │Reports   │
   └─────────┘           └──────────┘           └──────────┘
```

### Three Main Layers

| Layer | Function | Tools |
|-------|----------|-------|
| **Collection** | Read sites & extract jobs | Python + requests |
| **Analysis** | Scoring, matching, learning | Python + AI API (OpenAI/Gemini) |
| **Output** | Excel + Persian reports + emails | openpyxl + Markdown + SMTP |

---

## 📁 Project Structure

```
MigrationHunter/
│
├── 📄 README.md                    ← This file
├── 📄 EXECUTION_GUIDE.md           ← Execution guide
├── 🐍 setup.py                     ← Dynamic setup (creates config.json + .env)
├── 🐍 run.py                       ← Main pipeline runner
├── 🐍 email_analyzer.py            ← Gmail scanner
├── 🐍 email_dashboard.py           ← Email analysis Excel
├── 🐍 job_crawler.py               ← Job board crawler
├── 🐍 build_dashboard.py           ← Master dashboard builder
├── 🐍 followup_reminder.py         ← Follow-up tracker
├── 🐍 config_loader.py             ← Dynamic config loader
├── 📄 .env.example                  ← Environment template
├── 📄 requirements.txt              ← Dependencies
├── 📄 .gitignore                    ← Protects personal data
│
├── 📂 config.json                   ← Main config (applicants, emails, linkedins)
├── 📂 .env                          ← Secrets (never commit!)
│
├── 📂 profiles/
│   ├── APPLICANT1_PROFILE.md       ← Profile template
│   └── APPLICANT2_PROFILE.md       ← Profile template
│
├── 📂 memory/
│   ├── SOURCE_BANK.md              ← Source bank (auto-learning)
│   ├── EMPLOYER_BANK.md            ← Employer bank
│   ├── JOB_BANK.md                 ← Job bank
│   ├── RECRUITER_BANK.md           ← Recruiter bank
│   ├── APPLICATION_BANK.md         ← Application bank
│   ├── VISA_BANK.md                ← Visa info bank
│   ├── REGISTRATION_BANK.md        ← Registration bank
│   └── SEARCH_HISTORY.md           ← Search history
│
├── 📂 input/
│   └── LATEST_SEARCH.md            ← Latest search input
│
├── 📂 output/
│   ├── TOP_JOBS.md                 ← Top opportunities
│   ├── EMPLOYERS_TO_CONTACT.md     ← Employers to contact
│   ├── RECRUITMENT_AGENCIES.md     ← Recruitment agencies
│   ├── GOVERNMENT_SOURCES.md       ← Government sources
│   ├── APPLICATIONS_TO_PREPARE.md  ← Applications to prepare
│   ├── LANGUAGE_REGISTRATION.md    ← Language & registration status
│   ├── SOURCE_BANK_UPDATE.md       ← Source bank updates
│   ├── EMAILS_TO_SEND.md           ← Emails ready to send
│   └── DAILY_ACTIONS.md            ← Top 5 daily actions
│
├── 📂 dashboard/
│   └── MigrationHunter_Dashboard_YYYYMMDD_HHMM.xlsx  ← Master dashboard
│
└── 📂 dashboard/archive/
    └── [Previous dashboard versions]
```

---

## 🚀 Installation & Usage

### Prerequisites

```bash
Python 3.10+
pip install -r requirements.txt
```

### Install

```bash
git clone https://github.com/YOUR_USERNAME/MigrationHunter.git
cd MigrationHunter
pip install -r requirements.txt
```

### Quick Start

```bash
# 1. Initial setup (interactive)
python setup.py

# 2. Test email connection
python email_analyzer.py --dry-run

# 3. Run full pipeline (5 steps)
python run.py
```

### What `run.py` Does

| Step | Script | Description | Output |
|------|--------|-------------|--------|
| 1 | `email_analyzer.py` | Scans last 30 days of Gmail | `memory/EMAIL_ANALYSIS.json` |
| 2 | `email_dashboard.py` | Creates email analysis Excel | `dashboard/Email_Analysis_*.xlsx` |
| 3 | `job_crawler.py` | Searches 9 job boards | `dashboard/Job_Crawler_*.xlsx` |
| 4 | `followup_reminder.py` | Generates follow-up report | `output/FOLLOWUP_REMINDER.md` |
| 5 | `build_dashboard.py` | Builds 13-sheet master dashboard | `dashboard/MigrationHunter_Dashboard_*.xlsx` |

### AI API Configuration (Optional)

```bash
# OpenAI
AI_PROVIDER=openai
AI_API_KEY=your_openai_api_key

# Or Gemini
AI_PROVIDER=gemini
AI_API_KEY=your_gemini_api_key
```

### Email Configuration (Optional)

```bash
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your_email@gmail.com
SMTP_PASSWORD=your_app_password
```

---

## 👥 Applicant Profiles (Dynamic)

Run `python setup.py` to create profiles. Each applicant gets:

- **Name** (English + Persian)
- **Profession** & **Keywords** (auto-generated)
- **Gender** (for emoji/color)
- **Language Levels** (English, German)
- **Multiple Emails** (each with App Password)
- **Multiple LinkedIn Profiles**

Example structure in `config.json`:

```json
{
  "applicants": [
    {
      "id": "applicant1",
      "name": "First Last",
      "name_fa": "نام فارسی",
      "gender": "female",
      "emoji": "👩",
      "profession": "Software Engineer",
      "keywords": ["software", "engineer", "python", "نام فارسی"],
      "emails": ["email1@gmail.com", "email2@gmail.com"],
      "linkedins": ["https://linkedin.com/in/profile1"],
      "english": "C1",
      "german": "A1"
    }
  ]
}
```

---

## 🌍 Target Countries

### Tier 1 — Primary Focus

| Country | Key Pathways |
|---------|--------------|
| 🇳🇿 **New Zealand** | Green List, AEWV, Accredited Employer |
| 🇩🇪 **Germany** | EU Blue Card, Skilled Worker Visa (§18a) |
| 🇦🇺 **Australia** | TSS 482 Visa, Employer Sponsorship |
| 🇨🇦 **Canada** | Express Entry, Atlantic Immigration, LMIA |

### Tier 2 — European Options

| Country | Key Pathways |
|---------|--------------|
| 🇦🇹 Austria | Red-White-Red Card |
| 🇮🇪 Ireland | Critical Skills Visa |
| 🇳🇱 Netherlands | Highly Skilled Migrant |
| 🇸🇪 Sweden | Work Permit (Healthcare/IT) |
| 🇳🇴 Norway | Skilled Worker |
| 🇩🇰 Denmark | Positive List / Pay Limit |
| 🇫🇮 Finland | Residence Permit (Specialist) |
| 🇬🇧 UK | Health & Care Visa / Skilled Worker |
| 🇨🇭 Switzerland | Work Permit |
| 🇧🇪 Belgium | EU Blue Card |
| 🇵🇹 Portugal | D1 Work Visa |
| 🇪🇸 Spain | EU Blue Card |
| 🇮🇹 Italy | Decreto Flussi |
| 🇫🇷 France | Talent Passport |
| 🇵🇱 Poland | Work Permit (IT/Healthcare) |
| 🇨🇿 Czechia | EU Blue Card |
| 🇭🇺 Hungary | White Card |
| 🇷🇴 Romania | EU Blue Card |
| 🇭🇷 Croatia | EU Blue Card |

---

## 📊 Scoring System

### Path Fit Score (0-100)

| Criterion | Weight | Description |
|-----------|--------|-------------|
| **Professional Fit** | 20% | Job matches experience & skills |
| **Migration Fit** | 20% | Visa pathway & legal requirements |
| **Language Fit** | 15% | Employer/visa/registration language needs |
| **Sponsorship Fit** | 25% | Likelihood of employer sponsorship |
| **Family Fit** | 10% | Spouse/children relocation feasibility |
| **Speed** | 10% | Estimated processing time |

### Dashboard Color Coding

| Color | Meaning |
|-------|---------|
| 🟢 Green | Verified / Ready / High Score |
| 🟡 Yellow | Probable / Medium Score |
| 🟠 Orange | Possible / Low Score |
| 🔴 Red | Expired / Rejected |
| 🔵 Blue | New / Identified |
| ⚪ Gray | Unknown |

---

## 🧠 Learning System

System learns from previous searches:

### Source Score Increases

| Event | Score |
|-------|-------|
| Valid job found | +2 |
| Employer verified | +3 |
| International hiring confirmed | +4 |
| Sponsorship confirmed | +5 |
| Application sent | +2 |
| Employer response | +5 |
| Interview | +8 |
| Job offer | +15 |

### Source Score Decreases

| Event | Score |
|-------|-------|
| Job expired | -1 |
| Duplicate | -1 |
| Invalid job | -3 |
| Fake sponsorship | -10 |
| Fake employer | -10 |
| Scam detected | -20 |
| Verified scam | -50 |

> ⚠️ **Note:** Historical score only affects search priority. Every new opportunity must be re-verified.

---

## 📧 Email & Cover Letters

### Core Rule

```
AI generates email
↓
User reviews content
↓
User confirms (y/N)
↓
Send / Save / Copy
```

**Never sends without explicit user confirmation.**

### Usage

```bash
# Generate emails with AI
python run_auto.py --ai --email

# Interactive send
python run_auto.py --interactive-email
```

### Send Options

| Option | Description |
|--------|-------------|
| 1 | Send via SMTP (requires config) |
| 2 | Save as text file |
| 3 | Copy to clipboard |
| 4 | Cancel |

---

## 📋 Core Rules

### Language Rule — Search Never Stops

```
❌ Old: "No IELTS 7 → Discard opportunity"
✅ New: "Track opportunity → Language = verification step"
```

**Language is not a pre-filter.** Only a final verification step.

### No Fabrication

- Never create fake jobs
- If unknown: mark **UNKNOWN**
- Priority to official sources & real employers

### No Memory Deletion

- No history deleted
- Old info marked: OUTDATED / EXPIRED

### Anti-Scam Protection

Flag immediately:
- Guaranteed job offers
- Guaranteed visas
- Payment for job offers
- WhatsApp/Telegram only communication
- Pressure for upfront payment

---

## 📊 Master Dashboard (13 Sheets)

| Sheet | Title | Content |
|-------|-------|---------|
| 01 | **Dashboard** | KPI cards, country/applicant breakdown, Top 5 |
| 02 | **Opportunities** | Full table with filters |
| 03+ | **Applicant — Profession** | Per-applicant opportunity sheets (dynamic) |
| 05 | **Employers** | Verified employer bank with emails |
| 06 | **Emails** | Ready-to-send email list |
| 07 | **Applications** | Application pipeline tracker |
| 08 | **Follow-up** | 7-day follow-up tracker |
| 09 | **Visa** | Visa requirements per country |
| 10 | **Registration** | Professional registration pathways |
| 11 | **Evidence** | Evidence scoring matrix |
| 12 | **History** | Search history log |
| 13 | **Email Analysis** | Email category statistics |

### Formatting

- **Fonts:** B Mitra (Persian) + Times New Roman (English)
- **Direction:** RTL (Right-to-Left)
- **Conditional formatting:** Green/Yellow/Red based on decision
- **Freeze panes:** Headers always visible
- **Auto-filter:** All tables filterable

---

## 🔧 Available Commands

```bash
# Full pipeline
python run.py

# Test email only
python email_analyzer.py --dry-run

# Email analysis only
python email_dashboard.py

# Job search only
python job_crawler.py

# Follow-up reminder only
python followup_reminder.py

# Dashboard only
python build_dashboard.py

# Re-run setup (add/modify applicants)
python setup.py
```

---

## 🔒 Privacy

- Personal data stored locally only
- No data sent to external servers (unless user requests AI analysis)
- Emails sent only with explicit user confirmation
- `.env` and `config.json` are gitignored

---

## 🤝 Contributing

Designed for personal use but contributions welcome:

1. Fork
2. Create branch
3. Submit Pull Request

---

## 📜 License

MIT License — Free use with attribution

---

## 📞 Contact

**GitHub:** [MigrationHunter](https://github.com/YOUR_USERNAME/MigrationHunter)
**Issues:** [GitHub Issues](https://github.com/YOUR_USERNAME/MigrationHunter/issues)

---

<div align="center">

### 🎯 Mission

```
Real Job → Real Employer → Job Offer
→ Work Visa → Legal Relocation → Family Migration
```

**Every job link is a step closer to family.**

</div>

</div>