# Email Outreach Automation

A desktop app that uses any LLM to write and send personalized cold emails at scale — from your CV and a contacts CSV.

## Features

- **LLM-powered personalization** — reads your CV and generates a unique email for each contact
- **Smart CV extraction** — on first run the LLM extracts your name, phone, email, and LinkedIn from the CV automatically; re-runs on demand if you enter your API key after loading the CV
- **8 LLM providers** — OpenAI, Gemini, Groq, OpenRouter, Anthropic, Mistral, Together AI, Cohere
- **Email template** — optionally upload a `.txt`, `.pdf`, or `.docx` template; the LLM follows its structure and tone while still personalizing per recipient
- **Company research** — optionally scrapes each company's website (or DuckDuckGo) before writing
- **Draft review** — generate all emails first, edit subject and body inline in the review table, then send only what you approve
- **Persistent settings** — all inputs (files, credentials, provider, goal, checkboxes) are saved automatically and restored on next launch; you never have to re-enter them
- **Safe sending** — configurable delay between emails, pause/resume/stop at any time
- **CV attachment** — optionally attach your CV to every email
- **SMTP test** — verify your Gmail credentials before launching a campaign
- **Activity log** — color-coded log with entry count, saved to `outreach_log.json`

## Setup

```bash
# 1. Clone and enter the folder
git clone <repo-url>
cd email-outreach-app

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run
python email_outreach_app.py
```

> **Python 3.9+** required.

## Gmail App Password

Google requires an App Password (not your regular password) when using SMTP:

1. Enable 2-Step Verification on your Google account
2. Go to **Google Account → Security → App Passwords**
3. Create a password for "Mail" → copy the 16-character code
4. Paste it into the **App Password** field in the app

## Contacts CSV format

The CSV must have at least an `email` column. All other columns are optional but used for personalization:

| Column | Aliases recognized |
|--------|--------------------|
| `email` | `e-mail`, `email_address`, `mail` |
| `name` | `full_name`, `first_name`, `contact` |
| `company` | `organization`, `employer`, `firm` |
| `role` | `position`, `title`, `job_title` |
| `website` | `url`, `homepage`, `site` |

Any extra columns are included in the prompt as additional context.

**Example:**

```csv
name,email,company,role,website
Alice Smith,alice@example.com,Acme Corp,Engineering Lead,acme.com
Bob Jones,bob@techco.io,TechCo,CTO,
```

## Email Template (optional)

Upload a `.txt`, `.pdf`, or `.docx` file containing a sample email. The LLM will follow its structure, tone, and length while replacing all placeholder content with personalized details for each recipient. Leave blank to let the LLM write freely from the prompt rules.

## Workflow

### Option A — Generate Drafts (recommended)
1. Fill in all fields and click **Generate Drafts**
2. Review every email in the table — click any row to edit its subject or body directly
3. Uncheck rows you want to skip
4. Click **Send** — emails go out with the configured delay

### Option B — Start Campaign
Generates and sends each email immediately (one by one). Enable **Preview each email** in Settings to approve each one before it sends.

### Settings persistence
Every field is saved automatically to `outreach_settings.json` in the working directory. On the next launch the app restores all inputs including file paths — if the files still exist they are re-loaded silently.

## Supported LLM Providers

| Provider | Default model |
|----------|--------------|
| OpenAI | `openai/gpt-oss-120b` |
| Gemini | `gemini-2.0-flash` |
| Groq | `llama-3.3-70b-versatile` |
| OpenRouter | `google/gemini-2.0-flash-exp:free` |
| Anthropic | `claude-3-5-haiku-20241022` |
| Mistral | `mistral-small-latest` |
| Together AI | `meta-llama/Llama-3.3-70B-Instruct-Turbo` |
| Cohere | `command-r-plus-08-2024` |

You can override the model by typing any model ID into the **Model** field.

## Files created at runtime

| File | Purpose |
|------|---------|
| `outreach_settings.json` | Saved inputs restored on next launch |
| `outreach_log.json` | Full activity log of all campaigns |

## Optional dependencies

| Package | Purpose |
|---------|---------|
| `pandas` | Faster / more robust CSV parsing |
| `PyPDF2` or `pypdf` | Read PDF CVs and templates |
| `python-docx` | Read DOCX CVs and templates |

The app runs without these but will show a warning if the relevant file type is loaded.

## License

MIT
