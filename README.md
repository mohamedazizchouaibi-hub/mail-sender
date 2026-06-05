# Email Outreach Automation

A desktop application that uses your favorite LLM to generate and send personalized cold emails at scale using your CV and a contacts CSV file.

## 🎥 Demo Video

Watch the complete usage walkthrough:

**Video Tutorial:**
https://drive.google.com/file/d/1f8DgNx43uJV5GagRr3rjvDFdJNQYpeq6/view?usp=drive_link

---

## ✨ Features

### 🤖 AI-Powered Personalization

* Generates unique cold emails for every contact using any supported LLM.
* Uses your CV and recipient information to create tailored outreach messages.
* Produces personalized subject lines and email content automatically.

### 📄 Automatic CV Extraction

The application can automatically extract:

* Name
* Email
* Phone number
* LinkedIn profile

from your CV using AI.

* Runs automatically on first load.
* Can be re-run whenever you change or add an API key.

### 🧠 Multiple LLM Providers

Supports 8 providers out of the box:

* OpenAI
* Gemini
* Groq
* OpenRouter
* Anthropic
* Mistral
* Together AI
* Cohere

### 📝 Email Templates

Optionally upload a template in:

* `.txt`
* `.pdf`
* `.docx`

The AI follows the structure, tone, and style of your template while personalizing the content for each recipient.

### 🔍 Company Research

* Scrapes company websites when available.
* Falls back to DuckDuckGo search when necessary.
* Uses gathered information to improve personalization quality.

### ✅ Draft Review & Editing

* Generate all emails before sending.
* Review every draft in a table.
* Edit subjects and bodies directly inside the application.
* Select exactly which emails should be sent.

### 💾 Persistent Settings

The application automatically saves and restores:

* CV file path
* Contacts CSV path
* Template file path
* SMTP credentials
* Selected provider
* Selected model
* Campaign goal
* Application settings

No need to reconfigure everything each time you launch the application.

### ⏱ Safe Sending Controls

* Configurable delay between emails
* Pause campaign
* Resume campaign
* Stop campaign at any time

### 📎 CV Attachment Support

Optionally attach your CV to every outgoing email.

### 📧 SMTP Verification

Verify Gmail SMTP credentials before launching a campaign.

### 📊 Activity Logging

* Color-coded logs
* Campaign statistics
* Automatic logging to `outreach_log.json`

---

## 🚀 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/mohamedazizchouaibi-hub/mail-sender.git
cd mail-sender
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Launch the Application

```bash
python email_outreach_app.py
```

> Python 3.9+ is required.

---

## 🔐 Gmail App Password Setup

Google SMTP requires an App Password instead of your normal account password.

### Step 1

Enable **2-Step Verification** on your Google account.

### Step 2

Go to:

```text
Google Account → Security → App Passwords
```

### Step 3

Create an App Password for:

```text
Mail
```

### Step 4

Copy the generated 16-character password and paste it into the application's **App Password** field.

---

## 📁 Contacts CSV Format

The CSV file must contain at least an `email` column.

### Supported Columns

| Standard Column | Recognized Aliases             |
| --------------- | ------------------------------ |
| email           | e-mail, email_address, mail    |
| name            | full_name, first_name, contact |
| company         | organization, employer, firm   |
| role            | position, title, job_title     |
| website         | url, homepage, site            |

Additional columns are automatically included in the AI prompt and can be used for personalization.

### Example CSV

```csv
name,email,company,role,website
Alice Smith,alice@example.com,Acme Corp,Engineering Lead,acme.com
Bob Jones,bob@techco.io,TechCo,CTO,
```

---

## 📄 Email Template (Optional)

Upload a sample email in one of the following formats:

* TXT
* PDF
* DOCX

The AI will follow the template's:

* Structure
* Tone
* Writing style
* Approximate length

while replacing the content with recipient-specific information.

If no template is provided, emails are generated entirely from the campaign goal.

---

## 🔄 Workflow

### Option A — Generate Drafts (Recommended)

1. Fill in all required fields.
2. Click **Generate Drafts**.
3. Review generated emails.
4. Edit any subject or email body.
5. Uncheck emails you do not want to send.
6. Click **Send**.

### Option B — Start Campaign

1. Configure your campaign.
2. Click **Start Campaign**.

Emails will be generated and sent one-by-one.

Enable **Preview Each Email** in Settings if you want manual approval before sending.

---

## 💾 Settings Persistence

All application settings are automatically stored in:

```text
outreach_settings.json
```

On startup, the application restores all saved values and automatically reloads referenced files when available.

---

## 🧠 Supported LLM Providers

| Provider    | Default Model                             |
| ----------- | ----------------------------------------- |
| OpenAI      | `openai/gpt-oss-120b`                     |
| Gemini      | `gemini-2.0-flash`                        |
| Groq        | `llama-3.3-70b-versatile`                 |
| OpenRouter  | `google/gemini-2.0-flash-exp:free`        |
| Anthropic   | `claude-3-5-haiku-20241022`               |
| Mistral     | `mistral-small-latest`                    |
| Together AI | `meta-llama/Llama-3.3-70B-Instruct-Turbo` |
| Cohere      | `command-r-plus-08-2024`                  |

You can enter any compatible model identifier manually in the **Model** field.

---

## 📂 Runtime Files

| File                     | Purpose                                     |
| ------------------------ | ------------------------------------------- |
| `outreach_settings.json` | Stores application settings and user inputs |
| `outreach_log.json`      | Stores campaign logs and activity history   |

---

## 📦 Optional Dependencies

| Package        | Purpose                            |
| -------------- | ---------------------------------- |
| pandas         | Faster and more robust CSV parsing |
| PyPDF2 / pypdf | Read PDF CVs and templates         |
| python-docx    | Read DOCX CVs and templates        |

The application can run without these packages but may show warnings when related file types are used.

---

## 🛠 Technology Stack

* Python
* Tkinter
* SMTP (Gmail)
* OpenAI API
* Gemini API
* Anthropic API
* Groq API
* OpenRouter API
* Mistral API
* Together AI API
* Cohere API

---

## 📜 License

MIT License

Feel free to use, modify, and distribute this project under the terms of the MIT License.
