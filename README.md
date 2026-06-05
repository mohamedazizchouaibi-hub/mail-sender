# Email Outreach Automation

A desktop application that uses your favorite LLM to generate and send personalized cold emails at scale using your CV and a contacts CSV file.

## Demo Video

Watch the complete usage walkthrough:

**Video Tutorial:**  
https://drive.google.com/file/d/1f8DgNx43uJV5GagRr3rjvDFdJNQYpeq6/view?usp=drive_link

---

## Features

### AI-Powered Personalization

- Generates unique cold emails for every contact using any supported LLM.
- Uses your CV and recipient information to create tailored outreach messages.

### Automatic CV Extraction

- Extracts key information from your CV:
  - Name
  - Email
  - Phone number
  - LinkedIn profile
- Automatically runs on first load.
- Can be re-run at any time after adding or updating your API key.

### Multiple LLM Providers

Supports 8 providers:

- OpenAI
- Gemini
- Groq
- OpenRouter
- Anthropic
- Mistral
- Together AI
- Cohere

### Email Templates

Optionally upload a template in:

- `.txt`
- `.pdf`
- `.docx`

The AI follows the template's structure, tone, and style while personalizing each message.

### Company Research

- Scrapes company websites when available.
- Falls back to DuckDuckGo search if no website is provided.
- Uses gathered information to improve personalization.

### Draft Review & Editing

- Generate all emails before sending.
- Review subjects and email bodies in a dedicated table.
- Edit any draft directly inside the application.
- Send only the emails you approve.

### Persistent Settings

Automatically saves and restores:

- CV path
- Contacts CSV path
- Template path
- SMTP credentials
- Selected LLM provider
- Model name
- Campaign goal
- Application settings

No need to re-enter information every time.

### Safe Email Sending

- Configurable delay between emails.
- Pause campaigns.
- Resume campaigns.
- Stop campaigns at any time.

### CV Attachment Support

Optionally attach your CV to every outgoing email.

### SMTP Verification

Test your Gmail SMTP credentials before launching a campaign.

### Activity Logging

- Color-coded application log.
- Tracks campaign activity and status.
- Automatically saved to `outreach_log.json`.

---

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/mohamedazizchouaibi-hub/mail-sender.git
cd mail-sender
