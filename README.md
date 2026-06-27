# FitTrack — WhatsApp AI Fitness Platform

A production-ready WhatsApp bot that automates fitness client 
management for personal trainers using Claude AI.

## What it does

- Trainer registers clients via natural language WhatsApp message
- Claude AI extracts all details automatically
- Unique client ID generated (RAH-001, PRI-001, RAH-002)
- Styled Excel file auto-created and sent to both trainer and client
- Both parties notified on every update
- OCR — send blood report or weight scale photo → auto-updates Excel
- Conversational query engine — ask anything about client data
- Field-level permissions — trainer vs client access control
- Payment tracking with automatic due date reminders
- Backdated registration support
- Fully dynamic Excel columns — Claude decides based on data

## Tech Stack

- Python 3.10
- Flask (webhook server)
- Claude AI Sonnet 4.6 (natural language + vision/OCR)
- UltraMsg (WhatsApp Business API)
- OpenPyXL (Excel file creation)
- ngrok (internet tunnel)

## Version History

| Version | What was added |
|---|---|
| fittrack_final.py | Core bot — registration, updates, progress, trainer mode |
| fittrack_v14.py | Query engine, OCR vision, field permissions, payment tracking, backdated registration, dynamic columns |

## Setup

### 1. Install packages
\`\`\`
pip install flask requests anthropic openpyxl ngrok python-dotenv
\`\`\`

### 2. Create .env file
\`\`\`
ULTRAMSG_TOKEN=your_token_here
ANTHROPIC_API_KEY=your_key_here
OWN_NUMBER=917022206001
\`\`\`

### 3. Configure in fittrack_v14.py
\`\`\`python
ULTRAMSG_INSTANCE = "your_instance"
DESKTOP_PATH      = r"C:\Your\Path\FitTrack_Clients"
TRAINER_NUMBERS   = ["91XXXXXXXXXX"]
\`\`\`

### 4. Run
\`\`\`
Terminal 1: python fittrack_v14.py
Terminal 2: python start_ngrok.py
\`\`\`

### 5. Set webhook URL in UltraMsg dashboard
\`\`\`
https://your-ngrok-url.ngrok-free.dev/webhook
\`\`\`

## Trainer Commands

| Command | What it does |
|---|---|
| trainer register [details], phone [number] | Register new client |
| trainer update [ID] [field] [value] | Update client field |
| trainer note [ID] [text] | Add trainer note |
| trainer list | Show all clients |
| trainer progress [ID] | Client progress report |
| trainer details [ID] | Get client Excel file |

## Client Commands

| Command | What it does |
|---|---|
| update weight 79kg | Update weight |
| update blood sugar 95 | Update blood sugar |
| I have paid my fees | Mark payment as paid |
| my details | Get your Excel file |
| my progress | See weight journey |
| [any question] | Ask Claude about your data |

## Image/OCR Support

Send any of these as WhatsApp images:
- Weight scale photo → auto-reads weight
- Blood report → extracts all test values
- Body measurement chart → extracts chest, waist, hips etc

## Field Permissions

| Field | Trainer | Client |
|---|---|---|
| Weight | ✅ | ✅ |
| Blood Sugar | ✅ | ✅ |
| Diet | ✅ | ❌ |
| Calories | ✅ | ❌ |
| Payment Amount | ✅ | ❌ |
| Payment Status | Confirmed | Paid |
| Payment Due Date | ✅ | ❌ |

## Built by

Chenna Kesava Reddy — Software Engineer at Maersk  
Built as part of 4-day MCP Learning Journey — June 2026  
LinkedIn: linkedin.com/in/mckreddy-720859159

## Important

API keys and tokens are NOT included.  
Configure your own in `.env` file before running.  
Never commit `.env`, `credentials.json` or `token.json` to GitHub.