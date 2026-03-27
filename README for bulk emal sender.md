# Bulk Email Sender 📧

Send personalized emails to hundreds of contacts from a CSV file — automatically. No mail merge, no expensive tools, just Python.

Built with Python and Tkinter. Works with Gmail, Outlook, or any SMTP provider.

---

## Features

- 📋 **Load contacts from CSV** — any CSV with an `email` column works
- ✍️ **Personalized templates** — use `{name}`, `{company}`, or any column as a placeholder
- 👁️ **Email preview** — see exactly what the first contact will receive before sending
- 📊 **Live progress bar** — tracks emails sent in real time
- 📝 **Activity log** — see success and failure for every contact
- 🔒 **App Password support** — works with Gmail's 2FA security
- 🧵 **Background threading** — UI stays responsive while sending
- 🌑 **Dark themed UI** built with Tkinter

---

## Screenshots

> *(Add screenshots here after running the app)*

---

## How It Works

1. Run the script
2. Enter your SMTP settings (host, port, email, app password)
3. Load your CSV contact list
4. Write your subject and body with `{placeholders}`
5. Preview the first email
6. Click **Send Emails**

---

## CSV Format

Your CSV must have an `email` column. Any other columns become available as placeholders:

```csv
name,email,company
John Smith,john@example.com,Acme Corp
Sarah Lee,sarah@example.com,TechStart
```

Then in your email body use:
```
Hi {name}, I'd love to connect with you at {company}...
```

---

## Gmail Setup

1. Enable 2-Step Verification on your Google account
2. Go to **myaccount.google.com → Security → App Passwords**
3. Create an app password for "Mail"
4. Use that password in the app (not your regular password)

SMTP settings for Gmail:
- Host: `smtp.gmail.com`
- Port: `587`

---

## Requirements

```
Python 3.8+
tkinter (built-in)
```

No extra packages needed — uses Python's built-in `smtplib` and `csv`.

---

## Run

```bash
python bulk_email_sender.py
```

---

## Author

Built by **Areeb** — Python automation developer.  
[GitHub](https://github.com/) • [PeoplePerHour](#)
