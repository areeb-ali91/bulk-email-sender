# ============================================================
# BULK EMAIL SENDER v1.0
# Send personalized emails from a CSV contact list
# Clean dark Tkinter GUI with preview and progress tracking
# by Areeb
# ============================================================

import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
import csv
import smtplib
import threading
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ============================================================
# THEME
# ============================================================
T = {
    "bg":      "#0f1923",
    "bg2":     "#1a2634",
    "bg3":     "#243447",
    "accent":  "#2196f3",
    "accent2": "#1565c0",
    "green":   "#00e676",
    "red":     "#ff5252",
    "orange":  "#ff9800",
    "text":    "#e8f0fe",
    "text2":   "#90a4ae",
    "border":  "#243447",
}

FONT_HEAD  = ("Segoe UI", 14, "bold")
FONT_SUB   = ("Segoe UI", 8)
FONT_LABEL = ("Segoe UI", 9, "bold")
FONT_BODY  = ("Segoe UI", 9)
FONT_MONO  = ("Consolas", 9)

# ============================================================
# EMAIL LOGIC
# ============================================================
def send_emails(smtp_host, smtp_port, sender_email, password,
                contacts, subject_template, body_template,
                log_cb, progress_cb, done_cb):

    total   = len(contacts)
    sent    = 0
    failed  = 0

    try:
        server = smtplib.SMTP(smtp_host, smtp_port, timeout=10)
        server.starttls()
        server.login(sender_email, password)
        log_cb(f"✅  Logged in as {sender_email}", "green")
    except Exception as e:
        log_cb(f"❌  Login failed: {e}", "red")
        done_cb(0, total)
        return

    for i, contact in enumerate(contacts):
        try:
            # Personalize subject and body
            subject = subject_template
            body    = body_template
            for key, val in contact.items():
                placeholder = f"{{{key}}}"
                subject = subject.replace(placeholder, val)
                body    = body.replace(placeholder, val)

            msg = MIMEMultipart("alternative")
            msg["From"]    = sender_email
            msg["To"]      = contact.get("email", "")
            msg["Subject"] = subject

            msg.attach(MIMEText(body, "plain"))

            server.sendmail(sender_email, contact["email"], msg.as_string())
            log_cb(f"  ✓  Sent to {contact['email']}", "green")
            sent += 1

        except Exception as e:
            log_cb(f"  ✕  Failed for {contact.get('email', '?')}: {e}", "red")
            failed += 1

        progress_cb(i + 1, total)
        time.sleep(0.3)  # rate limit

    try:
        server.quit()
    except Exception:
        pass

    done_cb(sent, failed)


# ============================================================
# GUI
# ============================================================
class BulkEmailApp:
    def __init__(self, root):
        self.root     = root
        self.root.title("Bulk Email Sender")
        self.root.geometry("700x760")
        self.root.resizable(False, False)
        self.root.configure(bg=T["bg"])

        self.contacts      = []
        self.csv_path_var  = tk.StringVar(value="No file selected")
        self.status_var    = tk.StringVar(value="Ready")
        self.progress_var  = tk.IntVar(value=0)

        self._build()

    # ── BUILD ────────────────────────────────────────────────
    def _build(self):
        # HEADER
        header = tk.Frame(self.root, bg=T["accent"], pady=14)
        header.pack(fill=tk.X)
        tk.Label(header, text="Bulk Email Sender",
                 font=FONT_HEAD, bg=T["accent"], fg=T["text"]).pack()
        tk.Label(header, text="Send personalized emails to your contact list",
                 font=FONT_SUB, bg=T["accent"], fg=T["text2"]).pack()

        # SCROLL CONTAINER
        canvas   = tk.Canvas(self.root, bg=T["bg"],
                             highlightthickness=0)
        scrollbar = tk.Scrollbar(self.root, orient="vertical",
                                  command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.main = tk.Frame(canvas, bg=T["bg"])
        canvas.create_window((0,0), window=self.main, anchor="nw")
        self.main.bind("<Configure>", lambda e: canvas.configure(
            scrollregion=canvas.bbox("all")))

        pad = {"padx": 20, "pady": 6}

        # ── SMTP SETTINGS ────────────────────────────────────
        self._section("SMTP Settings")

        self.smtp_host  = self._field("SMTP Host",  "smtp.gmail.com")
        self.smtp_port  = self._field("SMTP Port",  "587")
        self.smtp_email = self._field("Your Email", "you@gmail.com")
        self.smtp_pass  = self._field("App Password", "", show="*")

        hint = tk.Label(self.main,
            text="💡  Use an App Password (Google: myaccount.google.com → Security → App Passwords)",
            font=("Segoe UI", 8), bg=T["bg"], fg=T["text2"],
            anchor="w", wraplength=640)
        hint.pack(fill=tk.X, **pad)

        self._divider()

        # ── CSV CONTACTS ─────────────────────────────────────
        self._section("Contact List (CSV)")

        csv_row = tk.Frame(self.main, bg=T["bg"])
        csv_row.pack(fill=tk.X, **pad)

        self.csv_lbl = tk.Label(csv_row,
            textvariable=self.csv_path_var,
            font=FONT_MONO, bg=T["bg2"], fg=T["text2"],
            anchor="w", padx=10)
        self.csv_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True, ipady=6)

        tk.Button(csv_row, text="Browse CSV",
                  font=FONT_BODY, bg=T["bg3"], fg=T["text"],
                  relief="flat", bd=0, padx=12, pady=6,
                  cursor="hand2",
                  activebackground=T["accent"],
                  activeforeground=T["text"],
                  command=self._load_csv).pack(side=tk.LEFT, padx=(6,0))

        csv_hint = tk.Label(self.main,
            text='💡  CSV must have a column named "email". Other columns (name, company, etc.) become placeholders.',
            font=("Segoe UI", 8), bg=T["bg"], fg=T["text2"],
            anchor="w", wraplength=640)
        csv_hint.pack(fill=tk.X, **pad)

        self.contact_count = tk.Label(self.main, text="",
            font=FONT_BODY, bg=T["bg"], fg=T["green"])
        self.contact_count.pack(anchor="w", **pad)

        self._divider()

        # ── EMAIL TEMPLATE ───────────────────────────────────
        self._section("Email Template")

        subj_row = tk.Frame(self.main, bg=T["bg"])
        subj_row.pack(fill=tk.X, **pad)
        tk.Label(subj_row, text="Subject:",
                 font=FONT_LABEL, bg=T["bg"], fg=T["text"],
                 width=10, anchor="w").pack(side=tk.LEFT)
        self.subject_entry = tk.Entry(subj_row,
            font=FONT_BODY, bg=T["bg2"], fg=T["text"],
            insertbackground=T["text"],
            relief="flat", bd=6)
        self.subject_entry.insert(0, "Hello {name}, here's something for you")
        self.subject_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Label(self.main, text="Body: (use {name}, {email}, {company} etc. as placeholders)",
                 font=FONT_LABEL, bg=T["bg"], fg=T["text"],
                 anchor="w").pack(fill=tk.X, padx=20, pady=(8,2))

        self.body_text = scrolledtext.ScrolledText(
            self.main, height=7,
            font=FONT_MONO, bg=T["bg2"], fg=T["text"],
            insertbackground=T["text"],
            relief="flat", bd=8,
            wrap=tk.WORD)
        self.body_text.insert("1.0",
            "Hi {name},\n\n"
            "I hope this message finds you well.\n\n"
            "This is a personalized message just for you at {company}.\n\n"
            "Best regards,\nAreeb")
        self.body_text.pack(fill=tk.X, padx=20, pady=(0,8))

        self._divider()

        # ── PREVIEW ──────────────────────────────────────────
        self._section("Preview (first contact)")

        self.preview_btn = tk.Button(self.main, text="👁  Preview First Email",
            font=FONT_BODY, bg=T["bg3"], fg=T["text"],
            relief="flat", bd=0, padx=14, pady=6,
            cursor="hand2",
            activebackground=T["accent"],
            activeforeground=T["text"],
            command=self._preview)
        self.preview_btn.pack(anchor="w", padx=20, pady=(4,6))

        self.preview_box = tk.Text(self.main, height=5,
            font=FONT_MONO, bg=T["bg2"], fg=T["orange"],
            relief="flat", bd=8,
            state=tk.DISABLED, wrap=tk.WORD)
        self.preview_box.pack(fill=tk.X, padx=20, pady=(0,8))

        self._divider()

        # ── SEND ─────────────────────────────────────────────
        btn_row = tk.Frame(self.main, bg=T["bg"])
        btn_row.pack(fill=tk.X, padx=20, pady=8)

        self.send_btn = tk.Button(btn_row,
            text="🚀  Send Emails",
            font=("Segoe UI", 11, "bold"),
            bg=T["accent"], fg=T["text"],
            relief="flat", bd=0,
            pady=10, cursor="hand2",
            activebackground=T["accent2"],
            activeforeground=T["text"],
            command=self._start_send)
        self.send_btn.pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(btn_row, text="✕ Clear Log",
            font=FONT_BODY, bg=T["bg3"], fg=T["text2"],
            relief="flat", bd=0, padx=14, pady=10,
            cursor="hand2",
            activebackground=T["bg2"],
            command=self._clear_log).pack(side=tk.LEFT, padx=(8,0))

        # PROGRESS
        self.prog_bar = tk.Canvas(self.main, height=6,
            bg=T["bg3"], highlightthickness=0)
        self.prog_bar.pack(fill=tk.X, padx=20, pady=(4,2))
        self._prog_rect = self.prog_bar.create_rectangle(
            0,0,0,6, fill=T["accent"], outline="")

        tk.Label(self.main, textvariable=self.status_var,
                 font=("Segoe UI", 8), bg=T["bg"], fg=T["text2"],
                 anchor="w").pack(fill=tk.X, padx=20)

        # LOG
        tk.Label(self.main, text="Activity Log",
                 font=FONT_LABEL, bg=T["bg"], fg=T["text"],
                 anchor="w").pack(fill=tk.X, padx=20, pady=(10,2))

        self.log_box = scrolledtext.ScrolledText(
            self.main, height=8,
            font=FONT_MONO, bg=T["bg2"], fg=T["text"],
            relief="flat", bd=8,
            state=tk.DISABLED, wrap=tk.WORD)
        self.log_box.tag_config("green",  foreground=T["green"])
        self.log_box.tag_config("red",    foreground=T["red"])
        self.log_box.tag_config("accent", foreground=T["accent"])
        self.log_box.tag_config("orange", foreground=T["orange"])
        self.log_box.pack(fill=tk.X, padx=20, pady=(0,20))

    # ── HELPERS ──────────────────────────────────────────────
    def _section(self, title):
        tk.Label(self.main, text=title,
                 font=("Segoe UI", 10, "bold"),
                 bg=T["bg"], fg=T["text"],
                 anchor="w").pack(fill=tk.X, padx=20, pady=(12,4))

    def _divider(self):
        tk.Frame(self.main, bg=T["bg3"], height=1).pack(
            fill=tk.X, padx=20, pady=6)

    def _field(self, label, default="", show=""):
        row = tk.Frame(self.main, bg=T["bg"])
        row.pack(fill=tk.X, padx=20, pady=3)
        tk.Label(row, text=label,
                 font=FONT_LABEL, bg=T["bg"], fg=T["text"],
                 width=14, anchor="w").pack(side=tk.LEFT)
        entry = tk.Entry(row, font=FONT_BODY,
                         bg=T["bg2"], fg=T["text"],
                         insertbackground=T["text"],
                         show=show,
                         relief="flat", bd=6)
        entry.insert(0, default)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        return entry

    # ── ACTIONS ──────────────────────────────────────────────
    def _load_csv(self):
        path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, newline="", encoding="utf-8") as f:
                reader = csv.DictReader(f)
                self.contacts = [row for row in reader]

            if not self.contacts:
                messagebox.showwarning("Empty CSV", "No contacts found.")
                return
            if "email" not in self.contacts[0]:
                messagebox.showerror("Missing Column",
                    'CSV must have a column named "email".')
                self.contacts = []
                return

            self.csv_path_var.set(path.split("/")[-1])
            self.contact_count.config(
                text=f"✅  {len(self.contacts)} contact(s) loaded.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read CSV:\n{e}")

    def _preview(self):
        if not self.contacts:
            self.preview_box.config(state=tk.NORMAL)
            self.preview_box.delete("1.0", tk.END)
            self.preview_box.insert("1.0", "Load a CSV first.")
            self.preview_box.config(state=tk.DISABLED)
            return

        contact = self.contacts[0]
        subject = self.subject_entry.get()
        body    = self.body_text.get("1.0", tk.END)
        for key, val in contact.items():
            subject = subject.replace(f"{{{key}}}", val)
            body    = body.replace(f"{{{key}}}", val)

        preview = f"To: {contact.get('email','')}\nSubject: {subject}\n\n{body}"
        self.preview_box.config(state=tk.NORMAL)
        self.preview_box.delete("1.0", tk.END)
        self.preview_box.insert("1.0", preview)
        self.preview_box.config(state=tk.DISABLED)

    def _log(self, msg, tag="text"):
        self.log_box.config(state=tk.NORMAL)
        self.log_box.insert(tk.END, msg + "\n", tag)
        self.log_box.see(tk.END)
        self.log_box.config(state=tk.DISABLED)

    def _clear_log(self):
        self.log_box.config(state=tk.NORMAL)
        self.log_box.delete("1.0", tk.END)
        self.log_box.config(state=tk.DISABLED)

    def _update_progress(self, done, total):
        def _update():
            w = self.prog_bar.winfo_width()
            frac = done / total if total else 0
            self.prog_bar.coords(self._prog_rect, 0, 0, w * frac, 6)
            self.status_var.set(f"Sending... {done}/{total}")
        self.root.after(0, _update)

    def _start_send(self):
        if not self.contacts:
            messagebox.showwarning("No Contacts", "Please load a CSV file first.")
            return

        host     = self.smtp_host.get().strip()
        port     = self.smtp_port.get().strip()
        email    = self.smtp_email.get().strip()
        password = self.smtp_pass.get().strip()
        subject  = self.subject_entry.get().strip()
        body     = self.body_text.get("1.0", tk.END).strip()

        if not all([host, port, email, password, subject, body]):
            messagebox.showwarning("Missing Info",
                "Please fill in all SMTP settings and the email template.")
            return

        if not messagebox.askyesno("Confirm",
            f"Send to {len(self.contacts)} contact(s)?\nThis cannot be undone."):
            return

        self.send_btn.config(state=tk.DISABLED, text="Sending...")
        self._clear_log()
        self._log(f"Starting — {len(self.contacts)} email(s) to send", "accent")

        def _log_safe(msg, tag): self.root.after(0, self._log, msg, tag)
        def _prog_safe(d, t):    self._update_progress(d, t)
        def _done(sent, failed):
            def _f():
                self._log(f"\nDone! ✅ {sent} sent  ❌ {failed} failed", "green")
                self.status_var.set(f"Done — {sent} sent, {failed} failed.")
                self.send_btn.config(state=tk.NORMAL, text="🚀  Send Emails")
            self.root.after(0, _f)

        threading.Thread(
            target=send_emails,
            args=(host, int(port), email, password,
                  self.contacts, subject, body,
                  _log_safe, _prog_safe, _done),
            daemon=True
        ).start()


# ── RUN ─────────────────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app  = BulkEmailApp(root)
    root.mainloop()
