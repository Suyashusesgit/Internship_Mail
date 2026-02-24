import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
from mailer import EmailSender

class MailerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Internship Email Automation Tool")
        self.root.geometry("700x700")
        
        self.db_path = tk.StringVar()
        self.resume_path = tk.StringVar()
        self.email_var = tk.StringVar()
        self.password_var = tk.StringVar()
        
        self.sender = None
        self.is_running = False
        
        self.create_widgets()
        
    def create_widgets(self):
        # Database Selection
        tk.Label(self.root, text="Step 1: Select Database (CSV/Excel)").pack(anchor="w", padx=10, pady=(10, 0))
        db_frame = tk.Frame(self.root)
        db_frame.pack(fill="x", padx=10)
        tk.Entry(db_frame, textvariable=self.db_path, width=60, state='readonly').pack(side="left", padx=(0, 10))
        tk.Button(db_frame, text="Browse", command=self.browse_db).pack(side="left")

        # Resume Selection
        tk.Label(self.root, text="Step 2: Select Resume (PDF)").pack(anchor="w", padx=10, pady=(10, 0))
        resume_frame = tk.Frame(self.root)
        resume_frame.pack(fill="x", padx=10)
        tk.Entry(resume_frame, textvariable=self.resume_path, width=60, state='readonly').pack(side="left", padx=(0, 10))
        tk.Button(resume_frame, text="Browse", command=self.browse_resume).pack(side="left")

        # SMTP Credentials
        tk.Label(self.root, text="Step 3: SMTP Credentials (Gmail App Password)").pack(anchor="w", padx=10, pady=(10, 0))
        cred_frame = tk.Frame(self.root)
        cred_frame.pack(fill="x", padx=10)
        tk.Label(cred_frame, text="Email:").pack(side="left")
        tk.Entry(cred_frame, textvariable=self.email_var, width=25).pack(side="left", padx=5)
        tk.Label(cred_frame, text="Password:").pack(side="left", padx=(10, 0))
        tk.Entry(cred_frame, textvariable=self.password_var, show="*", width=25).pack(side="left", padx=5)

        # Template Area
        tk.Label(self.root, text="Step 4: Email Template (e.g., Dear {Professor_Name})").pack(anchor="w", padx=10, pady=(10, 0))
        self.template_text = scrolledtext.ScrolledText(self.root, width=80, height=12)
        self.template_text.pack(padx=10, pady=5)
        self.template_text.insert(tk.END, "Subject: Internship Inquiry\n\nDear {Professor_Name},\n\nI am a B.Tech student exploring your work in {Research_Interest}...\n\nBest,\n[Your Name]")

        # Action Buttons
        btn_frame = tk.Frame(self.root)
        btn_frame.pack(pady=10)
        self.start_btn = tk.Button(btn_frame, text="START", width=15, bg="green", fg="white", font=("Arial", 12, "bold"), command=self.toggle_process)
        self.start_btn.pack()

        # Log Area
        tk.Label(self.root, text="Process Logs:").pack(anchor="w", padx=10)
        self.log_text = scrolledtext.ScrolledText(self.root, width=80, height=10, state='disabled')
        self.log_text.pack(padx=10, pady=5)

    def browse_db(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV/Excel Files", "*.csv *.xlsx *.xls")])
        if filename:
            self.db_path.set(filename)

    def browse_resume(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename:
            self.resume_path.set(filename)

    def log_message(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
        self.root.update_idletasks()

    def toggle_process(self):
        if not self.is_running:
            self.start_process()
        else:
            self.stop_process()

    def start_process(self):
        db = self.db_path.get()
        email = self.email_var.get()
        pwd = self.password_var.get()
        template = self.template_text.get("1.0", tk.END).strip()
        resume = self.resume_path.get()

        if not db or not email or not pwd or not template:
            messagebox.showerror("Error", "Please fill in all required fields (Database, Email, Password, Template).")
            return

        self.is_running = True
        self.start_btn.config(text="STOP", bg="red")
        
        self.log_text.config(state='normal')
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state='disabled')
        
        self.sender = EmailSender(smtp_email=email, smtp_password=pwd)
        
        # Run in a separate thread to prevent GUI freezing
        thread = threading.Thread(target=self.run_mailer_thread, args=(db, template, resume))
        thread.daemon = True
        thread.start()

    def stop_process(self):
        if self.sender:
            self.sender.stop()
        self.is_running = False
        self.start_btn.config(text="START", bg="green")
        self.log_message("Stopping process... (Please wait for the current action to finish)")

    def run_mailer_thread(self, db, template, resume):
        self.log_message("Starting email automation...")
        self.sender.process_and_send(
            database_path=db,
            template_text=template,
            resume_path=resume,
            progress_callback=self.log_message
        )
        # Reset button when finished
        self.root.after(0, self.reset_button)

    def reset_button(self):
        self.is_running = False
        self.start_btn.config(text="START", bg="green")
