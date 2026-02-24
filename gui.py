import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import pandas as pd
from mailer import EmailSender

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class MailerGUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Pro Outreach Suite")
        self.geometry("850x800")
        
        self.db_path = ctk.StringVar()
        self.resume_path = ctk.StringVar()
        self.email_var = ctk.StringVar()
        self.password_var = ctk.StringVar()
        self.is_html_var = ctk.BooleanVar(value=False)
        self.template_name_var = ctk.StringVar(value="")
        
        self.sender = None
        self.is_running = False
        
        self.create_widgets()
        self.load_template_list()
        
    def create_widgets(self):
        # Setup Grid
        self.grid_columnconfigure(0, weight=1)
        
        # Header
        header = ctk.CTkLabel(self, text="Internship Outreach Suite", font=ctk.CTkFont(size=24, weight="bold"))
        header.grid(row=0, column=0, pady=(20, 10))

        # 1. File Selection Frame
        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        file_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(file_frame, text="1. Files", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        ctk.CTkButton(file_frame, text="Browse Database (CSV/Excel)", command=self.browse_db, width=200).grid(row=1, column=0, padx=10, pady=5)
        ctk.CTkEntry(file_frame, textvariable=self.db_path, state="readonly").grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkButton(file_frame, text="Browse Resume (PDF)", command=self.browse_resume, width=200).grid(row=2, column=0, padx=10, pady=5)
        ctk.CTkEntry(file_frame, textvariable=self.resume_path, state="readonly").grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # 2. SMTP Frame
        smtp_frame = ctk.CTkFrame(self)
        smtp_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        smtp_frame.grid_columnconfigure((1,3), weight=1)

        ctk.CTkLabel(smtp_frame, text="2. SMTP Login (App Password)", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        ctk.CTkLabel(smtp_frame, text="Email:").grid(row=1, column=0, padx=10, pady=5, sticky="e")
        ctk.CTkEntry(smtp_frame, textvariable=self.email_var).grid(row=1, column=1, padx=10, pady=5, sticky="ew")
        
        ctk.CTkLabel(smtp_frame, text="Password:").grid(row=1, column=2, padx=10, pady=5, sticky="e")
        ctk.CTkEntry(smtp_frame, textvariable=self.password_var, show="*").grid(row=1, column=3, padx=10, pady=5, sticky="ew")

        # 3. Template Frame
        tpl_frame = ctk.CTkFrame(self)
        tpl_frame.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")
        self.grid_rowconfigure(3, weight=1)
        tpl_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(tpl_frame, text="3. Template Configuration", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=10, pady=5, sticky="w")
        
        # Profile controls
        prof_frame = ctk.CTkFrame(tpl_frame, fg_color="transparent")
        prof_frame.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10)
        
        self.template_dropdown = ctk.CTkOptionMenu(prof_frame, values=["Current Profile"], command=self.load_selected_template)
        self.template_dropdown.pack(side="left", padx=(0, 10))
        
        self.template_name_entry = ctk.CTkEntry(prof_frame, textvariable=self.template_name_var, placeholder_text="Profile Name")
        self.template_name_entry.pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(prof_frame, text="Save Profile", command=self.save_template, width=100).pack(side="left", padx=(0, 10))
        
        ctk.CTkSwitch(prof_frame, text="Send as HTML", variable=self.is_html_var).pack(side="right", padx=10)

        # Text Area
        self.template_text = ctk.CTkTextbox(tpl_frame, height=200)
        self.template_text.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
        tpl_frame.grid_rowconfigure(2, weight=1)
        
        default_tpl = "Subject: Internship Inquiry - {Professor_Name}\n\nDear {Professor_Name},\n\nI am reaching out regarding my strong interest in {Research_Interest|your lab's recent publications}. I have attached my resume..."
        self.template_text.insert("0.0", default_tpl)

        # 4. Actions & Logs
        act_frame = ctk.CTkFrame(self, fg_color="transparent")
        act_frame.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="ew")
        
        ctk.CTkButton(act_frame, text="Preview First Match", command=self.preview_match, fg_color="gray", hover_color="darkgray").pack(side="left", padx=10)
        self.start_btn = ctk.CTkButton(act_frame, text="START SEQUENCE", command=self.toggle_process, fg_color="green", hover_color="darkgreen")
        self.start_btn.pack(side="right", padx=10)

        self.log_text = ctk.CTkTextbox(self, height=120, state="disabled")
        self.log_text.grid(row=5, column=0, padx=20, pady=(0, 20), sticky="ew")

    def browse_db(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV/Excel Files", "*.csv *.xlsx *.xls")])
        if filename:
            self.db_path.set(filename)

    def browse_resume(self):
        filename = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if filename:
            self.resume_path.set(filename)

    def log_message(self, message):
        self.log_text.configure(state='normal')
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.configure(state='disabled')

    def load_template_list(self):
        if not os.path.exists("templates"):
            os.makedirs("templates")
        
        files = [f.replace(".txt", "") for f in os.listdir("templates") if f.endswith(".txt")]
        if files:
            self.template_dropdown.configure(values=files)
            self.template_dropdown.set(files[0])
            self.load_selected_template(files[0])

    def load_selected_template(self, name):
        filepath = f"templates/{name}.txt"
        if os.path.exists(filepath):
            with open(filepath, "r", encoding="utf-8") as f:
                content = f.read()
            self.template_text.delete("0.0", "end")
            self.template_text.insert("0.0", content)
            self.template_name_var.set(name)

    def save_template(self):
        name = self.template_name_var.get().strip()
        if not name:
            messagebox.showerror("Error", "Please enter a profile name first.")
            return
            
        content = self.template_text.get("0.0", "end").strip()
        if not os.path.exists("templates"):
            os.makedirs("templates")
            
        with open(f"templates/{name}.txt", "w", encoding="utf-8") as f:
            f.write(content)
            
        self.log_message(f"Saved template profile: {name}")
        self.load_template_list()
        self.template_dropdown.set(name)

    def preview_match(self):
        db = self.db_path.get()
        template = self.template_text.get("0.0", "end").strip()
        if not db or not template:
            messagebox.showerror("Error", "Please select a database and write a template first.")
            return
            
        try:
            df = pd.read_csv(db) if db.endswith('.csv') else pd.read_excel(db)
            if df.empty:
                messagebox.showerror("Error", "Database is empty.")
                return
            
            row = df.iloc[0]
            dummy_sender = EmailSender("", "")
            subj, body = dummy_sender.generate_preview(row, template)
            
            # Preview Modal
            prev_win = ctk.CTkToplevel(self)
            prev_win.title("Preview (First Row)")
            prev_win.geometry("600x500")
            prev_win.attributes("-topmost", True)
            
            ctk.CTkLabel(prev_win, text=f"Subject: {subj}", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=(20, 10))
            
            prev_text = ctk.CTkTextbox(prev_win)
            prev_text.pack(expand=True, fill="both", padx=20, pady=(0, 20))
            prev_text.insert("0.0", body)
            prev_text.configure(state="disabled")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not generate preview: {e}")

    def toggle_process(self):
        if not self.is_running:
            self.start_process()
        else:
            self.stop_process()

    def start_process(self):
        db = self.db_path.get()
        email = self.email_var.get()
        pwd = self.password_var.get()
        template = self.template_text.get("0.0", "end").strip()
        resume = self.resume_path.get()
        is_html = self.is_html_var.get()

        if not db or not email or not pwd or not template:
            messagebox.showerror("Error", "Please fill in all required fields.")
            return

        self.is_running = True
        self.start_btn.configure(text="STOP SEQUENCE", fg_color="red", hover_color="darkred")
        
        self.log_text.configure(state='normal')
        self.log_text.delete("0.0", "end")
        self.log_text.configure(state='disabled')
        
        self.sender = EmailSender(smtp_email=email, smtp_password=pwd)
        
        thread = threading.Thread(target=self.run_mailer_thread, args=(db, template, resume, is_html))
        thread.daemon = True
        thread.start()

    def stop_process(self):
        if self.sender:
            self.sender.stop()
        self.is_running = False
        self.start_btn.configure(text="START SEQUENCE", fg_color="green", hover_color="darkgreen")
        self.log_message("Stopping process... (Please wait for current action to finish)")

    def run_mailer_thread(self, db, template, resume, is_html):
        self.log_message("Starting email automation...")
        self.sender.process_and_send(
            database_path=db,
            template_text=template,
            resume_path=resume,
            is_html=is_html,
            progress_callback=self.log_message
        )
        self.after(0, self.reset_button)

    def reset_button(self):
        self.is_running = False
        self.start_btn.configure(text="START SEQUENCE", fg_color="green", hover_color="darkgreen")

if __name__ == "__main__":
    app = MailerGUI()
    app.mainloop()
