import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import ttkbootstrap as ttb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText
from tkinter import font
from tkhtmlview import HTMLText

class EmailAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Email Automation Tool")
        self.root.geometry("800x900")
        
        # Variables
        self.csv_data = None
        self.csv_columns = []
        self.email_vars = {}
        
        # Create main container with scrolling
        main_canvas = tk.Canvas(self.root)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        self.main_container = ttk.Frame(main_canvas, padding="10")

        # Configure the canvas
        self.main_container.bind("<Configure>", lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all")))
        main_canvas.create_window((0, 0), window=self.main_container, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        # Pack the scrollbar and canvas
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # CSV File Selection
        self.create_csv_section()
        
        # Gmail Authentication
        self.create_auth_section()
        
        # Email Template Editor
        self.create_template_section()
        
        # Send Email Section
        self.create_send_section()
    
    def create_csv_section(self):
        csv_frame = ttk.LabelFrame(self.main_container, text="CSV File Selection", padding="10")
        csv_frame.pack(fill=tk.X, pady=5)
        
        self.csv_path_var = tk.StringVar()
        ttk.Entry(csv_frame, textvariable=self.csv_path_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(csv_frame, text="Browse", command=self.load_csv).pack(side=tk.RIGHT, padx=5)
    
    def create_auth_section(self):
        auth_frame = ttk.LabelFrame(self.main_container, text="Gmail Authentication", padding="10")
        auth_frame.pack(fill=tk.X, pady=5)
        
        # Email Address
        ttk.Label(auth_frame, text="Gmail Address:").pack(anchor=tk.W)
        self.email_var = tk.StringVar()
        ttk.Entry(auth_frame, textvariable=self.email_var).pack(fill=tk.X)
        
        # 2FA Option
        self.use_2fa = tk.BooleanVar()
        ttk.Checkbutton(auth_frame, text="Use 2FA (Application-specific password)", 
                       variable=self.use_2fa, command=self.toggle_2fa).pack(anchor=tk.W, pady=5)
        
        # Password Entry
        ttk.Label(auth_frame, text="Password:").pack(anchor=tk.W)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(auth_frame, textvariable=self.password_var, show='*')
        self.password_entry.pack(fill=tk.X)
    
    def create_template_section(self):
        template_frame = ttk.LabelFrame(self.main_container, text="Email Template", padding="10")
        template_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Subject Line
        ttk.Label(template_frame, text="Subject:").pack(anchor=tk.W)
        self.subject_var = tk.StringVar()
        ttk.Entry(template_frame, textvariable=self.subject_var).pack(fill=tk.X, pady=(0, 5))
        
        # Template Variables Frame with Scrolling
        vars_container = ttk.Frame(template_frame)
        vars_container.pack(fill=tk.X, pady=5)
        
        # Create a canvas for scrolling
        canvas = tk.Canvas(vars_container, height=150)
        scrollbar = ttk.Scrollbar(vars_container, orient="vertical", command=canvas.yview)
        self.vars_frame = ttk.LabelFrame(canvas, text="Template Variables", padding="5")
        
        # Configure canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.X, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create window in canvas for variables
        canvas_frame = canvas.create_window((0, 0), window=self.vars_frame, anchor=tk.NW)
        
        # Update canvas scroll region when frame size changes
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_frame, width=canvas.winfo_width())
        
        self.vars_frame.bind("<Configure>", on_frame_configure)
        
        # Email Body Editor with Formatting Toolbar
        ttk.Label(template_frame, text="Email Body:").pack(anchor=tk.W)
        
        # Formatting Toolbar
        toolbar = ttk.Frame(template_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))
        

        
        # Hyperlink Button
        self.link_btn = ttk.Button(toolbar, text="🔗", width=3, style="Toolbutton")
        self.link_btn.pack(side=tk.LEFT, padx=2)
        
        # Rich Text Editor
        self.body_text = HTMLText(template_frame, height=15)
        self.body_text.pack(fill=tk.BOTH, expand=True)
        
        # Bind Events
        self.link_btn.config(command=self.insert_hyperlink)

    def insert_hyperlink(self):
        try:
            selection = self.body_text.get("sel.first", "sel.last")
            url = tk.simpledialog.askstring("Insert Hyperlink", "Enter URL:")
            if url:
                self.body_text.tag_add("hyperlink", "sel.first", "sel.last")
                self.body_text.tag_config("hyperlink", foreground="blue", underline=True)
                self.body_text.tag_bind("hyperlink", "<Button-1>", lambda e: self.open_url(url))
        except tk.TclError:
            pass

    def open_url(self, url):
        import webbrowser
        webbrowser.open(url)

    def create_send_section(self):
        send_frame = ttk.LabelFrame(self.main_container, text="Send Options", padding="10")
        send_frame.pack(fill=tk.X, pady=10)
        
        # Email Column Selection
        email_col_frame = ttk.Frame(send_frame)
        email_col_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(email_col_frame, text="Select Email Column:").pack(side=tk.LEFT, padx=(0, 5))
        self.email_column_var = tk.StringVar()
        self.email_column_combo = ttk.Combobox(email_col_frame, textvariable=self.email_column_var, state="readonly", height=5)
        self.email_column_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Progress and Send Button
        progress_frame = ttk.Frame(send_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress_var = tk.StringVar(value="Ready to send emails")
        ttk.Label(progress_frame, textvariable=self.progress_var).pack(side=tk.LEFT)
        ttk.Button(progress_frame, text="Send Emails", command=self.send_emails).pack(side=tk.RIGHT)
    
    def load_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                self.csv_data = pd.read_csv(file_path)
                self.csv_path_var.set(file_path)
                self.csv_columns = list(self.csv_data.columns)
                self.update_template_vars()
                messagebox.showinfo("Success", "CSV file loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV file: {str(e)}")
    
    def update_template_vars(self):
        # Clear existing variables
        for widget in self.vars_frame.winfo_children():
            widget.destroy()
        
        # Create buttons for each column
        for i, col in enumerate(self.csv_columns):
            btn = ttk.Button(self.vars_frame, text=f"{{{col}}}",
                            command=lambda c=col: self.insert_template_var(c))
            btn.grid(row=i//3, column=i%3, padx=2, pady=2, sticky='ew')
            
        # Update email column combobox with available email columns
        email_columns = [col for col in self.csv_columns if 'email' in col.lower()]
        if email_columns:
            self.email_column_combo['values'] = email_columns
            self.email_column_combo['values'] = email_columns
    
    def insert_template_var(self, column):
        self.body_text.insert(tk.INSERT, f"{{{column}}}")
    
    def toggle_2fa(self):
        if self.use_2fa.get():
            self.password_entry.configure(show='*')
        else:
            self.password_entry.configure(show='*')
    
    def send_emails(self):
        if not self.validate_inputs():
            return
        
        try:
            # Connect to Gmail
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(self.email_var.get(), self.password_var.get())
            
            total_rows = len(self.csv_data)
            selected_email_column = self.email_column_var.get()
            
            for index, row in self.csv_data.iterrows():
                try:
                    # Create message
                    msg = MIMEMultipart('alternative')
                    msg['From'] = self.email_var.get()
                    msg['To'] = row[selected_email_column]
                    
                    # Format subject and body
                    subject = self.subject_var.get()
                    body = self.body_text.get("1.0", tk.END)
                    
                    # Replace template variables
                    for col in self.csv_columns:
                        subject = subject.replace(f"{{{col}}}", str(row[col]))
                        body = body.replace(f"{{{col}}}", str(row[col]))
                    
                    msg['Subject'] = subject
                    msg.attach(MIMEText(body, 'html'))
                    
                    # Send email
                    server.send_message(msg)
                    
                    # Update progress
                    self.progress_var.set(f"Sent {index + 1} of {total_rows} emails")
                    self.root.update()
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to send email to {row[selected_email_column]}: {str(e)}")
            
            server.quit()
            messagebox.showinfo("Success", "All emails sent successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to connect to Gmail: {str(e)}")
    
    def validate_inputs(self):
        if not self.csv_data is not None:
            messagebox.showerror("Error", "Please load a CSV file first")
            return False
        
        if not self.email_var.get() or not self.password_var.get():
            messagebox.showerror("Error", "Please enter email and password")
            return False
        
        if not self.subject_var.get() or not self.body_text.get('1.0', tk.END).strip():
            messagebox.showerror("Error", "Please enter email subject and body")
            return False
        
        return True

if __name__ == "__main__":
    root = ttb.Window(themename="cosmo")
    app = EmailAutomationApp(root)
    root.mainloop()