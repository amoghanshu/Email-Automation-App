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
        
        # Enable window resizing
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        
        # Variables
        self.csv_data = None
        self.csv_columns = []
        self.email_vars = {}
        
        # Create main container with scrolling
        self.main_canvas = tk.Canvas(self.root)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.main_canvas.yview)
        self.main_container = ttk.Frame(self.main_canvas)

        # Configure the canvas
        self.main_container.bind("<Configure>", self._on_frame_configure)
        self.main_canvas.create_window((0, 0), window=self.main_container, anchor="nw")
        self.main_canvas.configure(yscrollcommand=self.scrollbar.set)

        # Pack the scrollbar and canvas with proper expansion
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.main_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Configure main_container for resizing
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.main_container.columnconfigure(0, weight=1)
        
        # Bind mousewheel to canvas
        self.main_canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.main_canvas.bind_all("<Shift-MouseWheel>", lambda event: self.main_canvas.xview_scroll(int(-1*(event.delta/120)), "units"))
        
        # Bind window resize event
        self.root.bind("<Configure>", lambda e: self._on_frame_configure())
        
        self.main_container.rowconfigure(0, weight=1)
        
        # CSV File Selection
        self.create_csv_section()
        
        # Gmail Authentication
        self.create_auth_section()
        
        # Email Template Editor
        self.create_template_section()
        
        # Send Email Section
        self.create_send_section()

    def _on_frame_configure(self, event=None):
        # Update the scroll region to encompass the entire frame
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))
        # Ensure the canvas width matches the window width
        width = self.main_canvas.winfo_width()
        if width > 1: # Only update if width is valid
            self.main_canvas.itemconfig(1, width=width)

    def _on_mousewheel(self, event):
        # Platform-specific scroll speed adjustment
        if event.num == 5 or event.delta < 0:
            self.main_canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:
            self.main_canvas.yview_scroll(-1, "units")
    
    def create_csv_section(self):
        csv_frame = ttk.LabelFrame(self.main_container, text="CSV File Selection", padding="10")
        csv_frame.pack(fill=tk.BOTH, pady=5)
        csv_frame.columnconfigure(0, weight=1)
        
        entry_frame = ttk.Frame(csv_frame)
        entry_frame.pack(fill=tk.X, expand=True)
        entry_frame.columnconfigure(0, weight=1)
        
        self.csv_path_var = tk.StringVar()
        ttk.Entry(entry_frame, textvariable=self.csv_path_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(entry_frame, text="Browse", command=self.load_csv).pack(side=tk.RIGHT, padx=5)
    
    def create_auth_section(self):
        auth_frame = ttk.LabelFrame(self.main_container, text="Gmail Authentication", padding="10")
        auth_frame.pack(fill=tk.BOTH, pady=5)
        auth_frame.columnconfigure(0, weight=1)
        
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
        template_frame.columnconfigure(0, weight=1)
        
        # Subject Line
        ttk.Label(template_frame, text="Subject:").pack(anchor=tk.W)
        self.subject_var = tk.StringVar()
        ttk.Entry(template_frame, textvariable=self.subject_var).pack(fill=tk.X, pady=(0, 5))
        
        # Template Variables Frame with Scrolling
        vars_container = ttk.Frame(template_frame)
        vars_container.pack(fill=tk.X, pady=5)
        vars_container.columnconfigure(0, weight=1)
        
        # Create a canvas for scrolling
        canvas = tk.Canvas(vars_container, height=150)
        
        # Horizontal scrollbar
        x_scrollbar = ttk.Scrollbar(vars_container, orient="horizontal", command=canvas.xview)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Vertical scrollbar
        y_scrollbar = ttk.Scrollbar(vars_container, orient="vertical", command=canvas.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Pack canvas after scrollbars
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.vars_frame = ttk.LabelFrame(canvas, text="Template Variables", padding="5")
        self.vars_frame.columnconfigure((0,1,2), weight=1)
        
        # Configure canvas with both scrollbars
        canvas.configure(xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
        
        # Create window in canvas for variables and set its width
        canvas_frame = canvas.create_window((0, 0), window=self.vars_frame, anchor=tk.NW, width=canvas.winfo_width())
        
        # Update canvas scroll region and frame width when size changes
        def on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_frame, width=canvas.winfo_width())
        
        def on_canvas_configure(event):
            canvas.itemconfig(canvas_frame, width=event.width)
        
        self.vars_frame.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", on_canvas_configure)
        
        # Email Body Editor with Formatting Toolbar
        ttk.Label(template_frame, text="Email Body:").pack(anchor=tk.W)
        
        # Formatting Toolbar
        toolbar = ttk.Frame(template_frame)
        toolbar.pack(fill=tk.X, pady=(0, 5))
        toolbar.columnconfigure(1, weight=1)
        
        # Hyperlink Button
        self.link_btn = ttk.Button(toolbar, text="🔗", width=3, style="Toolbutton")
        self.link_btn.pack(side=tk.LEFT, padx=2)
        
        # Rich Text Editor
        self.body_text = HTMLText(template_frame, height=15)
        self.body_text.pack(fill=tk.BOTH, expand=True, pady=5)

    def create_send_section(self):
        send_frame = ttk.LabelFrame(self.main_container, text="Send Options", padding="10")
        send_frame.pack(fill=tk.X, pady=10)
        send_frame.columnconfigure(0, weight=1)
        
        # Email Column Selection
        email_col_frame = ttk.Frame(send_frame)
        email_col_frame.pack(fill=tk.X, pady=5)
        email_col_frame.columnconfigure(1, weight=1)
        
        ttk.Label(email_col_frame, text="Select Email Column:").pack(side=tk.LEFT, padx=(0, 5))
        self.email_column_var = tk.StringVar()
        self.email_column_combo = ttk.Combobox(email_col_frame, textvariable=self.email_column_var, state="readonly", height=5)
        self.email_column_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Progress and Send Button
        progress_frame = ttk.Frame(send_frame)
        progress_frame.pack(fill=tk.X, pady=5)
        progress_frame.columnconfigure(0, weight=1)
        
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
            self.email_column_combo.set(email_columns[0] if email_columns else '')
    
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
if __name__ == "__main__":
    root = ttb.Window(themename="litera")
    app = EmailAutomationApp(root)
    root.mainloop()