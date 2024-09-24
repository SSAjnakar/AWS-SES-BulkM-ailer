import tkinter as tk
import tkinter.ttk as ttk
import tkinter.scrolledtext as scrolledtext
import threading
import time
import os
import boto3
import pandas as pd
import queue
from tkinter import filedialog, messagebox
from botocore.exceptions import ClientError

# AWS SES Configuration
AWS_REGION = "XX"
AWS_ACCESS_KEY_ID = "XXX"
AWS_SECRET_ACCESS_KEY = "XXX"

class ImprovedEmailClientGUI:
    def __init__(self, master):
        self.master = master
        master.title("© add your company name or  name of application")
        master.geometry("800x600")
        master.style = ttk.Style()
        master.style.theme_use("clam")  # Use themed style

        self.contact_files = []
        self.content_files = []
        self.sending_emails = False
        self.progress_queue = queue.Queue()

        # Initialize AWS SES client
        self.ses_client = boto3.client('ses',
                                       region_name=AWS_REGION,
                                       aws_access_key_id=AWS_ACCESS_KEY_ID,
                                       aws_secret_access_key=AWS_SECRET_ACCESS_KEY)

        self.create_widgets()
        # Status Bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.master, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var.set("Ready to send emails.")

        # Progress Bar
        self.progress_bar = ttk.Progressbar(self.master, orient=tk.HORIZONTAL, length=300, mode='indeterminate')
        self.progress_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.progress_bar.start(21) 

    def create_widgets(self):
        # Create a notebook for tabbed interface
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        # File Selection Tab
        self.file_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.file_frame, text="File Selection")

        # Contact Files
        ttk.Label(self.file_frame, text="Contact Files").grid(row=0, column=0, sticky="w", pady=10, padx=5)
        self.contact_listbox = tk.Listbox(self.file_frame, width=50, height=5)
        self.contact_listbox.grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(self.file_frame, text="Browse", command=self.browse_contact_file).grid(row=2, column=0, pady=5)
        ttk.Button(self.file_frame, text="Remove", command=lambda: self.remove_file('contact')).grid(row=2, column=1, pady=5)

        # Content Files
        ttk.Label(self.file_frame, text="Content Files").grid(row=3, column=0, sticky="w", pady=10, padx=5)
        self.content_listbox = tk.Listbox(self.file_frame, width=50, height=5)
        self.content_listbox.grid(row=4, column=0, padx=5, pady=5)
        ttk.Button(self.file_frame, text="Browse", command=self.browse_content_file).grid(row=5, column=0, pady=5)
        ttk.Button(self.file_frame, text="Remove", command=lambda: self.remove_file('content')).grid(row=5, column=1, pady=5)

        # Add some space between widgets
        ttk.Separator(self.file_frame, orient=tk.HORIZONTAL).grid(row=6, columnspan=2, sticky="ew", pady=10)

        # Email Configuration Tab
        self.config_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.config_frame, text="Email Configuration")

        ttk.Label(self.config_frame, text="Sender Email:").grid(row=1, column=0, sticky="w", pady=5, padx=5)
        self.email_entry = ttk.Entry(self.config_frame, width=50)
        self.email_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(self.config_frame, text="Delay (seconds):").grid(row=2, column=0, sticky="w", pady=5, padx=5)
        self.delay_entry = ttk.Entry(self.config_frame, width=10)
        self.delay_entry.grid(row=2, column=1, padx=5, pady=5)
        self.delay_entry.insert(0, "5")  # Default delay

        # "Next" Button to switch to Email Configuration
        ttk.Button(self.file_frame, text="Next", command=lambda: self.notebook.select(self.config_frame)).grid(row=7, column=0, pady=10)

        # Logs Tab
        self.logs_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.logs_frame, text="Error Logs")

        self.log_text = scrolledtext.ScrolledText(self.logs_frame, wrap=tk.WORD, width=80, height=20)
        self.log_text.pack(padx=10, pady=10, expand=True, fill="both")

        # Start Process Button in Email Configuration Tab
        self.start_button = ttk.Button(self.config_frame, text="Start Process", command=self.start_process)
        self.start_button.grid(row=3, column=1, pady=10)

        self.refresh_button = ttk.Button(self.config_frame, text="Refresh", command=self.refresh_gui)
        self.refresh_button.grid(row=3, column=0, pady=10)

        # Refresh File Selection Button
        #self.refresh_file_button = ttk.Button(self.file_frame, text="Refresh Files", command=self.refresh_file_selection)
        self.refresh_file_button = ttk.Button(self.file_frame, text="Refresh", command=self.refresh_file_selection)
    
        self.refresh_file_button.grid(row=8, column=0, pady=10)

    def refresh_file_selection(self):
        # Clear the file lists and update the listboxes
        self.contact_files.clear()
        self.content_files.clear()
        self.update_listbox(self.contact_listbox, self.contact_files)
        self.update_listbox(self.content_listbox, self.content_files)

        self.status_var.set("File selection refreshed.")
        self.log_text.delete(1.0, tk.END)  # Clear previous logs

    def refresh_gui(self):
        self.sending_emails = False
        self.start_button.config(state='normal')
        self.status_var.set("Ready to send emails.")
        self.log_text.delete(1.0, tk.END)  # Clear previous logs
        self.reset_gui()
        self.progress_bar.stop()  # Stop the progress bar

    def browse_contact_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path and file_path not in self.contact_files:
            self.contact_files.append(file_path)
            self.update_listbox(self.contact_listbox, self.contact_files)
        elif file_path in self.contact_files:
            messagebox.showwarning("Warning", "Contact file already added.")
    
    def browse_content_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path and file_path not in self.content_files:
            self.content_files.append(file_path)
            self.update_listbox(self.content_listbox, self.content_files)
        elif file_path in self.content_files:
            messagebox.showwarning("Warning", "Content file already added.")
    
    def update_listbox(self, listbox, file_list):
        listbox.delete(0, tk.END)
        for file_path in file_list:
            listbox.insert(tk.END, os.path.basename(file_path))
    
    def remove_file(self, file_type):
        if file_type == 'contact':
            selection = self.contact_listbox.curselection()
            if selection:
                index = selection[0]
                self.contact_files.pop(index)
                self.update_listbox(self.contact_listbox, self.contact_files)
            else:
                messagebox.showwarning("Warning", "No contact file selected.")
        elif file_type == 'content':
            selection = self.content_listbox.curselection()
            if selection:
                index = selection[0]
                self.content_files.pop(index)
                self.update_listbox(self.content_listbox, self.content_files)
            else:
                messagebox.showwarning("Warning", "No content file selected.")
    
    def start_process(self):
        if not self.contact_files or not self.content_files:
            messagebox.showerror("Error", "Please select both contact and content files.")
            return
        if not self.email_entry.get():
            messagebox.showerror("Error", "Please enter a sender email.")
            return
        
        self.sending_emails = True
        self.status_var.set("Processing... Please wait.")
        self.start_button.config(state='disabled')
        self.log_text.delete(1.0, tk.END)  # Clear previous logs
        self.progress_bar.start(5)  # Start the progress bar

        # Start the email sending process in a separate thread
        threading.Thread(target=self.send_emails_in_background).start()
        
        # Start checking the progress queue
        self.master.after(100, self.check_progress)
    
    def check_progress(self):
        try:
            message = self.progress_queue.get_nowait()
            self.status_var.set(message)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
            if message == "Process completed! Please check log files.":
                self.sending_emails = False
                self.start_button.config(state='normal')
                self.progress_bar.stop()  # Stop the progress bar
                messagebox.showinfo("Success", "Email sending process completed!")
                self.reset_gui()
            else:
                self.master.after(100, self.check_progress)
        except queue.Empty:
            if self.sending_emails:
                self.master.after(100, self.check_progress)
    
    def reset_gui(self):
        self.contact_files.clear()
        self.content_files.clear()
        self.update_listbox(self.contact_listbox, self.contact_files)
        self.update_listbox(self.content_listbox, self.content_files)
        self.email_entry.delete(0, tk.END)
        self.delay_entry.delete(0, tk.END)
        self.delay_entry.insert(0, "5")

        self.status_var.set("File selection refreshed.")
        self.log_text.delete(1.0, tk.END)

    def send_emails_in_background(self):
        try:
            delay_seconds = max(4, int(self.delay_entry.get()))
        except ValueError:
            delay_seconds = 5

        campaign_files = {}
        for file_path in self.contact_files + self.content_files:
            campaign_name = os.path.basename(file_path).split('_')[0]
            file_type = 'contact' if '_contact.csv' in file_path else 'content'
            campaign_files.setdefault(campaign_name, {})[file_type] = file_path

        try:
            for campaign_name, files in campaign_files.items():
                if 'contact' not in files or 'content' not in files:
                    raise Exception(f"Campaign '{campaign_name}' is missing either the contact or content file.")

                final_logs = []
                contact_file = files['contact']
                content_file = files['content']

                # Get directory of contact file to save log file
                contact_file_directory = os.path.dirname(contact_file)
                log_file_name = os.path.join(contact_file_directory, f"{campaign_name}_logs.csv")

                contacts_df = pd.read_csv(contact_file, encoding='latin1')
                with open(content_file, 'r', encoding='ISO-8859-1') as f:
                    content_df = pd.read_csv(f)
                content_df_iter = iter(content_df.iterrows())

                sender_email = self.email_entry.get()
                sender_name = sender_email.split('@')[0].capitalize()

                for index, contact in contacts_df.iterrows():
                    content_row = next(content_df_iter, None)
                    if content_row is None:
                        content_df_iter = iter(content_df.iterrows())
                        content_row = next(content_df_iter)

                    subject = content_row[1]['Subject']
                    body = content_row[1].get('Body', '')
                    recipient = contact['Email']

                    for column, value in contact.items():
                        placeholder = '{' + column.lower() + '}'
                        body = body.replace(placeholder, str(value))

                    try:
                        response = self.send_email(sender_name, sender_email, recipient, subject, body)
                        message_id = response.get('MessageId', 'N/A')

                        final_logs.append({
                            "Email": recipient,
                            "Status": "Success",
                            "MessageId": message_id
                        })

                    except ClientError as e:
                        final_logs.append({
                            "Email": recipient,
                            "Status": f"Error: {e.response['Error']['Message']}"
                        })

                    self.progress_queue.put(f"Sent to {recipient}. Waiting {delay_seconds} seconds...")
                    time.sleep(delay_seconds)

                logs_df = pd.DataFrame(final_logs)
                logs_df.to_csv(log_file_name, index=False)

            self.progress_queue.put("Process completed! Please check log files.")
        except Exception as e:
            self.progress_queue.put(f"Error: {str(e)}")

    def send_email(self, sender_name, sender_email, recipient_email, subject, body):
        charset = "UTF-8"
        response = self.ses_client.send_email(
            Destination={
                'ToAddresses': [recipient_email],
            },
            Message={
                'Body': {
                    'Html': {
                        'Charset': charset,
                        'Data': body,
                    },
                },
                'Subject': {
                    'Charset': charset,
                    'Data': subject,
                },
            },
            Source=f"{sender_name} | Your Company Name <{sender_email}>",
        )
        return response

if __name__ == "__main__":
    root = tk.Tk()
    gui = ImprovedEmailClientGUI(root)
    root.mainloop()