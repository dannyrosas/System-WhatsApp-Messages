import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import time
import os

from src.utils import create_template_file, load_excel_data, format_time
from src.sender import send_whatsapp_message

class WhatsAppSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Message Sender")
        self.root.geometry("900x800")
        self.root.configure(padx=20, pady=20)
        
        self.file_path = tk.StringVar()
        self.status_var = tk.StringVar(value="Waiting for file...")
        
        main_frame = ttk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        file_frame = ttk.LabelFrame(main_frame, text="Message File")
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(file_frame, text="Excel File:").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(file_frame, textvariable=self.file_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_file).grid(row=0, column=2, padx=5, pady=5)
        ttk.Button(file_frame, text="Create Template", command=self.create_template).grid(row=1, column=0, padx=5, pady=5)
        ttk.Button(file_frame, text="Load Data", command=self.load_data).grid(row=1, column=1, padx=5, pady=5)
        
        preview_frame = ttk.LabelFrame(main_frame, text="Data Preview")
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tree = ttk.Treeview(preview_frame)
        self.tree["columns"] = ("Name", "Indicative", "Number", "Message")
        self.tree.column("#0", width=0, stretch=tk.NO)
        self.tree.column("Name", anchor=tk.W, width=100)
        self.tree.column("Indicative", anchor=tk.W, width=70)
        self.tree.column("Number", anchor=tk.W, width=120)
        self.tree.column("Message", anchor=tk.W, width=250)
        
        self.tree.heading("#0", text="", anchor=tk.W)
        self.tree.heading("Name", text="Name", anchor=tk.W)
        self.tree.heading("Indicative", text="Indicative", anchor=tk.W)
        self.tree.heading("Number", text="Number", anchor=tk.W)
        self.tree.heading("Message", text="Message", anchor=tk.W)
        
        scrolly = ttk.Scrollbar(preview_frame, orient="vertical", command=self.tree.yview)
        scrollx = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
        
        scrolly.pack(side="right", fill="y")
        scrollx.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)
        
        send_frame = ttk.LabelFrame(main_frame, text="Send Messages")
        send_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(send_frame, text="Send Messages", command=self.send_messages).grid(row=0, column=0, padx=20, pady=5)
        
        self.progress = ttk.Progressbar(send_frame, orient="horizontal", length=300, mode="determinate")
        self.progress.grid(row=1, column=0, columnspan=1, padx=5, pady=5, sticky="ew")
        
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(status_frame, text="Status:").pack(side=tk.LEFT, padx=5)
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT, padx=5)
        
        self.result_text = tk.Text(main_frame, height=6, width=80)
        self.result_text.pack(fill=tk.X, padx=10, pady=10)
        
        self.df = None
    
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.file_path.set(file_path)
            self.status_var.set(f"File selected: {os.path.basename(file_path)}")
    
    def create_template(self):
        file_path = filedialog.asksaveasfilename(
            title="Save Template",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if file_path:
            try:
                create_template_file(file_path)
                self.file_path.set(file_path)
                self.status_var.set(f"Template created: {os.path.basename(file_path)}")
                messagebox.showinfo("Success", f"Template successfully created at {file_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Error creating template: {str(e)}")
    
    def load_data(self):
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showwarning("Warning", "Please select an Excel file first.")
            return
        
        try:
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            self.df = load_excel_data(file_path)
            
            for i, row in self.df.iterrows():
                self.tree.insert("", "end", iid=i, values=(
                    row["Name"],
                    row["Indicative"],
                    row["Number"],
                    row["Message"]
                ))
            
            self.status_var.set(f"Data loaded: {len(self.df)} messages")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error loading data: {str(e)}")
    
    def send_messages(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("Warning", "No data to send. Please load an Excel file.")
            return
        
        thread = threading.Thread(target=self._send_messages_thread)
        thread.daemon = True
        thread.start()
    
    def _send_messages_thread(self):
        try:
            wait_time = 15
            total_rows = len(self.df)
            self.progress["maximum"] = total_rows
            self.progress["value"] = 0
            
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, "Starting message sending process...\n")
            
            success_count = 0
            error_count = 0
            start_time = time.time()
            
            for index, row in self.df.iterrows():
                try:
                    name = row["Name"]
                    indicative = str(row["Indicative"])
                    number = str(row["Number"])
                    message = row["Message"]
                    
                    phone_no = f"+{indicative}{number}"
                    
                    self.status_var.set(f"Sending message to {name} ({phone_no})...")
                    self.result_text.insert(tk.END, f"⏳ Sending to {name} ({phone_no})...\n")
                    self.result_text.see(tk.END)
                    
                    send_whatsapp_message(phone_no, message, wait_time)
                    
                    success_count += 1
                    self.result_text.insert(tk.END, f"✅ Message sent to {name} ({phone_no})\n")
                    self.result_text.see(tk.END)
                    
                except Exception as e:
                    error_count += 1
                    self.result_text.insert(tk.END, f"⚠️ Error sending to {name} ({phone_no}): {str(e)}\n")
                    self.result_text.see(tk.END)
                
                self.progress["value"] = index + 1
                self.root.update_idletasks()
            
            end_time = time.time()
            elapsed_time = end_time - start_time
            
            hours, minutes, seconds = format_time(elapsed_time)
            
            summary = f"\n📊 **Summary Report** 📊\n"
            summary += f"✅ Successfully sent messages: {success_count}\n"
            summary += f"⚠️ Failed messages: {error_count}\n"
            summary += f"⏳ Total execution time: {hours}h {minutes}m {seconds}s"
            
            self.result_text.insert(tk.END, summary)
            self.result_text.see(tk.END)
            self.status_var.set(f"Completed: {success_count} sent, {error_count} failed")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during sending process: {str(e)}")
            self.status_var.set("Error during sending process")