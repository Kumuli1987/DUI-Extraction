import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
from pypdf import PdfReader
import os
import re
import threading

# Color theme
BG_COLOR = "#e6f2ff"
HEADER_BG = "#4682b4"
HEADER_FG = "#ffffff"
BUTTON_BG = "#4caf50"
BUTTON_FG = "#ffffff"
BUTTON_ACTIVE_BG = "#357a38"
LOG_BG = "#f7fbff"
LOG_FG = "#222222"
LABEL_FG = "#333333"

class DOIExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF DOI Extractor")
        self.root.configure(bg=BG_COLOR)
        self.root.geometry("900x600")
        self.root.minsize(900, 600)
        self.input_folder = ""
        self.output_folder = ""
        self.progress_style_name = "Blue.Horizontal.TProgressbar"
        self.green_progress_style_name = "Green.Horizontal.TProgressbar"
        
        # Header
        header = tk.Label(
            root, text="PDF DOI Extractor",
            font=("Arial Rounded MT Bold", 28, "bold"),
            bg=HEADER_BG, fg=HEADER_FG, pady=20, padx=10
        )
        header.pack(fill=tk.X, pady=(0, 25))

        frm1 = tk.Frame(root, bg=BG_COLOR)
        frm1.pack(pady=15, padx=40, fill=tk.X)
        btn1 = tk.Button(
            frm1, text="Select Input Folder", command=self.select_input_folder,
            bg=BUTTON_BG, fg=BUTTON_FG, activebackground=BUTTON_ACTIVE_BG, relief="groove",
            font=("Arial", 16, "bold"), padx=18, pady=10, bd=0, width=22
        )
        btn1.pack(side=tk.LEFT)
        self.input_label = tk.Label(
            frm1, text="No folder selected", width=60, anchor="w", bg=BG_COLOR, fg=LABEL_FG,
            font=("Arial", 14)
        )
        self.input_label.pack(side=tk.LEFT, padx=20)

        frm2 = tk.Frame(root, bg=BG_COLOR)
        frm2.pack(pady=15, padx=40, fill=tk.X)
        btn2 = tk.Button(
            frm2, text="Select Output Folder", command=self.select_output_folder,
            bg=BUTTON_BG, fg=BUTTON_FG, activebackground=BUTTON_ACTIVE_BG, relief="groove",
            font=("Arial", 16, "bold"), padx=18, pady=10, bd=0, width=22
        )
        btn2.pack(side=tk.LEFT)
        self.output_label = tk.Label(
            frm2, text="No folder selected", width=60, anchor="w", bg=BG_COLOR, fg=LABEL_FG,
            font=("Arial", 14)
        )
        self.output_label.pack(side=tk.LEFT, padx=20)

        run_btn = tk.Button(
            root, text="Run Extraction", command=self.run_extraction,
            bg="#1976d2", fg=BUTTON_FG, activebackground="#0d47a1", relief="groove",
            font=("Arial Rounded MT Bold", 18), padx=35, pady=14, bd=0, width=20
        )
        run_btn.pack(pady=30)

        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.style = ttk.Style()
        self.style.theme_use('default')
        self.style.configure(self.progress_style_name, thickness=32, troughcolor='#d0e2ff', background='#1976d2')
        self.style.configure(self.green_progress_style_name, thickness=32, troughcolor='#d0e2ff', background='#4caf50')
        self.progress = ttk.Progressbar(
            root, orient="horizontal", length=800, mode="determinate",
            variable=self.progress_var, style=self.progress_style_name
        )
        self.progress.pack(pady=(0, 20), padx=30)

        self.log_text = tk.Text(
            root, height=10, width=110, state=tk.DISABLED,
            bg=LOG_BG, fg=LOG_FG, font=("Consolas", 13), wrap=tk.WORD
        )
        self.log_text.pack(pady=(10, 18), padx=30, fill=tk.BOTH, expand=True)
        self.log_text.config(highlightbackground="#b0c4de", highlightcolor="#b0c4de", borderwidth=2)
    
    def select_input_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.input_folder = folder
            self.input_label.config(text=folder)
    
    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder
            self.output_label.config(text=folder)
    
    def log(self, msg):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, msg + '\n')
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
    
    def set_progress(self, value, total):
        if total > 0:
            percent = (value / total) * 100
        else:
            percent = 0
        self.progress_var.set(percent)
        self.root.update_idletasks()
    
    def set_progress_green(self):
        """Set the progress bar to green and full."""
        self.progress.config(style=self.green_progress_style_name)
        self.progress_var.set(100)
        self.root.update_idletasks()

    def reset_progress_bar(self):
        """Reset progress bar to blue and empty."""
        self.progress.config(style=self.progress_style_name)
        self.progress_var.set(0)
        self.root.update_idletasks()

    def on_finish(self, success, output_path):
        if success:
            self.set_progress_green()
            messagebox.showinfo("Done", f"Extraction complete.\nSaved to:\n{output_path}")
        else:
            self.reset_progress_bar()
            messagebox.showwarning("Warning", "No PDFs found or processed.")

    def run_extraction(self):
        if not self.input_folder or not self.output_folder:
            messagebox.showwarning("Missing Folder", "Please select both input and output folders.")
            return
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete('1.0', tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.reset_progress_bar()
        thread = threading.Thread(target=self.extract_doi_thread)
        thread.start()

    def extract_doi_thread(self):
        extract_doi(
            self.input_folder, self.output_folder, self.log, self.set_progress, self.on_finish
        )

def extract_doi(input_folder, output_folder, log_callback, progress_callback, finish_callback):
    DO_pattern = r'https?://doi\.org/(.+)'
    DO_list = []
    pdf_files = [f for f in os.listdir(input_folder) if f.lower().endswith('.pdf')]
    total = len(pdf_files)
    log_callback("Starting extraction...")
    for idx, filename in enumerate(pdf_files):
        pdf_path = os.path.join(input_folder, filename)
        log_callback(f'Processing: {filename}')
        try:
            reader = PdfReader(pdf_path)
            found = False
            for page_num, page in enumerate(reader.pages):
                text = page.extract_text() or ""
                matches = re.findall(DO_pattern, text)
                DOI_Match = re.search(DO_pattern, text)
                DOI_group = DOI_Match.group(0) if DOI_Match else ""
                if matches:
                    DO_list.append({
                        "File Name": filename,
                        "Page": page_num + 1,
                        "DOI Number": matches[0],
                        "DOI_Group": DOI_group
                    })
                    found = True
                    break
            reader.close()
            if not found:
                DO_list.append({
                    "File Name": filename,
                    "Page": "N/A",
                    "DOI Number": "Not Found"
                })
        except Exception as e:
            log_callback(f"Error with {filename}: {e}")
        progress_callback(idx + 1, total)
    if DO_list:
        df = pd.DataFrame(DO_list)
        output_path = os.path.join(output_folder, "Extracted_DO.xlsx")
        df.to_excel(output_path, index=False)
        log_callback(f"Extraction complete. Saved to {output_path}")
        finish_callback(True, output_path)
    else:
        log_callback("No PDFs found or processed.")
        finish_callback(False, "")

if __name__ == "__main__":
    root = tk.Tk()
    app = DOIExtractorApp(root)
    root.mainloop()