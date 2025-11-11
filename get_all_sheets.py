import sys
import subprocess
import importlib

# --- Dependency Check and Auto-Install ---
required_packages = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "tkinter": "tk",  # tkinter is included in Python/Anaconda, auto-install not needed
    "customtkinter": "customtkinter"
}

missing_packages = []

for package, install_name in required_packages.items():
    try:
        importlib.import_module(package)
    except ImportError:
        if package == "tkinter":
            missing_packages.append(package)
        else:
            try:
                print(f"Installing missing package: {install_name}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", install_name])
            except Exception as e:
                missing_packages.append(package)

if missing_packages:
    import tkinter
    from tkinter import messagebox
    msg = "The following required packages could not be installed:\n" + "\n".join(missing_packages)
    msg += "\n\nPlease install them manually and rerun the application."
    root = tkinter.Tk()
    root.withdraw()
    messagebox.showerror("Missing Dependencies", msg)
    sys.exit(1)

# --- All dependencies are available ---
import os
import threading
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

class ExcelCombinerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Excel Sheet Combiner")
        self.geometry("600x400")
        self.resizable(False, False)

        self.input_dir = None

        # --- UI Elements ---
        self.label_title = ctk.CTkLabel(self, text="Excel Sheet Combiner", font=("Segoe UI", 20, "bold"))
        self.label_title.pack(pady=10)

        self.btn_select = ctk.CTkButton(self, text="Select Folder", command=self.select_folder)
        self.btn_select.pack(pady=10)

        self.label_path = ctk.CTkLabel(self, text="No folder selected", wraplength=500, text_color="gray")
        self.label_path.pack(pady=5)

        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.set(0)
        self.progress_bar.pack(pady=10)

        self.text_output = ctk.CTkTextbox(self, width=500, height=150)
        self.text_output.pack(pady=10)

        self.btn_start = ctk.CTkButton(self, text="Start Combining", state="disabled", command=self.start_process)
        self.btn_start.pack(pady=10)

    # --- Folder Selection ---
    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing Excel Files")
        if folder:
            self.input_dir = folder
            self.label_path.configure(text=folder)
            self.btn_start.configure(state="normal")

    # --- Start Threaded Process ---
    def start_process(self):
        self.btn_start.configure(state="disabled")
        self.text_output.delete("1.0", "end")
        thread = threading.Thread(target=self.combine_excel_sheets)
        thread.start()

    # --- Excel Combining Logic ---
    def combine_excel_sheets(self):
        input_dir = self.input_dir
        output_dir = os.path.join(input_dir, "output")
        os.makedirs(output_dir, exist_ok=True)
        output_file = os.path.join(output_dir, "combined_workbook.xlsx")

        excel_files = [f for f in os.listdir(input_dir) if f.endswith(('.xlsx', '.xls'))]
        total_sheets = 0
        processed_sheets = 0
        errors = []

        for file_name in excel_files:
            try:
                total_sheets += len(pd.ExcelFile(os.path.join(input_dir, file_name)).sheet_names)
            except Exception as e:
                errors.append(f"Failed to read {file_name}: {e}")

        if total_sheets == 0:
            messagebox.showinfo("No Files", "No valid Excel files found.")
            self.btn_start.configure(state="normal")
            return

        writer = pd.ExcelWriter(output_file, engine='openpyxl')

        for file_name in excel_files:
            try:
                file_path = os.path.join(input_dir, file_name)
                excel_file = pd.ExcelFile(file_path)
                for sheet_name in excel_file.sheet_names:
                    try:
                        df = excel_file.parse(sheet_name)
                        safe_name = f"{os.path.splitext(file_name)[0]}_{sheet_name}"[:31]
                        df.to_excel(writer, sheet_name=safe_name, index=False)
                    except Exception as e:
                        errors.append(f"{file_name} → {sheet_name}: {e}")
                    processed_sheets += 1
                    self.progress_bar.set(processed_sheets / total_sheets)
            except Exception as e:
                errors.append(f"{file_name} skipped: {e}")

        writer.close()
        self.show_summary(output_file, processed_sheets, errors)

    # --- Display Summary ---
    def show_summary(self, output_file, processed, errors):
        self.text_output.insert("end", f"✅ Processing Complete!\n\n")
        self.text_output.insert("end", f"Sheets Combined: {processed}\n")
        self.text_output.insert("end", f"Output File:\n{output_file}\n\n")
        if errors:
            self.text_output.insert("end", "⚠️ Errors:\n")
            for e in errors:
                self.text_output.insert("end", f"- {e}\n")
        else:
            self.text_output.insert("end", "✅ No errors encountered.\n")

        self.btn_start.configure(state="normal")


if __name__ == "__main__":
    app = ExcelCombinerApp()
    app.mainloop()
