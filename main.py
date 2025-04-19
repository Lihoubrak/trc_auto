import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from logging_config import configure_logging
from driver_utils import terminate_chrome_processes, initialize_driver
from excel_utils import read_excel_data
from form_utils import get_form_headers, fill_google_form
from matching_utils import match_headers
from config import GOOGLE_FORM_URL, EXCEL_FILE, USER_DATA_DIR, PROFILE_DIR
from pathlib import Path
import openpyxl
import sys
import config
import threading
import importlib
import json
import os

CONFIG_PY = "config.py"

class ConfigGUI:
    """GUI for overriding configuration settings in config.py."""
    def __init__(self, root):
        self.root = root
        self.root.title("Configuration Settings")
        self.root.geometry("600x400")  # Increased height for account info labels
        self.root.resizable(False, False)  # Disable resizing and maximize button
        self.entries = {}
        self.is_running = False  # Flag to prevent multiple runs
        self.load_config()
        self.create_widgets()

    def load_config(self):
        """Load current configuration from config.py and account info from Local State."""
        self.config_values = {
            "GOOGLE_FORM_URL": GOOGLE_FORM_URL,
            "EXCEL_FILE": str(Path(EXCEL_FILE)) if EXCEL_FILE else "C:\\Users\\KHC\\Desktop\\trc_auto\\test.xlsx",
            "USER_DATA_DIR": str(Path(USER_DATA_DIR)) if USER_DATA_DIR else "C:\\Users\\KHC\\AppData\\Local\\Google\\Chrome\\User Data",
            "PROFILE_DIR": PROFILE_DIR if PROFILE_DIR else "Profile 1"
        }
        logging.info(f"Loaded config: {self.config_values}")
        self.account_name, self.email = self.get_account_info()

    def get_account_info(self):
        """Read account name and email from Local State file."""
        try:
            local_state_path = Path(self.config_values["USER_DATA_DIR"]) / "Local State"
            if not local_state_path.is_file():
                logging.warning(f"Local State file not found at {local_state_path}")
                return "Not available", "Not available"

            with open(local_state_path, 'r', encoding='utf-8') as f:
                local_state = json.load(f)

            profile_cache = local_state.get("profile", {}).get("info_cache", {})
            profile_info = profile_cache.get(self.config_values["PROFILE_DIR"], {})

            account_name = profile_info.get("name", "Not available")
            email = profile_info.get("user_name", "Not available")

            if account_name == "Not available" and email == "Not available":
                logging.info(f"No account info found for profile {self.config_values['PROFILE_DIR']}")
            else:
                logging.info(f"Retrieved account info - Name: {account_name}, Email: {email}")

            return account_name, email
        except Exception as e:
            logging.error(f"Failed to read Local State: {e}")
            return "Not available", "Not available"

    def save_config(self):
        """Save configuration by rewriting config.py."""
        try:
            google_form_url = self.entries["GOOGLE_FORM_URL"].get()
            excel_file = str(Path(self.entries["EXCEL_FILE"].get())).replace('\\', '\\\\')  # Escape for Python string
            user_data_dir = str(Path(self.entries["USER_DATA_DIR"].get())).replace('\\', '\\\\')  # Escape for Python string
            profile_dir = self.entries["PROFILE_DIR"].get()

            # Template for config.py
            config_content = f"""# Constants
GOOGLE_FORM_URL = "{google_form_url}"
EXCEL_FILE = r"{excel_file}"
CHROMEDRIVER_PATH = "chromedriver.exe"
SIMILARITY_THRESHOLD = 80
USER_DATA_DIR = r"{user_data_dir}"
PROFILE_DIR = "{profile_dir}"
"""
            with open(CONFIG_PY, "w") as f:
                f.write(config_content)
            logging.info("Configuration saved to config.py")
        except Exception as e:
            logging.error(f"Failed to save config.py: {e}")
            raise

    def clear_config(self):
        """Reset config.py to default values."""
        if messagebox.askyesno("Confirm", "Are you sure you want to reset settings to defaults?"):
            try:
                config_content = """# Constants
GOOGLE_FORM_URL = ""
EXCEL_FILE = ""
CHROMEDRIVER_PATH = "chromedriver.exe"
SIMILARITY_THRESHOLD = 80
USER_DATA_DIR = r""
PROFILE_DIR = ""
"""
                with open(CONFIG_PY, "w") as f:
                    f.write(config_content)
                logging.info("Configuration reset to defaults in config.py")
                
                # Update GUI fields
                self.config_values = {
                    "GOOGLE_FORM_URL": "",
                    "EXCEL_FILE": "",
                    "USER_DATA_DIR": "",
                    "PROFILE_DIR": ""
                }
                for key, entry in self.entries.items():
                    entry.delete(0, tk.END)
                    entry.insert(0, self.config_values[key])
                
                # Update account info
                self.account_name, self.email = "Not available", "Not available"
                self.account_name_label.config(text=f"Account Name: {self.account_name}")
                self.email_label.config(text=f"Email: {self.email}")
                
                messagebox.showinfo("Success", "Settings reset to defaults")
            except Exception as e:
                logging.error(f"Failed to reset config.py: {e}")
                messagebox.showerror("Error", f"Failed to reset config.py: {e}")

    def create_widgets(self):
        """Create GUI widgets for configuration inputs and account info."""
        frame = ttk.Frame(self.root, padding="10")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        config_fields = [
            ("Google Form URL", "GOOGLE_FORM_URL", False),
            ("Excel File", "EXCEL_FILE", True),
            ("User Data Directory", "USER_DATA_DIR", True),
            ("Profile Directory", "PROFILE_DIR", False),
        ]

        for idx, (label_text, config_key, is_file) in enumerate(config_fields):
            ttk.Label(frame, text=label_text + ":").grid(row=idx, column=0, sticky=tk.W, pady=5)
            entry = ttk.Entry(frame, width=50)
            entry.insert(0, self.config_values[config_key])
            entry.grid(row=idx, column=1, sticky=(tk.W, tk.E), pady=5)
            self.entries[config_key] = entry
            if is_file:
                browse_btn = ttk.Button(frame, text="Browse", command=lambda k=config_key: self.browse_file(k))
                browse_btn.grid(row=idx, column=2, padx=5)

        # Add account info labels
        self.account_name_label = ttk.Label(frame, text=f"Account Name: {self.account_name}")
        self.account_name_label.grid(row=len(config_fields), column=0, columnspan=2, sticky=tk.W, pady=5)
        
        self.email_label = ttk.Label(frame, text=f"Email: {self.email}")
        self.email_label.grid(row=len(config_fields)+1, column=0, columnspan=2, sticky=tk.W, pady=5)

        self.save_run_btn = ttk.Button(frame, text="Save and Run", command=self.save_and_run)
        self.save_run_btn.grid(row=len(config_fields)+2, column=0, columnspan=3, pady=10)
        ttk.Button(frame, text="Clear Saved Settings", command=self.clear_config).grid(row=len(config_fields)+3, column=0, columnspan=3, pady=5)

        # Add status label
        self.status_label = ttk.Label(frame, text="Ready", foreground="black")
        self.status_label.grid(row=len(config_fields)+4, column=0, columnspan=3, pady=10)

    def browse_file(self, config_key):
        """Open file/directory dialog for specific configuration fields and update account info."""
        if config_key == "EXCEL_FILE":
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        elif config_key == "USER_DATA_DIR":
            file_path = filedialog.askdirectory()
        else:
            return
        if file_path:
            self.entries[config_key].delete(0, tk.END)
            self.entries[config_key].insert(0, str(Path(file_path)))  # Normalize path
            logging.info(f"Selected {config_key}: {file_path}")
            
            # Update config_values and account info if User Data Directory or Profile Directory changed
            self.config_values[config_key] = str(Path(file_path))
            if config_key in ["USER_DATA_DIR", "PROFILE_DIR"]:
                self.account_name, self.email = self.get_account_info()
                self.account_name_label.config(text=f"Account Name: {self.account_name}")
                self.email_label.config(text=f"Email: {self.email}")

    def prevent_close(self):
        """Prevent closing the window during automation."""
        messagebox.showwarning("Warning", "Please wait for the Excel file to finish updating before closing the application.")
        return

    def save_and_run(self):
        """Validate inputs, save to config.py, update config module, and run automation in a separate thread."""
        if self.is_running:
            messagebox.showwarning("Warning", "Automation is already running. Please wait for it to complete.")
            return

        try:
            config.GOOGLE_FORM_URL = self.entries["GOOGLE_FORM_URL"].get()
            config.EXCEL_FILE = str(Path(self.entries["EXCEL_FILE"].get()))  # Normalize path
            config.USER_DATA_DIR = str(Path(self.entries["USER_DATA_DIR"].get()))  # Normalize path
            config.PROFILE_DIR = self.entries["PROFILE_DIR"].get()

            if not config.GOOGLE_FORM_URL.startswith("http"):
                raise ValueError("Invalid Google Form URL")
            if not Path(config.EXCEL_FILE).is_file():
                raise ValueError("Excel file does not exist")
            if not Path(config.USER_DATA_DIR).is_dir():
                raise ValueError("User data directory does not exist")

            self.save_config()
            # Reload config module to reflect new config.py values
            importlib.reload(config)

            # Update account info based on new config
            self.config_values["USER_DATA_DIR"] = config.USER_DATA_DIR
            self.config_values["PROFILE_DIR"] = config.PROFILE_DIR
            self.account_name, self.email = self.get_account_info()
            self.account_name_label.config(text=f"Account Name: {self.account_name}")
            self.email_label.config(text=f"Email: {self.email}")

            # Disable the "Save and Run" button and update status
            self.is_running = True
            self.save_run_btn.config(state="disabled")
            self.status_label.config(text="Running, please wait for Excel update...", foreground="red")
            
            # Prevent closing the window during automation
            self.root.protocol("WM_DELETE_WINDOW", self.prevent_close)

            # Run main() in a separate thread
            threading.Thread(target=self.run_automation, daemon=True).start()

        except ValueError as e:
            messagebox.showerror("Error", str(e))
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {e}")

    def run_automation(self):
        """Run the main automation process and display the result."""
        try:
            result = main()
            self.root.after(0, lambda: self.show_result(result))
        except Exception as e:
            self.root.after(0, lambda: self.show_result(f"Unexpected error: {e}"))
        finally:
            self.root.after(0, self.reset_gui)

    def show_result(self, result):
        """Display the result of the automation process."""
        google_form_name = config.GOOGLE_FORM_URL.split('/')[-2] if '/' in config.GOOGLE_FORM_URL else "Google Form"
        
        if result == "Success":
            message = f"Automation completed successfully! Excel file has been updated."
            messagebox.showinfo("Success", message)
            self.status_label.config(text="Completed: Excel file updated", foreground="green")
        else:
            message = f"{result}\nExcel file has been updated with error notes."
            messagebox.showerror("Error", message)
            self.status_label.config(text="Error: Check Excel file for details", foreground="red")

    def reset_gui(self):
        """Re-enable the GUI after automation completes."""
        self.is_running = False
        self.save_run_btn.config(state="normal")
        # Re-enable closing the window
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        # If no messagebox was shown, ensure status is updated
        if self.status_label.cget("text").startswith("Running"):
            self.status_label.config(text="Ready", foreground="black")

def main():
    """Main function to orchestrate the automation process."""
    configure_logging()
    driver = None
    try:
        filepath = Path(config.EXCEL_FILE)
        wb = openpyxl.load_workbook(filepath)
        sheet = wb.active
        
        excel_headers, rows = read_excel_data(config.EXCEL_FILE)
        logging.info(f"Total rows to process: {len(rows)}")

        note_column = None
        for col_idx, cell in enumerate(sheet[1], start=1):
            if cell.value and cell.value.lower() == "note":
                note_column = col_idx
                break
        if not note_column:
            note_column = len(excel_headers) + 1
            sheet.cell(row=1, column=note_column).value = "Note"
            logging.info(f"Added 'Note' column to Excel file at column {note_column}")

        if not rows:
            logging.info("No data rows to process in Excel file")
            wb.save(filepath)
            return "Success"

        terminate_chrome_processes()
        driver = initialize_driver()
        form_headers = get_form_headers(driver)
        header_mapping, unmatched_headers = match_headers(excel_headers, form_headers)

        for idx, row in enumerate(rows, start=2):
            note_cell = sheet.cell(row=idx, column=note_column).value
            if note_cell == "Inserted":
                logging.info(f"Row {idx} already inserted, skipping")
                continue

            logging.info(f"Processing row {idx}: {row}")
            success = fill_google_form(driver, row, excel_headers, header_mapping)

            if success:
                sheet.cell(row=idx, column=note_column).value = "Inserted"
                logging.info(f"Row {idx} processed successfully")
            else:
                error_message = f"Failed to insert row {idx}"
                sheet.cell(row=idx, column=note_column).value = error_message
                logging.error(f"{error_message} - Row data: {row}")
                
                wb.save(filepath)
                logging.info(f"Excel file saved with error note for row {idx}")
                
                return error_message

            wb.save(filepath)
            logging.info(f"Excel file saved after processing row {idx}")

        wb.save(filepath)  # Final save to ensure all updates are written
        logging.info("Final Excel file save completed")
        return "Success"

    except Exception as e:
        logging.error(f"Main process error: {e}")
        wb.save(filepath)
        logging.info("Excel file saved due to critical error")
        return f"Main process error: {e}"

    finally:
        if driver:
            try:
                driver.quit()
                logging.info("WebDriver closed")
            except Exception as e:
                logging.error(f"Error closing WebDriver: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConfigGUI(root)
    root.mainloop()