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
        self.root.geometry("600x400")
        self.root.resizable(False, False)
        self.entries = {}
        self.is_running = False
        self.load_config()
        self.create_widgets()

    def load_config(self):
        """Load current configuration from config.py and account info from Local State."""
        self.config_values = {
            "GOOGLE_FORM_URL": GOOGLE_FORM_URL or "",
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

            logging.info(f"Retrieved account info - Name: {account_name}, Email: {email}")
            return account_name, email
        except Exception as e:
            logging.error(f"Failed to read Local State: {e}")
            return "Not available", "Not available"

    def save_config(self):
        """Save configuration by rewriting config.py."""
        try:
            google_form_url = self.entries["GOOGLE_FORM_URL"].get().strip()
            excel_file = str(Path(self.entries["EXCEL_FILE"].get())).replace('\\', '\\\\')
            user_data_dir = str(Path(self.entries["USER_DATA_DIR"].get())).replace('\\', '\\\\')
            profile_dir = self.entries["PROFILE_DIR"].get().strip()

            config_content = f"""# Constants
GOOGLE_FORM_URL = "{google_form_url}"
EXCEL_FILE = r"{excel_file}"
CHROMEDRIVER_PATH = "chromedriver.exe"
SIMILARITY_THRESHOLD = 80
USER_DATA_DIR = r"{user_data_dir}"
PROFILE_DIR = "{profile_dir}"
"""
            with open(CONFIG_PY, "w", encoding='utf-8') as f:
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
                with open(CONFIG_PY, "w", encoding='utf-8') as f:
                    f.write(config_content)
                logging.info("Configuration reset to defaults in config.py")

                self.config_values = {
                    "GOOGLE_FORM_URL": "",
                    "EXCEL_FILE": "",
                    "USER_DATA_DIR": "",
                    "PROFILE_DIR": ""
                }
                for key, entry in self.entries.items():
                    entry.delete(0, tk.END)
                    entry.insert(0, self.config_values[key])

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

        self.account_name_label = ttk.Label(frame, text=f"Account Name: {self.account_name}")
        self.account_name_label.grid(row=len(config_fields), column=0, columnspan=2, sticky=tk.W, pady=5)

        self.email_label = ttk.Label(frame, text=f"Email: {self.email}")
        self.email_label.grid(row=len(config_fields)+1, column=0, columnspan=2, sticky=tk.W, pady=5)

        self.save_run_btn = ttk.Button(frame, text="Save and Run", command=self.save_and_run)
        self.save_run_btn.grid(row=len(config_fields)+2, column=0, columnspan=3, pady=10)
        ttk.Button(frame, text="Clear Saved Settings", command=self.clear_config).grid(row=len(config_fields)+3, column=0, columnspan=3, pady=5)

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
            self.entries[config_key].insert(0, str(Path(file_path)))
            logging.info(f"Selected {config_key}: {file_path}")

            self.config_values[config_key] = str(Path(file_path))
            if config_key in ["USER_DATA_DIR", "PROFILE_DIR"]:
                self.account_name, self.email = self.get_account_info()
                self.account_name_label.config(text=f"Account Name: {self.account_name}")
                self.email_label.config(text=f"Email: {self.email}")

    def prevent_close(self):
        """Prevent closing the window during automation."""
        messagebox.showwarning("Warning", "Please wait for the automation to complete before closing the application.")
        return

    def save_and_run(self):
        """Validate inputs, save to config.py, update config module, and run automation in a separate thread."""
        if self.is_running:
            messagebox.showwarning("Warning", "Automation is already running. Please wait for it to complete.")
            return

        try:
            # Get and validate inputs
            google_form_url = self.entries["GOOGLE_FORM_URL"].get().strip()
            excel_file = self.entries["EXCEL_FILE"].get().strip()
            user_data_dir = self.entries["USER_DATA_DIR"].get().strip()
            profile_dir = self.entries["PROFILE_DIR"].get().strip()

            # Log the input values for debugging
            logging.info(f"Inputs - Google Form URL: {google_form_url}, Excel File: {excel_file}, User Data Dir: {user_data_dir}, Profile Dir: {profile_dir}")

            # Validate inputs
            if not google_form_url:
                raise ValueError("Google Form URL cannot be empty")
            if not google_form_url.startswith("http"):
                raise ValueError("Invalid Google Form URL: Must start with 'http'")
            if not excel_file:
                raise ValueError("Excel file path cannot be empty")
            excel_path = Path(excel_file)
            if not excel_path.is_file():
                raise ValueError(f"Excel file does not exist: {excel_file}")
            if not user_data_dir:
                raise ValueError("User data directory cannot be empty")
            if not Path(user_data_dir).is_dir():
                raise ValueError(f"User data directory does not exist: {user_data_dir}")
            if not profile_dir:
                raise ValueError("Profile directory cannot be empty")

            # Update config module
            config.GOOGLE_FORM_URL = google_form_url
            config.EXCEL_FILE = str(excel_path)
            config.USER_DATA_DIR = user_data_dir
            config.PROFILE_DIR = profile_dir

            self.save_config()
            importlib.reload(config)

            self.config_values["USER_DATA_DIR"] = config.USER_DATA_DIR
            self.config_values["PROFILE_DIR"] = config.PROFILE_DIR
            self.account_name, self.email = self.get_account_info()
            self.account_name_label.config(text=f"Account Name: {self.account_name}")
            self.email_label.config(text=f"Email: {self.email}")

            self.is_running = True
            self.save_run_btn.config(state="disabled")
            self.status_label.config(text="Running, please wait for Excel update...", foreground="red")
            self.root.protocol("WM_DELETE_WINDOW", self.prevent_close)

            threading.Thread(target=self.run_automation, daemon=True).start()

        except ValueError as e:
            logging.error(f"Validation error: {e}")
            messagebox.showerror("Error", str(e))
        except Exception as e:
            logging.error(f"Unexpected error in save_and_run: {e}")
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
            # Avoid misleading message about Excel file update
            message = f"{result}"
            messagebox.showerror("Error", message)
            self.status_label.config(text="Error: Check logs for details", foreground="red")

    def reset_gui(self):
        """Re-enable the GUI after automation completes."""
        self.is_running = False
        self.save_run_btn.config(state="normal")
        self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
        if self.status_label.cget("text").startswith("Running"):
            self.status_label.config(text="Ready", foreground="black")

def main():
    """Main function to orchestrate the automation process."""
    configure_logging()
    driver = None
    wb = None
    filepath = Path(config.EXCEL_FILE)
    try:
        # Validate file existence
        if not filepath:
            raise ValueError("Excel file path is empty")
        if not filepath.is_file():
            raise FileNotFoundError(f"Excel file not found: {filepath}")

        # Load the workbook
        try:
            wb = openpyxl.load_workbook(filepath)
        except Exception as e:
            logging.error(f"Failed to load Excel file: {e}")
            raise ValueError(f"Failed to load Excel file: {e}")

        sheet = wb.active
        excel_headers, rows = read_excel_data(config.EXCEL_FILE)
        logging.info(f"Total rows to process: {len(rows)}")

        # Find or add 'Note' column
        note_column = None
        for col_idx, cell in enumerate(sheet[1], start=1):
            if cell.value and isinstance(cell.value, str) and cell.value.lower() == "note":
                note_column = col_idx
                break
        if not note_column:
            note_column = len(excel_headers) + 1
            sheet.cell(row=1, column=note_column).value = "Note"
            logging.info(f"Added 'Note' column to Excel file at column {note_column}")
            wb.save(filepath)

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

        wb.save(filepath)
        logging.info("Final Excel file save completed")
        return "Success"

    except Exception as e:
        error_message = f"Main process error: {e}"
        logging.error(error_message)
        if wb is not None:
            try:
                sheet = wb.active
                note_column = None
                for col_idx, cell in enumerate(sheet[1], start=1):
                    if cell.value and isinstance(cell.value, str) and cell.value.lower() == "note":
                        note_column = col_idx
                        break
                if not note_column:
                    note_column = sheet.max_column + 1
                    sheet.cell(row=1, column=note_column).value = "Note"
                sheet.cell(row=2, column=note_column).value = f"Error: {str(e)}"
                wb.save(filepath)
                logging.info("Excel file saved with error note due to critical error")
            except Exception as save_err:
                logging.error(f"Failed to save Excel file: {save_err}")
        return error_message

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