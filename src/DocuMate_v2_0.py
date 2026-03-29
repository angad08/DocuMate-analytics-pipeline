"""
====================================================================
 DocuMate v2.0 — Birth Registration
====================================================================

 Second version of DocuMate, created after v1 validated the idea.
 This version focuses on usability, safety, and portability while
 keeping the same core workflow.

--------------------------------------------------------------------
 WHAT'S NEW IN v2 (vs v1)
--------------------------------------------------------------------

 1. USER FEEDBACK (POPUPS)
    - Replaces console-only output with Windows popup messages.
    - Shows clear messages for:
        • Successful completion (count + time taken)
        • No records found
        • Errors with explanation

 2. BASIC ERROR HANDLING
    - Handles common failure cases gracefully instead of crashing.
    - Examples:
        • Excel file open → prompts user to close it
        • Unexpected error → displays error details in popup

 3. EARLY EXIT WHEN NO DATA
    - Stops immediately if no records are in "IN PROCESS" status,
      instead of running the full pipeline silently.

 4. STANDALONE EXECUTION SUPPORT
    - Designed to be packaged as a standalone .exe using PyInstaller.
    - Does not require Python to be installed on the target machine.

 5. PORTABLE FILE PATHS
    - Uses paths relative to the program location instead of
      hardcoded absolute paths, enabling use across machines.

--------------------------------------------------------------------
 CORE FUNCTIONALITY (UNCHANGED FROM v1)
--------------------------------------------------------------------

 1. Reads data from data/DocuMate_DataFrame.xlsx → DocuMateSRC
 2. Filters rows where STATUS == "IN PROCESS"
 3. Formats date fields to dd/mm/yyyy
 4. Generates one Word document per record using a template
    → saved to output_files/ as {serial}.docx
 5. Displays summary (records processed + execution time)

--------------------------------------------------------------------
 KNOWN LIMITATIONS
--------------------------------------------------------------------

 - No duplicate detection (existing files may be overwritten)
 - No logging to file
 - Popup notifications are Windows-only

====================================================================
"""


import pandas as pd              # For reading and manipulating Excel data
from docxtpl import DocxTemplate  # For filling Word templates with data
import time                       # For tracking how long the process takes
import os                         # For building file paths
import sys                        # For detecting if running as .exe or .py
import ctypes                     # For showing Windows popup messages


# ====================================================================
# BirthRegistrationProcessor
# --------------------------------------------------------------------
# This is the main engine of DocuMate v2.
# It wraps the entire workflow — read, filter, format, generate —
# into one object so you can configure it once and call .run()
# ====================================================================

class BirthRegistrationProcessor:

    def __init__(self, excel_path, sheet_name, template_path, output_folder, update_existing=False):
        """
        Sets up all the paths and settings the processor needs.

        - excel_path:      Full path to the Excel file with records
        - sheet_name:      Which sheet to read from (e.g. "Sheet2")
        - template_path:   Full path to the Word template (.docx)
        - output_folder:   Where generated documents will be saved
        - update_existing:  (Not yet used) Reserved for future logic
                           to decide whether to overwrite existing files
        """
        self.excel_path = excel_path
        self.sheet_name = sheet_name
        self.template_path = template_path
        self.output_folder = output_folder
        self.data = None                    # Will hold the Excel data once loaded
        self.update_existing = update_existing


    def load_data(self):
        """
        Step 1: Read the Excel file into memory.
        Loads all rows from the specified sheet into a DataFrame.
        No filtering happens here — just raw data loading.
        """
        self.data = pd.read_excel(self.excel_path, sheet_name=self.sheet_name)


    def filter_in_process(self):
        """
        Step 2: Keep only the records that need processing.
        Filters the data to rows where the STATUS column says
        "IN PROCESS" (case-insensitive).
        Everything else is ignored.
        """
        if self.data is not None:
            self.data = self.data[self.data["STATUS"].str.upper() == "IN PROCESS"]


    def format_dates(self, date_columns):
        """
        Step 3: Clean up date columns so they look good in documents.
        Converts raw Excel dates (which can look like "2025-01-15")
        into a clean dd/mm/yyyy format (e.g. "15/01/2025").

        - date_columns: list of column names to format
          (e.g. ["Registration_date"])
        """
        for col in date_columns:
            if col in self.data.columns:
                self.data[col] = pd.to_datetime(
                    self.data[col], errors="coerce"
                ).dt.strftime("%d/%m/%Y")


    def generate_documents(self):
        """
        Step 4: The main event — create Word documents.

        For each filtered record:
        1. Opens the Word template
        2. Fills in all placeholders with that row's data
        3. Saves it as a new .docx file named after the serial number

        Naming logic:
        - Serial "123/2025" → filename becomes "123.docx"
        - Only the part before "/" is used
        """
        for idx, row in self.data.iterrows():
            doc = DocxTemplate(self.template_path)

            # Convert the row into a dictionary so the template engine
            # can match column names to placeholders in the Word file
            context = row.to_dict()
            doc.render(context)

            # Extract just the number part from Serial for the filename
            # e.g. "123/2025" → "123"
            serial = str(row.get("Serial", f"row{idx+1}")).split("/")[0]

            output_file = f"{self.output_folder}/{serial}.docx"
            doc.save(output_file)
            print(f"✅ Created: {serial}.docx at {os.path.abspath(output_file)}")


    def run(self):
        """
        Runs the full pipeline in order:
        Load → Filter → Check → Format → Generate → Summary

        Wrapped in try/except so errors don't crash the program —
        instead, the user sees a popup explaining what went wrong.
        """
        try:
            start_time = time.time()

            # --- Load ---
            print("📄 Reading Excel data...")
            self.load_data()

            # --- Filter ---
            print("🔍 Searching new records...")
            self.filter_in_process()

            # --- Check: anything to process? ---
            # If no records match "IN PROCESS", stop here and tell the user.
            # This was a pain point in v1 — it would just silently do nothing.
            if self.data is None or self.data.empty:
                message = (
                    "🔍 Found no records to populate.\n\n"
                    "🧠 Powered by ENN — Fuelled by curiosity, refined by data-driven clarity."
                )
                ctypes.windll.user32.MessageBoxW(0, message, "DocuMate", 0)
                return

            # --- Format dates ---
            self.format_dates(["Registration_date"])
            print(f"DocuMate : 🔍 Found {len(self.data)} new records to populate.\n")

            # --- Generate documents ---
            print("⚙️ Populating Word documents...\n")
            self.generate_documents()

            # --- Success popup ---
            total_time = time.time() - start_time
            success_message = (
                f"✅ Mission accomplished!\nPopulated {len(self.data)} documents successfully "
                f"in {total_time:.2f} seconds.\n\n"
                "🧠 Powered by ENN — Fuelled by curiosity, refined by data-driven clarity."
            )
            ctypes.windll.user32.MessageBoxW(0, success_message, "DocuMate", 0)

        except PermissionError:
            # This happens when the Excel file is already open in another program.
            # Excel locks the file, so Python can't read it.
            error_message = (
                "❌ I cannot engage the populate launch because the Excel file is open.\n"
                "Please close it and try again later.\n\nMission aborted."
            )
            ctypes.windll.user32.MessageBoxW(0, error_message, "DocuMate", 0)

        except Exception as e:
            # Catch-all for anything unexpected.
            # Shows the actual error message so the user (or developer)
            # can figure out what went wrong.
            unknown_message = (
                f"⚠️ DocuMate encountered an unexpected error:\n{e}\n\n"
                "🧠 Powered by ENN — Fuelled by curiosity, refined by data-driven clarity."
            )
            ctypes.windll.user32.MessageBoxW(0, unknown_message, "DocuMate", 0)


# ====================================================================
# ENTRY POINT
# --------------------------------------------------------------------
# This block runs when you execute the script directly (or the .exe).
# It figures out where the program lives, sets up all the file paths,
# and kicks off the processor.
# ====================================================================

if __name__ == "__main__":
    print("\n🚀 ENN : Initializing hypersonic missile DocuMate launch...")
    print("DocuMate : Alright ENN , initializing populate launch 🚀\n")

    # Figure out the base directory:
    # - If running as a PyInstaller .exe → the folder containing the .exe
    # - If running as a .py script     → the folder containing the script
    # This makes all file paths work regardless of how it's launched.
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(os.path.dirname(sys.executable))
    else:
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    # Create the processor with all paths configured relative to base_dir.
    # This means the folder structure should look like:
    #
    #   base_dir/
    #   ├── data/DocuMate_DataFrame.xlsx  ← input data
    #   ├── templates/DOCUMENT_TEMPLATE_FILE.docx  ← Word template
    #   └── output_files/                               ← output goes here
    #
    docuMateAgent = BirthRegistrationProcessor(
        excel_path=os.path.join(base_dir, "data", "DocuMate_DataFrame.xlsx"),
        sheet_name="DocuMateSRC",
        template_path=os.path.join(base_dir, "templates", "DOCUMENT_TEMPLATE_FILE.docx"),
        output_folder=os.path.join(base_dir,"output_files"),
        update_existing=True
    )

    # Run the full pipeline
    docuMateAgent.run()