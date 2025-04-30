import tkinter
from tkinter import ttk, filedialog, messagebox
import sv_ttk

from script import run_patent_scraper
import os
from dotenv import load_dotenv


class MainWindow:

    def __init__(self, root):
        # dark theme
        sv_ttk.set_theme("dark")

        # root
        self.root = root
        self.root.title("Patent Webscraper")

        # frame1
        self.frame1 = ttk.Frame(root, padding=20)
        self.frame1.pack()

        # welcome text
        self.welcome = ttk.Label(
            self.frame1,
            text="Welcome to the Patent Webscraper",
            font=("Arial", 20),
            justify="center",
        )
        self.welcome.grid(column=0, row=0, pady=(0, 5))

        # filepath text
        self.file_path_label = ttk.Label(
            self.frame1,
            text="Step 1: Select a file. \nSelected file: N/A",
            justify="center",
        )
        self.file_path_label.grid(column=0, row=1, pady=(5, 0))

        # upload button
        self.upload = ttk.Button(
            self.frame1, text="Upload", command=self.get_upload_file
        )
        self.upload.grid(column=0, row=2, pady=(5, 10))

        # export file name label
        self.export_label = ttk.Label(
            self.frame1,
            text="Step 2: Give a name to the export file. You do not need to add '.xlsx' \nExport file name:",
            justify="center",
        )
        self.export_label.grid(column=0, row=3, pady=(5, 0))

        # export name entry
        self.export_var = tkinter.StringVar()
        self.export_entry = ttk.Entry(
            self.frame1, textvariable=self.export_var, width=40
        )
        self.export_entry.grid(column=0, row=4, pady=(5, 10))
        self.export_var.trace_add("write", self.update_export_name)

        # begin program button
        # this should be greyed out unless all info is ready
        self.begin_button = ttk.Button(
            self.frame1, text="Go!", command=self.check_all_info
        )
        self.begin_button.grid(column=0, row=5, pady=(10, 5))

        # status text box
        self.status_label = ttk.Label(
            self.frame1, text="Status: Nothing is currently happening"
        )
        self.status_label.grid(column=0, row=6)

        self.update_min_size()

    def update_min_size(self):
        # Prevent window from being resized smaller than initial size
        self.root.update_idletasks()  # Ensure widgets are fully rendered
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        self.root.minsize(width, height)

    def get_upload_file(self):
        self.file_path = filedialog.askopenfilename(
            title="Select a file",
            filetypes=[("Excel files", "*.xlsx")],
        )

        self.file_path_label.config(
            text=f"Step 1: Select a file. \nSelected file: {self.file_path}"
        )

    def update_export_name(self, *args):
        self.export_name = self.export_var.get()
        if self.export_name:
            label_text = f"Step 2: Give a name to the export file. You do not need to add '.xlsx'\nExport file name: {self.export_name}.xlsx"
        else:
            label_text = f"Step 2: Give a name to the export file. You do not need to add '.xlsx'\nExport file name:"
        self.export_label.config(text=label_text)

    def log_status(self, message):
        self.status_label.config(text=f"Status:\n{message}", justify="center")
        self.root.update()

    def check_all_info(self):
        missing = []

        if not hasattr(self, "file_path") or not self.file_path:
            missing.append("Import file")

        if not hasattr(self, "export_name") or not self.export_name:
            missing.append("Export filename")

        if missing:
            messagebox.showerror(
                "Missing Required Information",
                "Please provide the following before continuing:\n\n"
                + "\n".join(missing),
            )
            return False
        else:
            confirmed = messagebox.askokcancel(
                "Ready to go",
                f"Would you like to proceed with:\n{self.file_path} → {self.export_name}.xlsx?\n\nThis will open Firefox and this window may go unresponsive.",
            )

            if confirmed:
                self.status_label.config(text="Status: \nRunning script...")
                self.root.update()

                load_dotenv()
                gemini_key = os.getenv("GEMINI_KEY")
                export_path = f"Files/{self.export_name}.xlsx"

                try:
                    run_patent_scraper(
                        self.file_path, export_path, gemini_key, self.log_status
                    )
                    self.status_label.config(
                        text=f"Status: Done! Export saved to:\n{export_path}",
                        justify="center",
                    )
                except Exception as e:
                    messagebox.showerror("Error", f"Something went wrong:\n{e}")
                    self.status_label.config(text="Status: Failed. See error.")
