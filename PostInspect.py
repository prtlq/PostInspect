import os
import shutil
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from tkinter.scrolledtext import ScrolledText



# FIX THE SHEETSHIT

class PostInspectApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Post Inspect App")
        self.excel_file = ''
        self.photo_directory = ''
        self.columns = [chr(i) for i in range(ord('A'), ord('X'))]
        self.valid_extensions = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff'}
        
        self.build_gui()

    def build_gui(self):
        tab_control = ctk.CTkTabview(self.root, width=800, height=600)
        tab_control.pack(expand=1, fill="both")

        fotofinder_tab = tab_control.add("FotoFinder")
        help_tab = tab_control.add("Release Note")

        self.build_fotofinder_tab(fotofinder_tab)
        self.build_help_tab(help_tab)

    def build_fotofinder_tab(self, tab):

        left_frame = ctk.CTkFrame(tab)
        left_frame.grid(row=0, column=0, sticky="nswe", padx=20, pady=20)

        right_frame = ctk.CTkFrame(tab)
        right_frame.grid(row=0, column=1, sticky="nswe", padx=20, pady=20)

        tab.grid_columnconfigure(0, weight=1)
        tab.grid_columnconfigure(1, weight=3)
        tab.grid_rowconfigure(0, weight=1)
        
        left_frame.grid_rowconfigure(3, weight=1)
        left_frame.grid_rowconfigure(3, weight=1)

        right_frame.grid_rowconfigure(0, weight=1)

        ctk.CTkButton(left_frame, text="Select Photos Folder", command=self.select_photo_directory).grid(row=0, column=0, sticky="ew", pady=5)
        ctk.CTkButton(left_frame, text="Select Excel File", command=self.select_excel_file).grid(row=1, column=0, sticky="ew", pady=5)

        dropdown_frame = ctk.CTkFrame(left_frame)
        dropdown_frame.grid(row=2, column=0, sticky="ew", pady=5)

        self.spinner_col_skada_nr = ctk.CTkComboBox(dropdown_frame, values=self.columns)
        self.spinner_col_skada_nr.set("BaTMan Nr.")
        self.spinner_col_skada_nr.pack(side="left", fill="x", expand=True)

        self.spinner_col_photo = ctk.CTkComboBox(dropdown_frame, values=self.columns)
        self.spinner_col_photo.set("Photo Column")
        self.spinner_col_photo.pack(side="right", fill="x", expand=True)

        ctk.CTkButton(left_frame, text="Process Photos", command=self.process_photos).grid(row=4, column=0, sticky="ew", pady=5)

        self.message_box = ScrolledText(right_frame, wrap=tk.WORD, width=55, height=15)
        self.message_box.grid(row=0, column=0, sticky="nsew")


    def build_help_tab(self, tab):

        help_text = """
        Instructions

        1. Select the Photos Folder:
        - Click on the "Select Photos Folder" button.
        - From the dialog, select the parent directory (folder) that includes the inspection photos.
        - These photos may be stored in subdirectories (e.g., Photos Day 1).

        2. Select the Excel File:
        - Click on the "Select Excel File" button and choose the Excel file.
        - Ensure that the file contains the columns 'Skada Nr BatMan' and the photo column 'Foto' with a timestamp string in the format 24_04_17_00_40_58_tunnelXXX.

        3. Choose Columns:
        - From the "BatMan Nr." dropdown, select the Skada Nr BatMan column (usually Column C in the Excel file).
        - From the "Timestamps col." dropdown, select the Foto column (usually Column S in the Excel file).

        4. Process Photos:
        - Click on "Process Photos". New folders will be created, named according to the BatMan number, and will include the corresponding photos.
        - The newly created folders will be located in the same directory as the PostInspect.exe.

        --- 

        Release Notes

        Public VR0.2

        Bug Fix: 
        The issue with duplicating photos in the last folder is resolved.
        
        Visual Enhancements:
        Applied consistent padding and spacing for a more polished layout.
        Updated fonts and colors for buttons and labels to make the interface more user-friendly.
        Used customtkinter components consistently for a modern look.

        Public VR0.1

        - Integration of a message box for reporting.
        - Support for processing photos stored in subdirectories.
        - Enhanced indexing and matching algorithm.
        - Added a help tab.

        --- 

        This toolkit was developed by Pouria Taleghani, AFRY. For support and suggestions, please contact me.
        """
        help_text_box = ScrolledText(tab, wrap=tk.WORD, width=80, height=20)
        help_text_box.insert(tk.END, help_text)
        help_text_box.config(state=tk.DISABLED)
        help_text_box.pack(fill="both", expand=True)

    def select_excel_file(self):
        self.excel_file = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])
        self.update_message_box(f"Selected Excel File: {self.excel_file}")

    def select_photo_directory(self):
        self.photo_directory = filedialog.askdirectory(title="Select Photos Folder")
        self.update_message_box(f"Selected Photos Folder: {self.photo_directory}")

    def process_photos(self):
        if not self.excel_file or not self.photo_directory:
            messagebox.showerror("Error", "Please select both the Excel file and photos folder.")
            return

        col_photo = self.spinner_col_photo.get()  # Get the column for photos from dropdown
        col_skada_nr = self.spinner_col_skada_nr.get()  # Get the column for BaTMan Nr from dropdown

        if col_photo == 'Timestamps col.' or col_skada_nr == 'BaTMan Nr.':
            messagebox.showerror("Error", "Please specify both column names.")
            return

        try:
            df = pd.read_excel(self.excel_file)

            # Convert the dropdown letters to corresponding index
            col_photo_index = ord(col_photo.upper()) - ord('A')
            col_skada_nr_index = ord(col_skada_nr.upper()) - ord('A')

            # Check if these indices are within the range of available columns
            if col_photo_index >= len(df.columns) or col_skada_nr_index >= len(df.columns):
                messagebox.showerror("Error", "Selected column is out of bounds. Please check the Excel file.")
                return

            # Define the extract_timestamps function here
            def extract_timestamps(text):
                if isinstance(text, str):  # Ensure the input is a string
                    parts = text.split('|')  # Split by delimiter (assumed '|')
                    timestamps = [part.strip() for part in parts if part.strip()]  # Clean up and extract non-empty parts
                    return timestamps
                return []

            # Extract timestamps using the defined function
            df['Timestamps'] = df.iloc[:, col_photo_index].apply(extract_timestamps)

            report_lines = []
            processed_photos = set()

            def find_and_copy_photos(row):
                value_skada_nr = row.iloc[col_skada_nr_index]
                
                # Ensure that BaTMan Nr is a valid number or extract numeric part
                try:
                    index_value = int(float(value_skada_nr))
                except ValueError:
                    self.update_message_box(f"Skipping row with invalid BaTMan Nr: {value_skada_nr}")
                    return

                folder_name = f"sk{index_value}"
                timestamps = row['Timestamps']

                if not timestamps:
                    return

                copied_photos = []

                for root, _, files in os.walk(self.photo_directory):
                    if root.startswith(os.path.join(self.photo_directory, folder_name)):
                        continue
                    for filename in files:
                        if filename in processed_photos:
                            continue
                        if any(filename.lower().endswith(ext) for ext in self.valid_extensions):
                            source_file = os.path.join(root, filename)

                            matched = False
                            for timestamp in timestamps:
                                timestamp_pattern = r'\b' + re.escape(timestamp) + r'\b'
                                if re.search(timestamp_pattern, filename):
                                    matched = True
                                    break

                            if matched:
                                destination_folder = os.path.join(self.photo_directory, folder_name)
                                if not os.path.exists(destination_folder):
                                    os.makedirs(destination_folder, exist_ok=True)
                                destination_file = os.path.join(destination_folder, filename)
                                shutil.copy(source_file, destination_file)
                                copied_photos.append(filename)
                                processed_photos.add(filename)

                if copied_photos:
                    report_lines.append(f"Folder: {folder_name}")
                    for photo in copied_photos:
                        report_lines.append(f"  - {photo}")
                    report_lines.append("")

            df.apply(find_and_copy_photos, axis=1)

            # Unprocessed photos check
            unprocessed_photos = []
            for root, _, files in os.walk(self.photo_directory):
                if any(root.startswith(os.path.join(self.photo_directory, f"sk{int(float(b))}")) for b in df.iloc[:, col_skada_nr_index].dropna()):
                    continue
                for filename in files:
                    if any(filename.lower().endswith(ext) for ext in self.valid_extensions):
                        if filename not in processed_photos:
                            unprocessed_photos.append(filename)

            if unprocessed_photos:
                report_lines.append("Photos not placed in any folder:")
                for photo in unprocessed_photos:
                    report_lines.append(f"  - {photo}")
                report_lines.append("")

            report_file_path = os.path.join(self.photo_directory, "photo_organization_report.txt")
            with open(report_file_path, 'w') as report_file:
                report_file.write("\n".join(report_lines))

            self.update_message_box("\n".join(report_lines))

            messagebox.showinfo("Success", "Photos have been organized")

        except Exception as e:
            messagebox.showerror("Error", str(e))

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def update_message_box(self, message):
        self.message_box.insert(tk.END, f"\n{message}\n")
        self.message_box.yview(tk.END)


if __name__ == '__main__':
    ctk.set_appearance_mode("Light")  # Options: "System" (default), "Dark", "Light"
    ctk.set_default_color_theme("dark-blue")  # Options: "blue" (default), "green", "dark-blue"

    root = ctk.CTk()
    app = PostInspectApp(root)
    root.mainloop()
