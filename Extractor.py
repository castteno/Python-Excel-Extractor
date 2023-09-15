import tkinter as tk
from tkinter import filedialog, ttk
from ttkthemes import ThemedStyle
import pandas as pd
import tkinter.messagebox as messagebox
from PIL import Image, ImageTk
import os
import subprocess

from openpyxl import load_workbook
from pandas.api.types import CategoricalDtype


def combine_and_update(form_path, interns_detail_df, matching_criteria, error_label, root):
    try:
        form_df = pd.read_excel(form_path)

        for idx, intern_row in interns_detail_df.iterrows():
            criteria_value = intern_row[matching_criteria]
            if criteria_value in form_df[matching_criteria].values:
                idx = form_df.index[form_df[matching_criteria] == criteria_value].tolist()[0]
                for column in ['Track', 'Badge #', 'Name', 'Email']:
                    form_df.at[idx, column] = intern_row[column]
            else:
                new_data = pd.DataFrame([intern_row[['Track', 'Badge #', 'Name', 'Email']]])
                form_df = pd.concat([form_df, new_data], ignore_index=True)

        if 'GPA' in interns_detail_df.columns:
            gpa_dict = {row[matching_criteria]: row['GPA'] for _, row in interns_detail_df.iterrows()}
            form_df['GPA'] = form_df[matching_criteria].map(gpa_dict)
            form_df.sort_values(by='GPA', ascending=False, inplace=True)
        else:
            error_label.config(text="GPA column not found in interns details file.")
            return None

        with pd.ExcelWriter(form_path, engine='openpyxl', mode='a') as writer:
            sheet_name = 'Updated'
            count = 1
            while sheet_name in pd.ExcelFile(form_path).sheet_names:
                sheet_name = f'Updated_{count}'
                count += 1
            form_df.to_excel(writer, index=False, sheet_name=sheet_name)

        return form_path
    except PermissionError:
        error_label.config(text="The file is currently open. Please close it and try again.")
        return None

class GPAUpdaterApp:
    MAX_HEADERS = 5

    def __init__(self, root):
        self.root = root
        self.interns_detail_df = None
        self.form_path = None
        self.header_dropdowns = []
        self.header_labels = []

        image = Image.open(r"C:\Users\Castro\PycharmProjects\Data_Extractor\images\Background.jpeg")
        self.bg_image = ImageTk.PhotoImage(image)

        logo_image = Image.open(r"C:\Users\Castro\PycharmProjects\Data_Extractor\images\file.ico")
        self.logo = ImageTk.PhotoImage(logo_image)

        bg_label = tk.Label(self.root, image=self.bg_image)
        bg_label.place(relwidth=1, relheight=1)  # Fill the entire window

        self.root.title("Extractor")

        self.setup_custom_style()
        self.setup_gui()

    def setup_custom_style(self):
        self.custom_style = ThemedStyle(self.root)
        self.custom_style.theme_use("arc")

        # Making the background of the label transparent
        self.custom_style.configure("Custom.TLabel", background="transparent", foreground="black", font=("Helvetica", 12))
        # Configure a bold font for buttons
        self.custom_style.configure("Bold.TButton", font=("Helvetica", 10, "bold"))
        # Configure a small font for the + button
        self.custom_style.configure("Small.TButton", font=("Helvetica", 10, "bold"))

    def open_updated_file(self):
        if self.form_path:
            folder_path = os.path.dirname(self.form_path)  # Get the folder containing the file
            if os.name == 'nt':  # For Windows
                os.startfile(folder_path)
            else:
                opener = "open" if os.name == "posix" else "xdg-open"
                subprocess.call([opener, folder_path])
    def setup_gui(self):
        title_label = tk.Label(self.root, text="DATA TRANSFER", fg="black", bg="#EFEFEF",
                               font=("Helvetica", 20))
        title_label.place(relx=0.7, rely=0.15, anchor='center')

        btn1 = ttk.Button(self.root, text="SELECT MAIN DATA FILE", command=self.open_interns_file, style="Bold.TButton")
        btn1.place(relx=0.7, rely=0.22, anchor='center')

        btn2 = ttk.Button(self.root, text="SELECT FILE TO TRANSFER DATA TO IT", command=self.open_form_file, style="Bold.TButton")
        btn2.place(relx=0.7, rely=0.28, anchor='center',)

        self.matching_criteria_entry = ttk.Entry(self.root, justify=tk.CENTER)
        self.matching_criteria_entry.place(relx=0.7, rely=0.37, anchor='center')

        self.error_label = tk.Label(self.root, text="", fg="crimson", bg="#EFEFEF")
        self.error_label.place(relx=0.7, rely=0.46, anchor='center')

        combine_text = tk.Label(self.root, text="Enter related data in both files (Header)", fg="black", bg="#EFEFEF",
                               font=("Helvetica", 9,"bold"))
        combine_text.place(relx=0.7, rely=0.33, anchor='center')

        combine_button = ttk.Button(self.root, text="COMBINE", command=self.combine, style="Bold.TButton")
        combine_button.place(relx=0.7, rely=0.42, anchor='center')

        logo_button = tk.Button(self.root, image=self.logo, command=self.open_updated_file, bd=0, bg="#EFEFEF",
                                activebackground="#EFEFEF", relief="flat")
        logo_button.place(relx=0.7, rely=0.50, anchor='center')

    def open_interns_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.interns_detail_df = pd.read_excel(file_path)
            headers = self.interns_detail_df.columns.tolist()

            # Destroy all existing header dropdowns and labels
            for dropdown in self.header_dropdowns:
                dropdown.destroy()
            for label in self.header_labels:
                label.destroy()
            self.header_dropdowns = []
            self.header_labels = []

            # Create a fresh header dropdown
            self.create_header_selection(headers)

    def open_form_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.form_path = file_path
    def add_header(self, headers):
        if len(self.header_dropdowns) < self.MAX_HEADERS:
            self.add_header_dropdown(headers, len(self.header_dropdowns))
            self.error_label.config(text="")
        else:
            self.error_label.config(text="Maximum headers reached!")
        self.adjust_button_positions()

    def remove_last_header(self):
        if self.header_dropdowns:
            dropdown = self.header_dropdowns[-1]  # Get the last dropdown
            dropdown.destroy()
            self.header_dropdowns.pop()

            label = self.header_labels[-1]  # Get the last label
            label.destroy()
            self.header_labels.pop()

            self.adjust_button_positions()
            self.error_label.config(text="")

    def adjust_button_positions(self):
        self.add_button.place(relx=0.83, rely=0.5 + len(self.header_dropdowns) * 0.05, anchor='center')
        self.remove_button.place(relx=0.86, rely=0.5 + len(self.header_dropdowns) * 0.05, anchor='center')

    def add_header_dropdown(self, headers, idx):
        var = tk.StringVar(self.root)
        dropdown = ttk.Combobox(self.root, textvariable=var, values=headers)
        dropdown.place(relx=0.7, rely=0.55 + idx * 0.05, anchor='center')
        self.header_dropdowns.append(dropdown)

        num_label = ttk.Label(self.root, text=f"{len(self.header_dropdowns)}.", style="Custom.TLabel")
        num_label.place(relx=0.65, rely=0.55 + idx * 0.05, anchor='center')
        self.header_labels.append(num_label)

    def create_header_selection(self, headers):
        if hasattr(self, 'add_button'):
            self.add_button.destroy()
        if hasattr(self, 'remove_button'):
            self.remove_button.destroy()

        self.header_dropdowns = []
        self.header_labels = []

        self.add_header_dropdown(headers, 0)

        self.add_button = ttk.Button(self.root, text="+", width=1,
                                     command=lambda: self.add_header(headers),
                                     style="Small.TButton")
        self.add_button.place(relx=0.83, rely=0.5 + len(self.header_dropdowns) * 0.05, anchor='center')

        self.remove_button = ttk.Button(self.root, text="-", width=1,
                                        command=self.remove_last_header,
                                        style="Small.TButton")
        self.remove_button.place(relx=0.86, rely=0.5 + len(self.header_dropdowns) * 0.05, anchor='center')
    def combine(self):
        if self.form_path and self.interns_detail_df is not None and hasattr(self, 'header_dropdowns'):
            matching_criteria = self.matching_criteria_entry.get().strip()  # Trim whitespace

            if not matching_criteria:
                self.error_label.config(text="Please enter a matching criteria.", fg="crimson")
                return

            if matching_criteria not in self.interns_detail_df.columns:
                self.error_label.config(text="Matching criteria not found in interns details file.", fg="crimson")
                return

            selected_headers = [dropdown.get().strip() for dropdown in self.header_dropdowns if dropdown.get().strip()]

            if not selected_headers:
                self.error_label.config(text="Please select at least one header.", fg="crimson")
                return

            confirmation = messagebox.askyesno("Confirm Combine", "Are you sure you want to combine the data?")
            if confirmation:
                form_df = pd.read_excel(self.form_path)

                for header in selected_headers:
                    if header in form_df.columns:
                        for _, row in self.interns_detail_df.iterrows():
                            criteria_value = row[matching_criteria]
                            if criteria_value in form_df[matching_criteria].values:
                                idx = form_df.index[form_df[matching_criteria] == criteria_value].tolist()[0]
                                form_df.at[idx, header] = row[header]
                            else:
                                print(
                                    f"Couldn't import data for criteria value: {criteria_value}")  # This should be here
                    else:
                        form_df[header] = form_df[matching_criteria].map(
                            self.interns_detail_df.set_index(matching_criteria)[header])

                # Append missing records from the "Interns details" file to the "Preferences" file
                missing_records = self.interns_detail_df[
                    ~self.interns_detail_df[matching_criteria].isin(form_df[matching_criteria])]
                form_df = pd.concat([form_df, missing_records[selected_headers]], ignore_index=True)

                try:
                    with pd.ExcelWriter(self.form_path, engine='openpyxl', mode='a') as writer:
                        sheet_name = 'Updated'
                        count = 1
                        while sheet_name in pd.ExcelFile(self.form_path).sheet_names:
                            sheet_name = f'Updated_{count}'
                            count += 1
                        form_df.to_excel(writer, index=False, sheet_name=sheet_name)
                    self.error_label.config(text="Data combined successfully!", fg="green")
                except PermissionError:
                    self.error_label.config(text="The file is currently open. Please close it and try again.",
                                            fg="crimson")
            else:
                self.error_label.config(text="Please select both files and headers first.", fg="crimson")


def main():
    root = tk.Tk()
    root.configure(bg='#EFEFEF')

    # Center the window on screen
    window_width = 700
    window_height = 700
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calculate position
    x_position = (screen_width / 2) - (window_width / 2)
    y_position = (screen_height / 2) - (window_height / 2)

    # Set the position and size
    root.geometry(f"{window_width}x{window_height}+{int(x_position)}+{int(y_position)}")

    # Prevent the window from being resized
    root.resizable(False, False)

    # On Windows, remove maximize button
    root.state('normal')
    root.attributes('-disabled', True)
    root.attributes('-disabled', False)

    app = GPAUpdaterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
