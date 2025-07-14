import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class FolderCreatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Folder Creator from Excel")
        self.root.geometry("500x300")
        self.root.resizable(False, False)

        # Variables
        self.excel_path = None
        self.dest_folder = None
        self.df = None

        # Widgets
        self.label_status = ttk.Label(root, text="Step 1: Select Excel file", foreground="blue")
        self.label_status.pack(pady=10)

        self.btn_excel = tk.Button(root, text="Select Excel File", command=self.select_excel, bg="lightgray")
        self.btn_excel.pack(pady=5)

        self.label_excel = ttk.Label(root, text="No file selected", foreground="red")
        self.label_excel.pack()

        self.btn_folder = tk.Button(root, text="Select Destination Folder", command=self.select_folder, state="disabled", bg="lightgray")
        self.btn_folder.pack(pady=10)

        self.label_folder = ttk.Label(root, text="No folder selected", foreground="red")
        self.label_folder.pack()

        self.btn_create = tk.Button(root, text="Create Folders", command=self.create_folders, state="disabled", bg="gray")
        self.btn_create.pack(pady=15)

    def select_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.df = df
                self.excel_path = file_path
                self.label_excel.config(text=f"✔ File loaded: {os.path.basename(file_path)}", foreground="green")
                self.label_status.config(text="Step 2: Select destination folder", foreground="blue")
                self.btn_excel.config(bg="lightgreen")
                self.btn_folder.config(state="normal", bg="lightgray")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read Excel file: {e}")

    def select_folder(self):
        folder_path = filedialog.askdirectory(title="Select Destination Folder")
        if folder_path:
            self.dest_folder = folder_path
            self.label_folder.config(text=f"✔ Folder selected: {folder_path}", foreground="green")
            self.label_status.config(text="Step 3: Click 'Create Folders'", foreground="blue")
            self.btn_folder.config(bg="lightgreen")
            self.btn_create.config(state="normal", bg="orange")

    def create_folders(self):
        if self.df is None or self.dest_folder is None:
            messagebox.showerror("Error", "Missing Excel file or destination folder.")
            return

        if 'Folder' not in self.df.columns:
            available = ', '.join(self.df.columns)
            messagebox.showerror("Error", f"Column 'Folder' not found.\nAvailable columns: {available}")
            return

        try:
            created = 0
            for name in self.df['Folder']:
                if pd.notna(name):
                    path = os.path.join(self.dest_folder, str(name))
                    os.makedirs(path, exist_ok=True)
                    created += 1

            messagebox.showinfo("Success", f"{created} folders created successfully.")
            self.btn_create.config(bg="lightgreen")
            self.label_status.config(text="✅ Done", foreground="green")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while creating folders:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FolderCreatorApp(root)
    root.mainloop()
