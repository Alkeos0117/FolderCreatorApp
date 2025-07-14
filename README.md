# 📂 FolderCreatorApp

**FolderCreatorApp** is a Python-based desktop GUI application that allows users to automatically create folders from an Excel file. Designed for professionals in logistics, quality, and administration, this tool provides a clean interface and intelligent error handling to simplify repetitive folder creation tasks.

---

## 🧩 Features

✅ Intuitive step-by-step user interface  
✅ Select Excel file and destination folder visually  
✅ Automatic detection of folder names from column `Folder` or `S/N`  
✅ Smart button activation and visual feedback  
✅ Error messages for missing columns or file issues  
✅ Summary message on successful folder creation

---

## 🖥️ Screenshots

![App Screenshot](https://via.placeholder.com/600x300?text=GUI+Preview)

---

## 🚀 How to Use

1. Install Python 3.x and required libraries:
   ```bash
   pip install pandas openpyxl
2. Run the application:
python folder_creator.py

Follow these steps in the GUI:

Select the Excel file with folder names

Select the destination folder

Click Create Folders

📌 Your Excel file must contain a column named Folder or S/N.

📁 Excel File Format Example
S/N
1164751-001
1164751-002
1164751-003

The app will create folders using the values from this column.

⚙️ Requirements
Python 3.7+

pandas

openpyxl

tkinter (comes pre-installed with Python on Windows)

Install dependencies (if needed):

bash
Kopieren
Bearbeiten
pip install pandas openpyxl
📦 Repository Structure
bash
Kopieren
Bearbeiten
FolderCreatorApp/
├── folder_creator.py     # Main GUI script
├── README.md             # Documentation
└── .gitignore            # Optional file to exclude temp files
👤 Author
Wilmar Rodríguez
GitHub Profile

📄 License
This project is licensed under the MIT License – see the LICENSE file for details.
