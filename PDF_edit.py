from PyQt5.QtWidgets import QApplication, QWidget, QMessageBox, QPushButton, QLabel, QHBoxLayout, QVBoxLayout
from PyQt5.QtGui import QFont, QPixmap
from PyQt5.QtCore import Qt
import os          # For file and directory operations
from PyPDF2 import PdfMerger, PdfWriter, PdfReader 
import tkinter as tk  # For GUI dialog windows
from tkinter import filedialog, messagebox  # Specific GUI components
import pdfplumber  # For extracting text from PDF files


# Class
class Home(QWidget):
    # Constructor 
    def __init__(self):
        super().__init__()
        self.initUI()
        self.settings()
    # App Object and Design
    def initUI(self):
        self.split_button = QPushButton("Split PDF")  # Create split button
        self.merge_button = QPushButton("Merge PDF")  # Create merge button
        self.split_button.clicked.connect(self.split_click)
        self.merge_button.clicked.connect(self.merge_click)
        self.title = QLabel("pdf edit")              # Title
        self.title.setFont(QFont("Helvetica", 45, QFont.Bold))

    
        # CSS
        self.setStyleSheet("""              
            QWidget {
                background-color: #333;    
                color: #fff;
            }
                           
            QPushButton {
                background-color: #ffffff;
                color: #333;
                font-family: Segoe UI;
                font-weight: bold;
                font-size: 20px;
                border: 1px solid #fff;
                border-radius: 5px;
                padding: 5px 10px;               
            }
                           
            QPushButton:hover {
                background-color: #666;
                color: #dedede
            }
                           """)

        self.master = QVBoxLayout()

        self.title.setAlignment(Qt.AlignCenter)
        self.title.setContentsMargins(0, 20, 0, 40)
        self.master.addWidget(self.title)

        # Horizontal button layout
        button_row = QHBoxLayout()
        button_row.addWidget(self.split_button, alignment=Qt.AlignLeft)
        button_row.addSpacing(0)  # Spacer between buttons
        button_row.addWidget(self.merge_button, alignment=Qt.AlignRight)

        # Add spacer to push buttons to bottom
        self.master.addLayout(button_row)
        self.split_button.setFixedWidth(265)
        self.merge_button.setFixedWidth(265)
        self.split_button.setFixedHeight(100)
        self.merge_button.setFixedHeight(100)
        self.master.setContentsMargins(20, 100, 20, 150)

        self.setLayout(self.master)

    # App Settings
    def settings(self):
        self.setWindowTitle("PDF edit")      # Title on window
        self.setGeometry(250,250, 600, 500)  # Size of window

    # Button Click
    def button_click(self):
        self.split_pdf.clicked.connect(self.split_click)
        self.merge_pdf.clicked.connect(self.merge_click)


    # Split PDF
    def split_pdf(self):
        # Hide the main Tkinter root window
        root = tk.Tk()
        root.title("PDF Splitter")
        root.withdraw()

        # Ask user to select the large PDF file
        pdf_path = filedialog.askopenfilename(title="Select the PDF")

        def extract_first_valid_wbs(pages):
            for i, page in enumerate(pages):  
                found = False
                curr_page = pdf.pages[i]
                text = curr_page.extract_text()
                for line in text.splitlines():
                    words_list = line.split()
                    for word in words_list:
                        try:
                            if word[3:9] == "-3045-":
                                curr_WBS = word
                                if curr_WBS in wbs_elems.values():
                                    found = True
                                    break
                                if not curr_WBS in wbs_elems.values():
                                    wbs_elems[i + 1] = curr_WBS
                                    found = True
                                    break
                        except IndexError:
                            continue
                    if found == True:
                        break

        with pdfplumber.open(pdf_path) as pdf:
            wbs_elems = {} # List of WBS elements
            pages = pdf.pages
            extract_first_valid_wbs(pages)

            if len(wbs_elems) == 0:
                messagebox.showinfo("No WBS found", "No WBS found")

            print(wbs_elems)
                        
        reader = PdfReader(pdf_path)  
        output_dir = filedialog.askdirectory(title="Select folder to save split PDFs")
        if not output_dir:
            messagebox.showwarning("No Folder Selected", "Operation cancelled. No output folder selected.")
            exit()

        os.makedirs(output_dir, exist_ok=True)


        for i, key in enumerate(wbs_elems.keys()):
            start_page, filename = key, wbs_elems[key]

            # Determine the end page
            if i + 1 < len(wbs_elems):
                x = list(wbs_elems.keys())
                end_page = x[i + 1]
            else:
                end_page = len(reader.pages) + 1

            # Write pages
            writer = PdfWriter()
            for i in range(start_page, end_page):
                writer.add_page(reader.pages[i - 1])

            # Save the new PDF
            output_path = os.path.join(output_dir, f"{filename}.pdf")
            with open(output_path, "wb") as f_out:
                writer.write(f_out)

        # Reassurance
        messagebox.showinfo("Success", "Success")

    # Split Button Click
    def split_click(self):
        self.script = self.split_pdf()

    # How many pdfs
    def ask_how_many_pdfs(self):
        msg = QMessageBox(self)
        msg.setMinimumSize(500,400)
        msg.setWindowTitle("Merge PDFs")
        msg.setText("How many PDFs?")

        btn_2 = QPushButton("2")
        btn_3 = QPushButton("3")
        msg.addButton(btn_2, QMessageBox.AcceptRole)
        msg.addButton(btn_3, QMessageBox.AcceptRole)

        selected = msg.exec_()

        if msg.clickedButton() == btn_2:
            return 2
        elif msg.clickedButton() == btn_3:
            return 3
        else:
            return None
        
    # Merge PDF 
    def merge_pdf(self):
        # Hide the main Tkinter root window
        count = self.ask_how_many_pdfs()
        if count is None:
            return  # User closed the dialog


        # Continue with your merging logic, e.g.:
        if count == 2:
            self.merge_two_pdfs()
        elif count == 3:
            self.merge_three_pdfs()
        
    # Merge two
    def merge_two_pdfs(self):
        root = tk.Tk()
        root.title("PDF Joiner")
        root.withdraw()

        # Ask user to select the summary folder
        summary_folder = filedialog.askdirectory(title="Select the summary hours folder")
        summary_map = set()

        # Save all file names to a set
        if summary_folder:
            for filename in os.listdir(summary_folder):
                if not filename in summary_map:
                    summary_map.add(filename)

        # Ask user to select the detail folder
        detail_folder = filedialog.askdirectory(title="Select the detail hours folder")
        detail_map = set()

        # Save all file names to a set
        if detail_folder:
            for filename in os.listdir(detail_folder):
                if not filename in detail_map:
                    detail_map.add(filename)

        # Save same file names in a seperate map
        matched = set()
        for name in summary_map:
            name = name[:-4]
            matched.add(name)
        for name in detail_map:
            name = name[:-4]
            matched.add(name)
        
           

        # Ask user to choose an output directory
        output_dir = filedialog.askdirectory(title="Select folder to save combined PDFs")
        if not output_dir:
            messagebox.showwarning("No Folder Selected", "Operation cancelled. No output folder selected.")
            exit()
        os.makedirs(output_dir, exist_ok=True)

        # Combine PDFs
        for name in matched:
            merger = PdfMerger()

            # Add summary pdf
            for filename1 in os.listdir(summary_folder):
                if name == filename1[:-4]:  # Remove .pdf from end to compare 
                    full_path1 = os.path.join(summary_folder, filename1)
                    merger.append(full_path1)

            # Add detail pdf
            for filename2 in os.listdir(detail_folder):
                if name == filename2[:-4]:
                    full_path2 = os.path.join(detail_folder, filename2)
                    merger.append(full_path2)

            # Save to chosen file
            output_path = os.path.join(output_dir, f"{name}.pdf")
            with open(output_path, "wb") as f_out:
                merger.write(f_out)
            merger.close()

        # Reassurance
        messagebox.showinfo("Success", "Success")

    # Merge three
    def merge_three_pdfs(self):
        root = tk.Tk()
        root.title("PDF Joiner")
        root.withdraw()

        # Ask user to select the summary folder
        summary_folder = filedialog.askdirectory(title="Select the summary hours folder")
        summary_map = set()

        # Save all file names to a set
        if summary_folder:
            for filename in os.listdir(summary_folder):
                if not filename in summary_map:
                    summary_map.add(filename)

        # Ask user to select the detail folder
        detail_folder = filedialog.askdirectory(title="Select the detail hours folder")
        detail_map = set()

        # Save all file names to a set
        if detail_folder:
            for filename in os.listdir(detail_folder):
                if not filename in detail_map:
                    detail_map.add(filename)

        # Ask user to select the material folder
        material_folder = filedialog.askdirectory(title="Select the material folder")
        material_map = set()

        # Save all file names to a set
        if material_folder:
            for filename in os.listdir(material_folder):
                if not filename in material_map:
                    material_map.add(filename)

        # Save same file names in a seperate map
        matched = set()
        for name in summary_map:
            name = name[:-4]
            matched.add(name)
        for name in detail_map:
            name = name[:-4]
            matched.add(name)
        for name in material_map:
            name = name[:-4]
            matched.add(name)
                

        # Ask user to choose an output directory
        output_dir = filedialog.askdirectory(title="Select folder to save combined PDFs")
        if not output_dir:
            messagebox.showwarning("No Folder Selected", "Operation cancelled. No output folder selected.")
            exit()
        os.makedirs(output_dir, exist_ok=True)

        # combine PDFs
        for name in matched:
            merger = PdfMerger()

            # Add summary pdf
            for filename1 in os.listdir(summary_folder):
                if name == filename1[:-4]:  # Remove .pdf from end to compare 
                    full_path1 = os.path.join(summary_folder, filename1)
                    merger.append(full_path1)

            # Add detail pdf
            for filename2 in os.listdir(detail_folder):
                if name == filename2[:-4]:
                    full_path2 = os.path.join(detail_folder, filename2)
                    merger.append(full_path2)

            # Add material pdf
            for filename3 in os.listdir(material_folder):
                if name == filename3[:-4]:  # Remove .pdf from end to compare 
                    full_path3 = os.path.join(material_folder, filename3)
                    merger.append(full_path3)

            # Save to chosen file
            output_path = os.path.join(output_dir, f"{name}.pdf")
            with open(output_path, "wb") as f_out:
                merger.write(f_out)
            merger.close()

        # Reassurance
        messagebox.showinfo("Success", "Success")

    # Merge Button Click
    def merge_click(self):
        self.script = self.merge_pdf()
    
    

# Main Run
if __name__ == "__main__":
    app = QApplication([])
    main = Home()
    main.show()
    app.exec_()