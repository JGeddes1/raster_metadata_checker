import os
import openpyxl
import tkinter as tk
from tkinter import filedialog, Listbox, Scrollbar, END


def get_directory_path():
    directory = filedialog.askdirectory()
    return directory

def get_excel_file_path():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

def read_excel_file(file_path):
    filenames = []
    
    workbook = openpyxl.load_workbook(file_path)
    worksheet = workbook.active
    for row in worksheet.iter_rows(values_only=True):
        if row[0] and not str(row[0]).lower() in ["Filename","filename"]:  # Exclude common header names # Assuming filenames are in the first column
            filenames.append(row[0].strip().lower())
               # print("excel filenames: " +str(filenames))
        if row[2] and not str(row[2]).lower() in ["Subject Keyword 1","subjectword"]:
            subjectkeywords.add(row[2].strip().lower())
           
            print(subjectkeywords)
        if row[3] and not str(row[3]).lower() in ["Subject Keyword 2","subjectword"]:
            subjectkeywords.add(row[3].strip().lower())
        if row[4] and not str(row[4]).lower() in ["Subject Keyword 3","subjectword"]:
            subjectkeywords.add(row[4].strip().lower())
             
    return filenames

def list_files(directory):
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.tif']
    return [filename.lower() for filename in os.listdir(directory) if os.path.isfile(os.path.join(directory, filename)) and os.path.splitext(filename)[1].lower() in image_extensions]

def find_missing_files(directory_files, excel_filenames):
    missing_in_directory = [filename for filename in excel_filenames if filename not in directory_files]
    missing_in_metadata = [filename for filename in directory_files if filename not in excel_filenames]
    return missing_in_directory, missing_in_metadata

def browse_directory():
    directory_path.set(get_directory_path())

def browse_excel_file():
    excel_file_path.set(get_excel_file_path())

def check_files():
    directory = directory_path.get()
    excel_file1 = excel_file_path.get()
    excel_file2 = excel_file_path2.get()

    # Read filenames from both Excel files
    excel_filenames1 = read_excel_file(excel_file1)
    excel_filenames2 = read_excel_file(excel_file2) if excel_file2 else []
    # Combine filenames from both files
    excel_filenames = list(set(excel_filenames1 + excel_filenames2))

    directory_files = list_files(directory)

    missing_in_directory, missing_in_metadata = find_missing_files(directory_files, excel_filenames)

    # Clear existing items in the listbox
    missing_directory_listbox.delete(0, END)
    missing_metadata_listbox.delete(0, END)

    # Update listbox with missing files
    for file in missing_in_directory:
        missing_directory_listbox.insert(END, file)
    for file in missing_in_metadata:
        missing_metadata_listbox.insert(END, file)

    # Update listbox with subject keywords
    for keyword in subjectkeywords:
        keyword_listbox.insert(END, keyword)

    # Check if there are no issues and respond accordingly
    if not missing_in_directory and not missing_in_metadata:
        status_label.config(text="No missing files found! All good!", fg="green")
        # playsound('path_to_success_sound.mp3')  # Play a success notification sound
    else:
        status_label.config(text="Missing files detected. Check the lists above.", fg="red")

    
# Create main window
root = tk.Tk()
root.title("File Checker")

# Variables
directory_path = tk.StringVar()
excel_file_path = tk.StringVar()
subjectkeywords = set()

# Widgets
directory_label = tk.Label(root, text="Select Directory:")
directory_label.grid(row=0, column=0, sticky="w")

directory_entry = tk.Entry(root, textvariable=directory_path, width=50)
directory_entry.grid(row=0, column=1, padx=5, pady=5)

browse_directory_button = tk.Button(root, text="Browse", command=browse_directory)
browse_directory_button.grid(row=0, column=2, padx=5, pady=5)

excel_label = tk.Label(root, text="Select Excel File:")
excel_label.grid(row=1, column=0, sticky="w")

excel_entry = tk.Entry(root, textvariable=excel_file_path, width=50)
excel_entry.grid(row=1, column=1, padx=5, pady=5)

browse_excel_button = tk.Button(root, text="Browse", command=browse_excel_file)
browse_excel_button.grid(row=1, column=2, padx=5, pady=5)

check_button = tk.Button(root, text="Check Files", command=check_files)
check_button.grid(row=2, column=1, pady=10)

# Listbox for missing in directory
missing_directory_label = tk.Label(root, text="Missing from directory:")
missing_directory_label.grid(row=3, column=0, padx=5, pady=5)

scrollbar1 = Scrollbar(root)
scrollbar1.grid(row=4, column=1, sticky='ns')

missing_directory_listbox = Listbox(root, yscrollcommand=scrollbar1.set)
missing_directory_listbox.grid(row=4, column=0, columnspan=2, sticky='ew')
scrollbar1.config(command=missing_directory_listbox.yview)

# Listbox for missing in metadata
missing_metadata_label = tk.Label(root, text="Not listed in metadata:")
missing_metadata_label.grid(row=5, column=0, padx=5, pady=5)

scrollbar2 = Scrollbar(root)
scrollbar2.grid(row=6, column=1, sticky='ns')

missing_metadata_listbox = Listbox(root, yscrollcommand=scrollbar2.set)
missing_metadata_listbox.grid(row=6, column=0, columnspan=2, sticky='ew')
scrollbar2.config(command=missing_metadata_listbox.yview)


keyword_listbox = Listbox(root, yscrollcommand=scrollbar2.set)
keyword_listbox.grid(row=8, column=0, columnspan=2, sticky='ew')

# Variable for second Excel file path
excel_file_path2 = tk.StringVar()

# Label, entry, and button for the second Excel file
excel_label2 = tk.Label(root, text="Select Second Excel File:")
excel_label2.grid(row=2, column=0, sticky="w")

excel_entry2 = tk.Entry(root, textvariable=excel_file_path2, width=50)
excel_entry2.grid(row=2, column=1, padx=5, pady=5)

browse_excel_button2 = tk.Button(root, text="Browse", command=lambda: excel_file_path2.set(get_excel_file_path()))
browse_excel_button2.grid(row=2, column=2, padx=5, pady=5)

# Move the check button row down
check_button.grid(row=3, column=1, pady=10)


status_label = tk.Label(root, text="")
status_label.grid(row=7, column=0, columnspan=3, sticky="ew", padx=5, pady=5)

# Run the application
root.mainloop()
