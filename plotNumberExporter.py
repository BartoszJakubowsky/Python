import os
import re
import tkinter as tk
from tkinter import filedialog
import openpyxl

def choose_folder():
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory()
    return folder_selected

#extract numbers, numbners devided by "," or "_"
def extract_numbers(filename):
    # Wyszukiwanie liczb po przeciku, połączonych podłogą "_" oraz sekwencji wyłącznie z cyfr
    matches = re.findall(r'\d+[_]\d+|\d+(?=,)|\b\d+\b', filename)
    return matches

def create_excel(data, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for number in data:
        ws.append([number])
    wb.save(output_path)

def main():
    folder = choose_folder()
    if not folder:
        print("Folder not selected.")
        return

    all_numbers = []
    for filename in os.listdir(folder):
        numbers = extract_numbers(filename)
        all_numbers.extend(numbers)

    if not all_numbers:
        print("No numbers found in the filenames.")
        return

    output_path = os.path.join(folder, 'extracted_numbers.xlsx')
    create_excel(all_numbers, output_path)
    print(f"Excel file created at: {output_path}")

if __name__ == "__main__":
    main()
