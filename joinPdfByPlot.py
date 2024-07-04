import os
import re
import tkinter as tk
from tkinter import filedialog
import openpyxl
from docx2pdf import convert
import PyPDF2

def choose_folder():
    root = tk.Tk()
    root.withdraw()
    folder_docs = filedialog.askdirectory(title='Wybierz folder ze zgodami', mustexist=True)
    folder_maps = filedialog.askdirectory(title='Wybierz folder z mapami', initialdir=folder_docs)
    return [folder_docs, folder_maps]

def create_new_folder(docs_folder):
    target_path = os.path.join(docs_folder, "docs-pdf")
    if (os.path.exists(target_path)):
        return target_path
    
    os.makedirs(target_path)
    return target_path

def extract_numbers(filename):
    matches = re.findall(r'\d+[_]\d+|\d+(?=,)|\b\d+\b', filename)
    return matches

def merge_pdfs(pdf_list, output_path):
    pdf_merger = PyPDF2.PdfMerger()

    for pdf in pdf_list:
        pdf_merger.append(pdf)

    with open(output_path, 'wb') as output_pdf:
        pdf_merger.write(output_pdf)

def convert_pdf(file_path, final_path):
    return convert(file_path, final_path)["output"]
def main():
    folder_docs, folder_maps = choose_folder()
    if not folder_docs:
        print("Folder docs not selected.")
        return
    if not folder_maps:
        print("Folder maps not selected.")
        return

    pdfs_len = 0;
    for docname in os.listdir(folder_docs):
        doc_path = os.path.join(folder_docs, docname)
        plot_numbers = extract_numbers(docname)
        maps_to_merge = []
        for mapname in os.listdir(folder_maps):
            map_plot_number = extract_numbers(mapname)[0]
            if map_plot_number in plot_numbers:
                 map_path = os.path.join(folder_maps, mapname)
                 maps_to_merge.append(map_path)

        target_folder = create_new_folder(folder_docs)
        pdfname = convert_pdf(doc_path, target_folder)
        maps_to_merge.insert(0, pdfname)
        pdfname = os.path.join(folder_docs, os.path.basename(pdfname))

        merge_pdfs(maps_to_merge, pdfname)
        pdfs_len =+ 1

    print("All pdfs")
    print(pdfs_len)
if __name__ == "__main__":
    main()
