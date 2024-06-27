import os
import pandas as pd
import customtkinter as ctk
from customtkinter import filedialog

import tkinter as tk

documentsDirPath = os.path.join(os.path.expanduser("~"), "Documents")
root = ctk.CTk()
root.withdraw()

folder_path = filedialog.askdirectory(title="Wybierz folder")

if not folder_path:
    print("Nie wybrano folderu.")
    exit()

folders_list = os.listdir(folder_path)

df = pd.DataFrame(folders_list, columns=["Nazwa folderu"])

excel_file_path = os.path.join(folder_path, "nazwy_folderow.xlsx")
df.to_excel(excel_file_path, index=False)

print(f"Plik Excel został utworzony pomyślnie w ścieżce: {excel_file_path}")
