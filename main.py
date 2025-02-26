import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os
import shutil
import webbrowser

UPLOAD_FOLDER = "uploads"
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def convert_excel_to_html(filepath, title):
    try:
        df = pd.read_excel(filepath)
        df = df.fillna("")

        filename = os.path.basename(filepath)
        save_path = os.path.join(UPLOAD_FOLDER, filename)
        shutil.copy(filepath, save_path)  
        html_table = df.to_html(index=False)

        html_content = f"""
        <!DOCTYPE html>
        <html lang="de">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>{title}</title>
            <link rel="stylesheet" href="styles.css">
        </head>
        <body>
            <div class="container">
                <h1>{title}</h1>
                {html_table}
                <br>
                <a href="{save_path}" download class="download-btn">Excel-Datei herunterladen</a>
            </div>
        </body>
        </html>
        """

        with open("index.html", "w", encoding="utf-8") as file:
            file.write(html_content)

        upload_status.config(text="Upload erfolgreich!", fg="green")
        webbrowser.open("index.html")
    except Exception as e:
        upload_status.config(text=f"Fehler beim Upload: {str(e)}", fg="red")

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Dateien", "*.xlsx;*.xls")])
    if file_path:
        title = title_entry.get()
        if not title:
            title = "Excel-Tabelle"
        convert_excel_to_html(file_path, title)

root = tk.Tk()
root.title("ExcelToWeb")
root.geometry("1000x600")
root.configure(bg="#121212")

frame = tk.Frame(root, bg="#1E1E1E", padx=30, pady=30, relief="raised", bd=2)
frame.pack(pady=20, padx=20, fill="both", expand=True)

title_label = tk.Label(frame, text="Gib den Titel der HTML-Seite ein:", bg="#1E1E1E", fg="white", font=("Arial", 16, "bold"))
title_label.pack(pady=10)

title_entry = tk.Entry(frame, font=("Arial", 14), width=40, relief="solid", bd=2, bg="#333333", fg="white", insertbackground="white")
title_entry.pack(pady=10, ipady=5)

upload_button = tk.Button(frame, text="Excel-Datei auswählen", command=upload_file, font=("Arial", 14, "bold"), bg="#6200EE", fg="white", padx=30, pady=15, relief="flat", bd=0)
upload_button.pack(pady=20)

upload_status = tk.Label(frame, text="", bg="#1E1E1E", fg="white", font=("Arial", 14, "bold"))
upload_status.pack(pady=10)

root.mainloop()
