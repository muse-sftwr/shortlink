import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import requests
import time

# Funzione per accorciare un singolo URL
def shorten_url(url):
    try:
        url = url.strip()
        if not re.match(r'^https?://', url):
            return "INVALID_URL"
        time.sleep(0.2)  # pausa tra le richieste per evitare blocchi IP
        resp = requests.get(f"http://tinyurl.com/api-create.php?url={url}")
        if resp.status_code == 200:
            return resp.text
        else:
            return f"HTTP_ERROR_{resp.status_code}"
    except Exception as e:
        return f"EXCEPTION: {str(e)}"

# Normalizza i nomi colonna per cercare "OriginalLink" o simili
def find_column(df, target):
    target_clean = re.sub(r'[^a-zA-Z]', '', target).lower()
    for col in df.columns:
        col_clean = re.sub(r'[^a-zA-Z]', '', col).lower()
        if col_clean == target_clean:
            return col
    return None

# Funzione principale
def process_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, engine='openpyxl')

        # Trova colonne flessibili
        original_col = find_column(df, 'OriginalLink')
        short_col = find_column(df, 'ShortLink')

        if not original_col:
            messagebox.showerror("Errore", "Colonna 'OriginalLink' non trovata (accetta anche 'Original Link', 'original_link', ecc.)")
            return

        # Crea colonna shortlink se non esiste
        if not short_col:
            short_col = 'ShortLink'
            df[short_col] = ''

        for idx, row in df.iterrows():
            original_data = row[original_col]
            if pd.isna(original_data):
                continue

            urls = [url.strip() for url in str(original_data).split(',') if url.strip()]
            short_urls = [shorten_url(url) for url in urls]
            df.at[idx, short_col] = ', '.join(short_urls)

        # Salva direttamente sul file originale
        df.to_excel(file_path, index=False, engine='openpyxl')
        messagebox.showinfo("Fatto", f"Colonna '{short_col}' aggiornata nel file originale:\n{file_path}")

    except Exception as e:
        messagebox.showerror("Errore", f"Errore durante l'elaborazione:\n{str(e)}")

# GUI
root = tk.Tk()
root.title("Shortener Excel Tool")
root.geometry("400x200")

label = tk.Label(root, text="Accorciatore di URL da Excel", font=("Arial", 14))
label.pack(pady=20)

button = tk.Button(root, text="Scegli file Excel e genera ShortLink", command=process_excel, font=("Arial", 12), bg="#4CAF50", fg="white")
button.pack(pady=20)

root.mainloop()
