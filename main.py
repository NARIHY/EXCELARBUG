import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import getpass

def extraire_feuilles():
    fichier_excel = filedialog.askopenfilename(
        title="Sélectionner un fichier Excel",
        filetypes=[("Fichiers Excel", "*.xlsx")]
    )

    if not fichier_excel:
        return

    # Récupération user et date
    user = getpass.getuser()
    date_du_jour = datetime.now().strftime("%Y-%m-%d")

    # Chemin Documents
    documents = os.path.join(os.path.expanduser("~"), "Documents")

    dossier_sortie = os.path.join(
        documents,
        "bugExcel",
        user,
        date_du_jour,
        "fichier-excel-extraite"
    )

    os.makedirs(dossier_sortie, exist_ok=True)

    try:
        xls = pd.ExcelFile(fichier_excel)

        for sheet in xls.sheet_names:
            df = pd.read_excel(fichier_excel, sheet_name=sheet)
            fichier_sortie = os.path.join(dossier_sortie, f"{sheet}.xlsx")
            df.to_excel(fichier_sortie, index=False)

        messagebox.showinfo(
            "Succès",
            f"{len(xls.sheet_names)} feuilles extraites dans :\n{dossier_sortie}"
        )

    except Exception as e:
        messagebox.showerror("Erreur", str(e))


# Interface graphique
root = tk.Tk()
root.title("Extraction Excel – bugExcel")
root.geometry("420x200")

label = tk.Label(
    root,
    text="Extraction des feuilles Excel\nVers Documents/bugExcel",
    font=("Arial", 12)
)
label.pack(pady=20)

btn = tk.Button(
    root,
    text="Sélectionner le fichier Excel",
    command=extraire_feuilles,
    font=("Arial", 10),
    bg="#4CAF50",
    fg="white",
    padx=10,
    pady=5
)
btn.pack()

root.mainloop()
