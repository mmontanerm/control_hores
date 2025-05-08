import os
import re
import pandas as pd
from pathlib import Path
from tkinter import simpledialog, filedialog, Tk, Toplevel, Label
from tkinter.ttk import Progressbar
from PyPDF2 import PdfReader

# 🔒 Comprovació de la llibreria pywin32
try:
    import win32com.client  # Per accedir a Inventor via COM
except ImportError:
    print("❌ El mòdul 'pywin32' no està instal·lat.")
    print("Pots instal·lar-lo amb: pip install pywin32")
    exit()

# Funció per llegir PDFs
def cerca_en_pdf(fitxer, codi):
    try:
        reader = PdfReader(str(fitxer))
        for page in reader.pages:
            text = page.extract_text()
            if text and codi in text:
                return True
    except Exception:
        pass
    return False

# Funció per llegir fitxers de text
def cerca_dins_fitxer(fitxer, codi):
    try:
        if fitxer.suffix.lower() == ".pdf":
            return cerca_en_pdf(fitxer, codi)
        with open(fitxer, 'r', encoding='utf-8', errors='ignore') as f:
            contingut = f.read()
            if codi in contingut:
                return True
    except Exception:
        pass
    return False

# ✅ Funció per obtenir referències d'un fitxer .IAM utilitzant Inventor
def obtenir_referencies_iam(path_iam, app_inventor):
    referencies = []
    try:
        doc = app_inventor.Documents.Open(str(path_iam))
        for comp in doc.ComponentDefinition.Occurrences:
            fitxer_ref = comp.Definition.Document.FullFileName
            referencies.append(fitxer_ref)
        doc.Close(True)
    except Exception as e:
        print(f"⚠️ Error llegint IAM: {path_iam.name}")
    return referencies

# Inicia la interfície gràfica
root = Tk()
root.withdraw()

# Selecciona la carpeta
carpeta_seleccionada = filedialog.askdirectory(title="Selecciona la carpeta principal on buscar")
if not carpeta_seleccionada:
    print("No s'ha seleccionat cap carpeta. Sortint...")
    exit()
CARPETA_ARREL = Path(carpeta_seleccionada)

# Introdueix el codi
codi_buscat = simpledialog.askstring("Cerca de fitxers", "Introdueix el codi de màquina (ex: 23835-1000):")
if not codi_buscat:
    print("No s'ha introduït cap codi. Sortint...")
    exit()

# Classificació segons extensió
TIPUS_FITXERS = {
    "3D_Inventor": [".ipt", ".iam"],
    "2D_Planols": [".dwg", ".dxf", ".pdf"],
    "Llistat_parts": [".xls", ".xlsx", ".csv"],
    "Documentacio": [".txt", ".docx", ".doc", ".rtf"],
}

# Crea la finestra de progrés
progress_win = Toplevel()
progress_win.title("Cercant fitxers...")
progress_win.geometry("400x100")
progress_win.resizable(False, False)

label = Label(progress_win, text="S'estan cercant fitxers. Espera si us plau...")
label.pack(pady=10)

progress = Progressbar(progress_win, orient="horizontal", length=350, mode="determinate")
progress.pack(pady=5)
progress.update()

# Inicialitza la cerca
fitxers_a_analitzar = list(CARPETA_ARREL.rglob("*.*"))
progress["maximum"] = len(fitxers_a_analitzar)
fitxers_trobats = []
fitxers_afegits = set()

# Inicia Inventor
try:
    app = win32com.client.Dispatch("Inventor.Application")
    app.Visible = False
except Exception:
    print("❌ No s'ha pogut obrir Autodesk Inventor via COM.")
    progress_win.destroy()
    exit()

# Cerca principal
referencies_detectades = []

for i, fitxer in enumerate(fitxers_a_analitzar):
    ruta_completa = fitxer.resolve()
    nom = fitxer.name
    extensio = fitxer.suffix.lower()

    # Condició per afegir-lo a la llista
    afegir = False

    if codi_buscat in nom or cerca_dins_fitxer(fitxer, codi_buscat):
        afegir = True

    # Si és .iam, obtenir referències
    referencies = []
    if extensio == ".iam":
        referencies = obtenir_referencies_iam(fitxer, app)
        referencies_detectades.extend(referencies)

    if afegir and str(ruta_completa) not in fitxers_afegits:
        tipus = "Altres"
        for nom_tipus, extensions in TIPUS_FITXERS.items():
            if extensio in extensions:
                tipus = nom_tipus
                break
        fitxers_trobats.append({
            "Nom_fitxer": nom,
            "Ruta_completa": str(ruta_completa),
            "Tipus": tipus,
            "Extensio": extensio,
        })
        fitxers_afegits.add(str(ruta_completa))

    progress["value"] = i + 1
    progress.update()

# Tanca Inventor
app.Quit()
progress_win.destroy()

# 🔁 Buscar els fitxers referenciats si existeixen físicament
for ruta_ref in referencies_detectades:
    fitxer_ref = Path(ruta_ref)
    if fitxer_ref.exists() and str(fitxer_ref.resolve()) not in fitxers_afegits:
        extensio = fitxer_ref.suffix.lower()
        tipus = "3D_Inventor" if extensio in [".iam", ".ipt"] else "Altres"
        fitxers_trobats.append({
            "Nom_fitxer": fitxer_ref.name,
            "Ruta_completa": str(fitxer_ref.resolve()),
            "Tipus": tipus,
            "Extensio": extensio,
        })
        fitxers_afegits.add(str(fitxer_ref.resolve()))

# Exportar a Excel
if fitxers_trobats:
    df = pd.DataFrame(sorted(fitxers_trobats, key=lambda x: (x["Tipus"], x["Nom_fitxer"])))
    nom_excel = "fitxers_" + codi_buscat.replace("-", "_") + ".xlsx"
    fitxer_excel = CARPETA_ARREL / nom_excel
    df.to_excel(fitxer_excel, index=False)
    print("\n✅ Fitxer Excel generat:", fitxer_excel)
else:
    print("\nℹ️ No s'han trobat fitxers amb el codi", codi_buscat)
