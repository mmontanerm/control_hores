import os

# Ruta del directori on són els fitxers
directori = r'\\SAIRA\empresa\AIRA\Dept. Producción\12- HYGIENE\18_510011521\02- 3D\Amb referencies HY\10 Konstruktionszeichnungen'  # canvia-ho per la ruta real

os.chdir(directori)

for nom_fitxer in os.listdir():
    if nom_fitxer.startswith("HY-"):
        continue  # Ja té el prefix correcte, no cal fer res
    elif nom_fitxer.startswith("DB1ENG-"):
        nou_nom = "HY-" + nom_fitxer[len("DB1ENG-"):]  # substitueix DB1ENG per HY
    else:
        nou_nom = "HY-" + nom_fitxer  # només afegeix el prefix

    print(f"Renombrant {nom_fitxer} -> {nou_nom}")
    os.rename(nom_fitxer, nou_nom)
