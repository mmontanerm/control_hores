import os

# Defineix la carpeta on vols esborrar els fitxers
carpeta = "Z:/3D/2021"  # Substitueix per la ruta correcta

# Recórrer totes les subcarpetes i fitxers
for arrel, subcarpetes, fitxers in os.walk(carpeta):
    for fitxer in fitxers:
        if fitxer.endswith((".STL", ".stl", ".3MF", ".3mf", ".pdf", ".PDF", ".stp", ".STP", ".STEP", ".step")):  # Comprova si és un STL o STEP
            fitxer_path = os.path.join(arrel, fitxer)
            try:
                os.remove(fitxer_path)
                print("Esborrat:", fitxer_path)
            except Exception as e:
                print("No s'ha pogut esborrar {}: {}".format(fitxer_path, e))

print("Procés completat.")
