from flask import Flask, render_template, request, jsonify, send_file
from datetime import datetime
import csv
import os
import pytz
from collections import deque
from flask import request

app = Flask(__name__)
REGISTRE_PATH = 'registre.csv'
STATUS_PATH = 'status.txt'

def carregar_status():
    if not os.path.exists(STATUS_PATH):
        with open(STATUS_PATH, 'w') as f:
            f.write("ENTRADA")
    with open(STATUS_PATH, 'r') as f:
        return f.read().strip()

def guardar_status(estat):
    with open(STATUS_PATH, 'w') as f:
        f.write(estat)

def carregar_ultim_estat():
    if not os.path.exists(REGISTRE_PATH):
        return "ENTRADA"
    with open(REGISTRE_PATH, 'r', encoding='utf-8') as f:
        linies = f.readlines()
    if len(linies) < 2:
        return "ENTRADA"
    ultima_linia = linies[-1].strip().split(',')
    ultim_tipus = ultima_linia[1].strip().upper()
    return "SORTIDA" if ultim_tipus == "ENTRADA" else "ENTRADA"

def calcular_import_mes(mes, any_actual):
    if not os.path.exists(REGISTRE_PATH):
        return 0.0

    TARIFA_AIRA = 46.5
    TARIFA_REMOT = 37.5

    with open(REGISTRE_PATH, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        entrades = deque()
        total_aira = 0.0
        total_remot = 0.0

        for row in reader:
            data_str = row["Data i hora"]
            tipus = row["Tipus"]
            lloc = row["Lloc de feina"]

            if not data_str:
                continue

            data_hora = datetime.strptime(data_str, "%Y-%m-%d %H:%M:%S")
            if data_hora.month != mes or data_hora.year != any_actual:
                continue

            if tipus == "ENTRADA":
                entrades.append((data_hora, lloc.strip().upper()))
            elif tipus == "SORTIDA" and entrades:
                entrada_hora, lloc = entrades.popleft()
                durada = (data_hora - entrada_hora).total_seconds() / 3600.0
                if lloc == "AIRA":
                    total_aira += durada * 47.5
                elif lloc == "REMOT":
                    total_remot += durada * 37.5

    return round(total_aira + total_remot, 2)

@app.route('/')
def index():
    status = carregar_ultim_estat()
    zona = pytz.timezone("Europe/Madrid")
    ara = datetime.now(zona)
    import_mes = calcular_import_mes(ara.month, ara.year)
    return render_template('index.html', status=status, import_mes=import_mes)

@app.route('/marcar', methods=['POST'])
def marcar():
    data = request.json
    zona = pytz.timezone("Europe/Madrid")
    timestamp = datetime.now(zona).strftime('%Y-%m-%d %H:%M:%S')
    entrada_sortida = data.get('tipus')

    if entrada_sortida == "ENTRADA":
        fila = [timestamp, "ENTRADA", data.get('of'), data.get('lloc'), data.get('comentaris')]
    else:
        fila = [timestamp, "SORTIDA", "", "", ""]

    nou_fitxer = not os.path.exists(REGISTRE_PATH)
    with open(REGISTRE_PATH, 'a', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if nou_fitxer:
            writer.writerow(["Data i hora", "Tipus", "OF", "Lloc de feina", "Comentaris"])
        writer.writerow(fila)
        f.flush()                # <-- Afegeix això
        os.fsync(f.fileno())     # <-- I això també


    guardar_status("ENTRADA" if entrada_sortida == "ENTRADA" else "SORTIDA")
    return jsonify({"ok": True})

@app.route('/descarregar')
def descarregar_csv():
    if os.path.exists(REGISTRE_PATH):
        return send_file(REGISTRE_PATH, as_attachment=True)
    return "No hi ha cap registre disponible", 404

@app.route('/pujar', methods=['POST'])
def pujar():
    if 'fitxer' not in request.files:
        return "No s'ha enviat cap fitxer", 400

    fitxer = request.files['fitxer']
    if fitxer.filename == '':
        return "Fitxer buit", 400

    if fitxer:
        fitxer.save(REGISTRE_PATH)
        return "Fitxer pujat correctament", 200

from flask import request

@app.route('/registre')
def veure_registre():
    if not os.path.exists(REGISTRE_PATH):
        return render_template('registre.html', registres=[], data_inici="", data_fi="")

    data_inici = request.args.get("data_inici", "")
    data_fi = request.args.get("data_fi", "")

    with open(REGISTRE_PATH, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        linies = list(reader)

    registres = []
    entrada = None

    for fila in linies:
        if fila["Tipus"] == "ENTRADA":
            entrada = {
                "data": fila["Data i hora"],
                "of": fila["OF"],
                "lloc": fila["Lloc de feina"],
                "comentaris": fila["Comentaris"]
            }
        elif fila["Tipus"] == "SORTIDA" and entrada:
            data_entrada = datetime.strptime(entrada["data"], "%Y-%m-%d %H:%M:%S")
            data_sortida = datetime.strptime(fila["Data i hora"], "%Y-%m-%d %H:%M:%S")
            durada = round((data_sortida - data_entrada).total_seconds() / 3600.0, 2)

            # Filtrat per dates si cal
            if data_inici:
                dt_ini = datetime.strptime(data_inici, "%Y-%m-%d")
                if data_sortida < dt_ini:
                    continue
            if data_fi:
                dt_fi = datetime.strptime(data_fi, "%Y-%m-%d")
                if data_sortida > dt_fi:
                    continue

            registres.append({
                "data_inici": entrada["data"],
                "data_final": fila["Data i hora"],
                "of": entrada["of"],
                "lloc": entrada["lloc"],
                "comentaris": entrada["comentaris"],
                "hores": durada
            })
            entrada = None

    return render_template('registre.html', registres=registres, data_inici=data_inici, data_fi=data_fi)


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

