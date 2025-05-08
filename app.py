from flask import Flask, render_template, request, jsonify
from datetime import datetime
import csv
import os

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

@app.route('/')
def index():
    status = carregar_status()
    return render_template('index.html', status=status)

@app.route('/marcar', methods=['POST'])
def marcar():
    data = request.json
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
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

    guardar_status("ENTRADA" if entrada_sortida == "ENTRADA" else "SORTIDA")
    return jsonify({"ok": True})

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

