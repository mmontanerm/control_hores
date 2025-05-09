from flask import Flask, render_template, request, jsonify, send_file
from datetime import datetime
import os
import pytz
import psycopg2
from collections import deque
import csv
from io import StringIO

app = Flask(__name__)

DATABASE_URL = os.environ['DATABASE_URL']


def connect_db():
    return psycopg2.connect(DATABASE_URL, sslmode='require')


def carregar_ultim_estat():
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("SELECT tipus FROM registre ORDER BY data_hora DESC LIMIT 1")
    resultat = cur.fetchone()
    cur.close()
    conn.close()
    if not resultat:
        return "ENTRADA"
    return "SORTIDA" if resultat[0].upper() == "ENTRADA" else "ENTRADA"


def calcular_import_mes(mes, any_actual):
    TARIFA_AIRA = 46.5
    TARIFA_REMOT = 37.5

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT data_hora, tipus, lloc FROM registre 
        WHERE EXTRACT(MONTH FROM data_hora) = %s AND EXTRACT(YEAR FROM data_hora) = %s
        ORDER BY data_hora ASC
    """, (mes, any_actual))
    files = cur.fetchall()
    cur.close()
    conn.close()

    entrades = deque()
    total_aira = 0.0
    total_remot = 0.0

    for data_hora, tipus, lloc in files:
        if tipus == "ENTRADA":
            entrades.append((data_hora, lloc.strip().upper() if lloc else ""))
        elif tipus == "SORTIDA" and entrades:
            entrada_hora, lloc = entrades.popleft()
            durada = (data_hora - entrada_hora).total_seconds() / 3600.0
            if lloc == "AIRA":
                total_aira += durada * TARIFA_AIRA
            elif lloc == "REMOT":
                total_remot += durada * TARIFA_REMOT

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

    of = data.get('of') if entrada_sortida == "ENTRADA" else ""
    lloc = data.get('lloc') if entrada_sortida == "ENTRADA" else ""
    comentaris = data.get('comentaris') if entrada_sortida == "ENTRADA" else ""

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO registre (data_hora, tipus, of, lloc, comentaris)
        VALUES (%s, %s, %s, %s, %s)
    """, (timestamp, entrada_sortida, of, lloc, comentaris))
    conn.commit()
    cur.close()
    conn.close()

    return jsonify({"ok": True})


@app.route('/descarregar')
def descarregar_csv():
    conn = connect_db()
    cur = conn.cursor()
    cur.execute("SELECT data_hora, tipus, of, lloc, comentaris FROM registre ORDER BY data_hora ASC")
    registres = cur.fetchall()
    cur.close()
    conn.close()

    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(["Data i hora", "Tipus", "OF", "Lloc de feina", "Comentaris"])
    for fila in registres:
        writer.writerow(fila)

    output.seek(0)
    return send_file(output, as_attachment=True, download_name="registre.csv", mimetype="text/csv")


@app.route('/registre')
def veure_registre():
    data_inici = request.args.get("data_inici", "")
    data_fi = request.args.get("data_fi", "")

    conn = connect_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT data_hora, tipus, of, lloc, comentaris
        FROM registre
        ORDER BY data_hora ASC
    """)
    files = cur.fetchall()
    cur.close()
    conn.close()

    registres = []
    entrada = None

    for fila in files:
        data_hora, tipus, of, lloc, comentaris = fila
        if tipus == "ENTRADA":
            entrada = {
                "data": data_hora,
                "of": of,
                "lloc": lloc,
                "comentaris": comentaris
            }
        elif tipus == "SORTIDA" and entrada:
            if data_inici:
                dt_ini = datetime.strptime(data_inici, "%Y-%m-%d")
                if data_hora < dt_ini:
                    continue
            if data_fi:
                dt_fi = datetime.strptime(data_fi, "%Y-%m-%d")
                if data_hora > dt_fi:
                    continue

            durada = round((data_hora - entrada["data"]).total_seconds() / 3600.0, 2)

            registres.append({
                "data_inici": entrada["data"].strftime('%Y-%m-%d %H:%M:%S'),
                "data_final": data_hora.strftime('%Y-%m-%d %H:%M:%S'),
                "of": entrada["of"],
                "lloc": entrada["lloc"],
                "comentaris": entrada["comentaris"],
                "hores": durada
            })
            entrada = None

    return render_template('registre.html', registres=registres, data_inici=data_inici, data_fi=data_fi)

@app.route('/pujar', methods=['POST'])
def pujar():
    if 'fitxer' not in request.files:
        return "No s'ha enviat cap fitxer", 400

    fitxer = request.files['fitxer']
    if fitxer.filename == '':
        return "Fitxer buit", 400

    try:
        stream = StringIO(fitxer.stream.read().decode("UTF-8"))
        reader = csv.DictReader(stream)

        registres_nous = []

        for fila in reader:
            data_hora = fila["Data i hora"]
            tipus = fila["Tipus"]
            of = fila["OF"]
            lloc = fila["Lloc de feina"]
            comentaris = fila["Comentaris"]

            # Conversió segura de data
            try:
                data_obj = datetime.strptime(data_hora, "%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue  # salta línies mal formades

            registres_nous.append((data_obj, tipus, of, lloc, comentaris))

        if not registres_nous:
            return "No s'han trobat registres vàlids", 400

        conn = connect_db()
        cur = conn.cursor()

        for reg in registres_nous:
            cur.execute("""
                SELECT 1 FROM registre
                WHERE data_hora = %s AND tipus = %s AND of = %s AND lloc = %s AND comentaris = %s
            """, reg)

            if not cur.fetchone():
                cur.execute("""
                    INSERT INTO registre (data_hora, tipus, of, lloc, comentaris)
                    VALUES (%s, %s, %s, %s, %s)
                """, reg)

        conn.commit()
        cur.close()
        conn.close()

        return "Fitxer importat correctament (sense duplicats)", 200

    except Exception as e:
        return f"Error llegint el fitxer: {str(e)}", 500



if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)


