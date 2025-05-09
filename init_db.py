import os
import psycopg2

conn = psycopg2.connect(os.environ['DATABASE_URL'], sslmode='require')
cur = conn.cursor()

cur.execute("""
CREATE TABLE IF NOT EXISTS registre (
    id SERIAL PRIMARY KEY,
    data_hora TIMESTAMP,
    tipus VARCHAR(10),
    of TEXT,
    lloc TEXT,
    comentaris TEXT
);
""")

conn.commit()
cur.close()
conn.close()
print("Taula creada correctament.")
