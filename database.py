import sqlite3
import os

def conectar():
    # Crear carpeta data si no existe
    if not os.path.exists("data"):
        os.makedirs("data")


    conn = sqlite3.connect("data/profesores.db")
    return conn

def crear_tablas():
    conn = conectar()
    c = conn.cursor()

    c.execute("""
    CREATE TABLE IF NOT EXISTS profesores (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nombre TEXT NOT NULL
    )
    """)


    conn.commit()
    conn.close()