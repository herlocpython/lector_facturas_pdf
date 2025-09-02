import sqlite3

DBS = [
    r"D:\herloc.programacion\negocio_26\data\app.sqlite",
    r"D:\herloc.programacion\negocio_26\data\proveedor.sqlite"
]

for db_path in DBS:
    print(f"\n📂 Base de datos: {db_path}")
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # Listar todas las tablas
    cur.execute("SELECT name FROM sqlite_master WHERE type='table';")
    tablas = cur.fetchall()

    if not tablas:
        print("  ⚠ No hay tablas en esta base de datos.")
        continue

    for (tabla,) in tablas:
        print(f"  ▸ Tabla: {tabla}")
        # Listar columnas de la tabla
        cur.execute(f"PRAGMA table_info({tabla});")
        columnas = cur.fetchall()
        for col in columnas:
            print(f"     - {col[1]} ({col[2]})")

    conn.close()
