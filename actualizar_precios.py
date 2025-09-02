import os
import sqlite3
import pandas as pd
import csv

# Paths
EXCEL_DIR = r"D:\herloc.programacion\lector_facturas_pdf\files_repo\comprobados"
DB_PROYECTO = r"D:\herloc.programacion\negocio_26\data\app.sqlite"
DB_provider = r"D:\herloc.programacion\negocio_26\data\provider.sqlite"

# Logs
LOG_OK = "log_ok.csv"
LOG_ERR = "log_errores.csv"

# Margen en %
MARGEN = 20

def calcular_pvp(coste, iva, margen=MARGEN):
	if iva == 4:
		base = coste * 1.045
	elif iva == 21:
		base = coste * 1.262
	else:
		base = coste
	return round(base / ((100 - margen) / 100), 2)

def inicializar_logs():
	with open(LOG_OK, "w", newline="", encoding="utf-8") as f:
		writer = csv.writer(f)
		writer.writerow(["Operacion", "Codigo", "Referencia", "Descripcion", "Coste", "PVP"])
	with open(LOG_ERR, "w", newline="", encoding="utf-8") as f:
		writer = csv.writer(f)
		writer.writerow(["Codigo", "Referencia", "Descripcion", "Coste", "IVA", "Motivo"])

def log_ok(operacion, codigo, referencia, descripcion, coste, pvp):
	with open(LOG_OK, "a", newline="", encoding="utf-8") as f:
		writer = csv.writer(f)
		writer.writerow([operacion, codigo, referencia, descripcion, coste, pvp])

def log_err(codigo, referencia, descripcion, coste, iva, motivo):
	with open(LOG_ERR, "a", newline="", encoding="utf-8") as f:
		writer = csv.writer(f)
		writer.writerow([codigo, referencia, descripcion, coste, iva, motivo])

def procesar_excel(path_excel, conn_proyecto, conn_provider):
	df = pd.read_excel(path_excel)

	for _, row in df.iterrows():
		try:
			codigo = row["Código"]
			referencia = row["Referencia"]
			descripcion = row["Descripción"]
			coste = float(row["Precio"])
			iva_val = row["IVA"]

			if pd.isna(iva_val):
				print(f"⚠ Fila ignorada (sin IVA): {codigo} - {referencia} - {descripcion}")
				log_err(codigo, referencia, descripcion, coste, None, "Fila sin IVA en Excel")
				continue

			iva = int(iva_val)

		except Exception as e:
			print(f"⚠ Error leyendo fila: {e}")
			log_err(row.get("Código", ""), row.get("Referencia", ""), row.get("Descripción", ""), row.get("Precio", ""), row.get("IVA", ""), f"Error leyendo fila: {e}")
			continue

		# Calcular PVP
		pvp = calcular_pvp(coste, iva)
		
		# Buscar en la BD del proyecto
		cur = conn_proyecto.cursor()
		cur.execute("SELECT id FROM products WHERE codigo = ? AND referencia = ?", (codigo, referencia))
		existe = cur.fetchone()

		if existe:
			# Actualizar precio
			cur.execute("""
				UPDATE products
				SET pvcoste = ?, pvp = ?
				WHERE id = ?
			""", (coste, pvp, existe[0]))
			conn_proyecto.commit()
			print(f"✓ Actualizado en proyecto: {codigo} - {referencia}")
			log_ok("UPDATE", codigo, referencia, descripcion, coste, pvp)

		else:
			# Buscar en provider
			cur_prov = conn_provider.cursor()
			cur_prov.execute("SELECT * FROM products WHERE codigo = ? AND referencia = ?", (codigo, referencia))
			prov_articulo = cur_prov.fetchone()

			if prov_articulo:
				# Insertar en proyecto
				cur.execute("""
					INSERT INTO products
					(uid, codigo, referencia, subcategoria, descripcion, neto, iva,
					 ean_crc, ean_unidad, ean_unitario, ean_envase, ean_embalaje,
					 pvp, pvcoste, stock)
					VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
				""", (
					prov_articulo[1],  # uid
					prov_articulo[2],  # codigo
					prov_articulo[3],  # referencia
					prov_articulo[4],  # subcategoria
					prov_articulo[5],  # descripcion
					prov_articulo[6],  # neto
					prov_articulo[7],  # iva
					prov_articulo[8],  # ean_crc
					prov_articulo[9],  # ean_unidad
					prov_articulo[10], # ean_unitario
					prov_articulo[11], # ean_envase
					prov_articulo[12], # ean_embalaje
					pvp,               # pvp recalculado
					coste,             # pvcoste desde factura
					0                  # stock inicial
				))
				conn_proyecto.commit()
				print(f"+ Insertado desde provider: {codigo} - {referencia}")
				log_ok("INSERT", codigo, referencia, descripcion, coste, pvp)
			else:
				print(f"⚠ No encontrado en provider: {codigo} - {referencia}")
				log_err(codigo, referencia, descripcion, coste, iva, "No encontrado en provider")

def main():
	inicializar_logs()

	conn_proyecto = sqlite3.connect(DB_PROYECTO)
	conn_provider = sqlite3.connect(DB_provider)

	for file in os.listdir(EXCEL_DIR):
		if file.endswith(".xlsx"):
			path_excel = os.path.join(EXCEL_DIR, file)
			print(f"Procesando {path_excel}...")
			procesar_excel(path_excel, conn_proyecto, conn_provider)

	conn_proyecto.close()
	conn_provider.close()
	print("✔ Sincronización terminada")
	print(f"Log de éxitos: {LOG_OK}")
	print(f"Log de errores: {LOG_ERR}")

if __name__ == "__main__":
	main()
