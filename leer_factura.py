import pdfplumber
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
import os

def extract_invoice_data(pdf_path):
    """
    Extrae datos de productos de facturas PDF con el formato específico de Liderpapel
    """
    all_data = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extraer texto de la página
            text = page.extract_text()
            lines = text.split('\n')
            
            in_products_section = False
            current_product = None
            
            for i, line in enumerate(lines):
                # Detectar inicio de la sección de productos
                if 'Código Referencia Descripción' in line:
                    in_products_section = True
                    continue
                
                # Detectar fin de la sección de productos
                if any(end_marker in line for end_marker in ['Empresa adherida', 'Forma de Pago', 'CONDICIONES DE VENTA']):
                    in_products_section = False
                    if current_product:
                        all_data.append(current_product)
                        current_product = None
                    continue
                
                if in_products_section and line.strip():
                    # Patrón modificado: permite cualquier carácter en la referencia excepto espacios
                    pattern = r'^(\d+)\s+([^\s]+)\s+(.+?)\s+(\d+)\s+([\d,]+)\s+([\d,]+)\s+(\d+)$'
                    match = re.match(pattern, line.strip())
                    
                    if match:
                        # Si hay un producto actual pendiente, guardarlo
                        if current_product:
                            all_data.append(current_product)
                        
                        codigo, referencia, descripcion, cantidad, precio, importe, iva = match.groups()
                        current_product = {
                            'Código': codigo,
                            'Referencia': referencia.strip(),
                            'Descripción': descripcion.strip(),
                            'Cantidad': cantidad,
                            'Precio': precio,
                            'IVA': iva
                        }
                    elif current_product:
                        # Si ya tenemos un producto actual, podría ser continuación de descripción
                        # Verificar si la línea parece ser datos numéricos (no continuación)
                        if not re.search(r'\d+[\s,]+\d+[\s,]+\d+$', line.strip()):
                            current_product['Descripción'] += " " + line.strip()
                        else:
                            # Si parece ser una nueva línea de producto pero no matcheó el patrón completo
                            all_data.append(current_product)
                            current_product = None
            
            # Añadir el último producto si existe
            if current_product:
                all_data.append(current_product)
    
    # Limpiar y convertir los datos
    cleaned_data = []
    for item in all_data:
        try:
            # Limpiar descripción de caracteres extraños
            descripcion = re.sub(r'\s+', ' ', item['Descripción']).strip()
            
            cleaned_data.append({
                'Código': item['Código'],
                'Referencia': item['Referencia'],
                'Descripción': descripcion,
                'Cantidad': int(item['Cantidad']),
                'Precio': float(item['Precio'].replace(',', '.')),
                'IVA': int(item['IVA'])
            })
        except (ValueError, KeyError) as e:
            print(f"Error procesando item: {item}, error: {e}")
            continue
    
    return pd.DataFrame(cleaned_data)

def clean_descriptions(df):
    """
    Limpia y completa descripciones basado en patrones comunes
    """
    # Diccionario de mapeo para completar descripciones conocidas
    description_map = {
        'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL': {
            'LP541': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC AZUL DIA PAGINA',
            'LP542': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC NEGRO DIA PAGINA',
            'LP543': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC ROSA DIA PAGINA',
            'LP544': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC ROJO DIA PAGINA',
            'LP546': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC VERDE MENTA DIA PAGINA'
        },
        'BOLIGRAFO': {
            '8373602': 'BOLIGRAFO BIC CRISTAL ORIGINAL TINTA AZUL',
            'KF18625': 'BOLIGRAFO Q-CONNECT RETRACTIL BORRABLE 0,7 MM COLOR AZUL',
            'KF18626': 'BOLIGRAFO Q-CONNECT RETRACTIL BORRABLE 0,7 MM COLOR ROJO'
        }
    }
    
    for idx, row in df.iterrows():
        desc = row['Descripción']
        ref = row['Referencia']
        
        # Completar descripciones de agendas
        if 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL' in desc:
            if ref in description_map['AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL']:
                df.at[idx, 'Descripción'] = description_map['AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL'][ref]
        
        # Completar descripciones de bolígrafos
        elif 'BOLIGRAFO' in desc:
            if ref in description_map['BOLIGRAFO']:
                df.at[idx, 'Descripción'] = description_map['BOLIGRAFO'][ref]
    
    return df

def save_to_excel(df, excel_path, pdf_filename=None):
    """
    Guarda el DataFrame en un archivo Excel con formato profesional
    """
    # Crear un nuevo libro de Excel
    wb = Workbook()
    ws = wb.active
    ws.title = "Productos Factura"
    
    # Obtener el nombre del archivo PDF sin la extensión
    if pdf_filename:
        invoice_name = os.path.splitext(os.path.basename(pdf_filename))[0]
    else:
        invoice_name = "Factura"
    
    # Añadir título
    ws['A1'] = f"DETALLE DE PRODUCTOS - {invoice_name}"
    ws['A1'].font = Font(bold=True, size=14)
    ws.merge_cells('A1:G1')
    ws['A1'].alignment = Alignment(horizontal='center')
    
    # Añadir fecha de creación
    from datetime import datetime
    ws['A2'] = f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    ws.merge_cells('A2:G2')
    ws['A2'].alignment = Alignment(horizontal='center')
    
    # Añadir espacio
    ws.append([])
    
    # Convertir DataFrame a filas de Excel
    rows = dataframe_to_rows(df, index=False, header=True)
    
    # Añadir los datos
    for r_idx, row in enumerate(rows, 4):  # Empezar en fila 4
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            
            # Formato para la cabecera
            if r_idx == 4:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = openpyxl.styles.PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Ajustar anchos de columnas
    column_widths = {
        'A': 10,  # Código
        'B': 15,  # Referencia (más ancho para caracteres especiales)
        'C': 60,  # Descripción
        'D': 10,  # Cantidad
        'E': 10,  # Precio
        'F': 8    # IVA
    }
    
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    
    # Aplicar bordes a todos los datos
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    for row in ws.iter_rows(min_row=4, max_row=len(df)+4, min_col=1, max_col=6):
        for cell in row:
            cell.border = thin_border
            if cell.row > 4:  # No aplicar a la cabecera
                if cell.column in [4, 5]:  # Columnas Cantidad y Precio
                    cell.alignment = Alignment(horizontal='right')
                elif cell.column == 6:  # Columna IVA
                    cell.alignment = Alignment(horizontal='center')
    
    # Añadir totales
    total_row = len(df) + 6
    ws.cell(row=total_row, column=3, value="TOTAL PRODUCTOS:").font = Font(bold=True)
    ws.cell(row=total_row, column=4, value=len(df)).font = Font(bold=True)
    
    ws.cell(row=total_row+1, column=3, value="TOTAL UNIDADES:").font = Font(bold=True)
    ws.cell(row=total_row+1, column=4, value=df['Cantidad'].sum()).font = Font(bold=True)
    
    ws.cell(row=total_row+2, column=3, value="VALOR TOTAL:").font = Font(bold=True)
    total_value = (df['Cantidad'] * df['Precio']).sum()
    ws.cell(row=total_row+2, column=4, value=total_value).font = Font(bold=True)
    ws.cell(row=total_row+2, column=4).number_format = '#,##0.00€'
    
    # Guardar el archivo Excel
    wb.save(excel_path)
    return excel_path

def process_invoice_pdf(pdf_path, output_excel=None):
    """
    Procesa un PDF de factura y devuelve un DataFrame con los productos
    """
    print(f"Procesando factura: {pdf_path}")
    
    # Extraer datos
    df = extract_invoice_data(pdf_path)
    
    if df.empty:
        print("No se pudieron extraer datos del PDF.")
        # Mostrar texto crudo para debug
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                print(f"\n--- Página {i+1} (primeras 500 chars) ---")
                print(text[:500])
        return None
    
    # Limpiar descripciones
    df = clean_descriptions(df)
    
    # Mostrar resumen
    print(f"✓ Extraídos {len(df)} registros")
    print(f"✓ Total unidades: {df['Cantidad'].sum()}")
    print(f"✓ Valor total: {round((df['Cantidad'] * df['Precio']).sum(), 2)}€")
    print(f"✓ IVA aplicado: {df['IVA'].unique()}%")
    
    # Mostrar algunas referencias para verificar
    print(f"✓ Ejemplos de referencias: {df['Referencia'].head(5).tolist()}")
    
    # Guardar en Excel si se especifica
    if output_excel:
        excel_path = save_to_excel(df, output_excel, pdf_path)
        print(f"✓ Datos guardados en '{excel_path}'")
    
    return df

# Instalar openpyxl si no está instalado
try:
    import openpyxl
except ImportError:
    print("Instalando openpyxl...")
    import subprocess
    subprocess.check_call(["pip", "install", "openpyxl"])
    import openpyxl

# Ejemplo de uso
if __name__ == "__main__":
    # Procesar una factura
    pdf_file = r"D:\herloc.programacion\lector_facturas_pdf\files_repo\003723567.pdf"  # Cambia por la ruta de tu PDF
    excel_file = "datos_factura.xlsx"
    
    df = process_invoice_pdf(pdf_file, excel_file)
    
    # Mostrar los primeros 10 registros
    if df is not None:
        print("\nPrimeros 10 productos:")
        print(df.head(10).to_string(index=False))


# import pdfplumber
# import pandas as pd
# import re
# from openpyxl import Workbook
# from openpyxl.styles import Font, Alignment, Border, Side
# from openpyxl.utils.dataframe import dataframe_to_rows
# import os

# def extract_invoice_data(pdf_path):
#     """
#     Extrae datos de productos de facturas PDF con el formato específico de Liderpapel
#     """
#     all_data = []
    
#     with pdfplumber.open(pdf_path) as pdf:
#         for page in pdf.pages:
#             # Extraer texto de la página
#             text = page.extract_text()
#             lines = text.split('\n')
            
#             in_products_section = False
#             current_product = None
            
#             for i, line in enumerate(lines):
#                 # Detectar inicio de la sección de productos
#                 if 'Código Referencia Descripción' in line:
#                     in_products_section = True
#                     continue
                
#                 # Detectar fin de la sección de productos
#                 if any(end_marker in line for end_marker in ['Empresa adherida', 'Forma de Pago', 'CONDICIONES DE VENTA']):
#                     in_products_section = False
#                     if current_product:
#                         all_data.append(current_product)
#                         current_product = None
#                     continue
                
#                 if in_products_section and line.strip():
#                     # Patrón para líneas de producto completas
#                     pattern = r'^(\d+)\s+([A-Z0-9]+)\s+(.+?)\s+(\d+)\s+([\d,]+)\s+([\d,]+)\s+(\d+)$'
#                     match = re.match(pattern, line.strip())
                    
#                     if match:
#                         # Si hay un producto actual pendiente, guardarlo
#                         if current_product:
#                             all_data.append(current_product)
                        
#                         codigo, referencia, descripcion, cantidad, precio, importe, iva = match.groups()
#                         current_product = {
#                             'Código': codigo,
#                             'Referencia': referencia,
#                             'Descripción': descripcion.strip(),
#                             'Cantidad': cantidad,
#                             'Precio': precio,
#                             'IVA': iva
#                         }
#                     elif current_product:
#                         # Si ya tenemos un producto actual, podría ser continuación de descripción
#                         # Verificar si la línea parece ser datos numéricos (no continuación)
#                         if not re.search(r'\d+[\s,]+\d+[\s,]+\d+$', line.strip()):
#                             current_product['Descripción'] += " " + line.strip()
#                         else:
#                             # Si parece ser una nueva línea de producto pero no matcheó el patrón completo
#                             all_data.append(current_product)
#                             current_product = None
            
#             # Añadir el último producto si existe
#             if current_product:
#                 all_data.append(current_product)
    
#     # Limpiar y convertir los datos
#     cleaned_data = []
#     for item in all_data:
#         try:
#             # Limpiar descripción de caracteres extraños
#             descripcion = re.sub(r'\s+', ' ', item['Descripción']).strip()
            
#             cleaned_data.append({
#                 'Código': item['Código'],
#                 'Referencia': item['Referencia'],
#                 'Descripción': descripcion,
#                 'Cantidad': int(item['Cantidad']),
#                 'Precio': float(item['Precio'].replace(',', '.')),
#                 'IVA': int(item['IVA'])
#             })
#         except (ValueError, KeyError):
#             # Saltar items con problemas de conversión
#             continue
    
#     return pd.DataFrame(cleaned_data)

# def clean_descriptions(df):
#     """
#     Limpia y completa descripciones basado en patrones comunes
#     """
#     # Diccionario de mapeo para completar descripciones conocidas
#     description_map = {
#         'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL': {
#             'LP541': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC AZUL DIA PAGINA',
#             'LP542': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC NEGRO DIA PAGINA',
#             'LP543': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC ROSA DIA PAGINA',
#             'LP544': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC ROJO DIA PAGINA',
#             'LP546': 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL BASIC VERDE MENTA DIA PAGINA'
#         },
#         'BOLIGRAFO': {
#             '8373602': 'BOLIGRAFO BIC CRISTAL ORIGINAL TINTA AZUL',
#             'KF18625': 'BOLIGRAFO Q-CONNECT RETRACTIL BORRABLE 0,7 MM COLOR AZUL',
#             'KF18626': 'BOLIGRAFO Q-CONNECT RETRACTIL BORRABLE 0,7 MM COLOR ROJO'
#         }
#     }
    
#     for idx, row in df.iterrows():
#         desc = row['Descripción']
#         ref = row['Referencia']
        
#         # Completar descripciones de agendas
#         if 'AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL' in desc:
#             if ref in description_map['AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL']:
#                 df.at[idx, 'Descripción'] = description_map['AGENDA ESCOLAR LIDERPAPEL 25-26 ESPIRAL'][ref]
        
#         # Completar descripciones de bolígrafos
#         elif 'BOLIGRAFO' in desc:
#             if ref in description_map['BOLIGRAFO']:
#                 df.at[idx, 'Descripción'] = description_map['BOLIGRAFO'][ref]
    
#     return df

# def save_to_excel(df, excel_path, pdf_filename=None):
#     """
#     Guarda el DataFrame en un archivo Excel con formato profesional
#     """
#     # Crear un nuevo libro de Excel
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Productos Factura"
    
#     # Obtener el nombre del archivo PDF sin la extensión
#     if pdf_filename:
#         invoice_name = os.path.splitext(os.path.basename(pdf_filename))[0]
#     else:
#         invoice_name = "Factura"
    
#     # Añadir título
#     ws['A1'] = f"DETALLE DE PRODUCTOS - {invoice_name}"
#     ws['A1'].font = Font(bold=True, size=14)
#     ws.merge_cells('A1:G1')
#     ws['A1'].alignment = Alignment(horizontal='center')
    
#     # Añadir fecha de creación
#     from datetime import datetime
#     ws['A2'] = f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
#     ws.merge_cells('A2:G2')
#     ws['A2'].alignment = Alignment(horizontal='center')
    
#     # Añadir espacio
#     ws.append([])
    
#     # Convertir DataFrame a filas de Excel
#     rows = dataframe_to_rows(df, index=False, header=True)
    
#     # Añadir los datos
#     for r_idx, row in enumerate(rows, 4):  # Empezar en fila 4
#         for c_idx, value in enumerate(row, 1):
#             cell = ws.cell(row=r_idx, column=c_idx, value=value)
            
#             # Formato para la cabecera
#             if r_idx == 4:
#                 cell.font = Font(bold=True, color="FFFFFF")
#                 cell.fill = openpyxl.styles.PatternFill(start_color="366092", end_color="366092", fill_type="solid")
#                 cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
#     # Ajustar anchos de columnas
#     column_widths = {
#         'A': 10,  # Código
#         'B': 12,  # Referencia
#         'C': 60,  # Descripción
#         'D': 10,  # Cantidad
#         'E': 10,  # Precio
#         'F': 8    # IVA
#     }
    
#     for col, width in column_widths.items():
#         ws.column_dimensions[col].width = width
    
#     # Aplicar bordes a todos los datos
#     thin_border = Border(
#         left=Side(style='thin'),
#         right=Side(style='thin'),
#         top=Side(style='thin'),
#         bottom=Side(style='thin')
#     )
    
#     for row in ws.iter_rows(min_row=4, max_row=len(df)+4, min_col=1, max_col=6):
#         for cell in row:
#             cell.border = thin_border
#             if cell.row > 4:  # No aplicar a la cabecera
#                 if cell.column in [4, 5]:  # Columnas Cantidad y Precio
#                     cell.alignment = Alignment(horizontal='right')
#                 elif cell.column == 6:  # Columna IVA
#                     cell.alignment = Alignment(horizontal='center')
    
#     # Añadir totales
#     total_row = len(df) + 6
#     ws.cell(row=total_row, column=3, value="TOTAL PRODUCTOS:").font = Font(bold=True)
#     ws.cell(row=total_row, column=4, value=len(df)).font = Font(bold=True)
    
#     ws.cell(row=total_row+1, column=3, value="TOTAL UNIDADES:").font = Font(bold=True)
#     ws.cell(row=total_row+1, column=4, value=df['Cantidad'].sum()).font = Font(bold=True)
    
#     ws.cell(row=total_row+2, column=3, value="VALOR TOTAL:").font = Font(bold=True)
#     total_value = (df['Cantidad'] * df['Precio']).sum()
#     ws.cell(row=total_row+2, column=4, value=total_value).font = Font(bold=True)
#     ws.cell(row=total_row+2, column=4).number_format = '#,##0.00€'
    
#     # Guardar el archivo Excel
#     wb.save(excel_path)
#     return excel_path

# def process_invoice_pdf(pdf_path, output_excel=None):
#     """
#     Procesa un PDF de factura y devuelve un DataFrame con los productos
#     """
#     print(f"Procesando factura: {pdf_path}")
    
#     # Extraer datos
#     df = extract_invoice_data(pdf_path)
    
#     if df.empty:
#         print("No se pudieron extraer datos del PDF.")
#         return None
    
#     # Limpiar descripciones
#     df = clean_descriptions(df)
    
#     # Mostrar resumen
#     print(f"✓ Extraídos {len(df)} registros")
#     print(f"✓ Total unidades: {df['Cantidad'].sum()}")
#     print(f"✓ Valor total: {round((df['Cantidad'] * df['Precio']).sum(), 2)}€")
#     print(f"✓ IVA aplicado: {df['IVA'].unique()}%")
    
#     # Guardar en Excel si se especifica
#     if output_excel:
#         excel_path = save_to_excel(df, output_excel, pdf_path)
#         print(f"✓ Datos guardados en '{excel_path}'")
    
#     return df

# # Instalar openpyxl si no está instalado
# try:
#     import openpyxl
# except ImportError:
#     print("Instalando openpyxl...")
#     import subprocess
#     subprocess.check_call(["pip", "install", "openpyxl"])
#     import openpyxl

# # Ejemplo de uso
# if __name__ == "__main__":
#     # Procesar una factura
#     pdf_file = r"D:\herloc.programacion\papeleria_temporal\003704074.pdf"  # Cambia por la ruta de tu PDF
#     excel_file = "datos_factura.xlsx"
    
#     df = process_invoice_pdf(pdf_file, excel_file)
    
#     # Mostrar los primeros 5 registros
#     if df is not None:
#         print("\nPrimeros 5 productos:")
#         print(df.head(5).to_string(index=False))