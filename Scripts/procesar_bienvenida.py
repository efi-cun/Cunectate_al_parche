import os
import sys
import glob
import json
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASISTENCIA_DIR = os.path.join(BASE_DIR, "Asistencia")
PLANTA_DIR = os.path.join(BASE_DIR, "Planta")
EXCEL_OUTPUT = os.path.join(BASE_DIR, "Consolidado_Bienvenida_Facilitadores.xlsx")
HTML_OUTPUT = os.path.join(BASE_DIR, "tablero.html")
HTML_RESPALDO = os.path.join(BASE_DIR, "tablero_respaldo.html")

print("=== INICIANDO PROCESAMIENTO BIENVENIDA FACILITADORES CUN ===")

# --- DICCIONARIO DE LIMPIEZA DE CARACTERES (ESTILO procesar_reporte.ps1) ---
replacements = [
    ("Ã\xa0", "a"), ("Ã¡", "a"), ("Ã\xa2", "a"), ("Ã£", "a"), ("Ã¤", "a"),
    ("Ã¨", "e"), ("Ã©", "e"), ("Ãª", "e"), ("Ã«", "e"),
    ("Ã¬", "i"), ("Ã\xad", "i"), ("Ã®", "i"), ("Ã¯", "i"),
    ("Ã²", "o"), ("Ã³", "o"), ("Ã´", "o"), ("Ãµ", "o"), ("Ã¶", "o"),
    ("Ã¹", "u"), ("Ãº", "u"), ("Ã»", "u"), ("Ã¼", "u"),
    ("Ã±", "n"), ("Ã‘", "N"), ("Â¿", ""),
    ("Ã\x81", "A"), ("Ã\x89", "E"), ("Ã\x8d", "I"), ("Ã\x93", "O"), ("Ã\x9a", "U"),
    ("á", "a"), ("é", "e"), ("í", "i"), ("ó", "o"), ("ú", "u"), ("ñ", "n"),
    ("Á", "A"), ("É", "E"), ("Í", "I"), ("Ó", "O"), ("Ú", "U"), ("Ñ", "N")
]

def clean_text(txt):
    if pd.isna(txt) or txt is None:
        return ""
    s = str(txt)
    for old, new in replacements:
        s = s.replace(old, new)
    return s.strip()

# 1. Cargar y Limpiar Archivo de Asistencia
csv_files = glob.glob(os.path.join(ASISTENCIA_DIR, "*.csv"))
if not csv_files:
    csv_files = glob.glob(os.path.join(ASISTENCIA_DIR, "**", "*.csv"), recursive=True)

if not csv_files:
    print("ERROR: No se encontró ningún archivo .csv en la carpeta Asistencia.")
    sys.exit(1)

csv_files.sort(key=os.path.getmtime, reverse=True)
asistencia_file = csv_files[0]
print(f"[1/5] Limpiando y procesando archivo de Asistencia: {os.path.basename(asistencia_file)}")

try:
    df_asist_raw = pd.read_csv(asistencia_file, encoding='utf-8')
except Exception:
    df_asist_raw = pd.read_csv(asistencia_file, encoding='latin1')

print(f"      Total registros recibidos en Asistencia: {len(df_asist_raw)}")

col_ced_asist = next((c for c in df_asist_raw.columns if 'CEDULA' in str(c).upper() or 'IDENTIFICAC' in str(c).upper()), df_asist_raw.columns[0])
col_nom_asist = next((c for c in df_asist_raw.columns if 'NAME' in str(c).upper() or 'NOMBRE' in str(c).upper()), df_asist_raw.columns[1] if len(df_asist_raw.columns) > 1 else df_asist_raw.columns[0])

df_asist_clean = pd.DataFrame()
df_asist_clean['CEDULA'] = pd.to_numeric(df_asist_raw[col_ced_asist], errors='coerce').astype('Int64').astype(str).str.strip()

def format_nombre(val):
    txt = clean_text(val)
    if ',' in txt:
        parts = txt.split(',', 1)
        return f"{parts[1].strip()} {parts[0].strip()}".upper()
    return txt.upper()

df_asist_clean['NOMBRE'] = df_asist_raw[col_nom_asist].apply(format_nombre)

# 2. Cargar Base de Planta
xlsx_files = glob.glob(os.path.join(PLANTA_DIR, "*.xlsx"))
if not xlsx_files:
    xlsx_files = glob.glob(os.path.join(PLANTA_DIR, "**", "*.xlsx"), recursive=True)

if not xlsx_files:
    print("ERROR: No se encontró ningún archivo .xlsx en la carpeta Planta.")
    sys.exit(1)

planta_file = xlsx_files[0]
print(f"[2/5] Cargando base de referencia Planta: {os.path.basename(planta_file)}")

xl_planta = pd.ExcelFile(planta_file)
sheet_planta = "BASE FACILITADORES PRINCIPAL " if "BASE FACILITADORES PRINCIPAL " in xl_planta.sheet_names else xl_planta.sheet_names[0]
df_planta_raw = pd.read_excel(xl_planta, sheet_name=sheet_planta)

col_ced_planta = next((c for c in df_planta_raw.columns if 'IDENTIFICAC' in str(c).upper() or 'CEDULA' in str(c).upper()), df_planta_raw.columns[6])
col_nom_planta = next((c for c in df_planta_raw.columns if 'NOMBRE' in str(c).upper() and 'COMPLETO' in str(c).upper()), df_planta_raw.columns[7])
col_invitacion = next((c for c in df_planta_raw.columns if 'INVITAC' in str(c).upper()), 'INVITACIÓN')
col_nivel1 = next((c for c in df_planta_raw.columns if 'NIVEL 1' in str(c).upper()), 'Nombre Nivel 1')
col_nivel2 = next((c for c in df_planta_raw.columns if 'NIVEL 2' in str(c).upper()), 'Nombre Nivel 2')
col_nivel3 = next((c for c in df_planta_raw.columns if 'NIVEL 3' in str(c).upper()), 'Nombre Nivel 3')
col_modalidad_alt = next((c for c in df_planta_raw.columns if 'MODALIDAD' in str(c).upper()), 'Modalidad')
col_cargo = next((c for c in df_planta_raw.columns if 'CARGO' in str(c).upper()), 'Descripción Cargo')
col_centro_costo = next((c for c in df_planta_raw.columns if 'CENTRO COSTO' in str(c).upper()), 'Nombre Centro Costo')

df_planta_raw['CEDULA_CLEAN'] = pd.to_numeric(df_planta_raw[col_ced_planta], errors='coerce').astype('Int64').astype(str).str.strip()
planta_cedulas_set = set(df_planta_raw['CEDULA_CLEAN'].dropna().unique())
total_planta = len(df_planta_raw)

# Reorganizar PLANTA_BASE con CÉDULA en Columna A (Col 1)
planta_cols_ordered = [
    'CEDULA_CLEAN',
    col_invitacion,
    col_nom_planta,
    col_nivel1,
    col_nivel2,
    col_nivel3,
    col_modalidad_alt,
    col_cargo,
    col_centro_costo
]
for c in df_planta_raw.columns:
    if c not in planta_cols_ordered and c != col_ced_planta:
        planta_cols_ordered.append(c)

df_planta_export = df_planta_raw[planta_cols_ordered].copy()
df_planta_export.rename(columns={'CEDULA_CLEAN': 'CEDULA'}, inplace=True)

# 3. Deduplicación y Validación contra Planta
print("[3/5] Validando presencia en Planta y detectando duplicados...")

is_duplicated_in_asist = df_asist_clean.duplicated(subset=['CEDULA'], keep='first')
is_not_in_planta = ~df_asist_clean['CEDULA'].isin(planta_cedulas_set)

df_duplicados = df_asist_clean[is_duplicated_in_asist | is_not_in_planta].copy()
df_unicos = df_asist_clean[~is_duplicated_in_asist & ~is_not_in_planta].copy()

print(f"      Total Facilitadores en Base Planta: {total_planta}")
print(f"      Registros Válidos y Únicos (Consolidado_Unicos): {len(df_unicos)}")
print(f"      Registros Duplicados / No Encontrados en Planta (Duplicados): {len(df_duplicados)}")

# 4. Generar Libro de Excel con Fórmulas VLOOKUP Desplegadas en 1.500 Filas
TOTAL_FILAS_DESPLEGADAS = 1500
print(f"[4/5] Desplegando fórmulas BUSCARV a lo largo de {TOTAL_FILAS_DESPLEGADAS} filas en las columnas...")

wb = openpyxl.Workbook()
wb.remove(wb.active)

header_fill = PatternFill(start_color="004B28", end_color="004B28", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
data_font = Font(name="Calibri", size=11, color="1F2937")
border_thin = Side(style='thin', color="D1D5DB")
cell_border = Border(left=border_thin, right=border_thin, top=border_thin, bottom=border_thin)
alt_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

headers_consolidado = [
    'INVITACIÓN',
    'CEDULA',
    'NOMBRE',
    'Nombre Nivel 2',
    'Nombre Nivel 3',
    'Descripción Cargo',
    'Nombre Centro Costo'
]

def format_consolidado_with_full_column_formulas(ws, df_asist_data):
    ws.views.sheetView[0].showGridLines = True
    ws.append(headers_consolidado)
    
    for col_num, h in enumerate(headers_consolidado, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28

    data_rows = list(df_asist_data.itertuples(index=False))
    num_data_rows = len(data_rows)
    total_rows = max(num_data_rows, TOTAL_FILAS_DESPLEGADAS)

    for r_idx in range(2, total_rows + 2):
        if (r_idx - 2) < num_data_rows:
            row_data = data_rows[r_idx - 2]
            ced_val = str(row_data.CEDULA)
            nom_val = str(row_data.NOMBRE)
            ced_cell_value = int(ced_val) if ced_val.isdigit() else ced_val
            nom_formula = nom_val
        else:
            ced_cell_value = ""
            nom_formula = f'=IF(B{r_idx}="","",IFERROR(VLOOKUP(B{r_idx}, PLANTA_BASE!A:I, 3, FALSE), ""))'

        formula_invitacion = f'=IF(B{r_idx}="","",IFERROR(VLOOKUP(B{r_idx}, PLANTA_BASE!A:I, 2, FALSE), "NO ENCONTRADO EN PLANTA"))'
        formula_nivel2     = f'=IF(B{r_idx}="","",IFERROR(VLOOKUP(B{r_idx}, PLANTA_BASE!A:I, 5, FALSE), "NO ENCONTRADO EN PLANTA"))'
        formula_nivel3     = f'=IF(B{r_idx}="","",IFERROR(VLOOKUP(B{r_idx}, PLANTA_BASE!A:I, 6, FALSE), "NO ENCONTRADO EN PLANTA"))'
        formula_cargo      = f'=IF(B{r_idx}="","",IFERROR(VLOOKUP(B{r_idx}, PLANTA_BASE!A:I, 8, FALSE), "DOCENTE / FACILITADOR"))'
        formula_escuela    = f'=IF(B{r_idx}="","",IFERROR(VLOOKUP(B{r_idx}, PLANTA_BASE!A:I, 9, FALSE), "GENERAL"))'

        ws.cell(row=r_idx, column=1, value=formula_invitacion)
        ws.cell(row=r_idx, column=2, value=ced_cell_value)
        ws.cell(row=r_idx, column=3, value=nom_formula)
        ws.cell(row=r_idx, column=4, value=formula_nivel2)
        ws.cell(row=r_idx, column=5, value=formula_nivel3)
        ws.cell(row=r_idx, column=6, value=formula_cargo)
        ws.cell(row=r_idx, column=7, value=formula_escuela)
        
        ws.row_dimensions[r_idx].height = 20
        use_alt = (r_idx % 2 == 0)
        for c_idx in range(1, len(headers_consolidado) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = data_font
            cell.border = cell_border
            if use_alt:
                cell.fill = alt_fill
            
            if c_idx == 2: # CEDULA
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = '0'
            elif c_idx in [1, 4, 5]: # INVITACION, Nivel 2, Nivel 3
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = 22

def format_planta_sheet(ws, df_data):
    ws.views.sheetView[0].showGridLines = True
    headers = list(df_data.columns)
    ws.append(headers)
    
    for col_num, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 28

    for r_idx, row in enumerate(df_data.itertuples(index=False), start=2):
        ws.append(list(row))
        ws.row_dimensions[r_idx].height = 20
        use_alt = (r_idx % 2 == 0)
        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.font = data_font
            cell.border = cell_border
            if use_alt:
                cell.fill = alt_fill
            
            if headers[c_idx-1] == 'CEDULA':
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.number_format = '0'
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = 18

# Hoja 1: Consolidado_Unicos
ws_unicos = wb.create_sheet(title="Consolidado_Unicos")
format_consolidado_with_full_column_formulas(ws_unicos, df_unicos)

# Hoja 2: Duplicados
ws_duplicados = wb.create_sheet(title="Duplicados")
format_consolidado_with_full_column_formulas(ws_duplicados, df_duplicados)

# Hoja 3: PLANTA_BASE
ws_planta = wb.create_sheet(title="PLANTA_BASE")
format_planta_sheet(ws_planta, df_planta_export)

# Hoja 4: Resumen_y_Graficos
ws_resumen = wb.create_sheet(title="Resumen_y_Graficos")
ws_resumen.views.sheetView[0].showGridLines = True
ws_resumen.cell(row=1, column=1, value="RESUMEN EJECUTIVO DE ASISTENCIA CUN").font = Font(name="Calibri", size=16, bold=True, color="004B28")

planta_lookup = df_planta_raw.drop_duplicates(subset=['CEDULA_CLEAN']).set_index('CEDULA_CLEAN')

eval_records = []
for idx, row in df_unicos.iterrows():
    ced = row['CEDULA']
    nom = row['NOMBRE']
    
    if ced in planta_lookup.index:
        p_row = planta_lookup.loc[ced]
        inv_val = clean_text(p_row.get(col_invitacion, ''))
        if not inv_val or inv_val.upper() == 'NAN':
            inv_val = clean_text(p_row.get(col_modalidad_alt, 'VIRTUAL'))
        reg_val = clean_text(p_row.get(col_nivel2, 'NO ASIGNADO'))
        sed_val = clean_text(p_row.get(col_nivel3, 'NO ASIGNADO'))
        car_val = clean_text(p_row.get(col_cargo, 'DOCENTE / FACILITADOR'))
        esc_val = clean_text(p_row.get(col_centro_costo, 'GENERAL'))
    else:
        inv_val = "VIRTUAL"
        reg_val = "REGIONAL BOGOTA"
        sed_val = "BOGOTA CENTRO"
        car_val = "DOCENTE / FACILITADOR"
        esc_val = "GENERAL"

    inv_upper = inv_val.upper()
    tipo_mod = 'PRESENCIAL' if ('PRESENCIAL' in inv_upper or 'BOGOTA' in inv_upper or 'BOGOTÁ' in inv_upper or 'SEDE' in inv_upper) else 'VIRTUAL'

    eval_records.append({
        'INVITACIÓN': inv_val,
        'CEDULA': ced,
        'NOMBRE': nom,
        'Nombre Nivel 2': reg_val,
        'Nombre Nivel 3': sed_val,
        'Descripción Cargo': car_val,
        'Nombre Centro Costo': esc_val,
        'TIPO_MODALIDAD': tipo_mod
    })

df_eval = pd.DataFrame(eval_records)

ws_resumen.cell(row=3, column=1, value="Modalidad / Invitación").font = header_font
ws_resumen.cell(row=3, column=1).fill = header_fill
ws_resumen.cell(row=3, column=2, value="Cantidad Registros").font = header_font
ws_resumen.cell(row=3, column=2).fill = header_fill

if len(df_eval) > 0:
    modalidad_counts = df_eval['INVITACIÓN'].value_counts()
    row_curr = 4
    for mod_name, count_val in modalidad_counts.items():
        ws_resumen.cell(row=row_curr, column=1, value=str(mod_name)).border = cell_border
        ws_resumen.cell(row=row_curr, column=2, value=int(count_val)).border = cell_border
        ws_resumen.cell(row=row_curr, column=2).alignment = Alignment(horizontal="center")
        row_curr += 1

    ws_resumen.cell(row=row_curr, column=1, value="TOTAL GENERAL").font = Font(bold=True)
    ws_resumen.cell(row=row_curr, column=1).border = cell_border
    ws_resumen.cell(row=row_curr, column=2, value=f"=SUM(B4:B{row_curr-1})").font = Font(bold=True)
    ws_resumen.cell(row=row_curr, column=2).border = cell_border
    ws_resumen.cell(row=row_curr, column=2).alignment = Alignment(horizontal="center")

    pie = PieChart()
    labels = Reference(ws_resumen, min_col=1, min_row=4, max_row=row_curr-1)
    data = Reference(ws_resumen, min_col=2, min_row=3, max_row=row_curr-1)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Distribución por Modalidad (Únicos Válidos)"
    pie.width = 14
    pie.height = 8
    ws_resumen.add_chart(pie, "D3")

    row_reg_start = row_curr + 3
    ws_resumen.cell(row=row_reg_start, column=1, value="Regional (Nombre Nivel 2)").font = header_font
    ws_resumen.cell(row=row_reg_start, column=1).fill = header_fill
    ws_resumen.cell(row=row_reg_start, column=2, value="Presenciales").font = header_font
    ws_resumen.cell(row=row_reg_start, column=2).fill = header_fill

    regional_presencial = df_eval[df_eval['TIPO_MODALIDAD'] == 'PRESENCIAL']['Nombre Nivel 2'].value_counts()
    row_curr_reg = row_reg_start + 1
    for reg_name, count_val in regional_presencial.items():
        ws_resumen.cell(row=row_curr_reg, column=1, value=str(reg_name)).border = cell_border
        ws_resumen.cell(row=row_curr_reg, column=2, value=int(count_val)).border = cell_border
        ws_resumen.cell(row=row_curr_reg, column=2).alignment = Alignment(horizontal="center")
        row_curr_reg += 1

    if len(regional_presencial) > 0:
        chart_bar = BarChart()
        chart_bar.type = "col"
        chart_bar.style = 10
        chart_bar.title = "Participación Presencial por Regional"
        chart_bar.y_axis.title = "Cantidad"
        chart_bar.x_axis.title = "Regional"
        data_bar = Reference(ws_resumen, min_col=2, min_row=row_reg_start, max_row=row_curr_reg-1)
        cats_bar = Reference(ws_resumen, min_col=1, min_row=row_reg_start+1, max_row=row_curr_reg-1)
        chart_bar.add_data(data_bar, titles_from_data=True)
        chart_bar.set_categories(cats_bar)
        chart_bar.width = 16
        chart_bar.height = 8
        ws_resumen.add_chart(chart_bar, f"D{row_reg_start}")

# Guardar con manejo de bloqueo si el archivo está abierto en Excel
try:
    wb.save(EXCEL_OUTPUT)
    print("      Libro Excel consolidado guardado exitosamente.")
except PermissionError:
    print("      ADVERTENCIA: El archivo Consolidado_Bienvenida_Facilitadores.xlsx está abierto por el usuario.")
    alt_out = os.path.join(BASE_DIR, "Consolidado_Bienvenida_Facilitadores_Nuevo.xlsx")
    wb.save(alt_out)
    print(f"      Guardado exitosamente en copia alternativa: {alt_out}")

# 5. Generar Tablero HTML (CON PORCENTAJE DE ASISTENCIA CALCULADO SOBRE EL TOTAL DE PLANTA)
print("=== PREPARANDO DATOS PARA EL TABLERO HTML (PORCENTAJE RESPECTO AL TOTAL DE LA PLANTA) ===")

total_registros = len(df_eval)
presencial_count = len(df_eval[df_eval['TIPO_MODALIDAD'] == 'PRESENCIAL']) if total_registros > 0 else 0
virtual_count = len(df_eval[df_eval['TIPO_MODALIDAD'] == 'VIRTUAL']) if total_registros > 0 else 0

# Porcentaje Global de Asistencia respecto al Total de la Planta
pct_asistencia_global = round((total_registros / total_planta) * 100, 1) if total_planta > 0 else 0

regional_presencial_dict = df_eval[df_eval['TIPO_MODALIDAD'] == 'PRESENCIAL']['Nombre Nivel 2'].value_counts().to_dict() if total_registros > 0 else {}

# Agrupación por Escuela calculando % relativo al Total de la Planta o al Total Registros
escuelas_data = []
if total_registros > 0:
    for escuela, group in df_eval.groupby('Nombre Centro Costo'):
        t_count = len(group)
        p_count = len(group[group['TIPO_MODALIDAD'] == 'PRESENCIAL'])
        v_count = len(group[group['TIPO_MODALIDAD'] == 'VIRTUAL'])
        pct = round((t_count / total_registros) * 100, 1) if total_registros > 0 else 0
        escuelas_data.append({
            'escuela': str(escuela),
            'total': t_count,
            'presencial': p_count,
            'virtual': v_count,
            'porcentaje': pct
        })

    escuelas_data.sort(key=lambda x: x['total'], reverse=True)

html_content = f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Tablero Bienvenida Facilitadores CUN - Diapositiva Ejecutiva</title>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;600;700;800&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <style>
        :root {{
            --cun-green: #00A859;
            --cun-green-glow: #00FF87;
            --cun-dark: #050B14;
            --cun-card-bg: rgba(10, 20, 38, 0.78);
            --cun-cyan: #00F2FE;
            --cun-text: #F3F4F6;
            --cun-muted: #9CA3AF;
            --cun-border: rgba(0, 168, 89, 0.3);
        }}

        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: 'Plus Jakarta Sans', sans-serif;
        }}

        body {{
            background-color: var(--cun-dark);
            color: var(--cun-text);
            min-height: 100vh;
            overflow-x: hidden;
            position: relative;
        }}

        #galaxy-canvas {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            z-index: 0;
            pointer-events: none;
        }}

        .dashboard-container {{
            position: relative;
            z-index: 1;
            max-width: 1350px;
            margin: 0 auto;
            padding: 2rem 1.5rem;
        }}

        header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: var(--cun-card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--cun-border);
            padding: 1.25rem 2rem;
            border-radius: 20px;
            margin-bottom: 2rem;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
        }}

        .brand-section {{
            display: flex;
            align-items: center;
            gap: 1.25rem;
        }}

        .brand-logo-badge {{
            background: linear-gradient(135deg, #00A859, #00F2FE);
            width: 54px;
            height: 54px;
            border-radius: 16px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.6rem;
            color: #050B14;
            box-shadow: 0 0 24px rgba(0, 242, 254, 0.4);
        }}

        .brand-title h1 {{
            font-size: 1.6rem;
            font-weight: 800;
            letter-spacing: -0.5px;
            background: linear-gradient(90deg, #FFFFFF, #00FF87);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .brand-title p {{
            font-size: 0.85rem;
            color: var(--cun-muted);
        }}

        .header-actions {{
            display: flex;
            align-items: center;
            gap: 1rem;
        }}

        .live-badge {{
            display: flex;
            align-items: center;
            gap: 8px;
            background: rgba(0, 255, 135, 0.1);
            border: 1px solid rgba(0, 255, 135, 0.3);
            padding: 6px 16px;
            border-radius: 30px;
            font-size: 0.8rem;
            color: var(--cun-green-glow);
            font-weight: 600;
        }}

        .live-dot {{
            width: 8px;
            height: 8px;
            background-color: var(--cun-green-glow);
            border-radius: 50%;
            box-shadow: 0 0 10px var(--cun-green-glow);
            animation: pulse 1.5s infinite;
        }}

        @keyframes pulse {{
            0% {{ transform: scale(0.95); opacity: 0.8; }}
            50% {{ transform: scale(1.2); opacity: 1; }}
            100% {{ transform: scale(0.95); opacity: 0.8; }}
        }}

        /* HERO CARD PRINCIPAL: TOTAL REGISTROS ÚNICOS VÁLIDOS (TODO A LO LARGO) */
        .kpi-hero-card {{
            background: var(--cun-card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--cun-border);
            border-radius: 24px;
            padding: 2rem;
            margin-bottom: 1.5rem;
            position: relative;
            overflow: hidden;
            box-shadow: 0 12px 40px rgba(0, 168, 89, 0.15);
            transition: all 0.3s ease;
        }}

        .kpi-hero-card:hover {{
            border-color: var(--cun-green-glow);
            box-shadow: 0 16px 50px rgba(0, 255, 135, 0.25);
        }}

        .hero-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1.25rem;
        }}

        .hero-info .kpi-title {{
            font-size: 1rem;
            color: var(--cun-muted);
            font-weight: 700;
            letter-spacing: 0.5px;
        }}

        .hero-info .kpi-value {{
            font-size: 3.5rem;
            font-weight: 800;
            color: #FFFFFF;
            line-height: 1.1;
            margin: 0.25rem 0;
            background: linear-gradient(90deg, #FFFFFF, var(--cun-cyan));
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .hero-info .kpi-subtext {{
            font-size: 0.85rem;
            color: var(--cun-muted);
        }}

        .hero-progress-wrapper {{
            margin-top: 1rem;
        }}

        .hero-progress-meta {{
            display: flex;
            justify-content: space-between;
            font-size: 0.82rem;
            font-weight: 700;
            color: var(--cun-muted);
            margin-bottom: 8px;
            letter-spacing: 0.5px;
        }}

        .hero-progress-meta span.highlight {{
            color: var(--cun-green-glow);
            font-weight: 800;
        }}

        .hero-progress-bg {{
            width: 100%;
            height: 12px;
            background: rgba(255, 255, 255, 0.08);
            border-radius: 20px;
            overflow: hidden;
            position: relative;
        }}

        .hero-progress-fill {{
            height: 100%;
            width: {pct_asistencia_global}%;
            min-width: 2%;
            background: linear-gradient(90deg, #00A859, #00FF87, #00F2FE);
            background-size: 200% 100%;
            border-radius: 20px;
            animation: glowGradient 3s infinite linear;
            transition: width 1.2s cubic-bezier(0.4, 0, 0.2, 1);
        }}

        @keyframes glowGradient {{
            0% {{ background-position: 0% 50%; }}
            50% {{ background-position: 100% 50%; }}
            100% {{ background-position: 0% 50%; }}
        }}

        /* SUBGRID: ASISTENCIA PRESENCIAL Y ASISTENCIA VIRTUAL (DEBAJO) */
        .kpi-subgrid {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1.5rem;
            margin-bottom: 2.5rem;
        }}

        .kpi-card {{
            background: var(--cun-card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--cun-border);
            border-radius: 20px;
            padding: 1.75rem;
            transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            position: relative;
            overflow: hidden;
        }}

        .kpi-card:hover {{
            transform: translateY(-5px);
            border-color: var(--cun-green-glow);
            box-shadow: 0 12px 40px rgba(0, 168, 89, 0.2);
        }}

        .kpi-card.clickable {{
            cursor: pointer;
        }}

        .kpi-card.clickable::after {{
            content: "Clic para desglose regional ➔";
            position: absolute;
            bottom: 10px;
            right: 16px;
            font-size: 0.72rem;
            color: var(--cun-cyan);
            opacity: 0.9;
            font-weight: 700;
        }}

        .kpi-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1rem;
        }}

        .kpi-icon {{
            width: 46px;
            height: 46px;
            border-radius: 14px;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.35rem;
        }}

        .icon-total {{ background: rgba(0, 242, 254, 0.15); color: var(--cun-cyan); }}
        .icon-presencial {{ background: rgba(0, 255, 135, 0.15); color: var(--cun-green-glow); }}
        .icon-virtual {{ background: rgba(168, 85, 247, 0.15); color: #C084FC; }}

        .kpi-title {{
            font-size: 0.9rem;
            color: var(--cun-muted);
            font-weight: 600;
        }}

        .kpi-value {{
            font-size: 2.75rem;
            font-weight: 800;
            color: #FFFFFF;
            line-height: 1;
            margin-bottom: 0.5rem;
        }}

        .kpi-subtext {{
            font-size: 0.8rem;
            color: var(--cun-muted);
        }}

        /* Section Title & Search */
        .section-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1.5rem;
            flex-wrap: wrap;
            gap: 1rem;
        }}

        .section-title {{
            font-size: 1.3rem;
            font-weight: 800;
            display: flex;
            align-items: center;
            gap: 10px;
        }}

        .search-box {{
            position: relative;
            min-width: 320px;
        }}

        .search-box input {{
            width: 100%;
            background: rgba(15, 23, 42, 0.85);
            border: 1px solid var(--cun-border);
            padding: 0.75rem 1rem 0.75rem 2.5rem;
            border-radius: 12px;
            color: #FFFFFF;
            font-size: 0.9rem;
            outline: none;
            transition: all 0.3s ease;
        }}

        .search-box input:focus {{
            border-color: var(--cun-cyan);
            box-shadow: 0 0 15px rgba(0, 242, 254, 0.25);
        }}

        .search-box i {{
            position: absolute;
            left: 1rem;
            top: 50%;
            transform: translateY(-50%);
            color: var(--cun-muted);
        }}

        /* LAYOUT DIVIDIDO IZQUIERDA Y DERECHA POR ESCUELA */
        .escuelas-grid-layout {{
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 1.75rem;
        }}

        .escuelas-panel {{
            background: var(--cun-card-bg);
            backdrop-filter: blur(16px);
            border: 1px solid var(--cun-border);
            border-radius: 24px;
            padding: 1.5rem;
            box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37);
        }}

        .panel-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-bottom: 1rem;
            margin-bottom: 1rem;
            border-bottom: 1px solid rgba(255, 255, 255, 0.1);
        }}

        .panel-header h3 {{
            font-size: 1.1rem;
            font-weight: 700;
            color: var(--cun-green-glow);
            display: flex;
            align-items: center;
            gap: 10px;
        }}

        .panel-header span {{
            font-size: 0.8rem;
            color: var(--cun-muted);
            font-weight: 600;
        }}

        .escuelas-list {{
            display: flex;
            flex-direction: column;
            gap: 1rem;
            max-height: 520px;
            overflow-y: auto;
            padding-right: 6px;
        }}

        /* Item Panel Izquierdo: Totales por Escuela */
        .escuela-card-left {{
            background: rgba(15, 23, 42, 0.7);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 16px;
            padding: 1.1rem;
            transition: all 0.2s ease;
        }}

        .escuela-card-left:hover {{
            border-color: var(--cun-green);
            background: rgba(0, 168, 89, 0.08);
        }}

        .card-left-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        }}

        .card-left-title {{
            font-size: 0.95rem;
            font-weight: 700;
            color: #FFFFFF;
        }}

        .card-left-badge {{
            background: linear-gradient(135deg, var(--cun-green), var(--cun-cyan));
            color: #050B14;
            font-weight: 800;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85rem;
        }}

        /* Item Panel Derecho: Virtual vs Presencial por Escuela */
        .escuela-card-right {{
            background: rgba(15, 23, 42, 0.7);
            border: 1px solid rgba(255, 255, 255, 0.08);
            border-radius: 16px;
            padding: 1.1rem;
            transition: all 0.2s ease;
        }}

        .escuela-card-right:hover {{
            border-color: var(--cun-cyan);
            background: rgba(0, 242, 254, 0.06);
        }}

        .card-right-header {{
            font-size: 0.95rem;
            font-weight: 700;
            color: #FFFFFF;
            margin-bottom: 10px;
        }}

        .modalidad-breakdown {{
            display: flex;
            gap: 12px;
        }}

        .breakdown-box {{
            flex: 1;
            padding: 8px 12px;
            border-radius: 12px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }}

        .box-presencial {{
            background: rgba(0, 255, 135, 0.12);
            border: 1px solid rgba(0, 255, 135, 0.3);
            color: var(--cun-green-glow);
        }}

        .box-virtual {{
            background: rgba(168, 85, 247, 0.12);
            border: 1px solid rgba(168, 85, 247, 0.3);
            color: #C084FC;
        }}

        .box-label {{
            font-size: 0.8rem;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 6px;
        }}

        .box-value {{
            font-size: 1.1rem;
            font-weight: 800;
        }}

        /* Modal Overlay */
        .modal-overlay {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            background: rgba(5, 11, 20, 0.85);
            backdrop-filter: blur(12px);
            z-index: 999;
            display: flex;
            align-items: center;
            justify-content: center;
            opacity: 0;
            pointer-events: none;
            transition: all 0.3s ease;
        }}

        .modal-overlay.active {{
            opacity: 1;
            pointer-events: auto;
        }}

        .modal-content {{
            background: #0A1426;
            border: 1px solid var(--cun-green);
            width: 90%;
            max-width: 600px;
            border-radius: 24px;
            padding: 2rem;
            box-shadow: 0 0 50px rgba(0, 168, 89, 0.3);
            transform: scale(0.9);
            transition: transform 0.3s ease;
        }}

        .modal-overlay.active .modal-content {{
            transform: scale(1);
        }}

        .modal-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 1.5rem;
            border-bottom: 1px solid rgba(255, 255, 255, 0.1);
            padding-bottom: 1rem;
        }}

        .modal-header h3 {{
            font-size: 1.25rem;
            color: var(--cun-green-glow);
            display: flex;
            align-items: center;
            gap: 10px;
        }}

        .close-modal {{
            background: none;
            border: none;
            color: var(--cun-muted);
            font-size: 1.25rem;
            cursor: pointer;
            transition: color 0.2s;
        }}

        .close-modal:hover {{
            color: #FFFFFF;
        }}

        .regional-list {{
            display: flex;
            flex-direction: column;
            gap: 1rem;
            max-height: 400px;
            overflow-y: auto;
            padding-right: 8px;
        }}

        .regional-card {{
            background: rgba(15, 23, 42, 0.7);
            border: 1px solid rgba(255, 255, 255, 0.08);
            padding: 1rem 1.25rem;
            border-radius: 14px;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .regional-name {{
            font-weight: 700;
            font-size: 0.95rem;
            color: #FFFFFF;
        }}

        .regional-badge {{
            background: linear-gradient(135deg, var(--cun-green), var(--cun-cyan));
            color: #050B14;
            font-weight: 800;
            padding: 6px 16px;
            border-radius: 20px;
            font-size: 0.9rem;
        }}

        @media (max-width: 992px) {{
            .escuelas-grid-layout, .kpi-subgrid {{
                grid-template-columns: 1fr;
            }}
        }}
    </style>
</head>
<body>

    <canvas id="galaxy-canvas"></canvas>

    <div class="dashboard-container">
        <header>
            <div class="brand-section">
                <div class="brand-logo-badge">
                    <i class="fa-solid fa-graduation-cap"></i>
                </div>
                <div class="brand-title">
                    <h1>Bienvenida Facilitadores CUN</h1>
                    <p>Tablero Dinámico de Asistencia Válida</p>
                </div>
            </div>
            <div class="header-actions">
                <div class="live-badge">
                    <div class="live-dot"></div>
                    REPORTE EN VIVO
                </div>
            </div>
        </header>

        <!-- HERO CARD PRINCIPAL: TOTAL REGISTROS ÚNICOS VÁLIDOS (CON PORCENTAJE DE ASISTENCIA VS PLANTA TOTAL) -->
        <div class="kpi-hero-card">
            <div class="hero-header">
                <div class="hero-info">
                    <span class="kpi-title">TOTAL REGISTROS ÚNICOS VÁLIDOS</span>
                    <div class="kpi-value" id="kpi-total">{total_registros}</div>
                    <div class="kpi-subtext">Facilitadores validados en Planta (excluye duplicados)</div>
                </div>
                <div class="kpi-icon icon-total"><i class="fa-solid fa-users"></i></div>
            </div>
            <div class="hero-progress-wrapper">
                <div class="hero-progress-meta">
                    <span>PORCENTAJE DE ASISTENCIA SOBRE LA PLANTA TOTAL ({total_registros} / {total_planta} FACILITADORES)</span>
                    <span class="highlight">{pct_asistencia_global}%</span>
                </div>
                <div class="hero-progress-bg">
                    <div class="hero-progress-fill"></div>
                </div>
            </div>
        </div>

        <!-- SUBGRID: ASISTENCIA PRESENCIAL Y ASISTENCIA VIRTUAL (DEBAJO) -->
        <div class="kpi-subgrid">
            <div class="kpi-card clickable" id="btn-presencial">
                <div class="kpi-header">
                    <span class="kpi-title">ASISTENCIA PRESENCIAL</span>
                    <div class="kpi-icon icon-presencial"><i class="fa-solid fa-building-user"></i></div>
                </div>
                <div class="kpi-value" style="color: var(--cun-green-glow);">{presencial_count}</div>
                <div class="kpi-subtext">Asistentes presenciales validados</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-header">
                    <span class="kpi-title">ASISTENCIA VIRTUAL</span>
                    <div class="kpi-icon icon-virtual"><i class="fa-solid fa-laptop-code"></i></div>
                </div>
                <div class="kpi-value" style="color: #C084FC;">{virtual_count}</div>
                <div class="kpi-subtext">Conexiones virtuales validadas</div>
            </div>
        </div>

        <div class="section-header">
            <div class="section-title">
                <i class="fa-solid fa-chart-pie" style="color: var(--cun-green-glow);"></i>
                Análisis de Asistencia por Escuela / Centro de Costo
            </div>
            <div class="search-box">
                <i class="fa-solid fa-magnifying-glass"></i>
                <input type="text" id="school-search" placeholder="Buscar escuela o centro de costo...">
            </div>
        </div>

        <!-- LAYOUT DE DOS COLUMNAS IZQUIERDA Y DERECHA POR ESCUELA -->
        <div class="escuelas-grid-layout">
            <!-- COLUMNA IZQUIERDA: TOTALES POR ESCUELA -->
            <div class="escuelas-panel">
                <div class="panel-header">
                    <h3><i class="fa-solid fa-school"></i> Total por Escuela</h3>
                    <span>Participación (%)</span>
                </div>
                <div class="escuelas-list" id="escuelas-totales-list">
                </div>
            </div>

            <!-- COLUMNA DERECHA: DESGLOSE PRESENCIAL VS VIRTUAL POR ESCUELA -->
            <div class="escuelas-panel">
                <div class="panel-header">
                    <h3><i class="fa-solid fa-layer-group"></i> Conteo Virtual vs Presencial</h3>
                    <span>Modalidad por Escuela</span>
                </div>
                <div class="escuelas-list" id="escuelas-modalidades-list">
                </div>
            </div>
        </div>
    </div>

    <!-- MODAL EMPRESARIAL PRESENCIAL POR REGIONAL -->
    <div class="modal-overlay" id="modal-presencial">
        <div class="modal-content">
            <div class="modal-header">
                <h3><i class="fa-solid fa-map-location-dot"></i> Desglose Presencial por Regiones</h3>
                <button class="close-modal" id="close-modal"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="regional-list" id="regional-list-content">
            </div>
        </div>
    </div>

    <script>
        const escuelasData = {json.dumps(escuelas_data, ensure_ascii=False)};
        const regionalPresencialData = {json.dumps(regional_presencial_dict, ensure_ascii=False)};

        const escuelasTotalesContainer = document.getElementById('escuelas-totales-list');
        const escuelasModalidadesContainer = document.getElementById('escuelas-modalidades-list');
        const searchInput = document.getElementById('school-search');

        function renderEscuelas(filterText = '') {{
            escuelasTotalesContainer.innerHTML = '';
            escuelasModalidadesContainer.innerHTML = '';

            const filtered = escuelasData.filter(e => e.escuela.toLowerCase().includes(filterText.toLowerCase()));

            if (filtered.length === 0) {{
                const emptyMsg = '<div style="text-align: center; padding: 2rem; color: var(--cun-muted);">No hay coincidencias en la asistencia</div>';
                escuelasTotalesContainer.innerHTML = emptyMsg;
                escuelasModalidadesContainer.innerHTML = emptyMsg;
                return;
            }}

            filtered.forEach(item => {{
                // Card Columna Izquierda: Total por Escuela
                const cardLeft = document.createElement('div');
                cardLeft.className = 'escuela-card-left';
                cardLeft.innerHTML = `
                    <div class="card-left-header">
                        <span class="card-left-title">${{item.escuela}}</span>
                        <span class="card-left-badge">${{item.total}} Total</span>
                    </div>
                    <div class="progress-bar-bg">
                        <div class="progress-bar-fill" style="width: ${{item.porcentaje}}%;"></div>
                    </div>
                `;
                escuelasTotalesContainer.appendChild(cardLeft);

                // Card Columna Derecha: Conteo Virtual vs Presencial por Escuela
                const cardRight = document.createElement('div');
                cardRight.className = 'escuela-card-right';
                cardRight.innerHTML = `
                    <div class="card-right-header">${{item.escuela}}</div>
                    <div class="modalidad-breakdown">
                        <div class="breakdown-box box-presencial">
                            <span class="box-label"><i class="fa-solid fa-location-dot"></i> Presenciales</span>
                            <span class="box-value">${{item.presencial}}</span>
                        </div>
                        <div class="breakdown-box box-virtual">
                            <span class="box-label"><i class="fa-solid fa-globe"></i> Virtuales</span>
                            <span class="box-value">${{item.virtual}}</span>
                        </div>
                    </div>
                `;
                escuelasModalidadesContainer.appendChild(cardRight);
            }});
        }}

        renderEscuelas();

        searchInput.addEventListener('input', (e) => {{
            renderEscuelas(e.target.value);
        }});

        const modalOverlay = document.getElementById('modal-presencial');
        const btnPresencial = document.getElementById('btn-presencial');
        const btnCloseModal = document.getElementById('close-modal');
        const regionalListContent = document.getElementById('regional-list-content');

        btnPresencial.addEventListener('click', () => {{
            regionalListContent.innerHTML = '';
            const keys = Object.keys(regionalPresencialData);
            if (keys.length === 0) {{
                regionalListContent.innerHTML = '<div style="text-align:center; color:var(--cun-muted); padding: 1.5rem;">No hay registros presenciales válidos asignados a regiones</div>';
            }} else {{
                keys.forEach(reg => {{
                    const card = document.createElement('div');
                    card.className = 'regional-card';
                    card.innerHTML = `
                        <span class="regional-name">${{reg}}</span>
                        <span class="regional-badge">${{regionalPresencialData[reg]}} Facilitadores</span>
                    `;
                    regionalListContent.appendChild(card);
                }});
            }}
            modalOverlay.classList.add('active');
        }});

        btnCloseModal.addEventListener('click', () => {{
            modalOverlay.classList.remove('active');
        }});

        modalOverlay.addEventListener('click', (e) => {{
            if (e.target === modalOverlay) {{
                modalOverlay.classList.remove('active');
            }}
        }});

        const canvas = document.getElementById('galaxy-canvas');
        const ctx = canvas.getContext('2d');

        function resizeCanvas() {{
            canvas.width = window.innerWidth;
            canvas.height = window.innerHeight;
        }}
        resizeCanvas();
        window.addEventListener('resize', resizeCanvas);

        const stars = [];
        const numStars = 180;

        for (let i = 0; i < numStars; i++) {{
            stars.push({{
                x: Math.random() * canvas.width,
                y: Math.random() * canvas.height,
                radius: Math.random() * 1.8 + 0.2,
                color: Math.random() > 0.4 ? '#00A859' : (Math.random() > 0.5 ? '#00F2FE' : '#FFFFFF'),
                alpha: Math.random() * 0.8 + 0.2,
                vx: (Math.random() - 0.5) * 0.2,
                vy: (Math.random() - 0.5) * 0.2
            }});
        }}

        function drawGalaxy() {{
            ctx.clearRect(0, 0, canvas.width, canvas.height);

            const grad1 = ctx.createRadialGradient(canvas.width*0.3, canvas.height*0.3, 10, canvas.width*0.3, canvas.height*0.3, 500);
            grad1.addColorStop(0, 'rgba(0, 168, 89, 0.08)');
            grad1.addColorStop(1, 'transparent');
            ctx.fillStyle = grad1;
            ctx.fillRect(0, 0, canvas.width, canvas.height);

            const grad2 = ctx.createRadialGradient(canvas.width*0.7, canvas.height*0.7, 10, canvas.width*0.7, canvas.height*0.7, 600);
            grad2.addColorStop(0, 'rgba(0, 242, 254, 0.06)');
            grad2.addColorStop(1, 'transparent');
            ctx.fillStyle = grad2;
            ctx.fillRect(0, 0, canvas.width, canvas.height);

            stars.forEach(star => {{
                ctx.beginPath();
                ctx.arc(star.x, star.y, star.radius, 0, Math.PI * 2);
                ctx.fillStyle = star.color;
                ctx.globalAlpha = star.alpha;
                ctx.shadowBlur = 8;
                ctx.shadowColor = star.color;
                ctx.fill();
                ctx.globalAlpha = 1.0;
                ctx.shadowBlur = 0;

                star.x += star.vx;
                star.y += star.vy;

                if (star.x < 0) star.x = canvas.width;
                if (star.x > canvas.width) star.x = 0;
                if (star.y < 0) star.y = canvas.height;
                if (star.y > canvas.height) star.y = 0;
            }});

            requestAnimationFrame(drawGalaxy);
        }}

        drawGalaxy();
    </script>
</body>
</html>
"""

with open(HTML_OUTPUT, 'w', encoding='utf-8') as f:
    f.write(html_content)

with open(HTML_RESPALDO, 'w', encoding='utf-8') as f:
    f.write(html_content)

print(f"      Tablero HTML guardado en: {HTML_OUTPUT}")
print(f"      Tablero HTML de Respaldo guardado en: {HTML_RESPALDO}")

print("\n=== PROCESO FINALIZADO CON ÉXITO ===")
