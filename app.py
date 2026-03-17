import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
import re
import os
import subprocess
import tempfile
import zipfile
from io import BytesIO
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

st.set_page_config(
    page_title="Anexo ANMAT Natura",
    page_icon="📋",
    layout="wide"
)

# ─────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────
st.markdown("""
<style>
    .header-box {
        background: linear-gradient(135deg, #e8f4f8 0%, #d0eaf5 100%);
        border-left: 5px solid #00b4d8;
        border-radius: 8px;
        padding: 20px 28px;
        margin-bottom: 24px;
    }
    .header-box h1 { color: #00b4d8; font-size: 1.6rem; font-weight: 700; margin: 0 0 4px 0; }
    .header-box p { color: #444; font-size: 0.9rem; margin: 0; }
    .card { background: #ffffff; border: 1px solid #dde3ea; border-radius: 8px; padding: 20px; margin-bottom: 16px; box-shadow: 0 1px 4px rgba(0,0,0,0.06); }
    .card h3 { color: #00b4d8; font-size: 0.9rem; font-weight: 700; margin: 0 0 14px 0; text-transform: uppercase; letter-spacing: 0.06em; }
    .alert-box { background: #fff5f5; border: 1px solid #ffb3b3; border-radius: 6px; padding: 12px 16px; margin: 6px 0; color: #cc0000; font-size: 0.9rem; }
    .alert-box strong { color: #990000; }
    .success-box { background: #f0fff8; border: 1px solid #00c896; border-radius: 6px; padding: 12px 16px; margin: 6px 0; color: #007a5c; font-size: 0.9rem; }
    .stat-card { background: #ffffff; border: 1px solid #dde3ea; border-radius: 8px; padding: 16px; text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,0.05); }
    .stat-card .number { color: #00b4d8; font-size: 2rem; font-weight: 700; line-height: 1; }
    .stat-card .label { color: #888; font-size: 0.78rem; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.05em; }
    .step-badge { display: inline-block; background: #00b4d8; color: #ffffff; border-radius: 50%; width: 22px; height: 22px; text-align: center; line-height: 22px; font-weight: 700; font-size: 0.78rem; margin-right: 8px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────
st.markdown("""
<div class="header-box">
    <h1>📋 Generador de Anexo ANMAT</h1>
    <p>Natura · Avon · Operaciones de importación</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────
# FUNCIONES CORE (igual que el script)
# ─────────────────────────────────────────

def limpiar_str(s):
    s = str(s).strip()
    s = unicodedata.normalize('NFD', s)
    s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
    s = s.lower()
    s = s.replace(':', ' ').replace('.', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def normalizar_pais(origen_str):
    if pd.isna(origen_str):
        return ''
    s = str(origen_str).strip()
    if ':' in s:
        return s.split(':')[0].strip().lower()
    return s.split(' ')[0].strip().lower()

@st.cache_data
def cargar_anmat(file_bytes):
    buf = BytesIO(file_bytes)
    try:
        df = pd.read_excel(buf, sheet_name='HISTORICO', header=0, engine='pyxlsb')
    except:
        buf.seek(0)
        df = pd.read_excel(buf, sheet_name='HISTORICO', header=0)
    df['CM'] = df['CM'].astype(str).str.strip()
    return df

@st.cache_data
def cargar_avon(file_bytes):
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        f.write(file_bytes)
        tmp = f.name
    return pd.read_excel(tmp, header=0)

@st.cache_data
def cargar_fabricantes(file_bytes, suffix='.xlsx'):
    buf = BytesIO(file_bytes)
    if suffix == '.xls':
        df = pd.read_excel(buf, header=1, engine='xlrd')
    else:
        df = pd.read_excel(buf, header=1)
    df.columns = ['material', 'En Historico', 'Corresponde']
    return df

@st.cache_data
def cargar_ncm(file_bytes):
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        f.write(file_bytes)
        tmp = f.name
    df = pd.read_excel(tmp, header=0)
    df['Artículo'] = df['Artículo'].astype(str).str.strip()
    return df

def cargar_pl(file_bytes):
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        f.write(file_bytes)
        tmp = f.name
    xl = pd.ExcelFile(tmp)
    rows = []
    invoice = None
    for sh in xl.sheet_names:
        df = pd.read_excel(tmp, sheet_name=sh, header=None)
        header_row = None
        for i, row in df.iterrows():
            if 'MATERIAL CODE' in str(row.values):
                header_row = i
                break
        if header_row is None:
            continue
        if not invoice:
            for i, row in df.iterrows():
                for val in row.values:
                    if 'Nº INVOICE:' in str(val) or 'N° INVOICE:' in str(val):
                        idx = list(row.values).index(val)
                        if idx + 1 < len(row.values):
                            invoice = str(row.values[idx + 1]).strip()
                        break
                if invoice:
                    break
        data = df.iloc[header_row + 2:].copy().reset_index(drop=True)
        data.columns = range(len(data.columns))
        data = data[data[1].astype(str).str.match(r'^\d{5,}$')]
        rows.append(data)
    pl = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
    return pl, invoice

def cargar_proximas(file_bytes):
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        f.write(file_bytes)
        tmp = f.name
    df = pd.read_excel(tmp, header=0)
    df['Material'] = df['Material'].astype(str).str.strip()
    return df

def buscar_anmat(mat_code, df_anmat):
    found = df_anmat[df_anmat['CM'] == str(mat_code)]
    if len(found) == 0:
        return None
    if len(found) > 1:
        found = found.sort_values('Fecha Admision', ascending=False)
    return found.iloc[0]

def buscar_avon(mat_code, df_avon):
    mat_str = str(mat_code).strip()
    col_cm = 'CM / ZPAC'
    col_fi = 'FI Code Local'
    found = df_avon[df_avon[col_cm].astype(str).str.strip() == mat_str]
    if len(found) == 0:
        found = df_avon[df_avon[col_fi].astype(str).str.strip() == mat_str]
    return found.iloc[0] if len(found) > 0 else None

def buscar_fabricante(origen_str, mat_code, df_fab):
    origen = str(origen_str).strip()
    origen_limpio = limpiar_str(origen)
    match_norm = None
    match_parcial = None
    for _, row in df_fab.iterrows():
        en_hist = str(row['En Historico']).strip()
        if not en_hist or en_hist == 'nan':
            continue
        en_hist_limpio = limpiar_str(en_hist)
        if origen == en_hist:
            return row['Corresponde'], None
        if match_norm is None and origen_limpio == en_hist_limpio:
            match_norm = row['Corresponde']
        if match_parcial is None and en_hist_limpio in origen_limpio:
            match_parcial = row['Corresponde']
    if match_norm:
        return match_norm, None
    if match_parcial:
        return match_parcial, None
    return None, f"Fabricante no encontrado para material {mat_code} (origen: {origen_str})"

def buscar_ncm(mat_code, df_ncm):
    found = df_ncm[df_ncm['Artículo'] == str(mat_code)]
    if len(found) == 0:
        return None, f"NCM no encontrado para material {mat_code}"
    return found.iloc[0]['NCM'], None

def verificar_origen_proximas(origen_anmat, mat_code, df_prox):
    prox_row = df_prox[df_prox['Material'] == str(mat_code)]
    if len(prox_row) == 0:
        return None, f"Material {mat_code} no encontrado en Próximas Importaciones"
    origen_prox = str(prox_row.iloc[0]['Origen'])
    pais_anmat = normalizar_pais(origen_anmat)
    pais_prox = normalizar_pais(origen_prox)
    if pais_anmat != pais_prox:
        return None, f"Origen no coincide para {mat_code}: ANMAT='{origen_anmat}' vs ProxImp='{origen_prox}'"
    return origen_prox, None


from datetime import datetime, timedelta

SEPARADORES_REGISTRO = [' - ', ' / ', ' | ', '\n', '; ']

def separar_registros(registro_str):
    """Detecta si hay múltiples N° de inscripción y los separa."""
    if not registro_str or registro_str == 'nan':
        return [registro_str]
    s = str(registro_str).strip()
    for sep in SEPARADORES_REGISTRO:
        if sep in s:
            partes = [p.strip() for p in s.split(sep) if p.strip()]
            if len(partes) > 1:
                return partes
    return [s]

def buscar_por_registro(nro_registro, df_anmat):
    """Busca en ANMAT por número de registro exacto en columna Registros ANMAT."""
    nro = str(nro_registro).strip()
    found = df_anmat[df_anmat['Registros ANMAT'].astype(str).str.strip() == nro]
    if len(found) == 0:
        return None, f"⚠️ Registro {nro} no encontrado en ANMAT Histórico"
    if len(found) > 1:
        return None, f"⚠️ Registro {nro} aparece más de una vez en ANMAT Histórico"
    return found.iloc[0], None

def parsear_fecha_vencimiento(expire_str):
    """Convierte MM/YYYY a datetime."""
    try:
        if not expire_str or expire_str == 'nan':
            return None
        partes = str(expire_str).strip().split('/')
        if len(partes) == 2:
            mes, anio = int(partes[0]), int(partes[1])
            return datetime(anio, mes, 1)
        return None
    except:
        return None

def verificar_vencimiento(expire_str):
    """
    Retorna: ('ok', None) | ('vencido', msg) | ('proximo', msg)
    """
    fecha = parsear_fecha_vencimiento(expire_str)
    if fecha is None:
        return 'ok', None
    hoy = datetime.now()
    limite_90 = hoy + timedelta(days=90)
    if fecha < hoy:
        return 'vencido', f"⚠️ VENCIDO: {expire_str}"
    if fecha <= limite_90:
        return 'proximo', f"⚠️ Vence próximo en 90 días: {expire_str}"
    return 'ok', None

def procesar_pl(pl, df_anmat, df_avon, df_prox, df_fab, df_ncm):
    filas = []
    alertas_excluir = []   # no encontrados en ANMAT ni Avon
    alertas_avon = []      # encontrados en Avon pero faltan datos
    alertas_generales = []

    for _, pl_row in pl.iterrows():
        mat_code = str(pl_row[1]).strip()
        cantidad = pl_row[2]
        descripcion_pl = str(pl_row[3]).strip() if pd.notna(pl_row[3]) else ''
        lot_product = str(pl_row[5]).strip() if pd.notna(pl_row[5]) else ''
        expire_date = str(pl_row[6]).strip() if pd.notna(pl_row[6]) else ''

        fila = {
            'MATERIAL': mat_code,
            'descripcion_factura': descripcion_pl,
            'Marca y Nombre del producto': '',
            'Variedades': '',
            'Presentación': '',
            'Cantidad': cantidad,
            'N° de inscripcion': '',
            'Lote': lot_product,
            'Fecha de vencimiento': expire_date,
            'Origen': '',
            'Fabricante': '',
            'Posición Arancelaria': '',
            '_alertas': [],
            '_skip': False,
            '_avon': False,
            '_necesita_completar': False,
            '_vencimiento': None,
            '_multi_registro': False,
            '_expanded': False,
        }

        # ── Verificar vencimiento ──
        estado_venc, msg_venc = verificar_vencimiento(expire_date)
        if msg_venc:
            fila['_vencimiento'] = estado_venc
            fila['_alertas'].append(msg_venc)
            alertas_generales.append(f"{mat_code} — {msg_venc}")

        anmat_row = buscar_anmat(mat_code, df_anmat)

        if anmat_row is not None:
            nombre = str(anmat_row['NOMBRE']) if pd.notna(anmat_row['NOMBRE']) else ''
            variedad = str(anmat_row['Variedad']) if pd.notna(anmat_row['Variedad']) else ''
            contenido = str(anmat_row['CONTENIDO NETO']) if pd.notna(anmat_row['CONTENIDO NETO']) else ''
            registro = str(anmat_row['Registros ANMAT']) if pd.notna(anmat_row['Registros ANMAT']) else ''
            origen = str(anmat_row['ORIGEN']) if pd.notna(anmat_row['ORIGEN']) else ''

            if 'REFIL' in descripcion_pl.upper():
                nombre = nombre + ' (REPUESTO)'

            fila['Origen'] = normalizar_pais(origen).capitalize() if origen != 'nan' else ''

            _, alerta_origen = verificar_origen_proximas(origen, mat_code, df_prox)
            if alerta_origen:
                fila['_alertas'].append(alerta_origen)
                alertas_generales.append(alerta_origen)

            fab, alerta_fab = buscar_fabricante(origen, mat_code, df_fab)
            if alerta_fab:
                fila['_alertas'].append(alerta_fab)
                alertas_generales.append(alerta_fab)
            else:
                fila['Fabricante'] = fab

            # ── Múltiples registros ──
            registros = separar_registros(registro)
            if len(registros) <= 1:
                # Fila simple
                fila['Marca y Nombre del producto'] = nombre
                fila['Variedades'] = variedad if variedad != 'nan' else ''
                fila['Presentación'] = contenido if contenido != 'nan' else ''
                fila['N° de inscripcion'] = registro if registro != 'nan' else ''
            else:
                # Fila principal con todos los registros
                fila['Marca y Nombre del producto'] = nombre
                fila['Variedades'] = variedad if variedad != 'nan' else ''
                fila['Presentación'] = contenido if contenido != 'nan' else ''
                fila['N° de inscripcion'] = registro if registro != 'nan' else ''
                fila['_multi_registro'] = True
                filas.append(fila)
                # Filas expandidas por cada registro individual
                for nro in registros:
                    anmat_nro, alerta_nro = buscar_por_registro(nro, df_anmat)
                    fila_exp = {
                        'MATERIAL': '',
                        'descripcion_factura': '',
                        'Marca y Nombre del producto': '',
                        'Variedades': '',
                        'Presentación': '',
                        'Cantidad': '',
                        'N° de inscripcion': nro,
                        'Lote': '',
                        'Fecha de vencimiento': '',
                        'Origen': '',
                        'Fabricante': '',
                        'Posición Arancelaria': '',
                        '_alertas': [],
                        '_skip': False,
                        '_avon': False,
                        '_necesita_completar': False,
                        '_expanded': True,
                    }
                    if alerta_nro:
                        fila_exp['_alertas'].append(alerta_nro)
                        alertas_generales.append(alerta_nro)
                    else:
                        n = str(anmat_nro['NOMBRE']) if pd.notna(anmat_nro['NOMBRE']) else ''
                        v = str(anmat_nro['Variedad']) if pd.notna(anmat_nro['Variedad']) else ''
                        c = str(anmat_nro['CONTENIDO NETO']) if pd.notna(anmat_nro['CONTENIDO NETO']) else ''
                        if 'REFIL' in descripcion_pl.upper():
                            n = n + ' (REPUESTO)'
                        fila_exp['Marca y Nombre del producto'] = n
                        fila_exp['Variedades'] = v if v != 'nan' else ''
                        fila_exp['Presentación'] = c if c != 'nan' else ''
                    filas.append(fila_exp)
                # Saltar el append al final del loop
                ncm, alerta_ncm = buscar_ncm(mat_code, df_ncm)
                if alerta_ncm:
                    alertas_generales.append(alerta_ncm)
                else:
                    filas[-len(registros)-1]['Posición Arancelaria'] = ncm
                continue

        else:
            avon_row = buscar_avon(mat_code, df_avon)
            if avon_row is not None:
                fila['_avon'] = True
                cm_zpac = str(avon_row.get('CM / ZPAC', '')).strip()
                fi_code = str(avon_row.get('FI Code Local', '')).strip()
                fila['MATERIAL'] = cm_zpac if cm_zpac and cm_zpac != 'nan' else fi_code

                nombre_avon = str(avon_row.get('NOMBRE DE REGISTRO DE PRODUCTO', ''))
                contenido_avon = str(avon_row.get('CONTENIDO LEGAL', ''))
                registro_avon = str(avon_row.get('Reg. SP   (Trámite#)\nARGENTINA NATURA', ''))

                if 'REFIL' in descripcion_pl.upper():
                    nombre_avon = nombre_avon + ' (REPUESTO)'

                fila['Marca y Nombre del producto'] = nombre_avon if nombre_avon != 'nan' else ''
                fila['Presentación'] = contenido_avon if contenido_avon != 'nan' else ''
                fila['N° de inscripcion'] = registro_avon if registro_avon != 'nan' else ''
                fila['Variedades'] = ''
                fila['Origen'] = ''
                fila['_necesita_completar'] = True

                fab, alerta_fab = buscar_fabricante('', mat_code, df_fab)
                fila['Fabricante'] = fab if not alerta_fab else ''

                alertas_avon.append({
                    'material': mat_code,
                    'descripcion': descripcion_pl,
                    'fila_idx': len(filas)
                })
            else:
                fila['_skip'] = True
                alertas_excluir.append({
                    'material': mat_code,
                    'descripcion': descripcion_pl,
                    'fila_idx': len(filas)
                })

        ncm, alerta_ncm = buscar_ncm(mat_code, df_ncm)
        if alerta_ncm:
            fila['_alertas'].append(alerta_ncm)
            alertas_generales.append(alerta_ncm)
            fila['Posición Arancelaria'] = ''
        else:
            fila['Posición Arancelaria'] = ncm

        filas.append(fila)

    return filas, alertas_excluir, alertas_avon, alertas_generales

def separar_anexos(filas):
    principal, difusor, kit3x1, alertas_sep = [], [], [], []
    for fila in filas:
        if fila['_skip']:
            continue
        desc = fila['descripcion_factura'].upper()
        es_difusor = 'DIFUSOR' in desc
        es_3x1 = '3X1' in desc or '3 X 1' in desc
        if es_difusor and es_3x1:
            alertas_sep.append(f"Material {fila['MATERIAL']} tiene DIFUSOR y 3X1 — verificar.")
            principal.append(fila)
        elif es_difusor:
            difusor.append(fila)
        elif es_3x1:
            kit3x1.append(fila)
        else:
            principal.append(fila)
    return principal, difusor, kit3x1, alertas_sep

COLUMNAS_SALIDA = ['MATERIAL', 'descripcion_factura', 'Marca y Nombre del producto',
                   'Variedades', 'Presentación', 'Cantidad', 'N° de inscripcion',
                   'Lote', 'Fecha de vencimiento', 'Origen', 'Fabricante', 'Posición Arancelaria']
COLUMNAS_SIN_PRIMERAS = COLUMNAS_SALIDA[2:]

ANCHOS = {'MATERIAL': 12, 'descripcion_factura': 35, 'Marca y Nombre del producto': 45,
          'Variedades': 20, 'Presentación': 12, 'Cantidad': 10, 'N° de inscripcion': 22,
          'Lote': 10, 'Fecha de vencimiento': 18, 'Origen': 14, 'Fabricante': 38,
          'Posición Arancelaria': 20}

def escribir_excel_bytes(filas, incluir_primeras_cols=True):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Anexo de Productos'
    columnas = COLUMNAS_SALIDA if incluir_primeras_cols else COLUMNAS_SIN_PRIMERAS

    ws.merge_cells(f'A1:{get_column_letter(len(columnas))}1')
    titulo = ws['A1']
    titulo.value = 'ANEXO DE PRODUCTOS'
    titulo.font = Font(name='Arial', bold=True, size=11)
    titulo.alignment = Alignment(horizontal='center', vertical='center')
    titulo.fill = PatternFill('solid', start_color='D9D9D9')

    header_fill = PatternFill('solid', start_color='70AD47')
    for col_idx, col_name in enumerate(columnas, 1):
        cell = ws.cell(row=2, column=col_idx, value=col_name)
        cell.font = Font(name='Arial', bold=True, size=11, color='FFFFFF')
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    alerta_fill = PatternFill('solid', start_color='FFEB9C')
    for row_idx, fila in enumerate(filas, 3):
        tiene_alerta = len(fila.get('_alertas', [])) > 0 or fila.get('_necesita_completar', False)
        for col_idx, col_name in enumerate(columnas, 1):
            val = fila.get(col_name, '')
            # MATERIAL y Cantidad como número
            if col_name in ('MATERIAL', 'Cantidad') and val != '':
                try:
                    val = int(float(str(val)))
                except:
                    pass
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name='Calibri', size=11)
            cell.alignment = Alignment(vertical='center', wrap_text=True)
            if tiene_alerta:
                cell.fill = alerta_fill

    for col_idx, col_name in enumerate(columnas, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = ANCHOS.get(col_name, 15)

    ws.row_dimensions[2].height = 30
    ws.freeze_panes = 'A3'

    if not incluir_primeras_cols:
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

def excel_a_pdf_bytes(excel_bytes, nombre_base):
    """Convierte Excel a PDF usando openpyxl + reportlab."""
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
        from reportlab.lib.styles import getSampleStyleSheet
        from reportlab.lib import colors
        from reportlab.lib.units import cm

        # Leer datos del excel
        wb = load_workbook(BytesIO(excel_bytes))
        ws = wb.active
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append([str(v) if v is not None else '' for v in row])

        buf = BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=landscape(A4),
                                leftMargin=0.8*cm, rightMargin=0.8*cm,
                                topMargin=1*cm, bottomMargin=1*cm)
        styles = getSampleStyleSheet()

        if not data:
            return None

        # Anchos proporcionales por columna (landscape A4 = ~27.7cm util)
        ancho_total = landscape(A4)[0] - 1.6*cm
        # Pesos relativos por posicion de columna en COLUMNAS_SIN_PRIMERAS
        # Marca y Nombre | Variedades | Presentacion | Cantidad | N inscripcion | Lote | Fecha venc | Origen | Fabricante | NCM
        pesos = [4.5, 2.0, 1.2, 0.9, 2.2, 1.0, 1.5, 1.2, 3.5, 2.0]
        n_cols = len(data[0])
        if n_cols == len(pesos):
            total_pesos = sum(pesos)
            col_widths = [ancho_total * p / total_pesos for p in pesos]
        else:
            col_widths = [ancho_total / n_cols] * n_cols

        table = Table(data, colWidths=col_widths, repeatRows=2)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#D9D9D9')),
            ('BACKGROUND', (0,1), (-1,1), colors.HexColor('#70AD47')),
            ('TEXTCOLOR', (0,1), (-1,1), colors.white),
            ('FONTNAME', (0,0), (-1,1), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7),
            ('FONTSIZE', (0,1), (-1,1), 7),
            ('FONTSIZE', (0,2), (-1,-1), 7),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'),
            ('ALIGN', (0,0), (-1,1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.3, colors.HexColor('#CCCCCC')),
            ('ROWBACKGROUNDS', (0,2), (-1,-1), [colors.white, colors.HexColor('#F5F5F5')]),
            ('LEFTPADDING', (0,0), (-1,-1), 3),
            ('RIGHTPADDING', (0,0), (-1,-1), 3),
            ('TOPPADDING', (0,0), (-1,-1), 3),
            ('BOTTOMPADDING', (0,0), (-1,-1), 3),
            ('WORDWRAP', (0,0), (-1,-1), True),
        ]))
        doc.build([table])
        buf.seek(0)
        return buf.getvalue()
    except Exception as e:
        return None

def generar_zip(grupos, invoice):
    buf = BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        for nombre, filas in grupos:
            if not filas:
                continue
            nombre_base = f'ANEXO_{nombre}_{invoice}'
            # Excel completo
            xls_completo = escribir_excel_bytes(filas, incluir_primeras_cols=True)
            zf.writestr(f'{nombre_base}.xlsx', xls_completo)
            # Excel sin primeras cols
            xls_sin = escribir_excel_bytes(filas, incluir_primeras_cols=False)
            zf.writestr(f'{nombre_base}_SIN_MAT.xlsx', xls_sin)
            # PDF
            pdf = excel_a_pdf_bytes(xls_sin, f'{nombre_base}_SIN_MAT')
            if pdf:
                zf.writestr(f'{nombre_base}_SIN_MAT.pdf', pdf)
    buf.seek(0)
    return buf.getvalue()

# ─────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────
if 'filas_procesadas' not in st.session_state:
    st.session_state.filas_procesadas = None
if 'alertas_excluir' not in st.session_state:
    st.session_state.alertas_excluir = []
if 'alertas_avon' not in st.session_state:
    st.session_state.alertas_avon = []
if 'alertas_generales' not in st.session_state:
    st.session_state.alertas_generales = []
if 'invoice' not in st.session_state:
    st.session_state.invoice = None
if 'excluidos' not in st.session_state:
    st.session_state.excluidos = set()
if 'datos_avon_completados' not in st.session_state:
    st.session_state.datos_avon_completados = {}

# ─────────────────────────────────────────
# PASO 1: CARGA DE ARCHIVOS
# ─────────────────────────────────────────
st.markdown('<div class="card"><h3><span class="step-badge">1</span>Archivos de la operación</h3>', unsafe_allow_html=True)

st.markdown("**📌 Número de referencia de la operación**")
nro_referencia = st.text_input("", placeholder="ej: 4550595912", label_visibility="collapsed")

col1, col2 = st.columns(2)
with col1:
    f_pl = st.file_uploader("📦 Packing List", type=['xlsx'], key='pl')
    f_prox = st.file_uploader("📅 Próximas Importaciones", type=['xlsx'], key='prox')
    f_anmat = st.file_uploader("🏥 Registro ANMAT Histórico", type=['xlsb', 'xlsx'], key='anmat')
with col2:
    f_avon = st.file_uploader("🌸 Registros Avon", type=['xlsx'], key='avon')
    f_fab = st.file_uploader("🏭 Fabricantes", type=['xls', 'xlsx'], key='fab')
    f_ncm = st.file_uploader("📊 Catálogo NCM", type=['xlsx'], key='ncm')

st.markdown('</div>', unsafe_allow_html=True)

archivos_ok = all([f_pl, f_prox, f_anmat, f_avon, f_fab, f_ncm])

# ─────────────────────────────────────────
# PASO 2: PROCESAR
# ─────────────────────────────────────────
if archivos_ok:
    st.markdown('<div class="card"><h3><span class="step-badge">2</span>Procesar operación</h3>', unsafe_allow_html=True)
    
    if st.button("⚙️ Analizar y procesar", key='btn_procesar'):
        with st.spinner('Procesando...'):
            try:
                suffix_fab = '.xls' if f_fab.name.endswith('.xls') else '.xlsx'
                pl, invoice = cargar_pl(f_pl.read())
                df_prox = cargar_proximas(f_prox.read())
                df_anmat = cargar_anmat(f_anmat.read())
                df_avon = cargar_avon(f_avon.read())
                df_fab = cargar_fabricantes(f_fab.read(), suffix=suffix_fab)
                df_ncm = cargar_ncm(f_ncm.read())

                filas, alertas_excluir, alertas_avon, alertas_generales = procesar_pl(
                    pl, df_anmat, df_avon, df_prox, df_fab, df_ncm
                )

                st.session_state.filas_procesadas = filas
                st.session_state.alertas_excluir = alertas_excluir
                st.session_state.alertas_avon = alertas_avon
                st.session_state.alertas_generales = alertas_generales
                st.session_state.invoice = invoice
                st.session_state.excluidos = set()
                st.session_state.datos_avon_completados = {}

            except Exception as e:
                st.error(f"Error al procesar: {e}")

    st.markdown('</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────
# PASO 3: ALERTAS Y RESOLUCIÓN
# ─────────────────────────────────────────
if st.session_state.filas_procesadas is not None:
    filas = st.session_state.filas_procesadas
    invoice = st.session_state.invoice

    # Stats
    total = len(filas)
    skip = len(st.session_state.alertas_excluir)
    avon = len(st.session_state.alertas_avon)
    generales = len(st.session_state.alertas_generales)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f'<div class="stat-card"><div class="number">{total}</div><div class="label">Ítems PL</div></div>', unsafe_allow_html=True)
    with col2:
        st.markdown(f'<div class="stat-card"><div class="number" style="color:#00c896">{total - skip}</div><div class="label">A procesar</div></div>', unsafe_allow_html=True)
    with col3:
        st.markdown(f'<div class="stat-card"><div class="number" style="color:#ffd166">{avon}</div><div class="label">Avon / completar</div></div>', unsafe_allow_html=True)
    with col4:
        st.markdown(f'<div class="stat-card"><div class="number" style="color:#ff6b6b">{skip}</div><div class="label">No encontrados</div></div>', unsafe_allow_html=True)

    st.markdown('<br>', unsafe_allow_html=True)

    # ── Alertas: vencimientos ──
    alertas_venc = [f for f in st.session_state.filas_procesadas if f.get('_vencimiento') in ('vencido', 'proximo') and not f['_skip']]
    if alertas_venc:
        st.markdown('<div class="card"><h3>⏰ Alertas de vencimiento</h3>', unsafe_allow_html=True)
        for fila in alertas_venc:
            tipo = fila.get('_vencimiento')
            color = '#cc0000' if tipo == 'vencido' else '#cc7700'
            icono = '🔴' if tipo == 'vencido' else '🟡'
            msg = [a for a in fila['_alertas'] if 'venc' in a.lower() or 'Venc' in a]
            msg_str = msg[0] if msg else ''
            key_excl = f'venc_{fila["MATERIAL"]}_{fila["Lote"]}'
            excluido = key_excl in st.session_state.excluidos
            col1, col2 = st.columns([4, 1])
            with col1:
                st.markdown(f'<div class="alert-box" style="border-color:{color}; color:{color};">{icono} <strong>{fila["MATERIAL"]}</strong> — Lote {fila["Lote"]} — {msg_str}</div>', unsafe_allow_html=True)
            with col2:
                label = "✅ Excluido" if excluido else "Excluir"
                if st.button(label, key=f'btn_{key_excl}'):
                    st.session_state.excluidos.add(key_excl)
                    st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Alertas: no encontrados en ANMAT ni Avon ──
    if st.session_state.alertas_excluir:
        st.markdown('<div class="card"><h3>⚠️ No encontrados en ANMAT ni Avon — ¿excluir del anexo?</h3>', unsafe_allow_html=True)
        for item in st.session_state.alertas_excluir:
            col1, col2 = st.columns([4, 1])
            with col1:
                st.markdown(f'<div class="alert-box"><strong>{item["material"]}</strong> — {item["descripcion"]}</div>', unsafe_allow_html=True)
            with col2:
                excluido = item['material'] in st.session_state.excluidos
                label = "✅ Excluido" if excluido else "Excluir del Anexo"
                if st.button(label, key=f'excl_{item["material"]}'):
                    st.session_state.excluidos.add(item['material'])
                    st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Alertas: Avon sin fabricante/origen ──
    if st.session_state.alertas_avon:
        st.markdown('<div class="card"><h3>🌸 Ítems Avon — completar Fabricante y Origen</h3>', unsafe_allow_html=True)
        for item in st.session_state.alertas_avon:
            mat = item['material']
            st.markdown(f'<div class="alert-box"><strong>{mat}</strong> — {item["descripcion"]}</div>', unsafe_allow_html=True)
            prev = st.session_state.datos_avon_completados.get(mat, {})
            c1, c2, c3 = st.columns(3)
            with c1:
                fab_val = st.text_input("Fabricante", value=prev.get('fabricante', ''), key=f'fab_{mat}')
            with c2:
                orig_val = st.text_input("Origen", value=prev.get('origen', ''), key=f'orig_{mat}')
            with c3:
                var_val = st.text_input("Variedad", value=prev.get('variedad', ''), key=f'var_{mat}')
            if fab_val or orig_val or var_val:
                st.session_state.datos_avon_completados[mat] = {'fabricante': fab_val, 'origen': orig_val, 'variedad': var_val}
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Alertas generales ──
    if st.session_state.alertas_generales:
        with st.expander(f"⚠️ {len(st.session_state.alertas_generales)} alertas adicionales"):
            for a in st.session_state.alertas_generales:
                st.markdown(f'<div class="alert-box">{a}</div>', unsafe_allow_html=True)

    # ─────────────────────────────────────────
    # PASO 4: GENERAR ANEXO
    # ─────────────────────────────────────────
    st.markdown('<div class="card"><h3><span class="step-badge">4</span>Generar Anexo</h3>', unsafe_allow_html=True)

    if st.button("📄 Generar Anexo completo", key='btn_generar'):
        with st.spinner('Generando archivos...'):
            # Aplicar datos completados por usuario
            filas_final = []
            for fila in filas:
                mat = fila['MATERIAL']
                key_venc = f'{mat}_{fila.get("Lote","")}'
                if mat in st.session_state.excluidos:
                    continue
                if f'venc_{key_venc}' in st.session_state.excluidos:
                    continue
                f = fila.copy()
                if fila.get('_avon') and mat in st.session_state.datos_avon_completados:
                    datos = st.session_state.datos_avon_completados[mat]
                    if datos.get('fabricante'):
                        f['Fabricante'] = datos['fabricante']
                    if datos.get('origen'):
                        f['Origen'] = datos['origen']
                    if datos.get('variedad'):
                        f['Variedades'] = datos['variedad']
                filas_final.append(f)

            principal, difusor, kit3x1, _ = separar_anexos(filas_final)
            grupos = [('PRINCIPAL', principal), ('DIFUSOR', difusor), ('3x1', kit3x1)]
            ref = nro_referencia.strip() if nro_referencia.strip() else invoice
            zip_bytes = generar_zip(grupos, ref)

            st.markdown('<div class="success-box">✅ Anexo generado correctamente</div>', unsafe_allow_html=True)

            resumen = []
            for nombre, filas_g in grupos:
                if filas_g:
                    resumen.append(f"**{nombre}**: {len(filas_g)} ítems")
            st.markdown(' · '.join(resumen))

            st.download_button(
                label="⬇️ Descargar todos los archivos (ZIP)",
                data=zip_bytes,
                file_name=f"ANEXO_{ref}.zip",
                mime="application/zip",
                key='dl_zip'
            )

    st.markdown('</div>', unsafe_allow_html=True)
