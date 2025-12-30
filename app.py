"""
Asignador de CPEs a Pesadas
Match por: Patente + Fecha
"""

from flask import Flask, render_template, jsonify
from flask_cors import CORS
import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import os
import json
import re
from datetime import datetime

app = Flask(__name__)
CORS(app)

# Configuración de Spreadsheets
CONFIG = {
    # Spreadsheet de Cartas de Porte Afip (y Fletes facturados)
    'CPE_SPREADSHEET_ID': '1aSZalfUpSFHytq9sYEkzDvXqFC_nBF_9a99kg6qZSXc',
    'CPE_SHEET_NAME': 'Cartas de Porte Afip',

    # Columnas en CPE (base 0)
    'CPE_COL_CTG': 0,             # A: ctg
    'CPE_COL_NUMERO_CPE': 1,      # B: numero_cpe
    'CPE_COL_FECHA': 2,           # C: fecha_documento
    'CPE_COL_GRANO_TIPO': 17,     # R: grano_tipo (Soja, Maiz, etc)
    'CPE_COL_PATENTE': 26,        # AA: dominios_vehiculos

    # Spreadsheet de Pesadas
    'PESADAS_SPREADSHEET_ID': '1gTvXfwOsqbbc5lxpcsh8HMoB5F3Bix0qpdNKdyY5DME',
    'PESADAS_SHEET_NAME': 'Pesadas Todos',

    # Columnas en Pesadas (base 0)
    'PESADAS_COL_FECHA': 1,       # B: Fecha
    'PESADAS_COL_PRODUCTO': 4,    # E: Producto (soja, maiz, etc)
    'PESADAS_COL_NETO': 8,        # I: Neto
    'PESADAS_COL_PATENTE': 11,    # L: Placa Camion
    'PESADAS_COL_CPE': 13,        # N: Carta de porte/Remito
    'PESADAS_COL_VERIFICADO': 19, # T: Verificado Duplicado (OK = verificado)

    # Fletes facturados todos (mismo spreadsheet que CPE)
    'FLETES_SHEET_NAME': 'Fletes facturados todos',
    'FLETES_COL_CANTIDAD': 3,     # D: Cantidad (peso facturado)
    'FLETES_COL_CTG': 4,          # E: CTG
    'FLETES_COL_CPE': 5,          # F: CPE
    'FLETES_COL_NETO_PESADAS': 16, # Q: M Pesadas todos
    'FLETES_COL_M_CPES': 17,      # R: M CPE's
    'FLETES_COL_NETO_DESCARGAS': 18, # S: M Descargas todos

    # Descargas Todos (mismo spreadsheet que CPE)
    'DESCARGAS_SHEET_NAME': 'Descargas Todos',
    'DESCARGAS_COL_CTG': 12,      # M: CTG
    'DESCARGAS_COL_PESO_NETO': 16, # Q: Peso Neto
}

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]


def get_credentials():
    """Obtiene credenciales de Google OAuth"""
    creds = None
    token_path = os.path.join(os.path.dirname(__file__), 'token.json')
    creds_path = os.path.join(os.path.dirname(__file__), 'credentials.json')

    # Intentar cargar desde variable de entorno (para Render)
    token_json = os.environ.get('GOOGLE_TOKEN_JSON')
    if token_json:
        try:
            token_data = json.loads(token_json)
            creds = Credentials.from_authorized_user_info(token_data, SCOPES)
        except Exception as e:
            print(f"Error cargando token desde env: {e}")

    # Si no hay env var, intentar cargar desde archivo
    if not creds and os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
            # Solo guardar si es archivo local
            if not token_json and os.path.exists(token_path):
                with open(token_path, 'w') as token:
                    token.write(creds.to_json())
        else:
            # Solo en desarrollo local
            if os.path.exists(creds_path):
                flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
                creds = flow.run_local_server(port=0)
                with open(token_path, 'w') as token:
                    token.write(creds.to_json())
            else:
                raise Exception("No se encontraron credenciales. Configure GOOGLE_TOKEN_JSON en Render.")

    return creds


def normalizar_patente(patente):
    """Normaliza patente eliminando espacios, guiones y puntos"""
    if not patente:
        return ''
    return str(patente).strip().upper().replace(' ', '').replace('-', '').replace('.', '')


def extraer_patentes_de_array(patente_str):
    """
    Extrae patentes de un string que puede ser:
    - Un array JSON: '["ABC123","DEF456"]'
    - Una patente simple: 'ABC 123'
    Retorna lista de patentes normalizadas
    """
    if not patente_str:
        return []

    patente_str = str(patente_str).strip()
    patentes = []

    # Intentar parsear como JSON array
    if patente_str.startswith('['):
        try:
            lista = json.loads(patente_str)
            for p in lista:
                norm = normalizar_patente(p)
                if norm:
                    patentes.append(norm)
        except json.JSONDecodeError:
            # Si falla el JSON, intentar extraer con regex
            matches = re.findall(r'"([^"]+)"', patente_str)
            for m in matches:
                norm = normalizar_patente(m)
                if norm:
                    patentes.append(norm)
    else:
        # Es una patente simple
        norm = normalizar_patente(patente_str)
        if norm:
            patentes.append(norm)

    return patentes


def normalizar_fecha(fecha):
    """Normaliza fecha a formato YYYY-MM-DD para comparación"""
    if not fecha:
        return ''

    # Si ya es datetime
    if isinstance(fecha, datetime):
        return fecha.strftime('%Y-%m-%d')

    # Intentar parsear string
    fecha_str = str(fecha).strip()

    # Probar diferentes formatos
    formatos = [
        '%Y-%m-%d',
        '%d/%m/%Y',
        '%d-%m-%Y',
        '%Y/%m/%d',
        '%d/%m/%y',
    ]

    for fmt in formatos:
        try:
            dt = datetime.strptime(fecha_str, fmt)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            continue

    return fecha_str


def parse_number(value):
    """Convierte string a float, manejando formato argentino (puntos como miles, comas como decimales)"""
    if not value:
        return None
    value_str = str(value).strip()
    if not value_str:
        return None
    # Formato argentino: 30.000,00 -> 30000.00
    value_str = value_str.replace('.', '').replace(',', '.')
    try:
        return float(value_str)
    except ValueError:
        return None


def normalizar_producto(producto):
    """Normaliza el nombre del producto para comparación (soja, maiz, trigo, etc)"""
    if not producto:
        return ''
    prod = str(producto).strip().lower()
    # Normalizar variantes comunes
    if 'soja' in prod or 'soya' in prod:
        return 'soja'
    if 'maiz' in prod or 'maíz' in prod:
        return 'maiz'
    if 'trigo' in prod:
        return 'trigo'
    if 'girasol' in prod:
        return 'girasol'
    if 'cebada' in prod:
        return 'cebada'
    if 'sorgo' in prod:
        return 'sorgo'
    return prod


def cargar_cpes(gc):
    """Carga todos los CPE con sus datos - indexado por patente, guardando TODOS los CPEs (lista)"""
    ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
    hoja = ss.worksheet(CONFIG['CPE_SHEET_NAME'])
    datos = hoja.get_all_values()

    # Diccionario: patente -> [lista de {numero_cpe, fecha, ctg, grano_tipo}]
    cpes = {}

    # Saltar encabezado
    for fila in datos[1:]:
        if len(fila) <= CONFIG['CPE_COL_PATENTE']:
            continue

        ctg = fila[CONFIG['CPE_COL_CTG']] if len(fila) > CONFIG['CPE_COL_CTG'] else ''
        numero_cpe = fila[CONFIG['CPE_COL_NUMERO_CPE']] if len(fila) > CONFIG['CPE_COL_NUMERO_CPE'] else ''
        fecha = fila[CONFIG['CPE_COL_FECHA']] if len(fila) > CONFIG['CPE_COL_FECHA'] else ''
        grano_tipo = fila[CONFIG['CPE_COL_GRANO_TIPO']] if len(fila) > CONFIG['CPE_COL_GRANO_TIPO'] else ''
        patente_raw = fila[CONFIG['CPE_COL_PATENTE']] if len(fila) > CONFIG['CPE_COL_PATENTE'] else ''

        if not numero_cpe or not patente_raw:
            continue

        fecha_norm = normalizar_fecha(fecha)
        ctg_norm = ctg.strip() if ctg else ''
        grano_norm = normalizar_producto(grano_tipo)

        # Extraer todas las patentes del array (camión + acoplado)
        patentes = extraer_patentes_de_array(patente_raw)

        # Crear una entrada por cada patente (guardando TODOS los CPEs en lista)
        cpe_data = {'numero_cpe': numero_cpe, 'fecha': fecha_norm, 'ctg': ctg_norm, 'grano_tipo': grano_norm}
        for patente_norm in patentes:
            if patente_norm:
                if patente_norm not in cpes:
                    cpes[patente_norm] = []
                # Evitar duplicados (mismo numero_cpe)
                if not any(c['numero_cpe'] == numero_cpe for c in cpes[patente_norm]):
                    cpes[patente_norm].append(cpe_data)

    return cpes


def calcular_diferencia_fechas(fecha1, fecha2):
    """Calcula la diferencia en días entre dos fechas (formato YYYY-MM-DD)"""
    if not fecha1 or not fecha2:
        return float('inf')
    try:
        dt1 = datetime.strptime(fecha1, '%Y-%m-%d')
        dt2 = datetime.strptime(fecha2, '%Y-%m-%d')
        return abs((dt1 - dt2).days)
    except:
        return float('inf')


def calcular_dias_diferencia(fecha_pesada, fecha_cpe):
    """
    Calcula días de diferencia entre pesada y CPE.
    Retorna número positivo si CPE es después de pesada (válido).
    Retorna número negativo si CPE es antes de pesada (inválido).
    """
    if not fecha_pesada or not fecha_cpe:
        return float('inf')
    try:
        dt_pesada = datetime.strptime(fecha_pesada, '%Y-%m-%d')
        dt_cpe = datetime.strptime(fecha_cpe, '%Y-%m-%d')
        return (dt_cpe - dt_pesada).days
    except:
        return float('inf')


def asignar_cpes():
    """
    Proceso principal: asigna CPEs a Pesadas.

    LÓGICA:
    1. Cargar CPEs ya asignadas (para no reutilizar - relación 1:1)
    2. Para cada pesada sin CPE:
       - Buscar CPEs con misma patente
       - Filtrar por producto (palabra clave: maiz, soja, trigo, etc)
       - Filtrar por fecha: CPE >= pesada Y CPE <= pesada + 7 días
       - Excluir CPEs ya usadas
       - Si 1 candidato: asignar
       - Si 2+: asignar más cercana + marcar REVISAR
       - Si 0: dejar vacío
    """
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)

        # Cargar CPEs (indexado por patente, lista de todos los CPEs)
        cpes = cargar_cpes(gc)
        total_patentes_cpe = len(cpes)

        # Contar CPEs duplicados (patentes con múltiples CPEs)
        patentes_duplicadas = sum(1 for lista in cpes.values() if len(lista) > 1)

        # Abrir hoja de Pesadas
        ss_pesadas = gc.open_by_key(CONFIG['PESADAS_SPREADSHEET_ID'])
        hoja_pesadas = ss_pesadas.worksheet(CONFIG['PESADAS_SHEET_NAME'])
        datos_pesadas = hoja_pesadas.get_all_values()

        # PASO 1: Cargar CPEs ya asignadas (para no reutilizar)
        cpes_ya_usadas = set()
        for fila in datos_pesadas[1:]:
            cpe_existente = fila[CONFIG['PESADAS_COL_CPE']] if len(fila) > CONFIG['PESADAS_COL_CPE'] else ''
            if cpe_existente and str(cpe_existente).strip():
                cpes_ya_usadas.add(str(cpe_existente).strip())

        # Estadísticas
        total_pesadas = 0
        ya_tenian_cpe = 0
        matches_nuevos = 0
        matches_unicos = 0              # 1 solo candidato
        matches_con_empate = 0          # 2+ candidatos, se eligió más cercano
        sin_match = 0
        fuera_de_rango = 0              # Candidatos pero fuera de 7 días

        # Lista para tracking de duplicados (para la página de control)
        duplicados_info = []

        # Batch de actualizaciones
        actualizaciones_cpe = []
        actualizaciones_revisar = []

        # PASO 2: Procesar cada pesada (saltar encabezado)
        for idx, fila in enumerate(datos_pesadas[1:], start=2):
            if len(fila) <= CONFIG['PESADAS_COL_PATENTE']:
                continue

            fecha_pesada = fila[CONFIG['PESADAS_COL_FECHA']] if len(fila) > CONFIG['PESADAS_COL_FECHA'] else ''
            producto_pesada = fila[CONFIG['PESADAS_COL_PRODUCTO']] if len(fila) > CONFIG['PESADAS_COL_PRODUCTO'] else ''
            patente_pesada = fila[CONFIG['PESADAS_COL_PATENTE']] if len(fila) > CONFIG['PESADAS_COL_PATENTE'] else ''
            cpe_actual = fila[CONFIG['PESADAS_COL_CPE']] if len(fila) > CONFIG['PESADAS_COL_CPE'] else ''
            neto_pesada = fila[CONFIG['PESADAS_COL_NETO']] if len(fila) > CONFIG['PESADAS_COL_NETO'] else ''

            # Saltar filas sin fecha o patente
            if not fecha_pesada and not patente_pesada:
                continue

            total_pesadas += 1

            # Si ya tiene CPE, no sobrescribir
            if cpe_actual and str(cpe_actual).strip():
                ya_tenian_cpe += 1
                continue

            # Normalizar datos de la pesada
            patente_norm = normalizar_patente(patente_pesada)
            fecha_pesada_norm = normalizar_fecha(fecha_pesada)
            producto_pesada_norm = normalizar_producto(producto_pesada)

            if not patente_norm or patente_norm not in cpes:
                sin_match += 1
                continue

            lista_cpes = cpes[patente_norm]

            # FILTRO 1: Excluir CPEs ya usadas
            cpes_disponibles = [c for c in lista_cpes if c['numero_cpe'] not in cpes_ya_usadas]

            if not cpes_disponibles:
                sin_match += 1
                continue

            # FILTRO 2: Filtrar por producto (palabra clave)
            cpes_mismo_producto = [c for c in cpes_disponibles if c['grano_tipo'] == producto_pesada_norm]

            # Si no hay match por producto, intentar con todos
            candidatos = cpes_mismo_producto if cpes_mismo_producto else cpes_disponibles

            # FILTRO 3: Filtrar por fecha (CPE >= pesada, máximo 7 días)
            cpes_en_rango = []
            for c in candidatos:
                dias = calcular_dias_diferencia(fecha_pesada_norm, c['fecha'])
                if 0 <= dias <= 7:  # CPE mismo día o hasta 7 días después
                    c['dias_diferencia'] = dias
                    cpes_en_rango.append(c)

            if not cpes_en_rango:
                # Hay candidatos pero fuera de rango de fechas
                if candidatos:
                    fuera_de_rango += 1
                else:
                    sin_match += 1
                continue

            # ASIGNACIÓN
            numero_cpe = None
            marcar_revisar = False

            if len(cpes_en_rango) == 1:
                # Único candidato válido
                numero_cpe = cpes_en_rango[0]['numero_cpe']
                matches_unicos += 1
            else:
                # Múltiples candidatos: ordenar por fecha más cercana
                cpes_en_rango.sort(key=lambda x: x['dias_diferencia'])
                numero_cpe = cpes_en_rango[0]['numero_cpe']

                # Solo marcar REVISAR si hay empate real (2+ CPEs con la misma fecha)
                dias_mejor = cpes_en_rango[0]['dias_diferencia']
                cpes_mismo_dia = [c for c in cpes_en_rango if c['dias_diferencia'] == dias_mejor]

                if len(cpes_mismo_dia) > 1:
                    # Empate real: múltiples CPEs con la misma fecha más cercana
                    matches_con_empate += 1
                    marcar_revisar = True

                    # Registrar info para control
                    duplicados_info.append({
                        'fila': idx,
                        'fecha_pesada': fecha_pesada,
                        'producto_pesada': producto_pesada,
                        'patente': patente_pesada,
                        'neto': neto_pesada,
                        'cpe_asignado': numero_cpe,
                        'candidatos': len(cpes_en_rango),
                        'empate_real': len(cpes_mismo_dia),
                        'opciones': [{'cpe': c['numero_cpe'], 'fecha': c['fecha'], 'dias': c['dias_diferencia']}
                                    for c in cpes_en_rango[:5]]
                    })
                else:
                    # Sin empate: hay un ganador claro (fecha más cercana)
                    matches_unicos += 1

            if numero_cpe:
                # Marcar CPE como usada (para no reutilizar en esta corrida)
                cpes_ya_usadas.add(numero_cpe)

                # Agregar actualización de CPE
                col_cpe = CONFIG['PESADAS_COL_CPE'] + 1
                actualizaciones_cpe.append({
                    'range': f'{gspread.utils.rowcol_to_a1(idx, col_cpe)}',
                    'values': [[numero_cpe]]
                })

                # Si hay empate, marcar REVISAR en columna T
                if marcar_revisar:
                    col_verificado = CONFIG['PESADAS_COL_VERIFICADO'] + 1
                    actualizaciones_revisar.append({
                        'range': f'{gspread.utils.rowcol_to_a1(idx, col_verificado)}',
                        'values': [['REVISAR']]
                    })

                matches_nuevos += 1

        # Aplicar actualizaciones en batch
        if actualizaciones_cpe:
            hoja_pesadas.batch_update(actualizaciones_cpe)

        if actualizaciones_revisar:
            hoja_pesadas.batch_update(actualizaciones_revisar)

        return {
            'success': True,
            'total_patentes_cpe': total_patentes_cpe,
            'patentes_duplicadas': patentes_duplicadas,
            'total_pesadas': total_pesadas,
            'ya_tenian_cpe': ya_tenian_cpe,
            'matches_nuevos': matches_nuevos,
            'matches_unicos': matches_unicos,
            'matches_con_empate': matches_con_empate,
            'fuera_de_rango': fuera_de_rango,
            'duplicados_info': duplicados_info,
            'sin_match': sin_match,
            'cpes_ya_usadas': len(cpes_ya_usadas)
        }

    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }


def matchear_pesadas_fletes():
    """
    Lleva el Neto de Pesadas a Fletes facturados.
    Match: Pesadas.CPE -> CPE.numero_cpe -> CPE.ctg -> Fletes.CTG
    """
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)

        # 1. Cargar mapeo CPE: numero_cpe -> ctg
        ss_cpe = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
        hoja_cpe = ss_cpe.worksheet(CONFIG['CPE_SHEET_NAME'])
        datos_cpe = hoja_cpe.get_all_values()

        cpe_a_ctg = {}  # numero_cpe -> ctg
        for fila in datos_cpe[1:]:
            if len(fila) > CONFIG['CPE_COL_NUMERO_CPE']:
                ctg = fila[CONFIG['CPE_COL_CTG']] if len(fila) > CONFIG['CPE_COL_CTG'] else ''
                numero_cpe = fila[CONFIG['CPE_COL_NUMERO_CPE']] if len(fila) > CONFIG['CPE_COL_NUMERO_CPE'] else ''
                if ctg and numero_cpe:
                    cpe_a_ctg[numero_cpe.strip()] = ctg.strip()

        # 2. Cargar Pesadas con CPE asignado: obtener Neto por CPE
        ss_pesadas = gc.open_by_key(CONFIG['PESADAS_SPREADSHEET_ID'])
        hoja_pesadas = ss_pesadas.worksheet(CONFIG['PESADAS_SHEET_NAME'])
        datos_pesadas = hoja_pesadas.get_all_values()

        pesadas_por_cpe = {}  # cpe -> neto
        for fila in datos_pesadas[1:]:
            if len(fila) > CONFIG['PESADAS_COL_CPE']:
                cpe = fila[CONFIG['PESADAS_COL_CPE']] if len(fila) > CONFIG['PESADAS_COL_CPE'] else ''
                neto = fila[CONFIG['PESADAS_COL_NETO']] if len(fila) > CONFIG['PESADAS_COL_NETO'] else ''
                if cpe and str(cpe).strip() and neto:
                    pesadas_por_cpe[str(cpe).strip()] = neto

        # 3. Cargar Fletes y buscar matches por CTG
        hoja_fletes = ss_cpe.worksheet(CONFIG['FLETES_SHEET_NAME'])
        datos_fletes = hoja_fletes.get_all_values()

        # Estadísticas
        total_fletes = 0
        ya_tenian_neto = 0
        matches_nuevos = 0
        sin_match = 0

        # Batch de actualizaciones
        actualizaciones = []

        for idx, fila in enumerate(datos_fletes[1:], start=2):
            if len(fila) <= CONFIG['FLETES_COL_CTG']:
                continue

            ctg_flete = fila[CONFIG['FLETES_COL_CTG']] if len(fila) > CONFIG['FLETES_COL_CTG'] else ''
            neto_actual = fila[CONFIG['FLETES_COL_NETO_PESADAS']] if len(fila) > CONFIG['FLETES_COL_NETO_PESADAS'] else ''

            if not ctg_flete:
                continue

            total_fletes += 1
            ctg_flete = str(ctg_flete).strip()

            # Si ya tiene neto, no sobrescribir
            if neto_actual and str(neto_actual).strip():
                ya_tenian_neto += 1
                continue

            # Buscar: CTG -> numero_cpe -> Neto de Pesadas
            neto_encontrado = None

            # Buscar qué CPE tiene este CTG
            for numero_cpe, ctg in cpe_a_ctg.items():
                if ctg == ctg_flete:
                    # Encontramos el CPE, ahora buscar el Neto en Pesadas
                    if numero_cpe in pesadas_por_cpe:
                        neto_encontrado = pesadas_por_cpe[numero_cpe]
                        break

            if neto_encontrado:
                col_neto = CONFIG['FLETES_COL_NETO_PESADAS'] + 1
                actualizaciones.append({
                    'range': f'{gspread.utils.rowcol_to_a1(idx, col_neto)}',
                    'values': [[neto_encontrado]]
                })
                matches_nuevos += 1
            else:
                sin_match += 1

        # Aplicar actualizaciones en batch
        if actualizaciones:
            hoja_fletes.batch_update(actualizaciones)

            # Aplicar formato verde a todas las celdas en un solo batch
            if len(actualizaciones) > 0:
                celdas = [act['range'] for act in actualizaciones]
                # Agrupar en rangos para formatear de una sola vez (máximo 10 por llamada)
                for i in range(0, len(celdas), 10):
                    grupo = celdas[i:i+10]
                    try:
                        hoja_fletes.batch_format([{
                            'range': celda,
                            'format': {'backgroundColor': {'red': 0.71, 'green': 0.84, 'blue': 0.66}}
                        } for celda in grupo])
                    except Exception:
                        # Si batch_format no está disponible, formatear sin color
                        pass

        return {
            'success': True,
            'total_cpes_mapeados': len(cpe_a_ctg),
            'pesadas_con_cpe': len(pesadas_por_cpe),
            'total_fletes': total_fletes,
            'ya_tenian_neto': ya_tenian_neto,
            'matches_nuevos': matches_nuevos,
            'sin_match': sin_match
        }

    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }


def matchear_descargas_fletes():
    """
    Lleva el Peso Neto de Descargas a Fletes facturados.
    Match directo por CTG
    """
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)
        ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])

        # 1. Cargar Descargas: CTG -> Peso Neto
        hoja_descargas = ss.worksheet(CONFIG['DESCARGAS_SHEET_NAME'])
        datos_descargas = hoja_descargas.get_all_values()

        descargas_por_ctg = {}  # ctg -> peso_neto
        for fila in datos_descargas[1:]:
            if len(fila) > CONFIG['DESCARGAS_COL_PESO_NETO']:
                ctg = fila[CONFIG['DESCARGAS_COL_CTG']] if len(fila) > CONFIG['DESCARGAS_COL_CTG'] else ''
                peso_neto = fila[CONFIG['DESCARGAS_COL_PESO_NETO']] if len(fila) > CONFIG['DESCARGAS_COL_PESO_NETO'] else ''
                if ctg and str(ctg).strip() and peso_neto:
                    descargas_por_ctg[str(ctg).strip()] = peso_neto

        # 2. Cargar Fletes y buscar matches por CTG
        hoja_fletes = ss.worksheet(CONFIG['FLETES_SHEET_NAME'])
        datos_fletes = hoja_fletes.get_all_values()

        # Estadísticas
        total_fletes = 0
        ya_tenian_neto = 0
        matches_nuevos = 0
        sin_match = 0

        # Batch de actualizaciones
        actualizaciones = []

        for idx, fila in enumerate(datos_fletes[1:], start=2):
            if len(fila) <= CONFIG['FLETES_COL_CTG']:
                continue

            ctg_flete = fila[CONFIG['FLETES_COL_CTG']] if len(fila) > CONFIG['FLETES_COL_CTG'] else ''
            neto_actual = fila[CONFIG['FLETES_COL_NETO_DESCARGAS']] if len(fila) > CONFIG['FLETES_COL_NETO_DESCARGAS'] else ''

            if not ctg_flete:
                continue

            total_fletes += 1
            ctg_flete = str(ctg_flete).strip()

            # Si ya tiene neto, no sobrescribir
            if neto_actual and str(neto_actual).strip():
                ya_tenian_neto += 1
                continue

            # Buscar en Descargas por CTG
            if ctg_flete in descargas_por_ctg:
                peso_neto = descargas_por_ctg[ctg_flete]
                col_neto = CONFIG['FLETES_COL_NETO_DESCARGAS'] + 1
                actualizaciones.append({
                    'range': f'{gspread.utils.rowcol_to_a1(idx, col_neto)}',
                    'values': [[peso_neto]]
                })
                matches_nuevos += 1
            else:
                sin_match += 1

        # Aplicar actualizaciones en batch
        if actualizaciones:
            try:
                hoja_fletes.batch_update(actualizaciones)
            except Exception as batch_error:
                return {
                    'success': False,
                    'error': f'Error en batch_update: {str(batch_error)}',
                    'primera_actualizacion': actualizaciones[0] if actualizaciones else None
                }

            # Aplicar formato verde a todas las celdas en un solo batch
            if len(actualizaciones) > 0:
                celdas = [act['range'] for act in actualizaciones]
                # Agrupar en rangos para formatear de una sola vez (máximo 10 por llamada)
                for i in range(0, len(celdas), 10):
                    grupo = celdas[i:i+10]
                    try:
                        hoja_fletes.batch_format([{
                            'range': celda,
                            'format': {'backgroundColor': {'red': 0.71, 'green': 0.84, 'blue': 0.66}}
                        } for celda in grupo])
                    except Exception:
                        # Si batch_format no está disponible, formatear sin color
                        pass

        return {
            'success': True,
            'descargas_con_ctg': len(descargas_por_ctg),
            'total_fletes': total_fletes,
            'ya_tenian_neto': ya_tenian_neto,
            'matches_nuevos': matches_nuevos,
            'sin_match': sin_match
        }

    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }


def traer_cpes_a_fletes():
    """
    Busca el numero_cpe en Cartas de Porte por CTG y lo trae a Fletes.
    - Columna CPE: escribe el numero_cpe
    - Columna M CPE's: escribe "si" (verde) o "no" (rojo)
    """
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)
        ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])

        # 1. Cargar CPE: CTG -> numero_cpe
        hoja_cpe = ss.worksheet(CONFIG['CPE_SHEET_NAME'])
        datos_cpe = hoja_cpe.get_all_values()

        cpe_por_ctg = {}  # ctg -> numero_cpe
        for fila in datos_cpe[1:]:
            if len(fila) > CONFIG['CPE_COL_NUMERO_CPE']:
                ctg = fila[CONFIG['CPE_COL_CTG']] if len(fila) > CONFIG['CPE_COL_CTG'] else ''
                numero_cpe = fila[CONFIG['CPE_COL_NUMERO_CPE']] if len(fila) > CONFIG['CPE_COL_NUMERO_CPE'] else ''
                if ctg and str(ctg).strip() and numero_cpe:
                    cpe_por_ctg[str(ctg).strip()] = numero_cpe

        # 2. Cargar Fletes y buscar matches por CTG
        hoja_fletes = ss.worksheet(CONFIG['FLETES_SHEET_NAME'])
        datos_fletes = hoja_fletes.get_all_values()

        # Estadísticas
        total_fletes = 0
        con_cpe = 0
        sin_cpe = 0

        # Batch de actualizaciones
        actualizaciones_cpe = []
        actualizaciones_match = []
        formatos_verde = []
        formatos_rojo = []

        for idx, fila in enumerate(datos_fletes[1:], start=2):
            if len(fila) <= CONFIG['FLETES_COL_CTG']:
                continue

            ctg_flete = fila[CONFIG['FLETES_COL_CTG']] if len(fila) > CONFIG['FLETES_COL_CTG'] else ''

            if not ctg_flete:
                continue

            total_fletes += 1
            ctg_flete = str(ctg_flete).strip()

            col_cpe = CONFIG['FLETES_COL_CPE'] + 1  # F
            col_match = CONFIG['FLETES_COL_M_CPES'] + 1  # R
            celda_cpe = gspread.utils.rowcol_to_a1(idx, col_cpe)
            celda_match = gspread.utils.rowcol_to_a1(idx, col_match)

            # Obtener valor actual de M CPE's para no sobrescribir clasificaciones manuales
            m_cpes_actual = fila[CONFIG['FLETES_COL_M_CPES']] if len(fila) > CONFIG['FLETES_COL_M_CPES'] else ''
            m_cpes_actual = str(m_cpes_actual).strip().lower()

            # Valores que NO deben sobrescribirse (clasificaciones manuales + si/no existentes)
            valores_protegidos = [
                'traslado interno', 'flete en b', 'sin documentación',
                'pendiente de cpe', 'error de carga', 'cpe hecha por terceros',
                'si', 'no'  # tampoco sobrescribir si ya tiene si o no
            ]

            # Si ya tiene un valor protegido, no tocar esta fila
            if m_cpes_actual in valores_protegidos:
                if ctg_flete in cpe_por_ctg:
                    con_cpe += 1
                else:
                    sin_cpe += 1
                continue  # Saltar esta fila completamente

            # Solo llegamos aquí si M CPE's está vacío o tiene otro valor no protegido
            # Buscar en CPE por CTG
            if ctg_flete in cpe_por_ctg:
                numero_cpe = cpe_por_ctg[ctg_flete]
                actualizaciones_cpe.append({
                    'range': celda_cpe,
                    'values': [[numero_cpe]]
                })
                actualizaciones_match.append({
                    'range': celda_match,
                    'values': [['si']]
                })
                formatos_verde.append(celda_match)
                con_cpe += 1
            else:
                # Escribir "no" solo si está vacío
                actualizaciones_match.append({
                    'range': celda_match,
                    'values': [['no']]
                })
                formatos_rojo.append(celda_match)
                sin_cpe += 1

        # Aplicar actualizaciones en batch
        todas_actualizaciones = actualizaciones_cpe + actualizaciones_match
        if todas_actualizaciones:
            try:
                hoja_fletes.batch_update(todas_actualizaciones)
            except Exception as batch_error:
                return {
                    'success': False,
                    'error': f'Error en batch_update: {str(batch_error)}'
                }

            # Aplicar formato verde a los "si" en batch
            if formatos_verde:
                for i in range(0, len(formatos_verde), 10):
                    grupo = formatos_verde[i:i+10]
                    try:
                        hoja_fletes.batch_format([{
                            'range': celda,
                            'format': {'backgroundColor': {'red': 0.71, 'green': 0.84, 'blue': 0.66}}
                        } for celda in grupo])
                    except Exception:
                        pass

            # Aplicar formato rojo a los "no" en batch
            if formatos_rojo:
                for i in range(0, len(formatos_rojo), 10):
                    grupo = formatos_rojo[i:i+10]
                    try:
                        hoja_fletes.batch_format([{
                            'range': celda,
                            'format': {'backgroundColor': {'red': 0.92, 'green': 0.6, 'blue': 0.6}}
                        } for celda in grupo])
                    except Exception:
                        pass

        return {
            'success': True,
            'cpe_disponibles': len(cpe_por_ctg),
            'total_fletes': total_fletes,
            'con_cpe': con_cpe,
            'sin_cpe': sin_cpe
        }

    except Exception as e:
        return {
            'success': False,
            'error': str(e)
        }


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/asignar', methods=['POST'])
def asignar():
    resultado = asignar_cpes()
    return jsonify(resultado)


@app.route('/matchear-fletes', methods=['POST'])
def matchear_fletes():
    resultado = matchear_pesadas_fletes()
    return jsonify(resultado)


@app.route('/matchear-descargas', methods=['POST'])
def matchear_descargas():
    resultado = matchear_descargas_fletes()
    return jsonify(resultado)


@app.route('/health')
def health():
    return jsonify({'status': 'ok'})


@app.route('/ver-descargas')
def ver_descargas():
    """Ver encabezados de Descargas Todos"""
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)
        ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
        hoja = ss.worksheet('Descargas Todos')
        datos = hoja.get_all_values()
        return jsonify({
            'encabezados': datos[0] if datos else [],
            'ejemplo_fila': datos[1] if len(datos) > 1 else []
        })
    except Exception as e:
        return jsonify({'error': str(e)})


@app.route('/ver-fletes')
def ver_fletes():
    """Ver encabezados de Fletes facturados todos"""
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)
        ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
        hoja = ss.worksheet('Fletes facturados todos')
        datos = hoja.get_all_values()
        return jsonify({
            'total_columnas': len(datos[0]) if datos else 0,
            'encabezados': datos[0] if datos else [],
            'ejemplo_fila': datos[1] if len(datos) > 1 else []
        })
    except Exception as e:
        return jsonify({'error': str(e)})


@app.route('/ver-oc-fletes')
def ver_oc_fletes():
    """Ver encabezados de OC Fletes"""
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)
        ss = gc.open_by_key('1e_GIvBUY8uskXXL7c2TsBydxprT_h36VlsLhYooz72w')
        hoja = ss.worksheet('OC Fletes')
        datos = hoja.get_all_values()
        return jsonify({
            'total_columnas': len(datos[0]) if datos else 0,
            'encabezados': datos[0] if datos else [],
            'ejemplo_filas': datos[1:6] if len(datos) > 1 else []
        })
    except Exception as e:
        return jsonify({'error': str(e)})


@app.route('/analizar')
def analizar():
    """Análisis profundo de por qué no hay matches"""
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)

        # Cargar CPEs
        ss_cpe = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
        hoja_cpe = ss_cpe.worksheet(CONFIG['CPE_SHEET_NAME'])
        datos_cpe = hoja_cpe.get_all_values()

        # Cargar Pesadas
        ss_pesadas = gc.open_by_key(CONFIG['PESADAS_SPREADSHEET_ID'])
        hoja_pesadas = ss_pesadas.worksheet(CONFIG['PESADAS_SHEET_NAME'])
        datos_pesadas = hoja_pesadas.get_all_values()

        # Crear diccionario de CPEs: {(patente, fecha): numero_cpe}
        cpes_dict = {}
        for fila in datos_cpe[1:]:
            if len(fila) > CONFIG['CPE_COL_PATENTE']:
                numero_cpe = fila[CONFIG['CPE_COL_NUMERO_CPE']] if len(fila) > CONFIG['CPE_COL_NUMERO_CPE'] else ''
                fecha = normalizar_fecha(fila[CONFIG['CPE_COL_FECHA']] if len(fila) > CONFIG['CPE_COL_FECHA'] else '')
                patente_raw = fila[CONFIG['CPE_COL_PATENTE']] if len(fila) > CONFIG['CPE_COL_PATENTE'] else ''

                if numero_cpe and fecha:
                    for p in extraer_patentes_de_array(patente_raw):
                        cpes_dict[(p, fecha)] = numero_cpe

        # Analizar primeras 20 pesadas
        analisis = []
        for fila in datos_pesadas[1:21]:
            if len(fila) > CONFIG['PESADAS_COL_PATENTE']:
                fecha_pesada = normalizar_fecha(fila[CONFIG['PESADAS_COL_FECHA']] if len(fila) > CONFIG['PESADAS_COL_FECHA'] else '')
                patente_pesada = normalizar_patente(fila[CONFIG['PESADAS_COL_PATENTE']] if len(fila) > CONFIG['PESADAS_COL_PATENTE'] else '')

                # Buscar CPEs con misma fecha
                cpes_misma_fecha = [(k, v) for k, v in cpes_dict.items() if k[1] == fecha_pesada]

                # Buscar CPEs con misma patente
                cpes_misma_patente = [(k, v) for k, v in cpes_dict.items() if k[0] == patente_pesada]

                # Match exacto
                match = cpes_dict.get((patente_pesada, fecha_pesada))

                analisis.append({
                    'fecha_pesada': fecha_pesada,
                    'patente_pesada': patente_pesada,
                    'tiene_match': match is not None,
                    'cpe_match': match,
                    'cpes_misma_fecha': len(cpes_misma_fecha),
                    'cpes_misma_patente': len(cpes_misma_patente),
                    'ejemplo_misma_fecha': cpes_misma_fecha[:3] if cpes_misma_fecha else [],
                    'ejemplo_misma_patente': cpes_misma_patente[:3] if cpes_misma_patente else [],
                })

        return jsonify({
            'total_cpes_unicas': len(cpes_dict),
            'analisis_pesadas': analisis
        })

    except Exception as e:
        return jsonify({'error': str(e)})


@app.route('/debug')
def debug():
    """Endpoint para debug - muestra ejemplos de datos"""
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)

        # Cargar algunos CPEs de ejemplo
        ss_cpe = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
        hoja_cpe = ss_cpe.worksheet(CONFIG['CPE_SHEET_NAME'])
        datos_cpe = hoja_cpe.get_all_values()

        ejemplos_cpe = []
        for fila in datos_cpe[1:6]:  # Primeras 5 filas
            if len(fila) > CONFIG['CPE_COL_PATENTE']:
                patente_raw = fila[CONFIG['CPE_COL_PATENTE']] if len(fila) > CONFIG['CPE_COL_PATENTE'] else ''
                ejemplos_cpe.append({
                    'numero_cpe': fila[CONFIG['CPE_COL_NUMERO_CPE']] if len(fila) > CONFIG['CPE_COL_NUMERO_CPE'] else '',
                    'fecha_original': fila[CONFIG['CPE_COL_FECHA']] if len(fila) > CONFIG['CPE_COL_FECHA'] else '',
                    'fecha_normalizada': normalizar_fecha(fila[CONFIG['CPE_COL_FECHA']] if len(fila) > CONFIG['CPE_COL_FECHA'] else ''),
                    'patente_original': patente_raw,
                    'patentes_extraidas': extraer_patentes_de_array(patente_raw),
                })

        # Cargar algunos Pesadas de ejemplo
        ss_pesadas = gc.open_by_key(CONFIG['PESADAS_SPREADSHEET_ID'])
        hoja_pesadas = ss_pesadas.worksheet(CONFIG['PESADAS_SHEET_NAME'])
        datos_pesadas = hoja_pesadas.get_all_values()

        ejemplos_pesadas = []
        for fila in datos_pesadas[1:6]:  # Primeras 5 filas
            if len(fila) > CONFIG['PESADAS_COL_PATENTE']:
                ejemplos_pesadas.append({
                    'fecha_original': fila[CONFIG['PESADAS_COL_FECHA']] if len(fila) > CONFIG['PESADAS_COL_FECHA'] else '',
                    'fecha_normalizada': normalizar_fecha(fila[CONFIG['PESADAS_COL_FECHA']] if len(fila) > CONFIG['PESADAS_COL_FECHA'] else ''),
                    'patente_original': fila[CONFIG['PESADAS_COL_PATENTE']] if len(fila) > CONFIG['PESADAS_COL_PATENTE'] else '',
                    'patente_normalizada': normalizar_patente(fila[CONFIG['PESADAS_COL_PATENTE']] if len(fila) > CONFIG['PESADAS_COL_PATENTE'] else ''),
                    'cpe_actual': fila[CONFIG['PESADAS_COL_CPE']] if len(fila) > CONFIG['PESADAS_COL_CPE'] else '',
                })

        # Buscar CPEs con fechas de noviembre/diciembre 2025
        cpes_recientes = []
        for fila in datos_cpe[1:]:
            if len(fila) > CONFIG['CPE_COL_FECHA']:
                fecha = fila[CONFIG['CPE_COL_FECHA']] if len(fila) > CONFIG['CPE_COL_FECHA'] else ''
                fecha_norm = normalizar_fecha(fecha)
                if fecha_norm and fecha_norm.startswith('2025-1'):  # Nov o Dec
                    patente_raw = fila[CONFIG['CPE_COL_PATENTE']] if len(fila) > CONFIG['CPE_COL_PATENTE'] else ''
                    cpes_recientes.append({
                        'numero_cpe': fila[CONFIG['CPE_COL_NUMERO_CPE']] if len(fila) > CONFIG['CPE_COL_NUMERO_CPE'] else '',
                        'fecha': fecha_norm,
                        'patentes': extraer_patentes_de_array(patente_raw)
                    })
                    if len(cpes_recientes) >= 10:
                        break

        # Contar fechas únicas en CPE
        fechas_cpe = set()
        for fila in datos_cpe[1:]:
            if len(fila) > CONFIG['CPE_COL_FECHA']:
                fecha = normalizar_fecha(fila[CONFIG['CPE_COL_FECHA']])
                if fecha:
                    fechas_cpe.add(fecha)

        return jsonify({
            'ejemplos_cpe': ejemplos_cpe,
            'ejemplos_pesadas': ejemplos_pesadas,
            'cpes_recientes_nov_dic': cpes_recientes,
            'rango_fechas_cpe': {
                'min': min(fechas_cpe) if fechas_cpe else None,
                'max': max(fechas_cpe) if fechas_cpe else None,
                'total_fechas_unicas': len(fechas_cpe)
            },
            'encabezados_cpe': datos_cpe[0] if datos_cpe else [],
            'encabezados_pesadas': datos_pesadas[0] if datos_pesadas else []
        })

    except Exception as e:
        return jsonify({'error': str(e)})


if __name__ == '__main__':
    print("=" * 50)
    print("Asignador de CPEs a Pesadas")
    print("=" * 50)
    print("Servidor corriendo en: http://localhost:5015")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5015, debug=True)
