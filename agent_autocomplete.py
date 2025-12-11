"""
Agente Autocompletador de Fletes
Busca datos faltantes en las hojas vinculadas y los completa con formato gris.
"""

import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import os
import re

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Configuración
CONFIG = {
    'CPE_SPREADSHEET_ID': '1aSZalfUpSFHytq9sYEkzDvXqFC_nBF_9a99kg6qZSXc',
    'PESADAS_SPREADSHEET_ID': '1gTvXfwOsqbbc5lxpcsh8HMoB5F3Bix0qpdNKdyY5DME',

    'FLETES_SHEET': 'Fletes facturados todos',
    'PESADAS_SHEET': 'Pesadas Todos',
    'DESCARGAS_SHEET': 'Descargas Todos',
    'CPE_SHEET': 'Cartas de Porte Afip',

    # Columnas de Fletes (índice base 0)
    'FLETES_COL_FECHA': 1,
    'FLETES_COL_PRODUCTO': 2,
    'FLETES_COL_CTG': 4,
    'FLETES_COL_CPE': 5,
    'FLETES_COL_ORIGEN': 6,
    'FLETES_COL_DESTINO': 7,
    'FLETES_COL_TRANSPORTISTA': 8,
    'FLETES_COL_CHOFER': 9,
    'FLETES_COL_M_PESADAS': 16,
    'FLETES_COL_M_DESCARGAS': 18,

    # Columnas de Pesadas (índice base 0)
    'PESADAS_COL_PRODUCTO': 4,
    'PESADAS_COL_NETO': 8,
    'PESADAS_COL_ORIGEN': 10,
    'PESADAS_COL_TRANSPORTISTA': 12,
    'PESADAS_COL_CPE': 13,
    'PESADAS_COL_CHOFER': 14,

    # Columnas de Descargas (índice base 0)
    'DESCARGAS_COL_PRODUCTO': 6,
    'DESCARGAS_COL_ORIGEN': 9,
    'DESCARGAS_COL_CPE': 11,
    'DESCARGAS_COL_CTG': 12,
    'DESCARGAS_COL_PESO_NETO': 16,
    'DESCARGAS_COL_TRANSPORTISTA': 27,

    # Columnas de CPE (índice base 0)
    'CPE_COL_CTG': 0,
    'CPE_COL_NUMERO_CPE': 1,
    'CPE_COL_TRANSPORTISTA': 14,
    'CPE_COL_CHOFER': 16,
    'CPE_COL_PRODUCTO': 17,          # grano_tipo
    'CPE_COL_ORIGEN': 19,            # localidad_origen
    'CPE_COL_DESTINO': 22,           # localidad_destino
}

# Color gris para datos autocompletados
COLOR_GRIS = {'red': 0.5, 'green': 0.5, 'blue': 0.5}


def get_credentials():
    """Obtiene credenciales de Google OAuth"""
    creds = None
    base_path = os.path.dirname(__file__)
    token_path = os.path.join(base_path, 'token.json')
    creds_path = os.path.join(base_path, 'credentials.json')

    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, SCOPES)
            creds = flow.run_local_server(port=0)

        with open(token_path, 'w') as token:
            token.write(creds.to_json())

    return creds


def normalizar_cpe(cpe):
    """Normaliza número de CPE para comparación"""
    if not cpe:
        return ''
    return str(cpe).strip().upper().replace(' ', '')


def normalizar_ctg(ctg):
    """Normaliza CTG para comparación"""
    if not ctg:
        return ''
    # Remover ceros a la izquierda y espacios
    return str(ctg).strip().lstrip('0')


def cargar_datos_pesadas(gc):
    """Carga datos de Pesadas indexados por CPE"""
    ss = gc.open_by_key(CONFIG['PESADAS_SPREADSHEET_ID'])
    hoja = ss.worksheet(CONFIG['PESADAS_SHEET'])
    datos = hoja.get_all_values()

    pesadas_por_cpe = {}
    for fila in datos[1:]:
        if len(fila) > CONFIG['PESADAS_COL_CPE']:
            cpe = normalizar_cpe(fila[CONFIG['PESADAS_COL_CPE']] if len(fila) > CONFIG['PESADAS_COL_CPE'] else '')
            if cpe:
                pesadas_por_cpe[cpe] = {
                    'neto': fila[CONFIG['PESADAS_COL_NETO']] if len(fila) > CONFIG['PESADAS_COL_NETO'] else '',
                    'origen': fila[CONFIG['PESADAS_COL_ORIGEN']] if len(fila) > CONFIG['PESADAS_COL_ORIGEN'] else '',
                    'transportista': fila[CONFIG['PESADAS_COL_TRANSPORTISTA']] if len(fila) > CONFIG['PESADAS_COL_TRANSPORTISTA'] else '',
                    'chofer': fila[CONFIG['PESADAS_COL_CHOFER']] if len(fila) > CONFIG['PESADAS_COL_CHOFER'] else '',
                    'producto': fila[CONFIG['PESADAS_COL_PRODUCTO']] if len(fila) > CONFIG['PESADAS_COL_PRODUCTO'] else '',
                }

    return pesadas_por_cpe


def cargar_datos_descargas(gc):
    """Carga datos de Descargas indexados por CTG y CPE"""
    ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
    hoja = ss.worksheet(CONFIG['DESCARGAS_SHEET'])
    datos = hoja.get_all_values()

    descargas_por_ctg = {}
    descargas_por_cpe = {}

    for fila in datos[1:]:
        ctg = normalizar_ctg(fila[CONFIG['DESCARGAS_COL_CTG']] if len(fila) > CONFIG['DESCARGAS_COL_CTG'] else '')
        cpe = normalizar_cpe(fila[CONFIG['DESCARGAS_COL_CPE']] if len(fila) > CONFIG['DESCARGAS_COL_CPE'] else '')

        registro = {
            'peso_neto': fila[CONFIG['DESCARGAS_COL_PESO_NETO']] if len(fila) > CONFIG['DESCARGAS_COL_PESO_NETO'] else '',
            'origen': fila[CONFIG['DESCARGAS_COL_ORIGEN']] if len(fila) > CONFIG['DESCARGAS_COL_ORIGEN'] else '',
            'transportista': fila[CONFIG['DESCARGAS_COL_TRANSPORTISTA']] if len(fila) > CONFIG['DESCARGAS_COL_TRANSPORTISTA'] else '',
            'producto': fila[CONFIG['DESCARGAS_COL_PRODUCTO']] if len(fila) > CONFIG['DESCARGAS_COL_PRODUCTO'] else '',
        }

        if ctg:
            descargas_por_ctg[ctg] = registro
        if cpe:
            descargas_por_cpe[cpe] = registro

    return descargas_por_ctg, descargas_por_cpe


def cargar_datos_cpe(gc):
    """Carga datos de Cartas de Porte indexados por número CPE y CTG"""
    ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
    hoja = ss.worksheet(CONFIG['CPE_SHEET'])
    datos = hoja.get_all_values()

    cpe_por_numero = {}
    cpe_por_ctg = {}

    for fila in datos[1:]:
        ctg = normalizar_ctg(fila[CONFIG['CPE_COL_CTG']] if len(fila) > CONFIG['CPE_COL_CTG'] else '')
        numero_cpe = normalizar_cpe(fila[CONFIG['CPE_COL_NUMERO_CPE']] if len(fila) > CONFIG['CPE_COL_NUMERO_CPE'] else '')

        registro = {
            'ctg': ctg,
            'numero_cpe': numero_cpe,
            'transportista': fila[CONFIG['CPE_COL_TRANSPORTISTA']] if len(fila) > CONFIG['CPE_COL_TRANSPORTISTA'] else '',
            'chofer': fila[CONFIG['CPE_COL_CHOFER']] if len(fila) > CONFIG['CPE_COL_CHOFER'] else '',
            'producto': fila[CONFIG['CPE_COL_PRODUCTO']] if len(fila) > CONFIG['CPE_COL_PRODUCTO'] else '',
            'origen': fila[CONFIG['CPE_COL_ORIGEN']] if len(fila) > CONFIG['CPE_COL_ORIGEN'] else '',
            'destino': fila[CONFIG['CPE_COL_DESTINO']] if len(fila) > CONFIG['CPE_COL_DESTINO'] else '',
        }

        if numero_cpe:
            cpe_por_numero[numero_cpe] = registro
        if ctg:
            cpe_por_ctg[ctg] = registro

    return cpe_por_numero, cpe_por_ctg


def buscar_dato(campo, ctg, cpe, pesadas, descargas_ctg, descargas_cpe, cpe_data, cpe_por_ctg):
    """
    Busca un dato faltante en las hojas vinculadas.
    Retorna el valor encontrado o None.
    PRIORIDAD: Siempre buscar primero por CTG en CPE (Cartas de Porte).
    """
    ctg_norm = normalizar_ctg(ctg)
    cpe_norm = normalizar_cpe(cpe)

    valor = None

    if campo == 'origen':
        # PRIORIDAD: CPE por CTG (localidad_origen) -> Descargas -> Pesadas
        if ctg_norm and ctg_norm in cpe_por_ctg:
            valor = cpe_por_ctg[ctg_norm].get('origen')
        if not valor and cpe_norm and cpe_norm in cpe_data:
            valor = cpe_data[cpe_norm].get('origen')
        if not valor and ctg_norm and ctg_norm in descargas_ctg:
            valor = descargas_ctg[ctg_norm].get('origen')
        if not valor and cpe_norm and cpe_norm in pesadas:
            valor = pesadas[cpe_norm].get('origen')

    elif campo == 'destino':
        # PRIORIDAD: CPE por CTG (localidad_destino)
        if ctg_norm and ctg_norm in cpe_por_ctg:
            valor = cpe_por_ctg[ctg_norm].get('destino')
        if not valor and cpe_norm and cpe_norm in cpe_data:
            valor = cpe_data[cpe_norm].get('destino')

    elif campo == 'producto':
        # PRIORIDAD: CPE por CTG (grano_tipo) -> Descargas -> Pesadas
        if ctg_norm and ctg_norm in cpe_por_ctg:
            valor = cpe_por_ctg[ctg_norm].get('producto')
        if not valor and cpe_norm and cpe_norm in cpe_data:
            valor = cpe_data[cpe_norm].get('producto')
        if not valor and ctg_norm and ctg_norm in descargas_ctg:
            valor = descargas_ctg[ctg_norm].get('producto')
        if not valor and cpe_norm and cpe_norm in pesadas:
            valor = pesadas[cpe_norm].get('producto')

    elif campo == 'chofer':
        # PRIORIDAD: CPE por CTG -> Pesadas
        if ctg_norm and ctg_norm in cpe_por_ctg:
            valor = cpe_por_ctg[ctg_norm].get('chofer')
        if not valor and cpe_norm and cpe_norm in cpe_data:
            valor = cpe_data[cpe_norm].get('chofer')
        if not valor and cpe_norm and cpe_norm in pesadas:
            valor = pesadas[cpe_norm].get('chofer')

    elif campo == 'transportista':
        # PRIORIDAD: CPE por CTG -> Pesadas -> Descargas
        if ctg_norm and ctg_norm in cpe_por_ctg:
            valor = cpe_por_ctg[ctg_norm].get('transportista')
        if not valor and cpe_norm and cpe_norm in cpe_data:
            valor = cpe_data[cpe_norm].get('transportista')
        if not valor and cpe_norm and cpe_norm in pesadas:
            valor = pesadas[cpe_norm].get('transportista')
        if not valor and ctg_norm and ctg_norm in descargas_ctg:
            valor = descargas_ctg[ctg_norm].get('transportista')

    elif campo == 'm_pesadas':
        # Solo de Pesadas (por CPE)
        if cpe_norm and cpe_norm in pesadas:
            valor = pesadas[cpe_norm].get('neto')

    elif campo == 'm_descargas':
        # Solo de Descargas (por CTG)
        if ctg_norm and ctg_norm in descargas_ctg:
            valor = descargas_ctg[ctg_norm].get('peso_neto')
        if not valor and cpe_norm and cpe_norm in descargas_cpe:
            valor = descargas_cpe[cpe_norm].get('peso_neto')

    # Limpiar valor
    if valor and str(valor).strip():
        return str(valor).strip()

    return None


def ejecutar_autocompletado():
    """
    Ejecuta el autocompletado de campos vacíos en Fletes.
    Escribe los valores encontrados con formato de texto gris.
    """
    try:
        creds = get_credentials()
        gc = gspread.authorize(creds)

        print("Cargando datos de hojas vinculadas...")

        # Cargar datos de referencia
        pesadas = cargar_datos_pesadas(gc)
        descargas_ctg, descargas_cpe = cargar_datos_descargas(gc)
        cpe_data, cpe_por_ctg = cargar_datos_cpe(gc)

        print(f"  - Pesadas: {len(pesadas)} registros")
        print(f"  - Descargas: {len(descargas_ctg)} por CTG, {len(descargas_cpe)} por CPE")
        print(f"  - CPE: {len(cpe_data)} por número, {len(cpe_por_ctg)} por CTG")

        # Cargar Fletes
        ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
        hoja_fletes = ss.worksheet(CONFIG['FLETES_SHEET'])
        datos_fletes = hoja_fletes.get_all_values()

        print(f"Procesando {len(datos_fletes) - 1} filas de Fletes...")

        # Campos a autocompletar: (nombre, columna_fletes)
        campos = [
            ('producto', CONFIG['FLETES_COL_PRODUCTO']),
            ('origen', CONFIG['FLETES_COL_ORIGEN']),
            ('destino', CONFIG['FLETES_COL_DESTINO']),
            ('transportista', CONFIG['FLETES_COL_TRANSPORTISTA']),
            ('chofer', CONFIG['FLETES_COL_CHOFER']),
            ('m_pesadas', CONFIG['FLETES_COL_M_PESADAS']),
            ('m_descargas', CONFIG['FLETES_COL_M_DESCARGAS']),
        ]

        actualizaciones = []
        formatos = []
        campos_completados = 0
        filas_procesadas = 0

        transportistas_corregidos = 0

        for idx, fila in enumerate(datos_fletes[1:], start=2):  # start=2 por encabezado
            ctg = fila[CONFIG['FLETES_COL_CTG']] if len(fila) > CONFIG['FLETES_COL_CTG'] else ''
            cpe = fila[CONFIG['FLETES_COL_CPE']] if len(fila) > CONFIG['FLETES_COL_CPE'] else ''

            # Saltar filas sin CTG ni CPE (no hay forma de vincular)
            if not ctg and not cpe:
                continue

            fila_modificada = False

            for campo_nombre, col_idx in campos:
                valor_actual = fila[col_idx] if len(fila) > col_idx else ''
                valor_actual_str = str(valor_actual).strip() if valor_actual else ''

                # Buscar el valor correcto en las hojas vinculadas
                valor_nuevo = buscar_dato(
                    campo_nombre, ctg, cpe,
                    pesadas, descargas_ctg, descargas_cpe,
                    cpe_data, cpe_por_ctg
                )

                if not valor_nuevo:
                    continue

                # Para TRANSPORTISTA y ORIGEN: también corregir si no coincide exactamente
                if campo_nombre in ('transportista', 'origen'):
                    if not valor_actual_str:
                        # Está vacío, completar
                        celda = gspread.utils.rowcol_to_a1(idx, col_idx + 1)
                        actualizaciones.append({'range': celda, 'values': [[valor_nuevo]]})
                        formatos.append(celda)
                        campos_completados += 1
                        fila_modificada = True
                    elif valor_actual_str != valor_nuevo and valor_actual_str.upper() == valor_nuevo.upper():
                        # Mismo texto pero diferente capitalización, corregir
                        celda = gspread.utils.rowcol_to_a1(idx, col_idx + 1)
                        actualizaciones.append({'range': celda, 'values': [[valor_nuevo]]})
                        formatos.append(celda)
                        if campo_nombre == 'transportista':
                            transportistas_corregidos += 1
                        else:
                            campos_completados += 1
                        fila_modificada = True
                else:
                    # Para otros campos: solo completar si está vacío
                    if not valor_actual_str:
                        celda = gspread.utils.rowcol_to_a1(idx, col_idx + 1)
                        actualizaciones.append({'range': celda, 'values': [[valor_nuevo]]})
                        formatos.append(celda)
                        campos_completados += 1
                        fila_modificada = True

            if fila_modificada:
                filas_procesadas += 1

        print(f"Encontrados {campos_completados} campos para completar, {transportistas_corregidos} transportistas corregidos, en {filas_procesadas} filas")

        # Aplicar actualizaciones en batch
        if actualizaciones:
            print("Escribiendo datos...")
            hoja_fletes.batch_update(actualizaciones)

            # Aplicar formato gris a las celdas actualizadas
            print("Aplicando formato gris...")
            for celda in formatos:
                try:
                    hoja_fletes.format(celda, {
                        'textFormat': {
                            'foregroundColor': COLOR_GRIS
                        }
                    })
                except Exception as e:
                    print(f"  Error formateando {celda}: {e}")

        print("Autocompletado finalizado.")

        return {
            'success': True,
            'campos_completados': campos_completados,
            'transportistas_corregidos': transportistas_corregidos,
            'filas_procesadas': filas_procesadas,
            'detalles': {
                'pesadas_referencia': len(pesadas),
                'descargas_referencia': len(descargas_ctg),
                'cpe_referencia': len(cpe_data)
            }
        }

    except Exception as e:
        print(f"Error en autocompletado: {e}")
        return {
            'success': False,
            'error': str(e),
            'campos_completados': 0,
            'filas_procesadas': 0
        }


if __name__ == '__main__':
    print("=" * 50)
    print("Agente Autocompletador de Fletes")
    print("=" * 50)
    resultado = ejecutar_autocompletado()
    print("\nResultado:")
    print(f"  Éxito: {resultado['success']}")
    print(f"  Campos completados: {resultado['campos_completados']}")
    print(f"  Filas procesadas: {resultado['filas_procesadas']}")
