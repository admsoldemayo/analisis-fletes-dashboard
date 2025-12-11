"""
Data Loader - Carga datos de Google Sheets para el Dashboard de Fletes
"""

import gspread
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import pandas as pd
import os
from datetime import datetime
import re

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# IDs de los Spreadsheets
SPREADSHEET_IDS = {
    'cpe': '1aSZalfUpSFHytq9sYEkzDvXqFC_nBF_9a99kg6qZSXc',
    'pesadas': '1gTvXfwOsqbbc5lxpcsh8HMoB5F3Bix0qpdNKdyY5DME'
}

# Nombres de las hojas
SHEET_NAMES = {
    'fletes': 'Fletes facturados todos',
    'pesadas': 'Pesadas Todos',
    'descargas': 'Descargas Todos',
    'cpe': 'Cartas de Porte Afip'
}


class DataLoader:
    def __init__(self):
        self.gc = None
        self._cache = {}
        self._cache_time = {}
        self.cache_duration = 300  # 5 minutos

    def _get_credentials(self):
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

    def _get_client(self):
        """Obtiene cliente de gspread"""
        if self.gc is None:
            creds = self._get_credentials()
            self.gc = gspread.authorize(creds)
        return self.gc

    def _is_cache_valid(self, key):
        """Verifica si el cache es válido"""
        if key not in self._cache_time:
            return False
        elapsed = (datetime.now() - self._cache_time[key]).seconds
        return elapsed < self.cache_duration

    def _parse_number(self, value):
        """Convierte string a número, manejando formatos argentinos"""
        if pd.isna(value) or value == '' or value is None:
            return None

        value_str = str(value).strip()

        # Remover símbolos de moneda y espacios
        value_str = value_str.replace('$', '').replace(' ', '').strip()

        # Manejar formato argentino: 27.140,00 o 27,140.00
        if ',' in value_str and '.' in value_str:
            # Determinar cuál es el separador decimal
            if value_str.rfind(',') > value_str.rfind('.'):
                # Formato: 27.140,00 (punto miles, coma decimal)
                value_str = value_str.replace('.', '').replace(',', '.')
            else:
                # Formato: 27,140.00 (coma miles, punto decimal)
                value_str = value_str.replace(',', '')
        elif ',' in value_str:
            # Solo coma: puede ser decimal o miles
            parts = value_str.split(',')
            if len(parts) == 2 and len(parts[1]) <= 2:
                # Es decimal: 27140,00
                value_str = value_str.replace(',', '.')
            else:
                # Es miles: 27,140
                value_str = value_str.replace(',', '')

        try:
            return float(value_str)
        except (ValueError, TypeError):
            return None

    def _parse_date(self, value):
        """Convierte string a fecha"""
        if pd.isna(value) or value == '' or value is None:
            return None

        value_str = str(value).strip()

        formatos = [
            '%d/%m/%Y',
            '%Y-%m-%d',
            '%d-%m-%Y',
            '%Y/%m/%d',
            '%d/%m/%y',
        ]

        for fmt in formatos:
            try:
                return datetime.strptime(value_str, fmt)
            except ValueError:
                continue

        return None

    def get_fletes(self, use_cache=True):
        """Obtiene DataFrame de Fletes facturados todos"""
        cache_key = 'fletes'

        if use_cache and self._is_cache_valid(cache_key):
            return self._cache[cache_key].copy()

        gc = self._get_client()
        ss = gc.open_by_key(SPREADSHEET_IDS['cpe'])
        hoja = ss.worksheet(SHEET_NAMES['fletes'])
        datos = hoja.get_all_values()

        if not datos:
            return pd.DataFrame()

        # Crear DataFrame
        df = pd.DataFrame(datos[1:], columns=datos[0])

        # Renombrar columnas para facilitar uso
        column_map = {
            'Numero de factura': 'numero_factura',
            'Fecha': 'fecha',
            'Producto': 'producto',
            'Cantidad': 'cantidad',
            'CTG': 'ctg',
            'CPE': 'cpe',
            'Origen': 'origen',
            'Destino': 'destino',
            'Transportista': 'transportista',
            'Chofer': 'chofer',
            'Subtotal': 'subtotal',
            'IVA': 'iva',
            'Total': 'total',
            'Tarifa': 'tarifa',
            'KM': 'km',
            'Clasificacion': 'clasificacion',
            'M Pesadas todos': 'm_pesadas',
            "M CPE's": 'm_cpes',
            'M Descargas todos': 'm_descargas',
            'M Tarifa Segun la OC (manual)': 'm_tarifa_oc',
            'M Precio acordado (manual)': 'm_precio_acordado'
        }

        df = df.rename(columns=column_map)

        # Convertir tipos
        df['fecha_dt'] = df['fecha'].apply(self._parse_date)
        df['cantidad_num'] = df['cantidad'].apply(self._parse_number)
        df['m_pesadas_num'] = df['m_pesadas'].apply(self._parse_number)
        df['m_descargas_num'] = df['m_descargas'].apply(self._parse_number)
        df['subtotal_num'] = df['subtotal'].apply(self._parse_number)
        df['total_num'] = df['total'].apply(self._parse_number)
        df['tarifa_num'] = df['tarifa'].apply(self._parse_number)

        # Calcular merma
        df['merma_kg'] = df.apply(
            lambda row: (row['m_pesadas_num'] - row['m_descargas_num'])
            if pd.notna(row['m_pesadas_num']) and pd.notna(row['m_descargas_num'])
            else None,
            axis=1
        )

        df['merma_pct'] = df.apply(
            lambda row: (row['merma_kg'] / row['m_pesadas_num'] * 100)
            if pd.notna(row['merma_kg']) and row['m_pesadas_num'] and row['m_pesadas_num'] > 0
            else None,
            axis=1
        )

        # Calcular diferencia facturación
        df['dif_facturacion'] = df.apply(
            lambda row: (row['cantidad_num'] - row['m_descargas_num'])
            if pd.notna(row['cantidad_num']) and pd.notna(row['m_descargas_num'])
            else None,
            axis=1
        )

        # Flags útiles
        df['tiene_cpe'] = df['cpe'].apply(lambda x: bool(x and str(x).strip()))
        df['tiene_pesadas'] = df['m_pesadas_num'].notna()
        df['tiene_descargas'] = df['m_descargas_num'].notna()
        df['merma_sospechosa'] = df['merma_pct'].apply(lambda x: x > 0.3 if pd.notna(x) else False)

        # Guardar en cache
        self._cache[cache_key] = df
        self._cache_time[cache_key] = datetime.now()

        return df.copy()

    def get_pesadas(self, use_cache=True):
        """Obtiene DataFrame de Pesadas Todos"""
        cache_key = 'pesadas'

        if use_cache and self._is_cache_valid(cache_key):
            return self._cache[cache_key].copy()

        gc = self._get_client()
        ss = gc.open_by_key(SPREADSHEET_IDS['pesadas'])
        hoja = ss.worksheet(SHEET_NAMES['pesadas'])
        datos = hoja.get_all_values()

        if not datos:
            return pd.DataFrame()

        df = pd.DataFrame(datos[1:], columns=datos[0])

        # Renombrar columnas
        column_map = {
            'Nº': 'numero',
            'Fecha': 'fecha',
            'Mes': 'mes',
            'Año': 'anio',
            'Producto': 'producto',
            'Cantidad': 'cantidad',
            '     Bruto': 'bruto',
            '     Tara': 'tara',
            'Neto': 'neto',
            '           Destino': 'destino',
            'Origen': 'origen',
            'Placa Camion': 'patente',
            'Tranportista': 'transportista',
            'Carta de porte/Remito': 'cpe',
            'Chofer': 'chofer',
            'Rubro (GAN O AGR)': 'rubro',
            'Extras': 'extras',
            'Campo': 'campo',
            'Clasificacion': 'clasificacion'
        }

        df = df.rename(columns=column_map)

        # Convertir tipos
        df['fecha_dt'] = df['fecha'].apply(self._parse_date)
        df['neto_num'] = df['neto'].apply(self._parse_number)
        df['bruto_num'] = df['bruto'].apply(self._parse_number)
        df['tara_num'] = df['tara'].apply(self._parse_number)

        self._cache[cache_key] = df
        self._cache_time[cache_key] = datetime.now()

        return df.copy()

    def get_descargas(self, use_cache=True):
        """Obtiene DataFrame de Descargas Todos"""
        cache_key = 'descargas'

        if use_cache and self._is_cache_valid(cache_key):
            return self._cache[cache_key].copy()

        gc = self._get_client()
        ss = gc.open_by_key(SPREADSHEET_IDS['cpe'])
        hoja = ss.worksheet(SHEET_NAMES['descargas'])
        datos = hoja.get_all_values()

        if not datos:
            return pd.DataFrame()

        df = pd.DataFrame(datos[1:], columns=datos[0])

        # Renombrar columnas clave
        column_map = {
            'Comprador': 'comprador',
            'Fecha Descarga': 'fecha',
            'Destino': 'destino',
            'Cliente': 'cliente',
            'Producto': 'producto',
            'Orígen': 'origen',
            'Carta de porte/Remito': 'cpe',
            'CTG': 'ctg',
            'Peso Neto': 'peso_neto',
            'Placa Camion': 'patente',
            'Nombre Transporte': 'transportista'
        }

        df = df.rename(columns=column_map)

        # Convertir tipos
        df['fecha_dt'] = df['fecha'].apply(self._parse_date)
        df['peso_neto_num'] = df['peso_neto'].apply(self._parse_number)

        self._cache[cache_key] = df
        self._cache_time[cache_key] = datetime.now()

        return df.copy()

    def get_cpe(self, use_cache=True):
        """Obtiene DataFrame de Cartas de Porte"""
        cache_key = 'cpe'

        if use_cache and self._is_cache_valid(cache_key):
            return self._cache[cache_key].copy()

        gc = self._get_client()
        ss = gc.open_by_key(SPREADSHEET_IDS['cpe'])
        hoja = ss.worksheet(SHEET_NAMES['cpe'])
        datos = hoja.get_all_values()

        if not datos:
            return pd.DataFrame()

        df = pd.DataFrame(datos[1:], columns=datos[0])

        # Convertir fecha
        df['fecha_dt'] = df['fecha_documento'].apply(self._parse_date)

        self._cache[cache_key] = df
        self._cache_time[cache_key] = datetime.now()

        return df.copy()

    def get_summary_stats(self):
        """Obtiene estadísticas resumidas para KPIs"""
        df = self.get_fletes()

        stats = {
            'total_fletes': len(df),
            'con_cpe': df['tiene_cpe'].sum(),
            'sin_cpe': (~df['tiene_cpe']).sum(),
            'pesados_campo': df['tiene_pesadas'].sum(),
            'descargados': df['tiene_descargas'].sum(),
            'merma_promedio': df['merma_pct'].mean() if df['merma_pct'].notna().any() else 0,
            'mermas_sospechosas': df['merma_sospechosa'].sum(),
            'total_kg_pesados': df['m_pesadas_num'].sum(),
            'total_kg_descargados': df['m_descargas_num'].sum(),
        }

        return stats

    def get_merma_by_transportista(self):
        """Obtiene análisis de merma por transportista"""
        df = self.get_fletes()

        # Filtrar solo los que tienen ambos pesos
        df_completo = df[df['tiene_pesadas'] & df['tiene_descargas']].copy()

        if df_completo.empty:
            return pd.DataFrame()

        # Agrupar por transportista
        resumen = df_completo.groupby('transportista').agg({
            'merma_pct': 'mean',
            'merma_kg': 'sum',
            'm_pesadas_num': 'sum',
            'm_descargas_num': 'sum',
            'numero_factura': 'count'
        }).reset_index()

        resumen.columns = ['transportista', 'merma_pct_promedio', 'merma_kg_total',
                          'kg_pesados_total', 'kg_descargados_total', 'cantidad_fletes']

        # Clasificar riesgo
        def clasificar_riesgo(merma):
            if pd.isna(merma):
                return 'Sin datos'
            if merma <= 0.3:
                return 'Normal'
            elif merma <= 1.0:
                return 'Atención'
            else:
                return 'Alerta'

        resumen['clasificacion'] = resumen['merma_pct_promedio'].apply(clasificar_riesgo)

        return resumen.sort_values('merma_pct_promedio', ascending=False)

    def get_fletes_by_producto(self):
        """Obtiene distribución de fletes por producto"""
        df = self.get_fletes()

        resumen = df.groupby('producto').agg({
            'numero_factura': 'count',
            'm_pesadas_num': 'sum',
            'm_descargas_num': 'sum',
            'total_num': 'sum'
        }).reset_index()

        resumen.columns = ['producto', 'cantidad_fletes', 'kg_pesados', 'kg_descargados', 'total_facturado']

        return resumen.sort_values('cantidad_fletes', ascending=False)

    def clear_cache(self):
        """Limpia el cache"""
        self._cache = {}
        self._cache_time = {}


# Instancia global
data_loader = DataLoader()
