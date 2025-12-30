"""
Dashboard de Control de Fletes Agropecuario
Dash/Plotly - Puerto 5016
"""

import dash
from dash import dcc, html, dash_table, callback_context
from dash.dependencies import Input, Output, State, ALL
import dash_bootstrap_components as dbc
import plotly.express as px
import plotly.graph_objects as go
import pandas as pd
from datetime import datetime, timedelta
from data_loader import data_loader
import requests

# Colores del tema agro
COLORS = {
    'primary': '#2E7D32',
    'secondary': '#FFA000',
    'danger': '#D32F2F',
    'warning': '#F57C00',
    'success': '#388E3C',
    'background': '#F5F5F5',
    'card': '#FFFFFF',
    'text': '#212121',
    'text_light': '#757575'
}

# Inicializar app
app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True,
    title='Control de Fletes Agropecuario'
)

server = app.server


def create_kpi_card(title, value, icon, color, subtitle=None):
    """Crea una tarjeta KPI"""
    return dbc.Card([
        dbc.CardBody([
            html.Div([
                html.I(className=f"fas {icon}", style={'fontSize': '2rem', 'color': color}),
            ], className="text-center mb-2"),
            html.H3(value, className="text-center mb-1", style={'color': color, 'fontWeight': 'bold'}),
            html.P(title, className="text-center text-muted mb-0", style={'fontSize': '0.85rem'}),
            html.Small(subtitle, className="text-center text-muted d-block") if subtitle else None
        ])
    ], className="shadow-sm h-100")


def get_navbar():
    """Barra de navegación"""
    return dbc.Navbar(
        dbc.Container([
            dbc.Row([
                dbc.Col([
                    html.I(className="fas fa-truck me-2", style={'fontSize': '1.5rem'}),
                    dbc.NavbarBrand("Control de Fletes Agropecuario", className="ms-2 fw-bold")
                ], width="auto", className="d-flex align-items-center")
            ]),
            dbc.Nav([
                dbc.NavItem(dbc.NavLink("Dashboard", href="/", active="exact", className="px-3")),
                dbc.NavItem(dbc.NavLink("Procesadores", href="/procesadores", active="exact", className="px-3")),
                dbc.NavItem(dbc.NavLink("Trabajo Manual", href="/trabajo-manual", active="exact", className="px-3")),
            ], className="ms-auto", navbar=True)
        ], fluid=True),
        color="success",
        dark=True,
        className="mb-4"
    )


def get_dashboard_content():
    """Contenido de la página Dashboard"""
    try:
        df = data_loader.get_fletes()
        transportistas = sorted([t for t in df['transportista'].dropna().unique().tolist() if t])
        productos = sorted([p for p in df['producto'].dropna().unique().tolist() if p])
        origenes = sorted([o for o in df['origen'].dropna().unique().tolist() if o])

        fechas_validas = df['fecha_dt'].dropna()
        if not fechas_validas.empty:
            min_date = fechas_validas.min()
            max_date = fechas_validas.max()
        else:
            min_date = datetime.now() - timedelta(days=365)
            max_date = datetime.now()
    except Exception as e:
        print(f"Error cargando datos: {e}")
        transportistas, productos, origenes = [], [], []
        min_date = datetime.now() - timedelta(days=365)
        max_date = datetime.now()

    return html.Div([
        # Botones de acción
        dbc.Row([
            dbc.Col([
                dbc.Button([
                    html.I(className="fas fa-sync-alt me-2"),
                    "Actualizar Datos"
                ], id="btn-refresh", color="success", className="me-2"),
                dbc.Button([
                    html.I(className="fas fa-magic me-2"),
                    "Agente Corrección Datos Faltantes"
                ], id="btn-autocomplete", color="warning"),
            ], className="text-end mb-3")
        ]),

        # Filtros
        dbc.Card([
            dbc.CardBody([
                dbc.Row([
                    dbc.Col([
                        html.Label("Rango de Fechas", className="fw-bold mb-1"),
                        dcc.DatePickerRange(
                            id='filtro-fechas',
                            min_date_allowed=min_date,
                            max_date_allowed=max_date,
                            start_date=min_date,
                            end_date=max_date,
                            display_format='DD/MM/YYYY',
                            className="w-100"
                        )
                    ], md=3),
                    dbc.Col([
                        html.Label("Transportista", className="fw-bold mb-1"),
                        dcc.Dropdown(
                            id='filtro-transportista',
                            options=[{'label': t, 'value': t} for t in transportistas],
                            multi=True,
                            placeholder="Todos...",
                            style={'position': 'relative', 'zIndex': 10}
                        )
                    ], md=3, style={'zIndex': 10}),
                    dbc.Col([
                        html.Label("Producto", className="fw-bold mb-1"),
                        dcc.Dropdown(
                            id='filtro-producto',
                            options=[{'label': p, 'value': p} for p in productos],
                            multi=True,
                            placeholder="Todos...",
                            style={'position': 'relative', 'zIndex': 10}
                        )
                    ], md=2, style={'zIndex': 10}),
                    dbc.Col([
                        html.Label("Origen", className="fw-bold mb-1"),
                        dcc.Dropdown(
                            id='filtro-origen',
                            options=[{'label': o, 'value': o} for o in origenes],
                            multi=True,
                            placeholder="Todos...",
                            style={'position': 'relative', 'zIndex': 10}
                        )
                    ], md=2, style={'zIndex': 10}),
                    dbc.Col([
                        html.Label("Mostrar", className="fw-bold mb-1"),
                        dbc.Checklist(
                            id='filtro-problemas',
                            options=[{'label': ' Solo con problemas', 'value': 'problemas'}],
                            value=[],
                            switch=True
                        )
                    ], md=2),
                ])
            ])
        ], className="shadow-sm mb-4", style={'zIndex': 100, 'position': 'relative'}),

        # Mensaje explicativo
        dbc.Alert([
            html.I(className="fas fa-info-circle me-2"),
            html.Strong("Cómo interpretar este dashboard: "),
            html.Span("Fletes OK", style={'color': 'green', 'fontWeight': 'bold'}),
            " = tienen CPE + pesados + descargados (datos completos). ",
            html.Span("Filtro 'Solo con problemas'", style={'color': 'red', 'fontWeight': 'bold'}),
            " = sin CPE, merma >0.3%, o diferencia facturación >100kg. ",
            "El gráfico de merma muestra el % promedio y los kg totales por transportista."
        ], color="info", className="mb-3", dismissable=True),

        # KPIs - Fila 1 (Estado general)
        dbc.Row([
            dbc.Col(html.Div(id='kpi-total'), md=2),
            dbc.Col(html.Div(id='kpi-ok'), md=2),
            dbc.Col(html.Div(id='kpi-con-cpe'), md=2),
            dbc.Col(html.Div(id='kpi-no-llevan-cpe'), md=2),
            dbc.Col(html.Div(id='kpi-pesados'), md=2),
            dbc.Col(html.Div(id='kpi-descargados'), md=2),
        ], className="mb-3"),

        # KPIs - Fila 2 (Pendientes/Alertas - Rojos, alineados con fila 1)
        dbc.Row([
            dbc.Col(html.Div(id='kpi-merma'), md=2),
            dbc.Col(md=2),  # Espacio vacío debajo de Fletes OK
            dbc.Col(html.Div(id='kpi-falta-cpe'), md=2),
            dbc.Col(md=2),  # Espacio vacío debajo de No llevan CPE
            dbc.Col(html.Div(id='kpi-falta-pesadas'), md=2),
            dbc.Col(html.Div(id='kpi-falta-descargas'), md=2),
        ], className="mb-4"),

        # Gráficos principales
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.I(className="fas fa-chart-bar me-2"),
                        "Merma por Transportista"
                    ], className="fw-bold"),
                    dbc.CardBody([
                        dcc.Graph(id='grafico-merma-transportista', style={'height': '350px'})
                    ])
                ], className="shadow-sm h-100")
            ], md=6),
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.I(className="fas fa-chart-pie me-2"),
                        "Distribución por Cultivo"
                    ], className="fw-bold"),
                    dbc.CardBody([
                        dcc.Graph(id='grafico-cultivos', style={'height': '350px'})
                    ])
                ], className="shadow-sm h-100")
            ], md=6),
        ], className="mb-4"),

        # Tabla de alertas
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.I(className="fas fa-exclamation-triangle me-2 text-warning"),
                        "Fletes con Alertas",
                        dbc.Badge(id='badge-alertas', color="danger", className="ms-2")
                    ], className="fw-bold"),
                    dbc.CardBody([
                        dash_table.DataTable(
                            id='tabla-alertas',
                            columns=[
                                {'name': 'Fecha', 'id': 'fecha'},
                                {'name': 'Transportista', 'id': 'transportista'},
                                {'name': 'Producto', 'id': 'producto'},
                                {'name': 'Origen', 'id': 'origen'},
                                {'name': 'Pesado (kg)', 'id': 'm_pesadas'},
                                {'name': 'Descargado (kg)', 'id': 'm_descargas'},
                                {'name': 'Dif. (kg)', 'id': 'diferencia_kg'},
                                {'name': 'Merma %', 'id': 'merma_pct'},
                                {'name': 'Facturado', 'id': 'cantidad'},
                                {'name': 'Dif. Fact vs Desc', 'id': 'dif_facturacion'},
                            ],
                            style_table={'overflowX': 'auto'},
                            style_cell={'textAlign': 'left', 'padding': '10px', 'fontSize': '13px'},
                            style_header={
                                'backgroundColor': COLORS['primary'],
                                'color': 'white',
                                'fontWeight': 'bold'
                            },
                            style_data_conditional=[
                                {'if': {'filter_query': '{merma_pct} > 0.3'}, 'backgroundColor': '#ffebee'},
                                {'if': {'filter_query': '{merma_pct} > 1'}, 'backgroundColor': '#ffcdd2',
                                 'color': COLORS['danger'], 'fontWeight': 'bold'},
                            ],
                            page_size=15,
                            sort_action='native',
                            filter_action='native',
                            export_format='xlsx'
                        )
                    ])
                ], className="shadow-sm")
            ], md=12),
        ]),

        html.Div(id='ultima-actualizacion', className="text-muted text-center mt-3")
    ])


def get_procesadores_content():
    """Contenido de la página Procesadores"""
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.H4([
                    html.I(className="fas fa-cogs me-2"),
                    "Procesadores de Datos"
                ], className="mb-4"),
                html.P("Ejecuta los procesos de vinculación y asignación de datos entre hojas.",
                      className="text-muted mb-4")
            ])
        ]),

        # Botón EJECUTAR TODO
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardBody([
                        dbc.Button([
                            html.I(className="fas fa-rocket me-2"),
                            "EJECUTAR TODO (Pasos 0-3 + Agente)"
                        ], id="btn-ejecutar-todo", color="dark", size="lg", className="w-100"),
                        dcc.Loading(
                            id="loading-ejecutar-todo",
                            type="default",
                            children=html.Div(id="resultado-ejecutar-todo", className="mt-3")
                        )
                    ])
                ], className="shadow border-dark mb-4", style={'borderWidth': '2px'})
            ], md=8, className="mx-auto"),
        ]),

        # Nuevo: Traer CPEs a Fletes
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.I(className="fas fa-file-import me-2 text-danger"),
                        "Paso 0 - Automático"
                    ], className="bg-danger bg-opacity-25"),
                    dbc.CardBody([
                        html.H5("Traer CPEs a Fletes", className="card-title"),
                        html.P([
                            "Busca el ", html.Strong("numero_cpe"), " en Cartas de Porte por CTG y lo escribe en la columna ",
                            html.Strong("CPE"), " de Fletes. Marca ",
                            html.Span("si", style={'color': 'green', 'fontWeight': 'bold'}), " (verde) o ",
                            html.Span("no", style={'color': 'red', 'fontWeight': 'bold'}), " (rojo) en ",
                            html.Strong("M CPE's"), "."
                        ], className="card-text"),
                        dbc.Button([
                            html.I(className="fas fa-download me-2"),
                            "Traer CPEs"
                        ], id="btn-traer-cpes", color="danger", size="lg", className="w-100 mt-2"),
                        html.Div(id="resultado-traer-cpes", className="mt-3")
                    ])
                ], className="shadow h-100")
            ], md=6, className="mx-auto"),
        ], className="mb-4"),

        dbc.Row([
            # Asignar CPEs a Pesadas
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.I(className="fas fa-link me-2 text-primary"),
                        "Paso 1"
                    ], className="bg-light"),
                    dbc.CardBody([
                        html.H5("Asignar CPEs a Pesadas", className="card-title"),
                        html.P("Match por patente. Si hay duplicados, desambigua por peso (Neto vs Cantidad).",
                              className="card-text text-muted"),
                        dbc.Button([
                            html.I(className="fas fa-play me-2"),
                            "Ejecutar"
                        ], id="btn-asignar-cpe", color="primary", className="w-100 mt-2"),
                        html.Div(id="resultado-asignar-cpe", className="mt-3")
                    ])
                ], className="shadow-sm h-100")
            ], md=4),

            # Pesadas -> Fletes
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.I(className="fas fa-weight me-2 text-success"),
                        "Paso 2"
                    ], className="bg-light"),
                    dbc.CardBody([
                        html.H5("Pesadas → Fletes", className="card-title"),
                        html.P("Lleva el peso Neto de Pesadas a la columna M Pesadas de Fletes (via CTG).",
                              className="card-text text-muted"),
                        dbc.Button([
                            html.I(className="fas fa-play me-2"),
                            "Ejecutar"
                        ], id="btn-matchear-fletes", color="success", className="w-100 mt-2"),
                        html.Div(id="resultado-matchear-fletes", className="mt-3")
                    ])
                ], className="shadow-sm h-100")
            ], md=4),

            # Descargas -> Fletes
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.I(className="fas fa-warehouse me-2 text-info"),
                        "Paso 3"
                    ], className="bg-light"),
                    dbc.CardBody([
                        html.H5("Descargas → Fletes", className="card-title"),
                        html.P("Lleva el Peso Neto de Descargas a la columna M Descargas de Fletes (via CTG).",
                              className="card-text text-muted"),
                        dbc.Button([
                            html.I(className="fas fa-play me-2"),
                            "Ejecutar"
                        ], id="btn-matchear-descargas", color="info", className="w-100 mt-2"),
                        html.Div(id="resultado-matchear-descargas", className="mt-3")
                    ])
                ], className="shadow-sm h-100")
            ], md=4),
        ], className="mb-4"),

        html.Hr(className="my-4"),

        # Agente Corrección
        dbc.Row([
            dbc.Col([
                dbc.Card([
                    dbc.CardHeader([
                        html.I(className="fas fa-magic me-2 text-warning"),
                        "Agente Inteligente"
                    ], className="bg-warning bg-opacity-25"),
                    dbc.CardBody([
                        html.H5("Agente Corrección Datos Faltantes", className="card-title"),
                        html.P([
                            "Busca y completa automáticamente campos vacíos en Fletes: ",
                            html.Strong("Producto, Origen, Destino, Transportista, Chofer, M Pesadas, M Descargas"),
                            ". Los datos completados se muestran en ",
                            html.Span("gris", style={'color': 'gray', 'fontWeight': 'bold'}),
                            " para identificarlos."
                        ], className="card-text"),
                        dbc.Button([
                            html.I(className="fas fa-robot me-2"),
                            "Ejecutar Agente"
                        ], id="btn-agente-correccion", color="warning", size="lg", className="w-100 mt-2"),
                        html.Div(id="resultado-agente-correccion", className="mt-3")
                    ])
                ], className="shadow h-100")
            ], md=6, className="mx-auto"),
        ]),
    ])


# Layout principal con navegación
app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    get_navbar(),
    dbc.Container([
        html.Div(id='page-content')
    ], fluid=True, style={'backgroundColor': COLORS['background'], 'minHeight': '90vh', 'paddingBottom': '2rem'}),

    # Modales
    dbc.Modal([
        dbc.ModalHeader(dbc.ModalTitle("Resultado")),
        dbc.ModalBody(id='modal-body'),
        dbc.ModalFooter(dbc.Button("Cerrar", id="close-modal", className="ms-auto"))
    ], id="modal-resultado", is_open=False),

    dcc.Store(id='store-data'),
])


def get_trabajo_manual_content():
    """Página de Trabajo Manual con pestañas para Sin CPE y Duplicados"""
    # URL del sheet de Pesadas
    PESADAS_SHEET_URL = "https://docs.google.com/spreadsheets/d/1gTvXfwOsqbbc5lxpcsh8HMoB5F3Bix0qpdNKdyY5DME/edit"

    return dbc.Container([
        html.H3([
            html.I(className="fas fa-tools me-2"),
            "Trabajo Manual"
        ], className="mb-4"),

        dbc.Tabs([
            # Tab 1: Sin CPE
            dbc.Tab(label="Sin CPE", tab_id="tab-sin-cpe", children=[
                html.Div([
                    dbc.Alert([
                        html.I(className="fas fa-info-circle me-2"),
                        html.Strong("¿Para qué sirve? "),
                        "Clasificar fletes que no tienen Carta de Porte (traslados internos, en B, etc.)."
                    ], color="primary", className="mb-3 mt-3"),

                    dbc.Row([
                        dbc.Col([
                            dbc.Card([
                                dbc.CardBody([
                                    dbc.Button([
                                        html.I(className="fas fa-sync me-2"),
                                        "Cargar Fletes Sin CPE"
                                    ], id="btn-cargar-sin-cpe", color="danger", className="w-100"),
                                    dcc.Loading(
                                        id="loading-sin-cpe",
                                        type="default",
                                        children=html.Div(id="resultado-cargar-sin-cpe", className="mt-3")
                                    )
                                ])
                            ], className="shadow-sm")
                        ], md=4),
                        dbc.Col([
                            dbc.Card([
                                dbc.CardBody([
                                    html.Div(id="stats-sin-cpe", children=[
                                        html.P("Cargá los fletes para ver estadísticas.", className="text-muted mb-0")
                                    ])
                                ])
                            ], className="shadow-sm h-100")
                        ], md=8),
                    ], className="mb-4"),

                    dbc.Card([
                        dbc.CardHeader([
                            html.I(className="fas fa-table me-2"),
                            "Fletes sin Carta de Porte",
                            dbc.Badge(id="badge-sin-cpe", color="danger", className="ms-2")
                        ], className="bg-light"),
                        dbc.CardBody([
                            html.Div(id="tabla-sin-cpe", children=[
                                html.P("Cargá los fletes para ver el detalle.", className="text-muted text-center py-4")
                            ])
                        ])
                    ], className="shadow-sm"),

                    dcc.Store(id='store-fletes-sin-cpe'),
                ])
            ]),

            # Tab 2: Duplicados
            dbc.Tab(label="Duplicados", tab_id="tab-duplicados", children=[
                html.Div([
                    dbc.Alert([
                        html.I(className="fas fa-info-circle me-2"),
                        html.Strong("¿Para qué sirve? "),
                        "Revisar casos donde una patente aparece en varias CPEs y verificar que la asignación sea correcta."
                    ], color="primary", className="mb-3 mt-3"),

                    dbc.Alert([
                        html.I(className="fas fa-filter me-2"),
                        "Solo muestra casos con ambigüedad real: sin match exacto de fecha Y sin desempate único por producto."
                    ], color="info", className="mb-4"),

                    dbc.Row([
                        dbc.Col([
                            html.A(
                                dbc.Button([
                                    html.I(className="fas fa-external-link-alt me-2"),
                                    "Abrir Pesadas Todos (Google Sheets)"
                                ], color="success", className="mb-4"),
                                href=PESADAS_SHEET_URL,
                                target="_blank"
                            )
                        ])
                    ]),

                    dbc.Row([
                        dbc.Col([
                            dbc.Card([
                                dbc.CardHeader([
                                    html.I(className="fas fa-search me-2"),
                                    "Buscar Duplicados en CPE"
                                ], className="bg-warning text-dark"),
                                dbc.CardBody([
                                    html.P("Carga las patentes que aparecen en múltiples Cartas de Porte.",
                                          className="card-text text-muted"),
                                    dbc.Button([
                                        html.I(className="fas fa-sync me-2"),
                                        "Cargar Duplicados"
                                    ], id="btn-cargar-duplicados", color="warning", className="w-100"),
                                    html.Div(id="resultado-duplicados", className="mt-3")
                                ])
                            ], className="shadow-sm")
                        ], md=6),

                        dbc.Col([
                            dbc.Card([
                                dbc.CardHeader([
                                    html.I(className="fas fa-chart-bar me-2"),
                                    "Estadísticas"
                                ], className="bg-light"),
                                dbc.CardBody([
                                    html.Div(id="stats-duplicados", children=[
                                        html.P("Cargá los duplicados para ver estadísticas.", className="text-muted")
                                    ])
                                ])
                            ], className="shadow-sm")
                        ], md=6)
                    ], className="mb-4"),

                    dbc.Card([
                        dbc.CardHeader([
                            html.I(className="fas fa-table me-2"),
                            "Detalle de Patentes Duplicadas"
                        ], className="bg-light"),
                        dbc.CardBody([
                            html.Div(id="tabla-duplicados", children=[
                                html.P("Cargá los duplicados para ver el detalle.", className="text-muted text-center py-4")
                            ])
                        ])
                    ], className="shadow-sm")
                ])
            ]),
        ], id="tabs-trabajo-manual", active_tab="tab-sin-cpe"),
    ], fluid=True, className="py-4")


# Callback para navegación
@app.callback(
    Output('page-content', 'children'),
    Input('url', 'pathname')
)
def display_page(pathname):
    if pathname == '/procesadores':
        return get_procesadores_content()
    elif pathname == '/trabajo-manual':
        return get_trabajo_manual_content()
    else:
        return get_dashboard_content()


# Callbacks para los procesadores
@app.callback(
    Output('resultado-traer-cpes', 'children'),
    Input('btn-traer-cpes', 'n_clicks'),
    prevent_initial_call=True
)
def ejecutar_traer_cpes(n_clicks):
    if not n_clicks:
        return ""
    try:
        from app import traer_cpes_a_fletes
        resultado = traer_cpes_a_fletes()
        if resultado['success']:
            return dbc.Alert([
                html.Strong("Completado: "),
                html.Span(f"{resultado['con_cpe']} con CPE", style={'color': 'green', 'fontWeight': 'bold'}),
                ", ",
                html.Span(f"{resultado['sin_cpe']} sin CPE", style={'color': 'red', 'fontWeight': 'bold'}),
                f". Total: {resultado['total_fletes']}"
            ], color="success", className="mb-0")
        else:
            return dbc.Alert(f"Error: {resultado['error']}", color="danger", className="mb-0")
    except Exception as e:
        return dbc.Alert(f"Error: {str(e)}", color="danger", className="mb-0")


@app.callback(
    Output('resultado-asignar-cpe', 'children'),
    Input('btn-asignar-cpe', 'n_clicks'),
    prevent_initial_call=True
)
def ejecutar_asignar_cpe(n_clicks):
    if not n_clicks:
        return ""
    try:
        from app import asignar_cpes
        resultado = asignar_cpes()
        if resultado['success']:
            # Mostrar estadísticas detalladas
            detalles = []
            detalles.append(html.Strong("Completado: "))
            detalles.append(f"{resultado['matches_nuevos']} CPEs asignados. ")
            detalles.append(html.Br())
            detalles.append(f"Ya tenían: {resultado['ya_tenian_cpe']}. Sin match: {resultado['sin_match']}")

            # Mostrar desglose de matches
            detalles.append(html.Br())
            detalles.append(html.Small([
                html.I(className="fas fa-check-circle me-1 text-success"),
                f"Únicos: {resultado.get('matches_unicos', 0)} | ",
                f"Con empate (REVISAR): {resultado.get('matches_con_empate', 0)} | ",
                f"Fuera de rango (>7 días): {resultado.get('fuera_de_rango', 0)}"
            ], className="text-muted"))

            # Si hay casos marcados para revisar
            empates = resultado.get('matches_con_empate', 0)
            if empates > 0:
                detalles.append(html.Br())
                detalles.append(html.Small([
                    html.I(className="fas fa-exclamation-triangle me-1 text-warning"),
                    f"{empates} casos marcados 'REVISAR' en columna T - verificar en Trabajo Manual > Duplicados"
                ], className="text-warning"))

            return dbc.Alert(detalles, color="success", className="mb-0")
        else:
            return dbc.Alert(f"Error: {resultado['error']}", color="danger", className="mb-0")
    except Exception as e:
        return dbc.Alert(f"Error: {str(e)}", color="danger", className="mb-0")


@app.callback(
    Output('resultado-matchear-fletes', 'children'),
    Input('btn-matchear-fletes', 'n_clicks'),
    prevent_initial_call=True
)
def ejecutar_matchear_fletes(n_clicks):
    if not n_clicks:
        return ""
    try:
        from app import matchear_pesadas_fletes
        resultado = matchear_pesadas_fletes()
        if resultado['success']:
            return dbc.Alert([
                html.Strong("Completado: "),
                f"{resultado['matches_nuevos']} pesos llevados a Fletes. ",
                f"Ya tenían: {resultado['ya_tenian_neto']}"
            ], color="success", className="mb-0")
        else:
            return dbc.Alert(f"Error: {resultado['error']}", color="danger", className="mb-0")
    except Exception as e:
        return dbc.Alert(f"Error: {str(e)}", color="danger", className="mb-0")


@app.callback(
    Output('resultado-matchear-descargas', 'children'),
    Input('btn-matchear-descargas', 'n_clicks'),
    prevent_initial_call=True
)
def ejecutar_matchear_descargas(n_clicks):
    if not n_clicks:
        return ""
    try:
        from app import matchear_descargas_fletes
        resultado = matchear_descargas_fletes()
        if resultado['success']:
            return dbc.Alert([
                html.Strong("Completado: "),
                f"{resultado['matches_nuevos']} pesos llevados a Fletes. ",
                f"Ya tenían: {resultado['ya_tenian_neto']}"
            ], color="success", className="mb-0")
        else:
            return dbc.Alert(f"Error: {resultado['error']}", color="danger", className="mb-0")
    except Exception as e:
        return dbc.Alert(f"Error: {str(e)}", color="danger", className="mb-0")


@app.callback(
    Output('resultado-agente-correccion', 'children'),
    Input('btn-agente-correccion', 'n_clicks'),
    prevent_initial_call=True
)
def ejecutar_agente(n_clicks):
    if not n_clicks:
        return ""
    try:
        from agent_autocomplete import ejecutar_autocompletado
        resultado = ejecutar_autocompletado()
        if resultado['success']:
            return dbc.Alert([
                html.H6("Autocompletado exitoso", className="alert-heading"),
                html.P([
                    f"Campos completados: {resultado['campos_completados']}",
                    html.Br(),
                    f"Transportistas corregidos: {resultado.get('transportistas_corregidos', 0)}",
                    html.Br(),
                    f"Filas procesadas: {resultado['filas_procesadas']}"
                ], className="mb-0"),
                html.Small("Los datos se muestran en gris en la hoja.", className="text-muted")
            ], color="success")
        else:
            return dbc.Alert(f"Error: {resultado['error']}", color="danger")
    except Exception as e:
        return dbc.Alert(f"Error: {str(e)}", color="danger")


# Callback para EJECUTAR TODO
@app.callback(
    Output('resultado-ejecutar-todo', 'children'),
    Input('btn-ejecutar-todo', 'n_clicks'),
    prevent_initial_call=True
)
def ejecutar_todo(n_clicks):
    if not n_clicks:
        return ""

    import time
    DELAY_ENTRE_PASOS = 15  # segundos entre pasos para evitar quota exceeded

    resultados = []
    errores = []

    # Paso 0: Traer CPEs a Fletes
    try:
        from app import traer_cpes_a_fletes
        res0 = traer_cpes_a_fletes()
        if res0['success']:
            resultados.append(html.Div([
                html.I(className="fas fa-check text-success me-2"),
                html.Strong("Paso 0: "),
                f"{res0['con_cpe']} con CPE, {res0['sin_cpe']} sin CPE"
            ]))
        else:
            errores.append(f"Paso 0: {res0['error']}")
    except Exception as e:
        errores.append(f"Paso 0: {str(e)}")

    time.sleep(DELAY_ENTRE_PASOS)

    # Paso 1: Asignar CPEs a Pesadas
    try:
        from app import asignar_cpes
        res1 = asignar_cpes()
        if res1['success']:
            resultados.append(html.Div([
                html.I(className="fas fa-check text-success me-2"),
                html.Strong("Paso 1: "),
                f"{res1['matches_nuevos']} CPEs asignados"
            ]))
        else:
            errores.append(f"Paso 1: {res1['error']}")
    except Exception as e:
        errores.append(f"Paso 1: {str(e)}")

    time.sleep(DELAY_ENTRE_PASOS)

    # Paso 2: Pesadas -> Fletes
    try:
        from app import matchear_pesadas_fletes
        res2 = matchear_pesadas_fletes()
        if res2['success']:
            resultados.append(html.Div([
                html.I(className="fas fa-check text-success me-2"),
                html.Strong("Paso 2: "),
                f"{res2['matches_nuevos']} pesos de pesadas llevados"
            ]))
        else:
            errores.append(f"Paso 2: {res2['error']}")
    except Exception as e:
        errores.append(f"Paso 2: {str(e)}")

    time.sleep(DELAY_ENTRE_PASOS)

    # Paso 3: Descargas -> Fletes
    try:
        from app import matchear_descargas_fletes
        res3 = matchear_descargas_fletes()
        if res3['success']:
            resultados.append(html.Div([
                html.I(className="fas fa-check text-success me-2"),
                html.Strong("Paso 3: "),
                f"{res3['matches_nuevos']} pesos de descargas llevados"
            ]))
        else:
            errores.append(f"Paso 3: {res3['error']}")
    except Exception as e:
        errores.append(f"Paso 3: {str(e)}")

    time.sleep(DELAY_ENTRE_PASOS)

    # Agente Inteligente
    try:
        from agent_autocomplete import ejecutar_autocompletado
        res_agente = ejecutar_autocompletado()
        if res_agente['success']:
            resultados.append(html.Div([
                html.I(className="fas fa-check text-success me-2"),
                html.Strong("Agente: "),
                f"{res_agente['campos_completados']} campos completados"
            ]))
        else:
            errores.append(f"Agente: {res_agente['error']}")
    except Exception as e:
        errores.append(f"Agente: {str(e)}")

    # Construir resultado final
    contenido = []

    if resultados:
        contenido.append(dbc.Alert([
            html.H6([
                html.I(className="fas fa-check-circle me-2"),
                "Procesos completados"
            ], className="alert-heading"),
            html.Hr(),
            *resultados
        ], color="success"))

    if errores:
        contenido.append(dbc.Alert([
            html.H6([
                html.I(className="fas fa-exclamation-triangle me-2"),
                "Errores"
            ], className="alert-heading"),
            html.Hr(),
            *[html.Div(e) for e in errores]
        ], color="danger"))

    return html.Div(contenido)


# Callback principal del dashboard
@app.callback(
    [Output('kpi-total', 'children'),
     Output('kpi-ok', 'children'),
     Output('kpi-con-cpe', 'children'),
     Output('kpi-no-llevan-cpe', 'children'),
     Output('kpi-pesados', 'children'),
     Output('kpi-descargados', 'children'),
     Output('kpi-falta-cpe', 'children'),
     Output('kpi-falta-pesadas', 'children'),
     Output('kpi-falta-descargas', 'children'),
     Output('kpi-merma', 'children'),
     Output('grafico-merma-transportista', 'figure'),
     Output('grafico-cultivos', 'figure'),
     Output('tabla-alertas', 'data'),
     Output('badge-alertas', 'children'),
     Output('ultima-actualizacion', 'children')],
    [Input('filtro-fechas', 'start_date'),
     Input('filtro-fechas', 'end_date'),
     Input('filtro-transportista', 'value'),
     Input('filtro-producto', 'value'),
     Input('filtro-origen', 'value'),
     Input('filtro-problemas', 'value'),
     Input('btn-refresh', 'n_clicks'),
     Input('url', 'pathname')]
)
def update_dashboard(start_date, end_date, transportistas, productos, origenes, problemas, n_clicks, pathname):
    """Actualiza todo el dashboard según filtros"""

    # Solo actualizar si estamos en la página principal
    if pathname == '/procesadores':
        return [None] * 15

    ctx = callback_context
    if ctx.triggered and 'btn-refresh' in ctx.triggered[0]['prop_id']:
        data_loader.clear_cache()

    try:
        df = data_loader.get_fletes()
    except:
        empty_fig = go.Figure()
        empty_fig.add_annotation(text="Error cargando datos", showarrow=False)
        return [create_kpi_card("Error", "-", "fa-times", COLORS['danger'])] * 10 + [empty_fig, empty_fig, [], "0", "Error"]

    # Aplicar filtros
    if start_date:
        df = df[df['fecha_dt'] >= pd.to_datetime(start_date)]
    if end_date:
        df = df[df['fecha_dt'] <= pd.to_datetime(end_date)]
    if transportistas:
        df = df[df['transportista'].isin(transportistas)]
    if productos:
        df = df[df['producto'].isin(productos)]
    if origenes:
        df = df[df['origen'].isin(origenes)]
    if problemas and 'problemas' in problemas:
        df = df[(df['merma_sospechosa']) | (~df['tiene_cpe']) | (df['dif_facturacion'].abs() > 100)]

    # KPIs
    total = len(df)
    # Solo contar los que tienen un número real en M Pesadas / M Descargas (no vacío, no cero)
    pesados = len(df[(df['m_pesadas_num'].notna()) & (df['m_pesadas_num'] > 0)]) if total > 0 else 0
    descargados = len(df[(df['m_descargas_num'].notna()) & (df['m_descargas_num'] > 0)]) if total > 0 else 0
    merma_prom = df['merma_pct'].mean() if df['merma_pct'].notna().any() else 0

    # Calcular fletes que NO llevan CPE (Traslado interno, Flete en B, etc.)
    no_llevan_cpe_valores = ['Traslado interno', 'Flete en B', 'Sin documentación', 'CPE Hecha por Terceros']
    no_llevan_cpe = len(df[df['m_cpes'].isin(no_llevan_cpe_valores)]) if 'm_cpes' in df.columns else 0

    # Con CPE real = tienen CPE y NO están en la lista de "no llevan CPE"
    # Excluir los que tienen "No corresponde" en columna CPE
    con_cpe_real = len(df[
        (df['tiene_cpe']) &
        (~df['cpe'].str.lower().str.contains('no corresponde', na=False)) &
        (~df['m_cpes'].isin(no_llevan_cpe_valores) if 'm_cpes' in df.columns else True)
    ]) if total > 0 else 0

    # Calcular "Fletes OK": tienen CPE real + pesados + descargados (datos completos)
    fletes_ok = len(df[
        (df['tiene_cpe']) &
        (~df['cpe'].str.lower().str.contains('no corresponde', na=False)) &
        (df['tiene_pesadas']) &
        (df['tiene_descargas'])
    ])

    # Falta CPE = Total - Con CPE real - No llevan CPE (los que deberían tener CPE pero no lo tienen)
    falta_cpe = total - con_cpe_real - no_llevan_cpe

    # Falta Pesadas = celdas vacías en M Pesadas, excluyendo traslados internos
    # (los que NO tienen pesadas Y no son traslados internos)
    if 'm_cpes' in df.columns:
        falta_pesadas = len(df[
            (df['m_pesadas_num'].isna()) &
            (~df['m_cpes'].isin(['Traslado interno']))
        ]) if total > 0 else 0
    else:
        falta_pesadas = len(df[df['m_pesadas_num'].isna()]) if total > 0 else 0

    # Falta Descargas = celdas vacías en M Descargas, excluyendo traslados internos, flete en B y CPE terceros
    # (los que NO tienen descargas Y no son traslados ni en B ni CPE de terceros)
    if 'm_cpes' in df.columns:
        falta_descargas = len(df[
            (df['m_descargas_num'].isna()) &
            (~df['m_cpes'].isin(['Traslado interno', 'Flete en B', 'CPE Hecha por Terceros']))
        ]) if total > 0 else 0
    else:
        falta_descargas = len(df[df['m_descargas_num'].isna()]) if total > 0 else 0

    kpi_total = create_kpi_card("Total Fletes", f"{total:,}", "fa-truck", COLORS['primary'])
    kpi_ok = create_kpi_card("Fletes OK", f"{fletes_ok:,}", "fa-check-double", COLORS['success'],
                              f"{fletes_ok/total*100:.0f}% (CPE + Pesada + Descarga)" if total > 0 else "0%")
    kpi_con_cpe = create_kpi_card("Con CPE", f"{con_cpe_real:,}", "fa-file-alt", COLORS['success'],
                                   f"{con_cpe_real/total*100:.0f}%" if total > 0 else "0%")
    kpi_no_llevan_cpe = create_kpi_card("No llevan CPE", f"{no_llevan_cpe:,}", "fa-ban", COLORS['warning'],
                                         "Traslados / En B")
    kpi_pesados = create_kpi_card("Pesados Campo", f"{pesados:,}", "fa-weight", COLORS['secondary'])
    kpi_descargados = create_kpi_card("Descargados", f"{descargados:,}", "fa-warehouse", COLORS['primary'])
    kpi_falta_cpe = create_kpi_card("Falta CPE", f"{falta_cpe:,}", "fa-exclamation-circle", COLORS['danger'],
                                     "Pendientes de asignar")
    kpi_falta_pesadas = create_kpi_card("Falta Pesadas", f"{falta_pesadas:,}", "fa-weight", COLORS['danger'],
                                         "(no se consideran traslados int)")
    kpi_falta_descargas = create_kpi_card("Falta Descargas", f"{falta_descargas:,}", "fa-warehouse", COLORS['danger'],
                                           "(no se consideran traslados int/en B)")
    kpi_merma = create_kpi_card("Merma Prom.", f"{merma_prom:.2f}%", "fa-percentage",
                                 COLORS['danger'] if merma_prom > 0.3 else COLORS['success'])

    # Gráfico de merma por transportista (con kg totales)
    df_merma = df[df['tiene_pesadas'] & df['tiene_descargas']].copy()
    if not df_merma.empty:
        merma_trans = df_merma.groupby('transportista').agg({
            'merma_pct': 'mean',
            'merma_kg': 'sum',
            'numero_factura': 'count'
        }).reset_index()
        merma_trans.columns = ['transportista', 'merma_promedio', 'merma_kg_total', 'cantidad']
        merma_trans = merma_trans.sort_values('merma_promedio', ascending=True).tail(15)

        colors = ['#4CAF50' if x <= 0.3 else '#FFC107' if x <= 1 else '#F44336' for x in merma_trans['merma_promedio']]

        # Texto con % y kg totales
        text_labels = [f"{pct:.2f}% ({kg:,.0f} kg)" for pct, kg in zip(merma_trans['merma_promedio'], merma_trans['merma_kg_total'])]

        fig_merma = go.Figure(go.Bar(
            x=merma_trans['merma_promedio'], y=merma_trans['transportista'],
            orientation='h', marker_color=colors,
            text=text_labels, textposition='outside',
            hovertemplate="<b>%{y}</b><br>Merma: %{x:.2f}%<br>Total kg perdidos: %{customdata:,.0f}<br>Fletes: %{meta}<extra></extra>",
            customdata=merma_trans['merma_kg_total'],
            meta=merma_trans['cantidad']
        ))
        fig_merma.add_vline(x=0.3, line_dash="dash", line_color="red", annotation_text="0.3%")
        fig_merma.update_layout(margin=dict(l=20, r=120, t=20, b=20), xaxis_title="Merma %", showlegend=False, plot_bgcolor='white')
    else:
        fig_merma = go.Figure()
        fig_merma.add_annotation(text="Sin datos", showarrow=False, font_size=16)

    # Gráfico de cultivos
    if not df.empty:
        cultivos = df.groupby('producto').size().reset_index(name='cantidad')
        cultivos = cultivos[cultivos['producto'].notna() & (cultivos['producto'] != '')]
        if not cultivos.empty:
            fig_cultivos = px.pie(cultivos, values='cantidad', names='producto',
                                  color_discrete_sequence=px.colors.qualitative.Set3, hole=0.4)
            fig_cultivos.update_traces(textposition='inside', textinfo='percent+label')
            fig_cultivos.update_layout(margin=dict(l=20, r=20, t=20, b=20), showlegend=False)
        else:
            fig_cultivos = go.Figure()
            fig_cultivos.add_annotation(text="Sin datos", showarrow=False)
    else:
        fig_cultivos = go.Figure()
        fig_cultivos.add_annotation(text="Sin datos", showarrow=False)

    # Tabla de alertas
    df_alertas = df[(df['merma_sospechosa']) | (df['dif_facturacion'].abs() > 100) | (~df['tiene_cpe'])].copy()
    tabla_data = []
    for _, row in df_alertas.head(100).iterrows():
        # Calcular diferencia (Pesado - Descargado)
        if pd.notna(row['m_pesadas_num']) and pd.notna(row['m_descargas_num']):
            diferencia = row['m_pesadas_num'] - row['m_descargas_num']
            diferencia_str = f"{diferencia:,.0f}"
        else:
            diferencia_str = '-'

        tabla_data.append({
            'fecha': row['fecha'],
            'transportista': row['transportista'],
            'producto': row['producto'],
            'origen': row['origen'],
            'm_pesadas': f"{row['m_pesadas_num']:,.0f}" if pd.notna(row['m_pesadas_num']) else '-',
            'm_descargas': f"{row['m_descargas_num']:,.0f}" if pd.notna(row['m_descargas_num']) else '-',
            'diferencia_kg': diferencia_str,
            'merma_pct': f"{row['merma_pct']:.2f}" if pd.notna(row['merma_pct']) else '-',
            'cantidad': row['cantidad'],
            'dif_facturacion': f"{row['dif_facturacion']:,.0f}" if pd.notna(row['dif_facturacion']) else '-'
        })

    ultima_act = f"Última actualización: {datetime.now().strftime('%d/%m/%Y %H:%M')}"

    return (kpi_total, kpi_ok, kpi_con_cpe, kpi_no_llevan_cpe, kpi_pesados, kpi_descargados,
            kpi_falta_cpe, kpi_falta_pesadas, kpi_falta_descargas, kpi_merma,
            fig_merma, fig_cultivos, tabla_data, str(len(df_alertas)), ultima_act)


# Callback para el botón de autocomplete desde dashboard
@app.callback(
    Output('modal-resultado', 'is_open'),
    Output('modal-body', 'children'),
    [Input('btn-autocomplete', 'n_clicks'),
     Input('close-modal', 'n_clicks')],
    [State('modal-resultado', 'is_open')],
    prevent_initial_call=True
)
def toggle_modal(n_auto, n_close, is_open):
    ctx = callback_context
    if not ctx.triggered:
        return False, ""

    button_id = ctx.triggered[0]['prop_id'].split('.')[0]

    if button_id == 'btn-autocomplete' and n_auto:
        try:
            from agent_autocomplete import ejecutar_autocompletado
            resultado = ejecutar_autocompletado()
            if resultado['success']:
                body = dbc.Alert([
                    html.H5("Autocompletado exitoso"),
                    html.P(f"Campos completados: {resultado['campos_completados']}"),
                    html.P(f"Transportistas corregidos: {resultado.get('transportistas_corregidos', 0)}"),
                    html.P(f"Filas procesadas: {resultado['filas_procesadas']}"),
                ], color="success")
            else:
                body = dbc.Alert(f"Error: {resultado.get('error')}", color="danger")
        except Exception as e:
            body = dbc.Alert(f"Error: {str(e)}", color="danger")
        return True, body

    return False, ""


# Callback para cargar duplicados
@app.callback(
    [Output('resultado-duplicados', 'children'),
     Output('stats-duplicados', 'children'),
     Output('tabla-duplicados', 'children')],
    Input('btn-cargar-duplicados', 'n_clicks'),
    prevent_initial_call=True
)
def cargar_duplicados(n_clicks):
    if not n_clicks:
        return "", "", ""
    try:
        from app import cargar_cpes, get_credentials, normalizar_patente, normalizar_fecha, CONFIG
        import gspread

        creds = get_credentials()
        gc = gspread.authorize(creds)

        # Cargar CPEs (indexado por patente)
        cpes = cargar_cpes(gc)

        # Filtrar solo patentes con múltiples CPEs
        patentes_duplicadas = {pat: lista for pat, lista in cpes.items() if len(lista) > 1}

        if not patentes_duplicadas:
            return (
                dbc.Alert("No hay patentes duplicadas.", color="success"),
                html.P("0 patentes con múltiples CPEs", className="text-muted"),
                html.P("No hay duplicados para mostrar.", className="text-muted text-center py-4")
            )

        # Cargar Pesadas para encontrar las que tienen patente duplicada y NO coinciden en fecha
        ss_pesadas = gc.open_by_key(CONFIG['PESADAS_SPREADSHEET_ID'])
        hoja_pesadas = ss_pesadas.worksheet(CONFIG['PESADAS_SHEET_NAME'])
        datos_pesadas = hoja_pesadas.get_all_values()

        # Primero: recolectar todos los CPEs que ya fueron verificados (asignados a filas con OK)
        cpes_verificados = set()
        for fila in datos_pesadas[1:]:
            verificado = fila[CONFIG['PESADAS_COL_VERIFICADO']] if len(fila) > CONFIG['PESADAS_COL_VERIFICADO'] else ''
            cpe_asignado = fila[CONFIG['PESADAS_COL_CPE']] if len(fila) > CONFIG['PESADAS_COL_CPE'] else ''
            if verificado and str(verificado).strip().upper() == 'OK' and cpe_asignado:
                cpes_verificados.add(str(cpe_asignado).strip())

        # Buscar pesadas con patente duplicada donde la fecha NO coincide exactamente
        casos_revisar = []  # Lista de {fila, patente, fecha_pesada, cpe_asignado, cpes_disponibles}

        for idx, fila in enumerate(datos_pesadas[1:], start=2):
            if len(fila) <= CONFIG['PESADAS_COL_PATENTE']:
                continue

            fecha_pesada = fila[CONFIG['PESADAS_COL_FECHA']] if len(fila) > CONFIG['PESADAS_COL_FECHA'] else ''
            patente_pesada = fila[CONFIG['PESADAS_COL_PATENTE']] if len(fila) > CONFIG['PESADAS_COL_PATENTE'] else ''
            cpe_asignado = fila[CONFIG['PESADAS_COL_CPE']] if len(fila) > CONFIG['PESADAS_COL_CPE'] else ''
            producto_pesada = fila[CONFIG['PESADAS_COL_PRODUCTO']] if len(fila) > CONFIG['PESADAS_COL_PRODUCTO'] else ''

            if not patente_pesada or not cpe_asignado:
                continue

            # Verificar si ya fue marcado como OK
            verificado = fila[CONFIG['PESADAS_COL_VERIFICADO']] if len(fila) > CONFIG['PESADAS_COL_VERIFICADO'] else ''
            if verificado and str(verificado).strip().upper() == 'OK':
                continue

            patente_norm = normalizar_patente(patente_pesada)
            fecha_pesada_norm = normalizar_fecha(fecha_pesada)

            # Solo nos interesan las patentes duplicadas
            if patente_norm not in patentes_duplicadas:
                continue

            lista_cpes_original = patentes_duplicadas[patente_norm]

            # Filtrar CPEs que ya fueron verificados (asignados a otras pesadas con OK)
            lista_cpes = [cpe for cpe in lista_cpes_original
                         if cpe['numero_cpe'] not in cpes_verificados]

            # Si quedó solo 1 CPE después de filtrar verificados, ya no hay ambigüedad
            if len(lista_cpes) <= 1:
                continue

            # Normalizar producto de la pesada
            from app import normalizar_producto
            producto_norm = normalizar_producto(producto_pesada)

            # Buscar si algún CPE tiene fecha EXACTA con la pesada
            tiene_match_exacto_fecha = any(cpe['fecha'] == fecha_pesada_norm for cpe in lista_cpes)

            # Si hay match exacto de fecha, no hay ambigüedad
            if tiene_match_exacto_fecha:
                continue

            # Buscar cuántos CPEs tienen el mismo producto
            cpes_mismo_producto = [cpe for cpe in lista_cpes if cpe['grano_tipo'] == producto_norm]

            # Si hay exactamente 1 CPE con ese producto, el producto desempata -> no hay ambigüedad
            if len(cpes_mismo_producto) == 1:
                continue

            # SOLO agregar si hay ambigüedad real:
            # - No hay match exacto de fecha Y
            # - Hay >1 CPEs con el mismo producto (si hay 0, no hay opciones válidas)
            # Solo mostrar CPEs que coinciden en producto (los otros no aplican)
            if len(cpes_mismo_producto) > 1:
                casos_revisar.append({
                    'fila': idx,
                    'patente': patente_pesada,
                    'fecha_pesada': fecha_pesada,
                    'producto_pesada': producto_pesada,
                    'cpe_asignado': cpe_asignado,
                    'cpes_disponibles': cpes_mismo_producto,  # Solo CPEs del mismo producto
                    'cpes_mismo_producto': len(cpes_mismo_producto)
                })

        if not casos_revisar:
            return (
                dbc.Alert([
                    html.I(className="fas fa-check-circle me-2"),
                    "Todas las asignaciones se resolvieron por fecha exacta o producto único."
                ], color="success"),
                html.Div([
                    dbc.Row([
                        dbc.Col([
                            html.H4(str(len(patentes_duplicadas)), className="text-warning mb-0"),
                            html.Small("Patentes duplicadas", className="text-muted")
                        ], className="text-center"),
                        dbc.Col([
                            html.H4("0", className="text-success mb-0"),
                            html.Small("Casos a revisar", className="text-muted")
                        ], className="text-center"),
                    ])
                ]),
                html.P("No hay casos que necesiten revisión manual.", className="text-muted text-center py-4")
            )

        # Estadísticas
        stats = html.Div([
            dbc.Row([
                dbc.Col([
                    html.H4(str(len(patentes_duplicadas)), className="text-warning mb-0"),
                    html.Small("Patentes duplicadas", className="text-muted")
                ], className="text-center"),
                dbc.Col([
                    html.H4(str(len(casos_revisar)), className="text-danger mb-0"),
                    html.Small("Casos a revisar", className="text-muted")
                ], className="text-center"),
            ])
        ])

        # URL base del sheet con gid para Pesadas Todos
        PESADAS_SHEET_BASE = "https://docs.google.com/spreadsheets/d/1gTvXfwOsqbbc5lxpcsh8HMoB5F3Bix0qpdNKdyY5DME/edit#gid=0&range=A"

        # Construir tabla con casos a revisar
        filas_tabla = []
        for caso in casos_revisar:
            # Formatear los CPEs disponibles
            cpes_texto = []
            for cpe_data in caso['cpes_disponibles']:
                cpes_texto.append(f"{cpe_data['numero_cpe']} ({cpe_data['fecha']}, {cpe_data['grano_tipo'] or '-'})")

            fila_num = caso['fila']

            filas_tabla.append(html.Tr([
                html.Td(str(fila_num), style={'fontWeight': 'bold'}),
                html.Td(caso['patente']),
                html.Td(caso['fecha_pesada']),
                html.Td(caso['producto_pesada']),
                html.Td(caso['cpe_asignado'], className="text-primary fw-bold"),
                html.Td(html.Ul([html.Li(t, style={'fontSize': '0.85em'}) for t in cpes_texto],
                               style={'marginBottom': '0', 'paddingLeft': '1rem'})),
                html.Td([
                    dbc.Button("OK", id={'type': 'btn-ok-duplicado', 'index': fila_num},
                              color="success", size="sm", className="me-1"),
                    html.A(
                        dbc.Button("Ir", color="primary", size="sm", outline=True),
                        href=f"{PESADAS_SHEET_BASE}{fila_num}",
                        target="_blank"
                    )
                ], style={'whiteSpace': 'nowrap'})
            ]))

        tabla = dbc.Table([
            html.Thead(html.Tr([
                html.Th("Fila"),
                html.Th("Patente"),
                html.Th("Fecha Pesada"),
                html.Th("Producto"),
                html.Th("CPE Asignado"),
                html.Th("CPEs Disponibles"),
                html.Th("Acciones")
            ])),
            html.Tbody(filas_tabla)
        ], striped=True, bordered=True, hover=True, responsive=True, size="sm")

        return (
            dbc.Alert([
                html.I(className="fas fa-exclamation-triangle me-2"),
                f"{len(casos_revisar)} pesadas sin match exacto de fecha - revisar manualmente"
            ], color="warning"),
            stats,
            tabla
        )

    except Exception as e:
        import traceback
        return (
            dbc.Alert(f"Error: {str(e)}\n{traceback.format_exc()}", color="danger"),
            "",
            ""
        )


# Callback para marcar duplicado como OK
@app.callback(
    Output('resultado-duplicados', 'children', allow_duplicate=True),
    Input({'type': 'btn-ok-duplicado', 'index': ALL}, 'n_clicks'),
    State({'type': 'btn-ok-duplicado', 'index': ALL}, 'id'),
    prevent_initial_call=True
)
def marcar_ok_duplicado(n_clicks_list, ids):
    from dash import callback_context
    if not callback_context.triggered:
        return dash.no_update

    # Encontrar qué botón se clickeó
    triggered = callback_context.triggered[0]
    if not triggered['value']:
        return dash.no_update

    # Extraer el índice (número de fila) del botón clickeado
    prop_id = triggered['prop_id']
    # prop_id es algo como '{"index":94,"type":"btn-ok-duplicado"}.n_clicks'
    import json
    button_id = json.loads(prop_id.split('.')[0])
    fila_num = button_id['index']

    try:
        from app import get_credentials, CONFIG
        import gspread

        creds = get_credentials()
        gc = gspread.authorize(creds)

        # Abrir hoja de Pesadas
        ss_pesadas = gc.open_by_key(CONFIG['PESADAS_SPREADSHEET_ID'])
        hoja_pesadas = ss_pesadas.worksheet(CONFIG['PESADAS_SHEET_NAME'])

        # Escribir OK en la columna de verificado (columna T = 20 en base 1)
        col_verificado = CONFIG['PESADAS_COL_VERIFICADO'] + 1  # +1 porque gspread usa base 1
        celda = gspread.utils.rowcol_to_a1(fila_num, col_verificado)
        hoja_pesadas.update(celda, [['OK']])

        return dbc.Alert([
            html.I(className="fas fa-check me-2"),
            f"Fila {fila_num} marcada como OK. Recargá la lista para ver los cambios."
        ], color="success", dismissable=True)

    except Exception as e:
        return dbc.Alert(f"Error al marcar OK: {str(e)}", color="danger")


# Callback para cargar fletes sin CPE
@app.callback(
    [Output('resultado-cargar-sin-cpe', 'children'),
     Output('stats-sin-cpe', 'children'),
     Output('tabla-sin-cpe', 'children'),
     Output('badge-sin-cpe', 'children')],
    Input('btn-cargar-sin-cpe', 'n_clicks'),
    prevent_initial_call=True
)
def cargar_fletes_sin_cpe(n_clicks):
    if not n_clicks:
        return "", "", "", "0"
    try:
        from app import get_credentials, CONFIG
        import gspread

        creds = get_credentials()
        gc = gspread.authorize(creds)

        # Abrir hoja de Fletes
        ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
        hoja_fletes = ss.worksheet(CONFIG['FLETES_SHEET_NAME'])
        datos_fletes = hoja_fletes.get_all_values()

        # URL base del sheet de Fletes
        FLETES_SHEET_URL = f"https://docs.google.com/spreadsheets/d/{CONFIG['CPE_SPREADSHEET_ID']}/edit#gid=0&range=A"

        # Opciones para el dropdown
        OPCIONES_SIN_CPE = [
            {'label': 'Seleccionar...', 'value': ''},
            {'label': 'Traslado interno', 'value': 'Traslado interno'},
            {'label': 'Flete en B', 'value': 'Flete en B'},
            {'label': 'CPE Hecha por Terceros', 'value': 'CPE Hecha por Terceros'},
            {'label': 'Sin documentación', 'value': 'Sin documentación'},
            {'label': 'Pendiente de CPE', 'value': 'Pendiente de CPE'},
            {'label': 'Error de carga', 'value': 'Error de carga'},
        ]

        # Obtener índices de columnas por nombre del header
        headers = datos_fletes[0] if datos_fletes else []
        def get_col_idx(nombre):
            try:
                return headers.index(nombre)
            except ValueError:
                return -1

        col_fecha = get_col_idx('Fecha')
        col_transportista = get_col_idx('Transportista')
        col_producto = get_col_idx('Producto')
        col_ctg = get_col_idx('CTG')
        col_cantidad = get_col_idx('Cantidad')
        col_cpe = get_col_idx('CPE')
        col_m_cpes = get_col_idx("M CPE's")

        # Buscar fletes sin CPE (columna M CPE's = "no" o vacía, y columna CPE vacía)
        fletes_sin_cpe = []
        for idx, fila in enumerate(datos_fletes[1:], start=2):
            if len(fila) <= max(col_cpe, col_m_cpes):
                continue

            cpe = fila[col_cpe] if col_cpe >= 0 and len(fila) > col_cpe else ''
            m_cpes = fila[col_m_cpes] if col_m_cpes >= 0 and len(fila) > col_m_cpes else ''

            # Solo mostrar los que tienen "no" o están vacíos (no los ya clasificados)
            m_cpes_lower = str(m_cpes).strip().lower()
            if m_cpes_lower == 'no' or (m_cpes_lower == '' and cpe == ''):
                # Obtener info adicional para mostrar
                fecha = fila[col_fecha] if col_fecha >= 0 and len(fila) > col_fecha else ''
                transportista = fila[col_transportista] if col_transportista >= 0 and len(fila) > col_transportista else ''
                producto = fila[col_producto] if col_producto >= 0 and len(fila) > col_producto else ''
                ctg = fila[col_ctg] if col_ctg >= 0 and len(fila) > col_ctg else ''
                cantidad = fila[col_cantidad] if col_cantidad >= 0 and len(fila) > col_cantidad else ''

                fletes_sin_cpe.append({
                    'fila': idx,
                    'fecha': fecha,
                    'transportista': transportista,
                    'producto': producto,
                    'ctg': ctg,
                    'cantidad': cantidad,
                    'cpe': cpe,
                    'm_cpes': m_cpes
                })

        if not fletes_sin_cpe:
            return (
                dbc.Alert([
                    html.I(className="fas fa-check-circle me-2"),
                    "No hay fletes pendientes de clasificar."
                ], color="success"),
                html.P("Todos los fletes están clasificados.", className="text-success mb-0"),
                html.P("No hay fletes sin CPE pendientes.", className="text-muted text-center py-4"),
                "0"
            )

        # Estadísticas
        stats = html.Div([
            dbc.Row([
                dbc.Col([
                    html.H4(str(len(fletes_sin_cpe)), className="text-danger mb-0"),
                    html.Small("Fletes sin CPE", className="text-muted")
                ], className="text-center"),
                dbc.Col([
                    html.H4(str(len([f for f in fletes_sin_cpe if f['m_cpes'].lower() == 'no'])), className="text-warning mb-0"),
                    html.Small("Marcados 'no'", className="text-muted")
                ], className="text-center"),
                dbc.Col([
                    html.H4(str(len([f for f in fletes_sin_cpe if f['m_cpes'] == ''])), className="text-secondary mb-0"),
                    html.Small("Sin clasificar", className="text-muted")
                ], className="text-center"),
            ])
        ])

        # Construir tabla
        filas_tabla = []
        for flete in fletes_sin_cpe[:100]:  # Limitar a 100 para no sobrecargar
            fila_num = flete['fila']

            filas_tabla.append(html.Tr([
                html.Td(str(fila_num), style={'fontWeight': 'bold'}),
                html.Td(flete['fecha']),
                html.Td(flete['transportista'], style={'maxWidth': '200px', 'overflow': 'hidden', 'textOverflow': 'ellipsis', 'whiteSpace': 'nowrap'}),
                html.Td(flete['producto']),
                html.Td(flete['ctg']),
                html.Td(flete['cantidad']),
                html.Td(
                    dcc.Dropdown(
                        id={'type': 'dropdown-sin-cpe', 'index': fila_num},
                        options=OPCIONES_SIN_CPE,
                        value='',
                        placeholder='Clasificar...',
                        style={'minWidth': '160px'},
                        clearable=False
                    )
                ),
                html.Td([
                    html.A(
                        dbc.Button([html.I(className="fas fa-external-link-alt")],
                                  color="primary", size="sm", outline=True),
                        href=f"{FLETES_SHEET_URL}{fila_num}",
                        target="_blank",
                        title="Ver en Google Sheets"
                    )
                ])
            ]))

        tabla = dbc.Table([
            html.Thead(html.Tr([
                html.Th("Fila"),
                html.Th("Fecha"),
                html.Th("Transportista"),
                html.Th("Producto"),
                html.Th("CTG"),
                html.Th("Cantidad"),
                html.Th("Clasificar"),
                html.Th("Ver")
            ])),
            html.Tbody(filas_tabla)
        ], striped=True, bordered=True, hover=True, responsive=True, size="sm")

        return (
            dbc.Alert([
                html.I(className="fas fa-list me-2"),
                f"Se encontraron {len(fletes_sin_cpe)} fletes sin CPE"
            ], color="info"),
            stats,
            tabla,
            str(len(fletes_sin_cpe))
        )

    except Exception as e:
        import traceback
        return (
            dbc.Alert(f"Error: {str(e)}", color="danger"),
            "",
            "",
            "0"
        )


# Callback para clasificar un flete sin CPE
@app.callback(
    Output('resultado-cargar-sin-cpe', 'children', allow_duplicate=True),
    Input({'type': 'dropdown-sin-cpe', 'index': ALL}, 'value'),
    State({'type': 'dropdown-sin-cpe', 'index': ALL}, 'id'),
    prevent_initial_call=True
)
def clasificar_flete_sin_cpe(values, ids):
    from dash import callback_context
    if not callback_context.triggered:
        return dash.no_update

    # Encontrar qué dropdown cambió
    triggered = callback_context.triggered[0]
    prop_id = triggered['prop_id']
    valor_seleccionado = triggered['value']

    # Si no se seleccionó nada, ignorar
    if not valor_seleccionado:
        return dash.no_update

    # Extraer el índice (número de fila) del dropdown
    import json
    try:
        dropdown_id = json.loads(prop_id.split('.')[0])
        fila_num = dropdown_id['index']
    except:
        return dash.no_update

    try:
        from app import get_credentials, CONFIG
        import gspread

        creds = get_credentials()
        gc = gspread.authorize(creds)

        # Abrir hoja de Fletes
        ss = gc.open_by_key(CONFIG['CPE_SPREADSHEET_ID'])
        hoja_fletes = ss.worksheet(CONFIG['FLETES_SHEET_NAME'])

        # Obtener headers para encontrar columnas por nombre
        headers = hoja_fletes.row_values(1)

        def get_col_num(nombre):
            try:
                return headers.index(nombre) + 1  # Base 1 para gspread
            except ValueError:
                return -1

        col_cpe = get_col_num('CPE')
        col_m_cpes = get_col_num("M CPE's")

        if col_cpe < 1 or col_m_cpes < 1:
            return dbc.Alert("Error: No se encontraron las columnas CPE o M CPE's", color="danger")

        # Actualizar columna CPE con "No corresponde"
        celda_cpe = gspread.utils.rowcol_to_a1(fila_num, col_cpe)
        hoja_fletes.update(celda_cpe, [['No corresponde']])

        # Actualizar columna M CPE's con el motivo seleccionado
        celda_m_cpes = gspread.utils.rowcol_to_a1(fila_num, col_m_cpes)
        hoja_fletes.update(celda_m_cpes, [[valor_seleccionado]])

        return dbc.Alert([
            html.I(className="fas fa-check me-2"),
            f"Fila {fila_num} clasificada como '{valor_seleccionado}'. ",
            html.Small("Recargá la lista para actualizar.", className="text-muted")
        ], color="success", dismissable=True)

    except Exception as e:
        return dbc.Alert(f"Error al clasificar: {str(e)}", color="danger")


if __name__ == '__main__':
    print("=" * 50)
    print("Dashboard de Control de Fletes Agropecuario")
    print("=" * 50)
    print("Servidor corriendo en: http://localhost:5016")
    print("  - Dashboard: http://localhost:5016/")
    print("  - Procesadores: http://localhost:5016/procesadores")
    print("  - Trabajo Manual: http://localhost:5016/trabajo-manual")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5016, debug=True)
