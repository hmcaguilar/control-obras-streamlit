"""
Control de Obras - Dashboard Ejecutivo
============================================
Requiere:  pip install streamlit plotly pandas openpyxl
Ejecutar:  streamlit run Dashboard_Ejecutivo.py
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import warnings
warnings.filterwarnings("ignore")
from pathlib import Path
from datetime import datetime

# Ruta absoluta al Excel - funciona sin importar desde dónde se lanza streamlit
BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / "Concentrado_Sintetico_10Anios.xlsx"

# -- Paleta azul ejecutiva -------------------------------------------------------
ACCENT       = "#4A7FB5"   # Azul acero – marca, KPIs neutros
ACCENT_LIGHT = "#6A9FD5"   # Azul claro para gradientes
ACCENT_DIM   = "#3A6595"   # Azul profundo para contraste
POSITIVE     = "#4F9D8C"   # Verde azulado – indicadores positivos
POSITIVE_SOFT= "#6BB8A8"   # Verde azulado suave
NEUTRAL      = "#8697AC"   # Gris azulado – tesorería/neutral
NEUTRAL_DIM  = "#5E7391"   # Azul grisáceo profundo
ALERT        = "#A55B5B"   # Burdeos apagado – alertas (NO rojo puro)
ALERT_SOFT   = "#C07878"   # Burdeos suave
CHART_1      = "#5E7391"   # Gráfico extra 1
CHART_2      = "#3A5876"   # Gráfico extra 2
BG_DARK      = "#0F1620"   # Fondo principal grafito-azulado
BG_CARD      = "#182230"   # Fondo cards/sidebar
WHITE        = "#E8EDF3"   # Texto principal
OFF_WHITE    = "#C8D2DE"   # Texto secundario
MUTED        = "#7B8DA1"   # Texto apagado/labels

# Compatibilidad con nombres usados en el script
VINO         = CHART_2
VINO_MID     = ACCENT
GOLD         = ACCENT       # Donde antes era dorado, ahora azul acero
GOLD_SOFT    = "#D0DAE8"    # Tonalidad azul muy clara
GREEN        = POSITIVE
GREEN_SOFT   = "#C0D8D2"   # Verde-azulado claro
RED          = ALERT
BLUE         = ACCENT
TEAL         = "#4F9D8C"
ROSE         = NEUTRAL
CLAY         = ACCENT_LIGHT
SAGE         = NEUTRAL_DIM
STONE        = NEUTRAL
SAND         = "#D0DAE8"
EARTH        = ACCENT
EARTH_DARK   = CHART_2
GRAY         = NEUTRAL
AZUL_NOCHE   = BG_DARK

PALETTE = [ACCENT, POSITIVE, NEUTRAL_DIM, CHART_2, NEUTRAL, ACCENT_LIGHT, "#7BAFD4", CHART_1, "#2E4A66", "#4F9D8C"]

# -- Configuración de la página -------------------------------------------------
st.set_page_config(
    page_title="Control de Obras",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# -- CSS institucional -----------------------------------------------------------
st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono&display=swap');

  html, body, [class*="css"] {{
    font-family: 'DM Sans', sans-serif;
    background-color: {BG_DARK};
    color: {OFF_WHITE};
  }}

  /* Sidebar */
  [data-testid="stSidebar"] {{
    background: {BG_CARD} !important;
    border-right: 1px solid rgba(74,127,181,0.20);
  }}
  [data-testid="stSidebar"] * {{ color: {OFF_WHITE} !important; }}

  /* Encabezado */
  .brand-header {{
    background: linear-gradient(135deg, rgba(74,127,181,0.14) 0%, rgba(24,34,48,0.9) 100%);
    padding: 20px 28px;
    border-radius: 12px;
    margin-bottom: 24px;
    border: 1px solid rgba(74,127,181,0.25);
    display: flex;
    align-items: center;
    gap: 16px;
  }}
  .brand-mark {{
    width: 52px;
    height: 52px;
    border-radius: 14px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg, {ACCENT} 0%, {ACCENT_LIGHT} 100%);
    color: {BG_DARK};
    font-family: 'Bebas Neue', sans-serif;
    font-size: 28px;
    font-weight: 700;
    letter-spacing: 1px;
    box-shadow: 0 8px 20px rgba(74,127,181,0.22);
  }}
  .brand-title {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 28px;
    color: {WHITE};
    letter-spacing: 2px;
    line-height: 1;
  }}
  .brand-subtitle {{
    font-size: 13px;
    color: rgba(200,210,222,0.8);
    margin-top: 4px;
  }}

  /* Tarjetas KPI */
  .kpi-card {{
    background: rgba(24,34,48,0.7);
    border: 1px solid rgba(74,127,181,0.18);
    border-radius: 12px;
    padding: 20px 22px;
    border-top: 3px solid {ACCENT};
  }}
  .kpi-card-green  {{ border-top-color: {POSITIVE}; }}
  .kpi-card-red    {{ border-top-color: {ALERT}; }}
  .kpi-card-blue   {{ border-top-color: {ACCENT}; }}
  .kpi-card-teal   {{ border-top-color: {TEAL}; }}
  .kpi-label {{
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: {MUTED};
    margin-bottom: 6px;
  }}
  .kpi-value {{
    font-family: 'Bebas Neue', sans-serif;
    font-size: 32px;
    line-height: 1;
    color: {ACCENT};
  }}
  .kpi-value-green {{ color: {POSITIVE}; }}
  .kpi-value-red   {{ color: {ALERT}; }}
  .kpi-value-blue  {{ color: {ACCENT}; }}
  .kpi-value-white {{ color: {WHITE}; }}
  .kpi-delta {{ font-size: 12px; margin-top: 4px; color: {MUTED}; }}
  .kpi-delta-up   {{ color: {POSITIVE}; }}
  .kpi-delta-down {{ color: {ALERT}; }}

  /* Caja de alerta */
  .alert-box {{
    background: rgba(165,91,91,0.12);
    border: 1px solid rgba(165,91,91,0.30);
    border-radius: 8px;
    padding: 10px 14px;
    font-size: 13px;
    color: {ALERT_SOFT};
    margin: 8px 0;
  }}
  .tip-box {{
    background: rgba(74,127,181,0.08);
    border: 1px solid rgba(74,127,181,0.22);
    border-radius: 8px;
    padding: 10px 14px;
    font-size: 13px;
    color: {OFF_WHITE};
    margin: 8px 0;
  }}

  /* Main content */
  .main .block-container {{ padding: 1.5rem 2rem; max-width: 1600px; }}

  /* Plotly charts dark bg */
  .stPlotlyChart {{ border-radius: 12px; }}

  /* Divider */
  hr {{ border-color: rgba(255,255,255,0.07) !important; }}

  /* Metric overrides */
  [data-testid="stMetric"] {{
    background: rgba(24,34,48,0.7);
    border: 1px solid rgba(74,127,181,0.18);
    border-radius: 12px;
    padding: 16px !important;
  }}
  [data-testid="stMetricLabel"] {{ color: {MUTED} !important; font-size: 11px !important; }}
  [data-testid="stMetricValue"] {{ color: {ACCENT} !important; font-family: 'Bebas Neue' !important; font-size: 28px !important; }}
  [data-testid="stMetricDelta"] {{ font-size: 12px !important; }}

  h1, h2, h3 {{ font-family: 'Bebas Neue', sans-serif !important; letter-spacing: 2px !important; color: {WHITE} !important; }}
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# CARGA DE DATOS
# ==============================================================================

@st.cache_data(show_spinner="Cargando datos del modelo...")
def load_data(path):
    xl = pd.ExcelFile(str(path))

    # Helper: leer hoja, eliminar columnas sin nombre y limpiar texto
    def read(sheet, parse_dates=None):
        df = pd.read_excel(xl, sheet_name=sheet, parse_dates=parse_dates)
        df = df.loc[:, ~df.columns.str.startswith("Unnamed")]
        for col in df.select_dtypes(include="object").columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r"\n", " ", regex=True)
            df[col] = df[col].replace({"nan": None, "None": None})
        return df

    obras_df       = read("Obras")
    costos_df      = read("Costos",      parse_dates=["FECHA"])
    egresos_df     = read("Egresos",     parse_dates=["FECHA", "FECHA PAGO"])
    facturacion_df = read("Facturacion", parse_dates=["FECHA FACT.", "FECHA TENT. PAGO"])
    ingresos_df    = read("Ingresos",    parse_dates=["FECHA DE PAGO"])
    cxc_df         = read("CxC",         parse_dates=["FECHA FACT."])
    cxp_df         = read("CxP",         parse_dates=["FECHA"])
    contratos_df   = read("Contratos",   parse_dates=["FECHA INICIO", "FECHA FIN"])

    # Normalizar nombres de columnas
    obras_df.rename(columns={"No. Y NOMBRE OBRA": "OBRA"}, inplace=True)
    contratos_df.rename(columns={"No. Y NOMBRE OBRA": "OBRA"}, inplace=True)

    # Búsqueda: obra -> portafolio / estatus
    obra_map = obras_df.set_index("OBRA")[["PORTAFOLIO", "ESTATUS", "CLIENTE"]].to_dict("index")

    def add_port(df, obra_col="OBRA"):
        df["PORTAFOLIO"] = df[obra_col].map(
            lambda o: obra_map.get(o, {}).get("PORTAFOLIO", "OTROS") if o else "OTROS"
        )
        df["ESTATUS_OBRA"] = df[obra_col].map(
            lambda o: obra_map.get(o, {}).get("ESTATUS", "") if o else ""
        )
        return df

    costos_df      = add_port(costos_df)
    egresos_df     = add_port(egresos_df)
    facturacion_df = add_port(facturacion_df)
    ingresos_df    = add_port(ingresos_df)
    cxc_df         = add_port(cxc_df)
    cxp_df         = add_port(cxp_df)
    contratos_df   = add_port(contratos_df)

    # Helpers de fechas
    for df, col in [(costos_df, "FECHA"), (facturacion_df, "FECHA FACT."),
                    (ingresos_df, "FECHA DE PAGO"), (egresos_df, "FECHA PAGO"),
                    (cxp_df, "FECHA"), (cxc_df, "FECHA FACT."), (contratos_df, "FECHA INICIO")]:
        df["AÑO"]    = pd.to_datetime(df[col], errors="coerce").dt.year
        df["MES"]    = pd.to_datetime(df[col], errors="coerce").dt.month
        df["YEARMON"] = (pd.to_datetime(df[col], errors="coerce")
                         .dt.to_period("M").astype(str))

    # AÑO_ORIGEN en Egresos: año en que se INCURRIÓ el costo (columna FECHA),
    # distinto de "AÑO" que usa FECHA PAGO. Egresos["AÑO"] es correcto para
    # las vistas de flujo de caja (¿cuándo salió el dinero?), pero comparar
    # "Costo incurrido" (Costos, por FECHA) contra "Pagado" filtrado por FECHA
    # PAGO mezcla dos ventanas de tiempo distintas y puede dar >100% cuando se
    # están liquidando pagos rezagados de periodos anteriores. Esta columna se
    # usa exclusivamente para el ratio "Costos Pagados".
    egresos_df["AÑO_ORIGEN"] = pd.to_datetime(egresos_df["FECHA"], errors="coerce").dt.year

    # Seguridad numérica
    for df, col in [(costos_df, "TOTAL"), (facturacion_df, "IMPORTE"),
                    (facturacion_df, "ESTIMADO"), (ingresos_df, "INGRESOS"),
                    (cxc_df, "CxC"), (cxp_df, "TOTAL"), (egresos_df, "IMPORTE PAGADO")]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    for col in ["IMPORTE S/IVA", "IMPORTE C/IVA", "% ANTICIPO", "% F.G."]:
        contratos_df[col] = pd.to_numeric(contratos_df.get(col, 0), errors="coerce").fillna(0)

    # "% ANTICIPO" y "% F.G." vienen capturados como número entero (15 = 15%),
    # no como fracción (0.15). Se normalizan aquí, en el punto de carga, para
    # que cualquier cálculo que use estas columnas —hoy solo Anticipos/FG en
    # Backlog & Contratos, pero también a futuro— reciba la fracción correcta
    # y no un valor 100x inflado.
    for col in ["% ANTICIPO", "% F.G."]:
        contratos_df[col] = contratos_df[col] / 100

    return dict(
        obras=obras_df, costos=costos_df, egresos=egresos_df,
        facturacion=facturacion_df, ingresos=ingresos_df,
        cxc=cxc_df, cxp=cxp_df, contratos=contratos_df,
        obra_map=obra_map
    )

# -- Carga ---------------------------------------------------------------------
try:
    data = load_data(EXCEL_PATH)
except FileNotFoundError:
    st.error(f"⚠️ No se encontró el archivo **Concentrado_Optimizado.xlsx**.\n\n"
             f"Buscado en: `{EXCEL_PATH}`\n\n"
             "Asegúrate de que el Excel esté en la **misma carpeta** que `Dashboard_Ejecutivo.py`.")
    st.stop()

obras       = data["obras"]
costos      = data["costos"]
egresos     = data["egresos"]
facturacion = data["facturacion"]
ingresos    = data["ingresos"]
cxc         = data["cxc"]
cxp         = data["cxp"]
contratos   = data["contratos"]

# Metadatos del reporte
REPORT_CREATED = datetime.now()
REPORT_CREATED_STR = REPORT_CREATED.strftime("%d %b %Y")
LATEST_COSTOS_DATE = pd.to_datetime(costos["FECHA"], errors="coerce").max()
REPORT_CUTOFF_STR = (LATEST_COSTOS_DATE.strftime("%d %b %Y")
                     if not pd.isna(LATEST_COSTOS_DATE) else "N/A")

LAYOUT_BASE = dict(
    paper_bgcolor="rgba(15,22,32,0)",
    plot_bgcolor="rgba(15,22,32,0)",
    font=dict(family="DM Sans, Arial", color=OFF_WHITE, size=12),
    margin=dict(l=8, r=8, t=36, b=8),
    legend=dict(bgcolor="rgba(0,0,0,0)", font=dict(size=11)),
    colorway=PALETTE,
)

def apply_layout(fig, title="", height=320):
    fig.update_layout(
        title=dict(text=title, font=dict(size=13, color=OFF_WHITE), x=0),
        height=height,
        **LAYOUT_BASE
    )
    fig.update_xaxes(gridcolor="rgba(255,255,255,0.05)", linecolor="rgba(255,255,255,0.1)")
    fig.update_yaxes(gridcolor="rgba(255,255,255,0.05)", linecolor="rgba(255,255,255,0.1)")
    return fig

# ==============================================================================
# SIDEBAR - FILTROS
# ==============================================================================

with st.sidebar:
    logo_path = BASE_DIR / "Logo_cAc_hueso_sin fondo.png"
    st.image(str(logo_path), width=110)
    st.markdown(
        "<div style='text-align:center;font-size:11px;color:rgba(200,210,222,0.8);letter-spacing:2px;padding-top:10px'>CONTROL DE OBRAS</div>",
        unsafe_allow_html=True,
    )
    st.divider()

    # Selector de página
    page = st.radio(
        "📋 Página",
        ["🎯 Executive Summary", "💰 P&L Ejecutivo", "📊 Backlog & Contratos",
         "📈 Facturación & Cobro", "⚙️ Costos & Margen",
         "💳 Flujo de Caja", "🌊 Sankey · Costos"],
        label_visibility="collapsed"
    )
    st.divider()

    # Filtro de año
    all_years = sorted(facturacion["AÑO"].dropna().unique().astype(int), reverse=True)
    sel_years = st.multiselect("📅 Año", all_years, default=all_years[:3],
                               help="Selecciona uno o varios años")

    # Filtro de portafolio
    all_ports = sorted(obras["PORTAFOLIO"].dropna().unique())
    sel_ports = st.multiselect("🏗️ Portafolio", all_ports, default=[],
                               help="Vacío = todos")

    # Filtro de estatus
    sel_estatus = st.multiselect("🔖 Estatus", ["VIGENTE", "EN CIERRE", "CERRADA", "NO APLICA"],
                                 default=["VIGENTE"])

    st.divider()
    st.caption(
        f"Control de Obras\nFecha corte: {REPORT_CREATED_STR} · Reporte generado: {REPORT_CREATED_STR} · Por Hector Calderon"
    )
#"REPORT_CUTOFF_STR" utilizar para la fecha de corte real, ya que ls data que use llega posterior a la fecha de creacion
# -- Aplicar filtros -----------------------------------------------------------
def filter_df(df, year_col="AÑO", port_col="PORTAFOLIO", status_col="ESTATUS_OBRA"):
    d = df.copy()
    if sel_years and year_col in d.columns:
        d = d[d[year_col].isin(sel_years)]
    if sel_ports and port_col in d.columns:
        d = d[d[port_col].isin(sel_ports)]
    if sel_estatus and status_col in d.columns:
        d = d[d[status_col].isin(sel_estatus)]
    return d

f_fact = filter_df(facturacion)
f_ing  = filter_df(ingresos)
f_cos  = filter_df(costos)
f_cxc  = filter_df(cxc, year_col="AÑO", status_col="ESTATUS_OBRA")
f_cxp  = filter_df(cxp, year_col="AÑO", status_col="ESTATUS_OBRA")
f_egr  = filter_df(egresos, year_col="AÑO", status_col="ESTATUS_OBRA")
# f_egr_costos: egresos filtrados por AÑO_ORIGEN (fecha del costo, no fecha de
# pago). Se usa solo para "Costos Pagados", que compara contra Costos["TOTAL"]
# (también por fecha de origen) — así ambos lados de la razón corresponden a
# la misma ventana de tiempo. f_egr (por FECHA PAGO) se conserva para las
# vistas de flujo de caja, donde SÍ es correcto usar la fecha de pago.
f_egr_costos = filter_df(egresos, year_col="AÑO_ORIGEN", status_col="ESTATUS_OBRA")
f_contratos = filter_df(contratos, year_col=None, status_col="ESTATUS_OBRA")

# -- Cálculos base KPI ----------------------------------------------------------
# Documentación de la lógica central:
# 1) Facturación operativa: excluye registros 'EN PROCESO' y toma solo GRUPO = 'INGRESOS OPERATIVOS'.
# 2) Ventas netas: suma del campo IMPORTE de facturas operativas.
# 3) Costo de producción: suma TOTAL para CLASIFICACIÓN = 'COSTO PRODUCCION'.
# 4) Utilidad bruta: ventas netas - costo de producción.
# 5) EBITDA: utilidad bruta + otros ingresos - gastos no operativos.
# 6) EBIT: EBITDA - gastos administrativos - gastos financieros.
# 7) Utilidad neta: EBIT - impuestos.
# 8) Posición neta de tesorería: CxC + F.G. recuperable * ajuste (IVA/ajuste contractual) - (CxP + anticipos por amortizar * ajuste).

f_fact_operativos = f_fact[
    (f_fact["No. FACTURA"] != "EN PROCESO") &
    (f_fact["GRUPO"] == "INGRESOS OPERATIVOS")
]
f_fact_no_operativos = f_fact[
    (f_fact["No. FACTURA"] != "EN PROCESO") &
    (f_fact["GRUPO"] == "INGRESOS NO OPERATIVOS")
]

# NOTA: total_fact excluye "EN PROCESO" (facturas aún no emitidas formalmente).
# Antes se incluían aquí y sí se excluían en f_fact_operativos, inflando el KPI
# de "Facturación" del Resumen Ejecutivo por el monto de facturas pendientes.
total_fact     = f_fact[f_fact["No. FACTURA"] != "EN PROCESO"]["IMPORTE"].sum()
total_ing      = f_ing["INGRESOS"].sum()
total_cos_prod = f_cos[f_cos["CLASIFICACIÓN"] == "COSTO PRODUCCION"]["TOTAL"].sum()
total_cos_no_operativos = f_cos[f_cos["CLASIFICACIÓN"] == "GASTOS NO OPERATIVOS"]["TOTAL"].sum()
total_cos_admin = f_cos[f_cos["CLASIFICACIÓN"] == "GASTOS ADMINISTRATIVOS"]["TOTAL"].sum()
total_cos_financieros = f_cos[f_cos["CLASIFICACIÓN"] == "GASTOS FINANCIEROS"]["TOTAL"].sum()
total_cos_nop  = total_cos_no_operativos + total_cos_financieros
total_cos_admin_fin = total_cos_admin + total_cos_financieros
total_cos_imp  = f_cos[f_cos["CLASIFICACIÓN"] == "IMPUESTOS"]["TOTAL"].sum()
total_cos      = total_cos_prod + total_cos_admin + total_cos_nop + total_cos_imp
total_cxc      = f_cxc["CxC"].sum()
total_cxp      = f_cxp["TOTAL"].sum()
# "Costos Pagados" debe comparar contra pagos de la MISMA ventana de costo
# incurrido (f_egr_costos, filtrado por AÑO_ORIGEN), no por fecha de pago —
# ver nota en filter_df. f_egr (por FECHA PAGO) sigue existiendo para las
# vistas de flujo de caja mensual, donde la fecha de pago sí es lo correcto.
total_egr      = f_egr["IMPORTE PAGADO"].sum()
total_egr_costos_ventana = f_egr_costos["IMPORTE PAGADO"].sum()

# --- P&L: bases netas de IVA -------------------------------------------------
# ventas_netas / total_cos_prod (etc.) siguen en base CON IVA porque así están
# construidos CxC, CxP y Egresos en el dataset (ver DSO y Costos Pagados) — se
# conservan para esos cálculos de caja/cobranza. Pero usarlos para Utilidad
# Bruta/EBITDA/EBIT mezcla IMPORTE (con IVA, sin retenciones) contra TOTAL
# (con IVA, NETO de retenciones), lo que sobreestimaba la utilidad bruta en
# ~19% sobre los 10 años de este dataset. Para el Estado de Resultados se usa
# SUBTOTAL (neto de IVA) de ambos lados, que es la base contable correcta.
ventas_netas   = f_fact_operativos["IMPORTE"].sum()
ventas_sin_iva = f_fact_operativos["SUBTOTAL"].sum()
otros_ingresos = f_fact_no_operativos["IMPORTE"].sum()
otros_ingresos_sin_iva = f_fact_no_operativos["SUBTOTAL"].sum()
fg_x_recuperar = f_fact_operativos["F.G."].sum()
anticipos      = f_fact[f_fact["CONTRATO"] == "ANTICIPO"]["SUBTOTAL"].sum()
amortizado     = f_fact["AMORTIZACION"].abs().sum()
por_amortizar  = anticipos - amortizado

total_cos_prod_sin_iva  = f_cos[f_cos["CLASIFICACIÓN"] == "COSTO PRODUCCION"]["SUBTOTAL"].sum()
total_cos_nop_sin_iva   = f_cos[f_cos["CLASIFICACIÓN"].isin(["GASTOS NO OPERATIVOS", "GASTOS FINANCIEROS"])]["SUBTOTAL"].sum()
total_cos_admin_sin_iva = f_cos[f_cos["CLASIFICACIÓN"] == "GASTOS ADMINISTRATIVOS"]["SUBTOTAL"].sum()
total_cos_fin_sin_iva   = f_cos[f_cos["CLASIFICACIÓN"] == "GASTOS FINANCIEROS"]["SUBTOTAL"].sum()
total_cos_admin_fin_sin_iva = total_cos_admin_sin_iva + total_cos_fin_sin_iva
total_cos_imp_sin_iva   = f_cos[f_cos["CLASIFICACIÓN"] == "IMPUESTOS"]["SUBTOTAL"].sum()

ut_bruta       = ventas_sin_iva - total_cos_prod_sin_iva
ebitda         = ut_bruta + otros_ingresos_sin_iva - (total_cos_nop_sin_iva - total_cos_fin_sin_iva)
ebit           = ebitda - total_cos_admin_fin_sin_iva
beneficio_neto = ebit - total_cos_imp_sin_iva
margen_neto    = beneficio_neto / ventas_sin_iva * 100 if ventas_sin_iva else 0
margen_bruto   = ut_bruta / ventas_sin_iva * 100 if ventas_sin_iva else 0

# Aviso: estas dos clasificaciones no existen en el dataset actual de Costos
# (solo hay COSTO PRODUCCION, GASTOS NO OPERATIVOS, GASTOS ADMINISTRATIVOS).
# Mientras eso no cambie, total_cos_imp_sin_iva y total_cos_fin_sin_iva son
# siempre 0 y "Utilidad Neta" es en realidad EBIT (no se están deduciendo
# impuestos ni gastos financieros reales en ningún punto del modelo).
FALTAN_CLASIFICACIONES_FISCALES = not costos["CLASIFICACIÓN"].isin(
    ["IMPUESTOS", "GASTOS FINANCIEROS"]
).any()

# Ajuste contractual de 1.16: se mantiene para mantener la lógica de reporteo previa del modelo.
activo_contractual_neto = total_cxc + fg_x_recuperar * 1.16
pasivo_operativo_acumulado = total_cxp + por_amortizar * 1.16
posicion_neta_pbix = activo_contractual_neto - pasivo_operativo_acumulado

# -- DSO ----------------------------------------------------------------------
# DSO = CxC / (Ventas del periodo / días del periodo). La fórmula anterior
# dividía ventas_netas (suma de TODO el periodo filtrado, p. ej. 3 o 10 años)
# entre 30 como si fuera venta de un solo mes, lo que hundía el DSO a
# fracciones de día en vez de las ~26 días reales del ciclo de cobro. Aquí se
# usan los días reales que abarca la facturación operativa filtrada, así el
# resultado es correcto sin importar cuántos años se seleccionen en el filtro.
if not f_fact_operativos.empty and ventas_netas:
    fechas_periodo = pd.to_datetime(f_fact_operativos["FECHA FACT."], errors="coerce").dropna()
    periodo_dias = (fechas_periodo.max() - fechas_periodo.min()).days + 1 if not fechas_periodo.empty else 0
else:
    periodo_dias = 0
dso_days = (total_cxc / ventas_netas) * periodo_dias if ventas_netas and periodo_dias else 0

# -- Validaciones de coherencia ----
cxc_esperado = total_fact - total_ing
diferencia_cxc_pct = abs(total_cxc - cxc_esperado) / max(cxc_esperado, 0.01) * 100 if cxc_esperado != 0 else 0


# ==============================================================================
# ENCABEZADO DE PÁGINA (compartido)
# ==============================================================================

page_meta = {
    "🎯 Executive Summary":    ("Dashboard Ejecutivo", "Visión de negocio integral · Toma de decisiones estratégicas"),
    "💰 P&L Ejecutivo":        ("Estado de Resultados", "Vista consolidada · P&L completo"),
    "📊 Backlog & Contratos":  ("Backlog & Contratos", "Estado de cuenta · Saldo por ejecutar"),
    "📈 Facturación & Cobro":  ("Facturación & Cobranza", "Ciclo ingreso → cobro → CxC"),
    "⚙️ Costos & Margen":      ("Costos & Rentabilidad", "Producción real vs costo real"),
    "💳 Flujo de Caja":        ("Flujo de Caja", "CxC vs CxP · Posición neta"),
    "🌊 Sankey · Costos":      ("Flujo de Costos - Sankey", "Portafolio → Clasificación → Insumo"),
}

title, subtitle = page_meta[page]
st.markdown(f"""
<div class="brand-header">
  <div class="brand-mark">cAc</div>
  <div>
    <div class="brand-title">{title}</div>
    <div class="brand-subtitle">{subtitle}</div>
  </div>
</div>
""", unsafe_allow_html=True)


# ==============================================================================
# PÁGINA 0 - RESUMEN EJECUTIVO (INICIO)
# ==============================================================================

if page == "🎯 Executive Summary":

    # -- Header Narrativa - Calcular variables primero
    color_salud = POSITIVE if margen_neto > 15 else ACCENT
    texto_salud = "FUERTE ✅" if margen_neto > 15 else ("MODERADA ⚠️" if margen_neto > 5 else "CRÍTICA 🔴")
    
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, rgba(74,127,181,0.12) 0%, rgba(15,22,32,0.15) 100%); 
                padding: 24px; border-radius: 12px; margin-bottom: 24px;">
        <p style="font-size: 16px; line-height: 1.8; color: {OFF_WHITE}; margin: 0;">
        <strong>📊 Resumen del Período:</strong> Durante este período, hemos ejecutado <strong>${total_fact/1e6:.0f}M</strong> en facturación 
        con una utilidad neta de <strong>${beneficio_neto/1e6:.0f}M</strong> ({margen_neto:.1f}% margen). 
        La salud financiera es <strong style="color: {color_salud};">
        {texto_salud}</strong>.
        </p>
    </div>
    """, unsafe_allow_html=True)

    # -- FILA KPI PRINCIPAL --------------------------------------------------------
    col1, col2, col3, col4 = st.columns(4)
    
    # Calcular posicion_neta primero
    posicion_neta_kpi = posicion_neta_pbix
    color_posicion = POSITIVE if posicion_neta_kpi > 0 else ALERT
    
    with col1:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">💰 Ingresos Netos</div>
            <div class="kpi-value">${total_ing/1e6:.1f}M</div>
            <div class="kpi-delta">De ${total_fact/1e6:.1f}M facturado</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        color_margen = POSITIVE if margen_neto > 15 else (ACCENT if margen_neto > 5 else ALERT)
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">📊 Margen Neto</div>
            <div class="kpi-value" style="color: {color_margen};">{margen_neto:.1f}%</div>
            <div class="kpi-delta">Utilidad: ${beneficio_neto/1e6:.1f}M</div>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">⚖️ Posición Tesorería (ajustada)</div>
            <div class="kpi-value" style="color: {color_posicion};">${posicion_neta_kpi/1e6:.1f}M</div>
            <div class="kpi-delta">{'Favorable' if posicion_neta_kpi > 0 else 'A mejorar'} · incl. FG y anticipos</div>
        </div>
    """, unsafe_allow_html=True)

    with col4:
        # Se compara contra total_egr_costos_ventana (egresos filtrados por
        # AÑO_ORIGEN, misma base temporal que total_cos) en vez de total_egr
        # (por fecha de pago). Con la fecha de pago, este ratio podía superar
        # 100% al liquidarse en el periodo deuda incurrida en años anteriores.
        ratio_pago = total_egr_costos_ventana / total_cos * 100 if total_cos > 0 else 0
        color_pago = POSITIVE if ratio_pago > 80 else ACCENT
        st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">💳 Costos Pagados</div>
            <div class="kpi-value" style="color: {color_pago};">{ratio_pago:.0f}%</div>
            <div class="kpi-delta">Saldo: ${max(0, total_cxp)/1e6:.1f}M</div>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # -- TRES HISTORIAS PRINCIPALES -----------------------------------------------
    st.markdown("### 📖 La Historia de Tu Negocio")
    
    col_story1, col_story2, col_story3 = st.columns(3)

    # HISTORIA 1: RENTABILIDAD
    with col_story1:
        rentabilidad_status = "✅ SALUDABLE" if margen_neto > 15 else ("⚠️ MODERADA" if margen_neto > 5 else "🔴 EN RIESGO")
        st.markdown(f"""
        <div style="background: rgba(24,34,48,0.5); border: 1px solid rgba(74,127,181,0.18); 
                    border-radius: 8px; padding: 16px;">
            <h4 style="margin-top: 0; color: {ACCENT};">🎯 Rentabilidad</h4>
            <p><strong>{margen_neto:.1f}%</strong> margen neto</p>
            <p style="color: {MUTED}; font-size: 12px;">
            Por cada peso de venta, ganas <strong>${margen_neto/100:.2f}</strong>
            </p>
            <p style="color: {POSITIVE if margen_neto > 15 else ACCENT}; font-weight: bold;">
            Estado: {rentabilidad_status}
            </p>
        </div>
        """, unsafe_allow_html=True)

    # HISTORIA 2: OPERACIONES
    with col_story2:
        cxc_pct = total_cxc / ventas_netas * 100 if ventas_netas else 0
        dso_status = "✅ EXCELENTE" if dso_days < 30 else ("⚠️ NORMAL" if dso_days < 60 else "🔴 RETRASADO")
        st.markdown(f"""
        <div style="background: rgba(24,34,48,0.5); border: 1px solid rgba(74,127,181,0.18); 
                    border-radius: 8px; padding: 16px;">
            <h4 style="margin-top: 0; color: {ACCENT};">🚀 Operaciones</h4>
            <p><strong>{dso_days:.0f}</strong> días promedio de cobro</p>
            <p style="color: {MUTED}; font-size: 12px;">
            ${total_cxc/1e6:.1f}M por cobrar ({cxc_pct:.0f}% de facturación)
            </p>
            <p style="color: {POSITIVE if dso_days < 30 else ACCENT}; font-weight: bold;">
            Estado: {dso_status}
            </p>
        </div>
        """, unsafe_allow_html=True)

    # HISTORIA 3: LIQUIDEZ
    with col_story3:
        posicion_neta = total_cxc - total_cxp
        liquidez_status = "✅ FUERTE" if posicion_neta > 0 else "🔴 DÉBIL"
        liquidez_color = POSITIVE if posicion_neta > 0 else ALERT
        st.markdown(f"""
        <div style="background: rgba(24,34,48,0.5); border: 1px solid rgba(74,127,181,0.18); 
                    border-radius: 8px; padding: 16px;">
            <h4 style="margin-top: 0; color: {ACCENT};">💰 Liquidez (CxC − CxP)</h4>
            <p><strong style="color: {liquidez_color};">${abs(posicion_neta)/1e6:.1f}M</strong></p>
            <p style="color: {MUTED}; font-size: 12px;">
            {'Tienes saldo a favor' if posicion_neta > 0 else 'Necesitas cubrir deuda'}
            </p>
            <p style="color: {liquidez_color}; font-weight: bold;">
            Estado: {liquidez_status}
            </p>
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # -- MÉTRICAS COMPARATIVAS ---------------------------------------------------
    st.markdown("### 📊 Comparativas Clave")
    
    col_comp1, col_comp2 = st.columns(2)

    with col_comp1:
        # Ingresos vs Costos
        fig_comp1 = go.Figure()
        fig_comp1.add_trace(go.Bar(
            x=["Facturación", "Costo Producción", "Gastos Admin", "Otros"],
            y=[total_fact/1e6, total_cos_prod/1e6, total_cos_admin/1e6, 
               (total_cos_nop + total_cos_imp)/1e6],
            marker_color=[ACCENT, ALERT, NEUTRAL_DIM, NEUTRAL],
            text=[f"${v:.0f}M" for v in [total_fact/1e6, total_cos_prod/1e6, 
                  total_cos_admin/1e6, (total_cos_nop + total_cos_imp)/1e6]],
            textposition="outside",
            textfont=dict(size=10, color=OFF_WHITE),
        ))
        apply_layout(fig_comp1, "Ingresos vs Estructura de Costos", 300)
        st.plotly_chart(fig_comp1, use_container_width=True)

    with col_comp2:
        # CxC vs CxP
        fig_comp2 = go.Figure()
        fig_comp2.add_trace(go.Bar(
            x=["Cuentas por Cobrar", "Cuentas por Pagar"],
            y=[total_cxc/1e6, total_cxp/1e6],
            marker_color=[POSITIVE, ALERT],
            text=[f"${total_cxc/1e6:.1f}M", f"${total_cxp/1e6:.1f}M"],
            textposition="outside",
            textfont=dict(size=12, color=OFF_WHITE, family="Bebas Neue"),
        ))
        fig_comp2.update_layout(
            title=dict(text="Posición de Tesorería", font=dict(size=13), x=0),
            height=300, **LAYOUT_BASE
        )
        st.plotly_chart(fig_comp2, use_container_width=True)

    st.divider()

    # -- ALERTAS ESTRATÉGICAS ---------------------------------------------------
    st.markdown("### 🚨 Puntos de Atención Estratégicos")
    
    alertas_count = 0
    
    # Alerta 1: Margen bajo
    if margen_neto < 5:
        st.markdown(f"""
        <div class="alert-box">
        <strong>🔴 MARGEN CRÍTICO:</strong> El margen neto de {margen_neto:.1f}% está muy por debajo 
        del benchmark (15%+). Revisar estructura de costos urgentemente.
        </div>
        """, unsafe_allow_html=True)
        alertas_count += 1
    
    # Alerta 2: DSO alto
    if dso_days > 60:
        st.markdown(f"""
        <div class="alert-box">
        <strong>🔴 COBRANZA LENTA:</strong> {dso_days:.0f} días promedio es muy alto. 
        Recomendamos acelerar procesos de cobro.
        </div>
        """, unsafe_allow_html=True)
        alertas_count += 1
    
    # Alerta 3: Posición negativa
    if posicion_neta < 0:
        st.markdown(f"""
        <div class="alert-box">
        <strong>🔴 LIQUIDEZ NEGATIVA:</strong> Tienes ${abs(posicion_neta)/1e6:.1f}M más en deudas 
        que en cuentas por cobrar. Priorizar pagos y acelerar cobros.
        </div>
        """, unsafe_allow_html=True)
        alertas_count += 1
    
    # Alerta 4: Costos sin pagar (misma ventana temporal costo↔pago, ver nota arriba)
    pct_costos_pagados = total_egr_costos_ventana / total_cos * 100 if total_cos > 0 else 0
    if pct_costos_pagados < 70:
        st.markdown(f"""
        <div class="alert-box">
        <strong>🟡 DEUDA ACUMULADA:</strong> Solo {pct_costos_pagados:.0f}% de costos pagados. 
        ${total_cxp/1e6:.1f}M pendiente con proveedores.
        </div>
        """, unsafe_allow_html=True)
        alertas_count += 1
    
    if alertas_count == 0:
        st.markdown("""
        <div class="tip-box">
        <strong>✅ SIN ALERTAS CRÍTICAS:</strong> Tu negocio está operando dentro de parámetros normales.
        Continúa monitoreando los indicadores.
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    # -- RECOMENDACIONES POR ROL ------------------------------------------------
    st.markdown("### 💡 Recomendaciones Ejecutivas")
    
    col_rec1, col_rec2 = st.columns(2)

    with col_rec1:
        st.markdown(f"""
        **Para el CEO:**
        - Margen actual: **{margen_neto:.1f}%** (Objetivo: 15%)
        - Rentabilidad y márgenes deben ser la prioridad
        - {'✅ En buen camino' if margen_neto > 10 else '⚠️ Revisar urgentemente'}
        
        **Acciones sugeridas:**
        1. Analizar portafolios con margen < 10%
        2. Optimizar estructura de costos administrativos
        3. Mejorar eficiencia operativa
        """)

    with col_rec2:
        st.markdown(f"""
        **Para el Gerente Operativo:**
        - Cobranza: **{dso_days:.0f} días** (Objetivo: < 30 días)
        - Liquidez: **${posicion_neta/1e6:.1f}M** {'✅' if posicion_neta > 0 else '⚠️'}
        
        **Acciones sugeridas:**
        1. {'Acelerar cobranza a clientes' if dso_days > 45 else 'Mantener ritmo de cobranza'}
        2. {'Renegociar plazos con proveedores' if posicion_neta < 0 else 'Optimizar ciclo de pago'}
        3. Revisar CxC > 90 días
        """)


# ==============================================================================
# PÁGINA 1 - P&L EJECUTIVO
# ==============================================================================

elif page == "💰 P&L Ejecutivo":

    # -- Fila de KPI ------------------------------------------------------------
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("💰 Ingresos (Ventas Netas)", f"${ventas_sin_iva/1e6:.1f}M")
    c2.metric("📊 Utilidad Bruta",          f"${ut_bruta/1e6:.1f}M",
          delta=f"{margen_bruto:.1f}% margen bruto")
    c3.metric("✅ Utilidad Neta",            f"${beneficio_neto/1e6:.1f}M",
          delta=f"{margen_neto:.1f}% margen neto")
    c4.metric("🔧 Costo de Producción",     f"${total_cos_prod_sin_iva/1e6:.1f}M",
          delta=f"{total_cos_prod_sin_iva/ventas_sin_iva*100:.1f}% de ingresos" if ventas_sin_iva else None)
    c5.metric("📋 Gastos Operativos",       f"${total_cos_admin_sin_iva/1e6:.1f}M")
    
    # -- Narrativa de rentabilidad ------------------------------------------------
    st.markdown(f"""
    <div class="tip-box">
    <strong>📖 Lectura rápida:</strong> De cada $100 de ingresos, <strong>${margen_neto:.0f}</strong> 
    quedan como utilidad neta después de todos los costos. 
    El costo de producción consume <strong>${total_cos_prod_sin_iva/ventas_sin_iva*100:.0f}%</strong> de ingresos 
    y los gastos administrativos <strong>${total_cos_admin_sin_iva/ventas_sin_iva*100:.0f}%</strong>.
    </div>
    """, unsafe_allow_html=True)

    if FALTAN_CLASIFICACIONES_FISCALES:
        st.markdown("""
        <div class="alert-box">
        ⚠️ <strong>Nota:</strong> El dataset de Costos no tiene registros clasificados como
        "IMPUESTOS" ni "GASTOS FINANCIEROS". Mientras eso no se capture, esas líneas del
        cálculo son $0 y <strong>"Utilidad Neta" equivale a EBIT</strong> (no hay impuestos
        ni gastos financieros reales deducidos en ningún punto del modelo).
        </div>
        """, unsafe_allow_html=True)

    st.divider()

    col1, col2 = st.columns([3, 2])

    with col1:
        # Waterfall P&L — todo en base neta de IVA (SUBTOTAL), consistente
        # con ut_bruta/ebitda/ebit/beneficio_neto ya corregidos arriba.
        wf_labels = ["Ingresos", "Costo Prod.", "Ut. Bruta", "Otros Ingresos",
                     "EBITDA", "Gastos Admin + Fin", "Impuestos", "Ut. Neta"]
        wf_values = [ventas_sin_iva, -total_cos_prod_sin_iva, 0, otros_ingresos_sin_iva,
                     0, -total_cos_admin_fin_sin_iva, -total_cos_imp_sin_iva, 0]
        wf_measure = ["absolute", "relative", "total", "relative",
                      "total", "relative", "relative", "total"]

        fig_wf = go.Figure(go.Waterfall(
            name="P&L", measure=wf_measure,
            x=wf_labels, y=wf_values,
            connector=dict(line=dict(color="rgba(255,255,255,0.15)")),
            increasing=dict(marker_color=POSITIVE),
            decreasing=dict(marker_color=ALERT),
            totals=dict(marker_color=ACCENT),
            text=[f"${v/1e6:.1f}M" for v in
                  [ventas_sin_iva, total_cos_prod_sin_iva, ut_bruta, otros_ingresos_sin_iva,
                   ebitda, total_cos_admin_fin_sin_iva, total_cos_imp_sin_iva, beneficio_neto]],
            textposition="outside",
            textfont=dict(color=OFF_WHITE, size=11),
        ))
        apply_layout(fig_wf, "Cascada P&L - Del ingreso a la utilidad neta", 360)
        st.plotly_chart(fig_wf, use_container_width=True)

    with col2:
        # Donut composición de costos
        clas_totals = (f_cos.groupby("CLASIFICACIÓN")["TOTAL"]
                       .sum().sort_values(ascending=False).reset_index())
        fig_donut = go.Figure(go.Pie(
            labels=clas_totals["CLASIFICACIÓN"],
            values=clas_totals["TOTAL"],
            hole=0.55,
            marker_colors=PALETTE,
            textinfo="label+percent",
            textfont_size=10,
        ))
        fig_donut.add_annotation(
            text=f"${total_cos/1e6:.0f}M<br><span style='font-size:11px'>Costo total</span>",
            x=0.5, y=0.5, showarrow=False,
            font=dict(size=16, color=ACCENT)
        )
        apply_layout(fig_donut, "Composición de costos", 360)
        st.plotly_chart(fig_donut, use_container_width=True)

    st.divider()

    # P&L por portafolio
    st.markdown("#### P&L por Portafolio")
    # Antes esta tabla comparaba Ingresos (cobrado, CON IVA, hoja Ingresos)
    # contra Costos.TOTAL (CON IVA, NETO de retenciones) — el mismo problema
    # de bases distintas que en el P&L consolidado, más el hecho de mezclar
    # devengado (costo) con flujo de caja (ingreso cobrado). Se corrige usando
    # la misma base que el resto del Estado de Resultados: Facturación
    # operativa devengada (SUBTOTAL, neto de IVA) contra Costos (SUBTOTAL).
    f_fact_operativos_port = f_fact_operativos.groupby("PORTAFOLIO")["SUBTOTAL"].sum()
    pl_ing   = f_fact_operativos_port
    pl_cprod = f_cos[f_cos["CLASIFICACIÓN"]=="COSTO PRODUCCION"].groupby("PORTAFOLIO")["SUBTOTAL"].sum()
    pl_cadm  = f_cos[f_cos["CLASIFICACIÓN"]=="GASTOS ADMINISTRATIVOS"].groupby("PORTAFOLIO")["SUBTOTAL"].sum()
    pl_cnop  = f_cos[f_cos["CLASIFICACIÓN"].isin(["GASTOS NO OPERATIVOS","GASTOS FINANCIEROS"])].groupby("PORTAFOLIO")["SUBTOTAL"].sum()
    pl_cimp  = f_cos[f_cos["CLASIFICACIÓN"]=="IMPUESTOS"].groupby("PORTAFOLIO")["SUBTOTAL"].sum()

    pl_df = pd.DataFrame({
        "Ingresos":    pl_ing,
        "Costo Prod.": pl_cprod,
        "Gtos. Admin.":pl_cadm,
        "No Operativos":pl_cnop,
        "Impuestos":   pl_cimp,
    }).fillna(0)
    pl_df["Ut. Neta"] = (pl_df["Ingresos"] - pl_df["Costo Prod."]
                         - pl_df["Gtos. Admin."] - pl_df["No Operativos"] - pl_df["Impuestos"])
    pl_df["% Margen"] = (pl_df["Ut. Neta"] / pl_df["Ingresos"].replace(0, float("nan")) * 100).fillna(0)
    pl_df = pl_df[pl_df["Ingresos"] > 0].sort_values("Ut. Neta", ascending=False)

    col_a, col_b = st.columns([2, 3])

    with col_a:
        # Bar chart utilidad neta por portafolio
        top_pl = pl_df.head(12)
        colors = [POSITIVE if v >= 0 else ALERT for v in top_pl["Ut. Neta"]]
        fig_pl = go.Figure(go.Bar(
            x=top_pl["Ut. Neta"] / 1e6,
            y=top_pl.index,
            orientation="h",
            marker_color=colors,
            text=[f"${v/1e6:.1f}M" for v in top_pl["Ut. Neta"]],
            textposition="outside",
            textfont=dict(size=10, color=OFF_WHITE),
        ))
        apply_layout(fig_pl, "Utilidad Neta por Portafolio (Top 12)", 400)
        fig_pl.update_yaxes(tickfont_size=10)
        fig_pl.update_xaxes(title_text="Millones MXN")
        st.plotly_chart(fig_pl, use_container_width=True)

    with col_b:
        # Table
        display_pl = pl_df.copy()
        for col in ["Ingresos","Costo Prod.","Gtos. Admin.","No Operativos","Impuestos","Ut. Neta"]:
            display_pl[col] = display_pl[col].apply(lambda x: f"${x/1e6:.1f}M")
        display_pl["% Margen"] = pl_df["% Margen"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(
            display_pl[["Ingresos","Costo Prod.","Ut. Neta","% Margen"]].head(15),
            use_container_width=True, height=400
        )

    # Alertas
    obras_perdida = pl_df[pl_df["Ut. Neta"] < 0]
    if not obras_perdida.empty:
        st.markdown(f'<div class="alert-box">⚠️ <strong>{len(obras_perdida)} portafolios en pérdida:</strong> '
                    + ", ".join(obras_perdida.index.tolist()) + "</div>", unsafe_allow_html=True)


# ==============================================================================
# PÁGINA 2 - BACKLOG & CONTRATOS
# ==============================================================================

elif page == "📊 Backlog & Contratos":

    # KPI
    total_contratado = f_contratos["IMPORTE S/IVA"].sum()
    total_anticipo   = (f_contratos["IMPORTE S/IVA"] * f_contratos["% ANTICIPO"]).sum()
    total_estimado   = f_fact["ESTIMADO"].sum() if "ESTIMADO" in f_fact.columns else 0
    saldo_doc        = total_contratado - total_estimado
    total_fg         = (f_contratos["IMPORTE S/IVA"] * f_contratos["% F.G."]).sum()
    pct_avance       = total_estimado / total_contratado * 100 if total_contratado else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📜 Total Contratado",      f"${total_contratado/1e6:.1f}M")
    c2.metric("📄 Saldo por Documentar",  f"${saldo_doc/1e6:.1f}M",
              delta=f"{pct_avance:.1f}% estimado")
    c3.metric("💵 Anticipos Total",        f"${total_anticipo/1e6:.1f}M")
    c4.metric("🔒 FG Total Pactado",       f"${total_fg/1e6:.1f}M")

    # -- Narrativa de ejecución ---------------------------------------------------
    st.markdown(f"""
    <div class="tip-box">
    <strong>📖 Estado de ejecución:</strong> Del contrato total de <strong>${total_contratado/1e6:.1f}M</strong>, 
    has estimado <strong>${total_estimado/1e6:.1f}M ({pct_avance:.0f}%)</strong>. 
    Queda <strong>${saldo_doc/1e6:.1f}M</strong> por documentar. 
    Anticipos entregados: <strong>${total_anticipo/1e6:.1f}M</strong>.
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        # Avance por portafolio (stacked 100%)
        if sel_estatus:
            contratos_chart = (f_contratos[f_contratos["ESTATUS_OBRA"].isin(sel_estatus)]
                               if "ESTATUS_OBRA" in f_contratos.columns else f_contratos)
            fact_chart = (f_fact[f_fact["ESTATUS_OBRA"].isin(sel_estatus)]
                          if "ESTATUS_OBRA" in f_fact.columns else f_fact)
        else:
            contratos_chart = f_contratos
            fact_chart = f_fact

        port_contracted = contratos_chart.groupby("PORTAFOLIO")["IMPORTE S/IVA"].sum()
        port_estimado   = fact_chart.groupby("PORTAFOLIO")["ESTIMADO"].sum() if "ESTIMADO" in fact_chart.columns else pd.Series(0, index=port_contracted.index)
        avance_df = pd.DataFrame({
            "Contratado": port_contracted,
            "Estimado":  port_estimado,
        }).fillna(0)
        avance_df = avance_df[avance_df["Contratado"] > 1e6].copy()
        avance_df["Saldo"] = (avance_df["Contratado"] - avance_df["Estimado"]).clip(lower=0)
        avance_df["Pct"]   = avance_df["Estimado"] / avance_df["Contratado"] * 100
        avance_df = avance_df.sort_values("Pct", ascending=True).head(14)

        fig_avance = go.Figure()
        fig_avance.add_trace(go.Bar(
            name="Estimado", y=avance_df.index,
            x=avance_df["Pct"].clip(upper=100),
            orientation="h", marker_color=POSITIVE,
            text=[f"{p:.0f}%" for p in avance_df["Pct"]],
            textposition="inside", textfont=dict(size=9, color=WHITE),
        ))
        saldo_pct = (100 - avance_df["Pct"]).clip(lower=0)
        fig_avance.add_trace(go.Bar(
            name="Saldo", y=avance_df.index, x=saldo_pct,
            orientation="h", marker_color="rgba(255,255,255,0.12)",
        ))
        apply_layout(fig_avance, "% Avance de Estimación por Portafolio", 420)
        fig_avance.update_layout(barmode="stack", xaxis=dict(range=[0, 110]))
        fig_avance.update_yaxes(tickfont_size=9)
        st.plotly_chart(fig_avance, use_container_width=True)

    with col2:
        # Composición contractual por tipo
        tipo_df = f_contratos.groupby("TIPO")["IMPORTE C/IVA"].sum().sort_values(ascending=False)
        fig_tipo = go.Figure(go.Bar(
            x=tipo_df.index, y=tipo_df.values / 1e6,
            marker_color=PALETTE,
            text=[f"${v/1e6:.0f}M" for v in tipo_df.values],
            textposition="outside", textfont=dict(size=10, color=OFF_WHITE),
        ))
        apply_layout(fig_tipo, "Valor por Tipo de Contrato", 260)
        fig_tipo.update_yaxes(title_text="Millones MXN")
        st.plotly_chart(fig_tipo, use_container_width=True)

        # Anticipos vs Amortizado
        amort_total = f_fact["AMORTIZACION"].abs().sum()
        fig_ant = go.Figure(go.Bar(
            x=["Anticipos entregados", "Amortizado", "Por amortizar"],
            y=[total_anticipo/1e6, amort_total/1e6,
               max(0, total_anticipo - amort_total)/1e6],
            marker_color=[NEUTRAL_DIM, POSITIVE, ACCENT],
            text=[f"${v:.1f}M" for v in [
                total_anticipo/1e6, amort_total/1e6,
                max(0, total_anticipo - amort_total)/1e6]],
            textposition="outside", textfont=dict(size=10, color=OFF_WHITE),
        ))
        apply_layout(fig_ant, "Control de Anticipos", 220)
        st.plotly_chart(fig_ant, use_container_width=True)

    st.divider()

    # Tabla de detalle
    st.markdown("#### Estado de Cuenta por Portafolio")
    detail_df = pd.DataFrame({
        "Contratado":  port_contracted,
        "Estimado":    port_estimado,
    }).fillna(0)
    detail_df = detail_df[detail_df["Contratado"] > 0].copy()
    detail_df["Saldo x Doc."] = detail_df["Contratado"] - detail_df["Estimado"]
    detail_df["% Avance"]     = (detail_df["Estimado"] / detail_df["Contratado"] * 100).round(1)
    detail_df = detail_df.sort_values("Contratado", ascending=False)
    for col in ["Contratado","Estimado","Saldo x Doc."]:
        detail_df[col] = detail_df[col].apply(lambda x: f"${x/1e6:.1f}M")
    detail_df["% Avance"] = detail_df["% Avance"].apply(lambda x: f"{x:.1f}%")
    st.dataframe(detail_df, use_container_width=True, height=320)


# ==============================================================================
# PÁGINA 3 - FACTURACIÓN & COBRO
# ==============================================================================

elif page == "📈 Facturación & Cobro":

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🧾 Facturación",        f"${total_fact/1e6:.1f}M")
    c2.metric("💵 Ingresos Cobrados",  f"${total_ing/1e6:.1f}M")
    c3.metric("🔴 CxC Pendiente",      f"${total_cxc/1e6:.1f}M",
              delta=f"{total_cxc/total_fact*100:.1f}% de la facturación" if total_fact else None,
              delta_color="inverse")
    c4.metric("📅 DSO (días cobro)",   f"{dso_days:.0f} días" if dso_days > 0 else "N/D")

    # -- Narrativa de cobranza ----------------------------------------------------
    tasa_cobranza = (total_ing / total_fact * 100) if total_fact > 0 else 0
    st.markdown(f"""
    <div class="tip-box">
    <strong>📖 Ciclo de cobranza:</strong> Hemos facturado <strong>${total_fact/1e6:.1f}M</strong> 
    y cobrado <strong>${total_ing/1e6:.1f}M</strong> ({tasa_cobranza:.0f}% de lo facturado).
    Hay <strong>${total_cxc/1e6:.1f}M</strong> pendientes con un DSO (tiempo promedio de cobro) 
    de <strong>{dso_days:.0f} días</strong>.
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    col1, col2 = st.columns([3, 2])

    with col1:
        # Fact vs Ing mensual
        mon_f = (f_fact.groupby("YEARMON")["IMPORTE"].sum() / 1e6).reset_index()
        mon_i = (f_ing.groupby("YEARMON")["INGRESOS"].sum() / 1e6).reset_index()
        mon_f.columns = ["Mes", "Facturación"]
        mon_i.columns = ["Mes", "Ingresos"]
        mon_merged = pd.merge(mon_f, mon_i, on="Mes", how="outer").fillna(0)
        mon_merged = mon_merged.sort_values("Mes").tail(18)

        fig_mon = make_subplots(specs=[[{"secondary_y": True}]])
        fig_mon.add_trace(go.Bar(
            name="Facturación", x=mon_merged["Mes"], y=mon_merged["Facturación"],
            marker_color=ACCENT, opacity=0.85,
        ), secondary_y=False)
        fig_mon.add_trace(go.Scatter(
            name="Ingresos cobrados", x=mon_merged["Mes"], y=mon_merged["Ingresos"],
            mode="lines+markers", line=dict(color=POSITIVE, width=2.5),
            marker=dict(size=6, color=POSITIVE),
        ), secondary_y=True)
        apply_layout(fig_mon, "Facturación vs Ingresos cobrados - mensual", 340)
        fig_mon.update_yaxes(title_text="Facturación (M)", secondary_y=False,
                              gridcolor="rgba(255,255,255,0.05)")
        fig_mon.update_yaxes(title_text="Cobrado (M)", secondary_y=True,
                              gridcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig_mon, use_container_width=True)

    with col2:
        # CxC por portafolio
        cxc_port = f_cxc.groupby("PORTAFOLIO")["CxC"].sum().sort_values(ascending=False).head(10)
        fig_cxc = go.Figure(go.Bar(
            x=cxc_port.values / 1e6,
            y=cxc_port.index,
            orientation="h",
            marker_color=[ALERT if v == cxc_port.max() else ACCENT for v in cxc_port.values],
            text=[f"${v/1e6:.1f}M" for v in cxc_port.values],
            textposition="outside",
            textfont=dict(size=9, color=OFF_WHITE),
        ))
        apply_layout(fig_cxc, "CxC por Portafolio - Prioridad cobro", 340)
        fig_cxc.update_yaxes(tickfont_size=9)
        st.plotly_chart(fig_cxc, use_container_width=True)

    st.divider()

    # Comparación año sobre año (si se seleccionan varios años)
    if len(sel_years) > 1:
        st.markdown("#### Comparativa Año vs Año - Facturación mensual")
        fig_yoy = go.Figure()
        colors_yoy = [ACCENT, POSITIVE, NEUTRAL_DIM, ALERT, CHART_2]
        for i, yr in enumerate(sorted(sel_years)):
            yr_data = facturacion[facturacion["AÑO"] == yr].groupby("MES")["IMPORTE"].sum() / 1e6
            yr_data = yr_data[yr_data.index.notna()]
            meses = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]
            fig_yoy.add_trace(go.Scatter(
                x=[meses[int(m)-1] for m in yr_data.index if str(m) != "nan"],
                y=yr_data.values,
                name=str(yr),
                mode="lines+markers",
                line=dict(color=colors_yoy[i % len(colors_yoy)], width=2),
                marker=dict(size=5),
            ))
        apply_layout(fig_yoy, "", 280)
        st.plotly_chart(fig_yoy, use_container_width=True)

    # Tabla detallada de CxC
    st.markdown("#### Detalle CxC Pendiente")
    if not f_cxc.empty:
        cxc_detail = f_cxc[f_cxc["CxC"] != 0][
            ["CLIENTE","OBRA","No. FACTURA","FECHA FACT.","CONCEPTO","CxC","PORTAFOLIO"]
        ].copy()
        cxc_detail = cxc_detail.sort_values("CxC", ascending=False)
        cxc_detail["CxC"] = cxc_detail["CxC"].apply(lambda x: f"${x:,.0f}")
        st.dataframe(cxc_detail.head(30), use_container_width=True, height=280)


# ==============================================================================
# PÁGINA 4 - COSTOS & MARGEN
# ==============================================================================

elif page == "⚙️ Costos & Margen":

    # Igual que en el P&L: se usa la base neta de IVA (SUBTOTAL) para que
    # "ganancia" y "% utilidad" sean comparables ingreso-vs-costo sin el
    # sesgo de que Costos["TOTAL"] ya trae retenciones restadas e IMPORTE no.
    ganancia_prod  = ventas_sin_iva - total_cos_prod_sin_iva
    pct_util_prod  = ganancia_prod / ventas_sin_iva * 100 if ventas_sin_iva else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 Producción (Fact.)",   f"${ventas_sin_iva/1e6:.1f}M")
    c2.metric("🔧 Costo Real",           f"${total_cos_prod_sin_iva/1e6:.1f}M")
    c3.metric("📈 Ganancia Producción",  f"${ganancia_prod/1e6:.1f}M",
              delta=f"{pct_util_prod:.1f}% utilidad")
    c4.metric("💰 Utilidad Neta",         f"${beneficio_neto/1e6:.1f}M",
              delta=f"{margen_neto:.1f}%")

    # -- Narrativa de costos ------------------------------------------------------
    costo_pct_ing = total_cos_prod_sin_iva / ventas_sin_iva * 100 if ventas_sin_iva > 0 else 0
    st.markdown(f"""
    <div class="tip-box">
    <strong>📖 Gestión de costos:</strong> El costo de producción representa <strong>{costo_pct_ing:.0f}%</strong> 
    de los ingresos. Por cada peso de venta, inviertes <strong>${costo_pct_ing/100:.2f}</strong> en producción.
    La ganancia directa de producción es <strong>${ganancia_prod/1e6:.1f}M ({pct_util_prod:.1f}%)</strong>.
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        # Costos por clasificacion
        clas_df = f_cos.groupby("CLASIFICACIÓN")["TOTAL"].sum().sort_values(ascending=False).reset_index()
        fig_clas = go.Figure(go.Bar(
            x=clas_df["CLASIFICACIÓN"], y=clas_df["TOTAL"] / 1e6,
            marker_color=PALETTE[:len(clas_df)],
            text=[f"${v/1e6:.0f}M" for v in clas_df["TOTAL"]],
            textposition="outside", textfont=dict(size=10, color=OFF_WHITE),
        ))
        apply_layout(fig_clas, "Costos por Clasificación", 300)
        fig_clas.update_yaxes(title_text="Millones MXN")
        st.plotly_chart(fig_clas, use_container_width=True)

        # Insumos detalle (solo costo de produccion)
        cos_prod_df = f_cos[f_cos["CLASIFICACIÓN"] == "COSTO PRODUCCION"]
        insumo_df = (cos_prod_df.groupby("INSUMO / SUBCONTRATO")["TOTAL"]
                     .sum().sort_values(ascending=False).head(8).reset_index())
        fig_ins = go.Figure(go.Bar(
            x=insumo_df["INSUMO / SUBCONTRATO"],
            y=insumo_df["TOTAL"] / 1e6,
            marker_color=PALETTE[:len(insumo_df)],
            text=[f"${v/1e6:.0f}M" for v in insumo_df["TOTAL"]],
            textposition="outside", textfont=dict(size=9, color=OFF_WHITE),
        ))
        apply_layout(fig_ins, "Insumos - Costo de Producción", 280)
        st.plotly_chart(fig_ins, use_container_width=True)

    with col2:
        # Acumulado produccion vs costo mensual
        mon_prod = (f_fact.groupby("YEARMON")["IMPORTE"].sum() / 1e6).cumsum().reset_index()
        mon_cost = (f_cos[f_cos["CLASIFICACIÓN"]=="COSTO PRODUCCION"]
                    .groupby("YEARMON")["TOTAL"].sum() / 1e6).cumsum().reset_index()
        mon_prod.columns = ["Mes", "Producción acum."]
        mon_cost.columns = ["Mes", "Costo acum."]
        mon_acc = pd.merge(mon_prod, mon_cost, on="Mes", how="outer").ffill().fillna(0)
        mon_acc = mon_acc.sort_values("Mes").tail(24)

        fig_acc = go.Figure()
        fig_acc.add_trace(go.Scatter(
            x=mon_acc["Mes"], y=mon_acc["Producción acum."],
            name="Acum. Producción", mode="lines",
            line=dict(color=ACCENT, width=2.5, dash="dash"),
        ))
        fig_acc.add_trace(go.Scatter(
            x=mon_acc["Mes"], y=mon_acc["Costo acum."],
            name="Acum. Costo", mode="lines",
            line=dict(color=ALERT, width=2.5, dash="dash"),
        ))
        fig_acc.add_trace(go.Scatter(
            x=list(mon_acc["Mes"]) + list(mon_acc["Mes"])[::-1],
            y=list(mon_acc["Producción acum."]) + list(mon_acc["Costo acum."])[::-1],
            fill="toself", fillcolor="rgba(79,157,140,0.1)",
            line=dict(color="rgba(0,0,0,0)"), name="Brecha (ganancia)",
            showlegend=True,
        ))
        apply_layout(fig_acc, "Producción vs Costo acumulado", 320)
        st.plotly_chart(fig_acc, use_container_width=True)

        # Costos por portafolio (treemap)
        port_cos_df = (f_cos.groupby("PORTAFOLIO")["TOTAL"]
                       .sum().reset_index()
                       .sort_values("TOTAL", ascending=False).head(12))
        fig_tree = px.treemap(
            port_cos_df, path=["PORTAFOLIO"], values="TOTAL",
            color="TOTAL", color_continuous_scale=[[0, BG_DARK], [1, ACCENT]],
        )
        fig_tree.update_traces(textinfo="label+value+percent root",
                               textfont_size=11)
        apply_layout(fig_tree, "Distribución de costo por portafolio (Treemap)", 280)
        fig_tree.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig_tree, use_container_width=True)


# ==============================================================================
# PÁGINA 5 - FLUJO DE CAJA
# ==============================================================================

elif page == "💳 Flujo de Caja":

    posicion_neta = total_cxc - total_cxp

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🟢 CxC (por cobrar)",    f"${total_cxc/1e6:.1f}M")
    c2.metric("🔴 CxP (por pagar)",     f"${total_cxp/1e6:.1f}M")
    c3.metric("⚖️ Posición Neta",       f"${posicion_neta/1e6:.1f}M",
              delta="Favorable" if posicion_neta > 0 else "A mejorar",
              delta_color="normal" if posicion_neta > 0 else "inverse")
    c4.metric("💸 Egresos Pagados",     f"${total_egr/1e6:.1f}M")

    # -- Narrativa de tesorería ---------------------------------------------------
    st.markdown(f"""
    <div class="tip-box">
    <strong>📖 Posición de tesorería:</strong> Tienes <strong>${total_cxc/1e6:.1f}M</strong> por cobrar 
    de clientes y <strong>${total_cxp/1e6:.1f}M</strong> por pagar a proveedores. 
    Tu posición neta es <strong style="color: {'{0}'.format(POSITIVE if posicion_neta > 0 else ALERT)};">${abs(posicion_neta)/1e6:.1f}M</strong> 
    {'(FAVORABLE)' if posicion_neta > 0 else '(DESFAVORABLE)'}.
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        # Barra estilo gauge de CxC vs CxP
        fig_pos = go.Figure()
        fig_pos.add_trace(go.Bar(
            name="CxC (lo que nos deben)",
            x=["Posición de caja"],
            y=[total_cxc / 1e6],
            marker_color=POSITIVE, text=f"${total_cxc/1e6:.1f}M",
            textposition="inside", textfont=dict(color=WHITE, size=13),
        ))
        fig_pos.add_trace(go.Bar(
            name="CxP (lo que debemos)",
            x=["Posición de caja"],
            y=[total_cxp / 1e6],
            marker_color=ALERT, text=f"${total_cxp/1e6:.1f}M",
            textposition="inside", textfont=dict(color=WHITE, size=13),
        ))
        apply_layout(fig_pos, "CxC vs CxP - Posición de tesorería", 280)
        fig_pos.update_layout(barmode="group")
        st.plotly_chart(fig_pos, use_container_width=True)

        # CxP por proveedor
        cxp_prov = (f_cxp.groupby("NOMBRE PROVEEDOR")["TOTAL"]
                    .sum().sort_values(ascending=False).head(10).reset_index())
        fig_cxp_bar = go.Figure(go.Bar(
            x=cxp_prov["TOTAL"] / 1e6,
            y=cxp_prov["NOMBRE PROVEEDOR"],
            orientation="h",
            marker_color=ALERT,
            text=[f"${v/1e6:.1f}M" for v in cxp_prov["TOTAL"]],
            textposition="outside",
            textfont=dict(size=9, color=OFF_WHITE),
        ))
        apply_layout(fig_cxp_bar, "Top Proveedores - CxP Pendiente", 320)
        fig_cxp_bar.update_yaxes(tickfont_size=9)
        st.plotly_chart(fig_cxp_bar, use_container_width=True)

    with col2:
        # Flujo neto mensual (ingresos - egresos)
        egr_mon = (f_egr.groupby("YEARMON")["IMPORTE PAGADO"].sum() / 1e6).reset_index()
        ing_mon = (f_ing.groupby("YEARMON")["INGRESOS"].sum() / 1e6).reset_index()
        egr_mon.columns = ["Mes", "Egresos"]
        ing_mon.columns = ["Mes", "Ingresos"]
        flujo = pd.merge(ing_mon, egr_mon, on="Mes", how="outer").fillna(0)
        flujo = flujo.sort_values("Mes").tail(18)
        flujo["Flujo Neto"] = flujo["Ingresos"] - flujo["Egresos"]
        flujo_colors = [POSITIVE if v >= 0 else ALERT for v in flujo["Flujo Neto"]]

        fig_flujo = go.Figure()
        fig_flujo.add_trace(go.Bar(
            name="Ingresos cobrados", x=flujo["Mes"], y=flujo["Ingresos"],
            marker_color=POSITIVE, opacity=0.7,
        ))
        fig_flujo.add_trace(go.Bar(
            name="Egresos pagados", x=flujo["Mes"], y=-flujo["Egresos"],
            marker_color=ALERT, opacity=0.7,
        ))
        fig_flujo.add_trace(go.Scatter(
            name="Flujo Neto", x=flujo["Mes"], y=flujo["Flujo Neto"],
            mode="lines+markers",
            line=dict(color=ACCENT, width=2.5),
            marker=dict(color=flujo_colors, size=7),
        ))
        apply_layout(fig_flujo, "Flujo neto mensual (Ingresos − Egresos)", 320)
        fig_flujo.update_layout(barmode="relative")
        st.plotly_chart(fig_flujo, use_container_width=True)

        # Diagnóstico
        st.markdown("#### 🎯 Diagnóstico de tesorería")
        diagnosticos = []
        if total_cxc > 0:
            cxc_top = (f_cxc.groupby("PORTAFOLIO")["CxC"].sum()
                       .sort_values(ascending=False))
            if len(cxc_top) > 0:
                top_name = cxc_top.index[0]
                top_pct  = cxc_top.iloc[0] / total_cxc * 100
                if top_pct > 60:
                    diagnosticos.append(
                        f"🔴 <strong>{top_pct:.0f}% de CxC</strong> concentrado en <strong>{top_name}</strong>. Riesgo de concentración alto.")

        if posicion_neta < 0:
            diagnosticos.append(f"🔴 Posición neta <strong>negativa</strong>: debes más de lo que te deben (${abs(posicion_neta)/1e6:.1f}M).")
        else:
            diagnosticos.append(f"🟢 Posición neta <strong>positiva</strong>: te deben más de lo que debes (+${posicion_neta/1e6:.1f}M).")

        pct_costos_pagados = total_egr_costos_ventana / total_cos * 100 if total_cos > 0 else 0
        if pct_costos_pagados < 70:
            diagnosticos.append(f"🟡 Solo el <strong>{pct_costos_pagados:.0f}%</strong> de los costos han sido pagados. Deuda acumulada: ${total_cxp/1e6:.1f}M.")
        else:
            diagnosticos.append(f"🟢 <strong>{pct_costos_pagados:.0f}%</strong> de los costos pagados. Gestión de pagos en línea.")

        for msg in diagnosticos:
            color = "alert-box" if "🔴" in msg else "tip-box"
            st.markdown(f'<div class="{color}">{msg}</div>', unsafe_allow_html=True)


# ==============================================================================
# PÁGINA 6 - SANKEY · COSTOS
# ==============================================================================

elif page == "🌊 Sankey · Costos":

    st.markdown(f"""
    <div class="tip-box">
    🔍 <strong>Cómo interpretar:</strong> El diagrama muestra cómo se distribuyen los <strong>${f_cos['TOTAL'].sum()/1e6:.0f}M</strong> 
    en costos. Parte de cada <strong>Portafolio</strong> (izquierda) hacia la <strong>Clasificación de costo</strong> (centro) 
    y finalmente a los <strong>Insumos</strong> (derecha). El ancho de cada banda es proporcional al monto.
    </div>
    """, unsafe_allow_html=True)

    # Construir datos del Sankey a partir de costos filtrados
    cos_sankey = f_cos.copy()
    cos_sankey = cos_sankey[cos_sankey["TOTAL"] > 0]
    cos_sankey = cos_sankey[cos_sankey["PORTAFOLIO"].notna()]
    cos_sankey = cos_sankey[cos_sankey["CLASIFICACIÓN"].notna()]
    cos_sankey["INSUMO_CLEAN"] = cos_sankey["INSUMO / SUBCONTRATO"].fillna("OTROS")

    # Agregar datos
    agg1 = cos_sankey.groupby(["PORTAFOLIO", "CLASIFICACIÓN"])["TOTAL"].sum().reset_index()
    agg2 = cos_sankey.groupby(["CLASIFICACIÓN", "INSUMO_CLEAN"])["TOTAL"].sum().reset_index()

    # Top N portafolios
    top_ports = (cos_sankey.groupby("PORTAFOLIO")["TOTAL"].sum()
                 .sort_values(ascending=False).head(8).index.tolist())
    agg1 = agg1[agg1["PORTAFOLIO"].isin(top_ports)]

    # Top N insumos
    top_insumos = (cos_sankey.groupby("INSUMO_CLEAN")["TOTAL"].sum()
                   .sort_values(ascending=False).head(8).index.tolist())
    agg2 = agg2[agg2["INSUMO_CLEAN"].isin(top_insumos)]

    # Construir lista de nodos
    ports  = agg1["PORTAFOLIO"].unique().tolist()
    clas   = agg1["CLASIFICACIÓN"].unique().tolist()
    ins    = agg2["INSUMO_CLEAN"].unique().tolist()

    # Eliminar nodos duplicados entre niveles
    all_nodes = ports + [c for c in clas if c not in ports] + \
                [i for i in ins if i not in ports and i not in clas]
    node_idx = {n: i for i, n in enumerate(all_nodes)}

    src, tgt, val, colors_link = [], [], [], []
    color_map = {
        "COSTO PRODUCCION":       "rgba(74,127,181,0.4)",
        "GASTOS NO OPERATIVOS":   "rgba(165,91,91,0.4)",
        "GASTOS ADMINISTRATIVOS": "rgba(94,115,145,0.4)",
        "IMPUESTOS":              "rgba(134,151,172,0.4)",
        "GASTOS FINANCIEROS":     "rgba(58,88,118,0.4)",
    }

    for _, row in agg1.iterrows():
        if row["PORTAFOLIO"] in node_idx and row["CLASIFICACIÓN"] in node_idx:
            src.append(node_idx[row["PORTAFOLIO"]])
            tgt.append(node_idx[row["CLASIFICACIÓN"]])
            val.append(row["TOTAL"])
            colors_link.append(color_map.get(row["CLASIFICACIÓN"], "rgba(74,127,181,0.3)"))

    for _, row in agg2.iterrows():
        if row["CLASIFICACIÓN"] in node_idx and row["INSUMO_CLEAN"] in node_idx:
            src.append(node_idx[row["CLASIFICACIÓN"]])
            tgt.append(node_idx[row["INSUMO_CLEAN"]])
            val.append(row["TOTAL"])
            colors_link.append("rgba(74,127,181,0.2)")

    # Node colors
    node_colors = []
    for n in all_nodes:
        if n in ports:
            node_colors.append(CHART_2)
        elif n in clas:
            node_colors.append(color_map.get(n, ACCENT).replace("0.4", "0.9"))
        else:
            node_colors.append(POSITIVE)

    fig_sankey = go.Figure(go.Sankey(
        arrangement="snap",
        node=dict(
            pad=16, thickness=20,
            line=dict(color="rgba(255,255,255,0.1)", width=0.5),
            label=[f"{n}" for n in all_nodes],
            color=node_colors,
            hovertemplate="<b>%{label}</b><br>Total: $%{value:,.0f}<extra></extra>",
        ),
        link=dict(
            source=src, target=tgt, value=val,
            color=colors_link,
            hovertemplate="$%{value:,.0f}<extra></extra>",
        ),
    ))

    fig_sankey.update_layout(
        title=dict(
            text=f"Flujo de costos - Top portafolios → Clasificación → Insumo  |  Total: ${f_cos['TOTAL'].sum()/1e6:.0f}M",
            font=dict(size=13, color=OFF_WHITE), x=0
        ),
        height=580,
        **LAYOUT_BASE
    )
    st.plotly_chart(fig_sankey, use_container_width=True)

    st.divider()

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Desglose por Clasificación")
        clas_summary = (f_cos.groupby("CLASIFICACIÓN")["TOTAL"].sum()
                        .sort_values(ascending=False).reset_index())
        clas_summary["% del total"] = clas_summary["TOTAL"] / clas_summary["TOTAL"].sum() * 100
        clas_summary["Total (M)"] = clas_summary["TOTAL"].apply(lambda x: f"${x/1e6:.1f}M")
        clas_summary["% del total"] = clas_summary["% del total"].apply(lambda x: f"{x:.1f}%")
        st.dataframe(clas_summary[["CLASIFICACIÓN","Total (M)","% del total"]],
                     use_container_width=True, hide_index=True)

    with col2:
        st.markdown("#### Top Insumos - Costo de Producción")
        ins_summary = (f_cos[f_cos["CLASIFICACIÓN"]=="COSTO PRODUCCION"]
                       .groupby("INSUMO / SUBCONTRATO")["TOTAL"].sum()
                       .sort_values(ascending=False).head(8).reset_index())
        ins_summary["% del costo prod."] = (ins_summary["TOTAL"] /
            f_cos[f_cos["CLASIFICACIÓN"]=="COSTO PRODUCCION"]["TOTAL"].sum() * 100)
        ins_summary["Total (M)"] = ins_summary["TOTAL"].apply(lambda x: f"${x/1e6:.1f}M")
        ins_summary["% del costo prod."] = ins_summary["% del costo prod."].apply(lambda x: f"{x:.1f}%")
        st.dataframe(ins_summary[["INSUMO / SUBCONTRATO","Total (M)","% del costo prod."]],
                     use_container_width=True, hide_index=True)