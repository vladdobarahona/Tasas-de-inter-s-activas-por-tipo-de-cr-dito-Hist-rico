# -*- coding: utf-8 -*-
"""
Aplicación Streamlit - Tasas de interés Activas por tipo de crédito - Histórico (df hasta dos meses antes + df últimos dos meses)
Created on Thu Mar 26 20:33:13 2026

Author: Vladimir Alonso Barahona Palacios

Descripción:
------------
Aplicación interactiva desarrollada en Streamlit para la descarga,
validación y procesamiento de información financiera con fines de supervisión,
correspondiente a la información histórica desde la fecha de inicio de la serie
hasta la fecha publicado por la Superintendencia Financiera de Colombia (SFC).

La aplicación permite:
- Consultar la fecha máxima disponible del dataset.
- Descargar información por corte año mes.
- Validar la cantidad de registros antes de descargar.
- Procesar, limpiar y consolidar información agregada.
- Descargar los datos con según plantilla de CIIU.
- Generar un archivo Excel con información resumida.
- Permitir la descarga directa del archivo consolidado.

Fuente de datos:
----------------
Datos abiertos – Superintendencia Financiera de Colombia:
https://www.superfinanciera.gov.co

Repositorio oficial del dataset:
df1
https://www.datos.gov.co/Econom-a-y-Finanzas/Tasas-de-inter-s-activas-por-tipo-de-cr-dito-Hist-/w9zh-vetq/
df2
https://www.datos.gov.co/Econom-a-y-Finanzas/Tasas-de-inter-s-activas-por-tipo-de-cr-dito-ltimo/qzsc-9esp/
Categoría:
----------
Hacienda y Crédito Público

Notas:
------
- Los datos son consultados en tiempo real desde la API de datos.gov.co (Socrata).
- La aplicación soporta descargas masivas mediante paginación.
- Exporta el reporte final en formato Excel compatible con reportes regulatorios.
"""

import streamlit as st
import pandas as pd
import requests
from datetime import datetime, timedelta
from io import BytesIO

# ==============================
# Estilos personalizados
# ==============================
st.markdown("""
<style>
    /* Fondo de toda la aplicación */
    .stApp {
        background: #ffffff !important;
        font-family: "Segoe UI", "Frutiger", "Helvetica Neue", sans-serif;
        padding-top: 20px;
    }

    /* Título principal */
    .main-title {
        color: rgb(120,154,61);
        font-size: 2.5rem;
        font-weight: 700;
        line-height: 1.25;
        margin-top: 15px;
        margin-bottom: 0px;
    }

    /* Subtítulo */
    .sub-title {
        color: #4a4a4a;
        font-size: 1.1rem;
        margin-top: -5px;
        margin-bottom: 25px;
    }

    /* Fondo general de la página (fuera del contenedor blanco) */
    body {
        background-color: rgb(171,190,76) !important;
    }
</style>
""", unsafe_allow_html=True)


# ==============================
# LOGO + TÍTULO
# ==============================
col1, col2 = st.columns([1, 3])

with col1:
    st.image(
        "https://www.finagro.com.co/sites/default/files/logo-front-finagro.png",
        width=180
    )

with col2:
    st.markdown(
        """
        <h1 class="main-title">
            Tasas de interés activas por tipo de crédito (Histórico) – Consulta, Descarga y Procesamiento
        </h1>
        <div class="sub-title">
            Sistema de apoyo para traer información histórica e incluso la correspondiente a los dos últimos meses
        </div>
        """,
        unsafe_allow_html=True
    )


# ---------------- CONFIG ----------------

FUENTES = {
    "Fuente 1 - Tasas de interés activas por tipo de crédito – Últimos dos meses": "https://www.datos.gov.co/resource/qzsc-9esp.json",
    "Fuente 2 - Tasas de interés activas por tipo de crédito – Histórico": "https://www.datos.gov.co/resource/w9zh-vetq.json"
}

LIMIT = 50000

EXPECTED_SHEET = "CIIU"
EXPECTED_COLUMNS = ["Codigo_CIIU"]

# ---------------- FUNCIONES GENERALES ----------------

def obtener_min_max_fecha(base_url):
    min_q = f"{base_url}?$query=SELECT min(fecha_corte) as min_fecha"
    max_q = f"{base_url}?$query=SELECT max(fecha_corte) as max_fecha"

    min_fecha = requests.get(min_q).json()[0]["min_fecha"][:10]
    max_fecha = requests.get(max_q).json()[0]["max_fecha"][:10]

    return min_fecha, max_fecha

# ---------------- RANGO MES ----------------

def calcular_rango_mes(year, month):
    inicio = datetime(year, month, 1)

    if month == 12:
        siguiente = datetime(year + 1, 1, 1)
    else:
        siguiente = datetime(year, month + 1, 1)

    fin = siguiente - timedelta(days=1)

    return inicio.strftime("%Y-%m-%d"), fin.strftime("%Y-%m-%d")

# ---------------- EXCEL ----------------

def validar_excel(file):
    try:
        xls = pd.ExcelFile(file)

        if EXPECTED_SHEET not in xls.sheet_names:
            return False, "Hoja incorrecta"

        df = pd.read_excel(file, sheet_name=EXPECTED_SHEET)

        if list(df.columns) != EXPECTED_COLUMNS:
            return False, "Columnas incorrectas"

        return True, df

    except Exception as e:
        return False, str(e)

def crear_template():
    df = pd.DataFrame({"Codigo_CIIU": ["0111", "0112"]})
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="CIIU", index=False)

    return buffer.getvalue()

# ---------------- DISTINCT DINÁMICO ----------------

def obtener_valores_distintos(base_url, campo):
    query = f"{base_url}?$query=SELECT DISTINCT {campo} as valor"

    r = requests.get(query)

    if r.status_code != 200:
        st.error("Error consultando valores")
        return []

    data = r.json()
    return [x["valor"] for x in data if x["valor"]]

# ---------------- DESCARGA ----------------

def descargar(base_url, filtro_tipo, valores, fecha_desde, fecha_hasta):

    if filtro_tipo == "CIIU":
        ciiu_str = ",".join(f"'{str(x).zfill(4)}'" for x in valores)
        extra_filtro = f"Codigo_CIIU in ({ciiu_str})"
    else:
        val_str = ",".join(f"'{x}'" for x in valores)
        extra_filtro = f"tipo_de_garant_a in ({val_str})"

    where_clause = (
        f"fecha_corte between '{fecha_desde}T00:00:00' "
        f"and '{fecha_hasta}T23:59:59' AND {extra_filtro}"
    )

    filas = []
    offset = 0

    while True:
        params = {
            "$limit": LIMIT,
            "$offset": offset,
            "$where": where_clause
        }

        r = requests.get(base_url, params=params)

        if r.status_code != 200:
            st.error("Error en descarga")
            break

        data = r.json()

        if not data:
            break

        filas.extend(data)
        offset += LIMIT

    if filas:
        return pd.DataFrame(filas)

    return pd.DataFrame()

# ---------------- UI ----------------

st.title("📊 Descarga Datos Dinámicos")

# -------- Fuente --------
fuente_nombre = st.selectbox("Fuente de datos", list(FUENTES.keys()))
BASE_URL = FUENTES[fuente_nombre]

min_f, max_f = obtener_min_max_fecha(BASE_URL)

st.info(f"Datos disponibles: {min_f} → {max_f}")

# -------- Fecha --------
col1, col2 = st.columns(2)

with col1:
    year = st.number_input("Año", 2000, 2100, 2024)

with col2:
    month = st.selectbox("Mes", list(range(1, 13)))

fecha_desde, fecha_hasta = calcular_rango_mes(year, month)

st.success(f"📅 {fecha_desde} → {fecha_hasta}")

# -------- Tipo de filtro --------
st.subheader("Tipo de filtro")

filtro_tipo = st.radio("Filtrar por:", ["CIIU", "Tipo Garantía"])

# -------- CIIU --------

ciiu_list = []
excel_valido = False

if filtro_tipo == "CIIU":

    st.download_button("Descargar plantilla CIIU", crear_template(), "plantilla.xlsx")

    file = st.file_uploader("Subir Excel", type=["xlsx"])

    if file:
        ok, res = validar_excel(file)

        if ok:
            excel_valido = True
            ciiu_list = res["Codigo_CIIU"].astype(str).tolist()
            st.success("Excel válido")
        else:
            st.error(res)

# -------- VARIABLE DINÁMICA --------

valores_variable = []

if filtro_tipo == "Tipo Garantía":

    if st.button("Cargar valores únicos"):
        valores_variable = obtener_valores_distintos(BASE_URL, "tipo_de_garant_a")
        st.session_state["valores"] = valores_variable

    if "valores" in st.session_state:
        valores_variable = st.multiselect(
            "Seleccione valores",
            st.session_state["valores"]
        )

# -------- EJECUCIÓN --------

if st.button("📥 Descargar datos"):

    # VALIDACIONES
    if filtro_tipo == "CIIU" and not excel_valido:
        st.warning("Debe cargar un Excel válido")
        st.stop()

    if filtro_tipo == "Tipo Garantía" and not valores_variable:
        st.warning("Seleccione al menos un valor")
        st.stop()

    valores = ciiu_list if filtro_tipo == "CIIU" else valores_variable

    with st.spinner("Descargando..."):

        df = descargar(
            BASE_URL,
            filtro_tipo,
            valores,
            fecha_desde,
            fecha_hasta
        )

        if df.empty:
            st.warning("Sin datos")
        else:
            output = BytesIO()
            df.to_excel(output, index=False)
            output.seek(0)

            st.success(f"✅ Registros: {len(df)}")

            st.download_button(
                "⬇ Descargar Excel",
                output,
                file_name="datos.xlsx"
            )
