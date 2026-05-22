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
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import re
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
BASE_URL = "https://www.datos.gov.co/resource/qzsc-9esp.json"
LIMIT = 50000


def min_fecha():
    query = '''
https://www.datos.gov.co/resource/qzsc-9esp.json?$query=SELECT min(fecha_corte) AS min_fecha_corte
    '''
    response = requests.get(query)

    if response.status_code != 200:
        raise Exception(f"Error HTTP {response.status_code}: {response.text}")

    data = response.json()
    
    # Extraer solo YYYY-MM-DD
    fecha_inicio = data[0]['min_fecha_corte'][:10]
    
    return fecha_inicio

    
def max_fecha():
    query = '''
https://www.datos.gov.co/resource/qzsc-9esp.json?$query=%20SELECT%20max(fecha_corte) AS max_fecha_corte
    '''
    response = requests.get(query)

    if response.status_code != 200:
        raise Exception(f"Error HTTP {response.status_code}: {response.text}")

    data = response.json()
    
    # Extraer solo YYYY-MM-DD
    fecha_inicio = data[0]['max_fecha_corte'][:10]
    
    return fecha_inicio


FECHA_INICIO = min_fecha()
FECHA_FIN = max_fecha()
#FECHA_INICIO = "2022-07-01"
#FECHA_FIN = "2026-01-02"
Sector = 'Agricultura, ganadería, silvicultura y pesca'
Lista_CIIU = [
    '0111',
    '0112',
    '0113',
    '0114',
    '0115',
    '0119',
    '0121',
    '0122',
    '0123',
    '0124',
    '0125',
    '0126',
    '0127',
    '0129',
    '0130',
    '0141',
    '0144',
    '0145',
    '0149',
    '0150',
    '0161',
    '0162',
    '0164',
    '0210',
    '0220',
    '0240',
    '0311',
    '0312',
    '0321',
    '0322' 
    ]

#ciiu_str = ",".join(f"'{x}'" for x in Lista_CIIU)
ciiu_str = ",".join(f"'{str(x).zfill(4)}'" for x in Lista_CIIU)


# ---------------- FUNCIONES ----------------

def generar_rangos_mensuales(fecha_inicio, fecha_fin):
    """
    Genera tuplas (fecha_desde, fecha_hasta) por mes
    """
    inicio = datetime.strptime(fecha_inicio, "%Y-%m-%d")
    fin = datetime.strptime(fecha_fin, "%Y-%m-%d")

    rangos = []
    actual = inicio.replace(day=1)

    while actual <= fin:
        if actual.month == 12:
            siguiente = actual.replace(year=actual.year + 1, month=1)
        else:
            siguiente = actual.replace(month=actual.month + 1)

        fin_mes = siguiente - timedelta(days=1)
        if fin_mes > fin:
            fin_mes = fin

        rangos.append((actual.date(), fin_mes.date()))
        actual = siguiente

    return rangos


def descargar_mes(Lista_CIIU, fecha_desde, fecha_hasta):
    """
    Descarga todos los registros para un mes específico
    """
    where_clause = (
        f"fecha_corte between '{fecha_desde}T00:00:00' "
        f"and '{fecha_hasta}T23:59:59' "
        f"AND Codigo_CIIU in ({ciiu_str})"
    )

    offset = 0
    filas = []

    while True:
        params = {
            "$limit": LIMIT,
            "$offset": offset,
            "$where": where_clause
        }

        r = requests.get(BASE_URL, params=params)
        if r.status_code != 200:
            raise Exception(f"Error HTTP {r.status_code}: {r.text}")

        data = r.json()
        if not data:
            break

        filas.extend(data)
        offset += LIMIT

    if filas:
        return pd.DataFrame(filas)
    else:
        return pd.DataFrame()


# ---------------- EJECUCIÓN ----------------

if __name__ == "__main__":

    rangos = generar_rangos_mensuales(FECHA_INICIO, FECHA_FIN)

    dfs_mensuales = []

    print(f"Total de meses a procesar: {len(rangos)}")

    for fecha_desde, fecha_hasta in rangos:
        print(f"Procesando {fecha_desde} -> {fecha_hasta}")

        df_mes = descargar_mes(
            Lista_CIIU,
            fecha_desde.strftime("%Y-%m-%d"),
            fecha_hasta.strftime("%Y-%m-%d")
        )

        if not df_mes.empty:
            df_mes["periodo"] = fecha_desde.strftime("%Y-%m")
            dfs_mensuales.append(df_mes)

            print(f"  Registros descargados: {len(df_mes)}")
        else:
            print("  Sin registros")

    # Apilar todos los meses
    if dfs_mensuales:
        df_final = pd.concat(dfs_mensuales, ignore_index=True)
        print(f"\n✅ Total registros final: {len(df_final)}")
    else:
        df_final = pd.DataFrame()
        print("\n⚠️ No se descargaron registros")
        
        
df_final.to_excel(f"e://riesgos/{Sector}_datos_abiertos_últimos_dos_meses.xlsx", index=False)
