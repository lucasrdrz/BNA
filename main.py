import streamlit as st
import pandas as pd
from datetime import timedelta
from io import BytesIO

st.set_page_config(
    page_title="Generador Maestro BNA",
    layout="wide"
)

st.title("📊 Generador Maestro BNA")

# =====================================================
# FUNCIONES
# =====================================================

def calcular_inicio_sla_habil(fecha):

    if pd.isna(fecha):
        return pd.NaT

    if fecha.weekday() == 5:
        return (fecha + timedelta(days=2)).replace(
            hour=8,
            minute=30,
            second=0,
            microsecond=0
        )

    if fecha.weekday() == 6:
        return (fecha + timedelta(days=1)).replace(
            hour=8,
            minute=30,
            second=0,
            microsecond=0
        )

    if fecha.hour < 8 or (
        fecha.hour == 8 and fecha.minute < 30
    ):
        return fecha.replace(
            hour=8,
            minute=30,
            second=0,
            microsecond=0
        )

    if fecha.hour > 17 or (
        fecha.hour == 17 and fecha.minute > 30
    ):

        if fecha.weekday() == 4:
            return (fecha + timedelta(days=3)).replace(
                hour=8,
                minute=30,
                second=0,
                microsecond=0
            )

        return (fecha + timedelta(days=1)).replace(
            hour=8,
            minute=30,
            second=0,
            microsecond=0
        )

    return fecha


def generar_maestro(archivo_bna, archivos_diebold):

    # ======================================
    # BNA
    # ======================================

    df_BNA = pd.read_excel(
        archivo_bna,
        sheet_name="5-Incidentes con SLA en $Estado"
    )

    # ======================================
    # DIEBOLD
    # ======================================

    lista_dfs = []

    for archivo in archivos_diebold:

        df_temp = pd.read_excel(
            archivo,
            sheet_name="ListaTickets"
        )

        lista_dfs.append(df_temp)

    df_maestro = pd.concat(
        lista_dfs,
        ignore_index=True
    )

    df_maestro = df_maestro.drop_duplicates(
        subset="TICKET NUMBER"
    )

    # ======================================
    # NORMALIZACION
    # ======================================

    df_maestro["TICKET CLIENTE"] = (
        df_maestro["TICKET CLIENTE"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )

    df_BNA["NumIncidente"] = (
        df_BNA["NumIncidente"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(".0", "", regex=False)
    )

    coincidencias = (
        df_maestro["TICKET CLIENTE"]
        .isin(df_BNA["NumIncidente"])
        .sum()
    )

    # ======================================
    # MERGE
    # ======================================

    df_final = pd.merge(
        df_maestro,
        df_BNA,
        left_on="TICKET CLIENTE",
        right_on="NumIncidente",
        how="left"
    )

    # ======================================
    # REPETICIONES
    # ======================================

    df_final["Repeticion"] = (
        df_final.groupby("LUNO")["LUNO"]
        .transform("count")
    )

    # ======================================
    # FECHAS
    # ======================================

    df_final["FechaHoraGeneracion"] = pd.to_datetime(
        df_final["FECHA GENERACIÓN TICKET"].astype(str)
        + " "
        + df_final["HORA GENERACIÓN TICKET"].astype(str),
        format="%d/%m/%Y %H:%M",
        errors="coerce"
    )

    df_final["InicioSLA"] = pd.to_datetime(
        df_final["InicioSLA"],
        errors="coerce"
    )

    df_final["InicioSLA_Habil"] = (
        df_final["InicioSLA"]
        .apply(calcular_inicio_sla_habil)
    )

    diferencia = (
        df_final["FechaHoraGeneracion"]
        - df_final["InicioSLA_Habil"]
    )

    diferencia = diferencia.clip(
        lower=pd.Timedelta(0)
    )

    df_final["Diferencia en generacion"] = (
        diferencia.astype(str)
    )

    df_final["Diferencia HHMMSS"] = (
        diferencia.astype(str)
        .str.replace(
            "0 days ",
            "",
            regex=False
        )
    )

    df_final["Diferencia en generacion Min"] = (
        diferencia.dt.total_seconds() / 60
    ).round(2)

    # ======================================
    # PENALIDAD
    # ======================================

    df_final["Penalidad"] = (
        df_final["TiempoExcedidoHR"] * 10
    ).round(2)

    # ======================================
    # RESUMENES
    # ======================================

    def crear_resumen(nombre):

        df = df_final[
            df_final["UsuarioAsignado"]
            .fillna("")
            .str.upper()
            == nombre.upper()
        ]

        cantidad = len(df)

        sla_ok = (
            df["TiempoExcedidoHR"]
            .fillna(0)
            .le(0)
            .sum()
        )

        porcentaje = (
            round(sla_ok / cantidad * 100, 2)
            if cantidad > 0
            else 0
        )

        demora = round(
            df["Diferencia en generacion Min"]
            .fillna(0)
            .sum(),
            2
        )

        penalidad = round(
            df["Penalidad"]
            .fillna(0)
            .sum(),
            2
        )

        return pd.DataFrame({
            "Indicador": [
                "Cantidad equipos",
                "% SLA",
                "Demora Min",
                "Penalidad USD"
            ],
            "Valor": [
                cantidad,
                porcentaje,
                demora,
                penalidad
            ]
        })

    df_confalone = crear_resumen(
        "FEDERICO CONFALONE"
    )

    df_gilardoni = crear_resumen(
        "FEDERICO CONRADO GILARDONI"
    )

    # ======================================
    # EXCEL
    # ======================================

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        df_final.to_excel(
            writer,
            sheet_name="Detalle",
            index=False
        )

        df_confalone.to_excel(
            writer,
            sheet_name="Federico Confalone",
            index=False
        )

        df_gilardoni.to_excel(
            writer,
            sheet_name="Federico Gilardoni",
            index=False
        )

    output.seek(0)

    return output, coincidencias, df_final


# =====================================================
# INTERFAZ
# =====================================================

archivo_bna = st.file_uploader(
    "Archivo BNA",
    type=["xlsx"]
)

archivos_diebold = st.file_uploader(
    "Reportes Diebold",
    type=["xlsx"],
    accept_multiple_files=True
)

if st.button("Generar Maestro"):

    if archivo_bna is None:
        st.error("Subí archivo BNA")
        st.stop()

    if len(archivos_diebold) == 0:
        st.error("Subí al menos un reporte Diebold")
        st.stop()

    with st.spinner("Procesando..."):

        excel_file, coincidencias, df_final = generar_maestro(
            archivo_bna,
            archivos_diebold
        )

    col1, col2, col3 = st.columns(3)

    col1.metric(
        "Tickets",
        len(df_final)
    )

    col2.metric(
        "Coincidencias",
        coincidencias
    )

    col3.metric(
        "Penalidad Total",
        round(
            df_final["Penalidad"]
            .fillna(0)
            .sum(),
            2
        )
    )

    st.success("Archivo generado correctamente")

    st.download_button(
        label="📥 Descargar Maestro",
        data=excel_file,
        file_name="archivo_maestro_bna_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
