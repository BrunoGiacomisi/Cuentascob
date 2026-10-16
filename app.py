"""
Interfaz Streamlit para generar el Excel semanal "CUENTAS A COBRAR".
Toda la lógica está en procesador.py; acá solo va la pantalla.
"""

from datetime import date

import streamlit as st

import procesador as p

st.set_page_config(page_title="Cuentas a Cobrar", page_icon="📊", layout="centered")
st.title("Cuentas a Cobrar – Generador semanal")
st.write(
    "Cargá los 3 Excel exportados del ERP, elegí la fecha del reporte "
    "y apretá **Generar**."
)

# --- Carga de archivos -------------------------------------------------------
archivo_agencia = st.file_uploader(p.ETIQUETA_AGENCIA, type=["xlsx"])
archivo_intercompany = st.file_uploader(p.ETIQUETA_INTERCOMPANY, type=["xlsx"])
archivo_ffww = st.file_uploader(p.ETIQUETA_FFWW, type=["xlsx"])

fecha = st.date_input(
    "Fecha del reporte",
    value=date.today(),
    format="DD/MM/YYYY",
    help="Se usa para el nombre del archivo y para controlar que en SAP se haya "
    "exportado con fecha de vencimiento hasta este día + 1 mes.",
)

archivos = {
    p.ETIQUETA_AGENCIA: archivo_agencia,
    p.ETIQUETA_INTERCOMPANY: archivo_intercompany,
    p.ETIQUETA_FFWW: archivo_ffww,
}

# "Firma" de lo cargado: si el usuario cambia un archivo o la fecha,
# el resultado anterior deja de mostrarse.
firma = tuple((f.name, f.size) if f else None for f in archivos.values()) + (fecha,)

# --- Generar -----------------------------------------------------------------
if st.button("Generar", type="primary"):
    st.session_state.pop("resultado", None)
    faltan = [etiqueta for etiqueta, f in archivos.items() if f is None]
    if faltan:
        st.error("Falta cargar: " + ", ".join(faltan))
    else:
        try:
            with st.spinner("Procesando..."):
                resultado = p.generar_reporte(
                    archivo_agencia, archivo_intercompany, archivo_ffww, fecha
                )
            avisos_nombres = p.revisar_nombres_archivos(
                {etiqueta: f.name for etiqueta, f in archivos.items()}
            )
            st.session_state["resultado"] = (firma, resultado, avisos_nombres)
        except p.ErrorArchivo as e:
            st.error(str(e))
        except Exception as e:  # error inesperado: mostrarlo sin romper la app
            st.error(f"Ocurrió un error inesperado: {e}")

# --- Resultado ---------------------------------------------------------------
guardado = st.session_state.get("resultado")
if guardado and guardado[0] == firma:
    _, resultado, avisos_nombres = guardado

    st.success("Reporte generado correctamente.")
    st.subheader("Resumen")
    st.table([
        {
            "Hoja": h.nombre,
            "Clientes": h.cantidad_clientes,
            "Documentos": h.cantidad_documentos,
            "Total": p.formatear_monto(h.total, h.moneda),
            f"Vencido +{p.DIAS_PARA_RECLAMAR} días": p.formatear_monto(
                p.indicadores(h, resultado.fecha).vencido, h.moneda
            ),
        }
        for h in resultado.hojas
    ])
    st.caption(
        f"La planilla completa incluye además las hojas «{p.HOJA_RECLAMAR}» (documentos con "
        f"más de {p.DIAS_PARA_RECLAMAR} días de vencidos, para reclamar) y «{p.HOJA_ANALISIS}». "
        f"La de FFWW lleva solo la hoja FFWW y «{p.HOJA_ANALISIS}» calculada con FFWW."
    )

    avisos = avisos_nombres + resultado.avisos
    if avisos:
        st.subheader(f"Avisos ({len(avisos)})")
        for aviso in avisos:
            st.warning(aviso)

    st.subheader("Descargas")
    mime_excel = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    col_completa, col_ffww = st.columns(2)
    with col_completa:
        st.markdown("**Planilla completa**")
        st.download_button(
            label=f"Descargar {resultado.nombre_archivo}",
            data=resultado.contenido,
            file_name=resultado.nombre_archivo,
            mime=mime_excel,
            type="primary",
            on_click="ignore",
        )
    with col_ffww:
        st.markdown("**Solo FFWW (para Comercial)**")
        st.download_button(
            label=f"Descargar {resultado.nombre_archivo_ffww}",
            data=resultado.contenido_ffww,
            file_name=resultado.nombre_archivo_ffww,
            mime=mime_excel,
            type="primary",
            on_click="ignore",
        )
elif guardado:
    st.info("Cambiaste algún archivo o la fecha: apretá **Generar** de nuevo.")
