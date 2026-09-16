"""
Buscador de Convocatorias — FundsForNGOs Premium
+ Evaluación y Priorización CTI (Agencia Atenea / Bankpro)

App de Streamlit para que usuarios no técnicos consulten, descarguen y
prioricen convocatorias sin tocar el notebook original. Ver README.md para
instrucciones de despliegue en Streamlit Community Cloud.
"""

from __future__ import annotations

import os
from datetime import date, timedelta

import altair as alt
import pandas as pd
import streamlit as st

from cti_scoring import evaluar_cti, tabla_cruzada_top_donantes
from ffn_client import (
    FFNAuthError,
    FFNClient,
    ISSUES_DISPONIBLES,
    PAISES_DISPONIBLES,
    REGIONES_DISPONIBLES,
    TAMANOS_GRANT_DISPONIBLES,
    TIPOS_GRANT_DISPONIBLES,
    construir_dataframe,
    exportar_excel,
    exportar_excel_generico,
    obtener_paises_por_region,
)

RUTA_LOGO = os.path.join(os.path.dirname(__file__), "assets", "logo_atenea.png")


# =============================================================================
# Utilidades de presentación (gráficos, tablas, encabezado)
# =============================================================================

def grafico_barras_ordenado(conteo: pd.Series, nombre_categoria: str) -> alt.Chart:
    """Gráfico de barras en orden descendente por conteo, con la cifra
    encima de cada barra."""
    datos = conteo.rename("Conteo").rename_axis(nombre_categoria).reset_index()

    barras = alt.Chart(datos).mark_bar(color="#065094").encode(
        x=alt.X(f"{nombre_categoria}:N", sort="-y", title=None, axis=alt.Axis(labelFontSize=13)),
        y=alt.Y("Conteo:Q", title="Convocatorias"),
        tooltip=[nombre_categoria, "Conteo"],
    )
    etiquetas = alt.Chart(datos).mark_text(dy=-12, fontSize=18, fontWeight="bold", color="#020142").encode(
        x=alt.X(f"{nombre_categoria}:N", sort="-y"),
        y="Conteo:Q",
        text="Conteo:Q",
    )
    return (barras + etiquetas).properties(height=340)


def html_tabla_cruzada(df: pd.DataFrame) -> str:
    """HTML con las columnas numéricas centradas y estilo propio (Donante
    a la izquierda, cifras centradas, encabezados y fila de Total
    resaltados). Se renderiza con st.markdown(..., unsafe_allow_html=True)
    porque st.table() no respeta de forma confiable el centrado de un
    pandas Styler."""
    columnas_numericas = [c for c in df.columns if c != "Donante"]
    estilo = (
        df.style
        .hide(axis="index")
        .set_properties(subset=columnas_numericas, **{"text-align": "center"})
        .set_properties(subset=["Donante"], **{"text-align": "left", "font-weight": "600"})
        .set_table_styles([
            {"selector": "th", "props": [
                ("text-align", "center"), ("background-color", "#F2F4F9"),
                ("color", "#020142"), ("padding", "8px 12px"),
            ]},
            {"selector": "td", "props": [("padding", "6px 12px")]},
            {"selector": "table", "props": [("border-collapse", "collapse"), ("width", "100%")]},
        ])
    )
    # Resalta la fila de Total en negrita.
    def resaltar_total(fila):
        return ["font-weight: 700; background-color: #F2F4F9" if fila["Donante"] == "Total" else "" for _ in fila]
    estilo = estilo.apply(resaltar_total, axis=1)
    return estilo.to_html()


def config_columnas_tabla(df: pd.DataFrame) -> dict:
    """column_config compartido por las tablas de convocatorias: título fijo
    al hacer scroll horizontal, URLs como links clicleables, y las columnas
    de texto largo con un ancho fijo (la grilla las trunca con '...' y se
    puede hacer clic en la celda para ver el contenido completo)."""
    config: dict = {}
    if "Titulo" in df.columns:
        config["Titulo"] = st.column_config.TextColumn("Titulo", pinned=True, width="medium")
    if "URL referencia" in df.columns:
        config["URL referencia"] = st.column_config.LinkColumn("URL referencia", display_text="🔗 Abrir")
    if "URL en plataforma" in df.columns:
        config["URL en plataforma"] = st.column_config.LinkColumn("URL en plataforma", display_text="🔗 Abrir")
    if "Descripcion completa" in df.columns:
        config["Descripcion completa"] = st.column_config.TextColumn("Descripcion completa", width="medium")
    if "Descripcion corta" in df.columns:
        config["Descripcion corta"] = st.column_config.TextColumn("Descripcion corta", width="medium")
    return config


_EMOJI_CATEGORIA = {"ALTA": "🟢 ALTA", "MEDIA": "🟡 MEDIA", "BAJA": "🔴 BAJA"}


def con_categoria_visual(df: pd.DataFrame) -> pd.DataFrame:
    """Copia del DataFrame solo para mostrar en pantalla, con un emoji de
    semáforo antepuesto a la categoría para que se distinga de un vistazo.
    (El Excel que se descarga usa el DataFrame original, sin emojis)."""
    if "Categoria Priorizacion" not in df.columns:
        return df
    copia = df.copy()
    copia["Categoria Priorizacion"] = copia["Categoria Priorizacion"].map(_EMOJI_CATEGORIA).fillna(
        copia["Categoria Priorizacion"]
    )
    return copia


def mostrar_encabezado(titulo: str) -> None:
    col_logo, col_titulo = st.columns([1, 4], vertical_alignment="center")
    with col_logo:
        if os.path.exists(RUTA_LOGO):
            st.image(RUTA_LOGO, width=260)
    with col_titulo:
        st.title(titulo)


st.set_page_config(
    page_title="Buscador de Convocatorias — Agencia Atenea",
    page_icon=RUTA_LOGO if os.path.exists(RUTA_LOGO) else "🔎",
    layout="wide",
)


# =============================================================================
# 1. Acceso con contraseña de la app (protege la cuenta compartida de FFN)
# =============================================================================

def pedir_contrasena() -> bool:
    if st.session_state.get("autenticado"):
        return True

    mostrar_encabezado("Buscador de Convocatorias")
    st.write("Ingresa la contraseña de acceso para continuar.")

    with st.form("form_login_app"):
        clave = st.text_input("Contraseña de acceso", type="password")
        enviado = st.form_submit_button("Entrar")

    if enviado:
        clave_correcta = st.secrets.get("APP_PASSWORD")
        if not clave_correcta:
            st.error(
                "No se ha configurado APP_PASSWORD en los secretos de la app. "
                "Revisa la sección de Secrets en Streamlit Community Cloud."
            )
        elif clave == clave_correcta:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")

    return False


if not pedir_contrasena():
    st.stop()


# =============================================================================
# 2. Cliente de FFN (login con la cuenta compartida, guardado en la sesión)
# =============================================================================

@st.cache_resource(show_spinner=False)
def obtener_cliente() -> FFNClient:
    email = st.secrets.get("FFN_EMAIL")
    password = st.secrets.get("FFN_PASSWORD")
    if not email or not password:
        raise FFNAuthError(
            "Faltan FFN_EMAIL y/o FFN_PASSWORD en los secretos de la app."
        )
    cliente = FFNClient(email, password)
    cliente.login()
    return cliente


mostrar_encabezado("Buscador de Convocatorias — FundsForNGOs")

with st.status("Conectando con FundsForNGOs...", expanded=False) as status_login:
    try:
        cliente = obtener_cliente()
        status_login.update(label="Conectado a FundsForNGOs", state="complete")
    except FFNAuthError as exc:
        status_login.update(label="No se pudo conectar", state="error")
        st.error(f"No se pudo iniciar sesión en FundsForNGOs: {exc}")
        st.stop()


# =============================================================================
# 3. Filtros
# =============================================================================

st.sidebar.header("Filtros de búsqueda")


@st.cache_data(show_spinner=False, ttl=3600)
def _paises_por_region_cached(region_id: int) -> list[dict]:
    return obtener_paises_por_region(region_id)


def calcular_catalogo_paises(ids_regiones: list[int]) -> list[dict]:
    """Si hay regiones seleccionadas, trae solo los países de esas regiones
    (uniendo varias si el usuario eligió más de una). Si algo falla en la
    consulta, cae de vuelta al catálogo completo en vez de romper la app."""
    if not ids_regiones:
        return PAISES_DISPONIBLES

    vistos: dict[int, dict] = {}
    try:
        for rid in ids_regiones:
            for p in _paises_por_region_cached(rid):
                vistos[p["id"]] = p
    except RuntimeError as exc:
        st.sidebar.warning(f"No se pudo filtrar países por región ({exc}). Mostrando el catálogo completo.")
        return PAISES_DISPONIBLES

    resultado = sorted(vistos.values(), key=lambda x: x["name"])
    return resultado or PAISES_DISPONIBLES


# --- Filtros más usados: siempre visibles ---
opciones_regiones = {r["name"]: r["id"] for r in REGIONES_DISPONIBLES}
regiones_sel = st.sidebar.multiselect("Región", list(opciones_regiones.keys()))

ids_regiones_sel = sorted(opciones_regiones[n] for n in regiones_sel)
catalogo_paises_actual = calcular_catalogo_paises(ids_regiones_sel)
opciones_paises = {p["name"]: p["id"] for p in catalogo_paises_actual}

# La key cambia según la región seleccionada: así, cuando cambias de región,
# Streamlit muestra el multiselect limpio con las nuevas opciones en vez de
# fallar porque un país seleccionado antes ya no está en la lista nueva.
key_paises = "paises_" + "_".join(str(i) for i in ids_regiones_sel) if ids_regiones_sel else "paises_todas"
paises_sel = st.sidebar.multiselect("Países", list(opciones_paises.keys()), key=key_paises)

st.sidebar.markdown("**Fechas de publicación**")
activar_publicacion = st.sidebar.checkbox("Filtrar por fecha de publicación", value=False)
if activar_publicacion:
    col_desde, col_hasta = st.sidebar.columns(2)
    fecha_desde = col_desde.date_input(
        "Desde", value=date.today() - timedelta(days=30), format="DD/MM/YYYY", key="pub_desde"
    )
    fecha_hasta = col_hasta.date_input("Hasta", value=date.today(), format="DD/MM/YYYY", key="pub_hasta")
else:
    fecha_desde = None
    fecha_hasta = None

st.sidebar.markdown("**Fecha de cierre (deadline)**")
activar_deadline = st.sidebar.checkbox("Filtrar por fecha de cierre", value=False)
if activar_deadline:
    col_dd_desde, col_dd_hasta = st.sidebar.columns(2)
    deadline_desde = col_dd_desde.date_input(
        "Desde", value=date.today(), format="DD/MM/YYYY", key="deadline_desde"
    )
    deadline_hasta = col_dd_hasta.date_input(
        "Hasta", value=date.today() + timedelta(days=90), format="DD/MM/YYYY", key="deadline_hasta"
    )
else:
    deadline_desde = None
    deadline_hasta = None

# --- Filtros avanzados: colapsados por defecto ---
with st.sidebar.expander("Filtros avanzados"):
    opciones_issues = {i["name"]: i["id"] for i in ISSUES_DISPONIBLES}
    areas_sel = st.multiselect("Áreas temáticas", list(opciones_issues.keys()))

    opciones_tamanos = {t["name"]: t["id"] for t in TAMANOS_GRANT_DISPONIBLES}
    tamanos_sel = st.multiselect("Monto del grant", list(opciones_tamanos.keys()))

    opciones_tipos = {t["name"]: t["id"] for t in TIPOS_GRANT_DISPONIBLES}
    tipos_sel = st.multiselect("Tipo de convocatoria", list(opciones_tipos.keys()))

palabra_clave = st.sidebar.text_input("Palabra clave", placeholder="Ej: agua, mujeres, clima")
incluir_vencidas = st.sidebar.checkbox("Incluir convocatorias vencidas", value=False)

limite_resultados = st.sidebar.number_input(
    "Límite de resultados a extraer (0 = sin límite)",
    min_value=0,
    value=50,
    step=10,
    help="Un número alto puede tardar varios minutos, porque se consulta el detalle de cada convocatoria una por una.",
)


def armar_filtros() -> dict:
    return {
        "regions": [{"id": opciones_regiones[n], "name": n} for n in regiones_sel],
        "countries": [{"id": opciones_paises[n], "name": n} for n in paises_sel],
        "issues": [{"id": opciones_issues[n], "name": n} for n in areas_sel],
        "grantSizes": [{"id": opciones_tamanos[n], "name": n} for n in tamanos_sel],
        "grantTypes": [{"id": opciones_tipos[n], "name": n} for n in tipos_sel],
        "SearchText": palabra_clave.strip(),
        "fromDate": fecha_desde,
        "toDate": fecha_hasta,
        "deadlineActivo": activar_deadline,
        "deadlineDesde": deadline_desde,
        "deadlineHasta": deadline_hasta,
        "includeExpired": incluir_vencidas,
    }


# =============================================================================
# 4. Buscar (conteo rápido, sin traer el detalle completo)
# =============================================================================

col_buscar, col_ejecutar = st.columns(2)

if col_buscar.button("🔍 Buscar / contar resultados", width="stretch"):
    filtros = armar_filtros()
    with st.status("Buscando y filtrando por fecha...", expanded=True) as status_conteo:
        barra_conteo = st.progress(0.0)

        def progreso_conteo(encontrados: int, total: int, pagina: int, pagina_maxima: int) -> None:
            pagina_maxima = max(pagina_maxima, 1)
            barra_conteo.progress(min(pagina / pagina_maxima, 1.0))
            status_conteo.update(label=f"Página {pagina}/{pagina_maxima} — {encontrados} dentro del rango de fechas")

        try:
            listado_conteo = cliente.buscar_convocatorias(
                filtros, limite=None, pausa_entre_paginas=0.2, on_progreso=progreso_conteo
            )
        except RuntimeError as exc:
            status_conteo.update(label="Error consultando resultados", state="error")
            st.error(f"Error consultando resultados: {exc}")
        else:
            st.session_state["ultimo_total"] = len(listado_conteo)
            st.session_state["filtros_actuales"] = filtros
            status_conteo.update(
                label=f"Listo: {len(listado_conteo)} convocatorias dentro del rango de fechas",
                state="complete",
                expanded=False,
            )

if "ultimo_total" in st.session_state:
    st.metric("Convocatorias encontradas (ya filtradas por fecha)", st.session_state["ultimo_total"])


# =============================================================================
# 5. Ejecutar extracción completa
# =============================================================================

st.divider()
st.subheader("Extracción completa")
st.caption(
    "Este paso consulta el detalle de cada convocatoria (descripción completa, "
    "países elegibles, áreas temáticas), por lo que puede tardar varios minutos "
    "según cuántos resultados haya."
)

if st.button("▶️ Ejecutar extracción y preparar Excel", type="primary"):
    filtros = st.session_state.get("filtros_actuales") or armar_filtros()
    limite = None if limite_resultados == 0 else int(limite_resultados)

    with st.status("Extrayendo convocatorias...", expanded=True) as status_extraccion:
        status_extraccion.update(label="Buscando convocatorias...")
        barra_busqueda = st.progress(0.0)

        def progreso_busqueda(encontrados: int, total: int, pagina: int, pagina_maxima: int) -> None:
            pagina_maxima = max(pagina_maxima, 1)
            barra_busqueda.progress(min(pagina / pagina_maxima, 1.0))
            status_extraccion.update(
                label=f"Buscando... página {pagina}/{pagina_maxima} — {encontrados} encontradas"
            )

        try:
            listado = cliente.buscar_convocatorias(filtros, limite=limite, on_progreso=progreso_busqueda)
        except RuntimeError as exc:
            status_extraccion.update(label="Error durante la búsqueda", state="error")
            st.error(f"Error durante la búsqueda: {exc}")
            st.stop()

        if not listado:
            status_extraccion.update(label="Sin resultados", state="complete")
            st.warning("No se encontraron convocatorias con estos filtros.")
            st.stop()

        status_extraccion.update(label=f"Extrayendo detalle de {len(listado)} convocatorias...")
        barra_detalle = st.progress(0.0)

        def progreso_detalle(hecho: int, total: int) -> None:
            total = max(total, 1)
            barra_detalle.progress(hecho / total)
            status_extraccion.update(label=f"Extrayendo detalle... {hecho}/{total}")

        df = construir_dataframe(cliente, listado, on_progreso=progreso_detalle)

        st.session_state["df_extraido"] = df
        status_extraccion.update(label=f"Listo: {len(df)} convocatorias extraídas", state="complete", expanded=False)


# Se guarda en session_state para que las pestañas de abajo sobrevivan a los
# re-runs de Streamlit (por ejemplo, al hacer clic en un botón de descarga).
df = st.session_state.get("df_extraido")
hay_datos = df is not None and not df.empty
df_cti = evaluar_cti(df) if hay_datos else None


# =============================================================================
# 6. Resultados, en pestañas
# =============================================================================

st.divider()
tab_extraccion, tab_cti, tab_prioridad = st.tabs(
    ["📋 Extracción completa", "🔬 Evaluación y Priorización CTI", "🏆 Oportunidades de Máxima Prioridad"]
)

with tab_extraccion:
    if not hay_datos:
        st.info("Ejecuta la extracción de arriba para ver resultados aquí.")
    else:
        st.metric("Convocatorias extraídas", len(df))
        st.dataframe(
            df,
            width="stretch",
            hide_index=True,
            column_config=config_columnas_tabla(df),
        )
        st.download_button(
            label="⬇️ Descargar Excel (extracción cruda)",
            data=exportar_excel(df),
            file_name="convocatorias.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_cti:
    if not hay_datos:
        st.info("Ejecuta la extracción de arriba para ver la evaluación CTI aquí.")
    else:
        st.caption(
            "Análisis automático para evaluar y priorizar convocatorias de Ciencia, "
            "Tecnología e Innovación (CTI) provenientes de la plataforma FundsforONG "
            "para BankPro de la Agencia Atenea."
        )

        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        col_m1.metric("Total evaluadas", len(df_cti))
        col_m2.metric("🟢 ALTA", int((df_cti["Categoria Priorizacion"] == "ALTA").sum()))
        col_m3.metric("🟡 MEDIA", int((df_cti["Categoria Priorizacion"] == "MEDIA").sum()))
        col_m4.metric("🔴 BAJA", int((df_cti["Categoria Priorizacion"] == "BAJA").sum()))

        st.write("### Tabla completa de evaluación CTI")
        st.dataframe(
            con_categoria_visual(df_cti),
            width="stretch",
            hide_index=True,
            column_config=config_columnas_tabla(df_cti),
        )
        st.download_button(
            label="⬇️ Descargar evaluación CTI completa (Excel)",
            data=exportar_excel_generico(df_cti, "Evaluacion_CTI"),
            file_name="evaluacion_cti.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        col_g1, col_g2 = st.columns(2)
        with col_g1:
            st.write("### Convocatorias por Categoría de Priorización")
            conteo_categoria = df_cti["Categoria Priorizacion"].value_counts()
            st.altair_chart(grafico_barras_ordenado(conteo_categoria, "Categoria"), width="stretch")

        with col_g2:
            st.write("### Convocatorias por Estado de Plazo")
            conteo_estado = df_cti["Estado Plazo"].value_counts()
            st.altair_chart(grafico_barras_ordenado(conteo_estado, "Estado"), width="stretch")

        st.write("### Top 5 Donantes por Categoría de Priorización")
        cruzada = tabla_cruzada_top_donantes(df_cti, top_n=5)
        st.markdown(html_tabla_cruzada(cruzada), unsafe_allow_html=True)
        st.download_button(
            label="⬇️ Descargar tabla cruzada (Excel)",
            data=exportar_excel_generico(cruzada, "Top5_Donantes"),
            file_name="top5_donantes_por_categoria.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_prioridad:
    if not hay_datos:
        st.info("Ejecuta la extracción de arriba para ver las oportunidades de máxima prioridad aquí.")
    else:
        df_alta = df_cti.loc[
            df_cti["Categoria Priorizacion"] == "ALTA",
            ["Titulo", "Donante", "Monto", "Plazo Semanas", "Puntuacion Final"],
        ]
        st.metric("Oportunidades ALTA", len(df_alta))

        if df_alta.empty:
            st.info("Ninguna convocatoria quedó clasificada como ALTA con los filtros actuales.")
        else:
            st.dataframe(
                df_alta,
                width="stretch",
                hide_index=True,
                column_config=config_columnas_tabla(df_alta),
            )
            st.download_button(
                label="⬇️ Descargar Oportunidades ALTA (Excel)",
                data=exportar_excel_generico(df_alta, "Alta_Prioridad"),
                file_name="oportunidades_alta_prioridad.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
