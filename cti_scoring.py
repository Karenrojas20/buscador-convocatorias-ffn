"""
Algoritmo de Evaluación y Puntuación de Convocatorias CTI (v5.2)
Agencia Atenea / Bankpro

Implementa la Especificación Técnica v5.2 (Especificacion_Tecnica_Algoritmo_
Puntuacion_CTI_v5.2.docx) sobre el DataFrame que ya produce
ffn_client.construir_dataframe(). No hace llamadas a la API: solo evalúa
los datos que ya se extrajeron.

--------------------------------------------------------------------------
NOTA IMPORTANTE SOBRE SUPUESTOS (transparencia, no está 100% en el documento)
--------------------------------------------------------------------------
La especificación define bandas de puntuación (ej. "7.0-8.0 pts") en vez de
un número exacto por caso, y para C1/C3 describe categorías cualitativas
("temáticas de desarrollo social/económico", "restricción explícita") que
requieren una regla determinística para programarse. Las decisiones tomadas
para cerrar esas ambigüedades son:

- C1 (Alineación Temática): se usa el punto medio de cada banda del
  documento (9.5 / 7.5 / 5.5 / 3.0). El cruce se hace primero contra la
  columna "Areas tematicas" (nombres tal cual los devuelve FundsForNGOs);
  si no hay coincidencia con el catálogo, se hace un respaldo por palabras
  clave sobre título + descripciones.
- C2 (Tiempo Disponible): dentro de cada banda ("0-4 semanas": 1.0-4.0,
  etc.) se interpola linealmente según cuántas semanas faltan, en vez de
  fijar un único valor para toda la banda.
- C3 (Accesibilidad Geográfica): se interpreta "restricción explícita"
  como que el campo "Paises elegibles" tiene una lista concreta de países
  que NO incluye a Colombia y no contiene ninguna palabra clave de
  cobertura regional amplia (Latin America, Global South, etc.). Con los
  datos que trae la plataforma no es posible distinguir una restricción
  "incierta" de una "explícita" (regla 4.3.3 del documento), así que toda
  detección de "Geografía restrictiva" se trata como el caso explícito
  (regla 4.3.2: fuerza C3 a 0.0-2.0 y Categoria_Priorizacion = 'BAJA').
- C4 (Viabilidad Financiera): se usa el punto medio de cada banda oficial
  por ID (9.5 / 7.5 / 5.5 / 3.5 / 1.5), y 4.0 (banda conservadora 3.0-5.0)
  cuando el monto no coincide con ninguno de los 6 rangos oficiales.
"""

from __future__ import annotations

from datetime import date, datetime

import pandas as pd

from ffn_client import ISSUES_DISPONIBLES, TAMANOS_GRANT_DISPONIBLES

# =============================================================================
# CRITERIO 1: Alineación Temática (peso 0.25)
# =============================================================================

# IDs de ISSUES_DISPONIBLES agrupados según la Sección 1.1 / Criterio 1 del
# documento. Water & Sanitation (24) se agrupó junto con Sustainable
# Development por cercanía temática (no está explícito en el documento).
_TEMAS_9_10 = {38, 62, 17, 68}  # Science, Innovation, IT, Research
_TEMAS_7_8 = {1, 2, 5, 3, 55, 24}  # Education, Health, Environment, Agriculture, Sustainable Dev, Water & Sanitation
_TEMAS_5_6 = {9, 26, 33, 65, 69, 35, 14, 25, 8, 10, 28, 63}  # desarrollo social/económico e industrias creativas

_SCORE_C1_ALTO = 9.5
_SCORE_C1_MEDIO_ALTO = 7.5
_SCORE_C1_MEDIO = 5.5
_SCORE_C1_BAJO = 3.0

_ISSUE_ID_POR_NOMBRE = {i["name"].strip().lower(): i["id"] for i in ISSUES_DISPONIBLES}

_PALABRAS_C1_ALTO = ["science", "innovation", "information technology", "research"]
_PALABRAS_C1_MEDIO_ALTO = ["education", "health", "environment", "agriculture", "sustainable development", "water"]
_PALABRAS_C1_MEDIO = [
    "arts", "culture", "economic development", "business", "industry", "manufacturing",
    "transportation", "international affairs", "women", "gender", "youth",
    "civil society", "democracy", "governance", "community development", "leadership",
]


def calcular_c1(areas_tematicas: str, titulo: str, descripcion_corta: str, descripcion_completa: str) -> float:
    """Devuelve la mejor puntuación C1 según los Issues cruzados (o texto libre de respaldo)."""
    mejor_score: float | None = None

    if areas_tematicas:
        for nombre in areas_tematicas.split(","):
            issue_id = _ISSUE_ID_POR_NOMBRE.get(nombre.strip().lower())
            if issue_id is None:
                continue
            if issue_id in _TEMAS_9_10:
                return _SCORE_C1_ALTO  # ya es el máximo posible
            candidato = (
                _SCORE_C1_MEDIO_ALTO if issue_id in _TEMAS_7_8
                else _SCORE_C1_MEDIO if issue_id in _TEMAS_5_6
                else None
            )
            if candidato is not None and (mejor_score is None or candidato > mejor_score):
                mejor_score = candidato

    if mejor_score is not None:
        return mejor_score

    texto = " ".join([titulo or "", descripcion_corta or "", descripcion_completa or ""]).lower()
    if any(p in texto for p in _PALABRAS_C1_ALTO):
        return _SCORE_C1_ALTO
    if any(p in texto for p in _PALABRAS_C1_MEDIO_ALTO):
        return _SCORE_C1_MEDIO_ALTO
    if any(p in texto for p in _PALABRAS_C1_MEDIO):
        return _SCORE_C1_MEDIO

    return _SCORE_C1_BAJO


# =============================================================================
# CRITERIO 2: Tiempo Disponible (peso 0.25)
# =============================================================================

_SCORE_C2_SIN_FECHA = 3.0  # punto medio de la banda 2.0-4.0


def calcular_c2(fecha_limite_ddmmyyyy: str, hoy: date | None = None) -> tuple[float, float | None, str]:
    """Devuelve (score_C2, plazo_semanas, estado_plazo).

    `fecha_limite_ddmmyyyy` viene ya formateada como DD/MM/YYYY (como la
    produce ffn_client.formatear_fecha) o vacía si no se pudo interpretar.
    """
    hoy = hoy or date.today()

    if not fecha_limite_ddmmyyyy:
        return _SCORE_C2_SIN_FECHA, None, "SIN_FECHA_CLARA"

    try:
        fecha_cierre = datetime.strptime(fecha_limite_ddmmyyyy, "%d/%m/%Y").date()
    except ValueError:
        return _SCORE_C2_SIN_FECHA, None, "SIN_FECHA_CLARA"

    dias = (fecha_cierre - hoy).days
    semanas = round(dias / 7, 1)

    if semanas < 0:
        return 0.0, semanas, "CERRADA"
    if semanas < 4:
        score = 1.0 + (semanas / 4.0) * 3.0
        return round(score, 2), semanas, "CIERRE_INMINENTE"
    if semanas <= 8:
        score = 7.0 + ((semanas - 4.0) / 4.0) * 1.0
        return round(score, 2), semanas, "ABIERTA"

    score = 9.0 + min((semanas - 8.0) / 8.0, 1.0) * 1.0
    return round(score, 2), semanas, "ABIERTA"


# =============================================================================
# CRITERIO 3: Accesibilidad Geográfica (peso 0.15)
# =============================================================================

COLOMBIA_ID = 48
_COLOMBIA_NOMBRE = "colombia"

_PALABRAS_REGION_DIRECTA = ["latin america", "south america", "caribbean", "all countries"]
_PALABRAS_REGION_AMPLIA = ["global south", "developing countries", "americas"]

_TEXTOS_SIN_INFO = {"", "no especificado", "not specified", "n/a", "sin especificar"}


def calcular_c3(paises_elegibles: str) -> dict:
    """Devuelve score, y las dos banderas de geografía (ver notas del módulo)."""
    texto = (paises_elegibles or "").strip()
    texto_l = texto.lower()

    if texto_l in _TEXTOS_SIN_INFO:
        return {"score": 6.0, "flag_restrictiva": False, "flag_insuficiente": True, "degradado": False}

    if _COLOMBIA_NOMBRE in texto_l:
        return {"score": 9.5, "flag_restrictiva": False, "flag_insuficiente": False, "degradado": False}

    if any(p in texto_l for p in _PALABRAS_REGION_DIRECTA):
        return {"score": 9.5, "flag_restrictiva": False, "flag_insuficiente": False, "degradado": False}

    if any(p in texto_l for p in _PALABRAS_REGION_AMPLIA):
        return {"score": 8.5, "flag_restrictiva": False, "flag_insuficiente": False, "degradado": False}

    # Hay una lista concreta de países, Colombia no aparece, y no hay señal
    # de cobertura regional amplia -> restricción explícita (regla 4.3.2).
    return {"score": 1.0, "flag_restrictiva": True, "flag_insuficiente": False, "degradado": True}


# =============================================================================
# CRITERIO 4: Viabilidad Financiera (peso 0.10)
# =============================================================================

_NOMBRE_A_ID_MONTO = {t["name"].strip().lower(): t["id"] for t in TAMANOS_GRANT_DISPONIBLES}
_SCORE_POR_ID_MONTO = {5: 9.5, 4: 7.5, 3: 5.5, 2: 3.5, 1: 1.5, 0: 1.5}
_SCORE_C4_NO_ESTANDARIZADO = 4.0  # punto medio de la banda conservadora 3.0-5.0


def calcular_c4(monto: str) -> tuple[float, int | None, bool]:
    """Devuelve (score_C4, id_tamano_o_None, flag_monto_no_estandarizado)."""
    monto_id = _NOMBRE_A_ID_MONTO.get((monto or "").strip().lower())
    if monto_id is not None:
        return _SCORE_POR_ID_MONTO[monto_id], monto_id, False
    return _SCORE_C4_NO_ESTANDARIZADO, None, True


# =============================================================================
# Puntuación final, categoría, banderas rojas y recomendación (Secciones 3-4)
# =============================================================================

_RECOMENDACIONES_BASE = {
    "ALTA": "ALTA: Aplicar inmediatamente. Buena alineación, plazos razonables y financiamiento atractivo.",
    "MEDIA": "MEDIA: Considerar según capacidad. Alineación aceptable, pero con algunas restricciones.",
    "BAJA": "BAJA: No priorizar. Alineación y/o condiciones poco favorables.",
}
_RECOMENDACION_MEDIA_GEOGRAFIA = (
    "MEDIA: Interés potencial, pero con restricciones geográficas para Colombia. "
    "Validar elegibilidad antes de priorizar."
)
_RECOMENDACION_BAJA_GEOGRAFIA = (
    "BAJA: No priorizar. La convocatoria presenta restricciones geográficas incompatibles con Colombia."
)


def evaluar_convocatoria(fila: dict, hoy: date | None = None) -> dict:
    """Evalúa una convocatoria (fila del DataFrame base) y devuelve las
    columnas nuevas del Esquema de Salida Estandarizado."""

    c1 = calcular_c1(
        fila.get("Areas tematicas", ""),
        fila.get("Titulo", ""),
        fila.get("Descripcion corta", ""),
        fila.get("Descripcion completa", ""),
    )
    c2_score, semanas, estado_plazo = calcular_c2(fila.get("Fecha limite", ""), hoy)
    c3 = calcular_c3(fila.get("Paises elegibles", ""))
    c4_score, monto_id, flag_monto_no_std = calcular_c4(fila.get("Monto", ""))

    puntuacion_final = round(
        ((c1 * 0.25) + (c2_score * 0.25) + (c3["score"] * 0.15) + (c4_score * 0.10)) / 0.75, 2
    )

    if puntuacion_final >= 7.50:
        categoria = "ALTA"
    elif puntuacion_final >= 6.50:
        categoria = "MEDIA"
    else:
        categoria = "BAJA"

    # Regla 4.3: exclusión geográfica obligatoria — nunca puede ser ALTA.
    degradado_por_geografia = False
    if c3["flag_restrictiva"]:
        categoria = "BAJA"
        degradado_por_geografia = True

    # Banderas rojas, en el orden de la Sección 4.1 del documento.
    banderas = []
    if estado_plazo == "CIERRE_INMINENTE":
        banderas.append("Plazo <4 semanas")
    if c3["flag_restrictiva"]:
        banderas.append("Geografía restrictiva para Colombia")
    if c4_score <= 4.0 or monto_id in (0, 1, 2):
        banderas.append("Monto bajo para esfuerzo requerido")
    if estado_plazo == "SIN_FECHA_CLARA":
        banderas.append("Fecha no interpretable")
    if flag_monto_no_std:
        banderas.append("Monto no estandarizado")
    if c3["flag_insuficiente"]:
        banderas.append("Información geográfica insuficiente")

    banderas_texto = "; ".join(banderas) if banderas else "Ninguna"

    if categoria == "MEDIA" and degradado_por_geografia:
        recomendacion = _RECOMENDACION_MEDIA_GEOGRAFIA
    elif categoria == "BAJA" and degradado_por_geografia:
        recomendacion = _RECOMENDACION_BAJA_GEOGRAFIA
    else:
        recomendacion = _RECOMENDACIONES_BASE[categoria]

    return {
        "Plazo Semanas": semanas,
        "Estado Plazo": estado_plazo,
        "Puntuacion_C1": round(c1, 2),
        "Puntuacion_C2": round(c2_score, 2),
        "Puntuacion_C3": round(c3["score"], 2),
        "Puntuacion_C4": round(c4_score, 2),
        "Puntuacion Final": puntuacion_final,
        "Categoria Priorizacion": categoria,
        "Recomendacion": recomendacion,
        "Banderas Rojas": banderas_texto,
    }


# =============================================================================
# Tabla completa (Esquema de Salida Estandarizado, Sección 5)
# =============================================================================

COLUMNAS_ESQUEMA_CTI = [
    "Titulo", "Tipo", "Donante", "Monto", "Fecha publicación", "Fecha limite",
    "Plazo Semanas", "Estado Plazo", "Descripcion corta", "Descripcion completa",
    "URL referencia", "Areas tematicas", "Paises elegibles",
    "Puntuacion_C1", "Puntuacion_C2", "Puntuacion_C3", "Puntuacion_C4",
    "Puntuacion Final", "Categoria Priorizacion", "Recomendacion", "Banderas Rojas",
    "URL en plataforma",
]


def evaluar_cti(df: pd.DataFrame, hoy: date | None = None) -> pd.DataFrame:
    """A partir del DataFrame base (el que ya produce construir_dataframe),
    calcula las 10 columnas nuevas y arma la tabla de 22 columnas."""
    if df.empty:
        return pd.DataFrame(columns=COLUMNAS_ESQUEMA_CTI)

    columnas_nuevas = df.apply(lambda fila: pd.Series(evaluar_convocatoria(fila.to_dict(), hoy)), axis=1)
    resultado = pd.concat([df.reset_index(drop=True), columnas_nuevas.reset_index(drop=True)], axis=1)

    resultado = resultado.rename(columns={
        "Categoria": "Tipo",
        "Fecha publicacion": "Fecha publicación",
    })

    for col in ["Fecha publicación", "Fecha limite"]:
        resultado[col] = resultado[col].apply(lambda v: v if v else "Sin fecha clara")

    return resultado[COLUMNAS_ESQUEMA_CTI]


# =============================================================================
# Tabla cruzada: top 5 donantes x Categoria Priorizacion (Sección adicional)
# =============================================================================

def tabla_cruzada_top_donantes(df_cti: pd.DataFrame, top_n: int = 5) -> pd.DataFrame:
    """Filas: los `top_n` donantes con más convocatorias. Columnas: ALTA/MEDIA/
    BAJA. Incluye fila y columna de Total."""
    if df_cti.empty:
        return pd.DataFrame(columns=["Donante", "ALTA", "MEDIA", "BAJA", "Total"])

    conteo_donantes = df_cti["Donante"].value_counts()
    top_donantes = conteo_donantes.head(top_n).index.tolist()

    subconjunto = df_cti[df_cti["Donante"].isin(top_donantes)]
    cruzada = pd.crosstab(subconjunto["Donante"], subconjunto["Categoria Priorizacion"])

    for cat in ["ALTA", "MEDIA", "BAJA"]:
        if cat not in cruzada.columns:
            cruzada[cat] = 0
    cruzada = cruzada[["ALTA", "MEDIA", "BAJA"]]
    cruzada = cruzada.reindex(top_donantes)
    cruzada["Total"] = cruzada.sum(axis=1)

    fila_total = pd.DataFrame(
        [[cruzada["ALTA"].sum(), cruzada["MEDIA"].sum(), cruzada["BAJA"].sum(), cruzada["Total"].sum()]],
        columns=["ALTA", "MEDIA", "BAJA", "Total"],
        index=["Total"],
    )
    cruzada = pd.concat([cruzada, fila_total])

    return cruzada.reset_index().rename(columns={"index": "Donante"})
