# Buscador de Convocatorias — FundsForNGOs Premium

App de Streamlit que reemplaza el notebook `scraper_final_v2.ipynb` con una
interfaz web sencilla: filtros en la barra lateral, vista previa de
resultados, extracción completa con barra de progreso y descarga directa
del Excel.

## Estructura

```
ffn_streamlit/
├── app.py                          # Interfaz (Streamlit)
├── ffn_client.py                   # Login, búsqueda, detalle y export a Excel
├── cti_scoring.py                  # Algoritmo de evaluación y priorización CTI (v5.2)
├── requirements.txt
├── .gitignore
├── assets/
│   └── logo_atenea.png             # Logo de Agencia Atenea, usado en el encabezado y como favicon
└── .streamlit/
    ├── config.toml                 # Tema visual (colores de marca de Atenea)
    └── secrets.toml.example        # Plantilla — NO subir la versión real
```

## Desplegar en Streamlit Community Cloud

1. **Sube esta carpeta a un repositorio de GitHub** (puede ser privado).
   No incluyas nunca un archivo `secrets.toml` real en el repo — el
   `.gitignore` ya lo excluye.

2. Entra a [share.streamlit.io](https://share.streamlit.io), conecta tu
   cuenta de GitHub y selecciona el repositorio. Como archivo principal
   indica `app.py`.

3. Antes de desplegar (o después, en **App settings → Secrets**), pega lo
   siguiente reemplazando los valores reales:

   ```toml
   APP_PASSWORD = "contraseña-que-vas-a-compartir-con-el-equipo"
   FFN_EMAIL = "correo-de-la-cuenta-compartida"
   FFN_PASSWORD = "contraseña-de-la-cuenta-de-FundsForNGOs"
   ```

   - `APP_PASSWORD` es la contraseña que pedirá la app antes de mostrar
     los filtros (protege que gente ajena use la cuenta de FFN).
   - `FFN_EMAIL` / `FFN_PASSWORD` son las credenciales de la cuenta
     **compartida** de FundsForNGOs.

4. Comparte la URL pública que te da Streamlit + la contraseña de acceso
   con el equipo (por un canal seguro, no por correo abierto).

## Probar en tu computador antes de desplegar

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
# Edita .streamlit/secrets.toml con las credenciales reales
streamlit run app.py
```

## Catálogos de filtros

Todos los filtros (región, país, área temática, monto y tipo de
convocatoria) son listas desplegables con nombres reales, cargadas en
`ffn_client.py`:

- **Países**: 251, verificado contra
  `https://core.fundsforngospremium.com/api/Country/GetRegionCountries?id=20`
- **Áreas temáticas**: 58
- **Regiones**: 34
- **Montos de grant**: 6
- **Tipos de convocatoria**: 6

Si en el futuro FundsForNGOs agrega nuevos países, áreas o tipos, hay que
repetir el proceso: capturar el JSON desde las Herramientas de
Desarrollador del navegador (F12 → pestaña Network → abrir el
desplegable correspondiente en la web) y actualizar la lista
correspondiente en `ffn_client.py`.

## Países filtrados por región (cascada)

Cuando seleccionas una o más regiones en el sidebar, el dropdown de
Países ya no muestra el catálogo completo (251): consulta en vivo
`GetRegionCountries?id=<id_region>` (el mismo endpoint público de arriba,
pero pasándole el id de la región en vez de 20) y arma el dropdown solo
con esos países. Si eliges varias regiones a la vez, se unen los
resultados sin duplicar. Esto vive en dos funciones de `ffn_client.py`:

- `obtener_paises_por_region(region_id)`: hace la consulta HTTP y
  normaliza la respuesta.
- En `app.py`, `calcular_catalogo_paises(ids_regiones)` combina varias
  regiones y cae de vuelta al catálogo completo (con un aviso visible)
  si la consulta falla por cualquier motivo.

## Fecha de cierre (deadline) — filtra bien en el servidor

A diferencia de la fecha de publicación (ver abajo), el filtro de
**fecha de cierre** sí funciona correctamente en el servidor de
FundsForNGOs. Se activa enviando `filterBy: "DD"` junto con
`fromDate`/`toDate` (confirmado capturando el payload real del sitio).
En la app es un filtro opcional — checkbox "Filtrar por fecha de cierre"
en el sidebar — que, cuando está activo, manda el rango directo a la
API: no necesita el escaneo local de páginas que sí requiere la fecha
de publicación, y se puede combinar con ella sin problema (cada una usa
un mecanismo distinto: una en el servidor, la otra localmente).

## Filtro de fechas de publicación (workaround local)

**El filtro de "fecha de publicación" está roto en el propio sitio de
FundsForNGOs** — no solo en esta app. Se confirmó probando directamente
en `grants.fundsforngospremium.com`: al filtrar por un rango de fechas,
el sitio sigue devolviendo convocatorias con `Post Date` fuera de ese
rango. No hay forma de arreglar esto pidiéndole a la API que filtre
bien, porque el problema es del servidor — por eso usa un campo
distinto (`filterBy`) del que sí funciona para la fecha de cierre.

La solución implementada es descartar el filtrado del lado de FFN y
hacerlo en Python, después de traer los resultados:

1. `ffn_client.buscar_convocatorias()` le pide a la API página por
   página (10 resultados cada una) y, cuando hay un rango de fechas de
   publicación activo, revisa la fecha (`lastUpdated`) de cada
   convocatoria con `fecha_en_rango()` y descarta las que caen fuera —
   sin dejar de recorrer páginas solo porque ya "llenó" el cupo
   aparente, porque el total que reporta el sitio tampoco está
   filtrado por esta fecha.
2. Por seguridad hay un tope de `MAX_PAGINAS_ESCANEADAS = 300` páginas
   (3000 resultados). Si se llega a ese tope, la app deja de buscar sin
   avisar en pantalla — si genera resultados incompletos, hay que subir
   esa constante en `ffn_client.py`.
3. Si además el filtro de fecha de cierre está activo, el "total" que
   usa este recorrido ya viene filtrado por el servidor, así que en la
   práctica hay muchas menos páginas que escanear (la combinación de
   ambos filtros es más rápida que la fecha de publicación sola).

Dos detalles más del payload que se encontraron capturando el payload
real del sitio:

- Las fechas se envían **sin ceros a la izquierda** (`2026-9-1`, no
  `2026-09-01`) — así es como las manda el propio frontend de FFN.
  `_formato_fecha_api()` en `ffn_client.py` hace esta conversión.
- `regionID` se envía como número entero (el id de la primera región
  seleccionada, o `0` si no hay ninguna) en vez del valor fijo `"19"`
  que traía el notebook original por error.

## Evaluación y Priorización CTI (Agencia Atenea / Bankpro)

Implementa la Especificación Técnica v5.2 (algoritmo de evaluación de
convocatorias de Ciencia, Tecnología e Innovación) en `cti_scoring.py`.
Se ejecuta automáticamente, sobre los mismos datos que ya trajo el paso
de "Extracción completa" — no vuelve a consultar FundsForNGOs.

Qué calcula por cada convocatoria:

- **C1 (Alineación Temática, 25%)**: cruza "Areas tematicas" contra el
  catálogo `ISSUES_DISPONIBLES`; si no hay coincidencia, hace un respaldo
  por palabras clave en título/descripción.
- **C2 (Tiempo Disponible, 25%)**: `Plazo_Semanas = (fecha_cierre - HOY) / 7`,
  calculado dinámicamente contra la fecha del sistema en cada corrida.
- **C3 (Accesibilidad Geográfica, 15%)**: revisa si "Paises elegibles"
  incluye a Colombia (ID 48) o alguna región/cobertura amplia; si es una
  lista concreta de países sin Colombia, activa la bandera roja
  **"Geografía restrictiva para Colombia"**.
- **C4 (Viabilidad Financiera, 10%)**: mapea "Monto" contra
  `TAMANOS_GRANT_DISPONIBLES` (IDs 0-5).

`Puntuacion_Final = [(C1×0.25) + (C2×0.25) + (C3×0.15) + (C4×0.10)] / 0.75`,
categorizada ALTA (≥7.50) / MEDIA (6.50-7.49) / BAJA (<6.50).

**Regla de exclusión geográfica (obligatoria):** si se activa la bandera
"Geografía restrictiva para Colombia", la convocatoria **nunca** puede
quedar como ALTA, sin importar la puntuación — se fuerza a BAJA. Esto
está implementado como una verificación posterior al cálculo de la
puntuación, no como parte de la fórmula.

`cti_scoring.py` documenta en su propio encabezado las decisiones que
tomé para cerrar ambigüedades del documento (ej. valor exacto dentro de
cada banda de puntuación, cómo distinguir restricción "explícita" de
"incierta" con los datos disponibles). Vale la pena revisarlas si en
algún momento los resultados no coinciden con lo esperado.

En la app, esta sección aparece debajo de la extracción cruda e incluye:
tarjetas de conteo, la tabla completa de 22 columnas, una tabla filtrada
solo con las convocatorias ALTA ("Oportunidades de Máxima Prioridad"),
dos gráficos de barras (por Categoría y por Estado de Plazo) y una tabla
cruzada de los 5 donantes con más convocatorias por categoría. Las tres
tablas tienen su propio botón de descarga en Excel.

## Identidad visual (Agencia Atenea)

- **Logo**: `assets/logo_atenea.png` se muestra junto al título (pantalla de
  login y encabezado principal) y como ícono de la pestaña del navegador.
  Si el archivo llegara a faltar, la app no se rompe: cae de vuelta a un
  emoji como ícono y simplemente no muestra la imagen.
- **Colores**: `.streamlit/config.toml` define el tema con la paleta de
  Atenea. Solo se usaron 2 de los 6 colores de la tabla que compartieron
  (naranja `#E65A0F` como color primario — botones, checkboxes, elementos
  seleccionados — y azul oscuro `#020142` como color de texto), porque el
  tema clásico de Streamlit solo expone 5 colores base y el fondo
  secundario (sidebar, campos de filtro) necesita quedarse neutro para
  que los controles se sigan viendo legibles. El resto de la paleta
  (azul `#065094`, azul violeta `#5263AB`, naranja claro `#EF7A45`) se
  usó puntualmente en las barras de los gráficos, pero queda disponible
  para quien quiera extenderla con CSS personalizado más adelante.
- Nota: la tabla de colores que compartieron tiene "Azul" y "Azul claro"
  con el mismo valor (`#065094`) — asumo que fue un error de tipeo al
  armar la tabla, no lo interpreté como dos colores distintos.

## Organización de la interfaz

- **Filtros del sidebar**: Región, País y Fechas quedan siempre visibles
  (son los más usados). Áreas temáticas, Monto y Tipo de convocatoria
  quedan colapsados dentro de un expander "Filtros avanzados".
- **Resultados en pestañas**: "Extracción completa", "Evaluación y
  Priorización CTI" y "Oportunidades de Máxima Prioridad" son pestañas
  separadas (antes se apilaban todas en la misma página).
- **Tablas**: título fijo al hacer scroll horizontal (`pinned=True`,
  requiere Streamlit 1.41+, ya actualizado en requirements.txt), las
  columnas de URL son links clicleables ("🔗 Abrir"), y las columnas de
  texto largo (descripciones) tienen un ancho fijo — la grilla las trunca
  visualmente y se puede hacer clic en la celda para ver el contenido
  completo.
- **Estados de carga**: conectar, buscar y extraer usan `st.status()`, que
  muestra un log expandible con cada paso en vez de un spinner genérico.

## Notas de seguridad

- La cuenta de FundsForNGOs es compartida por todos los usuarios de la
  app — no hay sesiones individuales por persona.
- `APP_PASSWORD`, `FFN_EMAIL` y `FFN_PASSWORD` viven únicamente en los
  *Secrets* de Streamlit Community Cloud, nunca en el código ni en el
  repositorio.
- Si sospechas que la contraseña de la app se filtró, cámbiala en
  **App settings → Secrets** y la app se reinicia con la nueva.
