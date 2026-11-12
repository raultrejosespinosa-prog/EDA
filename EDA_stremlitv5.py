"""EDA Streamlit v4 — importar CSV/Excel, explorar hasta 50 columnas, visualizar, calcular y exportar.

Ejecutar: python -m streamlit run EDA_stremlitv4.py
Python 3.10+ | Compatible con Python 3.12
Dependencias: pandas, numpy, streamlit, openpyxl, matplotlib, python-pptx
"""
from __future__ import annotations

from datetime import date, datetime, time
from io import BytesIO
import csv
import re
import unicodedata

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

MAX_COLUMNAS = 50
COLUMNAS_EDA = [
    "Nombre_col", "Ubicación", "Tipo_dato", "Max", "Min", "Promedio", "DS",
    "Mediana", "Nivel_cat", "Q1", "Q3", "Conteo", "Categoría", "Nulls",
]
NULOS_TEXTO = {"", "null", "none", "nan", "n/a", "na", "<na>"}
BOOL_TEXTO = {"true", "false", "verdadero", "falso", "sí", "si", "yes", "no"}
PATRON_FECHA = re.compile(r"(?:\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{4})(?:\s|$|T)")


def limpiar_nulos(s: pd.Series) -> pd.Series:
    """Unifica marcadores nulos explícitos sin alterar el archivo fuente."""
    if pd.api.types.is_object_dtype(s.dtype) or pd.api.types.is_string_dtype(s.dtype):
        return s.map(lambda v: pd.NA if isinstance(v, str) and v.strip().casefold() in NULOS_TEXTO else v)
    return s


def clasificar(s: pd.Series, nombre: str) -> tuple[str, pd.Series]:
    """Retorna (tipo legible, serie preparada). Evita convertir IDs/números a fechas."""
    s = limpiar_nulos(s)
    validos = s.dropna()
    if validos.empty:
        return "sin datos", s
    if pd.api.types.is_bool_dtype(s.dtype):
        return "booleano", s
    if pd.api.types.is_datetime64_any_dtype(s.dtype) or all(
        isinstance(x, (date, datetime, time, pd.Timestamp)) and not isinstance(x, bool)
        for x in validos.head(30)
    ):
        return "fechayhora", s
    if pd.api.types.is_numeric_dtype(s.dtype):
        return "numérico", pd.to_numeric(s, errors="coerce")

    valores = validos.astype(str).str.strip()
    formas = set(valores.str.casefold().unique())
    if formas.issubset(BOOL_TEXTO):
        return "booleano", s

    # Solo reconoce texto-fecha con estructura explícita (no años, IDs o seriales).
    indicio_nombre = any(k in unicodedata.normalize("NFKD", str(nombre)).lower() for k in ("fecha", "date", "time"))
    muestras = valores.head(100)
    candidatos = muestras.map(lambda x: bool(PATRON_FECHA.search(x)))
    if candidatos.mean() >= 0.9 or (indicio_nombre and candidatos.mean() >= 0.75):
        try:
            fechas = pd.to_datetime(s, errors="coerce", format="mixed", dayfirst=True)
            if fechas.notna().sum() / len(validos) >= 0.9:
                return "fechayhora", fechas
        except (ValueError, TypeError, OverflowError):
            pass

    # Números almacenados como texto, excepto identificadores con ceros iniciales.
    # No interpreta automáticamente monedas ni porcentajes ambiguos.
    if not valores.str.match(r"0\d+").any():
        numeros = pd.to_numeric(s, errors="coerce")
        if numeros.notna().sum() == len(validos):
            return "numérico", numeros
    return "categórico", s


def decimales_serie(serie: pd.Series) -> int:
    """0 decimales si todos los valores válidos son enteros; en otro caso, 1 decimal."""
    x = pd.to_numeric(serie, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    if x.empty:
        return 1
    valores = x.to_numpy(dtype=float)
    return 0 if np.all(np.isclose(valores, np.round(valores), rtol=0.0, atol=1e-9)) else 1


def redondear_numero(valor, decimales: int):
    """Redondea para presentación conservando enteros cuando corresponde."""
    if valor is None or pd.isna(valor):
        return "NA"
    if decimales == 0:
        return int(round(float(valor)))
    return round(float(valor), 1)


def formato_vega(decimales: int) -> str:
    """Formato numérico para tooltips/ejes de Vega-Lite."""
    return ",.0f" if decimales == 0 else ",.1f"


def formato_slider(decimales: int) -> str:
    return "%.0f" if decimales == 0 else "%.1f"


def paso_slider(decimales: int) -> float:
    return 1.0 if decimales == 0 else 0.1


def clave_orden_nivel(serie: pd.Series, nombre: str = "") -> pd.Series:
    """Genera una clave ordenable para jerarquías: fechas, números, meses o aparición."""
    s = serie.copy()

    if pd.api.types.is_datetime64_any_dtype(s):
        return pd.to_datetime(s, errors="coerce")

    num = pd.to_numeric(s, errors="coerce")
    if num.notna().sum() == s.notna().sum() and s.notna().any():
        return num

    meses = {
        "enero": 1, "ene": 1, "january": 1, "jan": 1,
        "febrero": 2, "feb": 2, "february": 2,
        "marzo": 3, "mar": 3, "march": 3,
        "abril": 4, "abr": 4, "april": 4, "apr": 4,
        "mayo": 5, "may": 5,
        "junio": 6, "jun": 6, "june": 6,
        "julio": 7, "jul": 7, "july": 7,
        "agosto": 8, "ago": 8, "august": 8, "aug": 8,
        "septiembre": 9, "setiembre": 9, "sep": 9, "sept": 9, "september": 9,
        "octubre": 10, "oct": 10, "october": 10,
        "noviembre": 11, "nov": 11, "november": 11,
        "diciembre": 12, "dic": 12, "december": 12, "dec": 12,
    }
    normal = (
        s.astype("string")
        .str.strip()
        .str.casefold()
        .map(
            lambda x: "" if pd.isna(x) else "".join(
                c for c in unicodedata.normalize("NFKD", str(x))
                if not unicodedata.combining(c)
            )
        )
    )
    orden_mes = normal.map(meses)
    if orden_mes.notna().sum() >= max(1, int(0.8 * s.notna().sum())):
        return orden_mes

    # Orden estable por primera aparición para categorías generales.
    codigos, _ = pd.factorize(s.astype("string"), sort=False)
    return pd.Series(codigos, index=s.index, dtype="int64")

def analizar(df: pd.DataFrame) -> pd.DataFrame:
    """Produce exactamente las 14 columnas solicitadas, hasta 50 variables."""
    registros = []
    for pos in range(min(df.shape[1], MAX_COLUMNAS)):
        nombre = str(df.columns[pos])
        tipo, s = clasificar(df.iloc[:, pos], nombre)
        nulos = int(s.isna().sum())
        fila = {clave: "NA" for clave in COLUMNAS_EDA}
        fila.update({
            "Nombre_col": nombre,
            "Ubicación": get_column_letter(pos + 1),
            "Tipo_dato": tipo,
            "Conteo": len(s),
            "Nulls": nulos,
        })
        if tipo == "numérico":
            x = pd.to_numeric(s, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
            if not x.empty:
                dec = decimales_serie(x)
                stats = {
                    "Max": x.max(),
                    "Min": x.min(),
                    "Promedio": x.mean(),
                    "DS": x.std(ddof=1) if len(x) > 1 else np.nan,
                    "Mediana": x.median(),
                    "Q1": x.quantile(0.25),
                    "Q3": x.quantile(0.75),
                }
                fila.update({
                    k: redondear_numero(v, dec) if pd.notna(v) else "NA"
                    for k, v in stats.items()
                })
        elif tipo == "categórico":
            niveles = s.dropna().astype(str).unique().tolist()
            fila["Nivel_cat"] = len(niveles)
            if len(niveles) <= 5:
                fila["Categoría"] = ", ".join(niveles)
        registros.append(fila)
    return pd.DataFrame(registros, columns=COLUMNAS_EDA)


def importar_archivo(contenido: bytes, nombre: str, hoja: str | None = None) -> pd.DataFrame:
    ext = nombre.lower().rsplit(".", 1)[-1]
    if ext == "csv":
        # CSV: autodetección del separador, tolerancia a BOM y codificación occidental.
        for codificacion in ("utf-8-sig", "cp1252", "latin-1"):
            try:
                muestra = contenido[:10000].decode(codificacion)
                try:
                    separador = csv.Sniffer().sniff(muestra, delimiters=",;\t|").delimiter
                except csv.Error:
                    separador = ";" if muestra.count(";") > muestra.count(",") else ","
                return pd.read_csv(BytesIO(contenido), sep=separador, encoding=codificacion,
                                   low_memory=False)
            except UnicodeDecodeError:
                continue
        raise ValueError("No se pudo determinar la codificación del CSV")
    if ext in ("xlsx", "xls"):
        return pd.read_excel(BytesIO(contenido), sheet_name=hoja or 0)
    raise ValueError("Formato no admitido: seleccione CSV, XLSX o XLS")


def hojas_excel(contenido: bytes, nombre: str) -> list[str]:
    if nombre.lower().endswith((".xlsx", ".xls")):
        return pd.ExcelFile(BytesIO(contenido)).sheet_names
    return []


def crear_excel(contenido: bytes, nombre: str, datos: pd.DataFrame,
                eda: pd.DataFrame, hoja_origen: str | None) -> bytes:
    """XLSX: conserva hojas fuente. CSV/XLS: crea un XLSX con Datos y EDA."""
    if nombre.lower().endswith(".xlsx"):
        libro = load_workbook(BytesIO(contenido))
        # Al elegir otra hoja se inserta EDA a su lado, sin sobrescribir una hoja existente.
        nombre_eda = "EDA"
        if nombre_eda in libro.sheetnames:
            sufijo = 2
            while f"EDA_{sufijo}" in libro.sheetnames:
                sufijo += 1
            nombre_eda = f"EDA_{sufijo}"  # Preserva todas las hojas existentes.
        indice = libro.sheetnames.index(hoja_origen) + 1 if hoja_origen in libro.sheetnames else 1
        hoja = libro.create_sheet(nombre_eda, index=indice)
    else:
        libro = Workbook()
        hoja_datos = libro.active
        hoja_datos.title = "Datos"
        for j, nombre_col in enumerate(datos.columns, start=1):
            hoja_datos.cell(1, j, str(nombre_col))
        for i, fila in enumerate(datos.itertuples(index=False, name=None), start=2):
            for j, valor in enumerate(fila, start=1):
                if pd.isna(valor):
                    continue
                if isinstance(valor, (np.integer,)):
                    valor = int(valor)
                elif isinstance(valor, (np.floating,)):
                    valor = float(valor)
                elif isinstance(valor, pd.Timestamp):
                    valor = valor.to_pydatetime()
                if isinstance(valor, str) and valor.startswith(("=", "+", "-", "@")):
                    # No ejecutar fórmulas importadas de archivos de texto.
                    valor = "'" + valor
                hoja_datos.cell(i, j, valor)
        hoja = libro.create_sheet("EDA", index=1)

    for j, nombre_col in enumerate(COLUMNAS_EDA, 1):
        c = hoja.cell(1, j, nombre_col)
        c.fill = PatternFill("solid", fgColor="17365D")
        c.font = Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(wrap_text=True)
    for i, fila in enumerate(eda.itertuples(index=False, name=None), 2):
        for j, valor in enumerate(fila, 1):
            if isinstance(valor, (np.integer,)):
                valor = int(valor)
            elif isinstance(valor, (np.floating,)):
                valor = float(valor)
            hoja.cell(i, j, valor)
            if j in (4, 5, 6, 7, 8, 10, 11) and isinstance(valor, (float, int)):
                hoja.cell(i, j).number_format = '#,##0' if float(valor).is_integer() else '#,##0.0'
    for j in range(1, len(COLUMNAS_EDA) + 1):
        hoja.column_dimensions[get_column_letter(j)].width = 35 if j == 13 else 17
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = hoja.dimensions
    salida = BytesIO()
    libro.save(salida)
    return salida.getvalue()





# -----------------------------------------------------------------------------
# Visualizaciones interactivas y LogicCal
# -----------------------------------------------------------------------------

def paleta_azul(n: int) -> list[str]:
    """Genera n colores entre #052570 y #8FAFFA, inclusive."""
    if n <= 0:
        return []
    inicio = (0x05, 0x25, 0x70)
    fin = (0x8F, 0xAF, 0xFA)
    if n == 1:
        return ["#052570"]
    colores = []
    for i in range(n):
        t = i / (n - 1)
        rgb = tuple(round(a + (b - a) * t) for a, b in zip(inicio, fin))
        colores.append("#%02X%02X%02X" % rgb)
    return colores


def metadatos_graficos(df: pd.DataFrame) -> list[dict]:
    """Clasifica hasta MAX_COLUMNAS para construir selectores de gráficos robustos."""
    metadatos = []
    limite = min(df.shape[1], MAX_COLUMNAS)
    for pos in range(limite):
        nombre = str(df.columns[pos])
        tipo, _ = clasificar(df.iloc[:, pos], nombre)
        metadatos.append({
            "pos": pos,
            "nombre": nombre,
            "tipo": tipo,
            "etiqueta": f"{nombre} [{get_column_letter(pos + 1)}]",
        })
    return metadatos


def serie_grafico(df: pd.DataFrame, meta: dict) -> pd.Series:
    """Retorna la serie preparada según la clasificación EDA."""
    _, serie = clasificar(df.iloc[:, meta["pos"]], meta["nombre"])
    return serie


def limitar_filas_grafico(df: pd.DataFrame, max_filas: int = 50000) -> tuple[pd.DataFrame, bool]:
    """Evita enviar cantidades excesivas de puntos al navegador."""
    if len(df) <= max_filas:
        return df, False
    return df.sample(n=max_filas, random_state=42), True


def rango_numerico(serie: pd.Series) -> tuple[float, float] | None:
    """Devuelve (mínimo, máximo) finitos de una serie numérica o None si no hay datos."""
    x = pd.to_numeric(serie, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    if x.empty:
        return None
    return float(x.min()), float(x.max())


def resumen_boxplot(base: pd.DataFrame) -> pd.DataFrame:
    """Calcula estadísticos de Tukey (1.5*IQR) por categoría usando todos los datos válidos."""
    filas = []
    for categoria, grupo in base.groupby("_categoria", sort=False):
        x = pd.to_numeric(grupo["_valor"], errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
        if x.empty:
            continue
        q1 = float(x.quantile(0.25))
        mediana = float(x.median())
        q3 = float(x.quantile(0.75))
        iqr = q3 - q1
        limite_inf = q1 - 1.5 * iqr
        limite_sup = q3 + 1.5 * iqr
        dentro = x[(x >= limite_inf) & (x <= limite_sup)]
        bigote_inf = float(dentro.min()) if not dentro.empty else float(x.min())
        bigote_sup = float(dentro.max()) if not dentro.empty else float(x.max())
        outliers = int(((x < bigote_inf) | (x > bigote_sup)).sum())
        filas.append({
            "Categoría": str(categoria),
            "N": int(x.count()),
            "Promedio": float(x.mean()),
            "Mínimo": float(x.min()),
            "Q1": q1,
            "Mediana": mediana,
            "Q3": q3,
            "Máximo": float(x.max()),
            "IQR": float(iqr),
            "Bigote inferior": bigote_inf,
            "Bigote superior": bigote_sup,
            "Outliers": outliers,
        })
    return pd.DataFrame(filas)



def _figura_png(fig) -> bytes:
    """Convierte una figura Matplotlib a PNG en memoria y libera sus recursos."""
    import matplotlib.pyplot as plt

    salida = BytesIO()
    fig.savefig(salida, format="png", dpi=170, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    salida.seek(0)
    return salida.getvalue()


def crear_png_boxplot(base: pd.DataFrame, titulo_num: str, titulo_cat: str) -> bytes:
    """Genera una imagen PNG horizontal del boxplot actualmente seleccionado."""
    import matplotlib.pyplot as plt

    df = base.copy()
    df["_categoria"] = df["_categoria"].astype(str)
    categorias = df["_categoria"].drop_duplicates().tolist()
    grupos = [df.loc[df["_categoria"] == c, "_valor"].dropna().to_numpy() for c in categorias]
    colores = paleta_azul(len(categorias))

    alto = max(4.8, min(11.0, 0.42 * len(categorias) + 3.5))
    fig, ax = plt.subplots(figsize=(9.6, alto))
    # No se usa labels/tick_labels dentro de boxplot para mantener compatibilidad
    # entre versiones de Matplotlib.
    bp = ax.boxplot(grupos, patch_artist=True, vert=False)
    ax.set_yticks(range(1, len(categorias) + 1))
    ax.set_yticklabels(categorias)

    for caja, color in zip(bp["boxes"], colores):
        caja.set_facecolor(color)
        caja.set_edgecolor("#052570")
    for elemento in ("whiskers", "caps", "medians"):
        for linea in bp[elemento]:
            linea.set_color("#052570")
    for flier in bp.get("fliers", []):
        flier.set_markeredgecolor("#052570")
        flier.set_markerfacecolor("#8FAFFA")

    ax.set_xlabel(titulo_num)
    ax.set_ylabel(titulo_cat)
    ax.set_title(f"Boxplot de {titulo_num}")
    ax.grid(axis="x", alpha=0.20)
    fig.tight_layout()
    return _figura_png(fig)


def crear_png_barras(resumen: pd.DataFrame, titulo_x: str, titulo_y: str) -> bytes:
    """Genera PNG del gráfico de barras horizontal actualmente seleccionado."""
    import matplotlib.pyplot as plt

    df = resumen.copy()
    categorias = df["_categoria"].astype(str).tolist()
    colores = paleta_azul(len(categorias))

    fig, ax = plt.subplots(figsize=(9.2, 5.3))
    ax.barh(categorias, df["_valor"].to_numpy(), color=colores)
    ax.invert_yaxis()
    ax.set_xlabel(titulo_x)
    ax.set_ylabel(titulo_y)
    ax.set_title(f"{titulo_x} por {titulo_y}")
    ax.grid(axis="x", alpha=0.20)
    fig.tight_layout()
    return _figura_png(fig)


def crear_png_lineas(graf: pd.DataFrame, titulo_x: str, titulo_y: str, titulo_cat: str) -> bytes:
    """Genera PNG del gráfico de líneas actualmente seleccionado, incluyendo jerarquía en X."""
    import matplotlib.pyplot as plt

    df = graf.copy()
    df["_categoria"] = df["_categoria"].astype(str)
    categorias = df["_categoria"].drop_duplicates().tolist()
    colores = paleta_azul(len(categorias))

    etiquetas = df[["_x_label", "_x_order"]].drop_duplicates().sort_values("_x_order")
    orden_labels = etiquetas["_x_label"].tolist()
    posiciones = {et: i for i, et in enumerate(orden_labels)}

    fig, ax = plt.subplots(figsize=(10.2, 5.6))
    for categoria, color in zip(categorias, colores):
        g = df[df["_categoria"] == categoria].sort_values("_x_order")
        xs = [posiciones[v] for v in g["_x_label"]]
        ax.plot(xs, g["_y"], marker="o", markersize=3.5, linewidth=1.6,
                label=categoria, color=color)

    ax.set_xticks(range(len(orden_labels)))
    ax.set_xticklabels(orden_labels, rotation=45, ha="right")
    ax.set_xlabel(titulo_x)
    ax.set_ylabel(titulo_y)
    ax.set_title(f"{titulo_y} vs. {titulo_x}")
    ax.grid(alpha=0.20)
    if len(categorias) <= 20:
        ax.legend(title=titulo_cat, loc="best", fontsize=8)
    fig.tight_layout()
    return _figura_png(fig)


def mostrar_boxplot(st, datos: pd.DataFrame, metas: list[dict]) -> None:
    st.subheader("Boxplot horizontal")
    numericas = [m for m in metas if m["tipo"] == "numérico"]
    categoricas = [m for m in metas if m["tipo"] in ("categórico", "booleano")]

    if not numericas:
        st.info("No se detectaron variables numéricas para construir un boxplot.")
        return

    mapa_num = {m["etiqueta"]: m for m in numericas}
    mapa_cat = {m["etiqueta"]: m for m in categoricas}

    col1, col2 = st.columns(2)
    seleccion_num = col1.selectbox(
        "Variable numérica (eje X)", list(mapa_num), key="box_num"
    )
    opciones_cat = ["Sin agrupación"] + list(mapa_cat)
    seleccion_cat = col2.selectbox(
        "Variable categórica (eje Y y color)", opciones_cat, key="box_cat"
    )

    y = pd.to_numeric(serie_grafico(datos, mapa_num[seleccion_num]), errors="coerce")
    dec = decimales_serie(y)
    base = pd.DataFrame({"_valor": y})

    if seleccion_cat == "Sin agrupación":
        base["_categoria"] = "Todos"
        titulo_cat = "Grupo"
    else:
        cat = serie_grafico(datos, mapa_cat[seleccion_cat])
        base["_categoria"] = cat.astype("string")
        titulo_cat = mapa_cat[seleccion_cat]["nombre"]

    base = base.replace([np.inf, -np.inf], np.nan).dropna(subset=["_valor", "_categoria"])
    if base.empty:
        st.warning("No hay datos válidos para construir el boxplot con esta selección.")
        return

    tabla_resumen = resumen_boxplot(base)

    graf, muestreado = limitar_filas_grafico(base)
    if muestreado:
        st.warning(
            "El gráfico usa una muestra reproducible de 50 000 filas para mantener la aplicación fluida; "
            "la tabla resumen usa todos los datos válidos."
        )

    graf["_categoria"] = graf["_categoria"].astype(str)
    categorias = graf["_categoria"].unique().tolist()
    colores = paleta_azul(len(categorias))

    spec = {
        "mark": {"type": "boxplot", "extent": 1.5, "size": 38},
        "encoding": {
            "x": {
                "field": "_valor", "type": "quantitative",
                "title": mapa_num[seleccion_num]["nombre"],
                "axis": {"format": formato_vega(dec)},
            },
            "y": {
                "field": "_categoria", "type": "nominal",
                "title": titulo_cat, "sort": None,
                "axis": {"labelLimit": 320},
            },
            "color": {
                "field": "_categoria", "type": "nominal",
                "title": titulo_cat,
                "scale": {"domain": categorias, "range": colores},
                "legend": None if seleccion_cat == "Sin agrupación" else {},
            },
            "tooltip": [
                {"field": "_categoria", "type": "nominal", "title": titulo_cat},
                {
                    "field": "_valor", "type": "quantitative",
                    "title": mapa_num[seleccion_num]["nombre"],
                    "format": formato_vega(dec),
                },
            ],
        },
        "height": max(360, min(900, 34 * len(categorias) + 180)),
    }
    st.vega_lite_chart(graf, spec, use_container_width=True)

    # Se guarda siempre la imagen actual para Data Storytelling.
    try:
        st.session_state["story_box_png"] = crear_png_boxplot(
            base, mapa_num[seleccion_num]["nombre"], titulo_cat
        )
        st.session_state["story_box_desc"] = (
            f"Boxplot horizontal: {mapa_num[seleccion_num]['nombre']} por "
            f"{titulo_cat if seleccion_cat != 'Sin agrupación' else 'Todos'}"
        )
        st.session_state["story_box_selection"] = {
            "num": seleccion_num,
            "cat": seleccion_cat,
        }
    except Exception as error:
        st.session_state.pop("story_box_png", None)
        st.caption(f"No fue posible preparar la imagen del boxplot para Storytelling: {error}")

    st.markdown("#### Resumen estadístico del boxplot")
    st.caption(
        "Bigotes calculados con la regla de Tukey: Q1 − 1.5×IQR y Q3 + 1.5×IQR. "
        "Outliers = observaciones fuera de los bigotes."
    )
    tabla_mostrar = tabla_resumen.copy()
    columnas_decimales = [
        "Promedio", "Mínimo", "Q1", "Mediana", "Q3", "Máximo",
        "IQR", "Bigote inferior", "Bigote superior",
    ]
    for col in columnas_decimales:
        if dec == 0:
            tabla_mostrar[col] = tabla_mostrar[col].map(
                lambda v: int(round(v)) if pd.notna(v) else v
            )
        else:
            tabla_mostrar[col] = tabla_mostrar[col].round(1)
    st.dataframe(tabla_mostrar, use_container_width=True, hide_index=True)


def mostrar_barras(st, datos: pd.DataFrame, metas: list[dict]) -> None:
    st.subheader("Gráfico de barras horizontal")
    numericas = [m for m in metas if m["tipo"] == "numérico"]
    categoricas = [m for m in metas if m["tipo"] in ("categórico", "booleano")]

    if not numericas or not categoricas:
        st.info("Se necesita al menos una variable numérica y una variable categórica.")
        return

    mapa_num = {m["etiqueta"]: m for m in numericas}
    mapa_cat = {m["etiqueta"]: m for m in categoricas}

    c1, c2, c3 = st.columns(3)
    seleccion_num = c1.selectbox(
        "Variable numérica (eje X)", list(mapa_num), key="bar_num"
    )
    seleccion_cat = c2.selectbox(
        "Variable categórica (eje Y y color)", list(mapa_cat), key="bar_cat"
    )
    agregacion = c3.selectbox(
        "Resumen por categoría", ["Promedio", "Suma", "Mediana", "Conteo"], key="bar_agg"
    )

    x = pd.to_numeric(serie_grafico(datos, mapa_num[seleccion_num]), errors="coerce")
    dec = decimales_serie(x)
    cat = serie_grafico(datos, mapa_cat[seleccion_cat]).astype("string")
    base = pd.DataFrame({"_valor": x, "_categoria": cat})
    base = base.replace([np.inf, -np.inf], np.nan).dropna(subset=["_categoria"])

    rango = rango_numerico(base["_valor"])
    if rango is None:
        st.warning("La variable numérica seleccionada no contiene valores válidos.")
        return

    minimo, maximo = rango
    if minimo < maximo:
        seleccion_rango = st.slider(
            f"Filtro de {mapa_num[seleccion_num]['nombre']} (eje X)",
            min_value=float(minimo),
            max_value=float(maximo),
            value=(float(minimo), float(maximo)),
            step=paso_slider(dec),
            format=formato_slider(dec),
            key=f"bar_slider_x_{mapa_num[seleccion_num]['pos']}",
        )
        base = base[base["_valor"].between(seleccion_rango[0], seleccion_rango[1], inclusive="both")]
    else:
        st.info(
            f"{mapa_num[seleccion_num]['nombre']} tiene un único valor válido: "
            f"{formatear_resultado(minimo)}."
        )

    if base.empty:
        st.warning("El filtro seleccionado no deja observaciones para construir el gráfico.")
        return

    if agregacion == "Promedio":
        resumen = base.dropna(subset=["_valor"]).groupby("_categoria", as_index=False)["_valor"].mean()
    elif agregacion == "Suma":
        resumen = base.dropna(subset=["_valor"]).groupby("_categoria", as_index=False)["_valor"].sum()
    elif agregacion == "Mediana":
        resumen = base.dropna(subset=["_valor"]).groupby("_categoria", as_index=False)["_valor"].median()
    else:
        resumen = base.groupby("_categoria", as_index=False)["_valor"].count()

    if resumen.empty:
        st.warning("No hay datos válidos para construir el gráfico de barras.")
        return

    resumen["_categoria"] = resumen["_categoria"].astype(str)
    resumen = resumen.sort_values("_valor", ascending=False)
    if len(resumen) > 50:
        st.warning("La variable categórica tiene más de 50 niveles; se muestran los 50 con mayor valor resumido.")
        resumen = resumen.head(50)

    # Conteo siempre es entero; las demás agregaciones siguen el tipo visual de la variable.
    dec_graf = 0 if agregacion == "Conteo" else dec
    categorias = resumen["_categoria"].tolist()
    colores = paleta_azul(len(categorias))
    titulo_x = f"{agregacion} de {mapa_num[seleccion_num]['nombre']}"

    spec = {
        "mark": {"type": "bar"},
        "encoding": {
            "x": {
                "field": "_valor", "type": "quantitative", "title": titulo_x,
                "axis": {"format": formato_vega(dec_graf)},
            },
            "y": {
                "field": "_categoria", "type": "nominal",
                "title": mapa_cat[seleccion_cat]["nombre"], "sort": "-x",
            },
            "color": {
                "field": "_categoria", "type": "nominal",
                "title": mapa_cat[seleccion_cat]["nombre"],
                "scale": {"domain": categorias, "range": colores},
            },
            "tooltip": [
                {"field": "_categoria", "type": "nominal", "title": mapa_cat[seleccion_cat]["nombre"]},
                {
                    "field": "_valor", "type": "quantitative",
                    "title": titulo_x, "format": formato_vega(dec_graf),
                },
            ],
        },
        "height": max(350, min(900, 24 * len(resumen))),
    }
    st.vega_lite_chart(resumen, spec, use_container_width=True)

    try:
        st.session_state["story_bar_png"] = crear_png_barras(
            resumen, titulo_x, mapa_cat[seleccion_cat]["nombre"]
        )
        st.session_state["story_bar_desc"] = (
            f"Barras: {titulo_x} por {mapa_cat[seleccion_cat]['nombre']}"
        )
    except Exception as error:
        st.session_state.pop("story_bar_png", None)
        st.caption(f"No fue posible preparar la imagen de barras para PowerPoint: {error}")


def mostrar_lineas(st, datos: pd.DataFrame, metas: list[dict]) -> None:
    st.subheader("Gráfico de líneas con jerarquía en el eje X")

    numericas = [m for m in metas if m["tipo"] == "numérico"]
    categoricas = [m for m in metas if m["tipo"] in ("categórico", "booleano")]
    disponibles_x = [m for m in metas if m["tipo"] != "sin datos"]

    if not numericas:
        st.info("Se necesita al menos una variable numérica para el eje Y.")
        return
    if not categoricas:
        st.info("Se necesita al menos una variable categórica para definir el color.")
        return

    mapa_num = {m["etiqueta"]: m for m in numericas}
    mapa_cat = {m["etiqueta"]: m for m in categoricas}
    mapa_x = {m["etiqueta"]: m for m in disponibles_x}

    c1, c2, c3, c4 = st.columns(4)
    seleccion_n1 = c1.selectbox(
        "Nivel 1 eje X (detalle, ej. Mes)", list(mapa_x), key="line_x_n1"
    )
    opciones_n2 = ["Sin segundo nivel"] + list(mapa_x)
    seleccion_n2 = c2.selectbox(
        "Nivel 2 eje X (grupo, ej. Año)", opciones_n2, key="line_x_n2"
    )
    seleccion_y = c3.selectbox(
        "Variable numérica (eje Y)", list(mapa_num), key="line_y"
    )
    seleccion_cat = c4.selectbox(
        "Variable categórica (color)", list(mapa_cat), key="line_cat"
    )

    agregacion = st.selectbox(
        "Resumen de Y por combinación jerárquica",
        ["Promedio", "Suma", "Mediana", "Conteo"],
        key="line_agg",
    )

    n1 = serie_grafico(datos, mapa_x[seleccion_n1])
    y = pd.to_numeric(serie_grafico(datos, mapa_num[seleccion_y]), errors="coerce")
    dec_y = decimales_serie(y)
    cat = serie_grafico(datos, mapa_cat[seleccion_cat]).astype("string")

    base = pd.DataFrame({
        "_n1": n1.astype("string"),
        "_y": y,
        "_categoria": cat,
        "_ord1": clave_orden_nivel(n1, mapa_x[seleccion_n1]["nombre"]),
    })

    if seleccion_n2 == "Sin segundo nivel":
        base["_n2"] = ""
        base["_ord2"] = 0
        titulo_x = mapa_x[seleccion_n1]["nombre"]
    else:
        n2 = serie_grafico(datos, mapa_x[seleccion_n2])
        base["_n2"] = n2.astype("string")
        base["_ord2"] = clave_orden_nivel(n2, mapa_x[seleccion_n2]["nombre"])
        titulo_x = (
            f"{mapa_x[seleccion_n2]['nombre']} → "
            f"{mapa_x[seleccion_n1]['nombre']}"
        )

    base = base.replace([np.inf, -np.inf], np.nan).dropna(
        subset=["_n1", "_y", "_categoria", "_ord1", "_ord2"]
    )
    if base.empty:
        st.warning("No hay datos válidos para construir el gráfico de líneas.")
        return

    # Slider únicamente cuando el Nivel 1 es numérico.
    tipo_n1 = mapa_x[seleccion_n1]["tipo"]
    if tipo_n1 == "numérico":
        n1_num = pd.to_numeric(serie_grafico(datos, mapa_x[seleccion_n1]), errors="coerce")
        rango = rango_numerico(n1_num)
        if rango is not None:
            minimo, maximo = rango
            dec_x = decimales_serie(n1_num)
            if minimo < maximo:
                seleccion_rango = st.slider(
                    f"Filtro de {mapa_x[seleccion_n1]['nombre']} (Nivel 1 eje X)",
                    min_value=float(minimo),
                    max_value=float(maximo),
                    value=(float(minimo), float(maximo)),
                    step=paso_slider(dec_x),
                    format=formato_slider(dec_x),
                    key=f"line_slider_x_{mapa_x[seleccion_n1]['pos']}",
                )
                # Aplicar el filtro por índice original antes de resumir.
                mascara_rango = n1_num.between(
                    seleccion_rango[0], seleccion_rango[1], inclusive="both"
                )
                base = base.loc[base.index.intersection(datos.index[mascara_rango])]

    if base.empty:
        st.warning("El filtro seleccionado no deja observaciones para construir el gráfico.")
        return

    claves = ["_n2", "_ord2", "_n1", "_ord1", "_categoria"]
    if agregacion == "Promedio":
        graf = base.groupby(claves, as_index=False, dropna=False)["_y"].mean()
    elif agregacion == "Suma":
        graf = base.groupby(claves, as_index=False, dropna=False)["_y"].sum()
    elif agregacion == "Mediana":
        graf = base.groupby(claves, as_index=False, dropna=False)["_y"].median()
    else:
        graf = base.groupby(claves, as_index=False, dropna=False)["_y"].count()

    if graf.empty:
        st.warning("No hay datos después de resumir la jerarquía seleccionada.")
        return

    graf = graf.sort_values(["_ord2", "_ord1", "_categoria"], kind="stable")
    if seleccion_n2 == "Sin segundo nivel":
        graf["_x_label"] = graf["_n1"].astype(str)
    else:
        graf["_x_label"] = (
            graf["_n2"].astype(str) + " | " + graf["_n1"].astype(str)
        )

    pares_orden = (
        graf[["_ord2", "_ord1", "_x_label"]]
        .drop_duplicates()
        .sort_values(["_ord2", "_ord1"], kind="stable")
        .reset_index(drop=True)
    )
    pares_orden["_x_order"] = np.arange(len(pares_orden))
    graf = graf.merge(pares_orden, on=["_ord2", "_ord1", "_x_label"], how="left")
    graf["_categoria"] = graf["_categoria"].astype(str)

    categorias = graf["_categoria"].drop_duplicates().tolist()
    colores = paleta_azul(len(categorias))
    orden_x = pares_orden["_x_label"].tolist()
    dec_graf = 0 if agregacion == "Conteo" else dec_y

    tooltip = [
        {"field": "_n1", "type": "nominal", "title": mapa_x[seleccion_n1]["nombre"]},
    ]
    if seleccion_n2 != "Sin segundo nivel":
        tooltip.insert(
            0,
            {"field": "_n2", "type": "nominal", "title": mapa_x[seleccion_n2]["nombre"]},
        )
    tooltip.extend([
        {
            "field": "_y", "type": "quantitative",
            "title": f"{agregacion} de {mapa_num[seleccion_y]['nombre']}",
            "format": formato_vega(dec_graf),
        },
        {"field": "_categoria", "type": "nominal", "title": mapa_cat[seleccion_cat]["nombre"]},
    ])

    spec = {
        "mark": {"type": "line", "point": {"filled": True, "size": 55}},
        "encoding": {
            "x": {
                "field": "_x_label", "type": "nominal", "title": titulo_x,
                "sort": orden_x,
                "axis": {"labelAngle": -45, "labelLimit": 180},
            },
            "y": {
                "field": "_y", "type": "quantitative",
                "title": f"{agregacion} de {mapa_num[seleccion_y]['nombre']}",
                "axis": {"format": formato_vega(dec_graf)},
            },
            "color": {
                "field": "_categoria", "type": "nominal",
                "title": mapa_cat[seleccion_cat]["nombre"],
                "scale": {"domain": categorias, "range": colores},
            },
            "detail": {"field": "_categoria", "type": "nominal"},
            "order": {"field": "_x_order", "type": "quantitative"},
            "tooltip": tooltip,
        },
        "height": 500,
    }
    st.vega_lite_chart(graf, spec, use_container_width=True)

    try:
        st.session_state["story_line_png"] = crear_png_lineas(
            graf, titulo_x, f"{agregacion} de {mapa_num[seleccion_y]['nombre']}",
            mapa_cat[seleccion_cat]["nombre"]
        )
        st.session_state["story_line_desc"] = (
            f"Líneas: {agregacion} de {mapa_num[seleccion_y]['nombre']} por {titulo_x}, "
            f"color {mapa_cat[seleccion_cat]['nombre']}"
        )
    except Exception as error:
        st.session_state.pop("story_line_png", None)
        st.caption(f"No fue posible preparar la imagen de líneas para PowerPoint: {error}")



def aplicar_operador(serie: pd.Series, operador: str, umbral, tipo: str = "numérico") -> pd.Series:
    """Aplica comparadores numéricos o igualdad lógica para variables categóricas."""
    if operador == "==":
        if tipo == "numérico":
            x = pd.to_numeric(serie, errors="coerce")
            try:
                valor = float(umbral)
                return pd.Series(
                    np.isclose(x.to_numpy(dtype=float), valor, rtol=0.0, atol=1e-9, equal_nan=False),
                    index=serie.index,
                )
            except (TypeError, ValueError):
                return pd.Series(False, index=serie.index)
        return serie.astype("string") == str(umbral)

    if tipo != "numérico":
        raise ValueError("Las variables categóricas admiten únicamente el operador ==.")

    x = pd.to_numeric(serie, errors="coerce")
    valor = float(umbral)
    if operador == "<":
        return x < valor
    if operador == ">":
        return x > valor
    if operador == "<=":
        return x <= valor
    if operador in (">=", "=>"):
        return x >= valor
    raise ValueError(f"Operador no soportado: {operador}")


def calcular_agregacion(serie: pd.Series, funcion: str):
    """Calcula la agregación solicitada sobre valores numéricos válidos."""
    x = pd.to_numeric(serie, errors="coerce").replace([np.inf, -np.inf], np.nan).dropna()
    if x.empty:
        return None
    if funcion == "Suma":
        return float(x.sum())
    if funcion == "Conteo":
        return int(x.count())
    if funcion == "Promedio":
        return float(x.mean())
    if funcion == "Máximo":
        return float(x.max())
    if funcion == "Mínimo":
        return float(x.min())
    raise ValueError(f"Agregación no soportada: {funcion}")


def formatear_resultado(valor) -> str:
    """Enteros sin decimales; valores fraccionales con un decimal."""
    if valor is None or (isinstance(valor, float) and not np.isfinite(valor)):
        return "Sin resultado"
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return str(valor)
    if np.isclose(numero, round(numero), rtol=0.0, atol=1e-9):
        return f"{int(round(numero)):,}"
    return f"{numero:,.1f}"


def mostrar_logiccal(st, datos: pd.DataFrame, metas: list[dict]) -> None:
    st.subheader("LogicCal — Calculadora lógica de variables")
    st.write(
        "Construya reglas IF/THEN o IF/THEN/ELSE con una, dos o tres condiciones. "
        "Las condiciones pueden usar variables numéricas, categóricas o booleanas. "
        "Las variables categóricas utilizan igualdad lógica (==)."
    )

    condiciones_meta = [
        m for m in metas if m["tipo"] in ("numérico", "categórico", "booleano")
    ]
    numericas = [m for m in metas if m["tipo"] == "numérico"]

    if not condiciones_meta:
        st.info("No se detectaron variables compatibles para utilizar LogicCal.")
        st.session_state["logiccal_resultado_texto"] = "Sin resultado de LogicCal"
        return
    if not numericas:
        st.info("Se requiere al menos una variable numérica para aplicar la agregación THEN/ELSE.")
        st.session_state["logiccal_resultado_texto"] = "Sin resultado de LogicCal"
        return

    mapa_cond = {m["etiqueta"]: m for m in condiciones_meta}
    mapa_num = {m["etiqueta"]: m for m in numericas}

    activo = st.toggle("Activar lógica IF", value=True, key="logic_activo")
    if not activo:
        st.info("LogicCal está desactivado.")
        st.session_state["logiccal_resultado_texto"] = "LogicCal desactivado"
        return

    a, b = st.columns(2)
    modo = a.selectbox(
        "Tipo de lógica", ["IF → THEN", "IF → THEN / ELSE"], key="logic_modo"
    )
    niveles = int(b.selectbox(
        "Número de condiciones", [1, 2, 3], index=0, key="logic_niveles"
    ))

    condiciones = []
    conectores = []
    series_condicion = []
    validas_condicion = pd.Series(True, index=datos.index, dtype=bool)

    st.markdown("#### Condiciones")
    for i in range(niveles):
        if i > 0:
            conector = st.selectbox(
                f"Conector entre condición {i} y {i + 1}", ["Y", "O"],
                key=f"logic_conector_{i}"
            )
            conectores.append(conector)

        c1, c2, c3 = st.columns([2.4, 1.0, 1.8])
        seleccion = c1.selectbox(
            f"Variable {i + 1}", list(mapa_cond), key=f"logic_var_{i}"
        )
        meta = mapa_cond[seleccion]
        tipo = meta["tipo"]
        serie = serie_grafico(datos, meta)

        if tipo == "numérico":
            serie_num = pd.to_numeric(serie, errors="coerce").replace([np.inf, -np.inf], np.nan)
            validos = serie_num.dropna()
            if validos.empty:
                st.warning(f"{meta['nombre']} no contiene valores numéricos válidos.")
                st.session_state["logiccal_resultado_texto"] = "Sin resultado de LogicCal"
                return

            operador = c2.selectbox(
                f"Operador {i + 1}", ["<", ">", "<=", ">=", "=="], key=f"logic_op_{i}"
            )
            dec = decimales_serie(serie_num)
            valor_inicial = float(validos.median())
            umbral = c3.number_input(
                f"Valor {i + 1}",
                value=valor_inicial,
                step=paso_slider(dec),
                format=formato_slider(dec),
                key=f"logic_val_num_{i}_{meta['pos']}",
            )
            mascara_i = aplicar_operador(serie_num, operador, umbral, "numérico").fillna(False)
            validas_condicion &= serie_num.notna()
            valor_regla = formatear_resultado(umbral)
        else:
            serie_cat = limpiar_nulos(serie).astype("string")
            valores = serie_cat.dropna().drop_duplicates().astype(str).tolist()
            if not valores:
                st.warning(f"{meta['nombre']} no contiene categorías válidas.")
                st.session_state["logiccal_resultado_texto"] = "Sin resultado de LogicCal"
                return
            operador = c2.selectbox(
                f"Operador {i + 1}", ["=="], key=f"logic_op_{i}"
            )
            if len(valores) > 1000:
                st.caption(
                    f"{meta['nombre']} tiene {len(valores):,} niveles; se muestran los primeros 1,000."
                )
                valores = valores[:1000]
            umbral = c3.selectbox(
                f"Valor/categoría {i + 1}", valores,
                key=f"logic_val_cat_{i}_{meta['pos']}"
            )
            mascara_i = aplicar_operador(serie_cat, operador, umbral, tipo).fillna(False)
            validas_condicion &= serie_cat.notna()
            valor_regla = f"'{umbral}'"

        series_condicion.append(mascara_i)
        condiciones.append({
            "seleccion": seleccion,
            "nombre": meta["nombre"],
            "tipo": tipo,
            "operador": operador,
            "umbral": umbral,
            "valor_regla": valor_regla,
        })

    mascara = series_condicion[0].copy()
    for i, conector in enumerate(conectores, start=1):
        mascara = (mascara & series_condicion[i]) if conector == "Y" else (mascara | series_condicion[i])
    mascara = mascara & validas_condicion
    mascara_else = (~mascara) & validas_condicion

    st.markdown("#### THEN / ELSE")
    c1, c2 = st.columns(2)
    variable_resultado = c1.selectbox(
        "Aplicar agregación en", list(mapa_num), key="logic_target"
    )
    funcion_then = c2.selectbox(
        "Función THEN", ["Suma", "Conteo", "Promedio", "Máximo", "Mínimo"], key="logic_then"
    )

    funcion_else = None
    if modo == "IF → THEN / ELSE":
        funcion_else = st.selectbox(
            "Función ELSE", ["Suma", "Conteo", "Promedio", "Máximo", "Mínimo"], key="logic_else"
        )

    objetivo = pd.to_numeric(
        serie_grafico(datos, mapa_num[variable_resultado]), errors="coerce"
    ).replace([np.inf, -np.inf], np.nan)

    resultado_then = calcular_agregacion(objetivo[mascara], funcion_then)
    resultado_else = calcular_agregacion(objetivo[mascara_else], funcion_else) if funcion_else else None

    partes_regla = []
    for i, cond in enumerate(condiciones):
        if i > 0:
            partes_regla.append(conectores[i - 1])
        partes_regla.append(
            f"{cond['nombre']} {cond['operador']} {cond['valor_regla']}"
        )
    regla = " ".join(partes_regla)
    nombre_objetivo = mapa_num[variable_resultado]["nombre"]

    if modo == "IF → THEN":
        texto_resultado = (
            f"THEN {funcion_then} de {nombre_objetivo} = {formatear_resultado(resultado_then)}"
        )
        st.code(f"IF {regla} → {texto_resultado}", language="text")
    else:
        texto_resultado = (
            f"THEN {funcion_then} de {nombre_objetivo} = {formatear_resultado(resultado_then)} | "
            f"ELSE {funcion_else} de {nombre_objetivo} = {formatear_resultado(resultado_else)}"
        )
        st.code(f"IF {regla} → {texto_resultado}", language="text")

    st.session_state["logiccal_resultado_texto"] = texto_resultado
    st.session_state["logiccal_regla_texto"] = f"IF {regla}"

    if modo == "IF → THEN":
        m1, m2, m3 = st.columns(3)
        m1.metric("Resultado THEN", formatear_resultado(resultado_then))
        m2.metric("Registros IF", f"{int(mascara.sum()):,}")
        m3.metric("Registros evaluables", f"{int(validas_condicion.sum()):,}")
    else:
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Resultado THEN", formatear_resultado(resultado_then))
        m2.metric("Resultado ELSE", formatear_resultado(resultado_else))
        m3.metric("Registros IF", f"{int(mascara.sum()):,}")
        m4.metric("Registros ELSE", f"{int(mascara_else.sum()):,}")

    with st.expander("Vista previa de registros que cumplen IF"):
        indices = datos.index[mascara][:100]
        st.dataframe(datos.loc[indices], use_container_width=True, hide_index=True)
        if int(mascara.sum()) > 100:
            st.caption("Se muestran los primeros 100 registros que cumplen la regla IF.")

    if modo == "IF → THEN / ELSE":
        with st.expander("Vista previa de registros ELSE"):
            indices = datos.index[mascara_else][:100]
            st.dataframe(datos.loc[indices], use_container_width=True, hide_index=True)
            if int(mascara_else.sum()) > 100:
                st.caption("Se muestran los primeros 100 registros evaluables que no cumplen la regla IF.")


def _agregar_titulo_slide(slide, titulo: str) -> None:
    """Añade un título editable y una línea de acento azul."""
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor

    caja = slide.shapes.add_textbox(Inches(0.75), Inches(0.45), Inches(11.8), Inches(0.65))
    p = caja.text_frame.paragraphs[0]
    p.text = titulo
    p.font.name = "Aptos Display"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x05, 0x25, 0x70)

    linea = slide.shapes.add_shape(1, Inches(0.75), Inches(1.18), Inches(11.8), Inches(0.045))
    linea.fill.solid()
    linea.fill.fore_color.rgb = RGBColor(0x05, 0x25, 0x70)
    linea.line.fill.background()


def _agregar_texto_slide(slide, texto: str, left: float, top: float,
                         width: float, height: float, tamano: int = 20) -> None:
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor

    caja = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = caja.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = texto or ""
    p.font.name = "Aptos"
    p.font.size = Pt(tamano)
    p.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)


def crear_powerpoint_storytelling(anzuelo: str, contexto1: str, contexto2: str,
                                  contexto3: str, momento: str, resultado_logic: str,
                                  imagen_box: bytes | None, imagen_bar: bytes | None,
                                  imagen_line: bytes | None) -> bytes:
    """Crea un PowerPoint de 5 slides editable; los gráficos se insertan como imágenes."""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # A — Anzuelo
    slide = prs.slides.add_slide(blank)
    _agregar_titulo_slide(slide, "A. Anzuelo")
    _agregar_texto_slide(slide, anzuelo, 1.0, 2.0, 11.2, 3.3, 28)

    # B/C/D — Contextos + gráficos
    for titulo, texto, imagen in [
        ("B. Contexto 1 — Boxplot", contexto1, imagen_box),
        ("C. Contexto 2 — Barras", contexto2, imagen_bar),
        ("D. Contexto 3 — Líneas", contexto3, imagen_line),
    ]:
        slide = prs.slides.add_slide(blank)
        _agregar_titulo_slide(slide, titulo)
        _agregar_texto_slide(slide, texto, 0.8, 1.55, 4.0, 4.9, 18)
        if imagen:
            try:
                slide.shapes.add_picture(BytesIO(imagen), Inches(5.0), Inches(1.55), width=Inches(7.55))
            except Exception:
                _agregar_texto_slide(slide, "No se pudo insertar la imagen del gráfico.", 5.1, 2.4, 6.7, 1.0, 18)
        else:
            _agregar_texto_slide(slide, "Sin imagen de gráfico disponible.", 5.1, 2.4, 6.7, 1.0, 18)

    # E — Momento ajá
    slide = prs.slides.add_slide(blank)
    _agregar_titulo_slide(slide, "E. Momento ajá")
    _agregar_texto_slide(slide, momento, 0.8, 1.55, 5.0, 4.7, 20)

    circulo = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.0), Inches(1.55), Inches(4.5), Inches(4.5))
    circulo.fill.solid()
    circulo.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    circulo.line.color.rgb = RGBColor(0x05, 0x25, 0x70)
    circulo.line.width = Pt(4)
    tf = circulo.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = resultado_logic or "Sin resultado de LogicCal"
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Aptos Display"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x05, 0x25, 0x70)

    salida = BytesIO()
    prs.save(salida)
    salida.seek(0)
    return salida.getvalue()


def asegurar_boxplot_storytelling(st, datos: pd.DataFrame, metas: list[dict]) -> bytes | None:
    """Garantiza que Data Storytelling tenga una imagen del boxplot actual."""
    existente = st.session_state.get("story_box_png")
    if existente:
        return existente

    numericas = [m for m in metas if m["tipo"] == "numérico"]
    categoricas = [m for m in metas if m["tipo"] in ("categórico", "booleano")]
    if not numericas:
        return None

    mapa_num = {m["etiqueta"]: m for m in numericas}
    mapa_cat = {m["etiqueta"]: m for m in categoricas}

    seleccion_num = st.session_state.get("box_num")
    if seleccion_num not in mapa_num:
        seleccion_num = next(iter(mapa_num))

    seleccion_cat = st.session_state.get("box_cat", "Sin agrupación")
    if seleccion_cat != "Sin agrupación" and seleccion_cat not in mapa_cat:
        seleccion_cat = "Sin agrupación"

    y = pd.to_numeric(serie_grafico(datos, mapa_num[seleccion_num]), errors="coerce")
    base = pd.DataFrame({"_valor": y})
    if seleccion_cat == "Sin agrupación":
        base["_categoria"] = "Todos"
        titulo_cat = "Grupo"
    else:
        cat = serie_grafico(datos, mapa_cat[seleccion_cat])
        base["_categoria"] = cat.astype("string")
        titulo_cat = mapa_cat[seleccion_cat]["nombre"]

    base = base.replace([np.inf, -np.inf], np.nan).dropna(subset=["_valor", "_categoria"])
    if base.empty:
        return None

    try:
        imagen = crear_png_boxplot(
            base, mapa_num[seleccion_num]["nombre"], titulo_cat
        )
        st.session_state["story_box_png"] = imagen
        st.session_state["story_box_desc"] = (
            f"Boxplot horizontal: {mapa_num[seleccion_num]['nombre']} por "
            f"{titulo_cat if seleccion_cat != 'Sin agrupación' else 'Todos'}"
        )
        return imagen
    except Exception:
        return None

def mostrar_storytelling(st, datos: pd.DataFrame, metas: list[dict]) -> None:
    st.subheader("Data Storytelling")
    st.write(
        "Construya una narrativa de cinco momentos. El boxplot actual se adjunta automáticamente. "
        "Los otros gráficos también pueden incorporarse automáticamente al PowerPoint o sustituirse "
        "por una imagen PNG/JPG."
    )

    # Regeneración de respaldo: aunque el usuario no haya abierto el tab Boxplot
    # en la sesión actual, se intenta construir la imagen con la selección vigente.
    imagen_box_auto = asegurar_boxplot_storytelling(st, datos, metas)
    imagen_box = imagen_box_auto

    anzuelo = st.text_area(
        "Anzuelo", height=110, key="story_anzuelo",
        placeholder="Escriba la idea que capta la atención de la audiencia..."
    )

    st.markdown("#### Contexto 1 — Boxplot")
    contexto1 = st.text_area("Contexto 1", height=110, key="story_contexto1")
    if imagen_box_auto:
        st.success("Boxplot actual adjuntado automáticamente.")
        st.image(
            imagen_box_auto,
            caption=st.session_state.get(
                "story_box_desc", "Boxplot actual que se exportará en Contexto 1"
            ),
            width=620,
        )
        imagen_box = imagen_box_auto
    else:
        st.warning("No fue posible generar automáticamente el boxplot. Configure el tab Boxplot.")

    # Permite reemplazar manualmente solo si el usuario lo desea.
    upload_box = st.file_uploader(
        "Opcional: sustituir el Boxplot automático por otra imagen",
        type=["png", "jpg", "jpeg"], key="story_upload_box"
    )
    if upload_box:
        imagen_box = upload_box.getvalue()
        st.image(imagen_box, caption="Imagen manual que sustituirá al Boxplot", width=620)

    st.markdown("#### Contexto 2 — Barras")
    contexto2 = st.text_area("Contexto 2", height=110, key="story_contexto2")
    upload_bar = st.file_uploader(
        "Imagen opcional para Contexto 2 (si se omite, se usa Barras actual)",
        type=["png", "jpg", "jpeg"], key="story_upload_bar"
    )
    imagen_bar = upload_bar.getvalue() if upload_bar else st.session_state.get("story_bar_png")
    if imagen_bar:
        st.image(imagen_bar, caption="Imagen que se exportará en Contexto 2", width=620)
    else:
        st.info("Configure primero el tab Barras o adjunte una imagen.")

    st.markdown("#### Contexto 3 — Líneas")
    contexto3 = st.text_area("Contexto 3", height=110, key="story_contexto3")
    upload_line = st.file_uploader(
        "Imagen opcional para Contexto 3 (si se omite, se usa Líneas actual)",
        type=["png", "jpg", "jpeg"], key="story_upload_line"
    )
    imagen_line = upload_line.getvalue() if upload_line else st.session_state.get("story_line_png")
    if imagen_line:
        st.image(imagen_line, caption="Imagen que se exportará en Contexto 3", width=620)
    else:
        st.info("Configure primero el tab Líneas o adjunte una imagen.")

    st.markdown("#### Momento ajá")
    momento = st.text_area(
        "Momento ajá", height=110, key="story_momento",
        placeholder="Explique por qué el resultado de LogicCal es importante..."
    )
    resultado_logic = st.session_state.get("logiccal_resultado_texto", "Sin resultado de LogicCal")

    from html import escape
    st.markdown(
        f"""
        <div style="display:flex;justify-content:center;margin:18px 0 26px 0;">
          <div style="width:330px;height:330px;border:6px solid #052570;border-radius:50%;
                      background:#FFFFFF;display:flex;align-items:center;justify-content:center;
                      text-align:center;padding:28px;box-sizing:border-box;color:#052570;
                      font-size:25px;font-weight:700;line-height:1.25;">
            {escape(resultado_logic)}
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    try:
        pptx_bytes = crear_powerpoint_storytelling(
            anzuelo, contexto1, contexto2, contexto3, momento, resultado_logic,
            imagen_box, imagen_bar, imagen_line,
        )
        st.download_button(
            "⬇️ Exportar Data Storytelling a PowerPoint",
            data=pptx_bytes,
            file_name="EDA_Data_Storytelling.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="story_download_pptx",
        )
        st.caption(
            "El PowerPoint contiene cinco slides editables. Los textos, títulos y círculo del Momento ajá "
            "son editables; los gráficos se insertan como imágenes."
        )
    except ModuleNotFoundError as error:
        st.error(
            f"Falta una dependencia para exportar PowerPoint: {error}. "
            "Instale python-pptx y matplotlib en el mismo entorno virtual."
        )
    except Exception as error:
        st.error(f"No fue posible generar el PowerPoint: {error}")


def main() -> None:
    import streamlit as st

    st.set_page_config(page_title="EDA | Exploración de datos", page_icon="📊", layout="wide")
    st.title("📊 Exploratory Data Analysis (EDA)")
    st.write(
        "Cargue un archivo CSV o Excel, explore hasta 50 columnas, genere visualizaciones, "
        "construya reglas con LogicCal y prepare una narrativa de Data Storytelling."
    )

    entrada = st.file_uploader("Seleccione su archivo", type=["csv", "xlsx", "xls"])
    if entrada is None:
        st.info("Seleccione un archivo para comenzar.")
        return

    contenido = entrada.getvalue()
    try:
        hojas = hojas_excel(contenido, entrada.name)
        hoja = st.selectbox("Hoja de Excel", hojas, index=0) if hojas else None
        datos = importar_archivo(contenido, entrada.name, hoja)
        if datos.shape[1] == 0:
            st.error("El archivo no contiene columnas.")
            return
        if datos.shape[1] > MAX_COLUMNAS:
            st.warning(
                f"El archivo tiene {datos.shape[1]} columnas. Se analizan y grafican solo las primeras "
                f"{MAX_COLUMNAS}; se conserva el archivo original en la exportación XLSX."
            )
        eda = analizar(datos)
        metas = metadatos_graficos(datos)
    except Exception as error:
        st.error(f"No fue posible procesar el archivo: {error}")
        return

    a, b, c = st.columns(3)
    a.metric("Filas", f"{len(datos):,}")
    b.metric("Columnas originales", datos.shape[1])
    c.metric("Columnas analizadas", len(eda))

    tab_eda, tab_box, tab_barras, tab_lineas, tab_logic, tab_story = st.tabs(
        ["📋 EDA", "📦 Boxplot", "📊 Barras", "📈 Líneas", "🧮 LogicCal", "🎤 Data Storytelling"]
    )

    with tab_eda:
        st.subheader("Resultados del EDA")
        st.dataframe(eda, use_container_width=True, hide_index=True)
        with st.expander("Vista previa del archivo original"):
            st.dataframe(datos.head(100), use_container_width=True)
        with st.expander("Criterios del análisis"):
            st.write(
                "Conteo = total de filas importadas, incluidos los valores nulos. "
                "DS = desviación estándar muestral (ddof=1). "
                "Las categorías se enumeran únicamente cuando hay de 1 a 5 valores distintos. "
                "Las fechas ambiguas y códigos se tratan conservadoramente."
            )

        try:
            archivo = crear_excel(contenido, entrada.name, datos, eda, hoja)
            nombre_base = entrada.name.rsplit(".", 1)[0]
            st.download_button(
                "⬇️ Descargar Excel con hoja EDA",
                data=archivo,
                file_name=f"{nombre_base}_EDA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as error:
            st.error(f"El EDA se mostró correctamente, pero no se pudo exportar el Excel: {error}")

    with tab_box:
        mostrar_boxplot(st, datos, metas)

    with tab_barras:
        mostrar_barras(st, datos, metas)

    with tab_lineas:
        mostrar_lineas(st, datos, metas)

    with tab_logic:
        mostrar_logiccal(st, datos, metas)

    with tab_story:
        mostrar_storytelling(st, datos, metas)


if __name__ == "__main__":
    main()
