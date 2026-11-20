"""
ALMACEN_CLIE — réplica en Streamlit del reporte de Power BI (ALMACEN_CLIE.pbix)

Reporte original (1 página):
  - Segmentador NIT            (búsqueda + selección múltiple)
  - Segmentador NOMBRE         (búsqueda + selección múltiple)
  - Tarjeta  "cantidad total"  = SUM(Hoja1[cantidad total])
  - Tarjeta  "VALOR TOTAL"     = SUM(Hoja1[TOTAL])
  - Matriz   filas=GRUPO, valores = Suma cantidad total | Suma TOTAL | venta total
             ordenada por "venta total" descendente
  - Medida   venta total = DIVIDE(SUM(Hoja1[TOTAL]), SUM(Hoja1[cantidad total]), 0)

Extra (no existe en el .pbix):
  - Comparativo por año (ej. 2025 vs 2024, o 2026 vs 2025), por GRUPO y por mes.

Ejecutar:  streamlit run app.py
"""
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="ALMACEN CLIE", page_icon="📊", layout="wide")

CARPETA = Path(__file__).resolve().parent
DATOS_POR_DEFECTO = CARPETA / "datos_hoja1.csv.gz"   # copia rápida (se crea sola)

MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
         "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def buscar_excel():
    """Busca ALMACER*.xlsx en la carpeta de la app, la carpeta padre, la carpeta
    desde donde se ejecutó y la ruta que usaba el reporte de Power BI."""
    sitios = [
        CARPETA, CARPETA.parent, Path.cwd(), Path.cwd().parent,
        Path.home() / "OneDrive - AXXECOL S.A" / "Documentos" / "powerbi co",
        Path.home() / "OneDrive - AXXECOL S.A" / "Documentos",
        Path.home() / "Documents", Path.home() / "Downloads", Path.home() / "Desktop",
    ]
    for sitio in sitios:
        try:
            for p in sorted(sitio.glob("ALMACER*.xlsx")):
                if not p.name.startswith("~$"):
                    return p
        except OSError:
            continue
    return None


# ----------------------------------------------------------------------------
# Carga de datos (equivale a la consulta de Power Query)
# ----------------------------------------------------------------------------
def limpiar(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza encabezados como hace Power Query (quita espacios sobrantes)."""
    df = df.copy()
    cols = [" ".join(str(c).split()) for c in df.columns]
    # En el Excel hay dos columnas "TOTAL": la 1.ª es el valor y la 2.ª la cantidad.
    # Se identifican por posición para no depender de cómo las nombre cada lector.
    idx = [i for i, c in enumerate(cols) if c.startswith("TOTAL")]
    if len(idx) >= 2:
        cols[idx[0]], cols[idx[1]] = "TOTAL", "cantidad total"
    df.columns = cols
    df = df.rename(columns={c: c.upper() for c in cols
                            if c.upper() in ("NIT", "NOMBRE", "GRUPO", "PERIODO")})
    faltan = [c for c in ("NIT", "NOMBRE", "GRUPO") if c not in df.columns]
    if faltan:
        raise ValueError(f"Faltan columnas {faltan}. Columnas encontradas: {list(df.columns)[:15]}")
    # Descartar filas vacías (la hoja trae ~1 millón de filas con formato)
    df = df.dropna(subset=["NOMBRE", "GRUPO"], how="all").dropna(how="all")
    df["NIT"] = df["NIT"].astype(str).str.replace(r"\.0$", "", regex=True)
    for c in ("TOTAL", "cantidad total"):
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    for c in ("NOMBRE", "GRUPO"):
        df[c] = df[c].astype(str).str.strip()
    return df.reset_index(drop=True)


def _norm(c) -> str:
    return " ".join(str(c).split()).upper()


def leer_excel(archivo) -> pd.DataFrame:
    """Busca, en todas las hojas, la fila de encabezados que contenga NOMBRE y
    GRUPO (aunque no sea la fila 1 ni la hoja se llame «Hoja1») y lee desde ahí."""
    def rebobinar():
        if hasattr(archivo, "seek"):
            archivo.seek(0)

    rebobinar()
    xl = pd.ExcelFile(archivo)
    nombres = sorted(xl.sheet_names, key=lambda n: n.strip().lower() != "hoja1")
    vistas = {}
    for hoja in nombres:
        muestra = xl.parse(hoja, header=None, nrows=30)
        for i in range(len(muestra)):
            fila = {_norm(v) for v in muestra.iloc[i].tolist()}
            if any(x.startswith("NOMBRE") for x in fila) and "GRUPO" in fila:
                rebobinar()
                return pd.read_excel(archivo, sheet_name=hoja, header=i)
        vistas[hoja] = [_norm(v) for v in muestra.iloc[0].tolist()] if len(muestra) else []
    detalle = "; ".join(f"«{h}»: {c[:8]}" for h, c in vistas.items())
    raise ValueError("No encontré columnas NOMBRE y GRUPO en ninguna hoja. "
                     f"Primeras columnas por hoja → {detalle}")


def hay_datos_locales() -> bool:
    return DATOS_POR_DEFECTO.exists() or buscar_excel() is not None


@st.cache_data(show_spinner="Cargando datos…")
def cargar_por_defecto() -> pd.DataFrame:
    if DATOS_POR_DEFECTO.exists():
        return limpiar(pd.read_csv(DATOS_POR_DEFECTO, low_memory=False))
    # Primera vez: leer el Excel y guardar una copia rápida para las siguientes
    df = limpiar(leer_excel(buscar_excel()))
    try:
        df.to_csv(DATOS_POR_DEFECTO, index=False, compression="gzip")
    except OSError:
        pass
    return df


@st.cache_data(show_spinner="Leyendo Excel…")
def cargar_excel(archivo) -> pd.DataFrame:
    return limpiar(leer_excel(archivo))


# ----------------------------------------------------------------------------
# Año y mes a partir de PERIODO
# ----------------------------------------------------------------------------
def anio_mes(periodo: pd.Series) -> tuple[pd.Series, pd.Series]:
    """Extrae AÑO y MES de PERIODO. Entiende formatos como:
    2025, 202503, 2025-03, 2025/03, 2025-03-31 00:00:00, 03/2025, 3-2025."""
    s = periodo.astype(str).str.strip()
    # AAAA + (separador opcional) + MM
    a1 = s.str.extract(r"(20\d{2})[-/ .]?(\d{2})(?!\d)")
    # MM/AAAA
    a2 = s.str.extract(r"(?<!\d)(\d{1,2})[-/ .](20\d{2})")
    # Solo AAAA
    a3 = s.str.extract(r"(20\d{2})")

    anio = pd.to_numeric(a1[0], errors="coerce")
    mes = pd.to_numeric(a1[1], errors="coerce")

    anio = anio.fillna(pd.to_numeric(a2[1], errors="coerce"))
    mes = mes.fillna(pd.to_numeric(a2[0], errors="coerce"))
    anio = anio.fillna(pd.to_numeric(a3[0], errors="coerce"))

    mes = mes.where((mes >= 1) & (mes <= 12))
    return anio, mes


# ----------------------------------------------------------------------------
# Formato es-CO: 1.234.567,89
# ----------------------------------------------------------------------------
def fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def fmt_pct(x) -> str:
    """Variación porcentual con flecha: ▲ 12,3%  /  ▼ 5,0%"""
    if pd.isna(x) or np.isinf(x):
        return "—"
    flecha = "▲" if x > 0 else "▼" if x < 0 else "■"
    return f"{flecha} {fmt(abs(x), 1)}%"


def fmt_delta(x) -> str:
    """Texto para el delta de st.metric (+12,3% / -5,0%)."""
    if x is None or pd.isna(x) or np.isinf(x):
        return None
    signo = "+" if x > 0 else "-" if x < 0 else ""
    return f"{signo}{fmt(abs(x), 1)}%"


def pct(a, b):
    """Variación % de a contra b (b = base de comparación)."""
    if isinstance(a, pd.Series):
        return (a - b) / b.replace(0, np.nan) * 100
    return (a - b) / b * 100 if b else np.nan


# ----------------------------------------------------------------------------
# Lógica del reporte
# ----------------------------------------------------------------------------
def venta_total(total: float, cantidad: float) -> float:
    """DIVIDE(SUM(TOTAL), SUM(cantidad total), 0)"""
    return total / cantidad if cantidad else 0.0


def matriz_por_grupo(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby("GRUPO", as_index=False)
        .agg(**{"Suma de cantidad total": ("cantidad total", "sum"),
                "Suma de TOTAL": ("TOTAL", "sum")})
    )
    g["venta total"] = [
        venta_total(t, c) for t, c in zip(g["Suma de TOTAL"], g["Suma de cantidad total"])
    ]
    return g.sort_values("venta total", ascending=False).reset_index(drop=True)


def totales_anio(df: pd.DataFrame, anio: int) -> tuple[float, float, float]:
    d = df[df["ANIO"] == anio]
    cant = float(d["cantidad total"].sum())
    valor = float(d["TOTAL"].sum())
    return cant, valor, venta_total(valor, cant)


def comparar_grupos(df: pd.DataFrame, a: int, b: int) -> pd.DataFrame:
    """Por GRUPO: cantidad, TOTAL y venta total de los años a y b, con variaciones (a vs b)."""
    def por_anio(anio):
        return (df[df["ANIO"] == anio].groupby("GRUPO")
                .agg(cant=("cantidad total", "sum"), total=("TOTAL", "sum")))

    ga, gb = por_anio(a), por_anio(b)
    m = ga.join(gb, how="outer", lsuffix="_a", rsuffix="_b").fillna(0.0)
    m["venta_a"] = np.where(m["cant_a"] != 0, m["total_a"] / m["cant_a"].replace(0, np.nan), 0.0)
    m["venta_b"] = np.where(m["cant_b"] != 0, m["total_b"] / m["cant_b"].replace(0, np.nan), 0.0)
    m["d_cant_pct"] = pct(m["cant_a"], m["cant_b"])
    m["d_total"] = m["total_a"] - m["total_b"]
    m["d_total_pct"] = pct(m["total_a"], m["total_b"])
    m["d_venta_pct"] = pct(m["venta_a"], m["venta_b"])
    return m.sort_values("total_a", ascending=False).reset_index()


def segmentar(d: pd.DataFrame) -> pd.DataFrame:
    """Aplica los segmentadores NIT y NOMBRE."""
    if st.session_state.get("f_nit"):
        d = d[d["NIT"].isin(st.session_state["f_nit"])]
    if st.session_state.get("f_nombre"):
        d = d[d["NOMBRE"].isin(st.session_state["f_nombre"])]
    return d


# ----------------------------------------------------------------------------
# Interfaz
# ----------------------------------------------------------------------------
st.title("📊 ALMACEN CLIE")

with st.sidebar:
    st.header("Datos")
    subido = st.file_uploader("Actualizar con Excel (hoja «Hoja1»)", type=["xlsx", "xlsm"])
    st.caption("Sin archivo, la app busca ALMACER.xlsx sola.")

if subido:
    try:
        df = cargar_excel(subido)
    except ValueError as e:
        st.error(f"El archivo «{subido.name}» no tiene el formato esperado. {e}")
        st.stop()
elif hay_datos_locales():
    df = cargar_por_defecto()
else:
    st.warning("No encontré **ALMACER.xlsx** en ninguna carpeta habitual. "
               "Súbelo en la barra lateral (flecha «>» arriba a la izquierda).")
    st.stop()

# Año y mes derivados de PERIODO (para el comparativo)
if "PERIODO" in df.columns:
    df = df.copy()
    df["ANIO"], df["MES"] = anio_mes(df["PERIODO"])

df_total = df   # todos los periodos (el comparativo usa este)

with st.sidebar:
    # Extra (no existe en el .pbix): el reporte suma todos los periodos.
    periodos = sorted(df["PERIODO"].dropna().unique())
    sel_periodo = st.multiselect("PERIODO (opcional)", periodos,
                                 help="Filtra las tarjetas y la matriz. "
                                      "El comparativo por año ignora este filtro.")
if sel_periodo:
    df = df[df["PERIODO"].isin(sel_periodo)]

# --- Segmentadores con filtrado cruzado (como Power BI) ----------------------
sel_nit = st.session_state.get("f_nit", [])
sel_nom = st.session_state.get("f_nombre", [])

base_nit = df[df["NOMBRE"].isin(sel_nom)] if sel_nom else df
base_nom = df[df["NIT"].isin(sel_nit)] if sel_nit else df

c1, c2 = st.columns(2)
c1.multiselect("NIT", sorted(base_nit["NIT"].unique()), key="f_nit",
               placeholder="Buscar / seleccionar NIT")
c2.multiselect("NOMBRE", sorted(base_nom["NOMBRE"].unique()), key="f_nombre",
               placeholder="Buscar / seleccionar cliente")

f = segmentar(df)

# --- Tarjetas -----------------------------------------------------------------
cant = float(f["cantidad total"].sum())
valor = float(f["TOTAL"].sum())

k1, k2 = st.columns(2)
k1.metric("Cantidad total", fmt(cant))
k2.metric("VALOR TOTAL", fmt(valor))

# --- Matriz -------------------------------------------------------------------
st.subheader("Por GRUPO")
m = matriz_por_grupo(f)

if m.empty:
    st.info("No hay datos para los filtros seleccionados.")
else:
    st.dataframe(
        m.style.format({
            "Suma de cantidad total": lambda v: fmt(v),
            "Suma de TOTAL": lambda v: fmt(v),
            "venta total": lambda v: fmt(v),
        }),
        width="stretch", hide_index=True,
        height=min(40 + 35 * len(m), 520),
    )
    t1, t2, t3 = st.columns(3)
    t1.metric("Total · cantidad", fmt(cant))
    t2.metric("Total · TOTAL", fmt(valor))
    t3.metric("Total · venta total", fmt(venta_total(valor, cant)))

    st.download_button("⬇️ Descargar tabla (CSV)",
                       m.to_csv(index=False).encode("utf-8-sig"),
                       "almacen_clie_por_grupo.csv", "text/csv")

# ============================================================================
# COMPARATIVO POR AÑO
# ============================================================================
st.divider()
st.header("📈 Comparativo por año")

f_cmp = segmentar(df_total)   # respeta NIT / NOMBRE, ignora el filtro de PERIODO

if "ANIO" not in f_cmp.columns or f_cmp["ANIO"].dropna().empty:
    st.info("No pude identificar el año en la columna PERIODO, "
            "así que no se puede armar el comparativo.")
else:
    anios = sorted({int(x) for x in f_cmp["ANIO"].dropna().unique()})

    if len(anios) < 2:
        st.info(f"Solo hay datos de un año ({anios[0]}). "
                "Se necesitan al menos dos años para comparar.")
    else:
        s1, s2, _ = st.columns([1, 1, 2])
        anio_a = s1.selectbox("Año a analizar", anios, index=len(anios) - 1)
        opciones_b = [x for x in anios if x != anio_a]
        idx_b = opciones_b.index(anio_a - 1) if (anio_a - 1) in opciones_b else len(opciones_b) - 1
        anio_b = s2.selectbox("Comparar contra", opciones_b, index=idx_b)

        st.caption(f"Se compara **{anio_a}** contra **{anio_b}**. Las variaciones son de "
                   f"{anio_a} respecto a {anio_b}. Respeta los filtros NIT y NOMBRE.")

        ca, va, vta = totales_anio(f_cmp, anio_a)
        cb, vb, vtb = totales_anio(f_cmp, anio_b)

        # --- Tarjetas comparativas ---
        k1, k2, k3 = st.columns(3)
        k1.metric(f"Cantidad total {anio_a}", fmt(ca), fmt_delta(pct(ca, cb)))
        k1.caption(f"{anio_b}: {fmt(cb)}")
        k2.metric(f"VALOR TOTAL {anio_a}", fmt(va), fmt_delta(pct(va, vb)))
        k2.caption(f"{anio_b}: {fmt(vb)}")
        k3.metric(f"Venta total {anio_a}", fmt(vta), fmt_delta(pct(vta, vtb)))
        k3.caption(f"{anio_b}: {fmt(vtb)}")

        # --- Tabla por GRUPO ---
        st.subheader("Por GRUPO")
        cmp_g = comparar_grupos(f_cmp, anio_a, anio_b)

        if cmp_g.empty:
            st.info("No hay datos para los filtros seleccionados.")
        else:
            vista = pd.DataFrame({
                "GRUPO": cmp_g["GRUPO"],
                f"Cantidad {anio_a}": cmp_g["cant_a"].map(fmt),
                f"Cantidad {anio_b}": cmp_g["cant_b"].map(fmt),
                "Δ Cantidad %": cmp_g["d_cant_pct"].map(fmt_pct),
                f"TOTAL {anio_a}": cmp_g["total_a"].map(fmt),
                f"TOTAL {anio_b}": cmp_g["total_b"].map(fmt),
                "Δ TOTAL": cmp_g["d_total"].map(fmt),
                "Δ TOTAL %": cmp_g["d_total_pct"].map(fmt_pct),
                f"Venta total {anio_a}": cmp_g["venta_a"].map(fmt),
                f"Venta total {anio_b}": cmp_g["venta_b"].map(fmt),
                "Δ Venta total %": cmp_g["d_venta_pct"].map(fmt_pct),
            })
            st.dataframe(vista, width="stretch", hide_index=True,
                         height=min(40 + 35 * len(vista), 520))

            export = cmp_g.rename(columns={
                "cant_a": f"Cantidad {anio_a}", "cant_b": f"Cantidad {anio_b}",
                "total_a": f"TOTAL {anio_a}", "total_b": f"TOTAL {anio_b}",
                "venta_a": f"Venta total {anio_a}", "venta_b": f"Venta total {anio_b}",
                "d_cant_pct": "Δ Cantidad %", "d_total": "Δ TOTAL",
                "d_total_pct": "Δ TOTAL %", "d_venta_pct": "Δ Venta total %",
            })
            st.download_button("⬇️ Descargar comparativo (CSV)",
                               export.to_csv(index=False).encode("utf-8-sig"),
                               f"comparativo_{anio_a}_vs_{anio_b}.csv", "text/csv")

        # --- Comparativo mes a mes ---
        if f_cmp["MES"].notna().any():
            st.subheader("Mes a mes")
            medida = st.radio("Medida", ["TOTAL (valor)", "cantidad total"],
                              horizontal=True, key="medida_mes")
            col = "TOTAL" if medida.startswith("TOTAL") else "cantidad total"

            dm = f_cmp[f_cmp["ANIO"].isin([anio_a, anio_b]) & f_cmp["MES"].notna()]
            pv = (dm.pivot_table(index="MES", columns="ANIO", values=col,
                                 aggfunc="sum", fill_value=0.0)
                  .reindex(range(1, 13), fill_value=0.0))
            for y in (anio_a, anio_b):
                if y not in pv.columns:
                    pv[y] = 0.0

            graf = pv[[anio_b, anio_a]].copy()
            graf.columns = [str(anio_b), str(anio_a)]
            graf.index = MESES
            st.bar_chart(graf)

            tabla_mes = pd.DataFrame({
                "Mes": MESES,
                str(anio_a): pv[anio_a].map(fmt).values,
                str(anio_b): pv[anio_b].map(fmt).values,
                "Δ": (pv[anio_a] - pv[anio_b]).map(fmt).values,
                "Δ %": pct(pv[anio_a], pv[anio_b]).map(fmt_pct).values,
            })
            st.dataframe(tabla_mes, width="stretch", hide_index=True)