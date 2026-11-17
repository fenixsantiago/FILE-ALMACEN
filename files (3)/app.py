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

Ejecutar:  streamlit run app.py
"""
from pathlib import Path

import pandas as pd
import streamlit as st

st.set_page_config(page_title="ALMACEN CLIE", page_icon="📊", layout="wide")

CARPETA = Path(__file__).resolve().parent
DATOS_POR_DEFECTO = CARPETA / "datos_hoja1.csv.gz"   # copia rápida (se crea sola)

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
# Formato es-CO: 1.234.567,89
# ----------------------------------------------------------------------------
def fmt(x: float, dec: int = 2) -> str:
    s = f"{x:,.{dec}f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


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

with st.sidebar:
    # Extra (no existe en el .pbix): el reporte suma todos los periodos.
    periodos = sorted(df["PERIODO"].dropna().unique())
    sel_periodo = st.multiselect("PERIODO (opcional)", periodos,
                                 help="El reporte original suma todos los periodos.")
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

f = df
if st.session_state.get("f_nit"):
    f = f[f["NIT"].isin(st.session_state["f_nit"])]
if st.session_state.get("f_nombre"):
    f = f[f["NOMBRE"].isin(st.session_state["f_nombre"])]

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