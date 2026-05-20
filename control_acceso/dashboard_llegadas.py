import pandas as pd
import streamlit as st

# =========================
# CONFIGURACIÓN DE PÁGINA
# =========================
st.set_page_config(
    page_title="Dashboard Control de Acceso CPCS",
    page_icon="📊",
    layout="wide"
)

# =========================
# ESTILOS CSS
# =========================
st.markdown("""
<style>
    .main-title {
        font-size: 34px;
        font-weight: 800;
        color: #1A74A8;
        margin-bottom: 0;
    }
    .subtitle {
        font-size: 16px;
        color: #555;
        margin-top: 0;
    }
    .card {
        background: white;
        padding: 22px;
        border-radius: 18px;
        box-shadow: 0 4px 14px rgba(0,0,0,0.08);
        border-left: 7px solid #1A74A8;
        min-height: 120px;
    }
    .card-yellow {
        border-left: 7px solid rgb(252, 255, 76);
    }
    .card-title {
        font-size: 15px;
        font-weight: 600;
        color: #555;
    }
    .card-value {
        font-size: 34px;
        font-weight: 800;
        color: #111;
    }
    .section-title {
        font-size: 24px;
        font-weight: 700;
        margin-top: 30px;
        color: #1A74A8;
    }
</style>
""", unsafe_allow_html=True)

# =========================
# CARGAR DATOS
# =========================
@st.cache_data
def cargar_datos():
    # Tu archivo tiene el encabezado real en la fila 5 porque se exportó con startrow=4
    df = pd.read_excel(
        "./reportellegadasAdministrativos.xlsx",
        sheet_name="Reporte General",
        header=4
    )

    # Limpiar nombres de columnas
    df.columns = df.columns.str.strip()

    # Asegurar columnas numéricas
    columnas_numericas = [
        "llegadas_tarde",
        "llegadas_a_tiempo",
        "sin_marcar",
        "total_registros",
        "cumplimiento"
    ]

    for col in columnas_numericas:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df

try:
    df = cargar_datos()
except FileNotFoundError:
    st.error("No se encontró el archivo reportellegadasAdministrativos.xlsx. Debe estar en la misma carpeta que este dashboard.py")
    st.stop()

# =========================
# TÍTULO
# =========================
st.markdown('<h1 class="main-title">Dashboard Control de Acceso CPCS</h1>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Reporte de llegadas tarde, llegadas a tiempo y registros sin marcar.</p>', unsafe_allow_html=True)

# =========================
# FILTRO
# =========================
st.sidebar.header("Filtros")

empleados = ["Todos"] + sorted(df["nombre"].dropna().unique().tolist())
empleado_seleccionado = st.sidebar.selectbox("Empleado", empleados)

if empleado_seleccionado != "Todos":
    df_filtrado = df[df["nombre"] == empleado_seleccionado]
else:
    df_filtrado = df.copy()

# =========================
# KPIs
# =========================
total_empleados = df_filtrado["nombre"].nunique()
total_tardanzas = int(df_filtrado["llegadas_tarde"].sum())
total_a_tiempo = int(df_filtrado["llegadas_a_tiempo"].sum())
total_sin_marcar = int(df_filtrado["sin_marcar"].sum())
total_registros = int(df_filtrado["total_registros"].sum())

if total_registros > 0:
    cumplimiento_general = round((total_a_tiempo / total_registros) * 100, 2)
else:
    cumplimiento_general = 0

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown(f"""
    <div class="card">
        <div class="card-title">Total empleados</div>
        <div class="card-value">{total_empleados}</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="card card-yellow">
        <div class="card-title">Llegadas tarde</div>
        <div class="card-value">{total_tardanzas}</div>
    </div>
    """, unsafe_allow_html=True)

with col3:
    st.markdown(f"""
    <div class="card">
        <div class="card-title">Llegadas a tiempo</div>
        <div class="card-value">{total_a_tiempo}</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="card card-yellow">
        <div class="card-title">Sin marcar</div>
        <div class="card-value">{total_sin_marcar}</div>
    </div>
    """, unsafe_allow_html=True)

with col5:
    st.markdown(f"""
    <div class="card">
        <div class="card-title">Cumplimiento</div>
        <div class="card-value">{cumplimiento_general}%</div>
    </div>
    """, unsafe_allow_html=True)

# =========================
# GRÁFICOS
# =========================
st.markdown('<div class="section-title">Visualizaciones</div>', unsafe_allow_html=True)

col_g1, col_g2 = st.columns(2)

with col_g1:
    st.subheader("Top 10 llegadas tarde")
    top_tardanzas = (
        df_filtrado[["nombre", "llegadas_tarde"]]
        .sort_values("llegadas_tarde", ascending=False)
        .head(10)
        .set_index("nombre")
    )
    st.bar_chart(top_tardanzas)

with col_g2:
    st.subheader("Distribución general")
    resumen = pd.DataFrame({
        "tipo": ["Llegadas tarde", "Llegadas a tiempo", "Sin marcar"],
        "cantidad": [total_tardanzas, total_a_tiempo, total_sin_marcar]
    }).set_index("tipo")
    st.bar_chart(resumen)

# =========================
# TABLA DETALLADA
# =========================
st.markdown('<div class="section-title">Tabla detallada</div>', unsafe_allow_html=True)

st.dataframe(
    df_filtrado.sort_values("llegadas_tarde", ascending=False),
    use_container_width=True,
    hide_index=True
)
