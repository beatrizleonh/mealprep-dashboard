import streamlit as st
import pandas as pd
from fractions import Fraction
from pathlib import Path

# ---------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ---------------------------------------------------------

st.set_page_config(
    page_title="Meal Prep Dashboard",
    page_icon="🥗",
    layout="wide"
)

EXCEL_PATH = "base_datos_mealprep_streamlit_v2.xlsx"


# ---------------------------------------------------------
# FUNCIONES AUXILIARES
# ---------------------------------------------------------

@st.cache_data
def cargar_excel():
    """Carga las hojas necesarias del Excel."""
    if not Path(EXCEL_PATH).exists():
        st.error(
            f"No encontré el archivo {EXCEL_PATH}. "
            "Asegúrate de subirlo al repositorio de GitHub."
        )
        st.stop()

    xls = pd.ExcelFile(EXCEL_PATH)

    hojas = {
        "platillos": pd.read_excel(xls, sheet_name="Hoja 1 - Platillos"),
        "ingredientes": pd.read_excel(xls, sheet_name="Ingredientes_base"),
        "preparaciones": pd.read_excel(xls, sheet_name="Hoja 3 - Preparaciones"),
        "equivalencias": pd.read_excel(xls, sheet_name="Hoja 4 - Equivalencias"),
        "despensa": pd.read_excel(xls, sheet_name="Despensa"),
    }

    return hojas


def convertir_a_numero(valor):
    """Convierte cantidades tipo 1/2, 2/3 o 1 a número decimal."""
    if pd.isna(valor):
        return 0

    if isinstance(valor, (int, float)):
        return float(valor)

    valor = str(valor).strip()

    try:
        return float(valor)
    except ValueError:
        pass

    try:
        return float(Fraction(valor))
    except ValueError:
        return 0


def limpiar_texto(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip()


# ---------------------------------------------------------
# CARGA DE DATOS
# ---------------------------------------------------------

data = cargar_excel()

platillos_df = data["platillos"]
ingredientes_df = data["ingredientes"]
preparaciones_df = data["preparaciones"]
equivalencias_df = data["equivalencias"]
despensa_df = data["despensa"]

# Limpieza básica
platillos_df.columns = platillos_df.columns.str.strip().str.lower()
ingredientes_df.columns = ingredientes_df.columns.str.strip().str.lower()
preparaciones_df.columns = preparaciones_df.columns.str.strip().str.lower()
equivalencias_df.columns = equivalencias_df.columns.str.strip().str.lower()
despensa_df.columns = despensa_df.columns.str.strip().str.lower()

platillos_df["platillo"] = platillos_df["platillo"].apply(limpiar_texto)
platillos_df["tipo"] = platillos_df["tipo"].apply(limpiar_texto).str.lower()

ingredientes_df["platillo"] = ingredientes_df["platillo"].apply(limpiar_texto)
ingredientes_df["ingrediente"] = ingredientes_df["ingrediente"].apply(limpiar_texto)
ingredientes_df["grupo"] = ingredientes_df["grupo"].apply(limpiar_texto)
ingredientes_df["unidad"] = ingredientes_df["unidad"].apply(limpiar_texto)

if "cantidad_decimal" not in ingredientes_df.columns:
    ingredientes_df["cantidad_decimal"] = ingredientes_df["cantidad"].apply(convertir_a_numero)

preparaciones_df["platillo"] = preparaciones_df["platillo"].apply(limpiar_texto)


# ---------------------------------------------------------
# TÍTULO
# ---------------------------------------------------------

st.title("🥗 Meal Prep Dashboard")
st.write(
    "Organiza tu menú semanal, calcula ingredientes totales, consulta preparaciones "
    "y revisa equivalencias de alimentos."
)


# ---------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------

st.sidebar.header("Configuración")

personas = st.sidebar.number_input(
    "Personas:",
    min_value=1,
    max_value=50,
    value=1,
    step=1
)

st.sidebar.info(
    "Las cantidades se calculan con base en porciones para 1 persona "
    "y se multiplican por el número de personas seleccionado."
)


# ---------------------------------------------------------
# TABS PRINCIPALES
# ---------------------------------------------------------

tab_menu, tab_ingredientes, tab_preparaciones, tab_equivalencias, tab_despensa = st.tabs(
    [
        "📅 Menú semanal",
        "🛒 Ingredientes totales",
        "👩‍🍳 Preparaciones",
        "🔁 Equivalencias",
        "🏠 Despensa"
    ]
)


# ---------------------------------------------------------
# TAB 1: MENÚ SEMANAL
# ---------------------------------------------------------

with tab_menu:
    st.subheader("📅 Selecciona tu menú semanal")

    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

    tipos_comida = {
        "Desayuno": "desayuno",
        "Almuerzo": "almuerzo",
        "Comida": "comida",
        "Merienda": "merienda",
        "Cena": "cena",
    }

    # Ajuste por si en el Excel aparecen como colación
    equivalencias_tipo = {
        "almuerzo": ["almuerzo", "colación 1", "colacion 1"],
        "merienda": ["merienda", "colación 2", "colacion 2"],
        "desayuno": ["desayuno"],
        "comida": ["comida"],
        "cena": ["cena"],
    }

    seleccionados = []

    for nombre_visible, tipo_base in tipos_comida.items():
        st.markdown(f"### {nombre_visible}")

        columnas = st.columns(7)

        tipos_validos = equivalencias_tipo[tipo_base]

        opciones = platillos_df[
            platillos_df["tipo"].isin(tipos_validos)
        ]["platillo"].dropna().unique().tolist()

        opciones = ["—"] + sorted(opciones)

        for i, dia in enumerate(dias):
            with columnas[i]:
                platillo_elegido = st.selectbox(
                    dia,
                    opciones,
                    key=f"{nombre_visible}_{dia}"
                )

                if platillo_elegido != "—":
                    seleccionados.append({
                        "dia": dia,
                        "tipo": nombre_visible,
                        "platillo": platillo_elegido
                    })

    menu_semanal_df = pd.DataFrame(seleccionados)

    st.divider()

    st.subheader("Resumen del menú seleccionado")

    if menu_semanal_df.empty:
        st.warning("Todavía no has seleccionado platillos.")
    else:
        st.dataframe(menu_semanal_df, use_container_width=True)


# ---------------------------------------------------------
# TAB 2: INGREDIENTES TOTALES
# ---------------------------------------------------------

with tab_ingredientes:
    st.subheader("🛒 Lista total de ingredientes")

    if "menu_semanal_df" not in locals() or menu_semanal_df.empty:
        st.warning("Primero selecciona platillos en la pestaña de Menú semanal.")
    else:
        platillos_seleccionados = menu_semanal_df["platillo"].tolist()

        ingredientes_filtrados = ingredientes_df[
            ingredientes_df["platillo"].isin(platillos_seleccionados)
        ].copy()

        ingredientes_filtrados["cantidad_total"] = (
            ingredientes_filtrados["cantidad_decimal"] * personas
        )

        lista_super = (
            ingredientes_filtrados
            .groupby(["ingrediente", "grupo", "unidad"], as_index=False)["cantidad_total"]
            .sum()
            .sort_values(["grupo", "ingrediente"])
        )

        st.write(f"Ingredientes calculados para **{personas} persona(s)**.")

        st.dataframe(lista_super, use_container_width=True)

        csv = lista_super.to_csv(index=False).encode("utf-8")

        st.download_button(
            label="Descargar lista de ingredientes en CSV",
            data=csv,
            file_name="lista_ingredientes_mealprep.csv",
            mime="text/csv"
        )


# ---------------------------------------------------------
# TAB 3: PREPARACIONES
# ---------------------------------------------------------

with tab_preparaciones:
    st.subheader("👩‍🍳 Preparaciones de los platillos seleccionados")

    if "menu_semanal_df" not in locals() or menu_semanal_df.empty:
        st.warning("Primero selecciona platillos en la pestaña de Menú semanal.")
    else:
        platillos_unicos = menu_semanal_df["platillo"].drop_duplicates().tolist()

        prep_filtradas = preparaciones_df[
            preparaciones_df["platillo"].isin(platillos_unicos)
        ].copy()

        if prep_filtradas.empty:
            st.warning("No encontré preparaciones para los platillos seleccionados.")
        else:
            for _, row in prep_filtradas.iterrows():
                with st.expander(row["platillo"]):
                    columna_preparacion = None

                    for col in prep_filtradas.columns:
                        if "prepar" in col or "modo" in col:
                            columna_preparacion = col
                            break

                    if columna_preparacion:
                        st.write(row[columna_preparacion])
                    else:
                        st.write("No hay preparación disponible.")


# ---------------------------------------------------------
# TAB 4: EQUIVALENCIAS
# ---------------------------------------------------------

with tab_equivalencias:
    st.subheader("🔁 Consulta de equivalencias")

    st.write(
        "Aquí puedes buscar ingredientes equivalentes por categoría. "
        "Por ejemplo: cereales, frutas, verduras, leguminosas, grasas, proteínas, etc."
    )

    st.dataframe(equivalencias_df, use_container_width=True)

    if "categoría" in equivalencias_df.columns:
        categorias = equivalencias_df["categoría"].dropna().unique().tolist()
    elif "categoria" in equivalencias_df.columns:
        categorias = equivalencias_df["categoria"].dropna().unique().tolist()
    else:
        categorias = []

    if categorias:
        categoria_elegida = st.selectbox(
            "Filtrar por categoría:",
            ["Todas"] + sorted(categorias)
        )

        if categoria_elegida != "Todas":
            col_categoria = "categoría" if "categoría" in equivalencias_df.columns else "categoria"
            st.dataframe(
                equivalencias_df[equivalencias_df[col_categoria] == categoria_elegida],
                use_container_width=True
            )


# ---------------------------------------------------------
# TAB 5: DESPENSA
# ---------------------------------------------------------

with tab_despensa:
    st.subheader("🏠 Despensa base")

    st.write(
        "Esta hoja incluye ingredientes frecuentes y básicos de cocina como aceites, sal, "
        "pimienta, vinagre, salsa de soya, miel de agave, especias y otros complementos."
    )

    st.dataframe(despensa_df, use_container_width=True)
