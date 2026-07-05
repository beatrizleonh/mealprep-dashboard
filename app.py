import streamlit as st
import pandas as pd
from fractions import Fraction
from pathlib import Path
from io import BytesIO

# ---------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ---------------------------------------------------------

st.set_page_config(
    page_title="Meal Prep Dashboard",
    page_icon="🥗",
    layout="wide"
)

# OJO: este nombre debe coincidir EXACTAMENTE con el archivo que subiste a GitHub
EXCEL_PATH = "base_datos_mealprep_streamlit.xlsx"


# ---------------------------------------------------------
# FUNCIONES AUXILIARES
# ---------------------------------------------------------

@st.cache_data
def cargar_excel():
    """Carga las hojas necesarias del Excel."""
    if not Path(EXCEL_PATH).exists():
        st.error(
            f"No encontré el archivo: {EXCEL_PATH}. "
            "Revisa que el Excel esté subido a GitHub y que el nombre sea idéntico."
        )
        st.stop()

    xls = pd.ExcelFile(EXCEL_PATH)

    hojas = {
        "platillos": pd.read_excel(xls, sheet_name="Hoja 1 - Platillos"),
        "ingredientes": pd.read_excel(xls, sheet_name="Ingredientes_base"),
        "preparaciones": pd.read_excel(xls, sheet_name="Hoja 3 - Preparaciones"),
        "equivalencias": pd.read_excel(xls, sheet_name="Hoja 4 - Equivalencias"),
    }

    return hojas


def limpiar_texto(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip()


def convertir_a_numero(valor):
    """Convierte cantidades tipo 1/2, 2/3, 1.5 o 2 a número decimal."""
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float)):
        return float(valor)

    valor = str(valor).strip()

    if valor == "":
        return 0.0

    try:
        return float(valor)
    except ValueError:
        pass

    try:
        return float(Fraction(valor))
    except ValueError:
        return 0.0


def formato_cantidad(numero):
    """Convierte cantidades decimales a formato bonito tipo 1/2, 2/3, 1 1/2."""
    try:
        numero = float(numero)
    except Exception:
        return ""

    if numero == 0:
        return ""

    entero = int(numero)
    decimal = numero - entero

    if abs(decimal) < 0.001:
        return str(entero)

    frac = Fraction(decimal).limit_denominator(12)

    if entero == 0:
        return f"{frac.numerator}/{frac.denominator}"

    return f"{entero} {frac.numerator}/{frac.denominator}"


def encontrar_columna(df, posibles_nombres):
    """Busca una columna aunque tenga acentos, espacios o variantes."""
    columnas = list(df.columns)

    for posible in posibles_nombres:
        for col in columnas:
            if col.lower().strip() == posible.lower().strip():
                return col

    return None


def crear_excel_descarga(df):
    """Crea archivo Excel descargable desde un dataframe."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ingredientes faltantes")

    return output.getvalue()


# ---------------------------------------------------------
# CARGA DE DATOS
# ---------------------------------------------------------

data = cargar_excel()

platillos_df = data["platillos"]
ingredientes_df = data["ingredientes"]
preparaciones_df = data["preparaciones"]
equivalencias_df = data["equivalencias"]

# Limpieza de nombres de columnas
platillos_df.columns = platillos_df.columns.str.strip().str.lower()
ingredientes_df.columns = ingredientes_df.columns.str.strip().str.lower()
preparaciones_df.columns = preparaciones_df.columns.str.strip().str.lower()
equivalencias_df.columns = equivalencias_df.columns.str.strip().str.lower()

# Limpieza de texto
platillos_df["platillo"] = platillos_df["platillo"].apply(limpiar_texto)
platillos_df["tipo"] = platillos_df["tipo"].apply(limpiar_texto).str.lower()

ingredientes_df["platillo"] = ingredientes_df["platillo"].apply(limpiar_texto)
ingredientes_df["ingrediente"] = ingredientes_df["ingrediente"].apply(limpiar_texto)
ingredientes_df["grupo"] = ingredientes_df["grupo"].apply(limpiar_texto)
ingredientes_df["unidad"] = ingredientes_df["unidad"].apply(limpiar_texto)

preparaciones_df["platillo"] = preparaciones_df["platillo"].apply(limpiar_texto)

if "cantidad_decimal" not in ingredientes_df.columns:
    ingredientes_df["cantidad_decimal"] = ingredientes_df["cantidad"].apply(convertir_a_numero)

ingredientes_df["cantidad_decimal"] = ingredientes_df["cantidad_decimal"].apply(convertir_a_numero)


# ---------------------------------------------------------
# TÍTULO
# ---------------------------------------------------------

st.title("🥗 Meal Prep Dashboard")

st.write(
    "Planea tus comidas de la semana, calcula ingredientes totales, "
    "marca lo que ya tienes y revisa preparaciones y equivalencias."
)


# ---------------------------------------------------------
# SIDEBAR
# ---------------------------------------------------------

st.sidebar.header("Configuración")

personas = st.sidebar.number_input(
    "Personas:",
    min_value=1,
    max_value=100,
    value=1,
    step=1
)

st.sidebar.info(
    "Las recetas están en porciones para 1 persona. "
    "Aquí se multiplican automáticamente por el número de personas."
)


# ---------------------------------------------------------
# TABS
# ---------------------------------------------------------

tab_menu, tab_ingredientes, tab_preparaciones, tab_equivalencias = st.tabs(
    [
        "📅 Menú semanal",
        "🛒 Ingredientes totales",
        "👩‍🍳 Preparaciones",
        "🔁 Equivalencias",
    ]
)


# ---------------------------------------------------------
# TAB 1: MENÚ SEMANAL
# ---------------------------------------------------------

with tab_menu:
    st.subheader("📅 Menú semanal")

    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]

    filas_comida = {
        "Desayuno": ["desayuno"],
        "Almuerzo": ["almuerzo", "colación 1", "colacion 1"],
        "Comida": ["comida"],
        "Merienda": ["merienda", "colación 2", "colacion 2"],
        "Cena": ["cena"],
    }

    seleccionados = []

    # Encabezados de tabla
    header_cols = st.columns([1.2, 1, 1, 1, 1, 1, 1, 1])

    with header_cols[0]:
        st.markdown("**Comida**")

    for i, dia in enumerate(dias):
        with header_cols[i + 1]:
            st.markdown(f"**{dia}**")

    st.divider()

    # Filas de la tabla
    for comida_visible, tipos_validos in filas_comida.items():
        row_cols = st.columns([1.2, 1, 1, 1, 1, 1, 1, 1])

        with row_cols[0]:
            st.markdown(f"### {comida_visible}")

        opciones = (
            platillos_df[platillos_df["tipo"].isin(tipos_validos)]["platillo"]
            .dropna()
            .unique()
            .tolist()
        )

        opciones = ["—"] + sorted(opciones)

        for i, dia in enumerate(dias):
            with row_cols[i + 1]:
                platillo_elegido = st.selectbox(
                    label=f"{comida_visible} {dia}",
                    options=opciones,
                    key=f"menu_{comida_visible}_{dia}",
                    label_visibility="collapsed"
                )

                if platillo_elegido != "—":
                    seleccionados.append(
                        {
                            "dia": dia,
                            "tipo": comida_visible,
                            "platillo": platillo_elegido,
                        }
                    )

    menu_semanal_df = pd.DataFrame(seleccionados)

    st.divider()

    st.subheader("Resumen del menú seleccionado")

    if menu_semanal_df.empty:
        st.warning("Todavía no has seleccionado platillos.")
    else:
        resumen = menu_semanal_df.pivot(
            index="tipo",
            columns="dia",
            values="platillo"
        )

        orden_filas = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]
        orden_columnas = dias

        resumen = resumen.reindex(index=orden_filas, columns=orden_columnas)

        st.dataframe(resumen, use_container_width=True)


# ---------------------------------------------------------
# TAB 2: INGREDIENTES TOTALES
# ---------------------------------------------------------

with tab_ingredientes:
    st.subheader("🛒 Lista total de ingredientes")

    if menu_semanal_df.empty:
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
            .groupby(["grupo", "ingrediente", "unidad"], as_index=False)["cantidad_total"]
            .sum()
            .sort_values(["grupo", "ingrediente"])
        )

        lista_super["cantidad_formato"] = lista_super["cantidad_total"].apply(formato_cantidad)

        st.write(f"Ingredientes calculados para **{personas} persona(s)**.")

        ingredientes_faltantes = []

        grupos = sorted(lista_super["grupo"].dropna().unique().tolist())

        for grupo in grupos:
            grupo_df = lista_super[lista_super["grupo"] == grupo].copy()

            st.markdown(f"## {grupo}")

            for _, row in grupo_df.iterrows():
                ingrediente = row["ingrediente"]
                cantidad = row["cantidad_formato"]
                unidad = row["unidad"]

                if cantidad == "":
                    texto_ingrediente = f"{ingrediente} — {unidad}"
                else:
                    texto_ingrediente = f"{ingrediente} — {cantidad} {unidad}"

                key_checkbox = f"check_ingrediente_{grupo}_{ingrediente}_{unidad}"

                marcado = st.checkbox(
                    texto_ingrediente,
                    key=key_checkbox
                )

                if not marcado:
                    ingredientes_faltantes.append(
                        {
                            "grupo": grupo,
                            "ingrediente": ingrediente,
                            "cantidad": cantidad,
                            "unidad": unidad,
                        }
                    )

            st.divider()

        st.subheader("🧾 Ingredientes faltantes")

        if len(ingredientes_faltantes) == 0:
            st.success("Ya marcaste todos los ingredientes como listos ✅")
        else:
            faltantes_df = pd.DataFrame(ingredientes_faltantes)

            grupos_faltantes = sorted(faltantes_df["grupo"].dropna().unique().tolist())

            for grupo in grupos_faltantes:
                st.markdown(f"### {grupo}")

                grupo_faltantes = faltantes_df[faltantes_df["grupo"] == grupo]

                for _, row in grupo_faltantes.iterrows():
                    cantidad = row["cantidad"]
                    unidad = row["unidad"]
                    ingrediente = row["ingrediente"]

                    if cantidad == "":
                        st.write(f"- {ingrediente} — {unidad}")
                    else:
                        st.write(f"- {ingrediente} — {cantidad} {unidad}")

            excel_faltantes = crear_excel_descarga(faltantes_df)

            st.download_button(
                label="Descargar ingredientes faltantes en Excel",
                data=excel_faltantes,
                file_name="ingredientes_faltantes_mealprep.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )


# ---------------------------------------------------------
# TAB 3: PREPARACIONES
# ---------------------------------------------------------

with tab_preparaciones:
    st.subheader("👩‍🍳 Preparaciones")

    if menu_semanal_df.empty:
        st.warning("Primero selecciona platillos en la pestaña de Menú semanal.")
    else:
        platillos_unicos = sorted(menu_semanal_df["platillo"].drop_duplicates().tolist())

        prep_filtradas = preparaciones_df[
            preparaciones_df["platillo"].isin(platillos_unicos)
        ].copy()

        columna_preparacion = encontrar_columna(
            prep_filtradas,
            ["modo de preparación", "modo de preparacion", "preparación", "preparacion"]
        )

        if prep_filtradas.empty:
            st.warning("No encontré preparaciones para los platillos seleccionados.")
        else:
            for _, row in prep_filtradas.sort_values("platillo").iterrows():
                platillo = row["platillo"]
                key_hecho = f"prep_hecha_{platillo}"

                hecho = st.checkbox(
                    f"Ya está hecho: {platillo}",
                    key=key_hecho
                )

                if hecho:
                    st.success(f"✅ {platillo} ya está preparado")

                with st.expander(f"Ver preparación: {platillo}", expanded=not hecho):
                    if columna_preparacion:
                        st.write(row[columna_preparacion])
                    else:
                        st.write("No hay preparación disponible.")

                st.divider()


# ---------------------------------------------------------
# TAB 4: EQUIVALENCIAS
# ---------------------------------------------------------

with tab_equivalencias:
    st.subheader("🔁 Equivalencias por categoría")

    categoria_col = encontrar_columna(equivalencias_df, ["categoría", "categoria"])
    ingrediente_col = encontrar_columna(equivalencias_df, ["ingrediente"])
    cantidad_col = encontrar_columna(equivalencias_df, ["cantidad"])

    if not categoria_col or not ingrediente_col or not cantidad_col:
        st.error(
            "No encontré las columnas necesarias en la hoja de equivalencias. "
            "Necesito: categoría, ingrediente y cantidad."
        )
    else:
        equivalencias_limpias = equivalencias_df[
            [categoria_col, ingrediente_col, cantidad_col]
        ].copy()

        equivalencias_limpias.columns = ["categoría", "ingrediente", "cantidad"]

        equivalencias_limpias["categoría"] = equivalencias_limpias["categoría"].apply(limpiar_texto)
        equivalencias_limpias["ingrediente"] = equivalencias_limpias["ingrediente"].apply(limpiar_texto)
        equivalencias_limpias["cantidad"] = equivalencias_limpias["cantidad"].apply(limpiar_texto)

        equivalencias_limpias = equivalencias_limpias.sort_values(
            by=["categoría", "ingrediente"]
        )

        categorias = sorted(equivalencias_limpias["categoría"].dropna().unique().tolist())

        for categoria in categorias:
            st.markdown(f"## {categoria}")

            tabla_categoria = equivalencias_limpias[
                equivalencias_limpias["categoría"] == categoria
            ][["ingrediente", "cantidad"]].sort_values("ingrediente")

            st.dataframe(
                tabla_categoria,
                use_container_width=True,
                hide_index=True
            )

            st.divider()
