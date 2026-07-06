import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from fractions import Fraction
from pathlib import Path
import html
import json

# ---------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ---------------------------------------------------------

st.set_page_config(
    page_title="Meal Prep Dashboard",
    page_icon="🥗",
    layout="wide"
)

EXCEL_PATH = "base_datos_mealprep_streamlit.xlsx"


# ---------------------------------------------------------
# CSS GENERAL
# ---------------------------------------------------------

st.markdown(
    """
    <style>
    .ingredient-group {
        margin-top: 18px;
        margin-bottom: 6px;
        font-size: 23px;
        font-weight: 700;
    }

    .missing-list-box {
        background-color: #fafafa;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 16px 20px;
        margin-top: 10px;
        font-family: inherit;
        font-size: 16px;
        line-height: 1.35;
    }

    .missing-category {
        font-weight: 800;
        text-transform: uppercase;
        margin-top: 12px;
        margin-bottom: 4px;
    }

    .missing-category:first-child {
        margin-top: 0;
    }

    .missing-item {
        margin: 1px 0 1px 14px;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# ---------------------------------------------------------
# FUNCIONES AUXILIARES
# ---------------------------------------------------------

@st.cache_data
def cargar_excel():
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
    columnas = list(df.columns)

    for posible in posibles_nombres:
        for col in columnas:
            if col.lower().strip() == posible.lower().strip():
                return col

    return None


def copy_button(texto, label="Copiar", height=42):
    texto_json = json.dumps(texto)

    components.html(
        f"""
        <button class="copy-btn" onclick='navigator.clipboard.writeText({texto_json}); this.innerText="Copiado ✅"; setTimeout(() => this.innerText="{label}", 1500);'>
            {label}
        </button>

        <style>
        .copy-btn {{
            border: 1px solid #d1d5db;
            background-color: #ffffff;
            border-radius: 8px;
            padding: 7px 11px;
            font-size: 13px;
            cursor: pointer;
            font-family: sans-serif;
        }}

        .copy-btn:hover {{
            background-color: #f3f4f6;
        }}
        </style>
        """,
        height=height
    )


def generar_texto_dia(menu_df, dia):
    comidas = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]
    lineas = [f"{dia}:"]

    for comida in comidas:
        fila = menu_df[
            (menu_df["dia"] == dia) &
            (menu_df["tipo"] == comida)
        ]

        if fila.empty:
            platillo = ""
        else:
            platillo = fila.iloc[0]["platillo"]

        lineas.append(f"{comida}: {platillo}")

    return "\n".join(lineas)


def generar_texto_tabla(menu_df):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    comidas = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]

    lineas = []
    lineas.append("Comida\t" + "\t".join(dias))

    for comida in comidas:
        valores = []

        for dia in dias:
            fila = menu_df[
                (menu_df["dia"] == dia) &
                (menu_df["tipo"] == comida)
            ]

            if fila.empty:
                valores.append("")
            else:
                valores.append(str(fila.iloc[0]["platillo"]))

        lineas.append(comida + "\t" + "\t".join(valores))

    return "\n".join(lineas)


def render_resumen_menu(menu_df):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    comidas = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]

    html_table = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    body {
        margin: 0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        color: #111827;
    }

    .summary-table {
        width: 100%;
        border-collapse: collapse;
        table-layout: fixed;
        font-size: 15px;
    }

    .summary-table th {
        background-color: #f7f8fa;
        padding: 10px;
        border: 1px solid #e5e7eb;
        text-align: left;
        vertical-align: top;
        font-weight: 700;
    }

    .summary-table td {
        padding: 12px;
        border: 1px solid #e5e7eb;
        text-align: left;
        vertical-align: top;
        white-space: normal;
        word-wrap: break-word;
        overflow-wrap: break-word;
        line-height: 1.35;
        min-height: 65px;
    }

    .meal-col {
        width: 115px;
        font-weight: 700;
        background-color: #fafafa;
    }

    .day-header {
        display: flex;
        flex-direction: column;
        gap: 6px;
    }

    .copy-btn {
        border: 1px solid #d1d5db;
        background-color: #ffffff;
        border-radius: 8px;
        padding: 5px 8px;
        font-size: 12px;
        cursor: pointer;
        width: fit-content;
    }

    .copy-btn:hover {
        background-color: #f3f4f6;
    }
    </style>
    </head>
    <body>
    <table class="summary-table">
        <thead>
            <tr>
                <th class="meal-col">Tipo</th>
    """

    for dia in dias:
        texto_dia = generar_texto_dia(menu_df, dia)
        texto_dia_json = json.dumps(texto_dia)

        html_table += f"""
                <th>
                    <div class="day-header">
                        <span>{html.escape(dia)}</span>
                        <button class="copy-btn" onclick='navigator.clipboard.writeText({texto_dia_json}); this.innerText="Copiado ✅"; setTimeout(() => this.innerText="Copiar día", 1500);'>
                            Copiar día
                        </button>
                    </div>
                </th>
        """

    html_table += """
            </tr>
        </thead>
        <tbody>
    """

    for comida in comidas:
        html_table += f"""
            <tr>
                <td class="meal-col">{html.escape(comida)}</td>
        """

        for dia in dias:
            fila = menu_df[
                (menu_df["dia"] == dia) &
                (menu_df["tipo"] == comida)
            ]

            if fila.empty:
                valor = ""
            else:
                valor = str(fila.iloc[0]["platillo"])

            html_table += f"<td>{html.escape(valor)}</td>"

        html_table += "</tr>"

    html_table += """
        </tbody>
    </table>
    </body>
    </html>
    """

    components.html(html_table, height=430, scrolling=True)


def generar_texto_ingredientes(faltantes_df):
    if faltantes_df.empty:
        return "Todos los ingredientes están marcados como listos."

    lineas = []

    for grupo in sorted(faltantes_df["grupo"].dropna().unique().tolist()):
        grupo_df = faltantes_df[faltantes_df["grupo"] == grupo].sort_values("ingrediente")

        lineas.append(f"{grupo.upper()}:")

        for _, row in grupo_df.iterrows():
            ingrediente = row["ingrediente"]
            cantidad = row["cantidad"]
            unidad = row["unidad"]

            if cantidad:
                lineas.append(f"- {ingrediente} {cantidad} {unidad}")
            else:
                lineas.append(f"- {ingrediente} {unidad}")

        lineas.append("")

    return "\n".join(lineas).strip()


def render_ingredientes_faltantes(faltantes_df):
    if faltantes_df.empty:
        st.success("Ya marcaste todos los ingredientes como listos ✅")
        return

    html_lista = '<div class="missing-list-box">'

    for grupo in sorted(faltantes_df["grupo"].dropna().unique().tolist()):
        grupo_df = faltantes_df[faltantes_df["grupo"] == grupo].sort_values("ingrediente")

        html_lista += f'<div class="missing-category">{html.escape(grupo.upper())}:</div>'

        for _, row in grupo_df.iterrows():
            ingrediente = row["ingrediente"]
            cantidad = row["cantidad"]
            unidad = row["unidad"]

            if cantidad:
                texto = f"- {ingrediente} {cantidad} {unidad}"
            else:
                texto = f"- {ingrediente} {unidad}"

            html_lista += f'<div class="missing-item">{html.escape(texto)}</div>'

    html_lista += "</div>"

    st.markdown(html_lista, unsafe_allow_html=True)


# ---------------------------------------------------------
# CARGA DE DATOS
# ---------------------------------------------------------

data = cargar_excel()

platillos_df = data["platillos"]
ingredientes_df = data["ingredientes"]
preparaciones_df = data["preparaciones"]
equivalencias_df = data["equivalencias"]

platillos_df.columns = platillos_df.columns.str.strip().str.lower()
ingredientes_df.columns = ingredientes_df.columns.str.strip().str.lower()
preparaciones_df.columns = preparaciones_df.columns.str.strip().str.lower()
equivalencias_df.columns = equivalencias_df.columns.str.strip().str.lower()

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

    header_cols = st.columns([1.25, 1, 1, 1, 1, 1, 1, 1])

    with header_cols[0]:
        st.markdown("**Comida**")

    for i, dia in enumerate(dias):
        with header_cols[i + 1]:
            st.markdown(f"**{dia}**")

    st.divider()

    for comida_visible, tipos_validos in filas_comida.items():
        row_cols = st.columns([1.25, 1, 1, 1, 1, 1, 1, 1])

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

    title_col, copy_col = st.columns([5, 1])

    with title_col:
        st.subheader("Resumen del menú seleccionado")

    if menu_semanal_df.empty:
        st.warning("Todavía no has seleccionado platillos.")
    else:
        texto_tabla = generar_texto_tabla(menu_semanal_df)

        with copy_col:
            copy_button(texto_tabla, label="Copiar tabla")

        render_resumen_menu(menu_semanal_df)


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

        st.write(
            f"Ingredientes calculados para **{personas} persona(s)** "
            "con base en todos los platillos seleccionados en el menú semanal."
        )

        ingredientes_faltantes = []

        grupos = sorted(lista_super["grupo"].dropna().unique().tolist())

        for grupo in grupos:
            grupo_df = lista_super[lista_super["grupo"] == grupo].copy()

            st.markdown(
                f'<div class="ingredient-group">{html.escape(grupo)}</div>',
                unsafe_allow_html=True
            )

            for _, row in grupo_df.iterrows():
                ingrediente = row["ingrediente"]
                cantidad = row["cantidad_formato"]
                unidad = row["unidad"]

                if cantidad:
                    texto_ingrediente = f"{ingrediente} {cantidad} {unidad}"
                else:
                    texto_ingrediente = f"{ingrediente} {unidad}"

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
            texto_faltantes = generar_texto_ingredientes(faltantes_df)

            copy_button(texto_faltantes, label="Copiar lista de ingredientes faltantes")

            render_ingredientes_faltantes(faltantes_df)


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

                hecho = st.session_state.get(key_hecho, False)

                titulo_expander = f"~~{platillo}~~" if hecho else platillo

                with st.expander(titulo_expander, expanded=False):
                    st.checkbox(
                        "Marcar como preparado",
                        key=key_hecho
                    )

                    st.divider()

                    if columna_preparacion:
                        st.write(row[columna_preparacion])
                    else:
                        st.write("No hay preparación disponible.")


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
