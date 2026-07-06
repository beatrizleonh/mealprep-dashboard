import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from fractions import Fraction
from pathlib import Path
import html
import json
import re

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


def get_menu_value(menu_df, dia, comida):
    fila = menu_df[
        (menu_df["dia"] == dia) &
        (menu_df["tipo"] == comida)
    ]

    if fila.empty:
        return ""

    return str(fila.iloc[0]["platillo"]).strip()


def generar_texto_dia(menu_df, dia):
    comidas = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]
    lineas = [f"{dia}:"]

    for comida in comidas:
        platillo = get_menu_value(menu_df, dia, comida)
        lineas.append(f"{comida}: {platillo}")

    return "\n".join(lineas)


def generar_texto_tabla(menu_df):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

    bloques = []

    for dia in dias:
        bloques.append(generar_texto_dia(menu_df, dia))

    return "\n\n".join(bloques)


def render_resumen_menu(menu_df):
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    comidas = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]

    cards_html = """
    <!DOCTYPE html>
    <html>
    <head>
    <style>
    body {
        margin: 0;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        color: #111827;
    }

    .cards-container {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
        gap: 14px;
        width: 100%;
    }

    .day-card {
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 14px;
        background: #ffffff;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }

    .day-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        gap: 10px;
        margin-bottom: 10px;
    }

    .day-title {
        font-weight: 800;
        font-size: 18px;
    }

    .copy-btn {
        border: 1px solid #d1d5db;
        background-color: #ffffff;
        border-radius: 8px;
        padding: 5px 8px;
        font-size: 12px;
        cursor: pointer;
        white-space: nowrap;
    }

    .copy-btn:hover {
        background-color: #f3f4f6;
    }

    .meal-row {
        margin-bottom: 8px;
        line-height: 1.35;
    }

    .meal-label {
        font-weight: 700;
        display: block;
        font-size: 13px;
        color: #374151;
    }

    .meal-value {
        font-size: 14px;
        color: #111827;
    }

    .empty-value {
        color: #9ca3af;
        font-style: italic;
    }
    </style>
    </head>
    <body>
    <div class="cards-container">
    """

    for dia in dias:
        texto_dia = generar_texto_dia(menu_df, dia)
        texto_dia_json = json.dumps(texto_dia)

        cards_html += f"""
        <div class="day-card">
            <div class="day-header">
                <div class="day-title">{html.escape(dia)}</div>
                <button class="copy-btn" onclick='navigator.clipboard.writeText({texto_dia_json}); this.innerText="Copiado ✅"; setTimeout(() => this.innerText="Copiar día", 1500);'>
                    Copiar día
                </button>
            </div>
        """

        for comida in comidas:
            platillo = get_menu_value(menu_df, dia, comida)
            if platillo:
                valor_html = f'<span class="meal-value">{html.escape(platillo)}</span>'
            else:
                valor_html = '<span class="meal-value empty-value">Sin elegir</span>'

            cards_html += f"""
            <div class="meal-row">
                <span class="meal-label">{html.escape(comida)}:</span>
                {valor_html}
            </div>
            """

        cards_html += "</div>"

    cards_html += """
    </div>
    </body>
    </html>
    """

    components.html(cards_html, height=520, scrolling=True)


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


def parsear_menu_base(texto):
    dias_validos = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    comidas_validas = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]

    resultado = []
    dia_actual = None

    lineas = texto.splitlines()

    for linea in lineas:
        linea = linea.strip()

        if not linea:
            continue

        linea_normalizada = linea.replace("Miercoles", "Miércoles")

        match_dia = re.match(r"^(Lunes|Martes|Miércoles|Miercoles|Jueves|Viernes)\s*:\s*$", linea_normalizada, re.IGNORECASE)

        if match_dia:
            dia = match_dia.group(1)
            if dia.lower() == "miercoles":
                dia = "Miércoles"
            else:
                dia = dia.capitalize()
                if dia == "Miércoles":
                    dia = "Miércoles"

            if dia in dias_validos:
                dia_actual = dia

            continue

        if dia_actual:
            for comida in comidas_validas:
                patron = rf"^{comida}\s*:\s*(.*)$"
                match_comida = re.match(patron, linea, re.IGNORECASE)

                if match_comida:
                    platillo = match_comida.group(1).strip()

                    if platillo:
                        resultado.append(
                            {
                                "dia": dia_actual,
                                "tipo": comida,
                                "platillo": platillo
                            }
                        )

                    break

    return pd.DataFrame(resultado)


def cargar_menu_en_session_state(menu_df):
    for _, row in menu_df.iterrows():
        key = f"menu_{row['tipo']}_{row['dia']}"
        st.session_state[key] = row["platillo"]


def limpiar_menu_session_state():
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    comidas = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]

    for comida in comidas:
        for dia in dias:
            key = f"menu_{comida}_{dia}"
            st.session_state[key] = "—"


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
    "Planea tus comidas de lunes a viernes, calcula ingredientes totales, "
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

st.sidebar.divider()

if st.sidebar.button("Iniciar menú nuevo"):
    limpiar_menu_session_state()
    st.sidebar.success("Menú reiniciado. Cambia de pestaña o actualiza para verlo limpio.")


# ---------------------------------------------------------
# TABS
# ---------------------------------------------------------

tab_menu, tab_ingredientes, tab_preparaciones, tab_equivalencias, tab_base = st.tabs(
    [
        "📅 Menú semanal",
        "🛒 Ingredientes totales",
        "👩‍🍳 Preparaciones",
        "🔁 Equivalencias",
        "📋 Base",
    ]
)


# ---------------------------------------------------------
# TAB 1: MENÚ SEMANAL
# ---------------------------------------------------------

with tab_menu:
    st.subheader("📅 Menú semanal")

    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]

    filas_comida = {
        "Desayuno": ["desayuno"],
        "Almuerzo": ["almuerzo", "colación 1", "colacion 1"],
        "Comida": ["comida"],
        "Merienda": ["merienda", "colación 2", "colacion 2"],
        "Cena": ["cena"],
    }

    seleccionados = []

    header_cols = st.columns([1.25, 1, 1, 1, 1, 1])

    with header_cols[0]:
        st.markdown("**Comida**")

    for i, dia in enumerate(dias):
        with header_cols[i + 1]:
            st.markdown(f"**{dia}**")

    st.divider()

    for comida_visible, tipos_validos in filas_comida.items():
        row_cols = st.columns([1.25, 1, 1, 1, 1, 1])

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
            key = f"menu_{comida_visible}_{dia}"

            valor_actual = st.session_state.get(key, "—")

            if valor_actual not in opciones:
                opciones_mostradas = ["—", valor_actual] + [op for op in opciones if op not in ["—", valor_actual]]
            else:
                opciones_mostradas = opciones

            index_default = opciones_mostradas.index(valor_actual) if valor_actual in opciones_mostradas else 0

            with row_cols[i + 1]:
                platillo_elegido = st.selectbox(
                    label=f"{comida_visible} {dia}",
                    options=opciones_mostradas,
                    index=index_default,
                    key=key,
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
        st.warning("Primero selecciona platillos en la pestaña de Menú semanal o carga una base guardada.")
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
        st.warning("Primero selecciona platillos en la pestaña de Menú semanal o carga una base guardada.")
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


# ---------------------------------------------------------
# TAB 5: BASE
# ---------------------------------------------------------

with tab_base:
    st.subheader("📋 Base guardada del menú")

    st.write(
        "Pega aquí el texto que copiaste con el botón **Copiar tabla**. "
        "Después da clic en **Cargar menú desde base** para recuperar tu menú semanal."
    )

    ejemplo = """Lunes:
Desayuno: Avena con arándanos y nuez
Almuerzo: Chilaquiles con huevo y frijol
Comida: Bowl de tofu mediterraneo con quinoa
Merienda: Helado plátano y frutos
Cena: Mug cake chocolate

Martes:
Desayuno: Overnight oats sabor pay de limón
Almuerzo: Alambre de seitán con frijoles
Comida: Burrito salad
Merienda: Helado plátano y frutos
Cena: Cereal de soya con nuez y frutos rojos"""

    texto_base = st.text_area(
        "Pega aquí tu menú guardado:",
        height=350,
        placeholder=ejemplo,
        key="texto_base_menu"
    )

    col_cargar, col_limpiar = st.columns([1, 1])

    with col_cargar:
        if st.button("Cargar menú desde base"):
            menu_cargado_df = parsear_menu_base(texto_base)

            if menu_cargado_df.empty:
                st.error("No pude leer el menú. Revisa que tenga el formato: Día: y luego Comida: platillo.")
            else:
                cargar_menu_en_session_state(menu_cargado_df)
                st.success("Menú cargado correctamente. Ve a la pestaña Menú semanal para verlo actualizado.")

    with col_limpiar:
        if st.button("Limpiar menú actual"):
            limpiar_menu_session_state()
            st.success("Menú limpiado. Ve a la pestaña Menú semanal para confirmar.")
