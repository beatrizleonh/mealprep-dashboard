import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from fractions import Fraction
from pathlib import Path
import html
import json
import re
import unicodedata

# ---------------------------------------------------------
# CONFIGURACIÓN GENERAL
# ---------------------------------------------------------

st.set_page_config(
    page_title="Meal Prep Dashboard",
    page_icon="🥗",
    layout="wide"
)

EXCEL_PATH = "base_datos_mealprep_streamlit.xlsx"

DIAS = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
COMIDAS = ["Desayuno", "Almuerzo", "Comida", "Merienda", "Cena"]


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

    .ingredient-note {
        color: #6b7280;
        font-size: 0.92rem;
    }

    .legend-box {
        background-color: #fafafa;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 10px 14px;
        margin-bottom: 14px;
        font-size: 0.95rem;
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

    data = {
        "platillos": pd.read_excel(xls, sheet_name="Hoja 1 - Platillos"),
        "ingredientes": pd.read_excel(xls, sheet_name="Ingredientes_base"),
        "preparaciones": pd.read_excel(xls, sheet_name="Hoja 3 - Preparaciones"),
        "equivalencias": pd.read_excel(xls, sheet_name="Hoja 4 - Equivalencias"),
    }

    if "Hoja 7 - Hogar" in xls.sheet_names:
        data["hogar"] = pd.read_excel(xls, sheet_name="Hoja 7 - Hogar")
    else:
        data["hogar"] = pd.DataFrame(columns=["Nombre", "Categoría"])

    return data


def limpiar_texto(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip()


def normalizar(texto):
    texto = limpiar_texto(texto).lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = re.sub(r"\s+", " ", texto).strip()
    return texto


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


def separar_cantidad_unidad(texto):
    texto = limpiar_texto(texto)

    if texto == "":
        return 0.0, ""

    match = re.match(r"^(\d+\s+\d+/\d+|\d+/\d+|\d+(\.\d+)?)(.*)$", texto)

    if not match:
        return 0.0, texto

    cantidad_txt = match.group(1).strip()
    unidad = match.group(3).strip()

    if " " in cantidad_txt and "/" in cantidad_txt:
        partes = cantidad_txt.split()
        cantidad = float(partes[0]) + float(Fraction(partes[1]))
    else:
        cantidad = convertir_a_numero(cantidad_txt)

    return cantidad, unidad


def encontrar_columna(df, posibles_nombres):
    for posible in posibles_nombres:
        for col in df.columns:
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
    lineas = [f"{dia}:"]

    for comida in COMIDAS:
        platillo = get_menu_value(menu_df, dia, comida)
        lineas.append(f"{comida}: {platillo}")

    return "\n".join(lineas)


def generar_texto_tabla(menu_df):
    bloques = []

    for dia in DIAS:
        bloques.append(generar_texto_dia(menu_df, dia))

    return "\n\n".join(bloques)


def render_resumen_menu(menu_df):
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
        grid-template-columns: repeat(auto-fit, minmax(230px, 1fr));
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

    for dia in DIAS:
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

        for comida in COMIDAS:
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

    components.html(cards_html, height=540, scrolling=True)


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
            platillos = row.get("platillos", "")

            if cantidad:
                linea = f"- {ingrediente} {cantidad} {unidad}"
            else:
                linea = f"- {ingrediente} {unidad}"

            if platillos:
                linea += f" ({platillos})"

            lineas.append(linea)

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
            platillos = row.get("platillos", "")

            if cantidad:
                texto = f"- {ingrediente} {cantidad} {unidad}"
            else:
                texto = f"- {ingrediente} {unidad}"

            if platillos:
                texto += f" ({platillos})"

            html_lista += f'<div class="missing-item">{html.escape(texto)}</div>'

    html_lista += "</div>"

    st.markdown(html_lista, unsafe_allow_html=True)


def consolidar_ingredientes_faltantes(faltantes_df):
    if faltantes_df.empty:
        return faltantes_df

    df = faltantes_df.copy()
    df["cantidad_num"] = df["cantidad"].apply(convertir_a_numero)

    consolidado = (
        df.groupby(["grupo", "ingrediente", "unidad"], as_index=False)
        .agg(
            cantidad_num=("cantidad_num", "sum"),
            platillos=(
                "platillos",
                lambda x: ", ".join(
                    sorted(
                        set(
                            p.strip()
                            for item in x
                            for p in str(item).split(",")
                            if p.strip()
                        )
                    )
                )
            )
        )
    )

    consolidado["cantidad"] = consolidado["cantidad_num"].apply(formato_cantidad)

    consolidado = consolidado[
        ["grupo", "ingrediente", "cantidad", "unidad", "platillos"]
    ].sort_values(["grupo", "ingrediente"])

    return consolidado


def parsear_menu_base(texto):
    resultado = []
    dia_actual = None

    for linea in texto.splitlines():
        linea = linea.strip()

        if not linea:
            continue

        linea_normalizada = linea.replace("Miercoles", "Miércoles")

        match_dia = re.match(
            r"^(Lunes|Martes|Miércoles|Miercoles|Jueves|Viernes)\s*:\s*$",
            linea_normalizada,
            re.IGNORECASE
        )

        if match_dia:
            dia = match_dia.group(1)

            if normalizar(dia) == "miercoles":
                dia_actual = "Miércoles"
            else:
                dia_actual = dia.capitalize()

            continue

        if dia_actual:
            for comida in COMIDAS:
                match_comida = re.match(
                    rf"^{comida}\s*:\s*(.*)$",
                    linea,
                    re.IGNORECASE
                )

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
    for comida in COMIDAS:
        for dia in DIAS:
            key = f"menu_{comida}_{dia}"
            st.session_state[key] = "—"


def preparar_equivalencias(equivalencias_df):
    categoria_col = encontrar_columna(equivalencias_df, ["categoría", "categoria"])
    ingrediente_col = encontrar_columna(equivalencias_df, ["ingrediente"])
    cantidad_col = encontrar_columna(equivalencias_df, ["cantidad"])

    if not categoria_col or not ingrediente_col or not cantidad_col:
        return pd.DataFrame()

    eq = equivalencias_df[[categoria_col, ingrediente_col, cantidad_col]].copy()
    eq.columns = ["categoría", "ingrediente", "cantidad"]

    eq["categoría"] = eq["categoría"].apply(limpiar_texto)
    eq["ingrediente"] = eq["ingrediente"].apply(limpiar_texto)
    eq["cantidad"] = eq["cantidad"].apply(limpiar_texto)
    eq["categoria_norm"] = eq["categoría"].apply(normalizar)
    eq["ingrediente_norm"] = eq["ingrediente"].apply(normalizar)

    eq[["cantidad_num", "unidad_eq"]] = eq["cantidad"].apply(
        lambda x: pd.Series(separar_cantidad_unidad(x))
    )

    return eq


def obtener_equivalentes(eq_df, grupo):
    grupo_norm = normalizar(grupo)

    opciones = (
        eq_df[eq_df["categoria_norm"] == grupo_norm]["ingrediente"]
        .dropna()
        .unique()
        .tolist()
    )

    return sorted(opciones)


def calcular_sustitucion(eq_df, grupo, ingrediente_original, cantidad_total_original, unidad_original, ingrediente_elegido):
    grupo_norm = normalizar(grupo)
    original_norm = normalizar(ingrediente_original)
    elegido_norm = normalizar(ingrediente_elegido)

    original_eq = eq_df[
        (eq_df["categoria_norm"] == grupo_norm) &
        (eq_df["ingrediente_norm"] == original_norm)
    ]

    elegido_eq = eq_df[
        (eq_df["categoria_norm"] == grupo_norm) &
        (eq_df["ingrediente_norm"] == elegido_norm)
    ]

    if original_eq.empty or elegido_eq.empty:
        return formato_cantidad(cantidad_total_original), unidad_original

    cantidad_original_eq = float(original_eq.iloc[0]["cantidad_num"])
    cantidad_elegida_eq = float(elegido_eq.iloc[0]["cantidad_num"])
    unidad_elegida = str(elegido_eq.iloc[0]["unidad_eq"]).strip()

    if cantidad_original_eq == 0 or cantidad_elegida_eq == 0:
        return formato_cantidad(cantidad_total_original), unidad_original

    equivalentes = cantidad_total_original / cantidad_original_eq
    nueva_cantidad = equivalentes * cantidad_elegida_eq

    return formato_cantidad(nueva_cantidad), unidad_elegida


def formatear_opcion_platillo(platillo, mapa_nutriologa):
    if platillo == "—":
        return "—"

    estado = normalizar(mapa_nutriologa.get(platillo, ""))

    if estado == "si":
        return f"🟩 {platillo}"

    if estado == "no":
        return f"🟨 {platillo}"

    return platillo


# ---------------------------------------------------------
# CARGA DE DATOS
# ---------------------------------------------------------

data = cargar_excel()

platillos_df = data["platillos"]
ingredientes_df = data["ingredientes"]
preparaciones_df = data["preparaciones"]
equivalencias_df = data["equivalencias"]
hogar_df = data["hogar"]

platillos_df.columns = platillos_df.columns.str.strip().str.lower()
ingredientes_df.columns = ingredientes_df.columns.str.strip().str.lower()
preparaciones_df.columns = preparaciones_df.columns.str.strip().str.lower()
equivalencias_df.columns = equivalencias_df.columns.str.strip().str.lower()
hogar_df.columns = hogar_df.columns.str.strip().str.lower()

platillos_df["platillo"] = platillos_df["platillo"].apply(limpiar_texto)
platillos_df["tipo"] = platillos_df["tipo"].apply(limpiar_texto).str.lower()

if "nutriologa" not in platillos_df.columns:
    platillos_df["nutriologa"] = ""

platillos_df["nutriologa"] = platillos_df["nutriologa"].apply(limpiar_texto)

mapa_nutriologa = dict(zip(platillos_df["platillo"], platillos_df["nutriologa"]))

ingredientes_df["platillo"] = ingredientes_df["platillo"].apply(limpiar_texto)
ingredientes_df["ingrediente"] = ingredientes_df["ingrediente"].apply(limpiar_texto)
ingredientes_df["grupo"] = ingredientes_df["grupo"].apply(limpiar_texto)
ingredientes_df["unidad"] = ingredientes_df["unidad"].apply(limpiar_texto)

preparaciones_df["platillo"] = preparaciones_df["platillo"].apply(limpiar_texto)

if "cantidad_decimal" not in ingredientes_df.columns:
    ingredientes_df["cantidad_decimal"] = ingredientes_df["cantidad"].apply(convertir_a_numero)

ingredientes_df["cantidad_decimal"] = ingredientes_df["cantidad_decimal"].apply(convertir_a_numero)

if "nombre" not in hogar_df.columns:
    hogar_df["nombre"] = ""

if "categoría" not in hogar_df.columns and "categoria" in hogar_df.columns:
    hogar_df["categoría"] = hogar_df["categoria"]

if "categoría" not in hogar_df.columns:
    hogar_df["categoría"] = ""

hogar_df["nombre"] = hogar_df["nombre"].apply(limpiar_texto)
hogar_df["categoría"] = hogar_df["categoría"].apply(limpiar_texto)

hogar_df = hogar_df[hogar_df["nombre"] != ""].copy()
hogar_df = hogar_df.sort_values(["categoría", "nombre"])

equivalencias_limpias_global = preparar_equivalencias(equivalencias_df)

if "menu_pendiente_carga" in st.session_state:
    cargar_menu_en_session_state(pd.DataFrame(st.session_state["menu_pendiente_carga"]))
    del st.session_state["menu_pendiente_carga"]

if st.session_state.get("limpiar_menu_pendiente", False):
    limpiar_menu_session_state()
    st.session_state["limpiar_menu_pendiente"] = False


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
    st.sidebar.success("Menú reiniciado.")


# ---------------------------------------------------------
# TABS
# ---------------------------------------------------------

tab_menu, tab_ingredientes, tab_preparaciones, tab_equivalencias, tab_base, tab_hogar = st.tabs(
    [
        "📅 Menú semanal",
        "🛒 Ingredientes totales",
        "👩‍🍳 Preparaciones",
        "🔁 Equivalencias",
        "📋 Base",
        "🏠 Hogar",
    ]
)


# ---------------------------------------------------------
# TAB 1: MENÚ SEMANAL
# ---------------------------------------------------------

with tab_menu:
    st.subheader("📅 Menú semanal")

    st.markdown(
        """
        <div class="legend-box">
        🟩 Sugerencia de nutrióloga &nbsp;&nbsp; | &nbsp;&nbsp; 🟨 Platillo propio
        </div>
        """,
        unsafe_allow_html=True
    )

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

    for i, dia in enumerate(DIAS):
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

        for i, dia in enumerate(DIAS):
            key = f"menu_{comida_visible}_{dia}"
            valor_actual = st.session_state.get(key, "—")

            if valor_actual not in opciones:
                opciones_mostradas = ["—", valor_actual] + [
                    op for op in opciones if op not in ["—", valor_actual]
                ]
            else:
                opciones_mostradas = opciones

            index_default = opciones_mostradas.index(valor_actual) if valor_actual in opciones_mostradas else 0

            with row_cols[i + 1]:
                platillo_elegido = st.selectbox(
                    label=f"{comida_visible} {dia}",
                    options=opciones_mostradas,
                    index=index_default,
                    key=key,
                    label_visibility="collapsed",
                    format_func=lambda x, mapa=mapa_nutriologa: formatear_opcion_platillo(x, mapa)
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
            .groupby(["grupo", "ingrediente", "unidad"], as_index=False)
            .agg(
                cantidad_total=("cantidad_total", "sum"),
                platillos=("platillo", lambda x: ", ".join(sorted(set(x))))
            )
            .sort_values(["grupo", "ingrediente"])
        )

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
                ingrediente_original = row["ingrediente"]
                unidad_original = row["unidad"]
                cantidad_total = row["cantidad_total"]
                platillos_usados = row["platillos"]

                equivalentes = obtener_equivalentes(equivalencias_limpias_global, grupo)

                if ingrediente_original not in equivalentes:
                    opciones_equivalentes = [ingrediente_original] + equivalentes
                else:
                    opciones_equivalentes = equivalentes

                opciones_equivalentes = list(dict.fromkeys(opciones_equivalentes))

                key_select = f"equiv_{grupo}_{ingrediente_original}_{unidad_original}"
                key_checkbox = f"check_ingrediente_{grupo}_{ingrediente_original}_{unidad_original}"

                col_check, col_info, col_select = st.columns([0.15, 2.5, 1.7])

                with col_select:
                    ingrediente_elegido = st.selectbox(
                        "Equivalente",
                        opciones_equivalentes,
                        index=opciones_equivalentes.index(ingrediente_original)
                        if ingrediente_original in opciones_equivalentes else 0,
                        key=key_select,
                        label_visibility="collapsed"
                    )

                cantidad_mostrar, unidad_mostrar = calcular_sustitucion(
                    equivalencias_limpias_global,
                    grupo,
                    ingrediente_original,
                    cantidad_total,
                    unidad_original,
                    ingrediente_elegido
                )

                with col_check:
                    marcado = st.checkbox(
                        "",
                        key=key_checkbox,
                        label_visibility="collapsed"
                    )

                with col_info:
                    texto_ingrediente = f"**{ingrediente_elegido}** {cantidad_mostrar} {unidad_mostrar}"
                    st.markdown(
                        f"{texto_ingrediente} "
                        f"<span class='ingredient-note'>({html.escape(platillos_usados)})</span>",
                        unsafe_allow_html=True
                    )

                if not marcado:
                    ingredientes_faltantes.append(
                        {
                            "grupo": grupo,
                            "ingrediente": ingrediente_elegido,
                            "cantidad": cantidad_mostrar,
                            "unidad": unidad_mostrar,
                            "platillos": platillos_usados
                        }
                    )

            st.divider()

        st.subheader("🧾 Ingredientes faltantes")

        if len(ingredientes_faltantes) == 0:
            st.success("Ya marcaste todos los ingredientes como listos ✅")
        else:
            faltantes_df = pd.DataFrame(ingredientes_faltantes)
            faltantes_df = consolidar_ingredientes_faltantes(faltantes_df)

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

    if equivalencias_limpias_global.empty:
        st.error(
            "No encontré las columnas necesarias en la hoja de equivalencias. "
            "Necesito: categoría, ingrediente y cantidad."
        )
    else:
        equivalencias_tabla = equivalencias_limpias_global[
            ["categoría", "ingrediente", "cantidad"]
        ].copy()

        equivalencias_tabla = equivalencias_tabla.sort_values(
            by=["categoría", "ingrediente"]
        )

        categorias = sorted(equivalencias_tabla["categoría"].dropna().unique().tolist())

        for categoria in categorias:
            st.markdown(f"## {categoria}")

            tabla_categoria = equivalencias_tabla[
                equivalencias_tabla["categoría"] == categoria
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
                st.session_state["menu_pendiente_carga"] = menu_cargado_df.to_dict("records")
                st.success("Menú cargado. Actualizando...")
                st.rerun()

    with col_limpiar:
        if st.button("Limpiar menú actual"):
            st.session_state["limpiar_menu_pendiente"] = True
            st.success("Menú limpiado. Actualizando...")
            st.rerun()


# ---------------------------------------------------------
# TAB 6: HOGAR
# ---------------------------------------------------------

with tab_hogar:
    st.subheader("🏠 Hogar")

    st.write(
        "Lista general de productos de casa. Puedes filtrar por categoría para armar listas rápidas."
    )

    if hogar_df.empty:
        st.warning("No encontré información en la hoja `Hoja 7 - Hogar`.")
    else:
        categorias_hogar = sorted(hogar_df["categoría"].dropna().unique().tolist())

        categoria_elegida = st.selectbox(
            "Filtrar por categoría:",
            ["Todas"] + categorias_hogar
        )

        if categoria_elegida == "Todas":
            hogar_filtrado = hogar_df.copy()
        else:
            hogar_filtrado = hogar_df[hogar_df["categoría"] == categoria_elegida].copy()

        hogar_filtrado = hogar_filtrado[["nombre", "categoría"]].copy()
        hogar_filtrado.columns = ["Nombre", "Categoría"]

        st.dataframe(
            hogar_filtrado,
            use_container_width=True,
            hide_index=True
        )
