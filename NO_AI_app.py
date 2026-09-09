import io
import re
import datetime
import unicodedata

import pandas as pd
import requests
import streamlit as st


st.set_page_config(page_title="Tiempos de Taller OMODA & JAECOO", layout="wide")

URL_GITHUB_EXCEL = "https://github.com/MrFeudo/Catalogo-Operaciones/raw/main/DMS_Active_Spare_Parts.xlsb"
URL_GITHUB_VINES = "https://github.com/MrFeudo/Catalogo-Operaciones/raw/main/VINes.xlsb"
URL_FORMULARIO_LARK = "https://omodaeurope.jp.larksuite.com/share/base/form/shrjpjuwHJs8xglcJAIQ0v2Mnvf"
QR_FORMULARIO = "QR_FORMULARIO_OPS_GARANTIA.png"
SUGGESTIONS_WORKSHEET = "Sugerencias"

DEFAULT_SESSION_VALUES = {
    "authenticated": False,
    "idioma": "Español",
    "ultimo_vin_catalogo": "",
}
for key, value in DEFAULT_SESSION_VALUES.items():
    if key not in st.session_state:
        st.session_state[key] = value

IDIOMAS = {
    "Español": {
        "menu_titulo": "### 🗺️ Menú de Navegación",
        "menu_radio": "Selecciona una herramienta:",
        "menu_taller": "📋 Tiempos de Taller",
        "menu_solicitar": "📝 Operaciones no disponibles",
        "menu_buzon": "📮 Buzón de sugerencias",
        "pass_titulo": "🔐 Acceso Red de Dealers",
        "pass_input": "Introduce la contraseña de acceso:",
        "pass_boton": "Entrar",
        "pass_error": "❌ Contraseña incorrecta",
        "taller_titulo": "🚗 Catálogo Operaciones de mano de obra",
        "taller_sub": "Consulta piezas, modelos y tiempos asignados directamente desde el DMS.",
        "f_modelo": "1. Filtrar por Modelo:",
        "f_pieza": "2. Buscar por Nombre o Código de pieza:",
        "f_operacion": "3. Buscar por tipo de operación (ej: Remove, Paint...):",
        "f_mercado_taller": "Filtrar por Mercado / Organización (Taller):",
        "f_estado_taller": "Filtrar por Estado de Operación (Taller):",
        "res_taller": "### 📋 Resultados encontrados: {} operaciones",
        "warn_taller": "⚠️ No se encontraron operaciones con los criterios seleccionados.",
        "err_taller": "Error al procesar la base de datos de tiempos: {}",
        "todos": "Todos",
        "solicitar_titulo": "📝 Operaciones no disponibles",
        "solicitar_sub": "Si has buscado la operación y no aparece en el catálogo, repórtala mediante el formulario oficial.",
        "solicitar_btn": "📝 Abrir formulario de operaciones no disponibles",
        "solicitar_qr": "📱 O escanea el QR desde tu móvil",
        "solicitar_recordatorio": "Antes de reportarla, comprueba la operación por VIN, modelo, pieza y tipo de operación en el catálogo.",
        "buzon_titulo": "📮 Buzón de sugerencias",
        "buzon_sub": "Utiliza este espacio para proponer mejoras, reportar problemas o compartir ideas sobre la herramienta y el proceso.",
        "buzon_privacidad": "🔒 No se solicitan datos personales ni del concesionario.",
        "buzon_categoria": "Tipo de comentario *",
        "buzon_area": "¿A qué apartado se refiere? *",
        "buzon_comentario": "Comentario *",
        "buzon_placeholder": "Cuéntanos qué cambiarías, qué problema has encontrado o qué podría hacerse mejor...",
        "buzon_enviar": "Enviar sugerencia",
        "buzon_ok": "✅ Gracias. Tu sugerencia se ha registrado correctamente.",
        "buzon_error_vacio": "❌ Escribe un comentario antes de enviarlo.",
    },
    "English": {
        "menu_titulo": "### 🗺️ Navigation Menu",
        "menu_radio": "Select a tool:",
        "menu_taller": "📋 Workshop Times",
        "menu_solicitar": "📝 Missing Operations",
        "menu_buzon": "📮 Suggestions Box",
        "pass_titulo": "🔐 Dealer Network Access",
        "pass_input": "Enter access password:",
        "pass_boton": "Login",
        "pass_error": "❌ Incorrect password",
        "taller_titulo": "🚗 Labor Operations Catalog",
        "taller_sub": "Consult parts, models and assigned times directly from the DMS.",
        "f_modelo": "1. Filter by Model:",
        "f_pieza": "2. Search by Part Name or Code:",
        "f_operacion": "3. Search by operation type:",
        "f_mercado_taller": "Filter by Market / Organization:",
        "f_estado_taller": "Filter by Operation Status:",
        "res_taller": "### 📋 Results found: {} operations",
        "warn_taller": "⚠️ No operations found matching the selected criteria.",
        "err_taller": "Error processing workshop times database: {}",
        "todos": "All",
        "solicitar_titulo": "📝 Missing Operations",
        "solicitar_sub": "If you searched for the operation and it is not available in the catalog, report it using the official form.",
        "solicitar_btn": "📝 Open Missing Operations form",
        "solicitar_qr": "📱 Or scan the QR code from your phone",
        "solicitar_recordatorio": "Before reporting it, check the operation by VIN, model, part and operation type in the catalog.",
        "buzon_titulo": "📮 Suggestions Box",
        "buzon_sub": "Use this space to suggest improvements, report issues or share ideas about the tool and the process.",
        "buzon_privacidad": "🔒 No personal or dealer identification data is requested.",
        "buzon_categoria": "Type of comment *",
        "buzon_area": "Which area does it refer to? *",
        "buzon_comentario": "Comment *",
        "buzon_placeholder": "Tell us what you would change, what problem you found or what could be improved...",
        "buzon_enviar": "Send suggestion",
        "buzon_ok": "✅ Thank you. Your suggestion has been registered.",
        "buzon_error_vacio": "❌ Please enter a comment before submitting.",
    },
    "Chinese (中文)": {
        "menu_titulo": "### 🗺️ 导航菜单",
        "menu_radio": "选择工具:",
        "menu_taller": "📋 车间工时",
        "menu_solicitar": "📝 缺失操作",
        "menu_buzon": "📮 建议箱",
        "pass_titulo": "🔐 经销商网络访问",
        "pass_input": "输入访问密码:",
        "pass_boton": "登录",
        "pass_error": "❌ 密码错误",
        "taller_titulo": "🚗 工时操作目录",
        "taller_sub": "直接从 DMS 查询零件、车型和分配的时间。",
        "f_modelo": "1. 按车型筛选:",
        "f_pieza": "2. 按零件名称或代码搜索:",
        "f_operacion": "3. 按操作类型搜索:",
        "f_mercado_taller": "按市场 / 组织筛选:",
        "f_estado_taller": "按操作状态筛选:",
        "res_taller": "### 📋 找到的结果: {} 个操作",
        "warn_taller": "⚠️ 未找到符合条件的工时操作。",
        "err_taller": "处理车间工时数据库时出错: {}",
        "todos": "全部",
        "solicitar_titulo": "📝 缺失操作",
        "solicitar_sub": "如果已搜索但目录中没有该操作，请使用官方表单进行报告。",
        "solicitar_btn": "📝 打开缺失操作表单",
        "solicitar_qr": "📱 或使用手机扫描二维码",
        "solicitar_recordatorio": "报告前，请先按 VIN、车型、零件和操作类型在目录中进行检查。",
        "buzon_titulo": "📮 建议箱",
        "buzon_sub": "可在此提出改进建议、报告问题或分享有关工具和流程的想法。",
        "buzon_privacidad": "🔒 不要求提供个人或经销商身份信息。",
        "buzon_categoria": "意见类型 *",
        "buzon_area": "涉及哪个部分？ *",
        "buzon_comentario": "意见 *",
        "buzon_placeholder": "请告诉我们你会如何改进、遇到了什么问题或有什么建议...",
        "buzon_enviar": "提交建议",
        "buzon_ok": "✅ 谢谢。你的建议已成功记录。",
        "buzon_error_vacio": "❌ 提交前请输入意见。",
    },
}


def quitar_acentos(texto):
    texto = str(texto)
    return "".join(c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn")


def normalizar_modelo(modelo):
    modelo = quitar_acentos(modelo).upper().strip()
    modelo = modelo.replace("_", " ").replace("-", " ")
    modelo = re.sub(r"\s+", " ", modelo).strip()
    modelo = re.sub(r"\bMY\s*25\b", "", modelo)
    modelo = re.sub(r"\bMY\s*26\b", "", modelo)
    modelo = re.sub(r"\bMY\s*2025\b", "", modelo)
    modelo = re.sub(r"\bMY\s*2026\b", "", modelo)
    modelo = re.sub(r"\(\s*GASOLINA\s*\)", "", modelo)
    modelo = re.sub(r"\(\s*HIBRIDO\s*\)", "", modelo)
    modelo = re.sub(r"\(\s*ELECTRICO\s*\)", "", modelo)
    modelo = re.sub(r"\bBEV\b", "EV", modelo)
    return re.sub(r"\s+", " ", modelo).strip()


def obtener_tipo_propulsion(modelo):
    modelo_normalizado = normalizar_modelo(modelo)
    if re.search(r"\bPHEV\b", modelo_normalizado):
        return "PHEV"
    if re.search(r"\bHEV\b", modelo_normalizado):
        return "HEV"
    if re.search(r"\bEV\b", modelo_normalizado):
        return "EV"
    return "GASOLINA"


def buscar_modelo_equivalente(modelo_vin, modelos_candidatos, valor_todos=None):
    modelo_vin_normalizado = normalizar_modelo(modelo_vin)
    propulsion_vin = obtener_tipo_propulsion(modelo_vin)
    for candidato in modelos_candidatos:
        if valor_todos is not None and candidato == valor_todos:
            continue
        candidato_normalizado = normalizar_modelo(candidato)
        propulsion_candidato = obtener_tipo_propulsion(candidato)
        if propulsion_vin != propulsion_candidato:
            continue
        if candidato_normalizado == modelo_vin_normalizado:
            return candidato
    return None


@st.cache_data(ttl=600)
def load_data_vines():
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        response = requests.get(URL_GITHUB_VINES, headers=headers, timeout=15)
        if response.status_code != 200:
            st.sidebar.error(f"⚠️ Error HTTP {response.status_code} descargando VINes.xlsb")
            return pd.DataFrame(columns=["VIN", "Modelo_Excel"])
        file_bytes = io.BytesIO(response.content)
        df_vines = pd.read_excel(file_bytes, engine="pyxlsb")
        df_vines.columns = df_vines.columns.astype(str).str.strip()
        col_vin = "new_name"
        col_modelo = "new_productmodel_idname"
        if col_vin not in df_vines.columns or col_modelo not in df_vines.columns:
            st.sidebar.error("⚠️ No se encontraron las columnas esperadas en VINes.xlsb.")
            return pd.DataFrame(columns=["VIN", "Modelo_Excel"])
        df_clean = df_vines[[col_vin, col_modelo]].dropna(subset=[col_vin, col_modelo]).copy()
        df_clean.columns = ["VIN", "Modelo_Excel"]
        df_clean["VIN"] = df_clean["VIN"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip().str.upper()
        df_clean["Modelo_Excel"] = df_clean["Modelo_Excel"].astype(str).str.strip()
        return df_clean.drop_duplicates(subset=["VIN"], keep="first").reset_index(drop=True)
    except Exception as exc:
        st.sidebar.error(f"⚠️ Excepción al leer VINes.xlsb: {exc}")
        return pd.DataFrame(columns=["VIN", "Modelo_Excel"])


def obtener_modelo_desde_vin(vin, df_vines_db):
    vin = str(vin).strip().upper()
    if len(vin) != 17 or df_vines_db.empty:
        return None
    coincidencia = df_vines_db[df_vines_db["VIN"] == vin]
    if coincidencia.empty:
        return None
    return str(coincidencia.iloc[0]["Modelo_Excel"]).strip()


def render_sidebar_and_get_option():
    try:
        st.sidebar.image("logo_empresa.png", use_container_width=True)
    except Exception:
        st.sidebar.write("🏢 **OMODA & JAECOO**")
    st.sidebar.markdown("---")
    idiomas_disponibles = ["Español", "English", "Chinese (中文)"]
    idioma_seleccionado = st.sidebar.selectbox(
        "🌐 Language / Idioma / 语言:",
        idiomas_disponibles,
        index=idiomas_disponibles.index(st.session_state.idioma),
        key="selector_idioma_global",
    )
    st.session_state.idioma = idioma_seleccionado
    txt_local = IDIOMAS[st.session_state.idioma]
    st.sidebar.markdown("---")
    st.sidebar.markdown(txt_local["menu_titulo"])
    opciones = [txt_local["menu_taller"], txt_local["menu_solicitar"], txt_local["menu_buzon"]]
    opcion = st.sidebar.radio(txt_local["menu_radio"], opciones, key="menu_navegacion_app")
    return txt_local, opcion


def check_password(txt_local):
    if not st.session_state.authenticated:
        st.title(txt_local["pass_titulo"])
        password = st.text_input(txt_local["pass_input"], type="password", key="pass_input_unico")
        if st.button(txt_local["pass_boton"], key="pass_btn_unico"):
            if password == "DealersOJ2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error(txt_local["pass_error"])
        return False
    return True


@st.cache_data
def load_data_tiempos_v3():
    df = pd.read_excel(URL_GITHUB_EXCEL, sheet_name="new_srv_workhours", engine="pyxlsb")
    df.columns = df.columns.astype(str).str.strip()
    mapeo_columnas = {
        "new_productmodel_idname": "Modelo",
        "new_product_idname": "Nombre de la Pieza",
        "new_code": "Código de Referencia",
        "new_name": "Operación Técnica",
        "new_standardhour": "Tiempo Estándar (UT/Horas)",
        "new_remark": "Notas / Exclusiones",
        "Organization": "Mercado / Organización",
        "statecodename": "Estado",
    }
    cols_existentes = [col for col in mapeo_columnas if col in df.columns]
    df_limpio = df[cols_existentes].copy().rename(columns=mapeo_columnas)
    df_limpio = df_limpio.replace(to_replace=r"^0x.*$", value="", regex=True).fillna("").replace(["nan", "None", "NaN"], "")
    columnas_finales = [
        "Modelo",
        "Nombre de la Pieza",
        "Código de Referencia",
        "Operación Técnica",
        "Tiempo Estándar (UT/Horas)",
        "Notas / Exclusiones",
        "Mercado / Organización",
        "Estado",
    ]
    columnas_presentes = [col for col in columnas_finales if col in df_limpio.columns]
    return df_limpio[columnas_presentes].reset_index(drop=True)


def render_tiempos_taller(txt_local):
    try:
        data = load_data_tiempos_v3()
        df_vines_db = load_data_vines()
        st.title(txt_local["taller_titulo"])
        st.write(txt_local["taller_sub"])
        st.markdown("---")
        st.subheader("📊 Catálogo de operaciones")

        modelos_raw = [str(m).strip() for m in data["Modelo"].dropna().unique()] if "Modelo" in data.columns else []
        modelos_filtrados = [modelo for modelo in modelos_raw if any(marca in modelo.upper() for marca in ["OMODA", "JAECOO", "LEPAS"])]
        modelos_disponibles = [txt_local["todos"]] + sorted(list(set(modelos_filtrados)))

        col_vin, col1, col2, col3 = st.columns([1.5, 1.2, 1.5, 1.5])
        with col_vin:
            vin_busqueda = st.text_input("🔎 Buscar por VIN (Bastidor):", max_chars=17, placeholder="17 caracteres...", key="vin_taller_input").strip().upper()

        modelo_detectado_por_vin = None
        if len(vin_busqueda) == 17:
            modelo_raw_vin = obtener_modelo_desde_vin(vin_busqueda, df_vines_db)
            if modelo_raw_vin:
                modelo_detectado_por_vin = buscar_modelo_equivalente(modelo_raw_vin, modelos_disponibles, valor_todos=txt_local["todos"])
                if modelo_detectado_por_vin:
                    st.info(f"🚘 **VIN Detectado:** {vin_busqueda} ➔ **Modelo:** {modelo_detectado_por_vin}")
                else:
                    st.warning(f"⚠️ Bastidor localizado como **{modelo_raw_vin}**, pero no existe una coincidencia exacta en el catálogo de operaciones.")
            else:
                st.error("❌ Bastidor VIN no encontrado en la base de datos VINes.xlsb.")

        if vin_busqueda != st.session_state.ultimo_vin_catalogo:
            st.session_state.ultimo_vin_catalogo = vin_busqueda
            if modelo_detectado_por_vin:
                st.session_state["sb_modelo_taller_select"] = modelo_detectado_por_vin

        with col1:
            modelo_seleccionado = st.selectbox(txt_local["f_modelo"], modelos_disponibles, key="sb_modelo_taller_select")
        with col2:
            buscar_pieza = st.text_input(txt_local["f_pieza"], "").strip()
        with col3:
            buscar_operacion = st.text_input(txt_local["f_operacion"], "").strip()

        col_m, col_e = st.columns([2, 2])
        with col_m:
            if "Mercado / Organización" in data.columns:
                mercados_disponibles = [txt_local["todos"]] + [str(m).strip() for m in data["Mercado / Organización"].unique() if str(m).strip() != ""]
                indice_defecto = next((idx for idx, mercado in enumerate(mercados_disponibles) if "spain" in mercado.lower() or "oj spain" in mercado.lower()), 0)
                mercado_seleccionado = st.selectbox(txt_local["f_mercado_taller"], mercados_disponibles, index=indice_defecto)
            else:
                mercado_seleccionado = txt_local["todos"]
        with col_e:
            if "Estado" in data.columns:
                estados_disponibles = [txt_local["todos"]] + [str(e).strip() for e in data["Estado"].unique() if str(e).strip() != ""]
                indice_est_defecto = estados_disponibles.index("Active") if "Active" in estados_disponibles else 0
                estado_seleccionado = st.selectbox(txt_local["f_estado_taller"], estados_disponibles, index=indice_est_defecto)
            else:
                estado_seleccionado = txt_local["todos"]

        df_filtrado = data.copy()
        if modelo_seleccionado != txt_local["todos"] and "Modelo" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Modelo"] == modelo_seleccionado]
        if mercado_seleccionado != txt_local["todos"] and "Mercado / Organización" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Mercado / Organización"].astype(str).str.strip() == mercado_seleccionado]
        if estado_seleccionado != txt_local["todos"] and "Estado" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Estado"].astype(str).str.strip() == estado_seleccionado]
        if buscar_pieza and {"Nombre de la Pieza", "Código de Referencia"}.issubset(df_filtrado.columns):
            df_filtrado = df_filtrado[
                df_filtrado["Nombre de la Pieza"].astype(str).str.contains(buscar_pieza, case=False, na=False)
                | df_filtrado["Código de Referencia"].astype(str).str.contains(buscar_pieza, case=False, na=False)
            ]
        if buscar_operacion and "Operación Técnica" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Operación Técnica"].astype(str).str.contains(buscar_operacion, case=False, na=False)]

        st.markdown(txt_local["res_taller"].format(len(df_filtrado)))
        if not df_filtrado.empty:
            st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
        else:
            st.warning(txt_local["warn_taller"])
            st.markdown("### ¿La operación que necesitas no aparece?")
            st.write("Repórtala mediante el formulario oficial de Operaciones no disponibles.")
            st.link_button("📝 Reportar operación no disponible", URL_FORMULARIO_LARK, use_container_width=True)
    except Exception as exc:
        st.error(txt_local["err_taller"].format(exc))


def render_operaciones_no_disponibles(txt_local):
    st.title(txt_local["solicitar_titulo"])
    st.write(txt_local["solicitar_sub"])
    st.markdown("---")
    st.info(txt_local["solicitar_recordatorio"])
    st.link_button(txt_local["solicitar_btn"], URL_FORMULARIO_LARK, use_container_width=True)
    st.markdown("---")
    st.subheader(txt_local["solicitar_qr"])
    col1, col2, col3 = st.columns([1, 1.3, 1])
    with col2:
        try:
            st.image(QR_FORMULARIO, use_container_width=True)
        except Exception:
            st.warning("⚠️ No se ha encontrado la imagen del QR. Comprueba que `QR_FORMULARIO_OPS_GARANTIA.png` está en el repositorio.")


def obtener_url_google_sheets():
    if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
        return st.secrets["connections"]["gsheets"]["spreadsheet"]
    if "gsheets" in st.secrets and "spreadsheet" in st.secrets["gsheets"]:
        return st.secrets["gsheets"]["spreadsheet"]
    return st.secrets.get("spreadsheet", "")


def render_buzon_sugerencias(txt_local):
    st.title(txt_local["buzon_titulo"])
    st.write(txt_local["buzon_sub"])
    st.info(txt_local["buzon_privacidad"])
    st.markdown("---")

    if st.session_state.idioma == "Español":
        categorias = ["Mejora de la aplicación", "Mejora del proceso", "Error o problema encontrado", "Idea / sugerencia", "Otro"]
        areas = ["Tiempos de taller", "Operaciones no disponibles", "DMS", "Garantías", "Otro"]
    elif st.session_state.idioma == "English":
        categorias = ["Application improvement", "Process improvement", "Error / issue found", "Idea / suggestion", "Other"]
        areas = ["Workshop Times", "Missing Operations", "DMS", "Warranty", "Other"]
    else:
        categorias = ["应用改进", "流程改进", "发现错误 / 问题", "想法 / 建议", "其他"]
        areas = ["车间工时", "缺失操作", "DMS", "保修", "其他"]

    with st.form("form_sugerencias", clear_on_submit=True):
        categoria = st.selectbox(txt_local["buzon_categoria"], categorias)
        area = st.selectbox(txt_local["buzon_area"], areas)
        comentario = st.text_area(txt_local["buzon_comentario"], placeholder=txt_local["buzon_placeholder"], height=170)
        enviar = st.form_submit_button(txt_local["buzon_enviar"], use_container_width=True)

        if enviar:
            if not comentario.strip():
                st.error(txt_local["buzon_error_vacio"])
                return

            nueva_sugerencia = {
                "Fecha": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "Categoría": categoria,
                "Área": area,
                "Comentario": comentario.strip(),
            }

            try:
                from streamlit_gsheets import GSheetsConnection

                spreadsheet_url = obtener_url_google_sheets()
                if not spreadsheet_url:
                    raise ValueError("No se encontró la URL del Google Sheets en st.secrets.")

                conn = st.connection("gsheets", type=GSheetsConnection)
                columnas = ["Fecha", "Categoría", "Área", "Comentario"]

                try:
                    df_actual = conn.read(spreadsheet=spreadsheet_url, worksheet=SUGGESTIONS_WORKSHEET)
                except Exception:
                    df_actual = pd.DataFrame(columns=columnas)

                if df_actual.empty:
                    df_actual = pd.DataFrame(columns=columnas)
                else:
                    df_actual = df_actual.dropna(how="all").loc[:, ~df_actual.columns.astype(str).str.contains("^Unnamed")]
                    for columna in columnas:
                        if columna not in df_actual.columns:
                            df_actual[columna] = ""
                    df_actual = df_actual[columnas]

                df_nuevo = pd.DataFrame([nueva_sugerencia])
                df_actualizado = pd.concat([df_actual, df_nuevo], ignore_index=True)
                conn.update(spreadsheet=spreadsheet_url, worksheet=SUGGESTIONS_WORKSHEET, data=df_actualizado)
                st.success(txt_local["buzon_ok"])

            except Exception as exc:
                st.error(f"❌ No se ha podido enviar la sugerencia: {exc}")
                st.info(f"💡 Comprueba que exista una pestaña llamada `{SUGGESTIONS_WORKSHEET}` en el Google Sheets.")


txt, opcion_menu = render_sidebar_and_get_option()

if check_password(txt):
    if opcion_menu == txt["menu_taller"]:
        render_tiempos_taller(txt)
    elif opcion_menu == txt["menu_solicitar"]:
        render_operaciones_no_disponibles(txt)
    elif opcion_menu == txt["menu_buzon"]:
        render_buzon_sugerencias(txt)
