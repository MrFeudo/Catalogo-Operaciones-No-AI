import io
import re
import time
import datetime

import pandas as pd
import requests
import streamlit as st



# =========================================================================
# CONFIGURACIÓN GENERAL
# =========================================================================
st.set_page_config(page_title="Tiempos de Taller OMODA & JAECOO", layout="wide")

URL_GITHUB_EXCEL = "https://github.com/MrFeudo/Catalogo-Operaciones/raw/main/DMS_Active_Spare_Parts.xlsb"
URL_GITHUB_VINES = "https://github.com/MrFeudo/Catalogo-Operaciones/raw/main/VINes.xlsb"



# =========================================================================
# INICIALIZACIÓN SESSION STATE
# =========================================================================
DEFAULT_SESSION_VALUES = {
    "lista_solicitudes": [],
    "authenticated": False,
    "idioma": "Español",
    "solicitar_marca": "OMODA",
    "solicitar_modelo": "OMODA 5 (Gasolina)"
}

for key, value in DEFAULT_SESSION_VALUES.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================================
# DICCIONARIO DE TRADUCCIÓN
# =========================================================================
IDIOMAS = {
    "Español": {
        "menu_titulo": "### 🗺️ Menú de Navegación",
        "menu_radio": "Selecciona una herramienta:",
        "menu_taller": "📋 Tiempos de Taller",
        "menu_solicitar": "📝 Operaciones no disponibles",
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
        "solicitar_sub": "Usa este formulario para indicarnos una operación de taller que necesitas y no encuentras en el sistema.",
        "form_sub": "Datos de la Solicitud (Campos obligatorios *)",
        "form_marca": "Marca del vehículo *",
        "form_modelo": "INTRODUCIR MODELO *",
        "form_vin": "INTRODUCIR VIN (Bastidor) *",
        "form_vin_holder": "17 caracteres",
        "form_hq_code": "CÓDIGO DE PRODUCTO (Asignado por HQ)",
        "form_ref": "REFERENCIA DE PIEZA (Opcional)",
        "form_ref_holder": "Ej. 7365747465AA",
        "form_op": "OPERACIÓN QUE NO ENCUENTRAS *",
        "form_op_holder": "Describe la operación de taller que necesitas y no encuentras en el sistema...",
        "form_btn": "Enviar Solicitud a Central",
        "err_campos": "❌ Por favor, rellene todos los campos obligatorios (*).",
    },
    "English": {
        "menu_titulo": "### 🗺️ Navigation Menu",
        "menu_radio": "Select a tool:",
        "menu_taller": "📋 Workshop Times",
        "menu_solicitar": "📝 Missing Operations",
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
        "solicitar_titulo": "📝 Request for Additional Labor Operations",
        "solicitar_sub": "Use this form to request new operations or prices to be added to HQ master list.",
        "form_sub": "Request Details (* Required fields)",
        "form_marca": "Vehicle Brand *",
        "form_modelo": "ENTER MODEL *",
        "form_vin": "ENTER VIN (Chassis) *",
        "form_vin_holder": "17 characters",
        "form_hq_code": "PRODUCT CODE (Assigned by HQ)",
        "form_ref": "PART REFERENCE (Optional)",
        "form_ref_holder": "e.g., 7365747465AA",
        "form_op": "OPERATION REQUESTED TO BE ADDED *",
        "form_op_holder": "Describe in detail the technical operation or missing price required by the workshop...",
        "form_btn": "Send Request to HQ",
        "err_campos": "❌ Please fill in all required fields (*).",
    },
    "Chinese (中文)": {
        "menu_titulo": "### 🗺️ 导航菜单",
        "menu_radio": "选择工具:",
        "menu_taller": "📋 车间工时",
        "menu_solicitar": "📝 缺失操作",
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
        "warn_taller": "⚠️ 未找到符合选择条件的工时操作。",
        "err_taller": "处理车间工时数据库时出错: {}",
        "todos": "全部",
        "solicitar_titulo": "📝 申请新增工时操作",
        "solicitar_sub": "使用此表单申请在总部主数据中添加新工时操作或价格。",
        "form_sub": "申请信息 (* 为必填项)",
        "form_marca": "车辆品牌 *",
        "form_modelo": "输入车型 *",
        "form_vin": "输入 VIN (车架号) *",
        "form_vin_holder": "17位字符",
        "form_hq_code": "产品代码 (由总部分配)",
        "form_ref": "零件编号 (选填)",
        "form_ref_holder": "例如: 7365747465AA",
        "form_op": "申请添加的操作内容 *",
        "form_op_holder": "请详细描述车间所需的工时操作或缺失的价格...",
        "form_btn": "发送申请至总部",
        "err_campos": "❌ 请填写所有必填项 (*)。",
    }
}



# =========================================================================
# UTILIDADES Y CARGA CORREGIDA DE VINES (COLUMNAS DEFINIDAS EXPLICITAMENTE)
# =========================================================================
def normalizar_texto(texto):
    texto = str(texto)
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn').lower()

def normalize_text(text):
    return normalizar_texto(text)

@st.cache_data(ttl=600)
def load_data_vines():
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        response = requests.get(URL_GITHUB_VINES, headers=headers, timeout=15)
        
        if response.status_code == 200:
            file_bytes = io.BytesIO(response.content)
            df_vines = pd.read_excel(file_bytes, engine="pyxlsb")
            df_vines.columns = df_vines.columns.astype(str).str.strip()
            
            # Búsqueda explícita con tus columnas de Excel: new_name (VIN) y new_productmodel_idname (Modelo)
            col_vin = next((c for c in df_vines.columns if c.lower() == 'new_name' or 'vin' in c.lower() or 'bastidor' in c.lower()), None)
            col_modelo = next((c for c in df_vines.columns if c.lower() == 'new_productmodel_idname' or 'model' in c.lower()), None)
            
            if col_vin and col_modelo:
                df_clean = df_vines[[col_vin, col_modelo]].dropna().copy()
                df_clean.columns = ['VIN', 'Modelo_Excel']
                
                # Saneamiento del VIN
                df_clean['VIN'] = df_clean['VIN'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip().str.upper()
                df_clean['Modelo_Excel'] = df_clean['Modelo_Excel'].astype(str).str.strip()
                return df_clean
            else:
                st.sidebar.error(f"⚠️ Columnas esperadas no encontradas. Columnas en Excel: {list(df_vines.columns)}")
        else:
            st.sidebar.error(f"⚠️ Error HTTP {response.status_code} descargando VINes.xlsb")
    except Exception as exc:
        st.sidebar.error(f"⚠️ Excepción al leer VINes.xlsb: {exc}")
    
    return pd.DataFrame(columns=['VIN', 'Modelo_Excel'])




# =========================================================================
# AUTENTICACIÓN Y SIDEBAR
# =========================================================================
def render_sidebar_and_get_option():
    try:
        st.sidebar.image("logo_empresa.png", use_container_width=True)
    except Exception:
        st.sidebar.write("🏢 **OMODA & JAECOO**")

    st.sidebar.markdown("---")

    idioma_seleccionado = st.sidebar.selectbox(
        "🌐 Language / Idioma / 语言:", ["Español", "English", "Chinese (中文)"],
        index=["Español", "English", "Chinese (中文)"].index(st.session_state.idioma),
        key="selector_idioma_global"
    )
    st.session_state.idioma = idioma_seleccionado
    txt_local = IDIOMAS[st.session_state.idioma]

    st.sidebar.markdown("---")
    st.sidebar.markdown(txt_local["menu_titulo"])

    opciones = [txt_local["menu_taller"], txt_local["menu_solicitar"]]
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


# =========================================================================
# PANTALLA 1 - TIEMPOS DE TALLER
# =========================================================================
@st.cache_data
def load_data_tiempos_v3():
    df = pd.read_excel(URL_GITHUB_EXCEL, sheet_name="new_srv_workhours")
    df.columns = df.columns.astype(str).str.strip()

    mapeo_columnas = {
        "new_productmodel_idname": "Modelo", "new_product_idname": "Nombre de la Pieza",
        "new_code": "Código de Referencia", "new_name": "Operación Técnica",
        "new_standardhour": "Tiempo Estándar (UT/Horas)", "new_remark": "Notas / Exclusiones",
        "Organization": "Mercado / Organización", "statecodename": "Estado",
    }

    cols_existentes = [col for col in mapeo_columnas if col in df.columns]
    df_limpio = df[cols_existentes].copy().rename(columns=mapeo_columnas)
    df_limpio = df_limpio.replace(to_replace=r"^0x.*$", value="", regex=True).fillna("").replace(["nan", "None", "NaN"], "")

    columnas_finales = ["Modelo", "Nombre de la Pieza", "Código de Referencia", "Operación Técnica", "Tiempo Estándar (UT/Horas)", "Notas / Exclusiones", "Mercado / Organización", "Estado"]
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
        st.markdown("---")
        st.subheader("📊 Catálogo Completo (Filtros Manuales y Detección por VIN)")

        modelos_raw = [str(m).strip() for m in data["Modelo"].dropna().unique()] if "Modelo" in data.columns else []
        modelos_filtrados = [m for m in modelos_raw if any(marca in m.upper() for marca in ["OMODA", "JAECOO", "LEPAS"])]
        modelos_disponibles = [txt_local["todos"]] + sorted(list(set(modelos_filtrados)))

        col_vin, col1, col2, col3 = st.columns([1.5, 1.2, 1.5, 1.5])

        with col_vin:
            vin_busqueda = st.text_input(
                "🔎 Buscar por VIN (Bastidor):", 
                max_chars=17, 
                placeholder="17 caracteres...",
                key="vin_taller_input"
            ).strip().upper()

        # EVALUACIÓN DIRECTA DEL VIN EN CADA RENDER
        modelo_detectado_por_vin = None
        if len(vin_busqueda) == 17:
            if not df_vines_db.empty:
                coincidencia = df_vines_db[df_vines_db['VIN'] == vin_busqueda]
                if not coincidencia.empty:
                    modelo_raw = str(coincidencia.iloc[0]['Modelo_Excel']).upper().strip()
                    for mod in modelos_disponibles:
                        if mod != txt_local["todos"] and (mod.upper() in modelo_raw or modelo_raw in mod.upper()):
                            modelo_detectado_por_vin = mod
                            break
                    
                    if modelo_detectado_por_vin:
                        st.info(f"🚘 **VIN Detectado:** {vin_busqueda} ➔ **Modelo:** {modelo_detectado_por_vin}")
                    else:
                        st.warning(f"⚠️ Bastidor localizado ({modelo_raw}), pero no coincide con ningún modelo del desplegable.")
                else:
                    st.error("❌ Bastidor VIN no encontrado en la base de datos VINes.xlsb.")
            else:
                st.error("❌ BBDD de VINes vacía o no se pudo cargar desde GitHub.")

        with col1:
            default_index = 0
            if modelo_detectado_por_vin and modelo_detectado_por_vin in modelos_disponibles:
                default_index = modelos_disponibles.index(modelo_detectado_por_vin)

            modelo_seleccionado = st.selectbox(
                txt_local["f_modelo"], 
                modelos_disponibles, 
                index=default_index,
                key="sb_modelo_taller_select"
            )

        with col2:
            buscar_pieza = st.text_input(txt_local["f_pieza"], "").strip()

        with col3:
            buscar_operacion = st.text_input(txt_local["f_operacion"], "").strip()

        col_m, col_e = st.columns([2, 2])

        with col_m:
            if "Mercado / Organización" in data.columns:
                mercados_disponibles = [txt_local["todos"]] + [str(m).strip() for m in data["Mercado / Organización"].unique() if str(m).strip() != ""]
                indice_defecto = next((idx for idx, m in enumerate(mercados_disponibles) if "spain" in m.lower() or "oj spain" in m.lower()), 0)
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
                df_filtrado["Nombre de la Pieza"].astype(str).str.contains(buscar_pieza, case=False, na=False) |
                df_filtrado["Código de Referencia"].astype(str).str.contains(buscar_pieza, case=False, na=False)
            ]

        if buscar_operacion and "Operación Técnica" in df_filtrado.columns:
            df_filtrado = df_filtrado[df_filtrado["Operación Técnica"].astype(str).str.contains(buscar_operacion, case=False, na=False)]

        st.markdown(txt_local["res_taller"].format(len(df_filtrado)))
        if not df_filtrado.empty:
            st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
        else:
            st.warning(txt_local["warn_taller"])

    except Exception as exc:
        st.error(txt_local["err_taller"].format(exc))




# =========================================================================
# PANTALLA 3 - SOLICITAR OPERACIÓN (CON DETECCIÓN POR VIN FUNCIONAL)
# =========================================================================
def render_solicitar_operacion(txt_local):
    st.title(txt_local["solicitar_titulo"])
    st.write(txt_local["solicitar_sub"])
    st.markdown("---")

    MAPEO_MODELOS = {
        "OMODA 5 (Gasolina)": "T19C", "OMODA 5 HEV (Híbrido)": "T19C HEV", "OMODA 5 EV (Eléctrico)": "T19C EV",
        "OMODA 7 PHEV": "T1GC PHEV", "OMODA 9 PHEV": "T22 PHEV", "JAECOO 5 (Gasolina)": "T13J",
        "JAECOO 5 HEV": "T13J HEV", "JAECOO 5 BEV": "T13J BEV", "JAECOO 7 (Gasolina)": "T1EJ",
        "JAECOO 7 HEV": "T1EJ HEV", "JAECOO 7 PHEV": "T1EJ PHEV", "JAECOO 8 PHEV": "T26 PHEV",
        "LEPAS L8 PHEV": "T1G PHEV",
    }

    df_vines_db = load_data_vines()

    # Campo VIN fuera del form para actualizar los desplegables en vivo
    vin = st.text_input(
        txt_local["form_vin"], 
        max_chars=17, 
        placeholder=txt_local["form_vin_holder"],
        key="vin_solicitar_input",
        help="Al escribir los 17 caracteres, se detectará la marca y modelo automáticamente."
    ).strip().upper()

    # Detección por VIN
    if len(vin) == 17 and not df_vines_db.empty:
        coincidencia = df_vines_db[df_vines_db['VIN'] == vin]
        if not coincidencia.empty:
            modelo_detectado_raw = str(coincidencia.iloc[0]['Modelo_Excel']).upper().strip()
            for mod_comercial in MAPEO_MODELOS.keys():
                if mod_comercial.upper() in modelo_detectado_raw or modelo_detectado_raw in mod_comercial.upper():
                    nueva_marca = "OMODA" if mod_comercial.startswith("OMODA") else ("JAECOO" if mod_comercial.startswith("JAECOO") else "LEPAS")
                    
                    if st.session_state.solicitar_modelo != mod_comercial:
                        st.session_state.solicitar_marca = nueva_marca
                        st.session_state.solicitar_modelo = mod_comercial
                        st.toast(f"✅ Bastidor detectado: {mod_comercial}", icon="🚘")
                        st.rerun()
                    break

    st.subheader(txt_local["form_sub"])

    col1, col2 = st.columns(2)
    with col1:
        marcas = ["OMODA", "JAECOO", "LEPAS"]
        idx_m = marcas.index(st.session_state.solicitar_marca) if st.session_state.solicitar_marca in marcas else 0
        marca = st.selectbox(txt_local["form_marca"], marcas, index=idx_m, key="sb_marca_solicitar")
        st.session_state.solicitar_marca = marca

        modelos_filtrados = [mod for mod in MAPEO_MODELOS if mod.upper().startswith(marca.upper())]
        idx_mod = modelos_filtrados.index(st.session_state.solicitar_modelo) if st.session_state.solicitar_modelo in modelos_filtrados else 0
        modelo_comercial = st.selectbox(txt_local["form_modelo"], modelos_filtrados, index=idx_mod, key="sb_modelo_solicitar")
        st.session_state.solicitar_modelo = modelo_comercial

    with col2:
        codigo_producto_auto = MAPEO_MODELOS[modelo_comercial]
        st.text_input(txt_local["form_hq_code"], value=codigo_producto_auto, disabled=True)

    with st.form("hq_operation_form", clear_on_submit=True):
        numero_garantia = st.text_input("Nº de Garantía:", placeholder="Ej: CO202607290001", help="Formato: COYYYYMMDDXXXX").strip().upper()
        referencia = st.text_input(txt_local["form_ref"], placeholder=txt_local["form_ref_holder"]).strip().upper()
        operacion_solicitada = st.text_area(txt_local["form_op"], placeholder=txt_local["form_op_holder"]).strip()

        boton_enviar = st.form_submit_button(txt_local["form_btn"])

        if boton_enviar:
            patron_garantia = r"^CO\d{8}[A-Z0-9]{4}$"

            if not numero_garantia or not vin or not operacion_solicitada:
                st.error(txt_local["err_campos"])
            elif not re.match(patron_garantia, numero_garantia):
                st.error("❌ **Error en el Número de Garantía:** Debe cumplir el patrón **COYYYYMMDDXXXX**.")
            elif len(vin) != 17:
                st.error("❌ **Error en el VIN:** El número de bastidor debe tener exactamente 17 caracteres.")
            else:
                ahora = datetime.datetime.now()
                columnas_orden = [
                    "SN", "Submitted on", "Respondents", "Fecha del día",
                    "Marca del vehículo", "INTRODUCIR MODELO", "INTRODUCIR VIN",
                    "Mercado", "CÓDIGO DE PRODUCTO", "REFERENCIA DE PIEZA",
                    "OPERACIÓN QUE SE SOLICITA AÑADIR", "DEALER"
                ]

                nueva_solicitud = {
                    "SN": len(st.session_state.lista_solicitudes) + 1,
                    "Submitted on": ahora.strftime("%Y-%m-%d %H:%M:%S"),
                    "Respondents": f"Garantía: {numero_garantia}",
                    "Fecha del día": ahora.strftime("%Y-%m-%d"),
                    "Marca del vehículo": marca, "INTRODUCIR MODELO": modelo_comercial,
                    "INTRODUCIR VIN": vin, "Mercado": "Spain OJ",
                    "CÓDIGO DE PRODUCTO": codigo_producto_auto,
                    "REFERENCIA DE PIEZA": referencia if referencia else "NaN",
                    "OPERACIÓN QUE SE SOLICITA AÑADIR": operacion_solicitada,
                    "DEALER": numero_garantia,
                }

                subida_exitosa = False
                try:
                    from streamlit_gsheets import GSheetsConnection
                    conn = st.connection("gsheets", type=GSheetsConnection)

                    if "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                        spreadsheet_url = st.secrets["connections"]["gsheets"]["spreadsheet"]
                    elif "gsheets" in st.secrets and "spreadsheet" in st.secrets["gsheets"]:
                        spreadsheet_url = st.secrets["gsheets"]["spreadsheet"]
                    else:
                        spreadsheet_url = st.secrets.get("spreadsheet", "")

                    df_cloud = conn.read(spreadsheet=spreadsheet_url) if spreadsheet_url else pd.DataFrame(columns=columnas_orden)
                    if df_cloud.empty or len(df_cloud.columns) < 2:
                        df_cloud = pd.DataFrame(columns=columnas_orden)
                    else:
                        df_cloud = df_cloud.dropna(how="all").loc[:, ~df_cloud.columns.str.contains("^Unnamed")]

                    nueva_solicitud["SN"] = len(df_cloud) + 1
                    df_nuevo = pd.DataFrame([nueva_solicitud]).reindex(columns=columnas_orden)
                    df_cloud = df_cloud.reindex(columns=columnas_orden)
                    df_actualizado = pd.concat([df_cloud, df_nuevo], ignore_index=True)

                    if spreadsheet_url:
                        conn.update(spreadsheet=spreadsheet_url, data=df_actualizado)
                        subida_exitosa = True
                    else:
                        raise ValueError("No se encontró la URL del archivo de Sheets en st.secrets.")

                except Exception as exc:
                    st.error(f"❌ Error de conexión con Google Sheets: {exc}")
                    st.info("💡 Por seguridad, hemos guardado esta línea en la caché local.")

                st.session_state.lista_solicitudes.append(nueva_solicitud)

                if subida_exitosa:
                    st.success("✅ **Operación registrada con éxito.** La solicitud ha sido transmitida a Central.")
                    time.sleep(1.5)
                    st.rerun()

    if st.session_state.lista_solicitudes:
        st.markdown("---")
        st.subheader("📌 Solicitudes registradas en esta sesión")
        st.dataframe(pd.DataFrame(st.session_state.lista_solicitudes), use_container_width=True, hide_index=True)



# =========================================================================
# MAIN
# =========================================================================
txt, opcion_menu = render_sidebar_and_get_option()

if check_password(txt):
    if opcion_menu == txt["menu_taller"]:
        render_tiempos_taller(txt)
    elif opcion_menu == txt["menu_solicitar"]:
        render_solicitar_operacion(txt)

