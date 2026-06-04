import streamlit as st
import pandas as pd
import datetime
import io
import unicodedata

# Set page configuration
st.set_page_config(page_title="Buscador Técnico OMODA & JAECOO", layout="wide")

# =========================================================================
# 1. INICIALIZACIÓN ABSOLUTA DEL SESSION STATE
# =========================================================================
if "lista_solicitudes" not in st.session_state:
    st.session_state.lista_solicitudes = []

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if "idioma" not in st.session_state:
    st.session_state.idioma = "Español"

# URL Directa (Raw) al archivo Excel en GitHub
URL_GITHUB_EXCEL = "https://github.com/MrFeudo/Catalogo-Operaciones/raw/0d67464a70c2267e0f58aabe31f4530976f1aae8/DMS_Active_Spare_Parts.xlsx"

# =========================================================================
# 2. DICCIONARIO DE TRADUCCIÓN (Internacionalización - i18n para TFM)
# =========================================================================
IDIOMAS = {
    "Español": {
        "menu_titulo": "### 🗺️ Menú de Navegación",
        "menu_radio": "Selecciona una herramienta:",
        "menu_taller": "📋 Tiempos de Taller",
        "menu_solicitar": "📝 Solicitar Operación",
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
        "todas": "Todas",
        "solicitar_titulo": "📝 Solicitud de Operaciones Adicionales de Mano de Obra",
        "solicitar_sub": "Utilice este formulario para solicitar el alta de nuevas operaciones en el maestro de HQ.",
        "form_sub": "Datos de la Solicitud (Campos obligatorios *)",
        "form_marca": "Marca del vehículo *",
        "form_modelo": "INTRODUCIR MODELO *",
        "form_vin": "INTRODUCIR VIN (Bastidor) *",
        "form_vin_holder": "17 caracteres",
        "form_dealer": "DEALER (Concesionario) *",
        "form_hq_code": "CÓDIGO DE PRODUCTO (Asignado por HQ)",
        "form_ref": "REFERENCIA DE PIEZA (Opcional)",
        "form_ref_holder": "Ej. 7365747465AA",
        "form_op": "OPERACIÓN QUE SE SOLICITA AÑADIR *",
        "form_op_holder": "Describa detalladamente la operación técnica o falta de precio que requiere el taller...",
        "form_btn": "Enviar Solicitud a Central",
        "err_campos": "❌ Por favor, rellene todos los campos obligatorios (*).",
        "err_vin_corto": "❌ El VIN introducido es demasiado corto. Revíselo."
    },
    "English": {
        "menu_titulo": "### 🗺️ Navigation Menu",
        "menu_radio": "Select a tool:",
        "menu_taller": "📋 Workshop Times",
        "menu_solicitar": "📝 Request Operation",
        "pass_titulo": "🔐 Dealer Network Access",
        "pass_input": "Enter access password:",
        "pass_boton": "Login",
        "pass_error": "❌ Incorrect password",
        "taller_titulo": "🚗 Labor Operations Catalog",
        "taller_sub": "Consult parts, models, and assigned times directly from the DMS.",
        "f_modelo": "1. Filter by Model:",
        "f_pieza": "2. Search by Part Name or Code:",
        "f_operacion": "3. Search by operation type (e.g., Remove, Paint...):",
        "f_mercado_taller": "Filter by Market / Organization (Workshop):",
        "f_estado_taller": "Filter by Operation Status (Workshop):",
        "res_taller": "### 📋 Results found: {} operations",
        "warn_taller": "⚠️ No operations found matching the selected criteria.",
        "err_taller": "Error processing workshop times database: {}",
        "todos": "All",
        "todas": "All",
        "solicitar_titulo": "📝 Request for Additional Labor Operations",
        "solicitar_sub": "Use this form to request new operations to be added to HQ master list.",
        "form_sub": "Request Details (* Required fields)",
        "form_marca": "Vehicle Brand *",
        "form_modelo": "ENTER MODEL *",
        "form_vin": "ENTER VIN (Chassis) *",
        "form_vin_holder": "17 characters",
        "form_dealer": "DEALER *",
        "form_hq_code": "PRODUCT CODE (Assigned by HQ)",
        "form_ref": "PART REFERENCE (Optional)",
        "form_ref_holder": "e.g., 7365747465AA",
        "form_op": "OPERATION REQUESTED TO BE ADDED *",
        "form_op_holder": "Describe in detail the technical operation required by the workshop...",
        "form_btn": "Send Request to HQ",
        "err_campos": "❌ Please fill in all required fields (*).",
        "err_vin_corto": "❌ The entered VIN is too short. Please check it."
    },
    "Chinese (中文)": {
        "menu_titulo": "### 🗺️ 导航菜单",
        "menu_radio": "选择工具:",
        "menu_taller": "📋 车间工时",
        "menu_precios": "💰 零配件价格",
        "menu_solicitar": "📝 请求操作",
        "pass_titulo": "🔐 经销商网络访问",
        "pass_input": "输入访问密码:",
        "pass_boton": "登录",
        "pass_error": "❌ 密码错误",
        "taller_titulo": "🚗 工时操作目录",
        "taller_sub": "直接从 DMS 查询零件、车型和分配的时间。",
        "f_modelo": "1. 按车型筛选:",
        "f_pieza": "2. 按零件名称或代码搜索:",
        "f_operacion": "3. 按操作类型搜索 (例如: Remove, Paint...):",
        "f_mercado_taller": "按市场 / 组织筛选 (车间):",
        "f_estado_taller": "按操作状态筛选 (车间):",
        "res_taller": "### 📋 找到的结果: {} 个操作",
        "warn_taller": "⚠️ 未找到符合选择条件的工时操作。",
        "err_taller": "处理车间工时数据库时出错: {}",
        "todos": "全部",
        "todas": "全部",
        "solicitar_titulo": "📝 申请新增工时操作",
        "solicitar_sub": "使用此表单申请在总部(HQ)主数据中添加新工时操作。",
        "form_sub": "申请信息 (* 为必填项)",
        "form_marca": "车辆品牌 *",
        "form_modelo": "输入车型 *",
        "form_vin": "输入 VIN (车架号) *",
        "form_vin_holder": "17位字符",
        "form_dealer": "经销商 *",
        "form_hq_code": "产品代码 (由总部分配)",
        "form_ref": "零件编号 (选填)",
        "form_ref_holder": "例如: 7365747465AA",
        "form_op": "申请添加的操作内容 *",
        "form_op_holder": "请详细描述车间所需的工时操作...",
        "form_btn": "发送申请至总部",
        "err_campos": "❌ 请填写所有必填项 (*)。",
        "err_vin_corto": "❌ 输入的 VIN 太短，请检查。"
    }
}

# ==========================================
# 3. BARRA LATERAL: CONFIGURACIÓN E IDIOMA
# ==========================================
try:
    st.sidebar.image("logo_empresa.png", use_container_width=True)
except Exception:
    st.sidebar.write("🏢 **OMODA & JAECOO**")

st.sidebar.markdown("---")

idioma_seleccionado = st.sidebar.selectbox(
    "🌐 Language / Idioma / 语言:",
    ["Español", "English", "Chinese (中文)"],
    index=["Español", "English", "Chinese (中文)"].index(st.session_state.idioma),
    key="selector_idioma_global"
)
st.session_state.idioma = idioma_seleccionado
txt = IDIOMAS[st.session_state.idioma]

st.sidebar.markdown("---")
st.sidebar.markdown(txt["menu_titulo"])

opcion_menu = st.sidebar.radio(
    txt["menu_radio"],
    [txt["menu_taller"], txt["menu_solicitar"]],  
    key="menu_navegacion_app"
)

def buscador_tradicional_excel(consulta_usuario, df_contexto):
    try:
        # =====================================================================
        # 🎯 1. DICCIONARIO SEMÁNTICO ULTRA-EXPANDIDO (ACCIONES Y ENERGÍAS)
        # =====================================================================
        mapa_raices = {
            # --- 🛠️ ACCIONES Y VERBOS ---
            "cambiar": "remove and reinstall|replace|remove|reinstall",
            "cambio": "remove and reinstall|replace|remove|reinstall",
            "sustituir": "remove and reinstall|replace|remove|reinstall",
            "sustitucion": "remove and reinstall|replace|remove|reinstall",
            "reemplazar": "remove and reinstall|replace|remove|reinstall",
            "desmontar": "remove", "montar": "reinstall",
            "comprobar": "check|inspection|test|diagnostic|measurement",
            "verificar": "check|inspection|test|diagnostic",
            "revisar": "check|inspection|test|diagnostic",
            "actualizar": "refresh|update|software|flash",
            "programar": "refresh|update|software|flash|coding|program",
            "ajustar": "adjust|adjustment|alignment|calibrate|calibration",
            "limpiar": "clean|cleaning|wash",
            "pulir": "polishing|polish", "pulido": "polishing|polish",
            
            # --- 🔌 ELECTRÓNICA Y MÓDULOS ---
            "ecu": "ecu|engine control unit", "mcu": "mcu|motor control unit",
            "vcu": "vcu|vehicle control unit", "tcu": "tcu|transmission control unit",
            "bcm": "bcm|body control module", "bdm": "bcm|bdm",
            "centralita": "control unit|control module|ecu|bcm|mcu",
            "modulo": "control module|module", "modulos": "module",

            # --- ⚡ BATERÍAS Y ELEC. ---
            "bateria": "battery|storage battery|bms", "vateria": "battery", 
            "cable": "wiring|harness|wire", "cableado": "wiring|harness",
            "airbag": "airbag|air bag|srs|abm", "cinturon": "seatbelt|seat belt",
            "hebilla": "buckle", "capo": "hood", "paragolpes": "bumper",
            "faro": "headlamp|headlight", "piloto": "lamp|tail lamp",
            "canister": "canister|evap", "vapores": "canister|evap",
            "filtro": "filter", "junta": "gasket|seal", "tubo": "pipe|hose",

            # --- 📍 UBICACIONES ---
            "delantero": "fr|front", "delantera": "fr|front", "frontal": "fr|front",
            "trasero": "rr|rear", "trasera": "rr|rear", "posterior": "rr|rear",
            "izquierdo": "lh|left", "izquierda": "lh|left",
            "derecho": "rh|right", "derecha": "rh|right",
            
            # --- 🔋 MOTORIZACIONES Y ENERGÍAS (CRÍTICO) ---
            "hev": "hev|hybrid", "hibrido": "hev|hybrid", "hibrida": "hev|hybrid",
            "ev": "ev|bev|electric", "electrico": "ev|bev|electric", "electrica": "ev|bev|electric",
            "phev": "phev|plug-in", "enchufable": "phev|plug-in",
            "gasolina": "ice|gasoline", "termico": "ice", "atm": "ice"
        }

        # Diccionario rápido de conversión de modelos pegados
        abreviaturas_modelos = {
            "j5": "jaecoo 5", "jaecoo5": "jaecoo 5", "j-5": "jaecoo 5",
            "j7": "jaecoo 7", "jaecoo7": "jaecoo 7", "j-7": "jaecoo 7",
            "j8": "jaecoo 8", "jaecoo8": "jaecoo 8",
            "o5": "omoda 5", "omoda5": "omoda 5", "o-5": "omoda 5",
            "o7": "omoda 7", "omoda7": "omoda 7", "o-7": "omoda 7",
            "o9": "omoda 9", "omoda9": "omoda 9", "o-9": "omoda 9"
        }

        # 1. Normalización de la consulta del usuario
        consulta_limpia = consulta_usuario.lower().strip()
        for orig, dest in [("í", "i"), ("ó", "o"), ("á", "a"), ("é", "e"), ("ú", "u"), ("ñ", "n")]:
            consulta_limpia = consulta_limpia.replace(orig, dest)

        # Expandimos los modelos si vienen juntos (ej: j7 -> jaecoo 7)
        for abrev, mod_real in abreviaturas_modelos.items():
            if abrev in consulta_limpia.split() or abrev in consulta_limpia:
                consulta_limpia = consulta_limpia.replace(abrev, mod_real)

        lista_palabras_usuario = consulta_limpia.split()

        # =====================================================================
        # 🧪 2. PREPARACIÓN DE LA BASE DE DATOS (LIMPIEZA DE VACÍOS)
        # =====================================================================
        df_base = df_contexto.copy()
        
        # Filtro obligatorio: Eliminamos celdas corruptas o vacías de raíz
        df_base = df_base[
            (df_base['Nombre de la Pieza'].notna()) & (df_base['Nombre de la Pieza'].astype(str).str.strip() != "") &
            (df_base['Código de Referencia'].notna()) & (df_base['Código de Referencia'].astype(str).str.strip() != "") &
            (df_base['Operación Técnica'].notna()) & (df_base['Operación Técnica'].astype(str).str.strip() != "")
        ]

        # Creamos columnas auxiliares en minúsculas para realizar el escaneo interno
        df_base['mod_low'] = df_base['Modelo'].astype(str).str.lower().str.strip()
        df_base['pieza_low'] = df_base['Nombre de la Pieza'].astype(str).str.lower().str.strip()
        df_base['op_low'] = df_base['Operación Técnica'].astype(str).str.lower().str.strip()

        # Inicializamos el casillero de puntos
        df_base['score'] = 0

        # =====================================================================
        # 📊 3. ALGORITMO DE ASIGNACIÓN DE PUNTOS (SISTEMA INDEXADO)
        # =====================================================================
        
        # A. PUNTOS POR MARCA PRINCIPAL (Filtro base)
        if "omoda" in consulta_limpia:
            df_base['score'] += df_base['mod_low'].str.contains("omoda", na=False).astype(int) * 50
        elif "jaecoo" in consulta_limpia:
            df_base['score'] += df_base['mod_low'].str.contains("jaecoo", na=False).astype(int) * 50

        # B. PUNTOS POR NÚMERO DE MODELO EXACTO (Filtro estricto de silueta)
        modelos_numericos = ["5", "7", "8", "9"]
        num_detectado = None
        for num in modelos_numericos:
            if num in lista_palabras_usuario or any(f"omoda{num}" in w or f"jaecoo{num}" in w for w in lista_palabras_usuario) or num in consulta_limpia:
                num_detectado = num
                break
        
        if num_detectado:
            # Si coincide el número, gana 100 puntos. Si NO coincide, pierde 200 (penalización para borrar otros coches)
            df_base['score'] += df_base['mod_low'].str.contains(num_detectado, na=False).astype(int) * 100
            df_base['score'] -= (~df_base['mod_low'].str.contains(num_detectado, na=False)).astype(int) * 200

        # C. PUNTOS POR MOTORIZACIÓN / ENERGÍA (Gasolina, HEV, PHEV, EV)
        energias_claves = ["hev", "phev", "ev", "ice", "hybrid", "electric", "gasolina", "enchufable", "hibrido"]
        energia_solicitada = any(eng in consulta_limpia for eng in energias_claves)
        
        if energia_solicitada:
            # Buscamos qué raíces inglesas corresponden a lo que pidió el usuario
            raices_energia = []
            for esp, eng in mapa_raices.items():
                if esp in ["hev", "hibrido", "hibrida", "ev", "electrico", "electrica", "phev", "enchufable", "gasolina", "termico"] and esp in consulta_limpia:
                    raices_energia.extend(eng.split('|'))
            
            if raices_energia:
                regex_eng = '|'.join(set(raices_energia))
                # Si el modelo contiene la motorización exacta, gana 80 puntos de ventaja
                df_base['score'] += df_base['mod_low'].str.contains(regex_eng, na=False).astype(int) * 80
                # Si el usuario pide explícitamente una energía, penalizamos levemente los modelos puros de otras para limpiar
                if "ev" in consulta_limpia or "electrico" in consulta_limpia:
                    # Si es eléctrico, restamos puntos a los que digan hibrido, phev o gasolina en el nombre
                    df_base['score'] -= df_base['mod_low'].str.contains("hev|phev|ice|hybrid|gasoline", na=False).astype(int) * 50

        # D. PUNTOS POR PALABRAS CLAVE TÉCNICAS (El núcleo de la avería)
        for palabra in lista_palabras_usuario:
            # Ignoramos conectores comunes
            if palabra in ["quiero", "para", "con", "del", "una", "uno", "el", "la", "los", "las", "este", "un", "de", "omoda", "jaecoo", "5", "7", "8", "9"]:
                continue
            
            # Buscamos si la palabra tiene traducción técnica en nuestro mega diccionario
            terminos_ingleses = []
            if palabra in mapa_raices:
                terminos_ingleses.extend(mapa_raices[palabra].split('|'))
            else:
                if len(palabra) > 2: # Si no está en el mapa pero es larga, la buscamos tal cual (ej: códigos o siglas)
                    terminos_ingleses.append(palabra)
            
            if terminos_ingleses:
                regex_tecnica = '|'.join(set(terminos_ingleses))
                # Multiplicamos el score: coincidir en pieza u operación da la máxima puntuación (120 puntos)
                df_base['score'] += df_base['pieza_low'].str.contains(regex_tecnica, na=False).astype(int) * 120
                df_base['score'] += df_base['op_low'].str.contains(regex_tecnica, na=False).astype(int) * 120

        # =====================================================================
        # 🔴 4. FILTRADO Y EXTRACCIÓN DE RESULTADOS GANADORES
        # =====================================================================
        # Solo nos quedamos con las filas que tengan una puntuación positiva y superior a la penalización base
        # Esto elimina de golpe los "cambios de neumáticos" porque su puntuación se quedará en 0 o negativa
        df_filtrado_final = df_base[df_base['score'] > 30]

        if df_filtrado_final.empty:
            return None

        # Ordenamos los resultados de mayor a menor puntuación (relevancia pura)
        df_ordenado = df_filtrado_final.sort_values(by=['score', 'Modelo'], ascending=[False, True]).head(60)

        # Recuperamos las celdas limpias, inalteradas y en inglés oficial del df_contexto original
        df_output = df_contexto.loc[df_ordenado.index].copy()
        
        return df_output[['Modelo', 'Nombre de la Pieza', 'Código de Referencia', 'Operación Técnica', 'Tiempo Estándar (UT/Horas)']]

    except Exception:
        return None
# ==========================================
# 5. SISTEMA DE SEGURIDAD CONTRASEÑA
# ==========================================
def check_password():
    if not st.session_state.authenticated:
        st.title(txt["pass_titulo"])
        password = st.text_input(txt["pass_input"], type="password", key="pass_input_unico")
        if st.button(txt["pass_boton"], key="pass_btn_unico"):
            if password == "DealersOJ2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error(txt.get("pass_error", "❌ Contraseña incorrecta"))
        return False
    return True

if check_password():
    
    # =========================================================================
    # PANTALLA 1: TIEMPOS DE TALLER (CON MOTOR DE TRADUCCIÓN LOCAL)
    # =========================================================================
    if opcion_menu == txt["menu_taller"]:
        
        @st.cache_data
        def load_data_tiempos_v3():
            df = pd.read_excel(URL_GITHUB_EXCEL, sheet_name="new_srv_workhours")
            df.columns = df.columns.astype(str).str.strip()
            
            df = df.rename(columns={
                'new_productmodel_idname': 'Modelo',
                'new_product_idname': 'Nombre de la Pieza',
                'new_code': 'Código de Referencia',
                'new_name': 'Operación Técnica',
                'new_standardhour': 'Tiempo Estándar (UT/Horas)',
                'new_remark': 'Notas / Exclusiones',
                'Organization': 'Mercado / Organización',
                'statecodename': 'Estado'
            })
            
            columnas_finales = [
                'Modelo', 'Nombre de la Pieza', 'Código de Referencia', 
                'Operación Técnica', 'Tiempo Estándar (UT/Horas)', 'Notas / Exclusiones',
                'Mercado / Organización', 'Estado'
            ]
            
            df = df.fillna("")
            df = df.replace("nan", "")
            
            columnas_presentes = [col for col in columnas_finales if col in df.columns]
            return df[columnas_presentes].reset_index(drop=True)

        try:
            data = load_data_tiempos_v3()
            
            st.title(txt["taller_titulo"])
            st.write(txt["taller_sub"])
            st.markdown("---")

            if "resultado_tradicional_excel" not in st.session_state:
                st.session_state.resultado_tradicional_excel = None

            # =================================================================
            # 🤖 SECCIÓN A: ASISTENTE TRADICIONAL DE BÚSQUEDA BILINGÜE
            # =================================================================
            st.subheader("🤖 Buscar operación")
            st.write("Escribe tu consulta en español (ej: cambiar airbag delantero jaecoo 7). El sistema traducirá los términos y buscará en inglés.")
            
            consulta_rapida = st.text_input(
                "¿Qué operación, pieza o modelo necesitas localizar?",
                placeholder="Ejemplo: cambiar pastillas de freno delanteras del omoda 5 / desmontar paragolpes jaecoo 7...",
                key="campo_consulta_tradicional_excel"
            )

            st.warning("""
            ⚠️ **RECORDATORIO** Antes de tramitar cualquier reclamación, verifique obligatoriamente que **la pieza a reclamar coincide con el pedido exacto realizado a Recambios** para esta reparación. 
            """)

            if st.button("Buscar operación", type="primary"):
                if not consulta_rapida.strip():
                    st.warning("⚠️ Introduce una descripción o término para realizar la búsqueda.")
                else:
                    with st.spinner("🔍 Traduciendo y escaneando el catálogo de operaciones..."):
                        st.session_state.resultado_tradicional_excel = buscador_tradicional_excel(consulta_rapida, data)

            # RENDERIZADO DEL RESULTADO TRADICIONAL EN DATAFRAME INTERACTIVO
            if st.session_state.resultado_tradicional_excel is not None:
                st.markdown("#### 🎯 Operaciones encontradas en el catálogo oficial:")
                
                # Pintamos los resultados de forma limpia e interactiva
                st.dataframe(
                    st.session_state.resultado_tradicional_excel,
                    use_container_width=True,
                    hide_index=True
                )
                
                if st.button("🗑️ Limpiar búsqueda", key="btn_limpiar_tradicional"):
                    st.session_state.resultado_tradicional_excel = None
                    st.rerun()

            st.markdown("---")

            # =================================================================
            # 📊 SECCIÓN B: FILTROS Y TABLA TRADICIONAL (BÚSQUEDA MANUAL)
            # =================================================================
            st.subheader("📊 Catálogo Completo (Filtros Manuales)")
            
            col1, col2, col3 = st.columns([1, 1.5, 1.5])
            with col1:
                modelos_disponibles = [txt["todos"]] + list(data['Modelo'].dropna().unique())
                modelo_seleccionado = st.selectbox(txt["f_modelo"], modelos_disponibles)
            with col2:
                buscar_pieza = st.text_input(txt["f_pieza"], "").strip()
            with col3:
                buscar_operacion = st.text_input(txt["f_operacion"], "").strip()

            col_m, col_e = st.columns([2, 2])
            
            with col_m:
                if 'Mercado / Organización' in data.columns:
                    mercados_disponibles = [txt["todos"]] + [str(m).strip() for m in data['Mercado / Organización'].unique() if str(m).strip() != ""]
                    indice_defecto = 0
                    for idx, m in enumerate(mercados_disponibles):
                        if "spain" in m.lower() or "oj spain" in m.lower():
                            indice_defecto = idx
                            break
                    mercado_seleccionado = st.selectbox(txt["f_mercado_taller"], mercados_disponibles, index=indice_defecto)
                else:
                    mercado_seleccionado = txt["todos"]
                    
            with col_e:
                if 'Estado' in data.columns:
                    estados_disponibles = [txt["todos"]] + [str(e).strip() for e in data['Estado'].unique() if str(e).strip() != ""]
                    indice_est_defecto = estados_disponibles.index("Active") if "Active" in estados_disponibles else 0
                    estado_seleccionado = st.selectbox(txt["f_estado_taller"], estados_disponibles, index=indice_est_defecto)
                else:
                    estado_seleccionado = txt["todos"]

            df_filtrado = data.copy()
            
            if modelo_seleccionado != txt["todos"]:
                df_filtrado = df_filtrado[df_filtrado['Modelo'] == modelo_seleccionado]
                
            if mercado_seleccionado != txt["todos"] and 'Mercado / Organización' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Mercado / Organización'].astype(str).str.strip() == mercado_seleccionado]
                
            if estado_seleccionado != txt["todos"] and 'Estado' in df_filtrado.columns:
                df_filtrado = df_filtrado[df_filtrado['Estado'].astype(str).str.strip() == estado_seleccionado]

            if buscar_pieza:
                df_filtrado = df_filtrado[
                    df_filtrado['Nombre de la Pieza'].astype(str).str.contains(buscar_pieza, case=False, na=False) |
                    df_filtrado['Código de Referencia'].astype(str).str.contains(buscar_pieza, case=False, na=False)
                ]
                
            if buscar_operacion:
                df_filtrado = df_filtrado[df_filtrado['Operación Técnica'].astype(str).str.contains(buscar_operacion, case=False, na=False)]

            st.markdown(txt["res_taller"].format(len(df_filtrado)))
            if not df_filtrado.empty:
                st.dataframe(df_filtrado, width='stretch', hide_index=True)
            else:
                st.warning(txt["warn_taller"])
                
        except Exception as e:
            st.error(txt["err_taller"].format(e))

    # =========================================================================
    # PANTALLA 2: SOLICITUD DE OPERACIONES ADICIONALES (CONEXIÓN GOOGLE SHEETS)
    # =========================================================================
    elif opcion_menu == txt["menu_solicitar"]:
        st.title(txt["solicitar_titulo"])
        st.write(txt["solicitar_sub"])
        st.markdown("---")

        LISTA_DEALERS = sorted([
            "ACAI MOTOR MÁLAGA", "ALIFAVISA BILBAO", "ALIMOTOR ELCHE", "ANFERPA SEGOVIA", 
            "AUTO YALDE LOGROÑO", "AUTOCAM MOTOR VILAFRANCA", "AUTOCYL PALENCIA", "AUTOCYL VALLADOLID", 
            "AUTOTERMINAL", "AUTOVIDAL PALMA DE MALLORCA", "AXIS MOTORS", "BLENDIO LAREDO", 
            "BLENDIO LUGO", "BLENDIO OURENSE", "BLENDIO SANTANDER", "BLENDIO TORRELAVEGA", 
            "BORJAMOTOR ALICANTE", "CERVERA AVILA", "CERVERA SALAMANCA", "CHINARES GUADALAJARA", 
            "DUMOSA BENAVENTE", "ESLAUTO LEON", "GRUP BASOLS IGUALADA", "GRUPO JULIAN BURGOS", 
            "GRUPO NIETO MÁLAGA", "GRUPO NIETO MARBELLA", "HIMASA SEDAVÍ", "JEMOYA SORIA", 
            "LASACAR MIRANDA DE EBRO", "LASACAR VITORIA", "M TECNIK ALCALÁ DE HENARES", 
            "M TECNIK BARCELONA MAQUINIST", "M TECNIK CASTELLÓN", "M TECNIK GERONA", 
            "M TECNIK MATARÓ", "M TECNIK VINAROZ", "MARTIN LIZAGA", "MARTIN LIZAGA TERUEL", 
            "MAS AUTO LEGANÉS", "MAVEN BADAJOZ", "MAVEN CÁCERES", "MOLL MOTOR DENIA", 
            "MOLL MOTOR GANDIA", "MONECAR AUTOMOCION", "MONECAR CUENCA", "MOTOR NACIENTE", 
            "MOVINSUR GRANADA", "MOVINSUR JAÉN", "MOVINSUR MOTRIL", "MY CARS CÓRDOBA", 
            "NOATUM", "NOVACAR BCN SANT BOI", "PALAUSA ZAMORA", "PRUNA CAR GO GRANOLLERS", 
            "PROCHERY ALBACETE", "PROCHERY CARTAGENA", "PROCHERY MURCIA", "RAFAEL AFONSO AGUIMES", 
            "RAFAEL AFONSO LANZAROTE", "RAFAEL AFONSO LAS PALMAS", "RAFAEL AFONSO TENERIFE", 
            "RESNOVA MOTOR CORUÑA", "RESNOVA MOTOR GIJÓN", "RESNOVA MOTOR NARÓN", 
            "RESNOVA MOTOR OVIEDO", "RESNOVA MOTOR SANTIAGO", "RESNOVA MOTOR VIGO", 
            "SEGRE MOTORS LERIDA", "SERTECAUTO PONFERRADA", "SYRSA ALGECIRAS", 
            "SYRSA ALMERIA", "SYRSA EJIDO", "SYRSA HUELVA", "SYRSA SEVILLA", 
            "TALAUTO CAZALEGAS", "TALAUTO TOLEDO", "TALLERES CHINARES", "TECNOTARRACO TARRAGONA", 
            "TERRY MOBILITY JERÉZ", "TRADECAR GAMBOA ALCORCÓN", "TRADECAR GAMBOA MADRID", 
            "TRADECAR GAMBOA MAJADAHON", "TRADECAR GAMBOA RIVAS", "TUMASA HUESCA", 
            "TUMASA MONZÓN", "UNIONE ALCAZAR DE SAN JUAN", "UNIONE CIUDAD REAL", 
            "VALLESCAR SABADELL", "VALLESCAR TERRASSA", "VIAN AUTOMOBILE VILLALBA", 
            "ZEN MOTOR OLABERRIA", "ZEN MOTOR PAMPLONA", "ZEN MOTOR SAN SEBASTIÁN", 
            "ZEN MOTOR ZARAGOZA"
        ])
        
        MAPEO_MODELOS = {
            "OMODA 5 (Gasolina)": "T19C", "OMODA 5 HEV (Híbrido)": "T19C HEV", "OMODA 5 EV (Eléctrico)": "T19C EV",
            "OMODA 7 PHEV": "T1GC PHEV", "OMODA 9 PHEV": "T22 PHEV", "JAECOO 5 (Gasolina)": "T13J",
            "JAECOO 5 HEV": "T13J HEV", "JAECOO 5 BEV": "T13J BEV", "JAECOO 7 (Gasolina)": "T1EJ",
            "JAECOO 7 HEV": "T1EJ HEV", "JAECOO 7 PHEV": "T1EJ PHEV", "JAECOO 8 PHEV": "T26", "LEPAS L8 PHEV": "T1G PHEV"
        }
        
        st.subheader(txt["form_sub"])
        col1, col2 = st.columns(2)
        with col1:
            marca = st.selectbox(txt["form_marca"], ["OMODA", "JAECOO", "LEPAS"])
            modelos_filtrados = [mod for mod in MAPEO_MODELOS.keys() if mod.upper().startswith(marca.upper())]
            modelo_comercial = st.selectbox(txt["form_modelo"], modelos_filtrados)
        with col2:
            dealer = st.selectbox(txt["form_dealer"], LISTA_DEALERS)
            codigo_producto_auto = MAPEO_MODELOS[modelo_comercial]
            st.text_input(txt["form_hq_code"], value=codigo_producto_auto, disabled=True)
        
        with st.form("hq_operation_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            with c1:
                vin = st.text_input(txt["form_vin"], max_chars=17, placeholder=txt["form_vin_holder"]).strip().upper()
            with c2:
                referencia = st.text_input(txt["form_ref"], placeholder=txt["form_ref_holder"]).strip().upper()
            
            operacion_solicitada = st.text_area(txt["form_op"], placeholder=txt["form_op_holder"]).strip()
            boton_enviar = st.form_submit_button(txt["form_btn"])
            
            if boton_enviar:
                if not vin or not operacion_solicitada:
                    st.error(txt["err_campos"])
                elif len(vin) != 17:
                    st.error("❌ Error: El número de bastidor (VIN) debe tener exactamente 17 caracteres.")
                else:
                    ahora = datetime.datetime.now()
                    columnas_orden = [
                        "SN", "Submitted on", "Respondents", "Fecha del día", 
                        "Marca del vehículo", "INTRODUCIR MODELO", "INTRODUCIR VIN", 
                        "Mercado", "CÓDIGO DE PRODUCTO", "REFERENCIA DE PIEZA", 
                        "OPERACIÓN QUE SE SOLICITA AÑADIR", "DEALER"
                    ]
                    
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

                        df_cloud = conn.read(spreadsheet=spreadsheet_url)
                        
                        if df_cloud.empty or len(df_cloud.columns) < 2:
                            df_cloud = pd.DataFrame(columns=columnas_orden)
                        else:
                            df_cloud = df_cloud.dropna(how='all').loc[:, ~df_cloud.columns.str.contains('^Unnamed')]
                        
                        nuevo_sn = len(df_cloud) + 1
                    except Exception:
                        nuevo_sn = len(st.session_state.lista_solicitudes) + 1
                        df_cloud = pd.DataFrame(columns=columnas_orden)
                        spreadsheet_url = ""
                    
                    nueva_solicitud = {
                        "SN": int(nuevo_sn),
                        "Submitted on": str(ahora.strftime("%Y-%m-%d %H:%M:%S")),
                        "Respondents": str(f"Dealer App ({dealer})"),
                        "Fecha del día": str(ahora.strftime("%Y-%m-%d")),
                        "Marca del vehículo": str(marca),
                        "INTRODUCIR MODELO": str(modelo_comercial),
                        "INTRODUCIR VIN": str(vin),
                        "Mercado": "Spain OJ",
                        "CÓDIGO DE PRODUCTO": str(codigo_producto_auto),
                        "REFERENCIA DE PIEZA": str(referencia) if referencia else "NaN",
                        "OPERACIÓN QUE SE SOLICITA AÑADIR": str(operacion_solicitada),
                        "DEALER": str(dealer)
                    }
                    
                    try:
                        df_nuevo = pd.DataFrame([nueva_solicitud])
                        df_nuevo = df_nuevo.reindex(columns=columnas_orden)
                        df_cloud = df_cloud.reindex(columns=columnas_orden)
                        
                        df_actualizado = pd.concat([df_cloud, df_nuevo], ignore_index=True)
                        
                        if spreadsheet_url:
                            conn.update(
                                spreadsheet=spreadsheet_url,
                                data=df_actualizado
                            )
                            st.session_state.lista_solicitudes.append(nueva_solicitud)
                            subida_exitosa = True
                        else:
                            raise ValueError("No se encontró la URL de Sheets.")
                            
                    except Exception as e:
                        st.error(f"❌ Error de conexión con Google Sheets: {e}")
                        st.info("💡 Solicitud guardada temporalmente en caché local.")
                        st.session_state.lista_solicitudes.append(nueva_solicitud)

                    if subida_exitosa:
                        st.success("✅ **Operación registrada con éxito.** Transmitida al departamento de Garantías de Central.")
                        import time
                        time.sleep(1.5)
                        st.rerun()
