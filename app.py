import streamlit as st
import pandas as pd
import datetime
from streamlit_gsheets import GSheetsConnection
import google.generativeai as genai
import unicodedata

def normalizar_texto(texto):
    texto = str(texto)
    return ''.join(c for c in unicodedata.normalize('NFD', texto) if unicodedata.category(c) != 'Mn').lower()

st.set_page_config(page_title="Buscador Técnico OMODA & JAECOO", layout="wide")

# URL Directa (Raw) al archivo Excel en tu GitHub
URL_GITHUB_EXCEL = "https://github.com/MrFeudo/Catalogo-Operaciones/raw/0d67464a70c2267e0f58aabe31f4530976f1aae8/DMS_Active_Spare_Parts.xlsx"

# =========================================================================
# DICCIONARIO DE TRADUCCIÓN (Internacionalización - i18n para TFM)
# =========================================================================
IDIOMAS = {
    "Español": {
        "menu_titulo": "### 🗺️ Menú de Navegación",
        "menu_radio": "Selecciona una herramienta:",
        "menu_taller": "📋 Tiempos de Taller",
        "menu_precios": "💰 Precios de Recambios",
        "menu_solicitar": "📝 Solicitar Operación",
        "menu_ia": "🤖 Asistente IA",
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
        "precios_titulo": "💰 Maestro de Tarifas y Precios de Recambios",
        "precios_sub": "Consulta oficializada de precios y tarifas de distribución vigentes.",
        "f_buscar_recambio": "🔍 Buscar por Código de recambio o Descripción de pieza:",
        "f_mercado_precios": "Filtrar por Mercado / Organización:",
        "f_tarifa": "Filtrar por Tipo de Tarifa:",
        "res_precios": "### 📦 {} referencias de recambios localizadas",
        "warn_precios": "⚠️ No se encontraron recambios con los criterios seleccionados.",
        "err_precios": "Error al procesar el maestro de precios: {}",
        "todos": "Todos",
        "todas": "Todas",
        "filtro_modelo": "Filtrar por Modelo:",
        "solicitar_titulo": "📝 Solicitud de Operaciones Adicionales de Mano de Obra",
        "solicitar_sub": "Utilice este formulario para solicitar el alta de nuevas operaciones o precios en el maestro de HQ.",
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
        "err_vin_corto": "❌ El VIN introducido es demasiado corto. Revíselo.",
        "success_sheet": "✅ ¡Solicitud registrada con éxito! Los datos se han volcado a la plantilla de Central.",
        "warn_contingencia": "⚠️ Formulario correcto, guardado en modo de contingencia local."
    },
    "English": {
        "menu_titulo": "### 🗺️ Navigation Menu",
        "menu_radio": "Select a tool:",
        "menu_taller": "📋 Workshop Times",
        "menu_precios": "💰 Spare Parts Prices",
        "menu_solicitar": "📝 Request Operation",
        "menu_ia": "🤖 AI Assistant",
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
        "precios_titulo": "💰 Master Rate & Spare Parts Prices",
        "precios_sub": "Official consultation of current prices and distribution rates.",
        "f_buscar_recambio": "🔍 Search by Part Code or Description:",
        "f_mercado_precios": "Filter by Market / Organization:",
        "f_tarifa": "Filter by Rate Type:",
        "res_precios": "### 📦 {} spare parts references located",
        "warn_precios": "⚠️ No spare parts found matching the selected criteria.",
        "err_precios": "Error processing master price list: {}",
        "todos": "All",
        "todas": "All",
        "filtro_modelo": "Filter by Model:",
        "solicitar_titulo": "📝 Request for Additional Labor Operations",
        "solicitar_sub": "Use this form to request new operations or prices to be added to HQ master list.",
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
        "form_op_holder": "Describe in detail the technical operation or missing price required by the workshop...",
        "form_btn": "Send Request to HQ",
        "err_campos": "❌ Please fill in all required fields (*).",
        "err_vin_corto": "❌ The entered VIN is too short. Please check it.",
        "success_sheet": "✅ Request successfully registered! Data transferred to the HQ template.",
        "warn_contingencia": "⚠️ Form valid, saved in local contingency mode."
    },
    "Chinese (中文)": {
        "menu_titulo": "### 🗺️ 导航菜单",
        "menu_radio": "选择工具:",
        "menu_taller": "📋 车间工时",
        "menu_precios": "💰 零配件价格",
        "menu_solicitar": "📝 请求操作",
        "menu_ia": "🤖 智能助手",
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
        "precios_titulo": "💰 零售价与零配件价格总表",
        "precios_sub": "官方查询现行价格及分销费率。",
        "f_buscar_recambio": "🔍 按零件代码或描述搜索:",
        "f_mercado_precios": "按市场 / 组织筛选:",
        "f_tarifa": "按费率类型筛选:",
        "res_precios": "### 📦 已定位 {} 个零配件参考",
        "warn_precios": "⚠️ 未找到符合选择条件的零配件。",
        "err_precios": "处理价格总表时出错: {}",
        "todos": "全部",
        "todas": "全部",
        "filtro_modelo": "按车型过滤:",
        "solicitar_titulo": "📝 申请新增工时操作",
        "solicitar_sub": "使用此表单申请在总部(HQ)主数据中添加新工时操作 or 价格。",
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
        "form_op_holder": "请详细描述车间所需的工时操作 or 缺失的价格...",
        "form_btn": "发送申请至总部",
        "err_campos": "❌ 请填写所有必填项 (*)。",
        "err_vin_corto": "❌ 输入的 VIN 太短，请检查。",
        "success_sheet": "✅ 申请登记成功！数据已同步至总部模板。",
        "warn_contingencia": "⚠️ 表单正确，已保存至本地应急模式。"
    }
}

# ==========================================
# 1. BARRA LATERAL: LOGO + SELECCIÓN IDIOMA + MENÚ
# ==========================================
try:
    st.sidebar.image("logo_empresa.png", use_container_width=True)
except Exception:
    st.sidebar.write("🏢 **OMODA & JAECOO**")

st.sidebar.markdown("---")

if "idioma" not in st.session_state:
    st.session_state.idioma = "Español"

idioma_seleccionado = st.sidebar.selectbox(
    "🌐 Language / Idioma / 语言:",
    ["Español", "English", "Chinese (中文)"]
)
st.session_state.idioma = idioma_seleccionado
txt = IDIOMAS[st.session_state.idioma]

st.sidebar.markdown("---")
st.sidebar.markdown(txt["menu_titulo"])

opcion_menu = st.sidebar.radio(
    txt["menu_radio"],
    [txt["menu_taller"], txt["menu_precios"], txt["menu_solicitar"], txt["menu_ia"]]
)

# ==========================================
# 2. SEGURIDAD (Acceso único global)
# ==========================================
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        st.title(txt["pass_titulo"])
        password = st.text_input(txt["pass_input"], type="password")
        if st.button(txt["pass_boton"]):
            if password == "DealersOJ2026":
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error(txt["pass_error"])
        return False
    return True

if check_password():
    
    # =========================================================================
    # PANTALLA 1: TIEMPOS DE TALLER
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
                    
                    market_label = txt["f_mercado_taller"]
                    mercado_seleccionado = st.selectbox(market_label, mercados_disponibles, index=indice_defecto)
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
                st.dataframe(df_filtrado, use_container_width=True, hide_index=True)
            else:
                st.warning(txt["warn_taller"])
                
        except Exception as e:
            st.error(txt["err_taller"].format(e))

    # =========================================================================
    # PANTALLA 2: PRECIOS DE RECAMBIOS
    # =========================================================================
    elif opcion_menu == txt["menu_precios"]:
        
        @st.cache_data
        def load_prices_nueva_version():
            df = pd.read_excel(URL_GITHUB_EXCEL, sheet_name="Parts price")
            df.columns = df.columns.astype(str).str.strip()
            
            df = df.rename(columns={
                'Model': 'Modelo',
                'new_partscode': 'Código de Recambio',
                'new_product_idname': 'Descripción de la Pieza',
                'wholesale price (domestic)': 'Precio Mayorista Doméstico',
                'transactioncurrencyidname': 'Moneda',
                'new_pricetypename': 'Tipo de Tarifa',
                'new_businessunit_idname': 'Mercado / Organización',
                'statecodename': 'Estado'
            })
            
            columnas_finales_precios = [
                'Modelo', 'Código de Recambio', 'Descripción de la Pieza', 
                'Precio Mayorista Doméstico', 'Moneda', 'Tipo de Tarifa', 
                'Mercado / Organización', 'Estado'
            ]
            
            df = df.fillna("")
            df = df.replace("nan", "")
            
            columnas_visibles = [col for col in columnas_finales_precios if col in df.columns]
            return df[columnas_visibles].reset_index(drop=True)

        try:
            prices_data = load_prices_nueva_version()
            
            st.title(txt["precios_titulo"])
            st.write(txt["precios_sub"])
            st.markdown("---")
            
            col_busc, col_org_p, col_tar, col_mod = st.columns([2, 1, 1, 1])
            
            with col_busc:
                buscar_recambio = st.text_input(txt["f_buscar_recambio"], "").strip()
                
            with col_org_p:
                mercados_disponibles = [txt["todos"]] + [str(m).strip() for m in prices_data['Mercado / Organización'].unique() if str(m).strip() != ""]
                
                indice_defecto = 0
                for idx, m in enumerate(mercados_disponibles):
                    if "spain" in m.lower() or "oj spain" in m.lower():
                        indice_defecto = idx
                        break
                        
                market_name = txt["f_mercado_precios"]
                mercado_seleccionado = st.selectbox(market_name, mercados_disponibles, index=indice_defecto)
                
            with col_tar:
                tarifas_disponibles = [txt["todas"]] + [str(t).strip() for t in prices_data['Tipo de Tarifa'].unique() if str(t).strip() != ""]
                tarifa_seleccionada = st.selectbox(txt["f_tarifa"], tarifas_disponibles)

            with col_mod:
                modelos_disponibles = [txt["todos"]] + [str(mo).strip() for mo in prices_data['Modelo'].unique() if str(mo).strip() != ""]
                modelo_seleccionado = st.selectbox(txt["filtro_modelo"], modelos_disponibles)

            df_final_precios = prices_data.copy()

            if modelo_seleccionado != txt["todos"]:
                df_final_precios = df_final_precios[df_final_precios['Modelo'].astype(str).str.strip() == modelo_seleccionado]
                
            if mercado_seleccionado != txt["todos"]:
                df_final_precios = df_final_precios[df_final_precios['Mercado / Organización'].astype(str).str.strip() == mercado_seleccionado]
                
            if tarifa_seleccionada != txt["todas"]:
                df_final_precios = df_final_precios[df_final_precios['Tipo de Tarifa'].astype(str).str.strip() == tarifa_seleccionada]
                
            if buscar_recambio:
                df_final_precios = df_final_precios[
                    df_final_precios['Código de Recambio'].astype(str).str.contains(buscar_recambio, case=False) |
                    df_final_precios['Descripción de la Pieza'].astype(str).str.contains(buscar_recambio, case=False)
                ]

            st.markdown(txt["res_precios"].format(len(df_final_precios)))
            if not df_final_precios.empty:
                st.dataframe(df_final_precios, use_container_width=True, hide_index=True)
            else:
                st.warning(txt["warn_precios"])
                
        except Exception as e:
            st.error(txt["err_precios"].format(e))

    # =========================================================================
    # PANTALLA 3: SOLICITUD DE OPERACIONES ADICIONALES (Para HQ)
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
            "OMODA 5 (Gasolina)": "T19C",
            "OMODA 5 HEV (Híbrido)": "T19C HEV",
            "OMODA 5 EV (Eléctrico)": "T19C EV",
            "OMODA 7 PHEV": "T1GC PHEV",
            "OMODA 9 PHEV": "T22 PHEV",
            "JAECOO 5 (Gasolina)": "T13J",
            "JAECOO 5 HEV": "T13J HEV",
            "JAECOO 5 BEV": "T13J BEV",
            "JAECOO 7 (Gasolina)": "T1EJ",
            "JAECOO 7 HEV": "T1EJ HEV",
            "JAECOO 7 PHEV": "T1EJ PHEV",
            "JAECOO 8 PHEV": "T26",
            "LEPAS L8 PHEV": "T1G PHEV"
        }
        
        st.subheader(txt["form_sub"])
        
        col_dinamica1, col_dinamica2 = st.columns(2)
        with col_dinamica1:
            marca = st.selectbox(txt["form_marca"], ["OMODA", "JAECOO", "LEPAS"])
            modelos_filtrados = [mod for mod in MAPEO_MODELOS.keys() if mod.upper().startswith(marca.upper())]
            modelo_comercial = st.selectbox(txt["form_modelo"], modelos_filtrados)
            
        with col_dinamica2:
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
                elif len(vin) < 11:
                    st.error(txt["err_vin_corto"])
                else:
                    ahora = datetime.datetime.now()
                    
                    nueva_fila = pd.DataFrame([{
                        "SN": "",
                        "Submitted on": ahora.strftime("%Y-%m-%d %H:%M:%S"),
                        "Respondents": f"Dealer App ({dealer})",
                        "Fecha del día": ahora.strftime("%Y-%m-%d"),
                        "Marca del vehículo": marca,
                        "INTRODUCIR MODELO": modelo_comercial,
                        "INTRODUCIR VIN": vin,
                        "Mercado": "Spain OJ",
                        "CÓDIGO DE PRODUCTO": codigo_producto_auto,
                        "REFERENCIA DE PIEZA": referencia if referencia else "NaN",
                        "OPERACIÓN QUE SE SOLICITA AÑADIR": operacion_solicitada,
                        "DEALER": dealer
                    }])
                    
                    ruta_guardado = "C:/Users/PabloMorenoMartín/OneDrive - O&J Automotive Netherlands B.V/Documentos/Mis cosas/Recambios/Catalogo operaciones/solicitudes_operaciones_MO_HQ.xlsx"
                    
                    try:
                        df_local_existente = pd.read_excel(ruta_guardado)
                        df_local_final = pd.concat([df_local_existente, nueva_fila], ignore_index=True)
                    except Exception:
                        df_local_final = nueva_fila
                    
                    try:
                        df_local_final.to_excel(ruta_guardado, index=False)
                        st.success("¡Solicitud guardada con éxito localmente en tu carpeta de OneDrive!")
                        st.balloons()
                        # FIJADO AQUÍ: El st.dataframe queda metido dentro del bloque de éxito del envío
                        st.dataframe(nueva_fila, hide_index=True)
                    except Exception as e_archivo:
                        st.error(f"Error al escribir en el archivo Excel. Asegúrate de que no esté abierto: {e_archivo}")

    

# =========================================================================
    # PANTALLA 4: ASISTENTE IA AVANZADO (Spain OJ + Recambios + Tiempos)
    # =========================================================================
    elif opcion_menu == txt["menu_ia"]:
        import unicodedata
        
        st.title("🤖 Asistente Virtual de Posventa (Gemini)")
        st.write("Consulta tarifas, recambios y tiempos de taller en el DMS.")
        st.markdown("---")
        
        # 1. CONFIGURACIÓN API
        try:
            genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
            genai.api_version = 'v1' 
        except Exception:
            st.error("⚠️ Error: Configura 'GEMINI_API_KEY' en los Secrets de Streamlit.")
            st.stop()
            
        def normalizar(s):
            return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn').lower()

        pregunta = st.text_input("¿Qué necesitas consultar? (Ej: Alternador OMODA 5):").strip()
        
        if pregunta:
            with st.spinner("🤖 Consultando DMS (Análisis Técnico Senior)..."):
              try:
                    # 1. CARGA DE DATOS
                    df_precios = pd.read_csv("DMS_Active_Spare_Parts (1).xlsx - Parts price.csv")
                    df_tiempos = pd.read_csv("DMS_Active_Spare_Parts (1).xlsx - new_srv_workhours.csv")
                    
                    # 2. FILTRADO (Solo Spain OJ)
                    df_p = df_precios[df_precios.astype(str).apply(lambda x: x.str.contains("Spain OJ", na=False)).any(axis=1)]
                    df_t = df_tiempos[df_tiempos.astype(str).apply(lambda x: x.str.contains("Spain OJ", na=False)).any(axis=1)]
                    
                    # 3. CRUCE INTELIGENTE (La clave para que encuentre la relación)
                    # Unimos por el código de pieza común (ajusta los nombres de columna si es necesario)
                    df_unido = pd.merge(df_p, df_t, left_on='new_partscode', right_on='new_code', how='inner', suffixes=('_p', '_t'))
                    
                    # 4. BUSQUEDA POR NORMALIZACIÓN (Tildes)
                    def normalizar(s):
                        return ''.join(c for c in unicodedata.normalize('NFD', str(s)) if unicodedata.category(c) != 'Mn').lower()
                    
                    # Filtramos el DataFrame unido según la consulta
                    query_norm = normalizar(pregunta)
                    mask = df_unido.astype(str).apply(lambda x: x.str.contains(query_norm, na=False)).any(axis=1)
                    contexto = df_unido[mask].head(10).to_string()

                    # 5. LLAMADA A LA IA CON EL CONTEXTO YA CRUZADO
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    
                    prompt = f"""
                    Eres el Ingeniero Jefe Técnico de OMODA España. 
                    CONTEXTO TÉCNICO (Datos ya cruzados):
                    {contexto}
                    
                    CONSULTA: "{pregunta}"
                    
                    INSTRUCCIONES:
                    1. Identifica el código de pieza en la columna 'new_partscode'.
                    2. Extrae el precio de 'new_price'.
                    3. Si hay UTs en 'new_standardhour', DIVIDE ENTRE 100 para dar Horas.
                    4. Si no hay datos exactos, sé honesto y no inventes.
                    """
                    
                    respuesta = model.generate_content(prompt)
                    st.markdown("### 💬 Respuesta Técnica:")
                    st.write(respuesta.text)
                    
                except Exception as e:
                    st.error(f"❌ Error crítico: {e}")
