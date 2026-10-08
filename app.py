import streamlit as st
import pandas as pd
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
import plotly.express as px
import plotly.graph_objects as go
import unicodedata
import os

st.set_page_config(page_title="Tablero de Pagos ObSBA", layout="wide")
st.title("📊 Tablero de Pagos Diario y Acumulado - ObSBA")

# ==========================================
# 1. DICCIONARIO EXTERNO DE RUBROS
# ==========================================
@st.cache_data
def cargar_diccionario_rubros():
    archivo_txt = "categorias.txt"
    
    if not os.path.exists(archivo_txt):
        st.error(f"⚠️ No se encontró el archivo '{archivo_txt}'. Por favor, subilo a tu repositorio en GitHub.")
        return pd.DataFrame(columns=['Ente', 'Rubro', 'Nombre_Referencia']), {}

    # Leemos el archivo soportando caracteres especiales (tildes, ñ)
    try:
        with open(archivo_txt, "r", encoding="utf-8") as f:
            lineas = [l.strip() for l in f.readlines() if l.strip()]
    except UnicodeDecodeError:
        with open(archivo_txt, "r", encoding="latin1") as f:
            lineas = [l.strip() for l in f.readlines() if l.strip()]

    datos = []
    rubro_actual = "Sin Rubro"
    
    for linea in lineas:
        if linea.startswith("▸"):
            rubro_actual = linea.replace("▸", "").strip()
        # Evitamos arrastrar la fila de TOTAL GENERAL del txt
        elif "-" in linea and not linea.upper().startswith("TOTAL"):
            partes = linea.split("-", 1)
            if len(partes) == 2:
                ente = partes[0].strip()
                nombre = partes[1].strip()
                datos.append({'Ente': str(ente), 'Rubro': rubro_actual, 'Nombre_Referencia': nombre})
                
    df_diccionario = pd.DataFrame(datos)
    df_diccionario['Ente'] = df_diccionario['Ente'].astype(str).str.strip()
    
    # Función para normalizar nombres (quita tildes y pasa a mayúscula)
    def clean_str(s):
        if pd.isna(s): return ""
        s = str(s).upper().strip()
        return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        
    df_diccionario['Nombre_Limpio'] = df_diccionario['Nombre_Referencia'].apply(clean_str)
    
    # Creamos un mapa de rescate para buscar por nombre si falla el código
    dict_nombres = df_diccionario.set_index('Nombre_Limpio')['Rubro'].to_dict()
    
    return df_diccionario, dict_nombres

df_rubros_maestro, dict_nombres = cargar_diccionario_rubros()

# ==========================================
# 2. FUNCIONES DE FORMATO 
# ==========================================
def formatear_moneda(valor):
    try:
        return f"$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return "$ 0,00"

formato_pesos = JsCode(r"""
function(params) {
    var val = Number(params.value);
    if (isNaN(val)) { return "$ 0,00"; }
    var parts = val.toFixed(2).split(".");
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    return "$ " + parts.join(",");
}
""")

def clean_str(s):
    if pd.isna(s): return ""
    s = str(s).upper().strip()
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

# ==========================================
# 3. LÓGICA DE CARGA Y EXTRACCIÓN
# ==========================================
st.sidebar.header("Carga de Datos")
archivo_pagos = st.sidebar.file_uploader("Subí el Reporte de Pagos (.csv o .xlsx)", type=['csv', 'xlsx'])

if archivo_pagos is not None:
    try:
        if archivo_pagos.name.endswith('.csv'):
            df_pagos = pd.read_csv(archivo_pagos, skiprows=6, skipinitialspace=True, encoding='latin1', sep=';')
        else:
            df_pagos = pd.read_excel(archivo_pagos, skiprows=6)
            
        df_pagos.columns = df_pagos.columns.str.strip()
        df_pagos = df_pagos.dropna(how='all')
        
        # ELIMINAR BASURA ABSOLUTA Y TOTALES FINALES
        filtro_basura = df_pagos.astype(str).apply(lambda col: col.str.contains('TOTAL|MEP|Cheque', case=False, na=False)).any(axis=1)
        if filtro_basura.any():
            df_pagos = df_pagos[~filtro_basura].copy()
            st.sidebar.success("✅ Resumen final excluido.")
            
        # EXTRACCIÓN DEL ENTE NUMÉRICO Y NOMBRE DEL CSV
        if 'Benef.OP' in df_pagos.columns:
            df_pagos['Ente_Crudo'] = df_pagos['Benef.OP'].astype(str).str.strip()
            df_pagos['Ente'] = df_pagos['Ente_Crudo'].str.extract(r'(^\d+)', expand=False).fillna('')
            
            col_desc = 'Desc.' if 'Desc.' in df_pagos.columns else 'Descripción' if 'Descripción' in df_pagos.columns else None
            df_pagos['Nombre_Extraido'] = df_pagos['Ente_Crudo'].str.replace(r'^\d+\.?\d*\s*-?\s*', '', regex=True).str.strip()
            
            if col_desc and col_desc in df_pagos.columns:
                df_pagos['Nombre_CSV'] = df_pagos[col_desc].replace('', pd.NA).fillna(df_pagos['Nombre_Extraido'])
            else:
                df_pagos['Nombre_CSV'] = df_pagos['Nombre_Extraido']
                
            df_pagos['Nombre_CSV'] = df_pagos['Nombre_CSV'].fillna('Prestador Desconocido')
        else:
            st.error("No se encontró la columna 'Benef.OP'.")
            st.stop()
            
        # LIMPIEZA FINANCIERA
        columnas_dinero = ['Imp.OP', 'Imp.Neto', 'Imp.Ret.']
        col_anul = 'Imp.Anul.' if 'Imp.Anul.' in df_pagos.columns else ('Perc. por terceros' if 'Perc. por terceros' in df_pagos.columns else None)
        if col_anul:
            columnas_dinero.append(col_anul)
            
        for col in columnas_dinero:
            if col in df_pagos.columns:
                if df_pagos[col].dtype == object:
                    df_pagos[col] = df_pagos[col].astype(str).str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                df_pagos[col] = pd.to_numeric(df_pagos[col], errors='coerce').fillna(0.0)
            else:
                df_pagos[col] = 0.0 
                
        # FECHAS
        if 'Fecha Pago' in df_pagos.columns:
            df_pagos['Fecha_Obj'] = pd.to_datetime(df_pagos['Fecha Pago'], dayfirst=True, errors='coerce')
        else:
            df_pagos['Fecha_Obj'] = pd.NaT

        # CRUCE (MERGE) Y MOTOR DE RESCATE CON NORMALIZACIÓN
        df_final = pd.merge(df_pagos, df_rubros_maestro[['Ente', 'Rubro', 'Nombre_Referencia']], on='Ente', how='left')
        
        def rescatar_rubro(row):
            if pd.notna(row['Rubro']): 
                return row['Rubro']
            
            # Si el cruce numérico falló, normalizamos el nombre y lo buscamos en el diccionario
            nombre_limpio = clean_str(row['Nombre_CSV'])
            if dict_nombres and nombre_limpio in dict_nombres:
                return dict_nombres[nombre_limpio]
            return 'SIN RUBRO'

        df_final['Rubro'] = df_final.apply(rescatar_rubro, axis=1)
        df_final['Prestador'] = df_final['Nombre_Referencia'].fillna(df_final['Nombre_CSV'])
        
        prestadores_sin_rubro = df_final[df_final['Rubro'] == 'SIN RUBRO'][['Ente', 'Prestador']].drop_duplicates()
        if not prestadores_sin_rubro.empty:
            with st.sidebar.expander(f"⚠️ {len(prestadores_sin_rubro)} Entes SIN RUBRO"):
                st.dataframe(prestadores_sin_rubro, hide_index=True)

        # ==========================================
        # 4. INTERFAZ: TABS, MÉTRICAS Y GRÁFICOS
        # ==========================================
        if not df_final['Fecha_Obj'].dropna().empty:
            fecha_maxima = df_final['Fecha_Obj'].max()
            fecha_minima = df_final['Fecha_Obj'].min()
            fecha_max_str = fecha_maxima.strftime("%d/%m/%Y")
            fecha_min_str = fecha_minima.strftime("%d/%m/%Y")
            
            df_dia = df_final[df_final['Fecha_Obj'] == fecha_maxima].copy()
            titulo_tab_2 = f"📈 Acumulado Histórico (Desde {fecha_min_str} al {fecha_max_str})"
        else:
            df_dia = df_final.copy()
            fecha_max_str = "Desconocida"
            titulo_tab_2 = "📈 Acumulado Histórico"

        tab1, tab2 = st.tabs([f"📅 Pagos del Día ({fecha_max_str})", titulo_tab_2])
        
        def render_tablero(df_mostrar, es_acumulado=False):
            if df_mostrar.empty:
                st.info("No hay datos para mostrar.")
                return

            col_m1, col_m2, col_m3, col_m4 = st.columns(4)
            col_m1.metric("Total Gastado (Bruto)", formatear_moneda(df_mostrar['Imp.OP'].sum()))
            col_m2.metric("Total Gastado (Neto)", formatear_moneda(df_mostrar['Imp.Neto'].sum()))
            col_m3.metric("Total Retenciones", formatear_moneda(df_mostrar['Imp.Ret.'].sum()))
            if col_anul:
                col_m4.metric("Total Imp.Anul.", formatear_moneda(df_mostrar[col_anul].sum()))
            
            if es_acumulado and not df_mostrar['Fecha_Obj'].dropna().empty:
                st.markdown("---")
                st.subheader("Evolución Diaria de Pagos")
                df_evo = df_mostrar.groupby('Fecha_Obj')[columnas_dinero].sum().reset_index().sort_values('Fecha_Obj')
                
                fig_evo = go.Figure()
                nombres = {'Imp.OP': 'Total Bruto', 'Imp.Neto': 'Total Neto', 'Imp.Ret.': 'Retenciones', col_anul: 'Anulaciones'}
                colores = {'Imp.OP': '#1f77b4', 'Imp.Neto': '#2ca02c', 'Imp.Ret.': '#ff7f0e', col_anul: '#d62728'}
                
                for col in columnas_dinero:
                    fig_evo.add_trace(go.Scatter(
                        x=df_evo['Fecha_Obj'], y=df_evo[col],
                        mode='lines+markers', name=nombres.get(col, col),
                        line=dict(color=colores.get(col, '#333333'), width=2)
                    ))
                
                fig_evo.update_layout(
                    xaxis_title='Fecha de Pago', yaxis_title='Importe ($)',
                    hovermode='x unified', legend_title='Tipo de Importe',
                    xaxis=dict(tickformat="%d/%m/%Y"), yaxis=dict(tickformat="$,.0f"), separators=",."
                )
                st.plotly_chart(fig_evo, use_container_width=True)

            st.markdown("---")
            st.header("Análisis por Rubro")
            
            lista_rubros = ['Todos'] + sorted(df_mostrar['Rubro'].unique().tolist())
            clave_filtro = "filtro_rubro_acum" if es_acumulado else "filtro_rubro_dia"
            rubro_sel = st.selectbox("Seleccioná un rubro para filtrar la tabla y el gráfico:", lista_rubros, key=clave_filtro)
            
            if rubro_sel != 'Todos':
                df_mostrar = df_mostrar[df_mostrar['Rubro'] == rubro_sel]
                if df_mostrar.empty:
                    st.warning("No hay pagos para el rubro seleccionado.")
                    return

            df_resumen = df_mostrar.groupby(['Rubro', 'Prestador'])[columnas_dinero].sum().reset_index()
            df_resumen = df_resumen[(df_resumen[columnas_dinero] != 0).any(axis=1)]
            
            col_tabla, col_grafico = st.columns([2, 1])
            
            with col_tabla:
                st.subheader("Matriz Desplegable por Rubro")
                gb = GridOptionsBuilder.from_dataframe(df_resumen)
                gb.configure_column('Rubro', rowGroup=True, hide=True)
                for col in columnas_dinero:
                    gb.configure_column(col, type=["numericColumn"], valueFormatter=formato_pesos)
                grid_options = gb.build()
                AgGrid(df_resumen, gridOptions=grid_options, height=450, theme='streamlit', allow_unsafe_jscode=True)

            with col_grafico:
                st.subheader("Participación (Bruto)")
                df_torta = df_mostrar.groupby('Rubro')['Imp.OP'].sum().reset_index()
                df_torta = df_torta[df_torta['Imp.OP'] > 0]
                if not df_torta.empty:
                    fig_torta = px.pie(df_torta, values='Imp.OP', names='Rubro', hole=0.4)
                    fig_torta.update_traces(textposition='inside', textinfo='percent+label', hovertemplate='Rubro: %{label}<br>Monto: $ %{value:,.2f}')
                    fig_torta.update_layout(showlegend=False)
                    st.plotly_chart(fig_torta, use_container_width=True)
                else:
                    st.info("No hay pagos brutos para graficar.")

        with tab1:
            render_tablero(df_dia, es_acumulado=False)
            
        with tab2:
            render_tablero(df_final, es_acumulado=True)

        # ==========================================
        # 5. BUSCADOR DETALLADO UBICADO ABAJO
        # ==========================================
        st.markdown("---")
        st.header("🔍 Buscador Detallado de Prestador")
        busqueda_texto = st.text_input("Ingresá el nombre o número de ente de un prestador para ver su historial acumulado:")
        
        if busqueda_texto:
            df_busq = df_final[df_final['Prestador'].str.contains(busqueda_texto, case=False, na=False) | 
                               df_final['Ente'].astype(str).str.contains(busqueda_texto, case=False, na=False)]
            
            if not df_busq.empty:
                st.success(f"Se encontraron {len(df_busq)} registros para '{busqueda_texto}'.")
                col_b1, col_b2, col_b3, col_b4 = st.columns(4)
                col_b1.metric("Bruto", formatear_moneda(df_busq['Imp.OP'].sum()))
                col_b2.metric("Neto", formatear_moneda(df_busq['Imp.Neto'].sum()))
                col_b3.metric("Retenciones", formatear_moneda(df_busq['Imp.Ret.'].sum()))
                if col_anul:
                    col_b4.metric("Anulaciones", formatear_moneda(df_busq[col_anul].sum()))
                
                df_mostrar_b = df_busq[['Fecha_Obj', 'Ente', 'Prestador', 'Rubro'] + columnas_dinero].copy()
                df_mostrar_b['Fecha_Obj'] = df_mostrar_b['Fecha_Obj'].dt.strftime('%d/%m/%Y')
                df_mostrar_b.rename(columns={'Fecha_Obj': 'Fecha'}, inplace=True)
                st.dataframe(df_mostrar_b, use_container_width=True)
            else:
                st.warning("No se encontraron registros para tu búsqueda.")

    except Exception as e:
        st.error(f"Error procesando los datos: {e}")
else:
    st.info("Subí el reporte de pagos en el panel de la izquierda para comenzar.")
