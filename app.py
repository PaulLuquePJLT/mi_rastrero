# app.py
# =============================================================
#       Rastrero (Hasbro) – versión Streamlit
# =============================================================
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import re, gc, datetime as dt
import openpyxl
from streamlit_option_menu import option_menu

# -------------------------------------------------------------
# 1. Configuración global y estilos
# -------------------------------------------------------------
st.set_page_config(page_title="Generador de Rastrero Hasbro",
                   page_icon="📦",
                   layout="wide")

FA_LINK = """
<link rel="stylesheet"
      href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.5.0/css/all.min.css">
"""
st.markdown(FA_LINK, unsafe_allow_html=True)

CUSTOM_CSS = """
<style>
h1, h2, h3, h4 { font-family:'Inter', sans-serif; }
.header {
    display:flex; align-items:center; gap:18px;
    background:linear-gradient(135deg,#0d47a1 0%,#1976d2 100%);
    color:#fff; padding:18px 26px; border-radius:8px;
}
.header img { width:70px; height:70px; object-fit:contain;
              background:#fff; border-radius:50%;
              padding:6px; box-shadow:0 2px 6px rgba(0,0,0,.2); }
.status-ok   { color:#2e7d32; }
.status-warn { color:#c62828; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# -------------------------------------------------------------
# 2. Cabecera y barra de estado
# -------------------------------------------------------------
with st.container():
    col1, col2 = st.columns([1, 10])
    with col1:
        st.image(
            "https://images.seeklogo.com/logo-png/24/1/hasbro-logo-png_seeklogo-241282.png",
            width=120
        )
    with col2:
        st.markdown("""
        <h2 style="margin-bottom:0;">Generador de Rastrero Hasbro</h2>
        <small>Developed by: <strong>PJLT</strong></small>
        """, unsafe_allow_html=True)

status_placeholder = st.empty()
progress_bar       = st.progress(0)

def update_status(msg: str, pct: int|None = None, ok: bool = True):
    icon = "check-circle" if ok else "exclamation-triangle"
    clazz = "status-ok" if ok else "status-warn"
    status_placeholder.markdown(
        f"<span class='{clazz}'><i class='fa fa-{icon}'></i> {msg}</span>",
        unsafe_allow_html=True
    )
    if pct is not None:
        progress_bar.progress(pct)

# -------------------------------------------------------------
# 3. Utilidades comunes
# -------------------------------------------------------------
def clean_lote(series: pd.Series) -> pd.Series:
    return (series.astype(str)
                  .str.replace(r'\s+', '', regex=True)
                  .str.replace('\u00A0', '')
                  .str.strip())


def clean_number(col: pd.Series) -> pd.Series:
    s = (col.astype(str)
             .str.replace(r'[^0-9,.-]', '', regex=True)
             .str.replace('.', '', regex=False)
             .str.replace(',', '.', regex=False))
    return pd.to_numeric(s, errors='coerce').fillna(0.0)


def preparar_stock(file_bytes: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(file_bytes))
    df.columns = (df.columns.str.normalize('NFKD')
                            .str.encode('ascii','ignore').str.decode('utf-8')
                            .str.strip())
    df['Lote Proveedor']  = clean_lote(df['Lote Proveedor'])
    df['Cant. Final UMS'] = clean_number(df['Cant. Final UMS'])
    df['Concat_U_A']      = df['Ubicacion'] + df['Cod. Articulo']

    def factor(s):
        n = s.astype(str).str.extract(r'C(\d+(?:\.\d+)?)U', expand=False)
        return pd.to_numeric(n, errors='coerce').fillna(1)
    df['Factor'] = factor(df['Huella'])
    df['Cajas']  = df['Cant. Final UMS'] / df['Factor']
    df['UM']     = 'CAJ'

    def nivel(s):
        ult = s.str[-1]
        return np.where(ult.str.isnumeric() & (ult.astype(int) < 3), 'BAJO', 'ALTO')
    df['Nivel'] = nivel(df['Ubicacion'])

    keep = ['Concat_U_A','Ubicacion','Cod. Articulo','Factor','UM',
            'Nivel','Lote Proveedor','Cant. Final UMS','Cajas']
    grp  = ['Concat_U_A','Ubicacion','Cod. Articulo',
            'Factor','UM','Nivel','Lote Proveedor']
    df_bd = (df[keep].groupby(grp, as_index=False)
                     .agg({'Cant. Final UMS':'sum','Cajas':'sum'})
                     .rename(columns={'Ubicacion':'Ubicación',
                                      'Cod. Articulo':'Cod. Artículo'}))
    return df_bd

# Helpers para Rastrero Out

def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (df.columns.str.normalize('NFKD')
                              .str.encode('ascii','ignore').str.decode('utf-8')
                              .str.strip())
    return df


def factor(s: pd.Series) -> pd.Series:
    return (s.astype(str)
             .str.extract(r'C(\d+(?:\.\d+)?)U', expand=False)
             .astype(float)
             .fillna(1))


def calc_pasillo(u: str) -> str:
    if len(u) < 11:
        return 'Libre'
    if u[3:5] == 'MR':
        return 'Pasillo_1'
    tramo = u[8:11]
    if tramo in ('C06','C07','C08'):
        return 'Pasillo_1'
    if tramo in ('C09','C10'):
        return 'Pasillo_2'
    if tramo in ('C11','C12'):
        return 'Pasillo_3'
    return 'Libre'

# -------------------------------------------------------------
# 4. Estado global en sesión
# -------------------------------------------------------------
if 'state_in' not in st.session_state:
    st.session_state.state_in = {}
if 'state_out' not in st.session_state:
    st.session_state.state_out = {}

# =============================================================
#   5. MÓDULO – RASTRERO IN
# =============================================================
def rastrero_in():
    st.subheader("📥 Rastrero In")
    state = st.session_state.state_in

    c1, c2, c3 = st.columns(3)
    with c1:
        ing_file  = st.file_uploader("Flujo de Ingresos", type='xlsx', key='ing_in')
    with c2:
        stock_file= st.file_uploader("Stock", type='xlsx', key='stock_in')
    with c3:
        tmpl_file = st.file_uploader("Plantilla", type='xlsx', key='tmpl_in')
    fecha = st.date_input("Fecha del reporte", dt.date.today(), key='fecha_in')

    # Ingresos
    if ing_file and 'df_ing' not in state:
        name = ing_file.name
        if not name.startswith("ReportConsultasIngresosFlujoIngresos"):
            update_status("⚠️ Archivo de ingresos incorrecto", 0, False)
        else:
            update_status("Leyendo Flujo Ingresos…", 10)
            df = pd.read_excel(ing_file)
            df.columns = (df.columns.str.normalize('NFKD')
                                        .str.encode('ascii','ignore').str.decode('utf-8')
                                        .str.strip())
            df['Codigo Lote Proveedor'] = clean_lote(df['Codigo Lote Proveedor'])
            state['df_ing'] = df
            state['motivos'] = ['Todos'] + sorted(df['Motivo'].dropna().unique())
            update_status("Ingresos listos ✓", 40)

    # Stock
    if stock_file and 'df_stock' not in state:
        update_status("Procesando Stock…", 10)
        state['df_stock'] = preparar_stock(stock_file.read())
        update_status("Stock listo ✓", 60 if 'df_ing' in state else 40)

    # Selección lotes
    lotes_sel = []
    if 'df_ing' in state:
        motivo = st.selectbox("Motivo", state['motivos'], key='motivo_in')
        df_ing = state['df_ing']
        if motivo != 'Todos':
            df_ing = df_ing[df_ing['Motivo'] == motivo]
        vista = df_ing[['Codigo Lote Proveedor','Referencia']].drop_duplicates().reset_index(drop=True)
        lotes_sel = st.multiselect("Lotes", vista['Codigo Lote Proveedor'], default=vista['Codigo Lote Proveedor'].tolist())

    # Generar
    if st.button("Generar Rastrero", disabled=not(lotes_sel and 'df_stock' in state)):
        base = state['df_stock'].copy()
        base['Lote Proveedor'] = clean_lote(base['Lote Proveedor'])
        filtro = base[base['Lote Proveedor'].isin(lotes_sel)].copy()

        filtro = filtro.rename(columns={
            'Concat_U_A':'Concat_U_A_1','Ubicación':'Ubicación_Z',
            'Cod. Artículo':'Cod. Artículo_Z','Cajas':'Cajas_Z', 'Nivel':'Nivel_Z'
        })
        lookup = base.groupby('Concat_U_A')['Cajas'].sum()
        filtro['Stock_Final'] = filtro['Concat_U_A_1'].map(lookup).fillna(0)
        filtro['Stock_Inicial'] = filtro['Stock_Final'] - filtro['Cajas_Z']
        filtro['Check']=''; filtro['Observaciones']=''

        cols = ['Ubicación_Z','Cod. Artículo_Z','UM',
                'Stock_Inicial','Cajas_Z','Stock_Final','Check','Observaciones','Nivel_Z']
        ras_in = filtro[cols]
        state['ras_in'] = ras_in

        alto = ras_in[ras_in['Nivel_Z']=='ALTO'].drop(columns='Nivel_Z')
        bajo = ras_in[ras_in['Nivel_Z']=='BAJO'].drop(columns='Nivel_Z')

        st.success("Rastrero generado")
        with st.expander("🔼 Nivel ALTO", expanded=True): st.dataframe(alto)
        with st.expander("🔽 Nivel BAJO", expanded=False): st.dataframe(bajo)
        update_status("Rastrero In generado ✓", 100)

    # Descargar
    if tmpl_file and 'ras_in' in state:
        def make_xlsx_in():
            wb = openpyxl.load_workbook(BytesIO(tmpl_file.read()))
            alto = state['ras_in'][state['ras_in']['Nivel_Z']=='ALTO'].drop(columns='Nivel_Z')
            bajo = state['ras_in'][state['ras_in']['Nivel_Z']=='BAJO'].drop(columns='Nivel_Z')
            def paste(ws, df):
                for i,row in enumerate(df.itertuples(index=False), start=13):
                    for j,val in enumerate(row, start=2): ws.cell(row=i,column=j,value=val)
            paste(wb['R_Nivel_Alto'], alto); paste(wb['R_Nivel_Bajo'], bajo)
            for ws in (wb['R_Nivel_Alto'], wb['R_Nivel_Bajo']): ws['I1']=fecha.strftime('%d/%m/%Y')
            out=BytesIO(); wb.save(out); out.seek(0); return out
        fname=f"FORMATO_RASTRERO_INGRESOS_{fecha.strftime('%d.%m.%Y')}.xlsx"
        st.download_button("📥 Descargar Excel", data=make_xlsx_in(), file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =============================================================
# 6. Módulo Rastrero Out
# =============================================================
def calc_pasillo(u: str) -> str:
    if not isinstance(u, str) or len(u) < 11:
        return 'Libre'
    if u[3:5] == 'MR':
        return 'Pasillo_1'
    tramo = u[8:11]
    if tramo in ('C06','C07','C08'):
        return 'Pasillo_1'
    if tramo in ('C09','C10'):
        return 'Pasillo_2'
    if tramo in ('C11','C12'):
        return 'Pasillo_3'
    return 'Libre'


def calc_nivel(u_out: str) -> str:
    if not isinstance(u_out, str) or pd.isna(u_out):
        return ''
    if len(u_out) >= 6 and u_out[4:6] == 'MR':
        return 'B'
    last = u_out.strip()[-1]
    return 'B' if last.isdigit() and int(last) <= 2 else 'A'


def rastrero_out():
    st.subheader("📤 Rastrero Out")
    state = st.session_state.state_out

    # 1) Carga de archivos
    col1, col2, col3 = st.columns(3)
    with col1:
        asig_file = st.file_uploader("📑 Asignación (xlsx)", type="xlsx", key="asig_out")
    with col2:
        stock_file = st.file_uploader("📦 Stock (xlsx)", type="xlsx", key="stock_out")
    with col3:
        tmpl_file = st.file_uploader("🖋️ Plantilla (xlsx)", type="xlsx", key="tmpl_out")

    # 2) Lectura inicial de asignación
    headers = ['Estado','Nro. Picking','Usuario Picking','Cliente','Ubicacion',
               'Cod. Articulo','Articulo','Cant. Pick. UMS','Huella']
    if asig_file and 'df_asig_raw' not in state:
        update_status("Leyendo asignación…", 10)
        df_raw = norm_cols(pd.read_excel(asig_file))
        if not all(h in df_raw.columns for h in headers):
            update_status("⚠️ Cabeceras incorrectas", 0, False)
            return
        df = df_raw[headers].copy()
        df['Factor'] = factor(df['Huella'])
        df['Cajas_x'] = df['Cant. Pick. UMS'] / df['Factor']
        df['Concat1'] = df['Ubicacion'] + df['Cod. Articulo']
        df['Cliente_ext'] = df['Cliente'].str.split('|').str[1].fillna(df['Cliente'])
        state['df_asig_raw'] = df
        st.session_state.picks_sel = sorted(df['Nro. Picking'].dropna().unique())

        # 3) Filtro de Pickings y Fecha lado a lado
    if 'df_asig_raw' in state:
        all_picks = sorted(state['df_asig_raw']['Nro. Picking'].dropna().unique())
        # inicializar picks_sel en session_state una sola vez
        if 'picks_sel' not in st.session_state:
            st.session_state.picks_sel = all_picks
        col_date, col_filter = st.columns([1,2])
        with col_date:
            fecha = st.date_input(
                "📅 Fecha del reporte",
                st.session_state.get('fecha_out', dt.date.today()),
                key="fecha_out"
            )
        with col_filter:
            # multiselect sin default, ligando directamente a session_state
            st.multiselect(
                "🔎 Filtrar Pickings",
                options=all_picks,
                key="picks_sel"
            )

    # 4) Lectura de stock
    if stock_file and 'df_stock' not in state:
        update_status("Leyendo stock…", 10)
        df2 = norm_cols(pd.read_excel(stock_file))
        df2 = df2[['Ubicacion','Cod. Articulo','Cant. Final UMS','Huella']]
        df2['Factor'] = factor(df2['Huella'])
        df2['Cajas_y'] = df2['Cant. Final UMS'] / df2['Factor']
        df2['Concat2'] = df2['Ubicacion'] + df2['Cod. Articulo']
        state['df_stock'] = df2.groupby('Concat2', as_index=False).agg({
            'Cant. Final UMS':'sum','Cajas_y':'sum',
            'Ubicacion':'first','Cod. Articulo':'first'
        })
        update_status("Stock listo ✓", 60)

    # 5) Resúmenes automáticos tras multiselect
    picks = st.session_state.get('picks_sel', [])
    if 'df_asig_raw' in state and picks:
        df_f = state['df_asig_raw'][state['df_asig_raw']['Nro. Picking'].isin(picks)]
        state['tpick'] = df_f.groupby('Nro. Picking', as_index=False).agg({'Cant. Pick. UMS':'sum','Cajas_x':'sum'})
        state['tcli'] = df_f.groupby(['Nro. Picking','Cliente_ext'], as_index=False).agg({'Cant. Pick. UMS':'sum','Cajas_x':'sum'}).rename(columns={'Cliente_ext':'Cliente'})
        state['asign'] = df_f.groupby(['Concat1','Ubicacion','Cod. Articulo'], as_index=False)['Cajas_x'].sum()
        st.markdown("**T_Picking**"); st.dataframe(state['tpick'], use_container_width=True)
        st.markdown("**T_Clientes**"); st.dataframe(state['tcli'], use_container_width=True)
        update_status("Resúmenes listos ✓", 80)

    # 6) Generar rastrero Out
    if st.button("Generar Rastrero Out", disabled=not('asign' in state and 'df_stock' in state)):
        bd = pd.merge(state['asign'], state['df_stock'], left_on='Concat1', right_on='Concat2', how='left', suffixes=('_asig','_stk'))
        bd['UM']='CAJ'; bd['Salidas']=bd['Cajas_x']
        bd['Stock Final']=bd['Cajas_y'].fillna(0); bd['Stock Inicial']=bd['Salidas']+bd['Stock Final']
        bd['Check']=''; bd['Observacion']=''
        bd['Ubicacion_out']=bd['Ubicacion_stk'].combine_first(bd['Ubicacion_asig'])
        bd['Pasillo']=bd['Ubicacion_out'].apply(calc_pasillo)
        bd['Nivel']=bd['Ubicacion_out'].apply(calc_nivel)
        bd['Zona']=bd['Pasillo']+'_'+bd['Nivel']
        bd=bd[bd['Pasillo']!='Libre'].reset_index(drop=True)
        bd=bd.sort_values('Ubicacion_out').reset_index(drop=True)
        cols_export=['Ubicacion_out','Cod. Articulo_asig','UM','Stock Inicial','Salidas','Stock Final','Check','Observacion']
        tablas={}
        for zona, df_tab in bd.groupby('Zona'):
            if zona.startswith(('Pasillo_1','Pasillo_2','Pasillo_3')):
                df_tab=df_tab[cols_export].reset_index(drop=True)
                tablas[zona]=df_tab
                st.markdown(f"### {zona.replace('_',' — ')}"); st.dataframe(df_tab, use_container_width=True)
        state['ras_out']=tablas; update_status("Rastrero Out listo ✓", 100)


    # 7) Descarga a Excel
    if tmpl_file and 'ras_out' in state:
        tmpl_bytes = tmpl_file.read()
        try:
            wb = openpyxl.load_workbook(BytesIO(tmpl_bytes))
        except Exception as e:
            st.error(f"Error al leer la plantilla: {e}")
            return

        faltantes = [z for z in state['ras_out'] if z not in wb.sheetnames]
        if faltantes:
            st.error(f"No se encontraron las hojas: {', '.join(faltantes)} en la plantilla.")
            return

        def make_xlsx_out():
            wb2 = openpyxl.load_workbook(BytesIO(tmpl_bytes))
            for hoja, dfp in state['ras_out'].items():
                ws = wb2[hoja]
                for i, row in enumerate(dfp.itertuples(index=False), start=13):
                    for j, val in enumerate(row, start=2):
                        ws.cell(row=i, column=j, value=val)
                ws['I1'] = fecha.strftime('%d/%m/%Y')
                for idx, pk in enumerate(state['tpick']['Nro. Picking'], start=1):
                    ws.cell(row=idx, column=12, value=pk)
                last_row = 12 + len(dfp)
                if ws.max_row > last_row:
                    ws.delete_rows(last_row+1, ws.max_row - last_row)
                ws.print_area = f"B1:I{last_row}"
            buf = BytesIO(); wb2.save(buf); buf.seek(0); return buf

        fname = f"FORMATO_RASTRERO_SALIDAS_{fecha.strftime('%d.%m.%Y')}.xlsx"
        st.download_button(
            "📥 Descargar Excel Salidas",
            data=make_xlsx_out(),
            file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# -------------------------------------------------------------
# 7. Navegación principal con streamlit-option-menu
# -------------------------------------------------------------
with st.sidebar:
        # Inyectar CSS para centrar verticalmente el nav
    st.markdown(
        """
        <style>
        /* Centra cualquier nav (incluido option_menu) verticalmente */
        [data-testid="stSidebarNav"] {
            margin-top: auto;
            margin-bottom: auto;
        }
        /* Asegura que el sidebar ocupe toda la altura */
        [data-testid="stSidebar"] > div:first-child {
            display: flex;
            flex-direction: column;
            height: 200vh;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    # Logo centrado
    st.markdown(
        """
        <div style="text-align: center; margin-bottom: 1rem;">
            <img src="https://enlinea.dinet.com.pe/e/images/logo-login.png" width="250">
        </div>
        """,
        unsafe_allow_html=True
    )

    # Espacio opcional
    st.markdown(" ")

    selected = option_menu(
        menu_title=None,
        options=[" Inicio", " Rastrero In", " Rastrero Out"],
        icons=["house", "inbox", "box-arrow-up"],
        menu_icon="cast",
        default_index=0,
        orientation="vertical",
        styles={
            "container": {
                "padding": "1!important"
                # quitamos justify-content para usar el flujo normal
            },
            "nav-link": {
                "font-size": "14px",
                "text-align": "left",
                "margin": "3px 0"
            },
            "nav-link-selected": {
                "background": "linear-gradient(135deg,#0d47a1 0%,#1976d2 100%)",
                "color": "white",
                "text-align": "left"
            },
        }
    )

# Lógica de enrutamiento
if selected == " Inicio":
    st.markdown("### Bienvenido\nSelecciona una opción del menú.")
    update_status("Elige una opción…", 0)
elif selected == " Rastrero In":
    rastrero_in()
elif selected == " Rastrero Out":
    rastrero_out()
