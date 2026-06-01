import streamlit as st
import plotly.graph_objects as go
import pandas as pd
import re
import requests
from io import BytesIO

# --- 1. CONFIGURACIÓN ---
st.set_page_config(layout="wide", page_title="GALPON 3D - LUNDIN GOLD")

st.markdown("""
    <style>
    .title-lundin {
        color: #002D54; 
        font-size: 40px; 
        font-weight: 800;
        text-align: center; 
        margin-top: -50px; 
        margin-bottom: 15px;
        border-bottom: 2px solid #C5A059;
    }
    .metric-container {
        background-color: #4682B4; 
        padding: 20px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 4px 10px rgba(0,0,0,0.2);
        margin-bottom: 10px;
        border: 1px solid #FFFFFF;
    }
    .metric-value {
        font-size: 32px;
        font-weight: 800;
        color: #FFFFFF; 
        margin-bottom: 0px;
    }
    .metric-label {
        font-size: 14px;
        color: #FFFFFF; 
        text-transform: uppercase;
        letter-spacing: 1px;
        font-weight: 600;
    }
    </style>
    <div class="title-lundin">GALPON 3D</div>
    """, unsafe_allow_html=True)

# Inicializar estados de sesión
if 'llave_resaltada' not in st.session_state:
    st.session_state.update({
        'llave_resaltada': None, 'rack_filtro': None, 
        'lado_filtro': None, 'fila_filtro': None, 'nivel_filtro': None,
        'pozo_resaltado': None 
    })

# --- 2. CONEXIÓN A GOOGLE DRIVE ---
FILE_ID = "1Q9hMT2T5QMgxftxozlo6zEG9ZDOGTv2A" 
URL_DRIVE = f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx"
NOMBRE_HOJA = "Hoja2" 

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

@st.cache_data(ttl=10)
def procesar_datos():
    try:
        response = requests.get(URL_DRIVE)
        xls = pd.ExcelFile(BytesIO(response.content))
        df = pd.read_excel(xls, sheet_name=NOMBRE_HOJA, engine='openpyxl')
        
        df.iloc[:, 0] = df.iloc[:, 0].ffill().astype(str).str.strip().str.upper().str.replace(" ", "")
        df.iloc[:, 1] = df.iloc[:, 1].ffill().astype(str).str.strip().str.upper()
        df.iloc[:, 3] = df.iloc[:, 3].ffill() 
        
        mapa, inventario_lista = {}, []
        for _, row in df.iterrows():
            try:
                rack = str(row.iloc[0])
                lado_raw = str(row.iloc[1])
                lado = "SOLO" if rack in ["A9", "C8"] else ("IZQUIERDA" if "IZQ" in lado_raw else "DERECHA")
                
                f_val, n_val = int(float(row.iloc[3])), int(float(row.iloc[4]))
                id_raw = str(row.iloc[6]).strip().upper() 
                llave = f"{rack}-{lado}-{f_val}-{n_val}"
                
                if id_raw in ["0", "0.0", "VACIO", "VACÍO"]: color, estado = "ROJO", "Disponible"
                elif id_raw in ["NAN", "", "NONE", "NAT"]: color, estado = "NADA", "Vacio"
                else: color, estado = "AMARILLO", id_raw
                
                mapa[llave] = {"color": color, "id": estado, "desde": float(row.iloc[7] or 0), "hasta": float(row.iloc[8] or 0)}
                if color != "NADA":
                    inventario_lista.append({
                        "Rack": rack, "Lado": lado, "Fila": f_val, "Nivel": n_val, 
                        "Contenido": estado, "Desde": float(row.iloc[7] or 0), "Hasta": float(row.iloc[8] or 0), "Key": llave
                    })
            except: continue
        return mapa, pd.DataFrame(inventario_lista)
    except Exception as e:
        st.error(f"Error al cargar datos: {e}")
        return {}, pd.DataFrame()

@st.cache_data
def generar_excel(df):
    output = BytesIO()
    df_export = df.copy()
    df_export.rename(columns={"Nivel": "Altura (Piso)"}, inplace=True)
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Resultados_Pozo')
    return output.getvalue()

mapa_datos, df_inventario = procesar_datos()

# --- 3. RESUMEN DE INVENTARIO ---
if not df_inventario.empty:
    vacios = len(df_inventario[df_inventario["Contenido"] == "Disponible"])
    ocupados = len(df_inventario[df_inventario["Contenido"] != "Disponible"])
    total = len(df_inventario)

    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(f'''<div class="metric-container">
            <div class="metric-label">Capacidad Total</div>
            <div class="metric-value">{total}</div>
        </div>''', unsafe_allow_html=True)
    with c2:
        st.markdown(f'''<div class="metric-container" style="border-bottom: 4px solid #E74C3C;">
            <div class="metric-label">Espacios Vacíos</div>
            <div class="metric-value">{vacios}</div>
        </div>''', unsafe_allow_html=True)
    with c3:
        st.markdown(f'''<div class="metric-container" style="border-bottom: 4px solid #F1C40F;">
            <div class="metric-label">Cajas Llenas</div>
            <div class="metric-value">{ocupados}</div>
        </div>''', unsafe_allow_html=True)

st.markdown("---")

# --- 4. BUSCADORES ---
st.sidebar.title("🔍 Buscador")

if st.sidebar.button("🔄 Restablecer Vista", use_container_width=True):
    st.session_state.llave_resaltada = None
    st.session_state.pozo_resaltado = None
    st.session_state.rack_filtro = None
    st.session_state.lado_filtro = None
    st.rerun()

st.sidebar.markdown("---")

tab1, tab2 = st.sidebar.tabs(["Ubicación", "Hole ID"])

with tab1:
    if not df_inventario.empty:
        rack_sel = st.selectbox("Rack:", sorted(df_inventario["Rack"].unique(), key=natural_sort_key))
        df_rack = df_inventario[df_inventario["Rack"] == rack_sel]
        lado_sel = st.selectbox("Seleccionar Lado:", sorted(df_rack["Lado"].unique()))
        detalle = df_rack[df_rack["Lado"] == lado_sel].sort_values(by=["Fila", "Nivel"])
        seleccion_caja = st.selectbox("Ubicar Caja:", ["Ver Todo"] + detalle["Contenido"].tolist())
        if st.button("📍 Ubicar"):
            st.session_state.pozo_resaltado = None 
            if seleccion_caja != "Ver Todo":
                r = detalle[detalle["Contenido"] == seleccion_caja].iloc[0]
                st.session_state.update({'llave_resaltada': r["Key"], 'rack_filtro': rack_sel, 'lado_filtro': lado_sel, 'fila_filtro': r["Fila"], 'nivel_filtro': r["Nivel"]})
            else: st.session_state.llave_resaltada = None

with tab2:
    h_input = st.text_input("Hole ID:").upper()
    opcion_busqueda = st.radio("Opciones de búsqueda:", ["Buscar por metro específico", "Buscar todo el pozo"])
    
    m_input = 0.0
    if opcion_busqueda == "Buscar por metro específico":
        m_input = st.number_input("Metro a buscar:", value=0.0)
        
    if st.button("🔍 Buscar"):
        st.session_state.llave_resaltada = None
        st.session_state.pozo_resaltado = None
        
        if h_input and not df_inventario.empty:
            if opcion_busqueda == "Buscar por metro específico":
                res = df_inventario[(df_inventario["Contenido"] == h_input) & (df_inventario["Desde"] <= m_input) & (df_inventario["Hasta"] >= m_input)]
                if not res.empty:
                    r = res.iloc[0]
                    st.session_state.update({'llave_resaltada': r["Key"], 'rack_filtro': r["Rack"], 'lado_filtro': r["Lado"], 'fila_filtro': r["Fila"], 'nivel_filtro': r["Nivel"]})
                    st.success(f"Encontrada en {r['Rack']}")
                else: st.error("No encontrada")
            else:
                res = df_inventario[df_inventario["Contenido"] == h_input]
                if not res.empty:
                    st.session_state.pozo_resaltado = h_input
                else:
                    st.error("No se encontraron cajas para este pozo.")

    if st.session_state.get('pozo_resaltado') and h_input == st.session_state.pozo_resaltado:
        res = df_inventario[df_inventario["Contenido"] == h_input]
        st.success(f"✅ Se encontraron {len(res)} cajas para el pozo {h_input}.")
        
        df_mostrar = res[["Rack", "Lado", "Fila", "Nivel", "Desde", "Hasta"]].copy()
        df_mostrar.rename(columns={"Nivel": "Altura"}, inplace=True)
        st.dataframe(df_mostrar, use_container_width=True)
        
        excel_data = generar_excel(res[["Rack", "Lado", "Fila", "Nivel", "Desde", "Hasta", "Contenido"]])
        st.download_button(
            label="🟩 📥 Descargar Excel del Pozo",
            data=excel_data,
            file_name=f"Pozo_{h_input}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

posicion_caja = {}
cajas_masivas_coords = [] 
fig = go.Figure()

# --- 5. DIBUJO ---
def agregar_edificio(x, y, z, dx, dy, dz, color, nombre):
    fig.add_trace(go.Mesh3d(x=[x, x, x+dx, x+dx, x, x, x+dx, x+dx], y=[y, y+dy, y+dy, y, y, y+dy, y+dy, y], z=[z, z, z, z, z+dz, z+dz, z+dz, z+dz], i=[7,0,0,0,4,4,6,6,4,0,3,2], j=[3,4,1,2,5,6,5,2,0,1,6,3], k=[0,7,2,3,6,7,1,1,5,5,7,6], color=color, opacity=1.0, flatshading=True, name=nombre))
    xb, yb, zb = [x,x+dx,x+dx,x,x,None,x,x+dx,x+dx,x,x,None,x,x,None,x+dx,x+dx], [y,y,y+dy,y+dy,y,None,y,y,y+dy,y+dy,y,None,y,y+dy,None,y,y+dy], [z,z,z,z,z,None,z+dz,z+dz,z+dz,z+dz,z+dz,None,z,z,None,z,z]
    fig.add_trace(go.Scatter3d(x=xb, y=yb, z=zb, mode='lines', line=dict(color='black', width=3), showlegend=False))

def dibujar_hilera(lista, x_off, y_pos):
    for col_idx, (r, l) in enumerate(lista):
        r_k = r.upper()
        off_x = (1.5 if r_k.startswith("C") and int(r_k[1:])>=4 else 0) + (1.5 if r_k.startswith("C") and int(r_k[1:])>=8 else 0)
        x_p = (col_idx + x_off + off_x) * 1.4
        for f in range(1, 7):
            for n in range(1, 5):
                key = f"{r_k}-{l}-{f}-{n}"
                data = mapa_datos.get(key, {"color": "NADA"})
                if data["color"] == "NADA": continue
                y_p, z_p = y_pos + ((f-1) * 0.9), n * 0.85
                dx, dy, dz = (1.1 if l=="SOLO" else 0.8), 0.8, 0.7
                x_adj = x_p - 0.15 if l=="SOLO" else x_p
                
                if st.session_state.get('pozo_resaltado'):
                    if data["id"] == st.session_state.pozo_resaltado:
                        color_c = "#00BFFF" 
                        cajas_masivas_coords.append({'x': x_adj + (dx/2), 'y': y_p + (dy/2), 'z': z_p + dz, 'rack': r_k, 'fila': f, 'altura': n})
                    else:
                        color_c = "#E74C3C" if data["color"] == "ROJO" else "#F1C40F"
                else:
                    if st.session_state.rack_filtro == r_k and st.session_state.lado_filtro == l:
                        if key == st.session_state.llave_resaltada: 
                            color_c = "#00BFFF" 
                            posicion_caja.update({'x': x_adj + (dx/2), 'y': y_p + (dy/2), 'z': z_p + dz})
                        else: continue 
                    else: color_c = "#E74C3C" if data["color"] == "ROJO" else "#F1C40F"
                
                fig.add_trace(go.Mesh3d(x=[x_adj, x_adj, x_adj+dx, x_adj+dx, x_adj, x_adj, x_adj+dx, x_adj+dx], y=[y_p, y_p+dy, y_p+dy, y_p, y_p, y_p+dy, y_p+dy, y_p], z=[z_p, z_p, z_p, z_p, z_p+dz, z_p+dz, z_p+dz, z_p+dz], i=[7,0,0,0,4,4,6,6,4,0,3,2], j=[3,4,1,2,5,6,5,2,0,1,6,3], k=[0,7,2,3,6,7,1,1,5,5,7,6], color=color_c, opacity=1.0, flatshading=True))
                fig.add_trace(go.Scatter3d(x=[x_adj,x_adj+dx,x_adj+dx,x_adj,x_adj,None,x_adj,x_adj+dx,x_adj+dx,x_adj,x_adj,None,x_adj,x_adj,None,x_adj+dx,x_adj+dx], y=[y_p,y_p,y_p+dy,y_p+dy,y_p,None,y_p,y_p,y_p+dy,y_p+dy,y_p,None,y_p,y_p+dy,None,y_p,y_p+dy], z=[z_p,z_p,z_p,z_p,z_p,None,z_p+dz,z_p+dz,z_p+dz,z_p+dz,z_p+dz,None,z_p,z_p,None,z_p,z_p], mode='lines', line=dict(color='black', width=2), showlegend=False, hoverinfo="skip"))

f1_A = sorted([("A"+str(i), l) for i in range(1,14) for l in (["SOLO"] if i==9 else ["IZQUIERDA", "DERECHA"])], key=lambda x: natural_sort_key(x[0]))
f2_B = sorted([("B"+str(i), l) for i in range(1,12) for l in ["IZQUIERDA", "DERECHA"]], key=lambda x: natural_sort_key(x[0]))
f3_C = sorted([("C"+str(i), l) for i in range(1,19) for l in (["SOLO"] if i==8 else ["IZQUIERDA", "DERECHA"])], key=lambda x: natural_sort_key(x[0]))

dibujar_hilera(f1_A, 0, 7.5)    
dibujar_hilera(f2_B, 3.0, -1.0) 
dibujar_hilera(f3_C, 0, -10.0) 

x_ref = 35.0 
agregar_edificio(x_ref, 0.5, 0, 2, 1.2, 2, "#F1948A", "B.M")
agregar_edificio(x_ref, 2.0, 0, 2, 1.2, 2, "#F1948A", "B.F")
agregar_edificio(x_ref + 3.5, 0.5, 0, 3, 2.5, 2.2, "#D1F2EB", "BULK")
agregar_edificio(x_ref + 7.0, 0.5, 0, 3, 2.5, 2.2, "#D6DBDF", "LAB")
agregar_edificio(x_ref + 10.5, 0.5, 0, 3, 2.5, 2.2, "#ABEBC6", "CAFÉ")
agregar_edificio(x_ref, 7.0, 0, 13.5, 5, 3.5, "#AED6F1", "CORTE")

x_max = 55.0 
fig.add_trace(go.Mesh3d(x=[0, x_max, x_max, 0], y=[5.5, 5.5, 6.7, 6.7], z=[0.05]*4, color="#F1C40F", opacity=1.0, name="P1"))
fig.add_trace(go.Mesh3d(x=[0, x_max, x_max, 0], y=[-2.8, -2.8, -1.6, -1.6], z=[0.05]*4, color="#F1C40F", opacity=1.0, name="P2"))
fig.add_trace(go.Mesh3d(x=[-5, x_max + 5, x_max + 5, -5], y=[-15, -15, 15, 15], z=[0, 0, 0, 0], color="#FFFFFF", opacity=1.0, hoverinfo="skip"))

fig.add_trace(go.Scatter3d(x=[12, 12, 12], y=[10.5, 1.5, -7.5], z=[6, 6, 6], mode='text', text=["BLOQUE A", "BLOQUE B", "BLOQUE C"], textposition="middle center", textfont=dict(color="black", size=22, family="Arial Black"), showlegend=False, hoverinfo="skip"))

escena_config = dict(xaxis_visible=False, yaxis_visible=False, zaxis_visible=False, aspectmode='data')
anotaciones_pines = []

if st.session_state.get('pozo_resaltado') and cajas_masivas_coords:
    clusters = []
    for caja in cajas_masivas_coords:
        agregada = False
        for cluster in clusters:
            for caja_ref in cluster:
                dist = ((caja['x'] - caja_ref['x'])**2 + (caja['y'] - caja_ref['y'])**2 + (caja['z'] - caja_ref['z'])**2)**0.5
                if dist <= 3.5: cluster.append(caja); agregada = True; break
            if agregada: break
        if not agregada: clusters.append([caja])
            
    for cluster in clusters:
        representante = cluster[0] 
        anotaciones_pines.append(dict(x=representante['x'], y=representante['y'], z=representante['z'] + 0.5, ax=0, ay=-50, text="📍 AQUÍ", showarrow=True, arrowhead=2, arrowsize=1.5, arrowwidth=3, standoff=5, font=dict(color="white", size=16, family="Arial Black"), bgcolor="rgba(0,0,0,0.85)"))

elif posicion_caja:
    anotaciones_pines.append(dict(x=posicion_caja['x'], y=posicion_caja['y'], z=posicion_caja['z'] + 0.5, ax=0, ay=-50, text=f"📍 AQUÍ ➔ R:{st.session_state.rack_filtro} | Fila: {st.session_state.fila_filtro} | Altura: {st.session_state.nivel_filtro}", showarrow=True, arrowhead=2, arrowsize=1.5, arrowwidth=3, standoff=5, font=dict(color="white", size=16, family="Arial Black"), bgcolor="rgba(0,0,0,0.85)"))

escena_config['annotations'] = anotaciones_pines
fig.update_layout(template="plotly_dark", height=850, margin=dict(l=0, r=0, b=0, t=0), scene=escena_config)

st.plotly_chart(fig, use_container_width=True, key=f"grafico_{st.session_state.llave_resaltada}_{st.session_state.pozo_resaltado}")