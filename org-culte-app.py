
import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime

st.set_page_config(page_title="Gestión de Voluntarios - Iglesia Bautista", page_icon="⛪", layout="wide")

EXCEL_FILE = "planificacion_voluntarios_culto.xlsx"

@st.cache_data
def load_data():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    # Load Planificación
    ws_plan = wb["Planificación"]
    plan_data = []
    for row in range(5, ws_plan.max_row + 1):
        row_vals = [ws_plan.cell(row=row, column=col).value for col in range(1, 17)]
        if row_vals[0] is not None:
            plan_data.append(row_vals)
            
    headers_plan = [
        "Fecha", "Culto / Nota", "Puerta 1", "Puerta 2", "Entrada 1", "Entrada 2",
        "Presidencia", "Predicación", "Alabanza", "Sonido", "Proyección",
        "Santa Cena 1", "Santa Cena 2", "Santa Cena 3", "Ofrenda 1", "Ofrenda 2"
    ]
    df_plan = pd.DataFrame(plan_data, columns=headers_plan)
    
    # Load Voluntarios
    ws_vol = wb["Voluntarios"]
    vol_dict = {}
    for col in range(1, ws_vol.max_column + 1):
        area_name = ws_vol.cell(row=1, column=col).value
        if area_name:
            names = []
            for r in range(2, ws_vol.max_row + 1):
                val = ws_vol.cell(row=r, column=col).value
                if val:
                    names.append(val)
            vol_dict[area_name] = names
            
    return df_plan, vol_dict

df_plan, vol_dict = load_data()

st.title("⛪ Gestión de Cultos y Voluntarios")
st.markdown("Aplicación web para la planificación de servicios dominicales.")

tab1, tab2, tab3 = st.tabs(["📊 Resumen por Culto", "📅 Planificación General", "👥 Gestión de Voluntarios"])

with tab1:
    st.subheader("Informe Ejecutivo del Domingo")
    
    if not df_plan.empty:
        # Date selector
        dates_list = df_plan["Fecha"].dropna().tolist()
        selected_date = st.selectbox("Selecciona la fecha del culto:", dates_list)
        
        if selected_date:
            row_data = df_plan[df_plan["Fecha"] == selected_date].iloc[0]
            
            st.markdown(f"### Culto: {row_data['Culto / Nota']} ({str(selected_date)[:10]})")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### 🚪 Bienvenida y Acceso")
                st.info(f"**Puerta 1:** {row_data['Puerta 1'] or '-'}

**Puerta 2:** {row_data['Puerta 2'] or '-'}

**Entrada 1:** {row_data['Entrada 1'] or '-'}

**Entrada 2:** {row_data['Entrada 2'] or '-'}")
                
                st.markdown("#### 📖 Dirección y Palabra")
                st.success(f"**Presidencia:** {row_data['Presidencia'] or '-'}

**Predicación:** {row_data['Predicación'] or '-'}")
                
            with col2:
                st.markdown("#### 🎵 Técnica y Música")
                st.warning(f"**Alabanza:** {row_data['Alabanza'] or '-'}

**Sonido:** {row_data['Sonido'] or '-'}

**Proyección:** {row_data['Proyección'] or '-'}")
                
                st.markdown("#### 🍞 Ordenanzas y Colecta")
                st.error(f"**Santa Cena:** {row_data['Santa Cena 1'] or '-'}, {row_data['Santa Cena 2'] or '-'}, {row_data['Santa Cena 3'] or '-'}

**Ofrenda:** {row_data['Ofrenda 1'] or '-'}, {row_data['Ofrenda 2'] or '-'}")
    else:
        st.warning("No hay datos en la planificación.")

with tab2:
    st.subheader("Tabla Completa de Planificación")
    st.dataframe(df_plan, use_container_width=True)

with tab3:
    st.subheader("Listado de Voluntarios por Área")
    cols = st.columns(len(vol_dict))
    for i, (area, names) in enumerate(vol_dict.items()):
        with cols[i % len(cols)]:
            st.markdown(f"**{area}**")
            for name in names:
                st.text(f"- {name}")
