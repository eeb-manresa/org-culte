import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime

st.set_page_config(page_title="Gestión de Voluntarios - Iglesia Bautista", page_icon="⛪", layout="wide")

EXCEL_FILE = "planificacion.xlsx"

def load_wb_data():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    # 1. Cargar Planificación
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
    
    # 2. Cargar Voluntarios
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
            
    wb.close()
    return df_plan, vol_dict

df_plan, vol_dict = load_wb_data()

st.title("⛪ Gestión de Cultos y Voluntarios")
st.markdown("Aplicación web interactiva para la coordinación del servicio dominical.")

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Resumen por Culto", 
    "✏️ Asignar / Editar Culto", 
    "📅 Planificación General", 
    "👥 Gestionar Directorio"
])

# --- PESTAÑA 1: RESUMEN EJECUTIVO ---
with tab1:
    st.subheader("Informe Ejecutivo del Domingo")
    if not df_plan.empty:
        dates_list = df_plan["Fecha"].dropna().tolist()
        selected_date = st.selectbox("Selecciona la fecha del culto:", dates_list, key="resumen_date")
        
        if selected_date:
            row_data = df_plan[df_plan["Fecha"] == selected_date].iloc[0]
            st.markdown(f"### Culto: {row_data['Culto / Nota']} ({str(selected_date)[:10]})")
            
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("#### 🚪 Bienvenida y Acceso")
                st.info(f"""**Puerta 1:** {row_data['Puerta 1'] or '-'}
**Puerta 2:** {row_data['Puerta 2'] or '-'}
**Entrada 1:** {row_data['Entrada 1'] or '-'}
**Entrada 2:** {row_data['Entrada 2'] or '-'}""")
                
                st.markdown("#### 📖 Dirección y Palabra")
                st.success(f"""**Presidencia:** {row_data['Presidencia'] or '-'}
**Predicación:** {row_data['Predicación'] or '-'}""")
                
            with col2:
                st.markdown("#### 🎵 Técnica y Música")
                st.warning(f"""**Alabanza:** {row_data['Alabanza'] or '-'}
**Sonido:** {row_data['Sonido'] or '-'}
**Proyección:** {row_data['Proyección'] or '-'}""")
                
                st.markdown("#### 🍞 Ordenanzas y Colecta")
                st.error(f"""**Santa Cena:** {row_data['Santa Cena 1'] or '-'}, {row_data['Santa Cena 2'] or '-'}, {row_data['Santa Cena 3'] or '-'}
**Ofrenda:** {row_data['Ofrenda 1'] or '-'}, {row_data['Ofrenda 2'] or '-'}""")
    else:
        st.warning("No hay datos en la planificación.")

# --- PESTAÑA 2: ASIGNAR / EDITAR DESDE LA WEB ---
with tab2:
    st.subheader("Asignación de Voluntarios por Culto")
    if not df_plan.empty:
        dates_str = [str(d)[:10] for d in df_plan["Fecha"].dropna().tolist()]
        chosen_date_str = st.selectbox("Selecciona fecha a modificar:", dates_str, key="edit_date")
        
        # Encontrar la fila correspondiente en el Excel
        wb_edit = openpyxl.load_workbook(EXCEL_FILE)
        ws_p = wb_edit["Planificación"]
        
        target_row = None
        for r in range(5, ws_p.max_row + 1):
            cell_val = ws_p.cell(row=r, column=1).value
            if cell_val and str(cell_val)[:10] == chosen_date_str:
                target_row = r
                break
                
        if target_row:
            current_nota = ws_p.cell(row=target_row, column=2).value or "Culto General"
            
            with st.form(key="form_asignacion"):
                st.markdown(f"**Modificando el culto del día: {chosen_date_str}**")
                nota_input = st.text_input("Tipo de Culto / Nota", value=current_nota)
                
                # Listas de voluntarios disponibles por categoría
                l_puerta = [""] + vol_dict.get("Puerta", [])
                l_entrada = [""] + vol_dict.get("Entrada", [])
                l_presi = [""] + vol_dict.get("Presidencia", [])
                l_pred = [""] + vol_dict.get("Predicación", [])
                l_alab = [""] + vol_dict.get("Alabanza", [])
                l_son = [""] + vol_dict.get("Sonido", [])
                l_proy = [""] + vol_dict.get("Proyección", [])
                l_cena = [""] + vol_dict.get("Santa Cena", [])
                l_ofer = [""] + vol_dict.get("Ofrenda", [])
                
                def get_current_val(col_idx):
                    val = ws_p.cell(row=target_row, column=col_idx).value
                    return val if val in l_puerta or val in l_entrada or val in l_presi or val in l_pred or val in l_alab or val in l_son or val in l_proy or val in l_cena or val in l_ofer else ""

                col_a, col_b = st.columns(2)
                
                with col_a:
                    st.markdown("#### 🚪 Bienvenida")
                    p1 = st.selectbox("Puerta 1", l_puerta, index=l_puerta.index(ws_p.cell(row=target_row, column=3).value) if ws_p.cell(row=target_row, column=3).value in l_puerta else 0)
                    p2 = st.selectbox("Puerta 2", l_puerta, index=l_puerta.index(ws_p.cell(row=target_row, column=4).value) if ws_p.cell(row=target_row, column=4).value in l_puerta else 0)
                    e1 = st.selectbox("Entrada 1", l_entrada, index=l_entrada.index(ws_p.cell(row=target_row, column=5).value) if ws_p.cell(row=target_row, column=5).value in l_entrada else 0)
                    e2 = st.selectbox("Entrada 2", l_entrada, index=l_entrada.index(ws_p.cell(row=target_row, column=6).value) if ws_p.cell(row=target_row, column=6).value in l_entrada else 0)
                    
                    st.markdown("#### 📖 Palabra")
                    pres = st.selectbox("Presidencia", l_presi, index=l_presi.index(ws_p.cell(row=target_row, column=7).value) if ws_p.cell(row=target_row, column=7).value in l_presi else 0)
                    pred = st.selectbox("Predicación", l_pred, index=l_pred.index(ws_p.cell(row=target_row, column=8).value) if ws_p.cell(row=target_row, column=8).value in l_pred else 0)

                with col_b:
                    st.markdown("#### 🎵 Técnica y Música")
                    alab = st.selectbox("Alabanza", l_alab, index=l_alab.index(ws_p.cell(row=target_row, column=9).value) if ws_p.cell(row=target_row, column=9).value in l_alab else 0)
                    son = st.selectbox("Sonido", l_son, index=l_son.index(ws_p.cell(row=target_row, column=10).value) if ws_p.cell(row=target_row, column=10).value in l_son else 0)
                    proy = st.selectbox("Proyección", l_proy, index=l_proy.index(ws_p.cell(row=target_row, column=11).value) if ws_p.cell(row=target_row, column=11).value in l_proy else 0)
                    
                    st.markdown("#### 🍞 Ordenanzas y Colecta")
                    sc1 = st.selectbox("Santa Cena 1", l_cena, index=l_cena.index(ws_p.cell(row=target_row, column=12).value) if ws_p.cell(row=target_row, column=12).value in l_cena else 0)
                    sc2 = st.selectbox("Santa Cena 2", l_cena, index=l_cena.index(ws_p.cell(row=target_row, column=13).value) if ws_p.cell(row=target_row, column=13).value in l_cena else 0)
                    sc3 = st.selectbox("Santa Cena 3", l_cena, index=l_cena.index(ws_p.cell(row=target_row, column=14).value) if ws_p.cell(row=target_row, column=14).value in l_cena else 0)
                    of1 = st.selectbox("Ofrenda 1", l_ofer, index=l_ofer.index(ws_p.cell(row=target_row, column=15).value) if ws_p.cell(row=target_row, column=15).value in l_ofer else 0)
                    of2 = st.selectbox("Ofrenda 2", l_ofer, index=l_ofer.index(ws_p.cell(row=target_row, column=16).value) if ws_p.cell(row=target_row, column=16).value in l_ofer else 0)

                submitted = st.form_submit_button("💾 Guardar Asignación")
                if submitted:
                    ws_p.cell(row=target_row, column=2, value=nota_input)
                    ws_p.cell(row=target_row, column=3, value=p1)
                    ws_p.cell(row=target_row, column=4, value=p2)
                    ws_p.cell(row=target_row, column=5, value=e1)
                    ws_p.cell(row=target_row, column=6, value=e2)
                    ws_p.cell(row=target_row, column=7, value=pres)
                    ws_p.cell(row=target_row, column=8, value=pred)
                    ws_p.cell(row=target_row, column=9, value=alab)
                    ws_p.cell(row=target_row, column=10, value=son)
                    ws_p.cell(row=target_row, column=11, value=proy)
                    ws_p.cell(row=target_row, column=12, value=sc1)
                    ws_p.cell(row=target_row, column=13, value=sc2)
                    ws_p.cell(row=target_row, column=14, value=sc3)
                    ws_p.cell(row=target_row, column=15, value=of1)
                    ws_p.cell(row=target_row, column=16, value=of2)
                    
                    wb_edit.save(EXCEL_FILE)
                    st.success("¡Asignación actualizada correctamente! Recarga la página para ver los cambios reflejados.")
        wb_edit.close()

# --- PESTAÑA 3: PLANIFICACIÓN GENERAL ---
with tab3:
    st.subheader("Tabla Completa de Planificación")
    st.dataframe(df_plan, use_container_width=True)

# --- PESTAÑA 4: GESTIÓN DE DIRECTORIO DE VOLUNTARIOS ---
with tab4:
    st.subheader("Directorio de Voluntarios por Área")
    st.markdown("Añade o elimina voluntarios de forma rápida sin tocar celdas de Excel.")
    
    col_v1, col_v2 = st.columns(2)
    
    with col_v1:
        st.markdown("### ➕ Añadir Voluntario")
        with st.form("form_add_vol"):
            wb_v = openpyxl.load_workbook(EXCEL_FILE)
            ws_v = wb_v["Voluntarios"]
            
            # Mapeo de áreas a columnas de la hoja Voluntarios
            area_cols = {}
            for col in range(1, ws_v.max_column + 1):
                header = ws_v.cell(row=1, column=col).value
                if header:
                    area_cols[header] = col
            
            selected_area = st.selectbox("Selecciona el Área", list(area_cols.keys()))
            new_name = st.text_input("Nombre y Apellidos del Voluntario")
            
            add_btn = st.form_submit_button("Agregar al Directorio")
            if add_btn and new_name:
                col_idx = area_cols[selected_area]
                # Buscar la primera fila libre en esa columna
                next_r = 2
                while ws_v.cell(row=next_r, column=col_idx).value is not None:
                    next_r += 1
                ws_v.cell(row=next_r, column=col_idx, value=new_name)
                wb_v.save(EXCEL_FILE)
                st.success(f"Se ha añadido a {new_name} en {selected_area}.")
                wb_v.close()
                st.rerun()
            wb_v.close()

    with col_v2:
        st.markdown("### 📋 Listado Actual")
        for area, names in vol_dict.items():
            with st.expander(f"🔹 {area} ({len(names)} voluntarios)"):
                for name in names:
                    st.text(f"• {name}")