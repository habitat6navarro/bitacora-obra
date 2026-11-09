import streamlit as st
from supabase import create_client, Client
from datetime import date
import io
import uuid
import requests
from PIL import Image
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="Bitácora de Obra", page_icon="🏗️", layout="wide")

# --- CONEXIÓN A SUPABASE ---
@st.cache_resource
def init_connection() -> Client:
    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_connection()
except Exception as e:
    st.error("⚠️ Configura tus credenciales SUPABASE_URL y SUPABASE_KEY en `.streamlit/secrets.toml` o en Streamlit Cloud.")
    st.stop()

# --- COMPRESIÓN AUTOMÁTICA DE IMÁGENES ---
def comprimir_imagen(uploaded_file, max_width=1280, quality=80):
    img = Image.open(uploaded_file)
    if img.mode in ("RGBA", "P"):
        img = img.convert("RGB")
    if img.width > max_width:
        ratio = max_width / float(img.width)
        new_height = int(float(img.height) * ratio)
        img = img.resize((max_width, new_height), Image.Resampling.LANCZOS)
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=quality, optimize=True)
    buffer.seek(0)
    return buffer.getvalue()

# --- FUNCIÓN GENERADORA DE WORD (.DOCX) ---
def generar_word_bitacora(bitacora, personal, actividades, fotos):
    doc = Document()
    
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)

    titulo = doc.add_heading('BITÁCORA DIARIA DE OBRA', level=1)
    titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 1. Tabla de Datos Generales
    t_gen = doc.add_table(rows=3, cols=2)
    t_gen.style = 'Table Grid'
    t_gen.cell(0, 0).text = f"PROYECTO: {bitacora['proyecto']}"
    t_gen.cell(0, 1).text = f"FECHA: {bitacora['fecha']}"
    t_gen.cell(1, 0).text = f"RESIDENTE: {bitacora['residente']}"
    t_gen.cell(1, 1).text = f"HORAS LLUVIA: {bitacora['horas_lluvia']} hrs"
    t_gen.cell(2, 0).text = f"CLIMA MAÑANA: {bitacora['clima_manana']}"
    t_gen.cell(2, 1).text = f"CLIMA TARDE: {bitacora['clima_tarde']}"
    
    doc.add_paragraph()

    # 2. Tabla de Personal en Obra
    doc.add_heading('1. PERSONAL EN OBRA', level=2)
    if personal:
        # Resumen por cargo
        conteo_cargos = {}
        for p in personal:
            conteo_cargos[p["cargo"]] = conteo_cargos.get(p["cargo"], 0) + 1
        resumen_txt = " | ".join([f"{cargo}: {cant}" for cargo, cant in conteo_cargos.items()])
        p_res = doc.add_paragraph()
        p_res.add_run(f"Resumen de Fuerza Laboral ({len(personal)} personas): ").bold = True
        p_res.add_run(resumen_txt)

        t_per = doc.add_table(rows=1, cols=3)
        t_per.style = 'Table Grid'
        hdr_per = t_per.rows[0].cells
        hdr_per[0].text = 'Nombre del Trabajador'
        hdr_per[1].text = 'Cargo / Rol'
        hdr_per[2].text = 'Cuadrilla / Contratista'
        
        for p in personal:
            row = t_per.add_row().cells
            row[0].text = p['nombre_trabajador']
            row[1].text = p['cargo']
            row[2].text = p['cuadrilla_contratista']
    else:
        doc.add_paragraph("Sin registro de personal en esta fecha.")

    # 3. Actividades Ejecutadas por Frente
    doc.add_heading('2. ACTIVIDADES EJECUTADAS POR FRENTE', level=2)
    if actividades:
        t_act = doc.add_table(rows=1, cols=3)
        t_act.style = 'Table Grid'
        hdr_act = t_act.rows[0].cells
        hdr_act[0].text = 'Frente de Trabajo'
        hdr_act[1].text = 'Descripción de la Actividad'
        hdr_act[2].text = 'Cant. / Unidad'
        
        for a in actividades:
            row = t_act.add_row().cells
            row[0].text = a['frente_trabajo']
            row[1].text = a['descripcion']
            cant_str = f"{a['cantidad']} {a['unidad']}" if a['cantidad'] and float(a['cantidad']) > 0 else "N/A"
            row[2].text = cant_str
    else:
        doc.add_paragraph("Sin actividades registradas.")

    # 4. Novedades y Observaciones
    doc.add_heading('3. NOVEDADES, INTERVENTORÍA Y SST', level=2)
    doc.add_paragraph(bitacora['novedades'] if bitacora['novedades'] else "Sin novedades relevantes durante la jornada.")

    # 5. Mosaico Fotográfico (2 columnas)
    doc.add_heading('4. REGISTRO FOTOGRÁFICO', level=2)
    if fotos:
        t_fotos = doc.add_table(rows=0, cols=2)
        t_fotos.style = 'Table Grid'
        t_fotos.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        for i in range(0, len(fotos), 2):
            row_cells = t_fotos.add_row().cells
            for j in range(2):
                if i + j < len(fotos):
                    foto_item = fotos[i + j]
                    cell = row_cells[j]
                    p_img = cell.paragraphs[0]
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    try:
                        resp = requests.get(foto_item['url_foto'], timeout=10)
                        if resp.status_code == 200:
                            img_stream = io.BytesIO(resp.content)
                            run = p_img.add_run()
                            run.add_picture(img_stream, width=Inches(2.7))
                    except Exception:
                        p_img.add_run("[No se pudo cargar la imagen]")
                    
                    p_desc = cell.add_paragraph()
                    p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_desc.add_run(f"[{foto_item['frente_trabajo']}] ").bold = True
                    p_desc.add_run(foto_item['descripcion_foto'] or "")
    else:
        doc.add_paragraph("Sin registro fotográfico adjunto.")

    # 6. Firmas al final
    doc.add_paragraph("\n\n")
    t_firmas = doc.add_table(rows=1, cols=2)
    t_firmas.cell(0, 0).text = "__________________________________\nFIRMA RESIDENTE DE OBRA\n" + bitacora['residente']
    t_firmas.cell(0, 1).text = "__________________________________\nVo.Bo. INTERVENTORÍA / DIRECTOR"

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ PRINCIPAL ---
st.title("🏗️ Bitácora Diaria de Obra")

with st.sidebar:
    st.header("Selección de Jornada")
    proyecto_sel = st.selectbox(
        "Proyecto / Obra", 
        ["Castellana Mall", "Urbanización / Loteo", "Obra General"]
    )
    fecha_sel = st.date_input("Fecha de Bitácora", value=date.today())
    residente_def = st.text_input("Ing. Residente / Responsable", value="Ing. Daniel Navarro")

# Buscar si ya existe bitácora creada para ese proyecto y fecha
res_bitacora = supabase.table("bitacora_diaria").select("*").eq("proyecto", proyecto_sel).eq("fecha", str(fecha_sel)).execute()
bitacora_actual = res_bitacora.data[0] if res_bitacora.data else None

if not bitacora_actual:
    st.info(f"No se ha abierto la bitácora para **{proyecto_sel}** el día **{fecha_sel}**.")
    if st.button("➕ Abrir Bitácora del Día", type="primary"):
        nueva = {
            "proyecto": proyecto_sel,
            "fecha": str(fecha_sel),
            "residente": residente_def,
            "clima_manana": "Soleado",
            "clima_tarde": "Soleado",
            "horas_lluvia": 0,
            "novedades": ""
        }
        supabase.table("bitacora_diaria").insert(nueva).execute()
        st.rerun()
else:
    bitacora_id = bitacora_actual["id"]
    
    # Consultar datos hijos del día
    personal_dia = supabase.table("bitacora_personal").select("*").eq("bitacora_id", bitacora_id).execute().data
    actividades_dia = supabase.table("bitacora_actividades").select("*").eq("bitacora_id", bitacora_id).execute().data
    fotos_dia = supabase.table("bitacora_fotos").select("*").eq("bitacora_id", bitacora_id).execute().data

    # Barra superior de resumen y descarga de Word
    col_m1, col_m2, col_m3, col_m4 = st.columns([1, 1, 1, 1.5])
    col_m1.metric("👷 Personal Hoy", len(personal_dia))
    col_m2.metric("🔨 Actividades", len(actividades_dia))
    col_m3.metric("📷 Fotos", len(fotos_dia))
    
    with col_m4:
        docx_file = generar_word_bitacora(bitacora_actual, personal_dia, actividades_dia, fotos_dia)
        st.download_button(
            label="📄 Descargar Bitácora en Word (.docx)",
            data=docx_file,
            file_name=f"Bitacora_{proyecto_sel.replace(' ', '_')}_{fecha_sel}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    # PESTAÑAS DE TRABAJO
    tab1, tab2, tab3, tab4 = st.tabs([
        "👷 1. Personal", 
        "🔨 2. Actividades por Frente", 
        "📷 3. Fotos de Avance", 
        "🌦️ 4. Clima y Novedades"
    ])

    # --- PESTAÑA 1: PERSONAL (Con clonación del día anterior) ---
    with tab1:
        col_tit, col_btn = st.columns([2.5, 1.5])
        col_tit.subheader("Registro de Asistencia en Obra")
        
        if col_btn.button("🔄 Traer personal del último día", use_container_width=True):
            res_ant = (
                supabase.table("bitacora_diaria")
                .select("id, fecha")
                .eq("proyecto", proyecto_sel)
                .lt("fecha", str(fecha_sel))
                .order("fecha", desc=True)
                .limit(1)
                .execute()
            )
            
            if res_ant.data:
                id_anterior = res_ant.data[0]["id"]
                fecha_ant = res_ant.data[0]["fecha"]
                pers_anterior = (
                    supabase.table("bitacora_personal")
                    .select("nombre_trabajador, cargo, cuadrilla_contratista")
                    .eq("bitacora_id", id_anterior)
                    .execute()
                    .data
                )
                
                if pers_anterior:
                    nombres_hoy = {p["nombre_trabajador"].lower() for p in personal_dia}
                    nuevos_registros = [
                        {
                            "bitacora_id": bitacora_id,
                            "nombre_trabajador": p["nombre_trabajador"],
                            "cargo": p["cargo"],
                            "cuadrilla_contratista": p["cuadrilla_contratista"]
                        }
                        for p in pers_anterior
                        if p["nombre_trabajador"].lower() not in nombres_hoy
                    ]
                    
                    if nuevos_registros:
                        supabase.table("bitacora_personal").insert(nuevos_registros).execute()
                        st.success(f"Se copiaron {len(nuevos_registros)} trabajadores de la jornada del {fecha_ant}.")
                        st.rerun()
                    else:
                        st.info("Todo el personal de la jornada anterior ya está en la lista de hoy.")
                else:
                    st.warning(f"La bitácora del {fecha_ant} no tenía personal registrado.")
            else:
                st.warning("No se encontraron bitácoras anteriores para este proyecto.")

        with st.form("form_personal", clear_on_submit=True):
            c1, c2, c3 = st.columns([2, 1, 1.5])
            nombres_input = c1.text_area(
                "Nombre(s) del trabajador (Puedes pegar varios separados por línea)",
                placeholder="Ej:\nCarlos Pérez\nLuis Gómez\nPedro Martínez"
            )
            cargo_input = c2.selectbox("Cargo", ["Oficial", "Ayudante", "Maestro de Obra", "Operador de Equipo", "Residente / SST", "Otro"])
            cuadrilla_input = c3.text_input("Contratista / Frente", value="Obra Civil")
            
            if st.form_submit_button("➕ Agregar Personal Nuevo"):
                lista_nombres = [n.strip() for n in nombres_input.split("\n") if n.strip()]
                if lista_nombres:
                    registros = [
                        {
                            "bitacora_id": bitacora_id,
                            "nombre_trabajador": nombre,
                            "cargo": cargo_input,
                            "cuadrilla_contratista": cuadrilla_input
                        }
                        for nombre in lista_nombres
                    ]
                    supabase.table("bitacora_personal").insert(registros).execute()
                    st.success(f"Se agregaron {len(lista_nombres)} trabajador(es).")
                    st.rerun()

        if personal_dia:
            st.markdown("---")
            conteo_cargos = {}
            for p in personal_dia:
                conteo_cargos[p["cargo"]] = conteo_cargos.get(p["cargo"], 0) + 1
            resumen_str = " | ".join([f"**{cargo}s:** {cant}" for cargo, cant in conteo_cargos.items()])
            st.caption(f"📊 Resumen en obra hoy: {resumen_str}")

            for p in personal_dia:
                col_p1, col_p2, col_p3, col_p4 = st.columns([3, 2, 2, 1])
                col_p1.write(f"👷 **{p['nombre_trabajador']}**")
                col_p2.write(p['cargo'])
                col_p3.write(p['cuadrilla_contratista'])
                if col_p4.button("❌ Faltó", key=f"del_p_{p['id']}"):
                    supabase.table("bitacora_personal").delete().eq("id", p['id']).execute()
                    st.rerun()

    # --- PESTAÑA 2: ACTIVIDADES POR FRENTE ---
    with tab2:
        st.subheader("Avance Diario por Frente de Trabajo")
        with st.form("form_actividad", clear_on_submit=True):
            f1, f2, f3 = st.columns([2, 1, 1])
            frente_in = f1.text_input("Frente de Trabajo / Ubicación", placeholder="Ej: Losa Parqueadero, Cuarto Eléctrico, Viga Canal...")
            cant_in = f2.number_input("Cantidad ejecutada (Opcional)", min_value=0.0, step=0.5)
            unid_in = f3.selectbox("Unidad", ["glb", "m2", "m3", "ml", "kg", "und"])
            desc_in = st.text_area("Descripción de los trabajos realizados hoy")
            
            if st.form_submit_button("Guardar Actividad"):
                if frente_in and desc_in:
                    supabase.table("bitacora_actividades").insert({
                        "bitacora_id": bitacora_id,
                        "frente_trabajo": frente_in,
                        "descripcion": desc_in,
                        "cantidad": cant_in,
                        "unidad": unid_in
                    }).execute()
                    st.rerun()

        if actividades_dia:
            for a in actividades_dia:
                with st.container(border=True):
                    ca1, ca2 = st.columns([5, 1])
                    ca1.markdown(f"**📍 {a['frente_trabajo']}** — *({a['cantidad']} {a['unidad']})*")
                    ca1.write(a['descripcion'])
                    if ca2.button("🗑️ Eliminar", key=f"del_a_{a['id']}"):
                        supabase.table("bitacora_actividades").delete().eq("id", a['id']).execute()
                        st.rerun()

    # --- PESTAÑA 3: REGISTRO FOTOGRÁFICO ---
    with tab3:
        st.subheader("Subir Fotos del Día")
        with st.form("form_fotos", clear_on_submit=True):
            foto_file = st.file_uploader("Selecciona o toma una foto desde el celular", type=["jpg", "jpeg", "png"])
            frente_foto = st.text_input("Frente o zona de la foto", placeholder="Ej: Cerramiento perimetral")
            desc_foto = st.text_input("Pie de foto / Descripción corta")
            
            if st.form_submit_button("Subir Foto"):
                if foto_file:
                    nombre_archivo = f"{fecha_sel}/{uuid.uuid4()}.jpg"
                    file_bytes = comprimir_imagen(foto_file)
                    
                    supabase.storage.from_("fotos_bitacora").upload(
                        path=nombre_archivo,
                        file=file_bytes,
                        file_options={"content-type": "image/jpeg"}
                    )
                    url_publica = supabase.storage.from_("fotos_bitacora").get_public_url(nombre_archivo)
                    
                    supabase.table("bitacora_fotos").insert({
                        "bitacora_id": bitacora_id,
                        "frente_trabajo": frente_foto or "General",
                        "url_foto": url_publica,
                        "descripcion_foto": desc_foto
                    }).execute()
                    st.success("Foto comprimida y guardada correctamente.")
                    st.rerun()

        if fotos_dia:
            cols_f = st.columns(3)
            for idx, f in enumerate(fotos_dia):
                with cols_f[idx % 3]:
                    st.image(f["url_foto"], caption=f"[{f['frente_trabajo']}] {f['descripcion_foto']}", use_container_width=True)
                    if st.button("🗑️ Borrar foto", key=f"del_f_{f['id']}"):
                        supabase.table("bitacora_fotos").delete().eq("id", f['id']).execute()
                        st.rerun()

    # --- PESTAÑA 4: CLIMA Y NOVEDADES ---
    with tab4:
        st.subheader("Condiciones Climáticas y Observaciones")
        with st.form("form_clima"):
            cl1, cl2, cl3 = st.columns(3)
            opciones_clima = ["Soleado", "Parcialmente Nublado", "Nublado", "Lluvia Ligera", "Lluvia Fuerte / Tormenta"]
            
            c_man = cl1.selectbox("Clima Mañana", opciones_clima, index=opciones_clima.index(bitacora_actual["clima_manana"]))
            c_tar = cl2.selectbox("Clima Tarde", opciones_clima, index=opciones_clima.index(bitacora_actual["clima_tarde"]))
            h_llu = cl3.number_input("Horas perdidas por lluvia", min_value=0.0, max_value=12.0, step=0.5, value=float(bitacora_actual["horas_lluvia"]))
            
            nov_txt = st.text_area(
                "Novedades generales, visitas de Interventoría, pendientes de material o notas SST",
                value=bitacora_actual["novedades"] or "",
                height=150
            )
            
            if st.form_submit_button("Actualizar Clima y Novedades"):
                supabase.table("bitacora_diaria").update({
                    "clima_manana": c_man,
                    "clima_tarde": c_tar,
                    "horas_lluvia": h_llu,
                    "novedades": nov_txt
                }).eq("id", bitacora_id).execute()
                st.success("Datos actualizados.")
                st.rerun()
