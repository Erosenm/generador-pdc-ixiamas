import streamlit as st
import json
import re
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Intentamos importar Groq
try:
    from groq import Groq
except Exception:
    Groq = None

st.set_page_config(page_title="Generador de Planes Curriculares", layout="wide")

# --- Helpers de Configuración ---
def get_groq_client():
    api_key = st.secrets.get("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("Falta la clave de API en los secretos.")
    if Groq is None:
        raise RuntimeError("No se pudo importar `groq`.")
    return Groq(api_key=api_key)

def extract_json_from_text(text):
    try:
        return json.loads(text)
    except Exception:
        pass
    regex_obj = re.compile(r"(\{.*\})", re.DOTALL)
    regex_arr = re.compile(r"(\[.*\])", re.DOTALL)
    for regex in (regex_obj, regex_arr):
        m = regex.search(text)
        if m:
            candidate = m.group(1)
            try:
                return json.loads(candidate)
            except Exception:
                continue
    raise ValueError("No se pudo extraer JSON válido de la respuesta de IA.")

def set_cell_bg(cell, color_hex="D9E2F3"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def insert_formatted_text(cell, text):
    """Busca palabras clave y las pone en negrita automáticamente"""
    if not text:
        return
    p = cell.paragraphs[0] if cell.paragraphs else cell.add_paragraph()
    # Palabras a poner en negrita (Solo SER, SABER, HACER)
    keywords = [r"Lectura:", r"Práctica:", r"Teoría:", r"Valoración:", r"Producción:", r"SER:", r"SABER:", r"HACER:"]
    pattern = '(' + '|'.join(keywords) + ')'
    
    parts = re.split(pattern, text)
    for part in parts:
        run = p.add_run(part)
        if part in [k.replace("\\", "") for k in keywords]:  
            run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(9)

def build_prompt(tema, semanas, nivel, area, puntos_clave):
    weeks = int(semanas)
    instruccion_extra = f"\nEl profesor solicita que incluyas estos puntos clave o enfoques: {puntos_clave}" if puntos_clave else ""
    
    prompt = f"""
Actúa como un pedagogo boliviano experto en diseño curricular. 
Genera un Plan de Desarrollo Curricular (PDC) formal y riguroso.

Tema a desarrollar: "{tema}"
Área: "{area}"
Nivel educativo: "{nivel}"
Duración: {weeks} semanas.
Contexto socio-comunitario: Municipio de Ixiamas, Norte de La Paz.{instruccion_extra}

Devuelve ÚNICAMENTE un objeto JSON válido con la siguiente estructura exacta:

{{
  "objetivo_holistico_nivel": "Redacta el objetivo general del nivel (Ser, Saber, Hacer). Párrafo continuo.",
  "objetivo_aprendizaje": "Redacta el objetivo específico para este tema.",
  "semanas": [
    {{
      "semana": 1,
      "contenidos": "Subtítulo o contenido a avanzar.",
      "momentos": "Lectura: 10 minutos... Práctica: ... Teoría: ... Valoración: ... Producción: ...",
      "recursos": "Lista de materiales. Incluye recursos de Ixiamas.",
      "periodos": "2",
      "criterios_evaluacion": "SER:\\n- Valora...\\nSABER:\\n- Identifica...\\nHACER:\\n- Elabora..."
    }}
  ],
  "adaptaciones_curriculares": "Párrafo sobre adaptaciones para estudiantes con dificultades."
}}
Genera exactamente {weeks} objetos dentro del arreglo "semanas".
"""
    return prompt

def generar_docx(datos_formulario, plan_json):
    doc = Document()
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)
    
    # Forzar idioma Español
    rPr = font.element.get_or_add_rPr()
    lang = OxmlElement('w:lang')
    lang.set(qn('w:val'), 'es-ES')
    rPr.append(lang)

    # Configurar hoja a Carta y Horizontal
    for section in doc.sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11.0)
        section.page_height = Inches(8.5)
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Títulos
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run_t1 = p_titulo.add_run("PLAN DE DESARROLLO CURRICULAR PARA EDUCACIÓN SECUNDARIA COMUNITARIA PRODUCTIVA\n")
    run_t1.bold = True
    run_t1.font.size = Pt(11)
    
    run_t2 = p_titulo.add_run(f"PLAN DE DESARROLLO CURRICULAR Nº {datos_formulario.get('nro_plan', '1')}")
    run_t2.bold = True
    run_t2.font.size = Pt(11)

    p_datos = doc.add_paragraph()
    p_datos.add_run("1.  Datos referenciales").bold = True

    # --- TABLA DE DATOS REFERENCIALES (Anchos fijos para que no se estire de más) ---
    t_datos = doc.add_table(rows=5, cols=4)
    t_datos.style = 'Table Grid'
    t_datos.autofit = False
    
    # Ancho total: 9 pulgadas (centrado visualmente)
    for row in t_datos.rows:
        row.cells[0].width = Inches(1.5)
        row.cells[1].width = Inches(3.0)
        row.cells[2].width = Inches(1.5)
        row.cells[3].width = Inches(3.0)

    def set_cell_bold(cell, text):
        cell.text = text
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True

    set_cell_bold(t_datos.cell(0,0), "Distrito educativo")
    t_datos.cell(0,1).text = "IXIAMAS"
    set_cell_bold(t_datos.cell(0,2), "Unidad Educativa")
    t_datos.cell(0,3).text = datos_formulario.get("unidad", "").upper()

    set_cell_bold(t_datos.cell(1,0), "Nivel")
    t_datos.cell(1,1).text = datos_formulario.get("nivel", "").upper()
    set_cell_bold(t_datos.cell(1,2), "Año de escolaridad/Paralelo")
    t_datos.cell(1,3).text = datos_formulario.get("ano", "").upper()

    set_cell_bold(t_datos.cell(2,0), "Maestra/o")
    t_datos.cell(2,1).text = datos_formulario.get("docente", "").upper()
    t_datos.cell(2,1).merge(t_datos.cell(2,3))

    set_cell_bold(t_datos.cell(3,0), "Área")
    t_datos.cell(3,1).text = datos_formulario.get("area", "").upper() 
    t_datos.cell(3,1).merge(t_datos.cell(3,3))

    set_cell_bold(t_datos.cell(4,0), "Trimestre")
    t_datos.cell(4,1).text = datos_formulario.get("trimestre", "").upper()
    t_datos.cell(4,1).merge(t_datos.cell(4,3))
    
    row_fecha = t_datos.add_row()
    set_cell_bold(row_fecha.cells[0], "FECHA")
    row_fecha.cells[1].text = datos_formulario.get("fechas", "").upper()
    row_fecha.cells[1].merge(row_fecha.cells[3])

    doc.add_paragraph("") 

    # --- 2. DESARROLLO Y OBJETIVO HOLÍSTICO (FUERA DE LA TABLA) ---
    p_des = doc.add_paragraph()
    p_des.add_run("2.  Desarrollo\n").bold = True
    p_des.add_run("Objetivo holístico de nivel\n").bold = True
    p_des.add_run(plan_json.get("objetivo_holistico_nivel", ""))

    # --- TABLA PRINCIPAL (6 COLUMNAS) ---
    semanas = plan_json.get("semanas", [])
    num_semanas = len(semanas)
    
    # Filas: 1 de cabeceras + N semanas + 1 adaptaciones
    t_plan = doc.add_table(rows=(1 + num_semanas + 1), cols=6)
    t_plan.style = 'Table Grid'
    t_plan.autofit = False

    # Definir anchos para que quepa en la hoja horizontal (Total 10 pulgadas)
    anchos = [Inches(1.3), Inches(1.5), Inches(2.5), Inches(1.5), Inches(0.7), Inches(2.5)]

    # Fila 0: Cabeceras con fondo celeste
    cabeceras = ["Objetivo de aprendizaje", "Contenidos", "Momentos del proceso formativo", "Recursos", "Períodos", "Criterios de evaluación"]
    for i, txt in enumerate(cabeceras):
        celda = t_plan.cell(0,i)
        celda.width = anchos[i]
        set_cell_bg(celda, "D9E2F3")
        p = celda.paragraphs[0]
        p.add_run(txt).bold = True
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Filas de Semanas
    for idx, s in enumerate(semanas):
        fila_actual = idx + 1
        
        # Ajustar anchos de las celdas de la fila
        for col_idx in range(6):
            t_plan.cell(fila_actual, col_idx).width = anchos[col_idx]

        # Col 0: Objetivo de Aprendizaje (se llena solo en la primera semana y luego se fusionará)
        if idx == 0:
            t_plan.cell(fila_actual, 0).text = plan_json.get("objetivo_aprendizaje", "")
            
        # Col 1: Contenidos
        celda_cont = t_plan.cell(fila_actual, 1)
        p_cont = celda_cont.paragraphs[0]
        p_cont.add_run(f"Semana {s.get('semana', idx+1)}\n").bold = True
        p_cont.add_run(s.get("contenidos", ""))
        
        # Col 2, 3, 4, 5 (Usamos la función que pone negritas automáticamente)
        insert_formatted_text(t_plan.cell(fila_actual, 2), s.get("momentos", ""))
        insert_formatted_text(t_plan.cell(fila_actual, 3), s.get("recursos", "").replace("\\n", "\n"))
        
        t_plan.cell(fila_actual, 4).text = str(s.get("periodos", ""))
        insert_formatted_text(t_plan.cell(fila_actual, 5), s.get("criterios_evaluacion", "").replace("\\n", "\n"))

    # Fusionar la columna del Objetivo de Aprendizaje hacia abajo si hay más de 1 semana
    if num_semanas > 1:
        t_plan.cell(1, 0).merge(t_plan.cell(num_semanas, 0))

    # Última fila: Adaptaciones
    fila_adapt = num_semanas + 1
    cell_adapt = t_plan.cell(fila_adapt, 0)
    p_adapt = cell_adapt.paragraphs[0]
    p_adapt.add_run("Adaptaciones curriculares:\n").bold = True
    p_adapt.add_run(plan_json.get("adaptaciones_curriculares", ""))
    cell_adapt.merge(t_plan.cell(fila_adapt, 5))

    # Asegurar fuente Arial tamaño 9 en toda la tabla
    for row in t_plan.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if run.font.name is None:
                        run.font.name = 'Arial'
                    if run.font.size is None:
                        run.font.size = Pt(9)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- Interfaz Streamlit ---
st.title("Generador de Planes Curriculares")

MESES_ESPANOL = {1:"ENERO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}

with st.form("form_pdc"):
    st.subheader("Datos referenciales")
    c1, c2, c3 = st.columns(3)
    with c1:
        nro_plan = st.number_input("Nº de Plan Curricular", min_value=1, value=5, step=1)
        unidad = st.text_input("Unidad Educativa", value="GERMÁN BUSCH")
        nivel = st.selectbox("Nivel", options=["SECUNDARIA", "PRIMARIA"])
        area = st.text_input("Área (Ej: Artes Plásticas y Visuales)", value="ARTES PLÁSTICAS Y VISUALES")
    with c2:
        docente = st.text_input("Maestra/o", value="PEPE PÉREZ")
        ano = st.text_input("Año de Escolaridad", value="Quinto B")
        trimestre = st.selectbox("Trimestre", options=["PRIMERO", "SEGUNDO", "TERCERO"], index=1)
        semanas = st.selectbox("Semanas de duración", options=[1,2,3,4], index=1)
    with c3:
        tema = st.text_input("El Tema a avanzar", placeholder="Ej: La calidad de imagen...")
        st.write("Rango de Fechas")
        f1, f2 = st.columns(2)
        with f1:
            fecha_inicio = st.date_input("Desde")
        with f2:
            fecha_fin = st.date_input("Hasta")
            
    puntos_clave = st.text_area("Puntos clave o enfoques específicos (Opcional)", placeholder="Escribe aquí si quieres que la IA se enfoque en algo específico. Ej: Quiero que en la Práctica hagamos maquetas con material reciclado de la zona.")
            
    submitted = st.form_submit_button("Generar Plan en Word")

if submitted:
    if not tema or tema.strip() == "":
        st.error("Por favor ingrese el campo 'El Tema a avanzar'.")
    else:
        mes_ini = MESES_ESPANOL.get(fecha_inicio.month, "ENERO")
        mes_fin = MESES_ESPANOL.get(fecha_fin.month, "ENERO")
        cadena_fechas = f"DEL {fecha_inicio.day} DE {mes_ini} AL {fecha_fin.day} DE {mes_fin} DE {fecha_fin.year}"

        datos_form = {
            "nro_plan": nro_plan,
            "unidad": unidad,
            "nivel": nivel,
            "ano": ano,
            "docente": docente,
            "area": area,
            "trimestre": trimestre,
            "fechas": cadena_fechas,
        }

        try:
            with st.spinner("Creando estructura de 6 columnas y procesando textos..."):
                client = get_groq_client()
                prompt = build_prompt(tema=tema, semanas=semanas, nivel=nivel, area=area, puntos_clave=puntos_clave)

                response = client.chat.completions.create(
                    messages=[
                        {"role": "system", "content": "Eres un asistente experto. Debes responder ÚNICAMENTE con formato JSON válido."},
                        {"role": "user", "content": prompt}
                    ],
                    model="llama-3.3-70b-versatile",
                    temperature=0.2,
                )

                text = response.choices[0].message.content
                plan_json = extract_json_from_text(text)
                docx_bio = generar_docx(datos_form, plan_json)

                st.success("¡Plan generado exitosamente!")
                now_name = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"PDC_{unidad}_{tema[:10]}_{now_name}.docx"
                st.download_button(
                    label="📥 Descargar Plan Curricular",
                    data=docx_bio.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
            with st.expander("Detalles del error"):
                st.exception(e)