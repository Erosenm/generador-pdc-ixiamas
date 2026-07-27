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

st.set_page_config(page_title="Generador de Planes Curriculares - Ixiamas", layout="wide")

# --- Helpers de Configuración de Word ---
def get_groq_client():
    api_key = st.secrets.get("GROQ_API_KEY")
    if not api_key:
        raise RuntimeError("Falta la clave de API en los secretos (ej. 'GROQ_API_KEY').")
    if Groq is None:
        raise RuntimeError("No se pudo importar `groq`. Revisa el archivo requirements.txt.")
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

def set_cell_bg(cell, color_hex="D9E2F3"): # Color celeste/grisáceo
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def build_prompt(tema, semanas, nivel, contexto_extra=""):
    weeks = int(semanas)
    prompt = f"""
Actúa como un pedagogo boliviano experto en diseño curricular del Ministerio de Educación. 
Genera un Plan de Desarrollo Curricular (PDC) formal y riguroso.

Tema a desarrollar: "{tema}"
Nivel educativo: "{nivel}"
Duración: {weeks} semanas.
Contexto socio-comunitario: Municipio de Ixiamas, Norte de La Paz.

Devuelve ÚNICAMENTE un objeto JSON válido con la siguiente estructura exacta:

{{
  "objetivo_holistico_nivel": "Redacta el objetivo general del nivel según la Ley 070 (Ser, Saber, Hacer, Decidir), mencionando valores sociocomunitarios.",
  "objetivo_aprendizaje": "Redacta el objetivo específico para este tema (Ser, Saber, Hacer, Decidir).",
  "semanas": [
    {{
      "semana": 1,
      "contenidos": "Subtítulo o contenido específico a avanzar.",
      "momentos": "Lectura: 10 minutos de lectura... Práctica: ... Teoría: ... Valoración: ... Producción: ... (Redacta en un solo bloque continuo sin viñetas)",
      "recursos": "Lista de materiales (educativos, analógicos, de la vida). Incluye recursos de Ixiamas.",
      "periodos": "2",
      "criterios_evaluacion": "SER:\\n- Valora...\\nSABER:\\n- Identifica...\\nHACER:\\n- Elabora...\\nDECIDIR:\\n- Promueve..."
    }}
  ],
  "adaptaciones_curriculares": "Redacta un párrafo sobre cómo se adaptará el contenido para estudiantes con dificultades de aprendizaje."
}}
Genera exactamente {weeks} objetos dentro del arreglo "semanas".
"""
    return prompt

def generar_docx(datos_formulario, plan_json):
    doc = Document()
    
    # 1. Configuración de estilos e idioma Español (Adiós líneas rojas)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)
    
    # Forzar el idioma a Español para el corrector ortográfico
    rPr = font.element.get_or_add_rPr()
    lang = OxmlElement('w:lang')
    lang.set(qn('w:val'), 'es-ES')
    rPr.append(lang)

    # 2. Configurar hoja a Carta y Horizontal (Landscape)
    for section in doc.sections:
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Inches(11.0)  # Tamaño Carta a lo ancho
        section.page_height = Inches(8.5)  # Tamaño Carta a lo alto
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Título Principal
    p_titulo = doc.add_paragraph()
    p_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run_t1 = p_titulo.add_run("PLAN DE DESARROLLO CURRICULAR PARA EDUCACIÓN SECUNDARIA COMUNITARIA PRODUCTIVA\n")
    run_t1.bold = True
    run_t1.font.size = Pt(11)
    
    # Número de plan dinámico
    run_t2 = p_titulo.add_run(f"PLAN DE DESARROLLO CURRICULAR Nº {datos_formulario.get('nro_plan', '1')}")
    run_t2.bold = True
    run_t2.font.size = Pt(11)

    # Subtítulo 1. Datos referenciales
    p_datos = doc.add_paragraph()
    run_datos = p_datos.add_run("1.  Datos referenciales")
    run_datos.bold = True

    # --- TABLA DE DATOS REFERENCIALES ---
    t_datos = doc.add_table(rows=5, cols=4)
    t_datos.style = 'Table Grid'
    t_datos.autofit = False
    
    for row in t_datos.rows:
        row.cells[0].width = Inches(1.8)
        row.cells[1].width = Inches(3.2)
        row.cells[2].width = Inches(1.8)
        row.cells[3].width = Inches(3.2)

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
    t_datos.cell(3,1).text = "ARTES PLÁSTICAS Y VISUALES" # Puedes cambiar esto a un campo dinámico si quieres
    t_datos.cell(3,1).merge(t_datos.cell(3,3))

    set_cell_bold(t_datos.cell(4,0), "Trimestre")
    t_datos.cell(4,1).text = datos_formulario.get("trimestre", "").upper()
    t_datos.cell(4,1).merge(t_datos.cell(4,3))
    
    # Agregar fila extra para la FECHA (que ahora será con formato bonito)
    row_fecha = t_datos.add_row()
    set_cell_bold(row_fecha.cells[0], "FECHA")
    row_fecha.cells[1].text = datos_formulario.get("fechas", "").upper()
    row_fecha.cells[1].merge(row_fecha.cells[3])

    # Un párrafo limpio y vacío para separar tablas (evita que se peguen o formen recuadros raros)
    doc.add_paragraph("") 

    # Subtítulo 2. Desarrollo (Fuera del recuadro)
    p_des = doc.add_paragraph()
    run_des = p_des.add_run("2.  Desarrollo")
    run_des.bold = True

    # --- TABLA PRINCIPAL DEL PLAN ---
    semanas = plan_json.get("semanas", [])
    num_semanas = len(semanas)
    total_filas = 3 + num_semanas + 1
    t_plan = doc.add_table(rows=total_filas, cols=5)
    t_plan.style = 'Table Grid'

    # Fila 0 y 1: Objetivos
    cell_oh = t_plan.cell(0,0)
    p_oh = cell_oh.paragraphs[0]
    p_oh.add_run("Objetivo holístico de nivel\n").bold = True
    p_oh.add_run(plan_json.get("objetivo_holistico_nivel", ""))
    cell_oh.merge(t_plan.cell(0,4))

    cell_oa = t_plan.cell(1,0)
    p_oa = cell_oa.paragraphs[0]
    p_oa.add_run("Objetivo de aprendizaje\n").bold = True
    p_oa.add_run(plan_json.get("objetivo_aprendizaje", ""))
    cell_oa.merge(t_plan.cell(1,4))

    # Fila 2: Cabeceras con fondo celeste
    cabeceras = ["Contenidos", "Momentos del proceso formativo", "Recursos", "Períodos", "Criterios de evaluación"]
    for i, txt in enumerate(cabeceras):
        celda = t_plan.cell(2,i)
        set_cell_bg(celda, "D9E2F3") # Fondo celeste claro
        p = celda.paragraphs[0]
        p.add_run(txt).bold = True
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Filas de Semanas (Fila 3 en adelante)
    fila_actual = 3
    for idx, s in enumerate(semanas):
        p_cont = t_plan.cell(fila_actual, 0).paragraphs[0]
        p_cont.add_run(f"Semana {s.get('semana', idx+1)}\n").bold = True
        p_cont.add_run(s.get("contenidos", ""))
        
        t_plan.cell(fila_actual, 1).text = s.get("momentos", "")
        t_plan.cell(fila_actual, 2).text = s.get("recursos", "")
        t_plan.cell(fila_actual, 3).text = str(s.get("periodos", ""))
        
        # Aplicamos la viñeta con saltos de linea para los criterios
        t_plan.cell(fila_actual, 4).text = s.get("criterios_evaluacion", "").replace("\\n", "\n")
        
        fila_actual += 1

    # Última fila: Adaptaciones
    cell_adapt = t_plan.cell(fila_actual, 0)
    p_adapt = cell_adapt.paragraphs[0]
    p_adapt.add_run("Adaptaciones curriculares:\n").bold = True
    p_adapt.add_run(plan_json.get("adaptaciones_curriculares", ""))
    cell_adapt.merge(t_plan.cell(fila_actual, 4))

    # Dar formato (letra tamaño 9 para que todo entre bien en la tabla)
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
st.title("Generador de Planes Curriculares - Ixiamas")

MESES_ESPANOL = {1:"JUNIO", 2:"FEBRERO", 3:"MARZO", 4:"ABRIL", 5:"MAYO", 6:"JUNIO", 7:"JULIO", 8:"AGOSTO", 9:"SEPTIEMBRE", 10:"OCTUBRE", 11:"NOVIEMBRE", 12:"DICIEMBRE"}

with st.form("form_pdc"):
    st.subheader("Datos referenciales")
    c1, c2, c3 = st.columns(3)
    with c1:
        nro_plan = st.number_input("Nº de Plan Curricular", min_value=1, value=5, step=1)
        unidad = st.text_input("Unidad Educativa", value="IXIAMAS")
        nivel = st.selectbox("Nivel", options=["SECUNDARIA", "PRIMARIA"])
    with c2:
        docente = st.text_input("Maestra/o (Ej: PEPE PEREZ)")
        ano = st.text_input("Año de Escolaridad (Ej: Tercero C)")
        trimestre = st.selectbox("Trimestre", options=["PRIMERO", "SEGUNDO", "TERCERO"], index=1)
    with c3:
        tema = st.text_input("El Tema a avanzar", placeholder="Ej: La calidad de imagen...")
        semanas = st.selectbox("Semanas de duración", options=[1,2,3,4], index=3)
        st.write("Rango de Fechas")
        f1, f2 = st.columns(2)
        with f1:
            fecha_inicio = st.date_input("Desde")
        with f2:
            fecha_fin = st.date_input("Hasta")
            
    submitted = st.form_submit_button("Generar Plan en Word")

if submitted:
    if not tema or tema.strip() == "":
        st.error("Por favor ingrese el campo 'El Tema a avanzar'.")
    else:
        # Formatear la fecha como: DEL 1 DE JUNIO AL 26 DE JUNIO DE 2026
        mes_ini = MESES_ESPANOL.get(fecha_inicio.month, "ENERO")
        mes_fin = MESES_ESPANOL.get(fecha_fin.month, "ENERO")
        cadena_fechas = f"DEL {fecha_inicio.day} DE {mes_ini} AL {fecha_fin.day} DE {mes_fin} DE {fecha_fin.year}"

        datos_form = {
            "nro_plan": nro_plan,
            "unidad": unidad,
            "nivel": nivel,
            "ano": ano,
            "docente": docente,
            "trimestre": trimestre,
            "fechas": cadena_fechas,
            "tema": tema,
            "semanas": semanas
        }

        try:
            with st.spinner("Conectando con la IA ultra-rápida y estructurando tablas..."):
                client = get_groq_client()
                prompt = build_prompt(tema=tema, semanas=semanas, nivel=nivel)

                response = client.chat.completions.create(
                    messages=[
                        {
                            "role": "system",
                            "content": "Eres un asistente experto. Debes responder ÚNICAMENTE con formato JSON válido."
                        },
                        {
                            "role": "user",
                            "content": prompt
                        }
                    ],
                    model="llama-3.3-70b-versatile",
                    temperature=0.2,
                )

                text = response.choices[0].message.content
                plan_json = extract_json_from_text(text)
                docx_bio = generar_docx(datos_form, plan_json)

                st.success("¡Plan generado correctamente!")
                now_name = datetime.now().strftime("%Y%m%d_%H%M%S")
                filename = f"PDC_{unidad}_{tema[:10]}_{now_name}.docx"
                st.download_button(
                    label="📥 Descargar Plan_Curricular.docx",
                    data=docx_bio.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
            with st.expander("Detalles del error (depuración)"):
                st.exception(e)