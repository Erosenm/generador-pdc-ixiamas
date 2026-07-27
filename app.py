import streamlit as st
import json
import re
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Intentamos importar Groq
try:
    from groq import Groq
except Exception:
    Groq = None

st.set_page_config(page_title="Generador de Planes Curriculares - Ixiamas", layout="wide")

# --- Helpers ---
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

def build_prompt(tema, semanas, nivel, contexto_extra=""):
    weeks = int(semanas)
    prompt = f"""
Actúa como un pedagogo boliviano experto en diseño curricular del Ministerio de Educación (Ley 070 Avelino Siñani - Elizardo Pérez). 
Tu tarea es generar un Plan de Desarrollo Curricular (PDC) altamente detallado, formal y riguroso.

Tema a desarrollar: "{tema}"
Nivel educativo: "{nivel}"
Duración: {weeks} semanas.
Contexto socio-comunitario: Municipio de Ixiamas, Norte de La Paz, Amazonía boliviana (considerar flora, fauna, clima, vocación productiva maderera y agrícola, e identidad cultural).

Devuelve ÚNICAMENTE un objeto JSON válido con la siguiente estructura exacta. No agregues texto antes ni después del JSON.

{{
  "objetivo_holistico_nivel": "Redacta el objetivo general del nivel según la Ley 070 (Ser, Saber, Hacer, Decidir), mencionando valores sociocomunitarios, Madre Tierra y descolonización.",
  "objetivo_aprendizaje": "Redacta el objetivo específico para este tema, integrando las 4 dimensiones (Ser, Saber, Hacer, Decidir).",
  "semanas": [
    {{
      "semana": 1,
      "contenidos": "Detalla el subtítulo o contenido específico a avanzar esta semana.",
      "momentos": "Redacta en un solo bloque de texto continuo y detallado (no en diccionario) las actividades de Práctica, Teoría, Valoración y Producción. Empieza con 'Lectura: 10 minutos de lectura...' si corresponde. Ejemplo: 'Práctica: Observación del entorno... Teoría: Análisis de los conceptos... Valoración: Reflexión sobre la importancia... Producción: Elaboración de un mapa...'",
      "recursos": "Lista de materiales (educativos, analógicos, de la vida, para la producción). Incluye recursos específicos del contexto de Ixiamas (ej. hojas, semillas, madera). Usa viñetas o saltos de línea (\\n).",
      "periodos": "2",
      "criterios_evaluacion": "Redacta los criterios divididos por dimensiones. Usa saltos de línea (\\n). Ejemplo: 'SER: \\n- Valora...\\nSABER: \\n- Identifica...\\nHACER: \\n- Elabora...\\nDECIDIR: \\n- Promueve...'"
    }}
  ],
  "adaptaciones_curriculares": "Redacta un párrafo formal sobre cómo se adaptará el contenido para estudiantes con dificultades de aprendizaje, priorizando el uso de material del contexto de Ixiamas."
}}

Instrucciones críticas:
1. Genera exactamente {weeks} objetos dentro del arreglo "semanas".
2. La redacción de los "momentos" debe ser extensa y pedagógica, como un plan real.
3. En la semana 1, los "recursos" y "criterios_evaluacion" deben ser MUY extensos y detallados (como para abarcar todo el plan), en las siguientes semanas puedes dejarlos vacíos ("") o poner "Ídem" si el plan lo requiere para mantener la estética de tabla limpia, pero te sugiero llenarlos si aporta valor.
4. Usa saltos de línea `\\n` en los strings para que en Word se formen párrafos distintos.
"""
    return prompt

from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- Función auxiliar para el grosor de los bordes de tabla ---
def set_cell_border(cell, **kwargs):
    """
    Establece bordes en una celda de docx.
    Uso: set_cell_border(cell, top={"sz": 12, "val": "single", "color": "000000"})
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ["sz", "val", "color", "space", "shadow"]:
                if key in edge_data:
                    element.set(qn('w:{}'.format(key)), str(edge_data[key]))

def generar_docx(datos_formulario, plan_json):
    doc = Document()
    
    # Configuración de estilos generales
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)
    
    # Cambiar márgenes para que quepa la tabla ancha (ej: Margen estrecho)
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # Título Principal
    p = doc.add_paragraph()
    p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = p.add_run("PLAN DE DESARROLLO CURRICULAR PARA EDUCACIÓN SECUNDARIA COMUNITARIA PRODUCTIVA\nPLAN DE DESARROLLO CURRICULAR Nº 5")
    run.bold = True
    run.font.size = Pt(12)

    doc.add_paragraph("Datos referenciales").bold = True

    # --- TABLA DE DATOS REFERENCIALES ---
    # Creamos una tabla de 5 filas y 4 columnas
    t_datos = doc.add_table(rows=5, cols=4)
    t_datos.style = 'Table Grid'
    t_datos.autofit = False

    # Ajustar anchos (aproximados)
    for row in t_datos.rows:
        row.cells[0].width = Inches(1.5)
        row.cells[1].width = Inches(2.5)
        row.cells[2].width = Inches(1.5)
        row.cells[3].width = Inches(2.5)

    # Fila 0
    t_datos.cell(0,0).text = "Distrito educativo"
    t_datos.cell(0,1).text = "IXIAMAS"
    t_datos.cell(0,2).text = "Unidad Educativa"
    t_datos.cell(0,3).text = datos_formulario.get("unidad", "").upper()

    # Fila 1
    t_datos.cell(1,0).text = "Nivel"
    t_datos.cell(1,1).text = datos_formulario.get("nivel", "").upper()
    t_datos.cell(1,2).text = "Año de escolaridad/Paralelo"
    t_datos.cell(1,3).text = datos_formulario.get("ano", "").upper()

    # Fila 2
    t_datos.cell(2,0).text = "Maestra/o"
    t_datos.cell(2,1).text = datos_formulario.get("docente", "").upper()
    # Fusionar celdas restantes de la fila 2
    t_datos.cell(2,1).merge(t_datos.cell(2,3))

    # Fila 3
    t_datos.cell(3,0).text = "Área"
    t_datos.cell(3,1).text = "ARTES PLÁSTICAS Y VISUALES" # Puedes cambiarlo o hacerlo dinámico
    t_datos.cell(3,1).merge(t_datos.cell(3,3))

    # Fila 4
    t_datos.cell(4,0).text = "Trimestre"
    t_datos.cell(4,1).text = datos_formulario.get("trimestre", "").upper()
    t_datos.cell(4,1).merge(t_datos.cell(4,3))

    doc.add_paragraph("") # Espacio
    doc.add_paragraph("Desarrollo").bold = True

    # --- TABLA PRINCIPAL DEL PLAN ---
    semanas = plan_json.get("semanas", [])
    num_semanas = len(semanas)
    
    # Calculamos filas: 3 estáticas (Obj Holístico, Obj Aprendizaje, Cabeceras) + 1 por cada semana + 1 Adaptaciones
    total_filas = 3 + num_semanas + 1
    t_plan = doc.add_table(rows=total_filas, cols=5)
    t_plan.style = 'Table Grid'

    # Fila 0: Objetivo Holístico
    cell_oh_title = t_plan.cell(0,0)
    cell_oh_title.text = "Objetivo holístico de nivel\n" + plan_json.get("objetivo_holistico_nivel", "")
    cell_oh_title.merge(t_plan.cell(0,4))

    # Fila 1: Objetivo de aprendizaje
    cell_oa_title = t_plan.cell(1,0)
    cell_oa_title.text = "Objetivo de aprendizaje\n" + plan_json.get("objetivo_aprendizaje", "")
    cell_oa_title.merge(t_plan.cell(1,4))

    # Fila 2: Cabeceras de columnas
    cabeceras = ["Contenidos", "Momentos del proceso formativo", "Recursos", "Períodos", "Criterios de evaluación"]
    for i, txt in enumerate(cabeceras):
        p = t_plan.cell(2,i).paragraphs[0]
        p.add_run(txt).bold = True
        p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Filas de Semanas (Fila 3 en adelante)
    fila_actual = 3
    for idx, s in enumerate(semanas):
        titulo_semana = f"Semana {s.get('semana', idx+1)}\n"
        
        # Col 0: Contenidos
        t_plan.cell(fila_actual, 0).text = titulo_semana + s.get("contenidos", "")
        # Col 1: Momentos
        t_plan.cell(fila_actual, 1).text = s.get("momentos", "")
        # Col 2: Recursos
        t_plan.cell(fila_actual, 2).text = s.get("recursos", "")
        # Col 3: Periodos
        t_plan.cell(fila_actual, 3).text = str(s.get("periodos", ""))
        # Col 4: Criterios
        t_plan.cell(fila_actual, 4).text = s.get("criterios_evaluacion", "")
        
        fila_actual += 1

    # Última fila: Adaptaciones curriculares
    cell_adapt = t_plan.cell(fila_actual, 0)
    cell_adapt.text = "Adaptaciones curriculares:\n" + plan_json.get("adaptaciones_curriculares", "")
    cell_adapt.merge(t_plan.cell(fila_actual, 4))

    # Dar formato limpio a toda la tabla principal
    for row in t_plan.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(9) # Letra más pequeña para que quepa bien el texto

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio
# --- Interfaz Streamlit ---
st.title("Generador de Planes Curriculares - Ixiamas")

with st.form("form_pdc"):
    c1, c2, c3 = st.columns([3,3,2])
    with c1:
        unidad = st.text_input("Unidad Educativa", value="IXIAMAS")
        nivel = st.selectbox("Nivel", options=["PRIMARIA", "SECUNDARIA"])
        ano = st.text_input("Año de Escolaridad", placeholder="Ej: Tercero C")
    with c2:
        docente = st.text_input("Maestra/o")
        trimestre = st.selectbox("Trimestre", options=["PRIMERO", "SEGUNDO", "TERCERO"])
        fechas = st.text_input("Fechas (ej: Del 1 al 26 de Junio de 2026)")
    with c3:
        tema = st.text_input("El Tema a avanzar", placeholder="Ej: Robótica básica")
        semanas = st.selectbox("Semanas de duración", options=[1,2,3,4], index=0)
    submitted = st.form_submit_button("Generar Plan en Word")

if submitted:
    if not tema or tema.strip() == "":
        st.error("Por favor ingrese el campo 'El Tema a avanzar'.")
    else:
        datos_form = {
            "unidad": unidad,
            "nivel": nivel,
            "ano": ano,
            "docente": docente,
            "trimestre": trimestre,
            "fechas": fechas,
            "tema": tema,
            "semanas": semanas
        }

        try:
            with st.spinner("Conectando con la IA ultra-rápida y generando el plan..."):
                client = get_groq_client()
                prompt = build_prompt(tema=tema, semanas=semanas, nivel=nivel)

                # Llamada directa y veloz a la API de Groq
                response = client.chat.completions.create(
                    messages=[
                        {
                            "role": "system",
                            "content": "Eres un asistente experto. Debes responder ÚNICAMENTE con formato JSON válido, sin usar bloques de código markdown al inicio o al final."
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
                filename = f"Plan_Curricular_{now_name}.docx"
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