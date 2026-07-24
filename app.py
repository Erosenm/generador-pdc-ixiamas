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
Actúa como un educador boliviano experto en diseño curricular. Te daré un tema y un número de semanas. Devuelve SOLO un JSON bien formado con esta estructura exacta:

{{
  "tema": "{tema}",
  "nivel": "{nivel}",
  "objetivo_holistico": "<texto>",
  "objetivo_aprendizaje": "<texto>",
  "semanas": [
    {{
      "semana": 1,
      "contenidos": "<texto>",
      "momentos": {{
        "practica": "<texto>",
        "teoria": "<texto>",
        "valoracion": "<texto>",
        "produccion": "<texto>"
      }},
      "recursos": "<texto>",
      "periodos": "<numero o texto corto>",
      "criterios_evaluacion": "<texto corto>"
    }}
  ]
}}

Genera exactamente {weeks} objetos dentro del arreglo "semanas" numerados del 1 al {weeks}. Los recursos deben usar elementos del contexto natural de Ixiamas.
"""
    return prompt

def generar_docx(datos_formulario, plan_json):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    h = doc.add_heading("Plan de Desarrollo Curricular (PDC)", level=1)
    h.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    def add_row(name, value):
        row = table.add_row()
        row.cells[0].text = name
        row.cells[1].text = value

    add_row("Unidad Educativa", datos_formulario.get("unidad", ""))
    add_row("Nivel", datos_formulario.get("nivel", ""))
    add_row("Año de escolaridad", datos_formulario.get("ano", ""))
    add_row("Docente", datos_formulario.get("docente", ""))
    add_row("Trimestre", datos_formulario.get("trimestre", ""))
    add_row("Fechas", datos_formulario.get("fechas", ""))
    add_row("Tema a avanzar", datos_formulario.get("tema", ""))
    add_row("Semanas de duración", str(datos_formulario.get("semanas", "")))
    add_row("Generado el", datetime.now().strftime("%Y-%m-%d %H:%M"))

    doc.add_paragraph("")

    doc.add_heading("Objetivos", level=2)
    obj_hol = plan_json.get("objetivo_holistico", "")
    obj_apr = plan_json.get("objetivo_aprendizaje", "")
    p1 = doc.add_paragraph()
    p1.add_run("Objetivo Holístico: ").bold = True
    p1.add_run(obj_hol)
    p2 = doc.add_paragraph()
    p2.add_run("Objetivo de Aprendizaje: ").bold = True
    p2.add_run(obj_apr)

    doc.add_paragraph("")

    doc.add_heading("Programación por Semana", level=2)
    cols = ["Contenidos", "Momentos del proceso formativo", "Recursos", "Períodos", "Criterios de evaluación"]
    big_table = doc.add_table(rows=1, cols=5)
    big_table.style = 'Table Grid'
    hdr_cells = big_table.rows[0].cells
    for i, c in enumerate(cols):
        hdr_cells[i].text = c

    semanas = plan_json.get("semanas", [])
    if not isinstance(semanas, list):
        raise ValueError("El campo 'semanas' del JSON no es una lista.")

    for s in semanas:
        cont = s.get("contenidos", "")
        momentos = s.get("momentos", {})
        if isinstance(momentos, dict):
            momentos_str_lines = []
            for k in ("practica", "teoria", "valoracion", "produccion"):
                v = momentos.get(k, "")
                if v:
                    momentos_str_lines.append(f"{k.capitalize()}: {v}")
            momentos_str = "\n".join(momentos_str_lines)
        else:
            momentos_str = str(momentos)

        recursos = s.get("recursos", "")
        periodos = str(s.get("periodos", ""))
        criterios = s.get("criterios_evaluacion", "")

        row_cells = big_table.add_row().cells
        row_cells[0].text = cont
        row_cells[1].text = momentos_str
        row_cells[2].text = recursos
        row_cells[3].text = periodos
        row_cells[4].text = criterios

    for table_ref in (table, big_table):
        for row in table_ref.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(10)

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
                    model="llama3-70b-8192",
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