import streamlit as st
import json
import re
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Intentamos importar la librería oficial de Gemini / Generative AI
try:
    import google.generativeai as genai
except Exception:
    genai = None

st.set_page_config(page_title="Generador de Planes Curriculares - Ixiamas", layout="wide")

# --- Helpers ---
def configure_genai():
    # Busca claves en st.secrets (flexible)
    api_key = None
    if "GOOGLE_API_KEY" in st.secrets:
        api_key = st.secrets["GOOGLE_API_KEY"]
    elif "API_KEY" in st.secrets:
        api_key = st.secrets["API_KEY"]
    elif "google_api_key" in st.secrets:
        api_key = st.secrets["google_api_key"]
    if not api_key:
        raise RuntimeError("Falta la clave de API en `st.secrets` (ej. 'GOOGLE_API_KEY').")
    if genai is None:
        raise RuntimeError("No se pudo importar `google.generativeai`. Instale la dependencia.")
    genai.configure(api_key=api_key)
    return genai

def extract_text_from_response(resp):
    """
    Intenta extraer la respuesta textual de distintos formatos que puede devolver la SDK.
    """
    # resp puede ser un objeto con .text, .candidates, .last or dict
    try:
        if isinstance(resp, dict):
            # formes comunes
            if "candidates" in resp and isinstance(resp["candidates"], list) and len(resp["candidates"])>0:
                cand = resp["candidates"][0]
                if isinstance(cand, dict) and "content" in cand:
                    return cand["content"]
                if isinstance(cand, dict) and "text" in cand:
                    return cand["text"]
            if "content" in resp:
                return resp["content"]
            if "text" in resp:
                return resp["text"]
            # fallback: string representation
            return json.dumps(resp)
        # objeto con .text
        if hasattr(resp, "text"):
            return resp.text
        # objeto con .candidates
        if hasattr(resp, "candidates"):
            c = resp.candidates
            if isinstance(c, (list, tuple)) and len(c) > 0:
                first = c[0]
                if isinstance(first, dict) and "content" in first:
                    return first["content"]
                if hasattr(first, "content"):
                    return first.content
        # string
        return str(resp)
    except Exception:
        return str(resp)

def extract_json_from_text(text):
    # Buscar el primer bloque JSON válido en el texto
    # Primero intentar carga directa
    try:
        return json.loads(text)
    except Exception:
        pass
    # Extraer substring que parezca JSON (objeto o array)
    # Buscar desde la primera { hasta el último } que cierre correctamente
    # método simple: regex para {...}
    regex_obj = re.compile(r"(\{(?:.|\n)*\})", re.DOTALL)
    regex_arr = re.compile(r"(\[(?:.|\n)*\])", re.DOTALL)
    for regex in (regex_obj, regex_arr):
        m = regex.search(text)
        if m:
            candidate = m.group(1)
            # intentar balanceo de llaves si hay múltiples matches
            # intentar cargar
            try:
                return json.loads(candidate)
            except Exception:
                # intentar reparar comillas simples -> dobles (riesgoso, último recurso)
                cand2 = candidate.replace("'", '"')
                try:
                    return json.loads(cand2)
                except Exception:
                    continue
    raise ValueError("No se pudo extraer JSON válido de la respuesta de IA.")

def build_prompt(tema, semanas, nivel, contexto_extra=""):
    weeks = int(semanas)
    prompt = f"""
Actúa como un educador boliviano experto en diseño curricular. Te daré un tema y un número de semanas. Devuelve SOLO un JSON bien formado (sin texto adicional) con esta estructura exacta:

{{
  "tema": "<tema>",
  "nivel": "<nivel>",
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
    }},
    ...
  ]
}}

Genera exactamente {weeks} objetos dentro del arreglo "semanas" numerados del 1 al {weeks}. 
Cada campo debe ser texto simple (no arrays excepto el arreglo "semanas") y describir apropiadamente actividades y criterios para el contexto educativo boliviano.
Tema: \"{tema}\"
Nivel: \"{nivel}\"
{contexto_extra}

Responde únicamente con el JSON (asegúrate que sea parseable por json.loads en Python).
"""
    return prompt

def generar_docx(datos_formulario, plan_json):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)

    # Título
    h = doc.add_heading("Plan de Desarrollo Curricular (PDC)", level=1)
    h.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Info de cabecera (tabla)
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

    doc.add_paragraph("")  # espacio

    # Objetivos
    doc.add_heading("Objetivos", level=2)
    obj_hol = plan_json.get("objetivo_holistico", "")
    obj_apr = plan_json.get("objetivo_aprendizaje", "")
    p1 = doc.add_paragraph()
    p1.add_run("Objetivo Holístico: ").bold = True
    p1.add_run(obj_hol)
    p2 = doc.add_paragraph()
    p2.add_run("Objetivo de Aprendizaje: ").bold = True
    p2.add_run(obj_apr)

    doc.add_paragraph("")  # espacio

    # Tabla grande con 5 columnas
    doc.add_heading("Programación por Semana", level=2)
    cols = ["Contenidos", "Momentos del proceso formativo", "Recursos", "Períodos", "Criterios de evaluación"]
    big_table = doc.add_table(rows=1, cols=5)
    big_table.style = 'Table Grid'
    hdr_cells = big_table.rows[0].cells
    for i, c in enumerate(cols):
        hdr_cells[i].text = c

    # Agregar filas por cada semana
    semanas = plan_json.get("semanas", [])
    if not isinstance(semanas, list):
        raise ValueError("El campo 'semanas' del JSON no es una lista.")

    for s in semanas:
        cont = s.get("contenidos", "")
        momentos = s.get("momentos", {})
        # Formatear momentos en un solo string
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

    # Ajustes menores de estilo (opcional): fuente de todas las celdas
    for table in (table, big_table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Arial'
                        run.font.size = Pt(10)

    # Guardar en BytesIO
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- Interfaz Streamlit ---
st.title("Generador de Planes Curriculares - Ixiamas")

with st.form("form_pdc"):
    # Usamos columnas para un layout limpio
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
    # Validaciones básicas
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

        # Llamada a la IA y generación del documento
        try:
            with st.spinner("Conectando con la API de IA y generando el plan..."):
                # Configuración
                gi = configure_genai()

                prompt = build_prompt(tema=tema, semanas=semanas, nivel=nivel)

                # Llamada correcta a Gemini
                # Buscar automáticamente el nombre correcto del modelo en tu cuenta
                # Primero intentamos modelos preferidos (más nuevos)
                modelos_preferidos = [
                    "gemini-2.0-flash",
                    "gemini-1.5-pro", 
                    "gemini-1.5-flash",
                    "gemini-1.5-pro-latest",
                    "gemini-1.5-flash-latest"
                ]
                
                modelo_ideal = None
                # Primero intentar con los modelos preferidos
                for pref in modelos_preferidos:
                    try:
                        # Verificar si el modelo existe en la lista de modelos disponibles
                        for m in gi.list_models():
                            if pref in m.name.lower() and 'generateContent' in m.supported_generation_methods:
                                modelo_ideal = m.name
                                break
                        if modelo_ideal:
                            break
                    except Exception:
                        continue
                
                # Si no encontramos modelo preferido, buscar cualquier gemini disponible
                if not modelo_ideal:
                    for m in gi.list_models():
                        if 'generateContent' in m.supported_generation_methods:
                            if 'gemini' in m.name.lower():
                                # Evitar modelos antiguos descontinuados
                                if 'gemini-2.5-flash' not in m.name.lower():
                                    modelo_ideal = m.name
                                    break
                            
                if not modelo_ideal:
                    raise RuntimeError("Google no devolvió ningún modelo compatible. Revisa tu cuenta. Modelos disponibles: " + 
                                     ", ".join([m.name for m in gi.list_models()]))
                
                # Ahora sí, generamos el contenido con el modelo moderno que encontró
                model = gi.GenerativeModel(modelo_ideal)
                response = model.generate_content(prompt)
                
                text = extract_text_from_response(response)

                # Intentar parsear JSON
                plan_json = extract_json_from_text(text)

                # Validaciones del JSON
                if "semanas" not in plan_json or not isinstance(plan_json["semanas"], list):
                    st.error("La respuesta de la IA no contiene el campo 'semanas' como lista. Mostrar respuesta cruda para depuración.")
                    st.code(text)
                else:
                    # Generar el docx
                    docx_bio = generar_docx(datos_form, plan_json)

                    # Mostrar resumen y ofrecer descarga
                    st.success("Plan generado correctamente.")
                    now_name = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"Plan_Curricular_{now_name}.docx"
                    st.download_button(
                        label="Descargar Plan_Curricular.docx",
                        data=docx_bio.getvalue(),
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

        except Exception as e:
            st.error(f"Ocurrió un error: {e}")
            # Si hay más info, mostrarla en expander para depuración
            with st.expander("Detalles del error (depuración)"):
                st.exception(e)