import streamlit as st
import pandas as pd
import json
from io import BytesIO
import re

st.set_page_config(page_title="Resultados Examen de Admisión", layout="wide")

st.title("📊 Resultados del Examen de Admisión")
st.write("Sube tu archivo Excel con los datos de los estudiantes para procesar el informe de nivelación e ingresantes.")


def _to_number(value):
    """Convierte cadenas que representan porcentajes o números a float.
    Devuelve 0.0 si no se puede convertir.
    """
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        try:
            val = float(value)
            if 0 < val <= 1:
                val = val * 100
            return val
        except Exception:
            return 0.0
    if isinstance(value, str):
        s = value.strip()
        # Eliminar signo de porcentaje y espacios
        s = s.replace('%', '')
        s = s.replace(' ', '')
        # Normalizar separadores decimales: cambiar comas por puntos
        s = s.replace(',', '.')
        # Eliminar cualquier carácter no numérico (excepto punto y signo menos)
        s = re.sub(r'[^0-9.\-]', '', s)
        try:
            val = float(s) if s not in ['', '-', '.'] else 0.0
            if 0 < val <= 1:
                val = val * 100
            return val
        except Exception:
            return 0.0
    return 0.0

# Subida de archivo
uploaded_file = st.file_uploader("Elige un archivo Excel", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("Vista previa de los datos")
    st.dataframe(df.head(10))

    # Barra de progreso
    progress_bar = st.progress(0)
    st.write("Procesando los datos...")

    # Inicializamos listas para el resultado final
    resultados = []

    for i, row in df.iterrows():
        # ID incremental
        record_id = i + 1

        # Determinar asistió / condición
        asistio = "ASISTIÓ" if str(row.get("ASISTENCIA", "")).strip().upper() != "NO ASISTIÓ" else "NO ASISTIÓ"
        total_pct = _to_number(row.get("TOTAL %", 0))
        condicion = "INGRESÓ" if asistio == "ASISTIÓ" and total_pct >= 1 else "NO INGRESÓ"

        # Determinar áreas de nivelación
        areas_nivelacion = []

        if asistio == "ASISTIÓ":
            if _to_number(row.get("COMUNICACIÓN %", 0)) <= 30:
                areas_nivelacion.append({"curso": "COMUNICACIÓN"})
            if _to_number(row.get("HABILIDADES COMUNICATIVAS %", 0)) <= 30:
                areas_nivelacion.append({"curso": "HABILIDADES COMUNICATIVAS"})
            if _to_number(row.get("MATEMÁTICA %", 0)) <= 30:
                areas_nivelacion.append({"curso": "MATEMATICA"})
            if _to_number(row.get("CIENCIA, TECNOLOGÍA Y AMBIENTE %", 0)) <= 30:
                # Dependiendo de la carrera
                if row.get("CARRERA", "").upper() in ["DERECHO", "CONTABILIDAD", "ADMINISTRACIÓN DE EMPRESAS"]:
                    areas_nivelacion.append({"curso": "CIENCIAS SOCIALES"})
                else:
                    areas_nivelacion.append({"curso": "CIENCIA, TECNOLOGÍA Y AMBIENTE"})

        requiere_nivelacion = "SI" if len(areas_nivelacion) > 0 else "NO"

        # Agregar fila al resultado final
        resultados.append({
            "id": record_id,
            "periodo": "2026-1",
            "codigo_estudiante": row.get("CODIGO DE ESTUDIANTE", ""),
            "apellidos": row.get("APELLIDOS", ""),
            "nombres": row.get("NOMBRES", ""),
            "dni": row.get("DNI", ""),
            "area": row.get("AREA", ""),
            "programa": row.get("CARRERA", ""),
            "local_examen": row.get("SEDE DE ESTUDIO", ""),
            "MODALIDAD": row.get("MODALIDAD", ""),
            "puntaje": row.get("TOTAL", 0),
            "asistio": asistio,
            "condicion": condicion,
            "requiere_nivelacion": requiere_nivelacion,
            "areas_nivelacion": json.dumps(areas_nivelacion, ensure_ascii=False),
            "fecha_registro": pd.to_datetime(row.get("FECHA DE EXAMEN")).strftime("%Y-%m-%d 00:00:00") if row.get("FECHA DE EXAMEN") else "",
            "estado": 1
        })

        # Actualizar barra de progreso
        progress_bar.progress((i + 1) / len(df))

    # Convertimos a DataFrame
    df_resultados = pd.DataFrame(resultados)

    st.subheader("Vista previa de resultados")
    st.dataframe(df_resultados.head(10))

    # Descargar resultado como Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_resultados.to_excel(writer, index=False, sheet_name="Resultados")
    processed_data = output.getvalue()

    st.download_button(
        label="📥 Descargar Excel de resultados",
        data=processed_data,
        file_name="resultados_examen_admision.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success("✅ Procesamiento completado.")