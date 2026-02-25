import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io

st.set_page_config(layout="wide")
st.title("Sistema de Horarios Académicos - Elemaitre 2026")

archivo = st.file_uploader("Cargar archivo Excel", type=["xlsx"])

# =====================================================
# FUNCIÓN PARA GENERAR FRANJAS 50 MIN + 10 DESCANSO
# =====================================================

def generar_franjas(inicio="07:00", fin="22:00"):
    franjas = []
    hora_actual = datetime.strptime(inicio, "%H:%M")
    hora_fin_total = datetime.strptime(fin, "%H:%M")

    while hora_actual < hora_fin_total:
        hora_fin_clase = hora_actual + timedelta(minutes=50)
        franjas.append(f"{hora_actual.strftime('%H:%M')} - {hora_fin_clase.strftime('%H:%M')}")
        hora_actual += timedelta(minutes=60)

    return franjas


# =====================================================
# SI SE CARGA ARCHIVO
# =====================================================

if archivo is not None:

    # =========================
    # LEER ARCHIVO
    # =========================
    df = pd.read_excel(archivo, header=5)
    df.columns = df.columns.str.strip().str.lower()

    # =========================
    # SEPARAR HORARIO
    # =========================
    df[['dia', 'horas']] = df['horario'].str.split(' ', n=1, expand=True)
    df[['hora_inicio', 'hora_fin']] = df['horas'].str.split('-', expand=True)

    df['dia'] = df['dia'].str.strip().str.lower()
    df['hora_inicio'] = df['hora_inicio'].str.strip()
    df['hora_fin'] = df['hora_fin'].str.strip()

    # =========================
    # FILTRO DOCENTE
    # =========================
    docentes = df["docente"].dropna().unique()
    docente_sel = st.selectbox("Seleccione Docente", sorted(docentes))

    df_filtrado = df[df["docente"] == docente_sel]

    # =========================
    # CREAR MATRIZ
    # =========================
    franjas = generar_franjas()
    dias = ["lunes", "martes", "miercoles", "jueves", "viernes"]
    horario_matriz = pd.DataFrame("", index=franjas, columns=dias)

    # =========================
    # LLENAR CURSOS
    # =========================
    for _, fila in df_filtrado.iterrows():

        dia = fila["dia"]
        inicio = fila["hora_inicio"]
        fin = fila["hora_fin"]

        try:
            hora_inicio = datetime.strptime(inicio, "%H:%M")
            hora_fin = datetime.strptime(fin, "%H:%M")

            while hora_inicio < hora_fin:

                hora_fin_clase = hora_inicio + timedelta(minutes=50)
                franja_str = f"{hora_inicio.strftime('%H:%M')} - {hora_fin_clase.strftime('%H:%M')}"

                if franja_str in horario_matriz.index and dia in horario_matriz.columns:
                    horario_matriz.loc[franja_str, dia] = fila["curso"]

                hora_inicio += timedelta(minutes=60)

        except:
            pass

    # =========================
    # CONSEJO LUNES 13:30-16:30
    # =========================
    for franja in horario_matriz.index:
        hora_inicio_franja = franja.split(" - ")[0]
        if hora_inicio_franja in ["13:30", "14:30", "15:30"]:
            horario_matriz.loc[franja, "lunes"] = "CONSEJO"

    # =========================
    # ALMUERZO 12:00
    # =========================
    for franja in horario_matriz.index:
        if franja.startswith("12:00"):
            horario_matriz.loc[franja, :] = "ALMUERZO"

    st.subheader("Vista Previa")
    st.dataframe(horario_matriz)

    # =====================================================
    # EXPORTAR EXCEL
    # =====================================================

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:

        workbook = writer.book
        worksheet = workbook.add_worksheet("Horario")

        # FORMATOS
        formato_titulo = workbook.add_format({
            'bold': True,
            'align': 'center',
            'font_name': 'Georgia',
            'font_size': 24
        })

        formato_info = workbook.add_format({
            'align': 'center',
            'font_size': 14
        })

        formato_header = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center'
        })

        formato_curso = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True
        })

        formato_consejo = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#BDD7EE'
        })

        formato_almuerzo = workbook.add_format({
            'bold': True,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D9D9D9'
        })

        # DATOS DOCENTE
        datos_docente = df_filtrado.iloc[0]
        nombre_docente = datos_docente["docente"]
        correo = datos_docente.get("correo", "")
        telefono = datos_docente.get("telefono", "")

        # ENCABEZADO SUPERIOR
        worksheet.merge_range("A1:F1", f"{nombre_docente}", formato_titulo)
        worksheet.merge_range("A2:F2", f"HORARIO ACADÉMICO 2026", formato_info)
        worksheet.merge_range("A3:F3", f"Correo: {correo}", formato_info)
        worksheet.merge_range("A4:F4", f"Teléfono: {telefono}", formato_info)

        fila_inicio = 5

        worksheet.write_row(
            fila_inicio, 0,
            ["Hora", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes"],
            formato_header
        )

        horas_lista = list(horario_matriz.index)

        # ESCRIBIR HORAS
        for i, hora in enumerate(horas_lista):
            worksheet.write(fila_inicio + 1 + i, 0, hora)

        # FUSIÓN VERTICAL
        for col_idx, dia in enumerate(horario_matriz.columns):

            fila = 0

            while fila < len(horas_lista):

                valor = horario_matriz.iloc[fila, col_idx]

                if valor == "" or valor == "ALMUERZO":
                    fila += 1
                    continue

                inicio = fila
                fin = fila

                while (
                    fin + 1 < len(horas_lista)
                    and horario_matriz.iloc[fin + 1, col_idx] == valor
                ):
                    fin += 1

                formato_usar = formato_consejo if valor == "CONSEJO" else formato_curso

                worksheet.merge_range(
                    fila_inicio + 1 + inicio,
                    col_idx + 1,
                    fila_inicio + 1 + fin,
                    col_idx + 1,
                    valor,
                    formato_usar
                )

                fila = fin + 1

        # FUSIÓN HORIZONTAL ALMUERZO
        for i, hora in enumerate(horas_lista):

            fila_valores = horario_matriz.loc[hora]

            if all(valor == "ALMUERZO" for valor in fila_valores):

                worksheet.merge_range(
                    fila_inicio + 1 + i, 0,
                    fila_inicio + 1 + i, 5,
                    "ALMUERZO",
                    formato_almuerzo
                )

        worksheet.set_column("A:A", 15)
        worksheet.set_column("B:F", 20)

    excel_data = output.getvalue()

    st.download_button(
        "Descargar Horario Institucional",
        data=excel_data,
        file_name="horario_institucional.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )