# ============================
# BLOQUE 1 / 4
# IMPORTS + CONFIG + ALGORITMO
# ============================
import flet as ft
import pandas as pd
import re
import os
import sys
from datetime import datetime

# ---- PDF / ReportLab ----
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import ParagraphStyle


# =========================================================
# CONFIGURACIÓN RUTAS / DATASET / SALIDA
# =========================================================
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(__file__)

DATASET_FILE = "dataset.xlsx"
OUTPUT_DIR = BASE_DIR


# =========================================================
# ⭐ ALGORITMO CORREGIDO (SIEMPRE A FAVOR, PERO SIN PASARSE)
# =========================================================
def _subset_best_between(df: pd.DataFrame, min_needed: int, max_allowed: int):
    """
    Devuelve (indices, suma) de una combinación que:
    - suma <= max_allowed
    - si existe suma >= min_needed: elige la MENOR suma (mínimo exceso)
    - si no existe: elige la MAYOR suma < min_needed (lo más cercano por debajo)

    Importante:
    - Si df está ordenado por CR desc, este DP tenderá a elegir cursos grandes primero en empates.
    """
    if df.empty or max_allowed <= 0:
        return [], 0

    df2 = df.copy()
    df2["CR"] = pd.to_numeric(df2["CR"], errors="coerce").fillna(0).astype(int)

    items = []
    for idx, row in df2.iterrows():
        cr = int(row.get("CR", 0))
        if cr > 0:
            items.append((idx, cr))

    dp = {0: []}  # suma -> indices

    for idx, cr in items:
        for s in sorted(list(dp.keys()), reverse=True):
            ns = s + cr
            if ns <= max_allowed and ns not in dp:
                dp[ns] = dp[s] + [idx]

    sums = sorted(dp.keys())

    # 1) menor suma que cumpla el mínimo
    for s in sums:
        if s >= max(0, min_needed) and s <= max_allowed:
            return dp[s], s

    # 2) si no se puede, tomar la mayor posible por debajo
    best = max(sums) if sums else 0
    return dp.get(best, []), best


def seleccionar_convalidacion(df_conva: pd.DataFrame, crd: float, tolerancia: int = 2):
    """
    Selección SECUENCIAL por ciclo:
    - Primero ciclo 1, luego ciclo 2, luego 3, luego 4...
    - Solo pasas al siguiente ciclo si NO alcanzas CRD con lo acumulado.
    - Límite máximo total: CRD + tolerancia (ej. +2).
    - Dentro de cada ciclo: CR orden desc.
    """
    if df_conva.empty:
        return [], 0

    crd_int = int(float(crd))
    limite_total = crd_int + int(tolerancia)

    df = df_conva.copy()
    df["CR"] = pd.to_numeric(df["CR"], errors="coerce").fillna(0).astype(int)
    df["CICLO_NUM"] = (
        df["CICLO"].astype(str).str.extract(r"(\d+)")[0].fillna(0).astype(int)
    )

    ciclos_disponibles = sorted([c for c in df["CICLO_NUM"].unique().tolist() if c > 0])

    seleccion_total = []
    suma_total = 0

    for ciclo in ciclos_disponibles:
        # Si ya llegamos al CRD, NO seguir
        if suma_total >= crd_int:
            break

        max_restante = limite_total - suma_total
        if max_restante <= 0:
            break

        min_restante = crd_int - suma_total

        df_ciclo = df[df["CICLO_NUM"] == ciclo].copy()
        df_ciclo = df_ciclo.sort_values(by="CR", ascending=False)

        sel_c, suma_c = _subset_best_between(
            df=df_ciclo,
            min_needed=min_restante,
            max_allowed=max_restante
        )

        if suma_c <= 0:
            continue

        seleccion_total.extend(sel_c)
        suma_total += suma_c

        if suma_total >= crd_int:
            break

    return seleccion_total, suma_total


# =========================================================
# FUNCIÓN: CARGAR DATASET
# =========================================================
def cargar_dataset():
    ruta = os.path.join(BASE_DIR, DATASET_FILE)
    if not os.path.exists(ruta):
        raise FileNotFoundError(f"No se encontró el archivo: {ruta}")

    df = pd.read_excel(ruta, header=0)
    df.columns = df.columns.str.strip().str.upper()

    columnas_necesarias = {"CARRERA", "UNID. NEGOCIO", "CICLO", "CR", "CURSO"}
    faltantes = columnas_necesarias - set(df.columns)
    if faltantes:
        raise ValueError(f"Faltan columnas necesarias en el dataset: {faltantes}")

    if "MATERIA" not in df.columns:
        df["MATERIA"] = ""
    if "CÓD. CURSO" not in df.columns:
        df["CÓD. CURSO"] = ""
    if "REQUISITOS" not in df.columns:
        df["REQUISITOS"] = ""

    return df


# =========================================================
# UTIL: FORMATO APELLIDOS, NOMBRES
# =========================================================
def formatear_apellidos_nombres(apellidos: str, nombres: str) -> str:
    ap = (apellidos or "").strip()
    nm = (nombres or "").strip()
    if not ap and not nm:
        return ""
    if ap and nm:
        return f"{ap}, {nm}"
    return ap or nm
# ============================
# BLOQUE 2 / 4
# LÓGICA ACADÉMICA + PDF HELPERS
# ============================
def calcular_matriculables(df_resultado: pd.DataFrame, df_convalidados: pd.DataFrame):
    if "REQUISITOS" not in df_resultado.columns:
        df_resultado["PUEDE_MATRICULAR"] = False
        return df_resultado, df_resultado[df_resultado["PUEDE_MATRICULAR"]]

    cursos_ok = set(
        df_convalidados["CURSO"].astype(str).str.upper().str.strip().tolist()
    )

    def cumple_req(req):
        if req is None:
            return True
        req = str(req).strip()
        if req == "" or req.lower() == "nan":
            return True

        partes = [p.strip().upper() for p in re.split(r"[;,/]", req) if p.strip()]
        return any(r in cursos_ok for r in partes)

    puede = []
    for _, row in df_resultado.iterrows():
        if row.get("ESTADO_CONVALIDACION") == "CONVALIDADO":
            puede.append(False)
        else:
            puede.append(cumple_req(row.get("REQUISITOS")))

    df_resultado["PUEDE_MATRICULAR"] = puede
    df_matriculables = df_resultado[df_resultado["PUEDE_MATRICULAR"]].copy()
    return df_resultado, df_matriculables


# =========================================================
# PDF HELPERS (UPN)
# =========================================================
def _encabezado_pdf(c, titulo_principal, alumno, codigo, carrera_upn, campus, logo_path):
    width, height = A4

    if os.path.exists(logo_path):
        logo_w, logo_h = 45 * mm, 18 * mm
        c.drawImage(
            logo_path,
            (width - logo_w) / 2,
            height - 60,
            logo_w,
            logo_h,
            preserveAspectRatio=True,
            mask="auto",
        )
        titulo_y = height - 85
    else:
        titulo_y = height - 60

    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(width / 2, titulo_y, str(titulo_principal))

    c.setLineWidth(0.7)
    c.line(40, titulo_y - 6, width - 40, titulo_y - 6)

    c.setFont("Helvetica", 10)
    y = titulo_y - 22

    # ✅ alumno ya llega como "Apellidos, Nombres"
    c.drawString(40, y, f"Apellidos y Nombres: {str(alumno).upper()}")
    c.drawRightString(width - 40, y, f"Código: {str(codigo).upper()}")

    c.drawString(40, y - 14, f"Carrera en UPN: {str(carrera_upn).upper()}")
    c.drawString(40, y - 28, f"Campus: {str(campus).upper()}")

    return y - 45


def _dibujar_tabla_fija_27(c, y: float, df: pd.DataFrame, titulo_columna_curso: str, total_cr: int):
    estilo = ParagraphStyle(
        name="TablaUPN",
        fontName="Helvetica",
        fontSize=8,
        leading=11,
    )

    data = [[
        "Ciclo",
        titulo_columna_curso,
        "Materia",
        "Cód. Curso",
        "CR"
    ]]

    MAX_FILAS = 27

    for _, row in df.iterrows():
        if len(data) - 1 >= MAX_FILAS:
            break
        data.append(
            [
                Paragraph(str(row.get("CICLO", "")), estilo),
                Paragraph(str(row.get("CURSO", "")).upper(), estilo),
                Paragraph(str(row.get("MATERIA", "")).upper(), estilo),
                Paragraph(str(row.get("CÓD. CURSO", "")).upper(), estilo),
                Paragraph(str(row.get("CR", "")), estilo),
            ]
        )

    filas_actuales = len(data) - 1
    for _ in range(MAX_FILAS - filas_actuales):
        data.append(["", "", "", "", ""])

    data.append(["", "", "", "Total", str(total_cr)])

    row_heights = [14] + [16] * MAX_FILAS + [16]

    table = Table(
        data,
        colWidths=[15*mm, 95*mm, 30*mm, 25*mm, 15*mm],
        rowHeights=row_heights,
        repeatRows=1,
    )

    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (0, -1), "CENTER"),
                ("FONT", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONT", (0, -1), (-1, -1), "Helvetica-Bold"),
            ]
        )
    )

    _, table_height = table.wrap(0, 0)
    table.drawOn(c, 40, y - table_height)
    return y - table_height


def _dibujar_firmas(c, elaborado_nombre: str, elaborado_cargo: str, resp_nombre: str, resp_cargo: str):
    c.setFont("Helvetica", 9)
    c.drawString(40, 120, f"Nombre Elaborado por: {str(elaborado_nombre).upper()}")
    c.drawString(40, 108, f"Cargo: {str(elaborado_cargo).upper()}")
    c.drawString(40, 85,  f"Nombre Resp. Acad.: {str(resp_nombre).upper()}")
    c.drawString(40, 73,  f"Cargo: {str(resp_cargo).upper()}")
def _footer_pdf(c, plan_estudios: str, y_referencia: float):
    width, _ = A4
    now = datetime.now()
    fecha = f"{now.day:02d}/{now.month:02d}/{now.year}"

    c.setFont("Helvetica", 9)
    c.drawString(40, y_referencia, f"Plan de Estudios: {plan_estudios}")
    c.drawRightString(width - 40, y_referencia, f"Fecha: {fecha}")

    texto_legal = (
        "Este documento es meramente referencial y emitido por el área académica "
        "para que sirva de guía en el registro de cursos del estudiante. "
        "Es potestad del estudiante elegir y matricularse en los cursos que decida."
    )

    estilo_legal = ParagraphStyle(
        name="LegalFooter",
        fontName="Helvetica",
        fontSize=9,
        leading=11,
    )

    p = Paragraph(texto_legal, estilo_legal)
    _, h = p.wrap(width - 80, 100)
    p.drawOn(c, 40, y_referencia - h - 10)

    y_linea = 40
    c.setLineWidth(0.7)
    c.line(40, y_linea, width - 40, y_linea)

    c.setFont("Helvetica", 8)
    c.drawString(40, 25, "UNIVERSIDAD PRIVADA DEL NORTE S.A.C.")
    c.drawRightString(width - 40, 25, "Versión Conva2025G : 7.63")
# ============================
# BLOQUE 3 / 4
# GENERACIÓN DE PDFs
# ============================
def generar_pdf_convalidados(
    ruta_pdf: str,
    alumno: str,
    codigo: str,
    carrera_upn: str,
    campus: str,
    plan_estudios: str,
    elaborado_nombre: str,
    elaborado_cargo: str,
    resp_nombre: str,
    resp_cargo: str,
    df_convalidados: pd.DataFrame,
    logo_path: str,
):
    c = canvas.Canvas(ruta_pdf, pagesize=A4)

    total_cr = int(
        pd.to_numeric(df_convalidados.get("CR", 0), errors="coerce")
        .fillna(0)
        .sum()
    )

    y = _encabezado_pdf(
        c,
        "RESULTADO DE CONVALIDACIÓN",
        alumno,
        codigo,
        carrera_upn,
        campus,
        logo_path,
    )

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, y - 10, "Relación de cursos convalidados:")
    y -= 15

    _dibujar_tabla_fija_27(
        c=c,
        y=y,
        df=df_convalidados,
        titulo_columna_curso="Cursos Convalidados",
        total_cr=total_cr,
    )

    _dibujar_firmas(c, elaborado_nombre, elaborado_cargo, resp_nombre, resp_cargo)
    _footer_pdf(c, plan_estudios, 195)

    c.showPage()
    c.save()


def generar_pdf_proyeccion(
    ruta_pdf: str,
    alumno: str,
    codigo: str,
    carrera_upn: str,
    campus: str,
    plan_estudios: str,
    elaborado_nombre: str,
    elaborado_cargo: str,
    resp_nombre: str,
    resp_cargo: str,
    df_matriculables: pd.DataFrame,
    logo_path: str,
):
    c = canvas.Canvas(ruta_pdf, pagesize=A4)

    total_cr = int(
        pd.to_numeric(df_matriculables.get("CR", 0), errors="coerce")
        .fillna(0)
        .sum()
    )

    y = _encabezado_pdf(
        c,
        "CURSOS RECOMENDADOS PARA EL REGISTRO DE CURSO",
        alumno,
        codigo,
        carrera_upn,
        campus,
        logo_path,
    )

    c.setFont("Helvetica-Bold", 10)
    c.drawString(40, y - 10, "Relación de cursos recomendados para el registro de curso:")
    y -= 15

    _dibujar_tabla_fija_27(
        c=c,
        y=y,
        df=df_matriculables,
        titulo_columna_curso="Cursos recomendados",
        total_cr=total_cr,
    )

    _dibujar_firmas(c, elaborado_nombre, elaborado_cargo, resp_nombre, resp_cargo)
    _footer_pdf(c, plan_estudios, 195)

    c.showPage()
    c.save()
# ============================
# BLOQUE 4 / 4
# UI + REPORTES (MEJORAS SOLICITADAS)
# ============================
def main(page: ft.Page):
    page.title = "Proyección Malla Curricular UPN"
    page.horizontal_alignment = "center"
    page.scroll = "auto"

    # Dataset
    try:
        df_base = cargar_dataset()
    except Exception as e:
        page.add(ft.Text(f"Error cargando dataset: {e}", color="red", size=16, weight=ft.FontWeight.BOLD))
        return

    carreras = sorted(df_base["CARRERA"].dropna().unique().tolist())

    # ✅ Campos (NUEVO: Nombres y Apellidos separados)
    nombres_field = ft.TextField(label="Nombre(s)", width=300)
    apellidos_field = ft.TextField(label="Apellidos", width=300)

    codigo_field = ft.TextField(label="Código de estudiante", width=200)

    campus_dd = ft.Dropdown(
        label="Campus / Sede",
        options=[ft.dropdown.Option(s) for s in ["BREÑA","CAJAMARCA","CHORRILLOS","COMAS","LOS OLIVOS","TRUJILLO","SAN JUAN","ATE","VIRTUAL"]],
        width=250,
    )

    # ✅ Plan por defecto
    plan_field = ft.TextField(label="Plan de estudios", width=200, value="2025G")

    elaborado_nombre_field = ft.TextField(label="Nombre elaborado por", width=300)
    elaborado_cargo_field = ft.TextField(label="Cargo elaborado por", width=200, value="ASISTENTE")

    resp_nombre_field = ft.TextField(label="Nombre Resp. Académico", width=300)
    resp_cargo_field = ft.TextField(label="Cargo Resp. Académico", width=200, value="COORDINADOR")

    crd_field = ft.TextField(label="CRD (créditos a convalidar)", width=250, keyboard_type=ft.KeyboardType.NUMBER)

    carrera_dd = ft.Dropdown(label="Carrera", options=[ft.dropdown.Option(c) for c in carreras], width=400)
    unidad_dd = ft.Dropdown(label="Unidad de Negocio", options=[], width=400)

    resumen_text = ft.Text("", size=14)
    convalidados_table = ft.Column()
    matriculables_table = ft.Column()

    state = {
        "df_conva": pd.DataFrame(),
        "df_convalidados": pd.DataFrame(),
        "df_matriculables": pd.DataFrame(),
        "df_resultado": pd.DataFrame(),
        "carrera": None,
        "unidad": None,
        "suma_final": 0,
        "crd_int": 0,
        "limite_total": 0,
    }

    def btn_style(bg: str):
        return ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=14),
            padding=ft.padding.symmetric(horizontal=18, vertical=14),
            text_style=ft.TextStyle(size=14, weight=ft.FontWeight.BOLD),
            bgcolor=bg,
            color=ft.Colors.WHITE,
            elevation=2,
        )

    def tabla_convalidados_ui(df: pd.DataFrame, max_rows=20):
        columnas = [("CICLO","CICLO"),("CURSO","CURSO"),("MATERIA","MATERIA"),("CÓD. CURSO","CÓD. CURSO"),("CR","CR"),("REQUISITOS","REQUISITOS")]
        cols = [ft.DataColumn(ft.Text(t)) for t,_ in columnas]
        rows = []
        for _, row in df.head(max_rows).iterrows():
            rows.append(ft.DataRow(cells=[ft.DataCell(ft.Text(str(row.get(col,"")))) for _,col in columnas]))
        return ft.DataTable(columns=cols, rows=rows, column_spacing=20, horizontal_margin=10)

    def tabla_matriculables_ui(df: pd.DataFrame, max_rows=20):
        columnas = [("CICLO","CICLO"),("CURSO","CURSO"),("MATERIA","MATERIA"),("CÓD. CURSO","CÓD. CURSO"),("CR","CR")]
        cols = [ft.DataColumn(ft.Text(t)) for t,_ in columnas]
        rows = []
        for _, row in df.head(max_rows).iterrows():
            rows.append(ft.DataRow(cells=[ft.DataCell(ft.Text(str(row.get(col,"")))) for _,col in columnas]))
        return ft.DataTable(columns=cols, rows=rows, column_spacing=20, horizontal_margin=10)

    def on_carrera_change(e):
        unidad_dd.options.clear()
        unidad_dd.value = None
        if carrera_dd.value:
            df_tmp = df_base[df_base["CARRERA"] == carrera_dd.value]
            unidades = sorted(df_tmp["UNID. NEGOCIO"].dropna().unique())
            unidad_dd.options = [ft.dropdown.Option(u) for u in unidades]
        page.update()

    carrera_dd.on_change = on_carrera_change

    def limpiar_click(e):
        # ✅ Reset campos
        nombres_field.value = ""
        apellidos_field.value = ""
        codigo_field.value = ""
        campus_dd.value = None

        # ✅ vuelve a default 2025G
        plan_field.value = "2025G"

        elaborado_nombre_field.value = ""
        elaborado_cargo_field.value = "ASISTENTE"
        resp_nombre_field.value = ""
        resp_cargo_field.value = "COORDINADOR"

        crd_field.value = ""
        carrera_dd.value = None
        unidad_dd.options.clear()
        unidad_dd.value = None

        resumen_text.value = ""
        convalidados_table.controls.clear()
        matriculables_table.controls.clear()

        state.update(
            {
                "df_conva": pd.DataFrame(),
                "df_convalidados": pd.DataFrame(),
                "df_matriculables": pd.DataFrame(),
                "df_resultado": pd.DataFrame(),
                "carrera": None,
                "unidad": None,
                "suma_final": 0,
                "crd_int": 0,
                "limite_total": 0,
            }
        )

        page.snack_bar = ft.SnackBar(ft.Text("Formulario limpiado. Listo para una nueva convalidación."), bgcolor=ft.Colors.GREEN)
        page.snack_bar.open = True
        page.update()

    def procesar_click(e):
        resumen_text.value = ""
        convalidados_table.controls.clear()
        matriculables_table.controls.clear()

        # ✅ Validación ahora con nombres + apellidos
        if not nombres_field.value or not apellidos_field.value or not codigo_field.value:
            page.snack_bar = ft.SnackBar(ft.Text("Completa Nombre(s), Apellidos y Código"), bgcolor=ft.Colors.RED)
            page.snack_bar.open = True
            page.update()
            return

        if not campus_dd.value or not plan_field.value:
            page.snack_bar = ft.SnackBar(ft.Text("Completa Campus y Plan de estudios"), bgcolor=ft.Colors.RED)
            page.snack_bar.open = True
            page.update()
            return

        if not carrera_dd.value or not unidad_dd.value:
            page.snack_bar = ft.SnackBar(ft.Text("Selecciona Carrera y Unidad de Negocio"), bgcolor=ft.Colors.RED)
            page.snack_bar.open = True
            page.update()
            return

        try:
            crd = float(crd_field.value)
        except Exception:
            page.snack_bar = ft.SnackBar(ft.Text("CRD inválido"), bgcolor=ft.Colors.RED)
            page.snack_bar.open = True
            page.update()
            return

        df_conva = df_base[
            (df_base["CARRERA"] == carrera_dd.value)
            & (df_base["UNID. NEGOCIO"] == unidad_dd.value)
        ].copy()

        seleccion, suma_final = seleccionar_convalidacion(df_conva, crd, tolerancia=2)
        df_convalidados = df_conva.loc[seleccion].copy()

        df_resultado = df_conva.copy()
        df_resultado["ESTADO_CONVALIDACION"] = "NO CONVALIDADO"
        df_resultado.loc[df_convalidados.index, "ESTADO_CONVALIDACION"] = "CONVALIDADO"

        df_resultado, df_matriculables = calcular_matriculables(df_resultado, df_convalidados)

        crd_int = int(float(crd))
        limite_total = crd_int + 2

        suma_real = int(pd.to_numeric(df_convalidados["CR"], errors="coerce").fillna(0).sum())

        state.update(
            {
                "df_conva": df_conva,
                "df_convalidados": df_convalidados,
                "df_matriculables": df_matriculables,
                "df_resultado": df_resultado,
                "carrera": carrera_dd.value,
                "unidad": unidad_dd.value,
                "suma_final": suma_real,
                "crd_int": crd_int,
                "limite_total": limite_total,
            }
        )

        resumen_text.value = (
            f"CRD solicitado: {float(crd):.1f} | "
            f"Convalidados: {suma_real} | "
            f"Máx permitido (CRD+2): {limite_total}"
        )

        if not df_convalidados.empty:
            convalidados_table.controls.append(ft.Text("Cursos convalidados", weight=ft.FontWeight.BOLD))
            convalidados_table.controls.append(tabla_convalidados_ui(df_convalidados))

        if not df_matriculables.empty:
            matriculables_table.controls.append(ft.Text("Cursos matriculables", weight=ft.FontWeight.BOLD))
            matriculables_table.controls.append(tabla_matriculables_ui(df_matriculables))

        page.update()

    def generar_reportes_click(e):
        if state["df_conva"].empty:
            page.snack_bar = ft.SnackBar(ft.Text("Primero procesa la convalidación"), bgcolor=ft.Colors.RED)
            page.snack_bar.open = True
            page.update()
            return

        carrera = state["carrera"]
        unidad = state["unidad"]
        codigo = codigo_field.value.strip()

        # ✅ Formato requerido: "Apellidos, Nombres"
        alumno_fmt = formatear_apellidos_nombres(apellidos_field.value, nombres_field.value)

        nombre_excel = f"Proyeccion_Malla_{carrera}_{unidad}.xlsx"
        ruta_excel = os.path.join(OUTPUT_DIR, nombre_excel)

        ahora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        df_form = pd.DataFrame(
            {
                "Campo": [
                    "Fecha/Hora",
                    "Apellidos y Nombres",
                    "Código de estudiante",
                    "Campus / Sede",
                    "Plan de estudios",
                    "CRD solicitado",
                    "Máximo permitido (CRD+2)",
                    "Créditos convalidados (resultado)",
                    "Carrera",
                    "Unidad de Negocio",
                    "Nombre elaborado por",
                    "Cargo elaborado por",
                    "Nombre Resp. Académico",
                    "Cargo Resp. Académico",
                ],
                "Valor": [
                    ahora,
                    alumno_fmt,  # ✅ ya ordenado
                    codigo_field.value,
                    campus_dd.value,
                    plan_field.value,
                    crd_field.value,
                    state["limite_total"],
                    state["suma_final"],
                    carrera_dd.value,
                    unidad_dd.value,
                    elaborado_nombre_field.value,
                    elaborado_cargo_field.value,
                    resp_nombre_field.value,
                    resp_cargo_field.value,
                ],
            }
        )

        with pd.ExcelWriter(ruta_excel) as writer:
            df_form.to_excel(writer, "Formulario", index=False)
            state["df_conva"].to_excel(writer, "Malla", index=False)
            state["df_convalidados"].to_excel(writer, "Convalidados", index=False)
            state["df_matriculables"].to_excel(writer, "Matriculables", index=False)

        logo_path = os.path.join(BASE_DIR, "logo.jpg")
        carrera_upn = f"{carrera} - {unidad}"

        generar_pdf_convalidados(
            os.path.join(OUTPUT_DIR, f"Resultado_Convalidacion_{codigo}.pdf"),
            alumno_fmt,  # ✅ "Apellidos, Nombres"
            codigo,
            carrera_upn,
            campus_dd.value,
            plan_field.value,
            elaborado_nombre_field.value,
            elaborado_cargo_field.value,
            resp_nombre_field.value,
            resp_cargo_field.value,
            state["df_convalidados"],
            logo_path,
        )

        generar_pdf_proyeccion(
            os.path.join(OUTPUT_DIR, f"Proyeccion_Malla_{codigo}.pdf"),
            alumno_fmt,  # ✅ "Apellidos, Nombres"
            codigo,
            carrera_upn,
            campus_dd.value,
            plan_field.value,
            elaborado_nombre_field.value,
            elaborado_cargo_field.value,
            resp_nombre_field.value,
            resp_cargo_field.value,
            state["df_matriculables"],
            logo_path,
        )

        page.snack_bar = ft.SnackBar(ft.Text("Excel y PDFs generados correctamente"), bgcolor=ft.Colors.GREEN)
        page.snack_bar.open = True
        page.update()

    def copiar_tabla_click(e):
        if state["df_matriculables"].empty:
            page.snack_bar = ft.SnackBar(ft.Text("No hay cursos matriculables para copiar"), bgcolor=ft.Colors.RED)
            page.snack_bar.open = True
            page.update()
            return

        texto = state["df_matriculables"].to_string(index=False)
        page.set_clipboard(texto)

        page.snack_bar = ft.SnackBar(ft.Text("Tabla copiada al portapapeles"), bgcolor=ft.Colors.GREEN)
        page.snack_bar.open = True
        page.update()

    procesar_btn = ft.ElevatedButton(
        "Procesar",
        icon=ft.Icons.CALCULATE,
        on_click=procesar_click,
        height=56,
        width=240,
        style=btn_style("#2563EB"),
    )

    generar_btn = ft.ElevatedButton(
        "Generar Excel/PDF",
        icon=ft.Icons.PICTURE_AS_PDF,
        on_click=generar_reportes_click,
        height=56,
        width=300,
        style=btn_style("#16A34A"),
    )

    copiar_btn = ft.ElevatedButton(
        "Copiar Matriculables",
        icon=ft.Icons.COPY,
        on_click=copiar_tabla_click,
        height=56,
        width=300,
        style=btn_style("#0F766E"),
    )

    limpiar_btn = ft.ElevatedButton(
        "Nueva Convalidación",
        icon=ft.Icons.DELETE_SWEEP,
        on_click=limpiar_click,
        height=56,
        width=300,
        style=btn_style("#DC2626"),
    )

    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.Text("Proyección Malla Curricular UPN", size=20, weight=ft.FontWeight.BOLD),
                    ft.Text("⚠️ Todos los campos del formulario deben ser llenados antes de procesar.", size=12, color="#FF0000"),

                    # ✅ UI: Nombre(s) + Apellidos separados
                    ft.Row([nombres_field, apellidos_field, codigo_field]),
                    ft.Row([campus_dd, plan_field]),
                    ft.Row([elaborado_nombre_field, elaborado_cargo_field]),
                    ft.Row([resp_nombre_field, resp_cargo_field]),
                    ft.Divider(),

                    ft.Row([crd_field]),
                    ft.Row([carrera_dd]),
                    ft.Row([unidad_dd]),

                    ft.Row([procesar_btn, generar_btn], spacing=14),
                    ft.Row([copiar_btn, limpiar_btn], spacing=14),

                    ft.Divider(),
                    resumen_text,
                    ft.Divider(),

                    convalidados_table,
                    ft.Divider(),
                    matriculables_table,
                    ft.Divider(),

                    ft.Row(
                        [
                            ft.Icon(ft.Icons.CONTACT_PAGE, size=18, color="#555555"),
                            ft.Text("Elaborado por: Ing. Jesús Apolaya", size=11, italic=True, color="#555555"),
                        ],
                        alignment=ft.MainAxisAlignment.END,
                    ),
                ],
                expand=True,
                spacing=10,
            ),
            padding=20,
            width=950,
        )
    )


if __name__ == "__main__":
    ft.app(target=main)
