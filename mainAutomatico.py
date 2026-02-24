# ============================
# APP BATCH: EXCEL -> PDFs AUTOMÁTICOS (UPN) | SOLO PDF
# ============================
import flet as ft
import pandas as pd
import re
import os
import sys
import unicodedata
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
LOGO_FILE = "logo.jpg"


# =========================================================
# ⭐ ALGORITMO (selección de convalidación)
# =========================================================
def _subset_best_between(df: pd.DataFrame, min_needed: int, max_allowed: int):
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

    for s in sums:
        if s >= max(0, min_needed) and s <= max_allowed:
            return dp[s], s

    best = max(sums) if sums else 0
    return dp.get(best, []), best


def seleccionar_convalidacion(df_conva: pd.DataFrame, crd: float, tolerancia: int = 2):
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
        if suma_total >= crd_int:
            break

        max_restante = limite_total - suma_total
        if max_restante <= 0:
            break

        min_restante = crd_int - suma_total

        df_ciclo = df[df["CICLO_NUM"] == ciclo].copy().sort_values(by="CR", ascending=False)

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
# UTILIDADES
# =========================================================
def formatear_apellidos_nombres(apellidos: str, nombres: str) -> str:
    ap = (apellidos or "").strip()
    nm = (nombres or "").strip()
    if not ap and not nm:
        return ""
    if ap and nm:
        return f"{ap}, {nm}"
    return ap or nm


def safe_filename(s: str) -> str:
    s = (s or "").strip()
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:120] if len(s) > 120 else s


# =========================================================
# NORMALIZACIÓN ROBUSTA DE COLUMNAS (Excel Input)
# =========================================================
REQUIRED_INPUT_COLS_CANON = [
    "NOMBRE",
    "APELLIDO",
    "COD ESTUDIANTE",
    "SEDE",
    "PLAN DE ESTUDIOS",
    "CARGO ELABORADO POR",
    "CARGO RESP ACADEMICO",
    "CARRERA",
    "UNIDAD DE NEGOCIO",
    "CRD",
]

SYNONYMS = {
    "COD ESTUDIANTE": ["COD ESTUDIANTE", "CODIGO ESTUDIANTE", "CÓDIGO ESTUDIANTE", "COD. ESTUDIANTE", "COD EST."],
    "CARGO RESP ACADEMICO": [
        "CARGO RESP ACADEMICO",
        "CARGO RESP. ACADEMICO",
        "CARGO RESP. ACADÉMICO",
        "CARGO RESPONSABLE ACADEMICO",
    ],
    "UNIDAD DE NEGOCIO": ["UNIDAD DE NEGOCIO", "UNIDAD NEGOCIO", "UNIDAD"],
    "PLAN DE ESTUDIOS": ["PLAN DE ESTUDIOS", "PLAN ESTUDIOS", "PLAN"],
}

def _canon(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\u00A0", " ")  # NBSP
    s = s.strip().upper()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"[^\w\s]", " ", s)      # quita signos/puntos
    s = re.sub(r"\s+", " ", s).strip()  # colapsa espacios
    return s

def normalize_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [_canon(c) for c in df.columns]
    return df

def ensure_required_cols(df: pd.DataFrame):
    present = set(df.columns)
    missing = []
    for req in REQUIRED_INPUT_COLS_CANON:
        variants = [req] + SYNONYMS.get(req, [])
        variants = [_canon(v) for v in variants]
        if not any(v in present for v in variants):
            missing.append(req)

    if missing:
        raise ValueError(
            "Faltan columnas obligatorias en el Excel: "
            f"{missing}\n"
            f"Columnas detectadas: {sorted(list(present))}"
        )

def get_cell(row, col, default=""):
    col_c = _canon(col)
    v = row.get(col_c, default)
    if pd.isna(v):
        return default
    return str(v).strip()


# =========================================================
# LÓGICA: MATRICULABLES
# =========================================================
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
# PDF HELPERS
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

    c.drawString(40, y, f"Apellidos y Nombres: {str(alumno).upper()}")
    c.drawRightString(width - 40, y, f"Código: {str(codigo).upper()}")

    c.drawString(40, y - 14, f"Carrera en UPN: {str(carrera_upn).upper()}")
    c.drawString(40, y - 28, f"Sede: {str(campus).upper()}")

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


def generar_pdf_convalidados(
    ruta_pdf: str,
    alumno: str,
    codigo: str,
    carrera_upn: str,
    sede: str,
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
        sede,
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
    sede: str,
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
        sede,
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


# =========================================================
# UI (Flet) - SOLO CARGA Y PROCESO AUTOMÁTICO
# =========================================================
def main(page: ft.Page):
    page.title = "UPN - Proyección Malla (Proceso Masivo desde Excel) - SOLO PDF"
    page.horizontal_alignment = "center"
    page.scroll = "auto"

    # Cargar dataset base una sola vez
    try:
        df_base = cargar_dataset()
    except Exception as e:
        page.add(ft.Text(f"Error cargando dataset: {e}", color="red", size=16, weight=ft.FontWeight.BOLD))
        return

    # normalizar dataset para filtros exactos
    df_base_norm = df_base.copy()
    df_base_norm["CARRERA"] = df_base_norm["CARRERA"].astype(str).str.strip()
    df_base_norm["UNID. NEGOCIO"] = df_base_norm["UNID. NEGOCIO"].astype(str).str.strip()

    # UI controls
    status_text = ft.Text("Carga un Excel y el sistema generará PDFs automáticamente.", size=13)
    progress = ft.ProgressBar(width=700, value=0)
    log_box = ft.TextField(
        label="Log",
        multiline=True,
        min_lines=10,
        max_lines=14,
        read_only=True,
        width=980
    )

    def log(msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        log_box.value = (log_box.value or "") + f"[{ts}] {msg}\n"
        page.update()

    def btn_style(bg: str):
        return ft.ButtonStyle(
            shape=ft.RoundedRectangleBorder(radius=14),
            padding=ft.padding.symmetric(horizontal=18, vertical=14),
            text_style=ft.TextStyle(size=14, weight=ft.FontWeight.BOLD),
            bgcolor=bg,
            color=ft.Colors.WHITE,
            elevation=2,
        )

    # FilePicker
    picker = ft.FilePicker()
    page.overlay.append(picker)

    def on_pick_result(e: ft.FilePickerResultEvent):
        if not e.files:
            return
        ruta_excel = e.files[0].path
        if not ruta_excel or not os.path.exists(ruta_excel):
            page.snack_bar = ft.SnackBar(ft.Text("No se pudo leer la ruta del archivo."), bgcolor=ft.Colors.RED)
            page.snack_bar.open = True
            page.update()
            return

        run_batch(ruta_excel)

    picker.on_result = on_pick_result

    def run_batch(ruta_excel_in: str):
        log_box.value = ""
        progress.value = 0
        page.update()

        log(f"Archivo cargado: {ruta_excel_in}")

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_root = os.path.join(BASE_DIR, f"PROCESADOS_{stamp}")
        os.makedirs(out_root, exist_ok=True)
        log(f"Carpeta de salida: {out_root}")

        logo_path = os.path.join(BASE_DIR, LOGO_FILE)
        if not os.path.exists(logo_path):
            log("⚠️ No se encontró logo.jpg. Los PDFs saldrán sin logo.")
        else:
            log("Logo detectado correctamente (logo.jpg).")

        # Leer excel input
        df_in = pd.read_excel(ruta_excel_in, header=0)
        df_in = normalize_cols(df_in)

        # Log columnas reales detectadas
        log(f"Columnas detectadas: {list(df_in.columns)}")

        # Validar columnas
        ensure_required_cols(df_in)

        total = len(df_in)
        if total == 0:
            log("El Excel no tiene filas para procesar.")
            return

        ok_count = 0
        err_count = 0

        for i, row in df_in.iterrows():
            idx = i + 1
            progress.value = idx / total
            status_text.value = f"Procesando {idx}/{total}..."
            page.update()

            try:
                nombre = get_cell(row, "NOMBRE")
                apellido = get_cell(row, "APELLIDO")
                codigo = get_cell(row, "COD ESTUDIANTE")
                sede = get_cell(row, "SEDE")
                plan = get_cell(row, "PLAN DE ESTUDIOS")
                carrera = get_cell(row, "CARRERA")
                unidad = get_cell(row, "UNIDAD DE NEGOCIO")
                crd = float(get_cell(row, "CRD"))

                cargo_elab = get_cell(row, "CARGO ELABORADO POR")
                cargo_resp = get_cell(row, "CARGO RESP ACADEMICO")

                # Si quieres nombres también en el excel, agrega columnas:
                # "Nombre elaborado por" y "Nombre resp. Académico"
                # y aquí las leemos (si no existen, quedará vacío)
                nombre_elab = get_cell(row, "NOMBRE ELABORADO POR", default="")
                nombre_resp = get_cell(row, "NOMBRE RESP ACADEMICO", default="")

                if not codigo:
                    raise ValueError("COD ESTUDIANTE vacío")

                alumno_fmt = formatear_apellidos_nombres(apellido, nombre)

                folder_name = safe_filename(f"{codigo}_{apellido}_{nombre}")
                out_student = os.path.join(out_root, folder_name)
                os.makedirs(out_student, exist_ok=True)

                df_conva = df_base_norm[
                    (df_base_norm["CARRERA"] == carrera)
                    & (df_base_norm["UNID. NEGOCIO"] == unidad)
                ].copy()

                if df_conva.empty:
                    raise ValueError(f"No hay registros en dataset para Carrera='{carrera}' y Unidad='{unidad}'")

                seleccion, _ = seleccionar_convalidacion(df_conva, crd, tolerancia=2)
                df_convalidados = df_conva.loc[seleccion].copy()

                df_resultado = df_conva.copy()
                df_resultado["ESTADO_CONVALIDACION"] = "NO CONVALIDADO"
                df_resultado.loc[df_convalidados.index, "ESTADO_CONVALIDACION"] = "CONVALIDADO"
                df_resultado, df_matriculables = calcular_matriculables(df_resultado, df_convalidados)

                carrera_upn = f"{carrera} - {unidad}"

                pdf_conva = os.path.join(out_student, f"Resultado_Convalidacion_{codigo}.pdf")
                pdf_proy = os.path.join(out_student, f"Proyeccion_Malla_{codigo}.pdf")

                generar_pdf_convalidados(
                    pdf_conva,
                    alumno_fmt,
                    codigo,
                    carrera_upn,
                    sede,
                    plan,
                    nombre_elab,
                    cargo_elab,
                    nombre_resp,
                    cargo_resp,
                    df_convalidados,
                    logo_path,
                )

                generar_pdf_proyeccion(
                    pdf_proy,
                    alumno_fmt,
                    codigo,
                    carrera_upn,
                    sede,
                    plan,
                    nombre_elab,
                    cargo_elab,
                    nombre_resp,
                    cargo_resp,
                    df_matriculables,
                    logo_path,
                )

                ok_count += 1
                log(f"✅ {idx}/{total} OK - {codigo} - {alumno_fmt}")

            except Exception as ex:
                err_count += 1
                log(f"❌ {idx}/{total} ERROR - {get_cell(row, 'COD ESTUDIANTE')} -> {ex}")

        progress.value = 1
        status_text.value = f"Proceso finalizado. OK={ok_count} | ERROR={err_count}"
        page.update()

        page.snack_bar = ft.SnackBar(ft.Text(f"Terminado. Revisa: {out_root}"), bgcolor=ft.Colors.GREEN)
        page.snack_bar.open = True
        page.update()

    def seleccionar_excel_click(e):
        picker.pick_files(
            allow_multiple=False,
            allowed_extensions=["xlsx", "xls"]
        )

    seleccionar_btn = ft.ElevatedButton(
        "Cargar Excel y Generar PDFs",
        icon=ft.Icons.UPLOAD_FILE,
        on_click=seleccionar_excel_click,
        height=56,
        width=340,
        style=btn_style("#2563EB"),
    )

    page.add(
        ft.Container(
            content=ft.Column(
                [
                    ft.Text("UPN - Proyección Malla (Masivo desde Excel) - SOLO PDF", size=20, weight=ft.FontWeight.BOLD),
                    ft.Text(
                        "Excel requerido: Nombre | Apellido | Cod estudiante | Sede | Plan de estudios | "
                        "Cargo elaborado por | Cargo resp. Académico | Carrera | Unidad de negocio | CRD",
                        size=12,
                        color="#555555",
                    ),
                    ft.Row([seleccionar_btn], spacing=14),
                    status_text,
                    progress,
                    log_box,
                    ft.Row(
                        [
                            ft.Icon(ft.Icons.CONTACT_PAGE, size=18, color="#555555"),
                            ft.Text("Elaborado por: Ing. Jesús Apolaya", size=11, italic=True, color="#555555"),
                        ],
                        alignment=ft.MainAxisAlignment.END,
                    ),
                ],
                spacing=10,
            ),
            padding=20,
            width=1000,
        )
    )


if __name__ == "__main__":
    ft.app(target=main)