from __future__ import annotations

import json
import math
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import flet as ft
import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


# =========================================================
# APP CONVALIDACIÓN / PROYECCIÓN MALLA UPN
# Migrado de Apps Script a Python + Flet
# Desarrollado: Ing. Jesus Apolaya
# =========================================================

APP_TITLE = "App Convalidación / Proyección Malla UPN"
APP_VERSION = "2.0.0"
PDF_VERSION = "7.63"
DATASET_FILE = "dataset.xlsx"
OUTPUT_DIR = "APP_CONVALIDACION_SALIDAS"
TOLERANCIA_CRD = 2
RESPONSIVE_WIDTH = 1180

SHEET_PAQUETE = "PAQUETE"
SHEET_LOG = "LOG_APP"
SHEET_SEDES = "MAESTRO_SEDES"
SHEET_MALLA = "MALLA"
SHEET_CENTROS = "Maestro_Centro_Estudios"
SHEET_RESPONSABLES = "MAESTRO_RESPONSABLE_ACADEMICO"


# =========================================================
# HELPERS
# =========================================================

def normalize_key(value: Any) -> str:
    txt = str(value or "").strip()
    txt = re.sub(r"\s+", " ", txt)
    replacements = str.maketrans(
        "ÁÉÍÓÚáéíóúÑñ",
        "AEIOUaeiouNn",
    )
    txt = txt.translate(replacements)
    txt = re.sub(r"[^\w\s]", "", txt)
    txt = txt.replace(" ", "_")
    return txt.upper()


def normalize_text_search(value: Any) -> str:
    txt = str(value or "").strip()
    replacements = str.maketrans(
        "ÁÉÍÓÚáéíóúÑñ",
        "AEIOUaeiouNn",
    )
    txt = txt.translate(replacements).upper()
    txt = re.sub(r"\s+", " ", txt)
    return txt


def number_safe(value: Any) -> float:
    if value is None:
        return 0
    txt = str(value).strip().replace(",", ".")
    if txt == "" or txt.lower() == "nan":
        return 0
    try:
        return float(txt)
    except Exception:
        return 0


def format_date(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y")


def format_datetime(dt: datetime) -> str:
    return dt.strftime("%d/%m/%Y %H:%M:%S")


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|#%&{}$!\'@+=`]+', "", str(name or "SIN_NOMBRE")).strip()


def to_upper_safe(value: Any) -> str:
    return str(value or "").upper().strip()


def formatear_alumno(apellidos: str, nombres: str) -> str:
    ap = str(apellidos or "").strip()
    nm = str(nombres or "").strip()
    if ap and nm:
        return f"{ap}, {nm}"
    return ap or nm


def extract_cycle_number(value: Any) -> int:
    m = re.search(r"(\d+)", str(value or ""))
    return int(m.group(1)) if m else 0


def normalize_malla_value(value: Any) -> str:
    txt = str(value or "").strip().upper().replace(" ", "")
    if txt in ("2023", "PE2023"):
        return "2023"
    if txt in ("2025G", "PE2025G", "PE2025-G"):
        return "2025G"
    if txt in ("2026", "PE2026"):
        return "2026"
    return txt


# =========================================================
# DATASET REPOSITORY
# =========================================================

class DatasetRepository:
    def __init__(self, excel_path: str | Path):
        self.excel_path = Path(excel_path)
        if not self.excel_path.exists():
            raise FileNotFoundError(
                f"No se encontró el archivo {self.excel_path.name}. Debe estar en la misma carpeta del script."
            )
        self.book = pd.ExcelFile(self.excel_path)
        self._cache: Dict[str, pd.DataFrame] = {}

    def read_sheet(self, sheet_name: str) -> pd.DataFrame:
        if sheet_name in self._cache:
            return self._cache[sheet_name].copy()

        if sheet_name not in self.book.sheet_names:
            raise ValueError(f"No existe la hoja: {sheet_name}")

        df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
        df.columns = [normalize_key(c) for c in df.columns]
        df = df.fillna("")
        self._cache[sheet_name] = df.copy()
        return df.copy()

    def get_paquetes(self) -> List[str]:
        df = self.read_sheet(SHEET_PAQUETE)
        if "PAQUETE" not in df.columns:
            return []
        return sorted([str(x).strip() for x in df["PAQUETE"].tolist() if str(x).strip()])

    def get_sedes(self) -> List[str]:
        df = self.read_sheet(SHEET_SEDES)
        col = "DESCRIPCION" if "DESCRIPCION" in df.columns else (df.columns[0] if len(df.columns) else None)
        if not col:
            return []
        return sorted([str(x).strip() for x in df[col].tolist() if str(x).strip()])

    def get_carreras(self) -> List[str]:
        df = self.read_sheet(SHEET_MALLA)
        if "CARRERA" not in df.columns:
            return []
        return sorted(df["CARRERA"].astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())

    def get_responsables(self) -> List[Dict[str, str]]:
        df = self.read_sheet(SHEET_RESPONSABLES)
        out = []
        for _, r in df.iterrows():
            nombre = str(r.get("NOMBRE", "")).strip()
            if not nombre:
                continue
            out.append(
                {
                    "nombre": nombre,
                    "cargo": str(r.get("CARGO", "")).strip(),
                    "correo": str(r.get("CORREO", "")).strip(),
                    "grupo": str(r.get("GRUPO", "")).strip(),
                }
            )
        return out

    def get_responsable_by_nombre(self, nombre: str) -> Optional[Dict[str, str]]:
        objetivo = normalize_text_search(nombre)
        for item in self.get_responsables():
            if normalize_text_search(item["nombre"]) == objetivo:
                return item
        return None

    def search_instituciones(self, query: str) -> List[Dict[str, str]]:
        df = self.read_sheet(SHEET_CENTROS)
        q = normalize_text_search(query)
        out: List[Dict[str, str]] = []
        for _, r in df.iterrows():
            nombre = str(r.get("NOMBRE") or r.get("INSTITUCION_S") or r.get("INSTITUCION_C") or "").strip()
            if not nombre:
                continue
            tipo = str(r.get("INSTITUCION_TIPO", "")).strip().upper()
            codigo = str(r.get("COD_INSTITUCION", "")).strip()
            if q and q not in normalize_text_search(nombre):
                continue
            out.append({"nombre": nombre, "tipo": tipo, "codigo": codigo})
        out.sort(key=lambda x: x["nombre"].lower())
        return out[:50]

    def get_centro_by_nombre(self, nombre: str) -> Optional[Dict[str, str]]:
        objetivo = normalize_text_search(nombre)
        if not objetivo:
            return None
        for item in self.search_instituciones(""):
            if normalize_text_search(item["nombre"]) == objetivo:
                return item
        return None

    def resolver_malla_existente(self, malla_canonica: str) -> str:
        df = self.read_sheet(SHEET_MALLA)
        if "MALLA" not in df.columns:
            return malla_canonica
        disponibles = [str(x).strip() for x in df["MALLA"].tolist() if str(x).strip()]
        objetivo = normalize_malla_value(malla_canonica)
        for v in disponibles:
            if normalize_malla_value(v) == objetivo:
                return v
        return malla_canonica

    def get_unidades_by_carrera_and_malla(self, carrera: str, malla: str) -> List[str]:
        df = self.read_sheet(SHEET_MALLA)
        if not {"CARRERA", "MALLA", "UNID_NEGOCIO"}.issubset(df.columns):
            return []
        mask = (
            df["CARRERA"].astype(str).str.strip().eq(str(carrera).strip())
            & df["MALLA"].map(normalize_malla_value).eq(normalize_malla_value(malla))
        )
        unidades = df.loc[mask, "UNID_NEGOCIO"].astype(str).str.strip()
        return sorted(unidades.replace("", pd.NA).dropna().unique().tolist())

    def get_malla_preview(self, carrera: str, unidad: str, malla: str) -> List[Dict[str, Any]]:
        df = self.read_sheet(SHEET_MALLA)
        required = {"CARRERA", "UNID_NEGOCIO", "MALLA"}
        if not required.issubset(df.columns):
            return []
        mask = (
            df["CARRERA"].astype(str).str.strip().eq(str(carrera).strip())
            & df["UNID_NEGOCIO"].astype(str).str.strip().eq(str(unidad).strip())
            & df["MALLA"].map(normalize_malla_value).eq(normalize_malla_value(malla))
        )
        mdf = df.loc[mask].copy()
        if mdf.empty:
            return []

        mdf["_CICLO_NUM"] = mdf.get("CICLO", "").map(extract_cycle_number)
        mdf["_UBI"] = mdf.get("UBICACION_EN_EL_CICLO", "").map(number_safe)
        mdf = mdf.sort_values(["_CICLO_NUM", "_UBI"], ascending=[True, True])

        out = []
        for _, r in mdf.iterrows():
            out.append(
                {
                    "CICLO": r.get("CICLO", ""),
                    "CURSO": r.get("CURSO", ""),
                    "MATERIA": r.get("MATERIA", ""),
                    "COD_CURSO": r.get("COD_CURSO") or r.get("COD_CURSO_") or r.get("COD_CURSO___") or r.get("CODIGO_OFICIAL") or "",
                    "CR": number_safe(r.get("CR", 0)),
                    "REQUISITOS": r.get("REQUISITOS", ""),
                }
            )
        return out


# =========================================================
# REGLA DE MALLA EDITABLE
# =========================================================

@dataclass
class MallaRuleConfig:
    cad_2025g_min: int = 44
    cad_2023_min: int = 65
    general_2025g_min: int = 24
    general_2023_min: int = 65

    def get_malla_canonica(self, sede: str, crd: float) -> str:
        sede_txt = str(sede or "").strip().upper()
        n = int(number_safe(crd))
        if not sede_txt or n <= 0:
            return ""

        if sede_txt == "CAD":
            if n >= self.cad_2023_min:
                return "2023"
            if n >= self.cad_2025g_min:
                return "2025G"
            return "2026"

        if n >= self.general_2023_min:
            return "2023"
        if n >= self.general_2025g_min:
            return "2025G"
        return "2026"

    def build_rule_text(self, sede: str) -> str:
        sede_txt = str(sede or "").strip().upper()
        if sede_txt == "CAD":
            return (
                f"CAD 2.0 | PE2023: desde {self.cad_2023_min} | "
                f"PE2025G: entre {self.cad_2025g_min} y {self.cad_2023_min - 1} | "
                f"PE2026: hasta {self.cad_2025g_min - 1}"
            )
        return (
            f"Regla general | PE2023: desde {self.general_2023_min} | "
            f"PE2025G: entre {self.general_2025g_min} y {self.general_2023_min - 1} | "
            f"PE2026: hasta {self.general_2025g_min - 1}"
        )


# =========================================================
# ALGORITMO ACADÉMICO
# =========================================================

def subset_best_between(items: List[Dict[str, Any]], min_needed: float, max_allowed: float) -> Dict[str, Any]:
    if not items or max_allowed <= 0:
        return {"indexes": [], "sum": 0}

    dp: Dict[int, List[int]] = {0: []}
    for item in items:
        current_keys = sorted(dp.keys(), reverse=True)
        for s in current_keys:
            new_sum = int(s + item["cr"])
            if new_sum <= max_allowed and new_sum not in dp:
                dp[new_sum] = dp[s] + [item["index"]]

    sums = sorted(dp.keys())
    floor_needed = max(0, int(min_needed))
    for s in sums:
        if floor_needed <= s <= max_allowed:
            return {"indexes": dp[s], "sum": s}

    best = max(sums) if sums else 0
    return {"indexes": dp.get(best, []), "sum": best}


def seleccionar_convalidacion(rows: List[Dict[str, Any]], crd: float, tolerancia: int) -> Dict[str, Any]:
    if not rows:
        return {"seleccion": [], "suma": 0}

    crd_int = int(number_safe(crd))
    limite_total = crd_int + tolerancia

    rows2 = []
    for i, r in enumerate(rows):
        row = dict(r)
        row["__index"] = i
        row["CR_NUM"] = int(number_safe(r.get("CR", 0)))
        row["CICLO_NUM"] = extract_cycle_number(r.get("CICLO", ""))
        rows2.append(row)

    ciclos = sorted({r["CICLO_NUM"] for r in rows2 if r["CICLO_NUM"] > 0})

    seleccion_total: List[int] = []
    suma_total = 0

    for ciclo in ciclos:
        if suma_total >= crd_int:
            break
        max_restante = limite_total - suma_total
        if max_restante <= 0:
            break
        min_restante = crd_int - suma_total

        df_ciclo = [r for r in rows2 if r["CICLO_NUM"] == ciclo and r["CR_NUM"] > 0]
        df_ciclo.sort(key=lambda x: x["CR_NUM"], reverse=True)

        items = [{"index": r["__index"], "cr": r["CR_NUM"]} for r in df_ciclo]
        best = subset_best_between(items, min_restante, max_restante)
        if best["sum"] > 0:
            seleccion_total.extend(best["indexes"])
            suma_total += best["sum"]

    unicos = list(dict.fromkeys(seleccion_total))
    return {"seleccion": unicos, "suma": suma_total}


def calcular_matriculables(resultado_rows: List[Dict[str, Any]], convalidados_rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    cursos_ok = {str(r.get("CURSO", "")).strip().upper() for r in convalidados_rows if str(r.get("CURSO", "")).strip()}
    out = []
    for r in resultado_rows:
        estado = str(r.get("ESTADO_CONVALIDACION", ""))
        req = str(r.get("REQUISITOS", "")).strip()
        puede = False
        if estado == "CONVALIDADO":
            puede = False
        elif not req or req.lower() == "nan":
            puede = True
        else:
            partes = [x.strip().upper() for x in re.split(r"[;,/]", req) if x.strip()]
            puede = len(partes) == 0 or any(p in cursos_ok for p in partes)
        if puede:
            row = dict(r)
            row["PUEDE_MATRICULAR"] = True
            out.append(row)
    return out


# =========================================================
# EXPORTADOR PDF Y ARCHIVOS
# =========================================================

class ExportService:
    def __init__(self, base_dir: str | Path = OUTPUT_DIR):
        self.base_dir = Path(base_dir)
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self.styles = getSampleStyleSheet()

    def create_run_folder(self, codigo: str, alumno: str) -> Path:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        folder = self.base_dir / f"{sanitize_filename(codigo)} - {sanitize_filename(alumno)} - {ts}"
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    def build_pdf(self, datos: Dict[str, Any], rows: List[Dict[str, Any]], total_cr: float, out_path: Path,
                  title: str, subtitle: str, titulo_curso: str):
        doc = SimpleDocTemplate(
            str(out_path),
            pagesize=A4,
            rightMargin=12 * mm,
            leftMargin=12 * mm,
            topMargin=10 * mm,
            bottomMargin=10 * mm,
        )
        story = []

        title_style = self.styles["Title"].clone("title_custom")
        title_style.fontName = "Helvetica-Bold"
        title_style.fontSize = 11
        title_style.leading = 14
        title_style.alignment = 1

        normal = self.styles["Normal"].clone("normal_small")
        normal.fontName = "Helvetica"
        normal.fontSize = 8
        normal.leading = 10

        small_bold = self.styles["Normal"].clone("small_bold")
        small_bold.fontName = "Helvetica-Bold"
        small_bold.fontSize = 8
        small_bold.leading = 10

        story.append(Paragraph("UNIVERSIDAD PRIVADA DEL NORTE", small_bold))
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(title, title_style))
        story.append(Spacer(1, 2 * mm))

        alumno = f"Apellidos y Nombres: {to_upper_safe(datos['alumno'])}"
        codigo = f"Código: {to_upper_safe(datos['codigo'])}"
        hdr = Table([[alumno, codigo]], colWidths=[140 * mm, 40 * mm])
        hdr.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(hdr)
        story.append(Spacer(1, 2 * mm))

        info_lines = [
            f"Carrera en UPN: {to_upper_safe(datos['carrera'])} - {to_upper_safe(datos['unidad'])}",
            f"Campus: {to_upper_safe(datos['sede'])}",
            f"Institución de procedencia: {to_upper_safe(datos['institucionProcedencia'])}",
            f"Procedencia: {to_upper_safe(datos['tipoInstitucionProcedencia'])}",
            f"Carrera de procedencia: {to_upper_safe(datos['carreraProcedencia'])}",
        ]
        if datos["tipoCaso"] == "PAQUETE":
            info_lines.append(f"Tipo de Paquete: {to_upper_safe(datos['tipoPaquete'])}")
        for line in info_lines:
            story.append(Paragraph(line, normal))

        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(subtitle, small_bold))
        story.append(Spacer(1, 1 * mm))

        table_data = [["Ciclo", titulo_curso, "Materia", "Cód. Curso", "CR"]]
        max_rows = 27
        for r in rows[:max_rows]:
            table_data.append([
                str(r.get("CICLO", "")),
                str(r.get("CURSO", "")).upper(),
                str(r.get("MATERIA", "")).upper(),
                str(r.get("COD_CURSO", "")),
                str(int(number_safe(r.get("CR", 0)))) if number_safe(r.get("CR", 0)).is_integer() else str(r.get("CR", "")),
            ])
        while len(table_data) - 1 < max_rows:
            table_data.append(["", "", "", "", ""])
        table_data.append(["", "", "", "Total", str(int(total_cr) if float(total_cr).is_integer() else total_cr)])

        course_table = Table(table_data, repeatRows=1, colWidths=[15 * mm, 80 * mm, 45 * mm, 25 * mm, 15 * mm])
        course_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.4),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF1FB")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3),
            ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("ALIGN", (4, 1), (4, -1), "CENTER"),
            ("FONTNAME", (3, -1), (4, -1), "Helvetica-Bold"),
        ]))
        story.append(course_table)
        story.append(Spacer(1, 2 * mm))

        plan_table = Table([[f"Plan de Estudios: {datos['malla']}", f"Fecha: {format_date(datetime.now())}"]], colWidths=[120 * mm, 60 * mm])
        plan_table.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.0, colors.white),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]))
        story.append(plan_table)
        story.append(Spacer(1, 2 * mm))

        legal_text = (
            "Este documento es meramente referencial y emitido por el área académica para que sirva de guía en el "
            "registro de cursos del estudiante. Es potestad del estudiante elegir y matricularse en los cursos que decida."
        )
        story.append(Paragraph(legal_text, normal))
        story.append(Spacer(1, 3 * mm))

        story.append(Paragraph(f"Nombre Elaborado por: {to_upper_safe(datos['elaboradoNombre'])}", normal))
        story.append(Paragraph(f"Cargo: {to_upper_safe(datos['elaboradoCargo'])}", normal))
        story.append(Spacer(1, 2 * mm))
        story.append(Paragraph(f"Nombre Resp. Acad.: {to_upper_safe(datos['respNombre'])}", normal))
        story.append(Paragraph(f"Cargo: {to_upper_safe(datos['respCargo'])}", normal))
        story.append(Spacer(1, 3 * mm))
        story.append(Paragraph(f"UNIVERSIDAD PRIVADA DEL NORTE S.A.C. | Versión Conva{datos['malla']} : {PDF_VERSION}", normal))

        doc.build(story)

    def save_json(self, datos: Dict[str, Any], convalidados: List[Dict[str, Any]], matriculables: List[Dict[str, Any]], folder: Path) -> Path:
        payload = {
            "generadoEl": format_datetime(datetime.now()),
            "alumno": datos["alumno"],
            "codigo": datos["codigo"],
            "sede": datos["sede"],
            "carrera": datos["carrera"],
            "unidad": datos["unidad"],
            "malla": datos["malla"],
            "tipoCaso": datos["tipoCaso"],
            "tipoPaquete": datos["tipoPaquete"],
            "crd": datos["crd"],
            "institucionProcedencia": datos["institucionProcedencia"],
            "tipoInstitucionProcedencia": datos["tipoInstitucionProcedencia"],
            "codigoInstitucionProcedencia": datos["codigoInstitucionProcedencia"],
            "carreraProcedencia": datos["carreraProcedencia"],
            "elaboradoPor": datos["elaboradoNombre"],
            "cargoElaboradoPor": datos["elaboradoCargo"],
            "responsableAcademico": datos["respNombre"],
            "cargoResponsableAcademico": datos["respCargo"],
            "correoResponsableAcademico": datos["respCorreo"],
            "grupoResponsableAcademico": datos["respGrupo"],
            "convalidados": convalidados,
            "matriculables": matriculables,
        }
        path = folder / f"Resumen_{sanitize_filename(datos['codigo'])}.json"
        path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
        return path


# =========================================================
# LOG LOCAL EN EXCEL
# =========================================================

def ensure_log_book(log_path: Path):
    columns = [
        "FECHA_HORA", "CODIGO", "ALUMNO", "SEDE", "CARRERA", "UNIDAD", "MALLA",
        "TIPO_CASO", "TIPO_PAQUETE", "CRD", "INSTITUCION_PROCEDENCIA",
        "TIPO_INSTITUCION_PROCEDENCIA", "CARRERA_PROCEDENCIA", "RESPONSABLE_ACADEMICO",
        "CARGO_RESPONSABLE_ACADEMICO", "CREDITOS_CONVALIDADOS", "DOCENTE_CONVALIDA",
        "PDF_CONVALIDACION", "PDF_PROYECCION", "JSON_RESUMEN"
    ]
    if not log_path.exists():
        pd.DataFrame(columns=columns).to_excel(log_path, index=False)


def append_log(log_path: Path, row: Dict[str, Any]):
    ensure_log_book(log_path)
    df = pd.read_excel(log_path).fillna("")
    df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
    df.to_excel(log_path, index=False)


# =========================================================
# PROCESAMIENTO PRINCIPAL
# =========================================================

class ConvalidacionService:
    def __init__(self, repo: DatasetRepository, exporter: ExportService, rules: MallaRuleConfig):
        self.repo = repo
        self.exporter = exporter
        self.rules = rules
        self.log_path = Path(OUTPUT_DIR) / "LOG_APP.xlsx"

    def get_malla_automatica(self, sede: str, crd: float) -> Dict[str, str]:
        if not str(sede).strip():
            return {"mallaCanonica": "", "mallaReal": "", "regla": "Primero selecciona la sede."}
        regla = self.rules.build_rule_text(sede)
        if number_safe(crd) <= 0:
            return {"mallaCanonica": "", "mallaReal": "", "regla": regla}
        canonica = self.rules.get_malla_canonica(sede, crd)
        real = self.repo.resolver_malla_existente(canonica)
        return {"mallaCanonica": canonica, "mallaReal": real, "regla": regla}

    def validar_payload(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        malla_info = self.get_malla_automatica(payload.get("sede", ""), payload.get("crd", 0))
        centro = self.repo.get_centro_by_nombre(payload.get("institucionProcedencia", ""))
        responsable = self.repo.get_responsable_by_nombre(payload.get("respNombre", ""))

        data = {
            "nombres": str(payload.get("nombres", "")).strip(),
            "apellidos": str(payload.get("apellidos", "")).strip(),
            "alumno": formatear_alumno(payload.get("apellidos", ""), payload.get("nombres", "")),
            "codigo": str(payload.get("codigo", "")).strip(),
            "sede": str(payload.get("sede", "")).strip(),
            "carrera": str(payload.get("carrera", "")).strip(),
            "unidad": str(payload.get("unidad", "")).strip(),
            "malla": str(payload.get("malla", "") or malla_info["mallaReal"]).strip(),
            "tipoCaso": str(payload.get("tipoCaso", "")).strip().upper(),
            "tipoPaquete": str(payload.get("tipoPaquete", "")).strip(),
            "crd": int(number_safe(payload.get("crd", 0))),
            "elaboradoNombre": str(payload.get("elaboradoNombre", "")).strip(),
            "elaboradoCargo": str(payload.get("elaboradoCargo", "")).strip(),
            "institucionProcedencia": str((centro or {}).get("nombre") or payload.get("institucionProcedencia", "")).strip(),
            "tipoInstitucionProcedencia": str((centro or {}).get("tipo") or payload.get("tipoInstitucionProcedencia", "")).strip(),
            "codigoInstitucionProcedencia": str((centro or {}).get("codigo") or "").strip(),
            "carreraProcedencia": str(payload.get("carreraProcedencia", "")).strip(),
            "respNombre": str((responsable or {}).get("nombre") or payload.get("respNombre", "")).strip(),
            "respCargo": str((responsable or {}).get("cargo") or payload.get("respCargo", "")).strip(),
            "respCorreo": str((responsable or {}).get("correo") or "").strip(),
            "respGrupo": str((responsable or {}).get("grupo") or "").strip(),
        }

        if not data["nombres"]:
            raise ValueError("Completa el nombre.")
        if not data["apellidos"]:
            raise ValueError("Completa el apellido.")
        if not data["codigo"]:
            raise ValueError("Completa el código.")
        if not data["sede"]:
            raise ValueError("Selecciona la sede.")
        if data["crd"] <= 0:
            raise ValueError("CRD inválido.")
        if not data["malla"]:
            raise ValueError("No se pudo determinar la malla.")
        if not data["carrera"]:
            raise ValueError("Selecciona la carrera.")
        if not data["unidad"]:
            raise ValueError("Selecciona la unidad.")
        if not data["institucionProcedencia"]:
            raise ValueError("Selecciona la institución de procedencia.")
        if not data["tipoInstitucionProcedencia"]:
            raise ValueError("No se pudo determinar el tipo de institución.")
        if not data["carreraProcedencia"]:
            raise ValueError("Completa la carrera de procedencia.")
        if not data["elaboradoNombre"]:
            raise ValueError("Completa el nombre de elaborado por.")
        if not data["elaboradoCargo"]:
            raise ValueError("Completa el cargo de elaborado por.")
        if not data["respNombre"]:
            raise ValueError("Selecciona el responsable académico.")
        if not data["respCargo"]:
            raise ValueError("No se pudo determinar el cargo del responsable académico.")
        if data["tipoCaso"] not in {"PAQUETE", "REGULAR"}:
            raise ValueError("Selecciona el tipo de caso.")
        if data["tipoCaso"] == "PAQUETE" and not data["tipoPaquete"]:
            raise ValueError("Selecciona el tipo de paquete.")
        return data

    def generar_documentos(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        datos = self.validar_payload(payload)
        folder = self.exporter.create_run_folder(datos["codigo"], datos["alumno"])

        malla_rows = self.repo.get_malla_preview(datos["carrera"], datos["unidad"], datos["malla"])
        if not malla_rows:
            raise ValueError("No se encontró malla para los filtros seleccionados.")

        seleccion = seleccionar_convalidacion(malla_rows, datos["crd"], TOLERANCIA_CRD)
        convalidados = [r for i, r in enumerate(malla_rows) if i in seleccion["seleccion"]]
        resultado = []
        for i, r in enumerate(malla_rows):
            row = dict(r)
            row["ESTADO_CONVALIDACION"] = "CONVALIDADO" if i in seleccion["seleccion"] else "NO CONVALIDADO"
            resultado.append(row)
        matriculables = calcular_matriculables(resultado, convalidados)
        total_convalidados = sum(number_safe(r.get("CR", 0)) for r in convalidados)

        pdf_conva = folder / f"Resultado_Convalidacion_{sanitize_filename(datos['codigo'])}.pdf"
        pdf_proy = folder / f"Proyeccion_Malla_{sanitize_filename(datos['codigo'])}.pdf"
        self.exporter.build_pdf(
            datos, convalidados, total_convalidados, pdf_conva,
            "RESULTADO DE CONVALIDACIÓN", "Relación de cursos convalidados:", "Cursos Convalidados"
        )
        total_matriculables = sum(number_safe(r.get("CR", 0)) for r in matriculables)
        self.exporter.build_pdf(
            datos, matriculables, total_matriculables, pdf_proy,
            "CURSOS RECOMENDADOS PARA EL REGISTRO DE CURSO", "Relación de cursos recomendados:", "Cursos recomendados"
        )
        json_path = self.exporter.save_json(datos, convalidados, matriculables, folder)

        append_log(self.log_path, {
            "FECHA_HORA": format_datetime(datetime.now()),
            "CODIGO": datos["codigo"],
            "ALUMNO": datos["alumno"],
            "SEDE": datos["sede"],
            "CARRERA": datos["carrera"],
            "UNIDAD": datos["unidad"],
            "MALLA": datos["malla"],
            "TIPO_CASO": datos["tipoCaso"],
            "TIPO_PAQUETE": datos["tipoPaquete"],
            "CRD": datos["crd"],
            "INSTITUCION_PROCEDENCIA": datos["institucionProcedencia"],
            "TIPO_INSTITUCION_PROCEDENCIA": datos["tipoInstitucionProcedencia"],
            "CARRERA_PROCEDENCIA": datos["carreraProcedencia"],
            "RESPONSABLE_ACADEMICO": datos["respNombre"],
            "CARGO_RESPONSABLE_ACADEMICO": datos["respCargo"],
            "CREDITOS_CONVALIDADOS": total_convalidados,
            "DOCENTE_CONVALIDA": datos["elaboradoNombre"],
            "PDF_CONVALIDACION": str(pdf_conva),
            "PDF_PROYECCION": str(pdf_proy),
            "JSON_RESUMEN": str(json_path),
        })

        return {
            "resumen": {
                "alumno": datos["alumno"],
                "codigo": datos["codigo"],
                "crdSolicitado": datos["crd"],
                "convalidados": int(total_convalidados),
                "maximoPermitido": int(datos["crd"] + TOLERANCIA_CRD),
                "totalCursosConvalidados": len(convalidados),
                "totalCursosMatriculables": len(matriculables),
            },
            "tablas": {
                "convalidados": convalidados,
                "matriculables": matriculables,
                "malla": malla_rows,
            },
            "archivos": {
                "carpeta": str(folder.resolve()),
                "pdfConvalidacion": str(pdf_conva.resolve()),
                "pdfProyeccion": str(pdf_proy.resolve()),
                "jsonResumen": str(json_path.resolve()),
                "xlsxLog": str(self.log_path.resolve()),
            },
        }


# =========================================================
# UI FLET
# =========================================================

def main(page: ft.Page):
    page.title = APP_TITLE
    page.window_width = 1440
    page.window_height = 920
    page.scroll = ft.ScrollMode.AUTO
    page.theme_mode = ft.ThemeMode.LIGHT
    page.bgcolor = "#F5F7FB"
    page.padding = 18

    try:
        repo = DatasetRepository(DATASET_FILE)
    except Exception as e:
        page.add(
            ft.Container(
                padding=30,
                border_radius=16,
                bgcolor="#FFFFFF",
                content=ft.Column([
                    ft.Text(APP_TITLE, size=24, weight=ft.FontWeight.BOLD),
                    ft.Text(f"Error al cargar dataset.xlsx: {e}", color="#B91C1C", size=16),
                    ft.Text("Coloca el archivo dataset.xlsx en la misma carpeta del script y vuelve a ejecutar.")
                ])
            )
        )
        page.update()
        return

    exporter = ExportService()
    rules = MallaRuleConfig()
    service = ConvalidacionService(repo, exporter, rules)

    paquetes = repo.get_paquetes()
    sedes = repo.get_sedes()
    carreras = repo.get_carreras()
    responsables = repo.get_responsables()

    # Campos
    nombres = ft.TextField(label="Nombres", expand=True)
    apellidos = ft.TextField(label="Apellidos", expand=True)
    codigo = ft.TextField(label="Código del estudiante", width=220)

    sede = ft.Dropdown(label="Sede", width=220, options=[ft.dropdown.Option(x) for x in sedes])
    crd = ft.TextField(label="CRD", width=180, keyboard_type=ft.KeyboardType.NUMBER)
    malla = ft.TextField(label="Malla", width=180, read_only=False)
    carrera = ft.Dropdown(label="Carrera", width=340, options=[ft.dropdown.Option(x) for x in carreras])
    unidad = ft.Dropdown(label="Unidad", width=260, options=[])

    tipo_caso = ft.Dropdown(
        label="Tipo de caso",
        width=180,
        options=[ft.dropdown.Option("PAQUETE"), ft.dropdown.Option("REGULAR")],
    )
    tipo_paquete = ft.Dropdown(label="Tipo de paquete", width=220, visible=False, options=[ft.dropdown.Option(x) for x in paquetes])

    institucion_buscar = ft.TextField(label="Institución de procedencia", expand=True)
    institucion_resultados = ft.ListView(height=140, visible=False, spacing=2)
    tipo_institucion = ft.TextField(label="Procedencia", width=240, read_only=True)
    carrera_procedencia = ft.TextField(label="Nombre de Carrera", expand=True)

    elaborado_nombre = ft.TextField(label="Nombre Elaborado por", width=320, value="Ing. Jesus Apolaya")
    elaborado_cargo = ft.TextField(label="Cargo", width=240, value="DESARROLLADOR")

    resp_nombre = ft.Dropdown(label="Nombre Resp. Acad.", width=340, options=[ft.dropdown.Option(x["nombre"]) for x in responsables])
    resp_cargo = ft.TextField(label="Cargo", width=300, read_only=True)

    # Regla editable
    cad_2025g_min = ft.TextField(label="CAD: mínimo PE2025G", value=str(rules.cad_2025g_min), width=180)
    cad_2023_min = ft.TextField(label="CAD: mínimo PE2023", value=str(rules.cad_2023_min), width=180)
    general_2025g_min = ft.TextField(label="General: mínimo PE2025G", value=str(rules.general_2025g_min), width=200)
    general_2023_min = ft.TextField(label="General: mínimo PE2023", value=str(rules.general_2023_min), width=200)
    regla_texto = ft.Container(
        bgcolor="#ECFDF5",
        border=ft.border.all(1, "#86EFAC"),
        border_radius=12,
        padding=12,
        content=ft.Text("Primero selecciona la sede.", color="#166534", weight=ft.FontWeight.BOLD)
    )

    # Mensajes y KPIs
    msg = ft.Text(color="#1D4ED8")
    kpi_crd = ft.Text("0", size=22, weight=ft.FontWeight.BOLD)
    kpi_conva = ft.Text("0", size=22, weight=ft.FontWeight.BOLD)
    kpi_cursos_conva = ft.Text("0", size=22, weight=ft.FontWeight.BOLD)
    kpi_cursos_matri = ft.Text("0", size=22, weight=ft.FontWeight.BOLD)
    archivos_text = ft.Text(selectable=True)

    # Tablas
    tabla_malla = ft.DataTable(columns=[
        ft.DataColumn(ft.Text("Ciclo")),
        ft.DataColumn(ft.Text("Curso")),
        ft.DataColumn(ft.Text("Materia")),
        ft.DataColumn(ft.Text("Cód. Curso")),
        ft.DataColumn(ft.Text("CR")),
        ft.DataColumn(ft.Text("Requisitos")),
    ], rows=[])

    tabla_conva = ft.DataTable(columns=[
        ft.DataColumn(ft.Text("Ciclo")),
        ft.DataColumn(ft.Text("Curso")),
        ft.DataColumn(ft.Text("Materia")),
        ft.DataColumn(ft.Text("Cód. Curso")),
        ft.DataColumn(ft.Text("CR")),
    ], rows=[])

    tabla_matri = ft.DataTable(columns=[
        ft.DataColumn(ft.Text("Ciclo")),
        ft.DataColumn(ft.Text("Curso")),
        ft.DataColumn(ft.Text("Materia")),
        ft.DataColumn(ft.Text("Cód. Curso")),
        ft.DataColumn(ft.Text("CR")),
    ], rows=[])

    def card(title: str, content: ft.Control):
        return ft.Container(
            bgcolor="#FFFFFF",
            border_radius=18,
            padding=18,
            content=ft.Column([
                ft.Text(title, size=18, weight=ft.FontWeight.BOLD),
                content,
            ], spacing=12),
        )

    def kpi_box(label: str, value_control: ft.Text):
        return ft.Container(
            bgcolor="#F8FAFC",
            border=ft.border.all(1, "#E5E7EB"),
            border_radius=14,
            padding=12,
            width=250,
            content=ft.Column([
                ft.Text(label, color="#6B7280", size=12),
                value_control,
            ], spacing=4),
        )

    def update_rules_from_inputs():
        rules.cad_2025g_min = int(number_safe(cad_2025g_min.value)) or 44
        rules.cad_2023_min = int(number_safe(cad_2023_min.value)) or 65
        rules.general_2025g_min = int(number_safe(general_2025g_min.value)) or 24
        rules.general_2023_min = int(number_safe(general_2023_min.value)) or 65

    def refresh_regla():
        update_rules_from_inputs()
        regla_texto.content = ft.Text(
            service.get_malla_automatica(sede.value, crd.value).get("regla", "Primero selecciona la sede."),
            color="#166534",
            weight=ft.FontWeight.BOLD,
        )

    def refresh_malla_automatica(e=None):
        update_rules_from_inputs()
        info = service.get_malla_automatica(sede.value, crd.value)
        if not malla.focused:
            malla.value = info["mallaReal"]
        refresh_regla()
        refresh_unidades()
        page.update()

    def refresh_unidades(e=None):
        unidades = repo.get_unidades_by_carrera_and_malla(carrera.value or "", malla.value or "")
        unidad.options = [ft.dropdown.Option(x) for x in unidades]
        if unidad.value not in unidades:
            unidad.value = None
        load_malla_preview()

    def load_malla_preview(e=None):
        rows = repo.get_malla_preview(carrera.value or "", unidad.value or "", malla.value or "")
        tabla_malla.rows = [
            ft.DataRow(cells=[
                ft.DataCell(ft.Text(str(r.get("CICLO", "")))),
                ft.DataCell(ft.Text(str(r.get("CURSO", "")))),
                ft.DataCell(ft.Text(str(r.get("MATERIA", "")))),
                ft.DataCell(ft.Text(str(r.get("COD_CURSO", "")))),
                ft.DataCell(ft.Text(str(int(r.get("CR", 0)) if float(r.get("CR", 0)).is_integer() else r.get("CR", 0)))),
                ft.DataCell(ft.Text(str(r.get("REQUISITOS", "")))),
            ])
            for r in rows
        ]
        page.update()

    def render_simple_table(table: ft.DataTable, rows: List[Dict[str, Any]]):
        table.rows = [
            ft.DataRow(cells=[
                ft.DataCell(ft.Text(str(r.get("CICLO", "")))),
                ft.DataCell(ft.Text(str(r.get("CURSO", "")))),
                ft.DataCell(ft.Text(str(r.get("MATERIA", "")))),
                ft.DataCell(ft.Text(str(r.get("COD_CURSO", "")))),
                ft.DataCell(ft.Text(str(int(number_safe(r.get("CR", 0))) if float(number_safe(r.get("CR", 0))).is_integer() else r.get("CR", 0)))),
            ])
            for r in rows
        ]

    def buscar_instituciones(e=None):
        resultados = repo.search_instituciones(institucion_buscar.value or "")
        institucion_resultados.controls = []
        for item in resultados:
            def make_click(it=item):
                def _click(_):
                    institucion_buscar.value = it["nombre"]
                    tipo_institucion.value = it["tipo"]
                    institucion_resultados.visible = False
                    page.update()
                return _click
            institucion_resultados.controls.append(
                ft.Container(
                    padding=8,
                    border=ft.border.only(bottom=ft.BorderSide(1, "#EDF2F7")),
                    content=ft.Column([
                        ft.Text(item["nombre"], weight=ft.FontWeight.BOLD),
                        ft.Text(f"{item['tipo']} {('| ' + item['codigo']) if item['codigo'] else ''}", size=12, color="#6B7280"),
                    ], spacing=2),
                    on_click=make_click(),
                )
            )
        institucion_resultados.visible = len(resultados) > 0
        page.update()

    def on_responsable_change(e=None):
        found = repo.get_responsable_by_nombre(resp_nombre.value or "")
        resp_cargo.value = found["cargo"] if found else ""
        page.update()

    def on_tipo_caso_change(e=None):
        tipo_paquete.visible = (tipo_caso.value == "PAQUETE")
        if tipo_caso.value != "PAQUETE":
            tipo_paquete.value = None
        page.update()

    def limpiar(e=None):
        for ctrl in [nombres, apellidos, codigo, crd, malla, carrera_procedencia, institucion_buscar, tipo_institucion, resp_cargo]:
            ctrl.value = ""
        for ctrl in [sede, carrera, unidad, tipo_caso, tipo_paquete, resp_nombre]:
            ctrl.value = None
        tipo_paquete.visible = False
        elaborado_nombre.value = "Ing. Jesus Apolaya"
        elaborado_cargo.value = "DESARROLLADOR"
        tabla_malla.rows = []
        tabla_conva.rows = []
        tabla_matri.rows = []
        kpi_crd.value = "0"
        kpi_conva.value = "0"
        kpi_cursos_conva.value = "0"
        kpi_cursos_matri.value = "0"
        archivos_text.value = ""
        msg.value = "Formulario limpio."
        msg.color = "#047857"
        refresh_regla()
        page.update()

    def procesar(e=None):
        try:
            msg.value = "Procesando y generando documentos..."
            msg.color = "#1D4ED8"
            page.update()

            payload = {
                "nombres": nombres.value,
                "apellidos": apellidos.value,
                "codigo": codigo.value,
                "sede": sede.value,
                "carrera": carrera.value,
                "unidad": unidad.value,
                "malla": malla.value,
                "tipoCaso": tipo_caso.value,
                "tipoPaquete": tipo_paquete.value,
                "crd": crd.value,
                "elaboradoNombre": elaborado_nombre.value,
                "elaboradoCargo": elaborado_cargo.value,
                "institucionProcedencia": institucion_buscar.value,
                "tipoInstitucionProcedencia": tipo_institucion.value,
                "carreraProcedencia": carrera_procedencia.value,
                "respNombre": resp_nombre.value,
                "respCargo": resp_cargo.value,
            }
            res = service.generar_documentos(payload)

            kpi_crd.value = str(res["resumen"]["crdSolicitado"])
            kpi_conva.value = str(res["resumen"]["convalidados"])
            kpi_cursos_conva.value = str(res["resumen"]["totalCursosConvalidados"])
            kpi_cursos_matri.value = str(res["resumen"]["totalCursosMatriculables"])
            render_simple_table(tabla_conva, res["tablas"]["convalidados"])
            render_simple_table(tabla_matri, res["tablas"]["matriculables"])
            archivos_text.value = (
                f"Carpeta: {res['archivos']['carpeta']}\n"
                f"PDF Convalidación: {res['archivos']['pdfConvalidacion']}\n"
                f"PDF Proyección: {res['archivos']['pdfProyeccion']}\n"
                f"JSON Resumen: {res['archivos']['jsonResumen']}\n"
                f"Excel LOG_APP: {res['archivos']['xlsxLog']}"
            )
            msg.value = "Documentos generados correctamente."
            msg.color = "#047857"
            page.update()
        except Exception as ex:
            msg.value = f"Error: {ex}"
            msg.color = "#B91C1C"
            page.update()

    # Eventos
    sede.on_change = refresh_malla_automatica
    crd.on_change = refresh_malla_automatica
    carrera.on_change = refresh_unidades
    unidad.on_change = load_malla_preview
    resp_nombre.on_change = on_responsable_change
    tipo_caso.on_change = on_tipo_caso_change
    institucion_buscar.on_change = buscar_instituciones

    for rule_field in [cad_2025g_min, cad_2023_min, general_2025g_min, general_2023_min]:
        rule_field.on_change = refresh_malla_automatica

    hero = ft.Container(
        gradient=ft.LinearGradient(colors=["#1F4E79", "#2F6EA5"]),
        border_radius=24,
        padding=24,
        content=ft.Column([
            ft.Text(APP_TITLE, size=28, color="#FFFFFF", weight=ft.FontWeight.BOLD),
            ft.Text(
                "Versión migrada desde Apps Script a Python + Flet, sin login, usando dataset.xlsx y regla de malla editable.",
                size=14,
                color="#FFFFFF",
            ),
        ], spacing=8),
    )

    page.add(
        ft.Container(
            width=RESPONSIVE_WIDTH,
            alignment=ft.alignment.top_center,
            content=ft.Column([
                hero,
                card("Regla de malla editable", ft.Column([
                    ft.Row([cad_2025g_min, cad_2023_min, general_2025g_min, general_2023_min], wrap=True),
                    ft.Text(
                        "La lógica viene precargada con la regla original, pero ahora el usuario puede modificarla antes de procesar.",
                        size=13,
                        color="#6B7280",
                    ),
                    regla_texto,
                ])),
                card("Formulario", ft.Column([
                    ft.Row([nombres, apellidos, codigo], wrap=True),
                    ft.Row([sede, crd, malla, carrera], wrap=True),
                    ft.Row([unidad, tipo_caso, tipo_paquete], wrap=True),
                    ft.Divider(),
                    ft.Text("Carrera origen - Otros Centros de Estudio", weight=ft.FontWeight.BOLD),
                    ft.Row([institucion_buscar, tipo_institucion, carrera_procedencia], wrap=True),
                    institucion_resultados,
                    ft.Divider(),
                    ft.Text("Datos de elaboración", weight=ft.FontWeight.BOLD),
                    ft.Row([elaborado_nombre, elaborado_cargo], wrap=True),
                    ft.Divider(),
                    ft.Text("Responsable académico", weight=ft.FontWeight.BOLD),
                    ft.Row([resp_nombre, resp_cargo], wrap=True),
                    ft.Row([
                        ft.ElevatedButton("Procesar y generar", on_click=procesar, bgcolor="#2F6EA5", color="#FFFFFF"),
                        ft.OutlinedButton("Limpiar", on_click=limpiar),
                    ]),
                    msg,
                ])),
                card("Resumen", ft.Column([
                    ft.Row([
                        kpi_box("CRD solicitado", kpi_crd),
                        kpi_box("Convalidados", kpi_conva),
                        kpi_box("Cursos convalidados", kpi_cursos_conva),
                        kpi_box("Cursos matriculables", kpi_cursos_matri),
                    ], wrap=True),
                    ft.Text("Archivos generados", weight=ft.FontWeight.BOLD),
                    ft.Container(
                        bgcolor="#F8FAFC",
                        border=ft.border.all(1, "#E5E7EB"),
                        border_radius=12,
                        padding=12,
                        content=archivos_text,
                    )
                ])),
                card("Vista previa de malla", ft.Column([
                    ft.Text("Se carga según Carrera + Unidad + Malla", size=12, color="#6B7280"),
                    ft.Row([ft.Container(content=tabla_malla, scroll=ft.ScrollMode.AUTO)], scroll=ft.ScrollMode.AUTO),
                ])),
                card("Cursos convalidados", ft.Row([ft.Container(content=tabla_conva, scroll=ft.ScrollMode.AUTO)], scroll=ft.ScrollMode.AUTO)),
                card("Cursos matriculables", ft.Row([ft.Container(content=tabla_matri, scroll=ft.ScrollMode.AUTO)], scroll=ft.ScrollMode.AUTO)),
                ft.Container(
                    alignment=ft.alignment.center_right,
                    padding=ft.padding.only(top=4, bottom=20),
                    content=ft.Text("Desarrollado: Ing. Jesus Apolaya", size=12, color="#6B7280", italic=True),
                ),
            ], spacing=18),
        )
    )
    refresh_regla()
    page.update()


if __name__ == "__main__":
    ft.app(target=main)
