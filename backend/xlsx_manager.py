"""
xlsx_manager.py
Carga, edita y guarda el libro de cotizaciones Melectra.

Usa Polars para el manejo en memoria y openpyxl para leer/escribir
conservando estilos y fórmulas en el resto del libro.
"""

from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path
from typing import Optional

import openpyxl
import polars as pl
from openpyxl.worksheet.worksheet import Worksheet

from models import DatosCotizacion

# Nombres de meses en español para detectar la hoja del mes actual
_MESES_ES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE",
]

# ---------------------------------------------------------------------------
# Mapeo campo → columna (1-indexed)
# ---------------------------------------------------------------------------

COL_MAP: dict[str, int] = {
    "medio":                2,   # B — MEDIO POR EL CUAL SE DIO CUENTA
    "numero":               3,   # C — N° COTIZACIÓN
    "empresa":              4,   # D — NOMBRE EMPRESA
    "nombre":               5,   # E — ENCARGADO-SOLICITANTE
    "servicio":             6,   # F — SERVICIO
    "correo":               7,   # G — CORREO
    "telefono":             8,   # H — TELEFONO
    "valor_total":          9,   # I — VALOR
    "estado":               10,  # J — ESTADO
    "trabajo_realizado_en": 11,  # K — TRABAJO REALIZADO EN
    "orden_servicio":       12,  # L — ORDEN DE SERVICIO MELECTRA
    "fecha":                13,  # M — N° FACTURA (fecha de cotización)
    "observacion":          14,  # N — OBSERVACIÓN
}

CAMPOS = list(COL_MAP.keys())

MAX_BACKUPS = 3


# ---------------------------------------------------------------------------
# Helpers internos
# ---------------------------------------------------------------------------

def find_data_sheet(wb: openpyxl.Workbook) -> Worksheet:
    for name in wb.sheetnames:
        if "DESPLEGABLE" not in name.upper():
            return wb[name]
    return wb.active


def find_header_row(ws: Worksheet) -> int:
    """Detecta dinámicamente la fila de encabezados buscando 'MEDIO'."""
    for row in ws.iter_rows(max_row=20):
        for cell in row:
            if cell.value and "MEDIO" in str(cell.value).upper():
                return cell.row
    return 5  # fallback


def list_sheets(xlsx_path: Path) -> list[str]:
    """Retorna los nombres de hojas del libro, excluyendo DESPLEGABLE."""
    wb = openpyxl.load_workbook(xlsx_path, read_only=True)
    return [n for n in wb.sheetnames if "DESPLEGABLE" not in n.upper()]


def find_month_sheet(sheets: list[str]) -> Optional[str]:
    """Busca la hoja que coincida con el mes actual (ej: 'MAYO 2026')."""
    mes_actual = _MESES_ES[datetime.now().month - 1]
    anio_actual = str(datetime.now().year)
    # Primero buscar coincidencia exacta mes+año
    for s in sheets:
        if mes_actual in s.upper() and anio_actual in s:
            return s
    # Luego solo mes
    for s in sheets:
        if mes_actual in s.upper():
            return s
    return sheets[0] if sheets else None


def _cell_str(value) -> Optional[str]:
    if value is None:
        return None
    s = str(value).strip()
    return s if s else None


# ---------------------------------------------------------------------------
# Backup
# ---------------------------------------------------------------------------

def create_backup(xlsx_path: Path) -> Path:
    """
    Copia el xlsx en <xlsx_dir>/backups/<stem>_<timestamp>.xlsx.
    Guardarlo junto al archivo fuente garantiza una ruta escribible y fácil
    de encontrar para el usuario, tanto en dev como dentro de un AppImage.
    Conserva solo los MAX_BACKUPS más recientes por cada archivo.
    """
    backup_dir = xlsx_path.parent / "backups"
    backup_dir.mkdir(parents=True, exist_ok=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = backup_dir / f"{xlsx_path.stem}_{ts}{xlsx_path.suffix}"
    shutil.copy2(xlsx_path, dest)

    pattern = f"{xlsx_path.stem}_*{xlsx_path.suffix}"
    backups = sorted(backup_dir.glob(pattern))
    for old in backups[:-MAX_BACKUPS]:
        old.unlink(missing_ok=True)

    return dest


# ---------------------------------------------------------------------------
# Carga
# ---------------------------------------------------------------------------

def load_filas(xlsx_path: Path) -> pl.DataFrame:
    """
    Lee las filas de datos del xlsx y retorna un DataFrame de Polars.
    Todos los campos son Utf8 (o null). Incluye '_row' con el nro. de fila.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = find_data_sheet(wb)
    header_row = find_header_row(ws)
    data_start = header_row + 1

    records: list[dict] = []
    for row in ws.iter_rows(min_row=data_start, values_only=False):
        vals = {f: _cell_str(row[col - 1].value) for f, col in COL_MAP.items()}
        if not any(v for v in vals.values()):
            continue
        vals["_row"] = row[0].row
        records.append(vals)

    schema = {f: pl.Utf8 for f in CAMPOS}
    schema["_row"] = pl.Int32

    if not records:
        return pl.DataFrame(schema=schema)

    return pl.DataFrame(records, infer_schema_length=None).cast(schema)


# ---------------------------------------------------------------------------
# Guardado
# ---------------------------------------------------------------------------

def save_filas(xlsx_path: Path, filas: list[dict]) -> Path:
    """
    Escribe la lista de dicts al xlsx preservando estilos.
    Crea backup antes de guardar y retorna la ruta del backup.
    """
    backup_path = create_backup(xlsx_path)

    wb = openpyxl.load_workbook(xlsx_path)
    ws = find_data_sheet(wb)
    header_row = find_header_row(ws)
    data_start = header_row + 1

    # Borrar filas de datos existentes
    if ws.max_row >= data_start:
        ws.delete_rows(data_start, ws.max_row - data_start + 1)

    # Escribir nuevas filas
    for i, fila in enumerate(filas):
        row_idx = data_start + i
        for campo, col in COL_MAP.items():
            val = fila.get(campo)
            ws.cell(row=row_idx, column=col, value=val if val else None)

    wb.save(xlsx_path)
    return backup_path


# ---------------------------------------------------------------------------
# Conversión DatosCotizacion ↔ dict
# ---------------------------------------------------------------------------

def datos_to_dict(datos: DatosCotizacion) -> dict:
    return {f: getattr(datos, f, None) for f in CAMPOS}


def dict_to_datos(d: dict) -> DatosCotizacion:
    return DatosCotizacion(**{f: d.get(f) for f in CAMPOS})
