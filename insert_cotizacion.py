"""
insert_cotizacion.py
Inserta un DatosCotizacion en el xlsx de control de cotizaciones Melectra.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from extract_cotizacion import _parse_raw_valor
from models import DatosCotizacion
from xlsx_manager import COL_MAP, find_data_sheet, find_header_row

XLSX_PATH      = Path("docs/COTIZACIONES 2026. - copia.xlsx")
DEFAULT_ESTADO = "RECIBIDA"


def _parse_valor(raw: Optional[str]) -> Optional[float]:
    return _parse_raw_valor(raw) if raw else None


def _existing_numeros(ws: Worksheet, data_start: int) -> set[str]:
    col = COL_MAP["numero"] - 1
    numeros: set[str] = set()
    for row in ws.iter_rows(min_row=data_start, values_only=True):
        val = row[col]
        if val:
            numeros.add(str(val).strip())
    return numeros


def insert_row(ws: Worksheet, datos: DatosCotizacion, data_start: int) -> None:
    next_row = data_start
    for row in ws.iter_rows(min_row=data_start):
        if all(cell.value is None for cell in row):
            break
        next_row += 1

    ws.cell(row=next_row, column=COL_MAP["medio"],    value=datos.medio or "")
    ws.cell(row=next_row, column=COL_MAP["numero"],   value=datos.numero)
    ws.cell(row=next_row, column=COL_MAP["empresa"],  value=datos.empresa)
    ws.cell(row=next_row, column=COL_MAP["nombre"],   value=datos.nombre)
    ws.cell(row=next_row, column=COL_MAP["servicio"], value=datos.servicio)
    ws.cell(row=next_row, column=COL_MAP["correo"],   value=datos.correo)
    ws.cell(row=next_row, column=COL_MAP["telefono"], value=datos.telefono)
    ws.cell(row=next_row, column=COL_MAP["valor_total"], value=_parse_valor(datos.valor_total))
    ws.cell(row=next_row, column=COL_MAP["estado"],   value=datos.estado or DEFAULT_ESTADO)
    ws.cell(row=next_row, column=COL_MAP["fecha"],    value=datos.fecha)


def insert_cotizacion(
    datos: DatosCotizacion,
    xlsx_path: Path = XLSX_PATH,
    skip_duplicates: bool = True,
) -> bool:
    wb = load_workbook(xlsx_path)
    ws = find_data_sheet(wb)
    header_row = find_header_row(ws)
    data_start = header_row + 1

    if skip_duplicates and datos.numero:
        if datos.numero in _existing_numeros(ws, data_start):
            print(f"  [OMITIDO] {datos.numero} ya existe en la hoja.")
            return False

    insert_row(ws, datos, data_start)
    wb.save(xlsx_path)
    return True
