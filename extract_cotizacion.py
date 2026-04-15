"""
extract_cotizacion.py
Extrae datos clave de cotizaciones Melectra (.docx).

Campos extraídos:
  - numero        → después de "COTIZACIÓN No."
  - nombre        → debajo de "Señor" / "Señora"
  - empresa       → debajo del nombre
  - telefono      → después de "Móvil:" / "Movil:" / "Cel:"
  - correo        → después de "E-mail:"
  - servicio      → después de "ASUNTO:"
  - valor_total   → columna "VALOR TOTAL ANTES DE IVA" en la última tabla que la contenga
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from docx import Document
from docx.table import Table

from models import DatosCotizacion

# Mapeo de meses en español a número
MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12,
}

RE_FECHA_NUM = re.compile(r"(\d{1,2})\s+de\s+(\w+)\s+del?\s+(\d{4})", re.IGNORECASE)


def fecha_a_ddmmyyyy(texto: str | None) -> str | None:
    """Convierte '4 de abril del 2026' a '04/04/2026'."""
    if not texto:
        return None
    m = RE_FECHA_NUM.search(texto)
    if not m:
        return texto
    dia = int(m.group(1))
    mes_nombre = m.group(2).lower()
    anio = int(m.group(3))
    mes = MESES_ES.get(mes_nombre)
    if mes is None:
        return texto
    return f"{dia:02d}/{mes:02d}/{anio}"



# Regex patterns

RE_NUMERO   = re.compile(r"COTIZACI[OÓ]N\s+No\.?\s*(.+)", re.IGNORECASE)
RE_SENIOR   = re.compile(r"^Se[ñn]or[a]?\s*$", re.IGNORECASE)
RE_MOVIL    = re.compile(r"(?:M[oó]vil|Movil|Cel)\s*:\s*(.+)", re.IGNORECASE)
RE_EMAIL    = re.compile(r"E-?mail\s*:\s*(.+)", re.IGNORECASE)
RE_ASUNTO   = re.compile(r"ASUNTO\s*:\s*(.+)", re.IGNORECASE)
RE_FECHA    = re.compile(r"\d{1,2}\s+de\s+\w+\s+del?\s+\d{4}", re.IGNORECASE)
RE_VALOR_HDR = re.compile(r"VALOR\s+TOTAL[,]?\s+ANTES\s+(?:DE(?:L)?\s+)?IVA", re.IGNORECASE)


# Extractor modular — cada método se puede reemplazar por una llamada a AI

class CotizacionExtractor:
    """
    Extrae campos de un documento .docx de cotización Melectra.

    Cada método `_extract_*` recibe los textos de párrafos ya normalizados
    (lista de str) y/o las tablas del documento.  Si en algún momento el
    regex no alcanza, basta con sobreescribir el método correspondiente
    con una implementación basada en LLM.
    """

    def extract(self, path: str | Path) -> DatosCotizacion:
        doc = Document(str(path))
        paragraphs = [p.text.strip() for p in doc.paragraphs]
        tables = doc.tables

        datos = DatosCotizacion()
        datos.numero      = self._extract_numero(paragraphs)
        datos.nombre      = self._extract_nombre(paragraphs)
        datos.empresa     = self._extract_empresa(paragraphs)
        datos.telefono    = self._extract_telefono(paragraphs)
        datos.correo      = self._extract_correo(paragraphs)
        datos.servicio    = self._extract_servicio(paragraphs)
        datos.valor_total = self._extract_valor_total(tables)
        datos.fecha       = self._extract_fecha(paragraphs)
        return datos

    # Campos individuales

    def _extract_numero(self, paragraphs: list[str]) -> Optional[str]:
        for text in paragraphs:
            m = RE_NUMERO.search(text)
            if m:
                return m.group(1).strip()
        return None

    def _extract_nombre(self, paragraphs: list[str]) -> Optional[str]:
        """Línea inmediatamente después de 'Señor' / 'Señora'."""
        for i, text in enumerate(paragraphs):
            if RE_SENIOR.match(text) and i + 1 < len(paragraphs):
                candidate = paragraphs[i + 1].strip()
                if candidate:
                    return candidate
        return None

    def _extract_empresa(self, paragraphs: list[str]) -> Optional[str]:
        """Línea inmediatamente después del nombre (que está después de Señor/Señora)."""
        for i, text in enumerate(paragraphs):
            if RE_SENIOR.match(text) and i + 2 < len(paragraphs):
                candidate = paragraphs[i + 2].strip()
                if candidate:
                    return candidate
        return None

    def _extract_telefono(self, paragraphs: list[str]) -> Optional[str]:
        for text in paragraphs:
            m = RE_MOVIL.search(text)
            if m:
                return m.group(1).strip()
        return None

    def _extract_correo(self, paragraphs: list[str]) -> Optional[str]:
        for text in paragraphs:
            m = RE_EMAIL.search(text)
            if m:
                return m.group(1).strip()
        return None

    def _extract_servicio(self, paragraphs: list[str]) -> Optional[str]:
        """
        Captura todo el texto del párrafo ASUNTO (puede continuar en párrafos
        siguientes hasta encontrar una línea vacía).
        """
        for i, text in enumerate(paragraphs):
            m = RE_ASUNTO.search(text)
            if m:
                partes = [m.group(1).strip()]
                for siguiente in paragraphs[i + 1:]:
                    if not siguiente:
                        break
                    partes.append(siguiente)
                return " ".join(p for p in partes if p)
        return None

    def _extract_fecha(self, paragraphs: list[str]) -> Optional[str]:
        """Extrae la fecha y la convierte a DD/MM/YYYY."""
        for text in paragraphs:
            m = RE_FECHA.search(text)
            if m:
                return fecha_a_ddmmyyyy(m.group(0).strip())
        return None

    def _extract_valor_total(self, tables: list[Table]) -> Optional[str]:
        """
        Busca en todas las tablas la columna cuyo encabezado coincide con
        VALOR TOTAL ANTES DE(L) IVA y retorna el valor de la última tabla
        que la contenga (la más cercana al final del documento).
        """
        result = None
        for table in tables:
            if not table.rows:
                continue
            header_row = table.rows[0]
            header_cells = [c.text.strip() for c in header_row.cells]

            col_idx = next(
                (i for i, h in enumerate(header_cells) if RE_VALOR_HDR.search(h)),
                None,
            )
            if col_idx is None:
                continue

            # Tomar el valor de la primera fila de datos
            for row in table.rows[1:]:
                cells = row.cells
                if col_idx < len(cells):
                    val = cells[col_idx].text.strip()
                    if val:
                        result = val
                        break

        return result


