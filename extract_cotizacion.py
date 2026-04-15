"""
extract_cotizacion.py
Extrae datos clave de cotizaciones Melectra (.docx).

Campos extraídos:
  - numero        → cuerpo del doc (sin espacios) y/o nombre de archivo, se comparan
  - nombre        → debajo de "Señor" / "Señora"
  - empresa       → debajo del nombre, saltando cargo si lo hay; comparado con filename
  - telefono      → después de "Móvil:" / "Movil:" / "Cel:"
  - correo        → después de "E-mail:"
  - servicio      → después de "ASUNTO:" (párrafos continuos hasta línea vacía)
  - valor_total   → valor máximo en la columna "VALOR TOTAL ANTES DE IVA"
  - fecha         → fecha al inicio del documento, convertida a DD/MM/YYYY
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from docx import Document
from docx.table import Table

from models import DatosCotizacion


# ---------------------------------------------------------------------------
# Utilidades de fecha
# ---------------------------------------------------------------------------

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


# ---------------------------------------------------------------------------
# Regex patterns
# ---------------------------------------------------------------------------

RE_NUMERO    = re.compile(r"COTIZACI[OÓ]N\s+No\.?\s*(.+)", re.IGNORECASE)
RE_SENIOR    = re.compile(r"^Se[ñn]or[a]?\s*$", re.IGNORECASE)
RE_MOVIL     = re.compile(r"(?:M[oó]vil|Movil|Cel)\s*:\s*(.+)", re.IGNORECASE)
RE_EMAIL     = re.compile(r"E-?mail\s*:\s*(.+)", re.IGNORECASE)
RE_ASUNTO    = re.compile(r"ASUNTO\s*:\s*(.+)", re.IGNORECASE)
RE_FECHA     = re.compile(r"\d{1,2}\s+de\s+\w+\s+del?\s+\d{4}", re.IGNORECASE)
RE_VALOR_HDR = re.compile(r"VALOR\s+TOTAL[,]?\s+ANTES\s+(?:DE(?:L)?\s+)?IVA", re.IGNORECASE)

# Extrae el número de cotización del nombre del archivo: "COT 040401-26SV-W ..."
RE_NUM_FILE  = re.compile(r"COT\s+([\w-]+)", re.IGNORECASE)

# Extrae la empresa del nombre del archivo: palabras entre el número y la
# primera palabra clave que inicia la descripción del servicio
RE_EMP_FILE  = re.compile(
    r"COT\s+[\w-]+\s+(.+?)"
    r"(?=\s+(?:PROYECTO|PROPUESTA|PRUEBA[S]?|PARA\b|CORREC|DIAGNÓS|"
    r"MANTENIMIENTO|DISEÑO|INSTALA|VALIDAC|VLF\b)|\s+-\s|\.\w+$|$)",
    re.IGNORECASE,
)

# Patrones que indican que una línea es información de contacto (no empresa)
RE_CONTACTO  = re.compile(r"M[oó]vil|Movil|Cel|E-?mail|ASUNTO", re.IGNORECASE)


# ---------------------------------------------------------------------------
# Helpers de valor
# ---------------------------------------------------------------------------

def _parse_raw_valor(raw: str) -> Optional[float]:
    """'$ 9.800.000' | '$9,800,000' | '1.700.000' → float"""
    limpio = re.sub(r"[\$\s]", "", raw)   # quita $ y espacios
    # Si hay coma Y punto: coma es miles → quitarla, punto es decimal
    # Si sólo hay punto(s): todos son miles (formato colombiano)
    if "," in limpio and "." in limpio:
        limpio = limpio.replace(",", "")
    elif "," in limpio:
        # podría ser decimal europeo o miles; si hay más de 3 dígitos tras coma → miles
        partes = limpio.split(",")
        limpio = limpio.replace(",", "") if len(partes[-1]) == 3 else limpio.replace(",", ".")
    else:
        # sólo puntos → todos son separadores de miles
        limpio = limpio.replace(".", "")
    try:
        return float(limpio)
    except ValueError:
        return None


def _format_valor(n: float) -> str:
    """9800000 → '$9,800,000'"""
    return f"${n:,.0f}"


# ---------------------------------------------------------------------------
# Extractor modular
# ---------------------------------------------------------------------------

class CotizacionExtractor:
    """
    Extrae campos de un documento .docx de cotización Melectra.

    Cada método `_extract_*` es reemplazable por una implementación LLM.
    """

    def extract(self, path: str | Path) -> DatosCotizacion:
        path = Path(path)
        doc = Document(str(path))
        paragraphs = [p.text.strip() for p in doc.paragraphs]
        tables = doc.tables

        datos = DatosCotizacion()
        datos.numero      = self._extract_numero(paragraphs, path)
        datos.nombre      = self._extract_nombre(paragraphs)
        datos.empresa     = self._extract_empresa(paragraphs, path)
        datos.telefono    = self._extract_telefono(paragraphs)
        datos.correo      = self._extract_correo(paragraphs)
        datos.servicio    = self._extract_servicio(paragraphs)
        datos.valor_total = self._extract_valor_total(tables)
        datos.fecha       = self._extract_fecha(paragraphs)
        return datos

    # ------------------------------------------------------------------
    # Helpers de filename
    # ------------------------------------------------------------------

    def _numero_from_filename(self, path: Path) -> Optional[str]:
        m = RE_NUM_FILE.search(path.stem)
        if m:
            return re.sub(r"\s+", "", m.group(1))
        return None

    def _empresa_from_filename(self, path: Path) -> Optional[str]:
        m = RE_EMP_FILE.search(path.stem)
        if m:
            return m.group(1).strip()
        return None

    # ------------------------------------------------------------------
    # Campos individuales
    # ------------------------------------------------------------------

    def _extract_numero(self, paragraphs: list[str], path: Optional[Path] = None) -> Optional[str]:
        """
        Busca en el cuerpo del documento y en el nombre del archivo.
        Elimina espacios del resultado para que quede todo junto.
        """
        body = None
        for text in paragraphs:
            m = RE_NUMERO.search(text)
            if m:
                body = re.sub(r"\s+", "", m.group(1).strip())
                break

        filename = self._numero_from_filename(path) if path else None

        # Preferir el del cuerpo; si no hay, usar el del nombre de archivo
        return body or filename

    def _extract_nombre(self, paragraphs: list[str]) -> Optional[str]:
        """Línea inmediatamente después de 'Señor' / 'Señora'."""
        for i, text in enumerate(paragraphs):
            if RE_SENIOR.match(text) and i + 1 < len(paragraphs):
                candidate = paragraphs[i + 1].strip()
                if candidate:
                    return candidate
        return None

    def _extract_empresa(self, paragraphs: list[str], path: Optional[Path] = None) -> Optional[str]:
        """
        Busca la empresa debajo de Señor/Señora, saltando el cargo si lo hay.
        Estrategia: acumula líneas no-vacías y no-contacto tras el nombre;
        usa la última (la empresa suele aparecer después del cargo).
        Compara con el nombre del archivo como referencia.
        """
        empresa_file = self._empresa_from_filename(path) if path else None

        for i, text in enumerate(paragraphs):
            if not RE_SENIOR.match(text):
                continue

            # Recopilar candidatos: líneas no vacías que no sean contacto
            candidatos: list[str] = []
            for j in range(i + 2, min(i + 6, len(paragraphs))):
                linea = paragraphs[j].strip()
                if not linea:
                    break
                if RE_CONTACTO.search(linea):
                    break
                candidatos.append(linea)

            if not candidatos:
                continue

            # Si hay referencia del filename, preferir la que coincida
            if empresa_file:
                ef = empresa_file.upper()
                for c in candidatos:
                    if ef in c.upper():
                        return c

            # Sin referencia: la última candidata es la empresa
            # (cargo siempre viene antes que la empresa)
            return candidatos[-1]

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
        Captura todo el texto del párrafo ASUNTO y sus continuaciones
        hasta encontrar una línea vacía.
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
        Recorre todas las tablas buscando la columna VALOR TOTAL ANTES DE IVA.
        Devuelve el valor más alto encontrado formateado como '$X,XXX,XXX'.
        """
        max_val: Optional[float] = None

        for table in tables:
            if not table.rows:
                continue
            header_cells = [c.text.strip() for c in table.rows[0].cells]
            col_idx = next(
                (i for i, h in enumerate(header_cells) if RE_VALOR_HDR.search(h)),
                None,
            )
            if col_idx is None:
                continue

            for row in table.rows[1:]:
                if col_idx >= len(row.cells):
                    continue
                raw = row.cells[col_idx].text.strip()
                if not raw:
                    continue
                n = _parse_raw_valor(raw)
                if n is not None and (max_val is None or n > max_val):
                    max_val = n

        return _format_valor(max_val) if max_val is not None else None
