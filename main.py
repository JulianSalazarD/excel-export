"""
main.py
Flujo principal: extrae datos de cotizaciones .docx e inserta en el xlsx.

Uso:
    python main.py <archivo.docx> [archivo2.docx ...]
"""

import sys
from pathlib import Path

from extract_cotizacion import CotizacionExtractor
from insert_cotizacion import XLSX_PATH, insert_cotizacion


def main() -> None:
    if len(sys.argv) < 2:
        print("Uso: python main.py <archivo.docx> [archivo2.docx ...]")
        sys.exit(1)

    extractor = CotizacionExtractor()

    for arg in sys.argv[1:]:
        path = Path(arg)
        if not path.exists():
            print(f"[ERROR] Archivo no encontrado: {path}")
            continue

        print(f"\nProcesando: {path.name}")
        datos = extractor.extract(path)
        print(datos)

        insertado = insert_cotizacion(datos)
        if insertado:
            print(f"  [OK] Insertado en {XLSX_PATH.name}")


if __name__ == "__main__":
    main()
