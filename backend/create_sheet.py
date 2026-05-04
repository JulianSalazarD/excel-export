"""
Wrapper para crear una nueva hoja con plantilla desde un Excel existente.

Uso:
    python create_sheet.py <xlsx_path> <sheet_name>
    # Salida: JSON con nombre de la hoja creada o error
"""

import json
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")


def main() -> None:
    if len(sys.argv) < 3:
        print(json.dumps({"error": "Uso: create_sheet.py <xlsx_path> <sheet_name>"}))
        sys.exit(1)

    xlsx_path = Path(sys.argv[1])
    sheet_name = sys.argv[2]

    if not xlsx_path.exists():
        print(json.dumps({"error": f"Archivo no encontrado: {xlsx_path}"}))
        sys.exit(1)

    try:
        from xlsx_manager import create_template_sheet

        created = create_template_sheet(xlsx_path, sheet_name)
        print(json.dumps({"sheet": created}))
        sys.exit(0)

    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
