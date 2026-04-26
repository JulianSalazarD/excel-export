"""
Wrapper para insert_cotizacion.py que usa JSON para la comunicación.

Uso:
    python insert_wrapper.py '<datos_json>' [xlsx_path] [sheet_name]
    # Salida: JSON con resultado
"""

import json
import sys
from pathlib import Path

# Añadir el directorio padre al path para importar los módulos
sys.path.insert(0, str(Path(__file__).parent.parent))

from insert_cotizacion import insert_cotizacion, XLSX_PATH
from models import DatosCotizacion


def main() -> None:
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Uso: insert_wrapper.py '<datos_json>' [xlsx_path] [sheet_name]"}))
        sys.exit(1)

    try:
        # Parsear JSON de argumento
        datos_json = sys.argv[1]
        datos_dict = json.loads(datos_json)

        # Crear objeto DatosCotizacion
        datos = DatosCotizacion(
            medio=datos_dict.get("medio"),
            numero=datos_dict.get("numero"),
            nombre=datos_dict.get("nombre"),
            empresa=datos_dict.get("empresa"),
            telefono=datos_dict.get("telefono"),
            correo=datos_dict.get("correo"),
            servicio=datos_dict.get("servicio"),
            valor_total=datos_dict.get("valor_total"),
            estado=datos_dict.get("estado"),
            trabajo_realizado_en=datos_dict.get("trabajo_realizado_en"),
            orden_servicio=datos_dict.get("orden_servicio"),
            fecha=datos_dict.get("fecha"),
            observacion=datos_dict.get("observacion"),
        )

        # Determinar xlsx_path
        xlsx_path = XLSX_PATH
        if len(sys.argv) >= 3:
            xlsx_path = Path(sys.argv[2])

        # Determinar sheet_name
        sheet_name = None
        if len(sys.argv) >= 4:
            sheet_name = sys.argv[3]

        # Insertar
        resultado = insert_cotizacion(datos, xlsx_path=xlsx_path, sheet_name=sheet_name)

        # Solo imprimir JSON por stdout (sin líneas adicionales)
        print(json.dumps({"insertado": resultado}))
        sys.exit(0)

    except json.JSONDecodeError as e:
        print(json.dumps({"error": f"JSON inválido: {e}"}))
        sys.exit(1)
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
