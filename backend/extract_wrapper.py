"""
Wrapper para extract_cotizacion.py que usa JSON para la comunicación.

Uso:
    python extract_wrapper.py <archivo.docx>
    # Salida: JSON con datos extraídos o error
"""

import json
import sys
from pathlib import Path

from extract_cotizacion import CotizacionExtractor


def main() -> None:
    if len(sys.argv) < 2:
        print(json.dumps({"error": "Uso: extract_wrapper.py <archivo.docx>"}))
        sys.exit(1)

    docx_path = Path(sys.argv[1])

    if not docx_path.exists():
        print(json.dumps({"error": f"Archivo no encontrado: {docx_path}"}))
        sys.exit(1)

    try:
        extractor = CotizacionExtractor()
        datos = extractor.extract(docx_path)
        
        # Convertir a diccionario
        resultado = {
            "numero": datos.numero,
            "nombre": datos.nombre,
            "empresa": datos.empresa,
            "telefono": datos.telefono,
            "correo": datos.correo,
            "servicio": datos.servicio,
            "valor_total": datos.valor_total,
            "estado": datos.estado,
            "trabajo_realizado_en": datos.trabajo_realizado_en,
            "orden_servicio": datos.orden_servicio,
            "fecha": datos.fecha,
            "observacion": datos.observacion,
        }
        
        print(json.dumps(resultado, ensure_ascii=False))
        sys.exit(0)
        
    except Exception as e:
        print(json.dumps({"error": str(e)}))
        sys.exit(1)


if __name__ == "__main__":
    main()
