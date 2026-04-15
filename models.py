from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class DatosCotizacion:
    numero: Optional[str] = None
    nombre: Optional[str] = None
    empresa: Optional[str] = None
    telefono: Optional[str] = None
    correo: Optional[str] = None
    servicio: Optional[str] = None
    valor_total: Optional[str] = None
    fecha: Optional[str] = None

    def __str__(self) -> str:
        lines = [
            f"Número      : {self.numero}",
            f"Nombre      : {self.nombre}",
            f"Empresa     : {self.empresa}",
            f"Teléfono    : {self.telefono}",
            f"Correo      : {self.correo}",
            f"Servicio    : {self.servicio}",
            f"Valor total : {self.valor_total}",
            f"Fecha       : {self.fecha}",
        ]
        return "\n".join(lines)
