from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Optional


class Medio(str, Enum):
    NUEVO      = "NUEVO"
    REFERIDO   = "REFERIDO"
    CLIENTE    = "CLIENTE"
    WSP        = "WSP"
    LINEA      = "LÍINEA"
    PAGINA     = "PÁGINA"
    CORREO     = "CORREO"
    INSTAGRAM  = "INSTAGRAM"


class Estado(str, Enum):
    RECIBIDA    = "RECIBIDA"
    ACTUALIZADA = "ACTUALIZADA"
    RECHAZADA   = "RECHAZADA"
    APROBADA    = "APROBADA"
    FACTURADA   = "FACTURADA"


@dataclass
class DatosCotizacion:
    # ── Campos extraídos del .docx ──────────────────────────────────────
    numero:               Optional[str] = None
    nombre:               Optional[str] = None
    empresa:              Optional[str] = None
    telefono:             Optional[str] = None
    correo:               Optional[str] = None
    servicio:             Optional[str] = None
    valor_total:          Optional[str] = None
    fecha:                Optional[str] = None
    # ── Campos adicionales del Excel ────────────────────────────────────
    medio:                Optional[str] = None   # col B — enum Medio
    estado:               Optional[str] = None   # col J — enum Estado
    trabajo_realizado_en: Optional[str] = None   # col K
    orden_servicio:       Optional[str] = None   # col L
    observacion:          Optional[str] = None   # col N

    def __str__(self) -> str:
        lines = [
            f"Número             : {self.numero}",
            f"Nombre             : {self.nombre}",
            f"Empresa            : {self.empresa}",
            f"Teléfono           : {self.telefono}",
            f"Correo             : {self.correo}",
            f"Servicio           : {self.servicio}",
            f"Valor total        : {self.valor_total}",
            f"Fecha              : {self.fecha}",
            f"Medio              : {self.medio}",
            f"Estado             : {self.estado}",
            f"Trabajo realizado  : {self.trabajo_realizado_en}",
            f"Orden de servicio  : {self.orden_servicio}",
            f"Observación        : {self.observacion}",
        ]
        return "\n".join(lines)
