"""
app.py
Interfaz web para el flujo de cotizaciones Melectra.

Pasos:
  1. Subir .docx + seleccionar .xlsx
  2. Revisar y editar datos extraídos
  3. Insertar y mostrar confirmación

Ejecutar:
  uvicorn app:app --reload
"""

import json
import os
import signal
import sys
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from jinja2 import Environment, FileSystemLoader
from starlette.middleware.base import BaseHTTPMiddleware

from extract_cotizacion import CotizacionExtractor
from insert_cotizacion import XLSX_PATH, insert_cotizacion
from models import DatosCotizacion, Estado, Medio
from xlsx_manager import load_filas, save_filas

# Configuración

# Cuando está congelado por PyInstaller:
#   - sys._MEIPASS  → assets bundleados (templates)
#   - sys.executable parent → junto al .exe, para datos del usuario
if getattr(sys, "frozen", False):
    _ASSETS_DIR = Path(sys._MEIPASS)          # templates bundleados
    _DATA_DIR   = Path(sys.executable).parent  # json, backups
else:
    _ASSETS_DIR = Path(__file__).parent
    _DATA_DIR   = Path(__file__).parent

BASE_DIR   = _ASSETS_DIR
RUTAS_FILE = _DATA_DIR / ".rutas_xlsx.json"

app = FastAPI(title="Melectra Cotizaciones")

# Middleware: solo aceptar peticiones locales (anti-CSRF y anti-DNS-rebinding)

_ALLOWED_HOSTS = {"localhost:8000", "127.0.0.1:8000"}
_ALLOWED_ORIGINS = {"http://localhost:8000", "http://127.0.0.1:8000"}


class LocalhostOnlyMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        host = request.headers.get("host", "")
        if host not in _ALLOWED_HOSTS:
            raise HTTPException(status_code=403, detail="Host no permitido")

        # Para métodos que modifican estado, validar Origin/Referer
        if request.method in ("POST", "PUT", "DELETE", "PATCH"):
            origin = request.headers.get("origin")
            referer = request.headers.get("referer", "")
            if origin:
                if origin not in _ALLOWED_ORIGINS:
                    raise HTTPException(status_code=403, detail="Origen no permitido")
            elif referer:
                if not any(referer.startswith(o) for o in _ALLOWED_ORIGINS):
                    raise HTTPException(status_code=403, detail="Referer no permitido")
            else:
                raise HTTPException(status_code=403, detail="Falta Origin/Referer")

        return await call_next(request)


app.add_middleware(LocalhostOnlyMiddleware)


# Validación de rutas .xlsx

def _validar_xlsx(ruta: str) -> tuple[Path | None, str | None]:
    """Valida que la ruta sea un archivo .xlsx existente.

    Retorna (Path, None) si es válido, (None, mensaje_error) si no.
    """
    if not ruta or not ruta.strip():
        return None, "La ruta del archivo Excel está vacía."
    p = Path(ruta.strip())
    if p.suffix.lower() != ".xlsx":
        return None, f"El archivo debe tener extensión .xlsx (recibido: {p.suffix or 'sin extensión'})."
    if not p.exists():
        return None, f"Archivo no encontrado: {p}"
    if p.is_dir():
        return None, f"La ruta es una carpeta, no un archivo: {p}"
    return p, None


# Jinja2 directo (evita bug con Jinja2Templates en Python 3.14)
_jinja_env = Environment(loader=FileSystemLoader(str(BASE_DIR / "templates")))


# ---------------------------------------------------------------------------
# Gestión de rutas guardadas
# ---------------------------------------------------------------------------

def _load_rutas() -> list[str]:
    """Carga las rutas .xlsx guardadas desde JSON."""
    if RUTAS_FILE.exists():
        try:
            return json.loads(RUTAS_FILE.read_text())
        except (json.JSONDecodeError, OSError):
            pass
    return [str(XLSX_PATH)]


def _save_ruta(ruta: str) -> None:
    """Guarda una ruta .xlsx al historial (sin duplicados, máx 10)."""
    rutas = _load_rutas()
    if ruta in rutas:
        rutas.remove(ruta)
    rutas.insert(0, ruta)
    rutas = rutas[:10]
    RUTAS_FILE.write_text(json.dumps(rutas, indent=2))


def _delete_ruta(ruta: str) -> None:
    """Elimina una ruta del historial."""
    rutas = _load_rutas()
    rutas = [r for r in rutas if r != ruta]
    RUTAS_FILE.write_text(json.dumps(rutas, indent=2))


def _render(template_name: str, context: dict) -> HTMLResponse:
    """Helper para renderizar template Jinja2 con contexto."""
    template = _jinja_env.get_template(template_name)
    html = template.render(**context)
    return HTMLResponse(html)


extractor = CotizacionExtractor()

# Rutas

@app.post("/rutas/eliminar")
async def eliminar_ruta(ruta: str = Form(...)):
    """Elimina una ruta del historial JSON y devuelve la lista actualizada."""
    _delete_ruta(ruta)
    return JSONResponse({"ok": True, "rutas": _load_rutas()})


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Página principal: subir .docx y seleccionar .xlsx."""
    rutas = _load_rutas()
    return _render("index.html", {"request": request, "rutas": rutas})


@app.post("/extraer", response_class=HTMLResponse)
async def extraer(
    request: Request,
    docx_file: UploadFile = File(...),
    xlsx_path_input: str = Form(""),
):
    """Recibe el .docx, extrae datos y muestra preview editable."""

    # Guardar .docx en archivo temporal (necesario para python-docx)
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    tmp.write(await docx_file.read())
    tmp.close()
    docx_path = Path(tmp.name)

    # Resolver .xlsx: si el usuario escribió una ruta, validarla; sino el default
    if xlsx_path_input.strip():
        xlsx_path, err = _validar_xlsx(xlsx_path_input)
        if err:
            docx_path.unlink(missing_ok=True)
            return _render(
                "index.html",
                {"request": request, "rutas": _load_rutas(), "error": err},
            )
    else:
        xlsx_path, err = _validar_xlsx(str(XLSX_PATH))
        if err:
            docx_path.unlink(missing_ok=True)
            return _render(
                "index.html",
                {"request": request, "rutas": _load_rutas(), "error": err},
            )

    _save_ruta(str(xlsx_path))

    # Extraer datos del .docx
    try:
        datos = extractor.extract(docx_path)
    except Exception as e:
        docx_path.unlink(missing_ok=True)
        return _render(
            "index.html",
            {"request": request, "error": f"Error al procesar .docx: {e}"},
        )

    return _render(
        "preview.html",
        {
            "request": request,
            "datos": datos,
            "docx_name": docx_file.filename,
            "docx_tmp": str(docx_path),
            "xlsx_path": str(xlsx_path),
        },
    )


@app.post("/insertar", response_class=HTMLResponse)
async def insertar(
    request: Request,
    docx_tmp: str = Form(...),
    xlsx_path: str = Form(...),
    numero: str = Form(""),
    nombre: str = Form(""),
    empresa: str = Form(""),
    telefono: str = Form(""),
    correo: str = Form(""),
    servicio: str = Form(""),
    valor_total: str = Form(""),
    fecha: str = Form(""),
):
    """Inserta los datos (posiblemente editados) en el .xlsx."""

    # Construir DatosCotizacion desde los campos del form
    datos = DatosCotizacion(
        numero=numero or None,
        nombre=nombre or None,
        empresa=empresa or None,
        telefono=telefono or None,
        correo=correo or None,
        servicio=servicio or None,
        valor_total=valor_total or None,
        fecha=fecha or None,
    )

    # Limpiar archivo temporal del docx
    Path(docx_tmp).unlink(missing_ok=True)

    xlsx, err = _validar_xlsx(xlsx_path)
    if err:
        xlsx_fallback = Path(xlsx_path)
        return _render(
            "result.html",
            {
                "request": request,
                "success": False,
                "duplicate": False,
                "error_msg": err,
                "numero": datos.numero,
                "xlsx_name": xlsx_fallback.name,
                "xlsx_path": str(xlsx_fallback),
            },
        )

    # Intentar insertar directamente en el archivo del usuario
    try:
        inserted = insert_cotizacion(datos, xlsx_path=xlsx)
    except Exception as e:
        return _render(
            "result.html",
            {
                "request": request,
                "success": False,
                "duplicate": False,
                "error_msg": str(e),
                "numero": datos.numero,
                "xlsx_name": xlsx.name,
                "xlsx_path": str(xlsx),
            },
        )

    if inserted:
        _save_ruta(str(xlsx))

    return _render(
        "result.html",
        {
            "request": request,
            "success": inserted,
            "duplicate": not inserted,
            "error_msg": None,
            "numero": datos.numero,
            "xlsx_name": xlsx.name,
            "xlsx_path": str(xlsx),
        },
    )


# Excel manager

def _excel_ctx(xlsx_path: str, filas: list, error=None, success_msg=None) -> dict:
    return {
        "xlsx_path": xlsx_path,
        "rutas": _load_rutas(),
        "filas": filas,
        "medios": [m.value for m in Medio],
        "estados": [e.value for e in Estado],
        "error": error,
        "success_msg": success_msg,
    }


@app.get("/excel", response_class=HTMLResponse)
async def excel_view(request: Request, xlsx: str = ""):
    if not xlsx:
        rutas = _load_rutas()
        xlsx = rutas[0] if rutas else str(XLSX_PATH)

    xlsx_path, error = _validar_xlsx(xlsx)
    filas = []
    if xlsx_path:
        try:
            df = load_filas(xlsx_path)
            filas = df.to_dicts()
        except Exception as e:
            error = str(e)

    return _render("excel.html", {"request": request, **_excel_ctx(xlsx, filas, error=error)})


@app.post("/excel/guardar", response_class=HTMLResponse)
async def excel_guardar(
    request: Request,
    xlsx_path: str = Form(...),
    filas_json: str = Form(...),
):
    xlsx, err = _validar_xlsx(xlsx_path)
    if err:
        return _render("excel.html", {
            "request": request,
            **_excel_ctx(xlsx_path, [], error=err),
        })

    try:
        filas_raw: list[dict] = json.loads(filas_json)
        # Limpiar cadenas vacías a None
        filas = [{k: (v if v else None) for k, v in f.items()} for f in filas_raw]
        backup_path = save_filas(xlsx, filas)
        _save_ruta(xlsx_path)
        msg = f"Guardado correctamente. Backup: {backup_path.name}"
    except Exception as e:
        return _render("excel.html", {
            "request": request,
            **_excel_ctx(xlsx_path, [], error=str(e)),
        })

    # Recargar desde disco para mostrar estado actual
    df = load_filas(xlsx)
    return _render("excel.html", {
        "request": request,
        **_excel_ctx(xlsx_path, df.to_dicts(), success_msg=msg),
    })


# ---------------------------------------------------------------------------
# Shutdown
# ---------------------------------------------------------------------------

@app.post("/shutdown")
async def shutdown():
    """Termina el proceso del servidor liberando el puerto."""
    os.kill(os.getpid(), signal.SIGTERM)
    return JSONResponse({"ok": True})
