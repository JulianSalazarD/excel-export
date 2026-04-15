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
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from jinja2 import Environment, FileSystemLoader

from extract_cotizacion import CotizacionExtractor
from insert_cotizacion import XLSX_PATH, insert_cotizacion
from models import DatosCotizacion, Estado, Medio
from xlsx_manager import load_filas, save_filas

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent
RUTAS_FILE = BASE_DIR / ".rutas_xlsx.json"

app = FastAPI(title="Melectra Cotizaciones")

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


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------


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

    # Resolver .xlsx: si el usuario escribió una ruta, usarla; sino el default
    if xlsx_path_input.strip():
        xlsx_path = Path(xlsx_path_input.strip())
    else:
        xlsx_path = XLSX_PATH

    # Verificar que el xlsx existe
    if not xlsx_path.exists():
        docx_path.unlink(missing_ok=True)
        return _render(
            "index.html",
            {
                "request": request,
                "rutas": _load_rutas(),
                "error": f"Archivo Excel no encontrado: {xlsx_path}",
            },
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

    xlsx = Path(xlsx_path)

    # Limpiar archivo temporal del docx
    Path(docx_tmp).unlink(missing_ok=True)

    if not xlsx.exists():
        return _render(
            "result.html",
            {
                "request": request,
                "success": False,
                "duplicate": False,
                "error_msg": f"Archivo Excel no encontrado: {xlsx}",
                "numero": datos.numero,
                "xlsx_name": xlsx.name,
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
        },
    )


# ---------------------------------------------------------------------------
# Excel manager
# ---------------------------------------------------------------------------

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

    filas, error = [], None
    xlsx_path = Path(xlsx)
    if xlsx_path.exists():
        try:
            df = load_filas(xlsx_path)
            filas = df.to_dicts()
        except Exception as e:
            error = str(e)
    else:
        error = f"Archivo no encontrado: {xlsx}"

    return _render("excel.html", {"request": request, **_excel_ctx(xlsx, filas, error=error)})


@app.post("/excel/guardar", response_class=HTMLResponse)
async def excel_guardar(
    request: Request,
    xlsx_path: str = Form(...),
    filas_json: str = Form(...),
):
    xlsx = Path(xlsx_path)
    if not xlsx.exists():
        return _render("excel.html", {
            "request": request,
            **_excel_ctx(xlsx_path, [], error=f"Archivo no encontrado: {xlsx_path}"),
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
