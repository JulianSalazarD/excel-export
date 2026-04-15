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

from pathlib import Path

from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse
from jinja2 import Environment, FileSystemLoader

from extract_cotizacion import CotizacionExtractor
from insert_cotizacion import XLSX_PATH, insert_cotizacion
from models import DatosCotizacion

# ---------------------------------------------------------------------------
# Configuración
# ---------------------------------------------------------------------------

BASE_DIR = Path(__file__).parent
UPLOADS_DIR = BASE_DIR / "uploads"
UPLOADS_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Melectra Cotizaciones")

# Jinja2 directo (evita bug con Jinja2Templates en Python 3.14)
_jinja_env = Environment(loader=FileSystemLoader(str(BASE_DIR / "templates")))


def _render(template_name: str, context: dict) -> HTMLResponse:
    """Helper para renderizar template Jinja2 con contexto."""
    template = _jinja_env.get_template(template_name)
    html = template.render(**context)
    return HTMLResponse(html)


extractor = CotizacionExtractor()


# ---------------------------------------------------------------------------
# Rutas
# ---------------------------------------------------------------------------


@app.get("/", response_class=HTMLResponse)
async def index(request: Request):
    """Página principal: subir .docx y seleccionar .xlsx."""
    return _render("index.html", {"request": request})


@app.post("/extraer", response_class=HTMLResponse)
async def extraer(
    request: Request,
    docx_file: UploadFile = File(...),
    xlsx_file: UploadFile = File(None),
):
    """Recibe el .docx, extrae datos y muestra preview editable."""

    # Guardar .docx temporalmente
    docx_path = UPLOADS_DIR / "cotizacion.docx"
    docx_path.write_bytes(await docx_file.read())

    # Resolver .xlsx
    if xlsx_file and xlsx_file.filename:
        xlsx_path = UPLOADS_DIR / "destino.xlsx"
        xlsx_path.write_bytes(await xlsx_file.read())
    else:
        xlsx_path = XLSX_PATH

    # Verificar que el xlsx existe
    if not xlsx_path.exists():
        return _render(
            "index.html",
            {
                "request": request,
                "error": f"Archivo Excel no encontrado: {xlsx_path.name}",
            },
        )

    # Extraer datos del .docx
    try:
        datos = extractor.extract(docx_path)
    except Exception as e:
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
            "docx_path": str(docx_path),
            "xlsx_path": str(xlsx_path),
        },
    )


@app.post("/insertar", response_class=HTMLResponse)
async def insertar(
    request: Request,
    docx_path: str = Form(...),
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

    if not xlsx.exists():
        return _render(
            "result.html",
            {
                "request": request,
                "success": False,
                "duplicate": False,
                "error_msg": f"Archivo Excel no encontrado: {xlsx.name}",
                "numero": datos.numero,
                "xlsx_name": xlsx.name,
            },
        )

    # Intentar insertar
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
