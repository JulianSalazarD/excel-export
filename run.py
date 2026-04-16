"""
Punto de entrada para el ejecutable compilado con Nuitka.
Uso: ./run  (levanta el servidor en http://localhost:8000)
"""
import threading
import time
import webbrowser

import uvicorn
from app import app

URL = "http://127.0.0.1:8000"

def _abrir_navegador():
    time.sleep(1.5)  # espera a que uvicorn levante
    webbrowser.open(URL)

if __name__ == "__main__":
    threading.Thread(target=_abrir_navegador, daemon=True).start()
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=False)
