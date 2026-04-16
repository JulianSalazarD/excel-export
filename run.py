"""
Punto de entrada para el ejecutable compilado con Nuitka.
Uso: ./run  (levanta el servidor en http://localhost:8000)
"""
import uvicorn
from app import app

if __name__ == "__main__":
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=False)
