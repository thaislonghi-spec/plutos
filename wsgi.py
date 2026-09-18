# Porta de entrada alternativa: permite `gunicorn wsgi:app` a partir da raiz.
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "app"))
from server import app  # noqa: E402,F401
