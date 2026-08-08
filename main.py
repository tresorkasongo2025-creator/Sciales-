"""
Entrypoint déclaré dans .replit (entrypoint = "main.py").
Lance le serveur de développement Flask ; en production gunicorn importe
directement app:app et n'utilise pas ce fichier.
"""
from app import app

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
