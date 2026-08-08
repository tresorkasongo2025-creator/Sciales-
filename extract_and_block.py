"""Extrait tous les numéros de reçu des PDFs perdus et les bloque en base."""
import sys
sys.path.insert(0, '/home/runner/workspace')

import fitz
import re

pdf_files = [
    '/home/runner/workspace/attached_assets/Recus_SC_Lot1_1785509571493.pdf',
    '/home/runner/workspace/attached_assets/Recus_SC_Lot1_1785509814262.pdf',
]

numeros = set()

for pdf_path in pdf_files:
    doc = fitz.open(pdf_path)
    for page in doc:
        text = page.get_text()
        # Format "Preuve de paiement : B1-SC-XXX-S2-26"
        found = re.findall(r'[A-Z0-9]+-[A-Z]+-\d+-[A-Z0-9]+-\d+', text)
        for f in found:
            numeros.add(f.strip())
    doc.close()

print(f"Numéros extraits : {len(numeros)}")
for n in sorted(numeros):
    print(f"  {n}")

# Maintenant bloquer en base
from app import app, db, RecuPaiement
from zoneinfo import ZoneInfo
from datetime import datetime

_CAT = ZoneInfo('Africa/Lubumbashi')
now = datetime.now(_CAT).replace(tzinfo=None)

RAISON = 'BLOQUÉ - lot perdu, risque de fraude'

with app.app_context():
    bloques = 0
    deja_utilises = 0
    introuvables = []

    for numero in sorted(numeros):
        # Chercher par numero OU code_qr (les deux peuvent matcher)
        recu = RecuPaiement.query.filter(
            (RecuPaiement.numero == numero) | (RecuPaiement.code_qr == numero)
        ).first()
        
        if recu is None:
            introuvables.append(numero)
            continue
        
        if recu.utilise:
            deja_utilises += 1
            print(f"  DÉJÀ UTILISÉ : {numero} → matricule={recu.matricule_etudiant}")
            continue
        
        # Marquer comme utilisé avec raison
        recu.utilise = True
        recu.date_utilisation = now
        recu.nom_etudiant = RAISON
        recu.matricule_etudiant = 'BLOQUÉ'
        bloques += 1

    db.session.commit()
    
    print(f"\n=== RÉSULTAT ===")
    print(f"  Bloqués avec succès : {bloques}")
    print(f"  Déjà utilisés (légitimes?) : {deja_utilises}")
    print(f"  Introuvables en base : {len(introuvables)}")
    if introuvables:
        print(f"  Introuvables : {introuvables[:10]}")
    
    # Vérification finale : compter les reçus SC Lot1 non utilisés restants
    restants = RecuPaiement.query.filter_by(lot='1', dept='SC', utilise=False).count()
    print(f"\n  Reçus SC Lot1 encore libres en base : {restants}")
    
    # Afficher les reçus bloqués pour confirmation
    bloques_q = RecuPaiement.query.filter_by(
        lot='1', dept='SC', matricule_etudiant='BLOQUÉ'
    ).count()
    print(f"  Reçus SC Lot1 marqués BLOQUÉ : {bloques_q}")
