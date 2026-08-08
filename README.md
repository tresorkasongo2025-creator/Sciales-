# Application SCIENCES SOCIALES - UNILU

Application web de gestion des inscriptions et contrôle de présence pour la Faculté des Sciences Sociales, Politiques et Administratives de l'Université de Lubumbashi.

## 🎯 Fonctionnalités

### ✅ Inscription en ligne
- Formulaire d'inscription complet avec photo
- Génération automatique de matricule (ETU000001, ETU000002...)
- Création automatique de QR code personnalisé
- Base de données PostgreSQL

### 🏛️ Espace DECANAT
- Accès sécurisé par mot de passe
- Gestion des listes par département et promotion
- Export PDF et Excel
- Tableau de bord statistiques

### 📱 Contrôle de présence
- Scan de QR code via caméra
- Enregistrement automatique de l'heure et la date
- Support étudiants et professeurs
- Listes de présence en temps réel

## 🚀 Démarrage rapide

1. **L'application est déjà en ligne !** Cliquez simplement sur le bouton de prévisualisation web.

2. **Accès DECANAT** : 
   - Mot de passe par défaut : `DECANAT2026`
   - Peut être changé via variable d'environnement `DECANAT_PASSWORD`

## 📚 Départements

- Relations Internationales (RI)
- Sciences Politiques et Administratives - Politique (SPA POL)
- Sciences Politiques et Administratives - SAM (SPA SAM)
- Sociologie
- Anthropologie

## 🎓 Promotions

BAC 1, BAC 2, BAC 3, Master 1, Master 2

## 🔒 Sécurité

Pour un environnement de production, configurez ces variables d'environnement :
- `SECRET_KEY` - Clé secrète Flask
- `DECANAT_PASSWORD` - Mot de passe administration

## 📖 Documentation complète

Consultez `replit.md` pour la documentation technique complète.

## 🛠️ Technologies

- **Backend** : Flask, SQLAlchemy, PostgreSQL
- **Frontend** : HTML5, CSS3, JavaScript
- **QR Code** : qrcode, html5-qrcode
- **Export** : ReportLab (PDF), OpenPyXL (Excel)

---

© 2025 Université de Lubumbashi - Faculté des Sciences Sociales
