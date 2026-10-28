# TP RPA Python - Automatisation de Formulaire et d'Emails

Ce dépôt contient les travaux pratiques (TP) réalisés dans le cadre du module **Automation & Agentic AI** du **Master AI & DS (M2I, S3)**.

L'objectif de ce projet est de concevoir un robot RPA (Robotic Process Automation) capable de lire des données à partir d'un fichier Excel, de remplir automatiquement un formulaire web, puis d'envoyer un email de confirmation personnalisé à chaque utilisateur référencé.

---

## 🚀 Fonctionnalités

Le script principal réalise les étapes d'automatisation suivantes :
1. **Extraction de données** : Lecture et chargement des profils stockés dans un tableur Excel via la bibliothèque `pandas`.
2. **Saisie automatisée (RPA)** : Simulation des actions clavier/souris à l'aide de `pyautogui` pour remplir et soumettre un formulaire web local (`formulaire.html`).
3. **Notification par email** : Connexion sécurisée au protocole de messagerie Gmail (`SMTP_SSL`) et envoi de courriels personnalisés au format MIME `text`.

---

## 📁 Structure du Projet

```text
├── donnees.xlsx          # Fichier Excel contenant les données (Nom, Email, Message)
├── formulaire.html       # Formulaire web cible à automatiser
├── generer_excel.py      # Script utilitaire pour générer le fichier Excel de test
├── main.py               # Script principal d'automatisation (RPA + Emailing)
├── env.example           # Fichier d'exemple pour la configuration des variables d'environnement
├── requirements.txt      # Liste des dépendances Python requises
└── README.md             # Documentation du projet
```

---

## 🛠️ Installation et Configuration

### 1. Prérequis
Assurez-vous d'avoir Python 3.8+ installé sur votre machine.

### 2. Cloner le dépôt
```bash
git clone https://github.com
cd RPA-Python-TP
```

### 3. Installer les dépendances
Installez les bibliothèques requises (comme `pandas`, `pyautogui`, `openpyxl`, `python-dotenv`) à l'aide de la commande suivante :
```bash
pip install -r requirements.txt
```

### 4. Configurer les variables d'environnement
Créez un fichier `.env` à la racine du projet en vous basant sur le fichier modèle fourni :
```bash
cp env.example .env
```
Éditez le fichier `.env` et ajoutez vos identifiants de messagerie sécurisés :
```ini
EMAIL_USER=votre.email@gmail.com
EMAIL_PASSWORD=votre_mot_de_passe_d_application
```
*Note : Pour Gmail, vous devez générer un "Mot de passe d'application" depuis les paramètres de sécurité de votre compte Google.*

---

## 🎯 Utilisation

1. Ouvrez le fichier `formulaire.html` dans votre navigateur Web habituel et assurez-vous que la fenêtre est bien visible à l'écran.
2. Si vous n'avez pas de données de test, générez-les d'abord en exécutant :
   ```bash
   python generer_excel.py
   ```
3. Exécutez le script d'automatisation principal :
   ```bash
   python main.py
   ```
4. **Attention** : Ne manipulez pas votre clavier ou votre souris pendant l'exécution de `main.py`, car `pyautogui` prend temporairement le contrôle de l'interface graphique de votre machine.

---

## 📝 Perspectives & Évolutions
Conformément aux directives du TP, une extension future de ce projet consistera à adapter l'automatisation du formulaire web en utilisant des frameworks plus robustes et modernes de web scraping / test tels que **Selenium** et **Playwright**, permettant une exécution en arrière-plan (headless).

---

**Auteur :** [Votre Nom & Prénom]  
**Filière :** Master Artificial Intelligence & Data Science (AI & DS - M2I)  
