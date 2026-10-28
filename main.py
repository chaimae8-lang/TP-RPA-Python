# -*- coding: utf-8 -*-
"""
==============================================================================
 Master AI & DS - M2I S3 : Automation & Agentic AI  |  Prof. Imane Lahbari
 TP2 RPA Python : Excel  ->  Formulaire web  ->  Emails (Mailtrap)
==============================================================================
Usage :
    python main.py                          # PyAutoGUI + emails
    python main.py --mode selenium          # Selenium + emails
    python main.py --mode playwright        # Playwright + emails
    python main.py --mode none              # emails uniquement
    python main.py --no-email --mode selenium   # formulaire uniquement
    python main.py --url "https://mon-site/contact"
"""

import argparse
import logging
import os
import re
import smtplib
import sys
import time
import unicodedata
import webbrowser
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd

# python-dotenv est optionnel : sans lui on lit simplement les variables d'environnement
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ==============================================================================
# 1) CONFIGURATION (à adapter)
# ==============================================================================

# --- Fichier Excel ---
FICHIER_EXCEL = "donnees.xlsx"
COLONNES_REQUISES = ["Nom", "Email", "Message"]

# --- Formulaire web ---
# Par défaut : le formulaire HTML local fourni avec le TP
URL_FORMULAIRE = Path(__file__).with_name("formulaire.html").resolve().as_uri()

# Sélecteurs CSS des champs (Selenium / Playwright) -> à adapter à VOTRE formulaire
SELECTEURS = {
    "nom": "#nom",
    "email": "#email",
    "message": "#message",
    "bouton": "#envoyer",
}

# --- Délais (secondes) ---
DELAI_ACTION = 0.7          # entre deux actions (saisie, clic...)
DELAI_ENTRE_LIGNES = 1.5    # entre deux lignes Excel
DELAI_CHARGEMENT = 4        # attente d'ouverture du navigateur (PyAutoGUI)

# --- Mailtrap (valeurs lues depuis le fichier .env) ---
MAILTRAP_HOST = os.getenv("MAILTRAP_HOST", "smtp.mailtrap.io")
MAILTRAP_PORT = int(os.getenv("MAILTRAP_PORT", "2525"))
MAILTRAP_USER = os.getenv("MAILTRAP_USER", "")
MAILTRAP_PASSWORD = os.getenv("MAILTRAP_PASSWORD", "")
EXPEDITEUR = os.getenv("EXPEDITEUR", "bot-rpa@tp2.local")

# --- Logs ---
FICHIER_LOG = "execution_rpa.log"

# ==============================================================================
# 2) LOGGING (console + fichier)
# ==============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(FICHIER_LOG, encoding="utf-8"),
    ],
)
log = logging.getLogger("RPA")


# ==============================================================================
# 3) ÉTAPE 2 : LECTURE DU FICHIER EXCEL
# ==============================================================================
def lire_excel(chemin: str) -> pd.DataFrame:
    """Charge le fichier Excel, vérifie les colonnes et nettoie les données."""
    try:
        df = pd.read_excel(chemin)
    except FileNotFoundError:
        log.error(f"❌ Fichier introuvable : '{chemin}'. Lancez d'abord : python generer_excel.py")
        raise
    except Exception as e:
        log.error(f"❌ Impossible de lire '{chemin}' : {e}")
        raise

    manquantes = [c for c in COLONNES_REQUISES if c not in df.columns]
    if manquantes:
        raise ValueError(f"Colonnes manquantes dans l'Excel : {manquantes}")

    # Nettoyage : suppression des lignes vides et des espaces superflus
    df = df[COLONNES_REQUISES].dropna(how="all").fillna("").astype(str)
    df = df.apply(lambda col: col.str.strip()).reset_index(drop=True)

    log.info(f"✅ Fichier '{chemin}' chargé : {len(df)} ligne(s).")
    print("\n📄 Aperçu des données (head) :")
    print(df.head(), "\n")
    return df


# ==============================================================================
# 4) ÉTAPE 3 : AUTOMATISATION DU FORMULAIRE WEB
# ==============================================================================

# ---------- a) PyAutoGUI (simulation clavier) ----------
def _sans_accents(texte: str) -> str:
    """pyautogui.write() ne gère pas bien les caractères accentués -> on les retire."""
    nfkd = unicodedata.normalize("NFKD", texte)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def remplir_pyautogui(df: pd.DataFrame, url: str) -> int:
    """
    Remplit le formulaire en simulant le clavier.
    ⚠️ Ne touchez ni à la souris ni au clavier pendant l'exécution.
    ⚠️ Sécurité PyAutoGUI : déplacez la souris dans un coin de l'écran pour tout arrêter.
    Hypothèse : à l'ouverture, le curseur est dans le champ "Nom" ; l'ordre de tabulation
    est Nom -> Email -> Message -> Bouton, et le formulaire se vide après envoi.
    """
    import pyautogui  # import local : inutile si on n'utilise pas ce mode

    pyautogui.PAUSE = 0.2
    pyautogui.FAILSAFE = True
    reussies = 0

    try:
        log.info(f"🌐 Ouverture du navigateur : {url}")
        if not webbrowser.open(url):
            raise RuntimeError("Impossible d'ouvrir le navigateur.")
        log.info(f"⏳ Attente de {DELAI_CHARGEMENT}s pour le chargement de la page...")
        time.sleep(DELAI_CHARGEMENT)
    except Exception as e:
        log.error(f"❌ Formulaire inaccessible : {e}")
        return 0

    for i, ligne in df.iterrows():
        try:
            log.info(f"✍️  [PyAutoGUI] Ligne {i + 1}/{len(df)} : {ligne['Nom']}")
            pyautogui.write(_sans_accents(ligne["Nom"]), interval=0.03)
            time.sleep(DELAI_ACTION)
            pyautogui.press("tab")
            pyautogui.write(_sans_accents(ligne["Email"]), interval=0.03)
            time.sleep(DELAI_ACTION)
            pyautogui.press("tab")
            pyautogui.write(_sans_accents(ligne["Message"]), interval=0.02)
            time.sleep(DELAI_ACTION)
            pyautogui.press("tab")      # aller sur le bouton
            pyautogui.press("enter")    # envoyer
            time.sleep(DELAI_ENTRE_LIGNES)
            log.info(f"✅ Ligne {i + 1} saisie dans le formulaire.")
            reussies += 1
        except pyautogui.FailSafeException:
            log.error("🛑 Arrêt d'urgence (souris dans le coin de l'écran).")
            break
        except Exception as e:
            log.error(f"❌ Erreur ligne {i + 1} : {e}")
    return reussies


# ---------- b) Selenium ----------
def remplir_selenium(df: pd.DataFrame, url: str) -> int:
    """Remplit le formulaire avec Selenium WebDriver (Selenium >= 4.6 gère le driver seul)."""
    from selenium import webdriver
    from selenium.common.exceptions import TimeoutException, WebDriverException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait

    reussies = 0
    driver = None
    try:
        driver = webdriver.Chrome()
        driver.maximize_window()
        log.info(f"🌐 [Selenium] Ouverture : {url}")
        driver.get(url)
        attente = WebDriverWait(driver, 10)
        attente.until(EC.presence_of_element_located((By.CSS_SELECTOR, SELECTEURS["nom"])))
    except (TimeoutException, WebDriverException) as e:
        log.error(f"❌ Formulaire inaccessible (Selenium) : {e}")
        if driver:
            driver.quit()
        return 0

    try:
        for i, ligne in df.iterrows():
            try:
                log.info(f"✍️  [Selenium] Ligne {i + 1}/{len(df)} : {ligne['Nom']}")
                for cle, colonne in (("nom", "Nom"), ("email", "Email"), ("message", "Message")):
                    champ = attente.until(
                        EC.element_to_be_clickable((By.CSS_SELECTOR, SELECTEURS[cle]))
                    )
                    champ.clear()
                    champ.send_keys(ligne[colonne])
                    time.sleep(DELAI_ACTION)
                driver.find_element(By.CSS_SELECTOR, SELECTEURS["bouton"]).click()
                time.sleep(DELAI_ENTRE_LIGNES)
                log.info(f"✅ Ligne {i + 1} saisie dans le formulaire.")
                reussies += 1
            except Exception as e:
                log.error(f"❌ Erreur ligne {i + 1} (Selenium) : {e}")
    finally:
        time.sleep(2)
        driver.quit()
    return reussies


# ---------- c) Playwright ----------
def remplir_playwright(df: pd.DataFrame, url: str) -> int:
    """Remplit le formulaire avec Playwright (nécessite : playwright install chromium)."""
    from playwright.sync_api import Error as PlaywrightError
    from playwright.sync_api import sync_playwright

    reussies = 0
    try:
        with sync_playwright() as p:
            navigateur = p.chromium.launch(headless=False, slow_mo=200)
            page = navigateur.new_page()
            try:
                log.info(f"🌐 [Playwright] Ouverture : {url}")
                page.goto(url, timeout=15000)
                page.wait_for_selector(SELECTEURS["nom"], timeout=10000)
            except PlaywrightError as e:
                log.error(f"❌ Formulaire inaccessible (Playwright) : {e}")
                navigateur.close()
                return 0

            for i, ligne in df.iterrows():
                try:
                    log.info(f"✍️  [Playwright] Ligne {i + 1}/{len(df)} : {ligne['Nom']}")
                    page.fill(SELECTEURS["nom"], ligne["Nom"])
                    time.sleep(DELAI_ACTION)
                    page.fill(SELECTEURS["email"], ligne["Email"])
                    time.sleep(DELAI_ACTION)
                    page.fill(SELECTEURS["message"], ligne["Message"])
                    time.sleep(DELAI_ACTION)
                    page.click(SELECTEURS["bouton"])
                    time.sleep(DELAI_ENTRE_LIGNES)
                    log.info(f"✅ Ligne {i + 1} saisie dans le formulaire.")
                    reussies += 1
                except Exception as e:
                    log.error(f"❌ Erreur ligne {i + 1} (Playwright) : {e}")
            time.sleep(2)
            navigateur.close()
    except Exception as e:
        log.error(f"❌ Erreur Playwright : {e} (avez-vous lancé 'playwright install chromium' ?)")
    return reussies


# ==============================================================================
# 5) ÉTAPE 4 : ENVOI DES EMAILS VIA MAILTRAP
# ==============================================================================
REGEX_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def construire_email(nom: str, email: str, message: str) -> MIMEText:
    """Crée un email personnalisé à partir d'une ligne Excel."""
    corps = (
        f"Bonjour {nom},\n\n"
        f"Nous avons bien reçu votre message :\n"
        f"« {message} »\n\n"
        f"Nous vous répondrons à l'adresse {email} dans les meilleurs délais.\n\n"
        f"Cordialement,\nL'équipe RPA - Master AI & DS"
    )
    msg = MIMEText(corps, "plain", "utf-8")
    msg["Subject"] = f"Confirmation de votre message, {nom}"
    msg["From"] = EXPEDITEUR
    msg["To"] = email
    return msg


def envoyer_emails(df: pd.DataFrame) -> int:
    """Envoie un email par ligne via le SMTP de Mailtrap (TLS, port 2525)."""
    if not MAILTRAP_USER or not MAILTRAP_PASSWORD:
        log.error("❌ Identifiants Mailtrap absents : renseignez le fichier .env")
        return 0

    envoyes = 0
    try:
        log.info(f"📡 Connexion à {MAILTRAP_HOST}:{MAILTRAP_PORT} ...")
        with smtplib.SMTP(MAILTRAP_HOST, MAILTRAP_PORT, timeout=20) as serveur:
            serveur.ehlo()
            serveur.starttls()          # passage en TLS
            serveur.ehlo()
            serveur.login(MAILTRAP_USER, MAILTRAP_PASSWORD)
            log.info("🔐 Authentification Mailtrap réussie.")

            for i, ligne in df.iterrows():
                nom, email, message = ligne["Nom"], ligne["Email"], ligne["Message"]
                try:
                    if not REGEX_EMAIL.match(email):
                        raise ValueError(f"adresse email invalide : '{email}'")
                    msg = construire_email(nom, email, message)
                    serveur.sendmail(EXPEDITEUR, [email], msg.as_string())
                    log.info(f"✅ Email {i + 1}/{len(df)} envoyé à {nom} <{email}>")
                    envoyes += 1
                except Exception as e:
                    log.error(f"❌ Échec email {i + 1} ({nom}) : {e}")
                time.sleep(DELAI_ACTION)   # évite la limite de débit de Mailtrap
    except smtplib.SMTPAuthenticationError:
        log.error("❌ Authentification refusée : vérifiez MAILTRAP_USER / MAILTRAP_PASSWORD.")
    except (smtplib.SMTPException, OSError) as e:
        log.error(f"❌ Erreur de connexion SMTP : {e}")
    return envoyes


# ==============================================================================
# 6) ÉTAPE 5 : PROGRAMME PRINCIPAL + BILAN
# ==============================================================================
def main():
    parser = argparse.ArgumentParser(description="TP2 RPA : Excel -> Formulaire -> Emails")
    parser.add_argument("--mode", choices=["pyautogui", "selenium", "playwright", "none"],
                        default="pyautogui", help="outil d'automatisation du formulaire")
    parser.add_argument("--url", default=URL_FORMULAIRE, help="URL du formulaire")
    parser.add_argument("--excel", default=FICHIER_EXCEL, help="chemin du fichier Excel")
    parser.add_argument("--no-email", action="store_true", help="ne pas envoyer d'emails")
    args = parser.parse_args()

    print("=" * 60)
    print("🤖 TP2 RPA - Master AI & DS M2I - Automation & Agentic AI")
    print("=" * 60)

    # Étape 2 : lecture Excel
    try:
        df = lire_excel(args.excel)
    except Exception:
        sys.exit(1)
    total = len(df)

    # Étape 3 : formulaire
    saisies = None
    if args.mode != "none":
        outils = {"pyautogui": remplir_pyautogui,
                  "selenium": remplir_selenium,
                  "playwright": remplir_playwright}
        log.info(f"▶️  Étape 3 : remplissage du formulaire avec {args.mode.upper()}")
        saisies = outils[args.mode](df, args.url)

    # Étape 4 : emails
    envoyes = None
    if not args.no_email:
        log.info("▶️  Étape 4 : envoi des emails via Mailtrap")
        envoyes = envoyer_emails(df)

    # Étape 5 : bilan
    print("\n" + "=" * 60)
    print("📊 BILAN DE L'EXÉCUTION")
    print("=" * 60)
    print(f"Lignes Excel traitées : {total}")
    if saisies is not None:
        print(f"{'✅' if saisies == total else '❌'} Formulaire : {saisies}/{total} ligne(s) saisie(s)")
    if envoyes is not None:
        print(f"{'✅' if envoyes == total else '❌'} Emails     : {envoyes}/{total} envoyé(s)")
    print(f"📝 Détails dans : {FICHIER_LOG}")


if __name__ == "__main__":
    main()
