# -*- coding: utf-8 -*-
"""
ÉTAPE 1 : Génération automatique du fichier Excel de test (donnees.xlsx)
Colonnes : Nom, Email, Message
Usage : python generer_excel.py
"""
import pandas as pd

FICHIER_SORTIE = "donnees.xlsx"

# Données de test (5 lignes)
donnees = [
    {"Nom": "Amine El Idrissi", "Email": "amine.elidrissi@example.com",
     "Message": "Bonjour, je souhaite recevoir plus d'informations sur le Master AI et DS."},
    {"Nom": "Sara Benali", "Email": "sara.benali@example.com",
     "Message": "Pouvez-vous me confirmer la date de la prochaine session ?"},
    {"Nom": "Youssef Alaoui", "Email": "youssef.alaoui@example.com",
     "Message": "Je rencontre un probleme d'acces a la plateforme de cours."},
    {"Nom": "Khadija Mansouri", "Email": "khadija.mansouri@example.com",
     "Message": "Merci pour votre retour, je valide mon inscription."},
    {"Nom": "Omar Tazi", "Email": "omar.tazi@example.com",
     "Message": "Je voudrais proposer un sujet de projet sur l'Agentic AI."},
]

if __name__ == "__main__":
    try:
        df = pd.DataFrame(donnees, columns=["Nom", "Email", "Message"])
        df.to_excel(FICHIER_SORTIE, index=False)
        print(f"✅ Fichier '{FICHIER_SORTIE}' créé avec {len(df)} lignes.")
        print(df.head())
    except Exception as e:
        print(f"❌ Erreur lors de la création du fichier Excel : {e}")
