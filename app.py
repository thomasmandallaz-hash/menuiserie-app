import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Gestion Menuiserie", layout="wide")

# --- NAVIGATION DANS L'APPLICATION ---
st.sidebar.title("🛠️ Menu Principal")
menu = st.sidebar.radio(
    "Accéder à :",
    ["📊 Rentabilité & Planning", "📝 Pense-bête Chantiers", "💰 Trésorerie & Engagements", "📦 Stock & QR Codes"]
)

# Liste des chantiers réels de ton Kimai
LISTE_CHANTIERS = [
    "26/227 Réfection plan de travail - Cornachon",
    "26/221 Fabrication et pose étagères - Givre",
    "25/169 Fourniture et pose bloc portes - Mercier / Malinvaud",
    "25/19 Fabrication étagères et passe plat - Norma",
    "25/20 Pose cuisine Joubert - Ambiance Intérieur",
    "OE 25/180 Fabrication et pose placard - Murgier",
    "OE 25/00 Travail de bureau - Agencement Bois Création",
    "OE 25/01 Aménagement Atelier - Agencement Bois Création",
    "OE 24/112 Fabrication et pose façade meuble - Murgier",
    "OE 24/86 Fabrication meuble de chambre - Murgier",
    "OE 24/94 Porte d'entrée - Guerraz"
]

LISTE_TACHES = [
    "M1 - Montage atelier", 
    "D2 - Débit bois", 
    "P1 - Pose chantier", 
    "X5 - Bureau / Administration",
    "X8 - Devis & Métrés"
]

# Initialisation des bases de données en mémoire (session)
if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

if 'pense_bete' not in st.session_state:
    st.session_state['pense_bete'] = {
        "26/221 Fabrication et pose étagères - Givre": "Penser aux taquets invisibles + livraison coulisses Movento.",
        "25/169 Fourniture et pose bloc portes - Mercier / Malinvaud": "Vérifier la hauteur sous linteau avant la pose du bloc porte."
    }

if 'mouvements_treso' not in st.session_state:
    st.session_state['mouvements_treso'] = [
        {"Date": "2026-09-15", "Libellé": "Acompte Cuisine Joubert", "Type": "Encaissement", "Montant HT": 3500.0, "Statut": "Réalisé / Encaissé"},
        {"Date": "2026-09-18", "Libellé": "Commande Quincaillerie Foussier", "Type": "Engagement Dépense", "Montant HT": 480.0, "Statut": "Engagé / En attente"}
    ]

if 'stock_articles' not in st.session_state:
    st.session_state['stock_articles'] = [
        {"Réf": "VIS-3x10", "Désignation": "VIS 3x10", "Quantité": 50, "Unité": "centaine", "Prix unitaire": 0.99},
        {"Réf": "VIS-3x16", "Désignation": "VIS 3x16", "Quantité": 100, "Unité": "centaine", "Prix unitaire": 1.05},
        {"Réf": "VIS-3.5x30", "Désignation": "VIS 3,5X30", "Quantité": 50, "Unité": "centaine", "Prix unitaire": 1.23},
        {"Réf": "CHARN-BLUM", "Désignation": "Charnières Blum 110° Applique", "Quantité": 40, "Unité": "pc", "Prix unitaire": 3.20},
        {"Réf": "COUL-MOV-500", "Désignation": "Coulisses Movento 500mm", "Quantité": 6, "Unité": "paire", "Prix unitaire": 24.50}
    ]

# ==============================================================================
# MODULE 1 : RENTABILITÉ & PLANNING
# ==============================================================================
if menu == "📊 Rentabilité & Planning":
    st.title("🔨 Saisie des Heures & Rentabilité")

    st.sidebar.header("⏱️ Saisie Rapide")
    with st.sidebar.form("form_heures"):
        date_saisie = st.date_input("Date", datetime.today())
        chantier = st.selectbox("Chantier / Projet", LISTE_CHANTIERS)
        code_tache = st.selectbox("Code Tâche (Kimai)", LISTE_TACHES)
        heures = st.number_input("Nombre d'heures", min_value=0.5, max_value=12.0, value=7.5, step=0.5)
        valider = st.form_submit_button("Enregistrer les heures")

    if valider:
        est_prod = not code_tache.startswith("X")
        st.session_state['historique_heures'].append({
            "Date": str(date_saisie),
            "Chantier": chantier,
            "Code": code_tache,
            "Heures": heures,
            "Production": est_prod
        })
        st.sidebar.success("Heures enregistrées !")

    df_heures = pd.DataFrame(st.session_state['historique_heures'])
    heures_prod = df_heures[df_heures['Production'] == True]['Heures'].sum() if not df_heures.empty else 0.0
    
    objectif = 33.50
    taux_horaire = 75.0
    progression = min(heures_prod / objectif, 1.0)
    ca_realise = heures_prod * taux_horaire

    col1, col2, col3 = st.columns(3)
    col1.metric("Heures de Production", f"{heures_prod:.1f} h / {objectif} h")
    col2.metric("Chiffre d'Affaires Produit", f"{ca_realise:,.2f} € HT".replace(",", " "))

    if heures_prod >= objectif:
        col3.success("🟢 Objectif de rentabilité atteint !")
    elif heures_prod >= 25.0:
        col3.warning("🟠 Semaine en cours (objectif proche)")
    else:
        col3.error("🔴 Sous le seuil de rentabilité")

    st.progress(progression)

    if not df_heures.empty:
        st.subheader("📋 Saisies de la semaine")
        st.dataframe(df_heures, use_container_width=True)

# ==============================================================================
# MODULE 2 : PENSE-BÊTE CHANTIERS
# ==============================================================================
elif menu == "📝 Pense-bête Chantiers":
    st.title("📝 Pense-bête & Mémos par Chantier")
    
    chantier_sel = st.selectbox("Sélectionner un chantier", LISTE_CHANTIERS)
    
    note_actuelle = st.session_state['pense_bete'].get(chantier_sel, "")
    
    st.subheader(f"Notes pour : {chantier_sel}")
    nouvelle_note = st.text_area("Rédiger ou modifier les remarques / quincaillerie à prévoir :", value=note_actuelle, height=150)
    
    if st.button("💾 Enregistrer la note"):
        st.session_state['pense_bete'][chantier_sel] = nouvelle_note
        st.success("Pense-bête mis à jour avec succès !")

# ==============================================================================
# MODULE 3 : TRÉSORERIE & ENGAGEMENTS
# ==============================================================================
elif menu == "💰 Trésorerie & Engagements":
    st.title("💰 Suivi de Trésorerie & Engagements")

    st.sidebar.header("➕ Nouveau Mouvement")
    with st.sidebar.form("form_treso"):
        date_treso = st.date_input("Date", datetime.today())
        libelle = st.text_input("Libellé", value="Acompte client")
        type_mvt = st.selectbox("Type", ["Encaissement", "Engagement Dépense"])
        montant = st.number_input("Montant HT (€)", min_value=0.0, value=500.0, step=50.0)
        statut = st.selectbox("Statut", ["Engagé / En attente", "Réalisé / Encaissé"])
        valider_treso = st.form_submit_button("Enregistrer le mouvement")

    if valider_treso:
        st.session_state['mouvements_treso'].append({
            "Date": str(date_treso),
            "Libellé": libelle,
            "Type": type_mvt,
            "Montant HT": montant,
            "Statut": statut
        })
        st.sidebar.success("Mouvement enregistré !")

    df_treso = pd.DataFrame(st.session_state['mouvements_treso'])
    
    total_encaisse = df_treso[(df_treso['Type'] == 'Encaissement') & (df_treso['Statut'] == 'Réalisé / Encaissé')]['Montant HT'].sum()
    total_engag_depenses = df_treso[df_treso['Type'] == 'Engagement Dépense']['Montant HT'].sum()
    solde_previsionnel = total_encaisse - total_engag_depenses

    c1, c2, c3 = st.columns(3)
    c1.metric("Encaissements Réalisés", f"{total_encaisse:,.2f} € HT".replace(",", " "))
    c2.metric("Engagements Dépenses (Stocks/Achats)", f"{total_engag_depenses:,.2f} € HT".replace(",", " "))
    c3.metric("Solde Prévisionnel Net", f"{solde_previsionnel:,.2f} € HT".replace(",", " "))

    st.header("📋 Historique des Mouvements")
    st.dataframe(df_treso, use_container_width=True)

# ==============================================================================
# MODULE 4 : STOCK & QR CODES
# ==============================================================================
elif menu == "📦 Stock & QR Codes":
    st.title("📦 Gestion des Stocks & Quincaillerie")

    df_stock = pd.DataFrame(st.session_state['stock_articles'])
    st.dataframe(df_stock, use_container_width=True)
