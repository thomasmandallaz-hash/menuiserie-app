import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Gestion Menuiserie", layout="wide")

# --- NAVIGATION DANS L'APPLICATION ---
st.sidebar.title("🛠️ Menu Principal")
menu = st.sidebar.radio(
    "Accéder à :",
    ["📊 Rentabilité & Planning", "💰 Trésorerie & Engagements", "📦 Stock & QR Codes"]
)

# Initialisation des bases de données en mémoire (session)
if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

if 'mouvements_treso' not in st.session_state:
    st.session_state['mouvements_treso'] = [
        {"Date": "2026-09-15", "Libellé": "Acompte Cuisine Dupont", "Type": "Encaissement", "Montant HT": 3500.0, "Statut": "Réalisé"},
        {"Date": "2026-09-18", "Libellé": "Commande Quincaillerie Foussier", "Type": "Engagement Dépense", "Montant HT": 480.0, "Statut": "Engagé"}
    ]

if 'stock_articles' not in st.session_state:
    st.session_state['stock_articles'] = [
        {"Réf": "FOUS-283222", "Désignation": "Charnière Blum 110° Applique", "Fournisseur": "Foussier", "Quantité": 45, "Seuil Min": 10, "Prix HT": 3.20},
        {"Réf": "FOUS-241138", "Désignation": "Coulisse Movento 500mm", "Fournisseur": "Foussier", "Quantité": 8, "Seuil Min": 12, "Prix HT": 24.50},
        {"Réf": "WURTH-102345", "Désignation": "Vis Bois 3.5x30 (Boîte de 500)", "Fournisseur": "Würth", "Quantité": 3, "Seuil Min": 2, "Prix HT": 18.90}
    ]

# ==============================================================================
# MODULE 1 : RENTABILITÉ & PLANNING
# ==============================================================================
if menu == "📊 Rentabilité & Planning":
    st.title("🔨 Saisie des Heures & Rentabilité")

    # Barre latérale pour la saisie rapide
    st.sidebar.header("⏱️ Saisie Rapide")
    with st.sidebar.form("form_heures"):
        date_saisie = st.date_input("Date", datetime.today())
        chantier = st.text_input("Chantier / Devis", value="2026-DEV-018 Cuisine Dupont")
        code_tache = st.selectbox("Code Tâche (Kimai)", [
            "M1 - Montage atelier", 
            "D2 - Débit bois", 
            "P1 - Pose chantier", 
            "X8 - Administration / Devis"
        ])
        heures = st.number_input("Nombre d'heures", min_value=0.5, max_value=12.0, value=7.5, step=0.5)
        valider = st.form_submit_button("Enregistrer les heures")

    if valider:
        est_prod = not code_tache.startswith("X")
        st.session_state['historique_heures'].append({
            "Date": date_saisie,
            "Chantier": chantier,
            "Code": code_tache,
            "Heures": heures,
            "Production": est_prod
        })
        st.sidebar.success("Heures enregistrées !")

    # Calculs de rentabilité
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

    st.header("📅 Planning de la Semaine")
    col_mar, col_mer, col_jeu, col_ven = st.columns(4)
    with col_mar:
        st.subheader("Mardi")
        st.info("7,5 h — Débit / Usinage\n\n*Cuisine Dupont*")
    with col_mer:
        st.subheader("Mercredi")
        st.info("7,5 h — Montage Atelier\n\n*Cuisine Dupont*")
    with col_jeu:
        st.subheader("Jeudi")
        st.info("7,5 h — Pose Chantier\n\n*Placard Martin*")
    with col_ven:
        st.subheader("Vendredi")
        st.info("5,5 h — Pose & Finitions\n\n*Placard Martin*")

    if not df_heures.empty:
        st.subheader("📋 Historique des saisies de la semaine")
        st.dataframe(df_heures, use_container_width=True)

# ==============================================================================
# MODULE 2 : TRÉSORERIE & ENGAGEMENTS
# ==============================================================================
elif menu == "💰 Trésorerie & Engagements":
    st.title("💰 Suivi de Trésorerie & Engagements")

    # Formulaire d'ajout
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

    # Calculs Synthèse
    df_treso = pd.DataFrame(st.session_state['mouvements_treso'])
    
    total_encaisse = df_treso[(df_treso['Type'] == 'Encaissement') & (df_treso['Statut'] == 'Réalisé / Encaissé')]['Montant HT'].sum()
    total_engag_depenses = df_treso[df_treso['Type'] == 'Engagement Dépense']['Montant HT'].sum()
    solde_previsionnel = total_encaisse - total_engag_depenses

    c1, c2, c3 = st.columns(3)
    c1.metric("Encaissements Réalisés", f"{total_encaisse:,.2f} € HT".replace(",", " "))
    c2.metric("Engagements Dépenses (Stocks/Achats)", f"{total_engag_depenses:,.2f} € HT".replace(",", " "))
    c3.metric("Solde Prévisionnel Net", f"{solde_previsionnel:,.2f} € HT".replace(",", " "))

    st.header("📋 Historique des Mouvements de Trésorerie")
    st.dataframe(df_treso, use_container_width=True)

# ==============================================================================
# MODULE 3 : STOCK & QR CODES
# ==============================================================================
elif menu == "📦 Stock & QR Codes":
    st.title("📦 Gestion des Stocks & Scanner")

    df_stock = pd.DataFrame(st.session_state['stock_articles'])

    # Alertes Réapprovisionnement
    stock_bas = df_stock[df_stock['Quantité'] <= df_stock['Seuil Min']]
    if not stock_bas.empty:
        st.warning(f"⚠️ **{len(stock_bas)} référence(s) sous le seuil minimum !** Réapprovisionnement nécessaire.")

    st.header("📋 État du Stock")
    st.dataframe(df_stock, use_container_width=True)

    st.header("📷 Saisie / Ajustement de Stock")
    ref_select = st.selectbox("Sélectionner une référence", df_stock['Réf'] + " - " + df_stock['Désignation'])
    col_plus, col_moins = st.columns(2)
    
    with col_plus:
        if st.button("➕ Ajouter au stock (+1)"):
            st.success("Stock mis à jour !")
    with col_moins:
        if st.button("➖ Retirer du stock (-1)"):
            st.info("Stock mis à jour !")
