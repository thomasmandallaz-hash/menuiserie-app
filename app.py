import streamlit as st
import pandas as pd
from datetime import datetime
import os

st.set_page_config(page_title="Gestion Menuiserie", layout="wide")

# --- CHARGEMENT AUTOMATIQUE DES DONNÉES EXCEL ---
@st.cache_data
def charger_donnees():
    # Chargement du stock 2025
    if os.path.exists("Inventaire stock .xlsx"):
        df_inv = pd.read_excel("Inventaire stock .xlsx", sheet_name="Inventaire pour bilan 2025")
        df_inv = df_inv.dropna(subset=[df_inv.columns[0]])
        df_inv.columns = ["Désignation", "Quantité", "Prix Unitaire HT", "Unité", "Total HT", "Col6", "Col7"][:len(df_inv.columns)]
        df_inv = df_inv[df_inv["Désignation"] != "Désignation"]
        df_stock = df_inv[["Désignation", "Quantité", "Prix Unitaire HT", "Unité"]].dropna(subset=["Désignation"])
    else:
        df_stock = pd.DataFrame([{"Désignation": "VIS 3x10", "Quantité": 50, "Prix Unitaire HT": 0.99, "Unité": "centaine"}])

    # Chargement des projets Kimai
    if os.path.exists("kimai-projects_20260920152909.xlsx"):
        df_p = pd.read_excel("kimai-projects_20260920152909.xlsx")
        df_p = df_p.dropna(subset=["Nom"])
        liste_projets = df_p["Nom"].astype(str).tolist()
    else:
        liste_projets = ["26/221 Fabrication et pose d'étagères", "26/227 Réfection plan de travail"]

    return df_stock, liste_projets

df_stock_base, LISTE_CHANTIERS = charger_donnees()

LISTE_TACHES = [
    "M1 - Montage atelier", 
    "D2 - Débit bois", 
    "P1 - Pose chantier", 
    "X5 - Bureau / Administration",
    "X8 - Devis & Métrés"
]

# --- NAVIGATION ---
st.sidebar.title("🛠️ Menu Principal")
menu = st.sidebar.radio(
    "Accéder à :",
    ["📊 Rentabilité & Planning", "📝 Pense-bête Chantiers", "💰 Trésorerie & Engagements", "📦 Stock & QR Codes"]
)

# Initialisation des états
if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

if 'pense_bete' not in st.session_state:
    st.session_state['pense_bete'] = {}

if 'mouvements_treso' not in st.session_state:
    st.session_state['mouvements_treso'] = []

if 'stock_articles' not in st.session_state:
    st.session_state['stock_articles'] = df_stock_base.to_dict('records')

# ==============================================================================
# MODULE 1 : RENTABILITÉ & PLANNING
# ==============================================================================
if menu == "📊 Rentabilité & Planning":
    st.title("🔨 Saisie des Heures & Rentabilité")

    st.sidebar.header("⏱️ Saisie Rapide")
    with st.sidebar.form("form_heures"):
        date_saisie = st.date_input("Date", datetime.today())
        chantier = st.selectbox("Chantier / Projet (Kimai)", LISTE_CHANTIERS)
        code_tache = st.selectbox("Code Tâche", LISTE_TACHES)
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

    st.progress(min(heures_prod / objectif, 1.0))

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
    
    total_encaisse = df_treso[(df_treso['Type'] == 'Encaissement') & (df_treso['Statut'] == 'Réalisé / Encaissé')]['Montant HT'].sum() if not df_treso.empty else 0.0
    total_engag_depenses = df_treso[df_treso['Type'] == 'Engagement Dépense']['Montant HT'].sum() if not df_treso.empty else 0.0
    solde_previsionnel = total_encaisse - total_engag_depenses

    c1, c2, c3 = st.columns(3)
    c1.metric("Encaissements Réalisés", f"{total_encaisse:,.2f} € HT".replace(",", " "))
    c2.metric("Engagements Dépenses (Stocks/Achats)", f"{total_engag_depenses:,.2f} € HT".replace(",", " "))
    c3.metric("Solde Prévisionnel Net", f"{solde_previsionnel:,.2f} € HT".replace(",", " "))

    if not df_treso.empty:
        st.header("📋 Historique des Mouvements")
        st.dataframe(df_treso, use_container_width=True)

# ==============================================================================
# MODULE 4 : STOCK & QR CODES
# ==============================================================================
elif menu == "📦 Stock & QR Codes":
    st.title("📦 Gestion des Stocks (Données Bilan 2025)")

    df_stock = pd.DataFrame(st.session_state['stock_articles'])
    st.write(f"Total des références en base : **{len(df_stock)} articles**")
    st.dataframe(df_stock, use_container_width=True)
