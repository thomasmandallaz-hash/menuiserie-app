import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Gestion Menuiserie", layout="wide")

st.title("🔨 Gestion & Pilotage Menuiserie")

# --- BARRE LATÉRALE : SAISIE RAPIDE DES HEURES ---
st.sidebar.header("⏱️ Saisie Rapide des Heures")
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

if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

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

# --- SECTION 1 : JAUGE DE RENTABILITÉ ---
st.header("📊 Rentabilité de la Semaine")

df_heures = pd.DataFrame(st.session_state['historique_heures'])

if not df_heures.empty:
    heures_prod = df_heures[df_heures['Production'] == True]['Heures'].sum()
else:
    heures_prod = 0.0

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

# --- SECTION 2 : PLANNING DE LA SEMAINE ---
st.header("📅 Planning & Suivi des Journées")

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

# --- SECTION 3 : HISTORIQUE DES SAISIES ---
if not df_heures.empty:
    st.subheader("📋 Saisies de la semaine")
    st.dataframe(df_heures, use_container_width=True)
