import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io

# Import des bibliothèques d'impression
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import qrcode

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
        df_inv["Réf"] = df_inv["Désignation"].str.replace(r'[^a-zA-Z0-9\s-]', '', regex=True).str.strip().str.replace(' ', '-')
        df_stock = df_inv[["Réf", "Désignation", "Quantité", "Prix Unitaire HT", "Unité"]].dropna(subset=["Désignation"])
    else:
        df_stock = pd.DataFrame([{"Réf": "VIS-3x10", "Désignation": "VIS 3x10", "Quantité": 50, "Prix Unitaire HT": 0.99, "Unité": "centaine"}])

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
    [
        "📊 Rentabilité & Planning", 
        "🧮 Brouillon Devis & Marges",
        "📝 Pense-bête Chantiers", 
        "💰 Trésorerie & Engagements", 
        "🏦 Remises de Chèques",
        "📦 Stock & QR Codes"
    ]
)

# Initialisation des états
if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

if 'pense_bete' not in st.session_state:
    st.session_state['pense_bete'] = {}

if 'mouvements_treso' not in st.session_state:
    st.session_state['mouvements_treso'] = []

if 'remises_cheques' not in st.session_state:
    st.session_state['remises_cheques'] = []

if 'stock_articles' not in st.session_state:
    st.session_state['stock_articles'] = df_stock_base.to_dict('records')

# FONCTION GENERATION PDF (Format Avery 21/feuille : 33,5 mm × 38,1 mm)
def generer_pdf_etiquettes(df):
    buffer = io.BytesIO()
    p = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    
    w_label = 38.1 * mm
    h_label = 33.5 * mm
    margin_x = 0 * mm
    margin_y = 0 * mm

    p.setFont("Helvetica-Bold", 8)
    
    num_cols = 3
    num_rows = 7
    total_labels_page = num_cols * num_rows

    for i, row in df.iterrows():
        label_idx = i % total_labels_page
        if label_idx == 0 and i > 0:
            p.showPage()
            p.setFont("Helvetica-Bold", 8)
            label_idx = 0

        col = label_idx % num_cols
        r = label_idx // num_cols

        x = margin_x + (col * w_label)
        y = height - (margin_y + (r + 1) * h_label)

        # Correction : extraction de l'image au format PIL
        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data(str(row['Réf']))
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white").get_image()
        
        # Dessin direct de l'image PIL
        p.drawInlineImage(img_qr, x + 2*mm, y + 2*mm, width=15*mm, height=15*mm)

        designation = str(row['Désignation'])
        if len(designation) > 28:
            designation = designation[:26] + "..."
            
        p.drawString(x + 19*mm, y + 13*mm, designation)
        p.setFont("Helvetica", 7)
        p.drawString(x + 19*mm, y + 10*mm, f"Réf: {row['Réf']}")
        p.drawString(x + 19*mm, y + 7*mm, f"Prix: {row['Prix Unitaire HT']:.2f} €HT / {row['Unité']}")
        p.setFont("Helvetica-Bold", 8)
        p.rect(x, y, w_label, h_label)

    p.save()
    buffer.seek(0)
    return buffer

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
# MODULE 2 : BROUILLON DEVIS & MARGES
# ==============================================================================
elif menu == "🧮 Brouillon Devis & Marges":
    st.title("🧮 Brouillon de Devis & Calculateur de Marges")
    st.info("Ce calculateur reproduit la trame de votre fichier Excel 00 DEVIS.xlsx pour chiffrer vos fournitures et heures.")

    st.subheader("1. Fournitures & Matériaux")
    col_mat1, col_mat2 = st.columns(2)
    
    with col_mat1:
        m2_caissons = st.number_input("M² Panneaux de caissons", value=0.0, step=1.0)
        m2_facades = st.number_input("M² Panneaux de façade", value=0.0, step=1.0)
        m2_fond = st.number_input("M² Panneau de fond", value=0.0, step=1.0)
        ml_chant = st.number_input("Ml Rouleau de chant", value=0.0, step=1.0)

    with col_mat2:
        nb_charnieres = st.number_input("Nb Charnières + embases", value=0, step=1)
        nb_tiroirs = st.number_input("Nb Tiroirs (forfait)", value=0, step=1)
        forfait_materiel = st.number_input("Fourniture matériels divers (€)", value=0.0, step=10.0)

    achats_mat = (m2_caissons * 7.80) + (m2_facades * 15.00) + (m2_fond * 8.80) + (ml_chant * 0.65) + (nb_charnieres * 5.00) + (nb_tiroirs * 170.00) + forfait_materiel
    ventes_mat = (m2_caissons * 12.48) + (m2_facades * 24.00) + (m2_fond * 14.08) + (ml_chant * 1.04) + (nb_charnieres * 8.00) + (nb_tiroirs * 187.00) + forfait_materiel

    st.subheader("2. Main d'Œuvre & Déplacements")
    c_mo1, c_mo2, c_mo3 = st.columns(3)
    
    with c_mo1:
        h_etude = st.number_input("Heures Étude (70€/h)", value=0.0, step=0.5)
        h_fab = st.number_input("Heures M.O Fab (75€/h)", value=0.0, step=0.5)
    with c_mo2:
        h_pose = st.number_input("Heures M.O Pose (75€/h)", value=0.0, step=0.5)
        h_chargement = st.number_input("Heures M.O Chargement (75€/h)", value=0.0, step=0.5)
    with c_mo3:
        h_depl_pose = st.number_input("Heures Déplacement Pose (75€/h)", value=0.0, step=0.5)
        km_forfait = st.number_input("Distance Kilométrique (0.606 €/km)", value=0.0, step=10.0)

    ventes_mo = (h_etude * 70.0) + ((h_fab + h_pose + h_chargement + h_depl_pose) * 75.0) + (km_forfait * 0.606)
    total_heures = h_etude + h_fab + h_pose + h_chargement + h_depl_pose

    total_ca = ventes_mat + ventes_mo
    ca_moins_achats = total_ca - achats_mat
    pct_materiaux = (achats_mat / total_ca * 100) if total_ca > 0 else 0.0
    taux_horaire_effectif = (ca_moins_achats / total_heures) if total_heures > 0 else 0.0

    st.markdown("---")
    st.subheader("📊 Ratios & Synthèse du Devis")
    
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total CA Devis HT", f"{total_ca:,.2f} €".replace(",", " "))
    k2.metric("Achats Matériaux Brut", f"{achats_mat:,.2f} €".replace(",", " "))
    k3.metric("% Matériaux / CA", f"{pct_materiaux:.1f} %")
    k4.metric("Taux Horaire Réalisé", f"{taux_horaire_effectif:.2f} €/h", delta=f"{taux_horaire_effectif - 57.0:.2f} €/h vs réf 57€")

# ==============================================================================
# MODULE 3 : PENSE-BÊTE CHANTIERS
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
# MODULE 4 : TRÉSORERIE & ENGAGEMENTS
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
# MODULE 5 : REMISES DE CHÈQUES
# ==============================================================================
elif menu == "🏦 Remises de Chèques":
    st.title("🏦 Gestion des Remises de Chèques")
    st.info("Saisie des chèques reçus et génération de bordereaux de remise.")

    st.sidebar.header("➕ Saisir un Chèque")
    with st.sidebar.form("form_cheque"):
        nom_client = st.text_input("Nom du Client")
        num_cheque = st.text_input("N° de Chèque")
        num_remise = st.text_input("N° de la Remise", value="REM-2026-01")
        date_remise = st.date_input("Date de la Remise", datetime.today())
        montant_cheque = st.number_input("Montant du Chèque (€)", min_value=0.0, value=100.0, step=50.0)
        valider_cheque = st.form_submit_button("Ajouter à la remise")

    if valider_cheque:
        st.session_state['remises_cheques'].append({
            "Nom du client": nom_client,
            "N° de cheque": num_cheque,
            "N° de la remise": num_remise,
            "Date de la remise": str(date_remise),
            "Montant du chêque": montant_cheque
        })
        st.sidebar.success("Chèque enregistré !")

    df_cheques = pd.DataFrame(st.session_state['remises_cheques'])
    
    total_cheques = df_cheques['Montant du chêque'].sum() if not df_cheques.empty else 0.0
    nb_cheques = len(df_cheques)

    rc1, rc2 = st.columns(2)
    rc1.metric("Nombre de Chèques en Attente", f"{nb_cheques}")
    rc2.metric("Total Général Remise", f"{total_cheques:,.2f} €".replace(",", " "))

    if not df_cheques.empty:
        st.subheader("📋 Liste des chèques enregistrés")
        st.dataframe(df_cheques, use_container_width=True)

# ==============================================================================
# MODULE 6 : STOCK & QR CODES
# ==============================================================================
elif menu == "📦 Stock & QR Codes":
    st.title("📦 Gestion des Stocks (Données Bilan 2025)")

    df_stock = pd.DataFrame(st.session_state['stock_articles'])
    
    col1, col2 = st.columns([3, 1])
    col1.write(f"Total des références en base : **{len(df_stock)} articles**")

    pdf_etiquettes = generer_pdf_etiquettes(df_stock)
    col2.download_button(
        label="📄 Imprimer Épreuves Étiquettes (21/feuille)",
        data=pdf_etiquettes,
        file_name="epreuves_etiquettes_quincaillerie.pdf",
        mime="application/pdf",
        icon="🖨️"
    )

    st.dataframe(df_stock, use_container_width=True)
