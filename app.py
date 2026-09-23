import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
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
        "🔎 Synthèse par Chantier",
        "📅 Lundi Administratif & Devis",
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

if 'taches_lundi' not in st.session_state:
    st.session_state['taches_lundi'] = [
        {"Tâche / Devis": "Relance devis en attente", "Echéance": str(datetime.today().date()), "Fait": False},
        {"Tâche / Devis": "Métré & Rédaction devis cuisine", "Echéance": str(datetime.today().date()), "Fait": False},
        {"Tâche / Devis": "Facturation chantiers terminés", "Echéance": str(datetime.today().date()), "Fait": False}
    ]

# Modèles par défaut pour le brouillon devis
if 'devis_fournitures' not in st.session_state:
    st.session_state['devis_fournitures'] = pd.DataFrame([
        {"Désignation": "Panneaux de caissons", "Quantité": 0.0, "Unité": "m²", "Prix d'achat HT": 7.80, "Prix de vente HT": 12.48},
        {"Désignation": "Panneaux de façade", "Quantité": 0.0, "Unité": "m²", "Prix d'achat HT": 15.00, "Prix de vente HT": 24.00},
        {"Désignation": "Panneau de fond", "Quantité": 0.0, "Unité": "m²", "Prix d'achat HT": 8.80, "Prix de vente HT": 14.08},
        {"Désignation": "Rouleau de chant", "Quantité": 0.0, "Unité": "ml", "Prix d'achat HT": 0.65, "Prix de vente HT": 1.04},
        {"Désignation": "Charnières + embases", "Quantité": 0.0, "Unité": "U", "Prix d'achat HT": 5.00, "Prix de vente HT": 8.00},
        {"Désignation": "Tiroirs (forfait)", "Quantité": 0.0, "Unité": "U", "Prix d'achat HT": 170.00, "Prix de vente HT": 187.00}
    ])

if 'devis_mo' not in st.session_state:
    st.session_state['devis_mo'] = pd.DataFrame([
        {"Poste / Tâche": "Heures Étude & Devis (X8)", "Heures": 0.0, "Taux Horaire Vente (€/h)": 70.0},
        {"Poste / Tâche": "Heures Fabrication Atelier (M1/D2)", "Heures": 0.0, "Taux Horaire Vente (€/h)": 75.0},
        {"Poste / Tâche": "Heures Pose Chantier (P1)", "Heures": 0.0, "Taux Horaire Vente (€/h)": 75.0},
        {"Poste / Tâche": "Heures Chargement / Préparation", "Heures": 0.0, "Taux Horaire Vente (€/h)": 75.0},
        {"Poste / Tâche": "Heures Déplacement Pose", "Heures": 0.0, "Taux Horaire Vente (€/h)": 75.0}
    ])

# FONCTION GENERATION PDF
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

        qr = qrcode.QRCode(box_size=10, border=1)
        qr.add_data(str(row['Réf']))
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white").get_image()
        
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
# MODULE SYNTHÈSE PAR CHANTIER
# ==============================================================================
if menu == "🔎 Synthèse par Chantier":
    st.title("🔎 Synthèse Globale Récapitulative par Chantier")
    
    chantier_sel = st.selectbox("🎯 Sélectionner le chantier à analyser :", LISTE_CHANTIERS)

    # Filtrage des heures réelles du chantier
    df_heures = pd.DataFrame(st.session_state['historique_heures'])
    if not df_heures.empty and 'Chantier' in df_heures.columns:
        df_heures_chantier = df_heures[df_heures['Chantier'] == chantier_sel]
    else:
        df_heures_chantier = pd.DataFrame()

    heures_reelles_prod = df_heures_chantier[df_heures_chantier['Production'] == True]['Heures'].sum() if not df_heures_chantier.empty else 0.0
    heures_reelles_totales = df_heures_chantier['Heures'].sum() if not df_heures_chantier.empty else 0.0

    # Données du brouillon devis actuel
    df_fourn = st.session_state['devis_fournitures']
    df_mo = st.session_state['devis_mo']
    
    achats_mat = (df_fourn["Quantité"] * df_fourn["Prix d'achat HT"]).sum()
    ventes_mat = (df_fourn["Quantité"] * df_fourn["Prix de vente HT"]).sum()
    heures_devises = df_mo["Heures"].sum()
    ventes_mo = (df_mo["Heures"] * df_mo["Taux Horaire Vente (€/h)"]).sum()

    total_devis_ca = ventes_mat + ventes_mo
    marge_brute_mat = ventes_mat - achats_mat

    # Création des onglets récapitulatifs par chantier
    tab_marge, tab_devis, tab_heures, tab_memo = st.tabs([
        "📊 Rentabilité & Comparatif Prévu / Réel",
        "🧮 Devis & Chiffrage associé",
        "⏱️ Historique des Heures Réelles",
        "📝 Pense-bête & Remarques"
    ])

    with tab_marge:
        st.subheader("📊 Performance & Marges du Chantier")
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("CA Devisé HT", f"{total_devis_ca:,.2f} €".replace(",", " "))
        c2.metric("Heures Prévues (Devis)", f"{heures_devises:.1f} h")
        c3.metric("Heures Réelles Passées", f"{heures_reelles_totales:.1f} h", delta=f"{heures_devises - heures_reelles_totales:.1f} h solde devis")
        
        taux_realise = ((total_devis_ca - achats_mat) / heures_reelles_totales) if heures_reelles_totales > 0 else 0.0
        c4.metric("Taux Horaire Réel Dégagé", f"{taux_realise:.2f} €/h")

        if heures_reelles_totales > heures_devises and heures_devises > 0:
            st.error("⚠️ Attention : Le temps réel passé dépasse le nombre d'heures prévues au devis !")
        elif heures_reelles_totales > 0:
            st.success("🟢 Chantier sous contrôle au niveau du temps passé.")

    with tab_devis:
        st.subheader("🧮 Détail du Chiffrage / Devis du Chantier")
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            st.markdown("**Fournitures & Matériaux**")
            st.dataframe(df_fourn[df_fourn['Quantité'] > 0], use_container_width=True)
        with col_m2:
            st.markdown("**Main d'Œuvre & Prestations**")
            st.dataframe(df_mo[df_mo['Heures'] > 0], use_container_width=True)

    with tab_heures:
        st.subheader("⏱️ Détail des Pointages / Saisies")
        if not df_heures_chantier.empty:
            st.dataframe(df_heures_chantier, use_container_width=True)
        else:
            st.info("Aucune heure enregistrée pour ce chantier pour le moment.")

    with tab_memo:
        st.subheader("📝 Remarques & Mémos Chantier")
        note_ch = st.session_state['pense_bete'].get(chantier_sel, "Aucune note saisie pour ce chantier.")
        st.text_area("Note enregistrée :", value=note_ch, disabled=True, height=120)

# ==============================================================================
# MODULE 0 : LUNDI ADMINISTRATIF & DEVIS
# ==============================================================================
elif menu == "📅 Lundi Administratif & Devis":
    st.title("📅 Lundi Administratif — Devis & Tâches")
    st.info("Espace dédié à la journée du lundi : suivi des devis à effectuer, relances et tâches administratives. Les tâches non réalisées se répercutent sur le lundi suivant.")

    col_saisie, col_actions = st.columns([2, 1])

    with col_saisie:
        st.subheader("➕ Ajouter une tâche / un devis à réaliser")
        with st.form("form_lundi"):
            nouvelle_tache = st.text_input("Tâche ou Devis à traiter")
            date_ech = st.date_input("Échéance", datetime.today())
            if st.form_submit_button("Ajouter à la liste"):
                if nouvelle_tache:
                    st.session_state['taches_lundi'].append({
                        "Tâche / Devis": nouvelle_tache,
                        "Echéance": str(date_ech),
                        "Fait": False
                    })
                    st.success("Tâche ajoutée !")

    with col_actions:
        st.subheader("🔄 Report Automatique")
        if st.button("Reporter les tâches non cochées au lundi suivant"):
            prochain_lundi = datetime.today() + timedelta(days=(7 - datetime.today().weekday()))
            count = 0
            for item in st.session_state['taches_lundi']:
                if not item['Fait']:
                    item['Echéance'] = str(prochain_lundi.date())
                    count += 1
            st.success(f"{count} tâche(s) reportée(s) au {prochain_lundi.strftime('%d/%m/%Y')} !")

    st.markdown("---")
    st.subheader("📋 Liste des travaux administratifs du lundi")

    df_lundi = pd.DataFrame(st.session_state['taches_lundi'])
    if not df_lundi.empty:
        edited_df = st.data_editor(
            df_lundi,
            column_config={
                "Fait": st.column_config.CheckboxColumn("Statut", default=False),
                "Tâche / Devis": st.column_config.TextColumn("Description", width="large"),
                "Echéance": st.column_config.DateColumn("Date Échéance")
            },
            disabled=["Echéance"],
            num_rows="dynamic",
            use_container_width=True
        )
        st.session_state['taches_lundi'] = edited_df.to_dict('records')
    else:
        st.write("Aucune tâche en cours.")

# ==============================================================================
# MODULE 1 : RENTABILITÉ & PLANNING
# ==============================================================================
elif menu == "📊 Rentabilité & Planning":
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
# MODULE 2 : BROUILLON DEVIS & MARGES (DYNAMIQUE)
# ==============================================================================
elif menu == "🧮 Brouillon Devis & Marges":
    st.title("🧮 Brouillon de Devis & Calculateur de Marges")
    st.info("Vous pouvez ajouter, supprimer ou modifier directement toutes les lignes de fournitures et de main-d'œuvre dans les tableaux ci-dessous.")

    st.subheader("1. Fournitures & Matériaux (Modifiable)")
    df_fourn_edited = st.data_editor(
        st.session_state['devis_fournitures'],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Quantité": st.column_config.NumberColumn(min_value=0.0, step=1.0),
            "Prix d'achat HT": st.column_config.NumberColumn(format="%.2f €"),
            "Prix de vente HT": st.column_config.NumberColumn(format="%.2f €")
        }
    )
    st.session_state['devis_fournitures'] = df_fourn_edited

    achats_mat = (df_fourn_edited["Quantité"] * df_fourn_edited["Prix d'achat HT"]).sum()
    ventes_mat = (df_fourn_edited["Quantité"] * df_fourn_edited["Prix de vente HT"]).sum()

    st.subheader("2. Main d'Œuvre & Déplacements (Modifiable)")
    df_mo_edited = st.data_editor(
        st.session_state['devis_mo'],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Heures": st.column_config.NumberColumn(min_value=0.0, step=0.5),
            "Taux Horaire Vente (€/h)": st.column_config.NumberColumn(format="%.2f €/h")
        }
    )
    st.session_state['devis_mo'] = df_mo_edited

    km_forfait = st.number_input("Distance Kilométrique (0.606 €/km)", value=0.0, step=10.0)

    total_heures = df_mo_edited["Heures"].sum()
    ventes_mo = (df_mo_edited["Heures"] * df_mo_edited["Taux Horaire Vente (€/h)"]).sum() + (km_forfait * 0.606)

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
