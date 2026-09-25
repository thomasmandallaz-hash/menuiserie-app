import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import os
import io

# Import des bibliothèques pour l'impression et les QR Codes
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import qrcode

st.set_page_config(page_title="Gestion Menuiserie", layout="wide")

# --- CHARGEMENT AUTOMATIQUE DES DONNÉES EXCEL ---
@st.cache_data
def charger_donnees():
    # Chargement du stock
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

# Initialisation de la mémoire session
if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

# --- FONCTION GÉNÉRATION PDF ÉTIQUETTES (38.1 x 33.5 mm) ---
def generer_pdf_etiquettes(df_a_imprimer):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    
    # Dimensions de la planche
    w_label = 38.1 * mm
    h_label = 33.5 * mm
    margin_x = 7.0 * mm
    margin_y = 10.0 * mm
    gap_x = 3.0 * mm
    gap_y = 2.0 * mm
    
    cols = 3
    rows = 7
    page_height = 297 * mm
    
    col_idx = 0
    row_idx = 0
    
    for _, row in df_a_imprimer.iterrows():
        ref = str(row['Réf'])
        designation = str(row['Désignation'])
        
        # Position X et Y sur la page
        x = margin_x + col_idx * (w_label + gap_x)
        y = page_height - margin_y - (row_idx + 1) * h_label - row_idx * gap_y
        
        # Génération du QR code
        qr = qrcode.QRCode(box_size=2, border=1)
        qr.add_data(ref)
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white")
        
        qr_buffer = io.BytesIO()
        img_qr.save(qr_buffer, format='PNG')
        qr_buffer.seek(0)
        
        # Dessin sur le PDF : QR code en haut centré
        qr_size = 18 * mm
        qr_x = x + (w_label - qr_size) / 2
        qr_y = y + h_label - qr_size - 2 * mm
        c.drawInlineImage(qr_buffer, qr_x, qr_y, width=qr_size, height=qr_size)
        
        # Impression des textes sous le QR Code
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x + w_label / 2, y + 9 * mm, ref[:18])
        
        c.setFont("Helvetica", 6)
        c.drawCentredString(x + w_label / 2, y + 4 * mm, designation[:25])
        
        # Passages aux cases suivantes
        col_idx += 1
        if col_idx >= cols:
            col_idx = 0
            row_idx += 1
            if row_idx >= rows:
                row_idx = 0
                c.showPage()
                
    c.save()
    buffer.seek(0)
    return buffer

# --- NAVIGATION ---
st.sidebar.title("🛠️ Menu Principal")
menu = st.sidebar.radio(
    "Accéder à :",
    [
        "⏱️ Saisie des Heures",
        "📊 Suivi Temps & Production",
        "🧮 Brouillon Devis & Marges",
        "🏷️ Impression Étiquettes Stock"
    ]
)

# ---------------------------------------------------------
# 1. SAISIE DES HEURES
# ---------------------------------------------------------
if menu == "⏱️ Saisie des Heures":
    st.header("⏱️ Saisie Rapide des Heures Atelier & Chantier")
    
    with st.form("form_saisie_heures"):
        col1, col2 = st.columns(2)
        with col1:
            date_saisie = st.date_input("Date", datetime.now())
            chantier = st.selectbox("Chantier / Projet", LISTE_CHANTIERS)
        with col2:
            code_tache = st.selectbox("Tâche / Activité", LISTE_TACHES)
            heures = st.number_input("Nombre d'heures", min_value=0.25, max_value=12.0, step=0.25, value=1.0)
            
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
            st.success(f"Enregistré : {heures}h sur {chantier} ({code_tache})")

    st.subheader("Dernières saisies de la session")
    if st.session_state['historique_heures']:
        st.dataframe(pd.DataFrame(st.session_state['historique_heures']), use_container_width=True)

# ---------------------------------------------------------
# 2. SUIVI TEMPS & PRODUCTION
# ---------------------------------------------------------
elif menu == "📊 Suivi Temps & Production":
    st.header("📊 Suivi du Temps de Production")
    
    # Target horaire de production
    OBJECTIF_HEBDO = 33.50
    
    if st.session_state['historique_heures']:
        df_h = pd.DataFrame(st.session_state['historique_heures'])
        total_prod = df_h[df_h["Production"] == True]["Heures"].sum()
        total_hors_prod = df_h[df_h["Production"] == False]["Heures"].sum()
        total_general = total_prod + total_hors_prod
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Heures de Production", f"{total_prod:.2f} h", delta=f"{total_prod - OBJECTIF_HEBDO:.2f} h / objectif")
        c2.metric("Heures Admin / Hors-Prod", f"{total_hors_prod:.2f} h")
        c3.metric("Total Général", f"{total_general:.2f} h")
        
        st.subheader("Détail par chantier")
        st.dataframe(df_h.groupby(["Chantier", "Code"])["Heures"].sum().reset_index(), use_container_width=True)
    else:
        st.info("Aucune heure saisie pour le moment dans cette session.")

# ---------------------------------------------------------
# 3. BROUILLON DEVIS & MARGES
# ---------------------------------------------------------
elif menu == "🧮 Brouillon Devis & Marges":
    st.header("🧮 Brouillon de Devis et Calcul de Marge")
    
    st.subheader("1. Fournitures & Matériaux")
    col_mat1, col_mat2 = st.columns(2)
    achats_mat = col_mat1.number_input("Total Achats Matériaux HT (€)", min_value=0.0, value=500.0, step=50.0)
    ventes_mat = col_mat2.number_input("Total Facturé Matériaux HT (€)", min_value=0.0, value=750.0, step=50.0)
    
    marge_euro = ventes_mat - achats_mat
    taux_marque = (marge_euro / ventes_mat * 100) if ventes_mat > 0 else 0.0
    
    c_m1, c_m2 = st.columns(2)
    c_m1.caption(f"Marge brute fournitures : **{marge_euro:,.2f} € HT**")
    c_m2.caption(f"Taux de marque : **{taux_marque:.1f} %**")

    st.subheader("2. Main d'Œuvre Prévisionnelle")
    col_mo1, col_mo2 = st.columns(2)
    heures_prev = col_mo1.number_input("Heures estimées (Atelier + Pose)", min_value=0.0, value=15.0, step=0.5)
    taux_horaire = col_mo2.number_input("Taux Horaire Vendu HT (€/h)", min_value=0.0, value=55.0, step=5.0)
    
    total_mo = heures_prev * taux_horaire
    total_devis = ventes_mat + total_mo
    
    st.markdown("---")
    st.subheader(f"Total Général Estimé du Devis : **{total_devis:,.2f} € HT**")

# ---------------------------------------------------------
# 4. IMPRESSION ÉTIQUETTES STOCK
# ---------------------------------------------------------
elif menu == "🏷️ Impression Étiquettes Stock":
    st.header("🏷️ Impression d'Étiquettes QR Code pour Quincaillerie")
    st.write("Format : **Avery 33,5 mm × 38,1 mm** (21 étiquettes par page - 3 colonnes × 7 lignes)")
    
    st.dataframe(df_stock_base, use_container_width=True)
    
    articles_selectionnes = st.multiselect(
        "Sélectionnez les articles à imprimer :",
        options=df_stock_base["Réf"].tolist(),
        default=df_stock_base["Réf"].tolist()[:3]
    )
    
    if articles_selectionnes:
        df_filtr = df_stock_base[df_stock_base["Réf"].isin(articles_selectionnes)]
        pdf_data = generer_pdf_etiquettes(df_filtr)
        
        st.download_button(
            label="📄 Télécharger la planche d'étiquettes PDF",
            data=pdf_data,
            file_name="etiquettes_stock_avery.pdf",
            mime="application/pdf"
        )
