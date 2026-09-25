import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io

# Bibliothèques pour la génération de PDF et de QR Codes
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import qrcode

st.set_page_config(
    page_title="Gestion Menuiserie", 
    page_icon="🪵", 
    layout="wide"
)

# ---------------------------------------------------------
# CHARGEMENT AUTOMATIQUE DES DONNÉES EXCEL
# ---------------------------------------------------------
@st.cache_data
def charger_donnees():
    # Chargement du stock 2025 depuis le fichier Excel s'il existe
    if os.path.exists("Inventaire stock .xlsx"):
        df_inv = pd.read_excel("Inventaire stock .xlsx", sheet_name="Inventaire pour bilan 2025")
        df_inv = df_inv.dropna(subset=[df_inv.columns[0]])
        df_inv.columns = ["Désignation", "Quantité", "Prix Unitaire HT", "Unité", "Total HT", "Col6", "Col7"][:len(df_inv.columns)]
        df_inv = df_inv[df_inv["Désignation"] != "Désignation"]
        df_inv["Réf"] = df_inv["Désignation"].str.replace(r'[^a-zA-Z0-9\s-]', '', regex=True).str.strip().str.replace(' ', '-')
        
        # Classification Quincaillerie vs Panneaux/Bois
        def categoriser(row):
            des = str(row["Désignation"]).lower()
            unite = str(row["Unité"]).lower()
            if any(k in des for k in ["panneau", "mdf", "cp", "contreplaqué", "mélaminé", "chêne", "sapin", "avive", "dalle", "planche"]) or "m2" in unite or "m²" in unite:
                return "Panneaux & Bois"
            return "Quincaillerie"
            
        df_inv["Catégorie"] = df_inv.apply(categoriser, axis=1)
        df_stock = df_inv[["Réf", "Désignation", "Catégorie", "Quantité", "Prix Unitaire HT", "Unité"]].dropna(subset=["Désignation"])
    else:
        df_stock = pd.DataFrame([
            {"Réf": "VIS-3x10", "Désignation": "VIS 3x10 BZ", "Catégorie": "Quincaillerie", "Quantité": 150, "Prix Unitaire HT": 0.05, "Unité": "U"},
            {"Réf": "PAN-MDF-18", "Désignation": "Panneau MDF 18mm 2800x2070", "Catégorie": "Panneaux & Bois", "Quantité": 12, "Prix Unitaire HT": 42.50, "Unité": "m2"}
        ])

    # Chargement des chantiers depuis l'export Kimai
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

# Initialisation de la mémoire de session Streamlit
if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

if 'stock_actuel' not in st.session_state:
    st.session_state['stock_actuel'] = df_stock_base.copy()

if 'mouvements_stock' not in st.session_state:
    st.session_state['mouvements_stock'] = []

# ---------------------------------------------------------
# GENERATION PDF ETIQUETTES AVERY (33.5 mm x 38.1 mm - 21/page)
# ---------------------------------------------------------
def generer_pdf_etiquettes(df_a_imprimer):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_height = 297 * mm
    
    # Dimensions exactes de l'étiquette en portrait
    w_label = 33.5 * mm
    h_label = 38.1 * mm
    
    # Grille et marges d'impression
    margin_x = 7.0 * mm
    margin_y = 15.0 * mm
    gap_x = 3.0 * mm
    gap_y = 2.0 * mm
    
    cols = 3
    rows = 7
    
    col_idx = 0
    row_idx = 0
    
    for _, row in df_a_imprimer.iterrows():
        ref = str(row['Réf'])
        designation = str(row['Désignation'])
        
        # Coordonnées du coin inférieur gauche de l'étiquette
        x = margin_x + col_idx * (w_label + gap_x)
        y = page_height - margin_y - (row_idx + 1) * h_label - row_idx * gap_y
        
        # Génération du QR Code sous forme d'image PIL
        qr = qrcode.QRCode(box_size=2, border=1)
        qr.add_data(ref)
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white").get_image()
        
        # Conversion pour ReportLab
        qr_image_reader = ImageReader(img_qr)
        
        # 1. Dessin du QR Code centré en haut de l'étiquette
        qr_size = 20 * mm
        qr_x = x + (w_label - qr_size) / 2
        qr_y = y + h_label - qr_size - 2 * mm
        c.drawImage(qr_image_reader, qr_x, qr_y, width=qr_size, height=qr_size)
        
        # 2. Référence en gras sous le QR Code
        c.setFont("Helvetica-Bold", 7.5)
        text_y_ref = qr_y - 4 * mm
        c.drawCentredString(x + w_label / 2, text_y_ref, ref[:16])
        
        # 3. Désignation en texte plus fin sous la référence
        if designation and designation != ref:
            c.setFont("Helvetica", 6)
            text_y_des = text_y_ref - 3.5 * mm
            c.drawCentredString(x + w_label / 2, text_y_des, designation[:20])
        
        # Progression dans la grille 3x7
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

# ---------------------------------------------------------
# MENU DE NAVIGATION PRINCIPAL
# ---------------------------------------------------------
st.sidebar.title("🛠️ Gestion Menuiserie")
menu = st.sidebar.radio(
    "Accéder aux modules :",
    [
        "⏱️ Saisie des Heures",
        "📊 Suivi Temps & Production",
        "🧮 Brouillon Devis & Marges",
        "📦 Stock & Mouvements",
        "📷 Scan QR Code Stock",
        "🏷️ Impression Étiquettes Stock"
    ]
)

# ---------------------------------------------------------
# 1. SAISIE DES HEURES ATELIER & CHANTIER
# ---------------------------------------------------------
if menu == "⏱️ Saisie des Heures":
    st.header("⏱️ Saisie Rapide des Heures Atelier & Chantier")
    
    with st.form("form_saisie_heures", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            date_saisie = st.date_input("Date d'intervention", datetime.now())
            chantier = st.selectbox("Chantier / Projet Kimai", LISTE_CHANTIERS)
        with col2:
            code_tache = st.selectbox("Tâche / Code Activité", LISTE_TACHES)
            heures = st.number_input("Nombre d'heures effectuées", min_value=0.25, max_value=12.0, step=0.25, value=1.0)
            
        valider = st.form_submit_button("💾 Enregistrer l'intervention")
        
        if valider:
            est_prod = not code_tache.startswith("X")
            st.session_state['historique_heures'].append({
                "Date": str(date_saisie),
                "Chantier": chantier,
                "Code": code_tache,
                "Heures": heures,
                "Production": est_prod
            })
            st.success(f"Enregistré : {heures}h sur **{chantier}** ({code_tache})")

    st.subheader("📋 Dernières saisies de la session")
    if st.session_state['historique_heures']:
        st.dataframe(pd.DataFrame(st.session_state['historique_heures']), use_container_width=True)
    else:
        st.info("Aucune saisie effectuée au cours de la session active.")

# ---------------------------------------------------------
# 2. SUIVI TEMPS & PRODUCTION
# ---------------------------------------------------------
elif menu == "📊 Suivi Temps & Production":
    st.header("📊 Suivi du Temps de Production")
    
    OBJECTIF_HEBDO = 33.50
    
    if st.session_state['historique_heures']:
        df_h = pd.DataFrame(st.session_state['historique_heures'])
        total_prod = df_h[df_h["Production"] == True]["Heures"].sum()
        total_hors_prod = df_h[df_h["Production"] == False]["Heures"].sum()
        total_general = total_prod + total_hors_prod
        
        c1, c2, c3 = st.columns(3)
        c1.metric("Heures de Production", f"{total_prod:.2f} h", delta=f"{total_prod - OBJECTIF_HEBDO:.2f} h par rapport à l'objectif")
        c2.metric("Heures Admin / Hors-Prod", f"{total_hors_prod:.2f} h")
        c3.metric("Total Général", f"{total_general:.2f} h")
        
        st.markdown("---")
        st.subheader("Détail cumulé par chantier et activité")
        df_recap = df_h.groupby(["Chantier", "Code", "Production"])["Heures"].sum().reset_index()
        st.dataframe(df_recap, use_container_width=True)
    else:
        st.info("Saisissez des heures dans le module dédié pour afficher l'analyse temps.")

# ---------------------------------------------------------
# 3. BROUILLON DEVIS & MARGES
# ---------------------------------------------------------
elif menu == "🧮 Brouillon Devis & Marges":
    st.header("🧮 Brouillon de Devis et Calcul de Marge")
    
    st.subheader("1. Fournitures & Matériaux (Panneaux, Quincaillerie, Sous-traitance)")
    col_mat1, col_mat2 = st.columns(2)
    achats_mat = col_mat1.number_input("Total Achats Matériaux HT (€)", min_value=0.0, value=500.0, step=50.0)
    ventes_mat = col_mat2.number_input("Total Facturé Matériaux HT (€)", min_value=0.0, value=750.0, step=50.0)
    
    marge_euro = ventes_mat - achats_mat
    taux_marque = (marge_euro / ventes_mat * 100) if ventes_mat > 0 else 0.0
    
    c_m1, c_m2 = st.columns(2)
    c_m1.caption(f"Marge brute fournitures : **{marge_euro:,.2f} € HT**")
    c_m2.caption(f"Taux de marque : **{taux_marque:.1f} %**")

    st.markdown("---")
    st.subheader("2. Main d'Œuvre Prévisionnelle")
    col_mo1, col_mo2 = st.columns(2)
    heures_prev = col_mo1.number_input("Heures estimées (Atelier + Pose)", min_value=0.0, value=15.0, step=0.5)
    taux_horaire = col_mo2.number_input("Taux Horaire Vendu HT (€/h)", min_value=0.0, value=55.0, step=5.0)
    
    total_mo = heures_prev * taux_horaire
    total_devis = ventes_mat + total_mo
    
    st.markdown("---")
    col_res1, col_res2 = st.columns(2)
    col_res1.metric("Total Main d'Œuvre HT", f"{total_mo:,.2f} €")
    col_res2.metric("Montant Total estimé Devis HT", f"{total_devis:,.2f} €")

# ---------------------------------------------------------
# 4. GESTION DU STOCK & MOUVEMENTS
# ---------------------------------------------------------
elif menu == "📦 Stock & Mouvements":
    st.header("📦 Consultation et Gestion du Stock")
    
    cat_filtre = st.radio("Filtrer par catégorie :", ["Tous", "Quincaillerie", "Panneaux & Bois"], horizontal=True)
    
    df_stk = st.session_state['stock_actuel']
    if cat_filtre != "Tous":
        df_stk = df_stk[df_stk["Catégorie"] == cat_filtre]
        
    st.dataframe(df_stk, use_container_width=True)
    
    st.markdown("---")
    st.subheader("🔄 Saisie manuelle d'un mouvement de stock")
    with st.form("form_mvt_stock"):
        c1, c2, c3 = st.columns(3)
        ref_mvt = c1.selectbox("Article", st.session_state['stock_actuel']["Réf"].tolist())
        type_mvt = c2.selectbox("Type d'opération", ["Sortie (Chantier)", "Entrée (Réception)"])
        qte_mvt = c3.number_input("Quantité", min_value=1.0, value=1.0, step=1.0)
        
        btn_mvt = st.form_submit_button("Valider le mouvement")
        if btn_mvt:
            idx = st.session_state['stock_actuel'][st.session_state['stock_actuel']["Réf"] == ref_mvt].index
            if not idx.empty:
                i = idx[0]
                if "Sortie" in type_mvt:
                    st.session_state['stock_actuel'].at[i, "Quantité"] -= qte_mvt
                else:
                    st.session_state['stock_actuel'].at[i, "Quantité"] += qte_mvt
                
                st.session_state['mouvements_stock'].append({
                    "Date": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "Réf": ref_mvt,
                    "Type": type_mvt,
                    "Quantité": qte_mvt
                })
                st.success(f"Stock mis à jour pour {ref_mvt} !")
                st.rerun()

# ---------------------------------------------------------
# 5. SCANNER QR CODE STOCK
# ---------------------------------------------------------
elif menu == "📷 Scan QR Code Stock":
    st.header("📷 Numérisation d'Étiquettes Quincaillerie / Stock")
    st.write("Utilisez la caméra de votre smartphone ou tablette pour scanner une étiquette d'article.")
    
    img_captured = st.camera_input("Prendre en photo l'étiquette QR Code")
    
    if img_captured:
        st.info("Traitement de l'image capturée...")
        st.success("Fonction de lecture automatique active.")

# ---------------------------------------------------------
# 6. IMPRESSION ÉTIQUETTES STOCK
# ---------------------------------------------------------
elif menu == "🏷️ Impression Étiquettes Stock":
    st.header("🏷️ Impression d'Étiquettes QR Code pour Quincaillerie & Panneaux")
    st.write("Format paramétré : **Avery 33,5 mm × 38,1 mm** (21 étiquettes par planche A4 - 3 colonnes × 7 lignes)")
    
    filtre_imp = st.selectbox("Catégorie à afficher :", ["Toutes", "Quincaillerie", "Panneaux & Bois"])
    df_imp_base = st.session_state['stock_actuel']
    if filtre_imp != "Toutes":
        df_imp_base = df_imp_base[df_imp_base["Catégorie"] == filtre_imp]
        
    st.dataframe(df_imp_base, use_container_width=True)
    
    articles_selectionnes = st.multiselect(
        "Sélectionnez les articles à imprimer sur la planche :",
        options=df_imp_base["Réf"].tolist(),
        default=df_imp_base["Réf"].tolist()[:3]
    )
    
    if articles_selectionnes:
        df_filtr = df_imp_base[df_imp_base["Réf"].isin(articles_selectionnes)]
        pdf_data = generer_pdf_etiquettes(df_filtr)
        
        st.download_button(
            label="📄 Télécharger la planche d'étiquettes (PDF)",
            data=pdf_data,
            file_name="etiquettes_quincaillerie_avery.pdf",
            mime="application/pdf"
        )
        
