import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import glob

# Bibliothèques PDF et QR Codes
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import qrcode
from PIL import Image
import zxingcpp

st.set_page_config(
    page_title="Gestion Menuiserie", 
    page_icon="🪵", 
    layout="wide"
)

# ---------------------------------------------------------
# REFERENTIEL ACTIVITES
# ---------------------------------------------------------
ACTIVITES_INTEGRALES = [
    "D1 - Appro Débit", "D2 - Débit massif", "D3 - Scie à ruban", "D4 - Scie à format",
    "D5 - Scie à Panneaux", "D6 - Scie à ruban métaux", "D7 - Scie radiale",
    "C1 - Corroyeuse 4 faces", "C2 - Dégauchissage", "C3 - Rabotage",
    "U1 - Toupie", "U2 - Défonceuses", "U3 - Scie à ruban", "U4 - Plaqueuse",
    "U5 - CN (Commande Numérique)", "U6 - Tournage", "U7 - Mortaisage",
    "U8 - Tenonnage", "U9 - Pointage machine", "U10 - Usinage", "U11 - Perçage manuel",
    "B1 - Finition manuelle", "B2 - Ponceuse à bande", "B3 - Placage chant main",
    "B4 - Affleurage chants", "B5 - Vitrage", "B6 - Cassage d'arêtes", "B7 - Brossage",
    "M1 - Montage (cadrage)", "M2 - Montage caisses", "M3 - Collage Px", "M4 - Collage néoprène",
    "X1 - Étude de plan", "X2 - Établissement", "X3 - Traçage", "X4 - Affûtage",
    "X5 - Bureau", "X6 - RDV Clientèle", "X7 - Relevé de cotes", "X8 - Mise en page devis", "X9 - Création publication réseau",
    "V1 - Vernissage pistolet", "V2 - Laquage pistolet", "V3 - Pinceau", "V4 - Teinture", "V5 - Égrenage", "V6 - Traitement IFH",
    "Q1 - Pose quincaillerie",
    "N1 - Nettoyage atelier", "N2 - Rangement atelier", "N3 - Affûtage atelier", "N4 - Changement sac aspi", "N5 - Changement plaquettes",
    "T1 - Trajet", "T2 - Chargement", "T3 - Mise au séchoir", "T4 - Emballage", "T5 - Manutention",
    "P1 - Pose sur chantier", "P2 - Démontage", "P3 - Évacuation", "P4 - Fab gabarits", "P5 - Rangement nettoyage chantier"
]

@st.cache_data
def charger_activites_souche():
    fichier_souche = "Activités.xlsx"
    taches_extraites = []
    if os.path.exists(fichier_souche):
        try:
            xls = pd.ExcelFile(fichier_souche)
            sheet = "Réf" if "Réf" in xls.sheet_names else xls.sheet_names[0]
            df_act = pd.read_excel(xls, sheet_name=sheet)
            for idx, row in df_act.iterrows():
                vals = [str(val).strip() for val in row.values if pd.notna(val) and str(val).strip() != "nan"]
                for i in range(len(vals) - 1):
                    code, intitule = vals[i], vals[i+1]
                    if len(code) <= 4 and code[0].isalpha() and code[1:].isdigit():
                        elem = f"{code} - {intitule}"
                        if elem not in taches_extraites:
                            taches_extraites.append(elem)
        except Exception:
            pass
    return taches_extraites if len(taches_extraites) >= 30 else ACTIVITES_INTEGRALES.copy()

@st.cache_data
def charger_donnees_kimai_heures():
    fichiers_kimai = list(set(glob.glob("kimai-export*.xlsx") + glob.glob("*export*.xlsx")))
    donnees_cumulees = []
    
    for fichier in fichiers_kimai:
        if os.path.exists(fichier) and "activit" not in fichier.lower():
            try:
                df = pd.read_excel(fichier)
                col_nom, col_total = df.columns[0], df.columns[1]
                projet_actuel = "Général"
                
                for idx, row in df.iterrows():
                    val = str(row[col_nom]).strip()
                    try:
                        total_heures = float(str(row[col_total]).replace(',', '.'))
                    except ValueError:
                        total_heures = 0.0
                    
                    if val == "nan" or not val:
                        continue
                    
                    # Détection du nom de chantier
                    if any(val.startswith(p) for p in ["OE ", "25/", "26/", "Agencement"]):
                        projet_actuel = val
                    elif total_heures > 0:
                        donnees_cumulees.append({
                            "Projet": projet_actuel,
                            "Tâche": val.replace('\t', ' - ').strip(),
                            "Heures": total_heures
                        })
            except Exception:
                pass

    df_kimai = pd.DataFrame(donnees_cumulees)
    
    if not df_kimai.empty:
        def est_production(tache):
            t = str(tache).upper()
            return not (t.startswith("X") or "BUREAU" in t or "DEVIS" in t or "RDV" in t)

        df_kimai["Production"] = df_kimai["Tâche"].apply(est_production)
        projets_uniques = sorted(df_kimai["Projet"].unique().tolist())
    else:
        projets_uniques = ["OE 26/10 fabrication meuble enceinte", "26/221 Fabrication et pose d'étagères", "26/227 Réfection plan de travail"]

    return df_kimai, projets_uniques
@st.cache_data
def charger_stock():
    fichiers_inv = glob.glob("*stock*.xlsx") + glob.glob("*Inventaire*.xlsx")
    if fichiers_inv and os.path.exists(fichiers_inv[0]):
        try:
            df_inv = pd.read_excel(fichiers_inv[0], sheet_name="Inventaire pour bilan 2025")
            df_inv = df_inv.dropna(subset=[df_inv.columns[0]])
            df_inv.columns = ["Désignation", "Quantité", "Prix Unitaire HT", "Unité", "Total HT", "Col6", "Col7"][:len(df_inv.columns)]
            df_inv = df_inv[df_inv["Désignation"] != "Désignation"]
            df_inv["Réf"] = df_inv["Désignation"].str.replace(r'[^a-zA-Z0-9\s-]', '', regex=True).str.strip().str.replace(' ', '-')
            
            def categoriser(row):
                des, unite = str(row["Désignation"]).lower(), str(row["Unité"]).lower()
                if any(k in des for k in ["panneau", "mdf", "cp", "contreplaqué", "mélaminé", "chêne", "sapin", "avive", "dalle", "planche"]) or "m2" in unite or "m²" in unite:
                    return "Panneaux & Bois"
                return "Quincaillerie"
                
            df_inv["Catégorie"] = df_inv.apply(categoriser, axis=1)
            return df_inv[["Réf", "Désignation", "Catégorie", "Quantité", "Prix Unitaire HT", "Unité"]].dropna(subset=["Désignation"])
        except Exception:
            pass

    return pd.DataFrame([
        {"Réf": "VIS-3x10", "Désignation": "VIS 3x10 BZ", "Catégorie": "Quincaillerie", "Quantité": 150, "Prix Unitaire HT": 0.05, "Unité": "U"},
        {"Réf": "PAN-MDF-18", "Désignation": "Panneau MDF 18mm 2800x2070", "Catégorie": "Panneaux & Bois", "Quantité": 12, "Prix Unitaire HT": 42.50, "Unité": "m2"}
    ])

# ---------------------------------------------------------
# STRUCTURE DE BASE DE TON FICHIER DEVIS 00 DEVIS.XLSX
# ---------------------------------------------------------
DEVIS_MATRICE_BASE = [
    {"Poste": "Etude", "U": "heures", "Q": 0.0, "Prix Achat HT": 70.0, "Prix Vente HT": 70.0},
    {"Poste": "Panneaux de caissons", "U": "M²", "Q": 0.0, "Prix Achat HT": 7.80, "Prix Vente HT": 12.48},
    {"Poste": "Panneaux de façade", "U": "M²", "Q": 0.0, "Prix Achat HT": 15.00, "Prix Vente HT": 24.00},
    {"Poste": "Panneau de fond", "U": "M²", "Q": 0.0, "Prix Achat HT": 8.80, "Prix Vente HT": 14.08},
    {"Poste": "Rouleau de chant", "U": "Ml", "Q": 0.0, "Prix Achat HT": 0.65, "Prix Vente HT": 1.04},
    {"Poste": "Charnières + embases", "U": "pc", "Q": 0.0, "Prix Achat HT": 5.00, "Prix Vente HT": 8.00},
    {"Poste": "Tiroirs", "U": "Forf pour 1", "Q": 0.0, "Prix Achat HT": 170.0, "Prix Vente HT": 187.00},
    {"Poste": "Barre à penderie", "U": "pc", "Q": 0.0, "Prix Achat HT": 0.0, "Prix Vente HT": 0.0},
    {"Poste": "Fourniture materiel", "U": "Forf", "Q": 0.0, "Prix Achat HT": 0.0, "Prix Vente HT": 0.0},
    {"Poste": "M.O Fab", "U": "heures", "Q": 0.0, "Prix Achat HT": 75.0, "Prix Vente HT": 75.0},
    {"Poste": "M.O pose", "U": "heures", "Q": 0.0, "Prix Achat HT": 75.0, "Prix Vente HT": 75.0},
    {"Poste": "M.O chargement", "U": "heures", "Q": 0.0, "Prix Achat HT": 75.0, "Prix Vente HT": 75.0},
    {"Poste": "Deplacement pour pose", "U": "heures", "Q": 0.0, "Prix Achat HT": 75.0, "Prix Vente HT": 75.0},
    {"Poste": "Forfait kilometrique", "U": "Km", "Q": 0.0, "Prix Achat HT": 0.606, "Prix Vente HT": 0.606},
    {"Poste": "Relevé de cote", "U": "heures", "Q": 0.0, "Prix Achat HT": 70.0, "Prix Vente HT": 70.0},
    {"Poste": "Intendance (appel téléphonique etc)", "U": "forf", "Q": 0.0, "Prix Achat HT": 0.0, "Prix Vente HT": 0.0},
    {"Poste": "Dechetterie", "U": "% perte", "Q": 0.0, "Prix Achat HT": 0.0, "Prix Vente HT": 0.0},
    {"Poste": "Devis", "U": "forf", "Q": 0.0, "Prix Achat HT": 70.0, "Prix Vente HT": 0.0}
]

# Initialisations
LISTE_TACHES_SOUCHE = charger_activites_souche()
DF_KIMAI_HISTO, PROJETS_KIMAI = charger_donnees_kimai_heures()
DF_STOCK_BASE = charger_stock()

if 'liste_chantiers' not in st.session_state:
    st.session_state['liste_chantiers'] = PROJETS_KIMAI.copy()

if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

if 'liste_taches' not in st.session_state:
    st.session_state['liste_taches'] = LISTE_TACHES_SOUCHE.copy()

if 'stock_actuel' not in st.session_state:
    st.session_state['stock_actuel'] = DF_STOCK_BASE.copy()

# Dictionnaire regroupant les brouillons de devis par chantier
if 'devis_par_chantier' not in st.session_state:
    st.session_state['devis_par_chantier'] = {}
    for p in st.session_state['liste_chantiers']:
        st.session_state['devis_par_chantier'][p] = pd.DataFrame(DEVIS_MATRICE_BASE)

# ---------------------------------------------------------
# ETINETTES AVERY (33.5 x 38.1 mm)
# ---------------------------------------------------------
def generer_pdf_etiquettes(df_a_imprimer):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_height = 297 * mm
    w_label, h_label = 33.5 * mm, 38.1 * mm
    margin_x, margin_y = 7.2 * mm, 11.1 * mm
    gap_x, gap_y = 2.5 * mm, 1.167 * mm
    cols, rows = 3, 7
    col_idx, row_idx = 0, 0
    
    for _, row in df_a_imprimer.iterrows():
        ref = str(row['Réf'])
        designation = str(row['Désignation'])
        x = margin_x + col_idx * (w_label + gap_x)
        y = page_height - margin_y - (row_idx + 1) * h_label - row_idx * gap_y
        
        c.setStrokeColorRGB(0.85, 0.85, 0.85)
        c.setLineWidth(0.2)
        c.rect(x, y, w_label, h_label)
        
        qr = qrcode.QRCode(box_size=2, border=1)
        qr.add_data(ref)
        qr.make(fit=True)
        img_qr = qr.make_image(fill_color="black", back_color="white").get_image()
        qr_size = 21 * mm
        qr_x = x + (w_label - qr_size) / 2
        qr_y = y + h_label - qr_size - 2.0 * mm
        c.drawImage(ImageReader(img_qr), qr_x, qr_y, width=qr_size, height=qr_size)
        
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 8)
        text_y_ref = qr_y - 3.5 * mm
        c.drawCentredString(x + w_label / 2, text_y_ref, ref)
        
        if designation and designation != ref:
            c.setFont("Helvetica", 6.5)
            c.drawCentredString(x + w_label / 2, text_y_ref - 3.5 * mm, designation[:25])
        
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
# NAVIGATION
# ---------------------------------------------------------
st.sidebar.title("🛠️ Menuiserie App")
menu = st.sidebar.radio(
    "Modules :",
    [
        "🧮 Brouillon Devis & Calcul de Marge",
        "⏱️ Saisie des Heures",
        "📊 Suivi Temps & Rentabilité Chantier",
        "📦 Stock & Mouvements",
        "📷 Scan QR Code Stock",
        "🏷️ Impression Étiquettes"
    ]
)

# ---------------------------------------------------------
# 1. BROUILLON DEVIS PAR CHANTIER
# ---------------------------------------------------------
if menu == "🧮 Brouillon Devis & Calcul de Marge":
    st.header("🧮 Brouillon de Devis et Calcul de Marge")
    
    # Zone de sélection et de création de chantier
    col_sel, col_add = st.columns([2, 1])
    
    with col_sel:
        chantier_devis = st.selectbox("📂 Sélectionner le chantier :", st.session_state['liste_chantiers'], key="sel_chantier_devis")
        
    with col_add:
        st.write("➕ **Créer un nouveau chantier**")
        nouveau_chantier = st.text_input("Nom / Réf du nouveau chantier", placeholder="Ex: 26/230 Agencement Cuisine", label_visibility="collapsed")
        if st.button("➕ Ajouter le chantier"):
            nom_clean = nouveau_chantier.strip()
            if nom_clean and nom_clean not in st.session_state['liste_chantiers']:
                st.session_state['liste_chantiers'].append(nom_clean)
                st.session_state['devis_par_chantier'][nom_clean] = pd.DataFrame(DEVIS_MATRICE_BASE)
                st.success(f"Chantier '{nom_clean}' créé avec son devis type !")
                st.rerun()
            elif nom_clean in st.session_state['liste_chantiers']:
                st.warning("Ce chantier existe déjà.")
    
    st.markdown("---")
    st.subheader(f"📋 Brouillon de Devis : `{chantier_devis}`")
    st.caption("Chaque modification apportée ici est propre à ce chantier. Tu peux aussi utiliser le bouton en bas de tableau pour rajouter des lignes.")
    
    # Chargement du devis spécifique
    if chantier_devis not in st.session_state['devis_par_chantier']:
        st.session_state['devis_par_chantier'][chantier_devis] = pd.DataFrame(DEVIS_MATRICE_BASE)
        
    df_devis_actuel = st.session_state['devis_par_chantier'][chantier_devis].copy()
    
    edited_df = st.data_editor(
        df_devis_actuel,
        column_config={
            "Poste": st.column_config.TextColumn("Désignation / Poste", disabled=False),
            "U": st.column_config.SelectboxColumn("Unité (UV)", options=["heures", "M²", "Ml", "pc", "Forf pour 1", "Forf", "Km", "% perte", "Jour", "M3"], required=True),
            "Q": st.column_config.NumberColumn("Quantité (Q)", min_value=0.0, step=1.0, format="%.2f"),
            "Prix Achat HT": st.column_config.NumberColumn("Prix Achat HT (€)", min_value=0.0, step=0.1, format="%.2f €"),
            "Prix Vente HT": st.column_config.NumberColumn("Prix Vente HT (€)", min_value=0.0, step=0.1, format="%.2f €"),
        },
        num_rows="dynamic",
        use_container_width=True,
        key=f"editor_devis_{chantier_devis}"
    )
    
    # Sauvegarde automatique du brouillon du chantier
    st.session_state['devis_par_chantier'][chantier_devis] = edited_df
    
    # Calculs automatiques des sous-totaux et marges
    edited_df["Total Achat HT"] = edited_df["Q"] * edited_df["Prix Achat HT"]
    edited_df["Total Vente HT"] = edited_df["Q"] * edited_df["Prix Vente HT"]
    edited_df["Marge Brute HT"] = edited_df["Total Vente HT"] - edited_df["Total Achat HT"]
    
    total_achat = edited_df["Total Achat HT"].sum()
    total_vente = edited_df["Total Vente HT"].sum()
    marge_totale = edited_df["Marge Brute HT"].sum()
    taux_marque = (marge_totale / total_vente * 100) if total_vente > 0 else 0.0
    
    # Filtrage des heures devisées
    heures_df = edited_df[edited_df["U"] == "heures"]
    total_heures_devis = heures_df["Q"].sum()
    
    st.markdown("---")
    st.subheader(f"📊 Synthèse & Marges pour `{chantier_devis}`")
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Achat HT", f"{total_achat:,.2f} €")
    c2.metric("Total Devis HT", f"{total_vente:,.2f} €")
    c3.metric("Marge Brute Globale", f"{marge_totale:,.2f} €", f"{taux_marque:.1f}% de marque")
    c4.metric("Heures Devisées", f"{total_heures_devis:.2f} h")

# ---------------------------------------------------------
# 2. SAISIE DES HEURES
# ---------------------------------------------------------
elif menu == "⏱️ Saisie des Heures":
    st.header("⏱️ Saisie Rapide des Heures")
    with st.form("form_saisie_heures", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            date_saisie = st.date_input("Date", datetime.now())
            chantier = st.selectbox("Chantier / Projet", st.session_state['liste_chantiers'])
        with col2:
            code_tache = st.selectbox("Tâche / Activité", st.session_state['liste_taches'])
            heures = st.number_input("Heures", min_value=0.25, max_value=12.0, step=0.25, value=1.0)
            
        if st.form_submit_button("💾 Enregistrer"):
            est_prod = not (code_tache.startswith("X") or "BUREAU" in code_tache.upper())
            st.session_state['historique_heures'].append({
                "Date": str(date_saisie),
                "Chantier": chantier,
                "Code": code_tache,
                "Heures": heures,
                "Production": est_prod
            })
            st.success(f"Enregistré : {heures}h sur {chantier}")

# ---------------------------------------------------------
# 3. SUIVI TEMPS & RENTABILITE CHANTIER
# ---------------------------------------------------------
elif menu == "📊 Suivi Temps & Rentabilité Chantier":
    st.header("📊 Suivi Temps & Rentabilité par Chantier")
    
    df_global = DF_KIMAI_HISTO.copy()
    if st.session_state['historique_heures']:
        df_sess = pd.DataFrame(st.session_state['historique_heures'])
        df_sess.rename(columns={"Chantier": "Projet", "Code": "Tâche"}, inplace=True)
        df_global = pd.concat([df_global, df_sess[["Projet", "Tâche", "Heures", "Production"]]], ignore_index=True)

    projet_sel = st.selectbox("🎯 Choisir le chantier :", st.session_state['liste_chantiers'])
    df_proj = df_global[df_global["Projet"] == projet_sel].copy()
    
    if not df_proj.empty:
        # Filtre optionnel pour masquer les éléments parasites sans altérer les données brutes
        taches_dispo = df_proj["Tâche"].unique().tolist()
        taches_selectionnees = st.multiselect(
            "🔎 Tâches à inclure dans l'affichage :", 
            options=taches_dispo, 
            default=taches_dispo
        )
        
        df_filtr = df_proj[df_proj["Tâche"].isin(taches_selectionnees)]
        
        total_h = df_filtr["Heures"].sum()
        h_prod = df_filtr[df_filtr["Production"] == True]["Heures"].sum()
        h_bureau = df_filtr[df_filtr["Production"] == False]["Heures"].sum()
        
        c_h1, c_h2, c_h3 = st.columns(3)
        c_h1.metric("⏱️ Total Heures Passées", f"{total_h:,.2f} h")
        c_h2.metric("🔨 Production", f"{h_prod:,.2f} h")
        c_h3.metric("✏️ Etudes / Bureau", f"{h_bureau:,.2f} h")
        
        if total_h > 0:
            st.subheader("📉 Répartition des heures")
            st.bar_chart(df_filtr.groupby("Tâche")["Heures"].sum())
            
        with st.expander("📄 Voir le détail de toutes les lignes d'heures brutes"):
            st.dataframe(df_proj, use_container_width=True)
    else:
        st.info("Aucune heure enregistrée pour ce chantier.")

# ---------------------------------------------------------
# 4. STOCK & MOUVEMENTS
# ---------------------------------------------------------
elif menu == "📦 Stock & Mouvements":
    st.header("📦 Consultation & Gestion Stock")
    st.dataframe(st.session_state['stock_actuel'], use_container_width=True)

# ---------------------------------------------------------
# 5. SCAN QR CODE STOCK
# ---------------------------------------------------------
elif menu == "📷 Scan QR Code Stock":
    st.header("📷 Numérisation QR Code")
    img_captured = st.camera_input("Scanner le QR Code")
    if img_captured:
        results = zxingcpp.read_barcodes(Image.open(img_captured))
        if results:
            qr_data = results[0].text.strip()
            st.success(f"QR Code : `{qr_data}`")

# ---------------------------------------------------------
# 6. IMPRESSION ETIQUETTES
# ---------------------------------------------------------
elif menu == "🏷️ Impression Étiquettes":
    st.header("🏷️ Impression d'Étiquettes")
    df_imp = st.session_state['stock_actuel']
    articles = st.multiselect("Sélectionner les articles :", df_imp["Réf"].tolist(), default=df_imp["Réf"].tolist()[:3])
    if articles:
        pdf_data = generer_pdf_etiquettes(df_imp[df_imp["Réf"].isin(articles)])
        st.download_button("📄 Télécharger le PDF (Avery)", pdf_data, "etiquettes.pdf", "application/pdf")
