import streamlit as st
import pandas as pd
from datetime import datetime
import os
import io
import glob

# Bibliothèques pour la génération de PDF et de QR Codes
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import qrcode

# Bibliothèques pour la lecture de QR Codes
from PIL import Image
import zxingcpp

st.set_page_config(
    page_title="Gestion Menuiserie", 
    page_icon="🪵", 
    layout="wide"
)

# ---------------------------------------------------------
# CHARGEMENT DES ACTIVITÉS DEPUIS TON FICHIER "Activités.xlsx"
# ---------------------------------------------------------
@st.cache_data
def charger_activites_souche():
    fichier_souche = "Activités.xlsx"
    taches_uniques = []
    
    if os.path.exists(fichier_souche):
        try:
            df_act = pd.read_excel(fichier_souche, sheet_name="Réf")
            
            # Reconstruction propre des activités (Code - Libellé)
            for idx, row in df_act.iterrows():
                code = str(row['D1']).strip() if pd.notna(row['D1']) else ""
                nom = str(row['Appro Débit']).strip() if pd.notna(row['Appro Débit']) else ""
                
                if code and nom and code != "nan" and nom != "nan":
                    taches_uniques.append(f"{code} - {nom}")
        except Exception as e:
            st.error(f"Erreur lors de la lecture de Activités.xlsx : {e}")

    # En cas d'absence du fichier, liste de secours complète des 62 activités
    if not taches_uniques:
        taches_uniques = [
            "D2 - Débit massif", "D3 - Scie à ruban", "D4 - Scie à format", "D5 - Scie à Panneaux", "D6 - Scie à rubau metaux", "D7 - Scie radial",
            "C1 - Corroyeuse 4 faces", "C2 - Dégauchissage", "C3 - Rabotage",
            "U1 - Toupie", "U2 - Défonceuses", "U3 - Scie à ruban", "U4 - Plaqueuse", "U5 - CN", "U6 - Tournage", "U7 - Mortaisage", "U8 - Tennonage", "U9 - Pointage machine", "U10 - Usinage", "U11 - Perçage manuelle",
            "B1 - Finition manuelle", "B2 - Ponceuse à bande", "B3 - Placage chant main", "B4 - Affleurage chants", "B5 - Vitrage", "B6 - Cassage d'arrêtes", "B7 - Brossage",
            "M1 - Montage (cadrage)", "M2 - Montage caisses", "M3 - Collage Px", "M4 - Collage neoprene",
            "X1 - Etude de plan", "X2 - Etablissement", "X3 - Traçage", "X4 - Affutage", "X5 - Bureau", "X6 - RDV Clientèle", "X7 - Relevé de côtes", "X8 - Mise en page devis", "X9 - Création publication réseau",
            "V1 - Vernissage pistolet", "V2 - Laquage pistolet", "V3 - Pinceau", "V4 - Teinture", "V5 - Egrenage", "V6 - Traitement IFH",
            "Q1 - Pose quincaillerie",
            "N1 - Nettoyage atelier", "N2 - Rangement atelier", "N3 - Affutage", "N4 - Changement sac aspi", "N5 - Changement plaquettes",
            "T1 - Trajet", "T2 - Chargement", "T3 - Mise au sechoir", "T4 - Emballage", "T5 - Manutention",
            "P1 - Pose sur chantier", "P2 - Demontage", "P3 - Evacuation", "P4 - Fab gabarits", "P5 - Rangement nettoyage chantier"
        ]

    return taches_uniques

# ---------------------------------------------------------
# CHARGEMENT DES EXPORTS HEURES ET DU STOCK
# ---------------------------------------------------------
@st.cache_data
def charger_donnees_kimai_heures():
    fichiers_kimai = glob.glob("kimai-export*.xlsx") + glob.glob("*export*.xlsx")
    fichiers_kimai = list(set(fichiers_kimai))
    
    donnees_cumulees = []
    
    for fichier in fichiers_kimai:
        if os.path.exists(fichier) and "activit" not in fichier.lower():
            try:
                df = pd.read_excel(fichier)
                col_nom = df.columns[0]
                col_total = df.columns[1]
                
                projet_actuel = "Général"
                
                for idx, row in df.iterrows():
                    val = str(row[col_nom]).strip()
                    total_str = str(row[col_total]).replace(',', '.')
                    
                    try:
                        total_heures = float(total_str)
                    except ValueError:
                        total_heures = 0.0
                        
                    if val == "nan" or not val:
                        continue
                        
                    if any(val.startswith(p) for p in ["OE ", "25/", "26/", "Agencement"]):
                        projet_actuel = val
                    else:
                        tache_cleanee = val.replace('\t', ' - ').strip()
                        donnees_cumulees.append({
                            "Projet": projet_actuel,
                            "Tâche": tache_cleanee,
                            "Heures": total_heures,
                            "Fichier": fichier
                        })
            except Exception:
                pass

    df_kimai = pd.DataFrame(donnees_cumulees)
    
    def est_production(tache):
        t = str(tache).upper()
        if t.startswith("X") or "BUREAU" in t or "DEVIS" in t or "RDV" in t:
            return False
        return True

    if not df_kimai.empty:
        df_kimai["Production"] = df_kimai["Tâche"].apply(est_production)
        projets_uniques = sorted(df_kimai["Projet"].unique().tolist())
    else:
        projets_uniques = ["26/221 Fabrication et pose d'étagères", "26/227 Réfection plan de travail"]

    return df_kimai, projets_uniques

@st.cache_data
def charger_stock():
    fichier_inv = None
    fichiers_inv_trouves = glob.glob("*stock*.xlsx") + glob.glob("*Inventaire*.xlsx")
    if fichiers_inv_trouves:
        fichier_inv = fichiers_inv_trouves[0]

    if fichier_inv and os.path.exists(fichier_inv):
        try:
            df_inv = pd.read_excel(fichier_inv, sheet_name="Inventaire pour bilan 2025")
            df_inv = df_inv.dropna(subset=[df_inv.columns[0]])
            df_inv.columns = ["Désignation", "Quantité", "Prix Unitaire HT", "Unité", "Total HT", "Col6", "Col7"][:len(df_inv.columns)]
            df_inv = df_inv[df_inv["Désignation"] != "Désignation"]
            df_inv["Réf"] = df_inv["Désignation"].str.replace(r'[^a-zA-Z0-9\s-]', '', regex=True).str.strip().str.replace(' ', '-')
            
            def categoriser(row):
                des = str(row["Désignation"]).lower()
                unite = str(row["Unité"]).lower()
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

LISTE_TACHES_SOUCHE = charger_activites_souche()
DF_KIMAI_HISTO, LISTE_CHANTIERS = charger_donnees_kimai_heures()
DF_STOCK_BASE = charger_stock()

# Initialisation des états en session
if 'historique_heures' not in st.session_state:
    st.session_state['historique_heures'] = []

if 'liste_taches' not in st.session_state:
    st.session_state['liste_taches'] = LISTE_TACHES_SOUCHE.copy()

if 'stock_actuel' not in st.session_state:
    st.session_state['stock_actuel'] = DF_STOCK_BASE.copy()

if 'mouvements_stock' not in st.session_state:
    st.session_state['mouvements_stock'] = []

# ---------------------------------------------------------
# GENERATION PDF ETIQUETTES AVERY (63.5 mm x 38.1 mm)
# ---------------------------------------------------------
def generer_pdf_etiquettes(df_a_imprimer):
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    page_height = 297 * mm
    
    w_label = 63.5 * mm
    h_label = 38.1 * mm
    margin_x = 7.2 * mm
    margin_y = 11.1 * mm
    gap_x = 2.5 * mm
    gap_y = 1.167 * mm
    
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
        qr_image_reader = ImageReader(img_qr)
        
        qr_size = 21 * mm
        qr_x = x + (w_label - qr_size) / 2
        qr_y = y + h_label - qr_size - 2.0 * mm
        c.drawImage(qr_image_reader, qr_x, qr_y, width=qr_size, height=qr_size)
        
        c.setFillColorRGB(0, 0, 0)
        c.setFont("Helvetica-Bold", 8)
        text_y_ref = qr_y - 3.5 * mm
        c.drawCentredString(x + w_label / 2, text_y_ref, ref)
        
        if designation and designation != ref:
            c.setFont("Helvetica", 6.5)
            text_y_des = text_y_ref - 3.5 * mm
            c.drawCentredString(x + w_label / 2, text_y_des, designation[:35])
        
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
# MENU DE NAVIGATION
# ---------------------------------------------------------
st.sidebar.title("🛠️ Gestion Menuiserie")
menu = st.sidebar.radio(
    "Accéder aux modules :",
    [
        "⏱️ Saisie des Heures",
        "📊 Suivi Temps & Historique Kimai",
        "🧮 Brouillon Devis & Marges",
        "📦 Stock & Mouvements",
        "📷 Scan QR Code Stock",
        "🏷️ Impression Étiquettes Stock"
    ]
)

# ---------------------------------------------------------
# 1. SAISIE DES HEURES + AJOUT MANUEL DE TÂCHE
# ---------------------------------------------------------
if menu == "⏱️ Saisie des Heures":
    st.header("⏱️ Saisie Rapide des Heures Atelier & Chantier")
    
    with st.form("form_saisie_heures", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            date_saisie = st.date_input("Date d'intervention", datetime.now())
            chantier = st.selectbox("Chantier / Projet", LISTE_CHANTIERS)
        with col2:
            code_tache = st.selectbox("Tâche / Activité (issue de Activités.xlsx)", st.session_state['liste_taches'])
            heures = st.number_input("Nombre d'heures effectuées", min_value=0.25, max_value=12.0, step=0.25, value=1.0)
            
        valider = st.form_submit_button("💾 Enregistrer l'intervention")
        
        if valider:
            est_prod = not (code_tache.startswith("X") or "BUREAU" in code_tache.upper())
            st.session_state['historique_heures'].append({
                "Date": str(date_saisie),
                "Chantier": chantier,
                "Code": code_tache,
                "Heures": heures,
                "Production": est_prod
            })
            st.success(f"Enregistré : {heures}h sur **{chantier}** ({code_tache})")

    with st.expander("➕ Ajouter une nouvelle tâche / activité personnalisée"):
        with st.form("form_nouvelle_tache", clear_on_submit=True):
            col_nt1, col_nt2 = st.columns([3, 1])
            nouvelle_tache_nom = col_nt1.text_input("Nom de la nouvelle tâche", placeholder="Z1 - Maquette d'essai")
            btn_add_tache = col_nt2.form_submit_button("➕ Ajouter la tâche")
            
            if btn_add_tache and nouvelle_tache_nom.strip():
                nom_clean = nouvelle_tache_nom.strip()
                if nom_clean not in st.session_state['liste_taches']:
                    st.session_state['liste_taches'].append(nom_clean)
                    st.session_state['liste_taches'] = sorted(st.session_state['liste_taches'])
                    st.success(f"Tâche **'{nom_clean}'** ajoutée à la liste !")
                    st.rerun()
                else:
                    st.warning("Cette tâche existe déjà dans la liste.")

    st.subheader("📋 Saisies de la session en cours")
    if st.session_state['historique_heures']:
        st.dataframe(pd.DataFrame(st.session_state['historique_heures']), use_container_width=True)
    else:
        st.info("Aucune saisie effectuée au cours de la session active.")

# ---------------------------------------------------------
# 2. SUIVI TEMPS & HISTORIQUE KIMAI
# ---------------------------------------------------------
elif menu == "📊 Suivi Temps & Historique Kimai":
    st.header("📊 Historique Kimai Cumulé")
    
    if not DF_KIMAI_HISTO.empty:
        total_prod = DF_KIMAI_HISTO[DF_KIMAI_HISTO["Production"] == True]["Heures"].sum()
        total_hors_prod = DF_KIMAI_HISTO[DF_KIMAI_HISTO["Production"] == False]["Heures"].sum()
        total_global = total_prod + total_hors_prod
        
        col_k1, col_k2, col_k3 = st.columns(3)
        col_k1.metric("Total Heures Production", f"{total_prod:,.2f} h")
        col_k2.metric("Total Heures Hors-Prod / Bureau", f"{total_hors_prod:,.2f} h")
        col_k3.metric("Volume Total Enregistré Kimai", f"{total_global:,.2f} h")
        
        st.markdown("---")
        
        st.subheader("🔍 Recherche & Filtrage par Projet")
        projet_selectionne = st.selectbox("Sélectionner un projet Kimai :", ["Tous les projets"] + LISTE_CHANTIERS)
        
        if projet_selectionne != "Tous les projets":
            df_filtre = DF_KIMAI_HISTO[DF_KIMAI_HISTO["Projet"] == projet_selectionne]
        else:
            df_filtre = DF_KIMAI_HISTO
            
        st.dataframe(df_filtre[["Projet", "Tâche", "Heures", "Production"]], use_container_width=True)
        
        st.subheader("📈 Répartition par Tâche / Activité")
        df_recap_taches = df_filtre.groupby("Tâche")["Heures"].sum().reset_index().sort_values(by="Heures", ascending=False)
        st.bar_chart(df_recap_taches.set_index("Tâche"))
    else:
        st.warning("Aucun fichier d'export Kimai trouvé à la racine du projet.")

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
# 4. STOCK & MOUVEMENTS
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
# 5. SCAN QR CODE STOCK
# ---------------------------------------------------------
elif menu == "📷 Scan QR Code Stock":
    st.header("📷 Numérisation d'Étiquettes Stock")
    img_captured = st.camera_input("Prendre en photo l'étiquette QR Code")
    
    if img_captured:
        img = Image.open(img_captured)
        results = zxingcpp.read_barcodes(img)
        
        if results:
            qr_data = results[0].text.strip()
            st.success(f"✅ **QR Code détecté :** `{qr_data}`")
            
            df_stock = st.session_state['stock_actuel']
            article = df_stock[df_stock["Réf"] == qr_data]
            
            if not article.empty:
                art_info = article.iloc[0]
                st.markdown(f"### Article trouvé : **{art_info['Désignation']}**")
                st.write(f"* **Catégorie :** {art_info['Catégorie']}")
                st.write(f"* **Stock actuel :** {art_info['Quantité']} {art_info['Unité']}")
                st.write(f"* **Prix unitaire HT :** {art_info['Prix Unitaire HT']} €")
                
                st.markdown("---")
                c_act1, c_act2 = st.columns(2)
                qte_retrait = c_act1.number_input("Quantité à retirer / ajouter", min_value=1.0, value=1.0, step=1.0)
                
                if c_act1.button("➖ Sortie de stock"):
                    idx = df_stock[df_stock["Réf"] == qr_data].index[0]
                    st.session_state['stock_actuel'].at[idx, "Quantité"] -= qte_retrait
                    st.success(f"Retrait de {qte_retrait} effectué pour {qr_data}.")
                    st.rerun()
                    
                if c_act2.button("➕ Entrée en stock"):
                    idx = df_stock[df_stock["Réf"] == qr_data].index[0]
                    st.session_state['stock_actuel'].at[idx, "Quantité"] += qte_retrait
                    st.success(f"Ajout de {qte_retrait} effectué pour {qr_data}.")
                    st.rerun()
            else:
                st.warning(f"La référence `{qr_data}` a été lue mais elle n'existe pas dans le stock actuel.")
        else:
            st.error("❌ Aucun QR Code n'a pu être lu sur cette photo.")

# ---------------------------------------------------------
# 6. IMPRESSION ÉTIQUETTES STOCK
# ---------------------------------------------------------
elif menu == "🏷️ Impression Étiquettes Stock":
    st.header("🏷️ Impression d'Étiquettes QR Code")
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
