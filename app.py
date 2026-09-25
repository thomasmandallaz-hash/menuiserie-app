import json
import sqlite3
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "atelier.db"
DEFAULT_HOURLY_RATE = 57.0

QUOTE_ROWS = [
    {"Poste": "Etude", "Unité": "heures", "Prix achat HT": 70.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "Panneaux de caissons", "Unité": "M²", "Prix achat HT": 7.8, "Coefficient vente": 1.6, "Référence fournisseur": ""},
    {"Poste": "Panneaux de façade", "Unité": "M²", "Prix achat HT": 15.0, "Coefficient vente": 1.6, "Référence fournisseur": ""},
    {"Poste": "Panneau de fond", "Unité": "M²", "Prix achat HT": 8.8, "Coefficient vente": 1.6, "Référence fournisseur": ""},
    {"Poste": "Rouleau de chant", "Unité": "Ml", "Prix achat HT": 0.65, "Coefficient vente": 1.6, "Référence fournisseur": ""},
    {"Poste": "Charnières + embases", "Unité": "pc", "Prix achat HT": 5.0, "Coefficient vente": 1.6, "Référence fournisseur": ""},
    {"Poste": "Tiroirs", "Unité": "Forf pour 1", "Prix achat HT": 170.0, "Coefficient vente": 1.1, "Référence fournisseur": ""},
    {"Poste": "Barre à penderie", "Unité": "", "Prix achat HT": 0.0, "Coefficient vente": 1.6, "Référence fournisseur": ""},
    {"Poste": "Fourniture materiel", "Unité": "", "Prix achat HT": 0.0, "Coefficient vente": 1.6, "Référence fournisseur": ""},
    {"Poste": "M.O Fab", "Unité": "heures", "Prix achat HT": 75.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "M.O pose", "Unité": "heures", "Prix achat HT": 75.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "M.O chargement", "Unité": "heures", "Prix achat HT": 75.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "Deplacement pour pose", "Unité": "heures", "Prix achat HT": 75.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "Forfait kilometrique", "Unité": "Km", "Prix achat HT": 0.606, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "Relevé de cote", "Unité": "heures", "Prix achat HT": 0.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "Intendance (appel téléphonique etc)", "Unité": "forf", "Prix achat HT": 0.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "Déchetterie", "Unité": "% de perte", "Prix achat HT": 0.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "Déplacement pour devis", "Unité": "heures", "Prix achat HT": 0.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
    {"Poste": "Devis", "Unité": "forf", "Prix achat HT": 70.0, "Coefficient vente": 1.0, "Référence fournisseur": ""},
]


def db():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    con.executescript("""
    CREATE TABLE IF NOT EXISTS heures (id INTEGER PRIMARY KEY, jour TEXT, collaborateur TEXT, client TEXT, projet TEXT, heures REAL, production REAL, commentaire TEXT, statut TEXT DEFAULT 'Saisi');
    CREATE TABLE IF NOT EXISTS mouvements (id INTEGER PRIMARY KEY, horodatage TEXT, article TEXT, quantite REAL, type TEXT, emplacement TEXT, reference TEXT, commentaire TEXT);
    CREATE TABLE IF NOT EXISTS taches (id INTEGER PRIMARY KEY, semaine TEXT, titre TEXT, responsable TEXT, echeance TEXT, statut TEXT, archive INTEGER DEFAULT 0, note TEXT);
    CREATE TABLE IF NOT EXISTS etiquettes (id INTEGER PRIMARY KEY, article TEXT, reference TEXT, lot TEXT, quantite INTEGER, cree_le TEXT);
    """)
    return con


def money(v):
    return f"{float(v or 0):,.2f} €".replace(",", " ").replace(".", ",")


def load_quote(uploaded=None):
    rows = pd.DataFrame(QUOTE_ROWS)
    if uploaded is None:
        return rows
    try:
        book = pd.ExcelFile(uploaded)
        sheet = "Info complémentaires" if "Info complémentaires" in book.sheet_names else book.sheet_names[0]
        raw = pd.read_excel(uploaded, sheet_name=sheet, header=None)
        parsed = []
        for _, r in raw.iloc[3:27].iterrows():
            if pd.notna(r.iloc[0]):
                parsed.append({
                    "Poste": str(r.iloc[0]), "Unité": str(r.iloc[1]) if pd.notna(r.iloc[1]) else "",
                    "Prix achat HT": float(r.iloc[3]) if pd.notna(r.iloc[3]) and isinstance(r.iloc[3], (int, float)) else 0.0,
                    "Coefficient vente": 1.0, "Référence fournisseur": "",
                })
        if parsed:
            rows = pd.DataFrame(parsed)
    except Exception as exc:
        st.warning(f"Import Excel impossible : {exc}")
    return rows


def page_saisie_heures():
    st.header("⏱️ Saisie des heures")
    with st.form("hours_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        jour = c1.date_input("Jour", date.today())
        collaborateur = c2.text_input("Collaborateur")
        client = c3.text_input("Client / chantier")
        c4, c5, c6 = st.columns(3)
        projet = c4.text_input("Projet")
        heures = c5.number_input("Heures", min_value=0.0, step=0.25)
        production = c6.number_input("Production / unités", min_value=0.0, step=1.0)
        commentaire = st.text_area("Commentaire")
        if st.form_submit_button("Enregistrer", type="primary"):
            if not collaborateur.strip() or not client.strip():
                st.error("Le collaborateur et le client sont obligatoires.")
            else:
                con = db(); con.execute("INSERT INTO heures (jour,collaborateur,client,projet,heures,production,commentaire) VALUES (?,?,?,?,?,?,?)", (jour.isoformat(), collaborateur, client, projet, heures, production, commentaire)); con.commit(); con.close(); st.success("Saisie enregistrée.")
    con = db(); df = pd.read_sql_query("SELECT * FROM heures ORDER BY jour DESC, id DESC LIMIT 100", con); con.close()
    if not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)


def page_suivi():
    st.header("📊 Suivi Temps & Production")
    con = db(); df = pd.read_sql_query("SELECT * FROM heures", con); con.close()
    if df.empty:
        st.info("Aucune saisie d'heures pour le moment."); return
    df["jour"] = pd.to_datetime(df["jour"])
    c1, c2 = st.columns(2)
    start = c1.date_input("Du", df["jour"].min().date()); end = c2.date_input("Au", date.today())
    f = df[(df["jour"].dt.date >= start) & (df["jour"].dt.date <= end)]
    m1, m2, m3 = st.columns(3); m1.metric("Heures", f["heures"].sum()); m2.metric("Production", f["production"].sum()); m3.metric("Saisies", len(f))
    st.bar_chart(f.groupby("jour")["heures"].sum())
    st.dataframe(f, use_container_width=True, hide_index=True)


def page_devis():
    st.header("🧾 Brouillon devis & matrice")
    uploaded = st.file_uploader("Importer / remplacer la matrice 00 DEVIS.xlsx", type=["xlsx"])
    if "quote_df" not in st.session_state or uploaded is not None:
        st.session_state.quote_df = load_quote(uploaded)
    rate = st.number_input("Taux horaire de référence (€ HT)", min_value=0.0, value=DEFAULT_HOURLY_RATE, step=1.0)
    df = st.session_state.quote_df.copy()
    df["Quantité"] = 0.0
    edited = st.data_editor(df, num_rows="dynamic", use_container_width=True, hide_index=True, column_config={"Prix achat HT": st.column_config.NumberColumn(format="%.2f €"), "Coefficient vente": st.column_config.NumberColumn(format="%.2f"), "Quantité": st.column_config.NumberColumn(min_value=0.0, step=0.25)})
    edited["Prix vente unitaire HT"] = edited["Prix achat HT"].fillna(0) * edited["Coefficient vente"].fillna(1)
    edited["Total achat HT"] = edited["Quantité"].fillna(0) * edited["Prix achat HT"].fillna(0)
    edited["Total vente HT"] = edited["Quantité"].fillna(0) * edited["Prix vente unitaire HT"].fillna(0)
    hours = edited.loc[edited["Unité"].astype(str).str.lower().str.contains("heure"), "Quantité"].sum()
    ca = edited["Total vente HT"].sum(); achat = edited["Total achat HT"].sum(); marge = ca - achat
    marque = marge / ca if ca else 0
    a, b, c, d = st.columns(4); a.metric("Heures productives", f"{hours:.2f}"); b.metric("CA total HT", money(ca)); c.metric("Marge brute", money(marge)); d.metric("Taux de marque", f"{marque:.1%}")
    st.caption(f"Taux horaire devis calculé : {ca / hours if hours else 0:.2f} € / h — référence {rate:.2f} € / h")
    st.dataframe(edited[["Poste", "Unité", "Quantité", "Prix achat HT", "Prix vente unitaire HT", "Total achat HT", "Total vente HT", "Référence fournisseur"]], use_container_width=True, hide_index=True)
    out = BytesIO(); edited.to_excel(out, index=False, sheet_name="Devis"); st.download_button("Télécharger le brouillon Excel", out.getvalue(), "brouillon_devis.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_stock():
    st.header("📦 Stock & mouvements")
    with st.form("stock_form", clear_on_submit=True):
        c = st.columns(4); article = c[0].text_input("Article"); typ = c[1].selectbox("Type", ["Entrée", "Sortie", "Inventaire"]); qty = c[2].number_input("Quantité", min_value=0.0, step=1.0); ref = c[3].text_input("Référence")
        emp = st.text_input("Emplacement"); note = st.text_input("Commentaire")
        if st.form_submit_button("Enregistrer mouvement"):
            con = db(); con.execute("INSERT INTO mouvements (horodatage,article,quantite,type,emplacement,reference,commentaire) VALUES (?,?,?,?,?,?,?)", (datetime.now().isoformat(timespec="seconds"), article, qty, typ, emp, ref, note)); con.commit(); con.close(); st.success("Mouvement enregistré.")
    con = db(); moves = pd.read_sql_query("SELECT * FROM mouvements ORDER BY id DESC", con); con.close()
    if not moves.empty:
        moves["variation"] = moves.apply(lambda r: r["quantite"] if r["type"] == "Entrée" else -r["quantite"] if r["type"] == "Sortie" else r["quantite"], axis=1)
        st.subheader("Stock théorique par article"); st.dataframe(moves.groupby("article", as_index=False)["variation"].sum().rename(columns={"variation": "Stock"}), use_container_width=True, hide_index=True); st.dataframe(moves, use_container_width=True, hide_index=True)


def page_qr():
    st.header("🔎 Scan QR Code")
    st.info("Utilisez la caméra de votre appareil ou saisissez le contenu du QR code.")
    value = st.text_input("Contenu / référence scannée")
    image = st.camera_input("Scanner avec la caméra")
    if image:
        st.warning("Décodage automatique optionnel : installez opencv-python et pyzbar pour activer le décodage caméra.")
    if value: st.success(f"Code détecté : {value}"); st.session_state["last_qr"] = value


def page_labels():
    st.header("🏷️ Impression étiquettes")
    c = st.columns(4); article = c[0].text_input("Article", value=st.session_state.get("last_qr", "")); ref = c[1].text_input("Référence"); lot = c[2].text_input("Lot"); qty = c[3].number_input("Nombre", 1, 500, 1)
    cols = st.columns(3); width = cols[0].number_input("Largeur étiquette (mm)", 20.0, 210.0, 50.0); height = cols[1].number_input("Hauteur (mm)", 10.0, 150.0, 30.0); gap = cols[2].number_input("Espacement entre étiquettes (mm)", 0.0, 30.0, 3.0)
    if st.button("Générer la planche d'étiquettes", type="primary"):
        lines = ["<!doctype html><meta charset='utf-8'><style>", f"@page{{size:A4;margin:10mm}} body{{font-family:Arial;display:grid;grid-template-columns:repeat(auto-fill,{width}mm);gap:{gap}mm}} .label{{width:{width}mm;height:{height}mm;border:1px solid #111;box-sizing:border-box;padding:3mm;page-break-inside:avoid}}", "</style>"]
        for i in range(int(qty)): lines.append(f"<div class='label'><b>{article or 'Article'}</b><br>Réf. : {ref}<br>Lot : {lot}<br><small>Étiquette {i+1}/{int(qty)}</small></div>")
        html = "".join(lines); st.download_button("Télécharger la planche HTML à imprimer", html, "etiquettes.html", "text/html")


def page_admin():
    st.header("📅 Lundi administratif")
    monday = date.today() - timedelta(days=date.today().weekday())
    with st.form("task_form", clear_on_submit=True):
        c = st.columns(4); title = c[0].text_input("Tâche"); who = c[1].text_input("Responsable"); due = c[2].date_input("Échéance", monday); status = c[3].selectbox("Statut", ["À faire", "En cours", "Fait", "Bloqué"]); note = st.text_input("Note")
        if st.form_submit_button("Ajouter"):
            con = db(); con.execute("INSERT INTO taches (semaine,titre,responsable,echeance,statut,note) VALUES (?,?,?,?,?,?)", (monday.isoformat(), title, who, due.isoformat(), status, note)); con.commit(); con.close(); st.rerun()
    con = db(); tasks = pd.read_sql_query("SELECT * FROM taches WHERE archive=0 ORDER BY echeance,id", con); con.close()
    if tasks.empty: st.info("Aucune tâche active."); return
    for _, r in tasks.iterrows():
        c1, c2, c3, c4 = st.columns([0.08, 0.42, 0.2, 0.2])
        done = c1.checkbox("", key=f"done_{r['id']}", value=r["statut"] == "Fait")
        c2.write(f"**{r['titre']}**\n\n{r['note'] or ''}"); c3.write(f"{r['responsable']}\n\nÉchéance : {r['echeance']}")
        new_status = c4.selectbox("Statut", ["À faire", "En cours", "Fait", "Bloqué"], index=["À faire", "En cours", "Fait", "Bloqué"].index(r["statut"]), key=f"status_{r['id']}")
        if done != (r["statut"] == "Fait") or new_status != r["statut"]:
            con = db(); con.execute("UPDATE taches SET statut=? WHERE id=?", ("Fait" if done else new_status, int(r["id"]))); con.commit(); con.close(); st.rerun()
    st.divider(); st.subheader("Rapport automatique")
    report = tasks.groupby("statut").size().rename("Nombre").reset_index(); st.dataframe(report, hide_index=True, use_container_width=True)
    text = "Rapport administratif — " + date.today().isoformat() + "\n" + report.to_string(index=False)
    st.download_button("Télécharger le rapport", text, "rapport_lundi.txt", "text/plain")
    if st.button("Archiver les tâches terminées"):
        con = db(); con.execute("UPDATE taches SET archive=1 WHERE statut='Fait'"); con.commit(); con.close(); st.rerun()


def main():
    st.set_page_config(page_title="Atelier — Gestion", page_icon="🛠️", layout="wide")
    st.title("🛠️ Atelier — Gestion de production")
    pages = {"Saisie des Heures": page_saisie_heures, "Suivi Temps & Production": page_suivi, "Brouillon Devis & Matrice": page_devis, "Stock & Mouvements": page_stock, "Scan QR Code": page_qr, "Impression Étiquettes": page_labels, "Lundi Administratif": page_admin}
    choice = st.sidebar.radio("Modules", list(pages))
    st.sidebar.caption("Données locales SQLite · taux de référence devis : 57 € HT/h")
    pages[choice]()


if __name__ == "__main__":
    main()
