import streamlit as st
import pandas as pd
import io

# Configuration de la page
st.set_page_config(page_title="NT 02/DSB/2007 - Analyse Réglementaire", layout="wide")

# Titre
st.title("🏦 Notice Technique NT 02/DSB/2007")
st.markdown("### Structuration des exigences en fonds propres (Circulaire 26/G/2006)")

# --- 1. CHARGEMENT DES DONNÉES ---
# Dans un cas réel, vous chargeriez un fichier CSV : df = pd.read_csv("data_reglementation.csv")
# Ici, je recrée un extrait basé sur notre tableau précédent pour la démo

data = {
    "Article": [
        "Art. 1", "Art. 2", "Art. 3 à 9", "Art. 10", "Art. 11", "Art. 12", 
        "Art. 55", "Art. 56", "Art. 70", "Art. 76", "Art. 74", "Art. 81"
    ],
    "ACTIF - Risque de Crédit": [
        "OEEC éligibles", "", "Segmentation clientèle", "Types de créances", "Engagements Hors-Bilan", "Équivalent-risque dérivés",
        "", "", "", "Lignes de métier", "", ""
    ],
    "ACTIF - Risque de Marché": [
        "", "", "", "", "", "",
        "Def. Portefeuille Négociation", "Composition Portefeuille", "Calcul positions nettes", "", "", ""
    ],
    "ACTIF - Risque Opérationnel": [
        "", "", "", "", "", "",
        "", "", "", "Ventilation lignes de métier", "", ""
    ],
    "ACTIF - Autre / Général": [
        "", "", "", "", "", "",
        "", "", "", "", "", "Calcul base individuelle"
    ],
    "PASSIF - Risque de Crédit": [
        "", "Pondérations (après provisions)", "", "Pondérations souverains/BMD", "", "Méthode risque courant",
        "", "", "", "", "", ""
    ],
    "PASSIF - Risque de Marché": [
        "", "", "", "", "", "",
        "", "Éligibilité négociation", "Méthodes calcul EFP", "", "", ""
    ],
    "PASSIF - Risque Opérationnel": [
        "", "", "", "", "", "",
        "", "", "", "", "Formule KIB", ""
    ],
    "PASSIF - Autre / Général": [
        "", "", "", "", "", "",
        "", "", "Calcul consolidé", "", "Entrée en vigueur", ""
    ]
}

df = pd.DataFrame(data)

# --- 2. BARRE LATÉRALE (FILTRES) ---
st.sidebar.header("🔍 Filtres")

# Filtre par Type de Risque
risques = ["Tous", "Risque de Crédit", "Risque de Marché", "Risque Opérationnel", "Autre / Général"]
selected_risque = st.sidebar.selectbox("Sélectionner le Risque", risques)

# Filtre par Côté (Actif/Passif)
cotes = ["Tous", "ACTIF", "PASSIF"]
selected_cote = st.sidebar.selectbox("Sélectionner le Côté", cotes)

# Recherche par Article
search_article = st.sidebar.text_input("Rechercher un article (ex: 54)")

# --- 3. TRAITEMENT DES FILTRES ---
# On crée une copie pour ne pas modifier l'original
df_filtered = df.copy()

# Logique de filtrage simplifiée pour la démo
# Dans une app réelle, on pourrait "unpivot" le tableau pour filtrer plus facilement
if search_article:
    df_filtered = df_filtered[df_filtered["Article"].str.contains(search_article, case=False)]

# Affichage dynamique des colonnes selon les filtres
columns_to_show = ["Article"]

if selected_risque == "Tous" and selected_cote == "Tous":
    columns_to_show += [col for col in df.columns if col != "Article"]
elif selected_risque != "Tous" and selected_cote == "Tous":
    columns_to_show += [col for col in df.columns if selected_risque in col and col != "Article"]
elif selected_risque == "Tous" and selected_cote != "Tous":
    columns_to_show += [col for col in df.columns if selected_cote in col and col != "Article"]
else:
    # Les deux filtres actifs
    columns_to_show += [col for col in df.columns if selected_risque in col and selected_cote in col and col != "Article"]

df_display = df_filtered[columns_to_show]

# --- 4. AFFICHAGE PRINCIPAL ---
st.subheader("📊 Tableau de Structuration")
st.dataframe(
    df_display, 
    use_container_width=True, 
    height=600,
    hide_index=True
)

# --- 5. EXPORT EXCEL ---
st.subheader("💾 Exporter les données")
buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
    df_display.to_excel(writer, index=False, sheet_name='Reglementation')
    
st.download_button(
    label="Télécharger le tableau filtré (Excel)",
    data=buffer,
    file_name="NT02_Structuration.xlsx",
    mime="application/vnd.ms-excel"
)

# --- 6. PIED DE PAGE ---
st.markdown("---")
st.caption("Document de référence : NT n° 02/DSB/2007 du 13 avril 2007 (Bank Al-Maghrib)")