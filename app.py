import streamlit as st
import pandas as pd
import io
from datetime import datetime

# =============================================================================
# CONFIGURATION DE LA PAGE
# =============================================================================
st.set_page_config(
    page_title="NT 02/DSB/2007 - Analyse Réglementaire",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =============================================================================
# STYLE CSS PERSONNALISÉ (COULEURS PAR RISQUE)
# =============================================================================
st.markdown("""
    <style>
    /* Couleurs par type de risque */
    .credit-box { 
        background: linear-gradient(135deg, #e3f2fd 0%, #bbdefb 100%); 
        padding: 15px; 
        border-radius: 8px; 
        border-left: 6px solid #2196f3;
        margin: 10px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .market-box { 
        background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%); 
        padding: 15px; 
        border-radius: 8px; 
        border-left: 6px solid #ff9800;
        margin: 10px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .operational-box { 
        background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%); 
        padding: 15px; 
        border-radius: 8px; 
        border-left: 6px solid #9c27b0;
        margin: 10px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .general-box { 
        background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); 
        padding: 15px; 
        border-radius: 8px; 
        border-left: 6px solid #4caf50;
        margin: 10px 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    /* En-têtes de colonnes */
    .header-actif { 
        background-color: #1976d2 !important; 
        color: white !important;
        font-weight: bold;
    }
    .header-passif { 
        background-color: #388e3c !important; 
        color: white !important;
        font-weight: bold;
    }
    
    /* Métriques */
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        text-align: center;
        margin: 10px 0;
    }
    
    /* Tableau */
    .stDataFrame { 
        border: 1px solid #ddd; 
        border-radius: 8px;
    }
    
    /* Boutons */
    .stButton > button {
        border-radius: 8px;
        font-weight: bold;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: #f5f5f5;
        border-radius: 8px;
        padding: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# =============================================================================
# CHARGEMENT DES DONNÉES
# =============================================================================
@st.cache_data
def load_data():
    """Charge le fichier Excel avec cache pour performance"""
    try:
        df = pd.read_excel("Classeur1.xlsx", sheet_name="Feuil1")
        return df
    except FileNotFoundError:
        st.error("❌ Fichier 'Classeur1.xlsx' non trouvé. Veuillez le placer dans le même dossier que app.py")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erreur de lecture: {str(e)}")
        return pd.DataFrame()

df = load_data()

# =============================================================================
# FONCTIONS UTILITAIRES
# =============================================================================
def get_risque_color(risque):
    """Retourne la classe CSS selon le type de risque"""
    colors = {
        "Crédit": "credit-box",
        "Marché": "market-box",
        "Opérationnel": "operational-box",
        "Général": "general-box"
    }
    return colors.get(risque, "general-box")

def get_risque_icon(risque):
    """Retourne l'icône selon le type de risque"""
    icons = {
        "Crédit": "🛡️",
        "Marché": "📈",
        "Opérationnel": "⚙️",
        "Général": "📄"
    }
    return icons.get(risque, "📋")

def detect_risque(row):
    """Détecte le type de risque principal d'un article"""
    if pd.notna(row.get("ACTIF - Risque de Crédit", "")) or pd.notna(row.get("PASSIF - Risque de Crédit", "")):
        return "Crédit"
    elif pd.notna(row.get("ACTIF - Risque de Marché", "")) or pd.notna(row.get("PASSIF - Risque de Marché", "")):
        return "Marché"
    elif pd.notna(row.get("ACTIF - Risque Opérationnel", "")) or pd.notna(row.get("PASSIF - Risque Opérationnel", "")):
        return "Opérationnel"
    else:
        return "Général"

def get_actif_passif(row):
    """Détermine si l'article est plutôt Actif ou Passif"""
    actif_cols = ["ACTIF - Risque de Crédit", "ACTIF - Risque de Marché", 
                  "ACTIF - Risque Opérationnel", "ACTIF - Autre/Général"]
    passif_cols = ["PASSIF - Risque de Crédit", "PASSIF - Risque de Marché", 
                   "PASSIF - Risque Opérationnel", "PASSIF - Autre/Général"]
    
    actif_count = sum([1 for col in actif_cols if pd.notna(row.get(col, "")) and row.get(col, "") != ""])
    passif_count = sum([1 for col in passif_cols if pd.notna(row.get(col, "")) and row.get(col, "") != ""])
    
    if actif_count > passif_count:
        return "ACTIF"
    elif passif_count > actif_count:
        return "PASSIF"
    else:
        return "MIXTE"

# =============================================================================
# EN-TÊTE DE L'APPLICATION
# =============================================================================
st.title("🏦 Notice Technique NT 02/DSB/2007")
st.markdown("### **Analyse Interactive des Exigences en Fonds Propres**")
st.markdown("*Circulaire 26/G/2006 - Bank Al-Maghrib*")
st.divider()

# =============================================================================
# BARRE LATÉRALE - FILTRES
# =============================================================================
st.sidebar.header("🔍 Filtres de Recherche")
st.sidebar.markdown("---")

# Filtre par type de risque
st.sidebar.subheader("Type de Risque")
risques_dispo = ["Crédit", "Marché", "Opérationnel", "Général"]
if not df.empty:
    df["Risque_Detecte"] = df.apply(detect_risque, axis=1)
    risques_dispo = df["Risque_Detecte"].unique().tolist()

selected_risques = st.sidebar.multiselect(
    "Sélectionner les risques",
    options=risques_dispo,
    default=risques_dispo,
    key="risque_filter"
)

# Filtre par côté (Actif/Passif)
st.sidebar.subheader("Côté Bilan")
cote_options = ["Tous", "ACTIF", "PASSIF", "MIXTE"]
selected_cote = st.sidebar.selectbox(
    "Filtrer par côté",
    options=cote_options,
    key="cote_filter"
)

# Recherche par article
st.sidebar.subheader("Recherche")
search_article = st.sidebar.text_input(
    "🔎 Numéro d'article",
    placeholder="Ex: Art. 10, 54, 70",
    key="search_article"
)

search_titre = st.sidebar.text_input(
    "🔎 Mot-clé dans le titre",
    placeholder="Ex: pondération, garantie, négociation",
    key="search_titre"
)

# Bouton reset
st.sidebar.markdown("---")
if st.sidebar.button("🔄 Réinitialiser les filtres", use_container_width=True):
    st.rerun()

# =============================================================================
# TRAITEMENT DES FILTRES
# =============================================================================
if not df.empty:
    df_filtered = df.copy()
    
    # Ajout des colonnes détectées
    df_filtered["Risque_Detecte"] = df_filtered.apply(detect_risque, axis=1)
    df_filtered["Cote_Detectee"] = df_filtered.apply(get_actif_passif, axis=1)
    
    # Filtre par risque
    if selected_risques:
        df_filtered = df_filtered[df_filtered["Risque_Detecte"].isin(selected_risques)]
    
    # Filtre par côté
    if selected_cote != "Tous":
        df_filtered = df_filtered[df_filtered["Cote_Detectee"] == selected_cote]
    
    # Filtre par numéro d'article
    if search_article:
        df_filtered = df_filtered[df_filtered["Article"].str.contains(search_article, case=False, na=False)]
    
    # Filtre par titre
    if search_titre:
        df_filtered = df_filtered[df_filtered["Titre"].str.contains(search_titre, case=False, na=False)]
else:
    df_filtered = pd.DataFrame()

# =============================================================================
# TABLEAU DE BORD - MÉTRIQUES
# =============================================================================
st.subheader("📊 Tableau de Bord")

if not df.empty:
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric(
            label="📋 Total Articles",
            value=len(df),
            delta=None
        )
    with col2:
        st.metric(
            label="🔍 Articles Filtrés",
            value=len(df_filtered),
            delta=f"{len(df_filtered)}/{len(df)}"
        )
    with col3:
        credit_count = len(df[df["Risque_Detecte"] == "Crédit"]) if "Risque_Detecte" in df.columns else 0
        st.metric(
            label="🛡️ Risque Crédit",
            value=credit_count,
            delta=None
        )
    with col4:
        market_count = len(df[df["Risque_Detecte"] == "Marché"]) if "Risque_Detecte" in df.columns else 0
        st.metric(
            label="📈 Risque Marché",
            value=market_count,
            delta=None
        )
    with col5:
        op_count = len(df[df["Risque_Detecte"] == "Opérationnel"]) if "Risque_Detecte" in df.columns else 0
        st.metric(
            label="⚙️ Risque Opérationnel",
            value=op_count,
            delta=None
        )
else:
    st.warning("⚠️ Aucune donnée disponible. Vérifiez le fichier Excel.")

st.divider()

# =============================================================================
# AFFICHAGE PRINCIPAL - CARTES INTERACTIVES
# =============================================================================
st.subheader("📑 Liste des Articles Réglementaires")

if not df_filtered.empty:
    # Options d'affichage
    view_mode = st.radio(
        "Mode d'affichage",
        options=["📇 Cartes détaillées", "📊 Tableau compact"],
        horizontal=True,
        key="view_mode"
    )
    
    if view_mode == "📇 Cartes détaillées":
        # Affichage sous forme de cartes expandables
        for idx, row in df_filtered.iterrows():
            risque = row.get("Risque_Detecte", "Général")
            article = row.get("Article", "N/A")
            titre = row.get("Titre", "Sans titre")
            resume = row.get("Résumé", "")
            
            color_class = get_risque_color(risque)
            icon = get_risque_icon(risque)
            cote = row.get("Cote_Detectee", "MIXTE")
            
            with st.expander(f"{icon} **{article}** - {titre} | {risque} | {cote}", expanded=False):
                # Contenu de la carte
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    st.markdown(f'<div class="{color_class}"><strong>📝 Résumé:</strong><br/>{resume}</div>', 
                               unsafe_allow_html=True)
                
                with col2:
                    st.metric("Risque", risque)
                    st.metric("Côté", cote)
                
                # Détails Actif
                st.markdown("#### 📍 ACTIF")
                actif_cols = ["ACTIF - Risque de Crédit", "ACTIF - Risque de Marché", 
                             "ACTIF - Risque Opérationnel", "ACTIF - Autre/Général"]
                for col in actif_cols:
                    if pd.notna(row.get(col, "")) and row.get(col, "") != "":
                        st.info(f"**{col.split(' - ')[1]}**: {row[col]}")
                
                # Détails Passif
                st.markdown("#### 💰 PASSIF")
                passif_cols = ["PASSIF - Risque de Crédit", "PASSIF - Risque de Marché", 
                              "PASSIF - Risque Opérationnel", "PASSIF - Autre/Général"]
                for col in passif_cols:
                    if pd.notna(row.get(col, "")) and row.get(col, "") != "":
                        st.success(f"**{col.split(' - ')[1]}**: {row[col]}")
                
                # Bouton copier
                col_btn1, col_btn2 = st.columns([1, 4])
                with col_btn1:
                    if st.button("📋 Copier", key=f"copy_{article}"):
                        st.toast(f"✅ {article} copié !")
    
    else:
        # Affichage tableau
        display_cols = ["Article", "Titre", "Risque_Detecte", "Cote_Detectee", "Résumé"]
        st.dataframe(
            df_filtered[display_cols],
            use_container_width=True,
            height=600,
            hide_index=True
        )
else:
    st.info("📭 Aucun article ne correspond aux filtres sélectionnés.")

# =============================================================================
# EXPORT DES DONNÉES
# =============================================================================
st.divider()
st.subheader("💾 Exporter les Données")

if not df_filtered.empty:
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Export Excel
        buffer_xlsx = io.BytesIO()
        with pd.ExcelWriter(buffer_xlsx, engine='xlsxwriter') as writer:
            df_filtered.to_excel(writer, index=False, sheet_name='Articles_Filtres')
        
        st.download_button(
            label="📥 Télécharger Excel",
            data=buffer_xlsx,
            file_name=f"NT02_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with col2:
        # Export CSV
        csv_data = df_filtered.to_csv(index=False, encoding='utf-8-sig').encode('utf-8')
        st.download_button(
            label="📥 Télécharger CSV",
            data=csv_data,
            file_name=f"NT02_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col3:
        # Export PDF (simulation)
        st.button("📥 Télécharger PDF", disabled=True, use_container_width=True, help="Fonctionnalité à venir")
else:
    st.warning("⚠️ Aucune donnée à exporter.")

# =============================================================================
# VISUALISATION GRAPHIQUE (OPTIONNELLE)
# =============================================================================
st.divider()
st.subheader("📈 Répartition par Risque")

if not df.empty and "Risque_Detecte" in df.columns:
    col1, col2 = st.columns(2)
    
    with col1:
        # Camembert par risque
        risque_counts = df["Risque_Detecte"].value_counts()
        st.plotly_chart(
            pd.DataFrame({"Risque": risque_counts.index, "Nombre": risque_counts.values}),
            use_container_width=True
        )
    
    with col2:
        # Barres par côté
        cote_counts = df["Cote_Detectee"].value_counts()
        st.plotly_chart(
            pd.DataFrame({"Côté": cote_counts.index, "Nombre": cote_counts.values}),
            use_container_width=True
        )

# =============================================================================
# PIED DE PAGE
# =============================================================================
st.divider()
st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p><strong>🏦 Bank Al-Maghrib - Direction de la Supervision Bancaire</strong></p>
        <p>Notice Technique NT n° 02/DSB/2007 du 13 avril 2007</p>
        <p><em>Développé avec Streamlit | Dernière mise à jour: """ + datetime.now().strftime("%d/%m/%Y") + """</em></p>
    </div>
""", unsafe_allow_html=True)

# =============================================================================
# GESTION DES ERREURS
# =============================================================================
if df.empty:
    st.markdown("""
        ### ❌ Fichier Excel non trouvé
        **Solution:**
        1. Vérifiez que `Classeur1.xlsx` est dans le même dossier que `app.py`
        2. Vérifiez que la feuille s'appelle `Feuil1`
        3. Redémarrez l'application
        """)
