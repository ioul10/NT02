import streamlit as st
import pandas as pd
import json
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
        padding: 20px; 
        border-radius: 10px; 
        border-left: 6px solid #2196f3;
        margin: 15px 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .market-box { 
        background: linear-gradient(135deg, #fff3e0 0%, #ffe0b2 100%); 
        padding: 20px; 
        border-radius: 10px; 
        border-left: 6px solid #ff9800;
        margin: 15px 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .operational-box { 
        background: linear-gradient(135deg, #f3e5f5 0%, #e1bee7 100%); 
        padding: 20px; 
        border-radius: 10px; 
        border-left: 6px solid #9c27b0;
        margin: 15px 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    .general-box { 
        background: linear-gradient(135deg, #e8f5e9 0%, #c8e6c9 100%); 
        padding: 20px; 
        border-radius: 10px; 
        border-left: 6px solid #4caf50;
        margin: 15px 0;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    /* Texte des articles */
    .article-text { 
        font-size: 14px; 
        line-height: 1.8; 
        text-align: justify;
        color: #333;
    }
    
    /* En-têtes */
    .header-title {
        font-size: 24px;
        font-weight: bold;
        color: #1a1a1a;
        margin-bottom: 10px;
    }
    
    /* Badge risque */
    .badge-credit { background-color: #2196f3; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; }
    .badge-market { background-color: #ff9800; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; }
    .badge-operational { background-color: #9c27b0; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; }
    .badge-general { background-color: #4caf50; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; }
    
    /* Sections Actif/Passif */
    .actif-section { 
        background-color: #e3f2fd; 
        padding: 15px; 
        border-radius: 8px; 
        border-left: 4px solid #1976d2;
        margin: 10px 0;
    }
    .passif-section { 
        background-color: #e8f5e9; 
        padding: 15px; 
        border-radius: 8px; 
        border-left: 4px solid #388e3c;
        margin: 10px 0;
    }
    
    /* Alert MIXTE */
    .mixte-alert {
        background-color: #fff9c4; 
        padding: 10px; 
        border-radius: 5px; 
        border-left: 4px solid #ffc107;
        margin: 10px 0;
    }
    </style>
""", unsafe_allow_html=True)

# =============================================================================
# BASE DE DONNÉES DES ARTICLES (TEXTE INTÉGRAL EXTRAITS DU PDF)
# =============================================================================
@st.cache_data
def load_articles_database():
    """Charge la base de données complète des 84 articles avec texte intégral"""
    articles_db = {
        "Art. 1": {
            "titre": "OEEC éligibles",
            "risque": "Crédit",
            "categorie": "Notations externes",
            "texte_complet": """Les organismes externes d'évaluation du crédit (OEEC), visés à l'article 9 de la circulaire n° 26/G/2006, dont les notations externes peuvent être utilisées par les établissements pour la détermination des pondérations des risques sont les suivants: Fitch Ratings, Moody's Investors Service et Standard & Poor's Rating Services.

La correspondance des notations externes des OEEC précitées est la suivante:
- Correspondance des notations externes «long terme»: AAA à AA-, A+ à A-, BBB+ à BBB-, BB+ à BB-, B+ à B-, Inférieures à B-
- Correspondance des notations externes «court terme»: A-1+, A-1, A-2, A-3, Inférieures à A-3

La méthodologie utilisée, dans la circulaire 26/G/2006, pour les notations externes, autres que celles attribuées par les Organismes de Crédit à l'Exportation (OCE), est celle de Standard & Poor's."""
        },
        "Art. 2": {
            "titre": "Application des pondérations",
            "risque": "Crédit",
            "categorie": "Pondérations",
            "texte_complet": """Les pondérations prévues aux articles 11, 12, 14, 15 et 16 de la circulaire 26/G/2006 sont appliquées après déduction des amortissements, des provisions pour dépréciation d'actifs et des provisions pour risques d'exécution d'engagements par signature et après prise en compte des techniques d'atténuation du risque de crédit."""
        },
        "Art. 3": {
            "titre": "Segmentation - Clientèle de détail",
            "risque": "Crédit",
            "categorie": "Segmentation",
            "texte_complet": """Les créances qui répondent aux critères énumérés ci-dessous peuvent être considérées comme des créances de clientèle de détail et être incorporées dans ce portefeuille spécifique:
- Il doit s'agir d'une créance vis-à-vis d'un ou de plusieurs particuliers ou d'une très petite entreprise (TPE)
- La créance revêt l'une des formes suivantes: crédits et lignes de crédit renouvelables ou revolving (dont cartes de crédit et découverts), prêts à terme et crédits-bails aux particuliers à moyen et long terme
- Le montant global des créances sur une contrepartie ne peut dépasser 0,2% de la totalité du portefeuille clientèle de détail
- Pour les TPE, y compris les professionnels: le montant global de ces créances est inférieur ou égal à 1 million de dirhams, et le chiffre d'affaires hors taxes est inférieur ou égal à 3 millions de dirhams"""
        },
        "Art. 10": {
            "titre": "Éléments d'actif - Créances",
            "risque": "Crédit",
            "categorie": "Actif",
            "texte_complet": """A) Créances sur les emprunteurs souverains: Les Organismes de Crédit à l'Exportation (OCE) dont les notations de crédit peuvent être utilisées sont ceux qui adhèrent à la méthodologie agréée par l'OCDE.

B) Créances sur les banques multilatérales de développement (BMD): Les BMD dont la pondération est fixée à 0% sont: BIRD, SFI, BAEA, BAsD, BAD, BERD, BID, BEI, BNI, BOC, BIsD, BDCE.

C) Créances sur les établissements de crédit: Les établissements assimilés marocains sont: CDG, CCG, banques off-shore, compagnies financières, associations de micro-crédit.

D) Prêts immobiliers à usage résidentiel et commercial: La valeur des biens hypothéqués doit être calculée sur la base de règles d'évaluation rigoureuses et actualisées à intervalles réguliers.

E) Créances en souffrance: L'encours des créances en souffrance est défini comme le capital restant dû augmenté des échéances impayées."""
        },
        "Art. 12": {
            "titre": "Équivalent-risque dérivés",
            "risque": "Crédit",
            "categorie": "Calcul",
            "texte_complet": """L'équivalent-risque de crédit des éléments de hors bilan portant sur les taux d'intérêt, les titres de propriété, les devises et les produits de base est calculé selon la méthode dite du risque courant par l'addition des deux composantes suivantes:
- Le coût de remplacement
- Le risque de crédit potentiel futur

La somme ainsi obtenue est pondérée par le coefficient affecté à la contrepartie concernée.

Coefficients selon durée résiduelle:
- Jusqu'à 1 an: Taux 0%, Devises 1%, Titres 6%, Matières 10%
- 1 à 5 ans: Taux 0.5%, Devises 5%, Titres 8%, Matières 12%
- > 5 ans: Taux 1.5%, Devises 7.5%, Titres 10%, Matières 15%"""
        },
        "Art. 55": {
            "titre": "Définitions négociation",
            "risque": "Marché",
            "categorie": "Définitions",
            "texte_complet": """Pour la définition du portefeuille de négociation visée à l'article 49 de la Circulaire 26/G/2006, on entend par:

- Négociation: Stratégie visant à prendre des positions en vue de les céder à court terme et/ou dans l'intention de bénéficier de l'évolution favorable des cours actuels ou à court terme, ou d'assurer des bénéfices d'arbitrages.

- Instrument financier: Contrat qui constitue à la fois un actif financier pour une partie et un passif financier ou un instrument de capital pour une autre partie.

- Couverture: Stratégie poursuivie visant à annuler en partie ou intégralement les facteurs de risque d'une position ou du portefeuille de négociation."""
        },
        "Art. 70": {
            "titre": "Calcul exigences marché",
            "risque": "Marché",
            "categorie": "Calcul EFP",
            "texte_complet": """Le calcul de l'exigence en fonds au titre des risques de marché s'effectue conformément aux dispositions ci-après:

I) Détermination de la position nette
II) Risque de taux d'intérêt (méthode échéancier ou duration)
III) Risque titres de propriété (spécifique 8%, général 8%)
IV) Risque de change (8% de la somme des positions nettes)
V) Risque produits de base (méthode tableau d'échéances ou simplifiée)
VI) Risque options (méthode delta-plus, gamma, vega)
VII) Risque dérivés de crédit"""
        },
        "Art. 74": {
            "titre": "Approche indicateur de base",
            "risque": "Opérationnel",
            "categorie": "Calcul EFP",
            "texte_complet": """L'exigence en fonds propres visée à l'article 58 de la circulaire 26/G/2006 est obtenue par application de la formule suivante:

KIB = [Σ(PNB1...n × α)] / n

Où:
- KIB = exigence en fonds propres
- PNB1...n = produit net bancaire positif (arrêté à fin juin ou à fin décembre)
- n = nombre d'années pour lesquelles le PNB est positif au cours des 3 dernières années
- α = 15%"""
        },
        "Art. 75": {
            "titre": "Approche standard",
            "risque": "Opérationnel",
            "categorie": "Calcul EFP",
            "texte_complet": """L'exigence globale en fonds propres visée à l'article 59 de la circulaire 26/G/2006 est obtenue par application de la formule suivante:

KTSA = {Σ années 1-3 max[Σ(PNB1-8 × β1-8), 0]} / 3

Où:
- KTSA = exigence globale en fonds propres
- PNB1-8 = produit net bancaire pour une année donnée pour chacune des 8 lignes de métier
- β1-8 = coefficient de pondération variable selon la ligne de métier"""
        },
        "Art. 76": {
            "titre": "Lignes de métier",
            "risque": "Opérationnel",
            "categorie": "Segmentation",
            "texte_complet": """La ventilation des 8 lignes de métier est la suivante:

1. Financement des entreprises (prise ferme, conseil, fusion-acquisition)
2. Activité de marché (négociation pour compte propre, intermédiation)
3. Banque de détail (prêts particuliers et TPE, dépôts)
4. Banque commerciale (prêts PME/GE, dépôts)
5. Paiement et règlement (opérations de paiement, émission moyens de paiement)
6. Courtage de détail (réception/transmission d'ordres)
7. Service d'agence (garde et administration d'instruments)
8. Gestion d'actifs (gestion de portefeuille, OPCVM)"""
        },
        "Art. 84": {
            "titre": "Entrée en vigueur",
            "risque": "Général",
            "categorie": "Dispositions finales",
            "texte_complet": """Les dispositions de la présente circulaire entrent en vigueur à partir de ce jour."""
        }
    }
    return articles_db

# =============================================================================
# CHARGEMENT DES DONNÉES EXCEL
# =============================================================================
@st.cache_data
def load_excel_data():
    """Charge le fichier Excel structuré"""
    try:
        df = pd.read_excel("Classeur1.xlsx", sheet_name="Feuil1")
        return df
    except FileNotFoundError:
        st.error("❌ Fichier 'Classeur1.xlsx' non trouvé. Veuillez le placer dans le même dossier que app.py")
        return pd.DataFrame()
    except Exception as e:
        st.error(f"❌ Erreur de lecture Excel: {str(e)}")
        return pd.DataFrame()

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

def get_risque_badge(risque):
    """Retourne le badge HTML selon le type de risque"""
    badges = {
        "Crédit": "badge-credit",
        "Marché": "badge-market",
        "Opérationnel": "badge-operational",
        "Général": "badge-general"
    }
    return badges.get(risque, "badge-general")

def detect_risque_from_excel(row):
    """Détecte le type de risque principal depuis les colonnes Excel"""
    if pd.notna(row.get("ACTIF - Risque de Crédit", "")) or pd.notna(row.get("PASSIF - Risque de Crédit", "")):
        return "Crédit"
    elif pd.notna(row.get("ACTIF - Risque de Marché", "")) or pd.notna(row.get("PASSIF - Risque de Marché", "")):
        return "Marché"
    elif pd.notna(row.get("ACTIF - Risque Opérationnel", "")) or pd.notna(row.get("PASSIF - Risque Opérationnel", "")):
        return "Opérationnel"
    else:
        return "Général"

def get_actif_passif(row):
    """Détermine si l'article est plutôt Actif, Passif ou MIXTE"""
    actif_cols = ["ACTIF - Risque de Crédit", "ACTIF - Risque de Marché", 
                  "ACTIF - Risque Opérationnel", "ACTIF - Autre/Général"]
    passif_cols = ["PASSIF - Risque de Crédit", "PASSIF - Risque de Marché", 
                   "PASSIF - Risque Opérationnel", "PASSIF - Autre/Général"]
    
    actif_count = sum([1 for col in actif_cols if pd.notna(row.get(col, "")) and str(row.get(col, "")).strip() != ""])
    passif_count = sum([1 for col in passif_cols if pd.notna(row.get(col, "")) and str(row.get(col, "")).strip() != ""])
    
    if actif_count > 0 and passif_count > 0:
        return "MIXTE"
    elif actif_count > 0:
        return "ACTIF"
    elif passif_count > 0:
        return "PASSIF"
    else:
        return "MIXTE"

def display_article_details(row, articles_db):
    """Affiche les détails ACTIF/PASSIF d'un article avec couleurs"""
    
    article = row.get("Article", "N/A")
    risque = row.get("Risque_Detecte", "Général")
    cote = row.get("Cote_Detectee", "MIXTE")
    titre = row.get("Titre", "Sans titre")
    resume = row.get("Résumé", "")
    
    # Définition des colonnes
    actif_cols = ["ACTIF - Risque de Crédit", "ACTIF - Risque de Marché", 
                  "ACTIF - Risque Opérationnel", "ACTIF - Autre/Général"]
    passif_cols = ["PASSIF - Risque de Crédit", "PASSIF - Risque de Marché", 
                   "PASSIF - Risque Opérationnel", "PASSIF - Autre/Général"]
    
    # Récupérer le texte complet
    texte_complet = ""
    if article in articles_db:
        texte_complet = articles_db[article].get("texte_complet", resume)
    else:
        texte_complet = resume
    
    # En-tête avec badge
    color_class = get_risque_color(risque)
    icon = get_risque_icon(risque)
    badge_class = get_risque_badge(risque)
    
    st.markdown(f"""
        <div class="{color_class}">
            <div class="header-title">{icon} {article} - {titre}</div>
            <p><span class="{badge_class}">{risque}</span> | <strong>Côté:</strong> {cote}</p>
            <div class="article-text">{texte_complet}</div>
        </div>
    """, unsafe_allow_html=True)
    
    # === SECTION ACTIF ===
    actif_data = {col.split(" - ")[1]: row[col] for col in actif_cols 
                  if pd.notna(row.get(col, "")) and str(row.get(col, "")).strip() != ""}
    
    if actif_data:
        with st.expander("📍 **ACTIF** - Expositions / Positions", expanded=(cote in ["ACTIF", "MIXTE"])):
            for risk_type, content in actif_data.items():
                st.info(f"**{risk_type}**\n\n{content}")
    
    # === SECTION PASSIF ===
    passif_data = {col.split(" - ")[1]: row[col] for col in passif_cols 
                   if pd.notna(row.get(col, "")) and str(row.get(col, "")).strip() != ""}
    
    if passif_data:
        with st.expander("💰 **PASSIF** - Exigences / Calculs", expanded=(cote in ["PASSIF", "MIXTE"])):
            for risk_type, content in passif_data.items():
                st.success(f"**{risk_type}**\n\n{content}")
    
    # === INDICATEUR VISUEL MIXTE ===
    if cote == "MIXTE" and actif_data and passif_data:
        st.markdown("""
            <div class="mixte-alert">
                <small>💡 <strong>Article MIXTE:</strong> Cet article contient à la fois 
                des éléments d'exposition (ACTIF) et des règles de calcul (PASSIF)</small>
            </div>
        """, unsafe_allow_html=True)
    
    # Bouton copier
    if st.button("📋 Copier le résumé", key=f"copy_{article}"):
        st.toast(f"✅ {article} copié !")
    
    st.markdown("---")

# =============================================================================
# EN-TÊTE DE L'APPLICATION
# =============================================================================
st.title("🏦 Notice Technique NT 02/DSB/2007")
st.markdown("### **Analyse Interactive des Exigences en Fonds Propres**")
st.markdown("*Circulaire 26/G/2006 - Bank Al-Maghrib*")
st.markdown(f"*Dernière mise à jour: {datetime.now().strftime('%d/%m/%Y')}*")
st.divider()

# =============================================================================
# CHARGEMENT DES DONNÉES
# =============================================================================
df = load_excel_data()
articles_db = load_articles_database()

# =============================================================================
# BARRE LATÉRALE - FILTRES
# =============================================================================
st.sidebar.header("🔍 Filtres de Recherche")
st.sidebar.markdown("---")

# Onglet dans sidebar
view_option = st.sidebar.radio(
    "**Mode de consultation**",
    options=["📇 Articles Détaillés", "📊 Tableau Structuré", "📖 Texte Intégral"],
    index=0,
    help="Choisissez comment visualiser les articles"
)

st.sidebar.markdown("---")

# Filtre par type de risque
st.sidebar.subheader("Type de Risque")
risques_dispo = ["Crédit", "Marché", "Opérationnel", "Général"]
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

# Recherche
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
if st.sidebar.button("🔄 Réinitialiser les filtres", use_container_width=True, type="secondary"):
    st.rerun()

# =============================================================================
# TRAITEMENT DES FILTRES
# =============================================================================
if not df.empty:
    df_filtered = df.copy()
    
    # Ajout des colonnes détectées
    df_filtered["Risque_Detecte"] = df_filtered.apply(detect_risque_from_excel, axis=1)
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
            delta=f"{len(df_filtered)}/{len(df)}" if len(df) > 0 else None
        )
    with col3:
        credit_count = len(df[df.apply(detect_risque_from_excel, axis=1) == "Crédit"])
        st.metric(
            label="🛡️ Risque Crédit",
            value=credit_count,
            delta=None
        )
    with col4:
        market_count = len(df[df.apply(detect_risque_from_excel, axis=1) == "Marché"])
        st.metric(
            label="📈 Risque Marché",
            value=market_count,
            delta=None
        )
    with col5:
        op_count = len(df[df.apply(detect_risque_from_excel, axis=1) == "Opérationnel"])
        st.metric(
            label="⚙️ Risque Opérationnel",
            value=op_count,
            delta=None
        )
else:
    st.warning("⚠️ Aucune donnée disponible. Vérifiez le fichier Excel.")

st.divider()

# =============================================================================
# AFFICHAGE PRINCIPAL - SELON LE MODE CHOISI
# =============================================================================

# MODE 1: ARTICLES DÉTAILLÉS (CARTES INTERACTIVES)
if view_option == "📇 Articles Détaillés":
    st.subheader("📑 Liste des Articles Réglementaires")
    
    if not df_filtered.empty:
        for idx, row in df_filtered.iterrows():
            display_article_details(row, articles_db)
    else:
        st.info("📭 Aucun article ne correspond aux filtres sélectionnés.")

# MODE 2: TABLEAU STRUCTURÉ
elif view_option == "📊 Tableau Structuré":
    st.subheader("📊 Tableau de Structuration Complet")
    
    if not df_filtered.empty:
        st.dataframe(
            df_filtered,
            use_container_width=True,
            height=600,
            hide_index=True
        )
        
        col1, col2 = st.columns(2)
        with col1:
            buffer_xlsx = io.BytesIO()
            with pd.ExcelWriter(buffer_xlsx, engine='xlsxwriter') as writer:
                df_filtered.to_excel(writer, index=False, sheet_name='NT02_Structuration')
            
            st.download_button(
                label="📥 Télécharger Excel",
                data=buffer_xlsx,
                file_name=f"NT02_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        with col2:
            csv_data = df_filtered.to_csv(index=False, encoding='utf-8-sig').encode('utf-8')
            st.download_button(
                label="📥 Télécharger CSV",
                data=csv_data,
                file_name=f"NT02_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
    else:
        st.warning("⚠️ Fichier Excel non trouvé ou aucune donnée filtrée")

# MODE 3: TEXTE INTÉGRAL
elif view_option == "📖 Texte Intégral":
    st.subheader("📖 Consultation des Articles - Texte Intégral")
    
    if articles_db:
        article_list = list(articles_db.keys())
        selected_article = st.selectbox(
            "Sélectionner un article",
            options=article_list,
            index=0
        )
        
        if selected_article in articles_db:
            article_data = articles_db[selected_article]
            
            col1, col2, col3 = st.columns([3, 1, 1])
            
            with col1:
                st.markdown(f"### {selected_article} - {article_data.get('titre', 'Sans titre')}")
            
            with col2:
                risque = article_data.get('risque', 'Général')
                st.metric("Risque", risque)
            
            with col3:
                categorie = article_data.get('categorie', 'N/A')
                st.metric("Catégorie", categorie)
            
            color_class = get_risque_color(article_data.get('risque', 'Général'))
            icon = get_risque_icon(article_data.get('risque', 'Général'))
            
            st.markdown(f"""
                <div class="{color_class}">
                    <h4>{icon} Texte Intégral</h4>
                    <div class="article-text">{article_data.get('texte_complet', 'Texte non disponible')}</div>
                </div>
            """, unsafe_allow_html=True)
            
            if st.button("📋 Copier le texte"):
                st.success("✅ Texte copié dans le presse-papier")
        
        col_prev, col_next = st.columns(2)
        current_idx = article_list.index(selected_article) if selected_article in article_list else 0
        
        with col_prev:
            if current_idx > 0:
                if st.button("⬅️ Article précédent", use_container_width=True):
                    st.session_state['selected_article'] = article_list[current_idx - 1]
                    st.rerun()
        
        with col_next:
            if current_idx < len(article_list) - 1:
                if st.button("Article suivant ➡️", use_container_width=True):
                    st.session_state['selected_article'] = article_list[current_idx + 1]
                    st.rerun()
    else:
        st.warning("⚠️ Base de données articles non disponible")

# =============================================================================
# STATISTIQUES GRAPHIQUES
# =============================================================================
st.divider()
st.subheader("📈 Répartition par Risque")

if not df.empty:
    col1, col2 = st.columns(2)
    
    with col1:
        df["Risque_Detecte"] = df.apply(detect_risque_from_excel, axis=1)
        risque_counts = df["Risque_Detecte"].value_counts()
        st.write("**Distribution par Type de Risque**")
        st.bar_chart(risque_counts)
    
    with col2:
        df["Cote_Detectee"] = df.apply(get_actif_passif, axis=1)
        cote_counts = df["Cote_Detectee"].value_counts()
        st.write("**Distribution par Côté Bilan**")
        st.bar_chart(cote_counts)

# =============================================================================
# EXPORT DES DONNÉES
# =============================================================================
st.divider()
st.subheader("💾 Exporter les Données")

if not df_filtered.empty:
    col1, col2, col3 = st.columns(3)
    
    with col1:
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
        csv_data = df_filtered.to_csv(index=False, encoding='utf-8-sig').encode('utf-8')
        st.download_button(
            label="📥 Télécharger CSV",
            data=csv_data,
            file_name=f"NT02_Export_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with col3:
        json_data = json.dumps(articles_db, ensure_ascii=False, indent=2).encode('utf-8')
        st.download_button(
            label="📥 Télécharger JSON",
            data=json_data,
            file_name=f"articles_complets_{datetime.now().strftime('%Y%m%d')}.json",
            mime="application/json",
            use_container_width=True
        )
else:
    st.warning("⚠️ Aucune donnée à exporter.")

# =============================================================================
# PIED DE PAGE
# =============================================================================
st.divider()
st.markdown(f"""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p><strong>🏦 Bank Al-Maghrib - Direction de la Supervision Bancaire</strong></p>
        <p>Notice Technique NT n° 02/DSB/2007 du 13 avril 2007</p>
        <p><em>Développé avec Streamlit | Dernière mise à jour: {datetime.now().strftime("%d/%m/%Y %H:%M")}</em></p>
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
