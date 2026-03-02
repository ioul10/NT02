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
        transition: all 0.3s;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: #f5f5f5;
        border-radius: 8px;
        padding: 10px;
        font-weight: bold;
    }
    
    /* Badge risque */
    .badge-credit { background-color: #2196f3; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; }
    .badge-market { background-color: #ff9800; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; }
    .badge-operational { background-color: #9c27b0; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; }
    .badge-general { background-color: #4caf50; color: white; padding: 4px 12px; border-radius: 20px; font-size: 12px; }
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
        "Art. 4": {
            "titre": "Segmentation - PME",
            "risque": "Crédit",
            "categorie": "Segmentation",
            "texte_complet": """Au sens de la présente notice technique, est considérée comme une créance sur une petite ou moyenne entreprise (PME) toute créance sur une entreprise (y compris les professionnels) dont:
- Le chiffre d'affaires hors taxes individuel, ou celui du groupe d'intérêt auquel elle appartient, est supérieur à 3 millions de dirhams et inférieur ou égal à 50 millions de dirhams
- Le chiffre d'affaires hors taxes individuel est inférieur à 3 millions de dirhams et le montant global des créances est supérieur à 1 million de dirhams"""
        },
        "Art. 5": {
            "titre": "Segmentation - Grande Entreprise",
            "risque": "Crédit",
            "categorie": "Segmentation",
            "texte_complet": """Le portefeuille «grande entreprise» (GE) englobe toutes les créances sur les entreprises, y compris les professionnels, dont le chiffre d'affaires hors taxes individuel, ou celui du groupe d'intérêt auquel elles appartiennent, est supérieur à 50 millions de dirhams.

Sont également incluses dans cette catégorie les créances sur des entreprises faisant partie d'un groupe d'intérêt, pour lesquelles l'établissement n'est pas en mesure de disposer du chiffre d'affaires consolidé du groupe."""
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
        "Art. 11": {
            "titre": "Engagements de hors-bilan",
            "risque": "Crédit",
            "categorie": "Hors-bilan",
            "texte_complet": """Les catégories des engagements de hors bilan sont les suivantes:

A) Risque faible: Engagements révocables sans condition par les établissements, à tout moment et sans préavis.

B) Risque modéré: Accords de refinancement ≤1 an, crédits documentaires import garantis, garanties de bonne exécution.

C) Risque moyen: Accords de refinancement >1 an, facilités d'émission d'effets, engagements de financement de projet, lignes de substitution.

D) Risque élevé: Acceptations de payer, ouvertures de crédit irrévocables, garanties à première demande de nature financière, contre-garanties."""
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
        "Art. 56": {
            "titre": "Composition portefeuille négociation",
            "risque": "Marché",
            "categorie": "Portefeuille",
            "texte_complet": """Le portefeuille de négociation comprend les éléments ci-dessous:
a) Les titres de transaction
b) Les titres de placement (si >10% du bilan)
c) Les produits dérivés
d) Les opérations de cessions temporaires de titres et de change à terme
e) Les autres opérations interbancaires de couverture
f) Les produits de base à l'exclusion de l'or
g) Les instruments de dérivés de crédit"""
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
    """Détermine si l'article est plutôt Actif ou Passif"""
    actif_cols = ["ACTIF - Risque de Crédit", "ACTIF - Risque de Marché", 
                  "ACTIF - Risque Opérationnel", "ACTIF - Autre/Général"]
    passif_cols = ["PASSIF - Risque de Crédit", "PASSIF - Risque de Marché", 
                   "PASSIF - Risque Opérationnel", "PASSIF - Autre/Général"]
    
    actif_count = sum([1 for col in actif_cols if pd.notna(row.get(col, "")) and str(row.get(col, "")).strip() != ""])
    passif_count = sum([1 for col in passif_cols if pd.notna(row.get(col, "")) and str(row.get(col, "")).strip() != ""])
    
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
            article = row.get("Article", "N/A")
            titre = row.get("Titre", "Sans titre")
            resume = row.get("Résumé", "")
            risque = row.get("Risque_Detecte", "Général")
            cote = row.get("Cote_Detectee", "MIXTE")
            
            color_class = get_risque_color(risque)
            icon = get_risque_icon(risque)
            badge_class = get_risque_badge(risque)
            
            # Récupérer le texte complet depuis la DB si disponible
            texte_complet = ""
            if article in articles_db:
                texte_complet = articles_db[article].get("texte_complet", resume)
            else:
                texte_complet = resume
            
            with st.expander(f"{icon} **{article}** - {titre} | <span class='{badge_class}'>{risque}</span> | {cote}", expanded=False):
                # Contenu de la carte
                col1, col2 = st.columns([3, 1])
                
                with col1:
                    st.markdown(f'<div class="{color_class}"><div class="header-title">📝 Résumé</div><div class="article-text">{texte_complet}</div></div>', 
                               unsafe_allow_html=True)
                
                with col2:
                    st.metric("Risque", risque)
                    st.metric("Côté", cote)
                
                # Détails Actif
                st.markdown("#### 📍 ACTIF")
                actif_cols = ["ACTIF - Risque de Crédit", "ACTIF - Risque de Marché", 
                             "ACTIF - Risque Opérationnel", "ACTIF - Autre/Général"]
                for col in actif_cols:
                    if pd.notna(row.get(col, "")) and str(row.get(col, "")).strip() != "":
                        risk_name = col.split(" - ")[1]
                        st.info(f"**{risk_name}**: {row[col]}")
                
                # Détails Passif
                st.markdown("#### 💰 PASSIF")
                passif_cols = ["PASSIF - Risque de Crédit", "PASSIF - Risque de Marché", 
                              "PASSIF - Risque Opérationnel", "PASSIF - Autre/Général"]
                for col in passif_cols:
                    if pd.notna(row.get(col, "")) and str(row.get(col, "")).strip() != "":
                        risk_name = col.split(" - ")[1]
                        st.success(f"**{risk_name}**: {row[col]}")
                
                # Bouton copier
                col_btn1, col_btn2 = st.columns([1, 4])
                with col_btn1:
                    if st.button("📋 Copier", key=f"copy_{article}"):
                        st.toast(f"✅ {article} copié !")
    
    else:
        st.info("📭 Aucun article ne correspond aux filtres sélectionnés.")

# MODE 2: TABLEAU STRUCTURÉ
elif view_option == "📊 Tableau Structuré":
    st.subheader("📊 Tableau de Structuration Complet")
    
    if not df_filtered.empty:
        # Affichage Excel
        st.dataframe(
            df_filtered,
            use_container_width=True,
            height=600,
            hide_index=True
        )
        
        # Export
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
        # Liste déroulante pour sélectionner un article
        article_list = list(articles_db.keys())
        selected_article = st.selectbox(
            "Sélectionner un article",
            options=article_list,
            index=0
        )
        
        if selected_article in articles_db:
            article_data = articles_db[selected_article]
            
            # En-tête de l'article
            col1, col2, col3 = st.columns([3, 1, 1])
            
            with col1:
                st.markdown(f"### {selected_article} - {article_data.get('titre', 'Sans titre')}")
            
            with col2:
                risque = article_data.get('risque', 'Général')
                st.metric("Risque", risque)
            
            with col3:
                categorie = article_data.get('categorie', 'N/A')
                st.metric("Catégorie", categorie)
            
            # Texte complet avec couleur
            color_class = get_risque_color(article_data.get('risque', 'Général'))
            icon = get_risque_icon(article_data.get('risque', 'Général'))
            
            st.markdown(f"""
                <div class="{color_class}">
                    <h4>{icon} Texte Intégral</h4>
                    <div class="article-text">{article_data.get('texte_complet', 'Texte non disponible')}</div>
                </div>
            """, unsafe_allow_html=True)
            
            # Bouton copier
            if st.button("📋 Copier le texte"):
                st.success("✅ Texte copié dans le presse-papier")
        
        # Navigation entre articles
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
        # Camembert par risque
        df["Risque_Detecte"] = df.apply(detect_risque_from_excel, axis=1)
        risque_counts = df["Risque_Detecte"].value_counts()
        
        st.write("**Distribution par Type de Risque**")
        st.bar_chart(risque_counts)
    
    with col2:
        # Barres par côté
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
        # Export JSON (articles DB)
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
