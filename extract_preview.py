"""
Script d'export des données Onsite/Online par présentation.

Ce module Streamlit permet de se connecter à une base SQL Server,
d'extraire des données de présentations selon différentes vues,
et d'exporter les résultats vers Excel.
"""

import streamlit as st
import pyodbc
import pandas as pd
from io import BytesIO
import re
import math
from typing import Optional, Tuple, Dict, Any
from contextlib import contextmanager
from datetime import datetime

# ==============================================================================
# CONSTANTES
# ==============================================================================

ILLEGAL_XLSX_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
EXCEL_FILENAME_RE = re.compile(r'[\\/:*?"<>|\x00-\x1F]')
EXCEL_SHEET_RE = re.compile(r'[\[\]:\*\?\/\\]')

DEFAULT_SERVER = "sql-projectis.groupe.zzroot.com"
DEFAULT_DATABASE = "congres"
DEFAULT_USERNAME = "congres"
SQL_PORT = 1433
SQL_DRIVER = "ODBC Driver 18 for SQL Server"

MAX_FILENAME_LENGTH = 80
MAX_SHEET_NAME_LENGTH = 31
PREVIEW_ROWS = 100

VIEW_OPTIONS = [
    "Toutes les lignes brutes (uploads + previews)",
    "Dernier par type (max 2 lignes/presentation)",
    "1 ligne par présentation (règle métier)"
]

# ==============================================================================
# REQUÊTES SQL
# ==============================================================================

SQL_ALL_LINES = """
;WITH base_vis AS (
  SELECT
    vis.idIntervntn AS program_intervention_id,
    ses.event_id,
    vis.titreIntervntn AS presentation_title,
    ses.strTitre  AS session_name,
    ses.[start]   AS session_start
  FROM congres.dbo.v_evt_interventions_salles vis
  JOIN congres.dbo.[session] ses ON ses.intIdSession = vis.idSess
  WHERE ses.event_id = ?
),
base_pc AS (
  SELECT
    pc.presentation_id AS program_intervention_id,
    ses.event_id,
    i.strTitre AS presentation_title,
    ses.strTitre  AS session_name,
    ses.[start]   AS session_start
  FROM congres.dbo.v_presentation_client pc
  JOIN congres.dbo.[session] ses ON ses.intIdSession = pc.session_id
  JOIN congres.dbo.t_evt_interventions i ON i.intIdIntervention = pc.presentation_id
  WHERE ses.event_id = ?
),
base AS (
  SELECT DISTINCT program_intervention_id, event_id, presentation_title, session_name, session_start FROM base_vis
  UNION
  SELECT DISTINCT program_intervention_id, event_id, presentation_title, session_name, session_start FROM base_pc
),
map_ids AS (
  SELECT b.program_intervention_id, b.program_intervention_id AS intervention_id
  FROM base b
  UNION
  SELECT b.program_intervention_id, sync.intervention_id
  FROM base b
  JOIN congres.dbo.v_evt_synchronization_interventions sync
    ON sync.program_intervention_id = b.program_intervention_id
  UNION
  SELECT b.program_intervention_id, kc.intervention_id
  FROM base b
  JOIN congres.dbo.v_evt_keynoteCheck kc
    ON kc.customer_presentation_id = b.program_intervention_id
  UNION
  SELECT b.program_intervention_id, pc.intervention_id
  FROM base b
  JOIN congres.dbo.v_evt_powerpointCheck pc
    ON pc.customer_presentation_id = b.program_intervention_id
),
unioned AS (
  SELECT
    b.program_intervention_id AS presentation_id,
    b.presentation_title,
    b.session_name,
    b.session_start,
    s.[date] AS log_date,
    'online'   AS mode,
    'upload'   AS source,
    DATEDIFF(minute, s.[date], b.session_start) AS delay_minutes
  FROM base b
  JOIN map_ids m            ON m.program_intervention_id = b.program_intervention_id
  JOIN congres.dbo.stats s  ON s.presentation_id = m.intervention_id
  WHERE LOWER(LTRIM(RTRIM(s.client))) = 'upload'

  UNION ALL

  SELECT
    b.program_intervention_id AS presentation_id,
    b.presentation_title,
    b.session_name,
    b.session_start,
    COALESCE(p.datDebutIntervention, p.datFinIntervention) AS log_date,
    'onsite'   AS mode,
    'preview'  AS source,
    DATEDIFF(minute, COALESCE(p.datDebutIntervention, p.datFinIntervention), b.session_start) AS delay_minutes
  FROM base b
  JOIN map_ids m ON m.program_intervention_id = b.program_intervention_id
  JOIN congres.dbo.v_evt_preview_intervenant_interventions p
    ON p.idIntervention = m.intervention_id
   AND COALESCE(p.datDebutIntervention, p.datFinIntervention) IS NOT NULL

  UNION ALL

  SELECT
    b.program_intervention_id AS presentation_id,
    b.presentation_title,
    b.session_name,
    b.session_start,
    s.[date] AS log_date,
    'onsite'   AS mode,
    'terminal' AS source,
    DATEDIFF(minute, s.[date], b.session_start) AS delay_minutes
  FROM base b
  JOIN map_ids m            ON m.program_intervention_id = b.program_intervention_id
  JOIN congres.dbo.stats s  ON s.presentation_id = m.intervention_id
  WHERE LOWER(LTRIM(RTRIM(s.client))) = 'terminal'
)
SELECT
  presentation_id, presentation_title, session_name, session_start,
  log_date, mode, source, delay_minutes
FROM unioned
WHERE log_date IS NOT NULL
ORDER BY presentation_id, log_date;
"""

SQL_LAST_PER_TYPE = """
;WITH base AS (
  SELECT DISTINCT
         vis.idIntervntn AS presentation_id,
         ses.event_id,
         vis.titreIntervntn AS presentation_title,
         ses.strTitre    AS session_name,
         ses.[start]     AS session_start
  FROM congres.dbo.v_evt_interventions_salles AS vis
  JOIN congres.dbo.[session] AS ses
    ON ses.intIdSession = vis.idSess
  WHERE ses.event_id = ?
),
u_uploads AS (
  SELECT b.presentation_id, MAX(s.[date]) AS last_upload_date
  FROM base b
  JOIN congres.dbo.stats s
    ON s.presentation_id = b.presentation_id
   AND LOWER(LTRIM(RTRIM(s.client))) = 'upload'
  GROUP BY b.presentation_id
),
p_previews AS (
  SELECT b.presentation_id,
         MAX(COALESCE(p.datDebutIntervention, p.datFinIntervention)) AS last_preview_date
  FROM base b
  JOIN congres.dbo.v_evt_preview_intervenant_interventions p
    ON p.idIntervention = b.presentation_id
  GROUP BY b.presentation_id
),
t_terminals AS (
  SELECT b.presentation_id, MAX(s.[date]) AS last_terminal_date
  FROM base b
  JOIN congres.dbo.stats s
    ON s.presentation_id = b.presentation_id
   AND LOWER(LTRIM(RTRIM(s.client))) = 'terminal'
  GROUP BY b.presentation_id
),
unioned AS (
  SELECT b.presentation_id, b.presentation_title, b.session_name, b.session_start,
         'online' AS mode,  'upload' AS source,  u.last_upload_date AS last_log_date
  FROM base b JOIN u_uploads u ON u.presentation_id = b.presentation_id

  UNION ALL
  SELECT b.presentation_id, b.presentation_title, b.session_name, b.session_start,
         'onsite' AS mode, 'preview' AS source, p.last_preview_date
  FROM base b JOIN p_previews p ON p.presentation_id = b.presentation_id

  UNION ALL
  SELECT b.presentation_id, b.presentation_title, b.session_name, b.session_start,
         'onsite' AS mode, 'terminal' AS source, t.last_terminal_date
  FROM base b JOIN t_terminals t ON t.presentation_id = b.presentation_id
)
SELECT
  presentation_id, presentation_title, session_name, session_start,
  mode, source, last_log_date,
  DATEDIFF(minute, last_log_date, session_start) AS delay_minutes
FROM unioned
ORDER BY session_start, presentation_id, mode, source;
"""

SQL_ONE_LINE_RULE = """
;WITH base AS (
  SELECT DISTINCT
         vis.idIntervntn AS presentation_id,
         ses.event_id,
         vis.titreIntervntn AS presentation_title,
         ses.strTitre    AS session_name,
         ses.[start]     AS session_start,
         i.datMAJ        AS deposit_date
  FROM congres.dbo.v_evt_interventions_salles AS vis
  JOIN congres.dbo.[session] AS ses
    ON ses.intIdSession = vis.idSess
  JOIN congres.dbo.t_evt_interventions AS i
    ON i.intIdIntervention = vis.idIntervntn
  WHERE ses.event_id = ?
),
u_uploads AS (
  SELECT b.presentation_id, MAX(s.[date]) AS last_upload_date
  FROM base b
  JOIN congres.dbo.stats s
    ON s.presentation_id = b.presentation_id
   AND LOWER(LTRIM(RTRIM(s.client))) = 'upload'
  GROUP BY b.presentation_id
),
p_previews AS (
  SELECT b.presentation_id,
         MAX(COALESCE(p.datDebutIntervention, p.datFinIntervention)) AS last_preview_date
  FROM base b
  JOIN congres.dbo.v_evt_preview_intervenant_interventions p
    ON p.idIntervention = b.presentation_id
  GROUP BY b.presentation_id
),
t_terminals AS (
  SELECT b.presentation_id, MAX(s.[date]) AS last_terminal_date
  FROM base b
  JOIN congres.dbo.stats s
    ON s.presentation_id = b.presentation_id
   AND LOWER(LTRIM(RTRIM(s.client))) = 'terminal'
  GROUP BY b.presentation_id
),
merged AS (
  SELECT
    b.presentation_id, b.presentation_title, b.session_name, b.session_start,
    b.deposit_date,
    u.last_upload_date,
    p.last_preview_date,
    t.last_terminal_date,
    CASE
      WHEN COALESCE(p.last_preview_date, '1900-01-01') >= COALESCE(t.last_terminal_date, '1900-01-01')
           THEN p.last_preview_date
      ELSE t.last_terminal_date
    END AS last_onsite_date
  FROM base b
  LEFT JOIN u_uploads  u ON u.presentation_id = b.presentation_id
  LEFT JOIN p_previews p ON p.presentation_id = b.presentation_id
  LEFT JOIN t_terminals t ON t.presentation_id = b.presentation_id
)
SELECT
  presentation_id,
  presentation_title,
  session_name,
  session_start,
  CASE
    WHEN last_onsite_date IS NOT NULL THEN 'onsite'
    WHEN last_upload_date IS NOT NULL THEN 'online'
    ELSE 'unknown'
  END AS mode,
  deposit_date,
  last_upload_date,
  last_preview_date,
  last_terminal_date,
  COALESCE(last_onsite_date, last_upload_date, deposit_date) AS chosen_last_date,
  DATEDIFF(minute, COALESCE(last_onsite_date, last_upload_date, deposit_date), session_start) AS chosen_delay_minutes
FROM merged
ORDER BY session_start, presentation_id;
"""

QUERY_MAP = {
    VIEW_OPTIONS[0]: SQL_ALL_LINES,
    VIEW_OPTIONS[1]: SQL_LAST_PER_TYPE,
    VIEW_OPTIONS[2]: SQL_ONE_LINE_RULE,
}

# ==============================================================================
# UTILITAIRES
# ==============================================================================

def clean_excel_string(value: Any) -> Any:
    """
    Nettoie une chaîne pour la rendre compatible avec Excel.

    Supprime les caractères illégaux dans les fichiers XLSX.

    Args:
        value: Valeur à nettoyer

    Returns:
        Valeur nettoyée (chaîne sans caractères illégaux)
    """
    if value is None:
        return value
    value_str = str(value)
    return ILLEGAL_XLSX_RE.sub(" ", value_str)


def sanitize_filename(name: str, maxlen: int = MAX_FILENAME_LENGTH) -> str:
    """
    Crée un nom de fichier valide à partir d'une chaîne.

    Args:
        name: Nom à sanitiser
        maxlen: Longueur maximale du nom

    Returns:
        Nom de fichier valide
    """
    if not name:
        return "event"
    name = EXCEL_FILENAME_RE.sub(" ", name)
    name = re.sub(r"\s+", "_", name).strip("_")
    return name[:maxlen]


def sanitize_sheet_name(name: str) -> str:
    """
    Crée un nom de feuille Excel valide.

    Args:
        name: Nom à sanitiser

    Returns:
        Nom de feuille valide (max 31 caractères)
    """
    name = EXCEL_SHEET_RE.sub(" ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return (name or "Sheet")[:MAX_SHEET_NAME_LENGTH]


def format_delay_from_minutes(minutes: Optional[float]) -> Optional[str]:
    """
    Formate un délai en minutes vers un format lisible.

    Convertit un nombre de minutes en format 'X jours, Y heures, Z minutes'.
    Les valeurs négatives indiquent un retard.

    Args:
        minutes: Nombre de minutes (peut être négatif ou None)

    Returns:
        Chaîne formatée ou None si la valeur est invalide

    Examples:
        >>> format_delay_from_minutes(1500)
        '1 jours, 1 heures, 0 minutes'
        >>> format_delay_from_minutes(-60)
        '-0 jours, 1 heures, 0 minutes'
    """
    if minutes is None or (isinstance(minutes, float) and math.isnan(minutes)):
        return None

    try:
        total = int(minutes)
    except (ValueError, TypeError):
        return None

    sign = "-" if total < 0 else ""
    total_abs = abs(total)

    days = total_abs // 1440
    hours = (total_abs % 1440) // 60
    mins = total_abs % 60

    return f"{sign}{days} jours, {hours} heures, {mins} minutes"


@contextmanager
def get_sql_connection(server: str, database: str, username: str, password: str):
    """
    Context manager pour gérer les connexions SQL Server.

    Args:
        server: Adresse du serveur
        database: Nom de la base de données
        username: Nom d'utilisateur
        password: Mot de passe

    Yields:
        pyodbc.Connection: Connexion active

    Raises:
        pyodbc.Error: En cas d'erreur de connexion
    """
    conn = None
    try:
        connection_string = (
            f"Driver={{{SQL_DRIVER}}};"
            f"Server={server},{SQL_PORT};"
            f"Database={database};"
            f"UID={username};"
            f"PWD={password};"
            f"TrustServerCertificate=yes;"
        )
        conn = pyodbc.connect(connection_string)
        yield conn
    finally:
        if conn:
            conn.close()


def get_client_info(conn: pyodbc.Connection) -> Optional[Tuple[int, str]]:
    """
    Récupère les informations du client depuis la base.

    Args:
        conn: Connexion à la base de données

    Returns:
        Tuple (id, nom) du client ou None si non trouvé
    """
    try:
        query = "SELECT TOP 1 Id, Name FROM congres.dbo.Client"
        df = pd.read_sql_query(query, conn)
        if not df.empty:
            return int(df.loc[0, 'Id']), str(df.loc[0, 'Name'])
    except Exception as e:
        st.warning(f"Impossible de récupérer les infos client : {e}")
    return None


def get_events(conn: pyodbc.Connection) -> pd.DataFrame:
    """
    Récupère la liste des événements depuis la base.

    Args:
        conn: Connexion à la base de données

    Returns:
        DataFrame contenant les événements (Id, Title)
    """
    query = "SELECT Id, Title FROM congres.dbo.Event ORDER BY StartDate DESC"
    return pd.read_sql_query(query, conn)


def execute_query_with_params(
    conn: pyodbc.Connection,
    query: str,
    event_id: int
) -> pd.DataFrame:
    """
    Execute une requête SQL avec le bon nombre de paramètres.

    Args:
        conn: Connexion à la base de données
        query: Requête SQL (avec des ?)
        event_id: ID de l'événement

    Returns:
        DataFrame avec les résultats
    """
    param_count = query.count("?")
    params = [event_id] * param_count
    return pd.read_sql_query(query, conn, params=params)


def clean_dataframe_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """
    Nettoie un DataFrame pour l'export Excel.

    - Supprime les caractères illégaux
    - Ajoute les colonnes de délai formatées

    Args:
        df: DataFrame à nettoyer

    Returns:
        DataFrame nettoyé
    """
    if df.empty:
        return df

    # Nettoyage des chaînes
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].map(clean_excel_string)

    # Ajout des colonnes de délai formatées
    if "delay_minutes" in df.columns:
        df["delay_pretty"] = df["delay_minutes"].apply(format_delay_from_minutes)

    if "chosen_delay_minutes" in df.columns:
        df["chosen_delay_pretty"] = df["chosen_delay_minutes"].apply(format_delay_from_minutes)

    return df


def deduplicate_dataframe(
    df: pd.DataFrame,
    column: str,
    strategy: str = "last"
) -> Tuple[pd.DataFrame, int]:
    """
    Dédoublonne un DataFrame selon une colonne et une stratégie.

    Args:
        df: DataFrame à dédoublonner
        column: Nom de la colonne sur laquelle dédoublonner
        strategy: Stratégie de dédoublonnage
            - "first": Garde la première occurrence
            - "last": Garde la dernière occurrence
            - "most_recent": Garde l'occurrence la plus récente (selon log_date)

    Returns:
        Tuple (DataFrame dédoublonné, nombre de doublons supprimés)

    Raises:
        ValueError: Si la colonne n'existe pas ou si la stratégie est invalide
    """
    if df.empty:
        return df, 0

    if column not in df.columns:
        raise ValueError(f"La colonne '{column}' n'existe pas dans le DataFrame")

    valid_strategies = ["first", "last", "most_recent"]
    if strategy not in valid_strategies:
        raise ValueError(f"Stratégie invalide. Choisissez parmi : {valid_strategies}")

    initial_count = len(df)
    df_dedup = df.copy()

    if strategy == "most_recent":
        # Identifier la colonne de date à utiliser
        date_col = None
        for col in ["log_date", "last_log_date", "chosen_last_date", "session_start"]:
            if col in df.columns:
                date_col = col
                break

        if date_col is None:
            st.warning(
                "⚠️ Aucune colonne de date trouvée pour la stratégie 'plus récent'. "
                "Utilisation de 'last' à la place."
            )
            df_dedup = df.drop_duplicates(subset=[column], keep="last")
        else:
            # Trier par la colonne de dédoublonnage et la date (desc)
            df = df.sort_values([column, date_col], ascending=[True, False])
            df_dedup = df.drop_duplicates(subset=[column], keep="first")

    elif strategy == "first":
        df_dedup = df.drop_duplicates(subset=[column], keep="first")

    elif strategy == "last":
        df_dedup = df.drop_duplicates(subset=[column], keep="last")

    duplicates_removed = initial_count - len(df_dedup)

    return df_dedup, duplicates_removed


def create_excel_export(df: pd.DataFrame, event_title: str) -> Tuple[BytesIO, str]:
    """
    Crée un export Excel à partir d'un DataFrame.

    Args:
        df: DataFrame à exporter
        event_title: Titre de l'événement

    Returns:
        Tuple (buffer BytesIO, nom du fichier)
    """
    output = BytesIO()
    safe_title = sanitize_filename(event_title or "event")
    sheet_name = sanitize_sheet_name("OnsiteOnline")
    file_name = f"onsite_online_{safe_title}.xlsx"

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

    output.seek(0)
    return output, file_name


def display_mode_distribution(df: pd.DataFrame) -> None:
    """
    Affiche la répartition par mode dans l'interface.

    Args:
        df: DataFrame contenant une colonne 'mode'
    """
    if "mode" not in df.columns:
        return

    dist = df["mode"].value_counts(dropna=False).reset_index()
    dist.columns = ["mode", "count"]

    st.subheader("📊 Répartition par mode")

    # Affichage en colonnes pour un meilleur rendu
    cols = st.columns(len(dist))
    for idx, row in dist.iterrows():
        with cols[idx]:
            st.metric(label=str(row["mode"]).title(), value=int(row["count"]))

    # Table détaillée
    with st.expander("Voir les détails"):
        st.dataframe(dist, use_container_width=True)


# ==============================================================================
# INTERFACE STREAMLIT
# ==============================================================================

def render_sidebar() -> Dict[str, str]:
    """
    Affiche la sidebar et récupère les paramètres de connexion.

    Returns:
        Dictionnaire avec les paramètres de connexion
    """
    st.sidebar.header("🔌 Connexion au serveur SQL")

    config = {
        "server": st.sidebar.text_input("Serveur", DEFAULT_SERVER),
        "database": st.sidebar.text_input("Base de données", DEFAULT_DATABASE),
        "username": st.sidebar.text_input("Utilisateur", DEFAULT_USERNAME),
        "password": st.sidebar.text_input(
            "Mot de passe",
            type="password",
            help="Le mot de passe n'est pas sauvegardé"
        ),
    }

    return config


def main():
    """Point d'entrée principal de l'application."""

    # Configuration de la page
    st.set_page_config(
        page_title="Export Onsite/Online",
        page_icon="📥",
        layout="wide"
    )

    st.title("📥 Export Onsite / Online par Présentation")
    st.markdown("---")

    # Initialisation du session state
    if "connected" not in st.session_state:
        st.session_state.connected = False

    # Sidebar : paramètres de connexion
    config = render_sidebar()

    # Choix de la vue
    st.sidebar.markdown("---")
    st.sidebar.header("📋 Options d'export")
    view_choice = st.sidebar.radio(
        "Vue à générer",
        VIEW_OPTIONS,
        index=0,
        help="Sélectionnez le type de vue pour l'export"
    )

    # Options de dédoublonnage
    st.sidebar.markdown("---")
    st.sidebar.header("🔄 Dédoublonnage")
    enable_dedup = st.sidebar.checkbox(
        "Activer le dédoublonnage",
        value=False,
        help="Supprime les doublons selon une colonne spécifique"
    )

    dedup_column = None
    dedup_strategy = None

    if enable_dedup:
        dedup_column = st.sidebar.selectbox(
            "Colonne de dédoublonnage",
            ["presentation_id", "session_name"],
            index=0,
            help="Colonne sur laquelle dédoublonner les données"
        )

        dedup_strategy = st.sidebar.radio(
            "Stratégie",
            ["first", "last", "most_recent"],
            index=1,
            format_func=lambda x: {
                "first": "Première occurrence",
                "last": "Dernière occurrence",
                "most_recent": "Plus récente (selon date)"
            }[x],
            help="Quelle ligne garder en cas de doublon"
        )

    # Validation des paramètres de connexion
    if not all([config["server"], config["database"], config["username"], config["password"]]):
        st.warning("⚠️ Veuillez remplir tous les champs de connexion dans la sidebar.")
        st.stop()

    # Connexion et sélection de l'événement
    try:
        with get_sql_connection(**config) as conn:
            # Info client
            client_info = get_client_info(conn)
            if client_info:
                st.sidebar.success(f"✅ Client détecté : **{client_info[1]}**")

            # Récupération des événements
            events_df = get_events(conn)

            if events_df.empty:
                st.warning("⚠️ Aucun congrès disponible dans la base.")
                st.stop()

            # Sélection de l'événement
            st.subheader("🎯 Sélectionner un congrès")
            selected_event_title = st.selectbox(
                "Congrès",
                events_df["Title"].tolist(),
                label_visibility="collapsed"
            )

            selected_event_id = int(
                events_df.loc[events_df["Title"] == selected_event_title, "Id"].values[0]
            )

            st.info(f"📌 Événement sélectionné : **{selected_event_title}** (ID: {selected_event_id})")

    except pyodbc.Error as e:
        st.error(f"❌ Erreur de connexion à la base de données : {e}")
        st.stop()
    except Exception as e:
        st.error(f"❌ Erreur inattendue : {e}")
        st.stop()

    # Bouton de génération
    st.markdown("---")

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        generate_button = st.button(
            "🚀 Générer l'export",
            type="primary",
            use_container_width=True
        )

    if generate_button:
        with st.spinner("⏳ Génération de l'export en cours..."):
            try:
                with get_sql_connection(**config) as conn:
                    # Sélection de la requête
                    query = QUERY_MAP[view_choice]

                    # Exécution
                    df = execute_query_with_params(conn, query, selected_event_id)

                    if df.empty:
                        st.info("ℹ️ Aucune donnée trouvée pour ce congrès avec la vue sélectionnée.")
                        st.stop()

                    # Nettoyage
                    df = clean_dataframe_for_excel(df)

                    # Affichage des résultats initiaux
                    initial_count = len(df)
                    st.success(f"✅ **{initial_count}** ligne(s) récupérée(s).")

                    # Dédoublonnage si activé
                    duplicates_removed = 0
                    if enable_dedup and dedup_column and dedup_strategy:
                        try:
                            df, duplicates_removed = deduplicate_dataframe(
                                df,
                                dedup_column,
                                dedup_strategy
                            )
                            if duplicates_removed > 0:
                                st.info(
                                    f"🔄 Dédoublonnage effectué : **{duplicates_removed}** "
                                    f"doublon(s) supprimé(s) sur la colonne '{dedup_column}'. "
                                    f"**{len(df)}** ligne(s) restante(s)."
                                )
                            else:
                                st.info("ℹ️ Aucun doublon détecté.")
                        except ValueError as e:
                            st.error(f"❌ Erreur de dédoublonnage : {e}")

                    # Distribution par mode
                    display_mode_distribution(df)

                    # Aperçu des données
                    st.subheader("👁️ Aperçu des données")
                    st.dataframe(
                        df.head(PREVIEW_ROWS),
                        use_container_width=True,
                        hide_index=True
                    )

                    if len(df) > PREVIEW_ROWS:
                        st.caption(f"Affichage des {PREVIEW_ROWS} premières lignes sur {len(df)}")

                    # Export Excel
                    st.markdown("---")
                    output, file_name = create_excel_export(df, selected_event_title)

                    col1, col2, col3 = st.columns([1, 2, 1])
                    with col2:
                        st.download_button(
                            label="📥 Télécharger l'Excel",
                            data=output,
                            file_name=file_name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )

                    st.success(f"✅ Fichier prêt : **{file_name}**")

            except pyodbc.Error as e:
                st.error(f"❌ Erreur SQL : {e}")
            except Exception as e:
                st.error(f"❌ Erreur lors de l'exécution : {e}")
                st.exception(e)


if __name__ == "__main__":
    main()
