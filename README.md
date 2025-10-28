# Export Onsite / Online par Présentation

Application Streamlit pour l'extraction et l'export de données de présentations depuis SQL Server vers Excel.

## Description

Ce script permet de se connecter à une base de données SQL Server (congres), d'extraire des données de présentations selon différentes vues métier, et d'exporter les résultats vers Excel avec un formatage approprié.

### Fonctionnalités principales

- **Connexion sécurisée** à SQL Server avec gestion appropriée des connexions
- **3 vues d'export différentes** :
  - Toutes les lignes brutes (uploads + previews)
  - Dernier événement par type (max 2 lignes/présentation)
  - 1 ligne par présentation (règle métier avec priorité onsite)
- **Dédoublonnage flexible** :
  - Suppression des doublons par presentation_id ou session_name
  - 3 stratégies : première occurrence, dernière occurrence, ou plus récente
  - Statistiques de dédoublonnage en temps réel
- **Export Excel enrichi** :
  - Nettoyage des caractères incompatibles
  - Inclut le titre de présentation pour toutes les vues
  - Inclut l'ID de présentation, nom de session, dates
- **Formatage automatique** des délais (jours, heures, minutes)
- **Interface utilisateur intuitive** avec Streamlit
- **Statistiques visuelles** de répartition par mode (onsite/online)

## Prérequis

### Système

- Python 3.8 ou supérieur
- ODBC Driver 18 for SQL Server
- Accès réseau au serveur SQL

### Installation du driver ODBC (Linux/Ubuntu)

```bash
curl https://packages.microsoft.com/keys/microsoft.asc | apt-key add -
curl https://packages.microsoft.com/config/ubuntu/$(lsb_release -rs)/prod.list > /etc/apt/sources.list.d/mssql-release.list
apt-get update
ACCEPT_EULA=Y apt-get install -y msodbcsql18
```

### Installation du driver ODBC (Windows)

Téléchargez et installez depuis : https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server

### Installation du driver ODBC (macOS)

```bash
brew tap microsoft/mssql-release https://github.com/Microsoft/homebrew-mssql-release
brew update
HOMEBREW_ACCEPT_EULA=Y brew install msodbcsql18
```

## Installation

1. Clonez ce dépôt :
```bash
git clone <url-du-depot>
cd Preview
```

2. Créez un environnement virtuel (recommandé) :
```bash
python -m venv venv
source venv/bin/activate  # Linux/macOS
# ou
venv\Scripts\activate  # Windows
```

3. Installez les dépendances :
```bash
pip install -r requirements.txt
```

## Utilisation

### Lancement de l'application

```bash
streamlit run extract_preview.py
```

L'application s'ouvrira automatiquement dans votre navigateur par défaut (généralement http://localhost:8501).

### Configuration de la connexion

Dans la sidebar (barre latérale) :

1. **Serveur** : Adresse du serveur SQL (par défaut : sql-projectis.groupe.zzroot.com)
2. **Base de données** : Nom de la base (par défaut : congres)
3. **Utilisateur** : Nom d'utilisateur SQL
4. **Mot de passe** : Mot de passe (non sauvegardé)

### Sélection de la vue

Choisissez parmi les 3 vues disponibles :

#### 1. Toutes les lignes brutes
- Affiche toutes les entrées (uploads + previews + terminal)
- Plusieurs lignes possibles par présentation
- Idéal pour l'audit détaillé

#### 2. Dernier par type
- Une ligne par type de source (upload, preview, terminal)
- Maximum 2 lignes par présentation
- Utile pour comparer les modes

#### 3. 1 ligne par présentation (règle métier)
- Une seule ligne par présentation
- Priorité au mode onsite (preview/terminal)
- Vue consolidée pour reporting

### Options de dédoublonnage

L'application propose une fonctionnalité de dédoublonnage pour supprimer les lignes en double selon vos critères :

#### Activation
Cochez la case **"Activer le dédoublonnage"** dans la sidebar pour activer cette fonctionnalité.

#### Colonne de dédoublonnage
Choisissez la colonne sur laquelle identifier les doublons :
- **presentation_id** : Dédoublonne par ID de présentation (recommandé)
- **session_name** : Dédoublonne par nom de session

#### Stratégie de dédoublonnage
Sélectionnez quelle ligne conserver en cas de doublon :

1. **Première occurrence** : Garde la première ligne rencontrée
2. **Dernière occurrence** : Garde la dernière ligne rencontrée
3. **Plus récente (selon date)** : Garde la ligne avec la date la plus récente
   - Utilise automatiquement la colonne de date disponible (log_date, last_log_date, etc.)
   - Idéal pour conserver les données les plus à jour

#### Exemple d'utilisation
- Pour une vue avec plusieurs uploads par présentation, utilisez `presentation_id` + `Plus récente` pour ne garder que le dernier upload de chaque présentation
- Pour éliminer les sessions en double, utilisez `session_name` + stratégie de votre choix

#### Statistiques
Après le dédoublonnage, l'application affiche :
- Le nombre de doublons supprimés
- Le nombre de lignes restantes
- Un message si aucun doublon n'a été détecté

### Export des données

1. Sélectionnez un congrès dans la liste déroulante
2. Cliquez sur **"Générer l'export"**
3. Consultez les statistiques et l'aperçu
4. Téléchargez le fichier Excel avec le bouton **"Télécharger l'Excel"**

### Colonnes exportées

Selon la vue sélectionnée, le fichier Excel contient les colonnes suivantes :

#### Vue "Toutes les lignes brutes"
- `presentation_id` : ID de la présentation
- `presentation_title` : **Titre de la présentation**
- `session_name` : Nom de la session
- `session_start` : Date/heure de début de session
- `log_date` : Date/heure du log (upload/preview/terminal)
- `mode` : Mode (online/onsite)
- `source` : Source (upload/preview/terminal)
- `delay_minutes` : Délai en minutes
- `delay_pretty` : Délai formaté (X jours, Y heures, Z minutes)

#### Vue "Dernier par type"
- `presentation_id` : ID de la présentation
- `presentation_title` : **Titre de la présentation**
- `session_name` : Nom de la session
- `session_start` : Date/heure de début de session
- `mode` : Mode (online/onsite)
- `source` : Source (upload/preview/terminal)
- `last_log_date` : Date du dernier log par type
- `delay_minutes` : Délai en minutes
- `delay_pretty` : Délai formaté

#### Vue "1 ligne par présentation"
- `presentation_id` : ID de la présentation
- `presentation_title` : **Titre de la présentation**
- `session_name` : Nom de la session
- `session_start` : Date/heure de début de session
- `mode` : Mode retenu (onsite/online/unknown)
- `last_upload_date` : Date du dernier upload
- `last_preview_date` : Date du dernier preview
- `last_terminal_date` : Date du dernier terminal
- `chosen_last_date` : Date retenue selon la règle métier
- `chosen_delay_minutes` : Délai en minutes
- `chosen_delay_pretty` : Délai formaté

## Structure du code

### Constantes
- Configuration SQL Server
- Expressions régulières pour la validation
- Mappings des requêtes

### Utilitaires
- `clean_excel_string()` : Nettoyage des caractères illégaux pour Excel
- `sanitize_filename()` : Création de noms de fichiers valides
- `sanitize_sheet_name()` : Création de noms de feuilles Excel valides
- `format_delay_from_minutes()` : Formatage des délais en texte lisible

### Gestion SQL
- `get_sql_connection()` : Context manager pour les connexions
- `get_client_info()` : Récupération des infos client
- `get_events()` : Liste des événements
- `execute_query_with_params()` : Exécution de requêtes paramétrées

### Export et Traitement
- `clean_dataframe_for_excel()` : Nettoyage des DataFrames
- `deduplicate_dataframe()` : Dédoublonnage avec stratégies multiples
- `create_excel_export()` : Création du fichier Excel

### Interface
- `render_sidebar()` : Affichage de la barre latérale
- `display_mode_distribution()` : Statistiques visuelles
- `main()` : Point d'entrée principal

## Améliorations apportées

### Par rapport à la version originale

1. **Architecture** :
   - Code organisé en fonctions réutilisables
   - Séparation claire des responsabilités
   - Context manager pour la gestion des connexions

2. **Qualité du code** :
   - Type hints pour tous les paramètres et retours
   - Docstrings détaillées (format Google)
   - Constantes nommées et configurables
   - Code DRY (Don't Repeat Yourself)

3. **Robustesse** :
   - Gestion d'erreurs améliorée avec messages explicites
   - Fermeture automatique des connexions SQL
   - Validation des entrées
   - Gestion des cas limites (DataFrame vide, etc.)

4. **Expérience utilisateur** :
   - Interface plus claire et organisée
   - Messages d'information contextuels
   - Indicateurs de progression
   - Statistiques visuelles avec métriques
   - Boutons avec icônes et styles

5. **Performance** :
   - Utilisation du session state Streamlit
   - Réutilisation des connexions
   - Requêtes optimisées

6. **Maintenance** :
   - Documentation complète
   - Code facilement testable
   - Structure modulaire
   - Configuration centralisée

## Dépannage

### Erreur de connexion SQL

**Symptôme** : "Erreur de connexion à la base de données"

**Solutions** :
- Vérifiez que le serveur SQL est accessible (ping, telnet)
- Vérifiez les credentials (utilisateur/mot de passe)
- Vérifiez que le driver ODBC 18 est installé
- Vérifiez les règles de firewall

### Erreur "No module named 'pyodbc'"

**Solution** :
```bash
pip install pyodbc
```

### Caractères corrompus dans Excel

**Solution** : Le script nettoie automatiquement les caractères. Si le problème persiste, vérifiez l'encodage de la base de données source.

### Performances lentes

**Solutions** :
- Vérifiez la connexion réseau au serveur SQL
- Limitez la période d'extraction
- Vérifiez les index sur les tables SQL
- Optimisez les requêtes SQL si nécessaire

## Contribuer

Les contributions sont les bienvenues ! Pour contribuer :

1. Forkez le projet
2. Créez une branche pour votre fonctionnalité (`git checkout -b feature/AmazingFeature`)
3. Committez vos changements (`git commit -m 'Add some AmazingFeature'`)
4. Poussez vers la branche (`git push origin feature/AmazingFeature`)
5. Ouvrez une Pull Request

## Licence

Ce projet est à usage interne.

## Support

Pour toute question ou problème, contactez l'équipe de développement.