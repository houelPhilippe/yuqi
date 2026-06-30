# Éditeur Markdown

Application desktop pour éditer des fichiers Markdown (`.md`) et Quarto (`.qmd`) avec prévisualisation en temps réel, compilation de diagrammes et support avancé des documents techniques.

---

## Pourquoi cette application ?

Les éditeurs Markdown généralistes (VS Code, Obsidian, Typora…) ne couvrent pas tous les besoins d'un flux de travail documentaire technique. Cet éditeur est conçu pour :

- **Éditer et prévisualiser simultanément** des fichiers `.md` et `.qmd` dans une interface native légère, sans Node.js ni Electron.
- **Compiler des documents Quarto** (HTML, PDF) directement depuis l'éditeur, sans quitter la fenêtre.
- **Générer des diagrammes PlantUML** (.puml) avec aperçu intégré SVG/PNG.
- **Naviguer dans des projets documentaires** complexes via un explorateur de fichiers latéral.
- **Personnaliser finement** la typographie, les tailles de police, la mise en page et les polices par type de contenu (texte, code, tableaux, YAML).

L'application repose sur **pywebview** : une seule dépendance Python, une WebView système, aucun runtime JavaScript externe requis.

---

## Architecture

| Couche | Fichier(s) | Rôle |
|--------|-----------|------|
| **Entrée** | `main.py` | Initialise la fenêtre pywebview, enregistre l'API |
| **Backend** | `api.py` | Opérations fichiers, compilation, paramètres |
| **Frontend** | `frontend/index.html` | Structure UI (éditeur + prévisualisation) |
| **Styles** | `frontend/style.css` | Thème Catppuccin, mise en page responsive |
| **Logique** | `frontend/app.js` | Gestion des onglets, rendu Markdown, raccourcis |
| **Librairies** | `marked.min.js`, `turndown.min.js` | Parseur Markdown, convertisseur HTML→Markdown |

---

## Prérequis

- Python 3.9+
- [Quarto](https://quarto.org/) (optionnel — pour la compilation `.qmd`)
- [PlantUML](https://plantuml.com/) + Java (optionnel — pour les diagrammes `.puml`)

## Installation

```bash
python3 -m venv venv
source venv/bin/activate        # Windows : venv\Scripts\activate
pip install -r requirements.txt
```

## Lancer l'application

```bash
python main.py                  # Démarrage simple
python main.py chemin/fichier.md  # Ouvrir un fichier au démarrage
```

Taille de fenêtre par défaut : **1280×800 px** (minimum : 800×600 px).

---

## Fonctionnalités

### Gestion des fichiers et onglets

| Action | Raccourci |
|--------|-----------|
| Ouvrir un fichier | `Ctrl+O` |
| Enregistrer | `Ctrl+S` |
| Enregistrer sous… | `Ctrl+Shift+S` |
| Fermer l'onglet | `Ctrl+W` |
| Nouvel onglet vide | `Ctrl+T` |

- Plusieurs fichiers ouverts **simultanément** dans des onglets.
- **Menu contextuel sur les onglets** : renommer, fermer, fermer les autres.
- Indicateur de modification non enregistrée sur l'onglet.
- Formats supportés : `.md`, `.qmd`, `.html`, `.yml`, `.yaml`, `.css`, `.scss`, `.lua`, `.log`, `.tex`, `.puml`, images.

### Édition Markdown avancée

L'éditeur supporte une syntaxe Markdown étendue :

- **Tableaux grille** avec rendu structuré et styles appliqués.
- **Blocs YAML front-matter** avec coloration syntaxique dédiée.
- **Blocs de code** avec détection automatique du langage et coloration.
- **Shortcodes Quarto** `{{< ... >}}` reconnus et affichés.
- **Figures et images** avec alignement et redimensionnement (`width`, `height`).
- **Notes de bas de page** et références croisées.
- **Blocs div personnalisés** avec classes CSS (`::: {.ma-classe}`).
- **Texte surligné** et coloré.
- **HTML brut** intégré directement dans la source.

### Prévisualisation en temps réel

- Panneau de prévisualisation **latéral redimensionnable**.
- **Synchronisation du défilement** entre l'éditeur et la prévisualisation.
- Mode **source seul** ou **prévisualisation seule** commutable.
- Rendu fidèle du Markdown Quarto étendu (pandoc-like).

### Explorateur de fichiers

- **Panneau latéral** avec navigation dans les répertoires.
- Définition d'un **répertoire par défaut** mémorisé entre sessions.
- Ouverture d'un fichier par double-clic depuis l'arborescence.
- **Recherche dans les fichiers** du répertoire courant.

### Visionneuse d'images

- Affichage intégré des images (PNG, JPG, SVG…).
- **Contrôle du zoom** via curseur ou boutons +/−.

### Compilation Quarto

- **Compilation du fichier courant** en HTML ou PDF via Quarto.
- **Compilation du projet complet** (répertoire projet configurable).
- **Compilation d'un répertoire** entier.
- Panneau de **sortie de compilation** avec log en temps réel.

### Support PlantUML

- **Compilation de fichiers `.puml`** avec aperçu intégré.
- Affichage en **SVG ou PNG** au choix.
- Contrôles de zoom sur le diagramme.
- Configuration du chemin vers le fichier `.jar` PlantUML.

### Minimap et navigation

- **Minimap** du code source (panneau réduit représentant l'ensemble du fichier).
- Largeur de la minimap configurable.

### Autocomplétion CSS

- Extraction automatique des **classes CSS** du projet.
- Suggestions d'autocomplétion lors de la saisie de classes dans la source.

### Paramètres et personnalisation

Tous les paramètres sont sauvegardés dans `settings.json` :

| Paramètre | Description |
|-----------|-------------|
| `projectDir` | Répertoire du projet Quarto |
| `pdfFileName` | Nom du fichier PDF de sortie |
| `plantumlJar` | Chemin vers `plantuml.jar` |
| `zoomDefault` | Niveau de zoom de la prévisualisation (%) |
| `sidebarWidth` / `tocWidth` | Largeurs des panneaux latéraux |
| `fontText` / `fontCode` / `fontTable` / `fontYaml` | Polices par type de contenu |
| `fontSizeText` / `fontSizeCode` / … | Tailles de police par type |
| `lineHeight` | Interligne |
| `justify` | Justification du texte |
| `spellcheck` | Vérification orthographique |
| `lang` | Langue de l'interface (`fr`, `en`…) |

---

## Structure du projet

```
editeur_python/
├── main.py              # Point d'entrée — fenêtre pywebview
├── api.py               # Backend Python — fichiers, compilation, config
├── requirements.txt     # Dépendances Python (pywebview ≥ 4.0)
├── settings.json        # Préférences utilisateur
├── config.json          # État applicatif (dernier répertoire ouvert)
└── frontend/
    ├── index.html       # Structure UI complète
    ├── style.css        # Thème Catppuccin + mise en page
    ├── app.js           # Logique applicative (~7 800 lignes)
    ├── marked.min.js    # Parseur Markdown (local)
    └── turndown.min.js  # Convertisseur HTML→Markdown (local)
```

---

## Dépendances

| Dépendance | Version | Usage |
|-----------|---------|-------|
| [pywebview](https://pywebview.flowrl.com/) | ≥ 4.0 | Fenêtre native WebView |
| [marked.js](https://marked.js.org/) | embarquée | Rendu Markdown → HTML |
| [Turndown.js](https://github.com/mixmark-io/turndown) | embarquée | HTML → Markdown |
| Quarto | externe | Compilation `.qmd` |
| PlantUML + Java | externe | Compilation `.puml` |
