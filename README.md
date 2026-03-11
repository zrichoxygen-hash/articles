# 📊 Document Analyzer - Evaluation System

Application Streamlit pour évaluer automatiquement des documents académiques (PDF, PPTX) avec l'API OpenAI.

## Fonctionnalités

- 📤 **Upload en batch** de documents PDF et PPTX
- 🤖 **Évaluation automatique** avec critères personnalisables
- 💬 **Chat interactif** pour modifier les critères d'évaluation
- 📊 **Tableau des résultats** avec statistiques
- 💾 **Export en Excel** des notes

## Installation Locale

### Prérequis
- Python 3.8+
- Clé API OpenAI

### Étapes

1. **Cloner ou télécharger le référentiel**
```bash
cd "path/to/DOC ANALYSER"
```

2. **Créer un environnement virtuel**
```bash
python -m venv venv
source venv/Scripts/activate  # Windows
```

3. **Installer les dépendances**
```bash
pip install -r requirements.txt
```

4. **Configurer l'API Key**

Créez un fichier `API.env` à la racine:
```
OPENAI_API_KEY=sk-proj-xxxxxxxxxxxx
```

5. **Lancer l'app**
```bash
streamlit run app.py
```

## Déploiement en Ligne

### Avec Streamlit Community Cloud (Recommandé - Gratuit)

1. **Créer un compte GitHub**
   - Allez sur [github.com](https://github.com)
   - Créez un compte gratuit

2. **Créer un nouveau référentiel**
   - Cliquez sur "New Repository"
   - Nommez-le `doc-analyzer`
   - Laissez public
   - Créez le repo

3. **Pousser votre code**
```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/VOTRE_USERNAME/doc-analyzer.git
git push -u origin main
```

4. **Stocker la clé API de manière sûre**
   - Allez sur [share.streamlit.io](https://share.streamlit.io)
   - Connectez-vous avec GitHub
   - Allez dans "Settings" → "Secrets"
   - Ajoutez: `OPENAI_API_KEY=sk-proj-xxxxxxxxxxxx`

5. **Déployer l'app**
   - Sur [share.streamlit.io](https://share.streamlit.io), cliquez "New app"
   - Repository: `doc-analyzer`
   - Branch: `main`
   - Main file path: `app.py`
   - Cliquez "Deploy"

## Configuration de l'API Key

### En local
Créez `API.env`:
```
OPENAI_API_KEY=votre_clé_ici
```

### En ligne (Streamlit Cloud)
1. Ouvrez votre app déployée
2. Menu (≡) → "Settings"
3. "Secrets" → Ajoutez `OPENAI_API_KEY`

## Structure du Projet

```
DOC ANALYSER/
├── app.py                      # App principale Streamlit
├── batch_upload.py            # Upload en batch (CLI)
├── evaluate_documents.py      # Évaluation (CLI)
├── requirements.txt           # Dépendances Python
├── API.env                    # Clés API (à ignorer)
└── .streamlit/
    └── config.toml           # Configuration Streamlit
```

## Utilisation

1. **Upload des documents**
   - Entrez le chemin du dossier contenant PDF/PPTX
   - Cliquez "📤 Uploader les Documents"

2. **Personnaliser les critères** (optionnel)
   - Discutez avec l'IA pour modifier les critères
   - Ou éditez directement la zone de texte

3. **Évaluer**
   - Cliquez "🚀 Évaluer les Documents"
   - Attendez les résultats

4. **Exporter**
   - Cliquez "💾 Télécharger en Excel"

## Coûts

- **Streamlit Cloud**: Gratuit
- **API OpenAI**: Payant (flexible, basé sur l'usage)

## Support

Pour les questions sur Streamlit Cloud: https://docs.streamlit.io/deploy
Pour les questions sur OpenAI: https://platform.openai.com/docs
