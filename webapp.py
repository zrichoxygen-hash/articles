from flask import Flask, render_template_string, request, redirect, url_for, session
import os
import tempfile
from ia_engine import generer_article
import PyPDF2
import docx
import pptx
import pandas as pd
import requests
from bs4 import BeautifulSoup

from flask import session, redirect, url_for, flash
app = Flask(__name__)
app.secret_key = os.environ.get('FLASK_SECRET_KEY', 'dev-secret-key')  # À personnaliser en prod

LOGIN_FORM = '''
<!DOCTYPE html>
<html lang="fr">
<head><meta charset="UTF-8"><title>Connexion</title></head>
<body>
<h2>Connexion requise</h2>
<form method="post">
    <input type="password" name="password" placeholder="Mot de passe" required autofocus>
    <button type="submit">Se connecter</button>
    {% if error %}<div style="color:red;">{{ error }}</div>{% endif %}
</form>
</body></html>
'''

HTML_FORM = '''
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Générateur d'articles IA</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background: #f8f9fa; }
        textarea, input[type="file"] { width: 100%; margin-bottom: 10px; }
        textarea { height: 80px; border-radius: 6px; border: 1px solid #bbb; padding: 8px; }
        .result { margin-top: 30px; padding: 20px; background: #f4f4f4; border-radius: 8px; }
        .btn-main, .btn-redesign {
            background: linear-gradient(90deg, #007bff 60%, #0056b3 100%);
            color: #fff; border: none; border-radius: 6px; padding: 12px 28px;
            font-size: 1.1em; font-weight: bold; cursor: pointer; transition: background 0.2s, box-shadow 0.2s;
            box-shadow: 0 2px 8px #cce3ff;
        }
        .btn-main:hover, .btn-redesign:hover {
            background: linear-gradient(90deg, #0056b3 60%, #007bff 100%);
            box-shadow: 0 4px 16px #b3d1ff;
        }
        .file-list { margin: 8px 0 16px 0; padding: 8px; background: #e9ecef; border-radius: 6px; font-size: 0.98em; }
        .file-list span { display: inline-block; margin-right: 10px; }
        .feedback { margin: 10px 0; color: #007bff; font-size: 0.98em; }
    </style>
</head>
<body>
        <h1>Générateur d'articles IA</h1>
        <form method="post" id="mainForm" enctype="multipart/form-data">
            <label>Vos idées :</label><br>
            <textarea name="ideas" id="ideas" required>{{ ideas or '' }}</textarea><br>
            <label>Prompt de recherche :</label><br>
            <textarea name="prompt" id="prompt" required>{{ prompt or '' }}</textarea><br>
            <label>Joindre des fichiers (PDF, DOCX, TXT, etc.) :</label><br>
            <div style="position:relative;">
                <input type="file" id="documents" multiple style="opacity:0;position:absolute;left:0;top:0;width:100%;height:40px;z-index:2;cursor:pointer;">
                <button type="button" id="customUploadBtn" class="btn-main" style="width:100%;margin-bottom:10px;z-index:1;position:relative;">Sélectionner des fichiers</button>
            </div>
            <div class="file-list" id="fileList"></div>
            <label>Ajouter des liens web (un par ligne) :</label><br>
            <textarea name="links" id="links" style="height:60px;" placeholder="https://exemple.com/article1\nhttps://exemple.com/article2"></textarea><br>
            <button type="submit" class="btn-main">Générer l'article</button>
        </form>
        <script>
        // Sauvegarde et restauration des champs dans localStorage
        const ideasField = document.getElementById('ideas');
        const promptField = document.getElementById('prompt');
        const linksField = document.getElementById('links');
        if(localStorage.getItem('ideas')) ideasField.value = localStorage.getItem('ideas');
        if(localStorage.getItem('prompt')) promptField.value = localStorage.getItem('prompt');
        if(localStorage.getItem('links')) linksField.value = localStorage.getItem('links');
        ideasField.addEventListener('input', () => localStorage.setItem('ideas', ideasField.value));
        promptField.addEventListener('input', () => localStorage.setItem('prompt', promptField.value));
        linksField.addEventListener('input', () => localStorage.setItem('links', linksField.value));

        // Upload moderne avec suppression possible
        const fileInput = document.getElementById('documents');
        const fileList = document.getElementById('fileList');
        const customUploadBtn = document.getElementById('customUploadBtn');
        let filesArray = [];

        customUploadBtn.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (e) => {
            filesArray = Array.from(fileInput.files);
            renderFileList();
        });

        function renderFileList() {
            fileList.innerHTML = '';
            if (filesArray.length === 0) {
                fileList.innerHTML = '<span style="color:#888;">Aucun fichier sélectionné</span>';
                fileInput.value = '';
                return;
            }
            filesArray.forEach((file, idx) => {
                const fileItem = document.createElement('div');
                fileItem.style.display = 'flex';
                fileItem.style.alignItems = 'center';
                fileItem.style.marginBottom = '6px';
                fileItem.innerHTML = `<span style="background:#d1e7dd;padding:4px 10px;border-radius:5px;display:inline-block;margin-right:10px;">${file.name}</span>`;
                const removeBtn = document.createElement('button');
                removeBtn.textContent = '✖';
                removeBtn.type = 'button';
                removeBtn.style.background = '#ffdddd';
                removeBtn.style.color = '#c00';
                removeBtn.style.border = 'none';
                removeBtn.style.borderRadius = '50%';
                removeBtn.style.width = '28px';
                removeBtn.style.height = '28px';
                removeBtn.style.fontWeight = 'bold';
                removeBtn.style.cursor = 'pointer';
                removeBtn.style.marginLeft = '5px';
                removeBtn.addEventListener('click', () => {
                    filesArray.splice(idx, 1);
                    renderFileList();
                });
                fileItem.appendChild(removeBtn);
                fileList.appendChild(fileItem);
            });
        }

        // Avant soumission, reconstruire l'objet FileList à partir de filesArray
        document.getElementById('mainForm').addEventListener('submit', function(e) {
            if (fileInput.files.length !== filesArray.length) {
                // On doit reconstruire un DataTransfer
                const dt = new DataTransfer();
                filesArray.forEach(f => dt.items.add(f));
                fileInput.files = dt.files;
            }
        });
        // Affichage initial
        renderFileList();
        </script>
        {% if feedback_links %}
        <div class="feedback">
            <strong>Résultat extraction des liens :</strong><br>
            {% for line in feedback_links.split('<br>') if line.strip() %}
                {% if 'Succès extraction' in line %}
                    <span style="color:green;font-weight:bold;">✔</span> {{ line|safe }}<br>
                {% elif 'Échec extraction' in line %}
                    <span style="color:red;font-weight:bold;">✖</span> {{ line|safe }}<br>
                {% else %}
                    {{ line|safe }}<br>
                {% endif %}
            {% endfor %}
        </div>
        {% endif %}
        {% if html_output %}
        <div class="result">
            <h2>Aperçu visuel WordPress :</h2>
            <div id="wp-preview" style="background:#fff;max-width:700px;margin:0 auto 20px auto;padding:40px 30px 40px 30px;border:1px solid #ccc;box-shadow:0 2px 8px #eee;font-family:Georgia,serif;line-height:1.7;">
                {{ html_output|safe }}
            </div>
            <h2>HTML à copier pour WordPress :</h2>
            <textarea readonly style="width:100%;height:200px;">{{ html_output }}</textarea>
            <h2>Instructions de modification du design/layout :</h2>
            <form method="post">
                <input type="hidden" name="ideas" value="{{ ideas }}">
                <input type="hidden" name="prompt" value="{{ prompt }}">
                <input type="hidden" name="html_output" value="{{ html_output }}">
                <textarea name="redesign_prompt" id="redesign_prompt" style="width:100%;height:80px;" placeholder="Ex: Mets les titres en bleu, ajoute une bordure..." required>{{ redesign_prompt or '' }}</textarea><br>
                <button type="submit" name="action" value="redesign" class="btn-redesign">Améliorer le design/layout</button>
            </form>
        </div>
        <script>
        // Persistance du champ redesign_prompt
        const redesignField = document.getElementById('redesign_prompt');
        if(localStorage.getItem('redesign_prompt')) redesignField.value = localStorage.getItem('redesign_prompt');
        redesignField.addEventListener('input', () => localStorage.setItem('redesign_prompt', redesignField.value));
        </script>
        {% endif %}
    </body>
    </html>

    '''


def is_logged_in():
    return session.get('logged_in', False)


from functools import wraps

PASSWORD = os.environ.get('APP_PASSWORD', 'Freeredac')  # Mot de passe personnalisé

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


# Route de connexion par mot de passe uniquement
@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    feedback_links = ''
    if request.method == 'POST':
        if request.form.get('password') == PASSWORD:
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            error = 'Mot de passe incorrect.'
    return render_template_string(LOGIN_FORM, error=error)



@app.route('/', methods=['GET', 'POST'])
@login_required
def index():
    if not is_logged_in():
        return redirect(url_for('login'))
    html_output = None
    ideas = ''
    prompt = ''
    redesign_prompt = ''
    documents_content = ''
    links_content = ''
    links_extracted_text = ''
    feedback_links = ''
    if request.method == 'POST':
        # Récupérer les fichiers joints
        if 'documents' in request.files:
            files = request.files.getlist('documents')
            for file in files:
                if file.filename:

                    ext = os.path.splitext(file.filename)[1].lower()
                    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
                        file.save(tmp.name)
                        tmp.seek(0)
                        if ext == '.txt':
                            documents_content += tmp.read().decode('utf-8', errors='ignore') + '\n'
                        elif ext == '.pdf':
                            try:
                                tmp.close()  # Fermer avant lecture externe
                                reader = PyPDF2.PdfReader(tmp.name)
                                for page in reader.pages:
                                    documents_content += page.extract_text() + '\n'
                            except Exception as e:
                                documents_content += f"[Erreur PDF: {e}]\n"
                        elif ext == '.docx':
                            try:
                                tmp.close()
                                doc = docx.Document(tmp.name)
                                for para in doc.paragraphs:
                                    documents_content += para.text + '\n'
                            except Exception as e:
                                documents_content += f"[Erreur DOCX: {e}]\n"
                        elif ext == '.pptx':
                            try:
                                tmp.close()
                                pres = pptx.Presentation(tmp.name)
                                for slide in pres.slides:
                                    for shape in slide.shapes:
                                        if hasattr(shape, "text"):
                                            documents_content += shape.text + '\n'
                            except Exception as e:
                                documents_content += f"[Erreur PPTX: {e}]\n"
                        elif ext in ['.xlsx', '.csv']:
                            try:
                                tmp.close()
                                if ext == '.csv':
                                    df = pd.read_csv(tmp.name)
                                else:
                                    df = pd.read_excel(tmp.name)
                                documents_content += df.to_string(index=False) + '\n'
                            except Exception as e:
                                documents_content += f"[Erreur tableur: {e}]\n"
                        else:
                            tmp.close()
                        os.unlink(tmp.name)
        # Récupérer les liens web et extraire leur contenu
        links = request.form.get('links', '').strip()
        if links:
            links_content = links
            for link in links.splitlines():
                url = link.strip()
                if url:
                    try:
                        resp = requests.get(url, timeout=8)
                        resp.raise_for_status()
                        soup = BeautifulSoup(resp.text, 'html.parser')
                        for script in soup(['script', 'style', 'noscript']):
                            script.decompose()
                        text = soup.get_text(separator=' ', strip=True)
                        text = ' '.join(text.split())
                        links_extracted_text += f"\n[Contenu extrait de {url}]:\n{text[:3000]}\n"
                        feedback_links += f"<b>Succès extraction :</b> {url}<br>"
                    except Exception as e:
                        links_extracted_text += f"\n[Erreur lors de l'extraction de {url}: {e}]\n"
                        feedback_links += f"<b>Échec extraction :</b> {url} <span style='color:red;'>({e})</span><br>"
        # Si on demande un redesign
        if request.form.get('action') == 'redesign':
            redesign_prompt = request.form['redesign_prompt']
            html_output_old = request.form['html_output']
            ideas = request.form['ideas']
            prompt = request.form['prompt']
            redesign_full_prompt = (
                f"Voici un article HTML :\n{html_output_old}\n\n{redesign_prompt} "
                "Rends le HTML prêt pour WordPress, professionnel, académique, accessible, structuré, sans balises <html> ou <body>."
            )
            html_output = generer_article(ideas, redesign_full_prompt)
        else:
            ideas = request.form['ideas']
            prompt = request.form['prompt']
            redesign_prompt = ''
            # Combine tout pour le prompt IA
            prompt_html = prompt
            if documents_content:
                prompt_html += f"\n\nVoici le contenu de documents joints :\n{documents_content}"
            if links_content:
                prompt_html += f"\n\nVoici des liens à prendre en compte :\n{links_content}"
            if links_extracted_text:
                prompt_html += f"\n\nVoici le contenu extrait des liens web fournis :\n{links_extracted_text}"
            prompt_html += (
                "\n\nGénère un article au format HTML prêt à être inséré dans WordPress, avec : "
                "<h2> pour les titres principaux, <h3> pour les sous-titres, <p> pour les paragraphes, <ul> pour les listes, <blockquote> pour les citations. "
                "Commence par une introduction claire, structure l'article avec des titres explicites, des paragraphes courts, des exemples concrets, des citations si pertinent, et termine par une conclusion synthétique. "
                "Le style doit être professionnel, académique, intelligent mais accessible à tous. N'utilise pas de balises <html> ou <body>."
            )
            html_output = generer_article(ideas, prompt_html)
    return render_template_string(
        HTML_FORM + '<br><br><a href="/logout">Se déconnecter</a>',
        html_output=html_output, ideas=ideas, prompt=prompt, redesign_prompt=redesign_prompt, feedback_links=feedback_links
    )

if __name__ == '__main__':
    print('Flask démarre...')
    app.run(host='0.0.0.0', port=5000, debug=True)
