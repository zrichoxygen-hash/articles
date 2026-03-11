"""
Module pour gérer la génération d'article via l'API OpenAI.
"""

import os

from dotenv import load_dotenv
from openai import OpenAI

# Charge les variables locales si présentes (utile en dev), sans écraser Render.
load_dotenv("API.env", override=False)


def _get_api_key() -> str | None:
    # Priorité au nom standard, puis compat rétro avec l'ancien nom.
    return os.getenv("OPENAI_API_KEY") or os.getenv("OpenAI_API_KEY")


def generer_article(idees: str, prompt: str) -> str:
    """
    Prend les idées et le prompt, appelle l'API IA et retourne l'article généré.
    """
    try:
        api_key = _get_api_key()
        if not api_key:
            return "[ERREUR IA] Clé API manquante. Définis OPENAI_API_KEY dans l'environnement Render."

        model = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        client = OpenAI(api_key=api_key)

        response = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "Tu es un rédacteur d'articles expert."},
                {"role": "user", "content": f"Idées: {idees}\nPrompt: {prompt}"},
            ],
            max_tokens=900,
            temperature=0.7,
        )

        content = (response.choices[0].message.content or "").strip()
        if not content:
            return "[ERREUR IA] Réponse vide du modèle. Vérifie OPENAI_MODEL et les quotas API."
        return content
    except Exception as e:
        return f"[ERREUR IA] {e}"


if __name__ == "__main__":
    idee_test = "L'impact de l'IA sur l'éducation"
    prompt_test = "Rédige un article structuré, informatif et accessible sur ce sujet."
    print(generer_article(idee_test, prompt_test))
