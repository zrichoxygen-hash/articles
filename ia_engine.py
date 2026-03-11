if __name__ == "__main__":
    # Test manuel de la génération d'article IA
    idee_test = "L'impact de l'IA sur l'éducation"
    prompt_test = "Rédige un article structuré, informatif et accessible sur ce sujet."
    print(generer_article(idee_test, prompt_test))

# ia_engine.py
"""
Module pour gérer la génération d'article via une IA (OpenAI, Mistral, etc.)
"""

import os

from openai import OpenAI
from dotenv import load_dotenv

# Charger la clé API depuis API.env
load_dotenv("API.env")

def generer_article(idees: str, prompt: str) -> str:
    """
    Prend les idées et le prompt, appelle l'API IA et retourne l'article généré.
    """
    try:
        client = OpenAI(api_key=os.getenv("OpenAI_API_KEY"))
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "Tu es un rédacteur d'articles expert."},
                {"role": "user", "content": f"Idées: {idees}\nPrompt: {prompt}"}
            ],
            max_tokens=800,
            temperature=0.7
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"[ERREUR IA] {e}"
