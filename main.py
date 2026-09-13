from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from groq import Groq
from dotenv import load_dotenv
import os
import json
import re

load_dotenv()

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

client = Groq(api_key=os.getenv("GROQ_API_KEY"))

BASE_PROMPT = "Tu es Nexa AI, l'assistant intelligent créé par Nexora, fondée par Randriafenosoa Mamiratiniaina Kevin. Tu réponds TOUJOURS dans la même langue que l'utilisateur. Tu ne mentionnes jamais OpenAI ni ChatGPT."

def demander(prompt, system=None):
    messages = [{"role": "system", "content": system or BASE_PROMPT}]
    messages.append({"role": "user", "content": prompt})
    response = client.chat.completions.create(
        messages=messages,
        model="openai/gpt-oss-120b",
    )
    return response.choices[0].message.content

@app.get("/")
def racine():
    return {"message": "Nexa AI API fonctionne !"}

@app.post("/chat")
def chat(data: dict):
    message = data.get("message", "")
    historique = data.get("historique", [])
    persona = data.get("persona", "")

    system = BASE_PROMPT
    if persona:
        system += f" Comportement souhaité : {persona}."

    messages = [{"role": "system", "content": system}]
    for msg in historique[-20:]:
        role = "user" if msg["type"] == "user" else "assistant"
        messages.append({"role": role, "content": msg["texte"]})
    messages.append({"role": "user", "content": message})

    response = client.chat.completions.create(
        messages=messages,
        model="openai/gpt-oss-120b",
    )
    return {"reply": response.choices[0].message.content}

@app.post("/quiz")
def quiz(data: dict):
    sujet = data.get("sujet", "")
    nb = data.get("nb", 5)
    
    prompt = f"""Crée un quiz de {nb} questions à choix multiples (QCM) sur le sujet : {sujet}.

FORMAT STRICT (respecte exactement ce format) :

Question 1 : [ta question]
A) [option 1]
B) [option 2]
C) [option 3]
D) [option 4]
Réponse correcte : [A, B, C ou D]

Question 2 : [ta question]
A) [option 1]
B) [option 2]
C) [option 3]
D) [option 4]
Réponse correcte : [A, B, C ou D]

(continue jusqu'à {nb} questions)

Ne mets AUCUN autre texte. Juste les questions et les options."""
    
    contenu = demander(prompt, "Tu es un professeur qui crée des quiz QCM. Tu respectes STRICTEMENT le format demandé.")
    return {"reply": contenu, "quiz": contenu}

@app.post("/fiche")
def fiche(data: dict):
    sujet = data.get("sujet", "")
    prompt = f"""Crée une fiche de révision complète et structurée sur : {sujet}.

Structure :
1. Introduction
2. Points clés
3. Définitions importantes
4. Exemples
5. À retenir

Format clair, aéré."""
    return {"reply": demander(prompt, "Tu es un professeur qui crée des fiches de révision.")}

@app.post("/cv")
def cv(data: dict):
    infos = data.get("infos", "")
    prompt = f"""Crée un CV professionnel moderne basé sur ces informations :

{infos}

Format :
- En-tête
- Profil
- Formation
- Expérience
- Compétences
- Langues

Ton professionnel."""
    return {"reply": demander(prompt, "Tu es un expert en recrutement.")}

@app.post("/corriger")
def corriger(data: dict):
    texte = data.get("texte", "")
    prompt = f"""Corrige ce devoir :

{texte}

Donne :
1. Les fautes d'orthographe
2. Les fautes de grammaire
3. Les améliorations
4. Une note sur 20
5. Un commentaire encourageant"""
    return {"reply": demander(prompt, "Tu es un professeur qui corrige avec bienveillance.")}
