from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from groq import Groq
from dotenv import load_dotenv
import os

load_dotenv()

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

client = Groq(api_key=os.getenv("GROQ_API_KEY"))

PERSONAS = {
    "mentor": "Tu es un mentor bienveillant qui guide l'étudiant avec patience.",
    "ami": "Tu es un ami proche, tu parles de façon décontractée et amicale.",
    "prof": "Tu es un professeur strict mais juste, tu expliques clairement.",
    "coach": "Tu es un coach motivant, tu encourages et pousses à la réussite."
}

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
    else:
        system += " " + PERSONAS.get("mentor", "")

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
    niveau = data.get("niveau", "")
    nb = data.get("nb", 5)
    prompt = f"""Crée un quiz de {nb} questions à choix multiples (QCM) sur le sujet : {sujet}.
Niveau : {niveau}.

Format STRICT :
Question 1 : [question]
A) [réponse]
B) [réponse]
C) [réponse]
D) [réponse]
Réponse correcte : [lettre]

(et ainsi de suite)"""
    return {"reply": demander(prompt, "Tu es un professeur qui crée des quiz pédagogiques.")}

@app.post("/fiche")
def fiche(data: dict):
    sujet = data.get("sujet", "")
    prompt = f"""Crée une fiche de révision complète et structurée sur : {sujet}.

Structure :
1. Introduction
2. Points clés (avec titres)
3. Définitions importantes
4. Exemples
5. À retenir

Format clair, aéré, facile à réviser."""
    return {"reply": demander(prompt, "Tu es un professeur qui crée des fiches de révision claires.")}

@app.post("/cv")
def cv(data: dict):
    nom = data.get("nom", "")
    age = data.get("age", "")
    formation = data.get("formation", "")
    experience = data.get("experience", "")
    competences = data.get("competences", "")
    prompt = f"""Crée un CV professionnel moderne pour :
Nom : {nom}
Âge : {age}
Formation : {formation}
Expérience : {experience}
Compétences : {competences}

Format :
- En-tête (nom, âge)
- Profil / Accroche
- Formation
- Expérience
- Compétences
- Langues

Ton professionnel et personnalisé."""
    return {"reply": demander(prompt, "Tu es un expert en recrutement.")}

@app.post("/corriger")
def corriger(data: dict):
    texte = data.get("texte", "")
    prompt = f"""Corrige ce devoir :

{texte}

Donne :
1. Les fautes d'orthographe
2. Les fautes de grammaire
3. Les améliorations possibles
4. Une note sur 20
5. Un commentaire encourageant"""
    return {"reply": demander(prompt, "Tu es un professeur qui corrige avec bienveillance.")}
