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

@app.get("/")
def racine():
    return {"message": "Nexa AI API fonctionne !"}

@app.post("/chat")
def chat(data: dict):
    message = data.get("message", "")
    historique = data.get("historique", [])
    persona = data.get("persona", "mentor")

    base_prompt = f"Tu es Nexa AI, l'assistant intelligent créé par Nexora, fondée par Randriafenosoa Mamiratiniaina Kevin. {PERSONAS.get(persona, PERSONAS['mentor'])} Tu réponds TOUJOURS dans la même langue que l'utilisateur. Tu te souviens de toute la conversation. Tu ne mentionnes jamais OpenAI ni ChatGPT."

    messages = [{"role": "system", "content": base_prompt}]

    for msg in historique[-20:]:
        role = "user" if msg["type"] == "user" else "assistant"
        messages.append({"role": role, "content": msg["texte"]})

    messages.append({"role": "user", "content": message})

    response = client.chat.completions.create(
        messages=messages,
        model="openai/gpt-oss-120b",
    )
    return {"reply": response.choices[0].message.content}
