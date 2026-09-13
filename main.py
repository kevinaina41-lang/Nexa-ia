
from fastapi import FastAPI
from groq import Groq
from dotenv import load_dotenv
import os

load_dotenv()

app = FastAPI()
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

@app.get("/")
def racine():
    return {"message": "Nexa AI API fonctionne !"}

@app.post("/chat")
def chat(data: dict):
    message = data.get("message", "")
    response = client.chat.completions.create(
        messages=[
            {"role": "system", "content": "Tu es Nexa AI, créé par Nexora. Tu réponds dans la même langue que l'utilisateur."},
            {"role": "user", "content": message},
        ],
        model="openai/gpt-oss-120b",
    )
    return {"reply": response.choices[0].message.content}
