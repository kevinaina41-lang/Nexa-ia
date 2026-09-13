from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from groq import Groq
from dotenv import load_dotenv
from docx import Document
from openpyxl import Workbook, load_workbook
from pptx import Presentation
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from pypdf import PdfReader
from duckduckgo_search import DDGS
import os
import io
import requests

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


def rechercher_web(query):
    """Recherche sur Internet avec DuckDuckGo"""
    try:
        with DDGS() as ddgs:
            resultats = list(ddgs.text(query, max_results=5))
        if not resultats:
            return ""
        texte = "Résultats de recherche récents :\n\n"
        for i, r in enumerate(resultats, 1):
            texte += f"{i}. {r.get('title', '')}\n{r.get('body', '')}\nSource : {r.get('href', '')}\n\n"
        return texte
    except Exception:
        return ""


@app.get("/")
def racine():
    return {"message": "Nexa AI API fonctionne !"}


# ============ CHAT ============
@app.post("/chat")
def chat(data: dict):
    message = data.get("message", "")
    historique = data.get("historique", [])
    persona = data.get("persona", "")
    recherche = data.get("recherche", False)

    system = BASE_PROMPT
    if persona:
        system += f" Comportement souhaité : {persona}."

    if recherche:
        contexte = rechercher_web(message)
        if contexte:
            system += "\n\nUtilise ces résultats de recherche pour répondre :\n" + contexte

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


# ============ QUIZ ============
@app.post("/quiz")
def quiz(data: dict):
    sujet = data.get("sujet", "")
    nb = data.get("nb", 5)
    prompt = f"""Crée un quiz de {nb} questions à choix multiples (QCM) sur : {sujet}.

FORMAT STRICT :
Question 1 : [question]
A) [option]
B) [option]
C) [option]
D) [option]
Réponse correcte : [lettre]

(répète pour chaque question, sans autre texte)"""
    contenu = demander(prompt, "Tu crées des quiz QCM en respectant STRICTEMENT le format.")
    return {"reply": contenu, "quiz": contenu}


# ============ FICHE ============
@app.post("/fiche")
def fiche(data: dict):
    sujet = data.get("sujet", "")
    prompt = f"Crée une fiche de révision complète et structurée sur : {sujet}."
    return {"reply": demander(prompt, "Tu crées des fiches de révision claires.")}


# ============ CV ============
@app.post("/cv")
def cv(data: dict):
    infos = data.get("infos", "")
    prompt = f"Crée un CV professionnel moderne basé sur : {infos}"
    return {"reply": demander(prompt, "Tu es un expert en recrutement.")}


# ============ CORRECTION ============
@app.post("/corriger")
def corriger(data: dict):
    texte = data.get("texte", "")
    prompt = f"Corrige ce devoir :\n\n{texte}\n\nDonne : fautes, améliorations, note sur 20, commentaire."
    return {"reply": demander(prompt, "Tu corriges avec bienveillance.")}


# ============ GÉNÉRATION WORD ============
@app.post("/word")
def generer_word(data: dict):
    sujet = data.get("sujet", "")
    contenu = demander(f"Rédige un document complet et structuré sur : {sujet}. Utilise des titres et paragraphes.")

    doc = Document()
    doc.add_heading(sujet, 0)
    for ligne in contenu.split("\n"):
        ligne = ligne.strip()
        if not ligne:
            continue
        if ligne.startswith("# "):
            doc.add_heading(ligne[2:], level=1)
        elif ligne.startswith("## "):
            doc.add_heading(ligne[3:], level=2)
        elif ligne.startswith("- "):
            doc.add_paragraph(ligne[2:], style='List Bullet')
        else:
            doc.add_paragraph(ligne)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=nexa_document.docx"}
    )


# ============ GÉNÉRATION EXCEL ============
@app.post("/excel")
def generer_excel(data: dict):
    sujet = data.get("sujet", "")
    contenu = demander(f"Crée un tableau de données sur : {sujet}. Format CSV simple : en-têtes puis lignes séparées par des virgules. Aucun autre texte.")

    wb = Workbook()
    ws = wb.active
    ws.title = (sujet[:30] or "Données")

    for ligne in contenu.split("\n"):
        ligne = ligne.strip()
        if not ligne:
            continue
        cellules = [c.strip() for c in ligne.split(",")]
        ws.append(cellules)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=nexa_tableau.xlsx"}
    )


# ============ GÉNÉRATION POWERPOINT ============
@app.post("/pptx")
def generer_pptx(data: dict):
    sujet = data.get("sujet", "")
    contenu = demander(f"""Crée une présentation de 5 à 7 slides sur : {sujet}.

FORMAT STRICT :
Slide 1 : [Titre]
- [point]
- [point]

Slide 2 : [Titre]
- [point]
- [point]

(etc.)""")

    prs = Presentation()
    slides_data = []
    slide_actuelle = None

    for ligne in contenu.split("\n"):
        ligne = ligne.strip()
        if not ligne:
            continue
        if ligne.lower().startswith("slide"):
            if slide_actuelle:
                slides_data.append(slide_actuelle)
            titre = ligne.split(":", 1)[1].strip() if ":" in ligne else ligne
            slide_actuelle = {"titre": titre, "contenu": []}
        elif ligne.startswith("- ") and slide_actuelle:
            slide_actuelle["contenu"].append(ligne[2:])

    if slide_actuelle:
        slides_data.append(slide_actuelle)

    if not slides_data:
        slides_data = [{"titre": sujet, "contenu": [contenu]}]

    for s in slides_data:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = s["titre"]
        if len(slide.placeholders) > 1:
            slide.placeholders[1].text = "\n".join(s["contenu"])

    buffer = io.BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": "attachment; filename=nexa_presentation.pptx"}
    )


# ============ GÉNÉRATION PDF ============
@app.post("/pdf")
def generer_pdf(data: dict):
    sujet = data.get("sujet", "")
    contenu = demander(f"Rédige un document PDF sur : {sujet}.")

    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    largeur, hauteur = A4

    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, hauteur - 50, sujet[:80])

    c.setFont("Helvetica", 11)
    y = hauteur - 90
    for ligne in contenu.split("\n"):
        if y < 50:
            c.showPage()
            c.setFont("Helvetica", 11)
            y = hauteur - 50
        c.drawString(50, y, ligne[:100])
        y -= 15

    c.save()
    buffer.seek(0)
    return StreamingResponse(
        buffer,
        media_type="application/pdf",
        headers={"Content-Disposition": "attachment; filename=nexa_document.pdf"}
    )


# ============ GÉNÉRATION D'IMAGES ============
@app.post("/image")
def generer_image(data: dict):
    prompt = data.get("prompt", "")
    prompt_encode = requests.utils.quote(prompt)
    url = f"https://image.pollinations.ai/prompt/{prompt_encode}?width=1024&height=1024&nologo=true"
    return {"url": url}


# ============ ANALYSE DE FICHIERS ============
@app.post("/analyser")
async def analyser(file: UploadFile = File(...)):
    contenu = await file.read()
    nom = file.filename.lower()

    try:
        if nom.endswith(".docx"):
            doc = Document(io.BytesIO(contenu))
            texte = "\n".join([p.text for p in doc.paragraphs])
        elif nom.endswith(".xlsx"):
            wb = load_workbook(io.BytesIO(contenu))
            ws = wb.active
            lignes = []
            for row in ws.iter_rows(values_only=True):
                lignes.append(", ".join([str(c) if c is not None else "" for c in row]))
            texte = "\n".join(lignes)
        elif nom.endswith(".pdf"):
            reader = PdfReader(io.BytesIO(contenu))
            texte = "\n".join([p.extract_text() for p in reader.pages])
        else:
            texte = contenu.decode("utf-8", errors="ignore")

        analyse = demander(f"Analyse ce document :\n\n{texte[:3000]}\n\nRésume-le, donne les points clés et des recommandations.")
        return {"reply": analyse}
    except Exception as e:
        return {"reply": f"❌ Erreur d'analyse : {str(e)}"}
