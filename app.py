from fastapi import FastAPI
from engine.retriever import search_melc, search_matatag

app = FastAPI(title="Lulu AI Chat Engine")

@app.post("/chat")
def chat(payload: dict):
    curriculum = payload.get("curriculum")
    grade = payload.get("grade")
    query = payload.get("query")

    if not query:
        return {"reply": "Maglagay ng tanong."}

    if curriculum == "MELC":
        answers = search_melc(grade, query)
        source = "DepEd MELC"
    else:
        answers = search_matatag(query)
        source = "DepEd MATATAG"

    if not answers:
        return {
            "reply": "Walang eksaktong tugma sa DepEd curriculum para sa tanong na ito."
        }

    text = "\n".join(f"- {a}" for a in answers)

    return {
        "reply": f"Ayon sa {source}:\n{text}"
    }