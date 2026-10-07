import os
import requests
from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from openai import OpenAI

app = FastAPI()
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
PHONE_ID = os.getenv("PHONE_ID") or os.getenv("PHONE_NUMBER_ID")
client = OpenAI(api_key=os.getenv("OPENAI_KEY") or os.getenv("OPENAI_API_KEY"))

@app.get("/webhook")
async def verify(request: Request):
    mode = request.query_params.get("hub.mode")
    token = request.query_params.get("hub.verify_token")
    challenge = request.query_params.get("hub.challenge")
    if mode == "subscribe" and token == VERIFY_TOKEN:
        return PlainTextResponse(challenge)
    return PlainTextResponse("Error", status_code=403)

@app.post("/webhook")
async def webhook(request: Request):
    try:
        data = await request.json()
    except:
        return PlainTextResponse("ok", status_code=200)
    try:
        entry = data.get("entry", [{}])[0]
        changes = entry.get("changes", [{}])[0]
        value = changes.get("value", {})
        messages = value.get("messages", [])
        if not messages:
            return PlainTextResponse("ok", status_code=200)
        msg = messages[0]
        from_num = msg.get("from")
        msg_type = msg.get("type")
        texto = msg.get("text", {}).get("body", "") if msg_type == "text" else msg_type
        print(f"Recibido Automatyco de {from_num}: {texto}")
        if PHONE_ID and WHATSAPP_TOKEN:
            url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
            headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
            payload = {"messaging_product": "whatsapp", "to": from_num, "text": {"body": f"✅ Recibido Automatyco: {texto}"}}
            r = requests.post(url, json=payload, headers=headers)
            print(f"Respuesta WhatsApp: {r.status_code} - {r.text}")
    except Exception as e:
        print(f"Error procesando: {e}")
    return PlainTextResponse("ok", status_code=200)

@app.get("/")
def home():
    return {"status": "online"}
