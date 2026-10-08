import os
import json
import requests
import tempfile
import smtplib
from datetime import datetime
from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from openai import OpenAI
from pptx import Presentation
from pptx.util import Inches
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

app = FastAPI()

# --- CONFIG ---
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
PHONE_ID = os.getenv("PHONE_ID") or os.getenv("PHONE_NUMBER_ID") or "1335575689641146"
OPENAI_KEY = os.getenv("OPENAI_API_KEY") or os.getenv("OPENAI_KEY")
TEMPLATE_PATH = "Reporte.pptx" # tu plantilla que ya subiste

client = OpenAI(api_key=OPENAI_KEY) if OPENAI_KEY else None

# Memoria por técnico
sessions = {}

# --- FUNCIONES WHATSAPP ---
def send_whatsapp(to, text):
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    payload = {"messaging_product": "whatsapp", "to": to, "text": {"body": text}}
    requests.post(url, json=payload, headers=headers)

def send_document(to, file_path):
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    with open(file_path, "rb") as f:
        up = requests.post(f"https://graph.facebook.com/v20.0/{PHONE_ID}/media",
            headers=headers,
            files={"file": f},
            data={"messaging_product": "whatsapp", "type": "document"})
    media_id = up.json().get("id")
    if not media_id:
        print(f"Error subiendo doc: {up.text}")
        return
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    h2 = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    payload = {
        "messaging_product": "whatsapp", "to": to, "type": "document",
        "document": {"id": media_id, "filename": os.path.basename(file_path)}
    }
    requests.post(url, json=payload, headers=h2)

def download_media(media_id):
    info = requests.get(f"https://graph.facebook.com/v20.0/{media_id}",
        headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}).json()
    media_url = info.get("url")
    if not media_url:
        print(f"No se obtuvo url media: {info}")
        return None
    data = requests.get(media_url, headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}).content
    suffix = ".ogg" if "audio" in info.get("mime_type","") else ".jpg"
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(data)
    tmp.close()
    return tmp.name

# --- IA ---
def transcribe_audio(path):
    if not client:
        return "Transcripción no disponible - falta OPENAI_API_KEY"
    with open(path, "rb") as f:
        result = client.audio.transcriptions.create(model="whisper-1", file=f, language="es")
    return result.text

def extract_info(texto_transcrito):
    if not client:
        return {
            "fecha": datetime.now().strftime("%d/%m/%Y"),
            "empresa": "Cliente",
            "trabajo": texto_transcrito
        }
    prompt = f"""Del siguiente texto de un técnico extrae un JSON con exactamente estas 3 claves: fecha, empresa, trabajo_realizado.
Si no hay fecha usa hoy {datetime.now().strftime('%d/%m/%Y')}.
Texto: "{texto_transcrito}"
Responde SOLO el JSON."""
    try:
        r = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0
        )
        txt = r.choices[0].message.content.replace("```json","").replace("```","").strip()
        return json.loads(txt)
    except Exception as e:
        print(f"Error extrayendo info: {e}")
        return {"fecha": datetime.now().strftime("%d/%m/%Y"), "empresa": "Empresa", "trabajo": texto_transcrito}

# --- POWERPOINT ---
def generar_pptx(datos, fotos_paths, output_path):
    if os.path.exists(TEMPLATE_PATH):
        prs = Presentation(TEMPLATE_PATH)
    else:
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        slide.shapes.title.text = f"Reporte {datos.get('empresa','')}"

    # Reemplazar placeholders en todas las diapositivas existentes
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.text = run.text.replace("{FECHA}", datos.get("fecha","")) \
                                        .replace("{EMPRESA}", datos.get("empresa","")) \
                                        .replace("{TRABAJO}", datos.get("trabajo","")) \
                                        .replace("{{FECHA}}", datos.get("fecha","")) \
                                        .replace("{{EMPRESA}}", datos.get("empresa","")) \
                                        .replace("{{TRABAJO}}", datos.get("trabajo",""))

    # Si no se reemplazó nada (plantilla sin marcadores), crear slide resumen
    if not os.path.exists(TEMPLATE_PATH):
        slide = prs.slides.add_slide(prs.slide_layouts[5])
        tx = slide.shapes.add_textbox(Inches(0.5), Inches(1), Inches(9), Inches(2)).text_frame
        tx.text = f"Fecha: {datos.get('fecha')}\nEmpresa: {datos.get('empresa')}\nTrabajo: {datos.get('trabajo')}"

    # Agregar fotos
    for foto in fotos_paths:
        try:
            slide = prs.slides.add_slide(prs.slide_layouts[6]) # blank
            slide.shapes.add_picture(foto, Inches(0.5), Inches(0.5), Inches(9), Inches(5.5))
        except Exception as e:
            print(f"Error agregando foto: {e}")

    prs.save(output_path)
    return output_path

def enviar_correo(ruta_pptx, datos):
    host = os.getenv("SMTP_HOST")
    if not host:
        print("SMTP_HOST no configurado, no se envía correo")
        return
    to_email = os.getenv("REPORT_EMAIL") or "mencarnacion@automatyco.com"
    # Soporta múltiples correos separados por coma más adelante
    destinatarios = [e.strip() for e in to_email.split(",")]

    msg = MIMEMultipart()
    msg["From"] = os.getenv("SMTP_USER")
    msg["To"] = ", ".join(destinatarios)
    msg["Subject"] = f"Reporte Automatyco - {datos.get('empresa')} - {datos.get('fecha')}"
    msg.attach(MIMEText(f"Fecha: {datos.get('fecha')}\nEmpresa: {datos.get('empresa')}\nTrabajo realizado: {datos.get('trabajo')}\n\nGenerado automático desde WhatsApp.", "plain", "utf-8"))

    with open(ruta_pptx, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.presentationml.presentation")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(ruta_pptx)}")
        msg.attach(part)

    with smtplib.SMTP(host, int(os.getenv("SMTP_PORT","587"))) as s:
        s.starttls()
        s.login(os.getenv("SMTP_USER"), os.getenv("SMTP_PASS"))
        s.send_message(msg)
    print(f"Correo enviado a {destinatarios}")

# --- WEBHOOKS ---
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
        return PlainTextResponse("ok", 200)
    try:
        value = data["entry"][0]["changes"][0]["value"]
        messages = value.get("messages", [])
        if not messages:
            return PlainTextResponse("ok", 200)

        msg = messages[0]
        from_num = msg["from"]
        mtype = msg["type"]

        if from_num not in sessions:
            sessions[from_num] = {"fotos": [], "datos": {}}

        print(f"Mensaje Automatyco de {from_num}: {mtype}")

        if mtype in ("audio", "voice"):
            media_path = download_media(msg[mtype]["id"])
            texto = transcribe_audio(media_path)
            datos = extract_info(texto)
            sessions[from_num]["datos"] = datos
            send_whatsapp(from_num, f"✅ Audio recibido:\n📅 Fecha: {datos.get('fecha')}\n🏢 Empresa: {datos.get('empresa')}\n🔧 Trabajo: {datos.get('trabajo')}\n\nAhora envía las fotos y escribe *generar reporte*")

        elif mtype == "image":
            media_path = download_media(msg["image"]["id"])
            sessions[from_num]["fotos"].append(media_path)
            send_whatsapp(from_num, f"📸 Foto {len(sessions[from_num]['fotos'])} guardada. Envía más o escribe *generar reporte*")

        elif mtype == "text":
            body = msg["text"]["body"].lower()
            if "generar" in body or "reporte" in body:
                datos = sessions[from_num].get("datos")
                fotos = sessions[from_num].get("fotos", [])
                if not datos:
                    send_whatsapp(from_num, "❌ Primero manda el audio diciendo fecha, empresa y trabajo realizado.")
                    return PlainTextResponse("ok", 200)

                nombre = f"Reporte_{datos.get('empresa','Cliente').replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
                out_path = f"/tmp/{nombre}"
                generar_pptx(datos, fotos, out_path)
                enviar_correo(out_path, datos)
                send_document(from_num, out_path)
                send_whatsapp(from_num, f"✅ Reporte generado y enviado a {os.getenv('REPORT_EMAIL')} con {len(fotos)} fotos.")
                sessions[from_num] = {"fotos": [], "datos": {}} # limpiar
            else:
                send_whatsapp(from_num, "👋 Envía un audio con: fecha, empresa y trabajo realizado para empezar.")

    except Exception as e:
        print(f"Error webhook: {e}")
    return PlainTextResponse("ok", 200)

@app.get("/")
def home():
    return {"status": "online", "template": TEMPLATE_PATH, "sessions": len(sessions)}

@app.get("/privacy")
def privacy():
    return PlainTextResponse("Política de privacidad Automatyco: Uso interno para generación de reportes técnicos.")

@app.get("/terms")
def terms():
    return PlainTextResponse("Términos Automatyco - Uso interno.")
