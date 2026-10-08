import os
import json
import requests
import tempfile
import smtplib
import ssl
import re
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
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
PHONE_ID = os.getenv("PHONE_ID") or os.getenv("PHONE_NUMBER_ID") or "1335575689641146"
OPENAI_KEY = os.getenv("OPENAI_API_KEY") or os.getenv("OPENAI_KEY")
TEMPLATE_PATH = "Reporte.pptx"

client = OpenAI(api_key=OPENAI_KEY) if OPENAI_KEY else None
sessions = {}

def send_whatsapp(to, text):
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    requests.post(url, json={"messaging_product": "whatsapp", "to": to, "text": {"body": text}}, headers=headers)

def send_document(to, file_path):
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    with open(file_path, "rb") as f:
        up = requests.post(f"https://graph.facebook.com/v20.0/{PHONE_ID}/media", headers=headers, files={"file": f}, data={"messaging_product": "whatsapp", "type": "document"})
    media_id = up.json().get("id")
    if not media_id: return
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    h2 = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    requests.post(url, json={"messaging_product": "whatsapp", "to": to, "type": "document", "document": {"id": media_id, "filename": os.path.basename(file_path)}}, headers=h2)

def download_media(media_id):
    try:
        info = requests.get(f"https://graph.facebook.com/v20.0/{media_id}", headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}).json()
        url = info.get("url")
        if not url: return None
        data = requests.get(url, headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}).content
        suffix = ".ogg" if "audio" in info.get("mime_type","") else ".jpg"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(data); tmp.close()
        return tmp.name
    except: return None

def transcribe_audio(path):
    if not path: return "ERROR: No se pudo descargar el audio"
    if not client: return "ERROR: Falta OPENAI_API_KEY - usa texto por ahora"
    try:
        with open(path, "rb") as f:
            r = client.audio.transcriptions.create(model="whisper-1", file=f, language="es")
        return r.text
    except Exception as e:
        return f"ERROR TRANSCRIPCION: {e}"

def extract_info(texto, nombre_whatsapp=""):
    texto = texto.strip()
    if client:
        try:
            prompt = f'Extrae JSON con claves: fecha, empresa, trabajo_realizado, tecnico. Usa "{nombre_whatsapp}" como tecnico si no se menciona. Texto: "{texto}" Solo JSON.'
            r = client.chat.completions.create(model="gpt-4o-mini", messages=[{"role": "user", "content": prompt}], temperature=0)
            txt = r.choices[0].message.content.replace("```json","").replace("```","").strip()
            data = json.loads(txt)
            return {"fecha": data.get("fecha") or datetime.now().strftime("%d/%m/%Y"), "empresa": data.get("empresa") or "Empresa", "trabajo": data.get("trabajo_realizado") or texto, "tecnico": data.get("tecnico") or nombre_whatsapp or "Tecnico"}
        except: pass
    fecha = datetime.now().strftime("%d/%m/%Y")
    m = re.search(r"\d{1,2} de \w+ \d{4}", texto, re.IGNORECASE)
    if m: fecha = m.group(0)
    empresa = "Empresa"
    low = texto.lower()
    if "empresa" in low:
        try: empresa = texto.lower().split("empresa")[1].strip().split(",")[0].split(".")[0][:40]
        except: pass
    return {"fecha": fecha, "empresa": empresa.title(), "trabajo": texto, "tecnico": nombre_whatsapp or "Tecnico"}

def generar_pptx(datos, fotos, out_path):
    prs = Presentation(TEMPLATE_PATH) if os.path.exists(TEMPLATE_PATH) else Presentation()
    for slide in prs.slides:
        for shape in slide.shapes:
            if not shape.has_text_frame: continue
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    run.text = run.text.replace("{FECHA}", str(datos.get("fecha",""))).replace("{EMPRESA}", str(datos.get("empresa",""))).replace("{TRABAJO}", str(datos.get("trabajo",""))).replace("{TECNICO}", str(datos.get("tecnico","")))
    for foto in fotos:
        try:
            s = prs.slides.add_slide(prs.slide_layouts[6])
            s.shapes.add_picture(foto, Inches(0.5), Inches(0.5), Inches(9), Inches(5.5))
        except: pass
    prs.save(out_path)
    return out_path

def enviar_correo(ruta, datos):
    host = os.getenv("SMTP_HOST") or "smtp.gmail.com"
    port = int(os.getenv("SMTP_PORT","465"))
    user = (os.getenv("SMTP_USER") or "hal@automatyco.com").strip()
    pwd = (os.getenv("SMTP_PASS") or "lwqd jgli ximy wslf").replace("'","").strip()
    from_e = (os.getenv("SMTP_FROM") or user).strip()
    to_e = os.getenv("REPORT_EMAIL") or "mencarnacion@automatyco.com"
    dests = [e.strip() for e in to_e.split(",")]
    msg = MIMEMultipart()
    msg["From"] = from_e
    msg["To"] = ", ".join(dests)
    msg["Subject"] = f"Reporte {datos.get('empresa')} - {datos.get('tecnico')} - {datos.get('fecha')}"
    msg.attach(MIMEText(f"Técnico: {datos.get('tecnico')}\nFecha: {datos.get('fecha')}\nEmpresa: {datos.get('empresa')}\nTrabajo: {datos.get('trabajo')}", "plain", "utf-8"))
    with open(ruta, "rb") as f:
        part = MIMEBase("application", "vnd.openxmlformats-officedocument.presentationml.presentation")
        part.set_payload(f.read()); encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(ruta)}")
        msg.attach(part)
    if port == 465:
        with smtplib.SMTP_SSL(host, port, context=ssl.create_default_context()) as s:
            s.login(user, pwd); s.send_message(msg)
    else:
        with smtplib.SMTP(host, port) as s:
            s.starttls(); s.login(user, pwd); s.send_message(msg)

@app.get("/webhook")
async def verify(request: Request):
    if request.query_params.get("hub.mode") == "subscribe" and request.query_params.get("hub.verify_token") == VERIFY_TOKEN:
        return PlainTextResponse(request.query_params.get("hub.challenge"))
    return PlainTextResponse("Error", 403)

@app.post("/webhook")
async def webhook(request: Request):
    try: data = await request.json()
    except: return PlainTextResponse("ok", 200)
    try:
        entry = data["entry"][0]["changes"][0]["value"]
        msgs = entry.get("messages", [])
        if not msgs: return PlainTextResponse("ok", 200)
        msg = msgs[0]; from_num = msg["from"]; mtype = msg["type"]
        nombre_perfil = entry.get("contacts", [{}])[0].get("profile", {}).get("name","")
        if from_num not in sessions: sessions[from_num] = {"fotos": [], "datos": {}, "nombre": nombre_perfil}
        sessions[from_num]["nombre"] = nombre_perfil or sessions[from_num].get("nombre","")

        if mtype in ("audio","voice"):
            path = download_media(msg[mtype]["id"])
            texto = transcribe_audio(path)
            if "ERROR" in texto:
                send_whatsapp(from_num, f"⚠️ Audio no transcrito: {texto}\nManda TEXTO ejemplo:\n8 oct 2026 empresa Bimbo se hizo mantenimiento")
                return PlainTextResponse("ok", 200)
            datos = extract_info(texto, sessions[from_num]["nombre"])
            sessions[from_num]["datos"] = datos
            send_whatsapp(from_num, f"✅ Audio de {datos.get('tecnico')} guardado. Ahora fotos y *generar reporte*")
        elif mtype == "image":
            path = download_media(msg["image"]["id"])
            if path:
                sessions[from_num]["fotos"].append(path)
                send_whatsapp(from_num, f"📸 Foto {len(sessions[from_num]['fotos'])} de {sessions[from_num].get('nombre','')} guardada.")
        elif mtype == "text":
            body = msg["text"]["body"].strip()
            if "generar" in body.lower():
                datos = sessions[from_num].get("datos")
                fotos = sessions[from_num].get("fotos", [])
                if not datos:
                    send_whatsapp(from_num, "❌ Primero manda el texto con los datos. Ej: 8 octubre 2026 empresa Bimbo se hizo mantenimiento tablero")
                    return PlainTextResponse("ok", 200)
                out = f"/tmp/Reporte_{datos.get('empresa','').replace(' ','_')}_{datos.get('tecnico','').replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"
                generar_pptx(datos, fotos, out)
                enviar_correo(out, datos)
                send_document(from_num, out)
                send_whatsapp(from_num, f"✅ Reporte de {datos.get('tecnico')} enviado a {os.getenv('REPORT_EMAIL')}")
                sessions[from_num] = {"fotos": [], "datos": {}, "nombre": sessions[from_num].get("nombre","")}
            else:
                datos = extract_info(body, sessions[from_num]["nombre"])
                sessions[from_num]["datos"] = datos
                send_whatsapp(from_num, f"✅ Datos de {datos.get('tecnico')} guardados: {datos.get('fecha')} - {datos.get('empresa')}. Ahora fotos y *generar reporte*")
    except Exception as e:
        print(f"Error: {e}")
    return PlainTextResponse("ok", 200)

@app.get("/")
def home(): return {"status":"online", "mode":"texto"}
@app.get("/privacy")
def privacy(): return PlainTextResponse("Politica privacidad")
@app.get("/terms")
def terms(): return PlainTextResponse("Terminos")
