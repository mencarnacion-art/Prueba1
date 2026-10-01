fastapi
uvicorn
openai
python-pptx
requests

import os, time, json, requests, openai, smtplib
from fastapi import FastAPI, Request
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()

# TUS LLAVES - PONLAS EN RAILWAY COMO VARIABLES
openai.api_key = os.getenv("OPENAI_KEY")
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
PHONE_ID = os.getenv("PHONE_ID")
GMAIL_USER = os.getenv("GMAIL_USER", "mencarnacion@automatyco.com")
GMAIL_APP_PASS = os.getenv("GMAIL_PASS")
DESTINO = "mencarnacion@automatyco.com"

sesiones = {}

def enviar_whatsapp(para, texto):
    url = f"https://graph.facebook.com/v20.0/{PHONE_ID}/messages"
    payload = {"messaging_product":"whatsapp","to":para,"text":{"body":texto}}
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    requests.post(url, json=payload, headers=headers)

def crear_ppt(sesion):
    datos = sesion['datos']; fotos = sesion['fotos']
    prs = Presentation("Reporte.pptx")

    # Slide 1: Empresa
    for shape in prs.slides[0].shapes:
        if shape.has_text_frame and "Empresa" in shape.text:
            shape.text = f"Empresa: {datos.get('cliente','')}"

    # Insertar contenido ANTES de la última diapositiva
    def insertar_antes_del_final():
        xml_list = prs.slides._sldIdLst
        last = list(xml_list)[-1]
        xml_list.remove(last)
        xml_list.insert(len(list(xml_list))-0, last) # placeholder
        return last

    # Slide 2: Datos del audio
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    # Moverla a posición 1
    sldIdLst = prs.slides._sldIdLst
    new_slide = list(sldIdLst)[-1]
    sldIdLst.remove(new_slide)
    sldIdLst.insert(1, new_slide)

    prs.slides[1].shapes.title.text = f"Visita: {datos.get('visita','')} - {datos.get('cliente','')}"
    prs.slides[1].placeholders[1].text = f"PROBLEMA:\n{datos.get('problema','')}\n\nSOLUCIÓN:\n{datos.get('solucion','')}"

    # Slides de fotos
    for i, foto in enumerate(fotos):
        s = prs.slides.add_slide(prs.slide_layouts[5])
        # Mover antes del final
        sldIdLst = prs.slides._sldIdLst
        ns = list(sldIdLst)[-1]
        sldIdLst.remove(ns)
        sldIdLst.insert(len(list(sldIdLst))-1, ns)
        idx = len(prs.slides)-2
        try:
            prs.slides[idx].shapes.add_picture(foto, Inches(0.5), Inches(1.2), Inches(9), Inches(4.5))
        except: pass

    nombre = f"Reporte_{datos.get('cliente','Cliente')}.pptx"
    prs.save(nombre)
    return nombre

def enviar_correo(archivo, datos):
    msg = MIMEMultipart()
    msg['From'] = GMAIL_USER
    msg['To'] = DESTINO
    msg['Subject'] = f"Reporte {datos.get('cliente')} - {datos.get('visita')}"
    msg.attach(MIMEText(f"Cliente: {datos.get('cliente')}\nVisita: {datos.get('visita')}\nProblema: {datos.get('problema')}\n\nAdjunto reporte.", 'plain'))
    with open(archivo, "rb") as f:
        p = MIMEBase('application','octet-stream'); p.set_payload(f.read())
        encoders.encode_base64(p)
        p.add_header('Content-Disposition', f'attachment; filename={archivo}')
        msg.attach(p)
    s = smtplib.SMTP('smtp.gmail.com',587); s.starttls()
    s.login(GMAIL_USER, GMAIL_APP_PASS); s.send_message(msg); s.quit()

@app.post("/webhook")
async def webhook(req: Request):
    data = await req.json()
    try:
        v = data['entry'][0]['changes'][0]['value']
        if 'messages' not in v: return {"ok":True}
        m = v['messages'][0]; de = m['from']
        if de not in sesiones: sesiones[de]={"datos":{}, "fotos":[]}
        if 'audio' in m:
            aid=m['audio']['id']
            url=requests.get(f"https://graph.facebook.com/v20.0/{aid}", headers={"Authorization":f"Bearer {WHATSAPP_TOKEN}"}).json()['url']
            audio=requests.get(url, headers={"Authorization":f"Bearer {WHATSAPP_TOKEN}"}).content
            open("temp.ogg","wb").write(audio)
            with open("temp.ogg","rb") as f:
                t=openai.audio.transcriptions.create(model="whisper-1", file=f, language="es")
            chat=openai.chat.completions.create(model="gpt-4o-mini", messages=[{"role":"user","content":f"Extrae JSON cliente, visita, problema, solucion de: {t.text}. Solo JSON"}])
            sesiones[de]['datos']=json.loads(chat.choices[0].message.content)
            enviar_whatsapp(de, f"✅ Recibido {sesiones[de]['datos'].get('cliente')}. Envía fotos y escribe 'generar'")
        if 'image' in m:
            iid=m['image']['id']
            url=requests.get(f"https://graph.facebook.com/v20.0/{iid}", headers={"Authorization":f"Bearer {WHATSAPP_TOKEN}"}).json()['url']
            img=requests.get(url, headers={"Authorization":f"Bearer {WHATSAPP_TOKEN}"}).content
            ruta=f"/tmp/{de}_{len(sesiones[de]['fotos'])}.jpg"; open(ruta,"wb").write(img)
            sesiones[de]['fotos'].append(ruta)
            enviar_whatsapp(de, f"📸 Foto {len(sesiones[de]['fotos'])} guardada")
        if 'text' in m and 'generar' in m['text']['body'].lower():
            archivo=crear_ppt(sesiones[de])
            enviar_correo(archivo, sesiones[de]['datos'])
            enviar_whatsapp(de, f"✅ Reporte de {sesiones[de]['datos'].get('cliente')} generado y enviado a {DESTINO} por correo.")
            sesiones[de]={"datos":{}, "fotos":[]}
    except Exception as e: print(e)
    return {"ok":True}

@app.get("/webhook")
async def verify(request: Request):
    if request.query_params.get("hub.verify_token")=="automatyco123":
        return int(request.query_params.get("hub.challenge"))
    return {"error":"fail"}