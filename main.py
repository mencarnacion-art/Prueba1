import os, io, re, tempfile, requests
from datetime import datetime
from fastapi import FastAPI, Request
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID") or os.getenv("PHONE_NUMBER_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
SESSIONS = {}
TEMPLATE_PATH = "Reporte.pptx"

def send_text(to, text):
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    requests.post(url, headers=h, json={"messaging_product":"whatsapp","to":to,"type":"text","text":{"body":text}})

def send_doc(to, pptx_bytes, filename):
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media"
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    files = {'file': (filename, pptx_bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),'type': (None, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),'messaging_product': (None, 'whatsapp')}
    r = requests.post(url, headers=h, files=files)
    if r.status_code!=200: print(r.text); return
    media_id=r.json()["id"]
    requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
                  headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}","Content-Type":"application/json"},
                  json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":media_id,"filename":filename,"caption":filename}})

def mejorar_con_gemini(texto_raw, empresa):
    if not GOOGLE_API_KEY: return None
    try:
        prompt = f"""Eres redactor técnico senior de Automatyco. Corrige ortografía, gramática y redacta 100% profesional este reporte del cliente {empresa}. No te limites a una sola estructura, adapta según el texto:
- Si es instalación con problemas: Resumen Ejecutivo, Bitácora de Eventos con horas e incidencias, Plan de Acción.
- Si es mantenimiento: Objetivo, Trabajo Realizado, Hallazgos, Conclusión.
Usa lenguaje formal, técnico.

Texto del técnico:
{texto_raw}

Devuelve SOLO el texto mejorado en español, listo para PPT."""
        url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={GOOGLE_API_KEY}"
        r = requests.post(url, headers={"Content-Type":"application/json"},
                          json={"contents":[{"parts":[{"text":prompt}]}]}, timeout=30)
        print("GEMINI:", r.status_code)
        if r.status_code==200:
            return r.json()["candidates"][0]["content"]["parts"][0]["text"]
        print(r.text[:1000])
    except Exception as e:
        print(f"GEMINI ERROR: {e}")
    return None

def mejorar_con_openai(texto_raw, empresa):
    if not OPENAI_API_KEY: return None
    try:
        r = requests.post("https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
            json={"model":"gpt-4o-mini","messages":[{"role":"user","content":f"Redacta profesional reporte de {empresa}: {texto_raw}"}],"temperature":0.2}, timeout=20)
        if r.status_code==200: return r.json()["choices"][0]["message"]["content"]
    except: pass
    return None

def corregir(texto_raw, empresa):
    return mejorar_con_openai(texto_raw, empresa) or mejorar_con_gemini(texto_raw, empresa) or texto_raw

def parse_data(texts):
    full="\n".join(texts)
    empresa="CLIENTE"; tecnico="TECNICO"
    m=re.search(r'Empresa\s*:\s*(.+)', full, re.IGNORECASE)
    if m: empresa=m.group(1).split("\n")[0].strip()
    m2=re.search(r'(Tecnico|Técnico|By)\s*:\s*(.+)', full, re.IGNORECASE)
    if m2: tecnico=m2.group(2).split("\n")[0].strip()
    fecha_file=datetime.now().strftime("%Y-%m-%d")
    fecha_show=datetime.now().strftime("%d/%m/%Y")
    desc=re.sub(r'Empresa\s*:.*\n?','',full,flags=re.IGNORECASE)
    desc=re.sub(r'(Tecnico|Técnico|By)\s*:.*\n?','',desc,flags=re.IGNORECASE).strip()
    return empresa, tecnico, fecha_file, fecha_show, corregir(desc, empresa)

def add_slide_before_last(prs, layout):
    new_slide = prs.slides.add_slide(layout)
    sldIdLst = prs.slides._sldIdLst
    new_id = sldIdLst[-1]; sldIdLst.remove(new_id); sldIdLst.insert(len(sldIdLst)-1, new_id)
    return new_slide

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH)
    layout = prs.slide_layouts[1] if len(prs.slide_layouts)>1 else prs.slide_layouts[0]

    for shape in prs.slides[0].shapes:
        if shape.has_text_frame:
            if "{Empresa}" in shape.text: shape.text=shape.text.replace("{Empresa}",empresa)
            if "{EMPRESA}" in shape.text: shape.text=shape.text.replace("{EMPRESA}",empresa)
            if "{TECNICO}" in shape.text or "{Tecnico}" in shape.text:
                shape.text=shape.text.replace("{TECNICO}",tecnico).replace("{Tecnico}",tecnico)
            if "{FECHA}" in shape.text: shape.text=shape.text.replace("{FECHA}",fecha_show)

    for slide in list(prs.slides)[1:-1]:
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                shape.text = shape.text.replace("{TRABAJO}", trabajo)

    img_idx=0
    for i in range(2, len(list(prs.slides))-1):
        if img_idx>=len(images): break
        slide=prs.slides[i]
        if any(s.shape_type==13 for s in slide.shapes): continue
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path=tf.name
        slide.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3))
        os.unlink(path); img_idx+=1

    while img_idx < len(images):
        s = add_slide_before_last(prs, layout)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path=tf.name
        s.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3))
        os.unlink(path); img_idx+=1

    bio=io.BytesIO(); prs.save(bio); bio.seek(0)
    safe_empresa=re.sub(r'[^A-Za-z0-9]+','_',empresa).upper()[:20]
    safe_tecnico=re.sub(r'[^A-Za-z0-9]+','_',tecnico).upper()[:15]
    return bio.getvalue(), f"{safe_empresa}_{fecha_file}_{safe_tecnico}.pptx"

@app.get("/")
def home(): return {"ok":True}
@app.get("/webhook")
def verify(hub_verify_token: str = None, hub_challenge: str = None):
    if hub_verify_token==(os.getenv("VERIFY_TOKEN") or "automatyco123"): return int(hub_challenge)
    return "Token invalido"
@app.post("/webhook")
async def webhook(req: Request):
    data=await req.json()
    try:
        v=data["entry"][0]["changes"][0]["value"]
        if "messages" not in v: return "ok"
        msg=v["messages"][0]; from_num=msg["from"]
        SESSIONS.setdefault(from_num, {"texts":[],"images":[]})
        if msg["type"]=="text":
            body=msg["text"]["body"].strip()
            if body.lower() in ["generar reporte","generar","reporte"]:
                sess=SESSIONS[from_num]
                send_text(from_num, f"Consultando Gemini y armando {len(sess['images'])} fotos...")
                pptx_bytes, filename = create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx_bytes, filename=filename)
                SESSIONS[from_num]={"texts":[],"images":[]}
            else:
                SESSIONS[from_num]["texts"].append(body); send_text(from_num,"Guardado ✅")
        elif msg["type"]=="image":
            mid=msg["image"]["id"]; h={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info=requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h).json()
            img=requests.get(info["url"], headers=h).content
            SESSIONS[from_num]["images"].append(img); send_text(from_num,f"Foto {len(SESSIONS[from_num]['images'])} guardada")
    except Exception as e: print(e); import traceback; traceback.print_exc()
    return "ok"
