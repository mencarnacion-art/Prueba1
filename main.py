import os, io, re, tempfile, requests, time
from datetime import datetime
from fastapi import FastAPI, Request
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID") or os.getenv("PHONE_NUMBER_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
SESSIONS = {}; TEMPLATE_PATH = "Reporte.pptx"

def send_text(to, text):
    try:
        requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
            headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
            json={"messaging_product":"whatsapp","to":to,"type":"text","text":{"body":text[:4000]}})
    except: pass

def send_doc(to, pptx_bytes, filename):
    try:
        h={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
        files={'file':(filename,pptx_bytes,'application/vnd.openxmlformats-officedocument.presentationml.presentation'),'type':(None,'application/vnd.openxmlformats-officedocument.presentationml.presentation'),'messaging_product':(None,'whatsapp')}
        r=requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media",headers=h,files=files, timeout=20)
        if r.status_code!=200: print("MEDIA ERROR", r.text); return
        mid=r.json()["id"]
        requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
            headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}","Content-Type":"application/json"},
            json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":mid,"filename":filename}}, timeout=20)
    except Exception as e: print(f"SEND DOC ERROR {e}")

def mejorar_con_gemini(texto_raw, empresa):
    if not GOOGLE_API_KEY or not texto_raw: return None
    models = ["gemini-2.0-flash", "gemini-2.0-flash-001", "gemini-1.5-flash", "gemini-1.5-flash-latest", "gemini-3.8-flash"]
    for model in models:
        for attempt in range(2):
            try:
                prompt = f"Eres redactor técnico senior de Automatyco. Mejora a profesional este reporte de {empresa}. Texto técnico: {texto_raw[:8000]}. Devuelve SOLO texto mejorado en español, sin saludos."
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GOOGLE_API_KEY}"
                r = requests.post(url, headers={"Content-Type":"application/json"},
                    json={"contents":[{"parts":[{"text":prompt}]}], "generationConfig":{"temperature":0.2, "maxOutputTokens":2048}}, timeout=25)
                print(f"GEMINI {model}: {r.status_code}")
                if r.status_code==200:
                    txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                    if txt: return txt
                if r.status_code==503: # sobrecarga, espera
                    time.sleep(1)
                    continue
                if r.status_code==404: # no existe, pasa al siguiente modelo
                    break
                print(r.text[:1000])
            except Exception as e:
                print(f"GEMINI ERROR {model}: {e}")
                time.sleep(1)
    return None

def parse_data(texts):
    full="\n".join(texts); empresa="CLIENTE"; tecnico="TECNICO"
    m=re.search(r'Empresa\s*:\s*(.+)', full, re.IGNORECASE)
    if m: empresa=m.group(1).split("\n")[0].strip()[:40]
    m2=re.search(r'(Tecnico|Técnico|By)\s*:\s*(.+)', full, re.IGNORECASE)
    if m2: tecnico=m2.group(2).split("\n")[0].strip()[:30]
    fecha_file=datetime.now().strftime("%Y-%m-%d"); fecha_show=datetime.now().strftime("%d/%m/%Y")
    desc=re.sub(r'Empresa\s*:.*\n?','',full,flags=re.IGNORECASE); desc=re.sub(r'(Tecnico|Técnico|By)\s*:.*\n?','',desc,flags=re.IGNORECASE).strip()
    if not desc: desc=full
    mejorado = mejorar_con_gemini(desc, empresa)
    return empresa, tecnico, fecha_file, fecha_show, mejorado or desc

def add_slide_before_last(prs, layout):
    ns=prs.slides.add_slide(layout); lst=prs.slides._sldIdLst; nid=lst[-1]; lst.remove(nid); lst.insert(len(lst)-1,nid); return ns

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH); layout = prs.slide_layouts[1] if len(prs.slide_layouts)>1 else prs.slide_layouts[0]
    for shape in prs.slides[0].shapes:
        if shape.has_text_frame:
            t=shape.text
            if "{Empresa}" in t: shape.text=t.replace("{Empresa}",empresa)
            if "{TECNICO}" in t or "{Tecnico}" in t: shape.text=shape.text.replace("{TECNICO}",tecnico).replace("{Tecnico}",tecnico)
            if "{FECHA}" in t: shape.text=t.replace("{FECHA}",fecha_show)
    for slide in list(prs.slides)[1:-1]:
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                shape.text=shape.text.replace("{TRABAJO}", trabajo)
    img_idx=0
    for i in range(2, len(list(prs.slides))-1):
        if img_idx>=len(images): break
        slide=prs.slides[i]
        if any(s.shape_type==13 for s in slide.shapes): continue
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf: tf.write(images[img_idx]); path=tf.name
        slide.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3)); os.unlink(path); img_idx+=1
    while img_idx < len(images):
        s=add_slide_before_last(prs, layout)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf: tf.write(images[img_idx]); path=tf.name
        s.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3)); os.unlink(path); img_idx+=1
    bio=io.BytesIO(); prs.save(bio); bio.seek(0)
    safe_empresa=re.sub(r'[^A-Za-z0-9]+','_',empresa).upper()[:20]; safe_tecnico=re.sub(r'[^A-Za-z0-9]+','_',tecnico).upper()[:15]
    return bio.getvalue(), f"{safe_empresa}_{fecha_file}_{safe_tecnico}.pptx"

@app.get("/")
def home(): return {"ok":True}
@app.get("/webhook")
def verify(hub_verify_token: str = None, hub_challenge: str = None):
    if hub_verify_token==VERIFY_TOKEN: return int(hub_challenge)
    return "Token invalido"

@app.post("/webhook")
async def webhook(req: Request):
    data=await req.json()
    print(f"WEBHOOK: {str(data)[:800]}")
    try:
        entry=data.get("entry",[])[0]; change=entry.get("changes",[])[0]; value=change.get("value",{})
        # IGNORA statuses (evita el bucle infinito)
        if "statuses" in value: return "ok"
        if "messages" not in value: return "ok"

        msg=value["messages"][0]; from_num=msg["from"]
        # Evita que se responda a si mismo
        if from_num==WHATSAPP_PHONE_ID: return "ok"

        SESSIONS.setdefault(from_num, {"texts":[],"images":[]})

        if msg["type"]=="text":
            body=msg["text"]["body"].strip()
            if body.lower() in ["generar reporte","generar","reporte"]:
                sess=SESSIONS.get(from_num, {"texts":[],"images":[]})
                if not sess["texts"] and not sess["images"]:
                    send_text(from_num, "No tengo datos. Manda Empresa:, Técnico: y descripción primero.")
                    return "ok"
                send_text(from_num, f"Armando reporte con {len(sess['texts'])} textos y {len(sess['images'])} fotos... (si Gemini 3.8 está caído uso 2.0)")
                pptx_bytes, filename = create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx_bytes, filename=filename)
                SESSIONS[from_num]={"texts":[],"images":[]} # Limpia DESPUÉS de enviar
            else:
                SESSIONS[from_num]["texts"].append(body)
                send_text(from_num,"Guardado ✅ Cuando termines escribe: generar reporte")

        elif msg["type"]=="image":
            mid=msg["image"]["id"]; h={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info=requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h, timeout=15).json()
            img=requests.get(info["url"], headers=h, timeout=20).content
            SESSIONS[from_num]["images"].append(img)
            send_text(from_num,f"Foto {len(SESSIONS[from_num]['images'])} guardada")
    except Exception as e:
        print(e); import traceback; traceback.print_exc()
    return "ok"
