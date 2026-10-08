import os, io, re, tempfile, requests
from datetime import datetime
from fastapi import FastAPI, Request
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID") or os.getenv("PHONE_NUMBER_ID") or os.getenv("PHONE_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
# Lee las dos como tienes en la foto
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY") or os.getenv("OPENAI_API")
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
    print("UPLOAD MEDIA:", r.status_code, r.text[:500])
    if r.status_code!=200: send_text(to, f"Error subiendo: {r.text[:300]}"); return
    media_id=r.json()["id"]
    url_msg=f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    hj={"Authorization": f"Bearer {WHATSAPP_TOKEN}","Content-Type":"application/json"}
    requests.post(url_msg, headers=hj, json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":media_id,"filename":filename,"caption":f"Reporte {filename}"}})

def mejorar_profesional(texto_original):
    print(f"OPENAI KEY present: {bool(OPENAI_API_KEY)} len={len(OPENAI_API_KEY) if OPENAI_API_KEY else 0}")
    if not OPENAI_API_KEY:
        return texto_original

    prompt = f"""Eres redactor técnico senior de Automatyco. Corrige ortografía, gramática y redacta de forma 100% profesional este reporte de campo. No agregues datos falsos. Usa lenguaje técnico formal.

Texto del técnico:
{texto_original}

Devuelve SOLO el reporte mejorado en este formato:
SERVICIO REALIZADO:
- [bullet profesional]
- [bullet profesional]

Deja el equipo operativo. Recomendaciones si aplica.
Todo en español profesional."""

    try:
        r = requests.post("https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": "gpt-4o-mini",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.2,
                "max_tokens": 600
            }, timeout=30)
        print("OPENAI RESPONSE:", r.status_code, r.text[:1000])
        if r.status_code == 200:
            mejorado = r.json()["choices"][0]["message"]["content"].strip()
            if len(mejorado) > 20:
                return mejorado
        else:
            print(f"OPENAI ERROR {r.status_code}: {r.text}")
    except Exception as e:
        print(f"OPENAI EXCEPTION: {e}")
        import traceback; traceback.print_exc()

    return texto_original # si falla, devuelve original pero ya viste log

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
    if not desc: desc=full

    desc_mejorado = mejorar_profesional(desc)
    return empresa, tecnico, fecha_file, fecha_show, desc_mejorado

def replace_keep_format(shape, old, new):
    if not shape.has_text_frame: return False
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            if old in run.text:
                run.text=run.text.replace(old,new)
                return True
    # si no hay runs con el token, reemplaza todo el texto del shape (caso placeholder único)
    if shape.has_text_frame and old in shape.text:
        shape.text = shape.text.replace(old, new)
        return True
    return False

def add_slide_before_last(prs, layout):
    new_slide = prs.slides.add_slide(layout)
    sldIdLst = prs.slides._sldIdLst
    new_id = sldIdLst[-1]
    sldIdLst.remove(new_id)
    sldIdLst.insert(len(sldIdLst)-1, new_id) # antes del cierre
    return new_slide

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH)

    # 1. Portada
    for shape in prs.slides[0].shapes:
        replace_keep_format(shape,"{Empresa}",empresa)
        replace_keep_format(shape,"{EMPRESA}",empresa)
        replace_keep_format(shape,"{TECNICO}",tecnico)
        replace_keep_format(shape,"{Tecnico}",tecnico)
        replace_keep_format(shape,"{FECHA}",fecha_show)

    # 2. Trabajo - busca {TRABAJO} en todas menos la última
    for slide in list(prs.slides)[1:-1]:
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                # reemplazo conservando formato
                if not replace_keep_format(shape,"{TRABAJO}",trabajo):
                    # fallback: si el shape solo tiene el token
                    shape.text = trabajo

    # 3. Fotos - PROTEGE ULTIMA
    layout = prs.slide_layouts[6]
    img_idx = 0

    # a) Reusa slides vacías del medio (índice 2 hasta penúltima)
    for i in range(2, len(list(prs.slides))-1):
        if img_idx >= len(images): break
        slide = prs.slides[i]
        if any(s.shape_type==13 for s in slide.shapes): continue # ya tiene foto
        if any("{TRABAJO}" in (s.text if s.has_text_frame else "") for s in slide.shapes): continue
        try:
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
                tf.write(images[img_idx]); path=tf.name
            # FOTO DENTRO DE MARGEN: Ancho 8.5", centrado, alto automático
            slide.shapes.add_picture(path, Inches(2.4), Inches(1.5), Inches(8.5))
            os.unlink(path)
            img_idx+=1
        except Exception as e:
            print(e)

    # b) Resto de fotos en slides NUEVAS antes del cierre
    while img_idx < len(images):
        s = add_slide_before_last(prs, layout)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path=tf.name
        s.shapes.add_picture(path, Inches(2.4), Inches(1.5), Inches(8.5))
        os.unlink(path)
        img_idx+=1

    bio=io.BytesIO(); prs.save(bio); bio.seek(0)
    safe_empresa = re.sub(r'[^A-Za-z0-9_-]+','_',empresa).upper()[:20]
    safe_tecnico = re.sub(r'[^A-Za-z0-9_-]+','_',tecnico).upper()[:15]
    filename = f"{safe_empresa}_{fecha_file}_{safe_tecnico}.pptx"
    return bio.getvalue(), filename

@app.get("/")
def home(): return {"ok":True}

@app.get("/webhook")
def verify(hub_verify_token: str = None, hub_challenge: str = None):
    if hub_verify_token == (os.getenv("VERIFY_TOKEN") or "automatyco123"): return int(hub_challenge)
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
                send_text(from_num, f"Corrigiendo texto con IA y armando {len(sess['images'])} fotos...")
                pptx_bytes, filename = create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx_bytes, filename=filename)
                SESSIONS[from_num]={"texts":[],"images":[]}
            else:
                SESSIONS[from_num]["texts"].append(body)
                send_text(from_num, "Guardado ✅")
        elif msg["type"]=="image":
            mid=msg["image"]["id"]
            h={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info=requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h).json()
            img=requests.get(info["url"], headers=h).content
            SESSIONS[from_num]["images"].append(img)
            send_text(from_num, f"Foto {len(SESSIONS[from_num]['images'])} guardada")
    except Exception as e:
        print(e); import traceback; traceback.print_exc()
    return "ok"
