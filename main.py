import os, io, re, tempfile, requests
from datetime import datetime
from fastapi import FastAPI, Request
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID") or os.getenv("PHONE_NUMBER_ID") or os.getenv("PHONE_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
SESSIONS = {}
TEMPLATE_PATH = "Reporte.pptx"

def send_text(to, text):
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    requests.post(url, headers=h, json={"messaging_product":"whatsapp","to":to,"type":"text","text":{"body":text}})

def send_doc(to, pptx_bytes, filename="Reporte.pptx"):
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media"
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    files = {'file': (filename, pptx_bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),'type': (None, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),'messaging_product': (None, 'whatsapp')}
    r = requests.post(url, headers=h, files=files)
    if r.status_code!=200: print(r.text); return
    media_id=r.json()["id"]
    url_msg=f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    hj={"Authorization": f"Bearer {WHATSAPP_TOKEN}","Content-Type":"application/json"}
    requests.post(url_msg, headers=hj, json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":media_id,"filename":filename,"caption":"Reporte con formato"}})

def parse_data(texts):
    full="\n".join(texts)
    empresa=""; tecnico="";
    m=re.search(r'Empresa\s*:\s*(.+)', full, re.IGNORECASE)
    if m: empresa=m.group(1).split("\n")[0].strip()
    m2=re.search(r'(Tecnico|Técnico|By)\s*:\s*(.+)', full, re.IGNORECASE)
    if m2: tecnico=m2.group(2).split("\n")[0].strip()
    fecha=datetime.now().strftime("%d/%m/%Y")
    desc=re.sub(r'Empresa\s*:.*\n?','',full,flags=re.IGNORECASE)
    desc=re.sub(r'(Tecnico|Técnico|By)\s*:.*\n?','',desc,flags=re.IGNORECASE).strip()
    if not desc: desc=full
    return empresa, tecnico, fecha, desc

def replace_keep_format(shape, old, new):
    if not shape.has_text_frame: return False
    changed=False
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            if old in run.text:
                run.text=run.text.replace(old,new)
                changed=True
    return changed

def create_pptx(texts, images):
    empresa, tecnico, fecha, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH) if os.path.exists(TEMPLATE_PATH) else Presentation()

    # 1. Portada - sin romper formato
    first_slide = prs.slides[0]
    for shape in first_slide.shapes:
        replace_keep_format(shape,"{Empresa}",empresa)
        replace_keep_format(shape,"{EMPRESA}",empresa)
        replace_keep_format(shape,"{TECNICO}",tecnico)
        replace_keep_format(shape,"{Tecnico}",tecnico)
        replace_keep_format(shape,"{FECHA}",fecha)

    # 2. Rellenar {TRABAJO} - CORREGIDO sin slice
    all_slides = list(prs.slides)
    for slide in all_slides[1:]: # ahora si es lista Python normal
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if "{TRABAJO}" in run.text:
                            run.text=run.text.replace("{TRABAJO}", trabajo)

    # 3. Fotos - SIN mover _sldIdLst para no romper, solo usamos slides vacías y agregamos al final
    # Guardamos cuantas slides tenía originalmente
    original_count = len(prs.slides)
    img_idx = 0

    # Usa slides 2,3,etc que estén vacías (tu plantilla tiene 2 slides vacías)
    for idx in range(2, original_count):
        if img_idx >= len(images): break
        slide = prs.slides[idx]
        # si ya es la de trabajo, skip
        if any("{TRABAJO}" in s.text for s in slide.shapes if s.has_text_frame):
            continue
        # si no tiene ya foto, ponle
        try:
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
                tf.write(images[img_idx]); path=tf.name
            slide.shapes.add_picture(path, Inches(0.5), Inches(0.9), Inches(12), Inches(5.8))
            os.unlink(path)
            img_idx+=1
        except Exception as e:
            print(e)

    # Resto de fotos como slides nuevas al final
    layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    while img_idx < len(images):
        s = prs.slides.add_slide(layout)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path=tf.name
        s.shapes.add_picture(path, Inches(0.5), Inches(0.5), Inches(12), Inches(6.2))
        os.unlink(path)
        img_idx+=1

    # 4. Mover la ultima original al final si se agregaron nuevas (forma segura)
    if len(prs.slides) > original_count:
        sldIdLst = prs.slides._sldIdLst
        # la original ultima era la del index original_count-1, ahora está en medio
        # la movemos al final de verdad
        original_last = sldIdLst[original_count-1]
        sldIdLst.remove(original_last)
        sldIdLst.append(original_last)

    bio=io.BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()

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
                send_text(from_num, f"Generando... {len(sess['images'])} fotos")
                pptx=create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx, filename=f"Reporte_{from_num}.pptx")
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
        print(f"Error: {e}")
        import traceback; traceback.print_exc()
    return "ok"
