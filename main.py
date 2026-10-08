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
    found=False
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            if old in run.text:
                run.text=run.text.replace(old,new)
                found=True
    return found

def create_pptx(texts, images):
    empresa, tecnico, fecha, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH) if os.path.exists(TEMPLATE_PATH) else Presentation()

    # 1. Portada
    for shape in prs.slides[0].shapes:
        replace_keep_format(shape,"{Empresa}",empresa)
        replace_keep_format(shape,"{EMPRESA}",empresa)
        replace_keep_format(shape,"{TECNICO}",tecnico)
        replace_keep_format(shape,"{Tecnico}",tecnico)
        replace_keep_format(shape,"{FECHA}",fecha)

    # 2. Buscar slide con {TRABAJO} y rellenarlo
    trabajo_rellenado=False
    for slide in prs.slides[1:]: # desde slide 2
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                # reemplazo conservando formato
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if "{TRABAJO}" in run.text:
                            run.text=run.text.replace("{TRABAJO}", trabajo)
                            trabajo_rellenado=True
                # si quedó solo "Tabajo realizado:" + trabajo, ajusta
                if not trabajo_rellenado and "{TRABAJO}" in shape.text:
                    shape.text=shape.text.replace("{TRABAJO}", trabajo)
                    trabajo_rellenado=True

    # Si no encontró {TRABAJO}, crea la descripción en slide 2
    if not trabajo_rellenado and len(prs.slides)>1:
        for shape in prs.slides[1].shapes:
            if shape.has_text_frame and "Trabajo realizado" in shape.text:
                # agrega debajo
                shape.text_frame.paragraphs[0].runs[0].text = f"Trabajo realizado:\n{trabajo}"
                trabajo_rellenado=True
                break

    # 3. Fotos -> usar slides vacías de la plantilla primero (slide 3 en adelante)
    # Consideramos que la última slide es cierre y no se toca
    last_idx = len(prs.slides)-1
    img_idx=0
    from pptx.util import Inches
    # primero intenta llenar slides vacías que ya existen
    for i in range(2, len(prs.slides)):
        if i==last_idx: continue # no tocar última
        if img_idx>=len(images): break
        slide=prs.slides[i]
        # si la slide está vacía o solo tiene el header/footer de tu plantilla
        has_pic=len(slide.shapes)>0 and any(s.shape_type==13 for s in slide.shapes) # ya tiene foto
        if has_pic: continue
        # si es la de trabajo ya rellenada, saltar
        if trabajo_rellenado and i==1: continue
        # mete foto
        try:
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
                tf.write(images[img_idx]); path=tf.name
            slide.shapes.add_picture(path, Inches(0.5), Inches(0.9), Inches(12), Inches(5.8))
            os.unlink(path)
            img_idx+=1
        except Exception as e:
            print(e)

    # Si aún quedan fotos, crea nuevas antes de la última
    def add_before_last():
        layout=prs.slide_layouts[6] if len(prs.slide_layouts)>6 else prs.slide_layouts[-1]
        new=prs.slides.add_slide(layout)
        sldIdLst=prs.slides._sldIdLst
        new_id=sldIdLst[-1]
        sldIdLst.remove(new_id)
        sldIdLst.insert(-1, new_id) # antes de la última
        return new

    while img_idx < len(images):
        s=add_before_last()
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path=tf.name
        s.shapes.add_picture(path, Inches(0.5), Inches(0.5), Inches(12), Inches(6.2))
        os.unlink(path)
        img_idx+=1

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
                send_text(from_num, f"Generando reporte final con formato... {len(sess['images'])} fotos")
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
        print(e)
    return "ok"
