import os, io, re, tempfile, requests
from datetime import datetime
from fastapi import FastAPI, Request
from pptx import Presentation

app = FastAPI()
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID") or os.getenv("PHONE_NUMBER_ID") or os.getenv("PHONE_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
SESSIONS = {}
TEMPLATE_PATH = "Reporte.pptx"

def send_text(to, text):
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    d = {"messaging_product": "whatsapp", "to": to, "type": "text", "text": {"body": text}}
    requests.post(url, headers=h, json=d)

def send_doc(to, pptx_bytes, filename="Reporte.pptx"):
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media"
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    files = {'file': (filename, pptx_bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'), 'type': (None, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'), 'messaging_product': (None, 'whatsapp')}
    r = requests.post(url, headers=h, files=files)
    if r.status_code!=200: print(r.text); return
    media_id = r.json()["id"]
    url_msg = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    hj = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    d = {"messaging_product": "whatsapp", "to": to, "type": "document", "document": {"id": media_id, "filename": filename, "caption": "Reporte con formato"}}
    requests.post(url_msg, headers=hj, json=d)

def parse_data(texts):
    full = "\n".join(texts)
    empresa = "No especificada"
    tecnico = "Técnico"
    m = re.search(r'Empresa\s*:\s*(.+)', full, re.IGNORECASE)
    if m: empresa = m.group(1).split("\n")[0].strip()
    m2 = re.search(r'(Tecnico|Técnico|By)\s*:\s*(.+)', full, re.IGNORECASE)
    if m2: tecnico = m2.group(2).split("\n")[0].strip()
    fecha = datetime.now().strftime("%d/%m/%Y")
    # texto para descripción (quitamos las líneas de empresa/tecnico para no repetir)
    desc = re.sub(r'Empresa\s*:.*\n?', '', full, flags=re.IGNORECASE)
    desc = re.sub(r'(Tecnico|Técnico|By)\s*:.*\n?', '', desc, flags=re.IGNORECASE).strip()
    return empresa, tecnico, fecha, desc, full

def replace_keep_format(shape, replacements):
    """Reemplaza solo dentro de los runs, no toca fuente/tamaño"""
    if not shape.has_text_frame: return
    for p in shape.text_frame.paragraphs:
        for run in p.runs:
            for old, new in replacements.items():
                if old in run.text:
                    run.text = run.text.replace(old, new)

def create_pptx(texts, images):
    empresa, tecnico, fecha, desc, full = parse_data(texts)
    replacements = {"{Empresa}": empresa, "{EMPRESA}": empresa, "{empresa}": empresa,
                    "{TECNICO}": tecnico, "{Tecnico}": tecnico, "{tecnico}": tecnico,
                    "{FECHA}": fecha, "{Fecha}": fecha, "{fecha}": fecha}

    if not os.path.exists(TEMPLATE_PATH):
        # fallback
        prs = Presentation(); s=prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_textbox(100000,100000,1000000,1000000).text_frame.text=full
        bio=io.BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()

    prs = Presentation(TEMPLATE_PATH)

    # 1. RELLENAR PORTADA (slide 0) sin borrar formato
    if len(prs.slides)>0:
        for shape in prs.slides[0].shapes:
            replace_keep_format(shape, replacements)

    # 2. Guardar última diapositiva si existe (para que se quede al final)
    # La vamos a clonar al final después de insertar todo
    last_slide = prs.slides[-1] if len(prs.slides)>1 else None
    # Guardamos el ID de la última slide para mover nuevas antes de ella
    # Truco: insertamos slides nuevas en penúltimo lugar

    def add_slide_before_last(layout):
        new_slide = prs.slides.add_slide(layout)
        if last_slide and len(prs.slides)>1:
            # mover la nueva slide antes de la última
            sldIdLst = prs.slides._sldIdLst
            # la nueva es la última, la penúltima es la que era última
            new_sldId = sldIdLst[-1]
            last_sldId = sldIdLst[-2]
            # insertar new antes de last
            sldIdLst.remove(new_sldId)
            # insertar en -1 posición (antes del último)
            sldIdLst.insert(-1, new_sldId)
        return new_slide

    layout = prs.slide_layouts[6] if len(prs.slide_layouts)>6 else prs.slide_layouts[1]

    # 3. Slide de descripción
    if desc:
        s = add_slide_before_last(layout)
        # usa el mismo fondo de la plantilla, solo agrega textbox con formato neutro
        from pptx.util import Inches
        tx = s.shapes.add_textbox(Inches(0.5), Inches(0.8), Inches(12.3), Inches(6))
        tf = tx.text_frame; tf.word_wrap=True
        tf.text = f"Descripción del servicio:\n\n{desc}"

    # 4. Fotos - una por slide, antes de la última
    from pptx.util import Inches
    for img_b in images:
        try:
            s = add_slide_before_last(layout)
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
                tf.write(img_b); path=tf.name
            s.shapes.add_picture(path, Inches(0.5), Inches(0.5), Inches(12), Inches(6.2))
            os.unlink(path)
        except Exception as e:
            print(e)

    bio=io.BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()

@app.get("/")
def home(): return {"plantilla": os.path.exists(TEMPLATE_PATH)}

@app.get("/webhook")
def verify(hub_verify_token: str = None, hub_challenge: str = None):
    if hub_verify_token == (os.getenv("VERIFY_TOKEN") or "automatyco123"): return int(hub_challenge)
    return "Token invalido"

@app.post("/webhook")
async def webhook(req: Request):
    data = await req.json()
    try:
        v = data["entry"][0]["changes"][0]["value"]
        if "messages" not in v: return "ok"
        msg = v["messages"][0]; from_num=msg["from"]
        SESSIONS.setdefault(from_num, {"texts":[],"images":[]})
        if msg["type"]=="text":
            body=msg["text"]["body"].strip()
            if body.lower() in ["generar reporte","generar","reporte"]:
                sess=SESSIONS[from_num]
                send_text(from_num, f"Generando con tu formato... {len(sess['images'])} fotos")
                pptx=create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx, filename=f"Reporte_{from_num}.pptx")
                SESSIONS[from_num]={"texts":[],"images":[]}
            else:
                SESSIONS[from_num]["texts"].append(body)
                send_text(from_num, "Guardado. Escribe *generar reporte* al terminar")
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
