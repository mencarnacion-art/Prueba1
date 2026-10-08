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
    d = {"messaging_product": "whatsapp", "to": to, "type": "text", "text": {"body": text}}
    requests.post(url, headers=h, json=d)

def send_doc(to, pptx_bytes, filename="Reporte.pptx"):
    url_media = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media"
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    files = {'file': (filename, pptx_bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'), 'type': (None, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'), 'messaging_product': (None, 'whatsapp')}
    r = requests.post(url_media, headers=h, files=files)
    if r.status_code!=200:
        print(r.text); send_text(to, f"Error: {r.text[:400]}"); return
    media_id = r.json()["id"]
    url_msg = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    hj = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    d = {"messaging_product": "whatsapp", "to": to, "type": "document", "document": {"id": media_id, "filename": filename, "caption": "Listo ✅ con formato"}}
    requests.post(url_msg, headers=hj, json=d)

def parse_data(texts):
    full = "\n".join(texts)
    empresa = "No especificada"
    tecnico = "Técnico"
    # Busca Empresa: X o {Empresa} X
    m = re.search(r'Empresa\s*:\s*(.+)', full, re.IGNORECASE)
    if m: empresa = m.group(1).strip().split("\n")[0]
    # Busca Tecnico / By / Técnico
    m2 = re.search(r'(Tecnico|Técnico|By)\s*:\s*(.+)', full, re.IGNORECASE)
    if m2: tecnico = m2.group(2).strip().split("\n")[0]
    fecha = datetime.now().strftime("%d/%m/%Y")
    return empresa, tecnico, fecha, full

def replace_placeholder(slide, empresa, tecnico, fecha):
    for shape in slide.shapes:
        if not shape.has_text_frame: continue
        txt = shape.text
        if "{Empresa}" in txt or "{EMPRESA}" in txt or "Empresa:" in txt:
            txt = txt.replace("{Empresa}", empresa).replace("{EMPRESA}", empresa).replace("{empresa}", empresa)
        if "{TECNICO}" in txt or "{Tecnico}" in txt or "By:" in txt:
            txt = txt.replace("{TECNICO}", tecnico).replace("{Tecnico}", tecnico).replace("{TECNICO}", tecnico)
            # Si dice By:{TECNICO}, lo reemplazamos
            if "By:{TECNICO}" in shape.text or "By:{Tecnico}" in shape.text or "By:" in shape.text:
                 # limpieza
                 if "{TECNICO}" in shape.text or "{Tecnico}" in shape.text:
                     txt = txt.replace("By:{TECNICO}", f"By: {tecnico}").replace("By:{Tecnico}", f"By: {tecnico}")
        if "{FECHA}" in txt:
            txt = txt.replace("{FECHA}", fecha).replace("{Fecha}", fecha)
        shape.text = txt
        # reemplazo más agresivo por si está en runs
        for p in shape.text_frame.paragraphs:
            for run in p.runs:
                run.text = run.text.replace("{Empresa}", empresa).replace("{EMPRESA}", empresa).replace("{TECNICO}", tecnico).replace("{Tecnico}", tecnico).replace("{FECHA}", fecha)

def create_pptx(texts, images):
    empresa, tecnico, fecha, full_text = parse_data(texts)

    if os.path.exists(TEMPLATE_PATH):
        prs = Presentation(TEMPLATE_PATH)
        # 1. Rellenar portada (slide 1)
        if len(prs.slides) > 0:
            # reemplazo directo en la portada
            for shape in prs.slides[0].shapes:
                if shape.has_text_frame:
                    t = shape.text
                    if "{Empresa}" in t or "Empresa:" in t:
                        shape.text = f"Empresa:\n{empresa}"
                    if "{TECNICO}" in t or "By:" in t:
                        shape.text = f"By: {tecnico}"
                    if "{FECHA}" in t:
                        shape.text = shape.text.replace("{FECHA}", fecha)
                    # reemplazo en runs para respetar formato
                    for p in shape.text_frame.paragraphs:
                        for run in p.runs:
                            run.text = run.text.replace("{Empresa}", empresa).replace("{TECNICO}", tecnico).replace("{FECHA}", fecha)

        # 2. Diapositiva de descripción con el texto que envió
        layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
        slide_desc = prs.slides.add_slide(layout)
        # Si tu plantilla tiene fondo, se conserva
        txBox = slide_desc.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(12), Inches(6))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.text = f"DESCRIPCIÓN DEL SERVICIO - {fecha}\n\n{full_text}"

        # 3. Fotos
        for img_b in images:
            try:
                s = prs.slides.add_slide(layout)
                with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
                    tf.write(img_b); path = tf.name
                s.shapes.add_picture(path, Inches(0.5), Inches(0.5), Inches(12), Inches(6.5))
                os.unlink(path)
            except Exception as e:
                print(e)
    else:
        prs = Presentation()
        prs.slide_width = Inches(13.33); prs.slide_height = Inches(7.5)
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(6)).text_frame.text = full_text

    bio = io.BytesIO(); prs.save(bio); bio.seek(0); return bio.getvalue()

@app.get("/")
def home(): return {"ok": True, "plantilla": os.path.exists(TEMPLATE_PATH)}

@app.get("/webhook")
def verify(hub_verify_token: str = None, hub_challenge: str = None):
    if hub_verify_token == (os.getenv("VERIFY_TOKEN") or "automatyco123"): return int(hub_challenge)
    return "Token invalido"

@app.post("/webhook")
async def webhook(req: Request):
    data = await req.json()
    try:
        value = data["entry"][0]["changes"][0]["value"]
        if "messages" not in value: return "ok"
        msg = value["messages"][0]
        from_num = msg["from"]
        SESSIONS.setdefault(from_num, {"texts": [], "images": []})
        if msg["type"] == "text":
            body = msg["text"]["body"].strip()
            if body.lower() in ["generar reporte", "generar", "reporte"]:
                sess = SESSIONS[from_num]
                pptx = create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx, filename=f"Reporte_{from_num}.pptx")
                SESSIONS[from_num] = {"texts": [], "images": []}
            else:
                SESSIONS[from_num]["texts"].append(body)
                send_text(from_num, "Texto guardado. Sigue mandando info o escribe *generar reporte*")
        elif msg["type"] == "image":
            mid = msg["image"]["id"]
            h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info = requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h).json()
            img_data = requests.get(info["url"], headers=h).content
            SESSIONS[from_num]["images"].append(img_data)
            send_text(from_num, f"Foto {len(SESSIONS[from_num]['images'])} guardada")
    except Exception as e:
        print(e)
    return "ok"
