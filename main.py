import os, io, re, tempfile, requests
from datetime import datetime
from fastapi import FastAPI, Request
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID") or os.getenv("PHONE_NUMBER_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
SESSIONS = {}
TEMPLATE_PATH = "Reporte.pptx"

def send_text(to, text):
    requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
        headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
        json={"messaging_product":"whatsapp","to":to,"type":"text","text":{"body": text[:3500]}})

def send_doc(to, pptx_bytes, filename):
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    files = {'file': (filename, pptx_bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
             'type': (None, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
             'messaging_product': (None, 'whatsapp')}
    r = requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media", headers=h, files=files, timeout=30)
    if r.status_code!= 200: print(r.text); return
    mid = r.json()["id"]
    requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
        headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
        json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":mid,"filename":filename}})

def limpiar_markdown(t):
    return re.sub(r'[*_`#]', '', t).strip() if t else ""

def mejorar_texto_profesional(texto_raw, empresa):
    if not texto_raw: return texto_raw
    partes = [texto_raw[i:i+3500] for i in range(0, len(texto_raw), 3500)]
    resultado = ""
    for parte in partes:
        corregida = parte
        if GOOGLE_API_KEY:
            try:
                url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key={GOOGLE_API_KEY}"
                payload = {"contents": [{"parts": [{"text": f"Eres redactor técnico de Automatyco. Corrige ortografía y redacción profesional, mantén TODO el detalle sin resumir ni cortar, sin usar ** ni ##. Cliente {empresa}: {parte}"}]}],
                    "generationConfig": {"temperature": 0.2, "maxOutputTokens": 8192}}
                r = requests.post(url, headers={"Content-Type":"application/json"}, json=payload, timeout=40)
                if r.status_code == 200:
                    txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                    if len(txt) >= len(parte)*0.6: corregida = txt
            except Exception as e: print(e)
        resultado += " " + corregida
    return limpiar_markdown(resultado)

def replace_keep_format(shape, placeholder, value):
    if not shape.has_text_frame: return
    if placeholder not in shape.text: return
    for para in shape.text_frame.paragraphs:
        if placeholder in para.text:
            for run in para.runs:
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)
                    return
            para.text = para.text.replace(placeholder, value)
            return

def dividir_inteligente(texto, max_chars=850):
    oraciones = re.split(r'(?<=[.!?\n])\s+', texto)
    paginas, actual = [], ""
    for o in oraciones:
        if len(actual) + len(o) + 1 < max_chars:
            actual = (actual + " " + o).strip() if actual else o
        else:
            if actual: paginas.append(actual.strip())
            actual = o
    if actual: paginas.append(actual.strip())
    return paginas

def delete_slide(prs, index):
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]

def clear_placeholders(slide):
    for shape in list(slide.shapes):
        if shape.is_placeholder:
            sp = shape._element
            sp.getparent().remove(sp)

def parse_data(texts):
    full = "\n".join(texts)
    empresa = "CLIENTE"; tecnico = "TECNICO"
    m = re.search(r'Empresa\s*:\s*(.+)', full, re.IGNORECASE)
    if m: empresa = m.group(1).split("\n")[0].strip()
    m2 = re.search(r'(Tecnico|Técnico|By)\s*:\s*(.+)', full, re.IGNORECASE)
    if m2: tecnico = m2.group(2).split("\n")[0].strip()
    fecha_file = datetime.now().strftime("%Y-%m-%d")
    fecha_show = datetime.now().strftime("%d/%m/%Y")
    desc = re.sub(r'Empresa\s*:.*\n?', '', full, flags=re.IGNORECASE)
    desc = re.sub(r'(Tecnico|Técnico|By)\s*:.*\n?', '', desc, flags=re.IGNORECASE).strip()
    if not desc: desc = full
    final = mejorar_texto_profesional(desc, empresa)
    return empresa, tecnico, fecha_file, fecha_show, final

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH)

    for shape in prs.slides[0].shapes:
        replace_keep_format(shape, "{Empresa}", empresa)
        replace_keep_format(shape, "{EMPRESA}", empresa)
        replace_keep_format(shape, "{TECNICO}", tecnico)
        replace_keep_format(shape, "{Tecnico}", tecnico)
        replace_keep_format(shape, "{FECHA}", fecha_show)

    # Dejamos solo Portada y Gracias
    while len(prs.slides) > 2:
        delete_slide(prs, 1)

    sldLst = prs.slides._sldIdLst
    print(f"BASE: {len(sldLst)} slides [Portada, Gracias]")
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    paginas = dividir_inteligente(trabajo, 850)

    # Agregamos todo al final: quedará [Portada, Gracias, Texto1, Foto1...]
    for pagina in paginas:
        ns = prs.slides.add_slide(blank_layout)
        clear_placeholders(ns)
        txBox = ns.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9), Inches(5.8))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.text = pagina

    for img_bytes in images:
        if len(img_bytes) < 20000: continue
        ns = prs.slides.add_slide(blank_layout)
        clear_placeholders(ns)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(img_bytes); path = tf.name
        ns.shapes.add_picture(path, Inches(0.5), Inches(0.3), Inches(9), Inches(5.8))
        os.unlink(path)

    # REORDENAMIENTO BLINDADO - 100% al final
    sldLst = prs.slides._sldIdLst
    if len(sldLst) >= 3:
        portada_id = sldLst[0]
        gracias_id = sldLst[1]
        resto_ids = list(sldLst[2:]) # Todos los textos y fotos
        nuevo_orden = [portada_id] + resto_ids + [gracias_id]
        # Reasignamos en orden
        for idx in range(len(nuevo_orden)):
            sldLst[idx] = nuevo_orden[idx]
        print(f"REORDENADO: Portada + {len(resto_ids)} slides + Gracias al final")

    print(f"Slides finales: {len(prs.slides)}")
    bio = io.BytesIO(); prs.save(bio); bio.seek(0)
    safe = re.sub(r'[^A-Za-z0-9]+','_',empresa).upper()[:20]
    return bio.getvalue(), f"{safe}_{fecha_file}.pptx"

@app.get("/")
def home(): return {"ok": True}

@app.get("/webhook")
def verify(hub_verify_token: str = None, hub_challenge: str = None):
    if hub_verify_token == VERIFY_TOKEN: return int(hub_challenge)
    return "Token invalido"

@app.post("/webhook")
async def webhook(req: Request):
    data = await req.json()
    try:
        v = data["entry"][0]["changes"][0]["value"]
        if "statuses" in v or "messages" not in v: return "ok"
        msg = v["messages"][0]; frm = msg["from"]
        SESSIONS.setdefault(frm, {"texts":[],"images":[]})
        if msg["type"] == "text":
            body = msg["text"]["body"].strip()
            if body.lower() in ["generar reporte","generar","reporte"]:
                sess = SESSIONS.get(frm, {"texts":[],"images":[]})
                if not sess["texts"] and not sess["images"]:
                    send_text(frm, "No tengo datos"); return "ok"
                send_text(frm, f"Generando...")
                pptx, name = create_pptx(sess["texts"], sess["images"])
                send_doc(frm, pptx, name)
                SESSIONS[frm] = {"texts":[],"images":[]}
            else:
                SESSIONS[frm]["texts"].append(body)
                send_text(frm, "Guardado ✅")
        elif msg["type"] == "image":
            mid = msg["image"]["id"]; h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info = requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h, timeout=15).json()
            img = requests.get(info["url"], headers=h, timeout=20).content
            SESSIONS[frm]["images"].append(img)
            send_text(frm, f"Foto {len(SESSIONS[frm]['images'])} guardada")
    except Exception as e:
        print(e); import traceback; traceback.print_exc()
    return "ok"
