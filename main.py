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
    try:
        requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
            headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
            json={"messaging_product":"whatsapp","to":to,"type":"text","text":{"body": text[:3500]}})
    except: pass

def send_doc(to, pptx_bytes, filename):
    h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    files = {'file': (filename, pptx_bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
             'type': (None, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
             'messaging_product': (None, 'whatsapp')}
    r = requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media", headers=h, files=files, timeout=30)
    print("MEDIA", r.status_code)
    if r.status_code != 200: print(r.text); return
    mid = r.json()["id"]
    requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
        headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
        json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":mid,"filename":filename}})

def limpiar_markdown(t):
    if not t: return ""
    t = re.sub(r'\*\*(.*?)\*\*', r'\1', t)
    t = re.sub(r'[*_`#]', '', t)
    t = re.sub(r'^---+$', '', t, flags=re.MULTILINE)
    return t.strip()

def mejorar_con_gemini(texto_raw, empresa):
    if not GOOGLE_API_KEY: return None
    modelos = ["gemini-2.5-flash", "gemini-2.5-flash-latest", "gemini-3.8-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
    for modelo in modelos:
        try:
            prompt = f"Redacta profesional en texto plano sin ** ni ##, cliente {empresa}: {texto_raw[:6000]}"
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent?key={GOOGLE_API_KEY}"
            r = requests.post(url, headers={"Content-Type":"application/json"},
                json={"contents":[{"parts":[{"text":prompt}]}], "generationConfig":{"temperature":0.2}}, timeout=25)
            print(f"GEMINI {modelo}: {r.status_code}")
            if r.status_code == 200:
                txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                return limpiar_markdown(txt)
        except Exception as e:
            print(e)
    return None

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
    mejorado = mejorar_con_gemini(desc, empresa)
    final = limpiar_markdown(mejorado) if mejorado else limpiar_markdown(desc)
    return empresa, tecnico, fecha_file, fecha_show, final

def add_slide_before_last(prs, layout):
    ns = prs.slides.add_slide(layout)
    lst = prs.slides._sldIdLst
    nid = lst[-1]; lst.remove(nid); lst.insert(len(lst)-1, nid)
    return ns

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH)
    layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]

    for shape in prs.slides[0].shapes:
        replace_keep_format(shape, "{Empresa}", empresa)
        replace_keep_format(shape, "{EMPRESA}", empresa)
        replace_keep_format(shape, "{TECNICO}", tecnico)
        replace_keep_format(shape, "{Tecnico}", tecnico)
        replace_keep_format(shape, "{FECHA}", fecha_show)

    paginas = [trabajo[i:i+900] for i in range(0, len(trabajo), 900)] if len(trabajo) > 900 else [trabajo]
    
    for shape in prs.slides[1].shapes:
        if shape.has_text_frame and "{TRABAJO}" in shape.text:
            shape.text_frame.word_wrap = True
            replace_keep_format(shape, "{TRABAJO}", paginas[0])

    for extra in paginas[1:]:
        ns = add_slide_before_last(prs, layout)
        txBox = ns.shapes.add_textbox(Inches(0.5), Inches(0.8), Inches(9), Inches(5.5))
        txBox.text_frame.word_wrap = True
        txBox.text_frame.text = extra

    img_idx = 0
    for sld in list(prs.slides)[2:-1]:
        if img_idx >= len(images): break
        if any("{TRABAJO}" in s.text for s in sld.shapes if s.has_text_frame): continue
        for shp in list(sld.shapes):
            if shp.shape_type == 13: shp._element.getparent().remove(shp._element)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path = tf.name
        sld.shapes.add_picture(path, Inches(0.5), Inches(1.0), Inches(9), Inches(5))
        os.unlink(path); img_idx += 1

    while img_idx < len(images):
        ns = add_slide_before_last(prs, layout)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path = tf.name
        ns.shapes.add_picture(path, Inches(0.5), Inches(1.0), Inches(9), Inches(5))
        os.unlink(path); img_idx += 1

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
                send_text(frm, f"Generando {len(sess['texts'])} textos y {len(sess['images'])} fotos...")
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
