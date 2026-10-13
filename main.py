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
    except Exception as e:
        print(f"SEND TEXT ERROR {e}")

def send_doc(to, pptx_bytes, filename):
    try:
        h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
        files = {'file': (filename, pptx_bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
                 'type': (None, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
                 'messaging_product': (None, 'whatsapp')}
        r = requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media", headers=h, files=files, timeout=30)
        if r.status_code != 200:
            print("MEDIA ERROR", r.text)
            return
        mid = r.json()["id"]
        requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
            headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
            json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":mid,"filename":filename}})
    except Exception as e:
        print(f"SEND DOC ERROR {e}")

def limpiar_markdown(texto):
    texto = re.sub(r'\*\*(.*?)\*\*', r'\1', texto)
    texto = re.sub(r'__|\*\*|\*|`', '', texto)
    texto = re.sub(r'^#{1,6}\s*', '', texto, flags=re.MULTILINE)
    texto = re.sub(r'^---+\s*$', '', texto, flags=re.MULTILINE)
    texto = re.sub(r'^\s*[-]\s+', '• ', texto, flags=re.MULTILINE)
    return texto.strip()

def mejorar_con_gemini(texto_raw, empresa):
    if not GOOGLE_API_KEY or not texto_raw:
        return None
    models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-flash-latest"]
    for model in models:
        try:
            prompt = f"Eres redactor técnico de Automatyco. Mejora este reporte del cliente {empresa}. NO uses markdown, NO uses **, NO uses ###, NO uses ---. Texto plano profesional en español. Texto: {texto_raw[:7000]}"
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GOOGLE_API_KEY}"
            r = requests.post(url, headers={"Content-Type":"application/json"},
                json={"contents":[{"parts":[{"text":prompt}]}], "generationConfig":{"temperature":0.2, "maxOutputTokens":3000}}, timeout=30)
            print(f"GEMINI {model}: {r.status_code}")
            if r.status_code == 200:
                txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                return limpiar_markdown(txt)
        except Exception as e:
            print(f"GEMINI ERROR {model}: {e}")
    return None

def replace_keep_format(shape, placeholder, value):
    if not shape.has_text_frame: return
    if placeholder not in shape.text: return
    for paragraph in shape.text_frame.paragraphs:
        if placeholder in paragraph.text:
            paragraph.text = paragraph.text.replace(placeholder, value)

def parse_data(texts):
    full = "\n".join(texts)
    empresa = "CLIENTE"; tecnico = "TECNICO"
    m = re.search(r'Empresa\s*:\s*(.+)', full, re.IGNORECASE)
    if m: empresa = m.group(1).split("\n")[0].strip()[:40]
    m2 = re.search(r'(Tecnico|Técnico|By)\s*:\s*(.+)', full, re.IGNORECASE)
    if m2: tecnico = m2.group(2).split("\n")[0].strip()[:30]
    fecha_file = datetime.now().strftime("%Y-%m-%d")
    fecha_show = datetime.now().strftime("%d/%m/%Y")
    desc = re.sub(r'Empresa\s*:.*\n?', '', full, flags=re.IGNORECASE)
    desc = re.sub(r'(Tecnico|Técnico|By)\s*:.*\n?', '', desc, flags=re.IGNORECASE).strip()
    if not desc: desc = full
    mejorado = mejorar_con_gemini(desc, empresa)
    final = limpiar_markdown(mejorado or desc)
    return empresa, tecnico, fecha_file, fecha_show, final

def add_slide_before_last(prs, layout):
    new_slide = prs.slides.add_slide(layout)
    sldIdLst = prs.slides._sldIdLst
    new_id = sldIdLst[-1]
    sldIdLst.remove(new_id)
    sldIdLst.insert(len(sldIdLst)-1, new_id)
    return new_slide

def split_texto_en_paginas(texto, max_chars=850):
    parrafos = texto.split('\n')
    paginas = []; actual = ""
    for p in parrafos:
        if len(actual) + len(p) + 1 < max_chars:
            actual += p + "\n"
        else:
            if actual: paginas.append(actual.strip())
            if len(p) > max_chars:
                for i in range(0, len(p), max_chars):
                    paginas.append(p[i:i+max_chars])
                actual = ""
            else:
                actual = p + "\n"
    if actual: paginas.append(actual.strip())
    return paginas if paginas else [texto]

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH)
    layout_contenido = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]

    for shape in prs.slides[0].shapes:
        replace_keep_format(shape, "{Empresa}", empresa)
        replace_keep_format(shape, "{EMPRESA}", empresa)
        replace_keep_format(shape, "{TECNICO}", tecnico)
        replace_keep_format(shape, "{Tecnico}", tecnico)
        replace_keep_format(shape, "{FECHA}", fecha_show)

    paginas_texto = split_texto_en_paginas(trabajo, max_chars=850)

    slide_trabajo_idx = None
    for idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                slide_trabajo_idx = idx
                break
        if slide_trabajo_idx is not None: break

    if slide_trabajo_idx is not None:
        for shape in prs.slides[slide_trabajo_idx].shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                shape.text_frame.word_wrap = True
                replace_keep_format(shape, "{TRABAJO}", paginas_texto[0])

        for extra_text in paginas_texto[1:]:
            new_slide = add_slide_before_last(prs, layout_contenido)
            target = None; max_area = 0
            for sh in new_slide.shapes:
                if sh.has_text_frame:
                    area = sh.width * sh.height
                    if area > max_area:
                        max_area = area; target = sh
            if target:
                target.text_frame.word_wrap = True
                target.text_frame.clear()
                target.text_frame.paragraphs[0].text = extra_text

    img_idx = 0
    for i in range(2, len(list(prs.slides))-1):
        if img_idx >= len(images): break
        slide = prs.slides[i]
        if any(s.shape_type == 13 for s in slide.shapes): continue
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path = tf.name
        slide.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3))
        os.unlink(path); img_idx += 1

    while img_idx < len(images):
        s = add_slide_before_last(prs, layout_contenido)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path = tf.name
        s.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3))
        os.unlink(path); img_idx += 1

    bio = io.BytesIO(); prs.save(bio); bio.seek(0)
    safe_empresa = re.sub(r'[^A-Za-z0-9]+','_',empresa).upper()[:20]
    safe_tecnico = re.sub(r'[^A-Za-z0-9]+','_',tecnico).upper()[:15]
    return bio.getvalue(), f"{safe_empresa}_{fecha_file}_{safe_tecnico}.pptx"

@app.get("/")
def home():
    return {"ok": True}

@app.get("/webhook")
def verify(hub_verify_token: str = None, hub_challenge: str = None):
    if hub_verify_token == VERIFY_TOKEN:
        return int(hub_challenge)
    return "Token invalido"

@app.post("/webhook")
async def webhook(req: Request):
    data = await req.json()
    try:
        value = data["entry"][0]["changes"][0]["value"]
        if "statuses" in value or "messages" not in value:
            return "ok"
        msg = value["messages"][0]
        from_num = msg["from"]
        SESSIONS.setdefault(from_num, {"texts":[],"images":[]})

        if msg["type"] == "text":
            body = msg["text"]["body"].strip()
            if body.lower() in ["generar reporte","generar","reporte"]:
                sess = SESSIONS.get(from_num, {"texts":[],"images":[]})
                if not sess["texts"] and not sess["images"]:
                    send_text(from_num, "No tengo datos.")
                    return "ok"
                send_text(from_num, f"Generando reporte con {len(sess['texts'])} textos y {len(sess['images'])} fotos...")
                pptx_bytes, filename = create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx_bytes, filename=filename)
                SESSIONS[from_num] = {"texts":[],"images":[]}
            else:
                SESSIONS[from_num]["texts"].append(body)
                send_text(from_num, "Guardado ✅")

        elif msg["type"] == "image":
            mid = msg["image"]["id"]
            h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info = requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h, timeout=15).json()
            img = requests.get(info["url"], headers=h, timeout=20).content
            SESSIONS[from_num]["images"].append(img)
            send_text(from_num, f"Foto {len(SESSIONS[from_num]['images'])} guardada")

    except Exception as e:
        print(e)
        import traceback; traceback.print_exc()
    return "ok"
