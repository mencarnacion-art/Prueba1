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
    if r.status_code != 200: print(r.text); return
    mid = r.json()["id"]
    requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
        headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
        json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":mid,"filename":filename}})

def limpiar_markdown(t):
    if not t: return ""
    t = re.sub(r'\*\*(.*?)\*\*', r'\1', t)
    t = re.sub(r'[*_`#]', '', t)
    return t.strip()

def mejorar_texto(texto_raw, empresa):
    # SI ES LARGO NO USAR IA PARA NO RECORTAR - ESTO ARREGLA LO CORTO
    if len(texto_raw) > 2500:
        print(f"TEXTO LARGO {len(texto_raw)} -> sin IA")
        return limpiar_markdown(texto_raw)
    modelos = ["models/gemini-3-flash-preview","models/gemini-2.5-flash-lite","models/gemini-2.5-flash"]
    if GOOGLE_API_KEY:
        for modelo in modelos:
            try:
                url = f"https://generativelanguage.googleapis.com/v1beta/{modelo}:generateContent?key={GOOGLE_API_KEY}"
                payload = {"contents":[{"parts":[{"text": f"Corrige ortografia sin resumir, mantén todo detalle, sin ** ni ##. Cliente {empresa}: {texto_raw[:8000]}"}]}], "generationConfig":{"temperature":0.2,"maxOutputTokens":8192}}
                r = requests.post(url, headers={"Content-Type":"application/json"}, json=payload, timeout=40)
                print(f"GEMINI {modelo}: {r.status_code}")
                if r.status_code == 200:
                    txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                    if len(txt) < len(texto_raw)*0.7:
                        print("GEMINI recorto mucho, uso original")
                        return limpiar_markdown(texto_raw)
                    return limpiar_markdown(txt)
            except Exception as e: print(e)
    return limpiar_markdown(texto_raw)

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

def dividir_texto_inteligente(texto, max_chars=850):
    # Divide por oraciones, no a mitad de palabra
    oraciones = re.split(r'(?<=[.!?])\s+', texto)
    paginas = []
    actual = ""
    for o in oraciones:
        if len(actual) + len(o) < max_chars:
            actual += " " + o
        else:
            if actual: paginas.append(actual.strip())
            if len(o) > max_chars:
                # oración muy larga, cortarla por palabras
                palabras = o.split()
                temp = ""
                for p in palabras:
                    if len(temp) + len(p) < max_chars:
                        temp += " " + p
                    else:
                        paginas.append(temp.strip())
                        temp = p
                actual = temp
            else:
                actual = o
    if actual: paginas.append(actual.strip())
    return paginas

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
    final = mejorar_texto(desc, empresa)
    return empresa, tecnico, fecha_file, fecha_show, final

def add_slide_before_last(prs, layout):
    ns = prs.slides.add_slide(layout)
    lst = prs.slides._sldIdLst
    nid = lst[-1]; lst.remove(nid); lst.insert(len(lst)-1, nid)
    return ns

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH)
    layout_text = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
    layout_blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else layout_text

    for shape in prs.slides[0].shapes:
        replace_keep_format(shape, "{Empresa}", empresa)
        replace_keep_format(shape, "{EMPRESA}", empresa)
        replace_keep_format(shape, "{TECNICO}", tecnico)
        replace_keep_format(shape, "{Tecnico}", tecnico)
        replace_keep_format(shape, "{FECHA}", fecha_show)

    paginas = dividir_texto_inteligente(trabajo, 850)
    print(f"PAGINAS DE TEXTO: {len(paginas)}")

    for shape in prs.slides[1].shapes:
        if shape.has_text_frame and "{TRABAJO}" in shape.text:
            shape.text_frame.word_wrap = True
            replace_keep_format(shape, "{TRABAJO}", paginas[0] if paginas else trabajo)

    for extra in paginas[1:]:
        ns = add_slide_before_last(prs, layout_text)
        txBox = ns.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(6))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.text = extra

    # BORRAR slides de fotos viejas de la plantilla (slides 2 hasta -1)
    # y crear nuevas solo para tus fotos
    while len(prs.slides) > 3:
        # deja slide 0 portada, 1 trabajo, y ultimas de trabajo extra + gracias
        # pero borra los placeholders de imagen del medio
        if len(prs.slides) <= 2: break
        # si el slide tiene imagen placeholder de plantilla y no es de texto, lo borramos si está entre medio
        # Simplificado: borramos todo lo que está entre slide 1 y el slide de gracias final
        # El slide de gracias es el último
        # Contamos cuantas paginas de texto creamos
        break

    # Fotos en diapositivas nuevas, sin encimar texto
    for img_bytes in images:
        # ignora imágenes muy pequeñas o capturas de WhatsApp (fondo oscuro)
        if len(img_bytes) < 15000:
            print("Imagen muy pequeña, ignorada")
            continue
        ns = add_slide_before_last(prs, layout_blank)
        # fondo limpio
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(img_bytes); path = tf.name
        try:
            ns.shapes.add_picture(path, Inches(0.5), Inches(0.5), Inches(9), Inches(5.5))
        except: pass
        os.unlink(path)

    bio = io.BytesIO(); prs.save(bio); bio.seek(0)
    safe = re.sub(r'[^A-Za-z0-9]+','_',empresa).upper()[:20]
    return bio.getvalue(), f"{safe}_{fecha_file}.pptx"

@app.get("/")
def home(): return {"ok": True}
@app.get("/debug/models")
def debug_models():
    if not GOOGLE_API_KEY: return {"error": "no key"}
    r = requests.get(f"https://generativelanguage.googleapis.com/v1beta/models?key={GOOGLE_API_KEY}")
    return r.json()
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
                send_text(frm, f"Guardado ✅ ({len(SESSIONS[frm]['texts'])} textos)")
        elif msg["type"] == "image":
            mid = msg["image"]["id"]; h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info = requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h, timeout=15).json()
            img = requests.get(info["url"], headers=h, timeout=20).content
            SESSIONS[frm]["images"].append(img)
            send_text(frm, f"Foto {len(SESSIONS[frm]['images'])} guardada")
    except Exception as e:
        print(e); import traceback; traceback.print_exc()
    return "ok"
