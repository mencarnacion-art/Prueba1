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
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
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
    if not t: return ""
    t = re.sub(r'\*\*(.*?)\*\*', r'\1', t)
    t = re.sub(r'[*_`#]', '', t)
    return t.strip()

def mejorar_texto_profesional(texto_raw, empresa):
    texto_raw = texto_raw.strip()
    if not texto_raw: return texto_raw
    # Si es muy largo > 5000, lo partimos para que Gemini no lo recorte
    # Si es largo, lo corregimos por partes
    partes_a_corregir = [texto_raw[i:i+3500] for i in range(0, len(texto_raw), 3500)]
    resultado_final = ""

    modelos = ["models/gemini-3-flash-preview"] # SOLO EL QUE TE FUNCIONO

    for parte in partes_a_corregir:
        corregida = None
        if GOOGLE_API_KEY:
            try:
                url = f"https://generativelanguage.googleapis.com/v1beta/{modelos[0]}:generateContent?key={GOOGLE_API_KEY}"
                payload = {
                    "contents": [{"parts": [{"text": f"Eres redactor técnico profesional de Automatyco. Corrige ortografía, gramática y redacta de forma profesional y técnica, sin resumir, sin omitir detalles, sin cortar, prohibido usar ** o ## o markdown, solo texto plano. Mantén todo el contenido original. Cliente {empresa}: {parte}"}]}],
                    "generationConfig": {"temperature": 0.2, "maxOutputTokens": 8192}
                }
                r = requests.post(url, headers={"Content-Type":"application/json"}, json=payload, timeout=40)
                print(f"GEMINI {modelos[0]} {r.status_code} parte {len(parte)}")
                if r.status_code == 200:
                    txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                    # Si recorta más del 30%, no lo usamos
                    if len(txt) < len(parte) * 0.6:
                        print(f"Recorte excesivo {len(txt)} vs {len(parte)}, uso original")
                        corregida = parte
                    else:
                        corregida = txt
            except Exception as e:
                print(f"Gemini error {e}")

        if not corregida:
            # Fallback OpenAI si tienes key
            if OPENAI_API_KEY and len(parte) < 4000:
                try:
                    ro = requests.post("https://api.openai.com/v1/chat/completions",
                        headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type":"application/json"},
                        json={"model":"gpt-4o-mini","messages":[{"role":"user","content": f"Corrige ortografía profesional sin resumir ni cortar, sin **. Cliente {empresa}: {parte}"}], "max_tokens": 4000}, timeout=20)
                    if ro.status_code == 200:
                        corregida = ro.json()["choices"][0]["message"]["content"]
                except: pass

        if not corregida:
            corregida = parte

        resultado_final += " " + corregida

    return limpiar_markdown(resultado_final)

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
    paginas = []
    actual = ""
    for o in oraciones:
        if len(actual) + len(o) + 1 < max_chars:
            actual = (actual + " " + o).strip() if actual else o
        else:
            if actual: paginas.append(actual.strip())
            if len(o) > max_chars:
                palabras = o.split()
                tmp = ""
                for p in palabras:
                    if len(tmp) + len(p) < max_chars:
                        tmp = (tmp + " " + p).strip() if tmp else p
                    else:
                        paginas.append(tmp.strip())
                        tmp = p
                actual = tmp
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
    final = mejorar_texto_profesional(desc, empresa)
    print(f"TEXTO FINAL CORREGIDO {len(final)} chars")
    return empresa, tecnico, fecha_file, fecha_show, final

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH)

    # Layouts
    layout_text = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
    layout_blank = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else layout_text

    # 1. Portada
    for shape in prs.slides[0].shapes:
        replace_keep_format(shape, "{Empresa}", empresa)
        replace_keep_format(shape, "{EMPRESA}", empresa)
        replace_keep_format(shape, "{TECNICO}", tecnico)
        replace_keep_format(shape, "{Tecnico}", tecnico)
        replace_keep_format(shape, "{FECHA}", fecha_show)

    # 2. IMPORTANTE: Borrar todas las diapositivas del medio de la plantilla
    # Nos quedamos solo con portada (0) y final (última), el resto lo creamos nosotros en orden
    # Guardamos la diapositiva final de gracias
    last_slide = prs.slides[-1]
    # Borrar del final al principio para no mover índices
    # Dejamos slide 0 y slide 1 (texto base)
    while len(prs.slides) > 2:
        # Borra la slide 1 hacia adelante hasta dejar solo portada y la de gracias
        # Pero necesitamos conservar la última (gracias)
        # Así que borramos la slide en índice 1 si hay más de 2
        if len(prs.slides) > 2:
            # Si estamos borrando, necesitamos mantener la última
            # La técnica: borrar slide 2 en adelante, dejando 0 y 1 y última
            # Para simplificar: si len > 3, borra la slide 2
            if len(prs.slides) > 3:
                slide_id = prs.slides[2].slide_id
                prs.part.drop_rel(prs.slides._sldIdLst[2].rId)
                del prs.slides._sldIdLst[2]
            else:
                break
        else:
            break

    # Ahora tenemos: [0]=portada, [1]=texto base plantilla, [2]=gracias
    # Vamos a poner todo el texto en orden
    paginas = dividir_inteligente(trabajo, 850)
    print(f"PAGINAS TEXTO: {len(paginas)}")

    # Slide 1 = primera página de texto
    for shape in prs.slides[1].shapes:
        if shape.has_text_frame and "{TRABAJO}" in shape.text:
            shape.text_frame.word_wrap = True
            replace_keep_format(shape, "{TRABAJO}", paginas[0] if paginas else trabajo)

    # Crear slides de continuación de texto JUNTOS
    # Insertar antes de la de gracias
    for extra in paginas[1:]:
        new_slide = prs.slides.add_slide(layout_text)
        # Moverla antes de la última (gracias)
        sldIdLst = prs.slides._sldIdLst
        last = sldIdLst[-1]
        sldIdLst.remove(last)
        sldIdLst.append(last) # esto no, mejor insertar
        # Reordenar correctamente: la nueva está al final, la movemos a penúltima
        new_id = sldIdLst[-1]
        sldIdLst.remove(new_id)
        sldIdLst.insert(len(sldIdLst)-1, new_id)

        txBox = new_slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(6))
        tf = txBox.text_frame
        tf.word_wrap = True
        tf.text = extra

    # Ahora Fotos DESPUÉS del texto
    for img_bytes in images:
        if len(img_bytes) < 20000:
            print("Foto pequeña ignorada")
            continue
        new_slide = prs.slides.add_slide(layout_blank)
        sldIdLst = prs.slides._sldIdLst
        new_id = sldIdLst[-1]
        sldIdLst.remove(new_id)
        sldIdLst.insert(len(sldIdLst)-1, new_id)

        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(img_bytes); path = tf.name
        try:
            new_slide.shapes.add_picture(path, Inches(0.5), Inches(0.5), Inches(9), Inches(5.5))
        except: pass
        os.unlink(path)

    bio = io.BytesIO(); prs.save(bio); bio.seek(0)
    safe = re.sub(r'[^A-Za-z0-9]+','_',empresa).upper()[:20]
    return bio.getvalue(), f"{safe}_{fecha_file}.pptx"

@app.get("/")
def home(): return {"ok": True, "V13": "Texto corregido + orden texto/fotos"}

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
                send_text(frm, "Corrigiendo texto profesional y generando reporte...")
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
