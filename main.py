import os, io, re, tempfile, requests
from datetime import datetime
from fastapi import FastAPI, Request
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID") or os.getenv("PHONE_NUMBER_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY") or os.getenv("OPENAI_KEY")
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
    if r.status_code!=200: print(r.text); send_text(to, f"Error subiendo: {r.text[:300]}"); return
    media_id=r.json()["id"]
    url_msg=f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    hj={"Authorization": f"Bearer {WHATSAPP_TOKEN}","Content-Type":"application/json"}
    requests.post(url_msg, headers=hj, json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":media_id,"filename":filename,"caption":f"Reporte {filename}"}})

def mejorar_texto(texto):
    if not OPENAI_API_KEY or len(texto) < 10:
        return texto
    try:
        r = requests.post("https://api.openai.com/v1/chat/completions",
            headers={"Authorization": f"Bearer {OPENAI_API_KEY}", "Content-Type": "application/json"},
            json={
                "model": "gpt-4o-mini",
                "messages": [
                    {"role": "system", "content": "Eres corrector de reportes técnicos de mantenimiento industrial. Corrige ortografía, puntuación y mejora redacción profesional en español. No inventes datos. Mantén términos técnicos (robot, arnés, servo, etc). Devuelve SOLO el texto mejorado."},
                    {"role": "user", "content": texto}
                ],
                "temperature": 0.3
            }, timeout=20)
        if r.status_code==200:
            return r.json()["choices"][0]["message"]["content"].strip()
    except Exception as e:
        print(f"OpenAI error: {e}")
    return texto

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
    desc_mejorado = mejorar_texto(desc)
    return empresa, tecnico, fecha_file, fecha_show, desc_mejorado

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
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH) if os.path.exists(TEMPLATE_PATH) else Presentation()
    original_count = len(prs.slides)

    # 1. Portada
    for shape in prs.slides[0].shapes:
        replace_keep_format(shape,"{Empresa}",empresa)
        replace_keep_format(shape,"{EMPRESA}",empresa)
        replace_keep_format(shape,"{TECNICO}",tecnico)
        replace_keep_format(shape,"{Tecnico}",tecnico)
        replace_keep_format(shape,"{FECHA}",fecha_show)

    # 2. {TRABAJO}
    for slide in list(prs.slides)[1:]:
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                for p in shape.text_frame.paragraphs:
                    for run in p.runs:
                        if "{TRABAJO}" in run.text:
                            run.text=run.text.replace("{TRABAJO}", trabajo)

    # 3. Fotos - NUNCA tocar la última slide (es cierre)
    # Sacamos la última temporalmente para protegerla
    sldIdLst = prs.slides._sldIdLst
    last_sldId = sldIdLst[-1] if original_count > 1 else None
    if last_sldId and original_count > 1:
        sldIdLst.remove(last_sldId)

    layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    img_idx = 0

    # Usar slides vacías existentes (desde la 3 hasta la penúltima)
    # Rango 2 hasta original_count-1 (excluye última que ya sacamos)
    for idx in range(2, original_count-1):
        if img_idx >= len(images): break
        if idx >= len(prs.slides): break
        slide = prs.slides[idx]
        try:
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
                tf.write(images[img_idx]); path=tf.name
            # FOTO CENTRADA Y CON MARGEN (no 12 pulgadas)
            slide.shapes.add_picture(path, Inches(1.2), Inches(1.2), Inches(10.8), Inches(5.5))
            os.unlink(path)
            img_idx+=1
        except Exception as e:
            print(e)

    # Resto de fotos en slides nuevas
    while img_idx < len(images):
        s = prs.slides.add_slide(layout)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
            tf.write(images[img_idx]); path=tf.name
        s.shapes.add_picture(path, Inches(1.2), Inches(1.2), Inches(10.8), Inches(5.5))
        os.unlink(path)
        img_idx+=1

    # Regresar la última al final intacta
    if last_sldId:
        sldIdLst.append(last_sldId)

    bio=io.BytesIO(); prs.save(bio); bio.seek(0)

    # Nombre archivo: EMPRESA_FECHA_TECNICO.pptx
    safe_empresa = re.sub(r'[^A-Za-z0-9_-]+','_',empresa).upper()[:20]
    safe_tecnico = re.sub(r'[^A-Za-z0-9_-]+','_',tecnico).upper()[:15]
    filename = f"{safe_empresa}_{fecha_file}_{safe_tecnico}.pptx"
    return bio.getvalue(), filename

@app.get("/")
def home(): return {"ok":True, "plantilla": os.path.exists(TEMPLATE_PATH)}

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
                if not sess["texts"] and not sess["images"]:
                    send_text(from_num, "No tengo datos"); return "ok"
                send_text(from_num, f"Mejorando redacción y generando {len(sess['images'])} fotos...")
                pptx_bytes, filename = create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx_bytes, filename=filename)
                SESSIONS[from_num]={"texts":[],"images":[]}
            else:
                SESSIONS[from_num]["texts"].append(body)
                send_text(from_num, f"Guardado ✅ ({len(SESSIONS[from_num]['texts'])} textos)")
        elif msg["type"]=="image":
            mid=msg["image"]["id"]
            h={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info=requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h).json()
            img=requests.get(info["url"], headers=h).content
            SESSIONS[from_num]["images"].append(img)
            send_text(from_num, f"Foto {len(SESSIONS[from_num]['images'])} guardada 📸")
    except Exception as e:
        print(e); import traceback; traceback.print_exc()
    return "ok"
