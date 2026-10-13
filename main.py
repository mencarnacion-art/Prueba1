import os, io, re, tempfile, requests, time
from datetime import datetime
from fastapi import FastAPI, Request
from pptx import Presentation
from pptx.util import Inches

app = FastAPI()
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID") or os.getenv("PHONE_NUMBER_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN") or "automatyco123"
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY") or os.getenv("GEMINI_API_KEY")
SESSIONS = {}; TEMPLATE_PATH = "Reporte.pptx"

def send_text(to, text):
    requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
        headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"},
        json={"messaging_product":"whatsapp","to":to,"type":"text","text":{"body":text[:3500]}})

def send_doc(to, pptx_bytes, filename):
    h={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    files={'file':(filename,pptx_bytes,'application/vnd.openxmlformats-officedocument.presentationml.presentation'),'type':(None,'application/vnd.openxmlformats-officedocument.presentationml.presentation'),'messaging_product':(None,'whatsapp')}
    r=requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media",headers=h,files=files, timeout=30)
    if r.status_code!=200: print(r.text); return
    mid=r.json()["id"]
    requests.post(f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages",
        headers={"Authorization": f"Bearer {WHATSAPP_TOKEN}","Content-Type":"application/json"},
        json={"messaging_product":"whatsapp","to":to,"type":"document","document":{"id":mid,"filename":filename}})

def limpiar_markdown(texto):
    # Quita **, ###, --- y deja texto limpio para PPT
    texto = re.sub(r'\*\*(.*?)\*\*', r'\1', texto) # **negrita** -> negrita
    texto = re.sub(r'__|\*\*|\*', '', texto)
    texto = re.sub(r'^#{1,6}\s*', '', texto, flags=re.MULTILINE) # ### titulo
    texto = re.sub(r'^---+\s*$', '', texto, flags=re.MULTILINE)
    texto = re.sub(r'^\s*[-*]\s+', '• ', texto, flags=re.MULTILINE)
    texto = texto.replace('`','')
    return texto.strip()

def mejorar_con_gemini(texto_raw, empresa):
    if not GOOGLE_API_KEY or not texto_raw: return None
    models = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-flash-latest"]
    for model in models:
        try:
            prompt = f"""Eres redactor técnico de Automatyco. Mejora este reporte de {empresa}.
REGLAS OBLIGATORIAS:
- NO uses markdown. NO uses **, no uses ###, no uses ---, no uses asteriscos.
- Devuelve texto plano profesional en español, con saltos de linea.
- Estructura: Proveedor, Hora, Actividad, Descripción de Actividades numerada.
Texto: {texto_raw[:7000]}"""
            url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={GOOGLE_API_KEY}"
            r = requests.post(url, headers={"Content-Type":"application/json"},
                json={"contents":[{"parts":[{"text":prompt}]}], "generationConfig":{"temperature":0.2, "maxOutputTokens":3000}}, timeout=30)
            print(f"GEMINI {model}: {r.status_code}")
            if r.status_code==200:
                txt = r.json()["candidates"][0]["content"]["parts"][0]["text"]
                return limpiar_markdown(txt)
        except Exception as e: print(e)
    return None

# ESTA FUNCION ES LA QUE ARREGLA LO DE LA PRIMERA DIAPOSITIVA
def replace_keep_format(shape, placeholder, value):
    if not shape.has_text_frame: return
    if placeholder not in shape.text: return
    for paragraph in shape.text_frame.paragraphs:
        if placeholder in paragraph.text:
            # conserva formato del primer run
            full = paragraph.text
            new_text = full.replace(placeholder, value)
            # limpia runs y pone uno solo con el formato original
            p_format = paragraph.runs[0].font if paragraph.runs else None
            paragraph.text = new_text
            if p_format and paragraph.runs:
                paragraph.runs[0].font.name = p_format.name
                paragraph.runs[0].font.size = p_format.size
                paragraph.runs[0].font.bold = p_format.bold

def parse_data(texts):
    full="\n".join(texts); empresa="CLIENTE"; tecnico="TECNICO"
    m=re.search(r'Empresa\s*:\s*(.+)', full, re.IGNORECASE)
    if m: empresa=m.group(1).split("\n")[0].strip()[:40]
    m2=re.search(r'(Tecnico|Técnico|By)\s*:\s*(.+)', full, re.IGNORECASE)
    if m2: tecnico=m2.group(2).split("\n")[0].strip()[:30]
    fecha_file=datetime.now().strftime("%Y-%m-%d"); fecha_show=datetime.now().strftime("%d/%m/%Y")
    desc=re.sub(r'Empresa\s*:.*\n?','',full,flags=re.IGNORECASE); desc=re.sub(r'(Tecnico|Técnico|By)\s*:.*\n?','',desc,flags=re.IGNORECASE).strip()
    if not desc: desc=full
    mejorado = mejorar_con_gemini(desc, empresa)
    final = limpiar_markdown(mejorado or desc)
    return empresa, tecnico, fecha_file, fecha_show, final

def add_slide_before_last(prs, layout):
    ns=prs.slides.add_slide(layout); lst=prs.slides._sldIdLst; nid=lst[-1]; lst.remove(nid); lst.insert(len(lst)-1,nid); return ns

def create_pptx(texts, images):
    empresa, tecnico, fecha_file, fecha_show, trabajo = parse_data(texts)
    prs = Presentation(TEMPLATE_PATH); layout = prs.slide_layouts[1] if len(prs.slide_layouts)>1 else prs.slide_layouts[0]
    
    # PORTADA - respetando formato
    for shape in prs.slides[0].shapes:
        replace_keep_format(shape, "{Empresa}", empresa)
        replace_keep_format(shape, "{EMPRESA}", empresa)
        replace_keep_format(shape, "{TECNICO}", tecnico)
        replace_keep_format(shape, "{Tecnico}", tecnico)
        replace_keep_format(shape, "{FECHA}", fecha_show)

    # CUERPO - reemplaza {TRABAJO} y activa autoajuste para no tapar logo
    for slide in list(prs.slides)[1:-1]:
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                shape.text_frame.word_wrap = True
                shape.text_frame.auto_size = None
                replace_keep_format(shape, "{TRABAJO}", trabajo)
                # Si el texto es muy largo, reduce fuente un poco
                if len(trabajo) > 800:
                    for p in shape.text_frame.paragraphs:
                        for run in p.runs:
                            if run.font.size:
                                run.font.size = run.font.size * 0.9

    img_idx=0
    for i in range(2, len(list(prs.slides))-1):
        if img_idx>=len(images): break
        slide=prs.slides[i]
        if any(s.shape_type==13 for s in slide.shapes): continue
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf: tf.write(images[img_idx]); path=tf.name
        slide.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3)); os.unlink(path); img_idx+=1
    while img_idx < len(images):
        s=add_slide_before_last(prs, layout)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf: tf.write(images[img_idx]); path=tf.name
        s.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3)); os.unlink(path); img_idx+=1
    bio=io.BytesIO(); prs.save(bio); bio.seek(0)
    safe_empresa=re.sub(r'[^A-Za-z0-9]+','_',empresa).upper()[:20]; safe_tecnico=re.sub(r'[^A-Za-z0-9]+','_',tecnico).upper()[:15]
    return bio.getvalue(), f"{safe_empresa}_{fecha_file}_{safe_tecnico}.pptx"

@app.get("/")
def home(): return {"ok":True}
@app.get("/webhook")
def verify(hub_verify_token: str = None, hub_challenge: str = None):
    if hub_verify_token==VERIFY_TOKEN: return int(hub_challenge)
    return "Token invalido"
@app.post("/webhook")
async def webhook(req: Request):
    data=await req.json()
    try:
        v=data["entry"][0]["changes"][0]["value"]
        if "statuses" in v or "messages" not in v: return "ok"
        msg=v["messages"][0]; from_num=msg["from"]
        SESSIONS.setdefault(from_num, {"texts":[],"images":[]})
        if msg["type"]=="text":
            body=msg["text"]["body"].strip()
            if body.lower() in ["generar reporte","generar","reporte"]:
                sess=SESSIONS.get(from_num, {"texts":[],"images":[]})
                if not sess["texts"] and not sess["images"]:
                    send_text(from_num, "No tengo datos para reporte."); return "ok"
                send_text(from_num, f"Generando reporte limpio...")
                pptx_bytes, filename = create_pptx(sess["texts"], sess["images"])
                send_doc(from_num, pptx_bytes, filename=filename)
                SESSIONS[from_num]={"texts":[],"images":[]}
            else:
                SESSIONS[from_num]["texts"].append(body); send_text(from_num,"Guardado ✅")
        elif msg["type"]=="image":
            mid=msg["image"]["id"]; h={"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info=requests.get(f"https://graph.facebook.com/v20.0/{mid}", headers=h).json()
            img=requests.get(info["url"], headers=h).content
            SESSIONS[from_num]["images"].append(img); send_text(from_num,f"Foto {len(SESSIONS[from_num]['images'])} guardada")
    except Exception as e: print(e); import traceback; traceback.print_exc()
    return "ok"
