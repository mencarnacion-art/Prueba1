import os, io, base64, json, re, tempfile, requests
from flask import Flask, request, jsonify
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

app = Flask(__name__)

# --- CONFIG WHATSAPP ---
WHATSAPP_TOKEN = os.getenv("WHATSAPP_TOKEN")
WHATSAPP_PHONE_ID = os.getenv("WHATSAPP_PHONE_ID")
VERIFY_TOKEN = os.getenv("VERIFY_TOKEN", "automatyco123")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

# Memoria temporal por número
SESSIONS = {}

def send_whatsapp_text(to, text):
    url = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    data = {"messaging_product": "whatsapp", "to": to, "type": "text", "text": {"body": text}}
    requests.post(url, headers=headers, json=data)

def send_whatsapp_doc(to, pptx_bytes, filename="Reporte.pptx", caption="Aquí está tu reporte"):
    # 1. Subir archivo a Meta
    url_media = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/media"
    headers = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
    files = {'file': (filename, pptx_bytes, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'), 'type': (None, 'application/vnd.openxmlformats-officedocument.presentationml.presentation'), 'messaging_product': (None, 'whatsapp')}
    r = requests.post(url_media, headers=headers, files=files)
    if r.status_code!= 200:
        print("Error subiendo media:", r.text)
        send_whatsapp_text(to, "Error subiendo el PPTX: " + r.text)
        return
    media_id = r.json()["id"]
    # 2. Mandar documento
    url_msg = f"https://graph.facebook.com/v20.0/{WHATSAPP_PHONE_ID}/messages"
    headers_json = {"Authorization": f"Bearer {WHATSAPP_TOKEN}", "Content-Type": "application/json"}
    data = {"messaging_product": "whatsapp", "to": to, "type": "document", "document": {"id": media_id, "filename": filename, "caption": caption}}
    requests.post(url_msg, headers=headers_json, json=data)

def transcribe_audio(media_id):
    if not OPENAI_API_KEY:
        return "[Audio recibido - pon OPENAI_API_KEY para transcribir]"
    try:
        # bajar audio de WhatsApp
        url_info = f"https://graph.facebook.com/v20.0/{media_id}"
        h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
        info = requests.get(url_info, headers=h).json()
        audio_url = info["url"]
        audio_data = requests.get(audio_url, headers=h).content
        with tempfile.NamedTemporaryFile(suffix=".ogg", delete=False) as f:
            f.write(audio_data)
            fpath = f.name
        # whisper
        with open(fpath, "rb") as af:
            tr = requests.post("https://api.openai.com/v1/audio/transcriptions", headers={"Authorization": f"Bearer {OPENAI_API_KEY}"}, files={"file": af}, data={"model": "whisper-1", "language": "es"})
        os.unlink(fpath)
        return tr.json().get("text", "[No se pudo transcribir]")
    except Exception as e:
        print(e)
        return f"[Error transcribiendo audio: {e}]"

def create_pptx(texts, images_bytes):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)
    # Slide 1 - Resumen
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(12), Inches(6)).text_frame.text = "\n".join(texts) if texts else "Reporte sin texto"

    for img_b in images_bytes:
        try:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf:
                tf.write(img_b)
                tfpath = tf.name
            slide.shapes.add_picture(tfpath, Inches(0.5), Inches(0.5), Inches(12), Inches(6.5))
            os.unlink(tfpath)
        except Exception as e:
            print("Error imagen:", e)
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.getvalue()

@app.route("/", methods=["GET"])
def home():
    return "Bot Automatyco OK - Solo WhatsApp"

@app.route("/webhook", methods=["GET"])
def verify():
    if request.args.get("hub.verify_token") == VERIFY_TOKEN:
        return request.args.get("hub.challenge")
    return "Token invalido", 403

@app.route("/webhook", methods=["POST"])
def webhook():
    data = request.get_json()
    try:
        entry = data["entry"][0]["changes"][0]["value"]
        if "messages" not in entry:
            return "ok", 200
        msg = entry["messages"][0]
        from_num = msg["from"]
        if from_num not in SESSIONS:
            SESSIONS[from_num] = {"texts": [], "images": []}

        mtype = msg["type"]
        if mtype == "text":
            body = msg["text"]["body"].strip()
            if body.lower() in ["generar reporte", "generar", "reporte"]:
                sess = SESSIONS[from_num]
                if not sess["texts"] and not sess["images"]:
                    send_whatsapp_text(from_num, "No tengo nada aún. Mándame fotos y descripción primero.")
                    return "ok", 200
                send_whatsapp_text(from_num, f"Generando reporte con {len(sess['images'])} fotos... dame 10 seg")
                pptx = create_pptx(sess["texts"], sess["images"])
                send_whatsapp_doc(from_num, pptx, filename=f"Reporte_{from_num}.pptx", caption="Listo ✅ Reporte generado")
                SESSIONS[from_num] = {"texts": [], "images": []} # limpiar
            else:
                SESSIONS[from_num]["texts"].append(body)
                send_whatsapp_text(from_num, f"Texto guardado ({len(SESSIONS[from_num]['texts'])}). Manda más o escribe *generar reporte*")
        elif mtype == "image":
            media_id = msg["image"]["id"]
            # descargar imagen
            url_info = f"https://graph.facebook.com/v20.0/{media_id}"
            h = {"Authorization": f"Bearer {WHATSAPP_TOKEN}"}
            info = requests.get(url_info, headers=h).json()
            img_url = info["url"]
            img_data = requests.get(img_url, headers=h).content
            SESSIONS[from_num]["images"].append(img_data)
            send_whatsapp_text(from_num, f"Foto {len(SESSIONS[from_num]['images'])} guardada. Escribe *generar reporte* cuando termines")
        elif mtype == "audio":
            media_id = msg["audio"]["id"]
            txt = transcribe_audio(media_id)
            SESSIONS[from_num]["texts"].append(f"[AUDIO]: {txt}")
            send_whatsapp_text(from_num, f"Audio transcrito: {txt[:200]}... Sigue o pon *generar reporte*")
    except Exception as e:
        print("Error webhook:", e)
    return "ok", 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", 8080)))
