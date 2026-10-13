def split_texto_en_paginas(texto, max_chars=900):
    # Parte por párrafos para no cortar a mitad
    parrafos = texto.split('\n')
    paginas = []
    actual = ""
    for p in parrafos:
        if len(actual) + len(p) + 1 < max_chars:
            actual += p + "\n"
        else:
            if actual: paginas.append(actual.strip())
            # si un solo párrafo es larguísimo, córtalo
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
    layout_contenido = prs.slide_layouts[1] if len(prs.slide_layouts)>1 else prs.slide_layouts[0]

    # 1. PORTADA
    for shape in prs.slides[0].shapes:
        replace_keep_format(shape, "{Empresa}", empresa)
        replace_keep_format(shape, "{EMPRESA}", empresa)
        replace_keep_format(shape, "{TECNICO}", tecnico)
        replace_keep_format(shape, "{Tecnico}", tecnico)
        replace_keep_format(shape, "{FECHA}", fecha_show)

    # 2. TEXTO PAGINADO
    paginas_texto = split_texto_en_paginas(trabajo, max_chars=850) # ajusta aquí si quieres más/menos texto por diapo

    # Busca la diapositiva que tiene {TRABAJO}
    slide_trabajo_idx = None
    shape_trabajo_original = None
    for idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                slide_trabajo_idx = idx
                shape_trabajo_original = shape
                break
        if slide_trabajo_idx is not None: break

    if slide_trabajo_idx is not None:
        # Primera página en la diapo original
        for shape in prs.slides[slide_trabajo_idx].shapes:
            if shape.has_text_frame and "{TRABAJO}" in shape.text:
                shape.text_frame.word_wrap = True
                replace_keep_format(shape, "{TRABAJO}", paginas_texto[0])

        # Si hay más texto, crea diapositivas nuevas antes del final
        for extra_text in paginas_texto[1:]:
            new_slide = add_slide_before_last(prs, layout_contenido)
            # limpia la nueva diapo y pone título + texto
            # Busca el placeholder de texto más grande
            target_shape = None
            max_area = 0
            for sh in new_slide.shapes:
                if sh.has_text_frame:
                    area = sh.width * sh.height
                    if area > max_area:
                        max_area = area
                        target_shape = sh

            if target_shape:
                target_shape.text_frame.word_wrap = True
                target_shape.text_frame.clear()
                p = target_shape.text_frame.paragraphs[0]
                p.text = extra_text
                # Título opcional arriba
                for sh in new_slide.shapes:
                    if sh.has_text_frame and sh!= target_shape and "Haga clic" in sh.text:
                        sh.text = "Descripción de Actividades e Incidencias (Continuación)"

    # 3. FOTOS
    img_idx=0
    # Recorre de nuevo porque agregamos slides
    for i in range(2, len(list(prs.slides))-1):
        if img_idx>=len(images): break
        slide=prs.slides[i]
        # No poner foto donde hay texto de trabajo
        if any("{TRABAJO}" in s.text for s in slide.shapes if s.has_text_frame): continue
        if any(s.shape_type==13 for s in slide.shapes): continue
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf: tf.write(images[img_idx]); path=tf.name
        slide.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3)); os.unlink(path); img_idx+=1

    while img_idx < len(images):
        s=add_slide_before_last(prs, layout_contenido)
        with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tf: tf.write(images[img_idx]); path=tf.name
        s.shapes.add_picture(path, Inches(0.5), Inches(1.2), Inches(12.3), Inches(5.3)); os.unlink(path); img_idx+=1

    bio=io.BytesIO(); prs.save(bio); bio.seek(0)
    safe_empresa=re.sub(r'[^A-Za-z0-9]+','_',empresa).upper()[:20]; safe_tecnico=re.sub(r'[^A-Za-z0-9]+','_',tecnico).upper()[:15]
    return bio.getvalue(), f"{safe_empresa}_{fecha_file}_{safe_tecnico}.pptx"
