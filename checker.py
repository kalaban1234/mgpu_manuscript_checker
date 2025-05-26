import docx
from io import BytesIO
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
HEADER_KEYWORDS = [
    "введение", "цель исследования", "материалы и методы",
    "заключение", "список источников", "сведения об авторах", "references",
    "результаты", "результаты исследования", "обсуждение"
]

def is_probable_header(paragraph):
    text = paragraph.text.strip().lower()
    words = text.split()
    # Короткие заголовки или ключевые слова или жирный абзац
    if len(words) <= 12:
        return True
    if any(text.startswith(h) for h in HEADER_KEYWORDS):
        return True
    if all(run.bold for run in paragraph.runs if run.text.strip()):
        return True
    if paragraph.alignment == WD_ALIGN_PARAGRAPH.CENTER:
        return True
    return False

def check_docx(file_bytes):
    doc = docx.Document(BytesIO(file_bytes))
    report = []
    paragraphs = doc.paragraphs

    # 1. Проверка УДК
    idx = 0
    for i, p in enumerate(paragraphs[:5]):
        if p.text.strip().lower().startswith("удк"):
            udk_p = p
            idx = i
            font_names = [r.font.name for r in p.runs if r.text.strip()]
            font_sizes = [r.font.size.pt if r.font.size else None for r in p.runs if r.text.strip()]
            bolds = [r.bold for r in p.runs if r.text.strip()]
            if any(f != "Times New Roman" for f in font_names) or any(s != 14 for s in font_sizes) or any(bolds):
                report.append({"status": "warn", "msg": "УДК должен быть Times New Roman 14 пт, не жирный"})
            if udk_p.alignment not in [None, 0]:
                report.append({"status": "warn", "msg": "УДК должен быть по левому краю"})
            break
    else:
        report.append({"status": "error", "msg": "В первом абзаце не найден УДК"})
        idx = -1

    # 2. Поиск авторов (подряд идущие абзацы после УДК)
    authors_end = idx + 1
    for i in range(idx + 1, len(paragraphs)):
        text = paragraphs[i].text.strip()
        if not text:
            continue
        # Регулярки ФИО
        is_fio = (
                re.match(r"^[А-ЯЁA-Z][а-яёa-z]+\s[А-ЯЁA-Z]\.[А-ЯЁA-Z]\.$", text) or
                re.match(r"^[А-ЯA-Z]\.[А-ЯA-Z]\.\s*[А-ЯЁA-Z][а-яёa-z]+", text) or
                re.match(r"^[А-ЯЁA-Z]{1}\.[А-ЯЁA-Z]{1}\.\s*[А-ЯЁA-Z][а-яёa-z]+", text) or
                re.match(r"^[А-ЯЁA-Z]\.[А-ЯЁA-Z]\.\s*[А-ЯЁA-Z][а-яёa-z]+", text)
        )
        if is_fio:
            p = paragraphs[i]
            font_names = [r.font.name for r in p.runs if r.text.strip()]
            font_sizes = [r.font.size.pt if r.font.size else None for r in p.runs if r.text.strip()]
            bolds = [r.bold for r in p.runs if r.text.strip()]
            if any(f != "Times New Roman" for f in font_names):
                report.append({"status": "warn", "msg": f"ФИО автора '{text}' должен быть Times New Roman"})
            if any(s != 14 for s in font_sizes):
                report.append({"status": "warn", "msg": f"ФИО автора '{text}' должен быть 14 пт"})
            if not all(bolds):
                report.append({"status": "warn", "msg": f"ФИО автора '{text}' должен быть полужирным (bold)"})
            if p.alignment not in [None, 0]:
                report.append({"status": "warn", "msg": f"ФИО автора '{text}' должен быть по левому краю"})
            authors_end = i + 1
        else:
            break

    # 3. Название — первый жирный, по центру абзац после авторов
    found_title = False
    for i in range(authors_end, len(paragraphs)):
        p = paragraphs[i]
        text = p.text.strip()
        if not text:
            continue
        font_names = [r.font.name for r in p.runs if r.text.strip()]
        font_sizes = [r.font.size.pt if r.font.size else None for r in p.runs if r.text.strip()]
        bolds = [r.bold for r in p.runs if r.text.strip()]
        if text and all(bolds) and p.alignment == 1:
            found_title = True
            if any(f != "Times New Roman" for f in font_names) or any(s != 14 for s in font_sizes):
                report.append(
                    {"status": "warn", "msg": "Название статьи должно быть Times New Roman 14 пт, полужирным"})
            break
        else:
            if p.alignment != 1:
                report.append({"status": "warn", "msg": "Название статьи должно быть по центру"})
            if not all(bolds):
                report.append({"status": "warn", "msg": "Название статьи должно быть полужирным (bold)"})
            if any(f != "Times New Roman" for f in font_names) or any(s != 14 for s in font_sizes):
                report.append({"status": "warn", "msg": "Название статьи должно быть Times New Roman 14 пт"})
            break
    if not found_title:
        report.append({"status": "error",
                       "msg": "Название статьи не найдено или не соответствует требованиям (по центру, полужирное, Times New Roman 14)"})

    # --- Проверка кегля по всему основному тексту ---
    # Определяем границы основного текста
    start_idx = authors_end
    # Находим индекс "Список источников" (или "Список литературы"), чтобы не проверять библиографию
    end_idx = len(paragraphs)
    for i in range(start_idx, len(paragraphs)):
        if paragraphs[i].text.lower().startswith("список источников") or paragraphs[i].text.lower().startswith(
                "список литературы"):
            end_idx = i
            break

    # Проверяем кегль 14 по всему основному тексту
    for i in range(start_idx, end_idx):
        p = paragraphs[i]
        # Не трогаем подписи к рисункам (это отдельная логика)
        if re.match(r"(Рисунок|рисунок|Рис\.|рис\.)\s*\d+", p.text.strip(), re.IGNORECASE):
            continue
        wrong_size = None
        for run in p.runs:
            if run.text.strip() and run.font.size and run.font.size.pt != 14:
                wrong_size = run.font.size.pt
                break
        if wrong_size:
            report.append({
                "status": "warn",
                "msg": f"В абзаце найден неверный размер шрифта ({wrong_size} пт): «{p.text[:40]}...». Ожидалось 14 пт."
            })
    # --- Проверка выравнивания основного текста ---
    for i in range(start_idx, end_idx):
        p = paragraphs[i]
        if is_probable_header(p):
            continue  # Не трогаем заголовки!
        # Не подпись к рисунку
        if re.match(r"(рисунок|рис\.|рисунке|рисунку|рисунках)\s*\d+", p.text.strip(), re.IGNORECASE):
            continue
        if p.alignment not in [WD_ALIGN_PARAGRAPH.JUSTIFY]:
            report.append({
                "status": "warn",
                "msg": f"В абзаце выравнивание должно быть по ширине страницы: «{p.text[:40]}...»"
            })

    # --- 4. Подписи и ссылки на рисунки (строго: только если есть номер) ---
    drawing_captions = set()
    drawing_caption_idxs = set()
    for idx, p in enumerate(paragraphs):
        match = re.match(r"(Рисунок|рисунок|Рис\.|рис\.)\s*(\d+)", p.text.strip(), re.IGNORECASE)
        if match:
            drawing_captions.add(match.group(2))
            drawing_caption_idxs.add(idx)
            font_sizes = [r.font.size.pt if r.font.size else None for r in p.runs if r.text.strip()]
            # Для подписи к рисунку допускается кегль 12
            if any(s != 12 for s in font_sizes):
                report.append(
                    {"status": "warn", "msg": f"Подпись к рисунку '{p.text.strip()[:30]}...' должна быть 12 кеглем"})

    drawing_refs = set()
    for idx, p in enumerate(paragraphs):
        if idx in drawing_caption_idxs:
            continue  # Пропускаем подписи к рисункам!
        for m in re.findall(r"рисун[а-я]*\s*(\d+)|рис\.\s*(\d+)", p.text, re.IGNORECASE):
            num = next(filter(None, m), None)
            if num:
                drawing_refs.add(num)

    missed_in_text = drawing_captions - drawing_refs
    if missed_in_text:
        missed_str = ", ".join(sorted(missed_in_text))
        report.append({"status": "warn", "msg": f"Нет ссылок на рисунки {missed_str} в тексте"})

    missed_in_captions = drawing_refs - drawing_captions
    if missed_in_captions:
        missed_str = ", ".join(sorted(missed_in_captions))
        report.append(
            {"status": "warn", "msg": f"Есть ссылки на рисунки {missed_str} в тексте, но нет соответствующих подписей"})

    # --- 5. Проверка объема статьи (только до списка литературы) ---
    main_text = ""
    for p in paragraphs:
        if p.text.lower().startswith("список литературы") or p.text.lower().startswith("список источников"):
            break
        main_text += p.text + "\n"
    char_count = len(main_text.replace("\n", ""))
    if not (20000 <= char_count <= 40000):
        report.append({"status": "error",
                       "msg": f"Объем статьи {char_count} знаков (без списка литературы; ожидалось 20 000–40 000)."})

    # 6. Определяем индекс начала библиографии и само название (строгий стиль)
    biblio_idx = None
    biblio_title = None
    for i, p in enumerate(paragraphs):
        title = p.text.strip().lower()
        if title.startswith("список источников") or title.startswith("список литературы"):
            biblio_idx = i
            biblio_title = p.text.strip()
            break

    # Проверка наличия, кегля и выравнивания заголовка библиографии
    if biblio_idx is not None:
        biblio_p = paragraphs[biblio_idx]
        font_names = [r.font.name for r in biblio_p.runs if r.text.strip()]
        font_sizes = [r.font.size.pt if r.font.size else None for r in biblio_p.runs if r.text.strip()]
        if any(f != "Times New Roman" for f in font_names) or any(s != 14 for s in font_sizes):
            report.append({"status": "warn", "msg": f"Заголовок '{biblio_title}' должен быть Times New Roman 14 пт"})
        if biblio_p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            report.append({"status": "warn", "msg": f"Заголовок '{biblio_title}' должен быть по центру"})
        # Название должно быть строго "Список источников"
        if biblio_title.lower() != "список источников":
            report.append({"status": "error", "msg": "Название раздела должно быть строго 'Список источников'"})
    else:
        report.append(
            {"status": "error", "msg": "В тексте отсутствует заголовок 'Список источников' или 'Список литературы'"})

    if biblio_idx is not None:
        for p in paragraphs[biblio_idx + 1:]:
            if p.text.strip() == "":
                continue
            if p.text.strip().lower().startswith("references") or p.text.strip().lower().startswith(
                    "сведения об авторах"):
                break
            # Пропускаем подписи к рисункам (допускается только 12 пт)
            if re.match(r"(рисунок|рис\.|рисунке|рисунку|рисунках)\s*\d+", p.text.strip(), re.IGNORECASE):
                for run in p.runs:
                    if run.text.strip() and run.font.size and run.font.size.pt != 12:
                        report.append({
                            "status": "warn",
                            "msg": f"Подпись к рисунку в списке должна быть 12 пт: «{run.text[:40]}...»"
                        })
                continue
            # Для обычных элементов списка литературы — ловим любые отличия от 14 пт!
            has_size = False
            wrong_size = None
            for run in p.runs:
                if run.text.strip():
                    if run.font.size:
                        has_size = True
                        if run.font.size.pt != 14:
                            wrong_size = run.font.size.pt
                            break
            if has_size and wrong_size:
                report.append({
                    "status": "warn",
                    "msg": f"В абзаце найден неверный размер шрифта ({wrong_size} пт): «{p.text[:40]}...». Ожидалось 14 пт."
                })
            elif not has_size:
                report.append({
                    "status": "warn",
                    "msg": f"В абзаце не удалось определить размер шрифта: «{p.text[:40]}...». Ожидалось 14 пт."
                })
            # Проверка выравнивания абзаца
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                report.append({
                    "status": "warn",
                    "msg": f"В абзаце выравнивание должно быть по ширине страницы: «{p.text[:40]}...»"
                })

    # --- 8. Проверка наличия References ---
    has_references = False
    for p in paragraphs:
        if p.text.strip().lower() == "references":
            has_references = True
            # Проверяем, что заголовок References по центру и Times New Roman 14
            font_names = [r.font.name for r in p.runs if r.text.strip()]
            font_sizes = [r.font.size.pt if r.font.size else None for r in p.runs if r.text.strip()]
            if any(f != "Times New Roman" for f in font_names) or any(s != 14 for s in font_sizes):
                report.append({"status": "warn", "msg": "Заголовок 'References' должен быть Times New Roman 14 пт"})
            if p.alignment != 1:
                report.append({"status": "warn", "msg": "Заголовок 'References' должен быть по центру"})
            break
    if not has_references:
        report.append({"status": "error", "msg": "В тексте отсутствует раздел 'References'"})

    if not report:
        report.append({"status": "success", "msg": "Рукопись полностью соответствует инструкции!"})

    return report