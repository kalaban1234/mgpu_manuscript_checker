import docx
from io import BytesIO
import re

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

    # 3a. Проверка всего основного текста на размер шрифта 14 pt
    main_end = len(paragraphs)
    for i, p in enumerate(paragraphs):
        if p.text.lower().startswith("список источников") or p.text.lower().startswith("список литературы"):
            main_end = i
            break

    for p in paragraphs[:main_end]:
        wrong_size = None
        for run in p.runs:
            if run.text.strip() and (run.font.size and run.font.size.pt != 14):
                wrong_size = run.font.size.pt
                break
        if wrong_size:
            report.append({
                "status": "warn",
                "msg": f"В абзаце найден неверный размер шрифта ({wrong_size} пт): «{p.text[:40]}...». Ожидалось 14 пт."
            })

    # 4. Подписи и ссылки на рисунки (строго: только если есть номер)
    drawing_captions = set()
    for p in paragraphs:
        match = re.match(r"(Рисунок|рисунок|Рис\.|рис\.)\s*(\d+)", p.text.strip(), re.IGNORECASE)
        if match:
            drawing_captions.add(match.group(2))
            font_sizes = [r.font.size.pt if r.font.size else None for r in p.runs if r.text.strip()]
            if any(s != 12 for s in font_sizes):
                report.append(
                    {"status": "warn", "msg": f"Подпись к рисунку '{p.text.strip()[:30]}...' должна быть 12 кеглем"})

    drawing_refs = set()
    for p in paragraphs:
        if re.match(r"(Рисунок|рисунок|Рис\.|рис\.)\s*\d+", p.text.strip(), re.IGNORECASE):
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

    # 5. Проверка объема статьи (только до списка литературы)
    main_text = ""
    for p in paragraphs:
        if p.text.lower().startswith("список литературы") or p.text.lower().startswith("список источников"):
            break
        main_text += p.text + "\n"
    char_count = len(main_text.replace("\n", ""))
    if not (20000 <= char_count <= 40000):
        report.append({"status": "error", "msg": f"Объем статьи {char_count} знаков (без списка литературы; ожидалось 20 000–40 000)."})

    # Проверка правильного заголовка библиографического списка (раздела)
    has_sources_title = False
    for p in paragraphs:
        title = p.text.strip()
        # Требуется строго "Список источников"
        if title.lower() == "список источников":
            has_sources_title = True
            # Проверка шрифта и размера
            font_names = [r.font.name for r in p.runs if r.text.strip()]
            font_sizes = [r.font.size.pt if r.font.size else None for r in p.runs if r.text.strip()]
            if any(f != "Times New Roman" for f in font_names) or any(s != 14 for s in font_sizes):
                report.append(
                    {"status": "warn", "msg": "Заголовок 'Список источников' должен быть Times New Roman 14 пт"})
            if p.alignment not in [3, None]:  # По ширине страницы
                report.append({"status": "warn", "msg": "Заголовок 'Список источников' должен быть по ширине страницы"})
            break
        # Если встречено другое название
        if title.lower().startswith("список литер") or title.lower().startswith("список литературы"):
            report.append({"status": "error", "msg": "Название раздела должно быть строго 'Список источников'"})

    if not has_sources_title:
        report.append({"status": "error", "msg": "В тексте отсутствует заголовок 'Список источников'"})
    # 6. Список литературы — только блок между "Список литературы" и "Сведения об авторах"
    in_sources = False
    for p in paragraphs:
        if p.text.strip().lower().startswith(("список литературы", "список источников")):
            in_sources = True
            continue
        if in_sources:
            if p.text.strip().lower().startswith("сведения об авторах"):
                break
            for run in p.runs:
                if run.text.strip():
                    if run.font.name != "Times New Roman" or (run.font.size and run.font.size.pt != 14):
                        report.append({"status": "warn", "msg": f"В списке источников неверный шрифт или кегль: «{p.text[:40]}...». Ожидалось Times New Roman 14 пт"})
            if p.alignment not in [3, None]:
                report.append({"status": "warn", "msg": f"В списке источников выравнивание должно быть по ширине страницы"})

    if not report:
        report.append({"status": "success", "msg": "Рукопись полностью соответствует инструкции!"})

    return report