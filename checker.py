def is_keywords_ru(text):
    return bool(re.match(r"^ключ[её]вы[её][\s\-]*слова", text, re.IGNORECASE))

def is_keywords_en(text):
    return bool(re.match(r"^key[\s\-]*words", text, re.IGNORECASE))
import docx
from io import BytesIO
import re
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Паттерны стоп-заголовков для аннотаций/ключевых слов
STOP_HEADER_PATTERNS = [
    r"ключ[её]в[ыеё]+[\s\-]*слова?",
    r"key[\s\-]*words?",
    r"abstract",
    r"введение",
    r"материал[ыа]+ и методы",
    r"результат[ыа]+",
    r"заключение",
    r"список (источников|литературы)",
    r"сведения об авторах",
    r"references?"
]
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
                report.append({"status": "warn", "msg": "УДК должен быть Times New Roman 14 пт, не жирный", "section": "Оформление статьи"})
            if udk_p.alignment not in [None, 0]:
                report.append({"status": "warn", "msg": "УДК должен быть по левому краю", "section": "Оформление статьи"})
            break
    else:
        report.append({"status": "error", "msg": "В первом абзаце не найден УДК", "section": "Оформление статьи"})
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
                report.append({"status": "warn", "msg": f"ФИО автора '{text}' должен быть Times New Roman", "section": "Оформление статьи"})
            if any(s != 14 for s in font_sizes):
                report.append({"status": "warn", "msg": f"ФИО автора '{text}' должен быть 14 пт", "section": "Оформление статьи"})
            if not all(bolds):
                report.append({"status": "warn", "msg": f"ФИО автора '{text}' должен быть полужирным (bold)", "section": "Оформление статьи"})
            if p.alignment not in [None, 0]:
                report.append({"status": "warn", "msg": f"ФИО автора '{text}' должен быть по левому краю", "section": "Оформление статьи"})
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
                    {"status": "warn", "msg": "Название статьи должно быть Times New Roman 14 пт, полужирным", "section": "Оформление статьи"})
            break
        else:
            if p.alignment != 1:
                report.append({"status": "warn", "msg": "Название статьи должно быть по центру", "section": "Оформление статьи"})
            if not all(bolds):
                report.append({"status": "warn", "msg": "Название статьи должно быть полужирным (bold)", "section": "Оформление статьи"})
            if any(f != "Times New Roman" for f in font_names) or any(s != 14 for s in font_sizes):
                report.append({"status": "warn", "msg": "Название статьи должно быть Times New Roman 14 пт", "section": "Оформление статьи"})
            break
    if not found_title:
        report.append({"status": "error",
                       "msg": "Название статьи не найдено или не соответствует требованиям (по центру, полужирное, Times New Roman 14)",
                       "section": "Оформление статьи"})

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
                "msg": f"В абзаце найден неверный размер шрифта ({wrong_size} пт): «{p.text[:40]}...». Ожидалось 14 пт.",
                "section": "Оформление статьи"
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
                "msg": f"В абзаце выравнивание должно быть по ширине страницы: «{p.text[:40]}...»",
                "section": "Оформление статьи"
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
                    {"status": "warn", "msg": f"Подпись к рисунку '{p.text.strip()[:30]}...' должна быть 12 кеглем", "section": "Оформление статьи"})

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
        report.append({"status": "warn", "msg": f"Нет ссылок на рисунки {missed_str} в тексте", "section": "Оформление статьи"})

    missed_in_captions = drawing_refs - drawing_captions
    if missed_in_captions:
        missed_str = ", ".join(sorted(missed_in_captions))
        report.append(
            {"status": "warn", "msg": f"Есть ссылки на рисунки {missed_str} в тексте, но нет соответствующих подписей", "section": "Оформление статьи"})

    # --- 5. Проверка объема статьи (только до списка литературы) ---
    main_text = ""
    for p in paragraphs:
        if p.text.lower().startswith("список литературы") or p.text.lower().startswith("список источников"):
            break
        main_text += p.text + "\n"
    char_count = len(main_text.replace("\n", ""))
    if not (20000 <= char_count <= 40000):
        report.append({"status": "error",
                       "msg": f"Объем статьи {char_count} знаков (без списка литературы; ожидалось 20 000–40 000).",
                       "section": "Оформление статьи"})

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
            report.append({"status": "warn", "msg": f"Заголовок '{biblio_title}' должен быть Times New Roman 14 пт", "section": "Список источников"})
        if not all(b for b in biblio_p.runs if b.text.strip()):
            report.append({"status": "warn", "msg": f"Заголовок '{biblio_title}' должен быть полужирным (bold)", "section": "Список источников"})
        if biblio_p.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            report.append({"status": "warn", "msg": f"Заголовок '{biblio_title}' должен быть по центру", "section": "Список источников"})
        # Название должно быть строго "Список источников"
        if biblio_title.lower() != "список источников":
            report.append({"status": "error", "msg": "Название раздела должно быть строго 'Список источников'", "section": "Список источников"})
    else:
        report.append(
            {"status": "error", "msg": "В тексте отсутствует заголовок 'Список источников' или 'Список литературы'", "section": "Список источников"})

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
                            "msg": f"Подпись к рисунку в списке должна быть 12 пт: «{run.text[:40]}...»",
                            "section": "Список источников"
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
                    "msg": f"В абзаце найден неверный размер шрифта ({wrong_size} пт): «{p.text[:40]}...». Ожидалось 14 пт.",
                    "section": "Список источников"
                })
            elif not has_size:
                report.append({
                    "status": "warn",
                    "msg": f"В абзаце не удалось определить размер шрифта: «{p.text[:40]}...». Ожидалось 14 пт.",
                    "section": "Список источников"
                })
            # Проверка выравнивания абзаца
            if p.alignment != WD_ALIGN_PARAGRAPH.JUSTIFY:
                report.append({
                    "status": "warn",
                    "msg": f"В абзаце выравнивание должно быть по ширине страницы: «{p.text[:40]}...»",
                    "section": "Список источников"
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
                report.append({"status": "warn", "msg": "Заголовок 'References' должен быть Times New Roman 14 пт", "section": "Список источников"})
            if p.alignment != 1:
                report.append({"status": "warn", "msg": "Заголовок 'References' должен быть по центру", "section": "Список источников"})
            break
    if not has_references:
        report.append({"status": "error", "msg": "В тексте отсутствует раздел 'References'", "section": "Список источников"})

    # --- Новый блок: Проверка структуры статьи (обязательных разделов) ---
    def normalize_section(s):
        # Привести к нижнему регистру, убрать дефисы и лишние пробелы
        return re.sub(r"[\s\-]+", " ", s.lower().strip())

    # ОБНОВЛЁННЫЙ список: убран "введение"
    EXPECTED_SECTIONS = [
        "удк",
        "сведения об авторах",
        "аннотация",
        "abstract",
        "ключевые слова",
        "keywords",
        "материалы и методы",
        "результаты исследования",
        "заключение",
        "список источников"
    ]

    found_sections_map = {}  # ключ — нормализованный раздел, значение — как реально написано
    for p in paragraphs:
        text = normalize_section(p.text)
        for sec in EXPECTED_SECTIONS:
            if text.startswith(normalize_section(sec)):
                found_sections_map[sec] = p.text.strip()

    for sec in EXPECTED_SECTIONS:
        if sec == "сведения об авторах":
            continue  # проверяется отдельно
        if sec not in found_sections_map:
            # Пишем специальную ошибку с указанием на неправильное написание
            report.append({
                "status": "warn",
                "msg": f"В тексте отсутствует раздел '{sec.title()}' или он написан неверно",
                "section": "Структура"
            })
    # УБИРАЕМ ДУБЛЬ: больше не пишем про строгое соответствие "Список источников" в предыдущем блоке!

    # --- Новый блок: Корректная проверка аннотаций и ключевых слов как диапазонов ---

    def get_block_text(paragraphs, start_idx, stop_headers):
        """
        Возвращает все строки после start_idx, пока не встретится абзац,
        начинающийся на один из stop_headers.
        """
        block = []
        for p in paragraphs[start_idx + 1:]:
            text = p.text.strip()
            if not text:
                continue
            # Если абзац начинается на любой заголовок — стоп!
            if any(text.lower().startswith(h) for h in stop_headers):
                break
            block.append(text)
        return " ".join(block)

    stop_headers = [
        "ключевые слова", "keywords", "abstract", "введение", "материалы и методы",
        "результаты", "заключение", "список источников", "сведения об авторах", "references"
    ]

    annotation_ru = ""
    annotation_en = ""
    keywords_ru_block = ""
    keywords_en_block = ""
    for i, p in enumerate(paragraphs):
        text = p.text.strip().lower()
        if text.startswith("аннотация"):
            annotation_ru = get_block_text(paragraphs, i, stop_headers)
        if text.startswith("abstract"):
            annotation_en = get_block_text(paragraphs, i, stop_headers)
        if is_keywords_ru(text):
            keywords_ru_block = get_block_text(paragraphs, i, stop_headers)
            # если ключевые слова в одной строке — парсим только её
            if not keywords_ru_block:
                keywords_ru_block = p.text.split(":", 1)[-1] if ":" in p.text else ""
        if is_keywords_en(text):
            keywords_en_block = get_block_text(paragraphs, i, stop_headers)
            if not keywords_en_block:
                keywords_en_block = p.text.split(":", 1)[-1] if ":" in p.text else ""

    def extract_section_text(paragraphs, idx, header):
        """
        Ищет текст после header: в одной строке или сразу в следующем абзаце.
        """
        p = paragraphs[idx]
        text = p.text.strip()
        if text.lower().startswith(header):
            colon_idx = text.find(":")
            # В одной строке
            if colon_idx != -1 and colon_idx + 1 < len(text):
                return text[colon_idx + 1:].strip()
            # В следующем абзаце
            if idx + 1 < len(paragraphs):
                next_text = paragraphs[idx + 1].text.strip()
                # Не следующий заголовок
                if next_text and not any(next_text.lower().startswith(h) for h in stop_headers):
                    return next_text
        return ""

    # Удалено: старое извлечение keywords_ru и keywords_en, используем только keywords_ru_block и keywords_en_block

    def count_words(text):
        return len(re.findall(r"\w+", text))

    def count_keywords(text):
        return len([x.strip() for x in text.replace('\n', ',').replace(';', ',').split(",") if x.strip()])

def extract_annotation_block(paragraphs, header, stop_header_patterns):
    start = -1
    header_pattern = re.compile(rf"^{header}[\s\.\:\-]*", re.IGNORECASE)
    for i, p in enumerate(paragraphs):
        text = p.text.strip().lower()
        if header_pattern.match(text):
            start = i
            break
    if start == -1:
        return ""
    block = []
    # Первая строка может содержать часть аннотации сразу после "Аннотация."
    first_line = paragraphs[start].text.strip()
    after_header = re.sub(rf"^{header}[\.\:\-\s]*", "", first_line, flags=re.IGNORECASE).strip()
    if after_header:
        block.append(after_header)
    # Собираем все абзацы до первого стоп-заголовка
    for p in paragraphs[start + 1:]:
        txt = p.text.strip()
        # Если абзац начинается на любой стоп-заголовок — выходим
        if txt and any(re.match(pattern, txt.lower()) for pattern in stop_header_patterns):
            break
        block.append(txt)
    # Склеиваем, убирая пустые строки
    return " ".join([x for x in block if x])

    # Используем паттерны STOP_HEADER_PATTERNS для поиска аннотации
    annotation_ru = extract_annotation_block(paragraphs, "аннотация", STOP_HEADER_PATTERNS)
    print("DEBUG annotation_ru:", annotation_ru)
    if annotation_ru:
        word_count_ru = len(re.findall(r"\w+", annotation_ru))
        if word_count_ru > 250:
            report.append({
                "status": "warn",
                "msg": f"Аннотация на русском превышает 250 слов: {word_count_ru} слов",
                "section": "Аннотация"
            })
    else:
        report.append({
            "status": "warn",
            "msg": "В тексте отсутствует аннотация на русском языке",
            "section": "Аннотация"
        })

    if annotation_en:
        word_count_en = count_words(annotation_en)
        if word_count_en > 250:
            report.append({
                "status": "warn",
                "msg": f"Abstract превышает 250 слов: {word_count_en} слов",
                "section": "Аннотация"
            })
    else:
        report.append({
            "status": "warn",
            "msg": "В тексте отсутствует аннотация на английском языке (Abstract)",
            "section": "Аннотация"
        })

    # Проверка ключевых слов
    if keywords_ru_block:
        num = count_keywords(keywords_ru_block)
        if num < 3 or num > 15:
            report.append({
                "status": "warn",
                "msg": f"В русском языке количество ключевых слов вне диапазона 3–15 (найдено: {num})",
                "section": "Ключевые слова"
            })
    else:
        report.append({
            "status": "warn",
            "msg": f"В тексте отсутствует блок ключевых слов на русском языке",
            "section": "Ключевые слова"
        })

    if keywords_en_block:
        num = count_keywords(keywords_en_block)
        if num < 3 or num > 15:
            report.append({
                "status": "warn",
                "msg": f"В английском языке количество ключевых слов вне диапазона 3–15 (найдено: {num})",
                "section": "Ключевые слова"
            })
    else:
        report.append({
            "status": "warn",
            "msg": f"В тексте отсутствует блок ключевых слов на английском языке",
            "section": "Ключевые слова"
        })
    return report

def group_report(report):
    groups = {}
    for err in report:
        section = err.get('section', 'Оформление статьи')
        groups.setdefault(section, []).append(err)
    order = [
        "Оформление статьи",
        "Список источников",
        "Структура",
        "Аннотация",
        "Ключевые слова",
        "Прочее"
    ]
    result = []
    for sec in order:
        if sec in groups:
            result.append((sec, groups[sec]))
    # Добавить любые другие секции, если вдруг они новые
    for sec in groups:
        if sec not in order:
            result.append((sec, groups[sec]))
    return result