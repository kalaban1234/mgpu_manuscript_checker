import docx
from io import BytesIO
import re

# --- Справочник обязательных разделов ---
REQUIRED_SECTIONS = [
    "УДК", "Аннотация", "Abstract", "Ключевые слова", "Key words", "Введение",
    "Материалы и методы", "Результаты исследования", "Заключение", "Список литературы", "References"
]

def get_text_snippet(text, length=50):
    return text[:length] + ("..." if len(text) > length else "")

def check_docx(file_bytes):
    doc = docx.Document(BytesIO(file_bytes))
    report = []
    section_found = {key: False for key in REQUIRED_SECTIONS}

    # --- Проверка объема текста (без пробелов и переносов) ---
    full_text = "\n".join([p.text for p in doc.paragraphs])
    char_count = len(full_text.replace("\n", ""))
    if not (20000 <= char_count <= 40000):
        report.append({
            "status": "error",
            "msg": f"Объем статьи {char_count} знаков (ожидалось 20 000–40 000)."
        })

    # --- Проверка структуры и обязательных разделов ---
    for section in REQUIRED_SECTIONS:
        found = False
        for p in doc.paragraphs:
            # поиск разделов — по началу абзаца (или всему тексту)
            if p.text.strip().lower().startswith(section.lower()):
                found = True
                section_found[section] = True
                break
        if not found:
            report.append({
                "status": "error",
                "msg": f"Не найден обязательный раздел: «{section}»"
            })

    # --- Проверка сведений об авторах сразу после УДК ---
    if len(doc.paragraphs) > 1:
        authors_text = doc.paragraphs[1].text.strip()
        if not re.search(r"[А-ЯA-Z][а-яa-zё]+ [А-ЯA-Z]\.[А-ЯA-Z]\.", authors_text):
            report.append({
                "status": "warn",
                "msg": f"Второй абзац после УДК не похож на строку с ФИО автора (пример: Иванов И.И., Петров П.П.): «{get_text_snippet(authors_text)}»"
            })

    # --- Проверка шрифта, кегля, интервала, красной строки ---
    for p in doc.paragraphs:
        para_text = p.text.strip()
        if not para_text:
            continue
        for run in p.runs:
            font_name = run.font.name
            font_size = run.font.size.pt if run.font.size else None
            if font_name != "Times New Roman" or font_size != 14:
                # Показываем начало строки вместо номера
                report.append({
                    "status": "warn",
                    "msg": f"Неверный шрифт или кегль в тексте: «{get_text_snippet(para_text)}» (ожидалось Times New Roman 14)"
                })
                break
        # Межстрочный интервал
        if p.paragraph_format.line_spacing and p.paragraph_format.line_spacing != 1.5:
            report.append({
                "status": "warn",
                "msg": f"Некорректный интервал между строками в тексте: «{get_text_snippet(para_text)}» (ожидалось 1.5)"
            })
        # Красная строка (отступ первой строки)
        if p.paragraph_format.first_line_indent:
            indent = p.paragraph_format.first_line_indent.pt
            if abs(indent - 18) > 1:  # 1.25 см = ~18pt
                report.append({
                    "status": "warn",
                    "msg": f"Некорректный отступ первой строки (красная строка) в тексте: «{get_text_snippet(para_text)}» (ожидалось 1.25 см)"
                })

    # --- Проверка ключевых слов (без точки, 3–15 штук) ---
    keywords_ru = None
    keywords_en = None
    for p in doc.paragraphs:
        if p.text.lower().startswith("ключевые слова"):
            keywords_ru = p.text
        if p.text.lower().startswith("key words") or p.text.lower().startswith("keywords"):
            keywords_en = p.text
    for label, kw in [('Ключевые слова', keywords_ru), ('Key words', keywords_en)]:
        if kw:
            klist = [w.strip() for w in kw.split(':', 1)[-1].split(',')]
            if not (3 <= len(klist) <= 15):
                report.append({
                    "status": "warn",
                    "msg": f"{label}: обнаружено {len(klist)} ключевых слов (ожидалось 3–15)"
                })
            if kw.strip().endswith('.'):
                report.append({
                    "status": "warn",
                    "msg": f"{label}: не должно быть точки в конце"
                })

    # --- Проверка списка литературы (наличие, алфавитность, оформление DOI/URL) ---
    sources_block = ""
    in_sources = False
    for p in doc.paragraphs:
        if p.text.lower().startswith("список литературы"):
            in_sources = True
            continue
        if in_sources:
            if not p.text.strip(): break
            sources_block += p.text + "\n"
    if sources_block:
        lines = [l for l in sources_block.split('\n') if l.strip()]
        # Простейшая алфавитность: сравниваем первую букву (это очень упрощённо!)
        for i in range(1, len(lines)):
            if lines[i][0] < lines[i-1][0]:
                report.append({
                    "status": "warn",
                    "msg": f"В списке литературы нарушен алфавитный порядок: «{get_text_snippet(lines[i-1])}» и «{get_text_snippet(lines[i])}»"
                })
                break
        # Наличие DOI/URL хотя бы в половине ссылок (для примера)
        doi_count = sum(1 for l in lines if ('doi:' in l.lower() or 'url:' in l.lower() or 'http' in l.lower()))
        if doi_count < len(lines) // 2:
            report.append({
                "status": "warn",
                "msg": f"В списке литературы слишком мало DOI/URL (найдено {doi_count} из {len(lines)})"
            })
    else:
        report.append({
            "status": "error",
            "msg": "Не найден блок «Список литературы»"
        })

    # --- Проверка блока References ---
    has_references = False
    for p in doc.paragraphs:
        if p.text.strip().lower().startswith("references"):
            has_references = True
    if not has_references:
        report.append({
            "status": "error",
            "msg": "Отсутствует обязательный раздел References (список литературы на английском)"
        })

    # --- Проверка ссылок в тексте (квадратные скобки) ---
    main_text = "\n".join([p.text for p in doc.paragraphs if not p.text.lower().startswith("список литературы")])
    if not re.search(r"\[\d+\]", main_text):
        report.append({
            "status": "warn",
            "msg": "В тексте не обнаружено ссылок в квадратных скобках (например, [1], [2])"
        })

    # --- Языковая чистота ---
    # Оценка примитивная: ищем встречается ли много английских слов в основном тексте
    ru_letters = re.findall(r'[А-Яа-яё]+', main_text)
    en_letters = re.findall(r'[A-Za-z]+', main_text)
    if len(en_letters) > len(ru_letters) // 2:
        report.append({
            "status": "warn",
            "msg": "В тексте слишком много английских слов (допускается только один язык для основного текста)"
        })

    # --- Оценка разделов, рисунков и таблиц ---
    # (очень базово: ищем слово «Рисунок» и «Таблица»)
    for p in doc.paragraphs:
        if "рисунок" in p.text.lower():
            # Требует оформления подписи 12 кеглем (не проверяем, просто предупреждаем)
            report.append({
                "status": "info",
                "msg": f"Проверьте подписи к рисункам (должно быть 12 кеглем): «{get_text_snippet(p.text)}»"
            })
        if "таблица" in p.text.lower():
            report.append({
                "status": "info",
                "msg": f"Проверьте оформление таблицы: «{get_text_snippet(p.text)}»"
            })

    # --- Итог ---
    if not report:
        report.append({"status": "success", "msg": "Рукопись полностью соответствует инструкции!"})

    return report
