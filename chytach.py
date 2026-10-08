"""
Читач заяв: бере .docx або .pdf, знаходить у ньому заяви, витягає дані для форми
«Електронного суду», зберігає їх у JSON і робить HTML-сторінку для перевірки юристом.

Усе працює локально, без інтернету.

Запуск:
    python chytach.py samples/zrazky.pdf
    python chytach.py "C:\\Шлях\\до\\заяви.docx" інша_заява.pdf

Результат з'являється в папці «rezultat»:
    <назва файлу>.json  — дані для заповнювача форми
    <назва файлу>.html  — сторінка для перевірки (відкрийте двічі клацнувши)
"""

import html
import json
import re
import sys
from datetime import date
from pathlib import Path

PAPKA_REZULTATU = Path(__file__).parent / "rezultat"

# ---------------------------------------------------------------------------
# 1. Читання файлів у список абзаців
# ---------------------------------------------------------------------------


class ProblemaZFailom(Exception):
    """Помилка, яку треба показати користувачу простими словами."""


def chytaty_pdf(shliakh):
    """Повертає список абзаців. Рядки PDF склеюються в абзаци за відступами між ними."""
    import pdfplumber

    abzatsy = []
    storinok_bez_tekstu = []
    with pdfplumber.open(shliakh) as pdf:
        for nomer, storinka in enumerate(pdf.pages, start=1):
            riadky = storinka.extract_text_lines()
            if len(storinka.chars) < 20:
                storinok_bez_tekstu.append(nomer)
                continue
            poperednii = None
            for riadok in riadky:
                tekst = riadok["text"].strip()
                if not tekst:
                    continue
                if poperednii is None:
                    # Перший рядок сторінки: продовжуємо абзац із попередньої сторінки,
                    # лише якщо там явно обірване речення.
                    novyi = (not abzatsy or re.search(r"[.:;!?»)/_]$", abzatsy[-1])
                             or RE_DATA_RIADOK.match(abzatsy[-1]) or pochatok_punktu(tekst)
                             or RE_POCHATOK.match(tekst) or tekst.startswith("___"))
                else:
                    vysota = riadok["bottom"] - riadok["top"]
                    promizhok = riadok["top"] - poperednii["bottom"]
                    novyi = promizhok > vysota * 0.9 or pochatok_punktu(tekst)
                if novyi:
                    abzatsy.append(tekst)
                else:
                    abzatsy[-1] = skleity(abzatsy[-1], tekst)
                poperednii = riadok
        vsogo = len(pdf.pages)

    if storinok_bez_tekstu and len(storinok_bez_tekstu) == vsogo:
        raise ProblemaZFailom(
            "У цьому PDF немає текстового шару — схоже, це скан або фото. "
            "Щоб прочитати його, потрібне розпізнавання тексту (OCR). "
            "Найпростіше: попросіть надіслати заяву у форматі Word, або збережіть PDF "
            "через програму з функцією OCR (наприклад, ABBYY FineReader) і спробуйте знову."
        )
    poperedzhennia = []
    if storinok_bez_tekstu:
        poperedzhennia.append(
            "Сторінки " + ", ".join(map(str, storinok_bez_tekstu))
            + " не містять тексту (можливо, скан). Їхній вміст НЕ прочитано."
        )
    return abzatsy, poperedzhennia


def chytaty_docx(shliakh):
    """Повертає список абзаців Word, включно з текстом у таблицях (шапку часто роблять таблицею)."""
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    dokument = docx.Document(shliakh)
    abzatsy = []

    def obiity(element, batko):
        for dytyna in element.iterchildren():
            teg = dytyna.tag.rsplit("}", 1)[-1]
            if teg == "p":
                tekst = Paragraph(dytyna, batko).text
                # Розрив рядка всередині абзацу (Shift+Enter) вважаємо окремим рядком.
                for chastyna in tekst.split("\n"):
                    if chastyna.strip():
                        abzatsy.append(normalizuvaty(chastyna))
            elif teg == "tbl":
                tablytsia = Table(dytyna, batko)
                for riadok in tablytsia.rows:
                    bachyly = set()
                    for komirka in riadok.cells:
                        if id(komirka._tc) in bachyly:  # об'єднані клітинки
                            continue
                        bachyly.add(id(komirka._tc))
                        obiity(komirka._tc, komirka)

    obiity(dokument.element.body, dokument)
    return abzatsy, []


def pochatok_punktu(tekst):
    return bool(re.match(r"^(●|•|-|–|—|\d{1,2}[.)]\s)", tekst))


def skleity(pershyi, druhyi):
    # Слово, перенесене через дефіс у кінці рядка: «зобов'я-» + «зання».
    if re.search(r"[а-яіїєґ]-$", pershyi) and re.match(r"^[а-яіїєґ]", druhyi):
        return pershyi[:-1] + druhyi
    return pershyi + " " + druhyi


def normalizuvaty(tekst):
    tekst = tekst.replace("\u00a0", " ").replace("\u2019", "'").replace("ʼ", "'")
    return re.sub(r"[ \t]+", " ", tekst).strip()


def chytaty_fail(shliakh):
    shliakh = Path(shliakh)
    if not shliakh.exists():
        raise ProblemaZFailom(f"Файл не знайдено: {shliakh}")
    rozshyrennia = shliakh.suffix.lower()
    if rozshyrennia == ".pdf":
        abzatsy, poperedzhennia = chytaty_pdf(shliakh)
    elif rozshyrennia == ".docx":
        abzatsy, poperedzhennia = chytaty_docx(shliakh)
    elif rozshyrennia == ".doc":
        raise ProblemaZFailom(
            "Старий формат .doc не підтримується. Відкрийте файл у Word і збережіть як .docx "
            "(Файл → Зберегти як → «Документ Word (*.docx)»)."
        )
    else:
        raise ProblemaZFailom(f"Непідтримуваний формат {rozshyrennia}. Потрібен .docx або .pdf.")
    return [normalizuvaty(a) for a in abzatsy if a.strip()], poperedzhennia


# ---------------------------------------------------------------------------
# 2. Розбиття файлу на окремі заяви
# ---------------------------------------------------------------------------

RE_POCHATOK = re.compile(r"^До\s+\S")
RE_ZAHOLOVOK = re.compile(r"^(ПОЗОВНА\s+ЗАЯВА|ЗАЯВА|ЗУСТРІЧНА\s+ПОЗОВНА\s+ЗАЯВА|ВІДЗИВ|АПЕЛЯЦІЙНА\s+СКАРГА)\b")
RE_PIDPYS = re.compile(r"_{3,}.*?/\s*([^/]+?)\s*/|/\s*([А-ЯІЇЄҐ]\.\s?[А-ЯІЇЄҐ]?\.?\s?[А-ЯІЇЄҐ][а-яіїєґ'\-]+)\s*/")


def tse_pochatok_zaiavy(abzatsy, i):
    if not RE_POCHATOK.match(abzatsy[i]):
        return False
    # «суд» має бути в цьому або наступному абзаці (назва суду буває розірвана).
    shmatok = " ".join(abzatsy[i:i + 2])[:200]
    return re.search(r"\bсуд", shmatok, re.IGNORECASE) is not None


def rozbyty_na_zaiavy(abzatsy):
    """Повертає список (абзаци заяви) і список відкинутого службового тексту."""
    pochatky = [i for i in range(len(abzatsy)) if tse_pochatok_zaiavy(abzatsy, i)]
    zaiavy, vidkynuto = [], []
    if not pochatky:
        return [], abzatsy
    if pochatky[0] > 0:
        vidkynuto.extend(abzatsy[:pochatky[0]])
    for n, start in enumerate(pochatky):
        kinets = pochatky[n + 1] if n + 1 < len(pochatky) else len(abzatsy)
        chastyna = abzatsy[start:kinets]
        # Усе після рядка підпису — службовий текст (наприклад «Ось ще два зразки…»).
        for j, a in enumerate(chastyna):
            if RE_PIDPYS.search(a):
                vidkynuto.extend(chastyna[j + 1:])
                chastyna = chastyna[:j + 1]
                break
        zaiavy.append(chastyna)
    return zaiavy, vidkynuto


# ---------------------------------------------------------------------------
# 3. Розбір однієї заяви
# ---------------------------------------------------------------------------

ROLI = [
    ("Заявник (Стягувач)", r"Заявник\s*\(\s*Стягувач\s*\)"),
    ("Стягувач", r"Стягувач"),
    ("Заявник", r"Заявник|Заявниця"),
    ("Боржник", r"Боржник"),
    ("Позивач", r"Позивач(?:ка)?"),
    ("Відповідач", r"Відповідач(?:ка)?"),
    ("Третя особа", r"Третя\s+особа[^:]{0,80}"),
    ("Представник", r"Представник[^:]{0,40}"),
]
RE_ROL = re.compile(r"(?:^|(?<=\s)|(?<=\.))(" + "|".join(f"(?:{v})" for _, v in ROLI) + r")\s*:")

ADRESA = (r"(?:Адреса[^:]{0,60}|Зареєстроване\s+місце\s+проживання|Місце\s+проживання[^:]{0,30}|"
          r"Юридична\s+адреса|Місцезнаходження|Місце\s+знаходження|Поштова\s+адреса)\s*:")
KINETS_ADRESY = r"(?=\s*(?:РНОКПП|ІПН|Код\s+ЄДРПОУ|ЄДРПОУ|Код|Тел|Засоби\s+зв|E-?mail|Ел\.\s*пошта|Відомості|Паспорт|$))"
RE_RNOKPP = re.compile(r"(?:РНОКПП|ІПН|реєстраційний номер облікової картки платника податків)[^:\d]{0,40}:?\s*(\d{10})\b", re.I)
RE_EDRPOU = re.compile(r"(?:ЄДРПОУ)[^:\d]{0,10}:?\s*(\d{8})\b", re.I)
RE_TELEFON = re.compile(r"(?:Тел\.?|Телефон|Засоби\s+зв'язку|Моб\.?)\s*:?\s*(\+?[\d][\d\s()\-]{8,}\d)", re.I)
RE_EMAIL = re.compile(r"[\w.+\-]+@[\w\-]+(?:\.[\w\-]+)+")
RE_KABINET = re.compile(r"електронного\s+кабінету\s*:?\s*([а-яіїєґ]+)", re.I)
RE_SUMA = r"(\d{1,3}(?:[\s\u00a0]\d{3})*(?:[.,]\d{1,2})?|\d+(?:[.,]\d{1,2})?)\s*(?:грн|гривень)"
RE_TSINA = re.compile(r"Ціна\s+(позову|вимоги)\s*:?\s*" + RE_SUMA, re.I)
RE_ZBIR = re.compile(r"Судовий\s+збір\s*:?\s*" + RE_SUMA, re.I)
RE_ZVILNENO = re.compile(r"звільнен\w*\s+від\s+сплати(?:\s+(?:судового\s+збору\s+)?на\s+підставі\s+(.+?))?\)?\.?\s*$", re.I)
OZNAKY_YUROSOBY = re.compile(
    r"\b(ТОВ|ПрАТ|ПАТ|АТ|ПП|ДП|КП|ТзОВ|Товариств|Акціонерн|Підприємств|Корпораці|Банк|Установ|Організаці|"
    r"Управлінн|Департамент|Рада\b|Фонд|Кооператив|Об'єднанн|Спілк|Асоціаці|Компані)", re.I)

MISIATSI = {
    "січня": 1, "лютого": 2, "березня": 3, "квітня": 4, "травня": 5, "червня": 6,
    "липня": 7, "серпня": 8, "вересня": 9, "жовтня": 10, "листопада": 11, "грудня": 12,
}


def suma_v_chyslo(tekst):
    """«1 211,20» → 1211.2"""
    return round(float(re.sub(r"[\s\u00a0]", "", tekst).replace(",", ".")), 2)


def sud_u_nazyvnomu(nazva):
    """«Шевченківського районного суду м. Києва» → «Шевченківський районний суд м. Києва».
    Змінюємо лише слова до «суду» включно; решта (м. Києва, міста Києва) лишається."""
    slova = nazva.split()
    rezultat = []
    for i, slovo in enumerate(slova):
        if slovo.lower() == "суду":
            rezultat.append(slovo[:-1])
            rezultat.extend(slova[i + 1:])
            return " ".join(rezultat)
        for kinets, zamina in (("ського", "ський"), ("цького", "цький"), ("зького", "зький"),
                               ("ього", "ій"), ("ого", "ий")):
            if slovo.endswith(kinets):
                slovo = slovo[: -len(kinets)] + zamina
                break
        rezultat.append(slovo)
    return nazva  # «суду» не знайшли — не чіпаємо


def rozibraty_sud(tekst):
    tekst = re.sub(r"^До\s+", "", tekst)
    m = re.search(r"(?:,\s*|\s+)(?=(?:вул\.|вулиця|просп\.|проспект|пл\.|площа|бульв?\.|бульвар|пров\.|"
                  r"провулок|майдан|шосе|узвіз|\d{5}\b))", tekst)
    if m:
        nazva, adresa = tekst[: m.start()], tekst[m.end():]
    else:
        nazva, adresa = tekst, ""
    nazva = nazva.strip(" ,")
    return {
        "nazva_v_zaiavi": nazva,
        "nazva": sud_u_nazyvnomu(nazva),
        "adresa": adresa.strip(" ,.") or None,
    }


def rozibraty_uchasnyka(rol, tekst):
    u = {
        "rol": rol,
        "typ": None,
        "nazva": None,
        "adresa": None,
        "rnokpp": None,
        "edrpou": None,
        "telefon": None,
        "email": None,
        "elektronnyi_kabinet": None,
        "tekst_u_zaiavi": tekst,
    }
    m_adr = re.search(ADRESA, tekst, re.I)
    if m_adr:
        u["nazva"] = tekst[: m_adr.start()].strip(" ,.")
        m_kin = re.search(KINETS_ADRESY, tekst[m_adr.end():], re.I)
        adresa = tekst[m_adr.end(): m_adr.end() + (m_kin.start() if m_kin else len(tekst))]
        u["adresa"] = adresa.strip(" ,.") or None
    else:
        m_kin = re.search(KINETS_ADRESY, tekst, re.I)
        u["nazva"] = tekst[: m_kin.start()].strip(" ,.") if m_kin else tekst.strip(" ,.")

    if m := RE_RNOKPP.search(tekst):
        u["rnokpp"] = m.group(1)
    if m := RE_EDRPOU.search(tekst):
        u["edrpou"] = m.group(1)
    if m := RE_TELEFON.search(tekst):
        u["telefon"] = re.sub(r"\s+", " ", m.group(1)).strip()
    if m := RE_EMAIL.search(tekst):
        u["email"] = m.group(0).rstrip(".")
    if m := RE_KABINET.search(tekst):
        slovo = m.group(1).lower()
        if slovo.startswith("відсутн") or slovo == "немає":
            u["elektronnyi_kabinet"] = "відсутній"
        elif slovo.startswith("невідом"):
            u["elektronnyi_kabinet"] = "невідомо"
        elif slovo.startswith("зареєстр") or slovo.startswith("наявн") or slovo == "є":
            u["elektronnyi_kabinet"] = "зареєстрований"
        else:
            u["elektronnyi_kabinet"] = slovo

    if u["edrpou"] or (u["nazva"] and OZNAKY_YUROSOBY.search(u["nazva"])):
        u["typ"] = "юридична особа"
    elif u["nazva"] and re.match(r"(ФОП|Фізична\s+особа\s*[-–]\s*підприєм)", u["nazva"], re.I):
        u["typ"] = "фізична особа-підприємець"
    elif u["rnokpp"] or u["nazva"]:
        u["typ"] = "фізична особа"
    return u


def rozibraty_shapku(shapka):
    """shapka — увесь текст від «До … суду» до заголовка, склеєний в один рядок."""
    rezultat = {"sud": None, "uchasnyky": [], "tsina": None, "sudovyi_zbir": None}

    # Де закінчуються учасники: «Ціна позову» / «Судовий збір».
    m_groshi = re.search(r"\b(Ціна\s+(?:позову|вимоги)|Судовий\s+збір)\b", shapka)
    chastyna_uchasnykiv = shapka[: m_groshi.start()] if m_groshi else shapka
    chastyna_groshei = shapka[m_groshi.start():] if m_groshi else ""

    roli = list(RE_ROL.finditer(chastyna_uchasnykiv))
    sud_tekst = chastyna_uchasnykiv[: roli[0].start()] if roli else chastyna_uchasnykiv
    rezultat["sud"] = rozibraty_sud(sud_tekst.strip())

    for n, m in enumerate(roli):
        kinets = roli[n + 1].start() if n + 1 < len(roli) else len(chastyna_uchasnykiv)
        nazva_roli = next(nazva for nazva, vzirets in ROLI if re.fullmatch(vzirets, m.group(1).strip()))
        tekst = chastyna_uchasnykiv[m.end(): kinets].strip()
        rezultat["uchasnyky"].append(rozibraty_uchasnyka(nazva_roli, tekst))

    if m := RE_TSINA.search(chastyna_groshei):
        rezultat["tsina"] = {
            "vyd": "ціна позову" if m.group(1).lower() == "позову" else "ціна вимоги",
            "suma": suma_v_chyslo(m.group(2)),
        }
    if m := RE_ZBIR.search(chastyna_groshei):
        zbir = {"suma": suma_v_chyslo(m.group(1)), "zvilnenyi": False, "pidstava_zvilnennia": None}
        reshta = chastyna_groshei[m.end():]
        if mz := RE_ZVILNENO.search(reshta):
            zbir["zvilnenyi"] = True
            if mz.group(1):
                zbir["pidstava_zvilnennia"] = mz.group(1).strip(" ).")
                # Повернути закриту лапку назви закону, якщо обрізали разом із дужкою.
                if zbir["pidstava_zvilnennia"].count("«") > zbir["pidstava_zvilnennia"].count("»"):
                    zbir["pidstava_zvilnennia"] += "»"
        rezultat["sudovyi_zbir"] = zbir
    elif re.search(r"звільнен\w*\s+від\s+сплати", shapka, re.I):
        rezultat["sudovyi_zbir"] = {"suma": None, "zvilnenyi": True, "pidstava_zvilnennia": None}
    return rezultat


def rozibraty_zaholovok(abzatsy, i):
    """Повертає (тип документа, предмет, індекс першого абзацу основного тексту)."""
    zaholovok = abzatsy[i]
    m = RE_ZAHOLOVOK.match(zaholovok)
    reshta = zaholovok[m.end():].strip()
    slovo = re.sub(r"\s+", " ", m.group(1)).upper()
    typ = None
    if slovo == "ПОЗОВНА ЗАЯВА":
        typ = "позовна заява"
    elif slovo == "ЗАЯВА" and re.match(r"про\s+видачу\s+судового\s+наказу", reshta, re.I):
        typ = "заява про видачу судового наказу"
        reshta = re.sub(r"^про\s+видачу\s+судового\s+наказу\s*", "", reshta, flags=re.I)
    else:
        typ = (slovo.capitalize() + " " + reshta).strip()

    nastupnyi = i + 1
    if not reshta and nastupnyi < len(abzatsy) and abzatsy[nastupnyi].lower().startswith("про "):
        reshta = abzatsy[nastupnyi]
        nastupnyi += 1
    return typ, (reshta or None), nastupnyi


def rozibraty_datu(tekst):
    if m := re.search(r"«?\s*(\d{1,2})\s*»?\s+([а-яіїєґ]+)\s+(\d{4})", tekst):
        misiats = MISIATSI.get(m.group(2).lower())
        if misiats:
            try:
                return date(int(m.group(3)), misiats, int(m.group(1))).isoformat()
            except ValueError:
                return None
    if m := re.search(r"\b(\d{1,2})\.(\d{1,2})\.(\d{4})\b", tekst):
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
        except ValueError:
            return None
    return None


RE_DODATKY = re.compile(r"^(Додатки|Додаток|Перелік\s+документів.*додаються|До\s+заяви\s+додаються)\s*:?\s*$", re.I)
RE_DATA_RIADOK = re.compile(r"^«?\s*\d{1,2}\s*»?\s+[а-яіїєґ]+\s+\d{4}\s*(?:р\.|року)?\s*$|^\d{1,2}\.\d{1,2}\.\d{4}\s*(?:р\.|року)?$", re.I)


def rozibraty_zaiavu(abzatsy):
    poperedzhennia = []

    i_zahol = next((i for i, a in enumerate(abzatsy) if RE_ZAHOLOVOK.match(a)), None)
    if i_zahol is None:
        poperedzhennia.append("Не знайдено заголовок («ПОЗОВНА ЗАЯВА» / «ЗАЯВА …»). Розбір неповний.")
        i_zahol = len(abzatsy)
    shapka = " ".join(abzatsy[:i_zahol])
    dani = rozibraty_shapku(shapka)

    typ, predmet, i_tekst = (None, None, i_zahol)
    if i_zahol < len(abzatsy):
        typ, predmet, i_tekst = rozibraty_zaholovok(abzatsy, i_zahol)

    # Підпис і дата
    pidpysant, data_zaiavy, i_kinets_tekstu = None, None, len(abzatsy)
    for j in range(len(abzatsy) - 1, max(i_tekst, len(abzatsy) - 6) - 1, -1):
        if m := RE_PIDPYS.search(abzatsy[j]):
            pidpysant = (m.group(1) or m.group(2)).strip()
            i_kinets_tekstu = j
            # Дата може стояти в тому ж рядку, що й підпис.
            data_zaiavy = rozibraty_datu(abzatsy[j][: m.start()]) or data_zaiavy
            break
    for j in range(i_kinets_tekstu - 1, max(i_tekst, i_kinets_tekstu - 4) - 1, -1):
        if RE_DATA_RIADOK.match(abzatsy[j]):
            data_zaiavy = rozibraty_datu(abzatsy[j])
            break

    # Додатки
    dodatky = []
    i_dod = next((j for j in range(i_tekst, i_kinets_tekstu) if RE_DODATKY.match(abzatsy[j])), None)
    if i_dod is not None:
        for a in abzatsy[i_dod + 1: i_kinets_tekstu]:
            if m := re.match(r"^(\d{1,2})[.)]\s*(.+)$", a):
                dodatky.append({"nomer": int(m.group(1)), "nazva": m.group(2).strip()})
            elif RE_DATA_RIADOK.match(a):
                break
            elif dodatky:
                dodatky[-1]["nazva"] += " " + a

    tekst_zaiavy = abzatsy[i_zahol:i_kinets_tekstu]
    rezultat = {
        "typ_dokumenta": typ,
        "predmet": predmet,
        "sud": dani["sud"],
        "uchasnyky": dani["uchasnyky"],
        "tsina": dani["tsina"],
        "sudovyi_zbir": dani["sudovyi_zbir"],
        "dodatky": dodatky,
        "data": data_zaiavy,
        "pidpysant": pidpysant,
        "shapka": abzatsy[:i_zahol],
        # Текст для вбудованого редактора Е-суду: від заголовка до дати, без підпису.
        # Чи потрібна в редакторі ще й шапка — з'ясуємо за записом codegen (крок 3).
        "tekst_dlia_redaktora": tekst_zaiavy,
        "poperedzhennia": poperedzhennia,
    }
    perevirka(rezultat)
    return rezultat


def perevirka(z):
    """Додає попередження про поля, яких не знайдено або які виглядають підозріло."""
    p = z["poperedzhennia"]
    if not z["sud"] or not z["sud"].get("nazva"):
        p.append("Не знайдено назву суду.")
    elif not z["sud"].get("adresa"):
        p.append("Не знайдено адресу суду.")
    if not z["typ_dokumenta"]:
        p.append("Не визначено тип документа.")
    if not z["predmet"]:
        p.append("Не знайдено предмет (рядок після заголовка).")
    if not z["uchasnyky"]:
        p.append("Не знайдено жодного учасника.")
    for u in z["uchasnyky"]:
        hto = f"{u['rol']} ({u['nazva'] or 'без назви'})"
        for pole, nazva in (("nazva", "ПІБ / назву"), ("adresa", "адресу"), ("telefon", "телефон"),
                            ("email", "e-mail"), ("elektronnyi_kabinet", "відомості про ел. кабінет")):
            if not u[pole]:
                p.append(f"{hto}: не знайдено {nazva}.")
        if u["typ"] == "юридична особа" and not u["edrpou"]:
            p.append(f"{hto}: юридична особа, але не знайдено код ЄДРПОУ.")
        if u["typ"] != "юридична особа" and not u["rnokpp"]:
            p.append(f"{hto}: не знайдено РНОКПП.")
    if z["typ_dokumenta"] and z["uchasnyky"]:
        roli = {u["rol"] for u in z["uchasnyky"]}
        if z["typ_dokumenta"] == "позовна заява" and not {"Позивач", "Відповідач"} <= roli:
            p.append("У позовній заяві очікуються і Позивач, і Відповідач.")
    if not z["tsina"]:
        p.append("Ціну позову / вимоги не знайдено (для немайнових позовів, напр. розірвання шлюбу, це нормально).")
    if not z["sudovyi_zbir"]:
        p.append("Не знайдено суму судового збору.")
    if not z["dodatky"]:
        p.append("Не знайдено список додатків.")
    if not z["data"]:
        p.append("Не знайдено дату заяви.")
    if not z["pidpysant"]:
        p.append("Не знайдено підписанта.")


# ---------------------------------------------------------------------------
# 4. Сторінка для перевірки
# ---------------------------------------------------------------------------

def e(znachennia):
    if znachennia is None or znachennia == "":
        return '<span class="nema">не знайдено</span>'
    return html.escape(str(znachennia))


def groshi(suma):
    if suma is None:
        return None
    return f"{suma:,.2f}".replace(",", " ").replace(".", ",") + " грн"


def zrobyty_html(nazva_failu, zaiavy, vidkynuto, poperedzhennia_failu):
    chastyny = []
    for n, z in enumerate(zaiavy, start=1):
        zbir = z["sudovyi_zbir"] or {}
        zbir_tekst = groshi(zbir.get("suma"))
        if zbir.get("zvilnenyi"):
            zbir_tekst = (zbir_tekst or "") + " — звільнений від сплати" + (
                f" ({zbir['pidstava_zvilnennia']})" if zbir.get("pidstava_zvilnennia") else "")
        riadky = [
            ("Суд", z["sud"]["nazva"] if z["sud"] else None),
            ("Адреса суду", z["sud"]["adresa"] if z["sud"] else None),
            ("Тип документа", z["typ_dokumenta"]),
            ("Предмет", z["predmet"]),
            ((z["tsina"] or {}).get("vyd", "ціна позову / вимоги").capitalize(), groshi((z["tsina"] or {}).get("suma"))),
            ("Судовий збір", zbir_tekst),
            ("Дата заяви", z["data"]),
            ("Підписант", z["pidpysant"]),
        ]
        tablytsia = "".join(f"<tr><th>{html.escape(k)}</th><td>{e(v)}</td></tr>" for k, v in riadky)

        uchasnyky = "".join(
            "<tr>" + "".join(f"<td>{e(u[k])}</td>" for k in
                             ("rol", "typ", "nazva", "adresa")) +
            f"<td>{e(u['rnokpp'] or u['edrpou'])}" + (" <small>(ЄДРПОУ)</small>" if u["edrpou"] else
                                                     " <small>(РНОКПП)</small>" if u["rnokpp"] else "") + "</td>" +
            "".join(f"<td>{e(u[k])}</td>" for k in ("telefon", "email", "elektronnyi_kabinet")) +
            "</tr>"
            for u in z["uchasnyky"]
        )
        dodatky = "".join(f"<li value=\"{d['nomer']}\">{e(d['nazva'])}</li>" for d in z["dodatky"]) \
            or '<li class="nema">не знайдено</li>'
        poperedzhennia = "".join(f"<li>{e(p)}</li>" for p in z["poperedzhennia"])
        blok_poperedzhen = (f'<div class="uvaga"><b>Перевірте:</b><ul>{poperedzhennia}</ul></div>'
                            if poperedzhennia else '<div class="ok">Усі поля знайдено.</div>')
        tekst = "".join(f"<p>{e(a)}</p>" for a in z["tekst_dlia_redaktora"])
        shapka = "".join(f"<p>{e(a)}</p>" for a in z["shapka"])
        chastyny.append(f"""
<section>
  <h2>Заява {n} з {len(zaiavy)}: {e(z['typ_dokumenta'])} {e(z['predmet'])}</h2>
  {blok_poperedzhen}
  <table class="polia">{tablytsia}</table>
  <h3>Учасники</h3>
  <div class="prokrutka"><table class="uch">
    <tr><th>Роль</th><th>Тип</th><th>ПІБ / назва</th><th>Адреса</th><th>РНОКПП / ЄДРПОУ</th>
        <th>Телефон</th><th>E-mail</th><th>Ел. кабінет</th></tr>
    {uchasnyky}
  </table></div>
  <h3>Додатки</h3>
  <ol>{dodatky}</ol>
  <details><summary>Шапка заяви (як у документі)</summary><div class="tekst">{shapka}</div></details>
  <details><summary>Текст для редактора Е-суду ({len(z['tekst_dlia_redaktora'])} абзаців)</summary>
    <div class="tekst">{tekst}</div></details>
</section>""")

    zahalni = "".join(f"<li>{e(p)}</li>" for p in poperedzhennia_failu)
    vidk = "".join(f"<p>{e(a)}</p>" for a in vidkynuto)
    return f"""<!doctype html>
<html lang="uk"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Перевірка: {html.escape(nazva_failu)}</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", Arial, sans-serif; margin: 0 auto; max-width: 1100px;
         padding: 16px; color: #1d2433; background: #f6f7f9; line-height: 1.45; }}
  h1 {{ font-size: 22px; }} h2 {{ font-size: 18px; margin-top: 0; }} h3 {{ font-size: 15px; margin-bottom: 6px; }}
  section {{ background: #fff; border: 1px solid #dde1e7; border-radius: 8px; padding: 16px 20px; margin: 18px 0; }}
  table {{ border-collapse: collapse; width: 100%; font-size: 14px; }}
  th, td {{ border: 1px solid #dde1e7; padding: 6px 8px; text-align: left; vertical-align: top; }}
  .polia th {{ width: 180px; background: #f0f2f5; font-weight: 600; }}
  .uch th {{ background: #f0f2f5; }}
  .prokrutka {{ overflow-x: auto; }}
  .nema {{ background: #fff1c2; color: #7a5200; padding: 0 4px; border-radius: 3px; font-style: italic; }}
  .uvaga {{ background: #fff8e1; border: 1px solid #f0c94a; border-radius: 6px; padding: 8px 12px; margin-bottom: 12px; }}
  .uvaga ul {{ margin: 4px 0 0; }}
  .ok {{ background: #e8f6ec; border: 1px solid #8fd1a1; border-radius: 6px; padding: 8px 12px; margin-bottom: 12px; }}
  details {{ margin-top: 10px; }} summary {{ cursor: pointer; font-weight: 600; }}
  .tekst {{ background: #fafbfc; border: 1px solid #e5e8ec; padding: 8px 14px; font-size: 14px; }}
  small {{ color: #667; }}
</style></head><body>
<h1>Перевірка даних: {html.escape(nazva_failu)}</h1>
<p>Знайдено заяв: <b>{len(zaiavy)}</b>. Жовтим позначено поля, які програма не знайшла.
Будь ласка, звірте дані з оригіналом перед створенням чернетки.</p>
{f'<div class="uvaga"><b>Увага щодо файлу:</b><ul>{zahalni}</ul></div>' if zahalni else ''}
{''.join(chastyny)}
{f'<details><summary>Відкинутий службовий текст (не належить до жодної заяви)</summary><div class="tekst">{vidk}</div></details>' if vidk else ''}
</body></html>"""


# ---------------------------------------------------------------------------
# 5. Запуск
# ---------------------------------------------------------------------------

def obrobyty(shliakh):
    abzatsy, poperedzhennia_failu = chytaty_fail(shliakh)
    zaiavy_abzatsy, vidkynuto = rozbyty_na_zaiavy(abzatsy)
    if not zaiavy_abzatsy:
        raise ProblemaZFailom("Не знайдено жодної заяви. Кожна заява має починатися з рядка «До … суду».")
    zaiavy = [rozibraty_zaiavu(a) for a in zaiavy_abzatsy]
    return {
        "fail": str(Path(shliakh).name),
        "poperedzhennia_failu": poperedzhennia_failu,
        "zaiavy": zaiavy,
        "vidkynutyi_tekst": vidkynuto,
    }


def main(argumenty):
    if not argumenty:
        print("Вкажіть файл із заявою, наприклад:\n    python chytach.py samples/zrazky.pdf")
        return 1
    PAPKA_REZULTATU.mkdir(exist_ok=True)
    kod = 0
    for shliakh in argumenty:
        print(f"\n=== {shliakh}")
        try:
            dani = obrobyty(shliakh)
        except ProblemaZFailom as pomylka:
            print(f"ПОМИЛКА: {pomylka}")
            kod = 1
            continue
        stem = Path(shliakh).stem
        json_shliakh = PAPKA_REZULTATU / f"{stem}.json"
        html_shliakh = PAPKA_REZULTATU / f"{stem}.html"
        json_shliakh.write_text(json.dumps(dani, ensure_ascii=False, indent=2), encoding="utf-8")
        html_shliakh.write_text(
            zrobyty_html(dani["fail"], dani["zaiavy"], dani["vidkynutyi_tekst"], dani["poperedzhennia_failu"]),
            encoding="utf-8")
        for p in dani["poperedzhennia_failu"]:
            print(f"  УВАГА: {p}")
        for n, z in enumerate(dani["zaiavy"], start=1):
            stan = f"{len(z['poperedzhennia'])} зауваж." if z["poperedzhennia"] else "усе знайдено"
            print(f"  {n}. {z['typ_dokumenta']} {z['predmet'] or ''} — {z['sud']['nazva'] if z['sud'] else '?'} [{stan}]")
        print(f"  Дані:      {json_shliakh}")
        print(f"  Перевірка: {html_shliakh}")
    return kod


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")  # щоб українські літери не ламались у консолі Windows
    sys.exit(main(sys.argv[1:]))
