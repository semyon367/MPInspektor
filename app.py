import streamlit as st
import openpyxl
import re
import io
from datetime import datetime, date
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment

st.set_page_config(
    page_title="Анализ МП Инспектор",
    page_icon="📊",
    layout="centered"
)

# ================== Конфигурация ==================
SHEET_NAME = "Детализация МП Инспектор"

SUBJEKT_TO_DISTRICT = {
    "УНДиПР ГУ МЧС России по Чукотскому АО": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Забайкальскому краю": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Сахалинской области": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Приморскому краю": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Еврейской АО": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Камчатскому краю": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Республике Саха (Якутия)": "Дальневосточный ФО",
    "УНДиПР Главного управления МЧС России по Республике Бурятия": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Хабаровскому краю": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Магаданской области": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Амурской области": "Дальневосточный ФО",
    "УНДиПР ГУ МЧС России по Донецкой Народной Республике": "Новые регионы",
    "УНДиПР ГУ МЧС России по Запорожской области": "Новые регионы",
    "УНДиПР ГУ МЧС России по Херсонской области": "Новые регионы",
    "УНДиПР ГУ МЧС России по Луганской Народной Республике": "Новые регионы",
    "УНДиПР ГУ МЧС России по Пензенской области": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Оренбургской области": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Ульяновской области": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Республике Башкортостан": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Удмуртской Республике": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Самарской области": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Нижегородской области": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Кировской области": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Пермскому краю": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Саратовской области": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Республике Мордовия": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Чувашской Республике - Чувашии": "Приволжский ФО",
    "УНДиПР Главного управления МЧС России по Республике Марий Эл": "Приволжский ФО",
    "УНДиПР ГУ МЧС России по Республике Татарстан": "Приволжский ФО",
    "УНДПР Главного управления МЧС России по г. Санкт-Петербургу": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Ленинградской области": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Калининградской области": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Псковской области": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Республике Коми": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Архангельской области": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Вологодской области": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Новгородской области": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Республике Карелия": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Мурманской области": "Северо-Западный ФО",
    "ОНДиПР ГУ МЧС России по Ненецкому автономному округу": "Северо-Западный ФО",
    "УНДиПР ГУ МЧС России по Кабардино-Балкарской Республике": "Северо-Кавказский ФО",
    "УНДиПР ГУ МЧС России по Республике Северная Осетия - Алания": "Северо-Кавказский ФО",
    "УНДиПР ГУ МЧС России по Республике Дагестан": "Северо-Кавказский ФО",
    "УНДиПР ГУ МЧС России по Карачаево-Черкесской Республике": "Северо-Кавказский ФО",
    "УНДиПР ГУ МЧС России по Ставропольскому краю": "Северо-Кавказский ФО",
    "УНДиПР ГУ МЧС России по Республике Ингушетия": "Северо-Кавказский ФО",
    "УНДиПР ГУ МЧС России по Чеченской Республике": "Северо-Кавказский ФО",
    "УНДиПР ГУ МЧС России по Республике Тыва": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Новосибирской области": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Кемеровской области - Кузбассу": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Красноярскому краю": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Томской области": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Республике Алтай": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Омской области": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Республике Хакасия": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Алтайскому краю": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Иркутской области": "Сибирский ФО",
    "УНДиПР ГУ МЧС России по Курганской области": "Уральский ФО",
    "УНДиПР ГУ МЧС России по Свердловской области": "Уральский ФО",
    "УНДиПР ГУ МЧС России по Челябинской области": "Уральский ФО",
    "УНДиПР ГУ МЧС России по Ямало-Ненецкому АО": "Уральский ФО",
    "УНДиПР ГУ МЧС России по Ханты-Мансийскому АО - Югре": "Уральский ФО",
    "УНДиПР ГУ МЧС России по Тюменской области": "Уральский ФО",
    "УНДиПР ГУ МЧС России по Тверской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Курской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по г. Москве": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Московской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Владимирской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Тамбовской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Тульской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Липецкой области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Рязанской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Костромской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Ярославской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Ивановской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Воронежской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Калужской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Белгородской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Брянской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по Смоленской области": "Центральный ФО",
    "УНДПР ГУ МЧС России по Орловской области": "Центральный ФО",
    "УНДиПР ГУ МЧС России по г. Севастополю": "Южный ФО",
    "УНДиПР ГУ МЧС России по Волгоградской области": "Южный ФО",
    "УНДиПР ГУ МЧС России по Ростовской области": "Южный ФО",
    "УНДиПР ГУ МЧС России по Республике Адыгея": "Южный ФО",
    "УНДиПР ГУ МЧС России по Астраханской области": "Южный ФО",
    "УНДиПР ГУ МЧС России по Республике Крым": "Южный ФО",
    "УНДиПР ГУ МЧС России по Республике Калмыкия": "Южный ФО",
    "УНДиПР ГУ МЧС России по Краснодарскому краю": "Южный ФО",
}

COLUMN_KEYWORDS = {
    'subjekt': ['субъект рф'],
    'podrazdelenie': ['подразделение'],
    'vid_nadzora': ['вид надзора'],
    'nom_knm': ['номер кнм'],
    'vid': ['вид'],
    'status': ['статус кнм'],
    'narusheniya': ['нарушения выявлены'],
    'proverka_ogv': ['проверка огв/омсу'],
    'knd': ['кнд'],
    'ssylki': ['ссылки на файлы'],
    'date_act': ['дата составления акта о результате кнм', 'дата составления акта'],
    'tip_prof_vizita': ['тип проф. визита', 'тип профилактического визита'],
    's_vks': ['с вкс', 'вкс'],
}


# ================== Вспомогательные функции ==================

def normalize_str(s):
    if s is None:
        return ""
    return re.sub(r'\s+', ' ', str(s).strip()).lower()

def find_column_index(headers, possible_names):
    headers_norm = [normalize_str(h) for h in headers]
    possible_norm = [normalize_str(name) for name in possible_names]
    for idx, norm in enumerate(headers_norm):
        if norm in possible_norm:
            return idx
    for idx, norm in enumerate(headers_norm):
        for pname in possible_norm:
            if pname in norm:
                return idx
    return None

def parse_date(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        value = value.strip()
        for fmt in ('%d.%m.%Y', '%d/%m/%Y', '%d-%m-%Y'):
            try:
                return datetime.strptime(value, fmt).date()
            except ValueError:
                continue
    return None

def load_data(file_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"Лист '{SHEET_NAME}' не найден в файле. Доступные листы: {', '.join(wb.sheetnames)}")
    ws = wb[SHEET_NAME]
    headers = [cell.value if cell.value else "" for cell in ws[1]]
    col_indices = {}
    missing = []
    for key, possible_names in COLUMN_KEYWORDS.items():
        idx = find_column_index(headers, possible_names)
        if idx is None:
            missing.append(f"'{key}' (искали: {possible_names})")
        else:
            col_indices[key] = idx
    if missing:
        raise ValueError(f"Не найдены столбцы: {'; '.join(missing)}")
    data = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in row):
            continue
        data.append(row)
    return data, col_indices, headers

def filter_by_date(data, col_idx, date_from, date_to):
    date_col = col_idx['date_act']
    filtered, skipped = [], 0
    for row in data:
        parsed = parse_date(row[date_col])
        if parsed is None:
            skipped += 1
            continue
        if date_from <= parsed <= date_to:
            filtered.append(row)
        else:
            skipped += 1
    return filtered, skipped

def calculate_all_metrics(data, col_idx):
    podr_col = col_idx['podrazdelenie']
    knm_col = col_idx['nom_knm']
    vid_col = col_idx['vid']
    status_col = col_idx['status']
    proverka_col = col_idx['proverka_ogv']
    vid_nadzora_col = col_idx['vid_nadzora']
    knd_col = col_idx['knd']
    nar_col = col_idx['narusheniya']
    vks_col = col_idx['s_vks']
    ssylki_col = col_idx['ssylki']

    allowed_vids = {"", "выездная проверка", "рейдовый осмотр", "инспекционный визит"}

    knm_info = {}
    detail = {k: [] for k in ('vks_denom', 'vks_num', 'och_denom', 'och_num', 'och_nar_denom', 'och_nar_num')}
    rejected_vks, rejected_och = [], []
    denom_rows_vks, denom_rows_och = {}, {}

    for row in data:
        reasons_base = []
        if normalize_str(row[status_col]) != "завершена":
            reasons_base.append(f"Статус: {row[status_col]}")
        if normalize_str(row[proverka_col]) != "нет":
            reasons_base.append(f"Проверка ОГВ: {row[proverka_col]}")
        if normalize_str(row[vid_nadzora_col]) == "гнго":
            reasons_base.append("Вид надзора: ГНГО")
        podr = row[podr_col] if row[podr_col] else "Не указано"
        knm = row[knm_col]
        if not podr or not knm:
            reasons_base.append("Пустое подразделение или номер КНМ")

        if reasons_base:
            reason_str = "; ".join(reasons_base)
            rejected_vks.append(tuple(row) + (reason_str,))
            rejected_och.append(tuple(row) + (reason_str,))
            continue

        vid_val = normalize_str(row[vid_col]) if row[vid_col] else ""
        knd_str = normalize_str(row[knd_col]) if row[knd_col] else ""
        nar_str = normalize_str(row[nar_col]) if row[nar_col] else ""
        vks_str = normalize_str(row[vks_col]) if row[vks_col] else ""
        ssylki_val = row[ssylki_col]
        ssylki_not_empty = ssylki_val is not None and str(ssylki_val).strip() != ""

        if vid_val in allowed_vids:
            detail['vks_denom'].append(row)
            if vks_str == "да" and ssylki_not_empty:
                detail['vks_num'].append(row)
            if knm not in denom_rows_vks and not (vks_str == "да" and ssylki_not_empty):
                denom_rows_vks[knm] = (row, f"С ВКС ≠ 'да': {row[vks_col]}" if vks_str != "да" else "Ссылки пустые")
        else:
            rejected_vks.append(tuple(row) + (f"Вид КНМ не входит в список ВКС: {row[vid_col]}",))

        if vid_val in allowed_vids:
            if "осмотр" in knd_str:
                detail['och_denom'].append(row)
                if vks_str == "нет" and ssylki_not_empty:
                    detail['och_num'].append(row)
                if knm not in denom_rows_och and not (vks_str == "нет" and ssylki_not_empty):
                    denom_rows_och[knm] = (row, f"С ВКС ≠ 'нет': {row[vks_col]}" if vks_str != "нет" else "Ссылки пустые")
                if nar_str == "да":
                    detail['och_nar_denom'].append(row)
                    if vks_str == "нет" and ssylki_not_empty:
                        detail['och_nar_num'].append(row)
            else:
                rejected_och.append(tuple(row) + (f"КНД не содержит 'осмотр': {row[knd_col]}",))
        else:
            rejected_och.append(tuple(row) + (f"Вид КНМ не входит в список очных: {row[vid_col]}",))

        if knm not in knm_info:
            knm_info[knm] = {
                'podr': podr, 'vks_denom': False, 'vks_num': False,
                'och_denom': False, 'och_num': False, 'och_nar': False
            }
        info = knm_info[knm]
        if vid_val in allowed_vids:
            info['vks_denom'] = True
            if vks_str == "да" and ssylki_not_empty:
                info['vks_num'] = True
        if vid_val in allowed_vids and "осмотр" in knd_str:
            info['och_denom'] = True
            if vks_str == "нет" and ssylki_not_empty:
                info['och_num'] = True
            if nar_str == "да":
                info['och_nar'] = True

    metrics = defaultdict(lambda: [set(), set(), set(), set(), set()])
    for knm, info in knm_info.items():
        podr = info['podr']
        if info['vks_denom']: metrics[podr][0].add(knm)
        if info['vks_num']:   metrics[podr][1].add(knm)
        if info['och_denom']: metrics[podr][2].add(knm)
        if info['och_num']:   metrics[podr][3].add(knm)
        if info['och_nar']:   metrics[podr][4].add(knm)

    denom_not_num_vks = [
        tuple(row) + (reason,)
        for knm, (row, reason) in denom_rows_vks.items()
        if knm not in knm_info or not knm_info[knm]['vks_num']
    ]
    denom_not_num_och = [
        tuple(row) + (reason,)
        for knm, (row, reason) in denom_rows_och.items()
        if knm not in knm_info or not knm_info[knm]['och_num']
    ]

    return metrics, detail, rejected_vks, rejected_och, denom_not_num_vks, denom_not_num_och

def build_report_data(metrics):
    return [{
        'Подразделение': podr,
        'total_vks': len(sets[0]),
        'prim_vks': len(sets[1]),
        'total_och': len(sets[2]),
        'prim_och': len(sets[3]),
        'total_och_nar': len(sets[4]),
    } for podr, sets in metrics.items()]

def style_header(ws, fill_color):
    fill = PatternFill("solid", start_color=fill_color, end_color=fill_color)
    font = Font(bold=True, color="FFFFFF", name="Arial")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 40

def autowidth(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

def save_to_excel(report_data, filename, headers_orig, detail, rejected_vks, rejected_och,
                  denom_not_num_vks, denom_not_num_och, subjekt_name, date_from, date_to):
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "Итоги по подразделениям"

    summary_headers = [
        "№", "Подразделение",
        "Всего в ААС за выбранный период",
        "Всего в МП Инспектор за выбранный период (доля)",
        "Всего в ААС КНД за выбранный период (из них с нарушениями)",
        "Всего в МП Инспектор за выбранный период (доля)"
    ]
    ws_summary.append(summary_headers)
    style_header(ws_summary, "4472C4")
    ws_summary.row_dimensions[1].height = 40

    sorted_data = sorted(report_data, key=lambda x: x['Подразделение'])

    def fmt_pct(part, total):
        return f"{part} ({part/total*100:.2f}%)" if total > 0 else "0 (0.00%)"

    def fmt_och_total(total, nar):
        return f"{total} ({nar})"

    for row_num, d in enumerate(sorted_data, 1):
        ws_summary.append([
            row_num,
            d['Подразделение'],
            d['total_vks'],
            fmt_pct(d['prim_vks'], d['total_vks']),
            fmt_och_total(d['total_och'], d['total_och_nar']),
            fmt_pct(d['prim_och'], d['total_och'])
        ])

    total_vks   = sum(d['total_vks']   for d in sorted_data)
    prim_vks    = sum(d['prim_vks']    for d in sorted_data)
    total_och   = sum(d['total_och']   for d in sorted_data)
    prim_och    = sum(d['prim_och']    for d in sorted_data)
    total_och_nar = sum(d['total_och_nar'] for d in sorted_data)

    ws_summary.append([
        "", f"ИТОГО по субъекту {subjekt_name}",
        total_vks, fmt_pct(prim_vks, total_vks),
        fmt_och_total(total_och, total_och_nar), fmt_pct(prim_och, total_och)
    ])
    last_row = ws_summary.max_row
    for col in range(1, 7):
        ws_summary.cell(row=last_row, column=col).font = Font(bold=True)

    for col in ws_summary.columns:
        max_len = max((len(str(cell.value)) for cell in col if cell.value), default=10)
        ws_summary.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

    def write_detail_sheet(title, rows, fill_color, extra_header=None):
        ws = wb.create_sheet(title=title)
        hdrs = list(headers_orig) + ([extra_header] if extra_header else [])
        ws.append(hdrs)
        style_header(ws, fill_color)
        for row in rows:
            ws.append(list(row))
        autowidth(ws)

    write_detail_sheet("ВКС всего",        detail['vks_denom'],   "538135")
    write_detail_sheet("ВКС с МП",         detail['vks_num'],     "C55A11")
    write_detail_sheet("ВКС без МП",       denom_not_num_vks,    "7030A0", extra_header="Причина")
    write_detail_sheet("ВКС - Отсеянные",  rejected_vks,          "C00000", extra_header="Причина отклонения")
    write_detail_sheet("Очные всего",      detail['och_denom'],   "538135")
    write_detail_sheet("Очные с МП",       detail['och_num'],     "C55A11")
    write_detail_sheet("Очные без МП",     denom_not_num_och,    "7030A0", extra_header="Причина")
    write_detail_sheet("Очные - Отсеянные", rejected_och,         "C00000", extra_header="Причина отклонения")

    ws_info = wb.create_sheet("Информация")
    for row in [
        ["Выбранный файл:", filename],
        ["Период:", f"с {date_from.strftime('%d.%m.%Y')} по {date_to.strftime('%d.%m.%Y')}"],
        ["Субъект РФ:", subjekt_name],
        ["Дата формирования:", datetime.now().strftime('%d.%m.%Y %H:%M:%S')]
    ]:
        ws_info.append(row)
    for col in ws_info.columns:
        ws_info.column_dimensions[col[0].column_letter].width = 30

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.read()


# ================== Streamlit UI ==================

st.title("📊 Анализ применения МП Инспектор")
st.markdown("Загрузите файл выгрузки, выберите период — и получите готовый отчёт.")

st.divider()

# Шаг 1: Загрузка файла
uploaded_file = st.file_uploader(
    "Шаг 1. Загрузите файл выгрузки (.xlsx)",
    type=["xlsx", "xlsm"],
    help=f"Файл должен содержать лист «{SHEET_NAME}»"
)

if uploaded_file:
    file_bytes = uploaded_file.read()

    # Загружаем данные
    try:
        with st.spinner("Читаю файл..."):
            data, col_idx, headers_orig = load_data(file_bytes)
    except Exception as e:
        st.error(f"Ошибка при загрузке файла: {e}")
        st.stop()

    # Определяем субъект РФ
    subjekt_col = col_idx['subjekt']
    subjekt_name = None
    for row in data:
        if row[subjekt_col]:
            subjekt_name = str(row[subjekt_col]).strip()
            break

    if not subjekt_name:
        st.error("Не удалось определить субъект РФ из данных файла.")
        st.stop()

    district = SUBJEKT_TO_DISTRICT.get(" ".join(subjekt_name.split()), "Не определено")

    col1, col2 = st.columns(2)
    col1.success(f"✅ Файл загружен: **{uploaded_file.name}**")
    col1.info(f"📌 Субъект РФ: **{subjekt_name}**")
    col2.info(f"🗺️ Округ: **{district}**")
    col2.info(f"📋 Строк в файле: **{len(data)}**")

    st.divider()

    # Шаг 2: Выбор периода
    st.subheader("Шаг 2. Выберите период (дата составления акта)")
    c1, c2 = st.columns(2)
    today = date.today()
    default_from = date(today.year, 1, 1)

    date_from = c1.date_input("Начало периода", value=default_from, format="DD.MM.YYYY")
    date_to   = c2.date_input("Конец периода",  value=today,        format="DD.MM.YYYY")

    if date_from > date_to:
        st.warning("⚠️ Начальная дата позже конечной — они будут поменяны местами.")
        date_from, date_to = date_to, date_from

    st.divider()

    # Шаг 3: Расчёт
    if st.button("🚀 Сформировать отчёт", type="primary", use_container_width=True):
        with st.spinner("Фильтрую данные..."):
            data_filtered, skipped = filter_by_date(data, col_idx, date_from, date_to)

        if not data_filtered:
            st.error(f"За период {date_from.strftime('%d.%m.%Y')} — {date_to.strftime('%d.%m.%Y')} не найдено ни одной записи.")
            st.stop()

        st.info(f"После фильтра по дате: **{len(data_filtered)}** строк. Пропущено (вне периода / некорректная дата): **{skipped}**.")

        with st.spinner("Считаю показатели..."):
            metrics, detail, rejected_vks, rejected_och, denom_not_num_vks, denom_not_num_och = \
                calculate_all_metrics(data_filtered, col_idx)
            report_data = build_report_data(metrics)

        with st.spinner("Формирую Excel..."):
            result_bytes = save_to_excel(
                report_data, uploaded_file.name, headers_orig,
                detail, rejected_vks, rejected_och,
                denom_not_num_vks, denom_not_num_och,
                subjekt_name, date_from, date_to
            )

        st.success("✅ Отчёт готов!")

        # Превью сводной таблицы
        st.subheader("📋 Сводка по подразделениям")
        sorted_data = sorted(report_data, key=lambda x: x['Подразделение'])

        def fmt_pct(part, total):
            return f"{part} ({part/total*100:.2f}%)" if total > 0 else "0 (0.00%)"

        preview_rows = []
        for i, d in enumerate(sorted_data, 1):
            preview_rows.append({
                "№": i,
                "Подразделение": d['Подразделение'],
                "ВКС (всего)": d['total_vks'],
                "ВКС с МП (доля)": fmt_pct(d['prim_vks'], d['total_vks']),
                "Очные (всего / с нарушениями)": f"{d['total_och']} ({d['total_och_nar']})",
                "Очные с МП (доля)": fmt_pct(d['prim_och'], d['total_och']),
            })

        st.dataframe(preview_rows, use_container_width=True, hide_index=True)

        total_vks = sum(d['total_vks'] for d in sorted_data)
        prim_vks  = sum(d['prim_vks']  for d in sorted_data)
        total_och = sum(d['total_och'] for d in sorted_data)
        prim_och  = sum(d['prim_och']  for d in sorted_data)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("ВКС всего",   total_vks)
        m2.metric("ВКС с МП",   fmt_pct(prim_vks, total_vks))
        m3.metric("Очные всего", total_och)
        m4.metric("Очные с МП", fmt_pct(prim_och, total_och))

        st.divider()

        result_filename = f"результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button(
            label="⬇️ Скачать полный отчёт (.xlsx)",
            data=result_bytes,
            file_name=result_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )
