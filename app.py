import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.styles.borders import Border, Side
import re
import io
import math
from datetime import datetime, date
from collections import defaultdict

st.set_page_config(page_title="МП Инспектор — Анализ КНМ", page_icon="📊", layout="wide")

SHEET_NAME = "Детализация МП Инспектор"

COLUMN_KEYWORDS = {
    'subjekt':         ['субъект рф'],
    'podrazd':         ['подразделение'],
    'vid_nadzora':     ['вид надзора'],
    'nom_knm':         ['номер кнм'],
    'vid':             ['вид'],
    'status':          ['статус кнм'],
    'narusheniya':     ['нарушения выявлены'],
    'proverka_ogv':    ['проверка огв/омсу'],
    'knd':             ['кнд'],
    'ssylki':          ['ссылки на файлы'],
    'date_act':        ['дата составления акта о результате кнм', 'дата составления акта'],
    'tip_prof_vizita': ['тип проф. визита', 'тип профилактического визита'],
    's_vks':           ['с вкс', 'вкс'],
}

def normalize_str(s):
    if s is None:
        return ""
    return re.sub(r'\s+', ' ', str(s).strip()).lower()

def find_column_index(headers, possible_names):
    headers_norm  = [normalize_str(h) for h in headers]
    possible_norm = [normalize_str(n) for n in possible_names]
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
    if isinstance(value, date) and not isinstance(value, datetime):
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
        raise ValueError(
            f"Лист «{SHEET_NAME}» не найден.\n"
            f"Доступные листы: {', '.join(wb.sheetnames)}"
        )
    ws = wb[SHEET_NAME]
    headers      = [cell.value if cell.value else "" for cell in ws[1]]
    headers_orig = headers[:]

    col_indices = {}
    warnings_   = []
    missing     = []
    for key, possible_names in COLUMN_KEYWORDS.items():
        idx = find_column_index(headers, possible_names)
        if idx is None:
            if key == 'podrazd':
                idx = 17
                warnings_.append("Столбец «подразделение» не найден по имени — используем позицию 18 (индекс 17)")
            else:
                missing.append(f"«{key}» (искали: {possible_names})")
        col_indices[key] = idx

    if missing:
        raise ValueError("Не найдены обязательные столбцы:\n• " + "\n• ".join(missing))

    data = [
        row for row in ws.iter_rows(min_row=2, values_only=True)
        if not all(c is None for c in row)
    ]
    return data, col_indices, headers_orig, warnings_

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
    subj_col = col_idx["subjekt"]
    podrazd_col = col_idx["podrazd"]
    knm_col = col_idx["nom_knm"]
    vid_col = col_idx["vid"]
    status_col = col_idx["status"]
    proverka_col = col_idx["proverka_ogv"]
    vid_nadzora_col = col_idx["vid_nadzora"]
    knd_col = col_idx["knd"]
    nar_col = col_idx["narusheniya"]
    vks_col = col_idx["s_vks"]
    ssylki_col = col_idx["ssylki"]

    allowed_vids = {"", "выездная проверка", "рейдовый осмотр", "инспекционный визит"}

    knm_info = {}
    subj_of = {}
    seen = {key: set() for key in ("vks_denom", "vks_num", "och_denom", "och_num", "och_nar_denom", "och_nar_num")}
    detail = {key: [] for key in seen}
    denom_rows_vks = {}
    denom_rows_och = {}
    rejected_vks = []
    rejected_och = []

    for row in data:
        reasons_base = []

        if normalize_str(row[status_col]) != "завершена":
            reasons_base.append(f"Статус: {row[status_col]}")
        if normalize_str(row[proverka_col]) != "нет":
            reasons_base.append(f"Проверка ОГВ: {row[proverka_col]}")
        if normalize_str(row[vid_nadzora_col]) == "гнго":
            reasons_base.append("Вид надзора: ГНГО")

        subj = row[subj_col]
        podrazd = row[podrazd_col]
        knm = row[knm_col]

        if not subj or not knm:
            reasons_base.append("Пустой субъект или номер КНМ")

        if reasons_base:
            reason_text = "; ".join(reasons_base)
            rejected_vks.append(tuple(row) + (reason_text,))
            rejected_och.append(tuple(row) + (reason_text,))
            continue

        pod_key = str(podrazd).strip() if podrazd else "Не указано"
        if pod_key not in subj_of:
            subj_of[pod_key] = str(subj).strip() if subj else ""

        vid_val = normalize_str(row[vid_col]) if row[vid_col] else ""
        knd_str = normalize_str(row[knd_col]) if row[knd_col] else ""
        nar_str = normalize_str(row[nar_col]) if row[nar_col] else ""
        vks_str = normalize_str(row[vks_col]) if row[vks_col] else ""
        ssylki_ok = row[ssylki_col] is not None and str(row[ssylki_col]).strip() != ""
        sk = knm

        if vid_val in allowed_vids:
            append_unique(detail, seen, "vks_denom", sk, row)
            if vks_str == "да" and ssylki_ok:
                append_unique(detail, seen, "vks_num", sk, row)
            if sk not in denom_rows_vks:
                denom_rows_vks[sk] = (row, build_reason(vks_str, "да", ssylki_ok, row[vks_col]))
        else:
            rejected_vks.append(tuple(row) + (f"Вид КНМ не входит в список ВКС: {row[vid_col]}",))

        if vid_val in allowed_vids:
            if "осмотр" in knd_str:
                append_unique(detail, seen, "och_denom", sk, row)
                if vks_str == "нет" and ssylki_ok:
                    append_unique(detail, seen, "och_num", sk, row)
                if sk not in denom_rows_och:
                    denom_rows_och[sk] = (row, build_reason(vks_str, "нет", ssylki_ok, row[vks_col]))
                if nar_str == "да":
                    append_unique(detail, seen, "och_nar_denom", sk, row)
                    if vks_str == "нет" and ssylki_ok:
                        append_unique(detail, seen, "och_nar_num", sk, row)
            else:
                rejected_och.append(tuple(row) + (f"КНД не содержит 'осмотр': {row[knd_col]}",))
        else:
            rejected_och.append(tuple(row) + (f"Вид КНМ не входит в список очных: {row[vid_col]}",))

        if knm not in knm_info:
            knm_info[knm] = {
                "podr": pod_key,
                "vks_denom": False,
                "vks_num": False,
                "och_denom": False,
                "och_num": False,
                "och_nar": False,
            }

        info = knm_info[knm]

        if vid_val in allowed_vids:
            info["vks_denom"] = True
            if vks_str == "да" and ssylki_ok:
                info["vks_num"] = True

        if vid_val in allowed_vids and "осмотр" in knd_str:
            info["och_denom"] = True
            if vks_str == "нет" and ssylki_ok:
                info["och_num"] = True
            if nar_str == "да":
                info["och_nar"] = True

    metrics = defaultdict(lambda: [set(), set(), set(), set(), set()])

    for knm, info in knm_info.items():
        podr = info["podr"]
        if info["vks_denom"]:
            metrics[podr][0].add(knm)
        if info["vks_num"]:
            metrics[podr][1].add(knm)
        if info["och_denom"]:
            metrics[podr][2].add(knm)
        if info["och_num"]:
            metrics[podr][3].add(knm)
        if info["och_nar"]:
            metrics[podr][4].add(knm)

    denom_not_num_vks = [
        tuple(row) + (reason,)
        for sk, (row, reason) in denom_rows_vks.items()
        if sk not in seen["vks_num"]
    ]
    denom_not_num_och = [
        tuple(row) + (reason,)
        for sk, (row, reason) in denom_rows_och.items()
        if sk not in seen["och_num"]
    ]

    return metrics, subj_of, detail, rejected_vks, rejected_och, denom_not_num_vks, denom_not_num_och

def build_report_data(metrics, subj_of):
    return [{
        'Подразделение': pod,
        'Субъект':       subj_of.get(pod, ""),
        'total_vks':     len(sets[0]),
        'prim_vks':      len(sets[1]),
        'total_och':     len(sets[2]),
        'prim_och':      len(sets[3]),
        'total_och_nar': len(sets[4]),
    } for pod, sets in metrics.items()]

def build_excel(report_data, headers_orig, detail,
                rejected_vks, rejected_och, denom_not_num_vks, denom_not_num_och):
    wb = openpyxl.Workbook()

    def make_fill(hex_color):
        return PatternFill("solid", start_color=hex_color, end_color=hex_color)

    def style_row(ws, row_num, fill_color, bold=True, font_color="FFFFFF"):
        fill = make_fill(fill_color)
        font = Font(bold=bold, color=font_color, name="Arial", size=10)
        for cell in ws[row_num]:
            cell.fill = fill; cell.font = font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def autowidth(ws):
        for col in ws.columns:
            w = max((len(str(c.value)) for c in col if c.value), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(w + 3, 60)

    def write_detail_sheet(title, rows, fill_color, extra_header=None):
        ws = wb.create_sheet(title=title)
        hdrs = list(headers_orig) + ([extra_header] if extra_header else [])
        ws.append(hdrs)
        style_row(ws, 1, fill_color)
        ws.row_dimensions[1].height = 40
        for row in rows:
            ws.append(list(row))
        autowidth(ws)

    BLUE = "4472C4"
    ITOG = "1F3864"
    thin   = Side(style='thin', color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws = wb.active
    ws.title = "Итоги по подразделениям"
    ws.append(["", "", "ВКС", "", "ОЧНЫЕ", ""])
    ws.merge_cells("C1:D1"); ws.merge_cells("E1:F1")
    for cell in ws[1]:
        cell.fill      = make_fill(BLUE)
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    ws.append(["№ п/п", "Подразделение",
               "Всего в ААС КНД", "Всего в МП Инспектор (доля, %)",
               "Всего в ААС КНД\n(из них с нарушениями)", "Всего в МП Инспектор (доля, %)"])
    style_row(ws, 2, BLUE)
    ws.row_dimensions[2].height = 45
    for i, w in enumerate([7, 45, 20, 30, 28, 30], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    def fmt_vks(p, t):
        return f"{p} ({p/t*100:.2f}%)" if t else f"{p} (0.00%)"
    def fmt_och_d(t, n):
        return f"{t} ({n})"
    def fmt_och_n(p, t):
        return f"{p} ({p/t*100:.2f}%)" if t else f"{p} (0.00%)"
    def sort_key(i):
        return i['prim_och'] / i['total_och'] if i['total_och'] > 0 else 0.0

    n = 1
    for item in sorted(report_data, key=sort_key):
        ws.append([n, item['Подразделение'], item['total_vks'],
                   fmt_vks(item['prim_vks'], item['total_vks']),
                   fmt_och_d(item['total_och'], item['total_och_nar']),
                   fmt_och_n(item['prim_och'], item['total_och'])])
        cur = ws.max_row
        for cell in ws[cur]:
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.font      = Font(name="Arial", size=10)
            cell.border    = border
        ws.cell(cur, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if n % 2 == 0:
            for cell in ws[cur]: cell.fill = make_fill("DCE6F1")
        pod_text = str(item['Подразделение']) if item['Подразделение'] else ""
        lines    = math.ceil(len(pod_text) / 43) if pod_text else 1
        ws.row_dimensions[cur].height = max(lines * 14 + 6, 20)
        n += 1

    tv = sum(i['total_vks']    for i in report_data)
    pv = sum(i['prim_vks']     for i in report_data)
    to = sum(i['total_och']    for i in report_data)
    po = sum(i['prim_och']     for i in report_data)
    tn = sum(i['total_och_nar']for i in report_data)
    ws.append(["", "Итог по субъекту", tv, fmt_vks(pv, tv), fmt_och_d(to, tn), fmt_och_n(po, to)])
    itog = ws.max_row
    for cell in ws[itog]:
        cell.fill      = make_fill(ITOG)
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
    ws.cell(itog, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[itog].height = 24

    write_detail_sheet("ВКС всего",          detail['vks_denom'],     "538135")
    write_detail_sheet("ВКС с МП",           detail['vks_num'],       "C55A11")
    write_detail_sheet("ВКС без МП",         denom_not_num_vks,       "7030A0", "Причина")
    write_detail_sheet("ВКС — Отсеянные",    rejected_vks,            "C00000", "Причина отклонения")
    write_detail_sheet("Очные всего",        detail['och_denom'],     "538135")
    write_detail_sheet("Очные с МП",         detail['och_num'],       "C55A11")
    write_detail_sheet("Очные без МП",       denom_not_num_och,       "7030A0", "Причина")
    write_detail_sheet("Очные всего (нар.)", detail['och_nar_denom'], "375623")
    write_detail_sheet("Очные с МП (нар.)",  detail['och_nar_num'],   "843C0C")
    write_detail_sheet("Очные — Отсеянные",  rejected_och,            "C00000", "Причина отклонения")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ================== Streamlit UI ==================

st.title("📊 МП Инспектор — Итоги по подразделениям")
st.markdown("Загрузите файл выгрузки АИС КНД, выберите период и получите готовый отчёт.")

uploaded = st.file_uploader("**Шаг 1 — Загрузите файл выгрузки (.xlsx / .xlsm)**",
                             type=["xlsx", "xlsm"])

if uploaded:
    file_bytes = uploaded.read()

    with st.spinner("Читаем файл..."):
        try:
            data, col_idx, headers_orig, warnings_ = load_data(file_bytes)
        except Exception as e:
            st.error(f"Ошибка при загрузке: {e}")
            st.stop()

    for w in warnings_:
        st.warning(f"⚠️ {w}")

    st.success(f"✅ Файл загружен. Строк данных: **{len(data)}**")

    st.markdown("---")
    st.markdown("**Шаг 2 — Укажите период (дата составления акта КНМ)**")
    today = date.today()
    col1, col2 = st.columns(2)
    with col1:
        date_from = st.date_input("Дата начала", value=today)
    with col2:
        date_to = st.date_input("Дата окончания", value=today)

    if date_from > date_to:
        st.warning("⚠️ Начальная дата позже конечной — даты поменяны местами.")
        date_from, date_to = date_to, date_from

    st.markdown("---")
    if st.button("🚀 Запустить расчёт", type="primary", use_container_width=True):

        with st.spinner("Фильтруем по дате..."):
            data_filtered, skipped = filter_by_date(data, col_idx, date_from, date_to)

        if len(data_filtered) == 0:
            st.error("За выбранный период данных нет. Проверьте даты.")
            st.stop()

        st.info(
            f"Строк в периоде: **{len(data_filtered)}** &nbsp;|&nbsp; "
            f"Пропущено (вне периода / без даты): **{skipped}**"
        )

        with st.spinner("Считаем показатели..."):
            metrics, subj_of, detail, rej_vks, rej_och, dnn_vks, dnn_och = \
                calculate_all_metrics(data_filtered, col_idx)
            report_data = build_report_data(metrics, subj_of)

        st.markdown("---")
        st.subheader("📈 Итоги по субъекту")

        tv = sum(i['total_vks']    for i in report_data)
        pv = sum(i['prim_vks']     for i in report_data)
        to = sum(i['total_och']    for i in report_data)
        po = sum(i['prim_och']     for i in report_data)
        tn = sum(i['total_och_nar']for i in report_data)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("ВКС всего КНМ",       tv)
        c2.metric("ВКС с МП",            pv,
                  delta=f"{pv/tv*100:.1f}%" if tv else "—", delta_color="normal")
        c3.metric("Очные всего КНМ",     to)
        c4.metric("Очные с МП",          po,
                  delta=f"{po/to*100:.1f}%" if to else "—", delta_color="normal")
        c5.metric("Очные с нарушениями", tn)

        with st.expander("📋 Таблица по подразделениям", expanded=True):
            import pandas as pd
            rows_display = []
            for item in sorted(report_data,
                               key=lambda x: x['prim_och']/x['total_och'] if x['total_och'] > 0 else 0):
                rows_display.append({
                    "Подразделение":  item['Подразделение'],
                    "ВКС всего":      item['total_vks'],
                    "ВКС с МП":       item['prim_vks'],
                    "ВКС доля, %":    f"{item['prim_vks']/item['total_vks']*100:.1f}" if item['total_vks'] else "—",
                    "Очные всего":    item['total_och'],
                    "Очные с нар.":   item['total_och_nar'],
                    "Очные с МП":     item['prim_och'],
                    "Очные доля, %":  f"{item['prim_och']/item['total_och']*100:.1f}" if item['total_och'] else "—",
                })
            st.dataframe(pd.DataFrame(rows_display), use_container_width=True, hide_index=True)

        with st.spinner("Формируем Excel-файл..."):
            excel_bytes = build_excel(
                report_data, headers_orig, detail,
                rej_vks, rej_och, dnn_vks, dnn_och
            )

        filename = f"результат_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.success("✅ Готово!")
        st.download_button(
            label="⬇️ Скачать результат (.xlsx)",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
