"""
Нуртелеком · Контроль выплат по договорам.

Запуск:
    pip install streamlit pandas plotly openpyxl
    python -m streamlit run app.py

Рядом с app.py кладутся выгрузки по месяцам: «Результат_май.xlsx»,
«Результат_апрель.xlsx» и так далее — подходит любое имя вида «Результат*.xlsx».
Период берётся из колонки DAT внутри файла, а не из имени, поэтому новый месяц
достаточно положить в папку. Опционально рядом — «kg_regions.geojson».
Решения по выплатам сохраняются в «Решения.xlsx» рядом со скриптом — два листа
(«Выплаты», «Договоры»), можно открыть и посмотреть/поправить напрямую в Excel.
"""
import base64
import io
import json
import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

HERE = Path(__file__).resolve().parent
# Пути к рабочей машине в коде больше нет. При деплое (Streamlit Cloud и т.п.)
# выгрузки загружаются через кнопку в сайдбаре. Локально/в репозитории файлы
# берутся из папки, которую можно задать переменной окружения DATA_FOLDER,
# иначе — из папки «data» рядом со скриптом (если она есть).
DATA_FOLDER = Path(os.environ.get('DATA_FOLDER', HERE / 'data'))
DATA_GLOB = 'Результат *.xlsx'
GEOJSON_FILE = HERE / 'kg_regions.geojson'
def find_logo():
    # логотип может лежать и в assets/logo.png, и просто рядом с app.py —
    # проверяем оба варианта, плюс на всякий случай .jpg/.jpeg
    for name in ('assets/logo.png', 'logo.png', 'assets/logo.jpg', 'logo.jpg',
                 'assets/logo.jpeg', 'logo.jpeg'):
        p = HERE / name
        if p.exists():
            return p
    return None


LOGO_FILE = find_logo()
DECISIONS_NAME = 'Решения.xlsx'
DECISIONS_FILE = HERE / DECISIONS_NAME

# правило «оплата возобновилась, долг»: BALANCE_END == Sum reestr — начисление
# месяца остаётся на следующий, значит выплата закрывала долг.
RESUMED_RULE = True

# светлая корпоративная тема — фон белый, текст чёрный, акцент маджента (из логотипа)
INK, PANEL, LINE = '#FFFFFF', '#F7F5F7', '#E6E2E6'
TEXT, MUTED = '#0A0A0A', '#6B6470'
# фирменные цвета О! (из логотипа) — маджента как основной акцент.
# Остальные — подобраны так, чтобы хорошо читались на белом и не спорили с маджентой
MAGENTA = '#EF0277'
OK, WARN, COOL, CYAN, VIOLET = MAGENTA, '#B45309', '#2563EB', '#0891B2', '#3D4451'


REESTR = 'Выплаты по реестру'
EL_ONLY = 'Электроэнергия'
TECH_ONLY = 'Техобслуживание'
GR_ONLY = 'Групповые'
REVIEW = 'На проверку'

GROUPS = [REESTR, EL_ONLY, TECH_ONLY, GR_ONLY, REVIEW]
NOCHECK = [EL_ONLY, TECH_ONLY, GR_ONLY]
GROUP_COLOR = {REESTR: OK, EL_ONLY: COOL, TECH_ONLY: CYAN, GR_ONLY: VIOLET, REVIEW: WARN}
# короткие подписи для осей: полные названия налезают на легенду
SHORT = {REESTR: 'По реестру', EL_ONLY: 'Электро-<br>энергия', TECH_ONLY: 'Техобслу-<br>живание',
         GR_ONLY: 'Групповые', REVIEW: 'На проверку'}

# флаг из самой выгрузки: колонка "Не следовало платить" (bool) — build_result.py
# сам находит платежи, у которых по UNIQUE_ID в реестре нет ни одной активной записи
NOPAY = 'Платить не следовало'
NOPAY_COL = 'Не следовало платить'

# основания, по которым выплата попала в реестр — соответствуют реальным
# колонкам build_result.py, проверяются по порядку (см. load())
MATCHED_FINAL = 'FINAL (сумма сходится)'
MATCHED_RAN = 'RAN Sharing'
MATCHED_100050 = 'Проверка 100050'
MATCHED_NDS = 'Проверка НДС (1,14)'
MATCHED_NOMER = 'Один договор — точное совпадение'
MATCHED_PO_NOMERU = 'Один договор — по номеру'
MATCHED_PO_KONTR = 'Один договор — по контрагенту'
MATCHED_OSOBYI_NOVYI = 'Особый случай — новый договор'
MATCHED_OSOBYI_SMENA = 'Особый случай — смена номера'
MANUAL_REESTR = 'Оплата по реестру'
RESUMED = 'Оплата возобновилась, долг'
BASIS_COLOR = {MATCHED_FINAL: OK, 'Оплата долга': COOL, 'Аванс': VIOLET,
               'Оплата долга и аванс': VIOLET, RESUMED: CYAN, MANUAL_REESTR: OK}

# решения, которые ставят руками во время разбора
CLOSING = ['Оплата долга', 'Оплата на будущее', 'Оплата долга и на будущее',
           'Норма по договору', 'Разовая согласованная выплата']
OPEN = 'Не удалось объяснить'
DECISIONS = [''] + CLOSING + [OPEN]

# ручное решение не должно плодить отдельную категорию "Разобрано вручную: …" —
# оно просто попадает в ту же корзину по основанию, что и автоматика
MANUAL_BASIS_MAP = {
    'Оплата долга': 'Оплата долга',
    'Оплата на будущее': 'Аванс',
    'Оплата долга и на будущее': 'Оплата долга и аванс',
    'Норма по договору': MANUAL_REESTR,
    'Разовая согласованная выплата': MANUAL_REESTR,
}

REGION_ORDER = ['Чуйская', 'Ошская', 'Джалал-Абадская', 'Иссык-Кульская',
                'Нарынская', 'Баткенская', 'Таласская', 'Групповые', 'Другое']

# первая цифра кода сайта (N до первого "_") кодирует область — карта прислана
# вручную; 9 — групповые сайты (то же правило, что и колонка "Групповые" в
# самой выгрузке), всё, что не 1-7/9 (в т.ч. служебные коды вроде "__0") —
# "Другое". Определяем сами в дашборде — колонка "Область" в файле результата
# для этого не нужна.
REGION_MAP = {
    '1': 'Баткенская область', '2': 'Джалал-Абадская область',
    '3': 'Иссык-Кульская область', '4': 'Нарынская область',
    '5': 'Ошская область', '6': 'Таласская область',
    '7': 'Чуйская область', '9': 'Групповые',
}


def region_of(n):
    return n.astype(str).str[0].map(REGION_MAP).fillna('Другое')

MONTHS = ['', 'январь', 'февраль', 'март', 'апрель', 'май', 'июнь',
          'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']

st.set_page_config(page_title='Контроль выплат по договорам', page_icon='◆',
                   layout='wide', initial_sidebar_state='expanded')

st.markdown(f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@500;600&display=swap');

  .stApp {{ background:{INK}; color:{TEXT}; }}
  header[data-testid="stHeader"] {{ background:transparent; }}
  section[data-testid="stSidebar"] {{ background:{PANEL}; border-right:1px solid {LINE}; }}
  html, body, [class*="css"] {{ font-family:'IBM Plex Sans', sans-serif; }}
  .block-container {{ padding-top:2rem; max-width:1560px; }}

  .masthead {{ display:flex; align-items:center; gap:.8rem; flex-wrap:wrap;
               border-bottom:1px solid {LINE}; padding-bottom:.9rem; margin-bottom:1.3rem; }}
  .masthead img {{ height:30px; }}
  .masthead .mark {{ font-family:'IBM Plex Mono',monospace; font-weight:600;
                     font-size:.78rem; letter-spacing:.22em; color:{MAGENTA}; }}
  .masthead .date {{ margin-left:auto; font-family:'IBM Plex Mono',monospace;
                     font-size:.82rem; color:{MUTED}; }}

  .kpi {{ background:{PANEL}; border:1px solid {LINE}; border-radius:10px;
          padding:.85rem .9rem; height:100%; border-left:3px solid {LINE}; }}
  .kpi .label {{ font-size:.67rem; text-transform:uppercase; letter-spacing:.07em;
                 color:{MUTED}; margin-bottom:.4rem; min-height:2.1em; line-height:1.25; }}
  .kpi .value {{ font-family:'IBM Plex Mono',monospace; font-weight:600;
                 font-size:1.22rem; line-height:1.1; color:{TEXT};
                 font-variant-numeric:tabular-nums; }}
  .kpi .unit {{ font-size:.72rem; color:{MUTED}; font-weight:500; }}
  .kpi .sub {{ font-size:.72rem; color:{MUTED}; margin-top:.35rem; }}

  .mini {{ background:{PANEL}; border:1px solid {LINE}; border-radius:8px;
           padding:.6rem .8rem; height:100%; border-left:3px solid {LINE}; }}
  .mini .label {{ font-size:.68rem; text-transform:uppercase; letter-spacing:.08em;
                  color:{MUTED}; margin-bottom:.25rem; }}
  .mini .value {{ font-family:'IBM Plex Mono',monospace; font-weight:600;
                  font-size:1.02rem; color:{TEXT}; font-variant-numeric:tabular-nums; }}
  .mini .sub {{ font-size:.7rem; color:{MUTED}; margin-top:.15rem; }}

  .section-title {{ font-family:'IBM Plex Mono',monospace; font-size:.74rem;
                    letter-spacing:.16em; text-transform:uppercase; color:{MUTED};
                    margin:1.6rem 0 .6rem; }}
  .note {{ font-size:.8rem; color:{MUTED}; }}

  div[data-testid="stDataFrame"] {{ border:1px solid {LINE}; border-radius:10px; }}
  .stTabs [data-baseweb="tab-list"] {{ gap:1.2rem; border-bottom:1px solid {LINE}; }}
  .stTabs [data-baseweb="tab"] {{ font-size:.85rem; color:{MUTED}; padding:.4rem 0; }}
  .stTabs [aria-selected="true"] {{ color:{TEXT}; border-bottom-color:{MAGENTA} !important; }}
  .stButton>button[kind="primary"] {{ background:{MAGENTA}; border-color:{MAGENTA}; }}
  .stButton>button[kind="primary"]:hover {{ background:#c00161; border-color:#c00161; }}
  #MainMenu, footer {{ visibility:hidden; }}
  @media (prefers-reduced-motion: reduce) {{ * {{ transition:none !important; animation:none !important; }} }}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner=False)
def load(name, data):
    # data — байты xlsx (из файла на диске или загруженного через сайдбар).
    # Само содержимое служит ключом кэша: пересобрали файл — кэш сбросится сам,
    # отдельный аргумент со временем изменения больше не нужен.
    df = pd.read_excel(io.BytesIO(data))
    for c in ['BALANCE_START', 'CHARGE', 'PAYMENT', 'BALANCE_END', 'S']:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
    for c in ['FINAL', 'Один договор (номер+контрагент, точно)',
              'Один договор (по номеру, допуск)', 'Один договор (по контрагенту, допуск)',
              NOPAY_COL, 'Нужно разобрать']:
        if c in df.columns:
            df[c] = df[c].astype(str) == 'True'
    # «Сумма по реестру» — справочно, у большей части строк not found
    df['Реестр'] = pd.to_numeric(df['Sum reestr'], errors='coerce')
    df['Тип оплаты'] = df['Тип оплаты'].astype(str).replace({'False': '—', 'nan': '—'})
    df['Особый случай'] = df['Особый случай'].astype(str).replace({'False': '—', 'nan': '—'})
    # область определяем сами по первой цифре кода сайта (N до первого "_") —
    # в самом файле результата колонки "Область" может и не быть, но нам она
    # и не нужна: сопоставление цифра -> область задано вручную (REGION_MAP)
    if 'Область' in df.columns:
        df['Регион'] = df['Область'].astype(str).str.replace(' область', '', regex=False)
    else:
        df['Область'] = region_of(df['N'])
        df['Регион'] = df['Область'].astype(str).str.replace(' область', '', regex=False)

    # в контроль идут только фактические выплаты. нулевые ни на что не влияют,
    # отрицательные — ошибка системы.
    dropped = df[df['PAYMENT'] < 0]
    df = df[df['PAYMENT'] > 0].copy()

    # ── из чего состоит выплата ──
    # PAYMENT = BALANCE_START + CHARGE − BALANCE_END, поэтому выплату всегда можно
    # разложить: сначала гасится входящий долг, затем текущее начисление, остаток —
    # аванс (он же отрицательное сальдо на конец).
    debt = df['BALANCE_START'].clip(lower=0).clip(upper=df['PAYMENT'])
    adv = (-df['BALANCE_END']).clip(lower=0)
    adv = np.minimum(adv, df['PAYMENT'] - debt)
    df['Погашение долга'] = debt
    df['Аванс'] = adv
    df['Оплата месяца'] = df['PAYMENT'] - debt - adv

    el = df['Electro or not'].eq('Электро')
    tech = df['Electro or not'].eq('Техобслуживание сайтов')
    gr = df['Групповые'].astype(str).eq('Групповой')
    typed = df['Тип оплаты'].ne('—')
    osob = df['Особый случай'].ne('—')

    # ── основания «сходится с реестром» — по порядку, как в build_result.py:
    # FINAL -> RAN Sharing/Проверка 100050 -> Проверка НДС -> Один договор (3
    # варианта) -> Особый случай -> Тип оплаты (у него ещё и своя разбивка) ──
    basis = pd.Series('—', index=df.index, dtype=object)
    ran = df['RAN Sharing'].astype(str).eq('Runsharing') if 'RAN Sharing' in df.columns else pd.Series(False, index=df.index)
    p100050 = df['Проверка 100050'].astype(str).eq('Perfect') if 'Проверка 100050' in df.columns else pd.Series(False, index=df.index)
    pnds = df['Проверка НДС (1,14)'].astype(str).eq('Perfect') if 'Проверка НДС (1,14)' in df.columns else pd.Series(False, index=df.index)
    final = df['FINAL'] if 'FINAL' in df.columns else pd.Series(False, index=df.index)
    tochno = df['Один договор (номер+контрагент, точно)'] if 'Один договор (номер+контрагент, точно)' in df.columns else pd.Series(False, index=df.index)
    po_nomeru = df['Один договор (по номеру, допуск)'] if 'Один договор (по номеру, допуск)' in df.columns else pd.Series(False, index=df.index)
    po_kontr = df['Один договор (по контрагенту, допуск)'] if 'Один договор (по контрагенту, допуск)' in df.columns else pd.Series(False, index=df.index)

    for cond, name in [(final, MATCHED_FINAL), (ran, MATCHED_RAN), (p100050, MATCHED_100050),
                        (pnds, MATCHED_NDS), (tochno, MATCHED_NOMER), (po_nomeru, MATCHED_PO_NOMERU),
                        (po_kontr, MATCHED_PO_KONTR)]:
        take = cond & basis.eq('—')
        basis[take] = name
    take = osob & basis.eq('—')
    basis[take] = np.where(df.loc[take, 'Особый случай'].str.startswith('Новый'),
                            MATCHED_OSOBYI_NOVYI, MATCHED_OSOBYI_SMENA)
    take = typed & basis.eq('—')
    basis[take] = df.loc[take, 'Тип оплаты']
    df['Основание'] = basis
    matched = basis.ne('—')

    # ── «платить не следовало» — реальная колонка build_result.py ──
    if NOPAY_COL in df.columns:
        df[NOPAY] = df[NOPAY_COL].astype(bool)
    else:
        df[NOPAY] = False
    df['Метка выгрузки'] = np.where(df[NOPAY], 'Не следовало платить', '—')

    # электро/техобслуживание/групповые контролируются отдельно — забирают
    # строку раньше реестра, иначе одна выплата попала бы в два контура.
    # Всё, что раньше попадало в "электро + групповые", теперь целиком в "Электро".
    grp = pd.Series(REVIEW, index=df.index, dtype=object)
    grp[matched] = REESTR
    grp[gr & ~el & ~tech] = GR_ONLY
    grp[tech & ~el] = TECH_ONLY
    grp[el] = EL_ONLY
    grp[(~el) & (~tech) & (~gr) & (~matched) & df[NOPAY]] = REVIEW  # платить не следовало -> тоже на разбор, но с явной пометкой в таблице
    df['Группа'] = grp

    # стабильный ключ выплаты — переживает пересборку выгрузки
    df['key'] = (df['Номер договора'].astype(str) + '|' + df['N'].astype(str)
                 + '|' + df['PAYMENT'].astype(str) + '|' + df['DAT'].astype(str))
    df['ключ договора'] = df['Номер договора'].astype(str)

    first = pd.to_datetime(df['DAT'].iloc[0], dayfirst=True, errors='coerce')
    meta = {'файл': name,
            'период': period_label(df['DAT'].iloc[0]),
            # ключ сортировки: месяцы должны идти по календарю, а не по алфавиту
            'порядок': pd.Timestamp.min if pd.isna(first)
                       else pd.Timestamp(first.year, first.month, 1),
            'отброшено': len(dropped),
            'возобновилось': 0,
            'сумма возобновилось': 0.0}
    return df, meta


@st.cache_data(show_spinner=False)
def load_months(sources):
    """Читает все выгрузки и раскладывает их по периодам.

    sources — кортеж пар (имя, байты). Откуда взялся файл, не важно: лежит рядом
    со скриптом или загружен через сайдбар — обрабатывается одинаково. Ключом кэша
    выступает само содержимое файла, поэтому при замене выгрузки (build_result.py
    пересобрал её) цифры обновятся автоматически, без ручной чистки кэша.
    """
    out, broken = {}, []
    for name, data in sources:
        try:
            df, meta = load(name, data)
        except (ValueError, KeyError) as e:
            broken.append(f'{name}: {e}')
            continue
        out[meta['период']] = (df, meta)
    return dict(sorted(out.items(), key=lambda kv: kv[1][1]['порядок'])), broken


@st.cache_data(show_spinner=False)
def load_geo(path):
    if not Path(path).exists():
        return None
    with open(path, encoding='utf-8') as fh:
        return json.load(fh)


def blank_decisions():
    return {'payments': {}, 'contracts': {}}


def decisions_to_frames(dec):
    """payments/contracts (словарь) -> два DataFrame, по одному на лист Excel."""
    pay = pd.DataFrame(
        [{'key': k, 'Решение': v.get('решение', ''), 'Комментарий': v.get('комментарий', '')}
         for k, v in dec['payments'].items()],
        columns=['key', 'Решение', 'Комментарий'])
    con = pd.DataFrame(
        [{'ключ договора': k, 'Решение': v.get('решение', ''), 'Комментарий': v.get('комментарий', '')}
         for k, v in dec['contracts'].items()],
        columns=['ключ договора', 'Решение', 'Комментарий'])
    return pay, con


def frames_to_decisions(pay_df, con_df):
    payments = {str(r['key']): {'решение': r['Решение'], 'комментарий': r.get('Комментарий', '') or ''}
                for _, r in pay_df.iterrows()} if len(pay_df) else {}
    contracts = {str(r['ключ договора']): {'решение': r['Решение'], 'комментарий': r.get('Комментарий', '') or ''}
                 for _, r in con_df.iterrows()} if len(con_df) else {}
    return {'payments': payments, 'contracts': contracts}


def read_decisions(src=None):
    """Решения хранятся в отдельном Excel-файле — два листа, «Выплаты» и
    «Договоры» — чтобы бухгалтерия могла открыть и посмотреть/поправить
    их напрямую, а не только через дашборд.

    src — путь или файл-объект (например, загруженный через сайдбар); по
    умолчанию берётся файл рядом со скриптом."""
    if src is None:
        if not DECISIONS_FILE.exists():
            return blank_decisions()
        src = DECISIONS_FILE
    try:
        pay_df = pd.read_excel(src, sheet_name='Выплаты')
        con_df = pd.read_excel(src, sheet_name='Договоры')
        return frames_to_decisions(pay_df, con_df)
    except (ValueError, OSError, KeyError):
        return blank_decisions()


def write_decisions(dec):
    # На облачном хостинге файловая система эфемерна (сбрасывается при перезапуске),
    # а иногда доступна только на чтение — поэтому запись не должна ронять приложение.
    # Решения всегда живут в session_state; файл рядом со скриптом — лишь копия для
    # локальной работы. Для надёжного сохранения есть кнопки скачать/загрузить в сайдбаре.
    try:
        pay_df, con_df = decisions_to_frames(dec)
        with pd.ExcelWriter(DECISIONS_FILE, engine='openpyxl') as w:
            pay_df.to_excel(w, sheet_name='Выплаты', index=False)
            con_df.to_excel(w, sheet_name='Договоры', index=False)
    except OSError:
        pass


def decisions_bytes(dec):
    pay_df, con_df = decisions_to_frames(dec)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        pay_df.to_excel(w, sheet_name='Выплаты', index=False)
        con_df.to_excel(w, sheet_name='Договоры', index=False)
    return buf.getvalue()


def money(x):
    return f'{x:,.0f}'.replace(',', ' ')


def mln(x):
    return f'{x / 1e6:,.1f}'.replace(',', ' ').replace('.', ',') + ' млн'


def period_label(v):
    # DAT приходит то текстом «31.05.2026», то датой — период не должен от этого зависеть
    d = pd.to_datetime(v, dayfirst=True, errors='coerce')
    return str(v) if pd.isna(d) else f'{MONTHS[d.month]} {d.year}'


def logo_b64():
    if LOGO_FILE is not None and LOGO_FILE.exists():
        return base64.b64encode(LOGO_FILE.read_bytes()).decode()
    return None


LOGO = logo_b64()
LOGO_HTML = f'<img src="data:image/png;base64,{LOGO}" style="height:30px">' if LOGO else ''


def masthead(right=''):
    st.markdown(f'<div class="masthead">{LOGO_HTML}'
                '<span class="mark">КОНТРОЛЬ ВЫПЛАТ ПО ДОГОВОРАМ</span>'
                f'<span class="date">{right}</span></div>', unsafe_allow_html=True)


# ── откуда берём выгрузки ──
# Сайдбар с логотипом и загрузкой рисуем здесь, до чтения данных: сначала нужно
# получить файлы, а уже потом строить всё остальное. Логотип рисуем один раз тут.
st.sidebar.markdown(
    f'<div style="display:flex;align-items:center;gap:.5rem;margin-bottom:1rem">'
    f'{LOGO_HTML}<span style="font-family:\'IBM Plex Mono\',monospace;'
    f'font-size:.7rem;letter-spacing:.12em;color:{MUTED}">КОНТРОЛЬ ВЫПЛАТ</span></div>',
    unsafe_allow_html=True)
st.sidebar.markdown(
    f"<div style='font-family:IBM Plex Mono,monospace;font-size:.72rem;"
    f"letter-spacing:.18em;color:{MUTED};margin:.2rem 0 .4rem'>ВЫГРУЗКИ</div>",
    unsafe_allow_html=True)
uploaded = st.sidebar.file_uploader(
    'Файлы «Результат … .xlsx»', type='xlsx', accept_multiple_files=True,
    label_visibility='collapsed',
    help='Можно выбрать сразу несколько файлов за разные месяцы. Период берётся '
         'из колонки DAT внутри файла, а не из имени.')

# сначала то, что загрузили руками, затем файлы из локальной папки (если есть).
# в облаке папки обычно нет — работаем на загрузке; локально удобнее из папки.
sources = [(uf.name, uf.getvalue()) for uf in (uploaded or [])
           if not uf.name.startswith('~$')]
if DATA_FOLDER.exists():
    have = {n for n, _ in sources}
    for p in sorted(DATA_FOLDER.glob(DATA_GLOB)):
        if not p.name.startswith('~$') and p.name not in have:
            sources.append((p.name, p.read_bytes()))

if not sources:
    masthead()
    st.info(f'Загрузите выгрузки вида **{DATA_GLOB}** через панель слева '
            f'(раздел «Выгрузки»). Можно выбрать сразу несколько файлов за разные месяцы.')
    st.stop()

MONTHS_DATA, broken = load_months(tuple(sources))
if not MONTHS_DATA:
    masthead()
    st.error('Ни одну выгрузку не удалось прочитать: ' + '; '.join(broken))
    st.stop()

geo = load_geo(str(GEOJSON_FILE))
PERIODS = list(MONTHS_DATA)

if 'dec' not in st.session_state:
    st.session_state.dec = read_decisions()
dec = st.session_state.dec


# ─────────── ручные решения поверх автоматических групп ───────────
def apply_decisions(base, dec):
    d = base.copy()

    def field(series, name):
        return series.map(lambda v: v.get(name, '') if isinstance(v, dict) else '')

    per_pay = d['key'].map(dec['payments'])
    per_con = d['ключ договора'].map(dec['contracts'])

    reason = field(per_pay, 'решение')
    note = field(per_pay, 'комментарий')
    from_con = field(per_con, 'решение')
    take = (reason == '') & (from_con != '')
    reason[take] = from_con[take]
    note[take] = field(per_con, 'комментарий')[take]

    d['Решение'] = reason
    d['Комментарий'] = note

    closed = d['Решение'].isin(CLOSING) & d['Группа'].eq(REVIEW)
    d.loc[closed, 'Группа'] = REESTR
    d.loc[closed, 'Основание'] = d.loc[closed, 'Решение'].map(MANUAL_BASIS_MAP)
    d['Разобрано вручную'] = closed
    d['Открытый вопрос'] = d['Решение'].eq(OPEN)
    return d


# решения — общие для всех месяцев: правило по договору, поставленное в мае,
# автоматически применяется к июню
MONTHS_DATA = {p: (apply_decisions(d, dec), m) for p, (d, m) in MONTHS_DATA.items()}
ALL_REGIONS = sorted({r for d, _ in MONTHS_DATA.values() for r in d['Регион'].dropna()},
                     key=lambda r: REGION_ORDER.index(r) if r in REGION_ORDER else 99)


HAS_REGIONS = len(ALL_REGIONS) > 1 or (len(ALL_REGIONS) == 1 and ALL_REGIONS[0] != '—')


def side_title(text):
    st.markdown(f"<div style='font-family:IBM Plex Mono,monospace;font-size:.72rem;"
                f"letter-spacing:.18em;color:{MUTED};margin:.2rem 0 .6rem'>{text}</div>",
                unsafe_allow_html=True)


# ─────────────────────────── период и фильтры ───────────────────────────
with st.sidebar:
    side_title('ПЕРИОД')
    period = st.selectbox('Месяц', PERIODS, index=len(PERIODS) - 1,
                          label_visibility='collapsed')
    pos = PERIODS.index(period)
    prev_period = PERIODS[pos - 1] if pos > 0 else None
    st.markdown(f"<div class='note'>Найдено выгрузок: {len(PERIODS)} — "
                f"{', '.join(PERIODS)}.</div>", unsafe_allow_html=True)
    if broken:
        st.warning('Не прочитано: ' + '; '.join(broken))

    st.markdown('---')
    side_title('ФИЛЬТРЫ')
    regions = st.multiselect('Область', ALL_REGIONS, default=[]) if HAS_REGIONS else []
    groups = st.multiselect('Группа', GROUPS, default=[])
    search = st.text_input('Поиск по сайту, договору или контрагенту', '')
    st.markdown("<div class='note' style='margin-top:.8rem'>Пустой фильтр — значит все. "
                "Фильтры действуют и на сравнение периодов.</div>", unsafe_allow_html=True)

    st.markdown('---')
    side_title('РЕШЕНИЯ')
    st.download_button(f'Скачать {DECISIONS_NAME}', decisions_bytes(dec), file_name=DECISIONS_NAME,
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    dec_upload = st.file_uploader(f'Загрузить {DECISIONS_NAME}', type='xlsx',
                                  key='dec_upload', label_visibility='collapsed')
    if dec_upload is not None and st.session_state.get('dec_upload_name') != dec_upload.name:
        st.session_state.dec = read_decisions(dec_upload)
        st.session_state.dec_upload_name = dec_upload.name
        write_decisions(st.session_state.dec)
        st.rerun()
    st.markdown("<div class='note' style='margin-top:.4rem'>На облачном хостинге решения "
                "сбрасываются при перезапуске сервиса. Скачивайте файл, чтобы сохранить "
                "работу, и загружайте обратно, чтобы продолжить с того же места.</div>",
                unsafe_allow_html=True)


def sift(d):
    """Один и тот же фильтр для месяца и для сравнения — иначе цифры разъедутся."""
    if regions:
        d = d[d['Регион'].isin(regions)]
    if groups:
        d = d[d['Группа'].isin(groups)]
    if search:
        s = search.lower()
        d = d[d['N'].astype(str).str.lower().str.contains(s)
              | d['CONTRACTOR'].astype(str).str.lower().str.contains(s)
              | d['Номер договора'].astype(str).str.contains(s)]
    return d


df, meta = MONTHS_DATA[period]
n_dropped = meta['отброшено']
n_resumed, sum_resumed = meta['возобновилось'], meta['сумма возобновилось']
f = sift(df)

masthead(f'период · {period}'
         + (f'  ·  предыдущий · {prev_period}' if prev_period else ''))

if f.empty:
    st.warning('Под фильтры не попало ни одной выплаты. Снимите часть условий слева.')
    st.stop()

# ─────────────────────────── шесть квадратов ───────────────────────────
paid = f['PAYMENT'].sum()
by_group = f.groupby('Группа')['PAYMENT'].agg(['sum', 'size']).reindex(GROUPS).fillna(0)


prev_f = sift(MONTHS_DATA[prev_period][0]) if prev_period else None
prev_paid = prev_f['PAYMENT'].sum() if prev_f is not None else None
prev_group = (prev_f.groupby('Группа')['PAYMENT'].agg(['sum', 'size'])
              .reindex(GROUPS).fillna(0) if prev_f is not None else None)


def delta_note(cur, prev, worse_when_up=False):
    """Подпись «сколько к прошлому месяцу» — без неё цифра месяца ни о чём не говорит."""
    if prev is None:
        return ''
    d = cur - prev
    if abs(d) < 1:
        return f'<span style="color:{MUTED}">без изменений к {prev_period}</span>'
    sign, color = ('+', MUTED) if d > 0 else ('−', MUTED)
    if worse_when_up:
        color = WARN if d > 0 else OK
    pct = f' ({sign}{abs(d) / prev * 100:.0f}%)' if prev else ''
    return f'<span style="color:{color}">{sign}{money(abs(d))}{pct} к {prev_period}</span>'


def kpi(col, label, value, sub, color=LINE, unit='сом', extra=''):
    col.markdown(f'<div class="kpi" style="border-left-color:{color}">'
                 f'<div class="label">{label}</div>'
                 f'<div class="value">{value} <span class="unit">{unit}</span></div>'
                 f'<div class="sub">{sub}</div>'
                 + (f'<div class="sub">{extra}</div>' if extra else '')
                 + '</div>', unsafe_allow_html=True)


def mini(col, label, value, sub, color=LINE):
    col.markdown(f'<div class="mini" style="border-left-color:{color}">'
                 f'<div class="label">{label}</div>'
                 f'<div class="value">{value}</div>'
                 f'<div class="sub">{sub}</div></div>', unsafe_allow_html=True)


c = st.columns(6)
kpi(c[0], 'Всего выплачено за месяц', money(paid), f'{len(f)} выплат', MUTED,
    extra=delta_note(paid, prev_paid))
for col, g in zip(c[1:], GROUPS):
    s, n = by_group.loc[g, 'sum'], int(by_group.loc[g, 'size'])
    prev_s = prev_group.loc[g, 'sum'] if prev_group is not None else None
    kpi(col, g, money(s), f'{s / paid * 100:.1f}% суммы · {n} выплат' if paid else '',
        GROUP_COLOR[g], extra=delta_note(s, prev_s, worse_when_up=(g == REVIEW)))

notes = [f'Электроэнергия, техобслуживание и групповые сайты контролируются отдельно — здесь '
         f'только суммы. Если сайт одновременно и электро, и групповой — он в «{EL_ONLY}», '
         f'чтобы не считаться дважды. Группы не пересекаются и складываются в {money(paid)} сом.']
if n_resumed:
    notes.append(f'Правило «{RESUMED.lower()}» (сальдо на конец = сумма по реестру, начисление '
                 f'месяца остаётся на следующий) закрыло {n_resumed} выплат на '
                 f'{money(sum_resumed)} сом — они в реестре, а не на проверке.')
if n_dropped:
    notes.append(f'Отброшено строк с отрицательной выплатой: {n_dropped} (ошибка системы).')
n_nc = int(f[NOPAY].sum())
if n_nc:
    notes.append(f'Отдельно: {n_nc} выплат на {money(f.loc[f[NOPAY], "PAYMENT"].sum())} сом '
                 f'отмечены в выгрузке как «{NOPAY_COL.lower()}» — вкладка «{NOPAY}».')
st.markdown(f'<div class="note" style="margin-top:.7rem">{" ".join(notes)}</div>',
            unsafe_allow_html=True)

# ─────────────────────────── карта ───────────────────────────
reg = (f.pivot_table(index='Регион', columns='Группа', values='PAYMENT',
                     aggfunc='sum', fill_value=0)
       .reindex(columns=GROUPS, fill_value=0))
reg['Выплачено'] = reg.sum(axis=1)
reg['Начислено'] = f.groupby('Регион')['CHARGE'].sum()
reg_n = (f.pivot_table(index='Регион', columns='Группа', values='PAYMENT',
                       aggfunc='count', fill_value=0)
         .reindex(columns=GROUPS, fill_value=0))
reg = reg.reset_index()

if geo is not None and HAS_REGIONS:
    mcol, ctrl = st.columns([5, 1])
    with ctrl:
        st.markdown("<div style='height:2.6rem'></div>", unsafe_allow_html=True)
        metric = st.radio('Показать на карте', [REESTR, EL_ONLY, REVIEW],
                          index=0, label_visibility='collapsed')

    scale = {REESTR: [[0, PANEL], [1, OK]],
             EL_ONLY: [[0, PANEL], [1, COOL]],
             REVIEW: [[0, PANEL], [1, WARN]]}[metric]

    ids = {ft['properties']['region'] for ft in geo['features']}
    on_map = reg[reg['Регион'].isin(ids)].copy()
    on_map['_n'] = on_map['Регион'].map(reg_n[metric])

    # во всплывающей подсказке — только выбранная в легенде метрика:
    # сумма и количество выплат именно по ней, без остальных трёх
    fig_map = go.Figure(go.Choropleth(
        geojson=geo, locations=on_map['Регион'], featureidkey='properties.region',
        z=on_map[metric], colorscale=scale, marker_line_color=TEXT, marker_line_width=1,
        colorbar=dict(title='', thickness=10, len=.6, x=.98,
                      tickfont=dict(color=MUTED, size=11)),
        customdata=on_map[['_n']].values,
        hovertemplate='<b>%{location}</b><br>' + metric + ' %{z:,.0f} сом<br>'
                      '%{customdata[0]:.0f} выплат<extra></extra>'))
    fig_map.update_geos(fitbounds='locations', visible=False, bgcolor='rgba(0,0,0,0)')
    fig_map.update_layout(height=380, margin=dict(l=0, r=0, t=0, b=0),
                          paper_bgcolor='rgba(0,0,0,0)', font=dict(color=TEXT),
                          dragmode=False)
    with mcol:
        st.markdown(f'<div class="section-title">Карта областей · {metric.lower()}, сом</div>',
                    unsafe_allow_html=True)
        st.plotly_chart(fig_map, config={'displayModeBar': False, 'scrollZoom': False})

        off = reg[~reg['Регион'].isin(ids)].copy()
        off['_n'] = off['Регион'].map(reg_n[metric])
        if not off.empty:
            chips = '  ·  '.join(
                f"<b style='color:{TEXT}'>{r['Регион']}</b> "
                f"<span style='color:{MUTED}'>{metric.lower()} {money(r[metric])} сом, "
                f"{int(r['_n'])} выплат</span>" for _, r in off.iterrows())
            st.markdown(f"<div class='note'>вне карты — {chips}</div>",
                        unsafe_allow_html=True)
elif not HAS_REGIONS:
    st.markdown('<div class="note">Карта по областям скрыта: в файлах результата нет колонки '
                '«Область» (build_result.py её пока не формирует), поэтому реальных региональных '
                'данных для карты нет. Как только колонка появится в отчёте — карта включится сама.</div>',
                unsafe_allow_html=True)


def dark(fig, h=340, stack=None, legend_top=True, right=8):
    fig.update_layout(height=h, paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                      font=dict(color=MUTED, size=12, family='IBM Plex Sans, sans-serif'),
                      margin=dict(l=8, r=right, t=38 if legend_top else 8, b=8),
                      barmode=stack or 'group', bargap=.28,
                      legend=dict(orientation='h', y=1.16 if legend_top else -.12, x=0,
                                  yanchor='bottom', font=dict(color=MUTED)),
                      xaxis=dict(gridcolor=LINE, zerolinecolor=LINE, title=None,
                                 automargin=True),
                      yaxis=dict(gridcolor=LINE, zerolinecolor=LINE, title=None,
                                 automargin=True))
    st.plotly_chart(fig, config={'displayModeBar': False})


# ─────────────────────────── по регионам ───────────────────────────
left, right = st.columns([3, 2])

MAP_METRICS = [REESTR, EL_ONLY, REVIEW]

with left:
    st.markdown('<div class="section-title">Выплаты по областям · по реестру, электро, '
                'на проверку · сом</div>', unsafe_allow_html=True)
    r1 = reg.sort_values('Выплачено')
    fig = go.Figure()
    for g in MAP_METRICS:
        fig.add_bar(y=r1['Регион'], x=r1[g], orientation='h', name=g,
                    marker_color=GROUP_COLOR[g],
                    hovertemplate=g + ' %{x:,.0f} сом<extra></extra>')
    dark(fig, 380, stack='stack')

with right:
    st.markdown('<div class="section-title">Начислено против выплачено · по областям, сом</div>',
                unsafe_allow_html=True)
    r2 = reg.sort_values('Начислено')
    fig2 = go.Figure()
    fig2.add_bar(y=r2['Регион'], x=r2['Начислено'], orientation='h', name='начислено',
                 marker_color=MUTED, hovertemplate='начислено %{x:,.0f}<extra></extra>')
    fig2.add_bar(y=r2['Регион'], x=r2['Выплачено'], orientation='h', name='выплачено',
                 marker_color=MAGENTA, width=.45,
                 hovertemplate='выплачено %{x:,.0f}<extra></extra>')
    dark(fig2, 380, stack='overlay')

# ─────────────────────────── кого разбирать ───────────────────────────
st.markdown('<div class="section-title">Кого разбирать первым · сумма выплат '
            'на проверку</div>', unsafe_allow_html=True)
top = (f[f['Группа'] == REVIEW].groupby('CONTRACTOR')['PAYMENT']
       .agg(['sum', 'size']).sort_values('sum', ascending=False).head(12))
if top.empty:
    st.markdown('<div class="note">Под этими фильтрами разобраны все выплаты.</div>',
                unsafe_allow_html=True)
else:
    top = top.iloc[::-1]
    figs = go.Figure(go.Bar(
        x=top['sum'], y=[n[:44] for n in top.index], orientation='h', marker_color=MAGENTA,
        text=[money(v) for v in top['sum']], textposition='outside',
        textfont=dict(color=TEXT), customdata=top['size'],
        hovertemplate='%{x:,.0f} сом · %{customdata} выплат<extra></extra>'))
    figs.update_traces(cliponaxis=False)
    figs.update_xaxes(range=[0, top['sum'].max() * 1.22])
    dark(figs, 420, right=70)
    st.markdown(f'<div class="note">Топ-{len(top)} даёт {mln(top["sum"].sum())} сом '
                f'из {mln(by_group.loc[REVIEW, "sum"])} на проверку.</div>',
                unsafe_allow_html=True)

# ─────────────────────────── таблицы ───────────────────────────
# разложение выплаты считается всегда — оно нужно для сводки по реестру,
# но в построчных таблицах не показывается: руководителю важны сальдо и статус.
PARTS = ['Погашение долга', 'Оплата месяца', 'Аванс']
COLS = ['N', 'CONTRACTOR', 'Номер договора', 'Область', 'BALANCE_START', 'CHARGE',
        'PAYMENT', 'BALANCE_END', 'Группа', 'Основание', 'Комментарий']
NAMES = {'N': 'Сайт', 'CONTRACTOR': 'Контрагент', 'Номер договора': 'Договор',
         'BALANCE_START': 'Сальдо на начало', 'CHARGE': 'Начислено',
         'PAYMENT': 'Выплата', 'BALANCE_END': 'Сальдо на конец'}


def table(data, height=420, key='', extra=()):
    view = data[COLS + list(extra)].rename(columns=NAMES)
    st.dataframe(view, hide_index=True, height=height)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        view.to_excel(w, sheet_name='Выплаты', index=False)
    st.download_button('Скачать в Excel', buf.getvalue(), file_name='Выплаты.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                       key=f'dl{key}{height}{len(data)}')


# ─────────────────────────── сравнение периодов ───────────────────────────
def contract_roll(d):
    """Свод по договору: сколько выплачено и сколько из этого висит на проверке."""
    d = d.assign(_rev=d['PAYMENT'].where(d['Группа'].eq(REVIEW), 0))
    return d.groupby('ключ договора').agg(
        Контрагент=('CONTRACTOR', 'first'), Область=('Регион', 'first'),
        Выплата=('PAYMENT', 'sum'), Начислено=('CHARGE', 'sum'),
        Проверка=('_rev', 'sum'), Выплат=('PAYMENT', 'size'))


def money_cols(d, cols):
    d = d.copy()
    for c in cols:
        d[c] = d[c].map(money)
    return d


def xlsx_button(frames, name, key):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as w:
        for sheet, fr in frames.items():
            fr.to_excel(w, sheet_name=sheet[:31], index=False)
    st.download_button('Скачать в Excel', buf.getvalue(), file_name=name, key=key,
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def compare(a_label, b_label):
    """Сравнение двух периодов: что выросло, кто появился и кто на проверке не первый месяц."""
    a, b = sift(MONTHS_DATA[a_label][0]), sift(MONTHS_DATA[b_label][0])
    if a.empty or b.empty:
        st.warning('Под текущие фильтры в одном из периодов не попало ни одной выплаты.')
        return

    ga = a.groupby('Группа')['PAYMENT'].agg(['sum', 'size']).reindex(GROUPS).fillna(0)
    gb = b.groupby('Группа')['PAYMENT'].agg(['sum', 'size']).reindex(GROUPS).fillna(0)
    ca, cb = contract_roll(a), contract_roll(b)
    chronic_keys = ca.index[ca['Проверка'] > 0].intersection(cb.index[cb['Проверка'] > 0])
    chronic_sum = cb.loc[chronic_keys, 'Проверка'].sum()

    # ── вердикт одной фразой ──
    d_all = b['PAYMENT'].sum() - a['PAYMENT'].sum()
    d_rev = gb.loc[REVIEW, 'sum'] - ga.loc[REVIEW, 'sum']
    verdict = (f'Выплачено {money(b["PAYMENT"].sum())} сом против '
               f'{money(a["PAYMENT"].sum())} — {"+" if d_all >= 0 else "−"}{money(abs(d_all))}. '
               f'На проверке {money(gb.loc[REVIEW, "sum"])} сом против '
               f'{money(ga.loc[REVIEW, "sum"])} — {"+" if d_rev >= 0 else "−"}'
               f'{money(abs(d_rev))}.')
    if len(chronic_keys):
        verdict += (f' Из них {money(chronic_sum)} сом приходится на {len(chronic_keys)} '
                    f'договоров, которые были на проверке и в периоде {a_label}, — '
                    f'это системная история, а не разовый сбой.')
    st.markdown(f'<div class="note" style="margin:.5rem 0 1rem;color:{TEXT}">{verdict}</div>',
                unsafe_allow_html=True)

    # ── свод по группам ──
    st.markdown('<div class="section-title">Свод по группам · млн сом</div>',
                unsafe_allow_html=True)
    chart, tbl = st.columns([2, 3])
    with chart:
        fig = go.Figure()
        for lab, src, color in [(a_label, ga, MUTED), (b_label, gb, MAGENTA)]:
            fig.add_bar(x=[SHORT[g] for g in GROUPS],
                        y=[src.loc[g, 'sum'] / 1e6 for g in GROUPS], name=lab,
                        marker_color=color, customdata=[src.loc[g, 'sum'] for g in GROUPS],
                        hovertemplate='%{customdata:,.0f} сом<extra>' + lab + '</extra>')
        fig.update_layout(xaxis_tickangle=0, xaxis_tickfont=dict(size=11))
        dark(fig, 360, legend_top=True)
    with tbl:
        rows = []
        for g in GROUPS + ['ИТОГО']:
            if g == 'ИТОГО':
                sa, sb = ga['sum'].sum(), gb['sum'].sum()
                na, nb = ga['size'].sum(), gb['size'].sum()
            else:
                sa, sb, na, nb = (ga.loc[g, 'sum'], gb.loc[g, 'sum'],
                                  ga.loc[g, 'size'], gb.loc[g, 'size'])
            rows.append({'Группа': g, a_label: money(sa), b_label: money(sb),
                         'Δ сом': ('+' if sb >= sa else '−') + money(abs(sb - sa)),
                         'Δ %': f'{(sb - sa) / sa * 100:+.0f}%' if sa else '—',
                         f'выплат · {a_label}': int(na), f'выплат · {b_label}': int(nb)})
        st.dataframe(pd.DataFrame(rows), hide_index=True, height=340)

    # ── платить не следовало ──
    an, bn = a[a[NOPAY]], b[b[NOPAY]]
    st.markdown(f'<div class="section-title">{NOPAY}</div>', unsafe_allow_html=True)
    m = st.columns(4)
    mini(m[0], a_label, money(an['PAYMENT'].sum()), f'{len(an)} выплат', MUTED)
    mini(m[1], b_label, money(bn['PAYMENT'].sum()), f'{len(bn)} выплат', WARN)
    d_n = bn['PAYMENT'].sum() - an['PAYMENT'].sum()
    mini(m[2], 'Δ', ('+' if d_n >= 0 else '−') + money(abs(d_n)),
         f'{len(bn) - len(an):+d} выплат', WARN if d_n > 0 else OK)
    rep = len(set(an['ключ договора']) & set(bn['ключ договора']))
    mini(m[3], 'Повторяются', str(rep), f'договоров и в {a_label}', WARN if rep else OK)
    st.markdown(f'<div class="note" style="margin-top:.5rem">Выплаты, отмеченные в выгрузке '
                f'как «{NOPAY_COL.lower()}». Договор, который повторяется здесь второй '
                f'период, — это не ошибка оператора, а незакрытый договор в системе. '
                f'Разбор во вкладке «{NOPAY}».</div>', unsafe_allow_html=True)

    # ── договоры, которые на проверке не первый месяц ──
    st.markdown('<div class="section-title">На проверке второй период подряд · '
                'с них начинать разбор</div>', unsafe_allow_html=True)
    if not len(chronic_keys):
        st.success(f'Ни один договор не остался на проверке из {a_label} в {b_label}.')
    else:
        chronic = pd.DataFrame({
            'Договор': chronic_keys,
            'Контрагент': cb.loc[chronic_keys, 'Контрагент'].values,
            'Область': cb.loc[chronic_keys, 'Область'].values,
            a_label: ca.loc[chronic_keys, 'Проверка'].values,
            b_label: cb.loc[chronic_keys, 'Проверка'].values})
        chronic['Итого за два периода'] = chronic[a_label] + chronic[b_label]
        chronic = chronic.sort_values('Итого за два периода', ascending=False)
        st.markdown(f'<div class="note" style="margin-bottom:.6rem">{len(chronic)} договоров '
                    f'на {money(chronic["Итого за два периода"].sum())} сом за два периода. '
                    f'Повтор означает, что дело в условиях договора или в реестре, '
                    f'а не в разовой выплате.</div>', unsafe_allow_html=True)
        st.dataframe(money_cols(chronic, [a_label, b_label, 'Итого за два периода']),
                     hide_index=True, height=320)
        xlsx_button({'На проверке подряд': chronic}, 'Хронические_на_проверку.xlsx', 'dlchr')

    # ── кто появился и кто пропал ──
    st.markdown('<div class="section-title">Кто появился и кто пропал</div>',
                unsafe_allow_html=True)
    ka = a.groupby('CONTRACTOR')['PAYMENT'].sum()
    kb = b.groupby('CONTRACTOR')['PAYMENT'].sum()
    new = kb[~kb.index.isin(ka.index)].sort_values(ascending=False)
    gone = ka[~ka.index.isin(kb.index)].sort_values(ascending=False)
    l, r = st.columns(2)
    with l:
        mini(st.container(), f'Новые в {b_label}', money(new.sum()),
             f'{len(new)} контрагентов', OK)
        if len(new):
            st.dataframe(money_cols(new.reset_index().rename(
                columns={'CONTRACTOR': 'Контрагент', 'PAYMENT': 'Выплата'}).head(15),
                ['Выплата']), hide_index=True, height=260)
    with r:
        mini(st.container(), f'Были в {a_label}, в {b_label} выплат нет', money(gone.sum()),
             f'{len(gone)} контрагентов', WARN)
        if len(gone):
            st.dataframe(money_cols(gone.reset_index().rename(
                columns={'CONTRACTOR': 'Контрагент', 'PAYMENT': 'Выплата'}).head(15),
                ['Выплата']), hide_index=True, height=260)
    st.markdown('<div class="note">Пропавший контрагент — это либо закрытый договор, '
                'либо пропущенная выплата. Второе тоже нужно объяснить.</div>',
                unsafe_allow_html=True)

    # ── резкие изменения по договорам ──
    st.markdown('<div class="section-title">Сильнее всего изменились выплаты по договору'
                '</div>', unsafe_allow_html=True)
    j = ca[['Контрагент', 'Выплата']].join(cb[['Выплата']], how='inner',
                                           lsuffix=f' {a_label}', rsuffix=f' {b_label}')
    j['Δ'] = j[f'Выплата {b_label}'] - j[f'Выплата {a_label}']
    j['Δ %'] = np.where(j[f'Выплата {a_label}'] > 0,
                        j['Δ'] / j[f'Выплата {a_label}'] * 100, np.nan)
    j = j.assign(_abs=j['Δ'].abs()).sort_values('_abs', ascending=False).head(15)
    j = j.reset_index().rename(columns={'ключ договора': 'Договор'})
    j['Δ %'] = j['Δ %'].map(lambda v: '—' if pd.isna(v) else f'{v:+.0f}%')
    j['Δ'] = j['Δ'].map(lambda v: ('+' if v >= 0 else '−') + money(abs(v)))
    st.dataframe(money_cols(j.drop(columns='_abs'),
                            [f'Выплата {a_label}', f'Выплата {b_label}']),
                 hide_index=True, height=320)
    st.markdown('<div class="note">Договоры, где сумма изменилась сильнее всего. '
                'Скачок при неизменном начислении — повод проверить основание.</div>',
                unsafe_allow_html=True)


tab1, tabn, tabc, tab2, tab3, tab4 = st.tabs(
    ['Разбор · на проверку', NOPAY, 'Сравнение периодов', 'Выплаты по реестру',
     'Электроэнергия, техобслуживание и групповые', 'Все выплаты'])

# ── выплаты, которых не должно было быть ──
with tabn:
    nch = f[f[NOPAY]].sort_values('PAYMENT', ascending=False)
    st.markdown(
        f'<div class="note" style="margin:.4rem 0 .8rem">Строки, отмеченные в выгрузке '
        f'в колонке «{NOPAY_COL}»: договора в реестре нет, платить по нему не должны были, '
        f'а выплата всё равно прошла. Это вопрос к бухгалтерии, а не к реестру — '
        f'поэтому вынесен отдельно от «{REVIEW}».</div>', unsafe_allow_html=True)

    if nch.empty:
        st.success('В этом периоде таких выплат нет.')
    else:
        prev_n = prev_f[prev_f[NOPAY]] if prev_f is not None else None
        by_con = (nch.groupby('CONTRACTOR')['PAYMENT'].agg(['sum', 'size'])
                  .sort_values('sum', ascending=False))
        repeat = by_con[by_con['size'] > 1]

        m = st.columns(4)
        mini(m[0], 'Всего', money(nch['PAYMENT'].sum()),
             f'{len(nch)} выплат · {nch["PAYMENT"].sum() / paid * 100:.2f}% месяца'
             if paid else f'{len(nch)} выплат', WARN)
        mini(m[1], 'Контрагентов', str(len(by_con)),
             f'повторно у {len(repeat)}' if len(repeat) else 'по одной выплате', WARN)
        el_n = nch[nch['Electro or not'].eq('Электро')]
        mini(m[2], 'Из них электроэнергия', money(el_n['PAYMENT'].sum()),
             f'{len(el_n)} выплат · отдельный контур', COOL)
        mini(m[3], 'Уже стоят на проверке', money(nch.loc[nch['Группа'].eq(REVIEW),
                                                          'PAYMENT'].sum()),
             f'{int(nch["Группа"].eq(REVIEW).sum())} из {len(nch)} выплат', VIOLET)

        if prev_n is not None:
            st.markdown(f'<div class="note" style="margin-top:.5rem">'
                        f'{delta_note(nch["PAYMENT"].sum(), prev_n["PAYMENT"].sum(), True)}'
                        f'</div>', unsafe_allow_html=True)

        if len(repeat):
            st.markdown('<div class="section-title">Повторяются в одном периоде</div>',
                        unsafe_allow_html=True)
            rp = repeat.reset_index().rename(
                columns={'CONTRACTOR': 'Контрагент', 'sum': 'Сумма', 'size': 'Выплат'})
            st.markdown('<div class="note" style="margin-bottom:.6rem">Один и тот же '
                        'контрагент получил несколько таких выплат за месяц — значит дело '
                        'не в разовой ошибке оператора.</div>', unsafe_allow_html=True)
            st.dataframe(money_cols(rp, ['Сумма']), hide_index=True, height=180)

        st.markdown('<div class="section-title">Все выплаты по метке</div>',
                    unsafe_allow_html=True)
        table(nch, 380, key='nopay', extra=['Метка выгрузки'])

with tabc:
    if len(PERIODS) < 2:
        st.info('Для сравнения нужны минимум две выгрузки. Положите файл соседнего месяца '
                f'({DATA_GLOB}) в папку с app.py и обновите страницу.')
    else:
        # по умолчанию — выбранный месяц против предыдущего; для самого раннего
        # месяца сравнивать не с чем назад, поэтому берём его и следующий
        i_b = PERIODS.index(period)
        i_a = i_b - 1
        if i_a < 0:
            i_a, i_b = 0, 1
        s1, s2, _ = st.columns([1, 1, 3])
        a_label = s1.selectbox('База · с чем сравниваем', PERIODS, index=i_a)
        b_label = s2.selectbox('Период · что сравниваем', PERIODS, index=i_b)
        if a_label == b_label:
            st.warning('Выберите два разных периода.')
        else:
            compare(a_label, b_label)

# ── разбор: выплату можно определить руками ──
with tab1:
    rev = f[f['Группа'].eq(REVIEW)].sort_values('PAYMENT', ascending=False)
    open_q = f[f['Открытый вопрос']]

    st.markdown(
        f'<div class="note" style="margin:.4rem 0 .8rem">'
        f'{len(rev)} выплат на {money(rev["PAYMENT"].sum())} сом не объяснены ни одним '
        f'источником. Поставьте решение в колонке «Решение» — выплата уйдёт в «{REESTR}». '
        f'Галочка «Правило по договору» распространяет решение на все выплаты этого договора, '
        f'в том числе в следующих месяцах.</div>', unsafe_allow_html=True)

    if rev.empty:
        st.success('Все выплаты объяснены.')
    else:
        ed = rev[['key', 'ключ договора', 'N', 'CONTRACTOR', 'Номер договора', 'Область',
                  'BALANCE_START', 'CHARGE', 'PAYMENT', 'BALANCE_END',
                  'Решение', 'Комментарий']].copy()
        ed['Правило по договору'] = ed['ключ договора'].isin(dec['contracts'])
        ed = ed.rename(columns=NAMES)

        edited = st.data_editor(
            ed, hide_index=True, height=440, key='editor',
            disabled=['key', 'ключ договора', 'Сайт', 'Контрагент', 'Договор', 'Область',
                      'Сальдо на начало', 'Начислено', 'Выплата', 'Сальдо на конец'],
            column_config={
                'key': None, 'ключ договора': None,
                'Решение': st.column_config.SelectboxColumn(
                    'Решение', options=DECISIONS, width='medium',
                    help='Чем объясняется выплата'),
                'Комментарий': st.column_config.TextColumn(
                    'Комментарий', width='medium',
                    help='Обязателен, если объяснить не удалось'),
                'Правило по договору': st.column_config.CheckboxColumn(
                    'Правило по договору', width='small',
                    help='Применять решение ко всем выплатам договора, включая будущие месяцы'),
            })

        b1, b2, _ = st.columns([1, 1, 4])
        if b1.button('Сохранить решения', type='primary'):
            for _, r in edited.iterrows():
                k, ck = r['key'], r['ключ договора']
                res, note = (r['Решение'] or ''), (r['Комментарий'] or '')
                if res == '':
                    dec['payments'].pop(k, None)
                    if r['Правило по договору']:
                        dec['contracts'].pop(ck, None)
                elif r['Правило по договору']:
                    dec['contracts'][ck] = {'решение': res, 'комментарий': note}
                    dec['payments'].pop(k, None)
                else:
                    dec['payments'][k] = {'решение': res, 'комментарий': note}
            write_decisions(dec)
            st.rerun()

        if b2.button('Сбросить все решения'):
            st.session_state.dec = blank_decisions()
            write_decisions(st.session_state.dec)
            st.rerun()

    if not open_q.empty:
        st.markdown(f'<div class="section-title">Объяснить не удалось · '
                    f'{money(open_q["PAYMENT"].sum())} сом</div>', unsafe_allow_html=True)
        st.markdown('<div class="note" style="margin-bottom:.6rem">Эти выплаты разобрали, '
                    'но закрыть не смогли. Причина — в комментарии.</div>',
                    unsafe_allow_html=True)
        table(open_q, 240, key='open')

# ── реестр: за что заплатили ──
with tab2:
    ok = f[f['Группа'].eq(REESTR)]
    st.markdown('<div class="section-title">Из чего состоят выплаты по реестру</div>',
                unsafe_allow_html=True)
    p = st.columns(4)
    for col, part, color in zip(p, PARTS, [COOL, OK, VIOLET]):
        s = ok[part].sum()
        mini(col, part, money(s), f'{int((ok[part] > 0).sum())} выплат', color)
    mini(p[3], 'Итого по реестру', money(ok['PAYMENT'].sum()), f'{len(ok)} выплат', MUTED)
    st.markdown('<div class="note" style="margin-top:.5rem">Одна выплата может закрывать '
                'сразу долг, текущий месяц и аванс, поэтому суммы разнесены по частям, '
                'а не по выплатам. Пример: выплата 144 000 при долге 12 000 и начислении '
                '12 000 — это 12 000 долга, 12 000 за месяц и 120 000 аванса.</div>',
                unsafe_allow_html=True)

    st.markdown('<div class="section-title">По основанию</div>', unsafe_allow_html=True)
    agg = ok.groupby('Основание')['PAYMENT'].agg(['sum', 'size']).sort_values(
        'sum', ascending=False)
    for row_start in range(0, len(agg), 5):
        chunk = agg.iloc[row_start:row_start + 5]
        cs = st.columns(5)
        for col, (name, row) in zip(cs, chunk.iterrows()):
            mini(col, name[:38], money(row['sum']), f'{int(row["size"])} выплат',
                 BASIS_COLOR.get(name, OK))
        st.markdown('<div style="height:.5rem"></div>', unsafe_allow_html=True)

    table(ok, 420, key='ok')

with tab3:
    nc = f[f['Группа'].isin(NOCHECK)]
    st.markdown(f'<div class="note" style="margin:.4rem 0 .7rem">'
                f'{len(nc)} выплат на {money(nc["PAYMENT"].sum())} сом. Контроль по ним ведётся '
                f'отдельно, здесь они только для полноты картины.</div>',
                unsafe_allow_html=True)
    table(nc, 420, key='nc')

with tab4:
    table(f, 460, key='all')