# -*- coding: utf-8 -*-
"""Общий журнал (диспетчер, этап 2). Единая дневная запись по машине (работа+топливо)
-> Рапорт (топливо/пробег по машинам, листы АТ/ТТ/наши) + Расход-реализация.
Справочник Настройки_журнал.xlsx (правится заказчиком): машина -> категория, норма топлива.
v4 (фидбек диспетчера 22.10): затраты свободными строками (сумма+наименование, как внесли),
реализация по умолчанию на объекты РЭМ, итого по каждой машине в рапорте, колонки поставщиков
строятся из реальных значений. Листы единой записи бывают разной структуры — читаем по названиям."""
import os, sys, glob, datetime, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

HEAD = PatternFill('solid', fgColor='DDE6F0'); BOLD = Font(bold=True)
THIN = Border(*[Side('thin', color='B0B0B0')]*4)
# подзаголовки второй строки шапки (под объединённой «Спидометр»/«Моточасы»)
SUB = ('нач.','начало','нач','конеч.','конеч','конец','кон.','кон')

def num(v):
    try: return float(str(v).replace(',','.').replace(' ',''))
    except: return 0.0

def read_edinaya(folder):
    """Читает книги единой записи. Лист = машина. Колонки ищем по названию (листы разной структуры).
    Поддержка двухстрочной шапки: «Спидометр»/«Моточасы» в строке выше, «нач.»/«конеч.» под ними.
    Затраты: пара «сумма (колонка Затраты на ремонт) + наименование (колонка справа)», строк может
    быть несколько, названия произвольные."""
    rows = []; skipped = []
    for f in glob.glob(os.path.join(folder, '*.xlsx')):
        b = os.path.basename(f)
        if b.startswith('~$') or 'Настройки' in b or 'Рапорт' in b or 'Расход' in b: continue
        wb = openpyxl.load_workbook(f, data_only=True)
        for ws in wb.worksheets:
            # марка машины: первая непустая ячейка строк 1-3, не начинающаяся со слова «работа»
            top = [str(c.value).strip() for r in ws.iter_rows(min_row=1, max_row=3) for c in r if c.value]
            mark = ''
            for t in top:
                if t and not t.lower().startswith('работа'): mark = t; break
            # строка заголовков (с «Дата»)
            hdr = None
            for r in range(1, min(ws.max_row, 8)+1):
                if any(str(c.value).strip().lower()=='дата' for c in ws[r] if c.value): hdr = r; break
            if hdr is None: continue
            vals = [str(c.value).strip().lower() if c.value else '' for c in ws[hdr]]
            above = [str(c.value).strip().lower() if c.value else '' for c in ws[hdr-1]] if hdr > 1 else []
            col = {}
            for i, v in enumerate(vals):
                if not v: continue
                col.setdefault(v, i)
                if v in SUB and above:
                    g = ''
                    for j in range(min(i, len(above)-1), -1, -1):
                        if above[j]: g = above[j]; break
                    if g:
                        sub = 'нач' if v.startswith('нач') else 'кон'
                        col.setdefault(g + ' ' + sub, i)
            def gi(row, *names):
                for n in names:
                    i = col.get(n)
                    if i is not None and i < len(row): return row[i]
                return None
            zi = col.get('затраты на ремонт')
            if zi is None: zi = col.get('ремонт')
            for row in ws.iter_rows(min_row=hdr+1, values_only=True):
                date = gi(row,'дата')
                if not isinstance(date,(datetime.datetime,datetime.date)):
                    if any(gi(row,n) for n in ('заправка','отработано часов')) or (zi is not None and zi < len(row) and num(row[zi])):
                        skipped.append((ws.title, 'нет даты'))
                    continue
                items = []
                if zi is not None and zi < len(row):
                    s = num(row[zi])
                    if s:
                        nm = row[zi+1] if zi+1 < len(row) and row[zi+1] not in (None, '') else None
                        items.append((str(nm).strip() if nm else 'Затраты на ремонт', s))
                rows.append(dict(sheet=ws.title, mark=mark, book=b, date=date,
                    putlist=gi(row,'№ путевого','номер путевого','путевой'), fio=gi(row,'ф.и.о. водителя','фио','ф.и.о.'),
                    obj=gi(row,'объект'), work=gi(row,'вид работ'), hours=num(gi(row,'отработано часов')),
                    sp0=num(gi(row,'спидометр на начало','спидометр нач')), sp1=num(gi(row,'спидометр на конец','спидометр кон')),
                    mh0=num(gi(row,'мотосчётчик на начало','мотосчётчик нач','м/ч нач','моточасы нач')),
                    mh1=num(gi(row,'мотосчётчик на конец','мотосчётчик кон','м/ч кон','моточасы кон')),
                    bak0=num(gi(row,'бак на начало','остаток бака нач','бак нач')), bak1=num(gi(row,'бак на конец','остаток бака кон','бак кон')),
                    fuel=num(gi(row,'заправка','заправка литры','заправка, л')), post=str(gi(row,'поставщик') or '').strip(),
                    glonas=num(gi(row,'глонас','глонасс')),
                    fuel_rub=num(gi(row,'заправка руб','заправка рубли','заправка сумма','заправка дт руб')),
                    price=num(gi(row,'цена 1 часа работы','цена','цена 1 часа')),
                    stoim=num(gi(row,'стоимость работы','стоимость')),
                    fuel_src=str(gi(row,'источник заправки','источник') or '').strip(),
                    cost_items=items))
    return rows, skipped

def load_norms(folder):
    """Настройки_журнал.xlsx: лист 'Нормы' (Машина|Категория|Норма|Ед). Если нет — создаёт шаблон."""
    path = os.path.join(folder, 'Настройки_журнал.xlsx')
    norms = {}
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path, data_only=True)
        if 'Нормы' in wb.sheetnames:
            for r in wb['Нормы'].iter_rows(min_row=2, values_only=True):
                if r[0]: norms[str(r[0]).strip()] = dict(cat=(r[1] or 'ТТ'), norm=num(r[2]), unit=(r[3] or 'л/100км'))
    return norms, path

def load_storona_objects(folder):
    """Настройки_журнал.xlsx лист 'Сторонние объекты': объекты, реализация по которым идёт НА СТОРОНУ.
    По умолчанию (листа нет или он пуст) вся реализация считается на объекты РЭМ."""
    path = os.path.join(folder, 'Настройки_журнал.xlsx')
    st = set()
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path, data_only=True)
        for name in ('Сторонние объекты','Объекты на сторону','Сторона'):
            if name in wb.sheetnames:
                for r in wb[name].iter_rows(min_row=2, values_only=True):
                    if r and r[0]: st.add(str(r[0]).strip().lower())
    return st

def suppliers_from(rows):
    """Список поставщиков из реальных данных (Поставщик или Источник заправки), по алфавиту."""
    s = set()
    for r in rows:
        for v in (r.get('post'), r.get('fuel_src')):
            if v and str(v).strip(): s.add(str(v).strip())
    return sorted(s)

def build_report(rows, norms, outpath, suppliers=None):
    """Рапорт по топливу/пробегу, листы АТ/ТТ/наши. Колонки поставщиков — из реальных данных.
    После строк каждой машины — строка ИТОГО по машине."""
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    if suppliers is None: suppliers = suppliers_from(rows)
    by_cat = {'АТ':[], 'ТТ':[], 'наши':[]}
    for r in rows:
        cat = norms.get(r['mark'],{}).get('cat','ТТ')
        by_cat.setdefault(cat, []).append(r)
    cols = (['№ путевого','Дата','Ф.И.О.','Марка','Гос.№','спидометр нач','спидометр кон','мотосчётчик нач','мотосчётчик кон','бак нач']
            + suppliers + ['Заправка общая','бак кон','м/ч','Пробег','расход норма','расход факт','экономия','Стоимость работы','Вид работ','Часы'])
    def emit_machine(ws, group):
        sums = dict(hours=0.0, stoim=0.0, fuel=0.0, probeg=0.0, fakt=0.0, bysup={s:0.0 for s in suppliers})
        g = sorted(group, key=lambda x: x['date'])
        for r in g:
            nm = norms.get(r['mark'],{})
            probeg = max(0, r['sp1']-r['sp0']); mch = max(0, r['mh1']-r['mh0'])
            fakt = round(r['bak0'] + r['fuel'] - r['bak1'], 1)
            if nm.get('unit','').startswith('л/100'): norma = round(probeg*nm.get('norm',0)/100,1)
            else: norma = round(mch*nm.get('norm',0),1)
            econ = round(norma - fakt, 1)
            zap = {s:'' for s in suppliers}
            sup = r['post'] or r.get('fuel_src','')
            if sup in zap and r['fuel']: zap[sup] = r['fuel']
            stoim = r.get('stoim') or ((r.get('hours',0) or 0)*(r.get('price',0) or 0))
            line = [r['putlist'], r['date'], r['fio'], r['mark'], r['sheet'], r['sp0'] or '', r['sp1'] or '',
                    r['mh0'] or '', r['mh1'] or '', r['bak0'] or ''] + [zap[s] for s in suppliers] + \
                   [r['fuel'] or '', r['bak1'] or '', mch or '', probeg or '', norma, fakt, econ,
                    round(stoim) if stoim else '', r['work'], r['hours'] or '']
            ws.append(line); ws.cell(ws.max_row, 2).number_format = 'DD.MM.YYYY'
            sums['hours'] += r['hours'] or 0; sums['stoim'] += stoim or 0; sums['fuel'] += r['fuel'] or 0
            sums['probeg'] += probeg; sums['fakt'] += fakt
            for s in suppliers:
                if isinstance(zap[s], (int, float)): sums['bysup'][s] += zap[s]
        m = g[0]
        itg = ['', '', 'ИТОГО '+str(m['sheet']), m['mark'], '', '', '', '', '', ''] \
              + [round(sums['bysup'][s]) or '' for s in suppliers] \
              + [round(sums['fuel']) or '', '', '', round(sums['probeg']) or '', '', round(sums['fakt'],1) or '', '',
                 round(sums['stoim']) or '', '', round(sums['hours']) or '']
        ws.append(itg)
        for c in ws[ws.max_row]: c.font = BOLD
    for cat in ['АТ','ТТ','наши']:
        ws = wb.create_sheet(cat)
        ws.append(['Рапорт (собран из единой записи)'])
        ws.append(cols)
        for c in ws[2]: c.font, c.fill, c.border = BOLD, HEAD, THIN
        groups = {}
        for r in by_cat.get(cat, []): groups.setdefault(r['sheet'], []).append(r)
        for sheet in sorted(groups, key=str): emit_machine(ws, groups[sheet])
        ws.freeze_panes = 'A3'
    wb.save(outpath)
    return {c: len(by_cat.get(c,[])) for c in ['АТ','ТТ','наши']}

def build_rashod(rows, storona_objs, outpath):
    """Расход+реализация по машине: заправка ДТ (руб), затраты (свободные статьи как внесены),
    реализация (по умолчанию на объекты РЭМ; на сторону — если объект в списке сторонних)."""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Расход-реализация'
    ws.append(['Расход и реализация по машинам (собрано из единой записи)'])
    header = ['Машина','Марка','Заправка ДТ, руб','Итого затрат, руб',
              'Реализация на объекты РЭМ, руб','Реализация на сторону, руб','Итого реализация, руб']
    ws.append(header)
    for c in ws[2]: c.font, c.fill, c.border = BOLD, HEAD, THIN
    agg = {}; order = []
    for r in rows:
        k = (r['sheet'], r['mark'])
        if k not in agg: agg[k] = dict(fuel=0.0, items={}, rem=0.0, storona=0.0); order.append(k)
        a = agg[k]
        a['fuel'] += r.get('fuel_rub', 0.0) or 0.0
        for nm, s in (r.get('cost_items') or []):
            a['items'][nm] = a['items'].get(nm, 0.0) + s
        real = r.get('stoim') or ((r.get('hours',0.0) or 0.0)*(r.get('price',0.0) or 0.0))
        if str(r.get('obj') or '').strip().lower() in storona_objs: a['storona'] += real
        else: a['rem'] += real
    order = sorted(order, key=lambda x: str(x[0]))
    tot = dict(fuel=0.0, zat=0.0, rem=0.0, storona=0.0)
    for k in order:
        a = agg[k]; zat = sum(a['items'].values())
        ws.append([k[0], k[1], round(a['fuel']), round(zat), round(a['rem']), round(a['storona']), round(a['rem']+a['storona'])])
        for c in ws[ws.max_row]: c.border = THIN
        for i in range(3, 8): ws.cell(ws.max_row, i).number_format = '# ##0'
        tot['fuel'] += a['fuel']; tot['zat'] += zat; tot['rem'] += a['rem']; tot['storona'] += a['storona']
    ws.append(['ИТОГО','', round(tot['fuel']), round(tot['zat']), round(tot['rem']), round(tot['storona']), round(tot['rem']+tot['storona'])])
    for c in ws[ws.max_row]: c.font, c.border = BOLD, THIN
    for i in range(3, 8): ws.cell(ws.max_row, i).number_format = '# ##0'
    for i, w in enumerate([16,26,16,16,24,22,18], 1): ws.column_dimensions[get_column_letter(i)].width = w
    # детализация затрат «как внесены»
    ws.append([]); ws.append(['Затраты по статьям (как внесены в единую запись)'])
    ws.cell(ws.max_row, 1).font = BOLD
    ws.append(['Машина','Марка','Статья','Сумма, руб'])
    for c in ws[ws.max_row]: c.font, c.fill, c.border = BOLD, HEAD, THIN
    any_z = False
    for k in order:
        a = agg[k]
        if not a['items']: continue
        any_z = True
        for nm, s in a['items'].items():
            ws.append([k[0], k[1], nm, round(s)]); ws.cell(ws.max_row, 4).number_format = '# ##0'
            for c in ws[ws.max_row]: c.border = THIN
        ws.append(['', '', 'Итого по машине', round(sum(a['items'].values()))])
        for c in ws[ws.max_row]: c.font = BOLD
        ws.cell(ws.max_row, 4).number_format = '# ##0'
    if not any_z: ws.append(['', '', 'затрат не внесено', ''])
    wb.save(outpath)
    return {'машин': len(agg), 'затраты': round(tot['zat']), 'реализация': round(tot['rem']+tot['storona'])}

def run(folder, log=print):
    rows, skipped = read_edinaya(folder)
    norms, npath = load_norms(folder)
    storona = load_storona_objects(folder)
    out = os.path.join(folder, 'Рапорт (собран).xlsx')
    stats = build_report(rows, norms, out)
    out2 = os.path.join(folder, 'Расход-реализация (собран).xlsx')
    rstats = build_rashod(rows, storona, out2)
    log(f'строк: {len(rows)}, пропущено: {len(skipped)}, рапорт: {stats}, расход: {rstats}')
    return dict(rows=rows, stats=stats, rstats=rstats, out=out, out2=out2, skipped=skipped)

def selftest(outfile):
    import tempfile
    msg = 'FAIL'
    try:
        d = tempfile.mkdtemp()
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = '7895'
        ws['A1']='Работа на ООО "РЭМ"'; ws['A2']='Кран ZRT600'
        # двухстрочная шапка: Спидометр/Моточасы в строке 3, нач./конеч. в строке 4
        ws['J3']='Спидометр'; ws['L3']='Моточасы'
        ws.append([])  # на всякий случай не используется, шапку ставим явно ниже
        hdr = ['Дата','Ф.И.О. водителя','м/ч','Объект','Вид работ','Отработано часов','Цена 1 часа работы',
               'Стоимость работы','Маршрут','нач.','конеч.','нач.','конеч.','Бак нач','Заправка','Поставщик',
               'Бак кон','Глонас','Заправка руб','Источник заправки','Затраты на ремонт','']
        # шапку кладём в строку 4
        for i, v in enumerate(hdr, 1): ws.cell(4, i, v)
        # строка данных: объект УГМК-ЦХХ (на объекты РЭМ по умолчанию), стоимость 62700,
        # заправка 180 л Кушва, заправка 3000 руб, затраты 7907.36 ГСМ + 24426.25 Запчасти
        ws.cell(5,1, datetime.datetime(2026,9,1)); ws.cell(5,2,'Лагунов'); ws.cell(5,4,'УГМК-ЦХХ')
        ws.cell(5,5,'погрузка'); ws.cell(5,6,11); ws.cell(5,7,5700); ws.cell(5,8,62700)
        ws.cell(5,10,6994); ws.cell(5,11,6998); ws.cell(5,12,7095); ws.cell(5,13,7105); ws.cell(5,14,100)
        ws.cell(5,15,180); ws.cell(5,16,'Кушва'); ws.cell(5,17,200); ws.cell(5,19,3000); ws.cell(5,20,'Кушва')
        ws.cell(5,21,7907.36); ws.cell(5,22,'ГСМ, масла, жидкости')
        ws.cell(6,1, datetime.datetime(2026,9,2)); ws.cell(6,4,'УГМК-ЦХХ'); ws.cell(6,7,5700); ws.cell(6,8,0)
        ws.cell(6,21,24426.25); ws.cell(6,22,'Запчасти, материалы')
        norms_wb = openpyxl.Workbook(); nw = norms_wb.active; nw.title='Нормы'
        nw.append(['Машина','Категория','Норма','Ед']); nw.append(['Кран ZRT600','наши',0,'л/мч'])
        os.makedirs(d, exist_ok=True)
        wb.save(os.path.join(d,'Книга1.xlsx')); norms_wb.save(os.path.join(d,'Настройки_журнал.xlsx'))
        res = run(d, log=lambda *a: None)
        assert len(res['rows'])==2, len(res['rows'])
        # рапорт: колонка поставщика Кушва, заправка 180, итого по машине
        rb = openpyxl.load_workbook(res['out'])['наши']
        rhdr = [c.value for c in rb[2]]
        assert 'Кушва' in rhdr, rhdr
        r1 = dict(zip(rhdr, [c.value for c in rb[3]]))
        assert r1['Кушва']==180, r1['Кушва']
        assert r1['Стоимость работы']==62700, r1['Стоимость работы']
        # расход-реализация: реализация на объекты РЭМ (не сторона), затраты свободные
        rr = openpyxl.load_workbook(res['out2'])['Расход-реализация']
        rrh = [c.value for c in rr[2]]; rrow = dict(zip(rrh, [c.value for c in rr[3]]))
        assert rrow['Реализация на объекты РЭМ, руб']==62700, rrow['Реализация на объекты РЭМ, руб']
        assert rrow['Реализация на сторону, руб']==0, rrow['Реализация на сторону, руб']
        assert rrow['Итого затрат, руб']==round(7907.36+24426.25), rrow['Итого затрат, руб']
        assert rrow['Заправка ДТ, руб']==3000, rrow['Заправка ДТ, руб']
        # детализация затрат содержит наименования как внесены
        txt = []
        for row in rr.iter_rows(values_only=True):
            for c in row:
                if c: txt.append(str(c))
        joined = ' | '.join(txt)
        assert 'ГСМ, масла, жидкости' in joined, 'нет статьи ГСМ в детализации'
        assert 'Запчасти, материалы' in joined, 'нет статьи Запчасти'
        gui = 'ok'
        try:
            rt = build_gui(); rt.update(); rt.destroy()
        except ModuleNotFoundError as ge:
            if '_tkinter' in str(ge): gui = 'skip-local'
            else: raise AssertionError('GUI модуль потерян в сборке: %r' % ge)
        except Exception as ge:
            raise AssertionError('GUI не поднялся: %r' % ge)
        msg = 'OK рапорт(Кушва=%s стоимость=%s) расход(РЭМ=%s затраты=%s) статьи-детализация=да' % (
            r1['Кушва'], r1['Стоимость работы'],
            rrow['Реализация на объекты РЭМ, руб'], rrow['Итого затрат, руб'])
    except Exception as e:
        msg = 'FAIL %r' % (e,)
    with open(outfile,'w',encoding='utf-8') as f: f.write(msg)
    print(msg)

def build_gui():
    """Поднимает окно выбора папки; возвращает root. Для самопроверки exe (PyInstaller
    теряет модули, подгружаемые по имени — tkinter проверяем явно)."""
    import tkinter as tk
    root = tk.Tk(); root.withdraw()
    return root

def main():
    if '--selftest' in sys.argv:
        i = sys.argv.index('--selftest')
        selftest(sys.argv[i+1] if len(sys.argv) > i+1 else 'selftest.txt')
        return
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    if args:
        folder = args[0]
    elif getattr(sys, 'frozen', False):
        folder = os.path.dirname(sys.executable)
    else:
        folder = os.getcwd()
    gui = getattr(sys, 'frozen', False) and not args
    def show(title, text, err=False):
        if gui:
            from tkinter import messagebox
            build_gui()
            (messagebox.showerror if err else messagebox.showinfo)(title, text)
        else:
            print(text)
    try:
        res = run(folder, log=lambda *a: None)
        n = len(res['rows']); rst = res.get('rstats', {})
        if n == 0:
            show('Общий журнал',
                 'В этой папке не нашёл данных.\n\nПоложите книги единой записи в эту же папку, '
                 'рядом с программой, и запустите снова.')
            return
        show('Общий журнал',
             'Готово.\n\nОбработано строк: %d\nМашин в расходе: %s, реализация: %s руб\n\n'
             'Создано в этой папке, рядом с программой:\n  Рапорт (собран).xlsx\n  '
             'Расход-реализация (собран).xlsx\n\n'
             'В Расход-реализация затраты выводятся так, как внесены в единую запись (статья и сумма), '
             'плюс Итого затрат по машине. Реализация по умолчанию идёт на объекты РЭМ.'
             % (n, rst.get('машин'), rst.get('реализация')))
    except Exception as e:
        show('Общий журнал', 'Ошибка: %s' % e, err=True)

if __name__ == '__main__':
    main()
