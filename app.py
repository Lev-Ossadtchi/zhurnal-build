# -*- coding: utf-8 -*-
"""Общий журнал (диспетчер, этап 2). Единая дневная запись по машине (работа+топливо)
-> Рапорт (топливо/пробег по машинам, листы АТ/ТТ/наши) + задел под расход/реализацию.
Справочник Настройки_журнал.xlsx (правится заказчиком): машина -> категория, норма топлива,
список поставщиков. Данные нигде не теряются молча: строки без даты/машины выносятся отдельно."""
import os, sys, glob, datetime, openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

POST = ['Газпром КП','Лукойл','Газпром ЦХХ','Кушва бензовоз','Кушва база','ИП Белай Т.В.','УралКонтрактНефть','Наличка','Колонка']
HEAD = PatternFill('solid', fgColor='DDE6F0'); BOLD = Font(bold=True)
THIN = Border(*[Side('thin', color='B0B0B0')]*4)

def num(v):
    try: return float(str(v).replace(',','.').replace(' ',''))
    except: return 0.0

def read_edinaya(folder):
    """Читает книги единой записи. Лист = машина. Шапка: Дата,№путевого,ФИО,Объект,Вид работ,
    Отработано часов,Цена,Маршрут,Спидометр нач/кон,Мотосчётчик нач/кон,Бак нач,Заправка,Поставщик,Бак кон,Глонас."""
    rows = []; skipped = []
    for f in glob.glob(os.path.join(folder, '*.xlsx')):
        if os.path.basename(f).startswith('~$') or 'Настройки' in f or 'Рапорт' in f: continue
        wb = openpyxl.load_workbook(f, data_only=True)
        for ws in wb.worksheets:
            mark, gos = '', ''
            # шапка машины: строки 1-2 (марка, гос номер)
            top = [str(c.value) for r in ws.iter_rows(min_row=1, max_row=3) for c in r if c.value]
            if top: mark = top[1] if len(top)>1 else top[0]
            # ищем строку заголовков (с «Дата»)
            hdr = None
            for r in range(1, min(ws.max_row, 8)+1):
                vals = [str(c.value).strip().lower() if c.value else '' for c in ws[r]]
                if any(v=='дата' for v in vals): hdr = r; col = {v:i for i,v in enumerate(vals)}; break
            if hdr is None: continue
            def gi(row, *names):
                for n in names:
                    i = col.get(n)
                    if i is not None and i < len(row): return row[i]
                return None
            for row in ws.iter_rows(min_row=hdr+1, values_only=True):
                date = gi(row,'дата')
                if not isinstance(date,(datetime.datetime,datetime.date)):
                    if any(gi(row,n) for n in ('заправка','спидометр на конец','отработано часов')):
                        skipped.append((ws.title, 'нет даты')); 
                    continue
                rows.append(dict(sheet=ws.title, mark=mark, book=os.path.basename(f), date=date,
                    putlist=gi(row,'№ путевого','номер путевого','путевой'), fio=gi(row,'ф.и.о. водителя','фио','ф.и.о.'),
                    obj=gi(row,'объект'), work=gi(row,'вид работ'), hours=num(gi(row,'отработано часов')),
                    sp0=num(gi(row,'спидометр на начало','спидометр нач')), sp1=num(gi(row,'спидометр на конец','спидометр кон')),
                    mh0=num(gi(row,'мотосчётчик на начало','мотосчётчик нач','м/ч нач')), mh1=num(gi(row,'мотосчётчик на конец','мотосчётчик кон','м/ч кон')),
                    bak0=num(gi(row,'бак на начало','остаток бака нач','бак нач')), bak1=num(gi(row,'бак на конец','остаток бака кон','бак кон')),
                    fuel=num(gi(row,'заправка','заправка литры','заправка, л')), post=str(gi(row,'поставщик') or '').strip(),
                    glonas=num(gi(row,'глонас','глонасс')),
                    fuel_rub=num(gi(row,'заправка руб','заправка рубли','заправка сумма','заправка дт руб')),
                    price=num(gi(row,'цена 1 часа работы','цена','цена 1 часа')),
                    remont=num(gi(row,'затраты на ремонт','ремонт')),
                    fuel_src=str(gi(row,'источник заправки','источник') or '').strip()))
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

def build_report(rows, norms, outpath):
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    by_cat = {'АТ':[], 'ТТ':[], 'наши':[]}
    for r in rows:
        cat = norms.get(r['mark'],{}).get('cat','ТТ')
        by_cat.setdefault(cat, []).append(r)
    cols = ['№ путевого','Дата','Ф.И.О.','Марка','Гос.№','спидометр нач','спидометр кон','мотосчётчик нач','мотосчётчик кон','бак нач'] + POST + ['Заправка общая','бак кон','м/ч','Пробег','расход норма','расход факт','экономия','Вид работ','Часы']
    for cat in ['АТ','ТТ','наши']:
        ws = wb.create_sheet(cat)
        ws.append(['Рапорт (собран из единой записи)'])
        ws.append(cols)
        for c in ws[2]: c.font, c.fill, c.border = BOLD, HEAD, THIN
        for r in sorted(by_cat.get(cat,[]), key=lambda x:(str(x['sheet']), x['date'])):
            nm = norms.get(r['mark'],{})
            probeg = max(0, r['sp1']-r['sp0']); mch = max(0, r['mh1']-r['mh0'])
            fakt = round(r['bak0'] + r['fuel'] - r['bak1'], 1)
            if nm.get('unit','').startswith('л/100'): norma = round(probeg*nm.get('norm',0)/100,1)
            else: norma = round(mch*nm.get('norm',0),1)
            econ = round(norma - fakt, 1)
            zap = {p:'' for p in POST}
            if r['post'] in zap and r['fuel']: zap[r['post']] = r['fuel']
            line = [r['putlist'], r['date'], r['fio'], r['mark'], r['sheet'], r['sp0'] or '', r['sp1'] or '',
                    r['mh0'] or '', r['mh1'] or '', r['bak0'] or ''] + [zap[p] for p in POST] + \
                   [r['fuel'] or '', r['bak1'] or '', mch or '', probeg or '', norma, fakt, econ, r['work'], r['hours'] or '']
            ws.append(line)
            ws.cell(ws.max_row, 2).number_format = 'DD.MM.YYYY'
        ws.freeze_panes = 'A3'
    wb.save(outpath)
    return {c: len(by_cat.get(c,[])) for c in ['АТ','ТТ','наши']}

def load_rem_objects(folder):
    """Настройки_журнал.xlsx лист 'Объекты РЭМ': объекты, которые считаются внутренними (реализация на РЭМ).
    Остальные объекты идут в 'реализация на сторону'."""
    path = os.path.join(folder, 'Настройки_журнал.xlsx')
    rem = set()
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path, data_only=True)
        if 'Объекты РЭМ' in wb.sheetnames:
            for r in wb['Объекты РЭМ'].iter_rows(min_row=2, values_only=True):
                if r and r[0]: rem.add(str(r[0]).strip().lower())
    return rem

def build_rashod(rows, rem_objs, outpath):
    """Расход+реализация по машине: заправка ДТ (руб), ремонт, реализация (стоимость работ)
    с разбивкой на сторону и на объекты РЭМ."""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Расход-реализация'
    ws.append(['Расход и реализация по машинам (собрано из единой записи)'])
    ws.append(['Машина','Марка','Заправка ДТ, руб','Затраты на ремонт, руб','Реализация на сторону, руб','Реализация на объекты РЭМ, руб','Итого реализация, руб'])
    for c in ws[2]: c.font, c.fill, c.border = BOLD, HEAD, THIN
    agg = {}
    for r in rows:
        k = (r['sheet'], r['mark'])
        a = agg.setdefault(k, dict(fuel=0.0, remont=0.0, storona=0.0, rem=0.0))
        a['fuel'] += r.get('fuel_rub', 0.0); a['remont'] += r.get('remont', 0.0)
        real = (r.get('hours', 0.0) or 0.0) * (r.get('price', 0.0) or 0.0)
        if str(r.get('obj') or '').strip().lower() in rem_objs: a['rem'] += real
        else: a['storona'] += real
    tot = dict(fuel=0, remont=0, storona=0, rem=0)
    for (sheet, mark), a in sorted(agg.items(), key=lambda x: str(x[0][0])):
        ws.append([sheet, mark, round(a['fuel']), round(a['remont']), round(a['storona']), round(a['rem']), round(a['storona']+a['rem'])])
        for c in ws[ws.max_row]: c.border = THIN
        for i in (3,4,5,6,7): ws.cell(ws.max_row, i).number_format = '# ##0'
        for key in tot: tot[key] += a[key]
    ws.append(['ИТОГО','', round(tot['fuel']), round(tot['remont']), round(tot['storona']), round(tot['rem']), round(tot['storona']+tot['rem'])])
    for c in ws[ws.max_row]: c.font, c.border = BOLD, THIN
    for col,w in zip('ABCDEFG',[16,26,16,18,22,24,18]): ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A3'
    wb.save(outpath)
    return {'машин': len(agg), 'реализация': round(tot['storona']+tot['rem'])}

def run(folder, log=print):
    rows, skipped = read_edinaya(folder)
    norms, npath = load_norms(folder)
    rem_objs = load_rem_objects(folder)
    out = os.path.join(folder, 'Рапорт (собран).xlsx')
    stats = build_report(rows, norms, out)
    out2 = os.path.join(folder, 'Расход-реализация (собран).xlsx')
    rstats = build_rashod(rows, rem_objs, out2)
    log(f'строк: {len(rows)}, пропущено: {len(skipped)}, рапорт: {stats}, расход: {rstats}')
    return dict(rows=rows, stats=stats, rstats=rstats, out=out, out2=out2, skipped=skipped)

def selftest(outfile):
    import tempfile
    msg = 'FAIL'
    try:
        d = tempfile.mkdtemp()
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'К 839 ОМ'
        ws['A1']='Работа'; ws['A2']='ГАЗ-5204'; ws['B2']='К 839 ОМ'
        ws.append([])
        ws.append(['Дата','№ путевого','Ф.И.О. водителя','Объект','Вид работ','Отработано часов','Цена','Маршрут',
                   'Спидометр нач','Спидометр кон','Мотосчётчик нач','Мотосчётчик кон','Бак нач','Заправка','Поставщик','Бак кон','Глонас','Заправка руб','Затраты на ремонт'])
        # день: пробег 100 км, заправка 50 л Лукойл, бак 20->30 => факт=20+50-30=40, норма при 30л/100км = 30
        # реализация = 8ч x 700 = 5600 (объект ЦХХ не в РЭМ -> на сторону); заправка 3000 руб; ремонт 1000
        ws.append([datetime.datetime(2026,8,1),'123','Иванов','ЦХХ','перевозка',8,700,'база',
                   1000,1100,0,0,20,50,'Лукойл',30,0,3000,1000])
        norms_wb = openpyxl.Workbook(); nw = norms_wb.active; nw.title='Нормы'
        nw.append(['Машина','Категория','Норма','Ед']); nw.append(['ГАЗ-5204','АТ',30,'л/100км'])
        os.makedirs(d, exist_ok=True)
        wb.save(os.path.join(d,'Книга1.xlsx')); norms_wb.save(os.path.join(d,'Настройки_журнал.xlsx'))
        res = run(d, log=lambda *a: None)
        assert len(res['rows'])==1, res['rows']
        rb = openpyxl.load_workbook(res['out'])
        at = rb['АТ']; hdr = [c.value for c in at[2]]
        row = [c.value for c in at[3]]
        d2 = dict(zip(hdr,row))
        assert d2['Пробег']==100, d2['Пробег']
        assert d2['расход факт']==40, d2['расход факт']
        assert d2['расход норма']==30, d2['расход норма']
        assert d2['экономия']==-10, d2['экономия']
        assert d2['Лукойл']==50, d2['Лукойл']
        # проверка расход+реализация
        rr = openpyxl.load_workbook(res['out2'])['Расход-реализация']
        rhdr = [c.value for c in rr[2]]; rrow = dict(zip(rhdr, [c.value for c in rr[3]]))
        assert rrow['Заправка ДТ, руб']==3000, rrow['Заправка ДТ, руб']
        assert rrow['Затраты на ремонт, руб']==1000, rrow['Затраты на ремонт, руб']
        assert rrow['Реализация на сторону, руб']==5600, rrow['Реализация на сторону, руб']
        # проверка GUI (tkinter грузится в сборке). Локально на macOS _tkinter может
        # отсутствовать — это ок; на Windows-раннере tkinter есть, и потеря его в сборке — ошибка.
        gui = 'ok'
        try:
            r = build_gui(); r.update(); r.destroy()
        except ModuleNotFoundError as ge:
            if '_tkinter' in str(ge): gui = 'skip-local'
            else: raise AssertionError('GUI модуль потерян в сборке: %r' % ge)
        except Exception as ge:
            raise AssertionError('GUI не поднялся: %r' % ge)
        msg = 'OK рапорт(пробег %d факт %d норма %d) расход(заправка %d ремонт %d реализация %d)' % (
            d2['Пробег'], d2['расход факт'], d2['расход норма'],
            rrow['Заправка ДТ, руб'], rrow['Затраты на ремонт, руб'], rrow['Реализация на сторону, руб'])
    except Exception as e:
        import traceback; msg = 'FAIL %r' % (e,)
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
    folder = args[0] if args else None
    if not folder:
        from tkinter import filedialog, messagebox
        root = build_gui()
        folder = filedialog.askdirectory(title='Папка с единой записью по машинам')
        if not folder:
            return
        lines = []
        try:
            run(folder, log=lines.append)
            messagebox.showinfo('Общий журнал', 'Готово.\n' + '\n'.join(lines))
        except Exception as e:
            messagebox.showerror('Общий журнал', 'Ошибка: %s' % e)
        return
    run(folder)

if __name__ == '__main__':
    main()
