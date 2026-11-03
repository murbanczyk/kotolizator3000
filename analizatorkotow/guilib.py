import docx
import glob
import numpy as np
import os
import pandas as pd
import pathlib
import platform
import shutil
import subprocess
import sys
import tempfile
import wx
import wx.adv
import wx.lib.scrolledpanel as scrolled
from datetime import date, datetime
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Mm, Cm, Pt
from openpyxl.styles import Font
from threading import Thread
from . import dane
from .scraper import PaluchScraper, LIST_URLS
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_agg import FigureCanvasAgg


def resource_path(name):
    """Ścieżka do pliku z katalogu programu (działa też w aplikacji zbudowanej PyInstallerem)."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), name)


def mainGui():
    app = wx.App()
    try:
        bitmap = wx.Bitmap(resource_path('splash.png'))
        wx.adv.SplashScreen(bitmap, wx.adv.SPLASH_CENTRE_ON_SCREEN | wx.adv.SPLASH_TIMEOUT, 4500, None, -1).Show()
    except Exception as e:
        print('Brak ekranu powitalnego: ' + str(e))
    if os.name == 'nt':
        app.locale = wx.Locale(wx.LANGUAGE_POLISH)
    CatFrame().Show()
    app.MainLoop()


def Tooltipset(widget, tooltip):
    widget.SetToolTip(tooltip)


def error_dialog(parent, error):
    wx.MessageDialog(parent, str(error), 'Błąd', wx.OK | wx.ICON_ERROR).ShowModal()


def window_size(parent, width, height):
    """Rozmiar okna ograniczony do 90% ekranu."""
    area = wx.Display(wx.Display.GetFromWindow(parent) if parent else 0).GetClientArea()
    return min(width, int(area.width * 0.9)), min(height, int(area.height * 0.9))


class GeneralLayout(scrolled.ScrolledPanel):
    def __init__(self, parent):
        self.my_datadir = self.get_datadir() / "kotolizator"  # katalog na zapisane zestawienia
        self.today = date.today().strftime("%d-%m-%Y")
        self.my_datadir.mkdir(parents=True, exist_ok=True)
        print(self.my_datadir)
        self.my_datadir2 = str(self.my_datadir) + os.sep
        self.kotybool = True
        self.busy = False
        self.counts = None
        self.frames = {}  # pamięć wczytanych zestawień: (prefiks, dzień) -> DataFrame
        self.dates = []

        scrolled.ScrolledPanel.__init__(self, parent)
        self.SetupScrolling(scroll_y=True, scroll_x=False)
        main = wx.BoxSizer(wx.VERTICAL)

        # 1. gatunek
        self.gatunek = wx.RadioBox(self, label='1. Jakie zwierzęta?', choices=['Koty', 'Psy'],
                                   majorDimension=2, style=wx.RA_SPECIFY_COLS)
        Tooltipset(self.gatunek, 'Wybierz, czy analizujesz koty, czy psy')
        self.gatunek.Bind(wx.EVT_RADIOBOX, self.switchkotpies)
        main.Add(self.gatunek, 0, wx.ALL | wx.EXPAND, 8)

        # 2. dane
        box = wx.StaticBoxSizer(wx.VERTICAL, self, '2. Dane')
        sb = box.GetStaticBox()
        row = wx.BoxSizer(wx.HORIZONTAL)
        self.nowedane = wx.Button(sb, label='Pobierz aktualne dane ze strony')
        Tooltipset(self.nowedane, 'Pobiera aktualną listę zwierząt ze strony schroniska i zapisuje ją pod dzisiejszą datą')
        self.nowedane.Bind(wx.EVT_BUTTON, self.parsepaluch)
        row.Add(self.nowedane, 0, wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, 12)
        row.Add(wx.StaticText(sb, label='Dane z dnia:'), 0, wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, 6)
        self.backup = wx.Choice(sb, size=(140, -1))
        Tooltipset(self.backup, 'Z którego dnia pokazać zestawienia? Każde pobranie danych zapisuje nowy dzień.')
        self.backup.Bind(wx.EVT_CHOICE, self.combochange)
        row.Add(self.backup, 0, wx.ALIGN_CENTER_VERTICAL)
        box.Add(row, 0, wx.ALL, 6)
        self.daneinfo = wx.StaticText(sb, label='')
        box.Add(self.daneinfo, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        main.Add(box, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 8)

        # 3. zestawienia dla wybranego dnia
        box = wx.StaticBoxSizer(wx.VERTICAL, self, '3. Zestawienia dla wybranego dnia')
        sb = box.GetStaticBox()
        grid = wx.GridSizer(3, 2, 6, 6)
        self.browse = wx.Button(sb, label='Przeglądaj i szukaj')
        Tooltipset(self.browse, 'Tabela wszystkich zwierząt z wyszukiwarką; kliknij nagłówek kolumny, aby posortować')
        self.longest = wx.Button(sb, label='Najdłużej czekające')
        Tooltipset(self.longest, 'Zwierzęta, które są w schronisku najdłużej – warto je promować')
        self.word1 = wx.Button(sb, label='Lista zwierząt (Word)')
        Tooltipset(self.word1, 'Dokument Word ze zdjęciem, wiekiem, datą przyjęcia i opisem każdego zwierzęcia.\n'
                               'Pilne adopcje są podświetlone na czerwono, Ochota na Kota na fioletowo.')
        self.genxls = wx.Button(sb, label='Tabelka (Excel)')
        Tooltipset(self.genxls, 'Krótka tabelka: numer, imię, wiek i miejsce pobytu')
        self.histogram1 = wx.Button(sb, label='Histogram wieku')
        Tooltipset(self.histogram1, 'Wykresy wieku zwierząt: wszystkich, samic i samców')
        self.statcat = wx.Button(sb, label='Ile zwierząt w którym miejscu')
        Tooltipset(self.statcat, 'Liczba zwierząt w poszczególnych miejscach (np. Ochota na Kota)')
        self.day_buttons = (self.browse, self.longest, self.word1, self.genxls, self.histogram1, self.statcat)
        for b in self.day_buttons:
            grid.Add(b, 0, wx.EXPAND)
        box.Add(grid, 0, wx.ALL | wx.EXPAND, 6)
        main.Add(box, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 8)

        # 4. porównanie i historia
        box = wx.StaticBoxSizer(wx.VERTICAL, self, '4. Porównanie z innym dniem i historia')
        sb = box.GetStaticBox()
        row = wx.BoxSizer(wx.HORIZONTAL)
        row.Add(wx.StaticText(sb, label='Porównaj z dniem:'), 0, wx.RIGHT | wx.ALIGN_CENTER_VERTICAL, 6)
        self.backup2 = wx.Choice(sb, size=(140, -1))
        Tooltipset(self.backup2, 'Z którym dniem porównać wybrany dzień?')
        self.backup2.Bind(wx.EVT_CHOICE, self.combochange2)
        row.Add(self.backup2, 0, wx.ALIGN_CENTER_VERTICAL)
        box.Add(row, 0, wx.ALL, 6)
        grid = wx.FlexGridSizer(2, 3, 6, 6)
        grid.AddGrowableCol(0)
        self.txtnowe = wx.StaticText(sb, label='Nowe zwierzęta:')
        self.txtznikniete = wx.StaticText(sb, label='Zniknięte zwierzęta:')
        self.wordNew = wx.Button(sb, label='Lista (Word)')
        Tooltipset(self.wordNew, 'Zwierzęta, które są w wybranym dniu, a nie było ich w dniu porównania')
        self.histnew = wx.Button(sb, label='Histogram')
        Tooltipset(self.histnew, 'Wiek zwierząt, które przybyły')
        self.wordOld = wx.Button(sb, label='Lista (Word)')
        Tooltipset(self.wordOld, 'Zwierzęta, które były w dniu porównania, a nie ma ich w wybranym dniu (np. adoptowane)')
        self.histold = wx.Button(sb, label='Histogram')
        Tooltipset(self.histold, 'Wiek zwierząt, które zniknęły z listy')
        for w in (self.txtnowe, self.wordNew, self.histnew, self.txtznikniete, self.wordOld, self.histold):
            grid.Add(w, 0, wx.ALIGN_CENTER_VERTICAL | (wx.EXPAND if isinstance(w, wx.Button) else 0))
        box.Add(grid, 0, wx.ALL | wx.EXPAND, 6)
        self.porinfo = wx.StaticText(sb, label='')
        box.Add(self.porinfo, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 6)
        self.trendbtn = wx.Button(sb, label='Liczba zwierząt w czasie (wszystkie zapisane dni)')
        Tooltipset(self.trendbtn, 'Wykres liczby zwierząt oraz nowych i znikniętych zwierząt w kolejnych zapisanych dniach')
        box.Add(self.trendbtn, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 6)
        main.Add(box, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 8)

        # postęp i komunikaty
        self.postep = wx.Gauge(self, range=100, size=(-1, 18), name='Postęp')
        Tooltipset(self.postep, 'Postęp pobierania danych lub tworzenia zestawienia')
        main.Add(self.postep, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 8)
        self.statustxt = wx.StaticText(self, label='')
        main.Add(self.statustxt, 0, wx.ALL | wx.EXPAND, 8)

        self.browse.Bind(wx.EVT_BUTTON, self.showbrowse)
        self.longest.Bind(wx.EVT_BUTTON, self.showlongest)
        self.word1.Bind(wx.EVT_BUTTON, self.genword)
        self.genxls.Bind(wx.EVT_BUTTON, self.excellgen)
        self.histogram1.Bind(wx.EVT_BUTTON, self.plothist)
        self.statcat.Bind(wx.EVT_BUTTON, self.getcatstat)
        self.wordNew.Bind(wx.EVT_BUTTON, self.genwordNew)
        self.wordOld.Bind(wx.EVT_BUTTON, self.genwordOld)
        self.histnew.Bind(wx.EVT_BUTTON, self.plothistnew)
        self.histold.Bind(wx.EVT_BUTTON, self.plothistold)
        self.trendbtn.Bind(wx.EVT_BUTTON, self.plottrend)

        self.SetSizerAndFit(main)
        self.refresh_dates()

    # --- stan interfejsu ---

    def prefix(self):
        return 'Miau' if self.kotybool else 'Hau'

    def zwierzeta(self):
        return 'kotów' if self.kotybool else 'psów'

    def day(self):
        return self.backup.GetStringSelection()

    def day2(self):
        return self.backup2.GetStringSelection()

    def status(self, msg):
        """Komunikat pod paskiem postępu; można wołać z dowolnego wątku."""
        wx.CallAfter(self.set_status, msg)

    def set_status(self, msg):
        self.statustxt.SetLabel(msg)
        self.statustxt.Wrap(max(self.GetClientSize().width - 20, 200))  # długi komunikat w kilku liniach
        self.Layout()

    def find_dates(self):
        dates = []
        for k in glob.glob(self.my_datadir2 + self.prefix() + '*.kotolizer'):
            day = os.path.basename(k)[len(self.prefix()):-len('.kotolizer')]
            try:
                dates.append((datetime.strptime(day, '%d-%m-%Y'), day))
            except ValueError:
                pass
        return [d for _, d in sorted(dates, reverse=True)]  # najnowsze na górze

    def refresh_dates(self, select=None):
        """Odświeża listy dni; domyślnie wybiera najnowszy dzień i dzień przed nim do porównania."""
        self.dates = self.find_dates()
        self.backup.SetItems(self.dates)
        if self.dates:
            self.backup.SetSelection(self.dates.index(select) if select in self.dates else 0)
        if self.dates and self.dates[0] == self.today:
            self.nowedane.SetLabel('Pobierz dane ponownie')
        else:
            self.nowedane.SetLabel('Pobierz aktualne dane ze strony')
        self.combochange(None)

    def combochange(self, event):
        day = self.day()
        others = [d for d in self.dates if d != day]
        self.backup2.SetItems(others)
        if others:
            older = [d for d in others if self.dates.index(d) > self.dates.index(day)]
            self.backup2.SetStringSelection(older[0] if older else others[0])
        if not day:
            self.daneinfo.SetLabel('Brak zapisanych danych. Kliknij „Pobierz aktualne dane ze strony”.')
        else:
            try:
                n = len(self.getframe(day))
                info = 'W dniu %s na liście było %d %s.' % (day, n, self.zwierzeta())
            except Exception as e:
                info = 'Nie udało się wczytać danych z dnia %s: %s' % (day, e)
            if day != self.today:
                info += '\nTo nie są dzisiejsze dane – aby je mieć, kliknij „Pobierz aktualne dane ze strony”.'
            self.daneinfo.SetLabel(info)
        self.combochange2(None)

    def combochange2(self, event):
        self.txtnowe.SetLabel('Nowe zwierzęta:')
        self.txtznikniete.SetLabel('Zniknięte zwierzęta:')
        self.counts = None
        if not self.day2():
            self.porinfo.SetLabel('Do porównania potrzebne są dane z co najmniej dwóch dni.')
        else:
            try:
                self.counts = (len(self.newframe()), len(self.goneframe()))
                self.txtnowe.SetLabel('Nowe zwierzęta: %d' % self.counts[0])
                self.txtznikniete.SetLabel('Zniknięte zwierzęta: %d' % self.counts[1])
                self.porinfo.SetLabel('Porównanie dnia %s z dniem %s.' % (self.day(), self.day2()))
            except Exception as e:
                self.porinfo.SetLabel('Nie udało się porównać: %s' % e)
        self.update_buttons()
        self.Layout()
        self.GetParent().Layout()

    def update_buttons(self):
        idle = not self.busy
        have = idle and bool(self.day())
        for b in self.day_buttons:
            b.Enable(have)
        compare = have and self.counts is not None
        self.wordNew.Enable(compare and self.counts[0] > 0)
        self.histnew.Enable(compare and self.counts[0] > 0)
        self.wordOld.Enable(compare and self.counts[1] > 0)
        self.histold.Enable(compare and self.counts[1] > 0)
        self.trendbtn.Enable(idle and len(self.dates) >= 2)
        for w in (self.nowedane, self.gatunek, self.backup, self.backup2):
            w.Enable(idle)

    def run_task(self, target, args=(), done=None):
        """Uruchamia długą operację w tle, blokując w tym czasie przyciski."""
        self.busy = True
        self.update_buttons()
        self.postep.SetValue(0)

        def worker():
            try:
                result = target(*args)
            except Exception as e:
                import traceback
                traceback.print_exc()
                wx.CallAfter(self.task_finished, None, e)
            else:
                wx.CallAfter(self.task_finished, done(result) if done else None, None)

        t = Thread(target=worker)
        t.daemon = True
        t.start()

    def task_finished(self, msg, error):
        self.busy = False
        self.update_buttons()
        if error is not None:
            self.set_status('Wystąpił błąd.')
            error_dialog(self, error)
        elif msg:
            self.set_status(msg)

    def ask_path(self, title, wildcard, default):
        with wx.FileDialog(self, title, wildcard=wildcard, defaultFile=default,
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return None
            return fileDialog.GetPath()

    def filename(self, kind, ext, with_compare=False):
        name = kind + '_' + ('koty' if self.kotybool else 'psy') + '_' + self.day()
        if with_compare:
            name += '_vs_' + self.day2()
        return name + '.' + ext

    def switchkotpies(self, event):
        self.kotybool = self.gatunek.GetSelection() == 0
        self.set_status('')
        self.postep.SetValue(0)
        self.refresh_dates()

    # --- dane ---

    def getframe(self, day):
        key = (self.prefix(), day)
        if key not in self.frames:
            df = pd.read_pickle(self.my_datadir2 + key[0] + key[1] + '.kotolizer', compression='bz2')
            self.frames[key] = dane.sort_frame(df[dane.COLUMNS])
        return self.frames[key]

    def newframe(self):
        return dane.comparedf(self.getframe(self.day()), self.getframe(self.day2()))

    def goneframe(self):
        return dane.comparedf(self.getframe(self.day2()), self.getframe(self.day()))

    def parsepaluch(self, event):
        self.status('Pobieranie listy %s…' % self.zwierzeta())
        self.run_task(self.parsepaluch2, (), self.parse_done)

    def parse_done(self, result):
        n, fetched, found, ochota = result
        wx.CallAfter(self.refresh_dates, self.today)
        print('Nowych lub odświeżonych stron zwierząt: %d' % fetched)
        extra = []
        if ochota:
            extra.append('%d na Ochocie na Kota' % ochota)
        if found:
            extra.append('%d ostatnio znalezionych' % found)
        return 'Pobrano %d %s%s.' % (n, self.zwierzeta(), ' (' + ', '.join(extra) + ')' if extra else '')

    def parsepaluch2(self):
        species = 'koty' if self.kotybool else 'psy'
        scraper = PaluchScraper(self.my_datadir2 + 'cache.kotolizer')
        animals = scraper.get_animals(species)
        wx.CallAfter(self.postep.SetRange, len(animals))
        pending = []
        fetched = 0
        try:
            for count, a in enumerate(animals):
                self.status('Zwierzę %d z %d: %s' % (count + 1, len(animals), a['name'] or a['number']))
                # zdjęcie pobiera się w tle, równolegle ze stroną zwierzęcia
                image = scraper.thumbnail_later(a['image'], referer=a['link'])
                details, new = scraper.get_details(a['link'], referer=LIST_URLS[species])
                if new:
                    fetched += 1
                    if fetched % 20 == 0:
                        scraper.save_cache()
                pending.append((a, details, image))
                wx.CallAfter(self.postep.SetValue, count + 1)
            rows = []
            for count, (a, details, image) in enumerate(pending):
                if not image.done():
                    self.status('Kończę pobieranie zdjęć: %d z %d' % (count + 1, len(pending)))
                rows.append(dane.build_row(count + 1, a, details, image.result()))
        finally:
            scraper.close()
            scraper.save_cache()
        print('Zwierząt: %d, pobranych stron zwierząt: %d' % (len(animals), fetched))
        catframe = pd.DataFrame(rows, columns=dane.COLUMNS)
        catframe.to_pickle(self.my_datadir2 + self.prefix() + self.today + '.kotolizer', compression='bz2')
        self.frames.pop((self.prefix(), self.today), None)
        found = sum(1 for a in animals if a.get('found'))
        ochota = sum(1 for r in rows if dane.miejsce(r['Przyjęcie']) == 'Ochota na Kota')
        return len(animals), fetched, found, ochota

    # --- przeglądanie ---

    def showbrowse(self, event):
        table = dane.summary(self.getframe(self.day()))
        BrowseFrame(self, table, 'Zwierzęta – %s, %s' % (self.zwierzeta(), self.day()),
                    self.filename('przeglad', 'xlsx')).Show()

    def showlongest(self, event):
        table = dane.summary(self.getframe(self.day()))
        frame = BrowseFrame(self, table, 'Najdłużej czekające – %s, %s' % (self.zwierzeta(), self.day()),
                            self.filename('najdluzej_czekajace', 'xlsx'), sort_col='Dni w schronisku',
                            descending=True)
        frame.Show()
        if table['Dni w schronisku'].isna().all():
            wx.MessageDialog(frame, 'W danych z tego dnia nie ma dat przyjęcia (zapisano je starszą wersją '
                                    'programu). Pobierz aktualne dane, aby zobaczyć, kto czeka najdłużej.',
                             'Brak dat przyjęcia').ShowModal()

    # --- zestawienia ---

    def getcatstat(self, event):
        xlscat = self.gencatxls()
        wx.MessageDialog(self, xlscat.groupby(['Miejsce']).size().to_string(),
                         'Ile %s w którym miejscu (%s)' % (self.zwierzeta(), self.day())).ShowModal()

    def gencatxls(self):
        rows = []
        for nr, (_, b) in enumerate(self.getframe(self.day()).iterrows(), start=1):
            imie = str(b['Imię']).split('\n')[0]
            wiek = str(b['Wiek'])
            months = dane.age_months(wiek)
            if 'mies' in wiek and months is not None and months < 4:
                imie += '(m)'
            rows.append({'Nr': nr, 'ID': b['ID'], 'Imię': imie, 'Wiek': wiek, 'Miejsce': dane.miejsce(b['Przyjęcie'])})
        return pd.DataFrame(rows, columns=['Nr', 'ID', 'Imię', 'Wiek', 'Miejsce'])

    def excellgen(self, event):
        xlscat = self.gencatxls()
        pathname = self.ask_path('Zapisz tabelkę', 'Excel (*.xlsx)|*.xlsx', self.filename('tabelka', 'xlsx'))
        if pathname:
            with pd.ExcelWriter(pathname, engine='openpyxl') as writer:
                xlscat.to_excel(writer, sheet_name='Koty', index=False)
                # pogrubienie wierszy z Ochoty na Kota (bez pandas .style, które wymaga jinja2)
                sheet = writer.sheets['Koty']
                bold = Font(bold=True)
                for row, miejsce in enumerate(xlscat['Miejsce'], start=2):
                    if 'Ochota' in str(miejsce):
                        for cell in sheet[row]:
                            cell.font = bold
            self.set_status('Zapisano: ' + os.path.basename(pathname))
            self.openfile(pathname)

    def genword(self, event):
        self.wordtask('Zapisz listę zwierząt', self.filename('lista', 'docx'), self.getframe(self.day()),
                      'Lista %s – %s' % (self.zwierzeta(), self.day()))

    def genwordNew(self, event):
        self.wordtask('Zapisz listę nowych zwierząt', self.filename('nowe', 'docx', True), self.newframe(),
                      'Nowe %s – %s w porównaniu z %s' % (self.zwierzeta(), self.day(), self.day2()))

    def genwordOld(self, event):
        self.wordtask('Zapisz listę znikniętych zwierząt', self.filename('znikniete', 'docx', True), self.goneframe(),
                      'Zniknięte %s – %s w porównaniu z %s' % (self.zwierzeta(), self.day(), self.day2()))

    def wordtask(self, dialogtitle, default, catframe, title):
        pathname = self.ask_path(dialogtitle, 'Word (*.docx)|*.docx', default)
        if pathname:
            self.status('Tworzenie dokumentu Word…')
            self.run_task(self.createwordfile, (pathname, catframe, title),
                          lambda r: 'Zapisano: ' + os.path.basename(pathname))

    def createwordfile(self, pathname, catframe, title=''):
        catframe = catframe[dane.COLUMNS]
        doc = docx.Document()
        style = doc.styles['Normal']
        style.font.name = 'Arial'
        style.font.size = Pt(8)
        section = doc.sections[0]
        section.page_height = Mm(210)
        section.page_width = Mm(297)
        section.left_margin = section.right_margin = Mm(12)

        # nagłówek: tytuł, liczba zwierząt i legenda kolorów
        if title:
            doc.add_heading(title, level=1)
        doc.add_paragraph('Liczba zwierząt: %d. Utworzono %s.' % (len(catframe), date.today().strftime('%d.%m.%Y')))
        legend = doc.add_table(1, 2)
        for cell, text, color in zip(legend.rows[0].cells, ('Pilna adopcja', 'Ochota na Kota (kolumna z imieniem)'),
                                     ('f8ddd7', 'f0d7f8')):
            cell.text = text
            cell._tc.get_or_add_tcPr().append(parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), color)))
        doc.add_paragraph()

        t = doc.add_table(catframe.shape[0] + 1, catframe.shape[1])
        t.style = 'Table Grid'
        wx.CallAfter(self.postep.SetRange, max(catframe.shape[0], 1))
        for j in range(catframe.shape[-1]):
            t.cell(0, j).text = catframe.columns[j]
        thumbdir = tempfile.mkdtemp(prefix='kotolizator_')
        try:
            for i in range(catframe.shape[0]):
                wx.CallAfter(self.postep.SetValue, i + 1)
                przyjecie = str(catframe.values[i, 4])
                for j in range(catframe.shape[-1]):
                    value = catframe.values[i, j]
                    if j == 5:
                        if hasattr(value, 'save'):
                            thumb = os.path.join(thumbdir, '%d.jpg' % i)
                            value.convert('RGB').save(thumb, 'JPEG')
                            t.cell(i + 1, j).paragraphs[0].add_run().add_picture(thumb)
                    elif j == 0:
                        t.cell(i + 1, j).text = str(i + 1)
                    else:
                        t.cell(i + 1, j).text = str(value)
                    if 'PILNE' in przyjecie:
                        t.cell(i + 1, j)._tc.get_or_add_tcPr().append(
                            parse_xml(r'<w:shd {} w:fill="f8ddd7"/>'.format(nsdecls('w'))))
                if 'Ochota na Kota' in przyjecie:
                    t.cell(i + 1, 2)._tc.get_or_add_tcPr().append(
                        parse_xml(r'<w:shd {} w:fill="f0d7f8"/>'.format(nsdecls('w'))))
        finally:
            shutil.rmtree(thumbdir, ignore_errors=True)
        t.allow_autofit = True
        for j, width in enumerate((0.7, 1.5, 1.5, 1.2, 2, 6, 10)):
            t.columns[j].width = Cm(width)
        doc.save(pathname)
        self.openfile(pathname)

    # --- wykresy ---

    def gencathist(self, catframe, ax1, ax2, sex, kotbool):
        lata, lataF, lataM = [], [], []
        for wiek, imie in zip(catframe['Wiek'], catframe['Imię']):
            months = dane.age_months(wiek)
            if months is None:
                continue
            lata.append(months)
            if 'samiec' in imie:
                lataM.append(months)
            elif 'samica' in imie:
                lataF.append(months)

        if sex == 'all':
            values, color = lata, 'k'
            names = ('Koty', 'Koty') if kotbool else ('Psy', 'Psy')
        elif sex == 'f':
            values, color = lataF, 'pink'
            names = ('Kotki', 'Kotki') if kotbool else ('Suczki', 'Suczki')
        else:
            values, color = lataM, 'b'
            names = ('Kocury', 'Kocury') if kotbool else ('Psy (samce)', 'Psy (samce)')

        lataR = np.floor(np.array(values) / 12)
        oldest = int(np.max(lataR)) if len(lataR) else 0
        a1, _, patches = ax1.hist(np.array(values), np.arange(0, 13) - 0.5, linewidth=1, edgecolor='w', color=color)
        a2, _, patches2 = ax2.hist(lataR, np.arange(0, oldest + 2) - 0.5, linewidth=1, edgecolor='w', color=color)
        ax1.set_title('%s do 1 roku, suma: %d' % (names[0], round(sum(a1))))
        ax2.set_title('%s wszystkie, suma: %d\nstarsze od 8 lat: %d' % (names[1], round(sum(a2)),
                                                                        len(np.where(lataR >= 8)[0])))
        ax1.set_xticks(range(0, 12))
        ax1.set_xlabel('miesiące')
        ax2.set_xticks(range(0, oldest + 1))
        ax2.set_xlabel('lata')
        ylabel = 'Liczba kotów' if kotbool else 'Liczba psów'
        for ax, bars in ((ax1, patches), (ax2, patches2)):
            ax.set_ylabel(ylabel)
            ymax = [0]
            for pp in bars:
                y = pp.get_height() + 0.3
                ymax.append(y)
                ax.text(pp.get_x() + pp.get_width() / 2, y, int(pp.get_height()), ha='center')
            ax.set_ylim([0, max(ymax) * 1.15 + 0.5])
        return lata

    def plothist(self, event):
        self.showhist(self.getframe(self.day()), 'Wiek %s – %s' % (self.zwierzeta(), self.day()),
                      self.filename('histogram', 'png'))

    def plothistold(self, event):
        self.showhist(self.goneframe(),
                      'Wiek znikniętych %s – %s w porównaniu z %s' % (self.zwierzeta(), self.day(), self.day2()),
                      self.filename('histogram_znikniete', 'png', True))

    def plothistnew(self, event):
        self.showhist(self.newframe(),
                      'Wiek nowych %s – %s w porównaniu z %s' % (self.zwierzeta(), self.day(), self.day2()),
                      self.filename('histogram_nowe', 'png', True))

    def showhist(self, catframe, title, defaultfile):
        try:
            fig = self.histfigure(catframe)
        except Exception as e:
            error_dialog(self, e)
            return
        ChartFrame(self, fig, title, defaultfile).Show()
        self.set_status('Otwarto histogram: ' + title)

    def histfigure(self, catframe):
        """Rysuje histogramy wieku (wszystkie, samice, samce) i zwraca gotową figurę."""
        with matplotlib.rc_context(CHART_FONTS):
            fig = Figure(figsize=(13, 11), constrained_layout=True)
            axes = fig.subplots(3, 2)
            self.gencathist(catframe, axes[0][0], axes[0][1], 'all', self.kotybool)
            self.gencathist(catframe, axes[1][0], axes[1][1], 'f', self.kotybool)
            self.gencathist(catframe, axes[2][0], axes[2][1], 'm', self.kotybool)
        return fig

    def plottrend(self, event):
        self.status('Wczytywanie zapisanych dni…')
        self.run_task(self.loadtrend, (), self.trend_done)

    def loadtrend(self):
        days = sorted(self.dates, key=lambda d: datetime.strptime(d, '%d-%m-%Y'))
        wx.CallAfter(self.postep.SetRange, len(days))
        frames = []
        for i, day in enumerate(days):
            df = pd.read_pickle(self.my_datadir2 + self.prefix() + day + '.kotolizer', compression='bz2')
            frames.append((datetime.strptime(day, '%d-%m-%Y').date(), df[['ID']]))  # zdjęcia nie są potrzebne
            wx.CallAfter(self.postep.SetValue, i + 1)
        return dane.trend(frames)

    def trend_done(self, table):
        wx.CallAfter(self.showtrend, table)
        return 'Wczytano %d zapisanych dni.' % len(table)

    def showtrend(self, table):
        with matplotlib.rc_context(CHART_FONTS):
            fig = Figure(figsize=(13, 9), constrained_layout=True)
            ax1, ax2 = fig.subplots(2, 1, sharex=True)
            ax1.plot(table['Dzień'], table['Liczba'], marker='o', color='k')
            for d, n in zip(table['Dzień'], table['Liczba']):
                ax1.annotate(str(n), (d, n), textcoords='offset points', xytext=(0, 6), ha='center')
            ax1.set_title('Liczba %s na liście' % self.zwierzeta())
            ax1.set_ylabel('Liczba %s' % self.zwierzeta())
            ax1.set_ylim(bottom=0, top=max(table['Liczba']) * 1.15 + 1)
            rest = table.iloc[1:]
            gaps = np.diff([d.toordinal() for d in table['Dzień']])
            width = 0.35 * (min(gaps) if len(gaps) else 1)  # szerokość słupka w dniach
            ax2.bar([pd.Timestamp(d) - pd.Timedelta(days=width / 2) for d in rest['Dzień']], rest['Nowe'], width=width,
                    color='tab:green', label='nowe')
            ax2.bar([pd.Timestamp(d) + pd.Timedelta(days=width / 2) for d in rest['Dzień']], rest['Zniknięte'], width=width,
                    color='tab:red', label='zniknięte')
            ax2.set_title('Nowe i zniknięte w porównaniu z poprzednim zapisanym dniem')
            ax2.set_ylabel('Liczba %s' % self.zwierzeta())
            ax2.legend()
            fig.autofmt_xdate()
        ChartFrame(self, fig, 'Liczba %s w czasie' % self.zwierzeta(),
                   'liczba_%s_w_czasie.png' % ('kotow' if self.kotybool else 'psow'), save_size=(13, 9)).Show()

    def openfile(self, file):
        if platform.system() == 'Darwin':  # macOS
            subprocess.call(('open', file))
        elif platform.system() == 'Windows':
            os.startfile(file)
        else:
            subprocess.call(('xdg-open', file))

    def get_datadir(self) -> pathlib.Path:
        """Katalog na dane programu.

        linux: ~/.local/share, macOS: ~/Library/Application Support, windows: C:/Users/<USER>/AppData/Roaming
        """
        home = pathlib.Path.home()
        if sys.platform == "win32":
            return home / "AppData/Roaming"
        elif sys.platform == "darwin":
            return home / "Library/Application Support"
        return home / ".local/share"


CHART_FONTS = {'font.size': 9, 'axes.titlesize': 11, 'axes.labelsize': 10,
               'xtick.labelsize': 9, 'ytick.labelsize': 9, 'legend.fontsize': 9}


class BrowseFrame(wx.Frame):
    """Tabela zwierząt z wyszukiwarką, sortowaniem po kliknięciu nagłówka i zapisem do Excela."""

    COLUMNS = [('Nr', 45), ('ID', 75), ('Imię', 130), ('Płeć', 70), ('Wiek', 95), ('Przyjęto', 90),
               ('Dni w schronisku', 135), ('Pilne', 50), ('Miejsce', 120), ('Status', 380)]
    NUMERIC = {'Nr': 'Nr', 'Dni w schronisku': 'Dni w schronisku', 'Wiek': '_wiek'}
    URGENT_COLOUR = (248, 221, 215)  # pilne adopcje
    HOVER_COLOUR = (190, 222, 250)   # wiersz, którego zdjęcie jest właśnie wyświetlane

    def __init__(self, parent, table, title, defaultfile, sort_col=None, descending=False):
        wx.Frame.__init__(self, parent, title=title, size=window_size(parent, 1300, 700))
        self.table = table
        self.defaultfile = defaultfile
        self.sort_col = sort_col
        self.descending = descending
        self.visible = table
        panel = wx.Panel(self)
        self.search = wx.SearchCtrl(panel)
        self.search.ShowCancelButton(True)
        self.search.SetDescriptiveText('Szukaj: imię, numer, wiek, płeć, miejsce, status (np. „znaleziony”)…')
        self.search.Bind(wx.EVT_TEXT, lambda e: self.refresh())
        self.search.Bind(wx.EVT_SEARCHCTRL_CANCEL_BTN, lambda e: self.search.SetValue(''))
        self.list = wx.ListCtrl(panel, style=wx.LC_REPORT | wx.LC_SINGLE_SEL)
        for i, (name, width) in enumerate(self.COLUMNS):
            self.list.InsertColumn(i, name, width=width)
        self.list.Bind(wx.EVT_LIST_COL_CLICK, self.on_col_click)
        # podgląd zdjęcia po najechaniu myszką na wiersz; w wersji „generic” (macOS, Linux)
        # wiersze są w osobnym wewnętrznym oknie i HitTest liczy pozycję względem niego
        self.preview = ThumbPopup(self)
        self.hover_row = -1
        rows_window = self.list.GetMainWindow() if hasattr(self.list, 'GetMainWindow') else None
        self.rows_window = rows_window or self.list
        self.rows_window.Bind(wx.EVT_MOTION, self.on_motion)
        self.rows_window.Bind(wx.EVT_LEAVE_WINDOW, self.on_leave)
        self.list.Bind(wx.EVT_MOUSEWHEEL, lambda e: (self.on_leave(e), e.Skip()))
        self.Bind(wx.EVT_ACTIVATE, lambda e: (e.GetActive() or self.on_leave(e), e.Skip()))
        self.info = wx.StaticText(panel, label='')
        save = wx.Button(panel, label='Zapisz widoczne do Excela…')
        Tooltipset(save, 'Zapisuje do pliku Excel zwierzęta widoczne w tabeli, w tej samej kolejności')
        close = wx.Button(panel, label='Zamknij')
        save.Bind(wx.EVT_BUTTON, self.save)
        close.Bind(wx.EVT_BUTTON, lambda e: self.Close())
        bottom = wx.BoxSizer(wx.HORIZONTAL)
        bottom.Add(self.info, 1, wx.ALIGN_CENTER_VERTICAL)
        bottom.Add(save, 0, wx.RIGHT, 8)
        bottom.Add(close, 0)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.search, 0, wx.ALL | wx.EXPAND, 8)
        sizer.Add(self.list, 1, wx.LEFT | wx.RIGHT | wx.EXPAND, 8)
        sizer.Add(bottom, 0, wx.ALL | wx.EXPAND, 8)
        panel.SetSizer(sizer)
        self.Centre()
        self.refresh()

    def on_motion(self, event):
        event.Skip()
        result = self.list.HitTest(self.rows_window.ScreenToClient(wx.GetMousePosition()))
        row = result[0] if isinstance(result, tuple) else result
        if row < 0 or row >= len(self.visible):
            self.on_leave(event)
            return
        if row != self.hover_row:
            self.unhighlight()
            self.hover_row = row
            r = self.visible.iloc[row]
            caption = ('%s (%s)' % (r['Imię'], r['ID']) if r['Imię'] else r['ID']) + ('\n' + r['Wiek'] if r['Wiek'] else '')
            if self.preview.show(r.get('_foto'), caption, wx.GetMousePosition()):
                self.list.SetItemBackgroundColour(row, wx.Colour(*self.HOVER_COLOUR))
            else:
                self.preview.Hide()
        else:
            self.preview.move_near(wx.GetMousePosition())

    def row_colour(self, row):
        return wx.Colour(*self.URGENT_COLOUR) if self.visible.iloc[row]['Pilne'] else self.list.GetBackgroundColour()

    def unhighlight(self):
        """Przywraca zwykły kolor wiersza, który był podświetlony."""
        row = self.hover_row
        if 0 <= row < min(self.list.GetItemCount(), len(self.visible)):
            self.list.SetItemBackgroundColour(row, self.row_colour(row))

    def on_leave(self, event):
        self.unhighlight()
        self.hover_row = -1
        if self.preview.IsShown():
            self.preview.Hide()

    def on_col_click(self, event):
        col = self.COLUMNS[event.GetColumn()][0]
        if col == self.sort_col:
            self.descending = not self.descending
        else:
            self.sort_col, self.descending = col, col == 'Dni w schronisku'
        self.refresh()

    def refresh(self):
        table = self.table
        query = self.search.GetValue().strip().lower()
        if query:
            text = table[[c for c, _ in self.COLUMNS]].astype(str).agg(' '.join, axis=1).str.lower()
            table = table[text.str.contains(query, regex=False)]
        if self.sort_col:
            key = self.NUMERIC.get(self.sort_col, self.sort_col)
            missing = table[key].isna()
            ordered = table[~missing].sort_values(key, ascending=not self.descending, kind='stable')
            table = pd.concat([ordered, table[missing]])  # brak danych zawsze na końcu
        self.on_leave(None)  # zdejmij podświetlenie, zanim zmieni się zawartość tabeli
        self.visible = table
        self.list.DeleteAllItems()
        for i, (_, r) in enumerate(table.iterrows()):
            values = ['' if pd.isna(r[c]) else str(int(r[c]) if isinstance(r[c], float) else r[c])
                      for c, _ in self.COLUMNS]
            self.list.InsertItem(i, values[0])
            for j, v in enumerate(values[1:], start=1):
                self.list.SetItem(i, j, v)
            if r['Pilne']:
                self.list.SetItemBackgroundColour(i, wx.Colour(*self.URGENT_COLOUR))
        for i, (name, _) in enumerate(self.COLUMNS):
            arrow = (' ▼' if self.descending else ' ▲') if name == self.sort_col else ''
            column = self.list.GetColumn(i)
            column.SetText(name + arrow)
            self.list.SetColumn(i, column)
        self.info.SetLabel('Pokazano %d z %d. Kliknij nagłówek kolumny, aby posortować; '
                           'najedź myszką na zwierzę, aby zobaczyć zdjęcie.'
                           % (len(table), len(self.table)))

    def save(self, event):
        with wx.FileDialog(self, 'Zapisz do Excela', defaultFile=self.defaultfile, wildcard='Excel (*.xlsx)|*.xlsx',
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = dialog.GetPath()
        if not os.path.splitext(path)[1]:
            path += '.xlsx'
        try:
            self.visible[[c for c, _ in self.COLUMNS]].to_excel(path, index=False)
        except Exception as e:
            error_dialog(self, e)
            return
        self.info.SetLabel('Zapisano: ' + os.path.basename(path))


class ThumbPopup(wx.PopupWindow):
    """Małe okienko ze zdjęciem zwierzęcia wyświetlane obok kursora."""

    OFFSET = 18  # odległość od kursora, żeby okienko go nie zasłaniało

    def __init__(self, parent):
        wx.PopupWindow.__init__(self, parent, wx.BORDER_SIMPLE)
        self.SetBackgroundColour(wx.WHITE)
        self.bitmap = wx.StaticBitmap(self)
        self.caption = wx.StaticText(self)
        self.caption.SetForegroundColour(wx.BLACK)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.bitmap, 0, wx.ALL, 4)
        sizer.Add(self.caption, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, 4)
        self.SetSizer(sizer)
        self.cache = {}

    def show(self, image, caption, mouse):
        """Pokazuje zdjęcie obok kursora; False, gdy zwierzę nie ma zdjęcia."""
        if image is None or not hasattr(image, 'size'):
            return False
        key = id(image)
        if key not in self.cache:
            rgb = image.convert('RGB')
            self.cache[key] = wx.Bitmap(wx.Image(rgb.size[0], rgb.size[1], rgb.tobytes()))
        self.bitmap.SetBitmap(self.cache[key])
        self.caption.SetLabel(caption)
        self.GetSizer().Fit(self)
        self.move_near(mouse)
        if not self.IsShown():
            self.Show()
        return True

    def move_near(self, mouse):
        w, h = self.GetSize()
        area = wx.Display(max(wx.Display.GetFromPoint(mouse), 0)).GetClientArea()
        x, y = mouse.x + self.OFFSET, mouse.y + self.OFFSET
        if x + w > area.GetRight():
            x = mouse.x - self.OFFSET - w  # przy prawej krawędzi ekranu – na lewo od kursora
        if y + h > area.GetBottom():
            y = mouse.y - self.OFFSET - h  # przy dolnej krawędzi – nad kursorem
        self.SetPosition((x, y))


class ChartCanvas(wx.Panel):
    """Pokazuje figurę matplotlib dopasowaną do rozmiaru okna.

    Figura jest renderowana (Agg) do obrazka dokładnie w rozmiarze panelu,
    w rozdzielczości ekranu (np. 2x na Retinie), więc skaluje się poprawnie
    niezależnie od wersji matplotlib i wxPython.
    """

    def __init__(self, parent, fig):
        wx.Panel.__init__(self, parent, style=wx.FULL_REPAINT_ON_RESIZE)
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.fig = fig
        self.agg = FigureCanvasAgg(fig)
        self.bmp = None
        self.scale = 1.0
        self.timer = wx.CallLater(1, self.render)
        self.SetMinSize((300, 250))
        self.Bind(wx.EVT_SIZE, self.on_size)
        self.Bind(wx.EVT_PAINT, self.on_paint)

    def on_size(self, event):
        self.timer.Start(80)  # przerysuj dopiero, gdy użytkownik skończy zmieniać rozmiar
        event.Skip()

    def render(self):
        w, h = self.GetClientSize()
        if w < 10 or h < 10:
            return
        scale = self.GetContentScaleFactor() if hasattr(self, 'GetContentScaleFactor') else 1.0
        self.fig.set_size_inches(w / 100.0, h / 100.0)
        self.fig.set_dpi(100.0 * scale)
        self.agg.draw()
        W, H = self.agg.get_width_height()
        self.bmp = wx.Bitmap.FromBufferRGBA(W, H, bytes(self.agg.buffer_rgba()))
        self.scale = W / float(w)
        self.Refresh()

    def on_paint(self, event):
        dc = wx.PaintDC(self)
        dc.SetBackground(wx.WHITE_BRUSH)
        dc.Clear()
        if self.bmp is not None:
            dc.SetUserScale(1.0 / self.scale, 1.0 / self.scale)
            dc.DrawBitmap(self.bmp, 0, 0)


class ChartFrame(wx.Frame):
    """Okno z wykresem i przyciskiem do zapisania go jako obraz."""

    def __init__(self, parent, fig, title, defaultfile, save_size=(13, 11)):
        wx.Frame.__init__(self, parent, title=title, size=window_size(parent, 1100, 900))
        self.fig = fig
        self.defaultfile = defaultfile
        self.save_size = save_size  # rozmiar zapisywanego obrazka w calach (niezależny od rozmiaru okna)
        panel = wx.Panel(self)
        self.canvas = ChartCanvas(panel, fig)
        save = wx.Button(panel, label='Zapisz…')
        Tooltipset(save, 'Zapisz wykres jako obraz PNG, PDF lub SVG')
        close = wx.Button(panel, label='Zamknij')
        save.Bind(wx.EVT_BUTTON, self.save)
        close.Bind(wx.EVT_BUTTON, lambda e: self.Close())
        buttons = wx.BoxSizer(wx.HORIZONTAL)
        buttons.Add(save, 0, wx.RIGHT, 8)
        buttons.Add(close, 0)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.canvas, 1, wx.EXPAND)
        sizer.Add(buttons, 0, wx.ALL | wx.ALIGN_RIGHT, 8)
        panel.SetSizer(sizer)
        self.CreateStatusBar()
        self.SetStatusText('Wykres nie jest jeszcze zapisany – kliknij „Zapisz…”, aby zapisać go do pliku.')
        self.Centre()

    def save(self, event):
        with wx.FileDialog(self, 'Zapisz wykres', defaultFile=self.defaultfile,
                           wildcard='PNG (*.png)|*.png|PDF (*.pdf)|*.pdf|SVG (*.svg)|*.svg',
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = dialog.GetPath()
            ext = ('.png', '.pdf', '.svg')[dialog.GetFilterIndex()]
        if not os.path.splitext(path)[1]:
            path += ext
        try:
            self.fig.set_size_inches(*self.save_size)
            self.fig.savefig(path, dpi=150)
        except Exception as e:
            error_dialog(self, e)
            return
        finally:
            self.canvas.render()  # przywróć rozmiar dopasowany do okna
        self.SetStatusText('Zapisano: ' + path)


class CatFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title="Kotolizator 3000 - Analizator stanu zwierząt w schronisku na Paluchu")
        nolog = wx.LogNull()  # bez ostrzeżeń Windows przy wczytywaniu ikony
        try:
            self.SetIcon(wx.Icon(wx.Bitmap(resource_path('icon.ico'), wx.BITMAP_TYPE_ANY)))
        except Exception as e:
            print('Brak ikony: ' + str(e))
        del nolog
        self.p = wx.Panel(self)
        self.nb = GeneralLayout(self.p)
        self.verSizer = wx.BoxSizer(wx.VERTICAL)
        self.verSizer.Add(self.nb, 1, wx.ALL | wx.EXPAND, 0)
        self.statusbar = self.CreateStatusBar()
        self.p.SetSizer(self.verSizer)
        best = self.nb.GetSizer().GetMinSize()
        self.SetClientSize((max(best.width, 560) + 20, best.height + 10))
        self.SetMinClientSize((560, 300))
        self.Centre()
        self.statusbar.SetStatusText('Mateusz Urbańczyk z dedykacją dla Karoliny')
