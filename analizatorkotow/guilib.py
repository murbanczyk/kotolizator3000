import PIL
import docx
import glob
import numpy as np
import os
import pandas as pd
import pathlib
import platform
import re
import shutil
import subprocess
import sys
import wx
import wx.adv
import wx.lib.scrolledpanel as scrolled
# from IPython.core.display import HTML
from PIL import Image
from datetime import date, datetime
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Mm, Cm, Pt
from io import StringIO, BytesIO
from openpyxl.styles import Font
from pkg_resources import resource_filename
from threading import Thread
from .scraper import PaluchScraper, LIST_URLS
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas


def mainGui():
    app = wx.App()

    try:
        splashbmp= os.path.abspath(resource_filename('analizatorkotow', 'splash.png'))
        bitmap = wx.Bitmap(splashbmp)
        splash = wx.adv.SplashScreen(bitmap, wx.adv.SPLASH_CENTRE_ON_SCREEN | wx.adv.SPLASH_TIMEOUT, 4500, None, -1)

    except:
        bitmap = wx.Bitmap('analizatorkotow/splash.png')
        splash = wx.adv.SplashScreen(bitmap, wx.adv.SPLASH_CENTRE_ON_SCREEN | wx.adv.SPLASH_TIMEOUT, 4500, None, -1)

    splash.Show()
    if os.name == 'nt':
        app.locale = wx.Locale(wx.LANGUAGE_POLISH)

    CatFrame().Show()
    app.MainLoop()


def Tooltipset(object, tooltip):
    if float(wx.__version__[0]) < 4:
        object.SetToolTipString(tooltip)
    else:
        object.SetToolTip(tooltip)


class GeneralLayout(scrolled.ScrolledPanel):
    def __init__(self, parent):
        self.my_datadir = self.get_datadir() / "kotolizator" #generates path for application support to store backups
        self.today = date.today().strftime("%d-%m-%Y")
        try:
            self.my_datadir.mkdir(parents=True)
        except FileExistsError:
            pass
        print(self.my_datadir)
        self.my_datadir2 = str(self.my_datadir) + os.sep
        self.kotybool = True
        self.busy = False
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
        grid = wx.GridSizer(2, 2, 6, 6)
        self.word1 = wx.Button(sb, label='Lista zwierząt (Word)')
        Tooltipset(self.word1, 'Dokument Word ze zdjęciem, wiekiem, datą przyjęcia i opisem każdego zwierzęcia.\n'
                               'Pilne adopcje są podświetlone na czerwono, Ochota na Kota na fioletowo.')
        self.genxls = wx.Button(sb, label='Tabelka (Excel)')
        Tooltipset(self.genxls, 'Krótka tabelka: numer, imię, wiek i miejsce pobytu')
        self.histogram1 = wx.Button(sb, label='Histogram wieku')
        Tooltipset(self.histogram1, 'Wykresy wieku zwierząt: wszystkich, samic i samców')
        self.statcat = wx.Button(sb, label='Ile zwierząt w którym miejscu')
        Tooltipset(self.statcat, 'Liczba zwierząt w poszczególnych miejscach (np. Ochota na Kota)')
        for b in (self.word1, self.genxls, self.histogram1, self.statcat):
            grid.Add(b, 0, wx.EXPAND)
        box.Add(grid, 0, wx.ALL | wx.EXPAND, 6)
        main.Add(box, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 8)

        # 4. porównanie
        box = wx.StaticBoxSizer(wx.VERTICAL, self, '4. Porównanie z innym dniem')
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
        main.Add(box, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM | wx.EXPAND, 8)

        # postęp i komunikaty
        self.postep = wx.Gauge(self, range=100, size=(-1, 18), name='Postęp')
        Tooltipset(self.postep, 'Postęp pobierania danych lub tworzenia zestawienia')
        main.Add(self.postep, 0, wx.LEFT | wx.RIGHT | wx.EXPAND, 8)
        self.statustxt = wx.StaticText(self, label='')
        main.Add(self.statustxt, 0, wx.ALL | wx.EXPAND, 8)

        self.word1.Bind(wx.EVT_BUTTON, self.genword)
        self.genxls.Bind(wx.EVT_BUTTON, self.excellgen)
        self.histogram1.Bind(wx.EVT_BUTTON, self.plothist)
        self.statcat.Bind(wx.EVT_BUTTON, self.getcatstat)
        self.wordNew.Bind(wx.EVT_BUTTON, self.genwordNew)
        self.wordOld.Bind(wx.EVT_BUTTON, self.genwordOld)
        self.histnew.Bind(wx.EVT_BUTTON, self.plothistnew)
        self.histold.Bind(wx.EVT_BUTTON, self.plothistold)

        self.SetSizerAndFit(main)
        self.refresh_dates()

    # --- stan interfejsu ---

    def prefix(self):
        return 'Miau' if self.kotybool else 'Hau'

    def zwierzeta(self):
        return 'kotów' if self.kotybool else 'psów'

    def status(self, msg):
        """Komunikat pod paskiem postępu; można wołać z dowolnego wątku."""
        wx.CallAfter(self.statustxt.SetLabel, msg)

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
        day = self.backup.GetStringSelection()
        others = [d for d in self.dates if d != day]
        self.backup2.SetItems(others)
        if others:
            older = [d for d in others if self.dates.index(d) > self.dates.index(day)]
            self.backup2.SetStringSelection(older[0] if older else others[0])
        if not day:
            self.daneinfo.SetLabel('Brak zapisanych danych. Kliknij „Pobierz aktualne dane ze strony”.')
        else:
            try:
                n = len(self.getframecur(self.backup))
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
        if not self.backup2.GetStringSelection():
            self.porinfo.SetLabel('Do porównania potrzebne są dane z co najmniej dwóch dni.')
        else:
            try:
                catframe = self.getframecur(self.backup)
                catframe2 = self.getframecur(self.backup2)
                self.counts = (len(self.comparedf(catframe, catframe2)), len(self.comparedf(catframe2, catframe)))
                self.txtnowe.SetLabel('Nowe zwierzęta: %d' % self.counts[0])
                self.txtznikniete.SetLabel('Zniknięte zwierzęta: %d' % self.counts[1])
                self.porinfo.SetLabel('Porównanie dnia %s z dniem %s.' % (self.backup.GetStringSelection(),
                                                                          self.backup2.GetStringSelection()))
            except Exception as e:
                self.porinfo.SetLabel('Nie udało się porównać: %s' % e)
        self.update_buttons()
        self.Layout()
        self.GetParent().Layout()

    def update_buttons(self):
        idle = not self.busy
        have = idle and bool(self.backup.GetStringSelection())
        for b in (self.word1, self.genxls, self.histogram1, self.statcat):
            b.Enable(have)
        compare = have and self.counts is not None
        self.wordNew.Enable(compare and self.counts[0] > 0)
        self.histnew.Enable(compare and self.counts[0] > 0)
        self.wordOld.Enable(compare and self.counts[1] > 0)
        self.histold.Enable(compare and self.counts[1] > 0)
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
            self.statustxt.SetLabel('Wystąpił błąd.')
            wx.MessageDialog(self, str(error), 'Błąd', wx.OK | wx.ICON_ERROR).ShowModal()
        elif msg:
            self.statustxt.SetLabel(msg)

    def ask_path(self, title, wildcard, default):
        with wx.FileDialog(self, title, wildcard=wildcard, defaultFile=default,
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return None
            return fileDialog.GetPath()

    def filename(self, kind, ext, with_compare=False):
        name = kind + '_' + ('koty' if self.kotybool else 'psy') + '_' + self.backup.GetStringSelection()
        if with_compare:
            name += '_vs_' + self.backup2.GetStringSelection()
        return name + '.' + ext

    # --- zestawienia ---

    def getcatstat(self,event):
        xlscat=self.gencatxls()
        wx.MessageDialog(self, xlscat.groupby(['Miejsce']).size().to_string(),
                         'Ile %s w którym miejscu (%s)' % (self.zwierzeta(), self.backup.GetStringSelection()),
                         ).ShowModal()
    def gencatxls(self):

        catframe = self.getframecur(self.backup)
        rows = []
        opo=0
        for a, b in catframe.iterrows():
            opo+=1
            Imie = str(b['Imię']).split('\n')[0]
            Miejsce = str(b['Przyjęcie']).split('\n')[-1]
            miejscatocheck = ['Azyl', 'Ochota', 'właściciela']
            if not any(ext in Miejsce for ext in miejscatocheck):
                Miejsce = 'brak informacji'
            Wiek = str(b['Wiek'])
            if 'mies' in Wiek:
                try:
                    mies=int(Wiek.split()[0])
                except ValueError:
                    mies = None
                if mies is not None and mies<4:
                    Imie +='(m)'
            rows.append({'Nr':opo, 'ID': b['ID'], 'Imię': Imie, 'Wiek': Wiek, 'Miejsce': Miejsce})
        return pd.DataFrame(rows, columns=['Nr', 'ID', 'Imię', 'Wiek', 'Miejsce'])
    def excellgen(self,event):
        xlscat=self.gencatxls()
        # xlscat.style.applymap('font-weight: bold',
        #                        subset=pd.IndexSlice[xlscat.index[xlscat.Miejsce == 'Ochota na Kota'], :])
        # print(xlscat.style.applymap('font-weight: bold',
        #                        subset=pd.IndexSlice[xlscat.index[xlscat.Miejsce == 'Ochota na Kota'], :]))
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
            self.statustxt.SetLabel('Zapisano: ' + os.path.basename(pathname))
            self.openfile(pathname)

    def gencathist(self,catframe, ax1, ax2, sex,kotbool):


        lataM = []
        lataF = []
        lata = []
        for k in zip(catframe['Wiek'], catframe['Imię']):
            if 'samiec' in k[1]:
                if 'rok' in k[0] or 'lat' in k[0]:
                    lataM.append(int(k[0][0:2]) * 12)
                else:
                    lataM.append(int(k[0][0:2]))
            elif 'samica' in k[1]:
                if 'rok' in k[0] or 'lat' in k[0]:
                    lataF.append(int(k[0][0:2]) * 12)
                else:
                    lataF.append(int(k[0][0:2]))

            if 'rok' in k[0] or 'lat' in k[0]:
                lata.append(int(k[0][0:2]) * 12)
            else:
                lata.append(int(k[0][0:2]))


        bins = np.arange(0, 13) - 0.5

        if sex == 'all':
            lataR = np.floor(np.array(lata) / 12)
            oldest = int(np.max(lataR)) if len(lataR) else 0
            bins2 = np.arange(0, oldest + 2) - 0.5

            a1, _, patches = ax1.hist(np.array(lata), bins, linewidth=1, edgecolor='w', color='k')
            a2, _, patches2 = ax2.hist(lataR, bins2, linewidth=1, edgecolor='w', color='k')
            if kotbool:
                ax1.set_title('Koty do 1 roku, suma: ' + str(round(sum(a1))))
                ax2.set_title('Koty wszystkie, suma: ' + str(round(sum(a2))) + ', starsze od 8 lat: ' + str(
                    len(np.where(lataR >= 8)[0])))
            else:
                ax1.set_title('Psy do 1 roku, suma: ' + str(round(sum(a1))))
                ax2.set_title('Psy wszystkie, suma: ' + str(round(sum(a2))) + ', starsze od 8 lat: ' + str(
                    len(np.where(lataR >= 8)[0])))

        elif sex == 'f':
            lataR = np.floor(np.array(lataF) / 12)

            oldest = int(np.max(lataR)) if len(lataR) else 0
            bins2 = np.arange(0, oldest + 2) - 0.5

            a1, _, patches = ax1.hist(np.array(lataF), bins, linewidth=1, edgecolor='w', color='pink')
            a2, _, patches2 = ax2.hist(lataR, bins2, linewidth=1, edgecolor='w', color='pink')

            if kotbool:

                ax1.set_title('Kotki do 1 roku, suma: ' + str(round(sum(a1))))
                ax2.set_title('Kotki wszystkie, suma: ' + str(round(sum(a2))) + ', starsze od 8 lat: ' + str(
                    len(np.where(lataR >= 8)[0])))
            else:
                ax1.set_title('Suczki do 1 roku, suma: ' + str(round(sum(a1))))
                ax2.set_title('Suczki wszystkie, suma: ' + str(round(sum(a2))) + ', starsze od 8 lat: ' + str(
                    len(np.where(lataR >= 8)[0])))


        else:
            lataR = np.floor(np.array(lataM) / 12)

            oldest = int(np.max(lataR)) if len(lataR) else 0
            bins2 = np.arange(0, oldest + 2) - 0.5

            a1, _, patches = ax1.hist(np.array(lataM), bins, linewidth=1, edgecolor='w', color='b')
            # ax1.set_title('Kocury do 1 roku, suma: ' + str(round(sum(a1))))
            a2, _, patches2 = ax2.hist(lataR, bins2, linewidth=1, edgecolor='w', color='b')

            # ax2.set_title('Kocury wszystkie, suma: ' + str(round(sum(a2))) + ';starsze of 8lat:' + str(
            # len(np.where(lataR >= 8)[0])))
            if kotbool:

                ax1.set_title('Kocury do 1 roku, suma: ' + str(round(sum(a1))))
                ax2.set_title('Kocury wszystkie, suma: ' + str(round(sum(a2))) + ', starsze od 8 lat: ' + str(
                    len(np.where(lataR >= 8)[0])))
            else:
                ax1.set_title('Psy (samce) do 1 roku, suma: ' + str(round(sum(a1))))
                ax2.set_title('Psy (samce) wszystkie, suma: ' + str(round(sum(a2))) + ', starsze od 8 lat: ' + str(
                    len(np.where(lataR >= 8)[0])))

        ax1.set_xticks(range(0, 12))
        ymax=[]
        for pp in patches:
            # print(pp._x0)
            # print(pp._width)
            x = (pp._x0) + (pp._width) / 2
            y = pp._height  + 0.3
            ymax.append(y)
            ax1.text(x, y, int(pp._height), ha='center')
        ax1.set_ylim([0,max(ymax)+0.5])

        #     ax1.set_title('Koty do 1 roku, suma: '  + str(round(sum(a1))))
        ax1.set_xlabel('miesiące')
        lataR = np.floor(np.array(lata) / 12)

        #     oldest=int(np.max(lataR))
        #     bins = np.arange(0,oldest+2) -0.5

        #     a2, _, patches=ax2.hist(lataR,bins, linewidth=1, edgecolor='w')
        ax2.set_xticks(range(0, oldest + 1))

        ax2.set_xlabel('lata')
        if kotbool:
            ax2.set_ylabel('Liczba kotów')
            ax1.set_ylabel('Liczba kotów')
        else:
            ax2.set_ylabel('Liczba psów')
            ax1.set_ylabel('Liczba psów')
        ymax=[]

        for pp in patches2:
            x = (pp._x0) + (pp._width) / 2
            y = pp._height + 0.3
            ymax.append(y)

            ax2.text(x, y, int(pp._height), ha='center')
            #
            # x = (pp._x0 + pp._x1) / 2
            # y = pp._y1 + 0.3
            # ax2.text(x, y, int(pp._y1), ha='center')
        ax2.set_ylim([0,max(ymax)+0.5])

        return lata

    def plothist(self,event):
        self.showhist(self.getframecur(self.backup),
                      'Wiek %s – %s' % (self.zwierzeta(), self.backup.GetStringSelection()),
                      self.filename('histogram', 'png'))

    def plothistold(self,event):
        Zniknietekoty = self.comparedf(self.getframecur(self.backup2), self.getframecur(self.backup))
        self.showhist(Zniknietekoty,
                      'Wiek znikniętych %s – %s w porównaniu z %s' % (
                          self.zwierzeta(), self.backup.GetStringSelection(), self.backup2.GetStringSelection()),
                      self.filename('histogram_znikniete', 'png', True))

    def plothistnew(self,event):
        Nowekoty = self.comparedf(self.getframecur(self.backup), self.getframecur(self.backup2))
        self.showhist(Nowekoty,
                      'Wiek nowych %s – %s w porównaniu z %s' % (
                          self.zwierzeta(), self.backup.GetStringSelection(), self.backup2.GetStringSelection()),
                      self.filename('histogram_nowe', 'png', True))

    def showhist(self, catframe, title, defaultfile):
        try:
            fig = self.histfigure(catframe)
        except Exception as e:
            wx.MessageDialog(self, str(e), 'Błąd', wx.OK | wx.ICON_ERROR).ShowModal()
            return
        HistFrame(self, fig, title, defaultfile).Show()
        self.statustxt.SetLabel('Otwarto histogram: ' + title)

    def histfigure(self, catframe):
        """Rysuje histogramy wieku (wszystkie, samice, samce) i zwraca gotową figurę."""
        sizes = {'font.size': 9, 'axes.titlesize': 11, 'axes.labelsize': 10,
                 'xtick.labelsize': 9, 'ytick.labelsize': 9, 'legend.fontsize': 9}
        with matplotlib.rc_context(sizes):
            fig = Figure(figsize=(13, 11), constrained_layout=True)
            axes = fig.subplots(3, 2)
            self.gencathist(catframe, axes[0][0], axes[0][1], 'all', self.kotybool)
            self.gencathist(catframe, axes[1][0], axes[1][1], 'f', self.kotybool)
            self.gencathist(catframe, axes[2][0], axes[2][1], 'm', self.kotybool)
        return fig

    def openfile(self, file):
        if platform.system() == 'Darwin':  # macOS
            subprocess.call(('open',file))
        elif platform.system() == 'Windows':  # Windows
            os.startfile(file)
        else:  # linux variants
            subprocess.call(('xdg-open',file))


    def switchkotpies(self, event):
        self.kotybool = self.gatunek.GetSelection() == 0
        self.statustxt.SetLabel('')
        self.postep.SetValue(0)
        self.refresh_dates()

    # --- pobieranie danych ---

    def parsepaluch(self, event):
        self.status('Pobieranie listy %s…' % self.zwierzeta())
        self.run_task(self.parsepaluch2, (), self.parse_done)

    def parse_done(self, result):
        n, fetched = result
        wx.CallAfter(self.refresh_dates, self.today)
        return 'Pobrano dane: %d %s (nowych lub odświeżonych stron zwierząt: %d).' % (n, self.zwierzeta(), fetched)

    def parsepaluch2(self):
        species = 'koty' if self.kotybool else 'psy'
        scraper = PaluchScraper(self.my_datadir2 + 'cache.kotolizer')
        animals = scraper.get_animals(species)
        wx.CallAfter(self.postep.SetRange, len(animals))
        rows = []
        fetched = 0
        for count, a in enumerate(animals):
            self.status('Zwierzę %d z %d: %s' % (count + 1, len(animals), a['name'] or a['number']))
            details, new = scraper.get_details(a['link'], referer=LIST_URLS[species])
            if new:
                fetched += 1
                if fetched % 20 == 0:
                    scraper.save_cache()
            details = details or {}
            CatImage = scraper.get_thumbnail(a['image'], referer=a['link'])
            location = details.get('location', '')
            if location == 'Brak danych':
                location = ''
            if 'Ochota' in location and 'Ochota na Kota' not in location:
                location += ' (Ochota na Kota)'
            CatIn = details.get('date', '')
            CatIn += '\n\nZdjęć:' + str(details.get('photos', ''))
            CatIn += '\n\nRasa: ' + a['breed']
            if a['urgent']:
                CatIn += '\n\nPILNE'
            if location:
                CatIn += '\n\n' + location  # ostatnia linia jest używana w statystyce "Miejsca"
            rows.append({'Nr': count + 1, 'ID': int(a['number']),
                         'Imię': a['name'] + '\n' + details.get('sex', a['sex']),
                         'Wiek': a['age'], 'Przyjęcie': CatIn, 'Foto': CatImage,
                         'Ogłoszenie': details.get('desc', '')})
            wx.CallAfter(self.postep.SetValue, count + 1)
        scraper.save_cache()
        print('Zwierząt: %d, pobranych stron zwierząt: %d' % (len(animals), fetched))
        catframe = pd.DataFrame(rows, columns=['Nr', 'ID', 'Imię', 'Wiek', 'Przyjęcie', 'Foto', 'Ogłoszenie'])
        catframe.to_pickle(self.my_datadir2 + self.prefix() + self.today + '.kotolizer', compression='bz2')
        self.frames.pop((self.prefix(), self.today), None)
        return len(animals), fetched

    def getframecur(self, b):
        key = (self.prefix(), b.GetStringSelection())
        if key not in self.frames:
            df = pd.read_pickle(self.my_datadir2 + key[0] + key[1] + '.kotolizer', compression='bz2')
            self.frames[key] = df.sort_values('ID')
        return self.frames[key]

    def comparedf(self,one, two):
        kiki = pd.concat([one, two]).drop_duplicates(subset='ID', keep=False)
        new = pd.merge(one, kiki, how='inner', on='ID')
        df3 = new[['Nr_x', 'ID', 'Imię_x', 'Wiek_x', 'Przyjęcie_x', 'Foto_x', 'Ogłoszenie_x']].rename(
            columns={'Przyjęcie_x': 'Przyjęcie',
                     'Ogłoszenie_x': 'Ogłoszenie',
                     'Nr_x': 'Nr', 'Imię_x': 'Imię',
                     'Wiek_x': 'Wiek', 'Foto_x': 'Foto'})
        return df3
    def genword(self, event):
        pathname = self.ask_path('Zapisz listę zwierząt', 'Word (*.docx)|*.docx', self.filename('lista', 'docx'))
        if pathname:
            self.status('Tworzenie dokumentu Word…')
            self.run_task(self.createwordfile, (pathname, self.getframecur(self.backup)),
                          lambda r: 'Zapisano: ' + os.path.basename(pathname))

    def genwordNew(self, event):
        Nowekoty = self.comparedf(self.getframecur(self.backup), self.getframecur(self.backup2))
        pathname = self.ask_path('Zapisz listę nowych zwierząt', 'Word (*.docx)|*.docx',
                                 self.filename('nowe', 'docx', True))
        if pathname:
            self.status('Tworzenie dokumentu Word…')
            self.run_task(self.createwordfile, (pathname, Nowekoty), lambda r: 'Zapisano: ' + os.path.basename(pathname))

    def genwordOld(self, event):
        Zniknietekoty = self.comparedf(self.getframecur(self.backup2), self.getframecur(self.backup))
        pathname = self.ask_path('Zapisz listę znikniętych zwierząt', 'Word (*.docx)|*.docx',
                                 self.filename('znikniete', 'docx', True))
        if pathname:
            self.status('Tworzenie dokumentu Word…')
            self.run_task(self.createwordfile, (pathname, Zniknietekoty), lambda r: 'Zapisano: ' + os.path.basename(pathname))

    def createwordfile(self, pathname,catframe):

        try:
            os.mkdir(self.my_datadir2+'thumb')
        except:
            pass
        # open an existing document
        doc = docx.Document()

        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(8)

        section = doc.sections[0]
        section.page_height = Mm(210)
        section.page_width = Mm(297)
        t = doc.add_table(catframe.shape[0] + 1, catframe.shape[1])

        # add the header rows.
        # wx.CallAfter(self.postep.SetRange, 1)

        # pub.sendMessage('postep1', msg=(catframe.shape[0]))
        # pub.sendMessage('postep2', msg=0)
        wx.CallAfter(self.postep.SetRange, (catframe.shape[0]))

        for j in range(catframe.shape[-1]):
            t.cell(0, j).text = catframe.columns[j]

        # add the rest of the data frame
        for i in range(catframe.shape[0]):
            # pub.sendMessage('postep2', msg=i+1)
            wx.CallAfter(self.postep.SetValue, i + 1)

            for j in range(catframe.shape[-1]):

                if j == 5:
                    catframe.values[i, j].save(self.my_datadir2+'thumb/' + str(i) + '.jpg', "JPEG")
                    paragraph = t.cell(i + 1, j).paragraphs[0]
                    run = paragraph.add_run()
                    run.add_picture(self.my_datadir2+'thumb/' + str(i) + '.jpg')
                elif j == 0:
                    t.cell(i + 1, j).text = str(i + 1)


                else:
                    t.cell(i + 1, j).text = str(catframe.values[i, j])

                if 'PILNE' in catframe.values[i, 4]:
                    t.cell(i + 1, j)._tc.get_or_add_tcPr().append(
                        parse_xml(r'<w:shd {} w:fill="f8ddd7"/>'.format(nsdecls('w'))))
            if 'Ochota na Kota' in catframe.values[i, 4]:
                t.cell(i + 1, 2)._tc.get_or_add_tcPr().append(
                    parse_xml(r'<w:shd {} w:fill="f0d7f8"/>'.format(nsdecls('w'))))

        ####f0d7f8
        t.allow_autofit = True
        t.columns[0].width = Cm(0.7)
        t.columns[1].width = Cm(1.5)
        t.columns[2].width = Cm(1.5)
        t.columns[3].width = Cm(1.2)
        t.columns[4].width = Cm(2)
        t.columns[5].width = Cm(6)
        t.columns[6].width = Cm(10)
        doc.save(pathname)
        shutil.rmtree(self.my_datadir2+'thumb/')
        self.openfile(pathname)


    def get_datadir(self) -> pathlib.Path:

        """
        Returns a parent directory path
        where persistent application data can be stored.

        # linux: ~/.local/share
        # macOS: ~/Library/Application Support
        # windows: C:/Users/<USER>/AppData/Roaming
        """

        home = pathlib.Path.home()

        if sys.platform == "win32":
            return home / "AppData/Roaming"
        elif sys.platform == "linux":
            return home / ".local/share"
        elif sys.platform == "darwin":
            return home / "Library/Application Support"

    # create your program's directory


class HistFrame(wx.Frame):
    """Okno z histogramem i przyciskiem do zapisania go jako obraz."""

    def __init__(self, parent, fig, title, defaultfile):
        wx.Frame.__init__(self, parent, title=title, size=(1100, 900))
        self.fig = fig
        self.defaultfile = defaultfile
        panel = wx.Panel(self)
        self.canvas = FigureCanvas(panel, -1, fig)
        self.canvas.SetMinSize((300, 250))  # wykres dopasowuje się do rozmiaru okna
        save = wx.Button(panel, label='Zapisz…')
        Tooltipset(save, 'Zapisz histogram jako obraz PNG, PDF lub SVG')
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
        self.SetStatusText('Histogram nie jest jeszcze zapisany – kliknij „Zapisz…”, aby zapisać go do pliku.')
        self.Centre()

    def save(self, event):
        with wx.FileDialog(self, 'Zapisz histogram', defaultFile=self.defaultfile,
                           wildcard='PNG (*.png)|*.png|PDF (*.pdf)|*.pdf|SVG (*.svg)|*.svg',
                           style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT) as dialog:
            if dialog.ShowModal() == wx.ID_CANCEL:
                return
            path = dialog.GetPath()
        ext = ('.png', '.pdf', '.svg')[dialog.GetFilterIndex()]
        if not os.path.splitext(path)[1]:
            path += ext
        try:
            self.fig.savefig(path, dpi=150)
        except Exception as e:
            wx.MessageDialog(self, str(e), 'Błąd', wx.OK | wx.ICON_ERROR).ShowModal()
            return
        self.SetStatusText('Zapisano: ' + path)


class CatFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title="Kotolizator 3000 - Analizator stanu zwierząt w schronisku na Paluchu",
                          )
        nolog = wx.LogNull()  # windows alert fix

        if platform.system() == 'Linux' and getattr(sys, 'frozen', False):
            IconPath = sys._MEIPASS + '/icon.ico'
            if float(wx.__version__[0]) < 4:
                self.SetIcon(wx.IconFromBitmap(wx.Bitmap(IconPath, wx.BITMAP_TYPE_ANY)))
            else:
                self.SetIcon(wx.Icon(wx.Bitmap(IconPath, wx.BITMAP_TYPE_ANY)))
        else:
            try:  # Trend.Resources is for code used from whl distribution
                IconPath = os.path.abspath(resource_filename('analizatorkotow', 'icon.ico'))
                print(IconPath)
                if float(wx.__version__[0]) < 4:
                    self.SetIcon(wx.IconFromBitmap(wx.Bitmap(IconPath, wx.BITMAP_TYPE_ANY)))
                else:
                    self.SetIcon(wx.Icon(wx.Bitmap(IconPath, wx.BITMAP_TYPE_ANY)))
            except:  # If code is used from source code
                IconPath = 'icon.ico'
                if float(wx.__version__[0]) < 4:
                    self.SetIcon(wx.IconFromBitmap(wx.Bitmap(IconPath, wx.BITMAP_TYPE_ANY)))
                else:
                    self.SetIcon(wx.Icon(wx.Bitmap(IconPath, wx.BITMAP_TYPE_ANY)))
        del nolog
        self.p = wx.Panel(self)  # creates panel
        self.nb = GeneralLayout(self.p)
        self.verSizer = wx.BoxSizer(wx.VERTICAL)
        self.verSizer.Add(self.nb, 1, wx.ALL | wx.EXPAND, 0)
        self.firstimeclick = True
        self.statusbar = self.CreateStatusBar()
        self.p.SetSizer(self.verSizer)
        best = self.nb.GetSizer().GetMinSize()
        self.SetClientSize((max(best.width, 560) + 20, best.height + 10))
        self.SetMinClientSize((560, 300))
        self.Centre()
        self.statusbar.SetStatusText(
            'Mateusz Urbańczyk z dedykacją dla Karoliny')

    def updatecontrolers(self, msg):
        e = wx.SizeEvent(self.GetSize())
        self.ProcessEvent(e)
        self.Fit()
