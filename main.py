#### ZD MENÜ + ZD2027 + ZD Takip + ZD Kira + ZD Yapılacaklar (tek dosya)
####
#### Başlatma: bu dosyayı çalıştırın -> ZD Menüsü açılır, butonlardan programlar açılır.
#### Programlardaki ÇIKIŞ butonu (kısa veya uzun basma) menüye geri döndürür.
#### Menüdeki "Çıkış" butonu uygulamayı tamamen kapatır.
####
#### APK için: buildozer kullanırken bu dosyanın adı main.py olmalıdır.
####
import glob
import json
import math
import os
import shutil
import sqlite3
import sys
import traceback
from datetime import datetime              # ZDTakip ve ZDtodo: datetime.now()
import datetime as _dtmod                  # ZD2027 ve ZDKira: datetime.datetime.now() -> _dtmod.datetime.now()

import kivy
kivy.require("2.0.0")
from kivy.app import App
from kivy.clock import Clock
from kivy.core.window import Window
from kivy.graphics import Color, Line, Rectangle, RoundedRectangle
from kivy.lang import Builder
from kivy.logger import Logger
from kivy.metrics import dp
from kivy.properties import ObjectProperty
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.label import Label
from kivy.uix.modalview import ModalView
from kivy.uix.popup import Popup
from kivy.uix.scrollview import ScrollView
from kivy.uix.spinner import Spinner
from kivy.uix.switch import Switch
from kivy.uix.tabbedpanel import TabbedPanel, TabbedPanelItem
from kivy.uix.textinput import TextInput
from kivy.uix.widget import Widget
from kivy.utils import get_color_from_hex, platform

Window.fullscreen = 'auto'

# ============================================================
# AYARLAR: Her programın çalışma klasörü
# ============================================================
try:
    ZD_VERI_KOKU = os.path.dirname(os.path.abspath(__file__))
except NameError:
    ZD_VERI_KOKU = os.getcwd()

ZD_PROGRAM_KLASORLERI = {
    "ZD2027":  "",
    "ZDTakip": "",
    "ZDKira":  "",
    "ZDtodo":  "",
}

# ============================================================
# EXCEL DESTEĞİ: pandas varsa o kullanılır (Pydroid), yoksa (APK) openpyxl tabanlı
# küçük bir yedek devreye girer. Programlar yalnızca burada tanımlı işlevleri kullanır.
# ============================================================
class _MiniDataFrame:
    def __init__(self, data=None):
        self.columns = []
        self._rows = []
        if isinstance(data, dict):
            cols = list(data.keys())
            n = max((len(v) for v in data.values()), default=0)
            self.columns = cols
            self._rows = [[(data[c][i] if i < len(data[c]) else None) for c in cols] for i in range(n)]
        elif data is not None:
            data = [r for r in data if isinstance(r, dict)]
            cols = []
            for r in data:
                for k in r:
                    if k not in cols:
                        cols.append(k)
            self.columns = cols
            self._rows = [[r.get(c) for c in cols] for r in data]

    def __len__(self):
        return len(self._rows)

    def __getitem__(self, col):
        if col not in self.columns:
            raise KeyError(col)
        i = self.columns.index(col)
        return _MiniSeries(r[i] for r in self._rows)

    def iterrows(self):
        for n, r in enumerate(self._rows):
            yield n, dict(zip(self.columns, r))

    def to_excel(self, target, sheet_name='Sheet1', index=True, **kwargs):
        from openpyxl import Workbook
        writer = target if isinstance(target, _MiniExcelWriter) else None
        if writer is not None:
            ws = writer._yeni_sayfa(sheet_name)
        else:
            wb = Workbook()
            ws = wb.active
            ws.title = str(sheet_name)[:31]
        if self.columns:
            ws.append(([''] if index else []) + [_hucre(c) for c in self.columns])
        for n, r in enumerate(self._rows):
            ws.append(([n] if index else []) + [_hucre(v) for v in r])
        if writer is None:
            wb.save(target)

class _MiniSeries(list):
    def tolist(self):
        return list(self)

def _hucre(v):
    if v is None:
        return None
    if isinstance(v, float) and v != v:
        return None
    if isinstance(v, bool) or isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
        return ILLEGAL_CHARACTERS_RE.sub('', v)
    if isinstance(v, (datetime, _dtmod.date)):
        return v
    return str(v)

class _MiniExcelWriter:
    def __init__(self, path, engine=None, **kwargs):
        from openpyxl import Workbook
        self.path = path
        self.wb = Workbook()
        self._ilk = True
        self._kapali = False

    def _yeni_sayfa(self, ad):
        ad = str(ad)[:31]
        if self._ilk:
            ws = self.wb.active
            ws.title = ad
            self._ilk = False
        else:
            ws = self.wb.create_sheet(ad)
        return ws

    def close(self):
        if not self._kapali:
            self._kapali = True
            self.wb.save(self.path)

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

class _MiniExcelFile:
    def __init__(self, path):
        from openpyxl import load_workbook
        self.wb = load_workbook(path, data_only=True)
        self.sheet_names = list(self.wb.sheetnames)

def _mini_read_excel(io, sheet_name=0, **kwargs):
    xls = io if isinstance(io, _MiniExcelFile) else _MiniExcelFile(io)
    ws = xls.wb[sheet_name] if isinstance(sheet_name, str) else xls.wb.worksheets[sheet_name]
    rows = [[None if (v == '' or (isinstance(v, float) and v != v)) else v for v in r]
            for r in ws.iter_rows(values_only=True)]
    while rows and all(v is None for v in rows[-1]):
        rows.pop()
    if not rows:
        return _MiniDataFrame()
    baslik, govde = rows[0], rows[1:]
    n = len(baslik)
    while n > 0 and baslik[n - 1] is None and all((r[n - 1] if len(r) >= n else None) is None for r in govde):
        n -= 1
    df = _MiniDataFrame()
    df.columns = [h if h is not None else f"Unnamed: {i}" for i, h in enumerate(baslik[:n])]
    df._rows = [(list(r) + [None] * n)[:n] for r in govde]
    return df

def _mini_isna(x):
    return x is None or (isinstance(x, float) and x != x)

class _MiniPandas:
    DataFrame = _MiniDataFrame
    ExcelWriter = _MiniExcelWriter
    ExcelFile = _MiniExcelFile
    read_excel = staticmethod(_mini_read_excel)
    isna = staticmethod(_mini_isna)
    notna = staticmethod(lambda x: not _mini_isna(x))

try:
    import pandas as pd
except ImportError:
    pd = _MiniPandas

def get_external_storage_path():
    if platform == 'android':
        try:
            from android.storage import primary_external_storage_path
            return primary_external_storage_path()
        except ImportError:
            return os.path.expanduser('~')
    return os.path.expanduser('~')

# ============================================================
# BÖLÜM 5: ZD MENÜ (uygulamanın ana giriş noktası)
# ============================================================
if __name__ == "__main__":
    from kivy.app import App
    print("ZD Menü Başlatılıyor...")
