# -*- mode: python ; coding: utf-8 -*-
# PyInstaller: TimeSaver.exe (konsolsiz, papka ko'rinishida). build.ps1 dan ishga tushadi.
import os

from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs

APP = os.path.abspath(os.path.join(SPECPATH, "..", "app"))

datas = [
    (os.path.join(APP, "panel.html"), "."),
    (os.path.join(APP, "refresh.vbs"), "."),
    (os.path.join(APP, "config.example.ini"), "."),
    (os.path.join(APP, "branding", "logo.svg"), "branding"),
    (os.path.join(APP, "branding", "timesaver.ico"), "branding"),
    (os.path.join(APP, "branding", "logo-128.png"), "branding"),
    (os.path.join(APP, "vba"), "vba"),
]
datas += [(os.path.join(APP, "branding", "tray-%s.png" % k), "branding") for k in ("ok", "busy", "bad", "maint")]
# Panel oynasi (pywebview + Edge WebView2): WebView2 .NET DLL lari, JS fayllari, pythonnet / clr_loader DLL lari.
# (pywebview va pyinstaller-hooks-contrib hook lari ham shuni qiladi — bu yerda aniq ko'rsatilgan, takrorlar birlashtiriladi.)
datas += collect_data_files("webview", subdir="lib") + collect_data_files("webview", subdir="js")
binaries = collect_dynamic_libs("webview") + collect_dynamic_libs("clr_loader")

a = Analysis(
    [os.path.join(APP, "timesaver_tray.py")],
    pathex=[APP],
    datas=datas,
    binaries=binaries,
    hiddenimports=["timesaver_panel", "timesaver_window", "fom_sql", "pystray._win32", "pyodbc",
                   "webview", "webview.platforms.winforms", "webview.platforms.edgechromium", "clr", "clr_loader", "pythonnet"],
    excludes=["tkinter", "unittest", "pydoc", "test"],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="TimeSaver",
    icon=os.path.join(APP, "branding", "timesaver.ico"),
    console=False,
    version=os.path.join(SPECPATH, "version.txt"),
)
coll = COLLECT(exe, a.binaries, a.datas, name="TimeSaver")
