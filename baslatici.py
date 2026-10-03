# -*- coding: utf-8 -*-
"""Temizlik Sunum.exe — başlatıcı.

Açılışta GitHub'daki sürümü kontrol eder; yeniyse uygulama/ klasöründeki dosyaları indirir,
sonra uygulama/app.py'yi çalıştırır. İnternet yoksa mevcut sürümle açılır.
Bozuk bir güncelleme gelirse bir önceki sürüme (uygulama_yedek/) geri döner.
"""
import importlib.util
import json
import os
import shutil
import sys
import time
import traceback
import urllib.request

sys.dont_write_bytecode = True  # uygulama/ içinde __pycache__ oluşmasın

# PyInstaller'ın paketlemesi için app.py'nin kullandığı kütüphaneler burada içe aktarılır
import tkinter as tk
from tkinter import ttk, messagebox, filedialog  # noqa: F401
import PIL.Image, PIL.ImageOps, PIL.ImageTk  # noqa: F401,E401
import pptx  # noqa: F401
from pptx.util import Inches  # noqa: F401
from pptx.enum.shapes import MSO_SHAPE  # noqa: F401
from pptx.enum.text import PP_ALIGN  # noqa: F401
from pptx.dml.color import RGBColor  # noqa: F401
from pptx.opc.packuri import PackURI  # noqa: F401
import lxml.etree  # noqa: F401
import copy, datetime, re, subprocess, uuid, winreg  # noqa: F401,E401

DEPO = os.environ.get("TEMIZLIK_DEPO", "https://raw.githubusercontent.com/mert006/temizlik-sunum/main/uygulama/")

KOK = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
APP = os.path.join(KOK, "uygulama")
YEDEK = os.path.join(KOK, "uygulama_yedek")


def log(mesaj):
    try:
        with open(os.path.join(KOK, "baslatici_kaydi.txt"), "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S ") + mesaj + "\n")
    except Exception:
        pass


def yerel_surum():
    try:
        with open(os.path.join(APP, "surum.json"), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {"surum": 0}


def indir(ad, zaman_asimi=20):
    url = DEPO + ad + "?t=" + str(int(time.time()))
    req = urllib.request.Request(url, headers={"Cache-Control": "no-cache", "User-Agent": "TemizlikSunum"})
    with urllib.request.urlopen(req, timeout=zaman_asimi) as r:
        return r.read()


def guncelle():
    """Yeni sürüm indirildiyse notunu döner, yoksa None."""
    try:
        uzak = json.loads(indir("surum.json", 5).decode("utf-8"))
    except Exception as e:
        log(f"Sürüm kontrolü yapılamadı (internet yok olabilir): {e}")
        return None
    yerel = yerel_surum()
    if int(uzak.get("surum", 0)) <= int(yerel.get("surum", 0)):
        return None

    pencere = tk.Tk()
    pencere.title("Güncelleme")
    pencere.geometry("460x140")
    pencere.resizable(False, False)
    tk.Label(pencere, text=f"Yeni güncelleme bulundu (sürüm {uzak['surum']}).\nİndiriliyor, lütfen bekleyin…",
             font=("Segoe UI", 12), pady=20).pack()
    cubuk = ttk.Progressbar(pencere, mode="determinate", length=380, maximum=len(uzak.get("dosyalar", [])) + 1)
    cubuk.pack()
    pencere.update()

    gecici = os.path.join(KOK, "uygulama_yeni")
    try:
        shutil.rmtree(gecici, ignore_errors=True)
        os.makedirs(gecici)
        for ad in uzak.get("dosyalar", []):
            veri = indir(ad, 60)
            with open(os.path.join(gecici, ad), "wb") as f:
                f.write(veri)
            cubuk.step(1)
            pencere.update()
        with open(os.path.join(gecici, "surum.json"), "w", encoding="utf-8") as f:
            json.dump(uzak, f, ensure_ascii=False, indent=1)
        # mevcut sürümü yedekle, yenisini yerine koy
        shutil.rmtree(YEDEK, ignore_errors=True)
        if os.path.isdir(APP):
            shutil.copytree(APP, YEDEK, ignore=shutil.ignore_patterns("__pycache__"))
        for ad in os.listdir(gecici):
            shutil.copy2(os.path.join(gecici, ad), os.path.join(APP, ad))
        shutil.rmtree(gecici, ignore_errors=True)
        log(f"Sürüm {yerel.get('surum')} → {uzak['surum']} güncellendi")
        return uzak.get("notlar", "")
    except Exception as e:
        log(f"Güncelleme indirilemedi: {e}")
        shutil.rmtree(gecici, ignore_errors=True)
        return None
    finally:
        pencere.destroy()


def uygulamayi_yukle():
    spec = importlib.util.spec_from_file_location("app", os.path.join(APP, "app.py"))
    mod = importlib.util.module_from_spec(spec)
    sys.modules["app"] = mod
    spec.loader.exec_module(mod)
    return mod


def main():
    os.makedirs(APP, exist_ok=True)
    not_ = guncelle()
    try:
        mod = uygulamayi_yukle()
    except Exception:
        log("app.py yüklenemedi:\n" + traceback.format_exc())
        if os.path.isdir(YEDEK):
            for ad in os.listdir(YEDEK):
                if os.path.isfile(os.path.join(YEDEK, ad)):
                    shutil.copy2(os.path.join(YEDEK, ad), os.path.join(APP, ad))
            log("Önceki sürüme geri dönüldü")
            mod = uygulamayi_yukle()
            not_ = None
        else:
            raise
    mod.main(KOK, yerel_surum().get("surum", "?"), not_)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        hata = traceback.format_exc()
        log("HATA:\n" + hata)
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror("Hata", "Program açılırken bir hata oldu. Program klasöründeki "
                                     "'baslatici_kaydi.txt' dosyasını Mert Bey'e gönderin.\n\n" + hata[-500:])
