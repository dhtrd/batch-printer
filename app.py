# -*- coding: utf-8 -*-
"""
الطابعة الجماعية الآمنة — تطبيق سطح المكتب (واجهة أوركيد)
=========================================================
يعرض واجهة أوركيد (HTML) داخل نافذة سطح مكتب عبر pywebview،
ويشغّل خلفية Python حقيقية للطباعة الجماعية.

الأمان: يعمل محليًا بالكامل، لا يتصل بالشبكة، ويعطّل ماكرو أوفيس عند الطباعة.
النظام: Windows. الترخيص: مجاني ومفتوح المصدر.
"""

import os
import sys
import json
import time
import threading

import webview  # pywebview

try:
    import win32print
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

APP_VERSION = "1.1.0"

# ---------- الصيغ ----------
PDF_EXTS = {".pdf"}
WORD_EXTS = {".doc", ".docx", ".rtf", ".txt", ".odt", ".htm", ".html"}
EXCEL_EXTS = {".xls", ".xlsx", ".xlsm", ".csv", ".ods"}
PPT_EXTS = {".ppt", ".pptx", ".odp"}
IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
ALL_EXTS = PDF_EXTS | WORD_EXTS | EXCEL_EXTS | PPT_EXTS | IMAGE_EXTS


def ext_category(path):
    e = os.path.splitext(path)[1].lower()
    if e in PDF_EXTS:   return "PDF"
    if e in WORD_EXTS:  return "Word"
    if e in EXCEL_EXTS: return "Excel"
    if e in PPT_EXTS:   return "PowerPoint"
    if e in IMAGE_EXTS: return "Image"
    return None


def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def resource_path(name):
    base = getattr(sys, "_MEIPASS", None) or app_dir()
    return os.path.join(base, name)


def human(n):
    n = float(n)
    for u in ["B", "KB", "MB", "GB"]:
        if n < 1024:
            return ("%.0f %s" % (n, u)) if u == "B" else ("%.1f %s" % (n, u))
        n /= 1024.0
    return "%.1f TB" % n


# ===================== محرك الطباعة الآمن =====================
class PrintEngine:
    MSO_SECURITY_FORCE_DISABLE = 3

    def __init__(self, printer_name, copies=1, page_range="", dpi=300, log=None, cancel_event=None):
        self.printer = printer_name
        self.copies = max(1, int(copies))
        self.page_range = (page_range or "").strip()
        self.dpi = int(dpi)
        self.log = log or (lambda m: None)
        self.cancel = cancel_event or threading.Event()
        self._orig_default = None

    def __enter__(self):
        if HAS_WIN32:
            try:
                self._orig_default = win32print.GetDefaultPrinter()
                if self.printer and self.printer != self._orig_default:
                    win32print.SetDefaultPrinter(self.printer)
            except Exception as e:
                self.log("تنبيه: تعذّر ضبط الطابعة الافتراضية (%s)" % e)
        return self

    def __exit__(self, *exc):
        if HAS_WIN32 and self._orig_default:
            try:
                win32print.SetDefaultPrinter(self._orig_default)
            except Exception:
                pass
        return False

    def print_file(self, path):
        cat = ext_category(path)
        if cat == "PDF":        return self._print_pdf(path)
        if cat == "Word":       return self._print_word(path)
        if cat == "Excel":      return self._print_excel(path)
        if cat == "PowerPoint": return self._print_ppt(path)
        if cat == "Image":      return self._print_image(path)
        raise ValueError("صيغة غير مدعومة")

    @staticmethod
    def _parse_pages(rng, count):
        if not rng:
            return list(range(count))
        out = []
        for part in rng.replace(" ", "").split(","):
            if not part:
                continue
            if "-" in part:
                a, b = part.split("-", 1)
                a = int(a) if a else 1
                b = int(b) if b else count
                for p in range(a, b + 1):
                    if 1 <= p <= count:
                        out.append(p - 1)
            else:
                p = int(part)
                if 1 <= p <= count:
                    out.append(p - 1)
        return out or list(range(count))

    def _print_pdf(self, path):
        import fitz
        from PIL import Image
        doc = fitz.open(path)
        try:
            pages = self._parse_pages(self.page_range, doc.page_count)
            zoom = self.dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            images = []
            for i in pages:
                if self.cancel.is_set():
                    break
                page = doc.load_page(i)
                pix = page.get_pixmap(matrix=mat, alpha=False)
                images.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
        finally:
            doc.close()
        for _ in range(self.copies):
            if self.cancel.is_set():
                break
            self._draw_images(images, os.path.basename(path))
        return True

    def _print_image(self, path):
        from PIL import Image
        img = Image.open(path)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        for _ in range(self.copies):
            if self.cancel.is_set():
                break
            self._draw_images([img], os.path.basename(path))
        return True

    def _draw_images(self, images, docname="Document"):
        import win32ui
        import win32con
        from PIL import ImageWin
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(self.printer)
        try:
            horz = hDC.GetDeviceCaps(win32con.HORZRES)
            vert = hDC.GetDeviceCaps(win32con.VERTRES)
            hDC.StartDoc(docname)
            try:
                for img in images:
                    if self.cancel.is_set():
                        break
                    hDC.StartPage()
                    iw, ih = img.size
                    scale = min(horz / float(iw), vert / float(ih))
                    dw, dh = int(iw * scale), int(ih * scale)
                    x = (horz - dw) // 2
                    y = (vert - dh) // 2
                    ImageWin.Dib(img).draw(hDC.GetHandleOutput(), (x, y, x + dw, y + dh))
                    hDC.EndPage()
            finally:
                hDC.EndDoc()
        finally:
            hDC.DeleteDC()

    def _print_word(self, path):
        import win32com.client as win32
        app = None
        try:
            app = win32.DispatchEx("Word.Application")
            app.Visible = False
            try: app.DisplayAlerts = 0
            except Exception: pass
            try: app.AutomationSecurity = self.MSO_SECURITY_FORCE_DISABLE
            except Exception: pass
            doc = app.Documents.Open(os.path.abspath(path), ReadOnly=True,
                                     ConfirmConversions=False, AddToRecentFiles=False)
            try:
                try: app.ActivePrinter = self.printer
                except Exception: pass
                doc.PrintOut(Background=False, Copies=self.copies)
                time.sleep(1.0)
            finally:
                doc.Close(SaveChanges=False)
            return True
        finally:
            if app is not None:
                try: app.Quit()
                except Exception: pass

    def _print_excel(self, path):
        import win32com.client as win32
        app = None
        try:
            app = win32.DispatchEx("Excel.Application")
            app.Visible = False
            try: app.DisplayAlerts = False
            except Exception: pass
            try: app.AutomationSecurity = self.MSO_SECURITY_FORCE_DISABLE
            except Exception: pass
            wb = app.Workbooks.Open(os.path.abspath(path), ReadOnly=True, UpdateLinks=0)
            try:
                try:
                    wb.PrintOut(Copies=self.copies, ActivePrinter=self.printer)
                except Exception:
                    wb.PrintOut(Copies=self.copies)
                time.sleep(1.0)
            finally:
                wb.Close(SaveChanges=False)
            return True
        finally:
            if app is not None:
                try: app.Quit()
                except Exception: pass

    def _print_ppt(self, path):
        import win32com.client as win32
        app = None
        try:
            app = win32.DispatchEx("PowerPoint.Application")
            try: app.AutomationSecurity = self.MSO_SECURITY_FORCE_DISABLE
            except Exception: pass
            try: app.Visible = True
            except Exception: pass
            pres = app.Presentations.Open(os.path.abspath(path), WithWindow=False, ReadOnly=True)
            try:
                try: pres.PrintOptions.ActivePrinter = self.printer
                except Exception: pass
                try: pres.PrintOptions.NumberOfCopies = self.copies
                except Exception: pass
                pres.PrintOut()
                time.sleep(1.5)
            finally:
                pres.Close()
            return True
        finally:
            if app is not None:
                try: app.Quit()
                except Exception: pass


# ===================== الجسر بين الواجهة والخلفية =====================
_window = None
_cancel = threading.Event()


def _js(code):
    try:
        if _window is not None:
            _window.evaluate_js(code)
    except Exception:
        pass


def _q(s):
    return json.dumps("" if s is None else str(s), ensure_ascii=False)


class Api:
    def list_printers(self):
        printers, default = [], ""
        if HAS_WIN32:
            try:
                flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
                printers = [p[2] for p in win32print.EnumPrinters(flags)]
            except Exception:
                printers = []
            try:
                default = win32print.GetDefaultPrinter()
            except Exception:
                default = printers[0] if printers else ""
        return {"printers": printers, "default": default}

    def _meta(self, path):
        try:
            sz = os.path.getsize(path)
        except Exception:
            sz = 0
        return {"name": os.path.basename(path), "path": path, "size": human(sz)}

    def pick_files(self):
        types = ("المستندات المدعومة (*.pdf;*.doc;*.docx;*.rtf;*.txt;*.xls;*.xlsx;*.csv;*.ppt;*.pptx;*.jpg;*.jpeg;*.png;*.bmp;*.tif;*.tiff;*.gif)",
                 "كل الملفات (*.*)")
        res = _window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=True, file_types=types)
        out = []
        for p in (res or []):
            if ext_category(p):
                out.append(self._meta(p))
        return out

    def pick_folder(self):
        res = _window.create_file_dialog(webview.FOLDER_DIALOG)
        out = []
        if res:
            folder = res[0]
            for root, _, fs in os.walk(folder):
                for f in sorted(fs):
                    fp = os.path.join(root, f)
                    if ext_category(fp):
                        out.append(self._meta(fp))
        return out

    def start_print(self, payload):
        _cancel.clear()
        threading.Thread(target=self._run, args=(payload or {},), daemon=True).start()
        return True

    def cancel(self):
        _cancel.set()
        return True

    def _run(self, payload):
        printer = payload.get("printer") or (win32print.GetDefaultPrinter() if HAS_WIN32 else "")
        try:
            copies = max(1, int(payload.get("copies") or 1))
        except Exception:
            copies = 1
        rng = payload.get("range") or ""
        paths = payload.get("files") or []
        total = len(paths)
        ok = fail = 0
        try:
            with PrintEngine(printer, copies, rng, 300,
                             log=lambda m: _js("pyLog(%s)" % _q(m)),
                             cancel_event=_cancel) as eng:
                for i, p in enumerate(paths):
                    if _cancel.is_set():
                        _js("pyLog(%s,'warn')" % _q("تم الإيقاف بواسطة المستخدم."))
                        break
                    _js("pyRow(%d,%s)" % (i, _q("قيد الطباعة")))
                    _js("pyLog(%s)" % _q("[%d/%d] %s" % (i + 1, total, os.path.basename(p))))
                    try:
                        eng.print_file(p)
                        ok += 1
                        _js("pyRow(%d,%s)" % (i, _q("تمت")))
                    except Exception as e:
                        fail += 1
                        _js("pyRow(%d,%s)" % (i, _q("فشل")))
                        _js("pyLog(%s,'warn')" % _q("خطأ: %s" % e))
                    _js("pyProgress(%d)" % round((i + 1) * 100.0 / max(1, total)))
        except Exception as e:
            _js("pyLog(%s,'warn')" % _q("خطأ عام: %s" % e))
        _js("pyDone(%d,%d)" % (ok, fail))


def main():
    global _window
    _window = webview.create_window(
        "الطابعة الجماعية الآمنة — Secure Batch Printer",
        url=resource_path("ui.html"),
        js_api=Api(),
        width=1060, height=780, min_size=(880, 620),
    )
    webview.start()


if __name__ == "__main__":
    main()
