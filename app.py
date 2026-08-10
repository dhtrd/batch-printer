# -*- coding: utf-8 -*-
"""
الطابعة الجماعية الآمنة — تطبيق سطح المكتب (واجهة أوركيد)
واجهة HTML داخل نافذة سطح مكتب (pywebview) + خلفية Python للطباعة الفعلية.

المزايا:
  * الطباعة كـ«مهمة واحدة» مدمجة: كل الملفات تُطبع ضمن مهمة طباعة واحدة،
    فيمكن إيقافها فورًا دون الحاجة لإلغاء كل ملف على حدة من طابور ويندوز.
    (ملفات أوفيس تُحوّل داخليًا إلى PDF ثم تُدمج مع البقية في نفس المهمة.)
  * سحب وإفلات الملفات على النافذة لإضافتها مباشرة.
  * آمن: يعمل محليًا، بلا شبكة، ويعطّل ماكرو أوفيس.

النظام: Windows 10/11. الترخيص: مجاني ومفتوح المصدر.
"""

import os
import sys
import json
import base64
import tempfile
import threading

import webview

try:
    import win32print
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

APP_VERSION = "1.1.0"

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


def _draw_page(hDC, img):
    import win32con
    from PIL import ImageWin
    horz = hDC.GetDeviceCaps(win32con.HORZRES)
    vert = hDC.GetDeviceCaps(win32con.VERTRES)
    hDC.StartPage()
    iw, ih = img.size
    scale = min(horz / float(iw), vert / float(ih))
    dw, dh = int(iw * scale), int(ih * scale)
    x = (horz - dw) // 2
    y = (vert - dh) // 2
    ImageWin.Dib(img).draw(hDC.GetHandleOutput(), (x, y, x + dw, y + dh))
    hDC.EndPage()


class PrintEngine:
    MSO = 3  # msoAutomationSecurityForceDisable

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
            except Exception:
                pass
        return self

    def __exit__(self, *exc):
        if HAS_WIN32 and self._orig_default:
            try:
                win32print.SetDefaultPrinter(self._orig_default)
            except Exception:
                pass
        return False

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

    # تحويل ملف أوفيس إلى PDF مؤقت (لدمجه في المهمة الواحدة)
    def office_to_pdf(self, path):
        import win32com.client as win32
        cat = ext_category(path)
        fd, tmp = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        ap = os.path.abspath(path)
        out = os.path.abspath(tmp)
        if cat == "Word":
            app = win32.DispatchEx("Word.Application")
            try:
                app.Visible = False
                try: app.AutomationSecurity = self.MSO
                except Exception: pass
                doc = app.Documents.Open(ap, ReadOnly=True, ConfirmConversions=False, AddToRecentFiles=False)
                doc.ExportAsFixedFormat(out, 17)  # wdExportFormatPDF
                doc.Close(False)
            finally:
                try: app.Quit()
                except Exception: pass
        elif cat == "Excel":
            app = win32.DispatchEx("Excel.Application")
            try:
                app.Visible = False
                try: app.AutomationSecurity = self.MSO
                except Exception: pass
                wb = app.Workbooks.Open(ap, ReadOnly=True, UpdateLinks=0)
                wb.ExportAsFixedFormat(0, out)  # xlTypePDF
                wb.Close(False)
            finally:
                try: app.Quit()
                except Exception: pass
        elif cat == "PowerPoint":
            app = win32.DispatchEx("PowerPoint.Application")
            try:
                try: app.AutomationSecurity = self.MSO
                except Exception: pass
                try: app.Visible = True
                except Exception: pass
                pres = app.Presentations.Open(ap, WithWindow=False, ReadOnly=True)
                pres.SaveAs(out, 32)  # ppSaveAsPDF
                pres.Close()
            finally:
                try: app.Quit()
                except Exception: pass
        else:
            raise ValueError("ليست صيغة أوفيس")
        return out

    def pages_for(self, path):
        from PIL import Image
        if ext_category(path) == "Image":
            img = Image.open(path)
            if img.mode not in ("RGB", "L"):
                img = img.convert("RGB")
            return [img]
        import fitz
        doc = fitz.open(path)
        try:
            rng = self._parse_pages(self.page_range, doc.page_count)
            zoom = self.dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            out = []
            for i in rng:
                if self.cancel.is_set():
                    break
                pg = doc.load_page(i)
                pix = pg.get_pixmap(matrix=mat, alpha=False)
                out.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
            return out
        finally:
            doc.close()

    # الطباعة كمهمة طباعة واحدة لكل الملفات (قابلة للإيقاف الفوري)
    def print_batch(self, paths, on_start, on_done, on_log, on_progress):
        import win32ui
        prepared = []   # (orig_index, render_path, name)
        temps = []
        total = len(paths)
        for i, p in enumerate(paths):
            if self.cancel.is_set():
                break
            cat = ext_category(p)
            if cat in ("Word", "Excel", "PowerPoint"):
                on_log("تحويل: %s" % os.path.basename(p))
                try:
                    pdf = self.office_to_pdf(p)
                    temps.append(pdf)
                    prepared.append((i, pdf, os.path.basename(p)))
                except Exception as e:
                    on_done(i, False)
                    on_log("تعذّر تحويل %s: %s" % (os.path.basename(p), e))
            elif cat in ("PDF", "Image"):
                prepared.append((i, p, os.path.basename(p)))
            else:
                on_done(i, False)

        if not prepared:
            return (0, total)

        ok = fail = 0
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(self.printer)
        started = False
        try:
            hDC.StartDoc("Secure Batch Print")
            started = True
            for c in range(self.copies):
                if self.cancel.is_set():
                    break
                for (i, src, name) in prepared:
                    if self.cancel.is_set():
                        break
                    if c == 0:
                        on_start(i)
                    try:
                        imgs = self.pages_for(src)
                        for img in imgs:
                            if self.cancel.is_set():
                                break
                            _draw_page(hDC, img)
                        if c == 0:
                            on_done(i, True)
                            ok += 1
                            on_progress(round((ok + fail) * 100.0 / max(1, len(prepared))))
                    except Exception as e:
                        if c == 0:
                            on_done(i, False)
                            fail += 1
                            on_log("خطأ في %s: %s" % (name, e))
        finally:
            try:
                if started:
                    hDC.EndDoc()
            except Exception:
                pass
            try:
                hDC.DeleteDC()
            except Exception:
                pass
            for f in temps:
                try:
                    os.remove(f)
                except Exception:
                    pass
        return (ok, fail)


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
        return [self._meta(p) for p in (res or []) if ext_category(p)]

    def pick_folder(self):
        res = _window.create_file_dialog(webview.FOLDER_DIALOG)
        out = []
        if res:
            for root, _, fs in os.walk(res[0]):
                for f in sorted(fs):
                    fp = os.path.join(root, f)
                    if ext_category(fp):
                        out.append(self._meta(fp))
        return out

    def add_paths(self, paths):
        out = []
        for p in (paths or []):
            try:
                if p and os.path.isfile(p) and ext_category(p):
                    out.append(self._meta(p))
            except Exception:
                pass
        return out

    def add_dropped(self, name, data_url):
        try:
            if not ext_category("x" + os.path.splitext(name)[1]):
                return None
            raw = base64.b64decode((data_url or "").split(",")[-1])
            d = os.path.join(tempfile.gettempdir(), "sbp_drop")
            os.makedirs(d, exist_ok=True)
            fp = os.path.join(d, os.path.basename(name))
            with open(fp, "wb") as f:
                f.write(raw)
            return self._meta(fp)
        except Exception:
            return None

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

        def on_start(i):
            _js("pyRow(%d,%s)" % (i, _q("قيد الطباعة")))

        def on_done(i, okv):
            _js("pyRow(%d,%s)" % (i, _q("تمت" if okv else "فشل")))

        def on_log(m):
            _js("pyLog(%s)" % _q(m))

        def on_progress(p):
            _js("pyProgress(%d)" % p)

        on_log("بدء الطباعة كمهمة واحدة على: %s — %d ملف%s"
               % (printer, total, (" (نسخ: %d)" % copies) if copies > 1 else ""))
        ok = fail = 0
        try:
            with PrintEngine(printer, copies, rng, 300, log=on_log, cancel_event=_cancel) as eng:
                ok, fail = eng.print_batch(paths, on_start, on_done, on_log, on_progress)
        except Exception as e:
            on_log("خطأ عام: %s" % e)
        if _cancel.is_set():
            _js("pyLog(%s,'warn')" % _q("تم الإيقاف — أُلغيت بقية المهمة."))
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
