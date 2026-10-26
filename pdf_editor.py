"""Локальный редактор текста в PDF.

Запуск:  python pdf_editor.py [файл.pdf]
Откроется браузер с редактором. Всё работает на этом компьютере — файл никуда не отправляется.

Как устроена замена текста: старые символы физически удаляются из PDF (через redaction
PyMuPDF, картинки и линии не трогаются), а новый текст пишется на ту же базовую линию.
Шрифт по возможности берётся тот, что встроен в документ; если в нём нет нужных букв
(встроенные шрифты обычно урезаны до использованных символов) — похожий системный.
"""
import hashlib
import json
import os
import re
import sys
import threading
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

try:
    import pymupdf
except ImportError:
    sys.exit("Нужна библиотека PyMuPDF. Установите её командой:  python -m pip install pymupdf")

HERE = os.path.dirname(os.path.abspath(__file__))
UNDO_LIMIT = 40

FONT_DIRS = [
    os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
    os.path.expanduser("~/Library/Fonts"),
    "/Library/Fonts",
    "/System/Library/Fonts/Supplemental",
    "/usr/share/fonts",
]

# семейство -> файлы (обычный, жирный, курсив, жирный курсив)
FAMILIES = {
    "Arial": ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    "Times New Roman": ("times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf"),
    "Calibri": ("calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf"),
    "Arial Narrow": ("arialn.ttf", "arialnb.ttf", "arialni.ttf", "arialnbi.ttf"),
    "Courier New": ("cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"),
    "Verdana": ("verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf"),
    "Tahoma": ("tahoma.ttf", "tahomabd.ttf", "tahoma.ttf", "tahomabd.ttf"),
    "Georgia": ("georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf"),
    "Segoe UI": ("segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "segoeuiz.ttf"),
    "Liberation Sans": ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf",
                        "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
    "Liberation Serif": ("LiberationSerif-Regular.ttf", "LiberationSerif-Bold.ttf",
                         "LiberationSerif-Italic.ttf", "LiberationSerif-BoldItalic.ttf"),
    "DejaVu Sans": ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf",
                    "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
}

# фрагмент имени шрифта в PDF -> семейство (порядок важен: narrow раньше arial)
GUESSES = [
    ("arialnarrow", "Arial Narrow"), ("timesnewroman", "Times New Roman"), ("times", "Times New Roman"),
    ("tinos", "Times New Roman"), ("liberationserif", "Liberation Serif"), ("nimbusrom", "Times New Roman"),
    ("calibri", "Calibri"), ("carlito", "Calibri"), ("courier", "Courier New"), ("cousine", "Courier New"),
    ("liberationmono", "Courier New"), ("verdana", "Verdana"), ("tahoma", "Tahoma"), ("georgia", "Georgia"),
    ("segoe", "Segoe UI"), ("arial", "Arial"), ("helvetica", "Arial"), ("arimo", "Arial"),
    ("liberationsans", "Liberation Sans"), ("nimbussans", "Arial"), ("dejavusans", "DejaVu Sans"),
]
SERIF_FALLBACKS = ("Times New Roman", "Liberation Serif")
SANS_FALLBACKS = ("Arial", "Liberation Sans", "DejaVu Sans")

_font_index = None


def font_index():
    """Имя файла шрифта в нижнем регистре -> полный путь."""
    global _font_index
    if _font_index is None:
        _font_index = {}
        for d in FONT_DIRS:
            if d and os.path.isdir(d):
                for root, _, files in os.walk(d):
                    for f in files:
                        _font_index.setdefault(f.lower(), os.path.join(root, f))
    return _font_index


def available_families():
    return [fam for fam, files in FAMILIES.items() if files[0].lower() in font_index()]


def family_file(family, bold, italic):
    files = FAMILIES.get(family) or FAMILIES["Arial"]
    for name in (files[int(bold) + 2 * int(italic)], files[int(bold)], files[0]):
        path = font_index().get(name.lower())
        if path:
            return path
    return None


def clean_font_name(name):
    return re.sub(r"^[A-Z]{6}\+", "", name or "")


def font_key(name):
    """'ABCDEF+Times New Roman Regular' и 'TimesNewRomanPSMT' -> 'timesnewroman'."""
    key = re.sub(r"[^a-z0-9]", "", clean_font_name(name).lower())
    return re.sub(r"(regular|psmt|mt)$", "", key)


def guess_family(font, flags):
    avail = available_families()
    key = re.sub(r"[^a-z]", "", clean_font_name(font).lower())
    for part, fam in GUESSES:
        if part in key and fam in avail:
            return fam
    if flags & 8 and "Courier New" in avail:
        return "Courier New"
    for fam in (SERIF_FALLBACKS if flags & 4 or "serif" in key and "sans" not in key else SANS_FALLBACKS):
        if fam in avail:
            return fam
    return avail[0] if avail else "Arial"


def guess_style(font, flags):
    n = clean_font_name(font).lower()
    bold = bool(flags & 16) or any(w in n for w in ("bold", "black", "heavy", "semibold", "demi"))
    italic = bool(flags & 2) or "italic" in n or "oblique" in n
    return bold, italic


def hex_to_rgb(h):
    h = (h or "#000000").lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


class Editor:
    def __init__(self):
        self.lock = threading.RLock()
        self.doc = None
        self.filename = ""
        self.password = None
        self.undo = []
        self.dirty = False
        self.embedded = {}  # xref -> (buffer, Font, md5) | None
        self.ink = {}  # (md5 шрифта, буква) -> есть ли контур

    # ---------- документ ----------

    def open(self, data, filename, password=None):
        doc = pymupdf.open(stream=data, filetype="pdf")
        if doc.needs_pass:
            if not password or not doc.authenticate(password):
                raise PasswordNeeded()
        if not doc.is_pdf:
            raise ValueError("Это не PDF")
        self.doc, self.filename, self.password = doc, filename, password
        self.undo, self.dirty, self.embedded = [], False, {}

    def state(self):
        d = self.doc
        return {
            "loaded": d is not None,
            "filename": self.filename,
            "pages": [[p.rect.width, p.rect.height] for p in d] if d else [],
            "fonts": available_families(),
            "undo": len(self.undo),
            "dirty": self.dirty,
        }

    def checkpoint(self, page_no):
        self.undo.append((self.doc.tobytes(), page_no))
        del self.undo[:-UNDO_LIMIT]
        self.dirty = True

    def do_undo(self):
        if not self.undo:
            return None
        data, page_no = self.undo.pop()
        self.doc, self.embedded = self.reopen(data), {}
        self.dirty = True
        return page_no

    def reopen(self, data):
        doc = pymupdf.open(stream=data, filetype="pdf")
        if doc.needs_pass:
            doc.authenticate(self.password or "")
        return doc

    def compact(self):
        """garbage=4 склеивает одинаковые объекты — копии шрифтов, которые приносит каждый перенос."""
        self.doc, self.embedded = self.reopen(self.doc.tobytes(garbage=4)), {}

    def save_bytes(self):
        tmp = self.reopen(self.doc.tobytes(garbage=4))
        try:
            tmp.subset_fonts()  # системные шрифты встраиваются целиком — урезаем до нужных букв
        except Exception:
            pass
        data = tmp.tobytes(garbage=4, deflate=True)
        self.dirty = False
        return data

    # ---------- текст ----------

    def spans(self, n):
        return self.spans_raw(self.doc[n])

    @staticmethod
    def spans_raw(page):
        """Фрагменты текста. bbox/origin — в координатах PDF (для правки), box — как на экране."""
        rot = page.rotation_matrix
        out = []
        for b in page.get_text("dict", flags=pymupdf.TEXTFLAGS_TEXT)["blocks"]:
            for line in b.get("lines", []):
                dx, dy = line["dir"]
                horizontal = abs(dy) < 1e-3 and dx > 0
                for s in line["spans"]:
                    if not s["text"].strip():
                        continue
                    bold, italic = guess_style(s["font"], s["flags"])
                    out.append({
                        "bbox": [round(v, 2) for v in s["bbox"]],
                        "box": [round(v, 2) for v in pymupdf.Rect(s["bbox"]) * rot],
                        "origin": [round(v, 2) for v in s["origin"]],
                        "base": [round(v, 2) for v in pymupdf.Point(s["origin"]) * rot],  # начало базовой линии на экране
                        "text": s["text"],
                        "font": s["font"],
                        "flags": s["flags"],
                        "size": round(s["size"], 2),
                        "color": "#%06x" % s["color"],
                        "family": guess_family(s["font"], s["flags"]),
                        "bold": bold,
                        "italic": italic,
                        "horizontal": horizontal,
                    })
        return out

    def embedded_fonts(self, page, span_font):
        """Встроенные в PDF шрифты с именем как у фрагмента: [(buffer, Font, md5)]."""
        want = font_key(span_font)
        found = []
        for xref, ext, ftype, basefont, refname, *_ in page.get_fonts():
            if (ext == "n/a" or ftype == "Type3" or refname.startswith(("Ed", "EmbX"))
                    or font_key(basefont) != want):
                continue  # не встроен / Type3 / вставлен самим редактором
            if xref not in self.embedded:
                try:
                    buf = self.doc.extract_font(xref)[-1]
                    self.embedded[xref] = ((buf, pymupdf.Font(fontbuffer=buf), hashlib.md5(buf).hexdigest())
                                           if buf else None)
                except Exception:
                    self.embedded[xref] = None
            if self.embedded[xref]:
                found.append(self.embedded[xref])
        return found

    @staticmethod
    def drawn_glyphs(page, span):
        """{символ: номер глифа} — как исходный фрагмент реально нарисован в PDF."""
        r = pymupdf.Rect(span["bbox"]) + (-1, -1, 1, 1)
        key = font_key(span["font"])
        out = {}
        for sp in page.get_texttrace():
            if font_key(sp["font"]) == key:
                for u, g, origin, *_ in sp["chars"]:
                    if r.contains(pymupdf.Point(origin)):
                        out[chr(u)] = g
        return out

    def has_ink(self, digest, buf, ch):
        """Есть ли у буквы контур: урезанные шрифты нередко хранят пустые глифы для неиспользованных букв."""
        key = (digest, ch)
        if key not in self.ink:
            d = pymupdf.open()
            p = d.new_page(width=48, height=48)
            p.insert_font(fontname="F", fontbuffer=buf)
            p.insert_text((6, 34), ch, fontname="F", fontsize=28)
            self.ink[key] = min(p.get_pixmap(colorspace=pymupdf.csGRAY).samples) < 160
        return self.ink[key]

    def font_fits(self, buf, font, digest, text, drawn):
        """Шрифт документа годится, только если его таблица символов совпадает с тем, как фрагмент
        нарисован (иначе вместо букв выйдут чужие значки), и у всех новых букв есть контуры."""
        if not drawn or any(font.has_glyph(ord(c)) != g for c, g in drawn.items()):
            return False
        for ch in set(text) - {"\n"}:
            if not font.has_glyph(ord(ch)):
                return False
            if not ch.isspace() and ch not in drawn and not self.has_ink(digest, buf, ch):
                return False
        return True

    def choose_font(self, page, text, family, bold, italic, span=None):
        """-> (имя ресурса, kwargs для insert_font, Font, подпись для пользователя)."""
        if family == "auto" and span is not None:
            if (bold, italic) == guess_style(span["font"], span["flags"]):
                drawn = self.drawn_glyphs(page, span)
                for buf, font, digest in self.embedded_fonts(page, span["font"]):
                    if self.font_fits(buf, font, digest, text, drawn):
                        # имя — по содержимому шрифта: номера объектов меняются при пересборке файла,
                        # и имя по номеру могло совпасть с другим шрифтом на странице
                        return "EmbX" + digest[:10], {"fontbuffer": buf}, font, "шрифт документа"
            family = guess_family(span["font"], span["flags"])
        path = family_file(family, bold, italic)
        if not path:
            return "helv", {}, pymupdf.Font("helv"), "Helvetica (без кириллицы!)"
        name = "Ed" + hashlib.md5(path.lower().encode()).hexdigest()[:8]
        label = family + (" жирный" if bold else "") + (" курсив" if italic else "")
        return name, {"fontfile": path}, pymupdf.Font(fontfile=path), label

    def write_lines(self, page, text, x_left, x_right, baseline, size, color, fontinfo, align="left", fit=False):
        name, kw, font, _ = fontinfo
        if kw:
            page.insert_font(fontname=name, **kw)
        lines = text.split("\n")
        widths = [font.text_length(line, fontsize=size) for line in lines]
        box_w = x_right - x_left
        scale = 1.0
        if fit and box_w > 0 and max(widths) > box_w:
            scale = box_w / max(widths)
        for i, (line, w) in enumerate(zip(lines, widths)):
            if not line:
                continue
            w *= scale
            x = {"center": (x_left + x_right - w) / 2, "right": x_right - w}.get(align, x_left)
            p = pymupdf.Point(x, baseline + i * size * 1.2)
            morph = (p, pymupdf.Matrix(scale, 1)) if scale < 0.999 else None
            page.insert_text(p, line, fontname=name, fontsize=size, color=color, morph=morph)
        return (max(widths) * scale / box_w - 1) if box_w > 0 else 0

    @staticmethod
    def text_band(s):
        """Узкая полоса по середине строки фрагмента: удаляет его символы, не задевая соседние строки и фрагменты."""
        x0, y0, x1, y1 = s["bbox"]
        oy, size = s["origin"][1], s["size"]
        bbox = pymupdf.Rect(s["bbox"])
        inset = min(1.0, size * 0.08)
        band = pymupdf.Rect(x0 + inset, oy - size * 0.55, x1 - inset, oy - size * 0.2) & bbox
        if band.is_empty:
            band = bbox + (inset, (y1 - y0) * 0.3, -inset, -(y1 - y0) * 0.3)
        return band

    def remove_text(self, page, spans):
        for s in spans:
            page.add_redact_annot(self.text_band(s), fill=False, cross_out=False)
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                              graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)

    def edit(self, req):
        n, s = req["page"], req["span"]
        page = self.doc[n]
        text = req.get("text", "").replace("\r", "")
        size = float(req.get("size") or s["size"])
        bold, italic = bool(req.get("bold")), bool(req.get("italic"))
        self.checkpoint(n)
        # шрифт выбираем до удаления: проверка сверяет его с тем, как нарисован исходный фрагмент
        fontinfo = self.choose_font(page, text, req.get("font", "auto"), bold, italic, s) if text.strip() else None
        self.remove_text(page, [s])
        x1 = s["bbox"][2]
        ox, oy = s["origin"]

        note = ""
        if fontinfo:
            overflow = self.write_lines(page, text, ox, x1, oy, size, hex_to_rgb(req.get("color")),
                                        fontinfo, req.get("align", "left"), bool(req.get("fit")))
            note = "Шрифт: " + fontinfo[3]
            if overflow > 0.02:
                note += (f". Новый текст на {overflow:.0%} шире старого — если наезжает на соседний, "
                         "включите «Вписать в ширину»")
        return {"page": n, "note": note}

    def add(self, req):
        n = req["page"]
        page = self.doc[n]
        text = req.get("text", "").replace("\r", "")
        if not text.strip():
            return {"page": n, "note": ""}
        size = float(req.get("size") or 11)
        self.checkpoint(n)
        fontinfo = self.choose_font(page, text, req.get("font") or "Arial", bool(req.get("bold")),
                                    bool(req.get("italic")))
        name, kw, font, label = fontinfo
        if kw:
            page.insert_font(fontname=name, **kw)
        x, y = req["x"], req["y"] + font.ascender * size  # клик — верх строки, нужна базовая линия
        for i, line in enumerate(text.split("\n")):
            if line:
                # x, y — экранные координаты; на повёрнутой странице переводим их и поворачиваем текст
                p = pymupdf.Point(x, y + i * size * 1.2) * page.derotation_matrix
                page.insert_text(p, line, fontname=name, fontsize=size, color=hex_to_rgb(req.get("color")),
                                 rotate=page.rotation)
        return {"page": n, "note": "Шрифт: " + label}

    def delete(self, req):
        n, spans = req["page"], req["spans"]
        if spans:
            self.checkpoint(n)
            self.remove_text(self.doc[n], spans)
        return {"page": n, "note": ""}

    def move(self, req):
        """Переносит фрагменты, сохраняя исходные символы и шрифт.

        Запрос: {page, spans, dx, dy} — все на один сдвиг, или {page, items: [{spans, dx, dy}]} —
        у каждой группы свой (так работает выравнивание). Сдвиг — в экранных pt.
        Для каждой группы копия страницы очищается от картинок, линий и всего прочего текста,
        и её кусок вставляется на новое место как вложенная PDF-страница; исходные символы удаляются.
        """
        n = req["page"]
        page = self.doc[n]
        m = page.derotation_matrix
        groups = []
        for it in req.get("items") or [{"spans": req.get("spans"), "dx": req["dx"], "dy": req["dy"]}]:
            d = pymupdf.Point(it["dx"], it["dy"]) * m - pymupdf.Point(0, 0) * m
            if it["spans"] and (abs(d.x) >= 0.01 or abs(d.y) >= 0.01):
                groups.append((it["spans"], d))
        if not groups:
            return {"page": n, "note": ""}
        self.checkpoint(n)

        base = pymupdf.open()
        base.insert_pdf(self.doc, from_page=n, to_page=n)
        bp = base[0]
        bp.set_rotation(0)
        bp.add_redact_annot(bp.rect, fill=False, cross_out=False)
        bp.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_REMOVE,
                            graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED,
                            text=pymupdf.PDF_REDACT_TEXT_NONE)
        base_bytes, all_text = base.tobytes(), self.spans_raw(bp)
        same = lambda a, b: max(abs(p - q) for p, q in zip(a, b)) < 0.6

        pieces = []
        for spans, d in groups:
            keep = [s["bbox"] for s in spans]
            clip = pymupdf.Rect(keep[0])
            for r in keep[1:]:
                clip |= r
            clip += (-1, -1, 1, 1)
            tmp = pymupdf.open(stream=base_bytes, filetype="pdf")
            # убираем весь остальной текст, а не только попавший в кусок: иначе за пределами
            # видимой области остались бы скрытые копии, которые находит поиск в других программах
            others = [s for s in all_text if not any(same(s["bbox"], k) for k in keep)]
            if others:
                self.remove_text(tmp[0], others)
            pieces.append((tmp, clip, d))

        rot = page.rotation
        page.set_rotation(0)
        try:
            self.remove_text(page, [s for spans, _ in groups for s in spans])
            for tmp, clip, d in pieces:
                page.show_pdf_page(clip + (d.x, d.y, d.x, d.y), tmp, 0, clip=clip)
        finally:
            page.set_rotation(rot)
        self.compact()
        return {"page": n, "note": ""}

    def erase(self, req):
        n = req["page"]
        page = self.doc[n]
        rect = (pymupdf.Rect(req["rect"]) & page.rect) * page.derotation_matrix
        if rect.is_empty:
            return {"page": n, "note": ""}
        self.checkpoint(n)
        page.add_redact_annot(rect, fill=hex_to_rgb(req.get("color", "#ffffff")), cross_out=False)
        page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_PIXELS,
                              graphics=pymupdf.PDF_REDACT_LINE_ART_NONE)
        return {"page": n, "note": ""}

    def render(self, n, zoom):
        zoom = max(0.2, min(float(zoom), 5.0))
        return self.doc[n].get_pixmap(matrix=pymupdf.Matrix(zoom, zoom), annots=True).tobytes("png")


class PasswordNeeded(Exception):
    pass


ED = Editor()


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def send(self, code, body, ctype, headers=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def json(self, obj, code=200):
        self.send(code, json.dumps(obj, ensure_ascii=False).encode(), "application/json; charset=utf-8")

    def do_GET(self):
        url = urllib.parse.urlparse(self.path)
        q = {k: v[0] for k, v in urllib.parse.parse_qs(url.query).items()}
        try:
            with ED.lock:
                if url.path == "/":
                    with open(os.path.join(HERE, "index.html"), "rb") as f:
                        self.send(200, f.read(), "text/html; charset=utf-8")
                elif url.path == "/api/state":
                    self.json(ED.state())
                elif url.path == "/api/page":
                    self.send(200, ED.render(int(q["n"]), q.get("z", 2)), "image/png")
                elif url.path == "/api/spans":
                    self.json({"spans": ED.spans(int(q["n"]))})
                elif url.path == "/api/save":
                    base = os.path.splitext(ED.filename or "document")[0]
                    name = urllib.parse.quote(base + "_edited.pdf")
                    self.send(200, ED.save_bytes(), "application/pdf",
                              {"Content-Disposition": f"attachment; filename=\"edited.pdf\"; filename*=UTF-8''{name}"})
                else:
                    self.json({"error": "not found"}, 404)
        except Exception as e:
            self.json({"error": f"{type(e).__name__}: {e}"}, 500)

    def do_POST(self):
        url = urllib.parse.urlparse(self.path)
        body = self.rfile.read(int(self.headers.get("Content-Length") or 0))
        try:
            with ED.lock:
                if url.path == "/api/open":
                    name = urllib.parse.unquote(self.headers.get("X-Filename") or "document.pdf")
                    pwd = urllib.parse.unquote(self.headers.get("X-Password") or "") or None
                    ED.open(body, name, pwd)
                    self.json(ED.state())
                    return
                if ED.doc is None:
                    raise ValueError("Сначала откройте PDF")
                if url.path == "/api/undo":
                    result = {"page": ED.do_undo()}
                else:
                    req = json.loads(body or b"{}")
                    action = {"/api/edit": ED.edit, "/api/add": ED.add, "/api/erase": ED.erase,
                              "/api/move": ED.move, "/api/delete": ED.delete}.get(url.path)
                    if not action:
                        self.json({"error": "not found"}, 404)
                        return
                    result = action(req)
                result.update(undo=len(ED.undo), dirty=ED.dirty)
                self.json(result)
        except PasswordNeeded:
            self.json({"error": "password"}, 401)
        except Exception as e:
            self.json({"error": f"{type(e).__name__}: {e}"}, 500)


def main():
    paths = [a for a in sys.argv[1:] if not a.startswith("--")]
    if paths:
        path = paths[0]
        with open(path, "rb") as f:
            try:
                ED.open(f.read(), os.path.basename(path))
            except PasswordNeeded:
                print("PDF защищён паролем — откройте его через кнопку в редакторе.")
    server = None
    for port in (8765, 8766, 8767, 0):
        try:
            server = ThreadingHTTPServer(("127.0.0.1", port), Handler)
            break
        except OSError:
            continue
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"PDF-редактор запущен: {url}\nНе закрывайте это окно, пока работаете. Ctrl+C — выход.")
    if "--no-browser" not in sys.argv:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
