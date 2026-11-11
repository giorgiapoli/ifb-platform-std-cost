"""PDF di una tabella colorata (Styler di pandas): stessi testi e stessi colori che si vedono a schermo."""
import io
import os
import re

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

_FONT = {"ok": False}


def _font():
    """Font con Δ, €, ³ (Arial su Windows, altrimenti DejaVu); se non si trova, Helvetica."""
    if _FONT["ok"]:
        return _FONT["n"], _FONT["b"]
    cand = [("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
            ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")]
    n, b = "Helvetica", "Helvetica-Bold"
    for fn, fb in cand:
        if os.path.exists(fn) and os.path.exists(fb):
            pdfmetrics.registerFont(TTFont("StdN", fn))
            pdfmetrics.registerFont(TTFont("StdB", fb))
            n, b = "StdN", "StdB"
            break
    _FONT.update(ok=True, n=n, b=b)
    return n, b


def _css(coppie):
    """[(prop, val), ...] -> (sfondo, testo, grassetto); le regole piu' in basso valgono di piu'."""
    bg = fg = None
    bold = False
    for p, v in coppie:
        p, v = str(p).strip().lower(), str(v).strip()
        if p == "background-color" and v:
            bg = v
        elif p == "color" and v:
            fg = v
        elif p == "font-weight" and v:
            bold = v in ("600", "700", "bold")
    return bg, fg, bold


def _esc(t):
    return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def tabella_pdf(titolo, righe_intro, tab, sty, note=None, a_sinistra=(), larghe=None):
    """tab = DataFrame mostrato; sty = il suo Styler (formati e colori); righe_intro = testi sotto il titolo; note = legenda."""
    nome, grass = _font()
    sty._compute()
    ctx = sty.ctx
    cols = list(tab.columns)
    larghe = larghe or {}
    pag = landscape(A4)
    marg = 12
    doc = SimpleDocTemplate(io.BytesIO(), pagesize=pag, leftMargin=marg, rightMargin=marg, topMargin=marg, bottomMargin=marg)
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=pag, leftMargin=marg, rightMargin=marg, topMargin=marg, bottomMargin=marg,
                            title=titolo)
    disp = pag[0] - 2 * marg
    fixed = {c: larghe[c] for c in cols if c in larghe}
    resto = [c for c in cols if c not in fixed]
    w_res = (disp - sum(fixed.values())) / max(len(resto), 1)
    widths = [fixed.get(c, w_res) for c in cols]
    f = 5.9
    st_h = ParagraphStyle("h", fontName=nome, fontSize=f - 0.7, leading=f + 0.8, alignment=1, textColor=colors.HexColor("#667085"))
    st_t = ParagraphStyle("t", fontName=grass, fontSize=13, leading=16, textColor=colors.HexColor("#14213D"))
    st_i = ParagraphStyle("i", fontName=nome, fontSize=8, leading=10.5, textColor=colors.HexColor("#14213D"))
    st_n = ParagraphStyle("n", fontName=nome, fontSize=6.6, leading=8.6, textColor=colors.HexColor("#667085"))

    dati = [[Paragraph(_esc(c), st_h) for c in cols]]
    stile = [("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#F9FAFB")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
             ("LINEBELOW", (0, 0), (-1, 0), 0.6, colors.HexColor("#E4E7EC")),
             ("GRID", (0, 1), (-1, -1), 0.25, colors.HexColor("#E4E7EC")),
             ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2),
             ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5)]
    for i in range(len(tab)):
        riga = []
        for j, c in enumerate(cols):
            v = tab.iat[i, j]
            try:
                t = sty._display_funcs[(i, j)](v)
            except Exception:
                t = "" if v is None or (not isinstance(v, str) and v != v) else str(v)
            bg, fg, bold = _css(ctx.get((i, j), []))
            sx = c in a_sinistra
            p = ParagraphStyle("c", fontName=grass if bold else nome, fontSize=f, leading=f + 1.6, alignment=0 if sx else 2,
                               textColor=colors.HexColor(fg) if fg else colors.HexColor("#14213D"))
            riga.append(Paragraph(_esc(t), p))
            if bg:
                stile.append(("BACKGROUND", (j, i + 1), (j, i + 1), colors.HexColor(bg)))
        dati.append(riga)
    tb = Table(dati, colWidths=widths, repeatRows=1)
    tb.setStyle(TableStyle(stile))
    el = [Paragraph(_esc(titolo), st_t), Spacer(1, 4)]
    for r in righe_intro:
        el.append(Paragraph(r, st_i))
    el += [Spacer(1, 8), tb]
    if note:
        el += [Spacer(1, 8), Paragraph(note, st_n)]
    doc.build(el)
    return buf.getvalue()
