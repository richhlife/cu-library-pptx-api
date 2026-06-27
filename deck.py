"""Builds a branded board-package .pptx from a payload dict. Pure python-pptx."""
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

WW, HH, MX, CW = 13.33, 7.5, 0.6, 12.13
WHITE = RGBColor(0xFF, 0xFF, 0xFF); SOFT = RGBColor(0x6B, 0x76, 0x85); FAINT = RGBColor(0x9A, 0xA4, 0xAE)
PAPER = RGBColor(0xF4, 0xF6, 0xF8); LINE = RGBColor(0xE2, 0xE7, 0xEC); ROWALT = RGBColor(0xFA, 0xFB, 0xFC)
REDBG = RGBColor(0xFB, 0xEC, 0xEA); RED = RGBColor(0xB2, 0x3A, 0x30); REDINK = RGBColor(0x7A, 0x2A, 0x22)
INK = RGBColor(0x13, 0x20, 0x2B); RING1 = RGBColor(0x35, 0x55, 0x70); RING2 = RGBColor(0x2C, 0x4A, 0x64)
LABEL = RGBColor(0x8A, 0x9A, 0xA8); SUBWHITE = RGBColor(0xC7, 0xD2, 0xDC)


def _hex(h, default):
    try:
        h = (h or default).lstrip("#")
        return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except Exception:
        d = default.lstrip("#")
        return RGBColor(int(d[0:2], 16), int(d[2:4], 16), int(d[4:6], 16))


def _usd(n):
    try: return "$" + format(int(round(float(n))), ",")
    except Exception: return str(n) if n is not None else ""


def _M(n):
    try: return "$" + ("%.1f" % (float(n) / 1e6)) + "M"
    except Exception: return str(n) if n is not None else ""


def _var(a, b):
    try:
        d = float(a) - float(b)
        return ("+" if d >= 0 else "\u2212") + _usd(abs(d))
    except Exception:
        return ""


def build(p):
    NAVY = _hex((p.get("theme") or {}).get("primary"), "#16314E")
    ACCENT = _hex((p.get("theme") or {}).get("accent"), "#1F9ED4")
    name = ((p.get("theme") or {}).get("name") or "Credit Union")
    period = p.get("period") or {}
    plabel = period.get("label") or ""
    foot = (plabel.upper() + "  \u00b7  BOARD REPORT") if plabel else "BOARD REPORT"

    prs = Presentation(); prs.slide_width = Inches(WW); prs.slide_height = Inches(HH)
    BLANK = prs.slide_layouts[6]

    def slide(): return prs.slides.add_slide(BLANK)

    def rect(s, x, y, w, h, color, shape=MSO_SHAPE.RECTANGLE, radius=None):
        sp = s.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
        sp.fill.solid(); sp.fill.fore_color.rgb = color; sp.line.fill.background(); sp.shadow.inherit = False
        if radius is not None:
            try: sp.adjustments[0] = radius
            except Exception: pass
        return sp

    def outline(s, x, y, w, h, color, wpt=1.25):
        sp = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(w), Inches(h))
        sp.fill.background(); sp.line.color.rgb = color; sp.line.width = Pt(wpt); sp.shadow.inherit = False
        return sp

    def txt(s, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font="Calibri"):
        tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame
        tf.word_wrap = True; tf.vertical_anchor = anchor
        tf.margin_left = 0; tf.margin_right = 0; tf.margin_top = 0; tf.margin_bottom = 0
        for i, (t, sz, c, b) in enumerate(runs):
            par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            par.alignment = align
            r = par.add_run(); r.text = str(t); r.font.size = Pt(sz); r.font.bold = b; r.font.color.rgb = c; r.font.name = font
        return tb

    def header(s, title, page):
        rect(s, 0, 0, WW, 1.15, NAVY)
        txt(s, MX, 0, 9.5, 1.0, [(title, 26, WHITE, True)], font="Georgia", anchor=MSO_ANCHOR.MIDDLE)
        rect(s, MX, 0.95, 0.5, 0.045, ACCENT)
        rect(s, MX, 6.98, CW, 0.012, LINE)
        txt(s, MX, 7.04, 8, 0.3, [(foot, 9, FAINT, False)])
        txt(s, WW - MX - 1.8, 7.04, 1.8, 0.3, [("PAGE " + str(page), 9, FAINT, False)], align=PP_ALIGN.RIGHT)

    def tile(s, x, y, w, h, label, value, note):
        rect(s, x, y, w, h, PAPER, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.07)
        rect(s, x, y, 0.06, h, ACCENT)
        txt(s, x + 0.28, y + 0.2, w - 0.5, 0.3, [(label, 11.5, SOFT, False)])
        txt(s, x + 0.28, y + 0.5, w - 0.5, 0.62, [(value, 27, NAVY, True)], font="Georgia")
        if note: txt(s, x + 0.28, y + h - 0.46, w - 0.5, 0.35, [(note, 10.5, SOFT, False)])

    def tiles4(s, items, top=1.55):
        w = (WW - 2 * MX - 0.3) / 2; h = 1.95; gx = 0.3; gy = 0.28
        for i, (l, v, n) in enumerate(items):
            tile(s, MX + (i % 2) * (w + gx), top + (i // 2) * (h + gy), w, h, l, v, n)

    def set_cell(c, text, sz, color, bold, align):
        c.margin_left = Inches(0.12); c.margin_right = Inches(0.12); c.margin_top = Inches(0.03); c.margin_bottom = Inches(0.03)
        c.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = c.text_frame; tf.word_wrap = True; par = tf.paragraphs[0]; par.alignment = align
        r = par.add_run(); r.text = str(text); r.font.size = Pt(sz); r.font.bold = bold; r.font.color.rgb = color; r.font.name = "Calibri"

    def table(s, headers, rows, colw):
        nrows = len(rows) + 1; ncols = len(headers)
        gt = s.shapes.add_table(nrows, ncols, Inches(MX), Inches(1.7), Inches(CW), Inches(0.34 * nrows)).table
        gt.first_row = False; gt.horz_banding = False
        tot = sum(colw)
        for ci, cw in enumerate(colw): gt.columns[ci].width = Inches(cw / tot * CW)
        for ci, htext in enumerate(headers):
            cell = gt.cell(0, ci); cell.fill.solid(); cell.fill.fore_color.rgb = NAVY
            set_cell(cell, htext, 11, WHITE, True, PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT)
        for ri, (cells, total) in enumerate(rows):
            for ci, val in enumerate(cells):
                cell = gt.cell(ri + 1, ci); cell.fill.solid()
                cell.fill.fore_color.rgb = (PAPER if total else (WHITE if ri % 2 == 0 else ROWALT))
                set_cell(cell, val, 11, (NAVY if total else INK), total, PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.RIGHT)
        for rr in gt.rows: rr.height = Inches(0.34)

    # ---- Cover ----
    s = slide(); rect(s, 0, 0, WW, HH, NAVY)
    outline(s, WW - 3.2, -3.4, 7.4, 7.4, RING1)
    outline(s, WW - 1.7, -2.0, 5.0, 5.0, RING2)
    txt(s, MX, 0.85, 11, 0.4, [(name.upper(), 13, ACCENT, True)])
    txt(s, MX, 2.35, 11.5, 1.4, [("Board Report", 56, WHITE, True)], font="Georgia")
    if plabel: txt(s, MX + 0.03, 3.78, 11, 0.4, [(plabel.upper(), 14, SUBWHITE, False)])
    rect(s, MX, 5.05, 3.0, 0.03, ACCENT)
    txt(s, MX, 5.45, 5, 0.3, [("MEETING", 10, LABEL, True)])
    txt(s, MX, 5.78, 6, 0.4, [(period.get("meeting") or "", 14, WHITE, False)])
    txt(s, 6.8, 5.45, 5, 0.3, [("LOCATION", 10, LABEL, True)])
    txt(s, 6.8, 5.78, 6, 0.4, [(period.get("location") or "", 14, WHITE, False)])

    # ---- Financial highlights ----
    s = slide(); header(s, "Financial Highlights", 2)
    tiles4(s, [(h.get("label", ""), h.get("value", ""), h.get("note", "")) for h in (p.get("headline") or [])[:4]])

    # ---- Balance sheet ----
    s = slide(); header(s, "Statement of Financial Condition", 3)
    table(s, ["Line", "Actual", "Budget", "Variance"],
          [([r.get("label", ""), _usd(r.get("actual")), _usd(r.get("budget")), _var(r.get("actual"), r.get("budget"))], bool(r.get("total"))) for r in (p.get("balanceSheet") or [])],
          [4.33, 2.6, 2.6, 2.6])

    # ---- Income ----
    s = slide(); header(s, "Income Statement (YTD)", 4)
    table(s, ["Line (YTD)", "Actual", "Budget", "Variance"],
          [([r.get("label", ""), _usd(r.get("actual")), _usd(r.get("budget")), _var(r.get("actual"), r.get("budget"))], bool(r.get("total"))) for r in (p.get("incomeStatement") or [])],
          [4.33, 2.6, 2.6, 2.6])

    # ---- Lending ----
    s = slide(); header(s, "Lending & Collections", 5)
    L = p.get("lending") or {}; act = L.get("activity") or []
    a0 = act[0] if len(act) > 0 else {}; a1 = act[1] if len(act) > 1 else {}
    w = (WW - 2 * MX - 0.3) / 2; h = 1.7
    litems = [
        ("Consumer loans YTD", _M(a0.get("ytd")), (str(a0.get("goal")) + " to goal") if a0.get("goal") else ""),
        ("Mortgage loans YTD", _M(a1.get("ytd")), (str(a1.get("goal")) + " to goal") if a1.get("goal") else ""),
        ("Delinquency ratio", str(L.get("delinquencyRatio", "")) + "%", (str(L.get("over60Count")) + " loans over 60 days") if L.get("over60Count") is not None else ""),
        ("Recoveries YTD", _usd(L.get("recoveriesYTD")), ("charge-offs " + _usd(L.get("chargeOffTotal"))) if L.get("chargeOffTotal") is not None else ""),
    ]
    for i, (l, v, n) in enumerate(litems):
        tile(s, MX + (i % 2) * (w + 0.3), 1.55 + (i // 2) * (h + 0.25), w, h, l, v, n)
    fy = 1.55 + 2 * (h + 0.25)
    dr = L.get("delinquencyRatio"); drt = L.get("delinquencyRatioTable")
    flagged = (dr is not None and drt is not None and dr != drt)
    rect(s, MX, fy, CW, 0.82, (REDBG if flagged else PAPER), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
    rect(s, MX, fy, 0.06, 0.82, (RED if flagged else ACCENT))
    txt(s, MX + 0.28, fy + 0.13, 4, 0.3, [(("CONSISTENCY FLAG" if flagged else "CONSISTENCY CHECK"), 10, (RED if flagged else SOFT), True)])
    msg = ("Delinquency reads %s%% in the narrative vs %s%% in the by-product table. Reconcile before distribution." % (dr, drt)) if flagged else "No figure inconsistencies detected across the book."
    txt(s, MX + 0.28, fy + 0.4, CW - 0.6, 0.4, [(msg, 11.5, (REDINK if flagged else SOFT), False)])

    # ---- Operations ----
    s = slide(); header(s, "Operations & Risk", 6)
    O = p.get("ops") or {}; sb = O.get("sb1075") or {}
    tm = O.get("totalMembers"); tm = ("{:,}".format(tm) if isinstance(tm, (int, float)) else (tm or ""))
    oitems = [
        ("New members", O.get("newMembers", ""), ((tm + " total") + ((" \u00b7 " + str(O.get("closures")) + " closures") if O.get("closures") is not None else "")) if tm else ""),
        ("Member chatbot", "~" + str(O.get("sunnyConvosPerWeek", "")) + "/wk", ("live " + str(O.get("sunnyWeeks")) + " weeks") if O.get("sunnyWeeks") is not None else ""),
        ("SB 1075 fee cap", "$" + str(sb.get("planned", "")), ("down from $" + str(sb.get("current", "")) + ((" \u00b7 " + str(sb.get("effective"))) if sb.get("effective") else "")) if sb.get("current") is not None else ""),
        ("Escalated reviews", O.get("fraudReviews", ""), "member detail secured"),
    ]
    tiles4(s, oitems)

    # ---- CEO ----
    s = slide(); header(s, "CEO / President's Report", 7)
    tb = s.shapes.add_textbox(Inches(0.7), Inches(1.65), Inches(CW - 0.2), Inches(5.0)); tf = tb.text_frame; tf.word_wrap = True
    bl = p.get("ceoHighlights") or ["No CEO report highlights were provided."]
    for i, b in enumerate(bl):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.space_after = Pt(14); par.line_spacing = 1.05
        r = par.add_run(); r.text = "\u2022   " + str(b); r.font.size = Pt(15); r.font.color.rgb = INK; r.font.name = "Calibri"

    out = io.BytesIO(); prs.save(out); return out.getvalue()
