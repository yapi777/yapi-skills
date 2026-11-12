#!/usr/bin/env python3
"""進捗報告スライドの共通パーツ（配色・チップ・パネル・タイル・表）。

python-pptx から import して使う。スライドの「文字だけで下半分が真っ白」を避け、
テーマ配色（accent1=青 / accent2=オレンジ）の2色に統一するための道具一式。

    import sys
    sys.path.insert(0, r"<スキルフォルダ>/scripts")   # SKILL.md では ${CLAUDE_SKILL_DIR}/scripts
    import deck_style
    deck_style.set_palette(*deck_style.palette_from_theme(path))   # 自分の資料をベースにするとき（任意）
    from deck_style import *                   # set_palette の「後」に import する

    prs = Presentation(path)
    s = prs.slides[6]
    remove_placeholder(s)                      # 既定の本文プレースホルダを外す
    add_chip(s, FULL_LEFT, Emu(1180000), CHIP_W, Emu(600000), "指摘", ORANGE, sz=20)
    add_panel(s, FULL_LEFT + CHIP_W + Emu(140000), Emu(1180000),
              FULL_WIDTH - CHIP_W - Emu(140000), Emu(900000), ORANGE_LIGHT,
              [[("査読者の指摘文", False, None)]], sz=18)
"""
import re
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# --- テーマ由来の統一パレット -------------------------------------------
# 既定は同梱テンプレートのテーマ accent1=#156082（青） / accent2=#E97132（オレンジ）の2色構成。
# 別のテーマの資料をベースにするときは set_palette() / palette_from_theme() で差し替える。
# 独自の赤（#C00000 など）は使わない — 資料全体で浮く。
NAVY = RGBColor(0x10, 0x48, 0x62)          # accent1 lumMod75。見出しチップの背景
ACCENT1 = RGBColor(0x15, 0x60, 0x82)       # accent1 そのもの。強調文字
BLUE_LIGHT = RGBColor(0xDB, 0xE7, 0xF3)    # 本文パネルの淡青
ORANGE = RGBColor(0xE9, 0x71, 0x32)        # accent2。指摘・注意のチップ背景
ORANGE_DARK = RGBColor(0xC5, 0x5A, 0x11)   # accent2 lumMod75。オレンジの強調文字
ORANGE_LIGHT = RGBColor(0xFB, 0xE3, 0xD6)  # 指摘パネルの淡橙
GRAY = RGBColor(0x59, 0x59, 0x59)          # 補足・出典
GRAY_LIGHT = RGBColor(0xF0, 0xF0, 0xF0)    # 中立パネル
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x1A, 0x1A, 0x1A)

_DEFAULT_ACCENTS = ("156082", "E97132")
_DEFAULT_PALETTE = (NAVY, ACCENT1, BLUE_LIGHT, ORANGE, ORANGE_DARK, ORANGE_LIGHT)


def _hex(c):
    c = str(c).lstrip("#").upper()
    return RGBColor(int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16))


def _lum(c, mod, off=0.0):
    """PowerPoint の lumMod / lumOff と同じ計算（HSL の明度を mod 倍して off を足す）。"""
    import colorsys
    r, g, b = (int(str(c)[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = min(1.0, max(0.0, l * mod + off))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return RGBColor(*(round(v * 255) for v in (r, g, b)))


def set_palette(accent1, accent2):
    """配色をテーマの2色から作り直す（'156082' / '#E97132' / RGBColor のどれでもよい）。

    派生色（濃い色＝lumMod75、淡い色＝lumMod20+lumOff80）も計算し直す。
    `from deck_style import *` はこの関数を呼んだ「後」に行う（先に import した名前は古い色のまま）。
    """
    global NAVY, ACCENT1, BLUE_LIGHT, ORANGE, ORANGE_DARK, ORANGE_LIGHT
    a1, a2 = _hex(accent1), _hex(accent2)
    if (str(a1), str(a2)) == _DEFAULT_ACCENTS:     # 同梱テンプレートの2色なら実測値をそのまま使う
        NAVY, ACCENT1, BLUE_LIGHT, ORANGE, ORANGE_DARK, ORANGE_LIGHT = _DEFAULT_PALETTE
    else:
        ACCENT1, NAVY, BLUE_LIGHT = a1, _lum(a1, 0.75), _lum(a1, 0.2, 0.8)
        ORANGE, ORANGE_DARK, ORANGE_LIGHT = a2, _lum(a2, 0.75), _lum(a2, 0.2, 0.8)
    return dict(NAVY=NAVY, ACCENT1=ACCENT1, BLUE_LIGHT=BLUE_LIGHT,
                ORANGE=ORANGE, ORANGE_DARK=ORANGE_DARK, ORANGE_LIGHT=ORANGE_LIGHT)


def palette_from_theme(pptx_path):
    """pptx のテーマ（ppt/theme/theme1.xml）から accent1 / accent2 を読む。戻り値は ('156082', 'E97132') の形。"""
    import zipfile
    with zipfile.ZipFile(pptx_path) as z:
        names = sorted(n for n in z.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n))
        xml = z.read(names[0]).decode("utf-8")
    def pick(tag):
        m = re.search(rf"<a:{tag}>\s*<a:(?:srgbClr val|sysClr [^>]*lastClr)=\"([0-9A-Fa-f]{{6}})\"", xml)
        if not m:
            raise ValueError(f"{tag} が theme に見つからない")
        return m.group(1).upper()
    return pick("accent1"), pick("accent2")

# 意味と色の対応（資料全体で固定する）
#   指摘・問題・従来手法 → ORANGE / ORANGE_LIGHT / ORANGE_DARK
#   対応・提案・結論     → NAVY / BLUE_LIGHT / ACCENT1

# --- レイアウト定数（4:3, 9144000 x 6858000 EMU） -----------------------
SLIDE_W = Emu(9144000)
SLIDE_H = Emu(6858000)
FULL_LEFT = Emu(450000)      # 全幅パネルの左端
FULL_WIDTH = Emu(8244000)    # 全幅パネルの幅
CHIP_W = Emu(1150000)        # 「指摘」「対応」チップの幅
BODY_TOP = Emu(1150000)      # タイトル下線の下、本文開始位置
BODY_BOTTOM = Emu(6200000)   # ここより下には置かない（下余白の目安）
COL_W = Emu(4000000)         # 2段組みの1列幅
COL_GAP = Emu(244000)        # 2段組みの間隔

# 折り返しの目安: 全幅パネル（FULL_WIDTH - チップ）で 18pt なら日本語 約30文字。
# これを超えると1〜2文字だけ次行に落ちて不格好になるので、文を削るか改行を明示する。


def remove_placeholder(slide, idx=1):
    """本文プレースホルダ（既定の箇条書き枠）を削除する。チップ/パネルで作り直す前に呼ぶ。"""
    for shp in list(slide.shapes):
        if shp.is_placeholder and shp.placeholder_format.idx == idx:
            shp._element.getparent().remove(shp._element)
            return True
    return False


def find_shape(slide, shape_id):
    for shp in slide.shapes:
        if shp.shape_id == shape_id:
            return shp
    return None


def _no_shadow(shape):
    shape.shadow.inherit = False


def add_chip(slide, x, y, w, h, text, fill, text_color=WHITE, sz=18,
             align=PP_ALIGN.CENTER):
    """見出しチップ（塗りつぶし＋白抜き太字）。研究目的スライドの「目的」バッジと同じ型。"""
    chip = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    chip.fill.solid()
    chip.fill.fore_color.rgb = fill
    chip.line.fill.background()
    _no_shadow(chip)
    tf = chip.text_frame
    tf.margin_left = tf.margin_right = Emu(80000)
    tf.margin_top = tf.margin_bottom = Emu(20000)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(sz)
    r.font.bold = True
    r.font.color.rgb = text_color
    r.font.name = "メイリオ"
    return chip


def add_panel(slide, x, y, w, h, fill, paragraphs, sz=18, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.MIDDLE, line_spacing=None):
    """本文パネル。paragraphs は [[(text, bold, color), ...], ...]（段落 → ラン）。

    color に None を渡すと BLACK。sz は必ず 16 以上（スキルの絶対ルール）。
    """
    assert sz >= 16, "文字サイズは16pt以上（スキルの絶対ルール）"
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    box.fill.solid()
    box.fill.fore_color.rgb = fill
    box.line.fill.background()
    _no_shadow(box)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Emu(180000)
    tf.margin_right = Emu(180000)
    tf.margin_top = Emu(120000)
    tf.margin_bottom = Emu(120000)
    for i, para in enumerate(paragraphs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        for text, bold, color in para:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(sz)
            r.font.bold = bold
            r.font.color.rgb = color if color else BLACK
            r.font.name = "メイリオ"
    return box


def add_issue_response(slide, issue_paras, response_paras, issue_h=Emu(900000),
                       gap=Emu(270000), response_h=Emu(3050000), top=BODY_TOP,
                       issue_sz=18, response_sz=18, response_line_spacing=1.7):
    """「指摘（オレンジ）→ 対応（青）」の定番2段組みを一括で作る。

    査読対応・課題と対策など「問題 → 打ち手」の構造を持つスライドの標準形。
    戻り値は (指摘チップ, 指摘パネル, 対応チップ, 対応パネル)。
    """
    body_x = FULL_LEFT + CHIP_W + Emu(140000)
    body_w = FULL_WIDTH - CHIP_W - Emu(140000)
    c1 = add_chip(slide, FULL_LEFT, top, CHIP_W, issue_h, "指摘", ORANGE, sz=min(20, issue_sz + 2))
    p1 = add_panel(slide, body_x, top, body_w, issue_h, ORANGE_LIGHT, issue_paras, sz=issue_sz)
    y2 = top + issue_h + gap
    c2 = add_chip(slide, FULL_LEFT, y2, CHIP_W, response_h, "対応", NAVY, sz=min(20, response_sz + 2))
    p2 = add_panel(slide, body_x, y2, body_w, response_h, BLUE_LIGHT, response_paras,
                   sz=response_sz, line_spacing=response_line_spacing)
    return c1, p1, c2, p2


def add_tiles(slide, tiles, top=Emu(1300000), tile_w=COL_W, tile_h=Emu(1950000),
              gap=COL_GAP, sz=17):
    """2列グリッドのKPIタイル。tiles は [(fill, paragraphs), ...]（左上→右上→左下…）。

    箇条書き4点程度を面で見せたいときに使うと、下半分の空白が埋まる。
    """
    out = []
    for i, (fill, paras) in enumerate(tiles):
        col, row = i % 2, i // 2
        x = FULL_LEFT + col * (tile_w + gap)
        y = top + row * (tile_h + gap)
        out.append(add_panel(slide, x, y, tile_w, tile_h, fill, paras, sz=sz,
                             align=PP_ALIGN.CENTER, line_spacing=1.25))
    return out


def add_bars(slide, bars, top=Emu(3600000), bar_h=Emu(820000), gap=Emu(180000), sz=19):
    """全幅の帯を縦に積む。表やグラフの下に「読み取れること」を並べるのに使う。

    bars は [(fill, paragraphs), ...]。1本だけ NAVY 塗り＋白文字にすると主役が立つ。
    """
    out = []
    for i, (fill, paras) in enumerate(bars):
        y = top + i * (bar_h + gap)
        out.append(add_panel(slide, FULL_LEFT, y, FULL_WIDTH, bar_h, fill, paras,
                             sz=sz, align=PP_ALIGN.CENTER))
    return out


def add_checklist(slide, items, top=Emu(2650000), row_h=Emu(950000),
                  gap=Emu(180000), sz=17, mark_w=Emu(600000)):
    """チェックマーク付きの確認事項リスト。「やったこと／確認したこと」の列挙に。"""
    out = []
    for i, text in enumerate(items):
        y = top + i * (row_h + gap)
        add_chip(slide, FULL_LEFT, y, mark_w, row_h, "✓", NAVY, sz=26)
        out.append(add_panel(slide, FULL_LEFT + mark_w + Emu(120000), y,
                             FULL_WIDTH - mark_w - Emu(120000), row_h,
                             GRAY_LIGHT, [[(text, False, None)]], sz=sz,
                             line_spacing=1.15))
    return out


def set_runs(shape, paragraphs, sz=18, align=PP_ALIGN.LEFT, line_spacing=None,
             anchor=MSO_ANCHOR.MIDDLE):
    """既存パネルの中身を丸ごと入れ替える（文章量を削るときに使う）。"""
    assert sz >= 16, "文字サイズは16pt以上（スキルの絶対ルール）"
    tf = shape.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    first = tf.paragraphs[0]
    for r in list(first.runs):
        r._r.getparent().remove(r._r)
    for i, para in enumerate(paragraphs):
        p = first if i == 0 else tf.add_paragraph()
        p.alignment = align
        if line_spacing:
            p.line_spacing = line_spacing
        for text, bold, color in para:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(sz)
            r.font.bold = bold
            r.font.color.rgb = color if color else BLACK
            r.font.name = "メイリオ"


def recolor_deck(prs, fill_map=None, line_map=None, font_map=None):
    """既存の色を一括で置き換える（旧デザインをテーマ配色へ寄せるとき用）。

    キーは 'C00000' のような大文字16進文字列、値は RGBColor。
    表（GraphicFrame）のセル塗り・文字色も対象にする。
    """
    fill_map = fill_map or {}
    line_map = line_map or {}
    font_map = font_map or {}

    def fix_runs(tf):
        for p in tf.paragraphs:
            for r in p.runs:
                try:
                    if r.font.color and r.font.color.type is not None and r.font.color.rgb is not None:
                        cur = str(r.font.color.rgb)
                        if cur in font_map:
                            r.font.color.rgb = font_map[cur]
                except Exception:
                    pass

    for slide in prs.slides:
        for shp in slide.shapes:
            try:
                if shp.fill.type == 1 and str(shp.fill.fore_color.rgb) in fill_map:
                    shp.fill.fore_color.rgb = fill_map[str(shp.fill.fore_color.rgb)]
            except Exception:
                pass
            try:
                if shp.line.fill.type == 1 and str(shp.line.color.rgb) in line_map:
                    shp.line.color.rgb = line_map[str(shp.line.color.rgb)]
            except Exception:
                pass
            if shp.has_text_frame:
                fix_runs(shp.text_frame)
            if getattr(shp, "has_table", False) and shp.has_table:
                for row in shp.table.rows:
                    for cell in row.cells:
                        try:
                            if cell.fill.type == 1 and str(cell.fill.fore_color.rgb) in fill_map:
                                cell.fill.fore_color.rgb = fill_map[str(cell.fill.fore_color.rgb)]
                        except Exception:
                            pass
                        fix_runs(cell.text_frame)


# ---------------------------------------------------------------------------
# 主張行・結果表・太字の色戻し
# ---------------------------------------------------------------------------
from pptx.oxml.ns import qn as _qn
from lxml import etree as _etree


def add_claim(slide, text, y=Emu(860000), sz=20):
    """タイトル直下の「主張」1文（太字・標準色）。本文はこの下 y≈1,450,000 から始める。

    太字には色を付けない（青・オレンジの太字は淡色パネルの背景と紛れて見づらい）。
    全幅で 20pt なら日本語 約30文字まで。超えると折り返すので文言を削る。
    """
    box = slide.shapes.add_textbox(FULL_LEFT, y, FULL_WIDTH, Emu(500000))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tf.paragraphs[0].add_run()
    r.text = text
    r.font.size = Pt(sz)
    r.font.bold = True
    r.font.color.rgb = BLACK
    r.font.name = "メイリオ"
    return box


def _cell_bottom_border(cell, color):
    tcPr = cell._tc.get_or_add_tcPr()
    for side in ("L", "R", "T", "B"):
        for old in tcPr.findall(_qn(f"a:ln{side}")):
            tcPr.remove(old)
    for i, side in enumerate(("L", "R", "T", "B")):
        ln = _etree.Element(_qn(f"a:ln{side}"))
        if side == "B":
            ln.set("w", "12700")
            sf = _etree.SubElement(ln, _qn("a:solidFill"))
            _etree.SubElement(sf, _qn("a:srgbClr")).set("val", color)
        else:
            ln.set("w", "0")
            _etree.SubElement(ln, _qn("a:noFill"))
        tcPr.insert(i, ln)


def add_table(slide, x, y, col_w, row_h, rows, sz=16, highlight=(), name_colors=None,
              left_cols=(0,)):
    """結果表（濃紺ヘッダ＋白抜き太字、主役行は淡青、行の下に細い灰色線）。

    rows[0] がヘッダ。セルは str か (text, bold, color)。
    highlight: 淡青で塗る行番号（例: 提案手法の行）。主役行は太字にするが色は付けない。
    name_colors: {行番号: 色} — 1列目（手法名）の文字色。比較手法をオレンジにする用（太字にしない）。
    列幅の目安（16pt）: 「IDconst100」の太字ヘッダは 1.42M EMU、「平均トラック長」は 1.62M EMU 必要。
    足りないとヘッダが2行に割れて行高が伸び、下の要素と重なる。
    """
    nr, nc = len(rows), len(rows[0])
    gf = slide.shapes.add_table(nr, nc, x, y, sum(col_w, Emu(0)), row_h * nr)
    t = gf.table
    for j, w in enumerate(col_w):
        t.columns[j].width = w
    for i in range(nr):
        t.rows[i].height = row_h
        for j in range(nc):
            c = t.cell(i, j)
            v = rows[i][j]
            text, bold, color = (v, False, None) if isinstance(v, str) else v
            if i == 0:
                fill, bold, color = NAVY, True, WHITE
            elif i in highlight:
                fill, bold = BLUE_LIGHT, True
            else:
                fill = WHITE
            if i > 0 and j == 0 and name_colors and color is None and not bold:
                color = name_colors.get(i)
            if bold and i > 0:
                color = None            # 太字は標準色
            c.fill.solid(); c.fill.fore_color.rgb = fill
            c.margin_left = c.margin_right = Emu(50000)
            c.margin_top = c.margin_bottom = Emu(30000)
            c.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = c.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if (j in left_cols and i > 0) else PP_ALIGN.CENTER
            r = p.add_run()
            r.text = text
            r.font.size = Pt(sz); r.font.bold = bold; r.font.name = "メイリオ"
            r.font.color.rgb = color if color else BLACK
            _cell_bottom_border(c, "BFBFBF" if i > 0 else str(NAVY))
    tblPr = t._tbl.tblPr
    for k in ("bandRow", "firstRow"):
        if tblPr.get(k):
            tblPr.set(k, "0")
    return gf


def normalize_bold_color(prs, skip=(), colored=None):
    """太字ランに付いた青・オレンジの文字色を標準色(1A1A1A)に戻す。保存直前に呼ぶ。

    チップ・濃紺帯の白抜き太字（FFFFFF）は対象外なのでそのまま残る。
    skip: 触らないスライドの 0 始まり番号（研究背景など据え置きのスライド）。
    colored: 標準色に戻す色（16進）。省略時は今のパレットの青・オレンジ4色。
    戻り値は直した数。
    """
    if colored is None:
        colored = tuple(str(c).upper() for c in (ACCENT1, ORANGE_DARK, ORANGE, NAVY))
    def tfs(shapes):
        for sh in shapes:
            if sh.shape_type == 6:
                yield from tfs(sh.shapes)
            if sh.has_text_frame:
                yield sh.text_frame
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        yield c.text_frame
    n = 0
    for i, sl in enumerate(prs.slides):
        if i in skip:
            continue
        for tf in tfs(sl.shapes):
            for p in tf.paragraphs:
                for r in p.runs:
                    rPr = r._r.find(_qn("a:rPr"))
                    if rPr is None or rPr.get("b") != "1":
                        continue
                    sf = rPr.find(_qn("a:solidFill"))
                    if sf is None:
                        continue
                    clr = sf.find(_qn("a:srgbClr"))
                    if clr is not None and clr.get("val").upper() in colored:
                        clr.set("val", "1A1A1A")
                        n += 1
                        continue
                    sch = sf.find(_qn("a:schemeClr"))     # テーマ色（accent1/accent2）で付いた色も戻す
                    if sch is not None and sch.get("val") in ("accent1", "accent2"):
                        sf.remove(sch)
                        _etree.SubElement(sf, _qn("a:srgbClr")).set("val", "1A1A1A")
                        n += 1
    return n
