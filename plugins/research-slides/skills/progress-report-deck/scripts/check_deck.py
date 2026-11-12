#!/usr/bin/env python3
"""進捗報告 pptx の機械チェック（壊れ・はみ出し・泣き別れ・下の空白・配色・残り物・文言）。

    python check_deck.py <deck.pptx> [--base <ベースのpptx>] [--skip 3,4] [--min-pt 16]
    python check_deck.py --text <構成案.md>        # 構成案の文言だけ検査（create-doc から使う）

出力は 1 行 1 件:  slide7 [WARN] orphan: …
最後に必ず締めの行  DONE: <FAIL数> FAIL / <WARN数> WARN / <枚数> slides  を出す。
締めの行が無ければ検査が途中で落ちている（合格扱いにしない）。
終了コード: FAIL が 1 件以上なら 1、それ以外は 0（WARN だけなら 0）。

FAIL = PowerPoint で開けない・修復ダイアログが出る壊れ、16pt 未満の文字。
WARN = 見積もりや経験則による指摘。レンダリング画像で確かめてから直す（誤検出もある）。

文字の幅・行数は「全角=1字、半角=0.55字、メイリオの行の高さ=文字サイズ×1.5」の見積もり。
実際の折り返しとは 1〜2 字ずれることがあるので、orphan / overflow は画像で確認すること。

検査の考え方の一部（XML の健全性、テキスト量からの溢れ・泣き別れの見積もり、隠れ・空箱の判定、
AI 臭のする語のリスト化）は carnot-tech/jinba-consulting-pptx-skill（MIT License,
Copyright (c) 2026 Carnot AI Inc.）を参考に、4:3・16pt 以上のこのスキル向けに書き直したもの。
"""
import argparse
import math
import pathlib
import posixpath
import re
import sys
import unicodedata
import zipfile

from lxml import etree

HERE = pathlib.Path(__file__).resolve().parent
BUNDLED_TEMPLATE = (HERE.parent / "assets" / "template.pptx").resolve()
LEXICON = HERE.parent / "references" / "ai-smell-lexicon.md"

# --- 閾値（4:3 = 9144000 x 6858000 EMU） --------------------------------
EMU_PER_PT = 12700
SLIDE_W, SLIDE_H = 9144000, 6858000
FILL_MIN_BOTTOM = 4_700_000   # 本文の一番下の要素がここより上で終わる＝下の 1/3 が空白
HALF_W = 0.53                 # 半角1字の幅（全角=1）。メイリオの英数字を PowerPoint の PDF で実測
SPACE_W = 0.3                 # 半角スペースの幅
LINE_H = 1.2                  # 行送り＝文字サイズ×1.2×lnSpc（PowerPoint の PDF で実測）
ORPHAN_MAX = 4.0              # 最終行がこの字数（全角換算）以下なら泣き別れ
OVERFLOW_TOL = 1.08           # 見積もりの高さが箱の高さのこの倍を超えたらはみ出し
HIDDEN_RATIO = 0.85           # 文字のある図形のこの割合以上が後ろの不透明な図形に覆われたら隠れ
COLOR_TOL = 4                 # パレットの色との RGB 各成分の許容差
NEUTRAL_TOL = 10              # R/G/B の差がこの範囲なら灰色（中立色）として許す
DEFAULT_SZ = 18.0
DEFAULT_INS = 91440           # bodyPr の既定の左右余白
DEFAULT_TBINS = 45720         # bodyPr の既定の上下余白

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]
R = "{%s}" % NS["r"]
MC = "{%s}" % NS["mc"]
SCHEME_COLORS = {"bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4",
                 "accent5", "accent6", "hlink", "folHlink", "phClr", "dk1", "lt1", "dk2", "lt2"}
RPR_TAGS = {A + "rPr", A + "defRPr", A + "endParaRPr"}

LEFTOVER_RE = re.compile(r"〇〇|◯◯|○○|ＸＸ|\bXX+\b|\bTODO\b|\bTBD\b|YYYY|ここに.{0,8}(書く|入れる|入力)"
                         r"|テキストを入力|クリックして.{0,10}(入力|追加)|Click to (add|edit)|Lorem ipsum")
TITLE_POLITE_RE = re.compile(r"(です|ます|でした|ました|ません)[。！？]?$")
TITLE_ABOUT_RE = re.compile(r"(について|に関して|の紹介|を紹介(する|します)?|を(示す|整理する|まとめる|説明する))$")
CLAIM_LABEL_RE = re.compile(r"^(結果|考察|実験|検証)\s*[0-9０-９①-⑳]*$")
TOC_TITLE_RE = re.compile(r"今月行ったこと|今回行ったこと|目次|アジェンダ|本日の内容|Outline|Agenda", re.I)
YURE_PAIRS = [("サーバー", "サーバ"), ("ユーザー", "ユーザ"), ("パラメーター", "パラメータ"),
              ("コンピューター", "コンピュータ"), ("メモリー", "メモリ"), ("フォルダー", "フォルダ"),
              ("フィルター", "フィルタ"), ("データー", "データ"), ("モデル化", "モデリング")]


class Findings:
    def __init__(self):
        self.items = []

    def add(self, where, level, kind, msg):
        self.items.append((where, level, kind, msg))

    def fail(self, where, kind, msg):
        self.add(where, "FAIL", kind, msg)

    def warn(self, where, kind, msg):
        self.add(where, "WARN", kind, msg)

    def count(self, level):
        return sum(1 for it in self.items if it[1] == level)


def _short(text, n=30):
    text = re.sub(r"\s+", " ", text).strip()
    return text if len(text) <= n else text[:n] + "…"


def char_em(ch):
    """1字の幅（全角=1）。ギリシャ文字・数学記号（Δ, ≤, σ）はメイリオでは半角幅になる。"""
    if ch in "\t\r\n":
        return 0.0
    if ch == " ":
        return SPACE_W
    eaw = unicodedata.east_asian_width(ch)
    if eaw in ("F", "W"):
        return 1.0
    if eaw == "A":
        o = ord(ch)
        return HALF_W if o < 0x2000 or 0x2200 <= o < 0x2300 else 1.0
    return HALF_W


def text_em(s):
    """全角=1、半角=HALF_W として文字列の幅（字数）を返す。"""
    return sum(char_em(ch) for ch in s)


NO_LINE_START = set("、。，．・：；？！ー）」』】〕｝〉》ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ々ゝゞヽヾ),.:;!?]}%")
NO_LINE_END = set("（「『【〔｛〈《([{")
_WORD = re.compile(r"[A-Za-z0-9_\-./%+]")


def wrap_widths(s, cap):
    """PowerPoint の折り返しをまねて、各行の幅（全角換算）のリストを返す。

    1字ずつ詰めて、はみ出す字で改行する。行頭禁則（、。）ー っ など）は前の字ごと次の行へ送り、
    英数字の単語は途中で切らない（単語が1行より長いときは切る）。
    """
    if cap <= 0 or not s:
        return [text_em(s)]
    lines, start, w, i = [], 0, 0.0, 0
    n = len(s)
    while i < n:
        cw = char_em(s[i])
        if w + cw <= cap + 1e-6 or i == start:
            w += cw
            i += 1
            continue
        brk = i
        # 英数字の単語の途中なら単語の頭まで戻す
        if _WORD.match(s[brk]) and _WORD.match(s[brk - 1]):
            j = brk
            while j > start and _WORD.match(s[j - 1]):
                j -= 1
            if j > start:
                brk = j
        # 行頭禁則・行末禁則
        while brk - 1 > start and (s[brk] in NO_LINE_START or s[brk - 1] in NO_LINE_END):
            brk -= 1
        while brk < n and s[brk] == " ":      # 行頭の空白は消える
            brk += 1
        lines.append(text_em(s[start:brk].rstrip(" ")))
        start, i, w = brk, brk, 0.0
    lines.append(text_em(s[start:]))
    return lines


# --- 語のリスト（references/ai-smell-lexicon.md の機械チェック用ブロック） ---------
def load_lexicon(path=LEXICON):
    """```words … ``` ブロックを読む。1行1語、`re:` で始まる行は正規表現、`#` はコメント。"""
    pats = []
    try:
        txt = pathlib.Path(path).read_text(encoding="utf-8")
    except OSError:
        return pats
    m = re.search(r"```words\s*\n(.*?)```", txt, re.S)
    if not m:
        return pats
    for line in m.group(1).splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        word, _, hint = line.partition("=>")
        word, hint = word.strip(), hint.strip()
        if word.startswith("re:"):
            pats.append((re.compile(word[3:]), word[3:], hint))
        else:
            pats.append((re.compile(re.escape(word)), word, hint))
    return pats


def check_wording(texts, out, lexicon, one_sentence=True):
    """texts: [(where, text)]。AI 臭のする語・1項目2文以上・表記ゆれ・句読点の混在。"""
    seen_yure = {}
    puncts = {"、": [], "，": []}
    for where, t in texts:
        for rx, word, hint in lexicon:
            mm = rx.search(t)
            if mm:
                out.warn(where, "wording", f"「{mm.group(0)}」{('→ ' + hint) if hint else ''}（{_short(t)}）")
        if one_sentence and re.search(r"[。．](?=\s*\S)", t.strip()):
            out.warn(where, "wording", f"1項目に2文以上ある。箇条を分けるか1文に削る（{_short(t)}）")
        for long_, short_ in YURE_PAIRS:
            if long_ in t:
                seen_yure.setdefault((long_, short_), [set(), set()])[0].add(where)
            elif re.search(re.escape(short_) + r"(?!ー)", t):
                seen_yure.setdefault((long_, short_), [set(), set()])[1].add(where)
        for k in puncts:
            if k in t:
                puncts[k].append(where)
    for (long_, short_), (a, b) in seen_yure.items():
        if a and b:
            out.warn("deck", "wording", f"表記ゆれ「{long_}」({_where_list(a)})と「{short_}」({_where_list(b)})")
    if puncts["、"] and puncts["，"]:
        out.warn("deck", "wording",
                 f"読点の混在: 「、」({_where_list(puncts['、'])}) と「，」({_where_list(puncts['，'])})。どちらかに揃える")


def _where_list(ws):
    ws = sorted(set(ws), key=lambda w: (len(str(w)), str(w)))
    return ",".join(str(w) for w in ws[:6]) + ("…" if len(ws) > 6 else "")


# --- 1. パッケージ（ZIP・XML）の健全性 --------------------------------------
def check_package(path, out):
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        out.fail("deck", "xml", f"ZIP として開けない: {e}")
        return False
    names = z.namelist()
    dup = sorted({n for n in names if names.count(n) > 1})
    if dup:
        out.fail("deck", "xml", f"ZIP の中に同じ名前のエントリが重複: {', '.join(dup[:5])}（python-pptx で保存すると壊れる原因）")
    nameset = set(names)

    def rels_of(part):
        d, b = posixpath.split(part)
        rp = posixpath.join(d, "_rels", b + ".rels")
        if rp not in nameset:
            return {}
        root = etree.fromstring(z.read(rp))
        rels = {}
        for rel in root.findall("pr:Relationship", NS):
            tgt = rel.get("Target")
            if rel.get("TargetMode") != "External":
                tgt = posixpath.normpath(posixpath.join(d, tgt))
            rels[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], tgt, rel.get("TargetMode"))
        return rels

    try:
        ct = etree.fromstring(z.read("[Content_Types].xml"))
        overrides = {o.get("PartName").lstrip("/") for o in ct.findall("ct:Override", NS)}
        defaults = {d.get("Extension").lower() for d in ct.findall("ct:Default", NS)}
    except KeyError:
        out.fail("deck", "xml", "[Content_Types].xml が無い")
        return False

    pres_rels = rels_of("ppt/presentation.xml")
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    listed = []
    for sid in pres.iterfind(".//p:sldIdLst/p:sldId", NS):
        rid = sid.get(R + "id")
        if rid not in pres_rels:
            out.fail("deck", "xml", f"presentation.xml の sldIdLst が存在しないリレーション {rid} を指している")
            continue
        listed.append(pres_rels[rid][1])
    slide_rel_targets = {t for (typ, t, _) in pres_rels.values() if typ == "slide"}
    slide_parts = sorted((n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                         key=lambda n: int(re.search(r"(\d+)\.xml$", n).group(1)))
    for sp in slide_parts:
        if sp not in listed:
            how = "presentation.xml.rels には残っている" if sp in slide_rel_targets else "どこからも参照されていない"
            out.fail("deck", "xml", f"{sp} が sldIdLst に無いのに残っている（{how}）。rels・Content_Types・実体もまとめて消す")
    for t in slide_rel_targets:
        if t not in nameset:
            out.fail("deck", "xml", f"presentation.xml.rels が存在しない {t} を指している")

    for idx, part in enumerate(listed, start=1):
        where = f"slide{idx}"
        if part not in nameset:
            out.fail(where, "xml", f"{part} が ZIP に無い")
            continue
        if part not in overrides:
            out.fail(where, "xml", f"[Content_Types].xml に {part} の Override が無い")
        xml = z.read(part)
        root = etree.fromstring(xml)
        rels = rels_of(part)
        for el in root.iter():
            for attr in ("embed", "id", "link", "pict"):
                rid = el.get(R + attr)
                if rid is not None and rid not in rels:
                    out.fail(where, "xml", f"r:{attr}=\"{rid}\" が {posixpath.basename(part)}.rels に無い（切れた参照）")
            tag = el.tag
            if tag == A + "schemeClr" and el.get("val") not in SCHEME_COLORS:
                out.fail(where, "xml", f"不正なテーマ色名 schemeClr val=\"{el.get('val')}\"")
            if tag in (A + "ext", A + "chExt"):
                for k in ("cx", "cy"):
                    v = el.get(k)
                    if v is not None and v.lstrip("-").isdigit() and int(v) < 0:
                        out.fail(where, "xml", f"図形の大きさが負の値 ({k}={v})")
            if tag in RPR_TAGS and el.get("sz") is not None:
                sz = int(el.get("sz"))
                if not 100 <= sz <= 400000:
                    out.fail(where, "xml", f"文字サイズ sz={sz} が範囲外（100〜400000）")
        for rid, (typ, tgt, mode) in rels.items():
            if mode == "External":
                continue
            if tgt not in nameset:
                out.fail(where, "xml", f"リレーション {rid}（{typ}）の参照先 {tgt} が ZIP に無い")
            elif typ in ("image", "media", "video", "audio"):
                ext = tgt.rsplit(".", 1)[-1].lower()
                if tgt not in overrides and ext not in defaults:
                    out.fail(where, "xml", f"{tgt} の拡張子 .{ext} が [Content_Types].xml に登録されていない")
        spids = {c.get("id") for c in root.iterfind(".//p:cNvPr", NS)}
        for tn in root.iterfind(".//p:timing//p:spTgt", NS):
            if tn.get("spid") not in spids:
                out.fail(where, "xml", f"アニメーションが存在しない図形 spid={tn.get('spid')} を指している")
        if root.get("show") == "0":
            out.warn(where, "xml", "非表示スライドが残っている")
    return True


# --- 2. 文字サイズ（XML を直接走査。継承は見ない＝check-min-font.js と同じ） ------
def check_min_font(path, min_pt, out):
    z = zipfile.ZipFile(path)
    pres = etree.fromstring(z.read("ppt/presentation.xml"))
    rels = etree.fromstring(z.read("ppt/_rels/presentation.xml.rels"))
    rmap = {r.get("Id"): posixpath.normpath(posixpath.join("ppt", r.get("Target")))
            for r in rels.findall("pr:Relationship", NS)}
    for idx, sid in enumerate(pres.iterfind(".//p:sldIdLst/p:sldId", NS), start=1):
        part = rmap.get(sid.get(R + "id"))
        if not part or part not in z.namelist():
            continue
        root = etree.fromstring(z.read(part))
        where = f"slide{idx}"
        for el in root.iter(*RPR_TAGS):
            sz = el.get("sz")
            if sz and sz.isdigit() and int(sz) < round(min_pt * 100):
                para = el.getparent()
                while para is not None and para.tag != A + "p":
                    para = para.getparent()
                txt = "".join(para.itertext()) if para is not None else ""
                if el.tag == A + "endParaRPr" and txt.strip():
                    continue          # 文字のある段落の段落末記号は本文の run 側で判定する
                if el.tag == A + "defRPr":
                    continue          # 既定値（lstStyle 等）は実際の文字が無ければ見えない
                out.fail(where, "min-font", f"{int(sz) / 100:g}pt（{min_pt:g}pt 未満）「{_short(txt) or '(空の段落)'}」")
        for na in root.iterfind(".//a:normAutofit", NS):
            fs = na.get("fontScale")
            if fs and int(fs) < 100000:
                out.warn(where, "min-font", f"自動縮小 fontScale={int(fs) / 1000:g}% が掛かっている。継承の文字が {min_pt:g}pt 未満になっていないか確認")


# --- 3. 図形の幾何（python-pptx） --------------------------------------------
class Item:
    """スライド上の図形 1 つ（グループの中身はスライドの絶対座標に直したもの）。"""

    def __init__(self, shape, x, y, w, h, order):
        self.shape, self.x, self.y, self.w, self.h, self.order = shape, x, y, w, h, order
        self.text = ""
        self.has_text = False
        self.opaque = False
        self.is_picture = False
        self.is_line = False
        self.fill_rgb = None
        self.est_bottom = y + h
        self.ph_type = None


def _ph_type(shape):
    try:
        if shape.is_placeholder:
            return str(shape.placeholder_format.type).split(".")[-1].split(" ")[0]
    except Exception:
        pass
    return None


def _solid_fill_rgb(shape):
    """図形の塗りが単色なら 'RRGGBB'（テーマ色は解決せず 'scheme'）、塗りが無ければ None。"""
    spPr = shape._element.find(P + "spPr")
    if spPr is None:
        return None
    if spPr.find(A + "noFill") is not None:
        return None
    sf = spPr.find(A + "solidFill")
    if sf is not None:
        c = sf.find(A + "srgbClr")
        return c.get("val").upper() if c is not None else "scheme"
    if spPr.find(A + "gradFill") is not None or spPr.find(A + "pattFill") is not None or spPr.find(A + "blipFill") is not None:
        return "other"
    # 塗りの指定が無い → p:style の fillRef（idx>0 なら塗りあり）
    style = shape._element.find(P + "style")
    if style is not None:
        fr = style.find(A + "fillRef")
        if fr is not None and fr.get("idx", "0") != "0":
            return "scheme"
    return None


def _iter_items(slide):
    from pptx.shapes.shapetree import SlideShapeFactory
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    items = []
    order = [0]

    def add(shape, x, y, w, h):
        it = Item(shape, x, y, w, h, order[0])
        order[0] += 1
        it.ph_type = _ph_type(shape)
        st = None
        try:
            st = shape.shape_type
        except Exception:
            pass
        it.is_picture = st == MSO_SHAPE_TYPE.PICTURE or shape._element.tag == P + "pic"
        it.is_line = st in (MSO_SHAPE_TYPE.LINE,) or shape._element.tag == P + "cxnSp" or min(w, h) < 50000
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
            it.text = shape.text_frame.text
            it.has_text = bool(it.text.strip())
        if getattr(shape, "has_table", False) and shape.has_table:
            it.text = "\n".join(c.text_frame.text for r in shape.table.rows for c in r.cells)
            it.has_text = bool(it.text.strip())
            it.opaque = True
        it.fill_rgb = _solid_fill_rgb(shape) if shape._element.tag == P + "sp" else None
        if it.is_picture or (it.fill_rgb is not None and not it.is_line):
            it.opaque = True
        items.append(it)
        return it

    def walk(elms, parent_shapes, xf):
        for elm in elms:
            tag = elm.tag
            if tag == MC + "AlternateContent":
                choice = elm.find(MC + "Choice")
                inner = list(choice) if choice is not None else []
                walk(inner, parent_shapes, xf)
                continue
            if tag not in (P + "sp", P + "pic", P + "graphicFrame", P + "grpSp", P + "cxnSp"):
                continue
            shape = SlideShapeFactory(elm, parent_shapes)
            try:
                l, t, w, h = shape.left, shape.top, shape.width, shape.height
            except Exception:
                l = t = w = h = None
            if None in (l, t, w, h):
                continue
            x, y = xf(l, t)
            sx, sy = xf.scale
            if tag == P + "grpSp":
                gx = elm.find(P + "grpSpPr/" + A + "xfrm")
                choff = gx.find(A + "chOff") if gx is not None else None
                chext = gx.find(A + "chExt") if gx is not None else None
                cox, coy = (int(choff.get("x")), int(choff.get("y"))) if choff is not None else (l, t)
                cex, cey = (int(chext.get("cx")), int(chext.get("cy"))) if chext is not None else (w, h)
                kx = (w / cex if cex else 1.0) * sx
                ky = (h / cey if cey else 1.0) * sy
                walk(list(elm), shape.shapes, _Xf(x - cox * kx, y - coy * ky, kx, ky))
                continue
            add(shape, x, y, w * sx, h * sy)

    walk(list(slide.shapes._spTree), slide.shapes, _Xf(0, 0, 1.0, 1.0))
    return items


class _Xf:
    def __init__(self, ox, oy, kx, ky):
        self.ox, self.oy, self.scale = ox, oy, (kx, ky)

    def __call__(self, l, t):
        return self.ox + l * self.scale[0], self.oy + t * self.scale[1]


# 文字サイズ・段落の書式を、run → 段落 → 図形の lstStyle → レイアウト → マスターの順に解決する
def _lvl_ppr(lst, lvl):
    if lst is None:
        return None
    return lst.find(A + f"lvl{lvl + 1}pPr")


def _placeholder_chain(shape):
    """図形自身の lstStyle、レイアウト・マスターのプレースホルダの lstStyle、マスターの txStyles。"""
    chain = []
    txb = shape._element.find(P + "txBody")
    if txb is not None:
        chain.append(txb.find(A + "lstStyle"))
    if getattr(shape, "is_placeholder", False):
        try:
            lay = shape._base_placeholder
            while lay is not None:
                tb = lay._element.find(P + "txBody")
                if tb is not None:
                    chain.append(tb.find(A + "lstStyle"))
                lay = getattr(lay, "_base_placeholder", None)
        except Exception:
            pass
        try:
            master = shape.part.slide_layout.slide_master
            tx = master._element.find(P + "txStyles")
            pt = _ph_type(shape) or ""
            key = "titleStyle" if "TITLE" in pt else ("otherStyle" if pt in ("SLIDE_NUMBER", "DATE", "FOOTER") else "bodyStyle")
            if tx is not None:
                chain.append(tx.find(P + key))
        except Exception:
            pass
    else:
        try:
            pres = shape.part.package.presentation_part.presentation._element
            chain.append(pres.find(P + "defaultTextStyle"))
        except Exception:
            pass
    return chain


def _resolve(p_el, chain, getter, default):
    lvl = 0
    ppr = p_el.find(A + "pPr")
    if ppr is not None and ppr.get("lvl"):
        lvl = int(ppr.get("lvl"))
    v = getter(ppr) if ppr is not None else None
    if v is not None:
        return v
    for lst in chain:
        v = getter(_lvl_ppr(lst, lvl))
        if v is not None:
            return v
    return default


def _para_size(p_el, chain):
    for r in p_el.iter(A + "r"):
        rpr = r.find(A + "rPr")
        if rpr is not None and rpr.get("sz") and (r.findtext(A + "t") or "").strip():
            return int(rpr.get("sz")) / 100
    for tag in (A + "endParaRPr",):
        e = p_el.find(tag)
        if e is not None and e.get("sz") and not "".join(p_el.itertext()).strip():
            return int(e.get("sz")) / 100

    def g(ppr):
        if ppr is None:
            return None
        d = ppr.find(A + "defRPr")
        return int(d.get("sz")) / 100 if d is not None and d.get("sz") else None
    return _resolve(p_el, chain, g, DEFAULT_SZ)


def _para_lnspc(p_el, chain, sz):
    def g(ppr):
        if ppr is None:
            return None
        ls = ppr.find(A + "lnSpc")
        if ls is None:
            return None
        pct = ls.find(A + "spcPct")
        if pct is not None:
            return ("pct", int(pct.get("val")) / 100000)
        pts = ls.find(A + "spcPts")
        if pts is not None:
            return ("pts", int(pts.get("val")) / 100)
        return None
    v = _resolve(p_el, chain, g, ("pct", 1.0))
    return sz * LINE_H * v[1] if v[0] == "pct" else v[1]


def _para_spc(p_el, chain, sz, tag):
    def g(ppr):
        if ppr is None:
            return None
        s = ppr.find(A + tag)
        if s is None:
            return None
        pts = s.find(A + "spcPts")
        if pts is not None:
            return int(pts.get("val")) / 100
        pct = s.find(A + "spcPct")
        if pct is not None:
            return sz * LINE_H * int(pct.get("val")) / 100000
        return None
    return _resolve(p_el, chain, g, 0.0)


def _para_indent(p_el, chain):
    def g(ppr):
        if ppr is None or ppr.get("marL") is None:
            return None
        return int(ppr.get("marL"))
    return _resolve(p_el, chain, g, 0)


def _para_lines(p_el):
    """段落を <a:br/> で区切った行の文字列のリスト。"""
    lines, cur = [], []
    for el in p_el:
        if el.tag in (A + "r", A + "fld"):
            cur.append(el.findtext(A + "t") or "")
        elif el.tag == A + "br":
            lines.append("".join(cur))
            cur = []
        elif el.tag == "{http://schemas.microsoft.com/office/drawing/2010/main}m" or el.tag.endswith("}m"):
            cur.append("x" * 6)      # 数式は中身を数えず、短い式として扱う
    lines.append("".join(cur))
    return lines


def _body_props(shape):
    """bodyPr（図形自身 → レイアウトのプレースホルダ）から余白・折り返し・アンカー・自動調整を読む。"""
    out = {"l": None, "r": None, "t": None, "b": None, "wrap": None, "anchor": None, "autofit": None, "vert": None}
    bps = []
    txb = shape._element.find(P + "txBody")
    if txb is not None:
        bps.append(txb.find(A + "bodyPr"))
    if getattr(shape, "is_placeholder", False):
        try:
            lay = shape._base_placeholder
            while lay is not None:
                tb = lay._element.find(P + "txBody")
                if tb is not None:
                    bps.append(tb.find(A + "bodyPr"))
                lay = getattr(lay, "_base_placeholder", None)
        except Exception:
            pass
    for bp in bps:
        if bp is None:
            continue
        for k, attr in (("l", "lIns"), ("r", "rIns"), ("t", "tIns"), ("b", "bIns"), ("wrap", "wrap"),
                        ("anchor", "anchor"), ("vert", "vert")):
            if out[k] is None and bp.get(attr) is not None:
                out[k] = bp.get(attr)
        if out["autofit"] is None:
            for tag, name in ((A + "spAutoFit", "shape"), (A + "normAutofit", "norm"), (A + "noAutofit", "none")):
                if bp.find(tag) is not None:
                    out["autofit"] = name
                    break
    for k in ("l", "r"):
        out[k] = int(out[k]) if out[k] is not None else DEFAULT_INS
    for k in ("t", "b"):
        out[k] = int(out[k]) if out[k] is not None else DEFAULT_TBINS
    return out


def _text_frames(item):
    sh = item.shape
    if getattr(sh, "has_table", False) and sh.has_table:
        return []          # 表のセルは行の高さが伸びるので、幾何の見積もりはしない（文言の検査だけ）
    if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
        return [sh.text_frame]
    return []


def estimate_text(item, where, out, report=True):
    """図形内の文字の行数と高さを見積もり、orphan / overflow を報告する。戻り値は文字の高さ(EMU)。"""
    tfs = _text_frames(item)
    if not tfs or not item.has_text:
        return 0
    sh = item.shape
    bp = _body_props(sh)
    if bp["vert"] not in (None, "horz"):
        return 0
    chain = _placeholder_chain(sh)
    total_pt = 0.0
    total_lines = 0
    for p in tfs[0].paragraphs:
        p_el = p._p
        sz = _para_size(p_el, chain)
        lh = _para_lnspc(p_el, chain, sz)
        indent = _para_indent(p_el, chain)
        avail_pt = (item.w - bp["l"] - bp["r"] - indent) / EMU_PER_PT
        cap = avail_pt / sz if sz > 0 else 0
        nlines = 0
        for seg in _para_lines(p_el):
            if bp["wrap"] == "none" or cap <= 0:
                nlines += 1
                continue
            widths = wrap_widths(seg, cap)
            nlines += len(widths)
            if report and len(widths) >= 2 and 0 < widths[-1] <= ORPHAN_MAX:
                out.warn(where, "orphan",
                         f"最終行が約{math.ceil(widths[-1])}字だけになる見込み（{sz:g}pt・1行約{cap:.0f}字）"
                         f"→ 文を削るか <a:br/> で区切る「{_short(seg)}」")
        total_lines += nlines
        total_pt += nlines * lh + _para_spc(p_el, chain, sz, "spcBef") + _para_spc(p_el, chain, sz, "spcAft")
    text_h = total_pt * EMU_PER_PT + bp["t"] + bp["b"]
    if (report and total_lines >= 2 and bp["autofit"] != "shape" and bp["wrap"] != "none"
            and text_h > item.h * OVERFLOW_TOL):
        out.warn(where, "overflow",
                 f"文字が箱からはみ出す見込み（必要な高さ 約{text_h / 914400:.2f}in / 箱 {item.h / 914400:.2f}in）"
                 f"→ 削る・箱を広げる・スライドを分ける「{_short(item.text)}」")
    return text_h


def _area_overlap(a, b):
    w = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
    h = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
    return w * h if w > 0 and h > 0 else 0


def _is_title_slide(slide):
    return any(_ph_type(sh) == "CENTER_TITLE" for sh in slide.placeholders)


def _title_text(slide):
    for sh in slide.placeholders:
        if _ph_type(sh) in ("TITLE", "CENTER_TITLE"):
            return sh.text_frame.text.strip()
    return ""


# --- 4. 配色 -----------------------------------------------------------------
def build_palette(base_pptx):
    sys.path.insert(0, str(HERE))
    import deck_style
    a1, a2 = deck_style.palette_from_theme(base_pptx)
    named = deck_style.set_palette(a1, a2)
    allowed = {str(v).upper() for v in named.values()}
    for acc in (a1, a2):
        allowed.add(acc.upper())
        for mod, off in ((0.75, 0), (0.5, 0), (0.2, 0.8), (0.4, 0.6), (0.6, 0.4)):
            allowed.add(str(deck_style._lum(acc, mod, off)).upper())
    allowed |= {str(deck_style.GRAY).upper(), str(deck_style.GRAY_LIGHT).upper()}
    return allowed, (a1, a2)


def _rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def color_ok(hexv, allowed):
    r, g, b = _rgb(hexv)
    if max(r, g, b) - min(r, g, b) <= NEUTRAL_TOL:
        return True
    for c in allowed:
        cr, cg, cb = _rgb(c)
        if abs(r - cr) <= COLOR_TOL and abs(g - cg) <= COLOR_TOL and abs(b - cb) <= COLOR_TOL:
            return True
    return False


def check_colors(slide, where, allowed, out):
    root = slide._element
    bad = {}
    for el in root.iter(A + "srgbClr", A + "schemeClr"):
        anc = el.getparent()
        in_fallback = False
        while anc is not None:
            if anc.tag == MC + "Fallback":
                in_fallback = True
                break
            anc = anc.getparent()
        if in_fallback:
            continue
        if el.tag == A + "schemeClr":
            if el.get("val") in ("accent3", "accent4", "accent5", "accent6"):
                bad.setdefault(f"テーマの{el.get('val')}", 0)
                bad[f"テーマの{el.get('val')}"] += 1
            continue
        v = (el.get("val") or "").upper()
        if len(v) == 6 and not color_ok(v, allowed):
            bad.setdefault("#" + v, 0)
            bad["#" + v] += 1
    if bad:
        lst = "、".join(f"{k}×{n}" for k, n in sorted(bad.items()))
        out.warn(where, "color", f"パレットの2色（と灰色）以外の色: {lst}。写真の識別ラベル以外ならパレットの色に置き換える")


def check_bold_color(slide, where, out, allowed):
    """太字の run に色（中立色・白以外）が付いている。"""
    hits = []
    for r in slide._element.iter(A + "r"):
        rpr = r.find(A + "rPr")
        if rpr is None or rpr.get("b") != "1":
            continue
        sf = rpr.find(A + "solidFill")
        if sf is None:
            continue
        c = sf.find(A + "srgbClr")
        sch = sf.find(A + "schemeClr")
        colored = False
        if c is not None:
            v = c.get("val").upper()
            rr, gg, bb = _rgb(v)
            colored = max(rr, gg, bb) - min(rr, gg, bb) > NEUTRAL_TOL
        elif sch is not None and sch.get("val", "").startswith("accent"):
            colored = True
        if colored and (r.findtext(A + "t") or "").strip():
            hits.append(r.findtext(A + "t"))
    if hits:
        out.warn(where, "bold-color", f"太字に色が付いている {len(hits)} 箇所（{_short('／'.join(hits), 40)}）"
                 "→ normalize_bold_color() で標準色に戻す")


# --- 5. スライドごとの検査 --------------------------------------------------
def collect_texts(slide):
    """(図形, 段落テキスト) の列。表のセル・グループの中も含む。"""
    res = []

    def tfs(shapes):
        for sh in shapes:
            if getattr(sh, "shape_type", None) == 6:
                yield from tfs(sh.shapes)
                continue
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                yield sh, sh.text_frame
            if getattr(sh, "has_table", False) and sh.has_table:
                for row in sh.table.rows:
                    for c in row.cells:
                        yield sh, c.text_frame
    for sh, tf in tfs(slide.shapes):
        for p in tf.paragraphs:
            t = "".join(r.text for r in p.runs)
            if t.strip():
                res.append((sh, t))
    return res


def dummy_texts(template_path):
    from pptx import Presentation
    prs = Presentation(template_path)
    keep = set()
    n = len(prs.slides)
    for i, s in enumerate(prs.slides, start=1):
        if i in (1, n):
            continue          # タイトルとクロージングは枠として使い回すので比べない
        for sh, t in collect_texts(s):
            if _ph_type(sh) in ("TITLE", "SLIDE_NUMBER"):
                continue
            t = t.strip()
            if text_em(t) >= 8:
                keep.add(t)
    return keep


def check_slides(path, out, base, skip, lexicon):
    from pptx import Presentation
    try:
        prs = Presentation(path)
    except Exception as e:
        out.fail("deck", "xml", f"python-pptx で開けない: {e}")
        return 0
    if (prs.slide_width, prs.slide_height) != (SLIDE_W, SLIDE_H):
        out.warn("deck", "size", f"スライドサイズが 4:3（{SLIDE_W}x{SLIDE_H}）ではない: {prs.slide_width}x{prs.slide_height}")
    allowed, _ = build_palette(base or path)
    dummies = set()
    if base and pathlib.Path(base).resolve() == BUNDLED_TEMPLATE:
        dummies = dummy_texts(base)
    n = len(prs.slides)
    wording = []
    for idx, slide in enumerate(prs.slides, start=1):
        where = f"slide{idx}"
        if idx in skip:
            continue
        items = _iter_items(slide)
        title = _title_text(slide)
        is_title = _is_title_slide(slide)
        is_closing = idx == n and not title
        is_toc = bool(TOC_TITLE_RE.search(title))

        # 幾何: はみ出し・泣き別れ・枠外・隠れ・空箱
        for it in items:
            if it.ph_type in ("SLIDE_NUMBER", "DATE", "FOOTER"):
                continue
            text_h = estimate_text(it, where, out)
            if it.has_text and it.fill_rgb is None and not it.is_picture:
                bp = _body_props(it.shape)
                th = min(text_h, it.h) if bp["autofit"] != "shape" else text_h
                anc = bp["anchor"] or "t"
                if anc == "ctr":
                    it.est_bottom = it.y + it.h / 2 + th / 2
                elif anc == "b":
                    it.est_bottom = it.y + it.h
                else:
                    it.est_bottom = it.y + th
                if bp["autofit"] == "shape" and it.y + text_h > SLIDE_H:
                    out.warn(where, "outside", f"自動で伸びる文字の箱がスライドの下にはみ出す見込み「{_short(it.text)}」")
            if it.x < -SLIDE_W * 0.01 or it.y < -SLIDE_H * 0.02 or it.x + it.w > SLIDE_W * 1.01 or it.y + it.h > SLIDE_H * 1.01:
                if it.ph_type not in ("TITLE",):    # このテンプレートのタイトル枠は上に少しはみ出している
                    out.warn(where, "outside", f"図形がスライドの外にはみ出している（{_short(it.text) or '文字なしの図形'}）")
        for i, it in enumerate(items):
            if not it.has_text or it.w * it.h == 0:
                continue
            for other in items[i + 1:]:
                if not other.opaque or other.is_line:
                    continue
                ov = _area_overlap(it, other)
                if ov / (it.w * it.h) > HIDDEN_RATIO:
                    if other.has_text and ov / (other.w * other.h) < 0.9:
                        continue
                    out.warn(where, "hidden", f"文字が後ろに置いた{'画像' if other.is_picture else '塗りの図形'}に隠れている「{_short(it.text)}」")
                    break
        for it in items:
            if it.has_text or it.is_picture or it.is_line or it.fill_rgb is None:
                continue
            geom = it.shape._element.find(P + "spPr/" + A + "prstGeom")
            if geom is not None and geom.get("prst") not in ("rect", "roundRect"):
                continue          # 矢印・三角・円などは図の部品
            under_text = any(o.has_text and _area_overlap(it, o) > 0.5 * o.w * o.h for o in items if o is not it)
            if not under_text:
                out.warn(where, "empty", f"文字の無い塗りの箱がある（x={it.x / 914400:.1f}in, y={it.y / 914400:.1f}in）。消し忘れでないか確認")

        # 下の空白
        if not (is_title or is_closing or is_toc):
            body = [it for it in items if it.ph_type not in ("TITLE", "CENTER_TITLE", "SLIDE_NUMBER", "DATE", "FOOTER")]
            if body:
                bottom = max(it.est_bottom for it in body)
                if bottom < FILL_MIN_BOTTOM:
                    out.warn(where, "fill",
                             f"本文が y={bottom / 914400:.1f}in で終わり、下の {(SLIDE_H - bottom) / SLIDE_H:.0%} が空白。"
                             "チップ＋パネル・タイル・帯などの型で面を作る（中身を水増ししない）")
            else:
                out.warn(where, "fill", "タイトル以外に何も無い")

        # 配色・太字の色
        check_colors(slide, where, allowed, out)
        check_bold_color(slide, where, out, allowed)

        # 残り物・文言
        texts = collect_texts(slide)
        dummy_hits = []
        for sh, t in texts:
            if _ph_type(sh) == "SLIDE_NUMBER":
                continue
            m = LEFTOVER_RE.search(t)
            if m:
                out.warn(where, "leftover", f"仮の文字「{m.group(0)}」が残っている「{_short(t)}」")
            elif t.strip() in dummies:
                dummy_hits.append(t.strip())
            wording.append((where, t))
        if dummy_hits:
            out.warn(where, "leftover", f"同梱テンプレートのダミー文が {len(dummy_hits)} 段落残っている"
                     f"「{_short(dummy_hits[0])}」{' ほか' if len(dummy_hits) > 1 else ''}")
        if title:
            if TITLE_POLITE_RE.search(title):
                out.warn(where, "title", f"タイトルが です・ます で終わっている「{title}」")
            if TITLE_ABOUT_RE.search(title):
                out.warn(where, "title", f"タイトルが「〜について／〜を示す」の形。何が分かったか・何をしたかを書く「{title}」")
    check_wording(wording, out, lexicon)
    return n


# --- 6. 構成案（Markdown）の文言 ---------------------------------------------
def check_markdown(path, out, lexicon):
    lines = pathlib.Path(path).read_text(encoding="utf-8").splitlines()
    in_code = in_comment = False
    texts = []
    for no, line in enumerate(lines, start=1):
        s = line.strip()
        if s.startswith("```"):
            in_code = not in_code
            continue
        if in_code:
            continue
        if "<!--" in s:
            in_comment = "-->" not in s.split("<!--", 1)[1]
            continue
        if in_comment:
            in_comment = "-->" not in s
            continue
        if not s or s.startswith("|---") or s.startswith("#"):
            continue          # 見出しは「## スライドN — タイトル」の書式なので見ない
        s = re.sub(r"<[^>]*>", "", s)            # <30字以内の結論> のような見本の穴は除く
        s = re.sub(r"`[^`]*`", "", s)
        s = re.sub(r"\$[^$]*\$", "", s)
        where = f"L{no}"
        m = re.match(r"[-*]\s*\*\*主張\*\*\s*[:：]\s*(.*)", s)
        if m:
            claim = m.group(1).strip()
            if claim:
                if text_em(claim) > 30:
                    out.warn(where, "claim", f"主張が全角30字を超えている（約{text_em(claim):.0f}字）「{_short(claim, 40)}」")
                if TITLE_POLITE_RE.search(claim):
                    out.warn(where, "claim", f"主張が です・ます で終わっている「{claim}」")
                if TITLE_ABOUT_RE.search(claim) or CLAIM_LABEL_RE.match(claim):
                    out.warn(where, "claim", f"主張が結論になっていない（ラベル・作業の説明）「{claim}」")
        texts.append((where, re.sub(r"\*\*|__", "", s)))
    check_wording(texts, out, lexicon, one_sentence=False)   # 構成案は指示文やノートに複数文があってよい


# --- main --------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("pptx", nargs="?", type=pathlib.Path)
    ap.add_argument("--base", type=pathlib.Path, help="ベースにした pptx（配色の基準。同梱 template.pptx ならダミー文の残りも調べる）")
    ap.add_argument("--skip", default="", help="検査しないスライド番号（1始まり、カンマ区切り。据え置きの研究背景など）。xml と min-font は常に検査する")
    ap.add_argument("--min-pt", type=float, default=16.0)
    ap.add_argument("--text", type=pathlib.Path, help="構成案の Markdown の文言だけを検査する")
    args = ap.parse_args(argv)

    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    out = Findings()
    lexicon = load_lexicon()
    n = 0
    if args.text:
        check_markdown(args.text, out, lexicon)
    elif args.pptx:
        if not args.pptx.exists():
            ap.error(f"not found: {args.pptx}")
        skip = {int(s) for s in re.split(r"[,\s]+", args.skip) if s.strip().isdigit()}
        if check_package(args.pptx, out):
            check_min_font(args.pptx, args.min_pt, out)
            n = check_slides(args.pptx, out, args.base, skip, lexicon)
    else:
        ap.error("pptx か --text を指定する")

    def key(it):
        w = it[0]
        m = re.match(r"(slide|L)(\d+)$", str(w))
        return (0 if w == "deck" else 1, int(m.group(2)) if m else 0, 0 if it[1] == "FAIL" else 1)
    for where, level, kind, msg in sorted(out.items, key=key):
        print(f"{where} [{level}] {kind}: {msg}")
    print(f"DONE: {out.count('FAIL')} FAIL / {out.count('WARN')} WARN / {n} slides")
    return 1 if out.count("FAIL") else 0


if __name__ == "__main__":
    sys.exit(main())
