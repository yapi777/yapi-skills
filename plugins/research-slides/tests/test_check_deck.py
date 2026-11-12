"""check_deck.py のテスト。

    python -m unittest discover -s plugins/research-slides/tests

fixture はリポジトリに置かず、同梱の template.pptx を python-pptx でその場で書き換えて作る。
検査ごとに「壊した版で指摘が出る／直した版（または元の版）で出ない」を確かめる。
"""
import io
import pathlib
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout

SKILL = pathlib.Path(__file__).resolve().parents[1] / "skills" / "progress-report-deck"
sys.path.insert(0, str(SKILL / "scripts"))

import check_deck  # noqa: E402
from pptx import Presentation  # noqa: E402
from pptx.dml.color import RGBColor  # noqa: E402
from pptx.util import Emu, Pt  # noqa: E402

TEMPLATE = SKILL / "assets" / "template.pptx"


def run(*args):
    buf = io.StringIO()
    with redirect_stdout(buf):
        code = check_deck.main([str(a) for a in args])
    lines = buf.getvalue().splitlines()
    return code, lines


def hits(lines, kind, where=None):
    out = []
    for ln in lines:
        m = re.match(r"(\S+) \[(FAIL|WARN)\] ([\w-]+): ", ln)
        if m and m.group(3) == kind and (where is None or m.group(1) == where):
            out.append(ln)
    return out


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = pathlib.Path(tempfile.mkdtemp())

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def deck(self, name, edit):
        """template.pptx を開いて edit(prs) を当てて保存したパスを返す。"""
        prs = Presentation(TEMPLATE)
        edit(prs)
        path = self.tmp / f"{name}.pptx"
        prs.save(path)
        return path

    def patched(self, name, part, fn):
        """zip 内の1パーツの XML 文字列を fn で書き換えたコピーを返す。"""
        src = zipfile.ZipFile(TEMPLATE)
        path = self.tmp / f"{name}.pptx"
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == part:
                    data = fn(data.decode("utf-8")).encode("utf-8")
                dst.writestr(info, data)
        return path


def shape_with(slide, needle):
    for sh in slide.shapes:
        if sh.has_text_frame and needle in sh.text_frame.text:
            return sh
    raise AssertionError(f"{needle} が見つからない")


def set_text(shape, text):
    p = shape.text_frame.paragraphs[0]
    for extra in list(shape.text_frame.paragraphs)[1:]:
        extra._p.getparent().remove(extra._p)
    runs = p.runs
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)


class TestTemplate(Base):
    def test_template_has_no_fail_and_no_geometry_false_alarms(self):
        code, lines = run(TEMPLATE, "--base", TEMPLATE, "--skip", "3,4")
        self.assertEqual(code, 0, "\n".join(lines))
        self.assertTrue(lines[-1].startswith("DONE: 0 FAIL"), lines[-1])
        for kind in ("overflow", "outside", "hidden", "empty", "fill", "bold-color", "xml", "color"):
            self.assertEqual(hits(lines, kind), [], kind)

    def test_done_line_always_printed(self):
        _, lines = run(TEMPLATE)
        self.assertRegex(lines[-1], r"^DONE: \d+ FAIL / \d+ WARN / 18 slides$")


class TestPackage(Base):
    def test_dangling_image_rel_is_fail(self):
        def drop_image(xml):
            return re.sub(r'<Relationship [^>]*Type="[^"]*/image"[^>]*/>', "", xml)
        path = self.patched("dangling", "ppt/slides/_rels/slide3.xml.rels", drop_image)
        code, lines = run(path)
        self.assertEqual(code, 1)
        self.assertTrue(any("切れた参照" in ln for ln in hits(lines, "xml", "slide3")), lines)

    def test_bad_scheme_color_is_fail(self):
        path = self.patched("scheme", "ppt/slides/slide7.xml",
                            lambda x: x.replace('schemeClr val="accent1"', 'schemeClr val="FFC000"', 1))
        code, lines = run(path)
        self.assertEqual(code, 1)
        self.assertTrue(hits(lines, "xml", "slide7"), lines)

    def test_slide_removed_from_list_but_part_left_is_fail(self):
        def drop_last(xml):
            ids = re.findall(r"<p:sldId [^>]*/>", xml)
            return xml.replace(ids[-2], "")
        path = self.patched("orphanpart", "ppt/presentation.xml", drop_last)
        code, lines = run(path)
        self.assertEqual(code, 1)
        self.assertTrue(any("sldIdLst に無い" in ln for ln in hits(lines, "xml")), lines)


class TestMinFont(Base):
    def test_12pt_run_is_fail(self):
        def edit(prs):
            sh = shape_with(prs.slides[6], "可視化")
            sh.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
        code, lines = run(self.deck("small", edit))
        self.assertEqual(code, 1)
        self.assertTrue(hits(lines, "min-font", "slide7"), lines)

    def test_min_pt_option(self):
        code, lines = run(TEMPLATE, "--min-pt", "18")
        self.assertEqual(code, 1)
        self.assertTrue(hits(lines, "min-font"))


class TestOrphanOverflow(Base):
    # slide7 の「指摘」パネル: 幅 6954000 − 余白 360000 → 18pt で1行 28 字
    def test_one_or_two_chars_on_last_line(self):
        def edit(prs):
            set_text(shape_with(prs.slides[6], "実験結果について"), "あ" * 30)
        _, lines = run(self.deck("orphan", edit))
        self.assertTrue(hits(lines, "orphan", "slide7"), lines)

    def test_fits_on_one_line(self):
        def edit(prs):
            set_text(shape_with(prs.slides[6], "実験結果について"), "あ" * 28)
        _, lines = run(self.deck("fits", edit))
        self.assertEqual(hits(lines, "orphan", "slide7"), [])

    def test_text_too_long_for_box(self):
        def edit(prs):
            set_text(shape_with(prs.slides[6], "実験結果について"), "あ" * 28 * 4)
        _, lines = run(self.deck("overflow", edit))
        self.assertTrue(hits(lines, "overflow", "slide7"), lines)

    def test_wrap_widths_kinsoku_and_words(self):
        # 行頭に「。」が来るときは前の字ごと次の行へ送る
        self.assertEqual([round(w) for w in check_deck.wrap_widths("あいうえ。", 4)], [3, 2])
        # 英単語は途中で切らない
        widths = check_deck.wrap_widths("ab abcdefgh", 3)
        self.assertGreaterEqual(len(widths), 2)


class TestFill(Base):
    def test_bottom_half_empty(self):
        def edit(prs):
            s = prs.slides[13]          # 主要な結果: 表＋帯3本
            for sh in list(s.shapes):
                if sh.has_text_frame and sh.text_frame.text.startswith(("指標", "評価条件")):
                    sh._element.getparent().remove(sh._element)
        _, lines = run(self.deck("fill", edit))
        self.assertTrue(hits(lines, "fill", "slide14"), lines)

    def test_template_slide_is_filled(self):
        _, lines = run(TEMPLATE)
        self.assertEqual(hits(lines, "fill", "slide14"), [])


class TestShapes(Base):
    def test_empty_filled_box(self):
        def edit(prs):
            from pptx.enum.shapes import MSO_SHAPE
            box = prs.slides[13].shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(450000), Emu(6450000), Emu(2000000), Emu(300000))
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(0xDB, 0xE7, 0xF3)
        _, lines = run(self.deck("empty", edit))
        self.assertTrue(hits(lines, "empty", "slide14"), lines)

    def test_text_hidden_behind_later_box(self):
        def edit(prs):
            from pptx.enum.shapes import MSO_SHAPE
            sh = shape_with(prs.slides[6], "実験結果について")
            box = prs.slides[6].shapes.add_shape(MSO_SHAPE.RECTANGLE, sh.left, sh.top, sh.width, sh.height)
            box.fill.solid()
            box.fill.fore_color.rgb = RGBColor(0xF0, 0xF0, 0xF0)
        _, lines = run(self.deck("hidden", edit))
        self.assertTrue(hits(lines, "hidden", "slide7"), lines)

    def test_outside_slide(self):
        def edit(prs):
            sh = shape_with(prs.slides[6], "実験結果について")
            sh.left = Emu(8000000)
        _, lines = run(self.deck("outside", edit))
        self.assertTrue(hits(lines, "outside", "slide7"), lines)


class TestColor(Base):
    def test_red_text_is_off_palette(self):
        def edit(prs):
            sh = shape_with(prs.slides[6], "実験結果について")
            sh.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xC0, 0x00, 0x00)
        _, lines = run(self.deck("red", edit))
        self.assertTrue(any("C00000" in ln for ln in hits(lines, "color", "slide7")), lines)

    def test_bold_colored(self):
        def edit(prs):
            r = shape_with(prs.slides[6], "実験結果について").text_frame.paragraphs[0].runs[0]
            r.font.bold = True
            r.font.color.rgb = RGBColor(0x15, 0x60, 0x82)
        _, lines = run(self.deck("boldcolor", edit))
        self.assertTrue(hits(lines, "bold-color", "slide7"), lines)
        self.assertEqual(hits(lines, "color", "slide7"), [])      # 色自体はパレット内


class TestLeftoverAndWording(Base):
    def test_placeholder_marks(self):
        def edit(prs):
            set_text(shape_with(prs.slides[6], "実験結果について"), "〇〇の結果をここに書く")
        _, lines = run(self.deck("leftover", edit))
        self.assertTrue(hits(lines, "leftover", "slide7"), lines)

    def test_template_dummy_only_with_bundled_base(self):
        _, with_base = run(TEMPLATE, "--base", TEMPLATE)
        _, without = run(TEMPLATE)
        self.assertTrue(any("ダミー文" in ln for ln in hits(with_base, "leftover", "slide14")))
        self.assertFalse(any("ダミー文" in ln for ln in hits(without, "leftover")))

    def test_ai_smell_word(self):
        def edit(prs):
            set_text(shape_with(prs.slides[6], "実験結果について"), "精度が飛躍的に向上")
        _, lines = run(self.deck("smell", edit))
        self.assertTrue(any("飛躍的" in ln for ln in hits(lines, "wording", "slide7")), lines)

    def test_polite_title(self):
        def edit(prs):
            prs.slides[6].shapes.title.text_frame.text = "結果について説明します"
        _, lines = run(self.deck("title", edit))
        self.assertTrue(hits(lines, "title", "slide7"), lines)


class TestMarkdown(Base):
    def md(self, body):
        p = self.tmp / "doc.md"
        p.write_text(body, encoding="utf-8")
        return p

    def test_claim_rules(self):
        p = self.md("## スライド6 — 結果\n\n- **主張**: " + "あ" * 31 + "\n"
                    "## スライド7 — 結果\n\n- **主張**: 結果①\n"
                    "## スライド8 — 結果\n\n- **主張**: 速度が上がるほど差が開く\n")
        code, lines = run("--text", p)
        self.assertEqual(code, 0)
        self.assertTrue(hits(lines, "claim", "L3"))      # 31字
        self.assertTrue(hits(lines, "claim", "L6"))      # ラベルだけ
        self.assertEqual(hits(lines, "claim", "L9"), [])  # 結論になっている

    def test_bundled_doc_template_is_clean(self):
        tpl = SKILL.parent / "create-doc" / "references" / "template.md"
        _, lines = run("--text", tpl)
        self.assertEqual([ln for ln in lines if "[" in ln], [], lines)


if __name__ == "__main__":
    unittest.main()
