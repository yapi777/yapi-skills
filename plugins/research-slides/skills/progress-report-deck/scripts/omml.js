/**
 * PowerPoint の「ちゃんとした数式」（OMML）を pptx の生XMLに埋め込むためのヘルパ。
 *
 * ポイント（実際に PowerPoint で検証済み）:
 *  - DrawingML の中の数式は <a14:m><m:oMath>…</m:oMath></a14:m>。
 *  - m:r の書式は w:rPr ではなく **a:rPr** で指定し、フォントは Cambria Math。
 *  - a:rPr には **必ず文字色を明示する**。省略すると図形既定（p:style の fontRef が
 *    lt1＝白）が効いて、数式だけ極端に薄く表示される。
 *  - a14:m を a:p の直下にそのまま置くと、LibreOffice など a14 を解さないビューアでは
 *    数式が丸ごと消える。必ず p:sp 全体を mc:AlternateContent で包み、
 *    mc:Fallback にプレーンテキスト版を入れておく（PowerPoint は Choice 側を読む）。
 *  - この形なら pptx skill の validate.py（XSD 検証）も PASS する。
 *
 * 使い方:
 *   const { M, wrapShapeWithMath } = require(".../omml.js");
 *   const eq = M.oMath(M.sSub(M.mr("r"), M.mr("p")) + M.delim(M.mr("t") + M.mrU("+1"), "(", ")"));
 *   xml = wrapShapeWithMath(xml, 8, (useMath) => `<p:txBody>…${useMath ? eq : "<a:t>r_p(t)</a:t>"}…</p:txBody>`);
 */

const DEFAULT_SZ = 1800; // 16pt 未満にしないこと
const COLOR = "1A1A1A";

function rPr(sz, italic, color) {
  return (
    `<a:rPr lang="en-US" sz="${sz}" i="${italic ? 1 : 0}" smtClean="0">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>` +
    `<a:latin typeface="Cambria Math" panose="02040503050406030204" pitchFamily="18" charset="0"/>` +
    `</a:rPr>`
  );
}

const M = {
  /** 斜体のラン（変数: x, r, σ など） */
  mr(t, sz = DEFAULT_SZ, color = COLOR) {
    return `<m:r>${rPr(sz, true, color)}<m:t>${t}</m:t></m:r>`;
  },
  /** 立体のラン（演算子・数字: =, +, 2, T など） */
  mrU(t, sz = DEFAULT_SZ, color = COLOR) {
    return `<m:r><m:rPr><m:sty m:val="p"/></m:rPr>${rPr(sz, false, color)}<m:t>${t}</m:t></m:r>`;
  },
  ctrl(sz = DEFAULT_SZ, color = COLOR) {
    return `<m:ctrlPr>${rPr(sz, true, color)}</m:ctrlPr>`;
  },
  /** 下付き: base_sub */
  sSub(e, sub, sz = DEFAULT_SZ) {
    return `<m:sSub><m:sSubPr>${M.ctrl(sz)}</m:sSubPr><m:e>${e}</m:e><m:sub>${sub}</m:sub></m:sSub>`;
  },
  /** 上付き: base^sup */
  sSup(e, sup, sz = DEFAULT_SZ) {
    return `<m:sSup><m:sSupPr>${M.ctrl(sz)}</m:sSupPr><m:e>${e}</m:e><m:sup>${sup}</m:sup></m:sSup>`;
  },
  /** 上下付き: base_sub^sup（σ_q² など） */
  sSubSup(e, sub, sup, sz = DEFAULT_SZ) {
    return (
      `<m:sSubSup><m:sSubSupPr>${M.ctrl(sz)}</m:sSubSupPr><m:e>${e}</m:e>` +
      `<m:sub>${sub}</m:sub><m:sup>${sup}</m:sup></m:sSubSup>`
    );
  },
  /** 括弧: (…) [ … ] など。大きさは中身に自動追従する */
  delim(inner, beg = "(", end = ")", sz = DEFAULT_SZ) {
    return (
      `<m:d><m:dPr><m:begChr m:val="${beg}"/><m:endChr m:val="${end}"/>${M.ctrl(sz)}</m:dPr>` +
      `<m:e>${inner}</m:e></m:d>`
    );
  },
  /** 分数 */
  frac(num, den, sz = DEFAULT_SZ) {
    return `<m:f><m:fPr>${M.ctrl(sz)}</m:fPr><m:num>${num}</m:num><m:den>${den}</m:den></m:f>`;
  },
  /** アクセント（ハット等）: r̂ → M.hat(M.mrU(M.bold("r"))) */
  hat(e, chr = "̂", sz = DEFAULT_SZ) {
    return `<m:acc><m:accPr><m:chr m:val="${chr}"/>${M.ctrl(sz)}</m:accPr><m:e>${e}</m:e></m:acc>`;
  },
  /** ベクトル用の数学太字（Cambria Math の U+1D400 系）。M.mrU(M.bold("r")) のように立体で使う */
  bold(ch) {
    const c = ch.charCodeAt(0);
    if (c >= 97 && c <= 122) return String.fromCodePoint(0x1d41a + c - 97);
    if (c >= 65 && c <= 90) return String.fromCodePoint(0x1d400 + c - 65);
    return ch;
  },
  /** 全体を数式として包む。a:p の中に置く */
  oMath(inner) {
    return (
      `<a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
      inner +
      `</m:oMath></a14:m>`
    );
  },
};

/**
 * 指定 shape の txBody を差し替え、p:sp 全体を mc:AlternateContent で包む。
 *
 * @param {string} xml       スライドXML全体
 * @param {number} shapeId   <p:cNvPr id="N"> の N
 * @param {(useMath:boolean)=>string} bodyFn  true→数式入り / false→プレーンテキストの <p:txBody>
 * @param {number} [newHeight]  図形の高さ(EMU)を同時に変えたいとき
 */
function wrapShapeWithMath(xml, shapeId, bodyFn, newHeight) {
  const marker = `<p:cNvPr id="${shapeId}"`;
  const at = xml.indexOf(marker);
  if (at === -1) throw new Error(`shape ${shapeId} not found`);
  const spStart = xml.lastIndexOf("<p:sp>", at);
  const spEnd = xml.indexOf("</p:sp>", at) + "</p:sp>".length;
  let sp = xml.slice(spStart, spEnd);

  if (newHeight) {
    sp = sp.replace(/(<a:ext cx="\d+" cy=")\d+(")/, `$1${newHeight}$2`);
  }
  const skeleton = sp.replace(/<p:txBody>[\s\S]*<\/p:txBody>/, "@@BODY@@");

  const alt =
    `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">` +
    `<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">` +
    skeleton.replace("@@BODY@@", bodyFn(true)) +
    `</mc:Choice><mc:Fallback>` +
    skeleton.replace("@@BODY@@", bodyFn(false)) +
    `</mc:Fallback></mc:AlternateContent>`;

  return xml.slice(0, spStart) + alt + xml.slice(spEnd);
}

module.exports = { M, wrapShapeWithMath, DEFAULT_SZ };
