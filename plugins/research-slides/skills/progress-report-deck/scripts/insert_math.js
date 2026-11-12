/**
 * python-pptx で作ったスライドの「@@EQ:名前@@」という仮テキストを OMML 数式に置き換える。
 *
 * 流れ: python-pptx でパネル等を作るときに、数式を入れたい段落に
 *       ("@@EQ:D@@", False, None) のようなランを1つ置いておく（文字サイズ・色はそのランの書式を使う）
 *       → 保存 → unzip → このスクリプト → re-zip。
 * 同じ段落に普通の文字のランと並べてもよい（例: 「状態 」＋@@EQ:ST@@）。
 * 1つの図形に複数のマーカーがあっても、図形ごと mc:AlternateContent で包む（Fallback はプレーンテキスト）。
 *
 * 使い方（数式定義ファイルを書いて渡す）:
 *   node insert_math.js <unpacked-dir> <eqdefs.js>
 *
 * eqdefs.js の例:
 *   module.exports = (M) => {
 *     const r = M.mrU(M.bold("r")), v = M.mrU(M.bold("v"));
 *     return {
 *       LIN: [ M.hat(r) + M.delim(M.mr("t") + M.mrU("+1")) + M.mrU("=") + r + M.delim(M.mr("t"))
 *              + M.mrU("+") + v + M.delim(M.mr("t")),
 *              "r̂(t+1) = r(t) + v(t)" ],          // [OMMLの中身, Fallbackの平文]
 *     };
 *   };
 */
const fs = require("fs"), path = require("path");
const { M, wrapShapeWithMath } = require(path.join(__dirname, "omml.js"));

const [root, defsPath] = process.argv.slice(2);
if (!root || !defsPath) { console.error("usage: node insert_math.js <unpacked-dir> <eqdefs.js>"); process.exit(1); }
const EQ = require(path.resolve(defsPath))(M);
const S = path.join(root, "ppt", "slides");
let total = 0;
for (const f of fs.readdirSync(S).filter((f) => /^slide\d+\.xml$/.test(f))) {
  const p = path.join(S, f);
  let xml = fs.readFileSync(p, "utf8");
  if (!xml.includes("@@EQ:")) continue;
  const ids = new Set();
  const re = /@@EQ:(\w+)@@/g; let m;
  while ((m = re.exec(xml))) {
    const at = xml.lastIndexOf('<p:cNvPr id="', m.index);
    ids.add(xml.slice(at).match(/<p:cNvPr id="(\d+)"/)[1]);
  }
  for (const id of ids) {
    const at = xml.indexOf(`<p:cNvPr id="${id}"`);
    const s0 = xml.lastIndexOf("<p:sp>", at), s1 = xml.indexOf("</p:sp>", at);
    const body = xml.slice(s0, s1).match(/<p:txBody>[\s\S]*<\/p:txBody>/)[0];
    xml = wrapShapeWithMath(xml, +id, (useMath) =>
      body.replace(/<a:r>((?:(?!<a:r>)[\s\S])*?)<a:t>@@EQ:(\w+)@@<\/a:t><\/a:r>/g, (all, rpr, name) => {
        const e = EQ[name];
        if (!e) throw new Error(`eqdefs に ${name} がない`);
        if (useMath) total++;
        const sz = (rpr.match(/sz="(\d+)"/) || [])[1];
        let eq = M.oMath(e[0]);
        if (sz) eq = eq.replace(/sz="1800"/g, `sz="${sz}"`);   // ランの文字サイズに合わせる
        return useMath ? eq : `<a:r>${rpr}<a:t>${e[1]}</a:t></a:r>`;
      }));
  }
  fs.writeFileSync(p, xml);
  console.log(f, "shapes:", [...ids].join(","));
}
console.log("equations inserted:", total);
