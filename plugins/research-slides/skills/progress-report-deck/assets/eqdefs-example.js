// insert_math.js に渡す数式定義の例。名前: [OMML の中身, Fallback の平文]。
// 学習窓 D_t・探索半径 R・位置予測・状態ベクトル・カルマンフィルタの Q/R/P0 の書き方の見本。
module.exports = (M) => {
  const r = M.mrU(M.bold("r")), v = M.mrU(M.bold("v"));
  const sub = M.sSub, sup = M.sSup;
  const t1 = M.delim(M.mr("t") + M.mrU("+1")), rt = r + M.delim(M.mr("t"));
  return {
    D: [sub(M.mrU("\u{1D49F}"), M.mr("t")) + M.mrU("=") + M.delim(M.delim(sub(r, M.mr("i")) + M.mrU(",") + sub(v, M.mr("i"))) + M.mrU("\u2223") + M.mr("t") + M.mrU("\u2212") + M.mr("W") + M.mrU("\u2264") + sub(M.mr("f"), M.mr("i")) + M.mrU("\u2264") + M.mr("t") + M.mrU("\u22121"), "{", "}") + M.mrU(",\u2003") + M.mr("W") + M.mrU("=15"), "D_t = {(r_i, v_i) | t−W ≤ f_i ≤ t−1},  W = 15"],
    R1: [M.mr("R") + M.mrU("=") + M.mrU("min") + M.delim(sub(M.mr("R"), M.mrU("base")) + M.mrU("+") + M.mr("α") + sub(M.mr("U"), M.mrU("norm")) + M.mrU(",\u2009") + sub(M.mr("R"), M.mrU("max"))), "R = min(R_base + α U_norm, R_max)"],
    R2: [sub(M.mr("R"), M.mrU("base")) + M.mrU("=6,\u2003") + M.mr("α") + M.mrU("=10,\u2003") + sub(M.mr("R"), M.mrU("max")) + M.mrU("=10"), "R_base = 6,  α = 10,  R_max = 10"],
    POS: [M.hat(r) + t1 + M.mrU("=") + rt + M.mrU("+") + sub(M.hat(v), M.mrU("pred")) + M.delim(rt), "r̂(t+1) = r(t) + v̂_pred(r(t))"],
    LIN: [M.hat(r) + t1 + M.mrU("=") + rt + M.mrU("+") + v + M.delim(M.mr("t")), "r̂(t+1) = r(t) + v(t)"],
    ST: [sup(M.delim(M.mr("y") + M.mrU(",") + M.mr("x") + M.mrU(",") + sub(M.mr("v"), M.mr("y")) + M.mrU(",") + sub(M.mr("v"), M.mr("x")), "[", "]"), M.mrU("T")), "[y, x, v_y, v_x]^T"],
    KF1: [M.mr("Q") + M.mrU("=") + sup(M.mrU("0.5"), M.mrU("2")) + sub(M.mr("I"), M.mrU("4")) + M.mrU(",\u2003") + M.mr("R") + M.mrU("=") + sup(M.mrU("1.0"), M.mrU("2")) + sub(M.mr("I"), M.mrU("2")), "Q = 0.5² I_4,  R = 1.0² I_2"],
    KF2: [sub(M.mr("P"), M.mrU("0")) + M.mrU("=diag") + M.delim(sup(M.mrU("1"), M.mrU("2")) + M.mrU(",") + sup(M.mrU("1"), M.mrU("2")) + M.mrU(",") + sup(M.mrU("5"), M.mrU("2")) + M.mrU(",") + sup(M.mrU("5"), M.mrU("2"))), "P_0 = diag(1², 1², 5², 5²)"],
  };
};
