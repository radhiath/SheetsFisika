function hitungDeltaN(sigmaX, sigmaX2, n) {
  let hasil = (1 / n) * Math.sqrt((n * sigmaX2 - sigmaX ** 2) / (n - 1));
  if (isNaN(hasil)) {
    hasil = 0;
  }
  return hasil;
}

function hitungDelta3(sigmaX, sigmaX2) {
  let hasil = (1 / 3) * Math.sqrt((3 * sigmaX2 - sigmaX ** 2) / (3 - 1));
  if (isNaN(hasil)) {
    hasil = 0;
  }
  return hasil;
}

function hitungDelta5(sigmaX, sigmaX2) {
  let hasil = (1 / 5) * Math.sqrt((5 * sigmaX2 - sigmaX ** 2) / (5 - 1));
  if (isNaN(hasil)) {
    hasil = 0;
  }
  return hasil;
}

function hitungKSR(x, deltaX, angkaBlkgKoma=2) {
  let desimal = deltaX / x;
  let persen = (desimal * 100).toFixed(angkaBlkgKoma) + "%";
  
  if (desimal <= 0.001) persen += " (4 AP)";
  
  else if (desimal > 0.001 && desimal <= 0.01) persen += " (3 AP)";
  
  else if (desimal > 0.01 && desimal <= 0.1) persen += " (2 AP)";

  else persen += " (1 AP/ERROR)";

  return persen;
}

function hitungHasil(x, deltaX, KSR, style="default") {
  let formatX, formatDeltaX;
  let desimal = Number(KSR.split(" ")[1][1]) - 1;

  function toSuperScript(eksponen) {
    const superscriptMap = {
      "0": "⁰",
      "1": "¹",
      "2": "²",
      "3": "³",
      "4": "⁴",
      "5": "⁵",
      "6": "⁶",
      "7": "⁷",
      "8": "⁸",
      "9": "⁹",
      "-": "⁻"
    };
    return eksponen.split("").map((char) => superscriptMap[char] || char).join("");
  }

  formatX = String(x.toExponential(desimal)).replace("+", "").replace("e0", "");
  formatDeltaX = String(deltaX.toExponential(desimal)).replace("+", "").replace("e0", "");

  if (style === "default") {
    formatX = formatX.replace(/e(\-?\d+)/, (_, exp) => " ⋅ 10" + toSuperScript(exp));
    formatDeltaX = formatDeltaX.replace(/e(\-?\d+)/, (_, exp) => " ⋅ 10" + toSuperScript(exp));
    return "(" + formatX + " ± " + formatDeltaX + ")";

  } else if (style === "unicode") {
    expX = formatX.split("e")[1];
    expDeltaX = formatDeltaX.split("e")[1];
    formatX = formatX.replace(`e${expX}`, `\\bullet10^(${expX})`);
    formatDeltaX = formatDeltaX.replace(`e${expDeltaX}`, `\\bullet10^(${expDeltaX})`);
    return "(" + formatX + "±" + formatDeltaX + ")";
  
  } else if (style === "latex") {
    expX = formatX.split("e")[1];
    expDeltaX = formatDeltaX.split("e")[1];
    formatX = formatX.replace(`e${expX}`, `\\bullet10^{${expX}}`);
    formatDeltaX = formatDeltaX.replace(`e${expDeltaX}`, `\\bullet10^{${expDeltaX}}`);
    return "(" + formatX + " ± " + formatDeltaX + ")";
  }
}

function hitungAGrafik(n, sigmaX, sigmaY, sigmaX2, sigmaXY) {
  return ((sigmaY * sigmaX2) - (sigmaX * sigmaXY)) / ((n * sigmaX2) - (sigmaX ** 2)) ;
}

function hitungBGrafik(n, sigmaX, sigmaY, sigmaX2, sigmaXY) {
  return ((n * sigmaXY) - (sigmaX * sigmaY)) / ((n * sigmaX2) - (sigmaX ** 2)) ;
}

function hitungYGrafik(a, b, x) {
  return a + b * x;
}
