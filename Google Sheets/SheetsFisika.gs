function hitungDeltaN(sigmaX, sigma2X, n) {
  return (1 / n) * Math.sqrt((n * sigma2X - sigmaX ^ 2) / (n - 1));
}

function hitungDelta3(sigmaX, sigma2X) {
  return (1 / 3) * Math.sqrt((3 * sigma2X - sigmaX ^ 2) / (3 - 1));
}

function hitungDelta5(sigmaX, sigma2X) {
  return (1 / 5) * Math.sqrt((5 * sigma2X - sigmaX ^ 2) / (5 - 1));
}

function hitungKSR(rataRata, delta, angkaBlkgKoma=2) {
  let desimal = delta / rataRata;
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
  }
}
