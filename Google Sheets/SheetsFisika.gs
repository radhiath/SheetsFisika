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

