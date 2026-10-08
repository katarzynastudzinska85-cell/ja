// Wyszukuje aparat i klienta po numerze J w bazie Excel (lokalnie, bez internetu).
// Użycie: node znajdz.js 76000978 [kolejne numery...]
const agent = require("./live_diagnostic_agent.js");

const numbers = process.argv.slice(2);
if (!numbers.length) {
  console.log("Użycie: node znajdz.js 76000978 [kolejne numery...]");
  process.exit(1);
}

let db;
try {
  db = agent.loadAnalyzerDatabase();
} catch (error) {
  console.log(error.message);
  process.exit(1);
}

for (const jno of numbers) {
  const asset = agent.findAnalyzer(db, jno);
  console.log("");
  if (!asset) {
    console.log(`J${String(jno).replace(/^J/i, "")}: brak w bazie ${require("path").basename(agent.EXCEL_FILE)}.`);
    continue;
  }
  const loc = agent.getLocation(asset);
  console.log(`${asset.jno} — ${agent.getAnalyzerModel(asset)}`);
  console.log(`  Klient:        ${loc.customer}${asset.customerNumber ? ` (nr ${asset.customerNumber})` : ""}`);
  console.log(`  Adres:         ${loc.address}`);
  console.log(`  Miasto:        ${loc.city}`);
  console.log(`  Status:        ${asset.excelStatus || "brak danych"}`);
  console.log(`  Software:      ${asset.softwareVersion || "brak danych"}`);
  console.log(`  Numer seryjny: ${asset.serialNumber || "brak danych"}`);
  console.log(`  Instalacja:    ${asset.installDate || "brak danych"}`);
}
