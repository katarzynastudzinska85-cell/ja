// Testy funkcji bez przeglądarki: node --test
const test = require("node:test");
const assert = require("node:assert");
const agent = require("../live_diagnostic_agent.js");

test("parseUsDateTime uwzględnia AM/PM", () => {
  assert.strictEqual(agent.parseUsDateTime("9/24/2026 6:53:00 PM").getHours(), 18);
  assert.strictEqual(agent.parseUsDateTime("9/24/2026 12:05:00 AM").getHours(), 0);
  assert.strictEqual(agent.parseUsDateTime("9/24/2026 12:05:00 PM").getHours(), 12);
  assert.strictEqual(agent.parseUsDateTime("9/29/2026").getDate(), 29);
});

test("titleJnoCheck: zgodny, błąd OCR o 1 znak, inny aparat", () => {
  assert.strictEqual(agent.titleJnoCheck("Duty Cycle -- J46001812", "46001812").ok, true);
  const fuzzy = agent.titleJnoCheck("Step Loss -- J24000123", "34000123");
  assert.strictEqual(fuzzy.ok, true);
  assert.strictEqual(fuzzy.fuzzy, true);
  assert.strictEqual(agent.titleJnoCheck("Duty Cycle -- J46009999", "46001812").ok, false);
  assert.strictEqual(agent.titleJnoCheck("Duty Cycle", "46001812").ok, null);
});

test("chartKind toleruje błędy OCR", () => {
  assert.strictEqual(agent.chartKind("Supply 3 Buty Cycle"), "duty");
  assert.strictEqual(agent.chartKind("Thermistor Difference"), "thermDiff");
  assert.strictEqual(agent.chartKind("Supply 3 Thermistors"), "thermistors");
});

const base = {
  jno: "46002000", model: "VITROS 4600", status: "ORANGE", aaaVerified: true,
  dateRange: { start: "9/29/2026 12:00:00 AM", end: "10/6/2026 12:00:00 AM" },
  location: { customer: "X", address: "", city: "" }, asset: null,
  chartFiles: [], chartData: [], serviceDocuments: [],
  genericEvidence: { rules: [], conditionCodes: [], linkLabels: [] }, error: ""
};

test("Data Logger: priorytet ZDALNIE i czas bez łączności", () => {
  const days = ["D6", "D5", "D4", "D3", "D2", "D1", "Today"].map(day => ({ day, value: "0/0" }));
  const a = agent.analyzeAlert({
    ...base, alertName: "Data Logger Status",
    dataLogger: { detected: true, lastConnected: "9/24/2026 6:53:00 PM", aFile: days, bFile: days }
  }, new Date(2026, 9, 6, 12, 0));
  assert.strictEqual(a.priority, "ZDALNIE");
  assert.strictEqual(a.elapsedText, "≈ 11 d 17 h");
  assert.ok(a.observations.some(o => /brak transmisji/.test(o.text)));
});

test("Błąd AAA: brak oceny technicznej, niezależnie od koloru", () => {
  const a = agent.analyzeAlert({ ...base, alertName: "Supply 3 Thermal", error: "AAA nie otworzyło strony." });
  assert.strictEqual(a.priority, "ZDALNIE");
  assert.match(a.priorityReason, /Brak oceny technicznej/);
});

test("Kolor alertu nie wyznacza priorytetu", () => {
  const a = agent.analyzeAlert({ ...base, alertName: "Supply 3 Thermal", dataLogger: {} });
  assert.notStrictEqual(a.priority, "POMARAŃCZOWY");
  assert.ok(["WYSOKI", "ŚREDNI", "ZDALNIE", "NISKI"].includes(a.priority));
});

test("Treść AAA: wyróżniony wiersz tabeli staje się faktem, nagłówek z kolorowego wiersza", () => {
  const blocks = agent.sanitizeBlocks([
    { type: "heading", level: 3, text: "Processed Troubleshooting Guidance -- 76000981" },
    { type: "table", caption: "", rows: [
      { cells: ["Guidance (NEW)", "Noise %"], header: false, bg: "rgb(51, 68, 204)" },
      { cells: ["E: Preliminary Well Wash aspirate dispense actual volume", "8.33"], header: false, bg: "rgb(255, 255, 0)" },
      { cells: ["F: Preliminary Well Wash level sense", "3.45"], header: false, bg: "" }
    ] },
    { type: "para", text: "Save" },
    { type: "para", text: "Serwer 10.1.2.3" }
  ]);
  assert.strictEqual(blocks.length, 3, "samo słowo przycisku usunięte");
  assert.strictEqual(blocks[1].rows[0].header, true);
  assert.strictEqual(blocks[1].rows[1].highlight, "żółte");
  assert.match(blocks[2].text, /\[IP\]/);
  const facts = agent.aaaContentFacts({ aaaContent: [{ source: "strona AAA", blocks }] });
  assert.strictEqual(facts.length, 1);
  assert.match(facts[0], /Noise %: 8.33/);
});

test("Czerwony tekst w tabeli AAA staje się faktem", () => {
  const blocks = agent.sanitizeBlocks([
    { type: "table", caption: "", rows: [
      { cells: ["Guidance", "Noise %"], header: true, bg: "" },
      { cells: ["D: Final Well Wash", "12.5"], header: false, bg: "", fg: "rgb(255, 0, 0)" }
    ] },
    { type: "para", text: "Investigate: Signal Reagent", red: true }
  ]);
  const facts = agent.aaaContentFacts({ aaaContent: [{ source: "strona AAA", blocks }] });
  assert.strictEqual(facts.length, 2);
  assert.match(facts[0], /czerwony tekst/);
  assert.match(facts[1], /Investigate: Signal Reagent/);
});
