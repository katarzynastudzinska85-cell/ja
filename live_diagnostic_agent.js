// ============================================================
// QUIDELORTHO e-CONNECTIVITY SERVICE TRIAGE AGENT v3
// EXCEL ASSET DATABASE + AAA + POLSKI RAPORT FE
// READ-ONLY
//
// ŹRÓDŁA:
// 1. Dashboard -> aktywny alert + J-number
// 2. Excel -> model + placówka + adres + software + status
// 3. AAA -> diagnostyka + wykresy + historia
//
// WORKFLOW:
// - uruchamiasz skrypt
// - logujesz się ręcznie
// - przechodzisz wymagane podstrony
// - dochodzisz do dashboardu
// - wracasz do PowerShell
// - ENTER
// - dopiero wtedy agent analizuje
//
// READ-ONLY:
// - brak Save
// - brak Apply
// - brak zmian konfiguracji
// - brak adjustmentów
// - brak resetów
// - brak service commands
// ============================================================

const { chromium } = require("playwright");
const PDFDocument = require("pdfkit");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const readline = require("readline");

// Polskie znaki w konsoli Windows (PowerShell / cmd): przełącz stronę kodową na UTF-8.
try {
  if (process.platform === "win32") {
    require("child_process").execSync("chcp 65001 >nul", {
      stdio: "ignore",
      shell: true
    });
  }
  if (process.stdout && process.stdout.setDefaultEncoding) {
    process.stdout.setDefaultEncoding("utf8");
  }
} catch {}

// ============================================================
// CONFIG
// ============================================================

const START_URL =
  "https://orthoplus.orthoclinicaldiagnostics.com/eConnectivityDashboard/App/EHealthCheck.aspx";

const EDGE_PATH =
  "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe";

const EXCEL_FILE =
  path.join(
    __dirname,
    "stan Ortho na 23.04.2026.xlsx"
  );

const SESSION_FILE =
  path.join(
    __dirname,
    "session.json"
  );

const OUTPUT_DIR =
  path.join(
    __dirname,
    "aaa_output_live"
  );

const PDF_FILE =
  path.join(
    OUTPUT_DIR,
    "Raport_Serwisowy_eConnectivity_LIVE_V3.pdf"
  );

const JSON_FILE =
  path.join(
    OUTPUT_DIR,
    "Service_Triage_Data_LIVE_V3.json"
  );

const KNOWLEDGE_FILE = path.join(__dirname, "service_knowledge.json");

const ARIAL =
  "C:\\Windows\\Fonts\\arial.ttf";

const ARIAL_BOLD =
  "C:\\Windows\\Fonts\\arialbd.ttf";

// ============================================================
// BASIC HELPERS
// ============================================================

function ensureDir(dir) {
  fs.mkdirSync(dir, {
    recursive: true
  });
}

function sleep(ms) {
  return new Promise(resolve =>
    setTimeout(resolve, ms)
  );
}


// Tablica odwrotna Windows-1250: znak -> bajt (budowana z wbudowanego TextDecoder).
const CP1250_REVERSE = (() => {
  const map = new Map();
  try {
    const dec = new TextDecoder("windows-1250");
    for (let b = 0x80; b <= 0xff; b++) {
      map.set(dec.decode(Uint8Array.of(b)), b);
    }
  } catch {}
  return map;
})();

function cp1250RunToUtf8(run) {
  const bytes = [];
  for (const ch of run) {
    const code = ch.codePointAt(0);
    if (code < 0x80) {
      bytes.push(code);
    } else if (CP1250_REVERSE.has(ch)) {
      bytes.push(CP1250_REVERSE.get(ch));
    } else if (code >= 0x80 && code <= 0x9f) {
      bytes.push(code); // bajty niezdefiniowane w cp1250 (np. 0x81 po "Ĺ")
    } else {
      return null;
    }
  }
  const strict = new TextDecoder("utf-8", { fatal: true });

  try {
    return strict.decode(Uint8Array.from(bytes));
  } catch {}

  // Po kopiowaniu tekstu często znikają niewidzialne znaki 0x81 (część "Ł") i 0x98 (część "Ę").
  // Odtwórz je tylko wtedy, gdy po bajcie wiodącym nie ma bajtu kontynuacji.
  const healed = [];
  for (let i = 0; i < bytes.length; i++) {
    healed.push(bytes[i]);
    const next = bytes[i + 1];
    const missingNext = next === undefined || next < 0x80 || next > 0xbf;
    if (missingNext && bytes[i] === 0xc5) healed.push(0x81);
    if (missingNext && bytes[i] === 0xc4) healed.push(0x98);
  }

  try {
    return strict.decode(Uint8Array.from(healed));
  } catch {
    return null;
  }
}

// Naprawia tekst, w którym UTF-8 został omyłkowo odczytany jako Windows-1250
// (np. "ĹşrĂłdĹ‚o" -> "źródło"). Poprawne polskie znaki zostają nietknięte.
function repairPolishText(value) {
  let text = String(value ?? "");

  if (!/[^\x00-\x7F]/.test(text)) {
    return text;
  }

  for (let pass = 0; pass < 3; pass++) {
    let changed = false;

    text = text.replace(/[^\x00-\x7F]+/g, run => {
      const fixed = cp1250RunToUtf8(run);
      if (fixed !== null && fixed !== run && !fixed.includes("\uFFFD")) {
        changed = true;
        return fixed;
      }
      return run;
    });

    if (!changed) break;
  }

  return text;
}

function cleanText(value) {
  if (
    value === undefined ||
    value === null
  ) {
    return "";
  }

  return repairPolishText(value)
    .replace(
      /\b(?:\d{1,3}\.){3}\d{1,3}\b/g,
      "[IP]"
    )
    .replace(
      /https?:\/\/[^\s"'<>]+/gi,
      "[URL]"
    )
    .replace(
      /eqs=[^\s&"'<>]+/gi,
      "eqs=[REDACTED]"
    )
    .replace(
      /\b[A-Za-z0-9+/=_-]{60,}\b/g,
      "[TOKEN]"
    )
    .replace(/\s+/g, " ")
    .trim();
}

function safeName(value) {
  return String(value || "unknown")
    .replace(/[<>:"/\\|?*]/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 100);
}

function normalizeJNumber(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "");
}

function formatDate(date) {
  try {
    return new Intl.DateTimeFormat(
      "pl-PL",
      {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit"
      }
    ).format(date);
  } catch {
    return String(date);
  }
}

// ============================================================
// EXCEL DATABASE
// ============================================================

function loadAnalyzerDatabase() {
  if (
    !fs.existsSync(EXCEL_FILE)
  ) {
    throw new Error(
      `Nie znaleziono bazy Excel:\n${EXCEL_FILE}`
    );
  }

  console.log("");
  console.log(
    "Wczytuję bazę analizatorów z Excela..."
  );

  const workbook =
    XLSX.readFile(
      EXCEL_FILE,
      {
        cellDates: false
      }
    );

  const firstSheetName =
    workbook.SheetNames[0];

  if (!firstSheetName) {
    throw new Error(
      "Excel nie zawiera arkusza."
    );
  }

  const worksheet =
    workbook.Sheets[
      firstSheetName
    ];

  const rows =
    XLSX.utils.sheet_to_json(
      worksheet,
      {
        defval: "",
        raw: false
      }
    );

  const database =
    new Map();

  let valid = 0;

  for (const row of rows) {
    const rawJ =
      row["J Number"];

    const jno =
      normalizeJNumber(rawJ);

    if (!jno) {
      continue;
    }

    const asset = {
      jno,

      itemNumber:
        cleanText(
          row["Item number"]
        ),

      serialNumber:
        cleanText(
          row["Serial number"]
        ),

      description:
        cleanText(
          row["Description"]
        ),

      productGroup:
        cleanText(
          row["Product group"]
        ),

      productFamily:
        cleanText(
          row["Product family"]
        ),

      customerNumber:
        cleanText(
          row["Customer number"]
        ),

      installDate:
        cleanText(
          row["Install date"]
        ),

      country:
        cleanText(
          row["Country"]
        ),

      excelStatus:
        cleanText(
          row["Status"]
        ),

      softwareVersion:
        cleanText(
          row["Software version"]
        ),

      customerName:
        cleanText(
          row["Name"]
        ),

      address1:
        cleanText(
          row["Address 1"]
        ),

      address2:
        cleanText(
          row["Address 2"]
        ),

      address3:
        cleanText(
          row["Address 3"]
        ),

      city:
        cleanText(
          row["City"]
        ),

      zipCode:
        cleanText(
          row["Zip code"]
        )
    };

    database.set(
      jno,
      asset
    );

    // Dodatkowo bez J prefix,
    // ponieważ dashboard zwykle pokazuje sam numer.
    if (
      jno.startsWith("J")
    ) {
      database.set(
        jno.substring(1),
        asset
      );
    }

    valid++;
  }

  console.log(
    `Excel: ${valid} pozycji z J-number.`
  );

  console.log(
    `Arkusz: ${firstSheetName}`
  );

  return database;
}

function findAnalyzer(
  database,
  dashboardJno
) {
  const key =
    normalizeJNumber(
      dashboardJno
    );

  return (
    database.get(key) ||
    database.get(`J${key}`) ||
    null
  );
}

// ============================================================
// MODEL DISPLAY
// ============================================================

function getAnalyzerModel(asset) {
  if (!asset) {
    return "MODEL NIEZNANY — BRAK W BAZIE";
  }

  const family =
    cleanText(
      asset.productFamily
    );

  const description =
    cleanText(
      asset.description
    );

  const familyUpper =
    family.toUpperCase();

  const descUpper =
    description.toUpperCase();

  // Preferujemy Product family,
  // ale budujemy czytelną nazwę.

  if (
    familyUpper.includes("XT 3400") ||
    descUpper.includes("XT 3400")
  ) {
    return "VITROS XT 3400";
  }

  if (
    familyUpper.includes("XT 7600") ||
    descUpper.includes("XT 7600")
  ) {
    return "VITROS XT 7600";
  }

  if (
    familyUpper === "4600" ||
    descUpper.includes("VITROS 4600")
  ) {
    return "VITROS 4600";
  }

  if (
    familyUpper.includes("VISION MAX") ||
    descUpper.includes("VISION MAX")
  ) {
    return "ORTHO VISION Max";
  }

  if (
    familyUpper === "VISION" ||
    descUpper.includes("ORTHO VISION")
  ) {
    return "ORTHO VISION";
  }

  if (description) {
    return description;
  }

  if (family) {
    return family;
  }

  return "MODEL NIEZNANY";
}

function getLocation(asset) {
  if (!asset) {
    return {
      customer:
        "BRAK W BAZIE",

      address:
        "BRAK W BAZIE",

      city:
        "BRAK W BAZIE"
    };
  }

  const addressParts = [
    asset.address1,
    asset.address2,
    asset.address3
  ].filter(Boolean);

  return {
    customer:
      asset.customerName ||
      "brak danych",

    address:
      addressParts.length
        ? addressParts.join(", ")
        : "brak danych",

    city:
      [
        asset.zipCode,
        asset.city
      ]
        .filter(Boolean)
        .join(" ") ||
      "brak danych"
  };
}

// ============================================================
// ENTER
// ============================================================

function waitForEnter() {
  return new Promise(resolve => {
    const rl =
      readline.createInterface({
        input: process.stdin,
        output: process.stdout
      });

    console.log("");
    console.log(
      "================================================"
    );

    console.log(
      " NAWIGACJA RĘCZNA"
    );

    console.log(
      "================================================"
    );

    console.log("");
    console.log(
      "1. Zaloguj się do e-Connectivity."
    );

    console.log(
      "2. Przejdź przez wymagane podstrony."
    );

    console.log(
      "3. Wejdź na dashboard z tabelami analizatorów."
    );

    console.log(
      "4. Poczekaj, aż tabele się załadują."
    );

    console.log(
      "5. Wróć do PowerShell."
    );

    console.log("");
    console.log(
      "DOPIERO TERAZ naciśnij ENTER."
    );

    console.log("");

    rl.question(
      ">>> ENTER = START ANALYSIS: ",
      () => {
        rl.close();
        resolve();
      }
    );
  });
}

// ============================================================
// WAIT FOR DASHBOARD
// ============================================================

async function waitForDashboardStable(
  page
) {
  console.log("");
  console.log(
    "Sprawdzam kompletność dashboardu..."
  );

  let previous = "";
  let stableCount = 0;
  let bestState = null;

  const timeoutMs =
    90000;

  const start =
    Date.now();

  while (
    Date.now() - start <
    timeoutMs
  ) {
    const state =
      await page.evaluate(() => {
        const rows =
          Array.from(
            document.querySelectorAll(
              "tr[data-jno]"
            )
          );

        const jnos =
          rows
            .map(
              row =>
                row.getAttribute(
                  "data-jno"
                ) || ""
            )
            .filter(Boolean);

        let alertCount = 0;

        for (const row of rows) {
          for (
            const cell
            of Array.from(
              row.cells || []
            )
          ) {
            const classes =
              String(
                cell.className || ""
              )
                .split(/\s+/)
                .filter(Boolean);

            if (
              classes.includes("O") ||
              classes.includes("Y")
            ) {
              alertCount++;
            }
          }
        }

        return {
          analyzerCount:
            new Set(jnos).size,

          rowCount:
            rows.length,

          alertCount
        };
      });

    bestState =
      state;

    const signature =
      JSON.stringify(state);

    console.log(
      `  analizatory=${state.analyzerCount}, alerty=${state.alertCount}`
    );

    if (
      signature === previous &&
      state.analyzerCount > 1
    ) {
      stableCount++;
    } else {
      stableCount = 0;
    }

    previous =
      signature;

    if (
      stableCount >= 2
    ) {
      return state;
    }

    await sleep(2000);
  }

  return bestState;
}

// ============================================================
// INVENTORY
// ============================================================

async function inventoryDashboard(
  page
) {
  return page.evaluate(() => {

    function text(el) {
      return String(
        el?.innerText ||
        el?.textContent ||
        ""
      )
        .replace(/\s+/g, " ")
        .trim();
    }

    function getHeaders(table) {
      if (!table) return [];

      const rows =
        Array.from(
          table.querySelectorAll("tr")
        );

      for (const row of rows) {
        if (
          row.hasAttribute(
            "data-jno"
          )
        ) {
          continue;
        }

        const t =
          text(row);

        if (
          /status/i.test(t) ||
          /alert count/i.test(t)
        ) {
          return Array.from(
            row.cells || []
          ).map(text);
        }
      }

      return [];
    }

    const output = [];

    const rows =
      Array.from(
        document.querySelectorAll(
          "tr[data-jno]"
        )
      );

    for (const row of rows) {
      const jno =
        row.getAttribute(
          "data-jno"
        ) || "";

      if (!jno) continue;

      const table =
        row.closest("table");

      const headers =
        getHeaders(table);

      for (
        const cell
        of Array.from(
          row.cells || []
        )
      ) {
        const classes =
          String(
            cell.className || ""
          )
            .split(/\s+/)
            .filter(Boolean);

        let status = "";

        if (
          classes.includes("O")
        ) {
          status =
            "ORANGE";
        }

        if (
          classes.includes("Y")
        ) {
          status =
            "YELLOW";
        }

        if (!status) continue;

        const index =
          cell.cellIndex;

        let alertName =
          headers[index] || "";

        if (!alertName) {
          alertName =
            cell.getAttribute(
              "title"
            ) ||
            cell.getAttribute(
              "aria-label"
            ) ||
            "";
        }

        output.push({
          jno,

          status,

          alertName:
            String(
              alertName || ""
            )
              .replace(/\s+/g, " ")
              .trim(),

          cellId:
            cell.id || "",

          cellIndex:
            index
        });
      }
    }

    return output;
  });
}

// ============================================================
// FIND ALERT CELL
// ============================================================

async function findAlertCell(
  page,
  alert
) {
  if (alert.cellId) {
    const exact =
      page.locator(
        `[id="${alert.cellId}"]`
      );

    if (
      await exact.count()
    ) {
      return exact.first();
    }
  }

  const row =
    page.locator(
      `tr[data-jno="${alert.jno}"]`
    );

  if (
    !await row.count()
  ) {
    return null;
  }

  const cells =
    row.locator("td");

  if (
    alert.cellIndex <
    await cells.count()
  ) {
    return cells.nth(
      alert.cellIndex
    );
  }

  return null;
}

// ============================================================
// AAA POPUP
// ============================================================

async function findAAA(
  page,
  target
) {
  for (
    let attempt = 0;
    attempt < 6;
    attempt++
  ) {
    try {
      await target
        .scrollIntoViewIfNeeded();

      await target.hover();

      await sleep(350);

      await target.click({
        timeout: 3000
      });

      await sleep(500);

      await target.hover();

      await sleep(650);

      const aaa =
        page
          .locator(
            "#popupBalloon"
          )
          .getByText(
            "AAA",
            {
              exact: true
            }
          );

      if (
        await aaa.count() &&
        await aaa
          .first()
          .isVisible()
          .catch(
            () => false
          )
      ) {
        return aaa.first();
      }

    } catch {}

    await sleep(400);
  }

  return null;
}

async function clickAndCatchPage(
  context,
  link
) {
  const promise =
    context.waitForEvent(
      "page",
      {
        timeout: 7000
      }
    ).catch(
      () => null
    );

  await link.click();

  const newPage =
    await promise;

  if (!newPage) {
    return null;
  }

  await newPage.waitForLoadState(
    "domcontentloaded"
  );

  return newPage;
}

// ============================================================
// AAA VERIFY
// ============================================================

async function verifyAAA(
  page,
  jno
) {
  let pathname = "";

  try {
    pathname =
      new URL(
        page.url()
      ).pathname;
  } catch {}

  if (
    !pathname
      .toLowerCase()
      .endsWith(
        "/aaa.aspx"
      )
  ) {
    return false;
  }

  const body =
    cleanText(
      await page
        .locator("body")
        .innerText()
        .catch(
          () => ""
        )
    );

  return body.includes(
    jno
  );
}

// ============================================================
// AAA ALERT NAME
// ============================================================

async function extractAAAInformation(page, dashboardAlertName = "") {
  const raw = await page.locator("body").innerText().catch(() => "");
  let detectedName = cleanText(dashboardAlertName);
  if (!detectedName || /^alert count$/i.test(detectedName) || /^column\s+\d+/i.test(detectedName)) {
    const candidates = await page.evaluate(() => {
      const out=[];
      for (const el of document.querySelectorAll("h1,h2,h3,legend,caption,.title,.header,label,span")) {
        const t=String(el.innerText||el.textContent||"").replace(/\s+/g," ").trim();
        if(t.length>=4 && t.length<=160) out.push(t);
      }
      return [...new Set(out)].slice(0,300);
    }).catch(()=>[]);
    detectedName = cleanText(candidates.find(t => /status|alert|thermal|ring|logger|temperature|pressure|voltage|sensor|supply|incubator|wash|fluid|slide|tip|well|lamp|signal|connection/i.test(t) && !/alert analysis aid|date range|run|save|apply/i.test(t)) || "");
  }
  return { detectedName, aaaText: cleanText(raw) };
}

async function extractGenericAAAEvidence(page) {
  const raw = await page.locator("body").innerText().catch(() => "");
  const lines=[...new Set(raw.split(/\r?\n/).map(cleanText).filter(Boolean))];
  const ruleRx=/(?:desired|normal|range|threshold|trigger|yellow|orange|red|green|warning|critical|greater than|less than|above|below|between|difference|average|standard deviation|std|duty cycle|condition code|related code|temperature|pressure|voltage|current|speed|loss|correction|count|percent|%|hours?|days?)/i;
  const rules=lines.filter(x=>ruleRx.test(x) && !/\b(?:touch|click|press)\s+(?:save|apply)\b/i.test(x) && !/^apply$/i.test(x)).slice(0,250);
  const conditionCodes=[...new Set(lines.flatMap(x=>x.match(/\b[A-Z]{2,6}\d?(?:-\d{2,4}|[_-]\d{2,4}|\d{2,4})\b/g)||[]))].slice(0,100);
  const linkLabels=await page.locator("a").evaluateAll(els=>[...new Set(els.map(a=>String(a.innerText||a.textContent||"").replace(/\s+/g," ").trim()).filter(Boolean))].slice(0,200)).catch(()=>[]);
  return {rules,conditionCodes,linkLabels,rawTextSanitized:cleanText(raw)};
}

// ============================================================
// PEŁNA TREŚĆ STRONY AAA (tekst źródłowy do raportu)
// Struktura: nagłówki, akapity, listy z numeracją (1. / a. / i.), tabele
// z wyróżnionymi wierszami, pola tekstowe. Tylko odczyt DOM — bez klikania.
// ============================================================

// Funkcja wykonywana w przeglądarce (musi być samowystarczalna).
function collectPageBlocksInBrowser() {
  const SKIP = new Set(["SCRIPT", "STYLE", "NOSCRIPT", "TEMPLATE", "SVG", "CANVAS", "BUTTON", "SELECT", "OPTION", "INPUT", "MAP", "AREA", "IFRAME", "FRAME", "HEAD"]);
  const BLOCK = new Set(["DIV", "P", "SECTION", "ARTICLE", "MAIN", "HEADER", "FOOTER", "ASIDE", "NAV", "FORM", "FIELDSET", "TABLE", "UL", "OL", "LI", "DL", "DT", "DD", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE", "PRE", "CENTER", "TBODY", "THEAD", "TR", "TD", "TH", "TEXTAREA", "HR", "BR", "IMG", "SPAN-BLOCK"]);
  const blocks = [];
  const clean = s => String(s || "").replace(/\s+/g, " ").trim();

  const visible = el => {
    const cs = getComputedStyle(el);
    return cs.display !== "none" && cs.visibility !== "hidden";
  };

  const isBlockEl = el => {
    if (BLOCK.has(el.tagName)) return true;
    const d = getComputedStyle(el).display;
    return d === "block" || d === "list-item" || d === "table" || d === "flex" || d === "grid";
  };

  const roman = n => {
    const map = [[1000, "m"], [900, "cm"], [500, "d"], [400, "cd"], [100, "c"], [90, "xc"], [50, "l"], [40, "xl"], [10, "x"], [9, "ix"], [5, "v"], [4, "iv"], [1, "i"]];
    let out = "";
    for (const [v, s] of map) while (n >= v) { out += s; n -= v; }
    return out;
  };

  const marker = (list, index) => {
    if (list.tagName !== "OL") return "•";
    const start = Number(list.getAttribute("start") || 1);
    const n = start + index;
    const type = list.getAttribute("type") || "";
    const css = getComputedStyle(list).listStyleType;
    if (type === "a" || css === "lower-alpha" || css === "lower-latin") return String.fromCharCode(96 + ((n - 1) % 26) + 1) + ".";
    if (type === "A" || css === "upper-alpha" || css === "upper-latin") return String.fromCharCode(64 + ((n - 1) % 26) + 1) + ".";
    if (type === "i" || css === "lower-roman") return roman(n) + ".";
    if (type === "I" || css === "upper-roman") return roman(n).toUpperCase() + ".";
    return n + ".";
  };

  // Tekst elementu bez zagnieżdżonych list i tabel.
  const ownText = el => {
    let t = "";
    for (const node of el.childNodes) {
      if (node.nodeType === 3) t += node.textContent;
      else if (node.nodeType === 1) {
        if (SKIP.has(node.tagName) || node.tagName === "UL" || node.tagName === "OL" || node.tagName === "TABLE") continue;
        if (!visible(node)) continue;
        if (node.tagName === "BR") { t += " "; continue; }
        t += ownText(node);
      }
    }
    return t;
  };

  const bg = el => {
    const c = getComputedStyle(el).backgroundColor;
    return c && c !== "rgba(0, 0, 0, 0)" && c !== "transparent" ? c : "";
  };

  function walkList(list, level) {
    const items = [];
    let i = 0;
    for (const li of list.children) {
      if (li.tagName !== "LI" || !visible(li)) continue;
      const text = clean(ownText(li));
      if (text) items.push({ level, marker: marker(list, i), text });
      i++;
      for (const sub of li.querySelectorAll(":scope > ul, :scope > ol, :scope > div > ul, :scope > div > ol")) {
        items.push(...walkList(sub, level + 1));
      }
      for (const img of li.querySelectorAll(":scope > img, :scope > div > img")) {
        const alt = clean(img.getAttribute("alt") || img.getAttribute("title"));
        items.push({ level: level + 1, marker: "", text: alt ? `[obraz: ${alt}]` : "[obraz]" });
      }
    }
    return items;
  }

  function walkTable(table) {
    const rows = [];
    for (const tr of table.rows) {
      if (!visible(tr)) continue;
      const cells = [...tr.cells].map(c => clean(c.innerText || c.textContent));
      if (!cells.some(Boolean)) continue;
      const rowBg = bg(tr) || [...tr.cells].map(bg).find(Boolean) || "";
      rows.push({ cells, header: [...tr.cells].every(c => c.tagName === "TH"), bg: rowBg });
    }
    return rows;
  }

  function walk(el, depth) {
    if (!el || el.nodeType !== 1 || SKIP.has(el.tagName) || !visible(el)) return;
    const tag = el.tagName;

    if (/^H[1-6]$/.test(tag)) {
      const t = clean(el.innerText);
      if (t) blocks.push({ type: "heading", level: Number(tag[1]), text: t });
      return;
    }
    if (tag === "UL" || tag === "OL") {
      const items = walkList(el, 0);
      if (items.length) blocks.push({ type: "list", items });
      return;
    }
    if (tag === "TABLE") {
      // Tabela układu strony (z blokami w środku) — schodzimy głębiej zamiast spłaszczać.
      if (el.querySelector("table, ul, ol, h1, h2, h3, h4, textarea")) {
        for (const cell of el.querySelectorAll(":scope > tbody > tr > td, :scope > tr > td, :scope > thead > tr > th, :scope > tbody > tr > th")) walk(cell, depth + 1);
        return;
      }
      const rows = walkTable(el);
      if (rows.length) blocks.push({ type: "table", rows, caption: clean(el.caption && el.caption.innerText) });
      return;
    }
    if (tag === "TEXTAREA") {
      const t = clean(el.value || el.textContent);
      if (t) blocks.push({ type: "para", text: t, boxed: true });
      return;
    }
    if (tag === "IMG") {
      const alt = clean(el.getAttribute("alt") || el.getAttribute("title"));
      const src = el.getAttribute("src") || "";
      if (/chartaxd\.axd/i.test(src)) return; // wykresy mają własną sekcję
      if (el.width >= 120 && el.height >= 60) blocks.push({ type: "image", text: alt });
      return;
    }

    const hasBlockChild = [...el.children].some(c => !SKIP.has(c.tagName) && isBlockEl(c));
    if (!hasBlockChild) {
      const t = clean(el.innerText);
      if (t) {
        // Pogrubiony, krótki tekst bez kropki traktujemy jak nagłówek sekcji.
        const fw = Number(getComputedStyle(el).fontWeight) || 400;
        const fs = parseFloat(getComputedStyle(el).fontSize) || 12;
        const looksHeading = t.length <= 90 && !/[.:]$/.test(t) && (fw >= 600 || fs >= 16);
        blocks.push(looksHeading ? { type: "heading", level: fs >= 20 ? 2 : 3, text: t } : { type: "para", text: t });
      }
      return;
    }

    // Element mieszany: tekst bezpośredni + dzieci blokowe w kolejności dokumentu.
    let inline = "";
    const flush = () => {
      const t = clean(inline);
      if (t) blocks.push({ type: "para", text: t });
      inline = "";
    };
    for (const node of el.childNodes) {
      if (node.nodeType === 3) { inline += node.textContent; continue; }
      if (node.nodeType !== 1 || SKIP.has(node.tagName) || !visible(node)) continue;
      if (isBlockEl(node)) { flush(); walk(node, depth + 1); }
      else inline += " " + (node.innerText || node.textContent || "");
    }
    flush();
  }

  walk(document.body, 0);

  // Usuń sąsiadujące duplikaty (np. ten sam tekst w zagnieżdżonych kontenerach).
  const out = [];
  for (const b of blocks) {
    const prev = out[out.length - 1];
    if (prev && prev.type === b.type && prev.text && prev.text === b.text) continue;
    out.push(b);
  }
  return out.slice(0, 3000);
}

function isHighlightColor(rgb) {
  const m = String(rgb || "").match(/(\d+),\s*(\d+),\s*(\d+)/);
  if (!m) return "";
  const [r, g, b] = m.slice(1).map(Number);
  if (r > 200 && g > 180 && b < 140) return "żółte";
  if (r > 200 && g > 100 && g < 190 && b < 100) return "pomarańczowe";
  if (r > 190 && g < 100 && b < 100) return "czerwone";
  return "";
}

// Same słowa przycisków/linków-akcji to nie treść.
const ACTION_WORDS = /^(save|apply|run|reset|cancel|ok|close|print|submit|back|next|help|aaa)$/i;

function sanitizeBlocks(blocks) {
  const cut = s => cleanText(s).slice(0, 4000);
  return (blocks || []).filter(b => !(b.type === "para" && ACTION_WORDS.test(String(b.text).trim()))).map(b => {
    if (b.type === "list") return { ...b, items: b.items.map(i => ({ ...i, text: cut(i.text) })) };
    if (b.type === "table") {
      return {
        ...b,
        caption: cut(b.caption),
        rows: b.rows.map((r, i) => {
          const highlight = r.header ? "" : isHighlightColor(r.bg);
          // Pierwszy wiersz z kolorowym (nie wyróżniającym) tłem pełni rolę nagłówka, np. niebieski pasek.
          const header = r.header || (i === 0 && b.rows.length > 1 && !!r.bg && !highlight);
          return { cells: r.cells.map(cut), header, highlight: header ? "" : highlight };
        })
      };
    }
    return { ...b, text: cut(b.text) };
  });
}

async function extractAAAContent(page) {
  const sections = [];
  for (const frame of page.frames()) {
    try {
      const raw = await frame.evaluate(collectPageBlocksInBrowser);
      const blocks = sanitizeBlocks(raw);
      if (blocks.length) sections.push({ source: frame === page.mainFrame() ? "strona AAA" : "ramka strony AAA", blocks });
    } catch {}
  }
  return sections;
}

// Fakty z treści AAA: wiersze tabel wyróżnione kolorem na stronie (np. „Investigate”).
function aaaContentFacts(result) {
  const facts = [];
  for (const section of result.aaaContent || []) {
    let lastHeading = "";
    for (const b of section.blocks) {
      if (b.type === "heading") lastHeading = b.text;
      if (b.type !== "table") continue;
      const header = b.rows.find(r => r.header) || b.rows[0];
      for (const r of b.rows) {
        if (!r.highlight || r === header) continue;
        const pairs = r.cells.map((c, i) => (header && header !== r && header.cells[i] && i > 0 ? `${header.cells[i]}: ${c}` : c)).filter(Boolean);
        facts.push(`AAA wyróżnia (${r.highlight} tło) w tabeli „${b.caption || lastHeading || "bez tytułu"}”: ${pairs.join("; ")}.`);
      }
    }
  }
  return facts.slice(0, 12);
}

// Podstrony AAA (opt-in: AAA_FOLLOW_LINKS=1). Tylko linki z tej samej domeny, otwierane GET-em,
// bez linków wyglądających na akcje (Save, Apply, Run, Reset, Delete, logowanie).
const UNSAFE_LINK = /save|apply|run|reset|delete|remove|submit|update|logout|log\s*out|login|sign|wyloguj|javascript:|mailto:|__dopostback/i;

async function extractSubpages(context, page, limit = 8) {
  if (process.env.AAA_FOLLOW_LINKS !== "1") return [];

  const origin = new URL(page.url()).origin;
  const links = await page.locator("a[href]").evaluateAll(els => els.map(a => ({
    href: a.href,
    text: String(a.innerText || a.textContent || "").replace(/\s+/g, " ").trim()
  }))).catch(() => []);

  const seen = new Set();
  const picked = links.filter(l => {
    if (!l.text || !l.href) return false;
    let u;
    try { u = new URL(l.href); } catch { return false; }
    if (u.origin !== origin) return false;
    if (UNSAFE_LINK.test(l.text) || UNSAFE_LINK.test(u.pathname + u.search)) return false;
    const key = u.pathname + u.search;
    if (key === new URL(page.url()).pathname + new URL(page.url()).search || seen.has(key)) return false;
    seen.add(key);
    return true;
  }).slice(0, limit);

  const out = [];
  for (const link of picked) {
    const sub = await context.newPage();
    try {
      const response = await sub.goto(link.href, { waitUntil: "domcontentloaded", timeout: 20000 });
      const type = (response && response.headers()["content-type"]) || "";
      if (!/html/i.test(type)) {
        out.push({ title: cleanText(link.text), note: `dokument ${type.split(";")[0] || "nieznanego typu"} — nie odczytano treści`, sections: [] });
      } else {
        out.push({ title: cleanText(link.text), sections: await extractAAAContent(sub) });
      }
    } catch (error) {
      out.push({ title: cleanText(link.text), note: `brak pewnego odczytu: ${cleanText(error.message).slice(0, 120)}`, sections: [] });
    } finally {
      await sub.close().catch(() => {});
    }
  }
  return out;
}

function loadServiceKnowledge(){
  if(!fs.existsSync(KNOWLEDGE_FILE)){console.log("Service Knowledge: brak service_knowledge.json");return []}
  try{const p=JSON.parse(fs.readFileSync(KNOWLEDGE_FILE,"utf8"));const d=Array.isArray(p)?p:(p.documents||p.procedures||[]);console.log(`Service Knowledge: ${d.length} dokumentów/procedur.`);return d}catch(e){console.log(`Service Knowledge: błąd: ${cleanText(e.message)}`);return []}
}
function knowledgeTerms(r){const z=[r.alertName,...(r.genericEvidence?.conditionCodes||[]),...(r.genericEvidence?.rules||[]).slice(0,25)].join(" ").toLowerCase();const stop=new Set(["status","alert","full","the","and","for","with","from","this","that","desired","normal","range","yellow","orange","green","red","value","values","condition","code","codes","greater","less","than","above","below"]);return [...new Set(z.match(/[a-z0-9][a-z0-9_-]{2,}/g)||[])].filter(x=>!stop.has(x)).slice(0,35)}
function searchServiceKnowledge(k,r){
  if(!Array.isArray(k)||!k.length)return[];const terms=knowledgeTerms(r), model=String(r.model||"").toLowerCase(), compact=model.replace(/vitros|ortho/g,"").trim(), out=[];
  for(const d of k){const title=cleanText(d.title||d.name||d.procedure||d.heading||""), fp=cleanText(d.path||d.file||d.source||""), body=cleanText(d.text||d.content||d.summary||""), mt=cleanText(d.model||d.product||d.family||d.productFamily||"");const hay=`${title} ${fp} ${body} ${mt}`.toLowerCase();let score=0,h=[];if(model&&hay.includes(model))score+=8;if(compact&&hay.includes(compact))score+=5;for(const t of terms){if(title.toLowerCase().includes(t)){score+=5;h.push(t)}else if(fp.toLowerCase().includes(t)){score+=3;h.push(t)}else if(hay.includes(t)){score+=1;h.push(t)}}for(const c of (r.genericEvidence?.conditionCodes||[])){if(hay.includes(c.toLowerCase())){score+=12;h.push(c)}}if(score>=8&&h.length)out.push({title:title||path.basename(fp)||"Dokument serwisowy",path:fp,score,matchedTerms:[...new Set(h)].slice(0,12)})}
  return out.sort((x,y)=>y.score-x.score).slice(0,10)
}

// ============================================================
// DATE RANGE
// ============================================================

async function readDateRange(
  page
) {
  const start =
    page.locator(
      "#ctl00_AAA_HeaderMain1_DateRange_StartDT_TextBox1"
    );

  const end =
    page.locator(
      "#ctl00_AAA_HeaderMain1_DateRange_EndDT_TextBox1"
    );

  return {
    start:
      await start.count()
        ? await start
            .inputValue()
            .catch(
              () => ""
            )
        : "",

    end:
      await end.count()
        ? await end
            .inputValue()
            .catch(
              () => ""
            )
        : ""
  };
}

// ============================================================
// DATA LOGGER
// ============================================================

async function extractDataLogger(
  page
) {
  const body =
    cleanText(
      await page
        .locator("body")
        .innerText()
        .catch(
          () => ""
        )
    );

  if (
    !/Data Logger Status/i
      .test(body)
  ) {
    return {
      detected:
        false,

      lastConnected:
        "",

      aFile: [],

      bFile: []
    };
  }

  const match =
    body.match(
      /Last Connected[\s\S]{0,150}?(\d{1,2}\/\d{1,2}\/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[AP]M)/i
    );

  const lastConnected =
    match
      ? match[1]
      : "";

  function getPairs(label) {
    const idx =
      body
        .toLowerCase()
        .indexOf(
          label.toLowerCase()
        );

    if (
      idx < 0
    ) {
      return [];
    }

    const fragment =
      body.slice(
        idx,
        idx + 500
      );

    const values =
      fragment.match(
        /\b\d+\/\d+\b/g
      ) || [];

    const days = [
      "D6",
      "D5",
      "D4",
      "D3",
      "D2",
      "D1",
      "Today"
    ];

    return values
      .slice(0, 7)
      .map(
        (value, index) => ({
          day:
            days[index],

          value
        })
      );
  }

  return {
    detected:
      true,

    lastConnected,

    aFile:
      getPairs(
        "A-file upload status"
      ),

    bFile:
      getPairs(
        "B-file upload status"
      )
  };
}

// ============================================================
// AAA TECHNICAL RULES
// JSON ONLY
// ============================================================

async function extractDiagnosticRules(
  page
) {
  const raw =
    await page
      .locator("body")
      .innerText()
      .catch(
        () => ""
      );

  const lines =
    raw
      .split(/\r?\n/)
      .map(cleanText)
      .filter(Boolean);

  const keywords = [
    "step loss",
    "slot correction",
    "drag",
    "duty cycle",
    "thermistor",
    "ambient",
    "condition code",
    "belt",
    "motor",
    "roller",
    "bearing",
    "heat pump",
    "temperature",
    "adjustment"
  ];

  return [
    ...new Set(
      lines.filter(
        line => {
          const low =
            line.toLowerCase();

          if (
            low.includes(
              "touch save"
            ) ||
            low.includes(
              "click save"
            ) ||
            low.includes(
              "press save"
            ) ||
            low.includes(
              "apply"
            )
          ) {
            return false;
          }

          return keywords.some(
            keyword =>
              low.includes(
                keyword
              )
          );
        }
      )
    )
  ].slice(
    0,
    100
  );
}

// ============================================================
// CHART CAPTURE
// ============================================================

const CHART_URL_BY_FILE = new Map();

function createChartHandler(
  page,
  chartDir,
  chartFiles
) {
  let counter = 0;

  const handler =
    async response => {
      try {
        const url =
          new URL(
            response.url()
          );

        if (
          !url.pathname
            .toLowerCase()
            .endsWith(
              "/chartaxd.axd"
            )
        ) {
          return;
        }

        const type =
          response.headers()[
            "content-type"
          ] || "";

        if (
          !type
            .toLowerCase()
            .includes("image")
        ) {
          return;
        }

        const body =
          await response.body();

        counter++;

        const file =
          path.join(
            chartDir,
            `chart_${counter}.png`
          );

        fs.writeFileSync(
          file,
          body
        );

        CHART_URL_BY_FILE.set(file, response.url());

        chartFiles.push(
          file
        );

      } catch {}
    };

  page.on(
    "response",
    handler
  );

  return handler;
}



// ============================================================
// LOKALNY ODCZYT WYKRESÓW (bez wysyłania obrazów na zewnątrz)
// Wymaga: npm install pngjs tesseract.js @tesseract.js-data/eng
// Wyłączenie: AAA_LOCAL_CHARTS=0
// ============================================================

function buildChartReader() {
  // ============================================================
  // LOKALNY CZYTNIK WYKRESÓW (MS Chart / chartaxd.axd)
  // Odczyt wartości z obrazu PNG bez wysyłania go na zewnątrz:
  //  - kalibracja osi Y: podziałki + OCR etykiet (tesseract.js) + dopasowanie "ładnego" kroku
  //  - serie: dokładne kolory MS Chart (niebieski, brązowy, czerwony)
  //  - progi: fioletowe/żółte poziome linie
  // ============================================================

  const { PNG } = require("pngjs");

  const PALETTE = [
    { name: "niebieska", rgb: [0, 0, 255], kind: "series" },
    { name: "brązowa", rgb: [139, 69, 19], kind: "series" },
    { name: "czerwona", rgb: [255, 0, 0], kind: "series" },
    { name: "fioletowa", rgb: [148, 0, 211], kind: "threshold" },
    { name: "żółta", rgb: [255, 255, 0], kind: "threshold" }
  ];

  function decode(buffer) {
    const png = PNG.sync.read(buffer);
    return { w: png.width, h: png.height, data: png.data };
  }

  function pix(img, x, y) {
    const i = (y * img.w + x) * 4;
    return [img.data[i], img.data[i + 1], img.data[i + 2]];
  }

  function isDark(img, x, y) {
    const i = (y * img.w + x) * 4;
    return img.data[i] < 110 && img.data[i + 1] < 110 && img.data[i + 2] < 110;
  }

  function gray(img, x, y) {
    const [r, g, b] = pix(img, x, y);
    return 0.299 * r + 0.587 * g + 0.114 * b;
  }

  function near(c, rgb, tol) {
    return Math.abs(c[0] - rgb[0]) <= tol && Math.abs(c[1] - rgb[1]) <= tol && Math.abs(c[2] - rgb[2]) <= tol;
  }

  function findLeftAxis(img) {
    const maxX = Math.floor(img.w * 0.3);
    const counts = [];
    let mx = 0;
    for (let x = 15; x < maxX; x++) {
      let n = 0;
      for (let y = 0; y < img.h; y++) if (isDark(img, x, y)) n++;
      counts[x] = n;
      if (n > mx) mx = n;
    }
    for (let x = 15; x < maxX; x++) if (counts[x] >= 0.8 * mx) return x;
    return -1;
  }

  function groupRows(rows) {
    const groups = [];
    for (const y of rows) {
      if (groups.length && y - groups[groups.length - 1].slice(-1)[0] <= 1) groups[groups.length - 1].push(y);
      else groups.push([y]);
    }
    return groups.map(g => g.reduce((a, b) => a + b, 0) / g.length);
  }

  function findTicks(img, left) {
    const rows = [];
    for (let y = 15; y < img.h - 15; y++) {
      let n = 0;
      for (let x = left - 6; x < left; x++) if (isDark(img, x, y)) n++;
      if (n >= 4) rows.push(y);
    }
    const ticks = groupRows(rows);
    if (ticks.length < 3) return ticks;

    // Zostaw najdłuższy ciąg równo rozmieszczonych podziałek.
    let best = [];
    for (let i = 0; i < ticks.length - 1; i++) {
      const d = ticks[i + 1] - ticks[i];
      if (d < 8) continue;
      const run = [ticks[i], ticks[i + 1]];
      for (let j = i + 2; j < ticks.length; j++) {
        if (Math.abs(ticks[j] - run[run.length - 1] - d) <= 2.5) run.push(ticks[j]);
        else break;
      }
      if (run.length > best.length) best = run;
    }
    return best.length >= 2 ? best : ticks;
  }

  function labelPng(img, left, y, scale, thr) {
    const y0 = Math.max(0, Math.round(y) - 7);
    const y1 = Math.min(img.h, Math.round(y) + 8);
    const x0 = Math.max(0, left - 70);
    const x1 = left - 7;
    const w = x1 - x0;
    const h = y1 - y0;
    const g = new Float32Array(w * h);
    for (let yy = 0; yy < h; yy++) for (let xx = 0; xx < w; xx++) g[yy * w + xx] = gray(img, x0 + xx, y0 + yy);

    const W = w * scale;
    const H = h * scale;
    const bw = new Uint8Array(W * H).fill(255);
    let minX = W, maxX = -1, minY = H, maxY = -1, blackPx = 0;

    for (let Y = 0; Y < H; Y++) {
      const sy = Math.min(h - 1, Y / scale);
      const y0i = Math.floor(sy), y1i = Math.min(h - 1, y0i + 1), fy = sy - y0i;
      for (let X = 0; X < W; X++) {
        const sx = Math.min(w - 1, X / scale);
        const x0i = Math.floor(sx), x1i = Math.min(w - 1, x0i + 1), fx = sx - x0i;
        const v =
          g[y0i * w + x0i] * (1 - fx) * (1 - fy) + g[y0i * w + x1i] * fx * (1 - fy) +
          g[y1i * w + x0i] * (1 - fx) * fy + g[y1i * w + x1i] * fx * fy;
        if (v < thr) {
          bw[Y * W + X] = 0;
          blackPx++;
          if (X < minX) minX = X;
          if (X > maxX) maxX = X;
          if (Y < minY) minY = Y;
          if (Y > maxY) maxY = Y;
        }
      }
    }

    if (blackPx < 6 * scale || maxX < 0) return null;

    const m = 30;
    const cw = maxX - minX + 1 + 2 * m;
    const ch = maxY - minY + 1 + 2 * m;
    const out = new PNG({ width: cw, height: ch });
    out.data.fill(255);
    for (let Y = minY; Y <= maxY; Y++) {
      for (let X = minX; X <= maxX; X++) {
        if (bw[Y * W + X] === 0) {
          const o = ((Y - minY + m) * cw + (X - minX + m)) * 4;
          out.data[o] = out.data[o + 1] = out.data[o + 2] = 0;
        }
      }
    }
    return PNG.sync.write(out);
  }


  // Wykrywa kropkę dziesiętną na obrazie etykiety (OCR często ją gubi) i szacuje liczbę miejsc po przecinku.
  function detectDecimals(img, left, y) {
    const y0 = Math.max(0, Math.round(y) - 8);
    const y1 = Math.min(img.h - 1, Math.round(y) + 8);
    const x0 = Math.max(0, left - 70);
    const x1 = left - 8;

    const ink = [];
    for (let yy = y0; yy <= y1; yy++) {
      const row = [];
      for (let xx = x0; xx <= x1; xx++) row.push(gray(img, xx, yy) < 175);
      ink.push(row);
    }

    let r0 = -1, r1 = -1;
    ink.forEach((row, i) => { if (row.some(Boolean)) { if (r0 < 0) r0 = i; r1 = i; } });
    if (r0 < 0 || r1 - r0 < 5) return { dots: 0, k: 0 };

    const width = x1 - x0 + 1;
    const colInk = [];
    for (let c = 0; c < width; c++) {
      let top = -1, bottom = -1;
      for (let r = r0; r <= r1; r++) if (ink[r][c]) { if (top < 0) top = r; bottom = r; }
      colInk.push(top < 0 ? null : [top, bottom]);
    }

    // grupy kolumn rozdzielone pustymi kolumnami
    const groups = [];
    let cur = null;
    for (let c = 0; c < width; c++) {
      if (colInk[c]) {
        if (!cur) cur = { a: c, b: c, top: colInk[c][0], bottom: colInk[c][1] };
        else { cur.b = c; cur.top = Math.min(cur.top, colInk[c][0]); cur.bottom = Math.max(cur.bottom, colInk[c][1]); }
      } else if (cur) { groups.push(cur); cur = null; }
    }
    if (cur) groups.push(cur);

    const inkH = r1 - r0 + 1;
    const dotIdx = groups.findIndex((g, i) =>
      i > 0 && i < groups.length - 1 &&
      g.b - g.a <= 2 &&
      g.bottom >= r1 - 1 &&
      g.top >= r1 - 2 &&
      (g.bottom - g.top) <= 2
    );
    if (dotIdx < 0) return { dots: 0, k: 0 };

    const dot = groups[dotIdx];
    const rightInk = groups[groups.length - 1].b;
    const dW = 0.78 * inkH;
    const dotCenter = (dot.a + dot.b) / 2;
    const k = Math.max(1, Math.round((rightInk - dotCenter) / dW - 0.15));
    return { dots: 1, k };
  }

  function digitsOnly(t) {
    const s = String(t || "").replace(/[%\s]/g, "").replace(/-+$/, "");
    const neg = s.startsWith("-");
    const d = s.replace(/[^0-9]/g, "");
    return d ? { d, neg } : null;
  }

  function niceSteps() {
    const out = [];
    for (let e = -3; e <= 4; e++) for (const m of [1, 2, 2.5, 5]) out.push(m * Math.pow(10, e));
    return out;
  }

  function parseLabel(t) {
    const s = String(t || "").replace(/[%\s]/g, "").replace(/-+$/, "");
    return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : null;
  }

  // Zamienia surowe odczyty OCR + wykrycie kropki z pikseli na listę możliwych wartości etykiety.
  function labelValues(reads, k, globalK) {
    const out = [];
    for (const t of reads) {
      const p = digitsOnly(t);
      if (!p) continue;
      const base = Number(p.d) * (p.neg ? -1 : 1);
      if (k > 0) out.push(base / Math.pow(10, k));
      else {
        out.push(base);
        if (globalK > 0) out.push(base / Math.pow(10, globalK));
      }
    }
    return out;
  }

  function solveAxis(perTick) {
    const globalK = Math.max(0, ...perTick.map(t => t.k));
    const valuesPerTick = perTick.map(t => labelValues(t.reads, t.k, globalK));

    const cands = [];
    valuesPerTick.forEach((vals, k) => vals.forEach(v => cands.push([k, v])));

    let best = null;
    for (const step of niceSteps()) {
      for (const [k, v] of cands) {
        const v0 = v + k * step;
        let score = 0;
        valuesPerTick.forEach((vals, kk) => {
          const pred = v0 - kk * step;
          for (const rv of vals) if (Math.abs(rv - pred) < step * 0.01 + 1e-9) score++;
        });
        const key = [score, -String(step).length];
        if (!best || key[0] > best.key[0] || (key[0] === best.key[0] && key[1] > best.key[1])) {
          best = { key, v0, step, score };
        }
      }
    }

    if (!best || best.score < 2) return null;
    const dec = Math.max(0, -Math.floor(Math.log10(best.step) + 1e-9));
    return { v0: Number(best.v0.toFixed(dec + 1)), step: best.step, score: best.score };
  }

  async function readTitle(img, worker, plotTop) {
    const top = Math.max(0, Math.min(plotTop - 5, Math.floor(img.h * 0.16)));
    // Tytuł = pierwszy zwarty pas wierszy z tekstem (pomijamy ramkę u góry); kończy go pusta przerwa.
    let y0 = -1, y1 = -1, blank = 0;
    for (let y = 6; y < top; y++) {
      let n = 0;
      for (let x = 30; x < img.w - 30; x++) if (gray(img, x, y) < 120) n++;
      if (n > 0) {
        if (y0 < 0) y0 = y;
        y1 = y;
        blank = 0;
      } else if (y0 >= 0 && ++blank >= 4) {
        break;
      }
    }
    if (y0 < 0) return "";

    let xa = img.w, xb = 0;
    for (let y = y0; y <= y1; y++) {
      for (let x = 30; x < img.w - 30; x++) {
        if (gray(img, x, y) < 150) { if (x < xa) xa = x; if (x > xb) xb = x; }
      }
    }
    if (xb <= xa) return "";

    const scale = 3;
    const m = 14;
    const w = (xb - xa + 1) * scale;
    const h = (y1 - y0 + 1) * scale;
    const out = new PNG({ width: w + 2 * m, height: h + 2 * m });
    out.data.fill(255);
    for (let Y = 0; Y < h; Y++) {
      for (let X = 0; X < w; X++) {
        const sx = xa + Math.floor(X / scale), sy = y0 + Math.floor(Y / scale);
        const v = gray(img, sx, sy) < 150 ? 0 : 255;
        const o = ((Y + m) * out.width + (X + m)) * 4;
        out.data[o] = out.data[o + 1] = out.data[o + 2] = v;
        out.data[o + 3] = 255;
      }
    }

    await worker.setParameters({ tessedit_char_whitelist: "", tessedit_pageseg_mode: "7" });
    const r = await worker.recognize(PNG.sync.write(out));
    return String(r.data.text || "")
      .replace(/\}/g, ")")
      .replace(/\{/g, "(")
      .replace(/\s+/g, " ")
      .trim();
  }

  function median(arr) {
    const s = [...arr].sort((a, b) => a - b);
    const m = Math.floor(s.length / 2);
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  }

  // Odcinki, w których mediany kolumn leżą na prostej (±1 px), dłuższe niż minLen
  // i z różnicą wysokości ≥ 2 px. Tak AAA rysuje lukę w danych.
  function findLinearRuns(pts, minLen) {
    const runs = [];
    let i = 0;
    while (i < pts.length - 2) {
      let best = -1;
      for (let j = i + 2; j < pts.length && pts[j].x - pts[j - 1].x <= 2; j++) {
        const a = pts[i], b = pts[j], slope = (b.y - a.y) / (b.x - a.x);
        let fits = true;
        for (let k = i + 1; k < j; k++) {
          if (Math.abs(a.y + slope * (pts[k].x - a.x) - pts[k].y) > 1) { fits = false; break; }
        }
        if (!fits) break;
        best = j;
      }
      if (best > 0 && pts[best].x - pts[i].x >= minLen && Math.abs(pts[best].y - pts[i].y) >= 2) {
        runs.push({ x0: pts[i].x, x1: pts[best].x });
        i = best;
      } else {
        i++;
      }
    }
    return runs;
  }

  // Typowa "chropowatość" serii: mediana |2. różnicy| median kolumn (w pikselach).
  function roughness(pts) {
    const d = [];
    for (let i = 1; i < pts.length - 1; i++) {
      if (pts[i].x - pts[i - 1].x === 1 && pts[i + 1].x - pts[i].x === 1) {
        d.push(Math.abs(pts[i + 1].y - 2 * pts[i].y + pts[i - 1].y));
      }
    }
    return d.length ? median(d) : 0;
  }

  function mergeRuns(runs) {
    const out = [];
    for (const r of [...runs].sort((a, b) => a.x0 - b.x0)) {
      const last = out[out.length - 1];
      if (last && r.x0 <= last.x1) last.x1 = Math.max(last.x1, r.x1);
      else out.push({ ...r });
    }
    return out;
  }

  async function readChart(buffer, worker, opts = {}) {
    const img = decode(buffer);
    const result = { ok: false, reason: "", title: "", axis: null, thresholds: [], series: [] };

    const left = findLeftAxis(img);
    if (left < 0) { result.reason = "nie znaleziono osi Y"; return result; }

    const ticks = findTicks(img, left);
    if (ticks.length < 3) { result.reason = "za mało podziałek osi Y"; return result; }

    // OCR etykiet: kilka wariantów skali/progu na etykietę, potem głosowanie na ciąg arytmetyczny.
    await worker.setParameters({
      tessedit_char_whitelist: "0123456789.%-",
      tessedit_pageseg_mode: "7"
    });

    const perTick = [];
    for (const y of ticks) {
      const arr = [];
      for (const scale of [5, 7]) {
        for (const thr of [150, 175, 200]) {
          const png = labelPng(img, left, y, scale, thr);
          if (!png) continue;
          const r = await worker.recognize(png);
          arr.push(String(r.data.text || "").trim());
        }
      }
      perTick.push({ reads: arr, k: detectDecimals(img, left, y).k });
    }

    const axis = solveAxis(perTick);
    if (!axis) { result.reason = "nie udało się odczytać skali osi Y"; return result; }

    const yTop = ticks[0];
    const yBot = ticks[ticks.length - 1];
    const pxPerStep = (yBot - yTop) / (ticks.length - 1);
    const valueAt = y => axis.v0 - ((y - yTop) / pxPerStep) * axis.step;
    const places = Math.max(0, -Math.floor(Math.log10(axis.step) + 1e-9)) + 1;
    const resolution = axis.step / pxPerStep;
    const nice = v => {
      // Progi to zwykle liczby całkowite lub połówki — przyciągnij, jeśli mieści się w błędzie ~0,6 piksela.
      for (const snap of [Math.round(v), Math.round(v * 2) / 2]) {
        if (Math.abs(v - snap) <= 0.6 * resolution) return Number(snap.toFixed(places));
      }
      return Number(v.toFixed(places));
    };

    // Prawy kraniec obszaru wykresu: ciągła czarna oś X w wierszu najniższej podziałki.
    let right = left + 1;
    let gap = 0;
    for (let x = left + 1; x < img.w - 5; x++) {
      if (isDark(img, x, Math.round(yBot)) || isDark(img, x, Math.round(yBot) + 1)) { right = x; gap = 0; }
      else if (++gap > 3) break;
    }
    const plotW = right - left;

    result.axis = {
      top: Number(axis.v0.toFixed(4)),
      bottom: Number((axis.v0 - axis.step * (ticks.length - 1)).toFixed(4)),
      step: axis.step,
      resolution: Number((axis.step / pxPerStep).toFixed(4)),
      ocrAgreement: axis.score
    };
    result.title = await readTitle(img, worker, yTop).catch(() => "");

    const yMinRow = Math.round(yTop) - 3;
    const yMaxRow = Math.round(yBot) + 1;

    for (const col of PALETTE) {
      const colsMap = new Map();
      const rowCount = new Map();
      let total = 0;

      for (let y = yMinRow; y <= yMaxRow; y++) {
        for (let x = left + 2; x <= right; x++) {
          if (near(pix(img, x, y), col.rgb, 12)) {
            total++;
            rowCount.set(y, (rowCount.get(y) || 0) + 1);
            if (!colsMap.has(x)) colsMap.set(x, []);
            colsMap.get(x).push(y);
          }
        }
      }
      if (total < 25) continue;

      // Poziome linie = progi (wypełniają większość szerokości wykresu). Tylko kolory progów
      // (fioletowy, żółty, CLAUDE.md 5.1) — płaska seria nie jest progiem.
      const lineRows = col.kind !== "threshold" ? [] :
        [...rowCount.entries()].filter(([, n]) => n >= 0.6 * plotW).map(([y]) => y).sort((a, b) => a - b);
      const thrCenters = groupRows(lineRows);

      for (const c of thrCenters) {
        result.thresholds.push({ color: col.name, value: nice(valueAt(c)) });
      }

      if (col.name === "czerwona" && total < 400 && !thrCenters.length) {
        const ys = [...colsMap.values()].flat();
        const xs = [...colsMap.keys()];
        result.marker = {
          value: nice(valueAt(median(ys))),
          at: (median(xs) - left) / plotW,
          note: "czerwony znacznik (zwykle: wartość tego analizatora)"
        };
        continue;
      }

      const skip = new Set();
      for (const c of thrCenters) for (let d = -3; d <= 3; d++) skip.add(Math.round(c) + d);

      if (col.kind === "threshold" && thrCenters.length) continue;

      const medians = [];
      for (const [x, ys] of [...colsMap.entries()].sort((a, b) => a[0] - b[0])) {
        const clean = ys.filter(y => !skip.has(y));
        if (!clean.length) continue;
        medians.push({ x, y: median(clean), top: Math.min(...clean), bottom: Math.max(...clean) });
      }
      if (medians.length < 8) continue;

      // Luki w danych (CLAUDE.md 5.3): AAA łączy je prostą linią. Odcinek idealnie liniowy,
      // dłuższy niż ~3% szerokości i nie poziomy, traktujemy jako lukę i usuwamy ze statystyk.
      // Poziome plateau (np. nasycenie 100%) zostaje w danych.
      // W serii gładkiej (np. powolna zmiana temperatury) prosty odcinek może być prawdziwymi danymi,
      // więc tam oznaczamy go tylko jako możliwą lukę i nie usuwamy.
      const minGapPx = Math.max(6, 0.03 * plotW);
      const linearRuns = findLinearRuns(medians, minGapPx);
      const noisy = roughness(medians) >= 0.5;
      const gapRuns = noisy ? linearRuns : [];
      // Gdy "liniowa" jest ponad połowa szerokości, seria jest po prostu gładka — nic nie oznaczamy.
      const suspectedAll = noisy ? [] : mergeRuns(linearRuns);
      const suspected = suspectedAll.reduce((n, g) => n + g.x1 - g.x0, 0) > 0.5 * plotW ? [] : suspectedAll;
      for (let i = 1; i < medians.length; i++) {
        if (medians[i].x - medians[i - 1].x > minGapPx) gapRuns.push({ x0: medians[i - 1].x, x1: medians[i].x });
      }
      const gaps = mergeRuns(gapRuns);
      const kept = medians.filter(m => !gaps.some(g => m.x > g.x0 && m.x < g.x1));
      if (kept.length < 8) continue;

      const vals = kept.map(m => valueAt(m.y));
      const q = Math.max(1, Math.floor(vals.length / 4));
      const avg = a => a.reduce((s, v) => s + v, 0) / a.length;

      const maxVal = valueAt(Math.min(...kept.map(m => m.top)));
      const minVal = valueAt(Math.max(...kept.map(m => m.bottom)));
      const maxCol = kept.reduce((b, m) => (m.y < b.y ? m : b), kept[0]);
      const minCol = kept.reduce((b, m) => (m.y > b.y ? m : b), kept[0]);
      const gapWidth = gaps.reduce((sum, g) => sum + (g.x1 - g.x0), 0);

      const s = {
        color: col.name,
        pixels: total,
        min: Number(minVal.toFixed(3)),
        max: Number(maxVal.toFixed(3)),
        avg: Number(avg(vals).toFixed(3)),
        first: Number(vals[0].toFixed(3)),
        last: Number(vals[vals.length - 1].toFixed(3)),
        trend: Number((avg(vals.slice(-q)) - avg(vals.slice(0, q))).toFixed(3)),
        maxAt: (maxCol.x - left) / plotW,
        minAt: (minCol.x - left) / plotW,
        coverage: Number((medians.length / plotW).toFixed(3)),
        gapShare: Number((gapWidth / plotW).toFixed(3)),
        gaps: gaps.map(g => ({
          from: Number(((g.x0 - left) / plotW).toFixed(4)),
          to: Number(((g.x1 - left) / plotW).toFixed(4))
        })),
        suspectedGaps: suspected.map(g => ({
          from: Number(((g.x0 - left) / plotW).toFixed(4)),
          to: Number(((g.x1 - left) / plotW).toFixed(4))
        })),
        // Próbki bez luk: [ułamek osi X, wartość] — do analizy wg typu alertu i wykresów odtworzonych.
        samples: kept.map(m => [Number(((m.x - left) / plotW).toFixed(4)), Number(valueAt(m.y).toFixed(places))]),
        versus: []
      };

      result.series.push({ ...s, _vals: vals });
    }

    // Porównanie serii z progami: udział kolumn powyżej / poniżej.
    for (const s of result.series) {
      for (const t of result.thresholds) {
        const above = s._vals.filter(v => v > t.value + result.axis.resolution).length;
        const below = s._vals.filter(v => v < t.value - result.axis.resolution).length;
        const n = s._vals.length;
        s.versus.push({
          threshold: t.value,
          color: t.color,
          above: Number((above / n).toFixed(3)),
          below: Number((below / n).toFixed(3)),
          marginToMax: Number((s.max - t.value).toFixed(3)),
          marginToMin: Number((s.min - t.value).toFixed(3)),
          lastVsThreshold: Number((s.last - t.value).toFixed(3))
        });
      }
      delete s._vals;
    }

    result.ok = result.series.length > 0 || result.thresholds.length > 0;
    if (!result.ok) result.reason = "nie znaleziono serii w znanych kolorach";
    return result;
  }

  async function debugReads(buffer, worker) {
    const img = decode(buffer);
    const left = findLeftAxis(img);
    const ticks = findTicks(img, left);
    await worker.setParameters({ tessedit_char_whitelist: "0123456789.%-", tessedit_pageseg_mode: "7" });
    const out = [];
    for (const y of ticks) {
      const arr = [];
      for (const scale of [5, 7]) for (const thr of [150, 175, 200]) {
        const png = labelPng(img, left, y, scale, thr);
        if (!png) { arr.push(null); continue; }
        const r = await worker.recognize(png);
        arr.push(String(r.data.text || "").trim());
      }
      out.push({ reads: arr, dec: detectDecimals(img, left, y) });
    }
    return { left, ticks, out };
  }



  return { readChart, debugReads };
}

let LOCAL_READER = null;
let LOCAL_WORKER = null;
let LOCAL_DISABLED_REASON = "";

async function getLocalChartReader() {
  if (process.env.AAA_LOCAL_CHARTS === "0") {
    return null;
  }

  if (LOCAL_DISABLED_REASON) {
    return null;
  }

  if (LOCAL_READER && LOCAL_WORKER) {
    return { reader: LOCAL_READER, worker: LOCAL_WORKER };
  }

  try {
    require.resolve("pngjs");
    require.resolve("tesseract.js");
    require.resolve("@tesseract.js-data/eng");
  } catch {
    LOCAL_DISABLED_REASON = "brak pakietów";
    console.log("");
    console.log("Odczyt wartości z wykresów (lokalny) jest wyłączony — brak pakietów.");
    console.log("Zainstaluj je raz poleceniem:");
    console.log("  npm install pngjs tesseract.js @tesseract.js-data/eng");
    console.log("");
    return null;
  }

  try {
    const { createWorker } = require("tesseract.js");
    const eng = require("@tesseract.js-data/eng");

    LOCAL_READER = buildChartReader();
    LOCAL_WORKER = await createWorker("eng", 1, {
      langPath: eng.langPath,
      gzip: eng.gzip
    });

    return { reader: LOCAL_READER, worker: LOCAL_WORKER };
  } catch (error) {
    LOCAL_DISABLED_REASON = "błąd inicjalizacji";
    console.log(`Odczyt lokalny wykresów niedostępny: ${cleanText(error.message)}`);
    return null;
  }
}

async function shutdownChartReader() {
  if (LOCAL_WORKER) {
    try {
      await LOCAL_WORKER.terminate();
    } catch {}
    LOCAL_WORKER = null;
  }
}

async function readChartsLocally(chartFiles) {
  const ctx = await getLocalChartReader();
  if (!ctx) return chartFiles.map(() => null);

  const out = [];
  for (const file of chartFiles) {
    try {
      out.push(await ctx.reader.readChart(fs.readFileSync(file), ctx.worker));
    } catch (error) {
      out.push({ ok: false, reason: cleanText(error.message), series: [], thresholds: [] });
    }
  }
  return out;
}

// --- formatowanie wyników odczytu -------------------------------------------

// Format dashboardu: "9/24/2026 6:53:00 PM" (godzina może być pominięta).
function parseUsDateTime(value) {
  const m = String(value || "").match(
    /(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AP]M)?)?/i
  );
  if (!m) return null;
  let hour = Number(m[4] || 0);
  const ampm = String(m[7] || "").toUpperCase();
  if (ampm === "PM" && hour < 12) hour += 12;
  if (ampm === "AM" && hour === 12) hour = 0;
  return new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]), hour, Number(m[5] || 0), Number(m[6] || 0));
}

function fmtPlDateTime(date) {
  const p = n => String(n).padStart(2, "0");
  return `${p(date.getDate())}.${p(date.getMonth() + 1)}.${date.getFullYear()} ${p(date.getHours())}:${p(date.getMinutes())}`;
}

// Czas z osi X wykresu: oś przyjęta jako liniowa na zakres dat AAA, błąd ±1–2 h (CLAUDE.md 3, 5.1).
function approxTimeText(dateRange, fraction) {
  const a = parseUsDateTime(dateRange && dateRange.start);
  const b = parseUsDateTime(dateRange && dateRange.end);
  if (!a || !b || fraction === undefined || fraction === null) return "";
  const t = new Date(a.getTime() + Math.max(0, Math.min(1, fraction)) * (b.getTime() - a.getTime()));
  const p = n => String(n).padStart(2, "0");
  return ` (ok. ${p(t.getDate())}.${p(t.getMonth() + 1)} ${p(t.getHours())}:00)`;
}

function rangeHours(dateRange) {
  const a = parseUsDateTime(dateRange && dateRange.start);
  const b = parseUsDateTime(dateRange && dateRange.end);
  return a && b && b > a ? (b - a) / 3.6e6 : null;
}

function gapDurationText(dateRange, gap) {
  const h = rangeHours(dateRange);
  return h ? ` (brak danych ok. ${Math.max(1, Math.round((gap.to - gap.from) * h))} h)` : "";
}

function fmtChartNum(n, resolution) {
  if (n === null || n === undefined || !Number.isFinite(n)) return "?";
  const places = Math.min(3, Math.max(0, Math.round(-Math.log10(resolution || 1))));
  return n.toFixed(places).replace(".", ",");
}

function signedChartNum(n, resolution) {
  const text = fmtChartNum(n, resolution);
  return n > 0 ? `+${text}` : text;
}

function describeLocalChart(local, dateRange) {
  const lines = [];
  const findings = [];
  if (!local || !local.ok) return { lines, findings };

  const res = local.axis.resolution;
  const f = n => fmtChartNum(n, res);
  const sf = n => signedChartNum(n, res);
  const pct = x => `${Math.round(x * 100)}%`;
  const shortTitle = String(local.title || "wykres").replace(/\s*--?\s*[A-Z]?\d{5,}\s*$/i, "").trim();

  lines.push(
    `Oś Y: ${f(local.axis.bottom)} – ${f(local.axis.top)}, dokładność odczytu ±${f(res)}.`
  );

  for (const t of local.thresholds) {
    lines.push(`Linia pozioma (${t.color}): ${f(t.value)}`);
  }

  for (const s of local.series) {
    let line =
      `Seria ${s.color}: min ${f(s.min)}${approxTimeText(dateRange, s.minAt)}, ` +
      `maks ${f(s.max)}${approxTimeText(dateRange, s.maxAt)}, ` +
      `średnia ${f(s.avg)}, ostatnia ${f(s.last)}, ` +
      `trend ${sf(s.trend)} (ostatnia ćwiartka vs pierwsza)`;

    if (s.coverage < 0.6) line += " — seria częściowo zasłonięta inną, zakres szacunkowy";
    if (s.gaps && s.gaps.length) {
      line += ` — luki w danych usunięte ze statystyk: ${s.gaps
        .map(g => `${pct(g.from)}–${pct(g.to)} osi${gapDurationText(dateRange, g)}`)
        .join(", ")}`;
    }
    if (s.suspectedGaps && s.suspectedGaps.length) {
      line += ` — możliwe luki (odcinki liniowe w gładkiej serii, nie usunięto): ${s.suspectedGaps
        .map(g => `${pct(g.from)}–${pct(g.to)} osi`)
        .join(", ")}`;
    }
    lines.push(line);

    for (const v of s.versus) {
      lines.push(
        `   względem linii ${f(v.threshold)} (${v.color}): powyżej ${pct(v.above)} czasu, poniżej ${pct(v.below)}; ` +
        `ostatnia ${sf(v.lastVsThreshold)}, maks ${sf(v.marginToMax)}, min ${sf(v.marginToMin)}`
      );
    }

    // Wnioski: dwie linie tego samego koloru = pasmo (np. min/max), jedna linia = próg.
    const handled = new Set();
    for (const v of s.versus) {
      if (handled.has(v.color)) continue;
      const same = s.versus.filter(x => x.color === v.color);

      if (same.length >= 2) {
        handled.add(v.color);
        const hi = same.reduce((a, b) => (b.threshold > a.threshold ? b : a));
        const lo = same.reduce((a, b) => (b.threshold < a.threshold ? b : a));

        if (hi.above >= 0.01) {
          findings.push(
            `${shortTitle}: seria ${s.color} wychodzi powyżej górnej granicy ${f(hi.threshold)} w ${pct(hi.above)} czasu ` +
            `(maks ${f(s.max)}, ostatnia ${f(s.last)}).`
          );
        } else if (lo.below >= 0.01) {
          findings.push(
            `${shortTitle}: seria ${s.color} wychodzi poniżej dolnej granicy ${f(lo.threshold)} w ${pct(lo.below)} czasu ` +
            `(min ${f(s.min)}, ostatnia ${f(s.last)}).`
          );
        } else {
          findings.push(
            `${shortTitle}: seria ${s.color} w całości w pasmie ${f(lo.threshold)} – ${f(hi.threshold)} ` +
            `(min ${f(s.min)}, maks ${f(s.max)}).`
          );
        }
        continue;
      }

      handled.add(v.color);

      if (v.above >= 0.02 && v.above <= 0.98) {
        findings.push(
          `${shortTitle}: seria ${s.color} przekracza linię ${f(v.threshold)} w ${pct(v.above)} czasu ` +
          `(maks ${f(s.max)}, ostatnia ${f(s.last)}).`
        );
      } else if (v.above > 0.98) {
        findings.push(`${shortTitle}: seria ${s.color} cały czas powyżej linii ${f(v.threshold)} (ostatnia ${f(s.last)}).`);
      } else if (v.below > 0.98) {
        findings.push(
          `${shortTitle}: seria ${s.color} cały czas poniżej linii ${f(v.threshold)} ` +
          `(maks ${f(s.max)}, zapas ${f(Math.abs(v.marginToMax))}).`
        );
      }
    }
  }

  if (local.marker) {
    const ys = local.thresholds.filter(t => t.color === "żółta").map(t => t.value).sort((a, b) => a - b);
    let verdict = "";
    if (ys.length >= 2) {
      const lo = ys[0], hi = ys[ys.length - 1];
      verdict = local.marker.value >= lo && local.marker.value <= hi
        ? `w zakresie żółtych progów ${f(lo)} – ${f(hi)}`
        : `POZA zakresem żółtych progów ${f(lo)} – ${f(hi)}`;
    }
    lines.push(`Znacznik tego analizatora (czerwony): ${f(local.marker.value)} ${verdict}`.trim());
    findings.push(`${shortTitle}: wartość tego analizatora ${f(local.marker.value)} ${verdict}`.trim() + ".");
  }

  return { lines, findings };
}

// ============================================================
// CHART VALUES
// 1) DOM: tooltipy/alt/title z map obrazów, tabele, SVG text
// 2) OPCJONALNIE (opt-in): odczyt z obrazu przez API Anthropic
//    AAA_VISION=1 oraz ANTHROPIC_API_KEY w środowisku
// ============================================================

function parseNumbersFromText(text) {
  const out = [];
  const rx = /-?\d{1,3}(?:[ \u00A0]\d{3})*(?:[.,]\d+)?|-?\d+(?:[.,]\d+)?/g;
  for (const m of String(text || "").matchAll(rx)) {
    const n = Number(m[0].replace(/[ \u00A0]/g, "").replace(",", "."));
    if (Number.isFinite(n)) out.push(n);
  }
  return out;
}

function summarizePoints(points) {
  const nums = points
    .map(p => (typeof p.value === "number" ? p.value : null))
    .filter(v => v !== null);

  if (!nums.length) {
    return { count: points.length };
  }

  const sum = nums.reduce((a, b) => a + b, 0);

  return {
    count: nums.length,
    min: Math.min(...nums),
    max: Math.max(...nums),
    avg: Math.round((sum / nums.length) * 100) / 100,
    last: nums[nums.length - 1]
  };
}

async function extractChartDomData(page) {
  const raw = await page.evaluate(() => {
    const clean = s => String(s || "").replace(/\s+/g, " ").trim();
    const charts = [];

    const nearestTitle = el => {
      let node = el;
      for (let depth = 0; depth < 6 && node; depth++) {
        let prev = node.previousElementSibling;
        for (let hop = 0; hop < 4 && prev; hop++) {
          const t = clean(prev.innerText || prev.textContent);
          if (t && t.length <= 160) return t;
          prev = prev.previousElementSibling;
        }
        node = node.parentElement;
      }
      return "";
    };

    for (const img of document.querySelectorAll("img")) {
      const src = img.getAttribute("src") || img.src || "";
      if (!/chartaxd\.axd/i.test(src)) continue;

      const entry = {
        src: img.src || src,
        title: clean(img.getAttribute("title") || img.getAttribute("alt")) || nearestTitle(img),
        areas: [],
        tableRows: []
      };

      const useMap = (img.getAttribute("usemap") || "").replace(/^#/, "");
      if (useMap) {
        const map = [...document.querySelectorAll("map")].find(
          m => m.name === useMap || m.id === useMap
        );
        if (map) {
          for (const area of map.querySelectorAll("area")) {
            const label = clean(
              area.getAttribute("title") ||
              area.getAttribute("alt") ||
              area.getAttribute("data-tooltip") ||
              area.getAttribute("data-title") ||
              area.getAttribute("onmouseover") ||
              ""
            );
            if (label) entry.areas.push(label);
          }
        }
      }

      // Tabele z liczbami w pobliżu wykresu (ten sam kontener).
      const container = img.closest("table, div, td") || img.parentElement;
      const scope = container && container.parentElement ? container.parentElement : container;
      if (scope) {
        for (const tr of scope.querySelectorAll("table tr")) {
          const cells = [...tr.querySelectorAll("th,td")].map(c => clean(c.innerText || c.textContent));
          if (cells.length >= 2 && cells.some(c => /\d/.test(c)) && cells.join(" ").length < 220) {
            entry.tableRows.push(cells);
          }
          if (entry.tableRows.length >= 80) break;
        }
      }

      charts.push(entry);
    }

    // Wykresy rysowane jako SVG (jeśli strona ich używa)
    for (const svg of document.querySelectorAll("svg")) {
      const texts = [...svg.querySelectorAll("text")]
        .map(t => clean(t.textContent))
        .filter(Boolean);
      if (texts.length >= 3 && texts.some(t => /\d/.test(t))) {
        charts.push({
          src: "svg",
          title: nearestTitle(svg),
          areas: texts.slice(0, 200),
          tableRows: []
        });
      }
    }

    return charts;
  }).catch(() => []);

  return raw.map((chart, index) => {
    const points = [];

    for (const label of chart.areas) {
      const nums = parseNumbersFromText(label);
      points.push({
        label: cleanText(label).slice(0, 140),
        value: nums.length ? nums[nums.length - 1] : null,
        all: nums.slice(0, 6)
      });
    }

    for (const row of chart.tableRows) {
      const nums = row.slice(1).flatMap(parseNumbersFromText);
      points.push({
        label: cleanText(row.join(" | ")).slice(0, 140),
        value: nums.length ? nums[0] : null,
        all: nums.slice(0, 6)
      });
    }

    return {
      index,
      source: "DOM",
      url: chart.src,
      title: cleanText(chart.title),
      points: points.slice(0, 300),
      summary: summarizePoints(points)
    };
  });
}

async function readChartWithVision(file) {
  const key = process.env.ANTHROPIC_API_KEY;
  if (process.env.AAA_VISION !== "1" || !key) return null;

  if (typeof fetch !== "function") {
    console.log("AAA_VISION: wymaga Node 18+ (brak fetch).");
    return null;
  }

  try {
    const data = fs.readFileSync(file).toString("base64");

    const response = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "content-type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify({
        model: process.env.AAA_VISION_MODEL || "claude-sonnet-5-5",
        max_tokens: 1200,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: "image/png", data } },
            {
              type: "text",
              text:
                "To jest wykres diagnostyczny z analizatora. Odczytaj: tytuł, oś X (zakres, jednostka), " +
                "oś Y (zakres, jednostka), serie z legendy oraz wartości: minimum, maksimum, ostatnia " +
                "wartość i wszystkie wartości opisane na wykresie. Jeśli czegoś nie da się odczytać " +
                "pewnie, wpisz null — nie zgaduj. Odpowiedz WYŁĄCZNIE JSON: " +
                '{"title":"","x_axis":"","y_axis":"","series":[{"name":"","min":null,"max":null,"last":null,"labeled_values":[]}],"notes":""}'
            }
          ]
        }]
      })
    });

    if (!response.ok) {
      console.log(`AAA_VISION: błąd API ${response.status}`);
      return null;
    }

    const json = await response.json();
    const text = (json.content || [])
      .filter(b => b.type === "text")
      .map(b => b.text)
      .join("\n")
      .replace(/```json|```/g, "")
      .trim();

    return JSON.parse(text);
  } catch (error) {
    console.log(`AAA_VISION: ${cleanText(error.message)}`);
    return null;
  }
}

async function collectChartData(page, chartFiles) {
  const dom = await extractChartDomData(page);
  const result = [];
  const localReads = await readChartsLocally(chartFiles);

  for (let i = 0; i < chartFiles.length; i++) {
    const file = chartFiles[i];
    const url = CHART_URL_BY_FILE.get(file) || "";

    const match =
      dom.find(c => c.url && url && c.url === url) ||
      dom.find(c => c.url && url && c.url.split("?")[0] === url.split("?")[0] && c.points.length) ||
      null;

    const entry = {
      file: path.basename(file),
      title: match ? match.title : "",
      source: match && match.points.length ? "DOM" : "BRAK",
      points: match ? match.points : [],
      summary: match ? match.summary : { count: 0 },
      local: localReads[i] || null,
      vision: null
    };

    if (entry.local && entry.local.ok && entry.local.title) {
      entry.title = cleanText(entry.local.title);
    }

    const vision = await readChartWithVision(file);
    if (vision) {
      entry.vision = vision;
      if (!entry.title && vision.title) entry.title = cleanText(vision.title);
    }

    result.push(entry);
  }

  return result;
}

// ============================================================
// ANALYSIS (CLAUDE.md, rozdz. 3, 5.2, 6, 7)
// - każde zdanie ma status: Fakt / Odczyt / Hipoteza / Zalecenie
// - priorytet wynika z dowodów, nie z koloru alertu
// - progi zawsze z konkretnego wykresu, nie z założeń
// ============================================================

const PRIORITY_ORDER = ["WYSOKI", "ŚREDNI", "ZDALNIE", "NISKI"];

const PRIORITY_TERM = {
  WYSOKI: "najbliższa możliwa wizyta",
  "ŚREDNI": "zdalna weryfikacja / planowa wizyta",
  ZDALNIE: "od razu, zdalnie",
  NISKI: "obserwacja"
};

function fmtN(n, places = 1) {
  return Number.isFinite(n) ? n.toFixed(places).replace(".", ",") : "?";
}

function plural(n, one, few, many) {
  const word = n === 1 ? one : (n % 10 >= 2 && n % 10 <= 4 && (n % 100 < 12 || n % 100 > 14)) ? few : many;
  return `${n} ${word}`;
}

function fmtPct(share) {
  return `${fmtN(share * 100, 0)}%`;
}

function fmtHours(h) {
  if (!Number.isFinite(h)) return "";
  return h >= 1 ? `${fmtN(h, h >= 10 ? 0 : 1)} h` : `${Math.max(1, Math.round(h * 60))} min`;
}

// Tytuły pochodzą z OCR, więc dopasowanie jest luźne (np. "Buty Cycle" = "Duty Cycle").
function chartKind(title) {
  const t = String(title || "").toLowerCase();
  if (/diff|erence/.test(t)) return "thermDiff";
  if (/duty|cycle/.test(t)) return "duty";
  if (/ambi|bient/.test(t)) return "ambient";
  if (/therm|mistor/.test(t)) return "thermistors";
  if (/slot|correct/.test(t)) return "slotCorr";
  if (/step|loss/.test(t)) return "stepLoss";
  if (/stopp/.test(t)) return "cmStopping";
  if (/sync/.test(t)) return "readSync";
  return "";
}

const CHART_KIND_LABEL = {
  duty: "Duty Cycle",
  ambient: "Ambient Temperature",
  thermistors: "Thermistors",
  thermDiff: "Thermistor Difference",
  slotCorr: "Slot Corrections",
  stepLoss: "Step Loss",
  cmStopping: "CM RING Stopping",
  readSync: "READ SYNC"
};

function alertFamily(alertName) {
  const n = String(alertName || "").toLowerCase();
  if (/data\s*logger/.test(n)) return "dataLogger";
  if (/thermal/.test(n)) return "supplyThermal";
  if (/ring/.test(n) && /(cm|slide)/.test(n)) return "cmRing";
  return "generic";
}

function editDistance(a, b) {
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) d[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
  }
  return d[a.length][b.length];
}

// Walidacja odczytu wykresu (CLAUDE.md 5.2). J-number w tytule pochodzi z OCR:
// zgodny = identyczny; 1 znak różnicy = prawdopodobnie błąd OCR (odczyt niepewny);
// ≥ 2 znaki różnicy = wykres innego aparatu.
function titleJnoCheck(title, jno) {
  const want = String(jno || "").replace(/\D/g, "");
  const tail = String(title || "").split(/\s[-—–]+\s|J(?=[\dOoIl(])/).pop() || "";
  const found = String(title || "").match(/\d{5,}/g) || [];
  const tailDigits = tail
    .replace(/[OoQD()¢]/g, "0").replace(/[Il|!]/g, "1").replace(/S/g, "5").replace(/B/g, "8")
    .replace(/\D/g, "");
  if (tailDigits.length >= 5 && !found.includes(tailDigits)) found.push(tailDigits);
  if (!want || !found.length) return { ok: null, found: found.join(", "), fuzzy: false };
  if (found.includes(want)) return { ok: true, found: found.join(", "), fuzzy: false };
  const dist = Math.min(...found.map(f => editDistance(f, want)));
  return dist <= 1
    ? { ok: true, found: found.join(", "), fuzzy: true }
    : { ok: false, found: found.join(", "), fuzzy: false };
}

function validateCharts(result) {
  const usable = [];

  for (const c of result.chartData || []) {
    const local = c.local;
    const title = c.title || (local && local.title) || "";
    c.kind = chartKind(title);
    c.titleCheck = titleJnoCheck(title, result.jno);

    if (!local || !local.ok) {
      c.validity = "brak odczytu";
      continue;
    }

    const problems = [];
    const { top, bottom, resolution, ocrAgreement } = local.axis;

    if (c.titleCheck.ok === false) problems.push(`numer w tytule (${c.titleCheck.found}) inny niż J${result.jno}`);
    if (c.titleCheck.ok === null) problems.push("nie odczytano J-number z tytułu");
    if (c.titleCheck.fuzzy) problems.push(`J-number w tytule odczytany jako ${c.titleCheck.found} (1 znak różnicy, prawdopodobnie błąd OCR)`);
    if (ocrAgreement < 3) problems.push("skala osi Y potwierdzona przez mało etykiet");
    if (local.thresholds.some(t => t.value < bottom - resolution || t.value > top + resolution)) {
      problems.push("próg poza zakresem osi");
    }
    if (local.series.length && local.series.every(s => s.pixels < 60)) problems.push("mało pikseli serii");

    c.validity = problems.length ? `odczyt niepewny: ${problems.join("; ")}` : "wiarygodny";

    // Wykres innego aparatu nie wchodzi do analizy.
    if (c.titleCheck.ok === false) continue;

    usable.push({ ...c, title, uncertain: problems.length > 0 });
  }

  return usable;
}

function pickChart(charts, kind) {
  return charts.find(c => c.kind === kind && c.local.series.length) ||
    charts.find(c => c.kind === kind) ||
    null;
}

function mainSeries(chart) {
  return (chart && chart.local.series[0]) || null;
}

function uncertainTag(chart) {
  return chart && chart.uncertain ? " (odczyt niepewny)" : "";
}

// Wartość serii w punkcie osi X (najbliższa próbka w promieniu 1% szerokości).
function sampleAt(series, f) {
  if (!series || !series.samples || !series.samples.length) return null;
  const arr = series.samples;
  let lo = 0, hi = arr.length - 1;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    if (arr[mid][0] < f) lo = mid + 1; else hi = mid;
  }
  let best = arr[lo];
  if (lo > 0 && Math.abs(arr[lo - 1][0] - f) < Math.abs(best[0] - f)) best = arr[lo - 1];
  return Math.abs(best[0] - f) <= 0.01 ? best[1] : null;
}

// Kolejne próbki spełniające warunek, bez przerw dłuższych niż maxStep.
function episodes(samples, pred, maxStep = 0.01) {
  const out = [];
  let cur = null;
  for (const [f, v] of samples) {
    if (pred(v)) {
      if (cur && f - cur.to <= maxStep) { cur.to = f; cur.n++; }
      else { cur = { from: f, to: f, n: 1 }; out.push(cur); }
    } else {
      cur = null;
    }
  }
  return out;
}

function mean(values) {
  return values.length ? values.reduce((a, b) => a + b, 0) / values.length : NaN;
}

function correlation(pairs) {
  if (pairs.length < 10) return null;
  const mx = mean(pairs.map(p => p[0])), my = mean(pairs.map(p => p[1]));
  let sxy = 0, sxx = 0, syy = 0;
  for (const [x, y] of pairs) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; syy += (y - my) ** 2; }
  return sxx && syy ? sxy / Math.sqrt(sxx * syy) : null;
}

function linearFit(pairs) {
  if (pairs.length < 20) return null;
  const xs = pairs.map(p => p[0]);
  const mx = mean(xs), my = mean(pairs.map(p => p[1]));
  let sxy = 0, sxx = 0;
  for (const [x, y] of pairs) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; }
  if (!sxx) return null;
  const b = sxy / sxx;
  return { a: my - b * mx, b, xMin: Math.min(...xs), xMax: Math.max(...xs), n: pairs.length };
}

// Punkt podziału maksymalizujący różnicę średnich (skok poziomu, CLAUDE.md 6.2).
function levelShift(samples, minShare = 0.1) {
  const n = samples.length;
  const m = Math.max(5, Math.floor(n * minShare));
  if (n < 2 * m) return null;
  const prefix = [0];
  for (const [, v] of samples) prefix.push(prefix[prefix.length - 1] + v);
  let best = null;
  for (let k = m; k <= n - m; k++) {
    const before = prefix[k] / k;
    const after = (prefix[n] - prefix[k]) / (n - k);
    if (!best || Math.abs(after - before) > Math.abs(best.after - best.before)) {
      best = { at: samples[k][0], before, after };
    }
  }
  return best;
}

function newAnalysis(result, family) {
  return {
    family,
    priority: "",
    priorityReason: "",
    problem: "",
    missing: [],
    readings: [],      // { param, value, ref, assessment }
    observations: [],  // { status: "Fakt" | "Odczyt", text }
    hypotheses: [],    // tekst zaczynający się od "Hipoteza:"
    actions: [],       // { text, criterion, mode }
    dutyFit: null
  };
}

function checkCompleteness(result, A, requiredKinds, charts) {
  if (!result.alertName || result.alertName === "Alert AAA") A.missing.push("typ alertu (nie odczytano nazwy)");
  if (!result.jno) A.missing.push("J-number");
  if (!result.dateRange || !result.dateRange.start) A.missing.push("zakres dat AAA");
  if (!(result.genericEvidence && result.genericEvidence.rules.length)) A.missing.push("reguły/progi w tekście AAA");
  if (!result.chartFiles.length) A.missing.push("wykresy AAA");
  for (const kind of requiredKinds) {
    if (!pickChart(charts, kind)) A.missing.push(`wykres ${CHART_KIND_LABEL[kind]} (brak lub brak pewnego odczytu)`);
  }
}

function addDocsActions(result, A) {
  for (const d of (result.serviceDocuments || []).slice(0, 3)) {
    A.actions.push({ text: `Sprawdzić dokument: ${d.title}.`, criterion: "", mode: "na miejscu" });
  }
}

// ------------------------------------------------------------
// 6.1 Supply thermal (VITROS 4600)
// ------------------------------------------------------------

function analyzeSupplyThermal(result, charts) {
  const A = newAnalysis(result, "supplyThermal");
  A.problem = `Aktywny alert: ${result.alertName}.`;
  checkCompleteness(result, A, ["duty", "ambient", "thermistors", "thermDiff"], charts);

  const H = rangeHours(result.dateRange);
  const duty = pickChart(charts, "duty");
  const amb = pickChart(charts, "ambient");
  const th = pickChart(charts, "thermistors");
  const diff = pickChart(charts, "thermDiff");

  let satShare = 0, satHours = null, tempRise = null, aboveShare = null, ambientOut = false, dutyThr = null;

  if (duty && mainSeries(duty)) {
    const s = mainSeries(duty), res = duty.local.axis.resolution, tag = uncertainTag(duty);
    const vals = s.samples.map(x => x[1]);
    const thr = duty.local.thresholds.map(t => t.value).filter(v => v < 99).sort((a, b) => b - a)[0];
    dutyThr = Number.isFinite(thr) ? thr : null;

    A.readings.push({ param: "Duty Cycle — średnia", value: `${fmtN(s.avg)}%`, ref: `±${fmtN(res, 2)} pp`, assessment: "odczyt" + tag });

    if (dutyThr !== null) {
      aboveShare = vals.filter(v => v > dutyThr + res).length / vals.length;
      A.readings.push({
        param: "Duty Cycle powyżej poziomu pożądanego",
        value: fmtPct(aboveShare),
        ref: `${fmtN(dutyThr)}% (próg z wykresu)`,
        assessment: aboveShare >= 0.1 ? "odchylenie" : "w normie"
      });
      A.observations.push({ status: "Odczyt", text: `Duty powyżej ${fmtN(dutyThr)}% przez ${fmtPct(aboveShare)} czasu, średnio ${fmtN(s.avg)}% (±${fmtN(res, 2)} pp)${tag}.` });
    } else {
      A.observations.push({ status: "Odczyt", text: `Duty średnio ${fmtN(s.avg)}%${tag}. Brak pewnego odczytu poziomu pożądanego z wykresu.` });
    }

    A.observations.push({ status: "Odczyt", text: `Trend duty: ${s.trend > 0 ? "+" : ""}${fmtN(s.trend)} pp (ostatnia ćwiartka vs pierwsza).` });

    const sat = episodes(s.samples, v => v >= 99);
    satShare = sat.reduce((n, e) => n + e.n, 0) / vals.length;
    satHours = H ? satShare * H * (1 - (s.gapShare || 0)) : null;
    const longest = sat.reduce((b, e) => (e.to - e.from > b.to - b.from ? e : b), { from: 0, to: 0 });
    A.readings.push({
      param: "Nasycenie duty (≥ 99%)",
      value: sat.length ? `${fmtPct(satShare)} czasu${satHours !== null ? `, ok. ${fmtHours(satHours)}` : ""}` : "brak",
      ref: plural(sat.length, "epizod", "epizody", "epizodów"),
      assessment: sat.length ? "praca na granicy" : "w normie"
    });
    if (sat.length) {
      A.observations.push({
        status: "Odczyt",
        text: `Duty sięga ≥ 99% w ${sat.length} ${sat.length === 1 ? "epizodzie" : "epizodach"}, łącznie ${fmtPct(satShare)} czasu` +
          `${satHours !== null ? ` (ok. ${fmtHours(satHours)})` : ""}; najdłuższy epizod` +
          `${H ? ` ok. ${fmtHours((longest.to - longest.from) * H)}` : ""}${approxTimeText(result.dateRange, longest.from)}.`
      });
    } else {
      A.observations.push({ status: "Odczyt", text: "W tych danych nie widać nasycenia duty (≥ 99%)." });
    }

    // 3. Temperatura Supply 3 w nasyceniu vs poza nim.
    if (th && th.local.series.length && sat.length) {
      const inSat = f => sat.some(e => f >= e.from - 0.002 && f <= e.to + 0.002);
      const s0 = th.local.series[0], others = th.local.series.slice(1);
      const inside = [], outside = [];
      for (const [f, v] of s0.samples) {
        const vs = [v, ...others.map(o => sampleAt(o, f)).filter(x => x !== null)];
        (inSat(f) ? inside : outside).push(mean(vs));
      }
      if (inside.length >= 3 && outside.length >= 3) {
        const resT = th.local.axis.resolution;
        tempRise = mean(inside) - mean(outside);
        A.readings.push({
          param: "Supply 3: średnia w nasyceniu vs poza",
          value: `${fmtN(mean(inside), 2)} vs ${fmtN(mean(outside), 2)} °C`,
          ref: `szczyt ${fmtN(Math.max(...inside), 2)} vs ${fmtN(Math.max(...outside), 2)} °C, ±${fmtN(resT, 2)}`,
          assessment: tempRise > 2 * resT ? "temperatura rośnie w nasyceniu" : "bez wyraźnej różnicy"
        });
        A.observations.push({
          status: "Odczyt",
          text: `W okresach nasycenia temperatura Supply 3 jest średnio ${tempRise >= 0 ? "wyższa" : "niższa"} o ${fmtN(Math.abs(tempRise), 2)} °C niż poza nimi${uncertainTag(th)}.`
        });
        if (tempRise <= 2 * resT) tempRise = null;
      }
    }
  }

  // 4. Otoczenie vs pasmo.
  if (amb && mainSeries(amb)) {
    const s = mainSeries(amb), res = amb.local.axis.resolution;
    const thr = amb.local.thresholds.map(t => t.value).sort((a, b) => a - b);
    if (thr.length >= 2) {
      const lo = thr[0], hi = thr[thr.length - 1];
      const out = s.samples.filter(([, v]) => v > hi + res || v < lo - res).length / s.samples.length;
      const lastOut = s.last > hi + res || s.last < lo - res;
      ambientOut = out > 0;
      A.readings.push({
        param: "Otoczenie poza pasmem",
        value: fmtPct(out),
        ref: `${fmtN(lo)}–${fmtN(hi)} °C (z wykresu)`,
        assessment: lastOut ? "przekroczenie na końcu okna" : out ? "przekroczenia" : "w paśmie"
      });
      A.observations.push({
        status: "Odczyt",
        text: out
          ? `Otoczenie poza pasmem ${fmtN(lo)}–${fmtN(hi)} °C przez ${fmtPct(out)} czasu (maks ${fmtN(s.max)} °C${approxTimeText(result.dateRange, s.maxAt)})${lastOut ? "; przekroczenie trwa na końcu okna" : ""}.`
          : `Otoczenie w paśmie ${fmtN(lo)}–${fmtN(hi)} °C (min ${fmtN(s.min)}, maks ${fmtN(s.max)} °C).`
      });
    } else {
      A.observations.push({ status: "Odczyt", text: `Otoczenie: min ${fmtN(s.min)}, maks ${fmtN(s.max)} °C. Brak pewnego odczytu pasma z wykresu.` });
    }
    A.observations.push({ status: "Odczyt", text: `Trend otoczenia: ${s.trend > 0 ? "+" : ""}${fmtN(s.trend)} °C (ostatnia ćwiartka vs pierwsza).` });
  }

  // 5. Różnica termistorów vs trigger.
  let diffMarginLow = false;
  if (diff && mainSeries(diff)) {
    const s = mainSeries(diff);
    const trigger = Math.max(...diff.local.thresholds.map(t => Math.abs(t.value)));
    const maxAbs = Math.max(Math.abs(s.min), Math.abs(s.max));
    if (Number.isFinite(trigger) && trigger > 0) {
      const margin = trigger - maxAbs;
      diffMarginLow = margin < 0.25 * trigger;
      A.readings.push({
        param: "Różnica termistorów — maks.",
        value: fmtN(maxAbs, 2),
        ref: `trigger ${fmtN(trigger, 2)} (z wykresu)`,
        assessment: `zapas ≈ ${fmtN(margin, 2)}`
      });
      A.observations.push({ status: "Odczyt", text: `Różnica termistorów nie przekracza ${fmtN(maxAbs, 2)} przy triggerze ${fmtN(trigger, 2)} (zapas ≈ ${fmtN(margin, 2)})${uncertainTag(diff)}.` });
    } else {
      A.observations.push({ status: "Odczyt", text: `Różnica termistorów maks. ${fmtN(maxAbs, 2)}. Brak pewnego odczytu triggera z wykresu.` });
    }
  }

  // 6. Korelacja duty–otoczenie i dopasowanie do porównania (7).
  if (duty && amb && mainSeries(duty) && mainSeries(amb)) {
    const pairs = mainSeries(duty).samples
      .map(([f, d]) => [sampleAt(mainSeries(amb), f), d])
      .filter(p => p[0] !== null);
    const r = correlation(pairs);
    if (r !== null) {
      A.observations.push({ status: "Odczyt", text: `Korelacja duty–otoczenie r = ${fmtN(r, 2)} (${pairs.length} par). Korelacja nie dowodzi przyczyny.` });
    }
    const fit = linearFit(pairs.filter(p => p[1] < 99)); // nasycenie spłaszcza zależność
    if (fit && fit.xMax - fit.xMin >= 1) A.dutyFit = fit;
  }

  // Priorytet (rozdz. 7).
  if (!duty || !mainSeries(duty)) {
    A.priority = "ZDALNIE";
    A.priorityReason = "Brak pewnego odczytu Duty Cycle — nie da się ocenić wydajności układu; najpierw uzupełnić dane AAA.";
  } else if (satShare >= 0.05 && tempRise !== null) {
    A.priority = "WYSOKI";
    A.priorityReason = `Duty ≥ 99% przez ${fmtPct(satShare)} czasu${satHours !== null ? ` (ok. ${fmtHours(satHours)})` : ""}, a temperatura Supply 3 rośnie w tych okresach o ${fmtN(tempRise, 2)} °C.`;
  } else if (satShare > 0 || (aboveShare !== null && aboveShare >= 0.1) || ambientOut) {
    A.priority = "ŚREDNI";
    A.priorityReason = satShare > 0
      ? `Duty ≥ 99% przez ${fmtPct(satShare)} czasu, bez wyraźnego wzrostu temperatury Supply 3 w tych okresach.`
      : aboveShare !== null && aboveShare >= 0.1
        ? `Duty powyżej ${fmtN(dutyThr)}% przez ${fmtPct(aboveShare)} czasu, układ kompensuje (brak nasycenia).`
        : "Otoczenie poza pasmem — przyczyna może być zewnętrzna.";
  } else {
    A.priority = "NISKI";
    A.priorityReason = `Duty bez nasycenia${aboveShare !== null ? `, powyżej progu ${fmtPct(aboveShare)} czasu` : ""}; otoczenie w paśmie lub bez pewnego odczytu.`;
  }

  // Hipotezy.
  if (satShare > 0 || (aboveShare !== null && aboveShare >= 0.1)) {
    A.hypotheses.push("Hipoteza: obniżona wydajność chłodzenia (np. zabrudzony radiator lub filtr, wentylatory, ograniczony przepływ powietrza, spadek sprawności elementu chłodzącego, nieszczelność pokrywy). Do weryfikacji na miejscu.");
  }
  if (ambientOut) {
    A.hypotheses.push("Hipoteza: podwyższone obciążenie może wynikać z temperatury otoczenia. Do weryfikacji z laboratorium (klimatyzacja, wentylacja).");
  }
  if (diffMarginLow) {
    A.hypotheses.push("Hipoteza: rozbieżność termistorów zbliża się do triggera — może wskazywać na problem czujnika lub jego połączenia. Do weryfikacji pomiarem.");
  }
  if (!A.hypotheses.length) {
    A.hypotheses.push("W tych danych nie widać wzorca, który uzasadniałby hipotezę usterki. Alert może wynikać z krótkiego epizodu.");
  }

  // Zalecenia.
  const crit = dutyThr !== null
    ? `duty < ${fmtN(dutyThr)}% przy ≈ 27 °C, brak epizodów 100%`
    : "brak epizodów duty 100%";
  if (ambientOut) {
    A.actions.push({ text: "Zadzwonić do laboratorium: sprawdzić temperaturę i wentylację pomieszczenia.", criterion: "otoczenie w paśmie z wykresu AAA", mode: "zdalnie" });
  }
  if (A.priority === "WYSOKI" || A.priority === "ŚREDNI") {
    A.actions.push({ text: "Wykonać Checking Temperatures and Thermal Currents.", criterion: crit, mode: "na miejscu" });
    A.actions.push({ text: "Sprawdzić radiator, filtr, wentylatory, przepływ powietrza i szczelność pokrywy.", criterion: crit, mode: "na miejscu" });
  } else {
    A.actions.push({ text: "Obserwować Duty Cycle w kolejnym oknie AAA.", criterion: crit, mode: "zdalnie" });
  }
  addDocsActions(result, A);
  return A;
}

// ------------------------------------------------------------
// 6.2 Slide (CM) Ring (XT 3400)
// ------------------------------------------------------------

function markerVsBand(chart) {
  const m = chart.local.marker;
  const ys = chart.local.thresholds.filter(t => t.color === "żółta").map(t => t.value).sort((a, b) => a - b);
  if (!m) return null;
  if (ys.length < 2) return { value: m.value, band: null, outsideBy: null };
  const lo = ys[0], hi = ys[ys.length - 1];
  const outsideBy = m.value > hi ? m.value - hi : m.value < lo ? m.value - lo : 0;
  return { value: m.value, band: [lo, hi], outsideBy };
}

function analyzeCmRing(result, charts) {
  const A = newAnalysis(result, "cmRing");
  A.problem = `Aktywny alert: ${result.alertName}.`;
  checkCompleteness(result, A, ["slotCorr", "stepLoss", "cmStopping", "readSync"], charts);

  const slot = pickChart(charts, "slotCorr");
  const step = pickChart(charts, "stepLoss");
  const H = rangeHours(result.dateRange);

  let slotAbove = null, slotThr = null, shift = null, stepMargin = null, stepLimit = null, outsideMarkers = [];

  if (slot && mainSeries(slot)) {
    const s = mainSeries(slot), res = slot.local.axis.resolution, tag = uncertainTag(slot);
    const thr = slot.local.thresholds.map(t => t.value).sort((a, b) => a - b)[0];
    slotThr = Number.isFinite(thr) ? thr : null;
    if (slotThr !== null) {
      slotAbove = s.samples.filter(([, v]) => v > slotThr + res).length / s.samples.length;
      A.readings.push({
        param: "Slot Corrections powyżej poziomu pożądanego",
        value: fmtPct(slotAbove),
        ref: `${fmtN(slotThr)}/h (próg z wykresu)`,
        assessment: slotAbove >= 0.1 ? "odchylenie" : "w normie"
      });
    }
    A.readings.push({ param: "Slot Corrections — średnia / maks.", value: `${fmtN(s.avg)} / ${fmtN(s.max)}`, ref: `±${fmtN(res, 2)}`, assessment: "odczyt" + tag });
    A.observations.push({
      status: "Odczyt",
      text: `Slot Corrections: średnio ${fmtN(s.avg)}, maks ${fmtN(s.max)}${approxTimeText(result.dateRange, s.maxAt)}` +
        `${slotThr !== null ? `; powyżej ${fmtN(slotThr)} przez ${fmtPct(slotAbove)} czasu` : "; brak pewnego odczytu poziomu pożądanego"}${tag}.`
    });

    const ls = levelShift(s.samples);
    if (ls && Math.abs(ls.after - ls.before) >= Math.max(3 * res, 0.2 * Math.abs(ls.before))) {
      shift = ls;
      const nearGap = (s.gaps || []).find(g => ls.at >= g.from - 0.03 && ls.at <= g.to + 0.03);
      A.readings.push({
        param: "Skok poziomu Slot Corrections",
        value: `${fmtN(ls.before)} → ${fmtN(ls.after)}`,
        ref: approxTimeText(result.dateRange, ls.at).trim() || `ok. ${fmtPct(ls.at)} osi`,
        assessment: nearGap ? "pokrywa się z luką w danych" : "bez luki"
      });
      A.observations.push({
        status: "Odczyt",
        text: `Poziom Slot Corrections zmienia się z ${fmtN(ls.before)} na ${fmtN(ls.after)}${approxTimeText(result.dateRange, ls.at)}.` +
          (nearGap ? ` Skok pokrywa się z luką w danych${gapDurationText(result.dateRange, nearGap)} — sugeruje zdarzenie (restart, interwencja), a nie stopniowe zużycie.` : "")
      });
    }
  }

  if (step && mainSeries(step)) {
    const s = mainSeries(step);
    const limit = step.local.thresholds.map(t => t.value).sort((a, b) => a - b)[0];
    if (Number.isFinite(limit)) {
      stepLimit = limit;
      // Zapas liczony od najbardziej ujemnej wartości (CLAUDE.md 6.2.3).
      stepMargin = s.min - limit;
      A.readings.push({
        param: "Step Loss — najgorszy odczyt",
        value: fmtN(s.min),
        ref: `limit ${fmtN(limit)} (z wykresu)`,
        assessment: stepMargin <= 0 ? "limit osiągnięty" : `zapas ${fmtN(stepMargin)}`
      });
      A.observations.push({ status: "Odczyt", text: `Step Loss: najgorszy odczyt ${fmtN(s.min)}${approxTimeText(result.dateRange, s.minAt)}, limit ${fmtN(limit)}, zapas ${fmtN(stepMargin)}${uncertainTag(step)}.` });
    } else {
      A.observations.push({ status: "Odczyt", text: `Step Loss: najgorszy odczyt ${fmtN(s.min)}. Brak pewnego odczytu limitu z wykresu.` });
    }
  }

  for (const kind of ["cmStopping", "readSync"]) {
    const c = pickChart(charts, kind);
    if (!c) continue;
    const mv = markerVsBand(c);
    const label = CHART_KIND_LABEL[kind];
    if (!mv) {
      A.observations.push({ status: "Odczyt", text: `${label}: brak pewnego odczytu znacznika tego aparatu.` });
      continue;
    }
    const inBand = mv.band && mv.outsideBy === 0;
    A.readings.push({
      param: `${label} — wartość tego aparatu`,
      value: fmtN(mv.value),
      ref: mv.band ? `pasmo ${fmtN(mv.band[0])}–${fmtN(mv.band[1])}` : "brak pewnego odczytu pasma",
      assessment: !mv.band ? "?" : inBand ? "w paśmie" : `odstająca o ${fmtN(Math.abs(mv.outsideBy))}`
    });
    if (mv.band && !inBand) outsideMarkers.push(label);
    A.observations.push({
      status: "Odczyt",
      text: `${label}: wartość tego aparatu ${fmtN(mv.value)}` +
        (mv.band ? (inBand ? ` w paśmie ${fmtN(mv.band[0])}–${fmtN(mv.band[1])}.` : `, poza pasmem ${fmtN(mv.band[0])}–${fmtN(mv.band[1])} o ${fmtN(Math.abs(mv.outsideBy))}.`) : ".") +
        " Populacja z pikseli orientacyjna."
    });
  }

  if (!slot && !step) {
    A.priority = "ZDALNIE";
    A.priorityReason = "Brak pewnego odczytu Slot Corrections i Step Loss — najpierw uzupełnić dane AAA.";
  } else if (stepMargin !== null && stepMargin <= 0) {
    A.priority = "WYSOKI";
    A.priorityReason = `Step Loss osiąga limit ${fmtN(stepLimit)} (najgorszy odczyt ${fmtN(stepLimit + stepMargin)}).`;
  } else if (slotAbove !== null && slotAbove >= 0.5 && outsideMarkers.length) {
    A.priority = "WYSOKI";
    A.priorityReason = `Slot Corrections powyżej ${fmtN(slotThr)} przez ${fmtPct(slotAbove)} czasu, a ${outsideMarkers.join(" i ")} poza pasmem.`;
  } else if ((slotAbove !== null && slotAbove >= 0.1) || outsideMarkers.length || (shift && slotThr !== null && shift.after > slotThr)) {
    A.priority = "ŚREDNI";
    A.priorityReason = slotAbove !== null && slotAbove >= 0.1
      ? `Slot Corrections powyżej ${fmtN(slotThr)} przez ${fmtPct(slotAbove)} czasu${stepMargin !== null ? `, Step Loss z zapasem ${fmtN(stepMargin)}` : ""}.`
      : outsideMarkers.length
        ? `${outsideMarkers.join(" i ")} poza pasmem populacji.`
        : `Skok poziomu Slot Corrections do ${fmtN(shift.after)} (powyżej ${fmtN(slotThr)}).`;
  } else {
    A.priority = "NISKI";
    A.priorityReason = `Slot Corrections${slotAbove !== null ? ` powyżej progu ${fmtPct(slotAbove)} czasu` : " bez pewnego progu"}${stepMargin !== null ? `, Step Loss z zapasem ${fmtN(stepMargin)}` : ""}.`;
  }

  if (slotAbove !== null && slotAbove >= 0.1 && (stepMargin === null || stepMargin > 0)) {
    A.hypotheses.push("Hipoteza: aparat osiąga pozycję, ale potrzebuje więcej korekcji — może wskazywać na zwiększony opór lub zużycie mechaniki (CM Rotor belt i tracking, CM/RT Drive Motor i pinion, SAG rollers/bearings/spacers, Z-axis bearing pads). Do weryfikacji na miejscu.");
  }
  if (shift && (mainSeries(slot).gaps || []).some(g => shift.at >= g.from - 0.03 && shift.at <= g.to + 0.03)) {
    A.hypotheses.push("Hipoteza: zmiana poziomu po luce w danych może wynikać z restartu lub interwencji. Do weryfikacji w historii serwisowej.");
  }
  if (outsideMarkers.length) {
    A.hypotheses.push(`Hipoteza: wartość adjustment (${outsideMarkers.join(", ")}) odstająca od populacji może mieć związek z korekcjami. Do weryfikacji po ocenie mechaniki.`);
  }
  if (stepMargin !== null && stepMargin <= 0) {
    A.hypotheses.push("Hipoteza: gubienie kroków przez napęd CM Ring. Do weryfikacji na miejscu.");
  }
  if (!A.hypotheses.length) {
    A.hypotheses.push("W tych danych nie widać wzorca, który uzasadniałby hipotezę usterki.");
  }

  const crit = slotThr !== null ? `Slot Corrections < ${fmtN(slotThr)}/h, brak skoku poziomu` : "brak wzrostu Slot Corrections";
  if (A.priority === "WYSOKI" || A.priority === "ŚREDNI") {
    A.actions.push({ text: "Sprawdzić CM Rotor belt (zużycie, tracking na pinion CM/RT Drive Motor).", criterion: crit, mode: "na miejscu" });
    A.actions.push({ text: "Sprawdzić SAG rollers, bearings, spacers i Z-axis bearing pads.", criterion: crit, mode: "na miejscu" });
    if (outsideMarkers.length) {
      A.actions.push({ text: `Zweryfikować ${outsideMarkers.join(" i ")} dopiero po ocenie mechaniki, zgodnie z procedurą serwisową.`, criterion: "wartość w paśmie żółtych progów", mode: "na miejscu" });
    }
  } else {
    A.actions.push({ text: "Obserwować Slot Corrections i Step Loss w kolejnym oknie AAA.", criterion: crit, mode: "zdalnie" });
  }
  if (shift) {
    A.actions.push({ text: "Sprawdzić historię serwisową i restarty w okolicy skoku poziomu.", criterion: "wyjaśniona przyczyna zmiany poziomu", mode: "zdalnie" });
  }
  addDocsActions(result, A);
  return A;
}

// ------------------------------------------------------------
// 6.3 Data Logger Status
// ------------------------------------------------------------

function analyzeDataLogger(result, now = new Date()) {
  const A = newAnalysis(result, "dataLogger");
  A.problem = "e-Connectivity zgłasza brak lub opóźnienie danych z analizatora.";
  A.priority = "ZDALNIE";

  const dl = result.dataLogger || {};
  const last = parseUsDateTime(dl.lastConnected);
  if (last) {
    const h = Math.max(0, (now - last) / 3.6e6);
    A.elapsedHours = h;
    const d = Math.floor(h / 24), hh = Math.round(h - d * 24);
    A.elapsedText = `≈ ${d} d ${hh} h`;
    A.observations.push({ status: "Fakt", text: `Ostatnie połączenie: ${fmtPlDateTime(last)} (${A.elapsedText} przed raportem).` });
    A.priorityReason = `Brak łączności od ${fmtPlDateTime(last)} (${A.elapsedText}); aparat jest niewidoczny dla alertów technicznych.`;
  } else {
    A.missing.push("Last Connected (brak pewnego odczytu)");
    A.priorityReason = "Brak pewnego odczytu ostatniego połączenia; aparat może być niewidoczny dla alertów technicznych.";
  }

  for (const [label, list] of [["A-file", dl.aFile || []], ["B-file", dl.bFile || []]]) {
    if (!list.length) { A.missing.push(`${label} upload status`); continue; }
    const allZero = list.every(x => /^0\/0$/.test(x.value));
    A.observations.push({
      status: "Fakt",
      text: `${label}: ${list.map(x => `${x.day} ${x.value}`).join(", ")}${allZero ? " — 0/0 przez cały okres, brak transmisji" : ""}.`
    });
  }

  A.observations.push({ status: "Fakt", text: "Sam alert łączności nie dowodzi awarii mechanicznej, ale oznacza brak widoczności alertów technicznych tego aparatu." });
  A.hypotheses.push("Hipoteza: przerwa w zasilaniu, sieci, usłudze Data Logger lub regułach zapory. Do weryfikacji zdalnie.");

  A.actions.push({ text: "Zadzwonić do laboratorium: czy aparat jest włączony i podłączony do sieci.", criterion: "aparat pracuje, kabel sieciowy podłączony", mode: "zdalnie" });
  A.actions.push({ text: "Sprawdzić status usługi e-Connectivity/Data Logger.", criterion: "Last Connected aktualizuje się", mode: "zdalnie" });
  A.actions.push({ text: "Sprawdzić z IT klienta zaporę i sieć.", criterion: "A-file i B-file > 0 w kolejnym dniu", mode: "zdalnie" });
  A.actions.push({ text: "Zaplanować wizytę dopiero po wykluczeniu przyczyn zdalnych.", criterion: "", mode: "na miejscu" });
  return A;
}

// ------------------------------------------------------------
// 6.4 Inne alerty
// ------------------------------------------------------------

function analyzeGeneric(result, charts) {
  const A = newAnalysis(result, "generic");
  A.problem = `Aktywny alert: ${result.alertName}. Brak dedykowanych reguł dla tego typu alertu.`;
  checkCompleteness(result, A, [], charts);

  const ev = result.genericEvidence || {};
  A.observations.push({ status: "Fakt", text: `AAA: ${plural(ev.rules ? ev.rules.length : 0, "linia", "linie", "linii")} z regułami/parametrami; condition codes: ${ev.conditionCodes && ev.conditionCodes.length ? ev.conditionCodes.join(", ") : "brak pewnego odczytu"}.` });

  let exceed = 0;
  for (const c of charts) {
    for (const line of describeLocalChart(c.local, result.dateRange).findings) {
      A.observations.push({ status: "Odczyt", text: line + uncertainTag(c) });
    }
    for (const s of c.local.series) if (s.versus.some(v => v.above >= 0.02 && v.above <= 0.98)) exceed++;
  }

  const highlighted = aaaContentFacts(result);
  const hasContent = (result.aaaContent || []).some(s => s.blocks.length);

  if (highlighted.length) {
    // Wiersz wyróżniony przez AAA (np. „Investigate”) to odchylenie widoczne w danych źródłowych.
    A.priority = "ŚREDNI";
    A.priorityReason = highlighted[0].replace(/^AAA wyróżnia/, "AAA wyróżnia obszar") +
      (highlighted.length > 1 ? ` (+${highlighted.length - 1} kolejne)` : "");
  } else if (!charts.length && !hasContent) {
    A.priority = "ZDALNIE";
    A.priorityReason = "Brak pewnego odczytu wykresów i treści AAA — brak danych do oceny technicznej.";
  } else if (!charts.length) {
    A.priority = "NISKI";
    A.priorityReason = "AAA nie wyróżnia żadnego obszaru, a wykresów nie odczytano; w tych danych nie widać odchylenia.";
  } else if (exceed) {
    A.priority = "ŚREDNI";
    A.priorityReason = `${exceed} seria(e) przekracza(ją) linie progowe na wykresach; brak reguł dla tego alertu.`;
  } else {
    A.priority = "NISKI";
    A.priorityReason = "Serie na wykresach nie przekraczają linii progowych w tych danych.";
  }

  A.hypotheses.push("Brak dedykowanych reguł — nie stawiamy hipotezy przyczyny bez dokumentacji.");
  A.actions.push({ text: "Porównać wykresy AAA z progami opisanymi w AAA.", criterion: "", mode: "zdalnie" });
  A.actions.push({ text: "Sprawdzić powiązane condition codes.", criterion: "", mode: "zdalnie" });
  addDocsActions(result, A);
  return A;
}

function analyzeAlert(result, now = new Date()) {
  const family = alertFamily(result.alertName);

  if (result.error) {
    const A = newAnalysis(result, family);
    A.priority = "ZDALNIE";
    A.problem = `Aktywny alert: ${result.alertName || "nieznany"}.`;
    A.priorityReason = `Brak oceny technicznej — AAA nie zostało odczytane (${result.error}).`;
    A.missing.push("dane AAA");
    A.observations.push({ status: "Fakt", text: `Błąd analizy: ${result.error}` });
    A.actions.push({ text: "Otworzyć AAA ręcznie i powtórzyć analizę tego aparatu.", criterion: "AAA otwiera się z właściwym J-number", mode: "zdalnie" });
    return A;
  }

  const charts = validateCharts(result);
  const A =
    family === "dataLogger" ? analyzeDataLogger(result, now) :
    family === "supplyThermal" ? analyzeSupplyThermal(result, charts) :
    family === "cmRing" ? analyzeCmRing(result, charts) :
    analyzeGeneric(result, charts);

  for (const fact of aaaContentFacts(result)) A.observations.push({ status: "Fakt", text: fact });

  const hasGuidance = (result.aaaContent || []).some(s => s.blocks.some(b => b.type === "list"));
  if (hasGuidance) {
    A.actions.push({ text: "Przejść kroki rozwiązywania problemów z AAA (rozdział „Treść strony AAA” w tym raporcie), zaczynając od wyróżnionego obszaru.", criterion: "", mode: "na miejscu" });
  }
  return A;
}

// Porównanie aparatów tego samego modelu (CLAUDE.md 4.5, 6.1.7, 6.3.5).
function compareAnalyzers(results, refAmbient = 27) {
  const notes = [];

  const byModel = new Map();
  for (const r of results) {
    if (!r.analysis) continue;
    if (!byModel.has(r.model)) byModel.set(r.model, []);
    byModel.get(r.model).push(r);
  }

  for (const [model, group] of byModel) {
    // Duty przy tej samej temperaturze otoczenia.
    const fits = group
      .filter(r => r.analysis.dutyFit)
      .map(r => {
        const f = r.analysis.dutyFit;
        const inRange = refAmbient >= f.xMin - 1 && refAmbient <= f.xMax + 1;
        const extrapolated = refAmbient < f.xMin || refAmbient > f.xMax;
        return { r, at: inRange ? Math.max(0, Math.min(100, f.a + f.b * refAmbient)) : null, extrapolated };
      });

    const withValue = fits.filter(x => x.at !== null);
    if (withValue.length >= 2) {
      const best = withValue.reduce((a, b) => (b.at < a.at ? b : a));
      for (const x of withValue) {
        const delta = x.at - best.at;
        notes.push({ model, jno: x.r.jno, kind: "duty", value: x.at, delta, ref: best.r.jno, extrapolated: x.extrapolated });
        if (x !== best && delta >= 10) {
          x.r.analysis.observations.push({
            status: "Odczyt",
            text: `Przy ≈ ${refAmbient} °C duty ≈ ${fmtN(x.at)}% vs ${fmtN(best.at)}% na J${best.r.jno} (różnica ${fmtN(delta)} pp, dopasowanie liniowe${x.extrapolated || best.extrapolated ? ", częściowo ekstrapolacja poza zakres danych" : ""}). Wskazówka na obniżoną sprawność; samo porównanie nie dowodzi usterki.`
          });
        }
      }
    }

    // Wersja oprogramowania inna niż u pozostałych aparatów tego modelu (obserwacja, nie przyczyna).
    const versions = group.map(r => (r.asset && r.asset.softwareVersion) || "").filter(Boolean);
    if (new Set(versions).size > 1) {
      const counts = new Map();
      versions.forEach(v => counts.set(v, (counts.get(v) || 0) + 1));
      const common = [...counts.entries()].sort((a, b) => b[1] - a[1])[0][0];
      for (const r of group) {
        const v = r.asset && r.asset.softwareVersion;
        if (v && v !== common) {
          r.analysis.observations.push({ status: "Fakt", text: `Wersja oprogramowania ${v} różni się od najczęstszej u aparatów ${model} (${common}). Obserwacja, nie przyczyna.` });
        }
      }
    }
  }

  return notes;
}

function sortByPriority(results) {
  return [...results].sort((a, b) =>
    PRIORITY_ORDER.indexOf(a.analysis.priority) - PRIORITY_ORDER.indexOf(b.analysis.priority)
  );
}

// ============================================================
// PDF (CLAUDE.md, rozdz. 8)
// ============================================================

const SERIES_PDF_COLORS = { niebieska: "#0000ff", "brązowa": "#8b4513", czerwona: "#d00000" };
const COMPARE_COLORS = ["#1f5fbf", "#c0392b", "#2e8b57", "#8e44ad", "#d68910"];

function generatePDF(results, analyzerCount, comparisonNotes = []) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({
      size: "A4",
      margin: 40,
      info: { Title: "Raport serwisowy e-Connectivity v3", Author: "Service Triage Agent v3" }
    });

    const stream = fs.createWriteStream(PDF_FILE);
    stream.on("finish", resolve);
    stream.on("error", reject);
    doc.pipe(stream);

    let REGULAR = "Helvetica";
    let BOLD = "Helvetica-Bold";

    const fontCandidates = [
      [ARIAL, ARIAL_BOLD],
      ["C:\\Windows\\Fonts\\segoeui.ttf", "C:\\Windows\\Fonts\\segoeuib.ttf"],
      ["C:\\Windows\\Fonts\\calibri.ttf", "C:\\Windows\\Fonts\\calibrib.ttf"],
      ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"],
      ["/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"],
      ["/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]
    ];

    const fontPair = fontCandidates.find(([regular, bold]) => fs.existsSync(regular) && fs.existsSync(bold));

    if (fontPair) {
      doc.registerFont("ArialPL", fontPair[0]);
      doc.registerFont("ArialPLBold", fontPair[1]);
      REGULAR = "ArialPL";
      BOLD = "ArialPLBold";
    } else {
      console.log("UWAGA: nie znaleziono czcionki z polskimi znakami (Arial/Segoe/DejaVu) — PDF będzie bez polskich liter.");
    }

    const left = doc.page.margins.left;
    const contentWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
    const bottomLimit = () => doc.page.height - doc.page.margins.bottom;

    function pageCheck(needed = 130) {
      if (doc.y > doc.page.height - needed) doc.addPage();
    }

    function heading(value, size = 15) {
      pageCheck(90);
      doc.x = left;
      doc.font(BOLD).fontSize(size).fillColor("#111111").text(value);
      doc.moveDown(0.4);
    }

    function subheading(value) {
      pageCheck(80);
      doc.x = left;
      doc.moveDown(0.5);
      doc.font(BOLD).fontSize(10.5).fillColor("#111111").text(value);
      doc.moveDown(0.2);
    }

    function field(label, value) {
      doc.x = left;
      doc.font(BOLD).fontSize(9.5).fillColor("#222222").text(`${label}: `, { continued: true });
      doc.font(REGULAR).text(value || "brak danych");
    }

    function paragraph(text, size = 9.3, color = "#333333") {
      pageCheck(50);
      doc.x = left;
      doc.font(REGULAR).fontSize(size).fillColor(color).text(text, { lineGap: 1.5 });
    }

    // Zdanie ze statusem (CLAUDE.md 3): status słowem, nie kolorem.
    function statusLine(status, text) {
      pageCheck(45);
      doc.x = left;
      doc.font(BOLD).fontSize(9).fillColor("#222222").text(`${status}: `, { continued: true });
      doc.font(REGULAR).fillColor("#333333").text(text, { lineGap: 1.5 });
      doc.moveDown(0.1);
    }

    function table(columns, rows, size = 8, rowFills = []) {
      const pad = 3;
      const totalW = columns.reduce((s, c) => s + c.width, 0);
      const scale = contentWidth / totalW;
      const widths = columns.map(c => c.width * scale);

      const rowHeight = (cells, font) => {
        doc.font(font).fontSize(size);
        return Math.max(...cells.map((t, i) => doc.heightOfString(String(t ?? ""), { width: widths[i] - 2 * pad }))) + 2 * pad;
      };

      const drawRow = (cells, font, fill) => {
        const h = rowHeight(cells, font);
        if (doc.y + h > bottomLimit()) {
          doc.addPage();
          if (font !== BOLD) drawRow(columns.map(c => c.header), BOLD, "#e8e8e8");
        }
        const y = doc.y;
        if (fill) doc.save().rect(left, y, contentWidth, h).fill(fill).restore();
        let x = left;
        cells.forEach((t, i) => {
          doc.font(font).fontSize(size).fillColor("#222222")
            .text(String(t ?? ""), x + pad, y + pad, { width: widths[i] - 2 * pad });
          x += widths[i];
        });
        doc.save().moveTo(left, y + h).lineTo(left + contentWidth, y + h).lineWidth(0.4).strokeColor("#bbbbbb").stroke().restore();
        doc.x = left;
        doc.y = y + h;
      };

      pageCheck(60);
      drawRow(columns.map(c => c.header), BOLD, "#e8e8e8");
      rows.forEach((r, i) => drawRow(r, REGULAR, rowFills[i] || null));
      doc.x = left;
      doc.moveDown(0.5);
    }

    // Wykres odtworzony z próbek: luki szarym pasem, nie linią (CLAUDE.md 8).
    function drawReconstructed({ title, series, yMin, yMax, step = null, thresholds = [], dateRange, height = 125 }) {
      if (!series.length || !(yMax > yMin)) return;
      pageCheck(height + 50);

      const padL = 42, padB = 16, padT = 14;
      const x0 = left + padL, y0 = doc.y + padT;
      const w = contentWidth - padL - 10, h = height - padT - padB;
      const X = f => x0 + f * w;
      const Y = v => y0 + (yMax - Math.min(yMax, Math.max(yMin, v))) / (yMax - yMin) * h;

      doc.font(BOLD).fontSize(8.5).fillColor("#222222").text(title, left, doc.y, { width: contentWidth });

      for (const s of series) {
        for (const g of s.gaps || []) doc.save().rect(X(g.from), y0, X(g.to) - X(g.from), h).fill("#dddddd").restore();
      }

      doc.save().rect(x0, y0, w, h).lineWidth(0.5).strokeColor("#888888").stroke().restore();

      // Podziałki jak na wykresie źródłowym (krok osi z odczytu), inaczej 4 równe części.
      const ticks = step && (yMax - yMin) / step <= 10 ? Math.round((yMax - yMin) / step) : 4;
      for (let i = 0; i <= ticks; i++) {
        const v = yMin + (yMax - yMin) * i / ticks;
        const y = Y(v);
        doc.save().moveTo(x0, y).lineTo(x0 + w, y).lineWidth(0.3).strokeColor("#e2e2e2").stroke().restore();
        doc.font(REGULAR).fontSize(6.5).fillColor("#555555").text(fmtN(v, yMax - yMin < 5 ? 2 : 1), left, y - 3, { width: padL - 4, align: "right" });
      }

      for (const t of thresholds) {
        doc.save().moveTo(x0, Y(t.value)).lineTo(x0 + w, Y(t.value)).dash(3, { space: 2 }).lineWidth(0.7).strokeColor(t.color || "#9400d3").stroke().undash().restore();
      }

      for (const s of series) {
        doc.save().lineWidth(0.8).strokeColor(s.color);
        let prev = null;
        for (const [f, v] of s.samples) {
          if (prev && f - prev <= 0.01) doc.lineTo(X(f), Y(v));
          else doc.moveTo(X(f), Y(v));
          prev = f;
        }
        doc.stroke().restore();
      }

      const a = parseUsDateTime(dateRange && dateRange.start), b = parseUsDateTime(dateRange && dateRange.end);
      doc.font(REGULAR).fontSize(6.5).fillColor("#555555");
      if (a) doc.text(fmtPlDateTime(a), x0, y0 + h + 3, { width: 100 });
      if (b) doc.text(fmtPlDateTime(b), x0 + w - 100, y0 + h + 3, { width: 100, align: "right" });

      if (series.length > 1 || series.some(s => s.label)) {
        let lx = x0 + 110;
        for (const s of series) {
          doc.save().rect(lx, y0 + h + 5, 8, 3).fill(s.color).restore();
          doc.font(REGULAR).fontSize(6.5).fillColor("#333333").text(s.label || "", lx + 11, y0 + h + 3, { width: 90 });
          lx += 100;
        }
      }

      doc.x = left;
      doc.y = y0 + h + padB + 4;
    }

    function chartReconstruction(chart, dateRange, label) {
      const local = chart.local;
      drawReconstructed({
        title: `${label} — odtworzony z odczytu${chart.uncertain ? " (odczyt niepewny)" : ""}`,
        series: local.series.map(s => ({ color: SERIES_PDF_COLORS[s.color] || "#333333", samples: s.samples || [], gaps: s.gaps || [], label: `seria ${s.color}` })),
        yMin: local.axis.bottom,
        yMax: local.axis.top,
        step: local.axis.step,
        thresholds: local.thresholds.map(t => ({ value: t.value, color: t.color === "żółta" ? "#c9a800" : "#9400d3" })),
        dateRange
      });
    }

    // Treść strony AAA jako Fakt (tekst źródłowy, bez interpretacji).
    function renderContent(sections) {
      for (const section of sections) {
        if (sections.length > 1) paragraph(`Źródło: ${section.source}`, 7.8, "#777777");
        for (const b of section.blocks) {
          if (b.type === "heading") {
            pageCheck(70);
            doc.x = left;
            doc.moveDown(0.35);
            doc.font(BOLD).fontSize(b.level <= 2 ? 11 : 9.8).fillColor("#111111").text(b.text);
            doc.moveDown(0.15);
          } else if (b.type === "para") {
            if (b.boxed) {
              pageCheck(60);
              const h = doc.font(REGULAR).fontSize(8.8).heightOfString(b.text, { width: contentWidth - 12 }) + 10;
              const y = doc.y;
              doc.save().rect(left, y, contentWidth, h).lineWidth(0.5).strokeColor("#999999").stroke().restore();
              doc.font(REGULAR).fontSize(8.8).fillColor("#222222").text(b.text, left + 6, y + 5, { width: contentWidth - 12 });
              doc.x = left;
              doc.y = y + h + 4;
            } else {
              paragraph(b.text, 8.8, "#222222");
            }
          } else if (b.type === "list") {
            doc.moveDown(0.1);
            for (const it of b.items) {
              pageCheck(36);
              const indent = 6 + it.level * 16;
              doc.font(REGULAR).fontSize(8.8).fillColor("#222222")
                .text(`${it.marker ? it.marker + " " : ""}${it.text}`, left + indent, doc.y, { width: contentWidth - indent, lineGap: 1 });
            }
            doc.x = left;
            doc.moveDown(0.3);
          } else if (b.type === "table") {
            const n = Math.min(10, Math.max(...b.rows.map(r => r.cells.length)));
            const head = b.rows[0].header ? b.rows[0] : { cells: Array.from({ length: n }, () => "") };
            const body = b.rows[0].header ? b.rows.slice(1) : b.rows;
            if (b.caption) paragraph(b.caption, 8.8, "#111111");
            table(
              Array.from({ length: n }, (_, i) => ({ header: head.cells[i] || "", width: i === 0 ? 2.4 : 1 })),
              body.map(r => {
                const cells = r.cells.slice(0, n);
                while (cells.length < n) cells.push("");
                if (r.highlight) cells[0] = `[wyróżnione na stronie: ${r.highlight} tło] ${cells[0]}`;
                return cells;
              }),
              7.6,
              body.map(r => (r.highlight ? "#fff4c2" : null))
            );
          } else if (b.type === "image") {
            paragraph(`[obraz na stronie${b.text ? `: ${b.text}` : ""} — widoczny w kopii strony na końcu sekcji]`, 7.8, "#777777");
          }
        }
      }
    }

    // Kopia strony AAA jako obraz, pocięta na strony PDF.
    function renderScreenshot(file) {
      let img;
      try { img = doc.openImage(file); } catch { return; }
      const scale = contentWidth / img.width;
      const totalH = img.height * scale;
      let offset = 0;
      let first = true;
      while (offset < totalH - 1) {
        // Pierwszy fragment pod nagłówkiem, jeśli zmieści się sensowna część; kolejne na nowych stronach.
        if (!first || bottomLimit() - doc.y < 200) doc.addPage();
        const top = first && bottomLimit() - doc.y >= 200 ? doc.y : doc.page.margins.top;
        first = false;
        const avail = bottomLimit() - top;
        const sliceH = Math.min(avail, totalH - offset);
        doc.save();
        doc.rect(left, top, contentWidth, sliceH).clip();
        doc.image(img, left, top - offset, { width: contentWidth });
        doc.restore();
        offset += sliceH;
      }
      doc.x = left;
      doc.y = bottomLimit();
    }

    const ordered = sortByPriority(results);
    const placeOf = r => `J${r.jno} — ${r.model}, ${r.location.customer}`;
    const firstAction = r => (r.analysis.actions[0] && r.analysis.actions[0].text) || "";

    // ======================================================
    // STRONA 1: PODSUMOWANIE
    // ======================================================

    doc.font(BOLD).fontSize(21).fillColor("#111111").text("RAPORT SERWISOWY", { align: "center" });
    doc.fontSize(15).text("QuidelOrtho e-Connectivity", { align: "center" });
    doc.moveDown(0.3);
    doc.font(REGULAR).fontSize(9).fillColor("#555555").text(`Wygenerowano: ${formatDate(new Date())}`, { align: "center" });
    doc.moveDown(1);

    heading("PODSUMOWANIE", 14);
    field("Sprawdzone analizatory", String(analyzerCount));
    field("Aktywne alerty", String(results.length));
    field("Alerty z odczytanym AAA", String(results.filter(r => r.aaaVerified).length));
    doc.moveDown(0.6);

    table(
      [
        { header: "Poziom", width: 55 },
        { header: "Aparat i placówka", width: 120 },
        { header: "Alert", width: 80 },
        { header: "Co pokazują dane", width: 160 },
        { header: "Zalecenie", width: 110 }
      ],
      ordered.map(r => [r.analysis.priority, placeOf(r), r.alertName, r.analysis.priorityReason, firstAction(r)])
    );

    subheading("Najważniejsze wnioski");
    const conclusions = ordered
      .filter(r => r.analysis.priority === "WYSOKI" || r.analysis.priority === "ŚREDNI")
      .slice(0, 4)
      .map(r => `J${r.jno} (${r.location.customer}): ${r.analysis.priorityReason}`);
    const remote = results.filter(r => r.analysis.priority === "ZDALNIE");
    if (remote.length) {
      conclusions.push(`${remote.length} alert(y) bez oceny technicznej (brak łączności lub danych): ${remote.map(r => `J${r.jno}`).join(", ")}. Najpierw odzyskać widoczność.`);
    }
    if (!conclusions.length) conclusions.push("W tych danych nie widać parametrów pracujących na granicy wydajności.");
    conclusions.slice(0, 5).forEach((c, i) => paragraph(`${i + 1}. ${c}`));

    // ======================================================
    // SEKCJE APARATÓW
    // ======================================================

    for (const item of ordered) {
      const a = item.analysis;
      doc.addPage();

      doc.font(BOLD).fontSize(17).fillColor("#111111").text(`${item.model} — J${item.jno}`);
      doc.font(BOLD).fontSize(12).text(`${item.alertName} · priorytet ${a.priority}`);
      doc.moveDown(0.2);
      paragraph(a.priorityReason, 9.5, "#222222");
      doc.moveDown(0.4);

      field("Placówka", item.location.customer);
      field("Adres", item.location.address);
      field("Miasto", item.location.city);
      if (item.asset) {
        field("Status w bazie", item.asset.excelStatus);
        field("Software", item.asset.softwareVersion);
      }
      field("Status e-Connectivity", item.status);
      field("Zakres danych AAA", item.dateRange.start ? `${item.dateRange.start} – ${item.dateRange.end}` : "brak pewnego odczytu");
      if (item.error) field("Błąd analizy", item.error);

      if (a.missing.length) {
        subheading("Kompletność danych");
        paragraph(`Brakuje: ${a.missing.join("; ")}.`);
      }

      if (a.readings.length) {
        subheading("Tabela odczytów");
        table(
          [
            { header: "Parametr", width: 150 },
            { header: "Odczyt", width: 110 },
            { header: "Odniesienie", width: 130 },
            { header: "Ocena", width: 110 }
          ],
          a.readings.map(r => [r.param, r.value, r.ref, r.assessment])
        );
      }

      const charts = (item.chartData || []).filter(c => c.local && c.local.ok && c.local.series.length && !(c.titleCheck && c.titleCheck.ok === false));
      if (charts.length) {
        subheading("Wykresy odtworzone z danych");
        for (const c of charts) chartReconstruction(c, item.dateRange, CHART_KIND_LABEL[c.kind] || c.title || "Wykres");
      }

      subheading("Co widać");
      for (const o of a.observations) statusLine(o.status, o.text);

      subheading("Co to może oznaczać");
      for (const h of a.hypotheses) paragraph(h);

      subheading("Zalecane działania");
      a.actions.forEach((act, i) => {
        pageCheck(45);
        doc.x = left;
        doc.font(REGULAR).fontSize(9.3).fillColor("#333333")
          .text(`${i + 1}. ${act.text} (${act.mode})${act.criterion ? ` Kryterium sukcesu: ${act.criterion}.` : ""}`, { indent: 6, lineGap: 1.5 });
        doc.moveDown(0.1);
      });

      // Pełna treść strony AAA i podstron.
      if ((item.aaaContent || []).length || (item.aaaSubpages || []).length) {
        doc.addPage();
        heading(`TREŚĆ STRONY AAA — J${item.jno}`, 13);
        paragraph("Fakt: tekst źródłowy ze strony AAA w oryginalnej kolejności i języku. Usunięto tylko adresy IP, URL-e i tokeny. Wiersze wyróżnione kolorem na stronie oznaczono słownie.", 8.3, "#555555");
        doc.moveDown(0.3);
        renderContent(item.aaaContent || []);

        for (const sub of item.aaaSubpages || []) {
          doc.addPage();
          heading(`PODSTRONA AAA: ${sub.title}`, 12);
          if (sub.note) paragraph(sub.note, 8.5, "#777777");
          renderContent(sub.sections || []);
        }
      }

      // Wykresy źródłowe AAA z odczytem lokalnym.
      if (item.chartFiles.length) {
        doc.addPage();
        heading("WYKRESY ŹRÓDŁOWE AAA", 13);
        paragraph("Obrazy pobrane z AAA. Pod każdym: walidacja odczytu i wartości odczytane lokalnie z pikseli.", 8.5, "#555555");
        doc.moveDown(0.4);

        item.chartFiles.forEach((file, index) => {
          if (!fs.existsSync(file)) return;
          pageCheck(330);
          const info = (item.chartData || [])[index] || null;
          try {
            doc.x = left;
            doc.font(BOLD).fontSize(9).fillColor("#333333").text(`Wykres ${index + 1}${info && info.title ? `: ${info.title}` : ""}`);
            if (info && info.validity) {
              doc.font(REGULAR).fontSize(7.8).fillColor(info.validity === "wiarygodny" ? "#1a6b3a" : "#a04000").text(`Walidacja: ${info.validity}.`);
            }
            // Pozycja liczona jawnie, żeby tekst pod obrazem go nie nachodził.
            const img = doc.openImage(file);
            const k = Math.min(500 / img.width, 260 / img.height, 1);
            const iw = img.width * k, ih = img.height * k;
            if (doc.y + ih > bottomLimit()) doc.addPage();
            const iy = doc.y;
            doc.image(img, left + (contentWidth - iw) / 2, iy, { width: iw, height: ih });
            doc.x = left;
            doc.y = iy + ih + 4;

            if (info && info.local && info.local.ok) {
              doc.font(REGULAR).fontSize(7.6).fillColor("#222222");
              for (const line of describeLocalChart(info.local, item.dateRange).lines) {
                pageCheck(40);
                doc.x = left;
                doc.text(line.startsWith("   ") ? "      " + line.trim() : "• Odczyt: " + line);
              }
            } else if (info && info.local && info.local.reason) {
              doc.font(REGULAR).fontSize(7.5).fillColor("#999999").text(`Brak pewnego odczytu: ${info.local.reason}.`);
            } else if (info && !info.local) {
              doc.font(REGULAR).fontSize(7.5).fillColor("#999999").text("Brak pewnego odczytu — odczyt lokalny wyłączony (npm install pngjs tesseract.js @tesseract.js-data/eng).");
            }

            if (info && info.points && info.points.length) {
              const s = info.summary || {};
              doc.font(BOLD).fontSize(8).fillColor("#1a4d8f").text(
                `Fakt: wartości ze strony AAA (${info.points.length} pkt)` +
                (s.min !== undefined ? ` — min ${s.min}, maks ${s.max}, średnia ${s.avg}, ostatnia ${s.last}` : "")
              );
            }

            if (info && info.vision) {
              const v = info.vision;
              doc.font(BOLD).fontSize(8).fillColor("#8a5a00").text("Odczyt AI z obrazu (AAA_VISION) — szacunkowy, zweryfikować z wykresem:");
              doc.font(REGULAR).fontSize(7.5).fillColor("#333333");
              for (const serie of (v.series || []).slice(0, 8)) {
                doc.text(`• ${serie.name || "seria"}: min ${serie.min ?? "?"}, maks ${serie.max ?? "?"}, ostatnia ${serie.last ?? "?"}`);
              }
            }
            doc.moveDown(0.6);
          } catch {}
        });
      }
    }

    // ======================================================
    // KOPIE STRON AAA (obraz)
    // ======================================================

    for (const item of ordered) {
      if (!item.aaaScreenshot || !fs.existsSync(item.aaaScreenshot)) continue;
      doc.addPage();
      heading(`KOPIA STRONY AAA — J${item.jno} (${item.alertName})`, 13);
      paragraph("Zrzut całej strony AAA w chwili analizy, z obrazkami i wykresami. Pocięty na kolejne strony.", 8.3, "#555555");
      renderScreenshot(item.aaaScreenshot);
    }

    // ======================================================
    // PORÓWNANIE APARATÓW TEGO SAMEGO MODELU
    // ======================================================

    const groups = new Map();
    for (const r of results) {
      for (const c of r.chartData || []) {
        if (!c.kind || !c.local || !c.local.ok || !c.local.series.length) continue;
        if (c.titleCheck && c.titleCheck.ok === false) continue;
        const key = `${r.model}|${c.kind}`;
        if (!groups.has(key)) groups.set(key, { model: r.model, kind: c.kind, rows: [] });
        const g = groups.get(key);
        if (!g.rows.some(x => x.r.jno === r.jno)) g.rows.push({ r, c });
      }
    }
    const comparable = [...groups.values()].filter(g => g.rows.length >= 2);

    if (comparable.length || comparisonNotes.length) {
      doc.addPage();
      heading("PORÓWNANIE APARATÓW TEGO SAMEGO MODELU", 14);
      paragraph("Wspólna skala osi dla porównywanych aparatów. Pierwsza seria każdego wykresu. Luki w danych szarym pasem.", 8.5, "#555555");
      doc.moveDown(0.4);

      for (const g of comparable) {
        const yMin = Math.min(...g.rows.map(x => x.c.local.axis.bottom));
        const yMax = Math.max(...g.rows.map(x => x.c.local.axis.top));
        drawReconstructed({
          title: `${g.model} — ${CHART_KIND_LABEL[g.kind]}`,
          series: g.rows.map((x, i) => ({
            color: COMPARE_COLORS[i % COMPARE_COLORS.length],
            samples: x.c.local.series[0].samples || [],
            gaps: x.c.local.series[0].gaps || [],
            label: `J${x.r.jno}`
          })),
          yMin, yMax,
          step: g.rows[0].c.local.axis.step,
          thresholds: g.rows[0].c.local.thresholds.map(t => ({ value: t.value, color: t.color === "żółta" ? "#c9a800" : "#9400d3" })),
          dateRange: g.rows[0].r.dateRange,
          height: 150
        });

        const ref = g.rows[0].c.local.series[0];
        table(
          [
            { header: "Aparat", width: 90 },
            { header: "Średnia", width: 70 },
            { header: "Maks.", width: 70 },
            { header: "Ostatnia", width: 70 },
            { header: "Różnica średnich vs pierwszy", width: 120 }
          ],
          g.rows.map(x => {
            const s = x.c.local.series[0], res = x.c.local.axis.resolution;
            return [`J${x.r.jno}`, fmtChartNum(s.avg, res), fmtChartNum(s.max, res), fmtChartNum(s.last, res), signedChartNum(s.avg - ref.avg, res)];
          })
        );
      }

      const dutyNotes = comparisonNotes.filter(n => n.kind === "duty");
      if (dutyNotes.length) {
        subheading("Duty przy ≈ 27 °C otoczenia (dopasowanie liniowe)");
        table(
          [
            { header: "Model", width: 90 },
            { header: "Aparat", width: 80 },
            { header: "Duty przy ≈ 27 °C", width: 100 },
            { header: "Różnica vs najlepszy", width: 120 }
          ],
          dutyNotes.map(n => [n.model, `J${n.jno}`, `${fmtN(n.value)}%${n.extrapolated ? " (ekstrapolacja)" : ""}`, `${n.delta >= 0 ? "+" : ""}${fmtN(n.delta)} pp vs J${n.ref}`])
        );
        paragraph("Różnica ≥ 10 pp przy podobnym otoczeniu to wskazówka na obniżoną sprawność. Samo porównanie nie dowodzi usterki.", 8.5);
      }
    }

    // ======================================================
    // ALERTY ŁĄCZNOŚCI
    // ======================================================

    const connectivity = ordered.filter(r => r.analysis.family === "dataLogger");
    if (connectivity.length) {
      doc.addPage();
      heading("ALERTY ŁĄCZNOŚCI", 14);
      table(
        [
          { header: "Aparat i placówka", width: 140 },
          { header: "Ostatnie połączenie", width: 90 },
          { header: "Bez łączności", width: 70 },
          { header: "Kroki", width: 200 }
        ],
        connectivity.map(r => {
          const last = parseUsDateTime(r.dataLogger && r.dataLogger.lastConnected);
          return [
            placeOf(r),
            last ? fmtPlDateTime(last) : "brak pewnego odczytu",
            r.analysis.elapsedText || "?",
            r.analysis.actions.map((x, i) => `${i + 1}. ${x.text}`).join(" ")
          ];
        })
      );
      paragraph("Aparat bez łączności jest niewidoczny — jego stanu technicznego nie oceniamy.", 8.5);
    }

    // ======================================================
    // PLAN DZIAŁAŃ
    // ======================================================

    doc.addPage();
    heading("PLAN DZIAŁAŃ", 14);
    const planRows = [];
    for (const r of ordered) {
      for (const act of r.analysis.actions) {
        planRows.push([act.text, `J${r.jno}`, act.mode, PRIORITY_TERM[r.analysis.priority] || ""]);
      }
    }
    table(
      [
        { header: "Działanie", width: 260 },
        { header: "Aparat", width: 70 },
        { header: "Tryb", width: 70 },
        { header: "Termin", width: 110 }
      ],
      planRows
    );

    // ======================================================
    // METODYKA I OGRANICZENIA
    // ======================================================

    heading("METODYKA I OGRANICZENIA", 14);
    [
      "Źródła: dashboard e-Connectivity (alerty, J-number), baza Excel (model, placówka, software), strony AAA (zakres dat, reguły, wykresy).",
      "Wykresy AAA to obrazy PNG bez danych liczbowych w kodzie strony. Wartości odczytano lokalnie z pikseli: skala osi Y z OCR etykiet, serie i progi po kolorach MS Chart. Dokładność odczytu podano przy każdym wykresie (≈ ±1 piksel w jednostkach osi).",
      "Oś czasu przyjęto jako liniową na zakres dat AAA. Godziny z wykresów są przybliżone (±1–2 h) i oznaczone „ok.”. Dokładne czasy pochodzą tylko z danych tekstowych.",
      "Luki w danych wykryto jako odcinki idealnie liniowe dłuższe niż ≈ 3% szerokości wykresu (AAA łączy luki prostą linią) oraz jako odcinki bez pikseli serii. W seriach z szumem usunięto je ze statystyk. W seriach gładkich (np. powolna zmiana temperatury) prosty odcinek może być prawdziwymi danymi — oznaczono go jako możliwą lukę i zostawiono. Poziome plateau (np. nasycenie 100%) zostaje w danych; pozioma luka nie jest wykrywana.",
      "Odczyt uznano za wiarygodny, gdy tytuł zawiera J-number aparatu, skala ma równe kroki potwierdzone OCR, progi leżą w zakresie osi, a seria ma wystarczająco pikseli. Wykres z innym J-number wyłączono z analizy.",
      "Populacje na wykresach adjustment liczone z pikseli niedoszacowują gęste obszary — traktować orientacyjnie.",
      "Czas od ostatniego połączenia liczony względem zegara komputera generującego raport; strefa czasowa dashboardu nie jest weryfikowana.",
      `Odczyt AI z obrazów (AAA_VISION): ${process.env.AAA_VISION === "1" ? "WŁĄCZONY — obrazy wysłano do usługi zewnętrznej za zgodą użytkownika" : "wyłączony — żadne obrazy nie opuściły komputera"}.`,
      "Nie zweryfikowano: stanu aparatów na miejscu, historii serwisowej, przyczyn usterek. Każda przyczyna jest hipotezą do potwierdzenia pomiarem przez FSE.",
      "Treść stron AAA odczytano z drzewa strony (DOM) bez klikania: sekcje zwinięte lub ukryte na stronie nie są uwzględnione. Podstrony otwierane tylko przy AAA_FOLLOW_LINKS=1 (linki z tej samej domeny, bez linków-akcji).",
      "Tryb tylko do odczytu: agent nie wykonał Save, Apply, zmian konfiguracji, adjustmentów, resetów ani komend serwisowych."
    ].forEach(t => { paragraph(`• ${t}`, 8.8); doc.moveDown(0.15); });

    doc.end();
  });
}

// ============================================================
// MAIN
// ============================================================

async function runAgentInner() {

  ensureDir(
    OUTPUT_DIR
  );

  console.log("");
  console.log(
    "================================================"
  );

  console.log(
    " QUIDELORTHO LIVE DIAGNOSTIC AGENT"
  );

  console.log(
    " EXCEL + AAA + POLSKI RAPORT FE"
  );

  console.log(
    " READ-ONLY"
  );

  console.log(
    "================================================"
  );

  // ==========================================================
  // LOAD EXCEL FIRST
  // ==========================================================

  const analyzerDatabase =
    loadAnalyzerDatabase();

  const serviceKnowledge =
    loadServiceKnowledge();

  // ==========================================================
  // EDGE
  // ==========================================================

  const browser =
    await chromium.launch({
      headless: false,

      executablePath:
        EDGE_PATH
    });

  const options = {
    viewport: {
      width: 1500,
      height: 950
    }
  };

  if (
    fs.existsSync(
      SESSION_FILE
    )
  ) {
    options.storageState =
      SESSION_FILE;
  }

  const context =
    await browser.newContext(
      options
    );

  const page =
    await context.newPage();

  console.log("");
  console.log(
    "Otwieram e-Connectivity..."
  );

  await page.goto(
    START_URL,
    {
      waitUntil:
        "domcontentloaded",

      timeout:
        60000
    }
  ).catch(
    () => {}
  );

  // ==========================================================
  // USER NAVIGATION
  // ==========================================================

  await waitForEnter();

  // ==========================================================
  // DASHBOARD STABILITY
  // ==========================================================

  const state =
    await waitForDashboardStable(
      page
    );

  if (
    !state ||
    state.analyzerCount <= 1
  ) {
    console.log("");
    console.log(
      "BŁĄD: dashboard nie został poprawnie wykryty."
    );

    console.log(
      "Raport NIE został wygenerowany."
    );

    return;
  }

  console.log("");
  console.log(
    `Dashboard gotowy. Analizatory: ${state.analyzerCount}`
  );

  // ==========================================================
  // ALERT INVENTORY
  // ==========================================================

  const alerts =
    await inventoryDashboard(
      page
    );

  console.log(
    `Aktywne alerty O/Y: ${alerts.length}`
  );

  const results = [];

  // ==========================================================
  // LOOP
  // ==========================================================

  for (
    let index = 0;
    index < alerts.length;
    index++
  ) {
    const alert =
      alerts[index];

    const asset =
      findAnalyzer(
        analyzerDatabase,
        alert.jno
      );

    const model =
      getAnalyzerModel(
        asset
      );

    const location =
      getLocation(
        asset
      );

    console.log("");
    console.log(
      "------------------------------------------------"
    );

    console.log(
      `[${index + 1}/${alerts.length}] ${model} — J${alert.jno}`
    );

    console.log(
      `Placówka: ${location.customer}`
    );

    console.log(
      `Miasto:    ${location.city}`
    );

    if (!asset) {
      console.log(
        "UWAGA: J-number nie został znaleziony w Excelu."
      );
    }

    const result = {
      jno:
        alert.jno,

      model,

      asset,

      location,

      status:
        alert.status,

      alertName:
        alert.alertName,

      aaaVerified:
        false,

      dateRange: {
        start: "",
        end: ""
      },

      dataLogger: {
        detected:
          false,

        lastConnected:
          "",

        aFile: [],

        bFile: []
      },

      chartFiles: [],

      chartData: [],

      diagnosticRules: [],

      genericEvidence: { rules: [], conditionCodes: [], linkLabels: [], rawTextSanitized: "" },

      serviceDocuments: [],

      aaaContent: [],

      aaaSubpages: [],

      aaaScreenshot: "",

      error: ""
    };

    try {
      // ======================================================
      // CELL
      // ======================================================

      const cell =
        await findAlertCell(
          page,
          alert
        );

      if (!cell) {
        throw new Error(
          "Nie znaleziono komórki alertu."
        );
      }

      // ======================================================
      // AAA LINK
      // ======================================================

      const aaaLink =
        await findAAA(
          page,
          cell
        );

      if (!aaaLink) {
        throw new Error(
          "Nie znaleziono AAA."
        );
      }

      // ======================================================
      // AAA PAGE
      // ======================================================

      const aaaPage =
        await clickAndCatchPage(
          context,
          aaaLink
        );

      if (!aaaPage) {
        throw new Error(
          "AAA nie otworzyło strony."
        );
      }

      try {
        await sleep(
          1000
        );

        // ====================================================
        // VERIFY JNO
        // ====================================================

        result.aaaVerified =
          await verifyAAA(
            aaaPage,
            alert.jno
          );

        if (
          !result.aaaVerified
        ) {
          throw new Error(
            `AAA J-number mismatch dla J${alert.jno}.`
          );
        }

        console.log(
          "AAA: OK"
        );

        // ====================================================
        // ALERT NAME
        // ====================================================

        const aaaInfo =
          await extractAAAInformation(
            aaaPage,
            alert.alertName
          );

        if (
          aaaInfo.detectedName
        ) {
          result.alertName =
            aaaInfo.detectedName;
        }

        if (
          !result.alertName ||
          /^alert count$/i.test(
            result.alertName
          ) ||
          /^column\s+\d+/i.test(
            result.alertName
          )
        ) {
          result.alertName =
            "Alert AAA";
        }

        console.log(
          `Alert: ${result.alertName}`
        );

        // ====================================================
        // INITIAL DATA
        // ====================================================

        result.dateRange =
          await readDateRange(
            aaaPage
          );

        result.dataLogger =
          await extractDataLogger(
            aaaPage
          );

        result.diagnosticRules =
          await extractDiagnosticRules(
            aaaPage
          );

        result.genericEvidence =
          await extractGenericAAAEvidence(aaaPage);

        // ====================================================
        // FOLDER
        // ====================================================

        const alertDir =
          path.join(
            OUTPUT_DIR,

            `${safeName(model)}_J${alert.jno}_${safeName(result.alertName)}`
          );

        const chartDir =
          path.join(
            alertDir,
            "charts"
          );

        ensureDir(
          chartDir
        );

        // ====================================================
        // RUN
        // ====================================================

        const run =
          aaaPage.locator(
            "#ctl00_AAA_HeaderMain1_btnRun"
          );

        if (
          await run.count() &&
          await run
            .isVisible()
            .catch(
              () => false
            )
        ) {
          console.log(
            "Generuję wykresy AAA..."
          );

          const handler =
            createChartHandler(
              aaaPage,
              chartDir,
              result.chartFiles
            );

          await run.click();

          await aaaPage
            .waitForLoadState(
              "networkidle",
              {
                timeout:
                  15000
              }
            )
            .catch(
              () => {}
            );

          await sleep(
            2500
          );

          aaaPage.off(
            "response",
            handler
          );

          result.dateRange =
            await readDateRange(
              aaaPage
            );

          result.dataLogger =
            await extractDataLogger(
              aaaPage
            );

          result.diagnosticRules =
            await extractDiagnosticRules(
              aaaPage
            );

          result.chartData =
            await collectChartData(
              aaaPage,
              result.chartFiles
            );

          console.log(
            `Wykresy zapisane: ${result.chartFiles.length}, z odczytem wartości: ${
              result.chartData.filter(
                c => c.points.length || c.vision || (c.local && c.local.ok)
              ).length
            }`
          );
        }

        // ====================================================
        // PEŁNA TREŚĆ STRONY AAA + PODSTRONY + KOPIA OBRAZU
        // ====================================================

        result.aaaContent =
          await extractAAAContent(
            aaaPage
          );

        result.aaaSubpages =
          await extractSubpages(
            context,
            aaaPage
          );

        if (process.env.AAA_SCREENSHOT !== "0") {
          const shot = path.join(alertDir, "aaa_strona.png");
          await aaaPage
            .screenshot({ path: shot, fullPage: true })
            .then(() => { result.aaaScreenshot = shot; })
            .catch(() => {});
        }

        console.log(
          `Treść AAA: ${result.aaaContent.reduce((n, s) => n + s.blocks.length, 0)} bloków` +
          (result.aaaSubpages.length ? `, podstrony: ${result.aaaSubpages.length}` : "") +
          (result.aaaScreenshot ? ", kopia strony zapisana" : "")
        );

        // ====================================================
        // TECHNICAL JSON
        // ====================================================

        fs.writeFileSync(
          path.join(
            alertDir,
            "data.json"
          ),

          JSON.stringify(
            {
              jno:
                result.jno,

              model:
                result.model,

              asset:
                result.asset,

              location:
                result.location,

              status:
                result.status,

              alertName:
                result.alertName,

              dateRange:
                result.dateRange,

              dataLogger:
                result.dataLogger,

              diagnosticRules:
                result.diagnosticRules,

              genericEvidence:
                result.genericEvidence,

              chartFiles:
                result.chartFiles.map(file => path.basename(file)),

              chartData:
                result.chartData,

              serviceDocuments:
                result.serviceDocuments,

              aaaContent:
                result.aaaContent,

              aaaSubpages:
                result.aaaSubpages
            },
            null,
            2
          ),

          "utf8"
        );

      } finally {
        if (
          !aaaPage.isClosed()
        ) {
          await aaaPage
            .close()
            .catch(
              () => {}
            );
        }
      }

    } catch (error) {
      result.error =
        cleanText(
          error.message ||
          String(error)
        );

      console.log(
        `BŁĄD: ${result.error}`
      );
    }

    // ========================================================
    // ANALYSIS
    // ========================================================

    result.serviceDocuments =
      searchServiceKnowledge(
        serviceKnowledge,
        result
      );

    result.analysis =
      analyzeAlert(
        result
      );

    console.log(
      `PRIORYTET: ${result.analysis.priority} — ${result.analysis.priorityReason}`
    );

    results.push(
      result
    );

    await sleep(
      500
    );
  }

  // ==========================================================
  // PORÓWNANIE APARATÓW
  // ==========================================================

  const comparisonNotes =
    compareAnalyzers(
      results
    );

  // ==========================================================
  // MASTER JSON
  // ==========================================================

  fs.writeFileSync(
    JSON_FILE,

    JSON.stringify(
      {
        version:
          "v3",

        generatedAt:
          new Date()
            .toISOString(),

        excelDatabase:
          path.basename(
            EXCEL_FILE
          ),

        readOnly:
          true,

        analyzersChecked:
          state.analyzerCount,

        activeAlerts:
          results.length,

        comparison:
          comparisonNotes,

        results:
          results.map(
            r => ({
              jno:
                r.jno,

              model:
                r.model,

              asset:
                r.asset,

              location:
                r.location,

              status:
                r.status,

              alertName:
                r.alertName,

              aaaVerified:
                r.aaaVerified,

              dateRange:
                r.dateRange,

              chartCount:
                r.chartFiles.length,

              chartData:
                r.chartData,

              dataLogger:
                r.dataLogger,

              genericEvidence:
                r.genericEvidence,

              serviceDocuments:
                r.serviceDocuments,

              aaaContent:
                r.aaaContent,

              aaaSubpages:
                r.aaaSubpages,

              analysis:
                r.analysis,

              error:
                r.error
            })
          )
      },
      null,
      2
    ),

    "utf8"
  );

  // ==========================================================
  // PDF
  // ==========================================================

  console.log("");
  console.log(
    "Tworzę raport PDF..."
  );

  await generatePDF(
    results,
    state.analyzerCount,
    comparisonNotes
  );

  // ==========================================================
  // SAVE SESSION
  // ==========================================================

  try {
    await context.storageState({
      path:
        SESSION_FILE
    });
  } catch {}

  // ==========================================================
  // DONE
  // ==========================================================

  console.log("");
  console.log(
    "================================================"
  );

  console.log(
    " ANALIZA ZAKOŃCZONA"
  );

  console.log(
    "================================================"
  );

  console.log("");

  for (
    const result
    of results
  ) {
    console.log(
      `${result.model} — J${result.jno}`
    );

    console.log(
      `  ${result.location.customer}`
    );

    console.log(
      `  ${result.location.city}`
    );

    console.log(
      `  ${result.alertName}`
    );

    console.log(
      `  PRIORYTET: ${result.analysis.priority}`
    );

    console.log("");
  }

  console.log(
    `PDF:  ${PDF_FILE}`
  );

  console.log(
    `JSON: ${JSON_FILE}`
  );

  console.log("");
  console.log(
    "READ-ONLY:"
  );

  console.log(
    "- brak Save"
  );

  console.log(
    "- brak Apply"
  );

  console.log(
    "- brak zmian konfiguracji"
  );

  console.log(
    "- brak adjustmentów"
  );

  console.log(
    "- brak resetów"
  );

  console.log(
    "- brak service commands"
  );

}

async function runAgent() {
  try {
    await runAgentInner();
  } finally {
    await shutdownChartReader();
  }
}

if (require.main === module) {
  runAgent().catch(error => {

  console.error("");
  console.error(
    "FATAL ERROR:"
  );

  console.error(
    cleanText(
      error.stack ||
      error.message ||
      String(error)
    )
  );
});
}

module.exports = {
  repairPolishText,
  cleanText,
  parseNumbersFromText,
  summarizePoints,
  generatePDF,
  readChartsLocally,
  shutdownChartReader,
  describeLocalChart,
  analyzeAlert,
  compareAnalyzers,
  parseUsDateTime,
  titleJnoCheck,
  chartKind,
  collectPageBlocksInBrowser,
  sanitizeBlocks,
  aaaContentFacts,
  extractSubpages,
  PDF_FILE,
  OUTPUT_DIR
};
