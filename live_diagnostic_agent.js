// ============================================================
// QUIDELORTHO e-CONNECTIVITY SERVICE TRIAGE AGENT v12
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
function genericLiveAnalysis(r){
  const rules=r.genericEvidence?.rules||[],codes=r.genericEvidence?.conditionCodes||[],docs=r.serviceDocuments||[];
  if(/data\s*logger/i.test(r.alertName||""))return{priority:"ZDALNA WERYFIKACJA",decision:"REMOTE ONLY",problem:"e-Connectivity zgłasza brak lub opóźnienie danych z analizatora.",since:`Last Connected: ${r.dataLogger?.lastConnected||"brak pewnego odczytu"}.`,evidence:`AAA: ${rules.length} reguł/parametrów; ${codes.length} condition codes; ${r.chartFiles.length} wykresów.`,interpretation:"Sam Data Logger Status nie potwierdza awarii mechanicznej analizatora.",cause:"Najpierw zweryfikować dostępność analizatora, transmisję i e-Connectivity.",actions:["Sprawdzić Last Connected oraz A-file/B-file upload status.","Potwierdzić, czy analizator jest uruchomiony i dostępny.","Sprawdzić pozostałe aktywne alerty techniczne tego aparatu.","Wyjazd FSE rozważyć po wykluczeniu problemu komunikacyjnego."],conclusion:"Najpierw diagnostyka zdalna."};
  return{priority:r.status==="ORANGE"?"POMARAŃCZOWY":"ŻÓŁTY",decision:r.status==="ORANGE"?"WYMAGA OCENY SERWISOWEJ":"MONITOROWAĆ / OCENIĆ",problem:`Aktywny alert: ${r.alertName}.`,since:r.dateRange?.start?`Zakres danych AAA: ${r.dateRange.start} – ${r.dateRange.end}.`:"AAA nie podało pewnego początku problemu.",evidence:`AAA zweryfikowane dla J${r.jno}. Reguły/progi/parametry: ${rules.length}. Condition codes: ${codes.length?codes.join(", "):"brak pewnego odczytu"}. Wykresy: ${r.chartFiles.length}. Dokumenty ServicePublications: ${docs.length}.`,interpretation:r.chartFiles.length?(()=>{const cd=(r.chartData||[]);const withVals=cd.filter(c=>c.points&&c.points.length).length;const ai=cd.filter(c=>c.vision).length;const loc=cd.filter(c=>c.local&&c.local.ok).length;const fin=cd.flatMap(c=>describeLocalChart(c.local,r.dateRange).findings).slice(0,6);return `Wykresy źródłowe zapisano. Odczyt lokalny z obrazu: ${loc}/${r.chartFiles.length}; wartości ze strony: ${withVals}/${r.chartFiles.length}; odczyt AI: ${ai}/${r.chartFiles.length}.${fin.length?" Kluczowe odczyty: "+fin.join(" "):" Brak odczytu = wartości trzeba ocenić z samego wykresu."}`})():"Brak wykresów nie jest interpretowany jako brak problemu.",cause:docs.length?`Dopasowano dokumentację ServicePublications dla ${r.model}; należy skorelować ją z evidence AAA.`:"Brak wystarczająco pewnego dopasowania dokumentacji — bez zgadywania przyczyny.",actions:["Porównać evidence AAA z progami opisanymi przez AAA.","Sprawdzić powiązane condition codes.",...docs.slice(0,5).map(d=>`Dokumentacja: ${d.title}`),"Nie wykonywać adjustmentów ani zmian konfiguracji automatycznie."],conclusion:r.status==="ORANGE"?"Alert pomarańczowy wymaga oceny FSE; pilność musi wynikać z evidence, nie tylko z koloru.":"Alert żółty wymaga oceny trendu i evidence."}
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

      // Poziome linie = progi (wypełniają większość szerokości wykresu).
      const lineRows = [...rowCount.entries()].filter(([, n]) => n >= 0.6 * plotW).map(([y]) => y).sort((a, b) => a - b);
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
      let minY = Infinity, maxY = -Infinity;
      for (const [x, ys] of [...colsMap.entries()].sort((a, b) => a[0] - b[0])) {
        const clean = ys.filter(y => !skip.has(y));
        if (!clean.length) continue;
        medians.push({ x, y: median(clean) });
        minY = Math.min(minY, ...clean);
        maxY = Math.max(maxY, ...clean);
      }
      if (medians.length < 8) continue;

      const vals = medians.map(m => valueAt(m.y));
      const q = Math.max(1, Math.floor(vals.length / 4));
      const avg = a => a.reduce((s, v) => s + v, 0) / a.length;

      const maxVal = valueAt(minY);
      const minVal = valueAt(maxY);
      const maxCol = medians.reduce((b, m) => (m.y < b.y ? m : b), medians[0]);
      const minCol = medians.reduce((b, m) => (m.y > b.y ? m : b), medians[0]);

      // Luki w danych: odcinki >3% szerokości bez pikseli serii.
      let longestGap = 0;
      for (let i = 1; i < medians.length; i++) longestGap = Math.max(longestGap, medians[i].x - medians[i - 1].x);

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
        gapShare: Number((longestGap / plotW).toFixed(3)),
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

function parseUsDateTime(value) {
  const m = String(value || "").match(
    /(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?/
  );
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]), Number(m[4]), Number(m[5]), Number(m[6] || 0));
}

function approxTimeText(dateRange, fraction) {
  const a = parseUsDateTime(dateRange && dateRange.start);
  const b = parseUsDateTime(dateRange && dateRange.end);
  if (!a || !b || fraction === undefined || fraction === null) return "";
  const t = new Date(a.getTime() + Math.max(0, Math.min(1, fraction)) * (b.getTime() - a.getTime()));
  const p = n => String(n).padStart(2, "0");
  return ` (~${p(t.getDate())}.${p(t.getMonth() + 1)} ${p(t.getHours())}:${p(t.getMinutes())})`;
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
    if (s.gapShare > 0.05) line += ` — luka w danych ok. ${pct(s.gapShare)} zakresu`;
    lines.push(line);

    const byColor = new Map();
    for (const t of local.thresholds) {
      byColor.set(t.color, (byColor.get(t.color) || 0) + 1);
    }

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

function chartKey(title) {
  return String(title || "")
    .replace(/\s*--?\s*[A-Z]?\d{5,}\s*$/i, "")
    .replace(/[^A-Za-z0-9()\s.]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function buildChartComparison(results) {
  const groups = new Map();

  for (const r of results) {
    for (const c of r.chartData || []) {
      const local = c.local;
      if (!local || !local.ok || !local.series.length) continue;
      const key = chartKey(local.title);
      if (!key) continue;
      if (!groups.has(key)) groups.set(key, { title: local.title, rows: [] });
      const s = local.series[0];
      groups.get(key).rows.push({
        jno: r.jno,
        res: local.axis.resolution,
        last: s.last, avg: s.avg, max: s.max, min: s.min, trend: s.trend,
        thr: local.thresholds.map(t => t.value)
      });
    }
  }

  return [...groups.values()].filter(g => g.rows.length >= 2);
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
// ANALYSIS
// ============================================================

function analyzeAlert(result) {
  const name =
    String(
      result.alertName || ""
    ).toLowerCase();

  // ----------------------------------------------------------
  // DATA LOGGER
  // ----------------------------------------------------------

  if (
    name.includes(
      "data logger"
    )
  ) {
    let since =
      "Nie udało się ustalić ostatniego połączenia.";

    if (
      result.dataLogger
        ?.lastConnected
    ) {
      since =
        `Ostatnie połączenie zarejestrowane przez e-Connectivity: ${result.dataLogger.lastConnected}.`;
    }

    return {
      priority:
        "ŻÓŁTY",

      decision:
        "NAJPIERW WERYFIKACJA ZDALNA",

      problem:
        "e-Connectivity nie otrzymuje aktualnych danych z analizatora. Sam Data Logger Status nie potwierdza awarii mechanicznej analizatora.",

      since,

      evidence:
        "Należy ocenić Last Connected oraz status wysyłania A-file i B-file. Brak plików przez kolejne dni wskazuje przede wszystkim na problem z dostępnością danych lub komunikacją.",

      interpretation:
        "Pomarańczowy Data Logger nie powinien automatycznie powodować wyjazdu FE. Najpierw należy potwierdzić, czy analizator pracuje oraz czy działa jego połączenie e-Connectivity.",

      cause:
        "Najbardziej prawdopodobny kierunek: komunikacja e-Connectivity, brak transmisji danych, wyłączony analizator albo problem po stronie połączenia.",

      actions: [
        "Sprawdzić, czy analizator jest uruchomiony i pracuje.",
        "Zweryfikować zdalnie e-Connectivity.",
        "Sprawdzić Last Connected.",
        "Sprawdzić A-file i B-file upload status.",
        "Sprawdzić, czy analizator posiada inne aktywne alerty techniczne.",
        "Wyjazd FSE planować dopiero, jeżeli problemu nie można rozwiązać zdalnie lub występują inne objawy."
      ],

      conclusion:
        "Na podstawie samego Data Logger nie ma podstaw do pilnego wyjazdu FE."
    };
  }

  // ----------------------------------------------------------
  // SUPPLY 3 THERMAL
  // ----------------------------------------------------------

  if (
    name.includes(
      "supply 3 thermal"
    )
  ) {
    return {
      priority:
        "POMARAŃCZOWY",

      decision:
        "ZAPLANOWAĆ WIZYTĘ FE",

      problem:
        "Alert dotyczy układu termicznego Supply 3.",

      since:
        result.dateRange.start
          ? `AAA przedstawia dane od ${result.dateRange.start} do ${result.dateRange.end}.`
          : "Nie ustalono dokładnego początku problemu.",

      evidence:
        "Należy wspólnie ocenić oba thermistory, Thermistor Difference, Duty Cycle oraz Ambient Temperature. Wykresy AAA są dołączone do raportu.",

      interpretation:
        "Jeżeli oba thermistory przebiegają podobnie, a Thermistor Difference pozostaje niski, natomiast Duty Cycle często przekracza 85% lub osiąga 100%, bardziej prawdopodobnym kierunkiem jest spadek wydajności układu termicznego niż rozbieżność samych czujników.",

      cause:
        "Możliwy problem z wydajnością układu termicznego: przepływ powietrza, wentylator, zabrudzenie, cover/load door, heat pump lub jego połączenia. Thermistor należy podejrzewać szczególnie przy nieprawidłowej wartości lub dużej różnicy między czujnikami.",

      actions: [
        "Porównać Thermistor Right Side i Thermistor Left Side.",
        "Sprawdzić Thermistor Difference.",
        "Ocenić Duty Cycle, szczególnie okresy powyżej 85% i przy 100%.",
        "Ocenić Ambient Temperature.",
        "Sprawdzić warunki wentylacji analizatora.",
        "Sprawdzić cover/load door.",
        "Sprawdzić wentylator i zabrudzenia.",
        "Sprawdzić heat pump i jego połączenia zgodnie z procedurą serwisową."
      ],

      conclusion:
        "Zalecana planowana wizyta FE. Dokładną przyczynę należy potwierdzić na podstawie parametrów i kontroli mechaniczno-termicznej."
    };
  }

  // ----------------------------------------------------------
  // CM RING
  // ----------------------------------------------------------

  if (
    name.includes("cm") &&
    name.includes("ring")
  ) {
    return {
      priority:
        "POMARAŃCZOWY",

      decision:
        "ZAPLANOWAĆ WIZYTĘ FE",

      problem:
        "Alert dotyczy ruchu i pozycjonowania Slide (CM) Ring.",

      since:
        result.dateRange.start
          ? `AAA przedstawia trend od ${result.dateRange.start} do ${result.dateRange.end}.`
          : "Nie ustalono dokładnego początku problemu.",

      evidence:
        "Kluczowe jest porównanie Step Loss oraz Slot Corrections. Step Loss i liczba korekcji opisują różne zachowania mechanizmu i nie powinny być traktowane jako ten sam parametr.",

      interpretation:
        "Jeżeli Step Loss pozostaje stabilny, ale Slot Corrections rośnie lub utrzymuje się na podwyższonym poziomie, analizator nadal osiąga pozycję, lecz potrzebuje większej liczby korekcji. Taki obraz bardziej wskazuje na zwiększony opór lub pogarszającą się mechanikę pozycjonowania niż na klasyczne gubienie kroków.",

      cause:
        "Główne elementy do kontroli według AAA: CM Rotor belt i tracking, CM/RT Drive Motor oraz pinion, SAG rollers/bearings/spacers, Z-axis bearing pads oraz adjustment values związane z ruchem CM Ring.",

      actions: [
        "Ocenić trend Step Loss.",
        "Ocenić trend i zmienność Slot Corrections.",
        "Sprawdzić CM Rotor belt pod kątem zużycia i uszkodzeń.",
        "Sprawdzić belt tracking na CM/RT Drive Motor pinion.",
        "Sprawdzić CM/RT Drive Motor oraz pinion.",
        "Sprawdzić SAG rollers i bearings.",
        "Sprawdzić spacers SAG.",
        "Sprawdzić Z-axis bearing pads.",
        "Zweryfikować CM RING Stopping.",
        "Zweryfikować READ SYNC.",
        "Adjustment wykonywać dopiero po ocenie mechaniki i zgodnie z procedurą serwisową."
      ],

      conclusion:
        "Zalecana planowana wizyta FE. Wzrost Step Loss lub pojawienie się powiązanych condition codes powinno zwiększyć priorytet."
    };
  }

  // ----------------------------------------------------------
  // GENERIC
  // ----------------------------------------------------------

  return {
    priority:
      result.status === "ORANGE"
        ? "POMARAŃCZOWY"
        : "ŻÓŁTY",

    decision:
      result.status === "ORANGE"
        ? "WYMAGA OCENY SERWISOWEJ"
        : "MONITOROWAĆ",

    problem:
      "Wykryto aktywny alert bez dedykowanego modułu analitycznego.",

    since:
      "Nie ustalono dokładnego początku problemu.",

    evidence:
      "Dane AAA zostały zebrane, ale agent nie tworzy niepotwierdzonej diagnozy.",

    interpretation:
      "Alert wymaga oceny FSE.",

    cause:
      "Nie ustalono.",

    actions: [
      "Przeanalizować wykresy AAA.",
      "Sprawdzić Related Condition Codes.",
      "Skorelować alert z objawami klienta."
    ],

    conclusion:
      "Wymaga dalszej oceny."
  };
}

// ============================================================
// PDF
// ============================================================

function generatePDF(
  results,
  analyzerCount
) {
  return new Promise(
    (resolve, reject) => {

      const doc =
        new PDFDocument({
          size: "A4",
          margin: 40,

          info: {
            Title:
              "Raport serwisowy e-Connectivity v12",

            Author:
              "Service Triage Agent v12"
          }
        });

      const stream =
        fs.createWriteStream(
          PDF_FILE
        );

      stream.on(
        "finish",
        resolve
      );

      stream.on(
        "error",
        reject
      );

      doc.pipe(
        stream
      );

      let REGULAR =
        "Helvetica";

      let BOLD =
        "Helvetica-Bold";

      const fontCandidates = [
        [ARIAL, ARIAL_BOLD],
        ["C:\\Windows\\Fonts\\segoeui.ttf", "C:\\Windows\\Fonts\\segoeuib.ttf"],
        ["C:\\Windows\\Fonts\\calibri.ttf", "C:\\Windows\\Fonts\\calibrib.ttf"],
        ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"],
        ["/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf", "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"],
        ["/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"]
      ];

      const fontPair = fontCandidates.find(
        ([regular, bold]) => fs.existsSync(regular) && fs.existsSync(bold)
      );

      if (fontPair) {
        doc.registerFont("ArialPL", fontPair[0]);
        doc.registerFont("ArialPLBold", fontPair[1]);
        REGULAR = "ArialPL";
        BOLD = "ArialPLBold";
      } else {
        console.log(
          "UWAGA: nie znaleziono czcionki z polskimi znakami (Arial/Segoe/DejaVu) — PDF będzie bez polskich liter."
        );
      }

      function pageCheck(
        needed = 130
      ) {
        if (
          doc.y >
          doc.page.height -
          needed
        ) {
          doc.addPage();
        }
      }

      function heading(
        value,
        size = 15
      ) {
        doc
          .font(BOLD)
          .fontSize(size)
          .fillColor("#111111")
          .text(value);

        doc.moveDown(0.4);
      }

      function field(
        label,
        value
      ) {
        doc
          .font(BOLD)
          .fontSize(9.5)
          .fillColor("#222222")
          .text(
            `${label}: `,
            {
              continued: true
            }
          );

        doc
          .font(REGULAR)
          .text(
            value ||
            "brak danych"
          );
      }

      function section(
        title,
        body
      ) {
        pageCheck(110);

        doc.moveDown(0.6);

        doc
          .font(BOLD)
          .fontSize(10.5)
          .fillColor("#111111")
          .text(title);

        doc.moveDown(0.15);

        doc
          .font(REGULAR)
          .fontSize(9.5)
          .fillColor("#333333")
          .text(
            body ||
            "Brak danych.",
            {
              lineGap: 2
            }
          );
      }

      function actions(list) {
        pageCheck(170);

        doc.moveDown(0.7);

        doc
          .font(BOLD)
          .fontSize(10.5)
          .fillColor("#111111")
          .text(
            "CO POWINIEN ZROBIĆ FSE"
          );

        doc.moveDown(0.3);

        list.forEach(
          (value, index) => {
            pageCheck(50);

            doc
              .font(REGULAR)
              .fontSize(9.5)
              .fillColor("#333333")
              .text(
                `${index + 1}. ${value}`,
                {
                  indent: 8,
                  lineGap: 2
                }
              );

            doc.moveDown(0.15);
          }
        );
      }

      // ======================================================
      // COVER
      // ======================================================

      doc
        .font(BOLD)
        .fontSize(23)
        .fillColor("#111111")
        .text(
          "RAPORT SERWISOWY",
          {
            align: "center"
          }
        );

      doc
        .fontSize(17)
        .text(
          "QuidelOrtho e-Connectivity",
          {
            align: "center"
          }
        );

      doc.moveDown(0.5);

      doc
        .font(REGULAR)
        .fontSize(9)
        .fillColor("#555555")
        .text(
          `Wygenerowano: ${formatDate(new Date())}`,
          {
            align: "center"
          }
        );

      doc.moveDown(1.5);

      heading(
        "PODSUMOWANIE",
        15
      );

      field(
        "Sprawdzone analizatory",
        String(
          analyzerCount
        )
      );

      field(
        "Aktywne alerty",
        String(
          results.length
        )
      );

      doc.moveDown(1);

      for (
        const item
        of results
      ) {
        pageCheck(110);

        doc
          .font(BOLD)
          .fontSize(11)
          .fillColor("#111111")
          .text(
            `${item.model} — J${item.jno}`
          );

        doc
          .font(REGULAR)
          .fontSize(9.5)
          .text(
            item.location.customer
          );

        doc
          .fontSize(9)
          .fillColor("#555555")
          .text(
            `${item.location.city} | ${item.alertName}`
          );

        doc
          .font(BOLD)
          .fontSize(9.5)
          .fillColor("#111111")
          .text(
            `DECYZJA: ${item.analysis.decision}`
          );

        doc.moveDown(0.7);
      }

      // ======================================================
      // DETAILS
      // ======================================================

      for (
        const item
        of results
      ) {
        const a =
          item.analysis;

        doc.addPage();

        doc
          .font(BOLD)
          .fontSize(20)
          .fillColor("#111111")
          .text(
            item.model
          );

        doc
          .font(BOLD)
          .fontSize(15)
          .text(
            `J-number: ${item.jno}`
          );

        doc.moveDown(0.5);

        heading(
          "LOKALIZACJA",
          11
        );

        field(
          "Placówka",
          item.location.customer
        );

        field(
          "Adres",
          item.location.address
        );

        field(
          "Miasto",
          item.location.city
        );

        if (
          item.asset
        ) {
          field(
            "Status w bazie",
            item.asset.excelStatus
          );

          field(
            "Software",
            item.asset.softwareVersion
          );

          field(
            "Product family",
            item.asset.productFamily
          );
        }

        doc.moveDown(0.7);

        heading(
          item.alertName,
          14
        );

        doc
          .font(BOLD)
          .fontSize(13)
          .text(
            `DECYZJA: ${a.decision}`
          );

        doc.moveDown(0.5);

        field(
          "Status e-Connectivity",
          item.status
        );

        field(
          "Priorytet serwisowy",
          a.priority
        );

        if (
          item.dateRange.start
        ) {
          field(
            "Zakres danych AAA",
            `${item.dateRange.start} – ${item.dateRange.end}`
          );
        }

        section(
          "PROBLEM",
          a.problem
        );

        section(
          "OD KIEDY",
          a.since
        );

        section(
          "CO WIDAĆ W DANYCH",
          a.evidence
        );

        section(
          "INTERPRETACJA",
          a.interpretation
        );

        section(
          "NAJBARDZIEJ PRAWDOPODOBNY KIERUNEK",
          a.cause
        );

        actions(
          a.actions
        );

        section(
          "DECYZJA / PILNOŚĆ",
          a.conclusion
        );

        // ====================================================
        // DATA LOGGER
        // ====================================================

        if (
          item.dataLogger
            ?.detected
        ) {
          pageCheck(150);

          doc.moveDown(0.8);

          heading(
            "DANE e-CONNECTIVITY",
            11
          );

          field(
            "Ostatnie połączenie",
            item.dataLogger
              .lastConnected
          );

          if (
            item.dataLogger
              .aFile.length
          ) {
            field(
              "A-file",
              item.dataLogger
                .aFile
                .map(
                  x =>
                    `${x.day}: ${x.value}`
                )
                .join(" | ")
            );
          }

          if (
            item.dataLogger
              .bFile.length
          ) {
            field(
              "B-file",
              item.dataLogger
                .bFile
                .map(
                  x =>
                    `${x.day}: ${x.value}`
                )
                .join(" | ")
            );
          }
        }

        // ====================================================
        // CHARTS
        // ====================================================

        if (
          item.chartFiles.length
        ) {
          doc.addPage();

          heading(
            "WYKRESY DIAGNOSTYCZNE AAA",
            14
          );

          doc
            .font(REGULAR)
            .fontSize(8.5)
            .fillColor("#555555")
            .text(
              "Wykresy źródłowe wykorzystane podczas oceny alertu."
            );

          doc.moveDown(0.6);

          const allFindings = (item.chartData || [])
            .flatMap(c => describeLocalChart(c.local, item.dateRange).findings);

          if (allFindings.length) {
            doc
              .font(BOLD)
              .fontSize(9.5)
              .fillColor("#1a6b3a")
              .text("KLUCZOWE ODCZYTY (porównanie z progami na wykresach)");

            doc
              .font(REGULAR)
              .fontSize(8.5)
              .fillColor("#222222");

            for (const line of allFindings.slice(0, 12)) {
              pageCheck(60);
              doc.text(`• ${line}`);
            }

            doc.moveDown(0.6);
          }

          for (
            let index = 0;
            index <
            item.chartFiles.length;
            index++
          ) {
            const file =
              item.chartFiles[
                index
              ];

            if (
              !fs.existsSync(file)
            ) {
              continue;
            }

            pageCheck(315);

            try {
              doc
                .font(BOLD)
                .fontSize(9)
                .fillColor("#333333")
                .text(
                  `Wykres ${index + 1}`
                );

              doc.moveDown(0.2);

              const chartInfo =
                (item.chartData || [])[index] || null;

              if (chartInfo && chartInfo.title) {
                doc
                  .font(REGULAR)
                  .fontSize(8.5)
                  .fillColor("#333333")
                  .text(chartInfo.title);
              }

              doc.image(
                file,
                {
                  fit: [
                    500,
                    270
                  ],

                  align:
                    "center"
                }
              );

              doc.moveDown(0.3);

              // --- Odczyt lokalny (z obrazu, bez wysyłania na zewnątrz) ---
              if (chartInfo && chartInfo.local && chartInfo.local.ok) {
                const d = describeLocalChart(chartInfo.local, item.dateRange);

                pageCheck(120);

                doc
                  .font(BOLD)
                  .fontSize(8.5)
                  .fillColor("#1a6b3a")
                  .text("Odczyt wartości z wykresu (lokalnie, z obrazu):");

                doc
                  .font(REGULAR)
                  .fontSize(7.8)
                  .fillColor("#222222");

                for (const line of d.lines) {
                  pageCheck(40);
                  doc.text(line.startsWith("   ") ? "      " + line.trim() : "• " + line);
                }
              } else if (chartInfo && chartInfo.local && !chartInfo.local.ok && chartInfo.local.reason) {
                doc
                  .font(REGULAR)
                  .fontSize(7.5)
                  .fillColor("#999999")
                  .text(`Odczyt lokalny niemożliwy: ${chartInfo.local.reason}.`);
              }

              // --- Wartości z wykresu ---
              if (chartInfo && chartInfo.points && chartInfo.points.length) {
                const s = chartInfo.summary || {};

                doc
                  .font(BOLD)
                  .fontSize(8.5)
                  .fillColor("#1a4d8f")
                  .text(
                    `Wartości odczytane ze strony (${chartInfo.points.length} pkt)` +
                    (s.min !== undefined
                      ? ` — min: ${s.min}, max: ${s.max}, średnia: ${s.avg}, ostatnia: ${s.last}`
                      : "")
                  );

                doc
                  .font(REGULAR)
                  .fontSize(7.5)
                  .fillColor("#333333");

                for (const p of chartInfo.points.slice(0, 24)) {
                  pageCheck(60);
                  doc.text(
                    `• ${p.label}` +
                    (p.value !== null && p.value !== undefined && !/\d/.test(String(p.label))
                      ? `  [${p.value}]`
                      : "")
                  );
                }

                if (chartInfo.points.length > 24) {
                  doc.text(
                    `… oraz ${chartInfo.points.length - 24} kolejnych punktów (pełna lista w data.json)`
                  );
                }
              }

              if (chartInfo && chartInfo.vision) {
                const v = chartInfo.vision;

                doc
                  .font(BOLD)
                  .fontSize(8.5)
                  .fillColor("#8a5a00")
                  .text(
                    "Odczyt z obrazu (AI) — szacunkowy, zweryfikuj z wykresem:"
                  );

                doc
                  .font(REGULAR)
                  .fontSize(7.5)
                  .fillColor("#333333");

                if (v.x_axis || v.y_axis) {
                  doc.text(
                    `Oś X: ${v.x_axis || "?"} | Oś Y: ${v.y_axis || "?"}`
                  );
                }

                for (const serie of (v.series || []).slice(0, 8)) {
                  doc.text(
                    `• ${serie.name || "seria"}: min ${serie.min ?? "?"}, max ${serie.max ?? "?"}, ostatnia ${serie.last ?? "?"}` +
                    (Array.isArray(serie.labeled_values) && serie.labeled_values.length
                      ? `, opisane: ${serie.labeled_values.slice(0, 12).join("; ")}`
                      : "")
                  );
                }

                if (v.notes) {
                  doc.text(`Uwagi: ${cleanText(v.notes)}`);
                }
              }

              if (
                chartInfo &&
                !(chartInfo.points && chartInfo.points.length) &&
                !chartInfo.vision &&
                !(chartInfo.local && chartInfo.local.ok)
              ) {
                doc
                  .font(REGULAR)
                  .fontSize(7.5)
                  .fillColor("#999999")
                  .text(
                    "Brak odczytu wartości z tego wykresu. Zainstaluj pakiety: npm install pngjs tesseract.js @tesseract.js-data/eng (odczyt lokalny) lub włącz AAA_VISION=1."
                  );
              }

              doc.moveDown(0.7);

            } catch {}
          }
        }
      }

      // ======================================================
      // PORÓWNANIE ANALIZATORÓW (ten sam typ wykresu)
      // ======================================================

      const comparison = buildChartComparison(results);

      if (comparison.length) {
        doc.addPage();

        heading(
          "PORÓWNANIE ANALIZATORÓW — TE SAME WYKRESY",
          14
        );

        doc
          .font(REGULAR)
          .fontSize(8.5)
          .fillColor("#555555")
          .text("Wartości odczytane lokalnie z obrazów wykresów (pierwsza seria na wykresie). Odczyt przybliżony.");

        doc.moveDown(0.6);

        for (const group of comparison) {
          pageCheck(120);

          doc
            .font(BOLD)
            .fontSize(9.5)
            .fillColor("#1a4d8f")
            .text(group.title.replace(/\s*--?\s*[A-Z]?\d{5,}\s*$/i, ""));

          doc
            .font(REGULAR)
            .fontSize(8.3)
            .fillColor("#222222");

          for (const row of group.rows) {
            const f = n => fmtChartNum(n, row.res);
            doc.text(
              `• J${row.jno}: ostatnia ${f(row.last)}, średnia ${f(row.avg)}, min ${f(row.min)}, maks ${f(row.max)}, trend ${signedChartNum(row.trend, row.res)}` +
              (row.thr.length ? `, linie: ${row.thr.map(f).join(" / ")}` : "")
            );
          }

          doc.moveDown(0.5);
        }
      }

      // ======================================================
      // READ ONLY
      // ======================================================

      doc.addPage();

      heading(
        "TRYB READ-ONLY",
        15
      );

      doc
        .font(REGULAR)
        .fontSize(10)
        .fillColor("#333333")
        .text(
          "Agent analizuje dane e-Connectivity i AAA, ale nie wykonuje zmian na analizatorze."
        );

      doc.moveDown(0.7);

      doc.text(
        "Agent NIE wykonuje:\n\n" +
        "• Save\n" +
        "• Apply\n" +
        "• zmian konfiguracji\n" +
        "• automatycznych adjustmentów\n" +
        "• resetów\n" +
        "• service commands"
      );

      doc.end();
    }
  );
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
                result.serviceDocuments
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
      genericLiveAnalysis(
        result
      );

    console.log(
      `DECYZJA: ${result.analysis.decision}`
    );

    results.push(
      result
    );

    await sleep(
      500
    );
  }

  // ==========================================================
  // MASTER JSON
  // ==========================================================

  fs.writeFileSync(
    JSON_FILE,

    JSON.stringify(
      {
        version:
          "LIVE-v1",

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
    state.analyzerCount
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
      `  DECYZJA: ${result.analysis.decision}`
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
  genericLiveAnalysis,
  PDF_FILE,
  OUTPUT_DIR
};
