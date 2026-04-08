import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

export interface MatchedArticle {
  designation: string;
  appellation: string;
  color: string;
  millesime: string;
  taille: string;
  qty: number;
}

export interface MatchReport {
  clientName: string;
  fileName: string;
  matched: number;
  matchedArticles: MatchedArticle[];
  unmatched: string[];
  totalBottles: number;
}

export interface ConsolidationReport {
  reports: MatchReport[];
  totalClients: number;
  totalMatched: number;
  totalUnmatched: number;
}

/**
 * Normalize a string for fuzzy matching: trim, lowercase, remove accents, collapse spaces
 */
function normalize(str: string): string {
  return str
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Traduit les couleurs anglaises en français pour le matching avec la MATRICE
 * Fichiers EN : white, red, pink, rosé → Fichiers FR / MATRICE : blanc, rouge, rosé
 */
function normalizeColor(color: string): string {
  const n = normalize(color);
  // Traductions EN → FR
  if (n === 'white' || n === 'blanc') return 'blanc';
  if (n === 'red'   || n === 'rouge') return 'rouge';
  if (n === 'pink'  || n === 'rose'  || n === 'rose'  || n.includes('rose')) return 'rose';
  if (n === 'orange') return 'orange';
  if (n === 'sparkling' || n === 'effervescent') return 'effervescent';
  return n; // valeur normalisée brute si non reconnue
}

/**
 * Parse a numeric value from an Excel cell, handling formulas and rich text
 */
function parseNum(value: any): number {
  if (value === null || value === undefined) return 0;
  let val = value;
  if (typeof value === 'object' && 'result' in value) val = value.result;
  // Si richText → c'est un label/description, PAS un nombre → retourner 0
  // (évite d'extraire des chiffres parasites ex: "min 50 coffrets" → 50)
  if (typeof value === 'object' && 'richText' in value) return 0;
  if (typeof val === 'number') return val;
  const clean = val.toString().replace(/[^\d,.\-]/g, '').replace(',', '.');
  const parsed = parseFloat(clean);
  return isNaN(parsed) ? 0 : parsed;
}

/**
 * Get the string value from a cell, handling rich text objects
 */
function cellText(value: any): string {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object' && 'richText' in value) {
    return (value as any).richText.map((rt: any) => rt.text).join('').trim();
  }
  if (typeof value === 'object' && 'result' in value) {
    return value.result?.toString().trim() || '';
  }
  return value.toString().trim();
}

interface ParsedSourceFile {
  clientName: string;
  fileName: string;
  articles: { designation: string; appellation: string; color: string; millesime: string; taille: string; qty: number }[];
}

/**
 * Parse a single source file and extract client name + articles with qty
 */
async function parseSourceFile(file: File): Promise<ParsedSourceFile> {
  const extension = file.name.split('.').pop()?.toLowerCase();
  let sheet: any;

  if (extension === 'xls' || extension === 'xlsb') {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data, { type: 'array' });
    const ws = workbook.Sheets[workbook.SheetNames[0]];
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1');
    sheet = {
      getRow: (rowNum: number) => ({
        getCell: (colNum: number) => {
          const addr = XLSX.utils.encode_cell({ r: rowNum - 1, c: colNum - 1 });
          const cell = ws[addr];
          return { value: cell ? cell.v : null };
        },
      }),
      rowCount: range.e.r + 1,
    };
  } else {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await file.arrayBuffer());
    sheet = wb.worksheets[0];
  }

  // Client name from G4 (+ G5)
  const nom = sheet.getRow(4).getCell(7).value;
  const prenom = sheet.getRow(5).getCell(7).value;
  const clientName =
    nom instanceof Date
      ? '14 Juillet'
      : `${cellText(nom)} ${cellText(prenom)}`.trim();

  // ── Détection dynamique du header (FR: "Désignation" / EN: "Designation", "Product", etc.) ──
  const stripAccents = (s: string) =>
    s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const HEADER_KEYWORDS = ['designation', 'description', 'product', 'produit', 'libelle', 'article'];

  let headerRow = 13; // fallback ligne 13 (données démarrent à 14)
  for (let r = 1; r <= 50; r++) {
    for (let c = 1; c <= 5; c++) {
      const raw = sheet.getRow(r).getCell(c).value?.toString() || '';
      if (HEADER_KEYWORDS.some(kw => stripAccents(raw).includes(kw))) {
        headerRow = r;
        break;
      }
    }
    if (headerRow !== 13 || r === 13) break;
  }
  const firstDataRow = headerRow + 2; // ligne header + 1 ligne vide/sous-header + données
  const lastDataRow  = Math.min(headerRow + 170, 250); // max ~170 produits

  // ── Extraction des articles ──────────────────────────────────────────────
  const articles: { designation: string; appellation: string; color: string; millesime: string; taille: string; qty: number }[] = [];

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const row = sheet.getRow(r);
    const rawDesignation = row.getCell(2).value; // Col B
    const designation = cellText(rawDesignation);
    if (!designation) continue;

    const cartons = parseNum(row.getCell(12).value); // Col L
    const lineTotal = parseNum(row.getCell(13).value); // Col M — Montant HT
    const btlPerCarton = parseNum(row.getCell(9).value) || 1; // Col I, default 1
    const qty = Math.round(cartons * btlPerCarton);

    // Double protection : skip les lignes de titre/section
    if (lineTotal <= 0 && cartons <= 0) continue;

    const appellation = cellText(row.getCell(4).value); // Col D
    const color = cellText(row.getCell(6).value); // Col F
    const millesime = cellText(row.getCell(8).value); // Col H
    const taille = cellText(row.getCell(9).value); // Col I

    if (qty > 0) {
      articles.push({ designation, appellation, color, millesime, taille, qty });
    }
  }

  return { clientName: clientName || file.name, fileName: file.name, articles };
}

/**
 * Main consolidation: inject all source files into the MATRICE template
 */
export async function consolidateToMatrix(
  sourceFiles: File[],
  templateFile: File,
  onProgress?: (current: number, total: number, clientName: string) => void
): Promise<{ blob: Blob; report: ConsolidationReport }> {
  // 1. Load the MATRICE template with ExcelJS (preserves formulas, styles, etc.)
  const templateWb = new ExcelJS.Workbook();
  await templateWb.xlsx.load(await templateFile.arrayBuffer());
  const globalSheet = templateWb.worksheets[0];

  // 2. Build lookup maps.
  //
  // Règle : quand la taille est présente dans la commande, elle est OBLIGATOIRE dans le match —
  // pas de fallback vers une clé sans taille (évite de matcher le mauvais format).
  // Quand la taille est absente, on utilise les maps sans taille.
  //
  // Maps AVEC taille (utilisées si nt !== '')
  const mapT_DACMT = new Map<string, number>(); // desig|appel|color|mill|taille  (match complet)
  const mapT_DACT  = new Map<string, number>(); // desig|appel|color|taille       (sans mill)
  const mapT_DAT   = new Map<string, number>(); // desig|appel|taille             (sans color)
  const mapT_DT    = new Map<string, number>(); // desig|taille                   (fallback minimal)
  //
  // Maps SANS taille (utilisées si nt === '')
  const mapDACM    = new Map<string, number>(); // desig|appel|color|mill
  const mapDAC     = new Map<string, number>(); // desig|appel|color
  const mapDA      = new Map<string, number>(); // desig|appel
  const mapD       = new Map<string, number>(); // desig only

  for (let r = 4; r <= 172; r++) {
    const row = globalSheet.getRow(r);
    const desig  = normalize(cellText(row.getCell(1).value)); // Col A
    const appel  = normalize(cellText(row.getCell(2).value)); // Col B
    const color  = normalize(cellText(row.getCell(3).value)); // Col C
    const mill   = normalize(cellText(row.getCell(5).value)); // Col E — Millésime
    const taille = normalize(cellText(row.getCell(6).value)); // Col F — Taille

    // Skip les lignes de titre/section : elles n'ont que col A remplie.
    // Un vrai produit a toujours au minimum une appellation, une couleur OU une taille.
    if (!desig) continue;
    if (!appel && !color && !taille) continue; // ← ligne titre de section → ignorée

    // Maps avec taille
    mapT_DACMT.set(`${desig}|${appel}|${color}|${mill}|${taille}`, r);
    if (!mapT_DACT.has(`${desig}|${appel}|${color}|${taille}`)) mapT_DACT.set(`${desig}|${appel}|${color}|${taille}`, r);
    if (!mapT_DAT.has(`${desig}|${appel}|${taille}`)) mapT_DAT.set(`${desig}|${appel}|${taille}`, r);
    if (!mapT_DT.has(`${desig}|${taille}`)) mapT_DT.set(`${desig}|${taille}`, r);

    // Maps sans taille
    if (!mapDACM.has(`${desig}|${appel}|${color}|${mill}`)) mapDACM.set(`${desig}|${appel}|${color}|${mill}`, r);
    if (!mapDAC.has(`${desig}|${appel}|${color}`)) mapDAC.set(`${desig}|${appel}|${color}`, r);
    if (!mapDA.has(`${desig}|${appel}`)) mapDA.set(`${desig}|${appel}`, r);
    if (!mapD.has(desig)) mapD.set(desig, r);
  }

  // 3. Parse all source files first
  const parsedSources: ParsedSourceFile[] = [];
  for (let i = 0; i < sourceFiles.length; i++) {
    const parsed = await parseSourceFile(sourceFiles[i]);
    parsedSources.push(parsed);
    onProgress?.(i + 1, sourceFiles.length * 2, parsed.clientName);
  }

  // 4. Inject each source into the template
  const reports: MatchReport[] = [];
  let totalMatched = 0;
  let totalUnmatched = 0;

  for (let i = 0; i < parsedSources.length; i++) {
    const source = parsedSources[i];
    onProgress?.(
      sourceFiles.length + i + 1,
      sourceFiles.length * 2,
      source.clientName
    );

    // Column index for this client (0-based): bottles = 11 + i*2, price = 12 + i*2
    // ExcelJS columns are 1-based: bottles = 12 + i*2
    const bottlesCol = 12 + i * 2; // L=12, N=14, P=16...

    // Write client name in Row 1
    globalSheet.getRow(1).getCell(bottlesCol).value = source.clientName;

    // Match and inject
    const unmatched: string[] = [];
    const matchedArticles: MatchedArticle[] = [];
    let matched = 0;
    let totalBottles = 0;

    for (const article of source.articles) {
      const nd = normalize(article.designation);
      const na = normalize(article.appellation);
      const nc = normalizeColor(article.color); // traduit EN→FR : white→blanc, red→rouge...
      const nm = normalize(article.millesime);
      const nt = normalize(article.taille);

      // Si taille présente → elle est obligatoire dans le match, pas de fallback sans elle
      // Si taille absente → fallbacks classiques sans contrainte taille
      const targetRow = nt
        ? mapT_DACMT.get(`${nd}|${na}|${nc}|${nm}|${nt}`) ??
          mapT_DACT.get(`${nd}|${na}|${nc}|${nt}`) ??
          mapT_DAT.get(`${nd}|${na}|${nt}`) ??
          mapT_DT.get(`${nd}|${nt}`)
        : mapDACM.get(`${nd}|${na}|${nc}|${nm}`) ??
          mapDAC.get(`${nd}|${na}|${nc}`) ??
          mapDA.get(`${nd}|${na}`) ??
          mapD.get(nd);

      if (targetRow !== undefined) {
        globalSheet.getRow(targetRow).getCell(bottlesCol).value = article.qty;
        matched++;
        totalBottles += article.qty;
        matchedArticles.push({
          designation: article.designation,
          appellation: article.appellation,
          color: article.color,
          millesime: article.millesime,
          taille: article.taille,
          qty: article.qty,
        });
      } else {
        unmatched.push(article.designation);
      }
    }

    totalMatched += matched;
    totalUnmatched += unmatched.length;

    reports.push({
      clientName: source.clientName,
      fileName: source.fileName,
      matched,
      matchedArticles,
      unmatched,
      totalBottles,
    });
  }

  // 5. Generate the output Excel as a Blob
  const buffer = await templateWb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });

  return {
    blob,
    report: {
      reports,
      totalClients: parsedSources.length,
      totalMatched,
      totalUnmatched,
    },
  };
}
