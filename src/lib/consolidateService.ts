import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

export interface MatchReport {
  clientName: string;
  fileName: string;
  matched: number;
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
 * Parse a numeric value from an Excel cell, handling formulas and rich text
 */
function parseNum(value: any): number {
  if (value === null || value === undefined) return 0;
  let val = value;
  if (typeof value === 'object' && 'result' in value) val = value.result;
  if (typeof value === 'object' && 'richText' in value) {
    val = (value as any).richText.map((rt: any) => rt.text).join('');
  }
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
  articles: { designation: string; qty: number }[];
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

  // Extract articles from rows 14-179
  const articles: { designation: string; qty: number }[] = [];

  for (let r = 14; r <= 179; r++) {
    const row = sheet.getRow(r);
    const rawDesignation = row.getCell(2).value; // Col B
    const designation = cellText(rawDesignation);
    if (!designation) continue;

    const cartons = parseNum(row.getCell(12).value); // Col L
    const btlPerCarton = parseNum(row.getCell(9).value) || 1; // Col I, default 1
    const qty = Math.round(cartons * btlPerCarton);

    if (qty > 0) {
      articles.push({ designation, qty });
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

  // 2. Build a lookup map: normalized designation → row number in the global
  const designationMap = new Map<string, number>();
  for (let r = 4; r <= 172; r++) {
    const cellVal = globalSheet.getRow(r).getCell(1).value; // Col A
    const text = cellText(cellVal);
    if (text) {
      designationMap.set(normalize(text), r);
    }
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
    let matched = 0;
    let totalBottles = 0;

    for (const article of source.articles) {
      const normalizedDesig = normalize(article.designation);
      const targetRow = designationMap.get(normalizedDesig);

      if (targetRow !== undefined) {
        // Inject qty as Number in the bottles column
        globalSheet.getRow(targetRow).getCell(bottlesCol).value = article.qty;
        matched++;
        totalBottles += article.qty;
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
