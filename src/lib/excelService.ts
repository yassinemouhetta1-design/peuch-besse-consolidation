import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

export interface ArticleDetail {
  designation: string;
  appellation: string;
  color: string;
  millesime: string;
  taille: string;
  quantity: number;
  price: number;
  total: number;
}

export interface ConsolidationResult {
  fileName: string;
  clientName: string;
  articlesTotal: number;
  transport: number;
  globalTotal: number;
  totalBottles: number;
  isMatch: boolean;
  error?: string;
  articles?: ArticleDetail[];
}

export interface GlobalVerification {
  grandTotalSource: number;
  grandTotalGlobal: number;
  isMatch: boolean;
  totalBottles: number;
}

/**
 * Normalizes a string for comparison: trim, lowercase, remove special chars, replace newlines
 */
function normalize(str: string): string {
  return str
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // Remove accents
    .replace(/[^a-z0-9]/g, " ") // Replace non-alphanumeric with space
    .replace(/\s+/g, " ") // Collapse spaces
    .trim();
}

/**
 * Robustly parses a number from an Excel cell value, handling formulas
 */
function parseExcelNumber(value: any): number {
  if (value === null || value === undefined) return 0;
  
  // Handle ExcelJS formula objects
  let val = value;
  if (typeof value === 'object' && 'result' in value) {
    val = value.result;
  }
  
  if (typeof val === 'number') return val;
  
  // Clean string: remove currency symbols, spaces, and handle comma as decimal separator
  const clean = val.toString()
    .replace(/[^\d,.-]/g, '') // Keep only digits, comma, dot, minus
    .replace(/\s/g, '')       // Remove any remaining spaces
    .replace(',', '.');       // Convert comma to dot
    
  const parsed = parseFloat(clean);
  return isNaN(parsed) ? 0 : parsed;
}

export async function analyzeSourceFiles(
  sourceFiles: File[]
): Promise<ConsolidationResult[]> {
  const results: ConsolidationResult[] = [];

  for (const file of sourceFiles) {
    try {
      const extension = file.name.split('.').pop()?.toLowerCase();
      let sourceSheet: any;

      if (extension === 'xls' || extension === 'xlsb') {
        const data = await file.arrayBuffer();
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        
        sourceSheet = {
          getRow: (rowNum: number) => ({
            getCell: (colNum: number) => {
              const cellAddress = XLSX.utils.encode_cell({ r: rowNum - 1, c: colNum - 1 });
              const cell = worksheet[cellAddress];
              return { value: cell ? cell.v : null };
            }
          }),
          rowCount: range.e.r + 1
        };
      } else {
        const sourceWorkbook = new ExcelJS.Workbook();
        await sourceWorkbook.xlsx.load(await file.arrayBuffer());
        sourceSheet = sourceWorkbook.worksheets[0];
      }

      // Find Header "Désignation" / "Designation" (FR + EN)
      // On normalise le texte pour supprimer les accents et gérer les deux langues
      const stripAccents = (s: string) =>
        s.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

      // Mots-clés acceptés : français ET anglais
      const HEADER_KEYWORDS = ['designation', 'description', 'product', 'produit', 'libelle', 'article'];

      let sourceHeaderRow = -1;
      let designationCol = 2;
      for (let r = 1; r <= 50; r++) {
        for (let c = 1; c <= 5; c++) {
          const raw = sourceSheet.getRow(r).getCell(c).value?.toString() || '';
          const normalized = stripAccents(raw);
          if (HEADER_KEYWORDS.some(kw => normalized.includes(kw))) {
            sourceHeaderRow = r;
            designationCol = c;
            break;
          }
        }
        if (sourceHeaderRow !== -1) break;
      }

      if (sourceHeaderRow === -1) throw new Error(
        `En-tête non trouvé. Le fichier doit contenir une colonne "Désignation" (FR) ou "Designation" / "Product" (EN) dans les 50 premières lignes.`
      );
      const sourceFirstArticleRow = sourceHeaderRow + 2;

      // Client Info
      const nom = sourceSheet.getRow(4).getCell(7).value; // G4
      const prenom = sourceSheet.getRow(5).getCell(7).value; // G5
      const clientName = (nom instanceof Date) ? '14 Juillet' : `${nom?.toString() || ''} ${prenom?.toString() || ''}`.trim();

      // Transport (H185)
      const transportValue = parseExcelNumber(sourceSheet.getRow(185).getCell(8).value);

      // Articles
      const articles: ArticleDetail[] = [];
      let articlesTotal = 0;
      let totalBottles = 0;
      
      // Strict loop: only rows with text in Column B (Désignation) are included
      for (let srcRow = sourceFirstArticleRow; srcRow <= sourceSheet.rowCount; srcRow++) {
        const row = sourceSheet.getRow(srcRow);
        
        // Column B is the Designation
        const rawDesignation = row.getCell(2).value;
        let designation = '';
        if (rawDesignation && typeof rawDesignation === 'object' && 'richText' in rawDesignation) {
          designation = (rawDesignation as any).richText.map((rt: any) => rt.text).join('').trim();
        } else {
          designation = rawDesignation?.toString().trim() || '';
        }
        
        // Condition d'exclusion: If Column B is EMPTY, ignore the row
        if (!designation) continue;

        // Column M is the Montant HT (Total for the line)
        const lineTotal = parseExcelNumber(row.getCell(13).value);
        // Column L is the Number of Cartons
        const qtyL = parseExcelNumber(row.getCell(12).value);
        // Column I is the Bottles per Carton
        const btlPerCarton = parseExcelNumber(row.getCell(9).value) || 1;
        
        // Condition d'inclusion stricte: Only keep if there is a quantity OR an amount
        if (lineTotal <= 0 && qtyL <= 0) continue;

        const lineBottles = Math.round(qtyL * btlPerCarton);

        // Fix [object Object] bug by ensuring we get the text value (handles rich text)
        const rawAppellation = row.getCell(4).value;
        let appellation = '';
        if (rawAppellation && typeof rawAppellation === 'object' && 'richText' in rawAppellation) {
          appellation = (rawAppellation as any).richText.map((rt: any) => rt.text).join('');
        } else {
          appellation = rawAppellation?.toString() || '';
        }

        const rawColor = row.getCell(6).value; // Col F
        let color = '';
        if (rawColor && typeof rawColor === 'object' && 'richText' in rawColor) {
          color = (rawColor as any).richText.map((rt: any) => rt.text).join('');
        } else {
          color = rawColor?.toString() || '';
        }

        const rawMillesime = row.getCell(8).value; // Col H
        let millesime = '';
        if (rawMillesime && typeof rawMillesime === 'object' && 'richText' in rawMillesime) {
          millesime = (rawMillesime as any).richText.map((rt: any) => rt.text).join('');
        } else {
          millesime = rawMillesime?.toString() || '';
        }

        const rawTaille = row.getCell(9).value; // Col I
        let taille = '';
        if (rawTaille && typeof rawTaille === 'object' && 'richText' in rawTaille) {
          taille = (rawTaille as any).richText.map((rt: any) => rt.text).join('');
        } else {
          taille = rawTaille?.toString() || '';
        }

        articlesTotal += lineTotal;
        totalBottles += lineBottles;

        articles.push({
          designation,
          appellation,
          color,
          millesime,
          taille,
          quantity: lineBottles,
          price: lineBottles > 0 ? lineTotal / lineBottles : 0,
          total: lineTotal
        });
      }

      results.push({
        fileName: file.name,
        clientName: clientName || file.name,
        articlesTotal: articlesTotal,
        transport: transportValue,
        globalTotal: articlesTotal + transportValue,
        totalBottles: totalBottles,
        isMatch: true,
        articles
      });

    } catch (err) {
      results.push({
        fileName: file.name,
        clientName: file.name,
        articlesTotal: 0,
        transport: 0,
        globalTotal: 0,
        totalBottles: 0,
        isMatch: false,
        error: err instanceof Error ? err.message : String(err)
      });
    }
  }

  return results;
}
