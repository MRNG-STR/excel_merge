import * as XLSX from 'xlsx';

interface MergedCell {
  sheetName: string;
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
}

function getMergedCells(filePath: string): MergedCell[] {
  const workbook = XLSX.readFile(filePath);
  const sheets = workbook.SheetNames;

  const mergedCells: MergedCell[] = [];

  for (const sheetName of sheets) {
    const worksheet = workbook.Sheets[sheetName];
    const merged = worksheet['!merges'];

    if (merged) {
      for (const merge of merged) {
        mergedCells.push({
          sheetName,
          startRow: merge.s.r,
          endRow: merge.e.r,
          startColumn: merge.s.c,
          endColumn: merge.e.c,
        });
      }
    }
  }

  return mergedCells;
}

const filePath = process.argv[2];

if (!filePath) {
  console.error('Please provide an Excel file path as a command-line argument');
  process.exit(1);
}

const mergedCells = getMergedCells(filePath);

console.log('Merged Cells:'+'\n');
for (const cell of mergedCells) {
  console.log(
    `Sheet: ${cell.sheetName}, ` +
    `Start Row: ${cell.startRow}, ` +
    `End Row: ${cell.endRow}, ` +
    `Start Column: ${cell.startColumn}, ` +
    `End Column: ${cell.endColumn}`
  );
  console.log(`Range: ${cell.startRow},${cell.startColumn} To ${cell.endRow},${cell.endColumn}`+'\n');
}
