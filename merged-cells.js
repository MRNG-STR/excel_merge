"use strict";
exports.__esModule = true;
var XLSX = require("xlsx");
function getMergedCells(filePath) {
    var workbook = XLSX.readFile(filePath);
    var sheets = workbook.SheetNames;
    var mergedCells = [];
    for (var _i = 0, sheets_1 = sheets; _i < sheets_1.length; _i++) {
        var sheetName = sheets_1[_i];
        var worksheet = workbook.Sheets[sheetName];
        var merged = worksheet['!merges'];
        if (merged) {
            for (var _a = 0, merged_1 = merged; _a < merged_1.length; _a++) {
                var merge = merged_1[_a];
                mergedCells.push({
                    sheetName: sheetName,
                    startRow: merge.s.r,
                    endRow: merge.e.r,
                    startColumn: merge.s.c,
                    endColumn: merge.e.c
                });
            }
        }
    }
    return mergedCells;
}
var filePath = process.argv[2];
if (!filePath) {
    console.error('Please provide an Excel file path as a command-line argument');
    process.exit(1);
}
var mergedCells = getMergedCells(filePath);
console.log('Merged Cells:' + '\n');
for (var _i = 0, mergedCells_1 = mergedCells; _i < mergedCells_1.length; _i++) {
    var cell = mergedCells_1[_i];
    console.log("Sheet: ".concat(cell.sheetName, ", ") +
        "Start Row: ".concat(cell.startRow, ", ") +
        "End Row: ".concat(cell.endRow, ", ") +
        "Start Column: ".concat(cell.startColumn, ", ") +
        "End Column: ".concat(cell.endColumn));
    console.log("Range: ".concat(cell.startRow, ",").concat(cell.startColumn, " To ").concat(cell.endRow, ",").concat(cell.endColumn) + '\n');
}
