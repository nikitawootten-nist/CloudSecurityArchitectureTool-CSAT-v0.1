import * as ExcelJS from 'exceljs';

const CSAT_WORKBOOK_FILENAME = 'vault/_CSAT_2020.04.07_clean_modified.xlsx';
const CAP_WORKSHEET_NAME = 'Capabilities - Sec Controls';

// the first row that has any real data on it
const CAP_WORKSHEET_ROW_OFFSET = 8;

// const COMPARISON_WORKBOOK_FILENAME = 'vault/';
// const COMPARISON_WORKSHEET_NAME = '';

const OUTPUT_WORKBOOK_FILENAME = 'vault/output.xlsx';

// function excelCol(index: number) {
//     let column = '';
//     while (index > 26) {
//         column += 'Z';
//         index -= 26;
//     }
//     return column + String.fromCharCode(65 + index);
// }

async function init() {
    const csatWorkbook = new ExcelJS.Workbook();
    await csatWorkbook.xlsx.readFile(CSAT_WORKBOOK_FILENAME);
    const capWorksheet = csatWorkbook.getWorksheet(CAP_WORKSHEET_NAME);

    // const comparisonWorkbook = new ExcelJS.Workbook();
    // await comparisonWorkbook.xlsx.readFile(COMPARISON_WORKBOOK_FILENAME);
    // const comparisonWorksheet = comparisonWorkbook.getWorksheet(COMPARISON_WORKSHEET_NAME);

    capWorksheet.columns.slice(5, 21).forEach((column) => {
        const newRows = column.values?.slice(CAP_WORKSHEET_ROW_OFFSET).map((cell) => [cell, cell]);
        capWorksheet.spliceColumns(21, 0, newRows?.map((row) => row[0]) ?? [], newRows?.map((row) => row[1]) ?? []);
    });

    capWorksheet.getColumn('BJ').eachCell((cell) => {
        if (cell.value && typeof cell.value === 'object') {
            if ('sharedFormula' in cell.value) {
                cell.value.sharedFormula = cell.value.sharedFormula.replace('AD', 'BJ');
            } else if ('formula' in cell.value) {
                cell.value.formula = cell.value.formula.replace('AA', 'BG').replace('AB', 'BH').replace('AC', 'BI');
            }
        }
    });

    csatWorkbook.xlsx.writeFile(OUTPUT_WORKBOOK_FILENAME);
}

init();
