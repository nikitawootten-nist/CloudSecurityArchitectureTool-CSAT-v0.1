import * as ExcelJS from 'exceljs';

const CSAT_WORKBOOK_FILENAME = 'vault/_CSAT_2020.04.07_clean_modified.xlsx';
const CAP_WORKSHEET_NAME = 'Capabilities - Sec Controls';
const CAPv5_WORKSHEET_NAME = 'Capabilities v5 - Sec Controls';

const COMPARISON_WORKBOOK_FILENAME = 'vault/sp800-53r4-to-r5-comparison-workbook.xlsx';
const COMPARISON_WORKSHEET_NAME = 'Rev4 Rev5 Compared';

const OUTPUT_WORKBOOK_FILENAME = 'vault/output.xlsx';

const COLOR_GREEN = { argb: 'FF63BE7B' };
const COLOR_YELLOW = { argb: 'FFFFC300' };
const COLOR_RED = { argb: 'FFF8696B' };
const COLOR_PURPLE = { argb: 'FFBF40BF' };
const COLOR_BLACK = { argb: 'FF000000' };

function columnFromIndex(index: number) {
    let quotient = index + 1;
    let remainder: number;
    let column = '';

    do {
        remainder = quotient % 26;
        quotient = Math.floor(quotient / 26);

        column += String.fromCharCode(64 + remainder);
    } while (quotient !== 0);

    return column;
}

function indexFromColumn(column: string) {
    return (
        column
            .split('')
            .reverse()
            .map<number>((letter, index) => (letter.charCodeAt(0) - 64) * 26 ** index)
            .reduce((prev, curr) => prev + curr) - 1
    );
}

function range(length: number, start = 0) {
    return [...Array.from(new Array(length), (_, i) => i + start)];
}

function generateRange(startCol: string, startRow: number, endCol: string, endRow: number): string[] {
    const startColIndex = indexFromColumn(startCol);
    const endColIndex = indexFromColumn(endCol);

    return range(endColIndex - startColIndex + 1, startColIndex).flatMap((curCol) =>
        range(endRow - startRow + 1, startRow).map((curRow) => `${columnFromIndex(curCol)}${curRow}`),
    );
}

function duplicateWorksheet(workbook: ExcelJS.Workbook, oldName: string, newName: string): ExcelJS.Worksheet {
    const oldWorksheet = workbook.getWorksheet(oldName);
    const newWorksheet = workbook.addWorksheet(newName);

    newWorksheet.model = Object.assign(oldWorksheet.model, {
        // eslint-disable-next-line @typescript-eslint/ban-ts-comment
        // @ts-ignore
        mergeCells: oldWorksheet.model.merges,
    });
    newWorksheet.name = newName;

    return newWorksheet;
}

function populateControlMap(
    worksheet: ExcelJS.Worksheet,
): Map<string, 'unchanged' | 'editorial/administrative' | 'changed' | 'withdrawn'> {
    return new Map(
        worksheet
            .getRows(3, worksheet.rowCount - 2)
            ?.map<[string, 'unchanged' | 'editorial/administrative' | 'changed' | 'withdrawn']>((row) => {
                const rawId = row.getCell('A').text;
                const rawEditorialSwitch = row.getCell('G').text;
                const rawChangedElements = row.getCell('H').text;

                if (rawEditorialSwitch === 'N') {
                    if (rawChangedElements === 'N') {
                        return [rawId, 'unchanged'];
                    } else {
                        return [rawId, 'editorial/administrative'];
                    }
                } else if (rawChangedElements === 'Withdrawn') {
                    return [rawId, 'withdrawn'];
                } else {
                    return [rawId, 'changed'];
                }
            }),
    );
}

async function init() {
    const csat_wb = new ExcelJS.Workbook();
    await csat_wb.xlsx.readFile(CSAT_WORKBOOK_FILENAME);

    const comparison_wb = new ExcelJS.Workbook();
    await comparison_wb.xlsx.readFile(COMPARISON_WORKBOOK_FILENAME);
    const comparison_ws = comparison_wb.getWorksheet(COMPARISON_WORKSHEET_NAME);
    const controlMap = populateControlMap(comparison_ws);

    const capv5_ws = duplicateWorksheet(csat_wb, CAP_WORKSHEET_NAME, CAPv5_WORKSHEET_NAME);

    generateRange('I', 8, 'X', 353).forEach((cellName) => {
        const cell = capv5_ws.getCell(cellName);

        cell.value = {
            richText: cell.text
                .split(',')
                .map((control) => control.trim())
                .map((control) => [control, controlMap.get(control) ?? 'unknown'])
                .flatMap<ExcelJS.RichText>(([control, status]) => [
                    {
                        text: control,
                        font: {
                            strike: status === 'withdrawn',
                            bold: status === 'changed',
                            italic: status === 'editorial/administrative',
                            underline: status === 'unknown',
                            color:
                                status === 'unchanged'
                                    ? COLOR_GREEN
                                    : status === 'unknown'
                                    ? COLOR_PURPLE
                                    : status === 'changed'
                                    ? COLOR_YELLOW
                                    : status === 'editorial/administrative'
                                    ? COLOR_BLACK
                                    : COLOR_RED,
                        },
                    },
                    { text: ', ' },
                ]),
        };
        cell.value.richText.pop();
    });

    csat_wb.xlsx.writeFile(OUTPUT_WORKBOOK_FILENAME);
}

init();
