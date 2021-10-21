import * as ExcelJS from 'exceljs';

const CSAT_WORKBOOK_FILENAME = 'vault/_CSAT_2020.04.07_clean_modified.xlsx';
const CAP_WORKSHEET_NAME = 'Capabilities - Sec Controls';
const CAPv5_WORKSHEET_NAME = 'Capabilities v5 - Sec Controls';

const COMPARISON_WORKBOOK_FILENAME = 'vault/sp800-53r4-to-r5-comparison-workbook.xlsx';
const COMPARISON_WORKSHEET_NAME = 'Rev4 Rev5 Compared';

const OUTPUT_WORKBOOK_FILENAME = 'vault/output.xlsx';

const COLOR_GREEN = { argb: 'FF63BE7B' };
const COLOR_ORANGE = { argb: 'FFEDB126' };
const COLOR_RED = { argb: 'FFF8696B' };
const COLOR_PURPLE = { argb: 'FFBF40BF' };
const COLOR_BLACK = { argb: 'FF000000' };
const COLOR_BLUE = { argb: 'FF3A26ED' };

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
): Map<string, 'unchanged' | 'editorial/administrative' | 'changed' | 'withdrawn' | 'keep'> {
    return new Map(
        worksheet
            .getRows(3, worksheet.rowCount - 2)
            ?.map<[string, 'unchanged' | 'editorial/administrative' | 'changed' | 'withdrawn' | 'keep']>((row) => {
                const rawId = row.getCell('A').text;
                const rawEditorialSwitch = row.getCell('G').text;
                const rawChangedElements = row.getCell('H').text;
                // only exits on versions manually updated
                const cloudOverlayOverride = row.getCell('I').text;

                if (rawEditorialSwitch === 'N') {
                    if (rawChangedElements === 'N') {
                        return [rawId, 'unchanged'];
                    } else {
                        return [rawId, 'editorial/administrative'];
                    }
                } else if (rawChangedElements === 'Withdrawn') {
                    return [rawId, 'withdrawn'];
                } else if (cloudOverlayOverride === 'K') {
                    // if the cloud overlay column exists and the control is set to the "Keep" status
                    return [rawId, 'keep'];
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

    // loop through cloud overlay control list, for each control style the control depending on the control map
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
                            italic: status === 'editorial/administrative' || status === 'keep',
                            underline: status === 'unknown',
                            color:
                                status === 'unchanged'
                                    ? COLOR_GREEN
                                    : status === 'unknown'
                                    ? COLOR_PURPLE
                                    : status === 'changed'
                                    ? COLOR_ORANGE
                                    : status === 'editorial/administrative'
                                    ? COLOR_BLACK
                                    : status === 'keep'
                                    ? COLOR_BLUE
                                    : COLOR_RED,
                        },
                    },
                    { text: ', ' },
                ]),
        };
        cell.value.richText.pop(); // remove last comma
    });

    // replace all instancess of rev4 with rev5
    generateRange('I', 5, 'X', 6).forEach((cellName) => {
        const cell = capv5_ws.getCell(cellName);

        if (typeof cell.value === 'string') {
            cell.value = cell.value.replace('R4', 'R5').replace('r4', 'r5').replace('Rev4', 'Rev5');
        } else if (cell.value && typeof cell.value == 'object' && 'richText' in cell.value) {
            cell.value.richText = cell.value.richText.map((richText) => ({
                ...richText,
                text: richText.text.replace('R4', 'R5').replace('r4', 'r5').replace('Rev4', 'Rev5'),
            }));
        }
    });

    csat_wb.xlsx.writeFile(OUTPUT_WORKBOOK_FILENAME);
}

init();
