import * as ExcelJS from 'exceljs';

const CSAT_WORKBOOK_FILENAME = 'vault/CC_Overlay-SRA-CSAT_2018.04.14.xlsx';
const CAP_WORKSHEET_NAME = 'Capabilities - Sec Controls';
const CAPv5_WORKSHEET_NAME = 'Capabilities v5 - Sec Controls';

const COMPARISON_WORKBOOK_FILENAME = 'vault/sp800-53r4-to-r5-comparison-workbook.xlsx';
const COMPARISON_WORKSHEET_NAME = 'Rev4 Rev5 Compared';

const OVERRIDE_WORKBOOK_FILENAME = 'vault/CLOUD OVERLAY CHANGES_3-24-2022.xlsx';
const OVERRIDE_WORKSHEET_NAME = 'Sheet1';

const ADDITIONS_WORKBOOK_FILENAME = 'vault/CSAT_FEDRAMP REV 5 CONTROLS TO BE ADDED_7-6-2023.xlsx';
const ADDITIONS_WORKSHEETS_COLUMN_MAPPING: Record<string, string> = {};
ADDITIONS_WORKSHEETS_COLUMN_MAPPING[`${indexFromColumn('K')}`] = 'LOW';
ADDITIONS_WORKSHEETS_COLUMN_MAPPING[`${indexFromColumn('O')}`] = 'MODERATE';
ADDITIONS_WORKSHEETS_COLUMN_MAPPING[`${indexFromColumn('S')}`] = 'HIGH';

const OUTPUT_WORKBOOK_FILENAME = 'vault/output.xlsx';

const COLOR_GREEN = { argb: 'FF63BE7B' };
const COLOR_ORANGE = { argb: 'FFEDB126' };
const COLOR_RED = { argb: 'FFF8696B' };
const COLOR_PURPLE = { argb: 'FFBF40BF' };
const COLOR_BLACK = { argb: 'FF000000' };
const COLOR_BLUE = { argb: 'FF3A26ED' };
const COLOR_MAGENTA = { argb: 'FFFF00FF' };

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

/**
 * Get a numeric index from an excel column
 * @param column The column name
 * @returns The excel index for a given column (starting at 1)
 */
function indexFromColumn(column: string) {
    return column
        .toUpperCase()
        .split('')
        .reverse()
        .map<number>((letter, index) => (letter.charCodeAt(0) - 64) * 26 ** index)
        .reduce((prev, curr) => prev + curr);
}

function range(length: number, start = 0) {
    return [...Array.from(new Array(length), (_, i) => i + start)];
}

/**
 * Given starting and ending columns and rows, generate a list of cell names e.g. "A5"
 * @param startCol Starting column (e.g. "A")
 * @param startRow Starting row (e.g. 1)
 * @param endCol Ending column (e.g. "B")
 * @param endRow Ending row (e.g. 2)
 * @returns A list of cell names (e.g. ["A1", "A2", "B1", "B2"])
 */
function generateRange(startCol: string, startRow: number, endCol: string, endRow: number): string[] {
    const startColIndex = indexFromColumn(startCol);
    const endColIndex = indexFromColumn(endCol);

    return range(endColIndex - startColIndex, startColIndex - 1).flatMap((curCol) =>
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
                // // only exits on versions manually updated
                // const cloudOverlayOverride = row.getCell('I').text;

                if (rawEditorialSwitch === 'N') {
                    if (rawChangedElements === 'N') {
                        return [rawId, 'unchanged'];
                    } else {
                        return [rawId, 'editorial/administrative'];
                    }
                } else if (rawChangedElements === 'Withdrawn') {
                    return [rawId, 'withdrawn'];
                    // } else if (cloudOverlayOverride === 'K') {
                    //     // if the cloud overlay column exists and the control is set to the "Keep" status
                    //     return [rawId, 'keep'];
                } else {
                    return [rawId, 'changed'];
                }
            }),
    );
}

function populateOverrideMap(worksheet: ExcelJS.Worksheet): Map<string, string> {
    return new Map(
        worksheet
            .getRows(2, worksheet.rowCount - 1)
            ?.map<[string, string]>((row) => [row.getCell('A').text.trim(), row.getCell('B').text.trim()]),
    );
}

async function init() {
    const csat_wb = new ExcelJS.Workbook();
    await csat_wb.xlsx.readFile(CSAT_WORKBOOK_FILENAME);

    const comparison_wb = new ExcelJS.Workbook();
    await comparison_wb.xlsx.readFile(COMPARISON_WORKBOOK_FILENAME);
    const comparison_ws = comparison_wb.getWorksheet(COMPARISON_WORKSHEET_NAME);
    const controlMap = populateControlMap(comparison_ws);

    const override_wb = new ExcelJS.Workbook();
    await override_wb.xlsx.readFile(OVERRIDE_WORKBOOK_FILENAME);
    const override_ws = override_wb.getWorksheet(OVERRIDE_WORKSHEET_NAME);
    const overrideMap = populateOverrideMap(override_ws);

    const capv5_ws = duplicateWorksheet(csat_wb, CAP_WORKSHEET_NAME, CAPv5_WORKSHEET_NAME);

    const addition_wb = new ExcelJS.Workbook();
    await addition_wb.xlsx.readFile(ADDITIONS_WORKBOOK_FILENAME);

    let replacements_count = 0;
    let additions_count = 0;

    // loop through cloud overlay control list, for each control style the control depending on the control map
    generateRange('I', 4, 'X', 349).forEach((cellName) => {
        const cell = capv5_ws.getCell(cellName);

        if (cell.text.trim().length === 0) {
            return;
        }

        cell.value = {
            richText: cell.text
                .split(',')
                .map((control) => control.trim())
                .map((control) => [control, controlMap.get(control) ?? 'unknown', overrideMap.get(control) ?? ''])
                .flatMap<ExcelJS.RichText>(([control, status, replacement]) => {
                    let additions: ExcelJS.RichText[] = [];
                    if (cell.col in ADDITIONS_WORKSHEETS_COLUMN_MAPPING) {
                        const addition_ws = addition_wb.getWorksheet(ADDITIONS_WORKSHEETS_COLUMN_MAPPING[cell.col]);

                        addition_ws
                            .getRows(0, addition_ws.rowCount - 1)
                            ?.filter((row) => row.getCell('A').text.trim() === control)
                            .forEach((row) => {
                                additions_count += 1;
                                additions.push(
                                    {
                                        text: row.getCell('C').text.trim().split(',')[0].trim(),
                                        font: { color: COLOR_MAGENTA },
                                    },
                                    { text: ', ' },
                                );
                            });
                    }
                    if (additions.length != 0) {
                        additions.pop();
                        additions = [
                            { text: ' (', font: { color: COLOR_MAGENTA } },
                            ...additions,
                            { text: ')', font: { color: COLOR_MAGENTA } },
                        ];
                    }

                    if (replacement !== '') {
                        replacements_count += 1;
                        return [
                            { text: control, font: { color: COLOR_BLUE, bold: true, strike: true } },
                            { text: ` => ${replacement}`, font: { color: COLOR_BLUE, bold: true } },
                            ...additions,
                            { text: ', ' },
                        ];
                    } else {
                        return [
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
                                            ? COLOR_ORANGE
                                            : status === 'editorial/administrative'
                                            ? COLOR_BLACK
                                            : COLOR_RED,
                                },
                            },
                            ...additions,
                            { text: ', ' },
                        ];
                    }
                }),
        };
        cell.value.richText.pop(); // remove last comma
    });

    // replace all instances of rev4 with rev5
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

    console.log(`Additions complete!
    Replacements: ${replacements_count}
    Additions: ${additions_count}
    `);
}

init();
