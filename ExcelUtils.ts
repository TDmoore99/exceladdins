import { Excel } from "@microsoft/office-js";

export const checkForEmptyCells = async (range: Excel.Range, context: Excel.RequestContext): Promise<boolean> => {
  range.load("values");
  await context.sync();
  return range.values.some((row) => row.some((cell) => cell === null || cell === "" || cell === undefined));
};

export const deleteNamedRangeIfExists = async (context: Excel.RequestContext, rangeName: string) => {
  try {
    const existingName = context.workbook.names.getItemOrNullObject(rangeName);
    await context.sync();
    if (!existingName.isNullObject) {
      existingName.delete();
      await context.sync();
    }
  } catch (error) {
    console.warn(`Error deleting named range "${rangeName}":`, error);
  }
};

export const addNamedRange = async (context: Excel.RequestContext, rangeName: string, range: Excel.Range) => {
  context.workbook.names.add(rangeName, range);
  await context.sync();
};

export const updateMyRangeRow = async (fieldKey: string, rangeName: string, address: string) => {
  await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem("MY_RANGE");
    const dataRange = table.getDataBodyRange();
    dataRange.load("values");
    await context.sync();

    const tableData = dataRange.values;
    const matchIndex = tableData.findIndex((row) => row[0] === fieldKey);

    if (matchIndex !== -1) {
      const targetCell = dataRange.getCell(matchIndex, 2);
      targetCell.values = [[address]];
    } else {
      table.rows.add(null, [[fieldKey, rangeName, address]]);
    }

    await context.sync();
  });
};

export const sanitizeFormulaInput = (input: string): string => {
  return input.replace(/[^\w\d\=\+\-\*\/\(\)\:\,\.\!\$]/g, "");
};

export const parseFormulaFunction = (formula: string): string | null => {
  const match = formula.match(/^=\s*([A-Z]+)\s*\(/i);
  return match ? match[1].toUpperCase() : null;
};
