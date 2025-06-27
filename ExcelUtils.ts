import * as Excel from "exceljs";

type FieldConfig = {
  [key: string]: { range_name: string; initial_cells?: string };
};

let undoStack: any[] = [];

export const fetchFieldConfig = async (): Promise<FieldConfig> => {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItemOrNullObject("AIR_MODEL");
    table.load("name");
    await context.sync();

    if (table.isNullObject) return {};

    const rows = table.rows.load("items");
    await context.sync();

    const config: FieldConfig = {};
    rows.items.forEach((row) => {
      const key = row.values[0][0];
      const rangeName = row.values[0][1];
      const initialCells = row.values[0][2];
      config[key] = { range_name: rangeName, initial_cells: initialCells };
    });

    return config;
  });
};

export const checkForEmptyCells = async (range: Excel.Range, context: Excel.RequestContext): Promise<boolean> => {
  range.load("values");
  await context.sync();

  for (let row of range.values) {
    for (let cell of row) {
      if (cell === "" || cell === null) return true;
    }
  }
  return false;
};

export const addNamedRange = async (context: Excel.RequestContext, name: string, range: Excel.Range) => {
  context.workbook.names.add(name, range);
  await context.sync();
};

export const deleteNamedRangeIfExists = async (context: Excel.RequestContext, name: string) => {
  const existing = context.workbook.names.getItemOrNullObject(name);
  existing.load("name");
  await context.sync();

  if (!existing.isNullObject) {
    existing.delete();
    await context.sync();
  }
};

export const updateAirModelRow = async (key: string, rangeName: string, address: string) => {
  await Excel.run(async (context) => {
    const table = context.workbook.tables.getItem("AIR_MODEL");
    const dataBodyRange = table.getDataBodyRange().load("values");
    await context.sync();

    const rows = dataBodyRange.values;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === key) {
        table.rows.getItemAt(i).values = [[key, rangeName, address]];
        await context.sync();
        break;
      }
    }
  });
};

export const parseFormula = (formula: string): string => {
  if (!formula.startsWith("=")) return `=${formula}`;
  return formula;
};

export const pushUndoState = (state: any) => {
  undoStack.push(JSON.parse(JSON.stringify(state)));
};

export const popUndoState = (): any | null => {
  if (undoStack.length > 0) {
    return undoStack.pop();
  }
  return null;
};
