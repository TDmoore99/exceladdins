import React, { useEffect, useState } from "react";
import {
  checkForEmptyCells,
  deleteNamedRangeIfExists,
  addNamedRange,
  updatemyRangeRow,
  sanitizeFormulaInput,
  parseFormulaFunction
} from "../services/excelUtils";
import { Excel } from "@microsoft/office-js";

type FieldConfig = {
  [key: string]: {
    range_name: string;
  };
};

type FormulaInputs = {
  [key: string]: string;
};

const FinMappingPane: React.FC = () => {
  const [fieldConfig, setFieldConfig] = useState<FieldConfig>({});
  const [formulas, setFormulas] = useState<FormulaInputs>({});
  const [undoStack, setUndoStack] = useState<FormulaInputs[]>([]);
  const [redoStack, setRedoStack] = useState<FormulaInputs[]>([]);
  const [activeField, setActiveField] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    loadFieldsFromApi();
    registerExcelEvents();
  }, []);

  const loadFieldsFromApi = async () => {
    const mockApiResponse: FieldConfig = {
      StartDate: { range_name: "StartDateRange" },
      EndDate: { range_name: "EndDateRange" },
      Amount: { range_name: "AmountRange" }
    };

    setFieldConfig(mockApiResponse);
    await prefillFrommyRangeTable(mockApiResponse);
  };

  const prefillFrommyRangeTable = async (config: FieldConfig) => {
    await Excel.run(async (context) => {
      try {
        const table = context.workbook.tables.getItem("MY_RANGE");
        const dataRange = table.getDataBodyRange();
        dataRange.load("values");
        await context.sync();

        const tableData = dataRange.values;
        const newFormulas: FormulaInputs = {};

        Object.keys(config).forEach((key) => {
          const matchRow = tableData.find((row) => row[0] === key);
          if (matchRow && matchRow[2]) {
            newFormulas[key] = `=${matchRow[2]}`;
          }
        });

        setFormulas((prev) => ({
          ...prev,
          ...newFormulas
        }));
      } catch (e) {
        setError("Error reading MY_RANGE table.");
      }
    });
  };

  const pushUndoState = () => {
    setUndoStack((prev) => [...prev, { ...formulas }]);
    setRedoStack([]);
  };

  const undo = () => {
    if (undoStack.length > 0) {
      const last = undoStack.pop()!;
      setRedoStack((prev) => [...prev, { ...formulas }]);
      setFormulas(last);
    }
  };

  const redo = () => {
    if (redoStack.length > 0) {
      const next = redoStack.pop()!;
      setUndoStack((prev) => [...prev, { ...formulas }]);
      setFormulas(next);
    }
  };

  const registerExcelEvents = () => {
    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      sheet.onSelectionChanged.add(async () => {
        if (activeField) {
          const range = context.workbook.getSelectedRange();
          range.load(["address", "values"]);
          await context.sync();

          const hasEmptyCells = await checkForEmptyCells(range, context);
          if (hasEmptyCells) {
            setError(`Error: The selected range for "${activeField}" contains empty cells.`);
            return;
          }

          setError(null);
          const rangeName = fieldConfig[activeField]?.range_name;
          if (rangeName) {
            await deleteNamedRangeIfExists(context, rangeName);
            await addNamedRange(context, rangeName, range);

            pushUndoState();
            setFormulas((prev) => ({
              ...prev,
              [activeField]: `=${rangeName}`
            }));

            await updatemyRangeRow(activeField, rangeName, range.address);
          }
        }
      });

      await context.sync();
    }).catch((e) => setError("Error registering Excel events."));
  };

  const handlePickRange = (fieldKey: string) => {
    setActiveField(fieldKey);
  };

  const handleFormulaChange = (fieldKey: string, value: string) => {
    pushUndoState();
    setFormulas((prev) => ({
      ...prev,
      [fieldKey]: sanitizeFormulaInput(value)
    }));
  };

  const handleSubmit = async () => {
    try {
      await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem("MY_RANGE");
        const dataRange = table.getDataBodyRange();
        dataRange.load("values");
        await context.sync();

        const tableData = dataRange.values;

        for (const key of Object.keys(formulas)) {
          const formulaText = formulas[key].replace(/^=/, "");
          const matchIndex = tableData.findIndex((row) => row[0] === key);

          if (matchIndex !== -1) {
            const targetCell = dataRange.getCell(matchIndex, 2);
            targetCell.values = [[formulaText]];
          } else {
            table.rows.add(null, [[key, fieldConfig[key]?.range_name || "", formulaText]]);
          }
        }

        await context.sync();
      });
    } catch (e) {
      setError("Submit error updating Excel.");
    }
  };

  return (
    <div className="p-4">
      <h2 className="text-xl font-bold mb-2">Financial Field Mapping</h2>
      {error && <div className="text-red-600 mb-2">{error}</div>}

      <div className="flex gap-2 mb-4">
        <button onClick={undo} disabled={undoStack.length === 0} className="bg-gray-300 px-2 py-1 rounded">
          Undo
        </button>
        <button onClick={redo} disabled={redoStack.length === 0} className="bg-gray-300 px-2 py-1 rounded">
          Redo
        </button>
      </div>

      {Object.keys(fieldConfig).map((key) => (
        <FormulaInputRow
          key={key}
          fieldKey={key}
          value={formulas[key] || ""}
          onPickRange={() => handlePickRange(key)}
          onChange={(val) => handleFormulaChange(key, val)}
          detectedFunction={parseFormulaFunction(formulas[key] || "")}
        />
      ))}

      <button onClick={handleSubmit} className="mt-4 px-4 py-2 bg-blue-500 text-white rounded">
        Submit
      </button>
    </div>
  );
};

type FormulaInputRowProps = {
  fieldKey: string;
  value: string;
  onPickRange: () => void;
  onChange: (value: string) => void;
  detectedFunction: string | null;
};

const FormulaInputRow: React.FC<FormulaInputRowProps> = ({
  fieldKey,
  value,
  onPickRange,
  onChange,
  detectedFunction
}) => {
  const handleFunctionInsert = (funcName: string) => {
    if (funcName) {
      onChange(`=${funcName}()`);
    }
  };

  return (
    <div className="mb-4">
      <label className="font-semibold">{fieldKey}</label>
      <div className="flex items-center gap-2 mt-1">
        <input
          type="text"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          placeholder={`Formula for ${fieldKey}`}
          className="border p-1 flex-1"
        />
        <button onClick={onPickRange} className="bg-green-500 text-white px-2 py-1 rounded">
          Pick Range
        </button>
        <select onChange={(e) => handleFunctionInsert(e.target.value)} className="border p-1">
          <option value="">Function...</option>
          <option value="SUM">SUM</option>
          <option value="AVERAGE">AVERAGE</option>
          <option value="MAX">MAX</option>
          <option value="MIN">MIN</option>
        </select>
      </div>
      {detectedFunction && <div className="text-sm text-gray-500">Detected: {detectedFunction}</div>}
    </div>
  );
};

export default FinMappingPane;
