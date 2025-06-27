import React, { useEffect, useState } from "react";
import * as Excel from "exceljs";
import {
  fetchFieldConfig,
  checkForEmptyCells,
  addNamedRange,
  deleteNamedRangeIfExists,
  updateAirModelRow,
  parseFormula,
  pushUndoState,
  popUndoState,
} from "../services/excelUtils";

type FieldConfig = {
  [key: string]: { range_name: string; initial_cells?: string };
};

const FormulaInputRow: React.FC<{
  fieldKey: string;
  value: string;
  onChange: (newValue: string) => void;
  onPickRange: () => void;
  activeField: string | null;
}> = ({ fieldKey, value, onChange, onPickRange, activeField }) => {
  const handleFunctionInsert = (func: string) => {
    if (func) {
      onChange(`${func}()`);
    }
  };

  return (
    <div className="flex items-center space-x-2 mb-2">
      <label className="w-32">{fieldKey}</label>
      <input
        type="text"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={`Formula for ${fieldKey}`}
        className="border p-1 flex-1"
        disabled={activeField === fieldKey}
      />
      <button onClick={onPickRange} className="border px-2 py-1 bg-blue-100">
        Pick Range
      </button>
      <select
        title="Select a formula function"
        onChange={(e) => handleFunctionInsert(e.target.value)}
        className="border p-1"
      >
        <option value="">Function...</option>
        <option value="SUM">SUM</option>
        <option value="AVERAGE">AVERAGE</option>
        <option value="MAX">MAX</option>
        <option value="MIN">MIN</option>
      </select>
    </div>
  );
};

const FinMappingPane: React.FC = () => {
  const [formulas, setFormulas] = useState<{ [key: string]: string }>({});
  const [fieldConfig, setFieldConfig] = useState<FieldConfig>({});
  const [error, setError] = useState<string | null>(null);
  const [activeField, setActiveField] = useState<string | null>(null);

  useEffect(() => {
    const loadConfig = async () => {
      const config = await fetchFieldConfig();
      setFieldConfig(config);

      const initialFormulas: { [key: string]: string } = {};
      for (const key in config) {
        if (config[key].initial_cells) {
          initialFormulas[key] = config[key].initial_cells;
        }
      }
      setFormulas(initialFormulas);
    };

    loadConfig();

    Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.onSelectionChanged.add(async () => {
        if (!activeField) return;

        const range = context.workbook.getSelectedRange();
        range.load(["address", "values"]);
        await context.sync();

        const hasEmpty = await checkForEmptyCells(range, context);
        if (hasEmpty) {
          setError(`Error: Selected range for "${activeField}" contains empty cells.`);
          setActiveField(null);
          return;
        }

        setError(null);
        const rangeName = fieldConfig[activeField]?.range_name;
        if (rangeName) {
          await deleteNamedRangeIfExists(context, rangeName);
          await addNamedRange(context, rangeName, range);

          pushUndoState(formulas);
          setFormulas((prev) => ({
            ...prev,
            [activeField]: `=${rangeName}`,
          }));

          await updateAirModelRow(activeField, rangeName, range.address);
        }
        setActiveField(null);
      });
    });
  }, [activeField, fieldConfig]);

  const handleInputChange = (key: string, value: string) => {
    setFormulas((prev) => ({
      ...prev,
      [key]: value,
    }));
  };

  const handlePickRange = (key: string) => {
    setActiveField(key);
  };

  const handleUndo = () => {
    const previousState = popUndoState();
    if (previousState) setFormulas(previousState);
  };

  const handleSave = async () => {
    try {
      await Excel.run(async (context) => {
        for (const key in formulas) {
          const formula = parseFormula(formulas[key]);
          console.log(`Saving for ${key}: ${formula}`);
        }
      });
    } catch (err) {
      setError(`Error saving formulas: ${err}`);
    }
  };

  return (
    <div className="p-4">
      <h2 className="text-lg mb-4">Financial Mapping</h2>

      {Object.keys(fieldConfig).map((key) => (
        <FormulaInputRow
          key={key}
          fieldKey={key}
          value={formulas[key] || ""}
          onChange={(val) => handleInputChange(key, val)}
          onPickRange={() => handlePickRange(key)}
          activeField={activeField}
        />
      ))}

      {error && <div className="text-red-600">{error}</div>}

      <div className="flex space-x-2 mt-4">
        <button onClick={handleUndo} className="border px-2 py-1 bg-yellow-100">
          Undo
        </button>
        <button onClick={handleSave} className="border px-2 py-1 bg-green-200">
          Save
        </button>
      </div>
    </div>
  );
};

export default FinMappingPane;
