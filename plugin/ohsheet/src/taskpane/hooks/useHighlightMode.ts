import { useTableDrawing } from "./useTableDrawing";
import { useFormulaDrawing } from "./useFormulaDrawing";
import { COLORS } from "../constants/colors";
import { FormulaRanges, HighlightMode } from "../types";

export const useHighlightMode = (
  sheetResult: any,
  formulaRanges: FormulaRanges,
  setHighlightMode: (mode: HighlightMode) => void
) => {
  const { hideAllTables } = useTableDrawing();
  const { drawFormulaRanges } = useFormulaDrawing();

  const handleHighlightModeChange = async (e: any, data: any) => {
    const newMode = data.value as HighlightMode;
    
    // Clear current highlights based on what mode we're switching from
    if (newMode === "formula") {
      // Switching to formula mode, clear table highlights
      await hideAllTables(sheetResult);
    } else if (newMode === "table") {
      // Switching to table mode, clear formula highlights
      await drawFormulaRanges(formulaRanges.current, formulaRanges.precedents, formulaRanges.dependents, [
        COLORS.FORMULA_CLEAR,
        COLORS.FORMULA_CLEAR,
        COLORS.FORMULA_CLEAR,
      ]);
    }
    
    setHighlightMode(newMode);
  };

  return {
    handleHighlightModeChange,
  };
}; 