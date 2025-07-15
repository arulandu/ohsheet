import { COLORS } from "../constants/colors";

export const useFormulaDrawing = () => {
  const drawFormulaRange = async (range: string, color: string) => {
    if (!range) return;

    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        const rangeObj = worksheet.getRange(range);
        rangeObj.format.fill.color = color;
        await context.sync();
      });
    } catch (error) {
      console.error("Error drawing formula range:", error);
    }
  };

  // Draws all formula ranges with their respective colors
  const drawFormulaRanges = async (
    current: string,
    precedents: string[],
    dependents: string[],
    colors: string[] = [COLORS.FORMULA_CURRENT, COLORS.FORMULA_PRECEDENT, COLORS.FORMULA_DEPENDENT]
  ) => {
    console.log("drawing formula ranges", current, precedents, dependents);
    if (current) {
      await drawFormulaRange(current, colors[0]); // Red for current cell
    }
    for (const range of precedents || []) {
      await drawFormulaRange(range, colors[1]); // Green for precedents
    }
    for (const range of dependents || []) {
      await drawFormulaRange(range, colors[2]); // Orange for dependents
    }
  };

  return {
    drawFormulaRange,
    drawFormulaRanges,
  };
}; 