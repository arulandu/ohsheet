import { API_URL } from "../../util/constants";
import { useFormulaDrawing } from "./useFormulaDrawing";
import { COLORS } from "../constants/colors";
import { loadFilePath } from "../utils/fileUtils";
import { FormulaRanges } from "../types";

export const useFormulaRelations = () => {
  const { drawFormulaRanges } = useFormulaDrawing();

  const relateCell = async (
    context: Excel.RequestContext,
    setFormulaRanges: (ranges: FormulaRanges) => void,
    currentFormulaRanges: FormulaRanges
  ) => {
    try {
      const filePath = await loadFilePath();
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = context.workbook.getActiveCell();
      cell.load("address");
      worksheet.load("id, name");
      await context.sync();

      let prec,
        dec: string[] = [];
      try {
        let precedents = cell.getDirectPrecedents();
        precedents.load("addresses");
        await context.sync();
        prec = precedents.addresses.flatMap((addr) => addr.split(","));
      } catch (e) {}
      try {
        let dependents = cell.getDirectDependents();
        dependents.load("addresses");
        await context.sync();
        dec = dependents.addresses.flatMap((addr) => addr.split(","));
      } catch (e) {}

      const filterSameSheetAddresses = (addresses: string[]) => {
        if (!addresses) return [];
        return addresses
          .map((addr) => {
            if (addr.includes("!")) {
              const [sheetName, address] = addr.split("!");
              const cleanSheetName = sheetName.replace(/['"]/g, "");
              return [cleanSheetName, address];
            }

            return [worksheet.name, addr];
          })
          .filter((addr) => addr[0] === worksheet.name)
          .map((addr) => addr[1]);
      };

      const body = {
        filePath: filePath,
        sheetId: worksheet.id,
        address: cell.address.split("!")[1],
        precedents: prec ? filterSameSheetAddresses(prec) : [],
        dependents: dec ? filterSameSheetAddresses(dec) : [],
        debug: false, // TODO: Pass debug mode as parameter
      };

      const response = await fetch(`${API_URL}/api/relate`, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(body),
      });

      const result = await response.json();
      console.log(result);

      // Update state with new ranges
      const newCurrent = result.current || "";
      const newPrecedents = result.precedents || [];
      const newDependents = result.dependents || [];

      if (
        newCurrent !== currentFormulaRanges.current ||
        newPrecedents !== currentFormulaRanges.precedents ||
        newDependents !== currentFormulaRanges.dependents
      ) {
        console.log("range changed", currentFormulaRanges.current, currentFormulaRanges.precedents, currentFormulaRanges.dependents);
        await drawFormulaRanges(currentFormulaRanges.current, currentFormulaRanges.precedents, currentFormulaRanges.dependents, [
          COLORS.FORMULA_CLEAR,
          COLORS.FORMULA_CLEAR,
          COLORS.FORMULA_CLEAR,
        ]);
        setFormulaRanges({
          current: newCurrent,
          precedents: newPrecedents,
          dependents: newDependents,
        });
        await drawFormulaRanges(newCurrent, newPrecedents, newDependents);
      }
    } catch (error) {
      console.error("Error fetching cell relations:", error);
    }
  };

  return {
    relateCell,
  };
}; 