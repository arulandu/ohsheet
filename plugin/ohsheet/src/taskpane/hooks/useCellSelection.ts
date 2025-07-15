import { useState, useEffect } from "react";
import { isCellInRange } from "../../utils/range";
import { useTableDrawing } from "./useTableDrawing";
import { useFormulaRelations } from "./useFormulaRelations";
import { COLORS } from "../constants/colors";
import { FormulaRanges, HighlightMode, SheetResult } from "../types";

export const useCellSelection = (sheetResult: SheetResult | null, highlightMode: HighlightMode) => {
  const [currentCell, setCurrentCell] = useState<string>("");
  const [currentTable, setCurrentTable] = useState<any>(null);
  const [formulaRanges, setFormulaRanges] = useState<FormulaRanges>({
    current: "",
    precedents: [],
    dependents: [],
  });

  const { drawTable } = useTableDrawing();
  const { relateCell } = useFormulaRelations();

  const isCellInTable = (cellAddress: string, table: any) => {
    let cellType = "";
    if (isCellInRange(cellAddress, table.data)) {
      cellType = "data";
    } else if (isCellInRange(cellAddress, table.row_hdr)) {
      cellType = "row";
    } else if (isCellInRange(cellAddress, table.col_hdr)) {
      cellType = "column";
    }

    if (cellType) {
      return cellType;
    }

    return cellType;
  };

  const updateActiveTable = async (context: Excel.RequestContext, range: Excel.Range) => {
    if (sheetResult) {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.load("name");
      await context.sync();

      const sheetData = sheetResult.sheets.find((s) => s.name === worksheet.name);
      if (!sheetData) {
        console.log("No data found for current worksheet");
        return;
      }
      let setTable = false;
      for (const table of sheetData.tables) {
        const cellType = isCellInTable(range.address, table);
        if (cellType.length > 0) {
          if (currentTable !== table) {
            if (currentTable) {
              drawTable(currentTable, [
                COLORS.TABLE_INACTIVE,
                COLORS.TABLE_INACTIVE,
                COLORS.TABLE_INACTIVE,
              ]);
            }
          }

          setCurrentTable(table);
          drawTable(table);
          setTable = true;
          break;
        }
      }
      if (!setTable && currentTable) {
        drawTable(currentTable, [
          COLORS.TABLE_INACTIVE,
          COLORS.TABLE_INACTIVE,
          COLORS.TABLE_INACTIVE,
        ]);
        setCurrentTable(null);
      }
    }
  };

  const onCurrentCellChange = async (
    context: Excel.RequestContext,
    worksheet: Excel.Worksheet,
    range: Excel.Range,
    cellAddress: string
  ) => {
    if (highlightMode === "table") {
      await updateActiveTable(context, range);
    }
    if (highlightMode === "formula") {
      await relateCell(context, setFormulaRanges, formulaRanges);
    }
  };

  // Listen to cell selection changes
  useEffect(() => {
    const handleSelectionChanged = async () => {
      try {
        await Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          const worksheet = context.workbook.worksheets.getActiveWorksheet();
          range.load("address");
          worksheet.load("name");
          await context.sync();

          setCurrentCell(range.address);
          await onCurrentCellChange(context, worksheet, range, range.address);
        });
      } catch (error) {
        console.error("Error handling selection change:", error);
      }
    };

    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      handleSelectionChanged,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          console.error("Failed to register selection change handler:", result.error);
        } else {
          console.log("Successfully registered selection change handler");
        }
      }
    );

    return () => {
      Office.context.document.removeHandlerAsync(
        Office.EventType.DocumentSelectionChanged,
        handleSelectionChanged
      );
    };
  }, [
    sheetResult,
    currentTable,
    formulaRanges,
    highlightMode,
  ]); // Re-register when result changes

  return {
    currentCell,
    currentTable,
    formulaRanges,
  };
}; 