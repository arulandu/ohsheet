import { COLORS } from "../constants/colors";

export const useTableDrawing = () => {
  const drawTable = async (
    table: any,
    color: string[] = [COLORS.TABLE_DATA, COLORS.TABLE_ROW_HDR, COLORS.TABLE_COL_HDR]
  ) => {
    if (!table) return;

    try {
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getActiveWorksheet();

        if (table.data && table.data !== "") {
          const dataRange = worksheet.getRange(table.data);
          dataRange.format.fill.color = color[0];
        }

        if (table.row_hdr && table.row_hdr !== "") {
          const rowHdrRange = worksheet.getRange(table.row_hdr);
          rowHdrRange.format.fill.color = color[1];
        }

        if (table.col_hdr && table.col_hdr !== "") {
          const colHdrRange = worksheet.getRange(table.col_hdr);
          colHdrRange.format.fill.color = color[2];
        }

        await context.sync();
      });
    } catch (error) {
      console.error("Error drawing table:", error);
    }
  };

  const showAllTables = async (result: any) => {
    try {
      for (const sheetResult of result.sheets) {
        if (!sheetResult.tables || sheetResult.error) continue;

        sheetResult.tables.forEach((table: any, tableIndex: number) => {
          drawTable(table);
        });
      }
    } catch (error) {
      console.error("Error showing all tables:", error);
    }
  };

  const hideAllTables = async (result: any) => {
    try {
      for (const sheetResult of result.sheets) {
        if (!sheetResult.tables || sheetResult.error) continue;

        sheetResult.tables.forEach((table: any) => {
          drawTable(table, [COLORS.TABLE_INACTIVE, COLORS.TABLE_INACTIVE, COLORS.TABLE_INACTIVE]);
        });
      }
    } catch (error) {
      console.error("Error hiding all tables:", error);
    }
  };

  return {
    drawTable,
    showAllTables,
    hideAllTables,
  };
}; 