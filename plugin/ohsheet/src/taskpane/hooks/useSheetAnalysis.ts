import { useState } from "react";
import { API_URL } from "../../util/constants";
import { COLORS } from "../constants/colors";
import { useTableDrawing } from "./useTableDrawing";
import { useFormulaDrawing } from "./useFormulaDrawing";
import { loadFilePath } from "../utils/fileUtils";

export const useSheetAnalysis = (
  sheetSelection: string,
  invalidateCache: boolean,
  debugMode: boolean
) => {
  const [message, setMessage] = useState("");
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [sheetResult, setSheetResult] = useState(null);
  
  const { showAllTables, hideAllTables } = useTableDrawing();
  const { drawFormulaRanges } = useFormulaDrawing();

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setMessage("Analyzing...");

    try {
      await Excel.run(async (context) => {
        const filePath = await loadFilePath();
        const worksheets = context.workbook.worksheets;
        worksheets.load("items");
        await context.sync();

        let selectedWorksheets;
        if (sheetSelection === "current") {
          const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
          activeWorksheet.load("name, id");
          await context.sync();
          selectedWorksheets = [activeWorksheet];
        } else {
          selectedWorksheets = worksheets.items;
        }

        const worksheetPromises = selectedWorksheets.map(async (worksheet: Excel.Worksheet) => {
          const usedRange = worksheet.getUsedRange();
          worksheet.load("name, id");
          usedRange.load("address, values, formulas, numberFormat, text, rowCount, columnCount");
          return { worksheet, usedRange };
        });

        // load all worksheet data
        const worksheetData = await Promise.all(worksheetPromises);
        await context.sync();

        console.log("Loaded worksheet base data");
        let cellData = worksheetData.map(({ worksheet, usedRange }) => {
          const props = usedRange.getDisplayedCellProperties({
            address: true,
            style: true,
            format: {
              fill: {
                color: true,
              },
              font: {
                name: true,
                size: true,
                bold: true,
              },
            },
          });

          return props;
        });

        await context.sync();

        console.log("loaded all data");

        const sheets = worksheetData.map(
          (
            { worksheet, usedRange }: { worksheet: Excel.Worksheet; usedRange: Excel.Range },
            index
          ) => {
            const data = usedRange.values.map((row, r) =>
              row.map((col, c) => {
                const prop: Excel.CellProperties = cellData[index].value[r][c];

                return {
                  text: usedRange.text[r][c],
                  value: usedRange.values[r][c],
                  numberFormat: usedRange.numberFormat[r][c],
                  formula: String(usedRange.formulas[r][c]),
                  color: prop.format.fill.color,
                  font: {
                    name: prop.format.font.name,
                    size: prop.format.font.size,
                    bold: prop.format.font.bold,
                  },
                };
              })
            );

            const sheet = {
              id: worksheet.id,
              name: worksheet.name,
              usedRange: usedRange.address,
              shape: [usedRange.rowCount, usedRange.columnCount],
              data: data,
            };

            return sheet;
          }
        );

        const body = {
          filePath: filePath,
          sheets: sheets, // Send all selected sheets
          invalidate: invalidateCache,
          debug: debugMode,
        };

        try {
          const response = await fetch(`${API_URL}/api/index`, {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify(body),
          });

          if (!response.ok) {
            const errorText = await response.text();
            console.error("API Error Response:", errorText);
            setMessage(`Error calling API: ${response.status} ${response.statusText}`);
            return;
          }

          const result = await response.json();
          
          // Check if the response has the expected structure
          if (!result || !result.sheets) {
            console.error("Unexpected API response structure:", result);
            setMessage("Error: Unexpected response from server");
            return;
          }
          
          setSheetResult(result);

          const totalTables = result.sheets.reduce((total: number, sheet: any) => {
            return total + (sheet.tables ? sheet.tables.length : 0);
          }, 0);

          setMessage(`Found ${totalTables} table(s) across ${result.sheets.length} sheet(s)`);

          // Show tables briefly then hide them
          await showAllTables(result);
          setTimeout(async () => {
            await hideAllTables(result);
          }, 5000);
        } catch (apiError) {
          console.error("API Error:", apiError);
          setMessage("Error calling API: " + apiError.message);
        }
      });
    } catch (error) {
      console.error("Error analyzing worksheet:", error);
      setMessage("Error analyzing worksheet: " + error.message);
    } finally {
      setIsAnalyzing(false);
    }
  };

  return {
    message,
    isAnalyzing,
    sheetResult,
    handleAnalyze,
  };
}; 