import { useState } from "react";
import { API_URL } from "../../util/constants";
import { loadFilePath } from "../utils/fileUtils";
import { FormulaRanges } from "../types";

export const useQuery = () => {
  const [isQuerying, setIsQuerying] = useState(false);
  const [queryMessage, setQueryMessage] = useState("");

  const sendQuery = async (
    prompt: string,
    currentCell: string,
    currentTable: any,
    formulaRanges: FormulaRanges,
    sheetResult: any
  ) => {
    if (!prompt.trim()) {
      setQueryMessage("Please enter a query");
      return;
    }

    setIsQuerying(true);
    setQueryMessage("Sending query...");

    try {
      await Excel.run(async (context) => {
        const filePath = await loadFilePath();
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        worksheet.load("id, name");
        await context.sync();

        // Prepare the body data similar to what's sent in /api/relate
        const body = {
          filePath: filePath,
          sheetId: worksheet.id,
          address: currentCell.includes("!") ? currentCell.split("!")[1] : currentCell,
          currentCell: currentCell,
          currentTable: currentTable,
          formulaRanges: formulaRanges,
          prompt: prompt,
        };

        const response = await fetch(`${API_URL}/api/query`, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify(body),
        });

        if (!response.ok) {
          throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        setQueryMessage(result.message || "Query completed successfully");
      });
    } catch (error) {
      console.error("Error sending query:", error);
      setQueryMessage("Error sending query: " + error.message);
    } finally {
      setIsQuerying(false);
    }
  };

  return {
    isQuerying,
    queryMessage,
    sendQuery,
  };
}; 