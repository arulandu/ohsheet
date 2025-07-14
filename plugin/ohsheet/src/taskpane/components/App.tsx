import * as React from "react";
import { makeStyles, Button, Text, Title3, Select, Label, Checkbox } from "@fluentui/react-components";
import { useState } from "react";
import { API_URL } from "../../util/constants";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    padding: "20px",
    textAlign: "center",
  },
  title: {
    fontSize: "2.5rem",
    fontWeight: "bold",
    marginBottom: "8px",
    color: "#1ea363",
  },
  subtitle: {
    fontSize: "1.2rem",
    marginBottom: "40px",
    color: "#605e5c",
  },
  analyzeButton: {
    fontSize: "1.1rem",
    padding: "12px 32px",
    borderRadius: "6px",
    backgroundColor: "#1ea363",
  },
  message: {
    fontSize: "1.2rem",
    marginTop: "20px",
    color: "#605e5c",
    wordBreak: "break-word",
    whiteSpace: "pre-wrap",
    overflowY: "auto",
    maxWidth: "100%",
  },
});

async function loadFilePath() {
  return new Promise((resolve) => {
    Office.context.document.getFilePropertiesAsync(null, (res) => {
      if (res && res.value && res.value.url) {
        let name = res.value.url.substr(res.value.url.lastIndexOf("\\") + 1);
        resolve(name);
      }
      resolve("");
    });
  });
}

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const [message, setMessage] = useState("Hello, world!");
  const [sheetSelection, setSheetSelection] = useState("current");
  const [invalidateCache, setInvalidateCache] = useState(false);

  const drawTableBoundingBoxes = async (result: any) => {
    try {
      await Excel.run(async (context) => {
        // Define colors for different table components
        const colors = [
          { data: "#FF6B6B", rowHdr: "#4ECDC4", colHdr: "#45B7D1" }, // Red, Teal, Blue
          { data: "#96CEB4", rowHdr: "#FFEAA7", colHdr: "#DDA0DD" }, // Green, Yellow, Purple
          { data: "#FFB347", rowHdr: "#87CEEB", colHdr: "#98FB98" }, // Orange, Sky Blue, Pale Green
          { data: "#DDA0DD", rowHdr: "#F0E68C", colHdr: "#FFB6C1" }, // Plum, Khaki, Light Pink
        ];

        for (const sheetResult of result.sheets) {
          if (!sheetResult.tables || sheetResult.error) continue;
          
          // Get the worksheet by name
          const worksheet = context.workbook.worksheets.getItem(sheetResult.name);
          
          sheetResult.tables.forEach((table: any, tableIndex: number) => {
            const colorSet = colors[tableIndex % colors.length];
            
            // Draw bounding box around data range
            if (table.data && table.data !== '') {
              const dataRange = worksheet.getRange(table.data);
              dataRange.format.fill.color = colorSet.data;
              dataRange.format.borders.getItem('EdgeTop').style = 'Continuous';
              dataRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
              dataRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
              dataRange.format.borders.getItem('EdgeRight').style = 'Continuous';
            }
            
            // Draw bounding box around row headers
            if (table.row_hdr && table.row_hdr !== '') {
              const rowHdrRange = worksheet.getRange(table.row_hdr);
              rowHdrRange.format.fill.color = colorSet.rowHdr;
              rowHdrRange.format.borders.getItem('EdgeTop').style = 'Continuous';
              rowHdrRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
              rowHdrRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
              rowHdrRange.format.borders.getItem('EdgeRight').style = 'Continuous';
            }
            
            // Draw bounding box around column headers
            if (table.col_hdr && table.col_hdr !== '') {
              const colHdrRange = worksheet.getRange(table.col_hdr);
              colHdrRange.format.fill.color = colorSet.colHdr;
              colHdrRange.format.borders.getItem('EdgeTop').style = 'Continuous';
              colHdrRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
              colHdrRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
              colHdrRange.format.borders.getItem('EdgeRight').style = 'Continuous';
            }
            
            // Add a label above the table if there's a data range
            if (table.data && table.data !== '') {
              try {
                // Try to add a label in the cell above the data range
                const dataRange = worksheet.getRange(table.data);
                const topLeftCell = dataRange.getCell(0, 0);
                const labelCell = topLeftCell.getOffsetRange(-1, 0);
                labelCell.values = [[`Table ${tableIndex + 1}`]];
                labelCell.format.font.bold = true;
                labelCell.format.font.color = colorSet.data;
              } catch (e) {
                // If we can't add a label (e.g., at the top of the sheet), just continue
                console.log("Could not add table label:", e);
              }
            }
          });
        }
        
        await context.sync();
        console.log("Drew bounding boxes for detected tables");
      });
    } catch (error) {
      console.error("Error drawing bounding boxes:", error);
    }
  };

  const clearHighlighting = async () => {
    console.log("Clearing highlighting");
  };

  const handleAnalyze = async () => {
    try {
      await Excel.run(async (context) => {
        const filePath = await loadFilePath();

        const worksheets = context.workbook.worksheets;
        worksheets.load("items");
        await context.sync();

        console.log("loaded items")
        
        // Select worksheets based on user preference
        let selectedWorksheets;
        if (sheetSelection === "current") {
          // Get the active worksheet
          const activeWorksheet = context.workbook.worksheets.getActiveWorksheet();
          activeWorksheet.load("name, id");
          await context.sync();
          selectedWorksheets = [activeWorksheet];
        } else {
          // Get all worksheets
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

        console.log("Loaded worksheet base data")
        let cellData = worksheetData.map(({ worksheet, usedRange }) => {
          const props = usedRange.getDisplayedCellProperties({
            address: true,
            style: true,
            format: {
              fill: {
                color: true
              },
              font: {
                name: true,
                size: true,
                bold: true,
              }
            }
          })

          return props
        });
        
        await context.sync();

        console.log("loaded all data")

        const sheets = worksheetData.map(({ worksheet, usedRange }: { worksheet: Excel.Worksheet, usedRange: Excel.Range }, index) => {
          const data = usedRange.values.map((row, r) =>
            row.map((col, c) => {
              const prop : Excel.CellProperties = cellData[index].value[r][c];
              
              return {
                text: usedRange.text[r][c],
                value: usedRange.values[r][c],
                numberFormat: usedRange.numberFormat[r][c],
                formula: String(usedRange.formulas[r][c]),
                color: prop.format.fill.color,
                font: {
                  name: prop.format.font.name,
                  size: prop.format.font.size,
                  bold: prop.format.font.bold
                }
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
        });
console.log("done")
        const body = {
          filePath: filePath,
          sheets: sheets, // Send all selected sheets
          invalidate: invalidateCache,
        }

        try {
          const response = await fetch(`${API_URL}/api/index`, {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
            },
            body: JSON.stringify(body),
          });

          const result = await response.json();
          
          // Count total tables detected
          const totalTables = result.sheets.reduce((total: number, sheet: any) => {
            return total + (sheet.tables ? sheet.tables.length : 0);
          }, 0);
          
          setMessage(`Analysis complete! Found ${totalTables} tables across ${result.sheets.length} sheets. Check your spreadsheet for highlighted table regions.`);
          
          // Draw bounding boxes for detected tables
          await drawTableBoundingBoxes(result);
        } catch (apiError) {
          console.error("API Error:", apiError);
          setMessage("Error calling API: " + apiError.message);
        }
      });
    } catch (error) {
      console.error("Error analyzing worksheet:", error);
      setMessage("Error analyzing worksheet: " + error.message);
    }
  };

  return (
    <div className={styles.root}>
      <Title3 className={styles.title}>Oh Sheet!</Title3>
      <Text className={styles.subtitle}>AI Spreadsheet Analyzer</Text>
      
      <div style={{ marginBottom: "20px", width: "100%", maxWidth: "300px" }}>
        {/* <Label htmlFor="sheet-selection">Sheet Selection:</Label> */}
        <Select
          id="sheet-selection"
          value={sheetSelection}
          onChange={(e, data) => setSheetSelection(data.value)}
          style={{ marginTop: "8px" }}
        >
          <option value="current">Current Sheet</option>
          <option value="all">All Sheets</option>
        </Select>
      </div>
      
      <div style={{ marginBottom: "20px", width: "100%", maxWidth: "300px" }}>
        <Checkbox
          id="invalidate-cache"
          checked={invalidateCache}
          onChange={(e, data) => setInvalidateCache(data.checked === true)}
          label="Invalidate Cache"
        />
      </div>
      
      <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
        <Button
          appearance="primary"
          size="large"
          className={styles.analyzeButton}
          onClick={handleAnalyze}
        >
          Analyze
        </Button>
        <Button
          appearance="secondary"
          size="large"
          onClick={clearHighlighting}
        >
          Clear
        </Button>
      </div>
      <Text className={styles.message}>{message}</Text>
    </div>
  );
};

export default App;
