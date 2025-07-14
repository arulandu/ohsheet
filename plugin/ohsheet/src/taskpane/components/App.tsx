import * as React from "react";
import {
  makeStyles,
  Button,
  Text,
  Title3,
  Select,
  Label,
  Checkbox,
  Spinner,
} from "@fluentui/react-components";
import { useState } from "react";
import { API_URL } from "../../util/constants";
import { addressToCoord } from "../../utils/conversion";
import { isCellInRange } from "../../utils/range";

const COLORS = {
  TABLE_DATA: "#96CEB4",
  TABLE_ROW_HDR: "#6DAF8A",
  TABLE_COL_HDR: "#6DAF8A",
  TABLE_INACTIVE: "#FFFFFF",

  FORMULA_CURRENT: "#FF6B6B", // Red for current cell
  FORMULA_PRECEDENT: "#B366FF", // Purple for precedents 
  FORMULA_DEPENDENT: "#FFB366", // Orange for dependents
  FORMULA_CLEAR: "#FFFFFF", // White to clear highlights
} as const;

// Color Swatch Component
const ColorSwatch: React.FC<{ color: string; label: string }> = ({ color, label }) => (
  <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px" }}>
    <div
      style={{
        width: "16px",
        height: "16px",
        backgroundColor: color,
        border: "1px solid #ccc",
        borderRadius: "2px",
      }}
    />
    <Text style={{ fontSize: "12px"}}>{label}</Text>
  </div>
);

// Color Key Component
const ColorKey: React.FC = () => (
  <div style={{ marginTop: "10px", padding: "8px", border: "1px solid #e1e1e1", borderRadius: "4px"}}>
    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "8px" }}>
      <div>
        <Text style={{ fontSize: "12px", fontWeight: "bold", marginBottom: "4px"}}>Table</Text>
        <ColorSwatch color={COLORS.TABLE_DATA} label="Data" />
        <ColorSwatch color={COLORS.TABLE_ROW_HDR} label="Row Header" />
        <ColorSwatch color={COLORS.TABLE_COL_HDR} label="Column Header" />
      </div>
      <div>
        <Text style={{ fontSize: "12px", fontWeight: "bold", marginBottom: "4px" }}>Formula</Text>
        <ColorSwatch color={COLORS.FORMULA_CURRENT} label="Current" />
        <ColorSwatch color={COLORS.FORMULA_PRECEDENT} label="In Current Deps." />
        <ColorSwatch color={COLORS.FORMULA_DEPENDENT} label="Current in Deps." />
      </div>
    </div>
  </div>
);

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
    marginBottom: "10px",
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
  const [message, setMessage] = useState("");
  const [sheetSelection, setSheetSelection] = useState("current");
  const [invalidateCache, setInvalidateCache] = useState(false);
  const [debugMode, setDebugMode] = useState(false);
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [sheetResult, setSheetResult] = useState(null);
  const [sheetsData, setSheetsData] = useState(null);
  const [currentCell, setCurrentCell] = useState<string>("");
  const [currentTable, setCurrentTable] = useState<any>(null);
  const [currentFormulaRange, setCurrentFormulaRange] = useState<string>("");
  const [precedentRanges, setPrecedentRanges] = useState<string[]>([]);
  const [dependentRanges, setDependentRanges] = useState<string[]>([]);
  const [highlightMode, setHighlightMode] = useState<"formula" | "table">("formula");

  // Listen to cell selection changes
  React.useEffect(() => {
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
    currentFormulaRange,
    precedentRanges,
    dependentRanges,
    highlightMode,
  ]); // Re-register when result changes

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
      await relateCell(context);
    }
  };

  const handleAnalyze = async () => {
    setIsAnalyzing(true);
    setMessage("Analyzing...");

    await drawFormulaRanges(currentFormulaRange, precedentRanges, dependentRanges, [
      COLORS.FORMULA_CLEAR,
      COLORS.FORMULA_CLEAR,
      COLORS.FORMULA_CLEAR,
    ]);

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

          const result = await response.json();
          setSheetResult(result);

          const totalTables = result.sheets.reduce((total: number, sheet: any) => {
            return total + (sheet.tables ? sheet.tables.length : 0);
          }, 0);

          setMessage(`Found ${totalTables} table(s) across ${result.sheets.length} sheet(s)`);

          if (highlightMode === "table") {
            await showAllTables(result);
            setTimeout(async () => {
              await hideAllTables(result);
            }, 5000);
          }
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

  const relateCell = async (context: Excel.RequestContext) => {
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
        debug: debugMode,
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
        newCurrent !== currentFormulaRange ||
        newPrecedents !== precedentRanges ||
        newDependents !== dependentRanges
      ) {
        console.log("range changed", currentFormulaRange, precedentRanges, dependentRanges);
        await drawFormulaRanges(currentFormulaRange, precedentRanges, dependentRanges, [
          COLORS.FORMULA_CLEAR,
          COLORS.FORMULA_CLEAR,
          COLORS.FORMULA_CLEAR,
        ]);
        setCurrentFormulaRange(newCurrent);
        setPrecedentRanges(newPrecedents);
        setDependentRanges(newDependents);
        await drawFormulaRanges(newCurrent, newPrecedents, newDependents);
      }
    } catch (error) {
      console.error("Error fetching cell relations:", error);
    }
  };

  return (
    <div className={styles.root}>
      <Title3 className={styles.title}>Oh Sheet!</Title3>
      <Text className={styles.subtitle}>AI Spreadsheet Analyzer</Text>

      <div style={{ marginBottom: "10px", width: "100%", maxWidth: "300px" }}>
        <Label htmlFor="sheet-selection">Sheet Selection:</Label>
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

      <div
        style={{
          marginBottom: "20px",
          width: "100%",
          maxWidth: "300px",
          display: "flex",
          gap: "10px",
          justifyContent: "center",
        }}
      >
        <Checkbox
          id="invalidate-cache"
          checked={invalidateCache}
          onChange={(e, data) => setInvalidateCache(data.checked === true)}
          label="Invalidate"
        />
        <Checkbox
          id="debug-mode"
          checked={debugMode}
          onChange={(e, data) => setDebugMode(data.checked === true)}
          label="Debug"
        />
      </div>

      <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
        <Button
          appearance="primary"
          size="large"
          className={styles.analyzeButton}
          onClick={handleAnalyze}
          disabled={isAnalyzing}
        >
          {isAnalyzing ? (
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
              <Spinner size="tiny" />
              Analyzing...
            </div>
          ) : (
            "Analyze"
          )}
        </Button>
      </div>
      <Text className={styles.message}>{message}</Text>

      {sheetResult && (
        <div style={{ marginBottom: "10px", width: "100%", maxWidth: "300px" }}>
          <Label htmlFor="highlight-mode">Highlight Mode:</Label>
          <Select
            id="highlight-mode"
            value={highlightMode}
            onChange={async (e, data) => {
              const newMode = data.value as "formula" | "table";
              if (newMode !== highlightMode) {
                if (highlightMode === "table") {
                  await hideAllTables(sheetResult);
                }
                if (highlightMode === "formula") {
                  await drawFormulaRanges(currentFormulaRange, precedentRanges, dependentRanges, [
                    COLORS.FORMULA_CLEAR,
                    COLORS.FORMULA_CLEAR,
                    COLORS.FORMULA_CLEAR,
                  ]);
                  setCurrentFormulaRange("");
                  setPrecedentRanges([]);
                  setDependentRanges([]);
                }
                setHighlightMode(newMode);
              }
            }}
            style={{ marginTop: "8px" }}
          >
            <option value="formula">Formula</option>
                         <option value="table">Table</option>
           </Select>
           <ColorKey />
         </div>
       )}
    </div>
  );
};

export default App;
