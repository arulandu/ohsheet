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
import { useSheetAnalysis } from "../hooks/useSheetAnalysis";
import { useCellSelection } from "../hooks/useCellSelection";
import { useHighlightMode } from "../hooks/useHighlightMode";
import { useQuery } from "../hooks/useQuery";
import { ColorKey } from "./ColorKey";
import { COLORS } from "../constants/colors";
import { HighlightMode } from "../types";

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

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();
  const [sheetSelection, setSheetSelection] = useState("current");
  const [invalidateCache, setInvalidateCache] = useState(false);
  const [debugMode, setDebugMode] = useState(false);

  const {
    message,
    isAnalyzing,
    sheetResult,
    handleAnalyze,
  } = useSheetAnalysis(sheetSelection, invalidateCache, debugMode);

  const [highlightMode, setHighlightMode] = useState<HighlightMode>("table");

  const {
    currentCell,
    currentTable,
    formulaRanges,
  } = useCellSelection(sheetResult, highlightMode);

  const {
    handleHighlightModeChange,
  } = useHighlightMode(sheetResult, formulaRanges, setHighlightMode);

  const {
    isQuerying,
    queryMessage,
    sendQuery,
  } = useQuery();

  const [queryPrompt, setQueryPrompt] = useState("");

  const handleSendQuery = () => {
    sendQuery(queryPrompt, currentCell, currentTable, formulaRanges, sheetResult);
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
            onChange={handleHighlightModeChange}
            style={{ marginTop: "8px" }}
          >
            <option value="formula">Formula</option>
            <option value="table">Table</option>
          </Select>
          <ColorKey />
        </div>
      )}

      {sheetResult && (
        <div style={{ marginTop: "20px", width: "100%", maxWidth: "300px" }}>
          <Text style={{ fontWeight: "bold", marginBottom: "10px" }}>Query</Text>
          <div style={{ display: "flex", gap: "8px", marginBottom: "8px" }}>
            <input
              type="text"
              value={queryPrompt}
              onChange={(e) => setQueryPrompt(e.target.value)}
              placeholder="What is going on in this cell?"
              style={{
                flex: 1,
                padding: "8px 12px",
                border: "1px solid #d1d1d1",
                borderRadius: "4px",
                fontSize: "14px",
              }}
              onKeyPress={(e) => {
                if (e.key === "Enter") {
                  handleSendQuery();
                }
              }}
            />
            <Button
              appearance="outline"
              size="small"
              onClick={handleSendQuery}
              disabled={isQuerying || !queryPrompt.trim()}
              style={{
                backgroundColor: "#1ea363",
                borderColor: "#1ea363",
              }}
            >
              {isQuerying ? "Sending..." : "Send"}
            </Button>
          </div>
          {queryMessage && (
            <Text style={{ fontSize: "12px", color: "#605e5c", wordBreak: "break-word" }}>
              {queryMessage}
            </Text>
          )}
        </div>
      )}
    </div>
  );
};

export default App;
