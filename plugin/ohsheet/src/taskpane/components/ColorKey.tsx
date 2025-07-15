import * as React from "react";
import { Text } from "@fluentui/react-components";
import { COLORS } from "../constants/colors";

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
export const ColorKey: React.FC = () => (
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