export interface FormulaRanges {
  current: string;
  precedents: string[];
  dependents: string[];
}

export type HighlightMode = "formula" | "table";

export interface SheetResult {
  sheets: Array<{
    name: string;
    tables?: Array<{
      data: string;
      row_hdr: string;
      col_hdr: string;
    }>;
    error?: string;
  }>;
} 