import { addressToCoord } from "./conversion";

export const isCellInRange = (cellAddress: string, rangeAddress: string): boolean => {
    if (!rangeAddress || rangeAddress === "") return false;
    
    try {
      // Remove $ signs from addresses
      const cell = cellAddress.replace(/\$/g, '');
      const range = rangeAddress.replace(/\$/g, '');
      
      // Parse the cell coordinates
      const [cellRow, cellCol] = addressToCoord(cell);
      
      // Parse the range (e.g., "A1:B5" -> start and end coordinates)
      const rangeParts = range.split(':');
      if (rangeParts.length === 1) {
        // Single cell range
        const [rangeRow, rangeCol] = addressToCoord(rangeParts[0]);
        return cellRow === rangeRow && cellCol === rangeCol;
      } else if (rangeParts.length === 2) {
        // Multi-cell range
        const [startRow, startCol] = addressToCoord(rangeParts[0]);
        const [endRow, endCol] = addressToCoord(rangeParts[1]);
        
        // Check if cell is within the range bounds
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);
        
        return cellRow >= minRow && cellRow <= maxRow && 
               cellCol >= minCol && cellCol <= maxCol;
      }
      
      return false;
    } catch (error) {
      console.error("Error checking cell range:", error);
      return false;
    }
  };