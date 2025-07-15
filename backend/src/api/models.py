from pydantic import BaseModel
from typing import List, Union, Any, Optional


class FontData(BaseModel):
    """Represents font properties of a cell"""
    name: str
    size: int
    bold: Optional[bool] = False
    # Removed italic and underline as they're not used


class CellData(BaseModel):
    """Represents a single cell in the Excel worksheet"""
    text: str
    value: Union[str, int, float, bool, None]  # Excel cell values can be various types
    numberFormat: str
    formula: str = ""  # Cell formula
    color: Optional[str] = None  # Hex color string (e.g., "#FFFFFF")
    font: Optional[FontData] = None


class SheetData(BaseModel):
    """Represents a single worksheet from Excel"""
    id: str 
    name: str
    usedRange: str  # Excel address like "A1:D10"
    shape: List[int]  # [rowCount, columnCount]
    data: List[List[CellData]]  # 2D array of cells


class ExcelRequest(BaseModel):
    """Request model for the Excel analysis endpoint"""
    filePath: str
    sheets: List[SheetData]
    invalidate: bool = False # cache invalidation
    debug: bool = False # debug mode for plotting

class Table(BaseModel):
    """Represents a detected table in an Excel worksheet"""
    data: str
    row_hdr: str 
    col_hdr: str

class RegionData(BaseModel):
    format: List[dict]  # List of dicts with 'val' and 'ranges' keys
    formula: List[dict]  # List of dicts with 'val' and 'ranges' keys
    color: List[dict]    # List of dicts with 'val' and 'ranges' keys
    text: List[dict]     # List of dicts with 'val' and 'ranges' keys

class RelateRequest(BaseModel):
    """Request model for the relate endpoint"""
    filePath: str
    sheetId: str
    address: str
    precedents: List[str] = []  # Cells this cell depends on
    dependents: List[str] = []  # Cells that depend on this cell
    debug: bool = False # debug mode for plotting

class QueryRequest(BaseModel):
    """Request model for the query endpoint"""
    filePath: str
    sheetId: str
    address: str
    currentCell: str
    currentTable: Optional[Any] = None
    formulaRanges: Optional[dict] = None
    prompt: str

class SheetCacheData(BaseModel):
    id: str
    shape: List[int]
    info_ranges: List[str]
    tables: List[Table]
    regions: RegionData