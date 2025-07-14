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

class Table(BaseModel):
    """Represents a detected table in an Excel worksheet"""
    data: str
    row_hdr: str 
    col_hdr: str
    
    
class SheetCacheData(BaseModel):
    info_ranges: List[str]
    tables: List[Table]