from fastapi import APIRouter, Query
from fastapi.responses import Response
import json
from .models import ExcelRequest, SheetData, CellData
from services.detect import detect_tables
from services.cache import load_cache, SheetCacheData
import traceback

router = APIRouter()

@router.post("/index")
async def index_post(request: ExcelRequest):
    filePath = request.filePath
    sheets = request.sheets

    cache = load_cache(filePath)

    results = []
    for sheet in sheets:
        try:
            print(f"Processing sheet: {sheet}")

            sheet_cache = cache.get(sheet.id)
            if request.invalidate: sheet_cache = None
            if sheet_cache is None:
                info_ranges, tables = detect_tables(sheet)
                sheet_cache = SheetCacheData(info_ranges=info_ranges, tables=tables)
                cache.update(sheet.id, sheet_cache)
            else:
                print("Hit cache")
                info_ranges = sheet_cache.info_ranges
                tables = sheet_cache.tables
            
            sheet_result = {
                'name': sheet.name,
                'usedRange': sheet.usedRange,
                'shape': sheet.shape,
                'tables': tables,
            }
            results.append(sheet_result)
            
            print(f"Found {len(tables)} tables in sheet {sheet.name}")
            
        except Exception as e:
            print(f"Error processing sheet {sheet.name}: {e}")
            traceback.print_exc()
            # Add error information to results
            sheet_result = {
                'name': sheet.name,
                'usedRange': sheet.usedRange,
                'shape': sheet.shape,
                'tables': [],
                'error': str(e)
            }
            results.append(sheet_result)
    
    cache.save()
    
    return {
        "filePath": filePath,
        "sheets": results,
    }

@router.get("/")
async def root():
    return {"message": "Root route", "status": "active"} 
