from fastapi import APIRouter
from services.relate import relate
from services.cache import load_cache
from utils.path import extract_save_dir
from utils.plot import plot_sheet_cache
from .models import RelateRequest

router = APIRouter()

# What do we want? Given prec/dep addresses + req address, return formula range for the cell and the dependents and the precedents. 

@router.post("/relate")
async def relate_post(request: RelateRequest):
    """
    Analyze relationships for a specific cell in a sheet
    """
    sheet_id = request.sheetId
    cell_address = request.address

    cache = load_cache(request.filePath)
    sheet_cache = cache.get(sheet_id)
    if not sheet_cache:
        return {
            "success": False,
            "error": "Sheet not found"
        }
    
    result = relate(sheet_cache, request.address, request.precedents, request.dependents)
    
    return {
        "success": True,
        **result
    } 