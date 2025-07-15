from fastapi import APIRouter
from services.cache import load_cache
from services.query import query
from .models import QueryRequest

router = APIRouter()

@router.post("/query")
async def query_post(request: QueryRequest):
    """
    Handle user queries about spreadsheet data
    """
    sheet_id = request.sheetId
    cell_address = request.address
    prompt = request.prompt

    # Load the cached sheet data
    cache = load_cache(request.filePath)
    sheet_cache = cache.get(sheet_id)
    if not sheet_cache:
        return {
            "success": False,
            "error": "Sheet not found"
        }

    msg = query(prompt, cell_address, sheet_cache)
    
    return {
        "success": True,
        "message": msg,
    } 