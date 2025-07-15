from api.models import SheetCacheData

def query(prompt, cell_address, sheet_cache: SheetCacheData):
    return f"Answer to {prompt} in cell {cell_address} of sheet {sheet_cache.id}"