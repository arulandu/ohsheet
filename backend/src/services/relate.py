from services.cache import ExcelCache, SheetCacheData
from utils.conversion import address_to_coord, coord_to_address

def relate(sheet_cache: SheetCacheData, address: str, prec: list[str], dec: list[str]):
    regions = sheet_cache.regions.formula
    
    def get_formula_range(address: str):
        if ":" in address:
            st, end = address.split(":")
            return [get_formula_range(st), (get_formula_range(end))]
        
        for region in regions:
            for rng in region['ranges']:
                if address_in_range(address, rng):
                    return [rng]
    
        return []
    
    def flatten(lst):
        if not lst:
            return []
        if not isinstance(lst, list):
            return [lst]
        return sum((flatten(x) for x in lst), [])
    
    
    curr = flatten(get_formula_range(address))
    return {
        "current": curr[0] if len(curr) > 0 else "",
        "precedents": flatten([get_formula_range(addr) for addr in prec]),
        "dependents": flatten([get_formula_range(addr) for addr in dec]),
    }

def address_in_range(address: str, rng: str):
   st_c, end_c = [address_to_coord(x) for x in rng.split(":")]
   add = address_to_coord(address)

   return (st_c[0] <= add[0] <= end_c[0] and st_c[1] <= add[1] <= end_c[1])
