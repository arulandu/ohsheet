from pydantic import BaseModel
from typing import List
import os
import json
from api.models import Table, SheetCacheData

class ExcelCache:
    def __init__(self, filePath: str):
        self.filePath = filePath
        self.cache = {}
    
    def load(self, f):
        self.cache = json.load(f)
    
    def update(self, sheetId: str, sheetCacheData: SheetCacheData):
        if 'sheets' not in self.cache:
            self.cache['sheets'] = {}

        self.cache['sheets'][sheetId] = sheetCacheData.model_dump()
    
    def get(self, sheetId: str):
        if 'sheets' not in self.cache or sheetId not in self.cache['sheets']:
            return None
        return SheetCacheData(**self.cache['sheets'][sheetId])
    
    def save(self):
        cache_path = get_cache_path(self.filePath)
        with open(cache_path, 'w') as f:
            json.dump(self.cache, f)

def get_cache_path(filePath: str):
    base = filePath.rsplit('.', 1)[0]
    return base + '.osht'

def load_cache(filePath: str):
    cache_path = get_cache_path(filePath)
    cache = ExcelCache(filePath)
    if os.path.exists(cache_path):
        with open(cache_path, 'r') as f:
            cache.load(f)
            return cache
    
    return cache