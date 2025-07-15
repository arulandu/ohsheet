import numpy as np
from utils.conversion import address_to_coord, coord_to_address
from api.models import CellData, SheetData
from typing import List, Callable, Any

def merge_mask(cells: List[List[CellData]], key: Callable[[CellData], Any] = lambda x: x.value, cmp: Callable[[Any, Any], bool] = lambda x, y: x == y):
    """Group via RD DFS on cmp(key1, key2) = true"""
    st = []
    mask = [[-1 for c in r] for r in cells]

    def push(rt, nb):
        if cmp(key(cells[rt[0]][rt[1]]), key(cells[nb[0]][nb[1]])):
            if mask[nb[0]][nb[1]] == -1: # not visited
                st.append(nb)

    def dfs(rt, cat):
        if mask[rt[0]][rt[1]] >= 0: return # alr. visited

        st.append(rt)

        while len(st) > 0:
            s = st.pop()
            r, c = s
            mask[r][c] = cat

            if r < len(cells)-1:
                push(s, (r+1, c))
            
            if c < len(cells[0])-1:
                push(s, (r, c+1))

    cat = 1
    for r, row in enumerate(cells):
        for c, cell in enumerate(row):
            if mask[r][c] >= 0: continue

            if key(cell) is not None:
                dfs((r, c), cat)
                cat += 1
            else:
                mask[r][c] = 0 # empty

    return np.array(mask)

def region_to_ranges(region: np.array):
    def get_runs(r):
        runs = []
        for j in range(region.shape[1]):
            if region[r][j] > 0:
                if len(runs) > 0 and runs[-1][1] == j-1:
                    runs[-1][1] = j
                else:
                    runs.append([j, j])
        
        return runs
    
    def extend_run(r, run):
        er = r
        for j in range(r+1, region.shape[0]):
            if np.all(region[j][run[0]:run[1]+1] > 0):
                region[j][run[0]:run[1]+1] = 0
                er = j
            else: break
        
        return ((r, run[0]), (er, run[1]))

    rngs = []
    for r in range(region.shape[0]):
        runs = get_runs(r)
        for run in runs:
            rng = extend_run(r, run)
            rngs.append(rng)

    return rngs

def mask_to_ranges(cells: List[List[CellData]], mask, key: Callable[[CellData], Any]):
    n = np.max(mask)
    regions = []
    for t in range(1, n+1):
        ind = np.unravel_index(np.argmax(mask == t), mask.shape)
        val = key(cells[ind[0]][ind[1]])
        rs = region_to_ranges(mask == t)

        sm = np.sum(mask == t)
        check = 0
        for r in rs:
            a = (r[1][0] - r[0][0] + 1) * (r[1][1] - r[0][1] + 1)
            check += a
        # print(t, sm, check, rs)
        assert(sm == check)
        
        region = {
            'val': val,
            'ranges': rs
        }

        regions.append(region)

    return regions

def get_regions(sheet: SheetData, key: Callable[[CellData], Any] = lambda x: x.value, cmp: Callable[[Any, Any], bool] = lambda x, y: x == y):
    cells = sheet.data
    mask = merge_mask(cells, key=key, cmp=cmp)
    regions = mask_to_ranges(cells, mask, key)
    
    if len(cells) > 0 and len(cells[0]) > 0:
        st = sheet.usedRange.split(':')[0].split('!')[-1]
        row, col = address_to_coord(st)
        
        for i, region in enumerate(regions):
            ranges = [':'.join([coord_to_address(x[0] + row, x[1] + col) for x in r]) for r in region['ranges']]
            regions[i] = {
                'val': region['val'],
                'ranges': ranges
            }
        
    return regions