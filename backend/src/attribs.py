import difflib
from utils.conversion import excel_col_to_num
from api.models import CellData

def value_key(x: CellData):
    v = x.value
    return None if v is None or len(v) == 0 else 1

def text_key(x: CellData):
    k = x.text
    return None if k is None or len(k) == 0 else 1

def formula_cmp(x: str, y: str):
    if x == y: return True
    if x is None or y is None: return False
    
    diff = [(i, s) for i, s in enumerate(difflib.ndiff(x, y)) if s[0] in ['-', '+']]
    gdiff = []
    for d in diff:
        join = True
        if len(gdiff) == 0:
            join = False
        else:
            if gdiff[-1]['type'] != d[1][0] or abs(gdiff[-1]['ind'] - d[0]) > 1:
                join = False

        if join: 
            gdiff[-1]['value'] += d[1][-1]
        else:
            gdiff.append({'type': d[1][0], 'ind': d[0], 'value': d[1][-1]})

    similar = 0
    for i, d in enumerate(gdiff):
        if d['type'] == '-' and i < len(gdiff)-1 and gdiff[i+1]['type'] == '+':
            a, b = d['value'], gdiff[i+1]['value']
            if a.isdigit() and b.isdigit(): # consec row ref
                a, b = int(a), int(b)
                if abs(a - b) == 1:
                    similar += 1
            elif a.isalpha() and b.isalpha(): # consec col ref
                if abs(excel_col_to_num(a) - excel_col_to_num(b)) == 1:
                    similar += 1
    
    if similar > (len(gdiff)//2) * 0.9: return True # 0.9 for fuzziness

    return False

def formula_key(x: CellData):
    s = x.formula.strip()
    return s if len(s) > 0 else None

def format_key(x: CellData):
    s = x.numberFormat
    fk = font_key(x)
    if s == 'General' and fk is None:
        return None

    return f'{s}-{fk}'

def color_key(x: CellData):
    col = x.color
    if col is None or col in ['#FFFFFF', '#FFF']: return None
    return col

def font_key(x: CellData):
    f = x.font
    if f is None or text_key(x) is None:
        return None 

    return f'{f.name}-{f.size}-{f.bold}'