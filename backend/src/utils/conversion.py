def excel_col_to_num(col):
    if not col.isalpha(): return False

    return sum((ord(c) - ord('A') + 1) * (26 ** (len(col) - i - 1)) for i, c in enumerate(col))

def num_to_excel_col(num):
    col = ''
    while num > 0:
        num, rem = divmod(num, 26)
        if rem == 0:
            col = 'Z' + col
            num -= 1
        else:
            col = chr(ord('A') + rem - 1) + col
    return col

def coord_to_address(r, c):
    return f'${num_to_excel_col(c+1)}${r+1}'

def address_to_coord(addr):
    cola = ''.join(c for c in addr if c.isalpha())
    row = ''.join(c for c in addr if c.isdigit())
    col = excel_col_to_num(cola)
    row = int(row)
    return row-1, col-1
