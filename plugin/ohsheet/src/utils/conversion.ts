export function excelColToNum(col: string): number | false {
    if (!/^[A-Za-z]+$/.test(col)) return false;

    return col.split('').reduce((acc, c, i) => {
        return acc + (c.toUpperCase().charCodeAt(0) - 64) * Math.pow(26, col.length - i - 1);
    }, 0);
}

export function numToExcelCol(num: number): string {
    let col = '';
    while (num > 0) {
        let rem = num % 26;
        num = Math.floor(num / 26);
        if (rem === 0) {
            col = 'Z' + col;
            num--;
        } else {
            col = String.fromCharCode(64 + rem) + col;
        }
    }
    return col;
}

export function coordToAddress(r: number, c: number): string {
    return `$${numToExcelCol(c + 1)}$${r + 1}`;
}

export function addressToCoord(addr: string): [number, number] {
    addr = addr.split('!').slice(-1)[0];
    const cola = addr.match(/[A-Za-z]+/)?.[0] || '';
    const row = addr.match(/\d+/)?.[0] || '';
    const col = excelColToNum(cola);
    
    if (col === false) {
        throw new Error('Invalid column reference');
    }
    
    return [parseInt(row) - 1, col - 1];
}
