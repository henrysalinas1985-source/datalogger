const ExcelJS = require('C:/Users/hesalinas/Documents/SAP/TERUMO/Calibraciones/node_modules/exceljs');

async function inspectExcel() {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile('C:/Users/hesalinas/Documents/SAP/TERUMO/Calibraciones/CALIBRACIONES- DATALOG/DATALOGGER.xlsx');
    const ws = wb.worksheets[0];
    console.log('Sheet:', ws.name);
    console.log('RowCount:', ws.rowCount, 'ColCount:', ws.columnCount);

    for (let r = 1; r <= Math.min(ws.rowCount, 55); r++) {
        const row = ws.getRow(r);
        for (let c = 1; c <= Math.min(ws.columnCount, 16); c++) {
            const cell = row.getCell(c);
            const val = cell.text || (cell.value && typeof cell.value === 'object' ? JSON.stringify(cell.value).substring(0, 40) : String(cell.value || ''));
            if (val && val.trim() && val.trim() !== '0') {
                const addr = `${String.fromCharCode(64 + c)}${r}`;
                console.log(`${addr}: ${val.substring(0, 70)}`);
            }
        }
    }
}
inspectExcel().catch(console.error);
