const ExcelJS = require('exceljs');

async function test() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('srs_blank_template.xlsx');
    const sheet = workbook.getWorksheet('Page 1');
    
    console.log('Sheet methods:');
    console.log(Object.getOwnPropertyNames(Object.getPrototypeOf(sheet)).filter(name => !name.startsWith('_')).sort());
    
    // Check if clone exists
    console.log('\nHas clone?', typeof sheet.clone);
    console.log('Has duplicate?', typeof sheet.duplicate);
    
    // Check cell styles
    const cell = sheet.getCell('A1');
    console.log('\nCell A1 style keys:', Object.keys(cell));
    console.log('Cell style:', cell.style);
}

test().catch(console.error);