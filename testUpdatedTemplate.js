const ExcelJS = require('exceljs');

async function testUpdatedLogic() {
    console.log('Testing updated template logic...\n');
    
    try {
        // Load template
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('srs_blank_template.xlsx');
        
        console.log(`✅ Template loaded, sheets: ${workbook.worksheets.length}`);
        workbook.eachSheet((ws, id) => {
            console.log(`  Sheet ${id}: "${ws.name}"`);
        });
        
        // Test constants
        const studentsPerSheet = 34; // A10:A43 inclusive
        console.log(`\n✅ Students per sheet: ${studentsPerSheet}`);
        
        // Test different student counts
        const testCases = [
            { students: 5, expectedSheets: 1 },
            { students: 34, expectedSheets: 1 },
            { students: 35, expectedSheets: 2 },
            { students: 68, expectedSheets: 2 },
            { students: 69, expectedSheets: 3 },
            { students: 136, expectedSheets: 4 },
            { students: 137, expectedSheets: 5 }, // Needs sheet 5 (duplicate)
            { students: 95, expectedSheets: 3 },  // From logs: Think Cafe PM
            { students: 80, expectedSheets: 3 },  // Book Lab PM
            { students: 64, expectedSheets: 2 },  // Book Lab AM
        ];
        
        console.log('\n✅ Sheet calculation tests:');
        testCases.forEach(({ students, expectedSheets }) => {
            const sheetsNeeded = Math.ceil(students / studentsPerSheet);
            const passed = sheetsNeeded === expectedSheets;
            console.log(`  ${students} students → ${sheetsNeeded} sheets ${passed ? '✓' : `✗ (expected ${expectedSheets})`}`);
        });
        
        // Test sheet selection logic
        console.log('\n✅ Sheet selection logic:');
        const sheetNames = ['Page 1', 'Page 2', 'Page 3', 'Page 4'];
        const templateSheet = workbook.getWorksheet('Page 1');
        if (!templateSheet) throw new Error('Missing Page 1 sheet');
        
        for (let sheetIndex = 0; sheetIndex < 6; sheetIndex++) {
            let sheetName;
            if (sheetIndex < sheetNames.length) {
                sheetName = sheetNames[sheetIndex];
            } else {
                sheetName = `Page ${sheetIndex + 1}`;
            }
            console.log(`  Sheet index ${sheetIndex} → "${sheetName}"`);
        }
        
        // Test row population range
        console.log('\n✅ Row population range:');
        const startRow = 10;
        const totalRows = 34; // A10:A43 inclusive
        console.log(`  Start row: ${startRow}`);
        console.log(`  Total rows: ${totalRows}`);
        console.log(`  End row: ${startRow + totalRows - 1} (should be 43)`);
        
        // Verify template has empty rows 10-43
        const page1 = workbook.getWorksheet('Page 1');
        let emptyCount = 0;
        for (let row = 10; row <= 43; row++) {
            const cell = page1.getCell(`A${row}`);
            if (!cell.value || cell.value.toString().trim() === '') emptyCount++;
        }
        console.log(`\n✅ Template rows A10:A43 empty: ${emptyCount}/34 ${emptyCount === 34 ? '✓' : '✗'}`);
        
        console.log('\n✅ All tests completed!');
        return true;
        
    } catch (error) {
        console.error('❌ Test failed:', error.message);
        console.error(error.stack);
        return false;
    }
}

// Run test
testUpdatedLogic().then(success => {
    process.exit(success ? 0 : 1);
}).catch(err => {
    console.error('Test runner error:', err);
    process.exit(1);
});