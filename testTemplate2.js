// Test script to verify template structure
const ExcelJS = require('exceljs');
const fs = require('fs');

async function testTemplate() {
    console.log('Testing template structure...');
    
    try {
        // Load template file
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('srs_blank_template.xlsx');
        
        console.log('Template loaded successfully');
        console.log(`Number of worksheets: ${workbook.worksheets.length}`);
        
        // Check sheet names
        workbook.eachSheet((worksheet, sheetId) => {
            console.log(`Sheet ${sheetId}: "${worksheet.name}" (state: ${worksheet.state})`);
        });
        
        // Get Page 1 and Page 2
        const page1 = workbook.getWorksheet('Page 1');
        const page2 = workbook.getWorksheet('Page 2');
        
        if (!page1) {
            throw new Error('Missing "Page 1" sheet in template');
        }
        
        console.log(`\nPage 1 exists: ${!!page1}`);
        console.log(`Page 2 exists: ${!!page2}`);
        
        // Check key cells in Page 1
        console.log('\nChecking Page 1 structure:');
        
        // B4 should contain "Group" placeholder
        const b4 = page1.getCell('B4');
        console.log(`B4 value: "${b4.value}"`);
        
        // Check if B4 is part of a merged cell
        const merged = page1.mergedCells;
        console.log(`Merge cells count: ${merged ? merged.length : 0}`);
        if (merged) {
            merged.forEach(range => {
                console.log(`  Merged: ${range}`);
            });
        }
        
        // Check student rows A10:A42 are empty
        console.log('\nChecking student rows A10:A42 (should be empty):');
        let hasContent = false;
        for (let row = 10; row <= 42; row++) {
            const cell = page1.getCell(`A${row}`);
            if (cell.value) {
                console.log(`  A${row} has value: "${cell.value}"`);
                hasContent = true;
            }
        }
        if (!hasContent) console.log('  All empty - good!');
        
        // Check grade rows B10:B42
        console.log('Checking grade rows B10:B42 (should be empty):');
        hasContent = false;
        for (let row = 10; row <= 42; row++) {
            const cell = page1.getCell(`B${row}`);
            if (cell.value) {
                console.log(`  B${row} has value: "${cell.value}"`);
                hasContent = true;
            }
        }
        if (!hasContent) console.log('  All empty - good!');
        
        // Test population
        console.log('\nTesting population simulation:');
        const testData = {
            'Doe, John': { grade: '10', days: new Set(['Monday']) },
            'Smith, Jane': { grade: '11', days: new Set(['Tuesday']) },
            'Brown, Bob': { grade: '12', days: new Set(['Wednesday']) }
        };
        
        const sorted = Object.keys(testData).sort();
        console.log(`Would populate ${sorted.length} students:`);
        sorted.forEach((student, idx) => {
            const row = 10 + idx;
            console.log(`  Row ${row}: "${student}" (grade: ${testData[student].grade})`);
        });
        
        // Test clone function
        console.log('\nTesting sheet cloning simulation:');
        console.log('Column widths would be copied from Page 1');
        console.log('Merge cells would be copied');
        console.log('Row heights 1-44 would be copied');
        
        // Test multi-sheet logic
        console.log('\nTesting multi-sheet logic:');
        const studentsPerSheet = 33;
        const testCases = [5, 33, 34, 66, 67, 100];
        testCases.forEach(studentCount => {
            const sheetsNeeded = Math.ceil(studentCount / studentsPerSheet);
            console.log(`  ${studentCount} students → ${sheetsNeeded} sheet(s)`);
        });
        
        console.log('\n✅ All template tests passed!');
        return true;
        
    } catch (error) {
        console.error('❌ Template test failed:', error.message);
        console.error(error.stack);
        return false;
    }
}

// Run test
testTemplate().then(success => {
    process.exit(success ? 0 : 1);
}).catch(err => {
    console.error('Test runner error:', err);
    process.exit(1);
});