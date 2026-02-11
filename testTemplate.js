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
            
            // Check key cells if this is Page 1
                console.log('  Checking key cells for Page 1...');
                
                // Check group name cell (B4)
                const b4 = worksheet.getCell('B4');
                console.log(`    B4 value: "${b4.value}" (expected: maybe empty or placeholder)`);
                
                // Check student rows A10:A42
                console.log('    Checking student name rows (A10:A42):');
                for (let row = 10; row <= 42; row++) {
                    const cell = worksheet.getCell(`A${row}`);
                    if (cell.value) {
                        console.log(`      A${row}: "${cell.value}"`);
                    }
                }
                
                // Check grade rows B10:B42
                console.log('    Checking grade rows (B10:B42):');
                for (let row = 10; row <= 42; row++) {
                    const cell = worksheet.getCell(`B${row}`);
                    if (cell.value) {
                        console.log(`      B${row}: "${cell.value}"`);
                    }
                }
                
                // Check merged cells
                if (worksheet.mergeCells && worksheet.mergeCells.length) {
                    console.log(`    Merge cells: ${worksheet.mergeCells.length} ranges`);
                    worksheet.mergeCells.forEach(range => {
                        console.log(`      ${range}`);
                    });
                }
                
                // Check column widths
                console.log('    Column widths:');
                worksheet.columns.forEach((col, idx) => {
                    if (col && col.width) {
                        console.log(`      Column ${idx + 1}: ${col.width}`);
                    }
                });
            }
        });
        
        // Verify we have Page 1 and Page 2
        const page1 = workbook.getWorksheet('Page 1');
        const page2 = workbook.getWorksheet('Page 2');
        
        if (!page1) {
            throw new Error('Missing "Page 1" sheet in template');
        }
        
        console.log('\n✅ Template structure looks good!');
        console.log(`Page 1 exists: ${!!page1}`);
        console.log(`Page 2 exists: ${!!page2}`);
        
        // Test population function
        console.log('\nTesting population logic...');
        const testAttendanceData = {
            'Doe, John': { grade: '10', days: new Set(['Monday', 'Wednesday']) },
            'Smith, Jane': { grade: '11', days: new Set(['Tuesday', 'Thursday']) },
            'Brown, Bob': { grade: '12', days: new Set(['Friday']) }
        };
        
        const sortedStudents = Object.keys(testAttendanceData).sort();
        console.log(`Sample students: ${sortedStudents.join(', ')}`);
        
        // Test populateSheetWithStudents function (simulated)
        console.log('Simulating population of 3 students at row 10...');
        sortedStudents.forEach((student, index) => {
            const row = 10 + index;
            console.log(`  Row ${row}: A${row} = "${student}", B${row} = "${testAttendanceData[student].grade}"`);
        });
        
        // Test multi-sheet calculation
        const studentsPerSheet = 33;
        const totalStudents = 35; // Just over one sheet
        const sheetsNeeded = Math.ceil(totalStudents / studentsPerSheet);
        console.log(`\nMulti-sheet test: ${totalStudents} students need ${sheetsNeeded} sheets`);
        
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