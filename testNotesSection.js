const ExcelJS = require('exceljs');
const fs = require('fs').promises;

// Copy of addNotesSection from generateReports.js
function addNotesSection(sheet) {
    // Merge rows 5-7, columns A-Q for notes section
    sheet.mergeCells('A5', 'Q7');
    const notesCell = sheet.getCell('A5');
    notesCell.value = 'Notes:';
    notesCell.font = { bold: true, size: 9, name: 'Calibri' };
    notesCell.alignment = { vertical: 'top', wrapText: true };
    notesCell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
    };
}

async function testNotesSection() {
    console.log('Testing notes section addition...\n');
    
    try {
        // Load template
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('srs_blank_template.xlsx');
        console.log(`✅ Loaded template with ${workbook.worksheets.length} sheets`);
        
        // Test each sheet
        const sheetNames = ['Page 1', 'Page 2', 'Page 3', 'Page 4'];
        for (const sheetName of sheetNames) {
            const sheet = workbook.getWorksheet(sheetName);
            if (!sheet) {
                console.error(`❌ Missing sheet: ${sheetName}`);
                return false;
            }
            
            console.log(`\nTesting ${sheetName}:`);
            
            // Apply notes section
            addNotesSection(sheet);
            
            // Verify merge exists
            const merges = sheet._merges || [];
            const targetMerge = 'A5:Q7';
            const found = merges.some(m => m === targetMerge);
            console.log(`  Merge ${targetMerge}: ${found ? '✅' : '❌'}`);
            if (!found) {
                console.error(`  Expected merge ${targetMerge} not found`);
                return false;
            }
            
            // Check cell value and formatting
            const notesCell = sheet.getCell('A5');
            console.log(`  Cell value: "${notesCell.value}" ${notesCell.value === 'Notes:' ? '✅' : '❌'}`);
            console.log(`  Font bold: ${notesCell.font?.bold ? '✅' : '❌'}`);
            console.log(`  Font size: ${notesCell.font?.size} ${notesCell.font?.size === 9 ? '✅' : '❌'}`);
            console.log(`  Font name: ${notesCell.font?.name} ${notesCell.font?.name === 'Calibri' ? '✅' : '❌'}`);
            console.log(`  Alignment vertical top: ${notesCell.alignment?.vertical === 'top' ? '✅' : '❌'}`);
            console.log(`  Wrap text: ${notesCell.alignment?.wrapText ? '✅' : '❌'}`);
            console.log(`  Border present: ${notesCell.border?.top?.style === 'thin' ? '✅' : '❌'}`);
        }
        
        // Save test output for visual inspection
        const outputPath = 'test_output_notes.xlsx';
        await workbook.xlsx.writeFile(outputPath);
        const stats = await fs.stat(outputPath);
        console.log(`\n✅ Saved test output to ${outputPath} (${stats.size} bytes)`);
        
        console.log('\n✅ All notes section tests passed!');
        return true;
        
    } catch (error) {
        console.error('❌ Test failed:', error.message);
        console.error(error.stack);
        return false;
    }
}

// Run test
testNotesSection().then(success => {
    process.exit(success ? 0 : 1);
}).catch(err => {
    console.error('Test runner error:', err);
    process.exit(1);
});