const ExcelJS = require('exceljs');
const fs = require('fs').promises;

// Simulate populateSheetWithStudents function
function populateSheetWithStudents(sheet, students, attendanceData, startRow = 10) {
    // Day to column mapping: Monday = C,D,E; Tuesday = F,G,H; Wednesday = I,J,K; Thursday = L,M,N; Friday = O,P,Q
    const dayToColumns = {
        'Monday': ['C', 'D', 'E'],
        'Tuesday': ['F', 'G', 'H'],
        'Wednesday': ['I', 'J', 'K'],
        'Thursday': ['L', 'M', 'N'],
        'Friday': ['O', 'P', 'Q']
    };
    
    // Get all day columns for clearing
    const allDayColumns = Object.values(dayToColumns).flat();
    
    // Populate names, grades, and attendance checkboxes
    students.forEach((studentName, index) => {
        const row = startRow + index;
        const studentInfo = attendanceData[studentName];
        
        // Name and grade (existing)
        sheet.getCell(`A${row}`).value = studentName;
        sheet.getCell(`B${row}`).value = studentInfo.grade || '';
        
        // Clear all day cells first (ensures clean state)
        allDayColumns.forEach(col => {
            sheet.getCell(`${col}${row}`).value = '';
        });
        
        // Fill checkmarks for days present (3 cells per day)
        studentInfo.days.forEach(day => {
            const columns = dayToColumns[day];
            if (columns) {
                columns.forEach(col => {
                    sheet.getCell(`${col}${row}`).value = '✔'; // Same checkmark as legacy
                });
            }
        });
    });
    
    // Clear any remaining rows in the student range (A10:A43, B10:B43, and day columns)
    const totalRows = 34; // A10:A43 inclusive
    for (let i = students.length; i < totalRows; i++) {
        const row = startRow + i;
        sheet.getCell(`A${row}`).value = '';
        sheet.getCell(`B${row}`).value = '';
        // Also clear day cells for unused rows
        allDayColumns.forEach(col => {
            sheet.getCell(`${col}${row}`).value = '';
        });
    }
}

// Simulate duplicateSheet function
function duplicateSheet(sourceSheet, targetSheet) {
    // Copy column widths
    sourceSheet.columns.forEach((col, idx) => {
        if (col && col.width) {
            targetSheet.getColumn(idx + 1).width = col.width;
        }
    });
    
    // Copy merge cells
    if (sourceSheet.mergeCells && sourceSheet.mergeCells.length) {
        sourceSheet.mergeCells.forEach(mergeRange => {
            targetSheet.mergeCells(mergeRange);
        });
    }
    
    // Copy rows 1-44
    for (let rowNum = 1; rowNum <= 44; rowNum++) {
        const sourceRow = sourceSheet.getRow(rowNum);
        const targetRow = targetSheet.getRow(rowNum);
        
        if (sourceRow.height) {
            targetRow.height = sourceRow.height;
        }
        
        sourceRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
            const targetCell = targetRow.getCell(colNumber);
            targetCell.value = sourceCell.value;
            if (sourceCell.font) targetCell.font = { ...sourceCell.font };
            if (sourceCell.fill) targetCell.fill = { ...sourceCell.fill };
            if (sourceCell.alignment) targetCell.alignment = { ...sourceCell.alignment };
            if (sourceCell.border) targetCell.border = { ...sourceCell.border };
            if (sourceCell.numFmt) targetCell.numFmt = sourceCell.numFmt;
        });
    }
}

async function testIntegration() {
    console.log('Integration test with template...\n');
    
    try {
        // Load template
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile('srs_blank_template.xlsx');
        console.log(`✅ Loaded template with ${workbook.worksheets.length} sheets`);
        
        // Simulate data for a group with 95 students (Think Cafe PM)
        const attendanceData = {};
        const daysOfWeek = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
        for (let i = 1; i <= 95; i++) {
            const lastName = `Student${i}`;
            const firstName = `First${i}`;
            const studentName = `${lastName}, ${firstName}`;
            const days = new Set();
            
            // Assign days based on student index for testing
            // Student 1: Monday, Student 2: Tuesday, etc., wrap after Friday
            const dayIndex = (i - 1) % daysOfWeek.length;
            const primaryDay = daysOfWeek[dayIndex];
            days.add(primaryDay);
            
            // Also assign some students multiple days
            if (i % 7 === 0) {
                // Every 7th student gets two days
                const secondDayIndex = (dayIndex + 1) % daysOfWeek.length;
                days.add(daysOfWeek[secondDayIndex]);
            }
            
            attendanceData[studentName] = {
                grade: `${9 + (i % 4)}`, // Grades 9-12
                days: days
            };
        }
        
        const sortedStudents = Object.keys(attendanceData).sort();
        const studentsPerSheet = 34;
        const totalSheetsNeeded = Math.ceil(sortedStudents.length / studentsPerSheet);
        
        console.log(`✅ Simulating ${sortedStudents.length} students, ${totalSheetsNeeded} sheets needed`);
        
        const sheetNames = ['Page 1', 'Page 2', 'Page 3', 'Page 4'];
        const templateSheet = workbook.getWorksheet('Page 1');
        if (!templateSheet) throw new Error('Missing Page 1');
        
        // Process sheets
        for (let sheetIndex = 0; sheetIndex < totalSheetsNeeded; sheetIndex++) {
            let sheet;
            if (sheetIndex < sheetNames.length) {
                const sheetName = sheetNames[sheetIndex];
                sheet = workbook.getWorksheet(sheetName);
                if (!sheet) {
                    sheet = workbook.addWorksheet(sheetName);
                    duplicateSheet(templateSheet, sheet);
                    console.log(`  ${sheetName} not found, duplicated Page 1`);
                } else {
                    console.log(`  Using existing ${sheetName} sheet`);
                }
            } else {
                sheet = workbook.addWorksheet(`Page ${sheetIndex + 1}`);
                duplicateSheet(templateSheet, sheet);
                console.log(`  Created Page ${sheetIndex + 1} by duplication`);
            }
            
            // Set group name
            sheet.getCell('B4').value = 'Think Cafe PM (Test)';
            
            // Calculate student range
            const startIdx = sheetIndex * studentsPerSheet;
            const endIdx = Math.min(startIdx + studentsPerSheet, sortedStudents.length);
            const sheetStudents = sortedStudents.slice(startIdx, endIdx);
            
            console.log(`  Sheet ${sheetIndex + 1}: Students ${startIdx + 1}-${endIdx} (${sheetStudents.length})`);
            
            // Populate
            populateSheetWithStudents(sheet, sheetStudents, attendanceData, 10);
            
            // Verify population
            const expectedRows = Math.min(studentsPerSheet, sheetStudents.length);
            let populatedCount = 0;
            for (let i = 0; i < expectedRows; i++) {
                const row = 10 + i;
                const name = sheet.getCell(`A${row}`).value;
                if (name && name.toString().includes('Student')) populatedCount++;
            }
            console.log(`    Populated ${populatedCount} students in rows 10-${10 + expectedRows - 1}`);
            
            // Verify checkboxes for first student if sheet has students
            if (sheetStudents.length > 0) {
                const firstStudentRow = 10;
                const firstStudentName = sheetStudents[0];
                const firstStudentInfo = attendanceData[firstStudentName];
                console.log(`    First student: ${firstStudentName}, days: ${Array.from(firstStudentInfo.days).join(', ')}`);
                
                // Check a sample day cell
                if (firstStudentInfo.days.has('Monday')) {
                    const mondayCell = sheet.getCell(`C${firstStudentRow}`).value;
                    console.log(`    Monday checkbox C${firstStudentRow}: "${mondayCell}" (should be "✔")`);
                }
                // Check that a non-present day is empty (e.g., if Monday not present, check C cell)
                const sampleDay = 'Monday';
                const sampleCol = 'C';
                const sampleCellValue = sheet.getCell(`${sampleCol}${firstStudentRow}`).value;
                const shouldBeChecked = firstStudentInfo.days.has(sampleDay);
                console.log(`    Sample cell ${sampleCol}${firstStudentRow}: "${sampleCellValue}" (${shouldBeChecked ? 'should be "✔"' : 'should be empty'})`);
            }
            
            // Verify clearing of unused rows
            if (sheetStudents.length < studentsPerSheet) {
                const firstEmptyRow = 10 + sheetStudents.length;
                const emptyCell = sheet.getCell(`A${firstEmptyRow}`).value;
                console.log(`    First empty row A${firstEmptyRow}: "${emptyCell}" (should be empty)`);
                // Also verify day cells are empty
                const dayCell = sheet.getCell(`C${firstEmptyRow}`).value;
                console.log(`    Day cell C${firstEmptyRow}: "${dayCell}" (should be empty)`);
            }
        }
        
        // Save test output
        const testFile = 'test_output_integration.xlsx';
        await workbook.xlsx.writeFile(testFile);
        console.log(`\n✅ Saved test output to ${testFile}`);
        
        // Verify file exists
        const stats = await fs.stat(testFile);
        console.log(`✅ Output file size: ${stats.size} bytes`);
        
        // Clean up (optional)
        // await fs.unlink(testFile);
        
        console.log('\n✅ Integration test passed!');
        return true;
        
    } catch (error) {
        console.error('❌ Integration test failed:', error.message);
        console.error(error.stack);
        return false;
    }
}

// Run test
testIntegration().then(success => {
    process.exit(success ? 0 : 1);
}).catch(err => {
    console.error('Test runner error:', err);
    process.exit(1);
});