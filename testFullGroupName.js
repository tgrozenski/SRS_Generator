const ExcelJS = require('exceljs');
const Papa = require('papaparse');
const fs = require('fs');

// Replicate helper functions from generateReports.js
function getDayOfWeek(dateStr) {
    try {
        const date = new Date(dateStr);
        const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        return days[date.getUTCDay()];
    } catch (e) {
        return null;
    }
}

function reverseName(name) {
    if (!name || name.includes(',')) {
        return name;
    }
    const parts = name.split(' ').filter(p => p);
    if (parts.length > 1) {
        return `${parts[parts.length - 1]}, ${parts.slice(0, -1).join(' ')}`;
    }
    return name;
}

function getCleanGroupName(fullGroupName) {
    return fullGroupName.includes(' - ') ? fullGroupName.split(' - ').pop().trim() : fullGroupName;
}

function setGroupNameInSheet(sheet, groupName) {
    // Set full group name in merged cell B3:G3 (row 3, columns B through G)
    sheet.getCell('B3').value = groupName;
    
    // Clear placeholder data from sample template
    sheet.getCell('I3').value = '';
    sheet.getCell('N3').value = '';
}

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





async function testFullGroupName() {
    console.log('Testing full group name placement with real CSV data...\n');
    
    try {
        // 1. Load CSV data
        const csvData = fs.readFileSync('testReports/SAR.csv', 'utf8');
        const results = Papa.parse(csvData, { header: true, skipEmptyLines: true });
        
        const groupNameKey = 'Program Day: Group: Class Name';
        const uniqueGroups = [...new Set(results.data.map(row => row[groupNameKey]).filter(Boolean))];
        
        console.log(`Found ${uniqueGroups.length} groups in CSV`);
        
        // Process first group
        const firstGroup = uniqueGroups[0];
        console.log(`\nProcessing group: "${firstGroup}"`);
        console.log(`Cleaned name: "${getCleanGroupName(firstGroup)}"`);
        
        // Build attendance data
        const attendanceData = {};
        const presentRows = results.data.filter(row => row['Outcome'] === 'Present' && row[groupNameKey] === firstGroup);
        
        presentRows.forEach(row => {
            const studentName = reverseName(row['Student: Full Name'] || row['Student: Last, First']);
            const sessionDate = row['Session Date'];
            const dayOfWeek = getDayOfWeek(sessionDate);
            const grade = row['Grade'] || '';
            
            if (!attendanceData[studentName]) {
                attendanceData[studentName] = { grade: grade, days: new Set() };
            }
            if (!attendanceData[studentName].grade && grade) {
                attendanceData[studentName].grade = grade;
            }
            if (dayOfWeek) {
                attendanceData[studentName].days.add(dayOfWeek);
            }
        });
        
        const sortedStudents = Object.keys(attendanceData).sort();
        console.log(`Unique students: ${sortedStudents.length}`);
        
        // 2. Load template
        const templateWorkbook = new ExcelJS.Workbook();
        await templateWorkbook.xlsx.readFile('srs_blank_template.xlsx');
        console.log('Template loaded');
        
        const studentsPerSheet = 34;
        const totalSheetsNeeded = Math.ceil(sortedStudents.length / studentsPerSheet);
        console.log(`Sheets needed: ${totalSheetsNeeded}`);
        
        // Ensure we don't exceed 4 sheets (136 students max)
        if (totalSheetsNeeded > 4) {
            throw new Error(`Group has ${sortedStudents.length} students, exceeding maximum capacity of 136 students (4 sheets)`);
        }
        
        const templateSheet = templateWorkbook.getWorksheet('Page 1');
        if (!templateSheet) throw new Error('Missing Page 1 sheet');
        
        // 3. Generate sheets using only existing template sheets
        const sheetNames = ['Page 1', 'Page 2', 'Page 3', 'Page 4'];
        for (let sheetIndex = 0; sheetIndex < totalSheetsNeeded; sheetIndex++) {
            const sheetName = sheetNames[sheetIndex];
            const sheet = templateWorkbook.getWorksheet(sheetName);
            if (!sheet) {
                throw new Error(`Template missing "${sheetName}" sheet`);
            }
            console.log(`  Using ${sheetName} sheet`);
            
            // Set group name
            setGroupNameInSheet(sheet, firstGroup);
            
            // Populate students
            const startIdx = sheetIndex * studentsPerSheet;
            const endIdx = Math.min(startIdx + studentsPerSheet, sortedStudents.length);
            const sheetStudents = sortedStudents.slice(startIdx, endIdx);
            console.log(`  Sheet ${sheetIndex + 1}: ${sheetStudents.length} students`);
            
            populateSheetWithStudents(sheet, sheetStudents, attendanceData, 10);
        }
        
        // 4. Save output
        const outputFile = 'test_output_full_group.xlsx';
        await templateWorkbook.xlsx.writeFile(outputFile);
        console.log(`\n✅ Saved output to ${outputFile}`);
        
        // 5. Verify group names
        console.log('\nVerifying group name placement...');
        const verifyWorkbook = new ExcelJS.Workbook();
        await verifyWorkbook.xlsx.readFile(outputFile);
        
        let allCorrect = true;
        // Check only the sheets that should have been populated
        for (let sheetIndex = 0; sheetIndex < totalSheetsNeeded; sheetIndex++) {
            const sheetName = sheetNames[sheetIndex];
            const sheet = verifyWorkbook.getWorksheet(sheetName);
            if (!sheet) {
                console.log(`  ❌ Sheet "${sheetName}" not found in output`);
                allCorrect = false;
                continue;
            }
            
            const b3 = sheet.getCell('B3');
            console.log(`  Sheet "${sheet.name}": B3 = "${b3.value}"`);
            if (b3.value !== firstGroup) {
                console.log(`    ❌ Expected: "${firstGroup}"`);
                allCorrect = false;
            } else {
                console.log(`    ✅ Correct`);
            }
            
            // Check placeholders cleared
            if (sheet.getCell('I3').value !== '' || sheet.getCell('N3').value !== '') {
                console.log(`    ❌ Placeholders not cleared`);
                allCorrect = false;
            }
        }
        
        // Log other sheets as not populated
        verifyWorkbook.eachSheet((sheet, sheetId) => {
            if (!sheetNames.slice(0, totalSheetsNeeded).includes(sheet.name)) {
                console.log(`  Sheet "${sheet.name}": not populated (skipping)`);
            }
        });
        
        if (allCorrect) {
            console.log('\n✅ All group names correctly placed!');
            // Clean up (commented for inspection)
            // fs.unlinkSync(outputFile);
            return true;
        } else {
            console.log('\n❌ Group name verification failed');
            return false;
        }
        
    } catch (error) {
        console.error('❌ Test failed:', error.message);
        console.error(error.stack);
        return false;
    }
}

// Run test
testFullGroupName().then(success => {
    process.exit(success ? 0 : 1);
}).catch(err => {
    console.error('Test runner error:', err);
    process.exit(1);
});