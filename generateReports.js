document.addEventListener('DOMContentLoaded', () => {
    const csvFileInput = document.getElementById('csvFileInput');
    const fileNameElement = document.getElementById('fileName');
    const processBtn = document.getElementById('processBtn');
    const statusElement = document.getElementById('status');
    const reportListElement = document.getElementById('reportList');
    const errorsSection = document.getElementById('errors-section');
    const errorList = document.getElementById('errorList');

    let selectedFile;

    csvFileInput.addEventListener('change', (event) => {
        if (event.target.files.length > 0) {
            selectedFile = event.target.files[0];
            fileNameElement.textContent = `Selected file: ${selectedFile.name}`;
            processBtn.disabled = false;
        } else {
            selectedFile = null;
            fileNameElement.textContent = '';
            processBtn.disabled = true;
        }
    });

    processBtn.addEventListener('click', () => {
        if (!selectedFile) {
            alert("Please select a file first.");
            return;
        }

        // Reset UI
        processBtn.disabled = true;
        statusElement.textContent = 'Processing...';
        statusElement.classList.remove('text-red-600');
        reportListElement.innerHTML = '<p class="placeholder text-center text-gray-500">Generating reports...</p>';
        errorList.innerHTML = '';
        errorsSection.classList.add('hidden');

        const reader = new FileReader();
        reader.onload = (event) => {
            const csvData = event.target.result;
            try {
                // 1. Validate the raw CSV data first
                const validationErrors = validateCSVData(csvData);
                if (validationErrors.length > 0) {
                    errorsSection.classList.remove('hidden');
                    validationErrors.forEach(error => {
                        const errorItem = document.createElement('div');
                        errorItem.className = 'p-2 border-b border-red-100 last:border-b-0';
                        errorItem.textContent = error.toReadableMessage();
                        errorList.appendChild(errorItem);
                    });
                }

                // 2. Parse the data with Papa Parse to get key-value objects
                Papa.parse(csvData, {
                    header: true,
                    skipEmptyLines: true,
                    complete: (results) => {
                        try {
                            if (results.errors.length) {
                                let papaErrors = results.errors.map(e => `Error on row ${e.row}: ${e.message}`).join('\n');
                                throw new Error(`CSV parsing failed:\n${papaErrors}`);
                            }

                            // 3. Process the parsed data
                            processData(results.data);

                            // 4. Update final status
                            if (validationErrors.length > 0) {
                                statusElement.textContent = `Processing complete with ${validationErrors.length} warning(s).`;
                                statusElement.classList.add('text-red-600');
                            } else {
                                statusElement.textContent = 'Processing complete!';
                            }
                        } catch (e) {
                            console.error("Failed during processing:", e);
                            statusElement.textContent = `Error: ${e.message}`;
                            statusElement.classList.add('text-red-600');
                        } finally {
                            processBtn.disabled = false;
                        }
                    }
                });

            } catch (error) { // Catches errors from the initial validation step
                console.error("Failed during validation:", error);
                statusElement.textContent = `Error: ${error.message}`;
                statusElement.classList.add('text-red-600');
                processBtn.disabled = false;
            }
        };
        reader.onerror = () => {
            alert('Error reading file.');
            statusElement.textContent = 'Error reading file.';
            statusElement.classList.add('text-red-600');
            processBtn.disabled = false;
        };
        reader.readAsText(selectedFile);
    });



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
            return name
        }
        const parts = name.split(' ').filter(p => p);
        if (parts.length > 1) {
            return `${parts[parts.length - 1]}, ${parts.slice(0, -1).join(' ')}`;
        }
        return name;
    }

    function processData(allRows) {
        const groupNameKey = 'Program Day: Group: Class Name'
        const uniqueGroups = [...new Set(allRows.map(row => row[groupNameKey]).filter(Boolean))];

        if (uniqueGroups.length === 0) {
            if (uniqueGroups.length === 0) {
                throw new Error(`Could not find any groups under the column "${groupNameKey}".`);
            }
        }

        reportListElement.innerHTML = '';

        uniqueGroups.forEach(groupName => {
            const attendanceData = {};
            const presentRows = allRows.filter(row => row['Outcome'] === 'Present' && row[groupNameKey] === groupName);

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

            generateReport(groupName, attendanceData);
        });
    }

    // Helper functions for template-based SRS generation
    function getCleanGroupName(fullGroupName) {
        return fullGroupName.includes(' - ') ? fullGroupName.split(' - ').pop().trim() : fullGroupName;
    }

    function getSafeFileName(groupName) {
        const cleanName = getCleanGroupName(groupName);
        return cleanName.replace(/[^a-z0-9]/gi, '_');
    }

    function populateSheetWithStudents(sheet, students, attendanceData, startRow = 10) {
        // Populate names in column A and grades in column B
        students.forEach((studentName, index) => {
            const row = startRow + index;
            const studentInfo = attendanceData[studentName];
            sheet.getCell(`A${row}`).value = studentName;
            sheet.getCell(`B${row}`).value = studentInfo.grade || '';
        });
        
        // Clear any remaining rows in the student range (A10:A43, B10:B43)
        const totalRows = 34; // A10:A43 inclusive
        for (let i = students.length; i < totalRows; i++) {
            const row = startRow + i;
            sheet.getCell(`A${row}`).value = '';
            sheet.getCell(`B${row}`).value = '';
        }
    }

    async function loadTemplateWorkbook() {
        try {
            // Check if running from file:// protocol (fetch may fail due to CORS)
            if (window.location.protocol === 'file:') {
                console.warn('Running from file:// protocol - fetch may fail due to CORS. Consider using a local server.');
            }
            
            console.log('Fetching template file...');
            const response = await fetch('srs_blank_template.xlsx');
            if (!response.ok) {
                throw new Error(`Failed to load template: ${response.status} ${response.statusText}`);
            }
            const arrayBuffer = await response.arrayBuffer();
            console.log('Template fetched, loading with ExcelJS...');
            
            // Create workbook instance and load the template
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(arrayBuffer);
            console.log('Template loaded into workbook');
            return workbook;
        } catch (error) {
            console.error('Template loading failed:', error);
            // Provide more helpful error message
            if (error.message.includes('NetworkError') || error.message.includes('Failed to fetch')) {
                console.error('Network error detected. This may be due to CORS restrictions when running from file:// protocol.');
                console.error('Please run this application using a local HTTP server (e.g., python -m http.server 8000)');
            }
            throw error;
        }
    }

    // Clone a worksheet's structure (column widths, merge cells, basic formatting)
    function cloneSheetStructure(sourceSheet, targetSheet) {
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
        
        // Copy row heights for key rows (optional)
        // We'll copy rows 1-44 (the template range)
        for (let rowNum = 1; rowNum <= 44; rowNum++) {
            const sourceRow = sourceSheet.getRow(rowNum);
            const targetRow = targetSheet.getRow(rowNum);
            if (sourceRow.height) {
                targetRow.height = sourceRow.height;
            }
        }
        
        // Note: Cell styles and values are not copied - template should have them
        // but we will populate names/grades later
    }
    
    // Duplicate a worksheet including cell values and styles
    function duplicateSheet(sourceSheet, targetSheet) {
        console.log(`Duplicating sheet "${sourceSheet.name}" to "${targetSheet.name}"`);
        
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
        
        // Copy rows 1-44 (template range) - including values and basic styles
        for (let rowNum = 1; rowNum <= 44; rowNum++) {
            const sourceRow = sourceSheet.getRow(rowNum);
            const targetRow = targetSheet.getRow(rowNum);
            
            // Copy row height
            if (sourceRow.height) {
                targetRow.height = sourceRow.height;
            }
            
            // Copy each cell in the row
            sourceRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
                const targetCell = targetRow.getCell(colNumber);
                
                // Copy value
                targetCell.value = sourceCell.value;
                
                // Copy basic styling (font, fill, alignment, border)
                if (sourceCell.font) {
                    targetCell.font = { ...sourceCell.font };
                }
                if (sourceCell.fill) {
                    targetCell.fill = { ...sourceCell.fill };
                }
                if (sourceCell.alignment) {
                    targetCell.alignment = { ...sourceCell.alignment };
                }
                if (sourceCell.border) {
                    targetCell.border = { ...sourceCell.border };
                }
                if (sourceCell.numFmt) {
                    targetCell.numFmt = sourceCell.numFmt;
                }
            });
        }
        
        console.log(`Sheet duplication complete`);
    }

    async function generateReport(groupName, attendanceData) {
        // Legacy report generation (original implementation)
        async function generateReportLegacy(groupName, attendanceData) {
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Attendance', {
                views: [{ state: 'frozen', ySplit: 1 }]
            });

            worksheet.columns = [
                { header: 'Student Name', key: 'studentName', width: 30 },
                { header: 'Grade', key: 'grade', width: 10 },
                { header: 'Monday', key: 'monday', width: 15 },
                { header: 'Tuesday', key: 'tuesday', width: 15 },
                { header: 'Wednesday', key: 'wednesday', width: 15 },
                { header: 'Thursday', key: 'thursday', width: 15 },
                { header: 'Friday', key: 'friday', width: 15 },
            ];

            const sortedStudents = Object.keys(attendanceData).sort();

            sortedStudents.forEach(studentName => {
                const studentInfo = attendanceData[studentName];
                const daysAttended = studentInfo.days;
                worksheet.addRow({
                    studentName: studentName,
                    grade: studentInfo.grade,
                    monday: daysAttended.has('Monday') ? '✔' : '',
                    tuesday: daysAttended.has('Tuesday') ? '✔' : '',
                    wednesday: daysAttended.has('Wednesday') ? '✔' : '',
                    thursday: daysAttended.has('Thursday') ? '✔' : '',
                    friday: daysAttended.has('Friday') ? '✔' : '',
                });
            });

            // Style the header
            worksheet.getRow(1).eachCell(cell => {
                cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FF2E4756' }
                };
                cell.alignment = { vertical: 'middle', horizontal: 'center' };
            });
            
            // Style data rows with alternating colors
            worksheet.eachRow((row, rowNumber) => {
                if (rowNumber > 1) { // Start from the first data row
                    const isEvenRow = rowNumber % 2 === 0;
                    row.eachCell(cell => {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: isEvenRow ? 'FFFFFFFF' : 'FFF0F0F0' } // White or Light Grey
                        };
                        // Apply center alignment to checkmark cells (Monday to Friday columns)
                        if (cell.value === '✔') {
                            cell.alignment = { vertical: 'middle', horizontal: 'center' };
                        }
                    });
                }
            });

            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);

            const className = groupName.includes(' - ') ? groupName.split(' - ').pop().trim() : groupName;
            const safeGroupName = className.replace(/[^a-z0-9]/gi, '_');
            const reportItem = document.createElement('div');
            reportItem.className = 'report-item flex justify-between items-center p-3 border-b border-gray-200 last:border-b-0';
            reportItem.innerHTML = `
                <span class="font-medium text-gray-800">${className}</span>
                <a href="${url}" download="${safeGroupName}.xlsx" class="bg-green-500 text-white text-sm font-semibold py-1 px-3 rounded-md hover:bg-green-600 transition-colors">Download</a>
            `;
            reportListElement.appendChild(reportItem);
        }

        // Try template-based generation first, fall back to legacy if needed
        console.log(`Generating report for: ${groupName}`);
        try {
            console.log('Attempting template-based generation...');
            
            // 1. Load template
            const templateWorkbook = await loadTemplateWorkbook();
            console.log('Template loaded successfully');
            
            // 2. Prepare data
            const sortedStudents = Object.keys(attendanceData).sort();
            console.log(`Total students: ${sortedStudents.length}`);
            const studentsPerSheet = 34; // A10:A43 inclusive
            const totalSheetsNeeded = Math.ceil(sortedStudents.length / studentsPerSheet);
            console.log(`Sheets needed: ${totalSheetsNeeded}`);
            
            // 3. Get clean group name for sheet population
            const cleanGroupName = getCleanGroupName(groupName);
            console.log(`Clean group name: ${cleanGroupName}`);
            
            // 4. Ensure we have enough sheets (template has "Page 1", "Page 2", "Page 3", "Page 4")
            // Use existing sheets, duplicate "Page 1" if more needed
            const templateSheet = templateWorkbook.getWorksheet('Page 1');
            if (!templateSheet) {
                throw new Error('Template missing "Page 1" sheet');
            }
            console.log('Found Page 1 sheet');
            
            // Handle multiple sheets - template has 4 pre-formatted sheets
            const sheetNames = ['Page 1', 'Page 2', 'Page 3', 'Page 4'];
            for (let sheetIndex = 0; sheetIndex < totalSheetsNeeded; sheetIndex++) {
                let sheet;
                if (sheetIndex < sheetNames.length) {
                    // Use existing pre-formatted sheet if available
                    const sheetName = sheetNames[sheetIndex];
                    sheet = templateWorkbook.getWorksheet(sheetName);
                    if (sheet) {
                        console.log(`Using existing ${sheetName} sheet`);
                    } else {
                        // Sheet not found in template, duplicate Page 1
                        sheet = templateWorkbook.addWorksheet(sheetName);
                        duplicateSheet(templateSheet, sheet);
                        console.log(`${sheetName} not found in template, duplicated Page 1`);
                    }
                } else {
                    // Beyond 4 sheets, duplicate Page 1 template
                    sheet = templateWorkbook.addWorksheet(`Page ${sheetIndex + 1}`);
                    duplicateSheet(templateSheet, sheet);
                    console.log(`Created Page ${sheetIndex + 1} sheet by duplication`);
                }
                

                
                // Rename sheet to include group name? Keeping as Page X for simplicity
                // sheet.name = `${cleanGroupName} - Page ${sheetIndex + 1}`;
                
                // Calculate student range for this sheet
                const startIdx = sheetIndex * studentsPerSheet;
                const endIdx = Math.min(startIdx + studentsPerSheet, sortedStudents.length);
                const sheetStudents = sortedStudents.slice(startIdx, endIdx);
                console.log(`Sheet ${sheetIndex + 1}: Students ${startIdx + 1}-${endIdx} (${sheetStudents.length} students)`);
                
                // Populate names (A10:A43) and grades (B10:B43)
                populateSheetWithStudents(sheet, sheetStudents, attendanceData, 10);
            }
            
            // 5. Generate download
            console.log('Generating Excel file...');
            const buffer = await templateWorkbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            
            const safeGroupName = getSafeFileName(groupName);
            const className = cleanGroupName;
            const reportItem = document.createElement('div');
            reportItem.className = 'report-item flex justify-between items-center p-3 border-b border-gray-200 last:border-b-0';
            reportItem.innerHTML = `
                <span class="font-medium text-gray-800">${className} (Template)</span>
                <a href="${url}" download="${safeGroupName}.xlsx" class="bg-green-500 text-white text-sm font-semibold py-1 px-3 rounded-md hover:bg-green-600 transition-colors">Download</a>
            `;
            reportListElement.appendChild(reportItem);
            
            console.log(`✅ Template report generated for ${className}`);
            
        } catch (templateError) {
            console.error('❌ Template generation failed:', templateError);
            console.warn('Falling back to legacy method...');
            // Fallback to legacy generation
            await generateReportLegacy(groupName, attendanceData);
        }
    }
});