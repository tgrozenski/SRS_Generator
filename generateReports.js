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

                processData(csvData);

                if (validationErrors.length > 0) {
                    statusElement.textContent = `Processing complete with ${validationErrors.length} warning(s).`;
                    statusElement.classList.add('text-red-600');
                } else {
                    statusElement.textContent = 'Processing complete!';
                }

            } catch (error) {
                console.error("Failed to process data:", error);
                statusElement.textContent = `Error: ${error.message}`;
                statusElement.classList.add('text-red-600');
                alert(`An error occurred: ${error.message}`);
            } finally {
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

    function parseCSV(text) {
        const lines = text.trim().split('\n');
        const headerRegex = /"([^"]*)"/g;
        const headers = [...lines[0].matchAll(headerRegex)].map(match => match[1]);

        const rows = [];
        for (let i = 1; i < lines.length; i++) {
            const values = [...lines[i].matchAll(headerRegex)].map(match => match[1]);
            if (values.length === headers.length) {
                let row = {};
                headers.forEach((header, index) => {
                    row[header] = values[index];
                });
                rows.push(row);
            }
        }
        return rows;
    }

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
        const parts = name.split(' ').filter(p => p);
        if (parts.length > 1) {
            return `${parts[parts.length - 1]}, ${parts.slice(0, -1).join(' ')}`;
        }
        return name;
    }

    function processData(csvData) {
        const allRows = parseCSV(csvData);
        const groupNameKey = 'Program Day: Group: Class Name';

        const uniqueGroups = [...new Set(allRows.map(row => row[groupNameKey]).filter(Boolean))];

        if (uniqueGroups.length === 0) {
            throw new Error(`Could not find any groups under the column "${groupNameKey}".`);
        }

        reportListElement.innerHTML = '';

        uniqueGroups.forEach(groupName => {
            const attendanceData = {};
            const presentRows = allRows.filter(row => row['Outcome'] === 'Present' && row[groupNameKey] === groupName);

            presentRows.forEach(row => {
                const studentName = reverseName(row['Student: Full Name']);
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

    async function generateReport(groupName, attendanceData) {
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
});