document.addEventListener('DOMContentLoaded', () => {
    const csvFileInput = document.getElementById('csvFileInput');
    const fileNameElement = document.getElementById('fileName');
    const processBtn = document.getElementById('processBtn');
    const statusElement = document.getElementById('status');
    const reportListElement = document.getElementById('reportList');

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
        reportListElement.innerHTML = '<p class="placeholder text-center text-gray-500">Generating reports...</p>';

        const reader = new FileReader();
        reader.onload = (event) => {
            const csvData = event.target.result;
            try {
                processData(csvData);
                statusElement.textContent = 'Processing complete!';
            } catch (error) {
                console.error("Failed to process data:", error);
                statusElement.textContent = `Error: ${error.message}`;
                alert(`An error occurred: ${error.message}`);
            } finally {
                processBtn.disabled = false;
            }
        };
        reader.onerror = () => {
            alert('Error reading file.');
            statusElement.textContent = 'Error reading file.';
            processBtn.disabled = false;
        };
        reader.readAsText(selectedFile);
    });

    function parseCSV(text) {
        const lines = text.trim().split('\n');
        // Match fields that are enclosed in quotes. This is more robust.
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

        // Dynamically find unique groups
        const uniqueGroups = [...new Set(allRows.map(row => row[groupNameKey]).filter(Boolean))];

        if (uniqueGroups.length === 0) {
            throw new Error(`Could not find any groups under the column "${groupNameKey}".`);
        }

        reportListElement.innerHTML = ''; // Clear placeholder

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

    function generateReport(groupName, attendanceData) {
        const workbook = XLSX.utils.book_new();
        const header = ["Student Name", "Grade", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
        const sheetData = [header];

        const sortedStudents = Object.keys(attendanceData).sort();

        sortedStudents.forEach(studentName => {
            const studentInfo = attendanceData[studentName];
            const daysAttended = studentInfo.days;
            const row = [
                studentName,
                studentInfo.grade,
                daysAttended.has('Monday') ? '✔' : '',
                daysAttended.has('Tuesday') ? '✔' : '',
                daysAttended.has('Wednesday') ? '✔' : '',
                daysAttended.has('Thursday') ? '✔' : '',
                daysAttended.has('Friday') ? '✔' : '',
            ];
            sheetData.push(row);
        });

        const worksheet = XLSX.utils.aoa_to_sheet(sheetData);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance");

        const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        const blob = new Blob([wbout], { type: 'application/octet-stream' });
        const url = URL.createObjectURL(blob);

        // Create download link in the UI
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