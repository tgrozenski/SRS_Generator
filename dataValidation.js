class DataValidationError extends Error {
    constructor(message, errorType, student, lineNumber) {
        super(message);
        this.name = 'DataValidationError';
        this.errorType = errorType;
        this.student = student;
        this.lineNumber = lineNumber;
    }

    get studentName() { return this.student?.name; }
    get studentId() { return this.student?.id; }
    get group() { return this.student?.group; }
    get date() { return this.student?.date; }

    toReadableMessage() {
        const groupName = this.group?.includes(' - ') ? this.group.split(' - ').pop().trim() : this.group;
        let name = this.studentName || 'Unknown';
        if (!name.includes(',')) {
            const parts = name.split(' ').filter(p => p);
            if (parts.length > 1) {
                name = `${parts[parts.length - 1]}, ${parts.slice(0, -1).join(' ')}`;
            }
        }
        return `"${name}", ${this.errorType} - (${groupName}, ${this.date})`;
    }
}

function getStudentInfo(row) {
    return {
        name: row['Student: Full Name'] || row['Student: Last, First'] || 'Unknown',
        id: row['Student ID'] || null,
        group: row['Program Day: Group: Class Name'] || 'Unknown',
        date: row['Session Date'] || 'Unknown'
    };
}

function convert_to_12_hour(timeString) {
    if (!timeString) return '';

    if (timeString.includes('AM') || timeString.includes('PM')) {
        return timeString;
    }

    const parts = timeString.split(' ');
    let datePart = '';
    let timePart = '';

    if (parts.length > 1) {
        timePart = parts.pop();
        datePart = parts.join(' ');
    } else {
        timePart = timeString;
    }

    const timeParts = timePart.split(':');
    if (timeParts.length < 2) {
        return timeString;
    }

    let hours = parseInt(timeParts[0], 10);
    const minutes = timeParts[1];

    if (isNaN(hours) || isNaN(parseInt(minutes, 10))) {
        return timeString;
    }

    const ampm = hours >= 12 ? 'PM' : 'AM';
    hours = hours % 12;
    hours = hours === 0 ? 12 : hours;

    const secondsPart = timeParts.length > 2 ? `:${timeParts[2]}` : '';
    const newTime = `${hours}:${minutes}${secondsPart} ${ampm}`;

    return datePart ? `${datePart} ${newTime}` : newTime;
}

function parseCSVWithLineNumbers(text) {
    const lines = text.trim().split('\n');

    const parseLine = (line) => {
        const values = [];
        let current = '';
        let inQuotes = false;
        for (let i = 0; i < line.length; i++) {
            const char = line[i];
            if (char === '"') {
                inQuotes = !inQuotes;
                continue; // Don't add the quote to the value
            }
            if (char === ',' && !inQuotes) {
                values.push(current);
                current = '';
            } else {
                current += char;
            }
        }
        values.push(current);
        return values;
    };

    if (lines.length === 0 || !lines[0].trim()) return [];
    const headers = parseLine(lines[0]);
    const rows = [];

    for (let i = 1; i < lines.length; i++) {
        const line = lines[i];
        if (!line.trim()) continue;

        const values = parseLine(line);
        if (values.length === headers.length) {
            let row = {};
            headers.forEach((header, index) => {
                row[header] = values[index];
            });
            row.lineNumber = i + 1;
            rows.push(row);
        }
    }
    return rows;
}

function validateRow(row) {
    row['In Time'] = convert_to_12_hour(row['In Time']);
    row['Out Time'] = convert_to_12_hour(row['Out Time']);

    const {
        'Outcome': outcome,
        'In Time': inTime,
        'Out Time': outTime,
        'Program Day: Group: Class Name': groupName,
        'Session Date': sessionDate,
        'PG Drop-Off: Full Name': pgDropoff,
        lineNumber
    } = row;

    const student = getStudentInfo(row);

    // Rule 1: If Outcome is "Present", both "In time" and "Out time" must exist.
    if (outcome === 'Present' && (!inTime || !outTime)) {
        throw new DataValidationError('In time or Out time is missing for a "Present" outcome.', 'Missing Time', student, lineNumber);
    }

    // Rule 1b: If Outcome is "Present", "PG Drop-Off: Full Name" must exist.
    if (outcome === 'Present' && !pgDropoff) {
        throw new DataValidationError('PG Drop-Off name is missing for a "Present" outcome.', 'Missing Dropoff Name', student, lineNumber);
    }

    // Rule 2: "In time" must be at least 1 minute before "Out time".
    if (inTime && outTime) {
        const inDateTime = new Date(inTime)
        const outDateTime = new Date(outTime)

        if (inDateTime.getTime() + 60000 > outDateTime.getTime()) {
            throw new DataValidationError(`Out time (${outTime}) must be at least one minute after In time (${inTime}).`, 'Invalid Time Order', student, lineNumber);
        }
    }

    // Rule 3: If Outcome is "Scheduled" or "Absent", no time should exist.
    if ((outcome === 'Scheduled' || outcome === 'Absent') && (inTime || outTime)) {
        throw new DataValidationError('Time entries should not exist for a "Scheduled" or "Absent" outcome.', 'Unexpected Time', student, lineNumber);
    }
 
    // Rule 4: AM/PM time should match AM/PM group if group name ends with AM/PM.
    if (groupName && (groupName.endsWith(' AM') || groupName.endsWith(' PM'))) {
        if (inTime && outTime) {
            const isGroupAM = groupName.endsWith(' AM');
            const isInTimeAM = inTime.includes('AM');
            const isOutTimeAM = outTime.includes('AM');

            if (isGroupAM !== isInTimeAM || isGroupAM !== isOutTimeAM) {
                throw new DataValidationError(`Time AM/PM and group AM/PM should match.`, 'Time/Group Mismatch', student, lineNumber);
            }
        }
    }
}

function validateCSVData(csvData) {
    const allRows = parseCSVWithLineNumbers(csvData);
    const errors = [];

    allRows.forEach(row => {
        try {
            validateRow(row);
        } catch (error) {
            if (error instanceof DataValidationError) {
                errors.push(error);
            } else {
                const student = getStudentInfo(row);
                student.name = student.name || 'Unknown';
                student.group = student.group || 'Unknown';
                student.date = student.date || 'Unknown';
                
                errors.push(new DataValidationError(
                    `An unexpected error occurred: ${error.message}`,
                    'System Error',
                    student,
                    row.lineNumber || 0
                ));
            }
        }
    });

    return errors;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        validateCSVData,
        DataValidationError,
        convert_to_12_hour,
        getStudentInfo
    };
}
