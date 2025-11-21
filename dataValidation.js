class DataValidationError extends Error {
    constructor(message, lineNumber, studentName, group, errorType) {
        super(message);
        this.name = 'DataValidationError';
        this.lineNumber = lineNumber;
        this.studentName = studentName;
        this.group = group;
        this.errorType = errorType;
    }

    toReadableMessage() {
        const groupName = this.group.includes(' - ') ? this.group.split(' - ').pop().trim() : this.group;
        return `Error on line ${this.lineNumber}: For student "${this.studentName}" in group "${groupName}", a "${this.errorType}" error occurred`;
    }
}

function parseCSVWithLineNumbers(text) {
    const lines = text.trim().split('\n');
    const headerRegex = /"([^"]*)"/g;
    const headers = [...lines[0].matchAll(headerRegex)].map(match => match[1]);

    const rows = [];
    for (let i = 1; i < lines.length; i++) {
        const line = lines[i];
        if (!line.trim()) continue;

        const values = [...line.matchAll(headerRegex)].map(match => match[1]);
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
    const {
        'Outcome': outcome,
        'In Time': inTime,
        'Out Time': outTime,
        'Student: Full Name': studentName,
        'Program Day: Group: Class Name': groupName,
        lineNumber
    } = row;

    // Rule 1: If Outcome is "Present", both "In time" and "Out time" must exist.
    if (outcome === 'Present' && (!inTime || !outTime)) {
        throw new DataValidationError('In time or Out time is missing for a "Present" outcome.', lineNumber, studentName, groupName, 'Missing Time');
    }

    // Rule 2: "In time" must be at least 1 minute before "Out time".
    if (inTime && outTime) {
        const inDateTime = new Date(inTime)
        const outDateTime = new Date(outTime)

        if (inDateTime.getTime() + 60000 > outDateTime.getTime()) {
            throw new DataValidationError(`Out time (${outTime}) must be at least one minute after In time (${inTime}).`, lineNumber, studentName, groupName, 'Invalid Time Order');
        }
    }

    // Rule 3: If Outcome is "Scheduled" or "Absent", no time should exist.
    if ((outcome === 'Scheduled' || outcome === 'Absent') && (inTime || outTime)) {
        throw new DataValidationError('Time entries should not exist for a "Scheduled" or "Absent" outcome.', lineNumber, studentName, groupName, 'Unexpected Time');
    }
 
    // Rule 4: AM/PM time should match AM/PM group if group name ends with AM/PM.
    if (groupName.endsWith(' AM') || groupName.endsWith(' PM')) {
        if (inTime && outTime) {
            const isGroupAM = groupName.endsWith(' AM');
            const isInTimeAM = inTime.includes('AM');
            const isOutTimeAM = outTime.includes('AM');

            if (isGroupAM !== isInTimeAM || isGroupAM !== isOutTimeAM) {
                throw new DataValidationError(`Time AM/PM and group AM/PM should match.`, lineNumber, studentName, groupName, 'Time/Group Mismatch');
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
                const { lineNumber, 'Student: Full Name': studentName, 'Program Day: Group: Class Name': groupName } = row;
                errors.push(new DataValidationError(
                    `An unexpected error occurred: ${error.message}`,
                    lineNumber,
                    studentName || 'Unknown',
                    groupName || 'Unknown',
                    'System Error'
                ));
            }
        }
    });

    return errors;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        validateCSVData,
        DataValidationError
    };
}
