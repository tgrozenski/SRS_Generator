// Import the functions to be tested
const { validateCSVData, convert_to_12_hour } = require('./dataValidation');

// Simple assertion utility
function assert(condition, message) {
    if (!condition) {
        throw new Error(message || "Assertion failed");
    }
}

// Simple test runner
function test(name, fn) {
    try {
        fn();
        console.log(`✓ ${name}`);
    } catch (error) {
        console.error(`✗ ${name}`);
        console.error(error);
    }
}

// --- Test Cases ---
test('should return no errors for valid data', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 09:00 AM","2025-01-01 05:00 PM","John Doe","Some Other Group"
"Scheduled","","","Jane Smith","Group B"
"Absent","","","Jack Black","Group C"
"Present","2025-01-01 09:00 AM","2025-01-01 11:00 AM","Joan Rivers","New Study Hall AM"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 0, `Expected 0 errors, but got ${errors.length}`);
});

test('Rule 1: should detect missing In Time for "Present" outcome', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","","2025-01-01 05:00 PM","John Doe","Group A"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Missing Time', `Expected 'Missing Time', got ${errors[0].errorType}`);
});

test('Rule 1: should detect missing Out Time for "Present" outcome', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 09:00 AM","","John Doe","Group A"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Missing Time', `Expected 'Missing Time', got ${errors[0].errorType}`);
});

test('Rule 1: should detect missing In and Out Time for "Present" outcome', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","","","John Doe","Group A"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Missing Time', `Expected 'Missing Time', got ${errors[0].errorType}`);
});

test('Rule 2: should detect In Time greater than Out Time', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 05:00 PM","2025-01-01 09:00 AM","John Doe","Group A"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Invalid Time Order', `Expected 'Invalid Time Order', got ${errors[0].errorType}`);
});

test('Rule 2: should detect Out Time less than 1 minute after In Time', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 09:00 AM","2025-01-01 09:00 AM","John Doe","Group A"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Invalid Time Order', `Expected 'Invalid Time Order', got ${errors[0].errorType}`);
});

test('Rule 2: should pass when Out Time is exactly 1 minute after In Time', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 09:00:00 AM","2025-01-01 09:01:00 AM","John Doe","Group A"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 0, `Expected 0 errors, got ${errors.length}`);
});

test('Rule 3: should detect time entries for "Scheduled" outcome', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Scheduled","2025-01-01 09:00 AM","2025-01-01 05:00 PM","John Doe","Group A"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Unexpected Time', `Expected 'Unexpected Time', got ${errors[0].errorType}`);
});

test('Rule 3: should detect time entries for "Absent" outcome', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Absent","2025-01-01 09:00 AM","2025-01-01 05:00 PM","John Doe","Group A"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Unexpected Time', `Expected 'Unexpected Time', got ${errors[0].errorType}`);
});

test('Rule 4: should detect AM/PM mismatch for AM group with PM times', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 01:00 PM","2025-01-01 05:00 PM","John Doe","Think Cafe AM"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Time/Group Mismatch', `Expected 'Time/Group Mismatch', got ${errors[0].errorType}`);
});

test('Rule 4: should detect AM/PM mismatch for PM group with AM times', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 09:00 AM","2025-01-01 11:00 AM","John Doe","Book Lab PM"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Time/Group Mismatch', `Expected 'Time/Group Mismatch', got ${errors[0].errorType}`);
});

test('Rule 4: should be resilient to new group names ending in AM/PM', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 02:00 PM","2025-01-01 04:00 PM","John Doe","New Group PM"
"Present","2025-01-01 09:00 AM","2025-01-01 11:00 AM","Jane Doe","Another Group AM"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 0, `Expected 0 errors, but got ${errors.length}`);
});

test('Rule 4: should detect mismatch with new group names', () => {
    const csvData = `"Outcome","In Time","Out Time","Student: Full Name","Program Day: Group: Class Name"
"Present","2025-01-01 09:00 AM","2025-01-01 11:00 AM","John Doe","New Group PM"`;
    const errors = validateCSVData(csvData);
    assert(errors.length === 1, `Expected 1 error, got ${errors.length}`);
    assert(errors[0].errorType === 'Time/Group Mismatch', `Expected 'Time/Group Mismatch', got ${errors[0].errorType}`);
});

// Testing actual CSV's here
const fs = require('fs');
test('Testing an actual CSV here', () => {
        const data = fs.readFileSync('testReports/Report.csv', 'utf8');
        const errors = validateCSVData(data);
        assert(errors.length === 3, `Expected 3 errors, got ${errors.length}`);
    }
)

test('Testing an actual CSV here', () => {
        const data = fs.readFileSync('testReports/SAR.csv', 'utf8');
        const errors = validateCSVData(data);
        assert(errors.length === 85, `Expected 7 errors, got ${errors.length}`);
    }
)

test('Testing an actual CSV here', () => {
        const data = fs.readFileSync('testReports/FULLSAR.csv', 'utf8');
        const errors = validateCSVData(data);
        assert(errors.length === 7, `Expected 7 errors, got ${errors.length}`);
    }
)