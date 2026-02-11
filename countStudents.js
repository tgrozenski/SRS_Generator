const fs = require('fs');
const Papa = require('papaparse');

const csvData = fs.readFileSync('testReports/SAR.csv', 'utf8');
const results = Papa.parse(csvData, { header: true, skipEmptyLines: true });

const groups = {};
results.data.forEach(row => {
    const group = row['Program Day: Group: Class Name'];
    const student = row['Student: Full Name'];
    if (group && student) {
        if (!groups[group]) groups[group] = new Set();
        groups[group].add(student);
    }
});

console.log('Unique students per group:');
Object.keys(groups).forEach(group => {
    console.log(`${group}: ${groups[group].size} students`);
});