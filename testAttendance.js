const Papa = require('papaparse');
const fs = require('fs');

// Replicate the date parsing logic from generateReports.js
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

// Read and process CSV
const csvData = fs.readFileSync('testReports/SAR.csv', 'utf8');
const results = Papa.parse(csvData, { header: true, skipEmptyLines: true });

const groupNameKey = 'Program Day: Group: Class Name';
const uniqueGroups = [...new Set(results.data.map(row => row[groupNameKey]).filter(Boolean))];

console.log(`Found ${uniqueGroups.length} groups`);
console.log('Groups:', uniqueGroups.map(g => g.split(' - ').pop().trim()));

// Process first group
const firstGroup = uniqueGroups[0];
console.log(`\nProcessing group: ${firstGroup}`);
console.log(`Clean name: ${firstGroup.includes(' - ') ? firstGroup.split(' - ').pop().trim() : firstGroup}`);

const attendanceData = {};
const presentRows = results.data.filter(row => row['Outcome'] === 'Present' && row[groupNameKey] === firstGroup);

console.log(`Total present rows: ${presentRows.length}`);

presentRows.forEach((row, idx) => {
    if (idx < 5) { // Show first 5 rows
        const studentName = reverseName(row['Student: Full Name'] || row['Student: Last, First']);
        const sessionDate = row['Session Date'];
        const dayOfWeek = getDayOfWeek(sessionDate);
        const grade = row['Grade'] || '';
        
        console.log(`Row ${idx}: ${studentName}, Date: ${sessionDate}, Day: ${dayOfWeek}, Grade: ${grade}`);
    }
    
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

console.log(`\nTotal unique students in group: ${Object.keys(attendanceData).length}`);

// Show sample students with their days
const sampleStudents = Object.keys(attendanceData).slice(0, 10);
console.log('\nSample students and their attendance days:');
sampleStudents.forEach(student => {
    const info = attendanceData[student];
    console.log(`  ${student} (Grade: ${info.grade}): ${Array.from(info.days).join(', ')}`);
});

// Verify day mapping
console.log('\nDay to column mapping verification:');
const dayToColumns = {
    'Monday': ['C', 'D', 'E'],
    'Tuesday': ['F', 'G', 'H'],
    'Wednesday': ['I', 'J', 'K'],
    'Thursday': ['L', 'M', 'N'],
    'Friday': ['O', 'P', 'Q']
};

// Check which days appear in data
const allDays = new Set();
Object.values(attendanceData).forEach(info => {
    info.days.forEach(day => allDays.add(day));
});
console.log('Days present in data:', Array.from(allDays).sort());

// Show mapping for each day
Array.from(allDays).sort().forEach(day => {
    const columns = dayToColumns[day];
    console.log(`  ${day} → columns ${columns.join(', ')}`);
});

// Count students per day
const dayCounts = {};
Object.values(attendanceData).forEach(info => {
    info.days.forEach(day => {
        dayCounts[day] = (dayCounts[day] || 0) + 1;
    });
});
console.log('\nStudents per day:');
Object.entries(dayCounts).forEach(([day, count]) => {
    console.log(`  ${day}: ${count} students`);
});