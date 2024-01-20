const fs = require('fs');
const xlsx = require('xlsx');

function parseTime(timeStr) {
    // Convert time string to Date object
    return new Date(timeStr);
}

function analyzeFile(filePath) {
    let prevTimeOut;
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    for (const row of data.slice(1)) {
        // Assuming the file structure is consistent with the provided sample
        const positionId = row[0];
        const positionStatus = row[1];
        const timeInStr = row[2];
        const timeOutStr = row[3];

        // Skip rows with no time information or inactive positions
        if (!timeInStr || !timeOutStr || positionStatus !== 'Active') {
            continue;
        }

        const timeIn = parseTime(timeInStr);
        const timeOut = parseTime(timeOutStr);
        const timecardHours = parseFloat(row[4].split(':')[0]);

        // Check conditions and print relevant information
        if ((timeOut - timeIn) / (1000 * 60 * 60 * 24) === 6) {
            console.log(`Employee: ${row[7]}, Position: ${positionId} worked for 7 consecutive days.`);
        }

        const timeBetweenShifts = (timeIn - prevTimeOut) / (1000 * 60 * 60) || 0;
        if (1 < timeBetweenShifts && timeBetweenShifts < 10) {
            console.log(`Employee: ${row[7]}, Position: ${positionId} has less than 10 hours between shifts but greater than 1 hour.`);
        }

        if (timecardHours > 14) {
            console.log(`Employee: ${row[7]}, Position: ${positionId} worked for more than 14 hours in a single shift.`);
        }

        prevTimeOut = timeOut;
    }
}

const filePath = './Assignment_Timecard.xlsx'; // Replace with the actual path to your Excel file
analyzeFile(filePath);