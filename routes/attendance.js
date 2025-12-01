const express = require('express');
const router = express.Router();
const XLSX = require('xlsx');
const path = require('path');

const DATA_FILE_PATH = path.join(__dirname, '../data/DATA ABSEN.xlsx');

// Helper function to read Excel file
function readExcelFile() {
    try {
        const workbook = XLSX.readFile(DATA_FILE_PATH);
        return workbook;
    } catch (error) {
        console.error('Error reading Excel file:', error);
        return null;
    }
}

// Helper function to write Excel file
function writeExcelFile(workbook) {
    try {
        XLSX.writeFile(workbook, DATA_FILE_PATH);
        return true;
    } catch (error) {
        console.error('Error writing Excel file:', error);
        return false;
    }
}

// Get all employees
router.get('/employees', (req, res) => {
    const workbook = readExcelFile();
    if (!workbook) {
        return res.status(500).json({ error: 'Cannot read data file' });
    }

    const sheet = workbook.Sheets['NOVEMBER'];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    // Extract employees from column B (index 1)
    const employees = [];
    for (let i = 1; i < data.length; i++) {
        if (data[i] && data[i][1] && data[i][1].trim()) {
            employees.push({ name: data[i][1].trim() });
        }
    }
    
    res.json(employees);
});

// Get all attendance data
router.get('/attendance', (req, res) => {
    const workbook = readExcelFile();
    if (!workbook) {
        return res.status(500).json({ error: 'Cannot read data file' });
    }

    const sheet = workbook.Sheets['NOVEMBER'];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    const attendanceData = [];
    
    // Process data from columns E and F (index 4 and 5)
    for (let i = 1; i < data.length; i++) {
        if (data[i] && data[i][4] && data[i][5]) {
            const name = data[i][4];
            const datetimeStr = data[i][5];
            
            if (name && datetimeStr) {
                // Parse datetime string
                let date, time;
                if (datetimeStr.includes(' ')) {
                    const [datePart, timePart] = datetimeStr.split(' ');
                    date = datePart.replace(/\//g, '-');
                    time = timePart.substring(0, 5); // HH:MM format
                } else {
                    // Handle different date formats
                    date = datetimeStr.split(' ')[0].replace(/\//g, '-');
                    time = '00:00';
                }
                
                // Find existing record for this name and date or create new one
                let record = attendanceData.find(r => r.name === name && r.date === date);
                if (!record) {
                    record = { name, date, timeIn: null, timeOut: null };
                    attendanceData.push(record);
                }
                
                // Determine if this is time in or time out based on time
                if (!record.timeIn || time < record.timeIn) {
                    record.timeIn = time;
                } else {
                    record.timeOut = time;
                }
            }
        }
    }
    
    res.json(attendanceData);
});

// Add new attendance
router.post('/attendance', (req, res) => {
    const { name, type, date, time } = req.body;
    
    if (!name || !type || !date || !time) {
        return res.status(400).json({ error: 'All fields are required' });
    }

    const workbook = readExcelFile();
    if (!workbook) {
        return res.status(500).json({ error: 'Cannot read data file' });
    }

    const sheet = workbook.Sheets['NOVEMBER'];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    // Find the next empty row
    let newRowIndex = data.length;
    for (let i = 1; i < data.length; i++) {
        if (!data[i] || data[i].length === 0) {
            newRowIndex = i;
            break;
        }
    }
    
    // Format datetime for Excel
    const formattedDate = date.split('-').reverse().join('/');
    const datetime = `${formattedDate} ${time}`;
    
    // Add new attendance record
    if (!data[newRowIndex]) {
        data[newRowIndex] = [];
    }
    data[newRowIndex][4] = name; // Column E
    data[newRowIndex][5] = datetime; // Column F
    
    // Convert back to sheet and save
    const newSheet = XLSX.utils.aoa_to_sheet(data);
    workbook.Sheets['NOVEMBER'] = newSheet;
    
    if (writeExcelFile(workbook)) {
        res.json({ message: 'Attendance saved successfully' });
    } else {
        res.status(500).json({ error: 'Failed to save attendance' });
    }
});

// Calculate overtime
router.get('/overtime', (req, res) => {
    const workbook = readExcelFile();
    if (!workbook) {
        return res.status(500).json({ error: 'Cannot read data file' });
    }

    const sheet = workbook.Sheets['NOVEMBER'];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    const overtimeData = [];
    const attendanceMap = new Map();
    
    // Group attendance by name and date
    for (let i = 1; i < data.length; i++) {
        if (data[i] && data[i][4] && data[i][5]) {
            const name = data[i][4];
            const datetimeStr = data[i][5];
            
            if (name && datetimeStr) {
                let date, time;
                if (datetimeStr.includes(' ')) {
                    const [datePart, timePart] = datetimeStr.split(' ');
                    date = datePart.replace(/\//g, '-');
                    time = timePart.substring(0, 5);
                } else {
                    date = datetimeStr.split(' ')[0].replace(/\//g, '-');
                    time = '00:00';
                }
                
                const key = `${name}-${date}`;
                if (!attendanceMap.has(key)) {
                    attendanceMap.set(key, { name, date, times: [] });
                }
                attendanceMap.get(key).times.push(time);
            }
        }
    }
    
    // Calculate overtime for each day
    for (const [key, record] of attendanceMap) {
        if (record.times.length >= 2) {
            record.times.sort();
            const timeIn = record.times[0];
            const timeOut = record.times[record.times.length - 1];
            
            const [inHours, inMinutes] = timeIn.split(':').map(Number);
            const [outHours, outMinutes] = timeOut.split(':').map(Number);
            
            const totalInMinutes = inHours * 60 + inMinutes;
            const totalOutMinutes = outHours * 60 + outMinutes;
            
            let durationMinutes = totalOutMinutes - totalInMinutes;
            if (durationMinutes < 0) {
                durationMinutes += 24 * 60; // Next day
            }
            
            const hoursWorked = durationMinutes / 60;
            const overtimeHours = Math.max(0, hoursWorked - 9); // Assuming 9 hours normal work
            
            if (overtimeHours > 0) {
                overtimeData.push({
                    name: record.name,
                    date: record.date,
                    timeIn,
                    timeOut,
                    overtimeHours: overtimeHours.toFixed(1)
                });
            }
        }
    }
    
    res.json(overtimeData);
});

// Delete attendance
router.delete('/attendance', (req, res) => {
    const { date, name } = req.body;
    
    if (!date || !name) {
        return res.status(400).json({ error: 'Date and name are required' });
    }

    const workbook = readExcelFile();
    if (!workbook) {
        return res.status(500).json({ error: 'Cannot read data file' });
    }

    const sheet = workbook.Sheets['NOVEMBER'];
    let data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    // Filter out the record to delete
    const formattedDate = date.split('-').reverse().join('/');
    data = data.filter((row, index) => {
        if (index === 0) return true; // Keep header
        if (!row || !row[4] || !row[5]) return true;
        
        const rowName = row[4];
        const rowDateTime = row[5];
        
        if (rowName === name && rowDateTime.includes(formattedDate)) {
            return false; // Remove this row
        }
        return true;
    });
    
    // Convert back to sheet and save
    const newSheet = XLSX.utils.aoa_to_sheet(data);
    workbook.Sheets['NOVEMBER'] = newSheet;
    
    if (writeExcelFile(workbook)) {
        res.json({ message: 'Attendance deleted successfully' });
    } else {
        res.status(500).json({ error: 'Failed to delete attendance' });
    }
});

module.exports = router;
