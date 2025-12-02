// Excel Processor - Memproses file Excel data presensi

import { calculateHours } from './dataFormatter.js';

// Process Excel file
export async function processExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Get all sheets
                const sheets = workbook.SheetNames;
                let allData = [];
                
                sheets.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    // Process data based on sheet structure
                    const processedData = processSheetData(jsonData, sheetName);
                    allData = [...allData, ...processedData];
                });
                
                if (allData.length === 0) {
                    reject(new Error('Tidak ada data yang ditemukan dalam file Excel'));
                    return;
                }
                
                resolve(allData);
                
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Process data from a specific sheet
function processSheetData(data, sheetName) {
    if (data.length === 0) return [];
    
    // Try different data formats
    let processedData = [];
    
    // Format 1: Direct columns (Nama, Tanggal, Jam Masuk, Jam Keluar)
    if (data[0].Nama || data[0].nama || data[0]['Nama Lengkap']) {
        processedData = data.map(row => ({
            nama: row.Nama || row.nama || row['Nama Lengkap'] || '',
            tanggal: row.Tanggal || row.tanggal || row.TANGGAL || row.Date || row.date || '',
            jamMasuk: row['Jam Masuk'] || row['jam masuk'] || row.JamMasuk || row.jamMasuk || row['Jam masuk'] || row['Check-in'] || '',
            jamKeluar: row['Jam Keluar'] || row['jam keluar'] || row.JamKeluar || row.jamKeluar || row['Jam keluar'] || row['Check-out'] || '',
            durasi: calculateHours(
                row['Jam Masuk'] || row['jam masuk'] || row.JamMasuk || row.jamMasuk || row['Jam masuk'] || row['Check-in'] || '',
                row['Jam Keluar'] || row['jam keluar'] || row.JamKeluar || row.jamKeluar || row['Jam keluar'] || row['Check-out'] || ''
            )
        }));
    }
    // Format 2: Column E (Nama) and Column F (Waktu) from your sample
    else if (data[0]['__EMPTY_4'] || data[0]['__EMPTY_5']) {
        processedData = data
            .filter(row => row['__EMPTY_4'] && row['__EMPTY_5'])
            .map(row => {
                const nama = row['__EMPTY_4'];
                const waktu = row['__EMPTY_5'];
                
                // Parse datetime string
                let tanggal = '';
                let jamMasuk = '';
                
                if (typeof waktu === 'string') {
                    const [datePart, timePart] = waktu.split(' ');
                    tanggal = datePart;
                    jamMasuk = timePart ? timePart.split(':').slice(0, 2).join(':') : '';
                } else if (waktu instanceof Date) {
                    tanggal = waktu.toISOString().split('T')[0];
                    jamMasuk = waktu.toTimeString().split(' ')[0].substring(0, 5);
                }
                
                return {
                    nama: nama.toString().trim(),
                    tanggal: tanggal,
                    jamMasuk: jamMasuk,
                    jamKeluar: '', // This format doesn't have out time
                    durasi: 0
                };
            });
    }
    
    // Filter out invalid entries
    return processedData.filter(item => 
        item.nama && 
        item.nama.trim() !== '' && 
        !item.nama.toLowerCase().includes('nama') && 
        !item.nama.toLowerCase().includes('name')
    );
}

// Calculate salaries from attendance data
export function calculateSalaries(data, salaryPerHour = 50000, overtimeRate = 75000, taxRate = 5, workHours = 8) {
    // Group data by employee
    const employees = {};
    
    data.forEach(record => {
        const name = record.nama.trim();
        
        if (!employees[name]) {
            employees[name] = {
                nama: name,
                records: [],
                totalHari: 0,
                totalJam: 0,
                jamNormal: 0,
                jamLembur: 0
            };
        }
        
        // Calculate hours worked
        const hoursWorked = record.durasi || calculateHours(record.jamMasuk, record.jamKeluar);
        
        if (hoursWorked > 0) {
            employees[name].records.push({
                tanggal: record.tanggal,
                jamMasuk: record.jamMasuk,
                jamKeluar: record.jamKeluar,
                durasi: hoursWorked
            });
            
            employees[name].totalHari++;
            employees[name].totalJam += hoursWorked;
            
            // Regular hours (max workHours per day)
            const regular = Math.min(hoursWorked, workHours);
            employees[name].jamNormal += regular;
            
            // Overtime hours (hours beyond workHours)
            const overtime = Math.max(hoursWorked - workHours, 0);
            employees[name].jamLembur += overtime;
        }
    });
    
    // Calculate salaries for each employee
    const result = Object.values(employees).map(emp => {
        const gajiPokok = emp.jamNormal * salaryPerHour;
        const uangLembur = emp.jamLembur * overtimeRate;
        const gajiKotor = gajiPokok + uangLembur;
        const pajak = gajiKotor * (taxRate / 100);
        const gajiBersih = gajiKotor - pajak;
        
        return {
            nama: emp.nama,
            totalHari: emp.totalHari,
            totalJam: emp.totalJam,
            jamNormal: emp.jamNormal,
            jamLembur: emp.jamLembur,
            gajiPokok: gajiPokok,
            uangLembur: uangLembur,
            pajak: pajak,
            gajiBersih: gajiBersih,
            // For detailed report
            records: emp.records
        };
    });
    
    // Sort by name
    result.sort((a, b) => a.nama.localeCompare(b.nama));
    
    return result;
}

// Parse complex Excel formats (for your specific file structure)
export function parseComplexExcel(data) {
    const result = [];
    
    // Your specific Excel has data in column E (index 4) and F (index 5)
    data.forEach((row, index) => {
        // Skip header rows
        if (index < 5) return;
        
        const nama = row[4]; // Column E
        const waktu = row[5]; // Column F
        
        if (nama && waktu) {
            // Parse the datetime string
            let tanggal = '';
            let jam = '';
            
            if (typeof waktu === 'string') {
                // Handle different date formats
                if (waktu.includes('/')) {
                    // Format: DD/MM/YYYY HH:MM
                    const [datePart, timePart] = waktu.split(' ');
                    tanggal = datePart ? datePart.split('/').reverse().join('-') : '';
                    jam = timePart || '';
                } else if (waktu.includes('-')) {
                    // Format: YYYY-MM-DD HH:MM:SS
                    const [datePart, timePart] = waktu.split(' ');
                    tanggal = datePart || '';
                    jam = timePart ? timePart.split(':').slice(0, 2).join(':') : '';
                }
            } else if (waktu instanceof Date) {
                // Excel date object
                tanggal = waktu.toISOString().split('T')[0];
                jam = waktu.toTimeString().split(' ')[0].substring(0, 5);
            }
            
            result.push({
                nama: nama.toString().trim(),
                tanggal: tanggal,
                jamMasuk: jam,
                jamKeluar: '', // Need to pair in/out times
                durasi: 0
            });
        }
    });
    
    // Group by name and date to pair in/out times
    return pairInOutTimes(result);
}

// Pair in and out times for each employee on each date
function pairInOutTimes(data) {
    const grouped = {};
    
    // Group by name and date
    data.forEach(record => {
        const key = `${record.nama}_${record.tanggal}`;
        if (!grouped[key]) {
            grouped[key] = [];
        }
        grouped[key].push(record.jamMasuk);
    });
    
    // Create pairs
    const result = [];
    Object.keys(grouped).forEach(key => {
        const [nama, tanggal] = key.split('_');
        const times = grouped[key].sort();
        
        if (times.length >= 2) {
            // Assume first is in, last is out
            result.push({
                nama: nama,
                tanggal: tanggal,
                jamMasuk: times[0],
                jamKeluar: times[times.length - 1],
                durasi: calculateHours(times[0], times[times.length - 1])
            });
        }
    });
    
    return result;
}
