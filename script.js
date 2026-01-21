// Main JavaScript file - Sistem Pengolahan Data Presensi

// Global variables
let originalData = [];
let processedData = [];
let currentFile = null;
let uploadProgressInterval = null;
let hoursChart = null;
let salaryChart = null;
let currentWorkHours = 8;

// ============================
// KONFIGURASI ATURAN BARU
// ============================

const WORK_START_TIME = '07:00';    // Jam mulai kerja efektif
const WORK_END_TIME = '16:00';      // Jam pulang normal
const MIN_OVERTIME_MINUTES = 10;    // Minimal lembur yang dihitung (menit)

const overtimeRates = {
    'TU': 12500,
    'STAFF': 10000,
    'K3': 8000
};

const employeeCategories = {
    // TU
    'Bu Ati': 'TU',
    'Pak Irvan': 'TU',
    // STAFF
    'Pak Ardhi': 'STAFF',
    'Windy': 'STAFF',
    'Bu Elzi': 'STAFF',
    'Intan': 'STAFF',
    'Pak Rafly': 'STAFF',
    'Erni': 'STAFF',
    'Bu Dian': 'STAFF',
    'Pebi': 'STAFF',
    'Bu Wahyu': 'STAFF',
    'Devi': 'STAFF',
    'Alifah': 'STAFF',
    // K3
    'Pak Saji': 'K3',
    'Pa Nanang': 'K3',
    'Pak Nanang': 'K3'
};

// Nonaktifkan source map warnings
if (typeof console !== 'undefined' && console.warn) {
    const originalWarn = console.warn;
    console.warn = function(...args) {
        if (args[0] && typeof args[0] === 'string' && args[0].includes('DevTools failed to load source map')) {
            return; // Abaikan warning source map
        }
        originalWarn.apply(console, args);
    };
}

// ============================
// FUNGSI HELPER BARU: GET DAY NAME
// ============================

// Fungsi untuk mendapatkan nama hari dari tanggal (format: DD/MM/YYYY)
function getDayNameFromDate(dateString) {
    if (!dateString) return '-';
    
    try {
        // Parse tanggal dari format DD/MM/YYYY
        const [day, month, year] = dateString.split('/').map(Number);
        
        // Validasi input
        if (!day || !month || !year) return '-';
        
        // Buat objek Date (bulan dimulai dari 0 = Januari)
        const date = new Date(year, month - 1, day);
        
        // Validasi jika tanggal tidak valid
        if (isNaN(date.getTime())) return '-';
        
        const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
        return days[date.getDay()];
    } catch (error) {
        console.error('Error getting day name:', error);
        return '-';
    }
}

// Fungsi untuk format tanggal dengan hari
function formatDateWithDay(dateString) {
    if (!dateString) return '-';
    
    const dayName = getDayNameFromDate(dateString);
    const formattedDate = formatDate(dateString);
    
    return `${dayName}, ${formattedDate}`;
}

// DOM Elements
const loadingScreen = document.getElementById('loading-screen');
const mainContainer = document.getElementById('main-container');
const excelFileInput = document.getElementById('excel-file');
const uploadArea = document.getElementById('upload-area');
const browseBtn = document.getElementById('browse-btn');
const filePreview = document.getElementById('file-preview');
const processBtn = document.getElementById('process-data');
const resultsSection = document.getElementById('results-section');
const cancelUploadBtn = document.getElementById('cancel-upload');

// Event Listeners
document.addEventListener('DOMContentLoaded', initializeApp);
excelFileInput.addEventListener('change', handleFileSelect);
browseBtn.addEventListener('click', () => excelFileInput.click());
uploadArea.addEventListener('click', () => excelFileInput.click());
processBtn.addEventListener('click', processData);
cancelUploadBtn.addEventListener('click', cancelUpload);

// ============================
// HELPER FUNCTIONS
// ============================

// Format hours from decimal to hours:minutes
function formatHoursToHMS(hours) {
    if (!hours || hours <= 0) return '0 jam';
    
    try {
        const totalMinutes = Math.round(hours * 60);
        const h = Math.floor(totalMinutes / 60);
        const m = totalMinutes % 60;
        
        if (m === 0) {
            return `${h} jam`;
        } else {
            return `${h} jam ${m} menit`;
        }
    } catch (error) {
        console.error('Error formatting hours:', error);
        return `${hours.toFixed(2)} jam`;
    }
}

// Format jam untuk display (format baru)
function formatHoursToDisplay(hours) {
    if (!hours || hours <= 0) return "0 jam";
    
    const totalMinutes = Math.round(hours * 60);
    const h = Math.floor(totalMinutes / 60);
    const m = totalMinutes % 60;
    
    if (m === 0) {
        return `${h} jam`;
    } else {
        return `${h} jam ${m} menit`;
    }
}

// Format date from DD/MM/YYYY to readable format
function formatDate(dateString) {
    if (!dateString) return '-';
    
    try {
        if (typeof dateString === 'string') {
            if (dateString.includes('/')) {
                const [day, month, year] = dateString.split('/');
                return `${day.padStart(2, '0')}/${month.padStart(2, '0')}/${year}`;
            }
        }
        return dateString;
    } catch (error) {
        return dateString;
    }
}

// Format time from "HH:MM" or "H:MM" to "HH:MM"
function formatTime(timeString) {
    if (!timeString) return '-';
    
    try {
        if (typeof timeString === 'string') {
            if (timeString.includes(':')) {
                const [hours, minutes] = timeString.split(':');
                return `${hours.padStart(2, '0')}:${minutes.padStart(2, '0')}`;
            }
        }
        return timeString;
    } catch (error) {
        return timeString;
    }
}

// Format currency
function formatCurrency(amount) {
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(amount);
}

// Parse datetime string "DD/MM/YYYY HH:MM" to separate date and time
function parseDateTime(datetimeStr) {
    if (!datetimeStr) return { date: '', time: '' };
    
    try {
        if (typeof datetimeStr === 'string') {
            const parts = datetimeStr.split(' ');
            if (parts.length >= 2) {
                return {
                    date: parts[0],
                    time: parts[1]
                };
            }
        } else if (datetimeStr instanceof Date) {
            return {
                date: datetimeStr.toISOString().split('T')[0],
                time: datetimeStr.toTimeString().split(' ')[0].substring(0, 5)
            };
        }
    } catch (error) {
        console.error('Error parsing datetime:', error);
    }
    
    return { date: '', time: '' };
}

// Calculate hours between two time strings DENGAN ATURAN BARU
function calculateHours(timeIn, timeOut) {
    if (!timeIn || !timeOut) return 0;
    
    try {
        const parseTime = (timeStr) => {
            if (!timeStr) return null;
            
            if (typeof timeStr === 'string') {
                if (timeStr.includes(':')) {
                    const [hours, minutes] = timeStr.split(':').map(Number);
                    return { hours, minutes: minutes || 0 };
                }
            }
            return null;
        };
        
        const inTime = parseTime(timeIn);
        const outTime = parseTime(timeOut);
        
        if (!inTime || !outTime) return 0;
        
        // Konversi ke menit
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        const outMinutes = outTime.hours * 60 + outTime.minutes;
        
        // Aturan 1: Jam masuk efektif mulai 7:00
        const workStartMinutes = 7 * 60; // 07:00 = 420 menit
        const effectiveInMinutes = Math.max(inMinutes, workStartMinutes);
        
        // Hitung total menit kerja
        let totalMinutes = outMinutes - effectiveInMinutes;
        
        // Jika pulang sebelum masuk (melewati tengah malam)
        if (totalMinutes < 0) {
            totalMinutes += 24 * 60;
        }
        
        return Math.round((totalMinutes / 60) * 100) / 100;
        
    } catch (error) {
        console.error('Error calculating hours:', error);
        return 0;
    }
}

// ============================
// FUNGSI PERHITUNGAN BARU
// ============================

// Fungsi untuk menghitung lembur dengan aturan baru
function calculateOvertimeWithNewRules(jamKeluar) {
    if (!jamKeluar) return 0;
    
    try {
        const parseTime = (timeStr) => {
            if (!timeStr) return null;
            if (typeof timeStr === 'string') {
                if (timeStr.includes(':')) {
                    const [hours, minutes] = timeStr.split(':').map(Number);
                    return { hours, minutes: minutes || 0 };
                }
            }
            return null;
        };
        
        const outTime = parseTime(jamKeluar);
        if (!outTime) return 0;
        
        const outMinutes = outTime.hours * 60 + outTime.minutes;
        const workEndMinutes = 16 * 60; // 16:00 = 960 menit
        
        // Jika pulang sebelum atau tepat jam 16:00, tidak ada lembur
        if (outMinutes <= workEndMinutes) return 0;
        
        // Hitung menit lembur
        let overtimeMinutes = outMinutes - workEndMinutes;
        
        // Aturan 3: Abaikan lembur kurang dari 10 menit
        if (overtimeMinutes < MIN_OVERTIME_MINUTES) {
            return 0;
        }
        
        // Konversi ke jam (desimal)
        const overtimeHours = Math.round((overtimeMinutes / 60) * 100) / 100;
        
        return overtimeHours;
        
    } catch (error) {
        console.error('Error calculating overtime:', error);
        return 0;
    }
}

// Fungsi untuk mendapatkan jam masuk efektif
function getEffectiveInTime(jamMasuk) {
    if (!jamMasuk) return '-';
    
    try {
        const parseTime = (timeStr) => {
            if (!timeStr) return null;
            if (typeof timeStr === 'string') {
                if (timeStr.includes(':')) {
                    const [hours, minutes] = timeStr.split(':').map(Number);
                    return { hours, minutes: minutes || 0 };
                }
            }
            return null;
        };
        
        const inTime = parseTime(jamMasuk);
        if (!inTime) return jamMasuk;
        
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        const workStartMinutes = 7 * 60; // 07:00
        
        // Jika masuk sebelum 7:00, gunakan 7:00
        if (inMinutes < workStartMinutes) {
            return '07:00';
        }
        
        // Jika masuk setelah 7:00, gunakan jam masuk sebenarnya
        return jamMasuk;
        
    } catch (error) {
        console.error('Error getting effective in time:', error);
        return jamMasuk;
    }
}

// Process Excel file
function processExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                const sheets = workbook.SheetNames;
                let allData = [];
                
                sheets.forEach(sheetName => {
                    const worksheet = workbook.Sheets[sheetName];
                    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    const processedData = processYourExcelFormat(rawData);
                    allData = [...allData, ...processedData];
                });
                
                if (allData.length === 0) {
                    reject(new Error('Tidak ada data yang ditemukan dalam file Excel'));
                    return;
                }
                
                const pairedData = pairInOutTimes(allData);
                resolve(pairedData);
                
            } catch (error) {
                console.error('Error processing file:', error);
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Process your specific Excel format (kolom E dan F)
function processYourExcelFormat(rawData) {
    const result = [];
    
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        
        if (row[4] && row[5]) {
            const nama = row[4];
            const waktu = row[5];
            
            const { date, time } = parseDateTime(waktu);
            
            if (nama && date && time) {
                result.push({
                    nama: nama.toString().trim(),
                    tanggal: date,
                    waktu: time,
                    rawDatetime: waktu
                });
        }
        }
    }
    
    return result;
}

// Pair in and out times for each employee on each date
function pairInOutTimes(data) {
    const grouped = {};
    
    data.forEach(record => {
        const key = `${record.nama}_${record.tanggal}`;
        if (!grouped[key]) {
            grouped[key] = [];
        }
        grouped[key].push({
            time: record.waktu,
            raw: record
        });
    });
    
    const result = [];
    
    Object.keys(grouped).forEach(key => {
        const [nama, tanggal] = key.split('_');
        const times = grouped[key];
        
        times.sort((a, b) => {
            const timeA = a.time.split(':').map(Number);
            const timeB = b.time.split(':').map(Number);
            return (timeA[0] * 60 + timeA[1]) - (timeB[0] * 60 + timeB[1]);
        });
        
        if (times.length >= 2) {
            const jamMasuk = times[0].time;
            const jamKeluar = times[times.length - 1].time;
            const durasi = calculateHours(jamMasuk, jamKeluar);
            
            result.push({
                nama: nama,
                tanggal: tanggal,
                jamMasuk: jamMasuk,
                jamKeluar: jamKeluar,
                durasi: durasi,
                jumlahCatatan: times.length
            });
        } else if (times.length === 1) {
            result.push({
                nama: nama,
                tanggal: tanggal,
                jamMasuk: times[0].time,
                jamKeluar: '',
                durasi: 0,
                jumlahCatatan: 1,
                keterangan: 'Hanya satu catatan'
            });
        }
    });
    
    return result;
}

// Calculate overtime per day - DENGAN ATURAN BARU
function calculateOvertimePerDay(data, workHours = 8) {
    const result = data.map(record => {
        const hoursWorked = record.durasi || calculateHours(record.jamMasuk, record.jamKeluar);
        
        // Hitung lembur dengan aturan baru (berdasarkan jam pulang)
        const jamLemburDesimal = calculateOvertimeWithNewRules(record.jamKeluar);
        
        // Jam normal: total jam dikurangi jam lembur, maksimal workHours
        const jamNormal = Math.min(hoursWorked - jamLemburDesimal, workHours);
        
        // Format untuk display
        let jamLemburDisplay = "0 jam";
        if (jamLemburDesimal > 0) {
            const totalMenitLembur = Math.round(jamLemburDesimal * 60);
            const jamLemburJam = Math.floor(totalMenitLembur / 60);
            const jamLemburMenit = totalMenitLembur % 60;
            
            if (jamLemburJam === 0) {
                jamLemburDisplay = `${jamLemburMenit} menit`;
            } else if (jamLemburMenit === 0) {
                jamLemburDisplay = `${jamLemburJam} jam`;
            } else {
                jamLemburDisplay = `${jamLemburJam} jam ${jamLemburMenit} menit`;
            }
        }
        
        const jamNormalDisplay = formatHoursToDisplay(jamNormal);
        const durasiDisplay = formatHoursToDisplay(hoursWorked);
        
        // Tentukan jam masuk efektif
        const effectiveInTime = getEffectiveInTime(record.jamMasuk);
        
        // Keterangan khusus
        let keterangan = 'Tidak lembur';
        if (jamLemburDesimal > 0) {
            keterangan = `Lembur ${jamLemburDisplay}`;
            if (effectiveInTime !== record.jamMasuk) {
                keterangan += ` (masuk efektif: ${effectiveInTime})`;
            }
        } else if (effectiveInTime !== record.jamMasuk) {
            keterangan = `Masuk efektif: ${effectiveInTime}`;
        }
        
        return {
            nama: record.nama,
            tanggal: record.tanggal,
            jamMasuk: record.jamMasuk,
            jamMasukEfektif: effectiveInTime,
            jamKeluar: record.jamKeluar,
            durasi: hoursWorked,
            durasiFormatted: durasiDisplay,
            jamNormal: jamNormal,
            jamNormalFormatted: jamNormalDisplay,
            jamLembur: jamLemburDisplay,
            jamLemburDesimal: jamLemburDesimal,
            keterangan: keterangan,
            jamKerjaNormal: workHours
        };
    });
    
    result.sort((a, b) => {
        if (a.nama === b.nama) {
            const dateA = a.tanggal.split('/').reverse().join('-');
            const dateB = b.tanggal.split('/').reverse().join('-');
            return new Date(dateA) - new Date(dateB);
        }
        return a.nama.localeCompare(b.nama);
    });
    
    return result;
}

// Hitung total jam lembur per karyawan
function calculateOvertimeSummary(data) {
    const summary = {};
    
    data.forEach(item => {
        const employeeName = item.nama;
        const overtimeHours = item.jamLemburDesimal || 0;
        
        if (!summary[employeeName]) {
            summary[employeeName] = {
                nama: employeeName,
                totalLembur: 0,
                kategori: employeeCategories[employeeName] || 'STAFF',
                rate: overtimeRates[employeeCategories[employeeName]] || 10000
            };
        }
        
        summary[employeeName].totalLembur += overtimeHours;
    });
    
    // Hitung total gaji
    Object.keys(summary).forEach(employee => {
        const record = summary[employee];
        record.totalGaji = record.totalLembur * record.rate;
        record.totalGajiFormatted = formatCurrency(record.totalGaji);
        record.totalLemburFormatted = formatHoursToDisplay(record.totalLembur);
    });
    
    return Object.values(summary);
}

// ============================
// FUNGSI DOWNLOAD LAPORAN GAJI LEMBUR PER KARYAWAN
// ============================

function downloadOvertimeSalaryReport() {
    if (processedData.length === 0) {
        showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
        return;
    }
    
    try {
        const summary = calculateOvertimeSummary(processedData);
        
        if (summary.length === 0) {
            showNotification('Tidak ada data lembur untuk diunduh.', 'info');
            return;
        }
        
        // Format data untuk Excel
        const exportData = summary.map((item, index) => ({
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Kategori': item.kategori,
            'Rate Lembur (per jam)': `Rp ${item.rate.toLocaleString('id-ID')}`,
            'Total Jam Lembur': item.totalLembur.toFixed(2),
            'Total Jam Lembur (Format)': item.totalLemburFormatted,
            'Total Gaji Lembur': `Rp ${Math.round(item.totalGaji).toLocaleString('id-ID')}`
        }));
        
        // Tambahkan total di baris terakhir
        const totalLembur = summary.reduce((sum, item) => sum + item.totalLembur, 0);
        const totalGaji = summary.reduce((sum, item) => sum + item.totalGaji, 0);
        
        exportData.push({
            'No': '',
            'Nama Karyawan': 'TOTAL',
            'Kategori': '',
            'Rate Lembur (per jam)': '',
            'Total Jam Lembur': totalLembur.toFixed(2),
            'Total Jam Lembur (Format)': formatHoursToDisplay(totalLembur),
            'Total Gaji Lembur': `Rp ${Math.round(totalGaji).toLocaleString('id-ID')}`
        });
        
        // Generate Excel
        const worksheet = XLSX.utils.json_to_sheet(exportData);
        
        // Set column widths
        const wscols = [
            { wch: 5 },    // No
            { wch: 20 },   // Nama Karyawan
            { wch: 10 },   // Kategori
            { wch: 18 },   // Rate Lembur
            { wch: 15 },   // Total Jam Lembur
            { wch: 20 },   // Total Jam Lembur (Format)
            { wch: 20 }    // Total Gaji Lembur
        ];
        worksheet['!cols'] = wscols;
        
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Rekap Gaji Lembur');
        
        // Nama file dengan tanggal
        const now = new Date();
        const dateStr = `${now.getDate().toString().padStart(2, '0')}-${(now.getMonth()+1).toString().padStart(2, '0')}-${now.getFullYear()}`;
        const filename = `Rekap_Gaji_Lembur_${dateStr}.xlsx`;
        
        XLSX.writeFile(workbook, filename);
        
        showNotification('Laporan gaji lembur berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error generating overtime salary report:', error);
        showNotification('Gagal mengunduh laporan gaji lembur.', 'error');
    }
}

// Fungsi alias untuk kompatibilitas
function downloadOverTimeSalaryReport() {
    return downloadOvertimeSalaryReport();
}

// ============================
// FUNGSI DISPLAY TABEL DENGAN HARI
// ============================

// Tampilkan data original DENGAN HARI
function displayOriginalTable(data) {
    const tbody = document.getElementById('original-table-body');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    data.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>
                <div style="font-weight: 500;">${formatDate(item.tanggal)}</div>
                <div style="font-size: 0.85rem; color: #3498db;">
                    <i class="fas fa-calendar-day"></i> ${getDayNameFromDate(item.tanggal)}
                </div>
            </td>
            <td>${item.jamMasuk}</td>
            <td>${item.jamKeluar || '-'}</td>
            <td>${item.durasi ? item.durasi.toFixed(2) + ' jam' : '-'}</td>
        `;
        tbody.appendChild(row);
    });
}

// Display processed table (DATA PER HARI) DENGAN HARI
function displayProcessedTable(data) {
    const tbody = document.getElementById('processed-table-body');
    if (!tbody) {
        console.error('Element #processed-table-body tidak ditemukan');
        return;
    }
    
    tbody.innerHTML = '';
    
    data.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>
                <div style="font-weight: 500;">${formatDate(item.tanggal)}</div>
                <div style="font-size: 0.85rem; color: #3498db;">
                    <i class="fas fa-calendar-day"></i> ${getDayNameFromDate(item.tanggal)}
                </div>
            </td>
            <td>
                ${item.jamMasuk}
                ${item.jamMasuk !== item.jamMasukEfektif ? 
                    `<br><small style="color: #666;">(efektif: ${item.jamMasukEfektif})</small>` : 
                    ''}
            </td>
            <td>${item.jamKeluar}</td>
            <td>${item.durasiFormatted}</td>
            <td>${item.jamNormalFormatted}</td>
            <td style="color: ${item.jamLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-weight: bold;">
                ${item.jamLembur}
            </td>
            <td>
                <span style="color: ${item.jamLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-size: 0.85rem;">
                    ${item.keterangan}
                </span>
            </td>
        `;
        tbody.appendChild(row);
    });
}

// Display summaries DENGAN ATURAN BARU
function displaySummaries(data) {
    const employeeSummary = document.getElementById('employee-summary');
    const financialSummary = document.getElementById('financial-summary');
    
    if (!employeeSummary && !financialSummary) return;
    
    // Group by employee
    const employeeGroups = {};
    data.forEach(item => {
        if (!employeeGroups[item.nama]) {
            employeeGroups[item.nama] = [];
        }
        employeeGroups[item.nama].push(item);
    });
    
    let employeeHtml = '';
    Object.keys(employeeGroups).forEach(employee => {
        const records = employeeGroups[employee];
        const totalHari = records.length;
        const totalJam = records.reduce((sum, item) => sum + item.durasi, 0);
        const totalLemburDesimal = records.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
        const hariLembur = records.filter(item => item.jamLemburDesimal > 0).length;
        const totalNormal = records.reduce((sum, item) => sum + item.jamNormal, 0);
        
        // Hitung hari dengan masuk sebelum 7:00
        const hariMasukSebelum7 = records.filter(item => {
            const jamMasuk = item.jamMasuk;
            if (!jamMasuk) return false;
            const [hours] = jamMasuk.split(':').map(Number);
            return hours < 7;
        }).length;
        
        employeeHtml += `
            <div style="margin-bottom: 1rem; padding-bottom: 1rem; border-bottom: 1px solid #eee;">
                <strong>${employee}</strong><br>
                <small>
                    Total Hari: ${totalHari} | 
                    Total Jam: ${formatHoursToDisplay(totalJam)}<br>
                    Jam Normal: ${formatHoursToDisplay(totalNormal)}<br>
                    <span style="color: ${totalLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-weight: bold;">
                        Jam Lembur: ${formatHoursToDisplay(totalLemburDesimal)} (${hariLembur} hari)
                    </span>
                    ${hariMasukSebelum7 > 0 ? `<br>Masuk sebelum 7:00: ${hariMasukSebelum7} hari` : ''}
                </small>
            </div>
        `;
    });
    
    if (employeeSummary) employeeSummary.innerHTML = employeeHtml;
    
    // Financial summary dengan perhitungan gaji
    const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
    const totalLemburDesimal = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    const totalNormal = data.reduce((sum, item) => sum + item.jamNormal, 0);
    const hariDenganLembur = data.filter(item => item.jamLemburDesimal > 0).length;
    
    // Hitung statistik tambahan
    const hariMasukSebelum7 = data.filter(item => {
        const jamMasuk = item.jamMasuk;
        if (!jamMasuk) return false;
        const [hours] = jamMasuk.split(':').map(Number);
        return hours < 7;
    }).length;
    
    const hariPulangSetelah16 = data.filter(item => {
        const jamKeluar = item.jamKeluar;
        if (!jamKeluar) return false;
        const [hours] = jamKeluar.split(':').map(Number);
        return hours > 16 || (hours === 16 && jamKeluar.split(':')[1] > 0);
    }).length;
    
    const hariLemburDiabaikan = data.filter(item => {
        const jamKeluar = item.jamKeluar;
        if (!jamKeluar) return false;
        
        const [hours, minutes] = jamKeluar.split(':').map(Number);
        const outMinutes = hours * 60 + (minutes || 0);
        const workEndMinutes = 16 * 60;
        
        // Pulang setelah 16:00 tapi kurang dari 16:10
        return outMinutes > workEndMinutes && 
               (outMinutes - workEndMinutes) < MIN_OVERTIME_MINUTES;
    }).length;
    
    // Hitung gaji berdasarkan kategori
    const summary = calculateOvertimeSummary(data);
    let salaryHtml = '';
    
    const byCategory = {
        'TU': { totalJam: 0, totalGaji: 0 },
        'STAFF': { totalJam: 0, totalGaji: 0 },
        'K3': { totalJam: 0, totalGaji: 0 }
    };
    
    summary.forEach(item => {
        const category = item.kategori;
        if (byCategory[category]) {
            byCategory[category].totalJam += item.totalLembur;
            byCategory[category].totalGaji += item.totalGaji;
        }
    });
    
    let totalGajiAll = 0;
    Object.keys(byCategory).forEach(category => {
        if (byCategory[category].totalJam > 0) {
            salaryHtml += `
                <div style="margin: 0.5rem 0;">
                    <strong>${category}:</strong> ${formatHoursToDisplay(byCategory[category].totalJam)} 
                    x Rp ${overtimeRates[category].toLocaleString('id-ID')}/jam = 
                    Rp ${Math.round(byCategory[category].totalGaji).toLocaleString('id-ID')}
                </div>
            `;
            totalGajiAll += byCategory[category].totalGaji;
        }
    });
    
    if (financialSummary) {
        financialSummary.innerHTML = `
            <div><strong>Aturan Perhitungan:</strong></div>
            <div style="font-size: 0.85rem; color: #666; margin-bottom: 1rem;">
                • Jam masuk efektif: 07:00<br>
                • Jam pulang normal: 16:00<br>
                • Minimal lembur: 10 menit
            </div>
            
            <div>Konfigurasi Jam Kerja: <strong>${currentWorkHours} jam/hari</strong></div>
            <div>Total Entri Data: <strong>${data.length} hari</strong></div>
            <div>Masuk sebelum 07:00: <strong>${hariMasukSebelum7} hari</strong></div>
            <div>Pulang setelah 16:00: <strong>${hariPulangSetelah16} hari</strong></div>
            <div>Lembur diabaikan (<10 menit): <strong>${hariLemburDiabaikan} hari</strong></div>
            <div>Hari dengan Lembur: <strong>${hariDenganLembur} hari</strong></div>
            <div>Total Jam Kerja: <strong>${formatHoursToDisplay(totalJam)}</strong></div>
            <div>Total Jam Normal: <strong>${formatHoursToDisplay(totalNormal)}</strong></div>
            <div style="color: ${totalLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-weight: bold;">
                Total Jam Lembur: <strong>${formatHoursToDisplay(totalLemburDesimal)}</strong>
            </div>
            <div style="border-top: 2px solid #3498db; padding-top: 0.5rem; margin-top: 0.5rem;">
                <strong>Perhitungan Gaji Lembur:</strong><br>
                ${salaryHtml}
                <div style="font-weight: bold; color: #2c3e50; margin-top: 0.5rem;">
                    TOTAL GAJI LEMBUR: Rp ${Math.round(totalGajiAll).toLocaleString('id-ID')}
                </div>
            </div>
        `;
    }
}

// Prepare data for export DENGAN HARI
function prepareExportData(data) {
    if (data.length === 0) return [];
    
    const hasOvertimeData = data[0].jamLembur !== undefined;
    
    if (hasOvertimeData) {
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Hari': getDayNameFromDate(item.tanggal),
            'Tanggal': formatDate(item.tanggal),
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi Kerja': item.durasiFormatted,
            'Jam Normal': item.jamNormalFormatted,
            'Jam Lembur': item.jamLembur,
            'Durasi (Desimal)': item.durasi.toFixed(2),
            'Lembur (Desimal)': item.jamLemburDesimal.toFixed(2),
            'Keterangan': item.keterangan
        }));
    } else {
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama': item.nama,
            'Hari': getDayNameFromDate(item.tanggal),
            'Tanggal': formatDate(item.tanggal),
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi': item.durasi ? item.durasiFormatted : '',
            'Keterangan': item.jamKeluar ? 
                `${item.jamMasuk} - ${item.jamKeluar} (${item.durasiFormatted})` : 
                'Hanya jam masuk'
        }));
    }
}

// ============================
// MAIN APPLICATION FUNCTIONS
// ============================

// Initialize application
function initializeApp() {
    const now = new Date();
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    document.getElementById('current-date').textContent = now.toLocaleDateString('id-ID', options);
    
    // Debug info
    console.log('=== SISTEM PRESENSI LEMBUR ===');
    console.log('XLSX available:', typeof XLSX !== 'undefined');
    console.log('Chart.js available:', typeof Chart !== 'undefined');
    
    // Inisialisasi tombol download gaji lembur
    initializeDownloadButtons();
    
    // Tab functionality
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            switchTab(tabId);
        });
    });
    
    // Modal functionality
    const helpBtn = document.getElementById('help-btn');
    const closeHelpBtns = document.querySelectorAll('#close-help, #close-help-btn');
    const helpModal = document.getElementById('help-modal');
    
    if (helpBtn) helpBtn.addEventListener('click', () => helpModal.classList.add('active'));
    closeHelpBtns.forEach(btn => {
        btn.addEventListener('click', () => helpModal.classList.remove('active'));
    });
    
    // Template download
    const templateBtn = document.getElementById('template-btn');
    if (templateBtn) templateBtn.addEventListener('click', downloadTemplate);
    
    // Reset config
    const resetBtn = document.getElementById('reset-config');
    if (resetBtn) resetBtn.addEventListener('click', resetConfig);
    
    // Download buttons
    const downloadOriginal = document.getElementById('download-original');
    const downloadProcessed = document.getElementById('download-processed');
    const downloadBoth = document.getElementById('download-both');
    
    if (downloadOriginal) downloadOriginal.addEventListener('click', () => downloadReport('original'));
    if (downloadProcessed) downloadProcessed.addEventListener('click', () => downloadReport('processed'));
    if (downloadBoth) downloadBoth.addEventListener('click', () => downloadReport('both'));
    
    // Loading screen
    setTimeout(() => {
        if (loadingScreen) {
            loadingScreen.style.opacity = '0';
            setTimeout(() => {
                loadingScreen.style.display = 'none';
                if (mainContainer) mainContainer.classList.add('loaded');
            }, 500);
        }
    }, 2000);
}

// Fungsi untuk menginisialisasi tombol download
function initializeDownloadButtons() {
    // Cari tombol dengan id 'download-salary'
    let downloadSalaryBtn = document.getElementById('download-salary');
    
    console.log('Mencari tombol download-salary:', downloadSalaryBtn);
    
    if (downloadSalaryBtn) {
        // Jika tombol ditemukan, tambahkan event listener
        downloadSalaryBtn.addEventListener('click', downloadOvertimeSalaryReport);
        console.log('Tombol download-salary ditemukan dan diinisialisasi');
    } else {
        // Jika tidak ditemukan, buat tombol secara otomatis
        console.log('Tombol download-salary tidak ditemukan, membuat tombol...');
        createDownloadButton();
    }
}

// Fungsi untuk membuat tombol download jika tidak ada
function createDownloadButton() {
    // Tunggu hingga DOM siap
    setTimeout(() => {
        // Cari tempat yang cocok untuk menempatkan tombol
        const resultsSection = document.getElementById('results-section');
        const downloadButtonsContainer = document.querySelector('.download-buttons') || 
                                        document.querySelector('.button-group') ||
                                        document.querySelector('.section-actions');
        
        let container;
        
        if (downloadButtonsContainer) {
            container = downloadButtonsContainer;
        } else if (resultsSection) {
            // Cari atau buat container untuk tombol di results section
            let btnContainer = resultsSection.querySelector('.download-container');
            if (!btnContainer) {
                btnContainer = document.createElement('div');
                btnContainer.className = 'download-container';
                btnContainer.style.cssText = `
                    display: flex;
                    flex-wrap: wrap;
                    gap: 10px;
                    margin: 20px 0;
                    justify-content: center;
                `;
                
                const firstChild = resultsSection.firstChild;
                if (firstChild) {
                    resultsSection.insertBefore(btnContainer, firstChild);
                } else {
                    resultsSection.appendChild(btnContainer);
                }
            }
            container = btnContainer;
        } else {
            // Buat floating button jika tidak ada container yang cocok
            createFloatingDownloadButton();
            return;
        }
        
        // Buat tombol download gaji lembur
        const button = document.createElement('button');
        button.id = 'download-salary';
        button.className = 'btn btn-primary';
        button.innerHTML = `
            <i class="fas fa-file-excel"></i> 
            <span>Download Rekap Gaji Lembur Per Orang</span>
        `;
        button.style.cssText = `
            background: linear-gradient(135deg, #27ae60, #2ecc71);
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 8px;
            font-weight: bold;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 8px;
            box-shadow: 0 4px 12px rgba(39, 174, 96, 0.3);
            transition: all 0.3s ease;
        `;
        
        // Hover effect
        button.addEventListener('mouseenter', function() {
            this.style.transform = 'translateY(-2px)';
            this.style.boxShadow = '0 6px 16px rgba(39, 174, 96, 0.4)';
        });
        
        button.addEventListener('mouseleave', function() {
            this.style.transform = 'translateY(0)';
            this.style.boxShadow = '0 4px 12px rgba(39, 174, 96, 0.3)';
        });
        
        // Tambahkan event listener
        button.addEventListener('click', downloadOvertimeSalaryReport);
        
        // Tambahkan ke container
        container.appendChild(button);
        
        console.log('Tombol download gaji lembur berhasil dibuat');
        
    }, 1000); // Tunggu 1 detik untuk memastikan DOM siap
}

// Fungsi untuk membuat floating button
function createFloatingDownloadButton() {
    const button = document.createElement('button');
    button.id = 'download-salary';
    button.className = 'floating-download-btn';
    button.innerHTML = `
        <i class="fas fa-file-excel"></i> 
        <span>Download Gaji Lembur</span>
    `;
    button.style.cssText = `
        position: fixed;
        bottom: 20px;
        right: 20px;
        background: linear-gradient(135deg, #e74c3c, #c0392b);
        color: white;
        border: none;
        padding: 12px 20px;
        border-radius: 50px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        cursor: pointer;
        display: flex;
        align-items: center;
        gap: 8px;
        z-index: 1000;
        font-weight: bold;
    `;
    
    // Hover effect untuk floating button
    button.addEventListener('mouseenter', function() {
        this.style.transform = 'scale(1.05)';
        this.style.boxShadow = '0 6px 20px rgba(0,0,0,0.3)';
    });
    
    button.addEventListener('mouseleave', function() {
        this.style.transform = 'scale(1)';
        this.style.boxShadow = '0 4px 15px rgba(0,0,0,0.2)';
    });
    
    button.addEventListener('click', downloadOvertimeSalaryReport);
    
    document.body.appendChild(button);
    
    console.log('Floating button download gaji lembur berhasil dibuat');
}

// Handle file selection
async function handleFileSelect(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    currentFile = file;
    showFilePreview(file);
    simulateUploadProgress();
    
    try {
        const data = await processExcelFile(file);
        originalData = data;
        
        updateSidebarStats(data);
        showNotification('File berhasil diunggah!', 'success');
        processBtn.disabled = false;
        
        previewUploadedData(data);
        
    } catch (error) {
        console.error('Error processing file:', error);
        showNotification('Gagal memproses file. Pastikan format sesuai.', 'error');
        cancelUpload();
    }
}

// Show preview of uploaded data
function previewUploadedData(data) {
    const previewDiv = document.createElement('div');
    previewDiv.className = 'data-preview';
    previewDiv.style.cssText = `
        margin-top: 1rem;
        padding: 1rem;
        background: #f8f9fa;
        border-radius: 8px;
        border-left: 4px solid #3498db;
    `;
    
    const previewCount = Math.min(data.length, 5);
    let previewHtml = `<h4 style="margin-bottom: 0.5rem; color: #2c3e50;">
        <i class="fas fa-eye"></i> Preview Data (${previewCount} dari ${data.length} entri)
    </h4>`;
    
    previewHtml += `<div style="font-size: 0.9rem; color: #555;">`;
    
    for (let i = 0; i < previewCount; i++) {
        const record = data[i];
        const dayName = getDayNameFromDate(record.tanggal);
        previewHtml += `
            <div style="margin-bottom: 0.5rem; padding-bottom: 0.5rem; border-bottom: 1px solid #eee;">
                <strong>${record.nama}</strong> - ${formatDate(record.tanggal)} (${dayName})<br>
                Masuk: ${record.jamMasuk} | Pulang: ${record.jamKeluar || '-'} 
                ${record.durasi ? `| Durasi: ${record.durasi.toFixed(2)} jam` : ''}
            </div>
        `;
    }
    
    previewHtml += `</div>`;
    
    previewDiv.innerHTML = previewHtml;
    
    const filePreview = document.getElementById('file-preview');
    const existingPreview = filePreview.querySelector('.data-preview');
    if (existingPreview) {
        existingPreview.remove();
    }
    filePreview.appendChild(previewDiv);
}

// Show file preview
function showFilePreview(file) {
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const fileDate = document.getElementById('file-date');
    
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileDate.textContent = new Date(file.lastModified).toLocaleDateString('id-ID');
    
    filePreview.style.display = 'block';
    uploadArea.style.display = 'none';
}

// Simulate upload progress
function simulateUploadProgress() {
    if (uploadProgressInterval) clearInterval(uploadProgressInterval);
    
    const progressBar = document.getElementById('upload-progress');
    const progressText = document.getElementById('progress-text');
    
    let progress = 0;
    uploadProgressInterval = setInterval(() => {
        progress += Math.random() * 15;
        if (progress >= 100) {
            progress = 100;
            clearInterval(uploadProgressInterval);
        }
        
        progressBar.style.width = `${progress}%`;
        progressText.textContent = `${Math.round(progress)}%`;
    }, 100);
}

// Process data (HITUNG LEMBUR SAJA)
function processData() {
    if (originalData.length === 0) {
        showNotification('Tidak ada data untuk diproses.', 'warning');
        return;
    }
    
    currentWorkHours = parseFloat(document.getElementById('work-hours').value) || 8;
    
    processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Menghitung Lembur...';
    processBtn.disabled = true;
    
    setTimeout(() => {
        try {
            // Hitung lembur per hari dengan jam kerja yang diatur
            processedData = calculateOvertimePerDay(originalData, currentWorkHours);
            
            displayResults(processedData);
            createCharts(processedData);
            
            resultsSection.style.display = 'block';
            resultsSection.scrollIntoView({ behavior: 'smooth' });
            
            // Pastikan tombol download tersedia
            ensureDownloadButtonAvailable();
            
            showNotification(`Perhitungan lembur selesai! (Jam kerja: ${currentWorkHours} jam)`, 'success');
            
        } catch (error) {
            console.error('Error processing data:', error);
            showNotification('Terjadi kesalahan saat menghitung lembur.', 'error');
        } finally {
            processBtn.innerHTML = '<i class="fas fa-calculator"></i> Hitung Lembur';
            processBtn.disabled = false;
        }
    }, 1500);
}

// Fungsi untuk memastikan tombol download tersedia
function ensureDownloadButtonAvailable() {
    // Cek apakah tombol sudah ada
    let btn = document.getElementById('download-salary');
    
    if (!btn && processedData.length > 0) {
        // Coba buat tombol lagi
        createDownloadButton();
        
        // Tampilkan notifikasi
        setTimeout(() => {
            showNotification('Tombol download gaji lembur per orang sekarang tersedia!', 'info');
        }, 500);
    }
}

// Display results
function displayResults(data) {
    updateMainStatistics(data);
    displayOriginalTable(originalData);
    displayProcessedTable(data);
    displaySummaries(data);
}

// Update main statistics
function updateMainStatistics(data) {
    const totalKaryawan = new Set(data.map(item => item.nama)).size;
    const totalHari = data.length;
    const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
    const totalLembur = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    
    const totalKaryawanElem = document.getElementById('total-karyawan');
    const totalHariElem = document.getElementById('total-hari');
    const totalLemburElem = document.getElementById('total-lembur');
    
    if (totalKaryawanElem) totalKaryawanElem.textContent = totalKaryawan;
    if (totalHariElem) totalHariElem.textContent = totalHari;
    if (totalLemburElem) totalLemburElem.textContent = formatHoursToDisplay(totalLembur);
    
    // Update total gaji menjadi jam lembur
    const totalGajiElem = document.getElementById('total-gaji');
    if (totalGajiElem) totalGajiElem.textContent = formatHoursToDisplay(totalLembur) + ' lembur';
}

// Create charts
function createCharts(data) {
    // Destroy existing charts
    if (hoursChart) hoursChart.destroy();
    if (salaryChart) salaryChart.destroy();
    
    // Group by employee untuk chart jam kerja
    const employeeGroups = {};
    data.forEach(item => {
        if (!employeeGroups[item.nama]) {
            employeeGroups[item.nama] = { normal: 0, lembur: 0 };
        }
        employeeGroups[item.nama].normal += item.jamNormal;
        employeeGroups[item.nama].lembur += item.jamLemburDesimal;
    });
    
    const employeeNames = Object.keys(employeeGroups).slice(0, 10);
    const regularHours = employeeNames.map(name => employeeGroups[name].normal);
    const overtimeHours = employeeNames.map(name => employeeGroups[name].lembur);
    
    // Chart Jam Kerja
    const hoursCtx = document.getElementById('hoursChart').getContext('2d');
    hoursChart = new Chart(hoursCtx, {
        type: 'bar',
        data: {
            labels: employeeNames,
            datasets: [
                {
                    label: 'Jam Normal',
                    data: regularHours,
                    backgroundColor: 'rgba(52, 152, 219, 0.7)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 1
                },
                {
                    label: 'Jam Lembur',
                    data: overtimeHours,
                    backgroundColor: 'rgba(231, 76, 60, 0.7)',
                    borderColor: 'rgba(231, 76, 60, 1)',
                    borderWidth: 1
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Distribusi Jam Kerja per Karyawan'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Jam Kerja'
                    },
                    ticks: {
                        callback: function(value) {
                            return value + ' jam';
                        }
                    }
                },
                x: {
                    title: {
                        display: true,
                        text: 'Karyawan'
                    }
                }
            }
        }
    });
    
    // Chart Pie untuk komposisi jam kerja
    const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
    const totalNormal = data.reduce((sum, item) => sum + item.jamNormal, 0);
    const totalLembur = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    
    const salaryCtx = document.getElementById('salaryChart').getContext('2d');
    salaryChart = new Chart(salaryCtx, {
        type: 'doughnut',
        data: {
            labels: ['Jam Normal', 'Jam Lembur'],
            datasets: [{
                data: [totalNormal, totalLembur],
                backgroundColor: [
                    'rgba(52, 152, 219, 0.8)',
                    'rgba(231, 76, 60, 0.8)'
                ],
                borderColor: [
                    'rgba(52, 152, 219, 1)',
                    'rgba(231, 76, 60, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                title: {
                    display: true,
                    text: 'Komposisi Jam Kerja'
                },
                legend: {
                    position: 'bottom'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const label = context.label || '';
                            const value = context.raw || 0;
                            const percentage = context.dataset.data[context.dataIndex] / totalJam * 100;
                            return `${label}: ${value.toFixed(2)} jam (${percentage.toFixed(1)}%)`;
                        }
                    }
                }
            }
        }
    });
}

// Switch tabs
function switchTab(tabId) {
    console.log('Switch to tab:', tabId);
    
    // Update tab buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('data-tab') === tabId) {
            btn.classList.add('active');
        }
    });
    
    // Update tab contents
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
        if (content.id === `${tabId}-tab`) {
            content.classList.add('active');
        }
    });
}

// Cancel upload
function cancelUpload() {
    if (uploadProgressInterval) {
        clearInterval(uploadProgressInterval);
        uploadProgressInterval = null;
    }
    
    excelFileInput.value = '';
    filePreview.style.display = 'none';
    uploadArea.style.display = 'block';
    processBtn.disabled = true;
    
    resultsSection.style.display = 'none';
}

// Reset configuration
function resetConfig() {
    document.getElementById('work-hours').value = '8';
    currentWorkHours = 8;
    
    showNotification('Konfigurasi telah direset ke 8 jam kerja.', 'info');
}

// Download template
function downloadTemplate() {
    const templateData = [
        {
            'Nama': 'Windy',
            'Tanggal': '01/11/2025',
            'Jam Masuk': '09:43',
            'Jam Keluar': '16:36'
        },
        {
            'Nama': 'Windy',
            'Tanggal': '02/11/2025',
            'Jam Masuk': '08:30',
            'Jam Keluar': '17:30'
        },
        {
            'Nama': 'Bu Ali',
            'Tanggal': '01/11/2025',
            'Jam Masuk': '08:00',
            'Jam Keluar': '17:00'
        }
    ];
    
    generateReport(templateData, 'template_data_presensi.xlsx', 'Template Presensi');
    showNotification('Template berhasil diunduh.', 'success');
}

// Get column widths
function getColumnWidths(data) {
    if (data.length === 0) return [];
    
    const firstRow = data[0];
    const columns = Object.keys(firstRow);
    
    return columns.map(col => {
        const maxLength = Math.max(
            col.length,
            ...data.map(row => (row[col] ? row[col].toString().length : 0))
        );
        
        return { wch: Math.min(Math.max(maxLength + 2, 10), 50) };
    });
}

// Download report
async function downloadReport(type) {
    if (type === 'original' && originalData.length === 0) {
        showNotification('Tidak ada data asli untuk diunduh.', 'warning');
        return;
    }
    
    if ((type === 'processed' || type === 'both') && processedData.length === 0) {
        showNotification('Data belum diproses.', 'warning');
        return;
    }
    
    try {
        if (type === 'original') {
            await generateReport(originalData, 'data_presensi_asli.xlsx', 'Data Asli');
            showNotification('Data asli berhasil diunduh.', 'success');
        } else if (type === 'processed') {
            await generateReport(processedData, 'data_lembur_harian.xlsx', 'Data Lembur Harian');
            showNotification('Data lembur harian berhasil diunduh.', 'success');
        } else if (type === 'both') {
            await generateReport(originalData, 'data_presensi_asli.xlsx', 'Data Asli');
            setTimeout(async () => {
                await generateReport(processedData, 'data_lembur_harian.xlsx', 'Data Lembur Harian');
                showNotification('Kedua file berhasil diunduh.', 'success');
            }, 500);
        }
    } catch (error) {
        console.error('Error downloading report:', error);
        showNotification('Gagal mengunduh laporan.', 'error');
    }
}

// Generate Excel report untuk data biasa
function generateReport(data, filename, sheetName = 'Data Lembur Harian') {
    try {
        const exportData = prepareExportData(data);
        const worksheet = XLSX.utils.json_to_sheet(exportData);
        
        const wscols = getColumnWidths(exportData);
        worksheet['!cols'] = wscols;
        
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        XLSX.writeFile(workbook, filename);
        
        return true;
    } catch (error) {
        console.error('Error generating report:', error);
        throw error;
    }
}

// Update sidebar stats
function updateSidebarStats(data) {
    const uniqueEmployees = new Set(data.map(item => item.nama)).size;
    const totalDays = new Set(data.map(item => item.tanggal)).size;
    
    document.getElementById('stat-employees').textContent = uniqueEmployees;
    document.getElementById('stat-attendance').textContent = totalDays;
}

// Format file size
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Show notification
function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 1rem 1.5rem;
        background: ${type === 'success' ? '#2ecc71' : type === 'error' ? '#e74c3c' : '#3498db'};
        color: white;
        border-radius: 8px;
        box-shadow: 0 5px 20px rgba(0, 0, 0, 0.15);
        z-index: 1000;
        animation: slideInRight 0.3s ease;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    `;
    
    const icon = type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : 'info-circle';
    notification.innerHTML = `<i class="fas fa-${icon}"></i> ${message}`;
    
    document.body.appendChild(notification);
    
    setTimeout(() => {
        notification.style.animation = 'slideOutRight 0.3s ease';
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 300);
    }, 3000);
}

// Add CSS animations for notifications
if (!document.querySelector('#notification-styles')) {
    const style = document.createElement('style');
    style.id = 'notification-styles';
    style.textContent = `
        @keyframes slideInRight {
            from {
                transform: translateX(100%);
                opacity: 0;
            }
            to {
                transform: translateX(0);
                opacity: 1;
            }
        }
        
        @keyframes slideOutRight {
            from {
                transform: translateX(0);
                opacity: 1;
            }
            to {
                transform: translateX(100%);
                opacity: 0;
            }
        }
    `;
    document.head.appendChild(style);
}

// Inisialisasi aplikasi saat DOM siap
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initializeApp);
} else {
    initializeApp();
}
