// Main JavaScript file - Sistem Pengolahan Data Presensi

// Global variables
let originalData = [];
let processedData = [];
let currentFile = null;
let uploadProgressInterval = null;
let hoursChart = null;
let salaryChart = null;

// ============================
// FUNGSI PERHITUNGAN GAJI LEMBUR
// ============================

// Kategori karyawan dan rate lembur
const overtimeRates = {
    'TU': 12500,
    'STAFF': 10000,
    'K3': 8000
};

// Mapping nama karyawan ke kategori
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

// Format currency
function formatCurrency(amount) {
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0,
        maximumFractionDigits: 0
    }).format(amount);
}

// Generate laporan gaji lembur
function generateOvertimeSalaryReport(data) {
    const summary = calculateOvertimeSummary(data);
    
    const reportData = summary.map((item, index) => ({
        'No': index + 1,
        'Nama Karyawan': item.nama,
        'Kategori': item.kategori,
        'Total Jam Lembur': item.totalLembur.toFixed(2),
        'Rate Lembur (per jam)': `Rp ${item.rate.toLocaleString('id-ID')}`,
        'Total Gaji Lembur': `Rp ${Math.round(item.totalGaji).toLocaleString('id-ID')}`,
        'Keterangan': `Lembur ${item.totalLemburFormatted} x Rp ${item.rate.toLocaleString('id-ID')}`
    }));
    
    // Tambahkan total keseluruhan
    const totalJamLembur = summary.reduce((sum, item) => sum + item.totalLembur, 0);
    const totalGajiLembur = summary.reduce((sum, item) => sum + item.totalGaji, 0);
    
    reportData.push({});
    reportData.push({
        'Nama Karyawan': 'TOTAL KESELURUHAN',
        'Total Jam Lembur': totalJamLembur.toFixed(2),
        'Total Gaji Lembur': `Rp ${Math.round(totalGajiLembur).toLocaleString('id-ID')}`
    });
    
    return reportData;
}

// Download laporan gaji lembur
function downloadOvertimeSalaryReport(data) {
    if (data.length === 0) {
        showNotification('Tidak ada data lembur untuk diunduh.', 'warning');
        return;
    }
    
    try {
        const reportData = generateOvertimeSalaryReport(data);
        const worksheet = XLSX.utils.json_to_sheet(reportData);
        
        const wscols = [
            { wch: 5 },   // No
            { wch: 25 },  // Nama Karyawan
            { wch: 15 },  // Kategori
            { wch: 15 },  // Total Jam Lembur
            { wch: 20 },  // Rate Lembur
            { wch: 25 },  // Total Gaji Lembur
            { wch: 30 }   // Keterangan
        ];
        worksheet['!cols'] = wscols;
        
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Rekap Gaji Lembur');
        
        const filename = `rekap_gaji_lembur_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(workbook, filename);
        
        showNotification('Laporan gaji lembur berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error generating salary report:', error);
        showNotification('Gagal mengunduh laporan gaji lembur.', 'error');
    }
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

// Calculate hours between two time strings
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
        
        let totalMinutes = (outTime.hours * 60 + outTime.minutes) - 
                          (inTime.hours * 60 + inTime.minutes);
        
        if (totalMinutes < 0) {
            totalMinutes += 24 * 60;
        }
        
        return Math.round((totalMinutes / 60) * 100) / 100;
        
    } catch (error) {
        console.error('Error calculating hours:', error);
        return 0;
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

// Calculate overtime per day - LEMBUR jika total jam > 8 jam
function calculateOvertimePerDay(data, workHours = 8) {
    const result = data.map(record => {
        const hoursWorked = record.durasi || calculateHours(record.jamMasuk, record.jamKeluar);
        
        // Jam normal maksimal workHours (8 jam)
        const jamNormal = Math.min(hoursWorked, workHours);
        
        // Jam lembur jika total jam > workHours (desimal)
        const jamLemburDesimal = Math.max(hoursWorked - workHours, 0);
        
        // Format jam lembur untuk display
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
        
        // Format jam normal untuk display
        const jamNormalDisplay = formatHoursToDisplay(jamNormal);
        
        // Format durasi untuk display
        const durasiDisplay = formatHoursToDisplay(hoursWorked);
        
        // Keterangan
        const keterangan = jamLemburDesimal > 0 ? `Lembur ${jamLemburDisplay}` : 'Tidak lembur';
        
        return {
            nama: record.nama,
            tanggal: record.tanggal,
            jamMasuk: record.jamMasuk,
            jamKeluar: record.jamKeluar,
            durasi: hoursWorked,
            durasiFormatted: durasiDisplay,
            jamNormal: jamNormal,
            jamNormalFormatted: jamNormalDisplay,
            jamLembur: jamLemburDisplay,
            jamLemburDesimal: jamLemburDesimal,
            keterangan: keterangan
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

// Generate Excel report
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

// Prepare data for export
function prepareExportData(data) {
    if (data.length === 0) return [];
    
    const hasOvertimeData = data[0].jamLembur !== undefined;
    
    if (hasOvertimeData) {
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama Karyawan': item.nama,
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

// ============================
// MAIN APPLICATION FUNCTIONS
// ============================

// Initialize application
function initializeApp() {
    const now = new Date();
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    document.getElementById('current-date').textContent = now.toLocaleDateString('id-ID', options);
    
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
        previewHtml += `
            <div style="margin-bottom: 0.5rem; padding-bottom: 0.5rem; border-bottom: 1px solid #eee;">
                <strong>${record.nama}</strong> - ${formatDate(record.tanggal)}<br>
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
    
    const workHours = parseFloat(document.getElementById('work-hours').value) || 8;
    
    processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Menghitung Lembur...';
    processBtn.disabled = true;
    
    setTimeout(() => {
        try {
            // Hitung lembur per hari
            processedData = calculateOvertimePerDay(originalData, workHours);
            
            displayResults(processedData);
            createCharts(processedData);
            
            resultsSection.style.display = 'block';
            resultsSection.scrollIntoView({ behavior: 'smooth' });
            
            showNotification('Perhitungan lembur selesai!', 'success');
            
        } catch (error) {
            console.error('Error processing data:', error);
            showNotification('Terjadi kesalahan saat menghitung lembur.', 'error');
        } finally {
            processBtn.innerHTML = '<i class="fas fa-calculator"></i> Hitung Lembur';
            processBtn.disabled = false;
        }
    }, 1500);
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

// Tampilkan data original
function displayOriginalTable(data) {
    const tbody = document.getElementById('original-table-body');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    data.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>${formatDate(item.tanggal)}</td>
            <td>${item.jamMasuk}</td>
            <td>${item.jamKeluar || '-'}</td>
            <td>${item.durasi ? item.durasi.toFixed(2) + ' jam' : '-'}</td>
        `;
        tbody.appendChild(row);
    });
}

// Display processed table (DATA PER HARI)
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
            <td>${formatDate(item.tanggal)}</td>
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

// Display summaries
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
                </small>
            </div>
        `;
    });
    
    if (employeeSummary) employeeSummary.innerHTML = employeeHtml;
    
    // Financial summary
    const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
    const totalLemburDesimal = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    const totalNormal = data.reduce((sum, item) => sum + item.jamNormal, 0);
    const hariDenganLembur = data.filter(item => item.jamLemburDesimal > 0).length;
    
    if (financialSummary) {
        financialSummary.innerHTML = `
            <div>Total Entri Data: <strong>${data.length} hari</strong></div>
            <div>Hari dengan Lembur: <strong>${hariDenganLembur} hari</strong></div>
            <div>Total Jam Kerja: <strong>${formatHoursToDisplay(totalJam)}</strong></div>
            <div>Total Jam Normal: <strong>${formatHoursToDisplay(totalNormal)}</strong></div>
            <div style="color: ${totalLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-weight: bold;">
                Total Jam Lembur: <strong>${formatHoursToDisplay(totalLemburDesimal)}</strong>
            </div>
            <div style="border-top: 2px solid #3498db; padding-top: 0.5rem; margin-top: 0.5rem;">
                Rata-rata Jam Kerja per Hari: <strong>${(totalJam / data.length).toFixed(2)} jam</strong><br>
                Rata-rata Jam Lembur per Hari: <strong>${(totalLemburDesimal / data.length).toFixed(2)} jam</strong>
            </div>
        `;
    }
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
    
    showNotification('Konfigurasi telah direset.', 'info');
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
