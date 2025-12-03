// Main JavaScript file - Sistem Pengolahan Data Presensi & Gaji

// Global variables
let originalData = [];
let processedData = [];
let salaryData = [];
let currentFile = null;
let uploadProgressInterval = null;
let hoursChart = null;
let salaryChart = null;

// Configuration defaults
const DEFAULT_CONFIG = {
    workHours: 8,
    salaryPerHour: 50000,
    overtimeMultiplier: 1.5,
    taxPercentage: 5
};

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

// Format currency to Indonesian Rupiah
function formatRupiah(amount) {
    if (amount === 0) return 'Rp 0';
    
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0
    }).format(amount);
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
                    const processedData = processExcelFormat(rawData);
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

// Process Excel format
function processExcelFormat(rawData) {
    const result = [];
    
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        
        // Try different column configurations
        if (row[4] && row[5]) {
            // Format: kolom E dan F
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
        } else if (row[0] && row[1]) {
            // Format: kolom A dan B
            const nama = row[0];
            const waktu = row[1];
            const { date, time } = parseDateTime(waktu);
            
            if (nama && date && time) {
                result.push({
                    nama: nama.toString().trim(),
                    tanggal: date,
                    waktu: time,
                    rawDatetime: waktu
                });
            }
        } else if (row[0] && row[2] && row[3]) {
            // Format: Nama, Tanggal, Jam Masuk, Jam Keluar
            const nama = row[0];
            const tanggal = row[1];
            const jamMasuk = row[2];
            const jamKeluar = row[3];
            
            if (nama && tanggal) {
                result.push({
                    nama: nama.toString().trim(),
                    tanggal: tanggal,
                    waktu: jamMasuk || '',
                    tipe: 'masuk'
                });
                
                if (jamKeluar) {
                    result.push({
                        nama: nama.toString().trim(),
                        tanggal: tanggal,
                        waktu: jamKeluar,
                        tipe: 'keluar'
                    });
                }
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

// Calculate overtime and salary per day
function calculateOvertimeAndSalary(data, config) {
    const result = data.map(record => {
        const hoursWorked = record.durasi || calculateHours(record.jamMasuk, record.jamKeluar);
        
        // Calculate normal and overtime hours
        const jamNormal = Math.min(hoursWorked, config.workHours);
        const jamLembur = Math.max(hoursWorked - config.workHours, 0);
        
        // Calculate earnings
        const gajiNormal = jamNormal * config.salaryPerHour;
        const gajiLembur = jamLembur * config.salaryPerHour * config.overtimeMultiplier;
        const totalKotor = gajiNormal + gajiLembur;
        const pajak = totalKotor * (config.taxPercentage / 100);
        const gajiBersih = totalKotor - pajak;
        
        return {
            nama: record.nama,
            tanggal: record.tanggal,
            jamMasuk: record.jamMasuk,
            jamKeluar: record.jamKeluar,
            durasi: hoursWorked,
            jamNormal: jamNormal,
            jamLembur: jamLembur,
            gajiNormal: gajiNormal,
            gajiLembur: gajiLembur,
            totalKotor: totalKotor,
            pajak: pajak,
            gajiBersih: gajiBersih,
            keterangan: jamLembur > 0 ? `Lembur ${jamLembur.toFixed(2)} jam` : 'Tidak lembur'
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

// Calculate employee summary (aggregated per employee)
function calculateEmployeeSummary(dailyData, config) {
    const employeeGroups = {};
    
    dailyData.forEach(record => {
        if (!employeeGroups[record.nama]) {
            employeeGroups[record.nama] = {
                nama: record.nama,
                totalHari: 0,
                totalJam: 0,
                totalJamNormal: 0,
                totalJamLembur: 0,
                totalGajiNormal: 0,
                totalGajiLembur: 0,
                totalPajak: 0,
                totalGajiBersih: 0
            };
        }
        
        const employee = employeeGroups[record.nama];
        employee.totalHari++;
        employee.totalJam += record.durasi;
        employee.totalJamNormal += record.jamNormal;
        employee.totalJamLembur += record.jamLembur;
        employee.totalGajiNormal += record.gajiNormal;
        employee.totalGajiLembur += record.gajiLembur;
        employee.totalPajak += record.pajak;
        employee.totalGajiBersih += record.gajiBersih;
    });
    
    // Convert to array and sort
    const result = Object.values(employeeGroups);
    result.sort((a, b) => a.nama.localeCompare(b.nama));
    
    return result;
}

// Generate Excel report
function generateReport(data, filename, sheetName = 'Data') {
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
    
    // Check if it's daily data or employee summary
    const isDailyData = data[0].tanggal !== undefined;
    
    if (isDailyData) {
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi (jam)': item.durasi.toFixed(2),
            'Jam Normal': item.jamNormal.toFixed(2),
            'Jam Lembur': item.jamLembur.toFixed(2),
            'Gaji Normal': formatRupiah(item.gajiNormal),
            'Gaji Lembur': formatRupiah(item.gajiLembur),
            'Total Kotor': formatRupiah(item.totalKotor),
            'Pajak': formatRupiah(item.pajak),
            'Gaji Bersih': formatRupiah(item.gajiBersih),
            'Keterangan': item.keterangan
        }));
    } else {
        // Employee summary
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Total Hari': item.totalHari,
            'Total Jam Kerja': item.totalJam.toFixed(2),
            'Total Jam Normal': item.totalJamNormal.toFixed(2),
            'Total Jam Lembur': item.totalJamLembur.toFixed(2),
            'Total Gaji Normal': formatRupiah(item.totalGajiNormal),
            'Total Gaji Lembur': formatRupiah(item.totalGajiLembur),
            'Total Pajak': formatRupiah(item.totalPajak),
            'Total Gaji Bersih': formatRupiah(item.totalGajiBersih)
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
    // Set current date
    const now = new Date();
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    document.getElementById('current-date').textContent = now.toLocaleDateString('id-ID', options);
    
    // Setup tab navigation
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            switchTab(tabId);
        });
    });
    
    // Setup help modal
    const helpBtn = document.getElementById('help-btn');
    const closeHelpBtns = document.querySelectorAll('#close-help, #close-help-btn');
    const helpModal = document.getElementById('help-modal');
    
    helpBtn.addEventListener('click', () => helpModal.classList.add('active'));
    closeHelpBtns.forEach(btn => {
        btn.addEventListener('click', () => helpModal.classList.remove('active'));
    });
    
    // Setup template download
    const templateBtn = document.getElementById('template-btn');
    templateBtn.addEventListener('click', downloadTemplate);
    
    // Setup reset config
    const resetBtn = document.getElementById('reset-config');
    resetBtn.addEventListener('click', resetConfig);
    
    // Setup download buttons
    document.getElementById('download-original').addEventListener('click', () => downloadReport('original'));
    document.getElementById('download-daily').addEventListener('click', () => downloadReport('daily'));
    document.getElementById('download-employee').addEventListener('click', () => downloadReport('employee'));
    
    // Setup navigation
    setupNavigation();
    
    // Initialize with loading animation
    setTimeout(() => {
        loadingScreen.style.opacity = '0';
        setTimeout(() => {
            loadingScreen.style.display = 'none';
            mainContainer.classList.add('loaded');
        }, 500);
    }, 2000);
}

// Setup navigation menu
function setupNavigation() {
    // Header navigation
    document.querySelectorAll('.header-nav a').forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            const target = this.getAttribute('id');
            
            // Update active state
            document.querySelectorAll('.header-nav li').forEach(li => li.classList.remove('active'));
            this.parentElement.classList.add('active');
            
            // Show corresponding section
            switch(target) {
                case 'nav-upload':
                    showSection('upload-section');
                    break;
                case 'nav-analysis':
                    showAnalysis();
                    break;
                case 'nav-report':
                    showSection('results-section');
                    break;
                case 'nav-settings':
                    showSection('config-section');
                    break;
                case 'nav-help':
                    document.getElementById('help-modal').classList.add('active');
                    break;
                default:
                    showSection('upload-section');
            }
        });
    });
    
    // Sidebar menu
    document.querySelectorAll('.menu-item').forEach(item => {
        item.addEventListener('click', function(e) {
            e.preventDefault();
            
            // Update active state
            document.querySelectorAll('.menu-item').forEach(mi => mi.classList.remove('active'));
            this.classList.add('active');
            
            const menuId = this.getAttribute('id');
            switch(menuId) {
                case 'menu-dashboard':
                    showSection('upload-section');
                    break;
                case 'menu-import':
                    showSection('upload-section');
                    break;
                case 'menu-calculate':
                    processData();
                    break;
                case 'menu-analysis':
                    showAnalysis();
                    break;
                case 'menu-history':
                    // For now, show results
                    if (processedData.length > 0) {
                        showSection('results-section');
                    } else {
                        showNotification('Harap proses data terlebih dahulu', 'warning');
                    }
                    break;
                case 'menu-export':
                    if (processedData.length > 0) {
                        downloadReport('daily');
                    } else {
                        showNotification('Tidak ada data untuk diekspor', 'warning');
                    }
                    break;
            }
        });
    });
}

// Show specific section
function showSection(sectionId) {
    // Hide all main sections
    document.getElementById('upload-section').style.display = 'none';
    document.getElementById('config-section').style.display = 'none';
    document.getElementById('results-section').style.display = 'none';
    document.getElementById('analysis-section').style.display = 'none';
    
    // Show selected section
    document.getElementById(sectionId).style.display = 'block';
    
    // Scroll to top
    window.scrollTo({ top: 0, behavior: 'smooth' });
}

// Show analysis section
function showAnalysis() {
    if (processedData.length === 0) {
        showNotification('Harap proses data terlebih dahulu', 'warning');
        return;
    }
    
    showSection('analysis-section');
    displayAnalysisData();
}

// Display analysis data
function displayAnalysisData() {
    const analysisGrid = document.querySelector('.analysis-grid');
    if (!analysisGrid) return;
    
    // Calculate analysis metrics
    const totalEmployees = new Set(processedData.map(item => item.nama)).size;
    const totalDays = processedData.length;
    const totalHours = processedData.reduce((sum, item) => sum + item.durasi, 0);
    const totalOvertime = processedData.reduce((sum, item) => sum + item.jamLembur, 0);
    const avgHoursPerDay = totalHours / totalDays;
    const avgOvertimePerDay = totalOvertime / totalDays;
    
    const totalSalary = processedData.reduce((sum, item) => sum + item.gajiBersih, 0);
    const avgSalaryPerEmployee = totalSalary / totalEmployees;
    const totalTax = processedData.reduce((sum, item) => sum + item.pajak, 0);
    
    const analysisData = [
        {
            title: 'Rata-rata Jam Kerja/Hari',
            value: avgHoursPerDay.toFixed(2) + ' jam',
            description: 'Durasi kerja rata-rata per hari'
        },
        {
            title: 'Rata-rata Lembur/Hari',
            value: avgOvertimePerDay.toFixed(2) + ' jam',
            description: 'Jam lembur rata-rata per hari'
        },
        {
            title: 'Rata-rata Gaji/Karyawan',
            value: formatRupiah(avgSalaryPerEmployee),
            description: 'Gaji bersih rata-rata per karyawan'
        },
        {
            title: 'Total Pajak',
            value: formatRupiah(totalTax),
            description: 'Total pajak yang dipotong'
        },
        {
            title: 'Hari dengan Lembur',
            value: processedData.filter(item => item.jamLembur > 0).length + ' hari',
            description: 'Jumlah hari terjadi lembur'
        },
        {
            title: 'Efisiensi Kerja',
            value: ((totalHours / (totalDays * 8)) * 100).toFixed(1) + '%',
            description: 'Persentase pemanfaatan jam kerja'
        }
    ];
    
    analysisGrid.innerHTML = analysisData.map(item => `
        <div class="analysis-card">
            <h4><i class="fas fa-chart-line"></i> ${item.title}</h4>
            <div class="analysis-value">${item.value}</div>
            <div class="analysis-description">${item.description}</div>
        </div>
    `).join('');
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

// Process data (calculate overtime and salary)
function processData() {
    if (originalData.length === 0) {
        showNotification('Tidak ada data untuk diproses.', 'warning');
        return;
    }
    
    // Get configuration values
    const config = {
        workHours: parseFloat(document.getElementById('work-hours').value) || DEFAULT_CONFIG.workHours,
        salaryPerHour: parseFloat(document.getElementById('salary-per-hour').value) || DEFAULT_CONFIG.salaryPerHour,
        overtimeMultiplier: parseFloat(document.getElementById('overtime-multiplier').value) || DEFAULT_CONFIG.overtimeMultiplier,
        taxPercentage: parseFloat(document.getElementById('tax-percentage').value) || DEFAULT_CONFIG.taxPercentage
    };
    
    processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Menghitung Gaji...';
    processBtn.disabled = true;
    
    setTimeout(() => {
        try {
            // Calculate daily data
            processedData = calculateOvertimeAndSalary(originalData, config);
            
            // Calculate employee summary
            salaryData = calculateEmployeeSummary(processedData, config);
            
            // Display results
            displayResults(processedData, salaryData);
            
            // Show results section
            resultsSection.style.display = 'block';
            resultsSection.scrollIntoView({ behavior: 'smooth' });
            
            showNotification('Perhitungan gaji selesai!', 'success');
            
            // Update sidebar
            document.getElementById('stat-salary').textContent = formatRupiah(
                salaryData.reduce((sum, emp) => sum + emp.totalGajiBersih, 0)
            );
            
        } catch (error) {
            console.error('Error processing data:', error);
            showNotification('Terjadi kesalahan saat menghitung gaji.', 'error');
        } finally {
            processBtn.innerHTML = '<i class="fas fa-calculator"></i> Hitung Lembur & Gaji';
            processBtn.disabled = false;
        }
    }, 1500);
}

// Display results
function displayResults(dailyData, employeeData) {
    updateMainStatistics(dailyData, employeeData);
    displayDailyTable(dailyData);
    displayEmployeeTable(employeeData);
    displaySummaries(employeeData);
    createCharts(employeeData);
}

// Update main statistics
function updateMainStatistics(dailyData, employeeData) {
    const totalKaryawan = new Set(dailyData.map(item => item.nama)).size;
    const totalHari = dailyData.length;
    const totalJam = dailyData.reduce((sum, item) => sum + item.durasi, 0);
    const totalLembur = dailyData.reduce((sum, item) => sum + item.jamLembur, 0);
    const totalGaji = employeeData.reduce((sum, emp) => sum + emp.totalGajiBersih, 0);
    
    document.getElementById('total-karyawan').textContent = totalKaryawan;
    document.getElementById('total-hari').textContent = totalHari;
    document.getElementById('total-jam').textContent = totalJam.toFixed(1) + ' jam';
    document.getElementById('total-lembur').textContent = totalLembur.toFixed(1) + ' jam';
    document.getElementById('total-gaji').textContent = formatRupiah(totalGaji);
}

// Display daily table
function displayDailyTable(data) {
    const tbody = document.getElementById('daily-table-body');
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
            <td>${item.durasi.toFixed(2)}</td>
            <td>${item.jamNormal.toFixed(2)}</td>
            <td><strong style="color: ${item.jamLembur > 0 ? '#e74c3c' : '#27ae60'};">${item.jamLembur.toFixed(2)}</strong></td>
            <td>${formatRupiah(item.gajiNormal)}</td>
            <td>${formatRupiah(item.gajiLembur)}</td>
        `;
        tbody.appendChild(row);
    });
}

// Display employee table
function displayEmployeeTable(data) {
    const tbody = document.getElementById('employee-table-body');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    data.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>${item.totalHari}</td>
            <td>${item.totalJam.toFixed(2)}</td>
            <td><strong style="color: ${item.totalJamLembur > 0 ? '#e74c3c' : '#27ae60'};">${item.totalJamLembur.toFixed(2)}</strong></td>
            <td>${formatRupiah(item.totalGajiNormal)}</td>
            <td>${formatRupiah(item.totalGajiLembur)}</td>
            <td>${formatRupiah(item.totalPajak)}</td>
            <td><strong>${formatRupiah(item.totalGajiBersih)}</strong></td>
        `;
        tbody.appendChild(row);
    });
}

// Display summaries
function displaySummaries(data) {
    const employeeSummary = document.getElementById('employee-summary');
    const financialSummary = document.getElementById('financial-summary');
    
    if (!employeeSummary && !financialSummary) return;
    
    // Employee summary
    let employeeHtml = '';
    data.forEach(employee => {
        employeeHtml += `
            <div style="margin-bottom: 1rem; padding-bottom: 1rem; border-bottom: 1px solid #eee;">
                <strong>${employee.nama}</strong><br>
                <small>
                    Hari Kerja: ${employee.totalHari} hari | 
                    Total Jam: ${employee.totalJam.toFixed(2)} jam<br>
                    <span style="color: #e74c3c; font-weight: bold;">
                        Lembur: ${employee.totalJamLembur.toFixed(2)} jam
                    </span><br>
                    Gaji Bersih: ${formatRupiah(employee.totalGajiBersih)}
                </small>
            </div>
        `;
    });
    
    if (employeeSummary) employeeSummary.innerHTML = employeeHtml;
    
    // Financial summary
    const totalJam = data.reduce((sum, emp) => sum + emp.totalJam, 0);
    const totalLembur = data.reduce((sum, emp) => sum + emp.totalJamLembur, 0);
    const totalGajiPokok = data.reduce((sum, emp) => sum + emp.totalGajiNormal, 0);
    const totalGajiLembur = data.reduce((sum, emp) => sum + emp.totalGajiLembur, 0);
    const totalPajak = data.reduce((sum, emp) => sum + emp.totalPajak, 0);
    const totalGajiBersih = data.reduce((sum, emp) => sum + emp.totalGajiBersih, 0);
    
    if (financialSummary) {
        financialSummary.innerHTML = `
            <div>Total Karyawan: <strong>${data.length} orang</strong></div>
            <div>Total Jam Kerja: <strong>${totalJam.toFixed(2)} jam</strong></div>
            <div>Total Jam Lembur: <strong>${totalLembur.toFixed(2)} jam</strong></div>
            <div>Total Gaji Pokok: <strong>${formatRupiah(totalGajiPokok)}</strong></div>
            <div style="color: #e74c3c;">Total Uang Lembur: <strong>${formatRupiah(totalGajiLembur)}</strong></div>
            <div>Total Pajak: <strong>${formatRupiah(totalPajak)}</strong></div>
            <div style="border-top: 2px solid #3498db; padding-top: 0.5rem; margin-top: 0.5rem; font-weight: bold;">
                Total Penggajian: <strong>${formatRupiah(totalGajiBersih)}</strong>
            </div>
        `;
    }
}

// Create charts
function createCharts(employeeData) {
    // Destroy existing charts
    if (hoursChart) hoursChart.destroy();
    if (salaryChart) salaryChart.destroy();
    
    // Prepare data for hours chart
    const employeeNames = employeeData.map(emp => emp.nama);
    const regularHours = employeeData.map(emp => emp.totalJamNormal);
    const overtimeHours = employeeData.map(emp => emp.totalJamLembur);
    
    // Hours chart
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
            scales: {
                y: {
                    beginAtZero: true,
                    title: {
                        display: true,
                        text: 'Jam Kerja'
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
    
    // Prepare data for salary chart
    const salaryComponents = [
        employeeData.reduce((sum, emp) => sum + emp.totalGajiNormal, 0),
        employeeData.reduce((sum, emp) => sum + emp.totalGajiLembur, 0),
        employeeData.reduce((sum, emp) => sum + emp.totalPajak, 0)
    ];
    
    // Salary chart
    const salaryCtx = document.getElementById('salaryChart').getContext('2d');
    salaryChart = new Chart(salaryCtx, {
        type: 'doughnut',
        data: {
            labels: ['Gaji Pokok', 'Uang Lembur', 'Pajak'],
            datasets: [{
                data: salaryComponents,
                backgroundColor: [
                    'rgba(52, 152, 219, 0.7)',
                    'rgba(231, 76, 60, 0.7)',
                    'rgba(46, 204, 113, 0.7)'
                ],
                borderColor: [
                    'rgba(52, 152, 219, 1)',
                    'rgba(231, 76, 60, 1)',
                    'rgba(46, 204, 113, 1)'
                ],
                borderWidth: 1
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom'
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.raw;
                            return `${context.label}: ${formatRupiah(value)}`;
                        }
                    }
                }
            }
        }
    });
}

// Switch tabs
function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.getAttribute('data-tab') === tabId) {
            btn.classList.add('active');
        }
    });
    
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
    document.getElementById('work-hours').value = DEFAULT_CONFIG.workHours;
    document.getElementById('salary-per-hour').value = '';
    document.getElementById('overtime-multiplier').value = DEFAULT_CONFIG.overtimeMultiplier;
    document.getElementById('tax-percentage').value = DEFAULT_CONFIG.taxPercentage;
    
    showNotification('Konfigurasi telah direset ke nilai default.', 'info');
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
    
    if ((type === 'daily' || type === 'employee') && processedData.length === 0) {
        showNotification('Data belum diproses.', 'warning');
        return;
    }
    
    try {
        if (type === 'original') {
            await generateReport(originalData, 'data_presensi_asli.xlsx', 'Data Asli');
            showNotification('Data asli berhasil diunduh.', 'success');
        } else if (type === 'daily') {
            await generateReport(processedData, 'data_lembur_gaji_harian.xlsx', 'Data Harian');
            showNotification('Data lembur & gaji harian berhasil diunduh.', 'success');
        } else if (type === 'employee') {
            await generateReport(salaryData, 'rekap_gaji_karyawan.xlsx', 'Rekap Karyawan');
            showNotification('Rekap gaji karyawan berhasil diunduh.', 'success');
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
