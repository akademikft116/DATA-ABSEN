// Main JavaScript file - Sistem Pengolahan Data Presensi
// DIPERBAIKI UNTUK FORMAT EXCEL ANDA

// Global variables
let originalData = [];
let processedData = [];
let currentFile = null;
let uploadProgressInterval = null;
let hoursChart = null;
let salaryChart = null;

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

// Format currency (Rupiah)
function formatCurrency(amount) {
    return new Intl.NumberFormat('id-ID', {
        style: 'currency',
        currency: 'IDR',
        minimumFractionDigits: 0
    }).format(amount);
}

// Format date from DD/MM/YYYY to YYYY-MM-DD
function formatDate(dateString) {
    if (!dateString) return '-';
    
    try {
        if (typeof dateString === 'string') {
            if (dateString.includes('/')) {
                // Format: DD/MM/YYYY
                const [day, month, year] = dateString.split('/');
                return `${year}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
            } else if (dateString.includes('-')) {
                // Format: YYYY-MM-DD
                return dateString;
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
                    date: parts[0],  // DD/MM/YYYY
                    time: parts[1]   // HH:MM
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
                } else if (timeStr.includes('.')) {
                    const decimal = parseFloat(timeStr);
                    const hours = Math.floor(decimal);
                    const minutes = Math.round((decimal - hours) * 60);
                    return { hours, minutes };
                }
            }
            return null;
        };
        
        const inTime = parseTime(timeIn);
        const outTime = parseTime(timeOut);
        
        if (!inTime || !outTime) return 0;
        
        let totalMinutes = (outTime.hours * 60 + outTime.minutes) - 
                          (inTime.hours * 60 + inTime.minutes);
        
        // Jika jam pulang lebih kecil dari jam masuk (misal lembur sampai pagi)
        if (totalMinutes < 0) {
            totalMinutes += 24 * 60; // Tambah 24 jam
        }
        
        // Convert to hours with 2 decimal places
        return Math.round((totalMinutes / 60) * 100) / 100;
        
    } catch (error) {
        console.error('Error calculating hours:', error);
        return 0;
    }
}

// Process Excel file - KHUSUS UNTUK FORMAT ANDA
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
                    
                    // Convert to array of arrays untuk format spesifik Anda
                    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                    
                    // Process data based on your specific format
                    const processedData = processYourExcelFormat(rawData);
                    allData = [...allData, ...processedData];
                });
                
                if (allData.length === 0) {
                    reject(new Error('Tidak ada data yang ditemukan dalam file Excel'));
                    return;
                }
                
                // Pair in-out times
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
    
    // Loop melalui semua baris
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        
        // Kolom E adalah indeks 4 (Nama)
        // Kolom F adalah indeks 5 (Waktu)
        if (row[4] && row[5]) {
            const nama = row[4];
            const waktu = row[5];
            
            // Parse datetime
            const { date, time } = parseDateTime(waktu);
            
            if (nama && date && time) {
                result.push({
                    nama: nama.toString().trim(),
                    tanggal: date, // Format: DD/MM/YYYY
                    waktu: time,    // Format: HH:MM
                    rawDatetime: waktu
                });
            }
        }
    }
    
    return result;
}

// Pair in and out times for each employee on each date
function pairInOutTimes(data) {
    // Group by nama and tanggal
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
    
    // Create paired records
    const result = [];
    
    Object.keys(grouped).forEach(key => {
        const [nama, tanggal] = key.split('_');
        const times = grouped[key];
        
        // Sort by time
        times.sort((a, b) => {
            const timeA = a.time.split(':').map(Number);
            const timeB = b.time.split(':').map(Number);
            return (timeA[0] * 60 + timeA[1]) - (timeB[0] * 60 + timeB[1]);
        });
        
        // If we have at least 2 records (in and out)
        if (times.length >= 2) {
            // Take first as in, last as out (assuming multiple entries)
            const jamMasuk = times[0].time;
            const jamKeluar = times[times.length - 1].time;
            
            // Calculate duration
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
            // Only one record (either in or out only)
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

// Calculate salaries from attendance data
function calculateSalaries(data, salaryPerHour = 50000, overtimeRate = 75000, taxRate = 5, workHours = 8) {
    const employees = {};
    
    // Group by employee
    data.forEach(record => {
        const name = record.nama.trim();
        
        if (!employees[name]) {
            employees[name] = {
                nama: name,
                records: [],
                totalHari: 0,
                totalJam: 0,
                jamNormal: 0,
                jamLembur: 0,
                detailHari: []
            };
        }
        
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
            
            // Detail per hari
            employees[name].detailHari.push({
                tanggal: record.tanggal,
                jamMasuk: record.jamMasuk,
                jamKeluar: record.jamKeluar,
                durasi: hoursWorked,
                jamNormal: regular,
                jamLembur: overtime
            });
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
            gajiPokok: Math.round(gajiPokok),
            uangLembur: Math.round(uangLembur),
            pajak: Math.round(pajak),
            gajiBersih: Math.round(gajiBersih),
            // For detailed report
            records: emp.records,
            detailHari: emp.detailHari
        };
    });
    
    // Sort by name
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
    
    const isProcessedData = data[0].gajiPokok !== undefined;
    
    if (isProcessedData) {
        // Processed data format
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Total Hari Kerja': item.totalHari,
            'Total Jam Kerja': item.totalJam.toFixed(2),
            'Jam Normal': item.jamNormal.toFixed(2),
            'Jam Lembur': item.jamLembur.toFixed(2),
            'Gaji Pokok (Rp)': item.gajiPokok,
            'Uang Lembur (Rp)': item.uangLembur,
            'Potongan Pajak (Rp)': item.pajak,
            'Gaji Bersih (Rp)': item.gajiBersih,
            'Keterangan': 'Data terhitung otomatis'
        }));
    } else {
        // Original data format (for your specific format)
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi (jam)': item.durasi ? item.durasi.toFixed(2) : '',
            'Keterangan': item.jamKeluar ? 
                `${item.jamMasuk} - ${item.jamKeluar} (${item.durasi.toFixed(2)} jam)` : 
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

// Download file
function downloadFile(data, filename, sheetName) {
    return generateReport(data, filename, sheetName);
}

// ============================
// MAIN APPLICATION FUNCTIONS
// ============================

// Initialize application
function initializeApp() {
    const now = new Date();
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    document.getElementById('current-date').textContent = now.toLocaleDateString('id-ID', options);
    
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            switchTab(tabId);
        });
    });
    
    const helpBtn = document.getElementById('help-btn');
    const closeHelpBtns = document.querySelectorAll('#close-help, #close-help-btn');
    const helpModal = document.getElementById('help-modal');
    
    helpBtn.addEventListener('click', () => helpModal.classList.add('active'));
    closeHelpBtns.forEach(btn => {
        btn.addEventListener('click', () => helpModal.classList.remove('active'));
    });
    
    const templateBtn = document.getElementById('template-btn');
    templateBtn.addEventListener('click', downloadTemplate);
    
    const resetBtn = document.getElementById('reset-config');
    resetBtn.addEventListener('click', resetConfig);
    
    document.getElementById('download-original').addEventListener('click', () => downloadReport('original'));
    document.getElementById('download-processed').addEventListener('click', () => downloadReport('processed'));
    document.getElementById('download-both').addEventListener('click', () => downloadReport('both'));
    
    setTimeout(() => {
        loadingScreen.style.opacity = '0';
        setTimeout(() => {
            loadingScreen.style.display = 'none';
            mainContainer.classList.add('loaded');
        }, 500);
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
        
        // Show preview of first few records
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
    
    // Insert after file preview
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

// Process data
function processData() {
    if (originalData.length === 0) {
        showNotification('Tidak ada data untuk diproses.', 'warning');
        return;
    }
    
    const salaryPerHour = parseFloat(document.getElementById('salary-per-hour').value) || 50000;
    const overtimeRate = parseFloat(document.getElementById('overtime-rate').value) || 75000;
    const taxRate = parseFloat(document.getElementById('tax-rate').value) || 5;
    const workHours = parseFloat(document.getElementById('work-hours').value) || 8;
    
    processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Memproses...';
    processBtn.disabled = true;
    
    setTimeout(() => {
        try {
            processedData = calculateSalaries(originalData, salaryPerHour, overtimeRate, taxRate, workHours);
            
            displayResults(processedData);
            createCharts(processedData);
            
            resultsSection.style.display = 'block';
            resultsSection.scrollIntoView({ behavior: 'smooth' });
            
            showNotification('Data berhasil diproses!', 'success');
            
        } catch (error) {
            console.error('Error processing data:', error);
            showNotification('Terjadi kesalahan saat memproses data.', 'error');
        } finally {
            processBtn.innerHTML = '<i class="fas fa-calculator"></i> Proses Data';
            processBtn.disabled = false;
        }
    }, 1500);
}

// Display results
function displayResults(data) {
    updateMainStatistics(data);
    displayProcessedTable(data);
    displayOriginalTable();
    displaySummaries(data);
}

// Update main statistics
function updateMainStatistics(data) {
    const totalKaryawan = new Set(data.map(item => item.nama)).size;
    const totalHari = data.reduce((sum, item) => sum + item.totalHari, 0);
    const totalGaji = data.reduce((sum, item) => sum + item.gajiBersih, 0);
    const totalLembur = data.reduce((sum, item) => sum + item.jamLembur, 0);
    
    document.getElementById('total-karyawan').textContent = totalKaryawan;
    document.getElementById('total-hari').textContent = totalHari;
    document.getElementById('total-gaji').textContent = formatCurrency(totalGaji);
    document.getElementById('total-lembur').textContent = totalLembur.toFixed(1) + ' jam';
}

// Display processed table
function displayProcessedTable(data) {
    const tbody = document.getElementById('processed-table-body');
    tbody.innerHTML = '';
    
    data.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>${item.totalHari}</td>
            <td>${item.totalJam.toFixed(2)}</td>
            <td>${item.jamLembur.toFixed(2)}</td>
            <td>${formatCurrency(item.gajiPokok)}</td>
            <td>${formatCurrency(item.uangLembur)}</td>
            <td>${formatCurrency(item.pajak)}</td>
            <td><strong style="color: #2ecc71;">${formatCurrency(item.gajiBersih)}</strong></td>
        `;
        tbody.appendChild(row);
    });
}

// Display original table
function displayOriginalTable() {
    const tbody = document.getElementById('original-table-body');
    tbody.innerHTML = '';
    
    const previewData = originalData.slice(0, 10);
    
    previewData.forEach((item, index) => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${index + 1}</td>
            <td>${item.nama || '-'}</td>
            <td>${formatDate(item.tanggal) || '-'}</td>
            <td>${formatTime(item.jamMasuk) || '-'}</td>
            <td>${formatTime(item.jamKeluar) || '-'}</td>
            <td>${item.durasi ? item.durasi.toFixed(2) + ' jam' : '-'}</td>
        `;
        tbody.appendChild(row);
    });
}

// Display summaries
function displaySummaries(data) {
    const employeeSummary = document.getElementById('employee-summary');
    const uniqueEmployees = [...new Set(data.map(item => item.nama))];
    
    let employeeHtml = '';
    uniqueEmployees.forEach(employee => {
        const employeeData = data.find(item => item.nama === employee);
        employeeHtml += `
            <div style="margin-bottom: 1rem; padding-bottom: 1rem; border-bottom: 1px solid #eee;">
                <strong>${employee}</strong><br>
                <small>
                    Total Hari: ${employeeData.totalHari} | 
                    Total Jam: ${employeeData.totalJam.toFixed(2)} | 
                    Lembur: ${employeeData.jamLembur.toFixed(2)} jam<br>
                    Gaji Bersih: ${formatCurrency(employeeData.gajiBersih)}
                </small>
            </div>
        `;
    });
    employeeSummary.innerHTML = employeeHtml;
    
    const financialSummary = document.getElementById('financial-summary');
    const totalGajiPokok = data.reduce((sum, item) => sum + item.gajiPokok, 0);
    const totalUangLembur = data.reduce((sum, item) => sum + item.uangLembur, 0);
    const totalPajak = data.reduce((sum, item) => sum + item.pajak, 0);
    const totalGajiBersih = data.reduce((sum, item) => sum + item.gajiBersih, 0);
    
    financialSummary.innerHTML = `
        <div>Total Gaji Pokok: <strong>${formatCurrency(totalGajiPokok)}</strong></div>
        <div>Total Uang Lembur: <strong>${formatCurrency(totalUangLembur)}</strong></div>
        <div>Total Potongan Pajak: <strong>${formatCurrency(totalPajak)}</strong></div>
        <div style="border-top: 2px solid #3498db; padding-top: 0.5rem; margin-top: 0.5rem;">
            Total Gaji Bersih: <strong style="color: #2ecc71; font-size: 1.1em;">${formatCurrency(totalGajiBersih)}</strong>
        </div>
    `;
}

// Create charts
function createCharts(data) {
    if (hoursChart) hoursChart.destroy();
    if (salaryChart) salaryChart.destroy();
    
    const employeeNames = [...new Set(data.map(item => item.nama))];
    const totalHours = employeeNames.map(name => {
        const employeeData = data.find(item => item.nama === name);
        return employeeData ? employeeData.totalJam : 0;
    });
    
    const regularHours = employeeNames.map(name => {
        const employeeData = data.find(item => item.nama === name);
        return employeeData ? employeeData.totalJam - employeeData.jamLembur : 0;
    });
    
    const overtimeHours = employeeNames.map(name => {
        const employeeData = data.find(item => item.nama === name);
        return employeeData ? employeeData.jamLembur : 0;
    });
    
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
    
    const salaryCtx = document.getElementById('salaryChart').getContext('2d');
    const salaries = employeeNames.map(name => {
        const employeeData = data.find(item => item.nama === name);
        return employeeData ? employeeData.gajiBersih : 0;
    });
    
    salaryChart = new Chart(salaryCtx, {
        type: 'pie',
        data: {
            labels: employeeNames,
            datasets: [{
                data: salaries,
                backgroundColor: [
                    '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0',
                    '#9966FF', '#FF9F40', '#C9CBCF', '#7E57C2',
                    '#42A5F5', '#26C6DA', '#66BB6A', '#FFCA28'
                ]
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            const value = context.raw;
                            const percentage = ((value / salaries.reduce((a, b) => a + b, 0)) * 100).toFixed(1);
                            return `${context.label}: ${formatCurrency(value)} (${percentage}%)`;
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
    document.getElementById('salary-per-hour').value = '';
    document.getElementById('overtime-rate').value = '';
    document.getElementById('tax-rate').value = '';
    document.getElementById('work-hours').value = '8';
    
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
    
    downloadFile(templateData, 'template_format_anda.xlsx', 'Template Presensi');
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
            await generateReport(processedData, 'data_presensi_terhitung.xlsx', 'Data Terhitung');
            showNotification('Data terhitung berhasil diunduh.', 'success');
        } else if (type === 'both') {
            await generateReport(originalData, 'data_presensi_asli.xlsx', 'Data Asli');
            setTimeout(async () => {
                await generateReport(processedData, 'data_presensi_terhitung.xlsx', 'Data Terhitung');
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
    notification.className = `notification-${type}`;
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
const style = document.createElement('style');
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
