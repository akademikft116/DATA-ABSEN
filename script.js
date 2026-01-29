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
// KONFIGURASI ATURAN BARU - HARI JUMAT
// ============================

// Hari Senin-Kamis special rules
const NORMAL_DAY_RULES = {
    workStartTime: '07:00',
    workEndTime: '16:00',
    minOvertimeMinutes: 10,
    dailyWorkHours: 8
};

// Hari Jumat special rules
const FRIDAY_RULES = {
    workStartTime: '07:00',
    workEndTime: '15:00',
    minOvertimeMinutes: 10,
    dailyWorkHours: 8
};

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

// DOM Elements
let loadingScreen, mainContainer, excelFileInput, browseBtn, uploadArea;
let processBtn, resultsSection, cancelUploadBtn;

// Debug function untuk mengecek elemen
function debugElements() {
    console.log('=== DEBUG ELEMENTS ===');
    const elements = {
        'loading-screen': document.getElementById('loading-screen'),
        'main-container': document.getElementById('main-container'),
        'excel-file': document.getElementById('excel-file'),
        'browse-btn': document.getElementById('browse-btn'),
        'upload-area': document.getElementById('upload-area'),
        'process-data': document.getElementById('process-data'),
        'file-preview': document.getElementById('file-preview'),
        'results-section': document.getElementById('results-section')
    };
    
    Object.entries(elements).forEach(([id, element]) => {
        console.log(`${id}:`, element ? '✓ FOUND' : '✗ NOT FOUND');
    });
    console.log('======================');
}

// Event Listeners
document.addEventListener('DOMContentLoaded', initializeApp);

// ============================
// HELPER FUNCTIONS - DENGAN ATURAN JUMAT
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
function formatHoursToDisplay(hours, rounded = false) {
    if (!hours || hours <= 0) return "0 jam";
    
    if (rounded) {
        const roundedHours = roundOvertimeHours(hours);
        return `${roundedHours} jam`;
    }
    
    const totalMinutes = Math.round(hours * 60);
    const h = Math.floor(totalMinutes / 60);
    const m = totalMinutes % 60;
    
    if (m === 0) {
        return `${h} jam`;
    } else {
        return `${h} jam ${m} menit`;
    }
}

// ============================
// FUNGSI PEMBULATAN BARU DENGAN BATAS 7 JAM
// ============================

// Fungsi untuk membulatkan jam lembur DENGAN BATAS MAKSIMAL 7 JAM
function roundOvertimeHours(hours) {
    if (!hours || hours <= 0) return 0;
    
    // Jika jam lembur melebihi 7 jam, batasi menjadi 7 jam
    if (hours > 7) {
        return 7;
    }
    
    // Pembulatan standar: 0.5 ke atas dibulatkan ke atas
    return Math.round(hours);
}

// Fungsi untuk membulatkan dan format jam lembur
function roundAndFormatHours(hours) {
    const rounded = roundOvertimeHours(hours);
    return `${rounded} jam`;
}

// Fungsi untuk menampilkan informasi lembur dengan batas 7 jam
function formatOvertimeInfo(overtimeHours) {
    const rounded = roundOvertimeHours(overtimeHours);
    
    if (overtimeHours > 7) {
        return `${rounded} jam (maksimal, dari ${overtimeHours.toFixed(2)} jam)`;
    } else if (rounded !== Math.round(overtimeHours)) {
        return `${rounded} jam (dibulatkan dari ${overtimeHours.toFixed(2)} jam)`;
    } else {
        return `${rounded} jam`;
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
            // Coba format "DD/MM/YYYY HH:MM"
            if (datetimeStr.includes('/') && datetimeStr.includes(':')) {
                const spaceIndex = datetimeStr.indexOf(' ');
                if (spaceIndex > 0) {
                    return {
                        date: datetimeStr.substring(0, spaceIndex).trim(),
                        time: datetimeStr.substring(spaceIndex + 1).trim()
                    };
                }
            }
            
            // Coba format Excel date number
            if (!isNaN(datetimeStr) && datetimeStr.toString().length > 5) {
                const excelDate = parseFloat(datetimeStr);
                const date = new Date((excelDate - 25569) * 86400 * 1000);
                return {
                    date: `${date.getDate().toString().padStart(2, '0')}/${(date.getMonth() + 1).toString().padStart(2, '0')}/${date.getFullYear()}`,
                    time: `${date.getHours().toString().padStart(2, '0')}:${date.getMinutes().toString().padStart(2, '0')}`
                };
            }
        }
    } catch (error) {
        console.error('Error parsing datetime:', error);
    }
    
    return { date: '', time: '' };
}

// ============================
// FUNGSI UNTUK HARI JUMAT DAN SABTU
// ============================

// Fungsi untuk mengecek apakah suatu tanggal adalah hari Jumat
function isFriday(dateString) {
    if (!dateString) return false;
    
    try {
        const [day, month, year] = dateString.split('/');
        const date = new Date(year, month - 1, day);
        return date.getDay() === 5;
    } catch (error) {
        console.error('Error checking Friday:', error);
        return false;
    }
}

// Fungsi untuk mengecek apakah suatu tanggal adalah hari Sabtu
function isSaturday(dateString) {
    if (!dateString) return false;
    
    try {
        const [day, month, year] = dateString.split('/');
        const date = new Date(year, month - 1, day);
        return date.getDay() === 6;
    } catch (error) {
        console.error('Error checking Saturday:', error);
        return false;
    }
}

// Fungsi untuk mendapatkan nama hari dalam bahasa Indonesia
function getDayName(dateString) {
    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    try {
        const [day, month, year] = dateString.split('/');
        const date = new Date(year, month - 1, day);
        return days[date.getDay()];
    } catch (error) {
        return '';
    }
}

function getDayRules(dateString) {
    if (isSaturday(dateString)) {
        return SATURDAY_RULES;
    } else if (isFriday(dateString)) {
        return FRIDAY_RULES;
    } else {
        return NORMAL_DAY_RULES;
    }
}

// ============================
// KONFIGURASI ATURAN HARI SABTU
// ============================

const SATURDAY_RULES = {
    workHours: 6,
    minOvertimeMinutes: 10,
    k3SpecialHours: {
        startTime: '07:00',
        endTime: '22:00',
        isSpecialDay: true
    }
};

// Calculate hours between two time strings DENGAN ATURAN BARU (FLEKSIBEL JAM MASUK)
function calculateHoursWithFridayRules(timeIn, timeOut, dateString) {
    if (!timeIn || !timeOut || !dateString) return 0;
    
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
        
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        const outMinutes = outTime.hours * 60 + outTime.minutes;
        
        let totalMinutes = outMinutes - inMinutes;
        
        if (totalMinutes < 0) {
            totalMinutes += 24 * 60;
        }
        
        const totalHours = Math.round((totalMinutes / 60) * 100) / 100;
        
        return totalHours;
        
    } catch (error) {
        console.error('Error calculating hours with Friday rules:', error);
        return 0;
    }
}

// Fungsi untuk menghitung lembur dengan aturan baru (FLEKSIBEL JAM MASUK)
function calculateOvertimeWithDayRules(jamMasuk, jamKeluar, dateString) {
    if (!jamMasuk || !jamKeluar || !dateString) return 0;
    
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
        const outTime = parseTime(jamKeluar);
        if (!inTime || !outTime) return 0;
        
        const rules = getDayRules(dateString);
        
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        const outMinutes = outTime.hours * 60 + outTime.minutes;
        
        let totalMinutes = outMinutes - inMinutes;
        if (totalMinutes < 0) totalMinutes += 24 * 60;
        
        // Tentukan jam normal berdasarkan hari
        let maxNormalMinutes;
        if (isSaturday(dateString)) {
            maxNormalMinutes = 6 * 60; // 6 jam untuk Sabtu
        } else {
            maxNormalMinutes = 8 * 60; // 8 jam untuk hari lain
        }
        
        if (totalMinutes <= maxNormalMinutes) return 0;
        
        let overtimeMinutes = totalMinutes - maxNormalMinutes;
        
        if (overtimeMinutes < rules.minOvertimeMinutes) {
            return 0;
        }
        
        const overtimeHours = Math.round((overtimeMinutes / 60) * 100) / 100;
        
        // BATASI MAKSIMAL 7 JAM LEMBUR PER HARI
        const maxOvertimeHours = 7;
        const limitedOvertimeHours = Math.min(overtimeHours, maxOvertimeHours);
        
        return limitedOvertimeHours;
        
    } catch (error) {
        console.error('Error calculating overtime with day rules:', error);
        return 0;
    }
}

// Calculate hours between two time strings (untuk kompatibilitas)
function calculateHours(timeIn, timeOut) {
    return calculateHoursWithFridayRules(timeIn, timeOut, '01/01/2024');
}

// ============================
// FUNGSI UNTUK MEMISAHKAN DATA PER HARI DAN PER KARYAWAN
// ============================

// Fungsi untuk memisahkan data berdasarkan hari dan karyawan
function separateDataByDayAndEmployee(data) {
    const result = {
        seninKamis: {},
        jumat: {},
        sabtu: {}
    };
    
    data.forEach(item => {
        const employeeName = item.nama;
        
        if (item.isFriday) {
            // Data Jumat
            if (!result.jumat[employeeName]) {
                result.jumat[employeeName] = [];
            }
            result.jumat[employeeName].push(item);
        } else if (item.isSaturday) {
            // Data Sabtu
            if (!result.sabtu[employeeName]) {
                result.sabtu[employeeName] = [];
            }
            result.sabtu[employeeName].push(item);
        } else {
            // Data Senin-Kamis
            if (!result.seninKamis[employeeName]) {
                result.seninKamis[employeeName] = [];
            }
            result.seninKamis[employeeName].push(item);
        }
    });
    
    return result;
}

// Fungsi untuk menghitung summary per karyawan per hari
function calculateEmployeeSummaryByDay(data) {
    const separated = separateDataByDayAndEmployee(data);
    const summary = {
        seninKamis: {},
        jumat: {},
        sabtu: {}
    };
    
    // Hitung summary untuk Senin-Kamis
    Object.keys(separated.seninKamis).forEach(employeeName => {
        const employeeData = separated.seninKamis[employeeName];
        const totalLembur = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
        const totalLemburBulat = roundOvertimeHours(totalLembur);
        const category = employeeCategories[employeeName] || 'STAFF';
        const rate = overtimeRates[category];
        const totalGaji = totalLemburBulat * rate;
        
        summary.seninKamis[employeeName] = {
            nama: employeeName,
            kategori: category,
            rate: rate,
            totalHari: employeeData.length,
            totalLembur: totalLembur,
            totalLemburBulat: totalLemburBulat,
            totalGaji: totalGaji,
            data: employeeData
        };
    });
    
    // Hitung summary untuk Jumat
    Object.keys(separated.jumat).forEach(employeeName => {
        const employeeData = separated.jumat[employeeName];
        const totalLembur = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
        const totalLemburBulat = roundOvertimeHours(totalLembur);
        const category = employeeCategories[employeeName] || 'STAFF';
        const rate = overtimeRates[category];
        const totalGaji = totalLemburBulat * rate;
        
        summary.jumat[employeeName] = {
            nama: employeeName,
            kategori: category,
            rate: rate,
            totalHari: employeeData.length,
            totalLembur: totalLembur,
            totalLemburBulat: totalLemburBulat,
            totalGaji: totalGaji,
            data: employeeData
        };
    });
    
    // Hitung summary untuk Sabtu (dengan aturan khusus)
    Object.keys(separated.sabtu).forEach(employeeName => {
        const employeeData = separated.sabtu[employeeName];
        const category = employeeCategories[employeeName] || 'STAFF';
        const rate = overtimeRates[category];
        
        // Hitung lembur khusus Sabtu
        let totalLemburSabtu = 0;
        employeeData.forEach(item => {
            if (category === 'K3') {
                // K3: lembur jika pulang setelah 22:00
                const [outHour, outMinute] = item.jamKeluar.split(':').map(Number);
                const outMinutes = outHour * 60 + (outMinute || 0);
                const workEndMinutes = 22 * 60; // 22:00
                
                if (outMinutes > workEndMinutes) {
                    totalLemburSabtu += (outMinutes - workEndMinutes) / 60;
                }
            } else {
                // Non-K3: lembur setelah 6 jam kerja
                if (item.durasi > 6) {
                    totalLemburSabtu += Math.max(0, item.durasi - 6);
                }
            }
        });
        
        // Abaikan lembur kurang dari 10 menit
        if (totalLemburSabtu < (SATURDAY_RULES.minOvertimeMinutes / 60)) {
            totalLemburSabtu = 0;
        }
        
        // Batasi maksimal 7 jam
        if (totalLemburSabtu > 7) {
            totalLemburSabtu = 7;
        }
        
        const totalLemburBulat = roundOvertimeHours(totalLemburSabtu);
        const totalGaji = totalLemburBulat * rate;
        
        summary.sabtu[employeeName] = {
            nama: employeeName,
            kategori: category,
            rate: rate,
            totalHari: employeeData.length,
            totalLembur: totalLemburSabtu,
            totalLemburBulat: totalLemburBulat,
            totalGaji: totalGaji,
            data: employeeData
        };
    });
    
    return summary;
}

// ============================
// FUNGSI UPLOAD FILE YANG DIPERBAIKI
// ============================

// Inisialisasi event listeners dengan error handling
function setupEventListeners() {
    console.log('Setting up event listeners...');
    
    try {
        // Dapatkan elemen
        loadingScreen = document.getElementById('loading-screen');
        mainContainer = document.getElementById('main-container');
        excelFileInput = document.getElementById('excel-file');
        browseBtn = document.getElementById('browse-btn');
        uploadArea = document.getElementById('upload-area');
        processBtn = document.getElementById('process-data');
        resultsSection = document.getElementById('results-section');
        cancelUploadBtn = document.getElementById('cancel-upload');
        
        console.log('Elements initialized:');
        console.log('- uploadArea:', uploadArea);
        console.log('- excelFileInput:', excelFileInput);
        console.log('- browseBtn:', browseBtn);
        
        // 1. Event untuk klik di upload area
        if (uploadArea) {
            uploadArea.style.cursor = 'pointer';
            
            uploadArea.addEventListener('click', function(e) {
                console.log('Upload area clicked');
                // Pastikan tidak mengklik tombol browse
                if (!e.target.closest('#browse-btn')) {
                    if (excelFileInput) {
                        excelFileInput.click();
                    }
                }
            });
            
            // Drag and drop support
            uploadArea.addEventListener('dragover', function(e) {
                e.preventDefault();
                e.stopPropagation();
                uploadArea.classList.add('drag-over');
                uploadArea.style.borderColor = '#3498db';
                uploadArea.style.background = 'rgba(52, 152, 219, 0.1)';
            });
            
            uploadArea.addEventListener('dragleave', function(e) {
                e.preventDefault();
                e.stopPropagation();
                uploadArea.classList.remove('drag-over');
                uploadArea.style.borderColor = '';
                uploadArea.style.background = '';
            });
            
            uploadArea.addEventListener('drop', function(e) {
                e.preventDefault();
                e.stopPropagation();
                uploadArea.classList.remove('drag-over');
                uploadArea.style.borderColor = '';
                uploadArea.style.background = '';
                
                if (e.dataTransfer.files.length > 0) {
                    const file = e.dataTransfer.files[0];
                    console.log('File dropped:', file.name);
                    
                    // Validasi file
                    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
                        showNotification('Hanya file Excel (.xlsx, .xls) atau CSV yang didukung', 'error');
                        return;
                    }
                    
                    // Simulasi file input
                    if (excelFileInput) {
                        const dataTransfer = new DataTransfer();
                        dataTransfer.items.add(file);
                        excelFileInput.files = dataTransfer.files;
                        
                        // Trigger event change
                        const event = new Event('change', { bubbles: true });
                        excelFileInput.dispatchEvent(event);
                    }
                }
            });
        }
        
        // 2. Event untuk tombol browse
        if (browseBtn) {
            browseBtn.addEventListener('click', function(e) {
                console.log('Browse button clicked');
                e.stopPropagation(); // Hindari event bubbling
                if (excelFileInput) {
                    excelFileInput.click();
                }
            });
        }
        
        // 3. Event untuk file input
        if (excelFileInput) {
            excelFileInput.addEventListener('change', function(e) {
                console.log('File input changed');
                if (e.target.files && e.target.files[0]) {
                    handleFileSelect(e);
                }
            });
        }
        
        // 4. Event untuk tombol proses
        if (processBtn) {
            processBtn.addEventListener('click', processData);
        }
        
        // 5. Event untuk tombol cancel
        if (cancelUploadBtn) {
            cancelUploadBtn.addEventListener('click', cancelUpload);
        }
        
        // 6. Event untuk tabs
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                const tabId = this.getAttribute('data-tab');
                switchTab(tabId);
            });
        });
        
        // 7. Event untuk modal help
        const helpBtn = document.getElementById('help-btn');
        const closeHelpBtns = document.querySelectorAll('#close-help, #close-help-btn');
        const helpModal = document.getElementById('help-modal');
        
        if (helpBtn) helpBtn.addEventListener('click', () => helpModal?.classList.add('active'));
        closeHelpBtns.forEach(btn => {
            btn.addEventListener('click', () => helpModal?.classList.remove('active'));
        });
        
        // 8. Event untuk template download
        const templateBtn = document.getElementById('template-btn');
        if (templateBtn) templateBtn.addEventListener('click', downloadTemplate);
        
        // 9. Event untuk reset config
        const resetBtn = document.getElementById('reset-config');
        if (resetBtn) resetBtn.addEventListener('click', resetConfig);
        
        // 10. Event untuk download buttons
        const downloadOriginal = document.getElementById('download-original');
        const downloadProcessed = document.getElementById('download-processed');
        const downloadBoth = document.getElementById('download-both');
        const downloadSalary = document.getElementById('download-salary');
        const downloadSaturday = document.getElementById('download-saturday');
        
        if (downloadOriginal) downloadOriginal.addEventListener('click', () => downloadReport('original'));
        if (downloadProcessed) downloadProcessed.addEventListener('click', () => downloadReport('processed'));
        if (downloadBoth) downloadBoth.addEventListener('click', () => downloadReport('both'));
        if (downloadSalary) downloadSalary.addEventListener('click', showSeparatedDownloadOptions);
        if (downloadSaturday) downloadSaturday.addEventListener('click', () => {
            if (processedData.length === 0) {
                showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
                return;
            }
            showSaturdayDownloadModal();
        });
        
        console.log('Event listeners setup complete!');
        
    } catch (error) {
        console.error('Error in setupEventListeners:', error);
        showNotification('Error setting up event listeners', 'error');
        
        // Fallback: coba setup minimal
        setupMinimalEventListeners();
    }
}

// Fallback minimal event listeners
function setupMinimalEventListeners() {
    console.log('Setting up minimal event listeners...');
    
    const excelFileInput = document.getElementById('excel-file');
    const browseBtn = document.getElementById('browse-btn');
    const uploadArea = document.getElementById('upload-area');
    
    if (browseBtn && excelFileInput) {
        browseBtn.onclick = function() {
            excelFileInput.click();
        };
    }
    
    if (uploadArea && excelFileInput) {
        uploadArea.onclick = function() {
            excelFileInput.click();
        };
    }
    
    if (excelFileInput) {
        excelFileInput.onchange = handleFileSelect;
    }
    
    const processBtn = document.getElementById('process-data');
    if (processBtn) {
        processBtn.onclick = processData;
    }
}

function calculateOvertimeByDayType(data) {
    const separated = separateDataByDay(data);
    
    const fridayOvertime = separated.friday.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    const otherDaysOvertime = separated.otherDays.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    
    return {
        friday: {
            data: separated.friday,
            totalOvertime: fridayOvertime,
            totalDays: separated.friday.length,
            totalEmployees: new Set(separated.friday.map(item => item.nama)).size
        },
        otherDays: {
            data: separated.otherDays,
            totalOvertime: otherDaysOvertime,
            totalDays: separated.otherDays.length,
            totalEmployees: new Set(separated.otherDays.map(item => item.nama)).size
        }
    };
}

// ============================
// FUNGSI PERHITUNGAN BARU DENGAN ATURAN JUMAT DAN FLEKSIBEL JAM MASUK
// ============================

// Calculate overtime per day - DENGAN ATURAN BARU (FLEKSIBEL JAM MASUK)
function calculateOvertimePerDay(data, workHours = 8) {
    const result = data.map(record => {
        // Hitung total jam kerja (menggunakan jam masuk sebenarnya)
        const hoursWorked = calculateHoursWithFridayRules(
            record.jamMasuk, 
            record.jamKeluar, 
            record.tanggal
        );
        
        // Hitung lembur dengan aturan baru (FLEKSIBEL JAM MASUK)
        const jamLemburDesimal = calculateOvertimeWithDayRules(
            record.jamMasuk, 
            record.jamKeluar, 
            record.tanggal
        );
        
        // Tentukan jam normal berdasarkan hari
        let jamNormal;
        if (isSaturday(record.tanggal)) {
            jamNormal = Math.min(hoursWorked, 6); // 6 jam untuk Sabtu
        } else {
            jamNormal = Math.min(hoursWorked, workHours); // 8 jam untuk hari lain
        }
        
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
        
        // Cek apakah hari Jumat atau Sabtu
        const isFridayDay = isFriday(record.tanggal);
        const isSaturdayDay = isSaturday(record.tanggal);
        
        // Keterangan berdasarkan hari
        let keterangan = 'Tidak lembur';
        
        if (isSaturdayDay) {
            if (jamLemburDesimal > 0) {
                keterangan = `Lembur Sabtu ${jamLemburDisplay} (kerja > 6 jam)`;
            } else {
                keterangan = `Sabtu - Kerja ≤ 6 jam`;
            }
        } else if (isFridayDay) {
            if (jamLemburDesimal > 0) {
                keterangan = `Lembur ${jamLemburDisplay} (Jumat - kerja > 8 jam)`;
            } else {
                keterangan = `Hari Jumat - Kerja 8 jam normal`;
            }
        } else {
            if (jamLemburDesimal > 0) {
                keterangan = `Lembur ${jamLemburDisplay} (kerja > 8 jam)`;
            } else {
                keterangan = `Kerja ≤ 8 jam`;
            }
        }
        
        // Tambahkan info jika lembur dibatasi 7 jam
        if (jamLemburDesimal > 7) {
            keterangan = `Lembur ${jamLemburDisplay} (DIBATASI MAKSIMAL 7 JAM)`;
        }
        
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
            keterangan: keterangan,
            jamKerjaNormal: workHours,
            isFriday: isFridayDay,
            isSaturday: isSaturdayDay,
            hari: getDayName(record.tanggal)
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

// ============================
// FUNGSI UNTUK DOWNLOAD PER KARYAWAN
// ============================

// Fungsi untuk menampilkan modal pilihan download per karyawan
function showSeparatedDownloadOptions() {
    if (processedData.length === 0) {
        showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
        return;
    }
    
    // Hitung summary per karyawan per hari
    const summary = calculateEmployeeSummaryByDay(processedData);
    
    // Dapatkan daftar karyawan unik
    const allEmployees = new Set();
    Object.keys(summary.seninKamis).forEach(emp => allEmployees.add(emp));
    Object.keys(summary.jumat).forEach(emp => allEmployees.add(emp));
    Object.keys(summary.sabtu).forEach(emp => allEmployees.add(emp));
    
    const employees = Array.from(allEmployees);
    
    // Buat modal dengan opsi per karyawan
    const modalHtml = `
        <div class="modal" id="per-employee-download-modal">
            <div class="modal-content" style="max-width: 800px;">
                <div class="modal-header">
                    <h3><i class="fas fa-users"></i> Download Per Karyawan</h3>
                    <button class="modal-close" id="close-per-employee-download">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="download-stats" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 1rem; border-radius: 8px; margin-bottom: 1.5rem;">
                        <h4><i class="fas fa-chart-bar"></i> Statistik Data</h4>
                        <div style="display: flex; justify-content: space-around; text-align: center;">
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Total Karyawan</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${employees.length}</p>
                            </div>
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Senin-Kamis</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${Object.keys(summary.seninKamis).length}</p>
                            </div>
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Jumat</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${Object.keys(summary.jumat).length}</p>
                            </div>
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Sabtu</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${Object.keys(summary.sabtu).length}</p>
                            </div>
                        </div>
                    </div>
                    
                    <div style="margin-bottom: 1rem;">
                        <h4 style="color: #2c3e50; margin-bottom: 0.5rem;">Pilih Karyawan:</h4>
                        <div class="employee-list" style="max-height: 300px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px; padding: 0.5rem;">
                            ${employees.map(employee => `
                                <div class="employee-item" data-employee="${employee}" style="padding: 0.75rem; border-bottom: 1px solid #eee; cursor: pointer; transition: background-color 0.3s;">
                                    <div style="display: flex; justify-content: space-between; align-items: center;">
                                        <div>
                                            <strong>${employee}</strong>
                                            <div style="font-size: 0.85rem; color: #666;">
                                                <span>Senin-Kamis: ${summary.seninKamis[employee]?.totalHari || 0} hari</span>
                                                <span style="margin-left: 1rem;">Jumat: ${summary.jumat[employee]?.totalHari || 0} hari</span>
                                                <span style="margin-left: 1rem;">Sabtu: ${summary.sabtu[employee]?.totalHari || 0} hari</span>
                                            </div>
                                        </div>
                                        <div style="text-align: right;">
                                            <div style="font-weight: bold; color: #e74c3c;">
                                                ${(summary.seninKamis[employee]?.totalLemburBulat || 0) + 
                                                  (summary.jumat[employee]?.totalLemburBulat || 0) + 
                                                  (summary.sabtu[employee]?.totalLemburBulat || 0)} jam
                                            </div>
                                            <div style="font-size: 0.85rem; color: #27ae60;">
                                                ${formatCurrency(
                                                    (summary.seninKamis[employee]?.totalGaji || 0) + 
                                                    (summary.jumat[employee]?.totalGaji || 0) + 
                                                    (summary.sabtu[employee]?.totalGaji || 0)
                                                )}
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            `).join('')}
                        </div>
                    </div>
                    
                    <div class="download-options-per-employee" style="margin-top: 1.5rem;">
                        <h4 style="color: #2c3e50; margin-bottom: 0.5rem;">Pilih Format Download:</h4>
                        <div class="option-cards-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 1rem;">
                            <div class="option-card" id="option-per-employee-all">
                                <div class="option-icon">
                                    <i class="fas fa-file-excel" style="color: #217346; font-size: 2rem;"></i>
                                </div>
                                <div class="option-content">
                                    <h4>Semua Hari</h4>
                                    <p>Download semua data lembur karyawan terpilih</p>
                                    <ul>
                                        <li>Senin-Kamis, Jumat, dan Sabtu</li>
                                        <li>Dalam satu file Excel</li>
                                        <li>Format tabel lengkap</li>
                                    </ul>
                                </div>
                            </div>
                            
                            <div class="option-card" id="option-per-employee-seninkamis">
                                <div class="option-icon">
                                    <i class="fas fa-calendar-alt" style="color: #3498db; font-size: 2rem;"></i>
                                </div>
                                <div class="option-content">
                                    <h4>Hanya Senin-Kamis</h4>
                                    <p>Download data lembur hari Senin-Kamis</p>
                                    <ul>
                                        <li>Aturan normal: 8 jam kerja</li>
                                        <li>Lembur setelah jam 16:00</li>
                                        <li>Format per karyawan</li>
                                    </ul>
                                </div>
                            </div>
                            
                            <div class="option-card" id="option-per-employee-jumat">
                                <div class="option-icon">
                                    <i class="fas fa-calendar-day" style="color: #9b59b6; font-size: 2rem;"></i>
                                </div>
                                <div class="option-content">
                                    <h4>Hanya Jumat</h4>
                                    <p>Download data lembur hari Jumat</p>
                                    <ul>
                                        <li>Aturan khusus: pulang jam 15:00</li>
                                        <li>Lembur setelah jam 15:00</li>
                                        <li>Format per karyawan</li>
                                    </ul>
                                </div>
                            </div>
                            
                            <div class="option-card" id="option-per-employee-sabtu">
                                <div class="option-icon">
                                    <i class="fas fa-calendar-day" style="color: #e67e22; font-size: 2rem;"></i>
                                </div>
                                <div class="option-content">
                                    <h4>Hanya Sabtu</h4>
                                    <p>Download data lembur hari Sabtu</p>
                                    <ul>
                                        <li>Aturan khusus: kerja 6 jam</li>
                                        <li>K3: kerja 07:00-22:00</li>
                                        <li>Format per karyawan</li>
                                    </ul>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="cancel-per-employee-download">Batal</button>
                    <button class="btn btn-primary" id="download-selected-employee" disabled>
                        <i class="fas fa-download"></i> Download
                    </button>
                </div>
            </div>
        </div>
    `;
    
    // Tambahkan modal ke body jika belum ada
    let modal = document.getElementById('per-employee-download-modal');
    if (!modal) {
        document.body.insertAdjacentHTML('beforeend', modalHtml);
        modal = document.getElementById('per-employee-download-modal');
        
        // Variabel untuk menyimpan karyawan yang dipilih
        let selectedEmployee = null;
        let selectedFormat = null;
        
        // Setup event listeners untuk modal
        document.getElementById('close-per-employee-download').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        document.getElementById('cancel-per-employee-download').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        // Event untuk memilih karyawan
        document.querySelectorAll('.employee-item').forEach(item => {
            item.addEventListener('click', function() {
                // Hapus seleksi sebelumnya
                document.querySelectorAll('.employee-item').forEach(el => {
                    el.style.backgroundColor = '';
                    el.style.borderLeft = '';
                });
                
                // Tandai karyawan yang dipilih
                this.style.backgroundColor = '#e3f2fd';
                this.style.borderLeft = '3px solid #2196f3';
                
                selectedEmployee = this.getAttribute('data-employee');
                
                // Enable tombol download
                const downloadBtn = document.getElementById('download-selected-employee');
                if (downloadBtn && selectedEmployee && selectedFormat) {
                    downloadBtn.disabled = false;
                }
            });
        });
        
        // Event untuk memilih format
        document.querySelectorAll('.option-card').forEach(card => {
            card.addEventListener('click', function() {
                // Hapus seleksi sebelumnya
                document.querySelectorAll('.option-card').forEach(el => {
                    el.style.backgroundColor = '';
                    el.style.boxShadow = '';
                });
                
                // Tandai format yang dipilih
                this.style.backgroundColor = '#f8f9fa';
                this.style.boxShadow = '0 0 0 2px #3498db';
                
                selectedFormat = this.id.replace('option-per-employee-', '');
                
                // Enable tombol download
                const downloadBtn = document.getElementById('download-selected-employee');
                if (downloadBtn && selectedEmployee && selectedFormat) {
                    downloadBtn.disabled = false;
                }
            });
        });
        
        // Event untuk tombol download
        document.getElementById('download-selected-employee').addEventListener('click', () => {
            if (!selectedEmployee || !selectedFormat) {
                showNotification('Pilih karyawan dan format download terlebih dahulu', 'warning');
                return;
            }
            
            downloadPerEmployeeData(selectedEmployee, selectedFormat);
            modal.classList.remove('active');
        });
        
        // Close modal ketika klik di luar
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });
    }
    
    modal.classList.add('active');
}

// Fungsi untuk download data per karyawan
function downloadPerEmployeeData(employeeName, formatType) {
    if (!employeeName || !formatType) {
        showNotification('Data karyawan tidak valid', 'error');
        return;
    }
    
    const summary = calculateEmployeeSummaryByDay(processedData);
    let dataToDownload = [];
    let sheetName = '';
    let fileName = '';
    
    switch(formatType) {
        case 'all':
            // Gabungkan semua data
            dataToDownload = [
                ...(summary.seninKamis[employeeName]?.data || []),
                ...(summary.jumat[employeeName]?.data || []),
                ...(summary.sabtu[employeeName]?.data || [])
            ];
            sheetName = `Data Lembur ${employeeName}`;
            fileName = `lembur_${employeeName.replace(/\s+/g, '_').toLowerCase()}_semua_hari`;
            break;
            
        case 'seninkamis':
            dataToDownload = summary.seninKamis[employeeName]?.data || [];
            sheetName = `Lembur Senin-Kamis ${employeeName}`;
            fileName = `lembur_${employeeName.replace(/\s+/g, '_').toLowerCase()}_senin_kamis`;
            break;
            
        case 'jumat':
            dataToDownload = summary.jumat[employeeName]?.data || [];
            sheetName = `Lembur Jumat ${employeeName}`;
            fileName = `lembur_${employeeName.replace(/\s+/g, '_').toLowerCase()}_jumat`;
            break;
            
        case 'sabtu':
            dataToDownload = summary.sabtu[employeeName]?.data || [];
            sheetName = `Lembur Sabtu ${employeeName}`;
            fileName = `lembur_${employeeName.replace(/\s+/g, '_').toLowerCase()}_sabtu`;
            break;
    }
    
    if (dataToDownload.length === 0) {
        showNotification(`Tidak ada data ${formatType} untuk karyawan ${employeeName}`, 'warning');
        return;
    }
    
    try {
        // Buat data untuk export
        const exportData = dataToDownload.map((item, index) => {
            const jamLemburBulat = roundOvertimeHours(item.jamLemburDesimal);
            const category = employeeCategories[item.nama] || 'STAFF';
            const rate = overtimeRates[category];
            const gajiLembur = jamLemburBulat * rate;
            
            return {
                'No': index + 1,
                'Tanggal': formatDate(item.tanggal),
                'Hari': item.hari,
                'Jam Masuk': item.jamMasuk,
                'Jam Keluar': item.jamKeluar,
                'Durasi Kerja': item.durasiFormatted,
                'Jam Normal': item.jamNormalFormatted,
                'Jam Lembur (Bulat)': jamLemburBulat,
                'Jam Lembur (Desimal)': item.jamLemburDesimal.toFixed(2),
                'Rate per Jam': formatCurrency(rate),
                'Total Gaji Lembur': formatCurrency(gajiLembur),
                'Keterangan': item.keterangan
            };
        });
        
        // Tambahkan summary row di akhir
        const totalLemburBulat = dataToDownload.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
        const totalLemburDesimal = dataToDownload.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
        const category = employeeCategories[employeeName] || 'STAFF';
        const rate = overtimeRates[category];
        const totalGaji = totalLemburBulat * rate;
        
        exportData.push({
            'No': '',
            'Tanggal': 'TOTAL',
            'Hari': '',
            'Jam Masuk': '',
            'Jam Keluar': '',
            'Durasi Kerja': '',
            'Jam Normal': '',
            'Jam Lembur (Bulat)': totalLemburBulat,
            'Jam Lembur (Desimal)': totalLemburDesimal.toFixed(2),
            'Rate per Jam': formatCurrency(rate),
            'Total Gaji Lembur': formatCurrency(totalGaji),
            'Keterangan': `Total ${dataToDownload.length} hari kerja`
        });
        
        const worksheet = XLSX.utils.json_to_sheet(exportData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        
        // Set kolom width
        const wscols = [
            {wch: 5},   // No
            {wch: 12},  // Tanggal
            {wch: 10},  // Hari
            {wch: 10},  // Jam Masuk
            {wch: 10},  // Jam Keluar
            {wch: 12},  // Durasi
            {wch: 12},  // Jam Normal
            {wch: 15},  // Jam Lembur (Bulat)
            {wch: 15},  // Jam Lembur (Desimal)
            {wch: 15},  // Rate
            {wch: 20},  // Total Gaji
            {wch: 40}   // Keterangan
        ];
        worksheet['!cols'] = wscols;
        
        XLSX.writeFile(workbook, `${fileName}_${new Date().toISOString().split('T')[0]}.xlsx`);
        
        showNotification(`Data ${employeeName} berhasil diunduh!`, 'success');
        
    } catch (error) {
        console.error('Error downloading per employee data:', error);
        showNotification('Gagal mengunduh data karyawan', 'error');
    }
}

// ============================
// FUNGSI UNTUK DOWNLOAD SABTU PER KARYAWAN
// ============================

// Fungsi untuk modal download Sabtu per karyawan
function showSaturdayDownloadModal() {
    if (processedData.length === 0) {
        showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
        return;
    }
    
    // Filter data Sabtu
    const saturdayData = processedData.filter(item => item.isSaturday);
    
    if (saturdayData.length === 0) {
        showNotification('Tidak ada data lembur hari Sabtu', 'info');
        return;
    }
    
    // Kelompokkan data per karyawan
    const employeesData = {};
    saturdayData.forEach(item => {
        const employeeName = item.nama;
        if (!employeesData[employeeName]) {
            employeesData[employeeName] = [];
        }
        employeesData[employeeName].push(item);
    });
    
    const employees = Object.keys(employeesData);
    
    const modalHtml = `
        <div class="modal" id="saturday-download-modal">
            <div class="modal-content" style="max-width: 700px;">
                <div class="modal-header">
                    <h3><i class="fas fa-calendar-day"></i> Download Lembur Sabtu Per Karyawan</h3>
                    <button class="modal-close" id="close-saturday-download">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="saturday-info">
                        <h4><i class="fas fa-info-circle"></i> Aturan Hari Sabtu</h4>
                        <p><strong>Semua karyawan:</strong> Kerja 6 jam normal | <strong>K3:</strong> Bekerja 07:00-22:00</p>
                    </div>
                    
                    <div class="saturday-stats" style="margin-bottom: 1.5rem;">
                        <div style="display: flex; gap: 1rem; justify-content: space-around; text-align: center;">
                            <div>
                                <h5 style="color: #666; font-size: 0.9rem;">Total Karyawan</h5>
                                <p style="font-size: 1.5rem; font-weight: bold; color: #e67e22;">${employees.length}</p>
                            </div>
                            <div>
                                <h5 style="color: #666; font-size: 0.9rem;">Total Data</h5>
                                <p style="font-size: 1.5rem; font-weight: bold; color: #e67e22;">${saturdayData.length}</p>
                            </div>
                        </div>
                    </div>
                    
                    <div style="margin-bottom: 1rem;">
                        <h4 style="color: #2c3e50; margin-bottom: 0.5rem;">Pilih Karyawan:</h4>
                        <div class="saturday-employee-list" style="max-height: 250px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px; padding: 0.5rem;">
                            ${employees.map(employee => {
                                const employeeData = employeesData[employee];
                                const totalJam = employeeData.reduce((sum, item) => sum + item.durasi, 0);
                                const totalLembur = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
                                const totalLemburBulat = roundOvertimeHours(totalLembur);
                                const category = employeeCategories[employee] || 'STAFF';
                                const rate = overtimeRates[category];
                                const totalGaji = totalLemburBulat * rate;
                                
                                return `
                                    <div class="saturday-employee-item" data-employee="${employee}" style="padding: 0.75rem; border-bottom: 1px solid #eee; cursor: pointer; transition: background-color 0.3s;">
                                        <div style="display: flex; justify-content: space-between; align-items: center;">
                                            <div>
                                                <strong>${employee}</strong>
                                                <div style="font-size: 0.85rem; color: #666;">
                                                    <span>${employeeData.length} hari kerja</span>
                                                    <span style="margin-left: 1rem;">Kategori: ${category}</span>
                                                </div>
                                            </div>
                                            <div style="text-align: right;">
                                                <div style="font-weight: bold; color: #e74c3c;">
                                                    ${totalLemburBulat} jam lembur
                                                </div>
                                                <div style="font-size: 0.85rem; color: #27ae60;">
                                                    ${formatCurrency(totalGaji)}
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                `;
                            }).join('')}
                        </div>
                    </div>
                    
                    <div class="download-options-saturday-employee" style="margin-top: 1.5rem;">
                        <h4 style="color: #2c3e50; margin-bottom: 0.5rem;">Pilih Format Download:</h4>
                        <div class="option-cards-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 1rem;">
                            <div class="option-card" id="option-saturday-employee-detail">
                                <div class="option-icon">
                                    <i class="fas fa-file-excel" style="color: #e67e22; font-size: 2rem;"></i>
                                </div>
                                <div class="option-content">
                                    <h4>Detail Harian</h4>
                                    <p>Download detail lembur per hari</p>
                                    <ul>
                                        <li>Data detail tiap hari kerja</li>
                                        <li>Jam masuk & keluar</li>
                                        <li>Perhitungan lengkap</li>
                                    </ul>
                                </div>
                            </div>
                            
                            <div class="option-card" id="option-saturday-employee-summary">
                                <div class="option-icon">
                                    <i class="fas fa-chart-bar" style="color: #34495e; font-size: 2rem;"></i>
                                </div>
                                <div class="option-content">
                                    <h4>Rekap Bulanan</h4>
                                    <p>Download rekap bulanan</p>
                                    <ul>
                                        <li>Total jam lembur</li>
                                        <li>Total gaji lembur</li>
                                        <li>Format untuk payroll</li>
                                    </ul>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="cancel-saturday-download">Batal</button>
                    <button class="btn btn-primary" id="download-selected-saturday-employee" disabled>
                        <i class="fas fa-download"></i> Download
                    </button>
                </div>
            </div>
        </div>
    `;
    
    // Tambahkan modal ke body jika belum ada
    let modal = document.getElementById('saturday-download-modal');
    if (!modal) {
        document.body.insertAdjacentHTML('beforeend', modalHtml);
        modal = document.getElementById('saturday-download-modal');
        
        // Variabel untuk menyimpan pilihan
        let selectedEmployee = null;
        let selectedFormat = null;
        
        // Setup event listeners untuk modal
        document.getElementById('close-saturday-download').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        document.getElementById('cancel-saturday-download').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        // Event untuk memilih karyawan
        document.querySelectorAll('.saturday-employee-item').forEach(item => {
            item.addEventListener('click', function() {
                // Hapus seleksi sebelumnya
                document.querySelectorAll('.saturday-employee-item').forEach(el => {
                    el.style.backgroundColor = '';
                    el.style.borderLeft = '';
                });
                
                // Tandai karyawan yang dipilih
                this.style.backgroundColor = '#fff3cd';
                this.style.borderLeft = '3px solid #e67e22';
                
                selectedEmployee = this.getAttribute('data-employee');
                
                // Enable tombol download
                const downloadBtn = document.getElementById('download-selected-saturday-employee');
                if (downloadBtn && selectedEmployee && selectedFormat) {
                    downloadBtn.disabled = false;
                }
            });
        });
        
        // Event untuk memilih format
        document.querySelectorAll('.option-card').forEach(card => {
            card.addEventListener('click', function() {
                // Hapus seleksi sebelumnya
                document.querySelectorAll('.option-card').forEach(el => {
                    el.style.backgroundColor = '';
                    el.style.boxShadow = '';
                });
                
                // Tandai format yang dipilih
                this.style.backgroundColor = '#f8f9fa';
                this.style.boxShadow = '0 0 0 2px #e67e22';
                
                selectedFormat = this.id.replace('option-saturday-employee-', '');
                
                // Enable tombol download
                const downloadBtn = document.getElementById('download-selected-saturday-employee');
                if (downloadBtn && selectedEmployee && selectedFormat) {
                    downloadBtn.disabled = false;
                }
            });
        });
        
        // Event untuk tombol download
        document.getElementById('download-selected-saturday-employee').addEventListener('click', () => {
            if (!selectedEmployee || !selectedFormat) {
                showNotification('Pilih karyawan dan format download terlebih dahulu', 'warning');
                return;
            }
            
            downloadSaturdayPerEmployeeData(selectedEmployee, selectedFormat, employeesData[selectedEmployee]);
            modal.classList.remove('active');
        });
        
        // Close modal ketika klik di luar
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });
    }
    
    modal.classList.add('active');
}

// Fungsi untuk download data Sabtu per karyawan
function downloadSaturdayPerEmployeeData(employeeName, formatType, employeeData) {
    if (!employeeName || !formatType || !employeeData || employeeData.length === 0) {
        showNotification('Data tidak valid', 'error');
        return;
    }
    
    try {
        const category = employeeCategories[employeeName] || 'STAFF';
        const rate = overtimeRates[category];
        
        if (formatType === 'detail') {
            // Download detail harian
            const exportData = employeeData.map((item, index) => {
                const jamLemburBulat = roundOvertimeHours(item.jamLemburDesimal);
                const gajiLembur = jamLemburBulat * rate;
                
                return {
                    'No': index + 1,
                    'Tanggal': formatDate(item.tanggal),
                    'Hari': 'Sabtu',
                    'Jam Masuk': item.jamMasuk,
                    'Jam Keluar': item.jamKeluar,
                    'Durasi Kerja': item.durasiFormatted,
                    'Jam Normal (6 jam)': item.jamNormalFormatted,
                    'Jam Lembur (Bulat)': jamLemburBulat,
                    'Jam Lembur (Desimal)': item.jamLemburDesimal.toFixed(2),
                    'Rate per Jam': formatCurrency(rate),
                    'Gaji Lembur': formatCurrency(gajiLembur),
                    'Keterangan': item.keterangan
                };
            });
            
            // Tambahkan total row
            const totalLemburBulat = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
            const totalLemburDesimal = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
            const totalGaji = totalLemburBulat * rate;
            
            exportData.push({
                'No': '',
                'Tanggal': 'TOTAL',
                'Hari': '',
                'Jam Masuk': '',
                'Jam Keluar': '',
                'Durasi Kerja': '',
                'Jam Normal (6 jam)': '',
                'Jam Lembur (Bulat)': totalLemburBulat,
                'Jam Lembur (Desimal)': totalLemburDesimal.toFixed(2),
                'Rate per Jam': formatCurrency(rate),
                'Gaji Lembur': formatCurrency(totalGaji),
                'Keterangan': `Total ${employeeData.length} hari Sabtu`
            });
            
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, `Lembur Sabtu ${employeeName}`);
            
            XLSX.writeFile(workbook, `lembur_sabtu_detail_${employeeName.replace(/\s+/g, '_').toLowerCase()}_${new Date().toISOString().split('T')[0]}.xlsx`);
            
        } else if (formatType === 'summary') {
            // Download rekap bulanan
            const totalLemburBulat = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
            const totalLemburDesimal = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
            const totalGaji = totalLemburBulat * rate;
            
            const exportData = [{
                'Nama Karyawan': employeeName,
                'Kategori': category,
                'Rate per Jam': formatCurrency(rate),
                'Total Hari Sabtu': employeeData.length,
                'Total Jam Lembur (Desimal)': totalLemburDesimal.toFixed(2),
                'Total Jam Lembur (Bulat)': totalLemburBulat,
                'Total Gaji Lembur': formatCurrency(totalGaji),
                'Rata-rata per Hari': formatCurrency(totalGaji / employeeData.length),
                'Periode': `${employeeData[0].tanggal} - ${employeeData[employeeData.length - 1].tanggal}`
            }];
            
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, `Rekap Sabtu ${employeeName}`);
            
            XLSX.writeFile(workbook, `rekap_lembur_sabtu_${employeeName.replace(/\s+/g, '_').toLowerCase()}_${new Date().toISOString().split('T')[0]}.xlsx`);
        }
        
        showNotification(`Data Sabtu ${employeeName} berhasil diunduh!`, 'success');
        
    } catch (error) {
        console.error('Error downloading Saturday employee data:', error);
        showNotification('Gagal mengunduh data Sabtu karyawan', 'error');
    }
}

// ============================
// FUNGSI PROCESS EXCEL YANG DIPERBAIKI
// ============================

// Process Excel file
function processExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                console.log('Processing Excel file...');
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
                reject(new Error('Gagal membaca file Excel. Pastikan format file benar.'));
            }
        };
        
        reader.onerror = function() {
            reject(new Error('Gagal membaca file'));
        };
        
        reader.readAsArrayBuffer(file);
    });
}

// Process your specific Excel format (kolom E dan F)
function processYourExcelFormat(rawData) {
    const result = [];
    
    console.log('Processing Excel format, total rows:', rawData.length);
    
    // Cari header row untuk kolom E dan F
    for (let i = 0; i < rawData.length; i++) {
        const row = rawData[i];
        
        // Cek jika row memiliki minimal 6 kolom (A-F)
        if (row && row.length >= 6) {
            const nama = row[4];  // Kolom E (index 4)
            const waktu = row[5]; // Kolom F (index 5)
            
            if (nama && waktu) {
                const { date, time } = parseDateTime(waktu.toString());
                
                if (date && time) {
                    result.push({
                        nama: nama.toString().trim(),
                        tanggal: date,
                        waktu: time,
                        rawDatetime: waktu
                    });
                }
            }
        }
    }
    
    console.log('Processed', result.length, 'records');
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
            const durasi = calculateHoursWithFridayRules(jamMasuk, jamKeluar, tanggal);
            
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

// Hitung total jam lembur per karyawan DENGAN PEMBULATAN DAN BATAS 7 JAM
function calculateOvertimeSummary(data) {
    const summary = {};
    
    data.forEach(item => {
        const employeeName = item.nama;
        const overtimeHours = item.jamLemburDesimal || 0;
        
        if (!summary[employeeName]) {
            summary[employeeName] = {
                nama: employeeName,
                totalLembur: 0,
                totalLemburDesimal: 0,
                totalLemburBulat: 0,
                kategori: employeeCategories[employeeName] || 'STAFF',
                rate: overtimeRates[employeeCategories[employeeName]] || 10000,
                fridayOvertime: 0,
                otherDaysOvertime: 0
            };
        }
        
        summary[employeeName].totalLemburDesimal += overtimeHours;
        
        if (item.isFriday) {
            summary[employeeName].fridayOvertime += overtimeHours;
        } else {
            summary[employeeName].otherDaysOvertime += overtimeHours;
        }
    });
    
    // Hitung total dengan pembulatan DAN BATAS 7 JAM PER HARI
    Object.keys(summary).forEach(employee => {
        const record = summary[employee];
        
        // Hitung versi bulat DENGAN BATAS 7 JAM
        record.totalLemburBulat = roundOvertimeHours(record.totalLemburDesimal);
        record.fridayLemburBulat = roundOvertimeHours(record.fridayOvertime);
        record.otherDaysLemburBulat = roundOvertimeHours(record.otherDaysOvertime);
        
        // Versi lama (desimal) untuk kompatibilitas
        record.totalLembur = record.totalLemburDesimal;
        
        // Hitung gaji dengan versi BULAT
        record.totalGaji = record.totalLemburBulat * record.rate;
        record.fridayGaji = record.fridayLemburBulat * record.rate;
        record.otherDaysGaji = record.otherDaysLemburBulat * record.rate;
        
        // Format untuk display
        record.totalGajiFormatted = formatCurrency(record.totalGaji);
        record.fridayGajiFormatted = formatCurrency(record.fridayGaji);
        record.otherDaysGajiFormatted = formatCurrency(record.otherDaysGaji);
        
        // Format jam dengan pembulatan DAN INFO BATAS
        if (record.totalLemburDesimal > 7 * data.filter(item => item.nama === employee).length) {
            record.totalLemburFormatted = `${record.totalLemburBulat} jam (dibatasi maksimal 7 jam/hari)`;
        } else {
            record.totalLemburFormatted = `${record.totalLemburBulat} jam (${record.totalLemburDesimal.toFixed(2)})`;
        }
        
        record.fridayLemburFormatted = `${record.fridayLemburBulat} jam (${record.fridayOvertime.toFixed(2)})`;
        record.otherDaysLemburFormatted = `${record.otherDaysLemburBulat} jam (${record.otherDaysOvertime.toFixed(2)})`;
    });
    
    return Object.values(summary);
}

// ============================
// FUNGSI UPLOAD HANDLER YANG DIPERBAIKI
// ============================

// Handle file selection
async function handleFileSelect(event) {
    console.log('handleFileSelect called');
    
    const fileInput = event.target;
    const file = fileInput.files[0];
    
    if (!file) {
        console.log('No file selected');
        return;
    }
    
    console.log('File selected:', file.name, 'Type:', file.type, 'Size:', file.size);
    
    // Validasi file
    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
        showNotification('Hanya file Excel (.xlsx, .xls) atau CSV yang didukung', 'error');
        return;
    }
    
    if (file.size > 10 * 1024 * 1024) {
        showNotification('File terlalu besar. Maksimal 10MB', 'error');
        return;
    }
    
    currentFile = file;
    showFilePreview(file);
    simulateUploadProgress();
    
    try {
        // Disable process button selama upload
        if (processBtn) {
            processBtn.disabled = true;
            processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Memproses...';
        }
        
        const data = await processExcelFile(file);
        originalData = data;
        
        console.log('File processed successfully:', data.length, 'records');
        
        updateSidebarStats(data);
        showNotification(`File berhasil diunggah! ${data.length} data ditemukan.`, 'success');
        
        // Enable process button
        if (processBtn) {
            processBtn.disabled = false;
            processBtn.innerHTML = '<i class="fas fa-calculator"></i> Hitung Lembur';
        }
        
        previewUploadedData(data);
        
        // Stop progress bar
        if (uploadProgressInterval) {
            clearInterval(uploadProgressInterval);
            const progressBar = document.getElementById('upload-progress');
            const progressText = document.getElementById('progress-text');
            if (progressBar) progressBar.style.width = '100%';
            if (progressText) progressText.textContent = '100% - Selesai';
        }
        
    } catch (error) {
        console.error('Error processing file:', error);
        showNotification('Gagal memproses file: ' + error.message, 'error');
        
        // Reset state
        if (processBtn) {
            processBtn.disabled = true;
            processBtn.innerHTML = '<i class="fas fa-calculator"></i> Hitung Lembur';
        }
        
        cancelUpload();
    }
}

// Show preview of uploaded data
function previewUploadedData(data) {
    const filePreview = document.getElementById('file-preview');
    if (!filePreview) return;
    
    const existingPreview = filePreview.querySelector('.data-preview');
    if (existingPreview) {
        existingPreview.remove();
    }
    
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
        const isFridayDay = isFriday(record.tanggal);
        const isSaturdayDay = isSaturday(record.tanggal);
        let dayBadge = '';
        
        if (isFridayDay) dayBadge = '<span style="color: #9b59b6; font-weight: bold;">(JUMAT)</span>';
        if (isSaturdayDay) dayBadge = '<span style="color: #e67e22; font-weight: bold;">(SABTU)</span>';
        
        previewHtml += `
            <div style="margin-bottom: 0.5rem; padding-bottom: 0.5rem; border-bottom: 1px solid #eee;">
                <strong>${record.nama}</strong> - ${formatDate(record.tanggal)} ${dayBadge}<br>
                Masuk: ${record.jamMasuk} | Pulang: ${record.jamKeluar || '-'} 
                ${record.durasi ? `| Durasi: ${record.durasi.toFixed(2)} jam` : ''}
            </div>
        `;
    }
    
    previewHtml += `</div>`;
    
    previewDiv.innerHTML = previewHtml;
    filePreview.appendChild(previewDiv);
}

// Show file preview
function showFilePreview(file) {
    const fileName = document.getElementById('file-name');
    const fileSize = document.getElementById('file-size');
    const fileDate = document.getElementById('file-date');
    
    if (fileName) fileName.textContent = file.name;
    if (fileSize) fileSize.textContent = formatFileSize(file.size);
    if (fileDate) fileDate.textContent = new Date(file.lastModified).toLocaleDateString('id-ID');
    
    const filePreview = document.getElementById('file-preview');
    if (filePreview) {
        filePreview.style.display = 'block';
        filePreview.classList.add('active');
    }
    
    const uploadArea = document.getElementById('upload-area');
    if (uploadArea) {
        uploadArea.style.display = 'none';
    }
}

// Simulate upload progress
function simulateUploadProgress() {
    if (uploadProgressInterval) clearInterval(uploadProgressInterval);
    
    const progressBar = document.getElementById('upload-progress');
    const progressText = document.getElementById('progress-text');
    
    if (!progressBar || !progressText) return;
    
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

// ============================
// MAIN APPLICATION FUNCTIONS
// ============================

// Initialize application
function initializeApp() {
    console.log('=== INITIALIZING APP ===');
    
    // Debug elements
    debugElements();
    
    // Setup loading state
    if (loadingScreen) {
        loadingScreen.style.display = 'flex';
    }
    
    if (mainContainer) {
        mainContainer.style.opacity = '0';
        mainContainer.style.display = 'none';
    }
    
    try {
        // 1. Set current date
        const currentDate = document.getElementById('current-date');
        if (currentDate) {
            const now = new Date();
            const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
            currentDate.textContent = now.toLocaleDateString('id-ID', options);
        }
        
        // 2. Setup event listeners
        setupEventListeners();
        
        // 3. Hide loading screen after delay
        setTimeout(() => {
            console.log('Hiding loading screen...');
            
            if (loadingScreen) {
                loadingScreen.style.opacity = '0';
                loadingScreen.style.transition = 'opacity 0.5s ease';
                
                setTimeout(() => {
                    loadingScreen.style.display = 'none';
                    
                    if (mainContainer) {
                        mainContainer.style.display = 'block';
                        setTimeout(() => {
                            mainContainer.style.opacity = '1';
                            mainContainer.style.transition = 'opacity 0.5s ease';
                            mainContainer.classList.add('loaded');
                        }, 50);
                    }
                    
                    console.log('App initialized successfully!');
                    showNotification('Sistem Pengolahan Data Presensi siap digunakan', 'success');
                    
                }, 500);
            }
        }, 1500);
        
    } catch (error) {
        console.error('Error in initializeApp:', error);
        
        // Emergency fallback
        if (loadingScreen) {
            loadingScreen.style.display = 'none';
        }
        if (mainContainer) {
            mainContainer.style.display = 'block';
            mainContainer.style.opacity = '1';
        }
        
        showNotification('Aplikasi dijalankan dalam mode terbatas', 'warning');
    }
}

// Process data (HITUNG LEMBUR SAJA)
function processData() {
    if (originalData.length === 0) {
        showNotification('Tidak ada data untuk diproses. Silakan upload file Excel terlebih dahulu.', 'warning');
        return;
    }
    
    currentWorkHours = parseFloat(document.getElementById('work-hours').value) || 8;
    
    if (processBtn) {
        processBtn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Menghitung Lembur...';
        processBtn.disabled = true;
    }
    
    setTimeout(() => {
        try {
            // Hitung lembur per hari dengan jam kerja yang diatur
            processedData = calculateOvertimePerDay(originalData, currentWorkHours);
            
            displayResults(processedData);
            createCharts(processedData);
            
            if (resultsSection) {
                resultsSection.style.display = 'block';
                resultsSection.scrollIntoView({ behavior: 'smooth' });
            }
            
            showNotification(`Perhitungan lembur selesai! ${processedData.length} data diproses.`, 'success');
            
        } catch (error) {
            console.error('Error processing data:', error);
            showNotification('Terjadi kesalahan saat menghitung lembur: ' + error.message, 'error');
        } finally {
            if (processBtn) {
                processBtn.innerHTML = '<i class="fas fa-calculator"></i> Hitung Lembur';
                processBtn.disabled = false;
            }
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
    const totalLemburDesimal = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    const totalLemburBulat = roundOvertimeHours(totalLemburDesimal);
    
    const totalKaryawanElem = document.getElementById('total-karyawan');
    const totalHariElem = document.getElementById('total-hari');
    const totalLemburElem = document.getElementById('total-lembur');
    const totalGajiElem = document.getElementById('total-gaji');
    
    if (totalKaryawanElem) totalKaryawanElem.textContent = totalKaryawan;
    if (totalHariElem) totalHariElem.textContent = totalHari;
    if (totalLemburElem) totalLemburElem.textContent = formatOvertimeInfo(totalLemburDesimal);
    
    // Hitung total gaji
    let totalGaji = 0;
    const summary = calculateOvertimeSummary(data);
    summary.forEach(emp => {
        totalGaji += emp.totalGaji;
    });
    
    if (totalGajiElem) {
        totalGajiElem.innerHTML = `
            ${formatCurrency(totalGaji)}<br>
            <small style="font-size: 0.8rem;">
                (${totalLemburBulat} jam lembur)
            </small>
        `;
    }
}

// Tampilkan data original
function displayOriginalTable(data) {
    const tbody = document.getElementById('original-table-body');
    if (!tbody) return;
    
    tbody.innerHTML = '';
    
    data.forEach((item, index) => {
        const row = document.createElement('tr');
        const isFridayDay = isFriday(item.tanggal);
        const isSaturdayDay = isSaturday(item.tanggal);
        
        let dayBadge = '';
        if (isFridayDay) dayBadge = '<br><small style="color: #9b59b6; font-weight: bold;">(JUMAT)</small>';
        if (isSaturdayDay) dayBadge = '<br><small style="color: #e67e22; font-weight: bold;">(SABTU)</small>';
        
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>
                ${formatDate(item.tanggal)}
                ${dayBadge}
            </td>
            <td>${item.jamMasuk}</td>
            <td>${item.jamKeluar || '-'}</td>
            <td>${item.durasi ? item.durasi.toFixed(2) + ' jam' : '-'}</td>
        `;
        
        if (isFridayDay) {
            row.classList.add('friday-row');
        } else if (isSaturdayDay) {
            row.classList.add('saturday-row');
        }
        
        tbody.appendChild(row);
    });
}

// Tampilkan data terproses
function displayProcessedTable(data) {
    const tbody = document.getElementById('processed-table-body');
    if (!tbody) {
        console.error('Element #processed-table-body tidak ditemukan');
        return;
    }
    
    tbody.innerHTML = '';
    
    data.forEach((item, index) => {
        const row = document.createElement('tr');
        const jamLemburBulat = roundOvertimeHours(item.jamLemburDesimal);
        const isCapped = item.jamLemburDesimal > 7;
        
        // Tentukan badge berdasarkan hari
        let dayBadge = '';
        if (item.isFriday) dayBadge = '<br><small style="color: #9b59b6; font-weight: bold;">(JUMAT)</small>';
        if (item.isSaturday) dayBadge = '<br><small style="color: #e67e22; font-weight: bold;">(SABTU)</small>';
        
        // Tentukan warna teks berdasarkan jumlah lembur
        let overtimeColor = item.jamLemburDesimal > 0 ? '#e74c3c' : '#27ae60';
        if (isCapped) overtimeColor = '#e67e22';
        
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>
                ${formatDate(item.tanggal)}
                ${dayBadge}
            </td>
            <td>${item.jamMasuk}</td>
            <td>${item.jamKeluar}</td>
            <td>${item.durasiFormatted}</td>
            <td>${item.jamNormalFormatted}</td>
            <td style="color: ${overtimeColor}; font-weight: bold;">
                ${jamLemburBulat > 0 ? jamLemburBulat + ' jam' : item.jamLembur}
                ${isCapped ? 
                    '<br><small style="color: #e67e22; font-weight: bold;">(DIBATASI MAKSIMAL 7 JAM)</small>' : 
                    ''}
                ${item.isFriday && item.jamLemburDesimal > 0 ? 
                    '<br><small style="color: #9b59b6;">(pulang > 15:00)</small>' : 
                    ''}
                ${item.isSaturday && item.jamLemburDesimal > 0 ? 
                    '<br><small style="color: #e67e22;">(kerja > 6 jam)</small>' : 
                    ''}
                ${item.jamLemburDesimal > 0 && !isCapped ? 
                    `<br><small style="color: #666; font-size: 0.8rem;">(${item.jamLemburDesimal.toFixed(2)} jam desimal)</small>` : 
                    ''}
                ${isCapped ? 
                    `<br><small style="color: #666; font-size: 0.8rem;">(${item.jamLemburDesimal.toFixed(2)} jam → dibatasi 7 jam)</small>` : 
                    ''}
            </td>
            <td>
                <span style="color: ${overtimeColor}; font-size: 0.85rem;">
                    ${item.keterangan}
                </span>
            </td>
        `;
        
        if (item.isFriday) {
            row.classList.add('friday-row');
        } else if (item.isSaturday) {
            row.classList.add('saturday-row');
        }
        
        if (isCapped) {
            row.classList.add('overtime-capped-row');
        }
        
        tbody.appendChild(row);
    });
}

// Display summaries
function displaySummaries(data) {
    const employeeSummary = calculateOvertimeSummary(data);
    const employeeSummaryElem = document.getElementById('employee-summary');
    const financialSummary = document.getElementById('financial-summary');
    
    if (employeeSummaryElem) {
        let html = '<div style="max-height: 300px; overflow-y: auto;">';
        employeeSummary.forEach(emp => {
            if (emp.totalLemburDesimal > 0) {
                const isCapped = emp.totalLemburDesimal > 7 * data.filter(item => item.nama === emp.nama).length;
                html += `
                    <div style="margin-bottom: 1rem; padding: 0.75rem; background: #f8f9fa; border-radius: 8px;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div>
                                <strong>${emp.nama}</strong>
                                <br><small style="color: #666;">Kategori: ${emp.kategori}</small>
                            </div>
                            <div style="text-align: right;">
                                <div style="color: ${isCapped ? '#e67e22' : '#e74c3c'}; font-weight: bold;">${emp.totalLemburBulat} jam lembur${isCapped ? ' (MAKS)' : ''}</div>
                                <div style="font-size: 0.9rem;">(${emp.totalLemburDesimal.toFixed(2)} jam desimal)</div>
                            </div>
                        </div>
                        <div style="display: flex; justify-content: space-between; margin-top: 0.5rem; font-size: 0.85rem;">
                            <div>
                                Jumat: ${emp.fridayLemburBulat} jam<br>
                                Senin-Kamis: ${emp.otherDaysLemburBulat} jam
                            </div>
                            <div style="font-weight: bold; color: #27ae60;">
                                ${emp.totalGajiFormatted}
                            </div>
                        </div>
                    </div>
                `;
            }
        });
        html += '</div>';
        employeeSummaryElem.innerHTML = html;
    }
    
    // Financial summary
    const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
    const totalLemburDesimal = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    const totalLemburBulat = roundOvertimeHours(totalLemburDesimal);
    const totalNormal = data.reduce((sum, item) => sum + item.jamNormal, 0);
    
    let totalGajiAll = 0;
    employeeSummary.forEach(emp => {
        totalGajiAll += emp.totalGaji;
    });
    
    if (financialSummary) {
        financialSummary.innerHTML = `
            <div><strong>ATURAN PERHITUNGAN:</strong></div>
            <div style="font-size: 0.85rem; color: #666; margin-bottom: 1rem; padding: 0.75rem; background: #f8f9fa; border-radius: 4px;">
                • Jam kerja normal: 8 jam/hari (6 jam untuk Sabtu)<br>
                • Lembur dihitung setelah jam normal<br>
                • Minimal lembur: 10 menit<br>
                • Maksimal lembur: 7 jam/hari<br>
                • Jumat: pulang normal jam 15:00<br>
                • Sabtu: kerja 6 jam normal, K3: 07:00-22:00
            </div>
            
            <div style="margin-top: 1rem;">
                <div>Konfigurasi Jam Kerja: <strong>${currentWorkHours} jam/hari</strong></div>
                <div>Total Entri Data: <strong>${data.length} hari</strong></div>
                <div>Total Jam Kerja: <strong>${formatHoursToDisplay(totalJam)}</strong></div>
                <div>Total Jam Normal: <strong>${formatHoursToDisplay(totalNormal)}</strong></div>
                <div style="color: ${totalLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-weight: bold;">
                    Total Jam Lembur: <strong>${formatOvertimeInfo(totalLemburDesimal)}</strong>
                </div>
            </div>
            <div style="border-top: 2px solid #3498db; padding-top: 0.5rem; margin-top: 0.5rem;">
                <strong>Total Gaji Lembur:</strong><br>
                <div style="font-size: 1.5rem; font-weight: bold; color: #27ae60; margin-top: 0.5rem;">
                    ${formatCurrency(totalGajiAll)}
                </div>
                <div style="font-size: 0.85rem; color: #666; margin-top: 0.5rem;">
                    <i class="fas fa-info-circle"></i> Gaji dihitung berdasarkan jam lembur yang telah dibulatkan
                </div>
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
    const hoursCtx = document.getElementById('hoursChart');
    if (hoursCtx) {
        hoursChart = new Chart(hoursCtx.getContext('2d'), {
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
                        }
                    }
                }
            }
        });
    }
    
    // Chart Pie untuk komposisi jam kerja
    const salaryCtx = document.getElementById('salaryChart');
    if (salaryCtx) {
        const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
        const totalNormal = data.reduce((sum, item) => sum + item.jamNormal, 0);
        const totalLembur = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
        
        salaryChart = new Chart(salaryCtx.getContext('2d'), {
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
                    }
                }
            }
        });
    }
}

// Switch tabs
function switchTab(tabId) {
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
    
    const excelFileInput = document.getElementById('excel-file');
    const filePreview = document.getElementById('file-preview');
    const uploadArea = document.getElementById('upload-area');
    const processBtn = document.getElementById('process-data');
    
    if (excelFileInput) excelFileInput.value = '';
    if (filePreview) {
        filePreview.style.display = 'none';
        filePreview.classList.remove('active');
    }
    if (uploadArea) uploadArea.style.display = 'block';
    if (processBtn) processBtn.disabled = true;
    
    originalData = [];
    processedData = [];
    currentFile = null;
    
    const resultsSection = document.getElementById('results-section');
    if (resultsSection) resultsSection.style.display = 'none';
    
    showNotification('Upload dibatalkan', 'info');
}

// Reset configuration
function resetConfig() {
    document.getElementById('work-hours').value = '8';
    currentWorkHours = 8;
    showNotification('Konfigurasi direset ke 8 jam kerja', 'info');
}

// Download template
function downloadTemplate() {
    const templateData = [
        ['Nama', 'Tanggal', 'Jam Masuk', 'Jam Keluar'],
        ['Windy', '01/11/2025 07:00', '', ''],
        ['Windy', '01/11/2025 16:30', '', ''],
        ['Bu Ati', '01/11/2025 07:00', '', ''],
        ['Bu Ati', '01/11/2025 15:30', '', ''],
        ['Pak Saji', '02/11/2025 07:00', '', ''],
        ['Pak Saji', '02/11/2025 17:00', '', '']
    ];
    
    const worksheet = XLSX.utils.aoa_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Template');
    
    XLSX.writeFile(workbook, 'template_presensi.xlsx');
    showNotification('Template berhasil diunduh', 'success');
}

// Download report
async function downloadReport(type) {
    if (type === 'original' && originalData.length === 0) {
        showNotification('Tidak ada data asli untuk diunduh', 'warning');
        return;
    }
    
    if ((type === 'processed' || type === 'both') && processedData.length === 0) {
        showNotification('Data belum diproses', 'warning');
        return;
    }
    
    try {
        if (type === 'original') {
            const exportData = prepareExportData(originalData);
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Asli');
            XLSX.writeFile(workbook, 'data_presensi_asli.xlsx');
            showNotification('Data asli berhasil diunduh', 'success');
        } else if (type === 'processed') {
            const exportData = prepareExportData(processedData);
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Lembur');
            XLSX.writeFile(workbook, 'data_lembur_harian.xlsx');
            showNotification('Data lembur berhasil diunduh', 'success');
        } else if (type === 'both') {
            // Original data
            const worksheet1 = XLSX.utils.json_to_sheet(prepareExportData(originalData));
            const worksheet2 = XLSX.utils.json_to_sheet(prepareExportData(processedData));
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet1, 'Data Asli');
            XLSX.utils.book_append_sheet(workbook, worksheet2, 'Data Lembur');
            XLSX.writeFile(workbook, 'data_lengkap.xlsx');
            showNotification('Kedua file berhasil diunduh', 'success');
        }
    } catch (error) {
        console.error('Error downloading report:', error);
        showNotification('Gagal mengunduh laporan', 'error');
    }
}

// Prepare data for export
function prepareExportData(data) {
    if (data.length === 0) return [];
    
    const hasOvertimeData = data[0].jamLemburDesimal !== undefined;
    
    if (hasOvertimeData) {
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Hari': item.hari,
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi Kerja': item.durasiFormatted,
            'Jam Normal': item.jamNormalFormatted,
            'Jam Lembur': roundOvertimeHours(item.jamLemburDesimal) + ' jam',
            'Keterangan': item.keterangan
        }));
    } else {
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Hari': getDayName(item.tanggal),
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi': item.durasi ? item.durasiFormatted : ''
        }));
    }
}

// Update sidebar stats
function updateSidebarStats(data) {
    const uniqueEmployees = new Set(data.map(item => item.nama)).size;
    const uniqueDates = new Set(data.map(item => item.tanggal)).size;
    
    const statEmployees = document.getElementById('stat-employees');
    const statAttendance = document.getElementById('stat-attendance');
    
    if (statEmployees) statEmployees.textContent = uniqueEmployees;
    if (statAttendance) statAttendance.textContent = uniqueDates;
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
    // Remove existing notifications
    const existingNotifications = document.querySelectorAll('.notification');
    existingNotifications.forEach(notification => {
        if (notification.parentNode) {
            notification.parentNode.removeChild(notification);
        }
    });
    
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 1rem 1.5rem;
        background: ${type === 'success' ? '#2ecc71' : type === 'error' ? '#e74c3c' : type === 'warning' ? '#f39c12' : '#3498db'};
        color: white;
        border-radius: 8px;
        box-shadow: 0 5px 20px rgba(0, 0, 0, 0.15);
        z-index: 2000;
        animation: slideInRight 0.3s ease;
        display: flex;
        align-items: center;
        gap: 0.5rem;
        max-width: 400px;
    `;
    
    const icon = type === 'success' ? 'check-circle' : type === 'error' ? 'exclamation-circle' : type === 'warning' ? 'exclamation-triangle' : 'info-circle';
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

// Add CSS animations for notifications if not exists
if (!document.querySelector('#notification-animations')) {
    const style = document.createElement('style');
    style.id = 'notification-animations';
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

// Emergency fallback jika loading terlalu lama
setTimeout(() => {
    const loadingScreen = document.getElementById('loading-screen');
    const mainContainer = document.getElementById('main-container');
    
    if (loadingScreen && loadingScreen.style.display !== 'none') {
        console.log('Emergency timeout: Hiding loading screen');
        loadingScreen.style.display = 'none';
        if (mainContainer) {
            mainContainer.style.display = 'block';
            mainContainer.style.opacity = '1';
        }
        showNotification('Sistem siap digunakan', 'info');
    }
}, 10000);
