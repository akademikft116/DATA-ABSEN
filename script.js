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
// KONFIGURASI ATURAN
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

// Hari Sabtu special rules
const SATURDAY_RULES = {
    workHours: 6,
    minOvertimeMinutes: 10,
    k3SpecialHours: {
        startTime: '07:00',
        endTime: '22:00',
        isSpecialDay: true
    }
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

// Event Listeners
document.addEventListener('DOMContentLoaded', initializeApp);

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

// Format jam untuk display
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
// FUNGSI UNTUK HARI
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
// FUNGSI PERHITUNGAN JAM
// ============================

// Calculate hours between two time strings
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

// Fungsi untuk menghitung lembur dengan aturan baru
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

// ============================
// FUNGSI UPLOAD DAN EVENT LISTENERS
// ============================

// Inisialisasi event listeners
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
        
        // Setup event listeners untuk upload area
        if (uploadArea) {
            uploadArea.style.cursor = 'pointer';
            
            uploadArea.addEventListener('click', function(e) {
                console.log('Upload area clicked');
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
                    
                    if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
                        showNotification('Hanya file Excel (.xlsx, .xls) atau CSV yang didukung', 'error');
                        return;
                    }
                    
                    if (excelFileInput) {
                        const dataTransfer = new DataTransfer();
                        dataTransfer.items.add(file);
                        excelFileInput.files = dataTransfer.files;
                        
                        const event = new Event('change', { bubbles: true });
                        excelFileInput.dispatchEvent(event);
                    }
                }
            });
        }
        
        // Event untuk tombol browse
        if (browseBtn) {
            browseBtn.addEventListener('click', function(e) {
                e.stopPropagation();
                if (excelFileInput) {
                    excelFileInput.click();
                }
            });
        }
        
        // Event untuk file input
        if (excelFileInput) {
            excelFileInput.addEventListener('change', function(e) {
                if (e.target.files && e.target.files[0]) {
                    handleFileSelect(e);
                }
            });
        }
        
        // Event untuk tombol proses
        if (processBtn) {
            processBtn.addEventListener('click', processData);
        }
        
        // Event untuk tombol cancel
        if (cancelUploadBtn) {
            cancelUploadBtn.addEventListener('click', cancelUpload);
        }
        
        // Event untuk tabs
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', function() {
                const tabId = this.getAttribute('data-tab');
                switchTab(tabId);
            });
        });
        
        // Event untuk modal help
        const helpBtn = document.getElementById('help-btn');
        const closeHelpBtns = document.querySelectorAll('#close-help, #close-help-btn');
        const helpModal = document.getElementById('help-modal');
        
        if (helpBtn) helpBtn.addEventListener('click', () => helpModal?.classList.add('active'));
        closeHelpBtns.forEach(btn => {
            btn.addEventListener('click', () => helpModal?.classList.remove('active'));
        });
        
        // Event untuk template download
        const templateBtn = document.getElementById('template-btn');
        if (templateBtn) templateBtn.addEventListener('click', downloadTemplate);
        
        // Event untuk reset config
        const resetBtn = document.getElementById('reset-config');
        if (resetBtn) resetBtn.addEventListener('click', resetConfig);
        
        // ============================
        // EVENT LISTENERS UNTUK TOMBOL DOWNLOAD
        // ============================
        
        console.log('Setting up download buttons...');
        
        const downloadOriginal = document.getElementById('download-original');
        const downloadProcessed = document.getElementById('download-processed');
        const downloadBoth = document.getElementById('download-both');
        const downloadSalary = document.getElementById('download-salary');
        const downloadSaturday = document.getElementById('download-saturday');
        
        if (downloadOriginal) {
            downloadOriginal.addEventListener('click', function() {
                console.log('Download Original clicked!');
                downloadReport('original');
            });
        }
        
        if (downloadProcessed) {
            downloadProcessed.addEventListener('click', function() {
                console.log('Download Processed clicked!');
                downloadReport('processed');
            });
        }
        
        if (downloadBoth) {
            downloadBoth.addEventListener('click', function() {
                console.log('Download Both clicked!');
                downloadReport('both');
            });
        }
        
        if (downloadSalary) {
            downloadSalary.addEventListener('click', function() {
                console.log('Download Salary clicked!');
                showPerEmployeeDownloadOptions();
            });
        }
        
        if (downloadSaturday) {
            downloadSaturday.addEventListener('click', function() {
                console.log('Download Saturday clicked!');
                if (processedData.length === 0) {
                    showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
                    return;
                }
                downloadSaturdayAllInOne();
            });
        }
        
        console.log('Event listeners setup complete!');
        
    } catch (error) {
        console.error('Error in setupEventListeners:', error);
        showNotification('Error setting up event listeners', 'error');
        
        // Fallback
        setupMinimalEventListeners();
    }
}

// Fallback minimal event listeners
function setupMinimalEventListeners() {
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

// ============================
// FUNGSI PERHITUNGAN LEMBUR
// ============================

// Calculate overtime per day
function calculateOvertimePerDay(data, workHours = 8) {
    const result = data.map(record => {
        // Hitung total jam kerja
        const hoursWorked = calculateHoursWithFridayRules(
            record.jamMasuk, 
            record.jamKeluar, 
            record.tanggal
        );
        
        // Hitung lembur dengan aturan baru
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
        
        // Cek hari
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
// FUNGSI UNTUK DOWNLOAD SATU FILE PER KARYAWAN (SHEET TERPISAH)
// ============================

// Fungsi untuk menampilkan modal pilihan download
function showPerEmployeeDownloadOptions() {
    if (processedData.length === 0) {
        showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
        return;
    }
    
    // Dapatkan daftar karyawan unik
    const employees = [...new Set(processedData.map(item => item.nama))];
    
    // Pisahkan data berdasarkan hari
    const seninKamisData = processedData.filter(item => !item.isFriday && !item.isSaturday);
    const jumatData = processedData.filter(item => item.isFriday);
    const sabtuData = processedData.filter(item => item.isSaturday);
    
    const modalHtml = `
        <div class="modal" id="download-options-modal">
            <div class="modal-content" style="max-width: 700px;">
                <div class="modal-header">
                    <h3><i class="fas fa-download"></i> Pilih Format Download</h3>
                    <button class="modal-close" id="close-download-options">&times;</button>
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
                                <p style="font-size: 1.5rem; font-weight: bold;">${seninKamisData.length}</p>
                            </div>
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Jumat</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${jumatData.length}</p>
                            </div>
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Sabtu</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${sabtuData.length}</p>
                            </div>
                        </div>
                    </div>
                    
                    <div class="download-options-grid" style="display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 1rem;">
                        <!-- Opsi 1: Semua Karyawan dalam Satu File -->
                        <div class="option-card" id="option-all-employees">
                            <div class="option-icon">
                                <i class="fas fa-users" style="color: #217346; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Semua Karyawan</h4>
                                <p>Download data semua karyawan dalam satu file Excel</p>
                                <ul>
                                    <li>Satu file Excel dengan multiple sheet</li>
                                    <li>Setiap karyawan di sheet terpisah</li>
                                    <li>Semua hari (Senin-Kamis, Jumat, Sabtu)</li>
                                    <li>Total: ${employees.length} karyawan</li>
                                </ul>
                            </div>
                        </div>
                        
                        <!-- Opsi 2: Per Hari (Semua Karyawan) -->
                        <div class="option-card" id="option-by-day">
                            <div class="option-icon">
                                <i class="fas fa-calendar-alt" style="color: #3498db; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Per Jenis Hari</h4>
                                <p>Download data berdasarkan jenis hari</p>
                                <ul>
                                    <li>File 1: Senin-Kamis (Semua karyawan)</li>
                                    <li>File 2: Jumat (Semua karyawan)</li>
                                    <li>File 3: Sabtu (Semua karyawan)</li>
                                    <li>${employees.length} karyawan per file</li>
                                </ul>
                            </div>
                        </div>
                        
                        <!-- Opsi 3: Rekap Per Karyawan -->
                        <div class="option-card" id="option-summary">
                            <div class="option-icon">
                                <i class="fas fa-file-invoice-dollar" style="color: #9b59b6; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Rekap Gaji</h4>
                                <p>Download rekap gaji lembur per karyawan</p>
                                <ul>
                                    <li>Total jam lembur per karyawan</li>
                                    <li>Total gaji lembur per karyawan</li>
                                    <li>Breakdown per hari (Senin-Kamis, Jumat, Sabtu)</li>
                                    <li>Format untuk payroll</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="cancel-download-options">Batal</button>
                </div>
            </div>
        </div>
    `;
    
    // Tambahkan modal ke body jika belum ada
    let modal = document.getElementById('download-options-modal');
    if (!modal) {
        document.body.insertAdjacentHTML('beforeend', modalHtml);
        modal = document.getElementById('download-options-modal');
        
        // Setup event listeners untuk modal
        document.getElementById('close-download-options').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        document.getElementById('cancel-download-options').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        // Event untuk opsi download
        document.getElementById('option-all-employees').addEventListener('click', () => {
            downloadAllEmployeesInOneFile();
            modal.classList.remove('active');
        });
        
        document.getElementById('option-by-day').addEventListener('click', () => {
            downloadByDayType();
            modal.classList.remove('active');
        });
        
        document.getElementById('option-summary').addEventListener('click', () => {
            downloadSalarySummary();
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

// Fungsi untuk download semua karyawan dalam satu file (sheet terpisah)
function downloadAllEmployeesInOneFile() {
    if (processedData.length === 0) {
        showNotification('Tidak ada data untuk diunduh', 'warning');
        return;
    }
    
    try {
        // Buat workbook baru
        const workbook = XLSX.utils.book_new();
        const employees = [...new Set(processedData.map(item => item.nama))];
        
        console.log(`Membuat file Excel untuk ${employees.length} karyawan...`);
        
        // Untuk setiap karyawan, buat sheet terpisah
        employees.forEach(employee => {
            const employeeData = processedData.filter(item => item.nama === employee);
            
            // Buat data untuk sheet ini
            const exportData = employeeData.map((item, index) => {
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
            
            // Tambahkan total row
            const totalLemburBulat = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
            const totalLemburDesimal = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
            const category = employeeCategories[employee] || 'STAFF';
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
                'Keterangan': `Total ${employeeData.length} hari kerja`
            });
            
            // Buat worksheet
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            
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
            
            // Tambahkan sheet ke workbook
            // Batasi nama sheet maksimal 31 karakter
            let sheetName = employee.substring(0, 31);
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        });
        
        // Tambahkan sheet summary
        const summaryData = createSummaryData();
        const summaryWorksheet = XLSX.utils.json_to_sheet(summaryData);
        XLSX.utils.book_append_sheet(workbook, summaryWorksheet, 'REKAP');
        
        // Simpan file
        const fileName = `data_lembur_semua_karyawan_${new Date().toISOString().split('T')[0]}.xlsx`;
        XLSX.writeFile(workbook, fileName);
        
        showNotification(`File Excel berhasil diunduh! ${employees.length} karyawan dalam ${employees.length + 1} sheet.`, 'success');
        
    } catch (error) {
        console.error('Error downloading all employees file:', error);
        showNotification('Gagal mengunduh file Excel', 'error');
    }
}

// Fungsi untuk membuat data summary
function createSummaryData() {
    const employees = [...new Set(processedData.map(item => item.nama))];
    const summary = [];
    
    employees.forEach(employee => {
        const employeeData = processedData.filter(item => item.nama === employee);
        
        // Pisahkan berdasarkan hari
        const seninKamis = employeeData.filter(item => !item.isFriday && !item.isSaturday);
        const jumat = employeeData.filter(item => item.isFriday);
        const sabtu = employeeData.filter(item => item.isSaturday);
        
        // Hitung total lembur
        const totalLemburBulat = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
        const totalLemburDesimal = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
        const category = employeeCategories[employee] || 'STAFF';
        const rate = overtimeRates[category];
        const totalGaji = totalLemburBulat * rate;
        
        summary.push({
            'No': summary.length + 1,
            'Nama Karyawan': employee,
            'Kategori': category,
            'Rate per Jam': formatCurrency(rate),
            'Total Hari Kerja': employeeData.length,
            'Hari Senin-Kamis': seninKamis.length,
            'Hari Jumat': jumat.length,
            'Hari Sabtu': sabtu.length,
            'Total Jam Lembur (Bulat)': totalLemburBulat,
            'Total Jam Lembur (Desimal)': totalLemburDesimal.toFixed(2),
            'Total Gaji Lembur': formatCurrency(totalGaji),
            'Rata-rata per Hari': formatCurrency(totalGaji / employeeData.length)
        });
    });
    
    // Tambahkan total row
    const totalEmployees = employees.length;
    const totalHari = processedData.length;
    const totalLemburBulat = processedData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
    const totalLemburDesimal = processedData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    
    // Hitung total gaji
    let totalGajiAll = 0;
    employees.forEach(employee => {
        const employeeData = processedData.filter(item => item.nama === employee);
        const totalLemburBulatEmp = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
        const category = employeeCategories[employee] || 'STAFF';
        const rate = overtimeRates[category];
        totalGajiAll += totalLemburBulatEmp * rate;
    });
    
    summary.push({
        'No': '',
        'Nama Karyawan': 'TOTAL',
        'Kategori': '',
        'Rate per Jam': '',
        'Total Hari Kerja': totalHari,
        'Hari Senin-Kamis': processedData.filter(item => !item.isFriday && !item.isSaturday).length,
        'Hari Jumat': processedData.filter(item => item.isFriday).length,
        'Hari Sabtu': processedData.filter(item => item.isSaturday).length,
        'Total Jam Lembur (Bulat)': totalLemburBulat,
        'Total Jam Lembur (Desimal)': totalLemburDesimal.toFixed(2),
        'Total Gaji Lembur': formatCurrency(totalGajiAll),
        'Rata-rata per Hari': formatCurrency(totalGajiAll / totalHari)
    });
    
    return summary;
}

// Fungsi untuk download berdasarkan jenis hari
function downloadByDayType() {
    if (processedData.length === 0) {
        showNotification('Tidak ada data untuk diunduh', 'warning');
        return;
    }
    
    try {
        // Pisahkan data berdasarkan hari
        const seninKamisData = processedData.filter(item => !item.isFriday && !item.isSaturday);
        const jumatData = processedData.filter(item => item.isFriday);
        const sabtuData = processedData.filter(item => item.isSaturday);
        
        // 1. File Senin-Kamis
        if (seninKamisData.length > 0) {
            const seninKamisWorkbook = XLSX.utils.book_new();
            const employees = [...new Set(seninKamisData.map(item => item.nama))];
            
            employees.forEach(employee => {
                const employeeData = seninKamisData.filter(item => item.nama === employee);
                const exportData = createEmployeeExportData(employeeData, employee);
                const worksheet = XLSX.utils.json_to_sheet(exportData);
                let sheetName = employee.substring(0, 31);
                XLSX.utils.book_append_sheet(seninKamisWorkbook, worksheet, sheetName);
            });
            
            XLSX.writeFile(seninKamisWorkbook, `lembur_senin_kamis_${new Date().toISOString().split('T')[0]}.xlsx`);
        }
        
        // 2. File Jumat
        if (jumatData.length > 0) {
            const jumatWorkbook = XLSX.utils.book_new();
            const employees = [...new Set(jumatData.map(item => item.nama))];
            
            employees.forEach(employee => {
                const employeeData = jumatData.filter(item => item.nama === employee);
                const exportData = createEmployeeExportData(employeeData, employee);
                const worksheet = XLSX.utils.json_to_sheet(exportData);
                let sheetName = employee.substring(0, 31);
                XLSX.utils.book_append_sheet(jumatWorkbook, worksheet, sheetName);
            });
            
            XLSX.writeFile(jumatWorkbook, `lembur_jumat_${new Date().toISOString().split('T')[0]}.xlsx`);
        }
        
        // 3. File Sabtu
        if (sabtuData.length > 0) {
            const sabtuWorkbook = XLSX.utils.book_new();
            const employees = [...new Set(sabtuData.map(item => item.nama))];
            
            employees.forEach(employee => {
                const employeeData = sabtuData.filter(item => item.nama === employee);
                const exportData = createEmployeeExportData(employeeData, employee);
                const worksheet = XLSX.utils.json_to_sheet(exportData);
                let sheetName = employee.substring(0, 31);
                XLSX.utils.book_append_sheet(sabtuWorkbook, worksheet, sheetName);
            });
            
            XLSX.writeFile(sabtuWorkbook, `lembur_sabtu_${new Date().toISOString().split('T')[0]}.xlsx`);
        }
        
        showNotification('File per jenis hari berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error downloading by day type:', error);
        showNotification('Gagal mengunduh file', 'error');
    }
}

// Fungsi untuk membuat data export per karyawan
function createEmployeeExportData(employeeData, employeeName) {
    const exportData = employeeData.map((item, index) => {
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
    
    // Tambahkan total row
    const totalLemburBulat = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
    const totalLemburDesimal = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
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
        'Keterangan': `Total ${employeeData.length} hari kerja`
    });
    
    return exportData;
}

// Fungsi untuk download rekap gaji
function downloadSalarySummary() {
    if (processedData.length === 0) {
        showNotification('Tidak ada data untuk diunduh', 'warning');
        return;
    }
    
    try {
        const summaryData = createSummaryData();
        const worksheet = XLSX.utils.json_to_sheet(summaryData);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Rekap Gaji');
        
        // Set kolom width
        const wscols = [
            {wch: 5},   // No
            {wch: 25},  // Nama Karyawan
            {wch: 10},  // Kategori
            {wch: 15},  // Rate per Jam
            {wch: 12},  // Total Hari Kerja
            {wch: 15},  // Hari Senin-Kamis
            {wch: 10},  // Hari Jumat
            {wch: 10},  // Hari Sabtu
            {wch: 15},  // Total Jam Lembur (Bulat)
            {wch: 15},  // Total Jam Lembur (Desimal)
            {wch: 20},  // Total Gaji Lembur
            {wch: 15}   // Rata-rata per Hari
        ];
        worksheet['!cols'] = wscols;
        
        XLSX.writeFile(workbook, `rekap_gaji_lembur_${new Date().toISOString().split('T')[0]}.xlsx`);
        
        showNotification('Rekap gaji berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error downloading salary summary:', error);
        showNotification('Gagal mengunduh rekap gaji', 'error');
    }
}

// Fungsi untuk download Sabtu semua dalam satu file
function downloadSaturdayAllInOne() {
    const saturdayData = processedData.filter(item => item.isSaturday);
    
    if (saturdayData.length === 0) {
        showNotification('Tidak ada data lembur hari Sabtu', 'info');
        return;
    }
    
    try {
        const workbook = XLSX.utils.book_new();
        const employees = [...new Set(saturdayData.map(item => item.nama))];
        
        employees.forEach(employee => {
            const employeeData = saturdayData.filter(item => item.nama === employee);
            const exportData = createSaturdayExportData(employeeData, employee);
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            let sheetName = employee.substring(0, 31);
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
        });
        
        // Tambahkan sheet summary Sabtu
        const saturdaySummary = createSaturdaySummaryData(saturdayData);
        const summaryWorksheet = XLSX.utils.json_to_sheet(saturdaySummary);
        XLSX.utils.book_append_sheet(workbook, summaryWorksheet, 'REKAP SABTU');
        
        XLSX.writeFile(workbook, `lembur_sabtu_semua_karyawan_${new Date().toISOString().split('T')[0]}.xlsx`);
        
        showNotification(`File lembur Sabtu berhasil diunduh! ${employees.length} karyawan.`, 'success');
        
    } catch (error) {
        console.error('Error downloading Saturday file:', error);
        showNotification('Gagal mengunduh file Sabtu', 'error');
    }
}

// Fungsi untuk membuat data export Sabtu
function createSaturdayExportData(employeeData, employeeName) {
    const category = employeeCategories[employeeName] || 'STAFF';
    const rate = overtimeRates[category];
    
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
    
    return exportData;
}

// Fungsi untuk membuat summary data Sabtu
function createSaturdaySummaryData(saturdayData) {
    const employees = [...new Set(saturdayData.map(item => item.nama))];
    const summary = [];
    
    employees.forEach(employee => {
        const employeeData = saturdayData.filter(item => item.nama === employee);
        const totalLemburBulat = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
        const totalLemburDesimal = employeeData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
        const category = employeeCategories[employee] || 'STAFF';
        const rate = overtimeRates[category];
        const totalGaji = totalLemburBulat * rate;
        
        summary.push({
            'No': summary.length + 1,
            'Nama Karyawan': employee,
            'Kategori': category,
            'Rate per Jam': formatCurrency(rate),
            'Total Hari Sabtu': employeeData.length,
            'Total Jam Lembur (Bulat)': totalLemburBulat,
            'Total Jam Lembur (Desimal)': totalLemburDesimal.toFixed(2),
            'Total Gaji Lembur': formatCurrency(totalGaji),
            'Rata-rata per Hari': formatCurrency(totalGaji / employeeData.length)
        });
    });
    
    // Tambahkan total row
    const totalLemburBulatAll = saturdayData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
    const totalLemburDesimalAll = saturdayData.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    let totalGajiAll = 0;
    employees.forEach(employee => {
        const employeeData = saturdayData.filter(item => item.nama === employee);
        const totalLemburBulatEmp = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
        const category = employeeCategories[employee] || 'STAFF';
        const rate = overtimeRates[category];
        totalGajiAll += totalLemburBulatEmp * rate;
    });
    
    summary.push({
        'No': '',
        'Nama Karyawan': 'TOTAL SABTU',
        'Kategori': '',
        'Rate per Jam': '',
        'Total Hari Sabtu': saturdayData.length,
        'Total Jam Lembur (Bulat)': totalLemburBulatAll,
        'Total Jam Lembur (Desimal)': totalLemburDesimalAll.toFixed(2),
        'Total Gaji Lembur': formatCurrency(totalGajiAll),
        'Rata-rata per Hari': formatCurrency(totalGajiAll / saturdayData.length)
    });
    
    return summary;
}

// ============================
// FUNGSI LAINNYA (Upload, Process, dll)
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
    
    // Format info lembur dengan pembulatan
    let lemburInfo = `${totalLemburBulat} jam`;
    if (totalLemburDesimal > totalLemburBulat) {
        lemburInfo = `${totalLemburBulat} jam (dibulatkan dari ${totalLemburDesimal.toFixed(2)} jam)`;
    }
    if (totalLemburElem) totalLemburElem.textContent = lemburInfo;
    
    // Hitung total gaji
    let totalGaji = 0;
    const employees = [...new Set(data.map(item => item.nama))];
    employees.forEach(employee => {
        const employeeData = data.filter(item => item.nama === employee);
        const totalLemburBulatEmp = employeeData.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
        const category = employeeCategories[employee] || 'STAFF';
        const rate = overtimeRates[category];
        totalGaji += totalLemburBulatEmp * rate;
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
