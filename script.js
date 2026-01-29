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
// FUNGSI UNTUK HARI JUMAT
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
        
        const maxNormalMinutes = isSaturday(dateString) ? 6 * 60 : 8 * 60;
        
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
// FUNGSI UPLOAD FILE
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
        
        // 1. Event untuk klik di upload area
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
                    console.log('File dropped:', file.name);
                    
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
        
        // 2. Event untuk tombol browse
        if (browseBtn) {
            browseBtn.addEventListener('click', function(e) {
                console.log('Browse button clicked');
                e.stopPropagation();
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
        
        // ============================
        // FIXED: EVENT LISTENERS UNTUK TOMBOL DOWNLOAD
        // ============================
        
        console.log('Setting up download buttons...');
        
        // Debug: Tampilkan status tombol
        const downloadButtons = [
            'download-original',
            'download-processed', 
            'download-both',
            'download-salary',
            'download-saturday'
        ];
        
        downloadButtons.forEach(id => {
            const btn = document.getElementById(id);
            console.log(`Button ${id}:`, btn ? 'FOUND' : 'NOT FOUND');
        });
        
        // Setup event listeners untuk tombol download
        const downloadOriginal = document.getElementById('download-original');
        const downloadProcessed = document.getElementById('download-processed');
        const downloadBoth = document.getElementById('download-both');
        const downloadSalary = document.getElementById('download-salary');
        const downloadSaturday = document.getElementById('download-saturday');
        
        // Tambahkan debug pada setiap klik
        if (downloadOriginal) {
            console.log('Setting up download-original button');
            downloadOriginal.addEventListener('click', function(e) {
                console.log('Download Original clicked!');
                console.log('Original data length:', originalData.length);
                downloadReport('original');
            });
        }
        
        if (downloadProcessed) {
            console.log('Setting up download-processed button');
            downloadProcessed.addEventListener('click', function(e) {
                console.log('Download Processed clicked!');
                console.log('Processed data length:', processedData.length);
                downloadReport('processed');
            });
        }
        
        if (downloadBoth) {
            console.log('Setting up download-both button');
            downloadBoth.addEventListener('click', function(e) {
                console.log('Download Both clicked!');
                downloadReport('both');
            });
        }
        
        if (downloadSalary) {
            console.log('Setting up download-salary button');
            downloadSalary.addEventListener('click', function(e) {
                console.log('Download Salary clicked!');
                showSeparatedDownloadOptions();
            });
        }
        
        if (downloadSaturday) {
            console.log('Setting up download-saturday button');
            downloadSaturday.addEventListener('click', function(e) {
                console.log('Download Saturday clicked!');
                if (processedData.length === 0) {
                    showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
                    return;
                }
                showSaturdayDownloadModal();
            });
        }
        
        // Setup global click handler untuk testing
        document.addEventListener('click', function(e) {
            if (e.target.id && e.target.id.includes('download')) {
                console.log(`Global click detected on: ${e.target.id}`);
            }
        });
        
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
        
        // Jam normal: maksimal 8 jam, atau total jam jika kurang dari 8 jam
        const jamNormal = Math.min(hoursWorked, workHours);
        
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
        
        // Cek apakah hari Jumat
        const isFridayDay = isFriday(record.tanggal);
        const isSaturdayDay = isSaturday(record.tanggal);
        
        // Keterangan khusus dengan info hari
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
// FUNGSI PROCESS EXCEL
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

// ============================
// FUNGSI UNTUK DOWNLOAD
// ============================

// Fungsi utama untuk download laporan
function downloadReport(type) {
    console.log(`Download report type: ${type}`);
    
    // Validasi berdasarkan tipe
    if (type === 'original' && originalData.length === 0) {
        showNotification('Tidak ada data asli untuk diunduh', 'warning');
        return;
    }
    
    if ((type === 'processed' || type === 'both') && processedData.length === 0) {
        showNotification('Data belum diproses. Silakan klik "Hitung Lembur" terlebih dahulu.', 'warning');
        return;
    }
    
    try {
        if (type === 'original') {
            console.log('Downloading original data...');
            downloadOriginalData();
        } else if (type === 'processed') {
            console.log('Downloading processed data...');
            downloadProcessedData();
        } else if (type === 'both') {
            console.log('Downloading both files...');
            downloadBothFiles();
        }
    } catch (error) {
        console.error('Error in downloadReport:', error);
        showNotification('Gagal mengunduh laporan: ' + error.message, 'error');
    }
}

// Download data asli
function downloadOriginalData() {
    const exportData = originalData.map((item, index) => ({
        'No': index + 1,
        'Nama': item.nama,
        'Tanggal': formatDate(item.tanggal),
        'Hari': getDayName(item.tanggal),
        'Jam Masuk': item.jamMasuk,
        'Jam Keluar': item.jamKeluar,
        'Durasi': item.durasi ? item.durasi.toFixed(2) + ' jam' : ''
    }));
    
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Asli');
    
    // Set kolom width
    const wscols = [
        {wch: 5},   // No
        {wch: 20},  // Nama
        {wch: 12},  // Tanggal
        {wch: 10},  // Hari
        {wch: 10},  // Jam Masuk
        {wch: 10},  // Jam Keluar
        {wch: 12}   // Durasi
    ];
    worksheet['!cols'] = wscols;
    
    XLSX.writeFile(workbook, `data_presensi_asli_${new Date().toISOString().split('T')[0]}.xlsx`);
    showNotification('Data asli berhasil diunduh!', 'success');
}

// Download data terproses
function downloadProcessedData() {
    const exportData = processedData.map((item, index) => {
        const jamLemburBulat = roundOvertimeHours(item.jamLemburDesimal);
        const category = employeeCategories[item.nama] || 'STAFF';
        const rate = overtimeRates[category] || 10000;
        const gajiLembur = jamLemburBulat * rate;
        
        return {
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Hari': item.hari,
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi Kerja': item.durasiFormatted,
            'Jam Normal': item.jamNormalFormatted,
            'Jam Lembur (Desimal)': item.jamLemburDesimal.toFixed(2),
            'Jam Lembur (Bulat)': jamLemburBulat,
            'Kategori': category,
            'Rate per Jam': formatCurrency(rate),
            'Total Gaji Lembur': formatCurrency(gajiLembur),
            'Keterangan': item.keterangan
        };
    });
    
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Data Lembur');
    
    // Set kolom width
    const wscols = [
        {wch: 5},   // No
        {wch: 25},  // Nama
        {wch: 12},  // Tanggal
        {wch: 10},  // Hari
        {wch: 10},  // Jam Masuk
        {wch: 10},  // Jam Keluar
        {wch: 12},  // Durasi
        {wch: 12},  // Jam Normal
        {wch: 15},  // Jam Lembur (Desimal)
        {wch: 15},  // Jam Lembur (Bulat)
        {wch: 10},  // Kategori
        {wch: 15},  // Rate
        {wch: 20},  // Total Gaji
        {wch: 30}   // Keterangan
    ];
    worksheet['!cols'] = wscols;
    
    XLSX.writeFile(workbook, `data_lembur_harian_${new Date().toISOString().split('T')[0]}.xlsx`);
    showNotification('Data lembur berhasil diunduh!', 'success');
}

// Download kedua file
function downloadBothFiles() {
    // Buat workbook dengan dua sheet
    const workbook = XLSX.utils.book_new();
    
    // Sheet 1: Data Asli
    const exportDataOriginal = originalData.map((item, index) => ({
        'No': index + 1,
        'Nama': item.nama,
        'Tanggal': formatDate(item.tanggal),
        'Hari': getDayName(item.tanggal),
        'Jam Masuk': item.jamMasuk,
        'Jam Keluar': item.jamKeluar,
        'Durasi': item.durasi ? item.durasi.toFixed(2) + ' jam' : ''
    }));
    
    const worksheet1 = XLSX.utils.json_to_sheet(exportDataOriginal);
    XLSX.utils.book_append_sheet(workbook, worksheet1, 'Data Asli');
    
    // Sheet 2: Data Lembur
    const exportDataProcessed = processedData.map((item, index) => {
        const jamLemburBulat = roundOvertimeHours(item.jamLemburDesimal);
        const category = employeeCategories[item.nama] || 'STAFF';
        const rate = overtimeRates[category] || 10000;
        const gajiLembur = jamLemburBulat * rate;
        
        return {
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Hari': item.hari,
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi Kerja': item.durasiFormatted,
            'Jam Normal': item.jamNormalFormatted,
            'Jam Lembur (Bulat)': jamLemburBulat,
            'Total Gaji Lembur': formatCurrency(gajiLembur),
            'Keterangan': item.keterangan
        };
    });
    
    const worksheet2 = XLSX.utils.json_to_sheet(exportDataProcessed);
    XLSX.utils.book_append_sheet(workbook, worksheet2, 'Data Lembur');
    
    XLSX.writeFile(workbook, `data_lengkap_${new Date().toISOString().split('T')[0]}.xlsx`);
    showNotification('Kedua file berhasil diunduh dalam satu workbook!', 'success');
}

// Fungsi untuk modal pilihan download terpisah
function showSeparatedDownloadOptions() {
    if (processedData.length === 0) {
        showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
        return;
    }
    
    // Pisahkan data Jumat dan hari lain
    const fridayData = processedData.filter(item => item.isFriday);
    const otherDaysData = processedData.filter(item => !item.isFriday && !item.isSaturday);
    const saturdayData = processedData.filter(item => item.isSaturday);
    
    // Buat modal
    const modalHtml = `
        <div class="modal" id="separated-download-modal">
            <div class="modal-content">
                <div class="modal-header">
                    <h3><i class="fas fa-download"></i> Pilih Format Download</h3>
                    <button class="modal-close" id="close-separated-download">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="download-stats" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 1rem; border-radius: 8px; margin-bottom: 1.5rem;">
                        <h4><i class="fas fa-chart-bar"></i> Statistik Data</h4>
                        <div style="display: flex; justify-content: space-around; text-align: center;">
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Total Data</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${processedData.length}</p>
                            </div>
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Hari Jumat</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${fridayData.length}</p>
                            </div>
                            <div>
                                <p style="font-size: 0.9rem; margin-bottom: 0.25rem;">Hari Lain</p>
                                <p style="font-size: 1.5rem; font-weight: bold;">${otherDaysData.length}</p>
                            </div>
                        </div>
                    </div>
                    
                    <div class="download-options-separated">
                        <div class="option-card" id="option-all-together">
                            <div class="option-icon">
                                <i class="fas fa-file-excel" style="color: #217346; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Semua Data dalam Satu File</h4>
                                <p>Download semua data lembur dalam satu file Excel</p>
                                <ul>
                                    <li>Semua hari (Jumat, Sabtu & Senin-Kamis)</li>
                                    <li>Format tabel lengkap</li>
                                    <li>Termasuk perhitungan gaji</li>
                                </ul>
                            </div>
                        </div>
                        
                        <div class="option-card" id="option-friday-only">
                            <div class="option-icon">
                                <i class="fas fa-calendar-day" style="color: #9b59b6; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Hanya Data Hari Jumat</h4>
                                <p>Download data lembur khusus hari Jumat</p>
                                <ul>
                                    <li>${fridayData.length} entri data Jumat</li>
                                    <li>Aturan khusus: pulang jam 15:00</li>
                                    <li>Lembur setelah 8 jam kerja</li>
                                </ul>
                            </div>
                        </div>
                        
                        <div class="option-card" id="option-other-days-only">
                            <div class="option-icon">
                                <i class="fas fa-calendar-alt" style="color: #3498db; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Hanya Data Senin-Kamis</h4>
                                <p>Download data lembur hari Senin sampai Kamis</p>
                                <ul>
                                    <li>${otherDaysData.length} entri data</li>
                                    <li>Aturan normal: pulang jam 16:00</li>
                                    <li>Lembur setelah 8 jam kerja</li>
                                </ul>
                            </div>
                        </div>
                        
                        ${saturdayData.length > 0 ? `
                        <div class="option-card" id="option-saturday-only">
                            <div class="option-icon">
                                <i class="fas fa-calendar-day" style="color: #e67e22; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Hanya Data Sabtu</h4>
                                <p>Download data lembur hari Sabtu</p>
                                <ul>
                                    <li>${saturdayData.length} entri data Sabtu</li>
                                    <li>Aturan khusus: kerja 6 jam normal</li>
                                    <li>K3 bekerja 07:00-22:00</li>
                                </ul>
                            </div>
                        </div>
                        ` : ''}
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="cancel-separated-download">Batal</button>
                </div>
            </div>
        </div>
    `;
    
    // Tambahkan modal ke body jika belum ada
    let modal = document.getElementById('separated-download-modal');
    if (!modal) {
        document.body.insertAdjacentHTML('beforeend', modalHtml);
        modal = document.getElementById('separated-download-modal');
        
        // Setup event listeners untuk modal
        document.getElementById('close-separated-download').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        document.getElementById('cancel-separated-download').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        // Event untuk setiap opsi
        document.getElementById('option-all-together').addEventListener('click', () => {
            downloadProcessedData();
            modal.classList.remove('active');
        });
        
        document.getElementById('option-friday-only').addEventListener('click', () => {
            downloadFilteredData(fridayData, 'jumat');
            modal.classList.remove('active');
        });
        
        document.getElementById('option-other-days-only').addEventListener('click', () => {
            downloadFilteredData(otherDaysData, 'senin_kamis');
            modal.classList.remove('active');
        });
        
        if (saturdayData.length > 0) {
            document.getElementById('option-saturday-only').addEventListener('click', () => {
                downloadFilteredData(saturdayData, 'sabtu');
                modal.classList.remove('active');
            });
        }
        
        // Close modal ketika klik di luar
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });
    }
    
    modal.classList.add('active');
}

// Download data yang sudah difilter
function downloadFilteredData(data, typeName) {
    if (data.length === 0) {
        showNotification(`Tidak ada data ${typeName} untuk diunduh`, 'warning');
        return;
    }
    
    const exportData = data.map((item, index) => {
        const jamLemburBulat = roundOvertimeHours(item.jamLemburDesimal);
        const category = employeeCategories[item.nama] || 'STAFF';
        const rate = overtimeRates[category] || 10000;
        const gajiLembur = jamLemburBulat * rate;
        
        return {
            'No': index + 1,
            'Nama Karyawan': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Hari': item.hari,
            'Jam Masuk': item.jamMasuk,
            'Jam Keluar': item.jamKeluar,
            'Durasi Kerja': item.durasiFormatted,
            'Jam Normal': item.jamNormalFormatted,
            'Jam Lembur (Bulat)': jamLemburBulat,
            'Total Gaji Lembur': formatCurrency(gajiLembur),
            'Keterangan': item.keterangan
        };
    });
    
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    
    let sheetName = 'Data Lembur';
    let fileName = `data_lembur`;
    
    switch(typeName) {
        case 'jumat':
            sheetName = 'Data Lembur Jumat';
            fileName = `data_lembur_jumat`;
            break;
        case 'senin_kamis':
            sheetName = 'Data Lembur Senin-Kamis';
            fileName = `data_lembur_senin_kamis`;
            break;
        case 'sabtu':
            sheetName = 'Data Lembur Sabtu';
            fileName = `data_lembur_sabtu`;
            break;
    }
    
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, `${fileName}_${new Date().toISOString().split('T')[0]}.xlsx`);
    
    showNotification(`Data ${typeName} berhasil diunduh!`, 'success');
}

// ============================
// FUNGSI UNTUK MODAL SABTU
// ============================

// Fungsi untuk menghitung lembur hari Sabtu
function calculateSaturdayOvertime(data) {
    const saturdayData = data.filter(item => item.isSaturday);
    
    if (saturdayData.length === 0) {
        return {
            data: [],
            summary: [],
            totalEmployees: 0,
            totalOvertime: 0,
            totalOvertimeBulat: 0,
            totalGaji: 0
        };
    }
    
    // Proses perhitungan lembur Sabtu
    const processedData = saturdayData.map(record => {
        const category = employeeCategories[record.nama] || 'STAFF';
        const rate = overtimeRates[category];
        
        // Untuk hari Sabtu, jam kerja normal hanya 6 jam
        const normalWorkHours = 6;
        
        // Hitung durasi kerja
        const hoursWorked = calculateHoursWithFridayRules(
            record.jamMasuk, 
            record.jamKeluar, 
            record.tanggal
        );
        
        // Aturan khusus untuk K3 di hari Sabtu
        let overtimeHours = 0;
        if (category === 'K3') {
            // K3 bekerja dari 07:00 sampai 22:00 di hari Sabtu
            const [startHour, startMinute] = SATURDAY_RULES.k3SpecialHours.startTime.split(':').map(Number);
            const [endHour, endMinute] = SATURDAY_RULES.k3SpecialHours.endTime.split(':').map(Number);
            
            const workStartMinutes = startHour * 60 + startMinute;
            const workEndMinutes = endHour * 60 + endMinute;
            
            // Parse waktu keluar
            const [outHour, outMinute] = record.jamKeluar.split(':').map(Number);
            const outMinutes = outHour * 60 + (outMinute || 0);
            
            // Hitung lembur jika pulang setelah 22:00
            if (outMinutes > workEndMinutes) {
                overtimeHours = (outMinutes - workEndMinutes) / 60;
            }
        } else {
            // Untuk non-K3, lembur dihitung setelah 6 jam kerja
            if (hoursWorked > normalWorkHours) {
                overtimeHours = hoursWorked - normalWorkHours;
            }
        }
        
        // Abaikan lembur kurang dari 10 menit
        if (overtimeHours < (SATURDAY_RULES.minOvertimeMinutes / 60)) {
            overtimeHours = 0;
        }
        
        // BATASI MAKSIMAL 7 JAM LEMBUR
        if (overtimeHours > 7) {
            overtimeHours = 7;
        }
        
        // Hitung jam normal (maksimal 6 jam untuk Sabtu)
        const normalHours = Math.min(hoursWorked, normalWorkHours);
        
        // Pembulatan jam lembur
        const overtimeBulat = roundOvertimeHours(overtimeHours);
        
        return {
            ...record,
            kategori: category,
            rate: rate,
            jamNormalSabtu: normalHours,
            jamLemburSabtu: overtimeHours,
            jamLemburSabtuBulat: overtimeBulat,
            gajiLemburSabtu: overtimeBulat * rate,
            keteranganSabtu: category === 'K3' 
                ? `K3 - Kerja Sabtu 07:00-22:00 (${overtimeBulat} jam lembur)` 
                : `Sabtu - Kerja 6 jam normal (${overtimeBulat} jam lembur)`
        };
    });
    
    // Group by employee untuk summary
    const employeeGroups = {};
    processedData.forEach(item => {
        const employeeName = item.nama;
        if (!employeeGroups[employeeName]) {
            employeeGroups[employeeName] = {
                nama: employeeName,
                kategori: item.kategori,
                rate: item.rate,
                totalJam: 0,
                totalJamBulat: 0,
                totalGaji: 0,
                records: []
            };
        }
        
        employeeGroups[employeeName].totalJam += item.jamLemburSabtu;
        employeeGroups[employeeName].totalJamBulat += item.jamLemburSabtuBulat;
        employeeGroups[employeeName].totalGaji += item.gajiLemburSabtu;
        employeeGroups[employeeName].records.push(item);
    });
    
    const summary = Object.values(employeeGroups);
    
    // Hitung total
    const totalOvertime = processedData.reduce((sum, item) => sum + item.jamLemburSabtu, 0);
    const totalOvertimeBulat = processedData.reduce((sum, item) => sum + item.jamLemburSabtuBulat, 0);
    const totalGaji = processedData.reduce((sum, item) => sum + item.gajiLemburSabtu, 0);
    
    return {
        data: processedData,
        summary: summary,
        totalEmployees: summary.length,
        totalOvertime: totalOvertime,
        totalOvertimeBulat: totalOvertimeBulat,
        totalGaji: totalGaji
    };
}

// Fungsi untuk modal download Sabtu
function showSaturdayDownloadModal() {
    const saturdayData = calculateSaturdayOvertime(processedData);
    
    const modalHtml = `
        <div class="modal" id="saturday-download-modal">
            <div class="modal-content">
                <div class="modal-header">
                    <h3><i class="fas fa-calendar-day"></i> Download Lembur Hari Sabtu</h3>
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
                                <p style="font-size: 1.5rem; font-weight: bold; color: #e67e22;">${saturdayData.totalEmployees}</p>
                            </div>
                            <div>
                                <h5 style="color: #666; font-size: 0.9rem;">Total Data</h5>
                                <p style="font-size: 1.5rem; font-weight: bold; color: #e67e22;">${saturdayData.data.length}</p>
                            </div>
                            <div>
                                <h5 style="color: #666; font-size: 0.9rem;">Jam Lembur</h5>
                                <p style="font-size: 1.5rem; font-weight: bold; color: #e67e22;">${saturdayData.totalOvertimeBulat.toFixed(1)} jam</p>
                            </div>
                            <div>
                                <h5 style="color: #666; font-size: 0.9rem;">Total Gaji</h5>
                                <p style="font-size: 1.5rem; font-weight: bold; color: #27ae60;">${formatCurrency(saturdayData.totalGaji)}</p>
                            </div>
                        </div>
                    </div>
                    
                    <div class="download-options-saturday">
                        <div class="option-card" id="option-saturday-all">
                            <div class="option-icon">
                                <i class="fas fa-file-excel" style="color: #e67e22; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Data Lembur Sabtu Lengkap</h4>
                                <p>Download semua data lembur hari Sabtu</p>
                                <ul>
                                    <li>${saturdayData.data.length} entri data</li>
                                    <li>${saturdayData.totalEmployees} karyawan</li>
                                    <li>Format tabel detail per orang</li>
                                    <li>Termasuk perhitungan gaji</li>
                                    <li>Aturan khusus kategori K3</li>
                                </ul>
                            </div>
                        </div>
                        
                        ${saturdayData.summary.length > 0 ? `
                        <div class="option-card" id="option-saturday-summary">
                            <div class="option-icon">
                                <i class="fas fa-users" style="color: #3498db; font-size: 2rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Rekap per Karyawan</h4>
                                <p>Download ringkasan lembur per karyawan</p>
                                <ul>
                                    <li>${saturdayData.summary.length} karyawan</li>
                                    <li>Total jam lembur per orang</li>
                                    <li>Total gaji lembur per orang</li>
                                    <li>Format ringkas untuk laporan keuangan</li>
                                </ul>
                            </div>
                        </div>
                        ` : ''}
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="cancel-saturday-download">Batal</button>
                </div>
            </div>
        </div>
    `;
    
    // Tambahkan modal ke body jika belum ada
    let modal = document.getElementById('saturday-download-modal');
    if (!modal) {
        document.body.insertAdjacentHTML('beforeend', modalHtml);
        modal = document.getElementById('saturday-download-modal');
        
        // Setup event listeners untuk modal
        document.getElementById('close-saturday-download').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        document.getElementById('cancel-saturday-download').addEventListener('click', () => {
            modal.classList.remove('active');
        });
        
        // Event untuk opsi download
        document.getElementById('option-saturday-all').addEventListener('click', () => {
            downloadSaturdayReport(saturdayData, 'detail');
            modal.classList.remove('active');
        });
        
        if (saturdayData.summary.length > 0) {
            document.getElementById('option-saturday-summary').addEventListener('click', () => {
                downloadSaturdayReport(saturdayData, 'summary');
                modal.classList.remove('active');
            });
        }
        
        // Close modal ketika klik di luar
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                modal.classList.remove('active');
            }
        });
    }
    
    modal.classList.add('active');
}

// Fungsi untuk download laporan Sabtu
function downloadSaturdayReport(saturdayData, type = 'detail') {
    if (!saturdayData || saturdayData.data.length === 0) {
        showNotification('Tidak ada data lembur Sabtu', 'warning');
        return;
    }
    
    try {
        if (type === 'detail') {
            // Download detail lengkap
            const exportData = saturdayData.data.map((item, index) => ({
                'No': index + 1,
                'Nama Karyawan': item.nama,
                'Tanggal': formatDate(item.tanggal),
                'Hari': 'Sabtu',
                'Jam Masuk': item.jamMasuk,
                'Jam Keluar': item.jamKeluar,
                'Kategori': item.kategori,
                'Jam Kerja': item.durasi.toFixed(2) + ' jam',
                'Jam Normal (6 jam)': item.jamNormalSabtu.toFixed(2) + ' jam',
                'Jam Lembur': item.jamLemburSabtuBulat + ' jam',
                'Rate per Jam': formatCurrency(item.rate),
                'Total Gaji Lembur': formatCurrency(item.gajiLemburSabtu),
                'Keterangan': item.keteranganSabtu
            }));
            
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Lembur Sabtu Detail');
            XLSX.writeFile(workbook, `lembur_sabtu_detail_${new Date().toISOString().split('T')[0]}.xlsx`);
            
            showNotification('Laporan detail lembur Sabtu berhasil diunduh', 'success');
            
        } else if (type === 'summary') {
            // Download summary per karyawan
            const exportData = saturdayData.summary.map((item, index) => ({
                'No': index + 1,
                'Nama Karyawan': item.nama,
                'Kategori': item.kategori,
                'Rate per Jam': formatCurrency(item.rate),
                'Total Jam Lembur (Desimal)': item.totalJam.toFixed(2),
                'Total Jam Lembur (Bulat)': item.totalJamBulat,
                'Total Gaji Lembur': formatCurrency(item.totalGaji)
            }));
            
            const worksheet = XLSX.utils.json_to_sheet(exportData);
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, 'Rekap Lembur Sabtu');
            XLSX.writeFile(workbook, `rekap_lembur_sabtu_${new Date().toISOString().split('T')[0]}.xlsx`);
            
            showNotification('Rekap lembur Sabtu berhasil diunduh', 'success');
        }
        
    } catch (error) {
        console.error('Error downloading Saturday report:', error);
        showNotification('Gagal mengunduh laporan Sabtu', 'error');
    }
}

// ============================
// FUNGSI UPLOAD HANDLER
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
    console.log('Setting up download buttons...');
    
    // Cek semua tombol download
    const downloadButtons = [
        'download-original',
        'download-processed', 
        'download-both',
        'download-salary',
        'download-saturday'
    ];
    
    downloadButtons.forEach(id => {
        const btn = document.getElementById(id);
        console.log(`Button ${id}:`, btn ? '✓ DITEMUKAN' : '✗ TIDAK DITEMUKAN');
    });
    
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
                    
                    // Test tombol download setelah load
                    console.log('Testing download buttons after load...');
                    downloadButtons.forEach(id => {
                        const btn = document.getElementById(id);
                        if (btn) {
                            console.log(`Button ${id} is ready`);
                        }
                    });
                    
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
            
            // Enable download buttons setelah data diproses
            console.log('Processed data ready for download:', processedData.length, 'records');
            
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
    data.forEach(item => {
        const jamLemburBulat = roundOvertimeHours(item.jamLemburDesimal);
        const category = employeeCategories[item.nama] || 'STAFF';
        const rate = overtimeRates[category] || 10000;
        totalGaji += jamLemburBulat * rate;
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
    // Hitung summary per karyawan
    const employeeSummary = {};
    
    data.forEach(item => {
        const employeeName = item.nama;
        if (!employeeSummary[employeeName]) {
            employeeSummary[employeeName] = {
                nama: employeeName,
                totalLembur: 0,
                totalLemburBulat: 0,
                totalGaji: 0,
                kategori: employeeCategories[employeeName] || 'STAFF'
            };
        }
        
        const jamLemburBulat = roundOvertimeHours(item.jamLemburDesimal);
        const rate = overtimeRates[employeeSummary[employeeName].kategori] || 10000;
        
        employeeSummary[employeeName].totalLembur += item.jamLemburDesimal;
        employeeSummary[employeeName].totalLemburBulat += jamLemburBulat;
        employeeSummary[employeeName].totalGaji += jamLemburBulat * rate;
    });
    
    const employeeSummaryArray = Object.values(employeeSummary);
    const employeeSummaryElem = document.getElementById('employee-summary');
    const financialSummary = document.getElementById('financial-summary');
    
    if (employeeSummaryElem) {
        let html = '<div style="max-height: 300px; overflow-y: auto;">';
        
        // Urutkan berdasarkan total lembur tertinggi
        employeeSummaryArray.sort((a, b) => b.totalLemburBulat - a.totalLemburBulat);
        
        employeeSummaryArray.forEach(emp => {
            if (emp.totalLemburBulat > 0) {
                const isCapped = emp.totalLembur > emp.totalLemburBulat * 0.9 && emp.totalLembur > 7;
                
                html += `
                    <div style="margin-bottom: 1rem; padding: 0.75rem; background: #f8f9fa; border-radius: 8px;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div>
                                <strong>${emp.nama}</strong>
                                <br><small style="color: #666;">Kategori: ${emp.kategori}</small>
                            </div>
                            <div style="text-align: right;">
                                <div style="color: ${isCapped ? '#e67e22' : '#e74c3c'}; font-weight: bold;">
                                    ${emp.totalLemburBulat} jam lembur
                                    ${isCapped ? '<small style="font-size: 0.8rem;">(dibatasi)</small>' : ''}
                                </div>
                                <div style="font-size: 0.9rem;">(${emp.totalLembur.toFixed(2)} jam desimal)</div>
                            </div>
                        </div>
                        <div style="text-align: right; font-weight: bold; color: #27ae60; margin-top: 0.5rem;">
                            ${formatCurrency(emp.totalGaji)}
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
    employeeSummaryArray.forEach(emp => {
        totalGajiAll += emp.totalGaji;
    });
    
    // Hitung breakdown per hari
    const fridayData = data.filter(item => item.isFriday);
    const saturdayData = data.filter(item => item.isSaturday);
    const otherDaysData = data.filter(item => !item.isFriday && !item.isSaturday);
    
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
                    Total Jam Lembur: <strong>${totalLemburBulat} jam (${totalLemburDesimal.toFixed(2)} jam desimal)</strong>
                </div>
            </div>
            
            <div style="margin-top: 1rem; padding: 0.75rem; background: #e8f4f8; border-radius: 4px;">
                <strong>Breakdown per Hari:</strong>
                <div style="display: flex; justify-content: space-between; margin-top: 0.5rem;">
                    <div>Jumat: ${fridayData.length} hari</div>
                    <div>Sabtu: ${saturdayData.length} hari</div>
                    <div>Senin-Kamis: ${otherDaysData.length} hari</div>
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
        employeeGroups[item.nama].lembur += roundOvertimeHours(item.jamLemburDesimal);
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
        const totalNormal = data.reduce((sum, item) => sum + item.jamNormal, 0);
        const totalLembur = data.reduce((sum, item) => sum + roundOvertimeHours(item.jamLemburDesimal), 0);
        
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
