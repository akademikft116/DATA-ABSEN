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
    workStartTime: '07:00',      // Jam masuk minimal untuk dihitung
    workEndTime: '16:00',        // Jam pulang normal
    minOvertimeMinutes: 10,      // Minimal lembur yang dihitung
    dailyWorkHours: 8            // Jam kerja normal sehari
};

// Hari Jumat special rules (tetap sama)
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
        // Lembur maksimal 7 jam per hari
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
// FUNGSI UNTUK HARI JUMAT
// ============================

// Fungsi untuk mengecek apakah suatu tanggal adalah hari Jumat
function isFriday(dateString) {
    if (!dateString) return false;
    
    try {
        // Parse tanggal dari format DD/MM/YYYY
        const [day, month, year] = dateString.split('/');
        const date = new Date(year, month - 1, day);
        
        // Cek apakah hari Jumat (5 = Jumat)
        return date.getDay() === 5;
    } catch (error) {
        console.error('Error checking Friday:', error);
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
    return isFriday(dateString) ? FRIDAY_RULES : NORMAL_DAY_RULES;
}

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
        
        // PERUBAHAN: Gunakan jam masuk sebenarnya (tidak dipaksakan 07:00)
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        const outMinutes = outTime.hours * 60 + outTime.minutes;
        
        // Hitung total menit kerja
        let totalMinutes = outMinutes - inMinutes;
        
        // Jika pulang sebelum masuk (melewati tengah malam)
        if (totalMinutes < 0) {
            totalMinutes += 24 * 60;
        }
        
        // Konversi ke jam
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
        
        // Dapatkan aturan berdasarkan hari
        const rules = getDayRules(dateString);
        
        // PERUBAHAN: Gunakan jam masuk sebenarnya (FLEKSIBEL)
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        const outMinutes = outTime.hours * 60 + outTime.minutes;
        
        // Hitung total menit kerja
        let totalMinutes = outMinutes - inMinutes;
        if (totalMinutes < 0) totalMinutes += 24 * 60;
        
        // Lembur dihitung jika kerja > 8 jam dari jam masuk
        const maxNormalMinutes = rules.dailyWorkHours * 60; // 8 jam = 480 menit
        
        // Jika total kerja kurang dari 8 jam, tidak ada lembur
        if (totalMinutes <= maxNormalMinutes) return 0;
        
        // Hitung menit lembur (kelebihan dari 8 jam)
        let overtimeMinutes = totalMinutes - maxNormalMinutes;
        
        // Abaikan lembur kurang dari 10 menit
        if (overtimeMinutes < rules.minOvertimeMinutes) {
            return 0;
        }
        
        // Konversi ke jam (desimal)
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

// Fungsi untuk mendapatkan jam masuk efektif dengan aturan fleksibel
function getEffectiveInTimeWithDayRules(jamMasuk, dateString) {
    // PERUBAHAN: Kembalikan jam masuk sebenarnya (FLEKSIBEL)
    return jamMasuk;
}

// Calculate hours between two time strings (untuk kompatibilitas)
function calculateHours(timeIn, timeOut) {
    return calculateHoursWithFridayRules(timeIn, timeOut, '01/01/2024');
}

// ============================
// FUNGSI BARU: PISAH DATA JUMAT DAN HARI LAIN
// ============================

// Fungsi untuk memisahkan data berdasarkan hari
function separateDataByDay(data) {
    const fridayData = data.filter(item => item.isFriday);
    const otherDaysData = data.filter(item => !item.isFriday);
    
    return {
        friday: fridayData,
        otherDays: otherDaysData
    };
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
        
        // Tentukan jam masuk efektif (jam masuk sebenarnya)
        const effectiveInTime = getEffectiveInTimeWithDayRules(record.jamMasuk, record.tanggal);
        
        // Cek apakah hari Jumat
        const isFridayDay = isFriday(record.tanggal);
        
        // Keterangan khusus dengan info hari
        let keterangan = 'Tidak lembur';
        let dayInfo = isFridayDay ? ' (JUMAT)' : '';
        
        if (jamLemburDesimal > 0) {
            if (isFridayDay) {
                keterangan = `Lembur ${jamLemburDisplay} (Jumat - kerja > 8 jam dari jam masuk)`;
            } else {
                keterangan = `Lembur ${jamLemburDisplay} (kerja > 8 jam dari jam masuk)`;
            }
            
            // Tambahkan info jika lembur dibatasi 7 jam
            if (jamLemburDesimal > 7) {
                keterangan = `Lembur ${jamLemburDisplay} (DIBATASI MAKSIMAL 7 JAM)`;
            }
            
            if (effectiveInTime !== record.jamMasuk) {
                keterangan += ` (masuk efektif: ${effectiveInTime})`;
            }
        } else if (effectiveInTime !== record.jamMasuk) {
            keterangan = `Masuk efektif: ${effectiveInTime}${dayInfo}`;
        } else if (isFridayDay) {
            keterangan = `Hari Jumat - Kerja 8 jam normal`;
        } else {
            keterangan = `Kerja ≤ 8 jam dari jam masuk`;
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
            jamKerjaNormal: workHours,
            isFriday: isFridayDay,
            workEndTime: isFridayDay ? '15:00' : '16:00',
            hari: getDayName(record.tanggal),
            // Tambahan untuk debugging
            workStartEffective: record.jamMasuk,
            maxNormalHours: 8
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
// KONFIGURASI ATURAN HARI SABTU
// ============================

const SATURDAY_RULES = {
    workHours: 6,                // Hanya bekerja 6 jam di hari Sabtu
    minOvertimeMinutes: 10,      // Minimal lembur yang dihitung
    k3SpecialHours: {            // Khusus untuk kategori K3
        startTime: '07:00',      // Masuk jam 7 pagi
        endTime: '22:00',        // Pulang jam 10 malam
        isSpecialDay: true       // Hari kerja khusus
    }
};

// Fungsi untuk mengecek apakah suatu tanggal adalah hari Sabtu
function isSaturday(dateString) {
    if (!dateString) return false;
    
    try {
        const [day, month, year] = dateString.split('/');
        const date = new Date(year, month - 1, day);
        return date.getDay() === 6; // 6 = Sabtu
    } catch (error) {
        console.error('Error checking Saturday:', error);
        return false;
    }
}

// Fungsi untuk mendapatkan aturan berdasarkan hari dan kategori
function getDayRulesWithCategory(dateString, category) {
    if (isSaturday(dateString) && category === 'K3') {
        return SATURDAY_RULES.k3SpecialHours;
    } else if (isSaturday(dateString)) {
        return SATURDAY_RULES;
    } else if (isFriday(dateString)) {
        return FRIDAY_RULES;
    } else {
        return NORMAL_DAY_RULES;
    }
}

// Helper functions untuk tabel Excel
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

// Fungsi untuk menghitung lembur hari Sabtu dengan aturan khusus
function calculateSaturdayOvertime(data) {
    const saturdayData = data.filter(item => isSaturday(item.tanggal));
    
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
        
        // Pembulatan jam lembur DENGAN BATAS 7 JAM
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

function formatExcelDate(dateString) {
    try {
        const [day, month, year] = dateString.split('/');
        const shortYear = year.slice(-2);
        return `${day.padStart(2, '0')}-${month.padStart(2, '0')}-${shortYear}`;
    } catch (error) {
        return dateString;
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
                totalLemburDesimal: 0, // Tambah field desimal
                totalLemburBulat: 0,   // Tambah field bulat
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
            // Ada hari yang lembur > 7 jam
            record.totalLemburFormatted = `${record.totalLemburBulat} jam (dibatasi maksimal 7 jam/hari)`;
        } else {
            record.totalLemburFormatted = `${record.totalLemburBulat} jam (${record.totalLemburDesimal.toFixed(2)})`;
        }
        
        record.fridayLemburFormatted = `${record.fridayLemburBulat} jam (${record.fridayOvertime.toFixed(2)})`;
        record.otherDaysLemburFormatted = `${record.otherDaysLemburBulat} jam (${record.otherDaysOvertime.toFixed(2)})`;
        
        // Info detail
        if (record.totalLemburDesimal > 7 * data.filter(item => item.nama === employee).length) {
            record.infoDetail = `Dibatasi maksimal 7 jam/hari (${record.totalLemburDesimal.toFixed(2)} jam total)`;
        } else {
            record.infoDetail = `${record.totalLemburDesimal.toFixed(2)} jam → ${record.totalLemburBulat} jam (dibulatkan)`;
        }
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
    
    if (file.size > 10 * 1024 * 1024) { // 10MB limit
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
        previewHtml += `
            <div style="margin-bottom: 0.5rem; padding-bottom: 0.5rem; border-bottom: 1px solid #eee;">
                <strong>${record.nama}</strong> - ${formatDate(record.tanggal)} ${isFridayDay ? '<span style="color: #9b59b6; font-weight: bold;">(JUMAT)</span>' : ''}<br>
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
        
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>
                ${formatDate(item.tanggal)}
                <br><small style="color: ${isFridayDay ? '#9b59b6' : '#666'};">${getDayName(item.tanggal)} ${isFridayDay ? ' (JUMAT)' : ''}</small>
            </td>
            <td>${item.jamMasuk}</td>
            <td>${item.jamKeluar || '-'}</td>
            <td>${item.durasi ? item.durasi.toFixed(2) + ' jam' : '-'}</td>
        `;
        
        if (isFridayDay) {
            row.classList.add('friday-row');
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
        
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>
                ${formatDate(item.tanggal)}
                ${item.isFriday ? '<br><small style="color: #9b59b6; font-weight: bold;">(JUMAT)</small>' : ''}
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
                ${jamLemburBulat > 0 ? jamLemburBulat + ' jam' : item.jamLembur}
                ${isCapped ? 
                    '<br><small style="color: #e67e22; font-weight: bold;">(DIBATASI MAKSIMAL 7 JAM)</small>' : 
                    ''}
                ${item.isFriday && item.jamLemburDesimal > 0 ? 
                    '<br><small style="color: #9b59b6;">(pulang > 15:00)</small>' : 
                    ''}
                ${item.jamLemburDesimal > 0 && !isCapped ? 
                    `<br><small style="color: #666; font-size: 0.8rem;">(${item.jamLemburDesimal.toFixed(2)} jam desimal)</small>` : 
                    ''}
                ${isCapped ? 
                    `<br><small style="color: #666; font-size: 0.8rem;">(${item.jamLemburDesimal.toFixed(2)} jam → dibatasi 7 jam)</small>` : 
                    ''}
            </td>
            <td>
                <span style="color: ${item.jamLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-size: 0.85rem;">
                    ${item.keterangan}
                    ${isCapped ? '<br><small style="color: #e67e22; font-weight: bold;">(LEMBUR DIBATASI MAKSIMAL 7 JAM)</small>' : 
                     item.jamLemburDesimal > 0 ? '<br><small style="color: #666;">(dibulatkan: ' + jamLemburBulat + ' jam)</small>' : ''}
                </span>
            </td>
        `;
        
        if (item.isFriday) {
            row.classList.add('friday-row');
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
                • Jam kerja normal: 8 jam/hari<br>
                • Lembur dihitung setelah 8 jam kerja<br>
                • Minimal lembur: 10 menit<br>
                • Maksimal lembur: 7 jam/hari<br>
                • Pembulatan jam: 0.5 ke atas → 1
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
}, 10000); // 10 second timeout
