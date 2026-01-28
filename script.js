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

// ============================
// KONFIGURASI ATURAN BARU - SENIN-KAMIS
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

// DOM Elements (GLOBAL - hanya dideklarasikan sekali)
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
// FUNGSI PEMBULATAN BARU
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

// Fungsi untuk membulatkan dan format jam lembur DENGAN BATAS 7 JAM
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
        
        // Dapatkan aturan berdasarkan hari
        const rules = getDayRules(dateString);
        
        // PERUBAHAN: Gunakan jam masuk sebenarnya, bukan jam 07:00
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
        
        // Gunakan jam masuk sebenarnya
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
        
        // KONVERSI KE JAM (DESIMAL)
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
// Fungsi untuk mendapatkan jam masuk efektif dengan aturan Jumat
function getEffectiveInTimeWithDayRules(jamMasuk, dateString) {
    if (!jamMasuk || !dateString) return jamMasuk;
    
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
        
        // Dapatkan aturan berdasarkan hari
        const rules = getDayRules(dateString);
        
        const [startHour, startMinute] = rules.workStartTime.split(':').map(Number);
        const workStartMinutes = startHour * 60 + startMinute;
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        
        // Jika masuk sebelum jam mulai efektif, gunakan jam mulai efektif
        if (inMinutes < workStartMinutes) {
            return rules.workStartTime;
        }
        
        // Jika masuk setelah jam mulai efektif, gunakan jam masuk sebenarnya
        return jamMasuk;
        
    } catch (error) {
        console.error('Error getting effective in time:', error);
        return jamMasuk;
    }
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

function setupEventListeners() {
    console.log('Setting up event listeners...');
    
    try {
        // File upload
        const excelFileInput = document.getElementById('excel-file');
        const browseBtn = document.getElementById('browse-btn');
        const uploadArea = document.getElementById('upload-area');
        const processBtn = document.getElementById('process-data');
        const cancelUploadBtn = document.getElementById('cancel-upload');
        
        if (excelFileInput) excelFileInput.addEventListener('change', handleFileSelect);
        if (browseBtn) browseBtn.addEventListener('click', () => excelFileInput?.click());
        if (uploadArea) uploadArea.addEventListener('click', () => excelFileInput?.click());
        if (processBtn) processBtn.addEventListener('click', processData);
        if (cancelUploadBtn) cancelUploadBtn.addEventListener('click', cancelUpload);
        
    
    // Tabs
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            switchTab(tabId);
        });
    });
    
    // Modal
    const helpBtn = document.getElementById('help-btn');
    const closeHelpBtns = document.querySelectorAll('#close-help, #close-help-btn');
    const helpModal = document.getElementById('help-modal');
    
    if (helpBtn) helpBtn.addEventListener('click', () => helpModal?.classList.add('active'));
    closeHelpBtns.forEach(btn => {
        btn.addEventListener('click', () => helpModal?.classList.remove('active'));
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
 console.log('All event listeners setup complete!');
    } catch (error) {
        console.error('Error in setupEventListeners:', error);
    }
}

// Fungsi untuk menghitung total lembur berdasarkan hari
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
// FUNGSI PERHITUNGAN BARU DENGAN ATURAN JUMAT
// ============================

// Calculate overtime per day - DENGAN ATURAN BARU
function calculateOvertimePerDay(data, workHours = 8) {
    const result = data.map(record => {
        // Hitung total jam kerja (mulai dari jam 07:00)
        const hoursWorked = calculateHoursWithFridayRules(
            record.jamMasuk, 
            record.jamKeluar, 
            record.tanggal
        );
        
        // Hitung lembur dengan aturan baru (berdasarkan jam pulang dan hari)
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
        
        // Tentukan jam masuk efektif berdasarkan hari
        const effectiveInTime = getEffectiveInTimeWithDayRules(record.jamMasuk, record.tanggal);
        
        // Cek apakah hari Jumat
        const isFridayDay = isFriday(record.tanggal);
        
        // Keterangan khusus dengan info hari
        let keterangan = 'Tidak lembur';
        let dayInfo = isFridayDay ? ' (JUMAT)' : '';
        
        if (jamLemburDesimal > 0) {
            if (isFridayDay) {
                keterangan = `Lembur ${jamLemburDisplay} (Jumat - kerja > 8 jam dari jam 07:00)`;
            } else {
                keterangan = `Lembur ${jamLemburDisplay} (kerja > 8 jam dari jam 07:00)`;
            }
            
            if (effectiveInTime !== record.jamMasuk) {
                keterangan += ` (masuk efektif: ${effectiveInTime})`;
            }
        } else if (effectiveInTime !== record.jamMasuk) {
            keterangan = `Masuk efektif: ${effectiveInTime}${dayInfo}`;
        } else if (isFridayDay) {
            keterangan = `Hari Jumat - Kerja 8 jam normal`;
        } else {
            keterangan = `Kerja ≤ 8 jam dari jam 07:00`;
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
            workStartEffective: '07:00',
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

function formatExcelDate(dateString) {
    try {
        const [day, month, year] = dateString.split('/');
        const shortYear = year.slice(-2);
        return `${day.padStart(2, '0')}-${month.padStart(2, '0')}-${shortYear}`;
    } catch (error) {
        return dateString;
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

// Hitung total jam lembur per karyawan DENGAN PEMBULATAN
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
    
    // Hitung total dengan pembulatan
    Object.keys(summary).forEach(employee => {
        const record = summary[employee];
        
        // Hitung versi bulat
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
        
        // Format jam dengan pembulatan
        record.totalLemburFormatted = `${record.totalLemburBulat} jam (${record.totalLemburDesimal.toFixed(2)})`;
        record.fridayLemburFormatted = `${record.fridayLemburBulat} jam (${record.fridayOvertime.toFixed(2)})`;
        record.otherDaysLemburFormatted = `${record.otherDaysLemburBulat} jam (${record.otherDaysOvertime.toFixed(2)})`;
        
        // Info detail
        record.infoDetail = `${record.totalLemburDesimal.toFixed(2)} jam → ${record.totalLemburBulat} jam (dibulatkan)`;
    });
    
    return Object.values(summary);
}

// ============================
// FUNGSI BARU: DOWNLOAD TERPISAH JUMAT DAN HARI LAIN
// ============================

// Fungsi untuk menampilkan modal pilihan download terpisah
function showSeparatedDownloadOptions() {
    if (processedData.length === 0) {
        showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
        return;
    }
    
    // Pisahkan data
    const separated = calculateOvertimeByDayType(processedData);
    
    // Hitung statistik
    const fridayStats = separated.friday;
    const otherDaysStats = separated.otherDays;
    
    // Buat modal untuk pilihan download terpisah
    const modalHtml = `
        <div class="modal" id="separated-download-modal">
            <div class="modal-content" style="max-width: 700px;">
                <div class="modal-header">
                    <h3><i class="fas fa-calendar-alt"></i> Download Laporan Terpisah per Hari</h3>
                    <button class="modal-close" id="close-separated-download">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="download-stats">
                        <div class="stat-card" style="background: linear-gradient(135deg, #9b59b6 0%, #8e44ad 100%); color: white; padding: 1rem; border-radius: 8px; margin-bottom: 1rem;">
                            <h4><i class="fas fa-calendar-day"></i> Data Hari Jumat</h4>
                            <div style="display: flex; justify-content: space-between; margin-top: 0.5rem;">
                                <div>
                                    <div><strong>Total Hari:</strong> ${fridayStats.totalDays}</div>
                                    <div><strong>Total Karyawan:</strong> ${fridayStats.totalEmployees}</div>
                                </div>
                                <div>
                                    <div><strong>Total Jam Lembur:</strong> ${formatHoursToDisplay(fridayStats.totalOvertime)}</div>
                                    <div><strong>Pulang Normal:</strong> 15:00</div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="stat-card" style="background: linear-gradient(135deg, #3498db 0%, #2980b9 100%); color: white; padding: 1rem; border-radius: 8px; margin-bottom: 1.5rem;">
                            <h4><i class="fas fa-calendar-week"></i> Data Senin-Kamis</h4>
                            <div style="display: flex; justify-content: space-between; margin-top: 0.5rem;">
                                <div>
                                    <div><strong>Total Hari:</strong> ${otherDaysStats.totalDays}</div>
                                    <div><strong>Total Karyawan:</strong> ${otherDaysStats.totalEmployees}</div>
                                </div>
                                <div>
                                    <div><strong>Total Jam Lembur:</strong> ${formatHoursToDisplay(otherDaysStats.totalOvertime)}</div>
                                    <div><strong>Pulang Normal:</strong> 16:00</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="download-options-separated">
                        <div class="option-card" id="option-friday-only">
                            <div class="option-icon">
                                <i class="fas fa-calendar-day" style="color: #9b59b6; font-size: 2.5rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Download Data Jumat Saja</h4>
                                <p>File Excel berisi hanya data hari Jumat</p>
                                <ul style="text-align: left; margin-top: 0.5rem;">
                                    <li>Data lembur hari Jumat saja</li>
                                    <li>Format tabel per orang</li>
                                    <li>Catatan khusus: Pulang normal jam 15:00</li>
                                    <li>Worksheet summary terpisah</li>
                                </ul>
                            </div>
                        </div>
                        
                        <div class="option-card" id="option-other-days-only">
                            <div class="option-icon">
                                <i class="fas fa-calendar-week" style="color: #3498db; font-size: 2.5rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Download Data Senin-Kamis</h4>
                                <p>File Excel berisi data hari Senin sampai Kamis</p>
                                <ul style="text-align: left; margin-top: 0.5rem;">
                                    <li>Data lembur hari biasa (Senin-Kamis)</li>
                                    <li>Format tabel per orang</li>
                                    <li>Catatan: Pulang normal jam 16:00</li>
                                    <li>Worksheet summary terpisah</li>
                                </ul>
                            </div>
                        </div>
                        
                        <div class="option-card" id="option-both-separated">
                            <div class="option-icon">
                                <i class="fas fa-calendar-alt" style="color: #2ecc71; font-size: 2.5rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Download Keduanya (Terpisah)</h4>
                                <p>Dua file Excel terpisah untuk Jumat dan hari lain</p>
                                <ul style="text-align: left; margin-top: 0.5rem;">
                                    <li>File 1: Data Jumat</li>
                                    <li>File 2: Data Senin-Kamis</li>
                                    <li>Masing-masing format lengkap</li>
                                    <li>Summary terpisah</li>
                                </ul>
                            </div>
                        </div>
                        
                        <div class="option-card" id="option-all-together">
                            <div class="option-icon">
                                <i class="fas fa-file-excel" style="color: #217346; font-size: 2.5rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Download Semua (Gabungan)</h4>
                                <p>File Excel tunggal dengan semua data</p>
                                <ul style="text-align: left; margin-top: 0.5rem;">
                                    <li>Semua data dalam satu file</li>
                                    <li>Worksheet terpisah: Jumat dan hari lain</li>
                                    <li>Worksheet summary gabungan</li>
                                    <li>Rekap total semua hari</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="cancel-separated-download">Batal</button>
                </div>
            </div>
        </div>
    `;
    
    // Tambahkan modal ke body
    const existingModal = document.getElementById('separated-download-modal');
    if (existingModal) {
        existingModal.remove();
    }
    
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    const modal = document.getElementById('separated-download-modal');
    modal.classList.add('active');
    
    // Event listeners
    document.getElementById('close-separated-download').addEventListener('click', () => {
        modal.remove();
    });
    
    document.getElementById('cancel-separated-download').addEventListener('click', () => {
        modal.remove();
    });
    
    // Pilihan 1: Download Jumat saja
    document.getElementById('option-friday-only').addEventListener('click', () => {
        modal.remove();
        downloadFridayOnly();
    });
    
    // Pilihan 2: Download Senin-Kamis saja
    document.getElementById('option-other-days-only').addEventListener('click', () => {
        modal.remove();
        downloadOtherDaysOnly();
    });
    
    // Pilihan 3: Download kedua file terpisah
    document.getElementById('option-both-separated').addEventListener('click', () => {
        modal.remove();
        downloadBothSeparated();
    });
    
    // Pilihan 4: Download semua dalam satu file
    document.getElementById('option-all-together').addEventListener('click', () => {
        modal.remove();
        downloadAllTogether();
    });
}

// Fungsi untuk download hanya data Jumat
function downloadFridayOnly() {
    try {
        const separated = calculateOvertimeByDayType(processedData);
        
        if (separated.friday.data.length === 0) {
            showNotification('Tidak ada data lembur untuk hari Jumat.', 'info');
            return;
        }
        
        // Generate Excel untuk Jumat saja
        generateFridayOnlyExcel(separated.friday.data);
        showNotification('File Excel data Jumat berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error downloading Friday data:', error);
        showNotification('Gagal mengunduh data Jumat: ' + error.message, 'error');
    }
}

// Fungsi untuk download hanya data Senin-Kamis
function downloadOtherDaysOnly() {
    try {
        const separated = calculateOvertimeByDayType(processedData);
        
        if (separated.otherDays.data.length === 0) {
            showNotification('Tidak ada data lembur untuk hari Senin-Kamis.', 'info');
            return;
        }
        
        // Generate Excel untuk Senin-Kamis saja
        generateOtherDaysOnlyExcel(separated.otherDays.data);
        showNotification('File Excel data Senin-Kamis berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error downloading other days data:', error);
        showNotification('Gagal mengunduh data Senin-Kamis: ' + error.message, 'error');
    }
}

// Fungsi untuk download kedua file terpisah
function downloadBothSeparated() {
    try {
        const separated = calculateOvertimeByDayType(processedData);
        
        if (separated.friday.data.length === 0 && separated.otherDays.data.length === 0) {
            showNotification('Tidak ada data lembur untuk diunduh.', 'info');
            return;
        }
        
        // Generate kedua file
        if (separated.friday.data.length > 0) {
            generateFridayOnlyExcel(separated.friday.data);
        }
        
        setTimeout(() => {
            if (separated.otherDays.data.length > 0) {
                generateOtherDaysOnlyExcel(separated.otherDays.data);
                showNotification('Kedua file Excel berhasil diunduh!', 'success');
            }
        }, 500);
        
    } catch (error) {
        console.error('Error downloading separated data:', error);
        showNotification('Gagal mengunduh file terpisah: ' + error.message, 'error');
    }
}

// Fungsi untuk download semua dalam satu file
function downloadAllTogether() {
    try {
        if (processedData.length === 0) {
            showNotification('Tidak ada data lembur untuk diunduh.', 'warning');
            return;
        }
        
        // Generate file dengan worksheet terpisah
        generateAllInOneExcel(processedData);
        showNotification('File Excel gabungan berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error downloading all data:', error);
        showNotification('Gagal mengunduh file gabungan: ' + error.message, 'error');
    }
}

// Fungsi untuk menampilkan modal download lembur Sabtu
function showSaturdayDownloadModal() {
    if (processedData.length === 0) {
        showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
        return;
    }
    
    // Hitung data Sabtu
    const saturdayResult = calculateSaturdayOvertime(processedData);
    
    const modalHtml = `
        <div class="modal" id="saturday-download-modal">
            <div class="modal-content" style="max-width: 700px;">
                <div class="modal-header">
                    <h3><i class="fas fa-calendar-day"></i> Download Laporan Lembur Hari Sabtu</h3>
                    <button class="modal-close" id="close-saturday-download">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="saturday-info" style="background: linear-gradient(135deg, #e67e22 0%, #d35400 100%); color: white; padding: 1rem; border-radius: 8px; margin-bottom: 1.5rem;">
                        <h4><i class="fas fa-info-circle"></i> Aturan Khusus Hari Sabtu:</h4>
                        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-top: 0.5rem;">
                            <div>
                                <strong>Untuk Semua Karyawan:</strong>
                                <ul style="margin-top: 0.5rem; padding-left: 1rem;">
                                    <li>Jam kerja normal: 6 jam</li>
                                    <li>Masuk bebas jam berapa saja</li>
                                    <li>Lembur ≥ 10 menit setelah 6 jam kerja</li>
                                </ul>
                            </div>
                            <div>
                                <strong>Khusus Kategori K3:</strong>
                                <ul style="margin-top: 0.5rem; padding-left: 1rem;">
                                    <li>Kerja dari: 07:00 - 22:00</li>
                                    <li>Total: 15 jam kerja</li>
                                    <li>Lembur ≥ 10 menit setelah jam 22:00</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                    
                    <div class="saturday-stats" style="margin-bottom: 1.5rem;">
                        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem;">
                            <div style="background: #fff3cd; padding: 1rem; border-radius: 8px; text-align: center;">
                                <div style="font-size: 2rem; font-weight: bold; color: #e67e22;">${saturdayResult.totalEmployees}</div>
                                <div style="font-size: 0.9rem;">Karyawan Masuk</div>
                            </div>
                            <div style="background: #ffeaa7; padding: 1rem; border-radius: 8px; text-align: center;">
                                <div style="font-size: 2rem; font-weight: bold; color: #e67e22;">${saturdayResult.totalOvertime.toFixed(2)}</div>
                                <div style="font-size: 0.9rem;">Jam Lembur (Desimal)</div>
                            </div>
                            <div style="background: #fab1a0; padding: 1rem; border-radius: 8px; text-align: center;">
                                <div style="font-size: 2rem; font-weight: bold; color: #e67e22;">Rp ${Math.round(saturdayResult.totalGaji).toLocaleString('id-ID')}</div>
                                <div style="font-size: 0.9rem;">Total Gaji Lembur</div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="download-options-saturday">
                        <div class="option-card" id="option-saturday-detail">
                            <div class="option-icon">
                                <i class="fas fa-file-alt" style="color: #e67e22; font-size: 2.5rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Download Detail per Hari</h4>
                                <p>File Excel berisi detail lembur per hari Sabtu</p>
                                <ul style="text-align: left; margin-top: 0.5rem;">
                                    <li>Detail per hari kerja Sabtu</li>
                                    <li>Format tabel lengkap</li>
                                    <li>Perhitungan jam normal dan lembur</li>
                                    <li>Kolom keterangan khusus Sabtu</li>
                                </ul>
                            </div>
                        </div>
                        
                        <div class="option-card" id="option-saturday-summary">
                            <div class="option-icon">
                                <i class="fas fa-users" style="color: #e67e22; font-size: 2.5rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Download Summary per Karyawan</h4>
                                <p>File Excel berisi ringkasan lembur per karyawan</p>
                                <ul style="text-align: left; margin-top: 0.5rem;">
                                    <li>Ringkasan per karyawan</li>
                                    <li>Total jam lembur yang dibulatkan</li>
                                    <li>Perhitungan gaji lengkap</li>
                                    <li>Worksheet rekap kategori</li>
                                </ul>
                            </div>
                        </div>
                        
                        <div class="option-card" id="option-saturday-k3">
                            <div class="option-icon">
                                <i class="fas fa-hard-hat" style="color: #34495e; font-size: 2.5rem;"></i>
                            </div>
                            <div class="option-content">
                                <h4>Download Khusus K3</h4>
                                <p>File Excel khusus data K3 hari Sabtu</p>
                                <ul style="text-align: left; margin-top: 0.5rem;">
                                    <li>Data khusus kategori K3</li>
                                    <li>Format 07:00 - 22:00</li>
                                    <li>Perhitungan lembur setelah 22:00</li>
                                    <li>Worksheet terpisah</li>
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="cancel-saturday-download">Batal</button>
                </div>
            </div>
        </div>
    `;
    
    // Tambahkan modal ke body
    const existingModal = document.getElementById('saturday-download-modal');
    if (existingModal) {
        existingModal.remove();
    }
    
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    const modal = document.getElementById('saturday-download-modal');
    modal.classList.add('active');
    
    // Event listeners
    document.getElementById('close-saturday-download').addEventListener('click', () => {
        modal.remove();
    });
    
    document.getElementById('cancel-saturday-download').addEventListener('click', () => {
        modal.remove();
    });
    
    // Pilihan 1: Download detail per hari
    document.getElementById('option-saturday-detail').addEventListener('click', () => {
        modal.remove();
        downloadSaturdayDetail(saturdayResult.data);
    });
    
    // Pilihan 2: Download summary per karyawan
    document.getElementById('option-saturday-summary').addEventListener('click', () => {
        modal.remove();
        downloadSaturdaySummary(saturdayResult.summary);
    });
    
    // Pilihan 3: Download khusus K3
    document.getElementById('option-saturday-k3').addEventListener('click', () => {
        modal.remove();
        downloadSaturdayK3(saturdayResult.data);
    });
}

// ============================
// FUNGSI GENERATE EXCEL TERPISAH
// ============================

// Generate Excel hanya untuk data Jumat
function generateFridayOnlyExcel(fridayData) {
    try {
        // Buat workbook baru
        const workbook = XLSX.utils.book_new();
        
        // Buat worksheet untuk setiap karyawan di hari Jumat
        const employeeGroups = {};
        fridayData.forEach(item => {
            if (!employeeGroups[item.nama]) {
                employeeGroups[item.nama] = [];
            }
            employeeGroups[item.nama].push(item);
        });
        
        // Worksheet untuk setiap karyawan
        Object.keys(employeeGroups).forEach(employeeName => {
            const records = employeeGroups[employeeName];
            const category = employeeCategories[employeeName] || 'STAFF';
            const rate = overtimeRates[category];
            
            // Urutkan berdasarkan tanggal
            records.sort((a, b) => {
                const dateA = a.tanggal.split('/').reverse().join('-');
                const dateB = b.tanggal.split('/').reverse().join('-');
                return new Date(dateA) - new Date(dateB);
            });
            
            // Hitung total dengan pembulatan
            const totalLembur = records.reduce((sum, item) => sum + (item.jamLemburDesimal || 0), 0);
            const totalLemburBulat = roundOvertimeHours(totalLembur);
            const totalGaji = totalLemburBulat * rate;
            
            // Buat data worksheet
            const wsData = [];
            
            // Header
            wsData.push([`LEMBUR KARYAWAN - ${employeeName.toUpperCase()} (HARI JUMAT)`]);
            wsData.push(['Pulang Normal: 15:00 | Lembur ≥ 10 menit setelah 15:00']);
            wsData.push(['Catatan: Jam lembur telah dibulatkan (0.5 ke atas → 1, di bawah 0.5 → 0)']);
            wsData.push([]);
            
            // Header tabel
            wsData.push(['Tanggal', 'Jam Masuk', 'Jam Keluar', 'Jam Normal', 'Jam Lembur (Desimal)', 'Jam Lembur (Bulat)', 'Keterangan']);
            
            // Data
            records.forEach(record => {
                const jamLemburBulat = roundOvertimeHours(record.jamLemburDesimal);
                wsData.push([
                    formatExcelDate(record.tanggal),
                    record.jamMasuk,
                    record.jamKeluar,
                    `${record.jamNormal.toFixed(2)} jam`,
                    record.jamLemburDesimal > 0 ? `${record.jamLemburDesimal.toFixed(2)} jam` : '0 jam',
                    jamLemburBulat > 0 ? `${jamLemburBulat} jam` : '0 jam',
                    record.keterangan
                ]);
            });
            
            // Total
            wsData.push([]);
            wsData.push(['TOTAL JUMAT:', '', '', '', 
                `${totalLembur.toFixed(2)} jam`, 
                `${totalLemburBulat} jam (dibulatkan)`, 
                `Rp ${Math.round(totalGaji).toLocaleString('id-ID')}`]);
            
            // Buat worksheet
            const worksheet = XLSX.utils.aoa_to_sheet(wsData);
            
            // Set column widths
            worksheet['!cols'] = [
                { wch: 15 }, // Tanggal
                { wch: 12 }, // Jam Masuk
                { wch: 12 }, // Jam Keluar
                { wch: 12 }, // Jam Normal
                { wch: 18 }, // Jam Lembur (Desimal)
                { wch: 18 }, // Jam Lembur (Bulat)
                { wch: 30 }  // Keterangan
            ];
            
            // Tambahkan ke workbook
            const sheetName = employeeName.substring(0, 30).replace(/[\\/*\[\]:?]/g, '');
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName + ' (Jumat)');
        });
        
        // Buat worksheet summary untuk Jumat
        const summaryData = generateFridaySummary(fridayData);
        const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
        
        // Set column widths summary
        summaryWs['!cols'] = [
            { wch: 5 },  // No
            { wch: 25 }, // Nama Karyawan
            { wch: 15 }, // Kategori
            { wch: 20 }, // Total Jam Lembur (Desimal)
            { wch: 20 }, // Total Jam Lembur (Bulat)
            { wch: 25 }, // Total Gaji Lembur
            { wch: 15 }  // Rate
        ];
        
        XLSX.utils.book_append_sheet(workbook, summaryWs, 'SUMMARY JUMAT');
        
        // Simpan file
        const today = new Date();
        const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
        const filename = `lembur_jumat_${formattedDate}.xlsx`;
        
        XLSX.writeFile(workbook, filename);
        
    } catch (error) {
        console.error('Error generating Friday Excel:', error);
        throw error;
    }
}

// Fungsi untuk download detail lembur Sabtu
function downloadSaturdayDetail(data) {
    try {
        if (data.length === 0) {
            showNotification('Tidak ada data lembur untuk hari Sabtu.', 'info');
            return;
        }
        
        // Buat workbook baru
        const workbook = XLSX.utils.book_new();
        
        // Data untuk worksheet
        const wsData = [];
        
        // Header
        wsData.push(['LAPORAN LEMBUR KHUSUS HARI SABTU']);
        wsData.push(['FAKULTAS TEKNIK - UNIVERSITAS LANGLANGBUANA']);
        wsData.push(['ATURAN: Semua karyawan kerja 6 jam, K3 kerja 07:00-22:00']);
        wsData.push(['Catatan: Jam lembur telah dibulatkan (0.5 ke atas → 1, di bawah 0.5 → 0)']);
        wsData.push([]);
        
        // Header tabel
        wsData.push(['No', 'Nama Karyawan', 'Kategori', 'Tanggal', 'Hari', 'Jam Masuk', 'Jam Keluar', 'Durasi Kerja', 'Jam Normal (Sabtu)', 'Jam Lembur (Desimal)', 'Jam Lembur (Bulat)', 'Gaji Lembur', 'Keterangan']);
        
        // Data
        data.forEach((item, index) => {
            wsData.push([
                index + 1,
                item.nama,
                item.kategori,
                formatExcelDate(item.tanggal),
                'Sabtu',
                item.jamMasuk,
                item.jamKeluar,
                item.durasiFormatted,
                `${item.jamNormalSabtu.toFixed(2)} jam`,
                item.jamLemburSabtu > 0 ? `${item.jamLemburSabtu.toFixed(2)} jam` : '0 jam',
                item.jamLemburSabtuBulat > 0 ? `${item.jamLemburSabtuBulat} jam` : '0 jam',
                item.gajiLemburSabtu > 0 ? formatCurrency(item.gajiLemburSabtu) : 'Rp 0',
                item.keteranganSabtu
            ]);
        });
        
        // Total
        const totalJam = data.reduce((sum, item) => sum + item.jamLemburSabtu, 0);
        const totalJamBulat = data.reduce((sum, item) => sum + item.jamLemburSabtuBulat, 0);
        const totalGaji = data.reduce((sum, item) => sum + item.gajiLemburSabtu, 0);
        
        wsData.push([]);
        wsData.push(['TOTAL SABTU', '', '', '', '', '', '', '', '', 
            `${totalJam.toFixed(2)} jam`, 
            `${totalJamBulat} jam (dibulatkan)`, 
            formatCurrency(totalGaji), '']);
        
        // Buat worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(wsData);
        
        // Set column widths
        worksheet['!cols'] = [
            { wch: 5 },   // No
            { wch: 20 },  // Nama
            { wch: 12 },  // Kategori
            { wch: 15 },  // Tanggal
            { wch: 10 },  // Hari
            { wch: 10 },  // Jam Masuk
            { wch: 10 },  // Jam Keluar
            { wch: 15 },  // Durasi Kerja
            { wch: 20 },  // Jam Normal
            { wch: 18 },  // Jam Lembur (Desimal)
            { wch: 18 },  // Jam Lembur (Bulat)
            { wch: 20 },  // Gaji Lembur
            { wch: 35 }   // Keterangan
        ];
        
        // Tambahkan ke workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, 'LEMBUR SABTU');
        
        // Simpan file
        const today = new Date();
        const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
        const filename = `lembur_sabtu_detail_${formattedDate}.xlsx`;
        
        XLSX.writeFile(workbook, filename);
        showNotification('File Excel lembur Sabtu berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error downloading Saturday detail:', error);
        showNotification('Gagal mengunduh data Sabtu: ' + error.message, 'error');
    }
}

// Fungsi untuk download summary lembur Sabtu
function downloadSaturdaySummary(summary) {
    try {
        if (summary.length === 0) {
            showNotification('Tidak ada data lembur untuk hari Sabtu.', 'info');
            return;
        }
        
        // Buat workbook baru
        const workbook = XLSX.utils.book_new();
        
        // Data untuk worksheet
        const wsData = [];
        
        // Header
        wsData.push(['REKAPITULASI LEMBUR HARI SABTU']);
        wsData.push(['FAKULTAS TEKNIK - UNIVERSITAS LANGLANGBUANA']);
        wsData.push(['ATURAN KHUSUS: Kerja 6 jam, K3: 07:00-22:00']);
        wsData.push(['Catatan: Jam lembur telah dibulatkan (0.5 ke atas → 1, di bawah 0.5 → 0)']);
        wsData.push([]);
        
        // Header tabel summary
        wsData.push(['No', 'Nama Karyawan', 'Kategori', 'Rate per Jam', 'Total Jam Lembur (Desimal)', 'Total Jam Lembur (Bulat)', 'Total Gaji Lembur']);
        
        // Data summary
        summary.forEach((item, index) => {
            wsData.push([
                index + 1,
                item.nama,
                item.kategori,
                `Rp ${item.rate.toLocaleString('id-ID')}`,
                item.totalJam > 0 ? `${item.totalJam.toFixed(2)} jam` : '0 jam',
                item.totalJamBulat > 0 ? `${item.totalJamBulat} jam` : '0 jam',
                formatCurrency(item.totalGaji)
            ]);
        });
        
        // Total
        const totalJam = summary.reduce((sum, item) => sum + item.totalJam, 0);
        const totalJamBulat = summary.reduce((sum, item) => sum + item.totalJamBulat, 0);
        const totalGaji = summary.reduce((sum, item) => sum + item.totalGaji, 0);
        
        wsData.push([]);
        wsData.push(['TOTAL KESELURUHAN', '', '', '', 
            `${totalJam.toFixed(2)} jam`, 
            `${totalJamBulat} jam (dibulatkan)`, 
            formatCurrency(totalGaji)]);
        
        // Buat worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(wsData);
        
        // Set column widths
        worksheet['!cols'] = [
            { wch: 5 },   // No
            { wch: 25 },  // Nama
            { wch: 15 },  // Kategori
            { wch: 15 },  // Rate
            { wch: 20 },  // Jam Lembur (Desimal)
            { wch: 20 },  // Jam Lembur (Bulat)
            { wch: 25 }   // Total Gaji
        ];
        
        // Tambahkan ke workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, 'SUMMARY SABTU');
        
        // Worksheet untuk per kategori
        const byCategory = {};
        summary.forEach(item => {
            if (!byCategory[item.kategori]) {
                byCategory[item.kategori] = {
                    totalJam: 0,
                    totalJamBulat: 0,
                    totalGaji: 0,
                    count: 0
                };
            }
            byCategory[item.kategori].totalJam += item.totalJam;
            byCategory[item.kategori].totalJamBulat += item.totalJamBulat;
            byCategory[item.kategori].totalGaji += item.totalGaji;
            byCategory[item.kategori].count++;
        });
        
        const categoryData = [];
        categoryData.push(['REKAP PER KATEGORI - HARI SABTU']);
        categoryData.push([]);
        categoryData.push(['Kategori', 'Jumlah Karyawan', 'Total Jam (Desimal)', 'Total Jam (Bulat)', 'Total Gaji']);
        
        Object.keys(byCategory).forEach(category => {
            const data = byCategory[category];
            categoryData.push([
                category,
                data.count,
                `${data.totalJam.toFixed(2)} jam`,
                `${data.totalJamBulat} jam`,
                formatCurrency(data.totalGaji)
            ]);
        });
        
        const categoryWs = XLSX.utils.aoa_to_sheet(categoryData);
        categoryWs['!cols'] = [
            { wch: 15 },  // Kategori
            { wch: 15 },  // Jumlah
            { wch: 20 },  // Jam (Desimal)
            { wch: 20 },  // Jam (Bulat)
            { wch: 25 }   // Total Gaji
        ];
        
        XLSX.utils.book_append_sheet(workbook, categoryWs, 'REKAP KATEGORI');
        
        // Simpan file
        const today = new Date();
        const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
        const filename = `lembur_sabtu_summary_${formattedDate}.xlsx`;
        
        XLSX.writeFile(workbook, filename);
        showNotification('File Excel summary Sabtu berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error downloading Saturday summary:', error);
        showNotification('Gagal mengunduh summary Sabtu: ' + error.message, 'error');
    }
}

// Fungsi untuk download khusus K3 hari Sabtu
function downloadSaturdayK3(data) {
    try {
        const k3Data = data.filter(item => item.kategori === 'K3');
        
        if (k3Data.length === 0) {
            showNotification('Tidak ada data K3 untuk hari Sabtu.', 'info');
            return;
        }
        
        // Buat workbook baru
        const workbook = XLSX.utils.book_new();
        
        // Data untuk worksheet
        const wsData = [];
        
        // Header khusus K3
        wsData.push(['LAPORAN LEMBUR KHUSUS K3 - HARI SABTU']);
        wsData.push(['FAKULTAS TEKNIK - UNIVERSITAS LANGLANGBUANA']);
        wsData.push(['ATURAN KHUSUS K3: Kerja dari 07:00 sampai 22:00 (15 jam)']);
        wsData.push(['Lembur dihitung setelah jam 22:00, minimal 10 menit']);
        wsData.push(['Catatan: Jam lembur telah dibulatkan (0.5 ke atas → 1, di bawah 0.5 → 0)']);
        wsData.push([]);
        
        // Header tabel
        wsData.push(['No', 'Nama Karyawan', 'Tanggal', 'Jam Masuk', 'Jam Keluar', 'Durasi Kerja', 'Jam Lembur (Desimal)', 'Jam Lembur (Bulat)', 'Gaji Lembur', 'Keterangan']);
        
        // Data K3
        k3Data.forEach((item, index) => {
            wsData.push([
                index + 1,
                item.nama,
                formatExcelDate(item.tanggal),
                item.jamMasuk,
                item.jamKeluar,
                item.durasiFormatted,
                item.jamLemburSabtu > 0 ? `${item.jamLemburSabtu.toFixed(2)} jam` : '0 jam',
                item.jamLemburSabtuBulat > 0 ? `${item.jamLemburSabtuBulat} jam` : '0 jam',
                item.gajiLemburSabtu > 0 ? formatCurrency(item.gajiLemburSabtu) : 'Rp 0',
                `K3 Sabtu: 07:00-22:00 (${item.jamKeluar > '22:00' ? 'Lembur' : 'Tidak lembur'})`
            ]);
        });
        
        // Total K3
        const totalJam = k3Data.reduce((sum, item) => sum + item.jamLemburSabtu, 0);
        const totalJamBulat = k3Data.reduce((sum, item) => sum + item.jamLemburSabtuBulat, 0);
        const totalGaji = k3Data.reduce((sum, item) => sum + item.gajiLemburSabtu, 0);
        
        wsData.push([]);
        wsData.push(['TOTAL K3 SABTU', '', '', '', '', '', 
            `${totalJam.toFixed(2)} jam`, 
            `${totalJamBulat} jam (dibulatkan)`, 
            formatCurrency(totalGaji), '']);
        
        // Buat worksheet
        const worksheet = XLSX.utils.aoa_to_sheet(wsData);
        
        // Set column widths
        worksheet['!cols'] = [
            { wch: 5 },   // No
            { wch: 20 },  // Nama
            { wch: 15 },  // Tanggal
            { wch: 10 },  // Jam Masuk
            { wch: 10 },  // Jam Keluar
            { wch: 15 },  // Durasi Kerja
            { wch: 18 },  // Jam Lembur (Desimal)
            { wch: 18 },  // Jam Lembur (Bulat)
            { wch: 20 },  // Gaji Lembur
            { wch: 30 }   // Keterangan
        ];
        
        // Tambahkan ke workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, 'K3 SABTU');
        
        // Simpan file
        const today = new Date();
        const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
        const filename = `lembur_k3_sabtu_${formattedDate}.xlsx`;
        
        XLSX.writeFile(workbook, filename);
        showNotification('File Excel K3 Sabtu berhasil diunduh!', 'success');
        
    } catch (error) {
        console.error('Error downloading Saturday K3:', error);
        showNotification('Gagal mengunduh data K3 Sabtu: ' + error.message, 'error');
    }
}

// Generate Excel hanya untuk data Senin-Kamis
function generateOtherDaysOnlyExcel(otherDaysData) {
    try {
        // Buat workbook baru
        const workbook = XLSX.utils.book_new();
        
        // Buat worksheet untuk setiap karyawan di hari biasa
        const employeeGroups = {};
        otherDaysData.forEach(item => {
            if (!employeeGroups[item.nama]) {
                employeeGroups[item.nama] = [];
            }
            employeeGroups[item.nama].push(item);
        });
        
        // Worksheet untuk setiap karyawan
        Object.keys(employeeGroups).forEach(employeeName => {
            const records = employeeGroups[employeeName];
            const category = employeeCategories[employeeName] || 'STAFF';
            const rate = overtimeRates[category];
            
            // Urutkan berdasarkan tanggal
            records.sort((a, b) => {
                const dateA = a.tanggal.split('/').reverse().join('-');
                const dateB = b.tanggal.split('/').reverse().join('-');
                return new Date(dateA) - new Date(dateB);
            });
            
            // Hitung total dengan pembulatan
            const totalLembur = records.reduce((sum, item) => sum + (item.jamLemburDesimal || 0), 0);
            const totalLemburBulat = roundOvertimeHours(totalLembur);
            const totalGaji = totalLemburBulat * rate;
            
            // Buat data worksheet
            const wsData = [];
            
            // Header
            wsData.push([`LEMBUR KARYAWAN - ${employeeName.toUpperCase()} (SENIN-KAMIS)`]);
            wsData.push(['Pulang Normal: 16:00 | Lembur ≥ 10 menit setelah 16:00']);
            wsData.push(['Catatan: Jam lembur telah dibulatkan (0.5 ke atas → 1, di bawah 0.5 → 0)']);
            wsData.push([]);
            
            // Header tabel
            wsData.push(['Tanggal', 'Hari', 'Jam Masuk', 'Jam Keluar', 'Jam Normal', 'Jam Lembur (Desimal)', 'Jam Lembur (Bulat)', 'Keterangan']);
            
            // Data
            records.forEach(record => {
                const jamLemburBulat = roundOvertimeHours(record.jamLemburDesimal);
                wsData.push([
                    formatExcelDate(record.tanggal),
                    record.hari,
                    record.jamMasuk,
                    record.jamKeluar,
                    `${record.jamNormal.toFixed(2)} jam`,
                    record.jamLemburDesimal > 0 ? `${record.jamLemburDesimal.toFixed(2)} jam` : '0 jam',
                    jamLemburBulat > 0 ? `${jamLemburBulat} jam` : '0 jam',
                    record.keterangan
                ]);
            });
            
            // Total
            wsData.push([]);
            wsData.push(['TOTAL SENIN-KAMIS:', '', '', '', '', 
                `${totalLembur.toFixed(2)} jam`, 
                `${totalLemburBulat} jam (dibulatkan)`,
                `Rp ${Math.round(totalGaji).toLocaleString('id-ID')}`]);
            
            // Buat worksheet
            const worksheet = XLSX.utils.aoa_to_sheet(wsData);
            
            // Set column widths
            worksheet['!cols'] = [
                { wch: 15 }, // Tanggal
                { wch: 12 }, // Hari
                { wch: 12 }, // Jam Masuk
                { wch: 12 }, // Jam Keluar
                { wch: 12 }, // Jam Normal
                { wch: 18 }, // Jam Lembur (Desimal)
                { wch: 18 }, // Jam Lembur (Bulat)
                { wch: 30 }  // Keterangan
            ];
            
            // Tambahkan ke workbook
            const sheetName = employeeName.substring(0, 30).replace(/[\\/*\[\]:?]/g, '');
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName + ' (Senin-Kamis)');
        });
        
        // Buat worksheet summary untuk Senin-Kamis
        const summaryData = generateOtherDaysSummary(otherDaysData);
        const summaryWs = XLSX.utils.aoa_to_sheet(summaryData);
        
        // Set column widths summary
        summaryWs['!cols'] = [
            { wch: 5 },  // No
            { wch: 25 }, // Nama Karyawan
            { wch: 15 }, // Kategori
            { wch: 20 }, // Total Jam Lembur (Desimal)
            { wch: 20 }, // Total Jam Lembur (Bulat)
            { wch: 25 }, // Total Gaji Lembur
            { wch: 15 }  // Rate
        ];
        
        XLSX.utils.book_append_sheet(workbook, summaryWs, 'SUMMARY SENIN-KAMIS');
        
        // Simpan file
        const today = new Date();
        const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
        const filename = `lembur_senin_kamis_${formattedDate}.xlsx`;
        
        XLSX.writeFile(workbook, filename);
        
    } catch (error) {
        console.error('Error generating other days Excel:', error);
        throw error;
    }
}

// Generate Excel dengan semua data dalam satu file (worksheet terpisah)
function generateAllInOneExcel(allData) {
    try {
        // Pisahkan data
        const separated = calculateOvertimeByDayType(allData);
        
        // Buat workbook baru
        const workbook = XLSX.utils.book_new();
        
        // Worksheet untuk data Jumat
        if (separated.friday.data.length > 0) {
            const fridaySummary = generateFridaySummary(separated.friday.data);
            const fridayWs = XLSX.utils.aoa_to_sheet(fridaySummary);
            fridayWs['!cols'] = [
                { wch: 5 },  // No
                { wch: 25 }, // Nama Karyawan
                { wch: 15 }, // Kategori
                { wch: 20 }, // Total Jam Lembur (Desimal)
                { wch: 20 }, // Total Jam Lembur (Bulat)
                { wch: 25 }, // Total Gaji Lembur
                { wch: 15 }  // Rate
            ];
            XLSX.utils.book_append_sheet(workbook, fridayWs, 'JUMAT SUMMARY');
        }
        
        // Worksheet untuk data Senin-Kamis
        if (separated.otherDays.data.length > 0) {
            const otherDaysSummary = generateOtherDaysSummary(separated.otherDays.data);
            const otherDaysWs = XLSX.utils.aoa_to_sheet(otherDaysSummary);
            otherDaysWs['!cols'] = [
                { wch: 5 },  // No
                { wch: 25 }, // Nama Karyawan
                { wch: 15 }, // Kategori
                { wch: 20 }, // Total Jam Lembur (Desimal)
                { wch: 20 }, // Total Jam Lembur (Bulat)
                { wch: 25 }, // Total Gaji Lembur
                { wch: 15 }  // Rate
            ];
            XLSX.utils.book_append_sheet(workbook, otherDaysWs, 'SENIN-KAMIS SUMMARY');
        }
        
        // Worksheet untuk summary total
        const totalSummary = generateTotalSummary(allData);
        const totalWs = XLSX.utils.aoa_to_sheet(totalSummary);
        totalWs['!cols'] = [
            { wch: 5 },  // No
            { wch: 25 }, // Nama Karyawan
            { wch: 15 }, // Kategori
            { wch: 20 }, // Total Jam (Desimal)
            { wch: 15 }, // Jam (Bulat)
            { wch: 15 }, // Lembur Jumat
            { wch: 15 }, // Lembur Biasa
            { wch: 25 }, // Total Gaji
            { wch: 15 }  // Rate
        ];
        XLSX.utils.book_append_sheet(workbook, totalWs, 'TOTAL SUMMARY');
        
        // Worksheet untuk detail semua data
        const detailData = prepareExportData(allData);
        const detailWs = XLSX.utils.json_to_sheet(detailData);
        detailWs['!cols'] = [
            { wch: 5 },   // No
            { wch: 25 },  // Nama Karyawan
            { wch: 15 },  // Tanggal
            { wch: 12 },  // Hari
            { wch: 12 },  // Jam Masuk
            { wch: 12 },  // Jam Keluar
            { wch: 15 },  // Durasi Kerja
            { wch: 12 },  // Jam Normal
            { wch: 18 },  // Jam Lembur (Desimal)
            { wch: 18 },  // Jam Lembur (Bulat)
            { wch: 30 }   // Keterangan
        ];
        XLSX.utils.book_append_sheet(workbook, detailWs, 'DETAIL DATA');
        
        // Simpan file
        const today = new Date();
        const formattedDate = `${today.getFullYear()}-${(today.getMonth() + 1).toString().padStart(2, '0')}-${today.getDate().toString().padStart(2, '0')}`;
        const filename = `lembur_semua_hari_${formattedDate}.xlsx`;
        
        XLSX.writeFile(workbook, filename);
        
    } catch (error) {
        console.error('Error generating all in one Excel:', error);
        throw error;
    }
}

// Fungsi untuk membuat summary Jumat
function generateFridaySummary(fridayData) {
    const summary = calculateOvertimeSummary(fridayData);
    
    const summaryData = [];
    
    // Header
    summaryData.push(['REKAPITULASI LEMBUR KARYAWAN - HARI JUMAT']);
    summaryData.push(['Pulang Normal: 15:00 | Lembur ≥ 10 menit setelah 15:00']);
    summaryData.push(['Catatan: Jam lembur telah dibulatkan (0.5 ke atas → 1, di bawah 0.5 → 0)']);
    summaryData.push(['FAKULTAS TEKNIK - UNIVERSITAS LANGLANGBUANA']);
    summaryData.push([]);
    summaryData.push(['No', 'Nama Karyawan', 'Kategori', 'Total Jam Lembur (Desimal)', 'Total Jam Lembur (Bulat)', 'Total Gaji Lembur', 'Rate']);
    
    // Data per karyawan
    let totalJamAll = 0;
    let totalJamAllBulat = 0;
    let totalGajiAll = 0;
    
    summary.forEach((item, index) => {
        if (item.fridayOvertime > 0) {
            const jamBulat = roundOvertimeHours(item.fridayOvertime);
            const gaji = jamBulat * item.rate;
            
            summaryData.push([
                index + 1,
                item.nama,
                item.kategori,
                item.fridayOvertime.toFixed(2) + ' jam',
                jamBulat + ' jam',
                formatCurrency(gaji),
                `Rp ${item.rate.toLocaleString('id-ID')}/jam`
            ]);
            
            totalJamAll += item.fridayOvertime;
            totalJamAllBulat += jamBulat;
            totalGajiAll += gaji;
        }
    });
    
    summaryData.push([]);
    summaryData.push(['TOTAL JUMAT', '', '', 
        `${totalJamAll.toFixed(2)} jam`, 
        `${totalJamAllBulat} jam (dibulatkan)`, 
        formatCurrency(totalGajiAll), '']);
    
    return summaryData;
}

// Fungsi untuk membuat summary Senin-Kamis
function generateOtherDaysSummary(otherDaysData) {
    const summary = calculateOvertimeSummary(otherDaysData);
    
    const summaryData = [];
    
    // Header
    summaryData.push(['REKAPITULASI LEMBUR KARYAWAN - SENIN s/d KAMIS']);
    summaryData.push(['Pulang Normal: 16:00 | Lembur ≥ 10 menit setelah 16:00']);
    summaryData.push(['Catatan: Jam lembur telah dibulatkan (0.5 ke atas → 1, di bawah 0.5 → 0)']);
    summaryData.push(['FAKULTAS TEKNIK - UNIVERSITAS LANGLANGBUANA']);
    summaryData.push([]);
    summaryData.push(['No', 'Nama Karyawan', 'Kategori', 'Total Jam Lembur (Desimal)', 'Total Jam Lembur (Bulat)', 'Total Gaji Lembur', 'Rate']);
    
    // Data per karyawan
    let totalJamAll = 0;
    let totalJamAllBulat = 0;
    let totalGajiAll = 0;
    
    summary.forEach((item, index) => {
        if (item.otherDaysOvertime > 0) {
            const jamBulat = roundOvertimeHours(item.otherDaysOvertime);
            const gaji = jamBulat * item.rate;
            
            summaryData.push([
                index + 1,
                item.nama,
                item.kategori,
                item.otherDaysOvertime.toFixed(2) + ' jam',
                jamBulat + ' jam',
                formatCurrency(gaji),
                `Rp ${item.rate.toLocaleString('id-ID')}/jam`
            ]);
            
            totalJamAll += item.otherDaysOvertime;
            totalJamAllBulat += jamBulat;
            totalGajiAll += gaji;
        }
    });
    
    summaryData.push([]);
    summaryData.push(['TOTAL SENIN-KAMIS', '', '', 
        `${totalJamAll.toFixed(2)} jam`, 
        `${totalJamAllBulat} jam (dibulatkan)`, 
        formatCurrency(totalGajiAll), '']);
    
    return summaryData;
}

// Fungsi untuk membuat summary total semua hari
function generateTotalSummary(allData) {
    const summary = calculateOvertimeSummary(allData);
    
    const summaryData = [];
    
    // Header
    summaryData.push(['REKAPITULASI TOTAL LEMBUR KARYAWAN']);
    summaryData.push(['Jumat: Pulang 15:00 | Senin-Kamis: Pulang 16:00']);
    summaryData.push(['Catatan: Jam lembur telah dibulatkan (0.5 ke atas → 1, di bawah 0.5 → 0)']);
    summaryData.push(['FAKULTAS TEKNIK - UNIVERSITAS LANGLANGBUANA']);
    summaryData.push([]);
    summaryData.push(['No', 'Nama Karyawan', 'Kategori', 'Total Jam (Desimal)', 'Jam (Bulat)', 'Jumat', 'Biasa', 'Total Gaji', 'Rate']);
    
    // Data per karyawan
    let totalJamAll = 0;
    let totalJamAllBulat = 0;
    let totalJumatAll = 0;
    let totalBiasaAll = 0;
    let totalGajiAll = 0;
    
    summary.forEach((item, index) => {
        if (item.totalLemburDesimal > 0) {
            const totalBulat = roundOvertimeHours(item.totalLemburDesimal);
            const fridayBulat = roundOvertimeHours(item.fridayOvertime);
            const otherDaysBulat = roundOvertimeHours(item.otherDaysOvertime);
            const totalGaji = totalBulat * item.rate;
            
            summaryData.push([
                index + 1,
                item.nama,
                item.kategori,
                item.totalLemburDesimal.toFixed(2) + ' jam',
                totalBulat + ' jam',
                fridayBulat + ' jam',
                otherDaysBulat + ' jam',
                formatCurrency(totalGaji),
                `Rp ${item.rate.toLocaleString('id-ID')}/jam`
            ]);
            
            totalJamAll += item.totalLemburDesimal;
            totalJamAllBulat += totalBulat;
            totalJumatAll += fridayBulat;
            totalBiasaAll += otherDaysBulat;
            totalGajiAll += totalGaji;
        }
    });
    
    summaryData.push([]);
    summaryData.push(['TOTAL KESELURUHAN', '', '', 
        `${totalJamAll.toFixed(2)} jam`,
        `${totalJamAllBulat} jam (dibulatkan)`,
        `${totalJumatAll} jam`,
        `${totalBiasaAll} jam`,
        formatCurrency(totalGajiAll), '']);
    
    return summaryData;
}

// ============================
// FUNGSI UNTUK TABEL LEMBUR PER ORANG (LIKE EXCEL)
// ============================

// Fungsi untuk format tabel per orang
function createOvertimeTablePerPerson(data) {
    if (!data || data.length === 0) return [];
    
    // Kelompokkan data per orang
    const employeeGroups = {};
    data.forEach(record => {
        const employeeName = record.nama;
        if (!employeeGroups[employeeName]) {
            employeeGroups[employeeName] = [];
        }
        employeeGroups[employeeName].push(record);
    });
    
    const tables = [];
    
    Object.keys(employeeGroups).forEach(employeeName => {
        const records = employeeGroups[employeeName];
        const category = employeeCategories[employeeName] || 'STAFF';
        const rate = overtimeRates[category];
        
        // Urutkan berdasarkan tanggal
        records.sort((a, b) => {
            const dateA = a.tanggal.split('/').reverse().join('-');
            const dateB = b.tanggal.split('/').reverse().join('-');
            return new Date(dateA) - new Date(dateB);
        });
        
        // Hitung total per orang DENGAN PEMBULATAN
        const totalLemburDesimal = records.reduce((sum, item) => sum + (item.jamLemburDesimal || 0), 0);
        const totalLemburBulat = roundOvertimeHours(totalLemburDesimal);
        const totalGaji = totalLemburBulat * rate;
        
        // Buat data untuk tabel per orang - format seperti gambar
        const tableData = [];
        
        // Header untuk setiap orang (lebih sederhana seperti gambar)
        tableData.push([
            `LEMBUR KARYAWAN - ${employeeName.toUpperCase()}`,
            '', '', '', '', '', '', ''
        ]);
        
        tableData.push([
            'BULAN NOVEMBER 2025',
            '', '', '', '', '', '', ''
        ]);
        
        tableData.push([]); // Baris kosong
        
        // Rate bayaran lembur (sesuai gambar)
        tableData.push([
            'RATE BAYARAN LEMBUR',
            '', '', '', '', '', '', ''
        ]);
        
        tableData.push([
            'KABAG/K.TU',
            'Rp 12.500',
            '', '', '', '', '', ''
        ]);
        
        tableData.push([
            'STAF',
            'Rp 10.000',
            '', '', '', '', '', ''
        ]);
        
        tableData.push([
            'K3',
            'Rp 8.000',
            '', '', '', '', '', ''
        ]);
        
        tableData.push([]); // Baris kosong
        
        // Header tabel (sesuai gambar)
        tableData.push([
            'Name',
            'Hari',
            'Tanggal',
            'IN',
            'OUT',
            `JAM KERJA (${currentWorkHours} jam)`,
            'TOTAL (Bulat)',
            'TANDA TANGAN'
        ]);
        
        // Data per hari
        let cumulativeTotalBulat = 0;
        const filteredRecords = records.filter(record => record.jamLemburDesimal > 0);
        
        filteredRecords.forEach((record, index) => {
            const hari = record.hari;
            const jamMasuk = record.jamMasuk || '-';
            const jamKeluar = record.jamKeluar || '-';
            const lemburJamBulat = roundOvertimeHours(record.jamLemburDesimal);
            
            cumulativeTotalBulat += lemburJamBulat;
            
            tableData.push([
                employeeName,
                hari,
                formatExcelDate(record.tanggal),
                jamMasuk,
                jamKeluar,
                currentWorkHours,
                lemburJamBulat,
                record.isFriday ? 'Jumat' : '' // Tambahkan catatan untuk Jumat
            ]);
        });
        
        // Baris kosong jika tidak ada data lembur
        if (filteredRecords.length === 0) {
            tableData.push([
                employeeName,
                '', '', '', '', '', '', ''
            ]);
        }
        
        // Baris total (sesuai format gambar) - MENGGUNAKAN YANG SUDAH DIBULATKAN
        tableData.push([
            '', // Name kosong
            '', // Hari kosong
            '', // Tanggal kosong
            '', // IN kosong
            '', // OUT kosong
            '', // JAM KERJA kosong
            cumulativeTotalBulat, // Total jam lembur yang sudah dibulatkan
            '' // Tanda tangan kosong
        ]);
        
        // Baris jumlah jam lembur (dibulatkan)
        tableData.push([
            '', '', '', '', '', 
            cumulativeTotalBulat, // Jam lembur yang sudah dibulatkan
            '', ''
        ]);
        
        // Baris gaji lembur - MENGGUNAKAN JAM YANG SUDAH DIBULATKAN
        const totalGajiRounded = cumulativeTotalBulat * rate;
        
        tableData.push([
            '', '', '', '', '',
            `Rp ${totalGajiRounded.toLocaleString('id-ID')}`,
            '', ''
        ]);
        
        // Baris gaji lembur (duplikat seperti gambar)
        tableData.push([
            '', '', '', '', '',
            `Rp ${totalGajiRounded.toLocaleString('id-ID')}`,
            '', ''
        ]);
        
        tables.push({
            employee: employeeName,
            category: category,
            rate: rate,
            data: tableData,
            totalLemburDesimal: totalLemburDesimal,
            totalLemburBulat: cumulativeTotalBulat,
            totalGaji: totalGajiRounded
        });
    });
    
    return tables;
}

// ============================
// FUNGSI UNTUK TOMBOL SCROLL KE BAWAH
// ============================

// Fungsi untuk membuat tombol scroll ke bawah
function createScrollToBottomButton() {
    // Cek apakah tombol sudah ada
    if (document.getElementById('scroll-to-bottom-btn')) {
        return;
    }
    
    // Buat tombol
    const button = document.createElement('button');
    button.id = 'scroll-to-bottom-btn';
    button.className = 'scroll-to-bottom-btn';
    button.innerHTML = '<i class="fas fa-arrow-down"></i> Ke Bawah Tabel';
    button.title = 'Scroll ke akhir tabel';
    
    // Tambahkan ke body
    document.body.appendChild(button);
    
    // Event listener untuk tombol
    button.addEventListener('click', scrollToBottomOfTable);
}

// Fungsi untuk scroll ke bawah tabel
function scrollToBottomOfTable() {
    const tableContainer = document.querySelector('.table-responsive');
    const processedTab = document.getElementById('processed-tab');
    
    if (tableContainer && processedTab && processedTab.classList.contains('active')) {
        // Scroll ke bawah container tabel
        tableContainer.scrollTo({
            top: tableContainer.scrollHeight,
            behavior: 'smooth'
        });
        
        // Atau scroll ke elemen terakhir dalam tabel
        const lastRow = tableContainer.querySelector('tr:last-child');
        if (lastRow) {
            lastRow.scrollIntoView({
                behavior: 'smooth',
                block: 'end'
            });
        }
        
        showNotification('Scroll ke bawah tabel...', 'info');
    } else {
        // Jika tidak ada tabel, scroll ke bagian results
        const resultsSection = document.getElementById('results-section');
        if (resultsSection) {
            resultsSection.scrollIntoView({
                behavior: 'smooth',
                block: 'end'
            });
        }
    }
}

// Fungsi untuk scroll ke atas tabel
function scrollToTopOfTable() {
    const tableContainer = document.querySelector('.table-responsive');
    if (tableContainer) {
        tableContainer.scrollTo({
            top: 0,
            behavior: 'smooth'
        });
        
        // Atau scroll ke header tabel
        const firstRow = tableContainer.querySelector('tr:first-child');
        if (firstRow) {
            firstRow.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
        
        showNotification('Scroll ke atas tabel...', 'info');
    }
}

// Fungsi untuk menambahkan tombol scroll di dalam tabel
function addTableScrollButtons() {
    // Hapus tombol lama jika ada
    const oldButtons = document.querySelectorAll('.table-scroll-buttons');
    oldButtons.forEach(btn => btn.remove());
    
    // Cari container tabel
    const tableContainers = document.querySelectorAll('.table-responsive');
    
    tableContainers.forEach((container, index) => {
        // Buat tombol container
        const buttonContainer = document.createElement('div');
        buttonContainer.className = 'table-scroll-buttons';
        
        // Tombol ke bawah
        const bottomBtn = document.createElement('button');
        bottomBtn.className = 'table-scroll-btn';
        bottomBtn.innerHTML = '<i class="fas fa-arrow-down"></i> Ke Bawah';
        bottomBtn.onclick = () => {
            container.scrollTo({
                top: container.scrollHeight,
                behavior: 'smooth'
            });
        };
        
        // Tombol ke atas
        const topBtn = document.createElement('button');
        topBtn.className = 'table-scroll-btn';
        topBtn.innerHTML = '<i class="fas fa-arrow-up"></i> Ke Atas';
        topBtn.onclick = () => {
            container.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        };
        
        // Tambahkan tombol ke container
        buttonContainer.appendChild(bottomBtn);
        buttonContainer.appendChild(topBtn);
        
        // Tambahkan setelah container tabel
        container.parentNode.insertBefore(buttonContainer, container.nextSibling);
    });
}

// Fungsi untuk membuat floating action buttons
function createFloatingActionButtons() {
    // Hapus FAB lama jika ada
    const oldFab = document.getElementById('fab-container');
    if (oldFab) {
        oldFab.remove();
    }
    
    // Buat container FAB
    const fabContainer = document.createElement('div');
    fabContainer.id = 'fab-container';
    fabContainer.className = 'fab-container';
    
    // FAB untuk ke atas
    const fabTop = document.createElement('button');
    fabTop.className = 'fab fab-top';
    fabTop.innerHTML = '<i class="fas fa-arrow-up"></i>';
    fabTop.title = 'Ke Atas Tabel';
    
    // Label untuk FAB
    const topLabel = document.createElement('span');
    topLabel.className = 'fab-label';
    topLabel.textContent = 'Ke Atas Tabel';
    fabTop.appendChild(topLabel);
    
    fabTop.onclick = scrollToTopOfTable;
    
    // FAB untuk ke bawah
    const fabBottom = document.createElement('button');
    fabBottom.className = 'fab';
    fabBottom.innerHTML = '<i class="fas fa-arrow-down"></i>';
    fabBottom.title = 'Ke Bawah Tabel';
    
    const bottomLabel = document.createElement('span');
    bottomLabel.className = 'fab-label';
    bottomLabel.textContent = 'Ke Bawah Tabel';
    fabBottom.appendChild(bottomLabel);
    
    fabBottom.onclick = scrollToBottomOfTable;
    
    // Tambahkan FAB ke container
    fabContainer.appendChild(fabTop);
    fabContainer.appendChild(fabBottom);
    
    // Tambahkan ke body
    document.body.appendChild(fabContainer);
    
    // Sembunyikan FAB secara default
    fabContainer.style.display = 'none';
    
    // Tampilkan FAB hanya saat tab "Data Terproses" aktif
    const observer = new MutationObserver(() => {
        const processedTab = document.getElementById('processed-tab');
        if (processedTab && processedTab.classList.contains('active')) {
            fabContainer.style.display = 'flex';
        } else {
            fabContainer.style.display = 'none';
        }
    });
    
    // Observe perubahan pada tab content
    const tabContainer = document.querySelector('.tab-content');
    if (tabContainer) {
        observer.observe(tabContainer, {
            attributes: true,
            attributeFilter: ['class'],
            subtree: true
        });
    }
    
    return fabContainer;
}

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
            <td>${item.jamMasuk}</td>
            <td>${item.jamKeluar}</td>
            <td>${item.durasiFormatted}</td>
            <td>${item.jamNormalFormatted}</td>
            <td style="color: ${item.jamLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-weight: bold;">
                ${jamLemburBulat > 0 ? jamLemburBulat + ' jam' : item.jamLembur}
                ${isCapped ? 
                    '<br><small style="color: #e67e22; font-weight: bold;">(MAKSIMAL 7 JAM)</small>' : 
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
                    ${isCapped ? 
                        '<br><small style="color: #e67e22; font-weight: bold;">(LEMBUR DIBATASI MAKSIMAL 7 JAM)</small>' : 
                        ''}
                    ${item.jamLemburDesimal > 0 && !isCapped ? 
                        '<br><small style="color: #666;">(dibulatkan: ' + jamLemburBulat + ' jam)</small>' : 
                        ''}
                </span>
            </td>
        `;
        
        // Tambahkan class khusus untuk hari Jumat
        if (item.isFriday) {
            row.classList.add('friday-row');
        }
        
        // Tambahkan class khusus jika lembur dibatasi
        if (isCapped) {
            row.classList.add('overtime-capped-row');
        }
        
        tbody.appendChild(row);
    });
}
    
    // Tambahkan tombol scroll setelah tabel dimuat
    setTimeout(() => {
        addTableScrollButtons();
        createFloatingActionButtons();
        
        // Tampilkan tombol jika data banyak
        if (data.length > 10) {
            createScrollToBottomButton();
            
            // Tambahkan header dengan tombol cepat
            addQuickNavigationHeader();
        }
    }, 500);
}

// Fungsi untuk menambahkan header dengan tombol cepat
function addQuickNavigationHeader() {
    const cardHeader = document.querySelector('.data-preview-card .card-header');
    if (!cardHeader) return;
    
    // Cek apakah sudah ada tombol navigasi
    if (cardHeader.querySelector('.quick-navigation')) {
        return;
    }
    
    // Buat container tombol cepat
    const quickNav = document.createElement('div');
    quickNav.className = 'quick-navigation';
    quickNav.style.cssText = `
        display: flex;
        gap: 10px;
        margin-top: 10px;
    `;
    
    // Tombol ke bagian bawah tabel
    const toBottomBtn = document.createElement('button');
    toBottomBtn.className = 'btn btn-sm btn-primary';
    toBottomBtn.innerHTML = '<i class="fas fa-arrow-down"></i> Ke Bawah Tabel';
    toBottomBtn.style.cssText = `
        padding: 5px 10px;
        font-size: 12px;
        display: flex;
        align-items: center;
        gap: 5px;
    `;
    toBottomBtn.onclick = scrollToBottomOfTable;
    
    // Tombol ke bagian atas tabel
    const toTopBtn = document.createElement('button');
    toTopBtn.className = 'btn btn-sm btn-secondary';
    toTopBtn.innerHTML = '<i class="fas fa-arrow-up"></i> Ke Atas Tabel';
    toTopBtn.style.cssText = `
        padding: 5px 10px;
        font-size: 12px;
        display: flex;
        align-items: center;
        gap: 5px;
    `;
    toTopBtn.onclick = scrollToTopOfTable;
    
    // Tombol untuk mencari data tertentu
    const searchBtn = document.createElement('button');
    searchBtn.className = 'btn btn-sm btn-success';
    searchBtn.innerHTML = '<i class="fas fa-search"></i> Cari Karyawan';
    searchBtn.style.cssText = `
        padding: 5px 10px;
        font-size: 12px;
        display: flex;
        align-items: center;
        gap: 5px;
    `;
    searchBtn.onclick = showSearchModal;
    
    quickNav.appendChild(toBottomBtn);
    quickNav.appendChild(toTopBtn);
    quickNav.appendChild(searchBtn);
    
    // Tambahkan setelah header
    cardHeader.appendChild(quickNav);
}

// Fungsi untuk menampilkan modal pencarian
function showSearchModal() {
    if (processedData.length === 0) {
        showNotification('Tidak ada data untuk dicari.', 'warning');
        return;
    }
    
    // Dapatkan daftar karyawan unik
    const uniqueEmployees = [...new Set(processedData.map(item => item.nama))].sort();
    
    const modalHtml = `
        <div class="modal" id="search-modal">
            <div class="modal-content" style="max-width: 500px;">
                <div class="modal-header">
                    <h3><i class="fas fa-search"></i> Cari Data Karyawan</h3>
                    <button class="modal-close" id="close-search-modal">&times;</button>
                </div>
                <div class="modal-body">
                    <div class="search-container">
                        <div class="input-group" style="margin-bottom: 1rem;">
                            <input type="text" id="employee-search" placeholder="Cari nama karyawan..." 
                                   style="width: 100%; padding: 10px; border: 2px solid #ddd; border-radius: 4px;">
                        </div>
                        
                        <div class="employee-list" id="employee-list" 
                             style="max-height: 300px; overflow-y: auto; border: 1px solid #eee; border-radius: 4px; padding: 10px;">
                            ${uniqueEmployees.map(employee => `
                                <div class="employee-item" data-employee="${employee}" 
                                     style="padding: 8px; border-bottom: 1px solid #eee; cursor: pointer; transition: background 0.2s;">
                                    <strong>${employee}</strong>
                                    <small style="color: #666; float: right;">
                                        ${processedData.filter(item => item.nama === employee).length} entri
                                    </small>
                                </div>
                            `).join('')}
                        </div>
                        
                        <div style="margin-top: 1rem; padding: 10px; background: #f8f9fa; border-radius: 4px;">
                            <strong>Tips:</strong> Klik nama karyawan untuk scroll ke data mereka
                        </div>
                    </div>
                </div>
                <div class="modal-footer">
                    <button class="btn btn-secondary" id="cancel-search">Batal</button>
                </div>
            </div>
        </div>
    `;
    
    // Tambahkan modal ke body
    const existingModal = document.getElementById('search-modal');
    if (existingModal) {
        existingModal.remove();
    }
    
    document.body.insertAdjacentHTML('beforeend', modalHtml);
    const modal = document.getElementById('search-modal');
    modal.classList.add('active');
    
    // Event listeners
    document.getElementById('close-search-modal').addEventListener('click', () => {
        modal.remove();
    });
    
    document.getElementById('cancel-search').addEventListener('click', () => {
        modal.remove();
    });
    
    // Pencarian real-time
    const searchInput = document.getElementById('employee-search');
    const employeeList = document.getElementById('employee-list');
    
    searchInput.addEventListener('input', function() {
        const searchTerm = this.value.toLowerCase();
        const items = employeeList.querySelectorAll('.employee-item');
        
        items.forEach(item => {
            const employeeName = item.getAttribute('data-employee').toLowerCase();
            if (employeeName.includes(searchTerm)) {
                item.style.display = 'block';
            } else {
                item.style.display = 'none';
            }
        });
    });
    
    // Klik pada item karyawan
    employeeList.addEventListener('click', function(e) {
        const employeeItem = e.target.closest('.employee-item');
        if (employeeItem) {
            const employeeName = employeeItem.getAttribute('data-employee');
            scrollToEmployee(employeeName);
            modal.remove();
        }
    });
    
    // Fokus pada input pencarian
    setTimeout(() => {
        searchInput.focus();
    }, 100);
}

// Fungsi untuk scroll ke data karyawan tertentu
function scrollToEmployee(employeeName) {
    const tbody = document.getElementById('processed-table-body');
    if (!tbody) return;
    
    // Cari baris pertama untuk karyawan ini
    const rows = tbody.querySelectorAll('tr');
    let targetRow = null;
    
    for (let row of rows) {
        const nameCell = row.querySelector('td:nth-child(2) strong');
        if (nameCell && nameCell.textContent === employeeName) {
            targetRow = row;
            break;
        }
    }
    
    if (targetRow) {
        // Scroll ke baris target
        targetRow.scrollIntoView({
            behavior: 'smooth',
            block: 'center'
        });
        
        // Highlight baris
        targetRow.style.backgroundColor = '#fff3cd';
        targetRow.style.transition = 'background-color 0.5s';
        
        // Hapus highlight setelah 3 detik
        setTimeout(() => {
            targetRow.style.backgroundColor = '';
        }, 3000);
        
        showNotification(`Scroll ke data ${employeeName}...`, 'success');
    } else {
        showNotification(`Data untuk ${employeeName} tidak ditemukan`, 'warning');
    }
}

// Fungsi untuk menambahkan back to top button
function addBackToTopButton() {
    // Cek apakah button sudah ada
    if (document.getElementById('back-to-top')) {
        return;
    }
    
    // Buat button
    const button = document.createElement('button');
    button.id = 'back-to-top';
    button.className = 'back-to-top';
    button.innerHTML = '<i class="fas fa-arrow-up"></i>';
    button.title = 'Kembali ke atas';
    
    // Tambahkan ke body
    document.body.appendChild(button);
    
    // Event listener
    button.addEventListener('click', () => {
        window.scrollTo({
            top: 0,
            behavior: 'smooth'
        });
    });
    
    // Tampilkan/sembunyikan button berdasarkan scroll
    window.addEventListener('scroll', () => {
        if (window.scrollY > 300) {
            button.classList.add('show');
        } else {
            button.classList.remove('show');
        }
    });
}

// ============================
// MAIN APPLICATION FUNCTIONS - DENGAN UPDATE JUMAT
// ============================

// Initialize application
function initializeApp() {
    console.log('Initializing app...');
    
    // Debug: cek elemen
    const loadingScreen = document.getElementById('loading-screen');
    const mainContainer = document.getElementById('main-container');
    
    console.log('Loading screen:', loadingScreen);
    console.log('Main container:', mainContainer);
    
    if (!loadingScreen || !mainContainer) {
        console.error('Critical elements not found!');
        // Coba langsung tampilkan main container
        if (mainContainer) {
            mainContainer.style.opacity = '1';
            mainContainer.style.display = 'block';
        }
        return;
    }
    
    // 1. Set tanggal
    const currentDate = document.getElementById('current-date');
    if (currentDate) {
        const now = new Date();
        const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
        currentDate.textContent = now.toLocaleDateString('id-ID', options);
    }
    
    // 2. Setup event listeners
    try {
        setupEventListeners();
        console.log('Event listeners setup successfully');
    } catch (error) {
        console.error('Error setting up event listeners:', error);
    }
    
    // 3. Sembunyikan loading screen setelah 1.5 detik
    setTimeout(() => {
        console.log('Hiding loading screen...');
        loadingScreen.style.transition = 'opacity 0.5s ease';
        loadingScreen.style.opacity = '0';
        
        setTimeout(() => {
            loadingScreen.style.display = 'none';
            mainContainer.style.opacity = '1';
            mainContainer.style.transition = 'opacity 0.5s ease';
            mainContainer.style.display = 'block';
            console.log('App initialized successfully!');
            
            // Tambahkan class loaded untuk styling
            mainContainer.classList.add('loaded');
        }, 500);
    }, 1500);
    
    // 4. Setup lainnya
    addBackToTopButton();
    
    console.log('Initialization complete');
}
    
    // 5. Tambahkan info pembulatan ke sidebar
    const systemInfo = document.querySelector('.system-info');
    if (systemInfo) {
        const infoHtml = `
            <div style="margin-top: 1rem; padding: 0.75rem; background: #e3f2fd; border-radius: 8px; border-left: 4px solid #2196f3;">
                <h5 style="margin-bottom: 0.5rem; font-size: 0.9rem; color: #1565c0;">
                    <i class="fas fa-clock"></i> Aturan Baru Lembur (Senin-Kamis):
                </h5>
                <ul style="font-size: 0.8rem; color: #1565c0; padding-left: 1.2rem; margin-bottom: 0.5rem;">
                    <li><strong>Jam masuk efektif:</strong> 07:00</li>
                    <li><strong>Perhitungan dimulai dari:</strong> Jam 07:00</li>
                    <li><strong>Jam kerja normal:</strong> 8 jam</li>
                    <li><strong>Lembur jika:</strong> Kerja > 8 jam dari jam 07:00</li>
                    <li><strong>Minimal lembur:</strong> ≥ 10 menit</li>
                </ul>
                <h5 style="margin-bottom: 0.5rem; font-size: 0.9rem; color: #7b1fa2;">
                    <i class="fas fa-calendar-day"></i> Aturan Jumat:
                </h5>
                <ul style="font-size: 0.8rem; color: #7b1fa2; padding-left: 1.2rem; margin-bottom: 0;">
                    <li><strong>Jam masuk efektif:</strong> 07:00</li>
                    <li><strong>Jam pulang normal:</strong> 15:00</li>
                    <li><strong>Lembur jika:</strong> Pulang > 15:00</li>
                    <li><strong>Minimal lembur:</strong> ≥ 10 menit</li>
                </ul>
            </div>
            
            <div style="margin-top: 1rem; padding: 0.75rem; background: #fff3cd; border-radius: 8px; border-left: 4px solid #ffc107;">
                <h5 style="margin-bottom: 0.5rem; font-size: 0.9rem; color: #856404;">
                    <i class="fas fa-calculator"></i> Aturan Pembulatan Lembur:
                </h5>
                <ul style="font-size: 0.8rem; color: #856404; padding-left: 1.2rem; margin-bottom: 0;">
                    <li>Contoh: 7,18 jam → dibulatkan menjadi 7 jam</li>
                    <li>Contoh: 7,67 jam → dibulatkan menjadi 8 jam</li>
                    <li>Pembulatan standar (0.5 ke atas)</li>
                    <li>Gaji dihitung berdasarkan jam yang telah dibulatkan</li>
                </ul>
            </div>
        `;
        systemInfo.insertAdjacentHTML('beforeend', infoHtml);
    }
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

// Process data (HITUNG LEMBUR SAJA) - DENGAN ATURAN JUMAT DAN PEMBULATAN
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
            // Hitung lembur per hari dengan jam kerja yang diatur DAN ATURAN JUMAT
            processedData = calculateOvertimePerDay(originalData, currentWorkHours);
            
            displayResults(processedData);
            createCharts(processedData);
            
            resultsSection.style.display = 'block';
            resultsSection.scrollIntoView({ behavior: 'smooth' });
            
            showNotification(`Perhitungan lembur selesai! Jam lembur telah dibulatkan (${currentWorkHours} jam kerja, termasuk aturan Jumat)`, 'success');
            
        } catch (error) {
            console.error('Error processing data:', error);
            showNotification('Terjadi kesalahan saat menghitung lembur.', 'error');
        } finally {
            processBtn.innerHTML = '<i class="fas fa-calculator"></i> Hitung Lembur';
            processBtn.disabled = false;
        }
    }, 1500);
}

// Display results - DENGAN UPDATE JUMAT DAN PEMBULATAN
function displayResults(data) {
    updateMainStatistics(data);
    displayOriginalTable(originalData);
    displayProcessedTable(data);
    displaySummaries(data);
    
    // Inisialisasi tombol scroll
    setTimeout(() => {
        createScrollToBottomButton();
        createFloatingActionButtons();
    }, 1000);
}

// Update main statistics
function updateMainStatistics(data) {
    const totalKaryawan = new Set(data.map(item => item.nama)).size;
    const totalHari = data.length;
    const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
    const totalLemburDesimal = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    const totalLemburBulat = roundOvertimeHours(totalLemburDesimal);
    
    // Hitung statistik terpisah
    const separated = calculateOvertimeByDayType(data);
    const hariJumat = separated.friday.totalDays;
    const hariLain = separated.otherDays.totalDays;
    const lemburJumat = separated.friday.totalOvertime;
    const lemburLain = separated.otherDays.totalOvertime;
    const lemburJumatBulat = roundOvertimeHours(lemburJumat);
    const lemburLainBulat = roundOvertimeHours(lemburLain);
    
    const totalKaryawanElem = document.getElementById('total-karyawan');
    const totalHariElem = document.getElementById('total-hari');
    const totalLemburElem = document.getElementById('total-lembur');
    const totalGajiElem = document.getElementById('total-gaji');
    
    if (totalKaryawanElem) totalKaryawanElem.textContent = totalKaryawan;
    if (totalHariElem) totalHariElem.textContent = totalHari;
    if (totalLemburElem) totalLemburElem.textContent = formatHoursToDisplay(totalLemburDesimal) + ` (${totalLemburBulat} jam dibulatkan)`;
    
    // Update total gaji dengan info terpisah
    if (totalGajiElem) {
        totalGajiElem.innerHTML = `
            ${totalLemburBulat} jam lembur (dibulatkan)<br>
            <small style="font-size: 0.8rem;">
                Jumat: ${lemburJumatBulat} jam (${hariJumat} hari)<br>
                Senin-Kamis: ${lemburLainBulat} jam (${hariLain} hari)
            </small>
        `;
    }
}

// Tampilkan data original dengan info hari
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
        
        // Tambahkan class khusus untuk hari Jumat
        if (isFridayDay) {
            row.classList.add('friday-row');
        }
        
        tbody.appendChild(row);
    });
}

// Display summaries DENGAN ATURAN BARU
function displaySummaries(data) {
    // Hitung summary karyawan
    const employeeSummary = calculateOvertimeSummary(data);
    const employeeSummaryElem = document.getElementById('employee-summary');
    const financialSummary = document.getElementById('financial-summary');
    
    if (employeeSummaryElem) {
        let html = '<div style="max-height: 300px; overflow-y: auto;">';
        employeeSummary.forEach(emp => {
            if (emp.totalLemburDesimal > 0) {
                html += `
                    <div style="margin-bottom: 1rem; padding: 0.75rem; background: #f8f9fa; border-radius: 8px;">
                        <div style="display: flex; justify-content: space-between; align-items: center;">
                            <div>
                                <strong>${emp.nama}</strong>
                                <br><small style="color: #666;">Kategori: ${emp.kategori}</small>
                            </div>
                            <div style="text-align: right;">
                                <div style="color: #e74c3c; font-weight: bold;">${emp.totalLemburBulat} jam lembur</div>
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
    
    // Financial summary dengan perhitungan gaji DAN INFO ATURAN BARU
    const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
    const totalLemburDesimal = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    const totalLemburBulat = roundOvertimeHours(totalLemburDesimal);
    const totalNormal = data.reduce((sum, item) => sum + item.jamNormal, 0);
    const hariDenganLembur = data.filter(item => item.jamLemburDesimal > 0).length;
    
    // Hitung statistik tambahan
    const hariMasukSebelum7 = data.filter(item => {
        const jamMasuk = item.jamMasuk;
        if (!jamMasuk) return false;
        const [hours] = jamMasuk.split(':').map(Number);
        return hours < 7;
    }).length;
    
    const hariKerjaLebih8Jam = data.filter(item => {
        return item.durasi > 8;
    }).length;
    
    const hariLemburDiabaikan = data.filter(item => {
        // Lembur diabaikan jika < 10 menit setelah 8 jam kerja
        return item.durasi > 8 && item.jamLemburDesimal === 0;
    }).length;
    
    // Hitung total gaji dengan pembulatan
    let totalGajiAll = 0;
    let totalGajiJumat = 0;
    let totalGajiLain = 0;
    
    employeeSummary.forEach(emp => {
        totalGajiAll += emp.totalGaji;
        totalGajiJumat += emp.fridayGaji;
        totalGajiLain += emp.otherDaysGaji;
    });
    
    let salaryHtml = '';
    employeeSummary.forEach((emp, index) => {
        if (emp.totalLemburDesimal > 0) {
            salaryHtml += `
                <div style="display: flex; justify-content: space-between; margin-bottom: 0.5rem; padding: 0.5rem; background: ${index % 2 === 0 ? '#f8f9fa' : 'white'}; border-radius: 4px;">
                    <div>${emp.nama} (${emp.kategori})</div>
                    <div style="text-align: right;">
                        <div><strong>${emp.totalLemburBulat} jam</strong> × Rp ${emp.rate.toLocaleString('id-ID')}</div>
                        <div style="color: #27ae60; font-weight: bold;">${emp.totalGajiFormatted}</div>
                    </div>
                </div>
            `;
        }
    });
    
    if (financialSummary) {
        financialSummary.innerHTML = `
            <div><strong>ATURAN BARU PERHITUNGAN LEMBUR:</strong></div>
            <div style="font-size: 0.85rem; color: #666; margin-bottom: 1rem; padding: 0.75rem; background: #fff3cd; border-radius: 4px;">
                • <strong>Jam masuk efektif:</strong> 07:00 (semua hari)<br>
                • <strong>Perhitungan dimulai dari:</strong> Jam 07:00<br>
                • <strong>Jam kerja normal:</strong> 8 jam sehari<br>
                • <strong>Lembur dihitung jika:</strong> Kerja > 8 jam dari jam 07:00<br>
                • <strong>Minimal lembur:</strong> 10 menit<br>
                • <strong>Pembulatan jam:</strong> 0.5 ke atas → 1, di bawah 0.5 → 0
            </div>
            
            <div style="display: flex; justify-content: space-between; margin-bottom: 1rem;">
                <div style="flex: 1; padding: 0.5rem; background: rgba(52, 152, 219, 0.1); border-radius: 4px; margin-right: 0.5rem;">
                    <strong style="color: #3498db;">Statistik Kerja:</strong><br>
                    Masuk sebelum 07:00: <strong>${hariMasukSebelum7} hari</strong><br>
                    Kerja > 8 jam: <strong>${hariKerjaLebih8Jam} hari</strong><br>
                    Dengan lembur: <strong>${hariDenganLembur} hari</strong>
                </div>
                <div style="flex: 1; padding: 0.5rem; background: rgba(231, 76, 60, 0.1); border-radius: 4px;">
                    <strong style="color: #e74c3c;">Lembur Diabaikan:</strong><br>
                    Kerja > 8 jam tapi<br>lembur < 10 menit: <strong>${hariLemburDiabaikan} hari</strong>
                </div>
            </div>
            
            <div style="margin-top: 1rem;">
                <div>Konfigurasi Jam Kerja: <strong>${currentWorkHours} jam/hari</strong></div>
                <div>Total Entri Data: <strong>${data.length} hari</strong></div>
                <div>Total Jam Kerja: <strong>${formatHoursToDisplay(totalJam)}</strong></div>
                <div>Total Jam Normal: <strong>${formatHoursToDisplay(totalNormal)}</strong></div>
                <div style="color: ${totalLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-weight: bold;">
                    Total Jam Lembur: <strong>${formatHoursToDisplay(totalLemburDesimal)}</strong><br>
                    <small>Setelah dibulatkan: <strong>${roundAndFormatHours(totalLemburDesimal)}</strong></small>
                </div>
            </div>
            <div style="border-top: 2px solid #3498db; padding-top: 0.5rem; margin-top: 0.5rem;">
                <strong>Perhitungan Gaji Lembur (dengan pembulatan):</strong><br>
                ${salaryHtml}
                <div style="display: flex; justify-content: space-between; margin-top: 0.5rem;">
                    <div style="font-weight: bold; color: #9b59b6;">
                        TOTAL JUMAT: Rp ${Math.round(totalGajiJumat).toLocaleString('id-ID')}
                    </div>
                    <div style="font-weight: bold; color: #2c3e50;">
                        TOTAL GAJI LEMBUR: Rp ${Math.round(totalGajiAll).toLocaleString('id-ID')}
                    </div>
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
            employeeGroups[item.nama] = { normal: 0, lembur: 0, jumat: 0 };
        }
        employeeGroups[item.nama].normal += item.jamNormal;
        employeeGroups[item.nama].lembur += item.jamLemburDesimal;
        if (item.isFriday && item.jamLemburDesimal > 0) {
            employeeGroups[item.nama].jumat += item.jamLemburDesimal;
        }
    });
    
    const employeeNames = Object.keys(employeeGroups).slice(0, 10);
    const regularHours = employeeNames.map(name => employeeGroups[name].normal);
    const overtimeHours = employeeNames.map(name => employeeGroups[name].lembur);
    const fridayOvertime = employeeNames.map(name => employeeGroups[name].jumat);
    
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
                    label: 'Jam Lembur (Hari Biasa)',
                    data: overtimeHours.map((hour, idx) => hour - fridayOvertime[idx]),
                    backgroundColor: 'rgba(231, 76, 60, 0.7)',
                    borderColor: 'rgba(231, 76, 60, 1)',
                    borderWidth: 1
                },
                {
                    label: 'Jam Lembur (Jumat)',
                    data: fridayOvertime,
                    backgroundColor: 'rgba(155, 89, 182, 0.7)',
                    borderColor: 'rgba(155, 89, 182, 1)',
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
                    text: 'Distribusi Jam Kerja per Karyawan (Desimal)'
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
    const totalLemburJumat = data.reduce((sum, item) => sum + (item.isFriday ? item.jamLemburDesimal : 0), 0);
    const totalLemburBiasa = totalLembur - totalLemburJumat;
    
    const salaryCtx = document.getElementById('salaryChart').getContext('2d');
    salaryChart = new Chart(salaryCtx, {
        type: 'doughnut',
        data: {
            labels: ['Jam Normal', 'Lembur Hari Biasa', 'Lembur Jumat'],
            datasets: [{
                data: [totalNormal, totalLemburBiasa, totalLemburJumat],
                backgroundColor: [
                    'rgba(52, 152, 219, 0.8)',
                    'rgba(231, 76, 60, 0.8)',
                    'rgba(155, 89, 182, 0.8)'
                ],
                borderColor: [
                    'rgba(52, 152, 219, 1)',
                    'rgba(231, 76, 60, 1)',
                    'rgba(155, 89, 182, 1)'
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
                    text: 'Komposisi Jam Kerja (Desimal)'
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
            
            // Tambahkan tombol scroll jika tab "processed" aktif
            if (tabId === 'processed') {
                setTimeout(() => {
                    addTableScrollButtons();
                }, 300);
            }
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
            'Tanggal': '01/11/2025', // Jumat
            'Jam Masuk': '06:30',
            'Jam Keluar': '15:30' // Lembur 30 menit
        },
        {
            'Nama': 'Windy',
            'Tanggal': '02/11/2025', // Sabtu
            'Jam Masuk': '08:30',
            'Jam Keluar': '17:30'
        },
        {
            'Nama': 'Bu Ati',
            'Tanggal': '01/11/2025', // Jumat
            'Jam Masuk': '07:00',
            'Jam Keluar': '16:00' // Lembur 1 jam
        }
    ];
    
    generateReport(templateData, 'template_data_presensi.xlsx', 'Template Presensi');
    showNotification('Template berhasil diunduh.', 'success');
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
            'Jam Lembur (Desimal)': item.jamLemburDesimal.toFixed(2) + ' jam',
            'Jam Lembur (Bulat)': roundOvertimeHours(item.jamLemburDesimal) + ' jam',
            'Keterangan': item.keterangan,
            'Catatan Khusus': item.isFriday ? 'HARI JUMAT - Pulang normal jam 15:00' : ''
        }));
    } else {
        return data.map((item, index) => ({
            'No': index + 1,
            'Nama': item.nama,
            'Tanggal': formatDate(item.tanggal),
            'Hari': getDayName(item.tanggal),
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
