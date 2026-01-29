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
// PERBAIKAN ATURAN PERHITUNGAN
// ============================

// Aturan Jumat (special rules)
const FRIDAY_RULES = {
    workStartTime: '07:00',      // Jam masuk efektif
    standardWorkHours: 8,         // Jam kerja normal 8 jam
    minOvertimeMinutes: 10        // Minimal lembur yang dihitung
};

// Aturan hari lain
const NORMAL_DAY_RULES = {
    workStartTime: '07:00',      // Jam masuk efektif
    standardWorkHours: 8,         // Jam kerja normal 8 jam
    minOvertimeMinutes: 10        // Minimal lembur yang dihitung
};

// Fungsi untuk menghitung jam kerja dengan aturan yang benar
function calculateWorkHoursWithRules(timeIn, timeOut, dateString) {
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
        
        // Konversi jam mulai kerja efektif ke menit
        const [startHour, startMinute] = rules.workStartTime.split(':').map(Number);
        const workStartMinutes = startHour * 60 + startMinute;
        
        // Konversi input ke menit
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        const outMinutes = outTime.hours * 60 + outTime.minutes;
        
        // ATURAN PERBAIKAN: 
        // 1. Jika datang sebelum jam 7, tetap dihitung mulai jam 7
        // 2. Jika datang setelah jam 7, gunakan jam datang sebenarnya
        let effectiveInMinutes = inMinutes;
        if (inMinutes < workStartMinutes) {
            effectiveInMinutes = workStartMinutes; // Datang sebelum 7, hitung dari 7
        }
        
        // Hitung total menit kerja
        let totalMinutes = outMinutes - effectiveInMinutes;
        
        // Jika pulang sebelum masuk (melewati tengah malam)
        if (totalMinutes < 0) {
            totalMinutes += 24 * 60;
        }
        
        return Math.round((totalMinutes / 60) * 100) / 100;
        
    } catch (error) {
        console.error('Error calculating work hours:', error);
        return 0;
    }
}

// FUNGSI BARU: Hitung lembur berdasarkan jam kerja 8 jam, bukan jam pulang tetap
function calculateOvertimeBasedOnWorkHours(timeIn, timeOut, dateString, workHours = 8) {
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
        
        // Konversi jam mulai kerja efektif ke menit
        const [startHour, startMinute] = rules.workStartTime.split(':').map(Number);
        const workStartMinutes = startHour * 60 + startMinute;
        
        // Konversi input ke menit
        const inMinutes = inTime.hours * 60 + inTime.minutes;
        const outMinutes = outTime.hours * 60 + outTime.minutes;
        
        // ATURAN PERBAIKAN: 
        // 1. Jika datang sebelum jam 7, tetap dihitung mulai jam 7
        // 2. Jika datang setelah jam 7, gunakan jam datang sebenarnya
        let effectiveInMinutes = inMinutes;
        if (inMinutes < workStartMinutes) {
            effectiveInMinutes = workStartMinutes; // Datang sebelum 7, hitung dari 7
        }
        
        // Hitung total menit kerja
        let totalMinutes = outMinutes - effectiveInMinutes;
        
        // Jika pulang sebelum masuk (melewati tengah malam)
        if (totalMinutes < 0) {
            totalMinutes += 24 * 60;
        }
        
        // Konversi jam kerja normal ke menit
        const standardWorkMinutes = workHours * 60;
        
        // Hitung lembur: total menit kerja dikurangi jam kerja normal
        let overtimeMinutes = totalMinutes - standardWorkMinutes;
        
        // Jika tidak ada lembur (kerja kurang dari 8 jam)
        if (overtimeMinutes <= 0) {
            return 0;
        }
        
        // Aturan: Abaikan lembur kurang dari minimal yang ditentukan
        if (overtimeMinutes < rules.minOvertimeMinutes) {
            return 0;
        }
        
        // Konversi ke jam (desimal)
        const overtimeHours = Math.round((overtimeMinutes / 60) * 100) / 100;
        
        return overtimeHours;
        
    } catch (error) {
        console.error('Error calculating overtime based on work hours:', error);
        return 0;
    }
}

// Update fungsi calculateOvertimePerDay dengan aturan yang benar
function calculateOvertimePerDay(data, workHours = 8) {
    const result = data.map(record => {
        // Hitung jam kerja dengan aturan yang benar
        const hoursWorked = calculateWorkHoursWithRules(
            record.jamMasuk, 
            record.jamKeluar, 
            record.tanggal
        );
        
        // PERBAIKAN: Hitung lembur berdasarkan jam kerja 8 jam
        const jamLemburDesimal = calculateOvertimeBasedOnWorkHours(
            record.jamMasuk, 
            record.jamKeluar, 
            record.tanggal,
            workHours
        );
        
        // Jam normal: 8 jam atau jam kerja jika kurang dari 8 jam
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
        
        // Tentukan jam masuk efektif berdasarkan aturan
        const effectiveInTime = getEffectiveInTimeWithDayRules(record.jamMasuk, record.tanggal);
        
        // Cek apakah hari Jumat
        const isFridayDay = isFriday(record.tanggal);
        
        // Keterangan dengan aturan baru
        let keterangan = 'Tidak lembur';
        let dayInfo = isFridayDay ? ' (JUMAT)' : '';
        
        if (jamLemburDesimal > 0) {
            // Hitung jam selesai normal (jam masuk efektif + 8 jam)
            const [startHour, startMinute] = effectiveInTime.split(':').map(Number);
            const endNormalMinutes = startHour * 60 + startMinute + (workHours * 60);
            const endNormalHour = Math.floor(endNormalMinutes / 60) % 24;
            const endNormalMinute = endNormalMinutes % 60;
            const endNormalTime = `${endNormalHour.toString().padStart(2, '0')}:${endNormalMinute.toString().padStart(2, '0')}`;
            
            keterangan = `Lembur ${jamLemburDisplay} (kerja > 8 jam)`;
            keterangan += `${dayInfo}`;
            
            // Tambahkan info waktu normal selesai
            keterangan += `<br><small>Selesai normal: ${endNormalTime}</small>`;
            
            // Jika datang sebelum 7, tambahkan catatan
            if (effectiveInTime !== record.jamMasuk) {
                keterangan += `<br><small>Datang ${record.jamMasuk}, hitung dari ${effectiveInTime}</small>`;
            }
        } else {
            // Tidak lembur
            if (effectiveInTime !== record.jamMasuk) {
                keterangan = `Datang ${record.jamMasuk}, hitung dari ${effectiveInTime}${dayInfo}`;
            } else if (isFridayDay) {
                keterangan = `Hari Jumat - Kerja 8 jam${dayInfo}`;
            } else {
                keterangan = `Kerja ≤ 8 jam${dayInfo}`;
            }
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
            // Tambahan untuk info perhitungan
            jamSelesaiNormal: calculateNormalEndTime(effectiveInTime, workHours),
            datangSebelum7: effectiveInTime !== record.jamMasuk
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

// Fungsi tambahan untuk menghitung jam selesai normal
function calculateNormalEndTime(startTime, workHours = 8) {
    if (!startTime) return '';
    
    try {
        const [hours, minutes] = startTime.split(':').map(Number);
        const totalMinutes = hours * 60 + minutes + (workHours * 60);
        const endHour = Math.floor(totalMinutes / 60) % 24;
        const endMinute = totalMinutes % 60;
        
        return `${endHour.toString().padStart(2, '0')}:${endMinute.toString().padStart(2, '0')}`;
    } catch (error) {
        return '';
    }
}

// Update fungsi untuk mendapatkan jam masuk efektif
function getEffectiveInTimeWithDayRules(jamMasuk, dateString) {
    if (!jamMasuk) return '00:00';
    
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
        
        // PERBAIKAN: 
        // 1. Jika datang sebelum jam 7, hitung mulai jam 7
        // 2. Jika datang setelah jam 7, gunakan jam datang sebenarnya
        if (inMinutes < workStartMinutes) {
            return rules.workStartTime; // Datang sebelum 7, hitung dari 7
        } else {
            return jamMasuk; // Datang setelah 7, gunakan jam datang
        }
        
    } catch (error) {
        console.error('Error getting effective in time:', error);
        return jamMasuk;
    }
}

// Update displayProcessedTable untuk menampilkan info baru
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
        
        // Warna baris berdasarkan kondisi
        let rowClass = '';
        let rowStyle = '';
        
        if (item.jamLemburDesimal > 0) {
            rowClass = 'lembur-row';
            rowStyle = 'background-color: rgba(231, 76, 60, 0.05);';
        } else if (item.datangSebelum7) {
            rowClass = 'datang-awal-row';
            rowStyle = 'background-color: rgba(52, 152, 219, 0.05);';
        }
        
        if (item.isFriday) {
            rowClass += ' friday-row';
        }
        
        row.className = rowClass;
        row.setAttribute('style', rowStyle);
        
        row.innerHTML = `
            <td>${index + 1}</td>
            <td><strong>${item.nama}</strong></td>
            <td>
                ${formatDate(item.tanggal)}
                ${item.isFriday ? '<br><small style="color: #9b59b6; font-weight: bold;">(JUMAT)</small>' : ''}
                <br><small style="color: #666;">${item.hari}</small>
            </td>
            <td>
                ${item.jamMasuk}
                ${item.datangSebelum7 ? 
                    `<br><small style="color: #3498db;">(efektif: ${item.jamMasukEfektif})</small>` : 
                    ''}
            </td>
            <td>${item.jamKeluar}</td>
            <td>${item.durasiFormatted}</td>
            <td>${item.jamNormalFormatted}</td>
            <td style="color: ${item.jamLemburDesimal > 0 ? '#e74c3c' : '#27ae60'}; font-weight: bold;">
                ${jamLemburBulat > 0 ? jamLemburBulat + ' jam' : item.jamLembur}
                ${item.jamLemburDesimal > 0 ? 
                    `<br><small style="color: #666; font-size: 0.8rem;">
                        (${item.jamLemburDesimal.toFixed(2)} jam desimal)
                    </small>` : 
                    ''}
            </td>
            <td>
                <div style="font-size: 0.85rem;">
                    ${item.keterangan.replace(/<br>/g, '<br>')}
                    ${item.jamLemburDesimal > 0 ? 
                        `<br><small style="color: #666;">(dibulatkan: ${jamLemburBulat} jam)</small>` : 
                        ''}
                </div>
            </td>
        `;
        
        tbody.appendChild(row);
    });
    
    // Tambahkan CSS untuk row styling
    if (!document.querySelector('#row-styles')) {
        const style = document.createElement('style');
        style.id = 'row-styles';
        style.textContent = `
            .lembur-row:hover {
                background-color: rgba(231, 76, 60, 0.1) !important;
            }
            .datang-awal-row:hover {
                background-color: rgba(52, 152, 219, 0.1) !important;
            }
            .friday-row:hover {
                background-color: rgba(155, 89, 182, 0.1) !important;
            }
        `;
        document.head.appendChild(style);
    }
    
    // Tambahkan tombol scroll setelah tabel dimuat
    setTimeout(() => {
        addTableScrollButtons();
        createFloatingActionButtons();
        
        if (data.length > 10) {
            createScrollToBottomButton();
            addQuickNavigationHeader();
        }
    }, 500);
}

// Update sidebar info untuk menjelaskan aturan baru
function initializeApp() {
    // ... kode sebelumnya ...
    
    // Tambahkan info aturan baru ke sidebar
    const systemInfo = document.querySelector('.system-info');
    if (systemInfo) {
        const infoHtml = `
            <div style="margin-top: 1rem; padding: 0.75rem; background: #d4edda; border-radius: 8px; border-left: 4px solid #28a745;">
                <h5 style="margin-bottom: 0.5rem; font-size: 0.9rem; color: #155724;">
                    <i class="fas fa-clock"></i> Aturan Jam Kerja Baru:
                </h5>
                <ul style="font-size: 0.8rem; color: #155724; padding-left: 1.2rem; margin-bottom: 0;">
                    <li>Jam masuk efektif: 07:00 (semua hari)</li>
                    <li>Datang < 07:00 → hitung mulai 07:00</li>
                    <li>Datang > 07:00 → gunakan jam datang sebenarnya</li>
                    <li>Jam kerja normal: 8 jam mulai dari jam masuk efektif</li>
                    <li>Lembur: kerja > 8 jam dari jam masuk efektif</li>
                    <li>Pembulatan lembur: 0.5 ke atas → 1, di bawah 0.5 → 0</li>
                </ul>
            </div>
        `;
        
        // Cari dan hapus info lama jika ada
        const oldInfo = systemInfo.querySelector('div:nth-last-child(1)');
        if (oldInfo && oldInfo.querySelector('h5') && oldInfo.querySelector('h5').textContent.includes('Aturan Pembulatan Lembur')) {
            oldInfo.remove();
        }
        
        systemInfo.insertAdjacentHTML('beforeend', infoHtml);
    }
    
    // Event listener untuk tombol download gaji lembur
    const downloadSalaryBtn = document.getElementById('download-salary');
    if (downloadSalaryBtn) {
        // Ubah event listener untuk menampilkan modal download terpisah
        downloadSalaryBtn.addEventListener('click', () => {
            if (processedData.length === 0) {
                showNotification('Data belum diproses. Silakan hitung lembur terlebih dahulu.', 'warning');
                return;
            }
            // Tampilkan modal pilihan download terpisah
            showSeparatedDownloadOptions();
        });
    }
    
    // Tab functionality dengan scroll buttons
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            const tabId = this.getAttribute('data-tab');
            switchTab(tabId);
            
            // Tampilkan/sembunyikan tombol scroll berdasarkan tab aktif
            setTimeout(() => {
                const fabContainer = document.getElementById('fab-container');
                if (fabContainer) {
                    if (tabId === 'processed') {
                        fabContainer.style.display = 'flex';
                    } else {
                        fabContainer.style.display = 'none';
                    }
                }
            }, 300);
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
    
    // Tambahkan back to top button
    addBackToTopButton();
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

// Display summaries DENGAN ATURAN JUMAT DAN PEMBULATAN
function displaySummaries(data) {
    const employeeSummary = document.getElementById('employee-summary');
    const financialSummary = document.getElementById('financial-summary');
    
    if (!employeeSummary && !financialSummary) return;
    
    // Pisahkan data
    const separated = calculateOvertimeByDayType(data);
    const summary = calculateOvertimeSummary(data);
    
    let employeeHtml = '';
    summary.forEach(item => {
        if (item.totalLemburDesimal > 0) {
            employeeHtml += `
                <div style="margin-bottom: 1rem; padding-bottom: 1rem; border-bottom: 1px solid #eee;">
                    <strong>${item.nama} (${item.kategori})</strong><br>
                    <small>
                        Total Lembur: ${item.totalLemburFormatted}<br>
                        <span style="color: #9b59b6;">Jumat: ${item.fridayLemburFormatted}</span> | 
                        <span style="color: #3498db;">Senin-Kamis: ${item.otherDaysLemburFormatted}</span><br>
                        Total Gaji: ${item.totalGajiFormatted}
                    </small>
                </div>
            `;
        }
    });
    
    if (employeeSummary) employeeSummary.innerHTML = employeeHtml;
    
    // Financial summary dengan perhitungan gaji DAN INFO JUMAT DAN PEMBULATAN
    const totalJam = data.reduce((sum, item) => sum + item.durasi, 0);
    const totalLemburDesimal = data.reduce((sum, item) => sum + item.jamLemburDesimal, 0);
    const totalLemburBulat = roundOvertimeHours(totalLemburDesimal);
    const totalNormal = data.reduce((sum, item) => sum + item.jamNormal, 0);
    const hariDenganLembur = data.filter(item => item.jamLemburDesimal > 0).length;
    
    // Hitung statistik tambahan khusus Jumat
    const hariJumat = separated.friday.totalDays;
    const hariJumatDenganLembur = separated.friday.data.filter(item => item.jamLemburDesimal > 0).length;
    const hariJumatPulangSebelum3 = separated.friday.data.filter(item => {
        if (!item.jamKeluar) return false;
        const [hours, minutes] = item.jamKeluar.split(':').map(Number);
        const outMinutes = hours * 60 + (minutes || 0);
        return outMinutes <= 15 * 60; // 15:00 = 900 menit
    }).length;
    
    // Hitung statistik hari lain
    const hariMasukSebelum7 = data.filter(item => {
        const jamMasuk = item.jamMasuk;
        if (!jamMasuk) return false;
        const [hours] = jamMasuk.split(':').map(Number);
        return hours < 7;
    }).length;
    
    // Hitung hari pulang setelah jam normal (berdasarkan hari)
    const hariPulangSetelahNormal = data.filter(item => {
        const jamKeluar = item.jamKeluar;
        if (!jamKeluar) return false;
        
        const [hours, minutes] = jamKeluar.split(':').map(Number);
        const outMinutes = hours * 60 + (minutes || 0);
        
        // Tentukan jam pulang normal berdasarkan hari
        const workEndMinutes = item.isFriday ? 15 * 60 : 16 * 60;
        
        return outMinutes > workEndMinutes;
    }).length;
    
    const hariLemburDiabaikan = data.filter(item => {
        const jamKeluar = item.jamKeluar;
        if (!jamKeluar) return false;
        
        const [hours, minutes] = jamKeluar.split(':').map(Number);
        const outMinutes = hours * 60 + (minutes || 0);
        
        // Tentukan jam pulang normal berdasarkan hari
        const workEndMinutes = item.isFriday ? 15 * 60 : 16 * 60;
        
        // Pulang setelah jam normal tapi kurang dari 10 menit
        return outMinutes > workEndMinutes && 
               (outMinutes - workEndMinutes) < 10;
    }).length;
    
    // Hitung gaji berdasarkan kategori DENGAN PEMBULATAN
    let salaryHtml = '';
    
    const byCategory = {
        'TU': { totalJam: 0, totalGaji: 0, fridayJam: 0, fridayGaji: 0 },
        'STAFF': { totalJam: 0, totalGaji: 0, fridayJam: 0, fridayGaji: 0 },
        'K3': { totalJam: 0, totalGaji: 0, fridayJam: 0, fridayGaji: 0 }
    };
    
    summary.forEach(item => {
        const category = item.kategori;
        if (byCategory[category]) {
            byCategory[category].totalJam += item.totalLemburDesimal;
            byCategory[category].totalGaji += item.totalGaji;
            byCategory[category].fridayJam += item.fridayOvertime;
            byCategory[category].fridayGaji += item.fridayGaji;
        }
    });
    
    let totalGajiAll = 0;
    let totalGajiJumat = 0;
    
    Object.keys(byCategory).forEach(category => {
        if (byCategory[category].totalJam > 0) {
            const otherDaysJam = byCategory[category].totalJam - byCategory[category].fridayJam;
            const otherDaysGaji = byCategory[category].totalGaji - byCategory[category].fridayGaji;
            
            // Versi dibulatkan
            const totalJamBulat = roundOvertimeHours(byCategory[category].totalJam);
            const fridayJamBulat = roundOvertimeHours(byCategory[category].fridayJam);
            const otherDaysJamBulat = roundOvertimeHours(otherDaysJam);
            
            salaryHtml += `
                <div style="margin: 0.5rem 0; padding: 0.5rem; background: #f8f9fa; border-radius: 4px;">
                    <strong>${category}:</strong><br>
                    • Jumat: ${fridayJamBulat} jam (${byCategory[category].fridayJam.toFixed(2)}) = Rp ${Math.round(fridayJamBulat * overtimeRates[category]).toLocaleString('id-ID')}<br>
                    • Senin-Kamis: ${otherDaysJamBulat} jam (${otherDaysJam.toFixed(2)}) = Rp ${Math.round(otherDaysJamBulat * overtimeRates[category]).toLocaleString('id-ID')}<br>
                    <strong>Total: ${totalJamBulat} jam = Rp ${Math.round(totalJamBulat * overtimeRates[category]).toLocaleString('id-ID')}</strong>
                </div>
            `;
            
            // Gunakan versi dibulatkan untuk perhitungan total
            totalGajiAll += totalJamBulat * overtimeRates[category];
            totalGajiJumat += fridayJamBulat * overtimeRates[category];
        }
    });
    
    if (financialSummary) {
        financialSummary.innerHTML = `
            <div><strong>Aturan Perhitungan:</strong></div>
            <div style="font-size: 0.85rem; color: #666; margin-bottom: 1rem;">
                • Jam masuk efektif: 07:00 (semua hari)<br>
                • Jam pulang normal: <br>
                &nbsp;&nbsp;&nbsp;- Senin-Kamis: 16:00<br>
                &nbsp;&nbsp;&nbsp;- Jumat: 15:00<br>
                • Minimal lembur: 10 menit<br>
                • <strong style="color: #e74c3c;">Pembulatan: 0.5 ke atas → 1, di bawah 0.5 → 0</strong>
            </div>
            
            <div style="display: flex; justify-content: space-between;">
                <div style="flex: 1; padding: 0.5rem; background: rgba(155, 89, 182, 0.1); border-radius: 4px; margin-right: 0.5rem;">
                    <strong style="color: #9b59b6;">Hari Jumat:</strong><br>
                    Total Hari: ${hariJumat}<br>
                    Dengan Lembur: ${hariJumatDenganLembur}<br>
                    Pulang ≤ 15:00: ${hariJumatPulangSebelum3}
                </div>
                <div style="flex: 1; padding: 0.5rem; background: rgba(52, 152, 219, 0.1); border-radius: 4px;">
                    <strong style="color: #3498db;">Senin-Kamis:</strong><br>
                    Total Hari: ${separated.otherDays.totalDays}<br>
                    Dengan Lembur: ${separated.otherDays.data.filter(item => item.jamLemburDesimal > 0).length}<br>
                    Karyawan: ${separated.otherDays.totalEmployees}
                </div>
            </div>
            
            <div style="margin-top: 1rem;">
                <div>Konfigurasi Jam Kerja: <strong>${currentWorkHours} jam/hari</strong></div>
                <div>Total Entri Data: <strong>${data.length} hari</strong></div>
                <div>Masuk sebelum 07:00: <strong>${hariMasukSebelum7} hari</strong></div>
                <div>Pulang setelah jam normal: <strong>${hariPulangSetelahNormal} hari</strong></div>
                <div>Lembur diabaikan (<10 menit): <strong>${hariLemburDiabaikan} hari</strong></div>
                <div>Hari dengan Lembur: <strong>${hariDenganLembur} hari</strong></div>
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
