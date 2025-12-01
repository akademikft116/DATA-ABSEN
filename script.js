// Sistem Kalkulator Absensi & Lembur Excel
class AttendanceOvertimeCalculator {
    constructor() {
        this.rawData = [];
        this.attendanceData = [];
        this.overtimeData = [];
        this.employeeSummary = {};
        this.fileName = '';
        this.fileType = '';
        
        // Settings
        this.startTime = '08:00';
        this.endTime = '17:00';
        this.workingHours = 8;
        this.overtimeRate = 12500;
        this.roundingType = 'down';
        this.tolerance = 15;
        this.period = 'OKTOBER 2025';
        this.mealAllowance = 20000;
        this.transportAllowance = 15000;
        
        // Charts
        this.overtimeChart = null;
        this.employeeChart = null;
        
        this.init();
    }

    init() {
        this.setupEventListeners();
        this.loadSettings();
        this.updateUI();
        this.logActivity('Sistem siap. Upload file Excel untuk memulai.', 'info');
    }

    setupEventListeners() {
        // Tab navigation
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchTab(e.target.dataset.tab));
        });
        
        document.querySelectorAll('.results-tab-btn').forEach(btn => {
            btn.addEventListener('click', (e) => this.switchResultsTab(e.target.dataset.tab));
        });

        // Upload file
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('excelFile');
        
        uploadArea.addEventListener('click', () => fileInput.click());
        
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#3498db';
            uploadArea.style.background = '#e8f4fc';
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.style.borderColor = '#dee2e6';
            uploadArea.style.background = '#f8f9fa';
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#dee2e6';
            uploadArea.style.background = '#f8f9fa';
            
            const file = e.dataTransfer.files[0];
            if (file) {
                this.handleFileSelect(file);
            }
        });
        
        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                this.handleFileSelect(file);
            }
        });

        // Remove file
        document.getElementById('removeFile').addEventListener('click', () => this.removeFile());

        // Settings
        document.getElementById('startTime').addEventListener('change', (e) => {
            this.startTime = e.target.value;
            this.saveSettings();
        });
        
        document.getElementById('endTime').addEventListener('change', (e) => {
            this.endTime = e.target.value;
            this.saveSettings();
        });
        
        document.getElementById('workingHours').addEventListener('input', (e) => {
            this.workingHours = parseFloat(e.target.value) || 8;
            this.saveSettings();
        });
        
        document.getElementById('overtimeRate').addEventListener('input', (e) => {
            this.overtimeRate = parseInt(e.target.value) || 12500;
            this.saveSettings();
        });
        
        document.getElementById('roundingType').addEventListener('change', (e) => {
            this.roundingType = e.target.value;
            this.saveSettings();
        });
        
        document.getElementById('tolerance').addEventListener('input', (e) => {
            this.tolerance = parseInt(e.target.value) || 15;
            this.saveSettings();
        });
        
        document.getElementById('period').addEventListener('input', (e) => {
            this.period = e.target.value || 'OKTOBER 2025';
            this.saveSettings();
        });
        
        document.getElementById('mealAllowance').addEventListener('input', (e) => {
            this.mealAllowance = parseInt(e.target.value) || 20000;
            this.saveSettings();
        });
        
        document.getElementById('transportAllowance').addEventListener('input', (e) => {
            this.transportAllowance = parseInt(e.target.value) || 15000;
            this.saveSettings();
        });

        // Buttons
        document.getElementById('analyzeBtn').addEventListener('click', () => this.analyzeData());
        document.getElementById('calculateBtn').addEventListener('click', () => this.calculateOvertime());
        document.getElementById('exportBtn').addEventListener('click', () => this.exportToExcel());
        document.getElementById('resetBtn').addEventListener('click', () => this.resetAll());

        // Search
        document.getElementById('searchOvertime').addEventListener('input', (e) => {
            this.filterOvertimeData(e.target.value);
        });
    }

    loadSettings() {
        const saved = localStorage.getItem('attendanceCalculatorSettings');
        if (saved) {
            try {
                const settings = JSON.parse(saved);
                Object.assign(this, settings);
                this.updateUI();
            } catch (e) {
                console.log('Gagal load settings:', e);
            }
        }
    }

    saveSettings() {
        const settings = {
            startTime: this.startTime,
            endTime: this.endTime,
            workingHours: this.workingHours,
            overtimeRate: this.overtimeRate,
            roundingType: this.roundingType,
            tolerance: this.tolerance,
            period: this.period,
            mealAllowance: this.mealAllowance,
            transportAllowance: this.transportAllowance
        };
        localStorage.setItem('attendanceCalculatorSettings', JSON.stringify(settings));
    }

    updateUI() {
        document.getElementById('startTime').value = this.startTime;
        document.getElementById('endTime').value = this.endTime;
        document.getElementById('workingHours').value = this.workingHours;
        document.getElementById('overtimeRate').value = this.overtimeRate;
        document.getElementById('roundingType').value = this.roundingType;
        document.getElementById('tolerance').value = this.tolerance;
        document.getElementById('period').value = this.period;
        document.getElementById('mealAllowance').value = this.mealAllowance;
        document.getElementById('transportAllowance').value = this.transportAllowance;
    }

    switchTab(tabId) {
        document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        
        document.querySelector(`[data-tab="${tabId}"]`).classList.add('active');
        document.getElementById(tabId).classList.add('active');
        
        this.logActivity(`Beralih ke format: ${tabId}`, 'info');
    }

    switchResultsTab(tabId) {
        document.querySelectorAll('.results-tab-btn').forEach(btn => btn.classList.remove('active'));
        document.querySelectorAll('.results-tab-content').forEach(content => content.classList.remove('active'));
        
        document.querySelector(`[data-tab="${tabId}"]`).classList.add('active');
        document.getElementById(`${tabId}Tab`).classList.add('active');
    }

    handleFileSelect(file) {
        // Validasi file
        if (!file.name.match(/\.(xlsx|xls|csv)$/i)) {
            this.showError('File harus berformat Excel (.xlsx, .xls) atau CSV (.csv)');
            return;
        }

        if (file.size > 10 * 1024 * 1024) {
            this.showError('Ukuran file terlalu besar. Maksimal 10MB');
            return;
        }

        this.fileName = file.name;
        this.fileType = file.type;
        
        this.showProgress('Membaca file...', 10);
        
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                this.showProgress('Memproses data...', 30);
                
                let data;
                
                if (file.name.endsWith('.csv')) {
                    // Process CSV
                    data = this.parseCSV(e.target.result);
                } else {
                    // Process Excel
                    const workbook = XLSX.read(e.target.result, { type: 'binary' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    data = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                }
                
                this.rawData = data;
                
                this.showProgress('Validasi data...', 60);
                
                // Show file info
                this.showFileInfo(file);
                
                // Enable analyze button
                document.getElementById('analyzeBtn').disabled = false;
                document.getElementById('calculateBtn').disabled = true;
                document.getElementById('exportBtn').disabled = true;
                
                this.showProgress('Selesai!', 100);
                
                setTimeout(() => {
                    this.hideProgress();
                    this.logActivity(`File "${file.name}" berhasil diupload (${data.length} baris data)`, 'success');
                }, 500);
                
            } catch (error) {
                console.error('Error reading file:', error);
                this.showError('Gagal membaca file. Pastikan format file benar.');
                this.hideProgress();
            }
        };
        
        reader.onerror = () => {
            this.showError('Gagal membaca file');
            this.hideProgress();
        };
        
        if (file.name.endsWith('.csv')) {
            reader.readAsText(file);
        } else {
            reader.readAsBinaryString(file);
        }
    }

    parseCSV(csvText) {
        const rows = csvText.split('\n').map(row => row.split(',').map(cell => cell.trim()));
        return rows;
    }

    showFileInfo(file) {
        const fileInfo = document.getElementById('fileInfo');
        const fileName = document.getElementById('fileName');
        const fileSize = document.getElementById('fileSize');
        const fileStats = document.getElementById('fileStats');
        
        fileInfo.style.display = 'block';
        fileName.textContent = file.name;
        fileSize.textContent = this.formatFileSize(file.size);
        fileStats.textContent = `${this.rawData.length} baris data`;
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    removeFile() {
        this.rawData = [];
        this.attendanceData = [];
        this.overtimeData = [];
        this.employeeSummary = {};
        
        document.getElementById('fileInfo').style.display = 'none';
        document.getElementById('excelFile').value = '';
        document.getElementById('analyzeBtn').disabled = true;
        document.getElementById('calculateBtn').disabled = true;
        document.getElementById('exportBtn').disabled = true;
        
        this.logActivity('File telah dihapus', 'info');
    }

    analyzeData() {
        if (this.rawData.length === 0) {
            this.showError('Tidak ada data untuk dianalisis');
            return;
        }

        this.showProgress('Menganalisis data...', 20);
        
        try {
            this.attendanceData = [];
            this.employeeSummary = {};
            
            // Deteksi format data
            const format = this.detectFormat();
            
            this.showProgress('Memproses format data...', 40);
            
            if (format === 'attendance') {
                this.processAttendanceFormat();
            } else if (format === 'overtime') {
                this.processOvertimeFormat();
            } else if (format === 'raw') {
                this.processRawFormat();
            } else {
                // Coba otomatis
                this.processAutoFormat();
            }
            
            this.showProgress('Menyiapkan hasil...', 80);
            
            // Display attendance data
            this.displayAttendanceData();
            
            // Enable calculate button
            document.getElementById('calculateBtn').disabled = false;
            
            this.showProgress('Analisis selesai!', 100);
            
            setTimeout(() => {
                this.hideProgress();
                this.logActivity(`Data berhasil dianalisis: ${this.attendanceData.length} catatan absensi`, 'success');
                this.switchResultsTab('attendance');
            }, 500);
            
        } catch (error) {
            console.error('Error analyzing data:', error);
            this.showError('Terjadi kesalahan saat menganalisis data');
            this.hideProgress();
        }
    }

    detectFormat() {
        if (this.rawData.length < 2) return 'unknown';
        
        const firstRow = this.rawData[0];
        
        // Cek format absensi
        if (firstRow.includes('Tanggal') && firstRow.includes('Nama')) {
            return 'attendance';
        }
        
        // Cek format lembur
        if (firstRow.includes('Name') || firstRow.includes('Hari') || firstRow.includes('Tanggal')) {
            return 'overtime';
        }
        
        // Cek format mentah (nama dan waktu)
        if (firstRow.includes('Nama') && firstRow.includes('Waktu')) {
            return 'raw';
        }
        
        return 'auto';
    }

    processAttendanceFormat() {
        const headers = this.rawData[0];
        const tanggalIndex = headers.indexOf('Tanggal');
        const namaIndex = headers.indexOf('Nama');
        const masukIndex = headers.findIndex(h => h.includes('Masuk') || h.includes('IN'));
        const pulangIndex = headers.findIndex(h => h.includes('Pulang') || h.includes('OUT'));
        
        for (let i = 1; i < this.rawData.length; i++) {
            const row = this.rawData[i];
            if (!row || row.length < 4) continue;
            
            const tanggal = this.cleanValue(row[tanggalIndex]);
            const nama = this.cleanValue(row[namaIndex]);
            const jamMasuk = this.cleanValue(row[masukIndex]);
            const jamPulang = this.cleanValue(row[pulangIndex]);
            
            if (!tanggal || !nama || !jamMasuk) continue;
            
            this.attendanceData.push({
                tanggal,
                nama,
                jamMasuk: this.parseTime(jamMasuk),
                jamPulang: jamPulang ? this.parseTime(jamPulang) : null,
                status: this.getAttendanceStatus(jamMasuk)
            });
        }
    }

    processOvertimeFormat() {
        const headers = this.rawData[0];
        const nameIndex = headers.indexOf('Name') !== -1 ? headers.indexOf('Name') : 
                         headers.findIndex(h => h.toLowerCase().includes('nama'));
        const tanggalIndex = headers.indexOf('Tanggal') !== -1 ? headers.indexOf('Tanggal') : 
                           headers.findIndex(h => h.toLowerCase().includes('tanggal'));
        const inIndex = headers.indexOf('IN') !== -1 ? headers.indexOf('IN') : 
                       headers.findIndex(h => h.toLowerCase().includes('masuk') || h === 'IN');
        const outIndex = headers.indexOf('OUT') !== -1 ? headers.indexOf('OUT') : 
                        headers.findIndex(h => h.toLowerCase().includes('pulang') || h === 'OUT');
        
        for (let i = 1; i < this.rawData.length; i++) {
            const row = this.rawData[i];
            if (!row || row.length < 4) continue;
            
            const nama = this.cleanValue(row[nameIndex]);
            const tanggal = this.cleanValue(row[tanggalIndex]);
            const jamMasuk = this.cleanValue(row[inIndex]);
            const jamPulang = this.cleanValue(row[outIndex]);
            
            if (!nama || !tanggal || !jamMasuk) continue;
            
            this.attendanceData.push({
                tanggal,
                nama,
                jamMasuk: this.parseTime(jamMasuk),
                jamPulang: jamPulang ? this.parseTime(jamPulang) : null,
                status: 'Present'
            });
        }
    }

    processRawFormat() {
        const headers = this.rawData[0];
        const namaIndex = headers.indexOf('Nama');
        const waktuIndex = headers.indexOf('Waktu');
        const typeIndex = headers.indexOf('Type') !== -1 ? headers.indexOf('Type') : 
                         headers.findIndex(h => h.toLowerCase().includes('type') || h.toLowerCase().includes('tipe'));
        
        // Group by date and employee
        const grouped = {};
        
        for (let i = 1; i < this.rawData.length; i++) {
            const row = this.rawData[i];
            if (!row || row.length < 2) continue;
            
            const nama = this.cleanValue(row[namaIndex]);
            const waktu = this.cleanValue(row[waktuIndex]);
            const type = typeIndex !== -1 ? this.cleanValue(row[typeIndex]) : 'Unknown';
            
            if (!nama || !waktu) continue;
            
            // Parse datetime
            const datetime = this.parseDateTime(waktu);
            if (!datetime) continue;
            
            const dateStr = datetime.toISOString().split('T')[0];
            const timeStr = datetime.toTimeString().split(' ')[0].substring(0, 5);
            
            const key = `${dateStr}_${nama}`;
            
            if (!grouped[key]) {
                grouped[key] = {
                    nama,
                    tanggal: dateStr,
                    checkIn: null,
                    checkOut: null
                };
            }
            
            if (type.toLowerCase().includes('in') || timeStr < '12:00') {
                grouped[key].checkIn = timeStr;
            } else {
                grouped[key].checkOut = timeStr;
            }
        }
        
        // Convert grouped data to attendance data
        Object.values(grouped).forEach(item => {
            if (item.checkIn) {
                this.attendanceData.push({
                    tanggal: item.tanggal,
                    nama: item.nama,
                    jamMasuk: item.checkIn,
                    jamPulang: item.checkOut,
                    status: this.getAttendanceStatus(item.checkIn)
                });
            }
        });
    }

    processAutoFormat() {
        // Coba deteksi kolom otomatis
        const firstDataRow = this.rawData[1] || this.rawData[0];
        if (!firstDataRow) return;
        
        for (let i = 1; i < this.rawData.length; i++) {
            const row = this.rawData[i];
            if (!row || row.length < 3) continue;
            
            // Cari kolom yang berisi tanggal
            let tanggal = null;
            let nama = null;
            let waktu1 = null;
            let waktu2 = null;
            
            for (let j = 0; j < row.length; j++) {
                const cell = this.cleanValue(row[j]);
                if (!cell) continue;
                
                // Deteksi tanggal
                if (cell.match(/\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/) && !tanggal) {
                    tanggal = cell;
                }
                // Deteksi nama (bukan angka, bukan waktu)
                else if (!cell.match(/\d/) && !cell.match(/\d{1,2}[:\.]\d{2}/) && !nama) {
                    nama = cell;
                }
                // Deteksi waktu
                else if (cell.match(/\d{1,2}[:\.]\d{2}/)) {
                    if (!waktu1) {
                        waktu1 = cell;
                    } else if (!waktu2) {
                        waktu2 = cell;
                    }
                }
            }
            
            if (tanggal && nama && waktu1) {
                this.attendanceData.push({
                    tanggal,
                    nama,
                    jamMasuk: this.parseTime(waktu1),
                    jamPulang: waktu2 ? this.parseTime(waktu2) : null,
                    status: this.getAttendanceStatus(waktu1)
                });
            }
        }
    }

    cleanValue(value) {
        if (value === null || value === undefined || value === '') return '';
        return String(value).trim();
    }

    parseTime(timeStr) {
        if (!timeStr) return null;
        
        // Format: 7.15, 7:15, 07:15, 7.30, 16.02, 16:02
        let time = timeStr.toString();
        
        // Ganti titik dengan titik dua jika perlu
        if (time.includes('.')) {
            time = time.replace('.', ':');
        }
        
        // Tambahkan :00 jika hanya jam
        if (time.match(/^\d{1,2}$/)) {
            time += ':00';
        }
        
        // Format ke HH:MM
        const parts = time.split(':');
        if (parts.length >= 2) {
            const hours = parseInt(parts[0]).toString().padStart(2, '0');
            const minutes = parseInt(parts[1]).toString().padStart(2, '0');
            return `${hours}:${minutes}`;
        }
        
        return time;
    }

    parseDateTime(datetimeStr) {
        try {
            // Coba berbagai format
            let date = null;
            
            // Format: DD/MM/YYYY HH:MM
            if (datetimeStr.includes('/')) {
                const parts = datetimeStr.split(' ');
                if (parts.length === 2) {
                    const dateParts = parts[0].split('/');
                    const timeParts = parts[1].split(':');
                    
                    if (dateParts.length === 3 && timeParts.length >= 2) {
                        date = new Date(
                            parseInt(dateParts[2]),
                            parseInt(dateParts[1]) - 1,
                            parseInt(dateParts[0]),
                            parseInt(timeParts[0]),
                            parseInt(timeParts[1])
                        );
                    }
                }
            }
            // Format: YYYY-MM-DD HH:MM:SS
            else if (datetimeStr.includes('-')) {
                date = new Date(datetimeStr);
            }
            
            if (date && !isNaN(date.getTime())) {
                return date;
            }
        } catch (e) {
            console.log('Error parsing datetime:', e);
        }
        
        return null;
    }

    getAttendanceStatus(checkInTime) {
        if (!checkInTime) return 'Absent';
        
        const checkIn = this.timeToMinutes(checkInTime);
        const expectedStart = this.timeToMinutes(this.startTime);
        
        if (checkIn <= expectedStart + this.tolerance) {
            return 'Present';
        } else {
            return 'Late';
        }
    }

    timeToMinutes(timeStr) {
        if (!timeStr) return 0;
        const parts = timeStr.split(':');
        return parseInt(parts[0]) * 60 + parseInt(parts[1] || 0);
    }

    calculateOvertime() {
        if (this.attendanceData.length === 0) {
            this.showError('Tidak ada data absensi untuk dihitung');
            return;
        }

        this.showProgress('Menghitung lembur...', 20);
        
        try {
            this.overtimeData = [];
            this.employeeSummary = {};
            
            // Hitung lembur per catatan absensi
            for (const record of this.attendanceData) {
                if (!record.jamMasuk || !record.jamPulang) continue;
                
                const masuk = this.timeToMinutes(record.jamMasuk);
                const pulang = this.timeToMinutes(record.jamPulang);
                
                // Hitung durasi kerja
                let durasi = pulang - masuk;
                if (durasi < 0) durasi += 24 * 60; // Jika melewati tengah malam
                
                // Konversi ke jam
                const durasiJam = durasi / 60;
                
                // Hitung lembur
                const lemburJam = Math.max(durasiJam - this.workingHours, 0);
                
                if (lemburJam > 0) {
                    // Tentukan hari
                    const hari = this.getDayNameFromDate(record.tanggal);
                    
                    // Hitung pembayaran
                    const jamDibulatkan = this.roundHours(lemburJam);
                    const gajiLembur = jamDibulatkan * this.overtimeRate;
                    
                    this.overtimeData.push({
                        nama: record.nama,
                        hari: hari,
                        tanggal: this.formatDateForExcel(record.tanggal),
                        masuk: record.jamMasuk,
                        pulang: record.jamPulang,
                        jamKerja: this.workingHours,
                        lembur: lemburJam,
                        jamDibulatkan: jamDibulatkan,
                        gajiLembur: gajiLembur,
                        status: 'Lembur'
                    });
                    
                    // Update summary per karyawan
                    if (!this.employeeSummary[record.nama]) {
                        this.employeeSummary[record.nama] = {
                            hadir: 0,
                            terlambat: 0,
                            totalLembur: 0,
                            totalGaji: 0
                        };
                    }
                    
                    this.employeeSummary[record.nama].totalLembur += lemburJam;
                    this.employeeSummary[record.nama].totalGaji += gajiLembur;
                }
                
                // Update kehadiran
                if (!this.employeeSummary[record.nama]) {
                    this.employeeSummary[record.nama] = {
                        hadir: 0,
                        terlambat: 0,
                        totalLembur: 0,
                        totalGaji: 0
                    };
                }
                
                if (record.status === 'Present') {
                    this.employeeSummary[record.nama].hadir++;
                } else if (record.status === 'Late') {
                    this.employeeSummary[record.nama].terlambat++;
                    this.employeeSummary[record.nama].hadir++; // Tetap dihitung hadir
                }
            }
            
            this.showProgress('Menyiapkan hasil...', 80);
            
            // Display results
            this.displayOvertimeData();
            this.displaySummary();
            this.updateCharts();
            
            // Enable export button
            document.getElementById('exportBtn').disabled = false;
            
            this.showProgress('Perhitungan selesai!', 100);
            
            setTimeout(() => {
                this.hideProgress();
                this.logActivity(`Perhitungan lembur selesai: ${this.overtimeData.length} catatan lembur`, 'success');
                this.switchResultsTab('overtime');
            }, 500);
            
        } catch (error) {
            console.error('Error calculating overtime:', error);
            this.showError('Terjadi kesalahan saat menghitung lembur');
            this.hideProgress();
        }
    }

    roundHours(hours) {
        switch (this.roundingType) {
            case 'down':
                return Math.floor(hours);
            case 'up':
                return Math.ceil(hours);
            case 'nearest':
                return Math.round(hours);
            default:
                return Math.floor(hours);
        }
    }

    getDayNameFromDate(dateStr) {
        try {
            let date;
            
            if (dateStr.includes('-')) {
                const parts = dateStr.split('-');
                if (parts.length === 3) {
                    date = new Date(parts[2], parts[1] - 1, parts[0]);
                }
            } else if (dateStr.includes('/')) {
                const parts = dateStr.split('/');
                if (parts.length === 3) {
                    date = new Date(parts[2], parts[1] - 1, parts[0]);
                }
            } else {
                date = new Date(dateStr);
            }
            
            if (isNaN(date.getTime())) {
                return 'N/A';
            }
            
            const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
            return days[date.getDay()];
        } catch (error) {
            return 'N/A';
        }
    }

    formatDateForExcel(dateStr) {
        try {
            let date;
            
            if (dateStr.includes('-')) {
                const parts = dateStr.split('-');
                if (parts.length === 3) {
                    return `${parts[0].padStart(2, '0')}-${parts[1].padStart(2, '0')}-${parts[2]}`;
                }
            }
            
            return dateStr;
        } catch (error) {
            return dateStr;
        }
    }

    displayAttendanceData() {
        const tbody = document.getElementById('attendanceBody');
        tbody.innerHTML = '';
        
        if (this.attendanceData.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="7" class="text-center">Tidak ada data absensi</td>
                </tr>
            `;
            return;
        }
        
        this.attendanceData.forEach(record => {
            const durasi = record.jamPulang ? 
                this.calculateDuration(record.jamMasuk, record.jamPulang) : '-';
            
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${record.tanggal}</td>
                <td>${record.nama}</td>
                <td>${record.jamMasuk}</td>
                <td>${record.jamPulang || '-'}</td>
                <td>${durasi}</td>
                <td><span class="status-badge status-${record.status.toLowerCase()}">${record.status}</span></td>
                <td>${record.jamPulang ? 'Lengkap' : 'Tidak lengkap'}</td>
            `;
            tbody.appendChild(row);
        });
    }

    displayOvertimeData() {
        const tbody = document.getElementById('overtimeBody');
        tbody.innerHTML = '';
        
        if (this.overtimeData.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="8" class="text-center">Tidak ada data lembur</td>
                </tr>
            `;
            return;
        }
        
        this.overtimeData.forEach(record => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${record.nama}</td>
                <td>${record.hari}</td>
                <td>${record.tanggal}</td>
                <td>${record.masuk}</td>
                <td>${record.pulang}</td>
                <td>${record.jamKerja} jam</td>
                <td>${record.lembur.toFixed(2)} jam (${record.jamDibulatkan} jam dibulatkan)</td>
                <td><span class="status-badge status-overtime">Lembur</span></td>
            `;
            tbody.appendChild(row);
        });
    }

    displaySummary() {
        const tbody = document.getElementById('summaryBody');
        tbody.innerHTML = '';
        
        const employees = Object.keys(this.employeeSummary);
        
        if (employees.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="5" class="text-center">Tidak ada data summary</td>
                </tr>
            `;
            return;
        }
        
        let totalHadir = 0;
        let totalTerlambat = 0;
        let totalLembur = 0;
        let totalGaji = 0;
        
        employees.forEach(nama => {
            const summary = this.employeeSummary[nama];
            
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${nama}</td>
                <td>${summary.hadir}</td>
                <td>${summary.terlambat}</td>
                <td>${summary.totalLembur.toFixed(2)} jam</td>
                <td>Rp ${summary.totalGaji.toLocaleString('id-ID')}</td>
            `;
            tbody.appendChild(row);
            
            totalHadir += summary.hadir;
            totalTerlambat += summary.terlambat;
            totalLembur += summary.totalLembur;
            totalGaji += summary.totalGaji;
        });
        
        // Update summary cards
        document.getElementById('totalEmployees').textContent = employees.length;
        document.getElementById('totalOvertimeHours').textContent = totalLembur.toFixed(2) + ' jam';
        document.getElementById('totalOvertimePay').textContent = 'Rp ' + totalGaji.toLocaleString('id-ID');
        
        const attendanceRate = employees.length > 0 ? 
            Math.round((totalHadir / (employees.length * this.attendanceData.length / employees.length)) * 100) : 0;
        document.getElementById('attendanceRate').textContent = attendanceRate + '%';
    }

    calculateDuration(startTime, endTime) {
        const start = this.timeToMinutes(startTime);
        const end = this.timeToMinutes(endTime);
        
        let duration = end - start;
        if (duration < 0) duration += 24 * 60;
        
        const hours = Math.floor(duration / 60);
        const minutes = duration % 60;
        
        return `${hours} jam ${minutes} menit`;
    }

    filterOvertimeData(searchTerm) {
        const tbody = document.getElementById('overtimeBody');
        const rows = tbody.getElementsByTagName('tr');
        
        for (let row of rows) {
            const text = row.textContent.toLowerCase();
            if (text.includes(searchTerm.toLowerCase())) {
                row.style.display = '';
            } else {
                row.style.display = 'none';
            }
        }
    }

    updateCharts() {
        // Destroy existing charts
        if (this.overtimeChart) {
            this.overtimeChart.destroy();
        }
        if (this.employeeChart) {
            this.employeeChart.destroy();
        }
        
        // Prepare data for charts
        const employees = Object.keys(this.employeeSummary);
        const overtimeHours = employees.map(emp => this.employeeSummary[emp].totalLembur);
        const overtimePay = employees.map(emp => this.employeeSummary[emp].totalGaji / 1000); // dalam ribuan
        
        // Overtime Distribution Chart
        const overtimeCtx = document.createElement('canvas');
        document.getElementById('overtimeChart').innerHTML = '';
        document.getElementById('overtimeChart').appendChild(overtimeCtx);
        
        this.overtimeChart = new Chart(overtimeCtx, {
            type: 'pie',
            data: {
                labels: employees,
                datasets: [{
                    data: overtimeHours,
                    backgroundColor: [
                        '#FF6384', '#36A2EB', '#FFCE56', '#4BC0C0',
                        '#9966FF', '#FF9F40', '#8AC926', '#1982C4'
                    ]
                }]
            },
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'right'
                    },
                    title: {
                        display: true,
                        text: 'Distribusi Jam Lembur'
                    }
                }
            }
        });
        
        // Employee Overtime Chart
        const employeeCtx = document.createElement('canvas');
        document.getElementById('employeeChart').innerHTML = '';
        document.getElementById('employeeChart').appendChild(employeeCtx);
        
        this.employeeChart = new Chart(employeeCtx, {
            type: 'bar',
            data: {
                labels: employees,
                datasets: [{
                    label: 'Jam Lembur',
                    data: overtimeHours,
                    backgroundColor: '#36A2EB'
                }, {
                    label: 'Gaji Lembur (ribu)',
                    data: overtimePay,
                    backgroundColor: '#FFCE56',
                    yAxisID: 'y1'
                }]
            },
            options: {
                responsive: true,
                scales: {
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Jam Lembur'
                        }
                    },
                    y1: {
                        position: 'right',
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'Gaji (ribu Rp)'
                        }
                    }
                },
                plugins: {
                    title: {
                        display: true,
                        text: 'Jam & Gaji Lembur per Karyawan'
                    }
                }
            }
        });
    }

    exportToExcel() {
        if (this.overtimeData.length === 0 && this.attendanceData.length === 0) {
            this.showError('Tidak ada data untuk diexport');
            return;
        }

        try {
            this.showProgress('Membuat file Excel...', 30);
            
            // Create workbook
            const wb = XLSX.utils.book_new();
            
            // Sheet 1: Data Lembur (format seperti contoh)
            const overtimeSheet = this.createOvertimeSheet();
            XLSX.utils.book_append_sheet(wb, overtimeSheet, "LEMBUR");
            
            this.showProgress('Membuat sheet absensi...', 60);
            
            // Sheet 2: Data Absensi
            const attendanceSheet = this.createAttendanceSheet();
            XLSX.utils.book_append_sheet(wb, attendanceSheet, "ABSENSI");
            
            this.showProgress('Membuat sheet ringkasan...', 80);
            
            // Sheet 3: Ringkasan
            const summarySheet = this.createSummarySheet();
            XLSX.utils.book_append_sheet(wb, summarySheet, "RINGKASAN");
            
            // Export file
            const fileName = `LEMBUR_${this.period.replace(/\s+/g, '_')}_${new Date().getTime()}.xlsx`;
            
            this.showProgress('Mengekspor file...', 95);
            
            XLSX.writeFile(wb, fileName);
            
            this.showProgress('Export selesai!', 100);
            
            setTimeout(() => {
                this.hideProgress();
                this.logActivity(`File Excel berhasil diexport: ${fileName}`, 'success');
            }, 500);
            
        } catch (error) {
            console.error('Error exporting to Excel:', error);
            this.showError('Gagal mengekspor ke Excel: ' + error.message);
            this.hideProgress();
        }
    }

    createOvertimeSheet() {
        const data = [];
        
        // Header utama
        data.push([`LEMBUR SEMESTER GANJIL TAHUN AKADEMIK 2025/2026`]);
        data.push([`BULAN ${this.period.toUpperCase()}`]);
        data.push([]);
        
        // Header tabel
        data.push(['Name', 'Hari', 'Tanggal', 'IN', 'OUT', 'JAM KERJA', 'TOTAL LEMBUR', 'TANDA TANGAN']);
        
        // Data lembur
        this.overtimeData.forEach(record => {
            data.push([
                record.nama,
                record.hari,
                record.tanggal,
                record.masuk,
                record.pulang,
                record.jamKerja,
                record.lembur.toFixed(2),
                ''
            ]);
        });
        
        // Total per karyawan
        const employees = [...new Set(this.overtimeData.map(r => r.nama))];
        employees.forEach(employee => {
            const employeeRecords = this.overtimeData.filter(r => r.nama === employee);
            const totalLembur = employeeRecords.reduce((sum, r) => sum + r.lembur, 0);
            const totalDibulatkan = employeeRecords.reduce((sum, r) => sum + r.jamDibulatkan, 0);
            const totalGaji = employeeRecords.reduce((sum, r) => sum + r.gajiLembur, 0);
            
            data.push([]);
            data.push(['', '', '', '', '', `Total ${employee}:`, totalLembur.toFixed(2), '']);
            data.push(['', '', '', '', '', 'Dibulatkan:', totalDibulatkan, '']);
            data.push(['', '', '', '', '', 'Gaji Lembur:', `Rp ${totalGaji.toLocaleString('id-ID')}`, '']);
        });
        
        // Create worksheet
        const ws = XLSX.utils.aoa_to_sheet(data);
        
        // Set column widths
        const wscols = [
            {wch: 15}, {wch: 10}, {wch: 12}, {wch: 8}, 
            {wch: 8}, {wch: 10}, {wch: 12}, {wch: 20}
        ];
        ws['!cols'] = wscols;
        
        return ws;
    }

    createAttendanceSheet() {
        const data = [
            ['DATA ABSENSI'],
            [`Periode: ${this.period}`],
            []
        ];
        
        // Header
        data.push(['Tanggal', 'Nama', 'Jam Masuk', 'Jam Pulang', 'Durasi', 'Status', 'Keterangan']);
        
        // Data
        this.attendanceData.forEach(record => {
            const durasi = record.jamPulang ? 
                this.calculateDuration(record.jamMasuk, record.jamPulang) : '-';
            
            data.push([
                record.tanggal,
                record.nama,
                record.jamMasuk,
                record.jamPulang || '-',
                durasi,
                record.status,
                record.jamPulang ? 'Lengkap' : 'Tidak lengkap'
            ]);
        });
        
        return XLSX.utils.aoa_to_sheet(data);
    }

    createSummarySheet() {
        const data = [
            ['RINGKASAN LEMBUR DAN ABSENSI'],
            [`Periode: ${this.period}`],
            [''],
            ['Nama', 'Hadir', 'Terlambat', 'Total Jam Lembur', 'Gaji Lembur', 'Tanda Tangan']
        ];
        
        const employees = Object.keys(this.employeeSummary);
        let totalHadir = 0;
        let totalTerlambat = 0;
        let totalLembur = 0;
        let totalGaji = 0;
        
        employees.forEach(employee => {
            const summary = this.employeeSummary[employee];
            data.push([
                employee,
                summary.hadir,
                summary.terlambat,
                summary.totalLembur.toFixed(2),
                `Rp ${summary.totalGaji.toLocaleString('id-ID')}`,
                ''
            ]);
            
            totalHadir += summary.hadir;
            totalTerlambat += summary.terlambat;
            totalLembur += summary.totalLembur;
            totalGaji += summary.totalGaji;
        });
        
        // Total
        data.push(['']);
        data.push([
            'TOTAL',
            totalHadir,
            totalTerlambat,
            totalLembur.toFixed(2),
            `Rp ${totalGaji.toLocaleString('id-ID')}`,
            ''
        ]);
        
        // Settings info
        data.push(['']);
        data.push(['PENGATURAN:']);
        data.push(['Jam Mulai Kerja:', this.startTime]);
        data.push(['Jam Selesai Kerja:', this.endTime]);
        data.push(['Jam Kerja Normal:', `${this.workingHours} jam`]);
        data.push(['Rate Lembur:', `Rp ${this.overtimeRate.toLocaleString('id-ID')} per jam`]);
        data.push(['Pembulatan:', this.roundingType === 'down' ? 'Pembulatan ke Bawah' : 
                  this.roundingType === 'up' ? 'Pembulatan ke Atas' : 'Pembulatan Terdekat']);
        data.push(['Toleransi Keterlambatan:', `${this.tolerance} menit`]);
        
        return XLSX.utils.aoa_to_sheet(data);
    }

    resetAll() {
        if (!confirm('Apakah Anda yakin ingin mereset semua data?')) {
            return;
        }
        
        this.rawData = [];
        this.attendanceData = [];
        this.overtimeData = [];
        this.employeeSummary = {};
        this.fileName = '';
        
        // Reset UI
        document.getElementById('fileInfo').style.display = 'none';
        document.getElementById('excelFile').value = '';
        document.getElementById('analyzeBtn').disabled = true;
        document.getElementById('calculateBtn').disabled = true;
        document.getElementById('exportBtn').disabled = true;
        
        // Clear tables
        document.getElementById('attendanceBody').innerHTML = '';
        document.getElementById('overtimeBody').innerHTML = '';
        document.getElementById('summaryBody').innerHTML = '';
        
        // Clear summary cards
        document.getElementById('totalEmployees').textContent = '0';
        document.getElementById('totalOvertimeHours').textContent = '0';
        document.getElementById('totalOvertimePay').textContent = 'Rp 0';
        document.getElementById('attendanceRate').textContent = '0%';
        
        // Clear charts
        if (this.overtimeChart) {
            this.overtimeChart.destroy();
            this.overtimeChart = null;
        }
        if (this.employeeChart) {
            this.employeeChart.destroy();
            this.employeeChart = null;
        }
        
        document.getElementById('overtimeChart').innerHTML = '<p>Grafik akan tampil setelah perhitungan</p>';
        document.getElementById('employeeChart').innerHTML = '<p>Grafik akan tampil setelah perhitungan</p>';
        
        this.logActivity('Semua data telah direset', 'info');
        this.showError('');
    }

    showProgress(message, percent) {
        const container = document.getElementById('progressContainer');
        const fill = document.getElementById('progressFill');
        const text = document.getElementById('progressText');
        
        container.style.display = 'block';
        fill.style.width = percent + '%';
        text.textContent = message;
    }

    hideProgress() {
        const container = document.getElementById('progressContainer');
        container.style.display = 'none';
    }

    showError(message) {
        // Log error
        if (message) {
            this.logActivity(message, 'error');
        }
    }

    logActivity(message, type = 'info') {
        const logContainer = document.getElementById('activityLog');
        const logItem = document.createElement('div');
        logItem.className = `log-item ${type}`;
        
        const icon = type === 'success' ? 'fa-check-circle' : 
                     type === 'error' ? 'fa-exclamation-triangle' : 
                     type === 'warning' ? 'fa-exclamation-circle' : 'fa-info-circle';
        
        const now = new Date();
        const timeStr = now.toLocaleTimeString('id-ID', { 
            hour: '2-digit', 
            minute: '2-digit',
            second: '2-digit'
        });
        
        logItem.innerHTML = `
            <i class="fas ${icon}"></i>
            <span>${message}</span>
            <span class="log-time">${timeStr}</span>
        `;
        
        // Add to top
        logContainer.insertBefore(logItem, logContainer.firstChild);
        
        // Limit to 50 items
        if (logContainer.children.length > 50) {
            logContainer.removeChild(logContainer.lastChild);
        }
    }
}

// Initialize when page loads
document.addEventListener('DOMContentLoaded', () => {
    window.calculator = new AttendanceOvertimeCalculator();
});
