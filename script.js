// Data Storage dengan Excel Integration
class AttendanceSystem {
    constructor() {
        this.employees = this.getDefaultEmployees();
        this.attendance = [];
        this.overtimeData = []; // Data lembur
        this.excelData = null;
        this.currentFileName = 'LEMBUR_SEMESTER_GANJIL_OKTOBER_2025.xlsx';
        this.currentChart = null;
        this.overtimeRate = 12500; // Rate lembur per jam
        this.init();
    }

    getDefaultEmployees() {
        // Data dari Excel yang diberikan
        return [
            { id: 1, name: "ATI", position: "Staff", department: "Administrasi", status: "active" },
            { id: 2, name: "IRVAN", position: "Dosen", department: "Teknik Mesin", status: "active" },
            { id: 3, name: "WINDI", position: "Staff", department: "Administrasi", status: "active" },
            { id: 4, name: "ARDHI", position: "Dosen", department: "Teknik Informatika", status: "active" },
            { id: 5, name: "SAJI", position: "Staff", department: "Teknik Elektro", status: "active" },
            { id: 6, name: "WAHYU", position: "Admin", department: "Administrasi", status: "active" },
            { id: 7, name: "ELZI", position: "Staff", department: "Administrasi", status: "active" },
            { id: 8, name: "NANANG", position: "Staff", department: "Teknik Elektro", status: "active" },
            { id: 9, name: "DEVI", position: "Admin", department: "Administrasi", status: "active" },
            { id: 10, name: "INTAN", position: "Staff", department: "Teknik Informatika", status: "active" },
            { id: 11, name: "ALIFAH", position: "Staff", department: "Administrasi", status: "active" },
            { id: 12, name: "PEBI", position: "Staff", department: "Administrasi", status: "active" },
            { id: 13, name: "DIAN", position: "Staff", department: "Administrasi", status: "active" },
            { id: 14, name: "RAFLY", position: "Staff", department: "Administrasi", status: "active" },
            { id: 15, name: "ERNI", position: "Staff", department: "Administrasi", status: "active" }
        ];
    }

    init() {
        this.setupEventListeners();
        this.updateDateTime();
        this.loadEmployees();
        this.loadAttendanceData();
        this.updateDashboard();
        this.loadOvertimeData();
        setInterval(() => this.updateDateTime(), 1000);
        
        // Coba load data dari localStorage sebagai backup
        this.loadFromLocalStorage();
        
        // Set default date to today
        this.setCurrentTime();
    }

    setupEventListeners() {
        // Navigation
        document.querySelectorAll('.nav-item').forEach(item => {
            item.addEventListener('click', (e) => this.switchTab(e.currentTarget.dataset.tab));
        });

        // Attendance buttons
        document.getElementById('btnCheckIn')?.addEventListener('click', () => this.recordAttendance('in'));
        document.getElementById('btnCheckOut')?.addEventListener('click', () => this.recordAttendance('out'));
        document.getElementById('btnNow')?.addEventListener('click', () => this.setCurrentTime());
        document.getElementById('refreshData')?.addEventListener('click', () => this.loadAttendanceData());

        // Overtime Calculation
        const overtimeForm = document.getElementById('overtimeForm');
        if (overtimeForm) {
            overtimeForm.addEventListener('submit', (e) => this.addOvertimeRecord(e));
        }
        
        // Reset overtime form
        document.getElementById('btnResetOvertime')?.addEventListener('click', () => this.resetOvertimeForm());
        
        // Export & Import
        document.getElementById('exportExcel').addEventListener('click', () => this.exportToExcel());
        document.getElementById('importExcel')?.addEventListener('click', () => this.importExcel());
        document.getElementById('excelFileInput').addEventListener('change', (e) => this.handleFileImport(e));

        // Employee management
        document.getElementById('addEmployeeBtn')?.addEventListener('click', () => this.showEmployeeModal());
        document.getElementById('employeeForm')?.addEventListener('submit', (e) => this.saveEmployee(e));
        document.getElementById('cancelEmployee')?.addEventListener('click', () => this.hideEmployeeModal());

        // Reports
        document.getElementById('generateReport')?.addEventListener('click', () => this.generateReport());
        document.getElementById('reportPeriod')?.addEventListener('change', (e) => this.toggleCustomDateRange(e.target.value));

        // Modal close
        document.querySelector('.close')?.addEventListener('click', () => this.hideEmployeeModal());
        window.addEventListener('click', (e) => {
            if (e.target === document.getElementById('employeeModal')) {
                this.hideEmployeeModal();
            }
        });

        // Filters
        document.getElementById('filterEmployee')?.addEventListener('change', () => this.filterAttendance());
        document.getElementById('filterDate')?.addEventListener('change', () => this.filterAttendance());
        document.getElementById('clearFilters')?.addEventListener('click', () => this.clearFilters());
    }

    switchTab(tabName) {
        document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
        const targetTab = document.querySelector(`[data-tab="${tabName}"]`);
        if (targetTab) {
            targetTab.classList.add('active');
        }

        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        const targetContent = document.getElementById(tabName);
        if (targetContent) {
            targetContent.classList.add('active');
        }

        const titles = {
            'dashboard': 'Dashboard Overview',
            'attendance': 'Sistem Absensi Digital',
            'overtime': 'Perhitungan Lembur',
            'reports': 'Laporan & Analytics',
            'employees': 'Manajemen Karyawan'
        };
        
        const pageTitle = document.getElementById('pageTitle');
        if (pageTitle) {
            pageTitle.textContent = titles[tabName] || 'Sistem Absensi Digital';
        }

        if (tabName === 'dashboard') {
            this.updateDashboard();
        } else if (tabName === 'employees') {
            this.loadEmployeesTable();
        } else if (tabName === 'overtime') {
            this.loadOvertimeData();
        }
    }

    updateDateTime() {
        const now = new Date();
        const options = { 
            weekday: 'long', 
            year: 'numeric', 
            month: 'long', 
            day: 'numeric',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        };
        
        const dateTimeElement = document.getElementById('currentDateTime');
        if (dateTimeElement) {
            dateTimeElement.textContent = now.toLocaleDateString('id-ID', options);
        }
    }

    setCurrentTime() {
        const now = new Date();
        
        const attendanceDate = document.getElementById('attendanceDate');
        const attendanceTime = document.getElementById('attendanceTime');
        const overtimeDate = document.getElementById('overtimeDate');
        
        if (attendanceDate) attendanceDate.value = now.toISOString().split('T')[0];
        if (attendanceTime) attendanceTime.value = now.toTimeString().slice(0, 5);
        if (overtimeDate) overtimeDate.value = now.toISOString().split('T')[0];
    }

    loadEmployees() {
        const select = document.getElementById('employeeSelect');
        const filterSelect = document.getElementById('filterEmployee');
        
        if (select) {
            while (select.options.length > 1) select.remove(1);
        }
        
        if (filterSelect) {
            while (filterSelect.options.length > 1) filterSelect.remove(1);
        }

        this.employees.forEach(emp => {
            if (emp.status === 'active') {
                if (select) {
                    const option = new Option(emp.name, emp.id);
                    select.add(option);
                }
                
                if (filterSelect) {
                    const filterOption = new Option(emp.name, emp.id);
                    filterSelect.add(filterOption);
                }
            }
        });
    }

    addOvertimeRecord(e) {
        e.preventDefault();
        
        const employeeName = document.getElementById('overtimeEmployee').value;
        const tanggal = document.getElementById('overtimeDate').value;
        const jamMasuk = document.getElementById('overtimeIn').value;
        const jamPulang = document.getElementById('overtimeOut').value;
        const jamKerjaNormal = parseFloat(document.getElementById('normalHours').value) || 8;

        // Validasi input
        if (!employeeName || !tanggal || !jamMasuk || !jamPulang) {
            this.showToast('Harap lengkapi semua field!', 'error');
            return;
        }

        // Validasi format jam masuk
        if (!this.validateTimeFormat(jamMasuk)) {
            this.showToast('Format jam masuk tidak valid! Gunakan format seperti 7.15', 'error');
            return;
        }

        // Validasi format jam pulang
        if (!this.validateTimeFormat(jamPulang)) {
            this.showToast('Format jam pulang tidak valid! Gunakan format seperti 16.02', 'error');
            return;
        }

        // Konversi waktu ke format desimal
        const waktuMasuk = this.timeToDecimal(jamMasuk);
        const waktuPulang = this.timeToDecimal(jamPulang);
        
        // Validasi waktu
        if (waktuMasuk >= 24 || waktuPulang >= 24) {
            this.showToast('Waktu tidak valid! Jam harus kurang dari 24', 'error');
            return;
        }

        // Hitung total jam kerja
        let totalJam = waktuPulang - waktuMasuk;
        if (totalJam < 0) totalJam += 24; // Jika melewati tengah malam
        
        // Validasi total jam
        if (totalJam > 24) {
            this.showToast('Total jam kerja tidak boleh lebih dari 24 jam!', 'error');
            return;
        }

        // Hitung jam lembur
        const jamLembur = totalJam - jamKerjaNormal;
        const jamLemburDecimal = jamLembur > 0 ? Math.round(jamLembur * 100) / 100 : 0;

        // Tentukan hari
        const hari = this.getDayName(tanggal);

        // Simpan data lembur
        const record = {
            id: Date.now(),
            employeeName: employeeName,
            hari: hari,
            tanggal: this.formatDateExcel(tanggal),
            in: jamMasuk,
            out: jamPulang,
            jamKerja: jamKerjaNormal,
            total: jamLemburDecimal,
            timestamp: new Date().toISOString()
        };

        // Cari atau buat data karyawan
        let employeeData = this.overtimeData.find(emp => emp.employeeName === employeeName);
        if (!employeeData) {
            employeeData = {
                employeeName: employeeName,
                records: [],
                totalJam: 0,
                totalDibulatkan: 0,
                gajiLembur: 0
            };
            this.overtimeData.push(employeeData);
        }

        employeeData.records.push(record);
        this.updateOvertimeSummary();
        this.loadOvertimeData();
        
        const message = jamLemburDecimal > 0 
            ? `Lembur ${employeeName}: ${jamLemburDecimal} jam (${jamMasuk} - ${jamPulang})`
            : `Tidak ada lembur untuk ${employeeName}`;
        
        this.showToast(message, jamLemburDecimal > 0 ? 'success' : 'info');
        
        // Reset form
        this.resetOvertimeForm();
    }

    validateTimeFormat(timeStr) {
        // Validasi format: angka.titik.angka (contoh: 7.15, 16.02)
        const regex = /^\d{1,2}\.\d{2}$/;
        if (!regex.test(timeStr)) return false;
        
        const parts = timeStr.split('.');
        const jam = parseInt(parts[0]);
        const menit = parseInt(parts[1]);
        
        return jam >= 0 && jam < 24 && menit >= 0 && menit < 60;
    }

    timeToDecimal(timeStr) {
        // Konversi "7.15" menjadi 7.25 (15 menit = 0.25 jam)
        const parts = timeStr.split('.');
        const jam = parseInt(parts[0]);
        const menit = parts[1] ? parseInt(parts[1]) : 0;
        return jam + (menit / 60);
    }

    decimalToTime(decimal) {
        // Konversi 7.25 menjadi "7.15"
        const jam = Math.floor(decimal);
        const menit = Math.round((decimal - jam) * 60);
        return `${jam}.${menit.toString().padStart(2, '0')}`;
    }

    getDayName(dateString) {
        const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
        const date = new Date(dateString);
        return days[date.getDay()];
    }

    formatDateExcel(dateString) {
        // Format: DD-MM-YYYY
        const date = new Date(dateString);
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        return `${day}-${month}-${year}`;
    }

    resetOvertimeForm() {
        const form = document.getElementById('overtimeForm');
        if (form) {
            form.reset();
            document.getElementById('normalHours').value = '8';
            this.setCurrentTime();
        }
    }

    updateOvertimeSummary() {
        let totalJamAll = 0;
        let totalDibulatkanAll = 0;
        let totalGajiAll = 0;

        this.overtimeData.forEach(employee => {
            const totalJam = employee.records.reduce((sum, record) => sum + record.total, 0);
            employee.totalJam = Math.round(totalJam * 100) / 100;
            employee.totalDibulatkan = Math.floor(employee.totalJam); // Pembulatan ke bawah
            
            // Hitung gaji lembur
            employee.gajiLembur = employee.totalDibulatkan * this.overtimeRate;
            
            // Update totals
            totalJamAll += employee.totalJam;
            totalDibulatkanAll += employee.totalDibulatkan;
            totalGajiAll += employee.gajiLembur;
        });

        // Update summary display
        const totalHoursElement = document.getElementById('totalOvertimeHours');
        const roundedHoursElement = document.getElementById('totalRoundedHours');
        const totalPayElement = document.getElementById('totalOvertimePay');
        
        if (totalHoursElement) {
            totalHoursElement.textContent = `${totalJamAll.toFixed(2)} jam`;
        }
        
        if (roundedHoursElement) {
            roundedHoursElement.textContent = `${totalDibulatkanAll} jam`;
        }
        
        if (totalPayElement) {
            totalPayElement.textContent = `Rp ${totalGajiAll.toLocaleString('id-ID')}`;
        }
    }

    loadOvertimeData() {
        const container = document.getElementById('overtimeTable');
        if (!container) return;

        container.innerHTML = '';
        
        if (this.overtimeData.length === 0) {
            container.innerHTML = `
                <tr>
                    <td colspan="9" class="text-center" style="padding: 3rem; color: var(--gray);">
                        <i class="fas fa-clock" style="font-size: 3rem; margin-bottom: 1rem; display: block;"></i>
                        <h4>Belum ada data lembur</h4>
                        <p>Silakan tambahkan data lembur menggunakan form di atas</p>
                    </td>
                </tr>
            `;
            this.updateOvertimeSummary();
            return;
        }

        // Urutkan data berdasarkan nama karyawan
        const sortedData = [...this.overtimeData].sort((a, b) => 
            a.employeeName.localeCompare(b.employeeName)
        );

        sortedData.forEach(employee => {
            // Urutkan records berdasarkan tanggal
            const sortedRecords = [...employee.records].sort((a, b) => {
                const dateA = new Date(a.tanggal.split('-').reverse().join('-'));
                const dateB = new Date(b.tanggal.split('-').reverse().join('-'));
                return dateA - dateB;
            });

            // Tambahkan data lembur
            sortedRecords.forEach((record, index) => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${employee.employeeName}</td>
                    <td>${record.hari}</td>
                    <td>${record.tanggal}</td>
                    <td>${record.in}</td>
                    <td>${record.out}</td>
                    <td>${record.jamKerja}</td>
                    <td>${record.total.toFixed(2)}</td>
                    <td></td>
                    <td>
                        <button class="btn btn-outline btn-sm" onclick="attendanceSystem.editOvertimeRecord('${employee.employeeName}', ${record.id})">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn btn-outline btn-sm" onclick="attendanceSystem.deleteOvertimeRecord('${employee.employeeName}', ${record.id})">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                `;
                container.appendChild(row);
            });

            // Tambahkan total untuk karyawan
            const totalRow = document.createElement('tr');
            totalRow.className = 'employee-total';
            totalRow.innerHTML = `
                <td colspan="6" style="text-align: right; font-weight: bold;">Total ${employee.employeeName}:</td>
                <td>${employee.totalJam.toFixed(2)}</td>
                <td colspan="2"></td>
            `;
            container.appendChild(totalRow);

            // Tambahkan pembayaran
            const paymentRow = document.createElement('tr');
            paymentRow.className = 'payment-row';
            paymentRow.innerHTML = `
                <td colspan="6" style="text-align: right; font-weight: bold;">Dibulatkan (${employee.totalDibulatkan} jam × Rp ${this.overtimeRate.toLocaleString('id-ID')}):</td>
                <td>Rp ${employee.gajiLembur.toLocaleString('id-ID')}</td>
                <td colspan="2"></td>
            `;
            container.appendChild(paymentRow);

            // Baris pemisah
            const separator = document.createElement('tr');
            separator.innerHTML = '<td colspan="9" style="height: 20px; background-color: #f8f9fa;"></td>';
            container.appendChild(separator);
        });

        this.updateOvertimeSummary();
    }

    // Excel Export Function dengan format seperti gambar
    exportToExcel() {
        try {
            // Create workbook
            const wb = XLSX.utils.book_new();
            
            // Sheet 1: Data Lembur (Format seperti gambar)
            const overtimeData = this.prepareOvertimeDataForExcel();
            const ws1 = XLSX.utils.aoa_to_sheet(overtimeData);
            
            // Set column widths
            const wscols = [
                {wch: 15}, // Name
                {wch: 10}, // Hari
                {wch: 12}, // Tanggal
                {wch: 8},  // IN
                {wch: 8},  // OUT
                {wch: 10}, // JAM KENA
                {wch: 10}, // TOTAL
                {wch: 15}  // TANDA TANGAN
            ];
            ws1['!cols'] = wscols;
            
            // Merge cells untuk judul
            ws1['!merges'] = [
                { s: { r: 0, c: 0 }, e: { r: 0, c: 7 } }, // Judul utama
                { s: { r: 1, c: 0 }, e: { r: 1, c: 7 } }  // Sub judul
            ];
            
            XLSX.utils.book_append_sheet(wb, ws1, "LEMBUR OKTOBER 2025");
            
            // Sheet 2: Ringkasan
            const summaryData = this.prepareSummaryDataForExcel();
            const ws2 = XLSX.utils.aoa_to_sheet(summaryData);
            XLSX.utils.book_append_sheet(wb, ws2, "RINGKASAN");
            
            // Export file
            XLSX.writeFile(wb, this.currentFileName);
            this.showToast('Data lembur berhasil diexport ke Excel!', 'success');
        } catch (error) {
            console.error('Export error:', error);
            this.showToast('Error saat export Excel: ' + error.message, 'error');
        }
    }

    prepareOvertimeDataForExcel() {
        const data = [];
        
        // Header utama sesuai gambar
        data.push(['LEMBUR SEMESTER GANJIL TAHUN AKADEMIK 2025/2026']);
        data.push(['BULAN OKTOBER 2025']);
        data.push([]); // Baris kosong
        
        // Header tabel
        data.push(['Name', 'Hari', 'Tanggal', 'IN', 'OUT', 'JAM KENA', 'TOTAL', 'TANDA TANGAN']);
        
        // Data untuk setiap karyawan
        this.overtimeData.forEach(employee => {
            // Urutkan records berdasarkan tanggal
            const sortedRecords = [...employee.records].sort((a, b) => {
                const dateA = new Date(a.tanggal.split('-').reverse().join('-'));
                const dateB = new Date(b.tanggal.split('-').reverse().join('-'));
                return dateA - dateB;
            });

            sortedRecords.forEach(record => {
                data.push([
                    employee.employeeName,
                    record.hari,
                    record.tanggal,
                    record.in,
                    record.out,
                    record.jamKerja,
                    record.total.toFixed(2),
                    '' // Kolom tanda tangan kosong
                ]);
            });
            
            // Total untuk karyawan
            data.push([
                '', '', '', '', '', 'Total:',
                employee.totalJam.toFixed(2),
                ''
            ]);
            
            // Baris kosong
            data.push([]);
            
            // Pembulatan dan pembayaran
            data.push([
                '', '', '', '', '', 'Dibulatkan:',
                employee.totalDibulatkan,
                ''
            ]);
            
            data.push([
                '', '', '', '', '', 'Rp',
                employee.gajiLembur,
                ''
            ]);
            
            // Baris pemisah
            data.push([]);
            data.push([]);
        });
        
        return data;
    }

    prepareSummaryDataForExcel() {
        const data = [
            ['RINGKASAN PEMBAYARAN LEMBUR - OKTOBER 2025'],
            [''],
            ['Nama', 'Total Jam', 'Dibulatkan', 'Rate/Jam', 'Total Gaji', 'Tanda Tangan']
        ];
        
        let totalJamAll = 0;
        let totalDibulatkanAll = 0;
        let totalGajiAll = 0;

        this.overtimeData.forEach(employee => {
            data.push([
                employee.employeeName,
                employee.totalJam.toFixed(2),
                employee.totalDibulatkan,
                this.overtimeRate.toLocaleString('id-ID'),
                `Rp ${employee.gajiLembur.toLocaleString('id-ID')}`,
                ''
            ]);
            
            totalJamAll += employee.totalJam;
            totalDibulatkanAll += employee.totalDibulatkan;
            totalGajiAll += employee.gajiLembur;
        });

        // Baris kosong
        data.push(['']);
        
        // Total keseluruhan
        data.push(['TOTAL', 
            totalJamAll.toFixed(2), 
            totalDibulatkanAll, 
            '', 
            `Rp ${totalGajiAll.toLocaleString('id-ID')}`,
            ''
        ]);
        
        data.push(['']);
        data.push(['Catatan:']);
        data.push(['- Rate lembur per jam: Rp 12.500']);
        data.push(['- Pembulatan ke bawah']);
        data.push(['- Berlaku untuk periode 1-31 Oktober 2025']);
        
        return data;
    }

    // Fungsi bantuan untuk Overtime
    editOvertimeRecord(employeeName, recordId) {
        const employee = this.overtimeData.find(emp => emp.employeeName === employeeName);
        if (employee) {
            const record = employee.records.find(r => r.id === recordId);
            if (record) {
                // Isi form dengan data yang ada
                document.getElementById('overtimeEmployee').value = employeeName;
                document.getElementById('overtimeDate').value = this.convertToDateInput(record.tanggal);
                document.getElementById('overtimeIn').value = record.in;
                document.getElementById('overtimeOut').value = record.out;
                document.getElementById('normalHours').value = record.jamKerja;
                
                // Hapus record lama
                employee.records = employee.records.filter(r => r.id !== recordId);
                this.showToast(`Edit data lembur ${employeeName} tanggal ${record.tanggal}`, 'info');
            }
        }
    }

    convertToDateInput(dateStr) {
        // Convert DD-MM-YYYY to YYYY-MM-DD
        const parts = dateStr.split('-');
        return `${parts[2]}-${parts[1]}-${parts[0]}`;
    }

    deleteOvertimeRecord(employeeName, recordId) {
        if (confirm('Hapus data lembur ini?')) {
            const employee = this.overtimeData.find(emp => emp.employeeName === employeeName);
            if (employee) {
                employee.records = employee.records.filter(r => r.id !== recordId);
                if (employee.records.length === 0) {
                    this.overtimeData = this.overtimeData.filter(emp => emp.employeeName !== employeeName);
                }
                this.updateOvertimeSummary();
                this.loadOvertimeData();
                this.showToast('Data lembur berhasil dihapus!', 'success');
            }
        }
    }

    // Toast Notification
    showToast(message, type = 'info') {
        const toastContainer = document.getElementById('toastContainer');
        if (!toastContainer) return;

        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.innerHTML = `
            <div class="toast-icon">
                <i class="fas fa-${type === 'success' ? 'check-circle' : 
                                 type === 'error' ? 'exclamation-triangle' : 
                                 type === 'warning' ? 'exclamation-circle' : 'info-circle'}"></i>
            </div>
            <div class="toast-message">${message}</div>
        `;
        
        toastContainer.appendChild(toast);
        
        // Auto remove after 5 seconds
        setTimeout(() => {
            if (toast.parentNode === toastContainer) {
                toast.remove();
            }
        }, 5000);
    }

    // Local Storage
    saveToLocalStorage() {
        try {
            localStorage.setItem('overtimeData', JSON.stringify({
                overtimeData: this.overtimeData,
                lastUpdate: new Date().toISOString()
            }));
        } catch (error) {
            console.error('Error saving to localStorage:', error);
        }
    }

    loadFromLocalStorage() {
        try {
            const saved = localStorage.getItem('overtimeData');
            if (saved) {
                const data = JSON.parse(saved);
                this.overtimeData = data.overtimeData || [];
                this.showToast('Data lembur dimuat dari penyimpanan lokal', 'success');
            }
        } catch (error) {
            console.error('Error loading from localStorage:', error);
        }
    }

    // Fungsi lain yang sudah ada (untuk compatibility)
    recordAttendance(type) {
        // Implementation for attendance recording
        this.showToast(`Fitur absensi ${type} (dalam pengembangan)`, 'info');
    }

    loadAttendanceData() {
        // Implementation for loading attendance data
        this.showToast('Memuat data absensi...', 'info');
    }

    filterAttendance() {
        // Implementation for filtering attendance
    }

    clearFilters() {
        // Implementation for clearing filters
    }

    updateDashboard() {
        // Implementation for updating dashboard
    }

    showEmployeeModal() {
        const modal = document.getElementById('employeeModal');
        if (modal) modal.style.display = 'block';
    }

    hideEmployeeModal() {
        const modal = document.getElementById('employeeModal');
        if (modal) modal.style.display = 'none';
    }

    saveEmployee(e) {
        e.preventDefault();
        this.showToast('Fitur manajemen karyawan (dalam pengembangan)', 'info');
        this.hideEmployeeModal();
    }

    loadEmployeesTable() {
        // Implementation for loading employees table
    }

    generateReport() {
        this.showToast('Fitur laporan (dalam pengembangan)', 'info');
    }

    toggleCustomDateRange(value) {
        const customRange = document.getElementById('customDateRange');
        if (customRange) {
            customRange.style.display = value === 'custom' ? 'block' : 'none';
        }
    }

    importExcel() {
        document.getElementById('excelFileInput').click();
    }

    handleFileImport(event) {
        this.showToast('Fitur import Excel (dalam pengembangan)', 'info');
    }
}

// Initialize the system when page loads
document.addEventListener('DOMContentLoaded', () => {
    window.attendanceSystem = new AttendanceSystem();
});
