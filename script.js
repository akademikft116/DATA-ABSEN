// Data Storage dengan Excel Integration
class AttendanceSystem {
    constructor() {
        this.employees = this.getDefaultEmployees();
        this.attendance = [];
        this.overtimeData = []; // Data lembur
        this.excelData = null;
        this.currentFileName = 'LEMBUR_OKTOBER_2025.xlsx';
        this.currentChart = null;
        this.overtimeRate = 12500; // Rate lembur per jam
        this.mealAllowance = 20000; // Uang makan
        this.transportAllowance = 15000; // Transport
        this.init();
    }

    getDefaultEmployees() {
        // Data dari Excel yang diberikan
        return [
            { id: 1, name: "ATI", position: "Staff", department: "Administrasi", status: "active" },
            { id: 2, name: "IRVAN", position: "Dosen", department: "Teknik Mesin", status: "active" }
        ];
    }

    init() {
        this.setupEventListeners();
        this.updateDateTime();
        this.loadEmployees();
        this.loadAttendanceData();
        this.updateDashboard();
        setInterval(() => this.updateDateTime(), 1000);
        
        // Coba load data dari localStorage sebagai backup
        this.loadFromLocalStorage();
        
        // Load data lembur contoh dari gambar
        this.loadSampleOvertimeData();
    }

    loadSampleOvertimeData() {
        // Data contoh dari gambar Excel
        this.overtimeData = [
            {
                employeeName: "ATI",
                records: [
                    { hari: "Senin", tanggal: "02-10-2025", in: "7.15", out: "16.02", jamKerja: 8, total: 0.87 },
                    { hari: "Senin", tanggal: "06-10-2025", in: "7.22", out: "16.07", jamKerja: 8, total: 0.85 },
                    { hari: "Kamis", tanggal: "09-10-2025", in: "7.15", out: "16.13", jamKerja: 8, total: 0.98 },
                    { hari: "Senin", tanggal: "13-10-2025", in: "7.21", out: "16.11", jamKerja: 8, total: 0.79 },
                    { hari: "Senin", tanggal: "20-10-2025", in: "7.25", out: "16.11", jamKerja: 8, total: 0.86 },
                    { hari: "Kamis", tanggal: "23-10-2025", in: "7.19", out: "16.03", jamKerja: 8, total: 0.84 },
                    { hari: "Kamis", tanggal: "30-10-2025", in: "7.32", out: "16.17", jamKerja: 8, total: 0.85 }
                ],
                totalJam: 6.04,
                totalDibulatkan: 6,
                gajiLembur: 75000,
                totalGaji: 75000
            },
            {
                employeeName: "IRVAN",
                records: [
                    { hari: "Rabu", tanggal: "08-10-2025", in: "7.04", out: "16.12", jamKerja: 8, total: 0.72 },
                    { hari: "Rabu", tanggal: "15-10-2025", in: "7.17", out: "16.07", jamKerja: 8, total: 0.9 },
                    { hari: "Rabu", tanggal: "29-10-2025", in: "7.38", out: "16.05", jamKerja: 8, total: 0.7 },
                    { hari: "Jumat", tanggal: "31-10-2025", in: "7.52", out: "15.00", jamKerja: 7, total: 0.68 }
                ],
                totalJam: 3,
                totalDibulatkan: 3,
                gajiLembur: 37500,
                totalGaji: 37500
            }
        ];
    }

    setupEventListeners() {
        // Navigation
        document.querySelectorAll('.nav-item').forEach(item => {
            item.addEventListener('click', (e) => this.switchTab(e.currentTarget.dataset.tab));
        });

        // Attendance buttons
        document.getElementById('btnCheckIn').addEventListener('click', () => this.recordAttendance('in'));
        document.getElementById('btnCheckOut').addEventListener('click', () => this.recordAttendance('out'));
        document.getElementById('btnNow').addEventListener('click', () => this.setCurrentTime());
        document.getElementById('refreshData').addEventListener('click', () => this.loadAttendanceData());

        // Overtime Calculation
        document.getElementById('calculateOvertime')?.addEventListener('click', () => this.calculateOvertime());
        
        // Export & Import
        document.getElementById('exportExcel').addEventListener('click', () => this.exportToExcel());
        document.getElementById('importExcel').addEventListener('click', () => this.importExcel());
        document.getElementById('excelFileInput').addEventListener('change', (e) => this.handleFileImport(e));

        // Employee management
        document.getElementById('addEmployeeBtn').addEventListener('click', () => this.showEmployeeModal());
        document.getElementById('employeeForm').addEventListener('submit', (e) => this.saveEmployee(e));
        document.getElementById('cancelEmployee').addEventListener('click', () => this.hideEmployeeModal());

        // Reports
        document.getElementById('generateReport').addEventListener('click', () => this.generateReport());
        document.getElementById('reportPeriod').addEventListener('change', (e) => this.toggleCustomDateRange(e.target.value));

        // Modal close
        document.querySelector('.close').addEventListener('click', () => this.hideEmployeeModal());
        window.addEventListener('click', (e) => {
            if (e.target === document.getElementById('employeeModal')) {
                this.hideEmployeeModal();
            }
        });

        // Filters
        document.getElementById('filterEmployee').addEventListener('change', () => this.filterAttendance());
        document.getElementById('filterDate').addEventListener('change', () => this.filterAttendance());
        document.getElementById('clearFilters').addEventListener('click', () => this.clearFilters());

        // Overtime form
        const overtimeForm = document.getElementById('overtimeForm');
        if (overtimeForm) {
            overtimeForm.addEventListener('submit', (e) => this.addOvertimeRecord(e));
        }
    }

    calculateOvertime() {
        const employeeName = document.getElementById('overtimeEmployee').value;
        const tanggal = document.getElementById('overtimeDate').value;
        const jamMasuk = document.getElementById('overtimeIn').value;
        const jamPulang = document.getElementById('overtimeOut').value;
        const jamKerjaNormal = parseFloat(document.getElementById('normalHours').value) || 8;

        if (!employeeName || !tanggal || !jamMasuk || !jamPulang) {
            this.showToast('Harap lengkapi semua field!', 'error');
            return;
        }

        // Konversi waktu ke format desimal
        const waktuMasuk = this.timeToDecimal(jamMasuk);
        const waktuPulang = this.timeToDecimal(jamPulang);
        
        // Hitung total jam kerja
        let totalJam = waktuPulang - waktuMasuk;
        if (totalJam < 0) totalJam += 24; // Jika melewati tengah malam
        
        // Hitung jam lembur
        const jamLembur = totalJam - jamKerjaNormal;
        const jamLemburDecimal = jamLembur > 0 ? Math.round(jamLembur * 100) / 100 : 0;

        // Tentukan hari
        const hari = this.getDayName(tanggal);

        // Simpan data lembur
        const record = {
            employeeName: employeeName,
            hari: hari,
            tanggal: this.formatDateExcel(tanggal),
            in: jamMasuk,
            out: jamPulang,
            jamKerja: jamKerjaNormal,
            total: jamLemburDecimal
        };

        // Cari atau buat data karyawan
        let employeeData = this.overtimeData.find(emp => emp.employeeName === employeeName);
        if (!employeeData) {
            employeeData = {
                employeeName: employeeName,
                records: [],
                totalJam: 0,
                totalDibulatkan: 0,
                gajiLembur: 0,
                totalGaji: 0
            };
            this.overtimeData.push(employeeData);
        }

        employeeData.records.push(record);
        this.updateOvertimeSummary();
        this.loadOvertimeData();
        this.showToast(`Data lembur ${employeeName} berhasil ditambahkan!`, 'success');
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

    updateOvertimeSummary() {
        this.overtimeData.forEach(employee => {
            const totalJam = employee.records.reduce((sum, record) => sum + record.total, 0);
            employee.totalJam = Math.round(totalJam * 100) / 100;
            employee.totalDibulatkan = Math.floor(totalJam); // Pembulatan ke bawah
            
            // Hitung gaji lembur
            employee.gajiLembur = employee.totalDibulatkan * this.overtimeRate;
            employee.totalGaji = employee.gajiLembur;
        });
    }

    loadOvertimeData() {
        const container = document.getElementById('overtimeTable');
        if (!container) return;

        container.innerHTML = '';
        
        this.overtimeData.forEach(employee => {
            // Tambahkan header untuk setiap karyawan
            const headerRow = document.createElement('tr');
            headerRow.innerHTML = `
                <td colspan="9" style="background-color: #f0f0f0; font-weight: bold; padding: 10px;">
                    ${employee.employeeName}
                </td>
            `;
            container.appendChild(headerRow);

            // Tambahkan data lembur
            employee.records.forEach((record, index) => {
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${record.hari}</td>
                    <td>${record.tanggal}</td>
                    <td>${record.in}</td>
                    <td>${record.out}</td>
                    <td>${record.jamKerja}</td>
                    <td>${record.total}</td>
                    <td>${this.mealAllowance.toLocaleString('id-ID')}</td>
                    <td>${this.transportAllowance.toLocaleString('id-ID')}</td>
                    <td>
                        <button class="btn btn-outline btn-sm" onclick="attendanceSystem.editOvertimeRecord('${employee.employeeName}', ${index})">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn btn-outline btn-sm" onclick="attendanceSystem.deleteOvertimeRecord('${employee.employeeName}', ${index})">
                            <i class="fas fa-trash"></i>
                        </button>
                    </td>
                `;
                container.appendChild(row);
            });

            // Tambahkan total untuk karyawan
            const totalRow = document.createElement('tr');
            totalRow.innerHTML = `
                <td colspan="5" style="text-align: right; font-weight: bold;">Total:</td>
                <td>${employee.totalJam}</td>
                <td colspan="2"></td>
                <td>
                    <button class="btn btn-outline btn-sm" onclick="attendanceSystem.calculateEmployeeOvertime('${employee.employeeName}')">
                        <i class="fas fa-calculator"></i>
                    </button>
                </td>
            `;
            container.appendChild(totalRow);

            // Tambahkan ringkasan pembayaran
            const paymentRow = document.createElement('tr');
            paymentRow.innerHTML = `
                <td colspan="5" style="text-align: right; font-weight: bold;">Total Dibulatkan:</td>
                <td>${employee.totalDibulatkan}</td>
                <td colspan="2"></td>
                <td></td>
            `;
            container.appendChild(paymentRow);

            const gajiRow = document.createElement('tr');
            gajiRow.innerHTML = `
                <td colspan="5" style="text-align: right; font-weight: bold;">Gaji Lembur (${this.overtimeRate.toLocaleString('id-ID')}/jam):</td>
                <td>Rp ${employee.gajiLembur.toLocaleString('id-ID')}</td>
                <td colspan="3"></td>
            `;
            container.appendChild(gajiRow);

            // Baris kosong untuk pemisah
            const separator = document.createElement('tr');
            separator.innerHTML = '<td colspan="9" style="height: 20px;"></td>';
            container.appendChild(separator);
        });
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
                {wch: 10}, // Name
                {wch: 10}, // Hari
                {wch: 12}, // Tanggal
                {wch: 8},  // IN
                {wch: 8},  // OUT
                {wch: 10}, // JAM KENA
                {wch: 10}, // TOTAL
                {wch: 15}  // TANDA TANGAN
            ];
            ws1['!cols'] = wscols;
            
            XLSX.utils.book_append_sheet(wb, ws1, "LEMBUR OKTOBER 2025");
            
            // Sheet 2: Ringkasan Pembayaran
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
        data.push(['LEMBUR SEMESTER GANJIL TAHUN AKADEMIK 2025/2026', '', '', '', '', '', '', '']);
        data.push(['BULAN OKTOBER 2025', '', '', '', '', '', '', '']);
        data.push(['', '', '', '', '', '', '', '']);
        
        // Header tabel
        data.push(['Name', 'Hari', 'Tanggal', 'IN', 'OUT', 'JAM KENA', 'TOTAL', 'TANDA TANGAN']);
        
        // Data untuk setiap karyawan
        this.overtimeData.forEach(employee => {
            employee.records.forEach(record => {
                data.push([
                    employee.employeeName,
                    record.hari,
                    record.tanggal,
                    record.in,
                    record.out,
                    record.jamKerja,
                    record.total,
                    '' // Kolom tanda tangan kosong
                ]);
            });
            
            // Total untuk karyawan
            data.push([
                '', '', '', '', '', '',
                employee.totalJam,
                ''
            ]);
            
            // Baris kosong
            data.push(['', '', '', '', '', '', '', '']);
            
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
            data.push(['', '', '', '', '', '', '', '']);
        });
        
        return data;
    }

    prepareSummaryDataForExcel() {
        const data = [
            ['RINGKASAN PEMBAYARAN LEMBUR - OKTOBER 2025'],
            [''],
            ['Nama', 'Total Jam', 'Dibulatkan', 'Rate/Jam', 'Total Gaji']
        ];
        
        this.overtimeData.forEach(employee => {
            data.push([
                employee.employeeName,
                employee.totalJam,
                employee.totalDibulatkan,
                this.overtimeRate.toLocaleString('id-ID'),
                employee.totalGaji.toLocaleString('id-ID')
            ]);
        });
        
        // Total keseluruhan
        const totalGaji = this.overtimeData.reduce((sum, emp) => sum + emp.totalGaji, 0);
        data.push(['']);
        data.push(['TOTAL KESELURUHAN', '', '', '', totalGaji.toLocaleString('id-ID')]);
        
        return data;
    }

    // Tambahkan fungsi untuk tab Lembur
    switchTab(tabName) {
        document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        document.getElementById(tabName).classList.add('active');

        const titles = {
            'dashboard': 'Dashboard Overview',
            'attendance': 'Sistem Absensi Digital',
            'reports': 'Laporan & Analytics',
            'employees': 'Manajemen Karyawan',
            'overtime': 'Perhitungan Lembur'
        };
        document.getElementById('pageTitle').textContent = titles[tabName] || 'Sistem Absensi Digital';

        if (tabName === 'dashboard') {
            this.updateDashboard();
        } else if (tabName === 'employees') {
            this.loadEmployeesTable();
        } else if (tabName === 'overtime') {
            this.loadOvertimeData();
        }
    }

    // Fungsi bantuan untuk Overtime
    editOvertimeRecord(employeeName, index) {
        const employee = this.overtimeData.find(emp => emp.employeeName === employeeName);
        if (employee && employee.records[index]) {
            const record = employee.records[index];
            // Tampilkan modal edit
            this.showToast(`Edit data lembur ${employeeName} tanggal ${record.tanggal}`, 'info');
        }
    }

    deleteOvertimeRecord(employeeName, index) {
        if (confirm('Hapus data lembur ini?')) {
            const employee = this.overtimeData.find(emp => emp.employeeName === employeeName);
            if (employee) {
                employee.records.splice(index, 1);
                if (employee.records.length === 0) {
                    this.overtimeData = this.overtimeData.filter(emp => emp.employeeName !== employeeName);
                }
                this.updateOvertimeSummary();
                this.loadOvertimeData();
                this.showToast('Data lembur berhasil dihapus!', 'success');
            }
        }
    }

    calculateEmployeeOvertime(employeeName) {
        const employee = this.overtimeData.find(emp => emp.employeeName === employeeName);
        if (employee) {
            this.showToast(`Menghitung ulang lembur untuk ${employeeName}: ${employee.totalJam} jam`, 'info');
        }
    }
}

// Initialize the system
const attendanceSystem = new AttendanceSystem();
