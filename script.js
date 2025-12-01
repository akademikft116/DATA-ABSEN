// script.js - VERSI DIPERBAIKI
class AttendanceSystem {
    constructor() {
        this.employees = this.getDefaultEmployees();
        this.attendance = [];
        this.excelData = null;
        this.currentFileName = 'DATA_ABSEN.xlsx';
        this.currentChart = null;
        this.init();
    }

    getDefaultEmployees() {
        return [
            { id: 1, name: "Windy", position: "Staff", department: "Administrasi", status: "active" },
            { id: 2, name: "Bu Ati", position: "Kepala Bagian", department: "Administrasi", status: "active" },
            { id: 3, name: "Pa Ardhi", position: "Dosen", department: "Teknik Informatika", status: "active" },
            { id: 4, name: "Pa Saji", position: "Staff", department: "Teknik Elektro", status: "active" },
            { id: 5, name: "Pa Irvan", position: "Dosen", department: "Teknik Mesin", status: "active" },
            { id: 6, name: "Bu Wahyu", position: "Admin", department: "Administrasi", status: "active" },
            { id: 7, name: "Bu Elzi", position: "Staff", department: "Administrasi", status: "active" },
            { id: 8, name: "Pa Nanang", position: "Staff", department: "Teknik Elektro", status: "active" },
            { id: 9, name: "Devi", position: "Admin", department: "Administrasi", status: "active" },
            { id: 10, name: "Intan", position: "Staff", department: "Teknik Informatika", status: "active" },
            { id: 11, name: "Alifah", position: "Staff", department: "Administrasi", status: "active" },
            { id: 12, name: "Pebi", position: "Staff", department: "Administrasi", status: "active" },
            { id: 13, name: "Dian", position: "Staff", department: "Administrasi", status: "active" },
            { id: 14, name: "Rafly", position: "Staff", department: "Teknik Informatika", status: "active" },
            { id: 15, name: "Erni", position: "Staff", department: "Administrasi", status: "active" }
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

        // Export & Import
        document.getElementById('exportExcel').addEventListener('click', () => this.exportToExcel());
        document.getElementById('importExcel').addEventListener('click', () => this.triggerFileImport());
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
        
        // Filters
        document.getElementById('filterEmployee').addEventListener('change', () => this.filterAttendance());
        document.getElementById('filterDate').addEventListener('change', () => this.filterAttendance());
        document.getElementById('clearFilters').addEventListener('click', () => this.clearFilters());
    }

    switchTab(tabName) {
        console.log('Switching to tab:', tabName);
        document.querySelectorAll('.nav-item').forEach(item => {
            item.classList.remove('active');
        });
        
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        document.querySelectorAll('.tab-content').forEach(content => {
            content.classList.remove('active');
        });
        
        const tabElement = document.getElementById(tabName);
        if (tabElement) {
            tabElement.classList.add('active');
            
            const titles = {
                'dashboard': 'Dashboard Overview',
                'attendance': 'Sistem Absensi Digital',
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
            } else if (tabName === 'attendance') {
                this.loadAttendanceData();
            }
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
        const dateInput = document.getElementById('attendanceDate');
        const timeInput = document.getElementById('attendanceTime');
        
        if (dateInput) {
            dateInput.value = now.toISOString().split('T')[0];
        }
        
        if (timeInput) {
            const hours = now.getHours().toString().padStart(2, '0');
            const minutes = now.getMinutes().toString().padStart(2, '0');
            timeInput.value = `${hours}:${minutes}`;
        }
    }

    loadEmployees() {
        const select = document.getElementById('employeeSelect');
        const filterSelect = document.getElementById('filterEmployee');
        
        if (!select || !filterSelect) return;

        // Clear existing options except first one
        while (select.options.length > 1) select.remove(1);
        while (filterSelect.options.length > 1) filterSelect.remove(1);

        this.employees.forEach(emp => {
            if (emp.status === 'active') {
                const option = new Option(emp.name, emp.id);
                const filterOption = new Option(emp.name, emp.id);
                select.add(option);
                filterSelect.add(filterOption);
            }
        });
    }

    recordAttendance(type) {
        const employeeId = document.getElementById('employeeSelect')?.value;
        const date = document.getElementById('attendanceDate')?.value;
        const time = document.getElementById('attendanceTime')?.value;

        if (!employeeId || !date || !time) {
            this.showToast('Harap lengkapi semua field!', 'error');
            return;
        }

        const employee = this.employees.find(emp => emp.id == employeeId);
        if (!employee) {
            this.showToast('Karyawan tidak ditemukan!', 'error');
            return;
        }

        const record = {
            id: Date.now(),
            employeeId: parseInt(employeeId),
            employeeName: employee.name,
            date: date,
            type: type,
            time: time,
            timestamp: new Date().toISOString()
        };

        this.attendance.push(record);
        this.saveToLocalStorage();
        this.loadAttendanceData();
        this.updateDashboard();

        const action = type === 'in' ? 'Check In' : 'Check Out';
        this.showToast(`${action} berhasil dicatat untuk ${employee.name}`, 'success');
        
        // Reset form
        document.getElementById('employeeSelect').value = '';
        this.setCurrentTime();
    }

    loadAttendanceData() {
        const tbody = document.getElementById('attendanceTable');
        if (!tbody) return;

        tbody.innerHTML = '';

        const filteredData = this.getFilteredAttendance();

        if (filteredData.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="7" class="text-center" style="padding: 2rem; color: var(--gray);">
                        <i class="fas fa-inbox" style="font-size: 3rem; margin-bottom: 1rem; display: block;"></i>
                        Tidak ada data absensi
                    </td>
                </tr>
            `;
            return;
        }

        // Group by employee and date
        const grouped = {};
        filteredData.forEach(record => {
            const key = `${record.employeeId}-${record.date}`;
            if (!grouped[key]) {
                grouped[key] = {
                    employeeName: record.employeeName,
                    date: record.date,
                    checkIn: null,
                    checkOut: null,
                    records: []
                };
            }
            grouped[key].records.push(record);
            
            if (record.type === 'in') {
                grouped[key].checkIn = record.time;
            } else {
                grouped[key].checkOut = record.time;
            }
        });

        // Create table rows
        Object.values(grouped).forEach(session => {
            const duration = this.calculateDuration(session.checkIn, session.checkOut);
            const status = this.getAttendanceStatus(session.checkIn, session.checkOut);
            
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${session.employeeName}</td>
                <td>${this.formatDate(session.date)}</td>
                <td>${session.checkIn || '-'}</td>
                <td>${session.checkOut || '-'}</td>
                <td>${duration}</td>
                <td><span class="status-badge status-${status}">${this.getStatusText(status)}</span></td>
                <td>
                    <button class="btn btn-outline btn-sm" onclick="attendanceSystem.editAttendance('${session.date}', '${session.employeeName}')">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn btn-outline btn-sm" onclick="attendanceSystem.deleteAttendance('${session.date}', '${session.employeeName}')">
                        <i class="fas fa-trash"></i>
                    </button>
                </td>
            `;
            tbody.appendChild(row);
        });
    }

    getFilteredAttendance() {
        const employeeFilter = document.getElementById('filterEmployee')?.value || '';
        const dateFilter = document.getElementById('filterDate')?.value || '';

        let filtered = this.attendance;

        if (employeeFilter) {
            filtered = filtered.filter(record => record.employeeId == employeeFilter);
        }

        if (dateFilter) {
            filtered = filtered.filter(record => record.date === dateFilter);
        }

        return filtered;
    }

    filterAttendance() {
        this.loadAttendanceData();
    }

    clearFilters() {
        const filterEmployee = document.getElementById('filterEmployee');
        const filterDate = document.getElementById('filterDate');
        
        if (filterEmployee) filterEmployee.value = '';
        if (filterDate) filterDate.value = '';
        
        this.loadAttendanceData();
    }

    calculateDuration(checkIn, checkOut) {
        if (!checkIn || !checkOut) return '-';
        
        try {
            const [inHours, inMinutes] = checkIn.split(':').map(Number);
            const [outHours, outMinutes] = checkOut.split(':').map(Number);
            
            const totalIn = inHours * 60 + inMinutes;
            const totalOut = outHours * 60 + outMinutes;
            
            let duration = totalOut - totalIn;
            if (duration < 0) duration += 24 * 60;
            
            const hours = Math.floor(duration / 60);
            const minutes = duration % 60;
            
            return `${hours}h ${minutes}m`;
        } catch (e) {
            return '-';
        }
    }

    getAttendanceStatus(checkIn, checkOut) {
        if (!checkIn) return 'absent';
        if (!checkOut) return 'present';
        
        try {
            const [inHours] = checkIn.split(':').map(Number);
            if (inHours > 9) return 'late';
            
            const duration = this.calculateDuration(checkIn, checkOut);
            const hours = parseInt(duration.split('h')[0]);
            if (hours > 9) return 'overtime';
        } catch (e) {
            return 'present';
        }
        
        return 'present';
    }

    getStatusText(status) {
        const statusText = {
            'present': 'Hadir',
            'late': 'Terlambat',
            'absent': 'Tidak Hadir',
            'overtime': 'Lembur'
        };
        return statusText[status] || status;
    }

    formatDate(dateString) {
        try {
            const date = new Date(dateString);
            return date.toLocaleDateString('id-ID', {
                weekday: 'long',
                year: 'numeric',
                month: 'long',
                day: 'numeric'
            });
        } catch (e) {
            return dateString;
        }
    }

    showToast(message, type = 'info') {
        const toastContainer = document.getElementById('toastContainer');
        if (!toastContainer) return;

        const toast = document.createElement('div');
        toast.className = `toast ${type}`;
        toast.innerHTML = `
            <div class="toast-icon">
                <i class="fas fa-${type === 'success' ? 'check' : type === 'error' ? 'exclamation-triangle' : 'info'}"></i>
            </div>
            <div class="toast-message">${message}</div>
        `;
        
        toastContainer.appendChild(toast);
        
        setTimeout(() => {
            if (toast.parentNode === toastContainer) {
                toast.remove();
            }
        }, 3000);
    }

    // SIMPLIFIED Excel Export Function
    exportToExcel() {
        try {
            // Create workbook
            const wb = XLSX.utils.book_new();
            
            // Sheet 1: Data Absensi (SIMPLIFIED VERSION)
            const attendanceData = this.prepareSimpleAttendanceData();
            const ws1 = XLSX.utils.aoa_to_sheet(attendanceData);
            XLSX.utils.book_append_sheet(wb, ws1, "DATA ABSENSI");
            
            // Sheet 2: Data Karyawan
            const employeeData = this.prepareEmployeeData();
            const ws2 = XLSX.utils.aoa_to_sheet(employeeData);
            XLSX.utils.book_append_sheet(wb, ws2, "DATA KARYAWAN");
            
            // Export file
            XLSX.writeFile(wb, 'Data_Absensi_' + new Date().toISOString().slice(0,10) + '.xlsx');
            this.showToast('Data berhasil diexport ke Excel!', 'success');
        } catch (error) {
            console.error('Export error:', error);
            this.showToast('Error saat export Excel: ' + error.message, 'error');
        }
    }

    prepareSimpleAttendanceData() {
        const data = [
            ['LAPORAN ABSENSI - UNIVERSITAS LANGLANGBUANA'],
            ['Tanggal Export: ' + new Date().toLocaleDateString('id-ID')],
            [''],
            ['Nama Karyawan', 'Tanggal', 'Check In', 'Check Out', 'Durasi', 'Status']
        ];
        
        // Group attendance by employee and date
        const grouped = {};
        this.attendance.forEach(record => {
            const key = `${record.employeeName}-${record.date}`;
            if (!grouped[key]) {
                grouped[key] = {
                    employeeName: record.employeeName,
                    date: record.date,
                    checkIn: null,
                    checkOut: null
                };
            }
            
            if (record.type === 'in') {
                grouped[key].checkIn = record.time;
            } else {
                grouped[key].checkOut = record.time;
            }
        });
        
        // Add data rows
        Object.values(grouped).forEach(session => {
            const duration = this.calculateDuration(session.checkIn, session.checkOut);
            const status = this.getAttendanceStatus(session.checkIn, session.checkOut);
            
            data.push([
                session.employeeName,
                session.date,
                session.checkIn || '-',
                session.checkOut || '-',
                duration,
                this.getStatusText(status)
            ]);
        });
        
        return data;
    }

    prepareEmployeeData() {
        const data = [
            ['DATA KARYAWAN - UNIVERSITAS LANGLANGBUANA'],
            [''],
            ['ID', 'Nama', 'Jabatan', 'Departemen', 'Status']
        ];
        
        this.employees.forEach(emp => {
            data.push([
                emp.id,
                emp.name,
                emp.position,
                emp.department,
                emp.status === 'active' ? 'Aktif' : 'Non-Aktif'
            ]);
        });
        
        return data;
    }

    triggerFileImport() {
        const fileInput = document.getElementById('excelFileInput');
        if (fileInput) {
            fileInput.click();
        }
    }

    handleFileImport(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Reset file input
                event.target.value = '';
                
                this.showToast('File Excel berhasil dibaca!', 'success');
            } catch (error) {
                console.error('Import error:', error);
                this.showToast('Error membaca file Excel', 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    saveToLocalStorage() {
        try {
            const data = {
                employees: this.employees,
                attendance: this.attendance,
                lastUpdate: new Date().toISOString()
            };
            localStorage.setItem('attendanceSystemData', JSON.stringify(data));
        } catch (error) {
            console.error('Error saving to localStorage:', error);
        }
    }

    loadFromLocalStorage() {
        try {
            const saved = localStorage.getItem('attendanceSystemData');
            if (saved) {
                const data = JSON.parse(saved);
                this.employees = data.employees || this.employees;
                this.attendance = data.attendance || this.attendance;
                console.log('Data loaded from localStorage');
            }
        } catch (error) {
            console.error('Error loading from localStorage:', error);
        }
    }

    showEmployeeModal() {
        const modal = document.getElementById('employeeModal');
        if (modal) {
            modal.style.display = 'block';
        }
    }

    hideEmployeeModal() {
        const modal = document.getElementById('employeeModal');
        if (modal) {
            modal.style.display = 'none';
            document.getElementById('employeeForm').reset();
        }
    }

    saveEmployee(e) {
        e.preventDefault();
        
        const name = document.getElementById('empName')?.value;
        const position = document.getElementById('empPosition')?.value;
        const department = document.getElementById('empDepartment')?.value;
        
        if (!name || !position || !department) {
            this.showToast('Harap lengkapi semua data karyawan!', 'error');
            return;
        }
        
        const newEmployee = {
            id: this.employees.length > 0 ? Math.max(...this.employees.map(e => e.id)) + 1 : 1,
            name: name,
            position: position,
            department: department,
            status: 'active'
        };
        
        this.employees.push(newEmployee);
        this.saveToLocalStorage();
        this.loadEmployees();
        this.loadEmployeesTable();
        this.hideEmployeeModal();
        this.showToast('Karyawan berhasil ditambahkan!', 'success');
    }

    loadEmployeesTable() {
        const tbody = document.getElementById('employeesTable');
        if (!tbody) return;
        
        tbody.innerHTML = '';
        
        this.employees.forEach(emp => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${emp.id}</td>
                <td>${emp.name}</td>
                <td>${emp.position}</td>
                <td>${emp.department}</td>
                <td><span class="status-badge status-${emp.status}">${emp.status === 'active' ? 'Aktif' : 'Non-Aktif'}</span></td>
                <td>
                    <button class="btn btn-outline btn-sm" onclick="attendanceSystem.editEmployee(${emp.id})">
                        <i class="fas fa-edit"></i>
                    </button>
                    <button class="btn btn-outline btn-sm" onclick="attendanceSystem.deleteEmployee(${emp.id})">
                        <i class="fas fa-trash"></i>
                    </button>
                </td>
            `;
            tbody.appendChild(row);
        });
    }

    updateDashboard() {
        // Update employee count
        const totalEmployees = document.getElementById('totalEmployees');
        if (totalEmployees) {
            totalEmployees.textContent = this.employees.filter(emp => emp.status === 'active').length;
        }
        
        // Update present today
        const today = new Date().toISOString().split('T')[0];
        const presentTodayElement = document.getElementById('presentToday');
        if (presentTodayElement) {
            const presentToday = this.attendance.filter(record => 
                record.date === today && record.type === 'in'
            ).length;
            presentTodayElement.textContent = presentToday;
        }
        
        // Update overtime hours
        const overtimeHoursElement = document.getElementById('overtimeHours');
        if (overtimeHoursElement) {
            // Count records with overtime status
            let overtimeCount = 0;
            const grouped = {};
            this.attendance.forEach(record => {
                const key = `${record.employeeId}-${record.date}`;
                if (!grouped[key]) {
                    grouped[key] = { checkIn: null, checkOut: null };
                }
                if (record.type === 'in') grouped[key].checkIn = record.time;
                else grouped[key].checkOut = record.time;
            });
            
            Object.values(grouped).forEach(session => {
                if (this.getAttendanceStatus(session.checkIn, session.checkOut) === 'overtime') {
                    overtimeCount++;
                }
            });
            overtimeHoursElement.textContent = overtimeCount;
        }
        
        // Update attendance rate
        const attendanceRateElement = document.getElementById('attendanceRate');
        if (attendanceRateElement) {
            const totalEmployeesCount = this.employees.filter(emp => emp.status === 'active').length;
            const presentTodayCount = this.attendance.filter(record => 
                record.date === today && record.type === 'in'
            ).length;
            const attendanceRate = totalEmployeesCount > 0 ? 
                Math.round((presentTodayCount / totalEmployeesCount) * 100) : 0;
            attendanceRateElement.textContent = `${attendanceRate}%`;
        }
        
        this.updateRecentActivity();
        this.updateChart();
    }

    updateRecentActivity() {
        const container = document.getElementById('recentActivity');
        if (!container) return;
        
        container.innerHTML = '';
        
        const recentRecords = this.attendance
            .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
            .slice(0, 5);
        
        if (recentRecords.length === 0) {
            container.innerHTML = '<div class="text-center" style="padding: 2rem; color: var(--gray);">Tidak ada aktivitas terbaru</div>';
            return;
        }
        
        recentRecords.forEach(record => {
            const activityItem = document.createElement('div');
            activityItem.className = 'activity-item';
            activityItem.innerHTML = `
                <div class="activity-icon ${record.type === 'in' ? 'checkin' : 'checkout'}">
                    <i class="fas fa-${record.type === 'in' ? 'sign-in-alt' : 'sign-out-alt'}"></i>
                </div>
                <div class="activity-info">
                    <div class="activity-name">${record.employeeName}</div>
                    <div class="activity-time">${record.type === 'in' ? 'Check In' : 'Check Out'} - ${this.formatDateTime(record.timestamp)}</div>
                </div>
            `;
            container.appendChild(activityItem);
        });
    }

    formatDateTime(timestamp) {
        try {
            const date = new Date(timestamp);
            return date.toLocaleString('id-ID', {
                day: '2-digit',
                month: '2-digit',
                year: 'numeric',
                hour: '2-digit',
                minute: '2-digit'
            });
        } catch (e) {
            return timestamp;
        }
    }

    updateChart() {
        const ctx = document.getElementById('attendanceChart');
        if (!ctx) return;
        
        if (this.currentChart) {
            this.currentChart.destroy();
        }
        
        const last7Days = this.getLast7Days();
        const chartData = last7Days.map(day => {
            return this.attendance.filter(record => 
                record.date === day && record.type === 'in'
            ).length;
        });
        
        this.currentChart = new Chart(ctx.getContext('2d'), {
            type: 'bar',
            data: {
                labels: last7Days.map(day => this.formatChartDate(day)),
                datasets: [{
                    label: 'Jumlah Hadir',
                    data: chartData,
                    backgroundColor: 'rgba(52, 152, 219, 0.8)',
                    borderColor: 'rgba(52, 152, 219, 1)',
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            stepSize: 1
                        }
                    }
                }
            }
        });
    }

    getLast7Days() {
        const days = [];
        for (let i = 6; i >= 0; i--) {
            const date = new Date();
            date.setDate(date.getDate() - i);
            days.push(date.toISOString().split('T')[0]);
        }
        return days;
    }

    formatChartDate(dateString) {
        try {
            const date = new Date(dateString);
            return date.toLocaleDateString('id-ID', {
                day: '2-digit',
                month: '2-digit'
            });
        } catch (e) {
            return dateString;
        }
    }

    // Placeholder functions
    editAttendance(date, employeeName) {
        this.showToast(`Edit absensi ${employeeName} pada ${date}`, 'info');
    }

    deleteAttendance(date, employeeName) {
        if (confirm(`Hapus absensi ${employeeName} pada ${date}?`)) {
            // Find and remove all records for this employee on this date
            this.attendance = this.attendance.filter(record => 
                !(record.date === date && record.employeeName === employeeName)
            );
            this.saveToLocalStorage();
            this.loadAttendanceData();
            this.updateDashboard();
            this.showToast('Absensi berhasil dihapus!', 'success');
        }
    }

    editEmployee(id) {
        const employee = this.employees.find(emp => emp.id === id);
        if (employee) {
            document.getElementById('empName').value = employee.name;
            document.getElementById('empPosition').value = employee.position;
            document.getElementById('empDepartment').value = employee.department;
            this.showEmployeeModal();
            this.showToast(`Edit data ${employee.name}`, 'info');
        }
    }

    deleteEmployee(id) {
        if (confirm('Hapus karyawan ini?')) {
            this.employees = this.employees.filter(emp => emp.id !== id);
            this.saveToLocalStorage();
            this.loadEmployees();
            this.loadEmployeesTable();
            this.updateDashboard();
            this.showToast('Karyawan berhasil dihapus!', 'success');
        }
    }

    generateReport() {
        this.showToast('Laporan sedang diproses...', 'info');
        // Simple report generation
        setTimeout(() => {
            const today = new Date().toISOString().split('T')[0];
            const presentCount = this.attendance.filter(record => 
                record.date === today && record.type === 'in'
            ).length;
            
            const reportResults = document.getElementById('reportResults');
            if (reportResults) {
                reportResults.innerHTML = `
                    <div class="card">
                        <div class="card-body">
                            <h4>Laporan Harian - ${this.formatDate(today)}</h4>
                            <p>Jumlah karyawan hadir: ${presentCount}</p>
                            <p>Total karyawan: ${this.employees.filter(e => e.status === 'active').length}</p>
                            <p>Persentase kehadiran: ${Math.round((presentCount / this.employees.filter(e => e.status === 'active').length) * 100)}%</p>
                        </div>
                    </div>
                `;
            }
        }, 1000);
    }

    toggleCustomDateRange(value) {
        const customRange = document.getElementById('customDateRange');
        if (customRange) {
            customRange.style.display = value === 'custom' ? 'block' : 'none';
        }
    }
}

// Initialize the system
let attendanceSystem;

document.addEventListener('DOMContentLoaded', () => {
    console.log('DOM fully loaded, initializing system...');
    attendanceSystem = new AttendanceSystem();
    
    // Setup modal click outside to close
    const modal = document.getElementById('employeeModal');
    if (modal) {
        window.addEventListener('click', (e) => {
            if (e.target === modal) {
                attendanceSystem.hideEmployeeModal();
            }
        });
    }
    
    console.log('System initialized successfully');
});
