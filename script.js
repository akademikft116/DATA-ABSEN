// Data Storage dengan Excel Integration
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
            { id: 10, name: "Intan", position: "Staff", department: "Teknik Informatika", status: "active" }
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
    }

    switchTab(tabName) {
        document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        document.getElementById(tabName).classList.add('active');

        const titles = {
            'dashboard': 'Dashboard Overview',
            'attendance': 'Sistem Absensi Digital',
            'reports': 'Laporan & Analytics',
            'employees': 'Manajemen Karyawan'
        };
        document.getElementById('pageTitle').textContent = titles[tabName] || 'Sistem Absensi Digital';

        if (tabName === 'dashboard') {
            this.updateDashboard();
        } else if (tabName === 'employees') {
            this.loadEmployeesTable();
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
        document.getElementById('currentDateTime').textContent = 
            now.toLocaleDateString('id-ID', options);
    }

    setCurrentTime() {
        const now = new Date();
        document.getElementById('attendanceDate').value = now.toISOString().split('T')[0];
        document.getElementById('attendanceTime').value = now.toTimeString().slice(0, 5);
    }

    loadEmployees() {
        const select = document.getElementById('employeeSelect');
        const filterSelect = document.getElementById('filterEmployee');
        
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
        const employeeId = document.getElementById('employeeSelect').value;
        const date = document.getElementById('attendanceDate').value;
        const time = document.getElementById('attendanceTime').value;

        if (!employeeId || !date || !time) {
            this.showToast('Harap lengkapi semua field!', 'error');
            return;
        }

        const employee = this.employees.find(emp => emp.id == employeeId);
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
                    checkOut: null
                };
            }
            
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
        const employeeFilter = document.getElementById('filterEmployee').value;
        const dateFilter = document.getElementById('filterDate').value;

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
        document.getElementById('filterEmployee').value = '';
        document.getElementById('filterDate').value = '';
        this.loadAttendanceData();
    }

    calculateDuration(checkIn, checkOut) {
        if (!checkIn || !checkOut) return '-';
        
        const [inHours, inMinutes] = checkIn.split(':').map(Number);
        const [outHours, outMinutes] = checkOut.split(':').map(Number);
        
        const totalIn = inHours * 60 + inMinutes;
        const totalOut = outHours * 60 + outMinutes;
        
        let duration = totalOut - totalIn;
        if (duration < 0) duration += 24 * 60;
        
        const hours = Math.floor(duration / 60);
        const minutes = duration % 60;
        
        return `${hours}h ${minutes}m`;
    }

    getAttendanceStatus(checkIn, checkOut) {
        if (!checkIn) return 'absent';
        if (!checkOut) return 'present';
        
        const [inHours] = checkIn.split(':').map(Number);
        if (inHours > 9) return 'late';
        
        const duration = this.calculateDuration(checkIn, checkOut);
        const hours = parseInt(duration.split('h')[0]);
        if (hours > 9) return 'overtime';
        
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
        const date = new Date(dateString);
        return date.toLocaleDateString('id-ID', {
            weekday: 'long',
            year: 'numeric',
            month: 'long',
            day: 'numeric'
        });
    }

    showToast(message, type = 'info') {
        const toastContainer = document.getElementById('toastContainer');
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
            toast.remove();
        }, 3000);
    }

    // Excel Export Function
    exportToExcel() {
        try {
            // Create workbook
            const wb = XLSX.utils.book_new();
            
            // Sheet 1: Data Absensi (Format seperti Excel Anda)
            const attendanceData = this.prepareAttendanceDataForExcel();
            const ws1 = XLSX.utils.aoa_to_sheet(attendanceData);
            XLSX.utils.book_append_sheet(wb, ws1, "NOVEMBER");
            
            // Sheet 2: Data Lembur
            const overtimeData = this.prepareOvertimeDataForExcel();
            const ws2 = XLSX.utils.aoa_to_sheet(overtimeData);
            XLSX.utils.book_append_sheet(wb, ws2, "LEMBUR");
            
            // Sheet 3: Data K3
            const k3Data = this.prepareK3DataForExcel();
            const ws3 = XLSX.utils.aoa_to_sheet(k3Data);
            XLSX.utils.book_append_sheet(wb, ws3, "NOVEMBER.25 K3");
            
            // Export file
            XLSX.writeFile(wb, this.currentFileName);
            this.showToast('Data berhasil diexport ke Excel!', 'success');
        } catch (error) {
            console.error('Export error:', error);
            this.showToast('Error saat export Excel: ' + error.message, 'error');
        }
    }

    prepareAttendanceDataForExcel() {
        const data = [];
        
        // Header sesuai format Excel Anda
        data.push(['', 'Windy', '2', '', 'Nama', 'Waktu']);
        data.push(['', 'Bu Ati', '4', '']);
        data.push(['', 'Pa Ardhi', '5', '']);
        data.push(['', 'Pa Saji', '8', '']);
        data.push(['', 'Pa Irvan', '9', '']);
        data.push(['', 'Bu Wahyu', '10', '']);
        data.push(['', 'Bu Elzi', '14', '']);
        data.push(['', 'Pa Nanang', '18', '']);
        data.push(['', 'Devi', '35', '']);
        data.push(['', 'Intan', '37', '']);
        data.push(['', 'Alifah', '38', '']);
        data.push(['', 'Pebi', '40', '']);
        data.push(['', 'Dian', '43', '']);
        data.push(['', 'Rafly', '44', '']);
        data.push(['', 'Erni', '47', '']);
        data.push(['', '', '', '']);
        
        // Add attendance records
        this.attendance.forEach(record => {
            const excelDate = this.convertToExcelDate(record.date, record.time);
            data.push(['', '', '', '', record.employeeName, excelDate]);
        });
        
        return data;
    }

    prepareOvertimeDataForExcel() {
        // Format data lembur sesuai Excel Anda
        const data = [
            ['LEMBUR SEMESTER GANJIL TAHUN AKADEMIK 2025/2026', '', '', '', '', '', '', '', '', '', '', '', '', '', 'RATE BAYARAN LEMBUR', ''],
            ['BULAN OKTOBER 2025', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            ['Name', 'Hari', 'Tanggal', 'IN', 'OUT', 'JAM KERJA', 'TOTAL', 'UANG MAKAN', 'TRANSPORT', 'TANDA TANGAN', '', '', '', 'KABAG/K.TU', '12500'],
            // ... tambahkan data lembur sesuai kebutuhan
        ];
        
        return data;
    }

    prepareK3DataForExcel() {
        // Format data K3 sesuai Excel Anda
        const data = [
            ['HONOR PIKET LEMBUR (PAGI & MALAM) K3', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'AP', 'AQ', 'AR', 'AS'],
            ['PERIODE BULAN OKTOBER  2025', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
            // ... tambahkan data K3 sesuai kebutuhan
        ];
        
        return data;
    }

    convertToExcelDate(dateString, timeString) {
        // Convert to Excel date format: DD/MM/YYYY HH:MM
        const date = new Date(dateString);
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear();
        
        return `${day}/${month}/${year} ${timeString}`;
    }

    // Excel Import Function
    importExcel() {
        document.getElementById('excelFileInput').click();
    }

    handleFileImport(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                // Process attendance data from NOVEMBER sheet
                if (workbook.Sheets['NOVEMBER']) {
                    this.processAttendanceSheet(workbook.Sheets['NOVEMBER']);
                }
                
                this.showToast('Data Excel berhasil diimport!', 'success');
                this.loadAttendanceData();
                this.updateDashboard();
            } catch (error) {
                console.error('Import error:', error);
                this.showToast('Error saat import Excel: ' + error.message, 'error');
            }
        };
        reader.readAsArrayBuffer(file);
    }

    processAttendanceSheet(worksheet) {
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        // Process data based on your Excel structure
        jsonData.forEach((row, index) => {
            if (row[4] && row[5] && typeof row[4] === 'string' && typeof row[5] === 'string') {
                // This is an attendance record
                const employeeName = row[4];
                const dateTime = row[5];
                
                // Parse the date and time
                const [datePart, timePart] = this.parseExcelDateTime(dateTime);
                
                if (datePart && timePart) {
                    // Add to attendance records
                    const employee = this.employees.find(emp => emp.name === employeeName);
                    if (employee) {
                        const record = {
                            id: Date.now() + index,
                            employeeId: employee.id,
                            employeeName: employeeName,
                            date: datePart,
                            type: this.determineAttendanceType(timePart),
                            time: timePart,
                            timestamp: new Date().toISOString()
                        };
                        this.attendance.push(record);
                    }
                }
            }
        });
        
        this.saveToLocalStorage();
    }

    parseExcelDateTime(dateTimeString) {
        // Parse various Excel date formats
        if (dateTimeString.includes('/')) {
            // Format: DD/MM/YYYY HH:MM
            const parts = dateTimeString.split(' ');
            if (parts.length === 2) {
                const dateParts = parts[0].split('/');
                const timeParts = parts[1].split(':');
                
                if (dateParts.length === 3 && timeParts.length === 2) {
                    const date = `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}`;
                    const time = `${timeParts[0]}:${timeParts[1]}`;
                    return [date, time];
                }
            }
        }
        
        return [null, null];
    }

    determineAttendanceType(time) {
        // Simple logic: before 12:00 is check-in, after is check-out
        const [hours] = time.split(':').map(Number);
        return hours < 12 ? 'in' : 'out';
    }

    // Local Storage Backup
    saveToLocalStorage() {
        localStorage.setItem('attendanceData', JSON.stringify({
            employees: this.employees,
            attendance: this.attendance,
            lastUpdate: new Date().toISOString()
        }));
    }

    loadFromLocalStorage() {
        const saved = localStorage.getItem('attendanceData');
        if (saved) {
            try {
                const data = JSON.parse(saved);
                this.employees = data.employees || this.employees;
                this.attendance = data.attendance || this.attendance;
                this.showToast('Data berhasil dimuat dari penyimpanan lokal', 'success');
            } catch (error) {
                console.error('Error loading from localStorage:', error);
            }
        }
    }

    // Employee Management
    showEmployeeModal() {
        document.getElementById('employeeModal').style.display = 'block';
    }

    hideEmployeeModal() {
        document.getElementById('employeeModal').style.display = 'none';
        document.getElementById('employeeForm').reset();
    }

    saveEmployee(e) {
        e.preventDefault();
        
        const name = document.getElementById('empName').value;
        const position = document.getElementById('empPosition').value;
        const department = document.getElementById('empDepartment').value;
        
        const newEmployee = {
            id: this.employees.length + 1,
            name: name,
            position: position,
            department: department,
            status: 'active'
        };
        
        this.employees.push(newEmployee);
        this.saveToLocalStorage();
        this.loadEmployees();
        this.hideEmployeeModal();
        this.showToast('Karyawan berhasil ditambahkan!', 'success');
    }

    loadEmployeesTable() {
        const tbody = document.getElementById('employeesTable');
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

    // Dashboard Functions
    updateDashboard() {
        // Update employee count
        document.getElementById('totalEmployees').textContent = 
            this.employees.filter(emp => emp.status === 'active').length;
        
        // Update present today
        const today = new Date().toISOString().split('T')[0];
        const presentToday = this.attendance.filter(record => 
            record.date === today && record.type === 'in'
        ).length;
        document.getElementById('presentToday').textContent = presentToday;
        
        // Update overtime hours (simple calculation)
        const overtimeCount = this.attendance.filter(record => 
            this.getAttendanceStatus(record.checkIn, record.checkOut) === 'overtime'
        ).length;
        document.getElementById('overtimeHours').textContent = overtimeCount;
        
        // Update attendance rate
        const totalPossible = this.employees.length;
        const attendanceRate = totalPossible > 0 ? 
            Math.round((presentToday / totalPossible) * 100) : 0;
        document.getElementById('attendanceRate').textContent = `${attendanceRate}%`;
        
        this.updateRecentActivity();
        this.updateChart();
    }

    updateRecentActivity() {
        const container = document.getElementById('recentActivity');
        container.innerHTML = '';
        
        const recentRecords = this.attendance
            .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp))
            .slice(0, 10);
        
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
        const date = new Date(timestamp);
        return date.toLocaleString('id-ID', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    }

    updateChart() {
        const ctx = document.getElementById('attendanceChart').getContext('2d');
        
        if (this.currentChart) {
            this.currentChart.destroy();
        }
        
        // Simple chart data - you can enhance this
        const last7Days = this.getLast7Days();
        const chartData = last7Days.map(day => {
            return this.attendance.filter(record => 
                record.date === day && record.type === 'in'
            ).length;
        });
        
        this.currentChart = new Chart(ctx, {
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
        const date = new Date(dateString);
        return date.toLocaleDateString('id-ID', {
            day: '2-digit',
            month: '2-digit'
        });
    }

    // Placeholder functions for future implementation
    editAttendance(date, employeeName) {
        this.showToast(`Edit absensi ${employeeName} pada ${date}`, 'info');
    }

    deleteAttendance(date, employeeName) {
        if (confirm(`Hapus absensi ${employeeName} pada ${date}?`)) {
            this.attendance = this.attendance.filter(record => 
                !(record.date === date && record.employeeName === employeeName)
            );
            this.saveToLocalStorage();
            this.loadAttendanceData();
            this.showToast('Absensi berhasil dihapus!', 'success');
        }
    }

    editEmployee(id) {
        this.showToast(`Edit karyawan ID: ${id}`, 'info');
    }

    deleteEmployee(id) {
        if (confirm('Hapus karyawan ini?')) {
            this.employees = this.employees.filter(emp => emp.id !== id);
            this.saveToLocalStorage();
            this.loadEmployeesTable();
            this.showToast('Karyawan berhasil dihapus!', 'success');
        }
    }

    generateReport() {
        this.showToast('Laporan sedang diproses...', 'info');
        // Implement report generation logic here
    }

    toggleCustomDateRange(value) {
        const customRange = document.getElementById('customDateRange');
        customRange.style.display = value === 'custom' ? 'block' : 'none';
    }
}

// Initialize the system
const attendanceSystem = new AttendanceSystem();
