// Data Storage
class AttendanceSystem {
    constructor() {
        this.employees = JSON.parse(localStorage.getItem('employees')) || this.getDefaultEmployees();
        this.attendance = JSON.parse(localStorage.getItem('attendance')) || [];
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

        // Export
        document.getElementById('exportExcel').addEventListener('click', () => this.exportToExcel());

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
        // Update navigation
        document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
        document.querySelector(`[data-tab="${tabName}"]`).classList.add('active');

        // Update content
        document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
        document.getElementById(tabName).classList.add('active');

        // Update page title
        const titles = {
            'dashboard': 'Dashboard Overview',
            'attendance': 'Sistem Absensi Digital',
            'reports': 'Laporan & Analytics',
            'employees': 'Manajemen Karyawan'
        };
        document.getElementById('pageTitle').textContent = titles[tabName] || 'Sistem Absensi Digital';

        // Load tab-specific data
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
        
        // Clear existing options except the first one
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
        this.saveData();
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
        if (duration.includes('9') || duration.includes('1
