const API_BASE_URL = 'http://localhost:3000/api';

class AttendanceSystem {
    constructor() {
        this.initializeApp();
    }

    async initializeApp() {
        this.initializeTabs();
        await this.loadEmployees();
        this.initializeEventListeners();
        await this.loadAttendanceData();
        this.setCurrentDateTime();
    }

    initializeTabs() {
        const tabs = document.querySelectorAll('.tab');
        const tabContents = document.querySelectorAll('.tab-content');
        
        tabs.forEach(tab => {
            tab.addEventListener('click', () => {
                const tabId = tab.getAttribute('data-tab');
                
                tabs.forEach(t => t.classList.remove('active'));
                tabContents.forEach(content => content.classList.remove('active'));
                
                tab.classList.add('active');
                document.getElementById(tabId).classList.add('active');
            });
        });
    }

    async loadEmployees() {
        try {
            const response = await fetch(`${API_BASE_URL}/employees`);
            const employees = await response.json();
            
            this.populateEmployeeDropdowns(employees);
        } catch (error) {
            this.showNotification('Error loading employees', 'error');
            console.error('Error loading employees:', error);
        }
    }

    populateEmployeeDropdowns(employees) {
        const dropdowns = [
            'employeeName',
            'employeeFilter',
            'overtimeEmployeeFilter'
        ];

        dropdowns.forEach(dropdownId => {
            const dropdown = document.getElementById(dropdownId);
            dropdown.innerHTML = dropdownId === 'employeeName' 
                ? '<option value="">Pilih Karyawan</option>'
                : '<option value="all">Semua Karyawan</option>';
            
            employees.forEach(employee => {
                const option = document.createElement('option');
                option.value = employee.name;
                option.textContent = employee.name;
                dropdown.appendChild(option);
            });
        });
    }

    initializeEventListeners() {
        // Attendance form
        document.getElementById('submitAttendance').addEventListener('click', () => this.submitAttendance());
        document.getElementById('autoTime').addEventListener('click', () => this.setCurrentDateTime());
        
        // Filters
        document.getElementById('employeeFilter').addEventListener('change', () => this.filterAttendanceData());
        document.getElementById('dateFilter').addEventListener('change', () => this.filterAttendanceData());
        document.getElementById('overtimeEmployeeFilter').addEventListener('change', () => this.filterOvertimeData());
        document.getElementById('resetFilters').addEventListener('click', () => this.resetFilters());
        document.getElementById('refreshData').addEventListener('click', () => this.loadAttendanceData());
        document.getElementById('calculateOvertime').addEventListener('click', () => this.calculateOvertime());
        
        // Modal
        document.querySelector('.close').addEventListener('click', () => this.closeModal());
        document.getElementById('editForm').addEventListener('submit', (e) => this.updateAttendance(e));
        
        // Close modal when clicking outside
        window.addEventListener('click', (e) => {
            const modal = document.getElementById('editModal');
            if (e.target === modal) {
                this.closeModal();
            }
        });
    }

    setCurrentDateTime() {
        const now = new Date();
        const date = now.toISOString().split('T')[0];
        const time = now.toTimeString().split(' ')[0].substring(0, 5);
        
        document.getElementById('attendanceDate').value = date;
        document.getElementById('attendanceTime').value = time;
    }

    async submitAttendance() {
        const name = document.getElementById('employeeName').value;
        const type = document.getElementById('attendanceType').value;
        const date = document.getElementById('attendanceDate').value;
        const time = document.getElementById('attendanceTime').value;

        if (!name || !date || !time) {
            this.showNotification('Harap isi semua field', 'error');
            return;
        }

        try {
            const response = await fetch(`${API_BASE_URL}/attendance`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({
                    name,
                    type,
                    date,
                    time
                })
            });

            const result = await response.json();

            if (response.ok) {
                this.showNotification('Absensi berhasil disimpan', 'success');
                this.resetAttendanceForm();
                await this.loadAttendanceData();
            } else {
                this.showNotification(result.error || 'Error saving attendance', 'error');
            }
        } catch (error) {
            this.showNotification('Error saving attendance', 'error');
            console.error('Error:', error);
        }
    }

    resetAttendanceForm() {
        document.getElementById('employeeName').value = '';
        document.getElementById('attendanceType').value = 'in';
        this.setCurrentDateTime();
    }

    async loadAttendanceData() {
        try {
            const response = await fetch(`${API_BASE_URL}/attendance`);
            const data = await response.json();
            
            this.displayAttendanceData(data);
            this.calculateSummary(data);
        } catch (error) {
            this.showNotification('Error loading attendance data', 'error');
            console.error('Error:', error);
        }
    }

    displayAttendanceData(data) {
        const tbody = document.getElementById('attendanceData');
        
        if (data.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" class="empty-state">
                        <div>📊</div>
                        <p>Tidak ada data absensi</p>
                    </td>
                </tr>
            `;
            return;
        }
        
        tbody.innerHTML = '';
        
        data.forEach((record, index) => {
            const row = document.createElement('tr');
            const duration = this.calculateDuration(record.timeIn, record.timeOut);
            
            row.innerHTML = `
                <td><strong>${record.name}</strong></td>
                <td class="date-cell">${this.formatDate(record.date)}</td>
                <td class="time-cell">${record.timeIn || '-'}</td>
                <td class="time-cell">${record.timeOut || '-'}</td>
                <td class="time-cell">${duration}</td>
                <td>
                    <div class="action-buttons">
                        <button class="btn btn-primary btn-sm" onclick="attendanceSystem.editAttendance(${index})">Edit</button>
                        <button class="btn btn-danger btn-sm" onclick="attendanceSystem.deleteAttendance('${record.date}', '${record.name}')">Hapus</button>
                    </div>
                </td>
            `;
            tbody.appendChild(row);
        });
    }

    calculateDuration(timeIn, timeOut) {
        if (!timeIn || !timeOut) return '-';
        
        try {
            const [inHours, inMinutes] = timeIn.split(':').map(Number);
            const [outHours, outMinutes] = timeOut.split(':').map(Number);
            
            const totalInMinutes = inHours * 60 + inMinutes;
            const totalOutMinutes = outHours * 60 + outMinutes;
            
            const durationMinutes = totalOutMinutes >= totalInMinutes 
                ? totalOutMinutes - totalInMinutes 
                : (totalOutMinutes + 24*60) - totalInMinutes;
            
            const hours = Math.floor(durationMinutes / 60);
            const minutes = durationMinutes % 60;
            
            return `${hours} jam ${minutes} menit`;
        } catch (e) {
            return '-';
        }
    }

    formatDate(dateStr) {
        const date = new Date(dateStr);
        return date.toLocaleDateString('id-ID', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        });
    }

    filterAttendanceData() {
        // This would filter the displayed data without reloading from server
        // For now, we'll reload all data and filter client-side
        this.loadAttendanceData();
    }

    filterOvertimeData() {
        // Similar to filterAttendanceData
        this.calculateOvertime();
    }

    resetFilters() {
        document.getElementById('employeeFilter').value = 'all';
        document.getElementById('dateFilter').value = '';
        document.getElementById('overtimeEmployeeFilter').value = 'all';
        this.loadAttendanceData();
    }

    async calculateOvertime() {
        try {
            const response = await fetch(`${API_BASE_URL}/overtime`);
            const overtimeData = await response.json();
            this.displayOvertimeData(overtimeData);
        } catch (error) {
            this.showNotification('Error calculating overtime', 'error');
            console.error('Error:', error);
        }
    }

    displayOvertimeData(data) {
        const tbody = document.getElementById('overtimeData');
        
        if (data.length === 0) {
            tbody.innerHTML = `
                <tr>
                    <td colspan="6" class="empty-state">
                        <div>⏰</div>
                        <p>Tidak ada data lembur</p>
                    </td>
                </tr>
            `;
            return;
        }
        
        tbody.innerHTML = '';
        
        data.forEach(record => {
            const row = document.createElement('tr');
            const overtimePay = record.overtimeHours * 10000;
            
            row.innerHTML = `
                <td><strong>${record.name}</strong></td>
                <td class="date-cell">${this.formatDate(record.date)}</td>
                <td class="time-cell">${record.timeIn}</td>
                <td class="time-cell">${record.timeOut}</td>
                <td class="time-cell overtime-hours">${record.overtimeHours} jam</td>
                <td class="time-cell overtime-hours">Rp ${overtimePay.toLocaleString('id-ID')}</td>
            `;
            tbody.appendChild(row);
        });
    }

    calculateSummary(data) {
        const employees = [...new Set(data.map(record => record.name))];
        document.getElementById('totalEmployees').textContent = employees.length;
        document.getElementById('totalAttendance').textContent = data.length;
        
        // Calculate total overtime hours
        const totalOvertime = data.reduce((sum, record) => {
            if (record.timeIn && record.timeOut) {
                const [inHours] = record.timeIn.split(':').map(Number);
                const [outHours] = record.timeOut.split(':').map(Number);
                const hoursWorked = outHours - inHours;
                return sum + Math.max(0, hoursWorked - 8);
            }
            return sum;
        }, 0);
        
        document.getElementById('totalOvertime').textContent = totalOvertime.toFixed(1) + " jam";
        
        this.displaySummaryTable(data, employees);
    }

    displaySummaryTable(data, employees) {
        const tbody = document.getElementById('summaryData');
        tbody.innerHTML = '';
        
        employees.forEach(employee => {
            const employeeRecords = data.filter(record => record.name === employee);
            const workDays = employeeRecords.length;
            
            let totalWorkMinutes = 0;
            let validRecords = 0;
            let totalOvertime = 0;
            
            employeeRecords.forEach(record => {
                if (record.timeIn && record.timeOut) {
                    const [inHours, inMinutes] = record.timeIn.split(':').map(Number);
                    const [outHours, outMinutes] = record.timeOut.split(':').map(Number);
                    
                    const totalInMinutes = inHours * 60 + inMinutes;
                    const totalOutMinutes = outHours * 60 + outMinutes;
                    const durationMinutes = totalOutMinutes >= totalInMinutes 
                        ? totalOutMinutes - totalInMinutes 
                        : (totalOutMinutes + 24*60) - totalInMinutes;
                    
                    totalWorkMinutes += durationMinutes;
                    validRecords++;
                    
                    // Calculate overtime (more than 8 hours)
                    const hoursWorked = durationMinutes / 60;
                    totalOvertime += Math.max(0, hoursWorked - 8);
                }
            });
            
            const avgWorkHours = validRecords > 0 
                ? `${Math.floor(totalWorkMinutes / validRecords / 60)} jam ${Math.floor((totalWorkMinutes / validRecords) % 60)} menit`
                : "-";
            
            const row = document.createElement('tr');
            row.innerHTML = `
                <td><strong>${employee}</strong></td>
                <td class="time-cell">${workDays} hari</td>
                <td class="time-cell">${avgWorkHours}</td>
                <td class="time-cell overtime-hours">${totalOvertime.toFixed(1)} jam</td>
            `;
            tbody.appendChild(row);
        });
    }

    editAttendance(index) {
        // For now, we'll just show a simple edit modal
        // In a real implementation, you'd load the actual record data
        document.getElementById('editModal').style.display = 'block';
    }

    closeModal() {
        document.getElementById('editModal').style.display = 'none';
    }

    async updateAttendance(e) {
        e.preventDefault();
        // Implementation for updating attendance
        this.showNotification('Fitur edit sedang dalam pengembangan', 'success');
        this.closeModal();
    }

    async deleteAttendance(date, name) {
        if (confirm(`Hapus absensi ${name} pada ${this.formatDate(date)}?`)) {
            try {
                const response = await fetch(`${API_BASE_URL}/attendance`, {
                    method: 'DELETE',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ date, name })
                });

                if (response.ok) {
                    this.showNotification('Absensi berhasil dihapus', 'success');
                    await this.loadAttendanceData();
                } else {
                    this.showNotification('Error deleting attendance', 'error');
                }
            } catch (error) {
                this.showNotification('Error deleting attendance', 'error');
                console.error('Error:', error);
            }
        }
    }

    showNotification(message, type) {
        const notification = document.createElement('div');
        notification.className = `notification ${type}`;
        notification.textContent = message;
        
        document.body.appendChild(notification);
        
        setTimeout(() => {
            notification.remove();
        }, 3000);
    }
}

// Initialize the application
const attendanceSystem = new AttendanceSystem();
